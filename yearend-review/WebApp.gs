/**
 * 웹앱
 *
 * 배포: Apps Script 편집기 > 배포 > 새 배포 > 웹 앱
 *   - 다음 사용자 인증 정보로 실행: "웹 앱에 액세스하는 사용자"
 *   - 액세스 권한: 조직 내 모든 사용자 또는 "Google 계정이 있는 모든 사용자"
 * 접속한 사람의 이메일로 [담당자] 시트에서 이름/역할을 찾고, 시트 보호 권한도 그대로 적용된다.
 */

function doGet() {
  const team = String(getConfig()[CFG.TEAM] || '연말정산 검토');
  return HtmlService.createHtmlOutputFromFile('App')
    .setTitle(team)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    // 휴대폰 홈 화면에 추가했을 때 앱처럼 보이게
    .addMetaTag('mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-capable', 'yes');
}

/* ---------- 캐시 ----------
 * 화면을 열 때마다 시트 전체를 다시 읽지 않도록 60초간 캐시한다 (담당자 전체가 공유).
 * 웹앱에서 저장하면 해당 부분 캐시를 바로 갱신하고, 시트를 직접 고친 내용은 최대 60초 뒤 반영된다.
 *
 * 여러 사람이 동시에 쓸 때 꼬이지 않게:
 *  - 저장할 때마다 그 부분의 "버전"을 바꾼다. 시트를 읽어 캐시를 만드는 동안 누가 저장했으면
 *    그 캐시는 버린다 (저장 전 내용으로 덮어쓰지 않게).
 *  - 캐시 한 칸은 100KB까지라 바이트 기준으로 나눠 저장한다 (한글은 한 글자가 3바이트).
 *  - 대상자 목록은 열 이름을 반복하지 않는 배열로 줄여 저장·전송한다 (2천 명 기준 절반 이하).
 */

const CACHE_TTL = 60;
const CHUNK_BYTES = 95000;
const MAX_CHUNKS = 100;
// 연도별 데이터는 보고 있는 연도를 붙여 따로 캐시한다 (people:2026, people:2027 …)
const CK = {
  get people() { return 'people:' + viewYear_(); },
  get logs() { return 'logs:' + viewYear_(); },
  get notices() { return 'notices:' + viewYear_(); },
  get rules() { return 'rules:' + viewYear_(); },
  get files() { return 'files:' + viewYear_(); },
  get events() { return 'events:' + viewYear_(); },
  owner: 'owner',
  ctx: 'ctx',
};

/* 대상자 목록 압축: { k: [열 키…], r: [[행, 최종수정, 수정자, 값…]] } — 빈 칸·체크 안 됨은 0, 체크는 1 */
function encodePeople_(list) {
  const keys = COLS.map(c => c.k);
  return {
    k: keys,
    r: list.map(p => {
      const r = [p.row || 0, p.updated || 0, p.editor || 0];
      COLS.forEach(c => { const v = p[c.k]; r.push(c.t === 'bool' ? (v ? 1 : 0) : (v === '' || v == null ? 0 : v)); });
      while (r.length > 1 && r[r.length - 1] === 0) r.pop();
      return r;
    }),
  };
}

function decodePeople_(obj) {
  if (!obj || !obj.k || obj.k.join(',') !== COLS.map(c => c.k).join(',')) return null; // 열 구성이 바뀌었으면 새로 읽는다
  return obj.r.map(r => {
    const p = { row: r[0] || 0, updated: r[1] ? String(r[1]) : '', editor: r[2] ? String(r[2]) : '' };
    COLS.forEach((c, i) => { const v = r[i + 3]; p[c.k] = c.t === 'bool' ? v === 1 : (v === 0 || v == null ? '' : String(v)); });
    p.id = p.empNo || 'r' + p.row;
    return p;
  });
}

const CODECS_ = { people: { enc: encodePeople_, dec: decodePeople_ } };
const codec_ = key => CODECS_[String(key).split(':')[0]];

/** 문자열을 UTF-8 바이트 기준으로 자른다 */
function splitBytes_(s, max) {
  const parts = [];
  let start = 0, bytes = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    const b = c < 0x80 ? 1 : c < 0x800 ? 2 : c >= 0xD800 && c <= 0xDBFF ? 4 : c >= 0xDC00 && c <= 0xDFFF ? 0 : 3;
    if (bytes + b > max) { parts.push(s.slice(start, i)); start = i; bytes = 0; }
    bytes += b;
  }
  parts.push(s.slice(start));
  return parts;
}

/** 캐시에 있으면 값, 없으면 null (새로 만들지 않는다) */
function peekCache_(key) {
  try {
    const cache = CacheService.getDocumentCache();
    const hit = cache.get(key);
    let value = null;
    if (hit && hit.indexOf('CHUNKS:') === 0) {
      const n = Number(hit.slice(7));
      const keys = Array.from({ length: n }, (_, i) => `${key}#${i}`);
      const parts = cache.getAll(keys);
      if (keys.every(k => parts[k] != null)) value = JSON.parse(keys.map(k => parts[k]).join(''));
    } else if (hit) {
      value = JSON.parse(hit);
    }
    const cd = codec_(key);
    return value !== null && cd ? cd.dec(value) : value;
  } catch (e) { /* 캐시 오류는 없는 것으로 본다 */ }
  return null;
}

/** 저장(쓰기) 버전: 쓸 때마다 새 값. 캐시를 만드는 동안 바뀌었는지 확인하는 데 쓴다 */
function verOf_(key) {
  try { return CacheService.getDocumentCache().get('v:' + key) || ''; } catch (e) { return ''; }
}
function bumpVer_(key) {
  try { CacheService.getDocumentCache().put('v:' + key, Utilities.getUuid(), 21600); } catch (e) { /* 캐시 없음 */ }
}

function cached_(key, build, ttl) {
  const hit = peekCache_(key);
  if (hit !== null) return hit;
  const v0 = verOf_(key);
  const value = build();
  putCache_(key, value, ttl);
  // 읽는 동안 누가 저장했으면 방금 넣은 (저장 전) 캐시는 버린다
  if (verOf_(key) !== v0) removeCache_(key);
  return value;
}

function putCache_(key, value, ttl) {
  try {
    const cache = CacheService.getDocumentCache();
    const cd = codec_(key);
    const s = JSON.stringify(cd ? cd.enc(value) : value);
    const parts = splitBytes_(s, CHUNK_BYTES);
    if (parts.length === 1) {
      cache.put(key, s, ttl || CACHE_TTL);
      return true;
    }
    if (parts.length > MAX_CHUNKS) return false; // 너무 크면 캐시하지 않는다
    const map = {};
    parts.forEach((x, i) => { map[`${key}#${i}`] = x; });
    cache.putAll(map, ttl || CACHE_TTL);
    cache.put(key, 'CHUNKS:' + parts.length, ttl || CACHE_TTL);
    return true;
  } catch (e) { return false; /* 캐시할 수 없으면 건너뛴다 */ }
}

/**
 * 저장한 뒤 캐시를 그 자리에서 고친다 (없으면 다음 조회 때 시트에서 새로 읽으므로 그냥 둔다).
 * 같은 캐시를 동시에 고치면 한쪽이 사라질 수 있으므로 문서 잠금 안에서 부른다.
 */
function patchCache_(key, fn) {
  bumpVer_(key);
  const list = peekCache_(key);
  if (list === null) return;
  fn(list);
  if (!putCache_(key, list)) removeCache_(key);
}

function removeCache_(key) {
  try { CacheService.getDocumentCache().remove(key); } catch (e) { /* 캐시 없음 */ }
}

function dropCache_(key) {
  bumpVer_(key);
  removeCache_(key);
}

function clearCaches_() {
  try {
    const keys = [CK.ctx, 'menus', 'access'];
    const years = yearsWith_(findYears_(), baseYear_());
    years.forEach(y => ['people', 'logs', 'notices', 'rules', 'files', 'events'].forEach(k => keys.push(`${k}:${y}`)));
    CacheService.getDocumentCache().removeAll(keys);
  } catch (e) { /* 캐시 없음 */ }
}

const READERS = { people: readPeople_, logs: readLogs_, notices: readNotices_, rules: readRules_, files: readFiles_, events: readEvents_ };

/** 저장 후 바뀐 부분만 새로 읽어 캐시에 넣고 화면에 돌려준다 */
function refreshPart_(part) {
  const key = CK[part];
  bumpVer_(key);
  const v0 = verOf_(key);
  const value = READERS[part](SpreadsheetApp.getActive());
  putCache_(key, value);
  if (verOf_(key) !== v0) removeCache_(key); // 읽는 사이 다른 저장이 있었으면 캐시는 버린다
  const out = {};
  out[part] = forMe_(part, value, currentMember_(getContext().members));
  return out;
}

/**
 * 접속자에게 보낼 만큼만 거른다 (캐시는 전체를 두고, 보낼 때마다 거른다)
 *  - 대상자·응대기록·대상자 증빙: 열람범위 안의 대상자만
 *  - 첨부: 메뉴 권한이 있는 구분만
 */
function forMe_(part, value, me) {
  if (part === 'people') return encodePeople_(value.filter(p => canSee_(me, p)));
  if (part === 'logs' || part === 'files') {
    const list = part === 'files' ? visibleFiles_(value, me) : value;
    const seen = visibleEmpNos_(me);
    if (!seen) return list;
    if (part === 'logs') return list.filter(l => seen.has(String(l.empNo)));
    return list.filter(f => f.area !== 'person' || seen.has(String(f.ref)));
  }
  return value;
}


/** 열람범위로 볼 수 있는 대상자인가 */
function canSee_(me, p) {
  if (me.isLeader || me.isAdmin || me.scope === SCOPE_ALL) return true;
  if (!me.name) return false;
  if (p.owner === me.name || p.owner2 === me.name) return true;
  const list = me.scope.split(',');
  // 담당자가 정해지지 않은 대상자는 누구나 본다 (맡을 사람을 찾을 수 있게)
  return !p.owner || list.indexOf(p.owner) >= 0;
}

/** 범위가 제한된 사람이 볼 수 있는 사원번호(또는 행 id) 모음. 전체면 null */
function visibleEmpNos_(me) {
  if (me.isLeader || me.isAdmin || me.scope === SCOPE_ALL) return null;
  const people = cached_(CK.people, () => readPeople_(SpreadsheetApp.getActive()));
  return new Set(people.filter(p => canSee_(me, p)).map(p => p.empNo || p.id));
}

/** 메뉴 권한이 없는 구분의 첨부는 보내지 않는다 (예: 서류 검토 권한이 없으면 대상자 증빙 제외) */
function visibleFiles_(files, me) {
  const menus = me.menus;
  return files.filter(f => menus.indexOf(FILES.MENU[f.area] || 'files') >= 0);
}

/* ---------- 조회 ---------- */

/* ---------- 자동 새로고침 ----------
 * 화면은 1~2분마다 apiPoll(내가 가진 버전)을 부른다. 그 사이 누가 저장해 버전이 바뀐 부분만 다시 보낸다.
 * 바뀐 게 없으면 캐시에서 버전 몇 개만 읽고 끝나므로 가볍다.
 */
const POLL_PARTS_ = ['people', 'logs', 'notices', 'events', 'files'];

function partsFor_(me) {
  const can = k => me.menus.indexOf(k) >= 0;
  return {
    people: can('review') || can('dash'),
    logs: can('review'),
    notices: can('notice') || can('dash'),
    events: can('dash'),
    files: can('files') || can('review'),
  };
}

/** 버전은 자료를 읽기 전에 잰다 (읽는 사이 저장이 있으면 다음 번에 다시 받게) */
function stampOf_(year) {
  const keys = POLL_PARTS_.map(p => `v:${p}:${year}`);
  let got = {};
  try { got = CacheService.getDocumentCache().getAll(keys); } catch (e) { /* 캐시 없음 */ }
  const out = { year: year };
  POLL_PARTS_.forEach((p, i) => { out[p] = got[keys[i]] || ''; });
  return out;
}

function apiPoll(stamp) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  const want = partsFor_(me);
  const now = stampOf_(ctx.year);
  const out = { stamp: now, year: ctx.year };
  if (!stamp || stamp.year !== ctx.year) return out; // 다른 연도를 보고 있으면 화면이 알아서 다시 불러온다
  POLL_PARTS_.forEach(part => {
    if (!want[part] || stamp[part] === now[part]) return;
    out[part] = forMe_(part, cached_(CK[part], () => READERS[part](ctx.ss)), me);
  });
  return out;
}

/** 화면 전체 데이터. 권한 없는 메뉴의 데이터는 아예 보내지 않는다 */
function apiBootstrap() {
  const ctx = getContext();
  const ss = ctx.ss;
  const me = currentMember_(ctx.members);
  const can = k => me.menus.indexOf(k) >= 0;
  const seesPeople = can('review') || can('dash');
  const stamp = stampOf_(ctx.year);
  return {
    stamp: stamp,
    team: String(ctx.cfg[CFG.TEAM] || '연말정산 검토'),
    year: ctx.year,
    baseYear: ctx.baseYear,
    years: ctx.years,
    closedYears: ctx.closed,
    closed: ctx.closed.indexOf(ctx.year) >= 0,
    eventKinds: EVENT_KINDS,
    events: can('dash') ? cached_(CK.events, () => readEvents_(ss)) : [],
    sheetUrl: ss.getUrl(),
    me: me,
    staff: ctx.members.map(m => ({ name: m.name, role: m.role })),
    cols: COLS.map(c => ({ k: c.k, l: c.l, t: c.t || '', g: c.g, who: c.who, re: c.re || '', loose: !!c.loose })),
    logKinds: LOG_KINDS,
    noticeCategories: NOTICE.CATEGORIES,
    noticeLeaderOnly: NOTICE.LEADER_ONLY,
    ruleCategories: RULES.CATEGORIES,
    fileAreas: FILES.AREAS,
    people: seesPeople ? forMe_('people', cached_(CK.people, () => readPeople_(ss)), me) : [],
    logs: can('review') ? forMe_('logs', cached_(CK.logs, () => readLogs_(ss)), me) : [],
    notices: can('notice') || can('dash') ? cached_(CK.notices, () => readNotices_(ss)) : [],
    reads: getReads_(),
    rules: can('rules') ? cached_(CK.rules, () => readRules_(ss)) : [],
    files: can('files') || can('review') ? forMe_('files', cached_(CK.files, () => readFiles_(ss)), me) : [],
  };
}

/** 보는 연도 바꾸기 (사람마다 따로 저장). 바꾼 연도의 화면 데이터를 돌려준다 */
function apiSetYear(year) {
  const ctx = getContext();
  const y = String(year || '');
  if (ctx.years.indexOf(y) < 0) throw new Error(`${y}년 자료가 없습니다.`);
  PropertiesService.getUserProperties().setProperty('VIEW_YEAR', y);
  VIEW_MEMO_ = null;
  return apiBootstrap();
}

/** 접속자 → 담당자 정보 */
function currentMember_(members) {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const m = members.find(x => x.email && x.email.toLowerCase() === email);
  const owner = ownerEmail_();
  const isAdmin = (!!owner && owner === email) || !!(m && m.admin);
  const me = {
    email: email,
    name: m ? m.name : '',
    role: m ? m.role : '',
    isLeader: !!(m && m.role === LEADER_ROLE),
    isAdmin: isAdmin,
    // 총괄·관리자는 항상 전체
    scope: (m && m.role === LEADER_ROLE) || isAdmin ? SCOPE_ALL : (m ? m.scope : SCOPE_ALL),
  };
  me.menus = allowedMenus_(me);
  return me;
}
