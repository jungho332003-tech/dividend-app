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
 * 대상자가 많아 100KB를 넘으면 여러 조각으로 나눠 저장한다.
 */

const CACHE_TTL = 60;
const CHUNK = 90000;
const CK = {
  people: 'people',
  logs: 'logs',
  notices: 'notices',
  rules: 'rules',
  files: 'files',
  owner: 'owner',
  ctx: 'ctx',
};

/** 캐시에 있으면 값, 없으면 null (새로 만들지 않는다) */
function peekCache_(key) {
  try {
    const cache = CacheService.getDocumentCache();
    const hit = cache.get(key);
    if (hit && hit.indexOf('CHUNKS:') === 0) {
      const n = Number(hit.slice(7));
      const keys = Array.from({ length: n }, (_, i) => `${key}#${i}`);
      const parts = cache.getAll(keys);
      if (keys.every(k => parts[k] != null)) return JSON.parse(keys.map(k => parts[k]).join(''));
    } else if (hit) {
      return JSON.parse(hit);
    }
  } catch (e) { /* 캐시 오류는 없는 것으로 본다 */ }
  return null;
}

function cached_(key, build, ttl) {
  const hit = peekCache_(key);
  if (hit !== null) return hit;
  const value = build();
  putCache_(key, value, ttl);
  return value;
}

function putCache_(key, value, ttl) {
  try {
    const cache = CacheService.getDocumentCache();
    const s = JSON.stringify(value);
    if (s.length <= CHUNK) {
      cache.put(key, s, ttl || CACHE_TTL);
      return;
    }
    const parts = {};
    let n = 0;
    for (let i = 0; i < s.length; i += CHUNK) parts[`${key}#${n++}`] = s.slice(i, i + CHUNK);
    if (n > 40) return; // 너무 크면 캐시하지 않는다
    cache.putAll(parts, ttl || CACHE_TTL);
    cache.put(key, 'CHUNKS:' + n, ttl || CACHE_TTL);
  } catch (e) { /* 캐시할 수 없으면 건너뛴다 */ }
}

/** 캐시된 목록이 있으면 그 자리에서 고친다. 없으면 다음 조회 때 시트에서 새로 읽으므로 그냥 둔다 */
function patchCache_(key, fn) {
  const list = peekCache_(key);
  if (list === null) return;
  fn(list);
  putCache_(key, list);
}

function dropCache_(key) {
  try { CacheService.getDocumentCache().remove(key); } catch (e) { /* 캐시 없음 */ }
}

function clearCaches_() {
  try {
    CacheService.getDocumentCache().removeAll([CK.people, CK.logs, CK.notices, CK.rules, CK.files, CK.ctx, 'menus', 'access']);
  } catch (e) { /* 캐시 없음 */ }
}

const READERS = { people: readPeople_, logs: readLogs_, notices: readNotices_, rules: readRules_, files: readFiles_ };

/** 저장 후 바뀐 부분만 새로 읽어 캐시에 넣고 화면에 돌려준다 */
function refreshPart_(part) {
  const value = READERS[part](SpreadsheetApp.getActive());
  putCache_(CK[part], value);
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
  if (part === 'people') return value.filter(p => canSee_(me, p));
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
  return p.owner ? list.indexOf(p.owner) >= 0 : list.indexOf(SCOPE_UNASSIGNED) >= 0;
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

/** 화면 전체 데이터. 권한 없는 메뉴의 데이터는 아예 보내지 않는다 */
function apiBootstrap() {
  const ctx = getContext();
  const ss = ctx.ss;
  const me = currentMember_(ctx.members);
  const can = k => me.menus.indexOf(k) >= 0;
  const seesPeople = can('review') || can('dash');
  return {
    team: String(ctx.cfg[CFG.TEAM] || '연말정산 검토'),
    year: ctx.year,
    sheetUrl: ss.getUrl(),
    me: me,
    staff: ctx.members.map(m => ({ name: m.name, role: m.role })),
    cols: COLS.map(c => ({ k: c.k, l: c.l, t: c.t || '', g: c.g, who: c.who, re: c.re || '' })),
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
