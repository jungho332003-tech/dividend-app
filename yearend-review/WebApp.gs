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
};

function cached_(key, build, ttl) {
  const cache = CacheService.getDocumentCache();
  try {
    const hit = cache.get(key);
    if (hit && hit.indexOf('CHUNKS:') === 0) {
      const n = Number(hit.slice(7));
      const keys = Array.from({ length: n }, (_, i) => `${key}#${i}`);
      const parts = cache.getAll(keys);
      if (keys.every(k => parts[k] != null)) return JSON.parse(keys.map(k => parts[k]).join(''));
    } else if (hit) {
      return JSON.parse(hit);
    }
  } catch (e) { /* 캐시 오류는 무시하고 새로 읽는다 */ }
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

function dropCache_(key) {
  try { CacheService.getDocumentCache().remove(key); } catch (e) { /* 캐시 없음 */ }
}

function clearCaches_() {
  try {
    CacheService.getDocumentCache().removeAll([CK.people, CK.logs, CK.notices, CK.rules, CK.files, 'menus', 'access']);
  } catch (e) { /* 캐시 없음 */ }
}

const READERS = { people: readPeople_, logs: readLogs_, notices: readNotices_, rules: readRules_, files: readFiles_ };

/** 저장 후 바뀐 부분만 새로 읽어 캐시에 넣고 화면에 돌려준다 */
function refreshPart_(part) {
  const value = READERS[part](SpreadsheetApp.getActive());
  putCache_(CK[part], value);
  const out = {};
  out[part] = part === 'files' ? visibleFiles_(value, currentMember_(getMembers())) : value;
  return out;
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
    people: seesPeople ? cached_(CK.people, () => readPeople_(ss)) : [],
    logs: can('review') ? cached_(CK.logs, () => readLogs_(ss)) : [],
    notices: can('notice') || can('dash') ? cached_(CK.notices, () => readNotices_(ss)) : [],
    reads: getReads_(),
    rules: can('rules') ? cached_(CK.rules, () => readRules_(ss)) : [],
    files: can('files') || can('review') ? visibleFiles_(cached_(CK.files, () => readFiles_(ss)), me) : [],
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
  };
  me.menus = allowedMenus_(me);
  return me;
}
