/**
 * 웹앱 (대시보드 화면)
 *
 * 배포: Apps Script 편집기 > 배포 > 새 배포 > 웹 앱
 *   - 다음 사용자 인증 정보로 실행: "웹 앱에 액세스하는 사용자"
 *   - 액세스 권한: 조직 내 모든 사용자 또는 "Google 계정이 있는 모든 사용자"
 * 접속한 사람의 이메일로 [팀원] 시트에서 이름/역할을 찾고, 시트 보호 권한도 그대로 적용된다.
 */

function doGet() {
  // 첫 화면 데이터를 페이지에 같이 실어 보낸다 (화면이 뜬 뒤 서버를 한 번 더 부르지 않도록)
  let boot = 'null';
  try {
    boot = JSON.stringify(apiBootstrap())
      .replace(/</g, '\\u003c').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
  } catch (e) { /* 실패하면 화면이 열린 뒤 다시 불러온다 */ }
  const html = HtmlService.createHtmlOutputFromFile('App').getContent().replace('/*BOOT*/null', () => boot);
  return HtmlService.createHtmlOutput(html)
    .setTitle('인사팀 업무관리')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    // 휴대폰 홈 화면에 추가했을 때 앱처럼 보이게
    .addMetaTag('mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-capable', 'yes');
}

/* ---------- 캐시 ----------
 * 화면을 열 때마다 시트 전체를 다시 읽지 않도록 60초간 캐시한다 (팀 전체가 공유).
 * 웹앱에서 저장하면 해당 부분 캐시를 바로 갱신하고, 시트를 직접 고친 내용은 최대 60초 뒤 반영된다.
 */

const CACHE_TTL = 60;
const CK = {
  members: ctx => 'members:' + ctx.thisWeek.key,
  hr: ctx => 'hr:' + ctx.thisWeek.key,
  budget: 'budget',
  board: 'board',
  rules: 'rules',
  owner: 'owner',
};

function cached_(key, build, ttl) {
  const cache = CacheService.getDocumentCache();
  try {
    const hit = cache.get(key);
    if (hit) return JSON.parse(hit);
  } catch (e) { /* 캐시 오류는 무시하고 새로 읽는다 */ }
  const value = build();
  putCache_(key, value, ttl);
  return value;
}

function putCache_(key, value, ttl) {
  try {
    CacheService.getDocumentCache().put(key, JSON.stringify(value), ttl || CACHE_TTL);
  } catch (e) { /* 100KB 초과 등 캐시할 수 없으면 건너뛴다 */ }
}

/* ---------- 조회 ---------- */

/** 화면 전체 데이터. google.script.run 은 Date를 넘길 수 없어 문자열로 변환한다. */
function apiBootstrap() {
  const ctx = getContext();
  const ss = ctx.ss;
  const base = baseData_(ctx);
  const can = k => base.me.menus.indexOf(k) >= 0;
  // 권한 없는 메뉴의 데이터는 아예 보내지 않는다
  return Object.assign(base, {
    members: can('weekly') ? cached_(CK.members(ctx), () => readMembers_(ctx)) : [],
    hr: can('weekly') ? cached_(CK.hr(ctx), () => readHr_(ctx)) : { thisWeek: [], nextWeek: [] },
    budget: can('budget') ? cached_(CK.budget, () => readBudget_(ss)) : { accounts: [], requests: [] },
    board: can('board') ? cached_(CK.board, () => readBoard_(ss)) : [],
    rules: can('rules') ? cached_(CK.rules, () => readRules_(ss)) : [],
  });
}

/** 매번 계산해도 가벼운 부분 (접속자, 주차) */
function baseData_(ctx) {
  return {
    team: ctx.cfg[CFG.TEAM] || '',
    sheetUrl: ctx.ss.getUrl(),
    me: currentMember_(ctx.members),
    week: {
      key: ctx.thisWeek.key,
      thisLabel: ctx.thisWeek.label,
      nextLabel: ctx.nextWeek.label,
      deadlineIso: ctx.deadline.toISOString(),
      deadlineText: `${fmt(ctx.deadline, 'M/dd')} (${DAY_KO[ctx.deadline.getDay()]}) ${hourKo(ctx.deadline.getHours())}`,
      title: `${ctx.cfg[CFG.TEAM] || ''} 주간업무 ${deadlineText(ctx)}`,
      thisStart: ymd(ctx.thisWeek.start),
      thisEnd: ymd(ctx.thisWeek.end),
      afterLabel: ctx.afterWeek.label,
    },
    ruleCategories: RULES.CATEGORIES,
    hrCategories: HR.CATEGORIES,
  };
}

function readMembers_(ctx) {
  const report = ctx.ss.getSheetByName(SHEET.REPORT);
  const comments = report ? readLeaderComments_(report, ctx.thisWeek.key) : {};
  return ctx.members.filter(m => m.write).map(m => readMember_(ctx, m.name, comments[m.name]));
}

function readMember_(ctx, name, comment) {
  const input = readInput(ctx.ss, name);
  return {
    name: name,
    hasSheet: input.exists,
    done: input.done,
    note: input.note,
    thisWeek: input.thisWeek.map(toClientItem_),
    nextWeek: input.nextWeek.map(toClientItem_),
    comment: String(comment || ''),
    pre: {
      exists: input.pre.exists,
      note: input.pre.note,
      thisWeek: input.pre.thisWeek.map(toClientItem_),
      nextWeek: input.pre.nextWeek.map(toClientItem_),
    },
  };
}

/** 팀원 한 명만 다시 읽어 캐시를 고치고, 화면에 돌려줄 조각을 만든다 */
function memberPatch_(ctx, name, comment) {
  const key = CK.members(ctx);
  const list = cached_(key, () => readMembers_(ctx));
  const old = list.find(m => m.name === name);
  const patch = readMember_(ctx, name, comment !== undefined ? comment : (old ? old.comment : ''));
  const i = list.findIndex(m => m.name === name);
  if (i >= 0) list[i] = patch; else list.push(patch);
  putCache_(key, list);
  return { memberPatch: patch };
}

/** 저장 후 바뀐 부분만 새로 읽어 캐시에 넣고 화면에 돌려준다 */
function refreshPart_(part) {
  const ss = SpreadsheetApp.getActive();
  const readers = { budget: readBudget_, board: readBoard_, rules: readRules_ };
  const value = readers[part](ss);
  putCache_(CK[part], value);
  const out = {};
  out[part] = value;
  return out;
}

function toClientItem_(item) {
  const due = item.due;
  return {
    task: item.task,
    due: formatDue(due),
    dueIso: due instanceof Date ? ymd(due) : '',
  };
}

/** 접속자 → 팀원 정보 */
function currentMember_(members) {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const m = members.find(x => x.email && x.email.toLowerCase() === email);
  const owner = ownerEmail_();
  const isAdmin = (!!owner && owner === email) || !!(m && m.admin);
  const me = {
    email: email,
    name: m ? m.name : '',
    role: m ? m.role : '',
    writer: !!(m && m.write),
    isLeader: !!(m && m.role === '팀장'),
    isAdmin: isAdmin,
  };
  me.menus = allowedMenus_(me);
  return me;
}

function readBudget_(ss) {
  const reqSh = ss.getSheetByName(BUDGET.SHEET);
  const reqRows = reqSh && reqSh.getLastRow() >= 2
    ? reqSh.getRange(2, 1, reqSh.getLastRow() - 1, BCOL.MONTH).getValues()
    : [];

  return {
    accounts: getAccounts_(),
    requests: reqRows
      .map((r, i) => ({ r: r, row: i + 2 }))
      .filter(x => String(x.r[0]).trim() && String(x.r[BCOL.CANCELED - 1]).toUpperCase() !== 'Y')
      .map(x => ({
        row: x.row,
        id: String(x.r[0]),
        date: x.r[1] instanceof Date ? fmt(x.r[1], 'yyyy-MM-dd HH:mm') : String(x.r[1]),
        req: String(x.r[2]),
        from: String(x.r[3]),
        to: String(x.r[4]),
        amount: Number(x.r[5]) || 0,
        reason: String(x.r[6]),
        files: filesFromCell_(x.r[BCOL.FILES - 1]),
        month: monthOfRow_(x.r),
      }))
      .reverse(),
    months: budgetMonths_(new Date()),
  };
}

/* ---------- 내 주간업무 저장 ---------- */

/**
 * payload: { thisWeek:[{task, due}], nextWeek:[...], note, done }
 * due: 'yyyy-mm-dd' | '지속' | ''
 */
function apiSaveMyWeek(payload) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'weekly');
  if (!me.writer) throw new Error('작성 대상 팀원으로 등록되어 있지 않습니다. 관리자에게 [팀원] 시트 등록을 요청하세요.');

  const sh = ctx.ss.getSheetByName(SHEET.INPUT_PREFIX + me.name) || createMyInputSheet_(ctx, me);

  const toRows = list => (list || [])
    .filter(x => String(x.task || '').trim())
    .slice(0, INPUT.ROWS)
    .map(x => [String(x.task).trim(), parseDue_(x.due)]);

  const thisRows = toRows(payload.thisWeek);
  const nextRows = toRows(payload.nextWeek);

  // 지우기+쓰기를 한 번에: 빈 줄까지 채워서 통째로 덮어쓴다
  sh.getRange(INPUT.FIRST_ROW, 1, INPUT.ROWS, 2).setValues(padRows_(thisRows));
  sh.getRange(INPUT.FIRST_ROW, 4, INPUT.ROWS, 2).setValues(padRows_(nextRows));
  sh.getRange(INPUT.NOTE_CELL).setValue(String(payload.note || '').trim());
  sh.getRange(INPUT.DONE_CELL).setValue(!!payload.done);
  return memberPatch_(ctx, me.name);
}

/**
 * 팀원 목록에는 있는데 입력시트가 아직 없으면(권한 적용 전) 저장할 때 바로 만든다.
 * 본인과 관리자만 편집할 수 있게 잠그고, 다음 보고서 갱신 때 주간보고에 들어간다.
 */
function createMyInputSheet_(ctx, me) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const existing = ctx.ss.getSheetByName(SHEET.INPUT_PREFIX + me.name);
    if (existing) return existing;
    const sh = setupInputSheet_(ctx.ss, me.name, ctx);
    const admins = ctx.members.filter(m => m.admin && m.email).map(m => m.email);
    try {
      protectSheet_(sh, [me.email].concat(admins));
    } catch (e) {
      // 잠금은 관리자가 "권한 한 번에 적용"을 누르면 다시 걸린다
    }
    try { CacheService.getDocumentCache().remove(CK.members(ctx)); } catch (e) { /* 캐시 없음 */ }
    return sh;
  } catch (e) {
    throw new Error('입력시트를 만들 수 없습니다. 스프레드시트 편집 권한이 있는지 관리자에게 확인하세요. (관리 → 권한 한 번에 적용)');
  } finally {
    lock.releaseLock();
  }
}

/**
 * 다음 주 보고 미리 쓰기. payload: { thisWeek, nextWeek, note } (clear:true면 지우기)
 * 월요일 자동 이월 때 이 내용이 금주·차주로 옮겨지고 작성완료로 표시된다.
 */
function apiSavePreWeek(payload) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'weekly');
  if (!me.writer) throw new Error('작성 대상 팀원으로 등록되어 있지 않습니다. 관리자에게 [팀원] 시트 등록을 요청하세요.');
  const sh = ctx.ss.getSheetByName(SHEET.INPUT_PREFIX + me.name) || createMyInputSheet_(ctx, me);
  ensurePreArea_(sh, ctx);

  const toRows = list => payload.clear ? [] : (list || [])
    .filter(x => String(x.task || '').trim())
    .slice(0, INPUT.ROWS)
    .map(x => [String(x.task).trim(), parseDue_(x.due)]);
  // 지우기도 같은 경로: 빈 줄로 통째로 덮어쓴다
  sh.getRange(INPUT.FIRST_ROW, INPUT.PRE_THIS_COL, INPUT.ROWS, 2).setValues(padRows_(toRows(payload.thisWeek)));
  sh.getRange(INPUT.FIRST_ROW, INPUT.PRE_NEXT_COL, INPUT.ROWS, 2).setValues(padRows_(toRows(payload.nextWeek)));
  sh.getRange(INPUT.PRE_NOTE).setValue(payload.clear ? '' : String(payload.note || '').trim());
  sh.getRange(INPUT.PRE_FLAG).setValue(payload.clear ? '' : 'Y');
  return memberPatch_(ctx, me.name);
}

/** 입력 칸 수(INPUT.ROWS)만큼 빈 줄을 채운다 */
function padRows_(rows) {
  const out = rows.slice(0, INPUT.ROWS);
  while (out.length < INPUT.ROWS) out.push(['', '']);
  return out;
}

function parseDue_(v) {
  const s = String(v || '').trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return s;
}

/* ---------- 팀장 코멘트 ---------- */

function apiSaveComment(name, comment) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'weekly');
  if (!me.isLeader) throw new Error('팀장만 코멘트를 남길 수 있습니다.');

  const sh = ctx.ss.getSheetByName(SHEET.REPORT);
  const current = PropertiesService.getDocumentProperties().getProperty(PROP.REPORT_WEEK) === ctx.thisWeek.key;
  if (!sh || !current || sh.getLastRow() < REPORT.FIRST_ROW) {
    throw new Error('이번 주 주간보고가 아직 만들어지지 않았습니다. 잠시 후 다시 시도하세요.');
  }
  const names = sh.getRange(REPORT.FIRST_ROW, 1, sh.getLastRow() - REPORT.FIRST_ROW + 1, 1).getValues();
  const idx = names.findIndex(([v]) => String(v).trim() === name);
  if (idx < 0) throw new Error(`주간보고에서 ${name}님을 찾을 수 없습니다.`);

  const text = String(comment || '').trim();
  sh.getRange(REPORT.FIRST_ROW + idx, REPORT.COMMENT_COL).setValue(text);
  return memberPatch_(ctx, name, text);
}

/* ---------- 예산전용 ---------- */

function apiSubmitBudget(form) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'budget');
  const requester = me.name || String(form.requester || '').trim();
  const id = submitBudgetTransfer(Object.assign({}, form, { requester: requester }));
  return Object.assign(refreshPart_('budget'), { newId: id });
}

/** 신청 취소: 신청자 본인 또는 팀장·관리자. 취소한 건은 예산에서 빠진다. */
function apiCancelBudget(row, id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'budget');
  const sh = ctx.ss.getSheetByName(BUDGET.SHEET);
  const v = sh.getRange(row, 1, 1, 8).getValues()[0];
  if (String(v[BCOL.ID - 1]) !== id) throw new Error('신청 내역이 바뀌었습니다. 새로고침 후 다시 시도하세요.');
  if (String(v[BCOL.REQUESTER - 1]) !== me.name && !me.isLeader && !me.isAdmin) throw new Error('본인이 신청한 건만 취소할 수 있습니다.');
  sh.getRange(row, BCOL.CANCELED).setValue('Y');
  return refreshPart_('budget');
}

/* ---------- 알림 ---------- */

function apiRemind() {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  if (!me.isLeader && !me.isAdmin) throw new Error('팀장 또는 관리자만 알림을 보낼 수 있습니다.');
  return sendReminder_(ctx);
}
