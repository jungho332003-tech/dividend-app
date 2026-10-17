/**
 * 인사팀 업무관리 - 한 파일 설치본 (build.py로 자동 생성, 직접 고치지 말고 원본 .gs 파일을 고친 뒤 다시 생성)
 * Apps Script의 Code.gs 내용을 전부 지우고 이 파일 전체를 붙여넣으세요.
 */

// ===================================================================
// Config.gs
// ===================================================================

/**
 * 공통 설정 / 유틸
 */

const TZ = 'Asia/Seoul';
const DAY_KO = ['일', '월', '화', '수', '목', '금', '토'];

const SHEET = {
  CONFIG: '설정',
  MEMBERS: '팀원',
  HOLIDAYS: '공휴일',
  REPORT: '주간보고',
  INPUT_PREFIX: '입력_',
  ARCHIVE_PREFIX: '보관_',
};

// [설정] 시트의 항목명 (A열) — 값은 B열
const CFG = {
  TEAM: '팀명',
  MONDAY: '이번주 월요일',
  DEADLINE: '이번주 마감일시(비우면 자동)',
  DEADLINE_HOUR: '기본 마감 시각(시)',
  REMIND_HOURS: '미작성 알림(마감 몇 시간 전)',
  CALENDAR: '팀 캘린더 ID(비우면 각자 기본 캘린더)',
  EXTRA_CAL: '함께 볼 캘린더 ID(쉼표로 여러 개)',
  RECRUIT_CAL: '채용 캘린더 ID(채용담당만)',
  DRIVE_FOLDER: '첨부 폴더 ID(구글 드라이브)',
};

// 팀원 입력시트 레이아웃
const INPUT = {
  DONE_CELL: 'B2',
  NOTE_CELL: 'B3',
  THIS_LABEL: 'A5',
  NEXT_LABEL: 'D5',
  FIRST_ROW: 7,
  ROWS: 20,
  // 다음 주 미리 쓰기 (G~K열): 월요일에 금주·차주로 옮겨진다
  PRE_FLAG: 'H2',        // 'Y'면 미리 쓴 내용이 있음
  PRE_NOTE: 'H3',        // 다음 주 휴무계획/특이사항
  PRE_THIS_COL: 7,       // G:H 다음 주의 "금주" 업무·기한
  PRE_NEXT_COL: 10,      // J:K 다다음 주의 "차주" 업무·기한
  // 중요 표시(★)는 업무·기한 바로 오른쪽의 좁은 칸: C(금주) · F(차주) · I(미리 쓰기 금주) · L(미리 쓰기 차주)
  IMP_MARK: '★',
  PRE_THIS_LABEL: 'G5',
  PRE_NEXT_LABEL: 'J5',
};

// 주간보고 시트 레이아웃
const REPORT = {
  HEADER_ROW: 4,
  FIRST_ROW: 6,
  COLS: 7,
  COMMENT_COL: 7,
  MIN_BLOCK_ROWS: 1, // 빈 줄 없이 쓴 만큼만 (아무것도 안 썼으면 이름 줄 하나)
};

const PROP = {
  REPORT_WEEK: 'REPORT_WEEK',
  REPORT_HASH: 'REPORT_HASH', // 지난 보고서 내용 요약값 (바뀐 게 없으면 매시간 새로고침을 건너뛴다)
  REMINDED: 'REMINDED_WEEK',
  CLOSED: 'CLOSED_WEEK',
};

/* ---------- 날짜 ---------- */

function fmt(d, pattern) {
  return Utilities.formatDate(d, TZ, pattern);
}

function ymd(d) {
  return fmt(d, 'yyyy-MM-dd');
}

function addDays(d, n) {
  const x = new Date(d.getTime());
  x.setDate(x.getDate() + n);
  return x;
}

function mondayOf(date) {
  const d = new Date(date.getTime());
  d.setHours(0, 0, 0, 0);
  const dow = d.getDay();
  return addDays(d, dow === 0 ? -6 : 1 - dow);
}

function hourKo(h) {
  if (h < 12) return `오전 ${h}시`;
  return `오후 ${h === 12 ? 12 : h - 12}시`;
}

/** 기한 표시: 날짜면 "10/02 (금)", 그 외(지속 등)는 그대로 */
function formatDue(v) {
  if (v instanceof Date) return `${fmt(v, 'MM/dd')} (${DAY_KO[v.getDay()]})`;
  return String(v || '').trim();
}

/**
 * 주차 정보. 월요일이 속한 달 기준으로 "9월 4주" (그 달의 몇 번째 월요일인지).
 * 기간은 공휴일을 뺀 첫 근무일 ~ 마지막 근무일.
 */
function getWeekInfo(monday, holidays) {
  const days = [0, 1, 2, 3, 4].map(i => addDays(monday, i));
  const work = days.filter(d => !holidays.has(ymd(d)));
  const range = work.length ? work : days;
  const start = range[0];
  const end = range[range.length - 1];
  return {
    key: ymd(monday),
    monday: monday,
    start: start,
    end: end,
    label: `${monday.getMonth() + 1}월 ${Math.ceil(monday.getDate() / 7)}주 (${fmt(start, 'M/dd')}~${fmt(end, 'M/dd')})`,
  };
}

/** 기본 마감: 그 주 목요일 N시. 목요일이 공휴일이면 직전 근무일. */
function defaultDeadline(monday, holidays, hour) {
  let d = addDays(monday, 3);
  while (holidays.has(ymd(d)) && d > monday) d = addDays(d, -1);
  d.setHours(hour, 0, 0, 0);
  return d;
}

function deadlineText(ctx) {
  const regular = `매주 목요일 ${hourKo(ctx.hour)}까지`;
  const d = ctx.deadline;
  if (d.getDay() === 4 && d.getHours() === ctx.hour) return `(${regular})`;
  return `(${regular} / 이번주만 ${fmt(d, 'M/dd')} (${DAY_KO[d.getDay()]}) ${hourKo(d.getHours())}까지)`;
}

/* ---------- 시트 데이터 ---------- */

function getConfig() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.CONFIG);
  const map = {};
  if (!sh) return map;
  sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getValues().forEach(([k, v]) => {
    if (k) map[String(k).trim()] = v;
  });
  return map;
}

function setConfigValue(key, value) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.CONFIG);
  const keys = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      clearContextCache_();
      return;
    }
  }
  sh.appendRow([key, value]);
  clearContextCache_();
}

/** [팀원] 시트: 순서 | 이름 | 이메일(구글 계정) | 역할(팀원/팀장) | 작성대상(Y/N) | 관리자(Y/N) | 예산담당(Y/N) | 알림메일(회사 메일 등) */
function getMembers() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.MEMBERS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues()
    .filter(r => String(r[1]).trim())
    .map(r => ({
      order: Number(r[0]) || 999,
      name: String(r[1]).trim(),
      email: String(r[2]).trim(),
      role: String(r[3]).trim(),
      write: String(r[4]).trim().toUpperCase() !== 'N',
      admin: String(r[5]).trim().toUpperCase() === 'Y',
      budget: String(r[6]).trim().toUpperCase() === 'Y',
      notify: String(r[7] || '').trim(),
      recruit: String(r[8] || '').trim().toUpperCase() === 'Y', // 채용 담당: 수습평가·인터뷰 안내 메일을 받는다
    }))
    .sort((a, b) => a.order - b.order);
}

/** 알림 메일 받는 주소: [팀원] 알림메일 칸이 있으면 그 주소(회사 메일 등), 없으면 구글 계정 이메일 */
function mailTo_(m) {
  return (m && (m.notify || m.email)) || '';
}

function getLeaders(members) {
  return (members || getMembers()).filter(m => m.role === '팀장' && mailTo_(m));
}

function getHolidaySet() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.HOLIDAYS);
  const set = new Set();
  if (!sh || sh.getLastRow() < 2) return set;
  sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(([d]) => {
    if (d instanceof Date) set.add(ymd(d));
  });
  return set;
}

/** 이번 주 실행에 필요한 값 묶음 */
function getContext() {
  const ss = SpreadsheetApp.getActive();
  const base = sheetBase_();
  const cfg = base.cfg;
  const holidays = new Set(base.holidays);
  const monday = mondayOf(cfg[CFG.MONDAY] instanceof Date ? cfg[CFG.MONDAY] : new Date());
  const hour = Number(cfg[CFG.DEADLINE_HOUR]) || 14;
  return {
    ss: ss,
    cfg: cfg,
    holidays: holidays,
    hour: hour,
    thisWeek: getWeekInfo(monday, holidays),
    nextWeek: getWeekInfo(addDays(monday, 7), holidays),
    afterWeek: getWeekInfo(addDays(monday, 14), holidays),
    deadline: cfg[CFG.DEADLINE] instanceof Date ? cfg[CFG.DEADLINE] : defaultDeadline(monday, holidays, hour),
    members: base.members,
  };
}

/**
 * 설정·공휴일·팀원 시트는 거의 바뀌지 않으므로 60초간 캐시한다 (요청마다 시트 3개를 읽지 않도록).
 * 웹앱·메뉴에서 바꾸면 바로 지우고, 시트를 직접 고친 내용은 최대 60초 뒤 반영된다.
 */
const CTX_KEY = 'ctx';
const ISO_DATE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/;

function sheetBase_() {
  try {
    const hit = cacheGet_(CTX_KEY);
    if (hit) return JSON.parse(hit, (k, v) => (typeof v === 'string' && ISO_DATE.test(v) ? new Date(v) : v));
  } catch (e) { /* 캐시 오류는 무시하고 새로 읽는다 */ }
  const base = { cfg: getConfig(), holidays: Array.from(getHolidaySet()), members: getMembers() };
  putCache_(CTX_KEY, base);
  return base;
}

function clearContextCache_() {
  try { CacheService.getDocumentCache().remove(CTX_KEY); } catch (e) { /* 캐시 없음 */ }
}

/** 팀원 입력시트 읽기 */
function readInput(ss, name) {
  const sh = ss.getSheetByName(SHEET.INPUT_PREFIX + name);
  if (!sh) return { exists: false, done: false, note: '', thisWeek: [], nextWeek: [], pre: emptyPre_() };
  // 작성완료·특이사항·업무 목록·미리 쓰기를 한 번에 읽는다 (시트 요청 1회)
  const v = sh.getRange(1, 1, INPUT.FIRST_ROW + INPUT.ROWS - 1, INPUT.PRE_NEXT_COL + 2).getValues();
  const at = a1 => v[Number(a1.slice(1)) - 1][a1.charCodeAt(0) - 65];
  const rows = v.slice(INPUT.FIRST_ROW - 1);
  return {
    exists: true,
    done: at(INPUT.DONE_CELL) === true,
    note: String(at(INPUT.NOTE_CELL) || '').trim(),
    thisWeek: rows.filter(r => String(r[0]).trim()).map(r => ({ task: String(r[0]).trim(), due: r[1], imp: !!String(r[2]).trim() })),
    nextWeek: rows.filter(r => String(r[3]).trim()).map(r => ({ task: String(r[3]).trim(), due: r[4], imp: !!String(r[5]).trim() })),
    pre: readPre_(rows, at),
  };
}

function emptyPre_() {
  return { exists: false, note: '', thisWeek: [], nextWeek: [] };
}

/** 입력시트 G~K열의 "다음 주 미리 쓰기" */
function readPre_(rows, at) {
  const t = INPUT.PRE_THIS_COL - 1, n = INPUT.PRE_NEXT_COL - 1;
  const pre = {
    exists: String(at(INPUT.PRE_FLAG)).toUpperCase() === 'Y',
    note: String(at(INPUT.PRE_NOTE) || '').trim(),
    thisWeek: rows.filter(r => String(r[t]).trim()).map(r => ({ task: String(r[t]).trim(), due: r[t + 1], imp: !!String(r[t + 2]).trim() })),
    nextWeek: rows.filter(r => String(r[n]).trim()).map(r => ({ task: String(r[n]).trim(), due: r[n + 1], imp: !!String(r[n + 2]).trim() })),
  };
  if (!pre.exists) return emptyPre_();
  return pre;
}

function sheetUrl(ss, sh) {
  return `${ss.getUrl()}#gid=${sh.getSheetId()}`;
}

// ===================================================================
// Setup.gs
// ===================================================================

/**
 * 메뉴 / 초기 설정 / 권한
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 주간업무')
    .addItem('🌐 업무관리 웹앱 열기', 'showWebAppLink')
    .addItem('💰 예산전용 신청', 'openBudgetForm')
    .addSeparator()
    .addItem('주간보고 새로고침', 'generateReport')
    .addItem('다음 주차로 넘기기 (차주→금주 이월)', 'menuStartNewWeek')
    .addItem('미작성자 알림 메일 보내기', 'menuSendReminder')
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ 관리자 설정')
      .addItem('1. 기본 시트 만들기', 'initialize')
      .addItem('2. 팀원 입력시트 만들기 + 권한 적용', 'setupMembers')
      .addItem('3. 자동 실행(트리거) 켜기', 'setupTriggers'))
    .addToUi();
}

function showWebAppLink() {
  const url = ScriptApp.getService().getUrl();
  const html = url
    ? `<p style="font-family:sans-serif">아래 주소를 팀원에게 공유하세요.</p><p><a href="${url}" target="_blank">${url}</a></p>`
    : '<p style="font-family:sans-serif">아직 웹앱으로 배포되지 않았습니다. README의 "웹앱 배포" 순서를 따라 주세요.</p>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(460).setHeight(160), '업무관리 웹앱');
}

/** 1단계: 설정/팀원/공휴일/예산 시트 생성 (이미 있으면 건너뜀) */
function initialize() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);

  if (!ss.getSheetByName(SHEET.CONFIG)) {
    const sh = ss.insertSheet(SHEET.CONFIG);
    sh.getRange(1, 1, 10, 2).setValues([
      ['항목', '값'],
      [CFG.TEAM, '인사팀'],
      [CFG.MONDAY, mondayOf(new Date())],
      [CFG.DEADLINE, ''],
      [CFG.DEADLINE_HOUR, 14],
      [CFG.REMIND_HOURS, 3],
      [CFG.CALENDAR, ''],
      [CFG.DRIVE_FOLDER, ''],
      [CFG.EXTRA_CAL, ''],
      [CFG.RECRUIT_CAL, ''],
    ]);
    sh.getRange('B8').setNote('드라이브에서 폴더를 열었을 때 주소 .../folders/ 뒤의 값. 팀원에게 편집자로 공유해야 합니다.');
    sh.getRange('B9').setNote('회사 공용 캘린더·다른 팀 캘린더 등 [일정]에 같이 보여줄 캘린더 ID. 쉼표로 여러 개 (최대 6개). 보기만 하고 일정은 넣지 않습니다.');
    sh.getRange('B10').setNote('채용 전용 캘린더 ID. [팀원] 시트에서 채용담당=Y인 사람만 [일정]에서 보고 일정을 넣을 수 있습니다. 채용담당에게 "일정 변경" 권한으로 공유하세요.');
    sh.getRange('B3').setNumberFormat('yyyy-mm-dd (ddd)');
    sh.getRange('B7').setNote('구글 캘린더 > 설정 > 팀 캘린더 선택 > "캘린더 통합"의 캘린더 ID. 팀원에게 캘린더를 공유해야 합니다.');
    sh.getRange('B4').setNumberFormat('yyyy-mm-dd hh:mm').setNote('명절 등으로 이번 주만 마감이 다르면 입력. 예) 2026-09-23 14:00\n다음 주차로 넘어가면 자동으로 비워집니다.');
    styleHeader_(sh.getRange('A1:B1'));
    sh.setColumnWidth(1, 240).setColumnWidth(2, 200);
  }

  if (!ss.getSheetByName(SHEET.MEMBERS)) {
    const sh = ss.insertSheet(SHEET.MEMBERS);
    const rows = [['순서', '이름', '이메일', '역할', '작성대상(Y/N)', '관리자(Y/N)'], [0, '팀장', '', '팀장', 'N', 'N']];
    for (let i = 1; i <= 8; i++) rows.push([i, `팀원${i}`, '', '팀원', 'Y', 'N']);
    sh.getRange(1, 1, rows.length, 6).setValues(rows);
    styleHeader_(sh.getRange('A1:F1'));
    sh.getRange('F2:F50').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Y', 'N']).build());
    sh.getRange('D2:D50').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['팀원', '팀장']).build());
    sh.getRange('E2:E50').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Y', 'N']).build());
    sh.setColumnWidth(2, 120).setColumnWidth(3, 240).setColumnWidth(5, 110);
    sh.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SHEET.HOLIDAYS)) {
    const sh = ss.insertSheet(SHEET.HOLIDAYS);
    const rows = [['날짜', '명칭']].concat(DEFAULT_HOLIDAYS.map(([d, n]) => [new Date(d + 'T00:00:00+09:00'), n]));
    sh.getRange(1, 1, rows.length, 2).setValues(rows);
    sh.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd (ddd)');
    styleHeader_(sh.getRange('A1:B1'));
    sh.setColumnWidth(1, 150).setColumnWidth(2, 180);
    sh.setFrozenRows(1);
  }

  setupBudgetSheets_(ss);
  setupBoardSheets_(ss);
  setupRuleSheets_(ss);
  setupHrSheet_(ss);
  setupLeaveSheet_(ss);
  setupMenuSheet_(ss);
  ensureHeader_(ss.getSheetByName(SHEET.MEMBERS), 6, '관리자(Y/N)');
  ensureHeader_(ss.getSheetByName(SHEET.MEMBERS), 7, '예산담당(Y/N)');
  ensureHeader_(ss.getSheetByName(SHEET.MEMBERS), 8, '알림메일');
  ensureHeader_(ss.getSheetByName(SHEET.MEMBERS), 9, '채용담당(Y/N)');
  if (!ss.getSheetByName(SHEET.REPORT)) ss.insertSheet(SHEET.REPORT, 0);

  SpreadsheetApp.getUi().alert('기본 시트를 만들었습니다.\n\n[팀원] 시트에 이름·이메일을, [예산과목] 시트에 과목·편성액을 입력한 뒤\n"2. 팀원 입력시트 만들기 + 권한 적용"을 실행하세요.');
}

// 공휴일 기본값 (정부 발표에 따라 변경될 수 있으니 확인 후 수정하세요)
const DEFAULT_HOLIDAYS = [
  ['2026-09-24', '추석'], ['2026-09-25', '추석'], ['2026-09-26', '추석'],
  ['2026-10-03', '개천절'], ['2026-10-05', '개천절 대체공휴일'], ['2026-10-09', '한글날'], ['2026-12-25', '성탄절'],
  ['2027-01-01', '신정'], ['2027-02-08', '설날'], ['2027-02-09', '설날 대체공휴일'],
  ['2027-03-01', '삼일절'], ['2027-05-05', '어린이날'], ['2027-05-13', '부처님오신날'],
  ['2027-08-16', '광복절 대체공휴일'], ['2027-09-14', '추석'], ['2027-09-15', '추석'],
  ['2027-09-16', '추석'], ['2027-10-04', '개천절 대체공휴일'], ['2027-10-11', '한글날 대체공휴일'],
  ['2027-12-27', '성탄절 대체공휴일'],
];

/** 2단계: 팀원별 입력시트 생성 + 시트 보호(본인만 편집). 웹앱 관리 메뉴의 "권한 적용"도 같은 일을 한다. */
function setupMembers() {
  const ctx = getContext();
  const missing = applyProtections_(ctx);
  const msg = missing.length ? `\n\n⚠️ 이메일이 없어 권한을 적용하지 못한 팀원: ${missing.join(', ')}` : '';
  SpreadsheetApp.getUi().alert(`입력시트와 권한을 적용했습니다.\n스프레드시트를 팀원 전원에게 "편집자"로 공유해야 입력할 수 있습니다.\n(웹앱 관리 메뉴의 "권한 한 번에 적용"을 쓰면 공유까지 자동으로 합니다.)${msg}`);
}

/** 입력시트 생성 + 시트 보호. 이메일이 없어 잠그지 못한 팀원 이름 목록을 돌려준다. */
function applyProtections_(ctx) {
  const ss = ctx.ss;
  const leaders = getLeaders(ctx.members).map(m => m.email);
  const admins = ctx.members.filter(m => m.admin && m.email).map(m => m.email);
  const missing = [];

  let created = 0;
  ctx.members.filter(m => m.write).forEach(m => {
    if (!ss.getSheetByName(SHEET.INPUT_PREFIX + m.name)) created++;
    const sh = setupInputSheet_(ss, m.name, ctx);
    if (m.email) protectSheet_(sh, [m.email].concat(admins));
    else missing.push(m.name);
  });

  // 관리 시트: 관리자만
  [SHEET.CONFIG, SHEET.MEMBERS, SHEET.HOLIDAYS, MENU_SHEET].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) protectSheet_(sh, admins);
  });
  // 예산과목·코스트센터·계정과목: 관리자 + 예산 담당자 (웹앱 [예산 설정]에서 고친다)
  const budgetEditors = admins.concat(ctx.members.filter(m => m.budget && m.email).map(m => m.email));
  [BUDGET.ACCOUNTS, BUDGET.COSTS, BUDGET.ACCT_LIST].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) protectSheet_(sh, budgetEditors);
  });

  // 주간보고: 관리자만, 단 팀장 코멘트 칸은 팀장 편집 가능
  const report = ss.getSheetByName(SHEET.REPORT) || ss.insertSheet(SHEET.REPORT, 0);
  protectSheetWithLeaderRange_(report, report.getRange(REPORT.FIRST_ROW, REPORT.COMMENT_COL, report.getMaxRows() - REPORT.FIRST_ROW + 1, 1), leaders.concat(admins), false, admins);

  // 예산전용: 웹앱·메뉴로 신청 (직접 수정 시 경고만). 예전 양식이면 코스트 + 계정과목 양식으로 바꾼다
  migrateBudgetSheets_(ss);
  const budget = ss.getSheetByName(BUDGET.SHEET);
  if (budget) warnOnlyProtect_(budget);

  // 회사 기준: 팀장·관리자만 수정
  [RULES.SHEET, RULES.HISTORY].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) protectSheet_(sh, leaders.concat(admins));
  });

  setupHrSheet_(ss); // 예전 버전에서 업데이트한 경우
  setupLeaveSheet_(ss);
  // 게시판·인사현황·휴가: 누구나 등록 (직접 수정 시 경고만)
  [BOARD.SHEET, BOARD.COMMENTS, HR.SHEET, LEAVE.SHEET, MARK.SHEET, MINUTES.SHEET].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) warnOnlyProtect_(sh);
  });

  // 새 팀원 입력시트가 생겼거나 보고서가 비어 있을 때만 보고서를 다시 만든다 (평소엔 매시간 자동 갱신)
  if (created || report.getLastRow() < REPORT.FIRST_ROW) buildReport_(ctx);
  return missing;
}

function setupInputSheet_(ss, name, ctx) {
  const title = SHEET.INPUT_PREFIX + name;
  let sh = ss.getSheetByName(title);
  if (sh) return sh;

  sh = ss.insertSheet(title);
  sh.getRange('A1').setValue(`${name} 주간업무 입력`).setFontSize(14).setFontWeight('bold');
  sh.getRange('A2:A3').setValues([['작성완료 체크'], ['휴무계획/특이사항']]).setFontWeight('bold');
  sh.getRange(INPUT.DONE_CELL).insertCheckboxes();
  sh.getRange('B3:E3').merge().setBackground('#fffbea').setNote('예) 10/2(금) 연차');

  sh.getRange('A5:B5').merge();
  sh.getRange('D5:E5').merge();
  sh.getRange('A5:E6').setHorizontalAlignment('center').setFontWeight('bold');
  sh.getRange('A6:E6').setValues([['금주 주요업무', '기한', '', '차주 주요업무', '기한']]);
  styleHeader_(sh.getRange('A5:B6'));
  styleHeader_(sh.getRange('D5:E6'));
  writeInputLabels_(sh, ctx);

  const body = sh.getRange(INPUT.FIRST_ROW, 1, INPUT.ROWS, 5);
  body.setVerticalAlignment('middle').setWrap(true);
  sh.getRange(INPUT.FIRST_ROW, 1, INPUT.ROWS, 2).setBorder(true, true, true, true, true, true, '#bbbbbb', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(INPUT.FIRST_ROW, 4, INPUT.ROWS, 2).setBorder(true, true, true, true, true, true, '#bbbbbb', SpreadsheetApp.BorderStyle.SOLID);
  [2, 5].forEach(c => sh.getRange(INPUT.FIRST_ROW, c, INPUT.ROWS, 1).setNumberFormat('mm/dd (ddd)').setHorizontalAlignment('center'));

  sh.getRange(INPUT.FIRST_ROW + INPUT.ROWS + 1, 1)
    .setValue('※ 기한은 날짜(예: 10/02) 또는 "지속"으로 입력. "지속" 업무는 다음 주에도 차주 계획으로 자동 유지됩니다. 작성이 끝나면 작성완료에 체크하세요.')
    .setFontColor('#777777');

  sh.setColumnWidth(1, 380).setColumnWidth(2, 110).setColumnWidth(3, 16).setColumnWidth(4, 380).setColumnWidth(5, 110);
  sh.setFrozenRows(6);
  ensurePreArea_(sh, ctx);
  return sh;
}

function writeInputLabels_(sh, ctx) {
  sh.getRange(INPUT.THIS_LABEL).setValue(ctx.thisWeek.label);
  sh.getRange(INPUT.NEXT_LABEL).setValue(ctx.nextWeek.label);
  sh.getRange(INPUT.PRE_THIS_LABEL).setValue(ctx.nextWeek.label);
  sh.getRange(INPUT.PRE_NEXT_LABEL).setValue(ctx.afterWeek.label);
}

/** 입력시트 오른쪽(G~K열)에 "다음 주 미리 쓰기" 칸을 만든다. 예전에 만든 시트에도 처음 저장할 때 추가된다. */
function ensurePreArea_(sh, ctx) {
  if (sh.getRange('G1').getValue()) return;
  sh.getRange('G1').setValue('다음 주 미리 쓰기 (월요일에 왼쪽 금주·차주로 옮겨집니다)').setFontWeight('bold');
  sh.getRange('G2:G3').setValues([['미리 씀(Y)'], ['다음 주 휴무/특이사항']]).setFontWeight('bold');
  sh.getRange('H3:K3').merge();
  sh.getRange('G5:H5').merge();
  sh.getRange('J5:K5').merge();
  sh.getRange('G6:K6').setValues([['금주 주요업무', '기한', '', '차주 주요업무', '기한']]);
  styleHeader_(sh.getRange('G5:H6'));
  styleHeader_(sh.getRange('J5:K6'));
  sh.getRange(INPUT.PRE_THIS_LABEL).setValue(ctx.nextWeek.label);
  sh.getRange(INPUT.PRE_NEXT_LABEL).setValue(ctx.afterWeek.label);
  [8, 11].forEach(c => sh.getRange(INPUT.FIRST_ROW, c, INPUT.ROWS, 1).setNumberFormat('mm/dd (ddd)').setHorizontalAlignment('center'));
  sh.setColumnWidth(6, 16).setColumnWidth(7, 320).setColumnWidth(8, 110).setColumnWidth(9, 16).setColumnWidth(10, 320).setColumnWidth(11, 110);
}

function styleHeader_(range) {
  range.setBackground('#efefef').setFontWeight('bold').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
}

/* ---------- 보호 ---------- */

/** 시트 전체 보호: 실행한 관리자(소유자) + editors만 편집 */
function protectSheet_(sh, editors) {
  // 이미 원하는 사람들로만 잠겨 있으면 다시 잠그지 않는다 (권한 적용 속도)
  const current = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (current.length === 1 && !current[0].isWarningOnly() && sameEditors_(current[0], editors)) return current[0];
  current.forEach(p => p.remove());
  const p = sh.protect().setDescription(`${sh.getName()} 보호`);
  restrictEditors_(p, editors);
  return p;
}

/**
 * 시트는 관리자만(warningOnly면 경고만), leaderRange는 팀장만 편집.
 */
function protectSheetWithLeaderRange_(sh, leaderRange, leaders, warningOnly, admins) {
  const sheetP = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  const rangeP = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  if (sheetP.length === 1 && rangeP.length === 1
    && sheetP[0].isWarningOnly() === !!warningOnly
    && (warningOnly || sameEditors_(sheetP[0], admins || []))
    && rangeP[0].getRange().getA1Notation() === leaderRange.getA1Notation()
    && sameEditors_(rangeP[0], leaders)) return;
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

  const sp = sh.protect().setDescription(`${sh.getName()} 보호`);
  if (warningOnly) {
    sp.setWarningOnly(true);
  } else {
    restrictEditors_(sp, admins || []);
    sp.setUnprotectedRanges([leaderRange]);
  }

  const rp = leaderRange.protect().setDescription(`${sh.getName()} 팀장 전용`);
  restrictEditors_(rp, leaders);
}

/** 보호의 편집자가 (실행한 관리자 + 소유자 + editors)와 정확히 같은지 */
function sameEditors_(p, editors) {
  if (p.canDomainEdit()) return false;
  const want = new Set(editors.concat(Session.getEffectiveUser().getEmail(), ownerEmail_())
    .filter(Boolean).map(e => e.toLowerCase()));
  const have = new Set(p.getEditors().map(u => u.getEmail().toLowerCase()));
  // 소유자는 보호에서 뺄 수 없으므로 양쪽에 항상 있다고 본다
  const owner = ownerEmail_();
  if (owner) have.add(owner);
  if (want.size !== have.size) return false;
  for (const e of want) if (!have.has(e)) return false;
  return true;
}

/** 스프레드시트 소유자 이메일 (조회가 느려서 6시간 캐시) */
function ownerEmail_() {
  return cached_(CK.owner, () => {
    try {
      const o = SpreadsheetApp.getActive().getOwner();
      return o ? o.getEmail().toLowerCase() : '';
    } catch (e) { return ''; /* 공유 드라이브 등 소유자 조회 불가 */ }
  }, 21600);
}

/** 경고만 하는 시트 보호 (이미 그렇게 되어 있으면 그대로 둔다) */
function warnOnlyProtect_(sh) {
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
  const current = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (current.length === 1 && current[0].isWarningOnly()) return;
  current.forEach(p => p.remove());
  sh.protect().setDescription(`${sh.getName()} 보호`).setWarningOnly(true);
}

function restrictEditors_(p, editors) {
  const me = Session.getEffectiveUser().getEmail();
  p.addEditor(me);
  p.removeEditors(p.getEditors().map(u => u.getEmail()).filter(e => e && e !== me));
  if (p.canDomainEdit()) p.setDomainEdit(false);
  const list = editors.filter(e => e && e !== me);
  if (list.length) p.addEditors(list);
}

// ===================================================================
// Report.gs
// ===================================================================

/**
 * 주간보고 시트 생성 (보고 양식 그대로 조립)
 *
 *  이 름 | 금주 주요업무 | 기한 | 차주 주요업무 | 기한 | 휴무계획/특이사항 | 팀장 코멘트
 *  그 아래에 인사현황 표 (Personnel.gs)
 */

/** 메뉴 [주간보고 새로고침]: 내용이 같아도 다시 만든다 */
function generateReport() {
  withReportLock_(ctx => buildReport_(ctx));
}

/** 매시간 자동 새로고침: 내용이 바뀌었을 때만 다시 그린다 (시트 쓰기 수십 번을 아낀다) */
function refreshReportIfChanged_() {
  withReportLock_(ctx => buildReport_(ctx, true));
}

function withReportLock_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    fn(getContext());
  } finally {
    lock.releaseLock();
  }
}

/** 보고서를 만든다. onlyIfChanged 면 지난번과 내용이 같을 때 건너뛰고 false 를 돌려준다. */
function buildReport_(ctx, onlyIfChanged) {
  const ss = ctx.ss;
  const existing = ss.getSheetByName(SHEET.REPORT);
  const sh = existing || ss.insertSheet(SHEET.REPORT, 0);
  const comments = readLeaderComments_(sh, ctx.thisWeek.key);
  const title = `■ ${ctx.cfg[CFG.TEAM] || ''} 주간업무 ${deadlineText(ctx)}`;

  // 본문 (쓰기 전에 먼저 모은다)
  const rows = [];
  const blocks = [];
  const bold = []; // 팀장이 굵게 표시한 업무 칸 (A1)
  const marks = marksOf_(ctx);
  let row = REPORT.FIRST_ROW;
  ctx.members.filter(m => m.write).forEach(m => {
    const input = readInput(ss, m.name);
    const mk = marks[m.name] || { thisWeek: [], nextWeek: [] };
    const n = Math.max(REPORT.MIN_BLOCK_ROWS, input.thisWeek.length, input.nextWeek.length);
    for (let i = 0; i < n; i++) {
      const a = input.thisWeek[i];
      const b = input.nextWeek[i];
      // 담당자가 중요(★)로 표시했거나 팀장이 굵게 표시한 업무
      if (a && (a.imp || mk.thisWeek.indexOf(a.task) >= 0)) bold.push('B' + (row + i));
      if (b && (b.imp || mk.nextWeek.indexOf(b.task) >= 0)) bold.push('D' + (row + i));
      rows.push([
        i === 0 ? m.name : '',
        a ? a.task : '', a ? formatDue(a.due) : '',
        b ? b.task : '', b ? formatDue(b.due) : '',
        i === 0 ? combineNote_(memberLeaves_(ctx, m.name, ctx.thisWeek.monday, 2), input.note) : '',
        i === 0 ? (comments[m.name] || '') : '',
      ]);
    }
    blocks.push({ row: row, n: n, done: input.done });
    row += n;
  });
  const hrTable = hrTable_(readHr_(ctx));

  const props = PropertiesService.getDocumentProperties();
  const digest = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,
    JSON.stringify([title, ctx.thisWeek.label, ctx.nextWeek.label, rows, blocks, bold, hrTable]), Utilities.Charset.UTF_8));
  if (onlyIfChanged && existing && props.getProperty(PROP.REPORT_WEEK) === ctx.thisWeek.key
    && props.getProperty(PROP.REPORT_HASH) === digest) return false;

  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart().clear();

  // 제목
  sh.getRange('A1').setValue(title).setFontSize(13).setFontWeight('bold');
  sh.getRange('A2').setValue(`※ 이름 칸이 붉으면 작성완료 미체크  |  갱신 ${fmt(new Date(), 'M/dd HH:mm')}`)
    .setFontColor('#b45309');

  // 헤더 (2줄)
  const h = REPORT.HEADER_ROW;
  sh.getRange(h, 1, 2, REPORT.COLS).setValues([
    ['이 름', ctx.thisWeek.label, '', ctx.nextWeek.label, '', '휴무계획/특이사항', '팀장 코멘트'],
    ['', '금주 주요업무', '기한', '차주 주요업무', '기한', '', ''],
  ]);
  ['A', 'F', 'G'].forEach(c => sh.getRange(`${c}${h}:${c}${h + 1}`).merge());
  sh.getRange(h, 2, 1, 2).merge();
  sh.getRange(h, 4, 1, 2).merge();
  sh.getRange(h, 1, 2, REPORT.COLS)
    .setBackground('#efefef').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  if (rows.length) {
    const body = sh.getRange(REPORT.FIRST_ROW, 1, rows.length, REPORT.COLS);
    body.setNumberFormat('@').setValues(rows).setVerticalAlignment('middle').setWrap(true);
    [3, 5].forEach(c => sh.getRange(REPORT.FIRST_ROW, c, rows.length, 1).setHorizontalAlignment('center'));
    sh.getRange(REPORT.FIRST_ROW, REPORT.COMMENT_COL, rows.length, 1).setBackground('#fffbea');

    if (bold.length) sh.getRangeList(bold).setFontWeight('bold');
    blocks.forEach(b => {
      [1, 6, REPORT.COMMENT_COL].forEach(c => sh.getRange(b.row, c, b.n, 1).merge());
      sh.getRange(b.row, 1, b.n, 1).setHorizontalAlignment('center').setBackground(b.done ? null : '#fde2e2');
    });
  }

  const table = sh.getRange(h, 1, 2 + rows.length, REPORT.COLS);
  table.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  blocks.forEach(b => sh.getRange(b.row, 1, b.n, REPORT.COLS)
    .setBorder(true, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM));

  // 인사현황 표 (입사·휴직·복직·퇴사·인터뷰·기타)
  writeHrTable_(sh, ctx, h + 2 + rows.length + 2, hrTable);

  [110, 330, 95, 330, 95, 170, 240].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(h + 1);
  sh.hideColumns(REPORT.COMMENT_COL); // 팀장 코멘트 칸은 쓰지 않으므로 숨긴다

  const saved = {};
  saved[PROP.REPORT_WEEK] = ctx.thisWeek.key;
  saved[PROP.REPORT_HASH] = digest;
  props.setProperties(saved);
  return true;
}

/** 같은 주차 보고서를 다시 만들 때 팀장 코멘트를 유지하기 위해 먼저 읽어둔다 */
function readLeaderComments_(sh, weekKey) {
  const map = {};
  if (PropertiesService.getDocumentProperties().getProperty(PROP.REPORT_WEEK) !== weekKey) return map;
  const last = sh.getLastRow();
  if (last < REPORT.FIRST_ROW) return map;
  sh.getRange(REPORT.FIRST_ROW, 1, last - REPORT.FIRST_ROW + 1, REPORT.COLS).getValues().forEach(r => {
    if (r[0] && r[REPORT.COMMENT_COL - 1]) map[String(r[0]).trim()] = r[REPORT.COMMENT_COL - 1];
  });
  return map;
}

// ===================================================================
// Automation.gs
// ===================================================================

/**
 * 주차 넘기기 / 알림 / 트리거
 */

/* ---------- 주차 넘기기 ---------- */

function menuStartNewWeek() {
  const ui = SpreadsheetApp.getUi();
  const ctx = getContext();
  const target = addDays(ctx.thisWeek.monday, 7);
  const label = getWeekInfo(target, ctx.holidays).label;
  const ok = ui.alert('다음 주차로 넘기기',
    `현재 보고서를 보관하고 [${label}]로 넘어갑니다.\n각자의 "차주 주요업무"가 "금주 주요업무"로 이월됩니다. 진행할까요?`,
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  startNewWeek_(target);
  ui.alert(`[${label}] 주차로 넘어갔습니다.`);
}

/** 트리거: 매주 월요일 아침 자동 실행 */
function weeklyStart() {
  const ctx = getContext();
  const monday = mondayOf(new Date());
  if (monday > ctx.thisWeek.monday) startNewWeek_(monday);
}

function startNewWeek_(targetMonday) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const before = getContext();
    archiveReport_(before);

    setConfigValue(CFG.MONDAY, targetMonday);
    setConfigValue(CFG.DEADLINE, '');

    const ctx = getContext();
    ctx.members.filter(m => m.write).forEach(m => carryOver_(ctx.ss, m.name, ctx));

    const props = PropertiesService.getDocumentProperties();
    props.deleteProperty(PROP.REMINDED);
    props.deleteProperty(PROP.CLOSED);

    buildReport_(ctx);
  } finally {
    lock.releaseLock();
  }
}

/** 차주 → 금주 이월. "지속" 업무는 차주에도 그대로 남긴다. */
function carryOver_(ss, name, ctx) {
  const sh = ss.getSheetByName(SHEET.INPUT_PREFIX + name);
  if (!sh) return;

  const pre = readInput(ss, name).pre;
  // 업무 | 기한 | 중요(★) 를 함께 옮긴다
  const lastNext = sh.getRange(INPUT.FIRST_ROW, 4, INPUT.ROWS, 3).getValues().filter(r => String(r[0]).trim());
  sh.getRange(INPUT.FIRST_ROW, 1, INPUT.ROWS, 3).clearContent();
  sh.getRange(INPUT.FIRST_ROW, 4, INPUT.ROWS, 3).clearContent();

  if (pre.exists) {
    // 미리 써둔 다음 주 보고를 그대로 옮기고 작성완료로 표시
    const toRows = list => list.slice(0, INPUT.ROWS).map(x => [x.task, x.due, x.imp ? INPUT.IMP_MARK : '']);
    const thisRows = toRows(pre.thisWeek), nextRows = toRows(pre.nextWeek);
    if (thisRows.length) sh.getRange(INPUT.FIRST_ROW, 1, thisRows.length, 3).setValues(thisRows);
    if (nextRows.length) sh.getRange(INPUT.FIRST_ROW, 4, nextRows.length, 3).setValues(nextRows);
    sh.getRange(INPUT.DONE_CELL).setValue(true);
    sh.getRange(INPUT.NOTE_CELL).setValue(pre.note);
    clearPre_(sh);
  } else {
    // 기본: 차주 → 금주, "지속" 업무는 차주에도 유지
    const keep = lastNext.filter(r => String(r[1]).trim() === '지속');
    if (lastNext.length) sh.getRange(INPUT.FIRST_ROW, 1, lastNext.length, 3).setValues(lastNext);
    if (keep.length) sh.getRange(INPUT.FIRST_ROW, 4, keep.length, 3).setValues(keep);
    sh.getRange(INPUT.DONE_CELL).setValue(false);
    sh.getRange(INPUT.NOTE_CELL).clearContent();
  }
  writeInputLabels_(sh, ctx);
}

/** 미리 쓰기 칸 비우기 */
function clearPre_(sh) {
  sh.getRange(INPUT.PRE_FLAG).clearContent();
  sh.getRange(INPUT.PRE_NOTE).clearContent();
  sh.getRange(INPUT.FIRST_ROW, INPUT.PRE_THIS_COL, INPUT.ROWS, 3).clearContent();
  sh.getRange(INPUT.FIRST_ROW, INPUT.PRE_NEXT_COL, INPUT.ROWS, 3).clearContent();
}

/** 현재 주간보고를 "보관_yyyy-MM-dd" 시트로 복사해 숨김 */
function archiveReport_(ctx) {
  const ss = ctx.ss;
  const sh = ss.getSheetByName(SHEET.REPORT);
  if (!sh || sh.getLastRow() < REPORT.FIRST_ROW) return;

  const name = SHEET.ARCHIVE_PREFIX + ctx.thisWeek.key;
  const old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);

  const copy = sh.copyTo(ss).setName(name);
  protectSheet_(copy, []);
  copy.hideSheet();
}

/* ---------- 알림 ---------- */

function menuSendReminder() {
  const sent = sendReminder_(getContext());
  SpreadsheetApp.getUi().alert(sent.length ? `알림을 보냈습니다: ${sent.join(', ')}` : '미작성자가 없습니다.');
}

/** 작성완료 미체크 팀원에게 메일 (알림메일 칸이 있으면 그 주소로). 보낸 사람 이름 목록 반환 */
function sendReminder_(ctx) {
  const pending = pendingMembers_(ctx);
  pending.filter(m => mailTo_(m)).forEach(m => {
    const sh = ctx.ss.getSheetByName(SHEET.INPUT_PREFIX + m.name);
    MailApp.sendEmail({
      to: mailTo_(m),
      subject: `[${ctx.cfg[CFG.TEAM]}] 주간업무 작성 요청 - ${ctx.thisWeek.label}`,
      htmlBody:
        `${m.name}님, 이번 주 주간업무 작성을 부탁드립니다.<br>` +
        `마감: <b>${fmt(ctx.deadline, 'M/dd')} (${DAY_KO[ctx.deadline.getDay()]}) ${hourKo(ctx.deadline.getHours())}</b><br><br>` +
        `<a href="${sh ? sheetUrl(ctx.ss, sh) : ctx.ss.getUrl()}">내 입력시트 열기</a><br>` +
        `작성 후 "작성완료 체크"를 눌러주세요.`,
    });
  });
  return pending.filter(m => mailTo_(m)).map(m => m.name);
}

function pendingMembers_(ctx) {
  return ctx.members.filter(m => m.write && !readInput(ctx.ss, m.name).done);
}

/** 마감 시 팀장에게 보고서 링크 + 미작성자 안내 */
function sendClosingMail_(ctx) {
  const leaders = getLeaders(ctx.members);
  if (!leaders.length) return;
  const pending = pendingMembers_(ctx).map(m => m.name);
  const report = ctx.ss.getSheetByName(SHEET.REPORT);
  MailApp.sendEmail({
    to: leaders.map(mailTo_).join(','),
    subject: `[${ctx.cfg[CFG.TEAM]}] 주간업무 취합 완료 - ${ctx.thisWeek.label}`,
    htmlBody:
      `${ctx.thisWeek.label} 주간업무가 취합되었습니다.<br><br>` +
      `<a href="${sheetUrl(ctx.ss, report)}">주간보고 열기</a><br><br>` +
      (pending.length ? `⚠️ 작성완료 미체크: ${pending.join(', ')}` : '✅ 전원 작성완료'),
  });
}

/* ---------- 트리거 ---------- */

/**
 * 매시간 실행
 *  - 팀 캘린더의 휴가 일정을 [휴가] 시트로 가져오기
 *  - 평일 업무시간엔 보고서 자동 새로고침 (바뀐 내용이 있을 때만)
 *  - 마감 N시간 전 미작성자 알림 (주 1회)
 *  - 마감 후 팀장에게 취합 메일 (주 1회)
 *  - 채용 담당자 안내 메일 (입사 70일 수습평가, 인터뷰 전날)
 *  - 화면 데이터 미리 캐시 (첫 접속 속도)
 */
function hourlyCheck() {
  const now = new Date();
  const ctx = getContext();
  const props = PropertiesService.getDocumentProperties();
  const key = ctx.thisWeek.key;

  // 캘린더에 직접 넣은 휴가 일정 → [휴가] (보고서 새로고침 전에)
  syncLeavesFromCalendar_(ctx);

  const day = now.getDay();
  if (day >= 1 && day <= 5 && now.getHours() >= 8 && now.getHours() <= 20) refreshReportIfChanged_();

  const remindHours = Number(ctx.cfg[CFG.REMIND_HOURS]) || 3;
  const remindAt = new Date(ctx.deadline.getTime() - remindHours * 3600 * 1000);
  if (now >= remindAt && now < ctx.deadline && props.getProperty(PROP.REMINDED) !== key) {
    sendReminder_(ctx);
    props.setProperty(PROP.REMINDED, key);
  }

  if (now >= ctx.deadline && props.getProperty(PROP.CLOSED) !== key) {
    refreshReportIfChanged_();
    sendClosingMail_(ctx);
    props.setProperty(PROP.CLOSED, key);
  }

  try { syncTeamHr_(ctx); } catch (e) { /* 팀 캘린더 권한이 없으면 건너뛴다 */ }
  try { syncRecruitHr_(ctx); } catch (e) { /* 채용 캘린더 권한이 없으면 건너뛴다 */ }
  hrMailCheck_(ctx, now); // 채용 담당자: 입사 70일 수습평가·인터뷰 전날 안내 메일
  warmCaches_(ctx); // 다음 접속이 빠르도록 화면 데이터를 미리 캐시
}

// onBudgetEdit: 예전 버전(결재 기능)에서 만든 트리거도 정리한다
const TRIGGER_HANDLERS = ['hourlyCheck', 'weeklyStart', 'onBudgetEdit'];

/** 3단계: 자동 실행 트리거 등록 (다시 실행해도 중복 생성되지 않음) */
function setupTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => TRIGGER_HANDLERS.indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('hourlyCheck').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('weeklyStart').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();

  SpreadsheetApp.getUi().alert(
    '자동 실행을 켰습니다.\n\n' +
    '· 매주 월요일 7시: 다음 주차로 넘기기(이월)\n' +
    '· 매시간: 보고서 새로고침 / 마감 전 미작성 알림 / 마감 후 팀장 메일');
}

// ===================================================================
// Budget.gs
// ===================================================================

/**
 * 예산전용 신청 (결재 없이 신청 내역만 기록)
 *
 * [코스트센터] 코드 | 이름          ← 코스트센터 목록
 * [계정과목]   계정코드 | 계정과목  ← 계정 목록
 * [예산과목]  코스트 | 계정코드 | 계정과목 | 편성액 | 전입 | 전출 | 현재예산   ← 코스트 × 계정 조합별 예산
 * [예산전용]  신청번호 | 신청일시 | 신청자 | 변경전 코스트 | 변경전 계정코드 | 변경전 계정과목
 *             | 변경후 코스트 | 변경후 계정코드 | 변경후 계정과목 | 금액 | 사유 | 취소(Y) | 적용월
 *
 * 예산은 "코스트(코스트센터) + 계정과목" 한 쌍마다 관리한다. 변경전(감액) → 변경후(증액)로 옮긴다.
 * 계정코드는 [예산과목]에서 가져와 신청 내역에 함께 남긴다 (담당자가 ERP 등에 옮겨 적을 때 쓴다).
 * 적용월(yyyy-MM): 어느 달 예산을 전용하는지. 신청할 때 고르고, 내역은 월별로 볼 수 있다.
 * 신청하면 바로 예산에 반영된다. 잘못 신청한 건은 신청자 본인이나 팀장·관리자가 취소한다.
 */

const BUDGET = {
  SHEET: '예산전용',
  ACCOUNTS: '예산과목',
  COSTS: '코스트센터',   // 코드 | 이름
  ACCT_LIST: '계정과목', // 계정코드 | 계정과목
  ACCOUNT_ROWS: 100,
};

const BCOL = {
  ID: 1, DATE: 2, REQUESTER: 3,
  FROM_COST: 4, FROM_CODE: 5, FROM: 6,
  TO_COST: 7, TO_CODE: 8, TO: 9,
  AMOUNT: 10, REASON: 11, CANCELED: 12, MONTH: 13,
};
const BUDGET_HEADER = ['신청번호', '신청일시', '신청자', '변경전 코스트', '변경전 계정코드', '변경전 계정과목',
  '변경후 코스트', '변경후 계정코드', '변경후 계정과목', '금액', '사유', '취소(Y)', '적용월'];
const ACCOUNT_HEADER = ['코스트', '계정코드', '계정과목', '편성액', '전입', '전출', '현재예산'];

function setupBudgetSheets_(ss) {
  if (!ss.getSheetByName(BUDGET.SHEET)) {
    const sh = ss.insertSheet(BUDGET.SHEET);
    sh.getRange(1, 1, 1, BCOL.MONTH).setValues([BUDGET_HEADER]);
    formatBudgetSheet_(sh);
  }

  if (!ss.getSheetByName(BUDGET.ACCOUNTS)) {
    const sh = ss.insertSheet(BUDGET.ACCOUNTS);
    sh.getRange(1, 1, 1, 7).setValues([ACCOUNT_HEADER]);
    const samples = [
      ['1100', '52100', '복리후생비'], ['1100', '52300', '채용비'], ['1100', '52400', '회의비'],
      ['1200', '52500', '교육훈련비'], ['1200', '52600', '도서인쇄비'], ['1200', '52700', '행사비'],
    ];
    sh.getRange(2, 1, BUDGET.ACCOUNT_ROWS, 2).setNumberFormat('@'); // 0으로 시작하는 코드도 그대로
    sh.getRange(2, 1, samples.length, 4).setValues(samples.map(r => r.concat([0])));
    formatAccountSheet_(sh);
  }
  migrateBudgetSheets_(ss);
}

function formatBudgetSheet_(sh) {
  styleHeader_(sh.getRange(1, 1, 1, BCOL.MONTH));
  sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange('D2:I').setNumberFormat('@');
  sh.getRange('J2:J').setNumberFormat('#,##0');
  sh.getRange('K2:K').setWrap(true);
  sh.getRange('M2:M').setNumberFormat('@');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$L2="Y"')
      .setFontColor('#999999').setStrikethrough(true).setRanges([sh.getRange('A2:M')]).build(),
  ]);
  [110, 130, 80, 90, 90, 120, 90, 90, 120, 110, 300, 70, 80].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);
  sh.getRange('L1').setNote('Y면 취소된 건으로 예산에 반영되지 않습니다.');
}

/** 전입·전출·현재예산 수식: 코스트와 계정과목이 모두 같은 신청만 더한다 ("="&칸 → 코스트가 빈칸이어도 빈칸끼리 맞춘다) */
function formatAccountSheet_(sh, rows) {
  rows = rows || BUDGET.ACCOUNT_ROWS;
  styleHeader_(sh.getRange(1, 1, 1, 7));
  const b = `'${BUDGET.SHEET}'!`;
  const formulas = [];
  for (let r = 2; r < 2 + rows; r++) {
    formulas.push([
      `=IF($C${r}="","",SUMIFS(${b}$J:$J,${b}$G:$G,"="&$A${r},${b}$I:$I,"="&$C${r},${b}$L:$L,"<>Y"))`,
      `=IF($C${r}="","",SUMIFS(${b}$J:$J,${b}$D:$D,"="&$A${r},${b}$F:$F,"="&$C${r},${b}$L:$L,"<>Y"))`,
      `=IF($C${r}="","",N($D${r})+E${r}-F${r})`,
    ]);
  }
  sh.getRange(2, 5, rows, 3).setFormulas(formulas);
  sh.getRange(2, 4, rows, 4).setNumberFormat('#,##0');
  sh.getRange(2, 7, rows, 1).setFontWeight('bold');
  sh.setColumnWidth(1, 110).setColumnWidth(2, 110).setColumnWidth(3, 150);
  sh.setFrozenRows(1);
  sh.getRange('A1').setNote('코스트센터 코드. 코스트와 계정과목 한 쌍마다 한 줄씩 입력하고, 편성액은 연간(또는 기간) 예산을 입력하세요.');
}

/**
 * 예전 양식을 지금 양식(코스트 + 계정코드 + 계정과목)으로 바꾼다. 이미 지금 양식이면 아무것도 안 한다.
 *  v1: 과목명만 (전출과목 | 전입과목)
 *  v2: 코스트 + 계정과목 (계정코드 없음)
 * 계정코드는 [예산과목]에 같은 코스트·계정과목이 있으면 그 코드로 채운다.
 */
function migrateBudgetSheets_(ss) {
  const acc = ss.getSheetByName(BUDGET.ACCOUNTS);
  if (acc) {
    const a1 = String(acc.getRange(1, 1).getValue()).trim();
    const b1 = String(acc.getRange(1, 2).getValue()).trim();
    let changed = false;
    if (a1 !== '코스트') { acc.insertColumnsBefore(1, 2); changed = true; }          // v1: 과목명 | 편성액 …
    else if (b1 !== '계정코드') { acc.insertColumnBefore(2); changed = true; }      // v2: 코스트 | 계정과목 …
    if (changed) {
      acc.getRange(1, 1, 1, 7).setValues([ACCOUNT_HEADER]);
      acc.getRange(2, 1, BUDGET.ACCOUNT_ROWS, 2).setNumberFormat('@');
      formatAccountSheet_(acc);
    }
  }

  const req = ss.getSheetByName(BUDGET.SHEET);
  if (!req) return;
  const header = req.getRange(1, 1, 1, 6).getValues()[0].map(v => String(v).trim());
  if (header[BCOL.FROM_CODE - 1] === BUDGET_HEADER[BCOL.FROM_CODE - 1]) return;
  const v1 = header[3] !== '변경전 코스트';
  const last = req.getLastRow();
  const old = last >= 2 ? req.getRange(2, 1, last - 1, 11).getValues() : [];
  const codes = {};
  getAccounts_().forEach(a => { codes[a.cost + '\u0000' + a.name] = a.code; });
  const code = (cost, name) => codes[String(cost).trim() + '\u0000' + String(name).trim()] || '';
  const ym = v => (v instanceof Date ? fmt(v, 'yyyy-MM') : v);
  const rows = old.map(r => v1
    // 번호 | 일시 | 신청자 | 전출 | 전입 | 금액 | 사유 | 취소 | 첨부 | 적용월
    ? [r[0], r[1], r[2], '', code('', r[3]), r[3], '', code('', r[4]), r[4], r[5], r[6], r[7], ym(r[9])]
    // 번호 | 일시 | 신청자 | 변경전 코스트 | 변경전 계정과목 | 변경후 코스트 | 변경후 계정과목 | 금액 | 사유 | 취소 | 적용월
    : [r[0], r[1], r[2], r[3], code(r[3], r[4]), r[4], r[5], code(r[5], r[6]), r[6], r[7], r[8], r[9], ym(r[10])]);
  req.getRange(1, 1, Math.max(last, 1), Math.max(req.getLastColumn(), BCOL.MONTH)).clearContent().clearNote();
  req.getRange(1, 1, 1, BCOL.MONTH).setValues([BUDGET_HEADER]);
  if (rows.length) {
    req.getRange(2, BCOL.FROM_COST, rows.length, 6).setNumberFormat('@');
    req.getRange(2, BCOL.MONTH, rows.length, 1).setNumberFormat('@');
    req.getRange(2, 1, rows.length, BCOL.MONTH).setValues(rows);
  }
  formatBudgetSheet_(req);
}

/* ---------- 신청 폼 (스프레드시트 메뉴) ---------- */

function openBudgetForm() {
  // 한 파일 설치본(dist/Code.gs)은 화면을 BUDGET_FORM_HTML 문자열로 갖고 있다
  const html = (typeof BUDGET_FORM_HTML === 'string'
    ? HtmlService.createHtmlOutput(BUDGET_FORM_HTML)
    : HtmlService.createHtmlOutputFromFile('BudgetForm')).setWidth(520).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '💰 예산전용 신청');
}

/** 폼 초기값: 팀원 목록, 로그인 사용자, 코스트·계정코드·계정과목별 현재예산 */
function getBudgetFormData() {
  const members = getMembers();
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const me = members.find(m => m.email && m.email.toLowerCase() === email);
  return {
    members: members.map(m => m.name),
    me: me ? me.name : '',
    accounts: getAccounts_(),
    months: budgetMonths_(new Date()),
  };
}

/** 신청할 수 있는 달: 올해 1월 ~ 6개월 뒤. [{value:'2026-10', label:'2026년 10월'}], 이번 달은 current:true */
function budgetMonths_(now) {
  const out = [];
  const cur = fmt(now, 'yyyy-MM');
  const d = new Date(now.getFullYear(), 0, 1);
  const end = new Date(now.getFullYear(), now.getMonth() + 6, 1);
  while (d <= end) {
    const v = fmt(d, 'yyyy-MM');
    out.push({ value: v, label: `${d.getFullYear()}년 ${d.getMonth() + 1}월`, current: v === cur });
    d.setMonth(d.getMonth() + 1);
  }
  return out;
}

/** 'yyyy-MM' 검사. 비어 있으면 이번 달 */
function budgetMonth_(v) {
  const s = String(v || '').trim();
  if (!s) return fmt(new Date(), 'yyyy-MM');
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(s)) throw new Error('적용월을 올바르게 골라 주세요.');
  return s;
}

/** 시트의 적용월 칸 → 'yyyy-MM' (적용월이 없는 건은 신청일시 기준) */
function monthOfRow_(r) {
  const m = r[BCOL.MONTH - 1];
  if (m instanceof Date) return fmt(m, 'yyyy-MM');
  if (/^\d{4}-\d{2}$/.test(String(m).trim())) return String(m).trim();
  return r[BCOL.DATE - 1] instanceof Date ? fmt(r[BCOL.DATE - 1], 'yyyy-MM') : String(r[BCOL.DATE - 1]).slice(0, 7);
}

function monthLabel_(ym) {
  return `${Number(ym.slice(0, 4))}년 ${Number(ym.slice(5, 7))}월`;
}

/** "1100 / 52500 교육훈련비" (비어 있는 부분은 뺀다) */
function budgetLabel_(cost, code, account) {
  const acc = [code, account].filter(Boolean).join(' ');
  return cost ? `${cost} / ${acc}` : acc;
}

/* ---------- 코스트센터·계정과목 목록 ---------- */

/** [코스트센터]·[계정과목] 시트 목록. 시트가 아직 없으면 [예산과목]에 쓰인 것으로 만든다 (읽기만). */
function readBudgetLists_(ss) {
  const read = name => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return null;
    return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
      .map(r => ({ code: String(r[0]).trim(), name: String(r[1]).trim() }))
      .filter(x => x.code || x.name);
  };
  const accounts = getAccounts_();
  const uniq = (list, key) => list.filter((x, i, all) => all.findIndex(y => key(y) === key(x)) === i);
  return {
    costs: read(BUDGET.COSTS) || uniq(accounts.filter(a => a.cost).map(a => ({ code: a.cost, name: '' })), x => x.code),
    acctList: read(BUDGET.ACCT_LIST) || uniq(accounts.map(a => ({ code: a.code, name: a.name })), x => x.code + '|' + x.name),
  };
}

function writeBudgetList_(ss, name, header, rows) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() < 1 || String(sh.getRange(1, 1).getValue()) !== header[0]) {
    sh.getRange(1, 1, 1, 2).setValues([header]);
    styleHeader_(sh.getRange(1, 1, 1, 2));
    sh.setColumnWidth(1, 110).setColumnWidth(2, 180);
    sh.setFrozenRows(1);
  }
  const n = Math.max(sh.getLastRow() - 1, rows.length, 1);
  if (sh.getMaxRows() < n + 1) sh.insertRowsAfter(sh.getMaxRows(), n + 1 - sh.getMaxRows());
  sh.getRange(2, 1, n, 2).setNumberFormat('@').clearContent(); // 0으로 시작하는 코드도 그대로
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
}

/**
 * 웹앱 [예산 설정] 저장 (관리자·예산 담당자).
 * data: { costs: [{code, name}], accts: [{code, name}], rows: [{cost, code, name, budget}] }
 *  - costs → [코스트센터], accts → [계정과목], rows(코스트 × 계정 조합별 편성액) → [예산과목] A:D
 *  - 전입·전출·현재예산은 시트 수식이 계산한다.
 * 신청 내역이 있는 조합을 빼면 그 신청이 예산에 안 잡히므로, force 없이는 막고 알려준다.
 */
function apiSaveAccounts(data, force) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'budget');
  if (!me.isAdmin && !me.budgetManager) throw new Error('예산 설정은 관리자와 예산 담당자만 할 수 있습니다.');
  data = data || {};

  // 코스트센터 목록
  const costs = [], costSeen = {};
  (data.costs || []).forEach((c, i) => {
    const code = String(c.code || '').trim(), name = String(c.name || '').trim();
    if (!code && !name) return;
    if (!code) throw new Error(`코스트센터 ${i + 1}번째 줄: 코드를 입력해 주세요.`);
    if (costSeen[code]) throw new Error(`코스트센터 코드가 두 번 있습니다: ${code}`);
    costSeen[code] = true;
    costs.push([code, name]);
  });
  // 계정과목 목록
  const accts = [], acctSeen = {}, codeSeen = {};
  (data.accts || []).forEach((a, i) => {
    const code = String(a.code || '').trim(), name = String(a.name || '').trim();
    if (!code && !name) return;
    if (!name) throw new Error(`계정과목 ${i + 1}번째 줄: 계정과목 이름을 입력해 주세요.`);
    if (code && codeSeen[code]) throw new Error(`계정코드가 두 번 있습니다: ${code}`);
    if (acctSeen[code + '|' + name]) throw new Error(`계정과목이 두 번 있습니다: ${budgetLabel_('', code, name)}`);
    if (code) codeSeen[code] = true;
    acctSeen[code + '|' + name] = true;
    accts.push([code, name]);
  });
  // 조합별 편성액
  const rows = [], seen = {};
  (data.rows || []).forEach((a, i) => {
    const cost = String(a.cost || '').trim(), code = String(a.code || '').trim(), name = String(a.name || '').trim();
    if (!cost && !name && !Number(a.budget)) return; // 빈 줄
    if (!name) throw new Error(`편성 예산 ${i + 1}번째 줄: 계정과목을 골라 주세요.`);
    if (costs.length && cost && !costSeen[cost]) throw new Error(`편성 예산 ${i + 1}번째 줄: 코스트센터 목록에 없는 코드입니다 (${cost}).`);
    if (accts.length && !acctSeen[code + '|' + name]) throw new Error(`편성 예산 ${i + 1}번째 줄: 계정과목 목록에 없는 계정입니다 (${budgetLabel_('', code, name)}).`);
    const key = cost + '|' + name;
    if (seen[key]) throw new Error(`같은 코스트·계정과목이 두 번 있습니다: ${budgetLabel_(cost, code, name)}`);
    seen[key] = true;
    const budget = Number(String(a.budget == null ? '' : a.budget).replace(/[^\d.-]/g, '')) || 0;
    if (budget < 0) throw new Error(`${budgetLabel_(cost, code, name)}: 편성액은 0 이상이어야 합니다.`);
    rows.push([cost, code, name, budget]);
  });

  // 지금 [예산과목]에 있던 조합 중 신청 내역이 있는 것이 사라지는지 확인
  if (!force) {
    const used = {};
    const req = ctx.ss.getSheetByName(BUDGET.SHEET);
    if (req && req.getLastRow() >= 2) {
      req.getRange(2, 1, req.getLastRow() - 1, BCOL.MONTH).getValues().forEach(r => {
        if (String(r[BCOL.CANCELED - 1]).toUpperCase() === 'Y') return;
        [[r[BCOL.FROM_COST - 1], r[BCOL.FROM - 1]], [r[BCOL.TO_COST - 1], r[BCOL.TO - 1]]].forEach(([c, n]) => {
          if (String(n).trim()) used[String(c).trim() + '|' + String(n).trim()] = budgetLabel_(String(c).trim(), '', String(n).trim());
        });
      });
    }
    const before = {};
    getAccounts_().forEach(a => { before[a.cost + '|' + a.name] = true; });
    const lost = Object.keys(used).filter(k => before[k] && !seen[k]).map(k => used[k]);
    if (lost.length) return { needConfirm: true, lost: lost };
  }

  const ss = ctx.ss;
  try { migrateBudgetSheets_(ss); } catch (e) { /* 권한 없음 → 아래에서 알린다 */ }
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    writeBudgetList_(ss, BUDGET.COSTS, ['코드', '이름'], costs);
    writeBudgetList_(ss, BUDGET.ACCT_LIST, ['계정코드', '계정과목'], accts);
    const sh = ss.getSheetByName(BUDGET.ACCOUNTS) || (setupBudgetSheets_(ss), ss.getSheetByName(BUDGET.ACCOUNTS));
    const n = Math.max(BUDGET.ACCOUNT_ROWS, rows.length + 20);
    if (sh.getMaxRows() < n + 1) sh.insertRowsAfter(sh.getMaxRows(), n + 1 - sh.getMaxRows());
    sh.getRange(2, 1, n, 2).setNumberFormat('@'); // 0으로 시작하는 코드도 그대로
    sh.getRange(2, 1, n, 4).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
    formatAccountSheet_(sh, n); // 전입·전출·현재예산 수식 (줄이 늘었으면 그만큼 더)
  } catch (e) {
    throw new Error('예산 시트를 고칠 권한이 없습니다. 관리자가 [관리 → 팀원·권한 → 권한 한 번에 적용]을 한 번 눌러 주세요.');
  } finally {
    lock.releaseLock();
  }
  SpreadsheetApp.flush();
  try { CacheService.getDocumentCache().remove(CK.budget); } catch (e) { /* 캐시 없음 */ }
  return refreshPart_('budget');
}

function getAccounts_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(BUDGET.ACCOUNTS);
  if (!sh || sh.getLastRow() < 2) return [];
  const head = String(sh.getRange(1, 2).getValue()).trim();
  if (head !== '계정코드') return []; // 아직 예전 양식 (관리자가 "권한 한 번에 적용"을 누르면 바뀐다)
  return sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues()
    .filter(r => String(r[2]).trim())
    .map(r => ({
      cost: String(r[0]).trim(),
      code: String(r[1]).trim(),
      name: String(r[2]).trim(),
      budget: Number(r[3]) || 0,
      inAmt: Number(r[4]) || 0,
      outAmt: Number(r[5]) || 0,
      current: Number(r[6]) || 0,
    }));
}

/**
 * 신청 → [예산전용] 시트에 한 줄 추가 (알림 메일은 보내지 않는다). 신청번호 반환
 * form: { requester, month, fromCost, from, toCost, to, amount, reason }
 */
function submitBudgetTransfer(form) {
  const requester = String(form.requester || '').trim();
  const fromCost = String(form.fromCost || '').trim();
  const from = String(form.from || '').trim();
  const toCost = String(form.toCost || '').trim();
  const to = String(form.to || '').trim();
  const reason = String(form.reason || '').trim();
  const amount = Math.round(Number(String(form.amount || '').replace(/,/g, '')));
  const month = budgetMonth_(form.month);

  if (!requester || !from || !to || !reason) throw new Error('모든 항목을 입력해 주세요.');
  if (fromCost === toCost && from === to) throw new Error('변경전과 변경후가 같습니다.');
  if (!(amount > 0)) throw new Error('금액을 올바르게 입력해 주세요.');

  const ss = SpreadsheetApp.getActive();
  try { migrateBudgetSheets_(ss); } catch (e) {
    throw new Error('예산 시트를 새 양식(코스트 + 계정코드 + 계정과목)으로 바꾸지 못했습니다. 관리자에게 "권한 한 번에 적용"을 요청하세요.');
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const accounts = getAccounts_();
    const account = accounts.find(a => a.cost === fromCost && a.name === from);
    const target = accounts.find(a => a.cost === toCost && a.name === to);
    if (!account) throw new Error(`[${budgetLabel_(fromCost, '', from)}] 예산을 찾을 수 없습니다. [예산과목] 시트를 확인하세요.`);
    if (!target) throw new Error(`[${budgetLabel_(toCost, '', to)}] 예산을 찾을 수 없습니다. [예산과목] 시트를 확인하세요.`);
    const fromLabel = budgetLabel_(fromCost, account.code, from);
    if (amount > account.current) {
      throw new Error(`현재예산을 초과합니다. (${fromLabel} 현재예산: ${account.current.toLocaleString()}원)`);
    }

    const sh = ss.getSheetByName(BUDGET.SHEET);
    const now = new Date();
    const id = nextBudgetId_(sh, now);
    const row = sh.getLastRow() + 1;
    // 코드·적용월이 숫자나 날짜로 바뀌지 않게 (0으로 시작하는 코드, '2026-10')
    sh.getRange(row, BCOL.FROM_COST, 1, 6).setNumberFormat('@');
    sh.getRange(row, BCOL.MONTH).setNumberFormat('@');
    sh.getRange(row, 1, 1, BCOL.MONTH).setValues([[id, now, requester, fromCost, account.code, from, toCost, target.code, to, amount, reason, '', month]]);
    SpreadsheetApp.flush();
    try { CacheService.getDocumentCache().remove(CK.budget); } catch (e) { /* 캐시 없음 */ }

    return id;
  } finally {
    lock.releaseLock();
  }
}

/** BT-2026-001 형식 */
function nextBudgetId_(sh, now) {
  const prefix = `BT-${fmt(now, 'yyyy')}-`;
  let max = 0;
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, BCOL.ID, sh.getLastRow() - 1, 1).getValues().forEach(([v]) => {
      const s = String(v);
      if (s.indexOf(prefix) === 0) max = Math.max(max, Number(s.slice(prefix.length)) || 0);
    });
  }
  return prefix + String(max + 1).padStart(3, '0');
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

// ===================================================================
// Calendar.gs
// ===================================================================

/**
 * 일정 (구글 캘린더 연동)
 *
 * [설정] > "팀 캘린더 ID" 에 공유 캘린더 ID를 넣으면 그 캘린더를, 비우면 접속자의 기본 캘린더를 보여준다.
 * 캘린더 ID: 구글 캘린더 > 설정 > (캘린더 선택) > 캘린더 통합 > 캘린더 ID
 * 팀원이 일정을 보려면 그 캘린더가 팀원에게 공유되어 있어야 하고, 추가하려면 "일정 변경" 권한이 필요하다.
 */

/** [설정] "함께 볼 캘린더" — 회사 공용·다른 팀 캘린더 등을 보기 전용으로 같이 띄운다 (최대 6개) */
const EXTRA_CAL_MAX = 6;
function extraCalendarIds_(cfg) {
  const team = String(cfg[CFG.CALENDAR] || '').trim();
  return String(cfg[CFG.EXTRA_CAL] || '').split(/[\s,;]+/).map(s => s.trim()).filter(Boolean)
    .filter((v, i, a) => a.indexOf(v) === i && v !== team).slice(0, EXTRA_CAL_MAX);
}

/** 채용 캘린더: [팀원] 채용담당=Y 인 사람만 보고 쓴다 */
function recruitCalendarId_(cfg) { return String(cfg[CFG.RECRUIT_CAL] || '').trim(); }
function canRecruitCal_(ctx, me) { return !!(me && me.recruiter && recruitCalendarId_(ctx.cfg)); }
function getRecruitCalendar_(ctx) {
  const cal = CalendarApp.getCalendarById(recruitCalendarId_(ctx.cfg));
  if (!cal) throw new Error('채용 캘린더에 접근할 수 없습니다. 캘린더가 공유되어 있는지 관리자에게 확인하세요.');
  return cal;
}

function getTeamCalendar_() {
  const id = String(sheetBase_().cfg[CFG.CALENDAR] || '').trim();
  const cal = id ? CalendarApp.getCalendarById(id) : CalendarApp.getDefaultCalendar();
  if (!cal) throw new Error('팀 캘린더에 접근할 수 없습니다. 캘린더가 공유되어 있는지 관리자에게 확인하세요.');
  return cal;
}

/* ---------- 달력 캐시 ----------
 * 구글 캘린더 읽기가 느려서 결과를 캐시한다.
 *  - 팀 캘린더: 팀 전체가 5분간 공유 (팀 캘린더를 안 정했으면 각자 기본 캘린더라 본인 캐시에)
 *  - 내 캘린더(나만 보기): 본인 캐시에 2분
 * 웹앱에서 일정·휴가를 저장/삭제하면 버전을 바꿔 바로 새로 읽는다.
 */
const CAL_TTL = 300;
const MY_CAL_TTL = 120;

function calVersion_() {
  try { return CacheService.getDocumentCache().get('calver') || '0'; } catch (e) { return '0'; }
}

function bumpCalendar_() {
  try { CacheService.getDocumentCache().put('calver', String(Date.now()), 21600); } catch (e) { /* 캐시 없음 */ }
}

function userCached_(key, build, ttl) {
  let cache = null;
  try {
    cache = CacheService.getUserCache();
    const hit = cache.get(key);
    if (hit) return JSON.parse(hit);
  } catch (e) { /* 캐시 오류는 무시 */ }
  const value = build();
  try { if (cache) cache.put(key, JSON.stringify(value), ttl); } catch (e) { /* 100KB 초과 등 */ }
  return value;
}

function eventJson_(e, kind) {
  return {
    id: e.getId(),
    title: e.getTitle(),
    start: e.getStartTime().toISOString(),
    end: e.getEndTime().toISOString(),
    allDay: e.isAllDayEvent(),
    location: e.getLocation() || '',
    desc: e.getDescription() || '',
    kind: kind,
  };
}

/** startIso ~ endIso 사이 일정 + 공휴일 */
function apiCalendar(startIso, endIso, fresh) {
  const ctx = getContext(); // 팀원 목록은 캐시에서
  requireMenu_(currentMember_(ctx.members), 'calendar');
  if (fresh) bumpCalendar_(); // [새로고침]: 구글 캘린더에서 바로 다시 읽는다
  const start = new Date(startIso);
  const end = new Date(endIso);
  const shared = teamCalendarSet_(ctx);
  const ver = calVersion_();
  const key = `cal:${ver}:${startIso}:${endIso}`;

  const readTeam = () => {
    const cal = getTeamCalendar_();
    const names = ctx.members.map(m => m.name);
    const events = cal.getEvents(start, end).map(e => {
      const ev = eventJson_(e, isLeaveTitle_(e.getTitle(), names) ? 'leave' : 'event');
      ev.title = leaveTitle_(ev.title, names) || ev.title;
      // "연차] 민경"처럼 휴가 형식인데 [팀원]에 없는 이름이면 휴가 색으로만 보여준다 (주간보고에는 안 들어감)
      if (ev.kind === 'event') { const t = leaveTagTitle_(ev.title); if (t) { ev.kind = 'leave'; ev.title = t; ev.ro = true; } }
      return ev;
    });
    const sh = ctx.ss.getSheetByName(SHEET.HOLIDAYS);
    if (sh && sh.getLastRow() >= 2) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([d, name]) => {
        if (d instanceof Date && d >= addDays(start, -1) && d < end) {
          events.push({ id: 'h-' + ymd(d), title: String(name), start: ymd(d), end: ymd(addDays(d, 1)), allDay: true, location: '', desc: '', kind: 'holiday' });
        }
      });
    }
    return { name: cal.getName(), events: events };
  };
  // 팀 캘린더를 안 정했으면 각자 기본 캘린더 → 다른 사람에게 보이면 안 되므로 본인 캐시에
  const team = shared ? cached_(key, readTeam, CAL_TTL) : userCached_(key, readTeam, CAL_TTL);

  // 내 개인 캘린더 (나만 보기) — 팀 캘린더가 따로 있을 때만. 웹앱은 접속자 본인 계정으로 돌아서 본인 것만 보인다
  let mine = { name: '', events: [] };
  if (shared) {
    try {
      mine = userCached_('my' + key, () => {
        const my = CalendarApp.getDefaultCalendar();
        if (my.getId() === String(ctx.cfg[CFG.CALENDAR] || '').trim()) return { name: '', events: [] };
        return { name: my.getName(), events: my.getEvents(start, end).map(e => eventJson_(e, 'mine')) };
      }, MY_CAL_TTL);
    } catch (e) { /* 개인 캘린더를 못 읽으면 팀 일정만 보여준다 */ }
  }
  // 함께 볼 캘린더 — 볼 수 있는지는 사람마다 다르므로 본인 캐시에
  const ids = extraCalendarIds_(ctx.cfg);
  let extras = [];
  if (ids.length) {
    extras = userCached_('x' + key + ':' + ids.join(','), () => ids.map((id, i) => {
      try {
        const c = CalendarApp.getCalendarById(id);
        if (!c) return { name: '', ok: false, events: [] };
        return { name: c.getName(), ok: true, events: c.getEvents(start, end).map(e => Object.assign(eventJson_(e, 'extra'), { cal: i, calName: c.getName() })) };
      } catch (e) { return { name: '', ok: false, events: [] }; }
    }), CAL_TTL);
  }
  // 채용 캘린더 (채용담당만) — 팀 캘린더와 같이 등록한 일정은 한 번만 보이게 제목·시작으로도 겹침을 뺀다
  const me = currentMember_(ctx.members);
  let rec = null;
  if (canRecruitCal_(ctx, me)) {
    try {
      rec = userCached_('r' + key, () => {
        const c = getRecruitCalendar_(ctx);
        return { name: c.getName(), ok: true, events: c.getEvents(start, end).map(e => eventJson_(e, 'recruit')) };
      }, CAL_TTL);
    } catch (e) { rec = { name: '', ok: false, events: [] }; }
  }
  const seen = {};
  const once = list => list.filter(e => !seen[e.id] && (seen[e.id] = true));
  // 같은 일정이 여러 캘린더에 있으면(초대 등) 한 번만: 팀 → 내 캘린더 → 함께 볼 캘린더 순
  const events = once(team.events).concat(once(mine.events));
  if (rec) {
    const same = {};
    events.forEach(e => { same[e.title + '|' + e.start] = true; });
    events.push.apply(events, once(rec.events).filter(e => !same[e.title + '|' + e.start]));
  }
  extras.forEach(x => { events.push.apply(events, once(x.events)); });
  const merged = mergeLeaves_(events, ctx.members.map(m => m.name));
  return {
    calendarName: team.name,
    myCalendarName: mine.name,
    extraCalendars: extras.map((x, i) => ({ cal: i, name: x.name, ok: x.ok })),
    recruitCalendar: rec ? { name: rec.name, ok: rec.ok } : null,
    // 캘린더에 들어 있는 공휴일은 [공휴일] 시트와 겹치면 뺀다
    events: dropHolidayDupes_(merged),
  };
}

/**
 * 구글 캘린더에 공휴일이 일정으로 들어 있으면(공휴일 캘린더를 합쳐 둔 경우 등) [공휴일] 시트의 공휴일과 두 번 보인다.
 * 같은 날의 종일 일정 중 이름이 공휴일과 같거나 서로 포함되면 빼고, 빨간 공휴일만 남긴다.
 */
/** "[연차] 민경", "오전반차] 소영" → 팀 양식 "연차] 민경" (휴가 단어로 시작하는 제목). 아니면 '' */
function leaveTagTitle_(title) {
  const m = /^\s*\[?\s*([가-힣]+?)\s*\]\s*(.+)$/.exec(String(title || ''));
  if (!m) return '';
  const w = m[1].match(LEAVE.WORDS);
  return w && w[0] === m[1] ? leaveLabel_(leaveType_(w[1], m[1]), m[2].trim()) : '';
}

/**
 * 채용 캘린더·내 캘린더·함께 볼 캘린더에 적힌 휴가도 "[연차] 안지선"처럼 이름을 맞춰 휴가 색으로 보여주고,
 * 팀 캘린더에 같은 날 같은 휴가가 있으면 하나만 남긴다. (팀 캘린더 것이 아니면 여기서 지우지는 못하게 ro)
 */
function mergeLeaves_(events, names) {
  const dayOf = e => (/^\d{4}-\d{2}-\d{2}$/.test(e.start) ? e.start : ymd(new Date(e.start)));
  const keyOf = e => e.title.replace(/\s+/g, '') + '|' + dayOf(e);
  const seen = {};
  events.forEach(e => {
    if (e.kind === 'leave' || e.kind === 'event') {
      if (e.kind === 'leave') { seen[keyOf(e)] = true; e._base = true; } // 팀 캘린더 휴가가 기준
      return;
    }
    if (e.kind !== 'recruit' && e.kind !== 'mine' && e.kind !== 'extra') return;
    const t = (isLeaveTitle_(e.title, names) && leaveTitle_(e.title, names)) || leaveTagTitle_(e.title);
    if (!t) return;
    e.title = t;
    e.kind = 'leave';
    e.ro = true;
  });
  return events.filter(e => {
    if (e._base) { delete e._base; return true; }
    if (!e.ro || e.kind !== 'leave') return true;
    const k = keyOf(e);
    if (seen[k]) return false;
    seen[k] = true;
    return true;
  });
}

function dropHolidayDupes_(events) {
  const norm = t => String(t || '').replace(/[\s()\[\]·]/g, '');
  const words = t => String(t || '').split(/[\s()\[\]·]+/).filter(Boolean);
  // "개천절 대체공휴일" ↔ "대체공휴일(개천절)"처럼 순서가 달라도 같은 말이면 같은 공휴일로 본다
  // 거의 같은 이름만 ("한글날 행사 준비" 같은 진짜 일정은 남긴다)
  const EXTRA = ['', '공휴일', '휴일', '연휴', '대체공휴일', '대체휴일'];
  const same = (h, t) => h === t || (t.indexOf(h) >= 0 && EXTRA.indexOf(t.replace(h, '')) >= 0) || (h.indexOf(t) >= 0 && EXTRA.indexOf(h.replace(t, '')) >= 0);
  const dayOf = e => (/^\d{4}-\d{2}-\d{2}$/.test(e.start) ? e.start : ymd(new Date(e.start)));
  const holi = {};
  events.filter(e => e.kind === 'holiday').forEach(e => { (holi[e.start] = holi[e.start] || []).push(e.title); });
  return events.filter(e => {
    if (e.kind === 'holiday' || !e.allDay) return true;
    const names = holi[dayOf(e)];
    if (!names) return true;
    const t = norm(e.title);
    return !names.some(name => {
      const h = norm(name);
      if (!h || !t) return false;
      if (same(h, t)) return true;
      const ws = words(name);
      return ws.length > 1 && ws.every(w => t.indexOf(w) >= 0) && t.length <= h.length + 2;
    });
  });
}

/**
 * 캘린더에 넣은/지운 휴가 일정을 주간보고에 반영.
 * 달력을 그린 뒤 화면이 따로 부른다 → 달력 표시를 늦추지 않는다. 매시간 자동 실행 때도 맞춘다.
 */
function apiSyncLeaves() {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'calendar');
  let changed;
  if (teamCalendarSet_(ctx)) {
    // 이번 주~다다음 주 (주간보고에 나오는 범위)를 읽어 맞춘다
    const start = ctx.thisWeek.monday, end = addDays(ctx.afterWeek.monday, 7);
    const cal = getTeamCalendar_();
    changed = syncLeaveEvents_(ctx, cal.getEvents(start, end), start, end, cal);
  } else {
    changed = syncMyLeaves_(ctx, me); // 팀 캘린더가 없으면 내 캘린더에서 지운 내 휴가만
  }
  // 채용담당이 열면 채용 캘린더의 채용담당 휴가도 맞춘다
  if (canRecruitCal_(ctx, me)) { try { changed = syncRecruitLeaves_(ctx) || changed; } catch (e) { /* 다음에 다시 */ } }
  // 채용담당: 채용 캘린더에 넣은 입사·인터뷰를 인사현황으로
  let hrChanged = false;
  try { hrChanged = syncTeamHr_(ctx); } catch (e) { /* 다음에 다시 */ } // 팀 캘린더에서 고친 인사현황 일정
  if (canRecruitCal_(ctx, me)) { try { hrChanged = syncRecruitHr_(ctx) || hrChanged; } catch (e) { /* 다음에 다시 */ } }
  if (!changed && !hrChanged) return { changed: false };
  const out = { changed: true };
  if (changed) out.members = forViewer_(currentMember_(ctx.members), cached_(CK.members(ctx), () => readMembers_(ctx)));
  if (hrChanged) out.hr = cached_(CK.hr(ctx), () => readHr_(ctx));
  return out;
}

/**
 * ev: { title, date:'yyyy-mm-dd', allDay, startTime:'HH:mm', endTime:'HH:mm', location, desc, target:'team'|'mine' }
 * target 'mine' → 접속자 본인의 기본 캘린더에 "비공개"로 넣는다 (본인만 보기)
 */
function apiAddEvent(ev) {
  const title = String(ev.title || '').trim();
  const m = String(ev.date || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!title || !m) throw new Error('제목과 날짜를 입력해 주세요.');

  const me = currentMember_(sheetBase_().members); // 캐시 (시트를 매번 읽지 않도록)
  requireMenu_(me, 'calendar');
  const ctx = sheetBase_(); // 캐시 (설정·팀원)
  const recruitOnly = ev.target === 'recruit';
  const alsoRecruit = !!ev.recruit && !recruitOnly && ev.target !== 'mine';
  if ((recruitOnly || alsoRecruit) && !canRecruitCal_(ctx, currentMember_(ctx.members))) throw new Error('채용 캘린더는 채용담당만 쓸 수 있습니다.');
  const mine = ev.target === 'mine';
  // 팀 일정은 "이름) 업무명" 으로 만든다 (구글 캘린더에서도 누가 넣었는지 보이게). 나만 보기 일정은 그대로
  const fullTitle = !mine && me.name && title.indexOf(me.name + ')') !== 0 ? `${me.name}) ${title}` : title;
  const cal = mine ? CalendarApp.getDefaultCalendar() : recruitOnly ? getRecruitCalendar_(ctx) : getTeamCalendar_();
  const opts = {
    location: String(ev.location || '').trim(),
    description: [String(ev.desc || '').trim(), me.name ? `등록: ${me.name}` : ''].filter(Boolean).join('\n'),
  };
  const day = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  const make = c => {
    if (ev.allDay) return c.createAllDayEvent(fullTitle, day, opts);
    const at = (t, fallback) => {
      const [h, mi] = String(t || fallback).split(':').map(Number);
      const d = new Date(day.getTime());
      d.setHours(h || 0, mi || 0, 0, 0);
      return d;
    };
    const s = at(ev.startTime, '09:00');
    let e = at(ev.endTime, '10:00');
    if (e <= s) e = new Date(s.getTime() + 3600 * 1000);
    return c.createEvent(fullTitle, s, e, opts);
  };
  let created = null;
  try {
    created = make(cal);
  } catch (err) {
    if (mine) throw new Error('내 캘린더에 일정을 만들지 못했습니다. 잠시 후 다시 시도하세요.');
    if (recruitOnly) throw new Error('채용 캘린더에 일정을 추가할 권한이 없습니다. "일정 변경" 권한을 관리자에게 요청하세요.');
    throw new Error('일정을 추가할 권한이 없습니다. 팀 캘린더의 "일정 변경" 권한을 관리자에게 요청하세요.');
  }
  // 채용 캘린더에도 같이 (팀 캘린더에는 이미 들어갔으므로 실패해도 알리기만)
  let warning = '';
  if (alsoRecruit) {
    try { make(getRecruitCalendar_(ctx)); } catch (err) { warning = '팀 캘린더에는 넣었지만 채용 캘린더에는 권한이 없어 넣지 못했습니다.'; }
  }
  // 개인 캘린더를 다른 사람과 공유해 두었더라도 내용은 안 보이게
  if (mine && created) {
    try { created.setVisibility(CalendarApp.Visibility.PRIVATE); } catch (e) { /* 무시 */ }
  }
  bumpCalendar_();
  return warning ? { warning: warning } : true;
}

/* ---------- 내 메모 (일별, 나만 보기) ----------
 * 접속한 사람의 사용자 속성(UserProperties)에 저장한다 → 스프레드시트에 남지 않고, 관리자도 볼 수 없다.
 * 웹앱이 "접속한 사용자" 권한으로 돌아서 사람마다 따로 저장된다. 한 메모는 8,000자까지.
 */
const MEMO_PREFIX = 'memo:';

function apiMemos() {
  const all = PropertiesService.getUserProperties().getProperties();
  const out = {};
  Object.keys(all).forEach(k => { if (k.indexOf(MEMO_PREFIX) === 0) out[k.slice(MEMO_PREFIX.length)] = all[k]; });
  return out;
}

function apiSaveMemo(ymd, text) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(ymd))) throw new Error('날짜가 올바르지 않습니다.');
  const value = String(text || '').replace(/\s+$/, '');
  if (value.length > 8000) throw new Error('메모는 8,000자까지 쓸 수 있습니다.');
  const props = PropertiesService.getUserProperties();
  if (value.trim()) props.setProperty(MEMO_PREFIX + ymd, value);
  else props.deleteProperty(MEMO_PREFIX + ymd);
  return true;
}

// ===================================================================
// Board.gs
// ===================================================================

/**
 * 공유게시판
 *
 * [게시판]  글번호 | 작성일시 | 작성자 | 분류 | 제목 | 내용 | 링크 | 고정(Y) | 삭제(Y)
 * [댓글]    글번호 | 작성일시 | 작성자 | 내용
 */

const BOARD = {
  SHEET: '게시판',
  COMMENTS: '댓글',
  CATEGORIES: ['공지', '자료', '일반'],
};

const PCOL = { ID: 1, DATE: 2, AUTHOR: 3, CATEGORY: 4, TITLE: 5, BODY: 6, LINK: 7, PINNED: 8, DELETED: 9, FILES: 10 };

function setupBoardSheets_(ss) {
  if (!ss.getSheetByName(BOARD.SHEET)) {
    const sh = ss.insertSheet(BOARD.SHEET);
    sh.getRange(1, 1, 1, 9).setValues([['글번호', '작성일시', '작성자', '분류', '제목', '내용', '링크', '고정(Y)', '삭제(Y)']]);
    styleHeader_(sh.getRange(1, 1, 1, 9));
    sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('F2:F').setWrap(true);
    [70, 130, 80, 60, 260, 420, 200, 70, 70].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
  }
  ensureHeader_(ss.getSheetByName(BOARD.SHEET), PCOL.FILES, '첨부');
  if (!ss.getSheetByName(BOARD.COMMENTS)) {
    const sh = ss.insertSheet(BOARD.COMMENTS);
    sh.getRange(1, 1, 1, 4).setValues([['글번호', '작성일시', '작성자', '내용']]);
    styleHeader_(sh.getRange(1, 1, 1, 4));
    sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('D2:D').setWrap(true);
    [70, 130, 80, 480].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
  }
}

/** 고정글 먼저, 그다음 최신순 */
function readBoard_(ss) {
  const sh = ss.getSheetByName(BOARD.SHEET);
  if (!sh || sh.getLastRow() < 2) return [];

  const comments = {};
  const csh = ss.getSheetByName(BOARD.COMMENTS);
  if (csh && csh.getLastRow() >= 2) {
    csh.getRange(2, 1, csh.getLastRow() - 1, 4).getValues().forEach(r => {
      const id = String(r[0]);
      if (!id) return;
      (comments[id] = comments[id] || []).push({
        date: r[1] instanceof Date ? fmt(r[1], 'yyyy-MM-dd HH:mm') : String(r[1]),
        author: String(r[2]),
        body: String(r[3]),
      });
    });
  }

  return sh.getRange(2, 1, sh.getLastRow() - 1, PCOL.FILES).getValues()
    .filter(r => String(r[0]) && String(r[PCOL.DELETED - 1]).toUpperCase() !== 'Y')
    .map(r => ({
      id: String(r[0]),
      date: r[1] instanceof Date ? fmt(r[1], 'yyyy-MM-dd HH:mm') : String(r[1]),
      author: String(r[2]),
      category: String(r[3]),
      title: String(r[4]),
      body: String(r[5]),
      link: String(r[6] || ''),
      pinned: String(r[7]).toUpperCase() === 'Y',
      files: filesFromCell_(r[PCOL.FILES - 1]),
      comments: comments[String(r[0])] || [],
    }))
    .sort((a, b) => (b.pinned - a.pinned) || b.date.localeCompare(a.date));
}

function apiAddPost(post) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'board');
  if (!me.name) throw new Error('[팀원] 시트에 등록된 사람만 글을 쓸 수 있습니다.');

  const title = String(post.title || '').trim();
  const body = String(post.body || '').trim();
  const category = BOARD.CATEGORIES.indexOf(post.category) >= 0 ? post.category : '일반';
  if (!title || !body) throw new Error('제목과 내용을 입력해 주세요.');
  if (category === '공지' && !me.isLeader && !me.isAdmin) throw new Error('공지는 팀장 또는 관리자만 올릴 수 있습니다.');
  const pinned = !!post.pinned && (me.isLeader || me.isAdmin);

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sh = ctx.ss.getSheetByName(BOARD.SHEET);
    const last = sh.getLastRow();
    const ids = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(([v]) => Number(v) || 0) : [];
    const id = (ids.length ? Math.max.apply(null, ids) : 0) + 1;
    sh.getRange(last + 1, 1, 1, PCOL.FILES).setValues([[id, new Date(), me.name, category, title, body, String(post.link || '').trim(), pinned ? 'Y' : '', '', filesToCell_(post.files)]]);
  } finally {
    lock.releaseLock();
  }
  return refreshPart_('board');
}

function apiDeletePost(id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'board');
  const sh = ctx.ss.getSheetByName(BOARD.SHEET);
  const rows = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 3).getValues();
  const idx = rows.findIndex(r => String(r[0]) === String(id));
  if (idx < 0) throw new Error('글을 찾을 수 없습니다.');
  if (rows[idx][2] !== me.name && !me.isLeader && !me.isAdmin) throw new Error('본인 글만 삭제할 수 있습니다.');
  sh.getRange(idx + 2, PCOL.DELETED).setValue('Y');
  return refreshPart_('board');
}

function apiAddComment(id, body) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'board');
  if (!me.name) throw new Error('[팀원] 시트에 등록된 사람만 댓글을 쓸 수 있습니다.');
  const text = String(body || '').trim();
  if (!text) throw new Error('댓글 내용을 입력해 주세요.');
  ctx.ss.getSheetByName(BOARD.COMMENTS).appendRow([String(id), new Date(), me.name, text]);
  return refreshPart_('board');
}

// ===================================================================
// Rules.gs
// ===================================================================

/**
 * 회사 기준 (사내 규정·기준 모음)
 *
 * [회사기준]      번호 | 분류 | 제목 | 내용 | 시행일 | 담당 | 원문 링크 | 최종수정 | 수정자 | 삭제(Y)
 * [기준개정이력]  번호 | 일시 | 수정자 | 개정 내용
 *
 * 누구나 조회, 팀장·관리자만 등록/수정/삭제 (시트 보호로도 막는다).
 */

const RULES = {
  SHEET: '회사기준',
  HISTORY: '기준개정이력',
  CATEGORIES: ['인사·근태', '급여·수당', '복리후생', '출장·경비', '교육', '기타'],
};

const RCOL = { ID: 1, CATEGORY: 2, TITLE: 3, BODY: 4, EFFECTIVE: 5, OWNER: 6, LINK: 7, UPDATED: 8, EDITOR: 9, DELETED: 10, FILES: 11 };

function setupRuleSheets_(ss) {
  if (!ss.getSheetByName(RULES.SHEET)) {
    const sh = ss.insertSheet(RULES.SHEET);
    sh.getRange(1, 1, 1, 10).setValues([['번호', '분류', '제목', '내용', '시행일', '담당', '원문 링크', '최종수정', '수정자', '삭제(Y)']]);
    styleHeader_(sh.getRange(1, 1, 1, 10));
    sh.getRange('B2:B').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(RULES.CATEGORIES).build());
    sh.getRange('E2:E').setNumberFormat('yyyy-mm-dd');
    sh.getRange('H2:H').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('D2:D').setWrap(true);
    [50, 90, 220, 480, 100, 80, 200, 130, 80, 70].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
  }
  ensureHeader_(ss.getSheetByName(RULES.SHEET), RCOL.FILES, '첨부');
  if (!ss.getSheetByName(RULES.HISTORY)) {
    const sh = ss.insertSheet(RULES.HISTORY);
    sh.getRange(1, 1, 1, 4).setValues([['번호', '일시', '수정자', '개정 내용']]);
    styleHeader_(sh.getRange(1, 1, 1, 4));
    sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
    [50, 130, 80, 480].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
  }
}

function readRules_(ss) {
  const sh = ss.getSheetByName(RULES.SHEET);
  if (!sh || sh.getLastRow() < 2) return [];

  const history = {};
  const hsh = ss.getSheetByName(RULES.HISTORY);
  if (hsh && hsh.getLastRow() >= 2) {
    hsh.getRange(2, 1, hsh.getLastRow() - 1, 4).getValues().forEach(r => {
      const id = String(r[0]);
      if (!id) return;
      (history[id] = history[id] || []).push({
        date: r[1] instanceof Date ? fmt(r[1], 'yyyy-MM-dd') : String(r[1]).slice(0, 10),
        editor: String(r[2]),
        note: String(r[3]),
      });
    });
  }

  return sh.getRange(2, 1, sh.getLastRow() - 1, RCOL.FILES).getValues()
    .filter(r => String(r[0]) && String(r[RCOL.DELETED - 1]).toUpperCase() !== 'Y')
    .map(r => ({
      id: String(r[0]),
      category: String(r[1]),
      title: String(r[2]),
      body: String(r[3]),
      effective: r[4] instanceof Date ? ymd(r[4]) : String(r[4] || ''),
      owner: String(r[5] || ''),
      link: String(r[6] || ''),
      updated: r[7] instanceof Date ? fmt(r[7], 'yyyy-MM-dd') : String(r[7] || '').slice(0, 10),
      editor: String(r[8] || ''),
      files: filesFromCell_(r[RCOL.FILES - 1]),
      history: (history[String(r[0])] || []).reverse(),
    }));
}

/**
 * 등록/수정. rule: { id?, category, title, body, effective:'yyyy-mm-dd', owner, link, changeNote }
 */
function apiSaveRule(rule) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'rules');
  if (!me.isLeader && !me.isAdmin) throw new Error('회사 기준은 팀장 또는 관리자만 등록·수정할 수 있습니다.');

  const title = String(rule.title || '').trim();
  const body = String(rule.body || '').trim();
  if (!title || !body) throw new Error('제목과 내용을 입력해 주세요.');
  const category = RULES.CATEGORIES.indexOf(rule.category) >= 0 ? rule.category : '기타';
  const effective = parseDue_(rule.effective);
  const editor = me.name || me.email;
  const now = new Date();

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sh = ctx.ss.getSheetByName(RULES.SHEET);
    const hsh = ctx.ss.getSheetByName(RULES.HISTORY);
    const last = sh.getLastRow();
    const ids = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(([v]) => String(v)) : [];
    const values = [category, title, body, effective, String(rule.owner || '').trim(), String(rule.link || '').trim(), now, editor];

    let id = String(rule.id || '');
    const idx = id ? ids.indexOf(id) : -1;
    if (idx >= 0) {
      sh.getRange(idx + 2, RCOL.CATEGORY, 1, values.length).setValues([values]);
      sh.getRange(idx + 2, RCOL.FILES).setValue(filesToCell_(rule.files));
      hsh.appendRow([id, now, editor, String(rule.changeNote || '').trim() || '내용 수정']);
    } else {
      id = String((ids.length ? Math.max.apply(null, ids.map(Number).filter(n => !isNaN(n))) : 0) + 1);
      sh.getRange(last + 1, 1, 1, RCOL.FILES).setValues([[id].concat(values, ['', filesToCell_(rule.files)])]);
      hsh.appendRow([id, now, editor, String(rule.changeNote || '').trim() || '최초 등록']);
    }
  } finally {
    lock.releaseLock();
  }
  return refreshPart_('rules');
}

function apiDeleteRule(id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'rules');
  if (!me.isLeader && !me.isAdmin) throw new Error('팀장 또는 관리자만 삭제할 수 있습니다.');
  const sh = ctx.ss.getSheetByName(RULES.SHEET);
  const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(([v]) => String(v));
  const idx = ids.indexOf(String(id));
  if (idx < 0) throw new Error('기준을 찾을 수 없습니다.');
  sh.getRange(idx + 2, RCOL.DELETED).setValue('Y');
  ctx.ss.getSheetByName(RULES.HISTORY).appendRow([String(id), new Date(), me.name || me.email, '삭제']);
  return refreshPart_('rules');
}

// ===================================================================
// Personnel.gs
// ===================================================================

/**
 * 인사현황 (주간보고 아래 표)
 *
 *  구 분 | 금주: 대상 및 내용 | 일자/기간 | 차주: 대상 및 내용 | 발령일 | 비고
 *
 * [인사현황]  번호 | 등록일시 | 등록자 | 주차(월요일) | 구분 | 대상 및 내용 | 시작일 | 종료일 | 비고 | 삭제(Y) | 캘린더ID
 * 등록하면 팀 캘린더에도 "구분] 대상 및 내용" 종일 일정으로 만든다 (고치면 같이 고치고, 지우면 같이 지운다).
 * 주차는 일자(시작일)가 속한 주의 월요일로 정한다 → 날짜만 넣으면 그 주에 맞게 들어가고,
 * 다음 주가 되면 "차주"였던 항목이 자동으로 "금주"로 보인다. 더 뒤의 날짜는 미리 등록해 두면 그 주가 됐을 때 나타난다.
 */

const HR = {
  SHEET: '인사현황',
  CATEGORIES: ['입사', '휴직', '복직', '퇴사', '인터뷰', '기타'],
  MIN_ROWS: 1, // 빈 줄 없이 등록한 만큼만
};

const HCOL = { ID: 1, DATE: 2, AUTHOR: 3, WEEK: 4, CATEGORY: 5, TARGET: 6, START: 7, END: 8, MEMO: 9, DELETED: 10, EVENT: 11, REC_EVENT: 12, SOURCE: 13 };
// 채용 캘린더에도 같이 넣는 구분
const HR_RECRUIT_CATS = ['인터뷰', '입사'];

function setupHrSheet_(ss) {
  if (ss.getSheetByName(HR.SHEET)) return;
  const sh = ss.insertSheet(HR.SHEET);
  sh.getRange(1, 1, 1, HCOL.EVENT).setValues([['번호', '등록일시', '등록자', '주차(월요일)', '구분', '대상 및 내용', '시작일', '종료일', '비고', '삭제(Y)', '캘린더ID']]);
  styleHeader_(sh.getRange(1, 1, 1, HCOL.EVENT));
  sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange('D2:D').setNumberFormat('@');
  sh.getRange('G2:H').setNumberFormat('yyyy-mm-dd (ddd)');
  sh.getRange('E2:E').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(HR.CATEGORIES).build());
  [60, 130, 80, 100, 70, 280, 120, 120, 220, 70].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);
}

/** 이번 주·다음 주 항목 + 그 뒤에 미리 등록한 항목(later) */
function readHr_(ctx) {
  const sh = ctx.ss.getSheetByName(HR.SHEET);
  const out = { thisWeek: [], nextWeek: [], later: [] };
  if (!sh || sh.getLastRow() < 2) return out;
  const thisKey = ctx.thisWeek.key, nextKey = ctx.nextWeek.key;
  sh.getRange(2, 1, sh.getLastRow() - 1, HCOL.SOURCE).getValues().forEach(r => {
    if (!String(r[0]) || String(r[HCOL.DELETED - 1]).toUpperCase() === 'Y') return;
    const start = r[HCOL.START - 1], end = r[HCOL.END - 1];
    // 주차: 일자가 있으면 그 날짜의 주, 없으면 (예전에 금주/차주를 골라 저장한) 주차 칸
    const week = start instanceof Date ? ymd(mondayOf(start))
      : r[HCOL.WEEK - 1] instanceof Date ? ymd(r[HCOL.WEEK - 1]) : String(r[HCOL.WEEK - 1]).trim();
    const list = week === thisKey ? out.thisWeek : week === nextKey ? out.nextWeek : week > nextKey ? out.later : null;
    if (!list) return;
    list.push({
      id: String(r[0]),
      author: String(r[HCOL.AUTHOR - 1]),
      category: String(r[HCOL.CATEGORY - 1]),
      target: String(r[HCOL.TARGET - 1]),
      start: start instanceof Date ? ymd(start) : '',
      end: end instanceof Date ? ymd(end) : '',
      when: hrPeriod_(start, end),
      memo: String(r[HCOL.MEMO - 1] || ''),
      imported: hrImported_(r), // 채용 캘린더에서 가져온 항목
      week: week,
      weekLabel: list === out.later ? getWeekInfo(mondayOf(start), ctx.holidays).label : '',
    });
  });
  const byDate = (a, b) => (a.start || '9').localeCompare(b.start || '9') || Number(a.id) - Number(b.id);
  out.thisWeek.sort(byDate);
  out.nextWeek.sort(byDate);
  out.later.sort(byDate);
  return out;
}

/** 9/28(월) · 10/1(목)~14(수) · 9/30(수)~10/2(금) */
function hrPeriod_(start, end) {
  if (!(start instanceof Date)) return '';
  const md = d => `${d.getMonth() + 1}/${d.getDate()}(${DAY_KO[d.getDay()]})`;
  if (!(end instanceof Date) || ymd(end) === ymd(start)) return md(start);
  const tail = end.getMonth() === start.getMonth() && end.getFullYear() === start.getFullYear()
    ? `${end.getDate()}(${DAY_KO[end.getDay()]})` : md(end);
  return `${md(start)}~${tail}`;
}

/** 구분별로 금주·차주를 나란히 놓은 표 줄: [{category, rows:[{a, b}]}]. 등록한 항목이 없는 구분은 뺀다. */
function hrTable_(hr) {
  const used = HR.CATEGORIES.concat(hr.thisWeek.concat(hr.nextWeek).map(x => x.category))
    .filter((c, i, all) => all.indexOf(c) === i)
    .filter(c => hr.thisWeek.concat(hr.nextWeek).some(x => x.category === c));
  return used.map(cat => {
    const a = hr.thisWeek.filter(x => x.category === cat);
    const b = hr.nextWeek.filter(x => x.category === cat);
    const n = Math.max(HR.MIN_ROWS, a.length, b.length);
    const rows = [];
    for (let i = 0; i < n; i++) rows.push({ a: a[i] || null, b: b[i] || null });
    return { category: cat, rows: rows };
  });
}

/** 등록·수정. entry: { id?, category, target, start, end, memo } — 주차는 start 날짜로 정한다 */
function apiSaveHr(entry) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'weekly');
  if (!me.name) throw new Error('[팀원] 시트에 등록된 사람만 인사현황을 등록할 수 있습니다.');

  const category = HR.CATEGORIES.indexOf(entry.category) >= 0 ? entry.category : '기타';
  const target = String(entry.target || '').trim();
  if (!target) throw new Error('대상 및 내용을 입력해 주세요.');
  const start = parseDue_(entry.start);
  const end = parseDue_(entry.end);
  if (!(start instanceof Date)) throw new Error('일자를 골라 주세요. 날짜에 맞는 주에 들어갑니다.');
  if (end && !(end instanceof Date)) throw new Error('종료일 형식이 올바르지 않습니다.');
  if (start instanceof Date && end instanceof Date && end < start) throw new Error('종료일이 시작일보다 빠릅니다.');
  const week = ymd(mondayOf(start)); // 날짜가 속한 주
  const memo = String(entry.memo || '').trim();

  const sh = ctx.ss.getSheetByName(HR.SHEET) || (setupHrSheet_(ctx.ss), ctx.ss.getSheetByName(HR.SHEET));
  ensureHeader_(sh, HCOL.EVENT, '캘린더ID');
  ensureHeader_(sh, HCOL.REC_EVENT, '채용캘린더ID');
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  let row, oldEvent = '', oldRec = '', imported = false;
  try {
    const last = sh.getLastRow();
    const ids = last >= 2 ? sh.getRange(2, 1, last - 1, HCOL.SOURCE).getValues() : [];
    if (entry.id) {
      const idx = ids.findIndex(r => String(r[0]) === String(entry.id));
      if (idx < 0) throw new Error('항목을 찾을 수 없습니다.');
      if (!canEditHrRow_(me, ids[idx])) throw new Error('본인이 등록한 항목만 고칠 수 있습니다.');
      row = idx + 2;
      oldEvent = String(ids[idx][HCOL.EVENT - 1] || '');
      oldRec = String(ids[idx][HCOL.REC_EVENT - 1] || '');
      imported = hrImported_(ids[idx]);
      if (imported && HR_RECRUIT_CATS.indexOf(category) < 0) throw new Error('채용 캘린더에서 가져온 항목은 입사·인터뷰로만 둘 수 있습니다. 구분을 바꾸려면 채용 캘린더에서 제목을 고쳐 주세요.');
      sh.getRange(row, HCOL.WEEK, 1, 6).setValues([[week, category, target, start, end, memo]]);
    } else {
      const id = ids.reduce((m, r) => Math.max(m, Number(r[0]) || 0), 0) + 1;
      row = last + 1;
      sh.getRange(row, HCOL.WEEK).setNumberFormat('@');
      sh.getRange(row, 1, 1, HCOL.DELETED).setValues([[id, new Date(), me.name, week, category, target, start, end, memo, '']]);
    }
  } finally {
    lock.releaseLock();
  }
  // 팀 캘린더에도 같은 내용으로 (권한이 없으면 인사현황에만 저장)
  // 채용 캘린더에서 가져온 항목은 팀 캘린더에 새로 만들지 않는다 (팀 캘린더에 이미 같은 일정을 넣어 두는 경우가 많아 두 번 생기지 않게)
  const ev = imported && !oldEvent ? { id: '' } : hrEvent_(me, oldEvent, category, target, start, end, memo);
  if (ev.id !== oldEvent) sh.getRange(row, HCOL.EVENT).setValue(ev.id);
  // 인터뷰·입사는 채용 캘린더에도 (채용 캘린더가 정해져 있을 때만, 권한이 없으면 조용히 건너뛴다)
  const rec = hrRecruitEvent_(ctx, me, oldRec, category, target, start, end, memo, imported);
  if (rec !== oldRec) sh.getRange(row, HCOL.REC_EVENT).setValue(rec);
  const warning = [ev.warning, ctx._recWarn].filter(Boolean).join(' ');
  return Object.assign(refreshHr_(ctx), warning ? { warning: warning } : {});
}

/** 인사현황 1건 → 팀 캘린더 종일 일정 "구분] 대상 및 내용". 이미 있으면 고친다. { id, warning } */
function hrEvent_(me, eventId, category, target, start, end, memo) {
  if (allowedMenus_(me).indexOf('calendar') < 0) return { id: eventId };
  const title = hrLabel_(category, target);
  const last = end instanceof Date ? end : start;
  const desc = [memo, `인사현황 (등록 ${me.name})`].filter(Boolean).join('\n');
  try {
    const cal = getTeamCalendar_();
    const ev = eventId ? cal.getEventById(eventId) : null;
    if (ev) {
      ev.setTitle(title);
      ev.setAllDayDates(start, addDays(last, 1));
      ev.setDescription(desc);
      bumpCalendar_();
      return { id: eventId };
    }
    const created = cal.createAllDayEvent(title, start, addDays(last, 1), { description: desc });
    bumpCalendar_();
    return { id: created.getId() };
  } catch (e) {
    return { id: eventId, warning: '팀 캘린더에 일정을 만들 권한이 없어 인사현황에만 저장했습니다.' };
  }
}

/** 인사현황 → 채용 캘린더 일정 (인터뷰·입사만). 구분이 바뀌어 대상이 아니면 지운다. 새 일정 ID(없으면 '') */
function hrRecruitEvent_(ctx, me, eventId, category, target, start, end, memo, imported) {
  if (!recruitCalendarId_(ctx.cfg)) return eventId;
  try {
    const cal = getRecruitCalendar_(ctx);
    const ev = eventId ? cal.getEventById(eventId) : null;
    // 캘린더에서 가져온 일정: 원래 제목 형식("** 인터뷰] …")·시간·메모는 그대로 두고, 바뀐 내용과 날짜만 고친다
    if (ev && imported) {
      const title = recruitRetitle_(ev.getTitle(), category, target);
      if (title !== ev.getTitle()) ev.setTitle(title);
      if (ev.isAllDayEvent()) {
        ev.setAllDayDates(start, addDays(end instanceof Date ? end : start, 1));
      } else {
        const s0 = ev.getStartTime(), e0 = ev.getEndTime();
        const s1 = new Date(start.getFullYear(), start.getMonth(), start.getDate(), s0.getHours(), s0.getMinutes());
        if (s1.getTime() !== s0.getTime()) ev.setTime(s1, new Date(s1.getTime() + (e0 - s0)));
      }
      bumpCalendar_();
      return eventId;
    }
    if (HR_RECRUIT_CATS.indexOf(category) < 0) {
      if (ev) { ev.deleteEvent(); bumpCalendar_(); }
      return '';
    }
    const title = hrLabel_(category, target);
    const last = end instanceof Date ? end : start;
    const desc = [memo, `인사현황 (등록 ${me.name})`].filter(Boolean).join('\n');
    if (ev) {
      ev.setTitle(title);
      ev.setAllDayDates(start, addDays(last, 1));
      ev.setDescription(desc);
      bumpCalendar_();
      return eventId;
    }
    const created = cal.createAllDayEvent(title, start, addDays(last, 1), { description: desc });
    bumpCalendar_();
    return created.getId();
  } catch (e) {
    // 채용 캘린더 권한이 없는 사람이 등록한 경우 등 → 인사현황·팀 캘린더에는 그대로 두고 알린다
    ctx._recWarn = /찾을 수 없|접근할 수 없/.test(String(e && e.message))
      ? '채용 캘린더를 찾을 수 없어 팀 캘린더에만 넣었습니다. 관리 → 설정의 채용 캘린더 ID와 공유를 확인해 주세요.'
      : '채용 캘린더에 일정을 넣을 권한이 없어 팀 캘린더에만 넣었습니다. 채용 캘린더를 "일정 변경" 권한으로 공유받아야 합니다.';
    return eventId;
  }
}

/** 본인 등록분 · 팀장 · 관리자. 채용 캘린더에서 가져온 항목은 채용담당도 */
function canEditHrRow_(me, r) {
  return r[HCOL.AUTHOR - 1] === me.name || me.isLeader || me.isAdmin || (hrImported_(r) && !!me.recruiter);
}

/** 채용 캘린더에서 가져온 줄인지 (출처 칸, 예전에는 등록자 칸이 "채용캘린더") */
function hrImported_(r) {
  return String(r[HCOL.SOURCE - 1] || '') === RECRUIT_AUTHOR || String(r[HCOL.AUTHOR - 1] || '') === RECRUIT_AUTHOR;
}

/** 채용 캘린더 일정을 만든 사람이 채용담당이면 그 이름, 아니면 '' (담당자가 넣은 일정만 가져온다) */
function recruitCreator_(ev, ctx) {
  let emails = [];
  try { emails = (ev.getCreators() || []).map(x => String(x).toLowerCase()); } catch (e) { return ''; }
  const m = ctx.members.find(x => x.recruit && x.email && emails.indexOf(x.email.toLowerCase()) >= 0);
  return m ? m.name : '';
}

/* ---------- 채용 캘린더 → 인사현황 ----------
 * 채용담당이 채용 캘린더에 직접 넣은 입사·인터뷰(면접) 일정을 인사현황으로 가져온다. (등록자: 그 일정을 넣은 채용담당, 출처: "채용캘린더")
 *  - 채용담당이 아닌 사람이 넣은 일정은 가져오지 않는다 (구글 캘린더의 "만든 사람" 기준)
 *  - 제목이 "[입사] …", "[인터뷰] …", "… 면접 …"처럼 입사·인터뷰·면접이 들어간 일정만
 *  - 캘린더에서 날짜·제목을 고치면 따라 바뀌고, 지우면 인사현황에서도 지워진다
 *  - 웹앱에서 등록해 채용 캘린더로 보낸 항목은 날짜가 바뀐 것만 맞춘다
 * 채용담당이 앱을 열 때와 매시간 자동 실행 때 맞춘다.
 */
const RECRUIT_AUTHOR = '채용캘린더';
const RECRUIT_SYNC_DAYS = 90; // 이번 주 월요일부터 석 달

function recruitCategory_(title) {
  const t = String(title || '').replace(/^\s*[가-힣A-Za-z]{2,10}\)\s*/, '');
  const head = (/^\s*\[([^\]]+)\]/.exec(t) || [])[1] || '';
  const pick = s => /입사(?!지원|서류)/.test(s) ? '입사' : /인터뷰|면접/.test(s) ? '인터뷰' : '';
  return pick(head) || pick(t);
}

/**
 * 제목에서 구분 표시를 뺀 나머지 → 인사현황 "대상 및 내용"
 *  "정하은) [인터뷰] 경력직 후보 A" → "경력직 후보 A"
 *  "** 인터뷰] 경력직 후보 A" → "경력직 후보 A"  ( ']'만 붙인 형식, 앞의 기호는 버린다)
 *  "홍길동 입사] 재무팀" → "홍길동 재무팀"
 */
function recruitTarget_(title) {
  const t = String(title || '').trim();
  let s = t.replace(/^\s*[가-힣A-Za-z]{2,10}\)\s*/, '').replace(/^\s*\[[^\]]*\]\s*/, '');
  const m = /^([^\[\]]*?)\s*(?:입사|인터뷰|면접)\s*\]\s*(.*)$/.exec(s);
  if (m) s = [m[1].replace(/^[^0-9A-Za-z가-힣]+/, '').trim(), m[2].trim()].filter(Boolean).join(' ');
  return s.trim() || t;
}

function syncRecruitHr_(ctx) {
  if (!recruitCalendarId_(ctx.cfg)) return false;
  const from = ctx.thisWeek.monday, to = addDays(from, RECRUIT_SYNC_DAYS);
  let events;
  try { events = getRecruitCalendar_(ctx).getEvents(from, to); } catch (e) { return false; }

  const sh = ctx.ss.getSheetByName(HR.SHEET) || (setupHrSheet_(ctx.ss), ctx.ss.getSheetByName(HR.SHEET));
  ensureHeader_(sh, HCOL.EVENT, '캘린더ID');
  ensureHeader_(sh, HCOL.REC_EVENT, '채용캘린더ID');
  ensureHeader_(sh, HCOL.SOURCE, '출처');
  const day = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const hm = d => `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
  const key = v => (v instanceof Date ? ymd(v) : '');

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  let changed = false;
  try {
    const last = sh.getLastRow();
    const rows = last >= 2 ? sh.getRange(2, 1, last - 1, HCOL.SOURCE).getValues() : [];
    const linked = {};
    rows.forEach((r, i) => {
      const id = String(r[HCOL.REC_EVENT - 1] || '');
      if (id && String(r[HCOL.DELETED - 1]).toUpperCase() !== 'Y') linked[id] = i;
    });
    let nextId = rows.reduce((m, r) => Math.max(m, Number(r[0]) || 0), 0);
    const seen = {};
    const add = [];
    events.forEach(e => {
      const cat = recruitCategory_(e.getTitle());
      const id = e.getId();
      if (!cat || seen[id]) return; // 반복 일정은 첫 번만
      const creator = recruitCreator_(e, ctx);
      if (!creator) return; // 채용담당이 넣은 일정만 (다른 사람이 넣은 것은 가져오지 않는다)
      seen[id] = true;
      const start = day(e.getStartTime());
      let end = e.isAllDayEvent() ? addDays(day(e.getEndTime()), -1) : day(e.getEndTime());
      if (end <= start) end = '';
      const memo = [e.isAllDayEvent() ? '' : `${hm(e.getStartTime())}~${hm(e.getEndTime())}`, e.getLocation()].filter(Boolean).join(' · ');
      const target = recruitTarget_(e.getTitle());
      const week = ymd(mondayOf(start));
      const i = linked[id];
      if (i == null) {
        add.push([String(++nextId), new Date(), creator, week, cat, target, start, end, memo, '', '', id, RECRUIT_AUTHOR]);
        return;
      }
      const r = rows[i];
      const row = i + 2;
      if (hrImported_(r)) {
        // 캘린더에서 가져온 항목: 캘린더가 기준 (등록자 = 그 일정을 넣은 채용담당)
        if (r[HCOL.AUTHOR - 1] !== creator || String(r[HCOL.SOURCE - 1] || '') !== RECRUIT_AUTHOR) {
          sh.getRange(row, HCOL.AUTHOR).setValue(creator);
          sh.getRange(row, HCOL.SOURCE).setValue(RECRUIT_AUTHOR);
          changed = true;
        }
        if (r[HCOL.CATEGORY - 1] !== cat || r[HCOL.TARGET - 1] !== target || key(r[HCOL.START - 1]) !== ymd(start)
          || key(r[HCOL.END - 1]) !== key(end) || String(r[HCOL.MEMO - 1] || '') !== memo) {
          sh.getRange(row, HCOL.WEEK, 1, 6).setValues([[week, cat, target, start, end, memo]]);
          changed = true;
        }
      } else if (key(r[HCOL.START - 1]) !== ymd(start)) {
        // 웹앱에서 등록한 항목: 캘린더에서 날짜를 옮긴 것만 따라간다
        sh.getRange(row, HCOL.WEEK).setValue(week);
        sh.getRange(row, HCOL.START, 1, 2).setValues([[start, end]]);
        changed = true;
      }
    });
    if (add.length) {
      sh.getRange(last + 1, HCOL.WEEK, add.length, 1).setNumberFormat('@');
      sh.getRange(last + 1, 1, add.length, HCOL.SOURCE).setValues(add);
      changed = true;
    }
    // 캘린더에서 지운 일정 → 가져온 항목도 지운다 (맞추는 기간 안의 것만)
    rows.forEach((r, i) => {
      const id = String(r[HCOL.REC_EVENT - 1] || '');
      const st = r[HCOL.START - 1];
      if (!id || seen[id] || !hrImported_(r) || String(r[HCOL.DELETED - 1]).toUpperCase() === 'Y') return;
      if (!(st instanceof Date) || st < from || st >= to) return;
      sh.getRange(i + 2, HCOL.DELETED).setValue('Y');
      changed = true;
    });
  } finally {
    lock.releaseLock();
  }
  if (changed) refreshHr_(ctx);
  return changed;
}

/** 가져온 일정의 제목에서 "대상" 부분만 바꾼다. 형식을 알아볼 수 없으면 "구분] 대상" */
function recruitRetitle_(old, category, target) {
  const cur = recruitTarget_(old);
  if (cur === target) return old;
  const keep = old.lastIndexOf(cur) >= 0 && old.endsWith(cur) ? old.slice(0, old.length - cur.length) : '';
  const next = keep && /\S/.test(keep) ? (keep.endsWith(' ') || keep.endsWith(']') ? keep : keep + ' ') + target : '';
  // 새 제목에서도 같은 구분으로 읽혀야 다음 맞추기에서 사라지지 않는다
  return next && recruitCategory_(next) === category ? next.replace(/\]\S/, m => m[0] + ' ' + m[1]) : hrLabel_(category, target);
}

/* ---------- 팀 캘린더 → 인사현황 ----------
 * 웹앱 인사현황이 팀 캘린더에 만든 "구분] 대상" 일정을 구글 캘린더에서 고치면(제목·날짜) 인사현황도 따라간다.
 * 캘린더에서 지우면 인사현황에서도 지운다. 이번 주 월요일부터 석 달을 맞춘다.
 */
function syncTeamHr_(ctx) {
  if (!teamCalendarSet_(ctx)) return false;
  const sh = ctx.ss.getSheetByName(HR.SHEET);
  if (!sh || sh.getLastRow() < 2) return false;
  const from = ctx.thisWeek.monday, to = addDays(from, RECRUIT_SYNC_DAYS);
  let cal, events;
  try { cal = getTeamCalendar_(); events = cal.getEvents(from, to); } catch (e) { return false; }
  const byId = {};
  events.forEach(e => { byId[e.getId()] = e; });
  const key = v => (v instanceof Date ? ymd(v) : '');
  const width = Math.max(HCOL.EVENT, Math.min(HCOL.SOURCE, sh.getLastColumn()));
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
  const changes = [];
  rows.forEach((r, i) => {
    const id = String(r[HCOL.EVENT - 1] || '');
    if (!id || String(r[HCOL.DELETED - 1]).toUpperCase() === 'Y') return;
    const st = r[HCOL.START - 1];
    let ev = byId[id];
    if (!ev) {
      if (!(st instanceof Date) || st < from || st >= to) return; // 읽은 범위 밖은 건드리지 않는다
      try { ev = cal.getEventById(id); } catch (e) { ev = null; } // 범위 밖으로 옮겼는지 확인
      // 읽은 범위 안에 있다면 getEvents에 나왔어야 한다 → 안 나왔으면 지운 일정으로 본다
      try { if (ev && ev.getStartTime() >= from && ev.getStartTime() < to) ev = null; } catch (e) { ev = null; }
      if (!ev) { changes.push({ row: i + 2, del: true, r: r }); return; }
    }
    const title = String(ev.getTitle() || '').trim();
    const m = /^\s*\[?\s*([가-힣]+)\s*\]\s*(.*)$/.exec(title); // "입사] 대상" · 예전 "[입사] 대상"
    const category = m && HR.CATEGORIES.indexOf(m[1].trim()) >= 0 ? m[1].trim() : String(r[HCOL.CATEGORY - 1]);
    const target = (m ? m[2] : title).trim() || String(r[HCOL.TARGET - 1]);
    const allDay = ev.isAllDayEvent();
    const start = allDay ? ev.getAllDayStartDate() : startOfDay_(ev.getStartTime());
    let end = allDay ? addDays(ev.getAllDayEndDate(), -1) : startOfDay_(ev.getEndTime());
    if (end <= start) end = '';
    if (String(r[HCOL.CATEGORY - 1]) !== category || String(r[HCOL.TARGET - 1]) !== target
      || key(st) !== ymd(start) || key(r[HCOL.END - 1]) !== key(end)) {
      changes.push({ row: i + 2, r: r, category: category, target: target, start: start, end: end });
    }
  });
  if (!changes.length) return false;

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    changes.forEach(c => {
      const rec = String(c.r[HCOL.REC_EVENT - 1] || '');
      const author = String(c.r[HCOL.AUTHOR - 1] || '');
      if (c.del) {
        sh.getRange(c.row, HCOL.DELETED).setValue('Y');
        if (rec && recruitCalendarId_(ctx.cfg)) {
          try { const ev = getRecruitCalendar_(ctx).getEventById(rec); if (ev) ev.deleteEvent(); } catch (e) { /* 권한 없음 */ }
        }
        return;
      }
      sh.getRange(c.row, HCOL.WEEK, 1, 5).setValues([[ymd(mondayOf(c.start)), c.category, c.target, c.start, c.end]]);
      // 채용 캘린더에도 같이 걸려 있으면 그쪽도 맞춘다 (서로 되돌리지 않도록)
      if (rec || HR_RECRUIT_CATS.indexOf(c.category) >= 0) {
        const next = hrRecruitEvent_(ctx, { name: author }, rec, c.category, c.target, c.start, c.end instanceof Date ? c.end : '', String(c.r[HCOL.MEMO - 1] || ''), hrImported_(c.r));
        if (next !== rec && sh.getLastColumn() >= HCOL.EVENT) sh.getRange(c.row, HCOL.REC_EVENT).setValue(next);
      }
    });
  } finally {
    lock.releaseLock();
  }
  bumpCalendar_();
  refreshHr_(ctx);
  return true;
}

/** 인사현황 일정 제목 — 팀 양식 "입사] 대상" */
function hrLabel_(category, target) {
  return `${category}] ${target}`;
}

function apiDeleteHr(id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'weekly');
  const sh = ctx.ss.getSheetByName(HR.SHEET);
  if (!sh || sh.getLastRow() < 2) throw new Error('항목을 찾을 수 없습니다.');
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, HCOL.SOURCE).getValues();
  const idx = rows.findIndex(r => String(r[0]) === String(id));
  if (idx < 0) throw new Error('항목을 찾을 수 없습니다.');
  if (!canEditHrRow_(me, rows[idx])) throw new Error('본인이 등록한 항목만 삭제할 수 있습니다.');
  sh.getRange(idx + 2, HCOL.DELETED).setValue('Y');
  // 팀 캘린더 일정도 지운다
  const eventId = sh.getLastColumn() >= HCOL.EVENT ? String(sh.getRange(idx + 2, HCOL.EVENT).getValue() || '') : '';
  if (eventId) {
    try {
      const ev = getTeamCalendar_().getEventById(eventId);
      if (ev) { ev.deleteEvent(); bumpCalendar_(); }
    } catch (e) { /* 캘린더 권한이 없으면 인사현황에서만 지운다 */ }
  }
  const recId = sh.getLastColumn() >= HCOL.REC_EVENT ? String(sh.getRange(idx + 2, HCOL.REC_EVENT).getValue() || '') : '';
  if (recId && recruitCalendarId_(ctx.cfg)) {
    try {
      const ev = getRecruitCalendar_(ctx).getEventById(recId);
      if (ev) { ev.deleteEvent(); bumpCalendar_(); }
    } catch (e) { /* 권한 없음 */ }
  }
  return refreshHr_(ctx);
}

function refreshHr_(ctx) {
  const hr = readHr_(ctx);
  putCache_(CK.hr(ctx), hr);
  return { hr: hr };
}

/** 주간보고 시트 아래에 인사현황 표를 붙인다. 표가 끝난 다음 행 번호를 돌려준다. */
function writeHrTable_(sh, ctx, top, table) {
  table = table || hrTable_(readHr_(ctx));
  sh.getRange(top, 1).setValue('■ 인사 현황').setFontSize(12).setFontWeight('bold');
  const h = top + 1;
  sh.getRange(h, 1, 2, 6).setValues([
    ['구 분', ctx.thisWeek.label, '', ctx.nextWeek.label, '', '비고'],
    ['', '대상 및 내용', '일자/기간', '대상 및 내용', '발령일', ''],
  ]);
  ['A', 'F'].forEach(c => sh.getRange(`${c}${h}:${c}${h + 1}`).merge());
  sh.getRange(h, 2, 1, 2).merge();
  sh.getRange(h, 4, 1, 2).merge();
  sh.getRange(h, 1, 2, 6).setBackground('#efefef').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  const values = [];
  const blocks = [];
  let row = h + 2;
  table.forEach(t => {
    t.rows.forEach((r, i) => values.push([
      i === 0 ? t.category : '',
      r.a ? r.a.target : '', r.a ? r.a.when : '',
      r.b ? r.b.target : '', r.b ? r.b.when : '',
      [r.a && r.a.memo, r.b && r.b.memo].filter(Boolean).join(' / '),
    ]));
    blocks.push({ row: row, n: t.rows.length });
    row += t.rows.length;
  });
  if (!values.length) {
    values.push(['', '이번 주·다음 주에 등록된 인사현황이 없습니다.', '', '', '', '']);
    row++;
  }
  sh.getRange(h + 2, 1, values.length, 6).setNumberFormat('@').setValues(values)
    .setVerticalAlignment('middle').setWrap(true);
  [3, 5].forEach(c => sh.getRange(h + 2, c, values.length, 1).setHorizontalAlignment('center'));
  blocks.forEach(b => sh.getRange(b.row, 1, b.n, 1).merge().setHorizontalAlignment('center').setFontWeight('bold'));
  sh.getRange(h, 1, 2 + values.length, 6)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  blocks.forEach(b => sh.getRange(b.row, 1, b.n, 6)
    .setBorder(true, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM));
  return row;
}

/* ---------- 채용 담당자 안내 메일 (매시간 자동 실행에서) ----------
 *  - 입사: 입사일 + 70일이 되는 날 → 수습평가 안내
 *  - 인터뷰: 인터뷰 전날 → 내일 인터뷰 안내
 * 받는 사람: [팀원] 시트 "채용담당(Y)" — 알림메일(회사 메일) 칸이 있으면 그 주소로.
 * 같은 항목·같은 날짜로는 한 번만 보낸다 (날짜를 고치면 새 날짜 기준으로 다시 보낸다).
 */
const HR_MAIL = {
  PROBATION_DAYS: 70,
  HOUR: 9,      // 오전 9시 이후 첫 자동 실행 때 보낸다
  LATE_DAYS: 3, // 그날 못 보냈으면 3일 안에는 보낸다 (기능을 켜기 전 지난 입사자에게 한꺼번에 가지 않도록)
};

function hrMailCheck_(ctx, now) {
  now = now || new Date();
  if (now.getHours() < HR_MAIL.HOUR) return [];
  const to = ctx.members.filter(m => m.recruit).map(mailTo_).filter(Boolean);
  const sh = ctx.ss.getSheetByName(HR.SHEET);
  if (!to.length || !sh || sh.getLastRow() < 2) return [];
  const today = startOfDay_(now);
  const props = PropertiesService.getDocumentProperties();
  const team = ctx.cfg[CFG.TEAM] || '';
  const url = ScriptApp.getService().getUrl() || ctx.ss.getUrl();
  const md = d => `${d.getMonth() + 1}/${d.getDate()}(${DAY_KO[d.getDay()]})`;
  const sent = [];

  sh.getRange(2, 1, sh.getLastRow() - 1, HCOL.DELETED).getValues().forEach(r => {
    const id = String(r[0]).trim();
    const cat = String(r[HCOL.CATEGORY - 1]).trim();
    const start = r[HCOL.START - 1];
    if (!id || String(r[HCOL.DELETED - 1]).toUpperCase() === 'Y' || !(start instanceof Date)) return;
    const target = String(r[HCOL.TARGET - 1]).trim();
    const memo = String(r[HCOL.MEMO - 1] || '').trim();
    const day = startOfDay_(start);

    let due = null, mail = null;
    if (cat === '입사') {
      due = addDays(day, HR_MAIL.PROBATION_DAYS);
      mail = {
        subject: `[${team}] 수습평가 안내 - ${target} (입사 ${HR_MAIL.PROBATION_DAYS}일)`,
        htmlBody: `채용 담당자님, 아래 입사자가 입사 ${HR_MAIL.PROBATION_DAYS}일이 되었습니다. 수습평가 진행을 준비해 주세요.<br><br>` +
          `· 대상: <b>${target}</b><br>· 입사일: ${md(day)}<br>· 입사 ${HR_MAIL.PROBATION_DAYS}일: ${md(due)}<br>` +
          (memo ? `· 비고: ${memo}<br>` : '') + `<br><a href="${url}">업무관리 열기</a>`,
      };
    } else if (cat === '인터뷰') {
      due = addDays(day, -1);
      mail = {
        subject: `[${team}] 내일 인터뷰 안내 - ${target} (${md(day)})`,
        htmlBody: `채용 담당자님, 내일 인터뷰 일정이 있습니다.<br><br>` +
          `· 내용: <b>${target}</b><br>· 일자: ${hrPeriod_(start, r[HCOL.END - 1])}<br>` +
          (memo ? `· 비고: ${memo}<br>` : '') + `<br><a href="${url}">업무관리 열기</a>`,
      };
    } else return;

    if (today < due || today > addDays(due, cat === '인터뷰' ? 0 : HR_MAIL.LATE_DAYS)) return; // 인터뷰는 전날에만
    const key = `HRMAIL:${cat}:${id}:${ymd(day)}`;
    if (props.getProperty(key)) return;
    try {
      MailApp.sendEmail(Object.assign({ to: to.join(',') }, mail));
      props.setProperty(key, ymd(today));
      sent.push(`${cat} ${target}`);
    } catch (e) { /* 메일 한도 등 → 다음 실행 때 다시 */ }
  });
  return sent;
}

// ===================================================================
// Minutes.gs
// ===================================================================

/**
 * 회의록 (주간업무보고 > 회의록 탭)
 *
 * [회의록]  번호 | 일자 | 시간 | 장소 | 참석자 | 방향성 | 항목(JSON) | 수정자 | 수정일시 | 삭제(Y)
 * 항목: [{ cat: 구분, sub: 세부, content: 내용, owner: 담당(여러 명은 쉼표), due: 기한(yyyy-mm-dd / - / 직접 입력), status: 진행상태 }]
 * 새 회의록은 지난 회의록을 이어받아 시작한다 (완료한 항목은 빼고).
 */

const MINUTES = {
  SHEET: '회의록',
  STATUS: ['진행예정', '진행중', '완료', '보류', '-'],
  MAX_JSON: 45000, // 시트 한 칸은 5만 자까지
};

const MCOL = { ID: 1, DATE: 2, TIME: 3, PLACE: 4, ATTEND: 5, VISION: 6, ITEMS: 7, EDITOR: 8, UPDATED: 9, DELETED: 10 };

function setupMinutesSheet_(ss) {
  if (ss.getSheetByName(MINUTES.SHEET)) return ss.getSheetByName(MINUTES.SHEET);
  const sh = ss.insertSheet(MINUTES.SHEET);
  sh.getRange(1, 1, 1, MCOL.DELETED).setValues([['번호', '일자', '시간', '장소', '참석자', '방향성', '항목(JSON)', '수정자', '수정일시', '삭제(Y)']]);
  styleHeader_(sh.getRange(1, 1, 1, MCOL.DELETED));
  sh.getRange('B2:C').setNumberFormat('@');
  sh.getRange('I2:I').setNumberFormat('yyyy-mm-dd hh:mm');
  [60, 100, 70, 140, 200, 300, 400, 80, 130, 70].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);
  return sh;
}

/** 회의록 목록 (최근 회의가 앞) */
function readMinutes_(ss) {
  const sh = ss.getSheetByName(MINUTES.SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, MCOL.DELETED).getValues()
    .filter(r => String(r[0]).trim() && String(r[MCOL.DELETED - 1]).toUpperCase() !== 'Y')
    .map(r => {
      let items = [];
      try { items = JSON.parse(String(r[MCOL.ITEMS - 1] || '[]')); } catch (e) { items = []; }
      const d = r[MCOL.DATE - 1];
      return {
        id: String(r[0]),
        date: d instanceof Date ? ymd(d) : String(d),
        time: String(r[MCOL.TIME - 1] || ''),
        place: String(r[MCOL.PLACE - 1] || ''),
        attendees: String(r[MCOL.ATTEND - 1] || ''),
        vision: String(r[MCOL.VISION - 1] || ''),
        items: Array.isArray(items) ? items : [],
        editor: String(r[MCOL.EDITOR - 1] || ''),
        updated: r[MCOL.UPDATED - 1] instanceof Date ? fmt(r[MCOL.UPDATED - 1], 'yyyy-MM-dd HH:mm') : '',
      };
    })
    .sort((a, b) => b.date.localeCompare(a.date) || Number(b.id) - Number(a.id));
}

/** 저장(새로 만들기·고치기). m: { id?, date, time, place, attendees, vision, items } */
function apiSaveMinutes(m) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'weekly');
  if (!me.name) throw new Error('[팀원] 시트에 등록된 사람만 회의록을 쓸 수 있습니다.');
  const date = String(m.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('회의 일자를 골라 주세요.');
  const clean = v => String(v == null ? '' : v).trim();
  const items = (m.items || []).map(x => ({
    cat: clean(x.cat), sub: clean(x.sub), content: clean(x.content), owner: clean(x.owner), due: clean(x.due),
    status: MINUTES.STATUS.indexOf(clean(x.status)) >= 0 ? clean(x.status) : '진행중',
  })).filter(x => x.cat || x.sub || x.content);
  const json = JSON.stringify(items);
  if (json.length > MINUTES.MAX_JSON) throw new Error('항목이 너무 많습니다. 끝난 항목을 정리하거나 회의록을 나눠 주세요.');
  const row = [date, clean(m.time), clean(m.place), clean(m.attendees), clean(m.vision), json, me.name, new Date(), ''];

  const sh = setupMinutesSheet_(ctx.ss);
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  let id = clean(m.id);
  try {
    const last = sh.getLastRow();
    const ids = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
    const idx = id ? ids.indexOf(id) : -1;
    if (id && idx < 0) throw new Error('회의록을 찾을 수 없습니다. 다른 사람이 지웠을 수 있습니다.');
    if (idx >= 0) {
      sh.getRange(idx + 2, MCOL.DATE, 1, row.length).setValues([row]);
    } else {
      id = String(ids.reduce((mx, v) => Math.max(mx, Number(v) || 0), 0) + 1);
      sh.getRange(last + 1, MCOL.DATE, 1, 2).setNumberFormat('@');
      sh.getRange(last + 1, 1, 1, row.length + 1).setValues([[id].concat(row)]);
    }
  } finally {
    lock.releaseLock();
  }
  return Object.assign(refreshMinutes_(ctx.ss), { savedId: id });
}

/** 삭제: 쓴 사람(마지막 수정자)·팀장·관리자 */
function apiDeleteMinutes(id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'weekly');
  const sh = ctx.ss.getSheetByName(MINUTES.SHEET);
  if (!sh || sh.getLastRow() < 2) throw new Error('회의록을 찾을 수 없습니다.');
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, MCOL.EDITOR).getValues();
  const idx = rows.findIndex(r => String(r[0]) === String(id));
  if (idx < 0) throw new Error('회의록을 찾을 수 없습니다.');
  if (rows[idx][MCOL.EDITOR - 1] !== me.name && !me.isLeader && !me.isAdmin) throw new Error('쓴 사람이나 팀장·관리자만 지울 수 있습니다.');
  sh.getRange(idx + 2, MCOL.DELETED).setValue('Y');
  return refreshMinutes_(ctx.ss);
}

function refreshMinutes_(ss) {
  const list = readMinutes_(ss);
  putCache_(CK.minutes, list);
  return { minutes: list };
}

// ===================================================================
// Leave.gs
// ===================================================================

/**
 * 휴가·근태 → 주간보고 "휴무계획/특이사항" 자동 반영
 *
 * [휴가]  번호 | 등록일시 | 이름 | 종류 | 시작일 | 종료일 | 메모 | 캘린더 일정ID | 출처(웹앱/캘린더) | 삭제(Y)
 *
 * - 웹앱 일정 메뉴에서 휴가를 등록하면 이 시트에 기록하고 팀 캘린더에도 "[연차] 이름" 일정을 만든다.
 * - 구글 캘린더에 직접 "김민지 연차"처럼 넣은 일정도 (제목에 팀원 이름 + 휴가 단어) 자동으로 가져온다.
 *   일정 메뉴를 열 때와 매시간 자동 실행 때 맞춘다.
 * - 이번 주·다음 주에 걸친 휴가는 그 팀원의 휴무계획/특이사항에 "10/2(금) 연차"처럼 자동으로 붙는다.
 */

const LEAVE = {
  SHEET: '휴가',
  TYPES: ['연차', '오전반차', '오후반차', '오전반반차', '오후반반차', '정기휴가', '경조휴가', '대체휴무', '출장', '교육', '기타'],
  // 캘린더 일정 제목에서 휴가로 볼 단어 (긴 것부터)
  WORDS: /(오전반반차|오후반반차|반반차|오전반차|오후반차|반차|정기휴가|경조휴가|하계휴가|여름휴가|대체휴무|연차|휴가|휴무)/,
};

const LCOL = { ID: 1, DATE: 2, NAME: 3, TYPE: 4, START: 5, END: 6, MEMO: 7, EVENT: 8, SOURCE: 9, DELETED: 10, REC_EVENT: 11 };

function setupLeaveSheet_(ss) {
  if (ss.getSheetByName(LEAVE.SHEET)) return ss.getSheetByName(LEAVE.SHEET);
  const sh = ss.insertSheet(LEAVE.SHEET);
  sh.getRange(1, 1, 1, LCOL.DELETED).setValues([['번호', '등록일시', '이름', '종류', '시작일', '종료일', '메모', '캘린더 일정ID', '출처', '삭제(Y)']]);
  styleHeader_(sh.getRange(1, 1, 1, LCOL.DELETED));
  sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange('E2:F').setNumberFormat('yyyy-mm-dd (ddd)');
  [60, 130, 80, 80, 120, 120, 200, 160, 70, 70].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);
  return sh;
}

function leaveRows_(sh) {
  return sh && sh.getLastRow() >= 2 ? sh.getRange(2, 1, sh.getLastRow() - 1, LCOL.REC_EVENT).getValues() : [];
}

/** 이번 주 월요일 ~ 다다음 주 일요일에 걸친 휴가 (요청 안에서 한 번만 읽는다) */
function leavesOf_(ctx) {
  if (ctx._leaves) return ctx._leaves;
  const from = ctx.thisWeek.monday, to = addDays(ctx.afterWeek.monday, 6);
  ctx._leaves = leaveRows_(ctx.ss.getSheetByName(LEAVE.SHEET))
    .filter(r => String(r[0]) && String(r[LCOL.DELETED - 1]).toUpperCase() !== 'Y'
      && r[LCOL.START - 1] instanceof Date)
    .map(r => {
      const start = r[LCOL.START - 1];
      const end = r[LCOL.END - 1] instanceof Date && r[LCOL.END - 1] >= start ? r[LCOL.END - 1] : start;
      return {
        id: String(r[0]),
        name: String(r[LCOL.NAME - 1]).trim(),
        type: String(r[LCOL.TYPE - 1]).trim(),
        start: start,
        end: end,
        memo: String(r[LCOL.MEMO - 1] || '').trim(),
        source: String(r[LCOL.SOURCE - 1] || ''),
      };
    })
    .filter(l => l.end >= from && l.start <= to)
    .sort((a, b) => a.start - b.start);
  return ctx._leaves;
}

/** 그 사람의 휴가 중 monday 주 ~ (monday+weeks주) 에 걸친 것 → 화면용 */
function memberLeaves_(ctx, name, monday, weeks) {
  const to = addDays(monday, 7 * weeks - 1);
  return leavesOf_(ctx)
    .filter(l => l.name === name && l.end >= monday && l.start <= to)
    .map(l => ({
      id: l.id,
      type: l.type,
      start: ymd(l.start),
      end: ymd(l.end),
      source: l.source,
      text: leaveText_(l),
    }));
}

/** "10/2(금) 연차" · "9/30(수)~10/7(수) 정기휴가" */
function leaveText_(l) {
  return `${hrPeriod_(l.start, l.end)} ${l.type}${l.memo ? `(${l.memo})` : ''}`;
}

/** 자동 휴가 문구 + 직접 쓴 특이사항 (같은 내용을 이미 적었으면 한 번만) */
function combineNote_(leaves, note) {
  const n = String(note || '').trim();
  const auto = (leaves || []).map(l => l.text).filter(t => n.indexOf(t) < 0);
  return auto.concat(n ? [n] : []).join(', ');
}

/* ---------- 등록 / 삭제 ---------- */

/** leave: { type, start:'yyyy-mm-dd', end, memo } */
function apiAddLeave(leave) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'weekly');
  if (!me.name) throw new Error('[팀원] 시트에 등록된 사람만 휴가를 등록할 수 있습니다.');

  const type = LEAVE.TYPES.indexOf(leave.type) >= 0 ? leave.type : '기타';
  const start = parseDue_(leave.start);
  const end = leave.end ? parseDue_(leave.end) : start;
  if (!(start instanceof Date)) throw new Error('시작일을 골라 주세요.');
  if (!(end instanceof Date)) throw new Error('종료일 형식이 올바르지 않습니다.');
  if (end < start) throw new Error('종료일이 시작일보다 빠릅니다.');
  const memo = String(leave.memo || '').trim();

  // 팀 캘린더에도 종일 일정으로 만든다 (권한이 없으면 주간보고에만 반영)
  let eventId = '', warning = '';
  if (allowedMenus_(me).indexOf('calendar') >= 0) {
    try {
      const ev = getTeamCalendar_().createAllDayEvent(leaveLabel_(type, shortName_(me.name, ctx.members.map(m => m.name))), start, addDays(end, 1),
        { description: [memo, `등록: ${me.name} (업무관리 웹앱)`].filter(Boolean).join('\n') });
      eventId = ev.getId();
    } catch (e) {
      warning = '팀 캘린더에 일정을 만들 권한이 없어 주간보고에만 반영했습니다.';
    }
  }
  // 채용담당은 채용 캘린더에도 같이 넣는다
  let recEventId = '';
  if (canRecruitCal_(ctx, me)) {
    try {
      recEventId = getRecruitCalendar_(ctx).createAllDayEvent(leaveLabel_(type, shortName_(me.name, ctx.members.map(m => m.name))), start, addDays(end, 1),
        { description: [memo, `등록: ${me.name} (업무관리 웹앱)`].filter(Boolean).join('\n') }).getId();
    } catch (e) {
      warning = warning || '팀 캘린더에는 넣었지만 채용 캘린더에는 권한이 없어 넣지 못했습니다.';
    }
  }

  const sh = setupLeaveSheet_(ctx.ss);
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    if (recEventId) ensureHeader_(sh, LCOL.REC_EVENT, '채용캘린더 일정ID');
    sh.getRange(sh.getLastRow() + 1, 1, 1, LCOL.REC_EVENT)
      .setValues([[nextLeaveId_(sh), new Date(), me.name, type, start, end, memo, eventId, '웹앱', '', recEventId]]);
  } finally {
    lock.releaseLock();
  }
  ctx._leaves = null;
  if (eventId) bumpCalendar_();
  return Object.assign(leavePatch_(ctx, me.name), { warning: warning });
}

function apiDeleteLeave(id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'weekly');
  const sh = ctx.ss.getSheetByName(LEAVE.SHEET);
  const rows = leaveRows_(sh);
  const idx = rows.findIndex(r => String(r[0]) === String(id));
  if (idx < 0) throw new Error('휴가를 찾을 수 없습니다.');
  const name = String(rows[idx][LCOL.NAME - 1]).trim();
  if (name !== me.name && !me.isLeader && !me.isAdmin) throw new Error('본인 휴가만 지울 수 있습니다.');

  sh.getRange(idx + 2, LCOL.DELETED).setValue('Y');
  const eventId = String(rows[idx][LCOL.EVENT - 1] || '');
  // 여러 명이 같이 적힌 일정("…#이름")은 이 사람 휴가만 지우고 캘린더 일정은 그대로 둔다
  if (eventId && eventId.indexOf('#') < 0) {
    try {
      const fromRecruit = String(rows[idx][LCOL.SOURCE - 1]) === RECRUIT_LEAVE_SOURCE;
      const ev = (fromRecruit ? getRecruitCalendar_(ctx) : getTeamCalendar_()).getEventById(eventId);
      if (ev) { ev.deleteEvent(); bumpCalendar_(); }
    } catch (e) { /* 캘린더 권한이 없으면 시트에서만 지운다 */ }
  }
  dropRecruitLeaveEvent_(ctx, rows[idx][LCOL.REC_EVENT - 1]);
  ctx._leaves = null;
  return leavePatch_(ctx, name);
}

/**
 * 일정 화면에서 휴가 일정을 지운다. [휴가]에 그 일정으로 등록된 휴가가 있으면 휴가째 지우고(주간보고에서도 빠짐),
 * 캘린더에만 있던 휴가 일정이면 캘린더에서 지운다. 본인 휴가만 (팀장·관리자는 모두).
 */
function apiDeleteLeaveEvent(eventId) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'calendar');
  eventId = String(eventId || '');
  const sh = ctx.ss.getSheetByName(LEAVE.SHEET);
  const rows = leaveRows_(sh);
  const live = r => String(r[LCOL.DELETED - 1]).toUpperCase() !== 'Y';
  let idx = rows.findIndex(r => String(r[LCOL.EVENT - 1]) === eventId && live(r));
  if (idx < 0) idx = rows.findIndex(r => String(r[LCOL.EVENT - 1]) === `${eventId}#${me.name}` && live(r)); // 여러 명 일정 중 내 것
  if (idx >= 0) return apiDeleteLeave(rows[idx][0]);

  const ev = getTeamCalendar_().getEventById(eventId);
  if (!ev) throw new Error('일정을 찾을 수 없습니다. 이미 지워졌을 수 있습니다.');
  const title = ev.getTitle();
  if (!isLeaveTitle_(title, ctx.members.map(m => m.name))) throw new Error('휴가 일정만 여기서 지울 수 있습니다.');
  const who = leaveNames_(title, ctx.members.map(m => m.name));
  if (!me.isLeader && !me.isAdmin && (who.length !== 1 || who[0] !== me.name)) throw new Error(who.length > 1 ? '여러 명이 함께 적힌 휴가 일정은 팀장·관리자만 지울 수 있습니다.' : '본인 휴가만 지울 수 있습니다.');
  try { ev.deleteEvent(); } catch (e) { throw new Error('캘린더에서 지울 권한이 없습니다. 팀 캘린더 "일정 변경" 권한을 확인하세요.'); }
  bumpCalendar_();
  return {};
}

/** 주간보고 작성 대상이면 그 사람 조각만 돌려준다 (팀장 등은 주간보고에 없으므로 생략) */
/** 웹앱 휴가와 같이 만든 채용 캘린더 일정 지우기 */
function dropRecruitLeaveEvent_(ctx, id) {
  if (!id || !recruitCalendarId_(ctx.cfg)) return;
  try { const ev = getRecruitCalendar_(ctx).getEventById(String(id)); if (ev) { ev.deleteEvent(); bumpCalendar_(); } } catch (e) { /* 권한 없음 */ }
}

/** 웹앱 휴가와 같이 만든 채용 캘린더 일정 날짜 맞추기 */
function moveRecruitLeaveEvent_(ctx, id, start, end) {
  if (!id || !recruitCalendarId_(ctx.cfg)) return;
  try { const ev = getRecruitCalendar_(ctx).getEventById(String(id)); if (ev) ev.setAllDayDates(start, addDays(end, 1)); } catch (e) { /* 권한 없음 */ }
}

function leavePatch_(ctx, name) {
  return ctx.members.some(m => m.name === name && m.write) ? memberPatch_(ctx, name) : {};
}

function nextLeaveId_(sh) {
  return leaveRows_(sh).reduce((m, r) => Math.max(m, Number(r[0]) || 0), 0) + 1;
}

/* ---------- 캘린더 → [휴가] 맞추기 ---------- */

/**
 * 캘린더 일정 중 제목에 팀원 이름과 휴가 단어가 있는 것을 [휴가]에 넣고,
 * 캘린더에서 지워진 일정은 [휴가]에서도 지운다 (캘린더에서 넣은 것, 웹앱에서 넣은 것 모두).
 * events: CalendarEvent[] (from~to 사이를 모두 읽은 것). 바뀐 게 있으면 true.
 */
/**
 * opt.source: '캘린더'(팀 캘린더, 기본) | '채용캘린더'(채용 캘린더 — 채용담당 본인 휴가만)
 * 각 캘린더는 자기가 넣은 [휴가] 줄만 고치고 지운다.
 */
function syncLeaveEvents_(ctx, events, from, to, cal, opt) {
  const source = (opt && opt.source) || '캘린더';
  const fromRecruit = source === RECRUIT_LEAVE_SOURCE;
  const lo = ctx.thisWeek.monday > from ? ctx.thisWeek.monday : from;
  const afterEnd = addDays(ctx.afterWeek.monday, 7);
  const hi = afterEnd < to ? afterEnd : to;
  if (lo >= hi) return false;

  const names = ((opt && opt.names) || ctx.members.map(m => m.name)).filter(Boolean).sort((a, b) => b.length - a.length);
  if (!names.length) return false;
  const found = {};
  const present = {}; // 캘린더에 아직 있는 일정 (제목을 바꿨어도 지운 것으로 보지 않도록 전부)
  // 일정 하나 → { 키: 휴가 } (여러 명이면 사람마다 한 줄씩, [휴가] 캘린더 일정ID 칸: "일정ID#이름")
  const leavesOfEvent = ev => {
    const out = {};
    const title = ev.getTitle();
    const word = title.match(LEAVE.WORDS);
    if (!word) return out;
    const who = leaveNames_(title, names);
    const allDay = ev.isAllDayEvent();
    const start = allDay ? ev.getAllDayStartDate() : startOfDay_(ev.getStartTime());
    const end = allDay ? addDays(ev.getAllDayEndDate(), -1) : startOfDay_(ev.getEndTime());
    who.forEach(name => {
      out[who.length > 1 ? `${ev.getId()}#${name}` : ev.getId()] = { name: name, type: leaveType_(word[1], title), start: start, end: end < start ? start : end };
    });
    return out;
  };
  events.forEach(ev => {
    present[ev.getId()] = true;
    const one = leavesOfEvent(ev);
    Object.keys(one).forEach(k => { if (one[k].start < hi && one[k].end >= lo) found[k] = one[k]; });
  });
  // 읽은 범위 밖으로 옮겨진 일정은 하나씩 찾아본다 (지운 것과 구분)
  const lookup = {};
  const moved = key => {
    const id = key.split('#')[0];
    if (!(id in lookup)) {
      let one = null;
      try {
        const ev = cal ? cal.getEventById(id) : null;
        one = ev ? leavesOfEvent(ev) : null;
      } catch (e) { one = null; }
      lookup[id] = one;
    }
    const f = lookup[id] ? lookup[id][key] || null : null;
    // 읽은 범위 안에 있다면 getEvents에 나왔어야 한다 → 안 나왔으면 지운 일정으로 본다
    return f && (f.start >= hi || f.end < lo) ? f : null;
  };
  const day = d => (d instanceof Date ? ymd(d) : '');

  const sh = ctx.ss.getSheetByName(LEAVE.SHEET);
  const rows = leaveRows_(sh);
  const known = {};
  rows.forEach(r => {
    if (r[LCOL.EVENT - 1]) known[String(r[LCOL.EVENT - 1])] = true;
    if (r[LCOL.REC_EVENT - 1]) known[String(r[LCOL.REC_EVENT - 1])] = true; // 웹앱 휴가와 같이 만든 채용 캘린더 일정
  });
  const mine = r => (String(r[LCOL.SOURCE - 1]) === RECRUIT_LEAVE_SOURCE) === fromRecruit; // 이 캘린더가 맡은 줄
  // 다른 캘린더(또는 웹앱)에 같은 사람·같은 날짜·같은 종류 휴가가 이미 있으면 또 넣지 않는다
  const same = {};
  rows.forEach(r => {
    if (String(r[LCOL.DELETED - 1]).toUpperCase() === 'Y' || !(r[LCOL.START - 1] instanceof Date)) return;
    const e = r[LCOL.END - 1] instanceof Date ? r[LCOL.END - 1] : r[LCOL.START - 1];
    same[[String(r[LCOL.NAME - 1]).trim(), ymd(r[LCOL.START - 1]), ymd(e), String(r[LCOL.TYPE - 1]).trim()].join('|')] = true;
  });

  const add = Object.keys(found).filter(id => !known[id] && !same[[found[id].name, ymd(found[id].start), ymd(found[id].end), found[id].type].join('|')]);
  const gone = [];
  const update = []; // 캘린더에서 날짜·종류를 바꾼 휴가 [행, 휴가]
  rows.forEach((r, i) => {
    const eid = String(r[LCOL.EVENT - 1] || '');
    if (!eid || String(r[LCOL.DELETED - 1]).toUpperCase() === 'Y' || !mine(r)) return; // 캘린더 일정이 없는 휴가는 그대로
    const s = r[LCOL.START - 1];
    const inRange = s instanceof Date && s >= lo && s < hi;
    let f = found[eid];
    if (!f && inRange) {
      const fromCal = String(r[LCOL.SOURCE - 1]) === source;
      // 캘린더에서 넣은 것: 휴가 일정이 아니게 됐거나 지워지면 / 웹앱에서 넣은 것: 캘린더에서 지워지면
      // (단, 날짜를 읽은 범위 밖으로 옮긴 것이면 지우지 않고 날짜를 따라간다)
      const gonePlain = fromCal ? true : !present[eid.split('#')[0]];
      if (gonePlain) {
        f = moved(eid);
        if (!f) { gone.push(i + 2); return; }
      }
    }
    if (!f) return;
    const e0 = r[LCOL.END - 1] instanceof Date ? r[LCOL.END - 1] : s;
    if (day(s) !== ymd(f.start) || day(e0) !== ymd(f.end) || String(r[LCOL.TYPE - 1]).trim() !== f.type) update.push([i + 2, f]);
  });
  if (!add.length && !gone.length && !update.length) return false;

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const target = sh || setupLeaveSheet_(ctx.ss);
    gone.forEach(row => {
      target.getRange(row, LCOL.DELETED).setValue('Y');
      if (!fromRecruit) dropRecruitLeaveEvent_(ctx, rows[row - 2][LCOL.REC_EVENT - 1]);
    });
    update.forEach(([row, f]) => {
      target.getRange(row, LCOL.TYPE).setValue(f.type);
      target.getRange(row, LCOL.START, 1, 2).setValues([[f.start, f.end]]);
      if (!fromRecruit) moveRecruitLeaveEvent_(ctx, rows[row - 2][LCOL.REC_EVENT - 1], f.start, f.end);
    });
    if (add.length) {
      let id = nextLeaveId_(target);
      const now = new Date();
      const values = add.map(eid => {
        const f = found[eid];
        return [id++, now, f.name, f.type, f.start, f.end, '', eid, source, ''];
      });
      target.getRange(target.getLastRow() + 1, 1, values.length, LCOL.DELETED).setValues(values);
    }
  } finally {
    lock.releaseLock();
  }
  ctx._leaves = null;
  try { CacheService.getDocumentCache().remove(CK.members(ctx)); } catch (e) { /* 캐시 없음 */ }
  return true;
}

/**
 * 팀 캘린더를 안 정했을 때: 웹앱에서 등록한 내 휴가는 내 기본 캘린더에 만들어진다.
 * 그 일정을 내가 캘린더에서 지웠으면 [휴가]에서도 지운다. (남의 기본 캘린더는 볼 수 없으므로 본인 것만)
 */
function syncMyLeaves_(ctx, me) {
  const sh = ctx.ss.getSheetByName(LEAVE.SHEET);
  if (!sh || !me.name) return false;
  const lo = ctx.thisWeek.monday, hi = addDays(ctx.afterWeek.monday, 7);
  const present = {};
  CalendarApp.getDefaultCalendar().getEvents(lo, hi).forEach(e => { present[e.getId()] = true; });
  const gone = [];
  leaveRows_(sh).forEach((r, i) => {
    const eid = String(r[LCOL.EVENT - 1] || '');
    const s = r[LCOL.START - 1];
    if (!eid || String(r[LCOL.NAME - 1]).trim() !== me.name || (String(r[LCOL.SOURCE - 1]) === '캘린더' || String(r[LCOL.SOURCE - 1]) === RECRUIT_LEAVE_SOURCE)
      || String(r[LCOL.DELETED - 1]).toUpperCase() === 'Y' || !(s instanceof Date) || s < lo || s >= hi) return;
    if (!present[eid]) gone.push(i + 2);
  });
  if (!gone.length) return false;
  gone.forEach(row => sh.getRange(row, LCOL.DELETED).setValue('Y'));
  ctx._leaves = null;
  try { CacheService.getDocumentCache().remove(CK.members(ctx)); } catch (e) { /* 캐시 없음 */ }
  return true;
}

/** 매시간 자동 실행: 이번 주~다다음 주 캘린더를 읽어 맞춘다 */
function syncLeavesFromCalendar_(ctx) {
  let changed = false;
  if (teamCalendarSet_(ctx)) {
    const from = ctx.thisWeek.monday;
    const to = addDays(ctx.afterWeek.monday, 7);
    try {
      const cal = getTeamCalendar_();
      changed = syncLeaveEvents_(ctx, cal.getEvents(from, to), from, to, cal);
    } catch (e) { /* 캘린더 미설정·권한 없음 */ }
  }
  return syncRecruitLeaves_(ctx) || changed;
}

/**
 * 채용 캘린더에 적힌 연차·반차·휴가 → [휴가] (채용담당이 넣은 일정 중 채용담당 본인 휴가만, 출처 "채용캘린더")
 * 팀 캘린더에 같은 휴가가 이미 있으면 한 번만 들어간다.
 */
const RECRUIT_LEAVE_SOURCE = '채용캘린더';
function syncRecruitLeaves_(ctx) {
  if (!recruitCalendarId_(ctx.cfg)) return false;
  const names = ctx.members.filter(m => m.recruit && m.name).map(m => m.name);
  if (!names.length) return false;
  const from = ctx.thisWeek.monday;
  const to = addDays(ctx.afterWeek.monday, 7);
  try {
    const cal = getRecruitCalendar_(ctx);
    // 채용담당이 넣은 일정만
    const events = cal.getEvents(from, to).filter(e => recruitCreator_(e, ctx));
    return syncLeaveEvents_(ctx, events, from, to, cal, { source: RECRUIT_LEAVE_SOURCE, names: names });
  } catch (e) {
    return false; // 채용 캘린더 권한 없음
  }
}

/**
 * 팀 캘린더 ID가 있을 때만 맞춘다. 비어 있으면 사람마다 자기 기본 캘린더를 보게 되어
 * 다른 사람의 휴가 일정이 "지워진 것"처럼 보일 수 있기 때문이다.
 */
function teamCalendarSet_(ctx) {
  return !!String(ctx.cfg[CFG.CALENDAR] || '').trim();
}

/**
 * 휴가 일정 제목을 팀 양식 "연차] 지선" 으로 맞춘다 (화면 표시용, 캘린더 원본은 그대로).
 * 이름은 성을 뺀 이름 (같은 이름이 둘이면 성까지). 휴가 일정이 아니면 '' 를 돌려준다.
 */
function leaveTitle_(title, names) {
  const word = String(title).match(LEAVE.WORDS);
  const who = leaveNames_(title, names);
  if (!word || !who.length) return '';
  const type = leaveType_(word[1], title);
  // "[정기휴가] 윤수아", "출장] 수아"처럼 앞에 종류를 적어 두었으면 그 종류를 쓴다
  const tagged = String(title).match(/^\s*\[?\s*([가-힣]+)\s*\]/);
  const tag = tagged && LEAVE.TYPES.concat(['반차', '반반차', '휴가', '휴무']).indexOf(tagged[1]) >= 0 ? tagged[1] : type;
  return leaveLabel_(tag, who.map(n => shortName_(n, names)).join(', '));
}

/** 팀 양식: "연차] 지선" */
function leaveLabel_(type, who) {
  return `${type}] ${who}`;
}

/** 성을 뺀 이름 ("안지선" → "지선"). 같은 이름을 쓰는 팀원이 또 있거나 형식이 다르면 그대로 */
function shortName_(name, names) {
  const ok = n => /^[가-힣]{3,4}$/.test(n);
  if (!ok(name)) return name;
  const g = name.slice(1);
  return (names || []).filter(n => ok(n) && n.slice(1) === g).length <= 1 ? g : name;
}

/** 캘린더 화면에서 휴가 일정을 다른 색으로 보여주기 위한 판별 */
function isLeaveTitle_(title, names) {
  return LEAVE.WORDS.test(title) && !!leaveName_(title, names);
}

/**
 * 일정 제목에서 팀원 찾기 → [팀원] 시트의 이름(성 포함)
 *  1) "김정호 연차"처럼 이름 전체가 있으면 그 사람
 *  2) "연차] 정호"처럼 성을 뺀 이름만 있어도, 그 이름을 쓰는 팀원이 한 명뿐이면 그 사람
 *     (앞에 다른 한글이 붙은 "박정호"는 다른 사람으로 보고 건너뛴다)
 */
function leaveName_(title, names) {
  return leaveNames_(title, names)[0] || '';
}

/** "오전반차] 지애, 소영, 은지"처럼 여러 명이면 모두 (제목에 나온 순서대로) */
function leaveNames_(title, names) {
  const t = String(title || '');
  const list = names.filter(Boolean).slice().sort((a, b) => b.length - a.length);
  const hits = []; // [위치, 이름]
  const taken = [];
  const free = (at, len) => !taken.some(([a, b]) => at < b && at + len > a);
  list.forEach(n => {
    let at = t.indexOf(n);
    while (at >= 0) {
      if (free(at, n.length)) { hits.push([at, n]); taken.push([at, at + n.length]); }
      at = t.indexOf(n, at + 1);
    }
  });
  const given = {};
  list.forEach(n => {
    if (!/^[가-힣]{3,4}$/.test(n)) return;
    const g = n.slice(1);
    given[g] = given[g] ? '!' : n; // 같은 이름이 둘이면 누군지 모르므로 쓰지 않는다
  });
  Object.keys(given).filter(g => given[g] !== '!').forEach(g => {
    const re = new RegExp('(^|[^가-힣])' + g, 'g');
    let m;
    while ((m = re.exec(t))) {
      const at = m.index + m[1].length;
      if (free(at, g.length)) { hits.push([at, given[g]]); taken.push([at, at + g.length]); }
    }
  });
  return hits.sort((a, b) => a[0] - b[0]).map(h => h[1]).filter((n, i, a) => a.indexOf(n) === i);
}

function leaveType_(word, title) {
  // "오후 반차", "오전 반반차"처럼 띄어 쓴 것도 오전/오후로
  if (word === '반반차') return /오전/.test(title) ? '오전반반차' : /오후/.test(title) ? '오후반반차' : '반반차';
  if (word === '반차') return /오전/.test(title) ? '오전반차' : /오후/.test(title) ? '오후반차' : '반차';
  if (word === '하계휴가' || word === '여름휴가') return '정기휴가';
  return word;
}

function startOfDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

// ===================================================================
// Files.gs
// ===================================================================

/**
 * 파일 첨부 (구글 드라이브)
 *
 * [설정] > "첨부 폴더 ID" 의 드라이브 폴더 아래에 메뉴별 하위 폴더(게시판/회사기준/예산전용)를 만들어 저장한다.
 * 폴더 ID: 드라이브에서 폴더를 열었을 때 주소 .../folders/<여기> 부분
 * 팀원이 올리려면 그 폴더가 팀원에게 "편집자"로 공유되어 있어야 한다.
 *
 * 글/기준/신청에는 첨부 목록을 JSON 문자열로 한 칸에 저장한다: [{id, name, url, size, mime}]
 */

const FILES = {
  MAX_BYTES: 10 * 1024 * 1024,
  AREAS: { board: '게시판', rules: '회사기준', budget: '예산전용' },
};

/** file: { name, mimeType, data(base64), area } → 첨부 정보 */
function apiUploadFile(file) {
  const me = currentMember_(getMembers());
  if (!me.name && !me.isAdmin) throw new Error('[팀원] 시트에 등록된 사람만 파일을 올릴 수 있습니다.');
  requireMenu_(me, { board: 'board', rules: 'rules', budget: 'budget' }[file.area] || 'board');

  const name = String(file.name || '').trim() || '첨부파일';
  const bytes = Utilities.base64Decode(String(file.data || ''));
  if (!bytes.length) throw new Error('빈 파일은 올릴 수 없습니다.');
  if (bytes.length > FILES.MAX_BYTES) throw new Error('10MB 이하 파일만 올릴 수 있습니다. 큰 파일은 드라이브에 직접 올리고 링크를 붙여주세요.');

  const folder = areaFolder_(FILES.AREAS[file.area] || '기타');
  const created = folder.createFile(Utilities.newBlob(bytes, file.mimeType || 'application/octet-stream', name));
  if (me.name) created.setDescription(`올린 사람: ${me.name}`);
  return {
    id: created.getId(),
    name: created.getName(),
    url: created.getUrl(),
    size: bytes.length,
    mime: created.getMimeType(),
  };
}

function areaFolder_(areaName) {
  const id = String(getConfig()[CFG.DRIVE_FOLDER] || '').trim();
  if (!id) throw new Error('첨부 폴더가 설정되지 않았습니다. 관리자에게 [설정] > "첨부 폴더 ID" 입력을 요청하세요.');
  let root;
  try {
    root = DriveApp.getFolderById(id);
  } catch (e) {
    throw new Error('첨부 폴더에 접근할 수 없습니다. 폴더가 "편집자"로 공유되어 있는지 관리자에게 확인하세요.');
  }
  const it = root.getFoldersByName(areaName);
  return it.hasNext() ? it.next() : root.createFolder(areaName);
}

/** 화면에서 넘어온 첨부 목록을 검증해 시트에 저장할 JSON 문자열로 만든다 */
function filesToCell_(files) {
  const list = (files || [])
    .filter(f => f && f.id && f.name)
    .slice(0, 20)
    .map(f => ({ id: String(f.id), name: String(f.name), url: String(f.url || ''), size: Number(f.size) || 0, mime: String(f.mime || '') }));
  return list.length ? JSON.stringify(list) : '';
}

function filesFromCell_(v) {
  if (!v) return [];
  try {
    const list = JSON.parse(String(v));
    return Array.isArray(list) ? list : [];
  } catch (e) {
    return [];
  }
}

/** 예전에 만든 시트에 첨부 열 제목이 없으면 채운다 */
function ensureHeader_(sh, col, title) {
  const cell = sh.getRange(1, col);
  if (!cell.getValue()) {
    cell.setValue(title);
    styleHeader_(cell);
    sh.setColumnWidth(col, 220);
  }
}

// ===================================================================
// Admin.gs
// ===================================================================

/**
 * 관리 메뉴 (관리자 전용): 팀원 관리, 권한 한 번에 적용, 메뉴 권한, 설정
 *
 * 관리자 = 스프레드시트 소유자 + [팀원] 시트에서 관리자(Y)인 사람
 *
 * 권한 적용이 하는 일
 *  1. 스프레드시트를 팀원에게 편집자로 공유
 *  2. 첨부 폴더를 편집자로 공유
 *  3. 팀 캘린더를 "일정 변경" 권한으로 공유 (고급 서비스 Calendar 필요)
 *  4. 입력시트 생성 + 시트 보호 (본인/팀장/관리자 구분)
 */

const MENUS = [
  { key: 'dash', label: '대시보드' },
  { key: 'weekly', label: '주간업무보고' },
  { key: 'budget', label: '예산전용' },
  { key: 'calendar', label: '일정' },
  { key: 'board', label: '공유게시판' },
  { key: 'rules', label: '회사 기준' },
];
const ROLES = ['팀원', '팀장'];
const MENU_SHEET = '메뉴권한';

/* ---------- 메뉴 권한 ---------- */

function setupMenuSheet_(ss) {
  if (ss.getSheetByName(MENU_SHEET)) return;
  const sh = ss.insertSheet(MENU_SHEET);
  const rows = [['메뉴', '키'].concat(ROLES)].concat(MENUS.map(m => [m.label, m.key].concat(ROLES.map(() => 'Y'))));
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  styleHeader_(sh.getRange(1, 1, 1, rows[0].length));
  sh.getRange(2, 3, MENUS.length, ROLES.length).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Y', 'N']).build());
  sh.hideColumns(2);
  sh.setFrozenRows(1);
  sh.getRange('A1').setNote('역할별로 볼 수 있는 메뉴 (Y/N). 관리자는 항상 모든 메뉴를 봅니다.');
}

/** { 팀원: {dash:true, ...}, 팀장: {...} } */
function getMenuAccess_() {
  return cached_('menus', () => {
    const access = {};
    ROLES.forEach(r => { access[r] = {}; MENUS.forEach(m => { access[r][m.key] = true; }); });
    const sh = SpreadsheetApp.getActive().getSheetByName(MENU_SHEET);
    if (!sh || sh.getLastRow() < 2) return access;
    const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    const header = v[0].map(String);
    v.slice(1).forEach(row => {
      const key = String(row[1]);
      ROLES.forEach(r => {
        const c = header.indexOf(r);
        if (c >= 0 && key) access[r][key] = String(row[c]).toUpperCase() !== 'N';
      });
    });
    return access;
  });
}

/** 접속자가 볼 수 있는 메뉴 키 목록 */
function allowedMenus_(me) {
  if (me.isAdmin) return MENUS.map(m => m.key).concat('admin');
  const access = me.role ? getMenuAccess_()[me.role] || {} : null;
  const keys = access ? MENUS.map(m => m.key).filter(k => access[k] !== false) : [];
  // 예산 담당자는 역할 메뉴 권한과 상관없이 예산전용 메뉴를 쓴다 (관리 메뉴·다른 권한은 늘지 않음)
  if (me.budgetManager && keys.indexOf('budget') < 0) keys.push('budget');
  return keys;
}

function requireMenu_(me, key) {
  if (allowedMenus_(me).indexOf(key) < 0) {
    const label = (MENUS.find(m => m.key === key) || { label: key }).label;
    throw new Error(`[${label}] 메뉴를 사용할 권한이 없습니다. 관리자에게 문의하세요.`);
  }
}

function requireAdmin_(ctx) {
  const me = currentMember_(ctx.members);
  if (!me.isAdmin) throw new Error('관리자만 사용할 수 있습니다.');
  return me;
}

/* ---------- 조회 ---------- */

/** force=true면 공유 상태를 새로 확인 (상태 새로고침 버튼) */
function apiAdminData(force) {
  const ctx = getContext();
  requireAdmin_(ctx);
  return adminData_(ctx, !force);
}

/** 공유 상태만 (화면이 먼저 뜬 뒤 따로 부른다). force면 캐시를 무시하고 다시 확인한다. */
function apiAdminStatus(force) {
  const ctx = getContext();
  requireAdmin_(ctx);
  return force ? freshStatus_(ctx) : cached_('access', () => accessStatus_(ctx), ACCESS_TTL);
}

const ACCESS_TTL = 1800; // 공유 상태 캐시 30분 (권한 적용·팀원 저장 때는 바로 다시 확인)

/**
 * useCache=true면 공유 상태(시트·폴더·캘린더 확인, 가장 느린 부분)는 캐시에 있을 때만 싣고,
 * 없으면 null → 화면이 먼저 뜨고 apiAdminStatus 로 따로 받는다.
 */
function adminData_(ctx, useCache) {
  const cfg = ctx.cfg;
  return {
    members: ctx.members.map(m => ({ order: m.order, name: m.name, email: m.email, role: m.role, write: m.write, admin: m.admin, budget: m.budget, notify: m.notify, recruit: m.recruit })),
    menus: MENUS,
    roles: ROLES,
    menuAccess: getMenuAccess_(),
    settings: {
      team: String(cfg[CFG.TEAM] || ''),
      deadlineHour: Number(cfg[CFG.DEADLINE_HOUR]) || 14,
      remindHours: Number(cfg[CFG.REMIND_HOURS]) || 3,
      calendarId: String(cfg[CFG.CALENDAR] || ''),
      extraCalendars: extraCalendarIds_(cfg).join('\n'),
      recruitCalendarId: recruitCalendarId_(cfg),
      folderId: String(cfg[CFG.DRIVE_FOLDER] || ''),
    },
    webAppUrl: ScriptApp.getService().getUrl() || '',
    status: useCache ? cachedOnly_('access') : freshStatus_(ctx),
  };
}

function freshStatus_(ctx) {
  const status = accessStatus_(ctx);
  putCache_('access', status, ACCESS_TTL);
  return status;
}

function cachedOnly_(key) {
  try { const hit = cacheGet_(key); return hit ? JSON.parse(hit) : null; } catch (e) { return null; }
}

/**
 * 팀원별 접근 상태
 * sheet/folder: 'editor' | 'viewer' | 'none' | 'owner'   calendar: 'writer' | 'reader' | 'none' | 'owner'
 * 설정이 없거나 확인할 수 없으면 resources[].ok=false
 */
function accessStatus_(ctx) {
  const lower = list => list.map(u => u.getEmail().toLowerCase());
  const res = { sheet: { ok: true, name: ctx.ss.getName() }, folder: { ok: false }, calendar: { ok: false } };

  const file = DriveApp.getFileById(ctx.ss.getId());
  const fileOwner = file.getOwner() ? file.getOwner().getEmail().toLowerCase() : '';
  const fileEditors = lower(file.getEditors());
  const fileViewers = lower(file.getViewers());

  let folderEditors = [], folderViewers = [], folderOwner = '';
  const folderId = String(ctx.cfg[CFG.DRIVE_FOLDER] || '').trim();
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      folderOwner = folder.getOwner() ? folder.getOwner().getEmail().toLowerCase() : '';
      folderEditors = lower(folder.getEditors());
      folderViewers = lower(folder.getViewers());
      res.folder = { ok: true, name: folder.getName() };
    } catch (e) {
      res.folder = { ok: false, error: '폴더를 찾을 수 없거나 접근 권한이 없습니다.' };
    }
  } else {
    res.folder = { ok: false, error: '첨부 폴더 ID가 비어 있습니다.' };
  }

  let acl = {};
  const calId = String(ctx.cfg[CFG.CALENDAR] || '').trim();
  if (calId) {
    try {
      const cal = CalendarApp.getCalendarById(calId);
      if (!cal) throw new Error('not found');
      res.calendar = { ok: true, name: cal.getName() };
      (Calendar.Acl.list(calId).items || []).forEach(rule => {
        if (rule.scope && rule.scope.type === 'user') acl[String(rule.scope.value).toLowerCase()] = rule.role;
      });
    } catch (e) {
      res.calendar = {
        ok: false,
        error: /Calendar is not defined/.test(String(e))
          ? 'Apps Script에서 고급 서비스 "Google Calendar API"를 켜야 합니다.'
          : '캘린더를 찾을 수 없거나 공유 설정을 볼 권한이 없습니다.',
      };
    }
  } else {
    res.calendar = { ok: false, error: '팀 캘린더 ID가 비어 있습니다.' };
  }

  const level = (email, owner, editors, viewers) =>
    email === owner ? 'owner' : editors.indexOf(email) >= 0 ? 'editor' : viewers.indexOf(email) >= 0 ? 'viewer' : 'none';

  // 시트 보호는 한 번에 읽는다 (팀원마다 따로 읽으면 느림)
  const protectedFor = {};
  ctx.ss.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => {
    const name = p.getRange().getSheet().getName();
    if (name.indexOf(SHEET.INPUT_PREFIX) === 0) protectedFor[name] = lower(p.getEditors());
  });

  const people = {};
  ctx.members.filter(m => m.email).forEach(m => {
    const e = m.email.toLowerCase();
    const title = SHEET.INPUT_PREFIX + m.name;
    const sh = m.write ? ctx.ss.getSheetByName(title) : null;
    const inputProtected = !!sh && (protectedFor[title] || []).indexOf(e) >= 0;
    people[m.name] = {
      sheet: level(e, fileOwner, fileEditors, fileViewers),
      input: !m.write ? 'n/a' : !sh ? 'missing' : inputProtected ? 'ok' : 'unprotected',
      folder: res.folder.ok ? level(e, folderOwner, folderEditors, folderViewers) : 'n/a',
      calendar: res.calendar.ok ? (acl[e] || 'none') : 'n/a',
    };
  });
  // 함께 볼 캘린더: 관리자 계정으로 열리는지 (팀원도 각자 볼 권한이 있어야 보인다)
  res.extra = extraCalendarIds_(ctx.cfg).map(id => {
    try { const c = CalendarApp.getCalendarById(id); return c ? { id: id, ok: true, name: c.getName() } : { id: id, ok: false, error: '캘린더를 찾을 수 없거나 볼 권한이 없습니다.' }; }
    catch (e) { return { id: id, ok: false, error: '캘린더를 찾을 수 없거나 볼 권한이 없습니다.' }; }
  });
  const recId = recruitCalendarId_(ctx.cfg);
  if (recId) {
    try { const c = CalendarApp.getCalendarById(recId); res.recruit = c ? { ok: true, name: c.getName() } : { ok: false, error: '캘린더를 찾을 수 없거나 볼 권한이 없습니다.' }; }
    catch (e) { res.recruit = { ok: false, error: '캘린더를 찾을 수 없거나 볼 권한이 없습니다.' }; }
  }
  return { resources: res, people: people };
}

/* ---------- 저장 ---------- */

/** 팀원 목록 전체 저장. list: [{order, name, email, role, write, admin, budget, notify}] */
function apiSaveMembers(list) {
  const ctx = getContext();
  requireAdmin_(ctx);
  ensureHeader_(ctx.ss.getSheetByName(SHEET.MEMBERS), 6, '관리자(Y/N)');
  ensureHeader_(ctx.ss.getSheetByName(SHEET.MEMBERS), 7, '예산담당(Y/N)');
  ensureHeader_(ctx.ss.getSheetByName(SHEET.MEMBERS), 8, '알림메일');
  ensureHeader_(ctx.ss.getSheetByName(SHEET.MEMBERS), 9, '채용담당(Y/N)');

  const rows = (list || []).map((m, i) => ({
    order: Number(m.order) || i + 1,
    name: String(m.name || '').trim(),
    email: String(m.email || '').trim().toLowerCase(),
    role: ROLES.indexOf(m.role) >= 0 ? m.role : '팀원',
    write: !!m.write,
    admin: !!m.admin,
    budget: !!m.budget,
    notify: String(m.notify || '').trim().toLowerCase(),
    recruit: !!m.recruit,
  })).filter(m => m.name);

  const names = {};
  rows.forEach(m => {
    if (names[m.name]) throw new Error(`이름이 겹칩니다: ${m.name}. 동명이인은 "김민지A"처럼 구분해 주세요.`);
    names[m.name] = true;
    if (m.email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(m.email)) throw new Error(`이메일 형식을 확인해 주세요: ${m.name} (${m.email})`);
    if (m.notify && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(m.notify)) throw new Error(`알림메일 형식을 확인해 주세요: ${m.name} (${m.notify})`);
  });

  const sh = ctx.ss.getSheetByName(SHEET.MEMBERS);
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 9).clearContent();
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 9).setValues(rows.map(m => [m.order, m.name, m.email, m.role, m.write ? 'Y' : 'N', m.admin ? 'Y' : 'N', m.budget ? 'Y' : 'N', m.notify, m.recruit ? 'Y' : 'N']));
  }
  clearCaches_(ctx);
  // 저장만 하고 공유 상태는 다시 확인하지 않는다 (확인은 "권한 한 번에 적용"이나 "상태 새로고침"에서)
  return adminData_(getContext(), true);
}

function apiSaveMenuAccess(access) {
  const ctx = getContext();
  requireAdmin_(ctx);
  setupMenuSheet_(ctx.ss);
  const sh = ctx.ss.getSheetByName(MENU_SHEET);
  const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const header = v[0].map(String);
  const out = v.slice(1).map(row => ROLES.map(r => {
    const allowed = access && access[r] ? access[r][String(row[1])] !== false : true;
    return allowed ? 'Y' : 'N';
  }));
  const firstRoleCol = header.indexOf(ROLES[0]) + 1;
  sh.getRange(2, firstRoleCol, out.length, ROLES.length).setValues(out);
  try { CacheService.getDocumentCache().remove('menus'); } catch (e) { /* 캐시 없음 */ }
  return { menuAccess: getMenuAccess_() };
}

function apiSaveSettings(s) {
  const ctx = getContext();
  requireAdmin_(ctx);
  const hour = Math.round(Number(s.deadlineHour));
  const remind = Math.round(Number(s.remindHours));
  if (!(hour >= 6 && hour <= 22)) throw new Error('마감 시각은 6~22시 사이로 입력해 주세요.');
  if (!(remind >= 0 && remind <= 48)) throw new Error('미작성 알림은 0~48시간 전으로 입력해 주세요.');
  setConfigValue(CFG.TEAM, String(s.team || '').trim());
  setConfigValue(CFG.DEADLINE_HOUR, hour);
  setConfigValue(CFG.REMIND_HOURS, remind);
  setConfigValue(CFG.CALENDAR, String(s.calendarId || '').trim());
  const extra = String(s.extraCalendars || '').split(/[\s,;]+/).map(x => x.trim()).filter(Boolean);
  if (extra.length > EXTRA_CAL_MAX) throw new Error(`함께 볼 캘린더는 ${EXTRA_CAL_MAX}개까지 넣을 수 있습니다.`);
  setConfigValue(CFG.EXTRA_CAL, extra.filter((v, i, a) => a.indexOf(v) === i).join(', '));
  setConfigValue(CFG.RECRUIT_CAL, String(s.recruitCalendarId || '').trim());
  bumpCalendar_(); // 달력을 새로 읽게
  setConfigValue(CFG.DRIVE_FOLDER, String(s.folderId || '').trim().replace(/^.*\/folders\//, '').replace(/[?#].*$/, ''));
  clearCaches_(ctx);
  return adminData_(getContext(), false);
}

function clearCaches_(ctx) {
  try {
    CacheService.getDocumentCache().removeAll([CK.members(ctx), CK.hr(ctx), CK.budget, CK.board, CK.rules, CK.minutes, 'menus', CTX_KEY]);
  } catch (e) { /* 캐시 없음 */ }
}

/* ---------- 권한 한 번에 적용 ---------- */

/** 결과: { log: [{ok, text}], data } */
function apiApplyPermissions() {
  const ctx = getContext();
  requireAdmin_(ctx);
  const log = [];
  const ok = text => log.push({ ok: true, text: text });
  const warn = text => log.push({ ok: false, text: text });
  const people = ctx.members.filter(m => m.email);
  ctx.members.filter(m => !m.email).forEach(m => warn(`${m.name}: 이메일이 없어 건너뛰었습니다.`));

  // 1. 스프레드시트
  const file = DriveApp.getFileById(ctx.ss.getId());
  const fileOwner = file.getOwner() ? file.getOwner().getEmail().toLowerCase() : '';
  const fileEditors = file.getEditors().map(u => u.getEmail().toLowerCase());
  people.forEach(m => {
    const e = m.email.toLowerCase();
    if (e === fileOwner || fileEditors.indexOf(e) >= 0) return;
    try { file.addEditor(e); ok(`${m.name}: 스프레드시트 편집 권한을 줬습니다.`); }
    catch (err) { warn(`${m.name}: 스프레드시트 공유 실패 (${err.message})`); }
  });

  // 2. 첨부 폴더
  const folderId = String(ctx.cfg[CFG.DRIVE_FOLDER] || '').trim();
  if (!folderId) {
    warn('첨부 폴더 ID가 비어 있어 폴더 공유를 건너뛰었습니다.');
  } else {
    try {
      const folder = DriveApp.getFolderById(folderId);
      const owner = folder.getOwner() ? folder.getOwner().getEmail().toLowerCase() : '';
      const editors = folder.getEditors().map(u => u.getEmail().toLowerCase());
      people.forEach(m => {
        const e = m.email.toLowerCase();
        if (e === owner || editors.indexOf(e) >= 0) return;
        try { folder.addEditor(e); ok(`${m.name}: 첨부 폴더 편집 권한을 줬습니다.`); }
        catch (err) { warn(`${m.name}: 첨부 폴더 공유 실패 (${err.message})`); }
      });
    } catch (err) {
      warn('첨부 폴더에 접근할 수 없어 폴더 공유를 건너뛰었습니다.');
    }
  }

  // 3. 팀 캘린더
  const calId = String(ctx.cfg[CFG.CALENDAR] || '').trim();
  if (!calId) {
    warn('팀 캘린더 ID가 비어 있어 캘린더 공유를 건너뛰었습니다.');
  } else {
    try {
      const acl = {};
      (Calendar.Acl.list(calId).items || []).forEach(r => {
        if (r.scope && r.scope.type === 'user') acl[String(r.scope.value).toLowerCase()] = r.role;
      });
      people.forEach(m => {
        const e = m.email.toLowerCase();
        if (acl[e] === 'owner' || acl[e] === 'writer') return;
        try {
          Calendar.Acl.insert({ role: 'writer', scope: { type: 'user', value: e } }, calId, { sendNotifications: false });
          ok(`${m.name}: 팀 캘린더 "일정 변경" 권한을 줬습니다.`);
        } catch (err) {
          warn(`${m.name}: 캘린더 공유 실패 (${err.message})`);
        }
      });
    } catch (err) {
      warn(/Calendar is not defined/.test(String(err))
        ? '캘린더 공유를 하려면 Apps Script에서 고급 서비스 "Google Calendar API"를 켜야 합니다.'
        : '팀 캘린더 공유 설정을 바꿀 수 없습니다. 캘린더 소유자가 실행해 주세요.');
    }
  }

  // 4. 입력시트 + 시트 보호
  try {
    const missing = applyProtections_(ctx);
    ok('입력시트와 시트 보호를 적용했습니다.');
    missing.forEach(n => warn(`${n}: 이메일이 없어 입력시트를 본인 전용으로 잠그지 못했습니다.`));
  } catch (err) {
    warn(`시트 보호 적용 실패 (${err.message})`);
  }

  clearCaches_(ctx);
  if (!log.some(l => l.ok && /권한을 줬습니다/.test(l.text))) ok('새로 줄 공유 권한은 없었습니다. 모두 이미 공유되어 있습니다.');
  return { log: log, data: adminData_(getContext(), false) };
}

/** 팀원을 뺄 때 시트·폴더·캘린더 접근 해제 (소유자와 본인은 건드리지 않는다) */
function apiRevokeAccess(email) {
  const ctx = getContext();
  const me = requireAdmin_(ctx);
  const e = String(email || '').trim().toLowerCase();
  if (!e) throw new Error('이메일이 없습니다.');
  if (e === me.email) throw new Error('본인의 접근은 해제할 수 없습니다.');
  const log = [];

  const file = DriveApp.getFileById(ctx.ss.getId());
  if (file.getOwner() && file.getOwner().getEmail().toLowerCase() === e) throw new Error('스프레드시트 소유자의 접근은 해제할 수 없습니다.');
  try { file.removeEditor(e); } catch (err) { /* 편집자가 아님 */ }
  try { file.removeViewer(e); } catch (err) { /* 뷰어가 아님 */ }
  log.push({ ok: true, text: '스프레드시트 접근을 해제했습니다.' });

  const folderId = String(ctx.cfg[CFG.DRIVE_FOLDER] || '').trim();
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      try { folder.removeEditor(e); } catch (err) { /* 편집자가 아님 */ }
      try { folder.removeViewer(e); } catch (err) { /* 뷰어가 아님 */ }
      log.push({ ok: true, text: '첨부 폴더 접근을 해제했습니다.' });
    } catch (err) {
      log.push({ ok: false, text: '첨부 폴더 접근 해제 실패' });
    }
  }

  const calId = String(ctx.cfg[CFG.CALENDAR] || '').trim();
  if (calId) {
    try {
      Calendar.Acl.remove(calId, 'user:' + e);
      log.push({ ok: true, text: '팀 캘린더 공유를 해제했습니다.' });
    } catch (err) {
      log.push({ ok: false, text: '팀 캘린더 공유 해제 실패 (이미 해제되었거나 권한 없음)' });
    }
  }
  return { log: log, data: adminData_(getContext(), false) };
}

// ===================================================================
// WebApp.gs
// ===================================================================

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
 * 화면을 열 때마다 시트 전체를 다시 읽지 않도록 30분간 캐시한다 (팀 전체가 공유).
 * 웹앱에서 저장하면 해당 부분 캐시를 바로 갱신하고, 시트를 직접 고친 내용은 최대 30분 뒤 반영된다.
 * (매시간 자동 실행이 캐시를 미리 채워 두어, 아침 첫 접속도 빠르다)
 * 캐시 한 칸은 100KB까지라, 팀원·게시글·예산 내역이 쌓여 커지면 여러 칸에 나눠 담는다.
 */

const CACHE_TTL = 1800; // 30분. 웹앱에서 저장하면 바로 고치므로, 시트를 직접 고친 내용만 늦게 보인다
const CACHE_CHUNK = 30000; // 글자 수 (한글 3바이트 기준 90KB)
const CK = {
  members: ctx => 'members:' + ctx.thisWeek.key,
  hr: ctx => 'hr:' + ctx.thisWeek.key,
  budget: 'budget',
  board: 'board',
  rules: 'rules',
  minutes: 'minutes',
  owner: 'owner',
};

function cached_(key, build, ttl) {
  try {
    const hit = cacheGet_(key);
    if (hit) return JSON.parse(hit);
  } catch (e) { /* 캐시 오류는 무시하고 새로 읽는다 */ }
  const value = build();
  putCache_(key, value, ttl);
  return value;
}

function putCache_(key, value, ttl) {
  try {
    const cache = CacheService.getDocumentCache();
    const text = JSON.stringify(value);
    if (text.length <= CACHE_CHUNK) return cache.put(key, text, ttl || CACHE_TTL);
    // 큰 값: key#0, key#1 … 에 나눠 담고 key 에는 조각 수만 적는다
    const parts = {};
    let n = 0;
    for (let i = 0; i < text.length; i += CACHE_CHUNK) parts[key + '#' + n++] = text.slice(i, i + CACHE_CHUNK);
    parts[key] = '#chunks:' + n;
    cache.putAll(parts, ttl || CACHE_TTL);
  } catch (e) { /* 캐시할 수 없으면 건너뛴다 */ }
}

/** 캐시 문자열 (나눠 담은 값은 다시 잇는다). 없거나 조각이 빠졌으면 null */
function cacheGet_(key) {
  const cache = CacheService.getDocumentCache();
  const head = cache.get(key);
  if (!head || head.indexOf('#chunks:') !== 0) return head;
  const keys = [];
  for (let i = 0, n = Number(head.slice(8)); i < n; i++) keys.push(key + '#' + i);
  const got = cache.getAll(keys);
  return keys.every(k => got[k] != null) ? keys.map(k => got[k]).join('') : null;
}

/** 매시간 자동 실행 때 화면 데이터를 미리 읽어 캐시에 넣는다 → 아침 첫 접속도 시트를 읽지 않고 바로 뜬다 */
function warmCaches_(ctx) {
  const ss = ctx.ss, ttl = 3600;
  [
    [CK.members(ctx), () => readMembers_(ctx)],
    [CK.hr(ctx), () => readHr_(ctx)],
    [CK.budget, () => readBudget_(ss)],
    [CK.board, () => readBoard_(ss)],
    [CK.rules, () => readRules_(ss)],
    [CK.minutes, () => readMinutes_(ss)],
  ].forEach(([key, read]) => {
    try { putCache_(key, read(), ttl); } catch (e) { /* 한 부분이 실패해도 나머지는 채운다 */ }
  });
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
    members: can('weekly') ? forViewer_(base.me, cached_(CK.members(ctx), () => readMembers_(ctx))) : [],
    hr: can('weekly') ? cached_(CK.hr(ctx), () => readHr_(ctx)) : { thisWeek: [], nextWeek: [] },
    budget: can('budget') ? cached_(CK.budget, () => readBudget_(ss)) : { accounts: [], requests: [] },
    board: can('board') ? cached_(CK.board, () => readBoard_(ss)) : [],
    rules: can('rules') ? cached_(CK.rules, () => readRules_(ss)) : [],
    minutes: can('weekly') ? cached_(CK.minutes, () => readMinutes_(ss)) : [],
    people: can('weekly') ? ctx.members.filter(m => m.name).map(m => m.name) : [], // 회의록 담당 고르기용 (주간보고를 안 쓰는 팀장도 포함)
  });
}

/** 매번 계산해도 가벼운 부분 (접속자, 주차) */
function baseData_(ctx) {
  const me = currentMember_(ctx.members);
  return {
    team: ctx.cfg[CFG.TEAM] || '',
    // 스프레드시트 바로가기는 관리자에게만
    sheetUrl: me.isAdmin ? ctx.ss.getUrl() : '',
    me: me,
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
    // 팀 캘린더가 따로 있으면 일정 추가 때 "팀 / 내 캘린더(나만 보기)"를 고른다
    teamCalendar: teamCalendarSet_(ctx),
    // 채용담당이고 채용 캘린더가 정해져 있으면 [일정]에서 채용 캘린더를 보고 쓸 수 있다
    recruitCalendar: canRecruitCal_(ctx, me),
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
    // 캘린더·휴가 등록에서 자동으로 붙는 휴무 (이번 주~다음 주 / 미리 쓰기는 다음 주~다다음 주)
    leaves: memberLeaves_(ctx, name, ctx.thisWeek.monday, 2),
    preLeaves: memberLeaves_(ctx, name, ctx.nextWeek.monday, 2),
    // 팀장이 굵게 표시한 업무
    marks: marksOf_(ctx)[name] || { thisWeek: [], nextWeek: [] },
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
  return { memberPatch: forViewer_(currentMember_(ctx.members), [patch])[0] };
}

/** 팀장 코멘트는 지금 쓰지 않으므로 누구에게도 보내지 않는다 */
function forViewer_(me, list) {
  return list.map(m => Object.assign({}, m, { comment: '' }));
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
    imp: !!item.imp, // 담당자가 표시한 중요 업무 → 팀 주간보고에서 굵게
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
    // 예산담당자: 예산 메뉴 + [월별 정리]·[예산 설정] 탭 (관리 메뉴·스프레드시트 바로가기는 관리자만)
    budgetManager: !!(m && m.budget),
    recruiter: !!(m && m.recruit),
  };
  me.menus = allowedMenus_(me);
  return me;
}

function readBudget_(ss) {
  // 예전 양식이면 코스트 + 계정코드 + 계정과목 양식으로 바꾼다 (권한이 없으면 관리자가 "권한 한 번에 적용" 때 바뀐다)
  try { migrateBudgetSheets_(ss); } catch (e) { /* 권한 없음 */ }
  const reqSh = ss.getSheetByName(BUDGET.SHEET);
  const reqRows = reqSh && reqSh.getLastRow() >= 2
    ? reqSh.getRange(2, 1, reqSh.getLastRow() - 1, BCOL.MONTH).getValues()
    : [];

  const accounts = getAccounts_();
  // 계정코드가 비어 있는 예전 신청은 [예산과목]의 코드로 보여준다
  const codeOf = (cost, name, v) => String(v || '').trim()
    || (accounts.find(a => a.cost === String(cost).trim() && a.name === String(name).trim()) || {}).code || '';
  const lists = readBudgetLists_(ss);
  return {
    accounts: accounts,
    costs: lists.costs,       // 코스트센터 목록 [{code, name}]
    acctList: lists.acctList, // 계정과목 목록 [{code, name}]
    requests: reqRows
      .map((r, i) => ({ r: r, row: i + 2 }))
      .filter(x => String(x.r[0]).trim() && String(x.r[BCOL.CANCELED - 1]).toUpperCase() !== 'Y')
      .map(x => ({
        row: x.row,
        id: String(x.r[0]),
        date: x.r[1] instanceof Date ? fmt(x.r[1], 'yyyy-MM-dd HH:mm') : String(x.r[1]),
        req: String(x.r[2]),
        fromCost: String(x.r[BCOL.FROM_COST - 1]).trim(),
        fromCode: codeOf(x.r[BCOL.FROM_COST - 1], x.r[BCOL.FROM - 1], x.r[BCOL.FROM_CODE - 1]),
        from: String(x.r[BCOL.FROM - 1]).trim(),
        toCost: String(x.r[BCOL.TO_COST - 1]).trim(),
        toCode: codeOf(x.r[BCOL.TO_COST - 1], x.r[BCOL.TO - 1], x.r[BCOL.TO_CODE - 1]),
        to: String(x.r[BCOL.TO - 1]).trim(),
        amount: Number(x.r[BCOL.AMOUNT - 1]) || 0,
        reason: String(x.r[BCOL.REASON - 1]),
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
    .map(x => [String(x.task).trim(), parseDue_(x.due), x.imp ? INPUT.IMP_MARK : '']);

  const thisRows = toRows(payload.thisWeek);
  const nextRows = toRows(payload.nextWeek);

  // 지우기+쓰기를 한 번에: 빈 줄까지 채워서 통째로 덮어쓴다
  sh.getRange(INPUT.FIRST_ROW, 1, INPUT.ROWS, 3).setValues(padRows_(thisRows));
  sh.getRange(INPUT.FIRST_ROW, 4, INPUT.ROWS, 3).setValues(padRows_(nextRows));
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
    .map(x => [String(x.task).trim(), parseDue_(x.due), x.imp ? INPUT.IMP_MARK : '']);
  // 지우기도 같은 경로: 빈 줄로 통째로 덮어쓴다
  sh.getRange(INPUT.FIRST_ROW, INPUT.PRE_THIS_COL, INPUT.ROWS, 3).setValues(padRows_(toRows(payload.thisWeek)));
  sh.getRange(INPUT.FIRST_ROW, INPUT.PRE_NEXT_COL, INPUT.ROWS, 3).setValues(padRows_(toRows(payload.nextWeek)));
  sh.getRange(INPUT.PRE_NOTE).setValue(payload.clear ? '' : String(payload.note || '').trim());
  sh.getRange(INPUT.PRE_FLAG).setValue(payload.clear ? '' : 'Y');
  return memberPatch_(ctx, me.name);
}

/** 입력 칸 수(INPUT.ROWS)만큼 빈 줄을 채운다 */
function padRows_(rows) {
  const out = rows.slice(0, INPUT.ROWS);
  while (out.length < INPUT.ROWS) out.push(['', '', '']); // 업무 | 기한 | 중요(★)
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

/* ---------- 팀장 강조(굵게) ---------- */

/**
 * 팀장이 팀 주간보고에서 업무를 눌러 굵게 표시한다.
 * [강조]  주차(월요일) | 이름 | 구분(금주/차주) | 업무 | 표시자 | 일시
 * 업무 내용으로 찾기 때문에 팀원이 내용을 고치면 강조가 풀린다. 다음 주차가 되면 새로 시작한다.
 */
const MARK = { SHEET: '강조' };

function markSheet_(ss) {
  let sh = ss.getSheetByName(MARK.SHEET);
  if (sh) return sh;
  sh = ss.insertSheet(MARK.SHEET);
  sh.getRange(1, 1, 1, 6).setValues([['주차(월요일)', '이름', '구분', '업무', '표시자', '일시']]);
  styleHeader_(sh.getRange(1, 1, 1, 6));
  sh.getRange('A2:A').setNumberFormat('@');
  sh.getRange('F2:F').setNumberFormat('yyyy-mm-dd hh:mm');
  [100, 80, 60, 360, 80, 130].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);
  return sh;
}

/** 이번 주 강조: { 이름: { thisWeek: [업무], nextWeek: [업무] } } (요청 안에서 한 번만 읽는다) */
function marksOf_(ctx) {
  if (ctx._marks) return ctx._marks;
  const out = {};
  const sh = ctx.ss.getSheetByName(MARK.SHEET);
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(r => {
      const week = r[0] instanceof Date ? ymd(r[0]) : String(r[0]).trim();
      if (week !== ctx.thisWeek.key || !String(r[3]).trim()) return;
      const m = out[String(r[1]).trim()] = out[String(r[1]).trim()] || { thisWeek: [], nextWeek: [] };
      (String(r[2]).trim() === '차주' ? m.nextWeek : m.thisWeek).push(String(r[3]).trim());
    });
  }
  ctx._marks = out;
  return out;
}

/** 팀장: 업무 강조 켜기/끄기. which: 'this' | 'next' */
function apiToggleMark(name, which, task) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'weekly');
  if (!me.isLeader) throw new Error('팀장만 굵게 표시할 수 있습니다.');
  const label = which === 'next' ? '차주' : '금주';
  const text = String(task || '').trim();
  if (!text) throw new Error('업무를 찾을 수 없습니다.');

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sh = markSheet_(ctx.ss);
    const rows = sh.getLastRow() >= 2 ? sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues() : [];
    const idx = rows.findIndex(r => (r[0] instanceof Date ? ymd(r[0]) : String(r[0]).trim()) === ctx.thisWeek.key
      && String(r[1]).trim() === name && String(r[2]).trim() === label && String(r[3]).trim() === text);
    if (idx >= 0) sh.deleteRow(idx + 2);
    else {
      sh.getRange(sh.getLastRow() + 1, 1).setNumberFormat('@');
      sh.getRange(sh.getLastRow() + 1, 1, 1, 6).setValues([[ctx.thisWeek.key, name, label, text, me.name, new Date()]]);
    }
  } finally {
    lock.releaseLock();
  }
  ctx._marks = null;
  return memberPatch_(ctx, name);
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
  const v = sh.getRange(row, 1, 1, BCOL.MONTH).getValues()[0];
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

// ===================================================================
// BudgetForm.html (스프레드시트 메뉴의 예산전용 신청 화면)
// ===================================================================

const BUDGET_FORM_HTML = "<!DOCTYPE html>\n<html>\n<head>\n  <base target=\"_top\">\n  <style>\n    body { font-family: 'Malgun Gothic', sans-serif; font-size: 13px; color: #222; margin: 0; padding: 4px 8px; }\n    label { display: block; font-weight: bold; margin: 12px 0 4px; }\n    select, input, textarea { width: 100%; box-sizing: border-box; padding: 7px; font-size: 13px; border: 1px solid #bbb; border-radius: 4px; }\n    textarea { height: 90px; resize: vertical; }\n    .hint { color: #666; font-size: 12px; margin-top: 4px; min-height: 16px; }\n    .row { display: flex; gap: 10px; }\n    .row > div { flex: 1; }\n    .actions { margin-top: 18px; display: flex; gap: 8px; justify-content: flex-end; }\n    button { padding: 8px 18px; font-size: 13px; border-radius: 4px; border: 1px solid #bbb; background: #fff; cursor: pointer; }\n    button.primary { background: #1a73e8; border-color: #1a73e8; color: #fff; }\n    button:disabled { opacity: .5; cursor: default; }\n    #msg { margin-top: 12px; padding: 8px; border-radius: 4px; display: none; }\n    #msg.err { display: block; background: #fde2e2; color: #a11; }\n    #msg.ok { display: block; background: #e3f4e1; color: #1b5e20; }\n    h4 { margin: 16px 0 0; font-size: 13px; color: #1a73e8; border-top: 1px solid #eee; padding-top: 10px; }\n    .hint { margin-top: 6px; }\n  </style>\n</head>\n<body>\n  <form id=\"f\" onsubmit=\"submitForm(event)\">\n    <div class=\"row\">\n      <div>\n        <label>신청자</label>\n        <select name=\"requester\" id=\"requester\" required></select>\n      </div>\n      <div>\n        <label>적용월</label>\n        <select name=\"month\" id=\"month\" required></select>\n      </div>\n    </div>\n\n    <h4>변경전 (감액)</h4>\n    <div class=\"row\">\n      <div>\n        <label>변경전 코스트</label>\n        <select name=\"fromCost\" id=\"fromCost\" onchange=\"fillAccounts('from')\"></select>\n      </div>\n      <div>\n        <label>변경전 계정과목</label>\n        <select name=\"from\" id=\"from\" required onchange=\"showAvailable()\"></select>\n      </div>\n    </div>\n    <div class=\"hint\" id=\"fromHint\"></div>\n\n    <h4>변경후 (증액)</h4>\n    <div class=\"row\">\n      <div>\n        <label>변경후 코스트</label>\n        <select name=\"toCost\" id=\"toCost\" onchange=\"fillAccounts('to')\"></select>\n      </div>\n      <div>\n        <label>변경후 계정과목</label>\n        <select name=\"to\" id=\"to\" required></select>\n      </div>\n    </div>\n\n    <label>금액 (원)</label>\n    <input name=\"amount\" id=\"amount\" inputmode=\"numeric\" placeholder=\"예) 500,000\" required oninput=\"formatAmount(this)\">\n\n    <label>사유</label>\n    <textarea name=\"reason\" required placeholder=\"전용이 필요한 사유를 적어주세요\"></textarea>\n\n    <div id=\"msg\"></div>\n\n    <div class=\"actions\">\n      <button type=\"button\" onclick=\"google.script.host.close()\">닫기</button>\n      <button type=\"submit\" class=\"primary\" id=\"submitBtn\">신청</button>\n    </div>\n  </form>\n\n  <script>\n    let accounts = [];\n    const $ = id => document.getElementById(id);\n\n    function option(value, text) {\n      const o = document.createElement('option');\n      o.value = value;\n      o.textContent = text;\n      return o;\n    }\n\n    google.script.run.withSuccessHandler(data => {\n      accounts = data.accounts;\n      const req = $('requester');\n      req.appendChild(option('', '선택'));\n      data.members.forEach(n => req.appendChild(option(n, n)));\n      if (data.me) req.value = data.me;\n      (data.months || []).forEach(m => { $('month').appendChild(option(m.value, m.label)); if (m.current) $('month').value = m.value; });\n\n      // 코스트 목록 (빈 코스트는 \"(코스트 없음)\")\n      const costs = accounts.map(a => a.cost).filter((c, i, all) => all.indexOf(c) === i);\n      ['fromCost', 'toCost'].forEach(id => costs.forEach(c => $(id).appendChild(option(c, c || '(코스트 없음)'))));\n      fillAccounts('from');\n      fillAccounts('to');\n    }).withFailureHandler(showError).getBudgetFormData();\n\n    /** 고른 코스트의 계정과목만 보여준다 */\n    function fillAccounts(side) {\n      const cost = $(side + 'Cost').value;\n      const sel = $(side);\n      sel.innerHTML = '';\n      sel.appendChild(option('', '선택'));\n      accounts.filter(a => a.cost === cost).forEach(a => sel.appendChild(option(a.name, [a.code, a.name].filter(Boolean).join(' '))));\n      if (side === 'from') showAvailable();\n    }\n\n    function showAvailable() {\n      const a = accounts.find(x => x.cost === $('fromCost').value && x.name === $('from').value);\n      $('fromHint').textContent = a ? `${a.code ? '계정코드 ' + a.code + ' · ' : ''}현재예산 ${a.current.toLocaleString()}원` : '';\n    }\n\n    function formatAmount(el) {\n      const n = el.value.replace(/[^0-9]/g, '');\n      el.value = n ? Number(n).toLocaleString() : '';\n    }\n\n    function showError(err) {\n      const m = $('msg');\n      m.className = 'err';\n      m.textContent = err.message || err;\n      $('submitBtn').disabled = false;\n    }\n\n    function submitForm(e) {\n      e.preventDefault();\n      $('submitBtn').disabled = true;\n      $('msg').className = '';\n      google.script.run\n        .withSuccessHandler(id => {\n          const m = $('msg');\n          m.className = 'ok';\n          m.textContent = `신청되었습니다. (신청번호 ${id}) 예산에 바로 반영됩니다.`;\n          $('f').querySelectorAll('input, textarea, select').forEach(el => el.disabled = true);\n        })\n        .withFailureHandler(showError)\n        .submitBudgetTransfer(e.target);\n    }\n  </script>\n</body>\n</html>\n";
