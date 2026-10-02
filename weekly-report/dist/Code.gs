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
  PRE_THIS_LABEL: 'G5',
  PRE_NEXT_LABEL: 'J5',
};

// 주간보고 시트 레이아웃
const REPORT = {
  HEADER_ROW: 4,
  FIRST_ROW: 6,
  COLS: 7,
  COMMENT_COL: 7,
  MIN_BLOCK_ROWS: 4,
};

const PROP = {
  REPORT_WEEK: 'REPORT_WEEK',
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
      return;
    }
  }
  sh.appendRow([key, value]);
}

/** [팀원] 시트: 순서 | 이름 | 이메일 | 역할(팀원/팀장) | 작성대상(Y/N) | 관리자(Y/N) */
function getMembers() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.MEMBERS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues()
    .filter(r => String(r[1]).trim())
    .map(r => ({
      order: Number(r[0]) || 999,
      name: String(r[1]).trim(),
      email: String(r[2]).trim(),
      role: String(r[3]).trim(),
      write: String(r[4]).trim().toUpperCase() !== 'N',
      admin: String(r[5]).trim().toUpperCase() === 'Y',
    }))
    .sort((a, b) => a.order - b.order);
}

function getLeaders(members) {
  return (members || getMembers()).filter(m => m.role === '팀장' && m.email);
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
  const cfg = getConfig();
  const holidays = getHolidaySet();
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
    members: getMembers(),
  };
}

/** 팀원 입력시트 읽기 */
function readInput(ss, name) {
  const sh = ss.getSheetByName(SHEET.INPUT_PREFIX + name);
  if (!sh) return { exists: false, done: false, note: '', thisWeek: [], nextWeek: [], pre: emptyPre_() };
  // 작성완료·특이사항·업무 목록·미리 쓰기를 한 번에 읽는다 (시트 요청 1회)
  const v = sh.getRange(1, 1, INPUT.FIRST_ROW + INPUT.ROWS - 1, INPUT.PRE_NEXT_COL + 1).getValues();
  const at = a1 => { const r = sh.getRange(a1); return v[r.getRow() - 1][r.getColumn() - 1]; };
  const rows = v.slice(INPUT.FIRST_ROW - 1);
  return {
    exists: true,
    done: at(INPUT.DONE_CELL) === true,
    note: String(at(INPUT.NOTE_CELL) || '').trim(),
    thisWeek: rows.filter(r => String(r[0]).trim()).map(r => ({ task: String(r[0]).trim(), due: r[1] })),
    nextWeek: rows.filter(r => String(r[3]).trim()).map(r => ({ task: String(r[3]).trim(), due: r[4] })),
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
    thisWeek: rows.filter(r => String(r[t]).trim()).map(r => ({ task: String(r[t]).trim(), due: r[t + 1] })),
    nextWeek: rows.filter(r => String(r[n]).trim()).map(r => ({ task: String(r[n]).trim(), due: r[n + 1] })),
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
    sh.getRange(1, 1, 8, 2).setValues([
      ['항목', '값'],
      [CFG.TEAM, '인사팀'],
      [CFG.MONDAY, mondayOf(new Date())],
      [CFG.DEADLINE, ''],
      [CFG.DEADLINE_HOUR, 14],
      [CFG.REMIND_HOURS, 3],
      [CFG.CALENDAR, ''],
      [CFG.DRIVE_FOLDER, ''],
    ]);
    sh.getRange('B8').setNote('드라이브에서 폴더를 열었을 때 주소 .../folders/ 뒤의 값. 팀원에게 편집자로 공유해야 합니다.');
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
  setupMenuSheet_(ss);
  ensureHeader_(ss.getSheetByName(SHEET.MEMBERS), 6, '관리자(Y/N)');
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
  [SHEET.CONFIG, SHEET.MEMBERS, SHEET.HOLIDAYS, BUDGET.ACCOUNTS, MENU_SHEET].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) protectSheet_(sh, admins);
  });

  // 주간보고: 관리자만, 단 팀장 코멘트 칸은 팀장 편집 가능
  const report = ss.getSheetByName(SHEET.REPORT) || ss.insertSheet(SHEET.REPORT, 0);
  protectSheetWithLeaderRange_(report, report.getRange(REPORT.FIRST_ROW, REPORT.COMMENT_COL, report.getMaxRows() - REPORT.FIRST_ROW + 1, 1), leaders.concat(admins), false, admins);

  // 예산전용: 웹앱·메뉴로 신청 (직접 수정 시 경고만)
  const budget = ss.getSheetByName(BUDGET.SHEET);
  if (budget) warnOnlyProtect_(budget);

  // 회사 기준: 팀장·관리자만 수정
  [RULES.SHEET, RULES.HISTORY].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) protectSheet_(sh, leaders.concat(admins));
  });

  // 게시판: 누구나 글/댓글 작성 (직접 수정 시 경고만)
  [BOARD.SHEET, BOARD.COMMENTS].forEach(name => {
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
 */

function generateReport() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    buildReport_(getContext());
  } finally {
    lock.releaseLock();
  }
}

function buildReport_(ctx) {
  const ss = ctx.ss;
  const sh = ss.getSheetByName(SHEET.REPORT) || ss.insertSheet(SHEET.REPORT, 0);
  const comments = readLeaderComments_(sh, ctx.thisWeek.key);

  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart().clear();

  // 제목
  sh.getRange('A1').setValue(`■ ${ctx.cfg[CFG.TEAM] || ''} 주간업무 ${deadlineText(ctx)}`)
    .setFontSize(13).setFontWeight('bold');
  sh.getRange('A2')
    .setValue(`※ 각자 [${SHEET.INPUT_PREFIX}이름] 시트에 작성하면 자동 반영됩니다.  이름 칸이 붉으면 작성완료 미체크  |  갱신 ${fmt(new Date(), 'M/dd HH:mm')}`)
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

  // 본문
  const rows = [];
  const blocks = [];
  let row = REPORT.FIRST_ROW;
  ctx.members.filter(m => m.write).forEach(m => {
    const input = readInput(ss, m.name);
    const n = Math.max(REPORT.MIN_BLOCK_ROWS, input.thisWeek.length, input.nextWeek.length);
    for (let i = 0; i < n; i++) {
      const a = input.thisWeek[i];
      const b = input.nextWeek[i];
      rows.push([
        i === 0 ? m.name : '',
        a ? a.task : '', a ? formatDue(a.due) : '',
        b ? b.task : '', b ? formatDue(b.due) : '',
        i === 0 ? input.note : '',
        i === 0 ? (comments[m.name] || '') : '',
      ]);
    }
    blocks.push({ row: row, n: n, done: input.done });
    row += n;
  });

  if (rows.length) {
    const body = sh.getRange(REPORT.FIRST_ROW, 1, rows.length, REPORT.COLS);
    body.setNumberFormat('@').setValues(rows).setVerticalAlignment('middle').setWrap(true);
    [3, 5].forEach(c => sh.getRange(REPORT.FIRST_ROW, c, rows.length, 1).setHorizontalAlignment('center'));
    sh.getRange(REPORT.FIRST_ROW, REPORT.COMMENT_COL, rows.length, 1).setBackground('#fffbea');

    blocks.forEach(b => {
      [1, 6, REPORT.COMMENT_COL].forEach(c => sh.getRange(b.row, c, b.n, 1).merge());
      sh.getRange(b.row, 1, b.n, 1).setHorizontalAlignment('center').setBackground(b.done ? null : '#fde2e2');
    });
  }

  const table = sh.getRange(h, 1, 2 + rows.length, REPORT.COLS);
  table.setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
  blocks.forEach(b => sh.getRange(b.row, 1, b.n, REPORT.COLS)
    .setBorder(true, null, true, null, null, null, '#555555', SpreadsheetApp.BorderStyle.SOLID_MEDIUM));

  [110, 330, 95, 330, 95, 170, 240].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(h + 1);

  PropertiesService.getDocumentProperties().setProperty(PROP.REPORT_WEEK, ctx.thisWeek.key);
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
  const lastNext = sh.getRange(INPUT.FIRST_ROW, 4, INPUT.ROWS, 2).getValues().filter(r => String(r[0]).trim());
  sh.getRange(INPUT.FIRST_ROW, 1, INPUT.ROWS, 2).clearContent();
  sh.getRange(INPUT.FIRST_ROW, 4, INPUT.ROWS, 2).clearContent();

  if (pre.exists) {
    // 미리 써둔 다음 주 보고를 그대로 옮기고 작성완료로 표시
    const toRows = list => list.slice(0, INPUT.ROWS).map(x => [x.task, x.due]);
    const thisRows = toRows(pre.thisWeek), nextRows = toRows(pre.nextWeek);
    if (thisRows.length) sh.getRange(INPUT.FIRST_ROW, 1, thisRows.length, 2).setValues(thisRows);
    if (nextRows.length) sh.getRange(INPUT.FIRST_ROW, 4, nextRows.length, 2).setValues(nextRows);
    sh.getRange(INPUT.DONE_CELL).setValue(true);
    sh.getRange(INPUT.NOTE_CELL).setValue(pre.note);
    clearPre_(sh);
  } else {
    // 기본: 차주 → 금주, "지속" 업무는 차주에도 유지
    const keep = lastNext.filter(r => String(r[1]).trim() === '지속');
    if (lastNext.length) sh.getRange(INPUT.FIRST_ROW, 1, lastNext.length, 2).setValues(lastNext);
    if (keep.length) sh.getRange(INPUT.FIRST_ROW, 4, keep.length, 2).setValues(keep);
    sh.getRange(INPUT.DONE_CELL).setValue(false);
    sh.getRange(INPUT.NOTE_CELL).clearContent();
  }
  writeInputLabels_(sh, ctx);
}

/** 미리 쓰기 칸 비우기 */
function clearPre_(sh) {
  sh.getRange(INPUT.PRE_FLAG).clearContent();
  sh.getRange(INPUT.PRE_NOTE).clearContent();
  sh.getRange(INPUT.FIRST_ROW, INPUT.PRE_THIS_COL, INPUT.ROWS, 2).clearContent();
  sh.getRange(INPUT.FIRST_ROW, INPUT.PRE_NEXT_COL, INPUT.ROWS, 2).clearContent();
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

/** 작성완료 미체크 팀원에게 메일. 보낸 사람 이름 목록 반환 */
function sendReminder_(ctx) {
  const pending = pendingMembers_(ctx);
  pending.filter(m => m.email).forEach(m => {
    const sh = ctx.ss.getSheetByName(SHEET.INPUT_PREFIX + m.name);
    MailApp.sendEmail({
      to: m.email,
      subject: `[${ctx.cfg[CFG.TEAM]}] 주간업무 작성 요청 - ${ctx.thisWeek.label}`,
      htmlBody:
        `${m.name}님, 이번 주 주간업무 작성을 부탁드립니다.<br>` +
        `마감: <b>${fmt(ctx.deadline, 'M/dd')} (${DAY_KO[ctx.deadline.getDay()]}) ${hourKo(ctx.deadline.getHours())}</b><br><br>` +
        `<a href="${sh ? sheetUrl(ctx.ss, sh) : ctx.ss.getUrl()}">내 입력시트 열기</a><br>` +
        `작성 후 "작성완료 체크"를 눌러주세요.`,
    });
  });
  return pending.filter(m => m.email).map(m => m.name);
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
    to: leaders.map(m => m.email).join(','),
    subject: `[${ctx.cfg[CFG.TEAM]}] 주간업무 취합 완료 - ${ctx.thisWeek.label}`,
    htmlBody:
      `${ctx.thisWeek.label} 주간업무가 취합되었습니다.<br><br>` +
      `<a href="${sheetUrl(ctx.ss, report)}">주간보고 열기</a> (팀장 코멘트 칸에 의견을 남길 수 있습니다)<br><br>` +
      (pending.length ? `⚠️ 작성완료 미체크: ${pending.join(', ')}` : '✅ 전원 작성완료'),
  });
}

/* ---------- 트리거 ---------- */

/**
 * 매시간 실행
 *  - 평일 업무시간엔 보고서 자동 새로고침
 *  - 마감 N시간 전 미작성자 알림 (주 1회)
 *  - 마감 후 팀장에게 취합 메일 (주 1회)
 */
function hourlyCheck() {
  const now = new Date();
  const ctx = getContext();
  const props = PropertiesService.getDocumentProperties();
  const key = ctx.thisWeek.key;

  const day = now.getDay();
  if (day >= 1 && day <= 5 && now.getHours() >= 8 && now.getHours() <= 20) generateReport();

  const remindHours = Number(ctx.cfg[CFG.REMIND_HOURS]) || 3;
  const remindAt = new Date(ctx.deadline.getTime() - remindHours * 3600 * 1000);
  if (now >= remindAt && now < ctx.deadline && props.getProperty(PROP.REMINDED) !== key) {
    sendReminder_(ctx);
    props.setProperty(PROP.REMINDED, key);
  }

  if (now >= ctx.deadline && props.getProperty(PROP.CLOSED) !== key) {
    generateReport();
    sendClosingMail_(ctx);
    props.setProperty(PROP.CLOSED, key);
  }
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
 * [예산과목]  과목명 | 편성액 | 전입 | 전출 | 현재예산
 * [예산전용]  신청번호 | 신청일시 | 신청자 | 전출과목 | 전입과목 | 금액 | 사유 | 취소(Y)
 *
 * 신청하면 바로 예산에 반영된다. 잘못 신청한 건은 신청자 본인이나 팀장·관리자가 취소한다.
 */

const BUDGET = {
  SHEET: '예산전용',
  ACCOUNTS: '예산과목',
  ACCOUNT_ROWS: 50,
};

const BCOL = { ID: 1, DATE: 2, REQUESTER: 3, FROM: 4, TO: 5, AMOUNT: 6, REASON: 7, CANCELED: 8, FILES: 9 };

function setupBudgetSheets_(ss) {
  if (!ss.getSheetByName(BUDGET.SHEET)) {
    const sh = ss.insertSheet(BUDGET.SHEET);
    sh.getRange(1, 1, 1, 8).setValues([['신청번호', '신청일시', '신청자', '전출과목(감액)', '전입과목(증액)', '금액', '사유', '취소(Y)']]);
    styleHeader_(sh.getRange(1, 1, 1, 8));
    sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('F2:F').setNumberFormat('#,##0');
    sh.getRange('G2:G').setWrap(true);
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$H2="Y"')
        .setFontColor('#999999').setStrikethrough(true).setRanges([sh.getRange('A2:H')]).build(),
    ]);
    [110, 130, 80, 130, 130, 110, 320, 70].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
    sh.getRange('H1').setNote('Y면 취소된 건으로 예산에 반영되지 않습니다.');
  }

  ensureHeader_(ss.getSheetByName(BUDGET.SHEET), BCOL.FILES, '증빙 첨부');

  if (!ss.getSheetByName(BUDGET.ACCOUNTS)) {
    const sh = ss.insertSheet(BUDGET.ACCOUNTS);
    sh.getRange(1, 1, 1, 5).setValues([['과목명', '편성액', '전입', '전출', '현재예산']]);
    styleHeader_(sh.getRange(1, 1, 1, 5));

    const samples = ['교육훈련비', '복리후생비', '채용비', '행사비', '회의비', '도서인쇄비'];
    sh.getRange(2, 1, samples.length, 2).setValues(samples.map(n => [n, 0]));

    const b = `'${BUDGET.SHEET}'!`;
    const formulas = [];
    for (let r = 2; r < 2 + BUDGET.ACCOUNT_ROWS; r++) {
      formulas.push([
        `=IF($A${r}="","",SUMIFS(${b}$F:$F,${b}$E:$E,$A${r},${b}$H:$H,"<>Y"))`,
        `=IF($A${r}="","",SUMIFS(${b}$F:$F,${b}$D:$D,$A${r},${b}$H:$H,"<>Y"))`,
        `=IF($A${r}="","",N($B${r})+C${r}-D${r})`,
      ]);
    }
    sh.getRange(2, 3, BUDGET.ACCOUNT_ROWS, 3).setFormulas(formulas);
    sh.getRange(2, 2, BUDGET.ACCOUNT_ROWS, 4).setNumberFormat('#,##0');
    sh.getRange(2, 5, BUDGET.ACCOUNT_ROWS, 1).setFontWeight('bold');
    sh.setColumnWidth(1, 150);
    sh.setFrozenRows(1);
    sh.getRange('B1').setNote('연간(또는 기간) 편성 예산을 입력하세요.');
  }
}

/* ---------- 신청 폼 (스프레드시트 메뉴) ---------- */

function openBudgetForm() {
  // 한 파일 설치본(dist/Code.gs)은 화면을 BUDGET_FORM_HTML 문자열로 갖고 있다
  const html = (typeof BUDGET_FORM_HTML === 'string'
    ? HtmlService.createHtmlOutput(BUDGET_FORM_HTML)
    : HtmlService.createHtmlOutputFromFile('BudgetForm')).setWidth(480).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '💰 예산전용 신청');
}

/** 폼 초기값: 팀원 목록, 로그인 사용자, 과목별 현재예산 */
function getBudgetFormData() {
  const members = getMembers();
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const me = members.find(m => m.email && m.email.toLowerCase() === email);
  return {
    members: members.map(m => m.name),
    me: me ? me.name : '',
    accounts: getAccounts_(),
  };
}

function getAccounts_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(BUDGET.ACCOUNTS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({
      name: String(r[0]).trim(),
      budget: Number(r[1]) || 0,
      inAmt: Number(r[2]) || 0,
      outAmt: Number(r[3]) || 0,
      current: Number(r[4]) || 0,
    }));
}

/** 신청 → [예산전용] 시트에 한 줄 추가 + 팀장에게 알림 메일. 신청번호 반환 */
function submitBudgetTransfer(form) {
  const requester = String(form.requester || '').trim();
  const from = String(form.from || '').trim();
  const to = String(form.to || '').trim();
  const reason = String(form.reason || '').trim();
  const amount = Math.round(Number(String(form.amount || '').replace(/,/g, '')));

  if (!requester || !from || !to || !reason) throw new Error('모든 항목을 입력해 주세요.');
  if (from === to) throw new Error('전출과목과 전입과목이 같습니다.');
  if (!(amount > 0)) throw new Error('금액을 올바르게 입력해 주세요.');

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const account = getAccounts_().find(a => a.name === from);
    if (!account) throw new Error(`[${from}] 과목을 찾을 수 없습니다.`);
    if (amount > account.current) {
      throw new Error(`현재예산을 초과합니다. (${from} 현재예산: ${account.current.toLocaleString()}원)`);
    }

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(BUDGET.SHEET);
    const now = new Date();
    const id = nextBudgetId_(sh, now);
    sh.getRange(sh.getLastRow() + 1, 1, 1, BCOL.FILES).setValues([[id, now, requester, from, to, amount, reason, '', filesToCell_(form.files)]]);
    SpreadsheetApp.flush();
    try { CacheService.getDocumentCache().remove(CK.budget); } catch (e) { /* 캐시 없음 */ }

    const leaders = getLeaders();
    if (leaders.length) {
      MailApp.sendEmail({
        to: leaders.map(m => m.email).join(','),
        subject: `[예산전용 신청] ${id} ${requester} - ${from} → ${to} ${amount.toLocaleString()}원`,
        htmlBody:
          `<b>${requester}</b>님이 예산전용을 신청했습니다.<br><br>` +
          `신청번호: ${id}<br>전출과목: ${from}<br>전입과목: ${to}<br>` +
          `금액: <b>${amount.toLocaleString()}원</b><br>사유: ${escapeHtml_(reason)}<br>` +
          filesFromCell_(filesToCell_(form.files)).map(f => `첨부: <a href="${f.url}">${escapeHtml_(f.name)}</a><br>`).join('') + '<br>' +
          `<a href="${sheetUrl(ss, sh)}">예산전용 시트 열기</a>`,
      });
    }
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

function getTeamCalendar_() {
  const id = String(getConfig()[CFG.CALENDAR] || '').trim();
  const cal = id ? CalendarApp.getCalendarById(id) : CalendarApp.getDefaultCalendar();
  if (!cal) throw new Error('팀 캘린더에 접근할 수 없습니다. 캘린더가 공유되어 있는지 관리자에게 확인하세요.');
  return cal;
}

/** startIso ~ endIso 사이 일정 + 공휴일 + 주간보고 마감 */
function apiCalendar(startIso, endIso) {
  requireMenu_(currentMember_(getMembers()), 'calendar');
  const start = new Date(startIso);
  const end = new Date(endIso);
  const cal = getTeamCalendar_();

  const events = cal.getEvents(start, end).map(e => ({
    id: e.getId(),
    title: e.getTitle(),
    start: e.getStartTime().toISOString(),
    end: e.getEndTime().toISOString(),
    allDay: e.isAllDayEvent(),
    location: e.getLocation() || '',
    desc: e.getDescription() || '',
    kind: 'event',
  }));

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.HOLIDAYS);
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([d, name]) => {
      if (d instanceof Date && d >= addDays(start, -1) && d < end) {
        events.push({ id: 'h-' + ymd(d), title: String(name), start: ymd(d), end: ymd(addDays(d, 1)), allDay: true, location: '', desc: '', kind: 'holiday' });
      }
    });
  }

  const ctx = getContext();
  if (ctx.deadline >= start && ctx.deadline < end) {
    events.push({ id: 'deadline', title: '주간보고 마감', start: ctx.deadline.toISOString(), end: ctx.deadline.toISOString(), allDay: false, location: '', desc: '', kind: 'deadline' });
  }

  return { calendarName: cal.getName(), events: events };
}

/**
 * ev: { title, date:'yyyy-mm-dd', allDay, startTime:'HH:mm', endTime:'HH:mm', location, desc }
 */
function apiAddEvent(ev) {
  const title = String(ev.title || '').trim();
  const m = String(ev.date || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!title || !m) throw new Error('제목과 날짜를 입력해 주세요.');

  const me = currentMember_(getMembers());
  requireMenu_(me, 'calendar');
  const cal = getTeamCalendar_();
  const opts = {
    location: String(ev.location || '').trim(),
    description: [String(ev.desc || '').trim(), me.name ? `등록: ${me.name}` : ''].filter(Boolean).join('\n'),
  };
  const day = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  try {
    if (ev.allDay) {
      cal.createAllDayEvent(title, day, opts);
    } else {
      const at = (t, fallback) => {
        const [h, mi] = String(t || fallback).split(':').map(Number);
        const d = new Date(day.getTime());
        d.setHours(h || 0, mi || 0, 0, 0);
        return d;
      };
      const s = at(ev.startTime, '09:00');
      let e = at(ev.endTime, '10:00');
      if (e <= s) e = new Date(s.getTime() + 3600 * 1000);
      cal.createEvent(title, s, e, opts);
    }
  } catch (err) {
    throw new Error('일정을 추가할 권한이 없습니다. 팀 캘린더의 "일정 변경" 권한을 관리자에게 요청하세요.');
  }
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
  if (!me.role) return [];
  const access = getMenuAccess_()[me.role] || {};
  return MENUS.map(m => m.key).filter(k => access[k] !== false);
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

/**
 * useCache=true면 공유 상태(시트·폴더·캘린더 확인, 가장 느린 부분)를 5분 캐시에서 쓴다.
 */
function adminData_(ctx, useCache) {
  const cfg = ctx.cfg;
  return {
    members: ctx.members.map(m => ({ order: m.order, name: m.name, email: m.email, role: m.role, write: m.write, admin: m.admin })),
    menus: MENUS,
    roles: ROLES,
    menuAccess: getMenuAccess_(),
    settings: {
      team: String(cfg[CFG.TEAM] || ''),
      deadlineHour: Number(cfg[CFG.DEADLINE_HOUR]) || 14,
      remindHours: Number(cfg[CFG.REMIND_HOURS]) || 3,
      calendarId: String(cfg[CFG.CALENDAR] || ''),
      folderId: String(cfg[CFG.DRIVE_FOLDER] || ''),
    },
    webAppUrl: ScriptApp.getService().getUrl() || '',
    status: useCache ? cached_('access', () => accessStatus_(ctx), 300) : freshStatus_(ctx),
  };
}

function freshStatus_(ctx) {
  const status = accessStatus_(ctx);
  putCache_('access', status, 300);
  return status;
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

  const people = {};
  ctx.members.filter(m => m.email).forEach(m => {
    const e = m.email.toLowerCase();
    const sh = m.write ? ctx.ss.getSheetByName(SHEET.INPUT_PREFIX + m.name) : null;
    let inputProtected = false;
    if (sh) {
      const p = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
      inputProtected = !!p && p.getEditors().some(u => u.getEmail().toLowerCase() === e);
    }
    people[m.name] = {
      sheet: level(e, fileOwner, fileEditors, fileViewers),
      input: !m.write ? 'n/a' : !sh ? 'missing' : inputProtected ? 'ok' : 'unprotected',
      folder: res.folder.ok ? level(e, folderOwner, folderEditors, folderViewers) : 'n/a',
      calendar: res.calendar.ok ? (acl[e] || 'none') : 'n/a',
    };
  });
  return { resources: res, people: people };
}

/* ---------- 저장 ---------- */

/** 팀원 목록 전체 저장. list: [{order, name, email, role, write, admin}] */
function apiSaveMembers(list) {
  const ctx = getContext();
  requireAdmin_(ctx);
  ensureHeader_(ctx.ss.getSheetByName(SHEET.MEMBERS), 6, '관리자(Y/N)');

  const rows = (list || []).map((m, i) => ({
    order: Number(m.order) || i + 1,
    name: String(m.name || '').trim(),
    email: String(m.email || '').trim().toLowerCase(),
    role: ROLES.indexOf(m.role) >= 0 ? m.role : '팀원',
    write: !!m.write,
    admin: !!m.admin,
  })).filter(m => m.name);

  const names = {};
  rows.forEach(m => {
    if (names[m.name]) throw new Error(`이름이 겹칩니다: ${m.name}. 동명이인은 "김민지A"처럼 구분해 주세요.`);
    names[m.name] = true;
    if (m.email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(m.email)) throw new Error(`이메일 형식을 확인해 주세요: ${m.name} (${m.email})`);
  });

  const sh = ctx.ss.getSheetByName(SHEET.MEMBERS);
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 6).setValues(rows.map(m => [m.order, m.name, m.email, m.role, m.write ? 'Y' : 'N', m.admin ? 'Y' : 'N']));
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
  setConfigValue(CFG.DRIVE_FOLDER, String(s.folderId || '').trim().replace(/^.*\/folders\//, '').replace(/[?#].*$/, ''));
  clearCaches_(ctx);
  return adminData_(getContext(), false);
}

function clearCaches_(ctx) {
  try {
    CacheService.getDocumentCache().removeAll([CK.members(ctx), CK.budget, CK.board, CK.rules, 'menus']);
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
  return HtmlService.createHtmlOutputFromFile('App')
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
    ? reqSh.getRange(2, 1, reqSh.getLastRow() - 1, BCOL.FILES).getValues()
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
      }))
      .reverse(),
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

  sh.getRange(INPUT.FIRST_ROW, 1, INPUT.ROWS, 2).clearContent();
  sh.getRange(INPUT.FIRST_ROW, 4, INPUT.ROWS, 2).clearContent();
  if (thisRows.length) sh.getRange(INPUT.FIRST_ROW, 1, thisRows.length, 2).setValues(thisRows);
  if (nextRows.length) sh.getRange(INPUT.FIRST_ROW, 4, nextRows.length, 2).setValues(nextRows);
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

  clearPre_(sh);
  if (!payload.clear) {
    const toRows = list => (list || [])
      .filter(x => String(x.task || '').trim())
      .slice(0, INPUT.ROWS)
      .map(x => [String(x.task).trim(), parseDue_(x.due)]);
    const thisRows = toRows(payload.thisWeek), nextRows = toRows(payload.nextWeek);
    if (thisRows.length) sh.getRange(INPUT.FIRST_ROW, INPUT.PRE_THIS_COL, thisRows.length, 2).setValues(thisRows);
    if (nextRows.length) sh.getRange(INPUT.FIRST_ROW, INPUT.PRE_NEXT_COL, nextRows.length, 2).setValues(nextRows);
    sh.getRange(INPUT.PRE_NOTE).setValue(String(payload.note || '').trim());
    sh.getRange(INPUT.PRE_FLAG).setValue('Y');
  }
  return memberPatch_(ctx, me.name);
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

// ===================================================================
// BudgetForm.html (스프레드시트 메뉴의 예산전용 신청 화면)
// ===================================================================

const BUDGET_FORM_HTML = "<!DOCTYPE html>\n<html>\n<head>\n  <base target=\"_top\">\n  <style>\n    body { font-family: 'Malgun Gothic', sans-serif; font-size: 13px; color: #222; margin: 0; padding: 4px 8px; }\n    label { display: block; font-weight: bold; margin: 12px 0 4px; }\n    select, input, textarea { width: 100%; box-sizing: border-box; padding: 7px; font-size: 13px; border: 1px solid #bbb; border-radius: 4px; }\n    textarea { height: 90px; resize: vertical; }\n    .hint { color: #666; font-size: 12px; margin-top: 4px; min-height: 16px; }\n    .row { display: flex; gap: 10px; }\n    .row > div { flex: 1; }\n    .actions { margin-top: 18px; display: flex; gap: 8px; justify-content: flex-end; }\n    button { padding: 8px 18px; font-size: 13px; border-radius: 4px; border: 1px solid #bbb; background: #fff; cursor: pointer; }\n    button.primary { background: #1a73e8; border-color: #1a73e8; color: #fff; }\n    button:disabled { opacity: .5; cursor: default; }\n    #msg { margin-top: 12px; padding: 8px; border-radius: 4px; display: none; }\n    #msg.err { display: block; background: #fde2e2; color: #a11; }\n    #msg.ok { display: block; background: #e3f4e1; color: #1b5e20; }\n  </style>\n</head>\n<body>\n  <form id=\"f\" onsubmit=\"submitForm(event)\">\n    <label>신청자</label>\n    <select name=\"requester\" id=\"requester\" required></select>\n\n    <div class=\"row\">\n      <div>\n        <label>전출과목 (감액)</label>\n        <select name=\"from\" id=\"from\" required onchange=\"showAvailable()\"></select>\n        <div class=\"hint\" id=\"fromHint\"></div>\n      </div>\n      <div>\n        <label>전입과목 (증액)</label>\n        <select name=\"to\" id=\"to\" required></select>\n      </div>\n    </div>\n\n    <label>금액 (원)</label>\n    <input name=\"amount\" id=\"amount\" inputmode=\"numeric\" placeholder=\"예) 500,000\" required oninput=\"formatAmount(this)\">\n\n    <label>사유</label>\n    <textarea name=\"reason\" required placeholder=\"전용이 필요한 사유를 적어주세요\"></textarea>\n\n    <div id=\"msg\"></div>\n\n    <div class=\"actions\">\n      <button type=\"button\" onclick=\"google.script.host.close()\">닫기</button>\n      <button type=\"submit\" class=\"primary\" id=\"submitBtn\">신청</button>\n    </div>\n  </form>\n\n  <script>\n    let accounts = [];\n\n    function option(value, text) {\n      const o = document.createElement('option');\n      o.value = value;\n      o.textContent = text;\n      return o;\n    }\n\n    google.script.run.withSuccessHandler(data => {\n      accounts = data.accounts;\n      const req = document.getElementById('requester');\n      req.appendChild(option('', '선택'));\n      data.members.forEach(n => req.appendChild(option(n, n)));\n      if (data.me) req.value = data.me;\n\n      ['from', 'to'].forEach(id => {\n        const sel = document.getElementById(id);\n        sel.appendChild(option('', '선택'));\n        accounts.forEach(a => sel.appendChild(option(a.name, a.name)));\n      });\n    }).withFailureHandler(showError).getBudgetFormData();\n\n    function showAvailable() {\n      const a = accounts.find(x => x.name === document.getElementById('from').value);\n      document.getElementById('fromHint').textContent = a ? `현재예산 ${a.current.toLocaleString()}원` : '';\n    }\n\n    function formatAmount(el) {\n      const n = el.value.replace(/[^0-9]/g, '');\n      el.value = n ? Number(n).toLocaleString() : '';\n    }\n\n    function showError(err) {\n      const m = document.getElementById('msg');\n      m.className = 'err';\n      m.textContent = err.message || err;\n      document.getElementById('submitBtn').disabled = false;\n    }\n\n    function submitForm(e) {\n      e.preventDefault();\n      document.getElementById('submitBtn').disabled = true;\n      document.getElementById('msg').className = '';\n      google.script.run\n        .withSuccessHandler(id => {\n          const m = document.getElementById('msg');\n          m.className = 'ok';\n          m.textContent = `신청되었습니다. (신청번호 ${id}) 예산에 바로 반영됩니다.`;\n          document.getElementById('f').querySelectorAll('input, textarea, select').forEach(el => el.disabled = true);\n        })\n        .withFailureHandler(showError)\n        .submitBudgetTransfer(e.target);\n    }\n  </script>\n</body>\n</html>\n";
