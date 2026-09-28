/**
 * 메뉴 / 초기 설정 / 권한
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 주간업무')
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

/** 1단계: 설정/팀원/공휴일/예산 시트 생성 (이미 있으면 건너뜀) */
function initialize() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);

  if (!ss.getSheetByName(SHEET.CONFIG)) {
    const sh = ss.insertSheet(SHEET.CONFIG);
    sh.getRange(1, 1, 6, 2).setValues([
      ['항목', '값'],
      [CFG.TEAM, '인사팀'],
      [CFG.MONDAY, mondayOf(new Date())],
      [CFG.DEADLINE, ''],
      [CFG.DEADLINE_HOUR, 14],
      [CFG.REMIND_HOURS, 3],
    ]);
    sh.getRange('B3').setNumberFormat('yyyy-mm-dd (ddd)');
    sh.getRange('B4').setNumberFormat('yyyy-mm-dd hh:mm').setNote('명절 등으로 이번 주만 마감이 다르면 입력. 예) 2026-09-23 14:00\n다음 주차로 넘어가면 자동으로 비워집니다.');
    styleHeader_(sh.getRange('A1:B1'));
    sh.setColumnWidth(1, 240).setColumnWidth(2, 200);
  }

  if (!ss.getSheetByName(SHEET.MEMBERS)) {
    const sh = ss.insertSheet(SHEET.MEMBERS);
    const rows = [['순서', '이름', '이메일', '역할', '작성대상(Y/N)'], [0, '팀장', '', '팀장', 'N']];
    for (let i = 1; i <= 8; i++) rows.push([i, `팀원${i}`, '', '팀원', 'Y']);
    sh.getRange(1, 1, rows.length, 5).setValues(rows);
    styleHeader_(sh.getRange('A1:E1'));
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
  if (!ss.getSheetByName(SHEET.REPORT)) ss.insertSheet(SHEET.REPORT, 0);

  SpreadsheetApp.getUi().alert('기본 시트를 만들었습니다.\n\n[팀원] 시트에 이름·이메일을, [예산과목] 시트에 과목·편성액을 입력한 뒤\n"2. 팀원 입력시트 만들기 + 권한 적용"을 실행하세요.');
}

// 공휴일 기본값 (정부 발표에 따라 변경될 수 있으니 확인 후 수정하세요)
const DEFAULT_HOLIDAYS = [
  ['2026-09-24', '추석'], ['2026-09-25', '추석'],
  ['2026-10-05', '개천절 대체공휴일'], ['2026-10-09', '한글날'], ['2026-12-25', '성탄절'],
  ['2027-01-01', '신정'], ['2027-02-08', '설날'], ['2027-02-09', '설날 대체공휴일'],
  ['2027-03-01', '삼일절'], ['2027-05-05', '어린이날'], ['2027-05-13', '부처님오신날'],
  ['2027-08-16', '광복절 대체공휴일'], ['2027-09-14', '추석'], ['2027-09-15', '추석'],
  ['2027-09-16', '추석'], ['2027-10-04', '개천절 대체공휴일'], ['2027-10-11', '한글날 대체공휴일'],
  ['2027-12-27', '성탄절 대체공휴일'],
];

/** 2단계: 팀원별 입력시트 생성 + 시트 보호(본인만 편집) */
function setupMembers() {
  const ctx = getContext();
  const ss = ctx.ss;
  const leaders = getLeaders(ctx.members).map(m => m.email);
  const missing = [];

  ctx.members.filter(m => m.write).forEach(m => {
    const sh = setupInputSheet_(ss, m.name, ctx);
    if (m.email) protectSheet_(sh, [m.email]);
    else missing.push(m.name);
  });

  // 관리 시트: 관리자만
  [SHEET.CONFIG, SHEET.MEMBERS, SHEET.HOLIDAYS, BUDGET.ACCOUNTS].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) protectSheet_(sh, []);
  });

  // 주간보고: 관리자만, 단 팀장 코멘트 칸은 팀장 편집 가능
  const report = ss.getSheetByName(SHEET.REPORT) || ss.insertSheet(SHEET.REPORT, 0);
  protectSheetWithLeaderRange_(report, report.getRange(REPORT.FIRST_ROW, REPORT.COMMENT_COL, report.getMaxRows() - REPORT.FIRST_ROW + 1, 1), leaders, false);

  // 예산전용: 신청 칸(A~G)은 경고만, 결재 칸(H~J)은 팀장만
  const budget = ss.getSheetByName(BUDGET.SHEET);
  protectSheetWithLeaderRange_(budget, budget.getRange(2, BCOL.STATUS, budget.getMaxRows() - 1, 3), leaders, true);

  generateReport();
  const msg = missing.length ? `\n\n⚠️ 이메일이 없어 권한을 적용하지 못한 팀원: ${missing.join(', ')}` : '';
  SpreadsheetApp.getUi().alert(`입력시트와 권한을 적용했습니다.\n스프레드시트를 팀원 전원에게 "편집자"로 공유해야 입력할 수 있습니다.${msg}`);
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
  return sh;
}

function writeInputLabels_(sh, ctx) {
  sh.getRange(INPUT.THIS_LABEL).setValue(ctx.thisWeek.label);
  sh.getRange(INPUT.NEXT_LABEL).setValue(ctx.nextWeek.label);
}

function styleHeader_(range) {
  range.setBackground('#efefef').setFontWeight('bold').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
}

/* ---------- 보호 ---------- */

/** 시트 전체 보호: 실행한 관리자(소유자) + editors만 편집 */
function protectSheet_(sh, editors) {
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  const p = sh.protect().setDescription(`${sh.getName()} 보호`);
  restrictEditors_(p, editors);
  return p;
}

/**
 * 시트는 관리자만(warningOnly면 경고만), leaderRange는 팀장만 편집.
 */
function protectSheetWithLeaderRange_(sh, leaderRange, leaders, warningOnly) {
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

  const sp = sh.protect().setDescription(`${sh.getName()} 보호`);
  if (warningOnly) {
    sp.setWarningOnly(true);
  } else {
    restrictEditors_(sp, []);
    sp.setUnprotectedRanges([leaderRange]);
  }

  const rp = leaderRange.protect().setDescription(`${sh.getName()} 팀장 전용`);
  restrictEditors_(rp, leaders);
}

function restrictEditors_(p, editors) {
  const me = Session.getEffectiveUser().getEmail();
  p.addEditor(me);
  p.removeEditors(p.getEditors().map(u => u.getEmail()).filter(e => e && e !== me));
  if (p.canDomainEdit()) p.setDomainEdit(false);
  const list = editors.filter(e => e && e !== me);
  if (list.length) p.addEditors(list);
}
