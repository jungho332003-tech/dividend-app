/**
 * 메뉴 / 초기 설정 / 시트 보호
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🧾 연말정산')
    .addItem('🌐 검토 관리 웹앱 열기', 'showWebAppLink')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('⚙️ 관리자 설정')
      .addItem('1. 기본 시트 만들기', 'initialize')
      .addItem('2. 시트 보호 적용', 'setupProtections'))
    .addToUi();
}

function showWebAppLink() {
  const url = ScriptApp.getService().getUrl();
  const html = url
    ? `<p style="font-family:sans-serif">아래 주소를 담당자에게 공유하세요.</p><p><a href="${url}" target="_blank">${url}</a></p>`
    : '<p style="font-family:sans-serif">아직 웹앱으로 배포되지 않았습니다. README의 "웹앱 배포" 순서를 따라 주세요.</p>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(460).setHeight(160), '연말정산 검토 관리');
}

/** 1단계: 필요한 시트를 만든다 (이미 있으면 건너뛰고, 빠진 열만 채운다) */
function initialize() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);

  if (!ss.getSheetByName(SHEET.CONFIG)) {
    const sh = ss.insertSheet(SHEET.CONFIG);
    sh.getRange(1, 1, 4, 2).setValues([
      ['항목', '값'],
      [CFG.TEAM, '연말정산 검토'],
      [CFG.YEAR, defaultYear_()],
      [CFG.DRIVE_FOLDER, ''],
    ]);
    sh.getRange('B4').setNote('드라이브에서 폴더를 열었을 때 주소 .../folders/ 뒤의 값. 증빙 서류가 올라가므로 담당자에게만 공유하세요.');
    styleHeader_(sh.getRange('A1:B1'));
    sh.setColumnWidth(1, 240).setColumnWidth(2, 260);
  }

  if (!ss.getSheetByName(SHEET.MEMBERS)) {
    const sh = ss.insertSheet(SHEET.MEMBERS);
    const rows = [['순서', '이름', '이메일', '역할', '관리자(Y/N)', '열람범위'], [1, '총괄', '', '총괄', 'Y', '전체']];
    for (let i = 1; i <= 4; i++) rows.push([i + 1, `담당자${i}`, '', '응대담당', 'N', '본인']);
    sh.getRange(1, 1, rows.length, 6).setValues(rows);
    styleHeader_(sh.getRange('A1:F1'));
    sh.getRange('F1').setNote('전체 = 모든 대상자 / 본인 = 내 담당만 / 본인,이수민,(미배정) = 내 담당 + 고른 담당자의 대상자. 웹앱 권한 관리 > 열람 범위에서 바꾸는 것을 권장합니다.');
    sh.getRange('D2:D100').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ROLES).build());
    sh.getRange('E2:E100').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Y', 'N']).build());
    sh.setColumnWidth(2, 120).setColumnWidth(3, 240).setColumnWidth(5, 100);
    sh.setFrozenRows(1);
  }

  ensureHeader_(ss.getSheetByName(SHEET.MEMBERS), 6, '열람범위');
  setupPeopleSheets_(ss);
  setupNoticeSheets_(ss);
  setupRuleSheets_(ss);
  setupFileSheet_(ss);
  setupMenuSheet_(ss);
  clearCaches_();

  SpreadsheetApp.getUi().alert('기본 시트를 만들었습니다.\n\n' +
    '1. [담당자] 시트에 이름·이메일·역할을 입력하세요.\n' +
    '2. 기존 검토 시트 내용을 [대상자] 시트 2행부터 붙여넣으세요. (열 이름이 같으면 순서가 달라도 됩니다)\n' +
    '3. 웹앱을 배포한 뒤 관리 메뉴에서 "권한 한 번에 적용"을 누르세요.');
}

/** 2단계: 시트 보호만 다시 적용 (웹앱 관리 메뉴의 "권한 한 번에 적용"도 같은 일을 한다) */
function setupProtections() {
  applyProtections_(getContext());
  SpreadsheetApp.getUi().alert('시트 보호를 적용했습니다.\n담당자에게 스프레드시트를 "편집자"로 공유해야 웹앱에서 저장할 수 있습니다.');
}

/**
 * 시트 보호
 *  - 설정·담당자·메뉴권한: 관리자만
 *  - 연말정산 기준: 총괄·관리자만
 *  - 대상자·응대기록·공지·첨부: 웹앱으로 작성 (시트에서 직접 고치면 경고만)
 */
function applyProtections_(ctx) {
  const ss = ctx.ss;
  const admins = ctx.members.filter(m => m.admin && m.email).map(m => m.email);
  const leaders = ctx.members.filter(m => m.role === LEADER_ROLE && m.email).map(m => m.email);

  [SHEET.CONFIG, SHEET.MEMBERS, MENU_SHEET].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) protectSheet_(sh, admins);
  });
  [RULES.SHEET, RULES.HISTORY].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) protectSheet_(sh, leaders.concat(admins));
  });
  [SHEET.PEOPLE, SHEET.LOGS, NOTICE.SHEET, NOTICE.COMMENTS, FILES.SHEET].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) warnOnlyProtect_(sh);
  });
}

function styleHeader_(range) {
  range.setBackground('#efefef').setFontWeight('bold').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
}

/** 예전에 만든 시트에 열 제목이 없으면 채운다 */
function ensureHeader_(sh, col, title) {
  const cell = sh.getRange(1, col);
  if (!cell.getValue()) {
    cell.setValue(title);
    styleHeader_(cell);
    sh.setColumnWidth(col, 220);
  }
}

/* ---------- 보호 ---------- */

/** 시트 전체 보호: 실행한 관리자(소유자) + editors만 편집 */
function protectSheet_(sh, editors) {
  const current = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (current.length === 1 && !current[0].isWarningOnly() && sameEditors_(current[0], editors)) return current[0];
  current.forEach(p => p.remove());
  const p = sh.protect().setDescription(`${sh.getName()} 보호`);
  restrictEditors_(p, editors);
  return p;
}

/** 보호의 편집자가 (실행한 관리자 + 소유자 + editors)와 정확히 같은지 */
function sameEditors_(p, editors) {
  if (p.canDomainEdit()) return false;
  const want = new Set(editors.concat(Session.getEffectiveUser().getEmail(), ownerEmail_())
    .filter(Boolean).map(e => e.toLowerCase()));
  const have = new Set(p.getEditors().map(u => u.getEmail().toLowerCase()));
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
