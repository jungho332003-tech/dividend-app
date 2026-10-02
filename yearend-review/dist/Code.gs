/**
 * 연말정산 검토 관리 - 한 파일 설치본 (build.py로 자동 생성, 직접 고치지 말고 원본 .gs 파일을 고친 뒤 다시 생성)
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
  MEMBERS: '담당자',
  PEOPLE: '대상자',
  LOGS: '응대기록',
};

// [설정] 시트의 항목명 (A열) — 값은 B열
const CFG = {
  TEAM: '업무명',
  YEAR: '귀속연도',
  DRIVE_FOLDER: '첨부 폴더 ID(구글 드라이브)',
};

/**
 * [대상자] 시트 열. 기존에 쓰던 연말정산 검토 시트의 열 이름 그대로.
 * 읽고 쓸 때 1행 제목으로 열 위치를 찾으므로 열 순서가 달라도 된다.
 *  t:  text(기본) | bool(O/체크) | long(여러 줄)
 *  re: 해마다 이름이 바뀌는 열 매칭용 (공백 뺀 제목에 대해)
 *  who: 누가 고칠 수 있나
 *       info   기본 정보 — 1차 담당자·총괄
 *       assign 담당 배정 — 총괄 (비어 있으면 본인을 1차 담당자로 지정 가능)
 *       first  1차 검토 — 1차 담당자·총괄
 *       second 2차 검토 — 2차 검토 담당자·총괄
 */
const COLS = [
  { k: 'no', l: 'No.', g: '기본 정보', who: 'info' },
  { k: 'dept', l: '부서', g: '기본 정보', who: 'info' },
  { k: 'empNo', l: '사원번호', g: '기본 정보', who: 'info' },
  { k: 'name', l: '성명', g: '기본 정보', who: 'info' },
  { k: 'subgroup', l: '사원하위그룹명', g: '기본 정보', who: 'info' },
  { k: 'payArea', l: '급여영역', g: '기본 정보', who: 'info' },
  { k: 'rank', l: '직급', g: '기본 정보', who: 'info' },
  { k: 'phone', l: '전화번호', g: '기본 정보', who: 'info' },
  { k: 'email', l: '이메일주소', g: '기본 정보', who: 'info' },
  { k: 'leave', l: '휴직여부', t: 'bool', g: '기본 정보', who: 'info' },
  { k: 'owner', l: '담당자', g: '담당', who: 'assign' },
  { k: 'owner2', l: '2차검토 담당자', g: '담당', who: 'assign' },
  { k: 'review2', l: '2차 서류검토 여부', t: 'bool', g: '2차 검토', who: 'second' },
  { k: 'ehr', l: 'E-HR 등록', t: 'bool', g: '진행', who: 'first' },
  { k: 'arrived', l: '서류 도착여부', t: 'bool', g: '진행', who: 'first' },
  { k: 'manual', l: '수기서류 제출', t: 'bool', g: '진행', who: 'first' },
  { k: 'verified', l: '서류확인 및 검증', t: 'bool', g: '진행', who: 'first' },
  { k: 'prevWork', l: '종전근무지', g: '종전근무지', who: 'first' },
  { k: 'prevCount', l: '종전근무지 개수', g: '종전근무지', who: 'first' },
  { k: 'rentLoanApply', l: '주택임차차입금(신청여부)', t: 'bool', g: '주택자금 · 월세', who: 'first' },
  { k: 'rentLoan', l: '주택임차차입금', g: '주택자금 · 월세', who: 'first' },
  { k: 'mortApply', l: '장기주택저당차입금(신청여부)', t: 'bool', g: '주택자금 · 월세', who: 'first' },
  { k: 'mort', l: '장기주택저당차입금', g: '주택자금 · 월세', who: 'first' },
  { k: 'mortNts', l: '장기주택 국세청자료 여부', t: 'bool', g: '주택자금 · 월세', who: 'first' },
  { k: 'mortPrev', l: '전년도 장기주택 공제여부(차입일)', re: '^(\\d{2,4}년도?|전년도)장기주택공제여부', g: '주택자금 · 월세', who: 'first' },
  { k: 'savingApply', l: '주택마련저축(신청여부)', t: 'bool', g: '주택자금 · 월세', who: 'first' },
  { k: 'saving', l: '주택마련저축', g: '주택자금 · 월세', who: 'first' },
  { k: 'savingNts', l: '주택마련저축 국세청자료 여부', t: 'bool', g: '주택자금 · 월세', who: 'first' },
  { k: 'rentApply', l: '월세액(신청여부)', t: 'bool', g: '주택자금 · 월세', who: 'first' },
  { k: 'rent', l: '월세액', g: '주택자금 · 월세', who: 'first' },
  { k: 'note1', l: '특이사항', t: 'long', g: '1차 특이·수정사항', who: 'first' },
  { k: 'fix1', l: '수정사항', t: 'long', g: '1차 특이·수정사항', who: 'first' },
  { k: 'missing', l: '미비서류', t: 'long', g: '1차 특이·수정사항', who: 'first' },
  { k: 'note2', l: '특이사항(2차)', t: 'long', g: '2차 검토', who: 'second' },
  { k: 'fix2', l: '수정사항(2차)', t: 'long', g: '2차 검토', who: 'second' },
];
// 시트 끝에 자동으로 붙는 기록용 열
const META_COLS = [
  { k: 'updated', l: '최종수정' },
  { k: 'editor', l: '수정자' },
];

// 시트의 O, Y, ○, 완료 같은 표기를 체크로 인식
const TRUE_RE = /^(o|y|yes|true|1|○|●|◯|v|✓|✔|완료|등록|제출|도착|신청|있음|해당|유|휴직)$/i;
function toBool_(v) {
  return v === true || TRUE_RE.test(String(v == null ? '' : v).trim());
}

function normHeader_(s) {
  return String(s == null ? '' : s).replace(/\s/g, '').toLowerCase();
}

/* ---------- 날짜 ---------- */

function fmt(d, pattern) {
  return Utilities.formatDate(d, TZ, pattern);
}

function ymd(d) {
  return fmt(d, 'yyyy-MM-dd');
}

function cellText_(v) {
  if (v instanceof Date) return ymd(v);
  return String(v == null ? '' : v).trim();
}

/** 'yyyy-mm-dd' 문자열이면 Date로, 아니면 그대로 */
function parseDate_(v) {
  const s = String(v || '').trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return s;
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

/**
 * [담당자] 시트: 순서 | 이름 | 이메일 | 역할(응대담당/2차검토/총괄) | 관리자(Y/N) | 열람범위
 * 열람범위: 비우거나 '전체' = 모든 대상자 / '본인' = 내 담당만 / '본인,이수민,(미배정)' = 내 담당 + 고른 담당자의 대상자
 */
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
      admin: String(r[4]).trim().toUpperCase() === 'Y',
      scope: normScope_(r[5]),
    }))
    .sort((a, b) => a.order - b.order);
}

const SCOPE_ALL = '전체';
const SCOPE_UNASSIGNED = '(미배정)';

/** 열람범위 칸 → '전체' 또는 '본인,이름,…' (본인은 항상 포함) */
function normScope_(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s || s === SCOPE_ALL) return SCOPE_ALL;
  const names = s.split(/[,，\n]/).map(x => x.trim()).filter(x => x && x !== '본인' && x !== SCOPE_ALL);
  return ['본인'].concat(names.filter((x, i) => names.indexOf(x) === i)).join(',');
}

/**
 * 매 요청에 필요한 설정·담당자 목록. 두 시트를 매번 읽지 않도록 60초 캐시한다
 * (웹앱에서 담당자·설정을 저장하면 바로 지워지고, 시트를 직접 고치면 최대 60초 뒤 반영).
 */
function getContext() {
  const base = cached_(CK.ctx, () => {
    const c = getConfig();
    const cfg = {};
    Object.keys(c).forEach(k => { cfg[k] = c[k] instanceof Date ? ymd(c[k]) : c[k]; });
    return { cfg: cfg, members: getMembers() };
  });
  return {
    ss: SpreadsheetApp.getActive(),
    cfg: base.cfg,
    year: String(base.cfg[CFG.YEAR] || defaultYear_()),
    members: base.members,
  };
}

/** 하반기에는 올해 귀속, 상반기(연말정산 진행 중)에는 작년 귀속 */
function defaultYear_() {
  const d = new Date();
  return d.getMonth() >= 6 ? d.getFullYear() : d.getFullYear() - 1;
}

function sheetUrl(ss, sh) {
  return `${ss.getUrl()}#gid=${sh.getSheetId()}`;
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

// ===================================================================
// Setup.gs
// ===================================================================

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

// ===================================================================
// People.gs
// ===================================================================

/**
 * 대상자(서류 검토) / 응대기록
 *
 * [대상자]   기존 검토 시트의 열 그대로 (COLS) + 최종수정 | 수정자
 * [응대기록] 일시 | 사원번호 | 성명 | 작성자 | 구분 | 내용
 */

const LOG_KINDS = ['전화', '메일', '메신저', '방문', '보완요청', '기타'];

function setupPeopleSheets_(ss) {
  let sh = ss.getSheetByName(SHEET.PEOPLE);
  if (!sh) {
    sh = ss.insertSheet(SHEET.PEOPLE);
    const header = COLS.concat(META_COLS).map(c => c.l);
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    styleHeader_(sh.getRange(1, 1, 1, header.length));
    sh.getRange(1, 1, 1, header.length).setWrap(true);
    sh.setRowHeight(1, 42);
    sh.setFrozenRows(1);
    sh.setFrozenColumns(4);
    COLS.forEach((c, i) => {
      if (c.t === 'long') sh.setColumnWidth(i + 1, 220);
      else if (c.t === 'bool') sh.setColumnWidth(i + 1, 80);
    });
  } else {
    // 기존 시트를 그대로 쓴다: 없는 열만 오른쪽 끝에 추가
    const map = colMap_(sh);
    let last = sh.getLastColumn();
    COLS.concat(META_COLS).forEach(c => {
      if (map[c.k]) return;
      last += 1;
      sh.getRange(1, last).setValue(c.l);
      styleHeader_(sh.getRange(1, last));
    });
  }

  if (!ss.getSheetByName(SHEET.LOGS)) {
    const lg = ss.insertSheet(SHEET.LOGS);
    lg.getRange(1, 1, 1, 6).setValues([['일시', '사원번호', '성명', '작성자', '구분', '내용']]);
    styleHeader_(lg.getRange(1, 1, 1, 6));
    lg.getRange('A2:A').setNumberFormat('yyyy-mm-dd hh:mm');
    lg.getRange('F2:F').setWrap(true);
    [130, 90, 80, 80, 80, 480].forEach((w, i) => lg.setColumnWidth(i + 1, w));
    lg.setFrozenRows(1);
  }
}

/** 1행 제목 → { 키: 열번호(1부터) } */
function colMap_(sh) {
  const last = sh.getLastColumn();
  const map = {};
  if (!last) return map;
  const header = sh.getRange(1, 1, 1, last).getValues()[0].map(normHeader_);
  COLS.concat(META_COLS).forEach(c => {
    const want = normHeader_(c.l);
    const re = c.re ? new RegExp(c.re) : null;
    const i = header.findIndex(h => h === want || (re && re.test(h)));
    if (i >= 0) map[c.k] = i + 1;
  });
  return map;
}

function rowToPerson_(r, map, row) {
  const p = { row: row };
  COLS.forEach(c => {
    const v = map[c.k] ? r[map[c.k] - 1] : '';
    p[c.k] = c.t === 'bool' ? toBool_(v) : cellText_(v);
  });
  p.updated = map.updated && r[map.updated - 1] instanceof Date ? fmt(r[map.updated - 1], 'yyyy-MM-dd HH:mm') : cellText_(map.updated ? r[map.updated - 1] : '');
  p.editor = cellText_(map.editor ? r[map.editor - 1] : '');
  // 사원번호가 없으면 행 번호로 구분
  p.id = p.empNo || 'r' + row;
  return p;
}

function readPeople_(ss) {
  const sh = ss.getSheetByName(SHEET.PEOPLE);
  if (!sh || sh.getLastRow() < 2) return [];
  const map = colMap_(sh);
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
    .map((r, i) => rowToPerson_(r, map, i + 2))
    .filter(p => p.name || p.empNo);
}

function readLogs_(ss) {
  const sh = ss.getSheetByName(SHEET.LOGS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues()
    .filter(r => String(r[5]).trim())
    .map(r => ({
      date: r[0] instanceof Date ? fmt(r[0], 'yyyy-MM-dd HH:mm') : String(r[0]),
      empNo: cellText_(r[1]),
      name: String(r[2]),
      author: String(r[3]),
      kind: String(r[4]),
      body: String(r[5]),
    }))
    .reverse();
}

/**
 * id로 대상자 행을 찾아 그 행 값을 함께 돌려준다.
 * 캐시에 있는 행 번호부터 확인하고(읽기 1번), 맞지 않으면 사원번호 열에서 찾는다.
 */
function readPersonRow_(sh, map, id) {
  const width = sh.getLastColumn();
  const hit = (peekCache_(CK.people) || []).find(p => p.id === id);
  if (hit && hit.row >= 2) {
    const values = sh.getRange(hit.row, 1, 1, width).getValues()[0];
    const emp = map.empNo ? cellText_(values[map.empNo - 1]) : '';
    if ((emp || 'r' + hit.row) === id) return { row: hit.row, values: values };
  }
  const row = findPersonRow_(sh, map, id);
  return row < 0 ? null : { row: row, values: sh.getRange(row, 1, 1, width).getValues()[0] };
}

/** id(사원번호 또는 'r행번호')로 행 찾기 */
function findPersonRow_(sh, map, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  if (/^r\d+$/.test(id)) {
    const row = Number(id.slice(1));
    const emp = map.empNo ? cellText_(sh.getRange(row, map.empNo).getValue()) : '';
    return row <= last && !emp ? row : -1;
  }
  if (!map.empNo) return -1;
  const emps = sh.getRange(2, map.empNo, last - 1, 1).getValues().map(([v]) => cellText_(v));
  const i = emps.indexOf(String(id));
  return i < 0 ? -1 : i + 2;
}

/** 이 사람이 이 대상자의 이 칸을 고칠 수 있나 */
function canEdit_(me, cur, col, patch) {
  if (me.isLeader || me.isAdmin) return true;
  if (!me.name) return false;
  // 담당자가 비어 있으면 본인을 1차 담당자로 지정할 수 있다 (같은 저장에서 다른 칸도 함께 고칠 수 있다)
  const takes = !cur.owner && patch.owner === me.name;
  if (col.who === 'assign') return col.k === 'owner' && takes;
  if (col.who === 'second') return cur.owner2 === me.name;
  return cur.owner === me.name || takes;
}

/** 저장할 값. 체크박스 칸(지금 값이 true/false)이면 체크박스로, 아니면 기존 시트처럼 O로 쓴다 */
function cellValue_(col, now, value) {
  if (col.t === 'bool') return typeof now === 'boolean' ? !!value : (value ? 'O' : '');
  return String(value == null ? '' : value).trim();
}

/** 바뀐 칸만 저장한다 (1차·2차 담당자가 동시에 고쳐도 서로 덮어쓰지 않게) */
function apiSavePerson(id, patch) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'review');

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sh = ctx.ss.getSheetByName(SHEET.PEOPLE);
    const map = colMap_(sh);
    const found = readPersonRow_(sh, map, String(id));
    if (!found) throw new Error('대상자를 찾을 수 없습니다. 시트에서 사원번호가 바뀌었을 수 있으니 새로고침해 주세요.');
    const { row, values } = found;
    const cur = rowToPerson_(values, map, row);
    if (!canSee_(me, cur)) throw new Error('열람 범위 밖의 대상자입니다. 관리자에게 열람 범위를 요청하세요.');

    const keys = Object.keys(patch || {}).filter(k => COLS.some(c => c.k === k));
    if (!keys.length) return { personPatch: cur };
    keys.forEach(k => {
      const col = COLS.find(c => c.k === k);
      if (!canEdit_(me, cur, col, patch)) {
        throw new Error(`[${col.l}] 칸은 ${col.who === 'second' ? '2차검토 담당자' : col.who === 'assign' ? '총괄' : '담당자'}만 고칠 수 있습니다.`);
      }
      if (!map[k]) throw new Error(`[대상자] 시트에 "${col.l}" 열이 없습니다. 관리자에게 "기본 시트 만들기"를 다시 실행해 달라고 요청하세요.`);
    });
    if (keys.indexOf('empNo') >= 0) {
      const emp = String(patch.empNo || '').trim();
      if (emp && emp !== cur.empNo && findPersonRow_(sh, map, emp) > 0) throw new Error(`사원번호 ${emp}가 이미 있습니다.`);
    }
    // 바뀐 칸만 쓴다 (수식이 있는 다른 칸은 건드리지 않는다). 쓰기만 이어서 하면 Apps Script가 한 번에 보낸다
    const now = new Date();
    keys.forEach(k => {
      const v = cellValue_(COLS.find(c => c.k === k), values[map[k] - 1], patch[k]);
      sh.getRange(row, map[k]).setValue(v);
      values[map[k] - 1] = v;
    });
    if (map.updated) { sh.getRange(row, map.updated).setValue(now); values[map.updated - 1] = now; }
    if (map.editor) { sh.getRange(row, map.editor).setValue(me.name || me.email); values[map.editor - 1] = me.name || me.email; }

    // 다시 읽지 않고 저장한 값으로 결과를 만들고, 캐시의 그 사람만 바꾼다
    const saved = rowToPerson_(values, map, row);
    patchCache_(CK.people, list => {
      const i = list.findIndex(p => p.id === String(id));
      if (i >= 0) list[i] = saved; else list.push(saved);
    });
    return { personPatch: saved, oldId: String(id) };
  } finally {
    lock.releaseLock();
  }
}

/** 총괄·관리자: 대상자 한 명 추가 */
function apiAddPerson(fields) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'review');
  if (!me.isLeader && !me.isAdmin) throw new Error('대상자 추가는 총괄 또는 관리자만 할 수 있습니다.');
  if (!String(fields.name || '').trim()) throw new Error('성명을 입력해 주세요.');
  apiImportPeople([fields]);
  return refreshPart_('people');
}

/** 총괄·관리자: 대상자 삭제 (행 삭제) */
function apiDeletePerson(id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'review');
  if (!me.isLeader && !me.isAdmin) throw new Error('대상자 삭제는 총괄 또는 관리자만 할 수 있습니다.');
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sh = ctx.ss.getSheetByName(SHEET.PEOPLE);
    const row = findPersonRow_(sh, colMap_(sh), String(id));
    if (row < 0) throw new Error('대상자를 찾을 수 없습니다.');
    sh.deleteRow(row);
  } finally {
    lock.releaseLock();
  }
  return refreshPart_('people');
}

/**
 * 총괄·관리자: 시트에서 복사한 내용 붙여넣기. rows: [{empNo, name, ...}]
 * 사원번호가 같으면 붙여넣은 칸만 갱신, 없으면 추가. 시트 요청은 읽기 1번 + 쓰기 1번.
 */
function apiImportPeople(rows) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'review');
  if (!me.isLeader && !me.isAdmin) throw new Error('붙여넣기는 총괄 또는 관리자만 할 수 있습니다.');

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  let added = 0, updated = 0;
  try {
    const sh = ctx.ss.getSheetByName(SHEET.PEOPLE);
    const map = colMap_(sh);
    const width = sh.getLastColumn();
    const last = sh.getLastRow();
    const data = last >= 2 ? sh.getRange(2, 1, last - 1, width).getValues() : [];
    const byEmp = {};
    data.forEach((r, i) => { const e = map.empNo ? cellText_(r[map.empNo - 1]) : ''; if (e) byEmp[e] = i; });
    const now = new Date();

    (rows || []).forEach(src => {
      if (!String(src.name || '').trim()) return;
      const emp = String(src.empNo || '').trim();
      let r;
      if (emp && byEmp[emp] !== undefined) { r = data[byEmp[emp]]; updated++; }
      else { r = new Array(width).fill(''); data.push(r); if (emp) byEmp[emp] = data.length - 1; added++; }
      COLS.forEach(c => {
        if (!(c.k in src) || !map[c.k]) return;
        const old = r[map[c.k] - 1];
        r[map[c.k] - 1] = c.t === 'bool' ? (typeof old === 'boolean' ? toBool_(src[c.k]) : (toBool_(src[c.k]) ? 'O' : '')) : String(src[c.k] == null ? '' : src[c.k]).trim();
      });
      if (map.updated) r[map.updated - 1] = now;
      if (map.editor) r[map.editor - 1] = me.name || me.email;
    });
    if (data.length) sh.getRange(2, 1, data.length, width).setValues(data);
  } finally {
    lock.releaseLock();
  }
  dropCache_(CK.people);
  return Object.assign(refreshPart_('people'), { added: added, updated: updated });
}

/** 응대기록 추가: 누구든 응대한 사람이 남긴다 */
function apiAddLog(id, kind, body) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'review');
  if (!me.name && !me.isAdmin) throw new Error('[담당자] 시트에 등록된 사람만 응대기록을 남길 수 있습니다.');
  const text = String(body || '').trim();
  if (!text) throw new Error('응대 내용을 입력해 주세요.');
  const sh = ctx.ss.getSheetByName(SHEET.PEOPLE);
  const map = colMap_(sh);
  const found = readPersonRow_(sh, map, String(id));
  if (!found) throw new Error('대상자를 찾을 수 없습니다.');
  const p = rowToPerson_(found.values, map, found.row);
  if (!canSee_(me, p)) throw new Error('열람 범위 밖의 대상자입니다.');
  const now = new Date();
  const k = LOG_KINDS.indexOf(kind) >= 0 ? kind : '기타';
  ctx.ss.getSheetByName(SHEET.LOGS).appendRow([now, p.empNo, p.name, me.name || me.email, k, text]);
  // 응대기록 시트 전체를 다시 읽지 않고 캐시 맨 앞에 넣는다
  const entry = { date: fmt(now, 'yyyy-MM-dd HH:mm'), empNo: p.empNo, name: p.name, author: me.name || me.email, kind: k, body: text };
  patchCache_(CK.logs, list => { list.unshift(entry); });
  return { logAdded: entry };
}

/**
 * 대상자 업로드 양식 (총괄·관리자)
 * 접속한 사람의 드라이브에 "연말정산 대상자 업로드 양식" 스프레드시트를 하나 만들어 두고(다음부터는 다시 씀),
 * 열 제목·담당자 목록을 최신으로 맞춘 뒤 엑셀(.xlsx)로 내려받는 주소를 돌려준다.
 * withPeople=true면 지금 볼 수 있는 대상자 명단을 채워서 준다 (엑셀에서 고쳐 다시 올리기용).
 */
function apiTemplateLink(withPeople) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'review');
  if (!me.isLeader && !me.isAdmin) throw new Error('양식 다운로드는 총괄 또는 관리자만 할 수 있습니다.');

  const props = PropertiesService.getUserProperties();
  let ss = null;
  const id = props.getProperty('TEMPLATE_ID');
  if (id) {
    try {
      const f = DriveApp.getFileById(id);
      if (!f.isTrashed()) ss = SpreadsheetApp.openById(id);
    } catch (e) { ss = null; /* 지워졌거나 권한 없음: 새로 만든다 */ }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('연말정산 대상자 업로드 양식');
    props.setProperty('TEMPLATE_ID', ss.getId());
  }

  const header = COLS.map(c => c.l);
  let sh = ss.getSheetByName('대상자') || ss.getSheets()[0].setName('대상자');
  sh.clear();
  sh.getDataRange().clearDataValidations();
  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setBackground('#efefef').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(1, 42);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(4);
  COLS.forEach((c, i) => sh.setColumnWidth(i + 1, c.t === 'long' ? 220 : c.t === 'bool' ? 90 : 110));

  const ROWS = 1000;
  const boolRule = SpreadsheetApp.newDataValidation().requireValueInList(['O', 'X'], true).setAllowInvalid(true).build();
  COLS.forEach((c, i) => { if (c.t === 'bool') sh.getRange(2, i + 1, ROWS, 1).setDataValidation(boolRule).setHorizontalAlignment('center'); });
  const staff = ctx.members.map(m => m.name);
  if (staff.length) {
    const staffRule = SpreadsheetApp.newDataValidation().requireValueInList(staff, true).setAllowInvalid(true).build();
    ['owner', 'owner2'].forEach(k => sh.getRange(2, COLS.findIndex(c => c.k === k) + 1, ROWS, 1).setDataValidation(staffRule));
  }
  sh.getRange(2, COLS.findIndex(c => c.k === 'empNo') + 1, ROWS, 1).setNumberFormat('@'); // 사원번호 앞자리 0 유지

  let count = 0;
  if (withPeople) {
    const people = forMe_('people', cached_(CK.people, () => readPeople_(ctx.ss)), me);
    const rows = people.map(p => COLS.map(c => c.t === 'bool' ? (p[c.k] ? 'O' : '') : p[c.k]));
    if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
    count = rows.length;
  }

  let guide = ss.getSheetByName('작성 안내') || ss.insertSheet('작성 안내');
  guide.clear();
  const lines = [
    ['연말정산 대상자 업로드 양식 작성 안내'],
    [''],
    ['· [대상자] 시트 2행부터 한 사람씩 적고, 웹앱 서류 검토 > 대상자 업로드에 이 파일을 올리세요.'],
    ['· 성명은 꼭 적어야 합니다. 사원번호가 이미 있으면 그 사람 정보를 갱신하고, 없으면 새로 추가합니다.'],
    ['· 빈 칸은 기존 값을 그대로 둡니다. 필요 없는 열은 비워 두거나 지워도 됩니다.'],
    ['· 체크 항목(휴직여부, E-HR 등록, 서류 도착여부 등)은 O로 적습니다. 이미 체크된 것을 풀려면 X로 적습니다.'],
    [`· 담당자·2차검토 담당자는 [담당자] 목록의 이름과 똑같이 적어야 내 담당으로 연결됩니다: ${staff.join(', ') || '(담당자 없음)'}`],
    ['· 열 순서는 바꿔도 됩니다. 1행 제목(열 이름)으로 맞춥니다.'],
    [''],
    ['예시 (이 시트는 올려도 읽지 않습니다)'],
  ];
  guide.getRange(1, 1, lines.length, 1).setValues(lines);
  guide.getRange('A1').setFontSize(13).setFontWeight('bold');
  guide.getRange(lines.length + 1, 1, 2, 8).setValues([
    ['No.', '부서', '사원번호', '성명', '직급', '담당자', '2차검토 담당자', '서류 도착여부'],
    ['1', '인사팀', '20110321', '홍길동', '과장', staff[0] || '김담당', staff[1] || '이검토', 'O'],
  ]);
  guide.getRange(lines.length + 1, 1, 1, 8).setBackground('#efefef').setFontWeight('bold');
  guide.setColumnWidth(1, 120);
  ss.setActiveSheet(sh);
  SpreadsheetApp.flush();

  const stamp = fmt(new Date(), 'yyyyMMdd');
  return {
    url: `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=xlsx`,
    sheetUrl: ss.getUrl(),
    fileName: `연말정산_대상자_${withPeople ? '명단' : '양식'}_${stamp}.xlsx`,
    count: count,
  };
}

// ===================================================================
// Notice.gs
// ===================================================================

/**
 * 공지사항 / 특이사항 안내
 *
 * [공지]     번호 | 작성일시 | 작성자 | 분류 | 중요도 | 제목 | 내용 | 게시시작 | 게시종료 | 고정(Y) | 삭제(Y) | 첨부 | 수정일시
 * [공지댓글] 번호 | 작성일시 | 작성자 | 내용
 *
 * 공지·응대지침은 총괄·관리자만, 자료·질문은 누구나 올린다. 읽음 표시는 사람별로 저장(사용자 속성).
 */

const NOTICE = {
  SHEET: '공지',
  COMMENTS: '공지댓글',
  CATEGORIES: ['공지', '응대지침', '자료', '질문'],
  LEADER_ONLY: ['공지', '응대지침'],
  LEVELS: ['긴급', '중요', '일반'],
};

const NCOL = { ID: 1, DATE: 2, AUTHOR: 3, CATEGORY: 4, LEVEL: 5, TITLE: 6, BODY: 7, START: 8, END: 9, PINNED: 10, DELETED: 11, FILES: 12, UPDATED: 13 };

function setupNoticeSheets_(ss) {
  if (!ss.getSheetByName(NOTICE.SHEET)) {
    const sh = ss.insertSheet(NOTICE.SHEET);
    const header = ['번호', '작성일시', '작성자', '분류', '중요도', '제목', '내용', '게시시작', '게시종료', '고정(Y)', '삭제(Y)', '첨부', '수정일시'];
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    styleHeader_(sh.getRange(1, 1, 1, header.length));
    sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('H2:I').setNumberFormat('yyyy-mm-dd');
    sh.getRange('M2:M').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('G2:G').setWrap(true);
    [50, 130, 80, 80, 60, 260, 420, 100, 100, 60, 60, 200, 130].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
  }
  if (!ss.getSheetByName(NOTICE.COMMENTS)) {
    const sh = ss.insertSheet(NOTICE.COMMENTS);
    sh.getRange(1, 1, 1, 4).setValues([['번호', '작성일시', '작성자', '내용']]);
    styleHeader_(sh.getRange(1, 1, 1, 4));
    sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('D2:D').setWrap(true);
    [50, 130, 80, 480].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
  }
}

/** 고정 → 중요도 → 최신순 */
function readNotices_(ss) {
  const sh = ss.getSheetByName(NOTICE.SHEET);
  if (!sh || sh.getLastRow() < 2) return [];

  const comments = {};
  const csh = ss.getSheetByName(NOTICE.COMMENTS);
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

  const dt = v => v instanceof Date ? fmt(v, 'yyyy-MM-dd HH:mm') : String(v || '');
  return sh.getRange(2, 1, sh.getLastRow() - 1, NCOL.UPDATED).getValues()
    .filter(r => String(r[0]) && String(r[NCOL.DELETED - 1]).toUpperCase() !== 'Y')
    .map(r => ({
      id: String(r[0]),
      date: dt(r[1]),
      author: String(r[2]),
      category: String(r[3]) || '공지',
      level: NOTICE.LEVELS.indexOf(String(r[4])) >= 0 ? String(r[4]) : '일반',
      title: String(r[5]),
      body: String(r[6]),
      start: cellText_(r[7]),
      end: cellText_(r[8]),
      pinned: String(r[9]).toUpperCase() === 'Y',
      files: filesFromCell_(r[NCOL.FILES - 1]),
      updated: dt(r[NCOL.UPDATED - 1]) || dt(r[1]),
      comments: comments[String(r[0])] || [],
    }))
    .sort((a, b) => (b.pinned - a.pinned) || (NOTICE.LEVELS.indexOf(a.level) - NOTICE.LEVELS.indexOf(b.level)) || b.date.localeCompare(a.date));
}

/** 등록/수정. n: { id?, category, level, title, body, start, end, pinned, files } */
function apiSaveNotice(n) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'notice');
  if (!me.name && !me.isAdmin) throw new Error('[담당자] 시트에 등록된 사람만 글을 쓸 수 있습니다.');
  const lead = me.isLeader || me.isAdmin;

  const title = String(n.title || '').trim();
  const body = String(n.body || '').trim();
  if (!title || !body) throw new Error('제목과 내용을 입력해 주세요.');
  const category = NOTICE.CATEGORIES.indexOf(n.category) >= 0 ? n.category : '질문';
  if (NOTICE.LEADER_ONLY.indexOf(category) >= 0 && !lead) throw new Error(`[${category}]는 총괄 또는 관리자만 올릴 수 있습니다.`);
  const level = lead && NOTICE.LEVELS.indexOf(n.level) >= 0 ? n.level : '일반';
  const start = String(n.start || ''), end = String(n.end || '');
  if (start && end && start > end) throw new Error('게시 종료일이 시작일보다 빠릅니다.');
  const pinned = !!n.pinned && lead;
  const now = new Date();

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  let id;
  try {
    const sh = ctx.ss.getSheetByName(NOTICE.SHEET);
    const last = sh.getLastRow();
    const rows = last >= 2 ? sh.getRange(2, 1, last - 1, 3).getValues() : [];
    id = String(n.id || '');
    const idx = id ? rows.findIndex(r => String(r[0]) === id) : -1;
    const values = [category, level, title, body, parseDate_(start), parseDate_(end), pinned ? 'Y' : ''];
    if (idx >= 0) {
      if (String(rows[idx][2]) !== me.name && !lead) throw new Error('본인 글만 고칠 수 있습니다.');
      sh.getRange(idx + 2, NCOL.CATEGORY, 1, values.length).setValues([values]);
      sh.getRange(idx + 2, NCOL.FILES, 1, 2).setValues([[filesToCell_(n.files), now]]);
    } else {
      id = String((rows.length ? Math.max.apply(null, rows.map(r => Number(r[0]) || 0)) : 0) + 1);
      sh.getRange(last + 1, 1, 1, NCOL.UPDATED).setValues([[id, now, me.name || me.email].concat(values, ['', filesToCell_(n.files), now])]);
    }
  } finally {
    lock.releaseLock();
  }
  registerFiles_('notice', id, title, n.files, me);
  markRead_(id);
  return Object.assign(refreshPart_('notices'), { savedId: id, reads: getReads_() });
}

function apiDeleteNotice(id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'notice');
  const sh = ctx.ss.getSheetByName(NOTICE.SHEET);
  const rows = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 3).getValues();
  const idx = rows.findIndex(r => String(r[0]) === String(id));
  if (idx < 0) throw new Error('글을 찾을 수 없습니다.');
  if (rows[idx][2] !== me.name && !me.isLeader && !me.isAdmin) throw new Error('본인 글만 삭제할 수 있습니다.');
  sh.getRange(idx + 2, NCOL.DELETED).setValue('Y');
  return refreshPart_('notices');
}

function apiAddNoticeComment(id, body) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'notice');
  if (!me.name && !me.isAdmin) throw new Error('[담당자] 시트에 등록된 사람만 댓글을 쓸 수 있습니다.');
  const text = String(body || '').trim();
  if (!text) throw new Error('댓글 내용을 입력해 주세요.');
  const now = new Date();
  ctx.ss.getSheetByName(NOTICE.COMMENTS).appendRow([String(id), now, me.name || me.email, text]);
  const comment = { date: fmt(now, 'yyyy-MM-dd HH:mm'), author: me.name || me.email, body: text };
  patchCache_(CK.notices, list => { const n = list.find(x => x.id === String(id)); if (n) n.comments.push(comment); });
  return { commentAdded: { id: String(id), comment: comment } };
}

/* ---------- 읽음 표시 (사람마다 따로: 웹앱이 "접속한 사용자" 권한으로 실행되므로 사용자 속성에 저장) ---------- */

function getReads_() {
  try {
    return JSON.parse(PropertiesService.getUserProperties().getProperty('READ_NOTICES') || '[]');
  } catch (e) {
    return [];
  }
}

function markRead_(id) {
  const list = getReads_().filter(x => x !== String(id));
  list.push(String(id));
  PropertiesService.getUserProperties().setProperty('READ_NOTICES', JSON.stringify(list.slice(-500)));
  return list;
}

function apiMarkRead(id) {
  return { reads: markRead_(id) };
}

function apiMarkUnread(id) {
  const list = getReads_().filter(x => x !== String(id));
  PropertiesService.getUserProperties().setProperty('READ_NOTICES', JSON.stringify(list));
  return { reads: list };
}

// ===================================================================
// Rules.gs
// ===================================================================

/**
 * 연말정산 기준 (공제 요건·제출서류·응대 기준 모음)
 *
 * [기준]          번호 | 분류 | 제목 | 내용 | 적용 귀속연도 | 담당 | 원문 링크 | 최종수정 | 수정자 | 삭제(Y) | 첨부
 * [기준개정이력]  번호 | 일시 | 수정자 | 개정 내용
 *
 * 누구나 조회, 총괄·관리자만 등록/수정/삭제 (시트 보호로도 막는다).
 */

const RULES = {
  SHEET: '기준',
  HISTORY: '기준개정이력',
  CATEGORIES: ['인적공제', '주택자금·월세', '의료비', '교육비', '기부금', '신용카드 등', '연금·보험', '제출서류', '응대 기준', '기타'],
};

const RCOL = { ID: 1, CATEGORY: 2, TITLE: 3, BODY: 4, YEAR: 5, OWNER: 6, LINK: 7, UPDATED: 8, EDITOR: 9, DELETED: 10, FILES: 11 };

// 처음 설치할 때 넣는 예시 (세법은 해마다 바뀌므로 국세청 안내로 확인 후 고쳐 쓰도록 표시)
const SAMPLE_RULES = [
  ['인적공제', '기본공제 대상 부양가족 요건', '(예시 · 국세청 안내로 최신 기준 확인 후 사용)\n· 1인당 150만원 소득공제\n· 연간 소득금액 100만원 이하 (근로소득만 있으면 총급여 500만원 이하)\n· 나이: 직계존속 만 60세 이상, 직계비속 만 20세 이하, 형제자매 만 20세 이하 또는 60세 이상\n· 장애인은 나이 제한 없음 (소득 요건은 적용)\n· 같은 사람을 두 명이 공제할 수 없음 → 맞벌이 부부·형제 중복 여부 꼭 확인'],
  ['주택자금·월세', '월세액 세액공제', '(예시 · 국세청 안내로 최신 기준 확인 후 사용)\n· 무주택 세대의 세대주, 총급여 8천만원 이하\n· 국민주택규모(85㎡) 이하 또는 기준시가 4억원 이하 주택\n· 연 1,000만원 한도, 총급여 5,500만원 이하 17% / 그 외 15%\n· 제출: 임대차계약서 사본, 월세 이체 내역 (주민등록 주소지와 계약서 주소 일치 확인)'],
  ['주택자금·월세', '주택임차차입금 원리금 상환액', '(예시 · 국세청 안내로 최신 기준 확인 후 사용)\n· 무주택 세대의 세대주, 국민주택규모 주택 임차\n· 상환액의 40% 소득공제, 주택마련저축 납입액과 합산 연 400만원 한도\n· 제출: 주택자금상환증명서(금융기관), 임대차계약서, 주민등록등본'],
  ['의료비', '의료비 세액공제', '(예시 · 국세청 안내로 최신 기준 확인 후 사용)\n· 총급여 3% 초과분의 15% (난임시술비 30%, 미숙아·선천성이상아 20%)\n· 본인·65세 이상·장애인·6세 이하: 한도 없음 / 그 외 부양가족: 연 700만원\n· 실손보험금 받은 금액은 빼야 함\n· 간소화 자료에 없는 의료비(안경 등)는 영수증 별도 제출'],
  ['제출서류', '공통 제출 서류', '(예시 · 회사 기준에 맞게 고쳐 쓰세요)\n· 소득·세액공제신고서\n· 국세청 간소화 자료 PDF (부양가족 자료 제공 동의 포함)\n· 부양가족 변동 시 가족관계증명서\n· 중도입사자: 종전근무지 원천징수영수증·근로소득지급명세서'],
  ['응대 기준', '증빙 서류 받는 방법', '(예시 · 회사 기준에 맞게 고쳐 쓰세요)\n· 증빙은 지정된 경로(웹앱 첨부 또는 회사 메일)로만 받습니다. 개인 메신저 금지\n· 전화 문의는 본인 확인 후 안내\n· 안내한 내용은 응대기록에 남겨 다른 담당자도 볼 수 있게 합니다'],
];

function setupRuleSheets_(ss) {
  if (!ss.getSheetByName(RULES.SHEET)) {
    const sh = ss.insertSheet(RULES.SHEET);
    const header = ['번호', '분류', '제목', '내용', '적용 귀속연도', '담당', '원문 링크', '최종수정', '수정자', '삭제(Y)', '첨부'];
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    styleHeader_(sh.getRange(1, 1, 1, header.length));
    sh.getRange('B2:B').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(RULES.CATEGORIES).build());
    sh.getRange('H2:H').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('D2:D').setWrap(true);
    [50, 100, 220, 480, 90, 80, 200, 130, 80, 70, 200].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
    const year = String(getConfig()[CFG.YEAR] || defaultYear_());
    const now = new Date();
    const rows = SAMPLE_RULES.map((r, i) => [i + 1, r[0], r[1], r[2], year, '', '', now, '설치 예시', '', '']);
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
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
      year: cellText_(r[4]),
      owner: String(r[5] || ''),
      link: String(r[6] || ''),
      updated: r[7] instanceof Date ? fmt(r[7], 'yyyy-MM-dd') : String(r[7] || '').slice(0, 10),
      editor: String(r[8] || ''),
      files: filesFromCell_(r[RCOL.FILES - 1]),
      history: (history[String(r[0])] || []).reverse(),
    }));
}

/** 등록/수정. rule: { id?, category, title, body, year, owner, link, changeNote, files } */
function apiSaveRule(rule) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'rules');
  if (!me.isLeader && !me.isAdmin) throw new Error('연말정산 기준은 총괄 또는 관리자만 등록·수정할 수 있습니다.');

  const title = String(rule.title || '').trim();
  const body = String(rule.body || '').trim();
  if (!title || !body) throw new Error('제목과 내용을 입력해 주세요.');
  const category = RULES.CATEGORIES.indexOf(rule.category) >= 0 ? rule.category : '기타';
  const editor = me.name || me.email;
  const now = new Date();

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  let id;
  try {
    const sh = ctx.ss.getSheetByName(RULES.SHEET);
    const hsh = ctx.ss.getSheetByName(RULES.HISTORY);
    const last = sh.getLastRow();
    const ids = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(([v]) => String(v)) : [];
    const values = [category, title, body, String(rule.year || '').trim(), String(rule.owner || '').trim(), String(rule.link || '').trim(), now, editor];

    id = String(rule.id || '');
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
  registerFiles_('rules', id, title, rule.files, me);
  return Object.assign(refreshPart_('rules'), { savedId: id });
}

function apiDeleteRule(id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'rules');
  if (!me.isLeader && !me.isAdmin) throw new Error('총괄 또는 관리자만 삭제할 수 있습니다.');
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
 * 첨부 파일 (구글 드라이브)
 *
 * [설정] > "첨부 폴더 ID" 의 드라이브 폴더 아래에 구분별 하위 폴더를 만들어 저장한다.
 * 증빙 서류에는 개인정보가 들어 있으므로 이 폴더는 담당자에게만 공유한다 (관리 > 권한 한 번에 적용).
 *
 * [첨부] 시트에 모든 첨부를 한 줄씩 기록해 [첨부파일] 메뉴에서 한곳에 모아 본다.
 *   파일ID | 올린일시 | 올린사람 | 구분 | 대상번호 | 대상 | 파일명 | URL | 크기 | 형식 | 메모 | 삭제(Y)
 * 공지·기준 글에는 첨부 목록을 JSON 문자열로도 저장한다: [{id, name, url, size, mime}]
 */

const FILES = {
  SHEET: '첨부',
  MAX_BYTES: 10 * 1024 * 1024,
  AREAS: { person: '대상자 증빙', notice: '공지사항', rules: '연말정산 기준', etc: '일반 자료' },
  MENU: { person: 'review', notice: 'notice', rules: 'rules', etc: 'files' },
};

const FCOL = { ID: 1, DATE: 2, UPLOADER: 3, AREA: 4, REF: 5, REF_LABEL: 6, NAME: 7, URL: 8, SIZE: 9, MIME: 10, MEMO: 11, DELETED: 12 };

function setupFileSheet_(ss) {
  if (ss.getSheetByName(FILES.SHEET)) return;
  const sh = ss.insertSheet(FILES.SHEET);
  const header = ['파일ID', '올린일시', '올린사람', '구분', '대상번호', '대상', '파일명', 'URL', '크기', '형식', '메모', '삭제(Y)'];
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  styleHeader_(sh.getRange(1, 1, 1, header.length));
  sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
  [120, 130, 80, 100, 80, 160, 260, 220, 80, 120, 200, 60].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.hideColumns(1);
  sh.setFrozenRows(1);
}

/**
 * file: { name, mimeType, data(base64), area, ref?, refLabel?, memo? } → 첨부 정보
 * 대상자 증빙과 일반 자료는 올리자마자 [첨부]에 기록하고, 공지·기준은 글을 저장할 때 기록한다.
 */
function apiUploadFile(file) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  if (!me.name && !me.isAdmin) throw new Error('[담당자] 시트에 등록된 사람만 파일을 올릴 수 있습니다.');
  const area = FILES.AREAS[file.area] ? file.area : 'etc';
  requireMenu_(me, FILES.MENU[area]);

  if (area === 'person') {
    const p = cached_(CK.people, () => readPeople_(ctx.ss)).find(x => (x.empNo || x.id) === String(file.ref || ''));
    if (!p || !canSee_(me, p)) throw new Error('열람 범위 밖의 대상자에게는 파일을 올릴 수 없습니다.');
  }
  const name = String(file.name || '').trim() || '첨부파일';
  const bytes = Utilities.base64Decode(String(file.data || ''));
  if (!bytes.length) throw new Error('빈 파일은 올릴 수 없습니다.');
  if (bytes.length > FILES.MAX_BYTES) throw new Error('10MB 이하 파일만 올릴 수 있습니다. 큰 파일은 드라이브에 직접 올리고 링크를 붙여주세요.');

  const folder = areaFolder_(FILES.AREAS[area], ctx);
  const created = folder.createFile(Utilities.newBlob(bytes, file.mimeType || 'application/octet-stream', name));
  created.setDescription(`올린 사람: ${me.name || me.email}${file.refLabel ? ` / 대상: ${file.refLabel}` : ''}`);
  const info = { id: created.getId(), name: created.getName(), url: created.getUrl(), size: bytes.length, mime: created.getMimeType() };

  if (area === 'person' || area === 'etc') {
    registerFiles_(area, String(file.ref || ''), String(file.refLabel || ''), [info], me, String(file.memo || ''));
    return Object.assign({ file: info }, refreshPart_('files'));
  }
  return { file: info };
}

function areaFolder_(areaName, ctx) {
  const id = String((ctx ? ctx.cfg : getConfig())[CFG.DRIVE_FOLDER] || '').trim();
  if (!id) throw new Error('첨부 폴더가 설정되지 않았습니다. 관리자에게 관리 > 설정 > "첨부 폴더" 입력을 요청하세요.');
  let root;
  try {
    root = DriveApp.getFolderById(id);
  } catch (e) {
    throw new Error('첨부 폴더에 접근할 수 없습니다. 폴더가 "편집자"로 공유되어 있는지 관리자에게 확인하세요.');
  }
  const it = root.getFoldersByName(areaName);
  return it.hasNext() ? it.next() : root.createFolder(areaName);
}

/** [첨부] 시트에 아직 없는 파일만 기록 */
function registerFiles_(area, ref, refLabel, files, me, memo) {
  const list = (files || []).filter(f => f && f.id);
  if (!list.length) return;
  const sh = SpreadsheetApp.getActive().getSheetByName(FILES.SHEET);
  const last = sh.getLastRow();
  const have = new Set(last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(([v]) => String(v)) : []);
  const now = new Date();
  const rows = list.filter(f => !have.has(String(f.id))).map(f => [
    String(f.id), now, me.name || me.email, FILES.AREAS[area], ref, refLabel, String(f.name), String(f.url || ''), Number(f.size) || 0, String(f.mime || ''), memo || '', '',
  ]);
  if (rows.length) sh.getRange(last + 1, 1, rows.length, rows[0].length).setValues(rows);
  dropCache_(CK.files);
}

function readFiles_(ss) {
  const sh = ss.getSheetByName(FILES.SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const areaKey = {};
  Object.keys(FILES.AREAS).forEach(k => { areaKey[FILES.AREAS[k]] = k; });
  return sh.getRange(2, 1, sh.getLastRow() - 1, FCOL.DELETED).getValues()
    .filter(r => String(r[0]) && String(r[FCOL.DELETED - 1]).toUpperCase() !== 'Y')
    .map(r => ({
      id: String(r[0]),
      date: r[1] instanceof Date ? fmt(r[1], 'yyyy-MM-dd HH:mm') : String(r[1]),
      uploader: String(r[2]),
      area: areaKey[String(r[3])] || 'etc',
      ref: cellText_(r[4]),
      refLabel: String(r[5]),
      name: String(r[6]),
      url: String(r[7]),
      size: Number(r[8]) || 0,
      mime: String(r[9]),
      memo: String(r[10]),
    }))
    .reverse();
}

/** 첨부 삭제: 올린 사람 또는 총괄·관리자. 드라이브 파일은 휴지통으로 (30일 안에 복구 가능) */
function apiDeleteFile(id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'files');
  const sh = ctx.ss.getSheetByName(FILES.SHEET);
  const rows = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 3).getValues();
  const idx = rows.findIndex(r => String(r[0]) === String(id));
  if (idx < 0) throw new Error('파일을 찾을 수 없습니다.');
  if (String(rows[idx][2]) !== me.name && !me.isLeader && !me.isAdmin) throw new Error('본인이 올린 파일만 삭제할 수 있습니다.');
  sh.getRange(idx + 2, FCOL.DELETED).setValue('Y');
  try { DriveApp.getFileById(String(id)).setTrashed(true); } catch (e) { /* 이미 지워졌거나 권한 없음: 목록에서만 뺀다 */ }
  return refreshPart_('files');
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

// ===================================================================
// Admin.gs
// ===================================================================

/**
 * 관리 메뉴 (관리자 전용): 담당자 관리, 권한 한 번에 적용, 메뉴 권한, 설정
 *
 * 관리자 = 스프레드시트 소유자 + [담당자] 시트에서 관리자(Y)인 사람
 *
 * 권한 적용이 하는 일
 *  1. 스프레드시트를 담당자에게 편집자로 공유
 *  2. 첨부 폴더를 편집자로 공유
 *  3. 시트 보호 (설정·담당자는 관리자만, 기준은 총괄·관리자만)
 */

const MENUS = [
  { key: 'dash', label: '대시보드' },
  { key: 'review', label: '서류 검토' },
  { key: 'notice', label: '공지사항' },
  { key: 'rules', label: '연말정산 기준' },
  { key: 'files', label: '첨부파일' },
];
const ROLES = ['응대담당', '2차검토', '총괄'];
const LEADER_ROLE = '총괄';
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

/** { 응대담당: {dash:true, ...}, ... } */
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

function adminData_(ctx, useCache) {
  const cfg = ctx.cfg;
  return {
    members: ctx.members.map(m => ({ order: m.order, name: m.name, email: m.email, role: m.role, admin: m.admin, scope: m.scope })),
    // 열람 범위 화면용: 담당자 칸에 적힌 이름별 대상자 수와, 사람별 [1차, 2차] 담당 (이름만, 개인정보 없음)
    owners: ownerCounts_(ctx),
    assign: readPeople_(ctx.ss).map(p => [p.owner, p.owner2]),
    menus: MENUS,
    roles: ROLES,
    menuAccess: getMenuAccess_(),
    settings: {
      team: String(cfg[CFG.TEAM] || ''),
      year: ctx.year,
      folderId: String(cfg[CFG.DRIVE_FOLDER] || ''),
    },
    webAppUrl: ScriptApp.getService().getUrl() || '',
    status: useCache ? cached_('access', () => accessStatus_(ctx), 300) : freshStatus_(ctx),
  };
}

function ownerCounts_(ctx) {
  const counts = {};
  readPeople_(ctx.ss).forEach(p => { const k = p.owner || SCOPE_UNASSIGNED; counts[k] = (counts[k] || 0) + 1; });
  return counts;
}

function freshStatus_(ctx) {
  const status = accessStatus_(ctx);
  putCache_('access', status, 300);
  return status;
}

/**
 * 담당자별 접근 상태
 * sheet/folder: 'editor' | 'viewer' | 'none' | 'owner' ('n/a' = 폴더 설정 필요)
 */
function accessStatus_(ctx) {
  const lower = list => list.map(u => u.getEmail().toLowerCase());
  const res = { sheet: { ok: true, name: ctx.ss.getName() }, folder: { ok: false } };

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
    res.folder = { ok: false, error: '첨부 폴더가 비어 있습니다.' };
  }

  const level = (email, owner, editors, viewers) =>
    email === owner ? 'owner' : editors.indexOf(email) >= 0 ? 'editor' : viewers.indexOf(email) >= 0 ? 'viewer' : 'none';

  const people = {};
  ctx.members.filter(m => m.email).forEach(m => {
    const e = m.email.toLowerCase();
    people[m.name] = {
      sheet: level(e, fileOwner, fileEditors, fileViewers),
      folder: res.folder.ok ? level(e, folderOwner, folderEditors, folderViewers) : 'n/a',
    };
  });
  return { resources: res, people: people };
}

/* ---------- 저장 ---------- */

/** 담당자 목록 전체 저장. list: [{order, name, email, role, admin, scope}] */
function apiSaveMembers(list) {
  const ctx = getContext();
  requireAdmin_(ctx);

  const rows = (list || []).map((m, i) => ({
    order: Number(m.order) || i + 1,
    name: String(m.name || '').trim(),
    email: String(m.email || '').trim().toLowerCase(),
    role: ROLES.indexOf(m.role) >= 0 ? m.role : ROLES[0],
    admin: !!m.admin,
    scope: normScope_(m.scope),
  })).filter(m => m.name);

  const names = {};
  rows.forEach(m => {
    if (names[m.name]) throw new Error(`이름이 겹칩니다: ${m.name}. 동명이인은 "김민지A"처럼 구분해 주세요. ([대상자] 시트의 담당자 칸도 같은 이름으로 적어야 합니다)`);
    names[m.name] = true;
    if (m.email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(m.email)) throw new Error(`이메일 형식을 확인해 주세요: ${m.name} (${m.email})`);
  });

  const sh = ctx.ss.getSheetByName(SHEET.MEMBERS);
  ensureHeader_(sh, 6, '열람범위');
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 6).setValues(rows.map(m => [m.order, m.name, m.email, m.role, m.admin ? 'Y' : 'N', m.scope]));
  }
  clearCaches_();
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
  dropCache_('menus');
  return { menuAccess: getMenuAccess_() };
}

function apiSaveSettings(s) {
  const ctx = getContext();
  requireAdmin_(ctx);
  const year = String(s.year || '').trim();
  if (!/^\d{4}$/.test(year)) throw new Error('귀속연도는 2026처럼 네 자리로 입력해 주세요.');
  setConfigValue(CFG.TEAM, String(s.team || '').trim());
  setConfigValue(CFG.YEAR, Number(year));
  setConfigValue(CFG.DRIVE_FOLDER, String(s.folderId || '').trim().replace(/^.*\/folders\//, '').replace(/[?#].*$/, ''));
  clearCaches_();
  return adminData_(getContext(), false);
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
    warn('첨부 폴더가 비어 있어 폴더 공유를 건너뛰었습니다. (관리 > 설정)');
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

  // 3. 시트 보호
  try {
    applyProtections_(ctx);
    ok('시트 보호를 적용했습니다. (설정·담당자: 관리자만 / 기준: 총괄·관리자만)');
  } catch (err) {
    warn(`시트 보호 적용 실패 (${err.message})`);
  }

  clearCaches_();
  if (!log.some(l => l.ok && /권한을 줬습니다/.test(l.text))) ok('새로 줄 공유 권한은 없었습니다. 모두 이미 공유되어 있습니다.');
  return { log: log, data: adminData_(getContext(), false) };
}

/** 담당자를 뺄 때 시트·폴더 접근 해제 (소유자와 본인은 건드리지 않는다) */
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
  return { log: log, data: adminData_(getContext(), false) };
}

// ===================================================================
// WebApp.gs
// ===================================================================

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
