/**
 * 공통 설정 / 유틸
 */

const TZ = 'Asia/Seoul';
const DAY_KO = ['일', '월', '화', '수', '목', '금', '토'];

const SHEET = {
  CONFIG: '설정',
  MEMBERS: '담당자',
  // 아래 셋은 귀속연도별 시트: 대상자_2026, 응대기록_2026, 일정_2026 …
  PEOPLE: '대상자',
  LOGS: '응대기록',
  EVENTS: '일정',
};
const YEAR_SHEETS = [SHEET.PEOPLE, SHEET.LOGS, SHEET.EVENTS];

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
  // 퇴사자: 체크하면 명단에서 흐리게 보이고 대시보드 진행률에서 빠진다
  { k: 'retired', l: '퇴사여부', t: 'bool', re: '^(퇴사|퇴직)(여부)?$', g: '기본 정보', who: 'info' },
  { k: 'owner', l: '담당자', g: '담당', who: 'assign' },
  { k: 'owner2', l: '2차검토 담당자', g: '담당', who: 'assign' },
  { k: 'review2', l: '2차 서류검토 여부', t: 'bool', g: '2차 검토', who: 'second' },
  // 예전 시트의 "E-HR 등록" 열도 같은 칸으로 읽는다
  { k: 'ehr', l: '시스템 등록', t: 'bool', re: '^(시스템등록|e-?hr등록)$', g: '진행', who: 'first' },
  { k: 'arrived', l: '서류 도착여부', t: 'bool', g: '진행', who: 'first' },
  { k: 'verified', l: '서류확인 및 검증', t: 'bool', g: '진행', who: 'first' },
  // 예전 시트의 "종전근무지"(회사명) 열도 읽는다: 회사명이 적혀 있으면 있음(체크)으로 본다
  { k: 'prevWork', l: '종전근무지 여부', t: 'bool', loose: true, re: '^종전근무지(여부)?$', g: '진행', who: 'first' },
  { k: 'rentLoanApply', l: '주택임차차입금(신청여부)', t: 'bool', g: '주택자금 · 월세', who: 'first' },
  { k: 'mortApply', l: '장기주택저당차입금(신청여부)', t: 'bool', g: '주택자금 · 월세', who: 'first' },
  { k: 'mortNts', l: '장기주택 국세청자료 여부', t: 'bool', g: '주택자금 · 월세', who: 'first' },
  { k: 'savingApply', l: '주택마련저축(신청여부)', t: 'bool', g: '주택자금 · 월세', who: 'first' },
  { k: 'savingNts', l: '주택마련저축 국세청자료 여부', t: 'bool', g: '주택자금 · 월세', who: 'first' },
  { k: 'rentApply', l: '월세액(신청여부)', t: 'bool', g: '주택자금 · 월세', who: 'first' },
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
const TRUE_RE = /^(o|y|yes|true|1|○|●|◯|v|✓|✔|완료|등록|제출|도착|신청|있음|해당|유|휴직|퇴사|퇴직)$/i;
function toBool_(v) {
  return v === true || TRUE_RE.test(String(v == null ? '' : v).trim());
}

// loose 체크 칸(종전근무지 여부): 회사명처럼 무엇이든 적혀 있으면 체크, X·N·없음·빈 칸은 체크 안 함
const FALSE_RE = /^(x|n|no|false|0|없음|무|해당없음|-)$/i;
function boolOf_(col, v) {
  if (!col.loose) return toBool_(v);
  if (v === true || v === false) return v;
  const s = String(v == null ? '' : v).trim();
  return !!s && !FALSE_RE.test(s);
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

/**
 * 시트 전체를 한 번에 읽는다 (getDataRange 1번 = 행 수·열 수 조회 없이 읽기 1번).
 * 1행(제목)을 뺀 행들을 width 열 길이로 맞춰 돌려준다. 시트가 없거나 비었으면 [].
 */
function dataRows_(sh, width) {
  if (!sh) return [];
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  return v.slice(1).map(r => {
    if (!width) return r;
    const x = r.length > width ? r.slice(0, width) : r.slice();
    while (x.length < width) x.push('');
    return x;
  });
}

/** 기본 연도: 캐시된 설정이 있으면 그 값을 쓰고, 없을 때만 [설정] 시트를 읽는다 */
function baseYear_() {
  const base = peekCache_(CK.ctx);
  return String((base && base.cfg[CFG.YEAR]) || getConfig()[CFG.YEAR] || defaultYear_());
}

function getConfig() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.CONFIG);
  const map = {};
  if (!sh) return map;
  dataRows_(sh, 2).forEach(([k, v]) => {
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
 * 열람범위: 비우거나 '전체' = 모든 대상자(설정 안 함) / '본인' = 내 담당만 / '본인,이수민' = 내 담당 + 고른 담당자의 대상자
 * 담당자가 비어 있는 대상자는 열람범위와 상관없이 모두에게 보인다.
 */
function getMembers() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.MEMBERS);
  return dataRows_(sh, 6)
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
    return { cfg: cfg, members: getMembers(), years: findYears_() };
  });
  const baseYear = String(base.cfg[CFG.YEAR] || defaultYear_());
  return {
    ss: SpreadsheetApp.getActive(),
    cfg: base.cfg,
    baseYear: baseYear,                     // 관리자가 정한 기본(진행 중) 연도
    years: yearsWith_(base.years, baseYear), // 관리하는 연도 목록 (최근 연도 먼저)
    year: viewYear_(),                      // 지금 접속자가 보고 있는 연도
    members: base.members,
  };
}

/* ---------- 귀속연도 ---------- */

/** 시트 이름(대상자_2026 …)에서 연도 목록을 찾는다 */
function findYears_() {
  const set = {};
  SpreadsheetApp.getActive().getSheets().forEach(sh => {
    const m = /^대상자_(\d{4})$/.exec(sh.getName());
    if (m) set[m[1]] = true;
  });
  return Object.keys(set);
}

function yearsWith_(years, baseYear) {
  const all = years.indexOf(baseYear) >= 0 ? years.slice() : years.concat(baseYear);
  return all.sort((a, b) => Number(b) - Number(a));
}

/**
 * 접속자가 보고 있는 연도. 사람마다 사용자 속성(VIEW_YEAR)에 저장하고, 없으면 기본 연도.
 * 한 번 실행하는 동안은 같은 값을 쓴다 (접속자가 바뀌면 다시 읽는다).
 */
let VIEW_MEMO_ = null;
function viewYear_() {
  const email = Session.getActiveUser().getEmail();
  if (VIEW_MEMO_ && VIEW_MEMO_.email === email) return VIEW_MEMO_.year;
  const base = peekCache_(CK.ctx);
  const baseYear = baseYear_();
  const years = yearsWith_(base ? base.years : findYears_(), baseYear);
  let y = String(PropertiesService.getUserProperties().getProperty('VIEW_YEAR') || baseYear);
  if (years.indexOf(y) < 0) y = baseYear;
  VIEW_MEMO_ = { email: email, year: y };
  return y;
}

/**
 * 연도별 시트. 예전 설치본의 "대상자"·"응대기록"(연도 없는 이름)은 기본 연도 시트로 보고 이름을 바꿔 쓴다.
 */
function yearSheet_(ss, base, year) {
  const y = String(year || viewYear_());
  const sh = ss.getSheetByName(`${base}_${y}`);
  if (sh) return sh;
  const legacy = ss.getSheetByName(base);
  if (legacy && y === baseYear_()) {
    try { legacy.setName(`${base}_${y}`); } catch (e) { /* 이름을 못 바꾸면 그대로 쓴다 */ }
    return legacy;
  }
  return null;
}

/** 예전 이름의 연도별 시트를 지정한 연도 이름으로 바꾼다 (기본 연도를 바꾸기 전에 호출) */
function migrateLegacySheets_(ss, year) {
  YEAR_SHEETS.forEach(base => {
    const legacy = ss.getSheetByName(base);
    if (legacy && !ss.getSheetByName(`${base}_${year}`)) legacy.setName(`${base}_${year}`);
  });
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
