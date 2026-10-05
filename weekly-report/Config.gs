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
      clearContextCache_();
      return;
    }
  }
  sh.appendRow([key, value]);
  clearContextCache_();
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
    const hit = CacheService.getDocumentCache().get(CTX_KEY);
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
  const v = sh.getRange(1, 1, INPUT.FIRST_ROW + INPUT.ROWS - 1, INPUT.PRE_NEXT_COL + 1).getValues();
  const at = a1 => v[Number(a1.slice(1)) - 1][a1.charCodeAt(0) - 65];
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
