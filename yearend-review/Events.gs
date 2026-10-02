/**
 * 일정 (대시보드 캘린더). 구글 캘린더와 연결하지 않고 연도별 시트에 저장한다.
 *
 * [일정_2026]  번호 | 시작일 | 종료일 | 제목 | 구분 | 메모 | 작성자 | 삭제(Y)
 *
 * 누구나 조회, 총괄·관리자만 등록/수정/삭제.
 */

const EVENT_KINDS = ['마감', '일정', '교육', '휴무', '기타'];
const ECOL = { ID: 1, START: 2, END: 3, TITLE: 4, KIND: 5, MEMO: 6, AUTHOR: 7, DELETED: 8 };

function setupEventSheet_(ss, year) {
  if (yearSheet_(ss, SHEET.EVENTS, year)) return;
  const sh = ss.insertSheet(`${SHEET.EVENTS}_${year}`);
  const header = ['번호', '시작일', '종료일', '제목', '구분', '메모', '작성자', '삭제(Y)'];
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  styleHeader_(sh.getRange(1, 1, 1, header.length));
  sh.getRange('B2:C').setNumberFormat('yyyy-mm-dd');
  sh.getRange('E2:E').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(EVENT_KINDS).build());
  [50, 100, 100, 260, 70, 320, 80, 60].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);
  // 연말정산 기본 일정 예시 (회사 일정에 맞게 고쳐 쓰도록 표시). 지급명세서 제출 기한은 다음 해 3월 10일
  const next = Number(year) + 1;
  const d = s => parseDate_(s);
  const rows = [
    [1, d(`${next}-01-15`), '', '간소화 자료 조회 시작 (예시)', '일정', '국세청 연말정산 간소화 서비스 개시일을 확인해 고쳐 주세요', '설치 예시', ''],
    [2, d(`${next}-01-31`), '', '직원 서류 제출 마감 (예시)', '마감', '', '설치 예시', ''],
    [3, d(`${next}-02-01`), d(`${next}-02-10`), '1차 검토·보완 요청 (예시)', '일정', '', '설치 예시', ''],
    [4, d(`${next}-03-10`), '', '근로소득 지급명세서 제출 기한', '마감', '', '설치 예시', ''],
  ];
  sh.getRange(2, 1, rows.length, header.length).setValues(rows);
}

function readEvents_(ss) {
  const sh = yearSheet_(ss, SHEET.EVENTS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, ECOL.DELETED).getValues()
    .filter(r => String(r[0]) && cellText_(r[1]) && String(r[ECOL.DELETED - 1]).toUpperCase() !== 'Y')
    .map(r => {
      const start = cellText_(r[1]);
      const end = cellText_(r[2]) || start;
      return {
        id: String(r[0]),
        start: start,
        end: end < start ? start : end,
        title: String(r[3]),
        kind: EVENT_KINDS.indexOf(String(r[4])) >= 0 ? String(r[4]) : '기타',
        memo: String(r[5] || ''),
        author: String(r[6] || ''),
      };
    })
    .sort((a, b) => a.start.localeCompare(b.start));
}

/** 등록/수정. ev: { id?, start:'yyyy-mm-dd', end?, title, kind, memo } */
function apiSaveEvent(ev) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'dash');
  if (!me.isLeader && !me.isAdmin) throw new Error('일정은 총괄 또는 관리자만 등록·수정할 수 있습니다.');
  const title = String(ev.title || '').trim();
  const start = String(ev.start || '').trim();
  const end = String(ev.end || '').trim();
  if (!title || !/^\d{4}-\d{2}-\d{2}$/.test(start)) throw new Error('제목과 날짜를 입력해 주세요.');
  if (end && end < start) throw new Error('종료일이 시작일보다 빠릅니다.');
  const kind = EVENT_KINDS.indexOf(ev.kind) >= 0 ? ev.kind : '기타';

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  let id;
  try {
    let sh = yearSheet_(ctx.ss, SHEET.EVENTS);
    if (!sh) { setupEventSheet_(ctx.ss, ctx.year); sh = yearSheet_(ctx.ss, SHEET.EVENTS); }
    const last = sh.getLastRow();
    const ids = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(([v]) => String(v)) : [];
    const values = [parseDate_(start), end && end !== start ? parseDate_(end) : '', title, kind, String(ev.memo || '').trim(), me.name || me.email];
    id = String(ev.id || '');
    const idx = id ? ids.indexOf(id) : -1;
    if (idx >= 0) {
      sh.getRange(idx + 2, ECOL.START, 1, values.length).setValues([values]);
    } else {
      id = String((ids.length ? Math.max.apply(null, ids.map(Number).filter(n => !isNaN(n))) : 0) + 1);
      sh.getRange(last + 1, 1, 1, ECOL.DELETED).setValues([[id].concat(values, [''])]);
    }
  } finally {
    lock.releaseLock();
  }
  return Object.assign(refreshPart_('events'), { savedId: id });
}

function apiDeleteEvent(id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'dash');
  if (!me.isLeader && !me.isAdmin) throw new Error('일정은 총괄 또는 관리자만 삭제할 수 있습니다.');
  const sh = yearSheet_(ctx.ss, SHEET.EVENTS);
  if (!sh) throw new Error('일정 시트가 없습니다.');
  const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(([v]) => String(v));
  const idx = ids.indexOf(String(id));
  if (idx < 0) throw new Error('일정을 찾을 수 없습니다.');
  sh.getRange(idx + 2, ECOL.DELETED).setValue('Y');
  return refreshPart_('events');
}
