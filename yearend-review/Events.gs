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
}

function readEvents_(ss) {
  return dataRows_(yearSheet_(ss, SHEET.EVENTS), ECOL.DELETED)
    // 예전 설치본이 넣어 둔 예시 일정(작성자 '설치 예시')은 보여주지 않는다
    .filter(r => String(r[0]) && cellText_(r[1]) && String(r[ECOL.DELETED - 1]).toUpperCase() !== 'Y' && String(r[ECOL.AUTHOR - 1]) !== '설치 예시')
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
  requireOpen_(ctx, me);
  requireMenu_(me, 'dash');
  if (!me.isLeader && !me.isAdmin) throw new Error('일정은 총괄 또는 관리자만 등록·수정할 수 있습니다.');
  const title = String(ev.title || '').trim();
  const start = String(ev.start || '').trim();
  const end = String(ev.end || '').trim();
  if (!title || !/^\d{4}-\d{2}-\d{2}$/.test(start)) throw new Error('제목과 날짜를 입력해 주세요.');
  if (end && end < start) throw new Error('종료일이 시작일보다 빠릅니다.');
  const kind = EVENT_KINDS.indexOf(ev.kind) >= 0 ? ev.kind : '기타';

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error(LOCK_BUSY_MSG);
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
  requireOpen_(ctx, me);
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
