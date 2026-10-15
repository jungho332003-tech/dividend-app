/**
 * 회의록 (주간업무보고 > 회의록 탭)
 *
 * [회의록]  번호 | 일자 | 시간 | 장소 | 참석자 | 방향성 | 항목(JSON) | 수정자 | 수정일시 | 삭제(Y)
 * 항목: [{ cat: 구분, sub: 세부, content: 내용, owner: 담당, due: 기한(자유 입력), status: 진행상태 }]
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
