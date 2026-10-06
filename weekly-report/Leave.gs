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
  TYPES: ['연차', '오전반차', '오후반차', '정기휴가', '경조휴가', '병가', '대체휴무', '출장', '교육', '기타'],
  // 캘린더 일정 제목에서 휴가로 볼 단어 (긴 것부터)
  WORDS: /(오전반차|오후반차|반차|정기휴가|경조휴가|하계휴가|여름휴가|대체휴무|병가|연차|휴가|휴무)/,
};

const LCOL = { ID: 1, DATE: 2, NAME: 3, TYPE: 4, START: 5, END: 6, MEMO: 7, EVENT: 8, SOURCE: 9, DELETED: 10 };

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
  return sh && sh.getLastRow() >= 2 ? sh.getRange(2, 1, sh.getLastRow() - 1, LCOL.DELETED).getValues() : [];
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
      const ev = getTeamCalendar_().createAllDayEvent(`[${type}] ${me.name}`, start, addDays(end, 1),
        { description: [memo, `등록: ${me.name} (업무관리 웹앱)`].filter(Boolean).join('\n') });
      eventId = ev.getId();
    } catch (e) {
      warning = '팀 캘린더에 일정을 만들 권한이 없어 주간보고에만 반영했습니다.';
    }
  }

  const sh = setupLeaveSheet_(ctx.ss);
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    sh.getRange(sh.getLastRow() + 1, 1, 1, LCOL.DELETED)
      .setValues([[nextLeaveId_(sh), new Date(), me.name, type, start, end, memo, eventId, '웹앱', '']]);
  } finally {
    lock.releaseLock();
  }
  ctx._leaves = null;
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
  if (eventId) {
    try {
      const ev = getTeamCalendar_().getEventById(eventId);
      if (ev) ev.deleteEvent();
    } catch (e) { /* 캘린더 권한이 없으면 시트에서만 지운다 */ }
  }
  ctx._leaves = null;
  return leavePatch_(ctx, name);
}

/** 주간보고 작성 대상이면 그 사람 조각만 돌려준다 (팀장 등은 주간보고에 없으므로 생략) */
function leavePatch_(ctx, name) {
  return ctx.members.some(m => m.name === name && m.write) ? memberPatch_(ctx, name) : {};
}

function nextLeaveId_(sh) {
  return leaveRows_(sh).reduce((m, r) => Math.max(m, Number(r[0]) || 0), 0) + 1;
}

/* ---------- 캘린더 → [휴가] 맞추기 ---------- */

/**
 * 캘린더 일정 중 제목에 팀원 이름과 휴가 단어가 있는 것을 [휴가]에 넣고,
 * 캘린더에서 지워진 일정(출처 '캘린더')은 [휴가]에서도 지운다.
 * events: CalendarEvent[] (from~to 사이를 모두 읽은 것). 바뀐 게 있으면 true.
 */
function syncLeaveEvents_(ctx, events, from, to) {
  const lo = ctx.thisWeek.monday > from ? ctx.thisWeek.monday : from;
  const afterEnd = addDays(ctx.afterWeek.monday, 7);
  const hi = afterEnd < to ? afterEnd : to;
  if (lo >= hi) return false;

  const names = ctx.members.map(m => m.name).filter(Boolean).sort((a, b) => b.length - a.length);
  const found = {};
  events.forEach(ev => {
    const title = ev.getTitle();
    const word = title.match(LEAVE.WORDS);
    if (!word) return;
    const name = names.find(n => title.indexOf(n) >= 0);
    if (!name) return;
    const allDay = ev.isAllDayEvent();
    const start = allDay ? ev.getAllDayStartDate() : startOfDay_(ev.getStartTime());
    const end = allDay ? addDays(ev.getAllDayEndDate(), -1) : startOfDay_(ev.getEndTime());
    if (start >= hi || end < lo) return;
    found[ev.getId()] = { name: name, type: leaveType_(word[1], title), start: start, end: end < start ? start : end };
  });

  const sh = ctx.ss.getSheetByName(LEAVE.SHEET);
  const rows = leaveRows_(sh);
  const known = {};
  rows.forEach(r => { if (r[LCOL.EVENT - 1]) known[String(r[LCOL.EVENT - 1])] = true; });

  const add = Object.keys(found).filter(id => !known[id]);
  const gone = [];
  rows.forEach((r, i) => {
    if (String(r[LCOL.SOURCE - 1]) !== '캘린더' || String(r[LCOL.DELETED - 1]).toUpperCase() === 'Y') return;
    const s = r[LCOL.START - 1];
    if (!(s instanceof Date) || s < lo || s >= hi) return; // 이번에 읽은 범위 밖은 건드리지 않는다
    if (!found[String(r[LCOL.EVENT - 1])]) gone.push(i + 2);
  });
  if (!add.length && !gone.length) return false;

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const target = sh || setupLeaveSheet_(ctx.ss);
    gone.forEach(row => target.getRange(row, LCOL.DELETED).setValue('Y'));
    if (add.length) {
      let id = nextLeaveId_(target);
      const now = new Date();
      const values = add.map(eid => {
        const f = found[eid];
        return [id++, now, f.name, f.type, f.start, f.end, '', eid, '캘린더', ''];
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

/** 매시간 자동 실행: 이번 주~다다음 주 캘린더를 읽어 맞춘다 */
function syncLeavesFromCalendar_(ctx) {
  if (!teamCalendarSet_(ctx)) return false;
  const from = ctx.thisWeek.monday;
  const to = addDays(ctx.afterWeek.monday, 7);
  try {
    return syncLeaveEvents_(ctx, getTeamCalendar_().getEvents(from, to), from, to);
  } catch (e) {
    return false; // 캘린더 미설정·권한 없음
  }
}

/**
 * 팀 캘린더 ID가 있을 때만 맞춘다. 비어 있으면 사람마다 자기 기본 캘린더를 보게 되어
 * 다른 사람의 휴가 일정이 "지워진 것"처럼 보일 수 있기 때문이다.
 */
function teamCalendarSet_(ctx) {
  return !!String(ctx.cfg[CFG.CALENDAR] || '').trim();
}

/** 캘린더 화면에서 휴가 일정을 다른 색으로 보여주기 위한 판별 */
function isLeaveTitle_(title, names) {
  return LEAVE.WORDS.test(title) && names.some(n => n && title.indexOf(n) >= 0);
}

function leaveType_(word, title) {
  // "오후 반차"처럼 띄어 쓴 것도 오전/오후반차로
  if (word === '반차') return /오전/.test(title) ? '오전반차' : /오후/.test(title) ? '오후반차' : '반차';
  if (word === '하계휴가' || word === '여름휴가') return '정기휴가';
  return word;
}

function startOfDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
