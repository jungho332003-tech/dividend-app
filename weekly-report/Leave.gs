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
  TYPES: ['연차', '오전반차', '오후반차', '오전반반차', '오후반반차', '정기휴가', '경조휴가', '대체휴무', '출장', '교육', '기타'],
  // 캘린더 일정 제목에서 휴가로 볼 단어 (긴 것부터)
  WORDS: /(오전반반차|오후반반차|반반차|오전반차|오후반차|반차|정기휴가|경조휴가|하계휴가|여름휴가|대체휴무|연차|휴가|휴무)/,
  // 이 단어들은 제목 맨 앞 "출장] 정호"처럼 태그로 쓸 때만 휴가로 본다 (본문에 흔히 나오는 말이라)
  TAG_ONLY: ['출장', '교육', '기타'],
};

/** 기타 휴무는 제목에 사유를 붙인다: "기타] 정호 (UFO 대체)" — 메모 첫 줄 */
function leaveReason_(type, memo) {
  if (type !== '기타') return '';
  return String(memo || '').split('\n')[0].replace(/[()]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 30);
}

/** 기타] 제목의 괄호 사유는 이름 찾기에서 뺀다 ("기타] 정호 (하은 대신 근무)" → 하은은 휴가 아님) */
function leaveNameZone_(title) {
  const t = String(title || '');
  return /^\s*\[?\s*기타\s*\]/.test(t) ? t.replace(/\([^)]*\)\s*$/, '') : t;
}

/** 제목에서 휴가 종류 단어 → match 결과처럼 [전체, 단어] (없으면 null) */
function leaveWord_(title) {
  const t = String(title || '');
  const w = t.match(LEAVE.WORDS);
  if (w) return w;
  const tag = t.match(/^\s*\[?\s*([가-힣]+)\s*\]/);
  return tag && LEAVE.TAG_ONLY.indexOf(tag[1]) >= 0 ? [tag[1], tag[1]] : null;
}

const LCOL = { ID: 1, DATE: 2, NAME: 3, TYPE: 4, START: 5, END: 6, MEMO: 7, EVENT: 8, SOURCE: 9, DELETED: 10, REC_EVENT: 11 };

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
  return sh && sh.getLastRow() >= 2 ? sh.getRange(2, 1, sh.getLastRow() - 1, LCOL.REC_EVENT).getValues() : [];
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
      const ev = getTeamCalendar_().createAllDayEvent(leaveEventTitle_(type, shortName_(me.name, ctx.members.map(m => m.name)), memo), start, addDays(end, 1),
        { description: [memo, `등록: ${me.name} (업무관리 웹앱)`].filter(Boolean).join('\n') });
      eventId = ev.getId();
    } catch (e) {
      warning = '팀 캘린더에 일정을 만들 권한이 없어 주간보고에만 반영했습니다.';
    }
  }
  // 채용담당은 채용 캘린더에도 같이 넣는다
  let recEventId = '';
  if (me.recruiter && canRecruitCal_(ctx, me)) { // 채용담당 본인 휴가만 (팀장 휴가는 팀 캘린더에만)
    try {
      recEventId = getRecruitCalendar_(ctx).createAllDayEvent(leaveEventTitle_(type, shortName_(me.name, ctx.members.map(m => m.name)), memo), start, addDays(end, 1),
        { description: [memo, `등록: ${me.name} (업무관리 웹앱)`].filter(Boolean).join('\n') }).getId();
    } catch (e) {
      warning = warning || '팀 캘린더에는 넣었지만 채용 캘린더에는 권한이 없어 넣지 못했습니다.';
    }
  }

  const sh = setupLeaveSheet_(ctx.ss);
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    if (recEventId) ensureHeader_(sh, LCOL.REC_EVENT, '채용캘린더 일정ID');
    sh.getRange(sh.getLastRow() + 1, 1, 1, LCOL.REC_EVENT)
      .setValues([[nextLeaveId_(sh), new Date(), me.name, type, start, end, memo, eventId, '웹앱', '', recEventId]]);
  } finally {
    lock.releaseLock();
  }
  ctx._leaves = null;
  if (eventId) bumpCalendar_();
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
  // 여러 명이 같이 적힌 일정("…#이름")은 이 사람 휴가만 지우고 캘린더 일정은 그대로 둔다
  if (eventId && eventId.indexOf('#') < 0) {
    try {
      const fromRecruit = String(rows[idx][LCOL.SOURCE - 1]) === RECRUIT_LEAVE_SOURCE;
      const ev = (fromRecruit ? getRecruitCalendar_(ctx) : getTeamCalendar_()).getEventById(eventId);
      if (ev) { ev.deleteEvent(); bumpCalendar_(); }
    } catch (e) { /* 캘린더 권한이 없으면 시트에서만 지운다 */ }
  }
  dropRecruitLeaveEvent_(ctx, rows[idx][LCOL.REC_EVENT - 1]);
  ctx._leaves = null;
  return leavePatch_(ctx, name);
}

/**
 * 일정 화면에서 휴가 일정을 지운다. [휴가]에 그 일정으로 등록된 휴가가 있으면 휴가째 지우고(주간보고에서도 빠짐),
 * 캘린더에만 있던 휴가 일정이면 캘린더에서 지운다. 본인 휴가만 (팀장·관리자는 모두).
 */
function apiDeleteLeaveEvent(eventId) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'calendar');
  eventId = String(eventId || '');
  const sh = ctx.ss.getSheetByName(LEAVE.SHEET);
  const rows = leaveRows_(sh);
  const live = r => String(r[LCOL.DELETED - 1]).toUpperCase() !== 'Y';
  let idx = rows.findIndex(r => String(r[LCOL.EVENT - 1]) === eventId && live(r));
  if (idx < 0) idx = rows.findIndex(r => String(r[LCOL.EVENT - 1]) === `${eventId}#${me.name}` && live(r)); // 여러 명 일정 중 내 것
  if (idx >= 0) return apiDeleteLeave(rows[idx][0]);

  const ev = getTeamCalendar_().getEventById(eventId);
  if (!ev) throw new Error('일정을 찾을 수 없습니다. 이미 지워졌을 수 있습니다.');
  const title = ev.getTitle();
  if (!isLeaveTitle_(title, ctx.members.map(m => m.name))) throw new Error('휴가 일정만 여기서 지울 수 있습니다.');
  const who = leaveNames_(title, ctx.members.map(m => m.name));
  if (!me.isLeader && !me.isAdmin && (who.length !== 1 || who[0] !== me.name)) throw new Error(who.length > 1 ? '여러 명이 함께 적힌 휴가 일정은 팀장·관리자만 지울 수 있습니다.' : '본인 휴가만 지울 수 있습니다.');
  try { ev.deleteEvent(); } catch (e) { throw new Error('캘린더에서 지울 권한이 없습니다. 팀 캘린더 "일정 변경" 권한을 확인하세요.'); }
  bumpCalendar_();
  return {};
}

/**
 * 일정 화면에서 휴가 고치기: 종류·기간·메모. [휴가] 시트·팀 캘린더·채용 캘린더를 같이 고친다.
 * 본인 휴가만 (팀장·관리자는 모두). 여러 명이 함께 적힌 일정은 지우기만 된다.
 */
function apiUpdateLeaveEvent(eventId, leave) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'calendar');
  eventId = String(eventId || '');
  const type = LEAVE.TYPES.indexOf(leave.type) >= 0 ? leave.type : '기타';
  const start = parseDue_(leave.start);
  const end = leave.end ? parseDue_(leave.end) : start;
  if (!(start instanceof Date)) throw new Error('시작일을 골라 주세요.');
  if (!(end instanceof Date)) throw new Error('종료일 형식이 올바르지 않습니다.');
  if (end < start) throw new Error('종료일이 시작일보다 빠릅니다.');
  const memo = String(leave.memo || '').trim();
  const names = ctx.members.map(m => m.name);

  const sh = ctx.ss.getSheetByName(LEAVE.SHEET);
  const rows = leaveRows_(sh);
  const live = r => String(r[LCOL.DELETED - 1]).toUpperCase() !== 'Y';
  const idx = rows.findIndex(r => String(r[LCOL.EVENT - 1]) === eventId && live(r));
  const fromRecruit = idx >= 0 && String(rows[idx][LCOL.SOURCE - 1]) === RECRUIT_LEAVE_SOURCE;
  const cal = fromRecruit ? getRecruitCalendar_(ctx) : getTeamCalendar_();
  const ev = cal.getEventById(eventId);
  let name;
  if (idx >= 0) {
    name = String(rows[idx][LCOL.NAME - 1]).trim();
  } else {
    if (!ev) throw new Error('일정을 찾을 수 없습니다. 이미 지워졌을 수 있습니다.');
    const who = leaveNames_(leaveNameZone_(ev.getTitle()), names);
    if (who.length !== 1) throw new Error('여러 명이 함께 적힌 휴가 일정은 고칠 수 없습니다. 지우고 각자 다시 등록해 주세요.');
    name = who[0];
  }
  if (name !== me.name && !me.isLeader && !me.isAdmin) throw new Error('본인 휴가만 고칠 수 있습니다.');

  const title = leaveEventTitle_(type, shortName_(name, names), memo);
  const fix = e => {
    e.setTitle(title);
    e.setAllDayDates(start, addDays(end, 1));
    const keep = htmlText_(e.getDescription()).split('\n').filter(l => l.indexOf('등록: ') === 0);
    e.setDescription([memo].concat(keep).filter(Boolean).join('\n'));
  };
  if (ev) {
    try { fix(ev); } catch (e) { throw new Error('캘린더 일정을 고칠 권한이 없습니다. 팀 캘린더 "일정 변경" 권한을 확인하세요.'); }
  }
  if (idx >= 0) {
    const r = idx + 2;
    sh.getRange(r, LCOL.TYPE).setValue(type);
    sh.getRange(r, LCOL.START, 1, 3).setValues([[start, end, memo]]);
    const rec = rows[idx][LCOL.REC_EVENT - 1];
    if (rec && recruitCalendarId_(ctx.cfg)) {
      try { const re = getRecruitCalendar_(ctx).getEventById(String(rec)); if (re) fix(re); } catch (e) { /* 권한 없음 */ }
    }
  }
  bumpCalendar_();
  ctx._leaves = null;
  return leavePatch_(ctx, name);
}

/** 주간보고 작성 대상이면 그 사람 조각만 돌려준다 (팀장 등은 주간보고에 없으므로 생략) */
/** 웹앱 휴가와 같이 만든 채용 캘린더 일정 지우기 */
function dropRecruitLeaveEvent_(ctx, id) {
  if (!id || !recruitCalendarId_(ctx.cfg)) return;
  try { const ev = getRecruitCalendar_(ctx).getEventById(String(id)); if (ev) { ev.deleteEvent(); bumpCalendar_(); } } catch (e) { /* 권한 없음 */ }
}

/** 웹앱 휴가와 같이 만든 채용 캘린더 일정 날짜 맞추기 */
function moveRecruitLeaveEvent_(ctx, id, start, end) {
  if (!id || !recruitCalendarId_(ctx.cfg)) return;
  try { const ev = getRecruitCalendar_(ctx).getEventById(String(id)); if (ev) ev.setAllDayDates(start, addDays(end, 1)); } catch (e) { /* 권한 없음 */ }
}

function leavePatch_(ctx, name) {
  return ctx.members.some(m => m.name === name && m.write) ? memberPatch_(ctx, name) : {};
}

function nextLeaveId_(sh) {
  return leaveRows_(sh).reduce((m, r) => Math.max(m, Number(r[0]) || 0), 0) + 1;
}

/* ---------- 캘린더 → [휴가] 맞추기 ---------- */

/**
 * 캘린더 일정 중 제목에 팀원 이름과 휴가 단어가 있는 것을 [휴가]에 넣고,
 * 캘린더에서 지워진 일정은 [휴가]에서도 지운다 (캘린더에서 넣은 것, 웹앱에서 넣은 것 모두).
 * events: CalendarEvent[] (from~to 사이를 모두 읽은 것). 바뀐 게 있으면 true.
 */
/**
 * opt.source: '캘린더'(팀 캘린더, 기본) | '채용캘린더'(채용 캘린더 — 채용담당 본인 휴가만)
 * 각 캘린더는 자기가 넣은 [휴가] 줄만 고치고 지운다.
 */
function syncLeaveEvents_(ctx, events, from, to, cal, opt) {
  const source = (opt && opt.source) || '캘린더';
  const fromRecruit = source === RECRUIT_LEAVE_SOURCE;
  const lo = ctx.thisWeek.monday > from ? ctx.thisWeek.monday : from;
  const afterEnd = addDays(ctx.afterWeek.monday, 7);
  const hi = afterEnd < to ? afterEnd : to;
  if (lo >= hi) return false;

  const names = ((opt && opt.names) || ctx.members.map(m => m.name)).filter(Boolean).sort((a, b) => b.length - a.length);
  if (!names.length) return false;
  const found = {};
  const present = {}; // 캘린더에 아직 있는 일정 (제목을 바꿨어도 지운 것으로 보지 않도록 전부)
  // 일정 하나 → { 키: 휴가 } (여러 명이면 사람마다 한 줄씩, [휴가] 캘린더 일정ID 칸: "일정ID#이름")
  const leavesOfEvent = ev => {
    const out = {};
    const title = ev.getTitle();
    const word = leaveWord_(title);
    if (!word) return out;
    const who = leaveNames_(leaveNameZone_(title), names);
    const allDay = ev.isAllDayEvent();
    const start = allDay ? ev.getAllDayStartDate() : startOfDay_(ev.getStartTime());
    const end = allDay ? addDays(ev.getAllDayEndDate(), -1) : startOfDay_(ev.getEndTime());
    who.forEach(name => {
      out[who.length > 1 ? `${ev.getId()}#${name}` : ev.getId()] = { name: name, type: leaveType_(word[1], title), start: start, end: end < start ? start : end };
    });
    return out;
  };
  events.forEach(ev => {
    present[ev.getId()] = true;
    const one = leavesOfEvent(ev);
    Object.keys(one).forEach(k => { if (one[k].start < hi && one[k].end >= lo) found[k] = one[k]; });
  });
  // 읽은 범위 밖으로 옮겨진 일정은 하나씩 찾아본다 (지운 것과 구분)
  const lookup = {};
  const moved = key => {
    const id = key.split('#')[0];
    if (!(id in lookup)) {
      let one = null;
      try {
        const ev = cal ? cal.getEventById(id) : null;
        one = ev ? leavesOfEvent(ev) : null;
      } catch (e) { one = null; }
      lookup[id] = one;
    }
    const f = lookup[id] ? lookup[id][key] || null : null;
    // 읽은 범위 안에 있다면 getEvents에 나왔어야 한다 → 안 나왔으면 지운 일정으로 본다
    return f && (f.start >= hi || f.end < lo) ? f : null;
  };
  const day = d => (d instanceof Date ? ymd(d) : '');

  const sh = ctx.ss.getSheetByName(LEAVE.SHEET);
  const rows = leaveRows_(sh);
  const known = {};
  rows.forEach(r => {
    if (r[LCOL.EVENT - 1]) known[String(r[LCOL.EVENT - 1])] = true;
    if (r[LCOL.REC_EVENT - 1]) known[String(r[LCOL.REC_EVENT - 1])] = true; // 웹앱 휴가와 같이 만든 채용 캘린더 일정
  });
  const mine = r => (String(r[LCOL.SOURCE - 1]) === RECRUIT_LEAVE_SOURCE) === fromRecruit; // 이 캘린더가 맡은 줄
  // 다른 캘린더(또는 웹앱)에 같은 사람·같은 날짜·같은 종류 휴가가 이미 있으면 또 넣지 않는다
  const same = {};
  rows.forEach(r => {
    if (String(r[LCOL.DELETED - 1]).toUpperCase() === 'Y' || !(r[LCOL.START - 1] instanceof Date)) return;
    const e = r[LCOL.END - 1] instanceof Date ? r[LCOL.END - 1] : r[LCOL.START - 1];
    same[[String(r[LCOL.NAME - 1]).trim(), ymd(r[LCOL.START - 1]), ymd(e), String(r[LCOL.TYPE - 1]).trim()].join('|')] = true;
  });

  const add = Object.keys(found).filter(id => !known[id] && !same[[found[id].name, ymd(found[id].start), ymd(found[id].end), found[id].type].join('|')]);
  const gone = [];
  const update = []; // 캘린더에서 날짜·종류를 바꾼 휴가 [행, 휴가]
  rows.forEach((r, i) => {
    const eid = String(r[LCOL.EVENT - 1] || '');
    if (!eid || String(r[LCOL.DELETED - 1]).toUpperCase() === 'Y' || !mine(r)) return; // 캘린더 일정이 없는 휴가는 그대로
    const s = r[LCOL.START - 1];
    const inRange = s instanceof Date && s >= lo && s < hi;
    let f = found[eid];
    if (!f && inRange) {
      const fromCal = String(r[LCOL.SOURCE - 1]) === source;
      // 캘린더에서 넣은 것: 휴가 일정이 아니게 됐거나 지워지면 / 웹앱에서 넣은 것: 캘린더에서 지워지면
      // (단, 날짜를 읽은 범위 밖으로 옮긴 것이면 지우지 않고 날짜를 따라간다)
      const gonePlain = fromCal ? true : !present[eid.split('#')[0]];
      if (gonePlain) {
        f = moved(eid);
        if (!f) { gone.push(i + 2); return; }
      }
    }
    if (!f) return;
    const e0 = r[LCOL.END - 1] instanceof Date ? r[LCOL.END - 1] : s;
    if (day(s) !== ymd(f.start) || day(e0) !== ymd(f.end) || String(r[LCOL.TYPE - 1]).trim() !== f.type) update.push([i + 2, f]);
  });
  if (!add.length && !gone.length && !update.length) return false;

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const target = sh || setupLeaveSheet_(ctx.ss);
    gone.forEach(row => {
      target.getRange(row, LCOL.DELETED).setValue('Y');
      if (!fromRecruit) dropRecruitLeaveEvent_(ctx, rows[row - 2][LCOL.REC_EVENT - 1]);
    });
    update.forEach(([row, f]) => {
      target.getRange(row, LCOL.TYPE).setValue(f.type);
      target.getRange(row, LCOL.START, 1, 2).setValues([[f.start, f.end]]);
      if (!fromRecruit) moveRecruitLeaveEvent_(ctx, rows[row - 2][LCOL.REC_EVENT - 1], f.start, f.end);
    });
    if (add.length) {
      let id = nextLeaveId_(target);
      const now = new Date();
      const values = add.map(eid => {
        const f = found[eid];
        return [id++, now, f.name, f.type, f.start, f.end, '', eid, source, ''];
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

/**
 * 팀 캘린더를 안 정했을 때: 웹앱에서 등록한 내 휴가는 내 기본 캘린더에 만들어진다.
 * 그 일정을 내가 캘린더에서 지웠으면 [휴가]에서도 지운다. (남의 기본 캘린더는 볼 수 없으므로 본인 것만)
 */
function syncMyLeaves_(ctx, me) {
  const sh = ctx.ss.getSheetByName(LEAVE.SHEET);
  if (!sh || !me.name) return false;
  const lo = ctx.thisWeek.monday, hi = addDays(ctx.afterWeek.monday, 7);
  const present = {};
  CalendarApp.getDefaultCalendar().getEvents(lo, hi).forEach(e => { present[e.getId()] = true; });
  const gone = [];
  leaveRows_(sh).forEach((r, i) => {
    const eid = String(r[LCOL.EVENT - 1] || '');
    const s = r[LCOL.START - 1];
    if (!eid || String(r[LCOL.NAME - 1]).trim() !== me.name || (String(r[LCOL.SOURCE - 1]) === '캘린더' || String(r[LCOL.SOURCE - 1]) === RECRUIT_LEAVE_SOURCE)
      || String(r[LCOL.DELETED - 1]).toUpperCase() === 'Y' || !(s instanceof Date) || s < lo || s >= hi) return;
    if (!present[eid]) gone.push(i + 2);
  });
  if (!gone.length) return false;
  gone.forEach(row => sh.getRange(row, LCOL.DELETED).setValue('Y'));
  ctx._leaves = null;
  try { CacheService.getDocumentCache().remove(CK.members(ctx)); } catch (e) { /* 캐시 없음 */ }
  return true;
}

/** 매시간 자동 실행: 이번 주~다다음 주 캘린더를 읽어 맞춘다 */
function syncLeavesFromCalendar_(ctx) {
  let changed = false;
  if (teamCalendarSet_(ctx)) {
    const from = ctx.thisWeek.monday;
    const to = addDays(ctx.afterWeek.monday, 7);
    try {
      const cal = getTeamCalendar_();
      changed = syncLeaveEvents_(ctx, cal.getEvents(from, to), from, to, cal);
    } catch (e) { /* 캘린더 미설정·권한 없음 */ }
  }
  return syncRecruitLeaves_(ctx) || changed;
}

/**
 * 채용 캘린더에 적힌 연차·반차·휴가 → [휴가] (채용담당이 넣은 일정 중 채용담당 본인 휴가만, 출처 "채용캘린더")
 * 팀 캘린더에 같은 휴가가 이미 있으면 한 번만 들어간다.
 */
const RECRUIT_LEAVE_SOURCE = '채용캘린더';
function syncRecruitLeaves_(ctx) {
  if (!recruitCalendarId_(ctx.cfg)) return false;
  const names = ctx.members.filter(m => m.recruit && m.name).map(m => m.name);
  if (!names.length) return false;
  const from = ctx.thisWeek.monday;
  const to = addDays(ctx.afterWeek.monday, 7);
  try {
    const cal = getRecruitCalendar_(ctx);
    // 채용담당이 넣은 일정만
    const events = cal.getEvents(from, to).filter(e => recruitCreator_(e, ctx));
    return syncLeaveEvents_(ctx, events, from, to, cal, { source: RECRUIT_LEAVE_SOURCE, names: names });
  } catch (e) {
    return false; // 채용 캘린더 권한 없음
  }
}

/**
 * 팀 캘린더 ID가 있을 때만 맞춘다. 비어 있으면 사람마다 자기 기본 캘린더를 보게 되어
 * 다른 사람의 휴가 일정이 "지워진 것"처럼 보일 수 있기 때문이다.
 */
function teamCalendarSet_(ctx) {
  return !!String(ctx.cfg[CFG.CALENDAR] || '').trim();
}

/**
 * 휴가 일정 제목을 팀 양식 "연차] 지선" 으로 맞춘다 (화면 표시용, 캘린더 원본은 그대로).
 * 이름은 성을 뺀 이름 (같은 이름이 둘이면 성까지). 휴가 일정이 아니면 '' 를 돌려준다.
 */
function leaveTitle_(title, names) {
  const word = leaveWord_(title);
  const who = leaveNames_(leaveNameZone_(title), names);
  if (!word || !who.length) return '';
  const type = leaveType_(word[1], title);
  // "[정기휴가] 윤수아", "출장] 수아"처럼 앞에 종류를 적어 두었으면 그 종류를 쓴다
  const tagged = String(title).match(/^\s*\[?\s*([가-힣]+)\s*\]/);
  const tag = tagged && LEAVE.TYPES.concat(['반차', '반반차', '휴가', '휴무']).indexOf(tagged[1]) >= 0 ? tagged[1] : type;
  const label = leaveLabel_(tag, who.map(n => shortName_(n, names)).join(', '));
  if (tag !== '기타') return label;
  // 기타: 사유도 보여준다 — 괄호 안, 없으면 이름 뒤에 적힌 말
  const paren = String(title).match(/\(([^)]*)\)\s*$/);
  let reason = paren ? paren[1] : String(title).replace(/^\s*\[?\s*기타\s*\]/, '');
  if (!paren) who.forEach(n => { reason = reason.split(n).join(' ').split(shortName_(n, names)).join(' '); });
  reason = reason.replace(/^[\s,·/-]+|[\s,·/-]+$/g, '').trim();
  return reason ? `${label} (${reason})` : label;
}

/** 캘린더에 만드는 휴가 일정 제목 ("기타"는 사유를 괄호로) */
function leaveEventTitle_(type, who, memo) {
  const r = leaveReason_(type, memo);
  return leaveLabel_(type, who) + (r ? ` (${r})` : '');
}

/** 팀 양식: "연차] 지선" */
function leaveLabel_(type, who) {
  return `${type}] ${who}`;
}

/** 성을 뺀 이름 ("안지선" → "지선"). 같은 이름을 쓰는 팀원이 또 있거나 형식이 다르면 그대로 */
function shortName_(name, names) {
  const ok = n => /^[가-힣]{3,4}$/.test(n);
  if (!ok(name)) return name;
  const g = name.slice(1);
  return (names || []).filter(n => ok(n) && n.slice(1) === g).length <= 1 ? g : name;
}

/** 캘린더 화면에서 휴가 일정을 다른 색으로 보여주기 위한 판별 */
function isLeaveTitle_(title, names) {
  return !!leaveWord_(title) && !!leaveName_(leaveNameZone_(title), names);
}

/**
 * 일정 제목에서 팀원 찾기 → [팀원] 시트의 이름(성 포함)
 *  1) "김정호 연차"처럼 이름 전체가 있으면 그 사람
 *  2) "연차] 정호"처럼 성을 뺀 이름만 있어도, 그 이름을 쓰는 팀원이 한 명뿐이면 그 사람
 *     (앞에 다른 한글이 붙은 "박정호"는 다른 사람으로 보고 건너뛴다)
 */
function leaveName_(title, names) {
  return leaveNames_(title, names)[0] || '';
}

/** "오전반차] 지애, 소영, 은지"처럼 여러 명이면 모두 (제목에 나온 순서대로) */
function leaveNames_(title, names) {
  const t = String(title || '');
  const list = names.filter(Boolean).slice().sort((a, b) => b.length - a.length);
  const hits = []; // [위치, 이름]
  const taken = [];
  const free = (at, len) => !taken.some(([a, b]) => at < b && at + len > a);
  list.forEach(n => {
    let at = t.indexOf(n);
    while (at >= 0) {
      if (free(at, n.length)) { hits.push([at, n]); taken.push([at, at + n.length]); }
      at = t.indexOf(n, at + 1);
    }
  });
  const given = {};
  list.forEach(n => {
    if (!/^[가-힣]{3,4}$/.test(n)) return;
    const g = n.slice(1);
    given[g] = given[g] ? '!' : n; // 같은 이름이 둘이면 누군지 모르므로 쓰지 않는다
  });
  Object.keys(given).filter(g => given[g] !== '!').forEach(g => {
    const re = new RegExp('(^|[^가-힣])' + g, 'g');
    let m;
    while ((m = re.exec(t))) {
      const at = m.index + m[1].length;
      if (free(at, g.length)) { hits.push([at, given[g]]); taken.push([at, at + g.length]); }
    }
  });
  return hits.sort((a, b) => a[0] - b[0]).map(h => h[1]).filter((n, i, a) => a.indexOf(n) === i);
}

function leaveType_(word, title) {
  // "오후 반차", "오전 반반차"처럼 띄어 쓴 것도 오전/오후로
  if (word === '반반차') return /오전/.test(title) ? '오전반반차' : /오후/.test(title) ? '오후반반차' : '반반차';
  if (word === '반차') return /오전/.test(title) ? '오전반차' : /오후/.test(title) ? '오후반차' : '반차';
  if (word === '하계휴가' || word === '여름휴가') return '정기휴가';
  return word;
}

function startOfDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
