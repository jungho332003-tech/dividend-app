/**
 * 일정 (구글 캘린더 연동)
 *
 * [설정] > "팀 캘린더 ID" 에 공유 캘린더 ID를 넣으면 그 캘린더를, 비우면 접속자의 기본 캘린더를 보여준다.
 * 캘린더 ID: 구글 캘린더 > 설정 > (캘린더 선택) > 캘린더 통합 > 캘린더 ID
 * 팀원이 일정을 보려면 그 캘린더가 팀원에게 공유되어 있어야 하고, 추가하려면 "일정 변경" 권한이 필요하다.
 */

/** [설정] "함께 볼 캘린더" — 회사 공용·다른 팀 캘린더 등을 보기 전용으로 같이 띄운다 (최대 6개) */
const EXTRA_CAL_MAX = 6;
function extraCalendarIds_(cfg) {
  const team = String(cfg[CFG.CALENDAR] || '').trim();
  return String(cfg[CFG.EXTRA_CAL] || '').split(/[\s,;]+/).map(s => s.trim()).filter(Boolean)
    .filter((v, i, a) => a.indexOf(v) === i && v !== team).slice(0, EXTRA_CAL_MAX);
}

/** 채용 캘린더: [팀원] 채용담당=Y 인 사람만 보고 쓴다 */
function recruitCalendarId_(cfg) { return String(cfg[CFG.RECRUIT_CAL] || '').trim(); }
function canRecruitCal_(ctx, me) { return !!(me && me.recruiter && recruitCalendarId_(ctx.cfg)); }
function getRecruitCalendar_(ctx) {
  const cal = CalendarApp.getCalendarById(recruitCalendarId_(ctx.cfg));
  if (!cal) throw new Error('채용 캘린더에 접근할 수 없습니다. 캘린더가 공유되어 있는지 관리자에게 확인하세요.');
  return cal;
}

function getTeamCalendar_() {
  const id = String(sheetBase_().cfg[CFG.CALENDAR] || '').trim();
  const cal = id ? CalendarApp.getCalendarById(id) : CalendarApp.getDefaultCalendar();
  if (!cal) throw new Error('팀 캘린더에 접근할 수 없습니다. 캘린더가 공유되어 있는지 관리자에게 확인하세요.');
  return cal;
}

/* ---------- 달력 캐시 ----------
 * 구글 캘린더 읽기가 느려서 결과를 캐시한다.
 *  - 팀 캘린더: 팀 전체가 5분간 공유 (팀 캘린더를 안 정했으면 각자 기본 캘린더라 본인 캐시에)
 *  - 내 캘린더(나만 보기): 본인 캐시에 2분
 * 웹앱에서 일정·휴가를 저장/삭제하면 버전을 바꿔 바로 새로 읽는다.
 */
const CAL_TTL = 300;
const MY_CAL_TTL = 120;

function calVersion_() {
  try { return CacheService.getDocumentCache().get('calver') || '0'; } catch (e) { return '0'; }
}

function bumpCalendar_() {
  try { CacheService.getDocumentCache().put('calver', String(Date.now()), 21600); } catch (e) { /* 캐시 없음 */ }
}

function userCached_(key, build, ttl) {
  let cache = null;
  try {
    cache = CacheService.getUserCache();
    const hit = cache.get(key);
    if (hit) return JSON.parse(hit);
  } catch (e) { /* 캐시 오류는 무시 */ }
  const value = build();
  try { if (cache) cache.put(key, JSON.stringify(value), ttl); } catch (e) { /* 100KB 초과 등 */ }
  return value;
}

function eventJson_(e, kind) {
  return {
    id: e.getId(),
    title: e.getTitle(),
    start: e.getStartTime().toISOString(),
    end: e.getEndTime().toISOString(),
    allDay: e.isAllDayEvent(),
    location: e.getLocation() || '',
    desc: e.getDescription() || '',
    kind: kind,
  };
}

/** startIso ~ endIso 사이 일정 + 공휴일 */
function apiCalendar(startIso, endIso, fresh) {
  const ctx = getContext(); // 팀원 목록은 캐시에서
  requireMenu_(currentMember_(ctx.members), 'calendar');
  if (fresh) bumpCalendar_(); // [새로고침]: 구글 캘린더에서 바로 다시 읽는다
  const start = new Date(startIso);
  const end = new Date(endIso);
  const shared = teamCalendarSet_(ctx);
  const ver = calVersion_();
  const key = `cal:${ver}:${startIso}:${endIso}`;

  const readTeam = () => {
    const cal = getTeamCalendar_();
    const names = ctx.members.map(m => m.name);
    const events = cal.getEvents(start, end).map(e => {
      const ev = eventJson_(e, isLeaveTitle_(e.getTitle(), names) ? 'leave' : 'event');
      ev.title = leaveTitle_(ev.title, names) || ev.title;
      // "연차] 민경"처럼 휴가 형식인데 [팀원]에 없는 이름이면 휴가 색으로만 보여준다 (주간보고에는 안 들어감)
      if (ev.kind === 'event') { const t = leaveTagTitle_(ev.title); if (t) { ev.kind = 'leave'; ev.title = t; ev.ro = true; } }
      return ev;
    });
    const sh = ctx.ss.getSheetByName(SHEET.HOLIDAYS);
    if (sh && sh.getLastRow() >= 2) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([d, name]) => {
        if (d instanceof Date && d >= addDays(start, -1) && d < end) {
          events.push({ id: 'h-' + ymd(d), title: String(name), start: ymd(d), end: ymd(addDays(d, 1)), allDay: true, location: '', desc: '', kind: 'holiday' });
        }
      });
    }
    return { name: cal.getName(), events: events };
  };
  // 팀 캘린더를 안 정했으면 각자 기본 캘린더 → 다른 사람에게 보이면 안 되므로 본인 캐시에
  const team = shared ? cached_(key, readTeam, CAL_TTL) : userCached_(key, readTeam, CAL_TTL);

  // 내 개인 캘린더 (나만 보기) — 팀 캘린더가 따로 있을 때만. 웹앱은 접속자 본인 계정으로 돌아서 본인 것만 보인다
  let mine = { name: '', events: [] };
  if (shared) {
    try {
      mine = userCached_('my' + key, () => {
        const my = CalendarApp.getDefaultCalendar();
        if (my.getId() === String(ctx.cfg[CFG.CALENDAR] || '').trim()) return { name: '', events: [] };
        return { name: my.getName(), events: my.getEvents(start, end).map(e => eventJson_(e, 'mine')) };
      }, MY_CAL_TTL);
    } catch (e) { /* 개인 캘린더를 못 읽으면 팀 일정만 보여준다 */ }
  }
  // 함께 볼 캘린더 — 볼 수 있는지는 사람마다 다르므로 본인 캐시에
  const ids = extraCalendarIds_(ctx.cfg);
  let extras = [];
  if (ids.length) {
    extras = userCached_('x' + key + ':' + ids.join(','), () => ids.map((id, i) => {
      try {
        const c = CalendarApp.getCalendarById(id);
        if (!c) return { name: '', ok: false, events: [] };
        return { name: c.getName(), ok: true, events: c.getEvents(start, end).map(e => Object.assign(eventJson_(e, 'extra'), { cal: i, calName: c.getName() })) };
      } catch (e) { return { name: '', ok: false, events: [] }; }
    }), CAL_TTL);
  }
  // 채용 캘린더 (채용담당만) — 팀 캘린더와 같이 등록한 일정은 한 번만 보이게 제목·시작으로도 겹침을 뺀다
  const me = currentMember_(ctx.members);
  let rec = null;
  if (canRecruitCal_(ctx, me)) {
    try {
      rec = userCached_('r' + key, () => {
        const c = getRecruitCalendar_(ctx);
        return { name: c.getName(), ok: true, events: c.getEvents(start, end).map(e => eventJson_(e, 'recruit')) };
      }, CAL_TTL);
    } catch (e) { rec = { name: '', ok: false, events: [] }; }
  }
  const seen = {};
  const once = list => list.filter(e => !seen[e.id] && (seen[e.id] = true));
  // 같은 일정이 여러 캘린더에 있으면(초대 등) 한 번만: 팀 → 내 캘린더 → 함께 볼 캘린더 순
  const events = once(team.events).concat(once(mine.events));
  if (rec) {
    const same = {};
    events.forEach(e => { same[e.title + '|' + e.start] = true; });
    events.push.apply(events, once(rec.events).filter(e => !same[e.title + '|' + e.start]));
  }
  extras.forEach(x => { events.push.apply(events, once(x.events)); });
  const merged = mergeLeaves_(events, ctx.members.map(m => m.name));
  return {
    calendarName: team.name,
    myCalendarName: mine.name,
    extraCalendars: extras.map((x, i) => ({ cal: i, name: x.name, ok: x.ok })),
    recruitCalendar: rec ? { name: rec.name, ok: rec.ok } : null,
    // 캘린더에 들어 있는 공휴일은 [공휴일] 시트와 겹치면 뺀다
    events: dropHolidayDupes_(merged),
  };
}

/**
 * 구글 캘린더에 공휴일이 일정으로 들어 있으면(공휴일 캘린더를 합쳐 둔 경우 등) [공휴일] 시트의 공휴일과 두 번 보인다.
 * 같은 날의 종일 일정 중 이름이 공휴일과 같거나 서로 포함되면 빼고, 빨간 공휴일만 남긴다.
 */
/** "[연차] 민경", "오전반차] 소영" → 팀 양식 "연차] 민경" (휴가 단어로 시작하는 제목). 아니면 '' */
function leaveTagTitle_(title) {
  const m = /^\s*\[?\s*([가-힣]+?)\s*\]\s*(.+)$/.exec(String(title || ''));
  if (!m) return '';
  const w = m[1].match(LEAVE.WORDS);
  return w && w[0] === m[1] ? leaveLabel_(leaveType_(w[1], m[1]), m[2].trim()) : '';
}

/**
 * 채용 캘린더·내 캘린더·함께 볼 캘린더에 적힌 휴가도 "[연차] 안지선"처럼 이름을 맞춰 휴가 색으로 보여주고,
 * 팀 캘린더에 같은 날 같은 휴가가 있으면 하나만 남긴다. (팀 캘린더 것이 아니면 여기서 지우지는 못하게 ro)
 */
function mergeLeaves_(events, names) {
  const dayOf = e => (/^\d{4}-\d{2}-\d{2}$/.test(e.start) ? e.start : ymd(new Date(e.start)));
  const keyOf = e => e.title.replace(/\s+/g, '') + '|' + dayOf(e);
  const seen = {};
  events.forEach(e => {
    if (e.kind === 'leave' || e.kind === 'event') {
      if (e.kind === 'leave') { seen[keyOf(e)] = true; e._base = true; } // 팀 캘린더 휴가가 기준
      return;
    }
    if (e.kind !== 'recruit' && e.kind !== 'mine' && e.kind !== 'extra') return;
    const t = (isLeaveTitle_(e.title, names) && leaveTitle_(e.title, names)) || leaveTagTitle_(e.title);
    if (!t) return;
    e.title = t;
    e.kind = 'leave';
    e.ro = true;
  });
  return events.filter(e => {
    if (e._base) { delete e._base; return true; }
    if (!e.ro || e.kind !== 'leave') return true;
    const k = keyOf(e);
    if (seen[k]) return false;
    seen[k] = true;
    return true;
  });
}

function dropHolidayDupes_(events) {
  const norm = t => String(t || '').replace(/[\s()\[\]·]/g, '');
  const words = t => String(t || '').split(/[\s()\[\]·]+/).filter(Boolean);
  // "개천절 대체공휴일" ↔ "대체공휴일(개천절)"처럼 순서가 달라도 같은 말이면 같은 공휴일로 본다
  // 거의 같은 이름만 ("한글날 행사 준비" 같은 진짜 일정은 남긴다)
  const EXTRA = ['', '공휴일', '휴일', '연휴', '대체공휴일', '대체휴일'];
  const same = (h, t) => h === t || (t.indexOf(h) >= 0 && EXTRA.indexOf(t.replace(h, '')) >= 0) || (h.indexOf(t) >= 0 && EXTRA.indexOf(h.replace(t, '')) >= 0);
  const dayOf = e => (/^\d{4}-\d{2}-\d{2}$/.test(e.start) ? e.start : ymd(new Date(e.start)));
  const holi = {};
  events.filter(e => e.kind === 'holiday').forEach(e => { (holi[e.start] = holi[e.start] || []).push(e.title); });
  return events.filter(e => {
    if (e.kind === 'holiday' || !e.allDay) return true;
    const names = holi[dayOf(e)];
    if (!names) return true;
    const t = norm(e.title);
    return !names.some(name => {
      const h = norm(name);
      if (!h || !t) return false;
      if (same(h, t)) return true;
      const ws = words(name);
      return ws.length > 1 && ws.every(w => t.indexOf(w) >= 0) && t.length <= h.length + 2;
    });
  });
}

/**
 * 캘린더에 넣은/지운 휴가 일정을 주간보고에 반영.
 * 달력을 그린 뒤 화면이 따로 부른다 → 달력 표시를 늦추지 않는다. 매시간 자동 실행 때도 맞춘다.
 */
function apiSyncLeaves() {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'calendar');
  let changed;
  if (teamCalendarSet_(ctx)) {
    // 이번 주~다다음 주 (주간보고에 나오는 범위)를 읽어 맞춘다
    const start = ctx.thisWeek.monday, end = addDays(ctx.afterWeek.monday, 7);
    const cal = getTeamCalendar_();
    changed = syncLeaveEvents_(ctx, cal.getEvents(start, end), start, end, cal);
  } else {
    changed = syncMyLeaves_(ctx, me); // 팀 캘린더가 없으면 내 캘린더에서 지운 내 휴가만
  }
  // 채용담당이 열면 채용 캘린더의 채용담당 휴가도 맞춘다
  if (canRecruitCal_(ctx, me)) { try { changed = syncRecruitLeaves_(ctx) || changed; } catch (e) { /* 다음에 다시 */ } }
  // 채용담당: 채용 캘린더에 넣은 입사·인터뷰를 인사현황으로
  let hrChanged = false;
  try { hrChanged = syncTeamHr_(ctx); } catch (e) { /* 다음에 다시 */ } // 팀 캘린더에서 고친 인사현황 일정
  if (canRecruitCal_(ctx, me)) { try { hrChanged = syncRecruitHr_(ctx) || hrChanged; } catch (e) { /* 다음에 다시 */ } }
  if (!changed && !hrChanged) return { changed: false };
  const out = { changed: true };
  if (changed) out.members = forViewer_(currentMember_(ctx.members), cached_(CK.members(ctx), () => readMembers_(ctx)));
  if (hrChanged) out.hr = cached_(CK.hr(ctx), () => readHr_(ctx));
  return out;
}

/**
 * ev: { title, date:'yyyy-mm-dd', allDay, startTime:'HH:mm', endTime:'HH:mm', location, desc, target:'team'|'mine' }
 * target 'mine' → 접속자 본인의 기본 캘린더에 "비공개"로 넣는다 (본인만 보기)
 */
function apiAddEvent(ev) {
  const title = String(ev.title || '').trim();
  const m = String(ev.date || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!title || !m) throw new Error('제목과 날짜를 입력해 주세요.');

  const me = currentMember_(sheetBase_().members); // 캐시 (시트를 매번 읽지 않도록)
  requireMenu_(me, 'calendar');
  const ctx = sheetBase_(); // 캐시 (설정·팀원)
  const recruitOnly = ev.target === 'recruit';
  const alsoRecruit = !!ev.recruit && !recruitOnly && ev.target !== 'mine';
  if ((recruitOnly || alsoRecruit) && !canRecruitCal_(ctx, currentMember_(ctx.members))) throw new Error('채용 캘린더는 채용담당만 쓸 수 있습니다.');
  const mine = ev.target === 'mine';
  // 팀 일정은 "이름) 업무명" 으로 만든다 (구글 캘린더에서도 누가 넣었는지 보이게). 나만 보기 일정은 그대로
  const fullTitle = !mine && me.name && title.indexOf(me.name + ')') !== 0 ? `${me.name}) ${title}` : title;
  const cal = mine ? CalendarApp.getDefaultCalendar() : recruitOnly ? getRecruitCalendar_(ctx) : getTeamCalendar_();
  const opts = {
    location: String(ev.location || '').trim(),
    description: [String(ev.desc || '').trim(), me.name ? `등록: ${me.name}` : ''].filter(Boolean).join('\n'),
  };
  const day = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  const make = c => {
    if (ev.allDay) return c.createAllDayEvent(fullTitle, day, opts);
    const at = (t, fallback) => {
      const [h, mi] = String(t || fallback).split(':').map(Number);
      const d = new Date(day.getTime());
      d.setHours(h || 0, mi || 0, 0, 0);
      return d;
    };
    const s = at(ev.startTime, '09:00');
    let e = at(ev.endTime, '10:00');
    if (e <= s) e = new Date(s.getTime() + 3600 * 1000);
    return c.createEvent(fullTitle, s, e, opts);
  };
  let created = null;
  try {
    created = make(cal);
  } catch (err) {
    if (mine) throw new Error('내 캘린더에 일정을 만들지 못했습니다. 잠시 후 다시 시도하세요.');
    if (recruitOnly) throw new Error('채용 캘린더에 일정을 추가할 권한이 없습니다. "일정 변경" 권한을 관리자에게 요청하세요.');
    throw new Error('일정을 추가할 권한이 없습니다. 팀 캘린더의 "일정 변경" 권한을 관리자에게 요청하세요.');
  }
  // 채용 캘린더에도 같이 (팀 캘린더에는 이미 들어갔으므로 실패해도 알리기만)
  let warning = '';
  if (alsoRecruit) {
    try { make(getRecruitCalendar_(ctx)); } catch (err) { warning = '팀 캘린더에는 넣었지만 채용 캘린더에는 권한이 없어 넣지 못했습니다.'; }
  }
  // 개인 캘린더를 다른 사람과 공유해 두었더라도 내용은 안 보이게
  if (mine && created) {
    try { created.setVisibility(CalendarApp.Visibility.PRIVATE); } catch (e) { /* 무시 */ }
  }
  bumpCalendar_();
  return warning ? { warning: warning } : true;
}

/* ---------- 내 메모 (일별, 나만 보기) ----------
 * 접속한 사람의 사용자 속성(UserProperties)에 저장한다 → 스프레드시트에 남지 않고, 관리자도 볼 수 없다.
 * 웹앱이 "접속한 사용자" 권한으로 돌아서 사람마다 따로 저장된다. 한 메모는 8,000자까지.
 */
const MEMO_PREFIX = 'memo:';

function apiMemos() {
  const all = PropertiesService.getUserProperties().getProperties();
  const out = {};
  Object.keys(all).forEach(k => { if (k.indexOf(MEMO_PREFIX) === 0) out[k.slice(MEMO_PREFIX.length)] = all[k]; });
  return out;
}

function apiSaveMemo(ymd, text) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(ymd))) throw new Error('날짜가 올바르지 않습니다.');
  const value = String(text || '').replace(/\s+$/, '');
  if (value.length > 8000) throw new Error('메모는 8,000자까지 쓸 수 있습니다.');
  const props = PropertiesService.getUserProperties();
  if (value.trim()) props.setProperty(MEMO_PREFIX + ymd, value);
  else props.deleteProperty(MEMO_PREFIX + ymd);
  return true;
}
