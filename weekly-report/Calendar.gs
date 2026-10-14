/**
 * 일정 (구글 캘린더 연동)
 *
 * [설정] > "팀 캘린더 ID" 에 공유 캘린더 ID를 넣으면 그 캘린더를, 비우면 접속자의 기본 캘린더를 보여준다.
 * 캘린더 ID: 구글 캘린더 > 설정 > (캘린더 선택) > 캘린더 통합 > 캘린더 ID
 * 팀원이 일정을 보려면 그 캘린더가 팀원에게 공유되어 있어야 하고, 추가하려면 "일정 변경" 권한이 필요하다.
 */

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
  const seen = {};
  team.events.forEach(e => { seen[e.id] = true; });
  return {
    calendarName: team.name,
    myCalendarName: mine.name,
    events: team.events.concat(mine.events.filter(e => !seen[e.id])), // 팀 일정에 초대된 것은 한 번만
  };
}

/**
 * 캘린더에 직접 넣은 휴가 일정을 주간보고에 반영 (팀 캘린더가 설정된 경우).
 * 달력을 그린 뒤 화면이 따로 부른다 → 달력 표시를 늦추지 않는다. 매시간 자동 실행 때도 맞춘다.
 */
function apiSyncLeaves(startIso, endIso) {
  const ctx = getContext();
  requireMenu_(currentMember_(ctx.members), 'calendar');
  if (!teamCalendarSet_(ctx)) return { changed: false };
  const start = new Date(startIso), end = new Date(endIso);
  const changed = syncLeaveEvents_(ctx, getTeamCalendar_().getEvents(start, end), start, end);
  if (!changed) return { changed: false };
  return { changed: true, members: cached_(CK.members(ctx), () => readMembers_(ctx)) };
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
  const mine = ev.target === 'mine';
  const cal = mine ? CalendarApp.getDefaultCalendar() : getTeamCalendar_();
  const opts = {
    location: String(ev.location || '').trim(),
    description: [String(ev.desc || '').trim(), me.name ? `등록: ${me.name}` : ''].filter(Boolean).join('\n'),
  };
  const day = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  let created = null;
  try {
    if (ev.allDay) {
      created = cal.createAllDayEvent(title, day, opts);
    } else {
      const at = (t, fallback) => {
        const [h, mi] = String(t || fallback).split(':').map(Number);
        const d = new Date(day.getTime());
        d.setHours(h || 0, mi || 0, 0, 0);
        return d;
      };
      const s = at(ev.startTime, '09:00');
      let e = at(ev.endTime, '10:00');
      if (e <= s) e = new Date(s.getTime() + 3600 * 1000);
      created = cal.createEvent(title, s, e, opts);
    }
  } catch (err) {
    if (mine) throw new Error('내 캘린더에 일정을 만들지 못했습니다. 잠시 후 다시 시도하세요.');
    throw new Error('일정을 추가할 권한이 없습니다. 팀 캘린더의 "일정 변경" 권한을 관리자에게 요청하세요.');
  }
  // 개인 캘린더를 다른 사람과 공유해 두었더라도 내용은 안 보이게
  if (mine && created) {
    try { created.setVisibility(CalendarApp.Visibility.PRIVATE); } catch (e) { /* 무시 */ }
  }
  bumpCalendar_();
  return true;
}
