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

/** startIso ~ endIso 사이 일정 + 공휴일 + 주간보고 마감 */
function apiCalendar(startIso, endIso) {
  requireMenu_(currentMember_(getMembers()), 'calendar');
  const start = new Date(startIso);
  const end = new Date(endIso);
  const cal = getTeamCalendar_();
  const ctx = getContext();
  const raw = cal.getEvents(start, end);

  // 캘린더에 직접 넣은 휴가 일정을 주간보고에 반영 (팀 캘린더가 설정된 경우)
  if (teamCalendarSet_(ctx)) {
    try { syncLeaveEvents_(ctx, raw, start, end); } catch (e) { /* 시트 권한 등 — 매시간 자동 실행 때 다시 맞춘다 */ }
  }
  const names = ctx.members.map(m => m.name);

  const events = raw.map(e => ({
    id: e.getId(),
    title: e.getTitle(),
    start: e.getStartTime().toISOString(),
    end: e.getEndTime().toISOString(),
    allDay: e.isAllDayEvent(),
    location: e.getLocation() || '',
    desc: e.getDescription() || '',
    kind: isLeaveTitle_(e.getTitle(), names) ? 'leave' : 'event',
  }));

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.HOLIDAYS);
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([d, name]) => {
      if (d instanceof Date && d >= addDays(start, -1) && d < end) {
        events.push({ id: 'h-' + ymd(d), title: String(name), start: ymd(d), end: ymd(addDays(d, 1)), allDay: true, location: '', desc: '', kind: 'holiday' });
      }
    });
  }

  if (ctx.deadline >= start && ctx.deadline < end) {
    events.push({ id: 'deadline', title: '주간보고 마감', start: ctx.deadline.toISOString(), end: ctx.deadline.toISOString(), allDay: false, location: '', desc: '', kind: 'deadline' });
  }

  return { calendarName: cal.getName(), events: events };
}

/**
 * ev: { title, date:'yyyy-mm-dd', allDay, startTime:'HH:mm', endTime:'HH:mm', location, desc }
 */
function apiAddEvent(ev) {
  const title = String(ev.title || '').trim();
  const m = String(ev.date || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!title || !m) throw new Error('제목과 날짜를 입력해 주세요.');

  const me = currentMember_(getMembers());
  requireMenu_(me, 'calendar');
  const cal = getTeamCalendar_();
  const opts = {
    location: String(ev.location || '').trim(),
    description: [String(ev.desc || '').trim(), me.name ? `등록: ${me.name}` : ''].filter(Boolean).join('\n'),
  };
  const day = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  try {
    if (ev.allDay) {
      cal.createAllDayEvent(title, day, opts);
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
      cal.createEvent(title, s, e, opts);
    }
  } catch (err) {
    throw new Error('일정을 추가할 권한이 없습니다. 팀 캘린더의 "일정 변경" 권한을 관리자에게 요청하세요.');
  }
  return true;
}
