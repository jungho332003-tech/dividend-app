/**
 * 인사현황 (주간보고 아래 표)
 *
 *  구 분 | 금주: 대상 및 내용 | 일자/기간 | 차주: 대상 및 내용 | 발령일 | 비고
 *
 * [인사현황]  번호 | 등록일시 | 등록자 | 주차(월요일) | 구분 | 대상 및 내용 | 시작일 | 종료일 | 비고 | 삭제(Y) | 캘린더ID
 * 등록하면 팀 캘린더에도 "[구분] 대상 및 내용" 종일 일정으로 만든다 (고치면 같이 고치고, 지우면 같이 지운다).
 * 주차는 일자(시작일)가 속한 주의 월요일로 정한다 → 날짜만 넣으면 그 주에 맞게 들어가고,
 * 다음 주가 되면 "차주"였던 항목이 자동으로 "금주"로 보인다. 더 뒤의 날짜는 미리 등록해 두면 그 주가 됐을 때 나타난다.
 */

const HR = {
  SHEET: '인사현황',
  CATEGORIES: ['입사', '휴직', '복직', '퇴사', '인터뷰', '기타'],
  MIN_ROWS: 1, // 빈 줄 없이 등록한 만큼만
};

const HCOL = { ID: 1, DATE: 2, AUTHOR: 3, WEEK: 4, CATEGORY: 5, TARGET: 6, START: 7, END: 8, MEMO: 9, DELETED: 10, EVENT: 11, REC_EVENT: 12 };
// 채용 캘린더에도 같이 넣는 구분
const HR_RECRUIT_CATS = ['인터뷰', '입사'];

function setupHrSheet_(ss) {
  if (ss.getSheetByName(HR.SHEET)) return;
  const sh = ss.insertSheet(HR.SHEET);
  sh.getRange(1, 1, 1, HCOL.EVENT).setValues([['번호', '등록일시', '등록자', '주차(월요일)', '구분', '대상 및 내용', '시작일', '종료일', '비고', '삭제(Y)', '캘린더ID']]);
  styleHeader_(sh.getRange(1, 1, 1, HCOL.EVENT));
  sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange('D2:D').setNumberFormat('@');
  sh.getRange('G2:H').setNumberFormat('yyyy-mm-dd (ddd)');
  sh.getRange('E2:E').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(HR.CATEGORIES).build());
  [60, 130, 80, 100, 70, 280, 120, 120, 220, 70].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);
}

/** 이번 주·다음 주 항목 + 그 뒤에 미리 등록한 항목(later) */
function readHr_(ctx) {
  const sh = ctx.ss.getSheetByName(HR.SHEET);
  const out = { thisWeek: [], nextWeek: [], later: [] };
  if (!sh || sh.getLastRow() < 2) return out;
  const thisKey = ctx.thisWeek.key, nextKey = ctx.nextWeek.key;
  sh.getRange(2, 1, sh.getLastRow() - 1, HCOL.DELETED).getValues().forEach(r => {
    if (!String(r[0]) || String(r[HCOL.DELETED - 1]).toUpperCase() === 'Y') return;
    const start = r[HCOL.START - 1], end = r[HCOL.END - 1];
    // 주차: 일자가 있으면 그 날짜의 주, 없으면 (예전에 금주/차주를 골라 저장한) 주차 칸
    const week = start instanceof Date ? ymd(mondayOf(start))
      : r[HCOL.WEEK - 1] instanceof Date ? ymd(r[HCOL.WEEK - 1]) : String(r[HCOL.WEEK - 1]).trim();
    const list = week === thisKey ? out.thisWeek : week === nextKey ? out.nextWeek : week > nextKey ? out.later : null;
    if (!list) return;
    list.push({
      id: String(r[0]),
      author: String(r[HCOL.AUTHOR - 1]),
      category: String(r[HCOL.CATEGORY - 1]),
      target: String(r[HCOL.TARGET - 1]),
      start: start instanceof Date ? ymd(start) : '',
      end: end instanceof Date ? ymd(end) : '',
      when: hrPeriod_(start, end),
      memo: String(r[HCOL.MEMO - 1] || ''),
      week: week,
      weekLabel: list === out.later ? getWeekInfo(mondayOf(start), ctx.holidays).label : '',
    });
  });
  const byDate = (a, b) => (a.start || '9').localeCompare(b.start || '9') || Number(a.id) - Number(b.id);
  out.thisWeek.sort(byDate);
  out.nextWeek.sort(byDate);
  out.later.sort(byDate);
  return out;
}

/** 9/28(월) · 10/1(목)~14(수) · 9/30(수)~10/2(금) */
function hrPeriod_(start, end) {
  if (!(start instanceof Date)) return '';
  const md = d => `${d.getMonth() + 1}/${d.getDate()}(${DAY_KO[d.getDay()]})`;
  if (!(end instanceof Date) || ymd(end) === ymd(start)) return md(start);
  const tail = end.getMonth() === start.getMonth() && end.getFullYear() === start.getFullYear()
    ? `${end.getDate()}(${DAY_KO[end.getDay()]})` : md(end);
  return `${md(start)}~${tail}`;
}

/** 구분별로 금주·차주를 나란히 놓은 표 줄: [{category, rows:[{a, b}]}]. 등록한 항목이 없는 구분은 뺀다. */
function hrTable_(hr) {
  const used = HR.CATEGORIES.concat(hr.thisWeek.concat(hr.nextWeek).map(x => x.category))
    .filter((c, i, all) => all.indexOf(c) === i)
    .filter(c => hr.thisWeek.concat(hr.nextWeek).some(x => x.category === c));
  return used.map(cat => {
    const a = hr.thisWeek.filter(x => x.category === cat);
    const b = hr.nextWeek.filter(x => x.category === cat);
    const n = Math.max(HR.MIN_ROWS, a.length, b.length);
    const rows = [];
    for (let i = 0; i < n; i++) rows.push({ a: a[i] || null, b: b[i] || null });
    return { category: cat, rows: rows };
  });
}

/** 등록·수정. entry: { id?, category, target, start, end, memo } — 주차는 start 날짜로 정한다 */
function apiSaveHr(entry) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'weekly');
  if (!me.name) throw new Error('[팀원] 시트에 등록된 사람만 인사현황을 등록할 수 있습니다.');

  const category = HR.CATEGORIES.indexOf(entry.category) >= 0 ? entry.category : '기타';
  const target = String(entry.target || '').trim();
  if (!target) throw new Error('대상 및 내용을 입력해 주세요.');
  const start = parseDue_(entry.start);
  const end = parseDue_(entry.end);
  if (!(start instanceof Date)) throw new Error('일자를 골라 주세요. 날짜에 맞는 주에 들어갑니다.');
  if (end && !(end instanceof Date)) throw new Error('종료일 형식이 올바르지 않습니다.');
  if (start instanceof Date && end instanceof Date && end < start) throw new Error('종료일이 시작일보다 빠릅니다.');
  const week = ymd(mondayOf(start)); // 날짜가 속한 주
  const memo = String(entry.memo || '').trim();

  const sh = ctx.ss.getSheetByName(HR.SHEET) || (setupHrSheet_(ctx.ss), ctx.ss.getSheetByName(HR.SHEET));
  ensureHeader_(sh, HCOL.EVENT, '캘린더ID');
  ensureHeader_(sh, HCOL.REC_EVENT, '채용캘린더ID');
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  let row, oldEvent = '', oldRec = '';
  try {
    const last = sh.getLastRow();
    const ids = last >= 2 ? sh.getRange(2, 1, last - 1, HCOL.REC_EVENT).getValues() : [];
    if (entry.id) {
      const idx = ids.findIndex(r => String(r[0]) === String(entry.id));
      if (idx < 0) throw new Error('항목을 찾을 수 없습니다.');
      if (!canEditHrRow_(me, ids[idx][2])) throw new Error('본인이 등록한 항목만 고칠 수 있습니다.');
      row = idx + 2;
      oldEvent = String(ids[idx][HCOL.EVENT - 1] || '');
      oldRec = String(ids[idx][HCOL.REC_EVENT - 1] || '');
      sh.getRange(row, HCOL.WEEK, 1, 6).setValues([[week, category, target, start, end, memo]]);
    } else {
      const id = ids.reduce((m, r) => Math.max(m, Number(r[0]) || 0), 0) + 1;
      row = last + 1;
      sh.getRange(row, HCOL.WEEK).setNumberFormat('@');
      sh.getRange(row, 1, 1, HCOL.DELETED).setValues([[id, new Date(), me.name, week, category, target, start, end, memo, '']]);
    }
  } finally {
    lock.releaseLock();
  }
  // 팀 캘린더에도 같은 내용으로 (권한이 없으면 인사현황에만 저장)
  const ev = hrEvent_(me, oldEvent, category, target, start, end, memo);
  if (ev.id !== oldEvent) sh.getRange(row, HCOL.EVENT).setValue(ev.id);
  // 인터뷰·입사는 채용 캘린더에도 (채용 캘린더가 정해져 있을 때만, 권한이 없으면 조용히 건너뛴다)
  const rec = hrRecruitEvent_(ctx, me, oldRec, category, target, start, end, memo);
  if (rec !== oldRec) sh.getRange(row, HCOL.REC_EVENT).setValue(rec);
  return Object.assign(refreshHr_(ctx), ev.warning ? { warning: ev.warning } : {});
}

/** 인사현황 1건 → 팀 캘린더 종일 일정 "[구분] 대상 및 내용". 이미 있으면 고친다. { id, warning } */
function hrEvent_(me, eventId, category, target, start, end, memo) {
  if (allowedMenus_(me).indexOf('calendar') < 0) return { id: eventId };
  const title = `[${category}] ${target}`;
  const last = end instanceof Date ? end : start;
  const desc = [memo, `인사현황 (등록 ${me.name})`].filter(Boolean).join('\n');
  try {
    const cal = getTeamCalendar_();
    const ev = eventId ? cal.getEventById(eventId) : null;
    if (ev) {
      ev.setTitle(title);
      ev.setAllDayDates(start, addDays(last, 1));
      ev.setDescription(desc);
      bumpCalendar_();
      return { id: eventId };
    }
    const created = cal.createAllDayEvent(title, start, addDays(last, 1), { description: desc });
    bumpCalendar_();
    return { id: created.getId() };
  } catch (e) {
    return { id: eventId, warning: '팀 캘린더에 일정을 만들 권한이 없어 인사현황에만 저장했습니다.' };
  }
}

/** 인사현황 → 채용 캘린더 일정 (인터뷰·입사만). 구분이 바뀌어 대상이 아니면 지운다. 새 일정 ID(없으면 '') */
function hrRecruitEvent_(ctx, me, eventId, category, target, start, end, memo) {
  if (!recruitCalendarId_(ctx.cfg)) return eventId;
  try {
    const cal = getRecruitCalendar_(ctx);
    const ev = eventId ? cal.getEventById(eventId) : null;
    if (HR_RECRUIT_CATS.indexOf(category) < 0) {
      if (ev) { ev.deleteEvent(); bumpCalendar_(); }
      return '';
    }
    const title = `[${category}] ${target}`;
    const last = end instanceof Date ? end : start;
    const desc = [memo, `인사현황 (등록 ${me.name})`].filter(Boolean).join('\n');
    if (ev) {
      ev.setTitle(title);
      ev.setAllDayDates(start, addDays(last, 1));
      ev.setDescription(desc);
      bumpCalendar_();
      return eventId;
    }
    const created = cal.createAllDayEvent(title, start, addDays(last, 1), { description: desc });
    bumpCalendar_();
    return created.getId();
  } catch (e) {
    return eventId; // 채용 캘린더 권한이 없는 사람이 등록한 경우 등
  }
}

/** 본인 등록분 · 팀장 · 관리자. 채용 캘린더에서 가져온 항목은 채용담당도 */
function canEditHrRow_(me, author) {
  return author === me.name || me.isLeader || me.isAdmin || (author === RECRUIT_AUTHOR && !!me.recruiter);
}

/* ---------- 채용 캘린더 → 인사현황 ----------
 * 채용 캘린더에 직접 넣은 입사·인터뷰(면접) 일정을 인사현황으로 가져온다. (등록자: "채용캘린더")
 *  - 제목이 "[입사] …", "[인터뷰] …", "… 면접 …"처럼 입사·인터뷰·면접이 들어간 일정만
 *  - 캘린더에서 날짜·제목을 고치면 따라 바뀌고, 지우면 인사현황에서도 지워진다
 *  - 웹앱에서 등록해 채용 캘린더로 보낸 항목은 날짜가 바뀐 것만 맞춘다
 * 채용담당이 앱을 열 때와 매시간 자동 실행 때 맞춘다.
 */
const RECRUIT_AUTHOR = '채용캘린더';
const RECRUIT_SYNC_DAYS = 90; // 이번 주 월요일부터 석 달

function recruitCategory_(title) {
  const t = String(title || '').replace(/^\s*[가-힣A-Za-z]{2,10}\)\s*/, '');
  const head = (/^\s*\[([^\]]+)\]/.exec(t) || [])[1] || '';
  const pick = s => /입사(?!지원|서류)/.test(s) ? '입사' : /인터뷰|면접/.test(s) ? '인터뷰' : '';
  return pick(head) || pick(t);
}

/**
 * 제목에서 구분 표시를 뺀 나머지 → 인사현황 "대상 및 내용"
 *  "정하은) [인터뷰] 경력직 후보 A" → "경력직 후보 A"
 *  "** 인터뷰] 경력직 후보 A" → "경력직 후보 A"  ( ']'만 붙인 형식, 앞의 기호는 버린다)
 *  "홍길동 입사] 재무팀" → "홍길동 재무팀"
 */
function recruitTarget_(title) {
  const t = String(title || '').trim();
  let s = t.replace(/^\s*[가-힣A-Za-z]{2,10}\)\s*/, '').replace(/^\s*\[[^\]]*\]\s*/, '');
  const m = /^([^\[\]]*?)\s*(?:입사|인터뷰|면접)\s*\]\s*(.*)$/.exec(s);
  if (m) s = [m[1].replace(/^[^0-9A-Za-z가-힣]+/, '').trim(), m[2].trim()].filter(Boolean).join(' ');
  return s.trim() || t;
}

function syncRecruitHr_(ctx) {
  if (!recruitCalendarId_(ctx.cfg)) return false;
  const from = ctx.thisWeek.monday, to = addDays(from, RECRUIT_SYNC_DAYS);
  let events;
  try { events = getRecruitCalendar_(ctx).getEvents(from, to); } catch (e) { return false; }

  const sh = ctx.ss.getSheetByName(HR.SHEET) || (setupHrSheet_(ctx.ss), ctx.ss.getSheetByName(HR.SHEET));
  ensureHeader_(sh, HCOL.EVENT, '캘린더ID');
  ensureHeader_(sh, HCOL.REC_EVENT, '채용캘린더ID');
  const day = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const hm = d => `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
  const key = v => (v instanceof Date ? ymd(v) : '');

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  let changed = false;
  try {
    const last = sh.getLastRow();
    const rows = last >= 2 ? sh.getRange(2, 1, last - 1, HCOL.REC_EVENT).getValues() : [];
    const linked = {};
    rows.forEach((r, i) => {
      const id = String(r[HCOL.REC_EVENT - 1] || '');
      if (id && String(r[HCOL.DELETED - 1]).toUpperCase() !== 'Y') linked[id] = i;
    });
    let nextId = rows.reduce((m, r) => Math.max(m, Number(r[0]) || 0), 0);
    const seen = {};
    const add = [];
    events.forEach(e => {
      const cat = recruitCategory_(e.getTitle());
      const id = e.getId();
      if (!cat || seen[id]) return; // 반복 일정은 첫 번만
      seen[id] = true;
      const start = day(e.getStartTime());
      let end = e.isAllDayEvent() ? addDays(day(e.getEndTime()), -1) : day(e.getEndTime());
      if (end <= start) end = '';
      const memo = [e.isAllDayEvent() ? '' : `${hm(e.getStartTime())}~${hm(e.getEndTime())}`, e.getLocation()].filter(Boolean).join(' · ');
      const target = recruitTarget_(e.getTitle());
      const week = ymd(mondayOf(start));
      const i = linked[id];
      if (i == null) {
        add.push([String(++nextId), new Date(), RECRUIT_AUTHOR, week, cat, target, start, end, memo, '', '', id]);
        return;
      }
      const r = rows[i];
      const row = i + 2;
      if (r[HCOL.AUTHOR - 1] === RECRUIT_AUTHOR) {
        // 캘린더에서 가져온 항목: 캘린더가 기준
        if (r[HCOL.CATEGORY - 1] !== cat || r[HCOL.TARGET - 1] !== target || key(r[HCOL.START - 1]) !== ymd(start)
          || key(r[HCOL.END - 1]) !== key(end) || String(r[HCOL.MEMO - 1] || '') !== memo) {
          sh.getRange(row, HCOL.WEEK, 1, 6).setValues([[week, cat, target, start, end, memo]]);
          changed = true;
        }
      } else if (key(r[HCOL.START - 1]) !== ymd(start)) {
        // 웹앱에서 등록한 항목: 캘린더에서 날짜를 옮긴 것만 따라간다
        sh.getRange(row, HCOL.WEEK).setValue(week);
        sh.getRange(row, HCOL.START, 1, 2).setValues([[start, end]]);
        changed = true;
      }
    });
    if (add.length) {
      sh.getRange(last + 1, HCOL.WEEK, add.length, 1).setNumberFormat('@');
      sh.getRange(last + 1, 1, add.length, HCOL.REC_EVENT).setValues(add);
      changed = true;
    }
    // 캘린더에서 지운 일정 → 가져온 항목도 지운다 (맞추는 기간 안의 것만)
    rows.forEach((r, i) => {
      const id = String(r[HCOL.REC_EVENT - 1] || '');
      const st = r[HCOL.START - 1];
      if (!id || seen[id] || r[HCOL.AUTHOR - 1] !== RECRUIT_AUTHOR || String(r[HCOL.DELETED - 1]).toUpperCase() === 'Y') return;
      if (!(st instanceof Date) || st < from || st >= to) return;
      sh.getRange(i + 2, HCOL.DELETED).setValue('Y');
      changed = true;
    });
  } finally {
    lock.releaseLock();
  }
  if (changed) refreshHr_(ctx);
  return changed;
}

function apiDeleteHr(id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'weekly');
  const sh = ctx.ss.getSheetByName(HR.SHEET);
  if (!sh || sh.getLastRow() < 2) throw new Error('항목을 찾을 수 없습니다.');
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, HCOL.AUTHOR).getValues();
  const idx = rows.findIndex(r => String(r[0]) === String(id));
  if (idx < 0) throw new Error('항목을 찾을 수 없습니다.');
  if (!canEditHrRow_(me, rows[idx][2])) throw new Error('본인이 등록한 항목만 삭제할 수 있습니다.');
  sh.getRange(idx + 2, HCOL.DELETED).setValue('Y');
  // 팀 캘린더 일정도 지운다
  const eventId = sh.getLastColumn() >= HCOL.EVENT ? String(sh.getRange(idx + 2, HCOL.EVENT).getValue() || '') : '';
  if (eventId) {
    try {
      const ev = getTeamCalendar_().getEventById(eventId);
      if (ev) { ev.deleteEvent(); bumpCalendar_(); }
    } catch (e) { /* 캘린더 권한이 없으면 인사현황에서만 지운다 */ }
  }
  const recId = sh.getLastColumn() >= HCOL.REC_EVENT ? String(sh.getRange(idx + 2, HCOL.REC_EVENT).getValue() || '') : '';
  if (recId && recruitCalendarId_(ctx.cfg)) {
    try {
      const ev = getRecruitCalendar_(ctx).getEventById(recId);
      if (ev) { ev.deleteEvent(); bumpCalendar_(); }
    } catch (e) { /* 권한 없음 */ }
  }
  return refreshHr_(ctx);
}

function refreshHr_(ctx) {
  const hr = readHr_(ctx);
  putCache_(CK.hr(ctx), hr);
  return { hr: hr };
}

/** 주간보고 시트 아래에 인사현황 표를 붙인다. 표가 끝난 다음 행 번호를 돌려준다. */
function writeHrTable_(sh, ctx, top, table) {
  table = table || hrTable_(readHr_(ctx));
  sh.getRange(top, 1).setValue('■ 인사 현황').setFontSize(12).setFontWeight('bold');
  const h = top + 1;
  sh.getRange(h, 1, 2, 6).setValues([
    ['구 분', ctx.thisWeek.label, '', ctx.nextWeek.label, '', '비고'],
    ['', '대상 및 내용', '일자/기간', '대상 및 내용', '발령일', ''],
  ]);
  ['A', 'F'].forEach(c => sh.getRange(`${c}${h}:${c}${h + 1}`).merge());
  sh.getRange(h, 2, 1, 2).merge();
  sh.getRange(h, 4, 1, 2).merge();
  sh.getRange(h, 1, 2, 6).setBackground('#efefef').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  const values = [];
  const blocks = [];
  let row = h + 2;
  table.forEach(t => {
    t.rows.forEach((r, i) => values.push([
      i === 0 ? t.category : '',
      r.a ? r.a.target : '', r.a ? r.a.when : '',
      r.b ? r.b.target : '', r.b ? r.b.when : '',
      [r.a && r.a.memo, r.b && r.b.memo].filter(Boolean).join(' / '),
    ]));
    blocks.push({ row: row, n: t.rows.length });
    row += t.rows.length;
  });
  if (!values.length) {
    values.push(['', '이번 주·다음 주에 등록된 인사현황이 없습니다.', '', '', '', '']);
    row++;
  }
  sh.getRange(h + 2, 1, values.length, 6).setNumberFormat('@').setValues(values)
    .setVerticalAlignment('middle').setWrap(true);
  [3, 5].forEach(c => sh.getRange(h + 2, c, values.length, 1).setHorizontalAlignment('center'));
  blocks.forEach(b => sh.getRange(b.row, 1, b.n, 1).merge().setHorizontalAlignment('center').setFontWeight('bold'));
  sh.getRange(h, 1, 2 + values.length, 6)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  blocks.forEach(b => sh.getRange(b.row, 1, b.n, 6)
    .setBorder(true, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM));
  return row;
}

/* ---------- 채용 담당자 안내 메일 (매시간 자동 실행에서) ----------
 *  - 입사: 입사일 + 70일이 되는 날 → 수습평가 안내
 *  - 인터뷰: 인터뷰 전날 → 내일 인터뷰 안내
 * 받는 사람: [팀원] 시트 "채용담당(Y)" — 알림메일(회사 메일) 칸이 있으면 그 주소로.
 * 같은 항목·같은 날짜로는 한 번만 보낸다 (날짜를 고치면 새 날짜 기준으로 다시 보낸다).
 */
const HR_MAIL = {
  PROBATION_DAYS: 70,
  HOUR: 9,      // 오전 9시 이후 첫 자동 실행 때 보낸다
  LATE_DAYS: 3, // 그날 못 보냈으면 3일 안에는 보낸다 (기능을 켜기 전 지난 입사자에게 한꺼번에 가지 않도록)
};

function hrMailCheck_(ctx, now) {
  now = now || new Date();
  if (now.getHours() < HR_MAIL.HOUR) return [];
  const to = ctx.members.filter(m => m.recruit).map(mailTo_).filter(Boolean);
  const sh = ctx.ss.getSheetByName(HR.SHEET);
  if (!to.length || !sh || sh.getLastRow() < 2) return [];
  const today = startOfDay_(now);
  const props = PropertiesService.getDocumentProperties();
  const team = ctx.cfg[CFG.TEAM] || '';
  const url = ScriptApp.getService().getUrl() || ctx.ss.getUrl();
  const md = d => `${d.getMonth() + 1}/${d.getDate()}(${DAY_KO[d.getDay()]})`;
  const sent = [];

  sh.getRange(2, 1, sh.getLastRow() - 1, HCOL.DELETED).getValues().forEach(r => {
    const id = String(r[0]).trim();
    const cat = String(r[HCOL.CATEGORY - 1]).trim();
    const start = r[HCOL.START - 1];
    if (!id || String(r[HCOL.DELETED - 1]).toUpperCase() === 'Y' || !(start instanceof Date)) return;
    const target = String(r[HCOL.TARGET - 1]).trim();
    const memo = String(r[HCOL.MEMO - 1] || '').trim();
    const day = startOfDay_(start);

    let due = null, mail = null;
    if (cat === '입사') {
      due = addDays(day, HR_MAIL.PROBATION_DAYS);
      mail = {
        subject: `[${team}] 수습평가 안내 - ${target} (입사 ${HR_MAIL.PROBATION_DAYS}일)`,
        htmlBody: `채용 담당자님, 아래 입사자가 입사 ${HR_MAIL.PROBATION_DAYS}일이 되었습니다. 수습평가 진행을 준비해 주세요.<br><br>` +
          `· 대상: <b>${target}</b><br>· 입사일: ${md(day)}<br>· 입사 ${HR_MAIL.PROBATION_DAYS}일: ${md(due)}<br>` +
          (memo ? `· 비고: ${memo}<br>` : '') + `<br><a href="${url}">업무관리 열기</a>`,
      };
    } else if (cat === '인터뷰') {
      due = addDays(day, -1);
      mail = {
        subject: `[${team}] 내일 인터뷰 안내 - ${target} (${md(day)})`,
        htmlBody: `채용 담당자님, 내일 인터뷰 일정이 있습니다.<br><br>` +
          `· 내용: <b>${target}</b><br>· 일자: ${hrPeriod_(start, r[HCOL.END - 1])}<br>` +
          (memo ? `· 비고: ${memo}<br>` : '') + `<br><a href="${url}">업무관리 열기</a>`,
      };
    } else return;

    if (today < due || today > addDays(due, cat === '인터뷰' ? 0 : HR_MAIL.LATE_DAYS)) return; // 인터뷰는 전날에만
    const key = `HRMAIL:${cat}:${id}:${ymd(day)}`;
    if (props.getProperty(key)) return;
    try {
      MailApp.sendEmail(Object.assign({ to: to.join(',') }, mail));
      props.setProperty(key, ymd(today));
      sent.push(`${cat} ${target}`);
    } catch (e) { /* 메일 한도 등 → 다음 실행 때 다시 */ }
  });
  return sent;
}
