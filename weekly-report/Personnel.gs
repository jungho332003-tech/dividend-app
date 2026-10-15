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

const HCOL = { ID: 1, DATE: 2, AUTHOR: 3, WEEK: 4, CATEGORY: 5, TARGET: 6, START: 7, END: 8, MEMO: 9, DELETED: 10, EVENT: 11 };

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
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  let row, oldEvent = '';
  try {
    const last = sh.getLastRow();
    const ids = last >= 2 ? sh.getRange(2, 1, last - 1, HCOL.EVENT).getValues() : [];
    if (entry.id) {
      const idx = ids.findIndex(r => String(r[0]) === String(entry.id));
      if (idx < 0) throw new Error('항목을 찾을 수 없습니다.');
      if (ids[idx][2] !== me.name && !me.isLeader && !me.isAdmin) throw new Error('본인이 등록한 항목만 고칠 수 있습니다.');
      row = idx + 2;
      oldEvent = String(ids[idx][HCOL.EVENT - 1] || '');
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

function apiDeleteHr(id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'weekly');
  const sh = ctx.ss.getSheetByName(HR.SHEET);
  if (!sh || sh.getLastRow() < 2) throw new Error('항목을 찾을 수 없습니다.');
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, HCOL.AUTHOR).getValues();
  const idx = rows.findIndex(r => String(r[0]) === String(id));
  if (idx < 0) throw new Error('항목을 찾을 수 없습니다.');
  if (rows[idx][2] !== me.name && !me.isLeader && !me.isAdmin) throw new Error('본인이 등록한 항목만 삭제할 수 있습니다.');
  sh.getRange(idx + 2, HCOL.DELETED).setValue('Y');
  // 팀 캘린더 일정도 지운다
  const eventId = sh.getLastColumn() >= HCOL.EVENT ? String(sh.getRange(idx + 2, HCOL.EVENT).getValue() || '') : '';
  if (eventId) {
    try {
      const ev = getTeamCalendar_().getEventById(eventId);
      if (ev) { ev.deleteEvent(); bumpCalendar_(); }
    } catch (e) { /* 캘린더 권한이 없으면 인사현황에서만 지운다 */ }
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
