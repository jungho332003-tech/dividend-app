/**
 * 인사현황 (주간보고 아래 표)
 *
 *  구 분 | 금주: 대상 및 내용 | 일자/기간 | 차주: 대상 및 내용 | 발령일 | 비고
 *
 * [인사현황]  번호 | 등록일시 | 등록자 | 주차(월요일) | 구분 | 대상 및 내용 | 시작일 | 종료일 | 비고 | 삭제(Y)
 * 주차는 그 주 월요일 날짜로 저장한다 → 다음 주가 되면 "차주"였던 항목이 자동으로 "금주"로 보인다.
 */

const HR = {
  SHEET: '인사현황',
  CATEGORIES: ['입사', '휴직', '복직', '퇴사', '인터뷰', '기타'],
  MIN_ROWS: 3,
};

const HCOL = { ID: 1, DATE: 2, AUTHOR: 3, WEEK: 4, CATEGORY: 5, TARGET: 6, START: 7, END: 8, MEMO: 9, DELETED: 10 };

function setupHrSheet_(ss) {
  if (ss.getSheetByName(HR.SHEET)) return;
  const sh = ss.insertSheet(HR.SHEET);
  sh.getRange(1, 1, 1, HCOL.DELETED).setValues([['번호', '등록일시', '등록자', '주차(월요일)', '구분', '대상 및 내용', '시작일', '종료일', '비고', '삭제(Y)']]);
  styleHeader_(sh.getRange(1, 1, 1, HCOL.DELETED));
  sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange('D2:D').setNumberFormat('@');
  sh.getRange('G2:H').setNumberFormat('yyyy-mm-dd (ddd)');
  sh.getRange('E2:E').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(HR.CATEGORIES).build());
  [60, 130, 80, 100, 70, 280, 120, 120, 220, 70].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);
}

/** 이번 주·다음 주 항목만 읽는다 */
function readHr_(ctx) {
  const sh = ctx.ss.getSheetByName(HR.SHEET);
  const out = { thisWeek: [], nextWeek: [] };
  if (!sh || sh.getLastRow() < 2) return out;
  const thisKey = ctx.thisWeek.key, nextKey = ctx.nextWeek.key;
  sh.getRange(2, 1, sh.getLastRow() - 1, HCOL.DELETED).getValues().forEach(r => {
    if (!String(r[0]) || String(r[HCOL.DELETED - 1]).toUpperCase() === 'Y') return;
    const week = r[HCOL.WEEK - 1] instanceof Date ? ymd(r[HCOL.WEEK - 1]) : String(r[HCOL.WEEK - 1]).trim();
    const list = week === thisKey ? out.thisWeek : week === nextKey ? out.nextWeek : null;
    if (!list) return;
    const start = r[HCOL.START - 1], end = r[HCOL.END - 1];
    list.push({
      id: String(r[0]),
      author: String(r[HCOL.AUTHOR - 1]),
      category: String(r[HCOL.CATEGORY - 1]),
      target: String(r[HCOL.TARGET - 1]),
      start: start instanceof Date ? ymd(start) : '',
      end: end instanceof Date ? ymd(end) : '',
      when: hrPeriod_(start, end),
      memo: String(r[HCOL.MEMO - 1] || ''),
    });
  });
  const byDate = (a, b) => (a.start || '9').localeCompare(b.start || '9') || Number(a.id) - Number(b.id);
  out.thisWeek.sort(byDate);
  out.nextWeek.sort(byDate);
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

/** 구분별로 금주·차주를 나란히 놓은 표 줄: [{category, rows:[{a, b}]}] */
function hrTable_(hr) {
  return HR.CATEGORIES.map(cat => {
    const a = hr.thisWeek.filter(x => x.category === cat);
    const b = hr.nextWeek.filter(x => x.category === cat);
    const n = Math.max(HR.MIN_ROWS, a.length, b.length);
    const rows = [];
    for (let i = 0; i < n; i++) rows.push({ a: a[i] || null, b: b[i] || null });
    return { category: cat, rows: rows };
  });
}

/** 등록·수정. entry: { id?, week:'this'|'next', category, target, start, end, memo } */
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
  if (start && !(start instanceof Date)) throw new Error('일자 형식이 올바르지 않습니다.');
  if (end && !(end instanceof Date)) throw new Error('종료일 형식이 올바르지 않습니다.');
  if (start instanceof Date && end instanceof Date && end < start) throw new Error('종료일이 시작일보다 빠릅니다.');
  const week = entry.week === 'next' ? ctx.nextWeek.key : ctx.thisWeek.key;
  const memo = String(entry.memo || '').trim();

  const sh = ctx.ss.getSheetByName(HR.SHEET) || (setupHrSheet_(ctx.ss), ctx.ss.getSheetByName(HR.SHEET));
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const last = sh.getLastRow();
    const ids = last >= 2 ? sh.getRange(2, 1, last - 1, HCOL.AUTHOR).getValues() : [];
    if (entry.id) {
      const idx = ids.findIndex(r => String(r[0]) === String(entry.id));
      if (idx < 0) throw new Error('항목을 찾을 수 없습니다.');
      if (ids[idx][2] !== me.name && !me.isLeader && !me.isAdmin) throw new Error('본인이 등록한 항목만 고칠 수 있습니다.');
      sh.getRange(idx + 2, HCOL.WEEK, 1, 6).setValues([[week, category, target, start, end, memo]]);
    } else {
      const id = ids.reduce((m, r) => Math.max(m, Number(r[0]) || 0), 0) + 1;
      sh.getRange(last + 1, HCOL.WEEK).setNumberFormat('@');
      sh.getRange(last + 1, 1, 1, HCOL.DELETED).setValues([[id, new Date(), me.name, week, category, target, start, end, memo, '']]);
    }
  } finally {
    lock.releaseLock();
  }
  return refreshHr_(ctx);
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
  return refreshHr_(ctx);
}

function refreshHr_(ctx) {
  const hr = readHr_(ctx);
  putCache_(CK.hr(ctx), hr);
  return { hr: hr };
}

/** 주간보고 시트 아래에 인사현황 표를 붙인다. 표가 끝난 다음 행 번호를 돌려준다. */
function writeHrTable_(sh, ctx, top) {
  const table = hrTable_(readHr_(ctx));
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
  sh.getRange(h + 2, 1, values.length, 6).setNumberFormat('@').setValues(values)
    .setVerticalAlignment('middle').setWrap(true);
  [3, 5].forEach(c => sh.getRange(h + 2, c, values.length, 1).setHorizontalAlignment('center'));
  blocks.forEach(b => sh.getRange(b.row, 1, b.n, 1).merge().setHorizontalAlignment('center').setFontWeight('bold'));
  sh.getRange(h, 1, 2 + values.length, 6)
    .setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
  blocks.forEach(b => sh.getRange(b.row, 1, b.n, 6)
    .setBorder(true, null, true, null, null, null, '#555555', SpreadsheetApp.BorderStyle.SOLID_MEDIUM));
  return row;
}
