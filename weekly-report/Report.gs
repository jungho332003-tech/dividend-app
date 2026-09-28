/**
 * 주간보고 시트 생성 (보고 양식 그대로 조립)
 *
 *  이 름 | 금주 주요업무 | 기한 | 차주 주요업무 | 기한 | 휴무계획/특이사항 | 팀장 코멘트
 */

function generateReport() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    buildReport_(getContext());
  } finally {
    lock.releaseLock();
  }
}

function buildReport_(ctx) {
  const ss = ctx.ss;
  const sh = ss.getSheetByName(SHEET.REPORT) || ss.insertSheet(SHEET.REPORT, 0);
  const comments = readLeaderComments_(sh, ctx.thisWeek.key);

  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart().clear();

  // 제목
  sh.getRange('A1').setValue(`■ ${ctx.cfg[CFG.TEAM] || ''} 주간업무 ${deadlineText(ctx)}`)
    .setFontSize(13).setFontWeight('bold');
  sh.getRange('A2')
    .setValue(`※ 각자 [${SHEET.INPUT_PREFIX}이름] 시트에 작성하면 자동 반영됩니다.  이름 칸이 붉으면 작성완료 미체크  |  갱신 ${fmt(new Date(), 'M/dd HH:mm')}`)
    .setFontColor('#b45309');

  // 헤더 (2줄)
  const h = REPORT.HEADER_ROW;
  sh.getRange(h, 1, 2, REPORT.COLS).setValues([
    ['이 름', ctx.thisWeek.label, '', ctx.nextWeek.label, '', '휴무계획/특이사항', '팀장 코멘트'],
    ['', '금주 주요업무', '기한', '차주 주요업무', '기한', '', ''],
  ]);
  ['A', 'F', 'G'].forEach(c => sh.getRange(`${c}${h}:${c}${h + 1}`).merge());
  sh.getRange(h, 2, 1, 2).merge();
  sh.getRange(h, 4, 1, 2).merge();
  sh.getRange(h, 1, 2, REPORT.COLS)
    .setBackground('#efefef').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // 본문
  const rows = [];
  const blocks = [];
  let row = REPORT.FIRST_ROW;
  ctx.members.filter(m => m.write).forEach(m => {
    const input = readInput(ss, m.name);
    const n = Math.max(REPORT.MIN_BLOCK_ROWS, input.thisWeek.length, input.nextWeek.length);
    for (let i = 0; i < n; i++) {
      const a = input.thisWeek[i];
      const b = input.nextWeek[i];
      rows.push([
        i === 0 ? m.name : '',
        a ? a.task : '', a ? formatDue(a.due) : '',
        b ? b.task : '', b ? formatDue(b.due) : '',
        i === 0 ? input.note : '',
        i === 0 ? (comments[m.name] || '') : '',
      ]);
    }
    blocks.push({ row: row, n: n, done: input.done });
    row += n;
  });

  if (rows.length) {
    const body = sh.getRange(REPORT.FIRST_ROW, 1, rows.length, REPORT.COLS);
    body.setNumberFormat('@').setValues(rows).setVerticalAlignment('middle').setWrap(true);
    [3, 5].forEach(c => sh.getRange(REPORT.FIRST_ROW, c, rows.length, 1).setHorizontalAlignment('center'));
    sh.getRange(REPORT.FIRST_ROW, REPORT.COMMENT_COL, rows.length, 1).setBackground('#fffbea');

    blocks.forEach(b => {
      [1, 6, REPORT.COMMENT_COL].forEach(c => sh.getRange(b.row, c, b.n, 1).merge());
      sh.getRange(b.row, 1, b.n, 1).setHorizontalAlignment('center').setBackground(b.done ? null : '#fde2e2');
    });
  }

  const table = sh.getRange(h, 1, 2 + rows.length, REPORT.COLS);
  table.setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
  blocks.forEach(b => sh.getRange(b.row, 1, b.n, REPORT.COLS)
    .setBorder(true, null, true, null, null, null, '#555555', SpreadsheetApp.BorderStyle.SOLID_MEDIUM));

  [110, 330, 95, 330, 95, 170, 240].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(h + 1);

  PropertiesService.getDocumentProperties().setProperty(PROP.REPORT_WEEK, ctx.thisWeek.key);
}

/** 같은 주차 보고서를 다시 만들 때 팀장 코멘트를 유지하기 위해 먼저 읽어둔다 */
function readLeaderComments_(sh, weekKey) {
  const map = {};
  if (PropertiesService.getDocumentProperties().getProperty(PROP.REPORT_WEEK) !== weekKey) return map;
  const last = sh.getLastRow();
  if (last < REPORT.FIRST_ROW) return map;
  sh.getRange(REPORT.FIRST_ROW, 1, last - REPORT.FIRST_ROW + 1, REPORT.COLS).getValues().forEach(r => {
    if (r[0] && r[REPORT.COMMENT_COL - 1]) map[String(r[0]).trim()] = r[REPORT.COMMENT_COL - 1];
  });
  return map;
}
