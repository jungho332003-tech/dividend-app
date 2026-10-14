/**
 * 주간보고 시트 생성 (보고 양식 그대로 조립)
 *
 *  이 름 | 금주 주요업무 | 기한 | 차주 주요업무 | 기한 | 휴무계획/특이사항 | 팀장 코멘트
 *  그 아래에 인사현황 표 (Personnel.gs)
 */

/** 메뉴 [주간보고 새로고침]: 내용이 같아도 다시 만든다 */
function generateReport() {
  withReportLock_(ctx => buildReport_(ctx));
}

/** 매시간 자동 새로고침: 내용이 바뀌었을 때만 다시 그린다 (시트 쓰기 수십 번을 아낀다) */
function refreshReportIfChanged_() {
  withReportLock_(ctx => buildReport_(ctx, true));
}

function withReportLock_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    fn(getContext());
  } finally {
    lock.releaseLock();
  }
}

/** 보고서를 만든다. onlyIfChanged 면 지난번과 내용이 같을 때 건너뛰고 false 를 돌려준다. */
function buildReport_(ctx, onlyIfChanged) {
  const ss = ctx.ss;
  const existing = ss.getSheetByName(SHEET.REPORT);
  const sh = existing || ss.insertSheet(SHEET.REPORT, 0);
  const comments = readLeaderComments_(sh, ctx.thisWeek.key);
  const title = `■ ${ctx.cfg[CFG.TEAM] || ''} 주간업무 ${deadlineText(ctx)}`;

  // 본문 (쓰기 전에 먼저 모은다)
  const rows = [];
  const blocks = [];
  const bold = []; // 팀장이 굵게 표시한 업무 칸 (A1)
  const marks = marksOf_(ctx);
  let row = REPORT.FIRST_ROW;
  ctx.members.filter(m => m.write).forEach(m => {
    const input = readInput(ss, m.name);
    const mk = marks[m.name] || { thisWeek: [], nextWeek: [] };
    const n = Math.max(REPORT.MIN_BLOCK_ROWS, input.thisWeek.length, input.nextWeek.length);
    for (let i = 0; i < n; i++) {
      const a = input.thisWeek[i];
      const b = input.nextWeek[i];
      if (a && mk.thisWeek.indexOf(a.task) >= 0) bold.push('B' + (row + i));
      if (b && mk.nextWeek.indexOf(b.task) >= 0) bold.push('D' + (row + i));
      rows.push([
        i === 0 ? m.name : '',
        a ? a.task : '', a ? formatDue(a.due) : '',
        b ? b.task : '', b ? formatDue(b.due) : '',
        i === 0 ? combineNote_(memberLeaves_(ctx, m.name, ctx.thisWeek.monday, 2), input.note) : '',
        i === 0 ? (comments[m.name] || '') : '',
      ]);
    }
    blocks.push({ row: row, n: n, done: input.done });
    row += n;
  });
  const hrTable = hrTable_(readHr_(ctx));

  const props = PropertiesService.getDocumentProperties();
  const digest = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,
    JSON.stringify([title, ctx.thisWeek.label, ctx.nextWeek.label, rows, blocks, bold, hrTable]), Utilities.Charset.UTF_8));
  if (onlyIfChanged && existing && props.getProperty(PROP.REPORT_WEEK) === ctx.thisWeek.key
    && props.getProperty(PROP.REPORT_HASH) === digest) return false;

  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart().clear();

  // 제목
  sh.getRange('A1').setValue(title).setFontSize(13).setFontWeight('bold');
  sh.getRange('A2').setValue(`※ 이름 칸이 붉으면 작성완료 미체크  |  갱신 ${fmt(new Date(), 'M/dd HH:mm')}`)
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

  if (rows.length) {
    const body = sh.getRange(REPORT.FIRST_ROW, 1, rows.length, REPORT.COLS);
    body.setNumberFormat('@').setValues(rows).setVerticalAlignment('middle').setWrap(true);
    [3, 5].forEach(c => sh.getRange(REPORT.FIRST_ROW, c, rows.length, 1).setHorizontalAlignment('center'));
    sh.getRange(REPORT.FIRST_ROW, REPORT.COMMENT_COL, rows.length, 1).setBackground('#fffbea');

    if (bold.length) sh.getRangeList(bold).setFontWeight('bold');
    blocks.forEach(b => {
      [1, 6, REPORT.COMMENT_COL].forEach(c => sh.getRange(b.row, c, b.n, 1).merge());
      sh.getRange(b.row, 1, b.n, 1).setHorizontalAlignment('center').setBackground(b.done ? null : '#fde2e2');
    });
  }

  const table = sh.getRange(h, 1, 2 + rows.length, REPORT.COLS);
  table.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  blocks.forEach(b => sh.getRange(b.row, 1, b.n, REPORT.COLS)
    .setBorder(true, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM));

  // 인사현황 표 (입사·휴직·복직·퇴사·인터뷰·기타)
  writeHrTable_(sh, ctx, h + 2 + rows.length + 2, hrTable);

  [110, 330, 95, 330, 95, 170, 240].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(h + 1);

  const saved = {};
  saved[PROP.REPORT_WEEK] = ctx.thisWeek.key;
  saved[PROP.REPORT_HASH] = digest;
  props.setProperties(saved);
  return true;
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
