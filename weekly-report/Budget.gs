/**
 * 예산전용 신청 / 결재
 *
 * [예산과목]  과목명 | 편성액 | 전입(승인) | 전출(승인) | 현재예산 | 결재대기 전출 | 전용가능액
 * [예산전용]  신청번호 | 신청일시 | 신청자 | 전출과목 | 전입과목 | 금액 | 사유 | 결재 | 팀장 의견 | 처리일시
 *
 * 결재 칸이 비어 있으면 "결재대기". 팀장이 승인/반려를 고르면 처리일시가 찍히고 신청자에게 메일이 간다.
 */

const BUDGET = {
  SHEET: '예산전용',
  ACCOUNTS: '예산과목',
  ACCOUNT_ROWS: 50,
  DECISIONS: ['승인', '반려'],
};

const BCOL = { ID: 1, DATE: 2, REQUESTER: 3, FROM: 4, TO: 5, AMOUNT: 6, REASON: 7, STATUS: 8, OPINION: 9, DONE_AT: 10 };

function setupBudgetSheets_(ss) {
  if (!ss.getSheetByName(BUDGET.SHEET)) {
    const sh = ss.insertSheet(BUDGET.SHEET);
    sh.getRange(1, 1, 1, 10).setValues([['신청번호', '신청일시', '신청자', '전출과목(감액)', '전입과목(증액)', '금액', '사유', '결재', '팀장 의견', '처리일시']]);
    styleHeader_(sh.getRange(1, 1, 1, 10));
    sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('J2:J').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('F2:F').setNumberFormat('#,##0');
    sh.getRange('H2:H').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(BUDGET.DECISIONS).setAllowInvalid(false).build());
    sh.getRange('G2:G').setWrap(true);
    sh.getRange('I2:I').setWrap(true);

    const all = sh.getRange('A2:J');
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A2<>"",$H2="")').setBackground('#fff4d6').setRanges([all]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$H2="승인"').setBackground('#e3f4e1').setRanges([all]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$H2="반려"').setBackground('#fde2e2').setRanges([all]).build(),
    ]);
    [110, 130, 80, 130, 130, 110, 280, 70, 220, 130].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
    sh.getRange('H1').setNote('비어 있으면 결재대기. 팀장이 승인/반려를 선택하면 신청자에게 메일이 갑니다.');
  }

  if (!ss.getSheetByName(BUDGET.ACCOUNTS)) {
    const sh = ss.insertSheet(BUDGET.ACCOUNTS);
    sh.getRange(1, 1, 1, 7).setValues([['과목명', '편성액', '전입(승인)', '전출(승인)', '현재예산', '결재대기 전출', '전용가능액']]);
    styleHeader_(sh.getRange(1, 1, 1, 7));

    const samples = ['교육훈련비', '복리후생비', '채용비', '행사비', '회의비', '도서인쇄비'];
    sh.getRange(2, 1, samples.length, 2).setValues(samples.map(n => [n, 0]));

    const b = `'${BUDGET.SHEET}'!`;
    const formulas = [];
    for (let r = 2; r < 2 + BUDGET.ACCOUNT_ROWS; r++) {
      formulas.push([
        `=IF($A${r}="","",SUMIFS(${b}$F:$F,${b}$E:$E,$A${r},${b}$H:$H,"승인"))`,
        `=IF($A${r}="","",SUMIFS(${b}$F:$F,${b}$D:$D,$A${r},${b}$H:$H,"승인"))`,
        `=IF($A${r}="","",N($B${r})+C${r}-D${r})`,
        `=IF($A${r}="","",SUMIFS(${b}$F:$F,${b}$D:$D,$A${r},${b}$H:$H,""))`,
        `=IF($A${r}="","",E${r}-F${r})`,
      ]);
    }
    sh.getRange(2, 3, BUDGET.ACCOUNT_ROWS, 5).setFormulas(formulas);
    sh.getRange(2, 2, BUDGET.ACCOUNT_ROWS, 6).setNumberFormat('#,##0');
    sh.getRange(2, 7, BUDGET.ACCOUNT_ROWS, 1).setFontWeight('bold');
    sh.setColumnWidth(1, 150);
    sh.setFrozenRows(1);
    sh.getRange('B1').setNote('연간(또는 기간) 편성 예산을 입력하세요.');
  }
}

/* ---------- 신청 폼 ---------- */

function openBudgetForm() {
  const html = HtmlService.createHtmlOutputFromFile('BudgetForm').setWidth(480).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '💰 예산전용 신청');
}

/** 폼 초기값: 팀원 목록, 로그인 사용자, 과목별 전용가능액 */
function getBudgetFormData() {
  const members = getMembers();
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const me = members.find(m => m.email && m.email.toLowerCase() === email);
  return {
    members: members.map(m => m.name),
    me: me ? me.name : '',
    accounts: getAccounts_(),
  };
}

function getAccounts_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(BUDGET.ACCOUNTS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({ name: String(r[0]).trim(), balance: Number(r[4]) || 0, available: Number(r[6]) || 0 }));
}

/** 폼 제출 → [예산전용] 시트에 한 줄 추가 + 팀장에게 메일. 신청번호 반환 */
function submitBudgetTransfer(form) {
  const requester = String(form.requester || '').trim();
  const from = String(form.from || '').trim();
  const to = String(form.to || '').trim();
  const reason = String(form.reason || '').trim();
  const amount = Math.round(Number(String(form.amount || '').replace(/,/g, '')));

  if (!requester || !from || !to || !reason) throw new Error('모든 항목을 입력해 주세요.');
  if (from === to) throw new Error('전출과목과 전입과목이 같습니다.');
  if (!(amount > 0)) throw new Error('금액을 올바르게 입력해 주세요.');

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const account = getAccounts_().find(a => a.name === from);
    if (!account) throw new Error(`[${from}] 과목을 찾을 수 없습니다.`);
    if (amount > account.available) {
      throw new Error(`전용가능액을 초과합니다. (${from} 전용가능액: ${account.available.toLocaleString()}원)`);
    }

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(BUDGET.SHEET);
    const now = new Date();
    const id = nextBudgetId_(sh, now);
    const row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, 7).setValues([[id, now, requester, from, to, amount, reason]]);
    SpreadsheetApp.flush();

    const leaders = getLeaders();
    if (leaders.length) {
      MailApp.sendEmail({
        to: leaders.map(m => m.email).join(','),
        subject: `[예산전용 신청] ${id} ${requester} - ${from} → ${to} ${amount.toLocaleString()}원`,
        htmlBody:
          `<b>${requester}</b>님이 예산전용을 신청했습니다.<br><br>` +
          `신청번호: ${id}<br>전출과목: ${from}<br>전입과목: ${to}<br>` +
          `금액: <b>${amount.toLocaleString()}원</b><br>사유: ${escapeHtml_(reason)}<br><br>` +
          `<a href="${sheetUrl(ss, sh)}&range=H${row}">결재하러 가기</a> (결재 칸에서 승인/반려 선택)`,
      });
    }
    return id;
  } finally {
    lock.releaseLock();
  }
}

/** BT-2026-001 형식 */
function nextBudgetId_(sh, now) {
  const prefix = `BT-${fmt(now, 'yyyy')}-`;
  let max = 0;
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, BCOL.ID, sh.getLastRow() - 1, 1).getValues().forEach(([v]) => {
      const s = String(v);
      if (s.indexOf(prefix) === 0) max = Math.max(max, Number(s.slice(prefix.length)) || 0);
    });
  }
  return prefix + String(max + 1).padStart(3, '0');
}

/* ---------- 결재 ---------- */

/** 설치형 onEdit 트리거 (setupTriggers에서 등록). 결재 칸 변경 시 처리일시 기록 + 신청자 메일 */
function onBudgetEdit(e) {
  const range = e.range;
  const sh = range.getSheet();
  if (sh.getName() !== BUDGET.SHEET) return;
  if (range.getColumn() !== BCOL.STATUS || range.getNumColumns() !== 1 || range.getNumRows() !== 1 || range.getRow() < 2) return;

  const row = range.getRow();
  const status = String(range.getValue() || '').trim();
  if (BUDGET.DECISIONS.indexOf(status) < 0) {
    sh.getRange(row, BCOL.DONE_AT).clearContent();
    return;
  }
  sh.getRange(row, BCOL.DONE_AT).setValue(new Date());
  notifyBudgetDecision_(sh, row);
}

/** 결재 결과를 신청자에게 메일로 알린다 */
function notifyBudgetDecision_(sh, row) {
  const v = sh.getRange(row, 1, 1, 10).getValues()[0];
  const status = String(v[BCOL.STATUS - 1]);
  const requester = getMembers().find(m => m.name === String(v[BCOL.REQUESTER - 1]).trim());
  if (!requester || !requester.email) return;

  const amount = Number(v[BCOL.AMOUNT - 1]) || 0;
  const opinion = String(v[BCOL.OPINION - 1] || '').trim();
  MailApp.sendEmail({
    to: requester.email,
    subject: `[예산전용 ${status}] ${v[BCOL.ID - 1]} ${v[BCOL.FROM - 1]} → ${v[BCOL.TO - 1]} ${amount.toLocaleString()}원`,
    htmlBody:
      `신청하신 예산전용이 <b>${status}</b>되었습니다.<br><br>` +
      `신청번호: ${v[BCOL.ID - 1]}<br>전출과목: ${v[BCOL.FROM - 1]}<br>전입과목: ${v[BCOL.TO - 1]}<br>` +
      `금액: ${amount.toLocaleString()}원<br>` +
      (opinion ? `팀장 의견: ${escapeHtml_(opinion)}<br>` : '') +
      `<br><a href="${sheetUrl(SpreadsheetApp.getActive(), sh)}">예산전용 시트 열기</a>`,
  });
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}
