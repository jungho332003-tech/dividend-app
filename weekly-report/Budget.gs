/**
 * 예산전용 신청 (결재 없이 신청 내역만 기록)
 *
 * [예산과목]  과목명 | 편성액 | 전입 | 전출 | 현재예산
 * [예산전용]  신청번호 | 신청일시 | 신청자 | 전출과목 | 전입과목 | 금액 | 사유 | 취소(Y)
 *
 * 신청하면 바로 예산에 반영된다. 잘못 신청한 건은 신청자 본인이나 팀장·관리자가 취소한다.
 */

const BUDGET = {
  SHEET: '예산전용',
  ACCOUNTS: '예산과목',
  ACCOUNT_ROWS: 50,
};

const BCOL = { ID: 1, DATE: 2, REQUESTER: 3, FROM: 4, TO: 5, AMOUNT: 6, REASON: 7, CANCELED: 8, FILES: 9 };

function setupBudgetSheets_(ss) {
  if (!ss.getSheetByName(BUDGET.SHEET)) {
    const sh = ss.insertSheet(BUDGET.SHEET);
    sh.getRange(1, 1, 1, 8).setValues([['신청번호', '신청일시', '신청자', '전출과목(감액)', '전입과목(증액)', '금액', '사유', '취소(Y)']]);
    styleHeader_(sh.getRange(1, 1, 1, 8));
    sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('F2:F').setNumberFormat('#,##0');
    sh.getRange('G2:G').setWrap(true);
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$H2="Y"')
        .setFontColor('#999999').setStrikethrough(true).setRanges([sh.getRange('A2:H')]).build(),
    ]);
    [110, 130, 80, 130, 130, 110, 320, 70].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
    sh.getRange('H1').setNote('Y면 취소된 건으로 예산에 반영되지 않습니다.');
  }

  ensureHeader_(ss.getSheetByName(BUDGET.SHEET), BCOL.FILES, '증빙 첨부');

  if (!ss.getSheetByName(BUDGET.ACCOUNTS)) {
    const sh = ss.insertSheet(BUDGET.ACCOUNTS);
    sh.getRange(1, 1, 1, 5).setValues([['과목명', '편성액', '전입', '전출', '현재예산']]);
    styleHeader_(sh.getRange(1, 1, 1, 5));

    const samples = ['교육훈련비', '복리후생비', '채용비', '행사비', '회의비', '도서인쇄비'];
    sh.getRange(2, 1, samples.length, 2).setValues(samples.map(n => [n, 0]));

    const b = `'${BUDGET.SHEET}'!`;
    const formulas = [];
    for (let r = 2; r < 2 + BUDGET.ACCOUNT_ROWS; r++) {
      formulas.push([
        `=IF($A${r}="","",SUMIFS(${b}$F:$F,${b}$E:$E,$A${r},${b}$H:$H,"<>Y"))`,
        `=IF($A${r}="","",SUMIFS(${b}$F:$F,${b}$D:$D,$A${r},${b}$H:$H,"<>Y"))`,
        `=IF($A${r}="","",N($B${r})+C${r}-D${r})`,
      ]);
    }
    sh.getRange(2, 3, BUDGET.ACCOUNT_ROWS, 3).setFormulas(formulas);
    sh.getRange(2, 2, BUDGET.ACCOUNT_ROWS, 4).setNumberFormat('#,##0');
    sh.getRange(2, 5, BUDGET.ACCOUNT_ROWS, 1).setFontWeight('bold');
    sh.setColumnWidth(1, 150);
    sh.setFrozenRows(1);
    sh.getRange('B1').setNote('연간(또는 기간) 편성 예산을 입력하세요.');
  }
}

/* ---------- 신청 폼 (스프레드시트 메뉴) ---------- */

function openBudgetForm() {
  const html = HtmlService.createHtmlOutputFromFile('BudgetForm').setWidth(480).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '💰 예산전용 신청');
}

/** 폼 초기값: 팀원 목록, 로그인 사용자, 과목별 현재예산 */
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
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({
      name: String(r[0]).trim(),
      budget: Number(r[1]) || 0,
      inAmt: Number(r[2]) || 0,
      outAmt: Number(r[3]) || 0,
      current: Number(r[4]) || 0,
    }));
}

/** 신청 → [예산전용] 시트에 한 줄 추가 + 팀장에게 알림 메일. 신청번호 반환 */
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
    if (amount > account.current) {
      throw new Error(`현재예산을 초과합니다. (${from} 현재예산: ${account.current.toLocaleString()}원)`);
    }

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(BUDGET.SHEET);
    const now = new Date();
    const id = nextBudgetId_(sh, now);
    sh.getRange(sh.getLastRow() + 1, 1, 1, BCOL.FILES).setValues([[id, now, requester, from, to, amount, reason, '', filesToCell_(form.files)]]);
    SpreadsheetApp.flush();
    try { CacheService.getDocumentCache().remove(CK.budget); } catch (e) { /* 캐시 없음 */ }

    const leaders = getLeaders();
    if (leaders.length) {
      MailApp.sendEmail({
        to: leaders.map(m => m.email).join(','),
        subject: `[예산전용 신청] ${id} ${requester} - ${from} → ${to} ${amount.toLocaleString()}원`,
        htmlBody:
          `<b>${requester}</b>님이 예산전용을 신청했습니다.<br><br>` +
          `신청번호: ${id}<br>전출과목: ${from}<br>전입과목: ${to}<br>` +
          `금액: <b>${amount.toLocaleString()}원</b><br>사유: ${escapeHtml_(reason)}<br>` +
          filesFromCell_(filesToCell_(form.files)).map(f => `첨부: <a href="${f.url}">${escapeHtml_(f.name)}</a><br>`).join('') + '<br>' +
          `<a href="${sheetUrl(ss, sh)}">예산전용 시트 열기</a>`,
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

function escapeHtml_(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}
