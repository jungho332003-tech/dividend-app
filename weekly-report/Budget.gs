/**
 * 예산전용 신청 (결재 없이 신청 내역만 기록)
 *
 * [예산과목]  코스트 | 계정과목 | 편성액 | 전입 | 전출 | 현재예산
 * [예산전용]  신청번호 | 신청일시 | 신청자 | 변경전 코스트 | 변경전 계정과목 | 변경후 코스트 | 변경후 계정과목 | 금액 | 사유 | 취소(Y) | 적용월
 *
 * 예산은 "코스트 + 계정과목" 한 쌍마다 관리한다. 변경전(감액) → 변경후(증액)로 옮긴다.
 * 적용월(yyyy-MM): 어느 달 예산을 전용하는지. 신청할 때 고르고, 내역은 월별로 볼 수 있다.
 * 신청하면 바로 예산에 반영된다. 잘못 신청한 건은 신청자 본인이나 팀장·관리자가 취소한다.
 */

const BUDGET = {
  SHEET: '예산전용',
  ACCOUNTS: '예산과목',
  ACCOUNT_ROWS: 100,
};

const BCOL = { ID: 1, DATE: 2, REQUESTER: 3, FROM_COST: 4, FROM: 5, TO_COST: 6, TO: 7, AMOUNT: 8, REASON: 9, CANCELED: 10, MONTH: 11 };
const BUDGET_HEADER = ['신청번호', '신청일시', '신청자', '변경전 코스트', '변경전 계정과목', '변경후 코스트', '변경후 계정과목', '금액', '사유', '취소(Y)', '적용월'];
const ACCOUNT_HEADER = ['코스트', '계정과목', '편성액', '전입', '전출', '현재예산'];

function setupBudgetSheets_(ss) {
  if (!ss.getSheetByName(BUDGET.SHEET)) {
    const sh = ss.insertSheet(BUDGET.SHEET);
    sh.getRange(1, 1, 1, BCOL.MONTH).setValues([BUDGET_HEADER]);
    formatBudgetSheet_(sh);
  }

  if (!ss.getSheetByName(BUDGET.ACCOUNTS)) {
    const sh = ss.insertSheet(BUDGET.ACCOUNTS);
    sh.getRange(1, 1, 1, 6).setValues([ACCOUNT_HEADER]);
    const samples = [
      ['1100 인사운영', '복리후생비'], ['1100 인사운영', '채용비'], ['1100 인사운영', '회의비'],
      ['1200 인재개발', '교육훈련비'], ['1200 인재개발', '도서인쇄비'], ['1200 인재개발', '행사비'],
    ];
    sh.getRange(2, 1, samples.length, 3).setValues(samples.map(r => r.concat([0])));
    formatAccountSheet_(sh);
  }
  migrateBudgetSheets_(ss);
}

function formatBudgetSheet_(sh) {
  styleHeader_(sh.getRange(1, 1, 1, BCOL.MONTH));
  sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange('H2:H').setNumberFormat('#,##0');
  sh.getRange('I2:I').setWrap(true);
  sh.getRange('K2:K').setNumberFormat('@');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$J2="Y"')
      .setFontColor('#999999').setStrikethrough(true).setRanges([sh.getRange('A2:K')]).build(),
  ]);
  [110, 130, 80, 120, 120, 120, 120, 110, 320, 70, 80].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);
  sh.getRange('J1').setNote('Y면 취소된 건으로 예산에 반영되지 않습니다.');
}

/** 전입·전출·현재예산 수식: 코스트와 계정과목이 모두 같은 신청만 더한다 ("="&칸 → 코스트가 빈칸이어도 빈칸끼리 맞춘다) */
function formatAccountSheet_(sh) {
  styleHeader_(sh.getRange(1, 1, 1, 6));
  const b = `'${BUDGET.SHEET}'!`;
  const formulas = [];
  for (let r = 2; r < 2 + BUDGET.ACCOUNT_ROWS; r++) {
    formulas.push([
      `=IF($B${r}="","",SUMIFS(${b}$H:$H,${b}$F:$F,"="&$A${r},${b}$G:$G,"="&$B${r},${b}$J:$J,"<>Y"))`,
      `=IF($B${r}="","",SUMIFS(${b}$H:$H,${b}$D:$D,"="&$A${r},${b}$E:$E,"="&$B${r},${b}$J:$J,"<>Y"))`,
      `=IF($B${r}="","",N($C${r})+D${r}-E${r})`,
    ]);
  }
  sh.getRange(2, 4, BUDGET.ACCOUNT_ROWS, 3).setFormulas(formulas);
  sh.getRange(2, 3, BUDGET.ACCOUNT_ROWS, 4).setNumberFormat('#,##0');
  sh.getRange(2, 6, BUDGET.ACCOUNT_ROWS, 1).setFontWeight('bold');
  sh.setColumnWidth(1, 150).setColumnWidth(2, 150);
  sh.setFrozenRows(1);
  sh.getRange('A1').setNote('코스트(코스트센터)와 계정과목 한 쌍마다 한 줄. 편성액은 연간(또는 기간) 예산을 입력하세요.');
}

/**
 * 예전 양식(과목명만, 전출·전입과목)을 코스트 + 계정과목 양식으로 바꾼다. 이미 새 양식이면 아무것도 안 한다.
 * 기존 신청은 코스트가 빈칸인 채로 옮겨지고, 예산과목의 기존 과목도 코스트 빈칸으로 남는다 (나중에 채우면 됨).
 */
function migrateBudgetSheets_(ss) {
  const req = ss.getSheetByName(BUDGET.SHEET);
  if (req && String(req.getRange(1, BCOL.FROM_COST).getValue()).trim() !== BUDGET_HEADER[BCOL.FROM_COST - 1]) {
    // 예전: 번호 | 일시 | 신청자 | 전출 | 전입 | 금액 | 사유 | 취소 | 첨부 | 적용월
    const last = req.getLastRow();
    const old = last >= 2 ? req.getRange(2, 1, last - 1, 10).getValues() : [];
    const rows = old.map(r => [r[0], r[1], r[2], '', r[3], '', r[4], r[5], r[6], r[7], r[9] instanceof Date ? fmt(r[9], 'yyyy-MM') : r[9]]);
    req.getRange(1, 1, Math.max(last, 1), Math.max(req.getLastColumn(), BCOL.MONTH)).clearContent().clearNote();
    req.getRange(1, 1, 1, BCOL.MONTH).setValues([BUDGET_HEADER]);
    if (rows.length) {
      req.getRange(2, BCOL.MONTH, rows.length, 1).setNumberFormat('@');
      req.getRange(2, 1, rows.length, BCOL.MONTH).setValues(rows);
    }
    formatBudgetSheet_(req);
  }

  const acc = ss.getSheetByName(BUDGET.ACCOUNTS);
  if (acc && String(acc.getRange(1, 1).getValue()).trim() !== ACCOUNT_HEADER[0]) {
    // 예전: 과목명 | 편성액 | 전입 | 전출 | 현재예산 → 맨 앞에 코스트 열을 넣는다
    acc.insertColumnBefore(1);
    acc.getRange(1, 1, 1, 6).setValues([ACCOUNT_HEADER]);
    formatAccountSheet_(acc);
  }
}

/* ---------- 신청 폼 (스프레드시트 메뉴) ---------- */

function openBudgetForm() {
  // 한 파일 설치본(dist/Code.gs)은 화면을 BUDGET_FORM_HTML 문자열로 갖고 있다
  const html = (typeof BUDGET_FORM_HTML === 'string'
    ? HtmlService.createHtmlOutput(BUDGET_FORM_HTML)
    : HtmlService.createHtmlOutputFromFile('BudgetForm')).setWidth(520).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '💰 예산전용 신청');
}

/** 폼 초기값: 팀원 목록, 로그인 사용자, 코스트·계정과목별 현재예산 */
function getBudgetFormData() {
  const members = getMembers();
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const me = members.find(m => m.email && m.email.toLowerCase() === email);
  return {
    members: members.map(m => m.name),
    me: me ? me.name : '',
    accounts: getAccounts_(),
    months: budgetMonths_(new Date()),
  };
}

/** 신청할 수 있는 달: 올해 1월 ~ 6개월 뒤. [{value:'2026-10', label:'2026년 10월'}], 이번 달은 current:true */
function budgetMonths_(now) {
  const out = [];
  const cur = fmt(now, 'yyyy-MM');
  const d = new Date(now.getFullYear(), 0, 1);
  const end = new Date(now.getFullYear(), now.getMonth() + 6, 1);
  while (d <= end) {
    const v = fmt(d, 'yyyy-MM');
    out.push({ value: v, label: `${d.getFullYear()}년 ${d.getMonth() + 1}월`, current: v === cur });
    d.setMonth(d.getMonth() + 1);
  }
  return out;
}

/** 'yyyy-MM' 검사. 비어 있으면 이번 달 */
function budgetMonth_(v) {
  const s = String(v || '').trim();
  if (!s) return fmt(new Date(), 'yyyy-MM');
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(s)) throw new Error('적용월을 올바르게 골라 주세요.');
  return s;
}

/** 시트의 적용월 칸 → 'yyyy-MM' (적용월이 없는 건은 신청일시 기준) */
function monthOfRow_(r) {
  const m = r[BCOL.MONTH - 1];
  if (m instanceof Date) return fmt(m, 'yyyy-MM');
  if (/^\d{4}-\d{2}$/.test(String(m).trim())) return String(m).trim();
  return r[BCOL.DATE - 1] instanceof Date ? fmt(r[BCOL.DATE - 1], 'yyyy-MM') : String(r[BCOL.DATE - 1]).slice(0, 7);
}

function monthLabel_(ym) {
  return `${Number(ym.slice(0, 4))}년 ${Number(ym.slice(5, 7))}월`;
}

/** "1100 인사운영 / 교육훈련비" (코스트가 비어 있으면 계정과목만) */
function budgetLabel_(cost, account) {
  return cost ? `${cost} / ${account}` : account;
}

function getAccounts_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(BUDGET.ACCOUNTS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues()
    .filter(r => String(r[1]).trim())
    .map(r => ({
      cost: String(r[0]).trim(),
      name: String(r[1]).trim(),
      budget: Number(r[2]) || 0,
      inAmt: Number(r[3]) || 0,
      outAmt: Number(r[4]) || 0,
      current: Number(r[5]) || 0,
    }));
}

/**
 * 신청 → [예산전용] 시트에 한 줄 추가 + 팀장에게 알림 메일. 신청번호 반환
 * form: { requester, month, fromCost, from, toCost, to, amount, reason }
 */
function submitBudgetTransfer(form) {
  const requester = String(form.requester || '').trim();
  const fromCost = String(form.fromCost || '').trim();
  const from = String(form.from || '').trim();
  const toCost = String(form.toCost || '').trim();
  const to = String(form.to || '').trim();
  const reason = String(form.reason || '').trim();
  const amount = Math.round(Number(String(form.amount || '').replace(/,/g, '')));
  const month = budgetMonth_(form.month);

  if (!requester || !from || !to || !reason) throw new Error('모든 항목을 입력해 주세요.');
  if (fromCost === toCost && from === to) throw new Error('변경전과 변경후가 같습니다.');
  if (!(amount > 0)) throw new Error('금액을 올바르게 입력해 주세요.');

  const ss = SpreadsheetApp.getActive();
  try { migrateBudgetSheets_(ss); } catch (e) {
    throw new Error('예산 시트를 새 양식(코스트 + 계정과목)으로 바꾸지 못했습니다. 관리자에게 "권한 한 번에 적용"을 요청하세요.');
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const accounts = getAccounts_();
    const fromLabel = budgetLabel_(fromCost, from), toLabel = budgetLabel_(toCost, to);
    const account = accounts.find(a => a.cost === fromCost && a.name === from);
    if (!account) throw new Error(`[${fromLabel}] 예산을 찾을 수 없습니다. [예산과목] 시트를 확인하세요.`);
    if (!accounts.some(a => a.cost === toCost && a.name === to)) throw new Error(`[${toLabel}] 예산을 찾을 수 없습니다. [예산과목] 시트를 확인하세요.`);
    if (amount > account.current) {
      throw new Error(`현재예산을 초과합니다. (${fromLabel} 현재예산: ${account.current.toLocaleString()}원)`);
    }

    const sh = ss.getSheetByName(BUDGET.SHEET);
    const now = new Date();
    const id = nextBudgetId_(sh, now);
    const row = sh.getLastRow() + 1;
    sh.getRange(row, BCOL.MONTH).setNumberFormat('@'); // '2026-10'이 날짜로 바뀌지 않게
    sh.getRange(row, 1, 1, BCOL.MONTH).setValues([[id, now, requester, fromCost, from, toCost, to, amount, reason, '', month]]);
    SpreadsheetApp.flush();
    try { CacheService.getDocumentCache().remove(CK.budget); } catch (e) { /* 캐시 없음 */ }

    const leaders = getLeaders();
    if (leaders.length) {
      MailApp.sendEmail({
        to: leaders.map(m => m.email).join(','),
        subject: `[예산전용 신청] ${monthLabel_(month)}분 ${id} ${requester} - ${fromLabel} → ${toLabel} ${amount.toLocaleString()}원`,
        htmlBody:
          `<b>${escapeHtml_(requester)}</b>님이 예산전용을 신청했습니다.<br><br>` +
          `신청번호: ${id}<br>적용월: ${monthLabel_(month)}<br>` +
          `변경전: ${escapeHtml_(fromLabel)}<br>변경후: ${escapeHtml_(toLabel)}<br>` +
          `금액: <b>${amount.toLocaleString()}원</b><br>사유: ${escapeHtml_(reason)}<br><br>` +
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
