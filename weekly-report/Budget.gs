/**
 * 예산전용 신청 (결재 없이 신청 내역만 기록)
 *
 * [예산과목]  코스트 | 계정코드 | 계정과목 | 편성액 | 전입 | 전출 | 현재예산
 * [예산전용]  신청번호 | 신청일시 | 신청자 | 변경전 코스트 | 변경전 계정코드 | 변경전 계정과목
 *             | 변경후 코스트 | 변경후 계정코드 | 변경후 계정과목 | 금액 | 사유 | 취소(Y) | 적용월
 *
 * 예산은 "코스트(코스트센터) + 계정과목" 한 쌍마다 관리한다. 변경전(감액) → 변경후(증액)로 옮긴다.
 * 계정코드는 [예산과목]에서 가져와 신청 내역에 함께 남긴다 (담당자가 ERP 등에 옮겨 적을 때 쓴다).
 * 적용월(yyyy-MM): 어느 달 예산을 전용하는지. 신청할 때 고르고, 내역은 월별로 볼 수 있다.
 * 신청하면 바로 예산에 반영된다. 잘못 신청한 건은 신청자 본인이나 팀장·관리자가 취소한다.
 */

const BUDGET = {
  SHEET: '예산전용',
  ACCOUNTS: '예산과목',
  ACCOUNT_ROWS: 100,
};

const BCOL = {
  ID: 1, DATE: 2, REQUESTER: 3,
  FROM_COST: 4, FROM_CODE: 5, FROM: 6,
  TO_COST: 7, TO_CODE: 8, TO: 9,
  AMOUNT: 10, REASON: 11, CANCELED: 12, MONTH: 13,
};
const BUDGET_HEADER = ['신청번호', '신청일시', '신청자', '변경전 코스트', '변경전 계정코드', '변경전 계정과목',
  '변경후 코스트', '변경후 계정코드', '변경후 계정과목', '금액', '사유', '취소(Y)', '적용월'];
const ACCOUNT_HEADER = ['코스트', '계정코드', '계정과목', '편성액', '전입', '전출', '현재예산'];

function setupBudgetSheets_(ss) {
  if (!ss.getSheetByName(BUDGET.SHEET)) {
    const sh = ss.insertSheet(BUDGET.SHEET);
    sh.getRange(1, 1, 1, BCOL.MONTH).setValues([BUDGET_HEADER]);
    formatBudgetSheet_(sh);
  }

  if (!ss.getSheetByName(BUDGET.ACCOUNTS)) {
    const sh = ss.insertSheet(BUDGET.ACCOUNTS);
    sh.getRange(1, 1, 1, 7).setValues([ACCOUNT_HEADER]);
    const samples = [
      ['1100', '52100', '복리후생비'], ['1100', '52300', '채용비'], ['1100', '52400', '회의비'],
      ['1200', '52500', '교육훈련비'], ['1200', '52600', '도서인쇄비'], ['1200', '52700', '행사비'],
    ];
    sh.getRange(2, 1, BUDGET.ACCOUNT_ROWS, 2).setNumberFormat('@'); // 0으로 시작하는 코드도 그대로
    sh.getRange(2, 1, samples.length, 4).setValues(samples.map(r => r.concat([0])));
    formatAccountSheet_(sh);
  }
  migrateBudgetSheets_(ss);
}

function formatBudgetSheet_(sh) {
  styleHeader_(sh.getRange(1, 1, 1, BCOL.MONTH));
  sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange('D2:I').setNumberFormat('@');
  sh.getRange('J2:J').setNumberFormat('#,##0');
  sh.getRange('K2:K').setWrap(true);
  sh.getRange('M2:M').setNumberFormat('@');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$L2="Y"')
      .setFontColor('#999999').setStrikethrough(true).setRanges([sh.getRange('A2:M')]).build(),
  ]);
  [110, 130, 80, 90, 90, 120, 90, 90, 120, 110, 300, 70, 80].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);
  sh.getRange('L1').setNote('Y면 취소된 건으로 예산에 반영되지 않습니다.');
}

/** 전입·전출·현재예산 수식: 코스트와 계정과목이 모두 같은 신청만 더한다 ("="&칸 → 코스트가 빈칸이어도 빈칸끼리 맞춘다) */
function formatAccountSheet_(sh, rows) {
  rows = rows || BUDGET.ACCOUNT_ROWS;
  styleHeader_(sh.getRange(1, 1, 1, 7));
  const b = `'${BUDGET.SHEET}'!`;
  const formulas = [];
  for (let r = 2; r < 2 + rows; r++) {
    formulas.push([
      `=IF($C${r}="","",SUMIFS(${b}$J:$J,${b}$G:$G,"="&$A${r},${b}$I:$I,"="&$C${r},${b}$L:$L,"<>Y"))`,
      `=IF($C${r}="","",SUMIFS(${b}$J:$J,${b}$D:$D,"="&$A${r},${b}$F:$F,"="&$C${r},${b}$L:$L,"<>Y"))`,
      `=IF($C${r}="","",N($D${r})+E${r}-F${r})`,
    ]);
  }
  sh.getRange(2, 5, rows, 3).setFormulas(formulas);
  sh.getRange(2, 4, rows, 4).setNumberFormat('#,##0');
  sh.getRange(2, 7, rows, 1).setFontWeight('bold');
  sh.setColumnWidth(1, 110).setColumnWidth(2, 110).setColumnWidth(3, 150);
  sh.setFrozenRows(1);
  sh.getRange('A1').setNote('코스트센터 코드. 코스트와 계정과목 한 쌍마다 한 줄씩 입력하고, 편성액은 연간(또는 기간) 예산을 입력하세요.');
}

/**
 * 예전 양식을 지금 양식(코스트 + 계정코드 + 계정과목)으로 바꾼다. 이미 지금 양식이면 아무것도 안 한다.
 *  v1: 과목명만 (전출과목 | 전입과목)
 *  v2: 코스트 + 계정과목 (계정코드 없음)
 * 계정코드는 [예산과목]에 같은 코스트·계정과목이 있으면 그 코드로 채운다.
 */
function migrateBudgetSheets_(ss) {
  const acc = ss.getSheetByName(BUDGET.ACCOUNTS);
  if (acc) {
    const a1 = String(acc.getRange(1, 1).getValue()).trim();
    const b1 = String(acc.getRange(1, 2).getValue()).trim();
    let changed = false;
    if (a1 !== '코스트') { acc.insertColumnsBefore(1, 2); changed = true; }          // v1: 과목명 | 편성액 …
    else if (b1 !== '계정코드') { acc.insertColumnBefore(2); changed = true; }      // v2: 코스트 | 계정과목 …
    if (changed) {
      acc.getRange(1, 1, 1, 7).setValues([ACCOUNT_HEADER]);
      acc.getRange(2, 1, BUDGET.ACCOUNT_ROWS, 2).setNumberFormat('@');
      formatAccountSheet_(acc);
    }
  }

  const req = ss.getSheetByName(BUDGET.SHEET);
  if (!req) return;
  const header = req.getRange(1, 1, 1, 6).getValues()[0].map(v => String(v).trim());
  if (header[BCOL.FROM_CODE - 1] === BUDGET_HEADER[BCOL.FROM_CODE - 1]) return;
  const v1 = header[3] !== '변경전 코스트';
  const last = req.getLastRow();
  const old = last >= 2 ? req.getRange(2, 1, last - 1, 11).getValues() : [];
  const codes = {};
  getAccounts_().forEach(a => { codes[a.cost + '\u0000' + a.name] = a.code; });
  const code = (cost, name) => codes[String(cost).trim() + '\u0000' + String(name).trim()] || '';
  const ym = v => (v instanceof Date ? fmt(v, 'yyyy-MM') : v);
  const rows = old.map(r => v1
    // 번호 | 일시 | 신청자 | 전출 | 전입 | 금액 | 사유 | 취소 | 첨부 | 적용월
    ? [r[0], r[1], r[2], '', code('', r[3]), r[3], '', code('', r[4]), r[4], r[5], r[6], r[7], ym(r[9])]
    // 번호 | 일시 | 신청자 | 변경전 코스트 | 변경전 계정과목 | 변경후 코스트 | 변경후 계정과목 | 금액 | 사유 | 취소 | 적용월
    : [r[0], r[1], r[2], r[3], code(r[3], r[4]), r[4], r[5], code(r[5], r[6]), r[6], r[7], r[8], r[9], ym(r[10])]);
  req.getRange(1, 1, Math.max(last, 1), Math.max(req.getLastColumn(), BCOL.MONTH)).clearContent().clearNote();
  req.getRange(1, 1, 1, BCOL.MONTH).setValues([BUDGET_HEADER]);
  if (rows.length) {
    req.getRange(2, BCOL.FROM_COST, rows.length, 6).setNumberFormat('@');
    req.getRange(2, BCOL.MONTH, rows.length, 1).setNumberFormat('@');
    req.getRange(2, 1, rows.length, BCOL.MONTH).setValues(rows);
  }
  formatBudgetSheet_(req);
}

/* ---------- 신청 폼 (스프레드시트 메뉴) ---------- */

function openBudgetForm() {
  // 한 파일 설치본(dist/Code.gs)은 화면을 BUDGET_FORM_HTML 문자열로 갖고 있다
  const html = (typeof BUDGET_FORM_HTML === 'string'
    ? HtmlService.createHtmlOutput(BUDGET_FORM_HTML)
    : HtmlService.createHtmlOutputFromFile('BudgetForm')).setWidth(520).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '💰 예산전용 신청');
}

/** 폼 초기값: 팀원 목록, 로그인 사용자, 코스트·계정코드·계정과목별 현재예산 */
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

/** "1100 / 52500 교육훈련비" (비어 있는 부분은 뺀다) */
function budgetLabel_(cost, code, account) {
  const acc = [code, account].filter(Boolean).join(' ');
  return cost ? `${cost} / ${acc}` : acc;
}

/**
 * 웹앱 [예산 설정]: 코스트·계정코드·계정과목·편성액 목록을 통째로 저장한다 (관리자·예산 담당자).
 * list: [{ cost, code, name, budget }]. 전입·전출·현재예산은 시트 수식이 계산한다.
 * 신청 내역이 있는 코스트·계정과목을 빼거나 이름을 바꾸면 그 신청이 예산에 안 잡히므로, force 없이는 막고 알려준다.
 */
function apiSaveAccounts(list, force) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'budget');
  if (!me.isAdmin && !me.budgetManager) throw new Error('예산 설정은 관리자와 예산 담당자만 할 수 있습니다.');

  const rows = [];
  const seen = {};
  (list || []).forEach((a, i) => {
    const cost = String(a.cost || '').trim(), code = String(a.code || '').trim(), name = String(a.name || '').trim();
    if (!cost && !code && !name && !Number(a.budget)) return; // 빈 줄
    if (!name) throw new Error(`${i + 1}번째 줄: 계정과목을 입력해 주세요.`);
    const key = cost + '|' + name;
    if (seen[key]) throw new Error(`같은 코스트·계정과목이 두 번 있습니다: ${budgetLabel_(cost, code, name)}`);
    seen[key] = true;
    const budget = Number(String(a.budget == null ? '' : a.budget).replace(/[^\d.-]/g, '')) || 0;
    if (budget < 0) throw new Error(`${budgetLabel_(cost, code, name)}: 편성액은 0 이상이어야 합니다.`);
    rows.push([cost, code, name, budget]);
  });

  // 지금 신청 내역에서 쓰는 코스트·계정과목이 사라지는지 확인
  if (!force) {
    const used = {};
    const req = ctx.ss.getSheetByName(BUDGET.SHEET);
    if (req && req.getLastRow() >= 2) {
      req.getRange(2, 1, req.getLastRow() - 1, BCOL.MONTH).getValues().forEach(r => {
        if (String(r[BCOL.CANCELED - 1]).toUpperCase() === 'Y') return;
        [[r[BCOL.FROM_COST - 1], r[BCOL.FROM - 1]], [r[BCOL.TO_COST - 1], r[BCOL.TO - 1]]].forEach(([c, n]) => {
          if (String(n).trim()) used[String(c).trim() + '|' + String(n).trim()] = budgetLabel_(String(c).trim(), '', String(n).trim());
        });
      });
    }
    // 지금 [예산과목]에 있던 것만 본다 (원래부터 짝이 없던 예전 신청은 상관없음)
    const before = {};
    getAccounts_().forEach(a => { before[a.cost + '|' + a.name] = true; });
    const lost = Object.keys(used).filter(k => before[k] && !seen[k]).map(k => used[k]);
    if (lost.length) return { needConfirm: true, lost: lost };
  }

  const ss = ctx.ss;
  const sh = ss.getSheetByName(BUDGET.ACCOUNTS) || (setupBudgetSheets_(ss), ss.getSheetByName(BUDGET.ACCOUNTS));
  try { migrateBudgetSheets_(ss); } catch (e) { /* 권한 없음 → 아래에서 알린다 */ }
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const n = Math.max(BUDGET.ACCOUNT_ROWS, rows.length + 20);
    if (sh.getMaxRows() < n + 1) sh.insertRowsAfter(sh.getMaxRows(), n + 1 - sh.getMaxRows());
    sh.getRange(2, 1, n, 2).setNumberFormat('@'); // 0으로 시작하는 코드도 그대로
    sh.getRange(2, 1, n, 4).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
    // 전입·전출·현재예산 수식 (줄이 늘었으면 그만큼 더)
    formatAccountSheet_(sh, n);
  } catch (e) {
    throw new Error('[예산과목] 시트를 고칠 권한이 없습니다. 관리자가 [관리 → 팀원·권한 → 권한 한 번에 적용]을 한 번 눌러 주세요.');
  } finally {
    lock.releaseLock();
  }
  SpreadsheetApp.flush();
  try { CacheService.getDocumentCache().remove(CK.budget); } catch (e) { /* 캐시 없음 */ }
  return refreshPart_('budget');
}

function getAccounts_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(BUDGET.ACCOUNTS);
  if (!sh || sh.getLastRow() < 2) return [];
  const head = String(sh.getRange(1, 2).getValue()).trim();
  if (head !== '계정코드') return []; // 아직 예전 양식 (관리자가 "권한 한 번에 적용"을 누르면 바뀐다)
  return sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues()
    .filter(r => String(r[2]).trim())
    .map(r => ({
      cost: String(r[0]).trim(),
      code: String(r[1]).trim(),
      name: String(r[2]).trim(),
      budget: Number(r[3]) || 0,
      inAmt: Number(r[4]) || 0,
      outAmt: Number(r[5]) || 0,
      current: Number(r[6]) || 0,
    }));
}

/**
 * 신청 → [예산전용] 시트에 한 줄 추가 (알림 메일은 보내지 않는다). 신청번호 반환
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
    throw new Error('예산 시트를 새 양식(코스트 + 계정코드 + 계정과목)으로 바꾸지 못했습니다. 관리자에게 "권한 한 번에 적용"을 요청하세요.');
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const accounts = getAccounts_();
    const account = accounts.find(a => a.cost === fromCost && a.name === from);
    const target = accounts.find(a => a.cost === toCost && a.name === to);
    if (!account) throw new Error(`[${budgetLabel_(fromCost, '', from)}] 예산을 찾을 수 없습니다. [예산과목] 시트를 확인하세요.`);
    if (!target) throw new Error(`[${budgetLabel_(toCost, '', to)}] 예산을 찾을 수 없습니다. [예산과목] 시트를 확인하세요.`);
    const fromLabel = budgetLabel_(fromCost, account.code, from);
    if (amount > account.current) {
      throw new Error(`현재예산을 초과합니다. (${fromLabel} 현재예산: ${account.current.toLocaleString()}원)`);
    }

    const sh = ss.getSheetByName(BUDGET.SHEET);
    const now = new Date();
    const id = nextBudgetId_(sh, now);
    const row = sh.getLastRow() + 1;
    // 코드·적용월이 숫자나 날짜로 바뀌지 않게 (0으로 시작하는 코드, '2026-10')
    sh.getRange(row, BCOL.FROM_COST, 1, 6).setNumberFormat('@');
    sh.getRange(row, BCOL.MONTH).setNumberFormat('@');
    sh.getRange(row, 1, 1, BCOL.MONTH).setValues([[id, now, requester, fromCost, account.code, from, toCost, target.code, to, amount, reason, '', month]]);
    SpreadsheetApp.flush();
    try { CacheService.getDocumentCache().remove(CK.budget); } catch (e) { /* 캐시 없음 */ }

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
