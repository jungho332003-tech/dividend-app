/**
 * 회사 기준 (사내 규정·기준 모음)
 *
 * [회사기준]      번호 | 분류 | 제목 | 내용 | 시행일 | 담당 | 원문 링크 | 최종수정 | 수정자 | 삭제(Y)
 * [기준개정이력]  번호 | 일시 | 수정자 | 개정 내용
 *
 * 누구나 조회, 팀장·관리자만 등록/수정/삭제 (시트 보호로도 막는다).
 */

const RULES = {
  SHEET: '회사기준',
  HISTORY: '기준개정이력',
  CATEGORIES: ['인사·근태', '급여·수당', '복리후생', '출장·경비', '교육', '기타'],
};

const RCOL = { ID: 1, CATEGORY: 2, TITLE: 3, BODY: 4, EFFECTIVE: 5, OWNER: 6, LINK: 7, UPDATED: 8, EDITOR: 9, DELETED: 10, FILES: 11 };

function setupRuleSheets_(ss) {
  if (!ss.getSheetByName(RULES.SHEET)) {
    const sh = ss.insertSheet(RULES.SHEET);
    sh.getRange(1, 1, 1, 10).setValues([['번호', '분류', '제목', '내용', '시행일', '담당', '원문 링크', '최종수정', '수정자', '삭제(Y)']]);
    styleHeader_(sh.getRange(1, 1, 1, 10));
    sh.getRange('B2:B').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(RULES.CATEGORIES).build());
    sh.getRange('E2:E').setNumberFormat('yyyy-mm-dd');
    sh.getRange('H2:H').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('D2:D').setWrap(true);
    [50, 90, 220, 480, 100, 80, 200, 130, 80, 70].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
  }
  ensureHeader_(ss.getSheetByName(RULES.SHEET), RCOL.FILES, '첨부');
  if (!ss.getSheetByName(RULES.HISTORY)) {
    const sh = ss.insertSheet(RULES.HISTORY);
    sh.getRange(1, 1, 1, 4).setValues([['번호', '일시', '수정자', '개정 내용']]);
    styleHeader_(sh.getRange(1, 1, 1, 4));
    sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
    [50, 130, 80, 480].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
  }
}

function readRules_(ss) {
  const sh = ss.getSheetByName(RULES.SHEET);
  if (!sh || sh.getLastRow() < 2) return [];

  const history = {};
  const hsh = ss.getSheetByName(RULES.HISTORY);
  if (hsh && hsh.getLastRow() >= 2) {
    hsh.getRange(2, 1, hsh.getLastRow() - 1, 4).getValues().forEach(r => {
      const id = String(r[0]);
      if (!id) return;
      (history[id] = history[id] || []).push({
        date: r[1] instanceof Date ? fmt(r[1], 'yyyy-MM-dd') : String(r[1]).slice(0, 10),
        editor: String(r[2]),
        note: String(r[3]),
      });
    });
  }

  return sh.getRange(2, 1, sh.getLastRow() - 1, RCOL.FILES).getValues()
    .filter(r => String(r[0]) && String(r[RCOL.DELETED - 1]).toUpperCase() !== 'Y')
    .map(r => ({
      id: String(r[0]),
      category: String(r[1]),
      title: String(r[2]),
      body: String(r[3]),
      effective: r[4] instanceof Date ? ymd(r[4]) : String(r[4] || ''),
      owner: String(r[5] || ''),
      link: String(r[6] || ''),
      updated: r[7] instanceof Date ? fmt(r[7], 'yyyy-MM-dd') : String(r[7] || '').slice(0, 10),
      editor: String(r[8] || ''),
      files: filesFromCell_(r[RCOL.FILES - 1]),
      history: (history[String(r[0])] || []).reverse(),
    }));
}

/**
 * 등록/수정. rule: { id?, category, title, body, effective:'yyyy-mm-dd', owner, link, changeNote }
 */
function apiSaveRule(rule) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  if (!me.isLeader && !me.isAdmin) throw new Error('회사 기준은 팀장 또는 관리자만 등록·수정할 수 있습니다.');

  const title = String(rule.title || '').trim();
  const body = String(rule.body || '').trim();
  if (!title || !body) throw new Error('제목과 내용을 입력해 주세요.');
  const category = RULES.CATEGORIES.indexOf(rule.category) >= 0 ? rule.category : '기타';
  const effective = parseDue_(rule.effective);
  const editor = me.name || me.email;
  const now = new Date();

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sh = ctx.ss.getSheetByName(RULES.SHEET);
    const hsh = ctx.ss.getSheetByName(RULES.HISTORY);
    const last = sh.getLastRow();
    const ids = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(([v]) => String(v)) : [];
    const values = [category, title, body, effective, String(rule.owner || '').trim(), String(rule.link || '').trim(), now, editor];

    let id = String(rule.id || '');
    const idx = id ? ids.indexOf(id) : -1;
    if (idx >= 0) {
      sh.getRange(idx + 2, RCOL.CATEGORY, 1, values.length).setValues([values]);
      sh.getRange(idx + 2, RCOL.FILES).setValue(filesToCell_(rule.files));
      hsh.appendRow([id, now, editor, String(rule.changeNote || '').trim() || '내용 수정']);
    } else {
      id = String((ids.length ? Math.max.apply(null, ids.map(Number).filter(n => !isNaN(n))) : 0) + 1);
      sh.getRange(last + 1, 1, 1, RCOL.FILES).setValues([[id].concat(values, ['', filesToCell_(rule.files)])]);
      hsh.appendRow([id, now, editor, String(rule.changeNote || '').trim() || '최초 등록']);
    }
  } finally {
    lock.releaseLock();
  }
  return refreshPart_('rules');
}

function apiDeleteRule(id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  if (!me.isLeader && !me.isAdmin) throw new Error('팀장 또는 관리자만 삭제할 수 있습니다.');
  const sh = ctx.ss.getSheetByName(RULES.SHEET);
  const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(([v]) => String(v));
  const idx = ids.indexOf(String(id));
  if (idx < 0) throw new Error('기준을 찾을 수 없습니다.');
  sh.getRange(idx + 2, RCOL.DELETED).setValue('Y');
  ctx.ss.getSheetByName(RULES.HISTORY).appendRow([String(id), new Date(), me.name || me.email, '삭제']);
  return refreshPart_('rules');
}
