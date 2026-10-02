/**
 * 연말정산 기준 (공제 요건·제출서류·응대 기준 모음)
 *
 * [기준]          번호 | 분류 | 제목 | 내용 | 적용 귀속연도 | 담당 | 원문 링크 | 최종수정 | 수정자 | 삭제(Y) | 첨부
 * [기준개정이력]  번호 | 일시 | 수정자 | 개정 내용
 *
 * 누구나 조회, 총괄·관리자만 등록/수정/삭제 (시트 보호로도 막는다).
 */

const RULES = {
  SHEET: '기준',
  HISTORY: '기준개정이력',
  CATEGORIES: ['인적공제', '주택자금·월세', '의료비', '교육비', '기부금', '신용카드 등', '연금·보험', '제출서류', '응대 기준', '기타'],
};

const RCOL = { ID: 1, CATEGORY: 2, TITLE: 3, BODY: 4, YEAR: 5, OWNER: 6, LINK: 7, UPDATED: 8, EDITOR: 9, DELETED: 10, FILES: 11 };

// 처음 설치할 때 넣는 예시 (세법은 해마다 바뀌므로 국세청 안내로 확인 후 고쳐 쓰도록 표시)
const SAMPLE_RULES = [
  ['인적공제', '기본공제 대상 부양가족 요건', '(예시 · 국세청 안내로 최신 기준 확인 후 사용)\n· 1인당 150만원 소득공제\n· 연간 소득금액 100만원 이하 (근로소득만 있으면 총급여 500만원 이하)\n· 나이: 직계존속 만 60세 이상, 직계비속 만 20세 이하, 형제자매 만 20세 이하 또는 60세 이상\n· 장애인은 나이 제한 없음 (소득 요건은 적용)\n· 같은 사람을 두 명이 공제할 수 없음 → 맞벌이 부부·형제 중복 여부 꼭 확인'],
  ['주택자금·월세', '월세액 세액공제', '(예시 · 국세청 안내로 최신 기준 확인 후 사용)\n· 무주택 세대의 세대주, 총급여 8천만원 이하\n· 국민주택규모(85㎡) 이하 또는 기준시가 4억원 이하 주택\n· 연 1,000만원 한도, 총급여 5,500만원 이하 17% / 그 외 15%\n· 제출: 임대차계약서 사본, 월세 이체 내역 (주민등록 주소지와 계약서 주소 일치 확인)'],
  ['주택자금·월세', '주택임차차입금 원리금 상환액', '(예시 · 국세청 안내로 최신 기준 확인 후 사용)\n· 무주택 세대의 세대주, 국민주택규모 주택 임차\n· 상환액의 40% 소득공제, 주택마련저축 납입액과 합산 연 400만원 한도\n· 제출: 주택자금상환증명서(금융기관), 임대차계약서, 주민등록등본'],
  ['의료비', '의료비 세액공제', '(예시 · 국세청 안내로 최신 기준 확인 후 사용)\n· 총급여 3% 초과분의 15% (난임시술비 30%, 미숙아·선천성이상아 20%)\n· 본인·65세 이상·장애인·6세 이하: 한도 없음 / 그 외 부양가족: 연 700만원\n· 실손보험금 받은 금액은 빼야 함\n· 간소화 자료에 없는 의료비(안경 등)는 영수증 별도 제출'],
  ['제출서류', '공통 제출 서류', '(예시 · 회사 기준에 맞게 고쳐 쓰세요)\n· 소득·세액공제신고서\n· 국세청 간소화 자료 PDF (부양가족 자료 제공 동의 포함)\n· 부양가족 변동 시 가족관계증명서\n· 중도입사자: 종전근무지 원천징수영수증·근로소득지급명세서'],
  ['응대 기준', '증빙 서류 받는 방법', '(예시 · 회사 기준에 맞게 고쳐 쓰세요)\n· 증빙은 지정된 경로(웹앱 첨부 또는 회사 메일)로만 받습니다. 개인 메신저 금지\n· 전화 문의는 본인 확인 후 안내\n· 안내한 내용은 응대기록에 남겨 다른 담당자도 볼 수 있게 합니다'],
];

function setupRuleSheets_(ss) {
  if (!ss.getSheetByName(RULES.SHEET)) {
    const sh = ss.insertSheet(RULES.SHEET);
    const header = ['번호', '분류', '제목', '내용', '적용 귀속연도', '담당', '원문 링크', '최종수정', '수정자', '삭제(Y)', '첨부'];
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    styleHeader_(sh.getRange(1, 1, 1, header.length));
    sh.getRange('B2:B').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(RULES.CATEGORIES).build());
    sh.getRange('H2:H').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('D2:D').setWrap(true);
    [50, 100, 220, 480, 90, 80, 200, 130, 80, 70, 200].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
    const year = String(getConfig()[CFG.YEAR] || defaultYear_());
    const now = new Date();
    const rows = SAMPLE_RULES.map((r, i) => [i + 1, r[0], r[1], r[2], year, '', '', now, '설치 예시', '', '']);
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
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
      year: cellText_(r[4]),
      owner: String(r[5] || ''),
      link: String(r[6] || ''),
      updated: r[7] instanceof Date ? fmt(r[7], 'yyyy-MM-dd') : String(r[7] || '').slice(0, 10),
      editor: String(r[8] || ''),
      files: filesFromCell_(r[RCOL.FILES - 1]),
      history: (history[String(r[0])] || []).reverse(),
    }));
}

/** 등록/수정. rule: { id?, category, title, body, year, owner, link, changeNote, files } */
function apiSaveRule(rule) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'rules');
  if (!me.isLeader && !me.isAdmin) throw new Error('연말정산 기준은 총괄 또는 관리자만 등록·수정할 수 있습니다.');

  const title = String(rule.title || '').trim();
  const body = String(rule.body || '').trim();
  if (!title || !body) throw new Error('제목과 내용을 입력해 주세요.');
  const category = RULES.CATEGORIES.indexOf(rule.category) >= 0 ? rule.category : '기타';
  const editor = me.name || me.email;
  const now = new Date();

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  let id;
  try {
    const sh = ctx.ss.getSheetByName(RULES.SHEET);
    const hsh = ctx.ss.getSheetByName(RULES.HISTORY);
    const last = sh.getLastRow();
    const ids = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(([v]) => String(v)) : [];
    const values = [category, title, body, String(rule.year || '').trim(), String(rule.owner || '').trim(), String(rule.link || '').trim(), now, editor];

    id = String(rule.id || '');
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
  registerFiles_('rules', id, title, rule.files, me);
  return Object.assign(refreshPart_('rules'), { savedId: id });
}

function apiDeleteRule(id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'rules');
  if (!me.isLeader && !me.isAdmin) throw new Error('총괄 또는 관리자만 삭제할 수 있습니다.');
  const sh = ctx.ss.getSheetByName(RULES.SHEET);
  const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(([v]) => String(v));
  const idx = ids.indexOf(String(id));
  if (idx < 0) throw new Error('기준을 찾을 수 없습니다.');
  sh.getRange(idx + 2, RCOL.DELETED).setValue('Y');
  ctx.ss.getSheetByName(RULES.HISTORY).appendRow([String(id), new Date(), me.name || me.email, '삭제']);
  return refreshPart_('rules');
}
