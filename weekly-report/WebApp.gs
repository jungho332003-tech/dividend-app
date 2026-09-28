/**
 * 웹앱 (대시보드 화면)
 *
 * 배포: Apps Script 편집기 > 배포 > 새 배포 > 웹 앱
 *   - 다음 사용자 인증 정보로 실행: "웹 앱에 액세스하는 사용자"
 *   - 액세스 권한: 조직 내 모든 사용자 또는 "Google 계정이 있는 모든 사용자"
 * 접속한 사람의 이메일로 [팀원] 시트에서 이름/역할을 찾고, 시트 보호 권한도 그대로 적용된다.
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('App')
    .setTitle('인사팀 업무관리')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- 조회 ---------- */

/** 화면 전체 데이터. google.script.run 은 Date를 넘길 수 없어 문자열로 변환한다. */
function apiBootstrap() {
  const ctx = getContext();
  const ss = ctx.ss;
  const me = currentMember_(ctx.members);
  const report = ss.getSheetByName(SHEET.REPORT);
  const comments = report ? readLeaderComments_(report, ctx.thisWeek.key) : {};

  return {
    team: ctx.cfg[CFG.TEAM] || '',
    sheetUrl: ss.getUrl(),
    me: me,
    week: {
      key: ctx.thisWeek.key,
      thisLabel: ctx.thisWeek.label,
      nextLabel: ctx.nextWeek.label,
      deadlineIso: ctx.deadline.toISOString(),
      deadlineText: `${fmt(ctx.deadline, 'M/dd')} (${DAY_KO[ctx.deadline.getDay()]}) ${hourKo(ctx.deadline.getHours())}`,
      title: `${ctx.cfg[CFG.TEAM] || ''} 주간업무 ${deadlineText(ctx)}`,
      thisStart: ymd(ctx.thisWeek.start),
      thisEnd: ymd(ctx.thisWeek.end),
    },
    members: ctx.members.filter(m => m.write).map(m => {
      const input = readInput(ss, m.name);
      return {
        name: m.name,
        hasSheet: input.exists,
        done: input.done,
        note: input.note,
        thisWeek: input.thisWeek.map(toClientItem_),
        nextWeek: input.nextWeek.map(toClientItem_),
        comment: String(comments[m.name] || ''),
      };
    }),
    budget: readBudget_(ss),
    board: readBoard_(ss),
  };
}

function toClientItem_(item) {
  const due = item.due;
  return {
    task: item.task,
    due: formatDue(due),
    dueIso: due instanceof Date ? ymd(due) : '',
  };
}

/** 접속자 → 팀원 정보 */
function currentMember_(members) {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const m = members.find(x => x.email && x.email.toLowerCase() === email);
  let isAdmin = false;
  try {
    const owner = SpreadsheetApp.getActive().getOwner();
    isAdmin = !!owner && owner.getEmail().toLowerCase() === email;
  } catch (e) { /* 공유 드라이브 등 소유자 조회 불가 */ }
  return {
    email: email,
    name: m ? m.name : '',
    role: m ? m.role : '',
    writer: !!(m && m.write),
    isLeader: !!(m && m.role === '팀장'),
    isAdmin: isAdmin,
  };
}

function readBudget_(ss) {
  const accSh = ss.getSheetByName(BUDGET.ACCOUNTS);
  const accRows = accSh && accSh.getLastRow() >= 2
    ? accSh.getRange(2, 1, accSh.getLastRow() - 1, 7).getValues().filter(r => String(r[0]).trim())
    : [];

  const reqSh = ss.getSheetByName(BUDGET.SHEET);
  const reqRows = reqSh && reqSh.getLastRow() >= 2
    ? reqSh.getRange(2, 1, reqSh.getLastRow() - 1, 10).getValues()
    : [];

  return {
    accounts: accRows.map(r => ({
      name: String(r[0]).trim(),
      budget: Number(r[1]) || 0,
      inOk: Number(r[2]) || 0,
      outOk: Number(r[3]) || 0,
      current: Number(r[4]) || 0,
      pending: Number(r[5]) || 0,
      available: Number(r[6]) || 0,
    })),
    requests: reqRows
      .map((r, i) => ({ r: r, row: i + 2 }))
      .filter(x => String(x.r[0]).trim())
      .map(x => ({
        row: x.row,
        id: String(x.r[0]),
        date: x.r[1] instanceof Date ? fmt(x.r[1], 'yyyy-MM-dd HH:mm') : String(x.r[1]),
        req: String(x.r[2]),
        from: String(x.r[3]),
        to: String(x.r[4]),
        amount: Number(x.r[5]) || 0,
        reason: String(x.r[6]),
        status: String(x.r[7] || ''),
        opinion: String(x.r[8] || ''),
        doneAt: x.r[9] instanceof Date ? fmt(x.r[9], 'yyyy-MM-dd HH:mm') : String(x.r[9] || ''),
      }))
      .reverse(),
  };
}

/* ---------- 내 주간업무 저장 ---------- */

/**
 * payload: { thisWeek:[{task, due}], nextWeek:[...], note, done }
 * due: 'yyyy-mm-dd' | '지속' | ''
 */
function apiSaveMyWeek(payload) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  if (!me.writer) throw new Error('작성 대상 팀원으로 등록되어 있지 않습니다. 관리자에게 [팀원] 시트 등록을 요청하세요.');

  const sh = ctx.ss.getSheetByName(SHEET.INPUT_PREFIX + me.name);
  if (!sh) throw new Error('입력시트가 아직 없습니다. 관리자에게 "팀원 입력시트 만들기"를 요청하세요.');

  const toRows = list => (list || [])
    .filter(x => String(x.task || '').trim())
    .slice(0, INPUT.ROWS)
    .map(x => [String(x.task).trim(), parseDue_(x.due)]);

  const thisRows = toRows(payload.thisWeek);
  const nextRows = toRows(payload.nextWeek);

  sh.getRange(INPUT.FIRST_ROW, 1, INPUT.ROWS, 2).clearContent();
  sh.getRange(INPUT.FIRST_ROW, 4, INPUT.ROWS, 2).clearContent();
  if (thisRows.length) sh.getRange(INPUT.FIRST_ROW, 1, thisRows.length, 2).setValues(thisRows);
  if (nextRows.length) sh.getRange(INPUT.FIRST_ROW, 4, nextRows.length, 2).setValues(nextRows);
  sh.getRange(INPUT.NOTE_CELL).setValue(String(payload.note || '').trim());
  sh.getRange(INPUT.DONE_CELL).setValue(!!payload.done);
  return apiBootstrap();
}

function parseDue_(v) {
  const s = String(v || '').trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return s;
}

/* ---------- 팀장 코멘트 ---------- */

function apiSaveComment(name, comment) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  if (!me.isLeader) throw new Error('팀장만 코멘트를 남길 수 있습니다.');

  const sh = ctx.ss.getSheetByName(SHEET.REPORT);
  const current = PropertiesService.getDocumentProperties().getProperty(PROP.REPORT_WEEK) === ctx.thisWeek.key;
  if (!sh || !current || sh.getLastRow() < REPORT.FIRST_ROW) {
    throw new Error('이번 주 주간보고가 아직 만들어지지 않았습니다. 잠시 후 다시 시도하세요.');
  }
  const names = sh.getRange(REPORT.FIRST_ROW, 1, sh.getLastRow() - REPORT.FIRST_ROW + 1, 1).getValues();
  const idx = names.findIndex(([v]) => String(v).trim() === name);
  if (idx < 0) throw new Error(`주간보고에서 ${name}님을 찾을 수 없습니다.`);

  sh.getRange(REPORT.FIRST_ROW + idx, REPORT.COMMENT_COL).setValue(String(comment || '').trim());
  return apiBootstrap();
}

/* ---------- 예산전용 ---------- */

function apiSubmitBudget(form) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  const requester = me.name || String(form.requester || '').trim();
  submitBudgetTransfer(Object.assign({}, form, { requester: requester }));
  return apiBootstrap();
}

/** 팀장 결재: 시트에서 결재 칸을 바꾼 것과 같은 효과 (웹앱에서의 수정은 onEdit 트리거가 돌지 않으므로 메일도 여기서 보낸다) */
function apiDecideBudget(row, id, status, opinion) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  if (!me.isLeader) throw new Error('팀장만 결재할 수 있습니다.');
  if (BUDGET.DECISIONS.indexOf(status) < 0) throw new Error('승인 또는 반려를 선택하세요.');

  const sh = ctx.ss.getSheetByName(BUDGET.SHEET);
  const v = sh.getRange(row, 1, 1, 10).getValues()[0];
  if (String(v[BCOL.ID - 1]) !== id) throw new Error('신청 내역이 바뀌었습니다. 새로고침 후 다시 시도하세요.');
  if (String(v[BCOL.STATUS - 1] || '')) throw new Error(`이미 ${v[BCOL.STATUS - 1]}된 건입니다.`);

  sh.getRange(row, BCOL.STATUS, 1, 3).setValues([[status, String(opinion || '').trim(), new Date()]]);
  notifyBudgetDecision_(sh, row);
  return apiBootstrap();
}

/* ---------- 알림 ---------- */

function apiRemind() {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  if (!me.isLeader && !me.isAdmin) throw new Error('팀장 또는 관리자만 알림을 보낼 수 있습니다.');
  return sendReminder_(ctx);
}
