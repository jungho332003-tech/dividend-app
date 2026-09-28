/**
 * 주차 넘기기 / 알림 / 트리거
 */

/* ---------- 주차 넘기기 ---------- */

function menuStartNewWeek() {
  const ui = SpreadsheetApp.getUi();
  const ctx = getContext();
  const target = addDays(ctx.thisWeek.monday, 7);
  const label = getWeekInfo(target, ctx.holidays).label;
  const ok = ui.alert('다음 주차로 넘기기',
    `현재 보고서를 보관하고 [${label}]로 넘어갑니다.\n각자의 "차주 주요업무"가 "금주 주요업무"로 이월됩니다. 진행할까요?`,
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  startNewWeek_(target);
  ui.alert(`[${label}] 주차로 넘어갔습니다.`);
}

/** 트리거: 매주 월요일 아침 자동 실행 */
function weeklyStart() {
  const ctx = getContext();
  const monday = mondayOf(new Date());
  if (monday > ctx.thisWeek.monday) startNewWeek_(monday);
}

function startNewWeek_(targetMonday) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const before = getContext();
    archiveReport_(before);

    setConfigValue(CFG.MONDAY, targetMonday);
    setConfigValue(CFG.DEADLINE, '');

    const ctx = getContext();
    ctx.members.filter(m => m.write).forEach(m => carryOver_(ctx.ss, m.name, ctx));

    const props = PropertiesService.getDocumentProperties();
    props.deleteProperty(PROP.REMINDED);
    props.deleteProperty(PROP.CLOSED);

    buildReport_(ctx);
  } finally {
    lock.releaseLock();
  }
}

/** 차주 → 금주 이월. "지속" 업무는 차주에도 그대로 남긴다. */
function carryOver_(ss, name, ctx) {
  const sh = ss.getSheetByName(SHEET.INPUT_PREFIX + name);
  if (!sh) return;

  const rows = sh.getRange(INPUT.FIRST_ROW, 4, INPUT.ROWS, 2).getValues().filter(r => String(r[0]).trim());
  const keep = rows.filter(r => String(r[1]).trim() === '지속');

  sh.getRange(INPUT.FIRST_ROW, 1, INPUT.ROWS, 2).clearContent();
  sh.getRange(INPUT.FIRST_ROW, 4, INPUT.ROWS, 2).clearContent();
  if (rows.length) sh.getRange(INPUT.FIRST_ROW, 1, rows.length, 2).setValues(rows);
  if (keep.length) sh.getRange(INPUT.FIRST_ROW, 4, keep.length, 2).setValues(keep);

  sh.getRange(INPUT.DONE_CELL).setValue(false);
  sh.getRange(INPUT.NOTE_CELL).clearContent();
  writeInputLabels_(sh, ctx);
}

/** 현재 주간보고를 "보관_yyyy-MM-dd" 시트로 복사해 숨김 */
function archiveReport_(ctx) {
  const ss = ctx.ss;
  const sh = ss.getSheetByName(SHEET.REPORT);
  if (!sh || sh.getLastRow() < REPORT.FIRST_ROW) return;

  const name = SHEET.ARCHIVE_PREFIX + ctx.thisWeek.key;
  const old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);

  const copy = sh.copyTo(ss).setName(name);
  protectSheet_(copy, []);
  copy.hideSheet();
}

/* ---------- 알림 ---------- */

function menuSendReminder() {
  const sent = sendReminder_(getContext());
  SpreadsheetApp.getUi().alert(sent.length ? `알림을 보냈습니다: ${sent.join(', ')}` : '미작성자가 없습니다.');
}

/** 작성완료 미체크 팀원에게 메일. 보낸 사람 이름 목록 반환 */
function sendReminder_(ctx) {
  const pending = pendingMembers_(ctx);
  pending.filter(m => m.email).forEach(m => {
    const sh = ctx.ss.getSheetByName(SHEET.INPUT_PREFIX + m.name);
    MailApp.sendEmail({
      to: m.email,
      subject: `[${ctx.cfg[CFG.TEAM]}] 주간업무 작성 요청 - ${ctx.thisWeek.label}`,
      htmlBody:
        `${m.name}님, 이번 주 주간업무 작성을 부탁드립니다.<br>` +
        `마감: <b>${fmt(ctx.deadline, 'M/dd')} (${DAY_KO[ctx.deadline.getDay()]}) ${hourKo(ctx.deadline.getHours())}</b><br><br>` +
        `<a href="${sh ? sheetUrl(ctx.ss, sh) : ctx.ss.getUrl()}">내 입력시트 열기</a><br>` +
        `작성 후 "작성완료 체크"를 눌러주세요.`,
    });
  });
  return pending.filter(m => m.email).map(m => m.name);
}

function pendingMembers_(ctx) {
  return ctx.members.filter(m => m.write && !readInput(ctx.ss, m.name).done);
}

/** 마감 시 팀장에게 보고서 링크 + 미작성자 안내 */
function sendClosingMail_(ctx) {
  const leaders = getLeaders(ctx.members);
  if (!leaders.length) return;
  const pending = pendingMembers_(ctx).map(m => m.name);
  const report = ctx.ss.getSheetByName(SHEET.REPORT);
  MailApp.sendEmail({
    to: leaders.map(m => m.email).join(','),
    subject: `[${ctx.cfg[CFG.TEAM]}] 주간업무 취합 완료 - ${ctx.thisWeek.label}`,
    htmlBody:
      `${ctx.thisWeek.label} 주간업무가 취합되었습니다.<br><br>` +
      `<a href="${sheetUrl(ctx.ss, report)}">주간보고 열기</a> (팀장 코멘트 칸에 의견을 남길 수 있습니다)<br><br>` +
      (pending.length ? `⚠️ 작성완료 미체크: ${pending.join(', ')}` : '✅ 전원 작성완료'),
  });
}

/* ---------- 트리거 ---------- */

/**
 * 매시간 실행
 *  - 평일 업무시간엔 보고서 자동 새로고침
 *  - 마감 N시간 전 미작성자 알림 (주 1회)
 *  - 마감 후 팀장에게 취합 메일 (주 1회)
 */
function hourlyCheck() {
  const now = new Date();
  const ctx = getContext();
  const props = PropertiesService.getDocumentProperties();
  const key = ctx.thisWeek.key;

  const day = now.getDay();
  if (day >= 1 && day <= 5 && now.getHours() >= 8 && now.getHours() <= 20) generateReport();

  const remindHours = Number(ctx.cfg[CFG.REMIND_HOURS]) || 3;
  const remindAt = new Date(ctx.deadline.getTime() - remindHours * 3600 * 1000);
  if (now >= remindAt && now < ctx.deadline && props.getProperty(PROP.REMINDED) !== key) {
    sendReminder_(ctx);
    props.setProperty(PROP.REMINDED, key);
  }

  if (now >= ctx.deadline && props.getProperty(PROP.CLOSED) !== key) {
    generateReport();
    sendClosingMail_(ctx);
    props.setProperty(PROP.CLOSED, key);
  }
}

// onBudgetEdit: 예전 버전(결재 기능)에서 만든 트리거도 정리한다
const TRIGGER_HANDLERS = ['hourlyCheck', 'weeklyStart', 'onBudgetEdit'];

/** 3단계: 자동 실행 트리거 등록 (다시 실행해도 중복 생성되지 않음) */
function setupTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => TRIGGER_HANDLERS.indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('hourlyCheck').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('weeklyStart').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();

  SpreadsheetApp.getUi().alert(
    '자동 실행을 켰습니다.\n\n' +
    '· 매주 월요일 7시: 다음 주차로 넘기기(이월)\n' +
    '· 매시간: 보고서 새로고침 / 마감 전 미작성 알림 / 마감 후 팀장 메일');
}
