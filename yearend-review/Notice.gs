/**
 * 공지사항 / 특이사항 안내
 *
 * [공지]     번호 | 작성일시 | 작성자 | 분류 | 중요도 | 제목 | 내용 | 게시시작 | 게시종료 | 고정(Y) | 삭제(Y) | 첨부 | 수정일시
 * [공지댓글] 번호 | 작성일시 | 작성자 | 내용
 *
 * 공지·응대지침은 총괄·관리자만, 자료·질문은 누구나 올린다. 읽음 표시는 사람별로 저장(사용자 속성).
 */

const NOTICE = {
  SHEET: '공지',
  COMMENTS: '공지댓글',
  CATEGORIES: ['공지', '응대지침', '자료', '질문'],
  LEADER_ONLY: ['공지', '응대지침'],
  LEVELS: ['긴급', '중요', '일반'],
};

const NCOL = { ID: 1, DATE: 2, AUTHOR: 3, CATEGORY: 4, LEVEL: 5, TITLE: 6, BODY: 7, START: 8, END: 9, PINNED: 10, DELETED: 11, FILES: 12, UPDATED: 13 };

function setupNoticeSheets_(ss) {
  if (!ss.getSheetByName(NOTICE.SHEET)) {
    const sh = ss.insertSheet(NOTICE.SHEET);
    const header = ['번호', '작성일시', '작성자', '분류', '중요도', '제목', '내용', '게시시작', '게시종료', '고정(Y)', '삭제(Y)', '첨부', '수정일시'];
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    styleHeader_(sh.getRange(1, 1, 1, header.length));
    sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('H2:I').setNumberFormat('yyyy-mm-dd');
    sh.getRange('M2:M').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('G2:G').setWrap(true);
    [50, 130, 80, 80, 60, 260, 420, 100, 100, 60, 60, 200, 130].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
  }
  if (!ss.getSheetByName(NOTICE.COMMENTS)) {
    const sh = ss.insertSheet(NOTICE.COMMENTS);
    sh.getRange(1, 1, 1, 4).setValues([['번호', '작성일시', '작성자', '내용']]);
    styleHeader_(sh.getRange(1, 1, 1, 4));
    sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('D2:D').setWrap(true);
    [50, 130, 80, 480].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
  }
}

/** 고정 → 중요도 → 최신순 */
function readNotices_(ss) {
  const sh = ss.getSheetByName(NOTICE.SHEET);
  if (!sh || sh.getLastRow() < 2) return [];

  const comments = {};
  const csh = ss.getSheetByName(NOTICE.COMMENTS);
  if (csh && csh.getLastRow() >= 2) {
    csh.getRange(2, 1, csh.getLastRow() - 1, 4).getValues().forEach(r => {
      const id = String(r[0]);
      if (!id) return;
      (comments[id] = comments[id] || []).push({
        date: r[1] instanceof Date ? fmt(r[1], 'yyyy-MM-dd HH:mm') : String(r[1]),
        author: String(r[2]),
        body: String(r[3]),
      });
    });
  }

  const dt = v => v instanceof Date ? fmt(v, 'yyyy-MM-dd HH:mm') : String(v || '');
  return sh.getRange(2, 1, sh.getLastRow() - 1, NCOL.UPDATED).getValues()
    .filter(r => String(r[0]) && String(r[NCOL.DELETED - 1]).toUpperCase() !== 'Y')
    .map(r => ({
      id: String(r[0]),
      date: dt(r[1]),
      author: String(r[2]),
      category: String(r[3]) || '공지',
      level: NOTICE.LEVELS.indexOf(String(r[4])) >= 0 ? String(r[4]) : '일반',
      title: String(r[5]),
      body: String(r[6]),
      start: cellText_(r[7]),
      end: cellText_(r[8]),
      pinned: String(r[9]).toUpperCase() === 'Y',
      files: filesFromCell_(r[NCOL.FILES - 1]),
      updated: dt(r[NCOL.UPDATED - 1]) || dt(r[1]),
      comments: comments[String(r[0])] || [],
    }))
    .sort((a, b) => (b.pinned - a.pinned) || (NOTICE.LEVELS.indexOf(a.level) - NOTICE.LEVELS.indexOf(b.level)) || b.date.localeCompare(a.date));
}

/** 등록/수정. n: { id?, category, level, title, body, start, end, pinned, files } */
function apiSaveNotice(n) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'notice');
  if (!me.name && !me.isAdmin) throw new Error('[담당자] 시트에 등록된 사람만 글을 쓸 수 있습니다.');
  const lead = me.isLeader || me.isAdmin;

  const title = String(n.title || '').trim();
  const body = String(n.body || '').trim();
  if (!title || !body) throw new Error('제목과 내용을 입력해 주세요.');
  const category = NOTICE.CATEGORIES.indexOf(n.category) >= 0 ? n.category : '질문';
  if (NOTICE.LEADER_ONLY.indexOf(category) >= 0 && !lead) throw new Error(`[${category}]는 총괄 또는 관리자만 올릴 수 있습니다.`);
  const level = lead && NOTICE.LEVELS.indexOf(n.level) >= 0 ? n.level : '일반';
  const start = String(n.start || ''), end = String(n.end || '');
  if (start && end && start > end) throw new Error('게시 종료일이 시작일보다 빠릅니다.');
  const pinned = !!n.pinned && lead;
  const now = new Date();

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  let id;
  try {
    const sh = ctx.ss.getSheetByName(NOTICE.SHEET);
    const last = sh.getLastRow();
    const rows = last >= 2 ? sh.getRange(2, 1, last - 1, 3).getValues() : [];
    id = String(n.id || '');
    const idx = id ? rows.findIndex(r => String(r[0]) === id) : -1;
    const values = [category, level, title, body, parseDate_(start), parseDate_(end), pinned ? 'Y' : ''];
    if (idx >= 0) {
      if (String(rows[idx][2]) !== me.name && !lead) throw new Error('본인 글만 고칠 수 있습니다.');
      sh.getRange(idx + 2, NCOL.CATEGORY, 1, values.length).setValues([values]);
      sh.getRange(idx + 2, NCOL.FILES, 1, 2).setValues([[filesToCell_(n.files), now]]);
    } else {
      id = String((rows.length ? Math.max.apply(null, rows.map(r => Number(r[0]) || 0)) : 0) + 1);
      sh.getRange(last + 1, 1, 1, NCOL.UPDATED).setValues([[id, now, me.name || me.email].concat(values, ['', filesToCell_(n.files), now])]);
    }
  } finally {
    lock.releaseLock();
  }
  registerFiles_('notice', id, title, n.files, me);
  markRead_(id);
  return Object.assign(refreshPart_('notices'), { savedId: id, reads: getReads_() });
}

function apiDeleteNotice(id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'notice');
  const sh = ctx.ss.getSheetByName(NOTICE.SHEET);
  const rows = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 3).getValues();
  const idx = rows.findIndex(r => String(r[0]) === String(id));
  if (idx < 0) throw new Error('글을 찾을 수 없습니다.');
  if (rows[idx][2] !== me.name && !me.isLeader && !me.isAdmin) throw new Error('본인 글만 삭제할 수 있습니다.');
  sh.getRange(idx + 2, NCOL.DELETED).setValue('Y');
  return refreshPart_('notices');
}

function apiAddNoticeComment(id, body) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'notice');
  if (!me.name && !me.isAdmin) throw new Error('[담당자] 시트에 등록된 사람만 댓글을 쓸 수 있습니다.');
  const text = String(body || '').trim();
  if (!text) throw new Error('댓글 내용을 입력해 주세요.');
  ctx.ss.getSheetByName(NOTICE.COMMENTS).appendRow([String(id), new Date(), me.name || me.email, text]);
  return refreshPart_('notices');
}

/* ---------- 읽음 표시 (사람마다 따로: 웹앱이 "접속한 사용자" 권한으로 실행되므로 사용자 속성에 저장) ---------- */

function getReads_() {
  try {
    return JSON.parse(PropertiesService.getUserProperties().getProperty('READ_NOTICES') || '[]');
  } catch (e) {
    return [];
  }
}

function markRead_(id) {
  const list = getReads_().filter(x => x !== String(id));
  list.push(String(id));
  PropertiesService.getUserProperties().setProperty('READ_NOTICES', JSON.stringify(list.slice(-500)));
  return list;
}

function apiMarkRead(id) {
  return { reads: markRead_(id) };
}

function apiMarkUnread(id) {
  const list = getReads_().filter(x => x !== String(id));
  PropertiesService.getUserProperties().setProperty('READ_NOTICES', JSON.stringify(list));
  return { reads: list };
}
