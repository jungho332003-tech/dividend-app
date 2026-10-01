/**
 * 공유게시판
 *
 * [게시판]  글번호 | 작성일시 | 작성자 | 분류 | 제목 | 내용 | 링크 | 고정(Y) | 삭제(Y)
 * [댓글]    글번호 | 작성일시 | 작성자 | 내용
 */

const BOARD = {
  SHEET: '게시판',
  COMMENTS: '댓글',
  CATEGORIES: ['공지', '자료', '일반'],
};

const PCOL = { ID: 1, DATE: 2, AUTHOR: 3, CATEGORY: 4, TITLE: 5, BODY: 6, LINK: 7, PINNED: 8, DELETED: 9, FILES: 10 };

function setupBoardSheets_(ss) {
  if (!ss.getSheetByName(BOARD.SHEET)) {
    const sh = ss.insertSheet(BOARD.SHEET);
    sh.getRange(1, 1, 1, 9).setValues([['글번호', '작성일시', '작성자', '분류', '제목', '내용', '링크', '고정(Y)', '삭제(Y)']]);
    styleHeader_(sh.getRange(1, 1, 1, 9));
    sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('F2:F').setWrap(true);
    [70, 130, 80, 60, 260, 420, 200, 70, 70].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
  }
  ensureHeader_(ss.getSheetByName(BOARD.SHEET), PCOL.FILES, '첨부');
  if (!ss.getSheetByName(BOARD.COMMENTS)) {
    const sh = ss.insertSheet(BOARD.COMMENTS);
    sh.getRange(1, 1, 1, 4).setValues([['글번호', '작성일시', '작성자', '내용']]);
    styleHeader_(sh.getRange(1, 1, 1, 4));
    sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('D2:D').setWrap(true);
    [70, 130, 80, 480].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
  }
}

/** 고정글 먼저, 그다음 최신순 */
function readBoard_(ss) {
  const sh = ss.getSheetByName(BOARD.SHEET);
  if (!sh || sh.getLastRow() < 2) return [];

  const comments = {};
  const csh = ss.getSheetByName(BOARD.COMMENTS);
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

  return sh.getRange(2, 1, sh.getLastRow() - 1, PCOL.FILES).getValues()
    .filter(r => String(r[0]) && String(r[PCOL.DELETED - 1]).toUpperCase() !== 'Y')
    .map(r => ({
      id: String(r[0]),
      date: r[1] instanceof Date ? fmt(r[1], 'yyyy-MM-dd HH:mm') : String(r[1]),
      author: String(r[2]),
      category: String(r[3]),
      title: String(r[4]),
      body: String(r[5]),
      link: String(r[6] || ''),
      pinned: String(r[7]).toUpperCase() === 'Y',
      files: filesFromCell_(r[PCOL.FILES - 1]),
      comments: comments[String(r[0])] || [],
    }))
    .sort((a, b) => (b.pinned - a.pinned) || b.date.localeCompare(a.date));
}

function apiAddPost(post) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'board');
  if (!me.name) throw new Error('[팀원] 시트에 등록된 사람만 글을 쓸 수 있습니다.');

  const title = String(post.title || '').trim();
  const body = String(post.body || '').trim();
  const category = BOARD.CATEGORIES.indexOf(post.category) >= 0 ? post.category : '일반';
  if (!title || !body) throw new Error('제목과 내용을 입력해 주세요.');
  if (category === '공지' && !me.isLeader && !me.isAdmin) throw new Error('공지는 팀장 또는 관리자만 올릴 수 있습니다.');
  const pinned = !!post.pinned && (me.isLeader || me.isAdmin);

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sh = ctx.ss.getSheetByName(BOARD.SHEET);
    const last = sh.getLastRow();
    const ids = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(([v]) => Number(v) || 0) : [];
    const id = (ids.length ? Math.max.apply(null, ids) : 0) + 1;
    sh.getRange(last + 1, 1, 1, PCOL.FILES).setValues([[id, new Date(), me.name, category, title, body, String(post.link || '').trim(), pinned ? 'Y' : '', '', filesToCell_(post.files)]]);
  } finally {
    lock.releaseLock();
  }
  return refreshPart_('board');
}

function apiDeletePost(id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'board');
  const sh = ctx.ss.getSheetByName(BOARD.SHEET);
  const rows = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 3).getValues();
  const idx = rows.findIndex(r => String(r[0]) === String(id));
  if (idx < 0) throw new Error('글을 찾을 수 없습니다.');
  if (rows[idx][2] !== me.name && !me.isLeader && !me.isAdmin) throw new Error('본인 글만 삭제할 수 있습니다.');
  sh.getRange(idx + 2, PCOL.DELETED).setValue('Y');
  return refreshPart_('board');
}

function apiAddComment(id, body) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'board');
  if (!me.name) throw new Error('[팀원] 시트에 등록된 사람만 댓글을 쓸 수 있습니다.');
  const text = String(body || '').trim();
  if (!text) throw new Error('댓글 내용을 입력해 주세요.');
  ctx.ss.getSheetByName(BOARD.COMMENTS).appendRow([String(id), new Date(), me.name, text]);
  return refreshPart_('board');
}
