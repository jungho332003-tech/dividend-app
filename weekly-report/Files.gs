/**
 * 파일 첨부 (구글 드라이브)
 *
 * [설정] > "첨부 폴더 ID" 의 드라이브 폴더 아래에 메뉴별 하위 폴더(게시판/회사기준/예산전용)를 만들어 저장한다.
 * 폴더 ID: 드라이브에서 폴더를 열었을 때 주소 .../folders/<여기> 부분
 * 팀원이 올리려면 그 폴더가 팀원에게 "편집자"로 공유되어 있어야 한다.
 *
 * 글/기준/신청에는 첨부 목록을 JSON 문자열로 한 칸에 저장한다: [{id, name, url, size, mime}]
 */

const FILES = {
  MAX_BYTES: 10 * 1024 * 1024,
  AREAS: { board: '게시판', rules: '회사기준', budget: '예산전용' },
};

/** file: { name, mimeType, data(base64), area } → 첨부 정보 */
function apiUploadFile(file) {
  const me = currentMember_(getMembers());
  if (!me.name && !me.isAdmin) throw new Error('[팀원] 시트에 등록된 사람만 파일을 올릴 수 있습니다.');

  const name = String(file.name || '').trim() || '첨부파일';
  const bytes = Utilities.base64Decode(String(file.data || ''));
  if (!bytes.length) throw new Error('빈 파일은 올릴 수 없습니다.');
  if (bytes.length > FILES.MAX_BYTES) throw new Error('10MB 이하 파일만 올릴 수 있습니다. 큰 파일은 드라이브에 직접 올리고 링크를 붙여주세요.');

  const folder = areaFolder_(FILES.AREAS[file.area] || '기타');
  const created = folder.createFile(Utilities.newBlob(bytes, file.mimeType || 'application/octet-stream', name));
  if (me.name) created.setDescription(`올린 사람: ${me.name}`);
  return {
    id: created.getId(),
    name: created.getName(),
    url: created.getUrl(),
    size: bytes.length,
    mime: created.getMimeType(),
  };
}

function areaFolder_(areaName) {
  const id = String(getConfig()[CFG.DRIVE_FOLDER] || '').trim();
  if (!id) throw new Error('첨부 폴더가 설정되지 않았습니다. 관리자에게 [설정] > "첨부 폴더 ID" 입력을 요청하세요.');
  let root;
  try {
    root = DriveApp.getFolderById(id);
  } catch (e) {
    throw new Error('첨부 폴더에 접근할 수 없습니다. 폴더가 "편집자"로 공유되어 있는지 관리자에게 확인하세요.');
  }
  const it = root.getFoldersByName(areaName);
  return it.hasNext() ? it.next() : root.createFolder(areaName);
}

/** 화면에서 넘어온 첨부 목록을 검증해 시트에 저장할 JSON 문자열로 만든다 */
function filesToCell_(files) {
  const list = (files || [])
    .filter(f => f && f.id && f.name)
    .slice(0, 20)
    .map(f => ({ id: String(f.id), name: String(f.name), url: String(f.url || ''), size: Number(f.size) || 0, mime: String(f.mime || '') }));
  return list.length ? JSON.stringify(list) : '';
}

function filesFromCell_(v) {
  if (!v) return [];
  try {
    const list = JSON.parse(String(v));
    return Array.isArray(list) ? list : [];
  } catch (e) {
    return [];
  }
}

/** 예전에 만든 시트에 첨부 열 제목이 없으면 채운다 */
function ensureHeader_(sh, col, title) {
  const cell = sh.getRange(1, col);
  if (!cell.getValue()) {
    cell.setValue(title);
    styleHeader_(cell);
    sh.setColumnWidth(col, 220);
  }
}
