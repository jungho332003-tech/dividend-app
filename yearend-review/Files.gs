/**
 * 첨부 파일 (구글 드라이브)
 *
 * [설정] > "첨부 폴더 ID" 의 드라이브 폴더 아래에 구분별 하위 폴더를 만들어 저장한다.
 * 증빙 서류에는 개인정보가 들어 있으므로 이 폴더는 담당자에게만 공유한다 (관리 > 권한 한 번에 적용).
 *
 * [첨부] 시트에 모든 첨부를 한 줄씩 기록해 [첨부파일] 메뉴에서 한곳에 모아 본다.
 *   파일ID | 올린일시 | 올린사람 | 구분 | 대상번호 | 대상 | 파일명 | URL | 크기 | 형식 | 메모 | 삭제(Y)
 * 공지·기준 글에는 첨부 목록을 JSON 문자열로도 저장한다: [{id, name, url, size, mime}]
 */

const FILES = {
  SHEET: '첨부',
  MAX_BYTES: 10 * 1024 * 1024,
  AREAS: { person: '대상자 증빙', notice: '공지사항', rules: '연말정산 기준', etc: '일반 자료' },
  MENU: { person: 'review', notice: 'notice', rules: 'rules', etc: 'files' },
};

const FCOL = { ID: 1, DATE: 2, UPLOADER: 3, AREA: 4, REF: 5, REF_LABEL: 6, NAME: 7, URL: 8, SIZE: 9, MIME: 10, MEMO: 11, DELETED: 12, YEAR: 13 };

function setupFileSheet_(ss) {
  if (ss.getSheetByName(FILES.SHEET)) { ensureHeader_(ss.getSheetByName(FILES.SHEET), FCOL.YEAR, '귀속연도'); return; }
  const sh = ss.insertSheet(FILES.SHEET);
  const header = ['파일ID', '올린일시', '올린사람', '구분', '대상번호', '대상', '파일명', 'URL', '크기', '형식', '메모', '삭제(Y)', '귀속연도'];
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  styleHeader_(sh.getRange(1, 1, 1, header.length));
  sh.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
  [120, 130, 80, 100, 80, 160, 260, 220, 80, 120, 200, 60, 80].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.hideColumns(1);
  sh.setFrozenRows(1);
}

/**
 * file: { name, mimeType, data(base64), area, ref?, refLabel?, memo? } → 첨부 정보
 * 대상자 증빙과 일반 자료는 올리자마자 [첨부]에 기록하고, 공지·기준은 글을 저장할 때 기록한다.
 */
function apiUploadFile(file) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  if (!me.name && !me.isAdmin) throw new Error('[담당자] 시트에 등록된 사람만 파일을 올릴 수 있습니다.');
  const area = FILES.AREAS[file.area] ? file.area : 'etc';
  requireMenu_(me, FILES.MENU[area]);

  if (area === 'person') {
    const p = cached_(CK.people, () => readPeople_(ctx.ss)).find(x => (x.empNo || x.id) === String(file.ref || ''));
    if (!p || !canSee_(me, p)) throw new Error('열람 범위 밖의 대상자에게는 파일을 올릴 수 없습니다.');
  }
  const name = String(file.name || '').trim() || '첨부파일';
  const bytes = Utilities.base64Decode(String(file.data || ''));
  if (!bytes.length) throw new Error('빈 파일은 올릴 수 없습니다.');
  if (bytes.length > FILES.MAX_BYTES) throw new Error('10MB 이하 파일만 올릴 수 있습니다. 큰 파일은 드라이브에 직접 올리고 링크를 붙여주세요.');

  const folder = areaFolder_(FILES.AREAS[area], ctx);
  const created = folder.createFile(Utilities.newBlob(bytes, file.mimeType || 'application/octet-stream', name));
  created.setDescription(`올린 사람: ${me.name || me.email}${file.refLabel ? ` / 대상: ${file.refLabel}` : ''}`);
  const info = { id: created.getId(), name: created.getName(), url: created.getUrl(), size: bytes.length, mime: created.getMimeType() };

  if (area === 'person' || area === 'etc') {
    registerFiles_(area, String(file.ref || ''), String(file.refLabel || ''), [info], me, String(file.memo || ''));
    return Object.assign({ file: info }, refreshPart_('files'));
  }
  return { file: info };
}

function areaFolder_(areaName, ctx) {
  const id = String((ctx ? ctx.cfg : getConfig())[CFG.DRIVE_FOLDER] || '').trim();
  if (!id) throw new Error('첨부 폴더가 설정되지 않았습니다. 관리자에게 관리 > 설정 > "첨부 폴더" 입력을 요청하세요.');
  let root;
  try {
    root = DriveApp.getFolderById(id);
  } catch (e) {
    throw new Error('첨부 폴더에 접근할 수 없습니다. 폴더가 "편집자"로 공유되어 있는지 관리자에게 확인하세요.');
  }
  const it = root.getFoldersByName(areaName);
  return it.hasNext() ? it.next() : root.createFolder(areaName);
}

/** [첨부] 시트에 아직 없는 파일만 기록 */
function registerFiles_(area, ref, refLabel, files, me, memo) {
  const list = (files || []).filter(f => f && f.id);
  if (!list.length) return;
  const sh = SpreadsheetApp.getActive().getSheetByName(FILES.SHEET);
  const last = sh.getLastRow();
  const have = new Set(last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(([v]) => String(v)) : []);
  const now = new Date();
  const rows = list.filter(f => !have.has(String(f.id))).map(f => [
    String(f.id), now, me.name || me.email, FILES.AREAS[area], ref, refLabel, String(f.name), String(f.url || ''), Number(f.size) || 0, String(f.mime || ''), memo || '', '', viewYear_(),
  ]);
  if (rows.length) sh.getRange(last + 1, 1, rows.length, rows[0].length).setValues(rows);
  dropCache_(CK.files);
}

function readFiles_(ss) {
  const sh = ss.getSheetByName(FILES.SHEET);
  const areaKey = {};
  Object.keys(FILES.AREAS).forEach(k => { areaKey[FILES.AREAS[k]] = k; });
  const year = viewYear_();
  return dataRows_(sh, FCOL.YEAR)
    .filter(r => String(r[0]) && String(r[FCOL.DELETED - 1]).toUpperCase() !== 'Y')
    .filter(r => !cellText_(r[FCOL.YEAR - 1]) || cellText_(r[FCOL.YEAR - 1]) === year)
    .map(r => ({
      id: String(r[0]),
      date: r[1] instanceof Date ? fmt(r[1], 'yyyy-MM-dd HH:mm') : String(r[1]),
      uploader: String(r[2]),
      area: areaKey[String(r[3])] || 'etc',
      ref: cellText_(r[4]),
      refLabel: String(r[5]),
      name: String(r[6]),
      url: String(r[7]),
      size: Number(r[8]) || 0,
      mime: String(r[9]),
      memo: String(r[10]),
    }))
    .reverse();
}

/** 첨부 삭제: 올린 사람 또는 총괄·관리자. 드라이브 파일은 휴지통으로 (30일 안에 복구 가능) */
function apiDeleteFile(id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'files');
  const sh = ctx.ss.getSheetByName(FILES.SHEET);
  const rows = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 3).getValues();
  const idx = rows.findIndex(r => String(r[0]) === String(id));
  if (idx < 0) throw new Error('파일을 찾을 수 없습니다.');
  if (String(rows[idx][2]) !== me.name && !me.isLeader && !me.isAdmin) throw new Error('본인이 올린 파일만 삭제할 수 있습니다.');
  sh.getRange(idx + 2, FCOL.DELETED).setValue('Y');
  try { DriveApp.getFileById(String(id)).setTrashed(true); } catch (e) { /* 이미 지워졌거나 권한 없음: 목록에서만 뺀다 */ }
  return refreshPart_('files');
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
