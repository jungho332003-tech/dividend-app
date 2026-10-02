/**
 * 대상자(서류 검토) / 응대기록
 *
 * [대상자]   기존 검토 시트의 열 그대로 (COLS) + 최종수정 | 수정자
 * [응대기록] 일시 | 사원번호 | 성명 | 작성자 | 구분 | 내용
 */

const LOG_KINDS = ['전화', '메일', '메신저', '방문', '보완요청', '기타'];

function setupPeopleSheets_(ss) {
  let sh = ss.getSheetByName(SHEET.PEOPLE);
  if (!sh) {
    sh = ss.insertSheet(SHEET.PEOPLE);
    const header = COLS.concat(META_COLS).map(c => c.l);
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    styleHeader_(sh.getRange(1, 1, 1, header.length));
    sh.getRange(1, 1, 1, header.length).setWrap(true);
    sh.setRowHeight(1, 42);
    sh.setFrozenRows(1);
    sh.setFrozenColumns(4);
    COLS.forEach((c, i) => {
      if (c.t === 'long') sh.setColumnWidth(i + 1, 220);
      else if (c.t === 'bool') sh.setColumnWidth(i + 1, 80);
    });
  } else {
    // 기존 시트를 그대로 쓴다: 없는 열만 오른쪽 끝에 추가
    const map = colMap_(sh);
    let last = sh.getLastColumn();
    COLS.concat(META_COLS).forEach(c => {
      if (map[c.k]) return;
      last += 1;
      sh.getRange(1, last).setValue(c.l);
      styleHeader_(sh.getRange(1, last));
    });
  }

  if (!ss.getSheetByName(SHEET.LOGS)) {
    const lg = ss.insertSheet(SHEET.LOGS);
    lg.getRange(1, 1, 1, 6).setValues([['일시', '사원번호', '성명', '작성자', '구분', '내용']]);
    styleHeader_(lg.getRange(1, 1, 1, 6));
    lg.getRange('A2:A').setNumberFormat('yyyy-mm-dd hh:mm');
    lg.getRange('F2:F').setWrap(true);
    [130, 90, 80, 80, 80, 480].forEach((w, i) => lg.setColumnWidth(i + 1, w));
    lg.setFrozenRows(1);
  }
}

/** 1행 제목 → { 키: 열번호(1부터) } */
function colMap_(sh) {
  const last = sh.getLastColumn();
  const map = {};
  if (!last) return map;
  const header = sh.getRange(1, 1, 1, last).getValues()[0].map(normHeader_);
  COLS.concat(META_COLS).forEach(c => {
    const want = normHeader_(c.l);
    const re = c.re ? new RegExp(c.re) : null;
    const i = header.findIndex(h => h === want || (re && re.test(h)));
    if (i >= 0) map[c.k] = i + 1;
  });
  return map;
}

function rowToPerson_(r, map, row) {
  const p = { row: row };
  COLS.forEach(c => {
    const v = map[c.k] ? r[map[c.k] - 1] : '';
    p[c.k] = c.t === 'bool' ? toBool_(v) : cellText_(v);
  });
  p.updated = map.updated && r[map.updated - 1] instanceof Date ? fmt(r[map.updated - 1], 'yyyy-MM-dd HH:mm') : cellText_(map.updated ? r[map.updated - 1] : '');
  p.editor = cellText_(map.editor ? r[map.editor - 1] : '');
  // 사원번호가 없으면 행 번호로 구분
  p.id = p.empNo || 'r' + row;
  return p;
}

function readPeople_(ss) {
  const sh = ss.getSheetByName(SHEET.PEOPLE);
  if (!sh || sh.getLastRow() < 2) return [];
  const map = colMap_(sh);
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
    .map((r, i) => rowToPerson_(r, map, i + 2))
    .filter(p => p.name || p.empNo);
}

function readLogs_(ss) {
  const sh = ss.getSheetByName(SHEET.LOGS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues()
    .filter(r => String(r[5]).trim())
    .map(r => ({
      date: r[0] instanceof Date ? fmt(r[0], 'yyyy-MM-dd HH:mm') : String(r[0]),
      empNo: cellText_(r[1]),
      name: String(r[2]),
      author: String(r[3]),
      kind: String(r[4]),
      body: String(r[5]),
    }))
    .reverse();
}

/**
 * id로 대상자 행을 찾아 그 행 값을 함께 돌려준다.
 * 캐시에 있는 행 번호부터 확인하고(읽기 1번), 맞지 않으면 사원번호 열에서 찾는다.
 */
function readPersonRow_(sh, map, id) {
  const width = sh.getLastColumn();
  const hit = (peekCache_(CK.people) || []).find(p => p.id === id);
  if (hit && hit.row >= 2) {
    const values = sh.getRange(hit.row, 1, 1, width).getValues()[0];
    const emp = map.empNo ? cellText_(values[map.empNo - 1]) : '';
    if ((emp || 'r' + hit.row) === id) return { row: hit.row, values: values };
  }
  const row = findPersonRow_(sh, map, id);
  return row < 0 ? null : { row: row, values: sh.getRange(row, 1, 1, width).getValues()[0] };
}

/** id(사원번호 또는 'r행번호')로 행 찾기 */
function findPersonRow_(sh, map, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  if (/^r\d+$/.test(id)) {
    const row = Number(id.slice(1));
    const emp = map.empNo ? cellText_(sh.getRange(row, map.empNo).getValue()) : '';
    return row <= last && !emp ? row : -1;
  }
  if (!map.empNo) return -1;
  const emps = sh.getRange(2, map.empNo, last - 1, 1).getValues().map(([v]) => cellText_(v));
  const i = emps.indexOf(String(id));
  return i < 0 ? -1 : i + 2;
}

/** 이 사람이 이 대상자의 이 칸을 고칠 수 있나 */
function canEdit_(me, cur, col, patch) {
  if (me.isLeader || me.isAdmin) return true;
  if (!me.name) return false;
  // 담당자가 비어 있으면 본인을 1차 담당자로 지정할 수 있다 (같은 저장에서 다른 칸도 함께 고칠 수 있다)
  const takes = !cur.owner && patch.owner === me.name;
  if (col.who === 'assign') return col.k === 'owner' && takes;
  if (col.who === 'second') return cur.owner2 === me.name;
  return cur.owner === me.name || takes;
}

/** 저장할 값. 체크박스 칸(지금 값이 true/false)이면 체크박스로, 아니면 기존 시트처럼 O로 쓴다 */
function cellValue_(col, now, value) {
  if (col.t === 'bool') return typeof now === 'boolean' ? !!value : (value ? 'O' : '');
  return String(value == null ? '' : value).trim();
}

/** 바뀐 칸만 저장한다 (1차·2차 담당자가 동시에 고쳐도 서로 덮어쓰지 않게) */
function apiSavePerson(id, patch) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'review');

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sh = ctx.ss.getSheetByName(SHEET.PEOPLE);
    const map = colMap_(sh);
    const found = readPersonRow_(sh, map, String(id));
    if (!found) throw new Error('대상자를 찾을 수 없습니다. 시트에서 사원번호가 바뀌었을 수 있으니 새로고침해 주세요.');
    const { row, values } = found;
    const cur = rowToPerson_(values, map, row);
    if (!canSee_(me, cur)) throw new Error('열람 범위 밖의 대상자입니다. 관리자에게 열람 범위를 요청하세요.');

    const keys = Object.keys(patch || {}).filter(k => COLS.some(c => c.k === k));
    if (!keys.length) return { personPatch: cur };
    keys.forEach(k => {
      const col = COLS.find(c => c.k === k);
      if (!canEdit_(me, cur, col, patch)) {
        throw new Error(`[${col.l}] 칸은 ${col.who === 'second' ? '2차검토 담당자' : col.who === 'assign' ? '총괄' : '담당자'}만 고칠 수 있습니다.`);
      }
      if (!map[k]) throw new Error(`[대상자] 시트에 "${col.l}" 열이 없습니다. 관리자에게 "기본 시트 만들기"를 다시 실행해 달라고 요청하세요.`);
    });
    if (keys.indexOf('empNo') >= 0) {
      const emp = String(patch.empNo || '').trim();
      if (emp && emp !== cur.empNo && findPersonRow_(sh, map, emp) > 0) throw new Error(`사원번호 ${emp}가 이미 있습니다.`);
    }
    // 바뀐 칸만 쓴다 (수식이 있는 다른 칸은 건드리지 않는다). 쓰기만 이어서 하면 Apps Script가 한 번에 보낸다
    const now = new Date();
    keys.forEach(k => {
      const v = cellValue_(COLS.find(c => c.k === k), values[map[k] - 1], patch[k]);
      sh.getRange(row, map[k]).setValue(v);
      values[map[k] - 1] = v;
    });
    if (map.updated) { sh.getRange(row, map.updated).setValue(now); values[map.updated - 1] = now; }
    if (map.editor) { sh.getRange(row, map.editor).setValue(me.name || me.email); values[map.editor - 1] = me.name || me.email; }

    // 다시 읽지 않고 저장한 값으로 결과를 만들고, 캐시의 그 사람만 바꾼다
    const saved = rowToPerson_(values, map, row);
    patchCache_(CK.people, list => {
      const i = list.findIndex(p => p.id === String(id));
      if (i >= 0) list[i] = saved; else list.push(saved);
    });
    return { personPatch: saved, oldId: String(id) };
  } finally {
    lock.releaseLock();
  }
}

/** 총괄·관리자: 대상자 한 명 추가 */
function apiAddPerson(fields) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'review');
  if (!me.isLeader && !me.isAdmin) throw new Error('대상자 추가는 총괄 또는 관리자만 할 수 있습니다.');
  if (!String(fields.name || '').trim()) throw new Error('성명을 입력해 주세요.');
  apiImportPeople([fields]);
  return refreshPart_('people');
}

/** 총괄·관리자: 대상자 삭제 (행 삭제) */
function apiDeletePerson(id) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'review');
  if (!me.isLeader && !me.isAdmin) throw new Error('대상자 삭제는 총괄 또는 관리자만 할 수 있습니다.');
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sh = ctx.ss.getSheetByName(SHEET.PEOPLE);
    const row = findPersonRow_(sh, colMap_(sh), String(id));
    if (row < 0) throw new Error('대상자를 찾을 수 없습니다.');
    sh.deleteRow(row);
  } finally {
    lock.releaseLock();
  }
  return refreshPart_('people');
}

/**
 * 총괄·관리자: 시트에서 복사한 내용 붙여넣기. rows: [{empNo, name, ...}]
 * 사원번호가 같으면 붙여넣은 칸만 갱신, 없으면 추가. 시트 요청은 읽기 1번 + 쓰기 1번.
 */
function apiImportPeople(rows) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'review');
  if (!me.isLeader && !me.isAdmin) throw new Error('붙여넣기는 총괄 또는 관리자만 할 수 있습니다.');

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  let added = 0, updated = 0;
  try {
    const sh = ctx.ss.getSheetByName(SHEET.PEOPLE);
    const map = colMap_(sh);
    const width = sh.getLastColumn();
    const last = sh.getLastRow();
    const data = last >= 2 ? sh.getRange(2, 1, last - 1, width).getValues() : [];
    const byEmp = {};
    data.forEach((r, i) => { const e = map.empNo ? cellText_(r[map.empNo - 1]) : ''; if (e) byEmp[e] = i; });
    const now = new Date();

    (rows || []).forEach(src => {
      if (!String(src.name || '').trim()) return;
      const emp = String(src.empNo || '').trim();
      let r;
      if (emp && byEmp[emp] !== undefined) { r = data[byEmp[emp]]; updated++; }
      else { r = new Array(width).fill(''); data.push(r); if (emp) byEmp[emp] = data.length - 1; added++; }
      COLS.forEach(c => {
        if (!(c.k in src) || !map[c.k]) return;
        const old = r[map[c.k] - 1];
        r[map[c.k] - 1] = c.t === 'bool' ? (typeof old === 'boolean' ? toBool_(src[c.k]) : (toBool_(src[c.k]) ? 'O' : '')) : String(src[c.k] == null ? '' : src[c.k]).trim();
      });
      if (map.updated) r[map.updated - 1] = now;
      if (map.editor) r[map.editor - 1] = me.name || me.email;
    });
    if (data.length) sh.getRange(2, 1, data.length, width).setValues(data);
  } finally {
    lock.releaseLock();
  }
  dropCache_(CK.people);
  return Object.assign(refreshPart_('people'), { added: added, updated: updated });
}

/** 응대기록 추가: 누구든 응대한 사람이 남긴다 */
function apiAddLog(id, kind, body) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'review');
  if (!me.name && !me.isAdmin) throw new Error('[담당자] 시트에 등록된 사람만 응대기록을 남길 수 있습니다.');
  const text = String(body || '').trim();
  if (!text) throw new Error('응대 내용을 입력해 주세요.');
  const sh = ctx.ss.getSheetByName(SHEET.PEOPLE);
  const map = colMap_(sh);
  const found = readPersonRow_(sh, map, String(id));
  if (!found) throw new Error('대상자를 찾을 수 없습니다.');
  const p = rowToPerson_(found.values, map, found.row);
  if (!canSee_(me, p)) throw new Error('열람 범위 밖의 대상자입니다.');
  const now = new Date();
  const k = LOG_KINDS.indexOf(kind) >= 0 ? kind : '기타';
  ctx.ss.getSheetByName(SHEET.LOGS).appendRow([now, p.empNo, p.name, me.name || me.email, k, text]);
  // 응대기록 시트 전체를 다시 읽지 않고 캐시 맨 앞에 넣는다
  const entry = { date: fmt(now, 'yyyy-MM-dd HH:mm'), empNo: p.empNo, name: p.name, author: me.name || me.email, kind: k, body: text };
  patchCache_(CK.logs, list => { list.unshift(entry); });
  return { logAdded: entry };
}
