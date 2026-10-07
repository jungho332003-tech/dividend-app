/**
 * 대상자(서류 검토) / 응대기록
 *
 * [대상자]   기존 검토 시트의 열 그대로 (COLS) + 최종수정 | 수정자
 * [응대기록] 일시 | 사원번호 | 성명 | 작성자 | 구분 | 내용
 */

const LOG_KINDS = ['전화', '메일', '메신저', '방문', '보완요청', '기타'];

function setupPeopleSheets_(ss, year) {
  let sh = yearSheet_(ss, SHEET.PEOPLE, year);
  if (!sh) {
    sh = ss.insertSheet(`${SHEET.PEOPLE}_${year}`);
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

  if (!yearSheet_(ss, SHEET.LOGS, year)) {
    const lg = ss.insertSheet(`${SHEET.LOGS}_${year}`);
    lg.getRange(1, 1, 1, 6).setValues([['일시', '사원번호', '성명', '작성자', '구분', '내용']]);
    styleHeader_(lg.getRange(1, 1, 1, 6));
    lg.getRange('A2:A').setNumberFormat('yyyy-mm-dd hh:mm');
    lg.getRange('F2:F').setWrap(true);
    [130, 90, 80, 80, 80, 480].forEach((w, i) => lg.setColumnWidth(i + 1, w));
    lg.setFrozenRows(1);
  }
}

/** 새로 생긴 열(예: 퇴사여부)이 시트에 없으면 오른쪽 끝에 제목만 붙인다. 「기본 시트 만들기」를 다시 안 해도 저장된다 */
function ensureCol_(sh, map, col) {
  if (map[col.k]) return map[col.k];
  const at = sh.getLastColumn() + 1;
  sh.getRange(1, at).setValue(col.l);
  styleHeader_(sh.getRange(1, at));
  map[col.k] = at;
  return at;
}

/** 1행 제목 → { 키: 열번호(1부터) } */
function colMap_(sh) {
  const last = sh.getLastColumn();
  return last ? colMapFrom_(sh.getRange(1, 1, 1, last).getValues()[0]) : {};
}

/** 이미 읽은 제목 행으로 열 위치 찾기 */
function colMapFrom_(headerRow) {
  const map = {};
  const header = headerRow.map(normHeader_);
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
    p[c.k] = c.t === 'bool' ? boolOf_(c, v) : cellText_(v);
  });
  p.updated = map.updated && r[map.updated - 1] instanceof Date ? fmt(r[map.updated - 1], 'yyyy-MM-dd HH:mm') : cellText_(map.updated ? r[map.updated - 1] : '');
  p.editor = cellText_(map.editor ? r[map.editor - 1] : '');
  // 사원번호가 없으면 행 번호로 구분
  p.id = p.empNo || 'r' + row;
  return p;
}

function readPeople_(ss) {
  const sh = yearSheet_(ss, SHEET.PEOPLE);
  if (!sh) return [];
  const v = sh.getDataRange().getValues(); // 제목과 내용을 한 번에
  if (v.length < 2) return [];
  const map = colMapFrom_(v[0]);
  return v.slice(1)
    .map((r, i) => rowToPerson_(r, map, i + 2))
    .filter(p => p.name || p.empNo);
}

function readLogs_(ss) {
  return dataRows_(yearSheet_(ss, SHEET.LOGS), 6)
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
  if (col.who === 'lead') return !!(me.isLeader || me.isAdmin);
  if (me.isLeader || me.isAdmin) return true;
  if (!me.name) return false;
  // 담당자가 비어 있으면 본인을 1차 담당자로 지정할 수 있다 (같은 저장에서 다른 칸도 함께 고칠 수 있다)
  const takes = !cur.owner && patch.owner === me.name;
  if (col.who === 'assign') return col.k === 'owner' && takes;
  if (col.who === 'second') return cur.owner2 === me.name;
  return cur.owner === me.name || takes;
}

function whoLabel_(col) {
  return { lead: '총괄·관리자', second: '2차검토 담당자', assign: '총괄' }[col.who] || '담당자';
}

/** 저장할 값. 체크박스 칸(지금 값이 true/false)이면 체크박스로, 아니면 기존 시트처럼 O로 쓴다 */
function cellValue_(col, now, value) {
  // 종전근무지처럼 회사명이 적힌 칸은 체크를 유지할 때 그 글자를 그대로 둔다
  if (col.loose && value && boolOf_(col, now) && typeof now !== 'boolean') return now;
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
    const sh = peopleSheetOrThrow_(ctx.ss);
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
        throw new Error(`[${col.l}] 칸은 ${whoLabel_(col)}만 고칠 수 있습니다.`);
      }
      ensureCol_(sh, map, col);
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
    const sh = peopleSheetOrThrow_(ctx.ss);
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
    const sh = peopleSheetOrThrow_(ctx.ss);
    const map = colMap_(sh);
    COLS.forEach(c => { if (!map[c.k] && (rows || []).some(src => c.k in src)) ensureCol_(sh, map, c); });
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
        r[map[c.k] - 1] = c.t === 'bool' ? cellValue_(c, old, boolOf_(c, src[c.k])) : String(src[c.k] == null ? '' : src[c.k]).trim();
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
  const sh = peopleSheetOrThrow_(ctx.ss);
  const map = colMap_(sh);
  const found = readPersonRow_(sh, map, String(id));
  if (!found) throw new Error('대상자를 찾을 수 없습니다.');
  const p = rowToPerson_(found.values, map, found.row);
  if (!canSee_(me, p)) throw new Error('열람 범위 밖의 대상자입니다.');
  const now = new Date();
  const k = LOG_KINDS.indexOf(kind) >= 0 ? kind : '기타';
  yearSheet_(ctx.ss, SHEET.LOGS).appendRow([now, p.empNo, p.name, me.name || me.email, k, text]);
  // 응대기록 시트 전체를 다시 읽지 않고 캐시 맨 앞에 넣는다
  const entry = { date: fmt(now, 'yyyy-MM-dd HH:mm'), empNo: p.empNo, name: p.name, author: me.name || me.email, kind: k, body: text };
  patchCache_(CK.logs, list => { list.unshift(entry); });
  return { logAdded: entry };
}

/**
 * 대상자 업로드 양식 (총괄·관리자)
 * 접속한 사람의 드라이브에 "연말정산 대상자 업로드 양식" 스프레드시트를 하나 만들어 두고(다음부터는 다시 씀),
 * 열 제목·담당자 목록을 최신으로 맞춘 뒤 엑셀(.xlsx)로 내려받는 주소를 돌려준다.
 * withPeople=true면 지금 볼 수 있는 대상자 명단을 채워서 준다 (엑셀에서 고쳐 다시 올리기용).
 */
function apiTemplateLink(withPeople) {
  const ctx = getContext();
  const me = currentMember_(ctx.members);
  requireMenu_(me, 'review');
  if (!me.isLeader && !me.isAdmin) throw new Error('양식 다운로드는 총괄 또는 관리자만 할 수 있습니다.');

  const props = PropertiesService.getUserProperties();
  let ss = null;
  const id = props.getProperty('TEMPLATE_ID');
  if (id) {
    try {
      const f = DriveApp.getFileById(id);
      if (!f.isTrashed()) ss = SpreadsheetApp.openById(id);
    } catch (e) { ss = null; /* 지워졌거나 권한 없음: 새로 만든다 */ }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('연말정산 대상자 업로드 양식');
    props.setProperty('TEMPLATE_ID', ss.getId());
  }

  const header = COLS.map(c => c.l);
  let sh = ss.getSheetByName('대상자') || ss.getSheets()[0].setName('대상자');
  sh.clear();
  sh.getDataRange().clearDataValidations();
  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setBackground('#efefef').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(1, 42);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(4);
  COLS.forEach((c, i) => sh.setColumnWidth(i + 1, c.t === 'long' ? 220 : c.t === 'bool' ? 90 : 110));

  const ROWS = 1000;
  const boolRule = SpreadsheetApp.newDataValidation().requireValueInList(['O', 'X'], true).setAllowInvalid(true).build();
  COLS.forEach((c, i) => { if (c.t === 'bool') sh.getRange(2, i + 1, ROWS, 1).setDataValidation(boolRule).setHorizontalAlignment('center'); });
  const staff = ctx.members.map(m => m.name);
  if (staff.length) {
    const staffRule = SpreadsheetApp.newDataValidation().requireValueInList(staff, true).setAllowInvalid(true).build();
    ['owner', 'owner2'].forEach(k => sh.getRange(2, COLS.findIndex(c => c.k === k) + 1, ROWS, 1).setDataValidation(staffRule));
  }
  sh.getRange(2, COLS.findIndex(c => c.k === 'empNo') + 1, ROWS, 1).setNumberFormat('@'); // 사원번호 앞자리 0 유지

  let count = 0;
  if (withPeople) {
    const people = cached_(CK.people, () => readPeople_(ctx.ss)).filter(p => canSee_(me, p));
    const rows = people.map(p => COLS.map(c => c.t === 'bool' ? (p[c.k] ? 'O' : '') : p[c.k]));
    if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
    count = rows.length;
  }

  let guide = ss.getSheetByName('작성 안내') || ss.insertSheet('작성 안내');
  guide.clear();
  const lines = [
    ['연말정산 대상자 업로드 양식 작성 안내'],
    [''],
    ['· [대상자] 시트 2행부터 한 사람씩 적고, 웹앱 서류 검토 > 대상자 업로드에 이 파일을 올리세요.'],
    ['· 성명은 꼭 적어야 합니다. 사원번호가 이미 있으면 그 사람 정보를 갱신하고, 없으면 새로 추가합니다.'],
    ['· 빈 칸은 기존 값을 그대로 둡니다. 필요 없는 열은 비워 두거나 지워도 됩니다.'],
    ['· 체크 항목(휴직여부, 퇴사여부, 시스템 등록, 서류 도착여부, 종전근무지 여부 등)은 O로 적습니다. 이미 체크된 것을 풀려면 X로 적습니다.'],
    [`· 담당자·2차검토 담당자는 [담당자] 목록의 이름과 똑같이 적어야 내 담당으로 연결됩니다: ${staff.join(', ') || '(담당자 없음)'}`],
    ['· 열 순서는 바꿔도 됩니다. 1행 제목(열 이름)으로 맞춥니다.'],
    [''],
    ['예시 (이 시트는 올려도 읽지 않습니다)'],
  ];
  guide.getRange(1, 1, lines.length, 1).setValues(lines);
  guide.getRange('A1').setFontSize(13).setFontWeight('bold');
  guide.getRange(lines.length + 1, 1, 2, 8).setValues([
    ['No.', '부서', '사원번호', '성명', '직급', '담당자', '2차검토 담당자', '서류 도착여부'],
    ['1', '인사팀', '20110321', '홍길동', '과장', staff[0] || '김담당', staff[1] || '이검토', 'O'],
  ]);
  guide.getRange(lines.length + 1, 1, 1, 8).setBackground('#efefef').setFontWeight('bold');
  guide.setColumnWidth(1, 120);
  ss.setActiveSheet(sh);
  SpreadsheetApp.flush();

  const stamp = fmt(new Date(), 'yyyyMMdd');
  return {
    url: `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=xlsx`,
    sheetUrl: ss.getUrl(),
    fileName: `연말정산_대상자_${withPeople ? '명단' : '양식'}_${stamp}.xlsx`,
    count: count,
  };
}

/** 보고 있는 연도의 대상자 시트 (없으면 안내) */
function peopleSheetOrThrow_(ss) {
  const sh = yearSheet_(ss, SHEET.PEOPLE);
  if (!sh) throw new Error(`${viewYear_()}년 대상자 시트가 없습니다. 관리자에게 「새 연도 시작」을 요청하세요.`);
  return sh;
}
