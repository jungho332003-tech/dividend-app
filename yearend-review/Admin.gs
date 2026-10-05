/**
 * 관리 메뉴 (관리자 전용): 담당자 관리, 권한 한 번에 적용, 메뉴 권한, 설정
 *
 * 관리자 = 스프레드시트 소유자 + [담당자] 시트에서 관리자(Y)인 사람
 *
 * 권한 적용이 하는 일
 *  1. 스프레드시트를 담당자에게 편집자로 공유
 *  2. 첨부 폴더를 편집자로 공유
 *  3. 시트 보호 (설정·담당자는 관리자만, 기준은 총괄·관리자만)
 */

const MENUS = [
  { key: 'dash', label: '대시보드' },
  { key: 'review', label: '서류 검토' },
  { key: 'notice', label: '공지사항' },
  { key: 'rules', label: '연말정산 기준' },
  { key: 'files', label: '첨부파일' },
];
const ROLES = ['응대담당', '2차검토', '총괄'];
const LEADER_ROLE = '총괄';
const MENU_SHEET = '메뉴권한';

/* ---------- 메뉴 권한 ---------- */

function setupMenuSheet_(ss) {
  if (ss.getSheetByName(MENU_SHEET)) return;
  const sh = ss.insertSheet(MENU_SHEET);
  const rows = [['메뉴', '키'].concat(ROLES)].concat(MENUS.map(m => [m.label, m.key].concat(ROLES.map(() => 'Y'))));
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  styleHeader_(sh.getRange(1, 1, 1, rows[0].length));
  sh.getRange(2, 3, MENUS.length, ROLES.length).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Y', 'N']).build());
  sh.hideColumns(2);
  sh.setFrozenRows(1);
  sh.getRange('A1').setNote('역할별로 볼 수 있는 메뉴 (Y/N). 관리자는 항상 모든 메뉴를 봅니다.');
}

/** { 응대담당: {dash:true, ...}, ... } */
function getMenuAccess_() {
  return cached_('menus', () => {
    const access = {};
    ROLES.forEach(r => { access[r] = {}; MENUS.forEach(m => { access[r][m.key] = true; }); });
    const sh = SpreadsheetApp.getActive().getSheetByName(MENU_SHEET);
    if (!sh) return access;
    const v = sh.getDataRange().getValues();
    if (v.length < 2) return access;
    const header = v[0].map(String);
    v.slice(1).forEach(row => {
      const key = String(row[1]);
      ROLES.forEach(r => {
        const c = header.indexOf(r);
        if (c >= 0 && key) access[r][key] = String(row[c]).toUpperCase() !== 'N';
      });
    });
    return access;
  });
}

/** 접속자가 볼 수 있는 메뉴 키 목록 */
function allowedMenus_(me) {
  if (me.isAdmin) return MENUS.map(m => m.key).concat('admin');
  if (!me.role) return [];
  const access = getMenuAccess_()[me.role] || {};
  return MENUS.map(m => m.key).filter(k => access[k] !== false);
}

function requireMenu_(me, key) {
  if (allowedMenus_(me).indexOf(key) < 0) {
    const label = (MENUS.find(m => m.key === key) || { label: key }).label;
    throw new Error(`[${label}] 메뉴를 사용할 권한이 없습니다. 관리자에게 문의하세요.`);
  }
}

function requireAdmin_(ctx) {
  const me = currentMember_(ctx.members);
  if (!me.isAdmin) throw new Error('관리자만 사용할 수 있습니다.');
  return me;
}

/* ---------- 조회 ---------- */

/** force=true면 공유 상태를 새로 확인 (상태 새로고침 버튼) */
function apiAdminData(force) {
  const ctx = getContext();
  requireAdmin_(ctx);
  return adminData_(ctx, !force);
}

function adminData_(ctx, useCache) {
  const cfg = ctx.cfg;
  return {
    members: ctx.members.map(m => ({ order: m.order, name: m.name, email: m.email, role: m.role, admin: m.admin, scope: m.scope })),
    // 열람 범위 화면용: 담당자 칸에 적힌 이름별 대상자 수와, 사람별 [1차, 2차] 담당 (이름만, 개인정보 없음)
    owners: ownerCounts_(ctx),
    assign: readPeople_(ctx.ss).map(p => [p.owner, p.owner2]),
    menus: MENUS,
    roles: ROLES,
    menuAccess: getMenuAccess_(),
    settings: {
      team: String(cfg[CFG.TEAM] || ''),
      year: ctx.baseYear,
      years: ctx.years,
      folderId: String(cfg[CFG.DRIVE_FOLDER] || ''),
    },
    webAppUrl: ScriptApp.getService().getUrl() || '',
    status: useCache ? cached_('access', () => accessStatus_(ctx), 300) : freshStatus_(ctx),
  };
}

function ownerCounts_(ctx) {
  const counts = {};
  readPeople_(ctx.ss).forEach(p => { const k = p.owner || SCOPE_UNASSIGNED; counts[k] = (counts[k] || 0) + 1; });
  return counts;
}

function freshStatus_(ctx) {
  const status = accessStatus_(ctx);
  putCache_('access', status, 300);
  return status;
}

/**
 * 담당자별 접근 상태
 * sheet/folder: 'editor' | 'viewer' | 'none' | 'owner' ('n/a' = 폴더 설정 필요)
 */
function accessStatus_(ctx) {
  const lower = list => list.map(u => u.getEmail().toLowerCase());
  const res = { sheet: { ok: true, name: ctx.ss.getName() }, folder: { ok: false } };

  const file = DriveApp.getFileById(ctx.ss.getId());
  const fileOwner = file.getOwner() ? file.getOwner().getEmail().toLowerCase() : '';
  const fileEditors = lower(file.getEditors());
  const fileViewers = lower(file.getViewers());

  let folderEditors = [], folderViewers = [], folderOwner = '';
  const folderId = String(ctx.cfg[CFG.DRIVE_FOLDER] || '').trim();
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      folderOwner = folder.getOwner() ? folder.getOwner().getEmail().toLowerCase() : '';
      folderEditors = lower(folder.getEditors());
      folderViewers = lower(folder.getViewers());
      res.folder = { ok: true, name: folder.getName() };
    } catch (e) {
      res.folder = { ok: false, error: '폴더를 찾을 수 없거나 접근 권한이 없습니다.' };
    }
  } else {
    res.folder = { ok: false, error: '첨부 폴더가 비어 있습니다.' };
  }

  const level = (email, owner, editors, viewers) =>
    email === owner ? 'owner' : editors.indexOf(email) >= 0 ? 'editor' : viewers.indexOf(email) >= 0 ? 'viewer' : 'none';

  const people = {};
  ctx.members.filter(m => m.email).forEach(m => {
    const e = m.email.toLowerCase();
    people[m.name] = {
      sheet: level(e, fileOwner, fileEditors, fileViewers),
      folder: res.folder.ok ? level(e, folderOwner, folderEditors, folderViewers) : 'n/a',
    };
  });
  return { resources: res, people: people };
}

/* ---------- 저장 ---------- */

/** 담당자 목록 전체 저장. list: [{order, name, email, role, admin, scope}] */
function apiSaveMembers(list) {
  const ctx = getContext();
  requireAdmin_(ctx);

  const rows = (list || []).map((m, i) => ({
    order: Number(m.order) || i + 1,
    name: String(m.name || '').trim(),
    email: String(m.email || '').trim().toLowerCase(),
    role: ROLES.indexOf(m.role) >= 0 ? m.role : ROLES[0],
    admin: !!m.admin,
    scope: normScope_(m.scope),
  })).filter(m => m.name);

  const names = {};
  rows.forEach(m => {
    if (names[m.name]) throw new Error(`이름이 겹칩니다: ${m.name}. 동명이인은 "김민지A"처럼 구분해 주세요. ([대상자] 시트의 담당자 칸도 같은 이름으로 적어야 합니다)`);
    names[m.name] = true;
    if (m.email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(m.email)) throw new Error(`이메일 형식을 확인해 주세요: ${m.name} (${m.email})`);
  });

  const sh = ctx.ss.getSheetByName(SHEET.MEMBERS);
  ensureHeader_(sh, 6, '열람범위');
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 6).setValues(rows.map(m => [m.order, m.name, m.email, m.role, m.admin ? 'Y' : 'N', m.scope]));
  }
  clearCaches_();
  return adminData_(getContext(), true);
}

function apiSaveMenuAccess(access) {
  const ctx = getContext();
  requireAdmin_(ctx);
  setupMenuSheet_(ctx.ss);
  const sh = ctx.ss.getSheetByName(MENU_SHEET);
  const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const header = v[0].map(String);
  const out = v.slice(1).map(row => ROLES.map(r => {
    const allowed = access && access[r] ? access[r][String(row[1])] !== false : true;
    return allowed ? 'Y' : 'N';
  }));
  const firstRoleCol = header.indexOf(ROLES[0]) + 1;
  sh.getRange(2, firstRoleCol, out.length, ROLES.length).setValues(out);
  dropCache_('menus');
  return { menuAccess: getMenuAccess_() };
}

function apiSaveSettings(s) {
  const ctx = getContext();
  requireAdmin_(ctx);
  const year = String(s.year || '').trim();
  if (!/^\d{4}$/.test(year)) throw new Error('기본 연도는 2026처럼 네 자리로 입력해 주세요.');
  if (year !== ctx.baseYear) {
    if (ctx.years.indexOf(year) < 0) throw new Error(`${year}년 시트가 아직 없습니다. 「새 연도 시작」으로 만들어 주세요.`);
    // 예전 이름의 시트가 남아 있으면 지금 기본 연도 이름으로 먼저 고정한다
    migrateLegacySheets_(ctx.ss, ctx.baseYear);
  }
  setConfigValue(CFG.TEAM, String(s.team || '').trim());
  setConfigValue(CFG.YEAR, Number(year));
  setConfigValue(CFG.DRIVE_FOLDER, String(s.folderId || '').trim().replace(/^.*\/folders\//, '').replace(/[?#].*$/, ''));
  clearCaches_();
  return adminData_(getContext(), false);
}

/* ---------- 새 연도 시작 ---------- */

/**
 * opts: { year, carry: 이전 연도 대상자의 기본 정보·담당 가져오기, copyRules: 기준 복사, makeBase: 기본 연도로 지정 }
 * 대상자_YYYY·응대기록_YYYY·일정_YYYY 시트를 만든다. 이전 연도 시트는 그대로 남아 연도 선택으로 볼 수 있다.
 */
function apiStartYear(opts) {
  const ctx = getContext();
  const me = requireAdmin_(ctx);
  const year = String(opts && opts.year || '').trim();
  if (!/^\d{4}$/.test(year)) throw new Error('연도는 2027처럼 네 자리로 입력해 주세요.');
  if (ctx.years.indexOf(year) >= 0 && ctx.ss.getSheetByName(`${SHEET.PEOPLE}_${year}`)) throw new Error(`${year}년은 이미 있습니다.`);
  const from = String(opts.from || ctx.baseYear);

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  let carried = 0, copied = 0;
  try {
    migrateLegacySheets_(ctx.ss, ctx.baseYear);
    setupPeopleSheets_(ctx.ss, year);
    setupEventSheet_(ctx.ss, year);

    // 이전 연도 명단: 기본 정보와 담당 배정만 가져오고, 진행 체크·특이사항은 비운다
    if (opts.carry) {
      const src = ctx.ss.getSheetByName(`${SHEET.PEOPLE}_${from}`);
      const dst = ctx.ss.getSheetByName(`${SHEET.PEOPLE}_${year}`);
      if (src && src.getLastRow() >= 2) {
        const smap = colMap_(src), dmap = colMap_(dst);
        const keep = COLS.filter(c => c.who === 'info' || c.who === 'assign');
        const rows = src.getRange(2, 1, src.getLastRow() - 1, src.getLastColumn()).getValues()
          .filter(r => keep.some(c => smap[c.k] && String(r[smap[c.k] - 1]).trim()))
          .map(r => {
            const out = new Array(dst.getLastColumn()).fill('');
            keep.forEach(c => { if (smap[c.k] && dmap[c.k]) out[dmap[c.k] - 1] = c.t === 'bool' ? (boolOf_(c, r[smap[c.k] - 1]) ? 'O' : '') : r[smap[c.k] - 1]; });
            return out;
          });
        if (rows.length) dst.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
        carried = rows.length;
      }
    }

    // 기준 복사: 이전 연도 기준을 새 연도로 한 벌 더 만든다 (개정 이력은 새로 시작)
    if (opts.copyRules) {
      const sh = ctx.ss.getSheetByName(RULES.SHEET);
      const hsh = ctx.ss.getSheetByName(RULES.HISTORY);
      if (sh && sh.getLastRow() >= 2) {
        const all = sh.getRange(2, 1, sh.getLastRow() - 1, RCOL.FILES).getValues();
        let next = Math.max.apply(null, all.map(r => Number(r[0]) || 0)) + 1;
        const now = new Date();
        const rows = all.filter(r => String(r[0]) && String(r[RCOL.DELETED - 1]).toUpperCase() !== 'Y' && cellText_(r[RCOL.YEAR - 1]) === from)
          .map(r => { const x = r.slice(); x[0] = next++; x[RCOL.YEAR - 1] = year; x[RCOL.UPDATED - 1] = now; x[RCOL.EDITOR - 1] = me.name || me.email; return x; });
        if (rows.length) {
          sh.getRange(sh.getLastRow() + 1, 1, rows.length, RCOL.FILES).setValues(rows);
          hsh.getRange(hsh.getLastRow() + 1, 1, rows.length, 4).setValues(rows.map(x => [String(x[0]), now, me.name || me.email, `${from}년 기준에서 복사`]));
        }
        copied = rows.length;
      }
    }

    if (opts.makeBase !== false) setConfigValue(CFG.YEAR, Number(year));
    try { applyProtections_(getContext()); } catch (e) { /* 보호는 권한 한 번에 적용에서 다시 걸 수 있다 */ }
  } finally {
    lock.releaseLock();
  }
  clearCaches_();
  // 시작한 관리자는 바로 새 연도를 보게 한다
  PropertiesService.getUserProperties().setProperty('VIEW_YEAR', year);
  VIEW_MEMO_ = null;
  return { carried: carried, copied: copied, data: adminData_(getContext(), true) };
}

/* ---------- 권한 한 번에 적용 ---------- */

/** 결과: { log: [{ok, text}], data } */
function apiApplyPermissions() {
  const ctx = getContext();
  requireAdmin_(ctx);
  const log = [];
  const ok = text => log.push({ ok: true, text: text });
  const warn = text => log.push({ ok: false, text: text });
  const people = ctx.members.filter(m => m.email);
  ctx.members.filter(m => !m.email).forEach(m => warn(`${m.name}: 이메일이 없어 건너뛰었습니다.`));

  // 1. 스프레드시트
  const file = DriveApp.getFileById(ctx.ss.getId());
  const fileOwner = file.getOwner() ? file.getOwner().getEmail().toLowerCase() : '';
  const fileEditors = file.getEditors().map(u => u.getEmail().toLowerCase());
  people.forEach(m => {
    const e = m.email.toLowerCase();
    if (e === fileOwner || fileEditors.indexOf(e) >= 0) return;
    try { file.addEditor(e); ok(`${m.name}: 스프레드시트 편집 권한을 줬습니다.`); }
    catch (err) { warn(`${m.name}: 스프레드시트 공유 실패 (${err.message})`); }
  });

  // 2. 첨부 폴더
  const folderId = String(ctx.cfg[CFG.DRIVE_FOLDER] || '').trim();
  if (!folderId) {
    warn('첨부 폴더가 비어 있어 폴더 공유를 건너뛰었습니다. (관리 > 설정)');
  } else {
    try {
      const folder = DriveApp.getFolderById(folderId);
      const owner = folder.getOwner() ? folder.getOwner().getEmail().toLowerCase() : '';
      const editors = folder.getEditors().map(u => u.getEmail().toLowerCase());
      people.forEach(m => {
        const e = m.email.toLowerCase();
        if (e === owner || editors.indexOf(e) >= 0) return;
        try { folder.addEditor(e); ok(`${m.name}: 첨부 폴더 편집 권한을 줬습니다.`); }
        catch (err) { warn(`${m.name}: 첨부 폴더 공유 실패 (${err.message})`); }
      });
    } catch (err) {
      warn('첨부 폴더에 접근할 수 없어 폴더 공유를 건너뛰었습니다.');
    }
  }

  // 3. 시트 보호
  try {
    applyProtections_(ctx);
    ok('시트 보호를 적용했습니다. (설정·담당자: 관리자만 / 기준: 총괄·관리자만)');
  } catch (err) {
    warn(`시트 보호 적용 실패 (${err.message})`);
  }

  clearCaches_();
  if (!log.some(l => l.ok && /권한을 줬습니다/.test(l.text))) ok('새로 줄 공유 권한은 없었습니다. 모두 이미 공유되어 있습니다.');
  return { log: log, data: adminData_(getContext(), false) };
}

/** 담당자를 뺄 때 시트·폴더 접근 해제 (소유자와 본인은 건드리지 않는다) */
function apiRevokeAccess(email) {
  const ctx = getContext();
  const me = requireAdmin_(ctx);
  const e = String(email || '').trim().toLowerCase();
  if (!e) throw new Error('이메일이 없습니다.');
  if (e === me.email) throw new Error('본인의 접근은 해제할 수 없습니다.');
  const log = [];

  const file = DriveApp.getFileById(ctx.ss.getId());
  if (file.getOwner() && file.getOwner().getEmail().toLowerCase() === e) throw new Error('스프레드시트 소유자의 접근은 해제할 수 없습니다.');
  try { file.removeEditor(e); } catch (err) { /* 편집자가 아님 */ }
  try { file.removeViewer(e); } catch (err) { /* 뷰어가 아님 */ }
  log.push({ ok: true, text: '스프레드시트 접근을 해제했습니다.' });

  const folderId = String(ctx.cfg[CFG.DRIVE_FOLDER] || '').trim();
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      try { folder.removeEditor(e); } catch (err) { /* 편집자가 아님 */ }
      try { folder.removeViewer(e); } catch (err) { /* 뷰어가 아님 */ }
      log.push({ ok: true, text: '첨부 폴더 접근을 해제했습니다.' });
    } catch (err) {
      log.push({ ok: false, text: '첨부 폴더 접근 해제 실패' });
    }
  }
  return { log: log, data: adminData_(getContext(), false) };
}
