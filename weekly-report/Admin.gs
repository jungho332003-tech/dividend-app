/**
 * 관리 메뉴 (관리자 전용): 팀원 관리, 권한 한 번에 적용, 메뉴 권한, 설정
 *
 * 관리자 = 스프레드시트 소유자 + [팀원] 시트에서 관리자(Y)인 사람
 *
 * 권한 적용이 하는 일
 *  1. 스프레드시트를 팀원에게 편집자로 공유
 *  2. 첨부 폴더를 편집자로 공유
 *  3. 팀 캘린더를 "일정 변경" 권한으로 공유 (고급 서비스 Calendar 필요)
 *  4. 입력시트 생성 + 시트 보호 (본인/팀장/관리자 구분)
 */

const MENUS = [
  { key: 'dash', label: '대시보드' },
  { key: 'weekly', label: '주간업무보고' },
  { key: 'budget', label: '예산전용' },
  { key: 'calendar', label: '일정' },
  { key: 'board', label: '공유게시판' },
  { key: 'rules', label: '회사 기준' },
];
const ROLES = ['팀원', '팀장'];
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

/** { 팀원: {dash:true, ...}, 팀장: {...} } */
function getMenuAccess_() {
  return cached_('menus', () => {
    const access = {};
    ROLES.forEach(r => { access[r] = {}; MENUS.forEach(m => { access[r][m.key] = true; }); });
    const sh = SpreadsheetApp.getActive().getSheetByName(MENU_SHEET);
    if (!sh || sh.getLastRow() < 2) return access;
    const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
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

/**
 * useCache=true면 공유 상태(시트·폴더·캘린더 확인, 가장 느린 부분)를 5분 캐시에서 쓴다.
 */
function adminData_(ctx, useCache) {
  const cfg = ctx.cfg;
  return {
    members: ctx.members.map(m => ({ order: m.order, name: m.name, email: m.email, role: m.role, write: m.write, admin: m.admin })),
    menus: MENUS,
    roles: ROLES,
    menuAccess: getMenuAccess_(),
    settings: {
      team: String(cfg[CFG.TEAM] || ''),
      deadlineHour: Number(cfg[CFG.DEADLINE_HOUR]) || 14,
      remindHours: Number(cfg[CFG.REMIND_HOURS]) || 3,
      calendarId: String(cfg[CFG.CALENDAR] || ''),
      folderId: String(cfg[CFG.DRIVE_FOLDER] || ''),
    },
    webAppUrl: ScriptApp.getService().getUrl() || '',
    status: useCache ? cached_('access', () => accessStatus_(ctx), 300) : freshStatus_(ctx),
  };
}

function freshStatus_(ctx) {
  const status = accessStatus_(ctx);
  putCache_('access', status, 300);
  return status;
}

/**
 * 팀원별 접근 상태
 * sheet/folder: 'editor' | 'viewer' | 'none' | 'owner'   calendar: 'writer' | 'reader' | 'none' | 'owner'
 * 설정이 없거나 확인할 수 없으면 resources[].ok=false
 */
function accessStatus_(ctx) {
  const lower = list => list.map(u => u.getEmail().toLowerCase());
  const res = { sheet: { ok: true, name: ctx.ss.getName() }, folder: { ok: false }, calendar: { ok: false } };

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
    res.folder = { ok: false, error: '첨부 폴더 ID가 비어 있습니다.' };
  }

  let acl = {};
  const calId = String(ctx.cfg[CFG.CALENDAR] || '').trim();
  if (calId) {
    try {
      const cal = CalendarApp.getCalendarById(calId);
      if (!cal) throw new Error('not found');
      res.calendar = { ok: true, name: cal.getName() };
      (Calendar.Acl.list(calId).items || []).forEach(rule => {
        if (rule.scope && rule.scope.type === 'user') acl[String(rule.scope.value).toLowerCase()] = rule.role;
      });
    } catch (e) {
      res.calendar = {
        ok: false,
        error: /Calendar is not defined/.test(String(e))
          ? 'Apps Script에서 고급 서비스 "Google Calendar API"를 켜야 합니다.'
          : '캘린더를 찾을 수 없거나 공유 설정을 볼 권한이 없습니다.',
      };
    }
  } else {
    res.calendar = { ok: false, error: '팀 캘린더 ID가 비어 있습니다.' };
  }

  const level = (email, owner, editors, viewers) =>
    email === owner ? 'owner' : editors.indexOf(email) >= 0 ? 'editor' : viewers.indexOf(email) >= 0 ? 'viewer' : 'none';

  const people = {};
  ctx.members.filter(m => m.email).forEach(m => {
    const e = m.email.toLowerCase();
    const sh = m.write ? ctx.ss.getSheetByName(SHEET.INPUT_PREFIX + m.name) : null;
    let inputProtected = false;
    if (sh) {
      const p = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
      inputProtected = !!p && p.getEditors().some(u => u.getEmail().toLowerCase() === e);
    }
    people[m.name] = {
      sheet: level(e, fileOwner, fileEditors, fileViewers),
      input: !m.write ? 'n/a' : !sh ? 'missing' : inputProtected ? 'ok' : 'unprotected',
      folder: res.folder.ok ? level(e, folderOwner, folderEditors, folderViewers) : 'n/a',
      calendar: res.calendar.ok ? (acl[e] || 'none') : 'n/a',
    };
  });
  return { resources: res, people: people };
}

/* ---------- 저장 ---------- */

/** 팀원 목록 전체 저장. list: [{order, name, email, role, write, admin}] */
function apiSaveMembers(list) {
  const ctx = getContext();
  requireAdmin_(ctx);
  ensureHeader_(ctx.ss.getSheetByName(SHEET.MEMBERS), 6, '관리자(Y/N)');

  const rows = (list || []).map((m, i) => ({
    order: Number(m.order) || i + 1,
    name: String(m.name || '').trim(),
    email: String(m.email || '').trim().toLowerCase(),
    role: ROLES.indexOf(m.role) >= 0 ? m.role : '팀원',
    write: !!m.write,
    admin: !!m.admin,
  })).filter(m => m.name);

  const names = {};
  rows.forEach(m => {
    if (names[m.name]) throw new Error(`이름이 겹칩니다: ${m.name}. 동명이인은 "김민지A"처럼 구분해 주세요.`);
    names[m.name] = true;
    if (m.email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(m.email)) throw new Error(`이메일 형식을 확인해 주세요: ${m.name} (${m.email})`);
  });

  const sh = ctx.ss.getSheetByName(SHEET.MEMBERS);
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 6).setValues(rows.map(m => [m.order, m.name, m.email, m.role, m.write ? 'Y' : 'N', m.admin ? 'Y' : 'N']));
  }
  clearCaches_(ctx);
  // 저장만 하고 공유 상태는 다시 확인하지 않는다 (확인은 "권한 한 번에 적용"이나 "상태 새로고침"에서)
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
  try { CacheService.getDocumentCache().remove('menus'); } catch (e) { /* 캐시 없음 */ }
  return { menuAccess: getMenuAccess_() };
}

function apiSaveSettings(s) {
  const ctx = getContext();
  requireAdmin_(ctx);
  const hour = Math.round(Number(s.deadlineHour));
  const remind = Math.round(Number(s.remindHours));
  if (!(hour >= 6 && hour <= 22)) throw new Error('마감 시각은 6~22시 사이로 입력해 주세요.');
  if (!(remind >= 0 && remind <= 48)) throw new Error('미작성 알림은 0~48시간 전으로 입력해 주세요.');
  setConfigValue(CFG.TEAM, String(s.team || '').trim());
  setConfigValue(CFG.DEADLINE_HOUR, hour);
  setConfigValue(CFG.REMIND_HOURS, remind);
  setConfigValue(CFG.CALENDAR, String(s.calendarId || '').trim());
  setConfigValue(CFG.DRIVE_FOLDER, String(s.folderId || '').trim().replace(/^.*\/folders\//, '').replace(/[?#].*$/, ''));
  clearCaches_(ctx);
  return adminData_(getContext(), false);
}

function clearCaches_(ctx) {
  try {
    CacheService.getDocumentCache().removeAll([CK.members(ctx), CK.budget, CK.board, CK.rules, 'menus', CTX_KEY]);
  } catch (e) { /* 캐시 없음 */ }
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
    warn('첨부 폴더 ID가 비어 있어 폴더 공유를 건너뛰었습니다.');
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

  // 3. 팀 캘린더
  const calId = String(ctx.cfg[CFG.CALENDAR] || '').trim();
  if (!calId) {
    warn('팀 캘린더 ID가 비어 있어 캘린더 공유를 건너뛰었습니다.');
  } else {
    try {
      const acl = {};
      (Calendar.Acl.list(calId).items || []).forEach(r => {
        if (r.scope && r.scope.type === 'user') acl[String(r.scope.value).toLowerCase()] = r.role;
      });
      people.forEach(m => {
        const e = m.email.toLowerCase();
        if (acl[e] === 'owner' || acl[e] === 'writer') return;
        try {
          Calendar.Acl.insert({ role: 'writer', scope: { type: 'user', value: e } }, calId, { sendNotifications: false });
          ok(`${m.name}: 팀 캘린더 "일정 변경" 권한을 줬습니다.`);
        } catch (err) {
          warn(`${m.name}: 캘린더 공유 실패 (${err.message})`);
        }
      });
    } catch (err) {
      warn(/Calendar is not defined/.test(String(err))
        ? '캘린더 공유를 하려면 Apps Script에서 고급 서비스 "Google Calendar API"를 켜야 합니다.'
        : '팀 캘린더 공유 설정을 바꿀 수 없습니다. 캘린더 소유자가 실행해 주세요.');
    }
  }

  // 4. 입력시트 + 시트 보호
  try {
    const missing = applyProtections_(ctx);
    ok('입력시트와 시트 보호를 적용했습니다.');
    missing.forEach(n => warn(`${n}: 이메일이 없어 입력시트를 본인 전용으로 잠그지 못했습니다.`));
  } catch (err) {
    warn(`시트 보호 적용 실패 (${err.message})`);
  }

  clearCaches_(ctx);
  if (!log.some(l => l.ok && /권한을 줬습니다/.test(l.text))) ok('새로 줄 공유 권한은 없었습니다. 모두 이미 공유되어 있습니다.');
  return { log: log, data: adminData_(getContext(), false) };
}

/** 팀원을 뺄 때 시트·폴더·캘린더 접근 해제 (소유자와 본인은 건드리지 않는다) */
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

  const calId = String(ctx.cfg[CFG.CALENDAR] || '').trim();
  if (calId) {
    try {
      Calendar.Acl.remove(calId, 'user:' + e);
      log.push({ ok: true, text: '팀 캘린더 공유를 해제했습니다.' });
    } catch (err) {
      log.push({ ok: false, text: '팀 캘린더 공유 해제 실패 (이미 해제되었거나 권한 없음)' });
    }
  }
  return { log: log, data: adminData_(getContext(), false) };
}
