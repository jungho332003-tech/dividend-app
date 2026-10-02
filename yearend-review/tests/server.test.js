// 서버 시나리오 시험: python3 build.py 후 node tests/server.test.js
const { makeEnv, assert } = require('./gas_sim');
let pass = 0;
const ok = (name, fn) => { try { fn(); pass++; console.log('✓', name); } catch (e) { console.log('✗', name, '\n   ', e.message); process.exitCode = 1; } };
const throwsMsg = (fn, re) => { let err; try { fn(); } catch (e) { err = e; } assert(err, 'expected error'); assert(re.test(err.message), 'wrong error: ' + err.message); };

/* ---------- 1. 새 스프레드시트 ---------- */
{
  const { api, ss, state, books } = makeEnv();
  const as = email => { state.user = email; };
  ok('initialize creates sheets', () => {
    api.initialize();
    ['설정', '담당자', '대상자', '응대기록', '공지', '공지댓글', '기준', '기준개정이력', '첨부', '메뉴권한'].forEach(n => assert(ss.getSheetByName(n), n));
    const hdr = ss.getSheetByName('대상자').getRange(1, 1, 1, 40).getValues()[0].filter(Boolean);
    assert.strictEqual(hdr.length, 30);
    assert.strictEqual(hdr[0], 'No.'); assert.strictEqual(hdr[13], '시스템 등록'); assert.strictEqual(hdr[28], '최종수정');
    assert(!hdr.includes('수기서류 제출') && !hdr.includes('월세액') && hdr.includes('종전근무지 여부'));
    assert.strictEqual(ss.getSheetByName('기준').getLastRow(), 7);
  });
  ok('initialize twice is safe', () => { api.initialize(); assert.strictEqual(ss.getSheetByName('대상자').getLastColumn(), 30); assert.strictEqual(ss.getSheetByName('기준').getLastRow(), 7); });

  // 담당자 등록 (소유자 owner@x.com = 관리자)
  const mem = ss.getSheetByName('담당자');
  ok('new member rows default to 전체 (blank)', () => { assert.strictEqual(mem.get(3, 6), ''); });
  mem.getRange(2, 1, 5, 6).clearContent();
  mem.getRange(2, 1, 4, 6).setValues([[1, '정해린', 'lead@x.com', '총괄', 'N', ''], [2, '김도현', 'kim@x.com', '응대담당', 'N', '전체'], [3, '이수민', 'lee@x.com', '응대담당', 'N', ''], [4, '최유나', 'choi@x.com', '2차검토', 'N', '전체']]);
  ss.getSheetByName('설정').getRange('B4').setValue('FOLDER');

  ok('unregistered user has no menus', () => { as('stranger@x.com'); const b = api.apiBootstrap(); assert.strictEqual(b.me.menus.length, 0); assert.strictEqual(b.people.length, 0); });
  ok('owner is admin', () => { as('owner@x.com'); const b = api.apiBootstrap(); assert(b.me.isAdmin); assert(b.me.menus.includes('admin')); });

  ok('응대담당 cannot import', () => { as('kim@x.com'); throwsMsg(() => api.apiImportPeople([{ name: 'x' }]), /총괄 또는 관리자/); });
  ok('총괄 imports (bool O, multiline)', () => {
    as('lead@x.com');
    const r = api.apiImportPeople([
      { no: '1', empNo: '1001', name: '오민재', dept: '인사팀', owner: '김도현', owner2: '최유나', ehr: 'O', arrived: '○', mortPrev: '2019-03-15', note1: '줄1\n줄2' },
      { no: '2', empNo: '1002', name: '서지안', dept: '재무팀', owner: '이수민', owner2: '최유나', arrived: 'Y' },
      { no: '3', empNo: '1003', name: '전시우', dept: '품질팀' },
    ]);
    assert.strictEqual(r.added, 3); assert.strictEqual(r.people.length, 3);
    const p = r.people[0];
    assert.strictEqual(p.ehr, true); assert.strictEqual(p.arrived, true); assert.strictEqual(p.verified, false);
    assert.strictEqual(p.note1, '줄1\n줄2'); assert.strictEqual(p.editor, '정해린');
    const sh = ss.getSheetByName('대상자');
    assert.strictEqual(sh.get(2, 14), 'O'); // 시스템 등록 열(14번째)
  });
  ok('re-import updates by 사원번호', () => { const r = api.apiImportPeople([{ empNo: '1002', name: '서지안', dept: '재무2팀' }]); assert.strictEqual(r.updated, 1); assert.strictEqual(r.people.length, 3); assert.strictEqual(r.people[1].dept, '재무2팀'); assert.strictEqual(r.people[1].owner, '이수민'); });

  ok('응대담당 edits own person (first fields)', () => {
    as('kim@x.com');
    const r = api.apiSavePerson('1001', { verified: true, note1: '확인', missing: '' });
    assert.strictEqual(r.personPatch.verified, true); assert.strictEqual(r.personPatch.note1, '확인'); assert.strictEqual(r.personPatch.editor, '김도현');
  });
  ok('응대담당 blocked on others / second / assign', () => {
    throwsMsg(() => api.apiSavePerson('1002', { verified: true }), /담당자만/);
    throwsMsg(() => api.apiSavePerson('1001', { note2: 'x' }), /2차검토 담당자만/);
    throwsMsg(() => api.apiSavePerson('1001', { owner2: '이수민' }), /총괄만/);
  });
  ok('응대담당 takes unassigned person', () => {
    const r = api.apiSavePerson('1003', { owner: '김도현', ehr: true });
    assert.strictEqual(r.personPatch.owner, '김도현'); assert.strictEqual(r.personPatch.ehr, true);
    throwsMsg(() => { as('lee@x.com'); api.apiSavePerson('1003', { owner: '이수민' }); }, /총괄만/);
  });
  ok('2차검토 edits second fields only', () => {
    as('choi@x.com');
    const r = api.apiSavePerson('1001', { review2: true, note2: '2차 확인' });
    assert.strictEqual(r.personPatch.review2, true);
    throwsMsg(() => api.apiSavePerson('1001', { verified: false }), /담당자만/);
  });
  ok('총괄 edits anything, empNo change + dup check', () => {
    as('lead@x.com');
    api.apiSavePerson('1002', { owner: '김도현', verified: true });
    throwsMsg(() => api.apiSavePerson('1002', { empNo: '1001' }), /이미 있습니다/);
    const r = api.apiSavePerson('1002', { empNo: '2002' });
    assert.strictEqual(r.personPatch.id, '2002'); assert.strictEqual(r.oldId, '1002');
  });
  ok('add/delete person', () => {
    as('kim@x.com'); throwsMsg(() => api.apiAddPerson({ name: 'x' }), /총괄 또는 관리자/);
    as('lead@x.com');
    let r = api.apiAddPerson({ empNo: '1004', name: '홍나래' }); assert.strictEqual(r.people.length, 4);
    r = api.apiDeletePerson('1004'); assert.strictEqual(r.people.length, 3);
  });

  ok('응대기록', () => {
    as('lee@x.com'); // 담당이 아니어도 기록 가능
    const r = api.apiAddLog('1001', '전화', '영수증 요청');
    assert.strictEqual(r.logAdded.name, '오민재'); assert.strictEqual(r.logAdded.author, '이수민');
    assert.strictEqual(api.apiBootstrap().logs[0].body, '영수증 요청');
    throwsMsg(() => api.apiAddLog('1001', '전화', ' '), /내용/);
  });

  ok('notices: roles, read marks, files registry', () => {
    as('kim@x.com');
    throwsMsg(() => api.apiSaveNotice({ category: '공지', title: 't', body: 'b' }), /총괄 또는 관리자/);
    let r = api.apiSaveNotice({ category: '질문', level: '긴급', pinned: true, title: '질문', body: '내용' });
    const q = r.notices.find(n => n.id === r.savedId);
    assert.strictEqual(q.level, '일반'); assert.strictEqual(q.pinned, false); assert(r.reads.includes(r.savedId));
    as('lead@x.com');
    throwsMsg(() => api.apiSaveNotice({ category: '공지', title: 't', body: 'b', start: '2027-02-10', end: '2027-02-01' }), /종료일/);
    r = api.apiSaveNotice({ category: '공지', level: '긴급', pinned: true, title: '마감', body: '2/10', start: '2027-02-01', end: '2027-02-10', files: [{ id: 'X1', name: 'a.pdf', url: 'u', size: 1 }] });
    const n = r.notices[0];
    assert.strictEqual(n.title, '마감'); assert.strictEqual(n.level, '긴급'); assert.strictEqual(n.start, '2027-02-01'); assert.strictEqual(n.files.length, 1);
    assert.strictEqual(ss.getSheetByName('첨부').getLastRow(), 2);
    // 같은 첨부로 다시 저장해도 중복 기록 안 됨
    api.apiSaveNotice({ id: n.id, category: '공지', level: '중요', title: '마감(수정)', body: '2/10', files: [{ id: 'X1', name: 'a.pdf', url: 'u', size: 1 }] });
    assert.strictEqual(ss.getSheetByName('첨부').getLastRow(), 2);
    as('kim@x.com');
    throwsMsg(() => api.apiSaveNotice({ id: n.id, category: '질문', title: 'x', body: 'y' }), /본인 글만/);
    assert(!api.apiBootstrap().reads.includes(n.id));
    api.apiMarkRead(n.id); assert(api.apiBootstrap().reads.includes(n.id));
    api.apiMarkUnread(n.id); assert(!api.apiBootstrap().reads.includes(n.id));
    r = api.apiAddNoticeComment(n.id, '확인'); assert.strictEqual(r.commentAdded.comment.body, '확인');
    assert.strictEqual(api.apiBootstrap().notices.find(x => x.id === n.id).comments.length, 1);
  });

  ok('rules', () => {
    as('kim@x.com'); throwsMsg(() => api.apiSaveRule({ title: 't', body: 'b' }), /총괄 또는 관리자/);
    as('lead@x.com');
    let r = api.apiSaveRule({ category: '교육비', title: '교육비', body: '15%', year: '2026' });
    assert.strictEqual(r.rules.length, 7);
    r = api.apiSaveRule({ id: r.savedId, category: '교육비', title: '교육비', body: '15% 수정', year: '2026', changeNote: '한도 추가' });
    const x = r.rules.find(y => y.id === r.savedId);
    assert.strictEqual(x.history.length, 2); assert.strictEqual(x.history[0].note, '한도 추가');
    r = api.apiDeleteRule(x.id); assert.strictEqual(r.rules.length, 6);
  });

  ok('uploads + delete rights', () => {
    as('kim@x.com');
    const data = Buffer.from('pdfdata').toString('base64');
    let r = api.apiUploadFile({ name: '오민재_간소화.pdf', mimeType: 'application/pdf', data, area: 'person', ref: '1001', refLabel: '오민재(1001)' });
    assert(r.file.id); assert.strictEqual(r.files[0].area, 'person'); assert.strictEqual(r.files[0].ref, '1001');
    const fid = r.file.id;
    r = api.apiUploadFile({ name: '양식.xlsx', data, area: 'etc', memo: '양식' });
    assert.strictEqual(r.files[0].memo, '양식');
    throwsMsg(() => api.apiUploadFile({ name: 'big', data: Buffer.alloc(11 * 1024 * 1024).toString('base64'), area: 'etc' }), /10MB/);
    as('lee@x.com'); throwsMsg(() => api.apiDeleteFile(fid), /본인이 올린/);
    as('kim@x.com'); r = api.apiDeleteFile(fid); assert(!r.files.some(f => f.id === fid));
  });

  ok('menu access blocks server side', () => {
    as('kim@x.com'); throwsMsg(() => api.apiAdminData(), /관리자만/);
    as('owner@x.com');
    const d = api.apiAdminData(false);
    assert.strictEqual(d.members.length, 4);
    const acc = d.menuAccess; acc['응대담당'].files = false; acc['응대담당'].rules = false;
    api.apiSaveMenuAccess(acc);
    as('kim@x.com');
    const b = api.apiBootstrap();
    assert(!b.me.menus.includes('files')); assert(!b.me.menus.includes('rules'));
    assert.strictEqual(b.rules.length, 0);
    assert(b.files.every(f => f.area === 'person' || f.area === 'notice'), 'only visible areas');
    throwsMsg(() => api.apiUploadFile({ name: 'a', data: 'YQ==', area: 'etc' }), /첨부파일\] 메뉴/);
  });

  ok('열람 범위: 본인 / 본인+담당자 / 담당자 없는 대상자는 모두 / 전체', () => {
    as('owner@x.com');
    const d = api.apiAdminData(false);
    assert(d.owners['김도현'] >= 1);
    const save = scopes => api.apiSaveMembers(d.members.map(m => Object.assign({}, m, { scope: scopes[m.name] || '전체' })));
    // 사람 현황: 1001 김도현/최유나, 2002 김도현, 1003 김도현 → 1002 바꿔 이수민 담당 하나와 미배정 하나 만든다
    as('lead@x.com');
    api.apiSavePerson('2002', { owner: '이수민', owner2: '' });
    api.apiAddPerson({ empNo: '1005', name: '미배정씨' });
    as('owner@x.com');
    save({ '김도현': '본인' });
    assert.strictEqual(api.apiAdminData(false).members.find(m => m.name === '김도현').scope, '본인');
    as('kim@x.com');
    let b = api.apiBootstrap();
    assert.deepStrictEqual(b.people.map(p => p.empNo).sort().join(), '1001,1003,1005'); // 1005는 담당자 없음 → 보임
    assert(b.logs.every(l => ['1001', '1003', '1005'].includes(l.empNo)));
    assert(b.files.every(f => f.area !== 'person' || ['1001', '1003', '1005'].includes(f.ref)));
    throwsMsg(() => api.apiSavePerson('2002', { note1: 'x' }), /열람 범위/);
    throwsMsg(() => api.apiAddLog('2002', '전화', 'x'), /열람 범위/);
    throwsMsg(() => api.apiUploadFile({ name: 'a.pdf', data: 'YQ==', area: 'person', ref: '2002' }), /열람 범위/);
    as('owner@x.com'); save({ '김도현': '본인,이수민' });
    as('kim@x.com');
    b = api.apiBootstrap();
    assert.strictEqual(b.me.scope, '본인,이수민');
    assert.deepStrictEqual(b.people.map(p => p.empNo).sort().join(), '1001,1003,1005,2002');
    api.apiSavePerson('1005', { owner: '김도현' }); // 담당자 없는 대상자는 누구나 맡을 수 있다
    as('owner@x.com'); save({ '김도현': '' });
    as('kim@x.com'); assert.strictEqual(api.apiBootstrap().me.scope, '전체'); // 설정 안 하면 전체
    as('choi@x.com'); assert.strictEqual(api.apiBootstrap().people.length, 4); // 전체
    as('lead@x.com'); assert.strictEqual(api.apiBootstrap().me.scope, '전체'); // 총괄은 항상 전체
    as('owner@x.com'); save({});
  });

  ok('양식 다운로드: 총괄만, 한 번 만들고 다시 씀, 명단 채우기는 열람 범위대로', () => {
    as('kim@x.com'); throwsMsg(() => api.apiTemplateLink(false), /총괄 또는 관리자/);
    as('lead@x.com');
    const r1 = api.apiTemplateLink(false);
    assert(/\/export\?format=xlsx$/.test(r1.url)); assert.strictEqual(r1.count, 0);
    const book = books[Object.keys(books)[0]];
    const sh = book.getSheetByName('대상자');
    assert.strictEqual(sh.get(1, 1), 'No.'); assert.strictEqual(sh.get(1, 4), '성명'); assert.strictEqual(sh.getLastColumn(), 28);
    assert(book.getSheetByName('작성 안내'));
    const r2 = api.apiTemplateLink(true);
    assert.strictEqual(Object.keys(books).length, 1, 'reuses the same template');
    assert.strictEqual(r2.count, api.apiBootstrap().people.length);
    assert.strictEqual(sh.getLastRow(), r2.count + 1);
  });

  ok('admin save members / settings / apply', () => {
    as('owner@x.com');
    throwsMsg(() => api.apiSaveMembers([{ name: 'A', email: 'a@x.com' }, { name: 'A', email: 'b@x.com' }]), /겹칩니다/);
    throwsMsg(() => api.apiSaveMembers([{ name: 'A', email: 'bad' }]), /이메일 형식/);
    const d = api.apiSaveMembers([{ name: '정해린', email: 'lead@x.com', role: '총괄', admin: true }, { name: '김도현', email: 'kim@x.com', role: '응대담당' }, { name: '최유나', email: 'choi@x.com', role: '2차검토' }]);
    assert.strictEqual(d.members.length, 3); assert.strictEqual(d.members[0].admin, true);
    throwsMsg(() => api.apiSaveSettings({ team: 't', year: '26', folderId: '' }), /네 자리/);
    const s = api.apiSaveSettings({ team: '연말정산', year: '2026', folderId: 'https://drive.google.com/drive/folders/FOLDER?usp=sharing' });
    assert.strictEqual(s.settings.folderId, 'FOLDER'); assert.strictEqual(s.status.resources.folder.ok, true);
    const res = api.apiApplyPermissions();
    assert(res.log.some(l => /스프레드시트 편집 권한/.test(l.text)));
    assert(ss.getSheetByName('설정').getProtections().length === 1);
    assert(ss.getSheetByName('대상자').getProtections()[0].isWarningOnly());
    as('lee@x.com'); assert.strictEqual(api.apiBootstrap().me.menus.length, 0); // 담당자에서 빠짐
  });

  ok('cache chunking works for big lists', () => {
    const big = Array.from({ length: 3000 }, (_, i) => ({ i, s: '가나다라마바사아자차'.repeat(5) }));
    api.putCache_('bigkey', big);
    let built = false;
    const got = api.cached_('bigkey', () => { built = true; return []; });
    assert(!built); assert.strictEqual(got.length, 3000);
  });
}

/* ---------- 2. 기존 시트를 그대로 쓰는 경우 (열 순서 다름, 체크박스, 24년도 열 이름) ---------- */
{
  const { api, ss, state } = makeEnv();
  const sh = ss.insertSheet('대상자');
  const hdr = ['No.', '부서', '사원번호', '성명', '사원하위그룹명', '급여영역', '직급', '전화번호', '이메일주소', '휴직여부', '담당자', '2차검토 담당자', '2차 서류검토 여부', 'E-HR 등록', '서류 도착여부', '수기서류 제출', '서류확인 및 검증', '종전근무지', '종전근무지 개수', '주택임차차입금\n(신청여부)', '주택임차차입금', '장기주택\n저당차입금\n(신청여부)', '장기주택저당차입금', '장기주택\n국세청자료 여부', '24년도 장기주택 공제여부\n(차입일)', '주택마련저축\n(신청여부)', '주택마련저축', '주택마련저축\n국세청자료 여부', '월세액\n(신청여부)', '월세액', '특이사항', '수정사항', '특이사항(2차)', '수정사항(2차)', '미비서류'];
  sh.getRange(1, 1, 1, hdr.length).setValues([hdr]);
  const row = new Array(hdr.length).fill('');
  Object.assign(row, { 0: 1, 1: '인사팀', 2: 7001, 3: '기존직원', 10: '김도현', 11: '최유나', 13: true, 14: false, 17: '가나상사', 18: 1, 24: new api.CtxDate('2019-03-15T00:00:00+09:00'), 30: '기존 특이' });
  sh.getRange(2, 1, 1, hdr.length).setValues([row]);
  ok('existing sheet: only meta columns appended', () => {
    api.initialize();
    assert.strictEqual(sh.getLastColumn(), 37);
    assert.strictEqual(sh.get(1, 36), '최종수정');
  });
  const mem = ss.getSheetByName('담당자');
  mem.getRange(2, 1, 5, 5).clearContent();
  mem.getRange(2, 1, 1, 5).setValues([[1, '김도현', 'kim@x.com', '응대담당', 'N']]);
  ok('existing sheet: read with header matching (줄바꿈 제목, 24년도, 숫자 사번, 날짜)', () => {
    state.user = 'kim@x.com';
    const p = api.apiBootstrap().people[0];
    assert.strictEqual(p.id, '7001'); assert.strictEqual(p.ehr, true); assert.strictEqual(p.arrived, false);
    assert.strictEqual(p.prevWork, true, '회사명이 적힌 종전근무지 = 있음'); assert.strictEqual(p.note1, '기존 특이');
    assert.strictEqual(p.mortPrev, undefined); assert.strictEqual(p.manual, undefined);
  });
  ok('existing sheet: checkbox cells stay boolean on save', () => {
    const r = api.apiSavePerson('7001', { arrived: true, verified: true });
    assert.strictEqual(sh.get(2, 15), true); // 체크박스 칸 → true
    assert.strictEqual(sh.get(2, 17), 'O');  // 빈 칸이었던 곳 → O
    assert.strictEqual(r.personPatch.arrived, true);
    api.apiSavePerson('7001', { prevWork: true, note1: '저장' });
    assert.strictEqual(sh.get(2, 18), '가나상사', '체크 유지 시 회사명 그대로');
    api.apiSavePerson('7001', { prevWork: false });
    assert.strictEqual(sh.get(2, 18), '');
  });
}
console.log(`\n${pass} passed${process.exitCode ? ' (with failures)' : ''}`);
