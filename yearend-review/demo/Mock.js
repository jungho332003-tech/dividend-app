/* 예시 화면 전용 가짜 서버 (실제 웹앱에서는 Apps Script가 대신한다). build.py가 dist/demo.html을 만들 때 App.html의 <!-- MOCK --> 자리에 넣는다. */
window.MockApi = (function () {
  const pad = n => String(n).padStart(2, '0');
  const NOW = new Date('2027-02-04T10:00:00+09:00');
  let tick = 0;
  // 보는 사람의 시간대와 상관없이 한국 시간으로 찍는다
  const stamp = () => { const d = new Date(NOW.getTime() + (tick++) * 60000 + 9 * 3600000); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())} ${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}`; };
  const F = (name, size) => ({ id: 'demo-' + name, name, url: '#demo-file', size, mime: '' });
  let seq = 100;

  const ROLES = ['응대담당', '2차검토', '총괄'];
  const MENUS = [
    { key: 'dash', label: '대시보드' }, { key: 'review', label: '서류 검토' }, { key: 'notice', label: '공지사항' },
    { key: 'rules', label: '연말정산 기준' }, { key: 'files', label: '첨부파일' },
  ];
  const staff = [
    { name: '정해린', email: 'haerin.jung@example.com', role: '총괄', admin: true, scope: '전체' },
    { name: '김도현', email: 'dohyun.kim@example.com', role: '응대담당', admin: false, scope: '본인' },
    { name: '이수민', email: 'sumin.lee@example.com', role: '응대담당', admin: false, scope: '본인,박지호' },
    { name: '박지호', email: 'jiho.park@example.com', role: '응대담당', admin: false, scope: '본인' },
    { name: '최유나', email: 'yuna.choi@example.com', role: '2차검토', admin: false, scope: '전체' },
    { name: '한서준', email: 'seojun.han@example.com', role: '2차검토', admin: false, scope: '본인' },
  ];
  let me = 'dohyun.kim@example.com';
  let menuAccess = {};
  ROLES.forEach(r => { menuAccess[r] = {}; MENUS.forEach(m => { menuAccess[r][m.key] = true; }); });
  menuAccess['응대담당'].files = true;
  const settings = { team: '연말정산 검토', year: '2026', folderId: '1AbCdEfGhIjKlMnOp' };
  const reads = {}; // email -> [id]

  const COLS = [
    ['no', 'No.', '', '기본 정보', 'info'], ['dept', '부서', '', '기본 정보', 'info'], ['empNo', '사원번호', '', '기본 정보', 'info'], ['name', '성명', '', '기본 정보', 'info'],
    ['subgroup', '사원하위그룹명', '', '기본 정보', 'info'], ['payArea', '급여영역', '', '기본 정보', 'info'], ['rank', '직급', '', '기본 정보', 'info'],
    ['phone', '전화번호', '', '기본 정보', 'info'], ['email', '이메일주소', '', '기본 정보', 'info'], ['leave', '휴직여부', 'bool', '기본 정보', 'info'],
    ['owner', '담당자', '', '담당', 'assign'], ['owner2', '2차검토 담당자', '', '담당', 'assign'],
    ['review2', '2차 서류검토 여부', 'bool', '2차 검토', 'second'], ['ehr', '시스템 등록', 'bool', '진행', 'first', '^(시스템등록|e-?hr등록)$'],
    ['arrived', '서류 도착여부', 'bool', '진행', 'first'], ['verified', '서류확인 및 검증', 'bool', '진행', 'first'],
    ['prevWork', '종전근무지 여부', 'bool', '진행', 'first', '^종전근무지(여부)?$', true],
    ['rentLoanApply', '주택임차차입금(신청여부)', 'bool', '주택자금 · 월세', 'first'], ['mortApply', '장기주택저당차입금(신청여부)', 'bool', '주택자금 · 월세', 'first'],
    ['mortNts', '장기주택 국세청자료 여부', 'bool', '주택자금 · 월세', 'first'], ['savingApply', '주택마련저축(신청여부)', 'bool', '주택자금 · 월세', 'first'],
    ['savingNts', '주택마련저축 국세청자료 여부', 'bool', '주택자금 · 월세', 'first'], ['rentApply', '월세액(신청여부)', 'bool', '주택자금 · 월세', 'first'],
    ['note1', '특이사항', 'long', '1차 특이·수정사항', 'first'], ['fix1', '수정사항', 'long', '1차 특이·수정사항', 'first'], ['missing', '미비서류', 'long', '1차 특이·수정사항', 'first'],
    ['note2', '특이사항(2차)', 'long', '2차 검토', 'second'], ['fix2', '수정사항(2차)', 'long', '2차 검토', 'second'],
].map(([k, l, t, g, who, re, loose]) => ({ k, l, t, g, who, re: re || '', loose: !!loose }));

  const blank = () => { const p = {}; COLS.forEach(c => { p[c.k] = c.t === 'bool' ? false : ''; }); return p; };
  // 예전 열(수기서류 제출, 종전근무지 회사명·개수, 금액)로 적힌 예시 데이터를 지금 열로 맞춘다
  const person = (o) => {
    if (o.manual) o.arrived = true;
    o.prevWork = !!(o.prevWork || o.prevCount);
    ['manual', 'prevCount', 'rentLoan', 'mort', 'mortPrev', 'saving', 'rent'].forEach(k => delete o[k]);
    const p = Object.assign(blank(), o); p.id = p.empNo; p.updated = p.updated || '2027-02-03 17:20'; p.editor = p.editor || p.owner || ''; return p; };
  const people = [
    person({ no: '1', dept: '인사팀', empNo: '20110321', name: '오민재', subgroup: '정규직', payArea: '본사', rank: '차장', phone: '010-2201-1101', email: 'minjae.oh@example.com', owner: '김도현', owner2: '최유나', ehr: true, arrived: true, verified: true, review2: true, rentApply: true, rent: '7,200,000', note1: '월세 계약서 주소와 주민등록 주소 일치 확인함', editor: '최유나', updated: '2027-02-03 15:02' }),
    person({ no: '2', dept: '재무팀', empNo: '20150702', name: '서지안', subgroup: '정규직', payArea: '본사', rank: '과장', phone: '010-2201-1102', email: 'jian.seo@example.com', owner: '김도현', owner2: '최유나', ehr: true, arrived: true, verified: true, mortApply: true, mort: '4,350,000', mortNts: true, mortPrev: 'O (2019-03-15)', note1: '부친 인적공제 — 동생(타사 재직)과 중복 여부 본인 확인 완료, 동생은 공제 안 함' }),
    person({ no: '3', dept: '영업1팀', empNo: '20180115', name: '윤하람', subgroup: '정규직', payArea: '본사', rank: '대리', phone: '010-2201-1103', email: 'haram.yoon@example.com', owner: '김도현', owner2: '최유나', ehr: true, arrived: true, missing: '안경 구입 영수증(간소화 자료 없음), 월세 이체 내역 12월분', rentApply: true, rent: '6,000,000', note1: '의료비 중 실손보험 수령액 차감 필요 — 보험사 확인서 요청' }),
    person({ no: '4', dept: '영업1팀', empNo: '20190408', name: '강태윤', subgroup: '정규직', payArea: '본사', rank: '대리', phone: '010-2201-1104', email: 'taeyun.kang@example.com', owner: '김도현', owner2: '한서준', ehr: true }),
    person({ no: '5', dept: '영업2팀', empNo: '20200901', name: '임채원', subgroup: '정규직', payArea: '지사', rank: '사원', phone: '010-2201-1105', email: 'chaewon.lim@example.com', owner: '김도현', owner2: '한서준', prevWork: '㈜가나상사', prevCount: '1', ehr: true, arrived: true, note1: '중도입사(2026.04) — 종전근무지 원천징수영수증 받음' }),
    person({ no: '6', dept: '개발팀', empNo: '20160222', name: '조은우', subgroup: '정규직', payArea: '본사', rank: '과장', phone: '010-2201-1106', email: 'eunwoo.cho@example.com', owner: '이수민', owner2: '최유나', ehr: true, arrived: true, verified: true, review2: true, rentLoanApply: true, rentLoan: '3,000,000', note2: '주택임차차입금 대출기관 확인', fix2: '원리금 상환액 3,200,000 → 3,000,000 (중도상환 수수료 제외)' }),
    person({ no: '7', dept: '개발팀', empNo: '20170510', name: '신예준', subgroup: '정규직', payArea: '본사', rank: '대리', phone: '010-2201-1107', email: 'yejun.shin@example.com', owner: '이수민', owner2: '최유나', manual: true, arrived: true, verified: true, savingApply: true, saving: '2,400,000', savingNts: true }),
    person({ no: '8', dept: '개발팀', empNo: '20210303', name: '문서아', subgroup: '계약직', payArea: '본사', rank: '사원', phone: '010-2201-1108', email: 'seoa.moon@example.com', owner: '이수민', owner2: '최유나', ehr: true, arrived: true, missing: '장애인증명서(모친)', note1: '모친 장애인 추가공제 신청 — 증명서 미제출' }),
    person({ no: '9', dept: '총무팀', empNo: '20090817', name: '배준혁', subgroup: '정규직', payArea: '본사', rank: '부장', phone: '010-2201-1109', email: 'junhyuk.bae@example.com', owner: '이수민', owner2: '한서준', leave: true, note1: '육아휴직 중 (2026.09~) — 메일로 안내' }),
    person({ no: '10', dept: '총무팀', empNo: '20140130', name: '황지우', subgroup: '정규직', payArea: '본사', rank: '과장', phone: '010-2201-1110', email: 'jiwoo.hwang@example.com', owner: '이수민', owner2: '한서준', ehr: true, arrived: true }),
    person({ no: '11', dept: '마케팅팀', empNo: '20190923', name: '송민서', subgroup: '정규직', payArea: '본사', rank: '대리', phone: '010-2201-1111', email: 'minseo.song@example.com', owner: '박지호', owner2: '한서준', ehr: true, arrived: true, verified: true, review2: true }),
    person({ no: '12', dept: '마케팅팀', empNo: '20220704', name: '유다인', subgroup: '정규직', payArea: '본사', rank: '사원', phone: '010-2201-1112', email: 'dain.yoo@example.com', owner: '박지호', owner2: '한서준', ehr: true, arrived: true, verified: true, rentApply: true, rent: '8,400,000', note1: '월세 — 총급여 요건 확인함' }),
    person({ no: '13', dept: '생산팀', empNo: '20130612', name: '권도윤', subgroup: '정규직', payArea: '공장', rank: '차장', phone: '010-2201-1113', email: 'doyun.kwon@example.com', owner: '박지호', owner2: '최유나', manual: true, missing: '기부금 영수증 원본(교회)', note1: '수기 제출 — 기부금 영수증 사본만 옴' }),
    person({ no: '14', dept: '생산팀', empNo: '20230102', name: '남궁별', subgroup: '정규직', payArea: '공장', rank: '사원', phone: '010-2201-1114', email: 'byul.namgung@example.com', owner: '박지호', owner2: '최유나', prevWork: '다라물산, 마바테크', prevCount: '2' }),
    person({ no: '15', dept: '품질팀', empNo: '20240205', name: '전시우', subgroup: '정규직', payArea: '공장', rank: '사원', phone: '010-2201-1115', email: 'siwoo.jeon@example.com', owner: '', owner2: '' }),
    person({ no: '16', dept: '품질팀', empNo: '20240819', name: '홍나래', subgroup: '계약직', payArea: '공장', rank: '사원', phone: '010-2201-1116', email: 'narae.hong@example.com', owner: '', owner2: '', ehr: true }),
    // 17~30: 담당자들이 자주 만나는 연말정산 응대 사례
    person({ no: '17', dept: '인사팀', empNo: '20120905', name: '노승현', subgroup: '정규직', payArea: '본사', rank: '차장', phone: '010-2201-1117', email: 'seunghyun.noh@example.com', owner: '김도현', owner2: '최유나', ehr: true, arrived: true, note1: '모친 기본공제 신청 — 형(타사 재직)도 공제 신청했다고 함. 형제 중 한 명만 가능, 본인과 형 협의 결과 회신 대기' }),
    person({ no: '18', dept: '재무팀', empNo: '20170814', name: '구하늘', subgroup: '정규직', payArea: '본사', rank: '과장', phone: '010-2201-1118', email: 'haneul.koo@example.com', owner: '김도현', owner2: '최유나', ehr: true, arrived: true, verified: true, note1: '맞벌이 — 자녀 2명 중 첫째만 본인 공제, 둘째는 배우자 회사에서 공제. 자녀 교육비도 공제받는 쪽으로 맞춤 확인' }),
    person({ no: '19', dept: '해외영업팀', empNo: '20230310', name: '리밍', subgroup: '정규직', payArea: '본사', rank: '대리', phone: '010-2201-1119', email: 'liming@example.com', owner: '김도현', owner2: '한서준', ehr: true, arrived: true, note1: '외국인 근로자 — 단일세율 적용 신청 여부 문의. 일반 정산과 비교해 안내 후 본인 선택 받기로 함', fix1: '외국인등록증 사본 받음' }),
    person({ no: '20', dept: '영업2팀', empNo: '20261201', name: '차서윤', subgroup: '정규직', payArea: '지사', rank: '사원', phone: '010-2201-1120', email: 'seoyun.cha@example.com', owner: '김도현', owner2: '한서준', note1: '2026.12 입사 — 종전근무지 없음(첫 직장) 확인. 간소화 자료 제출 방법 안내 필요' }),
    person({ no: '21', dept: '개발팀', empNo: '20150603', name: '석준기', subgroup: '정규직', payArea: '본사', rank: '과장', phone: '010-2201-1121', email: 'jungi.seok@example.com', owner: '이수민', owner2: '최유나', ehr: true, arrived: true, mortApply: true, mortNts: false, mortPrev: 'X (2026-08-20 신규 차입)', missing: '장기주택저당차입금 이자상환증명서(간소화 자료에 없음), 주택 취득 시 등기사항증명서', note1: '2026년 8월 신규 차입 — 기준시가·차입 시기 요건 확인 필요' }),
    person({ no: '22', dept: '개발팀', empNo: '20181107', name: '어지민', subgroup: '정규직', payArea: '본사', rank: '대리', phone: '010-2201-1122', email: 'jimin.eo@example.com', owner: '이수민', owner2: '최유나', ehr: true, arrived: true, verified: true, review2: true, note1: '배우자 난임시술비 — 간소화 자료에 난임시술로 구분돼 있는지 확인함', note2: '의료비 공제율 구분 확인 완료' }),
    person({ no: '23', dept: '총무팀', empNo: '20080225', name: '하정우', subgroup: '정규직', payArea: '본사', rank: '부장', phone: '010-2201-1123', email: 'jungwoo.ha@example.com', owner: '이수민', owner2: '한서준', ehr: true, arrived: true, verified: true, note1: '대학생 자녀 교육비 — 장학금 받은 금액 빼고 공제. 학자금 대출 상환분은 자녀 본인 공제라 제외 안내' }),
    person({ no: '24', dept: '총무팀', empNo: '20160418', name: '편은지', subgroup: '정규직', payArea: '본사', rank: '과장', phone: '010-2201-1124', email: 'eunji.pyeon@example.com', owner: '이수민', owner2: '한서준', ehr: true, arrived: true, verified: true, review2: true, note1: '정치자금 기부금 10만원 + 고향사랑기부금 — 간소화 자료로 확인' }),
    person({ no: '25', dept: '마케팅팀', empNo: '20210927', name: '탁수빈', subgroup: '정규직', payArea: '본사', rank: '사원', phone: '010-2201-1125', email: 'subin.tak@example.com', owner: '박지호', owner2: '한서준', ehr: true, arrived: true, rentApply: true, rent: '7,800,000', savingApply: true, saving: '3,000,000', savingNts: true, note1: '월세 + 주택청약 — 무주택 세대주 여부 주민등록등본으로 확인 중' }),
    person({ no: '26', dept: '생산팀', empNo: '20070516', name: '마동철', subgroup: '정규직', payArea: '공장', rank: '차장', phone: '010-2201-1126', email: 'dongcheol.ma@example.com', owner: '박지호', owner2: '최유나', manual: true, arrived: true, missing: '부친 장애인증명서, 부친 의료비 영수증(요양병원)', note1: '수기 제출 — 부친(82세) 경로우대·장애인 추가공제 신청' }),
    person({ no: '27', dept: '생산팀', empNo: '20190722', name: '선우진', subgroup: '정규직', payArea: '공장', rank: '대리', phone: '010-2201-1127', email: 'woojin.sunwoo@example.com', owner: '박지호', owner2: '최유나', leave: true, note1: '육아휴직 중 (2026.06~) — 우편·메일로 서류 받기로 함' }),
    person({ no: '28', dept: '품질팀', empNo: '20260302', name: '봉예린', subgroup: '계약직', payArea: '공장', rank: '사원', phone: '010-2201-1128', email: 'yerin.bong@example.com', owner: '박지호', owner2: '한서준', prevWork: '사아물류, 자차유통', prevCount: '2', ehr: true, arrived: true, missing: '두 번째 종전근무지(자차유통) 원천징수영수증', note1: '2026.03 입사 — 종전근무지 2곳, 1곳 서류만 제출' }),
    person({ no: '29', dept: '연구소', empNo: '20140811', name: '피정민', subgroup: '정규직', payArea: '연구소', rank: '책임', phone: '010-2201-1129', email: 'jungmin.pi@example.com', owner: '', owner2: '', ehr: true, arrived: true, note1: '연금저축·IRP 납입 — 연간 한도 문의' }),
    person({ no: '30', dept: '연구소', empNo: '20220103', name: '도아름', subgroup: '정규직', payArea: '연구소', rank: '선임', phone: '010-2201-1130', email: 'areum.do@example.com', owner: '', owner2: '', ehr: true, note1: '배우자 카드 사용분 공제 문의 — 배우자 소득 확인 필요' }),
  ];

  const logs = [
    { date: '2027-02-04 09:40', empNo: '20120905', name: '노승현', author: '김도현', kind: '전화', body: '모친 중복공제 — 형과 상의 후 2/5까지 누가 공제할지 알려주기로 함' },
    { date: '2027-02-04 09:05', empNo: '20150603', name: '석준기', author: '이수민', kind: '보완요청', body: '이자상환증명서·등기사항증명서 요청 메일 발송' },
    { date: '2027-02-03 17:50', empNo: '20230310', name: '리밍', author: '김도현', kind: '방문', body: '단일세율과 일반 정산 비교표 보여주고 설명함. 본인이 일반 정산 선택' },
    { date: '2027-02-03 15:30', empNo: '20260302', name: '봉예린', author: '박지호', kind: '메신저', body: '자차유통 원천징수영수증은 이전 회사 인사팀에 요청 중이라고 함' },
    { date: '2027-02-03 13:20', empNo: '20070516', name: '마동철', author: '박지호', kind: '전화', body: '부친 장애인증명서는 주민센터 아닌 병원 발급 서류도 된다고 안내' },
    { date: '2027-02-02 16:10', empNo: '20080225', name: '하정우', author: '이수민', kind: '메일', body: '학자금 대출 상환분은 자녀 본인 공제라 부모 교육비에서 뺀다고 안내' },
    { date: '2027-02-02 11:45', empNo: '20190722', name: '선우진', author: '박지호', kind: '메일', body: '휴직자 서류 제출 안내 메일 (회신용 주소 포함)' },
    { date: '2027-02-01 10:20', empNo: '20220103', name: '도아름', author: '박지호', kind: '전화', body: '배우자 연 소득 100만원 넘으면 배우자 카드는 공제 대상 아님 안내 (담당 배정 전 응대)' },
    { date: '2027-02-04 09:12', empNo: '20180115', name: '윤하람', author: '김도현', kind: '전화', body: '안경 영수증·12월 월세 이체 내역 2/5(금)까지 제출하기로 함' },
    { date: '2027-02-03 16:40', empNo: '20210303', name: '문서아', author: '이수민', kind: '보완요청', body: '모친 장애인증명서 요청 메일 발송' },
    { date: '2027-02-03 14:05', empNo: '20150702', name: '서지안', author: '김도현', kind: '메신저', body: '부친 중복공제 여부 문의 → 동생은 공제 안 하기로 확인' },
    { date: '2027-02-03 11:30', empNo: '20130612', name: '권도윤', author: '박지호', kind: '방문', body: '기부금 영수증 사본 가져옴. 원본 또는 기부처 발급 영수증 필요 안내' },
    { date: '2027-02-02 17:15', empNo: '20090817', name: '배준혁', author: '이수민', kind: '메일', body: '휴직자 연말정산 안내 메일 발송 (간소화 자료 제출 방법)' },
    { date: '2027-02-02 10:02', empNo: '20200901', name: '임채원', author: '김도현', kind: '전화', body: '종전근무지 원천징수영수증 받는 방법 안내' },
  ];

  const notices = [
    { id: '5', date: '2027-02-03 18:00', author: '정해린', category: '공지', level: '긴급', title: '미비서류 보완 마감 2/10(수) 18시', body: '보완 요청한 직원은 2/10(수) 18시까지 받아주세요.\n마감 후 들어온 서류는 3월 급여 재정산으로 처리합니다.\n미비서류 칸이 비어 있어야 검토완료로 넘어갑니다.', start: '2027-02-01', end: '2027-02-10', pinned: true, files: [], updated: '2027-02-03 18:00', comments: [{ date: '2027-02-03 18:20', author: '박지호', body: '휴직자도 같은 마감인가요?' }, { date: '2027-02-03 18:31', author: '정해린', body: '휴직자는 2/12(금)까지 받아주세요.' }] },
    { id: '4', date: '2027-02-02 09:00', author: '정해린', category: '응대지침', level: '중요', title: '부양가족 중복공제 확인 방법', body: '맞벌이·형제가 같은 부모님을 공제하는 사례가 많습니다.\n1. 본인에게 다른 가족의 공제 여부를 확인\n2. 확인한 내용은 특이사항 칸에 적기\n3. 중복이 의심되면 2차 검토 담당자에게 메신저로 알려주세요', start: '', end: '', pinned: true, files: [F('부양가족_중복확인_체크리스트.pdf', 214000)], updated: '2027-02-02 09:00', comments: [] },
    { id: '3', date: '2027-01-27 13:40', author: '이수민', category: '자료', title: '종전근무지 원천징수영수증 요청 안내문', level: '일반', body: '중도입사자에게 보내는 안내문입니다. 필요하면 복사해서 쓰세요.', start: '', end: '', pinned: false, files: [F('종전근무지_서류요청_안내문.hwp', 31744)], updated: '2027-01-27 13:40', comments: [{ date: '2027-01-27 14:02', author: '김도현', body: '감사합니다!' }] },
    { id: '2', date: '2027-01-20 10:10', author: '박지호', category: '질문', title: '안경 구입비는 1인당 얼마까지인가요?', level: '일반', body: '직원 문의가 와서요. 기준 메뉴에 없어서 여쭤봅니다.', start: '', end: '', pinned: false, files: [], updated: '2027-01-20 10:10', comments: [{ date: '2027-01-20 11:00', author: '정해린', body: '1인당 연 50만원까지입니다. 안경점 영수증(사용자 성명·시력교정용 표시) 받아주세요. 기준 메뉴에도 추가할게요.' }] },
    { id: '1', date: '2027-01-12 09:00', author: '정해린', category: '공지', level: '일반', title: '2026 귀속 연말정산 일정', body: '· 1/15(금) 간소화 자료 조회 시작\n· 1/29(금) 직원 서류 제출 마감\n· 2/1(월)~2/10(수) 1차 검토·보완 요청\n· 2/11(목)~2/12(금) 2차 검토\n· 2월 급여에 정산 결과 반영', start: '', end: '', pinned: false, files: [], updated: '2027-01-12 09:00', comments: [] },
  ];

  const rules = [
    { id: '1', category: '인적공제', title: '기본공제 대상 부양가족 요건', year: '2026', owner: '정해린', link: '', files: [], updated: '2027-01-10', editor: '정해린', body: '(예시 · 국세청 안내로 최신 기준 확인 후 사용)\n· 1인당 150만원 소득공제\n· 연간 소득금액 100만원 이하 (근로소득만 있으면 총급여 500만원 이하)\n· 나이: 직계존속 만 60세 이상, 직계비속 만 20세 이하, 형제자매 만 20세 이하 또는 60세 이상\n· 장애인은 나이 제한 없음 (소득 요건은 적용)\n· 같은 사람을 두 명이 공제할 수 없음 → 맞벌이 부부·형제 중복 여부 꼭 확인', history: [{ date: '2027-01-10', editor: '정해린', note: '최초 등록' }] },
    { id: '2', category: '주택자금·월세', title: '월세액 세액공제', year: '2026', owner: '정해린', link: '', files: [], updated: '2027-01-10', editor: '정해린', body: '(예시 · 국세청 안내로 최신 기준 확인 후 사용)\n· 무주택 세대의 세대주, 총급여 8천만원 이하\n· 국민주택규모(85㎡) 이하 또는 기준시가 4억원 이하 주택\n· 연 1,000만원 한도, 총급여 5,500만원 이하 17% / 그 외 15%\n· 제출: 임대차계약서 사본, 월세 이체 내역 (주민등록 주소지와 계약서 주소 일치 확인)', history: [{ date: '2027-01-10', editor: '정해린', note: '최초 등록' }] },
    { id: '3', category: '주택자금·월세', title: '주택임차차입금 원리금 상환액', year: '2026', owner: '정해린', link: '', files: [], updated: '2027-01-10', editor: '정해린', body: '(예시 · 국세청 안내로 최신 기준 확인 후 사용)\n· 무주택 세대의 세대주, 국민주택규모 주택 임차\n· 상환액의 40% 소득공제, 주택마련저축 납입액과 합산 연 400만원 한도\n· 제출: 주택자금상환증명서(금융기관), 임대차계약서, 주민등록등본', history: [{ date: '2027-01-10', editor: '정해린', note: '최초 등록' }] },
    { id: '4', category: '의료비', title: '의료비 세액공제', year: '2026', owner: '정해린', link: '', files: [F('의료비_공제_안내.pdf', 388000)], updated: '2027-01-20', editor: '정해린', body: '(예시 · 국세청 안내로 최신 기준 확인 후 사용)\n· 총급여 3% 초과분의 15% (난임시술비 30%, 미숙아·선천성이상아 20%)\n· 본인·65세 이상·장애인·6세 이하: 한도 없음 / 그 외 부양가족: 연 700만원\n· 실손보험금 받은 금액은 빼야 함\n· 안경·콘택트렌즈: 1인당 연 50만원 (시력교정용, 사용자 성명 표시 영수증)', history: [{ date: '2027-01-20', editor: '정해린', note: '안경 구입비 한도 추가 (공지 질문 반영)' }, { date: '2027-01-10', editor: '정해린', note: '최초 등록' }] },
    { id: '5', category: '제출서류', title: '공통 제출 서류', year: '2026', owner: '이수민', link: '', files: [F('소득세액공제신고서_양식.xlsx', 58000)], updated: '2027-01-10', editor: '정해린', body: '(예시 · 회사 기준에 맞게 고쳐 쓰세요)\n· 소득·세액공제신고서\n· 국세청 간소화 자료 PDF (부양가족 자료 제공 동의 포함)\n· 부양가족 변동 시 가족관계증명서\n· 중도입사자: 종전근무지 원천징수영수증·근로소득지급명세서', history: [{ date: '2027-01-10', editor: '정해린', note: '최초 등록' }] },
    { id: '6', category: '응대 기준', title: '증빙 서류 받는 방법', year: '2026', owner: '정해린', link: '', files: [], updated: '2027-01-10', editor: '정해린', body: '(예시 · 회사 기준에 맞게 고쳐 쓰세요)\n· 증빙은 지정된 경로(웹앱 첨부 또는 회사 메일)로만 받습니다. 개인 메신저 금지\n· 전화 문의는 본인 확인 후 안내\n· 안내한 내용은 응대기록에 남겨 다른 담당자도 볼 수 있게 합니다', history: [{ date: '2027-01-10', editor: '정해린', note: '최초 등록' }] },
  ];

  const files = [
    { id: 'demo-f1', date: '2027-02-04 09:20', uploader: '김도현', area: 'person', ref: '20180115', refLabel: '윤하람(20180115)', name: '윤하람_간소화자료.pdf', url: '#demo-file', size: 512000, mime: '', memo: '' },
    { id: 'demo-f2', date: '2027-02-03 15:10', uploader: '김도현', area: 'person', ref: '20150702', refLabel: '서지안(20150702)', name: '서지안_주택자금상환증명서.pdf', url: '#demo-file', size: 230000, mime: '', memo: '' },
    { id: 'demo-f3', date: '2027-02-03 11:40', uploader: '박지호', area: 'person', ref: '20130612', refLabel: '권도윤(20130612)', name: '권도윤_기부금영수증_사본.jpg', url: '#demo-file', size: 1830000, mime: '', memo: '' },
    { id: 'demo-f4', date: '2027-02-02 10:30', uploader: '김도현', area: 'person', ref: '20200901', refLabel: '임채원(20200901)', name: '임채원_종전근무지_원천징수영수증.pdf', url: '#demo-file', size: 160000, mime: '', memo: '' },
    { id: 'demo-부양가족_중복확인_체크리스트.pdf', date: '2027-02-02 09:00', uploader: '정해린', area: 'notice', ref: '4', refLabel: '부양가족 중복공제 확인 방법', name: '부양가족_중복확인_체크리스트.pdf', url: '#demo-file', size: 214000, mime: '', memo: '' },
    { id: 'demo-종전근무지_서류요청_안내문.hwp', date: '2027-01-27 13:40', uploader: '이수민', area: 'notice', ref: '3', refLabel: '종전근무지 원천징수영수증 요청 안내문', name: '종전근무지_서류요청_안내문.hwp', url: '#demo-file', size: 31744, mime: '', memo: '' },
    { id: 'demo-의료비_공제_안내.pdf', date: '2027-01-20 11:10', uploader: '정해린', area: 'rules', ref: '4', refLabel: '의료비 세액공제', name: '의료비_공제_안내.pdf', url: '#demo-file', size: 388000, mime: '', memo: '' },
    { id: 'demo-소득세액공제신고서_양식.xlsx', date: '2027-01-10 09:00', uploader: '정해린', area: 'rules', ref: '5', refLabel: '공통 제출 서류', name: '소득세액공제신고서_양식.xlsx', url: '#demo-file', size: 58000, mime: '', memo: '' },
    { id: 'demo-f10', date: '2027-02-03 15:40', uploader: '박지호', area: 'person', ref: '20260302', refLabel: '봉예린(20260302)', name: '봉예린_종전근무지1_사아물류.pdf', url: '#demo-file', size: 148000, mime: '', memo: '' },
    { id: 'demo-f11', date: '2027-02-03 13:30', uploader: '박지호', area: 'person', ref: '20070516', refLabel: '마동철(20070516)', name: '마동철_수기제출_공제신고서.jpg', url: '#demo-file', size: 2240000, mime: '', memo: '' },
    { id: 'demo-f12', date: '2027-02-02 14:00', uploader: '이수민', area: 'person', ref: '20181107', refLabel: '어지민(20181107)', name: '어지민_간소화자료.pdf', url: '#demo-file', size: 486000, mime: '', memo: '' },
    { id: 'demo-f13', date: '2027-02-02 10:15', uploader: '김도현', area: 'person', ref: '20230310', refLabel: '리밍(20230310)', name: '리밍_외국인등록증_사본.pdf', url: '#demo-file', size: 205000, mime: '', memo: '' },
    { id: 'demo-f9', date: '2027-01-09 16:00', uploader: '정해린', area: 'etc', ref: '', refLabel: '', name: '2026귀속_연말정산_직원안내문.pdf', url: '#demo-file', size: 742000, mime: '', memo: '직원 배포용 안내문' },
  ];
  const AREAS = { person: '대상자 증빙', notice: '공지사항', rules: '연말정산 기준', etc: '일반 자료' };
  const AREA_MENU = { person: 'review', notice: 'notice', rules: 'rules', etc: 'files' };

  function meInfo() {
    const p = staff.find(x => x.email === me);
    const info = { email: p.email, name: p.name, role: p.role, isLeader: p.role === '총괄', isAdmin: !!p.admin };
    info.scope = info.isLeader || info.isAdmin ? '전체' : (p.scope || '전체');
    info.menus = info.isAdmin ? MENUS.map(x => x.key).concat('admin') : MENUS.map(x => x.key).filter(k => menuAccess[p.role][k] !== false);
    return info;
  }
  const lead = m => m.isLeader || m.isAdmin;
  // 서버의 canSee_와 같은 규칙
  const canSee = (m, p) => {
    if (lead(m) || m.scope === '전체') return true;
    if (p.owner === m.name || p.owner2 === m.name) return true;
    const list = m.scope.split(',');
    return !p.owner || list.indexOf(p.owner) >= 0;
  };
  const seen = m => new Set(people.filter(p => canSee(m, p)).map(p => p.empNo || p.id));
  const myPeople = m => people.filter(p => canSee(m, p));
  const myLogs = m => { const s = seen(m); return logs.filter(l => s.has(l.empNo)); };
  const need = (m, k) => { if (m.menus.indexOf(k) < 0) throw new Error(`[${(MENUS.find(x => x.key === k) || { label: k }).label}] 메뉴를 사용할 권한이 없습니다.`); };
  const myReads = () => (reads[me] = reads[me] || ['1', '2', '3']);
  const visFiles = m => { const s = seen(m); return files.filter(f => m.menus.indexOf(AREA_MENU[f.area]) >= 0 && (f.area !== 'person' || s.has(f.ref))); };
  const sortNotices = () => notices.sort((a, b) => (b.pinned - a.pinned) || (['긴급', '중요', '일반'].indexOf(a.level) - ['긴급', '중요', '일반'].indexOf(b.level)) || b.date.localeCompare(a.date));
  function canEdit(m, cur, c, patch) {
    if (lead(m)) return true;
    const takes = !cur.owner && patch.owner === m.name;
    if (c.who === 'assign') return c.k === 'owner' && takes;
    if (c.who === 'second') return cur.owner2 === m.name;
    return cur.owner === m.name || takes;
  }
  function register(area, ref, refLabel, list, m, memo) {
    (list || []).forEach(f => { if (!files.some(x => x.id === f.id)) files.unshift({ id: f.id, date: stamp(), uploader: m.name, area, ref, refLabel, name: f.name, url: f.url, size: f.size, mime: f.mime || '', memo: memo || '' }); });
  }
  function adminData() {
    const st = {};
    staff.forEach(p => { st[p.name] = { sheet: 'editor', folder: 'editor' }; });
    st['정해린'] = { sheet: 'owner', folder: 'owner' };
    st['한서준'] = { sheet: 'editor', folder: 'none' };
    return {
      members: staff.map((p, i) => ({ order: i + 1, name: p.name, email: p.email, role: p.role, admin: !!p.admin, scope: p.scope || '전체' })),
      owners: people.reduce((o, p) => { const k = p.owner || '(미배정)'; o[k] = (o[k] || 0) + 1; return o; }, {}),
      assign: people.map(p => [p.owner, p.owner2]),
      menus: MENUS, roles: ROLES, menuAccess: JSON.parse(JSON.stringify(menuAccess)), settings: Object.assign({}, settings),
      webAppUrl: 'https://script.google.com/macros/s/AKfy...예시.../exec',
      status: { resources: { sheet: { ok: true, name: '2026 귀속 연말정산 검토 (스프레드시트)' }, folder: { ok: true, name: '연말정산 증빙 (담당자 전용)' } }, people: st },
    };
  }

  return {
    delay: { apiUploadFile: 900, apiApplyPermissions: 1200, apiSavePerson: 700, apiAddLog: 600, apiSaveNotice: 700, apiSaveRule: 700, apiAddNoticeComment: 500, apiDeleteFile: 600 },
    users: () => staff.map(p => ({ email: p.email, name: p.name, role: p.role + (p.admin ? '·관리자' : '') })),
    setUser: email => { me = email; },
    now: () => NOW,

    apiBootstrap() {
      const m = meInfo();
      const can = k => m.menus.indexOf(k) >= 0;
      return {
        team: settings.team, year: settings.year, sheetUrl: 'https://docs.google.com/spreadsheets/', me: m,
        staff: staff.map(s => ({ name: s.name, role: s.role })), cols: COLS,
        logKinds: ['전화', '메일', '메신저', '방문', '보완요청', '기타'],
        noticeCategories: ['공지', '응대지침', '자료', '질문'], noticeLeaderOnly: ['공지', '응대지침'],
        ruleCategories: ['인적공제', '주택자금·월세', '의료비', '교육비', '기부금', '신용카드 등', '연금·보험', '제출서류', '응대 기준', '기타'],
        fileAreas: AREAS,
        people: can('review') || can('dash') ? myPeople(m) : [], logs: can('review') ? myLogs(m) : [],
        notices: can('notice') || can('dash') ? sortNotices() : [], reads: myReads().slice(),
        rules: can('rules') ? rules : [], files: can('files') || can('review') ? visFiles(m) : [],
      };
    },

    apiSavePerson(id, patch) {
      const m = meInfo(); need(m, 'review');
      const cur = people.find(p => p.id === id);
      if (!cur) throw new Error('대상자를 찾을 수 없습니다.');
      if (!canSee(m, cur)) throw new Error('열람 범위 밖의 대상자입니다.');
      Object.keys(patch).forEach(k => {
        const c = COLS.find(x => x.k === k);
        if (!canEdit(m, cur, c, patch)) throw new Error(`[${c.l}] 칸은 ${c.who === 'second' ? '2차검토 담당자' : c.who === 'assign' ? '총괄' : '담당자'}만 고칠 수 있습니다.`);
      });
      Object.assign(cur, patch, { updated: stamp(), editor: m.name });
      cur.id = cur.empNo || id;
      return { personPatch: cur, oldId: id };
    },
    apiAddPerson(f) {
      const m = meInfo(); if (!lead(m)) throw new Error('대상자 추가는 총괄 또는 관리자만 할 수 있습니다.');
      people.push(person(Object.assign({ no: String(people.length + 1) }, f, { editor: m.name, updated: stamp() })));
      return { people: myPeople(m) };
    },
    apiDeletePerson(id) {
      const m = meInfo(); if (!lead(m)) throw new Error('대상자 삭제는 총괄 또는 관리자만 할 수 있습니다.');
      people.splice(people.findIndex(p => p.id === id), 1);
      return { people };
    },
    apiImportPeople(rows) {
      const m = meInfo(); if (!lead(m)) throw new Error('붙여넣기는 총괄 또는 관리자만 할 수 있습니다.');
      let added = 0, updated = 0;
      rows.forEach(src => {
        const o = {};
        COLS.forEach(c => { if (!(c.k in src)) return; const v = String(src[c.k]).trim(); o[c.k] = c.t !== 'bool' ? v : c.loose ? !!v && !/^(x|n|no|false|0|없음|무|해당없음|-)$/i.test(v) : /^(o|y|yes|true|1|○|●|◯|v|✓|✔|완료|등록|제출|도착|신청|있음|해당|유|휴직)$/i.test(v); });
        const ex = o.empNo && people.find(p => p.empNo === o.empNo);
        if (ex) { Object.assign(ex, o, { updated: stamp(), editor: m.name }); updated++; }
        else { people.push(person(Object.assign(o, { editor: m.name, updated: stamp() }))); added++; }
      });
      return { people: myPeople(m), added, updated };
    },
    apiTemplateLink() { return { demo: true }; },
    apiAddLog(id, kind, body) {
      const m = meInfo(); need(m, 'review');
      const p = people.find(x => x.id === id);
      if (!canSee(m, p)) throw new Error('열람 범위 밖의 대상자입니다.');
      const entry = { date: stamp(), empNo: p.empNo, name: p.name, author: m.name, kind, body };
      logs.unshift(entry);
      return { logAdded: entry };
    },

    apiSaveNotice(n) {
      const m = meInfo(); need(m, 'notice');
      if (['공지', '응대지침'].indexOf(n.category) >= 0 && !lead(m)) throw new Error(`[${n.category}]는 총괄 또는 관리자만 올릴 수 있습니다.`);
      let id = n.id;
      const ex = id && notices.find(x => x.id === id);
      const vals = { category: n.category, level: lead(m) ? n.level : '일반', title: n.title, body: n.body, start: n.start, end: n.end, pinned: lead(m) && n.pinned, files: n.files, updated: stamp() };
      if (ex) Object.assign(ex, vals);
      else { id = String(++seq); notices.push(Object.assign({ id, date: stamp(), author: m.name, comments: [] }, vals)); }
      register('notice', id, n.title, n.files, m);
      if (myReads().indexOf(id) < 0) myReads().push(id);
      return { notices: sortNotices(), savedId: id, reads: myReads().slice() };
    },
    apiDeleteNotice(id) { notices.splice(notices.findIndex(n => n.id === id), 1); return { notices: sortNotices() }; },
    apiAddNoticeComment(id, body) { const c = { date: stamp(), author: meInfo().name, body }; notices.find(n => n.id === id).comments.push(c); return { commentAdded: { id, comment: c } }; },
    apiMarkRead(id) { if (myReads().indexOf(id) < 0) myReads().push(id); return { reads: myReads().slice() }; },
    apiMarkUnread(id) { reads[me] = myReads().filter(x => x !== id); return { reads: reads[me].slice() }; },

    apiSaveRule(rule) {
      const m = meInfo(); if (!lead(m)) throw new Error('연말정산 기준은 총괄 또는 관리자만 등록·수정할 수 있습니다.');
      let id = rule.id;
      const ex = id && rules.find(r => r.id === id);
      const vals = { category: rule.category, title: rule.title, body: rule.body, year: rule.year, owner: rule.owner, link: rule.link, files: rule.files, updated: stamp().slice(0, 10), editor: m.name };
      if (ex) { Object.assign(ex, vals); ex.history.unshift({ date: vals.updated, editor: m.name, note: rule.changeNote || '내용 수정' }); }
      else { id = String(++seq); rules.push(Object.assign({ id, history: [{ date: vals.updated, editor: m.name, note: '최초 등록' }] }, vals)); }
      register('rules', id, rule.title, rule.files, m);
      return { rules, savedId: id };
    },
    apiDeleteRule(id) { rules.splice(rules.findIndex(r => r.id === id), 1); return { rules }; },

    apiUploadFile(f) {
      const m = meInfo(); need(m, AREA_MENU[f.area] || 'files');
      const info = { id: 'demo-up-' + (++seq), name: f.name, url: '#demo-file', size: Math.round((f.data || '').length * 0.75), mime: f.mimeType || '' };
      if (f.area === 'person' || f.area === 'etc') { register(f.area, f.ref || '', f.refLabel || '', [info], m, f.memo); return { file: info, files: visFiles(m) }; }
      return { file: info };
    },
    apiDeleteFile(id) {
      const m = meInfo(); const f = files.find(x => x.id === id);
      if (f.uploader !== m.name && !lead(m)) throw new Error('본인이 올린 파일만 삭제할 수 있습니다.');
      files.splice(files.indexOf(f), 1);
      return { files: visFiles(m) };
    },

    apiAdminData() { if (!meInfo().isAdmin) throw new Error('관리자만 사용할 수 있습니다.'); return adminData(); },
    apiSaveMembers(list) {
      const keep = staff.slice();
      staff.length = 0;
      list.forEach(x => { const old = keep.find(s => s.email === x.email) || {}; staff.push(Object.assign(old, { name: x.name, email: x.email, role: x.role, admin: !!x.admin, scope: x.scope || '전체' })); });
      if (!staff.some(s => s.email === me)) me = staff[0].email;
      return adminData();
    },
    apiSaveMenuAccess(a) { menuAccess = JSON.parse(JSON.stringify(a)); return { menuAccess }; },
    apiSaveSettings(s) { if (!/^\d{4}$/.test(String(s.year).trim())) throw new Error('귀속연도는 2026처럼 네 자리로 입력해 주세요.'); Object.assign(settings, { team: s.team, year: String(s.year).trim(), folderId: s.folderId.replace(/^.*\/folders\//, '') }); return adminData(); },
    apiApplyPermissions() {
      return { log: [{ ok: true, text: '한서준: 첨부 폴더 편집 권한을 줬습니다.' }, { ok: true, text: '시트 보호를 적용했습니다. (설정·담당자: 관리자만 / 기준: 총괄·관리자만)' }], data: (() => { const d = adminData(); d.status.people['한서준'].folder = 'editor'; return d; })() };
    },
    apiRevokeAccess() { return { log: [{ ok: true, text: '스프레드시트 접근을 해제했습니다.' }, { ok: true, text: '첨부 폴더 접근을 해제했습니다.' }], data: adminData() }; },
  };
})();
