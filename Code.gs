/**
 * 청년부 청구함 — Apps Script 백엔드
 *
 * 이 스크립트는 '청년부 청구함' 구글 시트에 붙어서 동작해요.
 *  - 사이트(GitHub Pages)가 보내는 요청을 받아 시트에 기록하고, 본인 소속 데이터만 돌려줘요.
 *  - 영수증 사진은 드라이브 폴더에, 계좌번호는 회계 전용 스프레드시트에 따로 저장해요.
 *  - 시트 상단 메뉴 '청구함 관리'에서 정산 목록 만들기, 입금 완료 처리, PIN 초기화를 해요.
 *  - 금요일 마감 직후 한 번, 회계에게 이번 주 청구 요약 메일을 보내요. (청구자에게는 메일을 보내지 않아요)
 *
 * 설치 방법은 README.md 를 보세요.
 */

/* ───────── 시트 구조 ───────── */
const HEAD = {
  '설정':      ['키', '값', '설명'],
  '명단':      ['이름', '소속1', '직분1', '소속2', '직분2', '소속3', '직분3', '사용', 'PIN해시', 'PIN실패', '로그인', '마지막접속', '권한'],
  '소속':      ['소속명', '유형', '등록인원', '사용'],
  '예산항목':  ['항목명', '대상유형', '산정방식', '금액', '금액표기', '반기1회', '단체사진', '인원입력', '사전보고', '용도', '사용'],
  '청구내역':  ['청구번호', '제출시각', '이름', '소속', '예산항목', '사용일', '사용처', '내용', '결제금액', '인원', '청구금액', '개인부담', '영수증', '단체사진', '상태', '보류사유', '정산회차', '입금일', '재청구원번호', '반기', '결제내역', '처리자', '처리시각'],
  '정산':      ['회차', '정산일', '이름', '건수', '합계', '청구번호', '입금완료'],
  '규정':      ['구분', '내용', '대상유형', '노출'],
  'FAQ':       ['질문', '답변', '대상유형', '자주헷갈림', '노출'],
  '헷갈리는것': ['제목', '설명', '대상유형', '노출']
};
const ACCT_HEAD = ['청구번호', '제출시각', '이름', '은행', '계좌번호', '예금주'];
const TRANSFER_HEAD = ['회차', '이름', '은행', '계좌번호', '예금주', '금액', '받는분 통장 표시', '청구번호'];
const STATUS = ['접수', '확인', '보류', '입금완료', '취소'];
const TYPES = ['셀', '또래', '팀', '임원단'];
const DOW = ['일', '월', '화', '수', '목', '금', '토'];
const MAX_FAIL = 5;
const TOKEN_DAYS = 90;
const MAX_TOKENS = 5;
const SCHEMA = '3';
const CACHED_TABS = ['설정', '소속', '예산항목', '규정', 'FAQ', '헷갈리는것'];
let MEMO = {};   // 한 번의 요청 안에서 같은 탭을 두 번 읽지 않도록

/* ───────── 공통 도구 ───────── */
const props_ = () => PropertiesService.getScriptProperties();
function book_() {
  const id = props_().getProperty('BOOK_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
const sh_ = name => book_().getSheetByName(name);
function readRows_(name) {
  const s = sh_(name);
  if (!s || s.getLastRow() < 2) return [];
  const v = s.getRange(1, 1, s.getLastRow(), HEAD[name].length).getValues();
  v.shift();
  const h = HEAD[name];
  return v.map((r, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = r[j]); return o; })
          .filter(o => String(o[h[0]]).trim() !== '');
}
/** 탭 읽기: 요청 안에서는 한 번만, 자주 안 바뀌는 탭은 10분 캐시 (시트를 고치면 바로 지워짐) */
function rows_(name) {
  if (MEMO[name]) return MEMO[name];
  let rows = null;
  const cacheable = CACHED_TABS.indexOf(name) >= 0;
  if (cacheable) {
    try { const c = CacheService.getScriptCache().get('rows:' + name); if (c) rows = JSON.parse(c); } catch (e) {}
  }
  if (!rows) {
    rows = readRows_(name);
    if (cacheable) { try { CacheService.getScriptCache().put('rows:' + name, JSON.stringify(rows), 600); } catch (e) {} }
  }
  MEMO[name] = rows;
  return rows;
}
function forget_(name) { delete MEMO[name]; if (CACHED_TABS.indexOf(name) >= 0) { try { CacheService.getScriptCache().remove('rows:' + name); } catch (e) {} } }
/** 시트를 직접 고치면 캐시를 지워서 사이트에 바로 반영 (단순 트리거, 설치 불필요) */
function onEdit(e) { try { forget_(e.range.getSheet().getName()); } catch (x) {} }
const col_ = (name, key) => HEAD[name].indexOf(key) + 1;
const set_ = (name, row, key, val) => { sh_(name).getRange(row, col_(name, key)).setValue(val); forget_(name); };
const append_ = (name, row) => { sh_(name).appendRow(row); forget_(name); };
/** 예전 버전으로 설치한 시트에 새 칸 제목을 자동으로 채움 */
function ensureSchema_() {
  if (props_().getProperty('SCHEMA') === SCHEMA) return;
  Object.keys(HEAD).forEach(name => {
    const s = sh_(name); if (!s) return;
    const cur = s.getRange(1, 1, 1, HEAD[name].length).getValues()[0];
    HEAD[name].forEach((h, i) => { if (String(cur[i] || '').trim() === '') s.getRange(1, i + 1).setValue(h).setFontWeight('bold').setBackground('#FFF2E7'); });
  });
  props_().setProperty('SCHEMA', SCHEMA);
}
const yes_ = v => v === true || String(v).trim().toUpperCase() === 'Y';
const off_ = v => String(v).trim().toUpperCase() === 'N';
const tz_ = () => Session.getScriptTimeZone() || 'Asia/Seoul';
const fmt_ = (d, p) => Utilities.formatDate(d, tz_(), p || 'yyyy-MM-dd');
function cfg_() { const o = {}; rows_('설정').forEach(r => o[String(r['키']).trim()] = r['값']); return o; }
function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}
function secret_() {
  let s = props_().getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props_().setProperty('SECRET', s); }
  return s;
}
const pinHash_ = (name, pin) => hash_(secret_() + '|' + name + '|' + pin);
function alert_(msg) { try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); } }
function err_(msg, code) { const e = new Error(msg); e.userMsg = msg; e.code = code || ''; return e; }
function halfKey_(d) { const m = d.getMonth() + 1, y = d.getFullYear(); if (m === 12) return (y + 1) + 'H1'; if (m <= 5) return y + 'H1'; return y + 'H2'; }
function halfLabel_(k) { return k.slice(0, 4) + (k.endsWith('H1') ? ' 상반기 (12~5월)' : ' 하반기 (6~11월)'); }
function parseYmd_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}
const ymdOf_ = v => v instanceof Date ? fmt_(v) : String(v || '');
function parseLines_(v) { try { const a = JSON.parse(v || '[]'); return Array.isArray(a) ? a : []; } catch (e) { return []; } }
const typesOf_ = v => String(v || '').split(',').map(s => s.trim()).filter(Boolean);

/* ───────── 데이터 읽기 ───────── */
function units_() {
  return rows_('소속').filter(r => !off_(r['사용']))
    .map(r => ({ id: String(r['소속명']).trim(), type: String(r['유형']).trim(), size: Number(r['등록인원']) || 0 }));
}
function members_() {
  const U = units_();
  return rows_('명단').filter(r => !off_(r['사용'])).map(r => {
    const units = [];
    [1, 2, 3].forEach(k => {
      const u = U.find(x => x.id === String(r['소속' + k]).trim());
      if (u) units.push({ id: u.id, type: u.type, size: u.size, role: String(r['직분' + k] || '').trim() });
    });
    const role = String(r['권한'] || '').trim();
    return { name: String(r['이름']).trim(), units, row: r, role: ['회계', '열람'].indexOf(role) >= 0 ? role : '' };
  });
}
const member_ = name => members_().find(m => m.name === String(name || '').trim());
function items_() {
  return rows_('예산항목').filter(r => !off_(r['사용'])).map(r => ({
    id: String(r['항목명']).trim(),
    type: String(r['대상유형']).trim(),
    mode: String(r['산정방식']).trim() === '인당' ? 'perPerson' : 'pool',
    amount: Number(r['금액']) || 0,
    basis: String(r['금액표기'] || '').trim() || (String(r['산정방식']).trim() === '인당' ? '인당' : '반기'),
    once: yes_(r['반기1회']), photo: yes_(r['단체사진']), heads: yes_(r['인원입력']), pre: yes_(r['사전보고']),
    uses: String(r['용도'] || '')
  }));
}

/* ───────── 로그인 ───────── */
function checkPass_(pass) {
  if (String(pass || '') !== String(cfg_()['공동체암호'])) throw err_('암호가 맞지 않아요. 리더 단톡방 공지를 확인하세요.', 'BADPASS');
}
function tokens_(m) { try { return JSON.parse(m.row['로그인'] || '[]'); } catch (e) { return []; } }
function auth_(token) {
  if (!token) throw err_('다시 로그인해주세요.', 'AUTH');
  const h = hash_(String(token)), now = Date.now(), cache = CacheService.getScriptCache();
  const hit = cache.get('tok:' + h);
  const list = members_();
  let m = hit ? list.find(x => x.name === hit) : null;
  if (!m) {
    m = list.find(x => tokens_(x).some(t => t.h === h && t.exp > now));
    if (m) cache.put('tok:' + h, m.name, 21600);
  }
  if (!m) throw err_('다시 로그인해주세요.', 'AUTH');
  return m;
}
function admin_(token, needPay) {
  const m = auth_(token);
  if (!m.role) throw err_('정산 권한이 없어요.', 'ROLE');
  if (needPay && m.role !== '회계') throw err_('보기 전용 권한이에요. 처리는 회계가 해요.', 'ROLE');
  return m;
}

/* ───────── 사이트에 돌려줄 데이터 (본인 소속 것만) ───────── */
function state_(name) {
  const m = member_(name), c = cfg_();
  const types = m.units.map(u => u.type);
  const vis = t => { const ts = typesOf_(t); return !ts.length || ts.some(x => types.indexOf(x) >= 0); };
  const half = halfKey_(new Date());
  const all = rows_('청구내역');
  const myUnitIds = m.units.map(u => u.id);
  const used = {}, count = {};
  all.forEach(r => {
    if (['보류', '취소'].indexOf(r['상태']) >= 0 || r['반기'] !== half || myUnitIds.indexOf(r['소속']) < 0) return;
    const k = r['소속'] + '|' + r['예산항목'];
    used[k] = (used[k] || 0) + (Number(r['청구금액']) || 0);
    count[k] = (count[k] || 0) + 1;
  });
  const resubOf = {};
  all.forEach(r => { if (r['재청구원번호'] && r['상태'] !== '취소') resubOf[r['재청구원번호']] = r['청구번호']; });
  const claims = all.filter(r => r['이름'] === m.name && r['상태'] !== '취소')
    .sort((a, b) => new Date(b['제출시각']) - new Date(a['제출시각']))
    .map(r => ({
      no: r['청구번호'], unit: r['소속'], item: r['예산항목'], date: ymdOf_(r['사용일']), place: r['사용처'], memo: r['내용'],
      paid: Number(r['결제금액']) || 0, amount: Number(r['청구금액']) || 0, heads: Number(r['인원']) || undefined,
      status: r['상태'], reason: r['보류사유'], payDate: r['입금일'] ? ymdOf_(r['입금일']) : '',
      fixOf: r['재청구원번호'] || undefined, resub: resubOf[r['청구번호']],
      lines: parseLines_(r['결제내역'])
    }));
  return {
    me: { name: m.name, units: m.units, role: m.role },
    items: items_().filter(i => types.indexOf(i.type) >= 0),
    used, count, claims,
    halfLabel: halfLabel_(half),
    rules: rows_('규정').filter(r => !off_(r['노출']) && vis(r['대상유형'])).map(r => [r['구분'], r['내용']]),
    faq: rows_('FAQ').filter(r => !off_(r['노출']) && vis(r['대상유형'])).map(r => [r['질문'], r['답변'], yes_(r['자주헷갈림'])]),
    confuse: rows_('헷갈리는것').filter(r => !off_(r['노출']) && vis(r['대상유형'])).map(r => [r['제목'], r['설명']]),
    deadline: { day: Math.max(0, DOW.indexOf(String(c['마감요일'] || '금').trim())), hour: Number(c['마감시각']) || 18 }
  };
}

/* ───────── 사이트 요청 처리 ───────── */
function doGet() { return json_({ ok: true, message: '청년부 청구함 API가 동작 중이에요.' }); }
function doPost(e) {
  let out;
  try {
    MEMO = {};
    ensureSchema_();
    const p = JSON.parse(e.postData.contents);
    const fn = API[p.action];
    if (!fn) throw err_('잘못된 요청이에요.');
    out = Object.assign({ ok: true }, fn(p));
  } catch (x) {
    if (!x.userMsg) console.error(x);
    out = { ok: false, error: x.userMsg || '오류가 났어요. 잠시 후 다시 시도해주세요.', code: x.code || '' };
  }
  return json_(out);
}
const json_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

const API = {
  /** 공동체 암호 확인 후 이름 목록 */
  names(p) {
    checkPass_(p.pass);
    return { names: members_().map(m => ({ name: m.name, roles: m.units.map(u => u.id + ' ' + u.role).join(' · '), hasPin: !!m.row['PIN해시'] })) };
  },

  /** 이름 + PIN 으로 입장. 처음이면 PIN 을 새로 저장 */
  login(p) {
    checkPass_(p.pass);
    const m = member_(p.name);
    if (!m) throw err_('명단에 없는 이름이에요. 회계에게 추가를 요청하세요.');
    const fails = Number(m.row['PIN실패']) || 0;
    if (fails >= MAX_FAIL) throw err_('PIN을 5번 틀려서 잠겼어요. 회계에게 초기화를 요청하세요.', 'LOCKED');
    const pin = String(p.pin || '');
    if (!/^\d{4}$/.test(pin)) throw err_('PIN 4자리를 입력하세요.');
    const row = m.row._row;
    if (m.row['PIN해시']) {
      if (m.row['PIN해시'] !== pinHash_(m.name, pin)) {
        set_('명단', row, 'PIN실패', fails + 1);
        if (fails + 1 >= MAX_FAIL) throw err_('PIN을 5번 틀려서 잠겼어요. 회계에게 초기화를 요청하세요.', 'LOCKED');
        throw err_('PIN이 맞지 않아요. (' + (fails + 1) + '/' + MAX_FAIL + '회, ' + MAX_FAIL + '회 틀리면 잠겨요)', 'BADPIN');
      }
    } else {
      if (/^(\d)\1{3}$/.test(pin) || ['1234', '4321', '0123'].indexOf(pin) >= 0) throw err_('너무 쉬운 PIN이에요. 다른 숫자로 정해주세요.');
      set_('명단', row, 'PIN해시', pinHash_(m.name, pin));
    }
    const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    const now = Date.now();
    const list = tokens_(m).filter(t => t.exp > now).slice(-(MAX_TOKENS - 1));
    list.push({ h: hash_(token), exp: now + TOKEN_DAYS * 864e5 });
    set_('명단', row, 'PIN실패', 0);
    set_('명단', row, '로그인', JSON.stringify(list));
    set_('명단', row, '마지막접속', new Date());
    CacheService.getScriptCache().put('tok:' + hash_(token), m.name, 21600);
    return { token, state: state_(m.name) };
  },

  me(p) { return { state: state_(auth_(p.token).name) }; },

  logout(p) {
    try {
      const m = auth_(p.token), h = hash_(String(p.token));
      set_('명단', m.row._row, '로그인', JSON.stringify(tokens_(m).filter(t => t.h !== h)));
      CacheService.getScriptCache().remove('tok:' + h);
    } catch (e) { /* 이미 끝난 로그인 */ }
    return {};
  },

  /** 접수 상태일 때만 본인이 취소 */
  cancel(p) {
    const m = auth_(p.token);
    const r = rows_('청구내역').find(x => x['청구번호'] === p.no && x['이름'] === m.name);
    if (!r) throw err_('청구를 찾을 수 없어요.');
    if (r['상태'] !== '접수') throw err_('이미 회계가 확인한 청구라 취소할 수 없어요. 회계에게 연락하세요.');
    set_('청구내역', r._row, '상태', '취소');
    set_('청구내역', r._row, '처리자', m.name + '(본인)');
    set_('청구내역', r._row, '처리시각', new Date());
    return { state: state_(m.name) };
  },

  /** 청구 제출: 서버에서 규정과 한도를 다시 확인한 뒤 기록 */
  submit(p) {
    const m = auth_(p.token), c = p.claim || {};
    const unit = m.units.find(u => u.id === c.unit);
    if (!unit) throw err_('본인 소속의 예산만 청구할 수 있어요.');
    const item = items_().find(i => i.id === c.item && i.type === unit.type);
    if (!item) throw err_('예산 항목을 확인하세요.');
    const today = new Date(), d = parseYmd_(c.date);
    if (!d || d > today) throw err_('사용일을 확인하세요.');
    const half = halfKey_(d);
    if (half !== halfKey_(today)) throw err_('지난 반기 사용분은 이번 반기 예산으로 청구할 수 없어요.');
    const lines = (Array.isArray(c.lines) ? c.lines : []).slice(0, 5)
      .map(l => ({ place: String(l.place || '').trim(), amount: Math.round(Number(l.amount) || 0) }));
    if (!lines.length || lines.some(l => !l.place || l.amount <= 0)) throw err_('사용처와 결제 금액을 확인하세요.');
    if (lines.length > 1 && !item.once) throw err_('이 항목은 결제 1건씩 청구해요.');
    const paid = lines.reduce((s, l) => s + l.amount, 0);
    const place = lines.length > 1 ? lines.map(l => l.place + ' ' + l.amount.toLocaleString() + '원').join(' / ') : lines[0].place;
    const heads = Math.round(Number(c.heads) || 0);
    if (item.heads && heads <= 0) throw err_('참석 인원을 입력하세요.');
    if (item.pre && !c.prereport) throw err_('사역비는 회장단 사전 보고가 필요해요.');
    const receipts = (p.receipts || []).slice(0, 5);
    if (!receipts.length) throw err_('영수증 사진을 올려주세요.');
    if (receipts.length < lines.length) throw err_('결제한 가게마다 영수증을 올려주세요.');
    if (item.photo && !p.group) throw err_('단체사진을 올려주세요.');
    const acct = String(p.acct || '').replace(/\D/g, '');
    if (!p.bank || acct.length < 8) throw err_('입금 계좌를 확인하세요.');

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const st = state_(m.name), k = unit.id + '|' + item.id;
      if (item.once && (st.count[k] || 0) > 0) throw err_('이 항목은 이번 반기에 이미 청구했어요.');
      const limit = item.mode === 'pool' ? Math.max(0, item.amount - (st.used[k] || 0)) : Math.min(heads, unit.size) * item.amount;
      if (limit <= 0) throw err_('남은 예산이 없어요.');
      const amount = Math.min(paid, limit);
      const no = nextNo_();
      const folder = unitFolder_(d, unit.id);
      const clean = v => String(v).replace(/[\s\/\\:*?"<>|]+/g, '');
      const base = [fmt_(d, 'yyyyMMdd'), clean(unit.id), clean(item.id), amount + '원', clean(m.name)].join('_');
      const rLinks = receipts.map((b, i) => saveImage_(folder, b, base + '_영수증' + (i + 1) + '_' + no + '.jpg'));
      const gLink = p.group ? saveImage_(folder, p.group, base + '_단체사진_' + no + '.jpg') : '';
      append_('청구내역', [no, new Date(), m.name, unit.id, item.id, d, place, String(c.memo || ''), paid, heads || '',
        amount, paid - amount, rLinks.join('\n'), gLink, '접수', '', '', '', String(c.fixOf || ''), half, JSON.stringify(lines)]);
      acctSheet_('계좌').appendRow([no, new Date(), m.name, p.bank, acct, m.name]);
      MEMO = {};
      return { no, amount, state: state_(m.name) };
    } finally {
      lock.releaseLock();
    }
  }
,

  /* ───────── 회계 모드 ───────── */
  admin(p) { return { admin: adminState_(admin_(p.token)) }; },

  /** 영수증·단체사진 보기 (청구에 연결된 파일만) */
  photo(p) {
    admin_(p.token);
    const id = String(p.id || '');
    const known = rows_('청구내역').some(r => (String(r['영수증']) + String(r['단체사진'])).indexOf(id) >= 0);
    if (!id || !known) throw err_('사진을 찾을 수 없어요.');
    const blob = DriveApp.getFileById(id).getBlob();
    return { data: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
  },

  /** 확인 / 보류 / 접수로 되돌리기. 다른 회계가 먼저 바꿨으면 막음 */
  setStatus(p) {
    const m = admin_(p.token, true);
    const to = String(p.status || '');
    if (['확인', '보류', '접수'].indexOf(to) < 0) throw err_('잘못된 상태예요.');
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const r = rows_('청구내역').find(x => x['청구번호'] === p.no);
      if (!r) throw err_('청구를 찾을 수 없어요.');
      if (r['상태'] !== p.expect) throw err_('그사이 다른 사람이 처리했어요. 최신 내용으로 바꿨어요.', 'STALE');
      if (to === '보류' && !String(p.reason || '').trim()) throw err_('보류 사유를 적어주세요.');
      set_('청구내역', r._row, '상태', to);
      set_('청구내역', r._row, '보류사유', to === '보류' ? String(p.reason).trim() : '');
      set_('청구내역', r._row, '처리자', m.name);
      set_('청구내역', r._row, '처리시각', new Date());
    } finally { lock.releaseLock(); }
    return { admin: adminState_(m) };
  },

  /** 한 사람에게 송금한 뒤 그 사람의 확인된 청구를 입금 완료로 */
  paid(p) {
    const m = admin_(p.token, true);
    const nos = (p.nos || []).map(String);
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const rows = rows_('청구내역').filter(r => nos.indexOf(r['청구번호']) >= 0);
      if (!rows.length || rows.length !== nos.length) throw err_('청구를 찾을 수 없어요.', 'STALE');
      if (rows.some(r => r['상태'] !== '확인' || r['이름'] !== p.name)) throw err_('그사이 다른 사람이 처리했어요. 최신 내용으로 바꿨어요.', 'STALE');
      const now = new Date();
      rows.forEach(r => {
        set_('청구내역', r._row, '상태', '입금완료');
        set_('청구내역', r._row, '입금일', now);
        set_('청구내역', r._row, '정산회차', fmt_(now));
        set_('청구내역', r._row, '처리자', m.name);
        set_('청구내역', r._row, '처리시각', now);
      });
      const sum = rows.reduce((s, r) => s + (Number(r['청구금액']) || 0), 0);
      append_('정산', [fmt_(now), now, p.name, rows.length, sum, nos.join(', '), 'Y']);
    } finally { lock.releaseLock(); }
    return { admin: adminState_(m) };
  },

  /** 오늘 처리한 입금 완료를 실수로 눌렀을 때 되돌리기 */
  undoPaid(p) {
    const m = admin_(p.token, true);
    const nos = (p.nos || []).map(String), today = fmt_(new Date());
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const rows = rows_('청구내역').filter(r => nos.indexOf(r['청구번호']) >= 0);
      if (rows.some(r => r['상태'] !== '입금완료' || ymdOf_(r['입금일']) !== today)) throw err_('오늘 처리한 건만 되돌릴 수 있어요.', 'STALE');
      rows.forEach(r => {
        set_('청구내역', r._row, '상태', '확인');
        set_('청구내역', r._row, '입금일', '');
        set_('청구내역', r._row, '정산회차', '');
        set_('청구내역', r._row, '처리자', m.name + '(되돌림)');
        set_('청구내역', r._row, '처리시각', new Date());
      });
      const key = nos.join(', ');
      rows_('정산').filter(r => r['청구번호'] === key && yes_(r['입금완료'])).forEach(r => set_('정산', r._row, '입금완료', '되돌림'));
    } finally { lock.releaseLock(); }
    return { admin: adminState_(m) };
  }
};

function nextNo_() {
  const pre = fmt_(new Date(), 'yyMM');
  const n = rows_('청구내역').filter(r => String(r['청구번호']).indexOf(pre + '-') === 0).length + 1;
  return pre + '-' + String(n).padStart(4, '0');
}
/** 회계 모드 화면 데이터. 계좌번호는 '회계' 권한에게만 */
function adminState_(m) {
  const isPay = m.role === '회계';
  const all = rows_('청구내역');
  const ids = v => String(v || '').split('\n').map(u => (/\/d\/([\w-]+)/.exec(u) || [])[1]).filter(Boolean);
  const time = v => v instanceof Date ? fmt_(v, 'M/d HH:mm') : String(v || '');
  const resubbed = {};
  all.forEach(r => { if (r['재청구원번호'] && r['상태'] !== '취소') resubbed[r['재청구원번호']] = r['청구번호']; });
  const card = r => ({
    no: r['청구번호'], name: r['이름'], unit: r['소속'], item: r['예산항목'], date: ymdOf_(r['사용일']), place: r['사용처'], memo: r['내용'],
    paid: Number(r['결제금액']) || 0, amount: Number(r['청구금액']) || 0, own: Number(r['개인부담']) || 0, heads: Number(r['인원']) || 0,
    status: r['상태'], reason: r['보류사유'], submitted: time(r['제출시각']), submittedYmd: r['제출시각'] instanceof Date ? fmt_(r['제출시각']) : '',
    receipts: ids(r['영수증']), group: ids(r['단체사진']), fixOf: r['재청구원번호'] || '', resub: resubbed[r['청구번호']] || '',
    lines: parseLines_(r['결제내역']), by: r['처리자'] || '', at: time(r['처리시각']), payDate: r['입금일'] ? ymdOf_(r['입금일']) : ''
  });
  const byTime = (a, b) => new Date(a['제출시각']) - new Date(b['제출시각']);
  const toCheck = all.filter(r => r['상태'] === '접수').sort(byTime).map(card);
  const held = all.filter(r => r['상태'] === '보류' && !resubbed[r['청구번호']]).sort(byTime).map(card);
  const ready = all.filter(r => r['상태'] === '확인').sort(byTime);
  const accts = {};
  if (isPay && ready.length) acctSheet_('계좌').getDataRange().getValues().slice(1).forEach(r => accts[r[0]] = { bank: r[3], acct: String(r[4]), holder: r[5] });
  const groups = {};
  ready.forEach(r => {
    const g = groups[r['이름']] || (groups[r['이름']] = { name: r['이름'], total: 0, nos: [], claims: [], bank: '', acct: '', holder: '' });
    g.total += Number(r['청구금액']) || 0; g.nos.push(r['청구번호']); g.claims.push(card(r));
    const a = accts[r['청구번호']]; if (a) { g.bank = a.bank; g.acct = a.acct; g.holder = a.holder; }
  });
  const since = new Date(Date.now() - 45 * 864e5);
  const history = all.filter(r => r['상태'] === '입금완료' && r['입금일'] instanceof Date && r['입금일'] >= since)
    .sort((a, b) => b['입금일'] - a['입금일']).map(card);
  const ym = fmt_(new Date(), 'yyyy-MM');
  const month = all.filter(r => ['접수', '확인', '입금완료'].indexOf(r['상태']) >= 0 && ymdOf_(r['사용일']).slice(0, 7) === ym)
    .reduce((s, r) => s + (Number(r['청구금액']) || 0), 0);
  const sum = a => a.reduce((s, x) => s + x.amount, 0);
  const toPay = Object.keys(groups).map(k => groups[k]);
  return {
    role: m.role, isPay, today: fmt_(new Date()),
    summary: { check: { n: toCheck.length, sum: sum(toCheck) }, pay: { n: toPay.length, sum: toPay.reduce((s, g) => s + g.total, 0) }, held: held.length, month },
    toCheck, toPay, held, history
  };
}
function rootFolder_() {
  const id = String(cfg_()['영수증폴더ID'] || '');
  return DriveApp.getFolderById(id);
}
/** 영수증 폴더 / 2026-11 / 미디어팀 — 폴더 id 는 저장해 두고 다시 찾지 않음 */
function childFolder_(parent, name) {
  const key = 'F:' + parent.getId() + ':' + name, id = props_().getProperty(key);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const it = parent.getFoldersByName(name);
  const f = it.hasNext() ? it.next() : parent.createFolder(name);
  props_().setProperty(key, f.getId());
  return f;
}
function unitFolder_(d, unitId) { return childFolder_(childFolder_(rootFolder_(), fmt_(d, 'yyyy-MM')), String(unitId)); }
function saveImage_(folder, dataUrl, name) {
  const m = /^data:([^;]+);base64,(.+)$/.exec(String(dataUrl));
  if (!m || m[1].indexOf('image/') !== 0) throw err_('사진 파일만 올릴 수 있어요.');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name);
  return folder.createFile(blob).getUrl();
}
function acctBook_() { return SpreadsheetApp.openById(String(cfg_()['계좌파일ID'])); }
function acctSheet_(name) { return acctBook_().getSheetByName(name); }

/* ───────── 관리 메뉴 ───────── */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('청구함 관리')
    .addItem('정산 목록 만들기', 'makeSettlement')
    .addItem('입금 완료 처리 (최근 회차)', 'markPaid')
    .addItem('선택한 사람 PIN 초기화', 'resetPin')
    .addSeparator()
    .addItem('마감 메일 지금 보내보기', 'weeklyMail')
    .addItem('자동 메일 켜기 (금요일 마감 직후)', 'installTriggers')
    .addSeparator()
    .addItem('처음 설정 (탭·폴더 만들기)', 'setup')
    .addToUi();
}

/** '확인' 상태이고 아직 정산 안 된 청구를 사람별로 모아 정산 목록과 이체 목록을 만들어요 */
function makeSettlement() {
  const ui = SpreadsheetApp.getUi();
  const list = rows_('청구내역').filter(r => r['상태'] === '확인' && !r['정산회차']);
  if (!list.length) return ui.alert('정산할 청구가 없어요. 청구내역에서 상태를 ‘확인’으로 바꾼 건만 정산돼요.');
  const round = fmt_(new Date());
  const accts = {};
  acctSheet_('계좌').getDataRange().getValues().slice(1).forEach(r => accts[r[0]] = { bank: r[3], acct: String(r[4]), holder: r[5] });
  const byName = {};
  list.forEach(r => {
    const g = byName[r['이름']] || (byName[r['이름']] = { n: 0, sum: 0, nos: [], acct: null });
    g.n++; g.sum += Number(r['청구금액']) || 0; g.nos.push(r['청구번호']);
    g.acct = accts[r['청구번호']] || g.acct;
    set_('청구내역', r._row, '정산회차', round);
  });
  const settle = sh_('정산'), tr = transferSheet_();
  Object.keys(byName).forEach(name => {
    const g = byName[name], a = g.acct || {};
    settle.appendRow([round, new Date(), name, g.n, g.sum, g.nos.join(', '), '']);
    tr.appendRow([round, name, a.bank || '', a.acct || '', a.holder || name, g.sum, '청년부', g.nos.join(', ')]);
  });
  const total = Object.keys(byName).reduce((s, k) => s + byName[k].sum, 0);
  ui.alert('정산 목록을 만들었어요.\n\n' + Object.keys(byName).length + '명 · ' + list.length + '건 · ' + total.toLocaleString() +
    '원\n\n계좌가 들어간 이체 목록은 회계 전용 파일의 ‘이체목록’ 탭에 있어요. 입금을 마치면 ‘입금 완료 처리’를 눌러주세요.');
}
function transferSheet_() {
  const b = acctBook_();
  let s = b.getSheetByName('이체목록');
  if (!s) { s = b.insertSheet('이체목록'); s.appendRow(TRANSFER_HEAD); s.setFrozenRows(1); s.getRange('D:D').setNumberFormat('@'); }
  return s;
}

/** 가장 최근 회차를 입금 완료로 바꾸고, 메일을 등록한 사람에게 알려요 */
function markPaid() {
  const ui = SpreadsheetApp.getUi();
  const open = rows_('정산').filter(r => !yes_(r['입금완료']));
  if (!open.length) return ui.alert('입금 완료 처리할 회차가 없어요.');
  const round = ymdOf_(open[open.length - 1]['회차']);
  const target = open.filter(r => ymdOf_(r['회차']) === round);
  const total = target.reduce((s, r) => s + (Number(r['합계']) || 0), 0);
  if (ui.alert(round + ' 회차 ' + target.length + '명 · ' + total.toLocaleString() + '원을 입금 완료로 바꿀까요?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const today = new Date();
  rows_('청구내역').filter(r => ymdOf_(r['정산회차']) === round && r['상태'] === '확인').forEach(r => {
    set_('청구내역', r._row, '상태', '입금완료');
    set_('청구내역', r._row, '입금일', today);
  });
  target.forEach(r => set_('정산', r._row, '입금완료', 'Y'));
  ui.alert('입금 완료로 바꿨어요.');
}

/** 명단에서 선택한 줄의 PIN 을 초기화 (다음 입장 때 새 PIN 설정) */
function resetPin() {
  const ui = SpreadsheetApp.getUi(), s = SpreadsheetApp.getActiveSheet();
  if (s.getName() !== '명단') return ui.alert('명단 탭에서 초기화할 사람의 줄을 먼저 선택하세요.');
  const row = s.getActiveRange().getRow();
  if (row < 2) return ui.alert('사람의 줄을 선택하세요.');
  const name = s.getRange(row, col_('명단', '이름')).getValue();
  if (ui.alert(name + ' 님의 PIN을 초기화할까요? 다음 입장 때 새 PIN을 정하게 돼요.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  ['PIN해시', '로그인'].forEach(k => set_('명단', row, k, ''));
  set_('명단', row, 'PIN실패', 0);
  ui.alert(name + ' 님의 PIN을 초기화했어요.');
}

/** 금요일 마감 직후: 회계에게 이번 주 청구 요약 메일 한 통 */
function weeklyMail() {
  const c = cfg_();
  const to = String(c['회계메일'] || '').split(',').map(s => s.trim()).filter(Boolean).join(',');
  if (!to) { alert_('설정 탭의 ‘회계메일’에 받을 주소를 적어주세요.'); return; }
  const all = rows_('청구내역');
  const todo = all.filter(r => r['상태'] === '접수');
  const ready = all.filter(r => r['상태'] === '확인' && !r['정산회차']);
  const held = all.filter(r => r['상태'] === '보류' && !all.some(x => x['재청구원번호'] === r['청구번호']));
  const sum = a => a.reduce((s, r) => s + (Number(r['청구금액']) || 0), 0).toLocaleString() + '원';
  const td = 'style="padding:6px 10px;border-bottom:1px solid #eee"';
  const table = list => list.length ? '<table style="border-collapse:collapse;font-size:14px"><tr>' +
    ['번호', '이름', '소속', '항목', '사용일', '청구금액'].map(h => '<th ' + td + ' align="left">' + h + '</th>').join('') + '</tr>' +
    list.map(r => '<tr>' + [r['청구번호'], r['이름'], r['소속'], r['예산항목'], ymdOf_(r['사용일']), (Number(r['청구금액']) || 0).toLocaleString() + '원']
      .map(v => '<td ' + td + '>' + v + '</td>').join('') + '</tr>').join('') + '</table>' : '<p style="color:#888">없어요.</p>';
  const url = book_().getUrl();
  const html =
    '<div style="font-family:sans-serif;color:#2a1f17">' +
    '<h2 style="margin:0 0 4px">이번 주 청구 마감 요약</h2>' +
    '<p style="margin:0 0 16px;color:#666">' + fmt_(new Date(), 'M월 d일') + ' ' + (c['마감시각'] || 18) + '시 마감 기준</p>' +
    '<h3>확인이 필요한 청구 ' + todo.length + '건 · ' + sum(todo) + '</h3>' + table(todo) +
    '<h3>확인 끝, 정산 대기 ' + ready.length + '건 · ' + sum(ready) + '</h3>' + table(ready) +
    (held.length ? '<h3>보류 중 ' + held.length + '건</h3>' + table(held) : '') +
    '<p style="margin-top:20px">사이트에 회계 계정으로 로그인 → 정산 탭에서 확인·보류 → 송금 후 ‘입금 완료’</p>' +
    (c['사이트주소'] ? '<p><a href="' + c['사이트주소'] + '#admin" style="font-weight:700">휴대폰에서 정산하기 (사이트 정산 탭)</a></p>' : '') +
    '<p><a href="' + url + '">청구함 시트 열기</a></p></div>';
  MailApp.sendEmail({ to, subject: '[청년부 청구함] 이번 주 마감: 확인 필요 ' + todo.length + '건 (' + sum(todo) + ')', htmlBody: html });
}

/** 자동 메일 켜기: 금요일(설정의 마감요일) 마감 시각 15~45분 뒤 회계에게 요약 메일.
 *  메일은 이 메뉴를 누른 사람의 Gmail 에서 나가요. */
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'weeklyMail') ScriptApp.deleteTrigger(t);
  });
  const c = cfg_();
  const days = [ScriptApp.WeekDay.SUNDAY, ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY, ScriptApp.WeekDay.WEDNESDAY,
                ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY, ScriptApp.WeekDay.SATURDAY];
  const day = days[Math.max(0, DOW.indexOf(String(c['마감요일'] || '금').trim()))];
  const hour = Number(c['마감시각']) || 18;
  ScriptApp.newTrigger('weeklyMail').timeBased().onWeekDay(day).atHour(hour).nearMinute(30).create();
  alert_('자동 메일을 켰어요.\n\n매주 ' + (c['마감요일'] || '금') + '요일 ' + hour + '시 마감 직후, 회계에게 이번 주 청구 요약 메일이 한 통 가요.\n메일은 지금 이 메뉴를 누른 계정(' + Session.getEffectiveUser().getEmail() + ')에서 보내져요.');
}

/* ───────── 처음 설정 ───────── */
function setup() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  props_().setProperty('BOOK_ID', book.getId());
  secret_();
  Object.keys(HEAD).forEach(name => {
    let s = book.getSheetByName(name);
    if (!s) s = book.insertSheet(name);
    if (s.getLastRow() === 0) s.appendRow(HEAD[name]);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, HEAD[name].length).setFontWeight('bold').setBackground('#FFF2E7');
  });
  const first = book.getSheets()[0];
  if (Object.keys(HEAD).indexOf(first.getName()) < 0 && first.getLastRow() === 0) book.deleteSheet(first);

  // 기본 설정, 폴더, 회계 전용 계좌 파일
  const cfg = cfg_();
  const addCfg = (k, v, d) => { if (!(k in cfg)) sh_('설정').appendRow([k, v, d]); };
  addCfg('공동체암호', '1234', '사이트 입장 암호. 꼭 바꿔서 리더 단톡방에만 공지하세요.');
  addCfg('회계메일', '', '마감 요약 메일 받을 주소. 여러 명은 쉼표로 구분');
  addCfg('마감요일', '금', '일~토 중 하나');
  addCfg('마감시각', 18, '0~23');
  addCfg('사이트주소', '', '금요일 요약 메일에 넣을 사이트 주소 (GitHub Pages 주소)');
  if (!cfg['영수증폴더ID']) {
    const f = DriveApp.createFolder('청년부 청구함 영수증');
    sh_('설정').appendRow(['영수증폴더ID', f.getId(), '영수증 사진이 월별 폴더로 저장돼요. 회계와만 공유하세요.']);
  }
  if (!cfg['계좌파일ID']) {
    const a = SpreadsheetApp.create('청년부 청구함 계좌 (회계 전용)');
    const s = a.getSheets()[0]; s.setName('계좌'); s.appendRow(ACCT_HEAD); s.setFrozenRows(1); s.getRange('E:E').setNumberFormat('@');
    sh_('설정').appendRow(['계좌파일ID', a.getId(), '계좌번호는 이 파일에만 저장돼요. 회계·회장만 공유하세요.']);
  }

  // 입력 실수를 막는 드롭다운
  const dd = (name, key, list) => {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
    sh_(name).getRange(2, col_(name, key), 999, 1).setDataValidation(rule);
  };
  const ddRange = (name, key, src) => {
    const rule = SpreadsheetApp.newDataValidation().requireValueInRange(src, true).setAllowInvalid(false).build();
    sh_(name).getRange(2, col_(name, key), 999, 1).setDataValidation(rule);
  };
  dd('청구내역', '상태', STATUS);
  dd('소속', '유형', TYPES);
  dd('예산항목', '대상유형', TYPES);
  dd('예산항목', '산정방식', ['정액', '인당']);
  ['반기1회', '단체사진', '인원입력', '사전보고'].forEach(k => dd('예산항목', k, ['Y', 'N']));
  dd('명단', '권한', ['회계', '열람']);
  [['소속', '사용'], ['예산항목', '사용'], ['명단', '사용'], ['규정', '노출'], ['FAQ', '노출'], ['FAQ', '자주헷갈림'], ['헷갈리는것', '노출']].forEach(([n, k]) => dd(n, k, ['Y', 'N']));
  const unitList = sh_('소속').getRange('A2:A');
  [1, 2, 3].forEach(k => ddRange('명단', '소속' + k, unitList));
  sh_('청구내역').getRange('A:A').setNumberFormat('@');
  sh_('청구내역').getRange('F:F').setNumberFormat('yyyy-mm-dd');
  sh_('청구내역').getRange('I:L').setNumberFormat('#,##0');
  // 시스템이 쓰는 칸은 회색으로
  ['PIN해시', 'PIN실패', '로그인', '마지막접속'].forEach(k => sh_('명단').getRange(1, col_('명단', k), 1000, 1).setFontColor('#999999'));

  props_().deleteProperty('SCHEMA'); ensureSchema_();
  CACHED_TABS.forEach(forget_);
  seed_();
  alert_('처음 설정을 마쳤어요.\n\n1. 설정 탭에서 공동체암호와 회계메일을 바꿔주세요.\n2. 소속·명단·예산항목·규정·FAQ 탭의 예시 데이터를 실제 내용으로 바꿔주세요.\n3. README의 ‘웹 앱 배포’ 단계를 진행하세요.');
}

/** 탭이 비어 있을 때만 예시(가짜) 데이터를 넣어요 */
function seed_() {
  const put = (name, rows) => { const s = sh_(name); if (s.getLastRow() < 2) s.getRange(2, 1, rows.length, rows[0].length).setValues(rows); };
  put('소속', [['가나셀', '셀', 9, 'Y'], ['다라셀', '셀', 11, 'Y'], ['마바셀', '셀', 8, 'Y'], ['98또래', '또래', 18, 'Y'], ['00또래', '또래', 16, 'Y'],
              ['미디어팀', '팀', 10, 'Y'], ['새가족팀', '팀', 8, 'Y'], ['임원단', '임원단', 5, 'Y']]);
  put('명단', [['김하늘', '가나셀', '셀리더', '', '', '', '', 'Y', '', '', '', '', ''],
              ['이준서', '다라셀', '셀리더', '', '', '', '', 'Y', '', '', '', '', ''],
              ['박서윤', '마바셀', '셀리더', '98또래', '또래장', '', '', 'Y', '', '', '', '', ''],
              ['윤서아', '새가족팀', '팀장', '00또래', '또래장', '', '', 'Y', '', '', '', '', ''],
              ['장도현', '미디어팀', '팀장', '', '', '', '', 'Y', '', '', '', '', ''],
              ['한유진', '임원단', '회장', '', '', '', '', 'Y', '', '', '', '', '회계'],
              ['서다은', '임원단', '회계', '', '', '', '', 'Y', '', '', '', '', '회계']]);
  put('예산항목', [['셀 교제비', '셀', '정액', 80000, '셀당', 'Y', 'Y', 'Y', 'N', '식사, 간식', 'Y'],
                  ['또래 식탁교제비', '또래', '인당', 10000, '인당', 'Y', 'Y', 'Y', 'N', '식사', 'Y'],
                  ['팀 교제비', '팀', '인당', 10000, '인당', 'Y', 'Y', 'Y', 'N', '식사, 간식', 'Y'],
                  ['팀 사역비', '팀', '정액', 100000, '팀당', 'N', 'N', 'N', 'Y', '팀 사역 운영', 'Y'],
                  ['임원단 운영비', '임원단', '정액', 300000, '반기', 'N', 'N', 'N', 'N', '임원 회의, 행사 준비 (가데이터, 금액 확인 필요)', 'Y']]);
  put('규정', [['결제 수단', '예산은 카드(신용·체크)로만 결제해주세요. 현금 결제는 청구할 수 없어요.', '', 'Y'],
              ['인정 증빙', '카드 매출전표만 인정해요. 현금영수증, 계좌이체 내역 등은 안 돼요.', '', 'Y'],
              ['나눠 결제', '교제비는 모임 당일 여러 곳에서 나눠 결제해도 한 번에 청구할 수 있어요. 가게마다 영수증을 올려주세요.', '셀,또래,팀', 'Y'],
              ['영수증 원본', '사진으로 내역을 알아볼 수 있으면 원본은 내지 않아도 돼요.', '', 'Y'],
              ['영수증 분실', '상세 내용이 있는 카드 결제 전표만 가능해요. 결제 내역 목록만으로는 안 돼요.', '', 'Y'],
              ['한도 초과', '지정된 예산을 넘은 금액은 개인 부담이에요. 승인으로도 지원되지 않아요.', '', 'Y'],
              ['사전 보고', '사역비는 회장단에 사용 예정 내역을 먼저 보고하고, 보고한 금액만 청구할 수 있어요.', '팀', 'Y'],
              ['개인 결제', '개인 결제가 함께 찍힌 영수증은 처리할 수 없어요. 따로 결제해주세요.', '', 'Y'],
              ['참석 인원', '교제비는 단체사진이 필수예요. 사진에 나온 인원 수만큼만 지원해요.', '셀,또래,팀', 'Y'],
              ['합동 모임', '여러 소속이 함께 쓴 지출은 처리할 수 없어요.', '', 'Y'],
              ['입금 계좌', '본인 명의 계좌만 가능해요.', '', 'Y'],
              ['보류되는 경우', '영수증 형식이 규정과 다를 때, 사진 속 인원과 신청 인원이 다를 때, 예산 목적과 다른 지출일 때, 소액을 여러 번 나눠 쓴 뒤 따로 청구할 때', '', 'Y']]);
  put('FAQ', [['영수증은 어떤 걸 올려야 하나요?', '신용카드나 체크카드 매출전표만 인정돼요. 현금영수증과 계좌이체 내역은 안 돼요.', '', 'Y', 'Y'],
             ['현금으로 결제했어요.', '현금 결제는 청구할 수 없어요. 예산은 꼭 카드로 결제해주세요.', '', 'Y', 'Y'],
             ['교제비를 두 곳에서 나눠 결제했어요.', '모임 당일 결제라면 한 번에 청구하면 돼요. 청구하기 2단계에서 ‘다른 가게 결제 추가’를 누르고, 가게마다 영수증을 올려주세요.', '셀,또래,팀', 'N', 'Y'],
             ['영수증 사진은 어떻게 찍어야 하나요?', '배경 공백 없이 스캔한 것처럼 영수증만 꽉 차게 올려주세요. 아이폰 메모 앱이나 안드로이드 구글 드라이브의 ‘문서 스캔’을 쓰면 배경이 자동으로 잘려요.', '', 'N', 'Y'],
             ['서로 다른 예산 항목을 합쳐서 한 번에 써도 되나요?', '안 돼요. 각 예산은 용도에 맞는 지출에만 쓸 수 있어요.', '', 'Y', 'Y'],
             ['모임에 일부만 왔어요. 못 온 인원분은 다른 날 써도 되나요?', '안 돼요. 교제비는 모임 당일에 한꺼번에 써야 해요. 못 온 인원분은 지원되지 않아요.', '셀,또래,팀', 'Y', 'Y'],
             ['사역비는 언제 보고해야 하나요?', '쓰기 전에 회장단에 사용 예정 내역을 보고하세요. 보고한 금액까지만 청구할 수 있어요.', '팀', 'N', 'Y'],
             ['예산보다 많이 썼어요.', '초과한 금액은 개인 부담이에요. 사이트에서 청구 금액이 한도까지로 자동 계산돼요.', '', 'N', 'Y'],
             ['언제 입금되나요?', '매주 금요일 18시까지 들어온 청구를 확인해서 금요일부터 일요일 사이에 입금해요. 마감 후 청구는 다음 주에 입금돼요.', '', 'N', 'Y'],
             ['청구를 잘못 냈어요.', '아직 ‘접수’ 상태라면 ‘내 청구’에서 직접 취소하고 다시 내면 돼요. 회계가 확인한 뒤라면 회계에게 연락하세요.', '', 'N', 'Y'],
             ['청구가 보류됐어요.', '‘내 청구’에서 보류 사유를 확인하고 ‘고쳐서 다시 청구’를 누르세요. 기존 내용이 채워진 상태로 열려요.', '', 'N', 'Y'],
             ['영수증을 잃어버렸어요.', '상세 내용이 나온 카드 결제 전표를 카드사 앱 등에서 받아 올려주세요. 결제 내역 목록만으로는 안 돼요.', '', 'N', 'Y'],
             ['다른 사람 계좌로 받아도 되나요?', '본인 명의 계좌만 가능해요.', '', 'N', 'Y'],
             ['개인 주문이 같이 결제됐어요.', '개인 결제가 섞인 영수증은 처리할 수 없어요. 다음부터는 따로 결제해주세요.', '', 'N', 'Y'],
             ['원본 영수증도 내야 하나요?', '사진으로 내역을 알아볼 수 있으면 원본은 필요 없어요.', '', 'N', 'Y'],
             ['청구는 언제까지 해야 하나요?', '사용일로부터 2주 이내, 그리고 사용한 달 말일까지 해주세요. 달이 바뀐 뒤에도 청구는 되지만 피해주세요.', '', 'N', 'Y'],
             ['PIN을 잊어버렸어요.', '회계에게 초기화를 요청하세요. 다음 입장 때 새 PIN을 정하면 돼요.', '', 'N', 'Y'],
             ['명단에 제 이름이 없어요.', '회계에게 이름과 소속을 알려주면 명단에 추가해줘요.', '', 'N', 'Y']]);
  put('헷갈리는것', [['카드 결제, 카드 매출전표만', '현금 결제, 현금영수증, 계좌이체 내역, 결제 내역 목록은 안 돼요.', '', 'Y'],
                    ['교제비는 모임 당일에 한 번에', '반기에 한 번, 모임 날 한꺼번에 써요. 여러 곳에서 나눠 결제했으면 한 번에 청구하고, 못 온 인원분을 다른 날 나눠 쓸 수는 없어요.', '셀,또래,팀', 'Y'],
                    ['예산 항목끼리 합쳐 쓰기 불가', '항목마다 정해진 용도에만 쓸 수 있어요.', '', 'Y'],
                    ['한도 초과분은 개인 부담', '승인을 받아도 초과 금액은 지원되지 않아요.', '', 'Y'],
                    ['개인 결제는 따로', '개인 주문이 함께 찍힌 영수증은 처리할 수 없어요.', '', 'Y']]);
}
