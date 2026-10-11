/**
 * 유류대 지급신청서 HTML용 백엔드 (Google Apps Script 웹앱)  v7 (영수증: 총액만 반영, 공급가액·부가세 행 제외)
 *  - API 키는 "스크립트 속성"에만 보관합니다. HTML에는 키가 들어가지 않습니다.
 *
 * 스크립트 속성
 *   OPINET_KEY            오피넷 무료 API 키
 *   APP_TOKEN             HTML과 맞춰 쓰는 접속 토큰
 *   NCP_KEY_ID / NCP_KEY  (선택) 네이버클라우드 Maps Client ID / Secret  → 주소 → 좌표, 도로거리
 *   (회사명·장소명 검색 — 아래 중 하나)
 *   NAVER_HUB_ID / NAVER_HUB_SECRET
 *                         (선택) 네이버클라우드 NAVER API HUB 키. 둘 다 있을 때만 네이버 장소검색 사용
 *   KAKAO_REST_KEY        (선택) 카카오 developers REST API 키 → 카카오 로컬 키워드 검색 (무료)
 *   LOOKUP_SHEET_ID       (선택) 장소 LOOKUP 구글시트 ID. 첫 탭 A 명칭 | B 주소 | C 별칭(쉼표 구분)
 *   MAIL_DOMAIN           (선택) 메일 수신 허용 도메인, 쉼표 구분 (예: genmixtech.com)
 *
 * v2 변경 후 할 일
 *   1) 편집기에서 authorize() 한 번 실행 → 메일 발송 권한 승인
 *   2) 배포 > 배포 관리 > 수정 > 버전: 새 버전 > 배포  (URL 유지)
 */

const OPINET = { SIDO: '02', PROD: { B027: '휘발유', D047: '경유', K015: 'LPG' } };
const NCP_GEOCODE = 'https://maps.apigw.ntruss.com/map-geocode/v2/geocode';
const NCP_DRIVING = 'https://maps.apigw.ntruss.com/map-direction/v1/driving';
const NAVER_LOCAL = 'https://naverapihub.apigw.ntruss.com/search/v1/local';   // 2026.9 개발자센터 → API HUB 이관
const KAKAO_LOCAL = 'https://dapi.kakao.com/v2/local/search/keyword.json';
const P = k => (PropertiesService.getScriptProperties().getProperty(k) || '').trim();
const out = obj => ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);

/* ───── 공통: 토큰 확인 ───── */
function tokenError_(t) {
  const saved = P('APP_TOKEN'), sent = String(t || '').trim();
  if (!saved) return '백엔드에 APP_TOKEN 스크립트 속성이 없습니다. (속성 이름 철자·저장 여부 확인)';
  if (!sent) return '요청에 토큰이 없습니다. HTML 연결 설정에 토큰을 입력하세요.';
  if (sent !== saved) return `접속 토큰이 맞지 않습니다. (보낸 토큰 ${sent.length}자 / 저장된 토큰 ${saved.length}자)`;
  return '';
}
function checkProps() {
  const all = PropertiesService.getScriptProperties().getProperties();
  Object.keys(all).forEach(k => Logger.log(`[${k}] 길이 ${all[k].length}자, 앞뒤 공백 ${all[k] !== all[k].trim() ? '있음' : '없음'}`));
}
/* 새 권한(메일·드라이브) 승인용: 편집기에서 한 번 실행 */
function authorize() {
  Logger.log('메일 일일 잔여 발송량: ' + MailApp.getRemainingDailyQuota());
  Logger.log('드라이브 확인: ' + DriveApp.getRootFolder().getName());
  if (P('LOOKUP_SHEET_ID')) Logger.log('LOOKUP 시트: ' + SpreadsheetApp.openById(P('LOOKUP_SHEET_ID')).getName());
}

/* ───── 라우팅 ───── */
function doGet(e) {
  try {
    if (!e.parameter.action) return out({ ok: true, msg: '유류대 백엔드 동작 중 (v7)' });
    const te = tokenError_(e.parameter.token); if (te) return out({ ok: false, error: te });
    switch (e.parameter.action) {
      case 'price':    return out({ ok: true, ...gyeonggiPrice_() });
      case 'distance': return out({ ok: true, ...route_(e.parameter.from, e.parameter.to) });
      default:         return out({ ok: false, error: '알 수 없는 요청' });
    }
  } catch (err) { return out({ ok: false, error: String(err.message || err) }); }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const te = tokenError_(body.token); if (te) return out({ ok: false, error: te });
    if (body.action === 'ocr')  return out({ ok: true, ...ocrFiles_(body.files || body.images || []) });
    if (body.action === 'mail') return out({ ok: true, ...sendMail_(body) });
    return out({ ok: false, error: '알 수 없는 요청' });
  } catch (err) { return out({ ok: false, error: String(err.message || err) }); }
}

/* ───── 1. 오피넷 경기 평균가 (1시간 캐시) ───── */
function gyeonggiPrice_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('gg');
  if (hit) return JSON.parse(hit);
  const url = `https://www.opinet.co.kr/api/avgSidoPrice.do?out=json&sido=${OPINET.SIDO}&code=${P('OPINET_KEY')}`;
  const oil = JSON.parse(UrlFetchApp.fetch(url).getContentText()).RESULT.OIL;
  const prices = {};
  oil.filter(o => String(o.SIDOCD).padStart(2, '0') === OPINET.SIDO && OPINET.PROD[o.PRODCD])
     .forEach(o => { prices[OPINET.PROD[o.PRODCD]] = Number(o.PRICE); });
  if (!Object.keys(prices).length) throw new Error('오피넷 응답에 경기 가격이 없습니다. 키를 확인하세요.');
  const res = { region: '경기', date: Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd'), prices };
  cache.put('gg', JSON.stringify(res), 3600);
  return res;
}
function testPrice() { Logger.log(JSON.stringify(gyeonggiPrice_())); }

/* ───── 2. 거리 ─────
 * 위치 찾기: ⓪ LOOKUP 시트(명칭·별칭 → 주소) → ① 카카오 주소검색 → ② 카카오 장소검색 → ③ 네이버 Geocoding → ④ 네이버 API HUB 장소검색
 * 도로거리: ① 네이버 Directions 5 → ② 카카오모빌리티 길찾기
 * 한 단계가 실패(키 없음·권한 오류)해도 다음 단계로 넘어갑니다. 실패 사유는 모아서 보여줍니다.
 */
const KAKAO_ADDR = 'https://dapi.kakao.com/v2/local/search/address.json';
const KAKAO_NAVI = 'https://apis-navi.kakaomobility.com/v1/directions';

function http_(url, headers, tag) {
  const r = UrlFetchApp.fetch(url, { headers, muteHttpExceptions: true });
  const code = r.getResponseCode(), body = r.getContentText();
  if (code !== 200) throw new Error(`${tag} ${code} ${body.slice(0, 120)}`);
  return JSON.parse(body);
}
const kakaoH_ = () => ({ Authorization: 'KakaoAK ' + P('KAKAO_REST_KEY') });
const ncpH_ = () => ({ 'x-ncp-apigw-api-key-id': P('NCP_KEY_ID'), 'x-ncp-apigw-api-key': P('NCP_KEY') });

const FINDERS = [
  ['카카오 주소', () => P('KAKAO_REST_KEY'), q => {
    const d = (http_(`${KAKAO_ADDR}?size=1&query=${encodeURIComponent(q)}`, kakaoH_(), '카카오 주소').documents || [])[0];
    return d && { xy: `${d.x},${d.y}`, label: d.address_name };
  }],
  ['카카오 장소', () => P('KAKAO_REST_KEY'), q => {
    const d = (http_(`${KAKAO_LOCAL}?size=1&query=${encodeURIComponent(q)}`, kakaoH_(), '카카오 장소').documents || [])[0];
    return d && { xy: `${d.x},${d.y}`, label: `${d.place_name} (${d.road_address_name || d.address_name})` };
  }],
  ['네이버 주소', () => P('NCP_KEY_ID'), q => {
    const a = (http_(`${NCP_GEOCODE}?query=${encodeURIComponent(q)}`, ncpH_(), '네이버 주소').addresses || [])[0];
    return a && { xy: `${a.x},${a.y}`, label: a.roadAddress || a.jibunAddress };
  }],
  ['네이버 장소', () => P('NAVER_HUB_ID') && P('NAVER_HUB_SECRET'), q => {   // HUB 구독 키가 있을 때만 (Maps 키로는 401)
    const h = { 'X-NCP-APIGW-API-KEY-ID': P('NAVER_HUB_ID'), 'X-NCP-APIGW-API-KEY': P('NAVER_HUB_SECRET') };
    const it = (http_(`${NAVER_LOCAL}?display=1&query=${encodeURIComponent(q)}`, h, '네이버 장소').items || [])[0];
    if (!it) return null;
    const addr = it.roadAddress || it.address, g = FINDERS[0][1]() ? FINDERS[0][2](addr) : FINDERS[2][2](addr);
    return g && { xy: g.xy, label: `${it.title.replace(/<[^>]+>/g, '')} (${addr})` };
  }],
];

/* LOOKUP 시트: 지도 검색에 안 나오는 이름(예: 젠믹스텍)을 주소로 바꿔서 검색. 공백·대소문자 무시, 10분 캐시 */
const norm_ = s => String(s || '').replace(/\s+/g, '').toLowerCase();
function lookupTable_() {
  const id = P('LOOKUP_SHEET_ID'); if (!id) return {};
  const cache = CacheService.getScriptCache(), c = cache.get('lookup1');
  if (c) return JSON.parse(c);
  const map = {};
  SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getDisplayValues().slice(1).forEach(([name, addr, alias]) => {
    name = String(name || '').trim(); addr = String(addr || '').trim();
    if (!name || !addr) return;                                  // 주소가 빈 행은 건너뜀 → 지도 검색으로
    [name, ...String(alias || '').split(',')].map(norm_).filter(Boolean).forEach(k => { map[k] = { name, addr }; });
  });
  cache.put('lookup1', JSON.stringify(map), 600);
  return map;
}

function locate_(q, log) {
  q = String(q || '').trim();
  if (!q) throw new Error('출발지·도착지를 입력하세요.');
  let row = null;
  try { row = lookupTable_()[norm_(q)] || null; if (row && log) log(`LOOKUP 시트: ${q} → ${row.addr}`); }
  catch (e) { if (log) log('LOOKUP 시트: 오류 ' + e.message); }
  const query = row ? row.addr : q, named = hit => row ? { ...hit, label: `${row.name} (${hit.label})` } : hit;
  const cache = CacheService.getScriptCache(), k = 'loc2:' + query, c = cache.get(k);
  if (c) return named(JSON.parse(c));
  const errs = [];
  for (const [name, ready, fn] of FINDERS) {
    if (!ready()) continue;
    try { const hit = fn(query); if (log) log(`${name}: ${hit ? JSON.stringify(hit) : '결과 없음'}`);
      if (hit) { hit.via = name; cache.put(k, JSON.stringify(hit), 21600); return named(hit); } }
    catch (e) { errs.push(e.message); if (log) log(`${name}: 오류 ${e.message}`); }
  }
  if (row) throw new Error(`'${q}'의 LOOKUP 주소 '${row.addr}'를 지도에서 찾지 못했습니다. 시트의 주소를 확인하세요.` + (errs.length ? ` [${errs.join(' / ')}]` : ''));
  if (!P('KAKAO_REST_KEY') && !P('NCP_KEY_ID')) throw new Error('위치 검색 키가 없습니다. KAKAO_REST_KEY를 등록하세요.');
  throw new Error(`'${q}'을(를) 찾지 못했습니다. 주소나 '회사명 + 지역'으로 입력하세요.` + (errs.length ? ` [${errs.join(' / ')}]` : ''));
}

function route_(from, to, log) {
  const a = locate_(from, log), b = locate_(to, log), errs = [];
  if (P('NCP_KEY_ID')) {
    try {
      const s = http_(`${NCP_DRIVING}?start=${a.xy}&goal=${b.xy}&option=traoptimal`, ncpH_(), '네이버 길찾기').route.traoptimal[0].summary;
      return { km: Math.round(s.distance / 100) / 10, toll: s.tollFare || 0, fromLabel: a.label, toLabel: b.label, source: '네이버 길찾기' };
    } catch (e) { errs.push(e.message); if (log) log('네이버 길찾기: 오류 ' + e.message); }
  }
  if (P('KAKAO_REST_KEY')) {
    try {
      const j = http_(`${KAKAO_NAVI}?origin=${a.xy}&destination=${b.xy}&priority=RECOMMEND`, kakaoH_(), '카카오 길찾기');
      const rt = (j.routes || [])[0];
      if (!rt || rt.result_code !== 0) throw new Error('카카오 길찾기 ' + (rt ? rt.result_msg : '응답 없음'));
      const s = rt.summary;
      return { km: Math.round(s.distance / 100) / 10, toll: (s.fare && s.fare.toll) || 0, fromLabel: a.label, toLabel: b.label, source: '카카오 길찾기' };
    } catch (e) { errs.push(e.message); if (log) log('카카오 길찾기: 오류 ' + e.message); }
  }
  throw new Error('도로거리를 계산하지 못했습니다. ' + errs.join(' / '));
}

/* 진단: 편집기에서 실행 → 실행 로그에서 단계별 결과 확인 (출발지·도착지는 바꿔서) */
/* 진단: LOOKUP 시트를 캐시 없이 다시 읽어 등록된 명칭·별칭 목록을 로그로 확인 */
function testLookup() {
  CacheService.getScriptCache().remove('lookup1');
  const map = lookupTable_();
  Logger.log(`등록 ${Object.keys(map).length}건`);
  Object.keys(map).forEach(k => Logger.log(`${k} → ${map[k].addr}`));
}

function testRoute() {
  CacheService.getScriptCache().removeAll(['loc2:평택시청', 'loc2:수원시청']);
  try { Logger.log('결과: ' + JSON.stringify(route_('평택시청', '수원시청', m => Logger.log(m)))); }
  catch (e) { Logger.log('실패: ' + e.message); }
}

/* ───── 3. 영수증 OCR: 이미지·PDF, 한 장에 여러 건 ───── */
function ocrFiles_(files) {
  const rows = [], raws = [];
  files.forEach(f => {
    const blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.mime || 'image/jpeg', f.name || 'receipt');
    const text = ocr_(blob);
    const found = dropSupplyRows_(parseTolls_(text)).map(r => ({ ...r, file: f.name }));
    rows.push(...found);
    raws.push({ file: f.name, count: found.length, text: text.slice(0, 4000) });
  });
  return { rows, raws };
}
function ocr_(blob) {
  const doc = Drive.Files.create({ name: 'ocr_tmp', mimeType: 'application/vnd.google-apps.document' }, blob, { ocrLanguage: 'ko' });
  try { return DocumentApp.openById(doc.id).getBody().getText(); }
  finally { DriveApp.getFileById(doc.id).setTrashed(true); }
}

const SKIP_LINE = /합계|총\s?액|총\s?금액|소계|누계|잔액|충전|승인번호|카드번호|사업자|대표|전화|주소|부\s?가\s?세|세\s?액/;
const NOT_GATE = /^(통행료|요금|금액|합계|카드|승인|차종|일시|일자|날짜|시간|입구|출구|영수증|하이패스|한국도로공사|구분|원|정상|할인|경차|소형|대형|결제|현금|신용|체크|선불|후불|공급가액|공급|가액|부가세|세액|과세|면세)$/;
const DATE_RE = /(20\d{2})\s?[.\-/년]\s?(\d{1,2})\s?[.\-/월]\s?(\d{1,2})/;
const SUPPLY_RE = /공\s?급\s?가\s?액/;
const VAT_RE = /부\s?가\s?세|세\s?액/;
const LABEL_RE = /공\s?급\s?가\s?액|부\s?가\s?세|세\s?액|합\s?계|통\s?행\s?료|결\s?제\s?금\s?액|금\s?액/g;
const HAS_LABEL = new RegExp(LABEL_RE.source);
const kindOf_ = l => SUPPLY_RE.test(l) ? 's' : VAT_RE.test(l) ? 'v' : 't';     // 공급가액 / 부가세 / 통행료·합계
const ymd_ = d => `${d[1]}-${d[2].padStart(2, '0')}-${d[3].padStart(2, '0')}`;

/* strict: 영업소·TG·IC·출구처럼 톨게이트 표시가 있을 때만 (아무 한글 단어로 대신하지 않음) */
function gateIn_(s, strict) {
  const ex0 = s.match(/출\s?구\s*(?:영업소)?\s*[:：]?\s*([가-힣A-Za-z0-9]{2,12})/); if (ex0 && !NOT_GATE.test(ex0[1])) return ex0[1];
  const lb = s.match(/(?:영업소|요금소|톨게이트)(?:명)?\s*[:：]\s*([가-힣A-Za-z0-9]{2,12})/); if (lb) return lb[1];
  const m = s.match(/([가-힣A-Za-z0-9]{1,10}?)\s?(영업소|요금소|톨게이트|TG|IC|JC|나들목)/);
  if (m) return m[1] + (m[2] === '영업소' || m[2] === '요금소' ? '' : m[2]);
  const ex = s.match(/출구[:\s]*([가-힣]{2,10})/); if (ex) return ex[1];
  if (strict) return '';
  const words = (s.match(/[가-힣]{2,10}/g) || []).filter(w => !NOT_GATE.test(w));
  return words[0] || '';
}
function amountIn_(s) {
  const t = s.replace(DATE_RE, ' ').replace(/\d{1,2}:\d{2}(:\d{2})?/g, ' ').replace(/\*+\d*|\d{4}-\d{4}/g, ' ');
  const nums = (t.match(/\d{1,3}(?:,\d{3})+|\d{3,6}(?=\s*원)|\b\d{3,6}\b/g) || [])
    .map(n => Number(n.replace(/,/g, ''))).filter(n => n >= 100 && n <= 200000 && n % 10 === 0);
  return nums.length ? nums[nums.length - 1] : 0;
}
/* 줄 안의 숫자 목록 (날짜·시각·카드번호 제외). 공급가액은 10원 단위가 아니라서 amountIn_과 따로 */
function numsIn_(s) {
  const t = s.replace(DATE_RE, ' ').replace(/\d{1,2}:\d{2}(:\d{2})?/g, ' ').replace(/\*+\d*|\d{4}-\d{4}/g, ' ');
  return (t.match(/\d{1,3}(?:,\d{3})+|\d+/g) || []).map(n => Number(n.replace(/,/g, ''))).filter(n => n >= 10 && n <= 200000);
}
/* 줄마다 [라벨, 값] 뽑기: "통행료 2,900원", "공급가액 2,636 부가세 264",
 * 또는 "공급가액 부가세 합계" 머리글 + 다음 줄 "2,636 264 2,900" (칸 순서대로 짝지음) */
function labeledValues_(lines) {
  const items = [];
  lines.forEach((line, i) => {
    const ms = [...line.matchAll(LABEL_RE)]; if (!ms.length) return;
    const vals = ms.map((m, k) => numsIn_(line.slice(m.index + m[0].length, k + 1 < ms.length ? ms[k + 1].index : undefined))[0]);
    const next = lines[i + 1] || '';
    if (vals.every(v => !v) && !HAS_LABEL.test(next)) {
      const n = numsIn_(next);
      if (n.length === ms.length) n.forEach((v, k) => { vals[k] = v; });
      else if (ms.length === 1) vals[0] = n[0];
    }
    ms.forEach((m, k) => { if (vals[k]) items.push({ i, kind: kindOf_(m[0]), val: vals[k] }); });
  });
  return items;
}
/* 공급가액이 찍힌 영수증: '공급가액' 한 곳 = 한 건(중복 방지), 금액은 실제 낸 통행료(합계).
 * 통행료 후보 중 공급가액+부가세와 맞는 값을 고르고, 없으면 공급가액+부가세로 계산 (OCR 오독 보정) */
function parseSupply_(lines) {
  const items = labeledValues_(lines), anchors = [];
  let curDate = '', gate = '';
  lines.forEach((line, i) => {
    const d = line.match(DATE_RE);
    if (d) curDate = ymd_(d);
    if (!SUPPLY_RE.test(line)) {
      if (!SKIP_LINE.test(line)) { const g = gateIn_(line, true) || (gate ? '' : gateIn_(line)); if (g) gate = g; }
      return;
    }
    const s = items.find(it => it.kind === 's' && it.i === i);
    if (s) { anchors.push({ i, date: curDate, gate, supply: s.val }); gate = ''; }
  });
  const near = (a, kind) => items.filter(it => it.kind === kind).sort((x, y) => Math.abs(x.i - a.i) - Math.abs(y.i - a.i));
  return anchors.map(a => {
    const vat = (near(a, 'v').find(v => Math.abs(v.val - a.supply / 10) <= 1) || {}).val;
    const want = vat ? a.supply + vat : a.supply * 1.1, tol = vat ? 1 : 1.5;
    const paid = (near(a, 't').find(t => Math.abs(t.val - want) <= tol) || {}).val
      || (vat ? a.supply + vat : Math.round(a.supply * 1.1 / 10) * 10);
    return { date: a.date, gate: a.gate, amount: paid };
  });
}
/* ───── 공급가액·부가세·총액 '관계'로 찾기 (라벨 불필요) ─────
 * 영수증마다 숫자 세 개가 있고 공급가액 + 부가세 = 총액, 부가세 ≈ 공급가액의 10% 입니다.
 * 라벨(합계 등)이 없거나 OCR이 칸 순서를 섞어도, 가까이 있는 숫자들 중 이 관계를 만족하는
 * 세 개를 찾아 가장 큰 값(총액)을 실제 낸 통행료로 씁니다. 한 번 쓴 숫자는 다시 쓰지 않습니다.
 */
function moneyTokens_(text) {
  const t = text
    .replace(/(20\d{2})\s?[.\-/년]\s?\d{1,2}\s?[.\-/월]\s?\d{1,2}일?/g, m => ' '.repeat(m.length))   // 날짜
    .replace(/\d{1,2}:\d{2}(:\d{2})?/g, m => ' '.repeat(m.length))                                   // 시각
    .replace(/[\d*]{2,6}(-[\d*]{2,8}){1,3}/g, m => ' '.repeat(m.length))                              // 카드·전화·사업자번호
    .replace(/\d+\.\d+/g, m => ' '.repeat(m.length));                                                  // 소수(거리 km 등)
  const out = [];
  for (const m of t.matchAll(/\d{1,3}(?:,\d{3})+|\d+/g)) {
    const v = Number(m[0].replace(/,/g, ''));
    if (v >= 10 && v <= 300000) out.push({ v, pos: m.index, used: false });
  }
  return out;
}
const isVat_ = (supply, vat) => supply >= 300 && Math.abs(vat - supply / 10) <= 1;

function parseVatTriples_(text) {
  const nums = moneyTokens_(text), WIN = 8, found = [];
  for (let i = 0; i < nums.length; i++) {
    for (let j = Math.max(0, i - WIN); j < Math.min(nums.length, i + WIN + 1); j++) {
      const a = nums[i], b = nums[j];
      if (i === j || a.used || b.used || !isVat_(a.v, b.v)) continue;
      const lo = Math.max(0, Math.min(i, j) - WIN), hi = Math.min(nums.length, Math.max(i, j) + WIN + 1);
      let c = null;
      for (let k = lo; k < hi; k++) if (k !== i && k !== j && !nums[k].used && nums[k].v === a.v + b.v) { c = nums[k]; break; }
      if (!c) continue;
      a.used = b.used = c.used = true;
      found.push({ supply: a.v, vat: b.v, total: c.v, pos: Math.min(a.pos, b.pos, c.pos) });
      break;
    }
  }
  if (!found.length) return [];
  found.sort((x, y) => x.pos - y.pos);

  // 날짜·톨게이트 위치 목록
  const dates = [...text.matchAll(/(20\d{2})\s?[.\-/년]\s?(\d{1,2})\s?[.\-/월]\s?(\d{1,2})/g)].map(m => ({ pos: m.index, v: ymd_(m) }));
  const gates = []; let off = 0;
  text.split('\n').forEach(line => { const g = !SKIP_LINE.test(line) && gateIn_(line, true); if (g) gates.push({ pos: off, v: g }); off += line.length + 1; });
  // 개수가 같으면 순서대로 짝짓기 (표가 칸 단위로 읽힌 경우), 아니면 바로 앞(없으면 바로 뒤)의 값
  const pick = (list, f, k) => {
    const uniq = list.filter((x, n) => n === 0 || x.v !== list[n - 1].v || x.pos - list[n - 1].pos > 40);
    if (uniq.length === found.length) return uniq[k].v;
    const before = list.filter(x => x.pos <= f.pos).pop();
    return (before || list.find(x => x.pos > f.pos) || {}).v || '';
  };
  // 톨게이트 표시(영업소·TG·출구 등)가 없으면 금액이 있는 줄의 지명을 사용 (예: "2026-10-05 서울 2,636 264 2,900")
  const lineAt = pos => { const st = text.lastIndexOf('\n', pos) + 1, en = text.indexOf('\n', pos); return text.slice(st, en < 0 ? undefined : en); };
  return found.map((f, k) => ({ date: pick(dates, f, k), gate: gates.length ? pick(gates, f, k) : gateIn_(lineAt(f.pos)),
    amount: f.total, supply: f.supply, vat: f.vat }));
}

/* 공급가액(또는 부가세)이 별도 통행료 행으로 잡힌 것을 제거합니다. 실제 낸 금액(총액)만 남깁니다.
 *  - 다른 행 금액이 이 행 금액의 약 1.1배(= 공급가액 + 10%)이면 → 이 행은 공급가액이므로 제외
 *  - 다른 행 금액의 약 1/11(= 부가세)이면 → 이 행은 부가세이므로 제외
 *  - 짝이 없더라도 10원 단위가 아닌 금액은 통행료가 아니므로(공급가액) 제외
 *  - 100원 단위 금액은 실제 통행료로 보고 항상 유지 (같은 날 1,000원·1,100원 두 곳을 지난 경우 보호)
 *  같은 날짜끼리만 비교합니다(날짜가 비어 있으면 모두와 비교). */
function dropSupplyRows_(rows) {
  const sameDay = (a, b) => !a.date || !b.date || a.date === b.date;
  return rows.filter((a, i) => {
    const x = +a.amount || 0; if (!x) return true;
    if (x % 10 !== 0) return false;
    if (x % 100 === 0) return true;          // 100원 단위 금액은 실제 통행료로 보고 유지 (예: 1,000원·1,100원 두 곳 통과)
    return !rows.some((b, j) => {
      if (i === j || !sameDay(a, b)) return false;
      const y = +b.amount || 0;
      return (y > x && isVat_(x, y - x)) || (y > x && isVat_(y - x, x));
    });
  });
}

/* 줄 단위로 [날짜·톨게이트·금액]을 묶어 여러 건 추출 */
function parseTolls_(text) {
  const lines = text.split(/\n+/).map(l => l.trim()).filter(Boolean);
  const tri = parseVatTriples_(text);                   // ① 공급가액+부가세=총액 관계로 찾기 (라벨 불필요)
  if (tri.length) return tri;
  if (lines.some(l => SUPPLY_RE.test(l))) return parseSupply_(lines);   // ② 라벨 기준
  const rows = []; let curDate = '', pendingGate = '';
  lines.forEach(line => {
    const d = line.match(DATE_RE);
    if (d) curDate = `${d[1]}-${d[2].padStart(2, '0')}-${d[3].padStart(2, '0')}`;
    if (SKIP_LINE.test(line)) return;
    const gate = gateIn_(line), amt = amountIn_(line);
    if (amt && (gate || pendingGate)) { rows.push({ date: curDate, gate: gate || pendingGate, amount: amt }); pendingGate = ''; }
    else if (gate && !amt) pendingGate = gate;          // 톨게이트와 금액이 줄이 나뉜 경우
  });
  if (!rows.length) {                                   // 한 건짜리 영수증: 문서 전체에서 한 번 더
    const s = text.replace(/\s+/g, ' ');
    const amt = (s.match(/(?:통행료|요금|금액)[^0-9]{0,8}([\d,]{3,7})/) || [])[1];
    if (amt) { const d = s.match(DATE_RE);
      rows.push({ date: d ? `${d[1]}-${d[2].padStart(2, '0')}-${d[3].padStart(2, '0')}` : '', gate: gateIn_(s), amount: Number(amt.replace(/,/g, '')) }); }
  }
  return rows;
}

/* 영수증 인식 시험: 편집기에서 실행. 인식 원문을 아래 SAMPLE에 붙여넣고 돌려보면 됩니다 */
function testReceipt() {
  const SAMPLE = `통행료 영수증\n2026.10.05 14:22\n출구 서울\n2,636\n264\n2,900\n카드 1234-****-****-5678`;
  Logger.log(JSON.stringify(dropSupplyRows_(parseTolls_(SAMPLE))));
}

/* ───── 4. 메일 발송 (엑셀 첨부) ───── */
function sendMail_(b) {
  const to = String(b.to || '').split(/[,;\s]+/).filter(Boolean);
  const cc = String(b.cc || '').split(/[,;\s]+/).filter(Boolean);
  if (!to.length) throw new Error('받는 사람 이메일을 입력하세요.');
  const bad = [...to, ...cc].filter(a => !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(a));
  if (bad.length) throw new Error('이메일 형식이 맞지 않습니다: ' + bad.join(', '));
  const allow = P('MAIL_DOMAIN').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
  if (allow.length) {
    const out = [...to, ...cc].filter(a => !allow.some(d => a.toLowerCase().endsWith('@' + d)));
    if (out.length) throw new Error('허용되지 않은 수신 도메인입니다: ' + out.join(', '));
  }
  if (to.length + cc.length > 10) throw new Error('수신자는 10명까지 가능합니다.');
  const opts = { name: '유류대 지급신청', htmlBody: String(b.html || '').slice(0, 200000) };
  if (cc.length) opts.cc = cc.join(',');
  if (b.replyTo && /^[^@\s]+@[^@\s]+$/.test(b.replyTo)) opts.replyTo = b.replyTo;
  if (b.xlsx) opts.attachments = [Utilities.newBlob(Utilities.base64Decode(b.xlsx),
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', b.filename || '유류대_지급신청서.xlsx')];
  MailApp.sendEmail(to.join(','), String(b.subject || '유류대 지급신청서').slice(0, 200), '이 메일은 HTML 형식입니다.', opts);
  return { sent: to.length + cc.length, quotaLeft: MailApp.getRemainingDailyQuota() };
}
