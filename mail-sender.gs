/**
 * 출장비 정산 메일 발송기 (Google Apps Script)
 * - genmixtech.ai@gmail.com 계정으로 설치·배포합니다.
 * - 출장비 정산 페이지에서 보낸 엑셀 파일을 첨부해 Gmail로 발송합니다.
 * 원본: Obsidian 100_Intranet_경영정보/tools/mail-sender.gs
 */

// ===== 관리자 설정 =====
// 받는 사람/참조로 허용할 이메일 도메인. 비워 두면 모든 주소 허용.
// 예: ['genmixtech.com', 'gmail.com']
const ALLOWED_DOMAINS = [];
// 하루 최대 발송 건수 (오남용 방지)
const DAILY_LIMIT = 50;
// 모든 발송 메일을 관리 계정에도 숨은 참조로 남김 (기록용). 끄려면 ''.
const ADMIN_BCC = 'genmixtech.ai@gmail.com';
// 첨부 최대 크기 (바이트)
const MAX_BYTES = 5 * 1024 * 1024;
// =======================

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const to = clean(d.to), cc = clean(d.cc);
    if (!to.length) return out(false, '받는 사람 이메일이 없습니다.');
    const bad = to.concat(cc).filter(m => !isOk(m));
    if (bad.length) return out(false, '허용되지 않은 이메일입니다: ' + bad.join(', '));
    if (!d.file || !d.filename || !/\.xlsx$/i.test(d.filename)) return out(false, '엑셀 파일이 없습니다.');

    const bytes = Utilities.base64Decode(d.file);
    if (bytes.length > MAX_BYTES) return out(false, '첨부 파일이 너무 큽니다.');

    // 하루 발송 건수 제한
    const props = PropertiesService.getScriptProperties();
    const key = 'sent_' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd');
    const n = Number(props.getProperty(key) || 0);
    if (n >= DAILY_LIMIT) return out(false, '오늘 발송 한도를 넘었습니다. 관리자에게 문의하세요.');

    const blob = Utilities.newBlob(bytes,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', d.filename);
    const subject = String(d.subject || '[출장비 정산]').slice(0, 200);
    const body = String(d.body || '').slice(0, 5000);
    const opts = { attachments: [blob], name: '출장비 정산' };
    if (cc.length) { opts.cc = cc.join(','); opts.replyTo = cc[0]; }
    if (ADMIN_BCC) opts.bcc = ADMIN_BCC;

    GmailApp.sendEmail(to.join(','), subject, body, opts);
    props.setProperty(key, String(n + 1));
    return out(true, '발송했습니다.');
  } catch (err) {
    return out(false, '발송 중 오류: ' + err.message);
  }
}

// 배포 확인용 (브라우저로 주소를 열면 표시)
function doGet() { return out(true, '출장비 정산 메일 발송기 동작 중'); }

function clean(v) {
  return String(v || '').split(/[,;\s]+/).map(s => s.trim()).filter(Boolean).slice(0, 20);
}
function isOk(m) {
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(m)) return false;
  if (!ALLOWED_DOMAINS.length) return true;
  return ALLOWED_DOMAINS.indexOf(m.split('@')[1].toLowerCase()) >= 0;
}
function out(ok, message) {
  return ContentService.createTextOutput(JSON.stringify({ ok: ok, message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 설치 직후 한 번 실행해 Gmail 권한을 승인하는 용도
function authorize() { GmailApp.getInboxUnreadCount(); }
