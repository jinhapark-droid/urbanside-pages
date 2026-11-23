/**
 * 26FW 소진율 대시보드 — 업데이트 중계 (Google Apps Script 웹 앱)
 *
 * 공개 대시보드 페이지에는 GitHub 토큰을 둘 수 없어서, 토큰은 이 스크립트의
 * '스크립트 속성 GH_TOKEN'에만 보관하고 페이지 버튼은 이 웹 앱 주소만 연다.
 *  - doGet: 누가 버튼을 누르면 GitHub Actions 워크플로를 실행 (3분 쿨다운)
 *  - scheduled: 매일 09:30·13:30·18:40 KST 정시 실행 (GitHub 예약 실행 지연 보완)
 *
 * 설정: 스크립트 속성 GH_TOKEN = 세분화(fine-grained) 토큰
 *       (저장소 urbanside-pages만, 권한 Actions: Read and write)
 *       편집기에서 setupTriggers 한 번 실행 → 배포 > 웹 앱 (실행: 나, 액세스: 조직 내 모든 사용자 또는 모든 사용자)
 */
const REPO = 'jinhapark-droid/urbanside-pages';
const WORKFLOW = '26fw-sellthrough.yml';
const COOLDOWN_MIN = 3;

function dispatch_() {
  const token = PropertiesService.getScriptProperties().getProperty('GH_TOKEN');
  if (!token) return 'no-token';
  const res = UrlFetchApp.fetch(`https://api.github.com/repos/${REPO}/actions/workflows/${WORKFLOW}/dispatches`, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' },
    payload: JSON.stringify({ ref: 'main' }),
    muteHttpExceptions: true,
  });
  return res.getResponseCode();
}

function doGet() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  const props = PropertiesService.getScriptProperties();
  const last = Number(props.getProperty('LAST') || 0);
  let msg, ok = true;
  try {
    if (Date.now() - last < COOLDOWN_MIN * 60000) {
      msg = '방금 다른 분이 업데이트를 요청했어요. 1~2분 뒤 대시보드에 반영돼요.';
    } else {
      const code = dispatch_();
      if (code === 204) {
        props.setProperty('LAST', String(Date.now()));
        msg = '업데이트를 요청했어요. 1~2분 뒤 대시보드가 알아서 새로고침돼요.';
      } else {
        ok = false;
        msg = `요청이 실패했어요 (${code}). 담당자에게 알려 주세요.`;
      }
    }
  } finally {
    lock.releaseLock();
  }
  const html = `<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<body style="margin:0;font-family:-apple-system,'Apple SD Gothic Neo','Malgun Gothic',sans-serif;background:#f2f4f6;display:flex;min-height:100vh;align-items:center;justify-content:center">
<div style="background:#fff;border-radius:20px;padding:28px 32px;max-width:420px;margin:16px;box-shadow:0 2px 12px rgba(0,0,0,.06)">
<div style="font-size:15px;color:#6b7684;font-weight:600">26FW 소진율 대시보드</div>
<div style="font-size:20px;font-weight:700;margin:8px 0;color:${ok ? '#191f28' : '#f04452'}">${msg}</div>
<div style="font-size:14px;color:#8b95a1">이 탭은 닫아도 돼요.</div></div></body>`;
  return HtmlService.createHtmlOutput(html).setTitle('대시보드 업데이트');
}

function scheduled() {
  dispatch_();
}

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  [[9, 30], [13, 30], [18, 40]].forEach(([h, m]) =>
    ScriptApp.newTrigger('scheduled').timeBased().atHour(h).nearMinute(m).everyDays(1).inTimezone('Asia/Seoul').create());
}
