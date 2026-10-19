// 애플망고 나이스 연결 시험 — 화면 제목과 안내 띠만 바꾼다. 나이스 값은 읽거나 바꾸지 않는다.
(() => {
  const P = window.cpr && cpr.core && cpr.core.Platform && cpr.core.Platform.INSTANCE;
  let screen = '';
  try {
    const ids = P ? P.getAllRunningAppInstances().map((a) => a.app && a.app.id).filter(Boolean) : [];
    screen = ids.some((id) => /els_scres00_m02$/.test(id)) ? '교과별 평가 화면 찾음' : '교과별 평가 화면은 아직 안 열림';
  } catch (e) {
    screen = '화면 확인 실패';
  }
  const msg = `✔ 애플망고 연결됨 · 서명 확인 (${P ? '나이스 확인' : '나이스 아님'} · ${screen})`;
  if (!document.title.startsWith('✔')) document.title = '✔ ' + document.title;
  const old = document.getElementById('applemango-test-banner');
  if (old) old.remove();
  const bar = document.createElement('div');
  bar.id = 'applemango-test-banner';
  bar.textContent = msg;
  bar.style.cssText = 'position:fixed;top:12px;left:50%;transform:translateX(-50%);z-index:2147483647;background:#1f7a3a;color:#fff;font:600 15px/1.4 sans-serif;padding:10px 18px;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.25);cursor:pointer';
  bar.title = '누르면 닫혀요';
  bar.onclick = () => bar.remove();
  document.body.appendChild(bar);
  setTimeout(() => bar.remove(), 8000);
})();
