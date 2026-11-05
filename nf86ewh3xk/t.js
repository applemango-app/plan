// 예전 설치 주소(GitHub)의 🥭 버튼을 누르면 보이는 안내. 도우미는 applemango.app/neis 로 옮겼다(2026-11-05).
(() => {
  const NEW = 'https://applemango.app/neis/';
  const old = document.getElementById('applemango-moved');
  if (old) old.remove();
  const d = document.createElement('div');
  d.id = 'applemango-moved';
  d.style.cssText = 'position:fixed;top:12px;left:50%;transform:translateX(-50%);z-index:2147483647;background:#3A2A1E;color:#fff;font:600 15px sans-serif;padding:12px 18px;border-radius:10px;box-shadow:0 4px 16px rgba(0,0,0,.3)';
  d.append('🥭 애플망고 나이스 버튼의 주소가 바뀌었어요. ');
  const a = document.createElement('a');
  a.href = NEW; a.target = '_blank'; a.rel = 'noopener';
  a.textContent = '새 버튼 설치하기';
  a.style.cssText = 'color:#FFB25B;margin:0 10px';
  const x = document.createElement('button');
  x.textContent = '닫기';
  x.style.cssText = 'background:none;border:1px solid #fff;color:#fff;border-radius:6px;padding:2px 8px;cursor:pointer';
  x.onclick = () => d.remove();
  d.append(a, x);
  document.body.appendChild(d);
})();
