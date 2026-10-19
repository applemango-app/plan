// 서명 틀림 시험용. 이 파일은 t.js의 서명과 짝이 맞지 않으므로 즐겨찾기 버튼이 실행하지 않아야 한다.
// 혹시 실행되면 노란 띠가 나온다(시험 실패 표시). 나이스 값은 건드리지 않는다.
(() => {
  const bar = document.createElement('div');
  bar.textContent = '⚠ 시험 실패: 서명이 틀린 스크립트가 실행됐어요';
  bar.style.cssText = 'position:fixed;top:12px;left:50%;transform:translateX(-50%);z-index:2147483647;background:#e0b000;color:#000;font:600 15px sans-serif;padding:10px 18px;border-radius:8px';
  bar.onclick = () => bar.remove();
  document.body.appendChild(bar);
})();
