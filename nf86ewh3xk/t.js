// 애플망고 나이스 입력 도우미 (시험판)
// 나이스 '교과별 평가' 화면에서 애플망고 결과 엑셀의 평가 단계를 번호·영역에 맞춰 채운다.
// 원칙: 나이스 정보는 밖으로 보내지 않는다(네트워크 요청 없음). 저장은 선생님이 직접 누른다.
// 엑셀 내용은 화면에 글자로만 넣는다(innerHTML 사용 안 함).
(() => {
  const PANEL_ID = 'applemango-panel';

  // ---------- xlsx 읽기 (외부 라이브러리 없이: zip 목차 → deflate 풀기 → XML) ----------
  async function unzip(buf) {
    const dv = new DataView(buf);
    let eocd = -1;
    for (let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw Error('엑셀(xlsx) 파일이 아니에요');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const files = {};
    const dec = new TextDecoder();
    for (let n = 0; n < count; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw Error('엑셀 파일 목차가 이상해요');
      const method = dv.getUint16(p + 10, true);
      const csize = dv.getUint32(p + 20, true);
      const nameLen = dv.getUint16(p + 28, true), extraLen = dv.getUint16(p + 30, true), commentLen = dv.getUint16(p + 32, true);
      const off = dv.getUint32(p + 42, true);
      const name = dec.decode(new Uint8Array(buf, p + 46, nameLen));
      p += 46 + nameLen + extraLen + commentLen;
      files[name] = async () => {
        const start = off + 30 + dv.getUint16(off + 26, true) + dv.getUint16(off + 28, true);
        const data = new Uint8Array(buf, start, csize);
        if (method === 0) return dec.decode(data);
        if (method !== 8) throw Error('지원하지 않는 압축 방식');
        const out = await new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer();
        return dec.decode(out);
      };
    }
    return files;
  }

  const xml = (s) => new DOMParser().parseFromString(s, 'application/xml');
  const tags = (node, name) => Array.from(node.getElementsByTagNameNS('*', name));
  const textOf = (node) => tags(node, 't').map((t) => t.textContent).join('');
  const colIndex = (ref) => { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };

  async function readXlsx(buf) {
    const files = await unzip(buf);
    const get = async (path) => (files[path] ? files[path]() : null);
    const wb = xml(await get('xl/workbook.xml'));
    const rels = xml((await get('xl/_rels/workbook.xml.rels')) || '<r/>');
    const relTarget = {};
    for (const r of tags(rels, 'Relationship')) relTarget[r.getAttribute('Id')] = r.getAttribute('Target');
    const sstXml = await get('xl/sharedStrings.xml');
    const sst = sstXml ? tags(xml(sstXml), 'si').map(textOf) : [];
    const sheets = [];
    for (const s of tags(wb, 'sheet')) {
      const rid = s.getAttribute('r:id') || s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
      let target = relTarget[rid] || '';
      target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
      const sx = await get(target);
      if (!sx) continue;
      const rows = [];
      for (const row of tags(xml(sx), 'row')) {
        const r = Number(row.getAttribute('r')) - 1;
        const cells = [];
        for (const c of tags(row, 'c')) {
          const t = c.getAttribute('t');
          const v = tags(c, 'v')[0];
          let val = '';
          if (t === 's') val = sst[Number(v && v.textContent)] ?? '';
          else if (t === 'inlineStr') val = textOf(c);
          else val = v ? v.textContent : '';
          cells[colIndex(c.getAttribute('r'))] = val;
        }
        rows[r] = Array.from(cells, (x) => (x == null ? '' : String(x)));
      }
      sheets.push({ name: s.getAttribute('name'), rows: Array.from(rows, (r) => r || []) });
    }
    return sheets;
  }

  // ---------- 애플망고 결과 시트 해석 ----------
  // 1행: 영역(같은 영역이 이어지면 첫 칸에만), 2행: 번호 | 이름 | 단원(방법)… | (관찰 기록) | (특이사항) | 종합의견
  const TAIL = /^(관찰 기록|특이사항|종합의견|세부능력 및 특기사항)$/;
  function parseResultSheet(rows) {
    const head = rows[1] || [];
    if (String(head[0]).trim() !== '번호') return null;
    const columns = [];
    let area = '';
    for (let c = 2; c < head.length; c++) {
      const label = String(head[c] || '').trim();
      if (!label || TAIL.test(label)) continue;
      if (String((rows[0] || [])[c] || '').trim()) area = String(rows[0][c]).trim();
      columns.push({ col: c, area, label });
    }
    const students = new Map();
    for (const r of rows.slice(2)) {
      const no = Number(String(r[0] || '').trim());
      if (!Number.isInteger(no) || no <= 0) continue;
      students.set(no, { name: String(r[1] || '').trim(), levels: columns.map((x) => String(r[x.col] || '').trim()) });
    }
    return { columns, students };
  }

  const norm = (s) => String(s || '').replace(/[\s,·・‧.\-/()（）]/g, '');

  // 나이스 평가 항목마다 엑셀 열을 고른다. 영역 이름이 같은 열이 하나뿐일 때만 자동으로 짝짓는다.
  function autoMatch(items, columns) {
    const byArea = new Map();
    for (const it of items) { const k = norm(it.relmNm); byArea.set(k, (byArea.get(k) || 0) + 1); }
    return items.map((it) => {
      const k = norm(it.relmNm);
      const cands = columns.map((c, i) => [c, i]).filter(([c]) => norm(c.area) === k);
      return cands.length === 1 && byArea.get(k) === 1 ? cands[0][1] : -1;
    });
  }

  // 무엇을 바꿀지 계산만 한다(나이스에 아직 넣지 않음)
  function plan(neis, sheet, choice, overwrite) {
    const changes = [];
    const stats = neis.items.map(() => ({ fill: 0, same: 0, keep: 0, noLevel: 0, unknown: 0 }));
    for (const row of neis.rows) {
      const ii = neis.itemIndex.get(row.sccesCtrCd);
      const ci = choice[ii];
      if (ci < 0 || ci == null) continue;
      const st = stats[ii];
      const stu = sheet.students.get(row.clsNo);
      const level = stu ? stu.levels[ci] : '';
      if (!level) { st.noLevel++; continue; }
      const opt = neis.items[ii].options.get(norm(level));
      if (!opt) { st.unknown++; continue; }
      if (row.evlCtrCd.trim() === opt.cd.trim()) { st.same++; continue; }
      if (row.evlCtrCd.trim() && !overwrite) { st.keep++; continue; }
      st.fill++;
      changes.push({ r: row.r, opt });
    }
    const numbersOnlyInFile = [...sheet.students.keys()].filter((n) => !neis.numbers.has(n));
    const nameMismatch = [...sheet.students.entries()].filter(([n, s]) => neis.names.has(n) && s.name && norm(neis.names.get(n)) !== norm(s.name)).map(([n]) => n);
    return { changes, stats, numbersOnlyInFile, nameMismatch };
  }

  // ---------- 나이스 화면 읽기·쓰기 ----------
  function findScreen() {
    const P = window.cpr && cpr.core && cpr.core.Platform && cpr.core.Platform.INSTANCE;
    if (!P) return null;
    return P.getAllRunningAppInstances().find((a) => a.app && /els_scres00_m02$/.test(a.app.id)) || null;
  }

  function readNeis(app) {
    const ds = app.lookup('dsMain'), evl = app.lookup('dsEvlCn');
    const items = [], itemIndex = new Map(), rows = [], numbers = new Set(), names = new Map();
    for (let r = 0; r < ds.getRowCount(); r++) {
      const v = (k) => ds.getValue(r, k);
      const cd = v('sccesCtrCd');
      if (!itemIndex.has(cd)) { itemIndex.set(cd, items.length); items.push({ sccesCtrCd: cd, relmNm: v('relmNm'), lrnGoalCn: v('lrnGoalCn'), options: new Map() }); }
      const clsNo = Number(v('clsNo'));
      numbers.add(clsNo);
      names.set(clsNo, v('stuFlnm'));
      rows.push({ r, clsNo, sccesCtrCd: cd, evlCtrCd: String(v('evlCtrCd') || '') });
    }
    for (let r = 0; r < evl.getRowCount(); r++) {
      const ii = itemIndex.get(evl.getValue(r, 'sccesCtrCd'));
      const cd = String(evl.getValue(r, 'evlCtrCd') || '');
      if (ii == null || cd.trim() === '99') continue;
      items[ii].options.set(norm(evl.getValue(r, 'evlCtrNm')), { cd, nm: evl.getValue(r, 'evlCtrNm'), cn: evl.getValue(r, 'evlCtrCn') });
    }
    return { ds, subject: ds.getRowCount() ? ds.getValue(0, 'sbjtNm') : '', items, itemIndex, rows, numbers, names };
  }

  function apply(app, neis, changes) {
    for (const { r, opt } of changes) {
      neis.ds.setValue(r, 'evlCtrCd', opt.cd);
      neis.ds.setValue(r, 'evlCtrNm', opt.nm);
      neis.ds.setValue(r, 'evlCtrCn', opt.cn);
    }
    try { app.lookup('grdMain').redraw(); } catch (_) {}
  }

  // ---------- 화면 띠·패널 ----------
  function el(tag, style, text) { const e = document.createElement(tag); if (style) e.style.cssText = style; if (text != null) e.textContent = text; return e; }
  function banner(msg, color) {
    const d = el('div', `position:fixed;top:12px;left:50%;transform:translateX(-50%);z-index:2147483647;background:${color};color:#fff;font:600 15px/1.4 sans-serif;padding:10px 18px;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.25);cursor:pointer`, msg);
    d.onclick = () => d.remove();
    document.body.appendChild(d);
    setTimeout(() => d.remove(), 10000);
  }

  function openPanel(app) {
    document.getElementById(PANEL_ID)?.remove();
    const neis = readNeis(app);
    if (!neis.rows.length) return banner('먼저 과목을 고르고 조회를 눌러 주세요', '#b3261e');
    const panel = el('div', 'position:fixed;top:60px;right:20px;width:560px;max-height:80vh;overflow:auto;z-index:2147483646;background:#fff;color:#222;font:14px/1.5 "Malgun Gothic",sans-serif;border:2px solid #f2a516;border-radius:12px;box-shadow:0 8px 32px rgba(0,0,0,.3);padding:16px');
    panel.id = PANEL_ID;
    const close = el('button', 'float:right;border:0;background:none;font-size:20px;cursor:pointer', '✕');
    close.onclick = () => panel.remove();
    panel.append(close, el('div', 'font-weight:700;font-size:16px;margin-bottom:8px', `🥭 애플망고 → 나이스 (${neis.subject}, 학생 ${neis.numbers.size}명 · 평가 ${neis.items.length}개)`));
    const body = el('div');
    const pick = el('button', 'padding:8px 14px;border-radius:8px;border:1px solid #b77a00;background:#f2a516;font-weight:700;cursor:pointer', '애플망고 결과 엑셀 고르기');
    body.append(el('div', 'margin-bottom:8px', '앱에서 내보낸 수행평가 결과 엑셀(.xlsx)을 고르세요. 파일은 이 컴퓨터 안에서만 읽어요.'), pick);
    const probeBtn = el('button', 'display:block;margin-top:14px;padding:4px 10px;border-radius:6px;border:1px solid #aaa;background:#f4f4f4;font-size:12px;cursor:pointer', '열린 화면 구조 조사 저장 (개발용)');
    probeBtn.onclick = () => { try { probe(); } catch (e) { banner('⛔ 조사 실패: ' + e.message, '#b3261e'); } };
    panel.append(body, probeBtn);
    document.body.appendChild(panel);

    pick.onclick = () => {
      const input = el('input');
      input.type = 'file';
      input.accept = '.xlsx,.cell';
      input.onchange = async () => {
        const f = input.files && input.files[0];
        if (!f) return;
        try {
          const sheets = (await readXlsx(await f.arrayBuffer())).map((s) => ({ name: s.name, parsed: parseResultSheet(s.rows) })).filter((s) => s.parsed);
          if (!sheets.length) throw Error('애플망고 결과 시트를 찾지 못했어요');
          showMatch(app, neis, sheets, panel, body, f.name);
        } catch (e) { banner('⛔ ' + e.message, '#b3261e'); }
      };
      input.click();
    };
  }

  function showMatch(app, neis, sheets, panel, body, fileName) {
    body.textContent = '';
    const subj = norm(neis.subject);
    let si = Math.max(0, sheets.findIndex((s) => norm(s.name).startsWith(subj)));
    body.append(el('div', 'color:#666;margin-bottom:6px', '파일: ' + fileName));
    const sheetSel = el('select', 'margin:0 0 10px 6px;padding:4px');
    sheets.forEach((s, i) => { const o = el('option', null, s.name); o.value = i; sheetSel.append(o); });
    sheetSel.value = si;
    const sheetRow = el('div', null, '시트:');
    sheetRow.append(sheetSel);
    const table = el('div');
    const over = el('label', 'display:block;margin:10px 0');
    const overBox = el('input');
    overBox.type = 'checkbox';
    over.append(overBox, document.createTextNode(' 이미 입력된 단계도 엑셀 값으로 바꾸기'));
    const summary = el('div', 'margin:8px 0;padding:8px;background:#f7f7f7;border-radius:8px;white-space:pre-line');
    const go = el('button', 'padding:8px 14px;border-radius:8px;border:0;background:#1f7a3a;color:#fff;font-weight:700;cursor:pointer', '나이스 화면에 채우기');
    body.append(sheetRow, table, over, summary, go);
    let choice = [];

    const render = () => {
      const sheet = sheets[si].parsed;
      table.textContent = '';
      const t = el('table', 'width:100%;border-collapse:collapse;font-size:13px');
      const hr = el('tr');
      ['나이스 영역', '평가요소', '엑셀 열'].forEach((h) => hr.append(el('th', 'text-align:left;border-bottom:1px solid #ccc;padding:4px', h)));
      t.append(hr);
      neis.items.forEach((it, i) => {
        const tr = el('tr');
        tr.append(el('td', 'padding:4px;border-bottom:1px solid #eee', it.relmNm), el('td', 'padding:4px;border-bottom:1px solid #eee', it.lrnGoalCn));
        const sel = el('select', 'max-width:200px;padding:2px');
        const none = el('option', null, '— 넣지 않음 —'); none.value = -1; sel.append(none);
        sheet.columns.forEach((c, ci) => { const o = el('option', null, `${c.area} · ${c.label}`); o.value = ci; sel.append(o); });
        sel.value = choice[i];
        sel.onchange = () => { choice[i] = Number(sel.value); update(); };
        const td = el('td', 'padding:4px;border-bottom:1px solid #eee'); td.append(sel); tr.append(td);
        t.append(tr);
      });
      table.append(t);
      update();
    };
    let current = null;
    const update = () => {
      const sheet = sheets[si].parsed;
      current = plan(neis, sheet, choice, overBox.checked);
      const s = current.stats.reduce((a, x) => { for (const k in x) a[k] = (a[k] || 0) + x[k]; return a; }, {});
      const lines = [`채울 칸 ${s.fill}개`, `이미 같은 값 ${s.same}개`];
      if (s.keep) lines.push(`이미 다른 값이 있어 그대로 둘 칸 ${s.keep}개`);
      if (s.noLevel) lines.push(`엑셀에 단계가 없는 칸 ${s.noLevel}개`);
      if (s.unknown) lines.push(`⚠ 나이스 단계와 이름이 다른 칸 ${s.unknown}개 (예: 3단계/4단계 차이)`);
      if (current.numbersOnlyInFile.length) lines.push(`⚠ 나이스에 없는 번호: ${current.numbersOnlyInFile.join(', ')}`);
      if (current.nameMismatch.length) lines.push(`⚠ 이름이 다른 번호: ${current.nameMismatch.join(', ')} — 다른 반 파일인지 확인하세요`);
      const unmatched = choice.filter((c) => c < 0).length;
      if (unmatched) lines.push(`엑셀 열을 정하지 않은 평가 ${unmatched}개 (위 표에서 골라 주세요)`);
      summary.textContent = lines.join('\n');
      go.disabled = !current.changes.length;
      go.style.opacity = go.disabled ? 0.5 : 1;
    };
    sheetSel.onchange = () => { si = Number(sheetSel.value); choice = autoMatch(neis.items, sheets[si].parsed.columns); render(); };
    overBox.onchange = update;
    go.onclick = () => {
      if (!current || !current.changes.length) return;
      apply(app, neis, current.changes);
      body.textContent = '';
      body.append(el('div', 'font-size:15px;font-weight:700;color:#1f7a3a;margin-bottom:6px', `✔ ${current.changes.length}칸을 채웠어요.`),
        el('div', null, '화면에서 확인한 뒤 나이스의 [저장]을 눌러 주세요. 저장하지 않고 되돌리려면 [조회]를 다시 누르면 돼요.'));
    };
    choice = autoMatch(neis.items, sheets[si].parsed.columns);
    render();
  }

  // ---------- 화면 구조 조사(개발용) — 학생 이름·문구는 글자 수로만. 결과는 이 PC의 다운로드 폴더에 저장 ----------
  function probe() {
    const PERSONAL = /Flnm|Brdt|stdntInfo|stuInvlNo|Cn$|Cn\d$|Nm$/;
    const out = { at: new Date().toISOString(), screens: [] };
    const P = cpr.core.Platform.INSTANCE;
    for (const app of P.getAllRunningAppInstances().filter((a) => a.app && /^edu\//.test(a.app.id))) {
      let ctrls = [];
      try { ctrls = app.getContainer().getAllRecursiveChildren(); } catch (_) {}
      const s = { app: app.app.id, grids: [], buttons: ctrls.filter((c) => c.type === 'button').map((b) => [b.id, b.value]) };
      for (const g of ctrls.filter((c) => c && c.type === 'grid')) {
        const ds = g.dataSet;
        if (!ds || !ds.getRowCount()) continue;
        const cols = ds.getColumnNames();
        const r0 = {};
        for (const c of cols) {
          const v = ds.getValue(0, c);
          // 번호처럼 숫자 1~3자리는 그대로 보여 준다(어느 열이 번호인지 찾기 위해)
          r0[c] = v == null ? null : /^\d{1,3}$/.test(String(v)) ? String(v) : PERSONAL.test(c) ? `[${String(v).length}자]` : typeof v === 'number' ? v : String(v).length > 30 ? `[${String(v).length}자]` : v;
        }
        let detail = null;
        try { detail = [...Array(g.detail.cellCount).keys()].map((i) => { const cell = g.detail.getColumn(i); const ctl = cell && cell.control; return { col: cell && cell.columnName, ctrl: ctl && ctl.type, maxLength: ctl && ctl.maxLength }; }); } catch (_) {}
        s.grids.push({ id: g.id, dataSet: ds.id, rowCount: ds.getRowCount(), cols, firstRow: r0, detail });
      }
      s.textareas = ctrls.filter((c) => c.type === 'textarea').map((t) => ({ id: t.id, maxLength: t.maxLength, len: String(t.value || '').length }));
      if (s.grids.length || s.textareas.length) out.screens.push(s);
    }
    const a = el('a');
    a.href = URL.createObjectURL(new Blob([JSON.stringify(out, null, 1)], { type: 'application/json' }));
    a.download = 'applemango-neis-probe-' + Date.now() + '.json';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
    banner(`✔ 화면 구조를 다운로드 폴더에 저장했어요 (화면 ${out.screens.length}개)`, '#1f7a3a');
  }

  function probePanel(msg) {
    document.getElementById(PANEL_ID)?.remove();
    const panel = el('div', 'position:fixed;top:60px;right:20px;width:360px;z-index:2147483646;background:#fff;color:#222;font:14px/1.5 "Malgun Gothic",sans-serif;border:2px solid #f2a516;border-radius:12px;box-shadow:0 8px 32px rgba(0,0,0,.3);padding:16px');
    panel.id = PANEL_ID;
    const close = el('button', 'float:right;border:0;background:none;font-size:20px;cursor:pointer', '✕');
    close.onclick = () => panel.remove();
    const btn = el('button', 'margin-top:10px;padding:8px 14px;border-radius:8px;border:1px solid #b77a00;background:#f2a516;font-weight:700;cursor:pointer', '이 화면 구조 조사 저장 (개발용)');
    btn.onclick = () => { try { probe(); } catch (e) { banner('⛔ 조사 실패: ' + e.message, '#b3261e'); } };
    panel.append(close, el('div', 'font-weight:700;margin-bottom:6px', '🥭 애플망고'), el('div', null, msg), btn);
    document.body.appendChild(panel);
  }

  // ---------- 시작 ----------
  if (globalThis.__APPLEMANGO_TEST__) { globalThis.__APPLEMANGO_TEST__ = { readXlsx, parseResultSheet, autoMatch, plan, norm }; return; }
  const app = findScreen();
  if (!window.cpr) return banner('나이스 화면에서 눌러 주세요', '#b3261e');
  if (!app || !app.lookup('dsMain').getRowCount()) return probePanel('단계 입력은 [교과별 평가] 화면에서 조회한 뒤 눌러 주세요.');
  openPanel(app);
})();
