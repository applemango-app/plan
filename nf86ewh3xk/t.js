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
    let area = '', summaryCol = -1;
    for (let c = 2; c < head.length; c++) {
      const label = String(head[c] || '').trim();
      if (/^(종합의견|세부능력 및 특기사항)$/.test(label)) summaryCol = c;
      if (!label || TAIL.test(label)) continue;
      if (String((rows[0] || [])[c] || '').trim()) area = String(rows[0][c]).trim();
      columns.push({ col: c, area, label });
    }
    const students = new Map();
    for (const r of rows.slice(2)) {
      const no = Number(String(r[0] || '').trim());
      if (!Number.isInteger(no) || no <= 0) continue;
      students.set(no, { name: String(r[1] || '').trim(), levels: columns.map((x) => String(r[x.col] || '').trim()), summary: summaryCol < 0 ? '' : String(r[summaryCol] || '').trim() });
    }
    return { columns, students, summaryCol };
  }

  const norm = (s) => String(s || '').replace(/[\s,·・‧.\-/()（）]/g, '');
  // 이름 비교: 나이스의 동명이인·전입 표시(괄호 안, 끝의 영문 글자)는 빼고 본다
  const nameKey = (s) => norm(String(s || '').replace(/\([^)]*\)|（[^）]*）/g, '').replace(/\s*[A-Za-z]$/, ''));
  // 같은 학생으로 본다: 표시를 뺀 이름이 같거나, 한쪽이 다른 쪽으로 시작(앱에서 동명이인을 '홍길동2'처럼 구분한 경우)
  const sameName = (a, b) => { const x = nameKey(a), y = nameKey(b); return x === y || (Math.min(x.length, y.length) >= 2 && (x.startsWith(y) || y.startsWith(x))); };

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
    const nameMismatch = [...sheet.students.entries()].filter(([n, s]) => neis.names.has(n) && s.name && !sameName(neis.names.get(n), s.name)).map(([n]) => n);
    return { changes, stats, numbersOnlyInFile, nameMismatch };
  }

  // 행동발달사항 시트: 1행 번호 | 이름 | 행동발달사항 종합의견
  function parseBehaviorSheet(rows) {
    const h = rows[0] || [];
    if (String(h[0]).trim() !== '번호' || !/행동/.test(String(h[2] || ''))) return null;
    const m = new Map();
    for (const r of rows.slice(1)) {
      const no = Number(String(r[0] || '').trim());
      if (Number.isInteger(no) && no > 0) m.set(no, { name: String(r[1] || '').trim(), text: String(r[2] || '').trim() });
    }
    return m;
  }

  // 글 칸 채우기 계산. rows: [{ r, no, name, cur }], source: Map(번호 → { name, text })
  function planText(rows, source, overwrite) {
    const stats = { fill: 0, same: 0, keep: 0, noText: 0 };
    const changes = [];
    const numbers = new Set(rows.map((x) => x.no));
    for (const row of rows) {
      const src = source.get(row.no);
      const text = src ? src.text : '';
      if (!text) { stats.noText++; continue; }
      const cur = String(row.cur || '').trim();
      if (cur === text) { stats.same++; continue; }
      if (cur && !overwrite) { stats.keep++; continue; }
      stats.fill++;
      changes.push({ r: row.r, text });
    }
    const byNo = new Map(rows.map((x) => [x.no, x.name]));
    const numbersOnlyInFile = [...source.keys()].filter((n) => !numbers.has(n));
    const nameMismatch = [...source.entries()].filter(([n, v]) => byNo.has(n) && v.name && !sameName(byNo.get(n), v.name)).map(([n]) => n);
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
    const neis = readNeis(app);
    if (!neis.rows.length) return banner('먼저 과목을 고르고 조회를 눌러 주세요', '#b3261e');
    const { panel, body } = makePanel(`🥭 교과별 평가 · ${neis.subject} (학생 ${neis.numbers.size}명 · 평가 ${neis.items.length}개)`);
    const pick = el('button', 'padding:8px 14px;border-radius:8px;border:1px solid #b77a00;background:#f2a516;font-weight:700;cursor:pointer', '애플망고 결과 엑셀 고르기');
    body.append(el('div', 'margin-bottom:8px', '앱에서 내보낸 수행평가 결과 엑셀(.xlsx)을 고르세요. 파일은 이 컴퓨터 안에서만 읽어요.'), pick);
    const probeBtn = el('button', 'display:block;margin-top:14px;padding:4px 10px;border-radius:6px;border:1px solid #aaa;background:#f4f4f4;font-size:12px;cursor:pointer', '열린 화면 구조 조사 저장 (개발용)');
    probeBtn.onclick = () => { try { probe(); } catch (e) { banner('⛔ 조사 실패: ' + e.message, '#b3261e'); } };
    body.append(probeBtn);
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
      // 행동발달·누가기록·교과 문구 화면은 표에 묶이지 않은 데이터(기록 목록 등)도 열 이름과 줄 수를 본다
      if (/sdlbg|scres10/.test(app.app.id)) {
        try {
          s.dataSets = app.getAllDataControls().filter((d) => d.getColumnNames).map((d) => {
            const cols = d.getColumnNames();
            const n = d.getRowCount ? d.getRowCount() : null;
            const sample = {};
            if (n) for (const c of cols) { const v = d.getValue(0, c); sample[c] = v == null ? null : /Brdt|Flnm|Nm$|InvlNo|Info/.test(c) ? `[${String(v).length}자]` : /^\d{1,3}$/.test(String(v)) || /^(19|20)\d{2}[-.]?\d{2}[-.]?\d{2}$/.test(String(v)) || /^[A-Z]{1,2}\d*$|^[YN]$/.test(String(v)) ? String(v) : `[${String(v).length}자]`; }
            else if (n === null) for (const c of cols) { const v = d.getValue(c); sample[c] = v == null ? null : /Brdt|Flnm|Nm$|InvlNo|Info/.test(c) ? `[${String(v).length}자]` : /^\d{1,8}$|^[YN]$/.test(String(v)) ? String(v) : `[${String(v).length}자]`; }
            // 누가기록 목록은 모든 줄과 줄 상태(새 줄/수정/그대로)까지 본다. 내용은 글자 수만.
            let rows;
            if (n && /GicRec|^dsDt$/.test(d.id)) rows = [...Array(Math.min(n, 30)).keys()].map((r) => {
              const o = { _state: d.getRowState ? d.getRowState(r) : null };
              for (const c of cols) { const v = d.getValue(r, c); o[c] = v == null ? null : /Brdt|Flnm|Nm$|InvlNo|Info|Cn$/.test(c) ? `[${String(v).length}자]` : String(v); }
              return o;
            });
            return { id: d.id, type: d.type, rowCount: n, cols, sample, rows };
          });
        } catch (e) { s.dataSetsError = String(e.message); }
        s.grids2 = ctrls.filter((c) => c && c.type === 'grid').map((g) => ({ id: g.id, dataSet: g.dataSet && g.dataSet.id, rows: g.dataSet && g.dataSet.getRowCount() }));
      }
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

  // ---------- 공통 패널 ----------
  // 제목 줄: [←(메뉴로)] 제목 ………… [—(접기)] [✕(닫기)]. 제목 줄을 끌어 옮긴다. 옮긴 위치는 다음 창에도 쓴다.
  let lastPos = null;
  function makePanel(title, isMenu) {
    document.getElementById(PANEL_ID)?.remove();
    const panel = el('div', 'position:fixed;bottom:96px;right:20px;width:520px;max-height:70vh;display:flex;flex-direction:column;z-index:2147483646;background:#fff;color:#222;font:14px/1.5 "Malgun Gothic",sans-serif;border:2px solid #e8890c;border-radius:12px;box-shadow:0 8px 32px rgba(0,0,0,.3);overflow:hidden');
    panel.id = PANEL_ID;
    if (lastPos) Object.assign(panel.style, { left: lastPos.left, top: lastPos.top, right: 'auto', bottom: 'auto' });
    const head = el('div', 'display:flex;align-items:center;gap:6px;padding:8px 10px;background:#e8890c;color:#fff;cursor:move;user-select:none');
    const icon = (t, tip) => { const b = el('button', 'width:28px;height:28px;flex:none;border:0;border-radius:6px;background:rgba(255,255,255,.18);color:#fff;font:700 15px/1 sans-serif;cursor:pointer;padding:0', t); b.title = tip; b.onmousedown = (e) => e.stopPropagation(); return b; };
    if (!isMenu) { const back = icon('←', '메뉴로'); back.onclick = () => openMenu(); head.append(back); }
    head.append(el('div', 'flex:1;font-weight:700;font-size:15px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis', title));
    const fold = icon('—', '접기');
    const close = icon('✕', '닫기');
    head.append(fold, close);
    const body = el('div', 'padding:12px 16px;overflow:auto');
    panel.append(head, body);
    close.onclick = () => panel.remove();
    fold.onclick = () => {
      const folded = body.style.display !== 'none';
      body.style.display = folded ? 'none' : '';
      fold.textContent = folded ? '□' : '—';
      fold.title = folded ? '펴기' : '접기';
    };
    head.onmousedown = (e) => {
      const r = panel.getBoundingClientRect();
      const dx = e.clientX - r.left, dy = e.clientY - r.top;
      const move = (ev) => {
        const left = Math.max(0, Math.min(window.innerWidth - 120, ev.clientX - dx)) + 'px';
        const top = Math.max(0, Math.min(window.innerHeight - 44, ev.clientY - dy)) + 'px';
        Object.assign(panel.style, { left, top, right: 'auto', bottom: 'auto' });
        lastPos = { left, top };
      };
      const up = () => { document.removeEventListener('mousemove', move); document.removeEventListener('mouseup', up); };
      document.addEventListener('mousemove', move);
      document.addEventListener('mouseup', up);
      e.preventDefault();
    };
    return { panel, body };
  }
  function filePicker(body, hint, onSheets) {
    const pick = el('button', 'padding:8px 14px;border-radius:8px;border:1px solid #b77a00;background:#f2a516;font-weight:700;cursor:pointer', '엑셀 고르기');
    body.append(el('div', 'margin-bottom:8px', hint + ' 파일은 이 컴퓨터 안에서만 읽어요.'), pick);
    pick.onclick = () => {
      const input = el('input');
      input.type = 'file';
      input.accept = '.xlsx,.cell';
      input.onchange = async () => {
        const f = input.files && input.files[0];
        if (!f) return;
        try { onSheets(await readXlsx(await f.arrayBuffer()), f.name); } catch (e) { banner('⛔ ' + e.message, '#b3261e'); }
      };
      input.click();
    };
  }
  function done(body, msg) {
    body.textContent = '';
    body.append(el('div', 'font-size:15px;font-weight:700;color:#1f7a3a;margin-bottom:6px', msg),
      el('div', null, '화면에서 확인한 뒤 나이스의 [저장]을 눌러 주세요. 저장하지 않고 되돌리려면 [조회]를 다시 누르면 돼요.'));
  }

  // ---------- 학기말 종합의견 · 행동특성: 학생마다 글 칸 하나 ----------
  const TEXT_SCREENS = {
    subjText: { ds: 'dsMain', keyCol: 'stdntCn', nameCol: 'stdntNm', textCol: 'gnrlzOpiCn', subjCol: 'sbjtCdNm', grid: 'grdMain' },
    behavior: { ds: 'dsScrgRec', keyCol: 'clsNo', nameCol: 'stuFlnm', textCol: 'gnrlzOpiCn', grid: 'grdMain' },
  };

  function openTextPanel(app, kind, title) {
    const cfg = TEXT_SCREENS[kind];
    const ds = app.lookup(cfg.ds);
    const rows = [];
    for (let r = 0; r < ds.getRowCount(); r++) rows.push({ r, no: Number(ds.getValue(r, cfg.keyCol)), name: ds.getValue(r, cfg.nameCol), cur: ds.getValue(r, cfg.textCol) });
    const subject = cfg.subjCol && rows.length ? ds.getValue(0, cfg.subjCol) : '';
    const { panel, body } = makePanel(`🥭 ${title}${subject ? ' · ' + subject : ''} (학생 ${rows.length}명)`);
    const hint = kind === 'behavior' ? '앱에서 내보낸 행동발달사항 엑셀을 고르세요.' : '앱에서 내보낸 수행평가 결과 엑셀을 고르세요. 과목 시트의 종합의견 열을 넣어요.';
    filePicker(body, hint, (sheets, fileName) => {
      let sources;
      if (kind === 'behavior') sources = sheets.map((x) => ({ name: x.name, map: parseBehaviorSheet(x.rows) })).filter((x) => x.map);
      else sources = sheets.map((x) => ({ name: x.name, p: parseResultSheet(x.rows) })).filter((x) => x.p && x.p.summaryCol >= 0)
        .map((x) => ({ name: x.name, map: new Map([...x.p.students].map(([n, st]) => [n, { name: st.name, text: st.summary }])) }));
      if (!sources.length) throw Error(kind === 'behavior' ? '행동발달사항 시트를 찾지 못했어요' : '종합의견 열이 있는 결과 시트를 찾지 못했어요');
      body.textContent = '';
      body.append(el('div', 'color:#666;margin-bottom:6px', '파일: ' + fileName));
      let si = Math.max(0, sources.findIndex((x) => subject && norm(x.name).startsWith(norm(subject))));
      const sel = el('select', 'margin:0 0 10px 6px;padding:4px');
      sources.forEach((x, i) => { const o = el('option', null, x.name); o.value = i; sel.append(o); });
      sel.value = si;
      const row = el('div', null, '시트:');
      row.append(sel);
      const over = el('label', 'display:block;margin:10px 0');
      const box = el('input');
      box.type = 'checkbox';
      over.append(box, document.createTextNode(' 이미 입력된 글도 엑셀 내용으로 바꾸기'));
      const summary = el('div', 'margin:8px 0;padding:8px;background:#f7f7f7;border-radius:8px;white-space:pre-line');
      const go = el('button', 'padding:8px 14px;border-radius:8px;border:0;background:#1f7a3a;color:#fff;font-weight:700;cursor:pointer', '나이스 화면에 채우기');
      body.append(row, over, summary, go);
      let cur;
      const update = () => {
        cur = planText(rows, sources[si].map, box.checked);
        const st = cur.stats;
        const longest = Math.max(0, ...cur.changes.map((c) => c.text.length));
        const lines = [`채울 학생 ${st.fill}명`, `이미 같은 글 ${st.same}명`];
        if (st.keep) lines.push(`이미 다른 글이 있어 그대로 둘 학생 ${st.keep}명`);
        if (st.noText) lines.push(`엑셀에 글이 없는 학생 ${st.noText}명`);
        if (longest) lines.push(`가장 긴 글 ${longest}자`);
        if (cur.numbersOnlyInFile.length) lines.push(`⚠ 나이스에 없는 번호: ${cur.numbersOnlyInFile.join(', ')}`);
        if (cur.nameMismatch.length) lines.push(`⚠ 이름이 다른 번호: ${cur.nameMismatch.join(', ')} — 다른 반 파일인지 확인하세요`);
        summary.textContent = lines.join('\n');
        go.disabled = !cur.changes.length;
        go.style.opacity = go.disabled ? 0.5 : 1;
      };
      sel.onchange = () => { si = Number(sel.value); update(); };
      box.onchange = update;
      go.onclick = () => {
        for (const c of cur.changes) ds.setValue(c.r, cfg.textCol, c.text);
        try { app.lookup(cfg.grid).redraw(); } catch (_) {}
        done(body, `✔ ${cur.changes.length}명의 글을 채웠어요.`);
      };
      update();
    });
    document.body.appendChild(panel);
  }

  // ---------- 누가기록: 앱 누가기록 엑셀(번호 | 이름 | 날짜 | 누가기록)을 학생마다 넣는다 ----------
  // 나이스 [행추가]를 눌러 나이스가 새 줄을 만들게 하고(학생·학년·오늘 날짜), 날짜와 내용만 바꾼다.
  // 같은 날짜·같은 내용의 기록이 이미 있으면 건너뛴다(두 번 실행해도 중복되지 않게). 저장은 학생마다 선생님이 누른다.
  function toYmd(v) {
    const t = String(v == null ? '' : v).trim();
    let m = t.match(/^(\d{4})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})/);
    if (m) return m[1] + m[2].padStart(2, '0') + m[3].padStart(2, '0');
    if (/^\d{8}$/.test(t)) return t;
    if (/^\d{5}(\.\d+)?$/.test(t)) { // 엑셀 날짜 숫자
      const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(Number(t)) * 86400000);
      return d.getUTCFullYear() + String(d.getUTCMonth() + 1).padStart(2, '0') + String(d.getUTCDate()).padStart(2, '0');
    }
    return '';
  }
  function parseRecordSheet(rows) {
    const h = (rows[0] || []).map((x) => String(x || '').trim());
    if (h[0] !== '번호' || h[2] !== '날짜' || !/누가기록/.test(h[3] || '')) return null;
    const out = [];
    for (const r of rows.slice(1)) {
      const no = Number(String(r[0] || '').trim());
      const text = String(r[3] || '').trim();
      if (!Number.isInteger(no) || no <= 0 || !text) continue;
      out.push({ no, name: String(r[1] || '').trim(), date: toYmd(r[2]), text });
    }
    return out;
  }
  const recKey = (date, text) => date + '|' + norm(text);

  function openRecordPanel(app) {
    const { panel, body } = makePanel('🥭 누가기록');
    const rec = app.lookup('dsGicRecStu');
    const stu = app.lookup('dsStu');
    const roster = new Map(); // 번호 → { row, inv, name }
    for (let r = 0; r < stu.getRowCount(); r++) roster.set(Number(stu.getValue(r, 'clsNo')), { row: r, inv: stu.getValue(r, 'stuInvlNo'), name: stu.getValue(r, 'stuFlnm') });
    const dirtyRows = () => { let n = 0; for (let r = 0; r < rec.getRowCount(); r++) if (rec.getRowState(r) !== 1) n++; return n; };

    // 학생 고르기: 시험에서 확인한 방법을 차례로 쓴다. 그 학생 기록이 불러와질 때까지 기다린다.
    async function selectStudent(no) {
      const s = roster.get(no);
      const loaded = () => app.lookup('dmSearch').getValue('stuInvlNo') === s.inv && (rec.getRowCount() === 0 || rec.getValue(0, 'stuInvlNo') === s.inv);
      if (loaded()) return true;
      const g = app.lookup('grdStu');
      const wait = async (ms) => { const t0 = Date.now(); while (Date.now() - t0 < ms) { if (loaded()) return true; await new Promise((r) => setTimeout(r, 200)); } return loaded(); };
      for (const fn of [() => g.selectRows([s.row]), () => g.focusCell(s.row, 1), () => g.getCellElement(s.row, 1).click()]) {
        try { fn(); } catch (_) { continue; }
        if (await wait(5000)) { await new Promise((r) => setTimeout(r, 300)); return true; }
      }
      return false;
    }

    filePicker(body, '앱에서 내보낸 누가기록 엑셀을 고르세요.', (sheets, fileName) => {
      let all = null;
      for (const x of sheets) { all = parseRecordSheet(x.rows); if (all) break; }
      if (!all) throw Error('누가기록 시트(번호 | 이름 | 날짜 | 누가기록)를 찾지 못했어요');
      const badDate = all.filter((x) => !x.date).length;
      const byNo = new Map();
      for (const x of all) { if (!x.date) continue; if (!byNo.has(x.no)) byNo.set(x.no, []); byNo.get(x.no).push(x); }
      const nos = [...byNo.keys()].sort((a, b) => a - b);
      const missing = nos.filter((n) => !roster.has(n));
      const mismatch = nos.filter((n) => roster.has(n) && byNo.get(n)[0].name && !sameName(roster.get(n).name, byNo.get(n)[0].name));
      const todo = nos.filter((n) => roster.has(n));
      const status = new Map(todo.map((n) => [n, '대기']));

      body.textContent = '';
      body.append(el('div', 'color:#666;margin-bottom:6px', '파일: ' + fileName));
      const lines = [`학생 ${nos.length}명 · 기록 ${all.length - badDate}건`];
      if (badDate) lines.push(`⚠ 날짜를 읽지 못한 기록 ${badDate}건은 빼요`);
      if (missing.length) lines.push(`⚠ 나이스 학생 목록에 없는 번호: ${missing.join(', ')} (빼요)`);
      if (mismatch.length) lines.push(`⚠ 이름이 다른 번호: ${mismatch.join(', ')} — 다른 반 파일인지 확인하세요`);
      lines.push('같은 날짜·같은 내용의 기록이 나이스에 이미 있으면 건너뛰어요.');
      body.append(el('div', 'margin:8px 0;padding:8px;background:#f7f7f7;border-radius:8px;white-space:pre-line', lines.join('\n')));
      const list = el('div', 'margin:8px 0;font-size:13px;max-height:180px;overflow:auto;border:1px solid #eee;border-radius:8px;padding:6px');
      const msg = el('div', 'margin:8px 0;font-weight:700;white-space:pre-line');
      const btnStyle = (bg) => `margin:4px 6px 4px 0;padding:8px 14px;border-radius:8px;border:0;background:${bg};color:#fff;font-weight:700;cursor:pointer`;
      const one = el('button', btnStyle('#8a6d00'));
      const allBtn = el('button', btnStyle('#1f7a3a'), '전체 학생 넣기 (학생마다 자동 저장)');
      const confirmBox = el('div', 'display:none;margin:8px 0;padding:10px;border:2px solid #1f7a3a;border-radius:8px;background:#f2fbf4');
      const stop = el('button', btnStyle('#b3261e') + ';display:none', '멈추기');
      body.append(list, msg, one, allBtn, stop, confirmBox);
      let idx = 0, running = false, stopReq = false;
      const say = (t, bad) => { msg.textContent = t; msg.style.color = bad ? '#b3261e' : '#222'; };
      const drawList = () => {
        list.textContent = '';
        for (const n of todo) list.append(el('div', `padding:2px 4px;${n === todo[idx] ? 'background:#fff7e6;font-weight:700' : ''}`, `${n}번 · ${byNo.get(n).length}건 · ${status.get(n)}`));
        one.textContent = idx < todo.length ? `${todo[idx]}번 학생만 넣기 (저장은 직접)` : '모두 끝났어요';
        one.disabled = allBtn.disabled = running || idx >= todo.length;
        one.style.opacity = allBtn.style.opacity = one.disabled ? 0.5 : 1;
        stop.style.display = running ? 'inline-block' : 'none';
      };

      // 한 학생: 고르고, 없는 기록만 행추가로 넣는다. 저장은 하지 않는다.
      async function fillStudent(no) {
        if (!(await selectStudent(no))) return { err: `${no}번 학생 기록을 불러오지 못했어요` };
        const exact = new Map(), sameText = new Map();
        const bump = (m, k) => m.set(k, (m.get(k) || 0) + 1);
        for (let r = 0; r < rec.getRowCount(); r++) {
          bump(exact, recKey(String(rec.getValue(r, 'ghvrDevEnfcYmd') || ''), rec.getValue(r, 'ghvrDevCn')));
          bump(sameText, norm(rec.getValue(r, 'ghvrDevCn')));
        }
        const take = (m, k) => { const n = m.get(k) || 0; if (!n) return false; m.set(k, n - 1); return true; };
        const mine = byNo.get(no);
        const need = [];
        let skipped = 0;
        for (const x of mine) if (take(exact, recKey(x.date, x.text))) { take(sameText, norm(x.text)); skipped++; } else need.push(x);
        const toAdd = [];
        for (const x of need) if (take(sameText, norm(x.text))) skipped++; else toAdd.push(x);
        let added = 0;
        const addedDates = [];
        for (const x of toAdd) {
          const before = rec.getRowCount();
          app.lookup('btnAdd').click();
          if (rec.getRowCount() <= before) return { err: `${no}번 행추가가 되지 않았어요` };
          const r = rec.getRowCount() - 1;
          rec.setValue(r, 'ghvrDevEnfcYmd', x.date);
          rec.setValue(r, 'ghvrDevCn', x.text);
          addedDates.push(x.date);
          added++;
        }
        try { app.lookup('grdGicRecStu').redraw(); } catch (_) {}
        return { added, skipped, total: rec.getRowCount(), addedDates };
      }

      // 저장 누르기 → '저장되었습니다' 알림의 [확인]을 눌러 닫기 → 모든 줄이 저장된 상태인지 확인
      async function saveNow(expected) {
        const P = cpr.core.Platform.INSTANCE;
        const known = new Set(P.getAllRunningAppInstances());
        app.lookup('btnSaveStu').click();
        const t0 = Date.now();
        let closed = 0;
        while (Date.now() - t0 < 20000) {
          await new Promise((r) => setTimeout(r, 250));
          for (const inst of P.getAllRunningAppInstances()) {
            if (known.has(inst)) continue;
            let ok = null;
            try { ok = inst.getContainer().getAllRecursiveChildren().find((c) => c.type === 'button' && String(c.value || '').trim() === '확인'); } catch (_) {}
            if (!ok) continue;
            let text = '';
            try { text = inst.getContainer().getAllRecursiveChildren().filter((c) => c.type === 'output').map((c) => String(c.value || '')).join(' '); } catch (_) {}
            if (!/저장되었습니다/.test(text)) return { err: '나이스 알림: ' + (text.trim().slice(0, 80) || '(내용 없음)') + ' — 알림을 확인하고 닫아 주세요' };
            try { ok.click(); closed++; } catch (_) {}
            known.add(inst);
          }
          if (closed && dirtyRows() === 0 && rec.getRowCount() >= expected) return { ok: true };
        }
        return { err: closed ? '저장 뒤 목록이 바뀌지 않았어요' : "'저장되었습니다' 알림을 찾지 못했어요 — 화면을 확인해 주세요" };
      }

      one.onclick = async () => {
        if (dirtyRows()) return say('⚠ 지금 학생의 기록을 아직 저장하지 않았어요. 나이스 [저장]을 누른 뒤 다시 눌러 주세요.', true);
        if (idx > 0 && status.get(todo[idx - 1]).startsWith('넣음')) status.set(todo[idx - 1], status.get(todo[idx - 1]).replace('넣음', '저장됨 ✓'));
        const no = todo[idx];
        running = true; drawList(); say(`${no}번 학생을 고르는 중…`);
        const r = await fillStudent(no);
        running = false;
        if (r.err) { drawList(); return say('⚠ ' + r.err + `. 학생 목록에서 ${no}번을 직접 클릭한 뒤 다시 눌러 주세요.`, true); }
        status.set(no, r.added ? `넣음 ${r.added}건${r.skipped ? ` (이미 있음 ${r.skipped}건)` : ''}` : `모두 이미 있음 (${r.skipped}건)`);
        idx++;
        say(r.added ? `✔ ${no}번에 ${r.added}건을 넣었어요. 화면에서 확인하고 나이스 [저장]을 누른 뒤 다음 학생으로 가세요.` : `${no}번은 넣을 기록이 없어요(모두 이미 있음).`);
        drawList();
      };

      allBtn.onclick = () => {
        if (dirtyRows()) return say('⚠ 지금 학생의 기록을 아직 저장하지 않았어요. 나이스 [저장]이나 [조회]를 먼저 눌러 주세요.', true);
        const rest = todo.slice(idx);
        const cnt = rest.reduce((a, n) => a + byNo.get(n).length, 0);
        confirmBox.textContent = '';
        confirmBox.style.display = 'block';
        const yes = el('button', btnStyle('#1f7a3a'), '시작');
        const no = el('button', btnStyle('#777'), '취소');
        confirmBox.append(el('div', 'margin-bottom:6px', `엑셀에 있는 학생 ${rest.length}명의 기록 ${cnt}건 중 나이스에 아직 없는 기록만 넣고, 학생마다 저장해요. (이미 있는 기록은 건너뛰어요) 시작할까요?`), yes, no);
        no.onclick = () => { confirmBox.style.display = 'none'; };
        yes.onclick = () => { confirmBox.style.display = 'none'; runAll(); };
      };

      // 출결 날짜(학생 목록의 atteYmd)에서 날짜를 뽑는다. 저장이 거부됐을 때 겹치는 날짜를 알려 주는 데만 쓴다.
      const atteDates = (no) => {
        const s = roster.get(no);
        let t = '';
        try { t = String(stu.getValue(s.row, 'atteYmd') || ''); } catch (_) {}
        return new Set([...t.matchAll(/(\d{4})[-.]?(\d{2})[-.]?(\d{2})/g)].map((m) => m[1] + m[2] + m[3]));
      };
      const md = (d) => `${Number(d.slice(4, 6))}월 ${Number(d.slice(6, 8))}일`;
      const cont = el('button', btnStyle('#1f7a3a') + ';display:none', '계속 (다음 학생부터)');
      body.insertBefore(cont, confirmBox);
      let addedAll = 0;

      async function runAll() {
        running = true; stopReq = false; cont.style.display = 'none'; drawList();
        while (idx < todo.length) {
          if (stopReq) { say(`멈췄어요. 여기까지 저장했어요. (${idx}/${todo.length}명)`); break; }
          const n = todo[idx];
          say(`${n}번 진행 중… (${idx + 1}/${todo.length}명)`);
          const r = await fillStudent(n);
          if (r.err) { say('⚠ ' + r.err + '. 여기서 멈췄어요.', true); break; }
          if (r.added) {
            const sv = await saveNow(r.total);
            if (sv.err) {
              // 일시 정지: 넣은 기록은 화면에 남겨 두고, 선생님이 고쳐 저장한 뒤 [계속]을 누른다
              const att = atteDates(n);
              const hit = [...new Set(r.addedDates.filter((d) => att.has(d)))].map(md);
              status.set(n, `⚠ 저장 안 됨 · 확인 필요`);
              running = false; drawList();
              cont.style.display = 'inline-block';
              say(`⚠ ${n}번 기록이 저장되지 않았어요.\n${sv.err}\n` +
                (hit.length ? `출결과 겹치는 날짜: ${hit.join(', ')}\n` : '') +
                `나이스 알림을 닫고, 날짜를 고치거나 기록을 지운 뒤 나이스 [저장]을 누르세요. 그다음 [계속]을 누르면 ${todo[idx + 1] ? todo[idx + 1] + '번부터' : '마무리'} 이어서 해요.`, true);
              return;
            }
            status.set(n, `저장됨 ✓ ${r.added}건${r.skipped ? ` (이미 있음 ${r.skipped}건)` : ''}`);
            addedAll += r.added;
          } else status.set(n, `모두 이미 있음 (${r.skipped}건)`);
          idx++;
          drawList();
        }
        running = false;
        if (idx >= todo.length) say(`✔ 끝났어요. 학생 ${todo.length}명, 새 기록 ${addedAll}건을 저장했어요.`);
        drawList();
      }
      cont.onclick = () => {
        if (dirtyRows()) return say('⚠ 아직 [저장]하지 않은 기록이 있어요. 날짜를 고치거나 지운 뒤 나이스 [저장]을 먼저 눌러 주세요.', true);
        status.set(todo[idx], '선생님이 고쳐 저장함 ✓');
        idx++;
        runAll();
      };
      stop.onclick = () => { stopReq = true; say('지금 학생까지 마치고 멈출게요…'); };
      drawList();
    });
    document.body.appendChild(panel);
  }

  // ---------- 시작: 화면 메뉴 (버튼은 늘 보이고, 지금 보고 있는 조회된 화면만 누를 수 있다) ----------
  const SCREENS = [
    { re: /els_scres00_m02$/, ds: 'dsMain', grid: 'grdMain', label: '교과별 평가 (단계)', open: (a) => openPanel(a) },
    { re: /els_scres10_m00$/, ds: 'dsMain', grid: 'grdMain', label: '학기말 종합의견 (과목별)', open: (a) => openTextPanel(a, 'subjText', '학기말 종합의견') },
    { re: /els_sdlbg00_m00$/, ds: 'dsScrgRec', grid: 'grdMain', label: '행동특성 및 종합의견', open: (a) => openTextPanel(a, 'behavior', '행동특성 및 종합의견') },
    { re: /els_sdlbg00_m01$/, ds: 'dsStu', grid: 'grdStu', dirtyDs: 'dsGicRecStu', label: '누가기록', open: (a) => openRecordPanel(a) },
  ];
  if (globalThis.__APPLEMANGO_TEST__) { globalThis.__APPLEMANGO_TEST__ = { readXlsx, parseResultSheet, parseBehaviorSheet, parseRecordSheet, toYmd, planText, autoMatch, plan, norm }; return; }
  if (!window.cpr) return banner('나이스 화면에서 눌러 주세요', '#b3261e');
  // 탭이 화면에 보이는지: 그 화면 표의 실제 크기가 0이 아니면 보이는 탭
  const shown = (a, gridId) => { try { const r = a.lookup(gridId).getActualRect(); return r.width > 0 && r.height > 0; } catch (_) { return false; } };
  function menuState() {
    const running = cpr.core.Platform.INSTANCE.getAllRunningAppInstances();
    return SCREENS.map((sc) => {
      const inst = running.filter((a) => a.app && sc.re.test(a.app.id));
      const a = inst.find((x) => shown(x, sc.grid)) || null;
      let rows = 0;
      try { if (a) rows = a.lookup(sc.ds).getRowCount(); } catch (_) {}
      let why = '';
      if (!inst.length) why = '나이스에서 이 화면을 열고 조회해 주세요';
      else if (!a) why = '이 화면 탭으로 이동하면 쓸 수 있어요';
      else if (!rows) why = '조회를 눌러 학생 목록을 불러와 주세요';
      // 입력했지만 저장하지 않은 줄(새 줄·수정·삭제)
      let dirty = 0;
      for (const x of inst) { try { const d = x.lookup(sc.dirtyDs || sc.ds); for (let r = 0; r < d.getRowCount(); r++) if (d.getRowState(r) !== 1) dirty++; } catch (_) {} }
      return { sc, a, why, dirty };
    });
  }

  function openMenu() {
    const { panel, body } = makePanel('🥭 애플망고 → 나이스', true);
    body.append(el('div', 'margin-bottom:8px', '지금 보고 있는 화면에 넣을 수 있어요. 나이스 화면을 바꾸면 버튼도 따라 바뀌어요.'));
    const list = el('div');
    body.append(list);
    let sig = '';
    const render = () => {
      const st = menuState();
      const now = JSON.stringify(st.map((x) => [x.why, x.dirty, !!x.a]));
      if (now === sig) return;
      sig = now;
      list.textContent = '';
      for (const { sc, a, why, dirty } of st) {
        const on = !why;
        const b = el('button', `display:block;width:100%;text-align:left;margin:6px 0;padding:10px 14px;border-radius:8px;border:1px solid ${on ? '#b77a00' : '#ccc'};background:${on ? '#fff7e6' : '#f4f4f4'};color:${on ? '#222' : '#999'};font-weight:700;cursor:${on ? 'pointer' : 'default'}`);
        b.append(el('div', null, sc.label), el('div', `font-weight:400;font-size:12px;white-space:pre-line;color:${dirty ? '#b3261e' : '#888'}`,
          (why || '누르면 시작해요') + (dirty ? `\n⚠ 입력했지만 아직 [저장]하지 않은 칸이 ${dirty}개 있어요. 저장하거나 [조회]로 지워 주세요.` : '')));
        b.disabled = !on;
        if (on) b.onclick = () => { try { sc.open(a); } catch (e) { banner('⛔ ' + e.message, '#b3261e'); } };
        list.append(b);
      }
    };
    render();
    // 나이스 화면이 바뀔 때(메인 탭·안쪽 탭 전환, 조회)만 다시 본다. 변화가 몰리면 0.3초 모아서 한 번. 메뉴가 닫히면 멈춘다.
    let pending = 0;
    const obs = new MutationObserver((muts) => {
      if (!panel.isConnected) return obs.disconnect();
      if (muts.every((m) => panel.contains(m.target))) return;
      if (pending) return;
      pending = setTimeout(() => { pending = 0; try { render(); } catch (_) {} }, 300);
    });
    obs.observe(document.body, { subtree: true, childList: true, attributes: true, attributeFilter: ['style', 'class'] });
    const probeBtn = el('button', 'display:block;margin-top:14px;padding:4px 10px;border-radius:6px;border:1px solid #aaa;background:#f4f4f4;font-size:12px;cursor:pointer', '열린 화면 구조 조사 저장 (개발용)');
    probeBtn.onclick = () => { try { probe(); } catch (e) { banner('⛔ 조사 실패: ' + e.message, '#b3261e'); } };
    body.append(probeBtn);
    document.body.appendChild(panel);
  }

  // 오른쪽 아래 🥭 버튼: 즐겨찾기를 다시 누르지 않아도 메뉴를 연다(페이지를 새로 고치기 전까지). 다시 불러오면 새것으로 바꾼다.
  document.getElementById('applemango-fab')?.remove();
  const fabWrap = el('div', 'position:fixed;right:20px;bottom:44px;z-index:2147483645');
  fabWrap.id = 'applemango-fab';
  const fab = el('button', 'padding:10px 18px;border-radius:24px;border:3px solid #fff;background:#e8890c;color:#fff;font:700 16px/1 "Malgun Gothic",sans-serif;cursor:pointer;box-shadow:0 4px 16px rgba(0,0,0,.45)', '🥭 애플망고');
  fab.title = '애플망고 메뉴 열기';
  fab.onclick = () => openMenu();
  // 오른쪽 위 작은 X: 애플망고 버튼과 열린 창을 모두 닫는다(다시 쓰려면 즐겨찾기를 누른다)
  const x = el('button', 'position:absolute;top:-10px;right:-10px;width:24px;height:24px;border-radius:50%;border:2px solid #fff;background:#555;color:#fff;font:700 12px/1 sans-serif;cursor:pointer;padding:0', '✕');
  x.title = '애플망고 닫기';
  x.onclick = () => { document.getElementById(PANEL_ID)?.remove(); fabWrap.remove(); };
  fabWrap.append(fab, x);
  document.body.appendChild(fabWrap);
  openMenu();
})();
