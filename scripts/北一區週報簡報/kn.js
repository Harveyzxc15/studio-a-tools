// Keynote 橋接（JXA）：python 端用 osascript -l JavaScript 呼叫
//   dump  <deck.key>                    → 印出整份簡報結構 JSON（非隱藏頁的文字框、表格、圖表、圖片）
//   apply <deck.key> <ops.json> [pdf]   → 依 ops 修改後存檔；給 pdf 路徑就順便匯出 PDF
// 檔案一律用完整路徑開啟、用完一定關閉（不論成功失敗），不碰使用者自己開著的其他文件。

function docAt(kn, path) {
  for (const d of kn.documents()) {
    try { if (d.file().toString() === path) return d; } catch (e) {}
  }
  return null;
}

function openDoc(kn, path) {
  const name = path.split('/').pop();
  // 同一個檔若還在 Keynote 裡：可能是上一次操作剛關、還在關閉中 → 最多等 10 秒；仍在才視為使用者開著
  for (let i = 0; i < 10 && docAt(kn, path); i++) delay(1);
  if (docAt(kn, path)) throw new Error('檔案已在 Keynote 開啟中，請先關閉：' + name);
  kn.open(Path(path));
  for (let i = 0; i < 90; i++) {
    const d = docAt(kn, path);
    if (d) {
      try { d.slides.length; return d; } catch (e) {}   // 確認文件已可讀才交出去
    }
    delay(1);
  }
  throw new Error('開檔逾時：' + path);
}

function dumpTable(t, ti) {
  const R = t.rowCount(), C = t.columnCount();
  const fv = t.cells.formattedValue(), fm = t.cells.formula();
  const cells = [];
  for (let r = 0; r < R; r++) {
    const row = [];
    for (let c = 0; c < C; c++) row.push([fv[r * C + c] ?? '', fm[r * C + c] ?? '']);
    cells.push(row);
  }
  return { i: ti, name: t.name(), rows: R, cols: C, cells };
}

function safe(f, dflt) { try { return f(); } catch (e) { return dflt; } }

function dumpDoc(d) {
  const slides = [];
  d.slides().forEach((s, si) => {
    // 略過的頁也讀（維修明細表常被設成略過，但其他頁的公式仍引用它）；
    // 略過頁上的舊物件偶爾讀不到，逐項容錯，讀不到就當作沒有
    const skipped = s.skipped();
    const tables = [];
    safe(() => s.tables(), []).forEach((t, ti) => { const x = safe(() => dumpTable(t, ti), null); if (x) tables.push(x); });
    const texts = [];
    safe(() => s.textItems(), []).forEach((x, xi) => { const v = safe(() => x.objectText(), null); if (v !== null) texts.push({ i: xi, t: v }); });
    const geo = (o, i) => ({ i, pos: o.position(), w: o.width(), h: o.height() });
    slides.push({
      i: si, skipped, texts, tables,
      charts: safe(() => s.charts().map(geo), []),
      images: safe(() => s.images().map((m, mi) => Object.assign(geo(m, mi), { desc: m.description() })), []),
      lines: safe(() => s.lines().map(geo), []),
    });
  });
  return slides;
}

function cellOf(d, op) { return d.slides[op.s].tables[op.t].rows[op.r].cells[op.c]; }

function applyOps(kn, d, ops) {
  ops.forEach((op, k) => {
    try { applyOne(kn, d, op); }
    catch (e) { throw new Error(`第 ${k} 個操作失敗 ${JSON.stringify(op)}：${e}`); }
  });
}

function applyOne(kn, d, op) {
  {
    const s = d.slides[op.s];
    if (op.op === 'cell') {
      const c = cellOf(d, op);
      if ('v' in op) c.value = op.v;
      if (op.bg) c.backgroundColor = op.bg;
    } else if (op.op === 'text') {            // 依原文比對，避免索引位移
      const x = s.textItems().find(t => t.objectText() === op.old);
      if (x) x.objectText = op.v;
    } else if (op.op === 'deltext') {
      const x = s.textItems().find(t => t.objectText() === op.old);
      if (x) x.delete();
    } else if (op.op === 'delline') {
      const l = s.lines().find(l => Math.abs(l.position().x - op.x) < 2 && Math.abs(l.position().y - op.y) < 2);
      if (l) l.delete();
    } else if (op.op === 'image') {           // 用圖片取代原生圖表或上週的自動圖片
      let old = null;
      if (op.desc) old = s.images().find(m => m.description() === op.desc);
      if (!old && op.chart !== undefined) old = s.charts()[op.chart];
      const p = old.position(), w = old.width(), h = old.height();
      const img = kn.Image({ file: Path(op.file), position: p, width: w, height: h });
      s.images.push(img);
      img.description = op.desc;
      old.delete();
    } else if (op.op === 'yearmark') {        // 圖表上疊的「去年／今年」字樣與分隔虛線：改年份、虛線移到 op.x
      const ys = s.textItems().filter(t => /^\d{4}年$/.test(t.objectText())).sort((a, b) => a.position().x - b.position().x);
      if (ys.length === 2) { ys[0].objectText = op.prev; ys[1].objectText = op.cur; }
      const cur = ys[ys.length - 1];
      const ls = s.lines();
      if (ls.length === 1) {
        const p = ls[0].position(), dx = op.x - p.x;
        ls[0].position = { x: op.x, y: p.y };
        if (cur) { const q = cur.position(); cur.position = { x: q.x + dx, y: q.y }; }
      } else if (cur) {
        cur.position = { x: op.x - cur.width() / 2, y: cur.position().y };
      }
    } else if (op.op === 'rowcount') {        // 表格加列（加在最下面；JXA 沒有插入列指令）
      d.slides[op.s].tables[op.t].rowCount = op.n;
    } else {
      throw new Error('未知操作 ' + op.op);
    }
  }
}

function run(argv) {
  const kn = Application('com.apple.Keynote');
  const [cmd, path] = argv;
  // open：開著留給 chart_gui 操作畫面；saveclose／discard：操作完存檔或放棄後關閉
  if (cmd === 'open') { const d = openDoc(kn, path); return d.name(); }
  if (cmd === 'saveclose' || cmd === 'discard') {
    const d = docAt(kn, path);
    if (!d) throw new Error('文件沒有開著：' + path);
    if (cmd === 'saveclose') d.save();
    d.close({ saving: 'no' });
    return 'ok';
  }
  const d = openDoc(kn, path);
  try {
    if (cmd === 'dump') return JSON.stringify(dumpDoc(d));
    if (cmd === 'apply') {
      ObjC.import('Foundation');
      const txt = $.NSString.stringWithContentsOfFileEncodingError(argv[2], $.NSUTF8StringEncoding, null);
      if (!txt || txt.isNil()) throw new Error('讀不到 ops 檔：' + argv[2]);
      const ops = JSON.parse(txt.js);
      applyOps(kn, d, ops);
      d.save();
      if (argv[3]) kn.export(d, { to: Path(argv[3]), as: 'PDF' });
      return 'ok ' + ops.length;
    }
    throw new Error('未知指令 ' + cmd);
  } finally {
    d.close({ saving: 'no' });
  }
}
