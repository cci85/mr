// AI 回覆 → Word 報告（.docx）：各工具頁共用。
// 把貼上的 AI 回覆整理成段落與標題，「〔插入表N〕」換成工具產生的表格（保留合併儲存格、粗體、底線、斜體），
// 不經任何外部程式庫，直接在瀏覽器內組成 .docx 檔下載。
// 「〔插入圖N〕」換成工具產生的圖（PNG），下一行的圖稱置中，再接工具的圖註。
// 用法：DocxReport.bind({ textarea, button, msg, title, nameLine, onNames, tables, figures, ready, filename, toast })
(() => {
  const RFONT = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="新細明體" w:cs="Times New Roman"/>';
  const xesc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const rpr = f => '<w:rPr>' + RFONT + (f.b ? '<w:b/><w:bCs/>' : '') + (f.i ? '<w:i/><w:iCs/>' : '') + (f.u ? '<w:u w:val="single"/>' : '') + '<w:sz w:val="' + (f.sz || 24) + '"/><w:szCs w:val="' + (f.sz || 24) + '"/></w:rPr>';
  const run = (t, f) => '<w:r>' + rpr(f) + '<w:t xml:space="preserve">' + xesc(t) + '</w:t></w:r>';
  const brRun = f => '<w:r>' + rpr(f) + '<w:br/></w:r>';

  /* ---------- 解析 AI 回覆 ---------- */
  const cleanMd = t => t.replace(/^\s*#{1,6}\s*/, "").replace(/^\s*>\s?/, "").replace(/^\s*[-*•]\s+/, "")
    .replace(/\*\*(.+?)\*\*/g, "$1").replace(/__(.+?)__/g, "$1").trim();
  const FIG_RE = /^[〔\[［【]\s*插入圖\s*(\d+)\s*[〕\]］】]$/;
  function parse(txt, o) {
    let lines = String(txt || "").replace(/\r/g, "").split("\n").map(cleanMd);
    const ti = lines.findIndex(l => l === o.title);
    if (ti > 0) lines = lines.slice(ti);              // 去掉報告前的建議（例如「歸屬不明題目的內容判斷」）
    lines = lines.filter(l => l && !(o.nameLine && o.nameLine.test(l)) && !/^[-*_=]{3,}$/.test(l));
    const B = lines.map((l, i) => {
      const m = l.match(/^[〔\[［【]\s*插入表\s*(\d+)\s*[〕\]］】]$/);
      if (m) return { t: "table", n: +m[1] };
      const g = l.match(FIG_RE);
      if (g) return { t: "fig", n: +g[1] };
      if (i > 0 && FIG_RE.test(lines[i - 1]) && /^圖\s*\d+/.test(l) && l.length <= 40) return { t: "figcap", x: l };   // 緊接在〔插入圖N〕之後的圖稱
      if (i === 0 && l === o.title) return { t: "h0", x: l };
      if (l.length <= 30 && !/[。；]$/.test(l)) {
        if (/^（[一二三四五六七八九十]+）\S/.test(l)) return { t: "h1", x: l };
        if (/^(\d+\s*[.．、]\s*\S|第\s*\d+\s*群\s*[：:]\s*\S)/.test(l)) return { t: "h2", x: l };
      }
      return { t: "p", x: l };
    });
    if (!B.length || B[0].t !== "h0") B.unshift({ t: "h0", x: o.title });
    return B;
  }

  /* ---------- HTML 表格 → OOXML ---------- */
  function cellRuns(node, f, out) {
    node.childNodes.forEach(n => {
      if (n.nodeType === 3) { const t = n.textContent.replace(/\s+/g, " "); if (t.trim() || t === " ") out.push(run(t, f)); return; }
      if (n.nodeType !== 1) return;
      const tag = n.tagName;
      if (tag === "BR") { out.push(brRun(f)); return; }
      const g = Object.assign({}, f);
      if (tag === "B" || tag === "STRONG" || /bold|[6-9]00/.test(n.style && n.style.fontWeight || "")) g.b = true;
      if (tag === "I" || tag === "EM") g.i = true;
      if (tag === "U") g.u = true;
      cellRuns(n, g, out);
    });
    return out;
  }
  /* ---------- 欄寬：依內容字數計算，緊密為先 ---------- */
  // 版心寬 = A4 寬 11906 − 左右邊界各 1440；儲存格左右邊距 PAD；中文字寬 = 字級（半點）× 10 twips，英數約 0.55 字
  const PAGE_W = 11906 - 2880, PAD = 57;
  function linesOf(td) {
    const c = td.cloneNode(true);
    c.querySelectorAll("br").forEach(b => b.replaceWith("\n"));
    return c.textContent.split("\n").map(t => t.replace(/\s+/g, " ").trim());
  }
  const wide = ch => /[\u1100-\u115f\u2e80-\ua4cf\uac00-\ud7a3\uf900-\ufaff\ufe30-\ufe4f\uff00-\uff60\uffe0-\uffe6\u3000-\u303f\u2460-\u24ff\u25a0-\u25ff\u2605\u2606\u203b]/.test(ch);
  const textW = (t, sz) => { let n = 0; for (const ch of t) n += wide(ch) ? 1 : 0.55; return Math.ceil(n * sz * 10); };
  // 最小寬度：英數字串不可斷開；中文至少 4 字（內容較短時以內容為準）
  const minW = (t, sz) => {
    const words = t.match(/[\x21-\x7e\u00b0-\u024f\u0370-\u03ff]+/g) || [];
    return Math.max(textW("中中中中", sz), ...words.map(w => textW(w, sz)));
  };
  function fitWidths(grid, ncol, sz) {
    const slack = 2 * PAD + 30, nat = Array(ncol).fill(0), mn = Array(ncol).fill(0);
    grid.forEach(row => row.forEach((cell, c) => {
      if (!cell || !cell.lines || cell.cs > 1) return;
      cell.lines.forEach(t => {
        const w = textW(t, sz) + slack;
        nat[c] = Math.max(nat[c], w);
        mn[c] = Math.max(mn[c], Math.min(w, minW(t, sz) + slack));
      });
    }));
    // 第一欄是列名（例：市場區隔、消費動機），不換行
    mn[0] = Math.min(nat[0], Math.round(PAGE_W * 0.35));
    for (let c = 0; c < ncol; c++) { if (!nat[c]) nat[c] = mn[c] = textW("中", sz) + slack; mn[c] = Math.min(mn[c], nat[c]); }
    // 跨欄儲存格：所跨各欄合計不足時，平均補上
    grid.forEach(row => row.forEach((cell, c) => {
      if (!cell || !cell.lines || cell.cs < 2) return;
      const need = Math.max(...cell.lines.map(t => textW(t, sz))) + slack, have = nat.slice(c, c + cell.cs).reduce((a, b) => a + b, 0);
      if (need > have) for (let k = c; k < c + cell.cs; k++) nat[k] += Math.ceil((need - have) / cell.cs);
    }));
    const sum = a => a.reduce((x, y) => x + y, 0);
    if (sum(nat) <= PAGE_W) return nat;                      // 內容夠短：縮到剛好
    const extra = PAGE_W - sum(mn), give = nat.map((n, i) => n - mn[i]), g = sum(give);
    if (extra <= 0 || !g) return mn;
    return mn.map((m, i) => m + Math.floor(extra * give[i] / g));
  }
  function table(html, opt) {
    const tb = new DOMParser().parseFromString(html, "text/html").querySelector("table");
    const sz = 20, rows = [...tb.rows], grid = rows.map(() => []);   // 表格一律 10 pt
    let ncol = 0;
    rows.forEach((tr, ri) => {
      let c = 0;
      [...tr.cells].forEach(td => {
        while (grid[ri][c]) c++;
        const rs = td.rowSpan || 1, cs = td.colSpan || 1;
        const st = (td.getAttribute("style") || "") + ";" + (td.className || "");
        const al = /text-align:\s*left/.test(st) || /(^|\s|;)l(\s|$)/.test(td.className) ? "left" : "center";
        const b = /font-weight:\s*(bold|[6-9]00)/.test(td.getAttribute("style") || "");
        const runs = cellRuns(td, { sz, b }, []).join("");
        const lines = linesOf(td);
        for (let r = 0; r < rs; r++) {
          if (!grid[ri + r]) grid[ri + r] = [];
          grid[ri + r][c] = r === 0 ? { runs, al, cs, rs, lines, head: td.closest("thead") != null } : { cont: true, cs };
          for (let k = 1; k < cs; k++) grid[ri + r][c + k] = { skip: true };
        }
        c += cs; ncol = Math.max(ncol, c);
      });
    });
    const W = fitWidths(grid, ncol, sz);
    table.lastWidth = W.reduce((a, b) => a + b, 0);
    const ln = '<w:LINE w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
    const bd = ["top", "left", "bottom", "right", "insideH", "insideV"].map(s => ln.replace("LINE", s)).join("");
    let x = '<w:tbl><w:tblPr><w:tblW w:w="' + W.reduce((a, b) => a + b, 0) + '" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders>' + bd + '</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="' + PAD + '" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="' + PAD + '" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>'
      + W.map(w => '<w:gridCol w:w="' + w + '"/>').join("") + '</w:tblGrid>';
    grid.forEach(row => {
      x += '<w:tr><w:trPr><w:cantSplit/>' + (row.some(c => c && c.head) ? '<w:tblHeader/>' : '') + '</w:trPr>';
      for (let c = 0; c < ncol; c++) {
        const cell = row[c];
        if (cell && cell.skip) continue;
        const cs = cell && cell.cs || 1;
        let pr = '<w:tcW w:w="' + W.slice(c, c + cs).reduce((a, b) => a + b, 0) + '" w:type="dxa"/>';
        if (cs > 1) pr += '<w:gridSpan w:val="' + cs + '"/>';
        if (cell && cell.cont) pr += '<w:vMerge/>'; else if (cell && cell.rs > 1) pr += '<w:vMerge w:val="restart"/>';
        pr += '<w:vAlign w:val="center"/>';
        x += '<w:tc><w:tcPr>' + pr + '</w:tcPr><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="' + (cell && cell.al || "center") + '"/></w:pPr>' + (cell && cell.runs || '') + '</w:p></w:tc>';
      }
      x += '</w:tr>';
    });
    return x + '</w:tbl>';
  }

  /* ---------- 段落 ---------- */
  function para(text, o) {
    o = o || {};
    return '<w:p><w:pPr>' + (o.keepNext ? '<w:keepNext/>' : '')
      + '<w:spacing w:before="' + (o.before || 0) + '" w:after="' + (o.after || 0) + '" w:line="' + (o.line || 360) + '" w:lineRule="auto"/>'
      + (o.indent ? '<w:ind w:firstLineChars="200" w:firstLine="480"/>' : o.side ? '<w:ind w:left="' + o.side + '" w:right="' + o.side + '"/>' : '')
      + '<w:jc w:val="' + (o.jc || "both") + '"/></w:pPr>' + (text ? run(text, { b: o.b, sz: o.sz }) : '') + '</w:p>';
  }
  // 圖：寬度依 PNG 的顯示尺寸（px，96 dpi），超過版心時等比縮小
  const EMU = 9525, MAXW = PAGE_W * 635;
  function figXml(F, id) {
    let cx = Math.round(F.w * EMU), cy = Math.round(F.h * EMU);
    if (cx > MAXW) { cy = Math.round(cy * MAXW / cx); cx = MAXW; }
    return '<w:p><w:pPr><w:keepNext/><w:spacing w:before="120" w:after="60" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing>'
      + '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:docPr id="' + id + '" name="' + xesc(F.cap || ("圖" + id)) + '"/>'
      + '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
      + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
      + '<pic:nvPicPr><pic:cNvPr id="' + id + '" name="fig' + id + '.png"/><pic:cNvPicPr/></pic:nvPicPr>'
      + '<pic:blipFill><a:blip r:embed="rIdFig' + id + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
      + '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
      + '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
  }
  function bodyXml(blocks, tables, figures) {
    figures = figures || {};
    let b = "";
    blocks.forEach((k, i) => {
      if (k.t === "table") {
        const T = tables[k.n]; if (!T) return;
        // 表格置中；表號與註對齊表格左緣（寬度與表格相同）
        const tx = table(T.html, T), side = Math.max(0, Math.floor((PAGE_W - table.lastWidth) / 2));
        b += para(T.cap, { jc: "left", keepNext: true, before: 120, after: 60, line: 240, side })
          + tx
          + (T.note ? para(T.note, { sz: 20, line: 240, before: 40, side }) : '')
          + para("", { line: 240 });
      }
      else if (k.t === "fig") {
        const F = figures[k.n]; if (!F || !F.png) return;
        b += figXml(F, k.n);
        // 圖稱：沿用回覆中緊接的圖稱；沒有時用工具的圖稱
        const nx = blocks[i + 1];
        if (!(nx && nx.t === "figcap")) b += figCap(F.cap, F.note);
      }
      else if (k.t === "figcap") {
        const pv = blocks[i - 1], F = pv && pv.t === "fig" ? figures[pv.n] : null;
        b += figCap(k.x, F ? F.note : "");
      }
      else if (k.t === "h0") b += para(k.x, { b: true, sz: 28, jc: "center", after: 240 });
      else if (k.t === "h1") b += para(k.x, { b: true, jc: "left", keepNext: true, before: 240, after: 60 });
      else if (k.t === "h2") b += para(k.x, { b: true, jc: "left", keepNext: true, before: 120 });
      else b += para(k.x, { indent: true });
    });
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>' + b
      + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>';
  }

  function figCap(cap, note) {
    return para(cap || "", { jc: "center", line: 240, after: note ? 40 : 0 })
      + (note ? para(note, { sz: 20, line: 240, jc: "left" }) : '')
      + para("", { line: 240 });
  }

  /* ---------- 最小 zip（不壓縮）＋ CRC32 ---------- */
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = u => { let c = 0xFFFFFFFF; for (let i = 0; i < u.length; i++) c = CRC[(c ^ u[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  function zip(files) {
    const enc = new TextEncoder(), parts = [], central = []; let off = 0;
    const d = new Date(), dt = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate(), tm = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
    files.forEach(f => {
      const name = enc.encode(f.name), data = f.data instanceof Uint8Array ? f.data : enc.encode(f.data), crc = crc32(data);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(10, tm, true); h.setUint16(12, dt, true);
      h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true); h.setUint16(26, name.length, true);
      parts.push(new Uint8Array(h.buffer), name, data);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(12, tm, true); c.setUint16(14, dt, true);
      c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true); c.setUint16(28, name.length, true); c.setUint32(42, off, true);
      central.push(new Uint8Array(c.buffer), name);
      off += 30 + name.length + data.length;
    });
    const cs = central.reduce((a, b) => a + b.length, 0), e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true); e.setUint32(12, cs, true); e.setUint32(16, off, true);
    return new Blob(parts.concat(central, [new Uint8Array(e.buffer)]), { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  }
  const X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  function docx(blocks, tables, figures) {
    figures = figures || {};
    const used = [...new Set(blocks.filter(b => b.t === "fig" && figures[b.n] && figures[b.n].png).map(b => b.n))];
    return zip([
      { name: "[Content_Types].xml", data: X + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' + (used.length ? '<Default Extension="png" ContentType="image/png"/>' : '') + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>' },
      { name: "_rels/.rels", data: X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>' },
      { name: "word/_rels/document.xml.rels", data: X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        + used.map(n => '<Relationship Id="rIdFig' + n + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/fig' + n + '.png"/>').join("") + '</Relationships>' },
      { name: "word/document.xml", data: bodyXml(blocks, tables, figures) },
      ...used.map(n => ({ name: "word/media/fig" + n + ".png", data: figures[n].png })),
      { name: "word/styles.xml", data: X + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr>' + RFONT + '<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="zh-TW"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style></w:styles>' }
    ]);
  }

  /* ---------- 介面綁定 ---------- */
  // o.figures()（選用）：回傳 {1:{cap,note,svg}, …}，svg 為含 width／height 的 SVG 字串，下載時轉成 PNG 放進報告；
  // o.tables()：回傳 {1:{cap,html,note}, 2:…}；html 用各頁「複製表格」的 Word 版 HTML（表格一律 10 pt、欄寬依內容自動計算，舊的 sz／widths 不再使用）
  function bind(o) {
    const ta = document.getElementById(o.textarea), btn = document.getElementById(o.button), msg = document.getElementById(o.msg);
    if (!ta || !btn) return;
    const toast = o.toast || (m => alert(m));
    const names = v => { if (!o.onNames || !o.nameLine) return; const nl = v.split(/\n/).map(cleanMd).find(l => o.nameLine.test(l)); if (nl) o.onNames(nl); };
    // 找不到位置標記的表格，放在第一次提到「表N」的段落之後；都沒提到就放在最後
    const place = (B, T) => {
      Object.keys(T).map(Number).sort((a, b) => a - b).forEach(n => {
        if (B.some(b => b.t === "table" && b.n === n)) return;
        let i = B.findIndex(b => b.t === "p" && new RegExp("表\\s*" + n + "(?!\\d)").test(b.x));
        B.splice(i < 0 ? B.length : i + 1, 0, { t: "table", n });
      });
      return B;
    };
    // 找不到位置標記的圖，同樣放在第一次提到「圖N」的段落之後；都沒提到就放在最後
    const placeFig = (B, F) => {
      Object.keys(F).map(Number).sort((a, b) => a - b).forEach(n => {
        if (B.some(b => b.t === "fig" && b.n === n)) return;
        let i = B.findIndex(b => b.t === "p" && new RegExp("圖\\s*" + n + "(?!\\d)").test(b.x));
        B.splice(i < 0 ? B.length : i + 1, 0, { t: "fig", n });
      });
      return B;
    };
    const figs = () => (o.figures && !(o.ready && !o.ready()) ? o.figures() : {}) || {};
    const status = () => {
      const v = ta.value; if (!v.trim()) { msg.textContent = ""; return; }
      names(v);
      const T = o.ready && !o.ready() ? {} : o.tables(), B = parse(v, o);
      const np = B.filter(b => b.t === "p").length, nh = B.filter(b => b.t === "h1" || b.t === "h2").length;
      const keys = Object.keys(T).map(Number), found = keys.filter(n => B.some(b => b.t === "table" && b.n === n)), miss = keys.filter(n => !found.includes(n));
      const noTbl = B.filter(b => b.t === "table" && !T[b.n]).map(b => "表" + b.n);
      msg.textContent = "已讀取 " + np + " 段" + (nh ? "、" + nh + " 個小標題" : "")
        + (found.length ? "；表格位置：" + found.map(n => "表" + n).join("、") : "")
        + (miss.length ? "；找不到" + miss.map(n => "「〔插入表" + n + "〕」").join("") + "，該表會放在第一次提到它的段落之後" : "")
        + (noTbl.length ? "；" + noTbl.join("、") + "目前沒有表格可放，會略過" : "")
        + figMsg(B)
        + (o.extraMsg ? o.extraMsg() : "");
    };
    function figMsg(B) {
      if (!o.figures) return "";
      const F = figs(), keys = Object.keys(F).map(Number), found = keys.filter(n => B.some(b => b.t === "fig" && b.n === n)), miss = keys.filter(n => !found.includes(n));
      const none = B.filter(b => b.t === "fig" && !F[b.n]).map(b => "圖" + b.n);
      return (found.length ? "；圖的位置：" + found.map(n => "圖" + n).join("、") : "")
        + (miss.length ? "；找不到" + miss.map(n => "「〔插入圖" + n + "〕」").join("") + "，該圖會放在第一次提到它的段落之後" : "")
        + (none.length ? "；" + none.join("、") + "目前沒有圖可放，會略過" : "");
    }
    const toPng = svg => new Promise(res => {
      const m = /width="([\d.]+)" height="([\d.]+)"/.exec(svg || ""); if (!m) { res(null); return; }
      const w = +m[1], h = +m[2], k = 3, img = new Image();
      img.onload = () => {
        const c = document.createElement("canvas"); c.width = w * k; c.height = h * k;
        const x = c.getContext("2d"); x.fillStyle = "#fff"; x.fillRect(0, 0, c.width, c.height); x.drawImage(img, 0, 0, c.width, c.height);
        c.toBlob(b => { if (!b) { res(null); return; } b.arrayBuffer().then(ab => res({ png: new Uint8Array(ab), w, h })); }, "image/png");
      };
      img.onerror = () => res(null);
      img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
    });
    ta.addEventListener("input", status);
    btn.addEventListener("click", async () => {
      if (o.ready && !o.ready()) { toast(o.notReady || "尚未完成分析"); return; }
      const v = ta.value;
      if (!v.trim()) { toast("請先貼上 AI 回覆"); ta.focus(); return; }
      names(v);
      const T = o.tables(), F0 = figs(), B = placeFig(place(parse(v, o), T), F0), F = {};
      for (const n of Object.keys(F0)) { const r = await toPng(F0[n].svg); if (r) F[n] = Object.assign({}, F0[n], r); }
      if (!B.some(b => b.t === "p")) { toast("回覆中沒有可放入報告的段落"); return; }
      const d = new Date(), ds = d.getFullYear() + String(d.getMonth() + 1).padStart(2, "0") + String(d.getDate()).padStart(2, "0");
      const a = document.createElement("a"); a.href = URL.createObjectURL(docx(B, T, F)); a.download = o.filename + "_" + ds + ".docx";
      document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 5000);
      status();
      toast("已下載 Word 報告" + (o.extraMsg ? o.extraMsg() : ""));
    });
  }
  window.DocxReport = { bind, parse, docx };
})();
