// AI 回覆 → Word 報告（.docx）：各工具頁共用。
// 把貼上的 AI 回覆整理成段落與標題，「〔插入表N〕」換成工具產生的表格（保留合併儲存格、粗體、底線、斜體），
// 不經任何外部程式庫，直接在瀏覽器內組成 .docx 檔下載。
// 用法：DocxReport.bind({ textarea, button, msg, title, nameLine, onNames, tables, ready, filename, toast })
(() => {
  const RFONT = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="新細明體" w:cs="Times New Roman"/>';
  const xesc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const rpr = f => '<w:rPr>' + RFONT + (f.b ? '<w:b/><w:bCs/>' : '') + (f.i ? '<w:i/><w:iCs/>' : '') + (f.u ? '<w:u w:val="single"/>' : '') + '<w:sz w:val="' + (f.sz || 24) + '"/><w:szCs w:val="' + (f.sz || 24) + '"/></w:rPr>';
  const run = (t, f) => '<w:r>' + rpr(f) + '<w:t xml:space="preserve">' + xesc(t) + '</w:t></w:r>';
  const brRun = f => '<w:r>' + rpr(f) + '<w:br/></w:r>';

  /* ---------- 解析 AI 回覆 ---------- */
  const cleanMd = t => t.replace(/^\s*#{1,6}\s*/, "").replace(/^\s*>\s?/, "").replace(/^\s*[-*•]\s+/, "")
    .replace(/\*\*(.+?)\*\*/g, "$1").replace(/__(.+?)__/g, "$1").trim();
  function parse(txt, o) {
    let lines = String(txt || "").replace(/\r/g, "").split("\n").map(cleanMd);
    const ti = lines.findIndex(l => l === o.title);
    if (ti > 0) lines = lines.slice(ti);              // 去掉報告前的建議（例如「歸屬不明題目的內容判斷」）
    lines = lines.filter(l => l && !(o.nameLine && o.nameLine.test(l)) && !/^[-*_=]{3,}$/.test(l));
    const B = lines.map((l, i) => {
      const m = l.match(/^[〔\[［【]\s*插入表\s*(\d+)\s*[〕\]］】]$/);
      if (m) return { t: "table", n: +m[1] };
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
  function table(html, opt) {
    const tb = new DOMParser().parseFromString(html, "text/html").querySelector("table");
    const sz = opt.sz || 24, rows = [...tb.rows], grid = rows.map(() => []);
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
        for (let r = 0; r < rs; r++) {
          if (!grid[ri + r]) grid[ri + r] = [];
          grid[ri + r][c] = r === 0 ? { runs, al, cs, rs, head: td.closest("thead") != null } : { cont: true, cs };
          for (let k = 1; k < cs; k++) grid[ri + r][c + k] = { skip: true };
        }
        c += cs; ncol = Math.max(ncol, c);
      });
    });
    const W = opt.widths && opt.widths.length === ncol ? opt.widths : null;
    const ln = '<w:LINE w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
    const bd = ["top", "left", "bottom", "right", "insideH", "insideV"].map(s => ln.replace("LINE", s)).join("");
    let x = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:jc w:val="center"/><w:tblBorders>' + bd + '</w:tblBorders><w:tblLayout w:type="autofit"/><w:tblCellMar><w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>'
      + Array.from({ length: ncol }, (_, i) => '<w:gridCol w:w="' + (W ? W[i] : Math.floor(9000 / ncol)) + '"/>').join("") + '</w:tblGrid>';
    grid.forEach(row => {
      x += '<w:tr><w:trPr><w:cantSplit/>' + (row.some(c => c && c.head) ? '<w:tblHeader/>' : '') + '</w:trPr>';
      for (let c = 0; c < ncol; c++) {
        const cell = row[c];
        if (cell && cell.skip) continue;
        const cs = cell && cell.cs || 1;
        let pr = W ? '<w:tcW w:w="' + W.slice(c, c + cs).reduce((a, b) => a + b, 0) + '" w:type="dxa"/>' : '<w:tcW w:w="0" w:type="auto"/>';
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
      + (o.indent ? '<w:ind w:firstLineChars="200" w:firstLine="480"/>' : '')
      + '<w:jc w:val="' + (o.jc || "both") + '"/></w:pPr>' + (text ? run(text, { b: o.b, sz: o.sz }) : '') + '</w:p>';
  }
  function bodyXml(blocks, tables) {
    let b = "";
    blocks.forEach(k => {
      if (k.t === "table") {
        const T = tables[k.n]; if (!T) return;
        b += para(T.cap, { jc: "left", keepNext: true, before: 120, after: 60, line: 240 })
          + table(T.html, T)
          + (T.note ? para(T.note, { sz: 20, line: 240, before: 40 }) : '')
          + para("", { line: 240 });
      }
      else if (k.t === "h0") b += para(k.x, { b: true, sz: 28, jc: "center", after: 240 });
      else if (k.t === "h1") b += para(k.x, { b: true, jc: "left", keepNext: true, before: 240, after: 60 });
      else if (k.t === "h2") b += para(k.x, { b: true, jc: "left", keepNext: true, before: 120 });
      else b += para(k.x, { indent: true });
    });
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + b
      + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>';
  }

  /* ---------- 最小 zip（不壓縮）＋ CRC32 ---------- */
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = u => { let c = 0xFFFFFFFF; for (let i = 0; i < u.length; i++) c = CRC[(c ^ u[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  function zip(files) {
    const enc = new TextEncoder(), parts = [], central = []; let off = 0;
    const d = new Date(), dt = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate(), tm = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
    files.forEach(f => {
      const name = enc.encode(f.name), data = enc.encode(f.data), crc = crc32(data);
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
  function docx(blocks, tables) {
    return zip([
      { name: "[Content_Types].xml", data: X + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>' },
      { name: "_rels/.rels", data: X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>' },
      { name: "word/_rels/document.xml.rels", data: X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' },
      { name: "word/document.xml", data: bodyXml(blocks, tables) },
      { name: "word/styles.xml", data: X + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr>' + RFONT + '<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="zh-TW"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style></w:styles>' }
    ]);
  }

  /* ---------- 介面綁定 ---------- */
  // o.tables()：回傳 {1:{cap,html,note,sz,widths}, 2:…}；html 用各頁「複製表格」的 Word 版 HTML
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
        + (o.extraMsg ? o.extraMsg() : "");
    };
    ta.addEventListener("input", status);
    btn.addEventListener("click", () => {
      if (o.ready && !o.ready()) { toast(o.notReady || "尚未完成分析"); return; }
      const v = ta.value;
      if (!v.trim()) { toast("請先貼上 AI 回覆"); ta.focus(); return; }
      names(v);
      const T = o.tables(), B = place(parse(v, o), T);
      if (!B.some(b => b.t === "p")) { toast("回覆中沒有可放入報告的段落"); return; }
      const d = new Date(), ds = d.getFullYear() + String(d.getMonth() + 1).padStart(2, "0") + String(d.getDate()).padStart(2, "0");
      const a = document.createElement("a"); a.href = URL.createObjectURL(docx(B, T)); a.download = o.filename + "_" + ds + ".docx";
      document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 5000);
      status();
      toast("已下載 Word 報告" + (o.extraMsg ? o.extraMsg() : ""));
    });
  }
  window.DocxReport = { bind, parse, docx };
})();
