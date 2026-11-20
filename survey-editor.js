// 問卷編輯器：把貼回來的 AI 回覆解析成問卷（標題、問候語、部份、題目、選項、量表題），
// 可逐題檢視與編輯，下載 Word 問卷（同時存入資料夾）、複製純文字，並自動保存在這個瀏覽器內。
// AI 回覆中給教師檢查的附註、程式碼區塊外的說明文字另外收合顯示，不列入問卷。
// 用法：SurveyEditor.mount({ before: 元素, key: "pickany", step: 3, fileName: () => "Pick-any問卷_手搖飲料" })
(() => {
  if (window.SurveyEditor) return;

  /* ---------- 解析 ---------- */
  const MARK = /^（(標題|問候語[^）]*|靠右|置中|作答說明|問卷結尾[^）]*|李克特量表[^）]*)）\s*/;
  const PART = /^第\s*[一二三四五六七八九十\d]+\s*部[份分]/;
  const QNO = /^([SQsq]?\d+(?:\.\d+)*)(?:\s*[.．、:：)）]\s*|\s+)(\S.*)$/;
  const NOTE = /^[【\[（(]?\s*(附註|備註|格式檢查|教師檢查|檢查結果|產出前檢查|給授課教師|自我檢查|檢查說明)/;
  const clean = l => l.replace(/\t/g, " ").replace(/^\s*#{1,6}\s*/, "").replace(/^\s*>\s?/, "")
    .replace(/\*\*(.+?)\*\*/g, "$1").replace(/__(.+?)__/g, "$1").replace(/[☐❑]/g, "□")
    .replace(/^\s*[-*•]\s+(?=□)/, "").trim();

  function parse(raw) {
    let text = String(raw || "").replace(/\r/g, "");
    const notes = [];
    const fences = [...text.matchAll(/```[^\n]*\n([\s\S]*?)```/g)];
    if (fences.length) {
      const big = fences.reduce((a, b) => b[1].length > a[1].length ? b : a);
      const outside = (text.slice(0, big.index) + "\n" + text.slice(big.index + big[0].length)).trim();
      if (outside) notes.push(outside);
      text = big[1];
    }
    let lines = text.split("\n").map(clean);

    // 標題：「（標題）」標記，或前幾行中含「問卷」的短行；標題前的文字歸入附註
    let ti = lines.findIndex(l => /^（標題）/.test(l));
    if (ti < 0) {
      let seen = 0;
      for (let i = 0; i < lines.length && seen < 6; i++) {
        if (!lines[i] || /^[-—_=]{3,}$/.test(lines[i])) continue;
        seen++;
        if (lines[i].length <= 40 && /問卷|調查/.test(lines[i]) && !/^親愛/.test(lines[i]) && !/[。：:]$/.test(lines[i])) { ti = i; break; }
      }
    }
    let title = "";
    if (ti >= 0) {
      const pre = lines.slice(0, ti).filter(Boolean).join("\n");
      if (pre) notes.unshift(pre);
      title = lines[ti].replace(MARK, "").trim();
      lines = lines.slice(ti + 1);
    }

    const blocks = [];
    let q = null, para = null, lk = null, started = false, noteFrom = -1;
    const flush = () => { para = null; };
    const addText = (x, role, align) => { blocks.push({ t: "text", x, role, align: align || "left" }); };
    for (let i = 0; i < lines.length; i++) {
      let line = lines[i];
      if (!line || /^[-—_=]{3,}$/.test(line)) { flush(); continue; }
      if (started && NOTE.test(line)) { noteFrom = i; break; }
      const m = line.match(MARK);
      if (m) {
        const k = m[1], rest = line.slice(m[0].length).trim();
        flush();
        if (k === "標題") { if (!title) title = rest; continue; }
        if (/^李克特/.test(k)) { const n = (k.match(/\d+/g) || []).map(Number); lk = n.length ? Math.max(...n) : 5; q = null; continue; }
        if (k === "作答說明") { if (rest) { addText(rest, "inst"); scanScale(rest); } q = null; continue; }
        if (k === "靠右" || k === "置中") { if (rest) addText(rest, started ? "plain" : "greet", k === "靠右" ? "right" : "center"); continue; }
        if (/^問卷結尾/.test(k)) { if (rest) addText(rest, "end", "center"); q = null; continue; }
        if (!rest) continue;               // 問候語標記：下一行起為問候語
        line = rest;
      }
      if (/^\|/.test(line)) {               // Markdown 表格（李克特量表）
        flush(); q = null;
        const cells = line.replace(/^\||\|$/g, "").split("|").map(s => s.trim());
        if (cells.every(c => /^:?-{2,}:?$/.test(c) || !c)) continue;
        const nums = cells.slice(1).filter(c => /^\d+$/.test(c)).length;
        if (nums >= 3 && (cells.slice(1).every(c => /^\d+$/.test(c) || !c) || /題目|題項|敘述|項目/.test(cells[0]))) { lk = nums; continue; }
        const qm = cells[0].match(QNO);
        blocks.push({ t: "lk", no: qm ? qm[1] : "", x: qm ? qm[2] : cells[0], k: lk || Math.max(cells.length - 1, 5) });
        started = true;
        continue;
      }
      if (PART.test(line)) { flush(); q = null; lk = null; blocks.push({ t: "part", x: line }); started = true; continue; }
      if (/^□/.test(line)) {
        flush();
        const o = line.replace(/^□\s*/, "").trim();
        if (q && q.t === "q") q.opts.push(o);
        else addText(line, started ? "plain" : "greet");
        continue;
      }
      if (q && q.t === "q" && !q.opts.length && !q.scale && (line.match(/□/g) || []).length >= 3) { q.scale = line; continue; }
      const qm = line.match(QNO);
      if (qm && !(para && para.role === "greet" && !started && !/[？?）)]$/.test(line))) {
        flush();
        q = { t: "q", no: qm[1], x: qm[2], opts: [], scale: "", lk };
        blocks.push(q); started = true;
        continue;
      }
      // 一般文字：連續的行併成一段
      q = null;
      if (para) { para.x += "\n" + line; }
      else { para = { t: "text", x: line, role: started ? "plain" : "greet", align: "left" }; blocks.push(para); }
      scanScale(line);
    }
    function scanScale(s) {
      const n = [...s.matchAll(/(\d+)\s*分\s*代表/g)].map(x => +x[1]);
      if (n.length >= 2) lk = Math.max(...n);
    }
    if (noteFrom >= 0) notes.push(lines.slice(noteFrom).join("\n").trim());
    // 問候語最後的署名（…敬上）靠右；「本問卷到此結束」等結尾句置中
    for (let i = blocks.length - 1; i >= 0; i--) {
      const b = blocks[i];
      if (b.t !== "text" || b.role !== "greet" || b.align !== "left") continue;
      const ls = b.x.split("\n"), last = ls[ls.length - 1];
      if (/(敬上|謹啟|敬啟|謹上)\s*$/.test(last) && last.length <= 30) {
        if (ls.length === 1) b.align = "right";
        else { b.x = ls.slice(0, -1).join("\n"); blocks.splice(i + 1, 0, { t: "text", x: last, role: "greet", align: "right" }); }
      }
    }
    blocks.forEach(b => { if (b.t === "text" && b.role === "plain" && /^(本問卷到此結束|問卷到此結束|再次感謝|感謝您的(協助|填答|參與))/.test(b.x)) { b.role = "end"; b.align = "center"; } });
    // 沒有選項、也沒有量表列的題目：在李克特量表區內視為量表題，其餘為填答題
    const out = blocks.map(b => {
      if (b.t !== "q") return b;
      if (!b.opts.length && !b.scale && b.lk) return { t: "lk", no: b.no, x: b.x, k: b.lk };
      const r = { t: "q", no: b.no, x: b.x, opts: b.opts };
      if (b.scale) r.scale = b.scale;
      return r;
    });
    return { title, blocks: out, note: notes.filter(Boolean).join("\n\n").replace(/\n\s*\n(\s*\n)+/g, "\n\n").trim() };
  }

  /* ---------- 依序重新編號 ----------
     每個「部份」內重新起算；保留題號樣式：純數字（1、2…）、英文字首（S1、Q1…，各字首分開計數）、
     子題（1.1、1.2…：原本同一個主號的子題放在一起，主號與純數字共用同一個計數）。沒有題號的題目不編號。 */
  function renumber(doc) {
    let c = 0, pre = {}, lastMajor = null, major = 0, sub = 0, changed = 0;
    doc.blocks.forEach(b => {
      if (b.t === "part") { c = 0; pre = {}; lastMajor = null; return; }
      if ((b.t !== "q" && b.t !== "lk") || !String(b.no || "").trim()) return;
      const no = String(b.no).trim();
      let m, nw = no;
      if ((m = no.match(/^([A-Za-z]+)(\d+)$/))) { const k = m[1].toUpperCase(); pre[k] = (pre[k] || 0) + 1; nw = m[1] + pre[k]; lastMajor = null; }
      else if ((m = no.match(/^(\d+)\.(\d+)$/))) {
        if (m[1] !== lastMajor) { lastMajor = m[1]; major = ++c; sub = 0; }
        nw = major + "." + (++sub);
      }
      else if (/^\d+$/.test(no)) { nw = String(++c); lastMajor = null; }
      if (nw !== no) { b.no = nw; changed++; }
    });
    return changed;
  }

  /* ---------- 輸出：純文字 ---------- */
  const filled = a => a.filter(x => String(x).trim());
  const qLabel = no => !no ? "" : /\./.test(no) ? no + " " : no + ". ";
  function toText(doc) {
    const L = [];
    if (doc.title) L.push(doc.title, "");
    doc.blocks.forEach((b, i) => {
      const nx = doc.blocks[i + 1];
      if (b.t === "text") L.push(b.x, "");
      else if (b.t === "part") L.push(b.x, "");
      else if (b.t === "q") { L.push(qLabel(b.no) + b.x); filled(b.opts).forEach(o => L.push("□ " + o)); if (b.scale) L.push(b.scale); L.push(""); }
      else if (b.t === "lk") { L.push(qLabel(b.no) + b.x); if (!nx || nx.t !== "lk") L.push(""); }
    });
    return L.join("\n").replace(/\n{3,}/g, "\n\n").trim() + "\n";
  }

  /* ---------- 輸出：Word（.docx） ---------- */
  const RFONT = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="標楷體" w:cs="Times New Roman"/>';
  const xesc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const run = (t, o) => '<w:r><w:rPr>' + RFONT + (o.b ? '<w:b/><w:bCs/>' : '') + '<w:sz w:val="' + (o.sz || 24) + '"/><w:szCs w:val="' + (o.sz || 24) + '"/></w:rPr><w:t xml:space="preserve">' + xesc(t) + '</w:t></w:r>';
  const BOX = '<w:pBdr><w:top w:val="single" w:sz="6" w:space="4" w:color="000000"/><w:left w:val="single" w:sz="6" w:space="4" w:color="000000"/><w:bottom w:val="single" w:sz="6" w:space="4" w:color="000000"/><w:right w:val="single" w:sz="6" w:space="4" w:color="000000"/></w:pBdr>';
  function para(t, o) {
    o = o || {};
    return '<w:p><w:pPr>' + (o.keep ? '<w:keepNext/>' : '') + (o.box ? BOX : '') + (o.shd ? '<w:shd w:val="clear" w:color="auto" w:fill="EDEDED"/>' : '')
      + '<w:spacing w:before="' + (o.before || 0) + '" w:after="' + (o.after || 0) + '" w:line="276" w:lineRule="auto"/>'
      + (o.ind ? '<w:ind w:left="' + o.ind + '"/>' : '') + '<w:jc w:val="' + (o.jc || "left") + '"/></w:pPr>' + (t ? run(t, o) : '') + '</w:p>';
  }
  const jcOf = a => a === "right" ? "right" : a === "center" ? "center" : "both";
  function lkTable(items) {
    const k = Math.max(2, Math.min(11, +items[0].k || 5)), W = 9026, cw = 600, fw = W - cw * k;
    const bd = '<w:tblBorders>' + ["top", "left", "bottom", "right", "insideH", "insideV"].map(s => '<w:' + s + ' w:val="single" w:sz="4" w:space="0" w:color="000000"/>').join("") + '</w:tblBorders>';
    const cell = (t, w, o) => '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>' + para(t, Object.assign({ before: 40, after: 40 }, o)) + '</w:tc>';
    let x = '<w:tbl><w:tblPr><w:tblW w:w="' + W + '" w:type="dxa"/>' + bd + '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="' + fw + '"/>' + '<w:gridCol w:w="' + cw + '"/>'.repeat(k) + '</w:tblGrid>';
    x += '<w:tr><w:trPr><w:tblHeader/></w:trPr>' + cell("題目", fw, { jc: "center", b: true });
    for (let j = 1; j <= k; j++) x += cell(String(j), cw, { jc: "center", b: true });
    x += '</w:tr>';
    items.forEach(it => {
      x += '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + cell(qLabel(it.no) + it.x, fw, { jc: "left" });
      for (let j = 1; j <= k; j++) x += cell("□", cw, { jc: "center" });
      x += '</w:tr>';
    });
    return x + '</w:tbl>' + para("", {});
  }
  function bodyXml(doc) {
    let b = "";
    if (doc.title) b += para(doc.title, { b: true, sz: 32, jc: "center", after: 240 });
    const B = doc.blocks;
    for (let i = 0; i < B.length; i++) {
      const k = B[i];
      if (k.t === "text") {
        const ls = k.x.split("\n");
        ls.forEach((l, j) => {
          const o = { jc: jcOf(k.align) };
          if (k.role === "greet") { o.box = true; if (j === ls.length - 1 && !(B[i + 1] && B[i + 1].t === "text" && B[i + 1].role === "greet")) o.after = 240; }
          else if (k.role === "inst") { o.shd = true; if (j === ls.length - 1) o.after = 120; }
          else if (j === ls.length - 1) o.after = 120;
          if (k.role === "end") o.before = 240;
          b += para(l, o);
        });
      } else if (k.t === "part") b += para(k.x, { b: true, sz: 28, keep: true, before: 240, after: 120 });
      else if (k.t === "q") {
        const ops = filled(k.opts), last = !ops.length && !k.scale;
        b += para(qLabel(k.no) + k.x, { keep: !last, before: 120, after: last ? 120 : 0 });
        ops.forEach((o, j) => b += para("□ " + o, { ind: 480, keep: j < ops.length - 1 || !!k.scale, after: j === ops.length - 1 && !k.scale ? 120 : 0 }));
        if (k.scale) b += para(k.scale, { jc: "center", after: 120 });
      } else if (k.t === "lk") {
        const run2 = [k];
        while (B[i + 1] && B[i + 1].t === "lk" && (+B[i + 1].k || 5) === (+k.k || 5)) run2.push(B[++i]);
        b += lkTable(run2);
      }
    }
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + b
      + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>';
  }
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
  function toDocx(doc) {
    return zip([
      { name: "[Content_Types].xml", data: X + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>' },
      { name: "_rels/.rels", data: X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>' },
      { name: "word/_rels/document.xml.rels", data: X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' },
      { name: "word/document.xml", data: bodyXml(doc) },
      { name: "word/styles.xml", data: X + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr>' + RFONT + '<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="zh-TW"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style></w:styles>' }
    ]);
  }

  /* ---------- 介面 ---------- */
  const CSS = `
.sve-card{margin:20px 0 0;background:var(--surface,var(--panel,#fff));border:1px solid var(--border,var(--line,#dcdfd8));border-radius:12px;padding:22px 24px;color:var(--ink,#1f2a24);font-size:14px}
.sve-card h2{font-family:"Noto Serif TC",serif;font-size:18px;margin:0 0 6px;display:flex;align-items:center;gap:10px}
.sve-card .sve-step{display:inline-grid;place-items:center;min-width:26px;height:26px;padding:0 8px;border-radius:999px;background:var(--accent,#2f6b4f);color:var(--accent-contrast,var(--accent-ink,#fff));font-size:13px;font-weight:700}
.sve-lead{margin:0 0 10px;color:var(--ink-soft,var(--muted,#5f6b64))}
.sve-card textarea,.sve-card input,.sve-card select{font:inherit;color:inherit;background:var(--surface-2,var(--code-bg,#fbfbf8));border:1px solid var(--border,var(--line,#dcdfd8));border-radius:6px;padding:4px 8px;box-sizing:border-box}
.sve-raw{width:100%;min-height:110px;resize:vertical}
.sve-row{display:flex;flex-wrap:wrap;align-items:center;gap:8px;margin:8px 0}
.sve-card select{width:auto;flex:none}
.sve-card button{white-space:nowrap;font:inherit;font-size:13px;border:1px solid var(--accent,#2f6b4f);color:var(--accent,#2f6b4f);background:transparent;border-radius:6px;padding:3px 10px;cursor:pointer}
.sve-card button.pri{background:var(--accent,#2f6b4f);color:var(--accent-contrast,var(--accent-ink,#fff));font-weight:700;padding:7px 16px;border:none}
.sve-card button.mini{border:none;padding:1px 6px;color:var(--ink-soft,#5f6b64)}
.sve-card button.mini:hover{color:var(--accent,#2f6b4f)}
.sve-card button.warn{color:#9a3b00;border-color:#9a3b00}
.sve-msg{font-size:13px;color:var(--ink-soft,var(--muted,#5f6b64))}
.sve-tabs{display:flex;gap:6px;margin:14px 0 8px}
.sve-tabs button[aria-pressed=true]{background:var(--accent,#2f6b4f);color:var(--accent-contrast,var(--accent-ink,#fff))}
.sve-title{width:100%;font-weight:700;font-size:16px;text-align:center}
.sve-b{border:1px solid var(--border,var(--line,#dcdfd8));border-radius:8px;padding:8px 10px;margin:8px 0;background:var(--surface,#fff)}
.sve-b.part{background:var(--accent-soft,#e1ece7)}
.sve-bh{display:flex;align-items:center;gap:6px;margin-bottom:6px;font-size:12px;color:var(--ink-faint,#8f886f)}
.sve-bh .sp{flex:1}
.sve-tag{border:1px solid var(--border,#dcdfd8);border-radius:4px;padding:0 6px;white-space:nowrap}
.sve-b textarea{width:100%;resize:vertical;min-height:34px;line-height:1.6}
.sve-q{display:flex;gap:6px;align-items:flex-start}
.sve-q .no{width:64px;flex:none}
.sve-opt{display:flex;align-items:center;gap:4px;margin:4px 0 0 70px}
.sve-opt input{flex:1;min-width:0}
.sve-add{display:flex;flex-wrap:wrap;gap:4px;margin-top:6px;font-size:12px}
.sve-add button{font-size:12px;padding:0 8px;border-color:var(--border,#dcdfd8);color:var(--ink-soft,#5f6b64)}
.sve-view{border:1px solid var(--border,#dcdfd8);border-radius:8px;padding:18px 20px;background:#fff;color:#1d1d1d;font-family:"Times New Roman","DFKai-SB","BiauKai","標楷體",serif;line-height:1.7}
.sve-view h3{text-align:center;font-size:18px;margin:0 0 12px}
.sve-view .greet{border:1px solid #333;padding:8px 12px;margin-bottom:12px}
.sve-view .inst{background:#ededed;padding:4px 8px;margin:4px 0 8px}
.sve-view .part{font-weight:700;font-size:16px;margin:16px 0 6px}
.sve-view p{margin:0 0 6px;white-space:pre-wrap}
.sve-view .opt{margin:0 0 0 2em}
.sve-view table{border-collapse:collapse;width:100%;margin:6px 0 10px}
.sve-view td,.sve-view th{border:1px solid #333;padding:3px 6px}
.sve-view td.c,.sve-view th.c{text-align:center;width:2.4em}
.sve-note{margin-top:12px}
.sve-note summary{cursor:pointer;font-size:13px;color:var(--ink-soft,#5f6b64)}
.sve-note pre{white-space:pre-wrap;word-break:break-word;font-family:inherit;font-size:13px;line-height:1.7;background:var(--code-surface,var(--code-bg,#fbfbf8));border:1px solid var(--border,#dcdfd8);border-radius:8px;padding:10px 12px;margin:8px 0 0;max-height:320px;overflow:auto}
.sve-toast{position:fixed;left:50%;bottom:64px;transform:translateX(-50%);z-index:10000;max-width:90vw;padding:8px 16px;border-radius:8px;background:#28594f;color:#fff;font:14px/1.5 "Noto Sans TC","Microsoft JhengHei",sans-serif;opacity:0;transition:opacity .25s;pointer-events:none}
.sve-toast.show{opacity:1}
@media (prefers-color-scheme:dark){.sve-view{background:#f7f5ef}}`;

  const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const typeOf = b => b.t === "part" ? "部份標題" : b.t === "lk" ? "量表題" : b.t === "text" ? (b.role === "greet" ? "問候語" : b.role === "inst" ? "作答說明" : b.role === "end" ? "結尾" : "文字")
    : b.scale ? "量表題（單行）" : !b.opts.length ? "填答題" : /複選|最多選|可選/.test(b.x) ? "複選題" : "單選題";

  function mount(o) {
    if (!document.getElementById("sve-css")) { const s = document.createElement("style"); s.id = "sve-css"; s.textContent = CSS; document.head.appendChild(s); }
    const KEY = "mr.survey." + o.key;
    let st = { raw: "", doc: null };
    try { st = Object.assign(st, JSON.parse(localStorage.getItem(KEY) || "{}")); } catch (e) {}
    let mode = "edit", armed = null, saveT = null;

    const card = document.createElement("section");
    card.className = "sve-card";
    card.innerHTML = '<h2><span class="sve-step">' + (o.step || 3) + '</span>貼上 AI 回覆，檢視與編輯問卷</h2>'
      + '<p class="sve-lead">把 AI 回覆的問卷整段貼在下面（含其他說明文字也沒關係），網站會拆成標題、問候語、各題與選項，可逐題修改後下載 Word 問卷。內容會自動保存在這個瀏覽器內。</p>'
      + '<textarea class="sve-raw" placeholder="在這裡貼上 AI 回覆的完整問卷"></textarea>'
      + '<div class="sve-row"><button type="button" class="pri sve-parse">解析回覆</button><span class="sve-msg sve-pmsg"></span></div>'
      + '<div class="sve-main"></div>';
    o.before.parentNode.insertBefore(card, o.before);
    const $ = s => card.querySelector(s);
    const raw = $(".sve-raw"), main = $(".sve-main"), pmsg = $(".sve-pmsg"), pbtn = $(".sve-parse");
    raw.value = st.raw || "";

    let toastEl, toastT;
    const toast = m => {
      if (window.MRStoreToast) return window.MRStoreToast(m);
      if (!toastEl) { toastEl = document.createElement("div"); toastEl.className = "sve-toast"; document.body.appendChild(toastEl); }
      toastEl.textContent = m; toastEl.classList.add("show"); clearTimeout(toastT); toastT = setTimeout(() => toastEl.classList.remove("show"), 2200);
    };
    const save = () => { clearTimeout(saveT); saveT = setTimeout(() => { try { localStorage.setItem(KEY, JSON.stringify(st)); } catch (e) {} }, 300); };

    function summary(doc) {
      const qs = doc.blocks.filter(b => b.t === "q" || b.t === "lk");
      const c = t => qs.filter(b => typeOf(b) === t).length;
      const parts = [["單選題", c("單選題")], ["複選題", c("複選題")], ["量表題", c("量表題") + c("量表題（單行）")], ["填答題", c("填答題")]].filter(x => x[1]).map(x => x[0] + " " + x[1]);
      return "共 " + qs.length + " 題" + (parts.length ? "（" + parts.join("、") + "）" : "");
    }

    function doParse() {
      const d = parse(raw.value);
      if (!d.blocks.some(b => b.t === "q" || b.t === "lk")) { pmsg.textContent = "找不到題目：請確認貼上的是 AI 回覆的問卷內容（題目前有題號、選項前有「□」）。"; return; }
      st.doc = d; st.raw = raw.value; save();
      armed = null; pbtn.textContent = "解析回覆";
      pmsg.textContent = "已解析：" + summary(d) + "。請檢查下方內容。";
      render();
    }
    pbtn.onclick = () => {
      if (st.doc && armed !== "parse") { armed = "parse"; pbtn.textContent = "確定以新回覆取代目前的編輯內容？"; return; }
      doParse();
    };
    raw.addEventListener("input", () => { st.raw = raw.value; save(); if (armed === "parse") { armed = null; pbtn.textContent = "解析回覆"; } });
    raw.addEventListener("paste", () => setTimeout(() => { if (!st.doc) doParse(); }, 0));

    const newBlock = t => t === "q" ? { t: "q", no: "", x: "", opts: ["", ""] } : t === "lk" ? { t: "lk", no: "", x: "", k: 5 } : t === "part" ? { t: "part", x: "" } : { t: "text", x: "", role: "plain", align: "left" };
    const autoH = ta => { ta.style.height = "auto"; ta.style.height = (ta.scrollHeight + 2) + "px"; };

    function render() {
      const d = st.doc;
      if (!d) { main.innerHTML = ""; return; }
      let h = '<div class="sve-row" style="margin-top:14px"><button type="button" class="pri sve-word">下載 Word 問卷</button><button type="button" class="sve-copy">複製純文字</button><button type="button" class="sve-renum" title="各部份內依序重新編號；保留原本的題號樣式（1.、S1、Q1、1.1）">依序重新編號</button>'
        + '<span class="sve-msg">' + esc(summary(d)) + '</span><span style="flex:1"></span><button type="button" class="warn sve-clear">' + (armed === "clear" ? "確定清除貼上的回覆與編輯內容？" : "清除") + '</button>'
        + (armed === "clear" ? '<button type="button" class="sve-clear-no">取消</button>' : '') + '</div>'
        + '<div class="sve-tabs"><button type="button" data-m="edit" aria-pressed="' + (mode === "edit") + '">逐題編輯</button><button type="button" data-m="view" aria-pressed="' + (mode === "view") + '">問卷預覽</button></div>';
      if (mode === "view") h += viewHtml(d);
      else {
        h += '<input class="sve-title" data-f="title" value="' + esc(d.title) + '" placeholder="問卷標題">';
        h += '<div class="sve-add" style="margin:8px 0 0"><span class="sve-msg">在最前面新增：</span>' + addBtns(-1) + '</div>';
        d.blocks.forEach((b, i) => h += blockHtml(b, i));
      }
      if (d.note) h += '<details class="sve-note"><summary>AI 回覆中的其他內容（附註、說明；不列入問卷）</summary><pre>' + esc(d.note) + '</pre></details>';
      main.innerHTML = h;
      main.querySelectorAll("textarea").forEach(autoH);
    }
    const addBtns = i => '<button type="button" data-add="q" data-at="' + i + '">＋題目</button><button type="button" data-add="lk" data-at="' + i + '">＋量表題</button><button type="button" data-add="text" data-at="' + i + '">＋文字</button><button type="button" data-add="part" data-at="' + i + '">＋部份標題</button>';
    function blockHtml(b, i) {
      const del = armed === "b" + i;
      let h = '<div class="sve-b' + (b.t === "part" ? " part" : "") + '" data-i="' + i + '"><div class="sve-bh"><span class="sve-tag">' + typeOf(b) + '</span>';
      if (b.t === "text") h += '<select data-f="align" title="對齊"><option value="left"' + (b.align === "left" ? " selected" : "") + '>靠左</option><option value="center"' + (b.align === "center" ? " selected" : "") + '>置中</option><option value="right"' + (b.align === "right" ? " selected" : "") + '>靠右</option></select>';
      if (b.t === "lk") h += '<label>量表點數 <input type="number" min="2" max="11" data-f="k" value="' + esc(b.k || 5) + '" style="width:56px"></label>';
      h += '<span class="sp"></span><button type="button" class="mini" data-mv="-1" title="上移">↑</button><button type="button" class="mini" data-mv="1" title="下移">↓</button>'
        + '<button type="button" class="mini" data-del>' + (del ? "確定刪除？" : "刪除") + '</button>' + (del ? '<button type="button" class="mini" data-del-no>取消</button>' : '') + '</div>';
      if (b.t === "text") h += '<textarea data-f="x" rows="1">' + esc(b.x) + '</textarea>';
      else if (b.t === "part") h += '<input data-f="x" value="' + esc(b.x) + '" style="width:100%;font-weight:700">';
      else {
        h += '<div class="sve-q"><input class="no" data-f="no" value="' + esc(b.no) + '" placeholder="題號"><textarea data-f="x" rows="1" placeholder="題目">' + esc(b.x) + '</textarea></div>';
        if (b.t === "q") {
          b.opts.forEach((op, j) => h += '<div class="sve-opt"><span>□</span><input data-o="' + j + '" value="' + esc(op) + '"><button type="button" class="mini" data-omv="-1" data-j="' + j + '" title="上移">↑</button><button type="button" class="mini" data-omv="1" data-j="' + j + '" title="下移">↓</button><button type="button" class="mini" data-odel="' + j + '" title="刪除選項">✕</button></div>');
          if (b.scale != null && b.scale !== undefined && "scale" in b) h += '<div class="sve-opt"><span class="sve-msg">量表列</span><input data-f="scale" value="' + esc(b.scale) + '"></div>';
          h += '<div class="sve-opt"><button type="button" data-oadd>＋選項</button></div>';
        }
      }
      return h + '<div class="sve-add"><span class="sve-msg">在下方新增：</span>' + addBtns(i) + '</div></div>';
    }
    function viewHtml(d) {
      let h = '<div class="sve-view">' + (d.title ? '<h3>' + esc(d.title) + '</h3>' : '');
      const B = d.blocks;
      for (let i = 0; i < B.length; i++) {
        const b = B[i];
        if (b.t === "text") {
          const st2 = ' style="text-align:' + (b.align || "left") + '"';
          if (b.role === "greet") {
            let g = '<p' + st2 + '>' + esc(b.x) + '</p>';
            while (B[i + 1] && B[i + 1].t === "text" && B[i + 1].role === "greet") { const n = B[++i]; g += '<p style="text-align:' + (n.align || "left") + '">' + esc(n.x) + '</p>'; }
            h += '<div class="greet">' + g + '</div>';
          } else h += '<p class="' + (b.role === "inst" ? "inst" : "") + '"' + st2 + '>' + esc(b.x) + '</p>';
        } else if (b.t === "part") h += '<div class="part">' + esc(b.x) + '</div>';
        else if (b.t === "q") {
          h += '<p>' + esc(qLabel(b.no) + b.x) + '</p>' + filled(b.opts).map(op => '<p class="opt">□ ' + esc(op) + '</p>').join("");
          if (b.scale) h += '<p style="text-align:center">' + esc(b.scale) + '</p>';
        } else if (b.t === "lk") {
          const k = Math.max(2, Math.min(11, +b.k || 5)), items = [b];
          while (B[i + 1] && B[i + 1].t === "lk" && (+B[i + 1].k || 5) === k) items.push(B[++i]);
          h += '<table><tr><th>題目</th>' + Array.from({ length: k }, (_, j) => '<th class="c">' + (j + 1) + '</th>').join("") + '</tr>'
            + items.map(it => '<tr><td>' + esc(qLabel(it.no) + it.x) + '</td>' + '<td class="c">□</td>'.repeat(k) + '</tr>').join("") + '</table>';
        }
      }
      return h + '</div>';
    }

    // 輸入：只更新資料，不重畫（保留游標）
    main.addEventListener("input", e => {
      const t = e.target, d = st.doc; if (!d) return;
      if (t.tagName === "TEXTAREA") autoH(t);
      if (t.dataset.f === "title") { d.title = t.value; save(); return; }
      const bEl = t.closest(".sve-b"); if (!bEl) return;
      const b = d.blocks[+bEl.dataset.i];
      if (t.dataset.o != null) b.opts[+t.dataset.o] = t.value;
      else if (t.dataset.f === "k") b.k = Math.max(2, Math.min(11, parseInt(t.value, 10) || 5));
      else if (t.dataset.f) b[t.dataset.f] = t.value;
      save();
      if (t.dataset.f === "x" && b.t === "q") { const tag = bEl.querySelector(".sve-tag"); if (tag) tag.textContent = typeOf(b); }
    });
    main.addEventListener("change", e => {
      const t = e.target; if (t.dataset.f !== "align") return;
      const b = st.doc.blocks[+t.closest(".sve-b").dataset.i]; b.align = t.value; save();
    });
    main.addEventListener("click", e => {
      const t = e.target.closest("button"); if (!t) return;
      const d = st.doc;
      if (t.dataset.m) { mode = t.dataset.m; render(); return; }
      if (t.classList.contains("sve-word")) return downloadWord();
      if (t.classList.contains("sve-copy")) return copyText();
      if (t.classList.contains("sve-renum")) { const n = renumber(d); save(); render(); toast(n ? "已重新編號 " + n + " 題" : "題號都已依序，不需更動"); return; }
      if (t.classList.contains("sve-clear")) {
        if (armed !== "clear") { armed = "clear"; render(); return; }
        armed = null; st = { raw: "", doc: null }; raw.value = ""; pmsg.textContent = ""; pbtn.textContent = "解析回覆";
        try { localStorage.removeItem(KEY); } catch (er) {}
        render(); return;
      }
      if (t.classList.contains("sve-clear-no")) { armed = null; render(); return; }
      if (t.dataset.add) {
        const at = +t.dataset.at + 1;
        d.blocks.splice(at, 0, newBlock(t.dataset.add)); armed = null; save(); render();
        const nb = main.querySelector('.sve-b[data-i="' + at + '"] textarea, .sve-b[data-i="' + at + '"] input'); if (nb) nb.focus();
        return;
      }
      const bEl = t.closest(".sve-b"); if (!bEl) return;
      const i = +bEl.dataset.i, b = d.blocks[i];
      if (t.dataset.mv) { const j = i + +t.dataset.mv; if (j < 0 || j >= d.blocks.length) return; d.blocks.splice(j, 0, d.blocks.splice(i, 1)[0]); armed = null; }
      else if (t.hasAttribute("data-del")) { if (armed !== "b" + i) { armed = "b" + i; render(); return; } d.blocks.splice(i, 1); armed = null; }
      else if (t.hasAttribute("data-del-no")) armed = null;
      else if (t.dataset.omv) { const j = +t.dataset.j, k = j + +t.dataset.omv; if (k < 0 || k >= b.opts.length) return; b.opts.splice(k, 0, b.opts.splice(j, 1)[0]); }
      else if (t.dataset.odel != null) b.opts.splice(+t.dataset.odel, 1);
      else if (t.hasAttribute("data-oadd")) { b.opts.push(""); save(); render(); const ins = main.querySelectorAll('.sve-b[data-i="' + i + '"] input[data-o]'); if (ins.length) ins[ins.length - 1].focus(); return; }
      else return;
      save(); render();
    });

    function fname() { const n = String((o.fileName && o.fileName()) || "問卷").replace(/[\\/:*?"<>|]/g, "").trim() || "問卷"; return n + ".docx"; }
    function downloadWord() {
      const blob = toDocx(st.doc), a = document.createElement("a");
      a.href = URL.createObjectURL(blob); a.download = fname();
      document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    }
    async function copyText() {
      const t = toText(st.doc);
      try { await navigator.clipboard.writeText(t); toast("已複製問卷純文字"); }
      catch (e) {
        const ta = document.createElement("textarea"); ta.value = t; ta.style.position = "fixed"; ta.style.opacity = "0"; document.body.appendChild(ta); ta.select();
        try { document.execCommand("copy"); toast("已複製問卷純文字"); } catch (er) { toast("複製失敗"); }
        ta.remove();
      }
    }

    if (st.doc) pmsg.textContent = "已還原上次的編輯內容：" + summary(st.doc) + "。";
    render();
  }

  window.SurveyEditor = { mount, parse, toText, toDocx };
})();
