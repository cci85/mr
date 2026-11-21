// 資料夾：工具箱各頁共用的檔案存放處（整個工具箱只用一個資料夾，在首頁或各頁頂端設定）。
// 1. Chrome／Edge 可選電腦上的資料夾（例如 Google 雲端硬碟裡的資料夾），檔案直接存在該資料夾內；
//    其他瀏覽器改存在這個瀏覽器內。
// 2. 各頁下載的檔案（Excel、Word 報告、圖、CSV、設定檔）照常下載，同時存進資料夾；
//    使用者上傳的檔案也會存一份，下次不必再上傳。尚未選資料夾時只照常下載。
// 3. 進入分析頁時，依檔名元素（clean_、sim_clean_、fa_、seg_ 等前綴與關鍵字）自動載入最適合、最新的檔案，可改選。
// 4. 各工具的主要結果表格、命名與設定、AI 回覆合併寫入「研究紀錄_研究名稱.md」（一份研究一個檔，每個工具一節）；
//    檔尾資料區塊存各頁填寫內容，載入同一份研究時自動填回並套用；知覺定位問卷產生器自動帶入市場區隔分析的結果。
// 用法：<script src="../store.js"></script>，放在頁面程式之後、newver.js 之前。
(() => {
  if (window.MRStore) return;

  /* ---------- 檔名分類（依檔名元素） ---------- */
  const KINDS = [
    ["simclean", /^sim_clean_.*\.(xlsx|csv|omv)$/i, "擴充整理資料"],
    ["cleancfg", /^clean-config_.*\.json$/i, "整理設定"],
    ["clean", /^clean_.*\.(xlsx|csv|omv)$/i, "整理後資料"],
    ["simraw", /^sim_.*\.(xlsx|csv|omv)$/i, "擴充原始回覆"],
    ["fa", /^fa_.*\.(xlsx|csv|omv)$/i, "因素分析資料"],
    ["seg", /^seg_.*\.(xlsx|csv|omv)$/i, "市場區隔資料"],
    ["result", /^(seg-cluster|seg-result|efa-result|cfa-result)_|^(前測)?樣本結構分析_|^知覺定位分析_|^(0-1資料表|欄名對照表|題組分析|交叉表|共同勾選次數表|卡方檢定結果|競爭圖)/i, "分析結果"],
    ["report", /報告.*\.docx$/i, "報告"],
    ["word", /\.docx$/i, "問卷 Word"],
    ["image", /\.(png|svg|jpe?g)$/i, "圖"],
    ["record", /^研究紀錄_.*\.md$/i, "研究紀錄"],
    ["state", /^頁面紀錄_.*\.json$/i, "舊版頁面紀錄"],
    ["config", /\.json$/i, "設定"],
    ["raw", /\.(xlsx|csv|omv)$/i, "原始回覆"],
    ["other", /./, "其他"]
  ];
  const kindOf = name => KINDS.find(k => k[1].test(name)) || KINDS[KINDS.length - 1];

  /* 各頁要自動載入的檔案：kinds 依序為可接受的類型；prefer／avoid 是檔名關鍵字（加分／減分） */
  const PAGES = {
    "organizer":       { input: "input-csv", kinds: ["raw", "simraw"] },
    "simulate":        { input: "file", kinds: ["raw"] },
    "simulate-clean":  { input: "file", kinds: ["clean"] },
    "sample":          { input: "file", kinds: ["clean", "simclean"] },
    "group-pca":       { input: "input-csv", kinds: ["clean", "simclean"], prefer: /前測|pick|市場定義|市場範圍/i, avoid: /定位/ },
    "chisq":           { input: "input-csv", kinds: ["clean", "simclean", "seg"] },
    "factor-analysis": { input: "file", kinds: ["fa", "clean", "simclean"], avoid: /前測|定位|pick/i },
    "segment":         { input: "file", kinds: ["seg", "fa"] },
    "segmentation":    { input: "file", kinds: ["seg", "clean", "simclean"], avoid: /前測|定位|pick/i },
    "positioning":     { input: "file", kinds: ["clean", "simclean", "seg"], prefer: /定位|品牌/ },
    "google-form":     { input: "file", kinds: ["word"] }
  };
  const segs = location.pathname.split("/").filter(Boolean).filter(s => !/\.html?$/i.test(s));
  const PAGE = segs.length > 1 ? segs[segs.length - 1] : "";
  const CONF = PAGES[PAGE] || null;

  /* ---------- IndexedDB ---------- */
  let dbp = null;
  const db = () => dbp || (dbp = new Promise((res, rej) => {
    const r = indexedDB.open("mr-store", 1);
    r.onupgradeneeded = () => {
      const d = r.result;
      if (!d.objectStoreNames.contains("kv")) d.createObjectStore("kv");
      if (!d.objectStoreNames.contains("files")) d.createObjectStore("files", { keyPath: "id" }).createIndex("project", "project");
    };
    r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
  }));
  const tx = async (store, mode, fn) => {
    const d = await db();
    return new Promise((res, rej) => {
      const t = d.transaction(store, mode), s = t.objectStore(store);
      const r = fn(s); t.oncomplete = () => res(r && r.result); t.onerror = () => rej(t.error);
    });
  };
  const kvGet = k => tx("kv", "readonly", s => s.get(k));
  const kvSet = (k, v) => tx("kv", "readwrite", s => s.put(v, k));
  const kvDel = k => tx("kv", "readwrite", s => s.delete(k));


  /* ---------- 資料夾：整個工具箱只用一個資料夾 ---------- */
  // 資料夾（FileSystemDirectoryHandle）存在 IndexedDB（kv「folder」）。
  // 不支援選資料夾的瀏覽器（或尚未選資料夾時上傳的檔案）改存在這個瀏覽器內。
  // 舊版「專案」設定：第一次載入時沿用原本目前專案的資料夾。
  const HAS_FS = typeof window.showDirectoryPicker === "function";
  let dir = null;           // 資料夾
  let dirOK = false;        // 是否已取得讀寫權限
  let ready = false;        // 是否可存檔（已選資料夾，或瀏覽器不支援資料夾而改存在瀏覽器內）
  let loading = true;       // 尚在讀取資料夾設定（先顯示「讀取中」，避免卡片空白）
  let listing = false, listFail = false;   // 正在列出／無法列出資料夾檔案
  const timeout = (p, ms) => Promise.race([p, new Promise((_, rej) => setTimeout(() => rej(new Error("timeout")), ms))]);

  async function loadFolder() {
    try { dir = (await kvGet("folder")) || null; } catch (e) { dir = null; }
    if (!dir) {
      try {
        const old = (await kvGet("projects2")) || [];
        const want = localStorage.getItem("mr.currentProject");
        const p = old.find(x => x.id === want && x.dir) || old.find(x => x.dir);
        if (p) { dir = p.dir; await kvSet("folder", dir); }
      } catch (e) {}
    }
    if (dir && (typeof dir.queryPermission !== "function" || dir.kind !== "directory")) dir = null;   // 舊設定若不是有效的資料夾就略過
    if (!HAS_FS) dir = null;
  }

  const IDB = {
    async list() {
      const all = await tx("files", "readonly", s => s.getAll());
      return all.map(f => ({ name: f.name, mtime: f.mtime, size: f.blob.size, key: f.id, get: async () => new File([f.blob], f.name, { type: f.blob.type, lastModified: f.mtime }) }));
    },
    async write(name, blob) {
      const old = (await IDB.list()).filter(f => f.name === name);
      for (const f of old) await tx("files", "readwrite", s => s.delete(f.key));
      await tx("files", "readwrite", s => s.put({ id: "main/" + name, project: "main", name, blob, mtime: Date.now() }));
    },
    async remove(name) { for (const f of (await IDB.list()).filter(f => f.name === name)) await tx("files", "readwrite", s => s.delete(f.key)); }
  };

  const FS = {
    async list(d) {
      const out = [];
      for await (const [n, h] of d.entries()) {
        if (h.kind !== "file" || n.startsWith(".") || n.startsWith("~$") || /^desktop\.ini$/i.test(n)) continue;
        const f = await h.getFile();
        out.push({ name: n, mtime: f.lastModified, size: f.size, get: () => h.getFile() });
      }
      return out;
    },
    async write(d, name, blob) {
      const w = await (await d.getFileHandle(name, { create: true })).createWritable();
      await w.write(blob); await w.close();
    },
    async remove(d, name) { await d.removeEntry(name); }
  };
  const useFS = () => !!(dir && dirOK);
  const listCur = () => useFS() ? FS.list(dir) : dir ? Promise.resolve([]) : IDB.list();
  const writeCur = (name, blob) => useFS() ? FS.write(dir, name, blob) : IDB.write(name, blob);
  const removeCur = name => useFS() ? FS.remove(dir, name) : IDB.remove(name);

  async function checkPermission() {
    dirOK = false;
    if (dir) { try { dirOK = (await dir.queryPermission({ mode: "readwrite" })) === "granted"; } catch (e) { dirOK = false; } }
    ready = !!dir || !HAS_FS;
  }
  async function askPermission() {          // 須在使用者點擊時呼叫
    if (!dir) return false;
    try { dirOK = (await dir.requestPermission({ mode: "readwrite" })) === "granted"; } catch (e) { dirOK = false; }
    return dirOK;
  }
  async function pickDir() {
    let h;
    try { h = await window.showDirectoryPicker({ id: "mr-project", mode: "readwrite" }); } catch (e) { return; }
    dir = h; dirOK = true; ready = true;
    await kvSet("folder", h);
    toast("已設定資料夾「" + h.name + "」");
    loadedName = null;
    await refresh();
    try { await migrateOldState(); } catch (e) {}
    autoLoad();
  }

  /* ---------- 存檔 ---------- */
  async function save(name, blob, why) {
    if (!name || !blob || !ready) return;   // 尚未選資料夾時只照常下載
    if (dir && !dirOK) await askPermission();
    try {
      if (dir && !dirOK) throw new Error("no permission");
      await writeCur(name, blob);
      toast((dir ? "已存入資料夾「" + dir.name + "」：" : "已存在這個瀏覽器內：") + name);
    } catch (e) {
      try { await IDB.write(name, blob); toast("無法寫入資料夾，已暫存在瀏覽器內：" + name, true); }
      catch (e2) { toast("存檔失敗：" + name, true); }
    }
    refresh();
  }

  // 攔截各頁的下載：記住 blob 網址對應的內容，下載時同時存進資料夾
  const blobs = new Map();
  const origCreate = URL.createObjectURL;
  URL.createObjectURL = function (o) {
    const u = origCreate.apply(URL, arguments);
    if (o instanceof Blob) { blobs.set(u, o); if (blobs.size > 200) blobs.delete(blobs.keys().next().value); }
    return u;
  };
  const origClick = HTMLAnchorElement.prototype.click;
  HTMLAnchorElement.prototype.click = function () {
    try {
      if (this.hasAttribute("download") && this.download) {
        const href = this.href, name = this.download;
        if (blobs.has(href)) save(name, blobs.get(href), "download");
        else if (/^data:/.test(href)) fetch(href).then(r => r.blob()).then(b => save(name, b, "download")).catch(() => {});
      }
    } catch (e) { /* 不影響原本的下載 */ }
    return origClick.apply(this, arguments);
  };

  // 使用者自己上傳（選檔或拖曳）的檔案也存一份；同名且大小相同的就不重存
  let autoLoading = false;
  async function keepUploads(files) {
    if (!files || !files.length) return;
    let have = [];
    if (!ready) return;
    try { have = await listCur(); } catch (e) {}
    for (const f of files) {
      if (have.some(h => h.name === f.name && h.size === f.size)) continue;
      await save(f.name, f, "upload");
    }
  }
  document.addEventListener("change", e => {
    const t = e.target;
    if (autoLoading || !t || t.type !== "file" || !t.files) return;
    keepUploads([...t.files]);
  }, true);
  document.addEventListener("drop", e => {
    if (autoLoading || !e.dataTransfer || !e.dataTransfer.files || !e.dataTransfer.files.length) return;
    if (e.target && e.target.closest && e.target.closest(".mrs")) return;
    keepUploads([...e.dataTransfer.files]);
  }, true);

  /* ---------- 介面 ---------- */
  const css = document.createElement("style");
  css.textContent = `
.mrs{--a:#28594f;--b:#ddd7c9;--bg:#fbfaf6;--ink:#232019;--soft:#6b6457;font:14px/1.6 "Noto Sans TC","PingFang TC","Microsoft JhengHei",system-ui,sans-serif;color:var(--ink)}
@media (prefers-color-scheme:dark){.mrs{--a:#79bda7;--b:#373d33;--bg:#1c1f1a;--ink:#eae7dc;--soft:#a9a493}}
.mrs-bar{display:flex;flex-wrap:wrap;align-items:center;gap:6px 10px;padding:8px 16px;background:var(--bg);border-bottom:1px solid var(--b)}
.mrs-bar b{font-weight:600}
.mrs select,.mrs input[type=text]{font:inherit;color:inherit;background:transparent;border:1px solid var(--b);border-radius:6px;padding:2px 6px;max-width:260px}
.mrs button{font:inherit;font-size:13px;color:var(--a);background:transparent;border:1px solid var(--a);border-radius:6px;padding:2px 10px;cursor:pointer}
.mrs button.pri{background:var(--a);color:#fff}
@media (prefers-color-scheme:dark){.mrs button.pri{color:#141612}}
.mrs button:disabled{opacity:.45;cursor:default}
.mrs .soft{color:var(--soft);font-size:13px}
.mrs .warn{color:#9a5b00}
.mrs-panel{padding:10px 16px 14px;background:var(--bg);border-bottom:1px solid var(--b)}
.mrs-panel[hidden]{display:none}
.mrs-panel table{border-collapse:collapse;width:100%;max-width:980px;font-size:13.5px}
.mrs-panel td,.mrs-panel th{padding:4px 8px;border-bottom:1px solid var(--b);text-align:left;vertical-align:middle}
.mrs-panel th{font-weight:600;color:var(--soft)}
.mrs-panel td.n{word-break:break-all}
.mrs-tag{display:inline-block;padding:0 6px;border-radius:4px;border:1px solid var(--b);font-size:12px;color:var(--soft);white-space:nowrap}
.mrs-auto{margin:8px 0 10px;padding:8px 12px;border:1px dashed var(--a);border-radius:8px;background:var(--bg)}
.mrs-auto select{max-width:100%}
.mrs-home{margin:-8px 0 28px;border:1px solid var(--b);border-radius:12px;overflow:hidden;background:var(--bg)}
.mrs-home .mrs-bar{padding:12px 16px;border-bottom:0}
.mrs-home .mrs-title{font-size:16px;margin-right:4px}
.mrs-home .mrs-note{flex-basis:100%;margin:2px 0 0}
.mrs-home .mrs-panel{border-top:1px solid var(--b);border-bottom:0}
.mrs-toast{position:fixed;left:50%;bottom:22px;transform:translateX(-50%);z-index:10000;max-width:90vw;padding:8px 16px;border-radius:8px;background:#28594f;color:#fff;font:14px/1.5 "Noto Sans TC","Microsoft JhengHei",sans-serif;opacity:0;transition:opacity .25s;pointer-events:none}
.mrs-toast.err{background:#9a5b00}
.mrs-toast.show{opacity:1}`;
  document.head.appendChild(css);

  const el = (tag, attrs, html) => { const e = document.createElement(tag); Object.assign(e, attrs || {}); if (html != null) e.innerHTML = html; return e; };
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const p2 = n => String(n).padStart(2, "0");
  const when = t => { const d = new Date(t); return (d.getMonth() + 1) + "/" + d.getDate() + " " + p2(d.getHours()) + ":" + p2(d.getMinutes()); };
  const kb = n => n >= 1048576 ? (n / 1048576).toFixed(1) + " MB" : Math.max(1, Math.round(n / 1024)) + " KB";

  let toastEl, toastT;
  function toast(msg, err) {
    if (!toastEl) { toastEl = el("div", { className: "mrs-toast" }); document.body.appendChild(toastEl); }
    toastEl.textContent = msg; toastEl.classList.toggle("err", !!err); toastEl.classList.add("show");
    clearTimeout(toastT); toastT = setTimeout(() => toastEl.classList.remove("show"), 2600);
  }

  const wrap = el("div", { className: "mrs" });
  const bar = el("div", { className: "mrs-bar" });
  const panel = el("div", { className: "mrs-panel", hidden: true });
  wrap.append(bar, panel);

  const HOME = !PAGE;
  function mountBar() {
    if (HOME) {                             // 首頁：資料夾卡片放在說明文字下方
      wrap.classList.add("mrs-home");
      const lede = document.querySelector("main .lede");
      if (lede) { lede.after(wrap); return; }
    }
    const nav = document.querySelector("nav.toolbox-back:not(.bottom)");
    if (nav) nav.after(wrap); else document.body.insertBefore(wrap, document.body.firstChild);
  }

  let files = [];

  async function renderBar() {
    window.dispatchEvent(new CustomEvent("mr-research", { detail: { research: typeof research === "string" ? research : "" } }));
    let html = HOME ? "<b class=\"mrs-title\">資料夾</b>" : "<b>資料夾</b> ";
    if (loading) {
      bar.innerHTML = html + "<span class=\"soft\">讀取中…</span>";
      return;
    }
    if (!HAS_FS) {
      html += "<span>檔案存在這個瀏覽器內</span> <span class=\"soft\">（此瀏覽器不支援選擇電腦資料夾，請用 Chrome 或 Edge）</span>";
    } else if (!dir) {
      html += "<span class=\"soft\">尚未設定。</span> <button type=\"button\" class=\"mrs-pick pri\">選擇資料夾</button>";
    } else if (!dirOK) {
      html += "<span class=\"warn\">「" + esc(dir.name) + "」尚未連線</span> <button type=\"button\" class=\"mrs-grant pri\">連線資料夾</button>" +
        " <button type=\"button\" class=\"mrs-pick\">更換資料夾</button>";
    } else {
      html += "「" + esc(dir.name) + "」 <button type=\"button\" class=\"mrs-pick\" title=\"改用另一個資料夾；原資料夾裡的檔案不會移動或刪除\">更換資料夾</button>";
    }
    if (ready && (!dir || dirOK)) html += " <span class=\"soft\">｜</span> <button type=\"button\" class=\"mrs-files\">資料夾檔案（" + (listing ? "讀取中…" : listFail ? "無法讀取" : files.length) + "）" + (panel.hidden ? "▾" : "▴") + "</button>";
    if (HOME) {
      html += "<p class=\"soft mrs-note\">" + (!HAS_FS ? "各工具上傳與下載的檔案會存在這個瀏覽器內，分析頁會自動帶入；清除瀏覽資料或換電腦就會消失。"
        : !dir ? "選一個電腦上的資料夾（例如 Google 雲端硬碟裡的資料夾）。之後各工具上傳與下載的檔案都會自動存進去，分析頁會依檔名自動帶入，不必再手動上傳。"
        : !dirOK ? "瀏覽器重新開啟後，需要按一次「連線資料夾」，才能讀寫資料夾。"
        : "各工具上傳與下載的檔案會自動存進這個資料夾，分析頁會依檔名自動帶入最新、最適合的檔案；自行放進資料夾的檔案也讀得到。") + "</p>";
    }
    html += researchHtml();
    bar.innerHTML = html;
    bindResearch();
    const q = s => bar.querySelector(s);
    if (q(".mrs-grant")) q(".mrs-grant").onclick = async () => { if (await askPermission()) { await refresh(); try { await migrateOldState(); } catch (e) {} autoLoad(); } };
    if (q(".mrs-pick")) q(".mrs-pick").onclick = pickDir;
    if (q(".mrs-files")) q(".mrs-files").onclick = () => { panel.hidden = !panel.hidden; renderBar(); renderPanel(); };
  }

  let armed = null;   // 刪除前需再按一次
  async function renderPanel() {
    if (panel.hidden || !ready) return;
    let html = "";
    if (useFS()) {
      const n = (await IDB.list().catch(() => [])).length;
      if (n) html += "<p class=\"warn\">還有 " + n + " 個檔案暫存在瀏覽器內。<button type=\"button\" class=\"mrs-move\">搬到資料夾「" + esc(dir.name) + "」</button></p>";
    }
    if (!files.length) {
      html += "<p class=\"soft\">目前沒有檔案。在各工具下載的檔案與上傳的檔案，都會自動存到這裡。</p>";
    } else {
      html += "<table><thead><tr><th>檔名</th><th>類型</th><th>時間</th><th>大小</th><th></th></tr></thead><tbody>" +
        files.map((f, i) => "<tr><td class=\"n\">" + esc(f.name) + "</td><td><span class=\"mrs-tag\">" + kindOf(f.name)[2] + "</span></td><td>" + when(f.mtime) + "</td><td>" + kb(f.size) +
          "</td><td style=\"white-space:nowrap\"><button type=\"button\" data-dl=\"" + i + "\">下載</button> <button type=\"button\" data-rm=\"" + i + "\">" + (armed === f.name ? "確定刪除？" : "刪除") + "</button>" + (armed === f.name ? " <button type=\"button\" data-cancel>取消</button>" : "") + "</td></tr>").join("") +
        "</tbody></table>";
      if (!dir) html += "<p class=\"soft\">檔案存在這個瀏覽器內；清除瀏覽資料或換電腦就會消失。<button type=\"button\" class=\"mrs-zip\">全部下載（zip）</button></p>";
    }
    panel.innerHTML = html;
    panel.querySelectorAll("[data-dl]").forEach(b => b.onclick = async () => {
      const f = await files[+b.dataset.dl].get();
      const a = document.createElement("a"); a.href = origCreate.call(URL, f); a.download = f.name;
      document.body.appendChild(a); origClick.call(a); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 3000);
    });
    panel.querySelectorAll("[data-rm]").forEach(b => b.onclick = async () => {
      const f = files[+b.dataset.rm];
      if (armed !== f.name) { armed = f.name; renderPanel(); return; }
      armed = null;
      try { await removeCur(f.name); toast("已刪除：" + f.name); } catch (e) { toast("刪除失敗：" + f.name, true); }
      refresh();
    });
    panel.querySelectorAll("[data-cancel]").forEach(b => b.onclick = () => { armed = null; renderPanel(); });
    const mv = panel.querySelector(".mrs-move");
    if (mv) mv.onclick = moveToFolder;
    const zp = panel.querySelector(".mrs-zip");
    if (zp) zp.onclick = zipAll;
  }

  async function moveToFolder() {
    let n = 0;
    for (const f of await IDB.list()) {
      try { await FS.write(dir, f.name, await f.get()); await tx("files", "readwrite", s => s.delete(f.key)); n++; } catch (e) {}
    }
    toast("已搬移 " + n + " 個檔案到「" + dir.name + "」");
    refresh();
  }

  function loadScript(src) { return new Promise((res, rej) => { const s = document.createElement("script"); s.src = src; s.onload = res; s.onerror = rej; document.head.appendChild(s); }); }
  async function zipAll() {
    try {
      if (typeof JSZip === "undefined") await loadScript("https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js");
      const z = new JSZip();
      for (const f of files) z.file(f.name, await f.get());
      const blob = await z.generateAsync({ type: "blob" });
      const a = document.createElement("a"); a.href = origCreate.call(URL, blob); a.download = "行銷研究工具箱檔案.zip";
      document.body.appendChild(a); origClick.call(a); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 3000);
    } catch (e) { toast("壓縮失敗，請改為逐一下載", true); }
  }

  async function refresh() {
    listing = true; listFail = false; await renderBar();
    try { files = ready ? await timeout(listCur(), 20000) : []; } catch (e) { files = []; listFail = true; }
    listing = false;
    files.sort((a, b) => b.mtime - a.mtime);
    recNames = files.filter(f => /^研究紀錄_.+\.md$/.test(f.name)).map(f => f.name.slice(5, -3));
    if (GEN && !research && !genTried && ready && (!dir || dirOK)) {
      genTried = true;
      const cur = localStorage.getItem(LS_R);
      if (cur && recNames.includes(cur)) setResearch(cur);
    }
    await renderBar(); await renderPanel(); renderAuto();
  }

  /* ---------- 自動載入 ---------- */
  let autoBox = null, loadedName = null;
  function candidates() {
    if (!CONF) return [];
    const score = f => {
      const k = CONF.kinds.indexOf(kindOf(f.name)[0]);
      if (k < 0) return null;
      return (CONF.prefer && CONF.prefer.test(f.name) ? 2 : 0) - (CONF.avoid && CONF.avoid.test(f.name) ? 1 : 0);
    };
    return files.map(f => ({ f, s: score(f) })).filter(x => x.s !== null)
      .sort((a, b) => b.s - a.s || b.f.mtime - a.f.mtime).map(x => x.f);
  }
  function renderAuto() {
    if (!CONF) return;
    const input = document.getElementById(CONF.input);
    if (!input) return;
    if (!autoBox) {
      autoBox = el("div", { className: "mrs mrs-auto" });
      const anchor = input.closest(".dropzone") || document.getElementById("dropzone") || input;
      anchor.parentNode.insertBefore(autoBox, anchor);
    }
    const c = candidates();
    if (!ready) { autoBox.innerHTML = "<span class=\"soft\">按上方「選擇資料夾」設定資料夾後，上傳與下載的檔案會自動存進資料夾，下次直接帶入。</span>"; return; }
    if (dir && !dirOK) { autoBox.innerHTML = "<span class=\"warn\">資料夾尚未連線，請按上方「連線資料夾」，即可自動帶入資料夾中的檔案。</span>"; return; }
    if (!c.length) {
      autoBox.innerHTML = "<span class=\"soft\">資料夾中還沒有適合本頁的檔案，請在下方上傳；上傳的檔案會自動存進資料夾，下次直接帶入。</span>";
      return;
    }
    autoBox.innerHTML = (loadedName ? "已從資料夾帶入：" : "資料夾中的檔案：") +
      " <select>" + c.map((f, i) => "<option value=\"" + i + "\"" + (f.name === loadedName ? " selected" : "") + ">" + esc(f.name) + "（" + kindOf(f.name)[2] + "，" + when(f.mtime) + "）</option>").join("") + "</select>" +
      " <button type=\"button\" class=\"pri\">" + (loadedName ? "改用這個檔案" : "載入") + "</button>";
    const sel = autoBox.querySelector("select");
    autoBox.querySelector("button").onclick = () => loadInto(c[+sel.value]);
  }
  async function loadInto(f) {
    const input = CONF && document.getElementById(CONF.input);
    if (!input || !f) return;
    let file;
    try { file = await f.get(); } catch (e) { toast("讀取失敗：" + f.name, true); return; }
    try {
      const dt = new DataTransfer(); dt.items.add(file);
      autoLoading = true;
      input.files = dt.files;
      input.dispatchEvent(new Event("change", { bubbles: true }));
    } catch (e) { toast("此瀏覽器無法自動帶入，請手動上傳", true); return; }
    finally { autoLoading = false; }
    loadedName = f.name;
    renderAuto();
  }
  function autoLoad() {
    if (!CONF) return;
    const c = candidates();
    if (!c.length) return;
    loadInto(c[0]);
  }


  /* ---------- 研究紀錄：各工具的結果、命名與設定、AI 回覆合併成一份 .md ---------- */
  // 一份研究一個「研究紀錄_研究名稱.md」：每個工具一節（重新執行只更新自己那一節），
  // 檔尾的資料區塊（<!-- mr-data … -->）存各頁的填寫內容，載入同一份研究時自動填回並套用。
  // 分析頁依資料檔名對應研究（clean_、fa_、seg_ 等前綴去掉後相同者為同一份研究）；
  // 沒有資料檔的問卷產生器在資料夾列選研究。知覺定位問卷產生器自動帶入市場區隔分析的結果。
  const GEN = ["pickany", "segment-survey", "positioning-survey"].includes(PAGE);
  const REC_ON = !!CONF || GEN;
  const ORDER = ["pickany", "organizer", "simulate", "simulate-clean", "group-pca", "chisq", "sample", "segment-survey",
    "factor-analysis", "segment", "segmentation", "positioning-survey", "positioning", "google-form"];
  const stemOf = name => {
    let b = String(name || "").replace(/\.[^.]+$/, ""), prev;
    do { prev = b; b = b.replace(/^(?:sim_clean|clean|sim|fa|seg|efa-result|cfa-result|seg-result|seg-cluster)_/i, ""); } while (b !== prev);
    return b.replace(/_(因素分析資料|因素分數資料|集群資料|市場區隔資料|市場區隔分析結果|EFA分析結果|CFA分析結果)$/, "").replace(/[\\/:*?"<>|]/g, "_").trim();
  };
  const recName = r => "研究紀錄_" + r + ".md";
  const LS_R = "mr.research";
  let research = "", dataFile = "", stem = "", stFields = {}, stTimer = null, mutTimer = null, restoring = false, stGen = 0, lastBody = "", attachedNote = "";

  /* 讀寫研究紀錄 */
  const DATA_RE = /\n?<!-- mr-data\n([\s\S]*?)\n-->\s*$/;
  async function listRecords() {
    let L = [];
    try { L = await listCur(); } catch (e) {}
    return L.filter(f => /^研究紀錄_.+\.md$/.test(f.name)).sort((a, b) => b.mtime - a.mtime);
  }
  function parseRecord(text) {
    let data = { research: "", files: [], pages: {} };
    const m = DATA_RE.exec(text || "");
    if (m) { try { data = Object.assign(data, JSON.parse(m[1])); } catch (e) {} }
    const body = m ? text.slice(0, m.index) : (text || "");
    const secs = {};
    body.replace(/<!-- mr-section:([\w-]+) -->\n([\s\S]*?)<!-- \/mr-section:\1 -->/g, (all, p) => { secs[p] = all; return ""; });
    return { data, secs };
  }
  async function readRecord(r) {
    const f = (await listRecords()).find(x => x.name === recName(r));
    if (!f) return null;
    try { const text = await (await f.get()).text(); return Object.assign(parseRecord(text), { text, mtime: f.mtime }); } catch (e) { return null; }
  }
  function composeRecord(r, rec) {
    const keys = Object.keys(rec.secs).sort((a, b) => (ORDER.indexOf(a) + 1 || 99) - (ORDER.indexOf(b) + 1 || 99));
    return "# 研究紀錄：" + r + "\n\n最後更新：" + new Date().toLocaleString("zh-TW", { hour12: false }) +
      "\n\n本檔由行銷研究工具箱自動產生：每個工具一節，重新執行時只更新該節。檔尾的資料區塊供工具箱自動帶回填寫內容，請勿修改。\n\n" +
      keys.map(k => rec.secs[k].trim()).join("\n\n") + "\n\n<!-- mr-data\n" + JSON.stringify(rec.data) + "\n-->\n";
  }

  /* 依資料檔決定研究 */
  async function resolveResearch(st) {
    const recs = await listRecords();
    for (const f of recs) {
      try { const { data } = parseRecord(await (await f.get()).text()); if ((data.files || []).includes(st)) return { r: f.name.slice(5, -3), how: "map" }; } catch (e) {}
    }
    if (recs.some(f => f.name === recName(st))) return { r: st, how: "same" };
    const cur = localStorage.getItem(LS_R);
    if (cur && recs.some(f => f.name === recName(cur))) return { r: cur, how: "attach" };
    return { r: st, how: "new" };
  }

  /* 表單欄位 */
  const TEXTY = /^(text|search|number|email|url|tel|date|month|week|time)$/;
  function stEligible(e) {
    if (!e || !e.tagName || e.closest(".mrs") || e.hasAttribute("data-nosave") || (e.disabled && e.type !== "checkbox")) return false;
    if (e.tagName === "TEXTAREA") return !e.readOnly;
    if (e.tagName === "SELECT") return true;
    if (e.tagName === "INPUT") return !e.readOnly && (TEXTY.test(e.type) || e.type === "checkbox" || e.type === "radio");
    return false;
  }
  function stKeyOf(e) {
    if (e.type === "radio" && e.name) return "r:" + e.name;
    if (e.id) return "#" + e.id;
    const ds = [...e.attributes].filter(a => /^data-/.test(a.name));
    if (!e.name && !ds.length) return null;
    const sel = e.tagName.toLowerCase() + (e.name ? "[name=\"" + CSS.escape(e.name) + "\"]" : "") + ds.map(a => "[" + a.name + "=\"" + CSS.escape(a.value) + "\"]").join("");
    return "s:" + sel + "|" + [...document.querySelectorAll(sel)].indexOf(e);
  }
  function stFind(k) {
    if (k.startsWith("#")) return [document.getElementById(k.slice(1))].filter(Boolean);
    if (k.startsWith("r:")) return [...document.getElementsByName(k.slice(2))].filter(e => e.type === "radio");
    const i = k.lastIndexOf("|"), sel = k.slice(2, i), n = +k.slice(i + 1);
    try { const e = document.querySelectorAll(sel)[n]; return e ? [e] : []; } catch (err) { return []; }
  }
  const fieldEls = () => [...document.querySelectorAll("textarea,select,input")].filter(stEligible);
  function stSnapshot() {
    const out = Object.assign({}, stFields);
    fieldEls().forEach(e => {
      const k = stKeyOf(e); if (!k) return;
      if (e.type === "radio") { if (e.checked) out[k] = e.value; return; }
      out[k] = e.type === "checkbox" ? e.checked : e.value;
    });
    return out;
  }

  /* 本頁那一節的內容 */
  const isReply = e => e.tagName === "TEXTAREA" && (/reply|paste|import/i.test(e.id) || /AI/.test(e.placeholder || ""));
  const clean = s => String(s || "").replace(/\s+/g, " ").trim();
  function labelOf(e) {
    let t = "";
    if (e.id) { const l = document.querySelector("label[for=\"" + CSS.escape(e.id) + "\"]"); if (l) t = l.textContent; }
    if (!t) { const l = e.closest("label"); if (l) t = l.textContent; }
    if (!t && e.tagName === "TEXTAREA") t = capOf(e);
    if (!t && e.type === "radio") { const fs = e.closest("fieldset"); const lg = fs && fs.querySelector("legend"); t = lg ? lg.textContent : e.name; }
    t = clean(t || e.getAttribute("aria-label") || e.placeholder || e.name || e.id);
    return t.length > 40 ? t.slice(0, 40) + "…" : t;
  }
  function capOf(t) {
    let n = t;
    for (let up = 0; n && n !== document.body && up < 6; up++, n = n.parentElement) {
      let s = n.previousElementSibling;
      for (let i = 0; s && i < 4; i++, s = s.previousElementSibling) {
        if (/^H[1-5]$/.test(s.tagName) || s.classList.contains("fa-cap") || s.classList.contains("cap")) return clean(s.textContent);
        const h = s.querySelector && s.querySelector("h1,h2,h3,h4,h5,.fa-cap");
        if (h) return clean(h.textContent);
      }
    }
    return "";
  }
  const cellTxt = c => clean(c.innerText || c.textContent).replace(/\|/g, "｜");
  function tableMd(t) {
    const rows = [...t.rows], grid = rows.map(() => []);
    let w = 0;
    rows.forEach((tr, ri) => {
      let c = 0;
      [...tr.cells].forEach(td => {
        while (grid[ri][c] != null) c++;
        const rs = td.rowSpan || 1, cs = td.colSpan || 1, tx = cellTxt(td);
        for (let a = 0; a < rs; a++) for (let b = 0; b < cs; b++) if (grid[ri + a]) grid[ri + a][c + b] = a === 0 && b === 0 ? tx : "";
        c += cs; w = Math.max(w, c);
      });
    });
    if (rows.length < 2 || !w) return "";
    const line = r => "| " + Array.from({ length: w }, (_, i) => r[i] == null ? "" : r[i]).join(" | ") + " |";
    const MAX = 80, body = grid.slice(1);
    return [line(grid[0]), "|" + " --- |".repeat(w)].concat(body.slice(0, MAX).map(line)).join("\n") +
      (body.length > MAX ? "\n\n（共 " + body.length + " 列，僅列出前 " + MAX + " 列）" : "");
  }
  const fence = s => "~~~~text\n" + String(s).replace(/~~~~/g, "～～～～").trim() + "\n~~~~";
  function buildSection() {
    const title = clean((document.querySelector("h1") || {}).textContent || document.title);
    const L = ["<!-- mr-section:" + PAGE + " -->", "## " + title, "", "__TIME__" + (dataFile ? "｜資料檔：" + dataFile : "")];
    const els = fieldEls(), set = [], rep = [];
    els.forEach(e => {
      if (isReply(e)) { if (e.value.trim()) rep.push(e); return; }
      if (e.type === "radio") { if (!e.checked) return; }
      let v = e.type === "checkbox" ? (e.checked ? "是" : "否") : e.tagName === "SELECT" ? clean((e.selectedOptions[0] || {}).textContent || e.value) : e.type === "radio" ? clean((e.closest("label") || {}).textContent || e.value) : e.value;
      if (v == null || String(v).trim() === "") return;
      v = String(v).trim();
      const lab = e.type === "radio" ? clean((e.closest("fieldset") && e.closest("fieldset").querySelector("legend") || {}).textContent || "選項") : labelOf(e);
      set.push(/\n/.test(v) ? "- " + lab + "：\n" + v.split("\n").map(x => "  " + x).join("\n") : "- " + lab + "：" + v);
    });
    if (set.length) L.push("", "### 命名與設定", "", ...set);
    const tabs = [...document.querySelectorAll("table")].filter(t => {
      if (t.closest(".mrs") || t.querySelector("input,select,textarea,button")) return false;
      const h = t.closest("[hidden]"); if (h && !/^tab-/.test(h.id || "")) return false;
      return t.rows.length > 1;
    });
    const tmd = tabs.map(t => { const md = tableMd(t); if (!md) return ""; const c = capOf(t); return (c ? "#### " + c + "\n\n" : "") + md; }).filter(Boolean);
    if (tmd.length) L.push("", "### 主要結果", "", tmd.join("\n\n"));
    rep.forEach(e => L.push("", "### AI 回覆：" + labelOf(e), "", fence(e.value)));
    let extra = [];
    try { extra = typeof window.MRRecordExtra === "function" ? window.MRRecordExtra() || [] : []; } catch (e) { extra = []; }
    extra.forEach(x => { if (x && x.text) L.push("", "### " + x.title, "", fence(x.text)); });
    L.push("<!-- /mr-section:" + PAGE + " -->");
    return L.join("\n");
  }
  function sectionText(rec, page, title) {
    const s = rec && rec.secs[page]; if (!s) return "";
    const re = new RegExp("### " + title.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\n\\n~~~~text\\n([\\s\\S]*?)\\n~~~~");
    const m = re.exec(s); return m ? m[1] : "";
  }

  /* 寫入 */
  async function recWrite(force) {
    clearTimeout(stTimer); stTimer = null; clearTimeout(mutTimer); mutTimer = null;
    if (!research || !ready || restoring || (dir && !dirOK)) return;
    const body = buildSection();
    if (!force && body === lastBody) return;
    stFields = stSnapshot();
    const rec = (await readRecord(research)) || { data: { research, files: [], pages: {} }, secs: {} };
    rec.data.research = research;
    if (stem && !(rec.data.files || []).includes(stem)) rec.data.files = (rec.data.files || []).concat(stem);
    rec.data.pages = rec.data.pages || {};
    const now = new Date();
    rec.data.pages[PAGE] = { saved: now.toISOString(), file: dataFile, fields: stFields };
    rec.secs[PAGE] = body.replace("__TIME__", "更新：" + now.toLocaleString("zh-TW", { hour12: false }));
    const blob = new Blob([composeRecord(research, rec)], { type: "text/markdown" });
    try { await writeCur(recName(research), blob); } catch (e) { try { await IDB.write(recName(research), blob); } catch (e2) { return; } }
    lastBody = body;
    if (!files.some(f => f.name === recName(research))) refresh();
  }
  const stSchedule = () => { if (!research || restoring) return; clearTimeout(stTimer); stTimer = setTimeout(() => recWrite(), 1500); };

  /* 還原 */
  const settle = (quiet, max) => new Promise(res => {
    let t, done = false;
    const end = () => { if (done) return; done = true; mo.disconnect(); clearTimeout(t); clearTimeout(cap); res(); };
    const mo = new MutationObserver(() => { clearTimeout(t); t = setTimeout(end, quiet); });
    mo.observe(document.body, { childList: true, subtree: true, attributes: true, characterData: true });
    t = setTimeout(end, quiet); const cap = setTimeout(end, max);
  });
  function stApply(k, v) {
    const els = stFind(k); if (!els.length) return false;
    let e = els[0];
    if (k.startsWith("r:")) { e = els.find(x => x.value === v); if (!e || e.checked) return false; e.checked = true; }
    else if (e.type === "checkbox") { if (e.checked === !!v) return false; e.checked = !!v; }
    else { if (e.value === v) return false; e.value = v; if (e.tagName === "SELECT" && e.value !== v) return false; }
    e.dispatchEvent(new Event("input", { bubbles: true }));
    e.dispatchEvent(new Event("change", { bubbles: true }));
    return true;
  }
  async function applyFields(fields, gen) {
    const tries = {}, toggle = k => { if (k.startsWith("r:")) return true; const e = stFind(k)[0]; return !!e && (e.tagName === "SELECT" || e.type === "checkbox"); };
    const t0 = Date.now(); let idle = 0, n = 0;
    while (Date.now() - t0 < 20000 && gen === stGen) {
      const keys = Object.keys(fields).filter(k => (tries[k] || 0) < 3);
      let changed = false;
      for (const k of keys.filter(toggle)) if (stApply(k, fields[k])) { tries[k] = (tries[k] || 0) + 1; changed = true; n++; }
      if (!changed) for (const k of keys.filter(k => !toggle(k))) if (stApply(k, fields[k])) { tries[k] = (tries[k] || 0) + 1; changed = true; n++; }
      if (changed) { idle = 0; await settle(400, 5000); continue; }
      if (++idle >= 3) break;
      await settle(500, 2000);
    }
    return n;
  }
  async function recRestore(gen) {
    if (!ready || (dir && !dirOK) || !research) return;
    const rec = await readRecord(research);
    if (gen !== stGen) return;
    let pd = rec && rec.data.pages && rec.data.pages[PAGE];
    restoring = gen;
    let n = 0, imported = false;
    try {
      await settle(700, 8000);
      if (pd && pd.fields) { n = await applyFields(pd.fields, gen); if (gen === stGen) stFields = Object.assign({}, pd.fields); }
      if (PAGE === "positioning-survey" && gen === stGen) imported = rec ? await posImport(rec, pd) : (posEvt({ research, found: false }), false);
    } finally { if (restoring === gen) restoring = false; }
    if (gen !== stGen) return;
    if (n && !imported) toast("已從研究紀錄帶回上次的填寫內容與 AI 回覆");
    lastBody = ""; stSchedule();
  }
  /* 知覺定位問卷產生器：帶入市場區隔分析的結果 */
  const posEvt = d => window.dispatchEvent(new CustomEvent("mr-pos", { detail: d }));
  async function posImport(rec, pd, force) {
    const txt = sectionText(rec, "segment", "知覺定位問卷資料") || sectionText(rec, "segmentation", "知覺定位問卷資料");
    const ta = document.getElementById("import");
    if (!ta) return false;
    if (!txt) { posEvt({ research, found: false }); return false; }
    const segData = rec.data.pages.segment || rec.data.pages.segmentation || {};
    const newer = force || !pd || !pd.saved || (segData.saved && segData.saved > pd.saved);
    if ((!force && ta.value.trim() === txt.trim()) || !newer) { posEvt({ research, found: true, imported: false }); return false; }
    ta.value = txt;
    ta.dispatchEvent(new Event("input", { bubbles: true }));
    ta.dispatchEvent(new Event("paste", { bubbles: true }));
    await settle(400, 3000);
    toast("已從研究紀錄「" + research + "」帶入市場區隔分析的結果");
    posEvt({ research, found: true, imported: true });
    return true;
  }
  async function posReload() {
    if (!research) { posEvt({ research: "" }); return "none"; }
    const rec = await readRecord(research);
    if (!rec) { posEvt({ research, found: false }); return "missing"; }
    const ok = await posImport(rec, null, true);
    stSchedule();
    return ok ? "ok" : "missing";
  }

  async function setResearch(r, opts) {
    if (stTimer || mutTimer) await recWrite();
    research = r; lastBody = ""; stFields = {};
    if (r) { try { localStorage.setItem(LS_R, r); } catch (e) {} }
    renderBar();
    const gen = ++stGen;
    if (!opts || opts.restore !== false) await recRestore(gen);
    else if (r) recWrite(true);
  }
  async function stSetFile(name) {
    if (!CONF || !name) return;
    if ((stTimer || mutTimer) && research) await recWrite();
    dataFile = name; stem = stemOf(name);
    const gen = ++stGen;
    if (!ready || (dir && !dirOK)) return;
    const { r, how } = await resolveResearch(stem);
    if (gen !== stGen) return;
    attachedNote = how === "attach" ? stem : "";
    research = r; lastBody = ""; stFields = {};
    try { localStorage.setItem(LS_R, r); } catch (e) {}
    renderBar();
    await recRestore(gen);
  }

  /* 資料夾列上的研究選單 */
  let recNames = [];
  function researchHtml() {
    if (!REC_ON || !ready || (dir && !dirOK)) return "";
    const names = recNames.slice(); if (research && !names.includes(research)) names.unshift(research);
    return " <span class=\"soft\">｜</span> 研究紀錄：<select class=\"mrs-research\">" + (research ? "" : "<option value=\"\" selected>（未選擇）</option>") +
      names.map(n => "<option value=\"" + esc(n) + "\"" + (n === research ? " selected" : "") + ">" + esc(n) + "</option>").join("") +
      "<option value=\"__new\">＋ 新研究…</option></select>" +
      (attachedNote ? " <span class=\"soft\">資料檔「" + esc(attachedNote) + "」已歸入此研究</span> <button type=\"button\" class=\"mrs-split\">改為獨立研究</button>" : "");
  }
  function bindResearch() {
    const s = bar.querySelector(".mrs-research");
    if (s) s.onchange = async () => {
      let v = s.value;
      if (v === "__new") {
        v = (prompt("新研究的名稱（會用於檔名「研究紀錄_名稱.md」）", stem || "") || "").trim().replace(/[\\/:*?"<>|]/g, "_");
        if (!v) { renderBar(); return; }
      }
      attachedNote = "";
      // 分析頁：把目前的資料檔歸入選定的研究，並寫入本頁那一節；問卷產生器：帶回該研究的內容
      if (CONF && stem) await setResearch(v, { restore: false }); else await setResearch(v);
    };
    const sp = bar.querySelector(".mrs-split");
    if (sp) sp.onclick = () => { attachedNote = ""; setResearch(stem, { restore: false }); };
  }

  if (REC_ON) {
    if (CONF) {
      document.addEventListener("change", e => { const t = e.target; if (t && t.type === "file" && t.id === CONF.input && t.files && t.files[0]) stSetFile(t.files[0].name); }, true);
      document.addEventListener("drop", e => { const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]; if (f) stSetFile(f.name); }, true);
    }
    const onEdit = e => { if (!restoring && stEligible(e.target)) stSchedule(); };
    document.addEventListener("input", onEdit, true);
    document.addEventListener("change", onEdit, true);
    new MutationObserver(ms => {
      if (!research || restoring) return;
      if (ms.every(m => { const n = m.target.nodeType === 1 ? m.target : m.target.parentElement; return !n || n.closest(".mrs, .mrs-toast"); })) return;
      clearTimeout(mutTimer); mutTimer = setTimeout(() => recWrite(), 3000);
    }).observe(document.documentElement, { childList: true, subtree: true, characterData: true });
    window.addEventListener("pagehide", () => { if (stTimer || mutTimer) recWrite(); });
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "hidden" && (stTimer || mutTimer)) recWrite(); });
  }
  /* 舊版「頁面紀錄_頁面_資料檔.json」：內容併入對應的研究紀錄 .md 後刪除，資料夾裡只留 .md */
  async function migrateOldState() {
    if (!ready || (dir && !dirOK)) return;
    let L = [];
    try { L = await listCur(); } catch (e) { return; }
    const olds = L.filter(f => /^頁面紀錄_.+\.json$/i.test(f.name));
    if (!olds.length) return;
    let moved = 0;
    for (const f of olds) {
      const m = /^頁面紀錄_([a-z0-9-]+)_(.+)\.json$/i.exec(f.name);
      if (!m) continue;
      const page = m[1], st = m[2];
      let pd = null;
      try { pd = JSON.parse(await (await f.get()).text()); } catch (e) { continue; }
      if (!pd || !pd.fields) continue;
      const recs = await listRecords();
      let r = "";
      for (const x of recs) {
        try { const { data } = parseRecord(await (await x.get()).text()); if ((data.files || []).includes(st)) { r = x.name.slice(5, -3); break; } } catch (e) {}
      }
      if (!r) r = st;
      const rec = (await readRecord(r)) || { data: { research: r, files: [], pages: {} }, secs: {} };
      rec.data.research = r;
      rec.data.files = rec.data.files || [];
      if (!rec.data.files.includes(st)) rec.data.files.push(st);
      rec.data.pages = rec.data.pages || {};
      const cur = rec.data.pages[page];
      if (!cur || !cur.saved || (pd.saved && pd.saved > cur.saved))
        rec.data.pages[page] = { saved: pd.saved || new Date(f.mtime).toISOString(), file: pd.file || "", fields: pd.fields };
      try {
        await writeCur(recName(r), new Blob([composeRecord(r, rec)], { type: "text/markdown" }));
        await removeCur(f.name);
        moved++;
      } catch (e) {}
    }
    if (moved) { toast("已把 " + moved + " 個舊版頁面紀錄（.json）併入研究紀錄（.md）"); await refresh(); }
  }
  let genTried = false;
  async function recInit() {               // 載入頁面時：問卷產生器沿用最近使用的研究
    if (!REC_ON || !ready || (dir && !dirOK)) return;
    recNames = (await listRecords()).map(f => f.name.slice(5, -3));
    if (GEN && !research) {
      const cur = localStorage.getItem(LS_R);
      if (cur && recNames.includes(cur)) await setResearch(cur);
    }
    renderBar();
  }


  window.MRStore = { save, kindOf, posReload, get folder() { return dir ? dir.name : ""; }, get research() { return research; } };

  const start = async () => {
    mountBar();
    await renderBar();                      // 先顯示卡片（讀取中），之後的步驟即使卡住也不會空白
    try { await timeout(loadFolder(), 5000); } catch (e) {}
    try { await timeout(checkPermission(), 5000); } catch (e) { dirOK = false; ready = !!dir || !HAS_FS; }
    loading = false;
    await renderBar();
    loadedName = null;
    await refresh();
    try { await migrateOldState(); } catch (e) {}
    autoLoad();
  };
  if (document.readyState === "complete") start(); else window.addEventListener("load", start);
})();
