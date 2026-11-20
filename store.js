// 資料夾：工具箱各頁共用的檔案存放處（整個工具箱只用一個資料夾，在首頁或各頁頂端設定）。
// 1. Chrome／Edge 可選電腦上的資料夾（例如 Google 雲端硬碟裡的資料夾），檔案直接存在該資料夾內；
//    其他瀏覽器改存在這個瀏覽器內。
// 2. 各頁下載的檔案（Excel、Word 報告、圖、CSV、設定檔）照常下載，同時存進資料夾；
//    使用者上傳的檔案也會存一份，下次不必再上傳。尚未選資料夾時只照常下載。
// 3. 進入分析頁時，依檔名元素（clean_、sim_clean_、fa_、seg_ 等前綴與關鍵字）自動載入最適合、最新的檔案，可改選。
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
    await refresh(); autoLoad();
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
    let html = HOME ? "<b class=\"mrs-title\">資料夾</b>" : "<b>資料夾</b> ";
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
    if (ready && (!dir || dirOK)) html += " <span class=\"soft\">｜</span> <button type=\"button\" class=\"mrs-files\">資料夾檔案（" + files.length + "）" + (panel.hidden ? "▾" : "▴") + "</button>";
    if (HOME) {
      html += "<p class=\"soft mrs-note\">" + (!HAS_FS ? "各工具上傳與下載的檔案會存在這個瀏覽器內，分析頁會自動帶入；清除瀏覽資料或換電腦就會消失。"
        : !dir ? "選一個電腦上的資料夾（例如 Google 雲端硬碟裡的資料夾）。之後各工具上傳與下載的檔案都會自動存進去，分析頁會依檔名自動帶入，不必再手動上傳。"
        : !dirOK ? "瀏覽器重新開啟後，需要按一次「連線資料夾」，才能讀寫資料夾。"
        : "各工具上傳與下載的檔案會自動存進這個資料夾，分析頁會依檔名自動帶入最新、最適合的檔案；自行放進資料夾的檔案也讀得到。") + "</p>";
    }
    bar.innerHTML = html;
    const q = s => bar.querySelector(s);
    if (q(".mrs-grant")) q(".mrs-grant").onclick = async () => { if (await askPermission()) { await refresh(); autoLoad(); } };
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
    try { files = ready ? await listCur() : []; } catch (e) { files = []; }
    files.sort((a, b) => b.mtime - a.mtime);
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


  window.MRStore = { save, kindOf, get folder() { return dir ? dir.name : ""; } };

  const start = async () => {
    mountBar();
    await loadFolder();
    await checkPermission();
    loadedName = null;
    await refresh();
    autoLoad();
  };
  if (document.readyState === "complete") start(); else window.addEventListener("load", start);
})();
