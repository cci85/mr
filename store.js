// 專案資料夾：工具箱各頁共用的檔案存放處。
// 1. 專案可自行命名（隨時可重新命名），每個專案綁一個自己的資料夾。Chrome／Edge 可選電腦上的資料夾
//    （例如 Google 雲端硬碟裡的某個研究資料夾），檔案直接存在該資料夾內；其他瀏覽器改存在這個瀏覽器內。
// 2. 各頁下載的檔案（Excel、Word 報告、圖、CSV、設定檔）照常下載，同時存進目前專案；
//    使用者上傳的檔案也會存一份，下次不必再上傳。尚未建立專案時只照常下載。
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


  /* ---------- 專案：每個專案一個名稱、綁一個資料夾 ---------- */
  // 專案清單存在 IndexedDB（kv「projects2」）：[{ id, name, dir }]
  //   dir：電腦資料夾（FileSystemDirectoryHandle）；null 表示存在這個瀏覽器內（不支援選資料夾的瀏覽器）
  // 專案名稱只是網站上的顯示名稱，可隨時重新命名，不影響資料夾本身。
  const HAS_FS = typeof window.showDirectoryPicker === "function";
  const cleanName = s => String(s || "").replace(/[\\/:*?"<>|]/g, "").trim().slice(0, 60);
  let projects = [];        // 全部專案
  let cur = null;           // 目前專案
  let curOK = false;        // 目前專案的資料夾是否已取得讀寫權限
  const CUR_KEY = "mr.currentProject";

  const loadProjects = async () => { try { projects = (await kvGet("projects2")) || []; } catch (e) { projects = []; } };
  const saveProjects = () => kvSet("projects2", projects);

  const IDB = {
    async list(id) {
      const all = await tx("files", "readonly", s => s.index("project").getAll(id));
      return all.map(f => ({ name: f.name, mtime: f.mtime, size: f.blob.size, get: async () => new File([f.blob], f.name, { type: f.blob.type, lastModified: f.mtime }) }));
    },
    async write(id, name, blob) { await tx("files", "readwrite", s => s.put({ id: id + "/" + name, project: id, name, blob, mtime: Date.now() })); },
    async remove(id, name) { await tx("files", "readwrite", s => s.delete(id + "/" + name)); }
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
  const useFS = () => !!(cur && cur.dir && curOK);
  const listCur = () => !cur ? Promise.resolve([]) : useFS() ? FS.list(cur.dir) : cur.dir ? Promise.resolve([]) : IDB.list(cur.id);
  const writeCur = (name, blob) => useFS() ? FS.write(cur.dir, name, blob) : IDB.write(cur.id, name, blob);
  const removeCur = name => useFS() ? FS.remove(cur.dir, name) : IDB.remove(cur.id, name);

  async function checkPermission() {
    curOK = false;
    if (cur && cur.dir) { try { curOK = (await cur.dir.queryPermission({ mode: "readwrite" })) === "granted"; } catch (e) { curOK = false; } }
  }
  async function askPermission() {          // 須在使用者點擊時呼叫
    if (!cur || !cur.dir) return false;
    try { curOK = (await cur.dir.requestPermission({ mode: "readwrite" })) === "granted"; } catch (e) { curOK = false; }
    return curOK;
  }
  // 選資料夾；已被其他專案使用的資料夾不接受（一個資料夾只對應一個專案）
  async function pickDir(exceptId) {
    let h;
    try { h = await window.showDirectoryPicker({ id: "mr-project", mode: "readwrite" }); } catch (e) { return null; }
    for (const p of projects) {
      if (p.id === exceptId || !p.dir) continue;
      try { if (await p.dir.isSameEntry(h)) { toast("資料夾「" + h.name + "」已是專案「" + p.name + "」的資料夾", true); return null; } } catch (e) {}
    }
    return h;
  }
  const nameTaken = (n, exceptId) => projects.some(p => p.id !== exceptId && p.name === n);

  /* ---------- 存檔 ---------- */
  async function save(name, blob, why) {
    if (!name || !blob || !cur) return;     // 尚未建立專案時只照常下載
    if (cur.dir && !curOK) await askPermission();
    try {
      if (cur.dir && !curOK) throw new Error("no permission");
      await writeCur(name, blob);
      toast("已存入專案「" + cur.name + "」：" + name);
    } catch (e) {
      try { await IDB.write(cur.id, name, blob); toast("無法寫入資料夾，已暫存在瀏覽器內：" + name, true); }
      catch (e2) { toast("存入專案失敗：" + name, true); }
    }
    refresh();
  }

  // 攔截各頁的下載：記住 blob 網址對應的內容，下載時同時存進專案
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
    if (!cur) return;
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

  function mountBar() {
    const nav = document.querySelector("nav.toolbox-back:not(.bottom)");
    if (nav) nav.after(wrap); else document.body.insertBefore(wrap, document.body.firstChild);
  }

  let files = [];
  let mode = "";            // ""、"add"（新增專案）、"rename"（重新命名）

  async function renderBar() {
    const sorted = projects.slice().sort((a, b) => a.name.localeCompare(b.name, "zh-Hant"));
    let html = "<b>專案</b> ";
    if (mode === "add") {
      html += "<input type=\"text\" class=\"mrs-name\" placeholder=\"" + (HAS_FS ? "專案名稱（空白則用資料夾名稱）" : "專案名稱") + "\"> " +
        "<button type=\"button\" class=\"mrs-ok pri\">" + (HAS_FS ? "選擇資料夾並建立" : "建立") + "</button> <button type=\"button\" class=\"mrs-cancel\">取消</button>" +
        (HAS_FS ? " <span class=\"soft\">一個專案對應一個資料夾，例如 Google 雲端硬碟裡的某個研究資料夾。</span>" : "");
    } else if (mode === "rename") {
      html += "<input type=\"text\" class=\"mrs-name\" value=\"" + esc(cur.name) + "\"> <button type=\"button\" class=\"mrs-ok pri\">確定</button> <button type=\"button\" class=\"mrs-cancel\">取消</button>" +
        " <span class=\"soft\">只改網站上的專案名稱，資料夾名稱不變。</span>";
    } else if (!cur) {
      html += "<span class=\"soft\">尚未建立專案。</span> <button type=\"button\" class=\"mrs-add pri\">＋新增專案</button>" +
        " <span class=\"soft\">建立後，各工具下載與上傳的檔案會自動存進專案" + (HAS_FS ? "資料夾" : "") + "，下次直接帶入。</span>";
    } else {
      html += "<select class=\"mrs-p\">" + sorted.map(p => "<option value=\"" + esc(p.id) + "\"" + (p.id === cur.id ? " selected" : "") + ">" + esc(p.name) + "</option>").join("") + "</select>" +
        " <button type=\"button\" class=\"mrs-ren\">重新命名</button> <button type=\"button\" class=\"mrs-add\">＋新增專案</button> <span class=\"soft\">｜</span> ";
      if (cur.dir && curOK) html += "資料夾：「" + esc(cur.dir.name) + "」";
      else if (cur.dir) html += "<span class=\"warn\">資料夾「" + esc(cur.dir.name) + "」尚未授權</span> <button type=\"button\" class=\"mrs-grant pri\">連線資料夾</button>";
      else html += "存在這個瀏覽器內" + (HAS_FS ? "" : "<span class=\"soft\">（此瀏覽器不支援選擇電腦資料夾，請用 Chrome 或 Edge）</span>");
      if (HAS_FS) html += " <button type=\"button\" class=\"mrs-pick\" title=\"改用另一個資料夾作為這個專案的資料夾；原資料夾裡的檔案不會移動或刪除\">" + (cur.dir ? "更換資料夾" : "選擇資料夾") + "</button>";
      html += " <span class=\"soft\">｜</span> <button type=\"button\" class=\"mrs-files\">專案檔案（" + files.length + "）" + (panel.hidden ? "▾" : "▴") + "</button>";
    }
    bar.innerHTML = html;
    const q = s => bar.querySelector(s);
    if (q(".mrs-p")) q(".mrs-p").onchange = e => setProject(e.target.value);
    if (q(".mrs-add")) q(".mrs-add").onclick = () => { mode = "add"; renderBar().then(() => q(".mrs-name") && q(".mrs-name").focus()); };
    if (q(".mrs-ren")) q(".mrs-ren").onclick = () => { mode = "rename"; renderBar().then(() => { const i = q(".mrs-name"); if (i) { i.focus(); i.select(); } }); };
    if (q(".mrs-cancel")) q(".mrs-cancel").onclick = () => { mode = ""; renderBar(); };
    if (q(".mrs-ok")) {
      const ok = mode === "add" ? addProject : renameProject;
      q(".mrs-ok").onclick = ok;
      q(".mrs-name").onkeydown = e => { if (e.key === "Enter") ok(); if (e.key === "Escape") { mode = ""; renderBar(); } };
    }
    if (q(".mrs-grant")) q(".mrs-grant").onclick = async () => { if (await askPermission()) { await refresh(); autoLoad(); } };
    if (q(".mrs-pick")) q(".mrs-pick").onclick = changeDir;
    if (q(".mrs-files")) q(".mrs-files").onclick = () => { panel.hidden = !panel.hidden; renderBar(); renderPanel(); };
  }

  async function addProject() {
    let n = cleanName(bar.querySelector(".mrs-name").value), dir = null;
    if (HAS_FS) {
      dir = await pickDir(null);            // 必須直接在點擊中開啟選擇視窗
      if (!dir) return;
      if (!n) n = cleanName(dir.name) || "未命名專案";
    } else if (!n) { toast("請輸入專案名稱", true); return; }
    if (nameTaken(n)) { toast("已有名為「" + n + "」的專案，請換一個名稱", true); return; }
    const p = { id: "p" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), name: n, dir };
    projects.push(p); await saveProjects();
    mode = "";
    toast("已建立專案「" + n + "」" + (dir ? "，資料夾「" + dir.name + "」" : ""));
    await setProject(p.id);
  }
  async function renameProject() {
    const n = cleanName(bar.querySelector(".mrs-name").value);
    if (!n) { toast("請輸入專案名稱", true); return; }
    if (nameTaken(n, cur.id)) { toast("已有名為「" + n + "」的專案，請換一個名稱", true); return; }
    cur.name = n; await saveProjects();
    mode = ""; toast("專案已改名為「" + n + "」");
    await refresh();
  }
  async function changeDir() {
    const h = await pickDir(cur.id);
    if (!h) return;
    cur.dir = h; curOK = true; await saveProjects();
    toast("專案「" + cur.name + "」改用資料夾「" + h.name + "」");
    loadedName = null;
    await refresh(); autoLoad();
  }

  let armed = null;   // 刪除前需再按一次
  async function renderPanel() {
    if (panel.hidden || !cur) return;
    let html = "";
    if (useFS()) {
      const n = (await IDB.list(cur.id).catch(() => [])).length;
      if (n) html += "<p class=\"warn\">這個專案還有 " + n + " 個檔案暫存在瀏覽器內。<button type=\"button\" class=\"mrs-move\">搬到資料夾「" + esc(cur.dir.name) + "」</button></p>";
    }
    if (!files.length) {
      html += "<p class=\"soft\">專案「" + esc(cur.name) + "」目前沒有檔案。在各工具下載的檔案與上傳的檔案，都會自動存到這裡。</p>";
    } else {
      html += "<table><thead><tr><th>檔名</th><th>類型</th><th>時間</th><th>大小</th><th></th></tr></thead><tbody>" +
        files.map((f, i) => "<tr><td class=\"n\">" + esc(f.name) + "</td><td><span class=\"mrs-tag\">" + kindOf(f.name)[2] + "</span></td><td>" + when(f.mtime) + "</td><td>" + kb(f.size) +
          "</td><td style=\"white-space:nowrap\"><button type=\"button\" data-dl=\"" + i + "\">下載</button> <button type=\"button\" data-rm=\"" + i + "\">" + (armed === f.name ? "確定刪除？" : "刪除") + "</button>" + (armed === f.name ? " <button type=\"button\" data-cancel>取消</button>" : "") + "</td></tr>").join("") +
        "</tbody></table>";
      if (!cur.dir) html += "<p class=\"soft\">檔案存在這個瀏覽器內；清除瀏覽資料或換電腦就會消失。<button type=\"button\" class=\"mrs-zip\">全部下載（zip）</button></p>";
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
    for (const f of await IDB.list(cur.id)) {
      try { await FS.write(cur.dir, f.name, await f.get()); await IDB.remove(cur.id, f.name); n++; } catch (e) {}
    }
    toast("已搬移 " + n + " 個檔案到「" + cur.dir.name + "」");
    refresh();
  }

  function loadScript(src) { return new Promise((res, rej) => { const s = document.createElement("script"); s.src = src; s.onload = res; s.onerror = rej; document.head.appendChild(s); }); }
  async function zipAll() {
    try {
      if (typeof JSZip === "undefined") await loadScript("https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js");
      const z = new JSZip();
      for (const f of files) z.file(f.name, await f.get());
      const blob = await z.generateAsync({ type: "blob" });
      const a = document.createElement("a"); a.href = origCreate.call(URL, blob); a.download = cur.name + ".zip";
      document.body.appendChild(a); origClick.call(a); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 3000);
    } catch (e) { toast("壓縮失敗，請改為逐一下載", true); }
  }

  async function refresh() {
    try { files = await listCur(); } catch (e) { files = []; }
    files.sort((a, b) => b.mtime - a.mtime);
    await renderBar(); await renderPanel(); renderAuto();
  }

  async function setProject(id) {
    cur = projects.find(p => p.id === id) || null;
    if (cur) localStorage.setItem(CUR_KEY, cur.id); else localStorage.removeItem(CUR_KEY);
    await checkPermission();
    loadedName = null;
    await refresh();
    autoLoad();
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
    if (!cur) { autoBox.innerHTML = "<span class=\"soft\">建立專案（上方「＋新增專案」）後，上傳與下載的檔案會自動存進專案，下次直接帶入。</span>"; return; }
    if (cur.dir && !curOK) { autoBox.innerHTML = "<span class=\"warn\">專案資料夾尚未授權，請按上方「連線資料夾」，即可自動帶入專案中的檔案。</span>"; return; }
    if (!c.length) {
      autoBox.innerHTML = "<span class=\"soft\">專案「" + esc(cur.name) + "」中還沒有適合本頁的檔案，請在下方上傳；上傳的檔案會自動存進專案，下次直接帶入。</span>";
      return;
    }
    autoBox.innerHTML = (loadedName ? "已從專案帶入：" : "專案中的檔案：") +
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


  window.MRStore = { save, kindOf, get project() { return cur ? cur.name : ""; } };

  const start = async () => {
    mountBar();
    await loadProjects();
    if (!HAS_FS) projects = projects.filter(p => !p.dir);   // 無法使用資料夾的瀏覽器只列瀏覽器內的專案
    const want = localStorage.getItem(CUR_KEY);
    const first = projects.find(p => p.id === want) || projects[0];
    await setProject(first ? first.id : null);
  };
  if (document.readyState === "complete") start(); else window.addEventListener("load", start);
})();
