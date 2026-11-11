// 網站更新提示（與「你論了嗎」、教學工具箱相同）：
// 定期（與切回分頁時）讀取這一頁與本站共用程式檔的最新內容，和目前載入的版本比對；
// 不同時在頁面頂端顯示提示，按下即重新載入新版。
// 用法：<script src="newver.js" data-note="（說明文字）"></script>，路徑依頁面所在層級調整。
(() => {
  const me = document.currentScript;
  const note = me && me.dataset.note || "";
  const SITE = "行銷研究工具箱";

  const css = document.createElement("style");
  css.textContent =
    ".newver{position:sticky;top:0;z-index:9999;display:flex;flex-wrap:wrap;align-items:center;gap:10px;padding:10px 16px;background:#fbf0dc;color:#9a5b00;border-bottom:1px solid #9a5b00;font:600 15px/1.5 \"Noto Sans TC\",\"PingFang TC\",\"Microsoft JhengHei\",system-ui,sans-serif}" +
    ".newver[hidden]{display:none}" +
    ".newver button{padding:5px 12px;border:1px solid #9a5b00;border-radius:8px;background:#9a5b00;color:#fff;font:inherit;cursor:pointer}" +
    ".newver small{font-weight:400;font-size:13px}" +
    "@media (prefers-color-scheme:dark){.newver{background:#3a2f1c;color:#e7b360;border-color:#e7b360}.newver button{background:#e7b360;border-color:#e7b360;color:#141612}}";
  document.head.appendChild(css);

  const bar = document.createElement("div");
  bar.className = "newver"; bar.hidden = true;
  bar.innerHTML = SITE + "已更新。<button type=\"button\">重新整理，使用新版</button>" + (note ? "<small></small>" : "");
  if (note) bar.querySelector("small").textContent = note;
  const mount = () => document.body.insertBefore(bar, document.body.firstChild);
  if (document.body) mount(); else document.addEventListener("DOMContentLoaded", mount);

  // 要比對的檔案：這一頁本身 + 本站的共用程式檔（外部 CDN 與 newver.js 本身不算）
  const files = [location.pathname].concat(
    Array.from(document.querySelectorAll("script[src]"))
      .map(s => new URL(s.getAttribute("src"), location.href))
      .filter(u => u.origin === location.origin && !/\/newver\.js$/.test(u.pathname))
      .map(u => u.pathname));
  const get = (url, cache) => fetch(url, { cache }).then(r => r.ok ? r.text() : null).catch(() => null);

  // 基準：瀏覽器目前用的那一份（優先取快取，也就是畫面正在跑的版本）
  const base = {};
  files.forEach(f => get(f === location.pathname ? location.pathname + location.search : f, "force-cache").then(t => { base[f] = t; }));

  let last = Date.now();
  const check = async () => {
    if (!bar.hidden || Date.now() - last < 60000) return; last = Date.now();
    for (const f of files) {
      if (!base[f]) continue;
      const t = await get(f + "?nv=" + Date.now(), "no-store");
      if (t && t !== base[f]) { bar.hidden = false; return; }
    }
  };
  setInterval(check, 5 * 60000);
  document.addEventListener("visibilitychange", () => { if (!document.hidden) check(); });
  window.addEventListener("focus", check);

  // 先把新版寫進瀏覽器快取（含共用程式檔），再重新載入，避免讀到快取裡的舊檔
  bar.querySelector("button").addEventListener("click", async () => {
    await Promise.all(files.map(f => get(f, "reload")));
    const u = new URL(location.href);
    u.searchParams.set("r", Date.now());
    location.href = u.href;
  });
})();
