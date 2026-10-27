/*
 * jamovi .omv 讀檔程式（行銷研究工具箱共用）
 * 依 jamovi 原始碼 server/jamovi/server/formatio/omv.py 的寫檔規格：
 *   - metadata.json 的 dataSet.fields 依序對應 data.bin 中的欄位；每欄 rowCount 格
 *   - Decimal：8 位元組 little-endian double，NaN 為缺漏
 *   - Text + ID：4 位元組 int，為 strings.bin 中以 \0 結尾之 UTF-8 字串的位移
 *   - 其餘（Integer、Text 名目／次序）：4 位元組 int；-2147483648 為缺漏
 *   - xdata.json：{欄名: {labels: [[值, 標籤, 匯入值, ...], ...]}}
 * 處理方式：
 *   - 文字欄（Text）一律轉為標籤文字；整數欄只有「名目」（Nominal）才轉為標籤，次序與連續保留數字
 *   - 依啟用中的篩選欄（Filter，值 0 表示排除）刪除受訪者，篩選欄本身不輸出
 *   - 欄位設定的缺漏值條件（如 "== 99"）視為缺漏
 * 需要 JSZip 與 SheetJS（XLSX）。
 */
(function(global){
  "use strict";
  const INT_MISSING = -2147483648;

  function isOmv(file){ return !!file && /\.omv$/i.test(file.name || ""); }

  function parseMissing(conds){
    return (conds || []).map(function(c){
      const m = String(c).trim().match(/^(==|!=|<=|>=|<|>)\s*(.+)$/);
      if (!m) return null;
      let v = m[2].trim();
      if (/^(['"]).*\1$/.test(v)) v = v.slice(1, -1);
      return { op: m[1], v: v };
    }).filter(Boolean);
  }
  function hitMissing(rules, value, text){
    for (const r of rules){
      const n = Number(r.v), numeric = r.v !== "" && Number.isFinite(n) && typeof value === "number";
      const a = numeric ? value : String(text), b = numeric ? n : r.v;
      if ((r.op === "==" && a == b) || (r.op === "!=" && a != b) ||
          (r.op === "<" && a < b) || (r.op === "<=" && a <= b) ||
          (r.op === ">" && a > b) || (r.op === ">=" && a >= b)) return true;
    }
    return false;
  }

  /* 回傳 { headers, rows, removed, info }；rows 為二維陣列，缺漏為 null */
  async function read(arrayBuffer){
    if (typeof JSZip === "undefined") throw new Error("JSZip 未載入");
    const zip = await JSZip.loadAsync(arrayBuffer);
    const metaFile = zip.file("metadata.json");
    const binFile = zip.file("data.bin");
    if (!metaFile || !binFile) throw new Error("不是有效的 jamovi 檔案");
    const meta = JSON.parse(await metaFile.async("string"));
    const ds = meta.dataSet || meta;
    const fields = ds.fields || [];
    const rowCount = ds.rowCount || 0;
    let xdata = {};
    const xf = zip.file("xdata.json");
    if (xf){ try { xdata = JSON.parse(await xf.async("string")); } catch (e) { xdata = {}; } }
    const bin = await binFile.async("arraybuffer");
    const sf = zip.file("strings.bin");
    const strBytes = sf ? new Uint8Array(await sf.async("arraybuffer")) : null;
    const view = new DataView(bin);
    const dec = new TextDecoder("utf-8");

    function readStr(pos){
      if (!strBytes || pos < 0 || pos >= strBytes.length) return "";
      let end = pos; while (end < strBytes.length && strBytes[end] !== 0) end++;
      return dec.decode(strBytes.subarray(pos, end));
    }

    let off = 0;
    const cols = [];
    for (const f of fields){
      let dataType = f.dataType, measure = f.measureType || "Nominal";
      if (!dataType){
        if (measure === "NominalText"){ dataType = "Text"; measure = "Nominal"; }
        else if (measure === "Continuous") dataType = "Decimal";
        else dataType = "Integer";
      }
      const isDec = dataType === "Decimal";
      const isId = dataType === "Text" && measure === "ID";
      const width = isDec ? 8 : 4;
      if (off + width * rowCount > bin.byteLength) throw new Error("data.bin 長度與欄位設定不符");
      const labels = new Map();
      ((xdata[f.name] || {}).labels || []).forEach(function(l){ if (Array.isArray(l) && l.length >= 2) labels.set(l[0], String(l[1])); });
      const isFilter = f.columnType === "Filter";
      const useLabel = !isFilter && (dataType === "Text" || (dataType === "Integer" && measure === "Nominal"));
      const miss = parseMissing(f.missingValues);
      const vals = new Array(rowCount);
      for (let i = 0; i < rowCount; i++){
        let v;
        if (isDec){
          v = view.getFloat64(off + i * 8, true);
          if (Number.isNaN(v)) v = null;
        } else {
          const x = view.getInt32(off + i * 4, true);
          if (x === INT_MISSING) v = null;
          else if (isId) v = readStr(x);
          else if (useLabel && labels.has(x)) v = labels.get(x);
          else v = x;
        }
        if (v === "") v = null;
        if (v !== null && miss.length){
          const raw = (isDec || dataType === "Text") ? v : view.getInt32(off + i * 4, true);
          if (hitMissing(miss, raw, v)) v = null;
        }
        vals[i] = v;
      }
      off += width * rowCount;
      cols.push({ name: f.name, filter: isFilter, active: f.active !== false, vals: vals });
    }

    const filters = cols.filter(function(c){ return c.filter && c.active; });
    const outCols = cols.filter(function(c){ return !(c.filter && c.active); });
    const rows = [];
    let removed = 0;
    for (let i = 0; i < rowCount; i++){
      if (filters.some(function(c){ return c.vals[i] === 0; })){ removed++; continue; }
      rows.push(outCols.map(function(c){ return c.vals[i]; }));
    }
    return {
      headers: outCols.map(function(c){ return c.name; }),
      rows: rows,
      removed: removed,
      info: { rowCount: rowCount, filters: filters.length }
    };
  }

  /* 轉成 SheetJS 工作簿（工作表名稱「資料」），讓各頁面沿用 Excel 的處理流程 */
  async function toWorkbook(arrayBuffer){
    const r = await read(arrayBuffer);
    const ws = XLSX.utils.aoa_to_sheet([r.headers].concat(r.rows));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "資料");
    wb.omvInfo = { removed: r.removed, rowCount: r.info.rowCount, filters: r.info.filters };
    return wb;
  }

  global.OMV = { isOmv: isOmv, read: read, toWorkbook: toWorkbook };
})(window);
