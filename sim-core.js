// 練習資料擴充共用的模擬引擎（原始回覆版 simulate/ 與已整理資料版 simulate-clean/ 共用）
/* ===== SIM-BEGIN（純計算函式，可單獨測試） ===== */
const Sim=(function(){
  function rng(seed){ let a=seed>>>0; return function(){ a=(a+0x6D2B79F5)>>>0; let t=a; t=Math.imul(t^(t>>>15),t|1); t^=t+Math.imul(t^(t>>>7),t|61); return ((t^(t>>>14))>>>0)/4294967296; }; }
  const blank=v=>v===""||v===null||v===undefined;
  function splitMulti(s){ return String(s).split(/,\s*/).map(x=>x.trim()).filter(Boolean); }
  /* cols: [{type:"num"|"cat"|"multi"|"time"|"clear", min,max,int}]；rows：原始資料 */
  function prepare(rows,cols){
    const n=rows.length;
    const use=cols.map((c,j)=>c.type==="num"||c.type==="cat"||c.type==="multi"?j:-1).filter(j=>j>=0);
    const sets=rows.map(r=>cols.map((c,j)=>c.type==="multi"&&!blank(r[j])?new Set(splitMulti(r[j])):null));
    const pat=rows.map(r=>use.map(j=>blank(r[j])?"0":"1").join(""));
    function dist(a,b){
      let d=0, m=0;
      for (const j of use){
        const x=rows[a][j], y=rows[b][j], bx=blank(x), by=blank(y);
        m++;
        if (bx||by){ d+=bx&&by?0:1; continue; }
        const c=cols[j];
        if (c.type==="num") d+=c.max>c.min?Math.abs(x-y)/(c.max-c.min):0;
        else if (c.type==="multi"){ const A=sets[a][j], B=sets[b][j]; let inter=0; A.forEach(v=>{ if (B.has(v)) inter++; }); const uni=A.size+B.size-inter; d+=uni?1-inter/uni:0; }
        else d+=String(x).trim()===String(y).trim()?0:1;
      }
      return m?d/m:0;
    }
    return {n,use,pat,dist};
  }
  function neighbors(P,k){
    const {n,pat,dist}=P, out=new Array(n);
    for (let i=0;i<n;i++){
      let cand=[]; for (let j=0;j<n;j++) if (j!==i&&pat[j]===pat[i]) cand.push(j);
      let fallback=false;
      if (!cand.length){ fallback=true; for (let j=0;j<n;j++) if (j!==i) cand.push(j); }
      const ds=cand.map(j=>[j,dist(i,j)]).sort((a,b)=>a[1]-b[1]||a[0]-b[0]);
      out[i]={nb:ds.slice(0,k).map(x=>x[0]),fallback};
    }
    return out;
  }
  /* jit：{p:量表題微調機率, lo, hi}；量表題（整數且落在 lo～hi 內）以機率 p 上下調 1 分，超出範圍時改往另一方向 */
  function generate(rows,cols,m,k,seed,jit){
    const rand=rng(seed), P=prepare(rows,cols), NB=neighbors(P,k), n=rows.length;
    const out=[];
    for (let t=0;t<m;t++){
      const a=Math.floor(rand()*n), nb=NB[a].nb;
      const b=nb.length?nb[Math.floor(rand()*nb.length)]:a;
      const u=.2+.6*rand(), A=rows[a], B=rows[b]; /* 混合比例限制在 0.2～0.8，避免整列複製其中一人 */
      const row=cols.map((c,j)=>{
        const x=A[j], y=B[j];
        if (c.type==="clear") return "";
        if (c.type==="time"){ if (c.min==null) return blank(x)?"":x; return new Date(c.min+rand()*(c.max-c.min)); }
        if (c.type==="keep") return blank(x)?"":x;
        if (blank(x)) return "";
        if (c.type==="num"){
          let v=blank(y)?x:(rand()<u?y:x);
          if (jit&&jit.p>0&&c.int&&c.min>=jit.lo&&c.max<=jit.hi&&rand()<jit.p){
            let d=rand()<.5?-1:1; if (v+d<jit.lo||v+d>jit.hi) d=-d;
            if (v+d>=jit.lo&&v+d<=jit.hi) v+=d;
          }
          return v;
        }
        if (blank(y)) return x;
        return rand()<u?y:x;
      });
      out.push(row);
    }
    return {rows:out, fallback:NB.filter(x=>x.fallback).length};
  }
  return {generate, rng, splitMulti};
})();
/* ===== SIM-END ===== */
