import { useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import { INIT_PRODUCTS, INIT_SALES } from "./data.js";

const TRANSIT_DAYS = 60; // 下單後2個月出貨

const statusColors = { danger: "#ef4444", warning: "#f59e0b", ok: "#22c55e" };
const statusLabels  = { danger: "⚠️快斷貨", warning: "留意", ok: "正常" };

const CONTAINERS = [
  { key: "美麗多", label: "🪵 美麗多", color: "#8b5cf6", pallets: 48, fixedPallets: 0 },
  { key: "飼料",  label: "🥕 飼料",  color: "#f59e0b", pallets: 10, fixedPallets: 2 },
  { key: "牧草",  label: "🌿 牧草",  color: "#22c55e", pallets: 48, fixedPallets: 0 },
];

const makeShipments = () => ({
  美麗多: [
    { id:0, label:"TWMIL26/06", shipDate:"2026-05-01", arrivalDate:"2026-07-31", status:"past",    note:"已出貨，預計7/31到港", allocs:{MLMP2:960,MLMS2:840,MLMP4:400,MLMS07:30,MLMP07:90,MLMS10:20,MLMP10:40} },
    { id:1, label:"TWMIL26/07", shipDate:"2026-06-17", arrivalDate:"2026-09-15", status:"past",    note:"已出貨，預計9/15到港", allocs:{MLMP2:1200,MLMS2:720,MLMP4:320,MLMS07:60,MLMP07:60,MLMS10:20,MLMP10:40} },
    { id:2, label:"TWMIL26/08", shipDate:"2026-06-27", arrivalDate:"2026-09-25", status:"planned", note:"✅ 避開中秋後、國慶前", allocs:{MLMP2:3666,MLMS2:2538,MLMP4:280} },
    { id:3, label:"TWMIL26/09", shipDate:"2026-08-17", arrivalDate:"2026-11-15", status:"planned", note:"✅ 避開聖誕前兩週",     allocs:{MLMP2:2679,MLMS2:1833,MLMP4:490,MLMS07:240,MLMP07:120,MLMS10:39,MLMP10:78} },
    { id:4, label:"TWMIL26/10", shipDate:"2026-09-26", arrivalDate:"2026-12-25", status:"planned", note:"⚠️ 年前必到！",        allocs:{MLMP2:6627,MLMS2:141} },
  ],
  飼料: [
    { id:0, label:"FEED26/01", shipDate:"2026-07-09", arrivalDate:"2026-09-09", status:"planned", note:"", allocs:{} },
    { id:1, label:"FEED26/02", shipDate:"2026-09-09", arrivalDate:"2026-11-09", status:"planned", note:"", allocs:{} },
    { id:2, label:"FEED26/03", shipDate:"2026-11-09", arrivalDate:"2027-01-09", status:"planned", note:"", allocs:{} },
  ],
  牧草: [
    { id:0, label:"HAY26/01", shipDate:"2026-07-09", arrivalDate:"2026-09-09", status:"planned", note:"", allocs:{} },
    { id:1, label:"HAY26/02", shipDate:"2026-09-09", arrivalDate:"2026-11-09", status:"planned", note:"", allocs:{} },
    { id:2, label:"HAY26/03", shipDate:"2026-11-09", arrivalDate:"2027-01-09", status:"planned", note:"", allocs:{} },
  ],
});

function avgSales(id, months, data) {
  const d = (data[id]||[]).slice(0, months);
  return d.length ? d.reduce((a,b)=>a+b,0)/d.length : 0;
}
function addDays(dateStr, days) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  d.setDate(d.getDate()+days);
  return d.toISOString().slice(0,10);
}
function pallets(qty, p) {
  return qty>0 && p.upp>0 ? (qty/p.upp)*p.pslots : 0;
}

const Label = ({children}) => (
  <div style={{fontSize:11,color:"#94a3b8",marginBottom:3,fontWeight:600,textTransform:"uppercase",letterSpacing:0.5}}>{children}</div>
);
const BigNum = ({children,color,size=20}) => (
  <div style={{fontSize:size,fontWeight:800,color:color||"#1e293b",lineHeight:1.1}}>{children}</div>
);
const Card = ({children,borderColor,style={}}) => (
  <div style={{background:"#fff",borderRadius:14,border:`1.5px solid ${borderColor||"#e2e8f0"}`,padding:"16px 18px",marginBottom:14,...style}}>
    {children}
  </div>
);

const shipBadge = {
  past:    {bg:"#f1f5f9",color:"#64748b",text:"已出貨"},
  pending: {bg:"#fef9c3",color:"#92400e",text:"待確認"},
  planned: {bg:"#dcfce7",color:"#166534",text:"已規劃"},
};

export default function ProcurementApp() {
  const [products, setProducts]         = useState(INIT_PRODUCTS);
  const [salesData]                     = useState(INIT_SALES);
  const [avgMonths, setAvgMonths]       = useState(3);
  const [safetyDays, setSafetyDays]     = useState(7);
  const [activeTab, setActiveTab]       = useState("schedule");
  const [importStep, setImportStep]     = useState(null);
  const [allShipments, setAllShipments] = useState(makeShipments);
  const [container, setContainer]       = useState("美麗多");

  const today = new Date();
  const containerCfg = CONTAINERS.find(c=>c.key===container) || CONTAINERS[0];
  const containerColor = containerCfg.color;
  const CONTAINER_PALLETS = containerCfg.pallets - containerCfg.fixedPallets;

  const shipments = allShipments[container];
  const setShipments = useCallback((updater) => {
    setAllShipments(prev => ({...prev, [container]: typeof updater==="function" ? updater(prev[container]) : updater}));
  }, [container]);

  const catProducts = useMemo(() => products.filter(p => p.category === container), [products, container]);

  const inTransit = useMemo(() => {
    const t = {}; catProducts.forEach(p => { t[p.id]=0; });
    shipments.forEach(sh => {
      if (new Date(sh.arrivalDate) > today) {
        Object.entries(sh.allocs).forEach(([id,qty]) => { if(t[id]!==undefined) t[id]=(t[id]||0)+Number(qty||0); });
      }
    });
    return t;
  }, [shipments, catProducts]);

  const suggestions = useMemo(() => catProducts.map(p => {
    const avg   = avgSales(p.id, avgMonths, salesData);
    const daily = avg/30;
    const available = p.stock + (inTransit[p.id]||0);
    const daysLeft = daily>0 ? available/daily : 999;
    const status = daysLeft<45?"danger":daysLeft<90?"warning":"ok";
    const recQty = daily>0 ? Math.max(0, Math.ceil((TRANSIT_DAYS+15+safetyDays)*daily - available)) : 0;
    return {...p, avg, daily, daysLeft, status, inTransitQty:inTransit[p.id]||0, recQty};
  }), [catProducts, salesData, avgMonths, inTransit, safetyDays]);

  const shipmentPallets = useMemo(() => shipments.map(sh =>
    catProducts.reduce((sum,p) => sum + pallets(Number(sh.allocs[p.id]||0), p), 0)
  ), [shipments, catProducts]);

  const updateAlloc = useCallback((shipId, pid, val) => {
    setShipments(prev => prev.map(sh =>
      sh.id===shipId ? {...sh, allocs:{...sh.allocs,[pid]:Number(val)||0}} : sh
    ));
  }, [setShipments]);

  const updateShipment = useCallback((shipId, field, val) => {
    setShipments(prev => prev.map(sh => {
      if (sh.id!==shipId) return sh;
      const u = {...sh,[field]:val};
      if (field==="shipDate") u.arrivalDate = addDays(val, TRANSIT_DAYS);
      return u;
    }));
  }, [setShipments]);

  const addShipment = () => {
    const nextId = Math.max(...shipments.map(s=>s.id))+1;
    const prefix = container==="美麗多"?"TWMIL":container==="飼料"?"FEED":"HAY";
    const label = `${prefix}26/${String(nextId+1).padStart(2,"0")}`;
    setShipments(prev=>[...prev,{id:nextId,label,shipDate:"",arrivalDate:"",status:"planned",note:"",allocs:{}}]);
  };

  const handleStockImport = e => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const wb = XLSX.read(ev.target.result,{type:"array"});
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:0});
      setProducts(prev => prev.map(p => {
        const m = rows.find(r => String(r["品號"]||"").trim()===p.id);
        return m ? {...p, stock:Number(m["數量合計"]||m["公司倉"]||0)} : p;
      }));
      setImportStep(null); alert("庫存匯入完成！");
    };
    reader.readAsArrayBuffer(file);
  };

  const exportOrder = () => {
    const wb = XLSX.utils.book_new();
    shipments.forEach((sh,si) => {
      const rows = catProducts.filter(p=>Number(sh.allocs[p.id]||0)>0).map(p => {
        const qty = Number(sh.allocs[p.id]||0);
        return {"品號":p.id,"品名":p.name,"採購量":qty,"板數":pallets(qty,p).toFixed(1)};
      });
      if(rows.length===0) return;
      rows.push({"品號":"合計","品名":"","採購量":"","板數":shipmentPallets[si].toFixed(1)});
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), sh.label);
    });
    XLSX.writeFile(wb, `MOMI訂貨規劃_${container}_26-27.xlsx`);
  };

  return (
    <div style={{fontFamily:"'Noto Sans TC',system-ui,sans-serif",background:"#f1f5f9",minHeight:"100vh"}}>

      <div style={{background:"#1e293b",padding:"14px 18px"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
          <span style={{fontSize:22}}>📦</span>
          <div>
            <div style={{fontSize:16,fontWeight:800,color:"#fff"}}>MOMI 採購訂貨規劃</div>
            <div style={{fontSize:11,color:"#94a3b8"}}>輸入三個月銷售量＋即時庫存，自動計算建議採購量與到貨日</div>
          </div>
        </div>
        <div style={{display:"flex",gap:8}}>
          {CONTAINERS.map(c=>(
            <button key={c.key} onClick={()=>setContainer(c.key)} style={{
              background:container===c.key?c.color:"transparent",
              color:container===c.key?"#fff":"#94a3b8",
              border:`2px solid ${container===c.key?c.color:"#475569"}`,
              borderRadius:99,padding:"5px 14px",fontSize:13,fontWeight:700,cursor:"pointer",
              transition:"all 0.15s"
            }}>{c.label}</button>
          ))}
        </div>
      </div>

      {importStep==="stock" && (
        <div style={{position:"fixed",inset:0,background:"#0008",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:"#fff",borderRadius:14,padding:28,width:"100%",maxWidth:380}}>
            <div style={{fontSize:15,fontWeight:700,marginBottom:12}}>📥 匯入即時庫存</div>
            <div style={{fontSize:12,color:"#64748b",marginBottom:12}}>請上傳「多倉庫存表」Excel，系統會比對品號更新庫存</div>
            <input type="file" accept=".xlsx,.xls" onChange={handleStockImport} style={{marginBottom:16,width:"100%"}} />
            <button onClick={()=>setImportStep(null)} style={{width:"100%",background:"#f1f5f9",border:"none",borderRadius:8,padding:"10px",cursor:"pointer"}}>取消</button>
          </div>
        </div>
      )}

      <div style={{display:"flex",borderBottom:"1px solid #e2e8f0",background:"#fff",position:"sticky",top:0,zIndex:10}}>
        {[["schedule","🗓 船期規劃"],["status","📊 庫存狀況"],["pressure","📦 庫壓評估"],["products","✏️ 品項管理"],["info","⚙️ 設定"]].map(([key,label])=>(
          <button key={key} onClick={()=>setActiveTab(key)} style={{
            flex:1,padding:"12px 4px",fontSize:12,fontWeight:600,
            color:activeTab===key?containerColor:"#64748b",background:"none",border:"none",
            borderBottom:activeTab===key?`3px solid ${containerColor}`:"3px solid transparent",cursor:"pointer"
          }}>{label}</button>
        ))}
      </div>

      <div style={{padding:"16px 16px 40px",maxWidth:640,margin:"0 auto"}}>

        {activeTab==="schedule" && (<>
          <Card style={{background:"#f8fafc",borderColor:"#e2e8f0"}}>
            <div style={{fontSize:13,fontWeight:600,color:"#475569"}}>
              {container==="美麗多"?"🪵 美麗多貨櫃（愛沙尼亞）":container==="飼料"?"🥕 飼料貨櫃":"🌿 牧草貨櫃"}
            </div>
            <div style={{fontSize:11,color:"#94a3b8",marginTop:2}}>前置時間 {TRANSIT_DAYS} 天 ／ 貨櫃容量 {CONTAINER_PALLETS} 板{containerCfg.fixedPallets>0?`（+固定${containerCfg.fixedPallets}板，共${containerCfg.pallets}板）`:""}</div>
          </Card>

          {shipments.map((sh,si)=>{
            const pals = shipmentPallets[si];
            const isOver = pals>CONTAINER_PALLETS && sh.status!=="past";
            const pct = Math.min((pals/CONTAINER_PALLETS)*100,100);
            const badge = shipBadge[sh.status]||shipBadge.pending;
            return (
              <Card key={sh.id} borderColor={sh.status==="past"?"#e2e8f0":isOver?"#fecaca":sh.status==="planned"?containerColor+"44":"#e2e8f0"}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
                  <div>
                    <div style={{fontSize:15,fontWeight:800,color:sh.status==="past"?"#94a3b8":"#1e293b"}}>{sh.label}</div>
                    <span style={{background:badge.bg,color:badge.color,borderRadius:99,padding:"2px 8px",fontSize:11,fontWeight:700}}>{badge.text}</span>
                  </div>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontSize:18,fontWeight:800,color:isOver?"#ef4444":sh.status==="past"?"#94a3b8":"#1e293b"}}>
                      {pals.toFixed(1)}<span style={{fontSize:12,fontWeight:400,color:"#94a3b8"}}> / {CONTAINER_PALLETS}板</span>
                    </div>
                    <div style={{fontSize:11,color:isOver?"#ef4444":"#94a3b8"}}>
                      {isOver?`超出${(pals-CONTAINER_PALLETS).toFixed(1)}板`:`剩${(CONTAINER_PALLETS-pals).toFixed(1)}板`}
                    </div>
                  </div>
                </div>
                <div style={{height:6,background:"#f1f5f9",borderRadius:99,marginBottom:10}}>
                  <div style={{height:"100%",borderRadius:99,width:`${pct}%`,
                    background:sh.status==="past"?"#cbd5e1":isOver?"#ef4444":pct>85?"#f59e0b":containerColor}}/>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
                  <div>
                    <div style={{fontSize:10,color:"#94a3b8",marginBottom:2}}>出貨日</div>
                    {sh.status==="past"
                      ? <div style={{fontSize:13,fontWeight:600,color:"#94a3b8"}}>{sh.shipDate}</div>
                      : <input value={sh.shipDate} onChange={e=>updateShipment(sh.id,"shipDate",e.target.value)} type="date"
                          style={{width:"100%",border:"1px solid #e2e8f0",borderRadius:8,padding:"7px 10px",fontSize:13,boxSizing:"border-box"}}/>}
                  </div>
                  <div>
                    <div style={{fontSize:10,color:"#94a3b8",marginBottom:2}}>預計到港</div>
                    {sh.status==="past"
                      ? <div style={{fontSize:13,fontWeight:600,color:"#94a3b8"}}>{sh.arrivalDate}</div>
                      : <input value={sh.arrivalDate} onChange={e=>updateShipment(sh.id,"arrivalDate",e.target.value)} type="date"
                          style={{width:"100%",border:"1px solid #e2e8f0",borderRadius:8,padding:"7px 10px",fontSize:13,boxSizing:"border-box"}}/>}
                  </div>
                </div>
                {Object.entries(sh.allocs).filter(([,q])=>Number(q)>0).length>0 && (
                  <div style={{marginBottom:8}}>
                    {Object.entries(sh.allocs).filter(([,q])=>Number(q)>0).map(([pid,qty])=>{
                      const p = catProducts.find(x=>x.id===pid);
                      if(!p) return null;
                      return (
                        <div key={pid} style={{display:"flex",justifyContent:"space-between",marginBottom:3,color:sh.status==="past"?"#94a3b8":"#475569"}}>
                          <span style={{fontSize:12}}>{p.name}</span>
                          <span style={{fontWeight:600,whiteSpace:"nowrap"}}>{Number(qty).toLocaleString()} ／ {pallets(Number(qty),p).toFixed(1)}板</span>
                        </div>
                      );
                    })}
                  </div>
                )}
                <div style={{fontSize:11,color:sh.status==="past"?"#94a3b8":"#64748b",
                  background:sh.status==="past"?"#f8fafc":sh.note.includes("⚠️")?"#fef9c3":"#f0fdf4",
                  borderRadius:8,padding:"6px 10px"}}>
                  {sh.note||"（無備註）"}
                </div>
              </Card>
            );
          })}
          <button onClick={addShipment} style={{width:"100%",background:"none",border:"2px dashed #d1d5db",borderRadius:12,padding:"14px",fontSize:13,color:"#9ca3af",cursor:"pointer"}}>
            ＋ 新增船期
          </button>
          <div style={{marginTop:16,display:"flex",gap:8}}>
            <button onClick={exportOrder} style={{flex:1,background:containerColor,color:"#fff",border:"none",borderRadius:10,padding:"12px",fontSize:13,fontWeight:700,cursor:"pointer"}}>
              📥 匯出訂貨單 Excel
            </button>
            <button onClick={()=>setImportStep("stock")} style={{flex:1,background:"#f1f5f9",color:"#475569",border:"none",borderRadius:10,padding:"12px",fontSize:13,fontWeight:700,cursor:"pointer"}}>
              📤 更新庫存
            </button>
          </div>
        </>)}

        {activeTab==="pressure" && (<>
          <Card style={{background:"#f0fdf4",borderColor:"#bbf7d0"}}>
            <div style={{fontSize:13,fontWeight:600,color:"#166534",marginBottom:4}}>📦 {container} 庫壓評估</div>
            <div style={{fontSize:11,color:"#166534"}}>✅ 正常 0–30天 ／ 🟡 偏高 30–60天 ／ 🔴 過重 60天以上 ／ ⚠️ 斷貨風險</div>
          </Card>
          {(() => {
            const shipArr = shipments.map(sh => new Date(sh.arrivalDate));
            const pressureColors = { over:"#ef4444", high:"#f59e0b", ok:"#22c55e", zero:"#94a3b8" };
            const pressureLabels = { over:"🔴過重", high:"🟡偏高", ok:"✅正常", zero:"—" };
            return catProducts.map(prod => {
              const avg = avgSales(prod.id, avgMonths, salesData);
              const daily = avg / 30;
              if (daily === 0) return null;
              const rows = shipments.map((sh, si) => {
                const daysToArr = Math.max(0, (shipArr[si] - today) / 86400000);
                const prevOrders = shipments.slice(0, si).reduce((s, prev) => s + Number(prev.allocs[prod.id]||0), 0);
                const stockAtArr = prod.stock + prevOrders - daily * (shipArr[si] > today ? (shipArr[si]-today)/86400000 : 0);
                const orderQty = Number(sh.allocs[prod.id]||0);
                const stockAfter = stockAtArr + orderQty;
                const daysAfter = stockAfter / daily;
                const pLevel = daysAfter > 60 ? "over" : daysAfter > 30 ? "high" : daysAfter > 0 ? "ok" : "zero";
                return { sh, orderQty, daysAfter: Math.round(daysAfter), pLevel };
              });
              const p = Math.round(prod.stock / daily);
              const over = p > CONTAINER_PALLETS;
              return (
                <Card key={prod.id}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                    <div>
                      <div style={{fontSize:13,fontWeight:700}}>{prod.name}</div>
                      <div style={{fontSize:10,color:"#94a3b8",fontFamily:"monospace"}}>{prod.id}</div>
                    </div>
                    <div style={{textAlign:"right"}}>
                      <div style={{fontSize:16,fontWeight:800,color:over?"#ef4444":"#1e293b"}}>{p}天</div>
                      <div style={{fontSize:9,color:"#94a3b8"}}>/{CONTAINER_PALLETS}板</div>
                    </div>
                  </div>
                  <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                    {rows.map(({sh,orderQty,daysAfter,pLevel})=>(
                      <div key={sh.id} style={{background:"#fff",borderRadius:10,border:`1.5px solid ${over?"#fecaca":sh.status==="past"?"#f1f5f9":"#e2e8f0"}`,
                        padding:"8px 12px",minWidth:80,flexShrink:0,textAlign:"center"}}>
                        <div style={{fontSize:10,color:"#94a3b8",marginBottom:2}}>{sh.label}</div>
                        <div style={{fontSize:16,fontWeight:800,color:over?"#ef4444":sh.status==="past"?"#94a3b8":"#1e293b"}}>{orderQty>0?orderQty.toLocaleString():"—"}</div>
                        <div style={{fontSize:10,color:pressureColors[pLevel],fontWeight:600}}>{pressureLabels[pLevel]}</div>
                      </div>
                    ))}
                  </div>
                </Card>
              );
            });
          })()}
        </>)}

        {activeTab==="products" && (<>
          <Card style={{background:"#eff6ff",borderColor:"#bfdbfe"}}>
            <div style={{fontSize:13,fontWeight:600,color:"#1e40af",marginBottom:2}}>✏️ {container} 品項管理</div>
            <div style={{fontSize:11,color:"#3730a3"}}>可在各船期欄位直接填入採購數量。</div>
          </Card>
          {catProducts.map(prod=>{
            const s = suggestions.find(x=>x.id===prod.id)||{...prod,avg:0,daily:0,daysLeft:999,status:"ok",inTransitQty:0,recQty:0};
            return (
              <Card key={prod.id} borderColor={s.status==="danger"?"#fecaca":s.status==="warning"?"#fde68a":"#e2e8f0"}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
                  <div style={{flex:1,minWidth:0,paddingRight:8}}>
                    <div style={{fontSize:13,fontWeight:700,lineHeight:1.3}}>{prod.name}</div>
                    <div style={{fontSize:10,color:"#94a3b8",fontFamily:"monospace"}}>{prod.id}</div>
                  </div>
                  <span style={{background:statusColors[s.status]+"22",color:statusColors[s.status],
                    borderRadius:99,padding:"3px 10px",fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>
                    {statusLabels[s.status]}
                  </span>
                </div>
                <div style={{display:"flex",gap:12,marginBottom:10,flexWrap:"wrap"}}>
                  <div><div style={{fontSize:9,color:"#94a3b8"}}>現有庫存</div><div style={{fontSize:13,fontWeight:700}}>{prod.stock.toLocaleString()}</div></div>
                  <div><div style={{fontSize:9,color:"#94a3b8"}}>在途</div><div style={{fontSize:13,fontWeight:700,color:containerColor}}>{s.inTransitQty.toLocaleString()}</div></div>
                  <div><div style={{fontSize:9,color:"#94a3b8"}}>月均銷</div><div style={{fontSize:13,fontWeight:700}}>{Math.round(s.avg).toLocaleString()}</div></div>
                  <div><div style={{fontSize:9,color:"#94a3b8"}}>可撐天數</div>
                    <div style={{fontSize:13,fontWeight:700,color:statusColors[s.status]}}>{s.daysLeft>900?"—":Math.round(s.daysLeft)+"天"}</div>
                  </div>
                  {s.recQty>0 && <div><div style={{fontSize:9,color:"#94a3b8"}}>建議訂購</div>
                    <div style={{fontSize:13,fontWeight:800,color:s.status==="danger"?"#dc2626":"#b45309"}}>{s.recQty.toLocaleString()}</div>
                  </div>}
                </div>
                <div style={{display:"grid",gridTemplateColumns:`repeat(${Math.min(shipments.length,5)},1fr)`,gap:6}}>
                  {shipments.map((sh,si)=>{
                    const isPast = sh.status==="past";
                    return (
                      <div key={sh.id} style={{textAlign:"center"}}>
                        <div style={{fontSize:9,color:"#94a3b8",marginBottom:3}}>{sh.label}</div>
                        <input
                          type="number" min="0"
                          value={sh.allocs[prod.id]||""}
                          onChange={e=>updateAlloc(sh.id,prod.id,e.target.value)}
                          disabled={isPast}
                          style={{width:"100%",boxSizing:"border-box",
                            border:`1px solid ${isPast?"#f1f5f9":"#e2e8f0"}`,
                            borderRadius:8,padding:"6px 4px",fontSize:13,textAlign:"center",
                            background:isPast?"#f8fafc":"#fff",color:isPast?"#94a3b8":"#1e293b"}}
                        />
                      </div>
                    );
                  })}
                </div>
              </Card>
            );
          })}
        </>)}

        {activeTab==="status" && (<>
          <Card style={{background:"#eff6ff",borderColor:"#bfdbfe"}}>
            <div style={{fontSize:13,fontWeight:600,color:"#1e40af",marginBottom:2}}>📊 {container} 庫存狀況</div>
            <div style={{fontSize:11,color:"#3730a3"}}>可撐天數以月均銷量計算，在途含未到港船期數量。</div>
          </Card>
          {suggestions.map(s=>(
            <Card key={s.id} borderColor={s.status==="danger"?"#fecaca":s.status==="warning"?"#fde68a":"#e2e8f0"}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
                <div>
                  <div style={{fontSize:13,fontWeight:700}}>{s.name}</div>
                  <div style={{fontSize:10,color:"#94a3b8",fontFamily:"monospace"}}>{s.id}</div>
                </div>
                <span style={{background:statusColors[s.status]+"22",color:statusColors[s.status],borderRadius:99,padding:"3px 10px",fontSize:11,fontWeight:700}}>
                  {statusLabels[s.status]}
                </span>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"10px 20px"}}>
                <div><Label>現有庫存</Label><BigNum>{s.stock.toLocaleString()}</BigNum></div>
                <div><Label>在途庫存</Label><BigNum color={containerColor}>{s.inTransitQty.toLocaleString()}</BigNum></div>
                <div><Label>現有＋在途可撐</Label><BigNum color={statusColors[s.status]}>{s.daysLeft>900?"—":Math.round(s.daysLeft)+"天"}</BigNum></div>
                <div><Label>月均銷量</Label><BigNum>{Math.round(s.avg).toLocaleString()}</BigNum></div>
              </div>
              {s.recQty>0 && (
                <div style={{marginTop:10,padding:"8px 12px",background:s.status==="danger"?"#fef2f2":"#fffbeb",borderRadius:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <span style={{fontSize:11,fontWeight:600,color:"#92400e"}}>📦 建議訂購量</span>
                  <span style={{fontSize:18,fontWeight:800,color:s.status==="danger"?"#dc2626":"#b45309"}}>{s.recQty.toLocaleString()}</span>
                </div>
              )}
            </Card>
          ))}
        </>)}

        {activeTab==="info" && (<>
          <Card>
            <Label>計算月份數</Label>
            <select value={avgMonths} onChange={e=>setAvgMonths(Number(e.target.value))}
              style={{width:"100%",border:"1px solid #e2e8f0",borderRadius:8,padding:"8px 10px",fontSize:13,marginTop:4}}>
              {[1,2,3].map(m=><option key={m} value={m}>{m} 個月（{m===1?"6月":m===2?"5-6月":"4-6月"}）</option>)}
            </select>
          </Card>
          <Card>
            <Label>安全天數緩衝（建議訂購量用）</Label>
            <input type="number" min="0" max="30" value={safetyDays} onChange={e=>setSafetyDays(Number(e.target.value))}
              style={{width:"100%",border:"1px solid #e2e8f0",borderRadius:8,padding:"8px 10px",fontSize:13,marginTop:4,boxSizing:"border-box"}}/>
          </Card>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <Card><Label>運送天數</Label><BigNum size={24} color="#94a3b8">{TRANSIT_DAYS} 天</BigNum></Card>
            <Card><Label>貨櫃板數</Label><BigNum size={24} color="#94a3b8">{CONTAINER_PALLETS} 板</BigNum>{containerCfg.fixedPallets>0&&<div style={{fontSize:11,color:"#94a3b8",marginTop:2}}>共{containerCfg.pallets}板，固定{containerCfg.fixedPallets}板</div>}</Card>
          </div>
          <Card>
            <Label>建議訂購量公式</Label>
            <div style={{fontSize:12,color:"#64748b",lineHeight:1.6}}>
              (前置{TRANSIT_DAYS}天 + 目標庫存15天 + 安全{safetyDays}天) × 日均銷量 − (現有庫存 + 在途庫存)
            </div>
          </Card>
        </>)}

      </div>
    </div>
  );
}
