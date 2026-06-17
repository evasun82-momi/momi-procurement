import { useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";

const CONTAINER_PALLETS = 48;
const TRANSIT_DAYS = 90;

const statusColors = { danger: "#ef4444", warning: "#f59e0b", ok: "#22c55e" };
const statusLabels  = { danger: "⚠️快斷貨", warning: "留意", ok: "正常" };

const INIT_PRODUCTS = [
  { id: "MLMP2",   name: "美麗多木質墊料(大顆粒)-10L",   upp: 141, pslots: 1,   stock: 3691, moq: 141 },
  { id: "MLMS2",   name: "美麗多木質墊料(小顆粒)-10L",   upp: 141, pslots: 1,   stock: 4708, moq: 141 },
  { id: "MLMP4",   name: "美麗多木質墊料(大顆粒)-20L",   upp: 70,  pslots: 1,   stock: 34,   moq: 70  },
  { id: "MLMS07",  name: "美麗多木質墊料(小顆粒)-3.5L",  upp: 120, pslots: 1,   stock: 111,  moq: 120 },
  { id: "MLMP07",  name: "美麗多木質墊料(大顆粒)-3.5L",  upp: 120, pslots: 1,   stock: 0,    moq: 120 },
  { id: "MLMS10",  name: "美麗多木質墊料(小顆粒)-50L",   upp: 39,  pslots: 2,   stock: 3,    moq: 39  },
  { id: "MLMP10",  name: "美麗多木質墊料(大顆粒)-50L",   upp: 39,  pslots: 2,   stock: 12,   moq: 39  },
  { id: "MMC5",    name: "美麗多木刨花(小)-25L",         upp: 48,  pslots: 1,   stock: 616,  moq: 48  },
  { id: "MMCP5",   name: "美麗多木刨花(大)-25L",         upp: 48,  pslots: 1,   stock: 1471, moq: 48  },
  { id: "MLMN005", name: "美麗多木質墊料(絲)0.05kg",     upp: 200, pslots: 0.5, stock: 593,  moq: 200 },
  { id: "MMR10L",  name: "美麗多實驗室無塵純萃天然木梗", upp: 30,  pslots: 1,   stock: 545,  moq: 30  },
];

const INIT_SALES = {
  MLMP2:   [1823,2283,2012,1244,2013,1346,1035,1537,1899,1325],
  MLMS2:   [1445,1422,1433,1018,1555,1671,1215,1807,2118,1476],
  MLMP4:   [294, 303, 406, 251, 347, 325, 212, 256, 394, 299],
  MLMS07:  [189, 116, 157, 90,  102, 72,  51,  97,  51,  100],
  MLMP07:  [150, 150, 150, 150, 150, 136, 72,  134, 111, 109], // 依估計需求調整,
  MLMS10:  [33,  19,  28,  1,   20,  9,   11,  39,  10,  21],
  MLMP10:  [25,  39,  36,  21,  45,  17,  26,  20,  0,   17],
  MMC5:    [60,  102, 67,  53,  101, 77,  76,  127, 174, 85],
  MMCP5:   [99,  159, 165, 125, 150, 127, 97,  130, 165, 131],
  MLMN005: [58,  72,  56,  416, 714, 73,  16,  90,  52,  80],
  MMR10L:  [30,  97,  36,  47,  39,  53,  59,  54,  89,  72],
};

const INIT_SHIPMENTS = [
  {
    id:0, label:"TWMIL26/06", shipDate:"2026-05-01", arrivalDate:"2026-07-31",
    status:"past", note:"已出貨，預計7/31到港",
    allocs:{MLMP2:960,MLMS2:840,MLMP4:400,MLMS07:30,MLMP07:90,MLMS10:20,MLMP10:40,MMC5:0,MMCP5:0,MLMN005:0,MMR10L:0}
  },
  {
    id:1, label:"TWMIL26/07", shipDate:"2026-06-17", arrivalDate:"2026-09-15",
    status:"past", note:"已出貨，預計9/15到港",
    allocs:{MLMP2:1200,MLMS2:720,MLMP4:320,MLMS07:60,MLMP07:60,MLMS10:20,MLMP10:40,MMC5:0,MMCP5:0,MLMN005:0,MMR10L:0}
  },
  {
    id:2, label:"TWMIL26/08", shipDate:"2026-06-27", arrivalDate:"2026-09-25",
    status:"planned", note:"✅ 避開中秋後、國慶前",
    allocs:{MLMP2:3666,MLMS2:2538,MLMP4:280,MLMS07:0,MLMP07:0,MLMS10:0,MLMP10:0,MMC5:0,MMCP5:0,MLMN005:0,MMR10L:0}
  },
  {
    id:3, label:"TWMIL26/09", shipDate:"2026-08-17", arrivalDate:"2026-11-15",
    status:"planned", note:"✅ 避開聖誕前兩週",
    allocs:{MLMP2:2679,MLMS2:1833,MLMP4:490,MLMS07:240,MLMP07:120,MLMS10:39,MLMP10:78,MMC5:0,MMCP5:0,MLMN005:0,MMR10L:0}
  },
  {
    id:4, label:"TWMIL26/10", shipDate:"2026-09-26", arrivalDate:"2026-12-25",
    status:"planned", note:"⚠️ 年前必到！春節1/29，最晚1/15到港",
    allocs:{MLMP2:6627,MLMS2:141,MLMP4:0,MLMS07:0,MLMP07:0,MLMS10:0,MLMP10:0,MMC5:0,MMCP5:0,MLMN005:0,MMR10L:0}
  },
];

function avgSales(id, months, data) {
  const d = (data[id]||[]).slice(0, months);
  return d.length ? d.reduce((a,b)=>a+b,0)/d.length : 0;
}
function roundMOQ(qty, moq) { return qty<=0?0:Math.ceil(qty/moq)*moq; }
function addDays(dateStr, days) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  d.setDate(d.getDate()+days);
  return d.toISOString().slice(0,10);
}
function pallets(qty, p) {
  return qty>0 ? (qty/p.upp)*p.pslots : 0;
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
  const [products, setProducts] = useState(INIT_PRODUCTS);
  const [salesData]             = useState(INIT_SALES);
  const [avgMonths, setAvgMonths] = useState(3);
  const [safetyDays, setSafetyDays] = useState(7);
  const [activeTab, setActiveTab] = useState("schedule");
  const [importStep, setImportStep] = useState(null);
  const [shipments, setShipments] = useState(INIT_SHIPMENTS);

  const today = new Date("2026-06-12");

  // In-transit = not yet arrived
  const inTransit = useMemo(() => {
    const t = {}; products.forEach(p => { t[p.id]=0; });
    shipments.forEach(sh => {
      if (new Date(sh.arrivalDate) > today) {
        Object.entries(sh.allocs).forEach(([id,qty]) => { t[id]=(t[id]||0)+Number(qty||0); });
      }
    });
    return t;
  }, [shipments, products]);

  const totalOrdered = useMemo(() => {
    const t = {}; products.forEach(p => { t[p.id]=0; });
    shipments.forEach(sh => {
      Object.entries(sh.allocs).forEach(([id,qty]) => { t[id]=(t[id]||0)+Number(qty||0); });
    });
    return t;
  }, [shipments, products]);

  const suggestions = useMemo(() => products.map(p => {
    const avg   = avgSales(p.id, avgMonths, salesData);
    const daily = avg/30;
    const available = p.stock + inTransit[p.id];
    const daysLeft = daily>0 ? available/daily : 999;
    const status = daysLeft<45?"danger":daysLeft<90?"warning":"ok";
    return {...p, avg, daily, daysLeft, status, inTransitQty:inTransit[p.id], totalOrdered:totalOrdered[p.id]||0};
  }), [products, salesData, avgMonths, inTransit, totalOrdered]);

  const shipmentPallets = useMemo(() => shipments.map(sh =>
    products.reduce((sum,p) => sum + pallets(Number(sh.allocs[p.id]||0), p), 0)
  ), [shipments, products]);

  const updateAlloc = useCallback((shipId, pid, val) => {
    setShipments(prev => prev.map(sh =>
      sh.id===shipId ? {...sh, allocs:{...sh.allocs,[pid]:Number(val)||0}} : sh
    ));
  }, []);

  const updateShipment = useCallback((shipId, field, val) => {
    setShipments(prev => prev.map(sh => {
      if (sh.id!==shipId) return sh;
      const u = {...sh,[field]:val};
      if (field==="shipDate") u.arrivalDate = addDays(val, TRANSIT_DAYS);
      return u;
    }));
  }, []);

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
      const rows = products.filter(p=>Number(sh.allocs[p.id]||0)>0).map(p => {
        const qty = Number(sh.allocs[p.id]||0);
        return {"品號":p.id,"品名":p.name,"採購量":qty,"板數":pallets(qty,p).toFixed(1)};
      });
      rows.push({"品號":"合計","品名":"","採購量":"","板數":shipmentPallets[si].toFixed(1)});
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), sh.label);
    });
    XLSX.writeFile(wb, "MOMI訂貨規劃26-27.xlsx");
  };

  return (
    <div style={{fontFamily:"'Noto Sans TC',system-ui,sans-serif",background:"#f1f5f9",minHeight:"100vh"}}>

      {/* Header */}
      <div style={{background:"#1e293b",padding:"14px 18px"}}>
        <div style={{color:"#64748b",fontSize:11,letterSpacing:2}}>MOMI 美麗多</div>
        <div style={{color:"#fff",fontSize:17,fontWeight:800,marginTop:2,marginBottom:12}}>採購訂貨規劃 2026–27</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <button onClick={()=>setImportStep("stock")}
            style={{background:"#334155",color:"#e2e8f0",border:"none",borderRadius:8,padding:"8px 14px",fontSize:13,cursor:"pointer"}}>
            📥 匯入庫存
          </button>
          <button onClick={exportOrder}
            style={{background:"#3b82f6",color:"#fff",border:"none",borderRadius:8,padding:"8px 14px",fontSize:13,cursor:"pointer",fontWeight:700}}>
            📤 匯出訂單
          </button>
        </div>
      </div>

      {importStep && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:100,display:"flex",alignItems:"center",justifyContent:"center",padding:24}}>
          <div style={{background:"#fff",borderRadius:16,padding:28,width:"100%",maxWidth:380}}>
            <div style={{fontSize:16,fontWeight:700,marginBottom:8}}>匯入多倉庫存表</div>
            <div style={{color:"#64748b",fontSize:13,marginBottom:16}}>從鼎新A1「多倉庫存表」轉出Excel，需有欄位：品號、數量合計</div>
            <input type="file" accept=".xlsx,.xls" onChange={handleStockImport} style={{marginBottom:16,width:"100%"}} />
            <button onClick={()=>setImportStep(null)}
              style={{background:"#f1f5f9",border:"none",borderRadius:8,padding:"10px 20px",cursor:"pointer"}}>取消</button>
          </div>
        </div>
      )}

      {/* Tabs */}
      <div style={{background:"#fff",borderBottom:"1px solid #e2e8f0",display:"flex"}}>
        {[["schedule","船期總覽"],["alloc","分配訂貨"],["pressure","庫壓評估"],["suggest","庫存狀況"],["products","品項"],["settings","參數"]].map(([key,label])=>(
          <button key={key} onClick={()=>setActiveTab(key)} style={{
            flex:1,background:"none",border:"none",
            borderBottom:activeTab===key?"3px solid #3b82f6":"3px solid transparent",
            color:activeTab===key?"#3b82f6":"#64748b",
            padding:"11px 2px",fontSize:12,cursor:"pointer",fontWeight:activeTab===key?700:400
          }}>{label}</button>
        ))}
      </div>

      <div style={{padding:"16px 16px 40px",maxWidth:640,margin:"0 auto"}}>

        {/* ══ 船期總覽 ══ */}
        {activeTab==="schedule" && (<>
          <Card style={{background:"#fffbeb",borderColor:"#fde68a"}}>
            <div style={{fontSize:12,fontWeight:600,color:"#92400e",marginBottom:4}}>📅 船期建議已避開台灣＋愛沙尼亞假期</div>
            <div style={{fontSize:11,color:"#78350f",lineHeight:1.7}}>
              🇪🇪 愛沙尼亞：仲夏節(6/23)、獨立紀念日(2/24)、聖誕(12/24-26)、新年(1/1)<br/>
              🇹🇼 台灣：春節、228、清明、端午、中秋、國慶、聖誕/行憲
            </div>
          </Card>

          {shipments.map((sh,si)=>{
            const pals = shipmentPallets[si];
            const isOver = pals>CONTAINER_PALLETS && sh.status!=="past";
            const pct = Math.min((pals/CONTAINER_PALLETS)*100,100);
            const badge = shipBadge[sh.status]||shipBadge.pending;
            return (
              <Card key={sh.id} borderColor={sh.status==="past"?"#e2e8f0":isOver?"#fecaca":sh.status==="planned"?"#bbf7d0":"#e2e8f0"}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
                  <div>
                    <div style={{fontSize:15,fontWeight:800,color:sh.status==="past"?"#94a3b8":"#1e293b"}}>{sh.label}</div>
                    <span style={{background:badge.bg,color:badge.color,borderRadius:99,padding:"2px 8px",fontSize:11,fontWeight:700}}>{badge.text}</span>
                  </div>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontSize:18,fontWeight:800,color:isOver?"#ef4444":sh.status==="past"?"#94a3b8":"#1e293b"}}>
                      {pals.toFixed(1)}<span style={{fontSize:12,fontWeight:400,color:"#94a3b8"}}> / {CONTAINER_PALLETS}板</span>
                    </div>
                    <div style={{fontSize:11,color:isOver?"#ef4444":"#22c55e",fontWeight:600}}>
                      {isOver?`超出${(pals-CONTAINER_PALLETS).toFixed(1)}板`:`剩${(CONTAINER_PALLETS-pals).toFixed(1)}板`}
                    </div>
                  </div>
                </div>

                {pals>0 && (
                  <div style={{background:"#f1f5f9",borderRadius:99,height:8,marginBottom:12,overflow:"hidden"}}>
                    <div style={{height:"100%",borderRadius:99,width:`${pct}%`,
                      background:sh.status==="past"?"#cbd5e1":isOver?"#ef4444":pct>85?"#f59e0b":"#3b82f6"}}/>
                  </div>
                )}

                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:8}}>
                  <div>
                    <Label>出貨日</Label>
                    {sh.status==="past"
                      ? <div style={{fontSize:14,fontWeight:600,color:"#94a3b8"}}>{sh.shipDate}</div>
                      : <input type="date" value={sh.shipDate} onChange={e=>updateShipment(sh.id,"shipDate",e.target.value)}
                          style={{width:"100%",border:"1px solid #e2e8f0",borderRadius:8,padding:"7px 10px",fontSize:13,boxSizing:"border-box"}}/>
                    }
                  </div>
                  <div>
                    <Label>預計到港</Label>
                    {sh.status==="past"
                      ? <div style={{fontSize:14,fontWeight:600,color:"#94a3b8"}}>{sh.arrivalDate}</div>
                      : <input type="date" value={sh.arrivalDate} onChange={e=>updateShipment(sh.id,"arrivalDate",e.target.value)}
                          style={{width:"100%",border:"1px solid #e2e8f0",borderRadius:8,padding:"7px 10px",fontSize:13,boxSizing:"border-box"}}/>
                    }
                  </div>
                </div>

                {/* 品項摘要 */}
                {pals>0 && (
                  <div style={{background:"#f8fafc",borderRadius:10,padding:"10px 12px",fontSize:12}}>
                    {products.filter(p=>Number(sh.allocs[p.id]||0)>0).map(p=>{
                      const qty = Number(sh.allocs[p.id]);
                      const pls = pallets(qty,p);
                      return (
                        <div key={p.id} style={{display:"flex",justifyContent:"space-between",marginBottom:3,color:sh.status==="past"?"#94a3b8":"#475569"}}>
                          <span>{p.name.replace("美麗多木質墊料","").replace("美麗多","")||p.name}</span>
                          <span style={{fontWeight:600}}>{qty.toLocaleString()}個 ／ {pls.toFixed(1)}板</span>
                        </div>
                      );
                    })}
                  </div>
                )}

                {sh.note && (
                  <div style={{fontSize:11,color:sh.status==="past"?"#94a3b8":"#64748b",
                    background:sh.status==="past"?"#f8fafc":sh.note.includes("⚠️")?"#fef9c3":"#f0fdf4",
                    borderRadius:8,padding:"7px 10px",marginTop:8}}>
                    {sh.note}
                  </div>
                )}
              </Card>
            );
          })}
        </>)}

        {/* ══ 分配訂貨 ══ */}
        {activeTab==="alloc" && (<>
          {/* Mini pallet summary */}
          <div style={{display:"flex",gap:8,overflowX:"auto",marginBottom:14,paddingBottom:2}}>
            {shipments.map((sh,si)=>{
              const p = shipmentPallets[si];
              const over = p>CONTAINER_PALLETS;
              return (
                <div key={sh.id} style={{background:"#fff",borderRadius:10,border:`1.5px solid ${over?"#fecaca":sh.status==="past"?"#f1f5f9":"#e2e8f0"}`,
                  padding:"8px 12px",minWidth:80,flexShrink:0,textAlign:"center"}}>
                  <div style={{fontSize:9,fontWeight:700,color:"#64748b"}}>{sh.label}</div>
                  <div style={{fontSize:16,fontWeight:800,color:over?"#ef4444":sh.status==="past"?"#94a3b8":"#1e293b"}}>{p.toFixed(1)}</div>
                  <div style={{fontSize:9,color:"#94a3b8"}}>/{CONTAINER_PALLETS}板</div>
                  {over && <div style={{fontSize:9,color:"#ef4444",fontWeight:600}}>超出!</div>}
                </div>
              );
            })}
          </div>

          {products.map(prod=>{
            const s = suggestions.find(x=>x.id===prod.id)||prod;
            return (
              <Card key={prod.id} borderColor={s.status==="danger"?"#fecaca":s.status==="warning"?"#fde68a":"#e2e8f0"}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
                  <div style={{flex:1,marginRight:8}}>
                    <div style={{fontSize:13,fontWeight:700,lineHeight:1.4}}>{prod.name}</div>
                    <div style={{fontSize:10,color:"#94a3b8",fontFamily:"monospace"}}>{prod.id}</div>
                  </div>
                  <span style={{background:statusColors[s.status]+"22",color:statusColors[s.status],
                    borderRadius:99,padding:"2px 8px",fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>
                    {statusLabels[s.status]}
                  </span>
                </div>

                <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:12,
                  background:"#f8fafc",borderRadius:10,padding:"10px 12px"}}>
                  <div><div style={{fontSize:9,color:"#94a3b8"}}>現有庫存</div><div style={{fontSize:13,fontWeight:700}}>{prod.stock.toLocaleString()}</div></div>
                  <div><div style={{fontSize:9,color:"#94a3b8"}}>在途</div><div style={{fontSize:13,fontWeight:700,color:"#3b82f6"}}>{s.inTransitQty.toLocaleString()}</div></div>
                  <div><div style={{fontSize:9,color:"#94a3b8"}}>月均銷</div><div style={{fontSize:13,fontWeight:700}}>{Math.round(s.avg).toLocaleString()}</div></div>
                  <div>
                    <div style={{fontSize:9,color:"#94a3b8"}}>可撐天數</div>
                    <div style={{fontSize:13,fontWeight:700,color:statusColors[s.status]}}>
                      {s.daysLeft>900?"—":Math.round(s.daysLeft)+"天"}
                    </div>
                  </div>
                </div>

                <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:6}}>
                  {shipments.map((sh)=>{
                    const qty = Number(sh.allocs[prod.id]||0);
                    const pls = pallets(qty,prod);
                    const isPast = sh.status==="past";
                    return (
                      <div key={sh.id}>
                        <div style={{fontSize:9,color:isPast?"#94a3b8":"#475569",fontWeight:600,
                          marginBottom:3,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{sh.label}</div>
                        <input type="number" value={qty||""} placeholder="0"
                          disabled={isPast}
                          onChange={e=>updateAlloc(sh.id,prod.id,e.target.value)}
                          style={{width:"100%",boxSizing:"border-box",
                            border:`1.5px solid ${isPast?"#f1f5f9":qty>0?"#3b82f6":"#e2e8f0"}`,
                            borderRadius:8,padding:"7px 5px",fontSize:12,textAlign:"right",
                            background:isPast?"#f8fafc":qty>0?"#eff6ff":"#fff",
                            fontWeight:600,color:isPast?"#94a3b8":"#1e293b"}}/>
                        <div style={{fontSize:9,color:"#94a3b8",textAlign:"right",marginTop:2}}>{qty>0?pls.toFixed(1)+"板":"—"}</div>
                      </div>
                    );
                  })}
                </div>
              </Card>
            );
          })}
        </>)}

        {/* ══ 庫存狀況 ══ */}
        {activeTab==="suggest" && (<>
          <Card style={{background:"#eff6ff",borderColor:"#bfdbfe"}}>
            <div style={{fontSize:13,fontWeight:600,color:"#1e40af",marginBottom:2}}>📊 現有＋在途庫存狀況</div>
            <div style={{fontSize:11,color:"#3730a3"}}>在途包含06/07已出貨但未到港的數量。可撐天數以月均銷量計算。</div>
          </Card>
          {suggestions.map(s=>(
            <Card key={s.id} borderColor={s.status==="danger"?"#fecaca":s.status==="warning"?"#fde68a":"#e2e8f0"}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
                <div>
                  <div style={{fontSize:13,fontWeight:700}}>{s.name}</div>
                  <div style={{fontSize:10,color:"#94a3b8",fontFamily:"monospace"}}>{s.id}</div>
                </div>
                <span style={{background:statusColors[s.status]+"22",color:statusColors[s.status],
                  borderRadius:99,padding:"3px 10px",fontSize:11,fontWeight:700}}>
                  {statusLabels[s.status]}
                </span>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"10px 20px",marginBottom:12}}>
                <div><Label>現有庫存</Label><BigNum>{s.stock.toLocaleString()}</BigNum></div>
                <div><Label>在途庫存</Label><BigNum color="#3b82f6">{s.inTransitQty.toLocaleString()}</BigNum></div>
                <div>
                  <Label>現有＋在途可撐</Label>
                  <BigNum color={statusColors[s.status]}>{s.daysLeft>900?"—":Math.round(s.daysLeft)+"天"}</BigNum>
                </div>
                <div><Label>月均銷量</Label><BigNum>{Math.round(s.avg).toLocaleString()}</BigNum></div>
              </div>
              <div style={{background:"#f8fafc",borderRadius:10,padding:"10px 12px"}}>
                <Label>各期已規劃數量</Label>
                <div style={{display:"flex",gap:8,flexWrap:"wrap",marginTop:4}}>
                  {shipments.map(sh=>{
                    const qty = Number(sh.allocs[s.id]||0);
                    if (qty===0) return null;
                    return (
                      <div key={sh.id} style={{background:sh.status==="past"?"#f1f5f9":"#eff6ff",
                        borderRadius:8,padding:"4px 10px",fontSize:12}}>
                        <span style={{color:"#64748b"}}>{sh.label}：</span>
                        <span style={{fontWeight:700,color:sh.status==="past"?"#94a3b8":"#3b82f6"}}>{qty.toLocaleString()}</span>
                      </div>
                    );
                  })}
                </div>
              </div>
            </Card>
          ))}
        </>)}

        {/* ══ 品項設定 ══ */}
        {activeTab==="products" && (<>
          <div style={{fontSize:12,color:"#94a3b8",marginBottom:14}}>
            💡 每板數量：10L=141包/板，20L=70包/板，50L=39包（占2板位）
          </div>
          {products.map((p,i)=>(
            <Card key={p.id}>
              <div style={{fontSize:14,fontWeight:700,marginBottom:2}}>{p.name}</div>
              <div style={{fontSize:10,color:"#94a3b8",fontFamily:"monospace",marginBottom:14}}>{p.id}</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
                {[
                  {label:"現有庫存",  key:"stock",  step:1,   unit:"個"},
                  {label:"每板數量",  key:"upp",    step:1,   unit:"個/板"},
                  {label:"佔板位數",  key:"pslots", step:0.5, unit:"板位"},
                  {label:"MOQ",      key:"moq",    step:1,   unit:"個"},
                ].map(({label,key,step,unit})=>(
                  <div key={key}>
                    <Label>{label}</Label>
                    <div style={{display:"flex",alignItems:"center",gap:6}}>
                      <input type="number" step={step} value={p[key]}
                        onChange={e=>setProducts(prev=>prev.map((x,j)=>j===i?{...x,[key]:Number(e.target.value)}:x))}
                        style={{flex:1,minWidth:0,border:"1px solid #e2e8f0",borderRadius:8,padding:"9px 10px",fontSize:15,textAlign:"right",fontWeight:600}}/>
                      <span style={{fontSize:11,color:"#94a3b8",whiteSpace:"nowrap"}}>{unit}</span>
                    </div>
                  </div>
                ))}
              </div>
            </Card>
          ))}
        </>)}


        {/* ══ 庫壓評估 ══ */}
        {activeTab==="pressure" && (<>
          <Card style={{background:"#f0fdf4",borderColor:"#bbf7d0"}}>
            <div style={{fontSize:13,fontWeight:600,color:"#166534",marginBottom:4}}>📦 庫壓評估說明</div>
            <div style={{fontSize:11,color:"#166534",lineHeight:1.7}}>
              計算每個品項在各船期到港時的預估庫存（含現有＋在途＋已規劃採購量）。<br/>
              ✅ 正常：到港時剩 0–30天 庫存（庫壓最低）<br/>
              🟡 偏高：到港時剩 30–60天 庫存<br/>
              🔴 過重：到港時剩 超過60天 庫存<br/>
              ⚠️ 斷貨風險：到港前庫存歸零
            </div>
          </Card>
          {(() => {
            const today = new Date("2026-06-12");
            const shipArr = shipments.map(sh => new Date(sh.arrivalDate));
            return products.map(prod => {
              const avg = avgSales(prod.id, avgMonths, salesData);
              const daily = avg / 30;
              if (daily === 0) return null;
              // calc stock at each shipment arrival
              const rows = shipments.map((sh, si) => {
                const daysToArr = Math.max(0, (shipArr[si] - today) / 86400000);
                const prevOrders = shipments.slice(0, si).reduce((s, prev) => s + Number(prev.allocs[prod.id]||0), 0);
                const thisOrder = Number(sh.allocs[prod.id]||0);
                const stockAtArr = Math.max(0, prod.stock + (inTransit[prod.id]||0) + prevOrders - daily * daysToArr);
                const totalAtArr = stockAtArr + thisOrder;
                const daysOfStock = totalAtArr / daily;
                const pressure = daysOfStock > 60 ? "over" : daysOfStock > 30 ? "high" : daysOfStock > 0 ? "ok" : "zero";
                const stockoutBefore = stockAtArr <= 0 && sh.status !== "past";
                return { sh, si, stockAtArr, thisOrder, totalAtArr, daysOfStock, pressure, stockoutBefore };
              });
              const hasIssue = rows.some(r => r.pressure === "over" || r.pressure === "zero" || r.stockoutBefore);
              const pressureColors = { over:"#ef4444", high:"#f59e0b", ok:"#22c55e", zero:"#94a3b8" };
              const pressureLabels = { over:"🔴過重", high:"🟡偏高", ok:"✅正常", zero:"—" };
              return (
                <Card key={prod.id} borderColor={hasIssue ? "#fecaca" : "#e2e8f0"}>
                  <div style={{fontSize:13,fontWeight:700,marginBottom:2}}>{prod.name}</div>
                  <div style={{fontSize:10,color:"#94a3b8",fontFamily:"monospace",marginBottom:12}}>{prod.id} ｜ 月均銷量 {Math.round(avg).toLocaleString()} 個</div>
                  {rows.map(({sh, stockAtArr, thisOrder, totalAtArr, daysOfStock, pressure, stockoutBefore}) => (
                    <div key={sh.id} style={{marginBottom:10,padding:"10px 12px",borderRadius:10,
                      background: sh.status==="past" ? "#f8fafc" : pressure==="over" ? "#fef2f2" : pressure==="high" ? "#fffbeb" : "#f0fdf4",
                      border: }}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
                        <span style={{fontSize:12,fontWeight:700,color:sh.status==="past"?"#94a3b8":"#475569"}}>{sh.label} 到港 {sh.arrivalDate}</span>
                        {sh.status !== "past" && (
                          <span style={{fontSize:11,fontWeight:700,color:pressureColors[pressure]}}>{pressureLabels[pressure]}</span>
                        )}
                        {sh.status === "past" && <span style={{fontSize:11,color:"#94a3b8"}}>已出貨</span>}
                      </div>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
                        <div>
                          <div style={{fontSize:9,color:"#94a3b8"}}>到港時剩餘庫存</div>
                          <div style={{fontSize:13,fontWeight:700,color:stockoutBefore?"#ef4444":"#1e293b"}}>{stockoutBefore ? "⚠️斷貨" : Math.round(stockAtArr).toLocaleString()}</div>
                        </div>
                        <div>
                          <div style={{fontSize:9,color:"#94a3b8"}}>本批採購量</div>
                          <div style={{fontSize:13,fontWeight:700,color:thisOrder>0?"#3b82f6":"#94a3b8"}}>{thisOrder>0?thisOrder.toLocaleString():"未採購"}</div>
                        </div>
                        <div>
                          <div style={{fontSize:9,color:"#94a3b8"}}>到港後總庫存天數</div>
                          <div style={{fontSize:13,fontWeight:700,color:sh.status==="past"?"#94a3b8":pressureColors[pressure]}}>{Math.round(daysOfStock)}天</div>
                        </div>
                      </div>
                      {sh.status !== "past" && (
                        <div style={{marginTop:8,background:"rgba(255,255,255,0.6)",borderRadius:8,height:6,overflow:"hidden"}}>
                          <div style={{height:"100%",borderRadius:8,
                            width:,
                            background:pressureColors[pressure]}}/>
                        </div>
                      )}
                    </div>
                  ))}
                </Card>
              );
            });
          })()}
        </>)}

        {/* ══ 參數 ══ */}
        {activeTab==="settings" && (<>
          <Card>
            <Label>安全庫存天數</Label>
            <div style={{fontSize:12,color:"#94a3b8",marginBottom:10}}>到港後希望最少還有幾天庫存</div>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              <input type="number" value={safetyDays} onChange={e=>setSafetyDays(Number(e.target.value))}
                style={{width:100,border:"1px solid #e2e8f0",borderRadius:8,padding:"10px 14px",fontSize:18,fontWeight:700}}/>
              <span style={{color:"#64748b"}}>天</span>
            </div>
          </Card>
          <Card>
            <Label>銷量計算參考月數</Label>
            <div style={{fontSize:12,color:"#94a3b8",marginBottom:10}}>用最近幾個月的銷量計算日均值</div>
            <select value={avgMonths} onChange={e=>setAvgMonths(Number(e.target.value))}
              style={{width:"100%",border:"1px solid #e2e8f0",borderRadius:8,padding:"10px 14px",fontSize:16}}>
              {[1,2,3,4,5,6].map(m=><option key={m} value={m}>{m} 個月</option>)}
            </select>
          </Card>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <Card><Label>運送天數</Label><BigNum size={24} color="#94a3b8">{TRANSIT_DAYS} 天</BigNum></Card>
            <Card><Label>貨櫃板數</Label><BigNum size={24} color="#94a3b8">{CONTAINER_PALLETS} 板</BigNum></Card>
          </div>
        </>)}

      </div>
    </div>
  );
}
