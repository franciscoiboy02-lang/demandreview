import { useState, useEffect, useCallback, useMemo, useRef } from "react";
import { LineChart, Line, BarChart, Bar, ComposedChart, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, ReferenceLine } from "recharts";
import * as XLSX from "xlsx";

// ── CONSTANTES ────────────────────────────────────────────────────────────
const MO  = ["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"];
const MF  = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"];
const FIGS = ["S&OP","POA 2026","Estadístico","LY 2025"];
const FIG_LABEL  = { "S&OP":"FCST", "POA 2026":"POA", "Estadístico":"EST.", "LY 2025":"LY 2025" };
const FIG_COLOR  = { "S&OP":"#d97706", "POA 2026":"#059669", "Estadístico":"#dc2626", "LY 2025":"#4f46e5" };
const UOMS       = ["FDO","QQI","MON"];
const UOM_LABEL  = { FDO:"Fardos", QQI:"Quintales", MON:"Monetario" };
const CELL_COLORS= { AX:"#059669",AY:"#34d399",AZ:"#fbbf24",BX:"#60a5fa",BY:"#93c5fd",BZ:"#fcd34d",CX:"#94a3b8",CY:"#cbd5e1",CZ:"#e2e8f0" };
const CELL_LABEL = { AX:"Estrella",AY:"Alto Valor",AZ:"Crítico",BX:"Estable",BY:"Gestión",BZ:"Atención",CX:"Menor",CY:"Revisión",CZ:"Candidato" };
const STATUS_COLOR={ pendiente:"#d97706", realizada:"#059669", cancelada:"#94a3b8" };
const MES_NOMBRE = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
const DIA_NOMBRE = ["Dom","Lun","Mar","Mié","Jue","Vie","Sáb"];
const MEETING_TYPES = ["Revisión de Demanda","Alineación Comercial","S&OP Mensual","Follow-up","Kick-off","Otro"];

// ── HELPERS ───────────────────────────────────────────────────────────────
const sumArr = (a=[]) => a.reduce((s,v)=>s+(+v||0),0);
const fmtK   = (n) => { if(n==null||isNaN(n)) return "-"; const x=Math.abs(n); return x>=1e6?`${(n/1e6).toFixed(1)}M`:x>=1000?`${(n/1000).toFixed(0)}K`:`${Math.round(n)}`; };
const pct    = (a,b) => b?Math.round((a-b)/b*100):0;
const clr    = (d) => d>0?"#059669":d<0?"#dc2626":"#94a3b8";
const arw    = (d) => d>0?"▲":d<0?"▼":"–";

// ── TEMA ──────────────────────────────────────────────────────────────────
const THEMES = {
  ejecutivo: { label:"Ejecutivo",  sidebar:"#1e3055", bg:"#f0f4f8", card:"#ffffff", border:"#dde5f0", borderLt:"#eef2f8", t1:"#1a2540", t2:"#5a6f88", t3:"#9aaabb", accent:"#1e3055", accentLt:"#e8edf6", amber:"#d97706", green:"#059669", red:"#dc2626", indigo:"#4f46e5" },
  dark:       { label:"Dark",       sidebar:"#080f1e", bg:"#0d1526", card:"#0f1c30", border:"#1e3050", borderLt:"#162540", t1:"#e2e8f0", t2:"#7a98bb", t3:"#3d5070", accent:"#3b82f6", accentLt:"#1a2d4a", amber:"#f59e0b", green:"#34d399", red:"#f87171", indigo:"#818cf8" },
  forest:     { label:"Forest",     sidebar:"#1a4a35", bg:"#f0f7f4", card:"#ffffff", border:"#c8e6d8", borderLt:"#e8f5ee", t1:"#1a3025", t2:"#4a7060", t3:"#8aaa99", accent:"#1a4a35", accentLt:"#e0f0e8", amber:"#d97706", green:"#059669", red:"#dc2626", indigo:"#4f46e5" },
  slate:      { label:"Slate",      sidebar:"#334155", bg:"#f1f5f9", card:"#ffffff", border:"#e2e8f0", borderLt:"#f1f5f9", t1:"#0f172a", t2:"#475569", t3:"#94a3b8", accent:"#334155", accentLt:"#e2e8f0", amber:"#d97706", green:"#059669", red:"#dc2626", indigo:"#4f46e5" },
};
function makeC(theme="ejecutivo", custom={}) { return { ...THEMES[theme]||THEMES.ejecutivo, ...custom }; }

// ── DEMO DATA ─────────────────────────────────────────────────────────────
const DEMO_SKUS = [
  { item:"60000001", mat:"ARROZ PREC GALLO DORADO 5U 5LB",  cat:"Arroz",  sub:"Arroz Precocido",       marca:"Gallo Dorado", seg:"Empacado", cls:"AX", base:21000 },
  { item:"60000015", mat:"HARINA MAIZ MI ESTRELLA 9U 800G", cat:"Harina", sub:"Harina de Maíz Blanco", marca:"Mi Estrella",  seg:"Empacado", cls:"AX", base:14000 },
  { item:"60000052", mat:"ARROZ BLCO MACARENA 25U 400G",    cat:"Arroz",  sub:"Arroz Blanco",          marca:"Macarena",     seg:"Empacado", cls:"AX", base:18000 },
  { item:"60000038", mat:"ARROZ BLCO MOLINERO 5U 5LB",      cat:"Arroz",  sub:"Arroz Blanco",          marca:"Molinero",     seg:"Empacado", cls:"AX", base:13000 },
  { item:"60000024", mat:"ARROZ PREC SUNRICE 25U 400G",     cat:"Arroz",  sub:"Arroz Precocido",       marca:"Sunrice",      seg:"Empacado", cls:"BX", base:9000  },
  { item:"60000094", mat:"HARINA MAIZ ORO MAYA 9U 800G",    cat:"Harina", sub:"Harina de Maíz Blanco", marca:"Oro Maya",     seg:"Empacado", cls:"AX", base:19000 },
  { item:"60000175", mat:"AVENA H CAMPO RICO 1U 50LB",      cat:"Avena",  sub:"Avena",                 marca:"Campo Rico",   seg:"Granel",   cls:"BX", base:4500  },
];
const DEMO_CLIENTS = {
  "MODERNO":       [{ cad:"1014602-UNISUPER, S.A. BODEGA CENTRAL", w:0.40 },
                    { cad:"1015276-OPERADORA DE TIENDAS S.A.",      w:0.38 },
                    { cad:"1013588-MEGARED DE SUPERMERCADOS, S.A.", w:0.22 }],
  "MARCA PRIVADA": [{ cad:"1015276-OPERADORA DE TIENDAS MP",  w:0.55 },
                    { cad:"1014602-UNISUPER MP",               w:0.45 }],
};
const FIG_BASE = {
  "S&OP":        [0,0,0,0,21000,25000,24000,22500,22500,22500,22500,25000],
  "POA 2026":    [18149,18346,20171,21409,22947,18890,17885,20528,17311,23564,21355,23361],
  "LY 2025":     [21921,25536,18540,15324,24126,20789,17550,18695,18573,20132,21261,19180],
  "Estadístico": [20341,20047,20778,20954,21130,21306,21482,21658,21834,22009,22185,22361],
};
const QQI_F = {"60000001":0.25,"60000015":0.1587,"60000052":0.22,"60000038":0.25,"60000024":0.22,"60000094":0.1587,"60000175":0.5};
const MON_P = {"60000001":180,"60000015":120,"60000052":155,"60000038":165,"60000024":160,"60000094":115,"60000175":95};

function buildDemo() {
  const rows=[];
  DEMO_SKUS.forEach(sku=>{ Object.entries(DEMO_CLIENTS).forEach(([canal,clients])=>{ const cF=canal==="MODERNO"?1:0.28; clients.forEach(client=>{ UOMS.forEach(uom=>{ const uF=uom==="FDO"?1:uom==="QQI"?(QQI_F[sku.item]||0.25):(MON_P[sku.item]||150); FIGS.forEach(fig=>{ const meses=FIG_BASE[fig].map(v=>parseFloat((v*sku.base/21000*cF*client.w*uF).toFixed(2))); const q1=meses[0]+meses[1]+meses[2],q2=meses[3]+meses[4]+meses[5],q3=meses[6]+meses[7]+meses[8],q4=meses[9]+meses[10]+meses[11]; rows.push({canal,key:`${canal}-${sku.item}`,comentario:"",modificacion:"",item:sku.item,cadena:client.cad,codCliente:"",categoria:sku.cat,subCategoria:sku.sub,marca:sku.marca,segmento:sku.seg,uom,figurasClave:fig,skuDesc:`${sku.item}-${sku.mat}`,material:sku.mat,clasificacion:sku.cls,meses,q1,q2,q3,q4,total:q1+q2+q3+q4}); }); }); }); }); });
  return rows;
}

// ── ESTILOS (función para que el tema sea reactivo) ───────────────────────
function makeS(C) { return {
  app:   { display:"flex", height:"100vh", background:C.bg, color:C.t1, fontFamily:"'Inter','Segoe UI',system-ui,sans-serif", overflow:"hidden" },
  side:  { width:204, background:C.sidebar, display:"flex", flexDirection:"column", flexShrink:0 },
  main:  { flex:1, overflow:"auto", padding:"22px 26px" },
  card:  { background:C.card, border:`1px solid ${C.border}`, borderRadius:10, padding:16, marginBottom:14, boxShadow:"0 1px 4px rgba(30,48,85,0.06)" },
  hdr:   { fontSize:10, color:C.t3, fontWeight:700, letterSpacing:"0.12em", textTransform:"uppercase", marginBottom:10 },
  tag:   (c)=>({ display:"inline-flex", padding:"2px 8px", borderRadius:10, fontSize:9, fontWeight:700, background:c+"22", color:c }),
  btn:   (prim,act)=>({ padding:"6px 14px", borderRadius:6, fontSize:11, fontWeight:600, cursor:"pointer", whiteSpace:"nowrap", border:prim?"none":act?`1px solid ${C.accent}`:`1px solid ${C.border}`, background:prim?C.accent:act?C.accentLt:C.card, color:prim?"#fff":act?C.accent:C.t2 }),
  btnSm: (prim)=>({ padding:"4px 10px", borderRadius:5, fontSize:10, fontWeight:600, cursor:"pointer", border:prim?"none":`1px solid ${C.border}`, background:prim?C.accent:C.card, color:prim?"#fff":C.t2 }),
  btnDng:{ padding:"5px 10px", borderRadius:5, fontSize:10, fontWeight:600, cursor:"pointer", border:"1px solid #fca5a5", background:"#fff5f5", color:"#dc2626" },
  sel:   { background:C.card, border:`1px solid ${C.border}`, borderRadius:6, color:C.t1, padding:"5px 8px", fontSize:11, cursor:"pointer" },
  inp:   { background:C.bg, border:`1px solid ${C.border}`, borderRadius:5, color:C.t1, padding:"3px 6px", fontSize:11, width:70, textAlign:"right", outline:"none" },
  inpTxt:{ flex:1, background:C.bg, border:`1px solid ${C.border}`, borderRadius:6, color:C.t1, padding:"7px 12px", fontSize:11, outline:"none" },
  nb:    (a)=>({ display:"flex", alignItems:"center", gap:9, padding:"9px 16px", cursor:"pointer", fontSize:12, color:a?"#fff":"rgba(255,255,255,0.5)", background:a?"rgba(255,255,255,0.13)":"transparent", borderLeft:a?"3px solid #f59e0b":"3px solid transparent" }),
  modal: { position:"fixed", inset:0, background:"rgba(10,20,50,0.6)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:900 },
  mp:    { background:C.card, border:`1px solid ${C.border}`, borderRadius:12, boxShadow:"0 24px 64px rgba(0,0,0,0.2)", width:"94%", maxWidth:1100, maxHeight:"92vh", overflow:"auto", padding:24 },
  dlg:   { background:C.card, border:`1px solid ${C.border}`, borderRadius:10, boxShadow:"0 12px 40px rgba(0,0,0,0.15)", width:460, padding:24 },
  ttip:  { background:C.card, border:`1px solid ${C.border}`, borderRadius:8, fontSize:12, color:C.t1, boxShadow:"0 4px 12px rgba(0,0,0,0.1)", padding:"8px 12px" },
};}

// ── PYTHON MODELS ─────────────────────────────────────────────────────────
const DEF_MODELS = [
  { id:"sarima", name:"SARIMA", active:true, mape:70.1, description:"Seasonal ARIMA — patrones estacionales fuertes",
    script:`from statsmodels.tsa.statespace.sarimax import SARIMAX\n\ndef run_sarima(series, periods=12):\n    m = SARIMAX(series, order=(1,1,1), seasonal_order=(1,1,1,12),\n                enforce_stationarity=False, enforce_invertibility=False)\n    r = m.fit(disp=False)\n    fc = r.forecast(steps=periods)\n    mape = abs((series[-6:] - r.fittedvalues[-6:]) / series[-6:]).mean() * 100\n    return {"forecast": fc.tolist(), "mape_val": round(mape,2)}\n` },
  { id:"holtwinters", name:"Holt-Winters (ETS)", active:true, mape:73.5, description:"Suavizamiento exponencial triple amortiguado",
    script:`from statsmodels.tsa.holtwinters import ExponentialSmoothing\n\ndef run_hw(series, periods=12):\n    m = ExponentialSmoothing(series, seasonal_periods=12, trend="add", seasonal="add", damped_trend=True)\n    r = m.fit(optimized=True)\n    mape = abs((series[-6:] - r.fittedvalues[-6:]) / series[-6:]).mean() * 100\n    return {"forecast": r.forecast(periods).tolist(), "mape_val": round(mape,2)}\n` },
  { id:"prophet", name:"Prophet ★ MEJOR", active:true, mape:53.5, description:"MAPE 6m: 53.5% — estacionalidad múltiple",
    script:`from prophet import Prophet\nimport pandas as pd\n\ndef run_prophet(series, periods=12, start="2023-01-01"):\n    df = pd.DataFrame({"ds": pd.date_range(start=start, periods=len(series), freq="MS"), "y": series.values})\n    m = Prophet(seasonality_mode="multiplicative", yearly_seasonality=True,\n                weekly_seasonality=False, daily_seasonality=False, changepoint_prior_scale=0.05)\n    m.fit(df)\n    future = m.make_future_dataframe(periods=periods, freq="MS")\n    fc = m.predict(future).tail(periods)["yhat"].values\n    return {"forecast": fc.tolist()}\n` },
  { id:"ensemble", name:"Ensemble", active:true, mape:null, description:"Ponderado por inverso del MAPE",
    script:`import numpy as np\n\ndef run_ensemble(forecasts, mapes):\n    w = {m: 1.0/max(v,0.01) for m,v in mapes.items()}\n    tot = sum(w.values())\n    w = {m: x/tot for m,x in w.items()}\n    result = np.zeros(len(next(iter(forecasts.values()))))\n    for m,fc in forecasts.items(): result += w[m]*np.array(fc)\n    return {"forecast": result.tolist(), "weights": w}\n` },
];

// ── SIDEBAR ───────────────────────────────────────────────────────────────
function Sidebar({ view, setView, onImport, onExport, cfg, C, S }) {
  const fRef = useRef();
  const items = [
    {id:"dashboard", icon:"▦", lbl:"Dashboard"},
    {id:"revision",  icon:"⟳", lbl:"Revisión"},
    {id:"impacto",   icon:"⇄", lbl:"Impacto"},
    {id:"clientes",  icon:"⊞", lbl:"Clientes"},
    {id:"abcxyz",    icon:"◉", lbl:"ABC-XYZ"},
    {id:"gaps",      icon:"◈", lbl:"Gaps"},
    {id:"materia",   icon:"⬡", lbl:"Materia Prima"},
    {id:"reuniones", icon:"📅", lbl:"Reuniones"},
    {id:"modelos",   icon:"⌥", lbl:"Modelos"},
    {id:"historial", icon:"◷", lbl:"Historial"},
    {id:"tutorial",  icon:"📖", lbl:"Tutorial"},
  ];
  const sideBtn = { padding:"7px 14px", borderRadius:6, fontSize:10, fontWeight:600, cursor:"pointer", border:"1px solid rgba(255,255,255,0.2)", background:"transparent", color:"rgba(255,255,255,0.7)" };
  return (
    <div style={S.side}>
      <div style={{ padding:"14px 14px 12px", borderBottom:"1px solid rgba(255,255,255,0.1)", display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
        <div>
          {cfg.logoEmpresa
            ? <img src={cfg.logoEmpresa} alt="logo" style={{ height:28, objectFit:"contain", maxWidth:130, display:"block" }}/>
            : <div style={{ fontSize:14, fontWeight:800, color:"#fff", letterSpacing:"0.12em" }}>ALCSA</div>
          }
          <div style={{ fontSize:9, color:"rgba(255,255,255,0.35)", letterSpacing:"0.2em", marginTop:2 }}>DEMAND REVIEW</div>
        </div>
        <button onClick={()=>setView("config")} title="Configuración"
          style={{ background:view==="config"?"rgba(255,255,255,0.2)":"transparent", border:`1px solid rgba(255,255,255,${view==="config"?0.4:0.15})`, borderRadius:6, color:"rgba(255,255,255,0.65)", cursor:"pointer", fontSize:14, padding:"4px 7px", lineHeight:1 }}>
          ⚙
        </button>
      </div>
      <div style={{ flex:1, display:"flex", flexDirection:"column", gap:1, padding:"8px 0", overflowY:"auto" }}>
        {items.map(i=>(
          <div key={i.id} style={S.nb(view===i.id)} onClick={()=>setView(i.id)}>
            <span style={{ fontSize:13 }}>{i.icon}</span>
            <span>{i.lbl}</span>
          </div>
        ))}
      </div>
      <div style={{ padding:"12px 14px", borderTop:"1px solid rgba(255,255,255,0.1)", display:"flex", flexDirection:"column", gap:7 }}>
        <input type="file" ref={fRef} accept=".xlsx,.xlsm,.xls" style={{ display:"none" }} onChange={e=>{ onImport(e.target.files[0]); e.target.value=""; }} />
        <button style={sideBtn} onClick={()=>fRef.current.click()}>⬆ Importar Excel</button>
        <button style={sideBtn} onClick={onExport}>⬇ Exportar Excel</button>
      </div>
      <div style={{ padding:"5px 14px 10px", fontSize:9, color:"rgba(255,255,255,0.18)" }}>v2.5 · 2026</div>
    </div>
  );
}

// ── KPI CARD ──────────────────────────────────────────────────────────────
function KpiCard({ fig, value, uom, cmpValue, cmpLabel, C }) {
  const c=FIG_COLOR[fig], d=cmpValue!=null&&cmpValue>0?pct(value,cmpValue):null;
  return (
    <div style={{ background:C.card, border:`1px solid ${C.border}`, borderTop:`3px solid ${c}`, borderRadius:8, padding:"16px 20px", flex:1, minWidth:0, boxShadow:"0 1px 4px rgba(30,48,85,0.06)" }}>
      <div style={{ fontSize:9, color:C.t3, letterSpacing:"0.14em", textTransform:"uppercase", marginBottom:10, fontWeight:700 }}>{FIG_LABEL[fig]}</div>
      <div style={{ fontSize:30, fontWeight:800, color:C.t1, letterSpacing:"-0.02em", lineHeight:1 }}>{fmtK(value)}</div>
      <div style={{ fontSize:11, color:C.t3, marginTop:6 }}>{UOM_LABEL[uom]}</div>
      {d!=null&&<div style={{ marginTop:8, fontSize:11, color:clr(d), fontWeight:600 }}>{arw(d)} {Math.abs(d)}% vs {cmpLabel}</div>}
    </div>
  );
}

// ── BREAKDOWN TABLE ───────────────────────────────────────────────────────
function BreakdownTable({ title, groupKey, filtered, C, S }) {
  const groups = useMemo(()=>{
    const m={};
    filtered.forEach(r=>{ const gk=r[groupKey]||"—"; if(!m[gk]) m[gk]={k:gk,S:0,P:0,E:0,L:0}; const v=sumArr(r.meses); if(r.figurasClave==="S&OP") m[gk].S+=v; if(r.figurasClave==="POA 2026") m[gk].P+=v; if(r.figurasClave==="Estadístico") m[gk].E+=v; if(r.figurasClave==="LY 2025") m[gk].L+=v; });
    return Object.values(m).sort((a,b)=>b.S-a.S);
  },[filtered,groupKey]);
  const totS=groups.reduce((s,g)=>s+g.S,0),totP=groups.reduce((s,g)=>s+g.P,0),totE=groups.reduce((s,g)=>s+g.E,0),totL=groups.reduce((s,g)=>s+g.L,0);
  const thStyle={padding:"7px 12px",color:C.t2,fontSize:9,fontWeight:700,letterSpacing:"0.1em",borderBottom:`1px solid ${C.border}`};
  return (
    <div style={S.card}>
      <div style={S.hdr}>{title}</div>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
        <thead><tr><th style={{...thStyle,textAlign:"left"}}>{groupKey==="categoria"?"CATEGORÍA":"MARCA"}</th><th style={{...thStyle,textAlign:"right"}}>FCST</th><th style={{...thStyle,textAlign:"right"}}>POA</th><th style={{...thStyle,textAlign:"right"}}>EST.</th><th style={{...thStyle,textAlign:"right"}}>LY 2025</th><th style={{...thStyle,textAlign:"right"}}>GAP ABS</th><th style={{...thStyle,textAlign:"right"}}>GAP %</th></tr></thead>
        <tbody>
          {groups.map((g,i)=>{ const gap=g.S-g.P,gp=pct(g.S,g.P); return (
            <tr key={g.k} style={{borderBottom:`1px solid ${C.borderLt}`,background:i%2?C.bg:C.card}}>
              <td style={{padding:"7px 12px",color:C.t1,fontWeight:600}}>{g.k}</td>
              <td style={{padding:"7px 12px",textAlign:"right",color:FIG_COLOR["S&OP"],fontWeight:700}}>{fmtK(g.S)}</td>
              <td style={{padding:"7px 12px",textAlign:"right",color:FIG_COLOR["POA 2026"]}}>{fmtK(g.P)}</td>
              <td style={{padding:"7px 12px",textAlign:"right",color:FIG_COLOR["Estadístico"]}}>{fmtK(g.E)}</td>
              <td style={{padding:"7px 12px",textAlign:"right",color:FIG_COLOR["LY 2025"]}}>{fmtK(g.L)}</td>
              <td style={{padding:"7px 12px",textAlign:"right",color:clr(gap),fontWeight:700}}>{gap>0?"+":""}{fmtK(gap)}</td>
              <td style={{padding:"7px 12px",textAlign:"right"}}><span style={S.tag(clr(gp))}>{gp>0?"+":""}{gp}%</span></td>
            </tr>
          );})}
        </tbody>
        <tfoot><tr style={{borderTop:`2px solid ${C.border}`}}>
          <td style={{padding:"7px 12px",color:C.t1,fontWeight:700,fontSize:10}}>TOTAL</td>
          <td style={{padding:"7px 12px",textAlign:"right",color:FIG_COLOR["S&OP"],fontWeight:700}}>{fmtK(totS)}</td>
          <td style={{padding:"7px 12px",textAlign:"right",color:FIG_COLOR["POA 2026"],fontWeight:700}}>{fmtK(totP)}</td>
          <td style={{padding:"7px 12px",textAlign:"right",color:FIG_COLOR["Estadístico"],fontWeight:700}}>{fmtK(totE)}</td>
          <td style={{padding:"7px 12px",textAlign:"right",color:FIG_COLOR["LY 2025"],fontWeight:700}}>{fmtK(totL)}</td>
          <td style={{padding:"7px 12px",textAlign:"right",color:clr(totS-totP),fontWeight:700}}>{totS-totP>0?"+":""}{fmtK(totS-totP)}</td>
          <td style={{padding:"7px 12px",textAlign:"right"}}><span style={S.tag(clr(pct(totS,totP)))}>{pct(totS,totP)>0?"+":""}{pct(totS,totP)}%</span></td>
        </tr></tfoot>
      </table>
    </div>
  );
}

// ── DASHBOARD ─────────────────────────────────────────────────────────────
function DashboardView({ rows, uom, setUom, canal, setCanal, canales, C, S }) {
  const filtered=useMemo(()=>rows.filter(r=>r.uom===uom&&(canal==="ALL"||r.canal===canal)),[rows,uom,canal]);
  const totals=useMemo(()=>{ const t={}; FIGS.forEach(f=>{t[f]=0;}); filtered.forEach(r=>{t[r.figurasClave]=(t[r.figurasClave]||0)+sumArr(r.meses);}); return t; },[filtered]);
  const byMonth=useMemo(()=>MO.map((m,i)=>{ const p={name:m}; FIGS.forEach(f=>{p[f]=filtered.filter(r=>r.figurasClave===f).reduce((s,r)=>s+(r.meses[i]||0),0);}); return p; }),[filtered]);
  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",marginBottom:20}}>
        <div><div style={{fontSize:18,fontWeight:700,color:C.t1}}>Resumen Anual 2026</div><div style={{fontSize:11,color:C.t2,marginTop:3}}>Canal: <strong style={{color:C.t1}}>{canal==="ALL"?"Todos":canal}</strong> · UoM: <strong style={{color:uom==="FDO"?C.amber:uom==="QQI"?C.green:C.indigo}}>{UOM_LABEL[uom]}</strong></div></div>
        <div style={{display:"flex",gap:12,alignItems:"center"}}>
          <div style={{display:"flex",alignItems:"center",gap:6}}><span style={{fontSize:11,color:C.t2}}>Canal</span><select style={S.sel} value={canal} onChange={e=>setCanal(e.target.value)}>{canales.map(c=><option key={c} value={c}>{c==="ALL"?"Todos":c}</option>)}</select></div>
          <div style={{display:"flex",gap:3}}>{UOMS.map(u=><button key={u} style={S.btn(false,uom===u)} onClick={()=>setUom(u)}>{u}</button>)}</div>
        </div>
      </div>
      <div style={{display:"flex",gap:12,marginBottom:16}}>
        <KpiCard fig="S&OP"        value={totals["S&OP"]||0}        uom={uom} cmpValue={totals["POA 2026"]||0} cmpLabel="POA" C={C}/>
        <KpiCard fig="POA 2026"    value={totals["POA 2026"]||0}    uom={uom} C={C}/>
        <KpiCard fig="Estadístico" value={totals["Estadístico"]||0} uom={uom} cmpValue={totals["S&OP"]||0}     cmpLabel="FCST" C={C}/>
        <KpiCard fig="LY 2025"     value={totals["LY 2025"]||0}     uom={uom} C={C}/>
      </div>
      <div style={S.card}>
        <div style={S.hdr}>Evolución mensual 2026</div>
        <ResponsiveContainer width="100%" height={195}>
          <ComposedChart data={byMonth} margin={{top:4,right:8,left:0,bottom:0}}>
            <CartesianGrid stroke={C.borderLt} strokeDasharray="4 4"/>
            <XAxis dataKey="name" tick={{fill:C.t2,fontSize:11,fontWeight:600}}/>
            <YAxis tick={{fill:C.t2,fontSize:11}} tickFormatter={fmtK}/>
            <Tooltip contentStyle={S.ttip} formatter={(v,n)=>[fmtK(v),FIG_LABEL[n]||n]}/>
            {FIGS.map(f=><Line key={f} type="monotone" dataKey={f} stroke={FIG_COLOR[f]} strokeWidth={f==="S&OP"?2.5:1.5} strokeDasharray={f==="POA 2026"?"5 3":f==="LY 2025"?"2 4":f==="Estadístico"?"8 2":"none"} dot={false}/>)}
          </ComposedChart>
        </ResponsiveContainer>
        <div style={{display:"flex",gap:20,justifyContent:"center",marginTop:10}}>
          {FIGS.map(f=><span key={f} style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:C.t2}}><span style={{width:18,height:2,background:FIG_COLOR[f],display:"inline-block",borderRadius:1}}/>{FIG_LABEL[f]}</span>)}
        </div>
      </div>
      <BreakdownTable title="Desglose por Categoría" groupKey="categoria" filtered={filtered} C={C} S={S}/>
      <BreakdownTable title="Desglose por Marca"     groupKey="marca"     filtered={filtered} C={C} S={S}/>
    </div>
  );
}

// ── REVISION VIEW ─────────────────────────────────────────────────────────
function RevisionView({ rows, uom, canal, onSKUClick, C, S }) {
  const [fCanal,setFCanal]=useState(canal==="ALL"?"ALL":canal);
  const [fMarca,setFMarca]=useState("ALL");
  const [fCat,setFCat]=useState("ALL");
  const [fSub,setFSub]=useState("ALL");
  const [fItem,setFItem]=useState("ALL");
  const [fCadena,setFCadena]=useState("ALL");
  const opt=(arr)=>["ALL",...new Set(arr)];
  const filtered=useMemo(()=>rows.filter(r=>r.uom===uom&&(fCanal==="ALL"||r.canal===fCanal)&&(fMarca==="ALL"||r.marca===fMarca)&&(fCat==="ALL"||r.categoria===fCat)&&(fSub==="ALL"||r.subCategoria===fSub)&&(fItem==="ALL"||r.item===fItem)&&(fCadena==="ALL"||r.cadena===fCadena)),[rows,uom,fCanal,fMarca,fCat,fSub,fItem,fCadena]);
  const byItem=useMemo(()=>{ const m={}; filtered.forEach(r=>{ const k=`${r.canal}||${r.item}`; if(!m[k]) m[k]={key:k,canal:r.canal,item:r.item,mat:r.material,cls:r.clasificacion,figs:{}}; if(!m[k].figs[r.figurasClave]) m[k].figs[r.figurasClave]=Array(12).fill(0); r.meses.forEach((v,i)=>{m[k].figs[r.figurasClave][i]+=v;}); }); return Object.values(m).sort((a,b)=>sumArr(b.figs["S&OP"])-sumArr(a.figs["S&OP"])); },[filtered]);
  const selOpts=[[fCanal,setFCanal,opt(rows.map(r=>r.canal)),"Canal"],[fCat,setFCat,opt(rows.map(r=>r.categoria)),"Cat."],[fSub,setFSub,opt(rows.map(r=>r.subCategoria)),"Sub Cat."],[fMarca,setFMarca,opt(rows.map(r=>r.marca)),"Marca"],[fItem,setFItem,opt(rows.map(r=>r.item)),"SKU"],[fCadena,setFCadena,opt(rows.map(r=>r.cadena).filter(Boolean)),"Cliente"]];
  const thS={padding:"7px 6px",color:C.t2,fontWeight:700,fontSize:9,whiteSpace:"nowrap",letterSpacing:"0.06em",borderBottom:`2px solid ${C.border}`};
  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14,gap:12}}>
        <div style={{flexShrink:0}}><div style={{fontSize:17,fontWeight:700,color:C.t1}}>Revisión de Demanda</div><div style={{fontSize:11,color:C.t2}}>Clic en fila <strong>FCST</strong> para editar · UoM: <strong style={{color:C.amber}}>{UOM_LABEL[uom]}</strong></div></div>
        <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap",justifyContent:"flex-end"}}>
          {selOpts.map(([val,setter,opts,lbl])=>(
            <div key={lbl} style={{display:"flex",alignItems:"center",gap:4}}>
              <span style={{fontSize:9,color:C.t3,fontWeight:600}}>{lbl}</span>
              <select style={{...S.sel,maxWidth:150}} value={val} onChange={e=>setter(e.target.value)}>
                {opts.map(o=><option key={o} value={o}>{o==="ALL"?"Todos":o.length>22?o.slice(0,22)+"…":o}</option>)}
              </select>
            </div>
          ))}
        </div>
      </div>
      <div style={{...S.card,padding:0,overflow:"hidden"}}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:10}}>
            <thead><tr style={{background:C.bg}}>
              {["Canal","Item","Material","Cls","Figura",...MO,"TOTAL","vs POA","vs LY"].map((h,hi)=>(
                <th key={hi} style={{...thS,textAlign:hi<5?"left":"right",paddingLeft:hi===0?16:undefined}}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {byItem.map(d=>(
                FIGS.map((fig,fi)=>{
                  const vals=d.figs[fig]||Array(12).fill(0),tot=sumArr(vals);
                  const poa=sumArr(d.figs["POA 2026"]||[]),ly=sumArr(d.figs["LY 2025"]||[]);
                  const first=fi===0;
                  return (
                    <tr key={`${d.key}-${fig}`} style={{borderBottom:fi===3?`1px solid ${C.border}`:`1px solid ${C.borderLt}`,background:fig==="S&OP"?C.card:C.bg,cursor:fig==="S&OP"?"pointer":"default"}} onClick={fig==="S&OP"?()=>onSKUClick(d.key):undefined}>
                      <td style={{padding:"4px 6px 4px 16px",color:C.t2,whiteSpace:"nowrap"}}>{first?d.canal:""}</td>
                      <td style={{padding:"4px 6px",color:C.accent,fontWeight:700}}>{first?d.item:""}</td>
                      <td style={{padding:"4px 6px",color:C.t2,maxWidth:150,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{first?d.mat:""}</td>
                      <td style={{padding:"4px 6px"}}>{first?<span style={S.tag(C.t2)}>{d.cls}</span>:""}</td>
                      <td style={{padding:"4px 6px"}}><span style={S.tag(FIG_COLOR[fig])}>{FIG_LABEL[fig]}</span></td>
                      {vals.map((v,i)=><td key={i} style={{padding:"3px 6px",textAlign:"right",color:fig==="S&OP"?C.t1:C.t3,fontWeight:fig==="S&OP"?600:400}}>{fmtK(v)}</td>)}
                      <td style={{padding:"3px 6px",textAlign:"right",fontWeight:700,color:C.t1}}>{fmtK(tot)}</td>
                      <td style={{padding:"3px 6px",textAlign:"right"}}>{first&&fig==="S&OP"?<span style={{color:clr(pct(tot,poa)),fontWeight:600}}>{pct(tot,poa)>0?"+":""}{pct(tot,poa)}%</span>:""}</td>
                      <td style={{padding:"3px 6px",textAlign:"right"}}>{first&&fig==="S&OP"?<span style={{color:clr(pct(tot,ly)),fontWeight:600}}>{pct(tot,ly)>0?"+":""}{pct(tot,ly)}%</span>:""}</td>
                    </tr>
                  );
                })
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ── SKU MODAL ─────────────────────────────────────────────────────────────
function SKUModal({ skuKey, rows, uom, onChange, onClose, editableFrom, C, S }) {
  const [canal,item]=skuKey.split("||");
  const skuRows=useMemo(()=>rows.filter(r=>r.canal===canal&&r.item===item&&r.uom===uom),[rows,canal,item,uom]);
  const meta=skuRows[0]||{};
  const figs=useMemo(()=>{ const m={}; skuRows.forEach(r=>{ if(!m[r.figurasClave]) m[r.figurasClave]=Array(12).fill(0); r.meses.forEach((v,i)=>{m[r.figurasClave][i]+=v;}); }); return m; },[skuRows]);
  const [localSop,setLocalSop]=useState(()=>[...(figs["S&OP"]||Array(12).fill(0))]);
  const [comment,setComment]=useState(""); const [error,setError]=useState("");
  const poa=figs["POA 2026"]||Array(12).fill(0);
  const gap=localSop.map((v,i)=>({m:MO[i],gap:v-(poa[i]||0),pct:pct(v,poa[i])}));
  const hist=useMemo(()=>{ const base=sumArr(figs["S&OP"]||[])/12||500; return Array.from({length:24},(_,i)=>{ const yr=i<12?2024:2025,mi=i%12; return {label:`${MO[mi]}-${String(yr).slice(2)}`,real:Math.round(base*(0.85+Math.sin(i/2)*0.15)),fc:Math.round(base*(0.9+Math.sin(i/2+0.3)*0.1))}; }); },[figs]);
  const handleSave=()=>{ if(!comment.trim()){setError("El comentario es obligatorio antes de guardar.");return;} const original=figs["S&OP"]||Array(12).fill(0); localSop.forEach((v,i)=>{ if(i>=editableFrom&&Math.abs(v-(original[i]||0))>0.5) onChange(skuKey,i,v,comment,uom); }); onClose(); };
  const lS={fontSize:9,color:C.t2,fontWeight:700,letterSpacing:"0.1em",textTransform:"uppercase",borderBottom:`1px solid ${C.border}`,padding:"6px 8px",textAlign:"right",whiteSpace:"nowrap"};
  return (
    <div style={S.modal} onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={S.mp}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:16}}>
          <div>
            <div style={{fontSize:10,color:C.t3,marginBottom:4}}>{canal} · <strong style={{color:C.amber}}>{UOM_LABEL[uom]}</strong> · Editables: <strong style={{color:C.green}}>{MO.slice(editableFrom).join(", ")}</strong></div>
            <div style={{fontSize:16,fontWeight:700,color:C.t1}}>{item}</div>
            <div style={{fontSize:12,color:C.t2,marginTop:2}}>{meta.material}</div>
            <div style={{marginTop:8,display:"flex",gap:5,flexWrap:"wrap"}}>
              {[meta.categoria,meta.subCategoria,meta.marca,meta.clasificacion].filter(Boolean).map(t=><span key={t} style={{fontSize:10,color:C.t2,background:C.accentLt,padding:"2px 8px",borderRadius:12,fontWeight:500}}>{t}</span>)}
            </div>
          </div>
          <button onClick={onClose} style={{...S.btnSm(false),fontSize:16,padding:"4px 10px"}}>✕</button>
        </div>
        <div style={{overflowX:"auto",marginBottom:16,border:`1px solid ${C.border}`,borderRadius:8,overflow:"hidden"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:10}}>
            <thead><tr style={{background:C.bg}}>
              <th style={{...lS,textAlign:"left",paddingLeft:12}}>FIGURA</th>
              {MO.map((m,i)=><th key={m} style={{...lS,color:i>=editableFrom?C.t1:C.t3,fontWeight:i>=editableFrom?700:500}}>{m}</th>)}
              <th style={lS}>TOTAL</th>
            </tr></thead>
            <tbody>
              {FIGS.map(fig=>{ const vals=fig==="S&OP"?localSop:(figs[fig]||Array(12).fill(0)); return (
                <tr key={fig} style={{borderTop:`1px solid ${C.borderLt}`}}>
                  <td style={{padding:"5px 8px 5px 12px"}}><span style={S.tag(FIG_COLOR[fig])}>{FIG_LABEL[fig]}</span></td>
                  {vals.map((v,i)=>(
                    <td key={i} style={{padding:"2px 3px",textAlign:"right",background:i>=editableFrom&&fig==="S&OP"?"#fffbf0":"transparent"}}>
                      {fig==="S&OP"?<input type="number" style={{...S.inp,color:i<editableFrom?C.t3:C.t1,background:i<editableFrom?C.bg:"#fffbf0",borderColor:i<editableFrom?C.borderLt:C.amber+"55",fontWeight:i>=editableFrom?600:400}} value={Math.round(v)} disabled={i<editableFrom} onChange={e=>{setLocalSop(p=>p.map((x,idx)=>idx===i?+e.target.value:x));setError("");}}/>:<span style={{color:i>=editableFrom?FIG_COLOR[fig]:C.t3,opacity:i<editableFrom?0.5:1}}>{fmtK(v)}</span>}
                    </td>
                  ))}
                  <td style={{padding:"2px 8px",textAlign:"right",fontWeight:700,color:FIG_COLOR[fig]}}>{fmtK(sumArr(vals))}</td>
                </tr>
              );})}
              <tr style={{borderTop:`2px solid ${C.border}`,background:C.bg}}>
                <td style={{padding:"5px 8px 5px 12px",fontSize:9,color:C.t2,fontWeight:700}}>GAP FCST-POA</td>
                {gap.map((g,i)=><td key={i} style={{padding:"2px 3px",textAlign:"right",fontSize:10,color:clr(g.gap),opacity:i<editableFrom?0.4:1,fontWeight:600}}>{g.gap>0?"+":""}{fmtK(g.gap)}</td>)}
                <td style={{padding:"2px 8px",textAlign:"right",fontSize:10,color:clr(sumArr(gap.map(g=>g.gap))),fontWeight:700}}>{sumArr(gap.map(g=>g.gap))>0?"+":""}{fmtK(sumArr(gap.map(g=>g.gap)))}</td>
              </tr>
            </tbody>
          </table>
        </div>
        <div style={{background:error?"#fff8f8":C.bg,border:`1px solid ${error?"#fca5a5":C.border}`,borderRadius:8,padding:"12px 14px",marginBottom:14}}>
          <div style={{fontSize:10,color:error?"#dc2626":C.t2,fontWeight:600,marginBottom:8}}>{error?`⚠ ${error}`:"💬 Comentario del ajuste (obligatorio)"}</div>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            <input type="text" placeholder="Razón del ajuste: input comercial, promoción, temporada..." value={comment} onChange={e=>{setComment(e.target.value);setError("");}} style={{...S.inpTxt,borderColor:error?"#fca5a5":C.border}}/>
            <button style={S.btn(true)} onClick={handleSave}>Guardar y Cerrar</button>
          </div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"3fr 1fr",gap:12}}>
          <div>
            <div style={S.hdr}>Real vs Estadístico — 24 meses</div>
            <ResponsiveContainer width="100%" height={150}>
              <LineChart data={hist} margin={{top:0,right:8,left:0,bottom:0}}>
                <CartesianGrid stroke={C.borderLt} strokeDasharray="3 3"/>
                <XAxis dataKey="label" tick={{fill:C.t2,fontSize:9}} interval={3}/>
                <YAxis tick={{fill:C.t2,fontSize:9}} tickFormatter={fmtK}/>
                <Tooltip contentStyle={S.ttip} formatter={(v,n)=>[fmtK(v),n]}/>
                <Line type="monotone" dataKey="real" stroke="#3b82f6" strokeWidth={2} dot={false} name="Real"/>
                <Line type="monotone" dataKey="fc"   stroke={C.red}  strokeWidth={1.5} strokeDasharray="4 2" dot={false} name="Estadístico"/>
              </LineChart>
            </ResponsiveContainer>
          </div>
          <div>
            <div style={S.hdr}>Gap % vs POA</div>
            <ResponsiveContainer width="100%" height={150}>
              <BarChart data={gap.slice(editableFrom)} margin={{top:0,right:5,left:0,bottom:0}}>
                <CartesianGrid stroke={C.borderLt} horizontal={false}/>
                <XAxis dataKey="m" tick={{fill:C.t2,fontSize:9}}/>
                <YAxis tick={{fill:C.t2,fontSize:9}} unit="%"/>
                <Tooltip contentStyle={S.ttip} formatter={v=>[v+"%","Gap"]}/>
                <Bar dataKey="pct" fill={C.amber} radius={[3,3,0,0]} isAnimationActive={false}/>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── IMPACTO VIEW ──────────────────────────────────────────────────────────
function ImpactoView({ rows, log, sessions, uom, C, S }) {
  const [selSess,setSelSess]=useState(sessions[sessions.length-1]?.id||"");
  const [umbral,setUmbral]=useState(10);
  const [fCanal,setFCanal]=useState("ALL");
  const canalesOpts=useMemo(()=>["ALL",...new Set(rows.map(r=>r.canal))]  ,[rows]);
  const sessLog=useMemo(()=>log.filter(c=>c.session===selSess&&c.uom===uom&&(fCanal==="ALL"||c.canal===fCanal)),[log,selSess,uom,fCanal]);
  const changes=useMemo(()=>{ const m={}; sessLog.forEach(c=>{ const k=`${c.canal}||${c.item}`; if(!m[k]) m[k]={item:c.item,mat:c.material,canal:c.canal,deltas:Array(12).fill(0)}; if(c.monthIdx!=null) m[k].deltas[c.monthIdx]+=(c.delta||0); }); return Object.values(m); },[sessLog]);
  const comparison=useMemo(()=>{ return changes.map(ch=>{ const sopRows=rows.filter(r=>r.canal===ch.canal&&r.item===ch.item&&r.uom===uom&&r.figurasClave==="S&OP"); const after=Array(12).fill(0); sopRows.forEach(r=>r.meses.forEach((v,i)=>{after[i]+=v;})); const before=after.map((v,i)=>v-ch.deltas[i]); const totAfter=sumArr(after),totBefore=sumArr(before),deltaTotal=totAfter-totBefore,deltaPct=pct(totAfter,totBefore),significant=Math.abs(deltaPct)>=umbral; return {...ch,after,before,totAfter,totBefore,deltaTotal,deltaPct,significant}; }).sort((a,b)=>Math.abs(b.deltaTotal)-Math.abs(a.deltaTotal)); },[changes,rows,uom,umbral]);
  const totalBefore=comparison.reduce((s,d)=>s+d.totBefore,0),totalAfter=comparison.reduce((s,d)=>s+d.totAfter,0);
  const totalDelta=totalAfter-totalBefore,significant=comparison.filter(d=>d.significant);
  const monthChart=useMemo(()=>MO.map((m,i)=>({name:m,Antes:Math.round(comparison.reduce((s,d)=>s+d.before[i],0)),Después:Math.round(comparison.reduce((s,d)=>s+d.after[i],0))})),[comparison]);
  const sess=sessions.find(s=>s.id===selSess);
  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",marginBottom:18}}>
        <div><div style={{fontSize:17,fontWeight:700,color:C.t1}}>Impacto de Revisión</div><div style={{fontSize:11,color:C.t2,marginTop:2}}>S&OP antes vs después de los ajustes por sesión</div></div>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          <select style={{...S.sel,maxWidth:220}} value={selSess} onChange={e=>setSelSess(e.target.value)}>{sessions.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}</select>
          <select style={S.sel} value={fCanal} onChange={e=>setFCanal(e.target.value)}>{canalesOpts.map(o=><option key={o} value={o}>{o==="ALL"?"Todos":o}</option>)}</select>
          <select style={S.sel} value={umbral} onChange={e=>setUmbral(+e.target.value)}>{[5,10,15,20,25,30].map(v=><option key={v} value={v}>≥{v}%</option>)}</select>
        </div>
      </div>
      {comparison.length===0?(
        <div style={{...S.card,textAlign:"center",padding:"56px 0",color:C.t3}}><div style={{fontSize:32,marginBottom:12}}>📋</div><div style={{fontSize:13,fontWeight:600,color:C.t2,marginBottom:4}}>Sin cambios registrados</div><div style={{fontSize:11}}>Esta sesión no tiene ajustes guardados aún.</div></div>
      ):(
        <>
          <div style={{display:"flex",gap:12,marginBottom:16}}>
            {[["S&OP Antes",fmtK(totalBefore),"Punto de partida",C.indigo],["S&OP Después",fmtK(totalAfter),"Luego de ajustes",C.amber],[(totalDelta>0?"+":"")+fmtK(totalDelta),`${pct(totalAfter,totalBefore)>0?"+":""}${pct(totalAfter,totalBefore)}%`,"Δ Neto",clr(totalDelta)],[String(significant.length),`de ${comparison.length} SKUs`,"Significativos",C.red]].map(([val,sub,lbl,c],i)=>(
              <div key={i} style={{flex:1,background:C.card,border:`1px solid ${C.border}`,borderTop:`3px solid ${c}`,borderRadius:8,padding:"14px 18px",boxShadow:"0 1px 4px rgba(30,48,85,0.06)"}}>
                <div style={{fontSize:9,color:C.t3,letterSpacing:"0.12em",textTransform:"uppercase",marginBottom:8,fontWeight:700}}>{lbl}</div>
                <div style={{fontSize:26,fontWeight:800,color:C.t1,lineHeight:1}}>{val}</div>
                <div style={{fontSize:10,color:C.t3,marginTop:5}}>{sub}</div>
              </div>
            ))}
          </div>
          <div style={S.card}>
            <div style={S.hdr}>S&OP MENSUAL — ANTES vs DESPUÉS · {sess?.name}</div>
            <ResponsiveContainer width="100%" height={195}>
              <ComposedChart data={monthChart} margin={{top:4,right:8,left:0,bottom:0}}>
                <CartesianGrid stroke={C.borderLt} strokeDasharray="4 4"/>
                <XAxis dataKey="name" tick={{fill:C.t2,fontSize:11,fontWeight:600}}/>
                <YAxis tick={{fill:C.t2,fontSize:11}} tickFormatter={fmtK}/>
                <Tooltip contentStyle={S.ttip} formatter={(v,n)=>[fmtK(v),n]}/>
                <Line type="monotone" dataKey="Antes"   stroke={C.indigo} strokeWidth={2}   strokeDasharray="5 3" dot={false}/>
                <Line type="monotone" dataKey="Después" stroke={C.amber}  strokeWidth={2.5} dot={false}/>
              </ComposedChart>
            </ResponsiveContainer>
          </div>
          <div style={S.card}>
            <div style={S.hdr}>Cambios por SKU · {UOM_LABEL[uom]}{significant.length>0&&<span style={{marginLeft:10,...S.tag(C.red)}}>{significant.length} significativos (≥{umbral}%)</span>}</div>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                <thead><tr style={{background:C.bg}}>{["","Item","Material","Canal",...MO,"Antes","Después","Δ","Δ %","Comentario"].map((h,hi)=><th key={hi} style={{padding:"7px 6px",color:C.t2,textAlign:["Item","Material","Canal","","Comentario"].includes(h)?"left":"right",fontSize:9,fontWeight:700,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                <tbody>
                  {comparison.map((d,i)=>{ const comments=[...new Set(sessLog.filter(c=>c.item===d.item&&c.canal===d.canal).map(c=>c.comment).filter(Boolean))]; return (
                    <tr key={i} style={{borderBottom:`1px solid ${C.borderLt}`,background:d.significant?"#fffbf0":i%2?C.bg:C.card}}>
                      <td style={{padding:"6px 4px 6px 10px",textAlign:"center"}}>{d.significant&&<span style={{fontSize:10,fontWeight:700,color:C.amber}}>★</span>}</td>
                      <td style={{padding:"6px 6px",color:C.accent,fontWeight:700}}>{d.item}</td>
                      <td style={{padding:"6px 6px",color:C.t2,maxWidth:150,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{d.mat}</td>
                      <td style={{padding:"6px 6px"}}><span style={S.tag(C.accent)}>{d.canal}</span></td>
                      {d.deltas.map((v,mi)=><td key={mi} style={{padding:"6px 5px",textAlign:"right",fontSize:10,color:v>0?C.green:v<0?C.red:C.t3,fontWeight:v!==0?700:400}}>{v!==0?(v>0?"+":"")+fmtK(v):"·"}</td>)}
                      <td style={{padding:"6px 6px",textAlign:"right",color:C.t2}}>{fmtK(d.totBefore)}</td>
                      <td style={{padding:"6px 6px",textAlign:"right",color:C.t1,fontWeight:700}}>{fmtK(d.totAfter)}</td>
                      <td style={{padding:"6px 6px",textAlign:"right",color:clr(d.deltaTotal),fontWeight:700}}>{d.deltaTotal>0?"+":""}{fmtK(d.deltaTotal)}</td>
                      <td style={{padding:"6px 6px",textAlign:"right"}}><span style={S.tag(clr(d.deltaPct))}>{d.deltaPct>0?"+":""}{d.deltaPct}%</span></td>
                      <td style={{padding:"6px 6px",color:C.t2,fontSize:10,fontStyle:"italic",maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{comments.join(" · ")||"—"}</td>
                    </tr>
                  );})}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
    </div>
  );
}

// ── CLIENTES VIEW ─────────────────────────────────────────────────────────
function ClientesView({ rows, uom, canal, C, S }) {
  const [fCadena,setFCadena]=useState("ALL");
  const [fCanal,setFCanal]=useState(canal==="ALL"?"ALL":canal);
  const canalesOpts=useMemo(()=>["ALL",...new Set(rows.map(r=>r.canal))]  ,[rows]);
  const cadenasOpts=useMemo(()=>{ const base=rows.filter(r=>fCanal==="ALL"||r.canal===fCanal); return ["ALL",...new Set(base.map(r=>r.cadena).filter(Boolean))]; },[rows,fCanal]);
  const filtered=useMemo(()=>rows.filter(r=>r.uom===uom&&(fCanal==="ALL"||r.canal===fCanal)&&(fCadena==="ALL"||r.cadena===fCadena)),[rows,uom,fCanal,fCadena]);
  const byCadena=useMemo(()=>{ const m={}; filtered.forEach(r=>{ const k=r.cadena||"—"; if(!m[k]) m[k]={cadena:k,canal:r.canal,S:0,P:0,L:0}; const v=sumArr(r.meses); if(r.figurasClave==="S&OP") m[k].S+=v; if(r.figurasClave==="POA 2026") m[k].P+=v; if(r.figurasClave==="LY 2025") m[k].L+=v; }); return Object.values(m).sort((a,b)=>b.S-a.S); },[filtered]);
  const bySKU=useMemo(()=>{ if(fCadena==="ALL") return []; const m={}; filtered.forEach(r=>{ if(!m[r.item]) m[r.item]={item:r.item,mat:r.material,cat:r.categoria,marca:r.marca,S:0,P:0,L:0}; const v=sumArr(r.meses); if(r.figurasClave==="S&OP") m[r.item].S+=v; if(r.figurasClave==="POA 2026") m[r.item].P+=v; if(r.figurasClave==="LY 2025") m[r.item].L+=v; }); return Object.values(m).sort((a,b)=>b.S-a.S); },[filtered,fCadena]);
  const totS=byCadena.reduce((s,g)=>s+g.S,0);
  const barData=byCadena.slice(0,10).map(c=>({name:c.cadena.slice(0,20),FCST:Math.round(c.S),POA:Math.round(c.P)}));
  const thS={padding:"7px 10px",color:C.t2,fontSize:9,fontWeight:700,letterSpacing:"0.08em",borderBottom:`1px solid ${C.border}`};
  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",marginBottom:18}}>
        <div><div style={{fontSize:17,fontWeight:700,color:C.t1}}>Vista por Cliente (Cadena)</div><div style={{fontSize:11,color:C.t2,marginTop:2}}>S&OP por cadena · {UOM_LABEL[uom]}</div></div>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          <select style={S.sel} value={fCanal} onChange={e=>{setFCanal(e.target.value);setFCadena("ALL");}}>{canalesOpts.map(o=><option key={o} value={o}>{o==="ALL"?"Todos":o}</option>)}</select>
          <select style={{...S.sel,maxWidth:240}} value={fCadena} onChange={e=>setFCadena(e.target.value)}>{cadenasOpts.map(o=><option key={o} value={o}>{o==="ALL"?"Todas":o}</option>)}</select>
        </div>
      </div>
      <div style={{display:"flex",gap:12,marginBottom:16}}>
        {byCadena.slice(0,3).map((c,i)=>(
          <div key={c.cadena} style={{flex:1,background:C.card,border:`1px solid ${C.border}`,borderTop:`3px solid ${i===0?C.amber:i===1?C.accent:C.indigo}`,borderRadius:8,padding:"12px 16px",boxShadow:"0 1px 4px rgba(30,48,85,0.06)"}}>
            <div style={{fontSize:9,color:C.t3,letterSpacing:"0.12em",textTransform:"uppercase",fontWeight:700,marginBottom:6}}>#{i+1} Cliente</div>
            <div style={{fontSize:12,fontWeight:700,color:C.t1,marginBottom:4,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.cadena}</div>
            <div style={{fontSize:22,fontWeight:800,color:C.t1,lineHeight:1}}>{fmtK(c.S)}</div>
            <div style={{fontSize:10,color:C.t3,marginTop:4}}>{totS>0?Math.round(c.S/totS*100):0}% del total</div>
            <div style={{marginTop:6,fontSize:10,color:clr(pct(c.S,c.P)),fontWeight:600}}>{arw(pct(c.S,c.P))} {Math.abs(pct(c.S,c.P))}% vs POA</div>
          </div>
        ))}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"3fr 2fr",gap:12,marginBottom:14}}>
        <div style={S.card}>
          <div style={S.hdr}>S&OP vs POA por Cadena (top 10)</div>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={barData} margin={{top:0,right:10,left:0,bottom:60}}>
              <CartesianGrid stroke={C.borderLt} vertical={false}/>
              <XAxis dataKey="name" tick={{fill:C.t2,fontSize:10}} angle={-30} textAnchor="end" interval={0}/>
              <YAxis tick={{fill:C.t2,fontSize:10}} tickFormatter={fmtK}/>
              <Tooltip contentStyle={S.ttip} formatter={(v,n)=>[fmtK(v),n]}/>
              <Bar dataKey="FCST" fill={C.amber}       radius={[3,3,0,0]} barSize={18}/>
              <Bar dataKey="POA"  fill={C.accent+"55"} radius={[3,3,0,0]} barSize={18}/>
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div style={S.card}>
          <div style={S.hdr}>Participación % S&OP</div>
          <div style={{display:"flex",flexDirection:"column",gap:7,marginTop:4}}>
            {byCadena.slice(0,8).map((c,i)=>{ const share=totS>0?Math.round(c.S/totS*100):0; const colors=["#d97706","#1e3055","#4f46e5","#059669","#dc2626","#0891b2","#7c3aed","#b45309"]; return (
              <div key={c.cadena}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}><span style={{fontSize:10,color:C.t2,maxWidth:185,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.cadena}</span><span style={{fontSize:10,fontWeight:700,color:C.t1}}>{share}%</span></div>
                <div style={{height:5,background:C.bg,borderRadius:3,overflow:"hidden"}}><div style={{width:`${share}%`,height:"100%",background:colors[i%colors.length],borderRadius:3}}/></div>
              </div>
            );})}
          </div>
        </div>
      </div>
      <div style={S.card}>
        <div style={S.hdr}>{fCadena!=="ALL"?`SKUs — ${fCadena.slice(0,40)}`:"Resumen por cadena — clic en fila para ver sus SKUs"}</div>
        {fCadena==="ALL"?(
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
            <thead><tr style={{background:C.bg}}><th style={{...thS,textAlign:"left"}}>Cadena</th><th style={{...thS,textAlign:"left"}}>Canal</th><th style={{...thS,textAlign:"right"}}>FCST</th><th style={{...thS,textAlign:"right"}}>POA</th><th style={{...thS,textAlign:"right"}}>LY 2025</th><th style={{...thS,textAlign:"right"}}>Partic.%</th><th style={{...thS,textAlign:"right"}}>Gap %</th></tr></thead>
            <tbody>
              {byCadena.map((c,i)=>(
                <tr key={c.cadena} style={{borderBottom:`1px solid ${C.borderLt}`,background:i%2?C.bg:C.card,cursor:"pointer"}} onClick={()=>setFCadena(c.cadena)}>
                  <td style={{padding:"7px 10px",color:C.t1,fontWeight:600,maxWidth:220,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.cadena}</td>
                  <td style={{padding:"7px 10px"}}><span style={S.tag(C.accent)}>{c.canal}</span></td>
                  <td style={{padding:"7px 10px",textAlign:"right",color:FIG_COLOR["S&OP"],fontWeight:700}}>{fmtK(c.S)}</td>
                  <td style={{padding:"7px 10px",textAlign:"right",color:FIG_COLOR["POA 2026"]}}>{fmtK(c.P)}</td>
                  <td style={{padding:"7px 10px",textAlign:"right",color:FIG_COLOR["LY 2025"]}}>{fmtK(c.L)}</td>
                  <td style={{padding:"7px 10px",textAlign:"right",color:C.t2,fontWeight:600}}>{totS>0?Math.round(c.S/totS*100):0}%</td>
                  <td style={{padding:"7px 10px",textAlign:"right"}}><span style={S.tag(clr(pct(c.S,c.P)))}>{pct(c.S,c.P)>0?"+":""}{pct(c.S,c.P)}%</span></td>
                </tr>
              ))}
            </tbody>
          </table>
        ):(
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
            <thead><tr style={{background:C.bg}}><th style={{...thS,textAlign:"left"}}>Item</th><th style={{...thS,textAlign:"left"}}>Material</th><th style={{...thS,textAlign:"left"}}>Cat.</th><th style={{...thS,textAlign:"left"}}>Marca</th><th style={{...thS,textAlign:"right"}}>FCST</th><th style={{...thS,textAlign:"right"}}>POA</th><th style={{...thS,textAlign:"right"}}>LY 2025</th><th style={{...thS,textAlign:"right"}}>Partic.%</th></tr></thead>
            <tbody>
              {bySKU.map((s,i)=>{ const skuTot=bySKU.reduce((a,x)=>a+x.S,0); return (
                <tr key={s.item} style={{borderBottom:`1px solid ${C.borderLt}`,background:i%2?C.bg:C.card}}>
                  <td style={{padding:"7px 10px",color:C.accent,fontWeight:700}}>{s.item}</td>
                  <td style={{padding:"7px 10px",color:C.t2,maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.mat}</td>
                  <td style={{padding:"7px 10px"}}><span style={S.tag(C.t2)}>{s.cat}</span></td>
                  <td style={{padding:"7px 10px",color:C.t2}}>{s.marca}</td>
                  <td style={{padding:"7px 10px",textAlign:"right",color:FIG_COLOR["S&OP"],fontWeight:700}}>{fmtK(s.S)}</td>
                  <td style={{padding:"7px 10px",textAlign:"right",color:FIG_COLOR["POA 2026"]}}>{fmtK(s.P)}</td>
                  <td style={{padding:"7px 10px",textAlign:"right",color:FIG_COLOR["LY 2025"]}}>{fmtK(s.L)}</td>
                  <td style={{padding:"7px 10px",textAlign:"right",color:C.t2,fontWeight:600}}>{skuTot>0?Math.round(s.S/skuTot*100):0}%</td>
                </tr>
              );})}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

// ── ABC-XYZ VIEW ──────────────────────────────────────────────────────────
function AbcXyzView({ rows, uom, C, S }) {
  const [fCanal,setFCanal]=useState("ALL");
  const [highlight,setHighlight]=useState(null);
  const [showCfg,setShowCfg]=useState(false);
  const [selFig,setSelFig]=useState("S&OP");
  const [selMonths,setSelMonths]=useState([4,5,6,7,8,9,10,11]);
  const canalesOpts=useMemo(()=>["ALL",...new Set(rows.map(r=>r.canal))],[rows]);
  const toggleMonth=(i)=>setSelMonths(p=>p.includes(i)?p.filter(x=>x!==i):[...p,i]);
  const matrix=useMemo(()=>{ if(selMonths.length===0) return []; const src=rows.filter(r=>r.uom===uom&&r.figurasClave===selFig&&(fCanal==="ALL"||r.canal===fCanal)); const byItem={}; src.forEach(r=>{ const k=`${r.canal}||${r.item}`; if(!byItem[k]) byItem[k]={item:r.item,mat:r.material,cat:r.categoria,marca:r.marca,canal:r.canal,meses:Array(12).fill(0)}; r.meses.forEach((v,i)=>{byItem[k].meses[i]+=v;}); }); const items=Object.values(byItem).map(d=>{ const usedMeses=selMonths.map(i=>d.meses[i]); const total=sumArr(usedMeses),mean=total/selMonths.length; const variance=usedMeses.reduce((s,v)=>s+Math.pow(v-mean,2),0)/selMonths.length; const cv=mean>0?Math.sqrt(variance)/mean:0; return {...d,total,cv}; }).filter(d=>d.total>0).sort((a,b)=>b.total-a.total); const grand=items.reduce((s,d)=>s+d.total,0); let cum=0; items.forEach(d=>{ cum+=d.total; const sh=grand>0?cum/grand:0; d.abc=sh<=0.80?"A":sh<=0.95?"B":"C"; d.sharePct=grand>0?Math.round(d.total/grand*100):0; d.xyz=d.cv<0.2?"X":d.cv<0.5?"Y":"Z"; d.class=d.abc+d.xyz; }); return items; },[rows,uom,fCanal,selFig,selMonths]);
  const cellCount=useMemo(()=>{ const m={}; ["A","B","C"].forEach(a=>["X","Y","Z"].forEach(x=>{m[a+x]={count:0,vol:0,items:[]};})); matrix.forEach(d=>{m[d.class].count++;m[d.class].vol+=d.total;m[d.class].items.push(d);}); return m; },[matrix]);
  const grandTotal=matrix.reduce((s,d)=>s+d.total,0);
  const filtered=highlight?cellCount[highlight].items:matrix;
  return (
    <div>
      {showCfg&&(
        <div style={S.modal} onClick={e=>e.target===e.currentTarget&&setShowCfg(false)}>
          <div style={S.dlg}>
            <div style={{fontSize:14,fontWeight:700,color:C.t1,marginBottom:16}}>⚙ Configurar cálculo ABC-XYZ</div>
            <div style={{marginBottom:14}}>
              <div style={{fontSize:10,color:C.t2,fontWeight:700,marginBottom:8,textTransform:"uppercase",letterSpacing:"0.08em"}}>Figura de datos</div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>{FIGS.map(f=><button key={f} style={{...S.btn(false,selFig===f),fontSize:10,padding:"5px 10px"}} onClick={()=>setSelFig(f)}>{FIG_LABEL[f]}</button>)}</div>
            </div>
            <div style={{marginBottom:14}}>
              <div style={{fontSize:10,color:C.t2,fontWeight:700,marginBottom:8,textTransform:"uppercase",letterSpacing:"0.08em"}}>Meses a incluir ({selMonths.length} sel.)</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:6}}>{MO.map((m,i)=><button key={m} onClick={()=>toggleMonth(i)} style={{padding:"5px 10px",borderRadius:6,border:`1px solid ${selMonths.includes(i)?C.accent:C.border}`,background:selMonths.includes(i)?C.accentLt:C.card,color:selMonths.includes(i)?C.accent:C.t3,fontSize:10,fontWeight:600,cursor:"pointer"}}>{m}</button>)}</div>
              <div style={{marginTop:10,display:"flex",gap:8}}>
                {[["Todo el año",[0,1,2,3,4,5,6,7,8,9,10,11]],["Últimos 3m",[9,10,11]],["May–Dic",[4,5,6,7,8,9,10,11]]].map(([l,ms])=><button key={l} style={{...S.btn(false,false),fontSize:10}} onClick={()=>setSelMonths(ms)}>{l}</button>)}
              </div>
            </div>
            <div style={{padding:"10px 12px",background:C.accentLt,borderRadius:8,marginBottom:16,fontSize:11,color:C.t2}}>
              <strong>Figura:</strong> {FIG_LABEL[selFig]} &nbsp;·&nbsp; <strong>Meses:</strong> {selMonths.length>0?selMonths.map(i=>MO[i]).join(", "):"ninguno"}
            </div>
            <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
              <button style={S.btn(false)} onClick={()=>setShowCfg(false)}>Cancelar</button>
              <button style={S.btn(true)} onClick={()=>setShowCfg(false)}>Aplicar</button>
            </div>
          </div>
        </div>
      )}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",marginBottom:14}}>
        <div><div style={{fontSize:17,fontWeight:700,color:C.t1}}>Matriz ABC-XYZ</div><div style={{fontSize:11,color:C.t2,marginTop:2}}>ABC: volumen acumulado · XYZ: coeficiente de variación</div></div>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          <select style={S.sel} value={fCanal} onChange={e=>setFCanal(e.target.value)}>{canalesOpts.map(o=><option key={o} value={o}>{o==="ALL"?"Todos":o}</option>)}</select>
          <button style={S.btn(false,true)} onClick={()=>setShowCfg(true)}>⚙ Configurar</button>
          {highlight&&<button style={{...S.btn(false,true),fontSize:10}} onClick={()=>setHighlight(null)}>✕ Limpiar</button>}
        </div>
      </div>
      <div style={{display:"flex",gap:6,marginBottom:12,flexWrap:"wrap",alignItems:"center"}}>
        <span style={{fontSize:10,color:C.t3}}>Base:</span>
        <span style={S.tag(FIG_COLOR[selFig])}>{FIG_LABEL[selFig]}</span>
        {selMonths.map(i=><span key={i} style={{...S.tag(C.accent),fontSize:9}}>{MO[i]}</span>)}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"auto 1fr",gap:20,marginBottom:16,alignItems:"start"}}>
        <div style={{...S.card,padding:20,minWidth:360}}>
          <div style={S.hdr}>Clic en celda para filtrar</div>
          <div style={{display:"flex",gap:4,marginBottom:6,paddingLeft:64}}>
            {["X — Baja CV","Y — Media CV","Z — Alta CV"].map((l,i)=><div key={i} style={{flex:1,textAlign:"center",fontSize:9,color:C.t2,fontWeight:700}}>{l.split(" — ")[0]}<br/><span style={{fontWeight:400,color:C.t3,fontSize:8}}>{l.split(" — ")[1]}</span></div>)}
          </div>
          {["A","B","C"].map(abc=>(
            <div key={abc} style={{display:"flex",gap:4,marginBottom:4}}>
              <div style={{width:60,display:"flex",flexDirection:"column",justifyContent:"center",alignItems:"center",borderRadius:6,background:C.bg,padding:"4px 0"}}>
                <span style={{fontSize:12,fontWeight:800,color:C.accent}}>{abc}</span>
                <span style={{fontSize:8,color:C.t3}}>{abc==="A"?"≥80%":abc==="B"?"80-95%":"<95%"}</span>
              </div>
              {["X","Y","Z"].map(xyz=>{ const key=abc+xyz,cell=cellCount[key],isHl=highlight===key,bg=CELL_COLORS[key],volPct=grandTotal>0?Math.round(cell.vol/grandTotal*100):0; return (
                <div key={key} onClick={()=>setHighlight(isHl?null:key)} style={{flex:1,background:isHl?bg:bg+"33",border:`2px solid ${isHl?bg:bg+"66"}`,borderRadius:8,padding:"10px 8px",cursor:"pointer",textAlign:"center",transition:"all 0.15s"}}>
                  <div style={{fontSize:18,fontWeight:800,color:isHl?"#fff":C.t1}}>{cell.count}</div>
                  <div style={{fontSize:8,color:isHl?"rgba(255,255,255,0.8)":C.t3}}>SKUs</div>
                  <div style={{fontSize:10,fontWeight:700,color:isHl?"#fff":C.t2,marginTop:4}}>{fmtK(cell.vol)}</div>
                  <div style={{fontSize:8,color:isHl?"rgba(255,255,255,0.7)":C.t3}}>{volPct}% vol</div>
                  <div style={{marginTop:5,fontSize:8,color:isHl?"rgba(255,255,255,0.7)":C.t3,fontStyle:"italic"}}>{CELL_LABEL[key]}</div>
                </div>
              );})}
            </div>
          ))}
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <div style={S.card}>
            <div style={S.hdr}>Volumen por clase ABC</div>
            <ResponsiveContainer width="100%" height={130}>
              <BarChart data={["A","B","C"].map(abc=>({name:abc,vol:Math.round(matrix.filter(d=>d.abc===abc).reduce((s,d)=>s+d.total,0)),skus:matrix.filter(d=>d.abc===abc).length}))} margin={{top:0,right:10,left:0,bottom:0}}>
                <CartesianGrid stroke={C.borderLt} vertical={false}/>
                <XAxis dataKey="name" tick={{fill:C.t2,fontSize:12,fontWeight:700}}/>
                <YAxis tick={{fill:C.t2,fontSize:10}} tickFormatter={fmtK}/>
                <Tooltip contentStyle={S.ttip} formatter={(v,n)=>[n==="vol"?fmtK(v):v,n==="vol"?"Volumen":"SKUs"]}/>
                <Bar dataKey="vol" fill={C.amber} radius={[4,4,0,0]} barSize={40}/>
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div style={S.card}>
            <div style={S.hdr}>Variabilidad XYZ</div>
            <div style={{display:"flex",gap:10}}>
              {[["X","CV < 0.2","Alta estabilidad",C.green],["Y","CV 0.2–0.5","Moderada",C.amber],["Z","CV > 0.5","Alta volatilidad",C.red]].map(([k,cv,desc,c])=>(
                <div key={k} style={{flex:1,background:c+"10",border:`1px solid ${c}33`,borderRadius:8,padding:"10px 12px"}}>
                  <div style={{fontSize:18,fontWeight:800,color:c}}>{k}</div>
                  <div style={{fontSize:10,fontWeight:700,color:C.t1,marginTop:2}}>{cv}</div>
                  <div style={{fontSize:10,color:C.t2,marginTop:2}}>{desc}</div>
                  <div style={{fontSize:10,color:C.t2,marginTop:5}}><strong>{matrix.filter(d=>d.xyz===k).length}</strong> SKUs</div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
      <div style={S.card}>
        <div style={S.hdr}>{highlight?`Clase ${highlight} — ${CELL_LABEL[highlight]} (${cellCount[highlight].count})`:`Todos los SKUs (${matrix.length})`}</div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
            <thead><tr style={{background:C.bg}}>{["Item","Material","Cat.","Marca","Canal","ABC","XYZ","Clase","Volumen","Partic.%","CV"].map(h=><th key={h} style={{padding:"7px 10px",color:C.t2,textAlign:["Volumen","Partic.%","CV"].includes(h)?"right":"left",fontSize:9,fontWeight:700,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
            <tbody>
              {filtered.map((d,i)=>{ const bg=CELL_COLORS[d.class]; return (
                <tr key={d.item+(d.canal||"")+i} style={{borderBottom:`1px solid ${C.borderLt}`,background:i%2?C.bg:C.card}}>
                  <td style={{padding:"7px 10px",color:C.accent,fontWeight:700}}>{d.item}</td>
                  <td style={{padding:"7px 10px",color:C.t2,maxWidth:170,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{d.mat}</td>
                  <td style={{padding:"7px 10px"}}><span style={S.tag(C.t2)}>{d.cat}</span></td>
                  <td style={{padding:"7px 10px",color:C.t2}}>{d.marca}</td>
                  <td style={{padding:"7px 10px",color:C.t2}}>{d.canal}</td>
                  <td style={{padding:"7px 10px"}}><span style={S.tag(d.abc==="A"?C.green:d.abc==="B"?C.amber:C.red)}>{d.abc}</span></td>
                  <td style={{padding:"7px 10px"}}><span style={S.tag(d.xyz==="X"?C.green:d.xyz==="Y"?C.amber:C.red)}>{d.xyz}</span></td>
                  <td style={{padding:"7px 10px"}}><span style={{padding:"2px 8px",borderRadius:10,fontSize:9,fontWeight:700,background:bg+"33",color:C.t1,border:`1px solid ${bg}66`}}>{d.class}</span></td>
                  <td style={{padding:"7px 10px",textAlign:"right",color:FIG_COLOR["S&OP"],fontWeight:700}}>{fmtK(d.total)}</td>
                  <td style={{padding:"7px 10px",textAlign:"right",color:C.t2}}>{d.sharePct}%</td>
                  <td style={{padding:"7px 10px",textAlign:"right",color:C.t2}}>{d.cv.toFixed(2)}</td>
                </tr>
              );})}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ── GAPS VIEW ─────────────────────────────────────────────────────────────
function GapsView({ rows, uom, C, S }) {
  const [fCanal,setFCanal]=useState("ALL");
  const canalesOpts=useMemo(()=>["ALL",...new Set(rows.map(r=>r.canal))],[rows]);
  const filtered=useMemo(()=>rows.filter(r=>r.uom===uom&&(fCanal==="ALL"||r.canal===fCanal)),[rows,uom,fCanal]);
  const gaps=useMemo(()=>{ const m={}; filtered.forEach(r=>{ const k=`${r.canal}||${r.item}`; if(!m[k]) m[k]={item:r.item,mat:r.material.slice(0,30),canal:r.canal,S:0,P:0,E:0,L:0}; const v=sumArr(r.meses); if(r.figurasClave==="S&OP") m[k].S+=v; if(r.figurasClave==="POA 2026") m[k].P+=v; if(r.figurasClave==="Estadístico") m[k].E+=v; if(r.figurasClave==="LY 2025") m[k].L+=v; }); return Object.values(m).map(d=>({...d,gapAbs:d.S-d.P,gapPct:pct(d.S,d.P)})).sort((a,b)=>Math.abs(b.gapAbs)-Math.abs(a.gapAbs)).slice(0,12); },[filtered]);
  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
        <div><div style={{fontSize:17,fontWeight:700,color:C.t1}}>Análisis de Gaps</div><div style={{fontSize:11,color:C.t2}}>FCST vs referencias · {UOM_LABEL[uom]} · Canal: <strong>{fCanal==="ALL"?"Todos":fCanal}</strong></div></div>
        <div style={{display:"flex",alignItems:"center",gap:6}}><span style={{fontSize:11,color:C.t2}}>Canal</span><select style={S.sel} value={fCanal} onChange={e=>setFCanal(e.target.value)}>{canalesOpts.map(o=><option key={o} value={o}>{o==="ALL"?"Todos":o}</option>)}</select></div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:14}}>
        <div style={S.card}>
          <div style={S.hdr}>Top SKUs — Gap FCST vs POA</div>
          <ResponsiveContainer width="100%" height={310}>
            <BarChart data={gaps} layout="vertical" margin={{left:0,right:36}}>
              <CartesianGrid stroke={C.borderLt} horizontal={false}/>
              <XAxis type="number" tick={{fill:C.t2,fontSize:10}} tickFormatter={fmtK}/>
              <YAxis dataKey="mat" type="category" tick={{fill:C.t2,fontSize:10}} width={155}/>
              <Tooltip contentStyle={S.ttip} formatter={v=>[fmtK(v)]}/>
              <Bar dataKey="gapAbs" name="Gap" fill={C.amber} radius={[0,3,3,0]} barSize={18} label={{position:"right",fill:C.t2,fontSize:10,formatter:v=>(v>0?"+":"")+fmtK(v)}}/>
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div style={S.card}>
          <div style={S.hdr}>% Gap FCST vs POA</div>
          <ResponsiveContainer width="100%" height={310}>
            <BarChart data={gaps} layout="vertical" margin={{left:0,right:32}}>
              <CartesianGrid stroke={C.borderLt} horizontal={false}/>
              <XAxis type="number" tick={{fill:C.t2,fontSize:10}} unit="%"/>
              <YAxis dataKey="mat" type="category" tick={{fill:C.t2,fontSize:10}} width={155}/>
              <ReferenceLine x={0} stroke={C.t3}/>
              <Tooltip contentStyle={S.ttip} formatter={v=>[v+"%"]}/>
              <Bar dataKey="gapPct" name="Gap %" fill={C.indigo} radius={[0,3,3,0]} barSize={18}/>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
      <div style={S.card}>
        <div style={S.hdr}>Detalle completo</div>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
          <thead><tr style={{background:C.bg}}>{["Item","Material","Canal","FCST","POA","LY 2025","EST","Gap Abs","Gap %","vs LY"].map(h=><th key={h} style={{padding:"8px 10px",color:C.t2,textAlign:["Item","Material","Canal"].includes(h)?"left":"right",fontSize:9,fontWeight:700,borderBottom:`1px solid ${C.border}`}}>{h}</th>)}</tr></thead>
          <tbody>
            {gaps.map((g,i)=>(
              <tr key={i} style={{borderBottom:`1px solid ${C.borderLt}`,background:i%2?C.bg:C.card}}>
                <td style={{padding:"7px 10px",color:C.accent,fontWeight:700}}>{g.item}</td>
                <td style={{padding:"7px 10px",color:C.t2}}>{g.mat}</td>
                <td style={{padding:"7px 10px"}}><span style={S.tag(C.accent)}>{g.canal}</span></td>
                <td style={{padding:"7px 10px",textAlign:"right",color:FIG_COLOR["S&OP"],fontWeight:700}}>{fmtK(g.S)}</td>
                <td style={{padding:"7px 10px",textAlign:"right",color:FIG_COLOR["POA 2026"]}}>{fmtK(g.P)}</td>
                <td style={{padding:"7px 10px",textAlign:"right",color:FIG_COLOR["LY 2025"]}}>{fmtK(g.L)}</td>
                <td style={{padding:"7px 10px",textAlign:"right",color:FIG_COLOR["Estadístico"]}}>{fmtK(g.E)}</td>
                <td style={{padding:"7px 10px",textAlign:"right",color:clr(g.gapAbs),fontWeight:700}}>{g.gapAbs>0?"+":""}{fmtK(g.gapAbs)}</td>
                <td style={{padding:"7px 10px",textAlign:"right"}}><span style={S.tag(clr(g.gapPct))}>{g.gapPct>0?"+":""}{g.gapPct}%</span></td>
                <td style={{padding:"7px 10px",textAlign:"right",color:clr(pct(g.S,g.L))}}>{pct(g.S,g.L)>0?"+":""}{pct(g.S,g.L)}%</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ── MATERIA PRIMA ─────────────────────────────────────────────────────────
function MateriaPrimaView({ rows, C, S }) {
  const [fCanal,setFCanal]=useState("ALL");
  const canalesOpts=useMemo(()=>["ALL",...new Set(rows.map(r=>r.canal))],[rows]);
  const filtered=useMemo(()=>rows.filter(r=>r.uom==="QQI"&&r.figurasClave==="S&OP"&&(fCanal==="ALL"||r.canal===fCanal)),[rows,fCanal]);
  const bySub=useMemo(()=>{ const m={}; filtered.forEach(r=>{ const k=r.subCategoria||r.categoria||"—"; if(!m[k]) m[k]={sub:k,meses:Array(12).fill(0)}; r.meses.forEach((v,i)=>{m[k].meses[i]+=v;}); }); return Object.values(m).sort((a,b)=>sumArr(b.meses)-sumArr(a.meses)); },[filtered]);
  const totalMeses=MO.map((_,i)=>bySub.reduce((s,g)=>s+(g.meses[i]||0),0));
  const totalQqi=sumArr(totalMeses),totalTon=totalQqi/22;
  const chartData=MO.map((m,i)=>({name:m,QQI:Math.round(totalMeses[i]),TM:parseFloat((totalMeses[i]/22).toFixed(1))}));
  const subColors=["#d97706","#1e3055","#4f46e5","#059669","#dc2626","#0891b2","#7c3aed","#b45309"];
  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",marginBottom:18}}>
        <div><div style={{fontSize:17,fontWeight:700,color:C.t1}}>Necesidades de Materia Prima</div><div style={{fontSize:11,color:C.t2,marginTop:2}}>Basado en S&OP · Quintales y Toneladas <span style={{color:C.t3}}>(1 TM = 22 QQI)</span></div></div>
        <select style={S.sel} value={fCanal} onChange={e=>setFCanal(e.target.value)}>{canalesOpts.map(o=><option key={o} value={o}>{o==="ALL"?"Todos":o}</option>)}</select>
      </div>
      <div style={{display:"flex",gap:12,marginBottom:16}}>
        {[["Total Quintales","QQI",fmtK(totalQqi),C.amber],["Total Toneladas","TM",totalTon.toFixed(1),C.green],["Sub-categorías","",String(bySub.length),C.indigo]].map(([lbl,unit,val,c])=>(
          <div key={lbl} style={{flex:1,background:C.card,border:`1px solid ${C.border}`,borderTop:`3px solid ${c}`,borderRadius:8,padding:"14px 18px",boxShadow:"0 1px 4px rgba(30,48,85,0.06)"}}>
            <div style={{fontSize:9,color:C.t3,letterSpacing:"0.14em",textTransform:"uppercase",marginBottom:8,fontWeight:700}}>{lbl}</div>
            <div style={{fontSize:28,fontWeight:800,color:C.t1,lineHeight:1}}>{val}</div>
            {unit&&<div style={{fontSize:11,color:C.t3,marginTop:5}}>{unit}</div>}
          </div>
        ))}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:12,marginBottom:14}}>
        <div style={S.card}>
          <div style={S.hdr}>Necesidad mensual (QQI y TM)</div>
          <ResponsiveContainer width="100%" height={200}>
            <ComposedChart data={chartData} margin={{top:4,right:8,left:0,bottom:0}}>
              <CartesianGrid stroke={C.borderLt} strokeDasharray="4 4"/>
              <XAxis dataKey="name" tick={{fill:C.t2,fontSize:11}}/>
              <YAxis yAxisId="l" tick={{fill:C.t2,fontSize:10}} tickFormatter={fmtK}/>
              <YAxis yAxisId="r" orientation="right" tick={{fill:C.t2,fontSize:10}} tickFormatter={fmtK}/>
              <Tooltip contentStyle={S.ttip} formatter={(v,n)=>[n==="QQI"?fmtK(v)+" QQI":v.toFixed(1)+" TM",n]}/>
              <Bar yAxisId="l" dataKey="QQI" fill={C.amber+"bb"} radius={[3,3,0,0]} barSize={22}/>
              <Line yAxisId="r" type="monotone" dataKey="TM" stroke={C.green} strokeWidth={2.5} dot={{r:3,fill:C.green}} name="TM"/>
            </ComposedChart>
          </ResponsiveContainer>
        </div>
        <div style={S.card}>
          <div style={S.hdr}>Por sub-categoría</div>
          <div style={{display:"flex",flexDirection:"column",gap:7,marginTop:4}}>
            {bySub.map((s,i)=>{ const share=totalQqi>0?Math.round(sumArr(s.meses)/totalQqi*100):0; return (
              <div key={s.sub}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}><span style={{fontSize:10,color:C.t2,fontWeight:600}}>{s.sub}</span><span style={{fontSize:10,color:C.t1,fontWeight:700}}>{share}%</span></div>
                <div style={{height:5,background:C.bg,borderRadius:3,overflow:"hidden"}}><div style={{width:`${share}%`,height:"100%",background:subColors[i%subColors.length],borderRadius:3}}/></div>
              </div>
            );})}
          </div>
        </div>
      </div>
      <div style={S.card}>
        <div style={S.hdr}>Detalle mensual por sub-categoría</div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:10}}>
            <thead><tr style={{background:C.bg}}>
              <th style={{padding:"7px 10px",textAlign:"left",color:C.t2,fontSize:9,fontWeight:700,borderBottom:`1px solid ${C.border}`}}>SUB-CATEGORÍA</th>
              <th style={{padding:"7px 10px",textAlign:"left",color:C.t2,fontSize:9,fontWeight:700,borderBottom:`1px solid ${C.border}`}}>UND</th>
              {MO.map(m=><th key={m} style={{padding:"7px 8px",textAlign:"right",color:C.t2,fontSize:9,fontWeight:700,borderBottom:`1px solid ${C.border}`}}>{m}</th>)}
              <th style={{padding:"7px 10px",textAlign:"right",color:C.t2,fontSize:9,fontWeight:700,borderBottom:`1px solid ${C.border}`}}>TOTAL</th>
            </tr></thead>
            <tbody>
              {bySub.map((s,si)=>{ const qqi=s.meses,ton=s.meses.map(v=>parseFloat((v/22).toFixed(1))); return [
                <tr key={s.sub+"-q"} style={{borderBottom:`1px solid ${C.borderLt}`,background:si%2?C.bg:C.card}}>
                  <td rowSpan={2} style={{padding:"7px 10px",color:C.t1,fontWeight:700,verticalAlign:"middle"}}>{s.sub}</td>
                  <td style={{padding:"4px 10px"}}><span style={S.tag(C.amber)}>QQI</span></td>
                  {qqi.map((v,i)=><td key={i} style={{padding:"4px 8px",textAlign:"right",color:C.t2,fontWeight:600}}>{fmtK(v)}</td>)}
                  <td style={{padding:"4px 10px",textAlign:"right",color:C.amber,fontWeight:700}}>{fmtK(sumArr(qqi))}</td>
                </tr>,
                <tr key={s.sub+"-t"} style={{borderBottom:`2px solid ${C.border}`,background:si%2?"#ecf9f4":"#f5fcf7"}}>
                  <td style={{padding:"4px 10px"}}><span style={S.tag(C.green)}>TM</span></td>
                  {ton.map((v,i)=><td key={i} style={{padding:"4px 8px",textAlign:"right",color:C.green,fontWeight:600}}>{v.toFixed(1)}</td>)}
                  <td style={{padding:"4px 10px",textAlign:"right",color:C.green,fontWeight:700}}>{(sumArr(qqi)/22).toFixed(1)}</td>
                </tr>
              ];})}
            </tbody>
            <tfoot>
              <tr style={{borderTop:`2px solid ${C.border}`,background:C.accentLt}}>
                <td style={{padding:"8px 10px",fontWeight:700,color:C.t1}}>TOTAL</td><td style={{padding:"8px 10px"}}><span style={S.tag(C.amber)}>QQI</span></td>
                {totalMeses.map((v,i)=><td key={i} style={{padding:"8px 8px",textAlign:"right",color:C.amber,fontWeight:700}}>{fmtK(v)}</td>)}
                <td style={{padding:"8px 10px",textAlign:"right",color:C.amber,fontWeight:700}}>{fmtK(totalQqi)}</td>
              </tr>
              <tr style={{background:"#f5fcf7"}}>
                <td/><td style={{padding:"8px 10px"}}><span style={S.tag(C.green)}>TM</span></td>
                {totalMeses.map((v,i)=><td key={i} style={{padding:"8px 8px",textAlign:"right",color:C.green,fontWeight:700}}>{(v/22).toFixed(1)}</td>)}
                <td style={{padding:"8px 10px",textAlign:"right",color:C.green,fontWeight:700}}>{totalTon.toFixed(1)}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
    </div>
  );
}

// ── REUNIONES VIEW ────────────────────────────────────────────────────────
const DEMO_MEETINGS=[
  {id:"dm1",title:"Revisión Demanda Mayo 2026",type:"Revisión de Demanda",date:"2026-05-08",time:"09:00",duration:90,canal:"MODERNO",attendees:"ana.gonzalez@alcsa.com, pedro.ruiz@alcsa.com",teamsUrl:"",agenda:["Revisión S&OP vs POA","Ajustes rolling forecast","Compromisos comerciales"],status:"realizada",minuta:""},
  {id:"dm2",title:"Revisión Demanda Junio 2026",type:"Revisión de Demanda",date:"2026-06-07",time:"09:00",duration:90,canal:"ALL",attendees:"ana.gonzalez@alcsa.com, pedro.ruiz@alcsa.com",teamsUrl:"",agenda:["S&OP vs POA acumulado","Ajustes por marca","Acuerdos"],status:"pendiente",minuta:""},
  {id:"dm3",title:"Follow-up MODERNO",type:"Follow-up",date:"2026-06-22",time:"14:00",duration:45,canal:"MODERNO",attendees:"ana.gonzalez@alcsa.com",teamsUrl:"",agenda:["Avance compromisos"],status:"pendiente",minuta:""},
];

function ReunionesView({ C, S }) {
  const [meetings,setMeetings]=useState(DEMO_MEETINGS);
  const [curMonth,setCurMonth]=useState({year:2026,month:5});
  const [showNew,setShowNew]=useState(false);
  const [selM,setSelM]=useState(null);
  const [genLoading,setGenLoading]=useState(false);
  const [copied,setCopied]=useState(false);
  const [newForm,setNewForm]=useState({title:"",type:"Revisión de Demanda",date:"",time:"09:00",duration:90,canal:"ALL",attendees:"",teamsUrl:"",agenda:["","",""],status:"pendiente"});

  useEffect(()=>{(async()=>{ try{ const r=await window.storage?.get("dr2-meetings"); if(r) setMeetings(JSON.parse(r.value)); }catch(e){} })();},[]);
  const save=(m)=>{ setMeetings(m); window.storage?.set("dr2-meetings",JSON.stringify(m)).catch(()=>{}); };
  const updSel=(patch)=>{ const u=meetings.map(m=>m.id===selM?.id?{...m,...patch}:m); save(u); setSelM(p=>({...p,...patch})); };

  const daysInMonth=(y,mo)=>new Date(y,mo+1,0).getDate();
  const firstDayOf=(y,mo)=>new Date(y,mo,1).getDay();
  const today=new Date().toISOString().slice(0,10);
  const monthStr=`${curMonth.year}-${String(curMonth.month+1).padStart(2,"0")}`;
  const meetingsByDay=useMemo(()=>{ const m={}; meetings.filter(x=>x.date.startsWith(monthStr)).forEach(x=>{ const d=new Date(x.date+"T12:00").getDate(); if(!m[d]) m[d]=[]; m[d].push(x); }); return m; },[meetings,monthStr]);

  const generateMinuta=async()=>{
    if(!selM) return;
    setGenLoading(true);
    try {
      const prompt=`Eres el asistente de planeación de demanda de ALCSA (empresa guatemalteca de granos: arroz, maíz y avena). Genera una minuta profesional y concisa para la siguiente reunión de S&OP:\n\nTítulo: ${selM.title}\nTipo: ${selM.type}\nFecha: ${selM.date} | Hora: ${selM.time} | Duración: ${selM.duration} min\nCanal: ${selM.canal}\nAsistentes: ${selM.attendees}\n\nAgenda:\n${selM.agenda.map((a,i)=>`${i+1}. ${a}`).join("\n")}\n\nEstructura:\n1. ENCABEZADO\n2. RESUMEN EJECUTIVO (2-3 líneas)\n3. DESARROLLO POR PUNTO DE AGENDA\n4. ACUERDOS Y COMPROMISOS (tabla: Compromiso | Responsable | Fecha límite)\n5. PRÓXIMOS PASOS\n\nTono ejecutivo, todo en español.`;
      const res=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:1000,messages:[{role:"user",content:prompt}]})});
      const data=await res.json();
      const text=data.content?.find(b=>b.type==="text")?.text||"Error al generar la minuta.";
      updSel({minuta:text});
    } catch(e){ console.error(e); }
    setGenLoading(false);
  };

  const days=daysInMonth(curMonth.year,curMonth.month);
  const fd=firstDayOf(curMonth.year,curMonth.month);

  const createMeeting=()=>{
    const m={...newForm,id:"m"+Date.now(),minuta:"",agenda:newForm.agenda.filter(a=>a.trim())};
    const u=[...meetings,m]; save(u); setShowNew(false); setSelM(m);
    setNewForm({title:"",type:"Revisión de Demanda",date:"",time:"09:00",duration:90,canal:"ALL",attendees:"",teamsUrl:"",agenda:["","",""],status:"pendiente"});
  };

  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",marginBottom:18}}>
        <div><div style={{fontSize:17,fontWeight:700,color:C.t1}}>Reuniones S&OP</div><div style={{fontSize:11,color:C.t2,marginTop:2}}>Agenda, minutas generadas por IA y seguimiento</div></div>
        <button style={S.btn(true)} onClick={()=>setShowNew(true)}>+ Nueva Reunión</button>
      </div>

      {/* Nueva reunión modal */}
      {showNew&&(
        <div style={S.modal} onClick={e=>e.target===e.currentTarget&&setShowNew(false)}>
          <div style={{...S.dlg,width:560,maxHeight:"92vh",overflowY:"auto"}}>
            <div style={{fontSize:14,fontWeight:700,color:C.t1,marginBottom:16}}>Nueva Reunión</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
              <div style={{gridColumn:"1/-1"}}><label style={{fontSize:10,color:C.t2,fontWeight:600,display:"block",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.08em"}}>Título</label><input type="text" value={newForm.title} onChange={e=>setNewForm(p=>({...p,title:e.target.value}))} style={{...S.inpTxt,width:"100%",boxSizing:"border-box"}}/></div>
              <div><label style={{fontSize:10,color:C.t2,fontWeight:600,display:"block",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.08em"}}>Tipo</label><select value={newForm.type} onChange={e=>setNewForm(p=>({...p,type:e.target.value}))} style={{...S.sel,width:"100%",padding:"7px 8px"}}>{MEETING_TYPES.map(t=><option key={t}>{t}</option>)}</select></div>
              <div><label style={{fontSize:10,color:C.t2,fontWeight:600,display:"block",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.08em"}}>Canal</label><input type="text" value={newForm.canal} onChange={e=>setNewForm(p=>({...p,canal:e.target.value}))} style={{...S.inpTxt,width:"100%",boxSizing:"border-box"}}/></div>
              <div><label style={{fontSize:10,color:C.t2,fontWeight:600,display:"block",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.08em"}}>Fecha</label><input type="date" value={newForm.date} onChange={e=>setNewForm(p=>({...p,date:e.target.value}))} style={{...S.inpTxt,width:"100%",boxSizing:"border-box"}}/></div>
              <div><label style={{fontSize:10,color:C.t2,fontWeight:600,display:"block",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.08em"}}>Hora</label><input type="time" value={newForm.time} onChange={e=>setNewForm(p=>({...p,time:e.target.value}))} style={{...S.inpTxt,width:"100%",boxSizing:"border-box"}}/></div>
            </div>
            <div style={{marginBottom:10}}><label style={{fontSize:10,color:C.t2,fontWeight:600,display:"block",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.08em"}}>Asistentes (correos separados por coma)</label><input type="text" value={newForm.attendees} onChange={e=>setNewForm(p=>({...p,attendees:e.target.value}))} placeholder="ana@alcsa.com, pedro@alcsa.com..." style={{...S.inpTxt,width:"100%",boxSizing:"border-box"}}/></div>
            <div style={{marginBottom:10}}><label style={{fontSize:10,color:C.t2,fontWeight:600,display:"block",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.08em"}}>Enlace Teams (opcional)</label><input type="url" value={newForm.teamsUrl} onChange={e=>setNewForm(p=>({...p,teamsUrl:e.target.value}))} placeholder="https://teams.microsoft.com/..." style={{...S.inpTxt,width:"100%",boxSizing:"border-box"}}/></div>
            <div style={{marginBottom:16}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><label style={{fontSize:10,color:C.t2,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.08em"}}>Agenda</label><button style={{...S.btnSm(false),fontSize:10}} onClick={()=>setNewForm(p=>({...p,agenda:[...p.agenda,""]}))}>+ Punto</button></div>
              {newForm.agenda.map((a,i)=>(
                <div key={i} style={{display:"flex",gap:6,alignItems:"center",marginBottom:6}}>
                  <span style={{fontSize:10,color:C.t3,width:18,textAlign:"right",flexShrink:0}}>{i+1}.</span>
                  <input type="text" value={a} onChange={e=>setNewForm(p=>({...p,agenda:p.agenda.map((x,idx)=>idx===i?e.target.value:x)}))} placeholder={`Punto ${i+1}...`} style={{...S.inpTxt,flex:1}}/>
                  {newForm.agenda.length>1&&<button onClick={()=>setNewForm(p=>({...p,agenda:p.agenda.filter((_,idx)=>idx!==i)}))} style={{background:"none",border:"none",color:C.t3,cursor:"pointer",fontSize:16}}>×</button>}
                </div>
              ))}
            </div>
            <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
              <button style={S.btn(false)} onClick={()=>setShowNew(false)}>Cancelar</button>
              <button style={S.btn(true)} onClick={createMeeting}>Crear Reunión</button>
            </div>
          </div>
        </div>
      )}

      <div style={{display:"grid",gridTemplateColumns:"290px 1fr",gap:14,marginBottom:14}}>
        {/* Calendario */}
        <div style={S.card}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <button onClick={()=>setCurMonth(p=>p.month===0?{year:p.year-1,month:11}:{...p,month:p.month-1})} style={{...S.btnSm(false),padding:"3px 10px",fontSize:15}}>‹</button>
            <div style={{fontSize:13,fontWeight:700,color:C.t1}}>{MES_NOMBRE[curMonth.month]} {curMonth.year}</div>
            <button onClick={()=>setCurMonth(p=>p.month===11?{year:p.year+1,month:0}:{...p,month:p.month+1})} style={{...S.btnSm(false),padding:"3px 10px",fontSize:15}}>›</button>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",marginBottom:4}}>
            {DIA_NOMBRE.map(d=><div key={d} style={{textAlign:"center",fontSize:9,color:C.t3,fontWeight:700,padding:"2px 0"}}>{d}</div>)}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:2}}>
            {Array.from({length:fd}).map((_,i)=><div key={"e"+i}/>)}
            {Array.from({length:days}).map((_,i)=>{
              const day=i+1,ds=`${curMonth.year}-${String(curMonth.month+1).padStart(2,"0")}-${String(day).padStart(2,"0")}`,dm=meetingsByDay[day]||[];
              const isToday=ds===today,hasSel=dm.some(m=>m.id===selM?.id);
              return (
                <div key={day} onClick={()=>dm.length>0&&setSelM(dm[0])}
                  style={{textAlign:"center",padding:"5px 2px",borderRadius:6,cursor:dm.length>0?"pointer":"default",background:isToday?C.accent:hasSel?C.accentLt:dm.length>0?C.bg:"transparent",border:hasSel?`1px solid ${C.accent}`:"1px solid transparent"}}>
                  <div style={{fontSize:11,fontWeight:isToday||dm.length>0?700:400,color:isToday?"#fff":dm.length>0?C.accent:C.t3}}>{day}</div>
                  {dm.length>0&&<div style={{display:"flex",justifyContent:"center",gap:2,marginTop:2}}>{dm.slice(0,3).map(m=><div key={m.id} style={{width:5,height:5,borderRadius:"50%",background:STATUS_COLOR[m.status]||C.amber}}/>)}</div>}
                </div>
              );
            })}
          </div>
          <div style={{marginTop:12,borderTop:`1px solid ${C.borderLt}`,paddingTop:10,display:"flex",flexWrap:"wrap",gap:10}}>
            {Object.entries(STATUS_COLOR).map(([s,c])=><span key={s} style={{display:"flex",alignItems:"center",gap:4,fontSize:9,color:C.t3}}><span style={{width:7,height:7,borderRadius:"50%",background:c,display:"inline-block"}}/>{s.charAt(0).toUpperCase()+s.slice(1)}</span>)}
          </div>
        </div>

        {/* Lista del mes */}
        <div style={{...S.card,overflowY:"auto",maxHeight:380}}>
          <div style={S.hdr}>{MES_NOMBRE[curMonth.month]} {curMonth.year} — {meetings.filter(m=>m.date.startsWith(monthStr)).length} reuniones</div>
          {meetings.filter(m=>m.date.startsWith(monthStr)).length===0?(
            <div style={{textAlign:"center",color:C.t3,padding:"28px 0",fontSize:11}}>Sin reuniones programadas este mes</div>
          ):(
            meetings.filter(m=>m.date.startsWith(monthStr)).sort((a,b)=>a.date.localeCompare(b.date)).map(m=>(
              <div key={m.id} onClick={()=>setSelM(m)} style={{padding:"11px 14px",marginBottom:8,borderRadius:8,border:`1px solid ${selM?.id===m.id?C.accent:C.border}`,background:selM?.id===m.id?C.accentLt:C.bg,cursor:"pointer"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                  <div style={{flex:1}}>
                    <div style={{fontSize:12,fontWeight:700,color:C.t1,marginBottom:3}}>{m.title}</div>
                    <div style={{fontSize:10,color:C.t2}}>📅 {m.date} · 🕐 {m.time} ({m.duration} min)</div>
                    <div style={{fontSize:10,color:C.t2,marginTop:2}}>👥 {m.attendees.split(",").length} asistentes · {m.canal}</div>
                  </div>
                  <span style={{...S.tag(STATUS_COLOR[m.status]||C.amber),marginLeft:8,flexShrink:0,fontSize:9}}>{m.status}</span>
                </div>
                {m.agenda.length>0&&<div style={{marginTop:7,display:"flex",gap:4,flexWrap:"wrap"}}>{m.agenda.slice(0,3).map((a,i)=><span key={i} style={{fontSize:9,padding:"2px 6px",borderRadius:4,background:C.border,color:C.t2}}>{a.slice(0,30)}{a.length>30?"…":""}</span>)}{m.agenda.length>3&&<span style={{fontSize:9,color:C.t3}}>+{m.agenda.length-3} más</span>}</div>}
              </div>
            ))
          )}
        </div>
      </div>

      {/* Panel de detalle */}
      {selM&&(
        <div style={S.card}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:16,flexWrap:"wrap",gap:10}}>
            <div>
              <div style={{fontSize:14,fontWeight:700,color:C.t1}}>{selM.title}</div>
              <div style={{fontSize:10,color:C.t2,marginTop:4}}>📅 {selM.date} · 🕐 {selM.time} ({selM.duration} min) · 🏷 {selM.type}</div>
              <div style={{fontSize:10,color:C.t2,marginTop:2}}>📡 Canal: {selM.canal} · 👥 {selM.attendees}</div>
            </div>
            <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
              <span style={S.tag(STATUS_COLOR[selM.status]||C.amber)}>{selM.status}</span>
              <select value={selM.status} onChange={e=>updSel({status:e.target.value})} style={{...S.sel,fontSize:10}}>
                <option value="pendiente">Pendiente</option><option value="realizada">Realizada</option><option value="cancelada">Cancelada</option>
              </select>
              {selM.teamsUrl&&<button style={{...S.btn(false,false),fontSize:10}} onClick={()=>window.open(selM.teamsUrl,"_blank")}>🔗 Abrir en Teams</button>}
              <button style={S.btnDng} onClick={()=>{ const u=meetings.filter(m=>m.id!==selM.id); save(u); setSelM(null); }}>✕ Eliminar</button>
            </div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 2fr",gap:20}}>
            <div>
              <div style={S.hdr}>Agenda</div>
              {selM.agenda.map((a,i)=><div key={i} style={{display:"flex",alignItems:"flex-start",gap:8,padding:"7px 10px",marginBottom:5,background:C.bg,borderRadius:6}}><span style={{fontSize:10,color:C.t3,fontWeight:700,minWidth:18,marginTop:1}}>{i+1}.</span><span style={{fontSize:11,color:C.t1,lineHeight:1.4}}>{a}</span></div>)}
              {selM.teamsUrl&&<div style={{marginTop:12,padding:"10px 12px",background:C.accentLt,borderRadius:8,border:`1px solid ${C.accent}33`}}><div style={{fontSize:10,color:C.t2,fontWeight:600,marginBottom:6}}>🔗 Reunión Teams</div><button onClick={()=>window.open(selM.teamsUrl,"_blank")} style={{...S.btn(true),fontSize:10,width:"100%"}}>Unirse a la reunión</button></div>}
            </div>
            <div>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                <div style={S.hdr}>Minuta — generada por IA</div>
                <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                  <button style={{...S.btn(true),fontSize:10,opacity:genLoading?0.65:1}} onClick={generateMinuta} disabled={genLoading}>{genLoading?"⟳ Generando…":"✦ Generar con IA"}</button>
                  {selM.minuta&&<>
                    <button style={S.btn(false,copied)} onClick={()=>{ navigator.clipboard?.writeText(selM.minuta); setCopied(true); setTimeout(()=>setCopied(false),2000); }}>{copied?"✓ Copiado":"⎘ Copiar"}</button>
                    <a href={`mailto:${encodeURIComponent(selM.attendees||"")}?subject=${encodeURIComponent(`Minuta — ${selM.title} — ${selM.date}`)}&body=${encodeURIComponent(selM.minuta)}`} style={{...S.btn(false,false),fontSize:10,textDecoration:"none",display:"inline-flex",alignItems:"center"}}>✉ Enviar correo</a>
                  </>}
                </div>
              </div>
              {selM.minuta?(
                <textarea value={selM.minuta} onChange={e=>updSel({minuta:e.target.value})} style={{width:"100%",height:240,background:C.bg,border:`1px solid ${C.border}`,borderRadius:8,padding:14,fontFamily:"'Inter','Segoe UI',sans-serif",fontSize:11,lineHeight:1.65,color:C.t1,resize:"vertical",boxSizing:"border-box",outline:"none"}}/>
              ):(
                <div style={{height:240,background:C.bg,border:`2px dashed ${C.border}`,borderRadius:8,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:8,cursor:"pointer"}} onClick={generateMinuta}>
                  <div style={{fontSize:30,opacity:0.4}}>✦</div>
                  <div style={{fontSize:12,color:C.t2,fontWeight:600}}>Genera la minuta con IA</div>
                  <div style={{fontSize:10,color:C.t3,textAlign:"center",maxWidth:240,lineHeight:1.5}}>Claude redacta una minuta profesional basada en la agenda.</div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── MODELOS VIEW ──────────────────────────────────────────────────────────
function ModelsView({ models, setModels, editingModel, setEditingModel, C, S }) {
  const cur=models.find(m=>m.id===editingModel)||models[0];
  const [script,setScript]=useState(cur?.script||"");
  useEffect(()=>{ const m=models.find(x=>x.id===editingModel); if(m) setScript(m.script); },[editingModel,models]);
  const save=()=>{ const u=models.map(m=>m.id===cur.id?{...m,script}:m); setModels(u); window.storage?.set("dr2-models",JSON.stringify(u)).catch(()=>{}); };
  const exportTemplate=()=>{ const headers=["Canal","Cadena","Item","Material","Categoría","Sub Categoría","Marca"]; [2023,2024,2025].forEach(yr=>MF.forEach(m=>headers.push(`Real_${m}_${yr}`))); const sample=[["MODERNO","1014602-UNISUPER","60000001","ARROZ PREC GALLO DORADO 5U 5LB","Arroz","Arroz Precocido","Gallo Dorado",...Array(36).fill(0).map((_,i)=>Math.round(18000+Math.sin(i/2)*3000+i*50))]]; const ws=XLSX.utils.aoa_to_sheet([headers,...sample]); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Ventas_Reales"); XLSX.writeFile(wb,"Plantilla_Ventas_Modelos.xlsx"); };
  return (
    <div>
      <div style={{fontSize:17,fontWeight:700,color:C.t1,marginBottom:4}}>Modelos Estadísticos</div>
      <div style={{fontSize:11,color:C.t2,marginBottom:16}}>Edita scripts Python · activa/desactiva · agrega nuevos</div>
      <div style={{...S.card,background:C.amber+"0d",border:`1px solid ${C.amber}44`,marginBottom:14}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div><div style={{fontSize:12,fontWeight:700,color:C.amber,marginBottom:6}}>⚙ Cómo correr los modelos</div><div style={{fontSize:11,color:C.t2,lineHeight:1.7}}>1. Descarga la plantilla → 2. Rellena ventas históricas (2023·2024·2025) → 3. Ejecuta <code style={{background:C.bg,padding:"1px 5px",borderRadius:3,fontSize:10,color:C.accent}}>python run_models.py --input Plantilla.xlsx</code> → 4. Importa el resultado con ⬆ Importar Excel</div></div>
          <button style={{...S.btn(true),marginLeft:16,flexShrink:0}} onClick={exportTemplate}>⬇ Plantilla Excel</button>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"210px 1fr",gap:14}}>
        <div>
          {models.map(m=>(
            <div key={m.id} style={{...S.card,cursor:"pointer",borderColor:m.id===cur?.id?C.accent:C.border,borderWidth:m.id===cur?.id?2:1,marginBottom:8,padding:"10px 12px"}} onClick={()=>setEditingModel(m.id)}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
                <span style={{fontSize:11,fontWeight:700,color:m.id===cur?.id?C.accent:C.t1}}>{m.name}</span>
                <button onClick={e=>{ e.stopPropagation(); setModels(p=>p.map(x=>x.id===m.id?{...x,active:!x.active}:x)); }} style={{padding:"2px 7px",borderRadius:4,fontSize:9,border:"none",cursor:"pointer",background:m.active?"#d1fae5":"#f1f5f9",color:m.active?C.green:C.t3,fontWeight:700}}>{m.active?"● ON":"○ OFF"}</button>
              </div>
              {m.mape!=null&&<div style={{fontSize:10,color:C.green,fontWeight:600}}>MAPE 6m: {m.mape}%</div>}
              <div style={{fontSize:10,color:C.t2,marginTop:3}}>{m.description}</div>
            </div>
          ))}
          <button style={{...S.btn(false,false),width:"100%",padding:"8px"}} onClick={()=>{ const id="m"+Date.now(); setModels(p=>[...p,{id,name:"Nuevo Modelo",active:false,mape:null,description:"",script:"# script\n"}]); setEditingModel(id); }}>+ Nuevo modelo</button>
        </div>
        <div style={S.card}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <div><div style={{fontSize:12,fontWeight:700,color:C.t1}}>{cur?.name}</div><div style={{fontSize:11,color:C.t2}}>{cur?.description}</div></div>
            <div style={{display:"flex",gap:8,alignItems:"center"}}>{cur?.mape!=null&&<span style={S.tag(C.green)}>MAPE {cur.mape}%</span>}<button style={S.btn(true)} onClick={save}>Guardar Script</button></div>
          </div>
          <textarea value={script} onChange={e=>setScript(e.target.value)} style={{width:"100%",height:340,background:"#1e2940",color:"#a5f3fc",border:"none",borderRadius:8,padding:14,fontFamily:"'JetBrains Mono','Fira Code',monospace",fontSize:11,lineHeight:1.65,resize:"vertical",boxSizing:"border-box",outline:"none"}}/>
          <div style={{marginTop:8,fontSize:10,color:C.t3}}>Dependencias: statsmodels · prophet · pandas · numpy · scikit-learn</div>
        </div>
      </div>
    </div>
  );
}

// ── HISTORIAL VIEW ─────────────────────────────────────────────────────────
function HistorialView({ log, setLog, sessions, setSessions, activeSession, setActiveSession, C, S }) {
  const [filterSess,setFilterSess]=useState("ALL");
  const [showNewDlg,setShowNewDlg]=useState(false);
  const [confirmDel,setConfirmDel]=useState(null);
  const [newMonth,setNewMonth]=useState(4);
  const [newName,setNewName]=useState("");
  const filtered=log.filter(c=>filterSess==="ALL"||c.session===filterSess);
  const createSession=()=>{ const editableFrom=newMonth+2; const id="rev-"+Date.now(); const name=newName.trim()||`Revisión ${MES_NOMBRE[newMonth]} 2026`; const sess={id,name,date:new Date().toISOString().slice(0,10),status:"active",month:newMonth,editableFrom}; setSessions(p=>{ const ns=[...p,sess]; window.storage?.set("dr2-sess",JSON.stringify(ns)).catch(()=>{}); return ns; }); setActiveSession(id); setShowNewDlg(false); setNewName(""); };
  const deleteSession=(id)=>{ setSessions(p=>{ const ns=p.filter(s=>s.id!==id); window.storage?.set("dr2-sess",JSON.stringify(ns)).catch(()=>{}); return ns; }); setLog(p=>{ const nl=p.filter(c=>c.session!==id); window.storage?.set("dr2-log",JSON.stringify(nl)).catch(()=>{}); return nl; }); if(activeSession===id) setActiveSession(sessions.find(s=>s.id!==id)?.id||""); if(filterSess===id) setFilterSess("ALL"); setConfirmDel(null); };
  const sessToDel=sessions.find(s=>s.id===confirmDel);
  return (
    <div>
      <div style={{fontSize:17,fontWeight:700,color:C.t1,marginBottom:4}}>Historial de Cambios</div>
      <div style={{fontSize:11,color:C.t2,marginBottom:16}}>Auditoría completa · ajustes al S&OP por sesión</div>
      {confirmDel&&(
        <div style={S.modal}><div style={{...S.dlg,maxWidth:360}}><div style={{fontSize:14,fontWeight:700,color:C.red,marginBottom:8}}>⚠ Eliminar sesión</div><div style={{fontSize:11,color:C.t2,marginBottom:16}}>¿Eliminar <strong>{sessToDel?.name}</strong>? Se borrarán <strong>{log.filter(c=>c.session===confirmDel).length} cambios</strong>. No puede deshacerse.</div><div style={{display:"flex",gap:8,justifyContent:"flex-end"}}><button style={S.btn(false)} onClick={()=>setConfirmDel(null)}>Cancelar</button><button style={{...S.btn(true),background:C.red}} onClick={()=>deleteSession(confirmDel)}>Sí, eliminar</button></div></div></div>
      )}
      {showNewDlg&&(
        <div style={S.modal} onClick={e=>e.target===e.currentTarget&&setShowNewDlg(false)}>
          <div style={S.dlg}>
            <div style={{fontSize:14,fontWeight:700,color:C.t1,marginBottom:4}}>Nueva sesión de revisión</div>
            <div style={{fontSize:11,color:C.t2,marginBottom:16}}>Define el mes de revisión para habilitar los meses editables.</div>
            <div style={{marginBottom:12}}><label style={{fontSize:10,color:C.t2,fontWeight:600,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:"0.08em"}}>Mes de revisión</label><select style={{...S.sel,width:"100%",padding:"8px 10px",fontSize:12}} value={newMonth} onChange={e=>setNewMonth(+e.target.value)}>{MF.map((m,i)=><option key={i} value={i}>{m}</option>)}</select></div>
            <div style={{marginBottom:14,padding:"10px 12px",background:C.accentLt,borderRadius:8,border:`1px solid ${C.border}`}}>
              <div style={{fontSize:10,color:C.t2,fontWeight:600,marginBottom:6,textTransform:"uppercase",letterSpacing:"0.08em"}}>Meses editables en el rolling</div>
              <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>{MO.map((m,i)=><span key={m} style={{padding:"3px 8px",borderRadius:6,fontSize:10,fontWeight:600,background:i>=newMonth+2?C.green+"22":C.borderLt,color:i>=newMonth+2?C.green:C.t3}}>{m}</span>)}</div>
              <div style={{fontSize:10,color:C.t2,marginTop:6}}>Editables: <strong style={{color:C.green}}>{MO.slice(newMonth+2).join(" · ")}</strong></div>
            </div>
            <div style={{marginBottom:16}}><label style={{fontSize:10,color:C.t2,fontWeight:600,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:"0.08em"}}>Nombre</label><input type="text" placeholder={`Revisión ${MES_NOMBRE[newMonth]} 2026`} value={newName} onChange={e=>setNewName(e.target.value)} style={{...S.inpTxt,width:"100%",boxSizing:"border-box"}}/></div>
            <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}><button style={S.btn(false)} onClick={()=>setShowNewDlg(false)}>Cancelar</button><button style={S.btn(true)} onClick={createSession}>Crear Revisión</button></div>
          </div>
        </div>
      )}
      <div style={{display:"grid",gridTemplateColumns:"204px 1fr",gap:14}}>
        <div>
          <div style={S.hdr}>Sesiones</div>
          {sessions.map(s=>(
            <div key={s.id} style={{...S.card,padding:"10px 12px",marginBottom:8,borderLeft:`3px solid ${s.id===activeSession?C.amber:s.status==="active"?C.green:C.border}`,cursor:"pointer"}} onClick={()=>setFilterSess(s.id)}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                <div style={{fontSize:11,fontWeight:700,color:s.id===filterSess?C.accent:C.t1,flex:1,marginRight:6}}>{s.name}</div>
                <button style={S.btnDng} onClick={e=>{ e.stopPropagation(); setConfirmDel(s.id); }}>✕</button>
              </div>
              <div style={{fontSize:10,color:C.t3,marginTop:2}}>{s.date}</div>
              {s.editableFrom!=null&&<div style={{fontSize:9,color:C.t2,marginTop:3}}>Rolling: <strong style={{color:C.green}}>{MO.slice(s.editableFrom).join("·")}</strong></div>}
              <div style={{display:"flex",gap:6,alignItems:"center",marginTop:5}}>
                <span style={{fontSize:9,padding:"2px 7px",borderRadius:10,background:s.status==="active"?"#d1fae5":"#f1f5f9",color:s.status==="active"?C.green:C.t3,fontWeight:700}}>{s.status==="active"?"● Activa":"✓ Cerrada"}</span>
                <span style={{fontSize:10,color:C.t3}}>{log.filter(c=>c.session===s.id).length} cambios</span>
              </div>
            </div>
          ))}
          <button style={{...S.btn(true),width:"100%",padding:"8px",marginTop:2}} onClick={()=>setShowNewDlg(true)}>+ Nueva Sesión</button>
          {filterSess!=="ALL"&&<button style={{...S.btn(false,false),width:"100%",marginTop:6,padding:"7px",fontSize:10}} onClick={()=>setFilterSess("ALL")}>Ver Todos</button>}
        </div>
        <div style={S.card}>
          <div style={S.hdr}>Cambios registrados ({filtered.length})</div>
          {filtered.length===0?(
            <div style={{textAlign:"center",color:C.t3,padding:"48px 0",fontSize:12}}><div style={{fontSize:32,marginBottom:8}}>📋</div>Sin cambios en esta sesión</div>
          ):(
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
              <thead><tr style={{background:C.bg}}>{["Hora","SKU","Material","Canal","Mes","Valor ant.","Nuevo valor","Δ Cambio","Comentario"].map(h=><th key={h} style={{padding:"8px 10px",color:C.t2,textAlign:["Material","Comentario"].includes(h)?"left":"right",fontSize:9,fontWeight:700,letterSpacing:"0.08em",borderBottom:`1px solid ${C.border}`}}>{h}</th>)}</tr></thead>
              <tbody>
                {filtered.map((c,i)=>(
                  <tr key={c.id} style={{borderBottom:`1px solid ${C.borderLt}`,background:i%2?C.bg:C.card}}>
                    <td style={{padding:"7px 10px",color:C.t3,textAlign:"right"}}>{new Date(c.timestamp).toLocaleTimeString("es-GT")}</td>
                    <td style={{padding:"7px 10px",color:C.accent,fontWeight:700,textAlign:"right"}}>{c.item}</td>
                    <td style={{padding:"7px 10px",color:C.t2}}>{(c.material||"").slice(0,22)}</td>
                    <td style={{padding:"7px 10px",color:C.t2,textAlign:"right"}}>{c.canal}</td>
                    <td style={{padding:"7px 10px",textAlign:"right"}}><span style={S.tag(C.amber)}>{c.month}</span></td>
                    <td style={{padding:"7px 10px",textAlign:"right",color:C.t3}}>{fmtK(c.oldValue)}</td>
                    <td style={{padding:"7px 10px",textAlign:"right",color:C.t1,fontWeight:700}}>{fmtK(c.newValue)}</td>
                    <td style={{padding:"7px 10px",textAlign:"right"}}><span style={{...S.tag(clr(c.delta)),fontWeight:700}}>{c.delta>0?"+":""}{fmtK(c.delta)}</span></td>
                    <td style={{padding:"7px 10px",color:C.t2,fontStyle:"italic",maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.comment||"—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>
    </div>
  );
}

// ── CONFIG VIEW ───────────────────────────────────────────────────────────
function ConfigView({ cfg, setCfg, C, S }) {
  const [local,setLocal]=useState({...cfg});
  const logoRef=useRef(); const brandRef=useRef(); const [brandName,setBrandName]=useState("");
  const save=()=>{ setCfg(local); window.storage?.set("dr2-cfg",JSON.stringify(local)).catch(()=>{}); };
  const reset=()=>{ const def={theme:"ejecutivo",customColors:{},logoEmpresa:null,logosMarca:{}}; setLocal(def); setCfg(def); window.storage?.set("dr2-cfg",JSON.stringify(def)).catch(()=>{}); };
  const handleLogo=(file)=>{ if(!file) return; const r=new FileReader(); r.onload=e=>setLocal(p=>({...p,logoEmpresa:e.target.result})); r.readAsDataURL(file); };
  const handleBrand=(file)=>{ if(!file||!brandName.trim()) return; const r=new FileReader(); r.onload=e=>setLocal(p=>({...p,logosMarca:{...p.logosMarca,[brandName.trim()]:e.target.result}})); r.readAsDataURL(file); setBrandName(""); };
  const updColor=(k,v)=>setLocal(p=>({...p,customColors:{...p.customColors,[k]:v}}));
  const preC=makeC(local.theme,local.customColors);
  const Section=({title,children})=><div style={{...S.card,marginBottom:16}}><div style={S.hdr}>{title}</div>{children}</div>;
  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",marginBottom:20}}>
        <div><div style={{fontSize:17,fontWeight:700,color:C.t1}}>Configuración</div><div style={{fontSize:11,color:C.t2,marginTop:2}}>Personaliza tema visual, colores y logos</div></div>
        <div style={{display:"flex",gap:8}}><button style={S.btn(false,false)} onClick={reset}>↺ Restaurar</button><button style={S.btn(true)} onClick={save}>Guardar cambios</button></div>
      </div>
      <Section title="Tema visual">
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:14}}>
          {Object.entries(THEMES).map(([key,theme])=>(
            <div key={key} onClick={()=>setLocal(p=>({...p,theme:key}))} style={{borderRadius:8,border:`2px solid ${local.theme===key?C.accent:C.border}`,overflow:"hidden",cursor:"pointer",boxShadow:local.theme===key?`0 0 0 3px ${C.accent}33`:"none"}}>
              <div style={{display:"flex",height:52}}>
                <div style={{width:24,background:theme.sidebar}}/>
                <div style={{flex:1,background:theme.bg,padding:"6px 8px"}}>
                  <div style={{height:5,borderRadius:2,background:theme.border,marginBottom:4}}/>
                  <div style={{height:5,borderRadius:2,background:theme.accent,width:"70%",marginBottom:4}}/>
                  <div style={{height:5,borderRadius:2,background:theme.amber,width:"50%"}}/>
                </div>
              </div>
              <div style={{padding:"6px 10px",background:local.theme===key?C.accentLt:C.bg,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <span style={{fontSize:11,fontWeight:600,color:local.theme===key?C.accent:C.t2}}>{theme.label}</span>
                {local.theme===key&&<span style={{color:C.accent}}>✓</span>}
              </div>
            </div>
          ))}
        </div>
      </Section>
      <Section title="Colores personalizados">
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16}}>
          {[["Color acento","accent",preC.accent],["Color sidebar","sidebar",preC.sidebar],["Color destaque","amber",preC.amber]].map(([label,key,cur])=>(
            <div key={key}>
              <div style={{fontSize:10,color:C.t2,fontWeight:600,marginBottom:8,textTransform:"uppercase",letterSpacing:"0.08em"}}>{label}</div>
              <div style={{display:"flex",alignItems:"center",gap:10}}>
                <input type="color" value={local.customColors?.[key]||cur} onChange={e=>updColor(key,e.target.value)} style={{width:40,height:40,borderRadius:8,border:`1px solid ${C.border}`,cursor:"pointer",padding:2,background:"transparent"}}/>
                <div>
                  <div style={{width:90,height:16,borderRadius:4,background:local.customColors?.[key]||cur,marginBottom:4}}/>
                  <div style={{fontSize:10,color:C.t3,fontFamily:"monospace"}}>{local.customColors?.[key]||cur}</div>
                </div>
                {local.customColors?.[key]&&<button onClick={()=>setLocal(p=>{ const cc={...p.customColors}; delete cc[key]; return {...p,customColors:cc}; })} style={{...S.btnSm(false),fontSize:10,color:C.red}}>↺</button>}
              </div>
            </div>
          ))}
        </div>
      </Section>
      <Section title="Logos">
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
          <div>
            <div style={{fontSize:11,color:C.t2,fontWeight:600,marginBottom:10}}>Logo de empresa</div>
            <input type="file" ref={logoRef} accept="image/*" style={{display:"none"}} onChange={e=>handleLogo(e.target.files[0])}/>
            <div style={{display:"flex",gap:8,alignItems:"center"}}>
              {local.logoEmpresa?<img src={local.logoEmpresa} alt="logo" style={{height:40,objectFit:"contain",borderRadius:6,border:`1px solid ${C.border}`,padding:4,background:C.bg}}/>:<div style={{width:80,height:40,border:`2px dashed ${C.border}`,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center"}}><span style={{fontSize:9,color:C.t3}}>Sin logo</span></div>}
              <button style={S.btn(false,false)} onClick={()=>logoRef.current.click()}>Cargar imagen</button>
              {local.logoEmpresa&&<button style={S.btnDng} onClick={()=>setLocal(p=>({...p,logoEmpresa:null}))}>✕</button>}
            </div>
          </div>
          <div>
            <div style={{fontSize:11,color:C.t2,fontWeight:600,marginBottom:10}}>Logos de marcas</div>
            <input type="file" ref={brandRef} accept="image/*" style={{display:"none"}} onChange={e=>handleBrand(e.target.files[0])}/>
            <div style={{display:"flex",gap:6,marginBottom:10}}>
              <input type="text" placeholder="Nombre exacto de la marca..." value={brandName} onChange={e=>setBrandName(e.target.value)} style={{...S.inpTxt,width:160,flex:"none"}}/>
              <button style={S.btn(false,!!brandName.trim())} onClick={()=>brandName.trim()&&brandRef.current.click()}>Cargar logo</button>
            </div>
            <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
              {Object.entries(local.logosMarca||{}).map(([name,src])=>(
                <div key={name} style={{display:"flex",alignItems:"center",gap:5,padding:"4px 8px",background:C.bg,borderRadius:6,border:`1px solid ${C.border}`}}>
                  <img src={src} alt={name} style={{height:20,objectFit:"contain"}}/>
                  <span style={{fontSize:10,color:C.t2}}>{name}</span>
                  <button onClick={()=>setLocal(p=>{ const lm={...p.logosMarca}; delete lm[name]; return {...p,logosMarca:lm}; })} style={{background:"none",border:"none",cursor:"pointer",color:C.t3,fontSize:12,lineHeight:1}}>✕</button>
                </div>
              ))}
              {Object.keys(local.logosMarca||{}).length===0&&<span style={{fontSize:10,color:C.t3,fontStyle:"italic"}}>Sin logos cargados</span>}
            </div>
          </div>
        </div>
      </Section>
    </div>
  );
}

// ── TUTORIAL VIEW ─────────────────────────────────────────────────────────
function TutorialView({ C, S }) {
  const [tab,setTab]=useState("roadmap");
  const Code=({children})=><pre style={{background:"#1e2940",color:"#a5f3fc",borderRadius:8,padding:16,fontSize:10,lineHeight:1.65,overflow:"auto",fontFamily:"'JetBrains Mono','Fira Code',monospace",whiteSpace:"pre"}}>{children}</pre>;
  const Step=({n,title,children})=>(
    <div style={{display:"flex",gap:14,marginBottom:20}}>
      <div style={{width:28,height:28,borderRadius:"50%",background:C.accent,color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:800,flexShrink:0,marginTop:2}}>{n}</div>
      <div style={{flex:1}}><div style={{fontSize:13,fontWeight:700,color:C.t1,marginBottom:6}}>{title}</div><div style={{fontSize:11,color:C.t2,lineHeight:1.65}}>{children}</div></div>
    </div>
  );
  const TABS=[{id:"roadmap",label:"🗺 Hoja de ruta"},{id:"sharepoint",label:"☁ SharePoint + Power Automate"},{id:"schema",label:"🗄 Base de datos"},{id:"deploy",label:"🚀 Despliegue"},{id:"sop",label:"📋 S&OP completo"}];
  const SOPGaps=[
    {icon:"📥",title:"Input comercial pre-reunión",desc:"Formulario donde ventas carga su upside/downside por SKU antes de la reunión. Con tipo de evento, fecha e impacto estimado. Así la reunión tiene insumos preparados."},
    {icon:"🎯",title:"One Number / Consenso",desc:"El forecast que sale de la reunión se congela como versión oficial. Permite comparar qué prometiste en abril vs qué vendiste en abril real."},
    {icon:"📊",title:"FA / BIAS / MAPE histórico",desc:"Dashboard de accuracy del proceso. ¿Tu revisión de abril mejoró respecto a marzo? ¿Qué canal tiene más error?"},
    {icon:"📝",title:"Registro estructurado de supuestos",desc:"'Subí Gallo Dorado porque hay promoción en UNISUPER semana 3 de junio +15%.' Necesita: tipo de evento, fecha, SKU, canal, impacto % y responsable."},
    {icon:"✅",title:"Tracker de acuerdos de reunión",desc:"Los compromisos de la minuta necesitan su propio tablero con responsable, fecha límite y semáforo de cumplimiento."},
    {icon:"📅",title:"Calendario de eventos de demanda",desc:"Promociones, feriados, cambios de precio, lanzamientos. El estadístico no los captura — necesitan un registro que el planner consulte al hacer el rolling."},
  ];
  return (
    <div>
      <div style={{marginBottom:20}}><div style={{fontSize:17,fontWeight:700,color:C.t1}}>📖 Guía de implementación</div><div style={{fontSize:11,color:C.t2,marginTop:2}}>De artefacto en Claude → app real multi-usuario con Microsoft 365</div></div>
      <div style={{display:"flex",gap:4,marginBottom:20,flexWrap:"wrap"}}>{TABS.map(t=><button key={t.id} style={{...S.btn(false,tab===t.id),fontSize:11}} onClick={()=>setTab(t.id)}>{t.label}</button>)}</div>

      {tab==="roadmap"&&(
        <div>
          <div style={{...S.card,background:C.accentLt,border:`1px solid ${C.accent}33`,marginBottom:14}}>
            <div style={{fontSize:13,fontWeight:700,color:C.accent,marginBottom:8}}>Stack recomendado — ecosistema Microsoft</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10}}>
              {[["Frontend","React (este archivo)","Ya lo tienes listo","#059669"],["Datos","SharePoint Online","Excel en la nube de M365","#4f46e5"],["Automatización","Power Automate","Flujos sin código","#d97706"],["Despliegue","Azure Static Web Apps","URL en tu tenant M365","#0891b2"]].map(([rol,tech,desc,c])=>(
                <div key={rol} style={{padding:"12px 14px",background:C.card,borderRadius:8,border:`1px solid ${C.border}`,borderTop:`3px solid ${c}`}}>
                  <div style={{fontSize:9,color:C.t3,fontWeight:700,letterSpacing:"0.1em",textTransform:"uppercase",marginBottom:4}}>{rol}</div>
                  <div style={{fontSize:12,fontWeight:700,color:C.t1,marginBottom:3}}>{tech}</div>
                  <div style={{fontSize:10,color:C.t2}}>{desc}</div>
                </div>
              ))}
            </div>
          </div>
          <div style={S.card}>
            <div style={S.hdr}>Pasos — estimado total: 3-4 horas con IT</div>
            <Step n={1} title="Preparar los dos Excel en SharePoint (30 min)">Crear una biblioteca de documentos en SharePoint: <strong>Demand Review</strong>. Subir <code style={{background:C.bg,padding:"1px 5px",borderRadius:3,fontSize:10}}>DB_Demanda.xlsx</code> (tu DB actual) y <code style={{background:C.bg,padding:"1px 5px",borderRadius:3,fontSize:10}}>Cubo_Ventas.xlsx</code>. Asignar permisos: Planner = edición, Comerciales = edición solo en sus hojas, Gerencia = lectura.</Step>
            <Step n={2} title="Crear el flujo de Power Automate (60 min)">Flujo que se dispara cuando alguien modifica el Excel → extrae los cambios → los escribe en una tabla de log → notifica al planner por Teams. Ve a power.microsoft.com → Crear → Automatizado → "Cuando se modifica un archivo en SharePoint".</Step>
            <Step n={3} title="Exportar este archivo como proyecto React (15 min)">Descarga <code style={{background:C.bg,padding:"1px 5px",borderRadius:3,fontSize:10}}>demand_review.jsx</code> → <code style={{background:C.bg,padding:"1px 5px",borderRadius:3,fontSize:10}}>npm create vite@latest demand-review -- --template react</code> → copia el archivo como <code style={{background:C.bg,padding:"1px 5px",borderRadius:3,fontSize:10}}>src/App.jsx</code>.</Step>
            <Step n={4} title="Conectar la app a SharePoint vía Microsoft Graph (60 min)">Registrar la app en Azure AD → obtener credenciales → usar la Graph API para leer y escribir los Excel. El código exacto está en la pestaña ☁ SharePoint.</Step>
            <Step n={5} title="Desplegar en Azure Static Web Apps (30 min)">Push a GitHub → conectar con Azure Static Web Apps → en 5 min tienes tu URL. La autenticación es automática con cuentas de trabajo de ALCSA (Azure AD).</Step>
            <Step n={6} title="Publicar como pestaña en Teams (15 min)">Teams → canal de S&OP → + Agregar pestaña → Web site → pegar tu URL de Azure. Todo el equipo accede sin salir de Teams.</Step>
          </div>
        </div>
      )}

      {tab==="sharepoint"&&(
        <div>
          <div style={{...S.card,background:C.amber+"0d",border:`1px solid ${C.amber}44`,marginBottom:14}}>
            <div style={{fontSize:13,fontWeight:700,color:C.amber,marginBottom:4}}>Arquitectura SharePoint + Power Automate</div>
            <div style={{fontSize:11,color:C.t2,lineHeight:1.7}}>
              <strong>Excel 1 (DB_Demanda.xlsx)</strong> — vive en SharePoint. Comerciales editan sus inputs. Power Automate detecta cambios y notifica al planner.<br/>
              <strong>Excel 2 (Cubo_Ventas.xlsx)</strong> — también en SharePoint. El script Python de estadístico lo lee desde ahí y escribe el pronóstico de vuelta.<br/>
              <strong>La app React</strong> — lee ambos Excel vía Microsoft Graph API y los presenta. Los comerciales ven su vista, el planner ve todo.
            </div>
          </div>
          <div style={S.card}>
            <div style={S.hdr}>Código — conectar a SharePoint vía Microsoft Graph</div>
            <Code>{`// src/lib/sharepoint.js
import { PublicClientApplication } from "@azure/msal-browser";

const msalConfig = {
  auth: {
    clientId: "TU_APP_CLIENT_ID",       // Azure AD → App registrations
    authority: "https://login.microsoftonline.com/TU_TENANT_ID",
    redirectUri: window.location.origin,
  }
};
const msalInstance = new PublicClientApplication(msalConfig);

// Obtener token de acceso
async function getToken() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) {
    await msalInstance.loginPopup({ scopes: ["Files.ReadWrite", "User.Read"] });
  }
  const result = await msalInstance.acquireTokenSilent({
    scopes: ["Files.ReadWrite"],
    account: msalInstance.getAllAccounts()[0],
  });
  return result.accessToken;
}

// Leer el Excel de demanda desde SharePoint
export async function readDemandaExcel() {
  const token = await getToken();
  const siteId  = "TU_SITE_ID";   // SharePoint → Site Settings → Site ID
  const itemId  = "TU_ITEM_ID";   // ID del archivo en SharePoint

  // Obtener el archivo como binario
  const res = await fetch(
    \`https://graph.microsoft.com/v1.0/sites/\${siteId}/drive/items/\${itemId}/content\`,
    { headers: { Authorization: \`Bearer \${token}\` } }
  );
  const buffer = await res.arrayBuffer();
  
  // Parsear con SheetJS (ya está incluido en el artefacto)
  const wb = XLSX.read(buffer, { type: "array" });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets["DB"]);
  return rows.map(row => ({
    canal: row["Canal"] || "",
    item: String(row["Item"] || ""),
    cadena: row["Cadena"] || "",
    categoria: row["Categoría"] || "",
    subCategoria: row["Sub Categoría"] || "",
    marca: row["Marca"] || "",
    uom: row["Unidad de Medida"] || "FDO",
    figurasClave: row["Figura Clave"] || "",
    material: row["Material"] || "",
    clasificacion: row["Clasificación"] || "",
    meses: ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
            "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]
            .map(m => +(row[m] || 0)),
  }));
}

// Guardar un cambio de vuelta al Excel en SharePoint
export async function saveChangeToSharePoint(item, canal, mes, nuevoValor, comentario) {
  const token = await getToken();
  // Opción 1: Escribir directo al Excel via Graph (requiere Excel Online API)
  // Opción 2 (más simple): Agregar fila a una tabla de "Cambios Pendientes"
  // El flujo de Power Automate la procesa y actualiza el Excel principal

  const siteId = "TU_SITE_ID";
  const listId = "TU_LIST_ID"; // SharePoint List llamada "Cambios_Pendientes"

  await fetch(
    \`https://graph.microsoft.com/v1.0/sites/\${siteId}/lists/\${listId}/items\`,
    {
      method: "POST",
      headers: { Authorization: \`Bearer \${token}\`, "Content-Type": "application/json" },
      body: JSON.stringify({
        fields: {
          Item: item, Canal: canal, Mes: mes,
          NuevoValor: nuevoValor, Comentario: comentario,
          Estado: "pendiente", Timestamp: new Date().toISOString()
        }
      })
    }
  );
}`}</Code>
          </div>
          <div style={S.card}>
            <div style={S.hdr}>Flujo Power Automate — notificar al planner de cambios pendientes</div>
            <Code>{`Disparador: "Cuando se crea un elemento" en la lista "Cambios_Pendientes"

Acciones:
1. Condición: Estado == "pendiente"

2. Enviar mensaje en Teams al planner:
   Canal: S&OP Demand Review
   Mensaje: "📋 Cambio pendiente de aprobación
   ─────────────────────────
   SKU: [Item] — [Canal]
   Mes: [Mes]
   Valor propuesto: [NuevoValor]
   Comentario: [Comentario]
   
   ¿Apruebas este ajuste?"
   Botones: [✅ Aprobar] [❌ Rechazar]

3. Si Aprobar → actualizar la tabla DB en el Excel con el nuevo valor
4. Si Rechazar → marcar como rechazado + notificar al comercial
5. Actualizar campo Estado → "aprobado" o "rechazado"
6. Registrar en hoja "Historial" del Excel con timestamp`}</Code>
          </div>
        </div>
      )}

      {tab==="schema"&&(
        <div style={S.card}>
          <div style={S.hdr}>Estructura de los dos Excel en SharePoint</div>
          <Code>{`── DB_Demanda.xlsx ────────────────────────────────────────────────────
Hoja "DB" — tu estructura actual (28 columnas):
Canal | key | Comentario | Modificación | Item | Cadena | COD-CLIENTE
Categoría | Sub Categoría | Marca | Segmento | Unidad de Medida
Figura Clave | SKU&Descripción | Material | Clasificación
ENERO ... DICIEMBRE | Q1 | Q2 | Q3 | Q4 | Total

Hoja "Cambios_Pendientes" — nueva (para el flujo de aprobación):
Item | Canal | Mes | ValorAnterior | NuevoValor | Comentario
Solicitado_Por | Timestamp | Estado (pendiente/aprobado/rechazado)
Aprobado_Por | Fecha_Aprobacion

Hoja "Historial" — registro de todos los cambios aprobados:
Item | Canal | Mes | ValorAnterior | NuevoValor | Comentario
Aprobado_Por | Fecha | Sesion_Revision

── Cubo_Ventas.xlsx ───────────────────────────────────────────────────
Hoja "Ventas_Reales" — historial desde 2023:
Canal | Cadena | Item | Material | Categoría | Sub Categoría | Marca
Real_ENERO_2023 ... Real_DICIEMBRE_2025  (36 columnas de ventas reales)

Hoja "Estadistico_Output" — escrita por el script Python:
Canal | Item | Figura_Clave="Estadístico"
ENERO_2026 ... DICIEMBRE_2026 (pronóstico generado por los modelos)`}</Code>
        </div>
      )}

      {tab==="deploy"&&(
        <div style={S.card}>
          <div style={S.hdr}>Despliegue en Azure Static Web Apps — dentro del tenant de ALCSA</div>
          <Code>{`# 1. Crear el proyecto React
npm create vite@latest demand-review -- --template react
cd demand-review
npm install
npm install @azure/msal-browser xlsx recharts

# Copiar demand_review.jsx → src/App.jsx

# 2. Instalar dependencia de Graph API
npm install @azure/msal-browser

# 3. Build de producción
npm run build
# genera la carpeta /dist

# 4. Desplegar en Azure (pedir a IT que lo haga):
az staticwebapp create \\
  --name "demand-review-alcsa" \\
  --resource-group "rg-demand-review" \\
  --source "dist" \\
  --location "eastus" \\
  --branch "main" \\
  --app-location "/" \\
  --output-location "dist"

# 5. Configurar autenticación con Azure AD
# En Azure Portal → Static Web App → Authentication → Add identity provider
# Identity provider: Microsoft (Azure AD)
# Client ID: el mismo que usas en MSAL
# Resultado: solo cuentas @alcsa.com pueden acceder

# 6. Agregar como pestaña en Teams
# Teams → Canal S&OP → + → Website
# URL: https://demand-review-alcsa.azurestaticapps.net
# Nombre: Demand Review
# Todo el equipo accede sin salir de Teams`}</Code>
        </div>
      )}

      {tab==="sop"&&(
        <div>
          <div style={{...S.card,background:"#fffbf0",border:`1px solid ${C.amber}33`,marginBottom:14}}>
            <div style={{fontSize:13,fontWeight:700,color:C.amber,marginBottom:4}}>Lo que falta para un S&OP de clase mundial</div>
            <div style={{fontSize:11,color:C.t2}}>Este artefacto cubre el <strong>60%</strong> del proceso. Las piezas restantes transforman la herramienta en un proceso real de S&OP colaborativo.</div>
          </div>
          {SOPGaps.map((g,i)=>(
            <div key={i} style={{...S.card,padding:"14px 16px",marginBottom:0,display:"flex",gap:14,alignItems:"flex-start"}}>
              <div style={{fontSize:22,flexShrink:0,marginTop:1}}>{g.icon}</div>
              <div style={{flex:1}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                  <div style={{fontSize:13,fontWeight:700,color:C.t1,marginBottom:4}}>{g.title}</div>
                  <span style={{...S.tag(C.amber),fontSize:9,marginLeft:8,flexShrink:0}}>Por construir</span>
                </div>
                <div style={{fontSize:11,color:C.t2,lineHeight:1.6}}>{g.desc}</div>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ── MAIN APP ──────────────────────────────────────────────────────────────
const DEFAULT_CFG = { theme:"ejecutivo", customColors:{}, logoEmpresa:null, logosMarca:{} };

export default function App() {
  const [view,       setView]       = useState("dashboard");
  const [rows,       setRows]       = useState(buildDemo);
  const [uom,        setUom]        = useState("FDO");
  const [canal,      setCanal]      = useState("ALL");
  const [log,        setLog]        = useState([]);
  const [sessions,   setSessions]   = useState([
    {id:"rev-abr-26",name:"Revisión Abril 2026",date:"2026-04-05",status:"closed",month:3,editableFrom:5},
    {id:"rev-may-26",name:"Revisión Mayo 2026", date:"2026-05-08",status:"active",month:4,editableFrom:6},
  ]);
  const [activeSess, setActiveSess] = useState("rev-may-26");
  const [models,     setModels]     = useState(DEF_MODELS);
  const [editMod,    setEditMod]    = useState("prophet");
  const [selSKU,     setSelSKU]     = useState(null);
  const [cfg,        setCfg]        = useState(DEFAULT_CFG);

  const editableFrom=useMemo(()=>sessions.find(s=>s.id===activeSess)?.editableFrom??6,[sessions,activeSess]);
  const C=useMemo(()=>makeC(cfg.theme,cfg.customColors),[cfg.theme,cfg.customColors]);
  const S=useMemo(()=>makeS(C),[C]);

  useEffect(()=>{
    (async()=>{
      try{ const r=await window.storage?.get("dr2-log");    if(r) setLog(JSON.parse(r.value)); }catch(e){}
      try{ const r=await window.storage?.get("dr2-sess");   if(r) setSessions(JSON.parse(r.value)); }catch(e){}
      try{ const r=await window.storage?.get("dr2-models"); if(r) setModels(JSON.parse(r.value)); }catch(e){}
      try{ const r=await window.storage?.get("dr2-rows");   if(r) setRows(JSON.parse(r.value)); }catch(e){}
      try{ const r=await window.storage?.get("dr2-cfg");    if(r) setCfg({...DEFAULT_CFG,...JSON.parse(r.value)}); }catch(e){}
    })();
  },[]);

  const handleChange=useCallback((skuKey,mi,val,comment="",uomKey="FDO")=>{
    const [cn,item]=skuKey.split("||");
    setRows(prev=>{
      const sopRows=prev.filter(r=>r.canal===cn&&r.item===item&&r.uom===uomKey&&r.figurasClave==="S&OP");
      const oldAgg=sopRows.reduce((s,r)=>s+(r.meses[mi]||0),0);
      const newVal=Number(val),delta=Math.round(newVal-oldAgg);
      const entry={id:Date.now(),session:activeSess,timestamp:new Date().toISOString(),skuKey,item,canal:cn,uom:uomKey,month:MO[mi],monthIdx:mi,oldValue:Math.round(oldAgg),newValue:Math.round(newVal),delta,comment,material:prev.find(r=>r.item===item)?.material||""};
      setLog(p=>{ const nl=[entry,...p]; window.storage?.set("dr2-log",JSON.stringify(nl)).catch(()=>{}); return nl; });
      const curTotal=oldAgg;
      const updated=prev.map(r=>{ if(r.canal!==cn||r.item!==item||r.uom!==uomKey||r.figurasClave!=="S&OP") return r; const share=curTotal>0?(r.meses[mi]||0)/curTotal:1/Math.max(sopRows.length,1); const nm=[...r.meses]; nm[mi]=newVal*share; const q1=nm[0]+nm[1]+nm[2],q2=nm[3]+nm[4]+nm[5],q3=nm[6]+nm[7]+nm[8],q4=nm[9]+nm[10]+nm[11]; return {...r,meses:nm,q1,q2,q3,q4,total:q1+q2+q3+q4}; });
      window.storage?.set("dr2-rows",JSON.stringify(updated)).catch(()=>{});
      return updated;
    });
  },[activeSess]);

  const handleImport=useCallback((file)=>{ if(!file) return; const reader=new FileReader(); reader.onload=e=>{ try{ const wb=XLSX.read(e.target.result,{type:"binary"}); if(wb.SheetNames.includes("DB")){ const data=XLSX.utils.sheet_to_json(wb.Sheets["DB"]); const newRows=data.map(row=>{ const meses=MF.map(m=>+(row[m]||0)); const q1=meses[0]+meses[1]+meses[2],q2=meses[3]+meses[4]+meses[5],q3=meses[6]+meses[7]+meses[8],q4=meses[9]+meses[10]+meses[11]; return {canal:row["Canal"]||"",key:row["key"]||"",comentario:row["Comentario"]||"",modificacion:row["Modificación"]||"",item:String(row["Item"]||""),cadena:row["Cadena"]||"",codCliente:row["COD-CLIENTE"]||"",categoria:row["Categoría"]||"",subCategoria:row["Sub Categoría"]||"",marca:row["Marca"]||"",segmento:row["Segmento"]||"",uom:row["Unidad de Medida"]||"FDO",figurasClave:row["Figura Clave"]||"",skuDesc:row["SKU&Descripción"]||"",material:row["Material"]||"",clasificacion:row["Clasificación"]||"",meses,q1,q2,q3,q4,total:q1+q2+q3+q4}; }).filter(r=>r.item&&r.figurasClave&&FIGS.includes(r.figurasClave)); setRows(newRows); window.storage?.set("dr2-rows",JSON.stringify(newRows)).catch(()=>{}); } }catch(err){ console.error(err); } }; reader.readAsBinaryString(file); },[]);

  const handleExport=useCallback(()=>{ const data=rows.map(r=>{ const row={"Canal":r.canal,"key":r.key,"Comentario":r.comentario,"Modificación":r.modificacion,"Item":r.item,"Cadena":r.cadena,"COD-CLIENTE":r.codCliente,"Categoría":r.categoria,"Sub Categoría":r.subCategoria,"Marca":r.marca,"Segmento":r.segmento,"Unidad de Medida":r.uom,"Figura Clave":r.figurasClave,"SKU&Descripción":r.skuDesc,"Material":r.material,"Clasificación":r.clasificacion}; MF.forEach((m,i)=>{row[m]=r.meses[i]||0;}); row["Q1"]=r.q1;row["Q2"]=r.q2;row["Q3"]=r.q3;row["Q4"]=r.q4;row["Total"]=r.total; return row; }); const ws=XLSX.utils.json_to_sheet(data),wb2=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb2,ws,"DB"); XLSX.writeFile(wb2,`DemandReview_Export_${new Date().toISOString().slice(0,10)}.xlsx`); },[rows]);

  const canales=useMemo(()=>["ALL",...new Set(rows.map(r=>r.canal))],[rows]);

  return (
    <div style={S.app}>
      <Sidebar view={view} setView={setView} onImport={handleImport} onExport={handleExport} cfg={cfg} C={C} S={S}/>
      <div style={S.main}>
        {view==="dashboard" &&<DashboardView    rows={rows} uom={uom} setUom={setUom} canal={canal} setCanal={setCanal} canales={canales} C={C} S={S}/>}
        {view==="revision"  &&<RevisionView     rows={rows} uom={uom} canal={canal} onSKUClick={setSelSKU} C={C} S={S}/>}
        {view==="impacto"   &&<ImpactoView      rows={rows} log={log} sessions={sessions} uom={uom} C={C} S={S}/>}
        {view==="clientes"  &&<ClientesView     rows={rows} uom={uom} canal={canal} C={C} S={S}/>}
        {view==="abcxyz"    &&<AbcXyzView       rows={rows} uom={uom} C={C} S={S}/>}
        {view==="gaps"      &&<GapsView         rows={rows} uom={uom} C={C} S={S}/>}
        {view==="materia"   &&<MateriaPrimaView rows={rows} C={C} S={S}/>}
        {view==="reuniones" &&<ReunionesView    C={C} S={S}/>}
        {view==="modelos"   &&<ModelsView       models={models} setModels={setModels} editingModel={editMod} setEditingModel={setEditMod} C={C} S={S}/>}
        {view==="historial" &&<HistorialView    log={log} setLog={setLog} sessions={sessions} setSessions={setSessions} activeSession={activeSess} setActiveSession={setActiveSess} C={C} S={S}/>}
        {view==="tutorial"  &&<TutorialView     C={C} S={S}/>}
        {view==="config"    &&<ConfigView       cfg={cfg} setCfg={setCfg} C={C} S={S}/>}
      </div>
      {selSKU&&<SKUModal key={`${selSKU}-${uom}`} skuKey={selSKU} rows={rows} uom={uom} onChange={handleChange} onClose={()=>setSelSKU(null)} editableFrom={editableFrom} C={C} S={S}/>}
    </div>
  );
}
