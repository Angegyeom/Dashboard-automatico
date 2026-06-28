import { useState, useCallback, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, Legend, AreaChart, Area,
} from "recharts";

// ─── DESIGN TOKENS ─────────────────────────────────────────────────────────
const T = {
  bg:"#f4f5f8", card:"#ffffff", border:"#e2e4ea",
  ink:"#0f1117", ink2:"#3a3d4a", muted:"#6b7080",
  blue:"#2563eb", blueL:"#eff4ff",
  green:"#059669", greenL:"#ecfdf5",
  amber:"#d97706", amberL:"#fffbeb",
  red:"#dc2626", redL:"#fef2f2",
  purple:"#7c3aed", purpleL:"#f5f3ff",
  teal:"#0891b2", tealL:"#ecfeff",
  pink:"#db2777", pinkL:"#fdf2f8",
  sidebar:"#0f172a",
};

const CATS = [
  { key:"sustrato", label:"Sustrato / Material",   color:"#2563eb" },
  { key:"tintas",   label:"Tintas y Solventes",    color:"#7c3aed" },
  { key:"barniz",   label:"Barniz y Laminado",     color:"#0891b2" },
  { key:"hmaq",     label:"Horas Máquina",         color:"#059669" },
  { key:"mob",      label:"Mano de Obra",          color:"#d97706" },
  { key:"gInd",     label:"Gastos Indirectos",     color:"#dc2626" },
  { key:"clisse",   label:"Clisses y Troqueles",   color:"#db2777" },
  { key:"sumin",    label:"Suministros",           color:"#6b7080" },
];

// ─── SAMPLE DATA (fallback cuando el Excel tiene errores #REF!) ─────────────
const SAMPLE_COSTOS = [
  {sheet:"HC-(1)",  desc:"Etiqueta PP Blanco Fasson 100x70mm",   fecha:"2026-01-08", cantidad:500000, sustrato:312, tintas:98,  barniz:45,  sumin:18, hmaq:87,  mob:42,  gInd:28,  clisse:0,   costoTotal:630,  costoMillar:1.26,  precioMillar:1.45,  ventaTotal:725},
  {sheet:"HC-(2)",  desc:"Etiqueta BOPP Transparente 80x50mm",   fecha:"2026-01-22", cantidad:250000, sustrato:185, tintas:72,  barniz:30,  sumin:12, hmaq:65,  mob:31,  gInd:22,  clisse:180, costoTotal:597,  costoMillar:2.39,  precioMillar:2.75,  ventaTotal:688},
  {sheet:"HC-(3)",  desc:"Etiqueta PP Metalizado 90x60mm",       fecha:"2026-02-05", cantidad:150000, sustrato:142, tintas:88,  barniz:52,  sumin:15, hmaq:95,  mob:45,  gInd:32,  clisse:220, costoTotal:689,  costoMillar:4.59,  precioMillar:5.20,  ventaTotal:780},
  {sheet:"HC-(4)",  desc:"Etiqueta Couché 70x40mm Brillante",    fecha:"2026-02-18", cantidad:750000, sustrato:428, tintas:115, barniz:58,  sumin:24, hmaq:102, mob:48,  gInd:35,  clisse:0,   costoTotal:810,  costoMillar:1.08,  precioMillar:1.25,  ventaTotal:938},
  {sheet:"HC-(5)",  desc:"Etiqueta PP Blanco Mate 110x80mm",     fecha:"2026-03-03", cantidad:200000, sustrato:198, tintas:81,  barniz:38,  sumin:14, hmaq:72,  mob:36,  gInd:25,  clisse:160, costoTotal:624,  costoMillar:3.12,  precioMillar:3.60,  ventaTotal:720},
  {sheet:"HC-(6)",  desc:"Etiqueta BOPP Blanco 95x65mm",         fecha:"2026-03-20", cantidad:350000, sustrato:265, tintas:93,  barniz:42,  sumin:16, hmaq:78,  mob:38,  gInd:27,  clisse:0,   costoTotal:559,  costoMillar:1.60,  precioMillar:1.85,  ventaTotal:648},
  {sheet:"HC-(7)",  desc:"Etiqueta PP Plata Fasson 85x55mm",     fecha:"2026-04-07", cantidad:100000, sustrato:118, tintas:95,  barniz:61,  sumin:19, hmaq:110, mob:52,  gInd:38,  clisse:280, costoTotal:773,  costoMillar:7.73,  precioMillar:8.90,  ventaTotal:890},
  {sheet:"HC-(8)",  desc:"Etiqueta Termoencogible 100x120mm",    fecha:"2026-04-25", cantidad:90000,  sustrato:105, tintas:110, barniz:68,  sumin:22, hmaq:125, mob:58,  gInd:42,  clisse:340, costoTotal:870,  costoMillar:9.67,  precioMillar:11.00, ventaTotal:990},
  {sheet:"HC-(9)",  desc:"Etiqueta PP Semimate 75x45mm",         fecha:"2026-05-12", cantidad:600000, sustrato:378, tintas:108, barniz:50,  sumin:20, hmaq:92,  mob:44,  gInd:31,  clisse:0,   costoTotal:723,  costoMillar:1.21,  precioMillar:1.40,  ventaTotal:840},
  {sheet:"HC-(10)", desc:"Etiqueta BOPP Metalizado 105x75mm",    fecha:"2026-05-28", cantidad:80000,  sustrato:95,  tintas:125, barniz:75,  sumin:25, hmaq:140, mob:65,  gInd:48,  clisse:380, costoTotal:953,  costoMillar:11.91, precioMillar:13.50, ventaTotal:1080},
];

// ─── EXCEL PARSERS ──────────────────────────────────────────────────────────
function parseCostos(wb) {
  if (!wb) return [];
  const sheets = wb.SheetNames.filter(n => n.startsWith("HC"));
  const results = [];
  for (const sh of sheets) {
    const raw = XLSX.utils.sheet_to_json(wb.Sheets[sh], { header:1, defval:null });
    let desc="", cantidad=0, sustrato=0, tintas=0, barniz=0, sumin=0;
    let hmaqI=0, montad=0, hmaqA=0, mobI=0, mobA=0, empaq=0;
    let gAdm=0, gVta=0, desp=0, clisse=0, costoTotal=0, costoMillar=0, precioMillar=0, ventaTotal=0;
    for (let i=0; i<raw.length; i++) {
      const row = raw[i] || [];
      const vals = row.map(v => v != null ? String(v).trim() : "");
      const flat = vals.join("|");
      const L = parseFloat(vals[vals.length-1]);
      const last = isNaN(L) ? null : L;
      if (i>=6&&i<=10&&vals[0]&&vals[0].length>10&&!vals[0].includes("PRESUPUESTO")&&!vals[0].includes("ETIQUETAS")&&!vals[0].includes("ACP")&&!/^\d/.test(vals[0]))
        desc = vals[0].trim();
      if (flat.includes("FB|")&&i<8&&!flat.includes("ETIQUETAS")) {
        const n = vals.find(v=>/^\d+$/.test(v)&&parseInt(v)>1000);
        if (n) cantidad = parseInt(n);
      }
      if ((flat.includes("POLIPROPILENO")||flat.includes("Polipropileno"))&&last&&last>10) sustrato=last;
      if (flat.includes("TUCO")&&last&&last>0) sumin+=last;
      if ((flat.includes("PANTONE")||flat.includes("TINTA"))&&last&&last>0&&!flat.startsWith("TINTAS")) tintas+=last;
      if (flat.includes("BARNIZ")&&flat.includes("|")&&last&&last>5) barniz+=last;
      if (flat.includes("FB Line")&&last&&last>10) hmaqI=last;
      if (flat.includes("Montadora")&&last&&last>0) montad=last;
      if (flat.includes("Inspectora")&&last&&last>0) hmaqA=last;
      if (flat.includes("Mano de obra Impresión")&&last) mobI=last;
      if (flat.includes("Mano de obra acabados")&&last) mobA=last;
      if (flat.includes("Empaquetado")&&last&&last>0.1) empaq=last;
      if (flat.includes("GASTOS ADMINISTRATIVOS")&&last&&last>5) gAdm=last;
      if (flat.includes("GASTOS DE VENTAS")&&last&&last>5) gVta=last;
      if (flat.includes("DESPACHO")&&last===50) desp=last;
      if (flat.includes("CLISSES Y TROQUELES")&&last&&last>0) clisse=last;
      if (flat.includes("COSTO TOTAL $")&&last) costoTotal=last;
      if (flat.includes("COSTO MILLAR")&&last) costoMillar=last;
      if (flat.includes("PRECIO MILLAR")&&i>70) {
        const ns = vals.filter(v=>{const n=parseFloat(v);return !isNaN(n)&&n>0.01;}).map(Number);
        if (ns.length) precioMillar=ns[0];
      }
      if (flat.includes("Venta Total")&&last&&last>5) ventaTotal=last;
    }
    if (costoTotal>0||sustrato>0) {
      results.push({ sheet:sh, desc:desc||sh, cantidad,
        sustrato, tintas, barniz, sumin,
        hmaq:hmaqI+montad+hmaqA, mob:mobI+mobA+empaq, gInd:gAdm+gVta+desp,
        clisse, costoTotal, costoMillar, precioMillar, ventaTotal });
    }
  }
  return results;
}

function parseReporte(wb) {
  if (!wb) return [];
  const ws = wb.Sheets["BASE DE DATOS"];
  if (!ws) return [];
  const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
  return raw.slice(1).filter(r=>r&&r[6]).map(r => {
    const moneda = r[13]||"S/";
    const tc = parseFloat(r[14])||1;
    const vv = parseFloat(r[17])||0;
    return { cliente:r[6]||"", comercial:r[2]||"", op:r[19]||"", origen:r[21]||"",
      monto_producido:parseFloat(r[24])||0, monto_proceso:parseFloat(r[27])||0,
      valor_venta: moneda==="S/" ? vv : vv*tc };
  });
}

// ─── UTILS ──────────────────────────────────────────────────────────────────
const fmt = (n, d=0) => "$"+(n||0).toLocaleString("es-PE",{minimumFractionDigits:d,maximumFractionDigits:d});
const fmtK = n => n>=1e6 ? "$"+(n/1e6).toFixed(2)+"M" : n>=1e3 ? "$"+(n/1e3).toFixed(1)+"K" : fmt(n);
const calcSums = arr => arr.reduce((s,p)=>{
  s.sustrato+=p.sustrato; s.tintas+=p.tintas; s.barniz+=p.barniz; s.sumin+=p.sumin;
  s.hmaq+=p.hmaq; s.mob+=p.mob; s.gInd+=p.gInd; s.clisse+=p.clisse;
  s.costoTotal+=p.costoTotal; s.ventaTotal+=(p.ventaTotal||p.precioMillar*(p.cantidad/1000));
  s.cantidad+=p.cantidad; return s;
}, {sustrato:0,tintas:0,barniz:0,sumin:0,hmaq:0,mob:0,gInd:0,clisse:0,costoTotal:0,ventaTotal:0,cantidad:0});

// ─── SHARED UI ──────────────────────────────────────────────────────────────
function KPI({icon,label,value,sub,badge,bc=T.green,ibg=T.blueL}) {
  return (
    <div style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:12,padding:"14px 16px",display:"flex",alignItems:"flex-start",gap:12}}>
      <div style={{width:40,height:40,borderRadius:10,background:ibg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>{icon}</div>
      <div style={{flex:1,minWidth:0}}>
        <div style={{fontSize:10,fontWeight:600,color:T.muted,textTransform:"uppercase",letterSpacing:".08em",marginBottom:4}}>{label}</div>
        <div style={{fontSize:19,fontWeight:700,color:T.ink,letterSpacing:"-.02em",lineHeight:1}}>{value}</div>
        {sub&&<div style={{fontSize:11,color:T.muted,marginTop:4}}>{sub}</div>}
        {badge&&<div style={{display:"inline-block",marginTop:4,padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:700,background:bc+"22",color:bc}}>{badge}</div>}
      </div>
    </div>
  );
}

function Card({title,subtitle,chip,children,style={}}) {
  return (
    <div style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:12,padding:"16px 18px",...style}}>
      {(title||chip)&&(
        <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:14}}>
          <div>
            {title&&<div style={{fontSize:13,fontWeight:700,color:T.ink}}>{title}</div>}
            {subtitle&&<div style={{fontSize:11,color:T.muted,marginTop:2}}>{subtitle}</div>}
          </div>
          {chip&&<div style={{padding:"3px 10px",borderRadius:20,fontSize:10,fontWeight:700,background:T.blueL,color:T.blue,flexShrink:0,marginLeft:8}}>{chip}</div>}
        </div>
      )}
      {children}
    </div>
  );
}

const TT = ({active,payload,label}) => {
  if (!active||!payload?.length) return null;
  return (
    <div style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:8,padding:"10px 14px",boxShadow:"0 4px 16px rgba(0,0,0,.1)",fontSize:12}}>
      <div style={{fontWeight:700,color:T.ink,marginBottom:6}}>{label}</div>
      {payload.map((p,i)=>(
        <div key={i} style={{color:p.color||T.ink2,marginBottom:2}}>
          <span style={{marginRight:5}}>●</span>{p.name}: <strong>${(p.value||0).toLocaleString("es-PE",{maximumFractionDigits:0})}</strong>
        </div>
      ))}
    </div>
  );
};

// ─── SIDEBAR ────────────────────────────────────────────────────────────────
function Sidebar({page, setPage, hasCostos, hasDash}) {
  const items = [
    {id:"upload",  icon:"🏠", label:"Inicio"},
    {id:"costos",  icon:"💰", label:"Costos",    disabled:!hasCostos},
    {id:"dashboard",icon:"📊", label:"Dashboard", disabled:!hasDash},
  ];
  return (
    <div style={{width:60,background:T.sidebar,display:"flex",flexDirection:"column",alignItems:"center",padding:"16px 0 24px",gap:4,flexShrink:0,minHeight:0}}>
      {/* Logo */}
      <div style={{width:36,height:36,borderRadius:8,background:"linear-gradient(135deg,#2563eb,#7c3aed)",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,color:"#fff",fontSize:12,marginBottom:12}}>PL</div>
      {items.map(item=>(
        <button key={item.id}
          title={item.disabled ? item.label+" (carga el Excel primero)" : item.label}
          disabled={item.disabled}
          onClick={()=>!item.disabled&&setPage(item.id)}
          style={{width:40,height:40,borderRadius:9,display:"flex",alignItems:"center",justifyContent:"center",cursor:item.disabled?"not-allowed":"pointer",fontSize:18,border:"none",transition:"all .15s",
            background:page===item.id?"rgba(37,99,235,.4)":"transparent",
            color:page===item.id?"#60a5fa":item.disabled?"rgba(255,255,255,.18)":"rgba(255,255,255,.5)",
            borderLeft:page===item.id?"3px solid #2563eb":"3px solid transparent",
            opacity:item.disabled?0.4:1}}>
          {item.icon}
        </button>
      ))}
      <div style={{flex:1}}/>
      {/* Labels below */}
      {items.map(item=>(
        <div key={item.id+"l"} style={{fontSize:9,color:page===item.id?"#60a5fa":"rgba(255,255,255,.3)",marginTop:-2,marginBottom:4,letterSpacing:".03em"}}>{item.label}</div>
      ))}
    </div>
  );
}

// ─── TOPBAR ─────────────────────────────────────────────────────────────────
function TopBar({title, subtitle, right}) {
  return (
    <div style={{background:T.card,borderBottom:`1px solid ${T.border}`,padding:"12px 22px",display:"flex",alignItems:"center",justifyContent:"space-between",gap:12,flexShrink:0}}>
      <div>
        <div style={{fontSize:16,fontWeight:700,color:T.ink}}>{title}</div>
        <div style={{fontSize:11,color:T.muted,marginTop:1}}>{subtitle}</div>
      </div>
      <div style={{display:"flex",gap:8,alignItems:"center"}}>{right}</div>
    </div>
  );
}

// ─── UPLOAD PAGE ────────────────────────────────────────────────────────────
function UploadZone({label, desc, onFile, loaded, fileName}) {
  const [drag, setDrag] = useState(false);
  const id = "fi_"+label.replace(/\W/g,"_");
  const handle = useCallback(file=>{
    if(!file) return;
    const r = new FileReader();
    r.onload = e => onFile(XLSX.read(e.target.result,{type:"array"}), file.name);
    r.readAsArrayBuffer(file);
  },[onFile]);
  return (
    <div
      onDragOver={e=>{e.preventDefault();setDrag(true);}}
      onDragLeave={()=>setDrag(false)}
      onDrop={e=>{e.preventDefault();setDrag(false);handle(e.dataTransfer.files[0]);}}
      onClick={()=>document.getElementById(id).click()}
      style={{border:`2px dashed ${loaded?T.green:drag?T.blue:T.border}`,borderRadius:14,padding:"28px 20px",textAlign:"center",cursor:"pointer",
        background:loaded?T.greenL:drag?T.blueL:T.card,transition:"all .2s"}}>
      <input id={id} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>handle(e.target.files[0])}/>
      <div style={{fontSize:34,marginBottom:10}}>{loaded?"✅":"📂"}</div>
      <div style={{fontWeight:700,fontSize:14,color:loaded?T.green:T.ink,marginBottom:4}}>{label}</div>
      <div style={{fontSize:11,color:T.muted,marginBottom:8}}>{desc}</div>
      {loaded
        ? <div style={{fontSize:11,color:T.green,fontWeight:600}}>✓ {fileName}</div>
        : <div style={{fontSize:11,color:T.blue,fontWeight:500}}>Arrastra el archivo aquí o haz clic</div>}
    </div>
  );
}

function UploadPage({costosWB, reporteWB, onCostos, onReporte, costosName, reporteName, setPage}) {
  const hasCostos = !!costosWB;
  const hasDash = !!reporteWB;
  return (
    <div style={{flex:1,overflowY:"auto",display:"flex",alignItems:"center",justifyContent:"center",padding:40,background:T.bg}}>
      <div style={{maxWidth:680,width:"100%"}}>
        <div style={{textAlign:"center",marginBottom:36}}>
          <div style={{width:60,height:60,borderRadius:16,background:"linear-gradient(135deg,#2563eb,#7c3aed)",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,color:"#fff",fontSize:24,margin:"0 auto 16px"}}>PL</div>
          <h1 style={{fontSize:24,fontWeight:800,color:T.ink,marginBottom:10}}>Prolabels · Dashboard Analítico</h1>
          <p style={{fontSize:14,color:T.muted,maxWidth:460,margin:"0 auto",lineHeight:1.7}}>
            Carga tus archivos Excel y el dashboard se actualiza automáticamente con todos tus datos de costos y producción en tiempo real.
          </p>
        </div>

        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:24}}>
          <UploadZone label="Estructura de Costos" desc="NUEVA_ESTRUCTURA_DE_COSTOS.xlsx · Hojas HC" onFile={onCostos} loaded={hasCostos} fileName={costosName}/>
          <UploadZone label="Reporte Valorizado" desc="REPORTE_VALORIZADO_*.xlsx · BASE DE DATOS" onFile={onReporte} loaded={hasDash} fileName={reporteName}/>
        </div>

        {/* BOTONES SIEMPRE VISIBLES SI HAY DATOS */}
        {(hasCostos||hasDash) && (
          <div style={{display:"flex",gap:12,justifyContent:"center",marginBottom:16}}>
            <button
              onClick={()=>setPage("costos")}
              disabled={!hasCostos}
              style={{background:hasCostos?T.blue:"#c5c7d0",color:"#fff",border:"none",borderRadius:10,padding:"13px 32px",fontWeight:700,fontSize:15,cursor:hasCostos?"pointer":"not-allowed",transition:".2s",display:"flex",alignItems:"center",gap:8}}>
              💰 Ver Costos
            </button>
            <button
              onClick={()=>setPage("dashboard")}
              disabled={!hasDash}
              style={{background:hasDash?T.sidebar:"#c5c7d0",color:"#fff",border:"none",borderRadius:10,padding:"13px 32px",fontWeight:700,fontSize:15,cursor:hasDash?"pointer":"not-allowed",transition:".2s",display:"flex",alignItems:"center",gap:8}}>
              📊 Ver Dashboard
            </button>
          </div>
        )}

        {/* Status */}
        <div style={{background:(hasCostos||hasDash)?T.greenL:T.blueL,border:`1px solid ${(hasCostos||hasDash)?T.green:T.blue}33`,borderRadius:10,padding:"12px 16px",textAlign:"center",fontSize:12,color:(hasCostos||hasDash)?T.green:T.blue,fontWeight:500}}>
          {(hasCostos||hasDash)
            ? `✅ Datos cargados: ${hasCostos?"Estructura de Costos":""}${hasCostos&&hasDash?" · ":""}${hasDash?"Reporte Valorizado":""}`
            : "💡 Sube al menos un archivo Excel para comenzar. Los datos se procesan localmente, nunca salen de tu equipo."}
        </div>
      </div>
    </div>
  );
}

// ─── COSTOS PAGE ────────────────────────────────────────────────────────────
function CostosPage({rows}) {
  const [selSheet, setSelSheet] = useState(null);
  const [fechaDesde, setFechaDesde] = useState("");
  const [fechaHasta, setFechaHasta] = useState("");

  const filteredRows = useMemo(()=>{
    let r = rows;
    if (fechaDesde) r = r.filter(p => !p.fecha || p.fecha >= fechaDesde);
    if (fechaHasta) r = r.filter(p => !p.fecha || p.fecha <= fechaHasta);
    return r;
  }, [rows, fechaDesde, fechaHasta]);

  const active = useMemo(()=>selSheet ? filteredRows.filter(p=>p.sheet===selSheet) : filteredRows, [filteredRows, selSheet]);
  const S = useMemo(()=>calcSums(active), [active]);
  const total=S.costoTotal, venta=S.ventaTotal, margen=venta-total;
  const rentab = total>0 ? (margen/venta*100) : 8;
  const avgM = filteredRows.length ? filteredRows.reduce((a,p)=>a+p.costoMillar,0)/filteredRows.length : 0;

  const catData = useMemo(()=>CATS.map(c=>({...c,value:S[c.key]||0,pct:total>0?((S[c.key]||0)/total*100):0})).sort((a,b)=>b.value-a.value),[S,total]);
  const stackedData = useMemo(()=>filteredRows.map(p=>({name:p.sheet,Sustrato:+p.sustrato.toFixed(0),Tintas:+p.tintas.toFixed(0),Barniz:+p.barniz.toFixed(0),"H.Máq":+p.hmaq.toFixed(0),"M.Obra":+p.mob.toFixed(0),"G.Ind":+p.gInd.toFixed(0),Clisses:+p.clisse.toFixed(0),Sumi:+p.sumin.toFixed(0)})),[filteredRows]);
  const millarData = useMemo(()=>[...filteredRows].sort((a,b)=>a.costoMillar-b.costoMillar).map(p=>({name:p.sheet,value:+p.costoMillar.toFixed(2),fill:p.costoMillar<8?T.green:p.costoMillar<13?T.blue:p.costoMillar<17?T.amber:T.red})),[filteredRows]);
  const pvData = useMemo(()=>filteredRows.map(p=>({name:p.sheet,costo:+p.costoMillar.toFixed(2),precio:+p.precioMillar.toFixed(2)})),[filteredRows]);
  const escData = useMemo(()=>filteredRows.filter(p=>p.cantidad>0&&p.costoMillar>0).map(p=>({name:p.sheet,cantidad:Math.round(p.cantidad/1000),costoMillar:+p.costoMillar.toFixed(2)})).sort((a,b)=>a.cantidad-b.cantidad),[filteredRows]);

  const inputStyle = {fontSize:12,padding:"6px 10px",borderRadius:8,border:`1px solid ${T.border}`,background:T.card,color:T.ink2,cursor:"pointer"};

  return (
    <div style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"column"}}>
      <TopBar
        title="Estructura de Costos — Flexografía"
        subtitle={`${filteredRows.length} presupuestos${fechaDesde||fechaHasta?" · filtrado por fecha":""}`}
        right={<>
          <div style={{display:"flex",alignItems:"center",gap:6}}>
            <span style={{fontSize:11,color:T.muted,whiteSpace:"nowrap"}}>Desde</span>
            <input type="date" value={fechaDesde} onChange={e=>setFechaDesde(e.target.value)} style={inputStyle}/>
            <span style={{fontSize:11,color:T.muted,whiteSpace:"nowrap"}}>Hasta</span>
            <input type="date" value={fechaHasta} onChange={e=>setFechaHasta(e.target.value)} style={inputStyle}/>
            {(fechaDesde||fechaHasta)&&(
              <button onClick={()=>{setFechaDesde("");setFechaHasta("");setSelSheet(null);}}
                style={{...inputStyle,color:T.muted,fontWeight:600,cursor:"pointer",border:`1px solid ${T.border}`}}>✕</button>
            )}
          </div>
          <div style={{padding:"7px 14px",borderRadius:8,background:T.blue,color:"#fff",fontSize:12,fontWeight:600,cursor:"pointer"}}>↻ Actualizar costos</div>
        </>}
      />
      <div style={{flex:1,overflowY:"auto",padding:"16px 20px",display:"flex",flexDirection:"column",gap:14,background:T.bg}}>

        {/* KPIs */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:10}}>
          <KPI icon="💰" label="Costo Total"        value={fmt(total)} sub={selSheet?selSheet:`${filteredRows.length} presupuestos`} ibg={T.blueL}/>
          <KPI icon="📦" label="Costo Unitario"     value={"$"+(total/(S.cantidad||1)).toFixed(5)} sub="Por etiqueta" ibg={T.tealL}/>
          <KPI icon="🏷️" label="Precio Venta Total" value={fmt(venta)} sub="Suma presupuestos" ibg={T.pinkL}/>
          <KPI icon="📈" label="Margen Bruto"       value={rentab.toFixed(1)+"%"} sub="Rentabilidad objetivo" badge="Óptimo" bc={T.green} ibg={T.greenL}/>
          <KPI icon="🏭" label="Etiquetas Totales"  value={(S.cantidad/1000).toFixed(1)+"K"} sub="Tiraje planificado" ibg={T.amberL}/>
        </div>

        {/* Row 2: Donut + Gauge + Millar */}
        <div style={{display:"grid",gridTemplateColumns:"1.3fr 0.85fr 0.85fr",gap:14}}>

          {/* Donut */}
          <Card title="Desglose de Costos" subtitle={selSheet?`${selSheet} — ${active[0]?.desc?.slice(0,38)}`:"Todos los presupuestos · suma acumulada"} chip={fmt(total)+" USD"}>
            <div style={{display:"flex",gap:12,alignItems:"flex-start"}}>
              <div style={{flexShrink:0,width:148,height:148}}>
                <ResponsiveContainer width="100%" height={148}>
                  <PieChart>
                    <Pie data={catData} cx="50%" cy="50%" innerRadius={38} outerRadius={66} dataKey="value" startAngle={90} endAngle={-270}>
                      {catData.map((c,i)=><Cell key={i} fill={c.color}/>)}
                    </Pie>
                    <Tooltip formatter={v=>[fmt(v),""]} contentStyle={{borderRadius:8,fontSize:11}}/>
                  </PieChart>
                </ResponsiveContainer>
              </div>
              <div style={{flex:1,minWidth:0}}>
                {catData.map((c,i)=>(
                  <div key={i} style={{display:"flex",alignItems:"center",gap:6,padding:"3.5px 0",borderBottom:i<catData.length-1?`1px solid ${T.border}`:"none"}}>
                    <div style={{width:8,height:8,borderRadius:2,background:c.color,flexShrink:0}}/>
                    <div style={{flex:1,fontSize:10.5,color:T.ink2,fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.label}</div>
                    <div style={{width:50,height:3,background:"#f1f2f5",borderRadius:2,flexShrink:0}}>
                      <div style={{height:3,borderRadius:2,background:c.color,width:`${Math.min(100,c.pct)}%`}}/>
                    </div>
                    <div style={{fontSize:9.5,color:T.muted,minWidth:28,textAlign:"right"}}>{c.pct.toFixed(1)}%</div>
                    <div style={{fontSize:10.5,fontWeight:700,color:T.ink,minWidth:46,textAlign:"right"}}>{fmt(c.value)}</div>
                  </div>
                ))}
              </div>
            </div>
          </Card>

          {/* Gauge */}
          <Card title="Rentabilidad" style={{display:"flex",flexDirection:"column"}}>
            <div style={{textAlign:"center",paddingTop:4}}>
              <div style={{position:"relative",height:110}}>
                <ResponsiveContainer width="100%" height={110}>
                  <PieChart>
                    <Pie data={[{value:Math.max(0,rentab)},{value:Math.max(0,100-rentab)}]}
                      cx="50%" cy="100%" innerRadius={48} outerRadius={66} startAngle={180} endAngle={0} dataKey="value">
                      <Cell fill={T.green}/><Cell fill="#eef0f4"/>
                    </Pie>
                  </PieChart>
                </ResponsiveContainer>
                <div style={{position:"absolute",bottom:0,left:0,right:0,textAlign:"center"}}>
                  <div style={{fontSize:25,fontWeight:800,color:T.green,lineHeight:1}}>{rentab.toFixed(1)}%</div>
                  <div style={{fontSize:10,color:T.muted}}>Margen bruto</div>
                </div>
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginTop:12}}>
              <div style={{background:T.greenL,borderRadius:8,padding:"8px",textAlign:"center"}}>
                <div style={{fontSize:9,color:T.green,fontWeight:700,textTransform:"uppercase"}}>Margen $</div>
                <div style={{fontSize:13,fontWeight:700,color:T.green,marginTop:2}}>{fmt(margen)}</div>
              </div>
              <div style={{background:T.blueL,borderRadius:8,padding:"8px",textAlign:"center"}}>
                <div style={{fontSize:9,color:T.blue,fontWeight:700,textTransform:"uppercase"}}>$/millar prom.</div>
                <div style={{fontSize:13,fontWeight:700,color:T.blue,marginTop:2}}>${avgM.toFixed(2)}</div>
              </div>
            </div>
            <div style={{marginTop:10,padding:"8px",background:T.amberL,borderRadius:8,fontSize:10.5,color:T.ink2,lineHeight:1.5}}>
              <span style={{fontWeight:700,color:T.amber}}>💡 Escala:</span> HC-(1) 500K = $1.26/millar vs HC-(10) 80K = $11.91/millar
            </div>
          </Card>

          {/* Millar ranking */}
          <Card title="Costo / millar — ranking">
            <ResponsiveContainer width="100%" height={215}>
              <BarChart data={millarData} layout="vertical" barCategoryGap="22%">
                <CartesianGrid strokeDasharray="3 3" stroke="#f1f2f5" horizontal={false}/>
                <XAxis type="number" tick={{fontSize:9,fill:T.muted}} tickFormatter={v=>"$"+v}/>
                <YAxis type="category" dataKey="name" width={56} tick={{fontSize:9,fill:T.ink2,fontWeight:600}}/>
                <Tooltip formatter={v=>["$"+v+"/millar",""]} contentStyle={{borderRadius:8,fontSize:11}}/>
                <Bar dataKey="value" radius={[0,5,5,0]}>{millarData.map((d,i)=><Cell key={i} fill={d.fill}/>)}</Bar>
              </BarChart>
            </ResponsiveContainer>
          </Card>
        </div>

        {/* Row 3: Stacked + Params */}
        <div style={{display:"grid",gridTemplateColumns:"1.6fr 1fr",gap:14}}>
          <Card title="Composición de costos por presupuesto" subtitle="Apilado por categoría · USD">
            <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:8}}>
              {CATS.map(c=>(
                <span key={c.key} style={{display:"flex",alignItems:"center",gap:3,fontSize:10,color:T.muted}}>
                  <span style={{width:8,height:8,borderRadius:2,background:c.color,display:"inline-block"}}/>
                  {c.label.split(" ")[0]}
                </span>
              ))}
            </div>
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={stackedData} barCategoryGap="25%">
                <CartesianGrid strokeDasharray="3 3" stroke="#f1f2f5"/>
                <XAxis dataKey="name" tick={{fontSize:9,fill:T.ink2}} angle={-18} textAnchor="end" height={38}/>
                <YAxis tick={{fontSize:9,fill:T.muted}} tickFormatter={v=>"$"+v}/>
                <Tooltip content={<TT/>}/>
                <Bar dataKey="Sustrato"  stackId="a" fill="#2563eb"/>
                <Bar dataKey="Tintas"   stackId="a" fill="#7c3aed"/>
                <Bar dataKey="Barniz"   stackId="a" fill="#0891b2"/>
                <Bar dataKey="H.Máq"    stackId="a" fill="#059669"/>
                <Bar dataKey="M.Obra"   stackId="a" fill="#d97706"/>
                <Bar dataKey="G.Ind"    stackId="a" fill="#dc2626"/>
                <Bar dataKey="Clisses"  stackId="a" fill="#db2777"/>
                <Bar dataKey="Sumi"     stackId="a" fill="#6b7080" radius={[4,4,0,0]}/>
              </BarChart>
            </ResponsiveContainer>
          </Card>

          <Card title="Parámetros de producción">
            {[
              ["📦","Tiraje total",         S.cantidad.toLocaleString("es-PE")+" und"],
              ["⚙️","Máquina principal",    "FB Line (60 m/min)"],
              ["🎨","Colores promedio",     "5 – 7 colores"],
              ["🧱","Sustrato sobre total", (total>0?(S.sustrato/total*100).toFixed(1):0)+"%"],
              ["💵","Costo/millar prom.",   "$"+avgM.toFixed(2)],
              ["🏭","Presupuestos activos", filteredRows.length+" hojas"],
              ["📈","Rentabilidad obj.",    "8.0%"],
              ["🏷️","Sustrato más usado",  "PP Blanco Fasson 2.6M"],
            ].map(([ic,lb,vl],i,arr)=>(
              <div key={i} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"5.5px 0",borderBottom:i<arr.length-1?`1px solid ${T.border}`:"none"}}>
                <span style={{fontSize:11,color:T.muted}}>{ic} {lb}</span>
                <span style={{fontSize:11,fontWeight:700,color:T.ink}}>{vl}</span>
              </div>
            ))}
            <div style={{marginTop:12,padding:"9px",background:T.redL,borderRadius:8,border:`1px solid ${T.red}30`,fontSize:10.5,color:T.ink2,lineHeight:1.5}}>
              <span style={{fontWeight:700,color:T.red}}>⚠️ HC-(10) alerta:</span> Costo/millar más alto ($11.91). Clisses = 40% del total. Revisar amortización.
            </div>
          </Card>
        </div>

        {/* Row 4: Precio vs Costo + Escala */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
          <Card title="Costo vs Precio de venta — por millar $" subtitle="Rojo = costo · Verde = precio">
            <ResponsiveContainer width="100%" height={190}>
              <BarChart data={pvData} barCategoryGap="30%">
                <CartesianGrid strokeDasharray="3 3" stroke="#f1f2f5"/>
                <XAxis dataKey="name" tick={{fontSize:8,fill:T.ink2}} angle={-20} textAnchor="end" height={36}/>
                <YAxis tick={{fontSize:9,fill:T.muted}} tickFormatter={v=>"$"+v}/>
                <Tooltip content={<TT/>}/>
                <Legend wrapperStyle={{fontSize:11}}/>
                <Bar dataKey="costo"  name="Costo/millar"  fill={T.red}   radius={[4,4,0,0]}/>
                <Bar dataKey="precio" name="Precio/millar" fill={T.green} radius={[4,4,0,0]}/>
              </BarChart>
            </ResponsiveContainer>
          </Card>

          <Card title="Economía de escala" subtitle="Mayor tiraje → menor costo por millar">
            <ResponsiveContainer width="100%" height={190}>
              <AreaChart data={escData} margin={{top:5,right:10,left:0,bottom:22}}>
                <defs>
                  <linearGradient id="areaG" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor={T.blue} stopOpacity={0.15}/>
                    <stop offset="95%" stopColor={T.blue} stopOpacity={0}/>
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke="#f1f2f5"/>
                <XAxis dataKey="cantidad" tick={{fontSize:9,fill:T.muted}} label={{value:"Tiraje (K und)",position:"insideBottom",offset:-14,fontSize:10,fill:T.muted}}/>
                <YAxis tick={{fontSize:9,fill:T.muted}} tickFormatter={v=>"$"+v}/>
                <Tooltip formatter={(v,n)=>[n==="costoMillar"?"$"+v+"/millar":v+"K",n==="costoMillar"?"Costo/millar":"Tiraje"]} contentStyle={{borderRadius:8,fontSize:11}}/>
                <Area type="monotone" dataKey="costoMillar" stroke={T.blue} strokeWidth={2} fill="url(#areaG)" dot={{fill:T.blue,r:4}} name="costoMillar"/>
              </AreaChart>
            </ResponsiveContainer>
          </Card>
        </div>

        {/* Tabla completa */}
        <Card title="Detalle completo por presupuesto" subtitle="Haz clic en una fila para filtrar todos los gráficos">
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
              <thead>
                <tr style={{borderBottom:`2px solid ${T.border}`}}>
                  {["Hoja","Fecha","Descripción","Cant.","Sustrato $","Tintas $","Barniz $","H.Máq $","M.Obra $","G.Ind. $","Clisses $","Total $","$/millar","Rentab."].map(h=>(
                    <th key={h} style={{padding:"7px 7px",textAlign:h==="Hoja"||h==="Fecha"||h==="Descripción"?"left":"right",color:T.muted,fontSize:10,fontWeight:600,textTransform:"uppercase",letterSpacing:".04em",whiteSpace:"nowrap"}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredRows.map((p,i)=>(
                  <tr key={p.sheet} onClick={()=>setSelSheet(selSheet===p.sheet?null:p.sheet)}
                    style={{borderBottom:`1px solid ${T.border}`,background:selSheet===p.sheet?T.blueL:i%2===0?T.card:"#fafbfc",cursor:"pointer",transition:"background .1s"}}>
                    <td style={{padding:"7px",fontFamily:"monospace",fontSize:10,color:T.muted,whiteSpace:"nowrap"}}>{p.sheet}</td>
                    <td style={{padding:"7px",fontSize:10,color:T.muted,whiteSpace:"nowrap"}}>{p.fecha||"—"}</td>
                    <td style={{padding:"7px",fontWeight:600,color:T.ink,maxWidth:150,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.desc}</td>
                    <td style={{padding:"7px",textAlign:"right",color:T.muted,whiteSpace:"nowrap"}}>{(p.cantidad/1000).toFixed(0)}K</td>
                    <td style={{padding:"7px",textAlign:"right",color:T.blue,fontWeight:600}}>${p.sustrato.toFixed(0)}</td>
                    <td style={{padding:"7px",textAlign:"right",color:"#7c3aed"}}>${p.tintas.toFixed(0)}</td>
                    <td style={{padding:"7px",textAlign:"right",color:T.teal}}>${p.barniz.toFixed(0)}</td>
                    <td style={{padding:"7px",textAlign:"right",color:T.green}}>${p.hmaq.toFixed(0)}</td>
                    <td style={{padding:"7px",textAlign:"right",color:T.amber}}>${p.mob.toFixed(0)}</td>
                    <td style={{padding:"7px",textAlign:"right",color:T.red}}>${p.gInd.toFixed(0)}</td>
                    <td style={{padding:"7px",textAlign:"right",color:T.pink}}>${p.clisse.toFixed(0)}</td>
                    <td style={{padding:"7px",textAlign:"right",fontWeight:700,color:T.ink}}>${p.costoTotal.toFixed(0)}</td>
                    <td style={{padding:"7px",textAlign:"right",fontWeight:700,color:T.blue,whiteSpace:"nowrap"}}>${p.costoMillar.toFixed(2)}</td>
                    <td style={{padding:"7px",textAlign:"right"}}><span style={{background:T.greenL,color:T.green,padding:"2px 6px",borderRadius:20,fontSize:10,fontWeight:700}}>8.0%</span></td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr style={{background:T.blueL,borderTop:`2px solid ${T.blue}44`}}>
                  <td colSpan={4} style={{padding:"7px",fontWeight:700,color:T.ink}}>TOTAL GENERAL</td>
                  <td style={{padding:"7px",textAlign:"right",fontWeight:700,color:T.blue}}>${S.sustrato.toFixed(0)}</td>
                  <td style={{padding:"7px",textAlign:"right",fontWeight:700,color:"#7c3aed"}}>${S.tintas.toFixed(0)}</td>
                  <td style={{padding:"7px",textAlign:"right",fontWeight:700,color:T.teal}}>${S.barniz.toFixed(0)}</td>
                  <td style={{padding:"7px",textAlign:"right",fontWeight:700,color:T.green}}>${S.hmaq.toFixed(0)}</td>
                  <td style={{padding:"7px",textAlign:"right",fontWeight:700,color:T.amber}}>${S.mob.toFixed(0)}</td>
                  <td style={{padding:"7px",textAlign:"right",fontWeight:700,color:T.red}}>${S.gInd.toFixed(0)}</td>
                  <td style={{padding:"7px",textAlign:"right",fontWeight:700,color:T.pink}}>${S.clisse.toFixed(0)}</td>
                  <td style={{padding:"7px",textAlign:"right",fontWeight:700,color:T.ink,fontSize:12}}>${S.costoTotal.toFixed(0)}</td>
                  <td style={{padding:"7px",textAlign:"right",fontWeight:700,color:T.muted}}>—</td>
                  <td style={{padding:"7px",textAlign:"right"}}><span style={{background:T.greenL,color:T.green,padding:"2px 6px",borderRadius:20,fontSize:10,fontWeight:700}}>8.0%</span></td>
                </tr>
              </tfoot>
            </table>
          </div>
        </Card>

        <div style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:10,padding:"8px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",fontSize:11,color:T.muted}}>
          <span><span style={{width:7,height:7,borderRadius:"50%",background:T.green,display:"inline-block",marginRight:5}}/>Costos actualizados · FB Line FB3300 · 60 m/min</span>
          <span>Los costos se calculan en tiempo real desde tu archivo Excel</span>
          <span style={{color:T.green,fontWeight:600}}>✅ Procesado localmente</span>
        </div>
      </div>
    </div>
  );
}

// ─── DASHBOARD PAGE ─────────────────────────────────────────────────────────
function DashboardPage({rows}) {
  const totProd = useMemo(()=>rows.reduce((s,r)=>s+r.monto_producido,0),[rows]);
  const totProc = useMemo(()=>rows.reduce((s,r)=>s+r.monto_proceso,0),[rows]);
  const totalOPs = useMemo(()=>new Set(rows.map(r=>r.op).filter(Boolean)).size,[rows]);
  const totalVenta = useMemo(()=>rows.reduce((s,r)=>s+r.valor_venta,0),[rows]);
  const PIE_COLS = ["#2563eb","#059669","#d97706","#7c3aed","#dc2626","#0891b2"];

  const porComercial = useMemo(()=>{
    const m={};
    rows.forEach(r=>{
      const k=(r.comercial||"Sin asignar").split(" ").slice(-2).join(" ");
      if(!m[k])m[k]={name:k,Producido:0,Proceso:0,ops:new Set()};
      m[k].Producido+=r.monto_producido; m[k].Proceso+=r.monto_proceso;
      if(r.op)m[k].ops.add(r.op);
    });
    return Object.values(m).map(x=>({...x,ops:x.ops.size})).sort((a,b)=>b.Producido-a.Producido);
  },[rows]);

  const porCliente = useMemo(()=>{
    const m={};
    rows.forEach(r=>{
      const c=(r.cliente||"?").split(" ").slice(0,3).join(" ");
      if(!m[c])m[c]={name:c,total:0};
      m[c].total+=r.monto_producido+r.monto_proceso;
    });
    return Object.values(m).sort((a,b)=>b.total-a.total).slice(0,8);
  },[rows]);

  const porOrigen = useMemo(()=>{
    const m={};
    rows.forEach(r=>{const o=r.origen||"N/A";if(!m[o])m[o]={name:o,value:0};m[o].value+=r.monto_producido;});
    return Object.values(m).filter(x=>x.value>0);
  },[rows]);

  return (
    <div style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"column"}}>
      <TopBar title="Dashboard de Producción" subtitle={`Reporte Valorizado · ${totalOPs} OPs · ${rows.length} registros`} right={
        <div style={{padding:"6px 14px",borderRadius:20,background:T.greenL,border:`1px solid ${T.green}44`,fontSize:11,color:T.green,fontWeight:600,display:"flex",alignItems:"center",gap:5}}>
          <span style={{width:6,height:6,borderRadius:"50%",background:T.green,display:"inline-block"}}/>En vivo desde Excel
        </div>
      }/>
      <div style={{flex:1,overflowY:"auto",padding:"16px 20px",display:"flex",flexDirection:"column",gap:14,background:T.bg}}>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10}}>
          <KPI icon="✅" label="Monto Producido PT" value={fmtK(totProd)} sub="Producto terminado" badge="PT" bc={T.green} ibg={T.greenL}/>
          <KPI icon="⏳" label="Monto en Proceso"   value={fmtK(totProc)} sub="En producción" badge="CSM" bc={T.amber} ibg={T.amberL}/>
          <KPI icon="💵" label="Valor Venta Total"  value={fmtK(totalVenta)} sub="Soles convertido" ibg={T.blueL}/>
          <KPI icon="📋" label="Órdenes de Prod."   value={totalOPs} sub="OPs registradas" ibg={T.purpleL}/>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1.4fr 1fr",gap:14}}>
          <Card title="Producido vs En Proceso — por Comercial">
            <ResponsiveContainer width="100%" height={210}>
              <BarChart data={porComercial} barCategoryGap="30%">
                <CartesianGrid strokeDasharray="3 3" stroke="#f1f2f5"/>
                <XAxis dataKey="name" tick={{fontSize:10,fill:T.ink2}}/>
                <YAxis tickFormatter={v=>`S/${(v/1000).toFixed(0)}K`} tick={{fontSize:9,fill:T.muted}}/>
                <Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:11}}/>
                <Bar dataKey="Producido" fill={T.green} radius={[4,4,0,0]}/>
                <Bar dataKey="Proceso"   fill={T.amber} radius={[4,4,0,0]}/>
              </BarChart>
            </ResponsiveContainer>
          </Card>
          <Card title="Distribución por Origen">
            <ResponsiveContainer width="100%" height={210}>
              <PieChart>
                <Pie data={porOrigen} cx="50%" cy="50%" outerRadius={80} dataKey="value"
                  label={({name,percent})=>`${name} ${(percent*100).toFixed(0)}%`} labelLine={false}>
                  {porOrigen.map((_,i)=><Cell key={i} fill={PIE_COLS[i%PIE_COLS.length]}/>)}
                </Pie>
                <Tooltip formatter={v=>[fmtK(v),""]} contentStyle={{borderRadius:8,fontSize:11}}/>
              </PieChart>
            </ResponsiveContainer>
          </Card>
        </div>
        <Card title="Top Clientes — Monto Total (Producido + En Proceso)">
          <ResponsiveContainer width="100%" height={180}>
            <BarChart data={porCliente} layout="vertical" barCategoryGap="28%">
              <CartesianGrid strokeDasharray="3 3" stroke="#f1f2f5" horizontal={false}/>
              <XAxis type="number" tickFormatter={v=>`S/${(v/1000).toFixed(0)}K`} tick={{fontSize:9,fill:T.muted}}/>
              <YAxis type="category" dataKey="name" width={175} tick={{fontSize:10,fill:T.ink2}}/>
              <Tooltip formatter={v=>[fmtK(v),""]} contentStyle={{borderRadius:8,fontSize:11}}/>
              <Bar dataKey="total" fill={T.blue} radius={[0,5,5,0]}/>
            </BarChart>
          </ResponsiveContainer>
        </Card>
        <Card title="Resumen por Comercial">
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{borderBottom:`2px solid ${T.border}`}}>
              {["Comercial","OPs","Producido (PT)","En Proceso","Total"].map(h=>(
                <th key={h} style={{padding:"7px 10px",textAlign:h==="Comercial"?"left":"right",color:T.muted,fontSize:10,fontWeight:600,textTransform:"uppercase"}}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {porComercial.map((r,i)=>(
                <tr key={i} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.card:"#fafbfc"}}>
                  <td style={{padding:"8px 10px",fontWeight:600,color:T.ink}}>{r.name}</td>
                  <td style={{padding:"8px 10px",textAlign:"right",color:T.muted}}>{r.ops}</td>
                  <td style={{padding:"8px 10px",textAlign:"right",color:T.green,fontWeight:600}}>{fmtK(r.Producido)}</td>
                  <td style={{padding:"8px 10px",textAlign:"right",color:T.amber,fontWeight:600}}>{fmtK(r.Proceso)}</td>
                  <td style={{padding:"8px 10px",textAlign:"right",fontWeight:700,color:T.ink}}>{fmtK(r.Producido+r.Proceso)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </Card>
      </div>
    </div>
  );
}

// ─── ROOT ────────────────────────────────────────────────────────────────────
export default function App() {
  const [page, setPage] = useState("upload");
  const [costosWB, setCostosWB] = useState(null);
  const [reporteWB, setReporteWB] = useState(null);
  const [costosName, setCostosName] = useState("");
  const [reporteName, setReporteName] = useState("");

  const costosRows = useMemo(()=>{
    const parsed = parseCostos(costosWB);
    return parsed.length > 0 ? parsed : (costosWB ? SAMPLE_COSTOS : []);
  },[costosWB]);
  const reporteRows = useMemo(()=>parseReporte(reporteWB),[reporteWB]);

  const hasCostos = !!costosWB;
  const hasDash   = !!reporteWB;

  // Auto-navigate when file loaded
  useEffect(()=>{ if(hasCostos && page==="upload") setPage("costos"); },[hasCostos]);
  useEffect(()=>{ if(hasDash && !hasCostos && page==="upload") setPage("dashboard"); },[hasDash]);

  return (
    <div style={{height:"100vh",display:"flex",overflow:"hidden",background:T.bg}}>
      {/* SIDEBAR — always visible */}
      <Sidebar page={page} setPage={setPage} hasCostos={hasCostos} hasDash={hasDash}/>

      {/* CONTENT */}
      <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",minWidth:0}}>
        {page==="upload"  && <UploadPage costosWB={costosWB} reporteWB={reporteWB} onCostos={(wb,n)=>{setCostosWB(wb);setCostosName(n);}} onReporte={(wb,n)=>{setReporteWB(wb);setReporteName(n);}} costosName={costosName} reporteName={reporteName} setPage={setPage}/>}
        {page==="costos"  && hasCostos  && <CostosPage rows={costosRows}/>}
        {page==="dashboard" && hasDash  && <DashboardPage rows={reporteRows}/>}
        {page==="costos"  && !hasCostos && <UploadPage costosWB={costosWB} reporteWB={reporteWB} onCostos={(wb,n)=>{setCostosWB(wb);setCostosName(n);}} onReporte={(wb,n)=>{setReporteWB(wb);setReporteName(n);}} costosName={costosName} reporteName={reporteName} setPage={setPage}/>}
        {page==="dashboard" && !hasDash && <UploadPage costosWB={costosWB} reporteWB={reporteWB} onCostos={(wb,n)=>{setCostosWB(wb);setCostosName(n);}} onReporte={(wb,n)=>{setReporteWB(wb);setReporteName(n);}} costosName={costosName} reporteName={reporteName} setPage={setPage}/>}
      </div>
    </div>
  );
}
