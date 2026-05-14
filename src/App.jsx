import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ─── Theme ─────────────────────────────────────────────────────────
const C = {
  bg:"#080808", surface:"#101010", card:"#161616", border:"#2a2a2a",
  yellow:"#f0df00", yellowD:"#b8a800",
  green:"#22c55e", red:"#ef4444", blue:"#3b82f6",
  orange:"#f97316", purple:"#a855f7", cyan:"#06b6d4",
  text:"#f5f5f5", muted:"#666", subtle:"#999",
};

const G = `
  @import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Barlow+Condensed:wght@400;600;700;800;900&family=Barlow:wght@400;500;600&display=swap');
  *{box-sizing:border-box;margin:0;padding:0}
  body{background:${C.bg};color:${C.text};font-family:'Barlow',sans-serif;-webkit-tap-highlight-color:transparent}
  input,select,textarea,button{font-family:'Barlow',sans-serif}
  ::-webkit-scrollbar{width:4px;height:4px}
  ::-webkit-scrollbar-track{background:${C.surface}}
  ::-webkit-scrollbar-thumb{background:${C.border};border-radius:2px}
  @keyframes fadeIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
  @keyframes fadeInUp{from{opacity:0;transform:translateY(30px)}to{opacity:1;transform:none}}
  @keyframes spin{to{transform:rotate(360deg)}}
  .anim{animation:fadeIn .25s ease}
  .animUp{animation:fadeInUp .4s ease}
  @media print{.no-print{display:none!important}body{background:#fff;color:#000}}
`;

// ─── Helpers ───────────────────────────────────────────────────────
const fmt = n => Number(n||0).toLocaleString("pt-BR",{style:"currency",currency:"BRL"});
const today = () => new Date().toISOString().split("T")[0];
const fmtDate = iso => { if(!iso)return"—"; const [y,m,d]=iso.split("-"); return `${d}/${m}/${y}`; };
const monthName = m => ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"][m];
const fullMonth = m => ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"][m];
const getDaysInMonth = (y,m) => new Date(y,m+1,0).getDate();
const uid = () => Math.random().toString(36).slice(2,10);

// ─── Storage (Supabase / window.storage) ──────────────────────────
// Detecta se está rodando no Claude.ai (window.storage) ou no Vercel
const isClaudeEnv = () => typeof window !== "undefined" && typeof window.storage !== "undefined";

const KEY = "arced_locacoes_v1";

const DEFAULT = () => ({
  config: { companyName:"ArcD Locações", cnpj:"", contactName:"", contactPhone:"" },
  owners: [
    { id:"hygor", name:"Hygor", color: C.cyan },
    { id:"arcd",  name:"ArcD",  color: C.yellow },
  ],
  obras: [
    { id:uid(), name:"Obra 1", address:"", status:"active" },
    { id:uid(), name:"Obra 2", address:"", status:"active" },
  ],
  equipamentos: [],
  // alocações: { equipId: { obraId, startDate, endDate|null } }
  alocacoes: {},
  // histórico de movimentações
  movimentacoes: [],
});

const loadData = async () => {
  try {
    if (isClaudeEnv()) {
      const r = await window.storage.get(KEY);
      return r ? JSON.parse(r.value) : DEFAULT();
    } else {
      // Vercel + Supabase – mesma lógica do Ponto PRO
      const { createClient } = await import("@supabase/supabase-js");
      const sb = createClient(import.meta.env.VITE_SUPABASE_URL, import.meta.env.VITE_SUPABASE_ANON_KEY);
      const { data } = await sb.from("app_data").select("value").eq("key", KEY).single();
      return data ? JSON.parse(data.value) : DEFAULT();
    }
  } catch { return DEFAULT(); }
};

const saveData = async (obj) => {
  try {
    if (isClaudeEnv()) {
      await window.storage.set(KEY, JSON.stringify(obj));
    } else {
      const { createClient } = await import("@supabase/supabase-js");
      const sb = createClient(import.meta.env.VITE_SUPABASE_URL, import.meta.env.VITE_SUPABASE_ANON_KEY);
      await sb.from("app_data").upsert({ key: KEY, value: JSON.stringify(obj) }, { onConflict:"key" });
    }
  } catch(e) { console.error(e); }
};

// ─── Icons ─────────────────────────────────────────────────────────
const PATHS = {
  home:"M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z M9 22V12h6v10",
  box:"M21 16V8a2 2 0 00-1-1.73l-7-4a2 2 0 00-2 0l-7 4A2 2 0 003 8v8a2 2 0 001 1.73l7 4a2 2 0 002 0l7-4A2 2 0 0021 16z M3.27 6.96L12 12.01l8.73-5.05 M12 22.08V12",
  building:"M2 20h20 M6 20V10l6-6 6 6v10 M9 20v-5h6v5",
  chart:"M18 20V10 M12 20V4 M6 20v-6",
  settings:"M12 15a3 3 0 100-6 3 3 0 000 6z M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-2 2 2 2 0 01-2-2v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 01-2-2 2 2 0 012-2h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 012.83-2.83l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 012-2 2 2 2 0 012 2v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 012 2 2 2 0 01-2 2h-.09a1.65 1.65 0 00-1.51 1z",
  plus:"M12 5v14 M5 12h14",
  check:"M20 6L9 17l-5-5",
  x:"M18 6L6 18 M6 6l12 12",
  trash:"M3 6h18 M8 6V4h8v2 M19 6l-1 14H6L5 6",
  edit:"M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7 M18.5 2.5a2.12 2.12 0 013 3L12 15l-4 1 1-4z",
  excel:"M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z M14 2v6h6",
  print:"M6 9V2h12v7 M6 18H4a2 2 0 01-2-2v-5a2 2 0 012-2h16a2 2 0 012 2v5a2 2 0 01-2 2h-2 M6 14h12v8H6z",
  whatsapp:"M21 11.5a8.38 8.38 0 01-.9 3.8 8.5 8.5 0 01-7.6 4.7 8.38 8.38 0 01-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 01-.9-3.8 8.5 8.5 0 014.7-7.6 8.38 8.38 0 013.8-.9h.5a8.48 8.48 0 018 8v.5z",
  send:"M22 2L11 13 M22 2l-7 20-4-9-9-4 20-7",
  alert:"M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z M12 9v4 M12 17h.01",
  chevR:"M9 18l6-6-6-6",
  move:"M5 9l-3 3 3 3 M9 5l3-3 3 3 M15 19l-3 3-3-3 M19 9l3 3-3 3 M2 12h20 M12 2v20",
  dollar:"M12 1v22 M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6",
  backup:"M20 16v4a2 2 0 01-2 2H6a2 2 0 01-2-2v-4 M7 10l5 5 5-5 M12 15V3",
  upload:"M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4 M17 8l-5-5-5 5 M12 3v12",
  user:"M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2 M12 11a4 4 0 100-8 4 4 0 000 8z",
  logout:"M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4 M16 17l5-5-5-5 M21 12H9",
  history:"M1 4v6h6 M23 20v-6h-6 M20.49 9A9 9 0 005.64 5.64L1 10 M23 14l-4.64 4.36A9 9 0 013.51 15",
};
const Ic = ({n,s=18,style:sx})=>(
  <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={sx}>
    {(PATHS[n]||"").split(" M").map((d,i)=><path key={i} d={i===0?d:"M"+d}/>)}
  </svg>
);

// ─── UI Primitives ─────────────────────────────────────────────────
const Btn = ({children,onClick,v="primary",size="md",disabled,full,style:sx})=>{
  const base={
    display:"inline-flex",alignItems:"center",justifyContent:"center",gap:6,
    fontFamily:"'Barlow Condensed'",fontWeight:700,letterSpacing:".5px",
    border:"none",borderRadius:0,transition:"all .15s",
    cursor:disabled?"not-allowed":"pointer",opacity:disabled?.5:1,
    width:full?"100%":"auto",
    ...(size==="sm"?{padding:"5px 12px",fontSize:13}:size==="lg"?{padding:"14px 24px",fontSize:17}:{padding:"9px 16px",fontSize:15}),
    ...(v==="primary"?{background:C.yellow,color:C.bg}:
        v==="ghost"?{background:"transparent",color:C.subtle,border:`1px solid ${C.border}`}:
        v==="danger"?{background:C.red+"22",color:C.red,border:`1px solid ${C.red}44`}:
        v==="success"?{background:C.green+"22",color:C.green,border:`1px solid ${C.green}44`}:
        v==="info"?{background:C.blue+"22",color:C.blue,border:`1px solid ${C.blue}44`}:
        v==="cyan"?{background:C.cyan+"22",color:C.cyan,border:`1px solid ${C.cyan}44`}:
        v==="warning"?{background:C.yellow+"18",color:C.yellow,border:`1px solid ${C.yellow}44`}:
        {background:C.surface,color:C.text,border:`1px solid ${C.border}`}),
    ...sx
  };
  return <button style={base} onClick={onClick} disabled={disabled}>{children}</button>;
};

const IS={background:C.surface,border:`1.5px solid ${C.border}`,borderRadius:0,padding:"10px 13px",color:C.text,fontSize:15,outline:"none",width:"100%",transition:"border-color .15s"};
const Inp=({label,value,onChange,type="text",placeholder,required,max,min})=>(
  <div style={{display:"flex",flexDirection:"column",gap:5}}>
    {label&&<label style={{fontSize:11,fontWeight:700,color:C.subtle,textTransform:"uppercase",letterSpacing:".7px"}}>{label}{required&&<span style={{color:C.red}}> *</span>}</label>}
    <input type={type} value={value||""} onChange={e=>onChange(e.target.value)} placeholder={placeholder} required={required} max={max} min={min}
      style={IS} onFocus={e=>e.target.style.borderColor=C.yellow} onBlur={e=>e.target.style.borderColor=C.border}/>
  </div>
);
const Sel=({label,value,onChange,options})=>(
  <div style={{display:"flex",flexDirection:"column",gap:5}}>
    {label&&<label style={{fontSize:11,fontWeight:700,color:C.subtle,textTransform:"uppercase",letterSpacing:".7px"}}>{label}</label>}
    <select value={value||""} onChange={e=>onChange(e.target.value)} style={{...IS,appearance:"none"}}
      onFocus={e=>e.target.style.borderColor=C.yellow} onBlur={e=>e.target.style.borderColor=C.border}>
      {options.map(o=><option key={o.v??o} value={o.v??o}>{o.l??o}</option>)}
    </select>
  </div>
);
const Modal=({title,onClose,children,wide})=>(
  <div style={{position:"fixed",inset:0,background:"#000c",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:12}}
    onClick={e=>e.target===e.currentTarget&&onClose()}>
    <div className="anim" style={{background:C.card,border:`1px solid ${C.border}`,padding:22,width:"100%",maxWidth:wide?600:460,maxHeight:"92vh",overflowY:"auto",borderTop:`3px solid ${C.yellow}`}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
        <h2 style={{fontFamily:"'Barlow Condensed'",fontSize:21,fontWeight:800,textTransform:"uppercase",letterSpacing:1}}>{title}</h2>
        <button onClick={onClose} style={{background:"none",border:"none",color:C.muted,cursor:"pointer"}}><Ic n="x"/></button>
      </div>
      {children}
    </div>
  </div>
);
const Badge=({children,color=C.muted})=>(
  <span style={{background:color+"22",color,border:`1px solid ${color}44`,borderRadius:0,padding:"2px 7px",fontSize:11,fontWeight:700,letterSpacing:".3px",whiteSpace:"nowrap"}}>{children}</span>
);
const Divider=()=><div style={{height:1,background:C.border,margin:"14px 0"}}/>;
const Spinner=({small})=><div style={{width:small?14:20,height:small?14:20,border:`2px solid ${C.border}`,borderTopColor:C.yellow,borderRadius:"50%",animation:"spin 1s linear infinite",flexShrink:0}}/>;
const Toast=({toast})=>{
  if(!toast)return null;
  const col=toast.type==="error"?C.red:toast.type==="warn"?C.yellow:C.green;
  return(
    <div style={{position:"fixed",bottom:80,left:"50%",transform:"translateX(-50%)",background:col,color:col===C.yellow?C.bg:"#fff",padding:"10px 20px",fontSize:14,fontWeight:700,zIndex:400,whiteSpace:"nowrap",pointerEvents:"none",fontFamily:"'Barlow Condensed'",letterSpacing:.5,textTransform:"uppercase"}}>
      {toast.msg}
    </div>
  );
};

// ─── Login ────────────────────────────────────────────────────────
function LoginScreen({onEnter}){
  const [nome,setNome]=useState("");
  const [erro,setErro]=useState("");
  const entrar=()=>{ if(!nome.trim()){setErro("Insira seu nome.");return;} onEnter(nome.trim()); };
  return(
    <div className="animUp" style={{minHeight:"100vh",background:C.bg,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24,backgroundImage:`radial-gradient(circle at 50% 0%, ${C.yellow}08 0%, transparent 60%)`}}>
      <style>{G}</style>
      <div style={{fontSize:72,marginBottom:24}}>🔧</div>
      <div style={{width:"100%",maxWidth:360}}>
        <div style={{textAlign:"center",marginBottom:32}}>
          <h1 style={{fontFamily:"'Bebas Neue'",fontSize:28,color:C.yellow,letterSpacing:6,textTransform:"uppercase",marginBottom:4}}>ArcD Locações</h1>
          <p style={{fontSize:12,color:C.muted,letterSpacing:1}}>Gestão de equipamentos e locações</p>
          <div style={{height:1,background:`linear-gradient(90deg,transparent,${C.yellow},transparent)`,marginTop:12}}/>
        </div>
        <div style={{background:C.card,border:`1px solid ${C.border}`,borderTop:`3px solid ${C.yellow}`,padding:28,display:"flex",flexDirection:"column",gap:18}}>
          <div style={{textAlign:"center"}}>
            <p style={{fontSize:15,color:C.text,fontWeight:600}}>Bem-vindo!</p>
            <p style={{fontSize:13,color:C.muted,marginTop:4}}>Qual é o seu nome?</p>
          </div>
          <Inp label="Seu nome" value={nome} onChange={v=>{setNome(v);setErro("");}} placeholder="Ex.: João Silva"/>
          {erro&&<p style={{fontSize:13,color:C.red,fontWeight:600}}>⚠️ {erro}</p>}
          <button onClick={entrar} style={{background:C.yellow,color:C.bg,border:"none",padding:"14px",fontFamily:"'Bebas Neue'",fontSize:20,letterSpacing:3,cursor:"pointer",textTransform:"uppercase",display:"flex",alignItems:"center",justifyContent:"center",gap:10}}>
            <Ic n="check" s={18} style={{color:C.bg}}/> ENTRAR
          </button>
        </div>
        <p style={{textAlign:"center",fontSize:11,color:C.muted,marginTop:20}}>ArcD Construtora · Sistema de Locações</p>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════════════════════
export default function App(){
  const [data,setData]=useState(null);
  const [tab,setTab]=useState("dash");
  const [loading,setLoading]=useState(true);
  const [toast,setToast]=useState(null);
  const [loggedIn,setLoggedIn]=useState(false);

  useEffect(()=>{ loadData().then(d=>{ const r=d||DEFAULT(); setData(r); setLoading(false); if(r.userName?.trim()) setLoggedIn(true); }); },[]);

  const update=useCallback(async nd=>{ setData(nd); await saveData(nd); },[]);
  const showToast=useCallback((msg,type="success")=>{ setToast({msg,type}); setTimeout(()=>setToast(null),3200); },[]);

  const handleEnter=async nome=>{ const nd={...data,userName:nome}; await update(nd); setLoggedIn(true); };
  const handleLogout=()=>{ setLoggedIn(false); setTab("dash"); };

  if(loading) return(
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:C.bg,flexDirection:"column",gap:20}}>
      <style>{G}</style>
      <div style={{fontSize:72}}>🔧</div>
      <Spinner/>
      <p style={{color:C.muted,fontSize:13}}>Carregando dados...</p>
    </div>
  );

  if(!loggedIn) return <><style>{G}</style><LoginScreen onEnter={handleEnter}/></>;

  const TABS=[
    {id:"dash",n:"home",l:"Início"},
    {id:"equipamentos",n:"box",l:"Equip."},
    {id:"obras",n:"building",l:"Obras"},
    {id:"alocacoes",n:"move",l:"Alocações"},
    {id:"relatorios",n:"chart",l:"Relatórios"},
    {id:"config",n:"settings",l:"Config"},
  ];

  const props={data,update,showToast};

  return(
    <div style={{background:C.bg,minHeight:"100vh",maxWidth:480,margin:"0 auto",position:"relative"}}>
      <style>{G}</style>

      {/* Header */}
      <div style={{background:C.surface,borderBottom:`1px solid ${C.border}`,padding:"10px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:50}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <div style={{fontSize:28}}>🔧</div>
          <div>
            <p style={{fontFamily:"'Bebas Neue'",fontSize:13,color:C.yellow,letterSpacing:3,lineHeight:1}}>{data.config?.companyName||"ArcD Locações"}</p>
            <p style={{fontSize:10,color:C.muted}}>{data.userName?" Olá, "+data.userName:"Sistema de Locações"}</p>
          </div>
        </div>
        <button onClick={handleLogout} style={{background:"none",border:`1px solid ${C.border}`,padding:"7px 10px",cursor:"pointer",display:"flex",alignItems:"center",gap:4,color:C.muted}} title="Sair">
          <Ic n="logout" s={15}/>
        </button>
      </div>

      {/* Tab bar */}
      <div style={{background:C.surface,borderBottom:`1px solid ${C.border}`,display:"flex",overflowX:"auto",scrollbarWidth:"none"}}>
        {TABS.map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)} style={{flex:"0 0 auto",padding:"10px 14px",background:"none",border:"none",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:2,color:tab===t.id?C.yellow:C.muted,transition:"color .15s",borderBottom:tab===t.id?`2px solid ${C.yellow}`:"2px solid transparent"}}>
            <Ic n={t.n} s={18}/>
            <span style={{fontSize:9,fontWeight:700,fontFamily:"'Barlow Condensed'",letterSpacing:.7,textTransform:"uppercase"}}>{t.l}</span>
          </button>
        ))}
      </div>

      {/* Content */}
      <div style={{padding:"14px 14px 30px"}}>
        {tab==="dash"        && <Dashboard {...props} onTab={setTab}/>}
        {tab==="equipamentos" && <Equipamentos {...props}/>}
        {tab==="obras"       && <Obras {...props}/>}
        {tab==="alocacoes"   && <Alocacoes {...props}/>}
        {tab==="relatorios"  && <Relatorios {...props}/>}
        {tab==="config"      && <Config {...props} onLogout={handleLogout}/>}
      </div>

      <Toast toast={toast}/>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// DASHBOARD
// ═══════════════════════════════════════════════════════════════════
function Dashboard({data,onTab}){
  const now=new Date(); const m=now.getMonth(); const y=now.getFullYear();
  const equips=data.equipamentos||[];
  const alocados=equips.filter(e=>data.alocacoes[e.id]?.obraId);
  const disponiveis=equips.filter(e=>!data.alocacoes[e.id]?.obraId);
  const totalValorAquisicao=equips.reduce((s,e)=>s+(e.valorAquisicao||0),0);

  // Receita total mensal estimada (todos os equipamentos alocados)
  const receitaMes=equips.reduce((s,eq)=>{
    if(!data.alocacoes[eq.id]?.obraId) return s;
    if(eq.tipoLocacao==="mensal") return s+(eq.valorLocacao||0);
    const dias=getDaysInMonth(y,m);
    return s+(eq.valorLocacao||0)*dias;
  },0);

  // Por proprietário
  const resumoProp=(data.owners||[]).map(o=>{
    const eqs=equips.filter(e=>e.ownerId===o.id&&data.alocacoes[e.id]?.obraId);
    const rec=eqs.reduce((s,eq)=>{
      if(eq.tipoLocacao==="mensal") return s+(eq.valorLocacao||0);
      const dias=getDaysInMonth(y,m);
      return s+(eq.valorLocacao||0)*dias;
    },0);
    return{...o,qtd:eqs.length,receita:rec};
  });

  const Stat=({icon,label,value,sub,color,onClick})=>(
    <div onClick={onClick} style={{background:C.card,border:`1px solid ${C.border}`,padding:"14px",cursor:onClick?"pointer":"default",borderTop:`3px solid ${color}`}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
        <div>
          <p style={{fontSize:10,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:.7}}>{label}</p>
          <p style={{fontSize:26,fontWeight:900,fontFamily:"'Bebas Neue'",color:color,lineHeight:1.1,marginTop:4,letterSpacing:1}}>{value}</p>
          {sub&&<p style={{fontSize:11,color:C.subtle,marginTop:3}}>{sub}</p>}
        </div>
        <div style={{color,opacity:.5}}><Ic n={icon} s={22}/></div>
      </div>
    </div>
  );

  return(
    <div style={{display:"flex",flexDirection:"column",gap:14}} className="anim">
      {/* Period banner */}
      <div style={{background:C.yellow,padding:"10px 16px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div>
          <p style={{fontFamily:"'Bebas Neue'",fontSize:15,color:C.bg,letterSpacing:2}}>RECEITA ESTIMADA — {fullMonth(m).toUpperCase()} {y}</p>
          <p style={{fontSize:11,color:C.bg+"99"}}>{getDaysInMonth(y,m)} dias · {alocados.length} equipamentos ativos</p>
        </div>
        <p style={{fontFamily:"'Bebas Neue'",fontSize:22,color:C.bg,letterSpacing:1}}>{fmt(receitaMes)}</p>
      </div>

      {/* Stats */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        <Stat icon="box" label="Total Equipamentos" value={equips.length} color={C.yellow} sub={`${alocados.length} alocados`} onClick={()=>onTab("equipamentos")}/>
        <Stat icon="check" label="Disponíveis" value={disponiveis.length} color={C.green} sub="prontos p/ locar" onClick={()=>onTab("alocacoes")}/>
        <Stat icon="building" label="Obras Ativas" value={(data.obras||[]).filter(o=>o.status==="active").length} color={C.blue} sub="em andamento" onClick={()=>onTab("obras")}/>
        <Stat icon="dollar" label="Patrimônio" value={fmt(totalValorAquisicao)} color={C.purple} sub="valor de aquisição"/>
      </div>

      {/* Por proprietário */}
      <div style={{background:C.card,border:`1px solid ${C.border}`,padding:14}}>
        <p style={{fontSize:10,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:.7,marginBottom:12}}>Receita por Proprietário — {fullMonth(m)}</p>
        {resumoProp.map(o=>(
          <div key={o.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 0",borderBottom:`1px solid ${C.border}`}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{width:10,height:10,background:o.color,borderRadius:0}}/>
              <div>
                <p style={{fontFamily:"'Barlow Condensed'",fontWeight:800,fontSize:16}}>{o.name}</p>
                <p style={{fontSize:11,color:C.muted}}>{o.qtd} equip. alocados</p>
              </div>
            </div>
            <p style={{fontFamily:"'Bebas Neue'",fontSize:20,color:o.color,letterSpacing:1}}>{fmt(o.receita)}</p>
          </div>
        ))}
      </div>

      {/* Ações rápidas */}
      <div>
        <p style={{fontSize:10,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:.7,marginBottom:8}}>Ações Rápidas</p>
        {[
          {l:"Cadastrar Equipamento",t:"equipamentos",color:C.yellow,icon:"box"},
          {l:"Alocar / Desalocar",t:"alocacoes",color:C.cyan,icon:"move"},
          {l:"Ver Relatório Mensal",t:"relatorios",color:C.green,icon:"chart"},
          {l:"Gerenciar Obras",t:"obras",color:C.blue,icon:"building"},
        ].map(a=>(
          <button key={a.t} onClick={()=>onTab(a.t)} style={{background:C.card,border:`1px solid ${C.border}`,borderLeft:`3px solid ${a.color}`,padding:"12px 14px",display:"flex",alignItems:"center",justifyContent:"space-between",cursor:"pointer",width:"100%",marginBottom:6}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{color:a.color}}><Ic n={a.icon} s={16}/></div>
              <span style={{fontFamily:"'Barlow Condensed'",fontWeight:700,fontSize:15,color:C.text}}>{a.l}</span>
            </div>
            <Ic n="chevR" s={14} style={{color:C.muted}}/>
          </button>
        ))}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// EQUIPAMENTOS
// ═══════════════════════════════════════════════════════════════════
const CATEGORIAS=[
  "Betoneira","Andaime","Materlete / Elevador","Compactador","Gerador","Escora","Esmerilhadeira","Serra Circular","Perfuratriz","Compressor","Vibrador de Concreto","Outro"
];

function Equipamentos({data,update,showToast}){
  const emptyEq={id:"",nome:"",categoria:"Betoneira",ownerId:data.owners?.[0]?.id||"",modelo:"",numeroSerie:"",valorAquisicao:"",valorLocacao:"",tipoLocacao:"mensal",status:"disponivel",observacoes:"",dataAquisicao:""};
  const [modal,setModal]=useState(false);
  const [form,setForm]=useState(emptyEq);
  const [search,setSearch]=useState("");
  const [filterOwner,setFilterOwner]=useState("all");
  const [filterStatus,setFilterStatus]=useState("all");
  const F=k=>v=>setForm(f=>({...f,[k]:v}));

  const save=()=>{
    if(!form.nome.trim()){showToast("Nome obrigatório","error");return;}
    if(!form.valorLocacao){showToast("Valor de locação obrigatório","error");return;}
    const novo={...form,valorAquisicao:parseFloat(form.valorAquisicao)||0,valorLocacao:parseFloat(form.valorLocacao)||0};
    const equipamentos=form.id
      ?data.equipamentos.map(e=>e.id===form.id?novo:e)
      :[...data.equipamentos,{...novo,id:uid()}];
    update({...data,equipamentos});
    setModal(false);
    showToast(form.id?"Equipamento atualizado!":"Equipamento cadastrado!");
  };

  const remove=id=>{
    if(data.alocacoes[id]?.obraId){showToast("Desaloque o equipamento primeiro.","error");return;}
    const equipamentos=data.equipamentos.filter(e=>e.id!==id);
    const alocacoes={...data.alocacoes};delete alocacoes[id];
    update({...data,equipamentos,alocacoes});
    showToast("Equipamento removido.");
  };

  const ownerName=id=>(data.owners||[]).find(o=>o.id===id)?.name||"—";
  const ownerColor=id=>(data.owners||[]).find(o=>o.id===id)?.color||C.muted;
  const obraName=id=>(data.obras||[]).find(o=>o.id===id)?.name||"—";

  const list=(data.equipamentos||[]).filter(e=>{
    const matchOwner=filterOwner==="all"||e.ownerId===filterOwner;
    const alocado=!!data.alocacoes[e.id]?.obraId;
    const statusReal=alocado?"alocado":"disponivel";
    const matchStatus=filterStatus==="all"||filterStatus===statusReal;
    const matchSearch=!search||e.nome.toLowerCase().includes(search.toLowerCase())||e.categoria.toLowerCase().includes(search.toLowerCase());
    return matchOwner&&matchStatus&&matchSearch;
  });

  return(
    <div style={{display:"flex",flexDirection:"column",gap:12}} className="anim">
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div>
          <h2 style={{fontFamily:"'Bebas Neue'",fontSize:28,letterSpacing:2,color:C.yellow}}>Equipamentos</h2>
          <p style={{color:C.muted,fontSize:13}}>{(data.equipamentos||[]).length} cadastrados</p>
        </div>
        <Btn onClick={()=>{setForm({...emptyEq,ownerId:data.owners?.[0]?.id||""});setModal(true)}}><Ic n="plus" s={16}/>Novo</Btn>
      </div>

      <Inp value={search} onChange={setSearch} placeholder="🔍 Buscar equipamento..."/>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
        <Sel value={filterOwner} onChange={setFilterOwner} options={[{v:"all",l:"Todos proprietários"},...(data.owners||[]).map(o=>({v:o.id,l:o.name}))]}/>
        <Sel value={filterStatus} onChange={setFilterStatus} options={[{v:"all",l:"Todos"},{v:"disponivel",l:"Disponíveis"},{v:"alocado",l:"Alocados"}]}/>
      </div>

      {list.length===0&&<div style={{textAlign:"center",padding:"40px 0",color:C.muted}}><div style={{fontSize:36,marginBottom:8}}>🔧</div><p>Nenhum equipamento encontrado.</p></div>}

      {list.map(e=>{
        const alocado=!!data.alocacoes[e.id]?.obraId;
        const obraAtual=data.alocacoes[e.id]?.obraId;
        const oc=ownerColor(e.ownerId);
        return(
          <div key={e.id} style={{background:C.card,border:`1px solid ${C.border}`,borderLeft:`4px solid ${alocado?C.green:C.muted}`}}>
            <div style={{padding:"13px 15px"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                <div style={{flex:1}}>
                  <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",marginBottom:4}}>
                    <span style={{fontFamily:"'Barlow Condensed'",fontWeight:800,fontSize:17}}>{e.nome}</span>
                    <Badge color={oc}>{ownerName(e.ownerId)}</Badge>
                    {alocado?<Badge color={C.green}>Alocado</Badge>:<Badge color={C.muted}>Disponível</Badge>}
                  </div>
                  <p style={{fontSize:12,color:C.muted}}>{e.categoria}{e.modelo?` · ${e.modelo}`:""}</p>
                  {alocado&&<p style={{fontSize:12,color:C.green,marginTop:2}}>📍 {obraName(obraAtual)}</p>}
                  <div style={{display:"flex",gap:12,marginTop:6,flexWrap:"wrap"}}>
                    <span style={{fontSize:14,color:C.yellow,fontWeight:700}}>{fmt(e.valorLocacao)}<span style={{fontSize:10,color:C.muted,fontWeight:400}}>/{e.tipoLocacao==="mensal"?"mês":"dia"}</span></span>
                    {e.valorAquisicao>0&&<span style={{fontSize:12,color:C.subtle}}>Aq: {fmt(e.valorAquisicao)}</span>}
                  </div>
                </div>
                <div style={{display:"flex",gap:5}}>
                  <Btn onClick={()=>{setForm({...e,valorAquisicao:String(e.valorAquisicao),valorLocacao:String(e.valorLocacao)});setModal(true)}} v="ghost" size="sm"><Ic n="edit" s={13}/></Btn>
                  <Btn onClick={()=>remove(e.id)} v="danger" size="sm"><Ic n="trash" s={13}/></Btn>
                </div>
              </div>
              {e.observacoes&&<p style={{fontSize:11,color:C.subtle,marginTop:6,fontStyle:"italic"}}>"{e.observacoes}"</p>}
            </div>
          </div>
        );
      })}

      {modal&&(
        <Modal title={form.id?"Editar Equipamento":"Novo Equipamento"} onClose={()=>setModal(false)} wide>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <div style={{gridColumn:"1/-1"}}><Inp label="Nome do Equipamento *" value={form.nome} onChange={F("nome")} placeholder="Ex.: Betoneira 400L"/></div>
            <Sel label="Categoria *" value={form.categoria} onChange={F("categoria")} options={CATEGORIAS.map(c=>({v:c,l:c}))}/>
            <Inp label="Modelo / Marca" value={form.modelo} onChange={F("modelo")} placeholder="Ex.: Menegotti"/>
            <Inp label="Número de Série" value={form.numeroSerie} onChange={F("numeroSerie")} placeholder="SN-001"/>
            <Inp label="Data Aquisição" type="date" value={form.dataAquisicao} onChange={F("dataAquisicao")}/>
            <Inp label="Valor de Aquisição (R$)" type="number" value={form.valorAquisicao} onChange={F("valorAquisicao")} placeholder="0.00"/>
            <Sel label="Proprietário *" value={form.ownerId} onChange={F("ownerId")} options={(data.owners||[]).map(o=>({v:o.id,l:o.name}))}/>
            <Sel label="Tipo de Locação *" value={form.tipoLocacao} onChange={F("tipoLocacao")} options={[{v:"mensal",l:"Mensal (R$/mês)"},{v:"diario",l:"Diário (R$/dia)"}]}/>
            <Inp label="Valor de Locação (R$) *" type="number" value={form.valorLocacao} onChange={F("valorLocacao")} placeholder="0.00"/>
            <div style={{gridColumn:"1/-1"}}><Inp label="Observações" value={form.observacoes} onChange={F("observacoes")} placeholder="Estado, condições especiais..."/></div>
          </div>
          <div style={{display:"flex",gap:8,marginTop:16}}>
            <Btn v="ghost" onClick={()=>setModal(false)} full>Cancelar</Btn>
            <Btn onClick={save} full><Ic n="check" s={16}/>{form.id?"Salvar":"Cadastrar"}</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// OBRAS
// ═══════════════════════════════════════════════════════════════════
function Obras({data,update,showToast}){
  const empty={id:"",name:"",address:"",engineer:"",startDate:"",status:"active"};
  const [modal,setModal]=useState(false);
  const [form,setForm]=useState(empty);
  const F=k=>v=>setForm(f=>({...f,[k]:v}));
  const save=()=>{
    if(!form.name.trim()){showToast("Nome obrigatório","error");return;}
    const obras=form.id?data.obras.map(o=>o.id===form.id?form:o):[...data.obras,{...form,id:uid()}];
    update({...data,obras});setModal(false);showToast(form.id?"Obra atualizada!":"Obra cadastrada!");
  };
  const remove=id=>{
    const usada=(data.equipamentos||[]).some(e=>data.alocacoes[e.id]?.obraId===id);
    if(usada){showToast("Desaloque os equipamentos primeiro.","error");return;}
    update({...data,obras:(data.obras||[]).filter(o=>o.id!==id)});showToast("Obra removida.");
  };
  const STATUS={active:{l:"Ativa",c:C.green},paused:{l:"Pausada",c:C.yellow},done:{l:"Concluída",c:C.muted}};
  return(
    <div style={{display:"flex",flexDirection:"column",gap:12}} className="anim">
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div><h2 style={{fontFamily:"'Bebas Neue'",fontSize:28,letterSpacing:2,color:C.yellow}}>Obras</h2><p style={{color:C.muted,fontSize:13}}>{(data.obras||[]).length} cadastradas</p></div>
        <Btn onClick={()=>{setForm(empty);setModal(true)}}><Ic n="plus" s={16}/>Nova</Btn>
      </div>
      {(data.obras||[]).map(o=>{
        const s=STATUS[o.status]||STATUS.active;
        const equipsNaObra=(data.equipamentos||[]).filter(e=>data.alocacoes[e.id]?.obraId===o.id);
        return(
          <div key={o.id} style={{background:C.card,border:`1px solid ${C.border}`,borderLeft:`4px solid ${s.c}`,padding:"14px 16px"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
              <div style={{flex:1}}>
                <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
                  <span style={{fontFamily:"'Barlow Condensed'",fontWeight:800,fontSize:17}}>{o.name}</span>
                  <Badge color={s.c}>{s.l}</Badge>
                </div>
                {o.address&&<p style={{fontSize:12,color:C.muted,marginTop:4}}>📍 {o.address}</p>}
                {o.engineer&&<p style={{fontSize:12,color:C.subtle,marginTop:2}}>👷 {o.engineer}</p>}
                <p style={{fontSize:12,color:C.yellow,marginTop:4,fontWeight:700}}>{equipsNaObra.length} equipamento{equipsNaObra.length!==1?"s":""} alocado{equipsNaObra.length!==1?"s":""}</p>
                {equipsNaObra.length>0&&<p style={{fontSize:11,color:C.subtle,marginTop:2}}>{equipsNaObra.map(e=>e.nome).join(", ")}</p>}
              </div>
              <div style={{display:"flex",gap:6}}>
                <Btn onClick={()=>{setForm(o);setModal(true)}} v="ghost" size="sm"><Ic n="edit" s={13}/></Btn>
                <Btn onClick={()=>remove(o.id)} v="danger" size="sm"><Ic n="trash" s={13}/></Btn>
              </div>
            </div>
          </div>
        );
      })}
      {modal&&(
        <Modal title={form.id?"Editar Obra":"Nova Obra"} onClose={()=>setModal(false)}>
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            <Inp label="Nome da Obra *" value={form.name} onChange={F("name")} placeholder="Residencial Aurora"/>
            <Inp label="Endereço" value={form.address} onChange={F("address")} placeholder="Rua, número, bairro..."/>
            <Inp label="Engenheiro Responsável" value={form.engineer} onChange={F("engineer")}/>
            <Inp label="Data de Início" type="date" value={form.startDate} onChange={F("startDate")}/>
            <Sel label="Status" value={form.status} onChange={F("status")} options={[{v:"active",l:"Ativa"},{v:"paused",l:"Pausada"},{v:"done",l:"Concluída"}]}/>
            <div style={{display:"flex",gap:8,marginTop:8}}>
              <Btn v="ghost" onClick={()=>setModal(false)} full>Cancelar</Btn>
              <Btn onClick={save} full><Ic n="check" s={16}/>{form.id?"Salvar":"Cadastrar"}</Btn>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// ALOCAÇÕES
// ═══════════════════════════════════════════════════════════════════
function Alocacoes({data,update,showToast}){
  const [modalAlocar,setModalAlocar]=useState(null);// equipId
  const [obraDestino,setObraDestino]=useState("");
  const [dataInicio,setDataInicio]=useState(today());

  const equips=data.equipamentos||[];
  const alocados=equips.filter(e=>data.alocacoes[e.id]?.obraId);
  const disponiveis=equips.filter(e=>!data.alocacoes[e.id]?.obraId);
  const ownerName=id=>(data.owners||[]).find(o=>o.id===id)?.name||"—";
  const ownerColor=id=>(data.owners||[]).find(o=>o.id===id)?.color||C.muted;
  const obraName=id=>(data.obras||[]).find(o=>o.id===id)?.name||"—";

  const alocar=()=>{
    if(!obraDestino){showToast("Selecione uma obra.","error");return;}
    const alocacoes={...data.alocacoes,[modalAlocar]:{obraId:obraDestino,startDate:dataInicio,endDate:null}};
    const mov=[...(data.movimentacoes||[]),{id:uid(),equipId:modalAlocar,equipNome:equips.find(e=>e.id===modalAlocar)?.nome||"",tipo:"alocacao",obraId:obraDestino,obraNome:obraName(obraDestino),date:today()}];
    update({...data,alocacoes,movimentacoes:mov});
    showToast("Equipamento alocado!");
    setModalAlocar(null);setObraDestino("");setDataInicio(today());
  };

  const desalocar=(equipId)=>{
    const eq=equips.find(e=>e.id===equipId);
    const alocacoes={...data.alocacoes};
    const prevObra=alocacoes[equipId]?.obraId;
    delete alocacoes[equipId];
    const mov=[...(data.movimentacoes||[]),{id:uid(),equipId,equipNome:eq?.nome||"",tipo:"desalocacao",obraId:prevObra,obraNome:obraName(prevObra),date:today()}];
    update({...data,alocacoes,movimentacoes:mov});
    showToast("Equipamento desalocado!");
  };

  const transferir=(equipId,novaObra)=>{
    const eq=equips.find(e=>e.id===equipId);
    const prevObra=data.alocacoes[equipId]?.obraId;
    const alocacoes={...data.alocacoes,[equipId]:{obraId:novaObra,startDate:today(),endDate:null}};
    const mov=[...(data.movimentacoes||[]),{id:uid(),equipId,equipNome:eq?.nome||"",tipo:"transferencia",obraOrigemId:prevObra,obraOrigemNome:obraName(prevObra),obraId:novaObra,obraNome:obraName(novaObra),date:today()}];
    update({...data,alocacoes,movimentacoes:mov});
    showToast(`${eq?.nome||"Equip."} transferido!`);
  };

  const Card=({e,alocado})=>{
    const oc=ownerColor(e.ownerId);
    const obraAtual=data.alocacoes[e.id]?.obraId;
    const startDate=data.alocacoes[e.id]?.startDate;
    return(
      <div style={{background:C.card,border:`1px solid ${C.border}`,borderLeft:`4px solid ${alocado?C.green:C.muted}`,padding:"13px 15px",marginBottom:8}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:alocado?8:0}}>
          <div style={{flex:1}}>
            <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:3}}>
              <span style={{fontFamily:"'Barlow Condensed'",fontWeight:800,fontSize:16}}>{e.nome}</span>
              <Badge color={oc}>{ownerName(e.ownerId)}</Badge>
            </div>
            <p style={{fontSize:12,color:C.muted}}>{e.categoria}</p>
            <span style={{fontSize:13,color:C.yellow,fontWeight:700}}>{fmt(e.valorLocacao)}<span style={{fontSize:10,color:C.muted,fontWeight:400}}>/{e.tipoLocacao==="mensal"?"mês":"dia"}</span></span>
          </div>
          {!alocado&&(
            <Btn onClick={()=>{setModalAlocar(e.id);setObraDestino("");setDataInicio(today());}} size="sm"><Ic n="plus" s={13}/>Alocar</Btn>
          )}
        </div>
        {alocado&&(
          <div style={{borderTop:`1px solid ${C.border}`,paddingTop:8}}>
            <p style={{fontSize:12,color:C.green,marginBottom:6}}>📍 {obraName(obraAtual)}{startDate?` · desde ${fmtDate(startDate)}`:""}</p>
            <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
              <Btn onClick={()=>desalocar(e.id)} v="danger" size="sm"><Ic n="x" s={12}/>Desalocar</Btn>
              <Sel value="" onChange={v=>{ if(v)transferir(e.id,v); }}
                options={[{v:"",l:"→ Transferir para..."},...(data.obras||[]).filter(o=>o.id!==obraAtual&&o.status!=="done").map(o=>({v:o.id,l:o.name}))]}/>
            </div>
          </div>
        )}
      </div>
    );
  };

  return(
    <div style={{display:"flex",flexDirection:"column",gap:12}} className="anim">
      <div>
        <h2 style={{fontFamily:"'Bebas Neue'",fontSize:28,letterSpacing:2,color:C.yellow}}>Alocações</h2>
        <p style={{color:C.muted,fontSize:13}}>{alocados.length} alocados · {disponiveis.length} disponíveis</p>
      </div>

      {/* Alocados */}
      {alocados.length>0&&(
        <div>
          <div style={{background:C.green,padding:"6px 12px",display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
            <Ic n="check" s={13} style={{color:C.bg}}/>
            <span style={{fontFamily:"'Bebas Neue'",letterSpacing:1,color:C.bg,fontSize:14}}>EQUIPAMENTOS ALOCADOS ({alocados.length})</span>
          </div>
          {alocados.map(e=><Card key={e.id} e={e} alocado={true}/>)}
        </div>
      )}

      {/* Disponíveis */}
      {disponiveis.length>0&&(
        <div>
          <div style={{background:C.muted,padding:"6px 12px",display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
            <Ic n="box" s={13} style={{color:"#fff"}}/>
            <span style={{fontFamily:"'Bebas Neue'",letterSpacing:1,color:"#fff",fontSize:14}}>DISPONÍVEIS ({disponiveis.length})</span>
          </div>
          {disponiveis.map(e=><Card key={e.id} e={e} alocado={false}/>)}
        </div>
      )}

      {equips.length===0&&<div style={{textAlign:"center",padding:"40px 0",color:C.muted}}><div style={{fontSize:36,marginBottom:8}}>🔧</div><p>Nenhum equipamento cadastrado.</p></div>}

      {/* Histórico recente */}
      {(data.movimentacoes||[]).length>0&&(
        <div style={{background:C.card,border:`1px solid ${C.border}`,padding:14,marginTop:8}}>
          <p style={{fontSize:10,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:.7,marginBottom:12}}>Histórico Recente</p>
          {[...(data.movimentacoes||[])].reverse().slice(0,10).map(m=>(
            <div key={m.id} style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",padding:"7px 0",borderBottom:`1px solid ${C.border}`}}>
              <div>
                <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:2}}>
                  {m.tipo==="alocacao"&&<Badge color={C.green}>↓ Alocação</Badge>}
                  {m.tipo==="desalocacao"&&<Badge color={C.red}>↑ Devolução</Badge>}
                  {m.tipo==="transferencia"&&<Badge color={C.yellow}>↗ Transferência</Badge>}
                  <span style={{fontSize:13,fontWeight:700}}>{m.equipNome}</span>
                </div>
                {m.tipo==="alocacao"&&<p style={{fontSize:11,color:C.muted}}>→ {m.obraNome}</p>}
                {m.tipo==="desalocacao"&&<p style={{fontSize:11,color:C.muted}}>← de {m.obraNome}</p>}
                {m.tipo==="transferencia"&&<p style={{fontSize:11,color:C.muted}}>{m.obraOrigemNome} → <b style={{color:C.yellow}}>{m.obraNome}</b></p>}
              </div>
              <span style={{fontSize:11,color:C.muted,flexShrink:0,marginLeft:8}}>{fmtDate(m.date)}</span>
            </div>
          ))}
        </div>
      )}

      {/* Modal Alocar */}
      {modalAlocar&&(
        <Modal title="Alocar Equipamento" onClose={()=>setModalAlocar(null)}>
          <div style={{marginBottom:16,background:C.surface,border:`1px solid ${C.border}`,padding:12}}>
            <p style={{fontFamily:"'Barlow Condensed'",fontWeight:800,fontSize:16}}>{equips.find(e=>e.id===modalAlocar)?.nome||""}</p>
            <p style={{fontSize:12,color:C.muted}}>{equips.find(e=>e.id===modalAlocar)?.categoria||""}</p>
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            <Sel label="Obra de Destino *" value={obraDestino} onChange={setObraDestino} options={[{v:"",l:"— Selecionar obra —"},...(data.obras||[]).filter(o=>o.status!=="done").map(o=>({v:o.id,l:o.name}))]}/>
            <Inp label="Data de Início" type="date" value={dataInicio} onChange={setDataInicio}/>
            <div style={{display:"flex",gap:8,marginTop:6}}>
              <Btn v="ghost" onClick={()=>setModalAlocar(null)} full>Cancelar</Btn>
              <Btn onClick={alocar} full><Ic n="check" s={16}/>Alocar</Btn>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// RELATÓRIOS
// ═══════════════════════════════════════════════════════════════════
function Relatorios({data,showToast}){
  const now=new Date();
  const [month,setMonth]=useState(now.getMonth());
  const [year,setYear]=useState(now.getFullYear());

  const equips=data.equipamentos||[];
  const obras=data.obras||[];
  const owners=data.owners||[];
  const diasMes=getDaysInMonth(year,month);

  // Para cada equip, calcula valor mensal e verifica se estava alocado
  const calcValorMensal=eq=>{
    if(!data.alocacoes[eq.id]?.obraId) return 0;
    if(eq.tipoLocacao==="mensal") return eq.valorLocacao||0;
    return (eq.valorLocacao||0)*diasMes;
  };

  // Por obra
  const relObras=obras.map(o=>{
    const eqs=equips.filter(e=>data.alocacoes[e.id]?.obraId===o.id);
    const total=eqs.reduce((s,e)=>s+calcValorMensal(e),0);
    const detalhe=eqs.map(e=>({...e,valorMes:calcValorMensal(e),ownerName:owners.find(ow=>ow.id===e.ownerId)?.name||"—"}));
    return{...o,equips:detalhe,total};
  }).filter(o=>o.equips.length>0);

  // Por proprietário
  const relOwners=owners.map(ow=>{
    const eqs=equips.filter(e=>e.ownerId===ow.id&&data.alocacoes[e.id]?.obraId);
    const total=eqs.reduce((s,e)=>s+calcValorMensal(e),0);
    const detalhe=eqs.map(e=>({...e,valorMes:calcValorMensal(e),obraNome:obras.find(o=>o.id===data.alocacoes[e.id]?.obraId)?.name||"—"}));
    return{...ow,equips:detalhe,total};
  });

  const totalGeral=relOwners.reduce((s,o)=>s+o.total,0);

  const [expanded,setExpanded]=useState({});
  const toggle=k=>setExpanded(p=>({...p,[k]:!p[k]}));

  const exportXLS=()=>{
    const wb=XLSX.utils.book_new();
    const period=`${fullMonth(month)} ${year}`;

    // Aba por obra
    const hObra=["Obra","Equipamento","Categoria","Proprietário","Tipo Locação","Valor Unit.","Dias","Valor Mês"];
    const bObra=[];
    relObras.forEach(o=>o.equips.forEach(e=>{
      bObra.push([o.name,e.nome,e.categoria,e.ownerName,e.tipoLocacao==="mensal"?"Mensal":"Diário",e.valorLocacao,e.tipoLocacao==="mensal"?1:diasMes,e.valorMes]);
    }));
    const ws1=XLSX.utils.aoa_to_sheet([[`Relatório por Obra — ${period}`],[""],hObra,...bObra,["","","","","","","TOTAL GERAL:",totalGeral]]);
    ws1["!cols"]=[20,22,16,12,10,12,6,14].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb,ws1,"Por Obra");

    // Aba por proprietário
    const hOwn=["Proprietário","Equipamento","Categoria","Obra","Tipo Locação","Valor Unit.","Dias","Valor Mês"];
    const bOwn=[];
    relOwners.forEach(ow=>ow.equips.forEach(e=>{
      bOwn.push([ow.name,e.nome,e.categoria,e.obraNome,e.tipoLocacao==="mensal"?"Mensal":"Diário",e.valorLocacao,e.tipoLocacao==="mensal"?1:diasMes,e.valorMes]);
    }));
    const ws2=XLSX.utils.aoa_to_sheet([[`Relatório por Proprietário — ${period}`],[""],hOwn,...bOwn,["","","","","","","TOTAL GERAL:",totalGeral]]);
    ws2["!cols"]=[14,22,16,20,10,12,6,14].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb,ws2,"Por Proprietário");

    XLSX.writeFile(wb,`arced-locacoes-${year}-${String(month+1).padStart(2,"0")}.xlsx`);
    showToast("Excel gerado!");
  };

  const imprimirPDF=()=>{
    const period=`${fullMonth(month)} ${year}`;
    let html=`<!DOCTYPE html><html><head><title>Relatório Locações — ${period}</title>
    <style>*{box-sizing:border-box}body{font-family:Arial,sans-serif;padding:24px;color:#111}
    h1{font-size:22px;margin-bottom:4px}h2{font-size:15px;color:#555;font-weight:400;margin-bottom:20px}
    h3{font-size:14px;background:#f0df00;color:#080808;padding:6px 10px;margin:16px 0 8px}
    table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:12px}
    th{background:#080808;color:#f0df00;padding:6px 8px;text-align:left}td{border-bottom:1px solid #eee;padding:6px 8px}
    .tot{font-weight:bold;background:#fffde7}.right{text-align:right}
    @media print{button{display:none}}</style></head><body>
    <button onclick="window.print()" style="margin-bottom:16px;padding:8px 16px;background:#f0df00;border:none;cursor:pointer;font-weight:700">🖨️ Imprimir / PDF</button>
    <h1>ArcD Locações — Relatório Mensal</h1><h2>${period} · Total: R$ ${totalGeral.toFixed(2).replace(".",",")}</h2>
    <h3>📋 COBRANÇAS POR OBRA</h3>`;
    relObras.forEach(o=>{
      html+=`<p style="font-weight:700;margin:10px 0 4px">${o.name} — Total: R$ ${o.total.toFixed(2).replace(".",",")}</p>
      <table><thead><tr><th>Equipamento</th><th>Categoria</th><th>Proprietário</th><th>Tipo</th><th class="right">Valor/Período</th></tr></thead><tbody>`;
      o.equips.forEach(e=>{ html+=`<tr><td>${e.nome}</td><td>${e.categoria}</td><td>${e.ownerName}</td><td>${e.tipoLocacao==="mensal"?"Mensal":"Diário"}</td><td class="right">R$ ${e.valorMes.toFixed(2).replace(".",",")}</td></tr>`; });
      html+=`<tr class="tot"><td colspan="4">SUBTOTAL ${o.name}</td><td class="right">R$ ${o.total.toFixed(2).replace(".",",")}</td></tr></tbody></table>`;
    });
    html+=`<h3>💰 PAGAMENTOS POR PROPRIETÁRIO</h3>`;
    relOwners.forEach(ow=>{
      html+=`<p style="font-weight:700;margin:10px 0 4px">${ow.name} — Total a Receber: R$ ${ow.total.toFixed(2).replace(".",",")}</p>`;
      if(ow.equips.length>0){
        html+=`<table><thead><tr><th>Equipamento</th><th>Categoria</th><th>Obra</th><th>Tipo</th><th class="right">Valor</th></tr></thead><tbody>`;
        ow.equips.forEach(e=>{ html+=`<tr><td>${e.nome}</td><td>${e.categoria}</td><td>${e.obraNome}</td><td>${e.tipoLocacao==="mensal"?"Mensal":"Diário"}</td><td class="right">R$ ${e.valorMes.toFixed(2).replace(".",",")}</td></tr>`; });
        html+=`<tr class="tot"><td colspan="4">SUBTOTAL ${ow.name}</td><td class="right">R$ ${ow.total.toFixed(2).replace(".",",")}</td></tr></tbody></table>`;
      }else{html+=`<p style="color:#888;font-size:12px">Nenhum equipamento alocado.</p>`;}
    });
    html+=`<p style="margin-top:20px;font-size:11px;color:#777">Gerado em ${new Date().toLocaleDateString("pt-BR")}</p></body></html>`;
    const w=window.open("","_blank");w.document.write(html);w.document.close();
  };

  const copiarWpp=()=>{
    let txt=`🔧 RELATÓRIO LOCAÇÕES — ${fullMonth(month).toUpperCase()} ${year}\n\n`;
    txt+=`📋 COBRANÇAS POR OBRA:\n`;
    relObras.forEach(o=>{
      txt+=`\n▪ ${o.name}: ${fmt(o.total)}\n`;
      o.equips.forEach(e=>{ txt+=`  - ${e.nome} (${e.ownerName}): ${fmt(e.valorMes)}\n`; });
    });
    txt+=`\n💰 PAGAMENTOS:\n`;
    relOwners.forEach(ow=>{ txt+=`  ${ow.name}: ${fmt(ow.total)}\n`; });
    txt+=`\n💵 TOTAL GERAL: ${fmt(totalGeral)}`;
    navigator.clipboard.writeText(txt).then(()=>showToast("Copiado!")).catch(()=>showToast("Erro","error"));
  };

  const years=[];for(let i=now.getFullYear()-1;i<=now.getFullYear()+1;i++)years.push({v:String(i),l:String(i)});

  return(
    <div style={{display:"flex",flexDirection:"column",gap:14}} className="anim">
      <div><h2 style={{fontFamily:"'Bebas Neue'",fontSize:28,letterSpacing:2,color:C.yellow}}>Relatório Mensal</h2>
        <p style={{color:C.muted,fontSize:13}}>Cobranças e pagamentos</p></div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
        <Sel label="Mês" value={String(month)} onChange={v=>setMonth(Number(v))} options={Array.from({length:12},(_,i)=>({v:String(i),l:monthName(i)}))}/>
        <Sel label="Ano" value={String(year)} onChange={v=>setYear(Number(v))} options={years}/>
      </div>

      {/* Resumo total */}
      <div style={{background:C.yellow,padding:"16px 20px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div>
          <p style={{fontFamily:"'Bebas Neue'",fontSize:12,color:C.bg+"99",letterSpacing:2}}>TOTAL A COBRAR DAS OBRAS</p>
          <p style={{fontFamily:"'Bebas Neue'",fontSize:36,color:C.bg,letterSpacing:1,lineHeight:1}}>{fmt(totalGeral)}</p>
          <p style={{fontSize:11,color:C.bg+"88",marginTop:2}}>{fullMonth(month)} {year} · {diasMes} dias</p>
        </div>
        <div style={{textAlign:"right"}}>
          {relOwners.map(ow=>(
            <p key={ow.id} style={{fontSize:12,color:C.bg+"99"}}>{ow.name}: {fmt(ow.total)}</p>
          ))}
        </div>
      </div>

      {/* Exportar */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
        <Btn onClick={imprimirPDF} v="ghost"><Ic n="print" s={15}/>PDF / Imprimir</Btn>
        <Btn onClick={exportXLS} v="ghost" style={{color:C.green,borderColor:C.green+"44"}}><Ic n="excel" s={15}/>Excel .xlsx</Btn>
        <Btn onClick={copiarWpp} v="success" style={{gridColumn:"1/-1"}}><Ic n="whatsapp" s={15}/>Copiar Resumo WhatsApp</Btn>
      </div>

      {/* Seção por Obra */}
      <div>
        <div style={{background:C.surface,border:`1px solid ${C.border}`,borderTop:`3px solid ${C.yellow}`,padding:"10px 14px",marginBottom:2}}>
          <p style={{fontFamily:"'Bebas Neue'",fontSize:15,color:C.yellow,letterSpacing:2}}>📋 COBRANÇAS POR OBRA</p>
          <p style={{fontSize:12,color:C.muted}}>Valor a cobrar de cada obra</p>
        </div>
        {relObras.length===0&&<div style={{background:C.card,border:`1px solid ${C.border}`,padding:"20px",textAlign:"center",color:C.muted}}><p>Nenhum equipamento alocado em obras.</p></div>}
        {relObras.map(o=>(
          <div key={o.id} style={{marginBottom:4}}>
            <button onClick={()=>toggle(`obra_${o.id}`)} style={{background:C.card,border:`1px solid ${C.border}`,padding:"13px 15px",width:"100%",cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center",borderLeft:`4px solid ${C.yellow}`}}>
              <div style={{textAlign:"left"}}>
                <p style={{fontFamily:"'Barlow Condensed'",fontWeight:800,fontSize:16}}>{o.name}</p>
                <p style={{fontSize:11,color:C.muted}}>{o.equips.length} equipamento{o.equips.length!==1?"s":""}</p>
              </div>
              <div style={{display:"flex",alignItems:"center",gap:8}}>
                <span style={{fontFamily:"'Bebas Neue'",fontSize:20,color:C.yellow,letterSpacing:1}}>{fmt(o.total)}</span>
                <Ic n="chevR" s={15} style={{color:C.muted,transform:expanded[`obra_${o.id}`]?"rotate(90deg)":"none",transition:"transform .2s"}}/>
              </div>
            </button>
            {expanded[`obra_${o.id}`]&&(
              <div style={{background:C.surface,border:`1px solid ${C.border}`,borderTop:"none",padding:"10px 15px"}}>
                {o.equips.map(e=>(
                  <div key={e.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:`1px solid ${C.border}`}}>
                    <div>
                      <p style={{fontSize:13,fontWeight:600}}>{e.nome}</p>
                      <p style={{fontSize:11,color:C.muted}}>{e.categoria} · <span style={{color:ownerColor(e.ownerId)}}>{e.ownerName}</span> · {e.tipoLocacao==="mensal"?"Mensal":"Diário"}</p>
                    </div>
                    <span style={{color:C.yellow,fontWeight:700,fontSize:14}}>{fmt(e.valorMes)}</span>
                  </div>
                ))}
              </div>
            )}
          </div>
        ))}
      </div>

      {/* Seção por Proprietário */}
      <div>
        <div style={{background:C.surface,border:`1px solid ${C.border}`,borderTop:`3px solid ${C.cyan}`,padding:"10px 14px",marginBottom:2}}>
          <p style={{fontFamily:"'Bebas Neue'",fontSize:15,color:C.cyan,letterSpacing:2}}>💰 PAGAMENTOS POR PROPRIETÁRIO</p>
          <p style={{fontSize:12,color:C.muted}}>Valor a pagar para cada sócio</p>
        </div>
        {relOwners.map(ow=>(
          <div key={ow.id} style={{marginBottom:4}}>
            <button onClick={()=>toggle(`own_${ow.id}`)} style={{background:C.card,border:`1px solid ${C.border}`,padding:"13px 15px",width:"100%",cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center",borderLeft:`4px solid ${ow.color}`}}>
              <div style={{textAlign:"left"}}>
                <p style={{fontFamily:"'Barlow Condensed'",fontWeight:800,fontSize:16}}>{ow.name}</p>
                <p style={{fontSize:11,color:C.muted}}>{ow.equips.length} equipamento{ow.equips.length!==1?"s":""} alocado{ow.equips.length!==1?"s":""}</p>
              </div>
              <div style={{display:"flex",alignItems:"center",gap:8}}>
                <span style={{fontFamily:"'Bebas Neue'",fontSize:20,color:ow.color,letterSpacing:1}}>{fmt(ow.total)}</span>
                <Ic n="chevR" s={15} style={{color:C.muted,transform:expanded[`own_${ow.id}`]?"rotate(90deg)":"none",transition:"transform .2s"}}/>
              </div>
            </button>
            {expanded[`own_${ow.id}`]&&(
              <div style={{background:C.surface,border:`1px solid ${C.border}`,borderTop:"none",padding:"10px 15px"}}>
                {ow.equips.length===0&&<p style={{fontSize:12,color:C.muted}}>Nenhum equipamento alocado este mês.</p>}
                {ow.equips.map(e=>(
                  <div key={e.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:`1px solid ${C.border}`}}>
                    <div>
                      <p style={{fontSize:13,fontWeight:600}}>{e.nome}</p>
                      <p style={{fontSize:11,color:C.muted}}>{e.categoria} · {e.obraNome} · {e.tipoLocacao==="mensal"?"Mensal":"Diário"}</p>
                    </div>
                    <span style={{color:ow.color,fontWeight:700,fontSize:14}}>{fmt(e.valorMes)}</span>
                  </div>
                ))}
              </div>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// CONFIG
// ═══════════════════════════════════════════════════════════════════
function Config({data,update,showToast,onLogout}){
  const [form,setForm]=useState(data.config||{});
  const F=k=>v=>setForm(f=>({...f,[k]:v}));
  const save=()=>{ update({...data,config:form}); showToast("Configurações salvas!"); };
  const importRef=useRef();
  const exportBackup=()=>{
    const blob=new Blob([JSON.stringify(data,null,2)],{type:"application/json"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");a.href=url;a.download=`arced-locacoes-backup-${today()}.json`;a.click();
    URL.revokeObjectURL(url);showToast("Backup exportado!");
  };
  const importBackup=e=>{
    const file=e.target.files[0];if(!file)return;
    const reader=new FileReader();
    reader.onload=ev=>{try{update(JSON.parse(ev.target.result));showToast("Backup restaurado!");}catch{showToast("Arquivo inválido","error");}};
    reader.readAsText(file);e.target.value="";
  };

  // Proprietários
  const [ownerForm,setOwnerForm]=useState({name:""});
  const addOwner=()=>{
    if(!ownerForm.name.trim()){showToast("Nome obrigatório","error");return;}
    const COLORS=[C.cyan,C.yellow,C.green,C.orange,C.purple,C.blue,C.red];
    const color=COLORS[(data.owners||[]).length%COLORS.length];
    const owners=[...(data.owners||[]),{id:uid(),name:ownerForm.name.trim(),color}];
    update({...data,owners});setOwnerForm({name:""});showToast("Proprietário adicionado!");
  };
  const removeOwner=id=>{
    const emUso=(data.equipamentos||[]).some(e=>e.ownerId===id);
    if(emUso){showToast("Proprietário tem equipamentos.","error");return;}
    update({...data,owners:(data.owners||[]).filter(o=>o.id!==id)});showToast("Proprietário removido.");
  };

  return(
    <div style={{display:"flex",flexDirection:"column",gap:14}} className="anim">
      <div><h2 style={{fontFamily:"'Bebas Neue'",fontSize:28,letterSpacing:2,color:C.yellow}}>Configurações</h2></div>

      {/* Dados empresa */}
      <div style={{background:C.card,border:`1px solid ${C.border}`,borderTop:`3px solid ${C.yellow}`,padding:16}}>
        <p style={{fontSize:11,fontWeight:700,color:C.yellow,textTransform:"uppercase",letterSpacing:.7,marginBottom:12}}>Dados da Empresa</p>
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <Inp label="Nome da Empresa" value={form.companyName} onChange={F("companyName")} placeholder="ArcD Locações"/>
          <Inp label="CNPJ" value={form.cnpj} onChange={F("cnpj")} placeholder="00.000.000/0001-00"/>
          <Inp label="Responsável" value={form.contactName} onChange={F("contactName")} placeholder="Nome do responsável"/>
          <Inp label="Telefone" value={form.contactPhone} onChange={F("contactPhone")} placeholder="(00) 00000-0000"/>
        </div>
        <div style={{marginTop:14}}><Btn onClick={save} full><Ic n="check" s={16}/>Salvar</Btn></div>
      </div>

      {/* Proprietários */}
      <div style={{background:C.card,border:`1px solid ${C.border}`,borderTop:`3px solid ${C.cyan}`,padding:16}}>
        <p style={{fontSize:11,fontWeight:700,color:C.cyan,textTransform:"uppercase",letterSpacing:.7,marginBottom:12}}>Proprietários dos Equipamentos</p>
        {(data.owners||[]).map(o=>(
          <div key={o.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:`1px solid ${C.border}`}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{width:10,height:10,background:o.color}}/>
              <span style={{fontFamily:"'Barlow Condensed'",fontWeight:700,fontSize:16}}>{o.name}</span>
              <span style={{fontSize:11,color:C.muted}}>({(data.equipamentos||[]).filter(e=>e.ownerId===o.id).length} equip.)</span>
            </div>
            <Btn onClick={()=>removeOwner(o.id)} v="danger" size="sm"><Ic n="trash" s={12}/></Btn>
          </div>
        ))}
        <div style={{display:"flex",gap:8,marginTop:12}}>
          <div style={{flex:1}}><Inp value={ownerForm.name} onChange={v=>setOwnerForm({name:v})} placeholder="Nome do proprietário"/></div>
          <Btn onClick={addOwner}><Ic n="plus" s={16}/>Add</Btn>
        </div>
      </div>

      {/* Sessão */}
      <div style={{background:C.card,border:`1px solid ${C.border}`,borderTop:`3px solid ${C.yellow}`,padding:16}}>
        <p style={{fontSize:11,fontWeight:700,color:C.yellow,textTransform:"uppercase",letterSpacing:.7,marginBottom:12}}>Sessão</p>
        <div style={{background:C.surface,border:`1px solid ${C.border}`,padding:"12px 14px",marginBottom:10,display:"flex",alignItems:"center",gap:10}}>
          <Ic n="user" s={18} style={{color:C.yellow}}/>
          <div>
            <p style={{fontSize:12,color:C.muted}}>Conectado como</p>
            <p style={{fontSize:15,fontWeight:700,color:C.text}}>{data.userName||"Usuário"}</p>
          </div>
        </div>
        <Btn onClick={onLogout} v="ghost" full><Ic n="logout" s={16}/>Trocar de Usuário</Btn>
      </div>

      {/* Backup */}
      <div style={{background:C.card,border:`1px solid ${C.border}`,borderTop:`3px solid ${C.yellow}`,padding:16}}>
        <p style={{fontSize:11,fontWeight:700,color:C.yellow,textTransform:"uppercase",letterSpacing:.7,marginBottom:12}}>Backup de Dados</p>
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          <Btn onClick={exportBackup} v="warning" full><Ic n="backup" s={16}/>Exportar Backup JSON</Btn>
          <input type="file" accept=".json" ref={importRef} onChange={importBackup} style={{display:"none"}}/>
          <Btn onClick={()=>importRef.current.click()} v="ghost" full><Ic n="upload" s={16}/>Importar Backup</Btn>
        </div>
        <p style={{fontSize:11,color:C.muted,marginTop:10}}>⚠️ Importar substitui todos os dados atuais.</p>
      </div>

      {/* Estatísticas */}
      <div style={{background:C.card,border:`1px solid ${C.border}`,padding:16}}>
        <p style={{fontSize:11,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:.7,marginBottom:10}}>Estatísticas</p>
        {[
          ["Proprietários",(data.owners||[]).length],
          ["Equipamentos",(data.equipamentos||[]).length],
          ["Alocados",(data.equipamentos||[]).filter(e=>data.alocacoes[e.id]?.obraId).length],
          ["Obras",(data.obras||[]).length],
          ["Movimentações",(data.movimentacoes||[]).length],
        ].map(([l,v])=>(
          <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${C.border}`}}>
            <span style={{fontSize:13,color:C.subtle}}>{l}</span>
            <span style={{fontSize:13,fontWeight:700,color:C.yellow}}>{v}</span>
          </div>
        ))}
      </div>
    </div>
  );
}
