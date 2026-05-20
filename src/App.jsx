import { useState, useMemo, useRef, useCallback } from "react"
import * as XLSX from "xlsx"
import {
  ComposedChart, AreaChart, Area, BarChart, Bar, LineChart, Line,
  PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip,
  Legend, ResponsiveContainer, ReferenceLine
} from "recharts"
import {
  LayoutDashboard, Settings as SettingsIcon, TrendingUp, TrendingDown, ArrowUpRight, ArrowDownRight,
  Plus, Search, X, Clock, Activity, Calculator, Target, FileText,
  Bell, Moon, Sun, ChevronLeft, ChevronRight, Sparkles, CircleDollarSign,
  Wallet, Layers, ArrowLeftRight, TriangleAlert, BadgeCheck, PiggyBank,
  Upload, FileUp, CheckCircle2, AlertCircle, Zap, RefreshCw, Brain,
  Trash2, ShieldAlert, Package, Check, FileScan, Wand2, Lightbulb
} from "lucide-react"

// ─── STATIC COLOR MAP (Tailwind needs static class names) ────────────────────
const CLR = {
  emerald: { card:"bg-emerald-500/10 border border-emerald-500/20", text:"text-emerald-400", badge:"bg-emerald-500/10 border-emerald-500/20 text-emerald-400", alert:"bg-emerald-500/5 border-emerald-500/20 text-emerald-300" },
  rose:    { card:"bg-rose-500/10 border border-rose-500/20",       text:"text-rose-400",    badge:"bg-rose-500/10 border-rose-500/20 text-rose-400",       alert:"bg-rose-500/5 border-rose-500/20 text-rose-300" },
  amber:   { card:"bg-amber-500/10 border border-amber-500/20",     text:"text-amber-400",   badge:"bg-amber-500/10 border-amber-500/20 text-amber-400",     alert:"bg-amber-500/5 border-amber-500/20 text-amber-300" },
  indigo:  { card:"bg-indigo-500/10 border border-indigo-500/20",   text:"text-indigo-400",  badge:"bg-indigo-500/10 border-indigo-500/20 text-indigo-400",   alert:"bg-indigo-500/5 border-indigo-500/20 text-indigo-300" },
  violet:  { card:"bg-violet-500/10 border border-violet-500/20",   text:"text-violet-400",  badge:"bg-violet-500/10 border-violet-500/20 text-violet-400",   alert:"bg-violet-500/5 border-violet-500/20 text-violet-300" },
  sky:     { card:"bg-sky-500/10 border border-sky-500/20",         text:"text-sky-400",     badge:"bg-sky-500/10 border-sky-500/20 text-sky-400",             alert:"bg-sky-500/5 border-sky-500/20 text-sky-300" },
}

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const MONTHS = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
const CATS = {
  entrada: ["Serviços","Eventos","Vendas de Produtos","Reembolsos","Receitas Extras","Outros"],
  saida:   ["Aluguel","Energia/Água","Folha de Pagamento","Freelancers","Impostos","Marketing","Manutenção","Materiais","Plataformas/Software","Equipamentos","Despesas Operacionais","Outros"]
}
const PAYMENTS = ["PIX","Transferência","Cartão de Débito","Cartão de Crédito","Boleto","Dinheiro"]
const PIE_COLORS = ["#ef4444","#f59e0b","#6366f1","#8b5cf6","#10b981","#14b8a6","#f97316","#ec4899"]
const CHART_COLORS = { receita:"#10b981", custos:"#ef4444", lucro:"#818cf8", saldo:"#818cf8", margem:"#a78bfa" }

// ─── SEED DATA ────────────────────────────────────────────────────────────────
const SEED = [
  {id:1,name:"Festa Infantil - Família Silva",category:"Eventos",type:"entrada",value:1800,date:"2025-01-05",payment:"PIX",obs:"Pacote 3h",competencia:"2025-01",recurrence:"none",status:"pago"},
  {id:2,name:"Aluguel do espaço",category:"Aluguel",type:"saida",value:2200,date:"2025-01-10",payment:"Boleto",obs:"",competencia:"2025-01",recurrence:"mensal",status:"pago"},
  {id:3,name:"Aniversário Corporativo",category:"Eventos",type:"entrada",value:3500,date:"2025-01-12",payment:"Transferência",obs:"50 convidados",competencia:"2025-01",recurrence:"none",status:"pago"},
  {id:4,name:"Energia Elétrica",category:"Energia/Água",type:"saida",value:480,date:"2025-01-15",payment:"Boleto",obs:"",competencia:"2025-01",recurrence:"mensal",status:"pago"},
  {id:5,name:"Folha - Funcionários",category:"Folha de Pagamento",type:"saida",value:4800,date:"2025-01-30",payment:"Transferência",obs:"3 funcionários",competencia:"2025-01",recurrence:"mensal",status:"pago"},
  {id:6,name:"Pacote Mensal Kids",category:"Serviços",type:"entrada",value:2400,date:"2025-01-18",payment:"PIX",obs:"12x R$200",competencia:"2025-01",recurrence:"mensal",status:"pago"},
  {id:7,name:"Instagram Ads",category:"Marketing",type:"saida",value:600,date:"2025-01-20",payment:"Cartão de Crédito",obs:"",competencia:"2025-01",recurrence:"mensal",status:"pago"},
  {id:8,name:"Materiais decoração",category:"Materiais",type:"saida",value:320,date:"2025-01-22",payment:"PIX",obs:"Balões, fitas",competencia:"2025-01",recurrence:"none",status:"pago"},
  {id:9,name:"Festa Infantil - Família Costa",category:"Eventos",type:"entrada",value:2200,date:"2025-02-03",payment:"PIX",obs:"",competencia:"2025-02",recurrence:"none",status:"pago"},
  {id:10,name:"Aluguel do espaço",category:"Aluguel",type:"saida",value:2200,date:"2025-02-10",payment:"Boleto",obs:"",competencia:"2025-02",recurrence:"mensal",status:"pago"},
  {id:11,name:"Evento Carnaval Kids",category:"Eventos",type:"entrada",value:4800,date:"2025-02-15",payment:"PIX",obs:"Edição especial",competencia:"2025-02",recurrence:"none",status:"pago"},
  {id:12,name:"Energia Elétrica",category:"Energia/Água",type:"saida",value:510,date:"2025-02-15",payment:"Boleto",obs:"",competencia:"2025-02",recurrence:"mensal",status:"pago"},
  {id:13,name:"Folha - Funcionários",category:"Folha de Pagamento",type:"saida",value:4800,date:"2025-02-28",payment:"Transferência",obs:"",competencia:"2025-02",recurrence:"mensal",status:"pago"},
  {id:14,name:"Pacote Mensal Kids",category:"Serviços",type:"entrada",value:2800,date:"2025-02-18",payment:"PIX",obs:"14x R$200",competencia:"2025-02",recurrence:"mensal",status:"pago"},
  {id:15,name:"Instagram Ads",category:"Marketing",type:"saida",value:800,date:"2025-02-20",payment:"Cartão de Crédito",obs:"Campanha Carnaval",competencia:"2025-02",recurrence:"mensal",status:"pago"},
  {id:16,name:"Freelancer DJ",category:"Freelancers",type:"saida",value:700,date:"2025-02-15",payment:"PIX",obs:"Evento carnaval",competencia:"2025-02",recurrence:"none",status:"pago"},
  {id:17,name:"Festas Infantis - 3 eventos",category:"Eventos",type:"entrada",value:5400,date:"2025-03-08",payment:"PIX",obs:"3x R$1800",competencia:"2025-03",recurrence:"none",status:"pago"},
  {id:18,name:"Aluguel do espaço",category:"Aluguel",type:"saida",value:2200,date:"2025-03-10",payment:"Boleto",obs:"",competencia:"2025-03",recurrence:"mensal",status:"pago"},
  {id:19,name:"Pacote Mensal Kids",category:"Serviços",type:"entrada",value:3200,date:"2025-03-18",payment:"PIX",obs:"16 assinaturas",competencia:"2025-03",recurrence:"mensal",status:"pago"},
  {id:20,name:"Energia Elétrica",category:"Energia/Água",type:"saida",value:490,date:"2025-03-15",payment:"Boleto",obs:"",competencia:"2025-03",recurrence:"mensal",status:"pago"},
  {id:21,name:"Folha - Funcionários",category:"Folha de Pagamento",type:"saida",value:4800,date:"2025-03-31",payment:"Transferência",obs:"",competencia:"2025-03",recurrence:"mensal",status:"pago"},
  {id:22,name:"Instagram Ads",category:"Marketing",type:"saida",value:650,date:"2025-03-20",payment:"Cartão de Crédito",obs:"",competencia:"2025-03",recurrence:"mensal",status:"pago"},
  {id:23,name:"Manutenção brinquedos",category:"Manutenção",type:"saida",value:380,date:"2025-03-25",payment:"PIX",obs:"Revisão geral",competencia:"2025-03",recurrence:"none",status:"pago"},
  {id:24,name:"Evento Páscoa Kids",category:"Eventos",type:"entrada",value:6200,date:"2025-04-12",payment:"PIX",obs:"Edição Páscoa",competencia:"2025-04",recurrence:"none",status:"pago"},
  {id:25,name:"Festas Infantis - 2 eventos",category:"Eventos",type:"entrada",value:3600,date:"2025-04-20",payment:"PIX",obs:"",competencia:"2025-04",recurrence:"none",status:"pago"},
  {id:26,name:"Aluguel do espaço",category:"Aluguel",type:"saida",value:2200,date:"2025-04-10",payment:"Boleto",obs:"",competencia:"2025-04",recurrence:"mensal",status:"pago"},
  {id:27,name:"Pacote Mensal Kids",category:"Serviços",type:"entrada",value:3600,date:"2025-04-18",payment:"PIX",obs:"18 assinaturas",competencia:"2025-04",recurrence:"mensal",status:"pago"},
  {id:28,name:"Energia Elétrica",category:"Energia/Água",type:"saida",value:520,date:"2025-04-15",payment:"Boleto",obs:"",competencia:"2025-04",recurrence:"mensal",status:"pago"},
  {id:29,name:"Folha - Funcionários",category:"Folha de Pagamento",type:"saida",value:5200,date:"2025-04-30",payment:"Transferência",obs:"Reajuste",competencia:"2025-04",recurrence:"mensal",status:"pago"},
  {id:30,name:"Instagram + Google Ads",category:"Marketing",type:"saida",value:1200,date:"2025-04-20",payment:"Cartão de Crédito",obs:"Campanha Páscoa",competencia:"2025-04",recurrence:"none",status:"pago"},
  {id:31,name:"Materiais Páscoa",category:"Materiais",type:"saida",value:890,date:"2025-04-08",payment:"PIX",obs:"Ovos, cestas",competencia:"2025-04",recurrence:"none",status:"pago"},
  {id:32,name:"Imposto Simples Nacional",category:"Impostos",type:"saida",value:620,date:"2025-04-20",payment:"Boleto",obs:"",competencia:"2025-04",recurrence:"mensal",status:"pago"},
  {id:33,name:"Festa Infantil - Família Oliveira",category:"Eventos",type:"entrada",value:2200,date:"2025-05-04",payment:"PIX",obs:"",competencia:"2025-05",recurrence:"none",status:"pago"},
  {id:34,name:"Pacote Mensal Kids",category:"Serviços",type:"entrada",value:4000,date:"2025-05-05",payment:"PIX",obs:"20 assinaturas",competencia:"2025-05",recurrence:"mensal",status:"pago"},
  {id:35,name:"Aluguel do espaço",category:"Aluguel",type:"saida",value:2200,date:"2025-05-10",payment:"Boleto",obs:"",competencia:"2025-05",recurrence:"mensal",status:"pendente"},
  {id:36,name:"Energia Elétrica",category:"Energia/Água",type:"saida",value:545,date:"2025-05-15",payment:"Boleto",obs:"",competencia:"2025-05",recurrence:"mensal",status:"pendente"},
  {id:37,name:"Instagram Ads",category:"Marketing",type:"saida",value:700,date:"2025-05-10",payment:"Cartão de Crédito",obs:"",competencia:"2025-05",recurrence:"mensal",status:"pago"},
  {id:38,name:"Festa Infantil - Família Mendes",category:"Eventos",type:"entrada",value:2800,date:"2025-05-11",payment:"PIX",obs:"Pacote VIP",competencia:"2025-05",recurrence:"none",status:"pago"},
  {id:39,name:"Folha - Funcionários",category:"Folha de Pagamento",type:"saida",value:5200,date:"2025-05-31",payment:"Transferência",obs:"",competencia:"2025-05",recurrence:"mensal",status:"pendente"},
  {id:40,name:"Imposto Simples Nacional",category:"Impostos",type:"saida",value:680,date:"2025-05-20",payment:"Boleto",obs:"",competencia:"2025-05",recurrence:"mensal",status:"pendente"},
]

const SEED_SERVICES = [
  {id:1,name:"Festa Infantil Básica",category:"Eventos",unitCost:450,fixedShare:200,taxes:0.06,price:1800,description:"Até 20 crianças, 2h"},
  {id:2,name:"Festa Infantil VIP",category:"Eventos",unitCost:700,fixedShare:250,taxes:0.06,price:2800,description:"Até 30 crianças, 3h + decoração"},
  {id:3,name:"Pacote Mensal Kids",category:"Serviços",unitCost:40,fixedShare:60,taxes:0.06,price:200,description:"Acesso ilimitado semanal"},
  {id:4,name:"Evento Corporativo",category:"Eventos",unitCost:1200,fixedShare:400,taxes:0.06,price:3500,description:"Até 60 pessoas, 4h"},
  {id:5,name:"Workshop Infantil",category:"Serviços",unitCost:180,fixedShare:120,taxes:0.06,price:600,description:"2h, até 15 crianças"},
]

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const fmt    = v => new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL"}).format(v||0)
const fmtPct = v => `${(v||0).toFixed(1)}%`
const fmtNum = v => new Intl.NumberFormat("pt-BR").format(v||0)
const clamp  = (v,a,b) => Math.min(b,Math.max(a,v))
const uid    = () => Date.now() + Math.random()

function useFinancials(txs) {
  return useMemo(() => {
    const byM = {}
    txs.forEach(t => {
      const m = t.competencia || t.date?.slice(0,7) || ""
      if (!byM[m]) byM[m] = { e:0, s:0 }
      if (t.type==="entrada") byM[m].e += t.value
      else byM[m].s += t.value
    })
    const months = Object.keys(byM).sort()
    const monthly = months.map(m => ({
      month:m, label:MONTHS[parseInt(m.slice(5,7))-1]+"/"+m.slice(2,4),
      receita:byM[m].e, custos:byM[m].s, lucro:byM[m].e-byM[m].s
    }))
    const paid = s => t => t.status==="pago" && t.type===s
    const totalE = txs.filter(paid("entrada")).reduce((a,t)=>a+t.value,0)
    const totalS = txs.filter(paid("saida")).reduce((a,t)=>a+t.value,0)
    const fixCats = ["Aluguel","Folha de Pagamento","Energia/Água","Plataformas/Software"]
    const fixedCosts = txs.filter(t=>t.type==="saida"&&t.status==="pago"&&fixCats.includes(t.category)).reduce((a,t)=>a+t.value,0)
    const varCosts = totalS - fixedCosts
    const catMap = {}
    txs.filter(paid("saida")).forEach(t=>{catMap[t.category]=(catMap[t.category]||0)+t.value})
    const catData = Object.entries(catMap).map(([name,value])=>({name,value})).sort((a,b)=>b.value-a.value)
    const cm="2025-05", pm="2025-04"
    const getM = (m,type,status) => txs.filter(t=>(t.competencia||t.date?.slice(0,7))===m&&t.type===type&&(!status||t.status===status)).reduce((a,t)=>a+t.value,0)
    const cmE=getM(cm,"entrada"), cmS=getM(cm,"saida","pago"), cmSP=getM(cm,"saida","pendente")
    const pmE=getM(pm,"entrada"), pmS=getM(pm,"saida")
    const evtTxs = txs.filter(t=>t.type==="entrada"&&t.category==="Eventos"&&t.status==="pago")
    const ticketMedio = evtTxs.length>0 ? evtTxs.reduce((a,t)=>a+t.value,0)/evtTxs.length : 0
    const pending = txs.filter(t=>t.type==="saida"&&t.status==="pendente")
    return { totalE, totalS, lucro:totalE-totalS, margem:totalE>0?((totalE-totalS)/totalE)*100:0,
      monthly, fixedCosts, varCosts, catData, ticketMedio, pending,
      cmE, cmS, cmSP, cmL:cmE-cmS, cmM:cmE>0?((cmE-cmS)/cmE)*100:0,
      pmE, pmS, pmL:pmE-pmS }
  },[txs])
}

// ─── SHARED COMPONENTS ────────────────────────────────────────────────────────
function ChartTip({active,payload,label}) {
  if(!active||!payload?.length) return null
  return (
    <div className="bg-zinc-900 border border-zinc-700 rounded-xl p-3 shadow-2xl text-xs min-w-[140px]">
      <p className="text-zinc-400 mb-2 font-medium">{label}</p>
      {payload.map((p,i)=>(
        <div key={i} className="flex items-center gap-2 mb-1">
          <span className="w-2 h-2 rounded-full flex-shrink-0" style={{background:p.color}}/>
          <span className="text-zinc-400">{p.name}:</span>
          <span className="font-mono font-bold ml-auto" style={{color:p.color}}>{typeof p.value==="number"&&Math.abs(p.value)>1?fmt(p.value):`${p.value?.toFixed(1)}%`}</span>
        </div>
      ))}
    </div>
  )
}

function KPI({label,value,sub,trend,color="emerald",icon:Icon}) {
  const c = CLR[color] || CLR.emerald
  return (
    <div className={`rounded-xl p-4 ${c.card}`}>
      <div className="flex items-center justify-between mb-3">
        <div className={`p-2 rounded-lg bg-zinc-800/60`}>{Icon&&<Icon size={14} className={c.text}/>}</div>
        {trend!=null&&<span className={`text-xs font-semibold flex items-center gap-0.5 ${trend>=0?"text-emerald-400":"text-rose-400"}`}>{trend>=0?<ArrowUpRight size={11}/>:<ArrowDownRight size={11}/>}{Math.abs(trend).toFixed(1)}%</span>}
      </div>
      <p className="text-zinc-500 text-[11px] font-semibold uppercase tracking-wider mb-1">{label}</p>
      <p className={`text-xl font-bold font-mono leading-tight ${c.text}`}>{value}</p>
      {sub&&<p className="text-zinc-600 text-[11px] mt-1">{sub}</p>}
    </div>
  )
}

function Section({title,sub,children}) {
  return (
    <div className="space-y-5">
      <div><h1 className="text-xl font-bold text-zinc-100">{title}</h1>{sub&&<p className="text-zinc-500 text-sm mt-0.5">{sub}</p>}</div>
      {children}
    </div>
  )
}

function Card({children,className=""}) {
  return <div className={`bg-zinc-900/60 border border-zinc-800 rounded-xl p-4 ${className}`}>{children}</div>
}

const INP = "w-full bg-zinc-800 border border-zinc-700 rounded-lg px-3 py-2 text-sm text-zinc-200 focus:border-emerald-500 outline-none transition-colors"

// ─── FORM MODAL ───────────────────────────────────────────────────────────────
function TxForm({data,setData,onSave,onClose,onDelete,title}) {
  return (
    <Card className="border-zinc-700/60">
      <div className="flex items-center justify-between mb-4">
        <p className="text-sm font-semibold text-zinc-300">{title}</p>
        <button onClick={onClose} className="text-zinc-600 hover:text-zinc-300 transition-colors"><X size={15}/></button>
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
        <div className="sm:col-span-2 lg:col-span-2"><label className="text-[11px] text-zinc-500 mb-1 block">Nome *</label><input className={INP} value={data.name||""} onChange={e=>setData(p=>({...p,name:e.target.value}))} placeholder="Descrição"/></div>
        <div><label className="text-[11px] text-zinc-500 mb-1 block">Tipo *</label><select className={INP} value={data.type||"entrada"} onChange={e=>setData(p=>({...p,type:e.target.value,category:""}))}><option value="entrada">Entrada</option><option value="saida">Saída</option></select></div>
        <div><label className="text-[11px] text-zinc-500 mb-1 block">Categoria</label><select className={INP} value={data.category||""} onChange={e=>setData(p=>({...p,category:e.target.value}))}><option value="">Selecione</option>{(CATS[data.type||"entrada"]||[]).map(c=><option key={c}>{c}</option>)}</select></div>
        <div><label className="text-[11px] text-zinc-500 mb-1 block">Valor (R$) *</label><input type="number" className={`${INP} font-mono`} value={data.value||""} onChange={e=>setData(p=>({...p,value:e.target.value}))} placeholder="0.00"/></div>
        <div><label className="text-[11px] text-zinc-500 mb-1 block">Data *</label><input type="date" className={INP} value={data.date||""} onChange={e=>setData(p=>({...p,date:e.target.value}))}/></div>
        <div><label className="text-[11px] text-zinc-500 mb-1 block">Pagamento</label><select className={INP} value={data.payment||"PIX"} onChange={e=>setData(p=>({...p,payment:e.target.value}))}>{PAYMENTS.map(m=><option key={m}>{m}</option>)}</select></div>
        <div><label className="text-[11px] text-zinc-500 mb-1 block">Status</label><select className={INP} value={data.status||"pago"} onChange={e=>setData(p=>({...p,status:e.target.value}))}><option value="pago">Pago</option><option value="pendente">Pendente</option></select></div>
        <div><label className="text-[11px] text-zinc-500 mb-1 block">Recorrência</label><select className={INP} value={data.recurrence||"none"} onChange={e=>setData(p=>({...p,recurrence:e.target.value}))}><option value="none">Avulso</option><option value="mensal">Mensal</option><option value="semanal">Semanal</option><option value="anual">Anual</option></select></div>
        <div className="sm:col-span-2 lg:col-span-3"><label className="text-[11px] text-zinc-500 mb-1 block">Observação</label><input className={INP} value={data.obs||""} onChange={e=>setData(p=>({...p,obs:e.target.value}))} placeholder="Opcional"/></div>
      </div>
      <div className="flex items-center gap-2 mt-4 flex-wrap">
        <button onClick={onSave} className="bg-emerald-500 hover:bg-emerald-400 text-zinc-900 font-semibold text-sm rounded-lg px-5 py-2 transition-colors">Salvar</button>
        <button onClick={onClose} className="bg-zinc-800 hover:bg-zinc-700 text-zinc-300 text-sm rounded-lg px-4 py-2 transition-colors">Cancelar</button>
        {onDelete&&<button onClick={onDelete} className="ml-auto flex items-center gap-1.5 bg-rose-500/10 hover:bg-rose-500/20 text-rose-400 border border-rose-500/20 text-sm rounded-lg px-4 py-2 transition-colors"><Trash2 size={13}/>Excluir</button>}
      </div>
    </Card>
  )
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
function Dashboard({fin,txs}) {
  const te = fin.pmE>0?((fin.cmE/fin.pmE)-1)*100:0
  const tl = fin.pmL!==0?((fin.cmL/fin.pmL)-1)*100:0
  const cashFlow = useMemo(()=>{let a=0;return fin.monthly.map(m=>({...m,saldo:(a+=m.lucro)}));},[fin.monthly])
  const alerts = [
    fin.cmM<15 && {t:"rose",msg:`Margem crítica: ${fmtPct(fin.cmM)} — revisar preços urgente`},
    fin.cmM<25&&fin.cmM>=15 && {t:"amber",msg:`Margem em ${fmtPct(fin.cmM)} — abaixo do ideal de 25%`},
    fin.pending.length>0 && {t:"amber",msg:`${fin.pending.length} contas pendentes: ${fmt(fin.pending.reduce((a,t)=>a+t.value,0))}`},
    fin.cmL<fin.pmL && {t:"rose",msg:`Lucro caiu ${fmt(fin.pmL-fin.cmL)} em relação ao mês anterior`},
    te>5 && {t:"emerald",msg:`Faturamento cresceu ${fmtPct(te)} vs mês anterior — continue!`},
    fin.fixedCosts/(fin.totalS||1)>0.65 && {t:"amber",msg:`Custos fixos em ${fmtPct(fin.fixedCosts/(fin.totalS||1)*100)} das saídas — muito elevado`},
    {t:"sky",msg:"Eventos especiais geram picos de 40%+ — planeje campanhas antecipadas"},
  ].filter(Boolean).slice(0,5)
  const alertCls = {rose:CLR.rose.alert,amber:CLR.amber.alert,emerald:CLR.emerald.alert,sky:CLR.sky.alert}
  return (
    <Section title="Dashboard" sub="Visão estratégica em tempo real">
      <div className="grid grid-cols-2 xl:grid-cols-4 gap-3">
        <KPI label="Faturamento Mês" value={fmt(fin.cmE)} sub={`${te>=0?"+":""}${fmtPct(te)} vs ant.`} trend={te} color="emerald" icon={TrendingUp}/>
        <KPI label="Custos Mês" value={fmt(fin.cmS)} sub={`+${fmt(fin.cmSP)} pendente`} color="rose" icon={TrendingDown}/>
        <KPI label="Lucro Líquido" value={fmt(fin.cmL)} sub={`Margem ${fmtPct(fin.cmM)}`} trend={tl} color={fin.cmL>=0?"indigo":"rose"} icon={CircleDollarSign}/>
        <KPI label="Ticket Médio" value={fmt(fin.ticketMedio)} sub="Por evento" color="amber" icon={Wallet}/>
      </div>
      <div className="grid grid-cols-2 xl:grid-cols-4 gap-3">
        <KPI label="Total Entradas" value={fmt(fin.totalE)} color="emerald" icon={ArrowUpRight}/>
        <KPI label="Total Saídas" value={fmt(fin.totalS)} color="rose" icon={ArrowDownRight}/>
        <KPI label="Resultado Acum." value={fmt(fin.lucro)} sub={`Margem ${fmtPct(fin.margem)}`} color={fin.lucro>=0?"violet":"rose"} icon={PiggyBank}/>
        <KPI label="A Pagar" value={fmt(fin.pending.reduce((a,t)=>a+t.value,0))} sub={`${fin.pending.length} lançamentos`} color="amber" icon={Clock}/>
      </div>
      <div className="grid grid-cols-1 xl:grid-cols-3 gap-4">
        <Card className="xl:col-span-2">
          <p className="text-sm font-semibold text-zinc-300 mb-4">Receita × Custos × Lucro</p>
          <ResponsiveContainer width="100%" height={200}>
            <ComposedChart data={fin.monthly}>
              <CartesianGrid strokeDasharray="3 3" stroke="#27272a"/>
              <XAxis dataKey="label" tick={{fill:"#71717a",fontSize:10}} axisLine={false} tickLine={false}/>
              <YAxis tick={{fill:"#71717a",fontSize:10}} axisLine={false} tickLine={false} tickFormatter={v=>`R$${(v/1000).toFixed(0)}k`} width={36}/>
              <Tooltip content={<ChartTip/>}/><Legend wrapperStyle={{fontSize:10,color:"#a1a1aa"}}/>
              <Bar dataKey="receita" name="Receita" fill={CHART_COLORS.receita} opacity={0.85} radius={[3,3,0,0]}/>
              <Bar dataKey="custos" name="Custos" fill={CHART_COLORS.custos} opacity={0.75} radius={[3,3,0,0]}/>
              <Line dataKey="lucro" name="Lucro" stroke={CHART_COLORS.lucro} strokeWidth={2.5} dot={{r:2.5,fill:CHART_COLORS.lucro}}/>
            </ComposedChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <p className="text-sm font-semibold text-zinc-300 mb-3">Gastos por Categoria</p>
          <ResponsiveContainer width="100%" height={150}>
            <PieChart>
              <Pie data={fin.catData.slice(0,6)} cx="50%" cy="50%" innerRadius={40} outerRadius={65} dataKey="value" paddingAngle={2}>
                {fin.catData.slice(0,6).map((_,i)=><Cell key={i} fill={PIE_COLORS[i]}/>)}
              </Pie>
              <Tooltip formatter={v=>fmt(v)}/>
            </PieChart>
          </ResponsiveContainer>
          <div className="space-y-1 mt-2">
            {fin.catData.slice(0,4).map((c,i)=>(
              <div key={i} className="flex items-center justify-between text-xs">
                <div className="flex items-center gap-1.5"><span className="w-2 h-2 rounded-full" style={{background:PIE_COLORS[i]}}/><span className="text-zinc-500 truncate max-w-[100px]">{c.name}</span></div>
                <span className="text-zinc-300 font-mono">{fmt(c.value)}</span>
              </div>
            ))}
          </div>
        </Card>
      </div>
      <div className="grid grid-cols-1 xl:grid-cols-2 gap-4">
        <Card>
          <p className="text-sm font-semibold text-zinc-300 mb-4">Caixa Acumulado</p>
          <ResponsiveContainer width="100%" height={160}>
            <AreaChart data={cashFlow}>
              <defs><linearGradient id="gCA" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#818cf8" stopOpacity={0.3}/><stop offset="95%" stopColor="#818cf8" stopOpacity={0}/></linearGradient></defs>
              <CartesianGrid strokeDasharray="3 3" stroke="#27272a"/>
              <XAxis dataKey="label" tick={{fill:"#71717a",fontSize:10}} axisLine={false} tickLine={false}/>
              <YAxis tick={{fill:"#71717a",fontSize:10}} axisLine={false} tickLine={false} tickFormatter={v=>`R$${(v/1000).toFixed(0)}k`} width={36}/>
              <Tooltip content={<ChartTip/>}/>
              <Area dataKey="saldo" name="Saldo" stroke="#818cf8" fill="url(#gCA)" strokeWidth={2}/>
            </AreaChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <p className="text-sm font-semibold text-zinc-300 mb-3 flex items-center gap-2"><Sparkles size={14} className="text-amber-400"/>Inteligência</p>
          <div className="space-y-2">
            {alerts.map((a,i)=>(
              <div key={i} className={`flex items-start gap-2 border rounded-lg px-3 py-2 text-xs ${alertCls[a.t]}`}>
                {a.t==="rose"&&<ShieldAlert size={12} className="flex-shrink-0 mt-0.5"/>}
                {a.t==="amber"&&<TriangleAlert size={12} className="flex-shrink-0 mt-0.5"/>}
                {a.t==="emerald"&&<BadgeCheck size={12} className="flex-shrink-0 mt-0.5"/>}
                {a.t==="sky"&&<Lightbulb size={12} className="flex-shrink-0 mt-0.5"/>}
                <span>{a.msg}</span>
              </div>
            ))}
          </div>
        </Card>
      </div>
    </Section>
  )
}

// ─── LANÇAMENTOS ──────────────────────────────────────────────────────────────
function Lancamentos({txs,setTxs}) {
  const blank = {name:"",category:"",type:"entrada",value:"",date:new Date().toISOString().slice(0,10),payment:"PIX",obs:"",competencia:"",recurrence:"none",status:"pago"}
  const [form,setForm] = useState(blank)
  const [edit,setEdit] = useState(null)
  const [show,setShow] = useState(false)
  const [filt,setFilt] = useState("todos")
  const [q,setQ] = useState("")

  const filtered = useMemo(()=>txs.filter(t=>{
    if(filt==="entrada"&&t.type!=="entrada") return false
    if(filt==="saida"&&t.type!=="saida") return false
    if(filt==="pendente"&&t.status!=="pendente") return false
    if(q&&!t.name.toLowerCase().includes(q.toLowerCase())&&!t.category.toLowerCase().includes(q.toLowerCase())) return false
    return true
  }).sort((a,b)=>b.date.localeCompare(a.date)),[txs,filt,q])

  const save = () => {
    if(!form.name||!form.value||!form.date) return
    setTxs(p=>[...p,{...form,id:uid(),value:parseFloat(form.value),competencia:form.competencia||form.date.slice(0,7)}])
    setForm(blank); setShow(false)
  }
  const saveEdit = () => {
    if(!edit) return
    setTxs(p=>p.map(t=>t.id===edit.id?{...edit,value:parseFloat(edit.value)||0}:t))
    setEdit(null)
  }
  const del = id => { setTxs(p=>p.filter(t=>t.id!==id)); setEdit(null) }
  const toggleStatus = (e,id) => { e.stopPropagation(); setTxs(p=>p.map(t=>t.id===id?{...t,status:t.status==="pago"?"pendente":"pago"}:t)) }
  const openEdit = t => { setEdit({...t}); setShow(false) }

  return (
    <Section title="Lançamentos" sub="Entradas e saídas financeiras">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-1 bg-zinc-800/80 border border-zinc-700/40 rounded-lg p-1">
          {["todos","entrada","saida","pendente"].map(f=>(
            <button key={f} onClick={()=>setFilt(f)} className={`text-xs px-3 py-1.5 rounded-md font-medium transition-colors whitespace-nowrap ${filt===f?"bg-zinc-700 text-zinc-100":"text-zinc-500 hover:text-zinc-300"}`}>
              {f==="entrada"?"Entradas":f==="saida"?"Saídas":f==="pendente"?"Pendentes":"Todos"}
            </button>
          ))}
        </div>
        <button onClick={()=>{setEdit(null);setShow(s=>!s)}} className="flex items-center gap-2 bg-emerald-500 hover:bg-emerald-400 text-zinc-900 font-semibold text-sm rounded-lg px-4 py-2 transition-colors flex-shrink-0">
          <Plus size={14}/> Novo
        </button>
      </div>

      {show&&!edit&&<TxForm data={form} setData={setForm} onSave={save} onClose={()=>setShow(false)} title="Novo Lançamento"/>}
      {edit&&<TxForm data={edit} setData={setEdit} onSave={saveEdit} onClose={()=>setEdit(null)} onDelete={()=>del(edit.id)} title="Editar Lançamento"/>}

      <div className="flex items-center gap-2">
        <div className="flex items-center gap-2 bg-zinc-800/50 border border-zinc-700/40 rounded-lg px-3 py-2 flex-1">
          <Search size={13} className="text-zinc-500 flex-shrink-0"/>
          <input className="bg-transparent text-sm text-zinc-300 outline-none w-full placeholder-zinc-600" placeholder="Buscar..." value={q} onChange={e=>setQ(e.target.value)}/>
          {q&&<button onClick={()=>setQ("")}><X size={12} className="text-zinc-600"/></button>}
        </div>
        <span className="text-xs text-zinc-600 flex-shrink-0">{filtered.length}</span>
      </div>

      <div className="rounded-xl border border-zinc-800 overflow-hidden">
        <div className="hidden sm:grid bg-zinc-900/80 px-4 py-2.5 text-[11px] font-semibold text-zinc-500 uppercase tracking-wider border-b border-zinc-800" style={{gridTemplateColumns:"1fr 110px 100px 80px 75px 64px"}}>
          <span>Descrição</span><span>Categoria</span><span className="text-right">Valor</span><span>Data</span><span>Status</span><span/>
        </div>
        <div className="divide-y divide-zinc-800/40 max-h-[500px] overflow-y-auto">
          {!filtered.length&&<p className="text-center text-zinc-600 text-sm py-10">Nenhum lançamento.</p>}
          {filtered.map(t=>(
            <div key={t.id} onClick={()=>openEdit(t)} className={`px-4 py-3 hover:bg-zinc-800/20 cursor-pointer transition-colors ${edit?.id===t.id?"bg-sky-500/5":""}`}>
              {/* Mobile */}
              <div className="sm:hidden flex items-start justify-between gap-2">
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2 mb-0.5"><span className={`w-1.5 h-1.5 rounded-full flex-shrink-0 ${t.type==="entrada"?"bg-emerald-400":"bg-rose-400"}`}/><span className="text-sm text-zinc-200 font-medium truncate">{t.name}</span></div>
                  <div className="flex items-center gap-2 ml-3.5 text-xs text-zinc-600"><span>{t.category}</span><span>·</span><span>{t.date?.slice(0,10)}</span></div>
                </div>
                <span className={`font-mono font-bold text-sm flex-shrink-0 ${t.type==="entrada"?"text-emerald-400":"text-rose-400"}`}>{t.type==="entrada"?"+":"-"}{fmt(t.value)}</span>
              </div>
              {/* Desktop */}
              <div className="hidden sm:grid items-center gap-3" style={{gridTemplateColumns:"1fr 110px 100px 80px 75px 64px"}}>
                <div className="flex items-center gap-2 min-w-0">
                  <span className={`w-1.5 h-1.5 rounded-full flex-shrink-0 ${t.type==="entrada"?"bg-emerald-400":"bg-rose-400"}`}/>
                  <span className="text-sm text-zinc-200 font-medium truncate">{t.name}</span>
                  {t.recurrence!=="none"&&<span className="text-[10px] text-indigo-400 bg-indigo-400/10 rounded px-1">↻</span>}
                </div>
                <span className="text-zinc-500 text-xs truncate">{t.category}</span>
                <span className={`font-mono font-bold text-sm text-right ${t.type==="entrada"?"text-emerald-400":"text-rose-400"}`}>{t.type==="entrada"?"+":"-"}{fmt(t.value)}</span>
                <span className="text-zinc-500 text-xs">{t.date?.slice(0,10)}</span>
                <button onClick={e=>toggleStatus(e,t.id)} className={`text-[11px] rounded-full px-2 py-0.5 font-medium border w-fit ${t.status==="pago"?"bg-emerald-400/10 text-emerald-400 border-emerald-400/20":"bg-amber-400/10 text-amber-400 border-amber-400/20"}`}>{t.status==="pago"?"Pago":"Pend."}</button>
                <button onClick={e=>{e.stopPropagation();del(t.id)}} className="p-1.5 text-zinc-700 hover:text-rose-400 transition-colors ml-auto"><Trash2 size={13}/></button>
              </div>
            </div>
          ))}
        </div>
      </div>
    </Section>
  )
}

// ─── FLUXO DE CAIXA ───────────────────────────────────────────────────────────
function FluxoCaixa({fin}) {
  const [period,setPeriod] = useState("mensal")
  const proj = useMemo(()=>{
    const n = fin.monthly.length||1
    const avgR = fin.monthly.reduce((a,m)=>a+m.receita,0)/n
    const avgS = fin.monthly.reduce((a,m)=>a+m.custos,0)/n
    const future = ["Jun/25","Jul/25","Ago/25","Set/25"].map((label,i)=>({label,receita:avgR*(1+0.05*(i+1)),custos:avgS*(1+0.02*(i+1)),lucro:avgR*(1+0.05*(i+1))-avgS*(1+0.02*(i+1))}))
    return [...fin.monthly,...future]
  },[fin.monthly])
  const accData = fin.monthly.reduce((acc,m,i)=>[...acc,{...m,saldo:(acc[i-1]?.saldo||0)+m.lucro}],[])

  return (
    <Section title="Fluxo de Caixa" sub="Movimentação e projeção financeira">
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div className="grid grid-cols-3 gap-3 flex-1">
          {[{l:"Entradas",v:fin.cmE,c:"emerald"},{l:"Saídas",v:fin.cmS+fin.cmSP,sub:fmt(fin.cmSP)+" pend.",c:"rose"},{l:"Resultado",v:fin.cmL,sub:fmtPct(fin.cmM)+" margem",c:fin.cmL>=0?"indigo":"rose"}].map((k,i)=>(
            <div key={i} className={`border rounded-xl p-3 ${CLR[k.c].card}`}>
              <p className="text-xs text-zinc-500 mb-0.5">{k.l}</p>
              <p className={`text-lg font-bold font-mono ${CLR[k.c].text}`}>{fmt(k.v)}</p>
              {k.sub&&<p className="text-xs text-zinc-600">{k.sub}</p>}
            </div>
          ))}
        </div>
        <div className="flex items-center gap-1 bg-zinc-800/80 border border-zinc-700/40 rounded-lg p-1">
          {["mensal","acumulado","projeção"].map(p=><button key={p} onClick={()=>setPeriod(p)} className={`text-xs px-3 py-1.5 rounded-md font-medium capitalize whitespace-nowrap transition-colors ${period===p?"bg-zinc-700 text-zinc-100":"text-zinc-500 hover:text-zinc-300"}`}>{p.charAt(0).toUpperCase()+p.slice(1)}</button>)}
        </div>
      </div>
      <Card>
        {period==="mensal"&&<>
          <p className="text-sm font-semibold text-zinc-300 mb-4">Entradas × Saídas Mensais</p>
          <ResponsiveContainer width="100%" height={230}>
            <BarChart data={fin.monthly}><CartesianGrid strokeDasharray="3 3" stroke="#27272a"/><XAxis dataKey="label" tick={{fill:"#71717a",fontSize:10}} axisLine={false} tickLine={false}/><YAxis tick={{fill:"#71717a",fontSize:10}} axisLine={false} tickLine={false} tickFormatter={v=>`R$${(v/1000).toFixed(0)}k`} width={36}/><Tooltip content={<ChartTip/>}/><Legend wrapperStyle={{fontSize:10,color:"#a1a1aa"}}/><Bar dataKey="receita" name="Entradas" fill={CHART_COLORS.receita} radius={[4,4,0,0]}/><Bar dataKey="custos" name="Saídas" fill={CHART_COLORS.custos} opacity={0.8} radius={[4,4,0,0]}/></BarChart>
          </ResponsiveContainer>
        </>}
        {period==="acumulado"&&<>
          <p className="text-sm font-semibold text-zinc-300 mb-4">Saldo Acumulado</p>
          <ResponsiveContainer width="100%" height={230}>
            <AreaChart data={accData}>
              <defs><linearGradient id="gAcc2" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#818cf8" stopOpacity={0.4}/><stop offset="95%" stopColor="#818cf8" stopOpacity={0}/></linearGradient></defs>
              <CartesianGrid strokeDasharray="3 3" stroke="#27272a"/><XAxis dataKey="label" tick={{fill:"#71717a",fontSize:10}} axisLine={false} tickLine={false}/><YAxis tick={{fill:"#71717a",fontSize:10}} axisLine={false} tickLine={false} tickFormatter={v=>`R$${(v/1000).toFixed(0)}k`} width={36}/><Tooltip content={<ChartTip/>}/>
              <ReferenceLine y={0} stroke="#ef4444" strokeDasharray="4"/>
              <Area dataKey="saldo" name="Saldo" stroke="#818cf8" fill="url(#gAcc2)" strokeWidth={2.5}/>
            </AreaChart>
          </ResponsiveContainer>
        </>}
        {period==="projeção"&&<>
          <div className="flex items-center gap-2 mb-4"><p className="text-sm font-semibold text-zinc-300">Projeção 4 meses</p><span className="text-xs bg-violet-500/10 text-violet-400 border border-violet-500/20 rounded-full px-2 py-0.5">IA Preditiva</span></div>
          <ResponsiveContainer width="100%" height={230}>
            <ComposedChart data={proj}><CartesianGrid strokeDasharray="3 3" stroke="#27272a"/><XAxis dataKey="label" tick={{fill:"#71717a",fontSize:10}} axisLine={false} tickLine={false}/><YAxis tick={{fill:"#71717a",fontSize:10}} axisLine={false} tickLine={false} tickFormatter={v=>`R$${(v/1000).toFixed(0)}k`} width={36}/><Tooltip content={<ChartTip/>}/><Legend wrapperStyle={{fontSize:10,color:"#a1a1aa"}}/><Bar dataKey="receita" name="Receita" fill={CHART_COLORS.receita} opacity={0.7} radius={[4,4,0,0]}/><Bar dataKey="custos" name="Custos" fill={CHART_COLORS.custos} opacity={0.6} radius={[4,4,0,0]}/><Line dataKey="lucro" name="Lucro" stroke="#a78bfa" strokeWidth={2} strokeDasharray="5 5" dot={{r:2}}/></ComposedChart>
          </ResponsiveContainer>
          <p className="text-xs text-zinc-600 mt-2">* Crescimento estimado de 5%/mês baseado na média histórica</p>
        </>}
      </Card>
      {fin.pending.length>0&&(
        <Card>
          <p className="text-sm font-semibold text-zinc-300 mb-3 flex items-center gap-2"><Clock size={13} className="text-amber-400"/>Contas a Pagar ({fin.pending.length})</p>
          <div className="space-y-2">
            {fin.pending.map(t=>(
              <div key={t.id} className="flex items-center justify-between bg-amber-500/5 border border-amber-500/15 rounded-lg px-3 py-2.5">
                <div><p className="text-sm text-zinc-200 font-medium">{t.name}</p><p className="text-xs text-zinc-500">{t.category} · {t.date}</p></div>
                <span className="font-mono font-bold text-amber-400">{fmt(t.value)}</span>
              </div>
            ))}
          </div>
        </Card>
      )}
    </Section>
  )
}

// ─── PRECIFICAÇÃO ─────────────────────────────────────────────────────────────
function Precificacao({services,setServices}) {
  const [form,setForm] = useState({unitCost:"",fixedShare:"",taxes:"6",commission:"0",desiredMargin:"30"})
  const [sim,setSim] = useState("")
  const [editSvc,setEditSvc] = useState(null)
  const [showAdd,setShowAdd] = useState(false)
  const blankSvc = {name:"",category:"Serviços",unitCost:"",fixedShare:"",taxes:"6",price:"",description:""}
  const [svcForm,setSvcForm] = useState(blankSvc)

  const calc = useMemo(()=>{
    const uc=parseFloat(form.unitCost)||0,fc=parseFloat(form.fixedShare)||0
    const tx=(parseFloat(form.taxes)||0)/100,cm=(parseFloat(form.commission)||0)/100,dm=(parseFloat(form.desiredMargin)||0)/100
    const total=uc+fc,denom=1-dm-tx-cm
    const ideal=denom>0?total/denom:0,min=total/(1-tx-cm||1)
    const np=ideal-ideal*tx-ideal*cm-total
    const sp=parseFloat(sim)||0,spr=sp>0?sp-sp*tx-sp*cm-total:0
    return{total,ideal,min,markup:denom>0?1/denom:0,netProfit:np,netMargin:ideal>0?(np/ideal)*100:0,simProfit:spr,simMargin:sp>0?(spr/sp)*100:0}
  },[form,sim])

  const svcM = useMemo(()=>services.map(s=>{
    const t=s.unitCost+s.fixedShare,np=s.price-s.price*s.taxes-t,m=s.price>0?(np/s.price)*100:0
    const status=m<0?"loss":m<20?"low":m<35?"ok":"good"
    return {...s,totalCost:t,netProfit:np,margin:m,status}
  }),[services])

  const statusCls = {loss:"text-rose-400 bg-rose-500/10 border-rose-500/20",low:"text-amber-400 bg-amber-500/10 border-amber-500/20",ok:"text-sky-400 bg-sky-500/10 border-sky-500/20",good:"text-emerald-400 bg-emerald-500/10 border-emerald-500/20"}
  const statusLbl = {loss:"Prejuízo",low:"Baixa",ok:"Regular",good:"Saudável"}
  const barCls = {loss:"bg-rose-400",low:"bg-amber-400",ok:"bg-sky-400",good:"bg-emerald-400"}

  const addSvc = () => {
    if(!svcForm.name||!svcForm.price) return
    setServices(p=>[...p,{...svcForm,id:uid(),unitCost:parseFloat(svcForm.unitCost)||0,fixedShare:parseFloat(svcForm.fixedShare)||0,taxes:(parseFloat(svcForm.taxes)||6)/100,price:parseFloat(svcForm.price)||0}])
    setSvcForm(blankSvc); setShowAdd(false)
  }
  const saveSvc = () => {
    if(!editSvc) return
    setServices(p=>p.map(s=>s.id===editSvc.id?{...editSvc,unitCost:parseFloat(editSvc.unitCost)||0,fixedShare:parseFloat(editSvc.fixedShare)||0,taxes:(parseFloat(editSvc.taxes)||6)/100,price:parseFloat(editSvc.price)||0}:s))
    setEditSvc(null)
  }
  const delSvc = id => { setServices(p=>p.filter(s=>s.id!==id)); setEditSvc(null) }

  const SvcFieldForm = ({data,setData}) => (
    <div className="grid grid-cols-2 gap-2">
      {[{k:"name",l:"Nome"},{k:"price",l:"Preço (R$)"},{k:"unitCost",l:"Custo Direto"},{k:"fixedShare",l:"Fixo Rateado"},{k:"taxes",l:"Impostos (%)"},{k:"description",l:"Descrição"}].map(f=>(
        <div key={f.k}><label className="text-[10px] text-zinc-500 block mb-0.5">{f.l}</label><input className="w-full bg-zinc-700 border border-zinc-600 rounded px-2 py-1.5 text-xs text-zinc-200 outline-none focus:border-violet-400" value={data[f.k]||""} onChange={e=>setData(p=>({...p,[f.k]:e.target.value}))}/></div>
      ))}
    </div>
  )

  return (
    <Section title="Precificação Inteligente" sub="Calcule o preço correto com margem real">
      <div className="grid grid-cols-1 xl:grid-cols-2 gap-5">
        <Card>
          <p className="text-sm font-semibold text-zinc-300 mb-4 flex items-center gap-2"><Calculator size={14} className="text-emerald-400"/>Calculadora de Preço</p>
          <div className="space-y-3">
            {[{k:"unitCost",l:"Custo Direto / CMV (R$)"},{k:"fixedShare",l:"Rateio Custo Fixo (R$)"},{k:"taxes",l:"Impostos (%)"},{k:"commission",l:"Comissão (%)"},{k:"desiredMargin",l:"Margem Desejada (%)"}].map(f=>(
              <div key={f.k} className="flex items-center gap-3">
                <label className="text-xs text-zinc-500 w-44 flex-shrink-0">{f.l}</label>
                <input type="number" className={`flex-1 ${INP} font-mono`} value={form[f.k]} onChange={e=>setForm(p=>({...p,[f.k]:e.target.value}))}/>
              </div>
            ))}
          </div>
          <div className="mt-4 bg-zinc-800/50 rounded-xl p-4 space-y-2">
            <div className="flex justify-between text-sm"><span className="text-zinc-500">Custo Total</span><span className="font-mono text-zinc-300">{fmt(calc.total)}</span></div>
            <div className="flex justify-between text-sm"><span className="text-zinc-500">Preço Mínimo</span><span className="font-mono text-amber-400">{fmt(calc.min)}</span></div>
            <div className="h-px bg-zinc-700"/>
            <div className="flex justify-between font-bold"><span className="text-zinc-300">Preço Ideal</span><span className="font-mono text-emerald-400 text-xl">{fmt(calc.ideal)}</span></div>
            <div className="flex justify-between text-sm"><span className="text-zinc-500">Markup</span><span className="font-mono text-zinc-300">{calc.markup.toFixed(2)}x</span></div>
            <div className="flex justify-between text-sm"><span className="text-zinc-500">Lucro por venda</span><span className="font-mono text-indigo-400">{fmt(calc.netProfit)}</span></div>
            <div className="flex justify-between text-sm"><span className="text-zinc-500">Margem líquida</span><span className={`font-mono font-bold ${calc.netMargin>=20?"text-emerald-400":calc.netMargin>=0?"text-amber-400":"text-rose-400"}`}>{fmtPct(calc.netMargin)}</span></div>
          </div>
          <div className="mt-3 bg-zinc-800/30 border border-zinc-700/40 rounded-xl p-3">
            <p className="text-xs text-zinc-500 mb-2">🧮 Se eu vender por R$...</p>
            <div className="flex items-center gap-2">
              <input type="number" className={`flex-1 ${INP} font-mono text-sm`} value={sim} onChange={e=>setSim(e.target.value)} placeholder="Digite o preço"/>
              <div className="text-right flex-shrink-0">
                <p className={`font-mono font-bold ${calc.simProfit>=0?"text-emerald-400":"text-rose-400"}`}>{fmt(calc.simProfit)}</p>
                <p className={`text-xs ${calc.simMargin>=20?"text-emerald-500":calc.simMargin>=0?"text-amber-500":"text-rose-500"}`}>{fmtPct(calc.simMargin)}</p>
              </div>
            </div>
          </div>
        </Card>
        <Card>
          <div className="flex items-center justify-between mb-4">
            <p className="text-sm font-semibold text-zinc-300 flex items-center gap-2"><Package size={14} className="text-violet-400"/>Rentabilidade</p>
            <button onClick={()=>{setEditSvc(null);setShowAdd(s=>!s)}} className="text-xs bg-violet-500/10 text-violet-400 border border-violet-500/20 rounded-lg px-3 py-1.5 hover:bg-violet-500/20 transition-colors flex items-center gap-1"><Plus size={11}/>Novo</button>
          </div>
          {(showAdd||editSvc)&&(
            <div className="bg-zinc-800/50 border border-zinc-700/40 rounded-xl p-3 mb-3">
              <p className="text-xs font-semibold text-zinc-400 mb-2">{editSvc?"Editar":"Novo Serviço"}</p>
              <SvcFieldForm data={editSvc||svcForm} setData={editSvc?setEditSvc:setSvcForm}/>
              <div className="flex gap-2 mt-2">
                <button onClick={editSvc?saveSvc:addSvc} className="text-xs bg-violet-500 text-white rounded-lg px-3 py-1.5">Salvar</button>
                <button onClick={()=>{setEditSvc(null);setShowAdd(false)}} className="text-xs bg-zinc-700 text-zinc-300 rounded-lg px-3 py-1.5">Cancelar</button>
                {editSvc&&<button onClick={()=>delSvc(editSvc.id)} className="text-xs bg-rose-500/10 text-rose-400 border border-rose-500/20 rounded-lg px-3 py-1.5 ml-auto flex items-center gap-1"><Trash2 size={10}/>Excluir</button>}
              </div>
            </div>
          )}
          <div className="space-y-2.5 max-h-[440px] overflow-y-auto">
            {svcM.sort((a,b)=>b.margin-a.margin).map(s=>(
              <div key={s.id} onClick={()=>{setEditSvc({...s,taxes:String(Math.round(s.taxes*100)),fixedShare:String(s.fixedShare)});setShowAdd(false)}} className={`bg-zinc-800/30 border rounded-xl p-3 cursor-pointer hover:border-zinc-600 transition-colors ${editSvc?.id===s.id?"border-violet-500/50":"border-zinc-700/40"}`}>
                <div className="flex items-start justify-between mb-2">
                  <div><p className="text-sm font-semibold text-zinc-200">{s.name}</p><p className="text-xs text-zinc-600">{s.description}</p></div>
                  <div className="flex items-center gap-1.5">
                    <span className={`text-xs px-2 py-0.5 rounded-full border font-medium ${statusCls[s.status]}`}>{statusLbl[s.status]}</span>
                    <button onClick={e=>{e.stopPropagation();delSvc(s.id)}} className="p-1 text-zinc-700 hover:text-rose-400 transition-colors"><Trash2 size={11}/></button>
                  </div>
                </div>
                <div className="grid grid-cols-4 gap-1 text-center mb-2">
                  {[{l:"Preço",v:fmt(s.price)},{l:"Custo",v:fmt(s.totalCost)},{l:"Lucro",v:fmt(s.netProfit),e:s.netProfit>=0},{l:"Margem",v:fmtPct(s.margin),cls:statusCls[s.status].split(" ")[0]}].map((c,i)=>(
                    <div key={i}><p className="text-[10px] text-zinc-600">{c.l}</p><p className={`font-mono text-xs font-bold ${c.cls||(c.e!=null?(c.e?"text-emerald-400":"text-rose-400"):"text-zinc-300")}`}>{c.v}</p></div>
                  ))}
                </div>
                <div className="h-1 bg-zinc-700 rounded-full overflow-hidden">
                  <div className={`h-full rounded-full ${barCls[s.status]}`} style={{width:`${clamp(s.margin,0,100)}%`}}/>
                </div>
              </div>
            ))}
          </div>
        </Card>
      </div>
    </Section>
  )
}

// ─── PONTO DE EQUILÍBRIO ──────────────────────────────────────────────────────
function PontoEquilibrio({fin, txs}) {
  // All values computed from real data — override only if user explicitly edits
  const [cfOverride, setCFOverride] = useState("")
  const [atOverride, setATOverride] = useState("")
  const [amOverride, setAMOverride] = useState("")

  // Auto-calculated from transactions
  const fixCats = ["Aluguel","Folha de Pagamento","Energia/Água","Plataformas/Software","Impostos"]
  const varCats  = ["Materiais","Freelancers","Marketing","Manutenção","Equipamentos"]

  // Average monthly fixed costs (last 3 months or all months)
  const monthlyFixed = useMemo(() => {
    const byMonth = {}
    txs.filter(t => t.type==="saida" && t.status==="pago" && fixCats.includes(t.category)).forEach(t => {
      const m = t.competencia || t.date?.slice(0,7) || ""
      byMonth[m] = (byMonth[m]||0) + t.value
    })
    const vals = Object.values(byMonth)
    if (!vals.length) return 0
    return vals.reduce((a,v)=>a+v,0) / vals.length
  }, [txs])

  // Average monthly variable costs
  const monthlyVar = useMemo(() => {
    const byMonth = {}
    txs.filter(t => t.type==="saida" && t.status==="pago" && varCats.includes(t.category)).forEach(t => {
      const m = t.competencia || t.date?.slice(0,7) || ""
      byMonth[m] = (byMonth[m]||0) + t.value
    })
    const vals = Object.values(byMonth)
    if (!vals.length) return 0
    return vals.reduce((a,v)=>a+v,0) / vals.length
  }, [txs])

  // Average ticket from all entrada transactions
  const autoTicket = useMemo(() => {
    const entries = txs.filter(t => t.type==="entrada" && t.status==="pago")
    if (!entries.length) return 0
    return entries.reduce((a,t)=>a+t.value,0) / entries.length
  }, [txs])

  // Average monthly revenue
  const avgMonthlyRevenue = useMemo(() => {
    const byMonth = {}
    txs.filter(t => t.type==="entrada" && t.status==="pago").forEach(t => {
      const m = t.competencia || t.date?.slice(0,7) || ""
      byMonth[m] = (byMonth[m]||0) + t.value
    })
    const vals = Object.values(byMonth)
    if (!vals.length) return 0
    return vals.reduce((a,v)=>a+v,0) / vals.length
  }, [txs])

  // Auto contribution margin = (receita - custos variáveis) / receita
  const autoMargin = useMemo(() => {
    if (!avgMonthlyRevenue) return 35
    const cm = ((avgMonthlyRevenue - monthlyVar) / avgMonthlyRevenue) * 100
    return Math.max(5, Math.min(95, cm))
  }, [avgMonthlyRevenue, monthlyVar])

  // Use override or auto value
  const fc     = parseFloat(cfOverride) || monthlyFixed || 8000
  const ticket = parseFloat(atOverride) || autoTicket  || 1800
  const margin = (parseFloat(amOverride) || autoMargin) / 100

  const be  = margin > 0 ? fc / margin : 0
  const beU = be > 0 ? Math.ceil(be / ticket) : 0
  const cur = fin.cmE
  const pct = clamp((cur / (be||1)) * 100, 0, 100)
  const rem = Math.max(0, be - cur)
  const daily = be / 26
  const est   = rem > 0 ? Math.ceil(rem / ((cur/15)||1)) : 0

  const chartData = useMemo(() => {
    const pts = []
    const steps = Math.max(1, Math.floor(beU/8))
    for (let q = 0; q <= beU * 2; q += steps) {
      pts.push({ q, receita: q*ticket, custos: fc + q*ticket*(1-margin) })
    }
    return pts
  }, [beU, ticket, fc, margin])

  const Row = ({label, auto, override, setOverride, prefix="R$", suffix=""}) => (
    <div>
      <label className="text-xs text-zinc-500 block mb-1">{label}</label>
      <div className="relative">
        <input
          type="number"
          className={`${INP} font-mono ${!override ? "text-emerald-400" : "text-zinc-200"}`}
          value={override || ""}
          onChange={e => setOverride(e.target.value)}
          placeholder={`${auto.toFixed(0)} (calculado)`}
        />
        {override && (
          <button onClick={() => setOverride("")}
            className="absolute right-2 top-1/2 -translate-y-1/2 text-zinc-600 hover:text-zinc-300 text-xs">
            ↺ auto
          </button>
        )}
      </div>
      <p className="text-[11px] text-zinc-600 mt-0.5">
        Calculado dos dados: {prefix}{auto.toFixed(0)}{suffix}
      </p>
    </div>
  )

  return (
    <Section title="Ponto de Equilíbrio" sub="Calculado automaticamente com base em todos os lançamentos">
      <div className="grid grid-cols-1 xl:grid-cols-3 gap-4">
        <Card>
          <p className="text-sm font-semibold text-zinc-300 mb-1">Parâmetros</p>
          <p className="text-xs text-zinc-600 mb-4">Valores em verde = calculados automaticamente. Edite para sobrescrever.</p>
          <div className="space-y-4">
            <Row label="Custos Fixos Médios/Mês (R$)" auto={monthlyFixed} override={cfOverride} setOverride={setCFOverride}/>
            <Row label="Ticket Médio (R$)" auto={autoTicket} override={atOverride} setOverride={setATOverride}/>
            <Row label="Margem de Contribuição (%)" auto={autoMargin} override={amOverride} setOverride={setAMOverride} prefix="" suffix="%"/>
          </div>
          <div className="mt-4 bg-zinc-800/30 rounded-xl p-3 space-y-1 text-xs">
            <p className="text-zinc-500 font-semibold mb-2">Baseado em {txs.filter(t=>t.status==="pago").length} lançamentos</p>
            <div className="flex justify-between"><span className="text-zinc-500">Receita média/mês</span><span className="font-mono text-emerald-400">{fmt(avgMonthlyRevenue)}</span></div>
            <div className="flex justify-between"><span className="text-zinc-500">Custos fixos médios</span><span className="font-mono text-rose-400">{fmt(monthlyFixed)}</span></div>
            <div className="flex justify-between"><span className="text-zinc-500">Custos variáveis médios</span><span className="font-mono text-amber-400">{fmt(monthlyVar)}</span></div>
          </div>
        </Card>

        <div className="xl:col-span-2 grid grid-cols-2 gap-3">
          {[
            {l:"Break-Even Mensal", v:fmt(be),      c:CLR.indigo,   sub:"Meta mínima para cobrir custos"},
            {l:"Vendas Mínimas",    v:fmtNum(beU),  c:CLR.amber,    sub:`Baseado em ticket de ${fmt(ticket)}`},
            {l:"Meta Diária",       v:fmt(daily),   c:CLR.emerald,  sub:"26 dias úteis por mês"},
            {l:rem<=0?"🎉 Meta Atingida":"Falta este Mês", v:rem<=0?fmt(cur-be):fmt(rem), c:rem<=0?CLR.emerald:CLR.rose, sub:rem<=0?"Acima do break-even":`~${est} dias para zerar`}
          ].map((k,i)=>(
            <div key={i} className={`border rounded-xl p-3 text-center ${k.c.card}`}>
              <p className="text-[11px] text-zinc-500 mb-0.5 leading-tight">{k.l}</p>
              <p className={`text-lg font-bold font-mono ${k.c.text}`}>{k.v}</p>
              <p className="text-[10px] text-zinc-600 mt-0.5">{k.sub}</p>
            </div>
          ))}
        </div>
      </div>

      <Card>
        <div className="flex justify-between items-center mb-3">
          <p className="text-sm font-semibold text-zinc-300">Progresso do Mês Atual</p>
          <span className="text-xs font-mono text-zinc-400">{fmt(cur)} / {fmt(be)}</span>
        </div>
        <div className="h-3 bg-zinc-800 rounded-full overflow-hidden mb-2">
          <div className={`h-full rounded-full transition-all duration-700 ${pct>=100?"bg-emerald-400":pct>=70?"bg-sky-400":pct>=40?"bg-amber-400":"bg-rose-400"}`} style={{width:`${pct}%`}}/>
        </div>
        <div className="flex justify-between text-xs text-zinc-600">
          <span>{pct.toFixed(1)}% do break-even atingido</span>
          <span>{pct>=100?"✅ Meta alcançada!":`~${est} dias estimados`}</span>
        </div>
      </Card>

      <Card>
        <p className="text-sm font-semibold text-zinc-300 mb-4">Curva de Equilíbrio</p>
        <ResponsiveContainer width="100%" height={210}>
          <LineChart data={chartData}>
            <CartesianGrid strokeDasharray="3 3" stroke="#27272a"/>
            <XAxis dataKey="q" tick={{fill:"#71717a",fontSize:10}} axisLine={false} tickLine={false} label={{value:"Vendas",fill:"#71717a",position:"insideBottomRight",fontSize:9,offset:-5}}/>
            <YAxis tick={{fill:"#71717a",fontSize:10}} axisLine={false} tickLine={false} tickFormatter={v=>`R$${(v/1000).toFixed(0)}k`} width={36}/>
            <Tooltip content={<ChartTip/>}/><Legend wrapperStyle={{fontSize:10,color:"#a1a1aa"}}/>
            <ReferenceLine x={beU} stroke="#a78bfa" strokeDasharray="5 5" label={{value:`PE: ${beU}`,fill:"#a78bfa",fontSize:10}}/>
            <Line dataKey="receita" name="Receita" stroke="#10b981" strokeWidth={2} dot={false}/>
            <Line dataKey="custos" name="Custos Totais" stroke="#ef4444" strokeWidth={2} dot={false}/>
          </LineChart>
        </ResponsiveContainer>
      </Card>
    </Section>
  )
}

// ─── DRE ──────────────────────────────────────────────────────────────────────
function DRE({txs}) {
  const periods = [...new Set(txs.map(t=>t.competencia||t.date?.slice(0,7)))].filter(Boolean).sort()
  const [period,setPeriod] = useState("2025-05")
  const d = useMemo(()=>{
    const f = t=>(t.competencia||t.date?.slice(0,7))===period
    const sum = (arr,cond) => arr.filter(t=>f(t)&&cond(t)).reduce((a,t)=>a+t.value,0)
    const receita = sum(txs,t=>t.type==="entrada")
    const cmv = sum(txs,t=>t.type==="saida"&&["Materiais","Freelancers"].includes(t.category))
    const lb=receita-cmv, mb=receita>0?(lb/receita)*100:0
    const pessoal = sum(txs,t=>t.type==="saida"&&t.category==="Folha de Pagamento")
    const aluguel = sum(txs,t=>t.type==="saida"&&["Aluguel","Energia/Água"].includes(t.category))
    const despOp = sum(txs,t=>t.type==="saida"&&["Marketing","Manutenção","Plataformas/Software"].includes(t.category))
    const ebitda=lb-despOp-pessoal-aluguel
    const impostos = sum(txs,t=>t.type==="saida"&&t.category==="Impostos")
    const ll=ebitda-impostos, ml=receita>0?(ll/receita)*100:0
    return{receita,cmv,lb,mb,pessoal,aluguel,despOp,ebitda,impostos,ll,ml}
  },[txs,period])

  const Row = ({label,value,indent=0,bold=false,color="neutral",sep=false,sub}) => {
    const tc = color==="pos"?"text-emerald-400":color==="neg"?"text-rose-400":color==="acc"?"text-indigo-400":"text-zinc-300"
    return (<>
      {sep&&<div className="h-px bg-zinc-800 mx-3 my-0.5"/>}
      <div className={`flex justify-between items-center py-2 rounded-lg ${bold?"bg-zinc-800/40":""}`} style={{paddingLeft:`${12+indent*16}px`,paddingRight:"12px"}}>
        <div><span className={`text-sm ${bold?"font-bold text-zinc-100":"text-zinc-400"}`}>{label}</span>{sub&&<span className="text-xs text-zinc-600 ml-2">{sub}</span>}</div>
        {value!=null&&<span className={`font-mono text-sm ${bold?"font-bold":""} ${tc}`}>{fmt(value)}</span>}
      </div>
    </>)
  }

  return (
    <Section title="DRE — Demonstrativo de Resultado" sub="Visão gerencial do resultado financeiro">
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div className="grid grid-cols-2 xl:grid-cols-4 gap-3 flex-1">
          {[{l:"Receita Bruta",v:d.receita,c:CLR.emerald},{l:"Lucro Bruto",v:d.lb,c:CLR.sky,sub:fmtPct(d.mb)},{l:"EBITDA",v:d.ebitda,c:d.ebitda>=0?CLR.indigo:CLR.rose},{l:"Lucro Líquido",v:d.ll,c:d.ll>=0?CLR.violet:CLR.rose,sub:fmtPct(d.ml)}].map((k,i)=>(
            <div key={i} className={`border rounded-xl p-3 text-center ${k.c.card}`}>
              <p className="text-[11px] text-zinc-500">{k.l}</p>
              <p className={`font-mono font-bold text-base ${k.c.text}`}>{fmt(k.v)}</p>
              {k.sub&&<p className="text-xs text-zinc-600">{k.sub}</p>}
            </div>
          ))}
        </div>
        <select className="bg-zinc-800 border border-zinc-700 rounded-lg px-3 py-2 text-sm text-zinc-200 outline-none flex-shrink-0" value={period} onChange={e=>setPeriod(e.target.value)}>
          {periods.map(p=><option key={p} value={p}>{MONTHS[parseInt(p.slice(5,7))-1]}/{p.slice(2,4)}</option>)}
        </select>
      </div>
      <Card className="p-2">
        <div className="bg-zinc-800/80 px-4 py-3 rounded-lg flex justify-between mb-1">
          <span className="text-xs font-bold text-zinc-400 uppercase tracking-wider">DRE — {MONTHS[parseInt(period.slice(5,7))-1]}/{period.slice(2,4)}</span>
          <span className="text-xs font-bold text-zinc-400 uppercase tracking-wider">Valor</span>
        </div>
        <Row label="(+) RECEITA OPERACIONAL BRUTA" value={d.receita} bold color="pos"/>
        <Row label="(-) Custo do Serviço / CMV" value={-d.cmv} indent={1} color="neg"/>
        <Row label="(=) LUCRO BRUTO" value={d.lb} bold color={d.lb>=0?"pos":"neg"} sep sub={`Margem ${fmtPct(d.mb)}`}/>
        <Row label="(-) Despesas Operacionais" value={-(d.despOp+d.pessoal+d.aluguel)} color="neg" sep/>
        <Row label="Folha de Pagamento" value={-d.pessoal} indent={1}/>
        <Row label="Aluguel e Utilidades" value={-d.aluguel} indent={1}/>
        <Row label="Marketing e Operacional" value={-d.despOp} indent={1}/>
        <Row label="(=) EBITDA" value={d.ebitda} bold color={d.ebitda>=0?"pos":"neg"} sep/>
        <Row label="(-) Impostos e Tributos" value={-d.impostos} color="neg" sep/>
        <Row label="(=) LUCRO LÍQUIDO DO PERÍODO" value={d.ll} bold color={d.ll>=0?"acc":"neg"} sep/>
        <div className="px-3 py-2">
          <span className={`text-xs font-mono ${d.ml>=20?"text-emerald-500":d.ml>=10?"text-amber-500":"text-rose-500"}`}>
            Margem Líquida: {fmtPct(d.ml)} {d.ml>=20?"✅":d.ml>=10?"⚠️":"🚨"}
          </span>
        </div>
      </Card>
    </Section>
  )
}


// ─── PERSISTENCE HOOK ─────────────────────────────────────────────────────────
function useLocalStorage(key, initial) {
  const [state, setState] = useState(() => {
    try {
      const v = localStorage.getItem(key)
      return v ? JSON.parse(v) : initial
    } catch { return initial }
  })
  const set = (val) => setState(prev => {
    const next = typeof val === 'function' ? val(prev) : val
    try { localStorage.setItem(key, JSON.stringify(next)) } catch {}
    return next
  })
  return [state, set]
}

// ─── GROQ AI (free — 14.400 req/dia) ─────────────────────────────────────────
const GROQ_MODEL = "llama-3.3-70b-versatile"

async function callGroq(messages, apiKey) {
  if (!apiKey) throw new Error('Configure sua chave Groq em ⚙️ Configurações')
  const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` },
    body: JSON.stringify({ model: GROQ_MODEL, messages, max_tokens: 1000, temperature: 0.4 })
  })
  const data = await res.json()
  if (!res.ok || data.error) throw new Error(data.error?.message || `Erro ${res.status}`)
  return data.choices?.[0]?.message?.content || ''
}

// Text extraction for import (Groq doesn't do multimodal — we extract text client-side)
async function callGemini(parts, apiKey) {
  // Extract text content from parts and call Groq
  const textParts = parts.filter(p => p.text).map(p => p.text).join('\n')
  if (!textParts) throw new Error('Formato não suportado. Use CSV, XLSX ou cole o texto.')
  return callGroq([{ role:'user', content: textParts }], apiKey)
}

async function callGeminiChat(history, msg, apiKey, sysPrompt) {
  const messages = [
    { role: 'system', content: sysPrompt },
    ...history.slice(-8).map(m => ({ role: m.r==='user'?'user':'assistant', content: m.t })),
    { role: 'user', content: msg }
  ]
  return callGroq(messages, apiKey)
}

// ─── SETTINGS MODAL ───────────────────────────────────────────────────────────
function SettingsModal({ apiKey, setApiKey, onClose }) {
  const [input, setInput] = useState(apiKey || '')
  const save = () => { setApiKey(input.trim()); onClose() }
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-sm">
      <div className="bg-zinc-900 border border-zinc-800 rounded-2xl p-6 w-full max-w-md shadow-2xl">
        <div className="flex items-center justify-between mb-5">
          <div>
            <h2 className="text-lg font-bold text-zinc-100">Configurações de IA</h2>
            <p className="text-zinc-500 text-sm mt-0.5">Groq — 14.400 req/dia grátis</p>
          </div>
          <button onClick={onClose} className="text-zinc-500 hover:text-zinc-200 transition-colors"><X size={18}/></button>
        </div>
        <div className="space-y-4">
          <div className="bg-emerald-500/5 border border-emerald-500/20 rounded-xl p-3 text-xs text-emerald-300">
            ✅ Groq é <strong>100% gratuito</strong> — 14.400 requisições/dia, sem cartão de crédito.
          </div>
          <div>
            <label className="text-xs text-zinc-500 block mb-1">Chave da API Groq</label>
            <input
              type="password"
              className={`${INP} font-mono`}
              value={input}
              onChange={e => setInput(e.target.value)}
              placeholder="gsk_..."
              onKeyDown={e => e.key === 'Enter' && save()}
            />
          </div>
          <a
            href="https://console.groq.com/keys"
            target="_blank"
            rel="noopener noreferrer"
            className="flex items-center gap-2 text-sm text-violet-400 hover:text-violet-300 transition-colors"
          >
            <Sparkles size={14}/>
            Obter chave gratuita em console.groq.com →
          </a>
        </div>
        <div className="flex gap-2 mt-5">
          <button onClick={save} className="flex-1 bg-emerald-500 hover:bg-emerald-400 text-zinc-900 font-semibold text-sm rounded-lg py-2.5 transition-colors">
            Salvar
          </button>
          <button onClick={onClose} className="bg-zinc-800 hover:bg-zinc-700 text-zinc-300 text-sm rounded-lg px-4 py-2 transition-colors">
            Cancelar
          </button>
        </div>
        {apiKey && (
          <p className="text-xs text-zinc-600 text-center mt-3">
            Chave configurada: {apiKey.slice(0,8)}...{apiKey.slice(-4)}
          </p>
        )}
      </div>
    </div>
  )
}

// ─── IMPORTAÇÃO IA ────────────────────────────────────────────────────────────
const EXTRACTION_PROMPT = `Especialista em finanças brasileiras. Extraia TODAS as movimentações financeiras do conteúdo abaixo.

Categorias ENTRADA: Serviços, Eventos, Vendas de Produtos, Reembolsos, Receitas Extras, Outros
Categorias SAÍDA: Aluguel, Energia/Água, Folha de Pagamento, Freelancers, Impostos, Marketing, Manutenção, Materiais, Plataformas/Software, Equipamentos, Despesas Operacionais, Outros

Regras: Meta/Facebook/Google Ads→Marketing | Aluguel/Locação→Aluguel | Salário/Folha→Folha de Pagamento | Energia/Luz/Água→Energia/Água | DARF/DAS/Simples→Impostos | PIX recebido/Depósito/Crédito→Receitas Extras(entrada) | Pagamento/Débito→saida

RETORNE APENAS JSON ARRAY sem markdown sem backticks:
[{"name":"desc","date":"YYYY-MM-DD","value":123.45,"type":"entrada","category":"Eventos","payment":"PIX","obs":""}]
Se vazio retorne: []`

function ImportacaoIA({txs,setTxs}) {
  const [tab,setTab] = useState("arquivo")
  const [stage,setStage] = useState("idle")
  const [items,setItems] = useState([])
  const [sel,setSel] = useState(new Set())
  const [dupes,setDupes] = useState(new Set())
  const [analysis,setAnalysis] = useState("")
  const [msg,setMsg] = useState("")
  const [pasted,setPasted] = useState("")
  const [history,setHistory] = useState([])

  const callClaude = (parts) => callGemini(parts, apiKey)

  const process = async (blocks,name) => {
    setStage("loading"); setMsg("IA analisando..."); setAnalysis("")
    try {
      const raw = await callGemini(blocks, apiKey)
      let parsed=[]
      try{const clean=raw.replace(/```json|```/g,"").trim();const m=clean.match(/\[[\s\S]*\]/);if(m)parsed=JSON.parse(m[0]);else if(clean.startsWith("["))parsed=JSON.parse(clean)}catch(e){console.error(e)}
      if(!parsed.length){setAnalysis("Nenhuma movimentação encontrada. Tente colar o texto do extrato.");setItems([]);setStage("preview");return}
      const withMeta = parsed.map((t,i)=>({
        ...t,id:`i_${Date.now()}_${i}`,
        value:typeof t.value==="string"?parseFloat(String(t.value).replace(/\./g,"").replace(",","."))||0:Number(t.value)||0,
        status:"pago",recurrence:"none",competencia:t.date?.slice(0,7)||new Date().toISOString().slice(0,7),
        payment:t.payment||"PIX",type:["entrada","saida"].includes(t.type)?t.type:"saida",
      }))
      const dupeSet = new Set(withMeta.filter(t=>txs.some(e=>Math.abs(e.value-t.value)<0.01&&e.date===t.date&&e.type===t.type)).map(t=>t.id))
      setDupes(dupeSet); setSel(new Set(withMeta.filter(t=>!dupeSet.has(t.id)).map(t=>t.id))); setItems(withMeta)
      try{
        const tE=withMeta.filter(t=>t.type==="entrada").reduce((a,t)=>a+t.value,0)
        const tS=withMeta.filter(t=>t.type==="saida").reduce((a,t)=>a+t.value,0)
        const a=await callClaude([{text:`CFO virtual: analise em 2 bullets concisos (pt-BR): Entradas ${fmt(tE)}, Saídas ${fmt(tS)}, Resultado ${fmt(tE-tS)}.`}])
        setAnalysis(a)
      }catch{setAnalysis("")}
      setStage("preview")
    }catch(e){setAnalysis(e.message);setStage("error")}
  }

  const handleFile = async file => {
    if(!file) return
    const ext=file.name.split(".").pop().toLowerCase()
    setMsg("Lendo arquivo...")
    try{
      const toB64=f=>new Promise((r,j)=>{const rd=new FileReader();rd.onload=()=>r(rd.result.split(",")[1]);rd.onerror=j;rd.readAsDataURL(f)})
      const toTxt=f=>new Promise((r,j)=>{const rd=new FileReader();rd.onload=()=>r(rd.result);rd.onerror=j;rd.readAsText(f,"utf-8")})
      if(["jpg","jpeg","png","gif","webp"].includes(ext)){
        setErrorMsg("Para imagens, copie o texto do extrato e use a aba 'Colar Texto'.")
        setStage("error")
        return
      }else if(ext==="pdf"){
        // Try to read PDF as text first
        setMsg("Tentando extrair texto do PDF...")
        try {
          const txt = await toTxt(file)
          if (txt && txt.trim().length > 50) {
            await process([{text:EXTRACTION_PROMPT+"\n\n"+txt.slice(0,7000)}],file.name)
          } else {
            setErrorMsg("PDF sem texto legível. Abra o PDF, selecione todo o texto (Ctrl+A), copie e use a aba 'Colar Texto'.")
            setStage("error")
          }
        } catch {
          setErrorMsg("Não foi possível ler o PDF. Use a aba 'Colar Texto' e cole o conteúdo do extrato.")
          setStage("error")
        }
        return
      }else if(["xlsx","xls"].includes(ext)){
        const buf=await file.arrayBuffer()
        const wb=XLSX.read(new Uint8Array(buf),{type:"array"})
        const csv=XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]])
        await process([{text:EXTRACTION_PROMPT+"\n\n"+csv.slice(0,7000)}],file.name)
      }else{
        const txt=await toTxt(file)
        await process([{text:EXTRACTION_PROMPT+"\n\n"+txt.slice(0,7000)}],file.name)
      }
    }catch(e){setAnalysis(e.message);setStage("error")}
  }

  const importSel = () => {
    const toImport=items.filter(t=>sel.has(t.id))
    setTxs(p=>[...p,...toImport])
    setHistory(h=>[...h,{date:new Date().toLocaleDateString("pt-BR"),count:toImport.length,total:toImport.reduce((a,t)=>a+(t.type==="entrada"?t.value:-t.value),0)}])
    setStage("done")
  }
  const reset = () => {setStage("idle");setItems([]);setSel(new Set());setDupes(new Set());setAnalysis("");setMsg("");setPasted("")}
  const toggleSel = id => setSel(s=>{const n=new Set(s);n.has(id)?n.delete(id):n.add(id);return n})
  const updateItem = (id,k,v) => setItems(p=>p.map(t=>t.id===id?{...t,[k]:k==="value"?parseFloat(v)||0:v}:t))

  const totE=items.filter(t=>t.type==="entrada"&&sel.has(t.id)).reduce((a,t)=>a+t.value,0)
  const totS=items.filter(t=>t.type==="saida"&&sel.has(t.id)).reduce((a,t)=>a+t.value,0)

  return (
    <Section title="Importação com IA" sub="Extração automática de extratos, planilhas, PDFs e imagens">
      {stage!=="idle"&&<button onClick={reset} className="flex items-center gap-2 text-xs bg-zinc-800 hover:bg-zinc-700 border border-zinc-700 text-zinc-300 rounded-lg px-3 py-2 transition-colors w-fit"><RefreshCw size={12}/>Nova Importação</button>}
      {stage==="idle"&&(
        <div className="space-y-4">
          <div className="flex items-center gap-1 bg-zinc-800/80 border border-zinc-700/40 rounded-xl p-1 w-fit">
            <button onClick={()=>setTab("arquivo")} className={`flex items-center gap-2 text-sm px-4 py-2 rounded-lg font-medium transition-colors ${tab==="arquivo"?"bg-zinc-700 text-zinc-100":"text-zinc-500 hover:text-zinc-300"}`}><FileScan size={14}/>Arquivo</button>
            <button onClick={()=>setTab("colar")} className={`flex items-center gap-2 text-sm px-4 py-2 rounded-lg font-medium transition-colors ${tab==="colar"?"bg-zinc-700 text-zinc-100":"text-zinc-500 hover:text-zinc-300"}`}><FileText size={14}/>Colar Texto</button>
          </div>
          {tab==="arquivo"&&(
            <label className="block cursor-pointer">
              <input type="file" accept=".pdf,.xlsx,.xls,.csv,.ofx,.txt,.jpg,.jpeg,.png" className="sr-only" onChange={e=>{const f=e.target.files?.[0];if(f)handleFile(f);e.target.value=""}}/>
              <div className="border-2 border-dashed border-zinc-700 hover:border-violet-500 active:border-violet-400 rounded-2xl p-10 text-center transition-colors">
                <div className="w-16 h-16 rounded-2xl bg-zinc-800 flex items-center justify-center mx-auto mb-4"><Upload size={28} className="text-zinc-500"/></div>
                <p className="text-zinc-200 font-semibold mb-1">Toque para selecionar arquivo</p>
                <p className="text-zinc-500 text-sm mb-5">PDF · XLSX · CSV · JPG · PNG</p>
                <span className="inline-flex items-center gap-2 bg-violet-500 text-white text-sm font-semibold rounded-xl px-6 py-3"><FileScan size={15}/>Selecionar Arquivo</span>
              </div>
            </label>
          )}
          {tab==="colar"&&(
            <Card>
              <p className="text-sm font-semibold text-zinc-300 mb-1">Cole o conteúdo do extrato</p>
              <p className="text-xs text-zinc-500 mb-3">Funciona com extrato bancário copiado, CSV em texto, lista de pagamentos com datas e valores.</p>
              <textarea className={`${INP} resize-none font-mono text-xs`} rows={10} placeholder={"01/05/2025  PIX RECEBIDO FESTA INFANTIL  +1800,00\n05/05/2025  PAGAMENTO ALUGUEL  -2200,00\n10/05/2025  ENERGIA ELETRICA  -480,00"} value={pasted} onChange={e=>setPasted(e.target.value)}/>
              <button onClick={()=>process([{text:EXTRACTION_PROMPT+"\n\n"+pasted.slice(0,7000)}],"Texto colado")} disabled={!pasted.trim()} className="mt-3 w-full flex items-center justify-center gap-2 bg-violet-500 hover:bg-violet-400 disabled:opacity-40 text-white font-semibold text-sm rounded-xl py-3 transition-colors"><Brain size={15}/>Analisar com IA</button>
            </Card>
          )}
          {history.length>0&&(
            <Card>
              <p className="text-sm font-semibold text-zinc-300 mb-3">Histórico</p>
              {history.map((h,i)=>(
                <div key={i} className="flex items-center justify-between bg-zinc-800/30 rounded-lg px-3 py-2 text-xs mb-1">
                  <div className="flex items-center gap-2"><CheckCircle2 size={11} className="text-emerald-400"/><span className="text-zinc-400">{h.date}</span><span className="text-zinc-600">{h.count} lançamentos</span></div>
                  <span className={`font-mono font-bold ${h.total>=0?"text-emerald-400":"text-rose-400"}`}>{fmt(h.total)}</span>
                </div>
              ))}
            </Card>
          )}
        </div>
      )}
      {stage==="loading"&&(
        <Card className="text-center py-12">
          <div className="w-14 h-14 rounded-2xl bg-violet-500/10 flex items-center justify-center mx-auto mb-4 animate-pulse"><Brain size={26} className="text-violet-400"/></div>
          <p className="text-zinc-200 font-semibold mb-1">{msg}</p>
          <div className="flex justify-center gap-1.5 mt-4">{[0,1,2,3,4].map(i=><div key={i} className="w-1.5 h-1.5 rounded-full bg-violet-400 animate-bounce" style={{animationDelay:`${i*100}ms`}}/>)}</div>
        </Card>
      )}
      {stage==="error"&&(
        <div className="bg-rose-500/5 border border-rose-500/20 rounded-xl p-6 text-center">
          <AlertCircle size={28} className="text-rose-400 mx-auto mb-3"/>
          <p className="text-rose-300 font-semibold mb-2">Erro ao processar</p>
          <p className="text-zinc-400 text-sm mb-4">{analysis}</p>
          <div className="flex justify-center gap-3">
            <button onClick={reset} className="bg-zinc-800 hover:bg-zinc-700 text-zinc-300 text-sm rounded-lg px-4 py-2">Tentar novamente</button>
            <button onClick={()=>{reset();setTab("colar")}} className="bg-violet-500/10 text-violet-400 border border-violet-500/20 text-sm rounded-lg px-4 py-2">Colar Texto</button>
          </div>
        </div>
      )}
      {stage==="preview"&&(
        <div className="space-y-4">
          <div className="grid grid-cols-2 xl:grid-cols-4 gap-3">
            {[{l:"Encontrados",v:items.length,c:CLR.violet},{l:"Entradas",v:fmt(totE),c:CLR.emerald},{l:"Saídas",v:fmt(totS),c:CLR.rose},{l:"Resultado",v:fmt(totE-totS),c:totE-totS>=0?CLR.indigo:CLR.rose}].map((k,i)=>(
              <div key={i} className={`border rounded-xl p-3 text-center ${k.c.card}`}><p className="text-xs text-zinc-500">{k.l}</p><p className={`text-lg font-bold font-mono ${k.c.text}`}>{k.v}</p></div>
            ))}
          </div>
          {dupes.size>0&&<div className="flex items-center gap-2 bg-amber-500/5 border border-amber-500/20 rounded-xl p-3 text-sm text-amber-300"><TriangleAlert size={14} className="flex-shrink-0"/>{dupes.size} possível{dupes.size>1?"is":""} duplicata{dupes.size>1?"s":""} desmarcada{dupes.size>1?"s":""}.</div>}
          {analysis&&<div className="bg-violet-500/5 border border-violet-500/20 rounded-xl p-4"><div className="flex items-center gap-2 mb-2"><Sparkles size={13} className="text-violet-400"/><p className="text-sm font-semibold text-violet-300">Análise da IA</p></div><p className="text-sm text-zinc-400 whitespace-pre-line">{analysis}</p></div>}
          {items.length===0
            ?<Card className="text-center py-8"><p className="text-zinc-500">Nenhuma movimentação encontrada.</p><button onClick={()=>{reset();setTab("colar")}} className="mt-3 text-sm text-violet-400 underline">Tente colar o texto</button></Card>
            :<>
              <div className="flex items-center justify-between">
                <p className="text-sm font-semibold text-zinc-300">Revisar ({items.length})</p>
                <div className="flex items-center gap-3"><button onClick={()=>setSel(s=>s.size===items.length?new Set():new Set(items.map(t=>t.id)))} className="text-xs text-zinc-400 hover:text-zinc-200 underline">{sel.size===items.length?"Desmarcar":"Selecionar todos"}</button><span className="text-xs text-zinc-600">{sel.size}/{items.length}</span></div>
              </div>
              <Card className="p-0 overflow-hidden">
                <div className="divide-y divide-zinc-800/40 max-h-[400px] overflow-y-auto">
                  {items.map(t=>(
                    <div key={t.id} className={`px-3 py-3 transition-colors ${sel.has(t.id)?"":"opacity-40"} ${dupes.has(t.id)?"bg-amber-500/5":""}`}>
                      <div className="flex items-start gap-3">
                        <button onClick={()=>toggleSel(t.id)} className={`w-5 h-5 rounded flex items-center justify-center flex-shrink-0 mt-0.5 border transition-colors ${sel.has(t.id)?"bg-violet-500 border-violet-500":"border-zinc-600"}`}>{sel.has(t.id)&&<Check size={11} className="text-white"/>}</button>
                        <div className="flex-1 min-w-0 space-y-1.5">
                          <input className="w-full bg-transparent text-sm text-zinc-200 font-medium outline-none focus:bg-zinc-800 rounded px-1" value={t.name} onChange={e=>updateItem(t.id,"name",e.target.value)}/>
                          <div className="flex flex-wrap items-center gap-2">
                            <span className={`w-1.5 h-1.5 rounded-full ${t.type==="entrada"?"bg-emerald-400":"bg-rose-400"}`}/>
                            <select className="bg-zinc-800 border border-zinc-700 rounded px-1.5 py-0.5 text-xs text-zinc-300 outline-none" value={t.category} onChange={e=>updateItem(t.id,"category",e.target.value)}>{(CATS[t.type]||[]).map(c=><option key={c}>{c}</option>)}</select>
                            <input type="date" className="bg-zinc-800 border border-zinc-700 rounded px-1.5 py-0.5 text-xs text-zinc-400 outline-none" value={t.date} onChange={e=>updateItem(t.id,"date",e.target.value)}/>
                            {dupes.has(t.id)&&<span className="text-[10px] text-amber-400 bg-amber-400/10 rounded px-1">duplicata?</span>}
                          </div>
                        </div>
                        <div className="flex flex-col items-end gap-1 flex-shrink-0">
                          <input type="number" className={`font-mono font-bold text-sm text-right outline-none bg-transparent w-24 focus:bg-zinc-800 rounded px-1 ${t.type==="entrada"?"text-emerald-400":"text-rose-400"}`} value={t.value} onChange={e=>updateItem(t.id,"value",e.target.value)}/>
                          <button onClick={()=>setItems(p=>p.filter(x=>x.id!==t.id))} className="text-zinc-700 hover:text-rose-400 p-0.5"><X size={12}/></button>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </Card>
              <div className="flex items-center gap-3 flex-wrap">
                <button onClick={importSel} disabled={!sel.size} className="flex items-center gap-2 bg-violet-500 hover:bg-violet-400 disabled:opacity-40 text-white font-semibold text-sm rounded-xl px-5 py-3 transition-colors flex-1 justify-center sm:flex-none"><FileUp size={15}/>Importar {sel.size} lançamento{sel.size!==1?"s":""}</button>
                <button onClick={reset} className="text-sm text-zinc-500 hover:text-zinc-300 px-3">Cancelar</button>
              </div>
            </>
          }
        </div>
      )}
      {stage==="done"&&(
        <Card className="text-center py-10">
          <div className="w-14 h-14 rounded-2xl bg-emerald-500/20 flex items-center justify-center mx-auto mb-4"><CheckCircle2 size={28} className="text-emerald-400"/></div>
          <p className="text-emerald-300 font-bold text-lg mb-1">Importação concluída!</p>
          <p className="text-zinc-500 text-sm mb-6">Lançamentos adicionados com sucesso.</p>
          <button onClick={reset} className="bg-zinc-800 hover:bg-zinc-700 text-zinc-200 text-sm rounded-xl px-5 py-2.5 transition-colors">Nova Importação</button>
        </Card>
      )}
    </Section>
  )
}

// ─── INTELIGÊNCIA FINANCEIRA (CHAT IA) ────────────────────────────────────────
function InteligenciaFinanceira({fin,txs,apiKey}) {
  const [chat,setChat] = useState([])
  const [input,setInput] = useState("")
  const [loading,setLoading] = useState(false)
  const bottomRef = useRef()

  const sendMessage = async msg => {
    if(!msg.trim()||loading) return
    setChat(p=>[...p,{r:"user",t:msg}]); setInput(""); setLoading(true)
    try{
      const ctx = {faturamentoMes:fmt(fin.cmE),custosMes:fmt(fin.cmS),lucroMes:fmt(fin.cmL),margem:fmtPct(fin.cmM),faturamentoAnterior:fmt(fin.pmE),lucroAnterior:fmt(fin.pmL),custoFixo:fmt(fin.fixedCosts),custoVariavel:fmt(fin.varCosts),ticketMedio:fmt(fin.ticketMedio),contasVencer:fin.pending.length,totalLancamentos:txs.length,variacaoFaturamento:fin.pmE>0?`${((fin.cmE/fin.pmE-1)*100).toFixed(1)}%`:"N/A"}
      const sysPrompt = `Você é CFO virtual de PME brasileira. Dados financeiros: ${JSON.stringify(ctx)}. Responda em português, máximo 180 palavras, use bullets • quando listar. Direto e acionável.`
      const reply = await callGeminiChat(chat, msg, apiKey, sysPrompt)
      setChat(p=>[...p,{r:"ai",t:reply||"Sem resposta."}])
    }catch(e){setChat(p=>[...p,{r:"ai",t:`❌ ${e.message}`}])}
    setLoading(false)
    setTimeout(()=>bottomRef.current?.scrollIntoView({behavior:"smooth"}),80)
  }

  const marginData = fin.monthly.map(m=>({...m,margem:m.receita>0?(m.lucro/m.receita)*100:0}))
  const PROMPTS = [
    {l:"📊 Análise completa",q:"Faça uma análise completa da saúde financeira."},
    {l:"⚠️ Riscos",q:"Quais os principais riscos que devo agir agora?"},
    {l:"💰 Aumentar margem",q:"Como aumentar a margem de lucro com ações práticas?"},
    {l:"🎯 Break-even",q:"Como atingir o ponto de equilíbrio mais rápido?"},
    {l:"💸 Cortar custos",q:"Onde estou gastando acima do ideal?"},
    {l:"📈 Crescimento",q:"O que está limitando meu crescimento?"},
  ]

  return (
    <Section title="Inteligência Financeira" sub="CFO Virtual — chat baseado nos seus dados reais">
      <div className="grid grid-cols-2 xl:grid-cols-4 gap-3">
        {[{l:"Margem Atual",v:fmtPct(fin.cmM),c:fin.cmM>=20?"emerald":fin.cmM>=10?"amber":"rose"},{l:"Vs Mês Anterior",v:`${fin.pmE>0?((fin.cmE/fin.pmE-1)*100).toFixed(1):0}%`,c:fin.cmE>=fin.pmE?"emerald":"rose"},{l:"Custos Fixos",v:fmtPct(fin.fixedCosts/(fin.totalS||1)*100),c:fin.fixedCosts/(fin.totalS||1)<0.6?"emerald":"amber"},{l:"A Pagar",v:fmt(fin.pending.reduce((a,t)=>a+t.value,0)),c:fin.pending.length?"amber":"emerald"}].map((k,i)=>(
          <div key={i} className={`border rounded-xl p-3 ${CLR[k.c].card}`}>
            <p className="text-[11px] text-zinc-500 mb-0.5">{k.l}</p>
            <p className={`text-lg font-bold font-mono ${CLR[k.c].text}`}>{k.v}</p>
          </div>
        ))}
      </div>
      <div className="bg-zinc-900/60 border border-zinc-800 rounded-xl overflow-hidden" style={{minHeight:"380px",maxHeight:"520px",display:"flex",flexDirection:"column"}}>
        <div className="flex items-center gap-2 px-4 py-3 border-b border-zinc-800 flex-shrink-0">
          <Brain size={14} className="text-violet-400"/>
          <span className="text-sm font-semibold text-zinc-300">CFO Virtual</span>
          <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse ml-1"/>
          {chat.length>0&&<button onClick={()=>setChat([])} className="ml-auto text-xs text-zinc-600 hover:text-zinc-400 flex items-center gap-1"><Trash2 size={11}/>Limpar</button>}
        </div>
        <div className="flex-1 overflow-y-auto p-4 space-y-3">
          {!chat.length&&(
            <div className="text-center py-8">
              <Brain size={30} className="text-zinc-700 mx-auto mb-3"/>
              <p className="text-zinc-500 text-sm">Seu CFO virtual está pronto.</p>
              <p className="text-zinc-600 text-xs mt-1">Use os atalhos abaixo ou faça sua pergunta.</p>
            </div>
          )}
          {chat.map((m,i)=>(
            <div key={i} className={`flex ${m.r==="user"?"justify-end":"justify-start"}`}>
              {m.r==="ai"&&<div className="w-6 h-6 rounded-lg bg-violet-500/20 flex items-center justify-center mr-2 flex-shrink-0 mt-0.5"><Brain size={11} className="text-violet-400"/></div>}
              <div className={`max-w-[85%] rounded-xl px-3 py-2.5 text-sm leading-relaxed whitespace-pre-line ${m.r==="user"?"bg-violet-500/20 text-violet-100 rounded-br-sm":"bg-zinc-800 text-zinc-300 rounded-bl-sm"}`}>{m.t}</div>
            </div>
          ))}
          {loading&&<div className="flex"><div className="w-6 h-6 rounded-lg bg-violet-500/20 flex items-center justify-center mr-2"><Brain size={11} className="text-violet-400"/></div><div className="bg-zinc-800 rounded-xl rounded-bl-sm px-4 py-3 flex gap-1">{[0,1,2].map(i=><div key={i} className="w-1.5 h-1.5 rounded-full bg-zinc-500 animate-bounce" style={{animationDelay:`${i*150}ms`}}/>)}</div></div>}
          <div ref={bottomRef}/>
        </div>
        <div className="px-3 py-2 border-t border-zinc-800/60 flex gap-1.5 overflow-x-auto flex-shrink-0">
          {PROMPTS.map((p,i)=><button key={i} onClick={()=>sendMessage(p.q)} disabled={loading} className="text-xs bg-zinc-800 hover:bg-zinc-700 text-zinc-400 rounded-lg px-2.5 py-1.5 whitespace-nowrap transition-colors disabled:opacity-40 flex-shrink-0">{p.l}</button>)}
        </div>
        <div className="flex items-center gap-2 px-3 py-2.5 border-t border-zinc-800 flex-shrink-0">
          <input className="flex-1 bg-zinc-800 border border-zinc-700 rounded-lg px-3 py-2 text-sm text-zinc-200 outline-none focus:border-violet-500 placeholder-zinc-600" placeholder="Pergunte sobre suas finanças..." value={input} onChange={e=>setInput(e.target.value)} onKeyDown={e=>e.key==="Enter"&&!e.shiftKey&&sendMessage(input)} disabled={loading}/>
          <button onClick={()=>sendMessage(input)} disabled={loading||!input.trim()} className="w-9 h-9 flex items-center justify-center bg-violet-500 hover:bg-violet-400 disabled:opacity-40 rounded-lg transition-colors flex-shrink-0"><ArrowUpRight size={16} className="text-white"/></button>
        </div>
      </div>
      <Card>
        <p className="text-sm font-semibold text-zinc-300 mb-4">Evolução da Margem Líquida</p>
        <ResponsiveContainer width="100%" height={180}>
          <AreaChart data={marginData}>
            <defs><linearGradient id="gML" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#a78bfa" stopOpacity={0.3}/><stop offset="95%" stopColor="#a78bfa" stopOpacity={0}/></linearGradient></defs>
            <CartesianGrid strokeDasharray="3 3" stroke="#27272a"/>
            <XAxis dataKey="label" tick={{fill:"#71717a",fontSize:10}} axisLine={false} tickLine={false}/>
            <YAxis tick={{fill:"#71717a",fontSize:10}} axisLine={false} tickLine={false} tickFormatter={v=>`${v.toFixed(0)}%`} width={30}/>
            <Tooltip formatter={v=>[fmtPct(v),"Margem"]}/>
            <ReferenceLine y={20} stroke="#10b981" strokeDasharray="4" label={{value:"Meta 20%",fill:"#10b981",fontSize:10}}/>
            <Area dataKey="margem" name="Margem" stroke="#a78bfa" fill="url(#gML)" strokeWidth={2.5}/>
          </AreaChart>
        </ResponsiveContainer>
      </Card>
    </Section>
  )
}

// ─── NAV CONFIG ───────────────────────────────────────────────────────────────
const NAV = [
  {id:"dashboard",label:"Dashboard",icon:LayoutDashboard},
  {id:"lancamentos",label:"Lançamentos",icon:ArrowLeftRight},
  {id:"fluxo",label:"Fluxo de Caixa",icon:Activity},
  {id:"precificacao",label:"Precificação",icon:Calculator},
  {id:"breakeven",label:"Break-even",icon:Target},
  {id:"dre",label:"DRE",icon:FileText},
  {id:"importacao",label:"Importação IA",icon:Brain},
  {id:"inteligencia",label:"Inteligência Fin.",icon:Sparkles},
]

// ─── APP SHELL ────────────────────────────────────────────────────────────────
export default function App() {
  const [page,setPage] = useState("dashboard")
  const [dark,setDark] = useLocalStorage("erp_dark", true)
  const [collapsed,setCollapsed] = useState(false)
  const [mobileMenu,setMobileMenu] = useState(false)
  const [txs,setTxs] = useLocalStorage("erp_txs", SEED)
  const [svcs,setSvcs] = useLocalStorage("erp_svcs", SEED_SERVICES)
  const [apiKey,setApiKey] = useLocalStorage("erp_ai_key", import.meta.env.VITE_GROQ_KEY || "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")
  const [showSettings,setShowSettings] = useState(false)
  const [showReset,setShowReset] = useState(false)
  const fin = useFinancials(txs)
  const cur = NAV.find(n=>n.id===page)

  const go = id => { setPage(id); setMobileMenu(false) }

  const PAGES = {
    dashboard:<Dashboard fin={fin} txs={txs}/>,
    lancamentos:<Lancamentos txs={txs} setTxs={setTxs}/>,
    fluxo:<FluxoCaixa fin={fin}/>,
    precificacao:<Precificacao services={svcs} setServices={setSvcs}/>,
    breakeven:<PontoEquilibrio fin={fin} txs={txs}/>,
    dre:<DRE txs={txs}/>,
    importacao:<ImportacaoIA txs={txs} setTxs={setTxs} apiKey={apiKey}/>,
    inteligencia:<InteligenciaFinanceira fin={fin} txs={txs} apiKey={apiKey}/>,
  }

  const SideItem = ({n}) => (
    <button onClick={()=>go(n.id)} title={collapsed?n.label:""} className={`w-full flex items-center gap-3 rounded-lg h-9 transition-all ${collapsed?"justify-center px-0":"px-3"} ${page===n.id?"bg-emerald-500/10 text-emerald-400":"text-zinc-500 hover:text-zinc-200 hover:bg-zinc-800/60"}`}>
      <n.icon size={15} className="flex-shrink-0"/>
      {!collapsed&&<span className="text-sm font-medium truncate">{n.label}</span>}
      {!collapsed&&n.id==="importacao"&&<span className="ml-auto text-[9px] bg-violet-500/20 text-violet-400 border border-violet-500/20 rounded-full px-1.5 py-0.5 font-bold">IA</span>}
    </button>
  )

  return (
    <div className={`flex ${dark?"bg-zinc-950 text-zinc-100":"bg-slate-50 text-zinc-900"}`} style={{height:"100dvh",overflow:"hidden"}}>

      {/* Sidebar desktop */}
      <aside className={`hidden md:flex flex-col flex-shrink-0 transition-all duration-300 border-r ${dark?"bg-zinc-900 border-zinc-800":"bg-white border-zinc-200"} ${collapsed?"w-[60px]":"w-[220px]"}`}>
        <div className={`flex items-center gap-3 h-14 px-3 border-b ${dark?"border-zinc-800":"border-zinc-200"} flex-shrink-0 overflow-hidden`}>
          <div className="w-8 h-8 rounded-lg bg-emerald-500 flex items-center justify-center flex-shrink-0"><CircleDollarSign size={16} className="text-zinc-900"/></div>
          {!collapsed&&<div><p className="text-sm font-bold leading-none">FinanceiroERP</p><p className="text-[10px] text-zinc-500">Gestão Empresarial</p></div>}
        </div>
        <nav className="flex-1 py-2 px-2 space-y-0.5 overflow-y-auto overflow-x-hidden">
          {NAV.map(n=><SideItem key={n.id} n={n}/>)}
        </nav>
        <div className={`flex-shrink-0 border-t ${dark?"border-zinc-800":"border-zinc-200"} py-2 px-2 space-y-0.5`}>
          <button onClick={()=>setDark(!dark)} className={`w-full flex items-center gap-3 rounded-lg h-9 transition-colors ${collapsed?"justify-center px-0":"px-3"} text-zinc-500 hover:text-zinc-200 hover:bg-zinc-800/60`} title={collapsed?(dark?"Claro":"Escuro"):""}>
            {dark?<Sun size={14}/>:<Moon size={14}/>}
            {!collapsed&&<span className="text-sm">{dark?"Modo Claro":"Modo Escuro"}</span>}
          </button>
          <button onClick={()=>setShowReset(true)} className={`w-full flex items-center gap-3 rounded-lg h-9 transition-colors ${collapsed?"justify-center px-0":"px-3"} text-rose-800 hover:text-rose-400 hover:bg-rose-500/10`} title={collapsed?"Resetar dados":""}>
            <Trash2 size={14} className="flex-shrink-0"/>
            {!collapsed&&<span className="text-sm">Resetar dados</span>}
          </button>
          <button onClick={()=>setCollapsed(!collapsed)} className={`w-full flex items-center gap-3 rounded-lg h-9 transition-colors ${collapsed?"justify-center px-0":"px-3"} text-zinc-600 hover:text-zinc-300 hover:bg-zinc-800/60`}>
            {collapsed?<ChevronRight size={14}/>:<ChevronLeft size={14}/>}
            {!collapsed&&<span className="text-sm text-zinc-500">Recolher</span>}
          </button>
        </div>
      </aside>

      {showSettings && <SettingsModal apiKey={apiKey} setApiKey={setApiKey} onClose={()=>setShowSettings(false)}/>}

      {/* Reset confirmation */}
      {showReset && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-sm">
          <div className="bg-zinc-900 border border-zinc-800 rounded-2xl p-6 w-full max-w-sm shadow-2xl text-center">
            <div className="w-12 h-12 rounded-full bg-rose-500/10 flex items-center justify-center mx-auto mb-4">
              <Trash2 size={22} className="text-rose-400"/>
            </div>
            <h2 className="text-lg font-bold text-zinc-100 mb-2">Resetar todos os dados?</h2>
            <p className="text-zinc-500 text-sm mb-5">Todos os lançamentos, serviços e configurações serão apagados. Esta ação não pode ser desfeita.</p>
            <div className="flex gap-3">
              <button onClick={()=>setShowReset(false)} className="flex-1 bg-zinc-800 hover:bg-zinc-700 text-zinc-300 text-sm rounded-xl py-2.5 transition-colors">Cancelar</button>
              <button onClick={()=>{
                setTxs([])
                setSvcs([])
                try {
                  localStorage.removeItem("erp_txs")
                  localStorage.removeItem("erp_svcs")
                } catch{}
                setShowReset(false)
                setPage("lancamentos")
              }} className="flex-1 bg-rose-500 hover:bg-rose-400 text-white font-semibold text-sm rounded-xl py-2.5 transition-colors">
                Sim, resetar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Mobile drawer */}
      {mobileMenu&&(
        <div className="fixed inset-0 z-50 md:hidden">
          <div className="absolute inset-0 bg-black/60 backdrop-blur-sm" onClick={()=>setMobileMenu(false)}/>
          <div className={`absolute bottom-0 left-0 right-0 rounded-t-2xl border-t ${dark?"bg-zinc-900 border-zinc-800":"bg-white border-zinc-200"}`} style={{paddingBottom:"env(safe-area-inset-bottom)"}}>
            <div className={`flex items-center justify-between px-4 py-3 border-b ${dark?"border-zinc-800":"border-zinc-200"}`}>
              <div className="flex items-center gap-2"><div className="w-6 h-6 rounded-lg bg-emerald-500 flex items-center justify-center"><CircleDollarSign size={12} className="text-zinc-900"/></div><span className="text-sm font-bold">FinanceiroERP</span></div>
              <button onClick={()=>setMobileMenu(false)} className="text-zinc-500 hover:text-zinc-200 p-1"><X size={16}/></button>
            </div>
            <div className="grid grid-cols-2 gap-2 p-3">
              {NAV.map(n=>(
                <button key={n.id} onClick={()=>go(n.id)} className={`flex items-center gap-3 rounded-xl px-3 py-3 text-sm font-medium transition-colors ${page===n.id?"bg-emerald-500/10 text-emerald-400 border border-emerald-500/20":"text-zinc-400 hover:bg-zinc-800"}`}>
                  <n.icon size={15} className="flex-shrink-0"/>
                  <span className="truncate">{n.label}</span>
                  {n.id==="importacao"&&<span className="ml-auto text-[9px] bg-violet-500/20 text-violet-400 rounded-full px-1.5 py-0.5 font-bold">IA</span>}
                </button>
              ))}
            </div>
            <div className={`flex items-center justify-between px-4 py-3 border-t ${dark?"border-zinc-800":"border-zinc-200"}`}>
              <span className="text-xs text-zinc-500">Aparência</span>
              <button onClick={()=>setDark(!dark)} className={`flex items-center gap-2 text-xs rounded-lg px-3 py-1.5 ${dark?"bg-zinc-800 text-zinc-300":"bg-zinc-100 text-zinc-600"}`}>{dark?<Sun size={12}/>:<Moon size={12}/>}{dark?"Modo Claro":"Modo Escuro"}</button>
            </div>
          </div>
        </div>
      )}

      {/* Main */}
      <div className="flex-1 flex flex-col min-w-0" style={{overflow:"hidden"}}>
        {/* Topbar */}
        <header className={`flex-shrink-0 h-14 border-b flex items-center justify-between px-4 ${dark?"bg-zinc-950/90 border-zinc-800":"bg-white border-zinc-200"}`}>
          <div className="flex items-center gap-2 min-w-0">
            <div className="md:hidden flex items-center gap-2 flex-shrink-0">
              <div className="w-7 h-7 rounded-lg bg-emerald-500 flex items-center justify-center"><CircleDollarSign size={13} className="text-zinc-900"/></div>
            </div>
            <div className="hidden md:flex items-center gap-2 text-sm">
              <span className="text-zinc-600">ERP</span>
              <ChevronRight size={12} className="text-zinc-700"/>
            </div>
            <span className="font-semibold text-sm truncate">{cur?.label}</span>
          </div>
          <div className="flex items-center gap-2 flex-shrink-0">
            {fin.pending.length>0&&<button onClick={()=>setPage("fluxo")} className="relative p-2 rounded-lg text-zinc-400 hover:bg-zinc-800 transition-colors"><Bell size={15}/><span className="absolute top-1.5 right-1.5 w-1.5 h-1.5 rounded-full bg-rose-500"/></button>}
            <div className={`hidden md:flex items-center gap-2 rounded-lg px-3 py-1.5 border text-xs ${dark?"bg-zinc-800/50 border-zinc-700/40 text-zinc-400":"bg-zinc-50 border-zinc-200 text-zinc-600"}`}>
              <div className="w-5 h-5 rounded-full bg-emerald-500 flex items-center justify-center text-[10px] font-bold text-zinc-900">T</div>
              <span>Tia Carol</span>
            </div>
            <button onClick={()=>setShowSettings(true)} className="p-2 rounded-lg text-zinc-400 hover:bg-zinc-800 transition-colors" title="Configurações IA">
              <SettingsIcon size={15}/>
            </button>
            <button onClick={()=>setMobileMenu(true)} className="md:hidden p-2 rounded-lg text-zinc-400 hover:bg-zinc-800 transition-colors">
              <svg width="18" height="18" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" d="M4 6h16M4 12h16M4 18h16"/></svg>
            </button>
          </div>
        </header>

        {/* Content */}
        <main className="flex-1 overflow-y-auto p-4 md:p-5 lg:p-6 pb-20 md:pb-6">
          <div className="max-w-6xl mx-auto">{PAGES[page]||PAGES.dashboard}</div>
        </main>
      </div>

      {/* Bottom nav mobile */}
      <nav className={`md:hidden fixed bottom-0 left-0 right-0 z-40 border-t flex ${dark?"bg-zinc-900/95 border-zinc-800":"bg-white border-zinc-200"}`} style={{paddingBottom:"env(safe-area-inset-bottom)"}}>
        {NAV.slice(0,4).map(n=>(
          <button key={n.id} onClick={()=>setPage(n.id)} className={`flex-1 flex flex-col items-center justify-center gap-0.5 py-2.5 transition-colors ${page===n.id?"text-emerald-400":"text-zinc-600"}`}>
            <n.icon size={18} className={page===n.id?"text-emerald-400":"text-zinc-500"}/>
            <span className="text-[9px] font-medium leading-none">{n.label.split(" ")[0]}</span>
          </button>
        ))}
        <button onClick={()=>setMobileMenu(true)} className={`flex-1 flex flex-col items-center justify-center gap-0.5 py-2.5 transition-colors ${NAV.slice(4).some(n=>n.id===page)?"text-emerald-400":"text-zinc-600"}`}>
          <svg width="18" height="18" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" d="M4 6h16M4 12h16M4 18h16"/></svg>
          <span className="text-[9px] font-medium leading-none">Mais</span>
        </button>
      </nav>
    </div>
  )
}
