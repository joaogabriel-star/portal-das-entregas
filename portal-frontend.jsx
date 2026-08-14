import React, { useState, useMemo, useRef, useEffect, Fragment } from "react";
import * as d3 from "d3";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import {
  Search, Plus, X, Bot, Send, ChevronRight, ChevronDown, Check, Trash2,
  Sparkles, ClipboardList, Building2, LayoutGrid, GitBranch, PieChart,
  Info, Users, Copy, ArrowRight, FileSpreadsheet, ShieldCheck, Wand2,
  Sun, Flag, AlertTriangle, Upload, RefreshCw, Paperclip, FileText, Download, Pencil, Network,
  Menu, Columns, Home, Workflow, Target, ChevronLeft,
  GitMerge, Layers, ClipboardCheck, FilePlus2, RotateCcw, Trophy,
  TrendingDown, TrendingUp, Minus, BarChart3, Clock, CircleDot, ExternalLink, ZoomIn, ZoomOut, Maximize2,
  AlertCircle, ArrowLeft, BarChart2, CheckCircle2, MapPin, Minimize2, PlusCircle,
} from "lucide-react";

/* ---------- leitura de arquivos no navegador (planilha / pdf / docx) ---------- */
// retorna { tipo, nome, texto?, pdfB64? } para alimentar o conversor por IA
async function lerArquivo(file){
  const nome=file.name, ext=(nome.split(".").pop()||"").toLowerCase();
  if(ext==="xlsx"||ext==="xls"||ext==="csv"){
    const buf=await file.arrayBuffer();
    const wb=XLSX.read(buf,{type:"array"});
    let texto="";
    wb.SheetNames.forEach(sn=>{
      const linhas=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,blankrows:false});
      texto+=`# Planilha: ${sn}\n`+linhas.map(l=>l.join(" | ")).join("\n")+"\n\n";
    });
    return {tipo:"planilha",nome,texto:texto.slice(0,60000)};
  }
  if(ext==="docx"){
    const buf=await file.arrayBuffer();
    const {value}=await mammoth.extractRawText({arrayBuffer:buf});
    return {tipo:"documento",nome,texto:(value||"").slice(0,60000)};
  }
  if(ext==="pdf"){
    const b64=await new Promise((res,rej)=>{ const r=new FileReader(); r.onload=()=>res(String(r.result).split(",")[1]); r.onerror=rej; r.readAsDataURL(file); });
    return {tipo:"pdf",nome,pdfB64:b64};
  }
  throw new Error("Formato não suportado: ."+ext);
}

/* ---------- exportar Descrição de Área como .xlsx ---------- */
function exportarDescricaoXLSX(sel,notes,orgao,unidade){
  const NATROT={finalistico:"Finalístico",governanca:"Governança",suporte:"Suporte"};
  const cab=[["Descrição de Área — DFT/SISDIP"],["Órgão:",orgao||""],["Unidade:",unidade||""],
    ["Gerado em:",new Date().toLocaleString("pt-BR")],["Total de entregas:",sel.length],[]];
  const colunas=["#","Código","Entrega","Atividade","Serviço","Categoria de Serviço","Macroprocesso","Natureza","Metodologia","Situação","Observações da unidade"];
  const linhas=sel.map((e,i)=>[i+1,e.nova?"(nova)":e.codigo,e.entrega,e.atividade||"",e.servico||"",e.categoria||"",e.macro||"",e.nova?"—":(NATROT[e.natureza]||e.natureza),e.metod||(e.nova?"—":"Típica"),e.nova?"Nova — aguarda curadoria":"No catálogo",notes[e.codigo]||""]);
  const aoa=[...cab,colunas,...linhas];
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"]=[{wch:5},{wch:12},{wch:48},{wch:48},{wch:28},{wch:28},{wch:26},{wch:14},{wch:12},{wch:22},{wch:36}];
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"Descrição de Área");
  const nomeArq=`descricao-area-${(unidade||"unidade").replace(/[^a-z0-9]+/gi,"-").toLowerCase().slice(0,40)}.xlsx`;
  XLSX.writeFile(wb,nomeArq);
}


/* ---------- tokens gov.br / SIGEPE ---------- */
const C = {
  bg:"#F4F6F9", surface:"#FFFFFF", ink:"#1B1B1B", sub:"#55606E", faint:"#8C97A6",
  line:"#E4E7EB", primary:"#1351B4", primaryDark:"#0C326F", primarySoft:"#E8EEF9",
  green:"#168821", greenSoft:"#E3F2E5", navy:"#13315C", yellow:"#FFCD07",
};
const NAT = {
  finalistico:{ rot:"Finalístico", cor:"#168821", soft:"#E3F2E5" },
  governanca:{ rot:"Governança", cor:"#B86E00", soft:"#FBEEDB" },
  suporte:{ rot:"Suporte", cor:"#1351B4", soft:"#E8EEF9" },
};
const SIT = {
  ativa:{ rot:"Ativa", cor:"#168821", soft:"#E3F2E5" },
  proposta:{ rot:"Proposta", cor:"#B86E00", soft:"#FBEEDB" },
  ncat:{ rot:"Não categorizada", cor:"#8C97A6", soft:"#EEF1F4" },
};

/* ---------- dados demo enriquecidos ---------- */
// ── Banco real: 31 241 entregas · 216 macros · 1 311 categorias ───────
// Exibidos: 6 388 (cap 8/categoria). Cols A e B (Órgão/Unidade) nos detalhes.

// amostra mínima usada só se banco.json não estiver disponível (ex.: preview sem deploy)
const FALLBACK_BANCO=[
  {c:"06050016",e:"Entrevista de acolhimento e/ou atendimento a estudantes",a:"Realização de entrevista de acolhimento",s:"",k:"Acessibilidade e inclusão",m:"Educação superior",n:"finalistico",o:"FUFT",u:"FUFT",t:"Típica"},
  {c:"06050009",e:"Adequação pedagógica para Pessoas com Deficiência (PCD)",a:"Atualização da adequação pedagógica",s:"",k:"Acessibilidade e inclusão",m:"Educação superior",n:"finalistico",o:"UFS",u:"UFS",t:"Típica"},
  {c:"03050008",e:"Licença ambiental analisada",a:"Análise de requerimento de licença",s:"Licenciamento ambiental",k:"Licenciamento ambiental",m:"Gestão ambiental",n:"finalistico",o:"IBAMA",u:"IBAMA",t:"Típica"},
  {c:"01050266",e:"Folha de pagamento processada",a:"Processamento da folha",s:"Gestão de pessoas",k:"Gestão de pessoas",m:"Gestão de pessoas",n:"suporte",o:"MGI",u:"MGI",t:"Típica"},
  {c:"12620001",e:"Minuta normativa elaborada",a:"Elaboração de minuta",s:"Governança",k:"Normas e regulamentação",m:"Governança institucional",n:"governanca",o:"MGI",u:"MGI",t:"Típica"}
];

// ── Banco de entregas: carregado de banco.json em produção (31.241 entregas) ──
// Estruturas mutáveis populadas por loadBanco(). No preview do Claude, se o fetch
// falhar (arquivo ausente), cai no fallback embutido reduzido (FALLBACK_BANCO).
let ENTREGAS = [];           // lista plana de entregas (todas as colunas de navegação)
let BANCO_RAW = [];          // hierarquia nat→macro→cat→itens (modos processo/cadeia)
let CATEGORIAS = [];
let CAT_TXT = "";
let DET_CACHE = null;        // detalhes (fluxo/fonte) carregados sob demanda

const NATMAP={"Finalístico":"finalistico","Governança":"governanca","Suporte":"suporte"};

// monta ENTREGAS + BANCO_RAW + derivados a partir da lista plana "lite"
function ingestBanco(lite){
  ENTREGAS = lite.map(x=>({
    codigo:x.c, entrega:x.e, atividade:x.a,
    servico:x.s||x.k, categoria:x.k, macro:x.m, natureza:x.n,
    orgao:x.o, unidade:x.u, metod:x.t,
    sit: x.t==='Outra'?'proposta':'ativa',
  }));
  // hierarquia + contagem real por categoria
  const natOrder=["finalistico","governanca","suporte"];
  const byNat=new Map();
  ENTREGAS.forEach(e=>{
    if(!byNat.has(e.natureza)) byNat.set(e.natureza,new Map());
    const mM=byNat.get(e.natureza);
    if(!mM.has(e.macro)) mM.set(e.macro,new Map());
    const cM=mM.get(e.macro);
    if(!cM.has(e.categoria)) cM.set(e.categoria,[]);
    cM.get(e.categoria).push(e);
  });
  BANCO_RAW = natOrder.filter(n=>byNat.has(n)).map(nat=>({
    nat, macros:[...byNat.get(nat)].map(([mac,cM])=>({
      mac, cats:[...cM].map(([cat,itens])=>({
        cat, n:itens.length,
        itens:itens.map(e=>({cod:e.codigo,ent:e.entrega,ativ:e.atividade,serv:e.servico,metod:e.metod}))
      }))
    }))
  }));
  CATEGORIAS=[...new Set(ENTREGAS.map(e=>e.categoria))];
  // amostra p/ contexto da IA (uma por categoria, limitada — evita prompt gigante)
  const seen=new Set(); const sample=[];
  for(const e of ENTREGAS){ if(!seen.has(e.categoria)){ seen.add(e.categoria); sample.push(e); } if(sample.length>=700) break; }
  CAT_TXT=sample.map(e=>`${e.codigo} | ${e.entrega} | ${e.macro} > ${e.categoria}`).join("\n");
}

const CODIGO_RX = /\b\d{4}\.\d{4}\b/g;

// ── Enquadramento em 2 fases: (1) retrieval local no banco completo, custo zero;
//    (2) só as candidatas mais relevantes vão para a IA decidir. Barato + preciso. ──
const STOPWORDS=new Set(["de","da","do","das","dos","e","a","o","as","os","em","no","na","nos","nas","para","por","com","que","um","uma","ao","à","às","aos","se","the","of","and","pela","pelo","sobre","como","dos","seu","sua","ou","ser","the"]);
// stemming leve PT-BR: remove plural e sufixos flexionais comuns p/ casar "licenciamento"~"licenças"
function stem(w){
  w=w.replace(/(ções|ção|çao)$/,"ca").replace(/(mentos|mento)$/,"").replace(/(dores|dora|dor)$/,"").replace(/(agem|agens)$/,"ag");
  w=w.replace(/(izações|ização)$/,"iza").replace(/(icas|ico|ica)$/,"ic");
  if(w.length>5) w=w.replace(/(ais|eis|is|ns|s)$/,""); // plurais
  if(w.length>5) w=w.replace(/(ando|endo|indo|ada|ado|adas|ados|idas|idos)$/,""); // verbais
  return w;
}
function tokenize(s){ return norm(s).replace(/[^a-z0-9\s]/g," ").split(/\s+/).filter(w=>w.length>2&&!STOPWORDS.has(w)).map(stem); }

let _IDF=null; // peso inverso de frequência: termos raros valem mais
function construirIDF(){
  if(_IDF) return _IDF; _IDF=new Map();
  const df=new Map(); const N=ENTREGAS.length||1;
  for(const e of ENTREGAS){
    const toks=e._tokset||(e._tokset=new Set(tokenize(e.entrega+" "+e.atividade+" "+e.servico+" "+e.categoria)));
    for(const w of toks) df.set(w,(df.get(w)||0)+1);
  }
  for(const [w,c] of df) _IDF.set(w, Math.log(1+N/c));
  return _IDF;
}

// pontua cada entrega por similaridade ponderada (campo + IDF) e retorna as top-K
function buscarCandidatas(texto, k=60){
  const qt=tokenize(texto); if(!qt.length) return [];
  const idf=construirIDF();
  const qw=new Map(); qt.forEach(w=>qw.set(w,(qw.get(w)||0)+1));
  const scored=[];
  for(const e of ENTREGAS){
    // campos com pesos: entrega(3) > serviço(2) > categoria(1.5) > macro(1) > atividade(1)
    const campos=e._campos||(e._campos=[
      [new Set(tokenize(e.entrega)),3.0],
      [new Set(tokenize(e.servico||"")),2.0],
      [new Set(tokenize(e.categoria)),1.5],
      [new Set(tokenize(e.macro)),1.0],
      [new Set(tokenize(e.atividade||"")),0.8],
    ]);
    let s=0, hits=0;
    for(const [w,qc] of qw){
      const peso=idf.get(w)||1;
      for(const [cset,cw] of campos){ if(cset.has(w)){ s+=qc*peso*cw; hits++; break; } }
    }
    if(hits>0){ s*=(1+hits/qw.size); scored.push([s,e]); } // bônus por cobrir mais termos da consulta
  }
  scored.sort((a,b)=>b[0]-a[0]);
  return scored.slice(0,k).map(x=>x[1]);
}
function candidatasTxt(list){ return list.map(e=>`${e.codigo} | ${e.entrega}${e.servico?" | serviço: "+e.servico:""} | ${e.macro} > ${e.categoria}`).join("\n"); }

// endpoints: em produção, a Netlify Function protege a chave; no preview do Claude,
// a chamada direta à API é autenticada pelo ambiente. Tentamos a função e, se ela
// não existir (preview / sem backend), caímos para a API direta.
const AI_FUNCTION = "/.netlify/functions/claude";
const AI_DIRECT = "https://api.anthropic.com/v1/messages";
let _aiEndpoint = null; // memoiza qual funcionou
async function chamarIA(payload){
  const body=JSON.stringify(payload);
  const tentar=async(url)=>{ const r=await fetch(url,{method:"POST",headers:{"Content-Type":"application/json"},body}); if(!r.ok) throw new Error("HTTP "+r.status); return r.json(); };
  if(_aiEndpoint) return tentar(_aiEndpoint);
  try{ const d=await tentar(AI_FUNCTION); _aiEndpoint=AI_FUNCTION; return d; }
  catch(e1){ const d=await tentar(AI_DIRECT); _aiEndpoint=AI_DIRECT; return d; }
}



// detalhes (fluxo/fonte) sob demanda
async function ensureDetalhes(){
  if(DET_CACHE) return DET_CACHE;
  try{ const r=await fetch("banco-detalhes.json"); DET_CACHE=await r.json(); }
  catch{ DET_CACHE={}; }
  return DET_CACHE;
}
function getDetails(cod){ const d=(DET_CACHE&&DET_CACHE[cod])||{}; return {fluxo:d.flu||'',fonte:d.fon||''}; }

// carrega o banco principal; resolve com nº de entregas (0 = falhou)
async function loadBanco(){
  try{
    const r=await fetch("banco.json");
    if(!r.ok) throw new Error("HTTP "+r.status);
    const lite=await r.json();
    ingestBanco(lite);
    return ENTREGAS.length;
  }catch(err){
    ingestBanco(FALLBACK_BANCO);   // preview sem arquivo: usa amostra mínima
    return 0;
  }
}

const norm=s=>(s||"").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g,"");
const NAT_ORDER=["finalistico","governanca","suporte"];

/* ---------- órgão/unidade lembrados no navegador ----------
   O cadastro é a etapa 1 do caminho: uma vez informado, fica salvo neste
   computador e não é perguntado de novo (dá para trocar pelo cabeçalho). */
const CHAVE_UNIDADE="portal:unidade";
function lerUnidadeSalva(){
  try{
    const cru=window.localStorage.getItem(CHAVE_UNIDADE);
    if(!cru) return {orgao:"",unidade:""};
    const d=JSON.parse(cru);
    return {orgao:d.orgao||"",unidade:d.unidade||""};
  }catch{ return {orgao:"",unidade:""}; }
}
function salvarUnidade(orgao,unidade){
  try{
    if(!orgao && !unidade) window.localStorage.removeItem(CHAVE_UNIDADE);
    else window.localStorage.setItem(CHAVE_UNIDADE,JSON.stringify({orgao,unidade}));
  }catch{ /* navegador sem localStorage (aba anônima, etc.) — segue sem lembrar */ }
}

/* ---------- objetivos (editáveis no portal; futuramente vêm do PGD) ---------- */
const SEED_OBJ_ORGAO=[];  // começa vazio — a unidade inclui os objetivos do órgão
let _orgSeq=1;
const nextOrgId=()=>{ const id=String.fromCharCode(64+_orgSeq); _orgSeq++; return _orgSeq>27?`O${_orgSeq}`:id; };
const SEED_OBJ_UNIDADE=[];  // começa vazio
let _objSeq=1;
const nextObjId=()=>`u${_objSeq++}`;

/* ---------- tooltip reutilizável (hover) ---------- */
function Tip({text,children,pos="bottom",wide}){
  return (<span className="px-tipw">{children}<span className={`px-tip ${pos} ${wide?"wide":""}`} role="tooltip">{text}</span></span>);
}

export default function PortalEntregas(){
  const [query,setQuery]=useState("");
  const [dquery,setDquery]=useState("");        // busca com debounce
  const [natF,setNatF]=useState(null);
  const [macroF,setMacroF]=useState(null);
  const [catF,setCatF]=useState(null);
  const [servF,setServF]=useState(null);
  const [orgF,setOrgF]=useState(null);
  const [mode,setMode]=useState("lista");
  const [exp,setExp]=useState(null);
  const [sel,setSel]=useState([]);
  const [notes,setNotes]=useState({});          // observações por código → DFT
  // Órgão/unidade ficam lembrados no navegador: quem já cadastrou uma vez não
  // precisa redigitar a cada visita (troca pelo chip do cabeçalho).
  const [orgao,setOrgao]=useState(()=>lerUnidadeSalva().orgao);
  const [unidade,setUnidade]=useState(()=>lerUnidadeSalva().unidade);
  const [preview,setPreview]=useState(false);
  const [toast,setToast]=useState(null);
  /* Conversor agora vive dentro do Início (segunda porta), não é modal nem modo próprio */
  const [flags,setFlags]=useState([]);
  const [flagFor,setFlagFor]=useState(null);
  const [novaFor,setNovaFor]=useState(null);     // propor nova entrega (texto)
  const [onboard,setOnboard]=useState(false);    // boas-vindas antiga (desativada — Início assume)
  const [secao,setSecao]=useState("inicio");     // navegação principal: inicio · catalogo
  const [navPanel,setNavPanel]=useState(null);   // painel lateral de filtro: null · macro · categoria · servico · orgao
  const [escopo,setEscopo]=useState("sel");      // níveis: "sel" (o que a IA achou) · "tudo" (catálogo inteiro)
  const [vistaEntrega,setVistaEntrega]=useState("lista"); // nível 4: "lista" · "arvore"
  const [trocarUnidade,setTrocarUnidade]=useState(false); // modal de cadastro/troca da unidade
  const [descDrawer,setDescDrawer]=useState(false);
  const [descView,setDescView]=useState("lista");
  const [objUnidade,setObjUnidade]=useState(SEED_OBJ_UNIDADE);
  const [objOrgao,setObjOrgao]=useState(SEED_OBJ_ORGAO);
  // Estado de Marcos Referenciais fica aqui (não dentro do componente) para
  // não se perder quando o usuário troca de aba e volta — os arquivos
  // anexados e a resposta da IA continuam visíveis.
  const [marcosArqs,setMarcosArqs]=useState({cadeia:null,estrutura:null,rgi:null,regimento:null});
  const [marcosResposta,setMarcosResposta]=useState(null);
  const [marcosQtd,setMarcosQtd]=useState(0);
  const [marcosErro,setMarcosErro]=useState("");
  const [compare,setCompare]=useState([]);       // até 3 p/ comparar
  const [compareOpen,setCompareOpen]=useState(false);
  const [justAdded,setJustAdded]=useState(null); // micro-confirmação inline
  const [loading,setLoading]=useState(true);     // carregando banco
  const [bancoStatus,setBancoStatus]=useState("carregando"); // carregando|ok|fallback
  const [bancoN,setBancoN]=useState(0);
  /* Central de Revisão e Organizações agora são seções (secao), não overlays */
  const [toolsOpen,setToolsOpen]=useState(false);         // menu Ferramentas no cabeçalho
  const [descCollapsed,setDescCollapsed]=useState(false); // painel lateral recolhido (desktop)

  const selSet=useMemo(()=>new Set(sel.map(e=>e.codigo)),[sel]);
  const flash=m=>{setToast(m);setTimeout(()=>setToast(null),2400);};
  const add=e=>{ if(!selSet.has(e.codigo)){ setSel(s=>[...s,e]); setJustAdded(e.codigo); setTimeout(()=>setJustAdded(c=>c===e.codigo?null:c),1600); } };
  const rem=c=>{ setSel(s=>s.filter(e=>e.codigo!==c)); setNotes(o=>{const n={...o};delete n[c];return n;}); };
  const setNote=(c,v)=>setNotes(o=>({...o,[c]:v}));
  const toggleCompare=e=>setCompare(c=>{ if(c.find(x=>x.codigo===e.codigo)) return c.filter(x=>x.codigo!==e.codigo); if(c.length>=3){flash("Máximo de 3 entregas na comparação.");return c;} return [...c,e]; });

  // debounce da busca (≈300ms) — evita reconstruir gráficos a cada tecla
  useEffect(()=>{ const t=setTimeout(()=>setDquery(query),300); return ()=>clearTimeout(t); },[query]);
  // grava o cadastro da unidade neste navegador a cada mudança
  useEffect(()=>{ salvarUnidade(orgao,unidade); },[orgao,unidade]);
  // carga inicial real do banco de entregas (banco.json em produção)
  useEffect(()=>{ let vivo=true;
    loadBanco().then(n=>{ if(!vivo) return;
      setBancoN(ENTREGAS.length);
      setBancoStatus(n>0?"ok":"fallback");
      setLoading(false);
    });
    return ()=>{vivo=false;};
  },[]);

  const chooseNat=id=>{ setNatF(id); setMacroF(null); setCatF(null); setServF(null); setOrgF(null); };
  const pickMacro=(nat,mac)=>{ setNatF(nat); setMacroF(mac); setCatF(null); setServF(null); setOrgF(null); setMode("lista"); };
  const pickCategoria=cat=>{ setCatF(cat); setServF(null); setMode("lista"); };
  const pickServico=serv=>{ setServF(serv); setMode("lista"); };
  const pickOrgao=org=>{ setOrgF(org); setMode("lista"); };

  const macroOpts=useMemo(()=>{ const s=new Set(); ENTREGAS.forEach(e=>{ if(!natF||e.natureza===natF) s.add(e.macro); }); return [...s].sort((a,b)=>a.localeCompare(b)); },[natF,bancoN]);
  const catOpts=useMemo(()=>{ const s=new Set(); ENTREGAS.forEach(e=>{ if((!natF||e.natureza===natF)&&(!macroF||e.macro===macroF)) s.add(e.categoria); }); return [...s].sort((a,b)=>a.localeCompare(b)); },[natF,macroF,bancoN]);
  // opções para os painéis de filtro por categoria/serviço/órgão (mesmo padrão do MacroNav)
  const categoriaOpts=useMemo(()=>{ const m=new Map(); ENTREGAS.forEach(e=>{ if((!natF||e.natureza===natF)&&(!macroF||e.macro===macroF)) m.set(e.categoria,(m.get(e.categoria)||0)+1); }); return [...m.entries()].sort((a,b)=>b[1]-a[1]).map(([valor,n])=>({valor,n})); },[natF,macroF,bancoN]);
  const servicoOpts=useMemo(()=>{ const m=new Map(); ENTREGAS.forEach(e=>{ if(e.servico&&(!natF||e.natureza===natF)&&(!macroF||e.macro===macroF)&&(!catF||e.categoria===catF)) m.set(e.servico,(m.get(e.servico)||0)+1); }); return [...m.entries()].sort((a,b)=>b[1]-a[1]).map(([valor,n])=>({valor,n})); },[natF,macroF,catF,bancoN]);
  const orgaoOpts=useMemo(()=>{ const m=new Map(); ENTREGAS.forEach(e=>{ if(e.orgao) m.set(e.orgao,(m.get(e.orgao)||0)+1); }); return [...m.entries()].sort((a,b)=>b[1]-a[1]).map(([valor,n])=>({valor,n})); },[bancoN]);
  const naturezaOpts=useMemo(()=>NAT_ORDER.map(id=>({valor:id,label:NAT[id].rot,n:ENTREGAS.filter(e=>e.natureza===id).length})),[bancoN]);

  // números do banco para o painel do topo — recalculados quando o banco carrega
  const stats=useMemo(()=>{
    const macros=new Set(),cats=new Set(),servs=new Set(),orgs=new Set();
    const dist={finalistico:0,governanca:0,suporte:0}; let comServ=0;
    ENTREGAS.forEach(e=>{ macros.add(e.macro); cats.add(e.categoria);
      if(e.servico){ servs.add(e.servico); comServ++; } if(e.orgao) orgs.add(e.orgao);
      if(dist[e.natureza]!=null) dist[e.natureza]++; });
    const total=ENTREGAS.length;
    return { total, macros:macros.size, cats:cats.size, servs:servs.size, orgs:orgs.size,
      pctServ: total?Math.round(comServ/total*100):0, dist };
  },[bancoN]);
  // abre a Descrição da Área respeitando o tamanho de tela: gaveta no compacto, painel no amplo
  const openDesc=()=>{ const compacto = typeof window!=="undefined" && window.matchMedia && window.matchMedia("(max-width:1024px)").matches;
    if(compacto) setDescDrawer(true); else setDescCollapsed(false); };

  const res=useMemo(()=>{
    let r=ENTREGAS;
    if(natF) r=r.filter(e=>e.natureza===natF);
    if(macroF) r=r.filter(e=>e.macro===macroF);
    if(catF) r=r.filter(e=>e.categoria===catF);
    if(servF) r=r.filter(e=>e.servico===servF);
    if(orgF) r=r.filter(e=>e.orgao===orgF);
    if(dquery.trim()){ const q=norm(dquery); r=r.filter(e=>norm([e.codigo,e.entrega,e.atividade,e.servico,e.categoria,e.macro].join(" ")).includes(q)); }
    return r;
  },[dquery,natF,macroF,catF,servF,orgF,bancoN]);

  const total=ENTREGAS.length;
  const searching=!!(dquery.trim()||natF||macroF||catF||servF||orgF);

  const docHeader=(<DocHeaderEdit orgao={orgao} unidade={unidade} setOrgao={setOrgao} setUnidade={setUnidade}/>);

  /* ---- encadeamento dos 4 níveis do catálogo ----
     Descer leva o filtro do nível junto; a trilha no topo sobe de volta.
     Reaproveita os filtros que já existiam (macroF / catF / servF), então a
     Lista e a barra de filtros ativos continuam funcionando sem mudança. */
  const descerNivel=(nivel,valor,semServico)=>{
    if(nivel==="macro"){ setMacroF(valor||null); setCatF(null); setServF(null); setMode("processo"); return; }
    if(nivel==="processo"){ setCatF(valor||null); setServF(null); setMode("servico"); return; }
    // no nível Serviço, o grupo "sem serviço específico" não tem valor próprio
    // que sirva de filtro quando não há processo escolhido — nesse caso desce
    // sem fixar o serviço, mostrando as entregas do recorte atual.
    setServF(semServico && !catF ? null : (valor||null));
    setMode("lista");
  };
  const irNivel=id=>{
    if(id==="macro"){ setMacroF(null); setCatF(null); setServF(null); }
    if(id==="processo"){ setCatF(null); setServF(null); }
    if(id==="servico"){ setServF(null); }
    setMode(id);
  };
  const limparTrilha=()=>{ setMacroF(null); setCatF(null); setServF(null); };

  /* ---- caminho de 4 etapas (navegação principal) ---- */
  const etapaAtual = secao!=="catalogo" ? 0 : (mode==="marcos" ? 1 : 2);
  // "feito" só quando há sinal concreto; as etapas 3 e 4 vivem em páginas
  // separadas, então não têm como se marcar sozinhas aqui.
  const etapasFeitas = { 1: !!(orgao||unidade), 2: sel.length>0 };
  const irParaEtapa = n => {
    if(n===1){ setSecao("catalogo"); setMode("marcos"); return; }
    if(n===2){ setSecao("catalogo"); setMode("macro"); return; }
    if(typeof window==="undefined") return;
    window.location.href = n===3 ? "./jornada-gestor-unidade.html" : "./organizacoes.html";
  };

  return (
    <div className="px-app" style={{background:C.bg,minHeight:"100%",color:C.ink,fontFamily:"'Raleway','Segoe UI',system-ui,sans-serif"}}>
      <style>{css}</style>
      <div className="px-stripe"><span style={{background:C.green}}/><span style={{background:C.yellow}}/><span style={{background:C.primary}}/></div>

      <header className="px-head">
        <div className="px-logo" onClick={()=>setSecao("inicio")} style={{cursor:"pointer"}} title="Ir para o Início">
          <svg viewBox="0 0 32 24" width="32" height="25"><circle cx="7" cy="7" r="4" fill={C.yellow}/><circle cx="16" cy="6" r="4.4" fill={C.green}/><circle cx="25" cy="7" r="4" fill={C.primary}/><path d="M2 22c1.5-5 9-5 10.5 0z" fill={C.yellow}/><path d="M10.5 22c1.5-6 9.5-6 11 0z" fill={C.green}/><path d="M20 22c1.5-5 9-5 10.5 0z" fill={C.primary}/></svg>
          <div><div className="px-logo-w">Catálogo de Serviços</div><div className="px-logo-s">SIGEPE · SISDIP / DFT</div></div>
        </div>
        <nav className="px-nav" aria-label="Seções do portal">
          <button className={`px-nav-item ${secao==="inicio"?"on":""}`}
            onClick={()=>setSecao("inicio")}><Home size={14}/><span>Início</span></button>
        </nav>
        <div className="px-head-r">
          <button className="px-org" onClick={()=>setTrocarUnidade(true)}
            title={orgao||unidade?"Trocar órgão e unidade":"Cadastrar órgão e unidade"}>
            <Building2 size={13}/> {unidade||orgao||"Cadastrar unidade"} <Pencil size={11}/>
          </button>
        </div>
      </header>

      {/* faixa das 4 etapas — navegação principal do portal */}
      <FaixaEtapas etapa={etapaAtual} completo={etapasFeitas} expandida={secao==="inicio"} ir={irParaEtapa}/>

      {trocarUnidade && <ModalUnidade orgao={orgao} unidade={unidade}
        setOrgao={setOrgao} setUnidade={setUnidade} onClose={()=>setTrocarUnidade(false)}/>}

      {!loading && secao==="inicio" && <PainelNumeros stats={stats} bancoStatus={bancoStatus}/>}

      {secao==="inicio" && !loading && <SecaoInicioCaminho stats={stats}
        orgao={orgao} unidade={unidade} completo={etapasFeitas} ir={irParaEtapa}/>}
      {secao==="inicio" && loading && <BancoLoading/>}

      {/* Importar PGD é uma visualização da barra (mode==="pgd"); o Conversor
          passou a viver dentro do Início, como a segunda porta de entrada. */}

      {secao==="catalogo" && <div className="px-wrap">
        <div className="px-toolbar">
        <div className="px-controls">
          <button className={`px-navtoggle ${navPanel==="macro"?"on":""}`} onClick={()=>setNavPanel(p=>p==="macro"?null:"macro")} title="Navegar por macroprocesso"><Menu size={17}/></button>
          <div className="px-search full">
            <Search size={18} color={C.faint}/>
            <input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Buscar por entrega, atividade, serviço ou código…"/>
            {query && <button className="px-clear" onClick={()=>setQuery("")}><X size={15}/></button>}
          </div>
          <span className={`px-count ${searching?"filt":""}`}>{searching?`${res.length.toLocaleString("pt-BR")} de ${total.toLocaleString("pt-BR")}`:`${total.toLocaleString("pt-BR")} entregas`}</span>
        </div>

        <PainelNumeros stats={stats} bancoStatus={bancoStatus} filtro={{
          onTotal: ()=>{setNavPanel(null);setNatF(null);setMacroF(null);setCatF(null);setServF(null);setOrgF(null);},
          ativoNatureza: navPanel==="natureza"||!!natF,
          onNatureza: ()=>setNavPanel(p=>p==="natureza"?null:"natureza"),
          ativoMacro: navPanel==="macro"||!!macroF,
          onMacro: ()=>setNavPanel(p=>p==="macro"?null:"macro"),
          ativoCategoria: navPanel==="categoria"||!!catF,
          onCategoria: ()=>setNavPanel(p=>p==="categoria"?null:"categoria"),
          ativoServico: navPanel==="servico"||!!servF,
          onServico: ()=>setNavPanel(p=>p==="servico"?null:"servico"),
          ativoOrgao: navPanel==="orgao"||!!orgF,
          onOrgao: ()=>setNavPanel(p=>p==="orgao"?null:"orgao"),
        }}/>
        {(natF||macroF||catF||servF||orgF||dquery) && (
          <div className="px-filtros-ativos">
            <div className="px-fa-tags">
              {dquery && <span className="px-fa-tag">"{dquery.length>22?dquery.slice(0,21)+"…":dquery}"<button onClick={()=>setQuery("")}><X size={9}/></button></span>}
              {natF && NAT[natF] && <span className="px-fa-tag" style={{background:NAT[natF].soft,color:NAT[natF].cor,borderColor:NAT[natF].cor}}>{NAT[natF].rot}<button onClick={()=>chooseNat(null)} style={{color:NAT[natF].cor}}><X size={9}/></button></span>}
              {macroF && <span className="px-fa-tag">{macroF.length>38?macroF.slice(0,37)+"…":macroF}<button onClick={()=>setMacroF(null)}><X size={9}/></button></span>}
              {catF && <span className="px-fa-tag">{catF.length>38?catF.slice(0,37)+"…":catF}<button onClick={()=>setCatF(null)}><X size={9}/></button></span>}
              {servF && <span className="px-fa-tag">{servF.length>38?servF.slice(0,37)+"…":servF}<button onClick={()=>setServF(null)}><X size={9}/></button></span>}
              {orgF && <span className="px-fa-tag">{orgF.length>38?orgF.slice(0,37)+"…":orgF}<button onClick={()=>setOrgF(null)}><X size={9}/></button></span>}
            </div>
            <button className="px-fa-limpar" onClick={()=>{setNatF(null);setMacroF(null);setCatF(null);setServF(null);setOrgF(null);setQuery("");}}>Limpar tudo</button>
          </div>
        )}

        <div className="px-modes-row">
          <div className="px-modes">
            {/* os 5 passos do trabalho: começar → descer os níveis → chegar na entrega */}
            <button className={`px-mode ${mode==="marcos"?"on":""}`} onClick={()=>setMode("marcos")} title="Início — traga os documentos da unidade"><Home size={15}/> <span className="px-mode-lbl">Início</span></button>
            <button className={`px-mode ${mode==="macro"?"on":""}`} onClick={()=>setMode("macro")} title="Macroprocessos"><Layers size={15}/> <span className="px-mode-lbl">Macroprocessos</span></button>
            <button className={`px-mode ${mode==="processo"?"on":""}`} onClick={()=>setMode("processo")} title="Processo"><PieChart size={15}/> <span className="px-mode-lbl">Processo</span></button>
            <button className={`px-mode ${mode==="servico"?"on":""}`} onClick={()=>setMode("servico")} title="Serviço"><Sparkles size={15}/> <span className="px-mode-lbl">Serviço</span></button>
            <button className={`px-mode ${mode==="lista"?"on":""}`} onClick={()=>setMode("lista")} title="Entrega"><LayoutGrid size={15}/> <span className="px-mode-lbl">Entrega</span></button>
            <span className="px-modes-sep"/>
            {/* ferramentas de apoio */}
            <Tip pos="bottom" text="Importar PGD: a esteira de conversão do ciclo mensal — cada registro do PGD vira uma entrega do catálogo, com triagem humana e pré-inclusão na lista da unidade.">
              <button className={`px-mode cta-pgd ${mode==="pgd"?"on":""}`} onClick={()=>setMode("pgd")}><Upload size={15}/> <span className="px-mode-lbl">Importar PGD</span></button>
            </Tip>
            <Tip pos="bottom" text="Assistente de Entregas: descreva o que sua área faz e a IA sugere as entregas do catálogo — abre aqui mesmo, na área das visualizações.">
              <button className={`px-mode cta-bot ${mode==="assistente"?"on":""}`} onClick={()=>setMode("assistente")}><Bot size={15}/> <span className="px-mode-lbl">Assistente IA</span></button>
            </Tip>
            <Tip pos="bottom" text="Revisão do banco: console do curador para tratar entregas similares, sem serviço vinculado ou sinalizadas pelo uso do portal.">
              <button className={`px-mode ${mode==="revisao"?"on":""}`} onClick={()=>setMode("revisao")}><ShieldCheck size={15}/> <span className="px-mode-lbl">Revisão do banco</span>{flags.length>0 && <span className="px-mode-badge">{flags.length}</span>}</button>
            </Tip>
          </div>
        </div>
        </div>

        <div className={`px-body ${navPanel||descCollapsed?"solo":""} ${navPanel?"nav":""}`}>
          {navPanel==="natureza" && <ListaFiltroNav titulo="Natureza" icon={Layers} itens={naturezaOpts} valorAtual={natF} onPick={id=>{chooseNat(id);setNavPanel(null);}} onClear={()=>chooseNat(null)} onClose={()=>setNavPanel(null)} corDe={id=>NAT[id].cor} semBusca/>}
          {navPanel==="macro" && <MacroNav natF={natF} macroF={macroF} onPick={pickMacro} onClear={()=>setMacroF(null)} onClose={()=>setNavPanel(null)}/>}
          {navPanel==="categoria" && <ListaFiltroNav titulo="Categorias" icon={PieChart} itens={categoriaOpts} valorAtual={catF} onPick={pickCategoria} onClear={()=>setCatF(null)} onClose={()=>setNavPanel(null)}/>}
          {navPanel==="servico" && <ListaFiltroNav titulo="Serviços" icon={Sparkles} itens={servicoOpts} valorAtual={servF} onPick={pickServico} onClear={()=>setServF(null)} onClose={()=>setNavPanel(null)}/>}
          {navPanel==="orgao" && <ListaFiltroNav titulo="Órgãos" icon={Building2} itens={orgaoOpts} valorAtual={orgF} onPick={pickOrgao} onClear={()=>setOrgF(null)} onClose={()=>setNavPanel(null)}/>}
          <main className="px-main">
            {loading ? <BancoLoading/> : <>
              {bancoStatus==="fallback" && <div className="px-banco-aviso"><AlertTriangle size={14}/> Banco completo (<b>banco.json</b>) não encontrado — exibindo amostra de demonstração. Em produção, publique o arquivo junto ao portal para as 31.241 entregas.</div>}
              {/* nível 4: a entrega. Lista e árvore são duas leituras do mesmo
                  recorte, então viraram uma tela só com alternador. */}
              {mode==="lista" && <>
                <div className="px-vista-row">
                  <TrilhaNiveis nivel="lista" trilha={{macro:macroF,processo:catF,servico:servF}}
                    ir={irNivel} limpar={limparTrilha}/>
                  <div className="px-vista">
                    <button className={vistaEntrega==="lista"?"on":""} onClick={()=>setVistaEntrega("lista")} title="Ver como lista">
                      <LayoutGrid size={13}/> Lista
                    </button>
                    <button className={vistaEntrega==="arvore"?"on":""} onClick={()=>setVistaEntrega("arvore")} title="Ver como árvore">
                      <Network size={13}/> Árvore
                    </button>
                  </div>
                </div>
                {vistaEntrega==="lista"
                  ? <ListaEnriquecida res={res} searching={searching} query={dquery} selSet={selSet} add={add} rem={rem} exp={exp} setExp={setExp} onFlag={setFlagFor} onPropose={t=>setNovaFor(t||"")} compare={compare} toggleCompare={toggleCompare} justAdded={justAdded}/>
                  : <ArvoreDecomposicao res={res} add={add} rem={rem} selSet={selSet}/>}
              </>}
              {mode==="sunburst" && <Sunburst res={res} add={add} rem={rem} selSet={selSet} compare={compare} toggleCompare={toggleCompare} onFlag={setFlagFor} justAdded={justAdded}/>}
              {mode==="nova" && <NovaEntregaCatalogo onFlash={flash} onPropose={t=>setNovaFor(t||"")}/>}
              {mode==="marcos" && <MarcosReferenciais onAdd={add} onDone={()=>setMode("macro")}
                arqs={marcosArqs} setArqs={setMarcosArqs}
                resposta={marcosResposta} setResposta={setMarcosResposta}
                qtd={marcosQtd} setQtd={setMarcosQtd}
                erro={marcosErro} setErro={setMarcosErro}
                orgao={orgao} unidade={unidade} setOrgao={setOrgao} setUnidade={setUnidade}
                conversor={<ConversorUnificado add={add} selSet={selSet} sel={sel} notes={notes}
                  orgao={orgao} unidade={unidade} flash={flash} onAbrirAssistente={()=>setMode("assistente")}/>}/>}
              {(mode==="macro"||mode==="processo"||mode==="servico") &&
                <NivelCatalogo nivel={mode} sel={sel} selSet={selSet} add={add} rem={rem}
                  escopo={escopo} setEscopo={setEscopo}
                  trilha={{macro:macroF,processo:catF,servico:servF}}
                  descer={descerNivel} irNivel={irNivel} limparTrilha={limparTrilha}
                  onGoNova={()=>setMode("nova")} orgao={orgao} unidade={unidade}/>}
              {mode==="assistente" && <Assistente onAdd={add}/>}
              {mode==="pgd" && <ImportarPGD onConversor={()=>setMode("marcos")}/>}
              {mode==="revisao" && <CentralRevisao embutido flags={flags} onClose={()=>setMode("lista")}/>}
              {/* A Descrição da Área deixou de ter botão próprio na barra: vive no
                  painel lateral e no botão flutuante, que é onde ela é usada. */}
            </>}
          </main>

          {!navPanel && !descCollapsed && <aside className="px-doc">
            <div className="px-doc-h"><div className="px-doc-t"><ClipboardList size={16}/> Descrição da Área</div><span className="px-doc-badge">{sel.length}</span>
              <div className="px-doc-view">
                <button className={descView==="lista"?"on":""} onClick={()=>setDescView("lista")} title="Ver como lista"><ClipboardList size={13}/></button>
                <Tip pos="bottom" text="Cadeia de valor: veja suas entregas ligadas aos objetivos da unidade e do órgão, em tela cheia. Ótimo para enxergar o encadeamento do trabalho.">
                  <button className={`cadeia ${descView==="cadeia"?"on":""}`} onClick={()=>setDescView("cadeia")}><Workflow size={13}/></button>
                </Tip>
              </div>
              <button className="px-doc-collapse" onClick={()=>setDescCollapsed(true)} title="Recolher o painel e dar mais espaço ao catálogo"><ChevronRight size={15}/></button>
            </div>
            {docHeader}
            <DescPanel sel={sel} notes={notes} setNote={setNote} rem={rem} onInject={()=>sel.length&&setPreview(true)} orgao={orgao} unidade={unidade}/>
          </aside>}
        </div>
      </div>}

      {/* botões flutuantes (apenas na seção Catálogo) */}
      {secao==="catalogo" && navPanel && <button className="px-descfab" onClick={()=>setDescDrawer(true)}><ClipboardList size={16}/> Minha descrição <span>{sel.length}</span></button>}
      {secao==="catalogo" && !navPanel && <button className="px-descfab pxdesc" data-collapsed={descCollapsed?1:0} onClick={openDesc} title="Abrir a Descrição da Área"><ClipboardList size={16}/> Descrição <span>{sel.length}</span></button>}
      {secao==="catalogo" && compare.length>0 && <button className="px-cmpfab" onClick={()=>setCompareOpen(true)}><Columns size={16}/> Comparar <span>{compare.length}</span></button>}

      {descDrawer && <div className="px-drawer-bg" onClick={()=>setDescDrawer(false)}>
        <div className="px-drawer" onClick={e=>e.stopPropagation()}>
          <div className="px-doc-h"><div className="px-doc-t"><ClipboardList size={16}/> Descrição da Área</div><button className="px-x" onClick={()=>setDescDrawer(false)}><X size={17}/></button></div>
          {docHeader}
          <DescPanel sel={sel} notes={notes} setNote={setNote} rem={rem} onInject={()=>{setDescDrawer(false);sel.length&&setPreview(true);}} orgao={orgao} unidade={unidade}/>
        </div>
      </div>}

      {descView==="cadeia" && <CadeiaValor sel={sel} rem={rem} objUnidade={objUnidade} setObjUnidade={setObjUnidade} objOrgao={objOrgao} setObjOrgao={setObjOrgao} onClose={()=>setDescView("lista")}/>}

      {compareOpen && <CompareModal items={compare} onClose={()=>setCompareOpen(false)} onAdd={add} selSet={selSet} onDrop={c=>setCompare(x=>x.filter(e=>e.codigo!==c))} onClear={()=>{setCompare([]);setCompareOpen(false);}}/>}
      {preview && <PreviewDFT sel={sel} notes={notes} orgao={orgao} unidade={unidade} onClose={()=>setPreview(false)} onConfirm={()=>{setPreview(false);flash("Descrição enviada ao DFT (demonstração).");}}/>}
      {flagFor && <FlagModal entrega={flagFor} onClose={()=>setFlagFor(null)} onSubmit={(f)=>{setFlags(s=>[...s,f]);setFlagFor(null);flash("Sinalização enviada à curadoria do banco.");}}/>}
      {novaFor!==null && <FlagModal nova proposta={novaFor} onClose={()=>setNovaFor(null)} onSubmit={(f)=>{setFlags(s=>[...s,f]);setNovaFor(null);flash("Proposta de nova entrega enviada à curadoria.");}}/>}
      {/* Revisão e Qualidade do banco agora são visualizações (mode==="revisao"/"qualidade") dentro do Catálogo, não overlays/seções próprias. */}
      {/* Organizações agora é um app próprio (organizacoes.html) — ver OrganizacoesApp exportado no fim deste arquivo */}
      {toast && <div className="px-toast">{toast}</div>}
    </div>
  );
}

/* ---------- descrição da área (agrupada por natureza, com observações) ---------- */
/* Regra de criação por natureza: só o finalístico é aberto nos três níveis
   acima da entrega. Suporte e governança são estrutura comum a todo o governo
   — ali a unidade no máximo propõe um serviço novo, e mesmo assim o conferidor
   checa antes se já existe algo equivalente. */
const PODE_CRIAR={
  "Finalístico":{macro:true, cat:true, serv:true},
  "Governança": {macro:false,cat:false,serv:true},
  "Suporte":    {macro:false,cat:false,serv:true},
};
const NAT_CHAVE={"Finalístico":"finalistico","Governança":"governanca","Suporte":"suporte"};

function NovaEntregaCatalogo({onFlash,onPropose}){
  const [nat,setNat]=useState("Finalístico");
  const [mac,setMac]=useState(""),[cat,setCat]=useState(""),[serv,setServ]=useState(""),[ent,setEnt]=useState("");
  const [feito,setFeito]=useState(false);
  const regra=PODE_CRIAR[nat]||PODE_CRIAR["Finalístico"];

  // vocabulário restrito à natureza escolhida — o que já existe naquele ramo
  const vocab=useMemo(()=>{
    const chave=NAT_CHAVE[nat];
    const M=new Set(),Ca=new Set(),S=new Set();
    ENTREGAS.forEach(e=>{ if(e.natureza!==chave) return;
      if(e.macro)M.add(e.macro);
      if(e.categoria&&(!mac||e.macro===mac))Ca.add(e.categoria);
      if(e.servico&&e.servico!==e.categoria&&(!cat||e.categoria===cat))S.add(e.servico); });
    return {macros:[...M].sort(),cats:[...Ca].sort(),servs:[...S].sort().slice(0,1500)};
  },[nat,mac,cat]);

  // o que é novo em cada nível (não existe no vocabulário) → indicador
  const ehNovo=(v,lista)=>!!v.trim() && !lista.some(x=>norm(x)===norm(v));
  const novoMac=ehNovo(mac,vocab.macros), novoCat=ehNovo(cat,vocab.cats), novoServ=ehNovo(serv,vocab.servs);
  const violacao = (novoMac&&!regra.macro) ? "macroprocesso"
    : (novoCat&&!regra.cat) ? "processo"
    : (novoServ&&!regra.serv) ? "serviço" : null;

  const cand=useMemo(()=>{ if(norm(ent).length<6) return [];
    const toks=x=>new Set(norm(x).split(" ").filter(Boolean)); const T=toks(ent);
    return ENTREGAS.map(e=>{ const B=toks(e.entrega); let inter=0; T.forEach(t=>{if(B.has(t))inter++;});
      const sc=inter/((T.size+B.size-inter)||1);
      return {e,s:sc,nMac:norm(mac)===norm(e.macro),nCat:norm(cat)===norm(e.categoria),nServ:!!serv&&norm(serv)===norm(e.servico)}; })
      .filter(x=>x.s>0.34).sort((a,b)=>b.s-a.s).slice(0,6); },[ent,mac,cat,serv]);
  const top=cand[0];
  const veredito=!top?"ok":(top.s>=0.72&&top.nCat)?"bloqueio":(top.s>=0.55)?"atencao":"ok";

  function inserir(){
    if(violacao){ onFlash&&onFlash(`Em ${nat.toLowerCase()}, ${violacao} novo não pode ser criado — escolha um já existente.`); return; }
    if(veredito==="bloqueio"){onFlash&&onFlash("Muito parecida com uma entrega existente — verifique antes de propor.");return;}
    if(!ent.trim()||!mac||!cat){onFlash&&onFlash("Preencha ao menos macroprocesso, processo e o texto da entrega.");return;}
    onPropose&&onPropose(ent.trim()); setFeito(true);
  }

  // rótulo "novo"/"travado" ao lado de cada nível
  const marca=(novo,permitido)=> novo
    ? (permitido ? <span className="px-nova-tag nova"><Sparkles size={9}/> novo</span>
                 : <span className="px-nova-tag trava"><AlertTriangle size={9}/> não permitido nesta natureza</span>)
    : null;

  return (
    <div className="px-nova">
      <div className="px-nova-hero"><span className="px-nova-badge"><FilePlus2 size={18}/></span>
        <div><h2>Criar novo no catálogo</h2><p>Antes de propor ao banco, o sistema confere se já existe algo parecido nos 4 níveis (macroprocesso → processo → serviço → entrega).</p></div></div>
      <div className="px-nova-grid">
        <div className="px-nova-form">
          <label className="px-nova-l">Natureza</label>
          <div className="px-nova-nat">{Object.keys(NAT_CHAVE).map(n=>(
            <button key={n} className={`px-nova-chip ${nat===n?"on":""}`}
              style={nat===n?{background:NAT[NAT_CHAVE[n]].cor,borderColor:NAT[NAT_CHAVE[n]].cor,color:"#fff"}:{}}
              onClick={()=>{setNat(n);setMac("");setCat("");setServ("");}}>{n}</button>))}</div>

          <div className={`px-nova-regra ${regra.macro?"aberta":"restrita"}`}>
            {regra.macro
              ? <><Check size={12}/> Em finalístico você pode criar macroprocesso, processo e serviço novos.</>
              : <><AlertTriangle size={12}/> Em {nat.toLowerCase()}, macroprocesso e processo vêm da estrutura existente — só o serviço pode ser novo.</>}
          </div>

          <label className="px-nova-l">Macroprocesso {marca(novoMac,regra.macro)}</label>
          {regra.macro
            ? <input list="px-nova-macros" className="px-nova-in" value={mac} onChange={e=>{setMac(e.target.value);setCat("");setServ("");}} placeholder="Escolha um existente ou digite um novo…"/>
            : <select className="px-nova-in" value={mac} onChange={e=>{setMac(e.target.value);setCat("");setServ("");}}>
                <option value="">Escolha o macroprocesso…</option>
                {vocab.macros.map(m=><option key={m} value={m}>{m}</option>)}
              </select>}
          <datalist id="px-nova-macros">{vocab.macros.map(m=><option key={m} value={m}/>)}</datalist>

          <label className="px-nova-l">Processo (categoria de serviço) {marca(novoCat,regra.cat)}</label>
          {regra.cat
            ? <input list="px-nova-cats" className="px-nova-in" value={cat} onChange={e=>{setCat(e.target.value);setServ("");}} placeholder="Escolha um existente ou digite um novo…"/>
            : <select className="px-nova-in" value={cat} onChange={e=>{setCat(e.target.value);setServ("");}} disabled={!mac}>
                <option value="">{mac?"Escolha o processo…":"Escolha o macroprocesso primeiro"}</option>
                {vocab.cats.map(c=><option key={c} value={c}>{c}</option>)}
              </select>}
          <datalist id="px-nova-cats">{vocab.cats.map(c=><option key={c} value={c}/>)}</datalist>

          <label className="px-nova-l">Serviço {marca(novoServ,regra.serv)}</label>
          <input list="px-nova-servs" className="px-nova-in" value={serv} onChange={e=>setServ(e.target.value)} placeholder="Opcional — pode ser um serviço novo"/>
          <datalist id="px-nova-servs">{vocab.servs.map(x=><option key={x} value={x}/>)}</datalist>

          <label className="px-nova-l">Texto da entrega</label>
          <textarea className="px-nova-ta" value={ent} onChange={e=>{setEnt(e.target.value);setFeito(false);}} placeholder="Ex.: Relatório de gestão consolidado elaborado"/>
          <button className={`px-nova-btn ${veredito==="bloqueio"||violacao?"blocked":""}`} onClick={inserir} disabled={feito}>
            {feito?<><Check size={15}/> Proposta enviada à curadoria</>:<><Plus size={15}/> Propor ao catálogo</>}</button>
        </div>
        <div className="px-nova-check">
          <div className="px-nova-check-h"><Search size={15}/> Já existe algo parecido?</div>
          {violacao && <div className="px-nova-verd v-bloqueio" style={{marginBottom:"10px"}}>
            <AlertTriangle size={15}/> {violacao.charAt(0).toUpperCase()+violacao.slice(1)} novo não é permitido em {nat.toLowerCase()}.
          </div>}
          {norm(ent).length<6
            ? <div className="px-nova-empty">Digite o texto da entrega para ver as mais próximas no banco.</div>
            : <>
              <div className={`px-nova-verd v-${veredito}`}>
                {veredito==="bloqueio" && <><AlertTriangle size={15}/> Praticamente idêntica a uma entrega existente no mesmo processo.</>}
                {veredito==="atencao" && <><AlertTriangle size={15}/> Há entregas próximas — confirme que é realmente distinta.</>}
                {veredito==="ok" && <><Check size={15}/> Nada muito próximo. Parece uma entrega nova.</>}
              </div>
              <div className="px-nova-cands">{cand.map((c,i)=>(
                <div className="px-nova-cand" key={c.e.codigo+i}>
                  <span className="px-nova-cand-sim">{Math.round(c.s*100)}%</span>
                  <div className="px-nova-cand-tx"><b>{c.e.entrega}</b><span>{c.e.codigo} · {c.e.macro}</span></div>
                  <div className="px-nova-cand-niv" title="Níveis que coincidem">
                    <span className={c.nMac?"hit":""}>M</span>
                    <span className={c.nCat?"hit":""}>P</span>
                    <span className={c.nServ?"hit":""}>S</span>
                  </div>
                </div>))}</div>
            </>}
        </div>
      </div>
    </div>
  );
}

/* ============================================================================
   MARCOS REFERENCIAIS — upload dos 4 documentos institucionais da unidade
   (Cadeia de Valor, Estrutura Organizacional, Relatório de Gestão Integrado,
   Regimento Interno) e leitura por IA para identificar quais macroprocessos,
   processos e serviços do catálogo a unidade executa. Ao concluir, adiciona
   os achados à Descrição da Área e leva para a aba Macroprocessos.
============================================================================ */
function MarcosReferenciais({onAdd,onDone,arqs,setArqs,resposta,setResposta,qtd,setQtd,erro,setErro,
                             orgao,unidade,setOrgao,setUnidade,conversor}){
  const [carregando,setCarregando]=useState(false);
  // duas portas de entrada para o mesmo destino: os documentos institucionais
  // (marcos) ou o que a unidade já tem pronto em outro formato (conversor)
  const [porta,setPorta]=useState("marcos");
  const codMap=useMemo(()=>{const m=new Map();ENTREGAS.forEach(e=>m.set(e.codigo,e));return m;},[]);
  // Os 4 documentos institucionais são sugestões fixas; além deles a unidade
  // pode anexar quantos outros quiser (PGD, planejamento estratégico, etc.),
  // cada um com um rótulo livre que também vai para a IA.
  const campos=[["cadeia","Cadeia de Valor"],["estrutura","Estrutura Organizacional"],["rgi","Relatório de Gestão Integrado"],["regimento","Regimento Interno"]];
  const extras=Object.keys(arqs).filter(k=>k.startsWith("extra")&&arqs[k]);

  async function escolherArquivo(chave,ev,rotulo){
    const f=ev.target.files?.[0]; if(!f) return; setErro("");
    try{ const a=await lerArquivo(f); setArqs(s=>({...s,[chave]:{...a,rotulo:rotulo||s[chave]?.rotulo||""}})); }
    catch(err){ setErro(err.message||"Não consegui ler o arquivo."); }
    finally{ ev.target.value=""; }
  }
  async function adicionarExtra(ev){
    const f=ev.target.files?.[0]; if(!f) return; setErro("");
    try{
      const a=await lerArquivo(f);
      const chave="extra"+Date.now().toString(36);
      setArqs(s=>({...s,[chave]:{...a,rotulo:"Outro documento"}}));
    }catch(err){ setErro(err.message||"Não consegui ler o arquivo."); }
    finally{ ev.target.value=""; }
  }
  function renomearExtra(chave,rotulo){ setArqs(s=>({...s,[chave]:{...s[chave],rotulo}})); }
  function remover(chave){
    // extras somem de vez; os 4 fixos voltam a ser slot vazio
    if(chave.startsWith("extra")) setArqs(s=>{ const n={...s}; delete n[chave]; return n; });
    else setArqs(s=>({...s,[chave]:null}));
  }

  async function analisar(){
    const lista=Object.values(arqs).filter(Boolean);
    if(!lista.length||carregando) return;
    setCarregando(true); setResposta(null); setQtd(0); setErro("");
    const textos=lista.filter(a=>a.tipo!=="pdf").map(a=>`Documento "${a.rotulo?`${a.rotulo} — ${a.nome}`:a.nome}":\n${a.texto}`).join("\n\n");
    const cands=buscarCandidatas(textos||lista.map(a=>a.nome).join(" "),80);
    const catalogo=cands.length?candidatasTxt(cands):CAT_TXT;
    const ctx=[orgao&&`Órgão: ${orgao}`,unidade&&`Unidade: ${unidade}`].filter(Boolean).join(" · ");
    const sys=`Você é o Assistente de Marcos Referenciais do SISDIP/DFT (setor público federal). Recebe documentos institucionais de uma unidade — tipicamente Cadeia de Valor, Estrutura Organizacional, Relatório de Gestão Integrado e Regimento Interno, mas também outros que a unidade anexe (PGD, planejamento estratégico, etc.); o rótulo de cada documento vem junto do texto.${ctx?`\n${ctx}`:""}\nIdentifique quais macroprocessos, processos e serviços do catálogo essa unidade executa, citando sempre o código exatamente como aparece na lista de candidatas abaixo (8 dígitos, sem ponto — ex.: 02900067). Baseie-se SOMENTE nas entregas candidatas abaixo. Agrupe por macroprocesso e explique brevemente cada escolha.\nEntregas candidatas (código | entrega | macro > categoria):\n${catalogo}`;
    // O proxy de IA roda atrás de uma função serverless (Netlify/Vercel), que
    // tem limite de payload (~6MB). Relatórios de gestão com centenas de
    // páginas passam disso fácil e a chamada inteira falha com 413. Por isso
    // PDFs grandes demais ficam de fora do envio à IA — os outros documentos
    // seguem normalmente, e avisamos qual foi deixado de fora.
    const LIMITE_PDF_B64=4_000_000, LIMITE_TOTAL_B64=4_500_000; // ~3MB e ~3,4MB de arquivo original
    const pdfsGrandes=[]; let totalB64=0;
    const conteudo=[];
    lista.forEach(a=>{
      if(a.tipo!=="pdf") return;
      const tam=a.pdfB64?.length||0;
      if(tam>LIMITE_PDF_B64||totalB64+tam>LIMITE_TOTAL_B64){ pdfsGrandes.push(a.nome); return; }
      totalB64+=tam;
      conteudo.push({type:"document",source:{type:"base64",media_type:"application/pdf",data:a.pdfB64}});
    });
    const avisoTamanho=pdfsGrandes.length
      ? `Arquivo(s) grande(s) demais para a IA e não enviado(s): ${pdfsGrandes.join(", ")}. Reduza o tamanho (ex.: exporte só o sumário/resumo executivo) ou complete manualmente na Lista.`
      : "";
    if(!conteudo.length && !textos){
      setCarregando(false);
      setErro(avisoTamanho||"Não consegui extrair conteúdo dos arquivos enviados.");
      return;
    }
    conteudo.push({type:"text",text: textos ? `Documentos enviados (texto extraído):\n\n${textos}` : "Analise os documentos PDF anexados."});
    try{
      const d=await chamarIA({model:"claude-haiku-4-5-20251001",max_tokens:1500,system:sys,messages:[{role:"user",content:conteudo}]});
      const tx=(d.content||[]).map(b=>b.type==="text"?b.text:"").join("\n").trim();
      setResposta((tx||"Não consegui identificar nada nos documentos enviados.")+(avisoTamanho?`\n\n⚠️ ${avisoTamanho}`:""));
      // os códigos do banco são 8 dígitos sem ponto (ex.: 02900067) — a IA às
      // vezes segue à risca o "0000.0000" pedido no prompt, às vezes cita o
      // código cru como aparece na lista de candidatas. Aceita os dois.
      const cods=[...new Set((tx.match(/\b\d{4}\.?\d{4}\b/g)||[]).map(v=>v.replace(".","")))].filter(c=>codMap.has(c));
      const achados=cods.map(c=>codMap.get(c));
      achados.forEach(onAdd);
      setQtd(achados.length);
      if(achados.length>0 && onDone) onDone();
    }catch{
      const top=buscarCandidatas(textos||lista.map(a=>a.nome).join(" "),12);
      setErro("Não consegui falar com a IA agora."+(top.length?" Busquei direto no catálogo pelas palavras dos documentos — as entregas mais próximas já foram adicionadas à Descrição da Área:":"")+(avisoTamanho?` ${avisoTamanho}`:""));
      setResposta(top.length?top.map(e=>`${e.codigo} — ${e.entrega}\n   ${e.macro} > ${e.categoria}`).join("\n\n"):null);
      top.forEach(onAdd);
      setQtd(top.length);
      if(top.length>0 && onDone) onDone();
    }finally{ setCarregando(false); }
  }

  if(porta==="conversor") return (
    <div style={{padding:"32px 24px",maxWidth:"1080px"}}>
      <div className="px-etapa1-cad" style={{maxWidth:"720px"}}>
        <div className="px-etapa1-cad-h"><Building2 size={14}/> <b>Unidade</b>
          <span>fica salva neste computador</span></div>
        <DocHeaderEdit orgao={orgao} unidade={unidade} setOrgao={setOrgao} setUnidade={setUnidade}/>
      </div>
      <div className="px-porta">
        <button onClick={()=>setPorta("marcos")}><Bot size={13}/> Marcos referenciais</button>
        <button className="on"><Wand2 size={13}/> Conversor</button>
      </div>
      {conversor}
    </div>
  );

  return (
    <div style={{padding:"32px 24px",maxWidth:"720px"}}>
      {/* etapa 1 do caminho: quem é a unidade + o que ela produz, no mesmo lugar */}
      <div className="px-etapa1-cad">
        <div className="px-etapa1-cad-h"><Building2 size={14}/> <b>Unidade</b>
          <span>fica salva neste computador</span></div>
        <DocHeaderEdit orgao={orgao} unidade={unidade} setOrgao={setOrgao} setUnidade={setUnidade}/>
      </div>

      {conversor && <div className="px-porta">
        <button className="on"><Bot size={13}/> Marcos referenciais</button>
        <button onClick={()=>setPorta("conversor")}><Wand2 size={13}/> Conversor</button>
      </div>}

      <div style={{fontSize:"13px",lineHeight:1.5,marginBottom:"16px",display:"flex",gap:"8px",alignItems:"flex-start"}}>
        <Bot size={15}/>
        <span>Anexe os marcos referenciais da unidade (PDF ou Word). A IA lê os documentos, adiciona automaticamente à Descrição da Área e leva você à aba Macroprocessos com o que foi encontrado.</span>
      </div>
      <div style={{display:"flex",flexWrap:"wrap",gap:"10px",marginBottom:"10px"}}>
        {campos.map(([chave,rotulo])=>(
          <div key={chave} style={{border:`1px solid ${C.line}`,borderRadius:"10px",padding:"10px 12px",minWidth:"220px",flex:"1 1 220px"}}>
            <div style={{fontSize:"11px",fontWeight:700,marginBottom:"6px",opacity:.7}}>{rotulo}</div>
            {arqs[chave]
              ? <span className="px-anexo-chip">
                  <FileText size={12}/> {arqs[chave].nome}
                  <label style={{cursor:"pointer",display:"inline-flex",alignItems:"center",color:C.sub}} title="Trocar arquivo">
                    <RefreshCw size={11}/>
                    <input type="file" accept=".pdf,.docx" style={{display:"none"}} onChange={ev=>escolherArquivo(chave,ev,rotulo)}/>
                  </label>
                  <button onClick={()=>remover(chave)} title="Remover arquivo"><X size={11}/></button>
                </span>
              : <label className="px-mode" style={{cursor:"pointer"}}>
                  <Paperclip size={14}/> <span className="px-mode-lbl">Selecionar arquivo</span>
                  <input type="file" accept=".pdf,.docx" style={{display:"none"}} onChange={ev=>escolherArquivo(chave,ev,rotulo)}/>
                </label>}
          </div>
        ))}
        {/* documentos além dos 4 sugeridos — rótulo livre */}
        {extras.map(chave=>(
          <div key={chave} style={{border:`1px solid ${C.line}`,borderRadius:"10px",padding:"10px 12px",minWidth:"220px",flex:"1 1 220px"}}>
            <input className="px-extra-rot" value={arqs[chave].rotulo||""} placeholder="Nome do documento"
              onChange={ev=>renomearExtra(chave,ev.target.value)} title="Como este documento se chama"/>
            <span className="px-anexo-chip">
              <FileText size={12}/> {arqs[chave].nome}
              <button onClick={()=>remover(chave)} title="Remover arquivo"><X size={11}/></button>
            </span>
          </div>
        ))}
      </div>
      <label className="px-extra-add">
        <Plus size={13}/> Adicionar outro documento
        <input type="file" accept=".pdf,.docx,.xlsx,.xls,.csv" style={{display:"none"}} onChange={adicionarExtra}/>
      </label>
      {erro && <div className="px-anexo-erro" style={{marginBottom:"10px"}}><AlertTriangle size={12}/> {erro}</div>}
      <button className="px-mode cta-bot" disabled={carregando||!Object.values(arqs).some(Boolean)} onClick={analisar}>
        <Bot size={15}/> <span className="px-mode-lbl">{carregando?"Analisando…":"Analisar com IA"}</span>
      </button>
      {qtd>0 && <div style={{marginTop:"12px",fontSize:"12.5px",color:C.green,fontWeight:700,display:"flex",alignItems:"center",gap:"6px"}}>
        <Check size={14}/> {qtd} entrega{qtd===1?"":"s"} adicionada{qtd===1?"":"s"} à Descrição da Área — abrindo Macroprocessos…
      </div>}
      {resposta && <div className="px-chat-b" style={{marginTop:"16px"}}>
        <div className="px-msg assistant"><div className="px-msg-tx">{resposta}</div></div>
      </div>}
      <p style={{fontSize:"12px",opacity:.65,marginTop:"18px"}}>O que a IA não encontrar, complemente manualmente na aba <b>Lista</b>.</p>
    </div>
  );
}

/* ============================================================================
   MACROPROCESSOS — árvore Natureza → Macroprocesso → Categoria/Serviço (sem
   descer até as entregas individuais), com contagem de selecionadas em cada
   nível, busca para adicionar mais, link para propor o que não existe, e
   exportação em PDF no formato "Relatório de Macroprocessos".
============================================================================ */
/* Busca genérica para adicionar entregas fora do que a IA já encontrou —
   compartilhada por Macroprocessos, Processo e Serviço. */
function BuscaAdicionar({selSet,add,onGoNova}){
  const [busca,setBusca]=useState("");
  const candidatos=useMemo(()=>{
    const q=norm(busca.trim()); if(!q) return [];
    return ENTREGAS.filter(e=>!selSet.has(e.codigo) &&
      (norm(e.entrega).includes(q)||norm(e.macro).includes(q)||norm(e.categoria).includes(q)||e.codigo.includes(busca.trim()))
    ).slice(0,8);
  },[busca,selSet]);
  function adicionarBusca(e){ if(!selSet.has(e.codigo)){ add(e); setBusca(""); } }
  return (
    <div style={{position:"relative",marginTop:"18px"}}>
      <label style={{display:"block",fontSize:"11.5px",fontWeight:700,color:C.navy,marginBottom:"6px"}}>Buscar entrega para adicionar</label>
      <input value={busca} onChange={e=>setBusca(e.target.value)} placeholder="Buscar entrega para adicionar…"
        style={{width:"100%",padding:"8px 10px",border:`1px solid ${C.line}`,borderRadius:"8px",fontSize:"12.5px",boxSizing:"border-box"}}/>
      {busca.trim() && candidatos.length===0 && (
        <div style={{border:`1px solid ${C.line}`,borderRadius:"8px",marginTop:"4px",padding:"10px 12px",fontSize:"12px",color:C.sub,display:"flex",justifyContent:"space-between",alignItems:"center",gap:"8px",flexWrap:"wrap"}}>
          <span>Nenhuma entrega encontrada para "{busca.trim()}".</span>
          <button className="px-mode" style={{padding:"3px 8px"}} onClick={onGoNova}>Propor macroprocesso, processo e serviço novos</button>
        </div>
      )}
      {candidatos.length>0 && (
        <div style={{border:`1px solid ${C.line}`,borderRadius:"8px",marginTop:"4px",overflow:"hidden"}}>
          {candidatos.map(c=>(
            <div key={c.codigo} style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:"8px",padding:"7px 10px",fontSize:"12px",borderBottom:"1px solid #EFF2F6"}}>
              <span><b>{c.codigo}</b> — {c.entrega}</span>
              <button className="px-mode" style={{padding:"3px 8px"}} onClick={()=>adicionarBusca(c)}><Plus size={12}/> Adicionar</button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

/* Estado vazio compartilhado — aparece quando a Descrição da Área ainda não tem nada
   (antes de rodar os Marcos Referenciais, ou depois de limpar a seleção). */
function VazioSelecaoIA({Icone}){
  return (
    <div style={{border:`1px dashed ${C.line}`,borderRadius:"10px",padding:"22px",textAlign:"center",color:C.sub,fontSize:"13px"}}>
      <Icone size={22} color={C.faint} style={{marginBottom:"8px"}}/>
      <div>Nada aqui ainda. Vá em <b>Marcos Referenciais</b> pra deixar a IA identificar automaticamente, ou adicione manualmente na busca abaixo.</div>
    </div>
  );
}

/* ============================================================================
   OS 4 NÍVEIS DO CATÁLOGO — macroprocesso → processo → serviço → entrega.
   Um componente só serve os três primeiros níveis (o quarto é a Lista): mudam
   o campo agrupador e o nível seguinte. Clicar num item desce um nível
   levando o filtro junto; a trilha no topo permite subir de volta.
   O escopo alterna entre "só o que está na Descrição da Área" (o que a IA
   encontrou, que é o padrão) e "todo o catálogo".
============================================================================ */
const SEM_SERVICO="Sem serviço específico";
const NIVEIS={
  macro:    {campo:"macro",     rot:"Macroprocesso", plural:"macroprocessos", proximo:"processo", ic:Layers},
  processo: {campo:"categoria", rot:"Processo",      plural:"processos",      proximo:"servico",  ic:PieChart},
  servico:  {campo:"servico",   rot:"Serviço",       plural:"serviços",       proximo:"lista",    ic:Sparkles},
};

/* trilha clicável dos 4 níveis — mostra onde você está e deixa voltar */
function TrilhaNiveis({nivel,trilha,ir,limpar}){
  const passos=[
    {id:"macro",    rot:"Macroprocesso", val:trilha.macro},
    {id:"processo", rot:"Processo",      val:trilha.processo},
    {id:"servico",  rot:"Serviço",       val:trilha.servico},
    {id:"lista",    rot:"Entrega",       val:null},
  ];
  const idx=passos.findIndex(p=>p.id===nivel);
  return (
    <div className="px-trilha">
      {passos.map((p,i)=>(
        <Fragment key={p.id}>
          {i>0 && <ChevronRight size={13} className="px-trilha-sep"/>}
          <button className={`px-trilha-p ${i===idx?"on":""} ${i<idx?"feito":""}`}
            onClick={()=>ir(p.id)} title={`Ir para ${p.rot}`}>
            <span className="px-trilha-rot">{p.rot}</span>
            {p.val && <span className="px-trilha-val">{p.val.length>26?p.val.slice(0,25)+"…":p.val}</span>}
          </button>
        </Fragment>
      ))}
      {(trilha.macro||trilha.processo||trilha.servico) &&
        <button className="px-trilha-limpar" onClick={limpar} title="Limpar a trilha e voltar ao topo">
          <X size={11}/> limpar
        </button>}
    </div>
  );
}

function NivelCatalogo({nivel,sel,selSet,add,rem,escopo,setEscopo,trilha,descer,irNivel,limparTrilha,
                        onGoNova,orgao,unidade}){
  const cfg=NIVEIS[nivel];
  const [removidos,setRemovidos]=useState([]);   // permite remarcar sem refazer a análise

  // base = só o selecionado (padrão) ou o catálogo inteiro, sempre recortada
  // pelos níveis já escolhidos acima na trilha
  const agrupado=useMemo(()=>{
    // No escopo "só selecionados" os grupos desmarcados continuam na tela
    // (apagados), senão sumiriam no instante do clique e não haveria como
    // devolvê-los à Descrição da Área.
    let base;
    if(escopo!=="sel") base=ENTREGAS;
    else{
      const vistos=new Set(sel.map(e=>e.codigo));
      base=[...sel,...removidos.filter(e=>!vistos.has(e.codigo))];
    }
    const porNat=new Map();
    base.forEach(e=>{
      if(nivel!=="macro" && trilha.macro && e.macro!==trilha.macro) return;
      if(nivel==="servico" && trilha.processo && e.categoria!==trilha.processo) return;
      // ~73% das entregas não têm serviço próprio: ingestBanco usa a categoria
      // como serviço pra não deixar o campo vazio. No nível Serviço isso
      // enganaria (pareceria o nível Processo), então agrupamos à parte.
      const semServicoProprio = nivel==="servico" && e.servico===e.categoria;
      const valor = e[cfg.campo]||"";
      const nome = semServicoProprio ? SEM_SERVICO : (valor||"—");
      if(!porNat.has(e.natureza)) porNat.set(e.natureza,new Map());
      const porGrupo=porNat.get(e.natureza);
      if(!porGrupo.has(nome)) porGrupo.set(nome,{nome,valor,semServico:semServicoProprio,itens:[]});
      porGrupo.get(nome).itens.push(e);
    });
    return NAT_ORDER.filter(n=>porNat.has(n)).map(nat=>{
      const grupos=[...porNat.get(nat).values()]
        .sort((a,b)=> (a.semServico!==b.semServico) ? (a.semServico?1:-1) : a.nome.localeCompare(b.nome));
      return {nat,grupos,tot:grupos.reduce((s,g)=>s+g.itens.length,0)};
    });
  },[sel,removidos,escopo,nivel,trilha.macro,trilha.processo,cfg.campo]);

  const [natAberta,setNatAberta]=useState(()=>new Set(NAT_ORDER));
  function toggleNat(id){ setNatAberta(prev=>{ const s=new Set(prev); s.has(id)?s.delete(id):s.add(id); return s; }); }

  // caixinha: confere o que a IA trouxe. Desmarcar tira o grupo da Descrição
  // da Área (guardando de lado); marcar devolve.
  function alternarGrupo(g){
    const dentro=g.itens.filter(e=>selSet.has(e.codigo));
    if(dentro.length){
      setRemovidos(r=>[...r,...dentro]);
      dentro.forEach(e=>rem(e.codigo));
    }else{
      const voltar=removidos.filter(e=>(nivel==="servico"&&g.semServico)
        ? e.servico===e.categoria && (!trilha.processo||e.categoria===trilha.processo)
        : e[cfg.campo]===g.valor);
      voltar.forEach(add);
      setRemovidos(r=>r.filter(e=>!voltar.includes(e)));
    }
  }

  const conferindo = escopo==="sel";
  // no modo conferência os totais contam só o que segue marcado
  const contar=g=>conferindo ? g.itens.filter(e=>selSet.has(e.codigo)).length : g.itens.length;
  const totGrupos=agrupado.reduce((s,N)=>s+N.grupos.filter(g=>!conferindo||contar(g)>0).length,0);
  const totItens=agrupado.reduce((s,N)=>s+N.grupos.reduce((x,g)=>x+contar(g),0),0);

  return (
    <div style={{padding:"24px 24px 32px",maxWidth:"920px"}}>
      <TrilhaNiveis nivel={nivel} trilha={trilha} ir={irNivel} limpar={limparTrilha}/>

      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:"10px",marginBottom:"4px"}}>
        <h2 style={{margin:0,fontSize:"18px",color:C.navy}}>{cfg.rot}</h2>
        <div style={{display:"flex",gap:"8px",alignItems:"center"}}>
          <div className="px-escopo">
            <button className={escopo==="sel"?"on":""} onClick={()=>setEscopo("sel")}>Só os selecionados</button>
            <button className={escopo==="tudo"?"on":""} onClick={()=>setEscopo("tudo")}>Todo o catálogo</button>
          </div>
          {nivel==="macro" && <button className="px-mode" disabled={!sel.length}
            onClick={()=>gerarPdfRelatorioMacroprocessos(sel,orgao,unidade)}>
            <Download size={14}/> <span className="px-mode-lbl">Baixar PDF</span>
          </button>}
        </div>
      </div>
      <p style={{fontSize:"12.5px",color:C.faint,marginBottom:"14px"}}>
        {conferindo
          ? <>{sel.length} entrega{sel.length===1?"":"s"} na Descrição da Área — desmarque o que não é da unidade, clique para descer ao próximo nível.</>
          : <>Navegando por todo o catálogo. Clique para descer até a entrega.</>}
        {nivel==="macro" && (orgao||unidade) ? <> · {orgao||"—"} / {unidade||"—"}</> : null}
      </p>

      {totGrupos===0 ? <VazioSelecaoIA Icone={cfg.ic}/> : (
        <div className="px-tree">
          <div className="px-treebar">
            <div className="px-treebar-l">
              <b>{conferindo?"Descrição da Área":"Catálogo"}</b>
              <span>· {totGrupos} {totGrupos===1?cfg.rot.toLowerCase():cfg.plural} · {totItens.toLocaleString("pt-BR")} entregas</span>
            </div>
            <div className="px-treebar-r">
              <button onClick={()=>setNatAberta(new Set(NAT_ORDER))}>Expandir tudo</button><span>/</span>
              <button onClick={()=>setNatAberta(new Set())}>Recolher tudo</button>
            </div>
          </div>
          {agrupado.map(N=>{
            const nat=NAT[N.nat]; const aberta=natAberta.has(N.nat);
            const nGrupos=N.grupos.filter(g=>!conferindo||contar(g)>0).length;
            const nItens=N.grupos.reduce((x,g)=>x+contar(g),0);
            return (
              <div className="px-natacc" key={N.nat}>
                <button className="px-nathead" onClick={()=>toggleNat(N.nat)} style={{borderLeftColor:nat.cor}}>
                  {aberta?<ChevronDown size={16}/>:<ChevronRight size={16}/>}
                  <span className="px-nathead-dot" style={{background:nat.cor}}/>
                  <span className="px-nathead-t">{nat.rot}</span>
                  <span className="px-nathead-meta">{nGrupos} {nGrupos===1?cfg.rot.toLowerCase():cfg.plural} · {nItens.toLocaleString("pt-BR")} entregas</span>
                </button>
                {aberta && <div className="px-natbody">
                  {N.grupos.map(g=>{
                    const marcado=g.itens.some(e=>selSet.has(e.codigo));
                    return (
                      <div className={`px-nivel-row ${conferindo&&!marcado?"fora":""}`} key={g.nome}>
                        {conferindo && (
                          <button className={`px-nivel-check ${marcado?"on":""}`} onClick={()=>alternarGrupo(g)}
                            title={marcado?"Tirar da Descrição da Área":"Devolver à Descrição da Área"}>
                            {marcado && <Check size={11}/>}
                          </button>
                        )}
                        <button className="px-nivel-nome" onClick={()=>descer(nivel,g.valor,g.semServico)}
                          title={`Ver ${NIVEIS[cfg.proximo]?NIVEIS[cfg.proximo].plural:"as entregas"} de "${g.nome}"`}>
                          <span className="px-dot sel" style={{background:g.semServico?C.faint:nat.cor}}/>
                          <span className="px-nivel-t" style={g.semServico?{fontStyle:"italic",color:C.faint}:undefined}>{g.nome}</span>
                          <span className="px-catcount">{g.itens.length}</span>
                          <ChevronRight size={14} className="px-nivel-go"/>
                        </button>
                      </div>
                    );
                  })}
                </div>}
              </div>
            );
          })}
        </div>
      )}

      <BuscaAdicionar selSet={selSet} add={add} onGoNova={onGoNova}/>
    </div>
  );
}

/* ---------- Relatório de Macroprocessos em PDF (sem libs, gerado na mão) ---------- */
function winAnsiByte(code){
  if(code<256) return code;
  const map={0x2014:0x97,0x2013:0x96,0x2018:0x91,0x2019:0x92,0x201C:0x93,0x201D:0x94,0x2026:0x85};
  return map[code]!=null?map[code]:0x3F;
}
function pdfEscape(str){
  str=String(str==null?"":str); let out="";
  for(let i=0;i<str.length;i++){
    const b=winAnsiByte(str.charCodeAt(i));
    out+= (b===0x28||b===0x29||b===0x5C) ? "\\"+String.fromCharCode(b) : String.fromCharCode(b);
  }
  return out;
}
function larguraTexto(str,fontSize,bold){
  return String(str==null?"":str).length*fontSize*(bold?0.56:0.5);
}
function quebrarTexto(str,fontSize,maxWidth,bold){
  str=(str==null?"":String(str)).trim(); if(!str) return [""];
  const palavras=str.split(/\s+/); const linhas=[]; let atual="";
  for(const p of palavras){
    const teste=atual?atual+" "+p:p;
    if(larguraTexto(teste,fontSize,bold)<=maxWidth) atual=teste;
    else{ if(atual) linhas.push(atual); atual=p; }
  }
  if(atual) linhas.push(atual);
  return linhas.length?linhas:[""];
}
function agruparPorNaturezaRelatorio(itens){
  const finalistica=itens.filter(r=>r.natureza==="finalistico");
  const resto=itens.filter(r=>r.natureza!=="finalistico");
  const grupos=[];
  if(finalistica.length) grupos.push({label:"Finalística",itens:finalistica});
  if(resto.length) grupos.push({label:"Suporte/Governança",itens:resto});
  return grupos;
}
function gerarPdfRelatorioMacroprocessos(resultados,orgao,unidade){
  const PAGE_W=595, PAGE_H=842, MARGIN=42;
  const usableW=PAGE_W-MARGIN*2;
  const cols=[["Macroprocesso",122],["Processo",195]];
  cols.push(["Entrega",usableW-cols.reduce((s,c)=>s+c[1],0)]);
  const TITLE_FS=22, META_FS=11, INTRO_FS=10.5, SEC_FS=17, SUB_FS=12.5, TH_FS=11.5, ROW_FS=10.5;
  const LH=ROW_FS+4.3, TH_LH=TH_FS+6, META_LH=META_FS+6.5, INTRO_LH=INTRO_FS+5.5, FOOTER_FS=9.5;
  const grupos=agruparPorNaturezaRelatorio(resultados);
  const introTexto="Este relatório mapeia os macroprocessos, processos e entregas da unidade a partir dos marcos referenciais enviados (Cadeia de Valor, Estrutura Organizacional, Relatório de Gestão Integrado e Regimento Interno), com apoio de extração automática por IA.";
  const geradoEm=new Date().toLocaleString("pt-BR");

  let pagesContent=[]; let page=[]; let y=0;
  function texto(x,yy,size,str,bold){
    page.push("BT /"+(bold?"F2":"F1")+" "+size+" Tf "+x.toFixed(2)+" "+yy.toFixed(2)+" Td ("+pdfEscape(str)+") Tj ET");
  }
  function linha(yy){
    page.push("0.7 w "+MARGIN.toFixed(2)+" "+yy.toFixed(2)+" m "+(PAGE_W-MARGIN).toFixed(2)+" "+yy.toFixed(2)+" l S");
  }
  function novaPagina(){
    if(page.length) pagesContent.push(page.join("\n"));
    page=[]; y=PAGE_H-MARGIN;
    texto(MARGIN,y-TITLE_FS,TITLE_FS,"Relatório de Macroprocessos",true); y-=TITLE_FS+6;
    texto(MARGIN,y-META_FS,META_FS,"Órgão: "+(orgao||"—")); y-=META_LH;
    texto(MARGIN,y-META_FS,META_FS,"Unidade: "+(unidade||"—")); y-=META_LH;
    texto(MARGIN,y-META_FS,META_FS,"Data de geração: "+geradoEm); y-=META_LH+4;
    const intro=quebrarTexto(introTexto,INTRO_FS,usableW,false);
    intro.forEach(ln=>{ texto(MARGIN,y-INTRO_FS,INTRO_FS,ln); y-=INTRO_LH; });
    y-=6; linha(y); y-=22;
  }
  function cabecalhoTabela(){
    let x=MARGIN;
    for(const [rotulo,w] of cols){ texto(x,y-TH_FS,TH_FS,rotulo,true); x+=w; }
    y-=TH_LH;
  }
  function garantirEspaco(h){ if(y-h<MARGIN+16) novaPagina(); }

  novaPagina();
  for(const g of grupos){
    garantirEspaco(SEC_FS+SUB_FS+TH_LH+20);
    texto(MARGIN,y-SEC_FS,SEC_FS,g.label,true); y-=SEC_FS+8;
    texto(MARGIN,y-SUB_FS,SUB_FS,unidade||"—",true); y-=SUB_FS+6;
    cabecalhoTabela();
    for(const it of g.itens){
      const cel=[
        quebrarTexto(it.macro||"",ROW_FS,cols[0][1]-6,false),
        quebrarTexto(it.categoria||"",ROW_FS,cols[1][1]-6,false),
        quebrarTexto(it.entrega||"",ROW_FS,cols[2][1]-6,false),
      ];
      const nLinhas=Math.max(cel[0].length,cel[1].length,cel[2].length);
      const altura=nLinhas*LH;
      garantirEspaco(altura);
      let x=MARGIN;
      for(let c=0;c<cols.length;c++){
        cel[c].forEach((ln,li)=>texto(x,y-ROW_FS-LH*li,ROW_FS,ln));
        x+=cols[c][1];
      }
      y-=altura;
    }
    y-=14;
  }
  pagesContent.push(page.join("\n"));

  const total=pagesContent.length;
  pagesContent=pagesContent.map((stream,i)=>{
    const rotulo=(i+1)+" / "+total;
    const w=larguraTexto(rotulo,FOOTER_FS,false);
    const x=(PAGE_W-w)/2;
    return stream+"\nBT /F1 "+FOOTER_FS+" Tf "+x.toFixed(2)+" 24.00 Td ("+pdfEscape(rotulo)+") Tj ET";
  });

  let out="%PDF-1.4\n%\xE2\xE3\xCF\xD3\n";
  const offsets=[];
  function push(s){ out+=s; }
  function iniciarObj(num){ offsets[num]=out.length; push(num+" 0 obj\n"); }
  function fecharObj(){ push("endobj\n"); }
  const totalPages=pagesContent.length;
  const pageObjNums=[], contentObjNums=[];
  const fontBoldNum=4; let nextNum=5;
  for(let p=0;p<totalPages;p++){ pageObjNums.push(nextNum++); contentObjNums.push(nextNum++); }

  iniciarObj(1); push("<< /Type /Catalog /Pages 2 0 R >>\n"); fecharObj();
  iniciarObj(2); push("<< /Type /Pages /Kids ["+pageObjNums.map(n=>n+" 0 R").join(" ")+"] /Count "+totalPages+" >>\n"); fecharObj();
  iniciarObj(3); push("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\n"); fecharObj();
  iniciarObj(fontBoldNum); push("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>\n"); fecharObj();

  for(let q=0;q<totalPages;q++){
    iniciarObj(pageObjNums[q]);
    push("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 "+PAGE_W+" "+PAGE_H+"] /Resources << /Font << /F1 3 0 R /F2 "+fontBoldNum+" 0 R >> >> /Contents "+contentObjNums[q]+" 0 R >>\n");
    fecharObj();
    const stream=pagesContent[q];
    iniciarObj(contentObjNums[q]);
    push("<< /Length "+stream.length+" >>\nstream\n"+stream+"\nendstream\n");
    fecharObj();
  }
  const xrefStart=out.length;
  const totalObjs=nextNum-1;
  push("xref\n0 "+(totalObjs+1)+"\n0000000000 65535 f \n");
  for(let n=1;n<=totalObjs;n++) push(String(offsets[n]).padStart(10,"0")+" 00000 n \n");
  push("trailer\n<< /Size "+(totalObjs+1)+" /Root 1 0 R >>\nstartxref\n"+xrefStart+"\n%%EOF");

  const bytes=new Uint8Array(out.length);
  for(let bi=0;bi<out.length;bi++) bytes[bi]=out.charCodeAt(bi)&255;
  const blob=new Blob([bytes],{type:"application/pdf"});
  const url=URL.createObjectURL(blob);
  const lk=document.createElement("a");
  lk.href=url;
  lk.download="relatorio-macroprocessos-"+(unidade||"unidade").replace(/[^a-z0-9]+/gi,"-").toLowerCase().slice(0,40)+".pdf";
  document.body.appendChild(lk); lk.click(); lk.remove();
  setTimeout(()=>URL.revokeObjectURL(url),60000);
}

function DocHeaderEdit({orgao,unidade,setOrgao,setUnidade}){
  const [edit,setEdit]=useState(false);
  const [o,setO]=useState(orgao); const [u,setU]=useState(unidade);
  useEffect(()=>{ setO(orgao); },[orgao]);
  useEffect(()=>{ setU(unidade); },[unidade]);
  const salvar=()=>{ setOrgao(o.trim()); setUnidade(u.trim()); setEdit(false); };
  if(edit) return (<div className="px-doc-meta edit">
    <label>Órgão<input value={o} onChange={e=>setO(e.target.value)} placeholder="Nome do órgão" autoFocus/></label>
    <label>Unidade<input value={u} onChange={e=>setU(e.target.value)} placeholder="Nome da unidade"/></label>
    <div className="px-doc-meta-btns">
      <button className="px-dm-save" onClick={salvar}><Check size={13}/> Salvar</button>
      <button className="px-dm-cancel" onClick={()=>{setO(orgao);setU(unidade);setEdit(false);}}>Cancelar</button>
    </div>
  </div>);
  if(!orgao && !unidade) return (
    <button className="px-doc-setmeta" onClick={()=>setEdit(true)}>
      <Building2 size={14}/> <span>Informar órgão e unidade</span> <Pencil size={12}/>
    </button>
  );
  return (<div className="px-doc-meta">
    <div><span>Órgão</span>{orgao||"—"}</div>
    <div><span>Unidade</span>{unidade||"—"}</div>
    <button className="px-doc-meta-edit" onClick={()=>setEdit(true)} title="Editar órgão e unidade"><Pencil size={12}/></button>
  </div>);
}

function DescPanel({sel,notes,setNote,rem,onInject,orgao,unidade}){
  const groups=NAT_ORDER.map(nat=>({nat,itens:sel.filter(e=>!e.nova&&e.natureza===nat)})).filter(g=>g.itens.length);
  const novas=sel.filter(e=>e.nova);
  if(novas.length) groups.push({nat:"__novas__",itens:novas});
  const dist=NAT_ORDER.map(nat=>({nat,n:sel.filter(e=>!e.nova&&e.natureza===nat).length}));
  let idx=0;
  return (<>
    <div className="px-doc-list">
      {sel.length===0
        ? <div className="px-doc-guide">
            <div className="px-doc-guide-t">Como montar a descrição da área</div>
            <ol className="px-doc-steps">
              <li><span>1</span><div><b>Filtre por área de atuação</b>busca, natureza ou macroprocesso</div></li>
              <li><span>2</span><div><b>Selecione as entregas</b>marque as que a unidade executa</div></li>
              <li><span>3</span><div><b>Registre no DFT</b>ou exporte a planilha para a unidade</div></li>
            </ol>
          </div>
        : groups.map(g=>{ const ehNova=g.nat==="__novas__"; const n=ehNova?null:NAT[g.nat]; return (
          <div className="px-doc-group" key={g.nat}>
            <div className="px-doc-gh">{ehNova?<><Sparkles size={11} color={C.primary}/> Entregas novas (aguardam curadoria)</>:<><span className="px-doc-dot" style={{background:n.cor}}/>{n.rot}</>}<span className="px-doc-gc">{g.itens.length}</span></div>
            {g.itens.map(e=>{ idx++; const has=notes[e.codigo]!==undefined; return (
              <div className={`px-doc-item ${e.nova?"nova":""}`} key={e.codigo}>
                <span className="px-doc-n">{String(idx).padStart(2,"0")}</span>
                <div className="px-doc-body">
                  <div className="px-doc-row"><code>{e.nova?"nova":e.codigo}</code>{e.nova && <span className="px-doc-novatag" title="Entrega proposta — passará pela curadoria do banco antes de entrar no catálogo"><Sparkles size={9}/> nova · aguarda curadoria</span>}</div>
                  <div className="px-doc-name">{e.entrega}</div>
                  {has
                    ? <div className="px-doc-notebox"><textarea value={notes[e.codigo]} placeholder="Observação que vai para o DFT…" onChange={ev=>setNote(e.codigo,ev.target.value)}/></div>
                    : <button className="px-doc-addnote" onClick={()=>setNote(e.codigo,"")}><Plus size={11}/> Observação para o DFT</button>}
                </div>
                <button className="px-doc-del" onClick={()=>rem(e.codigo)} title="Remover"><Trash2 size={13}/></button>
              </div>); })}
          </div>); })}
    </div>
    <div className="px-doc-foot">
      {sel.length>0 && <div className="px-doc-dist">
        <span className="px-doc-dist-tot"><b>{sel.length}</b> entrega{sel.length===1?"":"s"}</span>
        <span className="px-doc-dist-sep"/>
        {dist.map(d=>{ const n=NAT[d.nat]; return (
          <span key={d.nat} className={`px-doc-dist-n ${d.n===0?"zero":""}`} title={n.rot}><i style={{background:n.cor}}/>{d.n} {n.rot.toLowerCase()}</span>
        ); })}
      </div>}
      <button className="px-inject" disabled={!sel.length} onClick={onInject}>
        <ShieldCheck size={16}/> Registrar no DFT {sel.length>0 && <em>({sel.length})</em>}
      </button>
      <button className="px-export" disabled={!sel.length} onClick={()=>exportarDescricaoXLSX(sel,notes,orgao,unidade)}>
        <Download size={15}/> Baixar planilha (.xlsx)
      </button>
    </div>
  </>);
}

/* ---------- skeleton de carga ---------- */
function SkeletonCat(){
  return (<div className="px-skel">
    {[0,1,2,3,4].map(i=>(<div className="px-skel-cat" key={i}>
      <div className="px-skel-head"><span className="px-skel-bar w40"/><span className="px-skel-pill"/></div>
      {i<2 && [0,1,2].map(j=><div className="px-skel-row" key={j}><span className="px-skel-box"/><span className="px-skel-bar w20"/><span className="px-skel-bar w60"/></div>)}
    </div>))}
  </div>);
}
function BancoLoading(){
  return (<div className="px-bancoload">
    <div className="px-bancoload-sp"/>
    <div className="px-bancoload-t">Carregando o banco de entregas…</div>
    <div className="px-bancoload-s">31.241 entregas · 696 categorias · 217 macroprocessos</div>
    <SkeletonCat/>
  </div>);
}

/* ============================================================================
   CAMINHO DE 4 ETAPAS — é a navegação principal do portal. Sai do cabeçalho
   ("Catálogo" e "Órgãos" não estão mais lá) e vira uma faixa logo abaixo dele:
   cards completos no Início, tira compacta nas demais telas.
============================================================================ */
const ETAPAS=[
  {n:1, ic:Bot,        t:"Unidade e marcos referenciais",
   d:"Cadastre a unidade e anexe os documentos institucionais — a IA identifica o que a área faz."},
  {n:2, ic:LayoutGrid, t:"Catálogo de entregas",
   d:"Confira o que a IA encontrou e desça de macroprocesso até a entrega, escolhendo o que é da unidade."},
  {n:3, ic:Users,      t:"Jornada do gestor",
   d:"Acompanhe entregas, equipe e o retorno do período. O planejamento é opcional."},
  {n:4, ic:Building2,  t:"Resultados do órgão",
   d:"Organograma, dimensionamento e o relatório de cada unidade."},
];

function FaixaEtapas({etapa,completo,expandida,ir}){
  if(expandida) return null;           // no Início os cards grandes já cumprem esse papel
  return (
    <nav className="px-etapas-tira" aria-label="Etapas do portal">
      {ETAPAS.map(e=>{
        const Ic=e.ic, feito=!!completo[e.n], ativo=etapa===e.n;
        return (
          <button key={e.n} className={`px-etapa-min ${ativo?"on":""} ${feito?"feito":""}`}
            onClick={()=>ir(e.n)} title={e.d}>
            <span className="px-etapa-min-n">{feito?<Check size={11}/>:e.n}</span>
            <Ic size={13}/>
            <span className="px-etapa-min-t">{e.t}</span>
          </button>
        );
      })}
    </nav>
  );
}

/* modal de cadastro/troca da unidade — acionado pelo chip do cabeçalho */
function ModalUnidade({orgao,unidade,setOrgao,setUnidade,onClose}){
  const [o,setO]=useState(orgao||"");
  const [u,setU]=useState(unidade||"");
  const salvar=()=>{ setOrgao(o.trim()); setUnidade(u.trim()); onClose(); };
  useEffect(()=>{ const h=ev=>{ if(ev.key==="Escape") onClose(); };
    window.addEventListener("keydown",h); return ()=>window.removeEventListener("keydown",h); },[onClose]);
  return (
    <div className="px-modal-bg" onClick={onClose}>
      <div className="px-unid-modal" onClick={ev=>ev.stopPropagation()}>
        <div className="px-unid-h"><Building2 size={16}/> <b>Órgão e unidade</b>
          <button className="px-unid-x" onClick={onClose}><X size={15}/></button></div>
        <p className="px-unid-d">Fica salvo neste computador — você não precisa informar de novo na próxima visita.</p>
        <div className="px-caminho-form">
          <label>Órgão<input value={o} onChange={e=>setO(e.target.value)} placeholder="Ex.: Ministério da Gestão e da Inovação" autoFocus/></label>
          <label>Unidade<input value={u} onChange={e=>setU(e.target.value)} placeholder="Ex.: Coordenação-Geral de Recursos Humanos" onKeyDown={e=>{ if(e.key==="Enter") salvar(); }}/></label>
        </div>
        <div className="px-caminho-foot">
          <button className="px-caminho-skip" onClick={onClose}>Cancelar</button>
          <button className="px-caminho-go" onClick={salvar}><Check size={15}/> Salvar</button>
        </div>
      </div>
    </div>
  );
}

/* ---------- Início — os 4 cards do caminho ---------- */
function SecaoInicioCaminho({stats,orgao,unidade,completo,ir}){
  const fmt=n=>Number(n||0).toLocaleString("pt-BR");
  return (
    <section className="px-caminho">
      <div className="px-caminho-hero">
        <h1>O trabalho do serviço público federal,<br/>organizado em um só lugar.</h1>
        <p>{(orgao||unidade)
          ? <>Caminho de <b>{unidade||orgao}</b>{unidade&&orgao ? <> — {orgao}</> : null}</>
          : "Quatro etapas, do cadastro da unidade ao relatório do órgão."}</p>
        <div className="px-home-nums">
          <span><b>{fmt(stats.total)}</b> entregas</span><i/>
          <span><b>{fmt(stats.servs)}</b> serviços</span><i/>
          <span><b>{fmt(stats.macros)}</b> macroprocessos</span>
        </div>
      </div>

      <div className="px-etapas-grid">
        {ETAPAS.map(e=>{
          const Ic=e.ic, feito=!!completo[e.n];
          return (
            <button className={`px-etapa-card ${feito?"feito":""}`} key={e.n} onClick={()=>ir(e.n)}>
              <span className="px-etapa-card-top">
                <span className="px-etapa-card-n">{feito?<Check size={13}/>:e.n}</span>
                <span className="px-etapa-card-ic"><Ic size={18}/></span>
              </span>
              <b>{e.t}</b>
              <span className="px-etapa-card-d">{e.d}</span>
              <span className="px-etapa-card-go">{feito?"Revisar":"Abrir"} <ArrowRight size={13}/></span>
            </button>
          );
        })}
      </div>

      <div className="px-home-foot">SGP · MGI — SIGEPE · SISDIP/DFT</div>
    </section>
  );
}

/* ---------- capa de abertura (clara e serena) ---------- */
/* ---------- Início — porta de entrada do portal (seção, não overlay; versão em cartões, mantida por
   compatibilidade — foi substituída pelo caminho guiado acima e não é mais renderizada) ---------- */
function SecaoInicio({stats,onCatalogo,onOrgaos,onChat,onConversor,onRevisao,onPGD,onMarcos}){
  const fmt=n=>Number(n||0).toLocaleString("pt-BR");
  const caminhos=[
    {ic:LayoutGrid, cor:C.primary, soft:C.primarySoft, t:"Catálogo de serviços",  d:"Navegue, busque e monte a descrição da sua área a partir das entregas oficiais.", onClick:onCatalogo, destaque:true},
    {ic:Wand2,      cor:C.green,   soft:C.greenSoft,   t:"Conversor de entregas",  d:"Traga seu PGD, regimento interno, planilha ou planejamento estratégico — a IA enquadra cada item no catálogo, com a confiança à mostra.", onClick:onConversor, destaque:true},
    {ic:ShieldCheck,cor:C.green,   soft:C.greenSoft,   t:"Curadoria do banco",    d:"Classifique serviços e una duplicatas — qualidade do catálogo.", onClick:onRevisao},
    {ic:Bot,        cor:C.primary, soft:C.primarySoft, t:"Assistente de Entregas",d:"Descreva o que sua área faz — a IA sugere as entregas.", onClick:onChat},
    {ic:Bot,        cor:"#0C7B93", soft:"#E2F1F4",     t:"Marcos Referenciais",  d:"Anexe a Cadeia de Valor, Estrutura Organizacional, Relatório de Gestão Integrado e Regimento Interno — a IA identifica os macroprocessos, processos e serviços da unidade.", onClick:onMarcos},
    {ic:Upload,     cor:"#5B3A9B", soft:"#F0EBF8",     t:"Importar PGD",         d:"A esteira de conversão: cada registro do PGD vira entrega do catálogo.", onClick:onPGD},
    {ic:Building2,  cor:C.navy,    soft:"#E7EDF7",     t:"Visão por órgão",       d:"Organograma, status do DFT e as entregas de cada unidade federal.", onClick:onOrgaos},
  ];
  return (
    <section className="px-home">
      <div className="px-home-hero">
        <h1>O trabalho do serviço público federal,<br/>organizado em um só lugar.</h1>
        <p>Catálogo oficial</p>
        <div className="px-home-nums">
          <span><b>{fmt(stats.total)}</b> entregas</span><i/>
          <span><b>{fmt(stats.servs)}</b> serviços</span><i/>
          <span><b>{fmt(stats.macros)}</b> macroprocessos</span>
        </div>
      </div>
      <div className="px-home-grid">
        {caminhos.map((c,i)=>{ const Ic=c.ic; return (
          <button className={`px-home-card ${c.destaque?"grande":""}`} key={i} onClick={c.onClick}>
            <span className="px-home-card-ic" style={{background:c.soft,color:c.cor}}><Ic size={20}/></span>
            <b>{c.t}</b>
            <span className="px-home-card-d">{c.d}</span>
            <span className="px-home-card-go" style={{color:c.cor}}>Abrir <ArrowRight size={13}/></span>
          </button>); })}
      </div>
      <div className="px-home-foot">SGP · MGI — SIGEPE · SISDIP/DFT</div>
    </section>
  );
}

/* ---------- capa antiga (mantida por compatibilidade; não renderizada) ---------- */
function Capa({stats,onEntrar,onChat,onConversor,onSunburst,onOrgaos,onRevisao}){
  const fmt=n=>Number(n||0).toLocaleString("pt-BR");
  useEffect(()=>{ const h=e=>{ if(e.key==="Escape") onEntrar(); }; window.addEventListener("keydown",h); return ()=>window.removeEventListener("keydown",h); },[onEntrar]);
  const caminhos=[
    {ic:LayoutGrid, cor:C.primary,  soft:C.primarySoft, t:"Explorar o catálogo", d:"Navegue por natureza, macroprocesso, categoria e serviço.", onClick:onEntrar,   delay:".08s"},
    {ic:Sun,        cor:"#B86E00",  soft:"#FBEEDB",     t:"Explosão solar",      d:"O catálogo inteiro em um só panorama visual.",              onClick:onSunburst, delay:".14s"},
    {ic:Bot,        cor:C.primary,  soft:C.primarySoft, t:"Assistente de Entregas", d:"Descreva o que sua área faz — a IA enquadra no catálogo.", onClick:onChat,    delay:".2s"},
    {ic:Wand2,      cor:C.green,    soft:C.greenSoft,   t:"Conversor",           d:"Do PGD, planilha ou PDF direto para o DFT.",                 onClick:onConversor,delay:".26s"},
    {ic:Building2,  cor:C.navy,     soft:"#EAEFF7",     t:"Visão por órgão",     d:"Organograma, unidades e as entregas de cada uma.",           onClick:onOrgaos,   delay:".32s"},
    {ic:ShieldCheck,cor:C.green,    soft:C.greenSoft,   t:"Curadoria do banco",  d:"Ajude a classificar serviços e unir duplicatas.",            onClick:onRevisao,  delay:".38s"},
  ];
  return (
    <div className="px-capa-bg" role="dialog" aria-label="Capa do Portal das Entregas">
      <div className="px-capa">
        <button className="px-capa-x" onClick={onEntrar} title="Fechar a capa"><X size={17}/></button>
        <div className="px-capa-top">
          <svg viewBox="0 0 32 24" width="46" height="35"><circle cx="7" cy="7" r="4" fill={C.yellow}/><circle cx="16" cy="6" r="4.4" fill={C.green}/><circle cx="25" cy="7" r="4" fill={C.primary}/><path d="M2 22c1.5-5 9-5 10.5 0z" fill={C.yellow}/><path d="M10.5 22c1.5-6 9.5-6 11 0z" fill={C.green}/><path d="M20 22c1.5-5 9-5 10.5 0z" fill={C.primary}/></svg>
          <h1>Portal das Entregas</h1>
          <p className="px-capa-tag">O trabalho do serviço público federal, organizado em um só catálogo — para consultar, descrever e dimensionar.</p>
          <div className="px-capa-nums">
            <span><b>{fmt(stats.total)}</b> entregas</span><i/>
            <span><b>{fmt(stats.servs)}</b> serviços</span><i/>
            <span><b>{fmt(stats.macros)}</b> macroprocessos</span>
          </div>
        </div>
        <div className="px-capa-grid">
          {caminhos.map((c,i)=>{ const Ic=c.ic; return (
            <button className="px-capa-card" key={i} style={{animationDelay:c.delay}} onClick={c.onClick}>
              <span className="px-capa-card-ic" style={{background:c.soft,color:c.cor}}><Ic size={19}/></span>
              <b>{c.t}</b>
              <span className="px-capa-card-d">{c.d}</span>
            </button>); })}
        </div>
        <div className="px-capa-foot">
          <span className="px-capa-org">SGP · MGI — SIGEPE · SISDIP/DFT</span>
          <button className="px-capa-entrar" onClick={onEntrar}>Entrar no portal <ArrowRight size={16}/></button>
        </div>
      </div>
    </div>
  );
}

/* ---------- onboarding (boas-vindas, 3 caminhos) ---------- */
function Onboarding({onChat,onPGD,onCatalog,onOrgaos,onRevisao,onSearch,onClose}){
  const [dont,setDont]=useState(false);
  const [hq,setHq]=useState("");
  const close=fn=>()=>{ fn(); };
  return (<div className="px-ob-bg">
    <div className="px-ob">
      <button className="px-ob-x" onClick={onClose}><X size={18}/></button>
      <div className="px-ob-logo">
        <svg viewBox="0 0 32 24" width="34" height="26"><circle cx="7" cy="7" r="4" fill={C.yellow}/><circle cx="16" cy="6" r="4.4" fill={C.green}/><circle cx="25" cy="7" r="4" fill={C.primary}/><path d="M2 22c1.5-5 9-5 10.5 0z" fill={C.yellow}/><path d="M10.5 22c1.5-6 9.5-6 11 0z" fill={C.green}/><path d="M20 22c1.5-5 9-5 10.5 0z" fill={C.primary}/></svg>
      </div>
      <h2>Bem-vindo ao Portal das Entregas</h2>
      <p>Por onde você quer começar? Você pode trocar de caminho a qualquer momento.</p>
      <div className="px-ob-cards">
        <button className="px-ob-card" onClick={onChat}>
          <span className="px-ob-ic" style={{background:C.primarySoft,color:C.primary}}><Bot size={22}/></span>
          <b>Me ajude a descrever</b>
          <span>Ajuda interativa: descreva o que sua área faz e a IA enquadra nas entregas do catálogo.</span>
          <span className="px-ob-go">Conversar com a IA <ArrowRight size={14}/></span>
        </button>
        <button className="px-ob-card" onClick={onPGD}>
          <span className="px-ob-ic" style={{background:"#E3F2E5",color:C.green}}><Upload size={22}/></span>
          <b>Transforme seu PGD em DFT</b>
          <span>Importe as entregas do PGD, confira a correspondência sugerida e registre no DFT.</span>
          <span className="px-ob-go">Abrir conversor <ArrowRight size={14}/></span>
        </button>
        <button className="px-ob-card" onClick={onCatalog}>
          <span className="px-ob-ic" style={{background:"#FBEEDB",color:"#B86E00"}}><LayoutGrid size={22}/></span>
          <b>Consultar o banco</b>
          <span>Navegue e busque no banco de entregas e monte sua descrição de área entrega por entrega.</span>
          <span className="px-ob-go">Explorar catálogo <ArrowRight size={14}/></span>
        </button>
      </div>
      <label className="px-ob-dont"><input type="checkbox" checked={dont} onChange={e=>setDont(e.target.checked)}/> Não mostrar esta tela novamente</label>
    </div>
  </div>);
}

/* ---------- menu lateral: navegação por macroprocesso ---------- */
function MacroNav({natF,macroF,onPick,onClear,onClose}){
  const tree=useMemo(()=>{
    const byNat={};
    ENTREGAS.forEach(e=>{ (byNat[e.natureza]=byNat[e.natureza]||{}); byNat[e.natureza][e.macro]=(byNat[e.natureza][e.macro]||0)+1; });
    return NAT_ORDER.filter(n=>byNat[n]).map(nat=>({ nat, macros:Object.entries(byNat[nat]).map(([mac,n])=>({mac,n})).sort((a,b)=>b.n-a.n) }));
  },[]);
  const [open,setOpen]=useState(()=>new Set(natF?[natF]:NAT_ORDER));
  const toggle=k=>setOpen(p=>{const n=new Set(p);n.has(k)?n.delete(k):n.add(k);return n;});
  return (<aside className="px-mnav">
    <div className="px-mnav-h"><span><Menu size={15}/> Macroprocessos</span><button onClick={onClose}><X size={16}/></button></div>
    <button className={`px-mnav-all ${!macroF?"on":""}`} onClick={onClear}>Todos os macroprocessos</button>
    <div className="px-mnav-body">
      {tree.map(g=>{ const n=NAT[g.nat]; const isOpen=open.has(g.nat); return (
        <div className="px-mnav-grp" key={g.nat}>
          <button className="px-mnav-nat" onClick={()=>toggle(g.nat)}>
            {isOpen?<ChevronDown size={13}/>:<ChevronRight size={13}/>}
            <span className="px-dot" style={{background:n.cor}}/>{n.rot}<span className="px-mnav-c">{g.macros.length}</span>
          </button>
          {isOpen && <div className="px-mnav-list">
            {g.macros.map(m=>(
              <button key={m.mac} className={`px-mnav-item ${macroF===m.mac?"on":""}`} onClick={()=>onPick(g.nat,m.mac)} title={m.mac}>
                <span>{m.mac}</span><b>{m.n}</b>
              </button>
            ))}
          </div>}
        </div>); })}
    </div>
  </aside>);
}

/* ---------- painel genérico de filtro por lista (categoria/serviço/órgão) ---------- */
function ListaFiltroNav({titulo,icon:Icon,itens,valorAtual,onPick,onClear,onClose,corDe,semBusca}){
  const [busca,setBusca]=useState("");
  const filtrados=useMemo(()=>{ const q=norm(busca); return q?itens.filter(it=>norm(it.label||it.valor).includes(q)):itens; },[itens,busca]);
  return (<aside className="px-mnav">
    <div className="px-mnav-h"><span><Icon size={15}/> {titulo}</span><button onClick={onClose}><X size={16}/></button></div>
    <button className={`px-mnav-all ${!valorAtual?"on":""}`} onClick={onClear}>Todos</button>
    {!semBusca && <div className="px-mnav-search"><Search size={13} color={C.faint}/><input value={busca} onChange={e=>setBusca(e.target.value)} placeholder="Buscar…"/></div>}
    <div className="px-mnav-body">
      <div className="px-mnav-list">
        {filtrados.map(it=>(
          <button key={it.valor} className={`px-mnav-item ${valorAtual===it.valor?"on":""}`} onClick={()=>onPick(it.valor)} title={it.label||it.valor}>
            {corDe && <span className="px-dot" style={{background:corDe(it.valor)}}/>}
            <span>{it.label||it.valor}</span><b>{it.n}</b>
          </button>
        ))}
        {filtrados.length===0 && <div className="px-mnav-empty">Nada encontrado.</div>}
      </div>
    </div>
  </aside>);
}

/* ---------- comparação rápida (até 3) ---------- */
function CompareModal({items,onClose,onAdd,selSet,onDrop,onClear}){
  const cols=items.map(e=>({e,det:getDetails(e.codigo)}));
  const linhas=[["Entrega","entrega"],["Atividade","atividade"],["Natureza","_nat"],["Macroprocesso","macro"],["Categoria","categoria"],["Serviço","servico"],["Fluxo de trabalho","_fluxo"],["Fonte de medição","_fonte"]];
  const val=(c,key)=>{ if(key==="_nat") return NAT[c.e.natureza]?NAT[c.e.natureza].rot:c.e.natureza; if(key==="_fluxo") return c.det.fluxo||"—"; if(key==="_fonte") return c.det.fonte||"—"; return c.e[key]||"—"; };
  return (<div className="px-modal-bg" onClick={onClose}>
    <div className="px-cmp" onClick={e=>e.stopPropagation()}>
      <div className="px-cmp-h"><div><Columns size={18} color={C.primary}/> <b>Comparar entregas</b><span>Decida lado a lado — adicione direto ou volte ao catálogo sem escolher.</span></div><button className="px-x" onClick={onClose}><X size={18}/></button></div>
      <div className="px-cmp-grid" style={{gridTemplateColumns:`150px repeat(${cols.length},1fr)`}}>
        <div className="px-cmp-cell px-cmp-rowh"></div>
        {cols.map(c=>{ const n=NAT[c.e.natureza]; const ok=selSet.has(c.e.codigo); return (
          <div className="px-cmp-cell px-cmp-colh" key={c.e.codigo}>
            <div className="px-cmp-top"><code style={{color:n.cor}}>{c.e.codigo}</code><button className="px-cmp-drop" onClick={()=>onDrop(c.e.codigo)} title="Tirar da comparação"><X size={13}/></button></div>
            <button className={`px-cmp-add ${ok?"done":""}`} onClick={()=>!ok&&onAdd(c.e)} disabled={ok}>{ok?<><Check size={13}/> Na descrição</>:<><Plus size={13}/> Adicionar</>}</button>
          </div>); })}
        {linhas.map(([lbl,key])=>(<React.Fragment key={key}>
          <div className="px-cmp-cell px-cmp-rowh">{lbl}</div>
          {cols.map(c=><div className="px-cmp-cell" key={c.e.codigo+key}>{val(c,key)}</div>)}
        </React.Fragment>))}
      </div>
      <div className="px-cmp-foot"><button className="px-btn-ghost" onClick={onClear}>Limpar comparação</button><button className="px-btn-primary" onClick={onClose}>Voltar ao catálogo</button></div>
    </div>
  </div>);
}

/* ---------- detalhe de entrega (fluxo/fonte carregados sob demanda) ---------- */
function EntregaDetalhe({e,inCmp,toggleCompare,onFlag}){
  const [d,setD]=useState(()=>getDetails(e.codigo));
  const [carregando,setCarregando]=useState(!DET_CACHE);
  useEffect(()=>{ let vivo=true;
    if(!DET_CACHE){ ensureDetalhes().then(()=>{ if(vivo){ setD(getDetails(e.codigo)); setCarregando(false); } }); }
    else { setD(getDetails(e.codigo)); }
    return ()=>{vivo=false;};
  },[e.codigo]);
  return (<div className="px-edet">
    {carregando
      ? <div className="px-edet-load">Carregando detalhes…</div>
      : <>
        {d.fluxo && <div className="px-edet-section"><b>Fluxo de trabalho</b><p>{d.fluxo}</p></div>}
        {d.fonte && <div className="px-edet-section"><b>Fonte de medição</b><p>{d.fonte}</p></div>}
      </>}
    <div className="px-edet-grid">
      <span><b>Atividade</b>{e.atividade}</span>
      <span><b>Metodologia</b>{e.metod||"Típica"}</span>
      <span><b>Natureza</b>{NAT[e.natureza]?NAT[e.natureza].rot:e.natureza}</span>
      <span><b>Macroprocesso</b>{e.macro}</span>
    </div>
    <div className="px-edet-credits">
      <span><b>Órgão</b>{e.orgao||"Banco de entregas"}</span>
      <span><b>Unidade de origem</b>{e.unidade||"—"}</span>
    </div>
    <div className="px-edet-acts">
      <button className={`px-cmplink ${inCmp?"on":""}`} onClick={()=>toggleCompare(e)}><Columns size={12}/> {inCmp?"Na comparação":"Comparar"}</button>
      <button className="px-flagbtn" onClick={()=>onFlag(e)}><AlertTriangle size={12}/> Sinalizar problema ao banco</button>
    </div>
  </div>);
}

/* ---------- painel de números (topo) ---------- */
function PainelNumeros({stats,bancoStatus,filtro}){
  const fmt=n=>Number(n||0).toLocaleString("pt-BR");
  const {total,macros,cats,servs,orgs,pctServ,dist}=stats;
  const seg=k=> total? (dist[k]/total*100):0;
  const Big = filtro
    ? <button className="px-mx-big px-mx-btn" onClick={filtro.onTotal} title="Ver todas as entregas">{fmt(total)} <em>entregas</em></button>
    : <span className="px-mx-big">{fmt(total)} <em>entregas</em></span>;
  const Mini = ({onClick,active,title,children}) => filtro
    ? <button className={`px-mx-mini px-mx-btn ${active?"on":""}`} onClick={onClick} title={title}>{children}</button>
    : <span className="px-mx-mini">{children}</span>;
  return (
    <section className="px-mx" aria-label="Números do banco de entregas">
      <div className="px-mx-lead">
        <span className="px-mx-lead-ic"><Layers size={19}/></span>
        <div className="px-mx-main">
          <div className="px-mx-nums">
            <span className="px-mx-eyebrow">Banco de entregas</span>
            <div className="px-mx-row">
              {Big}
              {filtro && <Mini onClick={filtro.onNatureza} active={filtro.ativoNatureza} title="Filtrar por natureza">Natureza</Mini>}
              <Mini onClick={filtro&&filtro.onMacro} active={filtro&&filtro.ativoMacro} title="Filtrar por macroprocesso"><b>{fmt(macros)}</b> macroprocessos</Mini>
              <Mini onClick={filtro&&filtro.onCategoria} active={filtro&&filtro.ativoCategoria} title="Filtrar por categoria"><b>{fmt(cats)}</b> categorias</Mini>
              <Mini onClick={filtro&&filtro.onServico} active={filtro&&filtro.ativoServico} title="Filtrar por serviço"><b>{fmt(servs)}</b> serviços</Mini>
              {orgs>0 && <Mini onClick={filtro&&filtro.onOrgao} active={filtro&&filtro.ativoOrgao} title="Filtrar por órgão"><b>{fmt(orgs)}</b> órgãos</Mini>}
            </div>
          </div>
          <div className="px-mx-bar" role="img" aria-label="Distribuição por natureza">
            {NAT_ORDER.map(k=> seg(k)>0 ? <span key={k} style={{width:seg(k)+"%",background:NAT[k].cor}} title={`${NAT[k].rot}: ${fmt(dist[k])}`}/> : null)}
          </div>
        </div>
      </div>
      {bancoStatus==="fallback" && <div className="px-mx-note">Amostra de demonstração — em produção, o banco completo traz 31.241 entregas.</div>}
    </section>
  );
}

/* ---------- MODO 1: catálogo expansível por categoria ---------- */
function ListaEnriquecida({res,searching,query,selSet,add,rem,exp,setExp,onFlag,onPropose,compare,toggleCompare,justAdded}){
  // hierarquia: natureza -> macroprocesso -> categoria de serviço -> entregas
  const arvore=useMemo(()=>{
    const natM=new Map();
    res.forEach(e=>{
      if(!natM.has(e.natureza)) natM.set(e.natureza,new Map());
      const macM=natM.get(e.natureza);
      if(!macM.has(e.macro)) macM.set(e.macro,new Map());
      const catM=macM.get(e.macro);
      if(!catM.has(e.categoria)) catM.set(e.categoria,[]);
      catM.get(e.categoria).push(e);
    });
    return NAT_ORDER.filter(nat=>natM.has(nat)).map(nat=>{
      const macM=natM.get(nat);
      const macros=[...macM].map(([mac,catM])=>{
        const cats=[...catM].map(([cat,itens])=>({cat,itens})).sort((a,b)=>a.cat.localeCompare(b.cat));
        const tot=cats.reduce((s,c)=>s+c.itens.length,0);
        return {mac,cats,tot};
      }).sort((a,b)=>a.mac.localeCompare(b.mac));
      const tot=macros.reduce((s,m)=>s+m.tot,0);
      return {nat,macros,tot};
    });
  },[res]);

  // expansão independente por nível. Naturezas começam RECOLHIDAS.
  const [openNat,setOpenNat]=useState(()=>new Set());
  const [openMac,setOpenMac]=useState(()=>new Set());
  const [openCat,setOpenCat]=useState(()=>new Set());
  const tgl=(set,fn)=>k=>fn(p=>{const n=new Set(p); n.has(k)?n.delete(k):n.add(k); return n;});
  const toggleNat=tgl(openNat,setOpenNat), toggleMac=tgl(openMac,setOpenMac), toggleCat=tgl(openCat,setOpenCat);
  const cmpSet=new Set((compare||[]).map(e=>e.codigo));
  const totCat=useMemo(()=>new Set(res.map(e=>e.categoria)).size,[res]);

  const expandirTudo=()=>{ const nn=new Set(),mm=new Set(),cc=new Set();
    arvore.forEach(N=>{ nn.add(N.nat); N.macros.forEach(M=>{ mm.add(N.nat+"|"+M.mac); M.cats.forEach(C=>cc.add(N.nat+"|"+M.mac+"|"+C.cat)); }); });
    setOpenNat(nn); setOpenMac(mm); setOpenCat(cc);
  };
  const recolherTudo=()=>{ setOpenNat(new Set()); setOpenMac(new Set()); setOpenCat(new Set()); };

  if(!arvore.length){
    const sugest=(()=>{ if(!query||!query.trim()) return [];
      const toks=norm(query).split(/\s+/).filter(t=>t.length>=4);
      const found=new Set();
      for(const t of toks){ for(const e of ENTREGAS){ if(norm(e.categoria).includes(t.slice(0,5))){ found.add(e.categoria); if(found.size>=5) break; } } if(found.size>=5) break; }
      return [...found].slice(0,5);
    })();
    return (<div className="px-empty2">
      <div className="px-empty2-ic"><Search size={26}/></div>
      <h3>Nenhuma entrega encontrada{query?<> para “<b>{query}</b>”</>:null}</h3>
      <p>Tente termos mais gerais, sinônimos, ou um código (ex.: 0305). Se de fato não existe, proponha — vai para a curadoria do banco.</p>
      {sugest.length>0 && <div className="px-empty2-sug">
        <span>Categorias parecidas:</span>
        {sugest.map(s=><span className="px-empty2-tag" key={s}>{s.length>40?s.slice(0,39)+"…":s}</span>)}
      </div>}
      <button className="px-empty2-btn" onClick={()=>onPropose(query)}><Plus size={15}/> Propor nova entrega ao banco</button>
    </div>);
  }

  // durante busca ativa, tudo expande para revelar os resultados
  const naturezaAberta=N=>searching||openNat.has(N.nat);
  const macroAberto=(N,M)=>searching||openMac.has(N.nat+"|"+M.mac);
  const catAberta=(N,M,C)=>searching||openCat.has(N.nat+"|"+M.mac+"|"+C.cat);

  return (<div className="px-tree">
    <div className="px-treebar">
      <div className="px-treebar-l"><b>Catálogo de serviços</b> <span>· {arvore.length} naturezas · {totCat} categorias · {res.length.toLocaleString("pt-BR")} entregas</span></div>
      <div className="px-treebar-r">
        <button onClick={expandirTudo}>Expandir tudo</button><span>/</span>
        <button onClick={recolherTudo}>Recolher tudo</button>
      </div>
    </div>

    {arvore.map(N=>{ const n=NAT[N.nat]; const natOpen=naturezaAberta(N);
      const nSelNat=(()=>{let s=0;N.macros.forEach(M=>M.cats.forEach(C=>C.itens.forEach(e=>{if(selSet.has(e.codigo))s++;})));return s;})();
      return (
      <div className="px-natacc" key={N.nat}>
        <button className="px-nathead" onClick={()=>toggleNat(N.nat)} style={{borderLeftColor:n.cor}}>
          {natOpen?<ChevronDown size={16}/>:<ChevronRight size={16}/>}
          <span className="px-nathead-dot" style={{background:n.cor}}/>
          <span className="px-nathead-t">{n.rot}</span>
          <span className="px-nathead-meta">{N.macros.length} macroprocessos · {N.tot.toLocaleString("pt-BR")} entregas</span>
          {nSelNat>0 && <span className="px-catsel"><Check size={9}/> {nSelNat}</span>}
        </button>

        {natOpen && <div className="px-natbody">
          {N.macros.map(M=>{ const macOpen=macroAberto(N,M); const macKey=N.nat+"|"+M.mac;
            const nSelMac=(()=>{let s=0;M.cats.forEach(C=>C.itens.forEach(e=>{if(selSet.has(e.codigo))s++;}));return s;})();
            return (
            <div className="px-macacc" key={macKey}>
              <button className="px-machead" onClick={()=>toggleMac(macKey)}>
                {macOpen?<ChevronDown size={14}/>:<ChevronRight size={14}/>}
                <GitBranch size={13} color={n.cor}/>
                <span className="px-machead-t">{M.mac}</span>
                <span className="px-machead-meta">{M.cats.length} categorias · {M.tot}</span>
                {nSelMac>0 && <span className="px-catsel sm"><Check size={9}/> {nSelMac}</span>}
              </button>

              {macOpen && <div className="px-macbody">
                {M.cats.map(C=>{ const catOpen=catAberta(N,M,C); const catKey=N.nat+"|"+M.mac+"|"+C.cat;
                  const nSel=C.itens.filter(e=>selSet.has(e.codigo)).length;
                  return (
                  <div className="px-catacc" key={catKey}>
                    <button className="px-cathead" onClick={()=>toggleCat(catKey)}>
                      {catOpen?<ChevronDown size={13}/>:<ChevronRight size={13}/>}
                      <span className={`px-dot${nSel>0?" sel":""}`} style={{background:n.cor}}/>
                      <span className="px-cathead-t">{C.cat}</span>
                      <span className="px-catcount">{C.itens.length}</span>
                      {nSel>0 && <span className="px-catsel"><Check size={9}/> {nSel}</span>}
                    </button>
                    {catOpen && <div className="px-erows">
                      {C.itens.map(e=>{ const ok=selSet.has(e.codigo); const op=exp===e.codigo; const sit=SIT[e.sit]; const inCmp=cmpSet.has(e.codigo); return (
                        <div className={`px-erow ${ok?"sel":""}`} key={e.codigo}>
                          <div className="px-erow-main">
                            <button className={`px-cbox ${ok?"on":""}`} onClick={()=>ok?rem(e.codigo):add(e)} title={ok?"Remover":"Adicionar à descrição"}>{ok&&<Check size={12}/>}</button>
                            <code className="px-ecode">{e.codigo}</code>
                            <span className="px-sit sm" style={{background:sit.soft,color:sit.cor}}>{sit.rot}</span>
                            <span className="px-ename">{e.entrega}</span>
                            <span className="px-eserv">{e.servico}</span>
                            {justAdded===e.codigo && <span className="px-added"><Check size={11}/> adicionada</span>}
                            <div className="px-erow-acts">
                              <button className={`px-cmpbtn ${inCmp?"on":""}`} onClick={()=>toggleCompare(e)} title={inCmp?"Na comparação":"Comparar"}><Columns size={13}/></button>
                              <button className="px-edetbtn" onClick={()=>setExp(op?null:e.codigo)} title="Informações gerais"><Info size={14}/></button>
                            </div>
                          </div>
                          {op && <EntregaDetalhe e={e} inCmp={inCmp} toggleCompare={toggleCompare} onFlag={onFlag}/>}
                        </div>); })}
                    </div>}
                  </div>); })}
              </div>}
            </div>); })}
        </div>}
      </div>); })}
  </div>);
}

/* ---------- MODO 2: por processo ---------- */
function PorProcesso({natF,selSet,add,rem}){
  const macros=useMemo(()=>{
    const arr=[];
    BANCO_RAW.forEach(n=>{
      if(natF&&n.nat!==natF) return;
      n.macros.forEach(m=>{
        const cats=m.cats.map(c=>({nome:c.cat,servicos:[{nome:c.cat,itens:c.itens.map(it=>({codigo:it.cod,entrega:it.ent,natureza:n.nat}))}]}));
        arr.push({nome:m.mac,natureza:n.nat,categorias:cats});
      });
    });
    return arr;
  },[natF]);
  return (<div className="px-proc">
    {macros.map(m=>{ const n=NAT[m.natureza]; return (
      <div className="px-proc-macro" key={m.nome}>
        <div className="px-proc-h"><span className="px-dot" style={{background:n.cor}}/><b>{m.nome}</b><span className="px-tag" style={{background:n.soft,color:n.cor}}>{n.rot}</span></div>
        <div className="px-proc-flow">
          {m.categorias.flatMap(c=>c.servicos).map((s,si,arr)=>(
            <React.Fragment key={s.nome}>
              <div className="px-stage">
                <div className="px-stage-t">{s.nome}</div>
                {s.itens.slice(0,12).map(e=>{ const ok=selSet.has(e.codigo); const full=ENTREGAS.find(x=>x.codigo===e.codigo)||e; return (
                  <button key={e.codigo} className={`px-stage-item ${ok?"ok":""}`} onClick={()=>ok?rem(e.codigo):add(full)}>
                    {ok?<Check size={12}/>:<Plus size={12}/>}<code>{e.codigo}</code><span>{e.entrega}</span>
                  </button>); })}
              </div>
              {si<arr.length-1 && <div className="px-stage-arrow"><ArrowRight size={18}/></div>}
            </React.Fragment>
          ))}
        </div>
      </div>); })}
  </div>);
}

/* ---------- MODO 5: cadeia de valor + camada de objetivos ---------- */
function trunc2(s,n){ s=s||""; return s.length>n?s.slice(0,n-1)+"…":s; }
function CadeiaValor({sel,rem,objUnidade,setObjUnidade,objOrgao,setObjOrgao,onClose}){
  const [focusObj,setFocusObj]=useState(null);
  const [order,setOrder]=useState({}); // natureza -> [macro,...] (ordem escolhida pela pessoa)
  const [objOpen,setObjOpen]=useState(false); // camada de objetivos retrátil (opcional)
  const abrirObjetivos=()=>setObjOpen(v=>!v);
  const fechar=()=>{ setFocusObj(null); onClose&&onClose(); };
  // trava o scroll do body enquanto a cadeia (tela cheia) está aberta
  useEffect(()=>{ const prev=document.body.style.overflow; document.body.style.overflow="hidden"; return ()=>{document.body.style.overflow=prev;}; },[]);
  // ESC fecha
  useEffect(()=>{ const h=e=>{ if(e.key==="Escape") fechar(); }; window.addEventListener("keydown",h); return ()=>window.removeEventListener("keydown",h); },[]);

  // códigos atualmente selecionados (para intersecção com os vínculos dos objetivos)
  const selCods=useMemo(()=>new Set(sel.map(e=>e.codigo)),[sel]);

  // (sem auto-distribuição: objetivos começam vazios e são criados pela pessoa)

  // mapa cod -> objetivos a que pertence (restrito aos selecionados)
  const codToObjs=useMemo(()=>{
    const m=new Map();
    objUnidade.forEach(o=>o.cods.forEach(c=>{ if(selCods.has(c)){ if(!m.has(c)) m.set(c,[]); m.get(c).push(o.id); } }));
    return m;
  },[objUnidade,selCods]);

  // cadeia: por natureza -> macros presentes -> entregas (macros na ordem escolhida pela pessoa)
  const faixas=useMemo(()=>NAT_ORDER.map(nat=>{
    const itens=sel.filter(e=>e.natureza===nat);
    const macMap=new Map();
    itens.forEach(e=>{ if(!macMap.has(e.macro)) macMap.set(e.macro,[]); macMap.get(e.macro).push(e); });
    let arr=[...macMap].map(([mac,ents])=>({mac,ents}));
    const ord=order[nat];
    if(ord&&ord.length){
      arr.sort((a,b)=>{ const ia=ord.indexOf(a.mac), ib=ord.indexOf(b.mac); return (ia<0?1e9:ia)-(ib<0?1e9:ib); });
    }
    return { nat, total:itens.length, macros:arr };
  }).filter(f=>f.macros.length),[sel,order]);

  // move um macroprocesso para a esquerda/direita dentro da sua faixa
  const moveMacro=(nat,mac,dir)=>{
    const cur=(faixas.find(f=>f.nat===nat)?.macros||[]).map(m=>m.mac);
    const i=cur.indexOf(mac), j=i+dir;
    if(i<0||j<0||j>=cur.length) return;
    const next=[...cur]; [next[i],next[j]]=[next[j],next[i]];
    setOrder(prev=>({...prev,[nat]:next}));
  };

  const focObj=objUnidade.find(o=>o.id===focusObj)||null;
  const focCods=focObj?new Set(focObj.cods.filter(c=>selCods.has(c))):null;
  // macros iluminados pelo objetivo em foco
  const macrosAcesos=useMemo(()=>{
    if(!focCods) return new Set();
    const s=new Set();
    sel.forEach(e=>{ if(focCods.has(e.codigo)) s.add(e.macro); });
    return s;
  },[focCods,sel]);

  const setObjField=(id,patch)=>setObjUnidade(prev=>prev.map(o=>o.id===id?{...o,...patch}:o));
  const addObj=()=>{ const id=nextObjId(); setObjUnidade(prev=>[...prev,{id,titulo:"",orgaoId:objOrgao[0]?.id||"",cods:[]}]); setFocusObj(null); };
  const delObj=id=>{ setObjUnidade(prev=>prev.filter(o=>o.id!==id)); if(focusObj===id) setFocusObj(null); };
  // objetivos do órgão (editáveis)
  const setOrgField=(id,patch)=>setObjOrgao(prev=>prev.map(o=>o.id===id?{...o,...patch}:o));
  const addOrg=()=>{ const id=nextOrgId(); setObjOrgao(prev=>[...prev,{id,titulo:""}]); };
  const delOrg=id=>{ setObjOrgao(prev=>prev.filter(o=>o.id!==id)); setObjUnidade(prev=>prev.map(o=>o.orgaoId===id?{...o,orgaoId:""}:o)); };
  const toggleCod=(cod)=>{ if(!focObj) return; setObjUnidade(prev=>prev.map(o=>{ if(o.id!==focObj.id) return o; const has=o.cods.includes(cod); return {...o,cods:has?o.cods.filter(c=>c!==cod):[...o.cods,cod]}; })); };
  const countVinc=o=>o.cods.filter(c=>selCods.has(c)).length;
  const orgaoRot=id=>{ const o=objOrgao.find(x=>x.id===id); return o?`${o.id} · ${o.titulo}`:"—"; };

  if(!sel.length) return (
    <div className="px-cv-overlay" onClick={fechar}>
      <div className="px-cv-sheet" onClick={e=>e.stopPropagation()}>
        <div className="px-cv-fullbar">
          <div className="px-cv-fullbar-t"><Workflow size={16} color={C.primary}/> <b>Cadeia de valor</b></div>
          <button className="px-cv-fullclose" onClick={fechar}><X size={16}/> Fechar</button>
        </div>
        <div className="px-cv px-cv-full">
          <div className="px-cv-empty">
            <Workflow size={30} color={C.faint}/>
            <h3>A cadeia de valor se monta com as entregas que você seleciona</h3>
            <p>Volte ao modo <b>Lista</b> e selecione as entregas da sua área. Aqui você verá em quais etapas (macroprocessos) elas atuam e como se conectam aos objetivos da unidade e do órgão.</p>
          </div>
        </div>
      </div>
    </div>
  );

  const corpo = (
    <div className="px-cv px-cv-full">
      {/* barra de topo (sempre em tela cheia) */}
      <div className="px-cv-fullbar">
        <div className="px-cv-fullbar-t"><Workflow size={16} color={C.primary}/> <b>Cadeia de valor</b><span>{sel.length} entregas</span></div>
        <button className="px-cv-fullclose" onClick={fechar}><X size={16}/> Fechar</button>
      </div>

      {/* ----- camada de objetivos (opcional, retrátil) ----- */}
      <div className={`px-cv-obj ${objOpen?"open":""}`}>
        <button className="px-cv-obj-head" onClick={abrirObjetivos}>
          <Target size={15} color={C.primary}/> <b>Vincular a objetivos</b>
          <span>opcional · conecte os objetivos da unidade aos do órgão e às entregas</span>
          {objUnidade.length>0 && <em className="px-cv-obj-count">{objUnidade.length}</em>}
          {objOpen?<ChevronDown size={16}/>:<ChevronRight size={16}/>}
        </button>

        {objOpen && <div className="px-cv-obj-body">
        <div className="px-cv-org">
          <span className="px-cv-org-l">Objetivos do órgão</span>
          {objOrgao.map(o=>{ const aceso=focObj&&focObj.orgaoId===o.id; return (
            <span key={o.id} className={`px-cv-orgpill editable ${aceso?"on":""}`}>
              <b>{o.id}</b>
              <input value={o.titulo} onChange={e=>setOrgField(o.id,{titulo:e.target.value})} placeholder="Objetivo do órgão"/>
              <button className="px-cv-orgdel" title="Remover objetivo do órgão" onClick={()=>delOrg(o.id)}><X size={11}/></button>
            </span>
          ); })}
          <button className="px-cv-orgadd" onClick={addOrg}><Plus size={13}/> Objetivo do órgão</button>
        </div>

        <div className="px-cv-uni">
          <span className="px-cv-org-l">Unidade</span>
          <div className="px-cv-uni-cards">
            {objUnidade.map(o=>{ const on=focusObj===o.id; const n=countVinc(o); return (
              <div key={o.id} className={`px-cv-ocard ${on?"on":""}`}>
                <div className="px-cv-ocard-top">
                  <input className="px-cv-otitle" value={o.titulo} onChange={e=>setObjField(o.id,{titulo:e.target.value})} placeholder="Objetivo da unidade…" autoFocus={!o.titulo}/>
                  <button className="px-cv-odel" title="Remover objetivo" onClick={()=>delObj(o.id)}><Trash2 size={13}/></button>
                </div>
                <div className="px-cv-ocard-link">
                  <span>Contribui para</span>
                  <select value={o.orgaoId} onChange={e=>setObjField(o.id,{orgaoId:e.target.value})}>
                    <option value="">— objetivo do órgão —</option>
                    {objOrgao.map(x=><option key={x.id} value={x.id}>{x.id} · {trunc2(x.titulo||"(sem título)",30)}</option>)}
                  </select>
                </div>
                <div className="px-cv-ocard-foot">
                  <span className="px-cv-ocount">{n} entrega{n===1?"":"s"}</span>
                  <button className={`px-cv-ofocus ${on?"on":""}`} onClick={()=>setFocusObj(on?null:o.id)}>{on?"Concluir":"Vincular entregas"}</button>
                </div>
              </div>
            ); })}
            <button className="px-cv-oadd" onClick={addObj}><Plus size={15}/> Novo objetivo</button>
          </div>
        </div>

        {focObj && <div className="px-cv-hint"><Info size={13}/> Vinculando <b>{trunc2(focObj.titulo,40)}</b> — clique nas entregas na cadeia para incluí-las ou removê-las. <button onClick={()=>setFocusObj(null)}>Concluir</button></div>}
        </div>}
      </div>

      {/* ----- cadeia de valor (faixas por natureza) ----- */}
      <div className="px-cv-chain">
        {faixas.map(f=>{ const n=NAT[f.nat]; const fin=f.nat==="finalistico"; return (
          <div className="px-cv-lane" key={f.nat} style={{background:n.soft}}>
            <div className="px-cv-lane-l" style={{color:n.cor}}><span className="px-cv-dot" style={{background:n.cor}}/>{n.rot}<i>{f.total}</i></div>
            <div className="px-cv-lane-flow">
              {f.macros.map((m,mi)=>{ const aceso=macrosAcesos.has(m.mac); const dim=focObj&&!aceso; return (
                <React.Fragment key={m.mac}>
                  <div className={`px-cv-macro ${aceso?"aceso":""} ${dim?"dim":""}`} style={aceso?{borderColor:C.primaryDark}:null}>
                    <div className="px-cv-macro-h" title={m.mac}>
                      <span>{trunc2(m.mac,40)}</span>
                      <div className="px-cv-macro-ord">
                        <button disabled={mi===0} onClick={()=>moveMacro(f.nat,m.mac,-1)} title="Mover para a esquerda"><ChevronLeft size={13}/></button>
                        <button disabled={mi===f.macros.length-1} onClick={()=>moveMacro(f.nat,m.mac,1)} title="Mover para a direita"><ChevronRight size={13}/></button>
                      </div>
                      <i style={{background:n.cor}}>{m.ents.length}</i>
                    </div>
                    <div className="px-cv-chips">
                      {m.ents.map(e=>{ const objs=codToObjs.get(e.codigo)||[]; const linked=focCods&&focCods.has(e.codigo); const dimc=focObj&&!linked;
                        return (
                          <button key={e.codigo} className={`px-cv-chip ${linked?"linked":""} ${dimc?"dimc":""} ${focObj?"clickable":""}`} title={focObj?"Clique para vincular/desvincular":e.entrega} onClick={()=>toggleCod(e.codigo)}>
                            <code>{e.codigo}</code><span>{trunc2(e.entrega,46)}</span>
                            {!focObj && objs.length>0 && <em className="px-cv-mark" title={`Em ${objs.length} objetivo(s)`}>{objs.length}</em>}
                          </button>
                        ); })}
                    </div>
                  </div>
                  {fin && mi<f.macros.length-1 && <div className="px-cv-arrow"><ChevronRight size={20}/></div>}
                </React.Fragment>
              ); })}
            </div>
          </div>
        ); })}
        <div className="px-cv-cap"><Info size={12}/> No Finalístico, as setas indicam um fluxo de valor ilustrativo — não uma sequência rígida entre macroprocessos. Mostramos apenas as etapas em que as entregas selecionadas atuam. Use ‹ › no topo de cada macroprocesso para reordená-los na sequência que fizer sentido para a sua área.</div>
      </div>
    </div>
  );

  return (
    <div className="px-cv-overlay" onClick={fechar}>
      <div className="px-cv-sheet" onClick={e=>e.stopPropagation()}>{corpo}</div>
    </div>
  );
}

/* ---------- MODO 3: panorama (treemap interativo com drill-down) ---------- */
/* ---------- MODO 2: Árvore de decomposição (fluida, estilo decomposition tree) ---------- */
const ARV_NIVEIS=[
  {key:"natureza", rot:"Natureza"},
  {key:"macro",    rot:"Macroprocesso"},
  {key:"categoria",rot:"Categoria de serviço"},
  {key:"servico",  rot:"Serviço"},
  {key:"entrega",  rot:"Entrega"},
];
function ArvoreDecomposicao({res,add,rem,selSet}){
  // caminho selecionado por nível: {natureza, macro, categoria, servico}
  const [path,setPath]=useState({});
  // colunas recolhidas manualmente pelo usuário (chaves de nível). Natureza tem regra automática à parte.
  const [recolhidasManuais,setRecolhidasManuais]=useState(()=>new Set());
  const recolher=(key)=>setRecolhidasManuais(p=>{ const n=new Set(p); n.add(key); return n; });
  const reabrir=(key)=>setRecolhidasManuais(p=>{ const n=new Set(p); n.delete(key); return n; });
  const niveisAtivos=useMemo(()=>{
    // sempre mostra o nível 0; cada nível seguinte aparece quando o anterior foi escolhido
    const out=[]; 
    for(let i=0;i<ARV_NIVEIS.length;i++){
      out.push(ARV_NIVEIS[i]);
      if(i<ARV_NIVEIS.length-1 && !path[ARV_NIVEIS[i].key]) break;
    }
    return out;
  },[path]);

  // dados de uma coluna: agrupa 'res' filtrado pelos níveis anteriores, contando pelo campo do nível
  const colunaDados=(nivelIdx)=>{
    const campo=ARV_NIVEIS[nivelIdx].key;
    let filtrado=res;
    for(let j=0;j<nivelIdx;j++){ const k=ARV_NIVEIS[j].key; const v=path[k]; if(v!=null) filtrado=filtrado.filter(e=>valNivel(e,k)===v); }
    if(campo==="entrega"){
      // folhas: cada entrega é um item selecionável
      return filtrado.map(e=>({nome:e.entrega,valor:1,entrega:e})).sort((a,b)=>a.nome.localeCompare(b.nome));
    }
    const m=new Map();
    filtrado.forEach(e=>{ const v=valNivel(e,campo); m.set(v,(m.get(v)||0)+1); });
    return [...m].map(([nome,valor])=>({nome,valor})).sort((a,b)=>b.valor-a.valor);
  };
  function valNivel(e,k){ return k==="natureza"?e.natureza : k==="macro"?e.macro : k==="categoria"?e.categoria : k==="servico"?(e.servico||e.categoria) : e.entrega; }

  const escolher=(nivelIdx,nome)=>{
    const campo=ARV_NIVEIS[nivelIdx].key;
    setPath(prev=>{
      const np={...prev};
      // define este nível e limpa os posteriores
      np[campo]=prev[campo]===nome?undefined:nome;
      for(let j=nivelIdx+1;j<ARV_NIVEIS.length;j++) delete np[ARV_NIVEIS[j].key];
      if(np[campo]===undefined) delete np[campo];
      return np;
    });
    // reabre automaticamente as colunas recolhidas dos níveis a partir deste (a árvore mudou de rumo)
    setRecolhidasManuais(prev=>{ const n=new Set(prev);
      for(let j=nivelIdx;j<ARV_NIVEIS.length;j++) n.delete(ARV_NIVEIS[j].key);
      return n; });
  };
  const toggleEntrega=e=>{ if(!e) return; if(selSet.has(e.codigo)) rem(e.codigo); else add(e); };

  const rotNivel=(k,v)=> k==="natureza"?(NAT[v]?NAT[v].rot:v):v;
  const corNivel=(nivelIdx)=>{
    const nat=path.natureza; const base=nat&&NAT[nat]?NAT[nat].cor:C.primary;
    return base;
  };

  return (<div className="px-arv">
    <div className="px-arv-h">
      <div><h3>Árvore de decomposição</h3><p>Clique em um bloco para abrir o próximo nível — Natureza › Macroprocesso › Categoria › Serviço › Entrega. No último nível, clique para adicionar à descrição de área. A busca filtra a árvore.</p></div>
      <div className="px-arv-crumb">
        <button className={Object.keys(path).length===0?"on":""} onClick={()=>setPath({})}>Tudo</button>
        {ARV_NIVEIS.map((nv,i)=> path[nv.key]!=null && (
          <span key={nv.key} className="px-arv-crumb-i">
            <ChevronRight size={12}/>
            <button onClick={()=>escolher(i,path[nv.key])} title={rotNivel(nv.key,path[nv.key])}>{trunc2(rotNivel(nv.key,path[nv.key]),22)}</button>
          </span>
        ))}
      </div>
    </div>

    <div className="px-arv-flow">
      {(()=>{ const entregaVisivel = niveisAtivos.some(nv=>nv.key==="entrega"); return niveisAtivos.map((nv,i)=>{
        const escolhido = path[nv.key]!=null;
        const ultimoAtivo = i===niveisAtivos.length-1;
        // AUTOMÁTICO: só a Natureza, e só quando a coluna Entrega já está aberta.
        const autoRecolher = nv.key==="natureza" && entregaVisivel;
        // MANUAL: qualquer coluna já escolhida que o usuário optou por recolher (e ainda não reabriu).
        const manualRecolher = escolhido && recolhidasManuais.has(nv.key);
        const recolhida = escolhido && (autoRecolher || manualRecolher);
        if(recolhida){
          const cor=corNivel(i);
          return (
            <button className="px-arv-col recolhida" key={nv.key} style={{animationDelay:`${i*60}ms`,"--cor":cor}}
              onClick={()=>reabrir(nv.key)} title={`${nv.rot}: ${rotNivel(nv.key,path[nv.key])} — clique para reabrir`}>
              <span className="px-arv-rec-lb">{nv.rot}</span>
              <span className="px-arv-rec-v">{rotNivel(nv.key,path[nv.key])}</span>
              <ChevronRight size={13} className="px-arv-rec-go"/>
            </button>
          );
        }
        const dados=colunaDados(i);
        const max=Math.max(1,...dados.map(d=>d.valor));
        const cor=corNivel(i);
        const total=dados.reduce((s,d)=>s+d.valor,0);
        const isEntrega=nv.key==="entrega";
        return (
        <div className={`px-arv-col${isEntrega?" larga":""}`} key={nv.key} style={{animationDelay:`${i*60}ms`}}>
          <div className="px-arv-col-h">
            <span>{nv.rot}</span>
            <em>{isEntrega?`${dados.length} entregas`:`${dados.length}`}</em>
            {escolhido && <button className="px-arv-col-min" title={`Recolher ${nv.rot}`} onClick={()=>recolher(nv.key)}><Minus size={13}/></button>}
          </div>
          <div className="px-arv-col-body">
            {dados.length===0 && <div className="px-arv-vazio">Sem itens</div>}
            {dados.slice(0,400).map((d,di)=>{
              const isLeaf=nv.key==="entrega";
              const sel=isLeaf&&d.entrega&&selSet.has(d.entrega.codigo);
              const ativo=!isLeaf&&path[nv.key]===d.nome;
              const pct=Math.round((d.valor/max)*100);
              return (
                <button key={di}
                  className={`px-arv-node ${ativo?"ativo":""} ${sel?"sel":""} ${isLeaf?"leaf":""}`}
                  style={{"--barw":pct+"%","--cor":cor}}
                  onClick={()=>isLeaf?toggleEntrega(d.entrega):escolher(i,d.nome)}
                  title={rotNivel(nv.key,d.nome)}>
                  <span className="px-arv-node-bar"/>
                  <span className="px-arv-node-tx">{rotNivel(nv.key,d.nome)}</span>
                  {!isLeaf && <span className="px-arv-node-val">{d.valor.toLocaleString("pt-BR")}</span>}
                  {isLeaf && <span className="px-arv-node-pick">{sel?<Check size={13}/>:<Plus size={13}/>}</span>}
                  {!isLeaf && <ChevronRight size={14} className="px-arv-node-go"/>}
                </button>
              );
            })}
            {dados.length>400 && <div className="px-arv-mais">+{(dados.length-400).toLocaleString("pt-BR")} — refine a busca para ver todos</div>}
          </div>
        </div>);
      }); })()}
    </div>
  </div>);
}

/* ---------- preview do que será gravado no DFT ---------- */
function PreviewDFT({sel,notes,orgao,unidade,onClose,onConfirm}){
  const comObs=sel.filter(e=>notes&&notes[e.codigo]&&notes[e.codigo].trim()).length;
  return (<div className="px-modal-bg" onClick={onClose}>
    <div className="px-modal" onClick={e=>e.stopPropagation()}>
      <div className="px-modal-h">
        <div><ShieldCheck size={18} color={C.primary}/> <b>Pré-visualização do envio ao DFT</b></div>
        <button onClick={onClose}><X size={18}/></button>
      </div>
      <p className="px-modal-sub">Confira exatamente o que será gravado. O gestor confirma e assume — nada vai ao sistema sem revisão.</p>
      <div className="px-modal-meta"><span><b>Órgão:</b> {orgao}</span><span><b>Unidade:</b> {unidade}</span><span><b>Entregas:</b> {sel.length}</span><span><b>Com observação:</b> {comObs}</span></div>
      <div className="px-modal-table">
        <div className="px-mt-head"><span>Código</span><span>Entrega</span><span>Observação (DFT)</span></div>
        {sel.map(e=>{ const ob=notes&&notes[e.codigo]; return (
          <div className="px-mt-row" key={e.codigo}><code>{e.codigo}</code><span>{e.entrega}</span>
            <span className="px-mt-obs">{ob&&ob.trim()?ob:<i>—</i>}</span>
          </div>); })}
      </div>
      <div className="px-modal-acts">
        <button className="px-btn-ghost" onClick={onClose}>Voltar e ajustar</button>
        <button className="px-btn-primary" onClick={onConfirm}><Send size={14}/> Confirmar e enviar ao DFT</button>
      </div>
    </div>
  </div>);
}

/* ---------- assistente (encaixe por IA) ---------- */
function Assistente({onAdd}){
  // inline: renderiza como visualização do .px-main (mesma área da Lista/Árvore/
  // Explosão solar). Padrão (flutuante): FAB + janela, controlados por open/setOpen.
  const [msgs,setMsgs]=useState([{role:"assistant",content:"Sou o Assistente de Entregas. Descreva o que sua área faz, cole trechos do PGD/planejamento, ou anexe uma planilha (.xlsx/.csv), PDF ou Word com a descrição de área — eu enquadro no catálogo (com o código), aviso se já existe algo parecido e proponho novas quando preciso. Os códigos viram botões de adicionar."}]);
  const [input,setInput]=useState(""); const [load,setLoad]=useState(false); const endRef=useRef(null);
  const [anexo,setAnexo]=useState(null);      // {tipo,nome,texto?,pdfB64?}
  const [anexoErro,setAnexoErro]=useState("");
  const fileRef=useRef(null);
  const codMap=useMemo(()=>{const m=new Map();ENTREGAS.forEach(e=>m.set(e.codigo,e));return m;},[]);
  useEffect(()=>{endRef.current?.scrollIntoView({behavior:"smooth"});},[msgs,open,load]);
  // FAB arrastável (segurar e mover); clique simples abre o chat
  const [fabPos,setFabPos]=useState(null);
  const fabDrag=useRef({on:false,moved:false,dx:0,dy:0});
  function fabDown(e){ const r=e.currentTarget.getBoundingClientRect(); fabDrag.current={on:true,moved:false,dx:e.clientX-r.left,dy:e.clientY-r.top}; try{e.currentTarget.setPointerCapture(e.pointerId);}catch(_){} }
  function fabMove(e){ const d=fabDrag.current; if(!d.on) return; if((e.movementX||0)||(e.movementY||0)) d.moved=true;
    setFabPos({left:Math.max(6,Math.min((window.innerWidth||1200)-58,e.clientX-d.dx)), top:Math.max(6,Math.min((window.innerHeight||800)-44,e.clientY-d.dy))}); }
  function fabUp(){ const d=fabDrag.current; d.on=false; if(!d.moved) setOpen(true); }

  async function escolherArquivo(ev){ const f=ev.target.files?.[0]; if(!f) return; setAnexoErro("");
    try{ const a=await lerArquivo(f); setAnexo(a); }
    catch(err){ setAnexoErro(err.message||"Não consegui ler o arquivo."); }
    finally{ if(fileRef.current) fileRef.current.value=""; }
  }

  async function send(){ const t=input.trim(); if((!t&&!anexo)||load) return;
    // monta a mensagem do usuário (texto + resumo do anexo) p/ exibição
    const rotulo=anexo?` 📎 ${anexo.nome}`:"";
    const userView=(t||(anexo?`Analisar arquivo anexado${rotulo}`:""))+(t&&anexo?rotulo:"");
    const nv=[...msgs,{role:"user",content:userView}];
    setMsgs(nv); setInput(""); setLoad(true);
    const anexoAtual=anexo; setAnexo(null);
    try{
      // fase 1 — candidatas do banco completo, com base no texto do usuário (e do anexo, se houver)
      const baseBusca=[t, anexoAtual?.texto||""].join(" ");
      const cands=buscarCandidatas(baseBusca, 50);
      const catalogo=cands.length?candidatasTxt(cands):CAT_TXT;
      const sys=`Você é o Assistente de Entregas do SISDIP/DFT (setor público federal). Funções: (1) enquadrar o trabalho descrito (ou o conteúdo de planilhas/PGD/planejamento/documentos enviados) em entregas do catálogo, citando o código no formato 0000.0000; (2) checar similaridade e evitar duplicidade; (3) propor nova entrega no padrão "nome no particípio | atividade | macroprocesso > serviço" quando não houver correspondência. Baseie-se SOMENTE nas entregas candidatas abaixo (as mais relevantes do banco para esta consulta). Quando o usuário enviar uma planilha, processe linha a linha. Conciso, cite sempre os códigos.\nEntregas candidatas (código | entrega | macro > categoria):\n${catalogo}`;
      // conteúdo da última mensagem do usuário: texto + anexo (texto extraído ou PDF nativo)
      let lastContent;
      if(anexoAtual?.tipo==="pdf"){
        lastContent=[{type:"document",source:{type:"base64",media_type:"application/pdf",data:anexoAtual.pdfB64}},{type:"text",text:t||"Enquadre as descrições de área deste PDF no catálogo, citando os códigos."}];
      } else if(anexoAtual?.texto){
        lastContent=`${t?t+"\n\n":""}Conteúdo do arquivo "${anexoAtual.nome}":\n${anexoAtual.texto}`;
      } else { lastContent=t; }
      const apiMsgs=nv.slice(0,-1).map(m=>({role:m.role,content:m.content})).concat([{role:"user",content:lastContent}]);
      const d=await chamarIA({model:"claude-haiku-4-5-20251001",max_tokens:1500,system:sys,messages:apiMsgs});
      const tx=(d.content||[]).map(b=>b.type==="text"?b.text:"").join("\n").trim();
      setMsgs(m=>[...m,{role:"assistant",content:tx||"Não consegui responder agora."}]);
    }catch{
      setMsgs(m=>[...m,{role:"assistant",content:"⚠️ Não consegui falar com o serviço de IA agora. Se o portal está publicado, verifique se a função de IA (backend) está configurada. O arquivo foi lido normalmente e você pode montar a descrição de área manualmente pelo catálogo."}]);
    } finally{ setLoad(false); }
  }
  const render=content=>{ const cs=[...new Set(content.match(CODIGO_RX)||[])].filter(c=>codMap.has(c));
    return (<><div className="px-msg-tx">{content}</div>{cs.length>0&&<div className="px-chips">{cs.map(c=><button key={c} className="px-chip2" onClick={()=>onAdd(codMap.get(c))}><Plus size={11}/> {c}</button>)}</div>}</>); };
  const corpo = (<>
      <div className="px-chat-b">
        {msgs.map((m,i)=><div key={i} className={`px-msg ${m.role}`}>{m.role==="assistant"?render(m.content):<div className="px-msg-tx">{m.content}</div>}</div>)}
        {load&&<div className="px-msg assistant"><div className="px-typing"><span/><span/><span/></div></div>}
        <div ref={endRef}/>
      </div>
      {(anexo||anexoErro) && <div className="px-chat-anexo">
        {anexo && <span className="px-anexo-chip"><FileText size={12}/> {anexo.nome} <button onClick={()=>setAnexo(null)}><X size={11}/></button></span>}
        {anexoErro && <span className="px-anexo-erro"><AlertTriangle size={12}/> {anexoErro}</span>}
      </div>}
      <div className="px-chat-i">
        <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv,.pdf,.docx" style={{display:"none"}} onChange={escolherArquivo}/>
        <button className="px-chat-clip" onClick={()=>fileRef.current?.click()} title="Anexar planilha, PDF ou Word"><Paperclip size={16}/></button>
        <textarea value={input} onChange={e=>setInput(e.target.value)} rows={1} onKeyDown={e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();send();}}} placeholder="Descreva, cole trechos, ou anexe um arquivo…"/>
        <button onClick={send} disabled={load||(!input.trim()&&!anexo)}><Send size={15}/></button>
      </div>
  </>);
  // O assistente vive apenas como visualização do Catálogo (mode==="assistente").
  // O antigo botão flutuante (FAB) foi removido — a porta de entrada é a barra de modos.
  return (<div className="px-chat inline">
    <div className="px-chat-h"><div className="px-chat-ht"><Bot size={16}/> Assistente de Entregas</div><span className="px-chat-h-sub">Descreva sua área ou anexe PGD/planilha/PDF — a IA enquadra no catálogo.</span></div>
    {corpo}
  </div>);
}

/* ============================================================================
   IMPORTAR PGD — esteira de conversão PGD -> DFT.
   Porte fiel (visual e funcionamento) do protótipo conversor-pgd-dft.html para
   dentro deste arquivo: nada de iframe, nada de HTML externo. Vive como uma
   visualização da barra de modos do Catálogo (mode==="pgd"), ao lado de Árvore,
   Explosão solar, Assistente e Conversor.
   Dados fictícios, como no protótipo — a engine real (UFRJ/LABORe) opera sobre
   o Banco de Entregas completo.
============================================================================ */
const CPD_CATALOGO = {
  "00014210": {nome:"Edital de licitação elaborado e publicado", cat:"Gestão de Contratações", macro:"Gestão de Logística e Aquisições"},
  "00014255": {nome:"Contrato administrativo fiscalizado", cat:"Gestão de Contratações", macro:"Gestão de Logística e Aquisições"},
  "00021033": {nome:"Processo de aposentadoria analisado e instruído", cat:"Concessão de Direitos e Benefícios", macro:"Gestão de Pessoas"},
  "00021047": {nome:"Benefício de pensão concedido", cat:"Concessão de Direitos e Benefícios", macro:"Gestão de Pessoas"},
  "00033801": {nome:"Painel gerencial desenvolvido e publicado", cat:"Gestão da Informação e Dados", macro:"Governança Digital"},
  "00033812": {nome:"Base de dados estruturada e disponibilizada", cat:"Gestão da Informação e Dados", macro:"Governança Digital"},
  "00027390": {nome:"Capacitação de servidores realizada", cat:"Desenvolvimento de Pessoas", macro:"Gestão de Pessoas"},
  "00027365": {nome:"Plano de desenvolvimento de equipe elaborado", cat:"Desenvolvimento de Pessoas", macro:"Gestão de Pessoas"},
  "00018904": {nome:"Nota técnica de política de pessoal elaborada", cat:"Formulação de Diretrizes de Gestão", macro:"Governança e Estratégia"},
  "00018922": {nome:"Norma interna revisada e publicada", cat:"Formulação de Diretrizes de Gestão", macro:"Governança e Estratégia"},
  "00041002": {nome:"Atendimento a demandas de órgãos centrais realizado", cat:"Atendimento Institucional", macro:"Gestão de Pessoas"},
};
const CPD_REGISTROS = [
  {id:1, servidor:"Servidor A", horas:34, texto:"Elaborei a minuta do edital de licitação para contratação do serviço de apoio administrativo e acompanhei a publicação no Compras.gov.",
    cands:[{cod:"00014210",conf:.98},{cod:"00014255",conf:.71},{cod:"00041002",conf:.42}]},
  {id:2, servidor:"Servidor B", horas:41, texto:"Análise e instrução de processos de aposentadoria no SEI, com verificação de tempo de contribuição e emissão de despacho.",
    cands:[{cod:"00021033",conf:.97},{cod:"00021047",conf:.80},{cod:"00041002",conf:.38}]},
  {id:3, servidor:"Servidor C", horas:28, texto:"Desenvolvimento do painel de indicadores de pessoal em Power BI e publicação para as áreas.",
    cands:[{cod:"00033801",conf:.96},{cod:"00033812",conf:.78},{cod:"00018904",conf:.31}]},
  {id:4, servidor:"Servidor D", horas:22, texto:"Conduzi a fiscalização mensal do contrato de limpeza, com ateste das notas e registro de ocorrências.",
    cands:[{cod:"00014255",conf:.95},{cod:"00014210",conf:.64},{cod:"00041002",conf:.29}]},
  {id:5, servidor:"Servidor E", horas:18, texto:"Preparação e condução de oficina de capacitação sobre o novo fluxo de concessão de benefícios.",
    cands:[{cod:"00027390",conf:.88},{cod:"00027365",conf:.74},{cod:"00021047",conf:.55}]},
  {id:6, servidor:"Servidor F", horas:25, texto:"Elaboração de nota técnica sobre a revisão da política de teletrabalho da unidade.",
    cands:[{cod:"00018904",conf:.84},{cod:"00018922",conf:.79},{cod:"00027365",conf:.40}]},
  {id:7, servidor:"Servidor G", horas:15, texto:"Estruturei a base de dados de lotação e disponibilizei o extrato mensal para a coordenação.",
    cands:[{cod:"00033812",conf:.82},{cod:"00033801",conf:.72},{cod:"00041002",conf:.36}]},
  {id:8, servidor:"Servidor H", horas:12, texto:"Respondi demandas do órgão central sobre o quadro de pessoal e prazos de recadastramento.",
    cands:[{cod:"00041002",conf:.77},{cod:"00021033",conf:.44},{cod:"00018904",conf:.33}]},
  {id:9, servidor:"Servidor I", horas:16, texto:"Reuniões de alinhamento semanal da equipe e organização da agenda de pautas do mês.",
    cands:[{cod:"00041002",conf:.52},{cod:"00027365",conf:.41},{cod:"00018904",conf:.24}]},
  {id:10, servidor:"Servidor J", horas:20, texto:"Mapeamento das trilhas de desenvolvimento da equipe e proposta de plano individual por servidor.",
    cands:[{cod:"00027365",conf:.86},{cod:"00027390",conf:.75},{cod:"00018904",conf:.47}]},
];
const CPD_ENTREGAS_ATUAIS = [
  {cod:"00021033", horas:120},{cod:"00014255", horas:80},{cod:"00027390", horas:45},{cod:"00018922", horas:30},
];
const CPD_AUTO=0.95, CPD_REV=0.70;

const cpdStatusPorConf = c => c>=CPD_AUTO ? "auto" : (c>=CPD_REV ? "revisar" : "sem");
const cpdConf   = r => r.cands[r.escolhido].conf;
const cpdCat    = r => { const c=r.cands[r.escolhido]; return {...CPD_CATALOGO[c.cod], cod:c.cod, conf:c.conf}; };
const cpdClasse = c => c>=CPD_AUTO?"g":(c>=CPD_REV?"a":"c");
const cpdCor    = c => c>=CPD_AUTO?"var(--cpd-green)":(c>=CPD_REV?"var(--cpd-amber)":"var(--cpd-coral)");
const cpdRotulo = c => c>=CPD_AUTO?"match automático":(c>=CPD_REV?"revisão humana":"sem correspondência");
const cpdFmt    = n => n==null?"—":Number(n).toLocaleString("pt-BR");

/* destaque dos termos em comum — simula a explicação do match */
const CPD_STOP=new Set(["sobre","para","com","mensal","novo","nova"]);
const cpdPrefixos = t => new Set((norm(t).match(/[a-z]{5,}/g)||[]).filter(w=>!CPD_STOP.has(w)).map(w=>w.slice(0,5)));
function cpdMarcar(a,b){
  const A=cpdPrefixos(a), comum=new Set([...cpdPrefixos(b)].filter(p=>A.has(p)));
  const wrap=t=>t.replace(/[A-Za-zÀ-ú]{5,}/g,w=>comum.has(norm(w).slice(0,5))?`<mark>${w}</mark>`:w);
  return [wrap(a),wrap(b)];
}
const CpdHTML = ({t,cls}) => <div className={cls} dangerouslySetInnerHTML={{__html:t}}/>;

function ImportarPGD({onConversor}){
  const inicial = () => CPD_REGISTROS.map(r=>({...r, escolhido:0, status:null, decidido:false}));
  const [regs,setRegs]     = useState(inicial);
  const [processado,setProcessado] = useState(false);
  const [revelados,setRevelados]   = useState(0);
  const [incluidas,setIncluidas]   = useState(false);
  const [tab,setTab]       = useState(0);
  const topoRef = useRef(null);

  /* fontes do protótipo (Archivo + Source Serif 4) — injetadas uma única vez */
  useEffect(()=>{
    if(typeof document==="undefined" || document.getElementById("cpd-fonts")) return;
    const l=document.createElement("link"); l.id="cpd-fonts"; l.rel="stylesheet";
    l.href="https://fonts.googleapis.com/css2?family=Archivo:wght@500;600;700;800&family=Source+Serif+4:opsz,wght@8..60,600;8..60,700&display=swap";
    document.head.appendChild(l);
  },[]);

  /* revelação progressiva da esteira, como no protótipo (170ms por linha) */
  useEffect(()=>{
    if(!processado || revelados>=regs.length) return;
    const t=setTimeout(()=>setRevelados(n=>n+1), revelados===0?120:170);
    return ()=>clearTimeout(t);
  },[processado,revelados,regs.length]);

  const aceitas    = regs.filter(r=>r.status==="auto"||(r.decidido&&r.status==="aceito"));
  const pendentes  = regs.filter(r=>r.status==="revisar"&&!r.decidido);
  const patch=(id,p)=>setRegs(l=>l.map(r=>r.id===id?{...r,...p}:r));

  function processar(){
    setRegs(l=>l.map(r=>({...r, status:cpdStatusPorConf(cpdConf(r))})));
    setRevelados(0); setProcessado(true);
  }
  function aceitar(id){ patch(id,{status:"aceito",decidido:true}); }
  function proporNova(id){ patch(id,{status:"novo",decidido:true}); }
  function trocar(id){
    setRegs(l=>l.map(r=>{
      if(r.id!==id) return r;
      const nr={...r, escolhido:(r.escolhido+1)%r.cands.length};
      const st=cpdStatusPorConf(cpdConf(nr));
      if(r.status!=="aceito"&&r.status!=="novo") nr.status = st==="auto" ? "revisar" : st;
      return nr;
    }));
  }
  function resetar(){ setRegs(inicial()); setProcessado(false); setRevelados(0); setIncluidas(false); setTab(0); }
  const irPara=i=>{ setTab(i); topoRef.current?.scrollIntoView({block:"start"}); };

  /* ---------------- hero ---------------- */
  const totalRegs=regs.length, horasCiclo=regs.reduce((s,r)=>s+r.horas,0);
  const nAuto=regs.filter(r=>r.status==="auto").length;
  const nRev=pendentes.length;
  const nSem=regs.filter(r=>r.status==="sem"&&!r.decidido).length;
  const stats = !processado
    ? [["",totalRegs,"Registros PGD"],["",cpdFmt(horasCiclo),"Horas no ciclo"],["faint","—","Match automático"],["faint","—","Para revisar"]]
    : [["",totalRegs,"Registros PGD"],["g",aceitas.length,"Aceitas"],["a",nRev,"Para revisar"],["c",nSem,"Sem match"]];
  const chipStatus = !processado ? "aguardando processamento"
    : (nRev>0 ? `${nRev} registros aguardando triagem`
              : (incluidas ? "entregas pré-incluídas na unidade" : "triagem concluída — pronto para pré-incluir"));

  /* ---------------- 01 · esteira ---------------- */
  const esteira = (<>
    <div className="cpd-kick"><span className="no">01</span><span className="t">Esteira de conversão</span></div>
    {!processado ? <>
      <div className="cpd-cta">
        <div className="cpd-exh-label">Ciclo fechado no PGD</div>
        <div className="cpd-exh-title" style={{margin:"6px 0 10px"}}>{totalRegs} registros de trabalho prontos para conversão</div>
        <p>Ao processar, a engine semântica compara cada descrição livre do PGD com o catálogo oficial de entregas
        e classifica por confiança: <b>≥95%</b> aceita automaticamente, <b>70–94%</b> vai para triagem humana,
        abaixo disso é candidata a nova entrega.</p>
        <button className="cpd-btn primary" onClick={processar}>Processar ciclo de junho</button>
      </div>
      {regs.map(r=>(
        <div className="cpd-row" key={r.id}>
          <div className="cpd-pgd">
            <div className="who"><span>PGD · {r.servidor}</span><span className="h">{r.horas}h</span></div>
            <div className="txt">{r.texto}</div>
          </div>
          <div className="cpd-conf"><div className="aguarda">aguardando</div></div>
          <div className="cpd-cat vazio">—</div>
        </div>
      ))}
    </> : <>
      <div className="cpd-exh">
        <div className="cpd-exh-head">
          <div className="cpd-exh-label">Exhibit P1</div>
          <div className="cpd-exh-title">Cada registro livre do PGD vira uma entrega do catálogo — com a confiança à mostra</div>
          <div className="cpd-exh-sub">termos em destaque = vocabulário em comum que sustentou o match · limiares: ≥95% automático · 70–94% triagem · &lt;70% sem correspondência</div>
        </div>
        {regs.map((r,i)=>{
          const c=cpdConf(r), ent=cpdCat(r), st=r.status;
          const semMatch=(st==="sem"&&!r.decidido);
          const [txtM,nomeM]=semMatch?[r.texto,""]:cpdMarcar(r.texto, ent.nome+" "+ent.cat);
          const nomeCurto = nomeM.split(" ").slice(0, ent.nome.split(" ").length).join(" ");
          return (
          <div className={`cpd-row ${i>=revelados?"oculta":""}`} key={r.id}>
            <div className="cpd-pgd">
              <div className="who"><span>PGD · {r.servidor}</span><span className="h">{r.horas}h</span></div>
              <CpdHTML cls="txt" t={txtM}/>
            </div>
            <div className="cpd-conf">
              <div className={`pc ${cpdClasse(c)}`}>{Math.round(c*100)}%</div>
              <div className="meter"><i style={{width:Math.round(c*100)+"%",background:cpdCor(c)}}/></div>
              <div className="lbl">{cpdRotulo(c)}</div>
            </div>
            {semMatch
              ? <div className="cpd-cat vazio">nenhuma entrega compatível no catálogo</div>
              : <div className="cpd-cat">
                  <div className="cod"><span>{ent.cod} · Banco de Entregas</span>
                    {st==="auto"   ? <span className="cpd-tag g">aceita automaticamente</span> :
                     st==="aceito" ? <span className="cpd-tag g">aceita na triagem</span> :
                     st==="novo"   ? <span className="cpd-tag p">nova entrega proposta</span> :
                                     <span className="cpd-tag a">aguardando triagem</span>}
                  </div>
                  <CpdHTML cls="nm" t={nomeCurto}/>
                  <div className="hier">{ent.cat} · {ent.macro}</div>
                </div>}
          </div>);
        })}
        <div className="cpd-src">Fonte (conceitual): registros do PGD do ciclo × catálogo Banco de Entregas · engine semântica UFRJ/LABORe</div>
      </div>
    </>}
  </>);

  /* ---------------- 02 · triagem ---------------- */
  const kcard=(r,acoes)=>{
    const c=cpdConf(r), ent=cpdCat(r);
    const cls=r.status==="aceito"?"resolvida":(r.status==="novo"?"rascunho":"");
    return (
      <div className={`cpd-kcard ${cls}`} key={r.id}>
        <div className="txt">{r.texto}</div>
        {r.status==="novo"
          ? <div className="match"><span className="cpd-tag p">rascunho de nova entrega enviado ao catálogo</span></div>
          : <div className="match"><div className="nm">{ent.nome}</div><div className="hier">{ent.cod} · {ent.cat}</div></div>}
        <div className="foot">
          <span className="conf-mini" style={{color:cpdCor(c)}}>
            <span className="m"><i style={{width:Math.round(c*100)+"%",background:cpdCor(c)}}/></span>{Math.round(c*100)}%
          </span>
          <span className="kacts">{acoes}</span>
        </div>
      </div>);
  };
  const colAuto=regs.filter(r=>r.status==="auto");
  const colRev =regs.filter(r=>r.status==="revisar"||r.status==="aceito");
  const colSem =regs.filter(r=>r.status==="sem"||r.status==="novo");
  const triagem = (<>
    <div className="cpd-kick"><span className="no">02</span><span className="t">Triagem humana</span></div>
    {!processado
      ? <div className="cpd-empty"><p><b>O ciclo ainda não foi processado.</b><br/>Volte à Esteira e clique em “Processar ciclo de junho” — a triagem recebe os matches de confiança intermediária.</p></div>
      : <div className="cpd-exh">
          <div className="cpd-exh-head">
            <div className="cpd-exh-label">Exhibit P2</div>
            <div className="cpd-exh-title">A máquina resolve o óbvio; a pessoa decide o resto</div>
            <div className="cpd-exh-sub">a coluna do meio é onde a triagem acontece — aceitar, trocar pela 2ª/3ª candidata ou propor entrega nova ao catálogo</div>
          </div>
          <div className="cpd-kanban">
            <div>
              <div className="cpd-kcol-h"><span className="t">Match automático · ≥95%</span><span className="n">{colAuto.length}</span></div>
              {colAuto.length?colAuto.map(r=>kcard(r)):<div className="cpd-kempty">nenhum match ≥95% neste ciclo</div>}
            </div>
            <div>
              <div className="cpd-kcol-h"><span className="t">Revisar · 70–94%</span><span className="n">{colRev.filter(r=>r.status==="revisar").length}</span></div>
              {colRev.length?colRev.map(r=>r.status==="aceito"?kcard(r):kcard(r,<>
                  <button className="ok" onClick={()=>aceitar(r.id)}>Aceitar</button>
                  <button onClick={()=>trocar(r.id)}>Trocar match</button>
                  <button className="nv" onClick={()=>proporNova(r.id)}>Propor nova</button>
                </>)):<div className="cpd-kempty">nada aguardando revisão</div>}
            </div>
            <div>
              <div className="cpd-kcol-h"><span className="t">Sem correspondência · &lt;70%</span><span className="n">{colSem.filter(r=>r.status==="sem").length}</span></div>
              {colSem.length?colSem.map(r=>r.status==="novo"?kcard(r):kcard(r,<>
                  <button onClick={()=>trocar(r.id)}>Ver candidatas</button>
                  <button className="nv" onClick={()=>proporNova(r.id)}>Propor nova entrega</button>
                </>)):<div className="cpd-kempty">todos os registros têm correspondência</div>}
            </div>
          </div>
          <div className="cpd-src">Decisões da triagem alimentam o re-treino da engine (ciclo virtuoso) e o fluxo de novas entregas do Catálogo — integração da fase seguinte.</div>
        </div>}
  </>);

  /* ---------------- 03 · resultado na unidade ---------------- */
  const porEntrega={};
  aceitas.forEach(r=>{ const e=cpdCat(r);
    porEntrega[e.cod]=porEntrega[e.cod]||{cod:e.cod,horas:0,regs:0};
    porEntrega[e.cod].horas+=r.horas; porEntrega[e.cod].regs++; });
  const jaTem=new Set(CPD_ENTREGAS_ATUAIS.map(e=>e.cod));
  const novas   = Object.values(porEntrega).filter(e=>!jaTem.has(e.cod));
  const reforco = Object.values(porEntrega).filter(e=> jaTem.has(e.cod));
  const horasMapeadas=aceitas.reduce((s,r)=>s+r.horas,0);
  const cob=Math.round(100*horasMapeadas/horasCiclo)||0;

  const resultado = (<>
    <div className="cpd-kick"><span className="no">03</span><span className="t">Resultado na unidade</span></div>
    {!processado
      ? <div className="cpd-empty"><p><b>O ciclo ainda não foi processado.</b><br/>Processe a esteira para ver o efeito na lista de entregas selecionadas da unidade.</p></div>
      : <div className="cpd-exh">
          <div className="cpd-exh-head">
            <div className="cpd-exh-label">Exhibit P3</div>
            <div className="cpd-exh-title">{cob}% do esforço do ciclo já chega <em>mapeado</em> à lista de entregas da unidade</div>
            <div className="cpd-exh-sub">o PGD deixa de ser registro paralelo e passa a alimentar o DFT — entregas novas chegam pré-incluídas, existentes ganham horas</div>
          </div>
          {pendentes.length>0 && <div className="cpd-exh-sub" style={{color:"var(--cpd-amber)",fontWeight:700,marginBottom:8}}>
            {pendentes.length} registros ainda aguardam triagem — as horas deles não estão contadas abaixo.</div>}
          <div className="cpd-cov">
            <div><div className="cov-num">{cob}%</div><div className="cov-lbl">cobertura do ciclo</div></div>
            <div className="cov-bar"><i style={{width:cob+"%"}}/></div>
            <div style={{fontSize:"10.5px",color:"var(--cpd-sub)",fontWeight:600}}>{cpdFmt(horasMapeadas)} de {cpdFmt(horasCiclo)} h mapeadas</div>
          </div>
          <table className="cpd-tab"><thead><tr><th>Entrega</th><th>Situação</th><th className="num">Horas/mês</th></tr></thead>
            <tbody>
              {novas.map(e=>{ const c=CPD_CATALOGO[e.cod]; return (
                <tr className="nova" key={e.cod}>
                  <td><b>{c.nome}</b><br/><span style={{fontSize:"10px",color:"var(--cpd-sub)"}}>{e.cod} · {c.cat}</span></td>
                  <td><span className="cpd-tag g">pré-incluída via PGD</span> <span className="cpd-tag k">{e.regs} {e.regs===1?"registro":"registros"}</span>
                    {incluidas && <> <span className="cpd-tag b">aguardando validação</span></>}</td>
                  <td className="num"><b style={{color:"var(--cpd-green)"}}>{cpdFmt(e.horas)}</b></td>
                </tr>); })}
              {CPD_ENTREGAS_ATUAIS.map(e=>{ const c=CPD_CATALOGO[e.cod], ref=reforco.find(x=>x.cod===e.cod); return (
                <tr key={e.cod}>
                  <td><b>{c.nome}</b><br/><span style={{fontSize:"10px",color:"var(--cpd-sub)"}}>{e.cod} · {c.cat}</span></td>
                  <td><span className="cpd-tag k">já selecionada</span>{ref && <> <span className="cpd-tag b">+{ref.regs} registros PGD</span></>}</td>
                  <td className="num">{cpdFmt(e.horas)}{ref && <> <b style={{color:"var(--cpd-green)"}}>+{cpdFmt(ref.horas)}</b></>}</td>
                </tr>); })}
            </tbody>
          </table>
          {incluidas
            ? <div className="cpd-confirm"><div className="ic">✓</div><p><b>{novas.length} entregas pré-incluídas na lista da COGEP</b> e {reforco.length} entregas existentes reforçadas com horas do PGD.
                O responsável pelo DFT da unidade valida (ou remove) as pré-inclusões no próximo ciclo de dimensionamento — nada entra no cálculo sem essa validação.</p></div>
            : <div style={{display:"flex",justifyContent:"flex-end",marginTop:16}}>
                <button className="cpd-btn primary" disabled={!novas.length} onClick={()=>setIncluidas(true)}>Pré-incluir {novas.length} entregas na lista da unidade</button>
              </div>}
          <div className="cpd-src">Fonte (conceitual): lista de entregas selecionadas da unidade no SISDIP + registros PGD aceitos na triagem deste ciclo.</div>
        </div>}
  </>);

  const VIEWS=[["01","Esteira"],["02","Triagem"],["03","Resultado na unidade"]];
  return (
    <div className="cpd" ref={topoRef}>
      <div className="cpd-govbar"/>
      <div className="cpd-hero"><div className="cpd-hero-in cpd-hero-grid">
        <div>
          <div className="cpd-kicker">Protótipo conceitual · Frente 2 × Frente 3</div>
          <div className="cpd-sig">PGD <span className="arrow">→</span> DFT</div>
          <div className="cpd-nome">Conversão dos registros mensais do Programa de Gestão em entregas do catálogo oficial — com pré-inclusão automática na lista de entregas selecionadas da unidade.</div>
          <div className="cpd-chips">
            <span className="cpd-chip">Unidade · <b>COGEP / SGP</b></span>
            <span className="cpd-chip">Ciclo · <b>junho 2026</b></span>
            <span className="cpd-chip live">{chipStatus}</span>
          </div>
        </div>
        <div>
          <div className="cpd-stats">
            {stats.map(([cls,v,l],i)=>(
              <div className="st" key={i}><div className={`v ${cls}`}>{v}</div><div className="l">{l}</div></div>
            ))}
          </div>
          <div className="cpd-engine">Engine de correspondência semântica · <b>UFRJ/LABORe (SBERT)</b> · limiar de auto-aceite <b>≥ 95%</b> · revisão humana entre 70–94%</div>
        </div>
      </div></div>

      <div className="cpd-subnav">
        <div className="cpd-tabs">
          {VIEWS.map(([k,t],i)=>(
            <button className={`cpd-tab ${i===tab?"on":""}`} key={k} onClick={()=>irPara(i)}>
              <span className="k">{k}</span>{t}
              {i===1&&processado&&pendentes.length>0 && <span className="cpd-badge">{pendentes.length}</span>}
            </button>))}
        </div>
        <div className="cpd-navr">
          {onConversor && <button className="cpd-btn" onClick={onConversor}>Ir ao Conversor</button>}
          <button className="cpd-btn" onClick={resetar}>Reiniciar demo</button>
        </div>
      </div>

      <div className="cpd-wrap">{[esteira,triagem,resultado][tab]}</div>

      <div className="cpd-foot">
        Protótipo conceitual — dados fictícios para demonstração do fluxo. A engine real de matching (UFRJ/LABORe) opera sobre o catálogo
        oficial do Banco de Entregas (~31 mil itens). A pré-inclusão alimenta a lista de entregas selecionadas da unidade no SISDIP/DFT,
        sujeita à validação do responsável pelo dimensionamento.
      </div>
    </div>
  );
}

/* ---------- MODO 4: explosão solar (sunburst zoomável animado) ---------- */
function trunc(s,n){ return s.length>n?s.slice(0,n-1)+"…":s; }
function buildSunFrom(list, capPerCat){
  const byNat=new Map(); const catCount=new Map();
  list.forEach(e=>{
    const ck=e.natureza+"|"+e.macro+"|"+e.categoria;
    const c=catCount.get(ck)||0; if(capPerCat && c>=capPerCat) return; catCount.set(ck,c+1);
    if(!byNat.has(e.natureza)) byNat.set(e.natureza,new Map());
    const macs=byNat.get(e.natureza);
    if(!macs.has(e.macro)) macs.set(e.macro,new Map());
    const cats=macs.get(e.macro);
    if(!cats.has(e.categoria)) cats.set(e.categoria,new Map());
    const servs=cats.get(e.categoria);
    if(!servs.has(e.servico)) servs.set(e.servico,[]);
    servs.get(e.servico).push(e);
  });
  // hierarquia até o SERVIÇO — as entregas ficam na folha (data.items), não como anéis
  return { name:"Catálogo", children:[...byNat].map(([nat,macs])=>({
    name:NAT[nat]?NAT[nat].rot:nat, natureza:nat,
    children:[...macs].map(([mac,cats])=>({ name:mac, natureza:nat,
      children:[...cats].map(([cat,servs])=>({ name:cat, natureza:nat,
        children:[...servs].map(([serv,items])=>({ name:serv||"(sem serviço)", natureza:nat, size:items.length, items }))
      }))
    }))
  }))};
}
function lighten(hex,t){ try{ return d3.interpolateRgb(hex,"#ffffff")(t); }catch(e){ return hex; } }
function textOn(hex){ const c=/^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex); if(!c) return "#fff";
  const r=parseInt(c[1],16),g=parseInt(c[2],16),b=parseInt(c[3],16); return (r*299+g*587+b*114)/1000>150?"#16283a":"#ffffff"; }
function Sunburst({res,add,rem,selSet,compare,toggleCompare,onFlag,justAdded}){
  const ref=useRef(null);
  const zoomRef=useRef(null);
  const [trail,setTrail]=useState([]);
  const [servSel,setServSel]=useState(null); // serviço escolhido → lista de entregas
  const [lexp,setLexp]=useState(null);       // detalhe expandido na lista
  const [zoom,setZoom]=useState(1);          // zoom in/out do gráfico
  const data=useMemo(()=>buildSunFrom(res, 0),[res]);
  useEffect(()=>{ setServSel(null); setLexp(null); setZoom(1); },[data]);

  // refs com os valores mais recentes (evita reconstruir o gráfico a cada seleção)
  const addRef=useRef(add); addRef.current=add;
  const remRef=useRef(rem); remRef.current=rem;
  const selRef=useRef(selSet); selRef.current=selSet;
  const pickRef=useRef(null); pickRef.current=(node)=>{ setServSel({
    name:node.data.name, natureza:node.data.natureza, items:node.data.items||[],
    trilha:node.ancestors().map(a=>a.data.name).reverse().slice(1) }); setLexp(null); };

  useEffect(()=>{
    const el=ref.current; if(!el) return;
    el.innerHTML="";
    if(!data.children || !data.children.length){ setTrail([]); return; }

    const width=820, radius=width/6;
    const hierarchy=d3.hierarchy(data).sum(d=>d.size||0).sort((a,b)=>(b.value||0)-(a.value||0));
    const root=d3.partition().size([2*Math.PI, hierarchy.height+1])(hierarchy);
    root.each(d=>d.current=d);

    const fillOf=d=>{ let a=d; while(a.depth>1 && a.parent) a=a.parent; const base=(a.data&&a.data.natureza&&NAT[a.data.natureza])?NAT[a.data.natureza].cor:"#9AA7B5";
      return lighten(base, Math.min(0.6,Math.max(0,(d.depth-1)*0.12))); };

    const arc=d3.arc().startAngle(d=>d.x0).endAngle(d=>d.x1)
      .padAngle(d=>Math.min((d.x1-d.x0)/2,0.004)).padRadius(radius*1.5)
      .innerRadius(d=>d.y0*radius).outerRadius(d=>Math.max(d.y0*radius,d.y1*radius-1));

    const svg=d3.select(el).append("svg")
      .attr("viewBox",[-width/2,-width/2,width,width].join(" "))
      .attr("width","100%").style("max-width",width+"px").style("height","auto")
      .style("font-family","'Raleway',system-ui,sans-serif");
    const g=svg.append("g");

    const arcVisible=d=>d.y1<=3 && d.y0>=1 && d.x1>d.x0;
    const labelVisible=d=>d.y1<=3 && d.y0>=1 && (d.y1-d.y0)*(d.x1-d.x0)>0.035;
    const labelTransform=d=>{ const x=(d.x0+d.x1)/2*180/Math.PI; const y=(d.y0+d.y1)/2*radius;
      return `rotate(${x-90}) translate(${y},0) rotate(${x<180?0:180})`; };

    const isServLeaf=d=> !!(d.data && d.data.items);
    const servHasSel=d=> isServLeaf(d) && selRef.current && d.data.items.some(it=>selRef.current.has(it.codigo));
    const path=g.append("g").selectAll("path").data(root.descendants().slice(1)).join("path")
      .attr("fill",d=>fillOf(d))
      .attr("fill-opacity",d=>arcVisible(d.current)?1:0)
      .attr("pointer-events",d=>arcVisible(d.current)?"auto":"none")
      .attr("stroke",d=>servHasSel(d)?C.primaryDark:"#fff").attr("stroke-width",d=>servHasSel(d)?2.4:1).attr("stroke-linejoin","round")
      .attr("d",d=>arc(d.current));
    path.style("cursor","pointer").on("click",clicked);
    path.append("title").text(d=>{
      const trilha=d.ancestors().map(a=>a.data.name).reverse().slice(1).join("  ›  ")||d.data.name;
      const serv=isServLeaf(d);
      return `${trilha}\n${d.value} entrega${d.value!==1?"s":""}${serv?"\nClique para ver as entregas deste serviço":""}`;
    });

    const label=g.append("g").attr("pointer-events","none").attr("text-anchor","middle").style("user-select","none")
      .selectAll("text").data(root.descendants().slice(1)).join("text")
      .attr("dy","0.34em")
      .attr("font-size",d=>d.depth===1?13.5:11.5).attr("font-weight",d=>d.depth===1?700:500)
      .attr("fill",d=>textOn(fillOf(d)))
      .attr("fill-opacity",d=>+labelVisible(d.current))
      .attr("transform",d=>labelTransform(d.current))
      .text(d=>trunc(d.data.name, d.depth===1?20:18));

    const parent=svg.append("circle").datum(root).attr("r",radius).attr("fill","#fff").attr("stroke",C.line)
      .attr("pointer-events","all").style("cursor","pointer").on("click",clicked);

    const cVal=svg.append("text").attr("text-anchor","middle").attr("dy","-0.15em")
      .attr("font-size",24).attr("font-weight",800).attr("fill",C.navy).text(root.value);
    const cLbl=svg.append("text").attr("text-anchor","middle").attr("dy","1.35em")
      .attr("font-size",11).attr("fill",C.sub).text("entregas");
    const cNm=svg.append("text").attr("text-anchor","middle").attr("dy","3.3em")
      .attr("font-size",10.5).attr("font-weight",600).attr("fill",C.faint).text("Catálogo");

    setTrail(root.ancestors().reverse());
    zoomRef.current=(node)=>clicked(null,node);

    function clicked(event,p){
      if(!p) return;
      // folha = serviço: abre a lista de entregas (estilo "Lista") em vez de dar zoom
      if(p.data && p.data.items){ pickRef.current(p); return; }
      setTrail(p.ancestors().reverse());
      parent.datum(p.parent||root);
      root.each(d=> d.target={
        x0:Math.max(0,Math.min(1,(d.x0-p.x0)/((p.x1-p.x0)||1)))*2*Math.PI,
        x1:Math.max(0,Math.min(1,(d.x1-p.x0)/((p.x1-p.x0)||1)))*2*Math.PI,
        y0:Math.max(0,d.y0-p.depth),
        y1:Math.max(0,d.y1-p.depth)
      });
      const t=g.transition().duration(740);
      path.transition(t).tween("data",d=>{ const i=d3.interpolate(d.current,d.target); return tt=>d.current=i(tt); })
        .filter(function(d){ return +this.getAttribute("fill-opacity") || arcVisible(d.target); })
        .attr("fill-opacity",d=>arcVisible(d.target)?1:0)
        .attr("pointer-events",d=>arcVisible(d.target)?"auto":"none")
        .attrTween("d",d=>()=>arc(d.current));
      label.filter(function(d){ return +this.getAttribute("fill-opacity") || labelVisible(d.target); })
        .transition(t)
        .attr("fill-opacity",d=>+labelVisible(d.target))
        .attrTween("transform",d=>()=>labelTransform(d.current));
      cVal.transition(t).textTween(()=>()=>String(p.value));
      cNm.transition(t).textTween(()=>()=>p.depth===0?"Catálogo":trunc(p.data.name,18));
    }

    return ()=>{ d3.select(el).selectAll("*").remove(); };
  },[data]);

  // aplica o zoom redimensionando o SVG (o container rola quando ampliado)
  useEffect(()=>{ const svg=ref.current&&ref.current.querySelector("svg"); if(!svg) return;
    if(zoom===1){ svg.style.width="100%"; svg.style.maxWidth="820px"; }
    else { svg.style.width=(820*zoom)+"px"; svg.style.maxWidth="none"; }
  },[zoom,data,servSel]);

  const empty=!data.children || !data.children.length;
  return (<div className="px-sun2">
    <div className="px-sun2-top">
      <div className="px-sun2-head">
        <h3>Explosão solar</h3>
        <p>Anéis por nível de agregação: natureza › macroprocesso › categoria › serviço. Clique para descer com animação; ao chegar em um <b>serviço</b>, veja as entregas dele em lista, logo abaixo. Clique no centro para voltar. A busca filtra o gráfico.</p>
      </div>
      <div className="px-sun2-leg">{Object.entries(NAT).map(([id,n])=><span key={id}><i style={{background:n.cor}}/>{n.rot}</span>)}</div>
    </div>
    {trail.length>0 && <div className="px-sun2-crumb">
      {trail.map((a,i)=>(<span key={i}><button onClick={()=>zoomRef.current&&zoomRef.current(a)}>{i===0?"Catálogo":trunc(a.data.name,30)}</button>{i<trail.length-1?<ChevronRight size={12}/>:null}</span>))}
    </div>}
    <div className="px-sun2-chartwrap">
      <div className="px-sun2-zoom">
        <button onClick={()=>setZoom(z=>Math.min(2.6,+(z+0.2).toFixed(2)))} disabled={zoom>=2.6} title="Aproximar"><ZoomIn size={15}/></button>
        <span className="px-sun2-zval">{Math.round(zoom*100)}%</span>
        <button onClick={()=>setZoom(z=>Math.max(0.6,+(z-0.2).toFixed(2)))} disabled={zoom<=0.6} title="Afastar"><ZoomOut size={15}/></button>
        <button onClick={()=>setZoom(1)} disabled={zoom===1} title="Redefinir zoom"><RotateCcw size={13}/></button>
      </div>
      <div className="px-sun2-chart" ref={ref}>{empty && <div className="px-sun-empty">Nenhuma entrega corresponde à busca.</div>}</div>
    </div>

    {servSel && <div className="px-sun2-list">
      <div className="px-sun2-list-h">
        <button className="px-sun2-list-back" onClick={()=>{setServSel(null);setLexp(null);}}><ChevronLeft size={14}/> Voltar ao gráfico</button>
        <div className="px-sun2-list-tt">
          <b>{servSel.name}</b>
          <span>{servSel.trilha.slice(0,-1).join("  ›  ")} · {servSel.items.length.toLocaleString("pt-BR")} entrega{servSel.items.length!==1?"s":""}</span>
        </div>
      </div>
      <div className="px-erows">
        {servSel.items.map(e=>{ const ok=selSet.has(e.codigo); const op=lexp===e.codigo; const sit=SIT[e.sit]; const inCmp=(compare||[]).some(x=>x.codigo===e.codigo); return (
          <div className={`px-erow ${ok?"sel":""}`} key={e.codigo}>
            <div className="px-erow-main">
              <button className={`px-cbox ${ok?"on":""}`} onClick={()=>ok?rem(e.codigo):add(e)} title={ok?"Remover":"Adicionar à descrição"}>{ok&&<Check size={12}/>}</button>
              <code className="px-ecode">{e.codigo}</code>
              {sit && <span className="px-sit sm" style={{background:sit.soft,color:sit.cor}}>{sit.rot}</span>}
              <span className="px-ename">{e.entrega}</span>
              <span className="px-eserv">{e.servico}</span>
              {justAdded===e.codigo && <span className="px-added"><Check size={11}/> adicionada</span>}
              <div className="px-erow-acts">
                {toggleCompare && <button className={`px-cmpbtn ${inCmp?"on":""}`} onClick={()=>toggleCompare(e)} title={inCmp?"Na comparação":"Comparar"}><Columns size={13}/></button>}
                <button className="px-edetbtn" onClick={()=>setLexp(op?null:e.codigo)} title="Informações gerais"><Info size={14}/></button>
              </div>
            </div>
            {op && <EntregaDetalhe e={e} inCmp={inCmp} toggleCompare={toggleCompare} onFlag={onFlag}/>}
          </div>); })}
      </div>
    </div>}

    <div className="px-sun2-note">Exibindo as {res.length.toLocaleString("pt-BR")} entregas do recorte atual, agregadas por serviço.</div>
  </div>);
}

/* ---------- Enquadrar Entregas (wizard: planilha/PGD/PDF/texto → DFT) ---------- */
const PGD_DEMO_LINHAS=[
  "Acompanhamento da política de minerais estratégicos",
  "Elaboração de estudos sobre economia circular na mineração",
  "Processamento mensal da folha de pagamento",
  "Concessão de aposentadorias a servidores",
  "Gestão de contratos de tecnologia da informação",
];
// itens de exemplo já "enquadrados" (só para ilustrar a tela na abertura — códigos reais do banco)
/* ============================================================================
   CONVERSOR UNIFICADO — vive dentro do Início (mode==="marcos", porta "conversor").
   Modelo de esteira do painel PGD->DFT (todas as linhas de uma vez, confianca
   a mostra, triagem por linha) + poderes do conversor avulso (anexar planilha/
   PDF/Word, IA real via backend, incluir na descricao da area, exportar).
============================================================================ */
const CONF_META={alta:{pct:95,cor:C.green,rot:"match automático"},media:{pct:80,cor:"#B8860B",rot:"revisar"},baixa:{pct:55,cor:"#C2410C",rot:"confiança baixa"}};

function ConversorUnificado({add,selSet,sel,notes,orgao,unidade,flash,onAbrirAssistente}){
  const codMap=useMemo(()=>{const m=new Map();ENTREGAS.forEach(e=>m.set(e.codigo,e));return m;},[]);
  const [itens,setItens]=useState([]);
  const [text,setText]=useState("");
  const [anexo,setAnexo]=useState(null); const [anexoErro,setAnexoErro]=useState("");
  const [aiLoad,setAiLoad]=useState(false); const [aviso,setAviso]=useState("");
  const [fase,setFase]=useState("entrada");
  const fileRef=useRef(null);

  async function escolherArquivo(ev){ const f=ev.target.files?.[0]; if(!f) return; setAnexoErro("");
    try{ const a=await lerArquivo(f); setAnexo(a);
      if(a.texto) setText(t=>(t?t+"\n":"")+a.texto.slice(0,8000));
    }catch(err){ setAnexoErro(err.message||"Não consegui ler o arquivo."); }
    finally{ if(fileRef.current) fileRef.current.value=""; }
  }
  function linhasDoTexto(t){ return t.split(/\r?\n/).map(s=>s.replace(/^[•\-\*\d\.\)\s]+/,"").trim()).filter(s=>s.length>2).slice(0,40); }

  async function enquadrar(){ const linhas=linhasDoTexto(text); if((!linhas.length&&!anexo)||aiLoad) return;
    setAiLoad(true); setAviso("");
    try{
      const bloco=linhas.map((ln,idx)=>({idx,ln,cands:buscarCandidatas(ln,12)}));
      const catalogoPorLinha=bloco.map(b=>`LINHA ${b.idx}: "${b.ln}"\ncandidatas:\n${candidatasTxt(b.cands)}`).join("\n\n");
      const sys=`Você enquadra entregas de uma área no catálogo do SISDIP. Para CADA linha, escolha entre as candidatas fornecidas as 5 MAIS pertinentes, ranqueadas da melhor para a pior. Responda APENAS um array JSON válido (sem markdown, sem texto fora). Formato: [{"idx":0,"opcoes":[{"cod":"0000.0000","conf":"alta|media|baixa","motivo":"3-6 palavras"}]}]. Use somente códigos presentes nas candidatas daquela linha. Traga até 5 opções por linha. Se realmente nenhuma servir, devolva "opcoes":[].`;
      let userContent;
      if(anexo?.tipo==="pdf" && !linhas.length){
        userContent=[{type:"document",source:{type:"base64",media_type:"application/pdf",data:anexo.pdfB64}},{type:"text",text:"Extraia as entregas deste PDF (uma por linha, idx sequencial) e enquadre com até 5 opções cada. Responda só o array JSON."}];
      } else {
        userContent=`Enquadre cada linha usando apenas suas candidatas. Responda só o array JSON.\n\n${catalogoPorLinha}`;
      }
      const d=await chamarIA({model:"claude-haiku-4-5-20251001",max_tokens:3500,system:sys,messages:[{role:"user",content:userContent}]});
      const tx=(d.content||[]).map(b=>b.type==="text"?b.text:"").join("").trim();
      const arr=JSON.parse(tx.replace(/```json|```/g,"").trim());
      const novos=linhas.map((ln,idx)=>{
        const r=arr.find(a=>a.idx===idx)||{opcoes:[]};
        const opcoes=(r.opcoes||[]).filter(o=>o.cod&&codMap.has(o.cod)).slice(0,5)
          .map(o=>({cod:o.cod,conf:["alta","media","baixa"].includes(o.conf)?o.conf:"media",motivo:o.motivo||""}));
        const melhor=opcoes[0]||null;
        return {pgd:ln, opcoes, escolhido:melhor?.cod||null, novaProposta:opcoes.length===0,
                incluir: !!melhor && melhor.conf==="alta",
                adicionada:false};
      });
      setItens(novos); setFase("esteira");
      const auto=novos.filter(n=>n.incluir).length, sem=novos.filter(n=>n.novaProposta).length;
      setAviso(`${novos.length} entrega(s) analisada(s) · ${auto} com match automático · ${sem} sem correspondência.`);
    }catch(err){
      setAviso("⚠️ Não consegui falar com o serviço de IA agora. Se o portal está publicado, confira se a função de IA (backend) está configurada e com créditos.");
    } finally{ setAiLoad(false); }
  }

  const setItem=(i,patch)=>setItens(list=>list.map((it,j)=>j===i?{...it,...patch}:it));
  const confDe=(it)=> it.novaProposta ? null : (it.opcoes.find(o=>o.cod===it.escolhido)?.conf||"media");

  const marcados=itens.filter(it=>it.incluir&&!it.adicionada);
  function incluirMarcadas(){
    let n=0;
    setItens(list=>list.map((it,i)=>{
      if(!(it.incluir&&!it.adicionada)) return it;
      if(it.novaProposta){ add({codigo:"nova-"+Date.now().toString(36)+"-"+i, entrega:it.pgd, nova:true, natureza:null, servico:"", categoria:"", macro:"", atividade:"", metod:""}); n++; }
      else if(it.escolhido){ const e=codMap.get(it.escolhido); if(e&&!selSet.has(e.codigo)){ add(e); n++; } }
      return {...it,adicionada:true};
    }));
    flash(`${n} entrega(s) incluída(s) na sua descrição da área.`);
  }
  const baixarPlanilha=()=>{
    const escolhidas=itens.filter(it=>it.adicionada&&!it.novaProposta&&it.escolhido).map(it=>codMap.get(it.escolhido)).filter(Boolean);
    const novas=itens.filter(it=>it.adicionada&&it.novaProposta).map((it,k)=>({codigo:"nova-dl-"+k,entrega:it.pgd,nova:true,natureza:null}));
    const mapa=new Map(); [...sel,...escolhidas].forEach(e=>mapa.set(e.codigo,e));
    const lista=[...mapa.values(),...novas];
    if(!lista.length){ setAviso("Nada para baixar ainda — inclua ao menos uma entrega."); return; }
    exportarDescricaoXLSX(lista,notes||{},orgao,unidade);
  };

  if(fase==="entrada") return (
    <div className="px-cu">
      <div className="px-cu-hero">
        <div className="px-cu-hero-ic"><Wand2 size={22}/></div>
        <h2>Traga seu <em>PGD</em>, <em>regimento interno</em> ou <em>planejamento estratégico</em></h2>
        <p>Cole o texto ou anexe o arquivo — a IA enquadra cada item no catálogo oficial, com a confiança à mostra, e você inclui direto na descrição da sua área.</p>
      </div>
      <div className="px-cu-input">
        <div className="px-cu-drop" onClick={()=>fileRef.current?.click()}>
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv,.pdf,.docx" style={{display:"none"}} onChange={escolherArquivo}/>
          <Upload size={18}/>
          <div><b>Anexar planilha, PDF ou Word</b><span>.xlsx · .csv · .pdf · .docx</span></div>
        </div>
        {anexo && <div className="px-conv-anexo"><FileText size={12}/> {anexo.nome} <button onClick={()=>setAnexo(null)}><X size={11}/></button></div>}
        {anexoErro && <div className="px-conv-anexo erro"><AlertTriangle size={12}/> {anexoErro}</div>}
        <textarea value={text} onChange={e=>setText(e.target.value)} placeholder="…ou cole aqui — uma entrega/atividade por linha. Ex.: relatórios do PGD, competências do regimento, iniciativas do planejamento."/>
        <div className="px-cu-acts">
          <button className="px-cu-go" onClick={enquadrar} disabled={aiLoad||(!text.trim()&&!anexo)}><Wand2 size={15}/> {aiLoad?"Analisando…":"Enquadrar no catálogo"}</button>
          <button className="px-cu-chat" onClick={onAbrirAssistente}><Bot size={14}/> Prefere conversar? Abra o Assistente</button>
        </div>
        {aviso && <div className="px-conv-aiout">{aviso}</div>}
      </div>
    </div>
  );

  return (
    <div className="px-cu">
      <div className="px-cu-topo">
        <button className="px-cu-voltar" onClick={()=>setFase("entrada")}><ChevronLeft size={14}/> Novo lote</button>
        <div className="px-cu-resumo">{aviso}</div>
        <div className="px-cu-topo-acts">
          <button className="px-cu-chat mini" onClick={onAbrirAssistente}><Bot size={13}/> Assistente</button>
          <button className="px-cu-dl" onClick={baixarPlanilha}><Download size={13}/> Planilha</button>
          <button className="px-cu-go" onClick={incluirMarcadas} disabled={!marcados.length}><ClipboardList size={14}/> Incluir {marcados.length||""} na minha descrição</button>
        </div>
      </div>

      {itens.map((it,i)=>{
        const conf=confDe(it); const meta=conf?CONF_META[conf]:null;
        const ent=it.escolhido?codMap.get(it.escolhido):null;
        return (
        <div key={i} className={`px-cu-row ${it.adicionada?"ok":""}`}>
          <label className="px-cu-check" title={it.adicionada?"Já incluída":"Marcar para incluir na descrição"}>
            <input type="checkbox" checked={it.incluir||it.adicionada} disabled={it.adicionada} onChange={e=>setItem(i,{incluir:e.target.checked})}/>
          </label>
          <div className="px-cu-pgd"><span className="px-cu-lbl">Sua entrega</span><div>{it.pgd}</div></div>
          <div className="px-cu-conf">
            {meta ? <>
              <b style={{color:meta.cor}}>{meta.pct}%</b>
              <span className="px-cu-meter"><i style={{width:meta.pct+"%",background:meta.cor}}/></span>
              <em>{it.adicionada?"incluída":meta.rot}</em>
            </> : <><b style={{color:"#C2410C"}}>—</b><em>sem correspondência</em></>}
          </div>
          <div className="px-cu-cat">
            {it.novaProposta
              ? <div className="px-cu-nova"><FilePlus2 size={13}/> Proposta de <b>nova entrega</b> ao catálogo{it.opcoes.length>0&&<button onClick={()=>setItem(i,{novaProposta:false,escolhido:it.opcoes[0].cod})}>voltar às opções</button>}</div>
              : <>
                <span className="px-cu-lbl">{ent?.codigo} · Banco de Entregas</span>
                <div className="px-cu-nome">{ent?.entrega}</div>
                <div className="px-cu-hier">{ent?.macro} › {ent?.categoria}</div>
                {it.opcoes.length>1 && <div className="px-cu-alts">{it.opcoes.map(o=>(
                  <button key={o.cod} className={o.cod===it.escolhido?"on":""} title={`${codMap.get(o.cod)?.entrega} ${o.motivo?("· "+o.motivo):""}`}
                    onClick={()=>setItem(i,{escolhido:o.cod,incluir:o.conf==="alta"?true:it.incluir})}>{codMap.get(o.cod)?.entrega.slice(0,34)}{codMap.get(o.cod)?.entrega.length>34?"…":""}</button>))}
                  <button className="nv" onClick={()=>setItem(i,{novaProposta:true,escolhido:null})}>propor nova</button>
                </div>}
              </>}
          </div>
        </div>);
      })}
      <div className="px-cu-rodape">A triagem é sua: alta confiança já vem marcada; revise as demais, troque a opção ou proponha uma nova entrega. Nada entra na descrição sem o seu clique.</div>
    </div>
  );
}

const DEMO_ITENS=[
  {pgd:"Processamento mensal da folha de pagamento", opcoes:[
    {cod:"01070152",conf:"alta",motivo:"folha no SIAPE"},
    {cod:"01070538",conf:"media",motivo:"benefício em folha"},
  ], escolhido:"01070152", novaProposta:false, adicionada:false},
  {pgd:"Concessão de aposentadorias a servidores", opcoes:[
    {cod:"01020078",conf:"alta",motivo:"portaria de concessão"},
    {cod:"01020080",conf:"baixa",motivo:"revisão no SIAPE"},
  ], escolhido:"01020078", novaProposta:false, adicionada:false},
];
const CONF={alta:{rot:"Alta",cor:"#168821",soft:"#E3F2E5"},media:{rot:"Média",cor:"#B86E00",soft:"#FBEEDB"},baixa:{rot:"Baixa",cor:"#8C97A6",soft:"#EEF1F4"},nova:{rot:"Propor nova",cor:"#1351B4",soft:"#E8EEF9"}};

function ConversorPGD({onClose,add,selSet,sel,notes,orgao,unidade,flash,inline}){
  const codMap=useMemo(()=>{const m=new Map();ENTREGAS.forEach(e=>m.set(e.codigo,e));return m;},[]);
  // cada item: {pgd, opcoes:[{cod,conf,motivo}], escolhido, novaProposta:bool, adicionada}
  const [itens,setItens]=useState(DEMO_ITENS.filter(d=>d.opcoes.every(o=>codMap.has(o.cod))));
  const [ehExemplo,setEhExemplo]=useState(true);
  const [passo,setPasso]=useState(0);           // índice da entrega no wizard
  const [text,setText]=useState("");
  const [anexo,setAnexo]=useState(null); const [anexoErro,setAnexoErro]=useState("");
  const [aiLoad,setAiLoad]=useState(false); const [aviso,setAviso]=useState("");
  const fileRef=useRef(null);

  const limparExemplo=()=>{ if(ehExemplo){ setEhExemplo(false); setItens([]); setText(""); setPasso(0); } };

  async function escolherArquivo(ev){ const f=ev.target.files?.[0]; if(!f) return; setAnexoErro(""); limparExemplo();
    try{ const a=await lerArquivo(f); setAnexo(a);
      if(a.texto) setText(t=>(t?t+"\n":"")+a.texto.slice(0,8000));
    }catch(err){ setAnexoErro(err.message||"Não consegui ler o arquivo."); }
    finally{ if(fileRef.current) fileRef.current.value=""; }
  }

  function linhasDoTexto(t){ return t.split(/\r?\n/).map(s=>s.replace(/^[•\-\*\d\.\)\s]+/,"").trim()).filter(s=>s.length>2).slice(0,40); }

  async function enquadrar(){ const linhas=linhasDoTexto(text); if((!linhas.length&&!anexo)||aiLoad) return;
    setAiLoad(true); setAviso(""); limparExemplo();
    try{
      const bloco=linhas.map((ln,idx)=>({idx,ln,cands:buscarCandidatas(ln,12)}));
      const catalogoPorLinha=bloco.map(b=>`LINHA ${b.idx}: "${b.ln}"\ncandidatas:\n${candidatasTxt(b.cands)}`).join("\n\n");
      const sys=`Você enquadra entregas de uma área no catálogo do SISDIP. Para CADA linha, escolha entre as candidatas fornecidas as 5 MAIS pertinentes, ranqueadas da melhor para a pior. Responda APENAS um array JSON válido (sem markdown, sem texto fora). Formato: [{"idx":0,"opcoes":[{"cod":"0000.0000","conf":"alta|media|baixa","motivo":"3-6 palavras"}]}]. Use somente códigos presentes nas candidatas daquela linha. Traga até 5 opções por linha. Se realmente nenhuma servir, devolva "opcoes":[].`;
      let userContent;
      if(anexo?.tipo==="pdf" && !linhas.length){
        userContent=[{type:"document",source:{type:"base64",media_type:"application/pdf",data:anexo.pdfB64}},{type:"text",text:"Extraia as entregas deste PDF (uma por linha, idx sequencial) e enquadre com até 5 opções cada. Responda só o array JSON."}];
      } else {
        userContent=`Enquadre cada linha usando apenas suas candidatas. Responda só o array JSON.\n\n${catalogoPorLinha}`;
      }
      const d=await chamarIA({model:"claude-haiku-4-5-20251001",max_tokens:3500,system:sys,messages:[{role:"user",content:userContent}]});
      const tx=(d.content||[]).map(b=>b.type==="text"?b.text:"").join("").trim();
      const arr=JSON.parse(tx.replace(/```json|```/g,"").trim());
      const novos=linhas.map((ln,idx)=>{
        const r=arr.find(a=>a.idx===idx)||{opcoes:[]};
        const opcoes=(r.opcoes||[]).filter(o=>o.cod&&codMap.has(o.cod)).slice(0,5)
          .map(o=>({cod:o.cod,conf:["alta","media","baixa"].includes(o.conf)?o.conf:"media",motivo:o.motivo||""}));
        return {pgd:ln, opcoes, escolhido:opcoes[0]?.cod||null, novaProposta:opcoes.length===0, adicionada:false};
      });
      setItens(novos); setPasso(0);
      const semOp=novos.filter(n=>!n.opcoes.length).length;
      setAviso(`${novos.length} entrega(s) analisada(s).${semOp?` ${semOp} sem correspondência — sugeridas como nova.`:""}`);
    }catch(err){
      setAviso("⚠️ Não consegui falar com o serviço de IA agora. Se o portal está publicado, confira se a função de IA (backend) está configurada e com créditos.");
    } finally{ setAiLoad(false); }
  }

  const escolher=(cod)=>setItens(list=>list.map((it,j)=>j===passo?{...it,escolhido:cod,novaProposta:false}:it));
  const marcarNova=()=>setItens(list=>list.map((it,j)=>j===passo?{...it,novaProposta:true,escolhido:null}:it));

  // adiciona o item atual à descrição de área (entrega do banco OU proposta nova)
  const adicionarAtual=()=>{ const it=itens[passo]; if(!it) return;
    if(it.novaProposta){
      const codNova="nova-"+Date.now().toString(36)+"-"+passo;
      add({codigo:codNova, entrega:it.pgd, nova:true, natureza:null, servico:"", categoria:"", macro:"", atividade:"", metod:""});
    } else if(it.escolhido){ const e=codMap.get(it.escolhido); if(e&&!selSet.has(e.codigo)) add(e); }
    setItens(list=>list.map((x,j)=>j===passo?{...x,adicionada:true}:x));
    if(passo<itens.length-1) setPasso(passo+1);
  };

  const baixarPlanilha=()=>{
    const escolhidasAqui=itens.filter(it=>it.adicionada&&!it.novaProposta&&it.escolhido).map(it=>codMap.get(it.escolhido)).filter(Boolean);
    const novasAqui=itens.filter(it=>it.adicionada&&it.novaProposta).map((it,k)=>({codigo:"nova-dl-"+k,entrega:it.pgd,nova:true,natureza:null}));
    const mapa=new Map(); [...sel,...escolhidasAqui].forEach(e=>mapa.set(e.codigo,e));
    const lista=[...mapa.values(),...novasAqui];
    if(!lista.length){ setAviso("Nada para baixar ainda — adicione ao menos uma entrega."); return; }
    exportarDescricaoXLSX(lista,notes||{},orgao,unidade);
  };

  const total=itens.length;
  const it=itens[passo];
  const totalAdicionadas=itens.filter(x=>x.adicionada).length;
  const progresso=total?Math.round((totalAdicionadas/total)*100):0;

  const painel=(
    <div className={`px-conv ${inline?"inline":""}`} onClick={e=>e.stopPropagation()}>
      <div className="px-conv-h">
        <div className="px-conv-ht"><span className="px-conv-hic"><Wand2 size={17}/></span><div><b>Conversor</b><span>Traga de planilha, PGD, PDF ou texto · a IA sugere · você escolhe uma a uma</span></div></div>
        {!inline && <button onClick={onClose} title="Fechar"><X size={18}/></button>}
      </div>
      <div className="px-conv-body">
        <div className="px-conv-left">
          <div className="px-conv-step">1 · Informe as entregas da área</div>
          <div className="px-conv-drop" onClick={()=>fileRef.current?.click()}>
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv,.pdf,.docx" style={{display:"none"}} onChange={escolherArquivo}/>
            <Upload size={18}/>
            <div><b>Anexar planilha, PDF ou Word</b><span>.xlsx · .csv · .pdf · .docx — ou digite abaixo</span></div>
          </div>
          {anexo && <div className="px-conv-anexo"><FileText size={12}/> {anexo.nome} <button onClick={()=>setAnexo(null)}><X size={11}/></button></div>}
          {anexoErro && <div className="px-conv-anexo erro"><AlertTriangle size={12}/> {anexoErro}</div>}
          <textarea value={text} onChange={e=>{limparExemplo();setText(e.target.value);}} placeholder="Uma entrega/atividade por linha — descreva o que a área faz…"/>
          <div className="px-conv-lbtns">
            <button className="px-conv-ai" onClick={enquadrar} disabled={aiLoad||(!text.trim()&&!anexo)}><Wand2 size={13}/> {aiLoad?"Analisando…":"Enquadrar com IA"}</button>
          </div>
          {aviso && <div className="px-conv-aiout">{aviso}</div>}
        </div>

        <div className="px-conv-right">
          <div className="px-conv-step">2 · Escolha a correspondência de cada entrega</div>
          {ehExemplo && <div className="px-conv-exlabel"><Info size={11}/> Exemplo — comece a digitar ou anexe um arquivo para montar a sua descrição</div>}

          {total===0
            ? <div className="px-conv-vazio">Informe suas entregas e clique em <b>Enquadrar com IA</b> — cada uma virá com até 5 opções para você escolher, uma por vez.</div>
            : <div className="px-wiz">
                {/* barra de progresso */}
                <div className="px-wiz-prog">
                  <div className="px-wiz-prog-top">
                    <span>Entrega <b>{passo+1}</b> de {total}</span>
                    <span>{totalAdicionadas} de {total} na descrição</span>
                  </div>
                  <div className="px-wiz-bar"><i style={{width:progresso+"%"}}/></div>
                  <div className="px-wiz-dots">
                    {itens.map((x,k)=><button key={k} className={`px-wiz-dot ${k===passo?"cur":""} ${x.adicionada?"ok":""}`} onClick={()=>setPasso(k)} title={`Entrega ${k+1}`}/>)}
                  </div>
                </div>

                {/* card da entrega atual */}
                <div className={`px-wiz-card ${it.adicionada?"add":""}`}>
                  <div className="px-wiz-orig"><span className="px-imp-num">{passo+1}</span> {it.pgd}</div>

                  <div className="px-wiz-ops">
                    {it.opcoes.map(op=>{ const e=codMap.get(op.cod); const cf=CONF[op.conf]; const on=!it.novaProposta&&it.escolhido===op.cod; if(!e) return null; const nat=NAT[e.natureza]; const cor=nat?nat.cor:C.primary; const soft=nat?nat.soft:C.primarySoft; return (
                      <button key={op.cod} className={`px-imp-op ${on?"on":""}`} style={{"--opcor":cor,"--opsoft":soft}} onClick={()=>escolher(op.cod)}>
                        <span className={`px-imp-radio ${on?"on":""}`}>{on&&<Check size={12}/>}</span>
                        <span className="px-imp-op-body">
                          <span className="px-imp-op-top"><code>{op.cod}</code><em style={{background:cf.soft,color:cf.cor}}>{cf.rot}</em></span>
                          <span className="px-imp-op-name">{e.entrega}</span>
                        </span>
                      </button>); })}

                    {/* opção sempre disponível: propor como nova */}
                    <button className={`px-imp-op nova ${it.novaProposta?"on":""}`} onClick={marcarNova}>
                      <span className={`px-imp-radio ${it.novaProposta?"on":""}`}>{it.novaProposta&&<Check size={11}/>}</span>
                      <span className="px-imp-op-body">
                        <span className="px-imp-op-top"><Sparkles size={12} color={C.primary}/> <b style={{color:C.primaryDark,fontSize:12}}>Propor como nova entrega</b></span>
                        <span className="px-imp-op-name">Nenhuma acima serve — enviar “{trunc2(it.pgd,44)}” para a curadoria avaliar.</span>
                      </span>
                    </button>
                  </div>
                </div>

                {/* navegação do wizard */}
                <div className="px-wiz-nav">
                  <button className="px-wiz-prev" disabled={passo===0} onClick={()=>setPasso(passo-1)}><ChevronLeft size={15}/> Anterior</button>
                  {it.adicionada
                    ? <span className="px-imp-added"><Check size={13}/> adicionada</span>
                    : <button className="px-wiz-add" disabled={!it.escolhido&&!it.novaProposta} onClick={adicionarAtual}>
                        <Plus size={14}/> {it.novaProposta?"Adicionar como nova":"Adicionar à descrição"}{passo<total-1?" e avançar":""}
                      </button>}
                  <button className="px-wiz-next" disabled={passo>=total-1} onClick={()=>setPasso(passo+1)}>Próxima <ChevronRight size={15}/></button>
                </div>
              </div>}
        </div>
      </div>
      <div className="px-conv-foot">
        <span className="px-conv-stat">{total===0?<em style={{color:C.faint}}>Aguardando suas entregas</em>:<><b>{totalAdicionadas}</b> de {total} adicionadas à descrição</>}</span>
        <div className="px-conv-footbtns">
          <button className="px-conv-dl" disabled={!totalAdicionadas&&!(sel&&sel.length)} onClick={baixarPlanilha}><Download size={14}/> Baixar planilha (.xlsx)</button>
          <button className="px-conv-inject" disabled={!totalAdicionadas} onClick={onClose}><Check size={15}/> Concluir</button>
        </div>
      </div>
    </div>
  );
  return inline ? painel : <div className="px-modal-bg" onClick={onClose}>{painel}</div>;
}

/* ---------- Sinalizar problema (loop de qualidade com a curadoria do banco) ---------- */
const FLAG_TIPOS=["Entrega duplicada","Descrição ambígua","Entrega desatualizada","Classificação incorreta","Outro"];
function FlagModal({entrega,nova,proposta,onClose,onSubmit}){
  const [tipo,setTipo]=useState(FLAG_TIPOS[0]);
  const [nome,setNome]=useState(proposta||"");
  const [obs,setObs]=useState("");
  if(nova){
    return (<div className="px-modal-bg" onClick={onClose}>
      <div className="px-flagm" onClick={e=>e.stopPropagation()}>
        <div className="px-flagm-h"><div><Plus size={17} color={C.primary}/> <b>Propor nova entrega ao banco</b></div><button onClick={onClose}><X size={17}/></button></div>
        <p className="px-flagm-sub">Não achou no catálogo? Descreva a entrega que falta. A proposta vai para a <b>curadoria do banco</b> avaliar.</p>
        <div className="px-flagm-l">Nome sugerido da entrega</div>
        <input className="px-flagm-in" value={nome} onChange={e=>setNome(e.target.value)} placeholder="Ex.: Relatório de gestão consolidado"/>
        <div className="px-flagm-l">Por que ela é necessária (opcional)</div>
        <textarea value={obs} onChange={e=>setObs(e.target.value)} placeholder="Contexto, atividade relacionada, macroprocesso sugerido…"/>
        <div className="px-flagm-acts"><button className="px-btn-ghost" onClick={onClose}>Cancelar</button><button className="px-btn-primary" disabled={!nome.trim()} onClick={()=>onSubmit({codigo:"nova",entrega:nome.trim(),tipo:"Propor nova entrega",obs})}><Plus size={14}/> Enviar proposta</button></div>
      </div>
    </div>);
  }
  return (<div className="px-modal-bg" onClick={onClose}>
    <div className="px-flagm" onClick={e=>e.stopPropagation()}>
      <div className="px-flagm-h"><div><AlertTriangle size={17} color="#B86E00"/> <b>Sinalizar problema ao banco</b></div><button onClick={onClose}><X size={17}/></button></div>
      <p className="px-flagm-sub">Sua sinalização vai para a <b>curadoria do banco</b>. Entrega <code>{entrega.codigo}</code> — {entrega.entrega}.</p>
      <div className="px-flagm-l">Tipo de problema</div>
      <div className="px-flagm-tipos">{FLAG_TIPOS.map(t=><button key={t} className={`px-flagm-tipo ${tipo===t?"on":""}`} onClick={()=>setTipo(t)}>{t}</button>)}</div>
      <div className="px-flagm-l">Observação (opcional)</div>
      <textarea value={obs} onChange={e=>setObs(e.target.value)} placeholder="Descreva o problema encontrado…"/>
      <div className="px-flagm-acts"><button className="px-btn-ghost" onClick={onClose}>Cancelar</button><button className="px-btn-primary" onClick={()=>onSubmit({codigo:entrega.codigo,entrega:entrega.entrega,tipo,obs})}><Flag size={14}/> Enviar à curadoria</button></div>
    </div>
  </div>);
}
function QualidadePanel({flags,onClose,embutido}){
  const corpo=(
    <div className={embutido?"px-qual px-qual-embutido":"px-qual"} onClick={e=>e.stopPropagation()}>
      <div className="px-flagm-h"><div><Flag size={17} color="#B86E00"/> <b>Qualidade do banco · sinalizações</b></div>{!embutido && <button onClick={onClose}><X size={17}/></button>}</div>
      <p className="px-flagm-sub">As entregas sinalizadas passam por uma curadoria do banco, que avalia e trata cada apontamento levantado pelo uso do portal.</p>
      {flags.length===0
        ? <div className="px-qual-empty">Nenhuma sinalização ainda. Abra os detalhes de uma entrega e use “Sinalizar problema ao banco”.</div>
        : <div className="px-qual-list">{flags.map((f,i)=>(<div className="px-qual-item" key={i}><span className="px-conv-conf" style={{background:"#FBEEDB",color:"#B86E00"}}>{f.tipo}</span><div className="px-qual-body"><div className="px-qual-row"><code>{f.codigo}</code> <b>{f.entrega}</b></div>{f.obs&&<p>{f.obs}</p>}</div></div>))}</div>}
    </div>
  );
  return embutido ? corpo : <div className="px-modal-bg" onClick={onClose}>{corpo}</div>;
}

/* ---------- css ---------- */
const css=`
*{box-sizing:border-box;}
/* escala tipográfica: 400 corpo · 600 label · 700 nome/ação · 800 título de seção */
.px-app,.px-app input,.px-app textarea,.px-app select,.px-app button{font-weight:400;}
.px-stripe{display:flex;height:4px;} .px-stripe span{flex:1;}
.px-head{background:#fff;padding:0 24px;height:57px;box-sizing:border-box;display:flex;justify-content:space-between;align-items:center;gap:18px;border-bottom:1px solid ${C.line};position:sticky;top:0;z-index:95;}
.px-nav{display:flex;align-items:center;gap:4px;flex:1;justify-content:center;min-width:0;}
.px-nav-item{display:flex;align-items:center;gap:7px;border:none;background:none;font-family:inherit;font-size:13px;font-weight:700;color:${C.faint};padding:8px 14px;border-radius:9px;cursor:pointer;transition:all .14s;white-space:nowrap;}
.px-nav-item:hover{color:${C.navy};background:${C.bg};}
.px-nav-item.on{color:#fff;background:${C.navy};}
@media(max-width:1100px){.px-nav-item span{display:none;}.px-nav-item{padding:9px 11px;}}
.px-wrap-secao{padding-top:18px;}
/* ---- Início ---- */
.px-home{max-width:1060px;margin:0 auto;padding:44px 28px 60px;animation:fadeUp .35s ease;}
@keyframes fadeUp{from{opacity:0;transform:translateY(10px);}to{opacity:1;transform:none;}}
.px-home-hero{text-align:center;max-width:720px;margin:0 auto 36px;}
.px-home-hero h1{font-size:31px;line-height:1.25;font-weight:800;color:${C.navy};letter-spacing:-.4px;margin:0 0 14px;}
.px-home-hero p{font-size:14.5px;line-height:1.65;color:${C.sub};margin:0 auto 20px;max-width:600px;}
.px-home-nums{display:flex;justify-content:center;align-items:center;gap:16px;font-size:13px;color:${C.sub};}
.px-home-nums b{color:${C.navy};font-size:16px;}
.px-home-nums i{width:1px;height:16px;background:${C.line};}
.px-home-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;}
@media(max-width:900px){.px-home-grid{grid-template-columns:repeat(2,1fr);}}
@media(max-width:600px){.px-home-grid{grid-template-columns:1fr;}}
.px-home-card{display:flex;flex-direction:column;align-items:flex-start;gap:9px;text-align:left;background:#fff;border:1px solid ${C.line};border-radius:15px;padding:20px 20px 17px;cursor:pointer;font-family:inherit;transition:all .16s;position:relative;}
.px-home-card:hover{border-color:${C.primary};box-shadow:0 8px 26px rgba(19,49,92,.1);transform:translateY(-2px);}
.px-home-card.grande{grid-column:span 1;background:linear-gradient(180deg,#fff, #FAFBFD);}
.px-home-card.grande b{font-size:16px;}
.px-home-card-ic{width:42px;height:42px;border-radius:12px;display:grid;place-items:center;}
.px-home-card b{font-size:14.5px;font-weight:800;color:${C.ink};}
.px-home-card-d{font-size:12px;line-height:1.55;color:${C.sub};min-height:36px;}
.px-home-card-go{display:flex;align-items:center;gap:5px;font-size:12px;font-weight:700;margin-top:2px;}
.px-home-foot{text-align:center;margin-top:40px;font-size:11px;color:${C.faint};letter-spacing:.4px;}
/* ---- Início 2.0: caminho guiado ---- */
.px-caminho{max-width:940px;margin:0 auto;padding:36px 28px 60px;animation:fadeUp .35s ease;}
.px-caminho-hero{text-align:center;margin:0 auto 30px;}
.px-caminho-hero h1{font-size:27px;line-height:1.25;font-weight:800;color:${C.navy};letter-spacing:-.4px;margin:0 0 12px;}
.px-caminho-hero p{font-size:14px;line-height:1.6;color:${C.sub};margin:0 auto 18px;}
.px-caminho-hero p b{color:${C.primaryDark};}
/* ---- os 4 cards do caminho (Início) ---- */
.px-etapas-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;}
@media(max-width:900px){.px-etapas-grid{grid-template-columns:repeat(2,1fr);}}
@media(max-width:560px){.px-etapas-grid{grid-template-columns:1fr;}}
.px-etapa-card{display:flex;flex-direction:column;align-items:flex-start;gap:8px;text-align:left;background:#fff;
  border:1px solid ${C.line};border-radius:15px;padding:18px 18px 15px;cursor:pointer;font-family:inherit;transition:all .16s;}
.px-etapa-card:hover{border-color:${C.primary};box-shadow:0 8px 26px rgba(19,49,92,.1);transform:translateY(-2px);}
.px-etapa-card-top{display:flex;align-items:center;gap:9px;width:100%;}
.px-etapa-card-n{width:22px;height:22px;border-radius:50%;display:grid;place-items:center;background:${C.primarySoft};
  color:${C.primary};font-size:12px;font-weight:800;flex-shrink:0;}
.px-etapa-card.feito .px-etapa-card-n{background:${C.greenSoft};color:${C.green};}
.px-etapa-card-ic{width:34px;height:34px;border-radius:10px;display:grid;place-items:center;background:${C.bg};color:${C.navy};margin-left:auto;}
.px-etapa-card b{font-size:14px;font-weight:800;color:${C.ink};line-height:1.3;}
.px-etapa-card-d{font-size:12px;line-height:1.5;color:${C.sub};min-height:54px;}
.px-etapa-card-go{display:flex;align-items:center;gap:5px;font-size:12px;font-weight:700;color:${C.primary};margin-top:2px;}
.px-etapa-card.feito .px-etapa-card-go{color:${C.green};}
/* ---- tira compacta das etapas (demais telas) ---- */
.px-etapas-tira{display:flex;gap:6px;background:#fff;border-bottom:1px solid ${C.line};padding:7px 24px;
  overflow-x:auto;position:sticky;top:57px;z-index:94;}
.px-etapa-min{display:inline-flex;align-items:center;gap:7px;background:none;border:1px solid transparent;border-radius:9px;
  padding:6px 11px;font-family:inherit;font-size:12.5px;font-weight:700;color:${C.faint};cursor:pointer;white-space:nowrap;transition:all .14s;}
.px-etapa-min:hover{background:${C.bg};color:${C.navy};}
.px-etapa-min.on{background:${C.primarySoft};border-color:${C.primary};color:${C.primaryDark};}
.px-etapa-min-n{width:18px;height:18px;border-radius:50%;display:grid;place-items:center;background:${C.line};color:${C.sub};font-size:10.5px;font-weight:800;flex-shrink:0;}
.px-etapa-min.on .px-etapa-min-n{background:${C.primary};color:#fff;}
.px-etapa-min.feito .px-etapa-min-n{background:${C.green};color:#fff;}
.px-etapa-min-t{max-width:180px;overflow:hidden;text-overflow:ellipsis;}
@media(max-width:760px){.px-etapa-min-t{display:none;}}
/* ---- criação: regra por natureza e marcas de nível novo ---- */
.px-nova-regra{display:flex;align-items:flex-start;gap:7px;font-size:11.5px;line-height:1.5;border-radius:9px;
  padding:9px 11px;margin:4px 0 14px;}
.px-nova-regra.aberta{background:${C.greenSoft};color:${C.green};}
.px-nova-regra.restrita{background:#FBF3DC;color:#9A6A00;}
.px-nova-tag{display:inline-flex;align-items:center;gap:3px;font-size:9.5px;font-weight:800;padding:2px 6px;
  border-radius:20px;margin-left:6px;text-transform:none;letter-spacing:0;}
.px-nova-tag.nova{background:${C.primarySoft};color:${C.primary};}
.px-nova-tag.trava{background:#FBEDE4;color:#C2410C;}
.px-nova-cand-niv{display:flex;gap:3px;margin-left:auto;flex-shrink:0;}
.px-nova-cand-niv span{width:16px;height:16px;border-radius:4px;display:grid;place-items:center;background:${C.bg};
  color:${C.faint};font-size:9px;font-weight:800;}
.px-nova-cand-niv span.hit{background:${C.primary};color:#fff;}
/* ---- trilha dos 4 níveis do catálogo ---- */
.px-trilha{display:flex;align-items:center;gap:2px;flex-wrap:wrap;background:#fff;border:1px solid ${C.line};
  border-radius:11px;padding:7px 10px;margin-bottom:16px;}
.px-trilha-sep{color:${C.line};flex-shrink:0;}
.px-trilha-p{display:flex;flex-direction:column;align-items:flex-start;gap:1px;background:none;border:none;
  border-radius:7px;padding:4px 9px;font-family:inherit;cursor:pointer;text-align:left;transition:all .14s;}
.px-trilha-p:hover{background:${C.bg};}
.px-trilha-rot{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.04em;color:${C.faint};}
.px-trilha-p.on .px-trilha-rot{color:${C.primary};}
.px-trilha-p.feito .px-trilha-rot{color:${C.green};}
.px-trilha-val{font-size:11.5px;font-weight:700;color:${C.ink};max-width:190px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.px-trilha-limpar{display:inline-flex;align-items:center;gap:4px;margin-left:auto;background:none;border:none;
  font-family:inherit;font-size:11px;font-weight:700;color:${C.faint};cursor:pointer;padding:4px 6px;}
.px-trilha-limpar:hover{color:${C.coral||"#9E3B1F"};}
/* ---- alternador de escopo (só selecionados / todo o catálogo) ---- */
.px-escopo{display:inline-flex;gap:2px;background:${C.bg};border:1px solid ${C.line};border-radius:9px;padding:3px;}
.px-escopo button{border:none;background:none;font-family:inherit;font-size:11.5px;font-weight:700;color:${C.faint};
  padding:5px 11px;border-radius:7px;cursor:pointer;transition:all .14s;white-space:nowrap;}
.px-escopo button:hover{color:${C.navy};}
.px-escopo button.on{background:#fff;color:${C.primaryDark};box-shadow:0 1px 3px rgba(19,49,92,.12);}
/* ---- linha de um item de nível (caixinha + nome clicável) ---- */
.px-nivel-row{display:flex;align-items:center;gap:8px;padding:2px 0 2px 6px;}
.px-nivel-row.fora .px-nivel-t{color:${C.faint};text-decoration:line-through;}
.px-nivel-row.fora .px-dot{opacity:.35;}
.px-nivel-check{width:17px;height:17px;flex-shrink:0;border:1.5px solid ${C.line};border-radius:5px;background:#fff;
  display:grid;place-items:center;cursor:pointer;color:#fff;transition:all .14s;}
.px-nivel-check:hover{border-color:${C.primary};}
.px-nivel-check.on{background:${C.green};border-color:${C.green};}
.px-nivel-nome{flex:1;display:flex;align-items:center;gap:8px;background:none;border:none;border-radius:8px;
  padding:7px 9px;font-family:inherit;cursor:pointer;text-align:left;min-width:0;transition:background .14s;}
.px-nivel-nome:hover{background:${C.bg};}
.px-nivel-t{flex:1;font-size:12.5px;font-weight:600;color:${C.ink};overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.px-nivel-go{color:${C.faint};flex-shrink:0;}
.px-nivel-nome:hover .px-nivel-go{color:${C.primary};}
/* ---- Início: duas portas de entrada (marcos referenciais / conversor) ---- */
.px-porta{display:inline-flex;gap:3px;background:${C.bg};border:1px solid ${C.line};border-radius:10px;padding:3px;margin-bottom:18px;}
.px-porta button{display:inline-flex;align-items:center;gap:7px;border:none;background:none;font-family:inherit;
  font-size:12.5px;font-weight:700;color:${C.faint};padding:7px 14px;border-radius:8px;cursor:pointer;transition:all .14s;}
.px-porta button:hover{color:${C.navy};}
.px-porta button.on{background:#fff;color:${C.primaryDark};box-shadow:0 1px 3px rgba(19,49,92,.12);}
/* ---- nível 4: alternador lista / árvore ---- */
.px-vista-row{display:flex;align-items:center;gap:12px;flex-wrap:wrap;padding:16px 24px 0;}
.px-vista-row .px-trilha{flex:1;margin-bottom:0;min-width:280px;}
.px-vista{display:inline-flex;gap:2px;background:${C.bg};border:1px solid ${C.line};border-radius:9px;padding:3px;flex-shrink:0;}
.px-vista button{display:inline-flex;align-items:center;gap:6px;border:none;background:none;font-family:inherit;
  font-size:11.5px;font-weight:700;color:${C.faint};padding:6px 12px;border-radius:7px;cursor:pointer;transition:all .14s;}
.px-vista button:hover{color:${C.navy};}
.px-vista button.on{background:#fff;color:${C.primaryDark};box-shadow:0 1px 3px rgba(19,49,92,.12);}
/* ---- etapa 1: cadastro da unidade no topo dos marcos referenciais ---- */
.px-etapa1-cad{background:#fff;border:1px solid ${C.line};border-radius:12px;padding:14px 16px;margin-bottom:20px;}
.px-etapa1-cad-h{display:flex;align-items:center;gap:7px;font-size:13px;color:${C.navy};margin-bottom:2px;}
.px-etapa1-cad-h span{font-size:11px;font-weight:600;color:${C.faint};margin-left:auto;}
.px-extra-rot{width:100%;border:none;border-bottom:1px dashed ${C.line};background:none;font-family:inherit;
  font-size:11px;font-weight:700;color:${C.navy};padding:0 0 4px;margin-bottom:6px;opacity:.85;}
.px-extra-rot:focus{outline:none;border-bottom-color:${C.primary};opacity:1;}
.px-extra-add{display:inline-flex;align-items:center;gap:7px;border:1px dashed ${C.line};border-radius:9px;
  padding:8px 13px;font-size:12.5px;font-weight:700;color:${C.sub};cursor:pointer;margin-bottom:14px;transition:all .14s;}
.px-extra-add:hover{border-color:${C.primary};color:${C.primaryDark};border-style:solid;}
/* ---- modal de órgão/unidade ---- */
.px-unid-modal{background:#fff;border-radius:16px;padding:24px;width:min(560px,100%);box-shadow:0 18px 50px rgba(19,49,92,.25);}
.px-unid-h{display:flex;align-items:center;gap:8px;font-size:15px;color:${C.navy};margin-bottom:6px;}
.px-unid-x{margin-left:auto;background:none;border:none;cursor:pointer;color:${C.faint};display:grid;place-items:center;}
.px-unid-x:hover{color:${C.ink};}
.px-unid-d{font-size:12.5px;color:${C.sub};line-height:1.55;margin:0 0 18px;}
.px-caminho-form{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:24px;}
@media(max-width:560px){.px-caminho-form{grid-template-columns:1fr;}}
.px-caminho-form label{display:block;font-size:11.5px;font-weight:700;color:${C.navy};margin-bottom:6px;}
.px-caminho-form input{width:100%;padding:10px 12px;border:1px solid ${C.line};border-radius:9px;font-size:14px;font-family:inherit;background:#fbfcfd;box-sizing:border-box;}
.px-caminho-form input:focus{outline:2px solid ${C.primarySoft};border-color:${C.primary};}
.px-caminho-foot{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;}
.px-caminho-skip{background:none;border:none;font-family:inherit;font-size:12.5px;font-weight:700;color:${C.faint};cursor:pointer;padding:6px 0;}
.px-caminho-skip:hover{color:${C.sub};}
.px-caminho-go{display:inline-flex;align-items:center;gap:8px;background:${C.primary};color:#fff;border:none;border-radius:9px;padding:12px 20px;font-size:14px;font-weight:700;cursor:pointer;font-family:inherit;}
.px-caminho-go:hover{background:${C.primaryDark};}
.px-caminho-go:disabled{opacity:.4;cursor:default;}
/* ---- seção PGD ---- */
.px-pgd-switch{display:inline-flex;gap:4px;background:#fff;border:1px solid ${C.line};border-radius:11px;padding:4px;margin-bottom:14px;}
.px-pgd-switch button{display:flex;align-items:center;gap:8px;border:none;background:none;font-family:inherit;font-size:13px;font-weight:700;color:${C.faint};padding:8px 15px;border-radius:8px;cursor:pointer;transition:all .14s;}
.px-pgd-switch button:hover{color:${C.navy};}
.px-pgd-switch button.on{background:${C.navy};color:#fff;}
.px-pgd-switch button em{font-style:normal;font-size:9px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;background:rgba(255,255,255,.18);border-radius:5px;padding:2px 6px;}
.px-pgd-switch button:not(.on) em{background:${C.greenSoft};color:${C.green};}
.px-pgd-frame-card{background:#fff;border:1px solid ${C.line};border-radius:14px;overflow:hidden;box-shadow:0 4px 18px rgba(19,49,92,.06);}
.px-pgd-frame{width:100%;height:calc(100vh - 196px);min-height:560px;border:none;display:block;}
/* nav global substitui os botões internos de voltar */
.od-back,.cr-back{display:none !important;}
/* ---- botão destaque do conversor na barra de modos ---- */
.px-mode.cta-conv{border-color:#B9D6BE;color:${C.green};background:#fff;font-weight:800;}
.px-mode.cta-conv:hover{background:${C.greenSoft};border-color:${C.green};color:${C.green};}
.px-mode.cta-conv.on{background:${C.green};border-color:${C.green};color:#fff;}
.px-modes-sep{width:1px;align-self:stretch;background:${C.line};margin:0 2px;}
.px-mode.cta-bot{border-color:${C.primary};color:${C.primary};font-weight:800;}
.px-mode.cta-bot:hover{background:${C.primarySoft};color:${C.primary};}
.px-mode.cta-bot.on{background:${C.primary};border-color:${C.primary};color:#fff;}
.px-mode.cta-pgd{border-color:#5B3A9B;color:#5B3A9B;font-weight:800;}
.px-mode.cta-pgd:hover{background:#F0EBF8;color:#5B3A9B;}
.px-mode.cta-pgd.on{background:#5B3A9B;border-color:#5B3A9B;color:#fff;}
/* ---- Conversor Unificado ---- */
.px-cu{animation:fadeUp .3s ease;}
.px-cu-hero{text-align:center;max-width:640px;margin:26px auto 22px;}
.px-cu-hero-ic{width:52px;height:52px;border-radius:15px;background:${C.greenSoft};color:${C.green};display:grid;place-items:center;margin:0 auto 14px;}
.px-cu-hero h2{font-size:23px;font-weight:800;color:${C.navy};line-height:1.35;margin:0 0 10px;}
.px-cu-hero h2 em{font-style:normal;color:${C.green};}
.px-cu-hero p{font-size:13px;color:${C.sub};line-height:1.6;margin:0;}
.px-cu-input{max-width:640px;margin:0 auto;background:#fff;border:1px solid ${C.line};border-radius:15px;padding:18px;box-shadow:0 5px 22px rgba(19,49,92,.06);}
.px-cu-drop{display:flex;gap:12px;align-items:center;border:1.5px dashed ${C.primary};border-radius:11px;padding:13px 16px;cursor:pointer;color:${C.primary};background:#FAFCFF;transition:background .15s;}
.px-cu-drop:hover{background:#F0F6FF;}
.px-cu-drop b{font-size:13px;display:block;}
.px-cu-drop span{font-size:11px;color:${C.faint};}
.px-cu-input textarea{width:100%;box-sizing:border-box;min-height:120px;margin-top:12px;border:1px solid ${C.line};border-radius:10px;padding:12px;font-family:inherit;font-size:13px;resize:vertical;}
.px-cu-acts{display:flex;gap:10px;align-items:center;margin-top:12px;flex-wrap:wrap;}
.px-cu-go{display:flex;align-items:center;gap:8px;border:none;background:${C.green};color:#fff;font-family:inherit;font-size:13px;font-weight:800;padding:11px 18px;border-radius:10px;cursor:pointer;}
.px-cu-go:hover{filter:brightness(1.07);}
.px-cu-go:disabled{opacity:.5;cursor:default;}
.px-cu-chat{display:flex;align-items:center;gap:7px;border:1px solid ${C.line};background:#fff;color:${C.primary};font-family:inherit;font-size:12px;font-weight:700;padding:10px 14px;border-radius:10px;cursor:pointer;}
.px-cu-chat:hover{border-color:${C.primary};}
.px-cu-chat.mini{padding:8px 11px;font-size:11px;}
.px-cu-topo{display:flex;align-items:center;gap:14px;flex-wrap:wrap;background:#fff;border:1px solid ${C.line};border-radius:13px;padding:11px 14px;margin-bottom:14px;position:sticky;top:63px;z-index:20;box-shadow:0 4px 14px rgba(19,49,92,.05);}
.px-cu-voltar{display:flex;align-items:center;gap:5px;border:none;background:none;color:${C.faint};font-family:inherit;font-size:12px;font-weight:700;cursor:pointer;}
.px-cu-voltar:hover{color:${C.navy};}
.px-cu-resumo{flex:1;font-size:11.5px;color:${C.sub};font-weight:600;min-width:180px;}
.px-cu-topo-acts{display:flex;gap:8px;align-items:center;flex-wrap:wrap;}
.px-cu-dl{display:flex;align-items:center;gap:6px;border:1px solid ${C.line};background:#fff;color:${C.sub};font-family:inherit;font-size:11.5px;font-weight:700;padding:9px 12px;border-radius:9px;cursor:pointer;}
.px-cu-dl:hover{border-color:${C.primary};color:${C.primary};}
.px-cu-row{display:grid;grid-template-columns:34px 1fr 120px 1.2fr;gap:14px;align-items:start;background:#fff;border:1px solid ${C.line};border-radius:12px;padding:13px 15px;margin-bottom:10px;}
@media(max-width:860px){.px-cu-row{grid-template-columns:30px 1fr;}.px-cu-conf,.px-cu-cat{grid-column:2;}}
.px-cu-row.ok{border-color:${C.green};background:#F7FBF8;}
.px-cu-check{display:flex;align-items:center;justify-content:center;padding-top:4px;}
.px-cu-check input{width:16px;height:16px;accent-color:${C.green};cursor:pointer;}
.px-cu-lbl{font-size:9px;letter-spacing:.08em;text-transform:uppercase;color:${C.faint};font-weight:800;display:block;margin-bottom:3px;}
.px-cu-pgd div{font-size:12.5px;color:${C.ink};line-height:1.5;}
.px-cu-conf{display:flex;flex-direction:column;gap:4px;align-items:flex-start;padding-top:2px;}
.px-cu-conf b{font-size:15px;font-weight:800;}
.px-cu-meter{width:76px;height:5px;border-radius:3px;background:${C.bg};overflow:hidden;display:block;}
.px-cu-meter i{display:block;height:100%;border-radius:3px;}
.px-cu-conf em{font-style:normal;font-size:9.5px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;color:${C.faint};}
.px-cu-nome{font-size:12.5px;font-weight:700;color:${C.navy};line-height:1.45;}
.px-cu-hier{font-size:10.5px;color:${C.sub};margin-top:2px;}
.px-cu-alts{display:flex;gap:6px;flex-wrap:wrap;margin-top:8px;}
.px-cu-alts button{border:1px solid ${C.line};background:#fff;border-radius:7px;padding:4px 9px;font-family:inherit;font-size:10.5px;font-weight:600;color:${C.sub};cursor:pointer;max-width:240px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.px-cu-alts button:hover{border-color:${C.primary};color:${C.primary};}
.px-cu-alts button.on{background:${C.navy};border-color:${C.navy};color:#fff;}
.px-cu-alts button.nv{border-color:#B7A3DC;color:#5B3A9B;}
.px-cu-alts button.nv:hover{background:#F4EFFB;}
.px-cu-nova{display:flex;align-items:center;gap:8px;font-size:12px;color:#5B3A9B;font-weight:600;flex-wrap:wrap;}
.px-cu-nova button{border:none;background:none;font-family:inherit;font-size:10.5px;color:${C.faint};text-decoration:underline;cursor:pointer;}
.px-cu-rodape{font-size:10.5px;color:${C.faint};font-weight:500;text-align:center;padding:10px 0 30px;}
.px-logo{display:flex;align-items:center;gap:11px;}
.px-logo-w{font-size:18px;font-weight:800;color:${C.navy};letter-spacing:-.01em;line-height:1;}
.px-logo-s{font-size:11px;color:${C.sub};margin-top:2px;}
.px-org{display:flex;align-items:center;gap:6px;font-size:12px;color:${C.sub};background:none;
  border:1px solid ${C.line};border-radius:8px;padding:6px 10px;font-family:inherit;font-weight:600;cursor:pointer;transition:all .14s;}
.px-org:hover{border-color:${C.primary};color:${C.primaryDark};background:${C.primarySoft};}
.px-org.vazio{background:#fff;border:1px dashed ${C.line};border-radius:20px;padding:5px 12px;cursor:pointer;font-weight:600;font-family:inherit;color:${C.primary};}
.px-org.vazio:hover{border-color:${C.primary};background:${C.primarySoft};}
.px-limpachip{display:flex;align-items:center;gap:6px;font-size:12px;font-weight:700;color:#fff;background:linear-gradient(100deg,#168821,#0f6b1a);border:none;border-radius:20px;padding:6px 14px;cursor:pointer;text-decoration:none;box-shadow:0 2px 8px rgba(22,136,33,.3);transition:.15s;}
.px-limpachip:hover{box-shadow:0 4px 14px rgba(22,136,33,.45);transform:translateY(-1px);}
.px-wrap{padding:16px 24px 40px;}
.px-controls{display:flex;gap:12px;align-items:center;margin-bottom:12px;flex-wrap:wrap;}
.px-search{flex:1;min-width:280px;display:flex;align-items:center;gap:10px;border:1.5px solid ${C.line};border-radius:11px;padding:11px 14px;background:#fff;}
.px-search:focus-within{border-color:${C.primary};box-shadow:0 0 0 3px ${C.primarySoft};}
.px-search input{flex:1;border:none;outline:none;font-size:14px;background:transparent;}
.px-clear{border:none;background:${C.line};color:${C.sub};border-radius:50%;width:21px;height:21px;display:grid;place-items:center;cursor:pointer;}
.px-pgdchip{display:flex;align-items:center;gap:7px;flex-shrink:0;font-size:12.5px;font-weight:700;color:${C.primaryDark};background:${C.primarySoft};border:1.5px solid #C9D8F0;border-radius:11px;padding:10px 14px;cursor:pointer;font-family:inherit;transition:.14s;}
.px-pgdchip:hover{border-color:${C.primary};background:#fff;box-shadow:0 3px 12px rgba(19,81,180,.12);}
.px-modebar{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:14px;padding:7px 10px;background:${C.surface};border:1px solid ${C.line};border-radius:13px;}
.px-modebar-l{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.06em;color:${C.faint};}
.px-modebar-sep{display:flex;align-items:center;color:#C4CDD9;}
.px-modes{display:flex;gap:4px;background:#fff;border:1px solid ${C.line};border-radius:10px;padding:4px;}
.px-mode{display:flex;align-items:center;gap:6px;font-size:12.5px;font-weight:600;padding:8px 13px;border-radius:8px;border:none;background:transparent;color:${C.sub};cursor:pointer;font-family:inherit;}
.px-mode:hover{background:${C.bg};color:${C.ink};}
.px-mode.on{background:${C.navy};color:#fff;}
.px-mode.compor{font-weight:700;color:${C.primaryDark};}
.px-enquadrar-btn{margin-left:auto;display:flex;align-items:center;gap:7px;font-family:inherit;font-size:12.5px;font-weight:700;color:#fff;background:linear-gradient(100deg,${C.primary},#0f6b1a);border:none;border-radius:10px;padding:9px 15px;cursor:pointer;box-shadow:0 2px 8px rgba(19,81,180,.25);transition:.15s;}
.px-enquadrar-btn:hover{background:linear-gradient(100deg,${C.primaryDark},#0d5c16);box-shadow:0 4px 14px rgba(19,81,180,.38);transform:translateY(-1px);}
.px-mode.compor:hover{background:${C.primarySoft};}
.px-mode.compor.on{background:${C.primary};color:#fff;}
.px-facets{display:flex;align-items:center;gap:8px;margin-bottom:10px;flex-wrap:wrap;padding-bottom:2px;}
.px-facets-l{font-size:12px;font-weight:700;color:${C.sub};}
.px-chip{display:flex;align-items:center;gap:6px;font-size:12px;font-weight:600;padding:6px 11px;line-height:1.15;border-radius:20px;border:1.5px solid ${C.line};background:#fff;color:${C.sub};cursor:pointer;}
.px-chip b{font-weight:800;opacity:.8;}
.px-chip.on{background:${C.navy};border-color:${C.navy};color:#fff;}
.px-chip-macro{display:flex;align-items:center;gap:6px;font-size:12px;font-weight:600;padding:6px 11px 6px 13px;line-height:1.15;border-radius:20px;border:1.5px solid ${C.line};background:#fff;color:${C.sub};cursor:pointer;transition:.14s;}
.px-chip-macro:hover{border-color:${C.primary};color:${C.primaryDark};}
.px-chip-macro.ativo{background:${C.primarySoft};border-color:${C.primary};color:${C.primaryDark};}
.px-chip-macro span{flex:1;}
.px-chip-macro em{font-style:normal;display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;border-radius:50%;background:rgba(19,81,180,.15);color:${C.primary};margin-left:2px;}
.px-chip-macro em:hover{background:${C.primary};color:#fff;}
.px-count{margin-left:auto;font-size:12px;color:${C.faint};white-space:nowrap;}
.px-count.filt{color:${C.primaryDark};font-weight:700;}
.px-search.full{padding:13px 16px;border-radius:12px;}
.px-search.full input{font-size:15px;}
.px-fbar{display:flex;align-items:center;gap:8px;margin-bottom:10px;flex-wrap:wrap;padding-bottom:2px;}
.px-modes-row{display:flex;margin-top:4px;}
.px-mode-badge{margin-left:2px;font-size:10px;font-weight:800;background:${C.primary};color:#fff;border-radius:10px;padding:1px 6px;line-height:1.4;}
.px-filtros-ativos{display:flex;align-items:center;gap:9px;flex-wrap:wrap;margin-bottom:14px;background:${C.primarySoft};border:1px solid ${C.line};border-radius:10px;padding:8px 12px;}
.px-fa-tags{display:flex;gap:6px;flex-wrap:wrap;flex:1;}
.px-fa-tag{display:inline-flex;align-items:center;gap:4px;font-size:11.5px;font-weight:600;padding:3px 8px 3px 10px;border-radius:20px;background:#fff;border:1px solid ${C.line};color:${C.ink};max-width:260px;overflow:hidden;}
.px-fa-tag button{display:inline-flex;align-items:center;justify-content:center;width:15px;height:15px;border-radius:50%;border:none;background:${C.line};color:${C.sub};cursor:pointer;padding:0;flex-shrink:0;}
.px-fa-tag button:hover{background:#c0c8d4;color:${C.ink};}
.px-fa-limpar{border:none;background:transparent;font-size:12px;font-weight:700;color:${C.primary};cursor:pointer;white-space:nowrap;flex-shrink:0;}
.px-fa-limpar:hover{text-decoration:underline;}

.px-body{display:grid;grid-template-columns:1fr 320px;gap:16px;align-items:start;}
.px-body.solo{grid-template-columns:1fr;}
.px-main{min-height:55vh;}
.px-empty{background:#fff;border:1px dashed ${C.line};border-radius:12px;padding:36px;text-align:center;color:${C.sub};}
.px-group{margin-bottom:22px;}
.px-group-h{display:flex;align-items:center;gap:9px;flex-wrap:wrap;}
.px-tag{font-size:10px;font-weight:700;padding:3px 9px;border-radius:20px;text-transform:uppercase;letter-spacing:.04em;}
.px-crumb{font-size:11px;color:${C.faint};display:flex;align-items:center;gap:3px;}
.px-group-t{font-size:15.5px;font-weight:800;color:${C.navy};margin:7px 0 11px;}
.px-cards{display:grid;grid-template-columns:1fr 1fr;gap:11px;}
.px-card{background:#fff;border:1px solid ${C.line};border-radius:13px;padding:14px 15px;transition:.14s;}
.px-card:hover{border-color:#cdd6e3;box-shadow:0 4px 14px rgba(19,49,92,.07);}
.px-card.ok{border-color:#bfe2c5;background:#fbfdfb;}
.px-card-top{display:flex;align-items:center;gap:8px;}
.px-code{font-family:'JetBrains Mono',ui-monospace,monospace;font-size:11.5px;font-weight:700;color:${C.primaryDark};background:${C.primarySoft};padding:2px 7px;border-radius:5px;}
.px-sit{font-size:9.5px;font-weight:700;padding:2px 8px;border-radius:20px;text-transform:uppercase;letter-spacing:.03em;}
.px-detbtn{margin-left:auto;display:flex;align-items:center;gap:4px;border:none;background:transparent;color:${C.faint};font-size:11px;font-weight:600;cursor:pointer;}
.px-detbtn:hover{color:${C.primary};}
.px-card-name{font-size:14px;font-weight:700;color:${C.ink};margin:9px 0 6px;line-height:1.3;}
.px-card-ativ{font-size:11.5px;color:${C.sub};line-height:1.4;}
.px-card-ativ span{color:${C.faint};font-weight:700;text-transform:uppercase;font-size:9px;letter-spacing:.05em;display:block;margin-bottom:1px;}
.px-card-foot{display:flex;align-items:center;justify-content:space-between;margin-top:11px;}
.px-card-areas{display:flex;align-items:center;gap:5px;font-size:11px;color:${C.faint};}
.px-add{display:flex;align-items:center;gap:6px;border:1.5px solid ${C.primary};background:#fff;color:${C.primary};font-size:12px;font-weight:700;padding:7px 12px;border-radius:8px;cursor:pointer;}
.px-add:hover{background:${C.primary};color:#fff;}
.px-add.done{background:${C.greenSoft};border-color:#bfe2c5;color:${C.green};}
.px-det{margin-top:12px;padding-top:12px;border-top:1px solid ${C.line};}
.px-det p{font-size:12px;color:${C.sub};margin:0 0 9px;line-height:1.5;}
.px-det-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;}
.px-det-grid div{font-size:11.5px;color:${C.ink};}
.px-det-grid span{display:block;font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:${C.faint};margin-bottom:1px;}

/* processo */
.px-proc{display:flex;flex-direction:column;gap:18px;}
.px-proc-macro{background:#fff;border:1px solid ${C.line};border-radius:13px;padding:15px 16px;}
.px-proc-h{display:flex;align-items:center;gap:9px;margin-bottom:13px;font-size:13.5px;color:${C.navy};}
.px-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0;transition:box-shadow .15s;}
.px-dot.sel{box-shadow:0 0 0 3px rgba(22,136,33,.22);}
.px-proc-h .px-dot{width:9px;height:9px;} /* reset for processo mode */
/* nível NATUREZA */
.px-natacc{margin-bottom:10px;border:1px solid ${C.line};border-radius:13px;overflow:hidden;background:#fff;}
.px-nathead{display:flex;align-items:center;gap:11px;width:100%;text-align:left;border:none;border-left:4px solid transparent;background:#fff;padding:15px 16px;cursor:pointer;color:${C.navy};}
.px-nathead:hover{background:${C.bg};}
.px-nathead-dot{width:12px;height:12px;border-radius:50%;flex-shrink:0;}
.px-nathead-t{font-size:15px;font-weight:800;flex-shrink:0;}
.px-nathead-meta{font-size:12px;color:${C.faint};font-weight:600;margin-left:6px;flex:1;}
.px-natbody{padding:6px 12px 12px;background:${C.surface};}
/* nível MACROPROCESSO */
.px-macacc{margin-bottom:6px;border:1px solid ${C.line};border-radius:10px;overflow:hidden;background:#fff;}
.px-machead{display:flex;align-items:center;gap:9px;width:100%;text-align:left;border:none;background:#fff;padding:11px 13px;cursor:pointer;color:${C.navy};}
.px-machead:hover{background:${C.bg};}
.px-machead-t{font-size:13px;font-weight:700;flex:1;line-height:1.3;}
.px-machead-meta{font-size:11px;color:${C.faint};font-weight:600;flex-shrink:0;}
.px-macbody{padding:5px 10px 9px;background:${C.surface};}
.px-catsel.sm{font-size:9.5px;padding:1px 7px;}
/* nível CATEGORIA (dentro do macro) */
.px-macbody .px-catacc{margin-bottom:5px;}
.px-macbody .px-cathead{padding:9px 12px;background:#fff;}
.px-macbody .px-cathead:hover{background:${C.bg};}
.px-macbody .px-cathead-t{font-size:12.5px;font-weight:600;}
.px-catacc{margin-bottom:7px;border:1px solid ${C.line};border-radius:11px;overflow:hidden;background:#fff;}
.px-cathead{display:flex;align-items:center;gap:10px;width:100%;text-align:left;border:none;background:${C.bg};padding:11px 14px;cursor:pointer;color:${C.navy};}
.px-cathead:hover{background:#eceff4;}
.px-cathead-t{font-size:13px;font-weight:700;flex:1;line-height:1.3;}
.px-catcount{font-size:11px;color:${C.faint};font-weight:600;flex-shrink:0;}
.px-catsel{display:inline-flex;align-items:center;gap:3px;font-size:10.5px;font-weight:700;color:${C.green};background:${C.greenSoft};border:1px solid #bfe2c5;border-radius:20px;padding:2px 8px;flex-shrink:0;}
.px-stage{min-width:230px;flex:1;background:${C.bg};border:1px solid ${C.line};border-radius:11px;padding:11px;}
.px-stage-t{font-size:11px;font-weight:700;color:${C.sub};text-transform:uppercase;letter-spacing:.04em;margin-bottom:9px;}
.px-stage-item{display:flex;align-items:center;gap:7px;width:100%;text-align:left;background:#fff;border:1px solid ${C.line};border-radius:8px;padding:8px 9px;margin-bottom:6px;cursor:pointer;transition:.12s;}
.px-stage-item:hover{border-color:${C.primary};}
.px-stage-item.ok{border-color:#bfe2c5;background:#fbfdfb;}
.px-stage-item code{font-family:'JetBrains Mono',monospace;font-size:10px;font-weight:700;color:${C.primaryDark};flex-shrink:0;}
.px-stage-item span{font-size:11px;color:${C.ink};line-height:1.25;}
.px-stage-arrow{display:flex;align-items:center;color:${C.faint};flex-shrink:0;}

/* panorama */
.px-pan{background:#fff;border:1px solid ${C.line};border-radius:14px;padding:18px;}
/* ---- árvore de decomposição ---- */
.px-arv{background:#fff;border:1px solid ${C.line};border-radius:14px;padding:18px;}
.px-arv-h{margin-bottom:14px;}
.px-arv-h h3{font-size:15px;font-weight:800;color:${C.navy};margin:0 0 3px;}
.px-arv-h p{font-size:12px;color:${C.faint};line-height:1.5;margin:0;}
.px-arv-crumb{display:flex;align-items:center;gap:3px;flex-wrap:wrap;margin-top:11px;}
.px-arv-crumb>button,.px-arv-crumb-i button{border:none;background:${C.bg};font-family:inherit;font-size:11.5px;font-weight:700;color:${C.primaryDark};cursor:pointer;padding:4px 10px;border-radius:20px;}
.px-arv-crumb>button.on{background:${C.primary};color:#fff;}
.px-arv-crumb>button:hover,.px-arv-crumb-i button:hover{background:${C.primarySoft};}
.px-arv-crumb-i{display:inline-flex;align-items:center;gap:3px;color:${C.faint};}
.px-arv-flow{display:flex;gap:14px;overflow-x:auto;padding:4px 2px 10px;align-items:flex-start;}
.px-arv-col{flex:0 0 248px;min-width:248px;background:${C.surface};border:1px solid ${C.line};border-radius:12px;padding:11px;opacity:0;transform:translateX(-8px);animation:arvIn .32s ease forwards;}
/* coluna de Entrega mais larga: nomes longos ficam legíveis, sobretudo com muitas entregas */
.px-arv-col.larga{flex:0 0 420px;min-width:420px;}
@media(max-width:1100px){.px-arv-col.larga{flex:0 0 340px;min-width:340px;}}
/* botão de recolher (−) opcional no cabeçalho da coluna já escolhida */
.px-arv-col-min{margin-left:6px;width:22px;height:22px;flex-shrink:0;border:1px solid ${C.line};background:#fff;border-radius:7px;display:inline-flex;align-items:center;justify-content:center;color:${C.sub};cursor:pointer;transition:border-color .14s,color .14s,background .14s;}
.px-arv-col-min:hover{border-color:${C.primary};color:${C.primary};background:${C.primarySoft};}
/* coluna recolhida: faixa estreita quando o nível já foi escolhido (evita a árvore crescer p/ o lado) */
.px-arv-col.recolhida{flex:0 0 46px;min-width:46px;width:46px;padding:11px 0;display:flex;flex-direction:column;align-items:center;gap:9px;cursor:pointer;background:#fff;border-color:var(--cor);transition:background .15s,box-shadow .15s;}
.px-arv-col.recolhida:hover{background:${C.bg};box-shadow:0 2px 8px rgba(13,49,111,.1);}
.px-arv-rec-lb{writing-mode:vertical-rl;transform:rotate(180deg);font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.08em;color:${C.faint};}
.px-arv-rec-v{writing-mode:vertical-rl;transform:rotate(180deg);font-size:12px;font-weight:800;color:var(--cor);max-height:230px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;}
.px-arv-rec-go{color:var(--cor);flex-shrink:0;}
@keyframes arvIn{to{opacity:1;transform:translateX(0);}}
.px-arv-col-h{display:flex;align-items:center;gap:7px;margin-bottom:9px;padding:0 3px;}
.px-arv-col-h>span:first-child{margin-right:auto;}
.px-arv-col-h>em{margin-left:auto;}
.px-arv-col-h>.px-arv-col-min{margin-left:0;}
.px-arv-col-h span{font-size:10.5px;font-weight:800;text-transform:uppercase;letter-spacing:.05em;color:${C.sub};}
.px-arv-col-h em{font-style:normal;font-size:11px;font-weight:700;color:${C.faint};}
.px-arv-col-body{display:flex;flex-direction:column;gap:5px;max-height:520px;overflow-y:auto;padding-right:3px;}
.px-arv-node{position:relative;display:flex;align-items:center;gap:7px;width:100%;text-align:left;border:1px solid ${C.line};background:#fff;border-radius:9px;padding:9px 10px;cursor:pointer;overflow:hidden;font-family:inherit;transition:border-color .14s, transform .1s;}
.px-arv-node:hover{border-color:var(--cor);transform:translateX(2px);}
.px-arv-node-bar{position:absolute;left:0;top:0;bottom:0;width:var(--barw);background:var(--cor);opacity:.1;transition:width .4s ease;}
.px-arv-node.ativo{border-color:var(--cor);box-shadow:0 0 0 2px var(--cor) inset;}
.px-arv-node.ativo .px-arv-node-bar{opacity:.18;}
.px-arv-node-tx{position:relative;flex:1;font-size:12px;font-weight:600;color:${C.ink};line-height:1.3;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.px-arv-node-val{position:relative;font-size:11px;font-weight:800;color:${C.navy};flex-shrink:0;}
.px-arv-node-go{position:relative;color:${C.faint};flex-shrink:0;}
.px-arv-node:hover .px-arv-node-go{color:var(--cor);}
.px-arv-node.leaf .px-arv-node-tx{white-space:normal;font-weight:500;}
.px-arv-node-pick{position:relative;flex-shrink:0;width:22px;height:22px;border-radius:50%;display:grid;place-items:center;background:${C.bg};color:${C.sub};}
.px-arv-node.leaf:hover .px-arv-node-pick{background:var(--cor);color:#fff;}
.px-arv-node.sel{border-color:var(--cor);background:${C.greenSoft};}
.px-arv-node.sel .px-arv-node-pick{background:${C.green};color:#fff;}
.px-arv-vazio{font-size:11.5px;color:${C.faint};padding:12px;text-align:center;}
.px-arv-mais{font-size:10.5px;color:${C.faint};padding:8px 4px 2px;text-align:center;font-style:italic;}
.px-pan-h{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;margin-bottom:14px;flex-wrap:wrap;}
.px-pan-h h3{font-size:16px;font-weight:800;color:${C.navy};margin:0 0 4px;}
.px-pan-h p{font-size:12px;color:${C.sub};margin:0;max-width:460px;line-height:1.4;}
.px-pan-leg{display:flex;gap:14px;flex-wrap:wrap;}
.px-pan-leg span{display:flex;align-items:center;gap:6px;font-size:12px;color:${C.sub};font-weight:600;}
.px-pan-leg i{width:11px;height:11px;border-radius:3px;}
.px-pan-chart{height:340px;}
.px-pan-crumb{display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin:0 0 12px;font-size:12px;}
.px-pan-crumb button{border:none;background:none;font-family:inherit;font-size:12px;font-weight:700;color:${C.primary};cursor:pointer;padding:3px 7px;border-radius:6px;transition:.15s;}
.px-pan-crumb button:hover{background:${C.primarySoft};}
.px-pan-crumb button.on{color:${C.ink};cursor:default;}
.px-pan-crumb button.on:hover{background:none;}
.px-pan-crumb svg{color:${C.faint};}
.px-pan-crumb-cur{font-weight:700;color:${C.ink};padding:3px 2px;}
.px-pan-crumb-lvl{margin-left:auto;font-size:11px;font-weight:700;color:${C.faint};text-transform:uppercase;letter-spacing:.4px;}
.px-pan-note{margin-top:12px;font-size:11.5px;color:${C.faint};text-align:center;}

/* documento / descrição */
.px-doc{background:#fff;border:1px solid ${C.line};border-radius:14px;padding:16px 15px;position:sticky;top:14px;}
.px-doc-h{display:flex;align-items:center;justify-content:space-between;gap:9px;}
.px-doc-t{display:flex;align-items:center;gap:7px;font-size:15px;font-weight:800;color:${C.navy};}
.px-doc-badge{background:${C.primary};color:#fff;font-size:12px;font-weight:700;min-width:24px;height:24px;border-radius:12px;display:grid;place-items:center;padding:0 7px;}
.px-doc-meta{display:flex;gap:14px;margin:10px 0 12px;position:relative;align-items:flex-start;}
.px-doc-setmeta{display:flex;align-items:center;gap:8px;width:100%;margin:10px 0 12px;background:${C.primarySoft};border:1px dashed #b9cdf0;border-radius:10px;padding:10px 13px;font-size:12.5px;font-weight:700;color:${C.primaryDark};cursor:pointer;font-family:inherit;}
.px-doc-setmeta span{flex:1;text-align:left;}
.px-doc-setmeta:hover{border-style:solid;border-color:${C.primary};}
.px-doc-meta>div{font-size:11px;color:${C.ink};font-weight:600;}
.px-doc-meta>div>span{display:block;font-size:9px;text-transform:uppercase;letter-spacing:.05em;color:${C.faint};margin-bottom:1px;font-weight:700;}
.px-doc-meta-edit{margin-left:auto;background:#fff;border:1px solid ${C.line};border-radius:7px;width:24px;height:24px;display:grid;place-items:center;cursor:pointer;color:${C.sub};flex-shrink:0;}
.px-doc-meta-edit:hover{border-color:${C.primary};color:${C.primary};}
.px-doc-meta.edit{flex-direction:column;gap:8px;background:${C.primarySoft};border:1px solid ${C.line};border-radius:10px;padding:11px;}
.px-doc-meta.edit label{display:flex;flex-direction:column;font-size:9px;text-transform:uppercase;letter-spacing:.05em;color:${C.faint};font-weight:700;gap:3px;}
.px-doc-meta.edit input{font-size:12px;font-weight:600;color:${C.ink};border:1.5px solid ${C.line};border-radius:7px;padding:6px 9px;font-family:inherit;text-transform:none;letter-spacing:0;}
.px-doc-meta.edit input:focus{outline:none;border-color:${C.primary};}
.px-doc-meta-btns{display:flex;gap:7px;margin-top:1px;}
.px-dm-save{display:flex;align-items:center;gap:5px;background:${C.primary};color:#fff;border:none;border-radius:7px;padding:6px 12px;font-size:11.5px;font-weight:700;cursor:pointer;}
.px-dm-save:hover{background:${C.primaryDark};}
.px-dm-cancel{background:#fff;border:1px solid ${C.line};border-radius:7px;padding:6px 12px;font-size:11.5px;font-weight:600;color:${C.sub};cursor:pointer;}
.px-doc-list{min-height:80px;max-height:44vh;overflow:auto;}
.px-doc-item{display:flex;gap:9px;align-items:flex-start;padding:9px 0;border-bottom:1px solid #f1f3f6;}
.px-doc-n{font-family:'JetBrains Mono',monospace;font-size:11px;font-weight:700;color:${C.faint};padding-top:2px;}
.px-doc-body{flex:1;}
.px-doc-row{display:flex;align-items:center;gap:7px;}
.px-doc-row code{font-family:'JetBrains Mono',monospace;font-size:10.5px;font-weight:700;color:${C.primaryDark};}
.px-doc-dot{width:7px;height:7px;border-radius:50%;}
.px-doc-name{font-size:12px;color:${C.ink};line-height:1.3;margin-top:2px;}
.px-doc-item>button{border:none;background:transparent;color:${C.faint};cursor:pointer;}
.px-doc-item>button:hover{color:#c0392b;}
.px-doc-guide{padding:6px 2px 2px;}
.px-doc-guide-t{font-size:12.5px;font-weight:800;color:${C.navy};margin-bottom:12px;}
.px-doc-steps{list-style:none;padding:0;margin:0;display:flex;flex-direction:column;gap:11px;}
.px-doc-steps li{display:flex;gap:11px;align-items:flex-start;}
.px-doc-steps li>span{flex-shrink:0;width:22px;height:22px;border-radius:50%;background:${C.primarySoft};color:${C.primary};font-size:12px;font-weight:800;display:grid;place-items:center;}
.px-doc-steps li>div{display:flex;flex-direction:column;font-size:11px;color:${C.faint};line-height:1.4;}
.px-doc-steps li>div>b{font-size:12.5px;font-weight:700;color:${C.ink};margin-bottom:1px;}
.px-doc-foot{border-top:1px solid ${C.line};margin:0 -15px -16px;padding:12px 15px 15px;background:#fff;position:sticky;bottom:-16px;}
.px-doc-dist{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:11px;font-size:11px;color:${C.sub};}
.px-doc-dist-tot{font-weight:600;color:${C.ink};}
.px-doc-dist-tot b{font-weight:800;color:${C.navy};}
.px-doc-dist-sep{width:1px;height:12px;background:${C.line};}
.px-doc-dist-n{display:inline-flex;align-items:center;gap:5px;font-weight:600;}
.px-doc-dist-n i{width:8px;height:8px;border-radius:50%;}
.px-doc-dist-n.zero{opacity:.4;}
.px-inject{width:100%;display:flex;align-items:center;justify-content:center;gap:8px;background:${C.primary};color:#fff;border:none;border-radius:10px;padding:12px;font-size:13.5px;font-weight:700;cursor:pointer;}
.px-inject em{font-style:normal;font-weight:800;opacity:.9;}
.px-inject:hover{background:${C.primaryDark};}
.px-inject:disabled{opacity:.4;cursor:default;}
.px-export{width:100%;margin-top:8px;display:flex;align-items:center;justify-content:center;gap:7px;background:#fff;color:${C.primaryDark};border:1.5px solid ${C.line};border-radius:10px;padding:10px;font-size:12.5px;font-weight:700;cursor:pointer;transition:.14s;}
.px-export:hover{border-color:${C.primary};background:${C.primarySoft};}
.px-export:disabled{opacity:.4;cursor:default;}

/* modal preview */
.px-modal-bg{position:fixed;inset:0;background:rgba(19,49,92,.45);display:grid;place-items:center;z-index:70;padding:20px;}
.px-modal{background:#fff;border-radius:16px;width:680px;max-width:100%;max-height:88vh;overflow:auto;padding:22px;box-shadow:0 20px 60px rgba(0,0,0,.3);}
.px-modal-h{display:flex;justify-content:space-between;align-items:center;font-size:16px;color:${C.navy};}
.px-modal-h div{display:flex;align-items:center;gap:8px;}
.px-modal-h button{border:none;background:${C.bg};border-radius:8px;width:32px;height:32px;display:grid;place-items:center;cursor:pointer;color:${C.sub};}
.px-modal-sub{font-size:12.5px;color:${C.sub};margin:8px 0 14px;line-height:1.5;}
.px-modal-meta{display:flex;gap:18px;flex-wrap:wrap;font-size:12px;color:${C.sub};margin-bottom:12px;}
.px-modal-table{border:1px solid ${C.line};border-radius:10px;overflow:hidden;}
.px-mt-head,.px-mt-row{display:grid;grid-template-columns:90px 1fr 1fr;gap:10px;padding:9px 12px;font-size:12px;align-items:center;}
.px-mt-head{background:${C.bg};font-weight:700;color:${C.sub};font-size:10.5px;text-transform:uppercase;letter-spacing:.04em;}
.px-mt-row{border-top:1px solid ${C.line};}
.px-mt-row code{font-family:'JetBrains Mono',monospace;font-weight:700;color:${C.primaryDark};font-size:11px;}
.px-mt-macro{color:${C.faint};font-size:11px;}
.px-modal-acts{display:flex;justify-content:flex-end;gap:10px;margin-top:16px;}
.px-btn-ghost{border:1.5px solid ${C.line};background:#fff;color:${C.ink};border-radius:9px;padding:10px 16px;font-size:13px;font-weight:600;cursor:pointer;}
.px-btn-primary{display:flex;align-items:center;gap:7px;border:none;background:${C.primary};color:#fff;border-radius:9px;padding:10px 18px;font-size:13px;font-weight:700;cursor:pointer;}
.px-btn-primary:hover{background:${C.primaryDark};}

/* assistente */
.px-chat{position:fixed;bottom:22px;left:22px;width:384px;max-width:calc(100vw - 24px);height:544px;max-height:calc(100vh - 44px);background:#fff;border-radius:16px;box-shadow:0 16px 50px rgba(13,50,111,.3);display:flex;flex-direction:column;overflow:hidden;z-index:51;border:1px solid ${C.line};}
.px-chat-h{background:${C.primary};color:#fff;padding:13px 16px;display:flex;justify-content:space-between;align-items:center;}
.px-chat-ht{display:flex;align-items:center;gap:8px;font-weight:700;font-size:14px;}
.px-chat-h button{background:rgba(255,255,255,.2);border:none;color:#fff;border-radius:7px;width:28px;height:28px;display:grid;place-items:center;cursor:pointer;}
.px-chat-b{flex:1;overflow:auto;padding:15px;display:flex;flex-direction:column;gap:11px;background:${C.bg};}
.px-msg{max-width:88%;font-size:13px;line-height:1.5;}
.px-msg.user{align-self:flex-end;}
.px-msg.user .px-msg-tx{background:${C.primary};color:#fff;padding:9px 13px;border-radius:13px 13px 4px 13px;}
.px-msg.assistant .px-msg-tx{background:#fff;border:1px solid ${C.line};padding:10px 13px;border-radius:13px 13px 13px 4px;white-space:pre-wrap;}
.px-chips{display:flex;flex-wrap:wrap;gap:6px;margin-top:7px;}
.px-chip2{display:flex;align-items:center;gap:4px;background:${C.greenSoft};color:${C.green};border:1px solid #bfe2c5;border-radius:7px;padding:4px 9px;font-size:11px;font-weight:700;font-family:'JetBrains Mono',monospace;cursor:pointer;}
.px-typing{display:flex;gap:4px;background:#fff;border:1px solid ${C.line};padding:12px 14px;border-radius:13px;width:fit-content;}
.px-typing span{width:7px;height:7px;border-radius:50%;background:${C.faint};animation:pxb 1.2s infinite;}
.px-typing span:nth-child(2){animation-delay:.2s;} .px-typing span:nth-child(3){animation-delay:.4s;}
@keyframes pxb{0%,60%,100%{opacity:.25;transform:translateY(0);}30%{opacity:1;transform:translateY(-3px);}}
.px-chat-i{display:flex;gap:8px;padding:12px;border-top:1px solid ${C.line};align-items:flex-end;}
.px-chat-i textarea{flex:1;border:1.5px solid ${C.line};border-radius:10px;padding:9px 12px;font-size:13px;outline:none;resize:none;max-height:88px;font-family:inherit;}
.px-chat-i textarea:focus{border-color:${C.primary};}
.px-chat-i button{background:${C.primary};color:#fff;border:none;border-radius:9px;width:40px;height:38px;display:grid;place-items:center;cursor:pointer;flex-shrink:0;}
.px-chat-i button:disabled{opacity:.4;cursor:default;}
.px-chat-i button.px-chat-clip{background:#fff;color:${C.sub};border:1.5px solid ${C.line};}
.px-chat-i button.px-chat-clip:hover{border-color:${C.primary};color:${C.primary};}
.px-chat-anexo{display:flex;flex-wrap:wrap;gap:8px;padding:0 12px 4px;}
.px-anexo-chip{display:inline-flex;align-items:center;gap:6px;font-size:11.5px;font-weight:600;color:${C.primaryDark};background:${C.primarySoft};border:1px solid ${C.line};border-radius:20px;padding:4px 6px 4px 10px;}
.px-anexo-chip button{background:none;border:none;cursor:pointer;color:${C.sub};display:inline-flex;padding:0;width:auto;height:auto;}
.px-anexo-erro{display:inline-flex;align-items:center;gap:5px;font-size:11.5px;color:#B42318;background:#FEF3F2;border:1px solid #FECDCA;border-radius:8px;padding:4px 9px;}
.px-toast{position:fixed;bottom:84px;left:50%;transform:translateX(-50%);background:${C.navy};color:#fff;padding:11px 20px;border-radius:10px;font-size:13px;font-weight:600;box-shadow:0 8px 24px rgba(0,0,0,.25);z-index:60;}

.px-edet-section{margin-bottom:9px;}
.px-edet-section>b{display:block;font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:${C.faint};margin-bottom:3px;}
.px-edet-section p{font-size:12px;color:${C.sub};margin:0;line-height:1.5;white-space:pre-wrap;}
.px-edet-credits{display:grid;grid-template-columns:1fr 1fr;gap:7px;margin-top:9px;padding-top:9px;border-top:1px dashed ${C.line};}
.px-edet-credits span{font-size:11px;color:${C.faint};}
.px-edet-credits b{display:block;font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;color:#c0c8d4;margin-bottom:2px;}
/* catálogo expansível */
.px-tree{}
.px-treebar{display:flex;justify-content:space-between;align-items:center;margin-bottom:11px;font-size:13px;flex-wrap:wrap;gap:8px;}
.px-treebar-l b{color:${C.navy};} .px-treebar-l span{color:${C.faint};}
.px-treebar-r{display:flex;align-items:center;gap:8px;color:${C.faint};}
.px-treebar-r button{border:none;background:transparent;color:${C.primary};font-weight:700;cursor:pointer;font-size:12.5px;}
.px-treebar-r button:hover{text-decoration:underline;}
.px-catacc{margin-bottom:7px;border:1px solid ${C.line};border-radius:11px;overflow:hidden;background:#fff;}
.px-erows{padding:3px 0;}
.px-erow{border-top:1px solid #f1f3f6;}
.px-erow.sel{background:${C.greenSoft};}
.px-erow-main{display:flex;align-items:center;gap:10px;padding:9px 14px;}
.px-erow-acts{display:flex;align-items:center;gap:2px;flex-shrink:0;opacity:0;transition:opacity .15s;}
.px-erow-main:hover .px-erow-acts{opacity:1;}
.px-erow.sel .px-erow-acts{opacity:1;}
.px-cbox{width:19px;height:19px;border-radius:6px;border:1.5px solid ${C.faint};background:#fff;display:grid;place-items:center;cursor:pointer;flex-shrink:0;color:#fff;padding:0;}
.px-cbox.on{background:${C.green};border-color:${C.green};}
.px-cbox:hover{border-color:${C.primary};}
.px-ecode{font-family:'JetBrains Mono',ui-monospace,monospace;font-size:11px;font-weight:700;color:${C.primaryDark};flex-shrink:0;width:62px;}
.px-sit.sm{font-size:8.5px;padding:2px 7px;flex-shrink:0;}
.px-ename{font-size:12.5px;color:${C.ink};font-weight:600;flex:1;line-height:1.25;}
.px-eserv{font-size:10.5px;color:${C.faint};flex-shrink:0;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.px-edetbtn{border:none;background:transparent;color:${C.faint};cursor:pointer;flex-shrink:0;padding:3px;display:grid;place-items:center;}
.px-edetbtn:hover{color:${C.primary};}
.px-edet{padding:0 14px 14px 44px;}
.px-edet-load{font-size:12px;color:${C.faint};padding:6px 0 10px;font-style:italic;}
.px-edet p{font-size:12px;color:${C.sub};margin:0 0 9px;line-height:1.5;}
.px-edet-grid{display:grid;grid-template-columns:1fr 1fr;gap:4px 18px;}
.px-edet-grid span{font-size:11.5px;color:${C.ink};padding:6px 0;border-bottom:1px solid #f0f2f5;line-height:1.4;}
.px-edet-grid b{display:block;font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:${C.faint};margin-bottom:3px;}

/* header direita + chip de sinalizações */
.px-head-r{display:flex;align-items:center;gap:12px;}
.px-flagchip{display:flex;align-items:center;gap:6px;font-size:12px;font-weight:700;color:#B86E00;background:#FBEEDB;border:1px solid #EBD7B0;border-radius:20px;padding:5px 12px;cursor:pointer;}
.px-flagchip:hover{filter:brightness(.97);}

/* banner do conversor PGD */
.px-pgdbanner{display:flex;align-items:center;gap:14px;width:100%;text-align:left;border:1.5px solid #C9D8F0;background:linear-gradient(90deg,${C.primarySoft},#fff);border-radius:13px;padding:13px 16px;margin-bottom:16px;cursor:pointer;transition:.14s;}
.px-pgdbanner:hover{border-color:${C.primary};box-shadow:0 4px 16px rgba(19,81,180,.12);}
.px-pgdbanner-ic{width:38px;height:38px;border-radius:10px;background:${C.primary};color:#fff;display:grid;place-items:center;flex-shrink:0;}
.px-pgdbanner-tx{flex:1;display:flex;flex-direction:column;}
.px-pgdbanner-tx b{font-size:14px;color:${C.navy};}
.px-pgdbanner-tx span{font-size:12px;color:${C.sub};margin-top:1px;}
.px-pgdbanner-cta{display:flex;align-items:center;gap:6px;font-size:13px;font-weight:700;color:${C.primary};flex-shrink:0;}

/* botão sinalizar no detalhe */
.px-flagbtn{margin-top:11px;display:inline-flex;align-items:center;gap:6px;border:1px solid #EBD7B0;background:#FBEEDB;color:#B86E00;font-size:11px;font-weight:700;padding:6px 11px;border-radius:8px;cursor:pointer;}
.px-flagbtn:hover{filter:brightness(.97);}

/* sunburst (explosão solar) — layout grande */
.px-sun2{background:#fff;border:1px solid ${C.line};border-radius:14px;padding:20px 22px;}
.px-sun2-top{display:flex;justify-content:space-between;align-items:flex-start;gap:20px;flex-wrap:wrap;}
.px-sun2-head h3{font-size:18px;font-weight:800;color:${C.navy};margin:0 0 6px;}
.px-sun2-head p{font-size:12.5px;color:${C.sub};margin:0;line-height:1.55;max-width:620px;}
.px-sun2-leg{display:flex;gap:16px;flex-wrap:wrap;flex-shrink:0;}
.px-sun2-leg span{display:flex;align-items:center;gap:6px;font-size:12px;color:${C.sub};font-weight:600;}
.px-sun2-leg i{width:11px;height:11px;border-radius:3px;}
.px-sun2-crumb{display:flex;align-items:center;flex-wrap:wrap;gap:3px;margin:14px 0 4px;font-size:12.5px;}
.px-sun2-crumb>span{display:inline-flex;align-items:center;gap:3px;color:${C.faint};}
.px-sun2-crumb button{border:none;background:transparent;color:${C.primary};font-weight:600;cursor:pointer;font-size:12.5px;padding:2px 4px;border-radius:6px;}
.px-sun2-crumb button:hover{background:${C.primarySoft};}
.px-sun2-chartwrap{position:relative;}
.px-sun2-chart{display:block;text-align:center;padding:8px 0 4px;min-height:480px;overflow:auto;}
.px-sun2-zoom{position:absolute;top:6px;right:6px;z-index:5;display:flex;align-items:center;gap:4px;background:rgba(255,255,255,.94);border:1px solid ${C.line};border-radius:10px;padding:4px 6px;box-shadow:0 2px 8px rgba(13,49,111,.12);}
.px-sun2-zoom button{width:28px;height:28px;border:1px solid ${C.line};background:#fff;border-radius:7px;display:grid;place-items:center;cursor:pointer;color:${C.sub};}
.px-sun2-zoom button:hover:not(:disabled){border-color:${C.primary};color:${C.primary};}
.px-sun2-zoom button:disabled{opacity:.4;cursor:default;}
.px-sun2-zval{font-size:11.5px;font-weight:800;color:${C.navy};min-width:38px;text-align:center;font-variant-numeric:tabular-nums;}
.px-sun2-pick{display:flex;align-items:center;gap:7px;font-size:12.5px;color:${C.ink};background:${C.greenSoft};border:1px solid ${C.green};border-radius:999px;padding:6px 13px;margin:4px 0 0;width:fit-content;}
.px-sun2-pick.rem{background:#FBEEDB;border-color:#B86E00;}
.px-sun2-pick b{font-weight:700;}
.px-sun2-chart svg{display:inline-block;}
.px-sun2-note{font-size:11px;color:${C.faint};line-height:1.5;text-align:center;border-top:1px solid ${C.line};padding-top:12px;margin-top:6px;}
.px-sun-empty{font-size:13px;color:${C.faint};text-align:center;padding:60px 20px;}

/* conversor PGD (modal anexo) */
.px-conv{background:#fff;border-radius:16px;width:1000px;max-width:100%;max-height:90vh;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 24px 70px rgba(0,0,0,.32);}
.px-conv-h{display:flex;justify-content:space-between;align-items:center;padding:16px 20px;border-bottom:1px solid ${C.line};}
.px-conv-ht{display:flex;align-items:center;gap:12px;}
.px-conv-hic{width:36px;height:36px;border-radius:10px;background:${C.primary};color:#fff;display:grid;place-items:center;}
.px-conv-ht b{font-size:16px;color:${C.navy};display:block;}
.px-conv-ht span{font-size:11.5px;color:${C.sub};}
.px-conv-h>button{border:none;background:${C.bg};border-radius:8px;width:32px;height:32px;display:grid;place-items:center;cursor:pointer;color:${C.sub};}
.px-conv-body{display:grid;grid-template-columns:340px 1fr;gap:18px;padding:18px 20px;overflow:auto;}
.px-conv-step{font-size:11px;font-weight:800;color:${C.primary};text-transform:uppercase;letter-spacing:.05em;margin-bottom:9px;}
.px-conv-left textarea{width:100%;height:120px;border:1.5px solid ${C.line};border-radius:10px;padding:11px;font-size:12.5px;outline:none;resize:none;font-family:inherit;line-height:1.5;margin-top:9px;}
.px-conv-left textarea:focus{border-color:${C.primary};}
.px-conv-drop{display:flex;align-items:center;gap:11px;border:1.5px dashed ${C.line};border-radius:11px;padding:13px 14px;cursor:pointer;color:${C.primary};transition:.14s;}
.px-conv-drop:hover{border-color:${C.primary};background:${C.primarySoft};}
.px-conv-drop b{display:block;font-size:12.5px;color:${C.ink};font-weight:700;}
.px-conv-drop span{display:block;font-size:10.5px;color:${C.faint};font-weight:500;margin-top:1px;}
.px-conv-anexo{display:inline-flex;align-items:center;gap:6px;font-size:11.5px;font-weight:600;color:${C.primaryDark};background:${C.primarySoft};border:1px solid ${C.line};border-radius:20px;padding:4px 6px 4px 10px;margin-top:9px;}
.px-conv-anexo.erro{color:#B42318;background:#FEF3F2;border-color:#FECDCA;}
.px-conv-anexo button{background:none;border:none;cursor:pointer;color:${C.sub};display:inline-flex;padding:0;}
.px-conv-vazio{font-size:12px;color:${C.faint};line-height:1.5;border:1px dashed ${C.line};border-radius:10px;padding:18px;text-align:center;}
.px-conv-exemplo{display:flex;align-items:flex-start;gap:7px;font-size:11.5px;line-height:1.45;color:#7a4d00;background:#FBEEDB;border:1px solid #E5C285;border-radius:9px;padding:8px 11px;margin-bottom:10px;}
.px-conv-exlabel{display:inline-flex;align-items:center;gap:5px;font-size:10.5px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;color:${C.faint};margin-bottom:10px;}
.px-conv-exemplo svg{flex-shrink:0;margin-top:1px;}
.px-conv-row.demo{opacity:.6;filter:grayscale(.3);}
.px-conv-lbtns{display:flex;gap:8px;margin-top:10px;}
.px-conv-ghost{flex:1;display:flex;align-items:center;justify-content:center;gap:6px;border:1.5px solid ${C.line};background:#fff;color:${C.ink};border-radius:9px;padding:9px;font-size:12px;font-weight:600;cursor:pointer;}
.px-conv-ai{flex:1;display:flex;align-items:center;justify-content:center;gap:6px;border:none;background:${C.primary};color:#fff;border-radius:9px;padding:9px;font-size:12px;font-weight:700;cursor:pointer;}
.px-conv-ai:disabled{opacity:.6;cursor:default;}
.px-conv-aiout{margin-top:11px;background:${C.bg};border:1px solid ${C.line};border-radius:10px;padding:11px;font-size:11.5px;color:${C.sub};white-space:pre-wrap;line-height:1.5;max-height:150px;overflow:auto;}
.px-conv-table{display:flex;flex-direction:column;gap:7px;}
.px-conv-row{display:flex;align-items:center;gap:10px;border:1px solid ${C.line};border-radius:10px;padding:10px 12px;background:#fff;}
.px-conv-row.on{border-color:#bfe2c5;background:#fbfdfb;}
.px-conv-pgd{font-size:12px;color:${C.ink};flex:1;min-width:0;line-height:1.3;}
.px-conv-match{font-size:11.5px;color:${C.sub};flex:1.1;min-width:0;line-height:1.3;}
.px-conv-match code{font-family:'JetBrains Mono',monospace;font-size:10.5px;font-weight:700;color:${C.primaryDark};}
.px-conv-nomatch{color:#B86E00;font-style:italic;}
.px-conv-conf{font-size:9.5px;font-weight:700;padding:3px 9px;border-radius:20px;text-transform:uppercase;letter-spacing:.03em;flex-shrink:0;}
.px-conv-foot{display:flex;justify-content:space-between;align-items:center;padding:14px 20px;border-top:1px solid ${C.line};background:${C.bg};gap:12px;}
.px-conv-stat{font-size:12.5px;color:${C.sub};}
.px-conv-footbtns{display:flex;gap:9px;}
.px-conv-dl{display:flex;align-items:center;gap:7px;border:1.5px solid ${C.line};background:#fff;color:${C.primaryDark};border-radius:10px;padding:10px 15px;font-size:12.5px;font-weight:700;cursor:pointer;font-family:inherit;}
.px-conv-dl:hover:not(:disabled){border-color:${C.primary};background:${C.primarySoft};}
.px-conv-dl:disabled{opacity:.4;cursor:default;}
.px-conv-inject{display:flex;align-items:center;gap:8px;border:none;background:${C.primary};color:#fff;border-radius:10px;padding:11px 18px;font-size:13px;font-weight:700;cursor:pointer;font-family:inherit;}
.px-conv-inject:hover:not(:disabled){background:${C.primaryDark};}
.px-conv-inject:disabled{opacity:.4;cursor:default;}
/* cards de escolha (rodadas) */
.px-conv-cards{display:flex;flex-direction:column;gap:11px;}
.px-imp-card{border:1px solid ${C.line};border-radius:12px;padding:12px;background:#fff;}
.px-imp-card.add{border-color:#bfe2c5;background:#fbfdfb;}
.px-imp-orig{display:flex;align-items:flex-start;gap:8px;font-size:13px;font-weight:700;color:${C.navy};margin-bottom:10px;line-height:1.35;}
.px-imp-num{flex-shrink:0;width:20px;height:20px;border-radius:50%;background:${C.navy};color:#fff;font-size:11px;font-weight:800;display:grid;place-items:center;}
.px-imp-ops{display:flex;flex-direction:column;gap:5px;}
.px-imp-op{display:flex;align-items:flex-start;gap:10px;width:100%;text-align:left;border:1.5px solid ${C.line};background:#fff;border-radius:11px;padding:11px 13px;cursor:pointer;font-family:inherit;transition:.14s;}
.px-imp-op:hover{border-color:var(--opcor,${C.primary});}
.px-imp-op.on{border-color:var(--opcor,${C.primary});background:var(--opsoft,${C.primarySoft});box-shadow:0 0 0 1px var(--opcor,${C.primary}) inset;}
.px-imp-radio{flex-shrink:0;width:22px;height:22px;border-radius:50%;border:2px solid ${C.line};display:grid;place-items:center;margin-top:0;color:#fff;transition:.14s;}
.px-imp-radio.on{background:var(--opcor,${C.primary});border-color:var(--opcor,${C.primary});}
.px-imp-op-body{flex:1;display:flex;flex-direction:column;gap:2px;}
.px-imp-op-top{display:flex;align-items:center;gap:7px;flex-wrap:wrap;}
.px-imp-op-top code{font-family:'JetBrains Mono',monospace;font-size:10.5px;font-weight:700;color:${C.primaryDark};}
.px-imp-op-top em{font-style:normal;font-size:9.5px;font-weight:700;padding:1px 7px;border-radius:20px;}
.px-imp-op-top i{font-style:normal;font-size:10px;color:${C.faint};}
.px-imp-op-name{font-size:12px;color:${C.ink};line-height:1.3;}
.px-imp-semop{display:flex;align-items:center;gap:6px;font-size:11.5px;color:#B86E00;background:#FBEEDB;border:1px solid #E5C285;border-radius:8px;padding:8px 10px;}
.px-imp-actions{margin-top:9px;display:flex;justify-content:flex-end;}
.px-imp-add{display:flex;align-items:center;gap:5px;border:1px solid ${C.primary};background:#fff;color:${C.primaryDark};border-radius:8px;padding:6px 12px;font-size:11.5px;font-weight:700;cursor:pointer;font-family:inherit;}
.px-imp-add:hover:not(:disabled){background:${C.primarySoft};}
.px-imp-add:disabled{opacity:.4;cursor:default;border-color:${C.line};color:${C.faint};}
.px-imp-added{display:flex;align-items:center;gap:5px;font-size:11.5px;font-weight:700;color:${C.green};}
/* wizard passo-a-passo */
.px-wiz{display:flex;flex-direction:column;gap:14px;}
.px-wiz-prog{background:${C.surface};border:1px solid ${C.line};border-radius:11px;padding:12px 14px;}
.px-wiz-prog-top{display:flex;justify-content:space-between;font-size:11.5px;color:${C.sub};margin-bottom:8px;}
.px-wiz-prog-top b{color:${C.navy};font-weight:800;}
.px-wiz-bar{height:7px;background:${C.line};border-radius:20px;overflow:hidden;}
.px-wiz-bar i{display:block;height:100%;background:linear-gradient(90deg,${C.primary},${C.green});border-radius:20px;transition:width .4s ease;}
.px-wiz-dots{display:flex;flex-wrap:wrap;gap:5px;margin-top:10px;}
.px-wiz-dot{width:18px;height:6px;border-radius:20px;border:none;background:${C.line};cursor:pointer;padding:0;transition:.15s;}
.px-wiz-dot.cur{background:${C.primary};width:26px;}
.px-wiz-dot.ok{background:${C.green};}
.px-wiz-card{border:1px solid ${C.line};border-radius:13px;padding:15px;background:#fff;}
.px-wiz-card.add{border-color:#bfe2c5;background:#fbfdfb;}
.px-wiz-orig{display:flex;align-items:flex-start;gap:9px;font-size:14px;font-weight:700;color:${C.navy};margin-bottom:13px;line-height:1.35;}
.px-wiz-ops{display:flex;flex-direction:column;gap:6px;}
.px-imp-op.nova{border-style:dashed;border-color:#b9cdf0;}
.px-imp-op.nova.on{border-style:solid;background:${C.primarySoft};}
.px-imp-op.nova .px-imp-op-name{color:${C.sub};}
.px-wiz-novaviso{display:flex;align-items:flex-start;gap:7px;margin-top:11px;font-size:11.5px;line-height:1.45;color:${C.primaryDark};background:${C.primarySoft};border:1px solid #b9cdf0;border-radius:9px;padding:9px 11px;}
.px-wiz-novaviso svg{flex-shrink:0;margin-top:1px;}
.px-wiz-nav{display:flex;align-items:center;justify-content:space-between;gap:10px;}
.px-wiz-prev,.px-wiz-next{display:flex;align-items:center;gap:5px;border:1px solid ${C.line};background:#fff;color:${C.sub};border-radius:9px;padding:9px 13px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit;}
.px-wiz-prev:hover:not(:disabled),.px-wiz-next:hover:not(:disabled){border-color:${C.primary};color:${C.primaryDark};}
.px-wiz-prev:disabled,.px-wiz-next:disabled{opacity:.35;cursor:default;}
.px-wiz-add{flex:1;display:flex;align-items:center;justify-content:center;gap:7px;border:none;background:${C.primary};color:#fff;border-radius:9px;padding:11px 16px;font-size:13px;font-weight:700;cursor:pointer;font-family:inherit;}
.px-wiz-add:hover:not(:disabled){background:${C.primaryDark};}
.px-wiz-add:disabled{opacity:.4;cursor:default;}
/* destaque de entrega nova na descrição de área */
.px-doc-item.nova .px-doc-row code{color:${C.primary};background:${C.primarySoft};padding:1px 6px;border-radius:5px;font-style:normal;}
.px-doc-novatag{display:inline-flex;align-items:center;gap:3px;font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.03em;color:${C.primaryDark};background:${C.primarySoft};border:1px solid #b9cdf0;border-radius:20px;padding:1px 7px;margin-left:6px;}

/* flag / qualidade modais */
.px-flagm,.px-qual{background:#fff;border-radius:16px;width:540px;max-width:100%;max-height:88vh;overflow:auto;padding:22px;box-shadow:0 22px 60px rgba(0,0,0,.3);}
.px-flagm-h{display:flex;justify-content:space-between;align-items:center;font-size:16px;color:${C.navy};}
.px-flagm-h div{display:flex;align-items:center;gap:8px;}
.px-flagm-h button{border:none;background:${C.bg};border-radius:8px;width:32px;height:32px;display:grid;place-items:center;cursor:pointer;color:${C.sub};}
.px-flagm-sub{font-size:12.5px;color:${C.sub};margin:10px 0 16px;line-height:1.5;}
.px-flagm-sub code{font-family:'JetBrains Mono',monospace;font-weight:700;color:${C.primaryDark};}
.px-flagm-l{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;color:${C.faint};margin-bottom:8px;}
.px-flagm-tipos{display:flex;flex-wrap:wrap;gap:7px;margin-bottom:16px;}
.px-flagm-tipo{font-size:12px;font-weight:600;padding:7px 12px;border-radius:20px;border:1.5px solid ${C.line};background:#fff;color:${C.sub};cursor:pointer;}
.px-flagm-tipo.on{background:#B86E00;border-color:#B86E00;color:#fff;}
.px-flagm textarea{width:100%;height:80px;border:1.5px solid ${C.line};border-radius:10px;padding:11px;font-size:13px;outline:none;resize:none;font-family:inherit;}
.px-flagm textarea:focus{border-color:${C.primary};}
.px-flagm-acts{display:flex;justify-content:flex-end;gap:10px;margin-top:16px;}
.px-qual-empty{font-size:13px;color:${C.faint};text-align:center;padding:28px;border:1px dashed ${C.line};border-radius:11px;line-height:1.5;}
.px-qual-list{display:flex;flex-direction:column;gap:9px;}
.px-qual-item{display:flex;gap:11px;align-items:flex-start;border:1px solid ${C.line};border-radius:11px;padding:12px;}
.px-qual-body{flex:1;}
.px-qual-row{font-size:12.5px;color:${C.ink};}
.px-qual-row code{font-family:'JetBrains Mono',monospace;font-size:11px;font-weight:700;color:${C.primaryDark};}
.px-qual-item p{font-size:12px;color:${C.sub};margin:5px 0 0;line-height:1.45;}
.px-qual-embutido{width:100%;max-width:820px;margin:0 auto;box-shadow:none;border:1px solid ${C.line};max-height:none;}
.px-desc-expandida{max-width:820px;margin:0 auto;background:#fff;border:1px solid ${C.line};border-radius:14px;padding:16px 18px;}
.px-desc-expandida .px-doc-h{margin-bottom:12px;}

/* header ghost btn */
.px-ghostbtn{display:flex;align-items:center;gap:6px;font-size:12px;font-weight:600;color:${C.sub};background:#fff;border:1px solid ${C.line};border-radius:20px;padding:5px 12px;cursor:pointer;}
.px-ghostbtn:hover{border-color:${C.primary};color:${C.primary};}

/* nav toggle + facet selects */
.px-navtoggle{flex-shrink:0;width:42px;height:42px;border-radius:11px;border:1.5px solid ${C.line};background:#fff;color:${C.sub};display:grid;place-items:center;cursor:pointer;}
.px-navtoggle:hover{border-color:${C.primary};color:${C.primary};}
.px-navtoggle.on{background:${C.primary};border-color:${C.primary};color:#fff;}


/* descrição: grupos por natureza + observação */
.px-doc-group{margin-bottom:10px;}
.px-doc-gh{display:flex;align-items:center;gap:7px;font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.04em;color:${C.sub};margin:6px 0 6px;}
.px-doc-gc{margin-left:auto;font-size:10px;background:${C.bg};border:1px solid ${C.line};border-radius:10px;padding:1px 8px;color:${C.faint};}
.px-doc-del{border:none;background:transparent;color:${C.faint};cursor:pointer;align-self:flex-start;padding:2px;}
.px-doc-del:hover{color:#c0392b;}
.px-doc-addnote{margin-top:5px;display:inline-flex;align-items:center;gap:4px;border:1px dashed ${C.line};background:transparent;color:${C.faint};font-size:10.5px;font-weight:600;padding:4px 8px;border-radius:7px;cursor:pointer;}
.px-doc-addnote:hover{border-color:${C.primary};color:${C.primary};}
.px-doc-notebox{margin-top:5px;}
.px-doc-notebox textarea{width:100%;min-height:46px;border:1.5px solid ${C.primarySoft};background:#fbfcfe;border-radius:8px;padding:7px;font-size:11.5px;font-family:inherit;resize:vertical;outline:none;color:${C.ink};}
.px-doc-notebox textarea:focus{border-color:${C.primary};}

/* micro-confirmação inline + comparar */
.px-added{display:inline-flex;align-items:center;gap:3px;font-size:10px;font-weight:700;color:${C.green};background:${C.greenSoft};border-radius:20px;padding:2px 8px;flex-shrink:0;animation:pxpop .2s ease;}
@keyframes pxpop{from{transform:scale(.7);opacity:0;}to{transform:scale(1);opacity:1;}}
.px-cmpbtn{border:none;background:transparent;color:${C.faint};cursor:pointer;flex-shrink:0;padding:3px;display:grid;place-items:center;border-radius:6px;}
.px-cmpbtn:hover{color:${C.primary};background:${C.primarySoft};}
.px-cmpbtn.on{color:#fff;background:${C.primary};}
.px-edet-acts{display:flex;gap:9px;flex-wrap:wrap;margin-top:11px;}
.px-cmplink{display:inline-flex;align-items:center;gap:6px;border:1px solid ${C.line};background:#fff;color:${C.sub};font-size:11px;font-weight:700;padding:6px 11px;border-radius:8px;cursor:pointer;}
.px-cmplink.on,.px-cmplink:hover{border-color:${C.primary};color:${C.primary};}

/* estado vazio melhorado */
.px-empty2{background:#fff;border:1px solid ${C.line};border-radius:14px;padding:42px 28px;text-align:center;}
.px-empty2-ic{width:56px;height:56px;border-radius:50%;background:${C.bg};color:${C.faint};display:grid;place-items:center;margin:0 auto 14px;}
.px-empty2 h3{font-size:16px;color:${C.navy};margin:0 0 7px;font-weight:700;}
.px-empty2 h3 b{color:${C.primary};}
.px-empty2 p{font-size:13px;color:${C.sub};max-width:460px;margin:0 auto 16px;line-height:1.55;}
.px-empty2-sug{display:flex;align-items:center;justify-content:center;gap:8px;flex-wrap:wrap;margin-bottom:18px;}
.px-empty2-sug>span:first-child{font-size:11.5px;color:${C.faint};}
.px-empty2-tag{font-size:11.5px;color:${C.sub};background:${C.bg};border:1px solid ${C.line};border-radius:20px;padding:4px 11px;}
.px-empty2-btn{display:inline-flex;align-items:center;gap:7px;background:${C.primary};color:#fff;border:none;border-radius:10px;padding:11px 18px;font-size:13px;font-weight:700;cursor:pointer;}
.px-empty2-btn:hover{background:${C.primaryDark};}

/* skeleton */
.px-skel{display:flex;flex-direction:column;gap:8px;}
.px-bancoload{display:flex;flex-direction:column;align-items:center;}
.px-bancoload-sp{width:30px;height:30px;border:3px solid ${C.line};border-top-color:${C.primary};border-radius:50%;animation:pxspin .8s linear infinite;margin:18px 0 12px;}
.px-bancoload-t{font-size:14px;font-weight:700;color:${C.navy};}
.px-bancoload-s{font-size:12px;color:${C.faint};margin:3px 0 20px;}
@keyframes pxspin{to{transform:rotate(360deg);}}
.px-banco-aviso{display:flex;align-items:flex-start;gap:8px;background:#FBEEDB;border:1px solid #E5C285;color:#7a4d00;font-size:12.5px;line-height:1.45;border-radius:10px;padding:10px 13px;margin-bottom:14px;}
.px-banco-aviso svg{flex-shrink:0;margin-top:1px;}
.px-skel-cat{border:1px solid ${C.line};border-radius:11px;overflow:hidden;background:#fff;}
.px-skel-head{display:flex;align-items:center;gap:10px;background:${C.bg};padding:13px 14px;}
.px-skel-row{display:flex;align-items:center;gap:11px;padding:10px 14px;border-top:1px solid #f1f3f6;}
.px-skel-bar,.px-skel-pill,.px-skel-box{background:linear-gradient(90deg,#eef1f5 25%,#e4e8ee 37%,#eef1f5 63%);background-size:400% 100%;animation:pxsh 1.3s ease infinite;border-radius:5px;height:12px;}
.px-skel-bar.w20{width:60px;} .px-skel-bar.w40{width:40%;} .px-skel-bar.w60{width:55%;}
.px-skel-pill{width:80px;height:18px;border-radius:20px;margin-left:auto;}
.px-skel-box{width:18px;height:18px;border-radius:5px;}
@keyframes pxsh{0%{background-position:100% 0;}100%{background-position:-100% 0;}}

/* onboarding */
.px-ob-bg{position:fixed;inset:0;background:rgba(12,24,44,.55);backdrop-filter:blur(3px);display:grid;place-items:center;z-index:60;padding:20px;}
.px-ob{background:#fff;border-radius:20px;max-width:880px;width:100%;padding:34px 32px 24px;position:relative;box-shadow:0 30px 80px rgba(0,0,0,.35);text-align:center;}
.px-ob-x{position:absolute;top:16px;right:16px;border:none;background:${C.bg};border-radius:8px;width:32px;height:32px;display:grid;place-items:center;cursor:pointer;color:${C.sub};}
.px-ob-logo{margin-bottom:10px;}
.px-ob h2{font-size:23px;font-weight:800;color:${C.navy};margin:0 0 6px;}
.px-ob>p{font-size:13.5px;color:${C.sub};margin:0 0 22px;}
.px-ob-cards{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;text-align:left;}
.px-ob-card{display:flex;flex-direction:column;gap:7px;border:1.5px solid ${C.line};border-radius:14px;padding:18px 16px;background:#fff;cursor:pointer;transition:.15s;}
.px-ob-card:hover{border-color:${C.primary};box-shadow:0 8px 24px rgba(19,81,180,.13);transform:translateY(-2px);}
.px-ob-ic{width:44px;height:44px;border-radius:12px;display:grid;place-items:center;margin-bottom:4px;}
.px-ob-card b{font-size:14.5px;color:${C.navy};}
.px-ob-card>span:not(.px-ob-ic):not(.px-ob-go){font-size:12px;color:${C.sub};line-height:1.5;flex:1;}
.px-ob-go{display:flex;align-items:center;gap:5px;font-size:12px;font-weight:700;color:${C.primary};margin-top:4px;}
.px-ob-dont{display:inline-flex;align-items:center;gap:8px;font-size:12px;color:${C.faint};margin-top:20px;cursor:pointer;}
.px-ob-dont input{width:15px;height:15px;cursor:pointer;}

/* menu lateral macroprocesso */
.px-body.nav{grid-template-columns:288px 1fr;}
.px-mnav{background:#fff;border:1px solid ${C.line};border-radius:14px;position:sticky;top:14px;max-height:calc(100vh - 30px);display:flex;flex-direction:column;overflow:hidden;}
.px-mnav-h{display:flex;align-items:center;justify-content:space-between;padding:13px 14px;border-bottom:1px solid ${C.line};}
.px-mnav-h>span{display:flex;align-items:center;gap:7px;font-size:13px;font-weight:800;color:${C.navy};}
.px-mnav-h button{border:none;background:${C.bg};border-radius:7px;width:28px;height:28px;display:grid;place-items:center;cursor:pointer;color:${C.sub};}
.px-mnav-all{margin:10px 12px 4px;text-align:left;border:1px solid ${C.line};background:#fff;border-radius:9px;padding:9px 11px;font-size:12.5px;font-weight:700;color:${C.sub};cursor:pointer;}
.px-mnav-all.on{border-color:${C.primary};color:${C.primary};background:${C.primarySoft};}
.px-mnav-search{display:flex;align-items:center;gap:8px;margin:0 12px 8px;border:1.5px solid ${C.line};border-radius:9px;padding:7px 10px;background:#fff;}
.px-mnav-search input{flex:1;border:none;outline:none;font-size:12.5px;background:transparent;}
.px-mnav-empty{font-size:12px;color:${C.faint};text-align:center;padding:16px;}
.px-mnav-body{overflow-y:auto;padding:6px 8px 10px;}
.px-mnav-grp{margin-bottom:2px;}
.px-mnav-nat{display:flex;align-items:center;gap:7px;width:100%;text-align:left;border:none;background:transparent;padding:8px 8px;cursor:pointer;font-size:12.5px;font-weight:800;color:${C.navy};}
.px-mnav-nat .px-dot{width:9px;height:9px;border-radius:50%;}
.px-mnav-c{margin-left:auto;font-size:10px;color:${C.faint};background:${C.bg};border-radius:10px;padding:1px 7px;font-weight:700;}
.px-mnav-list{display:flex;flex-direction:column;gap:1px;padding:0 2px 6px 6px;}
.px-mnav-item{display:flex;align-items:center;gap:8px;text-align:left;border:none;background:transparent;border-left:2px solid ${C.line};padding:7px 9px;cursor:pointer;border-radius:0 7px 7px 0;}
.px-mnav-item span{flex:1;font-size:11.5px;color:${C.sub};line-height:1.3;overflow:hidden;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;}
.px-mnav-item b{font-size:10.5px;color:${C.faint};flex-shrink:0;}
.px-mnav-item:hover{background:${C.bg};}
.px-mnav-item.on{background:${C.primarySoft};border-left-color:${C.primary};}
.px-mnav-item.on span{color:${C.primary};font-weight:700;}

/* botões flutuantes + drawer */
.px-descfab,.px-cmpfab{position:fixed;display:flex;align-items:center;gap:9px;border:none;border-radius:30px;padding:13px 20px;font-size:14px;font-weight:700;cursor:pointer;color:#fff;box-shadow:0 8px 24px rgba(0,0,0,.22);z-index:35;}
.px-descfab{right:22px;bottom:22px;background:${C.primary};}
.px-cmpfab{right:22px;bottom:80px;background:${C.navy};}
.px-descfab span,.px-cmpfab span{background:rgba(255,255,255,.25);border-radius:20px;min-width:22px;height:22px;display:grid;place-items:center;font-size:12px;}
.px-drawer-bg{position:fixed;inset:0;background:rgba(12,24,44,.45);z-index:50;display:flex;justify-content:flex-end;}
.px-drawer{width:380px;max-width:90vw;background:#fff;height:100%;padding:18px;overflow-y:auto;box-shadow:-12px 0 40px rgba(0,0,0,.2);animation:pxslide .22s ease;}
@keyframes pxslide{from{transform:translateX(30px);opacity:.6;}to{transform:translateX(0);opacity:1;}}
.px-x{border:none;background:${C.bg};border-radius:8px;width:32px;height:32px;display:grid;place-items:center;cursor:pointer;color:${C.sub};}

/* comparação */
.px-cmp{background:#fff;border-radius:16px;width:980px;max-width:100%;max-height:90vh;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 24px 70px rgba(0,0,0,.32);}
.px-cmp-h{display:flex;justify-content:space-between;align-items:flex-start;padding:16px 20px;border-bottom:1px solid ${C.line};}
.px-cmp-h b{font-size:16px;color:${C.navy};margin-left:7px;}
.px-cmp-h span{display:block;font-size:11.5px;color:${C.sub};margin-top:3px;}
.px-cmp-grid{display:grid;overflow:auto;padding:4px 0;}
.px-cmp-cell{padding:10px 14px;font-size:12px;color:${C.ink};border-bottom:1px solid #f1f3f6;line-height:1.45;}
.px-cmp-rowh{font-size:9.5px;font-weight:800;text-transform:uppercase;letter-spacing:.04em;color:${C.faint};background:${C.bg};position:sticky;left:0;}
.px-cmp-colh{border-bottom:2px solid ${C.line};}
.px-cmp-top{display:flex;align-items:center;justify-content:space-between;}
.px-cmp-top code{font-family:'JetBrains Mono',monospace;font-size:12px;font-weight:700;}
.px-cmp-drop{border:none;background:transparent;color:${C.faint};cursor:pointer;padding:2px;}
.px-cmp-drop:hover{color:#c0392b;}
.px-cmp-add{margin-top:7px;width:100%;display:flex;align-items:center;justify-content:center;gap:5px;border:1.5px solid ${C.primary};background:${C.primary};color:#fff;border-radius:8px;padding:7px;font-size:11.5px;font-weight:700;cursor:pointer;}
.px-cmp-add.done{background:${C.greenSoft};border-color:#bfe2c5;color:${C.green};cursor:default;}
.px-cmp-foot{display:flex;justify-content:space-between;padding:14px 20px;border-top:1px solid ${C.line};background:${C.bg};}
.px-mt-obs{font-size:11.5px;color:${C.sub};}
.px-mt-obs i{color:${C.faint};}
.px-flagm-in{width:100%;border:1.5px solid ${C.line};border-radius:9px;padding:10px;font-size:13px;font-family:inherit;outline:none;margin-bottom:4px;}
.px-flagm-in:focus{border-color:${C.primary};}

/* ----- cadeia de valor ----- */
.px-cv{display:flex;flex-direction:column;gap:16px;}
.px-cv-empty{background:${C.surface};border:1px solid ${C.line};border-radius:14px;padding:46px 28px;text-align:center;display:flex;flex-direction:column;align-items:center;gap:10px;}
.px-cv-empty h3{margin:4px 0 0;font-size:16px;color:${C.ink};}
.px-cv-empty p{margin:0;max-width:440px;font-size:13px;line-height:1.55;color:${C.sub};}
.px-cv-obj{background:${C.surface};border:1px solid ${C.line};border-radius:14px;padding:0;overflow:hidden;}
.px-cv-obj-head{display:flex;align-items:center;gap:7px;font-size:14px;color:${C.ink};width:100%;background:none;border:none;cursor:pointer;padding:15px 18px;text-align:left;font-family:inherit;}
.px-cv-obj.open .px-cv-obj-head{border-bottom:1px solid ${C.line};}
.px-cv-obj-head:hover{background:${C.bg};}
.px-cv-obj-head b{font-weight:800;}
.px-cv-obj-head span{font-size:11.5px;font-weight:500;color:${C.faint};margin-left:2px;}
.px-cv-obj-head>svg:last-child{margin-left:auto;color:${C.sub};flex-shrink:0;}
.px-cv-obj-count{font-size:11px;font-weight:800;color:${C.primary};background:${C.primarySoft};border-radius:20px;padding:1px 9px;font-style:normal;margin-left:auto;}
.px-cv-obj-count + svg{margin-left:8px;}
.px-cv-obj-body{padding:16px 18px;}
.px-cv-org{display:flex;align-items:center;flex-wrap:wrap;gap:7px;margin-top:12px;}
.px-cv-org-l{font-size:9.5px;font-weight:800;text-transform:uppercase;letter-spacing:.05em;color:${C.faint};min-width:54px;}
.px-cv-orgpill{font-size:11.5px;color:${C.navy};background:${C.bg};border:1px solid ${C.line};border-radius:999px;padding:4px 11px;transition:.15s;}
.px-cv-orgpill b{color:${C.primary};margin-right:3px;}
.px-cv-orgpill.on{background:${C.primarySoft};border-color:#b9cdf0;color:${C.primaryDark};font-weight:600;}
.px-cv-orgpill.editable{display:inline-flex;align-items:center;gap:5px;padding:3px 6px 3px 11px;}
.px-cv-orgpill.editable input{border:none;background:transparent;font-family:inherit;font-size:11.5px;color:${C.navy};outline:none;min-width:140px;width:160px;}
.px-cv-orgpill.editable input:focus{border-bottom:1px solid ${C.primary};}
.px-cv-orgdel{display:inline-flex;align-items:center;justify-content:center;width:17px;height:17px;border:none;background:${C.line};color:${C.sub};border-radius:50%;cursor:pointer;padding:0;flex-shrink:0;}
.px-cv-orgdel:hover{background:#e2b6b6;color:#8c2f2f;}
.px-cv-orgadd{display:inline-flex;align-items:center;gap:4px;font-size:11.5px;font-weight:700;color:${C.primary};background:#fff;border:1px dashed #b9cdf0;border-radius:999px;padding:5px 12px;cursor:pointer;font-family:inherit;}
.px-cv-orgadd:hover{background:${C.primarySoft};border-style:solid;}
.px-cv-uni{display:flex;align-items:flex-start;gap:8px;margin-top:14px;}
.px-cv-uni-cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(248px,1fr));gap:10px;flex:1;}
.px-cv-ocard{border:1.5px solid ${C.line};border-radius:11px;padding:10px 11px;background:${C.bg};display:flex;flex-direction:column;gap:8px;transition:.15s;}
.px-cv-ocard.on{border-color:${C.primaryDark};background:#fff;box-shadow:0 0 0 3px ${C.primarySoft};}
.px-cv-ocard-top{display:flex;align-items:flex-start;gap:6px;}
.px-cv-otitle{flex:1;border:none;background:transparent;font-size:13px;font-weight:700;color:${C.ink};font-family:inherit;outline:none;padding:2px 0;border-bottom:1px dashed transparent;}
.px-cv-otitle:focus{border-bottom-color:${C.primary};}
.px-cv-odel{border:none;background:transparent;color:${C.faint};cursor:pointer;padding:2px;}
.px-cv-odel:hover{color:#c0392b;}
.px-cv-ocard-link{display:flex;flex-direction:column;gap:3px;}
.px-cv-ocard-link span{font-size:9.5px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;color:${C.faint};}
.px-cv-ocard-link select{border:1px solid ${C.line};border-radius:7px;padding:5px 7px;font-size:11.5px;color:${C.ink};background:#fff;font-family:inherit;outline:none;cursor:pointer;}
.px-cv-ocard-foot{display:flex;align-items:center;justify-content:space-between;gap:6px;}
.px-cv-ocount{font-size:11px;font-weight:700;color:${C.sub};}
.px-cv-ofocus{border:1.5px solid ${C.primary};background:#fff;color:${C.primary};border-radius:7px;padding:4px 9px;font-size:11px;font-weight:700;cursor:pointer;transition:.15s;}
.px-cv-ofocus.on{background:${C.primary};color:#fff;}
.px-cv-oadd{display:flex;align-items:center;justify-content:center;gap:5px;border:1.5px dashed ${C.line};background:transparent;color:${C.sub};border-radius:11px;padding:10px;font-size:12px;font-weight:700;cursor:pointer;min-height:96px;transition:.15s;}
.px-cv-oadd:hover{border-color:${C.primary};color:${C.primary};}
.px-cv-hint{display:flex;align-items:center;gap:7px;margin-top:12px;padding:8px 12px;background:${C.primarySoft};border:1px solid #c5d7f3;border-radius:9px;font-size:12px;color:${C.primaryDark};}
.px-cv-hint b{font-weight:800;}
.px-cv-hint button{margin-left:auto;border:none;background:${C.primary};color:#fff;border-radius:6px;padding:4px 10px;font-size:11px;font-weight:700;cursor:pointer;}
.px-cv-chain{display:flex;flex-direction:column;gap:12px;}
.px-cv-lane{border-radius:13px;padding:12px 14px;}
.px-cv-lane-l{display:flex;align-items:center;gap:7px;font-size:12px;font-weight:800;text-transform:uppercase;letter-spacing:.04em;margin-bottom:10px;}
.px-cv-lane-l i{font-style:normal;font-size:11px;font-weight:800;background:rgba(255,255,255,.7);color:inherit;border-radius:999px;padding:1px 8px;}
.px-cv-dot{width:9px;height:9px;border-radius:50%;}
.px-cv-lane-flow{display:flex;align-items:stretch;gap:4px;overflow-x:auto;padding-bottom:4px;}
.px-cv-macro{flex:0 0 auto;width:230px;background:#fff;border:1.5px solid ${C.line};border-radius:11px;padding:10px;display:flex;flex-direction:column;gap:8px;transition:.18s;}
.px-cv-macro.aceso{box-shadow:0 0 0 3px ${C.primarySoft};}
.px-cv-macro.dim{opacity:.45;}
.px-cv-macro-h{display:flex;align-items:flex-start;justify-content:space-between;gap:6px;font-size:12px;font-weight:700;color:${C.navy};line-height:1.3;min-height:32px;}
.px-cv-macro-h>span{flex:1 1 auto;}
.px-cv-macro-ord{display:flex;gap:2px;flex:0 0 auto;}
.px-cv-macro-ord button{display:flex;align-items:center;justify-content:center;width:20px;height:20px;border:1px solid ${C.line};background:#fff;border-radius:6px;color:${C.sub};cursor:pointer;padding:0;transition:.15s;}
.px-cv-macro-ord button:hover:not(:disabled){border-color:${C.primary};color:${C.primary};}
.px-cv-macro-ord button:disabled{opacity:.3;cursor:default;}
.px-cv-macro-h i{flex:0 0 auto;font-style:normal;font-size:10.5px;font-weight:800;color:#fff;border-radius:999px;padding:1px 7px;}
.px-cv-chips{display:flex;flex-direction:column;gap:5px;}
.px-cv-chip{display:flex;align-items:center;gap:6px;text-align:left;width:100%;border:1px solid ${C.line};background:${C.bg};border-radius:8px;padding:5px 8px;font-family:inherit;cursor:default;transition:.15s;}
.px-cv-chip.clickable{cursor:pointer;}
.px-cv-chip.clickable:hover{border-color:${C.primary};}
.px-cv-chip code{font-family:'JetBrains Mono',monospace;font-size:10px;font-weight:700;color:${C.primary};flex:0 0 auto;}
.px-cv-chip span{font-size:11px;color:${C.sub};line-height:1.3;overflow:hidden;}
.px-cv-chip.linked{background:#fff;border-color:${C.primaryDark};box-shadow:0 0 0 2px ${C.primarySoft};}
.px-cv-chip.linked span{color:${C.ink};font-weight:600;}
.px-cv-chip.dimc{opacity:.4;}
.px-cv-mark{flex:0 0 auto;margin-left:auto;font-style:normal;font-size:9.5px;font-weight:800;color:${C.primary};background:${C.primarySoft};border-radius:999px;min-width:16px;height:16px;display:flex;align-items:center;justify-content:center;padding:0 4px;}
.px-cv-arrow{display:flex;align-items:center;color:${C.faint};flex:0 0 auto;}
.px-cv-cap{display:flex;align-items:flex-start;gap:6px;font-size:11px;line-height:1.5;color:${C.faint};padding:2px 2px 0;}
.px-cv-cap svg{flex:0 0 auto;margin-top:1px;}

@media (max-width:1000px){
  .px-body{grid-template-columns:1fr;} .px-doc{position:static;}
  .px-cards{grid-template-columns:1fr;}
  .px-conv-body{grid-template-columns:1fr;}
  .px-body.nav{grid-template-columns:1fr;}
  .px-ob-cards{grid-template-columns:1fr;}
  .px-mnav{position:static;max-height:none;}
  .px-cv-uni{flex-direction:column;} .px-cv-uni-cards{grid-template-columns:1fr;}
  .px-cv-macro{width:200px;}
}

/* menu Ferramentas + FAB arrastável */
.px-tools{position:relative;}
.px-ghostbtn.on{border-color:${C.primary};color:${C.primary};}
.px-tools-veil{position:fixed;inset:0;z-index:40;}
.px-tools-menu{position:absolute;right:0;top:calc(100% + 6px);z-index:41;background:#fff;border:1px solid ${C.line};border-radius:12px;box-shadow:0 16px 40px rgba(13,49,111,.16);padding:6px;min-width:252px;display:flex;flex-direction:column;gap:2px;}
.px-tools-menu button{display:flex;align-items:center;gap:11px;background:none;border:0;border-radius:9px;padding:9px 11px;cursor:pointer;font-family:inherit;text-align:left;color:${C.ink};}
.px-tools-menu button:hover{background:${C.bg};}
.px-tools-menu button>svg{color:${C.primary};flex-shrink:0;}
.px-tools-menu button b{display:block;font-size:13px;font-weight:700;}
.px-tools-menu button span{display:block;font-size:11.5px;color:${C.faint};}


/* hub: busca em destaque */
.px-ob-search{display:flex;align-items:center;gap:10px;background:#fff;border:2px solid ${C.line};border-radius:12px;padding:4px 6px 4px 14px;margin:0 0 18px;box-shadow:0 2px 10px rgba(13,49,111,.05);}
.px-ob-search input{flex:1;border:0;outline:0;background:none;font-size:15px;font-family:inherit;color:${C.ink};padding:9px 0;}
.px-ob-search button{border:0;background:${C.primary};color:#fff;border-radius:9px;padding:9px 18px;font-size:13.5px;font-weight:700;font-family:inherit;cursor:pointer;}
.px-ob-search button:hover{background:#0C326F;}
/* árvore de decomposição: texto da entrega inteiro */
.px-arv-node.leaf{height:auto;align-items:flex-start;}
.px-arv-node.leaf .px-arv-node-tx{white-space:normal;overflow:visible;display:block;line-height:1.35;}


/* barra de ações (3 acessos) */
.px-actionbar{display:flex;gap:12px;padding:12px 22px 0;max-width:1320px;margin:0 auto;width:100%;box-sizing:border-box;flex-wrap:wrap;}
.px-act{flex:1;min-width:220px;display:flex;align-items:center;gap:12px;background:#fff;border:1px solid ${C.line};border-left:4px solid ${C.primary};border-radius:12px;padding:12px 16px;cursor:pointer;font-family:inherit;text-align:left;transition:box-shadow .15s,transform .05s;box-shadow:0 1px 3px rgba(13,49,111,.05);}
.px-act:hover{box-shadow:0 6px 18px rgba(13,49,111,.10);transform:translateY(-1px);}
.px-act .px-act-ic{width:38px;height:38px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0;}
.px-act-tx{display:flex;flex-direction:column;}
.px-act-tx b{font-size:14px;color:${C.navy};font-weight:800;}
.px-act-tx span{font-size:11.5px;color:${C.faint};}
.px-act.rev{border-left-color:${C.green};} .px-act.rev .px-act-ic{background:#E3F2E5;color:${C.green};}
.px-act.org{border-left-color:#155BCB;} .px-act.org .px-act-ic{background:#E6EFFB;color:#155BCB;}
.px-act.enq{border-left-color:${C.primary};} .px-act.enq .px-act-ic{background:${C.primarySoft};color:${C.primary};}
/* toggle de visualização no painel de descrição */
.px-doc-view{display:flex;gap:3px;margin-left:auto;}
.px-doc-view button{width:28px;height:26px;border:1px solid ${C.line};background:#fff;border-radius:7px;cursor:pointer;display:flex;align-items:center;justify-content:center;color:${C.faint};}
.px-doc-view button.on{background:${C.primarySoft};border-color:${C.primary};color:${C.primary};}
.px-doc-cadeia{padding:6px;position:relative;}

/* ===== cadeia de valor: sempre em tela cheia (pop-up) ===== */
.px-cv{position:relative;}
.px-cv-overlay{position:fixed;inset:0;background:rgba(19,49,92,.5);z-index:75;display:grid;place-items:stretch;padding:24px;animation:cvOverlayIn .18s ease;}
@keyframes cvOverlayIn{from{opacity:0;}to{opacity:1;}}
.px-cv-sheet{background:${C.surface};border-radius:16px;overflow:auto;box-shadow:0 24px 70px rgba(0,0,0,.32);animation:cvSheetIn .22s cubic-bezier(.25,.7,.3,1);}
@keyframes cvSheetIn{from{opacity:0;transform:scale(.985) translateY(8px);}to{opacity:1;transform:none;}}
.px-cv-full{padding:0;}
.px-cv-fullbar{position:sticky;top:0;z-index:5;display:flex;align-items:center;justify-content:space-between;gap:12px;background:#fff;border-bottom:1px solid ${C.line};padding:13px 20px;box-shadow:0 1px 4px rgba(13,49,111,.05);}
.px-cv-fullbar-t{display:flex;align-items:center;gap:8px;font-size:15px;color:${C.navy};}
.px-cv-fullbar-t b{font-weight:800;} .px-cv-fullbar-t span{font-size:12px;font-weight:600;color:${C.faint};background:${C.bg};border-radius:20px;padding:2px 10px;}
.px-cv-fullclose{display:inline-flex;align-items:center;gap:6px;height:34px;padding:0 14px;border:1px solid ${C.line};background:#fff;border-radius:9px;font-size:13px;font-weight:700;font-family:inherit;color:${C.sub};cursor:pointer;transition:border-color .15s,color .15s;}
.px-cv-fullclose:hover{border-color:${C.coral};color:${C.coral};}
/* no modo tela cheia, a cadeia e os objetivos respiram: padding maior e faixas mais largas */
.px-cv-full .px-cv-obj,.px-cv-full .px-cv-chain{margin:0 20px;}
.px-cv-full .px-cv-obj{margin-top:16px;}
.px-cv-full .px-cv-chain{padding-bottom:24px;}
.px-cv-full .px-cv-uni-cards{grid-template-columns:repeat(auto-fill,minmax(300px,1fr))!important;}
.px-cv-full .px-cv-lane-flow{flex-wrap:wrap;}
.px-cv-full .px-cv-macro{max-width:none;}
.px-cv-full .px-cv-chip span{white-space:normal;}
@media(max-width:1000px){.px-cv-overlay{padding:10px;}}
/* nova entrega (modo do catálogo) */
.px-nova{max-width:1000px;margin:0 auto;}
.px-nova-hero{display:flex;gap:14px;align-items:flex-start;background:linear-gradient(120deg,${C.primarySoft},#fff);border:1px solid ${C.line};border-radius:16px;padding:18px 20px;margin-bottom:16px;}
.px-nova-badge{width:44px;height:44px;border-radius:12px;background:${C.primary};color:#fff;display:flex;align-items:center;justify-content:center;flex-shrink:0;}
.px-nova-hero h2{margin:0 0 4px;font-size:18px;color:${C.navy};}
.px-nova-hero p{margin:0;font-size:13px;color:${C.sub};line-height:1.5;}
.px-nova-grid{display:grid;grid-template-columns:1.1fr .9fr;gap:16px;}
.px-nova-form{background:#fff;border:1px solid ${C.line};border-radius:14px;padding:16px;}
.px-nova-l{display:block;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;color:${C.faint};margin:12px 0 5px;}
.px-nova-l:first-child{margin-top:0;}
.px-nova-nat{display:flex;gap:6px;flex-wrap:wrap;}
.px-nova-chip{border:1px solid ${C.line};background:#fff;border-radius:20px;padding:5px 12px;font-size:12px;font-family:inherit;cursor:pointer;color:${C.sub};}
.px-nova-in{width:100%;height:38px;border:1px solid ${C.line};border-radius:9px;padding:0 11px;font-size:13px;font-family:inherit;color:${C.ink};box-sizing:border-box;}
.px-nova-ta{width:100%;min-height:70px;border:1px solid ${C.line};border-radius:9px;padding:9px 11px;font-size:13px;font-family:inherit;color:${C.ink};box-sizing:border-box;resize:vertical;}
.px-nova-in:focus,.px-nova-ta:focus{outline:none;border-color:${C.primary};}
.px-nova-btn{margin-top:14px;width:100%;height:42px;border:0;border-radius:10px;background:${C.primary};color:#fff;font-size:14px;font-weight:700;font-family:inherit;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:8px;}
.px-nova-btn:hover{background:#0C326F;} .px-nova-btn.blocked{background:${C.faint};}
.px-nova-btn:disabled{background:${C.green};}
.px-nova-check{background:#fff;border:1px solid ${C.line};border-radius:14px;padding:16px;align-self:flex-start;}
.px-nova-check-h{display:flex;align-items:center;gap:7px;font-size:13px;font-weight:800;color:${C.navy};margin-bottom:10px;}
.px-nova-empty{font-size:12.5px;color:${C.faint};}
.px-nova-verd{display:flex;align-items:flex-start;gap:8px;font-size:12.5px;border-radius:9px;padding:9px 11px;line-height:1.4;margin-bottom:10px;}
.px-nova-verd.v-ok{background:#E3F2E5;color:#0F6B1E;} .px-nova-verd.v-atencao{background:#FBEEDB;color:#8A5200;} .px-nova-verd.v-bloqueio{background:#FAECE7;color:#9E3B1F;}
.px-nova-cand{display:flex;gap:10px;align-items:center;padding:8px 0;border-top:1px solid ${C.line};}
.px-nova-cand-sim{font-size:12px;font-weight:800;color:${C.primary};min-width:36px;}
.px-nova-cand-tx b{display:block;font-size:12.5px;color:${C.ink};font-weight:600;}
.px-nova-cand-tx span{font-size:11px;color:${C.faint};}
@media(max-width:820px){.px-nova-grid{grid-template-columns:1fr;}}


/* árvore: folha (entrega) com texto inteiro — sem corte */
.px-arv-node.leaf{overflow:visible!important;height:auto!important;align-items:flex-start!important;}
.px-arv-node.leaf .px-arv-node-tx{white-space:normal!important;overflow:visible!important;text-overflow:clip!important;display:block!important;line-height:1.35!important;}

/* ---------- painel de números (topo) — discreto ---------- */
.px-mx{max-width:1320px;margin:10px auto 0;padding:0 22px;display:flex;gap:14px;align-items:center;flex-wrap:wrap;}
.px-mx-lead{flex:1 1 420px;min-width:280px;display:flex;align-items:center;gap:13px;background:linear-gradient(115deg,#FFFFFF 55%,${C.primarySoft});border:1px solid ${C.line};border-left:3px solid ${C.primary};border-radius:11px;padding:11px 16px;box-shadow:0 1px 3px rgba(19,81,180,.06);}
.px-mx-lead-ic{width:40px;height:40px;border-radius:11px;display:grid;place-items:center;background:#fff;border:1px solid #D7E3F7;color:${C.primary};flex-shrink:0;box-shadow:0 1px 3px rgba(19,81,180,.12);}
.px-mx-main{flex:1;min-width:0;}
.px-mx-eyebrow{font-size:9.5px;font-weight:800;text-transform:uppercase;letter-spacing:.07em;color:${C.faint};}
.px-mx-row{display:flex;align-items:baseline;gap:16px;flex-wrap:wrap;margin-top:2px;}
.px-mx-big{font-size:22px;font-weight:800;color:${C.navy};line-height:1;letter-spacing:-.01em;}
.px-mx-big em{font-style:normal;font-size:11.5px;font-weight:600;color:${C.faint};letter-spacing:0;}
.px-mx-mini{font-size:12px;color:${C.sub};font-weight:500;}
.px-mx-mini b{font-weight:800;color:${C.ink};}
.px-mx-btn{border:none;background:transparent;font-family:inherit;cursor:pointer;padding:4px 8px;border-radius:8px;transition:.14s;margin:-4px -8px;}
.px-mx-btn:hover{background:#fff;box-shadow:0 1px 3px rgba(19,81,180,.12);}
.px-mx-btn.on{background:${C.primarySoft};color:${C.primaryDark};}
.px-mx-btn.on b{color:${C.primaryDark};}
.px-mx-bar{display:flex;height:5px;border-radius:4px;overflow:hidden;background:${C.line};margin-top:9px;}
.px-mx-bar span{height:100%;}
.px-mx-acts{display:flex;gap:10px;align-items:stretch;flex-wrap:wrap;}
.px-mx-cov{display:flex;align-items:center;gap:10px;background:${C.surface};border:1px solid ${C.line};border-radius:11px;padding:8px 13px;cursor:pointer;font-family:inherit;transition:.14s;}
.px-mx-cov:hover{border-color:${C.green};background:${C.greenSoft};}
.px-mx-cov-pct{display:flex;align-items:center;gap:5px;font-size:17px;font-weight:800;color:${C.green};white-space:nowrap;}
.px-mx-cov-tx{display:flex;flex-direction:column;line-height:1.25;text-align:left;}
.px-mx-cov-tx>span{font-size:10.5px;color:${C.faint};font-weight:600;}
.px-mx-cov-tx>b{display:flex;align-items:center;gap:4px;font-size:11.5px;font-weight:700;color:${C.green};}
.px-mx-org{display:flex;align-items:center;gap:11px;background:linear-gradient(120deg,#FFFDF5,#FFF8E6);border:1px solid #F0DFA6;border-left:3px solid ${C.yellow};border-radius:11px;padding:8px 15px 8px 13px;cursor:pointer;font-family:inherit;transition:.15s;box-shadow:0 1px 3px rgba(180,140,0,.08);}
.px-mx-org:hover{border-color:#E4B200;box-shadow:0 4px 14px rgba(180,140,0,.18);transform:translateY(-1px);}
.px-mx-org-ic{width:36px;height:36px;border-radius:9px;display:grid;place-items:center;background:${C.yellow};color:${C.navy};flex-shrink:0;}
.px-mx-org-tx{display:flex;flex-direction:column;line-height:1.25;text-align:left;}
.px-mx-org-tx>b{font-size:13.5px;font-weight:800;color:${C.navy};}
.px-mx-org-tx>span{font-size:11px;color:#8a7530;font-weight:600;}
.px-mx-org-arr{color:${C.navy};flex-shrink:0;}
.px-mx-note{flex:1 1 100%;font-size:11.5px;color:${C.sub};background:#FFF8E6;border:1px solid #F5E4A8;border-radius:9px;padding:8px 12px;}
@media(max-width:900px){.px-mx-acts{flex:1 1 100%;}.px-mx-org{flex:1;}}

/* ---------- tooltip ---------- */
.px-tipw{position:relative;display:inline-flex;}
.px-tip{position:absolute;z-index:60;left:50%;top:calc(100% + 9px);transform:translateX(-50%) translateY(4px);background:${C.navy};color:#fff;font-size:11.5px;font-weight:600;line-height:1.4;padding:8px 11px;border-radius:8px;width:max-content;max-width:230px;box-shadow:0 8px 22px rgba(13,49,111,.3);opacity:0;pointer-events:none;transition:opacity .14s,transform .14s;text-align:left;}
.px-tip.wide{max-width:300px;}
.px-tip::after{content:"";position:absolute;bottom:100%;left:50%;transform:translateX(-50%);border:6px solid transparent;border-bottom-color:${C.navy};}
.px-tipw:hover .px-tip{opacity:1;transform:translateX(-50%) translateY(0);}
.px-tip.top{top:auto;bottom:calc(100% + 9px);transform:translateX(-50%) translateY(-4px);}
.px-tip.top::after{bottom:auto;top:100%;border-bottom-color:transparent;border-top-color:${C.navy};}
.px-tipw:hover .px-tip.top{transform:translateX(-50%) translateY(0);}

/* ---------- modos: Nova entrega e Conversor com acento próprio ---------- */
.px-mode.nova{color:${C.green};font-weight:700;}
.px-mode.nova:hover{background:${C.greenSoft};color:${C.green};}
.px-mode.nova.on{background:${C.green};color:#fff;}
.px-mode.conversor{color:${C.primaryDark};background:${C.primarySoft};border:1px solid #C9D8F0;font-weight:700;}
.px-mode.conversor:hover{background:#fff;border-color:${C.primary};box-shadow:0 2px 8px rgba(19,81,180,.15);}
.px-mode.conversor.on{background:${C.primary};color:#fff;border-color:${C.primary};}
/* ---------- Importar PGD (esteira PGD → DFT) — porte do protótipo ---------- */
.cpd{--cpd-paper:#FCFCFD;--cpd-ink:#171B21;--cpd-sub:#4E5866;--cpd-faint:#98A1AD;--cpd-hair:#E7EAEE;--cpd-hair2:#F1F3F6;
  --cpd-navy:#12305B;--cpd-blue:#1351B4;--cpd-blue2:#7B99C9;--cpd-coral:#C2410C;--cpd-coral-s:#FBEDE4;
  --cpd-green:#1B7F3A;--cpd-green-s:#E9F4EC;--cpd-purple:#5B3A9B;--cpd-purple-s:#F0EBF8;--cpd-amber:#9A6A00;--cpd-amber-s:#FBF3DC;
  --cpd-num:"Archivo",-apple-system,sans-serif;--cpd-serif:"Source Serif 4",Georgia,serif;
  font-family:"Archivo",-apple-system,"Segoe UI",sans-serif;background:var(--cpd-paper);color:var(--cpd-ink);
  font-size:13px;line-height:1.55;border:1px solid var(--cpd-hair);border-radius:14px;overflow:hidden;animation:fadeUp .3s ease;}
.cpd b{font-weight:700;}
.cpd mark{background:#FFF3C4;color:inherit;border-radius:3px;padding:0 2px;}
.cpd-govbar{height:3px;background:linear-gradient(90deg,#168821 0 33%,#FFCD07 33% 66%,#1351B4 66% 100%);}
.cpd-hero{background:#fff;border-bottom:1px solid var(--cpd-hair);}
.cpd-hero-in{padding:26px 30px 22px;}
.cpd-hero-grid{display:grid;grid-template-columns:1.15fr .85fr;gap:38px;align-items:end;}
@media(max-width:1100px){.cpd-hero-grid{grid-template-columns:1fr;gap:22px;}}
.cpd-kicker{font-size:9.5px;letter-spacing:.16em;text-transform:uppercase;color:var(--cpd-faint);font-weight:700;}
.cpd-sig{font-family:var(--cpd-serif);font-size:32px;font-weight:700;color:var(--cpd-navy);letter-spacing:-.01em;line-height:1.1;margin-top:8px;}
.cpd-sig .arrow{color:var(--cpd-blue2);font-weight:600;}
.cpd-nome{font-size:12.5px;color:var(--cpd-sub);font-weight:500;margin-top:7px;max-width:520px;}
.cpd-chips{display:flex;gap:10px;align-items:center;margin-top:14px;flex-wrap:wrap;}
.cpd-chip{font-size:10px;font-weight:700;border:1px solid var(--cpd-hair);background:#fff;border-radius:7px;padding:5px 11px;color:var(--cpd-sub);}
.cpd-chip b{color:var(--cpd-navy);}
.cpd-chip.live{border-color:var(--cpd-green);color:var(--cpd-green);}
.cpd-stats{display:flex;align-items:baseline;border-bottom:1px solid var(--cpd-hair);padding-bottom:12px;margin-bottom:12px;}
.cpd-stats .st{padding:0 20px;}
.cpd-stats .st:first-child{padding-left:0;}
.cpd-stats .st+.st{border-left:1px solid var(--cpd-hair);}
.cpd-stats .v{font-family:var(--cpd-num);font-weight:700;font-size:24px;color:var(--cpd-navy);line-height:1;}
.cpd-stats .v.g{color:var(--cpd-green);}.cpd-stats .v.a{color:var(--cpd-amber);}.cpd-stats .v.c{color:var(--cpd-coral);}
.cpd-stats .v.faint{color:var(--cpd-faint);}
.cpd-stats .l{font-size:9.5px;letter-spacing:.1em;text-transform:uppercase;color:var(--cpd-faint);font-weight:700;margin-top:5px;}
.cpd-engine{font-size:10px;color:var(--cpd-faint);font-weight:500;}
.cpd-engine b{color:var(--cpd-sub);}
.cpd-subnav{position:sticky;top:0;z-index:30;background:rgba(252,252,253,.96);backdrop-filter:blur(10px);border-bottom:1px solid var(--cpd-hair);
  padding:0 30px;display:flex;align-items:stretch;gap:26px;flex-wrap:wrap;}
.cpd-tabs{display:flex;gap:26px;}
.cpd-tab{border:0;background:none;color:var(--cpd-faint);font-family:inherit;font-size:11.5px;font-weight:700;letter-spacing:.06em;
  text-transform:uppercase;padding:14px 2px 12px;cursor:pointer;border-bottom:2px solid transparent;display:flex;gap:8px;align-items:center;}
.cpd-tab .k{font-size:10px;color:var(--cpd-blue2);}
.cpd-tab:hover{color:var(--cpd-sub);}
.cpd-tab.on{color:var(--cpd-navy);border-bottom-color:var(--cpd-navy);}
.cpd-tab.on .k{color:var(--cpd-blue);}
.cpd-badge{font-family:var(--cpd-num);font-size:9px;background:var(--cpd-coral);color:#fff;border-radius:8px;padding:1px 6px;font-weight:700;}
.cpd-navr{margin-left:auto;display:flex;gap:9px;align-items:center;padding:9px 0;}
.cpd-btn{border:1px solid var(--cpd-hair);background:#fff;border-radius:8px;padding:7px 13px;font-size:11px;font-weight:700;color:var(--cpd-sub);cursor:pointer;font-family:inherit;transition:all .15s;}
.cpd-btn:hover{border-color:var(--cpd-blue2);color:var(--cpd-navy);}
.cpd-btn.primary{background:var(--cpd-navy);border-color:var(--cpd-navy);color:#fff;}
.cpd-btn.primary:hover{background:var(--cpd-blue);}
.cpd-btn:disabled{opacity:.45;cursor:default;}
.cpd-wrap{padding:8px 30px 40px;}
.cpd-kick{display:flex;align-items:baseline;gap:12px;margin:26px 0 4px;}
.cpd-kick .no{font-family:var(--cpd-num);font-size:11px;font-weight:700;color:var(--cpd-blue2);letter-spacing:.05em;}
.cpd-kick .t{font-size:10px;letter-spacing:.14em;text-transform:uppercase;color:var(--cpd-faint);font-weight:700;}
.cpd-exh{padding:6px 0 26px;}
.cpd-exh-head{max-width:780px;margin-bottom:16px;}
.cpd-exh-label{font-size:9px;letter-spacing:.14em;text-transform:uppercase;color:var(--cpd-blue2);font-weight:700;}
.cpd-exh-title{font-family:var(--cpd-serif);font-size:19px;font-weight:700;color:var(--cpd-ink);line-height:1.3;margin-top:5px;}
.cpd-exh-title em{font-style:normal;color:var(--cpd-coral);}
.cpd-exh-sub{font-size:11.5px;color:var(--cpd-sub);margin-top:6px;font-weight:500;}
.cpd-src{font-size:9.5px;color:var(--cpd-faint);margin-top:16px;font-weight:500;}
.cpd-empty{text-align:center;padding:44px 20px;color:var(--cpd-faint);}
.cpd-empty p{font-size:12px;max-width:440px;margin:8px auto 0;line-height:1.6;}
.cpd-tag{font-size:9px;font-weight:700;border-radius:5px;padding:2.5px 8px;white-space:nowrap;letter-spacing:.02em;display:inline-block;}
.cpd-tag.g{background:var(--cpd-green-s);color:var(--cpd-green);}
.cpd-tag.a{background:var(--cpd-amber-s);color:var(--cpd-amber);}
.cpd-tag.c{background:var(--cpd-coral-s);color:var(--cpd-coral);}
.cpd-tag.p{background:var(--cpd-purple-s);color:var(--cpd-purple);}
.cpd-tag.b{background:#EAF0FA;color:var(--cpd-blue);}
.cpd-tag.k{background:var(--cpd-hair2);color:var(--cpd-sub);}
/* 01 · esteira */
.cpd-row{display:grid;grid-template-columns:1fr 132px 1fr;align-items:stretch;border:1px solid var(--cpd-hair);border-radius:12px;margin-bottom:14px;background:#fff;overflow:hidden;opacity:1;transition:opacity .4s,transform .4s;}
.cpd-row.oculta{opacity:0;transform:translateY(8px);}
@media(max-width:900px){.cpd-row{grid-template-columns:1fr;}}
.cpd-pgd{padding:15px 18px;border-right:1px solid var(--cpd-hair2);}
.cpd-pgd .who{font-size:9.5px;color:var(--cpd-faint);font-weight:700;letter-spacing:.06em;text-transform:uppercase;display:flex;justify-content:space-between;gap:8px;}
.cpd-pgd .who .h{color:var(--cpd-navy);font-family:var(--cpd-num);}
.cpd-pgd .txt{font-size:12px;color:var(--cpd-ink);margin-top:7px;line-height:1.5;}
.cpd-conf{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:7px;padding:12px 10px;background:var(--cpd-hair2);}
.cpd-conf .pc{font-family:var(--cpd-num);font-size:16px;font-weight:700;line-height:1;}
.cpd-conf .pc.g{color:var(--cpd-green);}.cpd-conf .pc.a{color:var(--cpd-amber);}.cpd-conf .pc.c{color:var(--cpd-coral);}
.cpd-conf .meter{width:74px;height:5px;border-radius:3px;background:#fff;overflow:hidden;}
.cpd-conf .meter i{display:block;height:100%;border-radius:3px;transition:width .7s ease;}
.cpd-conf .lbl{font-size:8.5px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--cpd-faint);text-align:center;line-height:1.3;}
.cpd-conf .aguarda{color:var(--cpd-faint);font-size:11px;font-weight:600;}
.cpd-cat{padding:15px 18px;}
.cpd-cat .cod{font-family:var(--cpd-num);font-size:9.5px;color:var(--cpd-faint);font-weight:700;display:flex;justify-content:space-between;align-items:center;gap:8px;}
.cpd-cat .nm{font-size:12px;font-weight:700;color:var(--cpd-navy);margin-top:6px;line-height:1.45;}
.cpd-cat .hier{font-size:10px;color:var(--cpd-sub);margin-top:4px;}
.cpd-cat.vazio{display:flex;align-items:center;justify-content:center;color:var(--cpd-faint);font-size:11px;font-weight:600;}
.cpd-cta{border:1px dashed var(--cpd-blue2);border-radius:14px;padding:30px;text-align:center;background:#FBFCFE;margin-bottom:22px;}
.cpd-cta p{font-size:12px;color:var(--cpd-sub);max-width:520px;margin:0 auto 16px;}
/* 02 · triagem */
.cpd-kanban{display:grid;grid-template-columns:repeat(3,1fr);gap:22px;}
@media(max-width:1000px){.cpd-kanban{grid-template-columns:1fr;}}
.cpd-kcol-h{display:flex;align-items:baseline;justify-content:space-between;border-bottom:2px solid var(--cpd-ink);padding-bottom:7px;margin-bottom:12px;}
.cpd-kcol-h .t{font-size:10px;letter-spacing:.1em;text-transform:uppercase;font-weight:700;color:var(--cpd-navy);}
.cpd-kcol-h .n{font-family:var(--cpd-num);font-size:13px;font-weight:700;color:var(--cpd-navy);}
.cpd-kcard{border:1px solid var(--cpd-hair);border-radius:11px;background:#fff;padding:13px 15px;margin-bottom:12px;}
.cpd-kcard .txt{font-size:11px;color:var(--cpd-ink);line-height:1.5;}
.cpd-kcard .match{margin-top:9px;padding-top:9px;border-top:1px dashed var(--cpd-hair);}
.cpd-kcard .match .nm{font-size:11px;font-weight:700;color:var(--cpd-navy);}
.cpd-kcard .match .hier{font-size:9.5px;color:var(--cpd-sub);margin-top:2px;}
.cpd-kcard .foot{display:flex;justify-content:space-between;align-items:center;margin-top:10px;gap:8px;flex-wrap:wrap;}
.cpd-kcard .conf-mini{display:flex;align-items:center;gap:6px;font-family:var(--cpd-num);font-size:10.5px;font-weight:700;}
.cpd-kcard .conf-mini .m{width:44px;height:4px;border-radius:2px;background:var(--cpd-hair2);overflow:hidden;}
.cpd-kcard .conf-mini .m i{display:block;height:100%;}
.cpd-kcard .kacts{display:flex;gap:6px;flex-wrap:wrap;}
.cpd-kcard .kacts button{border:1px solid var(--cpd-hair);background:#fff;border-radius:7px;padding:4.5px 9px;font-size:10px;font-weight:700;color:var(--cpd-sub);cursor:pointer;font-family:inherit;}
.cpd-kcard .kacts button:hover{border-color:var(--cpd-blue2);color:var(--cpd-navy);}
.cpd-kcard .kacts button.ok{border-color:var(--cpd-green);color:var(--cpd-green);}
.cpd-kcard .kacts button.ok:hover{background:var(--cpd-green-s);}
.cpd-kcard .kacts button.nv{border-color:var(--cpd-purple);color:var(--cpd-purple);}
.cpd-kcard .kacts button.nv:hover{background:var(--cpd-purple-s);}
.cpd-kcard.resolvida{border-color:var(--cpd-green);background:#FBFEFC;}
.cpd-kcard.rascunho{border-color:var(--cpd-purple);background:#FDFCFE;}
.cpd-kempty{border:1px dashed var(--cpd-hair);border-radius:11px;padding:22px;text-align:center;font-size:10.5px;color:var(--cpd-faint);font-weight:600;}
/* 03 · resultado */
.cpd-tab-wrap{overflow-x:auto;}
table.cpd-tab{width:100%;border-collapse:collapse;font-size:11.5px;}
table.cpd-tab th{font-size:8.5px;letter-spacing:.08em;text-transform:uppercase;color:var(--cpd-faint);font-weight:700;text-align:left;padding:6px 8px 7px;border-bottom:1px solid var(--cpd-ink);}
table.cpd-tab td{padding:9px 8px;border-bottom:1px solid var(--cpd-hair2);vertical-align:middle;}
table.cpd-tab tr:last-child td{border-bottom:0;}
table.cpd-tab td.num,table.cpd-tab th.num{text-align:right;}
table.cpd-tab td.num{font-family:var(--cpd-num);font-weight:600;font-variant-numeric:tabular-nums;}
table.cpd-tab tr.nova td{background:#F7FAF7;}
table.cpd-tab tr.nova td:first-child{box-shadow:inset 3px 0 0 var(--cpd-green);}
.cpd-cov{display:flex;align-items:center;gap:16px;margin:4px 0 20px;flex-wrap:wrap;}
.cpd-cov .cov-bar{flex:1;min-width:200px;height:12px;border-radius:6px;background:var(--cpd-hair2);overflow:hidden;position:relative;}
.cpd-cov .cov-bar i{position:absolute;inset:0 auto 0 0;background:linear-gradient(90deg,var(--cpd-navy),var(--cpd-blue));border-radius:6px;transition:width .6s ease;}
.cpd-cov .cov-num{font-family:var(--cpd-num);font-size:22px;font-weight:700;color:var(--cpd-navy);line-height:1;}
.cpd-cov .cov-lbl{font-size:9.5px;color:var(--cpd-faint);font-weight:700;letter-spacing:.08em;text-transform:uppercase;}
.cpd-confirm{border:1px solid var(--cpd-green);border-radius:12px;background:var(--cpd-green-s);padding:18px 22px;margin-top:18px;display:flex;gap:14px;align-items:center;}
.cpd-confirm .ic{font-family:var(--cpd-num);font-size:22px;font-weight:700;color:var(--cpd-green);}
.cpd-confirm p{font-size:11.5px;color:var(--cpd-ink);margin:0;}
.cpd-foot{padding:14px 30px 26px;font-size:9px;color:var(--cpd-faint);line-height:1.6;border-top:1px solid var(--cpd-hair);}
/* assistente como visualização inline (ocupa o .px-main) */
.px-chat.inline{position:static;width:100%;max-width:100%;height:auto;min-height:calc(100vh - 340px);max-height:none;border-radius:12px;box-shadow:none;border:1px solid ${C.line};}
.px-chat.inline .px-chat-h{padding:14px 18px;gap:12px;flex-wrap:wrap;}
.px-chat-h-sub{color:rgba(255,255,255,.82);font-size:12px;font-weight:500;flex:1 1 auto;text-align:right;}
.px-chat.inline .px-chat-b{padding:22px 28px;gap:14px;min-height:400px;}
.px-chat.inline .px-msg{max-width:78%;font-size:13.5px;}
.px-chat.inline .px-chat-i{padding:14px 18px;}
.px-chat.inline .px-chat-i textarea{max-height:120px;font-size:13.5px;}
/* conversor como visualização inline (sem modal) */
.px-conv.inline{width:100%;max-width:100%;max-height:none;border-radius:12px;box-shadow:none;border:1px solid ${C.line};}
.px-conv.inline .px-conv-body{overflow:visible;}

/* ---------- cadeia de valor em destaque (Descrição da Área) ---------- */
.px-doc-view button.cadeia{color:${C.navy};}
.px-doc-view button.cadeia svg{color:#B86E00;}
.px-doc-view button.cadeia:hover{background:${C.yellow};color:${C.navy};}
.px-doc-view button.cadeia:hover svg{color:${C.navy};}
.px-doc-view button.cadeia.on{background:#B86E00;color:#fff;}
.px-doc-view button.cadeia.on svg{color:#fff;}

/* ---------- explosão solar: lista de entregas do serviço ---------- */
.px-sun2-list{margin-top:14px;background:#fff;border:1px solid ${C.line};border-radius:12px;overflow:hidden;}
.px-sun2-list-h{display:flex;align-items:center;gap:10px;padding:11px 14px;background:${C.bg};border-bottom:1px solid ${C.line};}
.px-sun2-list-back{display:flex;align-items:center;gap:5px;border:1px solid ${C.line};background:#fff;border-radius:8px;padding:6px 10px;font-family:inherit;font-size:12px;font-weight:700;color:${C.primary};cursor:pointer;}
.px-sun2-list-back:hover{border-color:${C.primary};background:${C.primarySoft};}
.px-sun2-list-tt{flex:1;min-width:0;}
.px-sun2-list-tt b{display:block;font-size:13px;font-weight:800;color:${C.navy};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.px-sun2-list-tt span{font-size:11px;color:${C.faint};}
.px-sun2-list .px-erows{padding:4px 8px 8px;max-height:360px;overflow:auto;}

/* ---------- capa de abertura (clara) ---------- */
.px-capa-bg{position:fixed;inset:0;z-index:90;background:rgba(240,243,248,.88);backdrop-filter:blur(5px);display:grid;place-items:center;padding:18px;overflow:auto;}
.px-capa{position:relative;background:#fff;border:1px solid ${C.line};width:820px;max-width:100%;border-radius:20px;padding:38px 42px 26px;box-shadow:0 18px 60px rgba(13,49,111,.14);animation:capaIn .4s cubic-bezier(.2,.8,.3,1) both;}
@keyframes capaIn{from{opacity:0;transform:translateY(12px) scale(.99);}to{opacity:1;transform:none;}}
.px-capa-x{position:absolute;top:16px;right:16px;width:32px;height:32px;border:1px solid ${C.line};background:#fff;border-radius:9px;display:grid;place-items:center;cursor:pointer;color:${C.sub};}
.px-capa-x:hover{border-color:${C.primary};color:${C.primary};}
.px-capa-top{text-align:center;max-width:560px;margin:0 auto 26px;}
.px-capa-top h1{font-size:29px;font-weight:800;letter-spacing:-.015em;color:${C.navy};margin:12px 0 8px;}
.px-capa-tag{font-size:13.5px;line-height:1.6;color:${C.sub};margin:0;}
.px-capa-nums{display:flex;justify-content:center;align-items:center;gap:14px;margin-top:16px;font-size:12.5px;color:${C.sub};flex-wrap:wrap;}
.px-capa-nums b{font-weight:800;color:${C.navy};margin-right:3px;}
.px-capa-nums i{width:4px;height:4px;border-radius:50%;background:${C.line};}
.px-capa-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;}
.px-capa-card{display:flex;flex-direction:column;align-items:flex-start;gap:8px;text-align:left;background:${C.surface};border:1px solid ${C.line};border-radius:13px;padding:16px 15px;cursor:pointer;font-family:inherit;transition:.15s;animation:capaUp .45s both;}
.px-capa-card:hover{border-color:${C.primary};background:#fff;box-shadow:0 4px 16px rgba(19,81,180,.1);transform:translateY(-2px);}
.px-capa-card:focus-visible{outline:2px solid ${C.primary};outline-offset:2px;}
.px-capa-card-ic{width:38px;height:38px;border-radius:10px;display:grid;place-items:center;}
.px-capa-card b{font-size:13.5px;font-weight:800;color:${C.navy};}
.px-capa-card-d{font-size:11.5px;line-height:1.45;color:${C.sub};}
.px-capa-foot{display:flex;justify-content:space-between;align-items:center;gap:14px;margin-top:24px;padding-top:18px;border-top:1px solid ${C.line};flex-wrap:wrap;}
.px-capa-org{font-size:10px;font-weight:700;letter-spacing:.09em;text-transform:uppercase;color:${C.faint};}
.px-capa-entrar{display:flex;align-items:center;gap:8px;background:${C.primary};color:#fff;border:none;border-radius:11px;padding:11px 20px;font-family:inherit;font-size:13.5px;font-weight:800;cursor:pointer;box-shadow:0 3px 12px rgba(19,81,180,.24);}
.px-capa-entrar:hover{background:${C.primaryDark};transform:translateY(-1px);}
@keyframes capaUp{from{opacity:0;transform:translateY(8px);}to{opacity:1;transform:none;}}
@media(max-width:760px){.px-capa{padding:28px 22px 20px;}.px-capa-grid{grid-template-columns:1fr 1fr;}}
@media(max-width:520px){.px-capa-grid{grid-template-columns:1fr;}}
@media(prefers-reduced-motion:reduce){.px-capa,.px-capa-card{animation:none;}}

/* ---------- barra de ferramentas fixa (busca + modos + facetas) ---------- */
.px-toolbar{position:sticky;top:0;z-index:20;background:${C.bg};padding-top:10px;padding-bottom:4px;margin-bottom:10px;box-shadow:0 6px 8px -8px rgba(13,49,111,.18);}
.px-toolbar .px-controls{margin-bottom:8px;}
.px-toolbar .px-facets{margin-bottom:8px;}
.px-toolbar .px-mx{margin:0 0 8px;padding:0;max-width:none;}
.px-toolbar .px-fbar{margin-bottom:8px;}
.px-toolbar .px-filtros-ativos{margin-bottom:2px;}
.px-mode-lbl{display:inline;}
@media(max-width:720px){.px-mode-lbl{display:none;}.px-mode{padding:8px 10px;}}

/* ---------- painel lateral recolhível ---------- */
.px-doc-collapse{margin-left:4px;background:#fff;border:1px solid ${C.line};border-radius:8px;width:26px;height:26px;display:grid;place-items:center;cursor:pointer;color:${C.sub};flex-shrink:0;}
.px-doc-collapse:hover{border-color:${C.primary};color:${C.primary};}
.px-descfab.pxdesc{display:none;background:${C.navy};}
.px-descfab.pxdesc[data-collapsed="1"]{display:flex;}
@media(max-width:1024px){
  .px-doc{display:none!important;}
  .px-body{grid-template-columns:1fr!important;}
  .px-descfab.pxdesc{display:flex!important;}
}

`;


/* ===== módulo embutido: CentralRevisao (tela cheia) ===== */
const CentralRevisao = (function(){

/* ============================================================
   Central de Revisão de Entregas — console do curador (Frente 1)
   3 abas: Similares · Serviços · Nova entrega
   Dados: snapshot embutido + round-trip por CSV (fácil, sem backend)
   Para ligar "ao vivo": trocar carregarBanco() por um fetch do JSON
   publicado do Google Sheets, ou pelo webservice do Serpro.
   ============================================================ */

const C = {
  bg:"#F4F6F9", surface:"#FFFFFF", ink:"#1B1B1B", sub:"#55606E", faint:"#8C97A6",
  line:"#E4E7EB", primary:"#1351B4", primaryDark:"#0C326F", primarySoft:"#E8EEF9",
  green:"#168821", greenSoft:"#E3F2E5", navy:"#13315C", yellow:"#FFCD07",
  amber:"#B86E00", amberSoft:"#FBEEDB", coral:"#9E3B1F", coralSoft:"#FAECE7",
};
const NAT = {
  "Finalístico":{cor:"#168821",soft:"#E3F2E5"},
  "Governança":{cor:"#B86E00",soft:"#FBEEDB"},
  "Suporte":{cor:"#5F5E5A",soft:"#F1EFE8"},
};

const SAMPLE = {"macros": ["Administrar as relações de trabalho", "Administrar bens de consumo e permanente", "Administrar infraestrutura de tecnologia da informação", "Administrar recursos financeiros", "Administrar serviços gerais", "Administrar suporte técnico", "Administração aduaneira", "Apoio ao trabalaho autogestionário, autônomo ou associado", "Apoio à produção agropecuária", "Atendimento ao público e ampliação do acesso à justiça", "Atenção hospitalar e urgência", "Atividade econômica", "Atuar na defesa e promoção dos direitos humanos", "Avaliação e controle", "Ações socioambientais e consolidação territorial em unidades de conservação", "Competitividade e política regulatória", "Comunicação institucional", "Comércio exterior", "Consultoria e assessoramento jurídico", "Cooperação jurídica internacional", "Criação e manejo de unidades de conservação", "Cultura", "Defesa agropecuária", "Defesa dos direitos", "Demarcação e regularização fundiária de terras indígenas", "Desenvolver comunicação institucional", "Desenvolver pessoas", "Desenvolver procedimentos correicionais", "Desenvolver sistemas corporativos", "Desenvolvimento de soluções agropecuárias", "Desenvolvimento e execução do programa espacial brasileiro", "Desenvolvimento e fortalecimento da economia da saúde", "Desenvolvimento estratégico do setor turístico", "Desenvolvimento industrial, inovação, comércio e serviços", "Desenvolvimento sistemas corporativos", "Diretrizes estratégicas para comunicações", "Economia verde, descarbonização e bioindústria", "Educação superior", "Enfrentamento e prevenção à corrupção", "Estímulo à eficiência, Inovação e competitividade", "Excluir", "Expansão e massificação das comunicações", "Financiamento público e operações de fomento e subvenções", "Fiscalização e conformidade regulatória do setor elétrico", "Fiscalização e monitoramento", "Fiscalização tributária", "Fomento ao desenvolvimento científico, tecnológico e produtivo em saúde", "Fomento ao desenvolvimento sustentável da bioeconomia", "Fomento à agregação de valor e diferenciação", "Fomento à sustentabilidade e qualidade da produção e das práticas agropecuárias", "Formulação, acompanhamento e avaliação de políticas públicas", "Fortalecimento dos produtores rurais", "Gerenciar a documentação arquivística", "Gerenciar a qualidade", "Gerenciar a renda do patrimônio indígena", "Gerenciar contratações", "Gerenciar desempenho de pessoas", "Gerenciar documentos arquivísticos", "Gerenciar informações cadastrais de pessoal", "Gerenciar manifestações de ouvidoria", "Gerenciar o acesso à informação", "Gerenciar processos de negócio", "Gerenciar programas e projetos", "Gerenciar recrutamento e seleção de pessoas", "Gerenciar riscos corporativos", "Gerir projetos de implementação de tecnologias sociais de acesso à água para consumo", "Gestão contábil e de custos do setor público", "Gestão da Infraestrutura hídrica regional", "Gestão da administração financeira e fiscal", "Gestão da cultura e do desenvolvimento socioeconômico afro-brasileiro", "Gestão da dívida ativa", "Gestão da logística pública", "Gestão da politica de geologia, mineração e transformação mineral", "Gestão da política da aviação civil", "Gestão da política das Infraestruturas hidroviárias", "Gestão da política de atenção especializada à saúde", "Gestão da política de atenção primária à saúde", "Gestão da política de desenvolvimento de atividades espaciais", "Gestão da política de fomento à produção agropecuária", "Gestão da política de hidrovias e navegação", "Gestão da política de infraestrutura aeroportuária", "Gestão da política de navegação marítima", "Gestão da política de pessoal da administração pública federal", "Gestão da política de petróleo, gás natural e biocombustíveis", "Gestão da política de preservação e difusão da produção bibliográfica e documental do país", "Gestão da política de proteção de dados pessoais", "Gestão da política de proteção e defesa do consumidor", "Gestão da política de saúde Indígena", "Gestão da política de segurança pública nacional", "Gestão da política de sustentabilidade, inovação e qualidade da produção agropecuária", "Gestão da política de transição energética e planejamento", "Gestão da política e salvaguarda da memória  bibliográfica e documental do país", "Gestão da política portuária", "Gestão da política pública de assistência social", "Gestão da transferência condicionada e direta de renda", "Gestão das infraestruturas hidroviárias", "Gestão das políticas tributária e aduaneira", "Gestão de ativos", "Gestão de benefícios trabalhistas", "Gestão de contabilidade pública", "Gestão de créditos e defesa do interesse público​", "Gestão de informações meteorológicas nacionais", "Gestão de informações populacionais, econômicas e geográficas nacionais", "Gestão de insumos estratégicos aplicados à saúde", "Gestão de investimentos aeroportuários e aviação civil", "Gestão de passivos", "Gestão de pessoas", "Gestão de política de exploração e produção de petróleo e gás natural", "Gestão de políticas de cooperação internacional, comércio exterior e integração regional e global​", "Gestão de políticas de energia elétrica", "Gestão de políticas de estímulo à eficiência, inovação e competitividade​", "Gestão de políticas de formação artística, livro e leitura", "Gestão de políticas de incentivo à agricultura familiar e a povos e comunidade tradicionais", "Gestão de políticas de justiça e cidadania", "Gestão de políticas de transporte aquaviário", "Gestão de políticas econômicas para transformação ecológica​", "Gestão de políticas públicas", "Gestão de políticas públicas de direitos humanos e cidadania", "Gestão de políticas públicas de segurança alimentar e nutricional", "Gestão de políticas públicas de vigilância em saúde e ambiente", "Gestão de processos de residência médica", "Gestão de registros administrativos e de informações estatísticas do mercado de trabalho", "Gestão de relações do trabalho no serviço público", "Gestão de tecnologias da informação e comunicação", "Gestão de trabalho em saúde", "Gestão de transferências e repasses de recursos da União", "Gestão do conhecimento e da informação corporativa", "Gestão do crédito tributário e da arrecadação", "Gestão do crédito tributário, administrativo e arrecadação", "Gestão do desenvolvimento da ciência, tecnologia e inovação (CT&I)", "Gestão do desenvolvimento organizacional e da inovação", "Gestão do planejamento setorial de mineração", "Gestão do processo administrativo digital", "Gestão do sistema de controle interno da administração pública federal", "Gestão do trabalho em saúde", "Gestão dos créditos e defesa do interesse público", "Gestão dos recursos FAT e FGTS", "Gestão e controle institucional", "Gestão estratégica e governança", "Gestão financeira, contábil e de custos", "Gestão fiscal e sustentabilidade financeira", "Gestão integrada da representação e defesa da fazenda nacional", "Gestão integrada das políticas econômicas", "Gestão jurídica institucional", "Gestão pública orientada por evidências", "Governança do setor espacial", "Governança do sistema nacional de cultura", "Governança e coordenação das empresas estatais", "Governança e gestão institucional", "Inspeção de produtos de origem animal e vegetal", "Inspeção do trabalho", "Justiça e cidadania", "Licenciamento de entidades e planos de previdência complementar", "Macroprocesso sem nome", "Modelar a estrutura organizacional", "Monitorar a gestão da comunicação institucional", "Monitorar a gestão das transferências da União", "Monitorar a gestão do desenvolvimento organizacional e da inovação", "Monitorar a manutenção predial", "Monitorar a vida funcional do servidor", "Normatização de setores econômicos sensíveis", "Normatização e efetivação das políticas de comunicações", "Normatização e orientação técnica", "Outorga e contratação de serviços de energia elétrica", "Participação e controle social", "Pesquisa, avaliação e monitoramento da biodiversidade", "Planejamento e gestão de recursos externos", "Planejamento e orçamento", "Planejamento estratégico fiscal", "Planejamento fiscal", "Planejar a gestão da informação e documentação", "Planejar a gestão de logística pública", "Política econômica no setor portuário", "Preservação do patrimônio cultural", "Preservação e difusão do patrimônio histórico, artístico e cultural", "Preservação e gestão do patrimônio documental e bibliográfico nacional", "Prestação de assistência jurídica integral e gratuita, individual e coletiva", "Prevenção controle e erradicação de doenças e pragas", "Prevenção e solução de litígios tributários e aduaneiros", "Produção de inteligência financeira contra a lavagem de dinheiro", "Programar orçamento institucional", "Promover o desenvolvimento sustentável", "Promover saúde, segurança e qualidade de vida", "Promoção da integridade pública, transparência e acesso a informação", "Promoção da política nacional de arquivos", "Promoção da segurança jurídica", "Promoção do acesso ao patrimônio documental nacional", "Promoção do acesso e do consumo de alimentação adequada e saudável", "Promoção do diálogo social nas relações de trabalho", "Promoção e fomento de ciência, tecnologia e inovação (CT&I)", "Promoção e proteção dos direitos e cidadania dos povos indígenas", "Promoção à inserção e permanência no trabalho", "Proteção", "Proteção de terras indígenas e de povos isolados e de recente contato", "Proteção às comunidades quilombolas", "Prover pessoas", "Recompensar direitos e vantagens", "Reconhecer direitos previdenciários e estatutários", "Reconhecimento das ocupações e profissões", "Regulamentação e fiscalização da atividade do setor turístico", "Regulação de serviços de saúde", "Regulação do Estado", "Regulação do setor elétrico", "Regulação e fiscalização", "Regulação e supervisão tributária e financeira", "Relacionamento com sociedade, governos e organismos nacionais e internacionais", "Relações institucionais e federativas", "Segurança jurídica e ambiente econômico​", "Segurança pública", "Subsidiar defesa da União", "Sustentabilidade e responsabilidade socioambiental na gestão da PGFN", "Transferência de recursos", "Transformação digital", "Transparência das contas públicas", "Transparência, articulação e responsabilização institucional", "Tributação"], "cats": ["Acessibilidade e inclusão", "Acesso ao espaço e missões espaciais", "Acesso à alimentação adequada e saudável", "Acesso à informação", "Acompanhamento da evolução da dívida pública", "Acompanhamento da legislação", "Acompanhamento da operação do  Sistema Elétrico Brasileiro (SEB)", "Acompanhamento de fundos", "Acompanhamento e avaliação de programas e ações de comunicações", "Acompanhamento e monitoramento do Bolsa Família", "Acompanhamento e sugestão de aprimoramentos na legislação", "Acompanhar e subsidiar a avaliação das políticas públicas implementadas com benefício fiscal", "Acordo de leniência", "Administração de haveres mobiliários", "Aferição de conformidade", "Alimentação juventudes", "Alimentação saudável", "Ambiente alimentar escolar", "Anistia política", "Análise de aspectos tributários no comércio exterior", "Análise de condutas e atos de concentração", "Análise de dados operacionais e padronização de procedimentos", "Análise de políticas econômicas", "Análise e aplicação da legislação de pessoal", "Análise epidemiológica", "Análise política comercial", "Aperfeiçoar atuação por recorte temático", "Aperfeiçoar o diálogo externo", "Aplicar os recursos da renda do patrimônio indígena", "Apoiar ações para o fortalecimento de equipamentos e tecnologias sociais de segurança alimentar e nutricional", "Apoiar ações voltadas para implementação de tecnologias sociais de acesso à água", "Apoio ao desenvolvimento da economia popular e solidária", "Apoio aos órgãos de controle", "Apoio e orientação à gestão descentralizada e participativa do SUAS", "Apoio à gestão descentralizada do programa bolsa família", "Apoio, fomento e economia do patrimônio cultural", "Arrecadação e inadimplência", "Articulação e Parceria para Viabilização de Políticas e Comunicações", "Articulação institucional", "Articulação institucional setorial", "Articulação institucional.", "Assessoramento", "Assessoria de Imprensa", "Assessoria de controle interno", "Assessoria especial de controle interno", "Assessoria parlamentar", "Assessoria técnica", "Assistência farmacêutica e insumos estratégicos para o SUS", "Atendimento", "Atendimento a demandas informacionais de empresas estatais federais", "Atendimento ao público", "Atendimento inicial ao público para prestação de assistência jurídica", "Atendimento à saúde", "Atenção especializada à saúde", "Atenção hospitalar, domiciliar e de urgência", "Atenção integral à saúde da criança e do adolescente", "Atenção integral à saúde da mulher", "Atenção integral à saúde da pessoa idosa", "Atividade correicional", "Atividades de inteligência e segurança institucional", "Atuação Técnica e Decisória do Tribunal", "Atuação junto a órgãos de controle, transparência pública e responsabilização administrativa", "Atuação na mediação e transação tributária para redução de litígios", "Auditoria", "Automação, padronização e racionalização de rotinas jurídicas e de cobrança", "Autorizar o ingresso em terra indígena para os casos que o requeiram", "Autorização para intervenção e circulação de bens culturais", "Avaliação de Impactos Ambientais", "Avaliação de conformidade", "Avaliação de impactos financeiros e fiscais de decisões normativas e judiciais consolidadas", "Avaliação de serviços públicos", "Avaliação e Monitoramento da Efetividade de Gestão das Unidades de Conservação", "Ações de atenção à saúde indígena", "Ações estruturantes em economia e desenvolvimento em saúde", "Ações investigativas", "Bioeconomia", "Bioindústria", "Boletim de monitoramento do Setor Elétrico", "CONDUZIR RELAÇÕES INSTITUCIONAIS EM ÂMBITO NACIONAL", "CONDUZIR RELAÇÕES INTERNACIONAIS", "CONTROLAR BENEFÍCIOS FISCAIS E REGIMES ESPECIAIS DE TRIBUTAÇÃO", "CONTROLAR E FISCALIZAR PROCESSOS ADUANEIROS DIFERENCIADOS", "CONTROLAR E FISCALIZAR PROCESSOS DE IMPORTAÇÃO E EXPORTAÇÃO", "CONTROLAR E FISCALIZAR REGIMES ADUANEIROS", "Cadastro e registro de pessoal", "Cadeias produtivas dos biomas e da Amazônia", "Cadeias produtivas e indicações geográficas", "Capacitação digital", "Central do disque 100", "Certificação de comunidades quilombolas", "Certificação de entidades beneficentes de assistência social em saúde", "Cesta Básica", "Classificação indicativa", "Cobrança e recuperação de crédito", "Coleta, tratamento e análise de informações financeiras", "Combate à Tortura", "Combate à corrupção e à lavagem de dinheiro", "Comercialização de energia elétrica no ambiente regulado", "Compensação previdenciária", "Competitividade exportadora e articulação internacional", "Complexo econômico-industrial da saúde e de inovação para o SUS", "Complexo eletroeletrônico e de semicondutores", "Compras e licitação", "Compras governamentais para o desenvolvimento industrial", "Comunicação institucional", "Comunicação interna", "Comércio e sustentabilidade", "Concessão de diárias e passagens", "Concessão do registro profissional e de contratante", "Concessões aeroportuárias", "Conciliação judicial e extrajudicial", "Condução de relações institucionais em âmbito nacional", "Conservação do patrimônio cultural material", "Consolidação e divulgação de informações e relatórios contábeis, orçamentários e fiscais da União", "Consolidação territorial das UCs", "Constituição de reservas indígenas", "Consultoria e assessoramento jurídico", "Consultoria jurídica institucional", "Contabilidade", "Contato nacional", "Contencioso administrativo", "Contencioso administrativo fiscal", "Contencioso internacional", "Contencioso judicial", "Controlar o Cumprimento das Obrigações Acessórias", "Controle interno", "Controle social na saúde indígena", "Convergência regulatória e barreiras às exportações", "Cooperação Internacional em Proteção de Dados Pessoais", "Cooperação jurídica internacional", "Coordenação da elaboração e revisão das normas regulamentadoras", "Corregedoria e responsabilização", "Criação e planejamento territorial de UCs", "Crédito e Garantia", "Defesa Agropecuária - Laboratórios", "Defesa Comercial", "Defesa agropecuária e vigilância  internacional", "Defesa judicial e administrativa da União em matéria fiscal e tributária", "Definição da estratégia de administração da Dívida Pública Federal", "Definição do posicionamento do Brasil​", "Demandas judiciais", "Demarcação de terras indígenas", "Descarbonização", "Descolamentos oficiais por necessidade de serviço", "Desenvolvimento científico e tecnológico em saúde", "Desenvolvimento da Indústria da Construção Civil", "Desenvolvimento da Infraestrutura", "Desenvolvimento da indústria automotiva", "Desenvolvimento de Programas, Instrumentos e Ferramentas de Gestão de Documentos", "Desenvolvimento de ações educativas, orientativas e de prevenção de acidentes do trabalho", "Desenvolvimento de capacitação, educação e ATER", "Desenvolvimento de políticas de conformidade tributária", "Desenvolvimento e capacitação", "Desenvolvimento organizacional", "Desenvolvimento produtivo e economia da saúde do Setor Portuário", "Desenvolvimento socioambiental do setor mineral", "Desintrusão de ocupantes não indígenas", "Deslocamentos oficiais por necessidade do serviço", "Diagnóstico e identificação de necessidade de intervenção regulatória", "Difusão de acervo e conhecimento científico, cultural e artístico", "Difusão do acervo e promoção do acesso ao acervo bibliográfico e documental nacional", "Direitos Humanos - Memória e verdade", "Disseminação, articulação e relações institucionais", "Divulgação de informativos institucionais", "Docência - Magistério Federal", "Doenças Crônicas e Raras", "Doenças Transmissíveis", "Economia Verde e de Impacto", "Economias da Sociobiodiversidade", "Educação Alimentar e Nutricional (EAN)", "Educação Patrimonial", "Eficiência energética", "Elaboração de políticas e diretrizes relacionadas ao federalismo fiscal​", "Elaboração e atualização de normas regulatórias", "Emergência em Saúde Pública", "Emissão de declaração de reconhecimento de limites", "Empreendedorismo Inovador e Inovação Aberta", "Enfrentamento à violência de Crianças e Adolescentes", "Engenharia e arquitetura", "Ensino - Gestão Acadêmica", "Ensino - Serviço de Laboratório Acadêmico", "Escrituração Fiscal Digital (EFD)", "Escuta ativa e participação cidadã", "Estabelecimento de normativos para a regulamentação de produtos agroalimentares artesanais", "Estabilização e previsibilidade de decisões e interpretações", "Estatísticas de Comércio Exterior", "Estratégia de Prevenção da Obesidade", "Estratégias em Epidemiologia e Saúde e Ambiente", "Estratégias para Conservação de Espécies e Habitats", "Estruturação e governança do setor espacial", "Estudos de Comércio Exterior", "Estudos econômicos aplicados à defesa da concorrência", "Estudos na área de petróleo e gás natural", "Etnodesenvolvimento", "Execução da contabilidade pública e de custos na União", "Execução fiscal e recuperação de ativos inadimplidos", "Expansão da geração de energia elétrica", "Expansão da transmissão de energia elétrica", "Facilitação do Comércio", "Facilitação do Transporte Aéreo", "Financeiro", "Financeiro - Contabilidade", "Financeiro -Pagamento", "Financiamento da Educação", "Financiamento externo", "Financiamento público", "Finanças Verdes", "Finanças e Conformidade", "Fiscalização", "Fiscalização Econômica, Financeira e de Mercado do Setor Elétrico", "Fiscalização das obrigações regulatórias", "Fiscalização de estabelecimentos e produtos de origem animal e vegetal", "Fiscalização do trabalho", "Fiscalização e Auditoria em Proteção de Dados Pessoais", "Fiscalização, Monitoramento e Avaliação do Patrimônio Cultural", "Fiscalização/Supervisão", "Fomentar a participação Integral da DPU em redes de Direitos Humanos", "Fomentar a participação integral da DPU em redes de Direitos Humanos", "Fomento a Inovação Agropecuária", "Fomento a equideocultura nacional", "Fomento a produção e uso de bioinsumos agropecuários", "Fomento ao Setor de Construção Naval, Hidroviária e Portuária", "Fomento ao desenvolvimento de pesquisas e estudos agropecuários", "Fomento ao ecossistema de inovação agropecuária", "Fomento da ciência, tecnologia e inovação (CT&I)", "Fomento e comercialização no setor espacial", "Fomento à produção agropecuária", "Fomento à sustentabilidade e qualidade da produção e das práticas agrícolas", "Fomento, promoção e divulgação da cultura afro-brasileira", "Formação de preço", "Formação e Qualificação em Oncologia", "Formular Atos Interpretativos e Normativos", "Formular e analisar propostas de política tributária e aduaneira", "Formulação da estratégia fiscal", "Formulação de Políticas Públicas de Comunicações", "Formulação de atos interpretativos e normativos", "Fortalecer e capacitar as estruturas de atuação coletiva", "Fortalecimento de mecanismos de saúde dos animais, sanidade das plantas e seus produtos", "Fortalecimento dos Equipamentos de Segurança Alimentar e Nutricional (EqSAN)", "GERENCIAR INTERVENIENTES NO COMÉRCIO EXTERIOR", "GERENCIAR O ATENDIMENTO", "GERENCIAR RISCOS DE CONFORMIDADE ADUANEIRA", "GERIR A ARRECADAÇÃO", "GERIR A CAPTAÇÃO E INTEGRAÇÃO DE DADOS DE INTERESSE TRIBUTÁRIO", "GERIR ATIVIDADES DE PROTOCOLO, TRIAGEM E PREPARO DE PROCESSO ADMINISTRATIVO DIGITAL", "GERIR CADASTROS TRIBUTÁRIOS E ADUANEIROS", "GERIR O CRÉDITO TRIBUTÁRIO", "GERIR TRATAMENTO DE EXPEDIENTE", "Gabinetes - especificar", "Garantir segurança na atuação", "Georreferenciamento de terras indígenas", "Gerenciamento dos riscos na produção, trânsito e comércio de animais, vegetais e seus produtos", "Gerir Acervo, Modelos, Classificação e Ciência de Documentos Administrativos", "Gerir Direito Creditório de Contribuinte", "Gerir a execução orçamentária, gestão de pessoas,  folhas de pagamento, rotinas de cadastro e concessão de benefícios", "Gerir as ações relacionadas à Proteção Social Especial", "Gerir ações relacionadas à proteção social especial em situações de calamidades públicas e emergências no SUAS", "Gerir informação territorial", "Gestão", "Gestão  de painéis de monitoramento", "Gestão Ambiental", "Gestão Bibliotecária", "Gestão Documental", "Gestão Hospitalar", "Gestão Patrimonial", "Gestão Socioambiental", "Gestão acadêmica e suporte a atividades de ensino", "Gestão administrativa de pessoal", "Gestão da Autorização de Funcionamento de Organização Estrangeira em Território Brasileiro – OE", "Gestão da Conta Única", "Gestão da Experimentação Animal", "Gestão da Informação no âmbito da Política de Pessoal de Empresas Estatais Federais", "Gestão da Política de Navegação interior", "Gestão da Qualificação de Organizações da Sociedade Civil de Interesse Público - OSCIP", "Gestão da Rede Socioassistencial Privada do SUAS", "Gestão da emissão da carteira de trabalho e previdência social", "Gestão da execução orçamentária e finaceira do FAT", "Gestão da extensão universitária", "Gestão da governança dos créditos tributários e dos mecanismos de incentivo fiscal", "Gestão da informação", "Gestão da informação corporativa", "Gestão da informação, do acervo e da memória da cultura afro-brasileira", "Gestão da inovação", "Gestão da política de biocombustíveis", "Gestão da política de derivados do petróleo", "Gestão da política de mineração", "Gestão da política nacional de gás natural", "Gestão da políticas sociais de energia elétrica", "Gestão da transição energética", "Gestão da universalização do acesso à energia elétrica", "Gestão das ações e serviços sine", "Gestão das captações e da execução orçamentária e financeira da Dívida Pública Federal", "Gestão das políticas de eólica offshore", "Gestão das políticas públicas agropecuárias", "Gestão das transferências voluntárias", "Gestão de Acervos Históricos e Culturais", "Gestão de Bens Culturais", "Gestão de Carreiras", "Gestão de Compensação Ambiental", "Gestão de Conversão de Multas", "Gestão de Custos", "Gestão de Custos para o SUS e Gestão de Custos do Ministério da Saúde", "Gestão de Demandas de Concursos Públicos", "Gestão de Demandas de Contratação Temporária", "Gestão de Demandas de Provimento de Cargos", "Gestão de Demandas do Banco de Professor e do Quadro Fixo das Instituições Federais de Ensino", "Gestão de Desempenho de Portos Delegados", "Gestão de Exames e Pré-testes educacionais", "Gestão de Fundos de Comunicações", "Gestão de Garantias e Contragarantias", "Gestão de Outorga de Geração de Energia", "Gestão de Outorga de Transmissão de Energia", "Gestão de Pagamentos", "Gestão de Processos", "Gestão de Riscos", "Gestão de Serviços Postais", "Gestão de análise para ressarcimento dos benefícios do Programa Bolsa Família", "Gestão de ativos apreendidos de tráfico / crimes conexos", "Gestão de ações para prevenção e vigilância do câncer", "Gestão de bases de dados e informações econômicas", "Gestão de benefícios do Programa Bolsa Família", "Gestão de cobrança de ressarcimento aos beneficiários do Programa Bolsa Família", "Gestão de condicionalidades do Programa Bolsa Família", "Gestão de contratações de TIC", "Gestão de contrato de estágio", "Gestão de contratos", "Gestão de contratos de prestação de serviços das ações de transferência direta de renda", "Gestão de contratos de residentes​", "Gestão de contratos de voluntários​", "Gestão de convênios", "Gestão de desempenho individual", "Gestão de eventos", "Gestão de haveres financeiros", "Gestão de indicadores", "Gestão de informações de sensoriamento remoto", "Gestão de infraestrutura de TIC", "Gestão de mercadorias apreendidas", "Gestão de obras hídricas", "Gestão de pagamentos", "Gestão de parcerias com organizações não governamentais", "Gestão de parcerias institucionais", "Gestão de parcerias institucionais.", "Gestão de passivos contingentes", "Gestão de políticas de Proteção Social Básica", "Gestão de políticas de incentivos fiscais para geração de energia elétrica", "Gestão de políticas de integração de energia elétrica internacional", "Gestão de políticas de promoção de eficiência econômica​", "Gestão de políticas públicas", "Gestão de processos", "Gestão de programas", "Gestão de programas de qualificação social e profissional", "Gestão de programas do PPA", "Gestão de projetos", "Gestão de projetos geológicos e controle de amostras", "Gestão de recursos hídricos no âmbito do setor elétrico", "Gestão de registro de entidades sindicais", "Gestão de relacionamento, diálogo e comunicação sobre a gestão de Benefícios do Programa Bolsa Família", "Gestão de riscos fiscais de médio e longo prazos", "Gestão de sistemas informatizados", "Gestão de soluções e serviços digitais", "Gestão do Fundo Nacional de Assistência Social", "Gestão do Setor Aeroportuário", "Gestão do Setor Hidroviário", "Gestão do Setor de Navegação Marítima", "Gestão do Sistema Nacional do Patrimônio Cultural (SNPC)", "Gestão do Sistema Único de Assistência Social (SUAS)", "Gestão do abono salarial", "Gestão do conhecimento institucional", "Gestão do contencioso administrativo de ações fiscais", "Gestão do planejamento setorial de mineração", "Gestão do planejamento territorial da mineração", "Gestão do portfólio de produtos meteorológicos", "Gestão do programa de geração de emprego e renda", "Gestão do programa energias da Amazônia", "Gestão do programa nacional do hidrogênio", "Gestão do programa nacional do microcrédito produtivo orientado", "Gestão do seguro-desemprego", "Gestão do setor espacial", "Gestão do sistema de gestão de documentos e arquivos - SIGA", "Gestão do sistema de orçamento público do SUS", "Gestão do Índice de Gestão Descentralizada do Programa Bolsa Família e do Cadastro Único - IGD-PBF", "Gestão documental e arquivística", "Gestão dos Contratos de Bens, Serviços e Insumos de Saúde Indígena", "Gestão dos Créditos e Defesa do Interesse Público", "Gestão dos Serviços de Radiodifusão", "Gestão dos contratos do Auxílio Emergencial (AE) e Auxílio Emergencial Residual (AER)", "Gestão dos débitos inscritos", "Gestão e Qualificação do Programa Nacional de Agricultura Urbana e Periurbana", "Gestão e tratamento de dados", "Gestão e tratamento de dados.", "Gestão estratégica do contencioso e da jurisprudência", "Gestão financeira", "Gestão, fomento e conservação de recursos genéticos", "Governança Estratégica de Sustentabilidade e Responsabilidade Socioambiental", "Governança de TIC", "Governança e Gestão Institucional", "Governança societária", "Grupos de Trabalho e Comitês", "Habilitação, monitoramento societário e econômico-financeiro", "IDENTIFICAR E MEDIR RISCOS DE CONFORMIDADE TRIBUTÁRIA", "Identificação do Patrimônio Cultural", "Identificação e delimitação de terras indígenas", "Implementação de políticas de apoio à produção e comercialização", "Implementação de soluções tecnológicas para gestão da dívida ativa e recuperação de créditos", "Incentivo a produtores rurais", "Inclusão Digital e Acesso à Informação", "Incorporação de Tecnologias em Saúde", "Indicadores econômicos e fiscais", "Indústria Aeroespacial, Defesa e Nuclear", "Indústria da Saúde", "Indústria de Bens de Consumo Não Duráveis e Semiduráveis", "Indústria de Metalurgia e de Base Florestal", "Indústria de Petróleo e Gás", "Informação em saúde", "Informações e relatórios contábeis, orçamentários e fiscais da União", "Infraestrutura da Qualidade", "Inovação Tecnológica e Produtiva", "Instâncias de Controle social", "Insumos estratégicos da saúde", "Integração com o SUS", "Integração urbana", "Integridade e gestão de riscos", "Integridade privada", "Integridade pública", "Inteligência Mercadológica do Setor Turístico", "Inteligência fiscalizatória", "Inteligência regulatória", "Intervir em processos de licenciamento ambiental de terceiros que impactem povos indígenas", "Investigação patrimonial e combate à evasão fiscal e fraudes estruturadas", "Investimentos e Parcerias Estratégicas para o Desenvolvimento do Setor Turístico", "Investimentos em Infraestrutura aeroportuária e infraestrutura aeronáutica regional", "JULGAR RECURSOS ADMINISTRATIVOS", "Julgamento de recursos administrativos", "Laboratórios de Saúde Pública", "Licenciamento", "Licenciamento Ambiental", "Logística - Segurança Interna", "Logística - importação / exportação", "MONITORAR E AVALIAR O SISTEMA TRIBUTÁRIO E ADUANEIRO", "Melhoria regulatória e do ambiente de negócios​", "Meteorologia Geral e Aplicada", "Modernização Regulatória", "Modernização da Gestão Portuária", "Modernização e estruturas organizacionais", "Monitoramento da qualidade e segurança de produtos de origem animal e vegetal", "Monitoramento das remunerações, receitas e aplicações do FAT", "Monitoramento de Mercado de Consumo", "Monitoramento do setor elétrico", "Monitoramento do setor energético", "Monitoramento e regulação da exploração e produção de petróleo e gás", "Monitoramento e regulação de contratos, concessões e instrumentos financeiros", "Negociação Tributária e Mediação para redução de litígios", "Negociações Extrarregionais", "Negociações Internacionais", "Normas de contabilidade pública, de custos e de elaboração dos demonstrativos fiscais", "Normatização", "Normatização de Patrimônio Cultural", "Obtenção de dados meteorológicos", "Operacionalizar a Gestão de Benefícios", "Operações", "Operações e Colegiados", "Organização/ promoção de eventos institucionais", "Orientação/Atendimento a órgãos setoriais", "Orçamento", "Outorga específica para exploração de aeródromos civis", "Outorgas de geração", "Outorgas portuárias", "Outorgas transmissão e distribuição e Planejamento da Transmissão", "Ouvidoria", "PLANEJAR A FISCALIZAÇÃO TRIBUTÁRIA", "PROPOR A UNIFORMIZAÇÃO DA JURISPRUDÊNCIA", "Padronizar a assistência jurídica", "Pagamento de Pessoal", "Participação em conselhos", "Patrimônio Genético", "Pesquisa e Monitoramento da Biodiversidade", "Pesquisa, Avaliação e Monitoramento da Biodiversidade", "Pesquisa, desenvolvimento e inovação", "Pesquisa, estudos e estatísticas", "Planejamento Financeiro e Territorial Integrado", "Planejamento Plurianual da APF", "Planejamento da Infraestrutura aeroportuária nacional", "Planejamento da Política Portuária", "Planejamento da Transmissão", "Planejamento da força de trabalho", "Planejamento das ações fiscalização do trabalho", "Planejamento de geração de energia elétrica", "Planejamento do Setor de Hidrovias", "Planejamento e a Programação Financeira", "Planejamento e orçamento", "Planejamento e relatoria anual de auditoria", "Planejamento energético", "Planejamento logístico", "Planejamento logístico​", "Planejamento, Avaliação e Dimensionamento de Profissionais de Saúde", "Planejamento, monitoramento e avaliação", "Planejamento, monitoramento e avaliação de ações prioritárias", "Planejamento, monitoramento e avaliação do Setor de Aviação Civil", "Planejamento, monitoramento e execução da atuação itinerante", "Planejar a proteção territorial de terras indígenas", "Política Nacional de Transição Energética - PNTE", "Política Regulatória", "Política de Pessoal das Empresas Estatais", "Política de gastos setoriais", "Políticas de investimento", "Políticas e programas do setor espacial", "Políticas migratórias", "Políticas para comércio e serviços digitais", "Políticas para serviços", "Políticas públicas para povos e comunidades tradicionais", "Povos e Comunidades Tradicionais (PCT)", "Preservação das manifestações culturais afro-brasileiras", "Preservação do patrimônio documental nacional", "Preservação dos acervos históricos, culturais e artísticos", "Preservação e salvaguarda do acervo bibliográfico e documental nacional", "Prestação de contas", "Prevenção às doenças crônicas", "Prevenção, controle e erradicação de doenças e pragas do cacaueiro", "Previdência e benefícios", "Priorizar atividades itinerantes com foco na redução da miséria extrema", "Procedimentos Administrativos", "Processo de residência  médica", "Processos de previsão numérica e ciência de dados meteorológicos", "Produção de dados e gestão do Conhecimento", "Produção de publicações institucionais e documentos oficiais", "Produção de publicações institucionais e documentos oficiais 1", "Produção e Divulgação Científica na Área de Oncologia", "Produção e disseminação de estatísticas e informações sobre mercado de trabalho", "Produção, análise e disseminação de dados e informações institucionais", "Programa de gestão e melhoria da qualidade - PGMQ", "Programa mais médicos para o Brasil", "Programas e Projetos Educacionais", "Programas e projetos para o setor de exploração e produção de petróleo (E&P)", "Programação de financiamento da atenção primária", "Projeção de riscos e impactos fiscais da litigiosidade tributária em curso", "Promover a Cidadania dos Povos Indígenas", "Promover a governança e a gestão estratégica interna", "Promover e Proteger os Direitos Sociais Indígenas", "Promoção da Inovação e Sustentabilidade da Cadeia Produtiva do Setor Turístico", "Promoção da alimentação e nutrição na saúde", "Promoção da aprendizagem profissional", "Promoção da atividade física", "Promoção da ciência, tecnologia e inovação (CT&I)", "Promoção da inclusão produtiva e mobilidade social do produtor rural", "Promoção da integração das políticas de Muncípios, Estados e Governo", "Promoção da sustentabilidade", "Promoção das ações voltadas à redução das perdas e desperdício de alimentos", "Promoção das exportações", "Promoção de Equidade e Determinantes Sociais em Saúde", "Promoção de ações de estruturação e fortalecimento das cadeias produtivas regionais", "Promoção de ações de formação em economia popular e solidária", "Promoção de ações de inclusão produtiva rural", "Promoção de ações de redução do impacto ambiental e do uso sustentável dos recursos naturais na produção", "Promoção de ações educativas", "Promoção de políticas e ações indutoras para transformação ecológica​", "Promoção do conhecimento sobre recursos naturais e das boas práticas agropecuárias", "Promoção do desenvolvimento socioeconômico do entorno do Centro Espacial de Alcântara", "Promoção do uso de boas práticas e sistemas agropecuários", "Promoção e Difusão do Patrimônio Cultural", "Promoção e Qualificação dos Processos Educativos Indígenas", "Promoção e divulgação cultural e científica", "Promoção à mediação e negociação coletiva", "Promoção, Estruturação e Qualificação de Produtos, Serviços e Destinos Turísticos", "Promoção, Expansão e Modernização das Comunicações", "Promoção/organização de eventos institucionais", "Propriedade Industrial - Contratos de Tecnologia", "Propriedade Industrial - Desenhos Industriais", "Propriedade Industrial - Estatísticas e Estudos", "Propriedade Industrial - Informação Tecnológica de Patentes", "Propriedade Industrial - Recursos e processos administrativos de nulidade", "Propriedade Industrial - Registro de Marcas", "Propriedade Industrial - concessão de patentes", "Propriedade Industrial - registro de indicação geográfica", "Propriedade Industrial - registro de marca via protocolo de Madri", "Propriedade Industrial - registro de software", "Propriedade Intelectual", "Proteção", "Proteção Ambiental", "Proteção e Defesa do Consumidor", "Provimento de informações - exceto órgãos judiciais e de controle", "Provimento de pessoal", "Provimento de subvenções a operações de crédito", "Provimento e Fixação de Profissionais da Saúde", "Provimento e movimentação de pessoal", "Práticas integrativas e complementares", "Publicações oficiais", "Publicidade e divulgação de informativos institucionais", "Publicidade e mídias institucionais", "Qualidade", "Qualidade ambiental", "Qualidade na saúde suplementar", "Qualidade regulatória", "Qualificação dos Equipamentos de Segurança Alimentar e Nutricional (EqSAN)", "Qualificação dos profissionais da saúde indígena", "REALIZAR A COMUNICAÇÃO", "REALIZAR A VIGILÂNCIA E REPRESSÃO", "REALIZAR AÇÕES DE PESQUISA E INVESTIGAÇÃO", "RH - Assessoramento", "RH - EMPREGADOS PÚBLICOS", "RH - Mobilidade - ATPS", "RH - Mobilidade - afastamento para servir em organismo internacional", "RH - Mobilidade - ex territórios", "RH - Monitoramento e Controle", "Realizar fiscalização de terras indígenas", "Realizar pesquisas de satisfação do assistido", "Recepção e manutenção de certidões, registros, declarações e cadastros", "Reconhecimento do Patrimônio Cultural", "Recrutamento e Seleção", "Recrutamento e seleção de pessoal", "Reflorestamento e Recuperação de Áreas Degradadas", "Regimes de Origem", "Regimes para bens de capital", "Registro de empresa de trabalho temporário", "Regulamentação das Comunicações", "Regulamentação e Fiscalização da Atividade do Setor Turístico", "Regulamentos Técnicos e Mobilidade Sustentável", "Regularização do domínio de terras indígenas", "Regulação assistencial", "Regulação da Infraestrutura", "Regulação da Proteção de Dados Pessoais", "Regulação da estrutura de produtos", "Regulação de Mercado", "Regulação do Mercado de Medicamentos", "Regulação do Relacionamento entre Operadoras e Fornecedores", "Regulação do Sistema Único de Assistência Social (SUAS)", "Regulação do setor de mineração", "Regulação do setor elétrico", "Regulação do setor mineral", "Regulação dos Serviços de Transmissão e Distribuição de Energia", "Regulação, licenciamento e fiscalização das atividades espaciais", "Relações Internacionais e diplomáticas", "Representação da União em Tribunais Superiores", "Representação da União na Segunda Instância", "Representação institucional", "Riscos Fiscais", "Saneamento ambiental em áreas indígenas", "Sanidade Vegetal e insumos Agrícolas", "Saúde Bucal", "Saúde da Família e Comunidade", "Saúde da População Negra", "Saúde e Qualidade de Vida do servidor", "Saúde e Segurança do trabalho", "Segurança Alimentar e Nutricional  (SAN) no Sistema Único de Assistência Social (SUAS)", "Segurança Alimentar e Nutricional (SAN) e Clima", "Segurança Alimentar e Nutricional (SAN) nas Cidades", "Segurança Alimentar e Nutricional (SAN) no Sistema Prisional", "Segurança da Aviação Civil contra atos de Interferência ilícita (AVSEC)", "Segurança da informação", "Segurança do Trabalho", "Segurança operacional do setor aéreo", "Segurança predial", "Segurança pública", "Serviços de auditoria", "Setor Audiovisual - Arrecadação de Receitas do Setor Audiovisual - Gestão de Créditos Não Tributários", "Setor Audiovisual - Arrecadação de Receitas do Setor Audiovisual - Gestão de Créditos Tributários", "Setor Audiovisual - Regulação do Setor Audiovisual - Fiscalização das Obrigações Regulatórias", "Setor Audiovisual - Regulação do Setor Audiovisual - Fiscalização das Obrigações Tributárias", "Sistemas de comércio exterior", "Subsidiar a Atuação da Defesa no Contencioso", "Supervisão da aplicação de normas de defesa do consumidor em âmbito nacional", "Supervisão do sistema de controle interno", "Suporte à governança dos recursos", "Suprimento de Sangue e Hemoderivados", "Sustentabilidade das Infraestruturas Hidroviárias", "Sustentabilidade do Setor Portuário", "Sustentabilidade na Navegação Marítima", "TI - Tecnologia da Informação", "TRATAR RISCOS DE CONFORMIDADE TRIBUTÁRIA", "Tecnologia e Pesquisa em Proteção de Dados Pessoais", "Temas Tarifários", "Temas multilaterais", "Tomada de Contas", "Transferência Internacional de Dados (TID)", "Transferência de conhecimento técnico para a sociedade e agentes de ATER", "Transformação digital e soberania tecnológica", "Transparência  ativa e dados abertos", "Transparência ativa e dados abertos", "Transplantes", "Transporte Aquaviário – Desempenho, Estudos e Sustentabilidade", "Transporte Aquaviário – Fiscalização", "Transporte Aquaviário – Licitação e Outorga", "Transporte Aquaviário – Regulação", "Transporte aéreo de carga", "Transporte e LLocomoção", "Transporte e Locomoção", "Transversais", "Uso Sustentável da Biodiversidade e Florestas", "Uso de jurimetria e análise preditiva para aprimoramento da atuação processual", "Vacinação e Imunização", "Vigilância das Doenças Imunopreveníveis", "Vigilância de Violências e Acidentes", "Vigilância de doenças não transmissíveis", "Vigilância em saúde Indígena", "Vigilância – HIV/aids, Tuberculose, Hepatites Virais e IST", "Visitação", "Zonas de Processamento de Exportação"], "servicos": ["ATUAR EM FÓRUNS E GRUPOS DE TRABALHO", "Abertura, acompanhamento e homologação das demandas de Tecnologia da Informação e Comunicação (TIC) e gestão de sistemas implantados.", "Abertura, registro e formalização de Processos de Assistência Jurídica (PAJ)", "Acesso ao sistema Petrvs para registro, preenchimento ou avaliação do plano de trabalho e das entregas mensais (servidor em Programa de Gestão e Desempenho - PGD)", "Acesso à informação em estratégias e", "Acompanhamento  e desenvolvivemento de  trabalhos sobre os Grupos Populacionais Tradicionais Específicos", "Acompanhamento  e desenvolvivemento de trabalhos sobre os Grupos Populacionais Tradicionais Específicos", "Acompanhamento Congresso, Câmara e Senado, plenários e comissões", "Acompanhamento Legislação", "Acompanhamento da Execução", "Acompanhamento da execução de obras hídricas", "Acompanhamento da execução do desenvolvimento físico e orçamentário dos programas, projetos e atividades sob a responsabilidade do Prevfogo", "Acompanhamento da execução orçamentária e financeira das ações de educação ambiental do Ibama", "Acompanhamento da logística de pagamento de benefícios do PBF", "Acompanhamento das Comissões Temáticas de interesse do Ministério", "Acompanhamento das Receitas Orçamentárias", "Acompanhamento das parcerias/ monitoramento dos Termos de Colaboração no Transferegov", "Acompanhamento de atendimento a vítimas de tortura", "Acompanhamento de atos normativos", "Acompanhamento de atos normativos na Rede Suas", "Acompanhamento de ações investigativas", "Acompanhamento de contratos de serviços", "Acompanhamento de instrumentos de legislação orçamentária", "Acompanhamento de instrumentos do planejamento", "Acompanhamento de projetos externos", "Acompanhamento de propostas legislativas sobre Assistência Social", "Acompanhamento do Programa Nacional de Capacitação do Sistema Único de Assistência Social (CapacitaSUAS)", "Acompanhamento do atendimento a demandas dos órgãos externos", "Acompanhamento dos saldos em contas correntes dos entes federados vinculadas ao Fundo Nacional de Assistência Social", "Acompanhamento e avaliação da execução do PDTIC", "Acompanhamento e controle das atividades de comércio exterior de substâncias e produtos químicos e biológicos", "Acompanhamento e controle dos recebimentos pelo Tesouro Nacional relativos aos financiamentos e refinanciamentos concedidos pela União", "Acompanhamento e divulgação de informações sobre haveres financeiros junto a Estados e Municípios", "Acompanhamento e gerenciamento da execução orçamentária e financeira", "Acompanhamento e instrução processual", "Acompanhamento e proposição de demandas para desenvolvimento de funcionalidades para outros sistemas, sob gestão de outros setores ou órgãos externos", "Acompanhamento e proposição de demandas para desenvolvimento de funcionalidades para sistemas", "Acompanhamento e subsídio técnico à defesa da União nas ações judiciais que envolvam os haveres financeiros", "Acompanhamento e suporte às operações de TIC terceirizadas", "Acompanhamento, análise e emissão de posicionamento sobre Projetos de Lei, propostas de Decreto, Portarias, Instruções Normativas, resoluções. dentre outras normativas que abordem temas relactivos so trabalho Infantil, Medidas Socioeducativas em meio aberto e trabalho análogo ao escravo", "Acompanhamentos internos", "Acompanhar e Gerenciar a Execução Orçamentária e Financeira", "Acompanhar e avaliar performance econômico-tributária", "Acompanhar e controlar os programas de financiamento e subvenção", "Acompanhar e divulgar a jurisprudência vinculante em matéria tributária e aduaneira", "Acompanhar e divulgar decisões judiciais relevantes em matéria tributária e aduaneira", "Acompanhar, avaliar e coordenar a atuação dos Defensores Públicos Federais", "Acordo de Cooperação Técnica", "Acordos de Cooperação internacional", "Aditivação ou Renovação de Contratos de TIC", "Administrativo", "Administração Condominial", "Administração da Base de Conhecimento de TIC", "Administração da concessão de garantias em operações de crédito a entes públicos", "Administração da execução orçamentária e financeira da DPF", "Administração da secretaria executiva da comissão de ética", "Administração de Acesso Remoto", "Administração de Almoxarifado", "Administração de Ambientes de Nuvem", "Administração de Bens Apreendidos", "Administração de Contratações de serviços", "Administração de Datacenter", "Administração de Ferramentas de Produtividade", "Administração de Patrimônio", "Administração de Plataformas de Desenvolvimento", "Administração de Rede Local", "Administração de Serviços Gerais", "Administração de Sistemas Operacionais", "Administração de Transporte", "Administração de benefícios do programa Bolsa Família", "Administração de contratações no âmbito de gestão de pessoas", "Administração de e-mails e identidades digitais", "Administração e gestão do Sistema de Informação de Manejo de Fauna (SIMAF) ou ferramenta correlata*", "Administração e gestão dos trabalhos", "Admissibilidades correcionais", "Aferir conformidade", "Agenda Internacional", "Ajustar notas de empenho", "Ajustes do encerramento do exercício financeiro", "Alimenta 1000", "Alimenta Cidades (SAN para PCT)", "Ampliar a disponibilidade e integração das informações para a sociedade sobre o monitoramento da qualidade do ar, da água e do solo", "Ampliação da Estratégia Alimenta Cidades", "Ampliação e desenvolvimento do acervo bibliográfico e documental", "Analisar Conformidade de financiamento e subvenção a programas de Governo acompanhados pela SUGEF", "Analisar Projetos de Lei com a temática de atuação da CGPClin", "Analisar Riscos dos Processos de Trabalho da Unidade", "Analisar atributos de ações orçamentárias", "Analisar crédito disponível", "Analisar e solicitar crédito suplemantar", "Analisar propostas de política tributária e aduaneira", "Análise Política Comercial", "Análise das necessidades de negócio para criação de soluções de Inteligência Artificial", "Análise de Haveres Mobiliários", "Análise de conformidade em licitações e contratos", "Análise de demandas relacionadas ao financiamento público", "Análise de demandas relacionadas à Transformação Digital", "Análise de impacto regulatório", "Análise de merito de demandas por movimentação de pessoas", "Análise de mérito de demandas por novas contratações e provimentos", "Análise de necessidades de negócio para criação de soluções de Inteligência de Negócio", "Análise de solicitações das ações de gestão de pessoas da SESAI", "Análise do financiamento e subvenção a programas de Governo", "Análise e Supervisão Contábil", "Análise e acompanhamento de saldos dos empenhos de diárias e passagens;", "Análise e elaboração de documentos  de Cooperação Técnica", "Análise e gestão de dados fundiários das UCs", "Análise normativa de demandas no âmbito da Proteção Social Especial", "Análise técnica para compartilhamento de dados", "Análise, validação e incorporação institucional de produtos técnicos de consultorias (UNODC)", "Análises Laboratoriais", "Análises de sobreposições territoriais e interesses concorrentes", "Aperfeiçoamento da atuação do IBAMA nos eventos externos de mudança do clima", "Aperfeiçoamento da atuação do Ibama no Plano Nacional de Contingência para Incidentes de Poluição por Óleo em Águas sob Jurisdição Nacional", "Aperfeiçoamento da participação do Ibama no Sistema de Proteção ao Programa Nuclear Brasileiro - Sipron", "Aperfeiçoamento das ações de Emergências Ambientais e Climáticas no IBAMA", "Aperfeiçoar o atendimento aos acidentes ambientais pelo Ibama", "Aplicação de metodologias e instrumentos técnico-profissionais em atendimento socioassistencial", "Apoiar na eliminação de documentos", "Apoiar na localização de documentos arquivados", "Apoiar na padronização do arquivamento de documentos", "Apoiar no recebimento e envio de documentos físicos ao arquivo central", "Apoiar no recebimento e transferência de arquivos/documentos físicos", "Apoiar os estados da região nordeste do Brasil na realização do serviço de acompanhamento familiar para inclusão produtiva rural", "Apoiar os gestores na implantação da gestão de continuidade de negócios", "Apoio Administrativo à Coordenação-Geral PPCAAM", "Apoio Diagnóstico e Terapêutico", "Apoio Jurídico em Atenção à Saúde Integral", "Apoio a elaboração, acompanhamento e avaliação do Plano Plurianual (PPA) da elaboração e avaliação da Lei Orçamentária Anual (LOA)", "Apoio a elaboração, acompanhamento e avaliação do Plano Plurianual (PPA) da elaboração e avaliação da Lei Orçamentária Anual (LOA).", "Apoio a elaboração, acompanhamento e avaliação dos programas e das ações do Planejamento Plurianual (PPA) e da Lei Orçamentária Anual (LOA), respectivamente", "Apoio a elaboração, implementação, monitoramento e avaliação do planejamento estratégico e gestão de projetos.", "Apoio a entes públicos e privados no atendimento a emergência climática", "Apoio administrativo à participação em foros internacionais", "Apoio administrativo às atividades da Comissão de Ética do Ibama, quando solicitado", "Apoio ao Desenvolvimento e Uso de Soluções de TI", "Apoio ao corpo gestor do IBAMA na tomada de decisões", "Apoio ao desenvolvimento da estratégia de cibersegurança junto ao Gestor de Segurança da Informação", "Apoio ao desenvolvimento e ao acompanhamento de indicadores ambientais", "Apoio ao fortalecimento da Rede Brasileira de Bancos de Alimentos (RBBA) e equipamentos de segurança alimentar nos territórios", "Apoio ao fortalecimento da Rede Brasileira de Bancos de Alimentos e equipamentos de segurança alimentar nos territórios", "Apoio ao fortalecimento das ações voltadas à perda e desperdício de alimentos", "Apoio ao fortalecimento das ações voltadas à perda e desperdício de alimentos.", "Apoio ao fortalecimento dos equipamentos de segurança alimentar e nutricional (EqSAN)", "Apoio ao gerenciamento da DPF", "Apoio aos órgãos de controle", "Apoio as ações de melhoria das condições de funcionamento da instituição, compreendidas as condições de caráter organizacional.", "Apoio e instrução dos debates e na tomada de decisões em relação ao uso sustentável da flora e dos recursos florestais", "Apoio e orientação à gestão descentralizada e participativa do SUAS", "Apoio jurídico de ações judiciais", "Apoio logístico à fiscalização ambiental", "Apoio na Elaboração do Banco de Projetos de Conversão de Multas", "Apoio na Gestão do Banco de Projetos de Conversão de Multas", "Apoio na digitalização de documentos", "Apoio na expedição de documentos", "Apoio na identificação, análise e tratamento de riscos nos processos e sistemas finalísticos", "Apoio na otimização da instrução processual no SEI", "Apoio nos procedimentos de doação de patrimônios", "Apoio técnico aos entes da Federação - Atendimento às demandas da STN, de órgãos de controle e de outros órgãos federais", "Apoio técnico aos entes da Federação - Promoção de capacitação, treinamentos e palestras", "Apoio técnico aos entes federados com temáticas relativas à Gestão do SUAS", "Apoio técnico aos entes federados com temáticas relativas à Proteção Social Especial", "Apoio técnico aos entes federados sobre Proteção Social Básica", "Apoio técnico aos entes federados sobre a gestão orçamentária e financeira do SUAS", "Apoio técnico aos entes federados sobre temáticas relativas à política de Assistência Social", "Apoio técnico e administrativo", "Apoio técnico e orientação sobre o Programa Bolsa Família", "Apoio técnico em demandas de direitos humanos e moradia", "Apoio técnico sobre custos", "Apoio técnico sobre temáticas relacionadas às entidades e organizações da sociedade civil de assistência social", "Apoio técnico à coordenação", "Apoio à Gestão de Bens e Serviços", "Apoio à Governança de Dados", "Apoio à Prestação de Contas Anual", "Apoio à celebração de acordos de cessão de uso de infraestrutura espacial pública", "Apoio à celebração de instrumentos de repasse", "Apoio à educação, pesquisa e ações formativas", "Apoio à emissão das Autorizações Diretas", "Apoio à gestão de aquisições e contratações da Unidade", "Apoio à gestão de continuidade de negócios", "Apoio à gestão de riscos", "Apoio à gestão do PEB", "Apoio à gestão dos processos administrativos", "Apoio à internacionalização de empresas e à promoção de negócios internacionais", "Apoio à mecanização de pequeno porte para agricultores familiares e povos e comunidades tradicionais", "Apoio, fomento e fiscalização da pesquisa, desenvolvimento e inovação nos institutos e universidades", "Aprimoramento da atuação do Ibama na gestão de riscos", "Aprimoramento da gestão da conformidade", "Aprimoramento da gestão das operações Ship to Ship (transferências de petróleo e derivados entre embarcações), incluindo sua normatização, autorização e fiscalização", "Aprimoramento da prestação de serviços ao público", "Aprimoramento dos instrumentos de regulação do SUAS", "Aprimoramento e fortalecimento do desenvolvimento técnico continuado dos servidores", "Aprimorar / ampliar as estratégias de monitoramento e avaliação das pesquisas", "Aprimorar metodogias", "Aprovação e publicação do PDTIC", "Apuração de dados e comunicação", "Apuração de infração a direito do consumidor", "Apuração de infrações ambientais", "Apuração e análise técnica dos processos tramitados à coordenação", "Apuração e consolidação dos processos tramitados à coordenação", "Articulação", "Articulação Institucional", "Articulação com Empresas licitadas para atendimento às demandas das cestas de alimentos", "Articulação de elaboração de materiais socioeducativos com as áreas finalísticas do Ibama e apoiando a produção destes nas Superintendências nos estados e nas Diretorias no Ibama Sede", "Articulação de políticas públicas para inclusão social dos povos e comunidades tradicionais", "Articulação e Cooperação Institucional em TICs", "Articulação e Relações Governamentais", "Articulação e encaminhamento à rede socioassistencial, de saúde e de proteção", "Articulação e monitoramento de planejamentos temáticos e específicos", "Articulação institucional", "Articulação institucional e apoio técnico em demandas de direitos humanos e moradia", "Articulação institucional em demandas coletivas de direitos humanos e moradia", "Articulação institucional enquanto Departamento de Proteção Social Especial", "Articulação institucional interna em direitos humanos", "Articulação institucional para Acordo de Não Persecução Penal", "Articulação institucional para aplicação do DFT nos órgãos do SIPEC", "Articulação institucional para implementação de ações de atenção primária à saúde l", "Articulação institucional para promoção da política regulatório nacional", "Articulação institucional para provimento de pessoal na APF", "Articulação interinstitucional", "Articulação intersetorial com órgãos governamentais e entes federados para organização da agenda da agricultura urbana e periurbana", "Articulação intragovernamental", "Articulações com demais parceiros estratégicos (sem Acordo de Cooperação Técnica) sobre temática da migração e refúgio", "Articulações em colegiados e grupos sobre temática da migração e refúgio", "Articulações intrassetoriais relativas à gestão de condicionalidades do Programa Bolsa Família", "Assegurar a promoção e a proteção dos direitos humanos", "Assessoramento", "Assessoramento Jurídico em projetos de leis e normativos sobre Atenção à Saúde Integral", "Assessoramento ao gabinete da Seplan", "Assessoramento da ASRCC no Comitê de Programação Financeira – CPF", "Assessoramento das áreas do Ministério da Fazenda na temática da integridade", "Assessoramento e apoio administrativo ao Presidente, Diretores e demais Dirigente do Ibama", "Assessoramento e articulação em assuntos internacionais", "Assessoramento técnico", "Assessoramento técnico à alta administração", "Assessoramento técnico-administrativo", "Assessoramento técnico-administrativo ao Gabinete do(a) Defensor(a) Público(a)-Chefe", "Assessoramento às áreas técnicas da SNDCA", "Assessorar na proposição de ações para a consolidação e o fortalecimento do SUAS", "Assessorar na proposição de ações para a consolidação e o fortalecimento dos instrumentos e instâncias de negociação e pactuação do SUAS", "Assessoria Parlamentar referente à prevenção e promoção da Saúde", "Assessoria Técnica e Resposta a Demandas Governamentais", "Assessoria de Imprensa e Relações Públicas relativo à Gestão da Atenção Primária à\nSaúde", "Assessoria de Imprensa e Relações Públicas sobre Atenção Integral à Saúde", "Assessoria e apoio na elaboração de conteúdos institucionais de ética", "Assessoria técnica aos Estados e municípios", "Assessoria técnica aos Estados e municípios sobre a regulação do SUAS", "Assessoria técnica à presidência do Ibama, ao MMA, à Casa Civil e aos demais Ministérios sobre questões relativas às competências da Diretoria de Proteção Ambiental.", "Assessoria às demandas de pessoal", "Assistência ao Presidente em sua representação política e social", "Assistência ao Presidente no preparo e despacho de seu expediente", "Assistência direta e imediata ao Presidente", "Assistência técnica em gestão documental", "Assistência técnica em perícias judiciais", "Assistência à saúde", "Atender a demandas de acesso à informação  da Rede SUAS", "Atender a demandas de transparência e acesso à informação", "Atender a demandas judiciais de solicitação de auxílio para pescadores atingidos pelo derramamento de petróleo no litoral nordestino em 2019", "Atender demandas da Lei de Acesso à Informação", "Atender diariamente o Grupo dos Estados da Rede SUAS no WhatsApp", "Atender o público interno e externo", "Atendimento", "Atendimento a demandas do cidadão no âmbito da Proteção Social Especial", "Atendimento a demandas institucionais externas", "Atendimento a demandas judiciais", "Atendimento a demandas por informações instutionais", "Atendimento a demandas por serviços contratados pelo órgão", "Atendimento a demandas relacionadas à Auditoria", "Atendimento a ocorrências em situações de emergência e estado de calamidade pública", "Atendimento a usuários sobre documentos e processos", "Atendimento a órgãos de controle e instrumentos de transparência", "Atendimento a órgãos requisitantes sobre a gestão das condicionalidades do Programa Bolsa Família", "Atendimento ao Cidadão", "Atendimento ao cidadão", "Atendimento ao público", "Atendimento ao público externo", "Atendimento ao público referente ao Programa de Aquisição de Alimentos (PAA)", "Atendimento ao público sobre acesso à informação", "Atendimento ao público visitante", "Atendimento ao usuário presencial", "Atendimento de Demandas Externas", "Atendimento de Demandas Funcionais", "Atendimento de Demandas Internas e Externas", "Atendimento de Demandas internas e externas", "Atendimento de Demandas referentes às Estratégias e Políticas de Saúde Comunitária", "Atendimento de demandas constantes em processos", "Atendimento de demandas de Órgãos de Controle", "Atendimento de demandas internas e externas", "Atendimento de demandas relativas ao controle administrativo, judicial e legislativo.", "Atendimento de pesquisas bibliográficas", "Atendimento de reuniões e pautas da Câmara de Comércio Exterior", "Atendimento dos Serviços Ibama", "Atendimento e Comunicação Institucional", "Atendimento e Gestão de Chamados (AtendeTI)", "Atendimento e apoio técnico em assistência social no âmbito dos Processos de Assistência Jurídica", "Atendimento e prestação de informações sobre compras e licitações", "Atendimento interno e externo à Administração Pública e sociedade civil", "Atendimento para informações sobre  serviços gerais", "Atendimento social e entrevistas socioeconômicas de assistidos e familiares no âmbito de programas, projetos e benefícios sociais", "Atendimento à Auditoria, demandas judiciais e administrativas", "Atendimento à demanda diversa", "Atendimento à demandas judiciais (tribunais, MPF, PFE)", "Atendimento à órgãos de controle", "Atendimento às demandas Interinstitucionais", "Atendimento às demandas judiciais", "Atenção especializada à saúde", "Atividades de capacitação e desenvolvimento de pessoal", "Atos de pessoal", "Atualizar e melhorar o Pesquisa Saúde;", "Atualizar painés oçamentários", "Atualização constante dos processos e procedimentos operacionais padronizados da área", "Atualização da Política Nacional de Desenvolvimento das Atividades Espaciais (PNDAE)", "Atualização das Agendas", "Atualização de norma ou política de TIC", "Atualização de normas regulatórias", "Atualização do Cadastro Nacional de Unidades de Conservação - CNUC", "Atualização do planejamento das contratações públicas de responsabilidade da unidade", "Atualização e disponibilização de informações nos diferentes canais de comunicação do Ibama", "Atualização e disponibilização de informações referentes aos temas da Cicam nos diferentes canais de comunicação do Ibama", "Atualização e manutenção das rotinas e procedimentos contábeis no SIAFI", "Atualização normativa referente a classificação e gerenciamento de áreas contaminadas", "Atuação administrativa em demandas de saúde", "Atuação como Órgão Central de Administração Financeira", "Atuação como órgão central de administração financeira", "Atuação em mutirões de atendimento", "Atuação em planos nacionais enquanto Proteção Social Especial", "Auditoria e fiscalização de POV", "Auditoria e fiscalização de produtos de origem vegetal", "Autorizar o ingresso em terra indígena para os casos que o requeiram", "Autorização e renovação de fundações de apoio à pesquisa", "Auxílio na execução das ações de reparação por danos ambientais, sob a competência do Ibama", "Auxílio na gestão da frequência dos servidores da unidade.", "Auxílio para pescadores atingidos pelo derramamento de petróleo no litoral nordestino em 2019", "Avaliar desempenho individual", "Avaliar execução físico-financeira do orçamento", "Avaliar projetos de pesquisa, incluindo os do Programa Antártico Brasileiro", "Avaliação  do Programa Cisternas", "Avaliação PLOA - fase qualitativa/quantitativa do PLOA", "Avaliação PLOA -fase qualitativa/quantitativa do PLOA", "Avaliação da resolução de vulnerabilidades cibernéticas", "Avaliação da sustentabilidade fiscal", "Avaliação de Tecnologias em Saúde (ATS) e Diretrizes Clínicas", "Avaliação de cursos de graduação", "Avaliação de desempenho institucional", "Avaliação de impactos ambientais", "Avaliação de imóveis", "Avaliação de maturidade de gestão de riscos", "Avaliação de políticas e programas", "Avaliação de produtos remediadores e dispersantes", "Avaliação de projetos de lei com impactos do setor agropecuário", "Avaliação de resultados e impacto do Programa de Aquisição de alimentos", "Avaliação do risco de extinção de espécies da fauna", "Avaliação dos processos em aberto nas unidades dos contratos do Auxílio Emergencial (AE) e Auxílio Emergencial Residual (AER)", "Avaliação e apoio técnico à gestão e destinação de documentos", "Avaliação e controle", "Avaliação e monitoramento da movimentação de pessoal", "Avaliação e qualificação das políticas de inclusão produtiva rural e de acesso à água para uma atuação mais adequada junto aos terrriórios atendidos e públicos beneficiários", "Avaliação e validação de modelos de dados", "Avaliação físico-financeira", "Ação governamental: Articular e incidir com os demais Ministérios e gestões estaduais de atendimento socioeducativo nas ações e iniciativas da Política Nacional de Atendimento Socioeducativo", "Ação governamental: Elaborar de diretrizes nacionais para o atendimento socioeducativo", "Ação governamental: Financiamento de equipamentos para a melhoria da infraestrutura nos espaços de atendimento socioeducativo", "Ação governamental: Fortalecer a Política Nacional de Formação Continuada: ENS", "Ações Socioambientais e Consolidação Territorial em Unidades de Conservação", "Ações de Acessibilidade e Inclusão", "Ações de apoio técnico referente à implatação e ao monitoramento das Comissões Intersetoriais do Programa Bolsa Família", "Ações voltadas para o aprimoramento contínuo das áreas.", "COMPARTILHAMENTO DE IMÓVEIS", "CONTRATAÇÃO DE OBRAS E SERVIÇOS DE ENGENHARIA", "CONTROLAR A CONFIGURAÇÃO E PADRONIZAÇÃO DOS SISTEMAS DE CONTROLE PROCESSUAL", "CONTROLE DE MERCADORIAS APREENDIDAS", "Cadastrar/Atualizar tetos orçamentários no SCDP", "Cadastro e registro de  servidores e estagiários", "Cadastros e informações fiscais/tributários", "Capacitar a equipe para utilizar as soluções tecnológicas", "Capacitação da equipe", "Capacitação da sociedade na temática educação ambiental", "Capacitação e Suporte aos Usuários de Ferramentas de Análise e Monitoramento de Efetividade de Gestão das UCs", "Capacitação e Treinamento da Equipe e\nProfissionais", "Capacitação e Treinamento da Equipe e\nProfissionais do Departamento", "Capacitação e Treinamento da Equipe e Profissionais", "Capacitação e desenvolvimento de pessoas", "Capacitação para a gestão do uso dos recursos faunísticos no âmbito do Sisnama", "Capacitação sistemática e continuada dos servidores do IBAMA para a utilização dos sistemas de informação de gestão do comércio exterior da biodiversidade", "Capacitações e eventos sobre custos", "Caracterização de povos e comunidades tradicionais", "Celebração de Projetos com Recursos Externos", "Celebração de Termos de Compromissos com Empreendedores", "Celebração de projetos com recursos de oriundos de TACs", "Centralizar o relacionamento da Instituição com órgãos de controle interno e externo", "Cerimonial e eventos", "Certificação de Entidades Beneficentes de Assistência Social (CEBAS)", "Cesta Básica de Referência", "Chamamentos de processos seletivos internos publicados", "Classificar a arrecadação", "Classificação e destinação de documentos", "Cobrança e gestão da arrecadação dos créditos tributários e não tributários constituídos", "Cofinanciamento federal da Proteção Social Especial", "Cofinanciamento federal da proteção social básica", "Colaboração e condução de ações associadas à fauna silvestre no enfrentamento de Emergências Ambientais e/ou Eventos Climáticos extremos", "Coleta de dados e realização de pesquisas com assistidos", "Coleta e análise de amostra e gestão de resultado", "Compatibilização de direitos em territórios sobrepostos a UCs", "Compensação Previdenciária", "Compensação previdenciária", "Composição da força de trabalho", "Composição estrutura organizacional da secretaria", "Compra pelo comprasnet", "Comunicação Institucional sobre a Gestão do SUAS", "Comunicação Institucional sobre a Proteção Social Básica", "Comunicação Institucional sobre a Proteção Social Especial", "Comunicação Interna", "Comunicação em saúde acerca das doenças sob atribuição do DPNI", "Comunicação institucional interna e externa.", "Comunicação pró-ativa com entes federados", "Concepção e gerenciamento de plataformas para disponibilização e divulgação interna e externa dos dados geográficos ambientais e informações ambientais tratadas ou produzidas pelo Ibama", "Concepção, manutenção e aprimoramento das funcionalidades dos sistemas de informação", "Concessão de Garantia a Entes Públicos", "Concessão de apoio técnico aos órgãos nas práticas contábeis e na elaboração de demonstrativos Fiscais", "Concessão de apoio técnico aos órgãos nas práticas contábeis e na elaboração de demonstrativos fiscais", "Concessão de diárias e passagens", "Concessão de licenças para atividades espaciais", "Conclusão dos processos contratuais do Auxílio Emergencial (AE) e Auxílio Emergencial Residual (AER) sem pendências", "Concorrência", "Condução de agendas que demandam diálogo intersetorial relacionadas aos públicos da Alta Complexidade (Crianças,  Adolescentes e Jovens, Adultos e Famílias, Pessoas Idosas, Pessoas com Deficiência e Mulheres em situação de violência doméstica).", "Condução de ações educativas em saneamento ambiental em áreas indígenas", "Condução dos processos de reparação por danos ambientais do território do Distrito Federal (Supes-DF)", "Condução dos processos relacionados a Cites de sua competência, em conjunto com as Autoridades Científicas", "Condução/ instrução dos processos licitatórios cujos objetos estejam nos planejamentos da coordenação", "Configurar Sistemas", "Conformidade Legal e Segurança Operacional", "Conformidade de registro de gestão", "Conformidade dos agentes operadores de apostas de quota fixa", "Conselhos de Unidades de Conservação", "Conservação do acervo arquivístico", "Consolidação das Contas Públicas da Federação", "Consolidação e Divulgação de Informações para Alocação de Recursos nas UCs", "Consolidação e Divulgação de Informações sobre Compromissos de Ajustamento de Conduta", "Consolidação e Divulgação de Informações sobre Projetos com Recursos Externos", "Consolidação e Divulgação de Informações sobre a Efetividade da Gestão das UCs", "Consolidação e Divulgação de Informações sobre dos Recursos Oriundos de Compensação Ambiental", "Constituir documento de arrecadação", "Constitução, controle e registro de haveres financeiros", "Construção e aprimoramento de normativos", "Consulta e levantamento de informações à gestão", "Consultoria e Assessoramento Jurídico", "Consultoria e Relações com Órgãos de Controle", "Consultoria relacionada à governança de TIC", "Contato com órgãos, entidades e terceiros relacionados à atividade finalística", "Contratação e gestão de consultorias técnicas via cooperação internacional (UNODC)", "Contribuir para o aperfeiçoamento contínuo das políticas públicas por meio do fortalecimento de uma cultura de monitoramento e avaliação  M&A", "Contribuição e fomento a estudos, pesquisas e eventos para as ações de agricultura urbana", "Contribuído com a institucionalização da governança da Integração e Desenvolvimento Sul- Americano (SEAI)", "Controlar créditos ativos da União", "Controlar e fiscalizar processo de exportação", "Controlar e fiscalizar processo de importação", "Controlar fruição de benefícios fiscais e regimes especiais de tributação", "Controlar os Prazos, Respostas e a Garantia de Efetivação da Ciência de Documentos Administrativos", "Controlar rede arrecadadora", "Controlar saldos Finaceiros", "Controle Estratégico dos Haveres Mobiliários", "Controle contábil referente às ações da Atenção Primária à Saúde/ Saúde Bucal", "Controle da Programação Financeira", "Controle da Qualidade e Auditoria em Radioterapia e Mamografia", "Controle da programação financeira", "Controle de Acessos", "Controle de Pagamento de Benefícios", "Controle de acessos e ativos de informação", "Controle de almoxarifado", "Controle de arrecadação atividades culturais e artísticas", "Controle de correspondências", "Controle de custos", "Controle de demandas de Ouvidoria", "Controle de legalidade dos atos administrativos", "Controle de prazos processuais", "Controle de registro de frequência", "Controle do cumprimento de legislação, resoluções, normas", "Controle e conformidade", "Controle em gestão de pessoas", "Cooperação Jurídica Internacional", "Cooperação internacional para promoção da alimentação e nutrição na saúde", "Coordenar Elaboração do Plano de Dados Abertos", "Coordenar o posicionamento institucional quanto a proposições legislativas com impacto fiscal ou afetas às atribuições da Secretaria do Tesouro Nacional", "Coordenar os estudos de cultura, memória e identidade", "Coordenação CRAS, PAIF, Equipe Volante, Lanchas da Assistência Social, e Carteira da Pessoa Idosa", "Coordenação Técnica da Câmara Nacional de Diálogo e Negociação Permanente entre o SUAS e o Sistema de Justiça e Garantia de Direitos (SJGD)", "Coordenação da Articulação com Comitê Nacional Interinstitucional de Implementação e Monitoramento da Política Antimanicomial do Poder Judiciário em interface com as Políticas Sociais (CONIMPA)", "Coordenação da Comissão Permanente de Avaliação de Documentos Sigilosos - CPads do Ibama", "Coordenação da Cooperação Técnica no contexto do departamento de prevenção e promoção da saúde", "Coordenação da Escola do Sistema Único de Assistência Social (SUAS) Simone Albuquerque", "Coordenação da Execução das Atividades Relacionadas ao Processo de Planejamento Estratégico do Ibama", "Coordenação da Força de Proteção do Sistema Único de Assistência Social (FORSUAS)", "Coordenação da Política de Planejamento da Força de Trabalho na APF", "Coordenação da atuação do Departamento em eventos de Foros Internacionais em foros internacionais", "Coordenação da elaboração e coordenação do Relatório de Qualidade do Meio Ambiente - RQMA", "Coordenação da estruturação, execução, implementação e monitoramento das ações de Governança no âmbito do Ibama", "Coordenação da execução das atividades relativas à execução contábil, à aplicação de dotações orçamentárias e recursos financeiros", "Coordenação da participação institucional da SNDCA na COP 30 (Belém/PA)", "Coordenação da política de aquisição, controle e manutenção dos acervos bibliográficos do Ibama, colocando-os à disposição do público interno e externo", "Coordenação da prestação de contas", "Coordenação das Atividades de Suporte Técnico e Administrativo à Gestão Ambiental", "Coordenação das atividades de capacitação e desenvolvimento de pessoal", "Coordenação das atividades relativas à execução orçamentária da Diretoria.", "Coordenação das atividades vinculadas às parcerias institucionais", "Coordenação das ações de comunicação e treinamento relacionados à conformidade", "Coordenação das ações de fiscalização ambiental", "Coordenação das ações institucionais em comemoração aos 35 anos do Estatuto da Criança e do Adolescente (ECA)", "Coordenação das ações que compõem o Planabio", "Coordenação das ações relacionadas à promoção da transparência ativa no Ibama", "Coordenação de Estudos, Projetos e Programas sobre Atenção Integral à Saúde", "Coordenação de Políticas e Programas de desenvolvimento tecnológico e inovação", "Coordenação de Projeto Piloto visando a implantação de novas modalidades de acolhimento  para os diversos públicos (Crianças,  Adolescentes e Jovens, Adultos e Famílias, Pessoas Idosas, Jovens e Adultos com Deficiência e Mulheres em situação de violência doméstica).", "Coordenação de ações com fornecedores e terceiros para Garantia de cumprimento de requisitos contratuais de cibersegurança", "Coordenação de ações com os integrantes do Sistema Brasileiro de Inteligência - Sisbin", "Coordenação de ações e prioridades táticas das unidades subordinadas", "Coordenação de infraestrutura tecnológica e suporte aos usuários", "Coordenação de instrumentos para transferência de recursos na condução de parcerias para promoção da agricultura urbana e periurbana", "Coordenação de processos de mudanças organizacional por meio de Comitê de gênero, raça e diversidade", "Coordenação de propostas de normas relativas ao uso sustentável e controle dos recursos da flora", "Coordenação do Grupo de Trabalho Intersetorial para o planejamento de estratégias e ações integradas voltadas à implantação, ampliação e qualificação do Serviço de Acolhimento em Família Acolhedora (GTI-SFA).", "Coordenação do Núcleo Nacional de Educação Permanente do Sistema Único de Assistência Social (NUNEP/SUAS)", "Coordenação do Sistema Nacional de Desenvolvimento das Atividades Espaciais (SINDAE)", "Coordenação do Subcomitê Federal para Acolhimento e Interiorização de Imigrantes em Situação de Vulnerabilidade (SUFAI)", "Coordenação do fornecimento de informações da PCPR", "Coordenação do procedimento de seleção e adoção de missões espaciais - ProSAME", "Coordenação do processo de avaliação continuada dos serviços públicos", "Coordenação do processo de elaboração da proposta orçamentária anual e das solicitações de alterações orçamentárias", "Coordenação do processo de regularização de obrigações oriundas de passivos contingentes reconhecidos", "Coordenação do sistema nacional de prevenção e combate aos incêndios florestais", "Coordenação e Supervisão de Atividades de Inteligência e Contrainteligência de interesse do Ibama", "Coordenação e apoio à execução de projetos e convênios institucionais na área de direitos humanos.", "Coordenação e apoio à execução de projetos, acordos e convênios institucionais", "Coordenação e consolidação da elaboração dos relatórios anuais de atividade e de gestão", "Coordenação e elaboração de manifestações técnicas", "Coordenação e elaboração de normas e procedimentos para auxiliar ações de uso sustentável, por meio do Plano de Manejo Florestal Sustentável - PMFS", "Coordenação e execução de ações de inovação para a sistematização da produção e da disseminação do conhecimento e da informação ambiental", "Coordenação e execução de repasses financeiros da União a entes subnacionais e entidades privadas, com foco em regularidade, controle e resultado.", "Coordenação e fomento a estudos e pesquisas sobre a cultura afro-brasileira", "Coordenação e realização de perícias médicas e sociais", "Coordenação na SNAS quanto à implementação do Plano Nacional para Enfrentamento ao Estado de Coisas nas Prisões Brasileiras (Pena Justa)", "Coordenação na gestão de documentos sigilosos", "Coordenação no âmbito da Proteção Social Especial/SNAS do plano de fortalecimento do SUAS junto aos povos da Terra Indígena Yanomami frente à Emergência", "Coordenação,  junto às unidades descentralizadas, da execução das ações que compõem o Planabio, relativas às competências das suas coordenações de área", "Coordenação, elaboração de normas e procedimentos para auxiliar ações de uso sustentável, por meio do Plano de Manejo Florestal Sustentável - PMFS", "Coordenação, elaboração de normas e procedimentos para auxiliar ações de uso sustentável, por meio do Plano de Manejo Florestal Sustentável - PMFS (concessões floestais federais)", "Coordenação, elaboração de normas e procedimentos para auxiliar ações de uso sustentável, por meio do Plano de Manejo Florestal Sustentável - PMFS Comunitário (Unidades de Conservação Federal)", "Coordenação, elaboração de normas e procedimentos para auxiliar ações de uso sustentável, por meio do Plano de Manejo Florestal Sustentável Comunitário- PMFS (Unidades de Conservação Federal)", "Coordenação, gestão e acompanhamento das propostas de conversão de multas relacionadas a Projetos Institucionais propostos pela Coordenação no âmbito do PCMAI", "Coordenação, instrução processual e execução de análises e procedimentos técnico-administrativos relacionados ao processo de compensação ambiental decorrentes do licenciamento ambiental federal, conforme art. 36 da Lei nº 9.985, de 2000 e Decreto que o regulamenta", "Coordenação, instrução processual e execução de análises e procedimentos técnico-administrativos relacionados ao processo de compensação ambiental decorrentes do licenciamento ambiental federal, conforme art. 36 da Lei nº 9.985, de 2000, e Decreto que o regulamenta", "Coordenação, orientação técnica e acompanhamento de ações de uso, por meio de manejo sustentável da vegetação nativa", "Coordenação, orientação, subsídios e execução das obrigações incumbidas ao Ibama em acordos nacionais e internacionais sobre o comércio e o controle da biodiversidade do qual o País é signatário", "Coordenação, supervisionamento, instrução de proposta de normas, bem como orientação técnica a ações e programas permanentes de monitoramento do uso da fauna, de âmbito regional ou nacional, de forma articulada com outras instituições", "Coordenação, supervisão, instrução processual e execução de análises e procedimentos técnico-administrativos relacionados ao licenciamento ambiental federal de atividades e empreendimentos", "Credenciar agente arrecadador", "Credenciar intervenientes", "Criar apresentações sobre temas estratégicos das áreas de pesquisa clínica e de saúde pública de precisão", "Criação de comunidades por temas ou competências", "Criação e Implementação de um Plano Coordenado para Revisão e Otimização de Gastos Públicos", "Criação e alteração de limites e categorias de UCs", "Criação e estruturação de Câmara Consultiva", "Criação e publicação de conteúdos nos sites e repositórios da BNDigital", "Cruzamento e análise de dados para pagamento de benefícios do PBF", "Cumprimento das exigências legais impostas as entidades da Adminstração Pública", "Cumprimento de sentença judicial de ressocialização", "Curadoria cinematográfica", "Curadoria de exposições do acervo da FUNDAJ", "Curadoria de materiais e documentos para público da biblioteca", "DAR SUPORTE À INTEGRAÇÃO E AO COMPARTILHAMENTO DE DADOS DE INTERESSE TRIBUTÁRIO", "DEFINIR OS MAIORES CONTRIBUINTES DE INTERESSE FISCAL", "DESTINAÇÃO DE MERCADORIAS APREENDIDAS", "Decidir sobre Direito Creditório", "Definição, implementação e sustentação da arquitetura e plataforma para disponibilização de dados, produtos ou serviços que suportem o desenvolvimento de soluções analíticas de dados", "Demandas Administrativas de TI", "Demandas sociais de participação e diversidade; atendimento a movimentos sociais, organização de eventos com sociedade civil", "Demarcação e sinalização do perímetro das UCs", "Desapropriação de imóveis e indenização de benfeitorias em UCs", "Desemvolvimento de ações investigativas", "Desenho de Programas, Projetos e Instrumentos Técnicos", "Desenvolver comunicação institucional", "Desenvolver pessoas", "Desenvolver processos de gestão orçamentária", "Desenvolvimento de aplicações e sistemas corporativos", "Desenvolvimento de atividades em cooperação SUPES e outras instituições, sob demanda", "Desenvolvimento de ações de promoção e difusão cultural", "Desenvolvimento de estudos, pesquisas e inovação envolvendo geotecnologias, com ênfase em monitoramento ambiental", "Desenvolvimento de novas soluções", "Desenvolvimento de política afeta a sociobioeconomia", "Desenvolvimento de produto técnico ou tecnológico", "Desenvolvimento de programas de inovação", "Desenvolvimento de projetos no âmbito da gestão do Programa Bolsa Família", "Desenvolvimento de rotas tecnológicas", "Desenvolvimento de sistemas corporativos", "Desenvolvimento de sistemas de comércio exterior", "Desenvolvimento de soluções de TI internamente", "Desenvolvimento de um Plano Integrado de Melhoria e Eficiência das Políticas Públicas", "Desenvolvimento do acervo (aquisição, doação, compra ou produção e descarte)", "Desenvolvimento e Execução de Capacitações Técnicas", "Desenvolvimento e Gestão de Competências Aeronáuticas", "Desenvolvimento e Manutenção de Solução de TI", "Desenvolvimento e gestão do sistema SISPRO", "Desenvolvimento e implantação de soluções de TI", "Desenvolvimento e monitoramento da estratégia de administração da DPF", "Desenvolvimento e monitoramento do planejamento fiscal de médio e longo prazos", "Deslocamentos oficiais por necessidade do serviço", "Diagnóstico da Efetividade da Gestão das UCs", "Diagnóstico das ações de Agricultura Urbana e Periurbana dentro da Estratégia Nacional de Segurança Alimentar e Nutricional nas Cidades – Alimenta Cidades", "Difusão de conhecimento e orientações sobre a gestão descentralizada do Programa Bolsa Família e do cumprimento de suas condicionalidades", "Difusão de conteúdos originários de pesquisas e estudos.", "Diligência à Gestão Municipal do Cadastro Único e a Órgãos Parceiros", "Dimensionamento da Força de Trabalho", "Diretrizes Estratégicas para Comunicação", "Diretrizes Estratégicas para Comunicações", "Disponibilização de informações ambientais", "Disponibilização de meios aéreos", "Disponibilização de meios aéreos não tripulados", "Disponibilização de publicações institucionais", "Disponibilização e disseminação de informações e acervos", "Disseminar a legislação", "Disseminação da cultura da gestão de continuidade de negócios", "Disseminação da cultura da gestão de integridade", "Disseminação da cultura da gestão de riscos", "Disseminação da informação", "Disseminação da ética pública", "Disseminação do conhecimento em Ciência, Tecnologia e Inovação (CT&I)", "Disseminação do conhecimento, comercializar e controlar os produtos editoriais impressos", "Diversificação da captação de recursos internacionais e nacionais de forma estratégica", "Divulgação de informações contábeis e fiscais da Federação", "Divulgação de informações da área de gestão de pessoas", "Divulgação de informações de custos do Governo Federal", "Divulgação de orientações e informações institucionais", "Divulgação e apoio às ações de capacitação institucional", "Divulgação e popularização da ciência", "Divulgação proativa e espontânea de informações de interesse público no site do Ministério do Desenvolvimento Social, Família e Combate à Fome consoante diretrizes da Lei de Acesso à Informação acerca dos contratos, termos aditivos e termos de apostilamentos de transferência de renda", "EDIÇÃO DE TEXTO/EDITORAÇÃO", "EXECUTAR ATIVIDADES PREPARATÓRIAS, ACESSÓRIAS E COMPLEMENTARES DE INSTRUÇÃO PROCESSUAL", "EXECUTAR PROCEDIMENTOS DE CLASSIFICAÇÃO DE DOCUMENTOS E DE PROCESSOS ADMINISTRATIVOS", "EXECUTAR PROCEDIMENTOS DE RECEPÇÃO DE DOCUMENTOS E PROCESSOS DE EXPEDIENTE", "EXECUTAR PROCEDIMENTOS DE RECEPÇÃO E TRIAGEM DE PROCESSOS ADMINISTRATIVOS ENTRE UNIDADES", "Editoração?", "Edição de atos normativos referentes às condicionalidades do Programa Bolsa Família", "Educação ambiental na gestão pública da biodiversidade", "Efetuar controle de carga", "Efetuar o controle e a fiscalização da remessa expressa", "Efetuar o controle e fiscalização da remessa postal", "Efetuar o controle e fiscalização de bagagem", "Efetuar o controle e fiscalização de movimentação física internacional de valores", "Efetuar previsão de pagamentos", "Elaborar boletim de decisões judiciais", "Elaborar e divulgar estatísticas de finanças públicas em conformidade", "Elaborar estudos e assessorar os dirigentes no tocante a temas econômico-fiscais", "Elaborar propostas de súmula ou resolução do CARF", "Elaborar respostas a demandas institucionais", "Elaboraçao de conteúdos visando garantir o acesso à informação e transparência", "Elaboração da Programação Financeira", "Elaboração da proposta orçamentária", "Elaboração das normas das atividades espaciais", "Elaboração de Carta Acordo", "Elaboração de demostrativo de execução orçamentária e financeira", "Elaboração de diagnóstico dos dados de resíduos sólidos do relatório anual de atividades potencialmente poluidoras RAPP", "Elaboração de diagnósticos, avaliações, modelos, relatórios temáticos e outros produtos vinculados ao processamento dos dados brutos dos sistemas de informação sobre o uso de espécimes, produtos e subprodutos da fauna", "Elaboração de diretrizes para classificação, gerenciamento e remediação de áreas contaminadas", "Elaboração de diretrizes técnicas (manuais, planos, documentos instrutivos)", "Elaboração de documentos e diligências no âmbito da atividade finalística", "Elaboração de documentos e petições judiciais da Unidade", "Elaboração de documentos técnicos para atendimento a manifestação quanto às demandas de órgãos de controle interno e fiscalização e/ou auditoria externo de defesa do Estado", "Elaboração de edital para apoiar a oferta de refeições em cozinhas solidárias.", "Elaboração de estudos estratégicos", "Elaboração de guia de governança de dados", "Elaboração de informes e materiais sobre as condicionalidades do Programa Bolsa Família", "Elaboração de minutas, no âmbito da atividade finalística", "Elaboração de norma regulatória de bancos comerciais", "Elaboração de normas e fluxos para parcerias", "Elaboração de normas regulatórias sobre o mercado de fundos", "Elaboração de normas, procedimentos e instrumentos para projetos de serviços ambientais", "Elaboração de projetos de unidades de saúde indígena e infrestrutura de saneamento (Sistemas de Abastecimento de Água - SAA, Melhoria Sanitária Domiciliar - MSD, Polo Base - PB, Unidade Básica de Saúde Indígena - UBSI e Casa de Saúde Indígena - CASAI)", "Elaboração de subsídio à CGFAU, no tema fauna e biodiversidade aquática, no contexto de conservação associada a acordos e tratados internacionais (CDB e CMS)", "Elaboração de subsídio à CGFAU, no tema fauna, em sua atuação como Autoridade Científica da Cites", "Elaboração de subsídio à CGFlo, no tema flora, em sua atuação como Autoridade Científica da Cites.", "Elaboração de termos de referência e editais de chamamento público", "Elaboração do Mapa de Risco", "Elaboração do PTD (Plano de Transformação Digital)", "Elaboração do Plano Estratégico de Tecnologia da Informação e Comunicação (PETIC)", "Elaboração do Plano de Conformidade", "Elaboração do planejamento de longo prazo", "Elaboração dos Relatórios de Análise de Impacto Regulatório (AIR) e de Análise de Resultado Regulatório (ARR)", "Elaboração dos planos de contratações de TIC e projeções orçamentárias", "Elaboração e Coordenação dos Planos de Gestão Anual, Bianual e Quadrianual", "Elaboração e Implementação de um Plano Estratégico para Otimização de Políticas Públicas", "Elaboração e Monitoramento do Acordo de Gestão com o MMA", "Elaboração e difusão de estudos sobre o setor espacial", "Elaboração e disponibilização das informações de Contabilidade Pública da União", "Elaboração e disponibilização de informações e estatísticas da DPF", "Elaboração e disponibilização dos relatórios das demonstrações contábeis e notas explicativas", "Elaboração e encaminhamento de comunicações institucionais em direitos humanos", "Elaboração e implementação da Rede Nacional de M&A", "Elaboração e implementação do Sistema nacional M&A", "Elaboração e manutenção do Plano de Continuidade de TIC", "Elaboração e manutenção dos instrumentos de planejamento estratégico de TIC e demais planos derivados, e Monitoramento de sua execução e o alcance dos resultados definidos", "Elaboração e proposição de requisitos e especificações técnicas para a importação e exportação de espécies, produtos e subprodutos da biodiversidade", "Elaboração e proposição de requisitos e especificações técnicas para reposição florestal obrigatória, e uso da matéria prima florestal de empreendimentos licenciados", "Elaboração e revisão de normas relativas a avaliação ambiental de agrotóxicos, substâncias químicas e produtos perigosos", "Elaboração e revisão de planos de manejo", "Elaboração ou alteração de atos normativos", "Elaboração subsídios técnicos em acordos judiciais e termos de compromisso a serem firmados no Ibama Sede, referentes a obrigações de recuperação ambiental e reparação por danos ambientais à flora e fauna, observada competência das demais áreas na forma de normativa própria", "Elaboração, análise e validação da documentação normativa institucional", "Elaboração, consolidação, acompanhamento e avaliação dos planos e programas anuais e plurianuais do Ibama", "Elaboração, implementação e monitoramento da política de geoinformação do Ibama", "Elaboração, revisão e consolidação de atos normativos e mapeamento e melhoria de processos organizacionais", "Elaboração, revisão e consolidação de atos normativos, definição de procedimentos e mapeamento e melhoria de processos organizacionais.", "Elucidação de questionamentos sobre patrimônio", "Emissão de declaração de reconhecimento de limites", "Emissão de documentos funcionais", "Emissão de licenças e anuição da importação e exportação de espécies, produtos e subprodutos da biodiversidade", "Emitir empenho, pagamento e Realizar a conformidade.", "Emitir empenho, pagamento e realizar a conformidade.", "Encapsulamento e encadernação de acervo", "Ensino", "Ensino - Gestão Acadêmica", "Ensino e Supervisão Acadêmica", "Escrituração Fiscal Digital EFD-Reinf e Transmissão da DCTFWeb", "Estabelecer e gerenciar o orçamento dos gastos da unidade nas despesas diversas e orçamento destinado ao subsídio das refeições;", "Estabelecimento da gestão estratégica de dados sob gestão do DIQUA", "Estabelecimento de Normas de Contabilidade Pública, e de Elaboração dos Demonstrativos Fiscais para a União.", "Estabelecimento de diretrizes, padrões e normas técnicas de TIC", "Estabelecimento de governança, supervisão e apoio técnico ao uso de soluções com IAGen", "Estabelecimento de mecanismos financeiros de conversão de multas indireta - contratação de banco público", "Estabelecimento de parceria para a execução das ações planejadas para o PEB", "Estabelecimento de parcerias estratégicas para projeção do setor espacial junto a outras instituições", "Estabelecimento de zonas de amortecimento", "Estabelecimento e acompanhamento de convênios, acordos, termos de cooperação e parcerias estratégicas com entidades públicas e privadas.", "Estratégia Alimenta Cidades", "Estratégia Alimenta Cidades RS", "Estratégia de Prevenção da Obesidade", "Estratégia de rastreio de não acompanhados da saúde dos participantes do Bolsa Família", "Estratégia de rastreio de não localizados (Nloc) da educação dos participantes do Programa Bolsa Família", "Estratégia e inovação", "Estratégias de Comunicação com a Juventude de Recife", "Estruturação da gestão administrativa para atuar de forma inovadora, padronizada e transparente", "Estruturação da visitação", "Estruturação do sistema de governança dos ecossistemas estaduais de inovação", "Estruturação e operação do Escritório de projetos da SMA", "Estruturação e padronização da documentação organizacional", "Estudos e pesquisas para a difusão dos acervos", "Estudos e projetos de aperfeiçoamento e inovação institucional", "Evoluções no sistema da APS", "Executar o financiamento e subvenção a programas de Governo", "Executar os Procedimentos de Ciência de Documentos Administrativos", "Executar políticas públicas", "Execução Financeira - DARF Numerado", "Execução ações acessórias aos haveres financeiros da União não relacionados a entes federativos", "Execução da cooperação técnica entre instituições nacionais e internacionais em assuntos relacionados ao manejo integrado do fogo", "Execução das ações planejadas para o  Programa Espacial Brasileiro (PEB)", "Execução de Auditoria Interna", "Execução de Convênios", "Execução de atividades relacionadas à comunicação interna e externa e às relações intra e interinstitucionais, incluindo atendimento de agentes externos e usuários dos serviços públicos prestados", "Execução de atividades relacionadas às relações intra e interinstitucionais, incluindo atendimento de agentes externos e usuários dos serviços públicos prestados.", "Execução de ações de fiscalização ambiental", "Execução e monitoramento do orçamento", "Expansão e Massificação das Comunicações", "Extinção do Mercado Ilegal de Apostas de Quota Fixa", "FISCALIZAÇÃO DE OBRAS E SERVIÇOS DE ENGENHARIA", "Fiscalização ambiental", "Fiscalização das atividades licenciadas", "Fiscalização de Contratos de Atenção Integral à Saúde", "Fiscalização de Contratos de TIC", "Fiscalização de Contratos de TIC para o MAPA, MDA e MPA", "Fiscalização de contratos", "Fiscalização de contratos de TIC", "Fiscalização federal em produtos de origem vegetal que possuem padrão oficial definidos pelo MAPA", "Fomento da iniciativa privada", "Fomento e apoio à formação e capacitação contínua de Agentes de Emergências Ambientais e de agentes de órgãos parceiros", "Fomento à produção nacional de insumos estratégicos", "Fonecimento Técnico de  Subsídio à instância de apoio à governança com base nos dados de risco dos relatórios de Monitoramento", "Formalização de TED para apoiar os estados da região nordeste do Brasil na realização do serviço de acompanhamento familiar para inclusão produtiva rural", "Formalização de Termo de adesão entre Fomento Rural e o Estado do Rio grande do Norte", "Formalização e acompanhamento dos instrumentos de parceria do Programa Cisternas", "Formação continuada e profissional de agentes públicos e sociedade civil", "Formação e desenvolvimento de coleção", "Formação e disseminação de conhecimento em DFT", "Formular Atos Interpretativos", "Formular Atos Normativos", "Formular propostas de política tributária e aduaneira", "Formulação da Política de movimentação de pessoal", "Formulação da estratégia", "Formulação de modelos de prestação de serviços socioassistenciais culturalmente adequados", "Formulação de políticas e programas de desenvolvimento tecnológico e inovação", "Formulação de políticas públicas em Bioeconomia e Ciências Exatas, Humanas e Sociais", "Formulação de políticas públicas ou programas de popularização da ciência e educação científica", "Formulação do modelo de governança e sustentação do CPNU", "Fornecer informações de arrecadação", "Fornecer informações e documentos solicitados por órgãos de controle interno ou externo, como auditorias, SISPNAES, CGU, entre outros.", "Fornecimento de Subsídios e Informações", "Fornecimento de informações, metodologias e tecnologias geoespaciais para apoiar as atividades das áreas finalísticas do Ibama", "Fornecimento de subsídios à CGFau/DBFlo, no tema fauna e biodiversidade aquática, em sua atuação como Autoridade Científica da Cites", "Fornecimento de suporte técnico e metodológico às atividades relativas ao gerenciamento de projetos e produtos de Tecnologia da Informação e Comunicação", "Fornecimento do acesso e disponibilização de informações ambientais e do conhecimento ao público interno e externo", "Fortalecimento da capacidade de resposta e resiliência da STN diante de incidentes", "Fortalecimento das ações de Agricultura Urbana e Periurbana dentro da Estratégia Nacional de Segurança Alimentar e Nutricional nas Cidades – Alimenta Cidades no apoio ao Eixo IV da Estratégia Alimenta Cidades", "Funcionalidades de sistemas desenvolvidos", "FÉRIAS", "GERENCIAMENTO DA ESCRITURAÇÃO CONTÁBIL", "GERENCIAR O ACERVO DE PROCESSOS ADMINISTRATIVOS FISCAIS EM CONTENCIOSO ADMINISTRATIVO", "GERIR A ATIVIDADE NÁUTICA", "GERIR A GOVERNANÇA DE DADOS", "GERIR A UNIDADE DE AVIAÇÃO PÚBLICA", "GERIR ATIVIDADE DE CÃES DE FARO", "GERIR COOPERAÇÃO INTERNACIONAL", "GERIR INTERCÂMBIO INTERNACIONAL DE INFORMAÇÕES", "GERIR OS RECURSOS OPERACIONAIS DE REPRESSÃO ADUANEIRA", "GERIR PROGRAMA DE CONFORMIDADE TRIBUTÁRIA COOPERATIVA", "GERIR RELACIONAMENTO COM O CONGRESSO NACIONAL", "GERIR RELAÇÕES COM ÓRGÃOS DE CONTROLE", "GERIR SOLUÇÕES VOLTADAS À CAPTAÇÃO DE DADOS DE INTERESSE TRIBUTÁRIO", "GESTÃO DA CERTIFICAÇÃO DIGITAL", "GESTÃO DA SEGURANÇA EM INFRAESTRUTURA E SISTEMAS", "GESTÃO DE AQUISIÇÕES E CONTRATAÇÕES DE MATERIAIS E SERVIÇOS", "GESTÃO DE DADOS DE IMÓVEIS", "GESTÃO DE DOCUMENTOS FÍSICOS", "GESTÃO DE MUDANÇAS", "GESTÃO DE OBRAS E SERVIÇOS DE ENGENHARIA", "GESTÃO DO SUPORTE DE TECNOLOGIA", "GESTÃO DOS SERVIÇOS DE TI", "Garantia de conformidade com normas e políticas de segurança da informação", "Garantia de conformidade das ações de TIC com as orientações institucionais", "Garantia de orçamento para aquisição de insumos necessários para controle e prevenção das doenças imunopreveníveis", "Garantia de que as mudanças sejam registradas com plano contendo prazos, recursos, etapas e plano de reversão", "Garantia de que mudanças nos sistemas e processos considerem impactos na segurança da informação", "Garantia de que os serviços críticos de TI possam ser restaurados dentro de tempos aceitáveis após interrupções significativas", "Garantir acesso a sistemas", "Garantir excelência dos programas de financiamento e subvenção", "Garantir representatividade em fóruns consultivos/ deliberativos", "Garantir sustentação a Sistemas", "Gerar e consolidar informações para fins de disponibilização de demonstrativos e relatórios relacionados aos temas de responsabilidade da SUGEF", "Geração da folha de pagamento do Programa Bolsa Família", "Geração, adaptação e difusão de conhecimentos técnicos envolvendo análise de dados e geotecnologia", "Geração, análise, integração e disseminação de forma sistemática os dados e as informações ambientais produzidas, em articulação com as áreas finalísticas do Ibama", "Geração, análise, integração e disseminação de forma sistemática os dados e as informações ambientais produzidas, em articulação com as áreas finalísticas do Ibama;", "Gerenciamento da segurança operacional", "Gerenciamento das atividades relacionadas aos processos de infrações administrativas", "Gerenciamento das licenças/autorizações emitidas", "Gerenciamento de Ambiente de Trabalho", "Gerenciamento de Ativos e Configurações", "Gerenciamento de Backup", "Gerenciamento de Bens e Patrimônio", "Gerenciamento de Custos e Orçamento de TI", "Gerenciamento de Ferramentas de Análise e Monitoramento de Efetividade de Gestão das UCs", "Gerenciamento de Portfólio", "Gerenciamento de Problemas e Incidentes", "Gerenciamento de Processos e Qualidade", "Gerenciamento de Projetos de Sistemas, Dados e BI", "Gerenciamento de Projetos de TI", "Gerenciamento de Riscos", "Gerenciamento de Serviços de Atendimento aos Usuários", "Gerenciamento de Serviços de Banco de Dados", "Gerenciamento de Serviços de Relacionamento com o Cliente", "Gerenciamento de Sistemas Estruturantes", "Gerenciamento de atendimento ao usuário", "Gerenciamento de atividades de organização, tratamento e alimentação da base dados física e digital do acervo bibliográfico e memória institucional do Ibama", "Gerenciamento de atividades de organização, tratamento e alimentação da base de dados relativas aos acervos normativos", "Gerenciamento de bens e patrimônios", "Gerenciamento de demandas do serviço de informação ao cidadão", "Gerenciamento de fluxo de trabalho e padronização de processos", "Gerenciamento de operações de TI", "Gerenciamento de riscos", "Gerenciamento de segurança na nuvem", "Gerenciamento de serviços de administração de dados", "Gerenciamento de serviços de administração de dados para todas as soluções sustentadas pela área de Tecnologia da Informação e Comunicação", "Gerenciamento do ciclo de recursos humanos e administrativos, desde a nomeação e exoneração de pessoal até o planejamento anual de contratações, recursos contratuais, viagens e gestão do sistema SEI.", "Gerenciamento dos processos administrativos e apoio na construção de indicadores ambientais", "Gerenciamento dos serviços espaciais nacionais", "Gerenciamento e divulgação do repositório institucional e banco de imagens do Ibama", "Gerenciar a documentação arquivística", "Gerenciar a sustentação de Sistemas, Dados e BI", "Gerenciar demandas do serviço de informação ao cidadão", "Gerenciar informações cadastrais de pessoal", "Gerenciar o ressarcimento ao SUS​", "Gerenciar operadores econômicos autorizados", "Gerenciar relações institucionais", "Gerenciar riscos de conformidade aduaneira  nas remessas internacionais", "Gerenciar riscos de conformidade aduaneira na exportação", "Gerenciar riscos de conformidade aduaneira na importação", "Gerenciar riscos de conformidade aduaneira nas bagagens acompanhadas", "Gerenciar riscos em TI", "Gerencimento de dados e informações obtidas por meio das atividades de ouvidoria para aprimoramento da prestação de serviços públicos oferecidos pelo Ibama", "Gerir Acordos de Cooperação Técnica", "Gerir Capacitação", "Gerir Demandas Administrativas do Gabinete da SEAID", "Gerir Desempenho Organizacional", "Gerir Dimensionamento da Força de Trabalho", "Gerir Estratégia", "Gerir Processos", "Gerir Programa de Incentivos", "Gerir a Ouvidoria do Tesouro Nacional", "Gerir a comunicação e assessorar os dirigentes quanto à representação", "Gerir a execução", "Gerir a habilitação de operadores no comércio exterior", "Gerir acervo técnico e documental", "Gerir acesso a Sistemas", "Gerir afastamentos", "Gerir alfandegamento de locais e recintos", "Gerir as ações de gestão de continuidade de negócios", "Gerir ações de integridade", "Gerir ações prioritárias", "Gerir bases de dados", "Gerir bens patrimoniais", "Gerir competências", "Gerir comunicações da autoridade nacional de proteção de dados", "Gerir conformidade ao plano de dados abertos RFB", "Gerir contas", "Gerir contratações", "Gerir contratos", "Gerir créditos", "Gerir custos", "Gerir dados", "Gerir demandas", "Gerir demandas internas", "Gerir demandas no SEI", "Gerir desenvolvimento", "Gerir desenvolvimento em cargos/ carreiras", "Gerir desenvolvimento profissional", "Gerir documentos", "Gerir documentos técnicos oficiais", "Gerir débito automático", "Gerir débito on line pré autorizado", "Gerir entregas periódicas", "Gerir estrutura corporativa", "Gerir estrutura organizacional", "Gerir eventos de capacitação", "Gerir execução de contratos", "Gerir execução de conttratos", "Gerir folha", "Gerir foruns consultivos", "Gerir frequência", "Gerir férias", "Gerir garantias", "Gerir informação corporativa", "Gerir instrumentos de parceria e cooperação técnica", "Gerir instrumentos de parceria/ cooperação técnica", "Gerir lançamentos no SIAFI", "Gerir licenças", "Gerir movimentação", "Gerir o Inventário de Dados Pessoais", "Gerir o Programa de Gestão e Desempenho - PGD", "Gerir o Programa de Privacidade e Segurança da Informação (PPSI)", "Gerir o Relatório de Impacto à Proteção de Dados Pessoais", "Gerir o orçamento", "Gerir os avisos de privacidade", "Gerir pagamentos", "Gerir processo seletivo", "Gerir processos", "Gerir processos de remoção", "Gerir regularidade fiscal", "Gerir repositório informacional", "Gerir repositórios informacionais", "Gerir responsabilidades", "Gerir riscos", "Gerir substituições", "Gerir transformação organizacional", "Gerir tributos", "Gerir trilhas de desenvovimentos", "Gestão", "Gestão  da Publicação de conteúdos factuais e/ou noticiosos nas redes do MDHC", "Gestão  da Transmissão de agendas eventos", "Gestão  de processos", "Gestão  do Planejamento de campanhas e projetos especiais", "Gestão  do Relacionamento com públicos estratégicos do Ministério nas redes sociais", "Gestão  estratégica de ações intersetoriais e interfederativas, em parceria com as áreas de educação, saúde e assistência social", "Gestão Acadêmica", "Gestão Administrativa da Unidade.", "Gestão Contratual", "Gestão Contábil e Financeira dos Recursos Oriundos de Compensação Ambiental", "Gestão Documental - necessidade de gestão dos documentos produzidos e movimentados  nas diversas fases de vida dos processos, nos termos da Lei 8.159/1991 e Decreto 10.278/2020", "Gestão Documental - necessidade de gestão dos documentos produzidos e movimentados  nas diversas fases de vida dos processos, nos termos da Lei 8.159/1991 e Decreto 10.278/2021", "Gestão Documental - necessidade de gestão dos documentos produzidos e movimentados  nas diversas fases de vida dos processos, nos termos da Lei 8.159/1991 e Decreto 10.278/2023", "Gestão Documental - necessidade de gestão dos documentos produzidos e movimentados  nas diversas fases de vida dos processos, nos termos da Lei 8.159/1991 e Decreto 10.278/2024", "Gestão Documental - necessidade de gestão dos documentos produzidos e movimentados  nas diversas fases de vida dos processos, nos termos da Lei 8.159/1991 e Decreto 10.278/2025", "Gestão Documental - necessidade de gestão dos documentos produzidos e movimentados  nas diversas fases de vida dos processos, nos termos da Lei 8.159/1991 e Decreto 10.278/2027", "Gestão Documental - necessidade de gestão dos documentos produzidos e movimentados  nas diversas fases de vida dos processos, nos termos da Lei 8.159/1991 e Decreto 10.278/2028", "Gestão Documental - necessidade de gestão dos documentos produzidos e movimentados  nas diversas fases de vida dos processos, nos termos da Lei 8.159/1991 e Decreto 10.278/2029", "Gestão Estratégica de cursos e programas de capacitação", "Gestão Executiva do Gabinete", "Gestão Nacional da Rede Cetas do Ibama", "Gestão Nacional da Reparação de Danos Ambientais", "Gestão Nacional de Bens Apreendidos", "Gestão Nacional do Sancionador Ambiental", "Gestão Operacional Termos de Compromissos com Empreendedores", "Gestão Operacional da Execução dos Recursos de Compensação Ambiental", "Gestão Operacional de Projetos com Recursos Externos", "Gestão Operacional de Projetos com Recursos de Oriundos de TACs", "Gestão Orçamentaria", "Gestão Orçamentária e Financeira", "Gestão Parceria com a Associação Voluntários para o Serviço Internacional - BRASIL (AVSI Brasil)", "Gestão Patrimonial", "Gestão Predial", "Gestão acadêmica", "Gestão adaptativa do programa Bolsa Família", "Gestão administrativa da Unidade", "Gestão administrativa de ações de capacitação", "Gestão administrativa de contratos", "Gestão administrativa de contratos de estagiários", "Gestão administrativa de contratos de residentes", "Gestão administrativa de estagiários e residentes", "Gestão administrativa de folha de ponto", "Gestão administrativa de plantões institucionais", "Gestão administrativa de seleção de pessoal", "Gestão administrativa de voluntários", "Gestão arquivística", "Gestão cadastral funcional de membros, servidores, voluntários e estagiários", "Gestão compartilhada e dos BD espaciais CGTI", "Gestão condominial", "Gestão contábil da SVSA", "Gestão da  Identidade Visual de campanhas e projetos especiais", "Gestão da Articulação Institucional", "Gestão da Autorização de Funcionamento de Organização Estrangeira em Território Brasileiro – OE", "Gestão da Avaliação Ambiental de Agrotóxicos", "Gestão da Avaliação e Registro de Produtos de Controle Ambiental", "Gestão da Capacitação", "Gestão da Capacitação da equipe", "Gestão da Cobertura fotojornalística", "Gestão da Comunicação Interna", "Gestão da Comunicação interna do Tesouro Nacional", "Gestão da Conta Única e participação do Sistema de Pagamentos Brasileiros-SPB", "Gestão da Diagramação de peças", "Gestão da Equipagem de Conselhos Tutelares e Conselhos de Direitos", "Gestão da Equipagem de Conselhos Tutelares e Conselhos de Direitos - EQUIPADH+", "Gestão da Estratégia Nacional de Comércio Exterior no âmbito da Câmara do Comércio Exterior", "Gestão da Estratégia Organizacional", "Gestão da Estrutura Corporativa", "Gestão da Estrutura de Mobilidade no Sistema Único de Assistência Social (MOBSUAS)", "Gestão da Execução de campanhas e projetos especiais", "Gestão da Formalização da Política Nacional de Formação continuada", "Gestão da Formação Nacional em Direitos Humanos para Pessoas, Lideranças Comunitárias, Gestores Públicos e Conselhos de Direitos", "Gestão da Força de Trabalho da STN", "Gestão da Força de Trabalho do Departamento", "Gestão da IDV (Identidade Visual)", "Gestão da Infraestrutura de TIC", "Gestão da Mesa Nacional de Negociação Permanente", "Gestão da Ouvidoria do Tesouro Nacional", "Gestão da Pesquisa Nacional por Amostra de Domicílios Contínua - PNAD", "Gestão da Política Nacional de Alimentação e Nutrição (PNAN)", "Gestão da Política Nacional de Proteção de Crianças e Adolescentes no Ambiente Digital", "Gestão da Política Nacional de Saúde Bucal", "Gestão da Política Pública em Atenção Integral à Saúde da Pessoa Idosa", "Gestão da Política Pública em Atenção Integral à Saúde do Homem", "Gestão da Política de Atenção à Saúde e Segurança do Trabalho", "Gestão da Política de Formação do SGDCA", "Gestão da Política de Pessoal da APF do DGCI", "Gestão da Produção de materiais específicos de atendimento à imprensa", "Gestão da Produção de notícias de caráter factual e de serviço", "Gestão da Produção e publicação de conteúdos relacionados a agendas e eventos", "Gestão da Provocação proativa da imprensa", "Gestão da Qualidade, confiabilidade e segurança do Setor Espacial brasileiro - Qualiespaço", "Gestão da Relação com o Fundo de Compensação Ambiental", "Gestão da Reprodução de notícias na Intranet", "Gestão da Revista da Propriedade Industrial - Patentes, Desenhos Industriais, Contratos e Programas de Computador", "Gestão da aplicação da marca do Governo Federal", "Gestão da atividade correcional", "Gestão da auditoria interna", "Gestão da avaliação de desempenho", "Gestão da avaliação e melhoria dos processos de controle de preservativos de madeira", "Gestão da central nacional de transportes", "Gestão da cobrança e arrecadação de taxas e Auto de Infração", "Gestão da comunicação institucional em direitos humanos", "Gestão da comunicação relativa às condicionalidades do Programa Bolsa Família", "Gestão da concessão de benefícios, aposentadorias e pensões", "Gestão da consultoria sobre as condicionalidades do Programa Bolsa Família", "Gestão da consultoria técnica sobre as condicionalidades do Programa Bolsa Família", "Gestão da educação ambiental em temas de substâncias químicas", "Gestão da educação continuada", "Gestão da equipe e das atividades da Coordenação Geral de Proteção Social Especial de Alta Complexidade", "Gestão da formação do SGD", "Gestão da formação inicial", "Gestão da formulação e implementação do planejamento estratégico e orçamentário", "Gestão da formulação e implementação do planejamento orçamentário da SVSA", "Gestão da força de trabalho da SVSA", "Gestão da frequência, licenças/afastamentos e PGD dos servidores do MPO", "Gestão da identidade visual nos documentos e publicações", "Gestão da informação", "Gestão da informação corporativa", "Gestão da informação e acompanhamento da Rede Socioassistencial do SUAS", "Gestão da informação para o licenciamento ambiental", "Gestão da informação por meio do atendimento às demandas do Serviço de Informação ao Cidadão (SIC/Lei de Acesso à Informação) e da OuviSUS", "Gestão da informação sobre famílias beneficiárias em UCs", "Gestão da integridade", "Gestão da inteligência ambiental", "Gestão da intranet da AEB", "Gestão da jornada de trabalho", "Gestão da pesca artesanal", "Gestão da plataforma AEB escola virtual", "Gestão da política de educação contínua e permanente", "Gestão da política de internacionalização da literatura brasileira", "Gestão da política pública de Conversão de Multas no Ibama", "Gestão da presença digital do Ministério nas redes sociais", "Gestão da progressão funcional", "Gestão da proteção, conservação, manejo e uso sustentável da fauna e da biodiversidade aquática", "Gestão da qualidade e segurança de produtos espaciais", "Gestão da rede de comunicação de dados", "Gestão da rede de órgãos do SIPEC", "Gestão da remuneração", "Gestão da representação do TN em Conselhos Fiscais", "Gestão da segurança e manutenção predial", "Gestão da seleção interna", "Gestão da sustentabilidade", "Gestão da tecnologia da informação", "Gestão da transferência de recursos destinados aos programas e ações tutelados pela SVSA", "Gestão da vigilância de violências e acidentes", "Gestão da vigilância epidemiológica das doenças imunopreveníveis", "Gestão das Atividades de Execução Orçamentária e Financeira da Unidade.", "Gestão das Ações de Determinantes Sociais em Saúde", "Gestão das Ações de Risco", "Gestão das Ações de Risco Gestão do  Planejamento Estratégicos Repetida LINHA 10", "Gestão das Cartas Acordo", "Gestão das Consultorias de Pessoa Jurídica", "Gestão das Demandas Oriundas de Órgãos de Controle e dos Cidadãos.", "Gestão das Licenças, Afastamentos e Concessões aos Servidores", "Gestão das atividades dos servidores da coordenação", "Gestão das atividades dos terceirizados da coordenação", "Gestão das atividades relativas à governança e uso de soluções de IA", "Gestão das ações de comunicação da SVSA", "Gestão das ações de gestão da integridade na STN", "Gestão das ações de gestão de continuidade de negócios", "Gestão das ações de gestão de riscos e  apoio aos gestores na sua implantação", "Gestão das ações de gestão do trabalho e educação permanente do SUAS no âmbito da Estratégia Yanomami", "Gestão das ações e compromissos de interesse nacional, estabelecidos no âmbito de \nacordos, convenções e tratados internacionais (Convenção de Basileia, Convenção de Minamata e Protocolo de Montreal)", "Gestão das ações e compromissos de interesse nacional, estabelecidos no âmbito de acordos, convenções e tratados internacionais", "Gestão das ações relacionadas à Carreira, recrutamento, seleção e movimentação de pessoas", "Gestão das demandas de ouvidoria", "Gestão das metas e indicadores pactuados pela SVSA no planejamento estratégico", "Gestão das obrigações tributárias", "Gestão das solicitações de materiais de almoxrifado", "Gestão de APIs e integração de dados", "Gestão de Acervo", "Gestão de Arquitetura Tecnológica", "Gestão de Arquivo Físico", "Gestão de Assessoria Técnica às Unidades.", "Gestão de Assuntos Internacionais", "Gestão de Atendimento reativo de demandas da imprensa", "Gestão de Ações de Apoio a Comitês e Mecanismos Estaduais de Prevenção e Combate à Tortura", "Gestão de Ações de Apoio ao CNCPT (Comitê Nacinal de Prevenção e Combate à Tortura)", "Gestão de Ações de Apoio ao CNPCT (Comitê Nacional de Prevenção e Combate à Tortura)", "Gestão de Ações de Apoio ao SNCPT (Sistema Nacinal de Prevenção e Combate à Tortura)", "Gestão de Ações de Apoio ao SNPCT (Sistema Nacional de Prevenção e Combate à Tortura)", "Gestão de Capital Humano e Comunicação na TI", "Gestão de Catálogo de serviços e entrega da APF", "Gestão de Conectividade e Redes Corporativas", "Gestão de Conteúdo referente às Estratégias e Políticas de Saúde referente à Atenção Integral à Saúde", "Gestão de Contratos", "Gestão de Créditos Não Tributários", "Gestão de Cursos e Disciplinas", "Gestão de Dados", "Gestão de Demandas", "Gestão de Demandas da Diretoria", "Gestão de Demandas de Cidadania e Controle Social referentes ao Programa Bolsa Família", "Gestão de Demandas dos Órgãos de Controle", "Gestão de Diárias e Passagens", "Gestão de Eventos", "Gestão de Identidade", "Gestão de Informações do IGD-PBF", "Gestão de Infraestrutura em Nuvem", "Gestão de Inovação Tecnológica", "Gestão de Integridade Institucional", "Gestão de Licenças, Afastamentos e Concessões aos Servidores", "Gestão de Pagamentos", "Gestão de Parcerias", "Gestão de Parcerias no âmbito da Proteção Social Especial", "Gestão de Pessoas", "Gestão de Pilotos Remotos", "Gestão de Plataforma de Automação", "Gestão de Políticas Quilombolas", "Gestão de Processos da Unidade", "Gestão de Processos de Apoio Administrativo", "Gestão de Processos no Sistema Eletrônico de Informações", "Gestão de Processos: Política de Governança de Processos; Guia de Gestão de Processos: Mapeamento e redesenho de processos (publicada a metodologia de gestão de processos (DAGE/SE);", "Gestão de Produção de notícias para Intranet", "Gestão de Programas e Projetos Estratégicos", "Gestão de Projetos, TEDs e ACTs celebrados pela COINT", "Gestão de Publicações no site", "Gestão de Recursos de Emendas Parlamentares", "Gestão de Segurança da Informação", "Gestão de Segurança de dados - LGDP", "Gestão de Serviços de Infraestrutura Governamental", "Gestão de Sistemas de Informação", "Gestão de Soluções de Inteligência Artificial", "Gestão de Soluções de TI", "Gestão de Tecnologia da Informação e Comunicação (TIC) no âmbito da STN", "Gestão de Telefonia IP Institucional", "Gestão de Trabalhos de Busca em Cemitérios", "Gestão de Trabalhos de Pesquisas sobre os trabalhos da CEMDP e os grupos atingidos pela ditadura militar", "Gestão de Transparência", "Gestão de Vulnerabilidades", "Gestão de acervos", "Gestão de acessos para interoperabilidade de dados", "Gestão de acordos de cooperação", "Gestão de acordos de leniência", "Gestão de afastamentos para servir em organismo internacional", "Gestão de almoxarifado", "Gestão de alterações contratuais de estágio", "Gestão de amostras laboratoriais", "Gestão de aquisição de insumos para ações voltadas ao saneamento ambiental e serviços de engenharia", "Gestão de aquisição de serviços de saúde Indígena", "Gestão de aquisições", "Gestão de arquivos e documentos da Unidade", "Gestão de atividades", "Gestão de audiência", "Gestão de audiências na área criminal", "Gestão de avaliação e melhoria contínua", "Gestão de ações de vigilância, prevenção e controle de doenças e agravos, de modo complementar ou suplementar", "Gestão de ações investigativas", "Gestão de ações prioritárias", "Gestão de ações relacionadas à movimentação e desempenho de pessoas", "Gestão de ações voltadas à perda e desperdício de alimentos", "Gestão de backups e armazenamento de dados", "Gestão de banco de dados", "Gestão de benefícios nos programas de transferência direta de renda", "Gestão de bens de consumo e permanente", "Gestão de bens e serviços", "Gestão de cadastros de usuários internos", "Gestão de canais de comunicação do programa Bolsa Família", "Gestão de capacitação", "Gestão de cenários em alimentação e nutrição populacional", "Gestão de cenários em saúde", "Gestão de ciclo documental com garantia de cumprimento de prazos", "Gestão de cobranças administrativas", "Gestão de comentários nos canais proprietários do Ministério e redes sociais", "Gestão de comunicação institucional", "Gestão de consultorias no âmbito da Gestão do SUAS", "Gestão de consultorias no âmbito da Proteção Social Básica", "Gestão de consultorias no âmbito da Proteção Social Especial", "Gestão de consultorias relacionadas à Proteção Social Especial de Alta Complexidade", "Gestão de conteúdo para a difusão da literatura infantil e juvenil brasileira", "Gestão de conteúdos", "Gestão de conteúdos de interesse institucional", "Gestão de conteúdos e equipes", "Gestão de conteúdos para canais de comunicação", "Gestão de contratos", "Gestão de contratos - atendimento aos quesitos dos art.s 84 a 115 da Lei 14.133/2021", "Gestão de contratos - atendimento aos quesitos dos art.s 84 a 115 da Lei 14.133/2022", "Gestão de contratos - atendimento aos quesitos dos art.s 84 a 115 da Lei 14.133/2023", "Gestão de contratos - atendimento aos quesitos dos art.s 84 a 115 da Lei 14.133/2024", "Gestão de contratos - atendimento aos quesitos dos art.s 84 a 115 da Lei 14.133/2027", "Gestão de contratos - atendimento aos quesitos dos art.s 84 a 115 da Lei 14.133/2028", "Gestão de contratos - atendimento aos quesitos dos art.s 84 a 115 da Lei 14.133/2032", "Gestão de contratos diversos da SVSA", "Gestão de controles e segurança institucionais", "Gestão de convênios", "Gestão de créditos orçamentários e recursos financeiros", "Gestão de dados da política de atenção à saúde indígena", "Gestão de dados e informações ambientais", "Gestão de dados pertinentes à execução de políticas públicas e convênios da SESAI", "Gestão de demanda judicial", "Gestão de demandas", "Gestão de demandas advindas da Auditoria e demandas judiciais", "Gestão de demandas da Ouvidoria", "Gestão de demandas de Design", "Gestão de demandas de acesso à informação", "Gestão de demandas de atendimento externas e internas", "Gestão de demandas de audiovisual", "Gestão de demandas de governança interna", "Gestão de demandas de pessoal", "Gestão de demandas de reprodução de acervo", "Gestão de demandas de órgãos de controle", "Gestão de demandas do sistema PGC-PCA", "Gestão de demandas dos órgãos de controle", "Gestão de demandas externas e internas relativas aos haveres financeiros junto a Estados e Municípios", "Gestão de demandas institucionais", "Gestão de demandas judiciais", "Gestão de demandas oriundas de órgãos externos", "Gestão de demandas por informação referentes aos serviços da Câmara do Comércio Exterior", "Gestão de demandas por informação sobre comércio exterior", "Gestão de demandas por informação sobre operações de exportações", "Gestão de demandas por informação sobre operações de importações", "Gestão de demandas por informação sobre os serviços do Comex", "Gestão de demandas relacionadas a débitos imputados a servidor", "Gestão de desempenho individual", "Gestão de design e validação", "Gestão de deslocamentos institucionais", "Gestão de direitos e vantagens", "Gestão de diárias e passagens", "Gestão de equipamentos de TIC", "Gestão de equipamentos/ sistemas", "Gestão de estoques e materiais de consumo", "Gestão de estoques e materiais de consumo da área", "Gestão de eventos", "Gestão de eventos contratuais de residentes", "Gestão de eventos institucionais promovidos pelo DAENT", "Gestão de exposições e mostras envolvendo o acervo da FBN", "Gestão de formação continuada dos profissionais do SGD", "Gestão de frequência", "Gestão de informações em canais de comunicação interna", "Gestão de instrumentos de qualidade ambiental: \nCadastro Técnico Federal de Atividades e Instrumentos de Defesa Ambiental - CTF/AIDA; Cadastro Técnico Federal de Atividades Potencialmente Poluidoras ou Utilizadoras dos Recursos Ambientais - CTF/APP;\nCadastro Nacional de Operadores de Resíduos Perigosos - CNORP; Relatório Anual de Atividades Potencialmente Poluidoras ou Utilizadoras de Recursos Ambientais (RAPP) e Certificado de Regularidade (CR)", "Gestão de insumos estratégicos e amostras biológicas aplicados à Vigilância em Saúde e Ambiente", "Gestão de insumos estratégicos para o controle das doenças imunopreveníveis", "Gestão de laboratórios de Conservação", "Gestão de laboratórios de preservação documental", "Gestão de laboratórios de restauração", "Gestão de licenças de softwares", "Gestão de licenças e afastamentos", "Gestão de licenças médicas funcionais", "Gestão de licenças, afastamentos e férias", "Gestão de materiais de consumo", "Gestão de materias de expediente", "Gestão de metadados do acervo documental digitalizado", "Gestão de mostras internas do acervo cultural", "Gestão de movimentação", "Gestão de movimentação de pessoal", "Gestão de movimentações de ATPS", "Gestão de movimentações de empregados públicos", "Gestão de movimentações de servidores de ex-territórios", "Gestão de movimentações e retorno de anistiados", "Gestão de negócios florestais", "Gestão de normativos", "Gestão de pagamentos do programa Bolsa Família", "Gestão de paineis de BI", "Gestão de painéis de BI", "Gestão de parcerias", "Gestão de parcerias para gestão de processos de contencioso administrativo fiscal", "Gestão de periódicos", "Gestão de pesquisas de índice de preços", "Gestão de políticas de Ciência, Tecnologia e Inovação (CT&I) para tecnologias digitais", "Gestão de políticas e normas de segurança", "Gestão de políticas intersetoriais relacionadas a propriedade Intelectual", "Gestão de políticas públicas em saúde", "Gestão de portfólio de projetos", "Gestão de processos", "Gestão de processos PAJs no sistema SIS-DPU", "Gestão de processos SEI", "Gestão de processos administrativos", "Gestão de processos administrativos de viagens", "Gestão de processos administrativos referentes à saúde da pessoa idosa", "Gestão de processos aministrativos no SEI", "Gestão de processos da unidade", "Gestão de processos de material permanente", "Gestão de processos estratégicos", "Gestão de processos no SEI", "Gestão de produtos de TI no âmbito da STN", "Gestão de programas de pesquisa e fomento para estudo do acervo da BN", "Gestão de programas e projetos de pesquisa", "Gestão de projetos", "Gestão de projetos de parceria para qualificação dos profissinais da Saúde indígena", "Gestão de projetos de recuperação de documentos históricos e obras literárias nacionais no exterior", "Gestão de projetos e processos estratégicos no âmbito da CGGE", "Gestão de projetos e processos estratégicos no âmbito da Diretoria", "Gestão de projetos e programas de educação espacial", "Gestão de propostas de normativas legais e infralegais que abordem temas relativos aos diversos Serviços de Acolhimento coordenados pela Coordenação Geral de Proteção Social Especial de Alta Complexidade", "Gestão de protocolo, registro processual e distribuição de documentos", "Gestão de prêmios literários da BN", "Gestão de publicações oficiais", "Gestão de reagentes, padrões e soluções", "Gestão de recursos do Programa Mais Médicos para o Brasil", "Gestão de relações institucionais e federativas", "Gestão de rendimentos do patrimônio indígena", "Gestão de respostas às solicitações", "Gestão de retenção e recolhimento de tributos", "Gestão de riscos", "Gestão de riscos e compliance", "Gestão de riscos fiscais de médio e longo prazos", "Gestão de riscos na preservação do acervo bibliográfico e documental", "Gestão de riscos na preservação do acervo documental", "Gestão de serviços ambientais", "Gestão de serviços de conectividade e redes corporativas", "Gestão de serviços de interoperabilidade", "Gestão de serviços de software contratados", "Gestão de serviços gerais e segurança institucional", "Gestão de sistema de preservação digital", "Gestão de sistema de recebimento de informações contábeis e fiscais do Setor Público Brasileiro", "Gestão de sistemas da informação da inteligência ambiental", "Gestão de sistemas de comércio exterior", "Gestão de sistemas e bases informacionais", "Gestão de sistemas informatizados no âmbito da Proteção Social Básica", "Gestão de soluções de automação", "Gestão de testes e controle de qualidade", "Gestão de transportes", "Gestão de tutoriais e material instrucional", "Gestão de uniformes e EPI", "Gestão de vigência de contratos de voluntários", "Gestão descentralizada do programa Bolsa Família", "Gestão do Banco de Projetos com Recursos Externos", "Gestão do Bloco de Assinaturas", "Gestão do Cadastro Funcional", "Gestão do Cadastro Funcional\nDesenvolvimento e", "Gestão do Cadastro Nacional de Entidades de Assistência Social (CNEAS)", "Gestão do Colegiado do Departamento", "Gestão do Contrato de Locação de Veículos Utilitários", "Gestão do Desempenho Individual", "Gestão do Monitoramento", "Gestão do Monitoramento da implementação do SIPIA", "Gestão do Núcleo de Inovação Tecnológica", "Gestão do Orçamento da SGP", "Gestão do Planejamento e alocação do quadro técnico", "Gestão do Planejamento estratégico", "Gestão do Plano Pena Justa", "Gestão do Plano Setorial de Agricultura de Baixa Emissão de Carbono/                                                       Promoção da adoção de tecnologias e práticas (SPSABC) voltadas simultaneamente à Mitigação de emissão de carbono, e à Adaptação da produção agropecuária aos efeitos da mudança do clima", "Gestão do Programa Envelhecer nos Territórios", "Gestão do Programa Envelhecer nos territórios", "Gestão do Programa Farmácia Popular do Brasil", "Gestão do Programa de Alimentação Escolar", "Gestão do Programa de Aquisição de Alimentos (PAA)", "Gestão do Programa de Aquisição de Alimentos (PAA)\n\nElaboração, revisão e publicação de diretrizes e atos normativos para as ações de promoção da alimentação saudável\nNormatização das ações de promoção da alimentação saudável", "Gestão do Programa de Fortalecimento da Capacidade Institucional para a Gestão em Regulação - PRO-REG", "Gestão do Programa de Gestão e Desempenho - PGD", "Gestão do Projeto MPO de CORPO E ALMA", "Gestão do SISDIP", "Gestão do Sancionador Ambiental", "Gestão do Sistema\nIntegrado de Comércio Exterior - SISCOMEX.", "Gestão do Sistema Acolhedor", "Gestão do Sistema de Cadastro das Instituições de Uso Científico de Animais (CIUCA)", "Gestão do Sistema de Informação do Programa de Aquisição de Alimentos (SISPAA)", "Gestão do Viva mais Cidadania", "Gestão do acervo arquivístico e bibliográfico sob custódia da FUNDAJ", "Gestão do acompanhamento e monitoramento das execuções diretas do Programa", "Gestão do atendimento ao público e serviços de recepção", "Gestão do atendimento inicial ao público", "Gestão do banco de dados da APS", "Gestão do cadastro e registro de pessoal do departamento", "Gestão do ciclo de vida de plataformas e ferramentas de produtividade", "Gestão do conhecimento", "Gestão do conhecimento em serviços de TIC", "Gestão do conhecimento sobre as ofertas do Sistema Único de Assistência Social (SUAS) em emergências", "Gestão do contrato de Normas ABNT E ISO, aquisição de normas técnicas para os usuários internos do ibama", "Gestão do contrato de publicações, além de acompanhamento da execução desse serviço no âmbito do Ibama Sede", "Gestão do cumprimento das decisões judiciais, decisões administrativas e diligências relacionadas a matéria de pessoal", "Gestão do cumprimento de determinação judicial de convocação de servior para participar de audiência judicial", "Gestão do desenvolvimento servidores do MPO (PDP, capacitação, mestrado)", "Gestão do espaço físico", "Gestão do estágio probatório", "Gestão do novo plano de Integridade: Integra+MDHC - 2024-2025", "Gestão do novo plano de integridade: Integra+MDHC - 2024-2025", "Gestão do pagamento do cofinanciamento federal da Proteção Social Básica", "Gestão do portal da AEB", "Gestão do portifólio de ativos de TIC", "Gestão do provimento, da alocação e da vacância", "Gestão do recrutamento externo", "Gestão do relacionamento com assistidos", "Gestão do relacionamento com assistidos e órgãos externos", "Gestão do sistema de prêmios literários", "Gestão do site do sistemas de comércio exterior Siscomex", "Gestão do site institucional e do portal Tesouro Transparente", "Gestão do sites institucionais de Recuperação ambiental", "Gestão documental", "Gestão documental dos PAJs", "Gestão documental e fiscalização de TEDs, convênios, termos de fomento, emendas de bancada", "Gestão dos Contratos de Bens, Serviços e Insumos de Saúde Indígena", "Gestão dos Honorários de Sucumbência", "Gestão dos Instrumentos Jurídicos de Transferência de Recursos e demais parcerias", "Gestão dos Planos de Risco", "Gestão dos Programas Nacionais de Controle da Poluição do Ar por Veículos Automotores - Proconve e por Motociclos e Veículos Similares - Promot e a utilização do Selo Ruído do Programa Nacional de Educação e Controle da Poluição Sonora", "Gestão dos Projetos de Cooperação Técnica Internacional", "Gestão dos Vídeos", "Gestão dos bens patrimoniais da unidade", "Gestão dos ciclos de planejamento e monitoramento estratégico da SVSA", "Gestão dos dados Cadastrais de Pessoal", "Gestão dos dados Financeiros de Pessoal", "Gestão dos dados disponibilizados da Infraestrutura Nacional de Dados Abertos - INDA e da Infraestrutura Nacional de Dados Espaciais - INDE", "Gestão dos dados pessoais e funcionais", "Gestão dos planos estratégicos do departamento de prevenção e promoção da saúde", "Gestão dos procedimentos administrativos e judiciais para cobrança e recuperação de créditos públicos, inclusive em parceria com órgãos de controle.", "Gestão dos processos de trabalho", "Gestão dos sistemas de comércio Exterior", "Gestão dos sistemas de informação referente às Estratégias e Políticas de Saúde Comunitá", "Gestão e Acompanhamento dos Instrumentos de Parceria Relativo ao PPCAAMM", "Gestão e Coordenação de Atividades Administrativas e Operacionais", "Gestão e Fiscalização de Contratos", "Gestão e Fiscalização de Contratos de extração de dados orçamentários", "Gestão e Monitoramento de Fundos", "Gestão e Resposta a Demandas Externas Relacionadas à Qualidade Ambiental", "Gestão e Suporte ao Sistema SIPIA PPCAAM", "Gestão e Suporte de Sistemas de Informação em Saúde", "Gestão e Tratamento de Dados", "Gestão e acompanhamento de compras de materiais", "Gestão e acompanhamento dos instrumentos de parceria e cooperação técnica", "Gestão e controle de bens patrimoniais permanentes", "Gestão e disseminação da documentação institucional", "Gestão e disseminação de normas, manuais e procedimentos", "Gestão e fiscalização de contratos administrativos", "Gestão e manutenção do acervo", "Gestão e participação em órgãos colegiados", "Gestão e tratamento de dados", "Gestão estratégica", "Gestão estratégica da Diretoria de Proteção Ambiental", "Gestão estratégica de  ações intersetoriais e interfederativas, em parceria com as áreas de educação, saúde e assistência social", "Gestão estratégica de pessoas e das relações de trabalho na Diretoria", "Gestão executiva da política de combate à tortura", "Gestão executiva da política de erradicação do trabalho escravo", "Gestão federal do Serviço de Proteção em situações de Calamidades Públicas e de Emergências", "Gestão financeira, contábil e de custos", "Gestão geral das ações pactuadas no âmbito de Acordo de Cooperação Técnica (ACT) no âmbito da Operação Acolhida", "Gestão integrada dos acervos bibliográfico, e museológico", "Gestão interna da Coordenação", "Gestão local de projetos e articulação institucional", "Gestão normativa e procedimental sobre prevenção, detecção precoce, análise de risco, análise de rotas de vetores e de dispersores, e controle das espécies exóticas invasoras", "Gestão orçamentária das ações de atenção primária à saúde", "Gestão orçamentária e financeira de repasses", "Gestão patrimonial da unidade", "Gestão qualificada da informação técnica frente às demandas recebidas via SEI, CITSMART, FalaBR, Ouvidoria, Lei de Acesso a Informação (LAI), e-mail institucional, whatsapp e telefone.", "Gestão técnico-normativa sobre prevenção, detecção precoce, análise de risco, análise de rotas de vetores e de dispersores, e controle das espécies exóticas invasoras", "Gestão viatura", "Gestão, Acompanhamento e Avaliação dos Instrumentos de Parceria relativos ao PPCAAM", "Gestão, controle e recebimento dos haveres financeiros da União não relacionados a entes federativos", "Gestão, desenvolvimento, adaptação e difusão de tecnologias envolvendo banco de dados geográficos ambientais e geotecnologias para apoiar as atividades finalísticas", "Governança da Infraestrutura e Serviços Digitais", "Governança de Dados", "Governança de TIC", "Governança e gestão de TI", "Governança e gestão normativa da informação", "Governança: Elaboração do Regimento Interno,  Relatório de Gestão e Relatórios Trimestrais de Monitoramento; Página transparência e prestação de contas;", "Guarda, preservação e recuperação dos acervos", "IDENTIFICAR RISCOS DE CONFORMIDADE", "INATIVAR - ITP", "INATIVAR - incompreensível", "INATIVAR - parte do fluxo", "Identificar entendimentos divergentes no contencioso administrativo", "Identificação e classificação dos ativos de informação para garantir sua proteção", "Impacto das Mudanças no Consumo Alimentar sobre a Mortalidade", "Implantação de inovações em processos, serviços ou produtos", "Implantação do projeto Educação para toda vida", "Implantação do projeto Pronatec -  Saberes da Vida", "Implantação do projeto Vida Digna em Casa", "Implantação do projeto Vida Digna em casa", "Implementar ações de projetos e melhoria de processos", "Implementar os serviços farmacêuticos nos estados no âmbito do CESAF", "Implementação da Norma Operacional Básica de Recursos Humanos do Sistema Único de Assistência Social (NOB-RH/SUAS, 2006)", "Implementação da Vigilância Laboratorial", "Implementação da plataforma do Observatório do Planejamento", "Implementação da política e o plano de gestão do conhecimento e da informação ambiental", "Implementação de Programa de Gestão de Desempenho", "Implementação de Projetos de Fortalecimento Institucional", "Implementação de ações Disque 100", "Implementação de ações para o aumento da força de trabalho", "Implementação de mecanismos de participação social.", "Implementação de plano de resposta a incidentes cibernéticos", "Implementação de política pública em Práticas Integrativas e Complementares em Saúde", "Implementação de políticas públicas em na área de Atenção à Saúde da Criança, do Adolescente e dos Jovens", "Implementação do DFT em órgãos com processo de provimento / autorização de novas vagas em curso", "Implementação do Decreto nº 11.821/2023", "Implementação do Guia de Gestão e Governança Institucional", "Implementação do Programa de Desenvolvimento Integrado para o Centro Espacial de Alcântara (PDI-CEA) por meio de parcerias", "Implementação e difusão da Plataforma Integradora de Dados e Análise", "Implementação e execução do Programa de Brigadas Federais", "Implementação, integração e gerenciamento de sistemas de informação para a gestão do comércio exterior da biodiversidade", "Implementação, integração, gerenciamento, auditoria e monitoramento de sistemas de informação de monitoramento e controle do uso da fauna", "Importação de bens adquiridos ou doados às entidades ou organizações da sociedade civil de assistência social", "Inclusão social e produtiva", "Incorporação não onerosa de imóveis inseridos em UCs", "Inexigibilidade ou dispensa de licitação", "Iniciativas de incentivo ao desenvolvimento de sistemas de custos de órgãos e entidades", "Inquérito CadÚnico", "Instruir processo no sei de afastamento de servidor para participação presencial de servidores em eventos internacionais", "Instrumentalizar a execução", "Instrumentos de Planejamento", "Instrumentos de conectividade e gestão integrada", "Instrução de proposta para normas, requisitos e procedimentos para conclusão dos projetos de plantio florestal incentivados", "Instrução de proposta para normas, requisitos e procedimentos para manejo sustentável comunitário em Unidades de Conservação de Uso Sustentável - IBAMA, ICMBio e SFB.", "Instrução de proposta, implementação e monitoramento do Plano Nacional de Gestão da Educação Ambiental - Pangea, em conjunto com as Superintendências nos estados e as Diretorias no Ibama Sede, visando o fortalecimento da gestão ambiental pública", "Instrução de propostas de normas, orientação técnica e acompanhamento e execução da emissão de autorizações, anuências ou licenças para o uso sustentável da flora, coordenação, elaboração de normas e procedimentos para auxiliar as ações de uso sustentável, por meio do Plano de Manejo Florestal Sustentável - PMFS", "Instrução de propostas de normas, orientação técnica e acompanhamento e execução da emissão de autorizações, anuências ou licenças para o uso sustentável da flora, coordenação, elaboração de normas e procedimentos para auxiliar as ações de uso sustentável, por meio do Plano de Manejo Florestal Sustentável - PMFS.", "Instrução de propostas de normas, orientação técnica, acompanhamento e execução de programas e ações relativas ao uso sustentável da flora", "Instrução de propostas de normas, orientação técnica, acompanhamento e execução programas e ações relativas ao uso sustentável da flora", "Instrução de propostas de normas, orientação técnica, acompanhamento e execução programas e ações relativas ao uso sustentável da flora;", "Instrução de propostas de normas, orientação, acompanhamento, elaboração e execução de programas e ações relativas a espécies não Cites.", "Instrução de propostas de normas, orientação, acompanhamento, elaboração e execução de programas e ações relativas à implementação da Cites", "Instrução de propostas de normas, orientação, acompanhamento, elaboração e execução de programas e ações relativas à implementação da Cites e de espécies não Cites.", "Instrução de propostas de normas, orientação, acompanhamento, elaboração e execução de programas e ações relativas à implementação da Cites.", "Instrução de propostas de normas, padrões, metodologias e processos de reparação por  dano ambiental e recuperação ambiental", "Instrução e acompanhamento de processos cuja condução do licenciamento ambiental tenha sido delegada pelo Ibama a outro ente da federação", "Instrução propostas de normas, padrões, metodologias e processos de reparação pelo dano ambiental e recuperação ambiental", "Integração entre PGD e DFT", "Integridade e transparência", "Inventário e controle do acervo documental bibliográfico da BEC", "Investigar irregularidades funcionais", "JULGAR IMPUGNAÇÕES E MANIFESTAÇÕES DE INCONFORMIDADE E RECURSOS", "JULGAR RECURSOS HIERÁRQUICOS EM MATÉRIA TRIBUTÁRIA E ADUANEIRA", "Julgamento correcional", "Julgamento das impugnações contra o lançamento PPA", "Julgamento dos Recursos contra o lançamento do PPA", "Julgamento em segunda instância de recursos apresentado  contra a cobrança da TCFA", "LOCAÇÃO DE IMÓVEIS", "Laboratório de Inovação em Políticas Públicas para Sistemas Alimentares Saudáveis e Sustentáveis (AlimentaLAB)", "Lançamento tributário das taxas decorrentes do poder de polícia do Ibama", "Legislação e Normas", "Lei Geral de Proteção de Dados Pessoais (LGPD)", "Levantamento das informações e avaliação do cumprimento das metas de gestão e uso da flora", "Levantamento das informações e avaliação do cumprimento das metas de recuperação ambiental", "Levantamento de habilidades e competências necessárias pelos setores que atuam com recuperação ambiental no IBAMA", "Levantamento de indicadores", "Levantamento e busca de obras desaparecidas", "Licitação por pregão Sistema de Registro de Preços (SRP)", "Logística Patrimonial - alteração de Layout das áreas do MPO para melhor ocupação de espaço físico, em atendimento as necessidades das áreas cumprindo o contido da Portaria Conjunta nº 38, de 31 de julho de 2021", "Logística patrimonial - Atendimento ao DECRETO Nº 9.373/2018 sobre a obrigatoriedade de controle dos bens patrimoniáveis", "MACROPROCESSOS MS", "MANUTENÇÃO E GESTÃO DE IMÓVEIS", "MONITORAR AÇÕES DE COOPERAÇÃO", "Manejo e uso da sociobiodiversidade", "Manejo florestal comunitário em UCs", "Manejo integrado do fogo", "Manejo, controle e erradicação de espécies exóticas invasoras", "Manifestação para licenciamento ambiental", "Manter atualizadas relações de recursos extraordinários com repercussão geral e de recursos especiais repetitivos", "Manter registros funcionais", "Manutenção / sustentação de sistemas corporativos", "Manutenção da implementação de controles de segurança aos recursos físicos e lógicos para proteção dos dados e comunicação", "Manutenção das ferramentas de gestão", "Manutenção de Catálogo de Procedimentos Operacionais", "Manutenção de aeronaves", "Manutenção de catálogo de dados disponíveis", "Manutenção de catálogo de serviços de TI atualizado e disponível para os usuários, garantindo que os serviços tenham informações claras, procedimentos de atendimento, níveis de suporte, prazos e responsáveis", "Manutenção de escala de tripulação de aeronaves", "Manutenção de normas e padrões relativos aos processos de desenvolvimento interno e descentralizado de soluções", "Manutenção do portfólio de projetos de Tecnologia da Informação e Comunicação e provisão de informações relacionadas à execução, aos riscos, aos custos, aos marcos e aos benefícios desses projetos", "Manutenção e ajustes na parceria para execução das ações planejadas para o PEB", "Manutenção e evolução de sistemas informatizados de cobrança administrativa", "Manutenção predial", "Mapeamento / levantamento de inovações aplicáveis aos processos da saúde bucal", "Mapeamento das Ações de Agricultura Urbana e Periurbana em Serviços de Saúde e Assistência Social: produção de alimentos em espaços comunitários ou institucionais", "Mapeamento das Ações de Agricultura Urbana e Periurbana em unidades escolares: promoção de hortas pedagógicas em espaços educacionais", "Mapeamento de Ações de Agricultura Urbana e Periurbana em Saúde e Assistência Social: elaboração de edital de chamamento público para a produção de E-book", "Mapeamento de central de custos", "Mapeamento de competências", "Mapeamento de processos", "Mapeamento dos serviços prestados  com padronização de procedimentos e melhoria nos fluxos de trabalho de trabalho", "Mapeamento e articulação de parceiros", "Mapeamento e definição de produtos de dados", "Mapeamento e melhoria contínua dos processos relacionados às condicionalidades do Programa Bolsa Família", "Mapeamento e qualificação dos Equipamentos de Segurança Alimentar e Nutricional (EqSAN) para uma ação integrada com outras politicas de SAN nos terrítorios", "Marco de Referência de Sistemas Alimentares e Clima", "Marco de Referência sobre Sistemas Alimentares e Clima para Políticas Públicas", "Materiais institucionais (manuais, cartilhas e documentos orientadores) analisados, adequados e validados conforme metodologia do Programa", "Matriz Diagnóstica de Sistemas Alimentares Indígenas", "Maturidade de Ouvidoria", "Mediação cultural e artística", "Mediação de conflitos", "Mediação e suporte ao desenvolvimento e uso de soluções de TI", "Melhoramento do controle de acidentes em empreendimentos licenciados", "Melhorar os mecanismos de transparência das pesquisas financiadas;", "Melhoria de processos internos relativos à conversão de multas", "Mensurar gastos tributários", "Missões para experimentos em Microgravidade", "Monitoramento contínuo de sistemas, redes e serviços", "Monitoramento da Cesta Básica", "Monitoramento da Política", "Monitoramento da dimensão Estratégica do PPA", "Monitoramento da execução da estratégia e do desempenho organizacional", "Monitoramento da execução dos serviços, programas, projetos e benefícios de Proteção Social Básica", "Monitoramento da execução dos serviços, programas, projetos e benefícios de Proteção Social Especial", "Monitoramento da iniciativas de promoção da  Ciência, Tecnologia e Inovação (CT&I)", "Monitoramento da oferta dos Serviços de Acolhimento.", "Monitoramento da operação dos serviços e soluções de Tecnologia da Informação e Comunicação para identificação de falhas ou degradações de desempenho", "Monitoramento da pauta das sessões do TCU", "Monitoramento da poluição por hidrocarboneto das águas marinhas de jurisdição nacional", "Monitoramento da visitação", "Monitoramento de Serviços de TI", "Monitoramento de acordos internacionais e termos de colaboração no âmbito da Proteção Social Especial", "Monitoramento de alertas de desastres", "Monitoramento de benefícios do Programa Bolsa Família", "Monitoramento de dados de APIs para carga em banco de dados", "Monitoramento de integridade na gestão de benefícios do Programa Bolsa Família", "Monitoramento do acompanhamento de condicionalidades em municípios com fluxo migratório e promover medidas para a redução dos impactos da aplicação dos efeitos do não cumprimento de condicionalidades para as famílias beneficiárias", "Monitoramento do acompanhamento de condicionalidades em municípios em situação de calamidades e e emergência e promover medidas para a redução dos impactos da aplicação dos efeitos do não cumprimento de condicionalidades para as famílias beneficiárias", "Monitoramento do atendimento às recomendações da CGU", "Monitoramento do cumprimento do calendário de obrigações", "Monitoramento do mapeamento e gerência dos processos", "Monitoramento dos Estados, Municípios e Distrito Federal nas temáticas relativas à Gestão do SUAS", "Monitoramento dos resultados da gestão de riscos e proposição de medidas para seu aperfeiçoamento", "Monitoramento dos serviços e programas de Proteção Social Básica", "Monitoramento e Avaliação da Entregas Institucionais", "Monitoramento e Avaliação de Resultados de Projetos com Recursos Externos", "Monitoramento e Avaliação de Resultados de Projetos com Recursos de Oriundos de TACs", "Monitoramento e Avaliação de Resultados dos Recursos Oriundos de Compensação Ambiental", "Monitoramento e Avaliação do Planejamento Integrado de Recursos das UCs", "Monitoramento e avaliação das atividades espaciais", "Monitoramento e avaliação de indicadores da saude indigena", "Monitoramento e avaliação de indicadores de governança e gestão de TIC", "Monitoramento e avaliação do Programa de Aquisição de alimentos", "Monitoramento e avaliação do desempenho institucional", "Monitoramento e controle das informações de pessoal do departamento", "Monitoramento e estudos de crianças na primeira infância beneficiárias do Programa Bolsa Famíia em relação ao cumprimento das condicionalidades e superação do ciclo integeracional da pobreza", "Monitoramento e estudos jovens beneficiários do Programa Bolsa Famíia em relação ao cumprimento das condicionalidades e superação do ciclo integeracional da pobreza", "Monitoramento e orientação dos agentes do Sistema de Contabilidade Federal", "Monitoramento e Acompanhamentoda Proteção", "Monitoramento geoespacial da proteção ambiental", "Monitoramento institucional do PEI do Departamento de Proteção Social Especial", "Monitoramento, Resposta e Avaliação de Emergências Ambientais", "Monitorar Plano de Dados Abertos vigente", "Monitorar a gestão de pessoas", "Monitorar a vida funcional do servidor", "Monitorar a vida funcional do servidor\nGerir os atos de pessoal e os atos de gestão de pessoas", "Monitorar e Apoiar o Processo Orçamentário", "Movimentação de pessoal", "NEGOCIAR ACORDOS E TRATADOS INTERNACIONAIS", "Normalização Bibliográfica", "Normas contábeis para a Federação - Coordenação da Câmara Técnica de Normas Contábeis e Demonstrativos Fiscais da Federação (CTCONF)", "Normas contábeis para a Federação - Elaboração de instruções de procedimentos contábeis (IPC)", "Normas contábeis para a Federação - Gestão do Manual de Contabilidade aplicada ao Setor Público (MCASP)", "Normas contábeis para a Federação - Gestão do Manual de Demonstrativos Fiscais (MDF)", "Normas contábeis para a Federação - Gestão dos classificadores orçamentários e gerenciais", "Normas contábeis para a Federação - Harmonização das normas de contabilidade do setor público aos padrões internacionais.", "Normas contábeis para a Federação - Projeto - Acordo de Cooperação STN/IRB/Atricon nº 30/2023", "Normas contábeis para a Federação - Projeto - Acordo de Cooperação Técnica –STN/Ministério da Saúde", "Normas contábeis para a Federação - Projeto - Criação do Assistente IA para responder às ouvidorias.", "Normas contábeis para a Federação - Projeto - Definição do Relatório Financeiro Consolidado de Sustentabilidade para o Setor Público Brasileiro", "Normas contábeis para a Federação - Projeto - Definição dos Impactos da Reforma Tributária", "Normas contábeis para a Federação - Projeto - Definições sobre Transferências da União", "Normatização", "Normatização do contencioso administrativo fiscal", "Normatização e Efetivação das Políticas de Comunicações", "Normatização e Efetividade das Politicas de Comunicações", "Normatização e monitoramento relacionados às normas contábeis", "Normatização sobre Custos", "Notificação via Edital", "Novos dados na base de processos de cobrança de ressarcimento", "ORGANIZAÇÃO DE/PARTICIPAÇÃO EM EVENTOS LITERÁRIOS", "Oferta de atividades EAD", "Oferta de capacitaçãoes em TIC", "Oferta de programas de formação em políticas de conservação e proteção do meio ambiente", "Oferta de subsídio técnico para a execução qualificada dos Serviços da Proteção Social Especial de Alta Complexidade", "Operacionaizar pagamentos", "Operacionalizar registros diversos", "Operacionalização do Programa Bolsa Família, Auxílio Gás, Fomento e Legado Cadastro Único", "Operacionalização do Sistema de Concessão de Diárias e Passagens (SCDP)", "Operação e atualização do observatório do setor espacial", "Organização de agenda", "Organização de documentos para publicação e/ou revisão de diretrizes, atos normativos e criação de leis e decretos para as ações de promoção da agricultura urbana e periurbana", "Organização de eventos de promoção cultural", "Organização do processo de residência médica", "Organização do trabalho da unidade por meio da gestão de processos SEI", "Organização documental  para prestação de contas", "Organização e Desenvolvimento Documental do Departamento de Prevenção e Promoção à Saúde", "Organização e Desenvolvimento Documental referente à Vigilância em Saúde e Ambiente", "Organização e Desenvolvimento Documental referente às Estratégias e\nPolíticas de Saúde Comunitária", "Organização e Gestão de Eventos em Atenção Integral à Saúde", "Organização e Gestão de Eventos sobre Promoção de Equidade e Determinantes Sociais em Saúde", "Organização e controle de acervo documental", "Organização, contribuição na realização e participação de simulados de acidentes ambientais para aperfeiçoamento da preparação e resposta", "Orientação científica", "Orientação das unidades delegatárias quanto as atribuições referentes à Autoridade Administrativa Cites", "Orientação das unidades do Ibama sobre a elaboração de atos normativos em consonância com o Manual de Redação da Presidência da República e demais legislações vigentes", "Orientação de Cumprimento de Decisão Judicial", "Orientação e melhoria nas interpretações Técnicas e normativas", "Orientação para encaminhamento de informação oficial", "Orientação sobre estágio obrigatório", "Orientação técnica sobre ações de recuperação ambiental e reparação por dano ambiental", "Orientaçãoes ao público interessado nas áreas técnica e laborátórios da FBN", "Orientaçãoes ao público interessado no acervo da BN", "Orientações diversas sobre controle de frequência do servidor", "Outras demandas", "PARTICIPAR DE FÓRUNS E GRUPOS DE TRABALHO INTERNACIONAIS", "PLANEJAMENTO DA APLICAÇÃO INTERNA DO ORÇAMENTO", "PLANEJAMENTO EDITORIAL", "PRESTAR ATENDIMENTO E ORIENTAÇÃO", "PRODUÇÃO EDITORIAL", "Padronização de documentos", "Padronização e Normatização de Procedimentos Técnicos", "Pagamento de substituição", "Parceria com instituições e órgãos para realização de estudos, pesquisas e desenvolvimento de tecnologias sociais sobre temas relacionados à oferta dos serviços de acolhimento para os diversos públicos (Crianças,  Adolescentes e Jovens, Adultos e Famílias, Pessoas Idosas, Jovens e Adultos com Deficiência e Mulheres em situação de violência doméstica).", "Parcerias Estratégicas e Institucionais", "Parcerias para viabilizar a execução das Políticas Públicas a cargo da Diretoria", "Participação da gestão do Comitê Interfederativo (CIF)", "Participação das atividades referentes à Convenção de Biodiversidade - CDB e outros fóruns internacionais de regulação do uso e conservação da biodiversidade", "Participação em Grupos de Trabalho", "Participação em Grupos de trabalho e comitês", "Participação em agendas e atividades institucionais em direitos humanos", "Participação em atividades acadêmico-administrativas", "Participação em eventos", "Participação em grupos de trabalho nas materias de sua competência", "Participação em grupos de trabalhos, comitês, colegiados, sindicâncias e comissões e representacao institucional", "Participação em grupos de trabalhos, comitês, colegiados, sindicâncias e comissões e representacao institucional, encaminhamento de demandas", "Participação em grupos de trabalhos, comitês, colegiados, sindicâncias e comissões referentes ao curso de vida", "Participação em grupos de trabalhos, comitês, colegiados, sindicâncias, comissões e representacao institucional sobre cadeias produtivas e mudanças climáticas", "Participação em instâncias colegiadas", "Participação em outras instâncias de governança do setor espacial", "Participação em processos de tomada de decisão colegiada (na Câmara de Regulação do Mercado de Medicamentos)", "Participação institucional em Grupos de Trabalho Interministeriais", "Participação institucional enquanto Departamento de Proteção Social", "Participação técnica da Coordenação Nacional em estudos de caso relacionados à proteção.", "Participação técnica em instâncias colegiadas e grupos de trabalho relacionados a Povos e Comunidades Tradicionais", "Particpação do secretário ou representante em eventos internacionais de interesse do governo, a fim de apoiar na discussão da garantia do direito ao acesso e qualidade dos serviços da atenção primária à saúde", "Passivos referentes a organismos internacionais dos quais Brasil faça parte quitados.", "Pedidos de Compras e Contratações", "Pesquisa Nacional sobre Soberania e Segurança Alimentar e Nutricional dos Povos Indígenas no Brasil", "Pesquisa de Segurança Alimentar e Nutricional no Sistema Prisional", "Pesquisa e gestão da informação sobre biodiversidade", "Pesquisa jurisprudencial", "Pesquisa para fundamentar decisão", "Pesquisa sobre estratégias para promoção da alimentação saudável", "Pesquisa, Desenvolvimento e Inovação", "Planajamento de Curso de Formação em Agricultura Urbana e Periurbana em Serviços de Saúde e Assistência Social para servidores e usuários", "Planejamento", "Planejamento Integrado de Recursos das UCs", "Planejamento da Contratação - Fase preparatória - atendimento ao arts. 18 e 19  da Lei 14.133/2023", "Planejamento da Contratação - Fase preparatória - atendimento ao arts. 18 e 19  da Lei 14.133/2030", "Planejamento da Contratação - Fase preparatória - atendimento ao arts. 18 e 19  da Lei 14.133/2031", "Planejamento da Execução dos  Recursos Oriundos de Compensação Ambiental", "Planejamento da Execução dos  Recursos Oriundos de Recursos Externos", "Planejamento da gestão de tecnologia da informação", "Planejamento da visitação", "Planejamento das ações que compõem o Planabio", "Planejamento de Ações de Redução de Impactos sobre a Biodiversidade (PRIM)", "Planejamento de Ações para Conservação de Espécies Ameaçadas de Extinção (PAN)", "Planejamento de Contratação de TIC", "Planejamento de Contratação de TIC para o MAPA, MDA e MPA", "Planejamento de Curso de Formação em Agricultura Urbana e Periurbana em Serviços de Saúde e Assistência Social para servidores e usuários", "Planejamento de Instrumentos de Gestão", "Planejamento de Mudanças e Liberações", "Planejamento de Níveis de Serviço", "Planejamento de contratação de TIC", "Planejamento de longo, médio e curto prazos das atividades espaciais (PNAE, PPA, Lei Orçamentária Anual, respectivamente)", "Planejamento do Orçamento do departamento de prevenção e promoção da saúde", "Planejamento do orçamento", "Planejamento e Conformidade de Contratações", "Planejamento e Coordenação de Contratações", "Planejamento e Execução de Operações Aéreas", "Planejamento e Gestão Orçamentária do Centro", "Planejamento e Gestão da Auditoria", "Planejamento e Orçamento", "Planejamento e controle do orçamento anual das atividades do Prevfogo, conforme PLOA e diretrizes do Pnapa", "Planejamento e coordenação das atividades da Unidade", "Planejamento e coordenação do Programa Quelônios da Amazônia (PQA)", "Planejamento e desenvolvimento de políticas nacionais e inciativas de CT&I em tecnologias digitais", "Planejamento e execução de ações de articulação institucionais", "Planejamento e execução de ações de gestão documental e arquivística", "Planejamento e execução de operações logísticas", "Planejamento e gestão Estratégica", "Planejamento e implementação de  programas, projetos e ações educativas no contexto das atividades finalísticas, visando ao fortalecimento da gestão ambiental pública", "Planejamento e organização do trabalho", "Planejamento e relatoria anual de auditoria", "Planejamento estratégico", "Planejamento estratégico do Departamento de Proteção Social Especial no PPA", "Planejamento operacional", "Planejamento, Gestão e Monitoramento de instrumentos de repasse", "Planejamento, coordenação e Acompanhamento dos haveres financeiros da União junto às administrações direta e indireta de Estados, do Distrito Federal e dos Municípios", "Planejamento, coordenação e acompanhamento a operacionalização dos pagamentos de compromissos internos de responsabilidade do Tesouro Nacional", "Planejamento, coordenação e supervisão da execução das atividades relacionadas ao Sistema de Organização e Inovação Institucional do Governo Federal - Siorg", "Planejamento, coordenação e supervisão da execução das metas globas e intermediárias da Instituição", "Planejamento, coordenação e supervisão das atividades de comunicação social, relações institucionais e ainda a publicação, a divulgação e o acompanhamento das matérias de interesse do Ibama", "Planejamento, desenvolvimento e execução do manejo integrado do fogo (MIF)", "Planejamento, elaboração e avaliação de recursos de divulgação institucional no contexto do manejo integrado do fogo (MIF)", "Planejamento, execução e avaliação da formação de gestores estaduais de programas em fase de implementação", "Planejamento, execução e avaliação da formação técnica das equipes estaduais do PPCAAM", "Planejamento, execução e sistematização do Encontro anual de Coordenadores Gerais e Técnicos do PPCAAM", "Planejar Projetos de TI", "Planejar e especificar sistemas para ampliação das modalidades de pagamento", "Plano Anual de Acompanhamanto  e Fiscalização do  Programa de Aquisção de Alimentos (PAA)", "Plano Anual de Acompanhamanto  e Fiscalizção do  Programa de Aquisção de Alimentos (PAA)", "Plano Anual de Acompanhamento e Fiscalização do Programa de Aquisição de Alimentos (PAA)", "Planos Estratégicos: Elaboração, revisão e monitoramento do PEI; Monitoramento do PPA;", "Política de inovação tecnológica", "Pregão Eletrônico", "Preparação e Planejamento Estratégico para Revisão e Otimização de Políticas Públicas", "Preparo físico e arquivamento de obras", "Preservação do acervo bibliográfico e documental", "Preservação do acervo museológico sob custódia da FUNDAJ", "Preservação e Difusão de acervo cinematográfico", "Prestar contas", "Prestar subsídios à Procuradoria Federal", "Prestação de apoio técnico aos órgãos federais nas práticas contábeis e na elaboração de Demonstrativos Fiscais", "Prestação de apoio técnico nas ações de controle, monitoramento e fiscalização de queimadas irregulares e incêndios florestais", "Prestação de assessoria técnica à implantação e ao monitoramento das Comissões Intersetoriais do Programa Bolsa Família", "Prestação de contas", "Prestação de contas à Auditoria sobre a elaboração do Relatório de Gestão", "Prestação de serviço de administração de pessoal", "Prestação de suporte na avaliação e na atualização das comissões de gestão dos contratos do Auxílio Emergencial (AE) e Auxílio Emergencial Residual (AER)", "Prevenir ilícitos funcionais", "Prevenção de ilícito administrativo", "Prevenção e Resposta às Emergências Climáticas e Epizootias", "Prevenção e apuração de desvios éticos", "Procedimentos de Reanálise Ambiental de Agrotóxicos", "Procedimentos de contabilidade de custos em soluções tecnológicas de apoio à geração de informações de custos", "Procedimentos para conclusão dos projetos de plantio florestal incentivados", "Processamento e Preservação", "Processamento técnico bibliográfico", "Processamento técnico de obras do acervo captadas por Projeto de captação de obras (DDL), doação e intercâmbio", "Processamento técnico de obras do acervo captadas por depósito legal, doção, permuta e compra.", "Processamento técnico de periódicos, acervo especial (Cartografia, iconografia, manuscritos, obras raras e obras musicias) e acervo geral retrospectivo", "Processar mensalmente os pagamentos dos auxílios financeiros concedidos aos estudantes, conforme critérios estabelecidos nos editais vigentes.", "Processo administrativo disciplinar", "Processo administrativo disciplinar\nAdmissibilidades correcionais", "Processo administrativo disciplinar\nProcesso de responsablização de entes privados PAR", "Processo administrativo disciplinar\nProcesso de responsablização de entes privados PAR\nAdmissibilidades correcionais", "Processo administrativo disciplinar\nProcesso de responsablização de entes privados PAR\nJulgamento correcional\nAdmissibilidades correcionais", "Processo administrativo disciplinar/ Processo de responsablização de entes privados PAR", "Processo de pagamento das integralizações de cotas, recomposições a fundos internacionais e das contribuições  a organismos internacionais coordenado e planejado.", "Processo de responsablização de entes privados PAR", "Processo de responsablização de entes privados PAR\nJulgamento Correcional", "Processos de microfilmagem e Digitalização do acervo", "Produzir documentação técnica oficial", "Produzir informação oficial", "Produzir, analisar e dar transparência às projeções da despesa pública", "Produção de Conhecimento", "Produção de Conteúdos Estratégicos e\nNormativos sobre Estratégias e Políticas de Saúde Comunitária", "Produção de Conteúdos Estratégicos e\nNormativos sobre Saúde Bucal", "Produção de Conteúdos Estratégicos e\nNormativos sobre Saúde da Família e Comunidade", "Produção de Conteúdos Estratégicos e Normativo para prevenção e promoção da saúde", "Produção de Conteúdos Estratégicos e Normativos sobre Atenção Integral à Saúde", "Produção de Conteúdos Estratégicos e Normativos sobre Atenção à Saúde das Crianças, Adolescentes e Jovens", "Produção de Conteúdos Estratégicos e Normativos sobre Estratégias e Políticas de Saúde Comunitária", "Produção de Conteúdos Estratégicos e Normativos sobre Prevenção e Promoção da Saúde", "Produção de Conteúdos Estratégicos e Normativos sobre Saúde da População Negra", "Produção de Documentação Técnica referente à Atenção Integral à Saúde", "Produção de Materiais de Comunicação\nDigital, Publicidade e Identidade Visual em Estratégias e Políticas de Saúde Comunitária", "Produção de Materiais de Comunicação Digital, Publicidade e Identidade Visual para o site institucional na temática do curso da vida", "Produção de conhecimento de inteligência", "Produção de conteúdo sobre provimento de pessoal na APF", "Produção de documentos técnico-administrativos", "Produção de documentos técnico-administrativos para o Bolsa Família", "Produção de estudos e análises técnicas", "Produção de pesquisas e estudos socioeconômicos para subsidiar a assistência jurídica", "Produção de peças jurídicas em direitos humanos", "Produção de relatório de inteligência financeira", "Produção de relatórios e pareceres socioeconômicos para subsidiar a concessão da assistência jurídica", "Produção e difusão de notícias e informativos institucionais", "Produção e disseminação de informações fundiárias de terras indígenas", "Produção e disseminação do conhecimento", "Produção e divulgação de conteúdo informativo institucional", "Produção e divulgação de conteúdo informativo institucional da CGU", "Produção e edição de materiais técnicos sobre as condicionalidades do Programa Bolsa Família", "Produção ou atualização de normas temáticas, conforme lacunas e demandas", "Produção, análise e disseminação de dados e informações", "Produção, análise e disseminação de dados e informaçõessobre registro civil", "Produção, processamento, análise e quantificação de dados geoespaciais e informações ambientais", "Programa Nacional de Monitoramento da Biodiversidade – Programa Monitora", "Programa Nacional de SAN no Sistema Prisional (PN-SANSisP)", "Programa de Gestão e Melhoria da Qualidade - PGMQ", "Programação Financeira", "Programação Orçamentária e Financeira", "Programação de aquisição dos medicamentos e insumos estratégicos sob a gestão do Dathi", "Projetar despesas Orçamentárias e Financeiras da SUGEF", "Projeto \"Comer pra quê?\"", "Projeto InovaSAN", "Projetos do portfólio da AEB", "Projetos referentes às operações de TIC", "Projeção de indicadores fiscais de médio prazo para a condução da política fiscal", "Promover articulação com o Parlamento", "Promover e aprimorar a gestão e a governança nas ações do Dathi para HIV, aids, Tuberculose, Hepatites Virais e Infecções Sexualmente Transmissíveis", "Promover estudos e  avaliar programas e políticas públicas", "Promoção Cultural Afro-brasileira", "Promoção da Conservação das exposições\nsob custódia da FUNDAJ", "Promoção da Transformação Digital da Secretaria do Tesouro Nacional", "Promoção da gestão, manutenção e aperfeiçoamento do sistema de emissão de licenças de exportação, importação e reexportação de espécies, produtos e subprodutos da fauna e flora pertencentes ou não aos anexos da Cites - SisCites", "Promoção da gestão, manutenção e aperfeiçoamento dos sistemas de informação de controle do uso dos recursos da flora e cadastro de áreas de interesse ambiental", "Promoção da interoperabilidade dos sistemas da saúde para o registro do acompanhamento das condicionalidades", "Promoção da proteção de dados pessoais e a adequação do DAENT/SVSA à Lei Geral de Proteção de Dados", "Promoção da publicação de matérias do Ibama no Diário Oficial da União - DOU", "Promoção da qualidade de vida no trabalho", "Promoção da transparência ativa e passiva", "Promoção da valorização dos servidores", "Promoção das revisões no Programa de Integridade, Política de Gestão de Riscos e Integridade, Plano de Gestão de Riscos no processos finalísticos e no Plano de dados Abertos", "Promoção de ações de facilitação de comércio exterior junto aos demais órgãos anuentes", "Promoção de ações educacionais", "Promoção de estratégias de capacitação e educação permanente para atuação na vigilância em saúde", "Promoção de estudos e pesquisas educativas", "Promoção de eventos de difusão de conhecimento", "Promoção de eventos institucionais", "Promoção de eventos técnicos científicos", "Promoção de parcerias nacionais e internacionais", "Promoção e implantação de Colegiados de Governança", "Promoção e implementação de programas de treinamento e sensibilização para técnicos do Prevfogo e comunidades", "Propor Informação em Ações Judiciais e Prestar Subsídio à Defesa da Fazenda Nacional no Contencioso Judicial", "Propor atividades acadêmicas (seminários, minicursos, palestras) para professores e alunos da pós-graduação", "Proposição de Trilhas de Capacitação para TI em Negócios", "Proposição de ajustes no Sistema de Condicionalidades (SICON)", "Proposição de ações para a consolidação e o fortalecimento dos instrumentos e instâncias de negociação e pactuação do SUAS", "Proposição de diretrizes para consumo em nuvem (FinOps) das soluções envolvendo Inteligência Artificial", "Proposição de normas e metodologia de gestão de continuidade de negócios", "Proposição de normas e metodologia de gestão de integridade", "Proposição de normas e metodologia de gestão de riscos e mudanças no limite de tolerância a riscos", "Proposição e apoio ações compartilhadas de educação ambiental e ações de formação continuada em parceria com os órgãos do Sisnama, entidades públicas e organizações da sociedade civil que desenvolvam atividades ligadas à área ambiental", "Proposição, coordenção e execução de convênios e cooperações técnicas nacionais e internacionais, visando ao aprimoramento e a atuação complementar e compartilhada das ações relacionadas ao monitoramento e à gestão das informações ambientais", "Prospecção, análise e classificação de tecnologias espaciais", "Proteção das informações da STN contra ameaças e vulnerabilidades", "Proteção de dados e conhecimentos sensíveis da organização", "Protolo de documentos no processo judicial eletrônico (PJe)", "Prover informações sob demanda", "Prover pessoas", "Provimento e nomeação de pessoal", "Provisão de infraestrutura para preservação do acervo bibliográfico", "Publicar atos administrativos", "Publicação de manuais e guias para acesso a serviços de TI", "Publicação de relatórios de transparência e prestação de contas sobre as ações e impactos dos EqSAN", "Qualificar o cuidado de pessoas vivendo com HIV ou aids (PVHA) na rede de serviços do SUS", "Qualificação Cadastral da folha de pagamentos do Programa Bolsa Família", "Qualificação cadastral da folha de pagamentos do programa Bolsa Família", "Qualificação das políticas de inclusão produtiva rural e acesso à água a partir do diálogo com os parceiros", "Qualificação de Ações de SAN no SUAS", "REALIZAR A COMUNICAÇÃO INTERNA", "REALIZAR A COMUNICAÇÃO SOCIAL", "REALIZAR A PESQUISA E SELEÇÃO PARA REPRESSÃO ADUANEIRA", "REALIZAR ASSISTÊNCIA PARA CONFORMIDADE TRIBUTÁRIA", "REALIZAR LAVRATURA DO AUTO DE INFRAÇÃO E REPRESENTAÇÕES CORRELATAS", "REALIZAR O PLANEJAMENTO INTEGRADO DA FISCALIZAÇÃO", "REALIZAR OPERAÇÕES DE REPRESSÃO", "REALIZAR PESQUISA E SELEÇÃO PARA FISCALIZAÇÃO ADUANEIRA", "REALIZAR PROCEDIMENTO FISCAL DE DILIGÊNCIA", "REALIZAR PROCEDIMENTO FISCAL DE FISCALIZAÇÃO", "REALIZAR PROCEDIMENTO FISCAL DE INFORMAÇÃO FISCAL", "REALIZAR PROCEDIMENTO FISCAL DE REVISÃO DE DECLARAÇÕES", "REALIZAR RETENÇÃO E APREENSÃO DE MERCADORIAS", "REALIZAR TRATAMENTO DE CONTROLES ESPECIAIS", "REALIZAÇÃO DA EXECUÇÃO ORÇAMENTÁRIA E FINANCEIRA", "Realizada a análise de projetos do setor público submetidos à COFIEX", "Realizar 05 oficinas regionais de qualificação da gestão das Assistências farmacêuticas estaduais no âmbito do componente estratégico", "Realizar a Execução Orçamentária e Financeira", "Realizar a Gestão do Dia a Dia dos Processos de Trabalho", "Realizar a Governança da Proteção de Dados Pessoais", "Realizar a organização do processo de trabalho da equipe por meio da gestão de processos no SEI", "Realizar análise de dados e estudos diagnósticos para implementação de ações / melhorias", "Realizar atividades de combate a fraudes fiscais estruturadas", "Realizar auditoria fiscal", "Realizar cobrança administrativa", "Realizar estudos de jurisprudência", "Realizar estudos de viabilidade para fornecimento de medicamentos pediátricos no âmbito do CESAF (Componente Estratégico da Assistência Farmacêutica)", "Realizar gestão administrativa", "Realizar interlocuções com redes e atores do sistema de ciência, tecnologia e inovação", "Realizar investigação de origem de mercadoria", "Realizar manifestações técninas sobre os projetos fomentados (existentes ou novos)", "Realizar o controle da admissão temporária", "Realizar o controle do repetro", "Realizar o controle do trânsito aduaneiro", "Realizar orientação sobre proteção de dados", "Realizar pagamento", "Realizar procedimentos de pesquisa e investigação", "Realizar reuniões estratégicas e trocas de informações com setores externos do Decit/SECTICS/MS", "Realizar seleção para conferência aduaneira na exportação", "Realizar seleção para conferência aduaneira na importação", "Realização  bianual do Prêmio  de Direitos Humanos", "Realização  bianual do Prêmio Luiz Gama de Direitos Humanos", "Realização  bianualmente do Prêmio Luiz Gama de Direitos Humanos", "Realização da Câmara Técnica para migrantes e refugiados da Comissão Intergestores Tripartite (CIT)", "Realização da análise de demanda de Segurança da Informação e Comunicação (SIC)", "Realização da consolidação dos atos normativos editados pelo Ibama", "Realização da gestão do Sistema Nacional de Emergências Ambientais (Siema), incluindo coordenação do monitoramento e análise de dados sobre acidentes ambientais", "Realização da gestão e da publicidade do acervo normativo na página do Ibama na internet", "Realização de análise de negócio", "Realização de análises para o monitoramento da mpox, e divulgação do diagnóstico situacional", "Realização de atividades educativas", "Realização de auditorias", "Realização de ações relacionadas a administração predial, contratações, gestão de transporte e gestão de materiais, bens e patrimônio.", "Realização de ações relacionadas à gestão do conhecimento, informações, dados e documentos.", "Realização de ações relacionadas à gestão dos recursos financeiros.", "Realização de ações voltadas a desenvolvimento, capacitação, avaliação registro e qualidade de vida dos servidores e funcionários.", "Realização de consultas públicas", "Realização de diagnóstico de competências", "Realização de eventos/encontros sobre programa Bolsa Família", "Realização de iniciativas de mobilização e articulação intersetorial e interfederativa para desenvolver ações complementares para mulheres beneficiárias do Programa Bolsa Família", "Realização de triagem hospitalar.", "Realização do alinhamento dos processos internos e projetos com as prioridades estratégicas", "Realização e registro de operações de crédito", "Recebimento, planejamento, distribuição, monitoramento e acompanhamento das demandas internas e externas.", "Recepção, tratamento e encaminhamento das demandas oriundas do Ministério Público Federal", "Recepção, tratamento e encaminhamento das respostas a Mandado de Segurança", "Recomendações Cesta Básica", "Reconhecer Direitos Previdenciários e Estatutários", "Recrutamento e Seleção da Equipe e Profissionais de Saúde", "Recurso apresentado pelos beneficiários", "Reformulação da Política de movimentação de pessoal", "Registro das atividades de execução e da conformidade de gestão no Sistema Integrado de Administração Financeira do Governo Federal - SIAFI", "Registro de atos de pessoal", "Registro de objetos espaciais", "Registro e Proteção de Obras Intelectuais", "Registro e reconhecimento contábil", "Regulamentação da Fundação Escola do SUAS Simone Albuquerque", "Regulamentação de pensão especial para crianças ou adolescentes, órfãos em razão do crime de feminicídio", "Regulamentação do Marco Legal do Setor Espacial", "Regularização do limite de saque", "Regulação do novo prontuário eletrônico do SUAS", "Relacionamento com organismos Internacionais do Setor Espacial", "Relatórios consolidados de desenvolvimento rural publicados. prognósticos, diagnósticos e estudos relativos ao setor agropecuário", "Relatórios de repasse aos entes Estaduais dos tributos estaduais recolhidos conjuntamente com o tributo federal (TCFA)", "Relatórios finais PIBIC e de conclusão da pós-graduação", "Relação do Sistema Único de Assistência Social com o Sistema de Justiça", "Relações Institucionais", "Renovação de contratos de residentes", "Renovação de contratos de voluntários", "Representar o Departamento de Condicionalidades  em cursos, seminários, rodas de conversa e debates", "Representar unidade em foruns consultivos/ decisórios", "Representar unidade em foruns consultivos/ deliberativos", "Representação do MDS na Comissão Intersetorial do Sistema Nacional de Atendimento Socioeducativo (SINASE)", "Representação do MDS na Comissão Nacional de Erradicação do Trabalho Infantil (CONAETI)", "Representação do MDS na Comissão Nacional de Erradicação do Tráfico de Pessoas CONATRAP", "Representação do MDS no ComitêNacional de Prevenção e Combate à Tortura (CNPCT)", "Representação do Ministério do Desenvolvimento e Assistência Social, Família e Combate\nà Fome no Comitê Gestor da Política Nacional de Gestão Territorial e Ambiental Quilombola - PNGTAQ", "Representação em colegiados intersetoriais relacionados à proteção social básica", "Representação institucional", "Representação institucional da Diretoria do Departamento de Proteção Social Especial", "Representação institucional da Unidade em espaços de assistência e serviços sociais", "Representação institucional em eventos", "Representação institucional em eventos e ações de inovação do agro.", "Representação institucional na camex", "Representação no Comissão Intersetorial de enfrentamento à Violência Sexual Contra Crianças e Adolescentes (CIEVSCA)", "Representação no Comitê Gestor da Política Nacional de Busca de Pessoas Desaparecidas", "Representação no Comitê Gestor da Política Nacional de Prevenção da Automutilação e do Suicído (CGPNPAS)", "Representação no Comitê Nacional de Prevenção e Combate a Tortura (CNPCT)", "Representação no Conselho Nacional da Criança e do Adolescente (CONANDA)", "Representação no Conselho Nacional da Pessoa Idosa (CNDI)", "Requerimento de Informação Parlamentar (RIC e RQS) analisado preliminarmente", "Requerimento de Informação Parlamentar (RIC e RQS) monitorado", "Responder a demandas de informação em HIV/Aids", "Restauração Ecológica", "Revisão das diretrizes vinculadas ao Relatório de Gestão da Casa", "Revisão do planejamento de longo prazo", "Revisão, editoração e tradução de publicações produzidas pelas unidades do Ibama. Coordenação do Comitê Editorial.", "Risco Fiscal das Estatais", "Rotinas de trabalho", "Saúde Bucal", "Saúde da Família e Comunidade", "Secretaria Executiva da Mesa Nacional de Negociação Permanente do Sistema Único de Assistência Social (MNN-SUAS)", "Secretaria Executiva do Comitê Pemanente MDS para Proteção Social em Emergências e Calamidades Públicas", "Secretaria Executiva do Grupo de Trabalho Intersetorial para o planejamento de estratégias e ações integradas voltadas à implantação, ampliação e qualificação do Serviço de Acolhimento em Família Acolhedora.", "Secretaria técnica da Comissão Intergestores Tripartite (CIT) e suas Câmaras Técnicas e Grupos de Trabalho", "Secretaria-executiva de conselhos e comitês da Câmara de Comércio Exterior", "Secretariado das reuniões do Conselho Gestor", "Seleção do Fornecedor", "Serviço de Intercâmbio Bibliográfico", "Simplificado e aumentado  transparência e previsibilidade em 100% dos processos de autorização para preparação de projetos no âmbito da Cofiex (SEAID);", "Sistema de Fiscalização (SISFIS)", "Sistema de Registro de Preços (SRP)", "Sistemas de Comércio Exterior", "Solicitar Descentralização Orçamentária", "Solicitar recursos financeiros", "Solicitação de Afastamento a serviço para o exterior", "Solicitação de diárias, transporte e passagens para ações de formação e desenvolvimento de pessoas", "Solicitação de serviços e materiais de expediente", "Solucionar consultas sobre legislação tributária e aduaneira e classificação de mercadorias, serviços e intangíveis", "Space Farming", "Subsidiar a avaliação dos resultados das políticas públicas implementadas com benefício fiscal", "Subsidiar defesa da União", "Subsídio à ASRCC no atendimento a demandas de órgãos de controle", "Subsídio à proposta orçamentária de TIC", "Subsídio às áreas de negócio para a priorização das necessidades da organização relacionadas à Tecnologia da Informação e Comunicação", "Subsídios e resposta às demandas", "Subsídios internos para a AEB", "Subsídios à PFE/AGU e Unidades Administrativas do Ibama acerca da cobrança contenciosa da TCFA.", "Supervisionar as proposições normativas", "Supervisão da adequação processual e documental de todos os processos e procedimentos", "Supervisão da adequação processual e documental de todos os processos e procedimentos.", "Supervisão do sistema de controle interno", "Supervisão técnica das entidades/OSCs certificadas com a Certificação de Entidades Beneficentes de Assistência Social (CEBAS)", "Suporte", "Suporte Administrativo", "Suporte Administrativo ao Colegiado sobre Projetos e Parcerias", "Suporte Administrativo à Destinação dos Recursos de Compensação Ambiental", "Suporte a projetos e soluções de Inteligência de Negócio", "Suporte ao acesso de informações da unidade", "Suporte aos usuários do Sistema de Cadastro das Instituições de Uso Científico de Animais (CIUCA)", "Suporte e atendimento aos usuários de TIC do Ibama", "Suporte em operações aéreas", "Suporte quanto às Ações Civis Públicas (ACPs) relacionadas às emergências ambientais", "Suporte técnico para acesso ou uso de sistemas corporativos", "Suporte técnico para instalações de hardware e equipamentos", "Suporte técnico para uso ou instalações de hardware e equipamentos", "Suporte técnico-administrativo à gestão de processos consultivos quanto à adequação jurídica", "Suporte técnico-administrativo à gestão de processos de consulta ética", "Suporte à fiscalização ambiental", "Suporte às demandas referentes a proposições legislativas pertinentes as emergências ambientais", "Suprimento de Fundos - Cadastro de Portador", "Suprimento de Fundos - Concessão e Pagamento", "Suprimento de Fundos - Prestação de Contas", "Sustentação de soluções de Inteligência de Negócio", "TRATAR EXPEDIENTES RELATIVOS À SOLICITAÇÃO DE ANÁLISES OU INSTAURAÇÃO DE PROCEDIMENTOS FISCAIS", "Tomada de Contas", "Tomada de Preços", "Tomadas de Contas Especial", "Tramitação Legislativa", "Transferência de recursos", "Transparência e Atendimento Institucional", "Transparência e Prestação de Contas", "Tratamento adequado dos dados pessoais conforme a LGPD e demais normativos", "Tratamento de demandas de Ouvidoria", "Tratamento e arquivamentos de documentos físicos", "Tratamento e publicização de dados oficiais", "Tratar incidentes com dados pessoais", "Triagem de Auto de Infração - AI e Taxa de Controle de Fiscalização Ambiental - TCFA prescritos", "Triagem e distribuição de processos", "VENDA/DOAÇÃO DE LIVROS", "VIABILIZAR AÇÕES DE COOPERAÇÃO", "Validar dados da arrecadação", "Validação de estruturas e scripts de dados", "Viabilizar e coordenar a realização de estudos e pesquisas destinados à produção de conhecimento na área do financiamento federal da Atenção Primária à Saúde", "Vigilância, Promoção e Comunicação em Saúde", "Voluntariado", "Zelar pela transparência das estatísticas fiscais dos assuntos acompanhados pela SUGEF", "comunicação institucional. Orientação e Articulação de Cadeias Produtivas  EM seminários,fóruns, reuniões", "não é entrega", "Índice de preços da Nova Cesta Básica de Alimentos (IPCBA)"], "corpus": [{"cod": "08470121", "ent": "Produto 2 (estudo sobre o processo de participação social no ciclo regulatório) desenvolvido", "mac": "Competitividade e política regulatória", "cat": "Política Regulatória", "serv": "", "nat": "Finalístico"}, {"cod": "10970008", "ent": "Sessões de julgamento de processos sancionadores realizadas", "mac": "Gestão de políticas de estímulo à eficiência, inovação e competitividade​", "cat": "Melhoria regulatória e do ambiente de negócios​", "serv": "", "nat": "Finalístico"}, {"cod": "10530004", "ent": "PSP anual homologado", "mac": "Diretrizes estratégicas para comunicações", "cat": "Gestão de Serviços Postais", "serv": "", "nat": "Finalístico"}, {"cod": "01210003", "ent": "Painéis Orçamentários atualizados", "mac": "Programar orçamento institucional", "cat": "Orçamento", "serv": "Atualizar painés oçamentários", "nat": "Finalístico"}, {"cod": "06580009", "ent": "Supervisão de estágio e preceptoria de residência realizada", "mac": "Gestão do trabalho em saúde", "cat": "Ações de atenção à saúde indígena", "serv": "", "nat": "Finalístico"}, {"cod": "01070608", "ent": "Desligamento de servidor realizado", "mac": "Reconhecer direitos previdenciários e estatutários", "cat": "Cadastro e registro de pessoal", "serv": "", "nat": "Finalístico"}, {"cod": "08250044", "ent": "Carteira de projetos e pipeline gerenciados.", "mac": "Gestão de políticas de cooperação internacional, comércio exterior e integração regional e global​", "cat": "Negociações Internacionais", "serv": "", "nat": "Finalístico"}, {"cod": "02570152", "ent": "Dados e informações lançados no SDA e em sistemas diversos", "mac": "Gestão do conhecimento e da informação corporativa", "cat": "Gestão da informação corporativa", "serv": "", "nat": "Governança"}, {"cod": "03500408", "ent": "Informações detalhadas sobre renegociação de créditos devidamente levantadas.", "mac": "Gestão de créditos e defesa do interesse público​", "cat": "Pesquisa, estudos e estatísticas", "serv": "", "nat": "Finalístico"}, {"cod": "03790871", "ent": "Tutoriais e guias", "mac": "Gerenciar a qualidade", "cat": "Produção, análise e disseminação de dados e informações institucionais", "serv": "Gestão do conhecimento", "nat": "Finalístico"}, {"cod": "11990018", "ent": "Pesquisa realizada sobre estratégias efetivas para promoção da alimentação saudável, em parceria com MGI e Virgínia Tech", "mac": "Promoção do acesso e do consumo de alimentação adequada e saudável", "cat": "Estratégia de Prevenção da Obesidade", "serv": "Pesquisa sobre estratégias para promoção da alimentação saudável", "nat": "Finalístico"}, {"cod": "11120015", "ent": "Mídia veiculada comprovada e faturada", "mac": "Desenvolvimento estratégico do setor turístico", "cat": "Promoção, Estruturação e Qualificação de Produtos, Serviços e Destinos Turísticos", "serv": "", "nat": "Finalístico"}, {"cod": "14150020", "ent": "Proposta ou atos normativos de política ou programa de popularização da ciência e educação científica elaborados", "mac": "Promoção e fomento de ciência, tecnologia e inovação (CT&I)", "cat": "Promoção da ciência, tecnologia e inovação (CT&I)", "serv": "Formulação de políticas públicas ou programas de popularização da ciência e educação científica", "nat": "Finalístico"}, {"cod": "08660017", "ent": "Avaliações dos progamas de integridade de pessoas jurídicas de direito privado realizadas", "mac": "Promoção da integridade pública, transparência e acesso a informação", "cat": "Integridade privada", "serv": "", "nat": "Finalístico"}, {"cod": "03200162", "ent": "Participação em Comitês/Grupos de Trabalho/Colegiados do MDHC realizada", "mac": "Monitorar a gestão do desenvolvimento organizacional e da inovação", "cat": "Grupos de Trabalho e Comitês", "serv": "Gestão de relações institucionais e federativas", "nat": "Finalístico"}, {"cod": "09940001", "ent": "Estudos e propostas sobre tendências e perspectivas de evolução no setor de Exploração e Produção (E&P) de petróleo e gás avaliados", "mac": "Gestão da política de petróleo, gás natural e biocombustíveis", "cat": "Monitoramento do setor energético", "serv": "", "nat": "Finalístico"}, {"cod": "12960003", "ent": "Regime Repetro fiscalizado", "mac": "Administração aduaneira", "cat": "CONTROLAR E FISCALIZAR REGIMES ADUANEIROS", "serv": "Realizar o controle do repetro", "nat": "Finalístico"}, {"cod": "01080104", "ent": "Certificado de Estágio Emitido", "mac": "Desenvolver pessoas", "cat": "Recrutamento e seleção de pessoal", "serv": "", "nat": "Finalístico"}, {"cod": "10430053", "ent": "Processo analisados - despachos", "mac": "Monitorar a gestão da comunicação institucional", "cat": "Gestão dos Serviços de Radiodifusão", "serv": "", "nat": "Finalístico"}, {"cod": "02590463", "ent": "Acompanhamento do desenvolvimento do novo módulo de manifestações de disponibilidade orçamentária de Pessoal e Encargos Sociais realizado", "mac": "Desenvolver sistemas corporativos", "cat": "Gestão de sistemas informatizados", "serv": "", "nat": "Finalístico"}, {"cod": "10180013", "ent": "Plano nacional de auditorias presenciais do PMLS - Programa Mais Leite Saudável planejado  Etapa 2", "mac": "Gestão da política de fomento à produção agropecuária", "cat": "Fomento à produção agropecuária", "serv": "", "nat": "Finalístico"}, {"cod": "01400492", "ent": "Processo SEI do contrato via ABDI instruído", "mac": "Gestão da logística pública", "cat": "Gestão de contratos", "serv": "", "nat": "Suporte"}, {"cod": "12780084", "ent": "Processo saneado", "mac": "Prevenção e solução de litígios tributários e aduaneiros", "cat": "Formular Atos Interpretativos e Normativos", "serv": "Formular Atos Interpretativos", "nat": "Finalístico"}, {"cod": "09120010", "ent": "Minuta das Diretrizes do Planejamento da Inspeção do Trabalho, no âmbito do combate ao trabalho análogo ao de escravizado e tráfico de pessoas, produzida.", "mac": "Inspeção do trabalho", "cat": "Planejamento das ações fiscalização do trabalho", "serv": "", "nat": "Finalístico"}, {"cod": "11250040", "ent": "Ajuizamento coletivo realizado", "mac": "Gestão da dívida ativa", "cat": "Gestão dos Créditos e Defesa do Interesse Público", "serv": "", "nat": "Finalístico"}, {"cod": "01090018", "ent": "Solicitação de FG/FCT analisada", "mac": "Gerenciar recrutamento e seleção de pessoas", "cat": "Provimento e movimentação de pessoal", "serv": "Gestão de licenças e afastamentos", "nat": "Finalístico"}, {"cod": "06090002", "ent": "Pedido de parcelamento analisado", "mac": "Tributação", "cat": "Setor Audiovisual - Arrecadação de Receitas do Setor Audiovisual - Gestão de Créditos Não Tributários", "serv": "Gestão de Créditos Não Tributários", "nat": "Finalístico"}, {"cod": "10180033", "ent": "Demandas de produção vegetal atendidas", "mac": "Gestão da política de fomento à produção agropecuária", "cat": "Fomento à produção agropecuária", "serv": "", "nat": "Finalístico"}, {"cod": "09580001", "ent": "Objetos e localidades das ações e programas definidos.", "mac": "Gestão de investimentos aeroportuários e aviação civil", "cat": "Investimentos em Infraestrutura aeroportuária e infraestrutura aeronáutica regional", "serv": "", "nat": "Finalístico"}, {"cod": "05110069", "ent": "Solicitação de apoio operacional, com relação à política de segurança, analisada", "mac": "Justiça e cidadania", "cat": "Segurança pública", "serv": "", "nat": "Finalístico"}, {"cod": "05100135", "ent": "Relatório final, nota técnica e despacho de procedimentos e processos correcionais analisados", "mac": "Desenvolver procedimentos correicionais", "cat": "Corregedoria e responsabilização", "serv": "", "nat": "Finalístico"}, {"cod": "09760015", "ent": "Poligonal de Portos Organizados avaliado e publicado", "mac": "Gestão da política portuária", "cat": "Planejamento da Política Portuária", "serv": "", "nat": "Finalístico"}, {"cod": "12710003", "ent": "Atos publicados", "mac": "Gestão do processo administrativo digital", "cat": "Gerir Acervo, Modelos, Classificação e Ciência de Documentos Administrativos", "serv": "Executar os Procedimentos de Ciência de Documentos Administrativos", "nat": "Finalístico"}, {"cod": "03640006", "ent": "Suplementos nutricionais relacionados com programas estratégicos de prevenção e controle de micronutrientes adquiridos", "mac": "Gestão de insumos estratégicos aplicados à saúde", "cat": "Promoção da alimentação e nutrição na saúde", "serv": "Fomento à produção nacional de insumos estratégicos", "nat": "Finalístico"}, {"cod": "12860006", "ent": "Gestão de meios de arrecadação - Débito automático processado", "mac": "Gestão do crédito tributário e da arrecadação", "cat": "GERIR A ARRECADAÇÃO", "serv": "Gerir débito automático", "nat": "Finalístico"}, {"cod": "12750002", "ent": "Editais de fomento divulgados", "mac": "Gestão da cultura e do desenvolvimento socioeconômico afro-brasileiro", "cat": "Preservação das manifestações culturais afro-brasileiras", "serv": "Promoção Cultural Afro-brasileira", "nat": "Finalístico"}, {"cod": "01070723", "ent": "Registro de assentamentos funcionais de colaboradores acompanhado", "mac": "Reconhecer direitos previdenciários e estatutários", "cat": "Cadastro e registro de pessoal", "serv": "", "nat": "Finalístico"}, {"cod": "09510004", "ent": "Conselheiros capacitados", "mac": "Gestão da política de hidrovias e navegação", "cat": "Fomento ao Setor de Construção Naval, Hidroviária e Portuária", "serv": "", "nat": "Finalístico"}, {"cod": "12810010", "ent": "Propostas de política aduaneira analisadas", "mac": "Gestão das políticas tributária e aduaneira", "cat": "Formular e analisar propostas de política tributária e aduaneira", "serv": "Analisar propostas de política tributária e aduaneira", "nat": "Finalístico"}, {"cod": "13900001", "ent": "Atendimento em mutirão realizado", "mac": "Atendimento ao público e ampliação do acesso à justiça", "cat": "Priorizar atividades itinerantes com foco na redução da miséria extrema", "serv": "Atuação em mutirões de atendimento", "nat": "Finalístico"}, {"cod": "02470003", "ent": "Atestado de Habilitação emitido", "mac": "Licenciamento de entidades e planos de previdência complementar", "cat": "Licenciamento", "serv": "Suporte técnico-administrativo à gestão de processos consultivos quanto à adequação jurídica", "nat": "Finalístico"}, {"cod": "13020024", "ent": "Resposta a intercâmbio encaminhada", "mac": "Relacionamento com sociedade, governos e organismos nacionais e internacionais", "cat": "CONDUZIR RELAÇÕES INTERNACIONAIS", "serv": "GERIR INTERCÂMBIO INTERNACIONAL DE INFORMAÇÕES", "nat": "Finalístico"}, {"cod": "01500475", "ent": "Prestação de contas (Suprimento de Fundos-CPGF) realizada tempestivamente", "mac": "Administrar bens de consumo e permanente", "cat": "Gestão Patrimonial", "serv": "Gestão de estoques e materiais de consumo da área", "nat": "Finalístico"}, {"cod": "09600001", "ent": "Ações das entidades vinculadas ao Ministério nos programas de aceleração do crescimento do setor portuário monitoradas", "mac": "Política econômica no setor portuário", "cat": "Desenvolvimento produtivo e economia da saúde do Setor Portuário", "serv": "", "nat": "Finalístico"}, {"cod": "12580005", "ent": "Tecnologia Crítica Desenvolvida", "mac": "Desenvolvimento e execução do programa espacial brasileiro", "cat": "Acesso ao espaço e missões espaciais", "serv": "Pesquisa, Desenvolvimento e Inovação", "nat": "Finalístico"}, {"cod": "06100012", "ent": "Relatório de coordenação da operação reversa e do plano de roteirização da distribuição interestadual e local, nas UF´s sede das aplicadoras, das provas do Enem e Pré-Testes de Intens, elaborado.", "mac": "Gerenciar programas e projetos", "cat": "Gestão de Exames e Pré-testes educacionais", "serv": "Implementação de Programa de Gestão de Desempenho", "nat": "Finalístico"}, {"cod": "07620296", "ent": "Transferências voluntárias realizadas no estado de São Paulo", "mac": "Gestão estratégica e governança", "cat": "Gestão de projetos", "serv": "", "nat": "Governança"}, {"cod": "13180012", "ent": "Cadastramento de famílias afetadas (demandas referentes a moradia), realizadas.", "mac": "Atuar na defesa e promoção dos direitos humanos", "cat": "Aperfeiçoar o diálogo externo", "serv": "Articulação institucional e apoio técnico em demandas de direitos humanos e moradia", "nat": "Finalístico"}, {"cod": "03770008", "ent": "Grupos de Trabalho criados", "mac": "Fomento ao desenvolvimento científico, tecnológico e produtivo em saúde", "cat": "Incorporação de Tecnologias em Saúde", "serv": "Gestão dos Contratos de Bens, Serviços e Insumos de Saúde Indígena", "nat": "Finalístico"}, {"cod": "11120016", "ent": "Uso da marca aprovado e documentado", "mac": "Desenvolvimento estratégico do setor turístico", "cat": "Promoção, Estruturação e Qualificação de Produtos, Serviços e Destinos Turísticos", "serv": "", "nat": "Finalístico"}, {"cod": "02400028", "ent": "Tratamento de resíduos sólidos de interesse da fiscalização federal agropecuária fiscalizado", "mac": "Defesa agropecuária", "cat": "Defesa agropecuária e vigilância  internacional", "serv": "", "nat": "Finalístico"}, {"cod": "10010003", "ent": "Politica publica relacionados ao PNB monitorada", "mac": "Gestão da política de sustentabilidade, inovação e qualidade da produção agropecuária", "cat": "Gestão das políticas públicas agropecuárias", "serv": "", "nat": "Finalístico"}, {"cod": "09270009", "ent": "Planilha Resumo das Receitas do FAT atualizada", "mac": "Administrar recursos financeiros", "cat": "Monitoramento das remunerações, receitas e aplicações do FAT", "serv": "", "nat": "Finalístico"}, {"cod": "09990013", "ent": "Resultados do Plano de Ação  dos ecossistemas  de inovação estaduais monitorados", "mac": "Desenvolvimento de soluções agropecuárias", "cat": "Fomento a Inovação Agropecuária", "serv": "", "nat": "Finalístico"}, {"cod": "02300083", "ent": "Expedientes e demandas internas distribuídas", "mac": "Consultoria e assessoramento jurídico", "cat": "Consultoria jurídica institucional", "serv": "", "nat": "Finalístico"}, {"cod": "02610213", "ent": "Relatório de Monitoramento Clínico das Pessoas Vivendo com HIV ou aids Coinfectadas com Mpox no Brasil elaborado", "mac": "Gestão do conhecimento e da informação corporativa", "cat": "Gestão e tratamento de dados", "serv": "", "nat": "Governança"}, {"cod": "02310208", "ent": "Atos normativos elaborados", "mac": "Gestão de políticas públicas", "cat": "Normatização", "serv": "Gestão dos Contratos de Bens, Serviços e Insumos de Saúde Indígena", "nat": "Finalístico"}, {"cod": "07620377", "ent": "Empresas dos projetos do DigitalBR e-commerce regional atendidas.", "mac": "Gerenciar programas e projetos", "cat": "Gestão de projetos", "serv": "", "nat": "Finalístico"}, {"cod": "09200003", "ent": "Base de dados do Gov.br Geridas", "mac": "Gestão de registros administrativos e de informações estatísticas do mercado de trabalho", "cat": "Gestão da emissão da carteira de trabalho e previdência social", "serv": "", "nat": "Finalístico"}, {"cod": "11250051", "ent": "Créditos brasileiros no exterior renegociados", "mac": "Gestão de créditos e defesa do interesse público​", "cat": "Gestão dos Créditos e Defesa do Interesse Público", "serv": "", "nat": "Finalístico"}, {"cod": "06070001", "ent": "Processo Administrativo Fiscal instaurado", "mac": "Tributação", "cat": "Setor Audiovisual - Arrecadação de Receitas do Setor Audiovisual - Gestão de Créditos Tributários", "serv": "Gestão de Créditos Não Tributários", "nat": "Finalístico"}, {"cod": "12040002", "ent": "Desenvolvimento de novo serviço no SouGov realizado", "mac": "Transformação digital", "cat": "Capacitação digital", "serv": "", "nat": "Finalístico"}, {"cod": "06090001", "ent": "Processo para cobrança administrativa instruído", "mac": "Tributação", "cat": "Setor Audiovisual - Arrecadação de Receitas do Setor Audiovisual - Gestão de Créditos Não Tributários", "serv": "Gestão de Créditos Não Tributários", "nat": "Finalístico"}, {"cod": "01300413", "ent": "Despacho com análise sobre o pedido de inclusão de demanda no Plano de Contratações Anual (PCA) elaborado", "mac": "Gerenciar contratações", "cat": "Compras e licitação", "serv": "Análise de conformidade em licitações e contratos", "nat": "Finalístico"}, {"cod": "09750004", "ent": "Estudos  para inovação e sustentabilidade de hidrovias realizados", "mac": "Gestão das infraestruturas hidroviárias", "cat": "Sustentabilidade das Infraestruturas Hidroviárias", "serv": "", "nat": "Finalístico"}, {"cod": "09800005", "ent": "Atividades da COP relacionadas ao setor mineral coordenadas", "mac": "Gestão do planejamento setorial de mineração", "cat": "Desenvolvimento socioambiental do setor mineral", "serv": "", "nat": "Finalístico"}, {"cod": "10450009", "ent": "Processo de impugnação do lançamento do crédito tributário (análise de tempetividade) analisado", "mac": "Normatização e efetivação das políticas de comunicações", "cat": "Gestão de Fundos de Comunicações", "serv": "", "nat": "Finalístico"}, {"cod": "01230172", "ent": "Faturas de Cartão de Pagamento do Governo Federal Encaminhadas para Pagamento", "mac": "Administrar recursos financeiros", "cat": "Contabilidade", "serv": "", "nat": "Finalístico"}, {"cod": "07050081", "ent": "Relatório de participação em evento internacional de radiodifusão elaborado", "mac": "Cooperação jurídica internacional", "cat": "Relações Internacionais e diplomáticas", "serv": "", "nat": "Governança"}, {"cod": "10440011", "ent": "Projetos decorrentes da emissão de valores mobiliários acompanhados", "mac": "Normatização e efetivação das políticas de comunicações", "cat": "Acompanhamento e avaliação de programas e ações de comunicações", "serv": "Normatização e Efetividade das Politicas de Comunicações", "nat": "Finalístico"}, {"cod": "08570002", "ent": "Instrumento legal que ampare a gestão interministerial do sistema e a contratação de consultorias técnicas afins elaborado", "mac": "Gestão de tecnologias da informação e comunicação", "cat": "Políticas de investimento", "serv": "", "nat": "Suporte"}, {"cod": "05040077", "ent": "Medidas de insdisponibilidade advindas de RCSNU - Respostas recebidas dos órgãos e instituições indicados no  art. 10, § 1° da Lei N° 13.810 incluídas nos processos SEI respectivos", "mac": "Gestão jurídica institucional", "cat": "Cooperação jurídica internacional", "serv": "", "nat": "Governança"}, {"cod": "03110034", "ent": "Processos Eleitorais ocorridos (Portarias de Designação dos Coordenadores de Cursos)", "mac": "Educação superior", "cat": "Ensino - Gestão Acadêmica", "serv": "Compensação Previdenciária", "nat": "Finalístico"}, {"cod": "08110031", "ent": "Solicitações de controle de material de apoio às atividades do laboratório realizadas", "mac": "Gestão de políticas de estímulo à eficiência, inovação e competitividade​", "cat": "Inovação Tecnológica e Produtiva", "serv": "", "nat": "Finalístico"}, {"cod": "10440004", "ent": "Reunião do GAPE realizada", "mac": "Normatização e efetivação das políticas de comunicações", "cat": "Acompanhamento e avaliação de programas e ações de comunicações", "serv": "Normatização e Efetividade das Politicas de Comunicações", "nat": "Finalístico"}, {"cod": "00060037", "ent": "Rede de cooperação e informações arquivísticas Memórias Reveladas gerida", "mac": "Promoção do acesso ao patrimônio documental nacional", "cat": "Preservação do patrimônio documental nacional", "serv": "Gestão de recursos do Programa Mais Médicos para o Brasil", "nat": "Finalístico"}, {"cod": "01920064", "ent": "Petições 261 de de retificação por erro do INPI analisadas", "mac": "Atividade econômica", "cat": "Propriedade Industrial - concessão de patentes", "serv": "", "nat": "Finalístico"}, {"cod": "07130013", "ent": "Diagnóstico dos benefícios concedidos por empresa estatal em Acordo Coletivo de Trabalho (ACT) atualizado", "mac": "Governança e coordenação das empresas estatais", "cat": "Política de Pessoal das Empresas Estatais", "serv": "Suporte técnico-administrativo à gestão de processos consultivos quanto à adequação jurídica", "nat": "Finalístico"}, {"cod": "11400069", "ent": "Controle geral de frequências", "mac": "Administrar as relações de trabalho", "cat": "Gestão administrativa de pessoal", "serv": "", "nat": "Finalístico"}, {"cod": "06120158", "ent": "Relatório da medida cautelar Comunidade Guapoý Guarani Kaiowá enviado", "mac": "Relações institucionais e federativas", "cat": "Articulação institucional", "serv": "", "nat": "Governança"}, {"cod": "06010414", "ent": "Comunicados ou boletins de oportunidades internacionais para empresas do setor emitidos", "mac": "Gestão da política de desenvolvimento de atividades espaciais", "cat": "Comunicação institucional", "serv": "Apoio à internacionalização de empresas e à promoção de negócios internacionais", "nat": "Finalístico"}, {"cod": "13090004", "ent": "Reforma Tributária do Consumo - Programa de Incentivo implementado", "mac": "Gestão do crédito tributário e da arrecadação", "cat": "Gerir Direito Creditório de Contribuinte", "serv": "", "nat": "Finalístico"}, {"cod": "12240094", "ent": "Relatório de acompanhamento de envio de informações necessárias à consolidação gerado", "mac": "Transparência das contas públicas", "cat": "Informações e relatórios contábeis, orçamentários e fiscais da União", "serv": "Consolidação das Contas Públicas da Federação", "nat": "Finalístico"}, {"cod": "02310058", "ent": "Contribuições a proposta de ato normativo, guias e manuais apresentada", "mac": "Regulação e fiscalização", "cat": "Normatização", "serv": "Elaboração, revisão e consolidação de atos normativos, definição de procedimentos e mapeamento e melhoria de processos organizacionais.", "nat": "Finalístico"}, {"cod": "01070281", "ent": "Concessão de pensão civil vitálicia realizada", "mac": "Administrar as relações de trabalho", "cat": "Cadastro e registro de pessoal", "serv": "Fomento à produção nacional de insumos estratégicos", "nat": "Finalístico"}, {"cod": "12710001", "ent": "Comprovante de ciência registrada", "mac": "Gestão do processo administrativo digital", "cat": "Gerir Acervo, Modelos, Classificação e Ciência de Documentos Administrativos", "serv": "Executar os Procedimentos de Ciência de Documentos Administrativos", "nat": "Finalístico"}, {"cod": "10410002", "ent": "Acompanhamento de certificados de capacitação realizado", "mac": "Administrar suporte técnico", "cat": "Inclusão Digital e Acesso à Informação", "serv": "Expansão e Massificação das Comunicações", "nat": "Finalístico"}, {"cod": "06680012", "ent": "Comunicações realizadas", "mac": "Gestão e controle institucional", "cat": "Atividade correicional", "serv": "", "nat": "Finalístico"}, {"cod": "12800031", "ent": "Auto de Infração de aplicação de multa lavrado", "mac": "Administração aduaneira", "cat": "REALIZAR A VIGILÂNCIA E REPRESSÃO", "serv": "REALIZAR RETENÇÃO E APREENSÃO DE MERCADORIAS", "nat": "Finalístico"}, {"cod": "01540068", "ent": "Apoio aos gestores prestado", "mac": "Gestão da logística pública", "cat": "Concessão de diárias e passagens", "serv": "Solicitação de diárias, transporte e passagens para ações de formação e desenvolvimento de pessoas", "nat": "Suporte"}, {"cod": "07820025", "ent": "Qualificação sobre os processos educativos indígenas no âmbito de instâncias de monitoramento, avaliação, comitês, grupos de trabalho.", "mac": "Promoção e proteção dos direitos e cidadania dos povos indígenas", "cat": "Promoção e Qualificação dos Processos Educativos Indígenas", "serv": "Gestão de recursos do Programa Mais Médicos para o Brasil", "nat": "Finalístico"}, {"cod": "12530039", "ent": "Modelo de análise de dados ou de projeção desenvolvido", "mac": "Gestão da administração financeira e fiscal", "cat": "Planejamento e a Programação Financeira", "serv": "Atuação como órgão central de administração financeira", "nat": "Finalístico"}, {"cod": "01230091", "ent": "Débito relacionado a diárias e passagens registrado no SIAFI", "mac": "Gestão financeira, contábil e de custos", "cat": "Contabilidade", "serv": "", "nat": "Suporte"}, {"cod": "02280010", "ent": "Certidão de revelia lavrada", "mac": "Gestão do crédito tributário, administrativo e arrecadação", "cat": "Cobrança e recuperação de crédito", "serv": "Lançamento tributário das taxas decorrentes do poder de polícia do Ibama", "nat": "Finalístico"}, {"cod": "00050018", "ent": "Documentos digitais e não digitais recolhidos, transferidos e doados", "mac": "Promoção do acesso ao patrimônio documental nacional", "cat": "Gestão do sistema de gestão de documentos e arquivos - SIGA", "serv": "Gestão de rendimentos do patrimônio indígena", "nat": "Finalístico"}, {"cod": "02050004", "ent": "Autorização de afretamento - docagem emitido", "mac": "Gestão de políticas de transporte aquaviário", "cat": "Transporte Aquaviário – Licitação e Outorga", "serv": "", "nat": "Finalístico"}, {"cod": "01030038", "ent": "Conferência do conteúdo publicado", "mac": "Excluir", "cat": "RH - Monitoramento e Controle", "serv": "Monitoramento e controle das informações de pessoal do departamento", "nat": "Finalístico"}, {"cod": "13400001", "ent": "Eventos para mobilização das cadeias produtivas para o uso de bioinsumos realizados", "mac": "Fomento à sustentabilidade e qualidade da produção e das práticas agropecuárias", "cat": "Promoção do uso de boas práticas e sistemas agropecuários", "serv": "", "nat": "Finalístico"}, {"cod": "06030043", "ent": "Sensibilização de órgãos/ entidades do SIPEC  sobre as novas regras de negociação coletiva", "mac": "Gestão de relações do trabalho no serviço público", "cat": "Disseminação, articulação e relações institucionais", "serv": "", "nat": "Finalístico"}, {"cod": "02510285", "ent": "Respostas aos pedidos de informações provenientes de órgãos e instituições nacionais elaboradas", "mac": "Relações institucionais e federativas", "cat": "Atendimento", "serv": "", "nat": "Governança"}, {"cod": "01140009", "ent": "Recebimento e arquivamento no SEI dos laudos de manutenção dos extintores e mangueiras", "mac": "Monitorar a manutenção predial", "cat": "Segurança do Trabalho", "serv": "", "nat": "Finalístico"}, {"cod": "02590547", "ent": "Caixa Lockss FBN gerida", "mac": "Preservação e gestão do patrimônio documental e bibliográfico nacional", "cat": "Gestão de sistemas informatizados", "serv": "Gestão de sistema de preservação digital", "nat": "Finalístico"}, {"cod": "05050059", "ent": "Notas Técnicas de Cadeira de Custódia e Reestruturação da Forense  elaboradas", "mac": "Segurança pública", "cat": "Combate à corrupção e à lavagem de dinheiro", "serv": "", "nat": "Finalístico"}, {"cod": "11070038", "ent": "Painel de Gestão de Dados de Emergências elaborado", "mac": "Pesquisa, avaliação e monitoramento da biodiversidade", "cat": "Pesquisa e Monitoramento da Biodiversidade", "serv": "Prevenção e Resposta às Emergências Climáticas e Epizootias", "nat": "Finalístico"}, {"cod": "13920051", "ent": "Cirurgias gerais realizadas", "mac": "Atenção hospitalar e urgência", "cat": "Atendimento à saúde", "serv": "", "nat": "Finalístico"}, {"cod": "07560038", "ent": "Avaliação Setorial das operadoras com programas de operadoras Certificadas elaborada", "mac": "Gerenciar a qualidade", "cat": "Qualidade na saúde suplementar", "serv": "", "nat": "Finalístico"}, {"cod": "09080019", "ent": "Minuta de Acordo de Cooperação Técnica de Parceria com a Petrobras elaborada", "mac": "Apoio ao trabalaho autogestionário, autônomo ou associado", "cat": "Apoio ao desenvolvimento da economia popular e solidária", "serv": "", "nat": "Finalístico"}, {"cod": "06880363", "ent": "Projetos Premiados", "mac": "Transferência de recursos", "cat": "Gestão de parcerias institucionais", "serv": "", "nat": "Finalístico"}, {"cod": "11720009", "ent": "Eventos culturais produzidos", "mac": "Gestão da política de preservação e difusão da produção bibliográfica e documental do país", "cat": "Difusão do acervo e promoção do acesso ao acervo bibliográfico e documental nacional", "serv": "Organização de eventos de promoção cultural", "nat": "Finalístico"}, {"cod": "12480045", "ent": "Participação de reuniões, revisão de trabalhos das gerências e atendimento a diversas demandas relacionadas à COSEC realizadas", "mac": "Financiamento público e operações de fomento e subvenções", "cat": "Provimento de subvenções a operações de crédito", "serv": "Garantir excelência dos programas de financiamento e subvenção", "nat": "Finalístico"}, {"cod": "01210014", "ent": "Despesas reprogramadas", "mac": "Programar orçamento institucional", "cat": "Orçamento", "serv": "Promover e aprimorar a gestão e a governança nas ações do Dathi para HIV, aids, Tuberculose, Hepatites Virais e Infecções Sexualmente Transmissíveis", "nat": "Finalístico"}, {"cod": "11720012", "ent": "Apoio concedido aos tradutores estrangeiros", "mac": "Gestão da política de preservação e difusão da produção bibliográfica e documental do país", "cat": "Difusão do acervo e promoção do acesso ao acervo bibliográfico e documental nacional", "serv": "Gestão de processos estratégicos", "nat": "Finalístico"}, {"cod": "07600030", "ent": "Relatório de acompanhamento dos processos administrativos julgados pela DICOL elaborado", "mac": "Regulação de serviços de saúde", "cat": "Qualidade regulatória", "serv": "", "nat": "Finalístico"}, {"cod": "12460016", "ent": "Atos normativo do setor espacial elaborados", "mac": "Desenvolvimento e execução do programa espacial brasileiro", "cat": "Regulação, licenciamento e fiscalização das atividades espaciais", "serv": "Elaboração das normas das atividades espaciais", "nat": "Finalístico"}, {"cod": "09270014", "ent": "Boletins de Informações Financeiras do FAT elaborado", "mac": "Gestão dos recursos FAT e FGTS", "cat": "Monitoramento das remunerações, receitas e aplicações do FAT", "serv": "", "nat": "Finalístico"}, {"cod": "07620601", "ent": "MAPEAMENTO DAS ATIVIDADES DA COMEQ PRODUZIDO", "mac": "Gestão do desenvolvimento organizacional e da inovação", "cat": "Gestão de processos", "serv": "", "nat": "Finalístico"}, {"cod": "11080009", "ent": "Planos de Visitação em Territórios Tradicionais em Unidades de Conservação - PVIS elaborados", "mac": "Criação e manejo de unidades de conservação", "cat": "Visitação", "serv": "Planejamento da visitação", "nat": "Finalístico"}, {"cod": "07270010", "ent": "Demandas institucionais sobre ações laboratoriais respondidas", "mac": "Gestão da política de atenção primária à saúde", "cat": "Laboratórios de Saúde Pública", "serv": "Gestão dos Contratos de Bens, Serviços e Insumos de Saúde Indígena", "nat": "Finalístico"}, {"cod": "09420015", "ent": "Reuniões ordinárias e extraordinárias do CGEE/PROCEL organizadas", "mac": "Gestão da política de transição energética e planejamento", "cat": "Eficiência energética", "serv": "", "nat": "Finalístico"}, {"cod": "09280002", "ent": "Despesas do FAT monitoradas", "mac": "Gestão dos recursos FAT e FGTS", "cat": "Gestão da execução orçamentária e finaceira do FAT", "serv": "", "nat": "Finalístico"}, {"cod": "09240005", "ent": "Notas técnicas de valores a serem inscritos pela setorial contábil no passivo dos benefícios do seguro-desemprego elaboradas", "mac": "Gestão de benefícios trabalhistas", "cat": "Gestão do seguro-desemprego", "serv": "", "nat": "Finalístico"}, {"cod": "06020018", "ent": "Atendimento ao público interno e externo realizado", "mac": "Participação e controle social", "cat": "Ouvidoria", "serv": "Prover informações sob demanda", "nat": "Governança"}, {"cod": "00070018", "ent": "Realização de apresentações musicais", "mac": "Cultura", "cat": "Difusão de acervo e conhecimento científico, cultural e artístico", "serv": "", "nat": "Finalístico"}, {"cod": "11270003", "ent": "PRFNs pelo LABJUD, durante os prazos de editais de transação, para resolução de dúvidas atendidas.", "mac": "Gestão integrada da representação e defesa da fazenda nacional", "cat": "Atuação na mediação e transação tributária para redução de litígios", "serv": "", "nat": "Finalístico"}, {"cod": "11460005", "ent": "Processos apontados para redução a termo das sessões de julgamentos distribuídos", "mac": "Segurança jurídica e ambiente econômico​", "cat": "Participação em conselhos", "serv": "", "nat": "Finalístico"}, {"cod": "12320002", "ent": "Projeções fiscais realizadas", "mac": "Planejamento estratégico fiscal", "cat": "Indicadores econômicos e fiscais", "serv": "Projeção de indicadores fiscais de médio prazo para a condução da política fiscal", "nat": "Finalístico"}, {"cod": "12840007", "ent": "Transação por adesão deferida", "mac": "Gestão do crédito tributário e da arrecadação", "cat": "GERIR O CRÉDITO TRIBUTÁRIO", "serv": "Realizar cobrança administrativa", "nat": "Finalístico"}, {"cod": "12870010", "ent": "Metodologia de carteiras setoriais revisada.", "mac": "Fiscalização tributária", "cat": "PLANEJAR A FISCALIZAÇÃO TRIBUTÁRIA", "serv": "DEFINIR OS MAIORES CONTRIBUINTES DE INTERESSE FISCAL", "nat": "Finalístico"}, {"cod": "09010001", "ent": "Pedidos de concessão de registro profissional concedidos", "mac": "Reconhecimento das ocupações e profissões", "cat": "Concessão do registro profissional e de contratante", "serv": "", "nat": "Finalístico"}, {"cod": "09970019", "ent": "Subcomitê 03 do Comitê da Resolução CNPE nº 10/2024 sobre Combustíveis Sustentáveis de Aviação - SAF coordenado", "mac": "Gestão da política de petróleo, gás natural e biocombustíveis", "cat": "Gestão da política de biocombustíveis", "serv": "", "nat": "Finalístico"}, {"cod": "03790005", "ent": "Dados e informações do System of Health Accounts (SHA) anual do Brasil enviados à Organização para a Cooperação e Desenvolvimento Econômico/OCDE", "mac": "Desenvolvimento e fortalecimento da economia da saúde", "cat": "Produção de dados e gestão do Conhecimento", "serv": "", "nat": "Finalístico"}, {"cod": "03630006", "ent": "Documento com contribuições técnico-científica para subsidiar  grupos de trabalhos, comitês, colegiados, sindicâncias e comissões voltadas as doenças crônicas não transmissíveis elaborado", "mac": "Gestão da política de atenção primária à saúde", "cat": "Prevenção às doenças crônicas", "serv": "Saúde da Família e Comunidade", "nat": "Finalístico"}, {"cod": "13590008", "ent": "Entregas Programadas de maquinário agrícola executadas", "mac": "Apoio à produção agropecuária", "cat": "Implementação de políticas de apoio à produção e comercialização", "serv": "", "nat": "Finalístico"}, {"cod": "09680272", "ent": "Planos, programas, eventos e missões definidos", "mac": "Gestão da política de desenvolvimento de atividades espaciais", "cat": "Planejamento, monitoramento e avaliação", "serv": "Planejamento operacional", "nat": "Finalístico"}, {"cod": "05040028", "ent": "Subtração Internacional de Crianças - Pedido ativo analisado e em conformidade", "mac": "Gestão jurídica institucional", "cat": "Cooperação jurídica internacional", "serv": "", "nat": "Governança"}, {"cod": "01150356", "ent": "Desembolso do valor de repasse do instrumento solicitado", "mac": "Gestão de transferências e repasses de recursos da União", "cat": "Gestão das transferências voluntárias", "serv": "", "nat": "Finalístico"}, {"cod": "11240080", "ent": "Aperfeiçoamento do arcabouço normativo e regulatório do Programa de Aquisição de Alimentos (PAA) publicado", "mac": "Gestão de políticas públicas de segurança alimentar e nutricional", "cat": "Acesso à alimentação adequada e saudável", "serv": "Gestão do Programa de Aquisição de Alimentos (PAA)\n\nElaboração, revisão e publicação de diretrizes e atos normativos para as ações de promoção da alimentação saudável\nNormatização das ações de promoção da alimentação saudável", "nat": "Finalístico"}, {"cod": "08670007", "ent": "Proposição de novos casos de operações especiais coordenados", "mac": "Enfrentamento e prevenção à corrupção", "cat": "Ações investigativas", "serv": "", "nat": "Finalístico"}, {"cod": "11540005", "ent": "Soluções tecnológicas propostas e desenvolvidas.", "mac": "Estímulo à eficiência, Inovação e competitividade", "cat": "Implementação de soluções tecnológicas para gestão da dívida ativa e recuperação de créditos", "serv": "", "nat": "Finalístico"}, {"cod": "07610001", "ent": "Levantamento das comissões da verdade subnacionais resultante de consultoria realizado", "mac": "Gestão de políticas públicas de direitos humanos e cidadania", "cat": "Direitos Humanos - Memória e verdade", "serv": "Compensação Previdenciária", "nat": "Finalístico"}, {"cod": "01700616", "ent": "Acompanhamento dos processos e documentos no SEI", "mac": "Gerenciar a documentação arquivística", "cat": "Gestão documental e arquivística", "serv": "Atendimento de demandas constantes em processos", "nat": "Finalístico"}, {"cod": "02310102", "ent": "Nota Técnica de Atendimento ao Parecer Jurídico elaborada", "mac": "Gestão de trabalho em saúde", "cat": "Normatização", "serv": "Gestão dos Contratos de Bens, Serviços e Insumos de Saúde Indígena", "nat": "Finalístico"}, {"cod": "11660005", "ent": "Recurso interposto pela parte decidido", "mac": "Gestão da política de proteção de dados pessoais", "cat": "Fiscalização e Auditoria em Proteção de Dados Pessoais", "serv": "", "nat": "Finalístico"}, {"cod": "11250017", "ent": "Apoio na Relação com os Órgãos de Origem realizado;", "mac": "Gestão da dívida ativa", "cat": "Gestão dos Créditos e Defesa do Interesse Público", "serv": "", "nat": "Finalístico"}, {"cod": "03600137", "ent": "Referência de nível medida", "mac": "Desenvolver pessoas", "cat": "Produção, análise e disseminação de dados e informações institucionais", "serv": "Gestão dos Contratos de Bens, Serviços e Insumos de Saúde Indígena", "nat": "Finalístico"}, {"cod": "07620613", "ent": "Procedimento Operacional Padrão (como se faz), com base no mapeamento do processo de trabalho feito elaborado - CNDH", "mac": "Gerenciar processos de negócio", "cat": "Gestão de processos", "serv": "Gestão da Estratégia Organizacional", "nat": "Finalístico"}, {"cod": "11020043", "ent": "Cooperação e participação em eventos nacionais e internacionais realizadas", "mac": "Gestão de políticas de cooperação internacional, comércio exterior e integração regional e global​", "cat": "Definição do posicionamento do Brasil​", "serv": "", "nat": "Finalístico"}, {"cod": "02590205", "ent": "Rotinas de gestão do sistema corrigidas", "mac": "Transformação digital", "cat": "Gestão de sistemas informatizados", "serv": "Gestão da Autorização de Funcionamento de Organização Estrangeira em Território Brasileiro – OE", "nat": "Finalístico"}, {"cod": "03760075", "ent": "Conexão ao sistema Gestão de farmácias que não efetuaram a renovação do credenciamento suspensa", "mac": "Gestão de insumos estratégicos aplicados à saúde", "cat": "Assistência farmacêutica e insumos estratégicos para o SUS", "serv": "Fomento à produção nacional de insumos estratégicos", "nat": "Finalístico"}, {"cod": "03800001", "ent": "Diagnóstico Situacional dos Núcleos de Econômica da Saúde/NES publicado", "mac": "Desenvolvimento e fortalecimento da economia da saúde", "cat": "Ações estruturantes em economia e desenvolvimento em saúde", "serv": "Emissão de declaração de reconhecimento de limites", "nat": "Finalístico"}, {"cod": "09690004", "ent": "Processo minerário de indeferimento analisado", "mac": "Gestão da politica de geologia, mineração e transformação mineral", "cat": "Regulação do setor mineral", "serv": "", "nat": "Finalístico"}, {"cod": "01670006", "ent": "Alteração de ato constitutivo analisada", "mac": "Fiscalização e conformidade regulatória do setor elétrico", "cat": "Fiscalização Econômica, Financeira e de Mercado do Setor Elétrico", "serv": "Compensação Previdenciária", "nat": "Finalístico"}, {"cod": "09630002", "ent": "Portaria publicada para efetivação da cessão da área comercial", "mac": "Gestão da política de infraestrutura aeroportuária", "cat": "Concessões aeroportuárias", "serv": "", "nat": "Finalístico"}, {"cod": "07800041", "ent": "Participação de indígenas em eventos de artesanatos apoiada", "mac": "Promover o desenvolvimento sustentável", "cat": "Etnodesenvolvimento", "serv": "Compensação Previdenciária", "nat": "Finalístico"}, {"cod": "12530056", "ent": "Estrutura de dados e do Painel de receitas e despesas definida", "mac": "Gestão da administração financeira e fiscal", "cat": "Planejamento e a Programação Financeira", "serv": "Atuação como órgão central de administração financeira", "nat": "Finalístico"}, {"cod": "06870020", "ent": "Painéis das Transferências voluntárias em execução e em prestação de contas acessivéis na internet e na intranet do MDHC  atualizado quinzenalmente", "mac": "Gestão da logística pública", "cat": "Gestão de pagamentos", "serv": "", "nat": "Suporte"}, {"cod": "03790271", "ent": "Visitas técnicas às unidades regionais realizadas", "mac": "Gestão de informações populacionais, econômicas e geográficas nacionais", "cat": "Produção, análise e disseminação de dados e informações institucionais", "serv": "", "nat": "Finalístico"}, {"cod": "02590461", "ent": "Acompanhamento das novas funcionalidades relativas às Informações Complementares de Pessoal e Benefícios realizadas.", "mac": "Desenvolver sistemas corporativos", "cat": "Gestão de sistemas informatizados", "serv": "", "nat": "Finalístico"}, {"cod": "11350001", "ent": "Bases para publicações no  Portal da Cidadania Tributária  sistematizadas", "mac": "Regulação e supervisão tributária e financeira", "cat": "Desenvolvimento de políticas de conformidade tributária", "serv": "", "nat": "Finalístico"}, {"cod": "01070251", "ent": "Cadastro de Requisição registrada no SIPAC", "mac": "Administrar serviços gerais", "cat": "Cadastro e registro de pessoal", "serv": "Fomento à produção nacional de insumos estratégicos", "nat": "Finalístico"}, {"cod": "01580002", "ent": "Cobrança de Inadimplência Realizada", "mac": "Gestão do crédito tributário, administrativo e arrecadação", "cat": "Arrecadação e inadimplência", "serv": "Gestão do Programa Farmácia Popular do Brasil", "nat": "Finalístico"}, {"cod": "06640016", "ent": "Agenda transversal ou prioridade do PPA analisada e marcada", "mac": "Gestão de políticas de energia elétrica", "cat": "Planejamento Plurianual da APF", "serv": "Gestão orçamentária das ações de atenção primária à saúde", "nat": "Finalístico"}, {"cod": "12770002", "ent": "Decisão da Certificação de Entidades Beneficentes de Assistência Social (CEBAS) publicada no DOU", "mac": "Gestão da política pública de assistência social", "cat": "Gestão da Rede Socioassistencial Privada do SUAS", "serv": "Certificação de Entidades Beneficentes de Assistência Social (CEBAS)", "nat": "Finalístico"}, {"cod": "01400764", "ent": "Termo de Recebimento Provisório validado", "mac": "Gestão da logística pública", "cat": "Gestão de contratos", "serv": "", "nat": "Suporte"}, {"cod": "10930012", "ent": "Acordos  fomentados e propostos", "mac": "Gestão integrada das políticas econômicas", "cat": "Análise de políticas econômicas", "serv": "", "nat": "Finalístico"}, {"cod": "11150009", "ent": "Prestadores de serviços turísticos fiscalizados", "mac": "Regulamentação e fiscalização da atividade do setor turístico", "cat": "Regulamentação e Fiscalização da Atividade do Setor Turístico", "serv": "", "nat": "Finalístico"}, {"cod": "10010002", "ent": "Plano de ação de Bioeconomia monitorado.", "mac": "Fomento ao desenvolvimento sustentável da bioeconomia", "cat": "Gestão das políticas públicas agropecuárias", "serv": "", "nat": "Finalístico"}, {"cod": "11740016", "ent": "Maquinários e materiais do Laboratório de Restauração higienizados", "mac": "Gestão da política e salvaguarda da memória  bibliográfica e documental do país", "cat": "Preservação e salvaguarda do acervo bibliográfico e documental nacional", "serv": "Gestão de laboratórios de restauração", "nat": "Finalístico"}, {"cod": "01590002", "ent": "Limites DEC e FEC estabelecidos", "mac": "Regulação do setor elétrico", "cat": "Regulação dos Serviços de Transmissão e Distribuição de Energia", "serv": "Gestão de Créditos Não Tributários", "nat": "Finalístico"}, {"cod": "02380061", "ent": "Interlocução institucional estabelecida", "mac": "Promoção da segurança jurídica", "cat": "Consultoria e assessoramento jurídico", "serv": "", "nat": "Finalístico"}, {"cod": "02310124", "ent": "Minuta de IN sobre manejo de produtos florestais não madeireiros elaborada", "mac": "Proteção", "cat": "Normatização", "serv": "Instrução de propostas de normas, orientação técnica, acompanhamento e execução de programas e ações relativas ao uso sustentável da flora", "nat": "Finalístico"}, {"cod": "10180011", "ent": "Projetos, planos de trabalho com foco no fomento às Boas Práticas Agropecuárias (BPA) elaborado", "mac": "Gestão da política de fomento à produção agropecuária", "cat": "Fomento à produção agropecuária", "serv": "", "nat": "Finalístico"}, {"cod": "12760002", "ent": "processo dee Certificação de Comunidades Remanescentes de Quilombos recebidos e analisadosra Garantia de Direitos Territoriais, culturais, sociais acesso a Politicas Públicas Certificação de Comunidades Quilombolas.", "mac": "Proteção às comunidades quilombolas", "cat": "Certificação de comunidades quilombolas", "serv": "Gestão de Políticas Quilombolas", "nat": "Finalístico"}, {"cod": "11480005", "ent": "Política Geral/Específica de Gestão de Riscos divulgada", "mac": "Gerenciar riscos corporativos", "cat": "Gestão de Riscos", "serv": "", "nat": "Finalístico"}, {"cod": "06110045", "ent": "Parecer de ateste dos requisitos para liberação dos recursos para obras e aquisição de equipamentos(instrumento-TED)", "mac": "Relações institucionais e federativas", "cat": "Gestão de parcerias institucionais", "serv": "", "nat": "Governança"}, {"cod": "00020004", "ent": "Despacho/Parecer sobre Relatório Anual de Avaliação do Programa de Gestão Orientada para Resultados elaborado", "mac": "Gerenciar desempenho de pessoas", "cat": "Gestão de desempenho individual", "serv": "Gestão da Autorização de Funcionamento de Organização Estrangeira em Território Brasileiro – OE", "nat": "Finalístico"}, {"cod": "13470010", "ent": "Relatório de Auditoria referente à BPF/APPCC elaborado", "mac": "Inspeção de produtos de origem animal e vegetal", "cat": "Fiscalização de estabelecimentos e produtos de origem animal e vegetal", "serv": "Auditoria e fiscalização de produtos de origem vegetal", "nat": "Finalístico"}, {"cod": "02470006", "ent": "Atualização cadastral de Dirigentes e Entidades realizada", "mac": "Licenciamento de entidades e planos de previdência complementar", "cat": "Licenciamento", "serv": "Suporte técnico-administrativo à gestão de processos consultivos quanto à adequação jurídica", "nat": "Finalístico"}, {"cod": "07820030", "ent": "\"Apoio e qualificação para a oferta de ações e políticas públicas referentes aos processos educativos comunitários e escolares dos povos indígenas em contexto urbano, incluindo as comunidades indígenas residentes no Distrito Federal realizados \"", "mac": "Promoção e proteção dos direitos e cidadania dos povos indígenas", "cat": "Promoção e Qualificação dos Processos Educativos Indígenas", "serv": "Gestão de recursos do Programa Mais Médicos para o Brasil", "nat": "Finalístico"}, {"cod": "10440012", "ent": "Documentos técnicos elaborados, pareceres emitidos e contribuições estratégicas produzidasl.", "mac": "Normatização e efetivação das políticas de comunicações", "cat": "Acompanhamento e avaliação de programas e ações de comunicações", "serv": "Diretrizes Estratégicas para Comunicação", "nat": "Finalístico"}, {"cod": "11170003", "ent": "Acompanhamento técnico de projetos de restauração ecológica e produtiva através de sistemas agroflorestais realizado", "mac": "Ações socioambientais e consolidação territorial em unidades de conservação", "cat": "Economias da Sociobiodiversidade", "serv": "Inclusão social e produtiva", "nat": "Finalístico"}, {"cod": "11210239", "ent": "Instrumentos de parcerias da Proteção Social Especial gerenciados", "mac": "Gestão da política pública de assistência social", "cat": "Gerir as ações relacionadas à Proteção Social Especial", "serv": "Monitoramento da execução dos serviços, programas, projetos e benefícios de Proteção Social Especial", "nat": "Finalístico"}, {"cod": "12670002", "ent": "Análise de Risco Fiscal de Empresa Estatal realizado", "mac": "Planejamento fiscal", "cat": "Riscos Fiscais", "serv": "Risco Fiscal das Estatais", "nat": "Finalístico"}, {"cod": "13590001", "ent": "Apoio a eventos de entrega de maquinário realizado", "mac": "Apoio à produção agropecuária", "cat": "Implementação de políticas de apoio à produção e comercialização", "serv": "", "nat": "Finalístico"}, {"cod": "09330031", "ent": "Mudanças implementadas", "mac": "Administrar infraestrutura de tecnologia da informação", "cat": "Gestão de infraestrutura de TIC", "serv": "", "nat": "Finalístico"}, {"cod": "12430002", "ent": "Manutenção Detacustos realizada", "mac": "Gestão contábil e de custos do setor público", "cat": "Execução da contabilidade pública e de custos na União", "serv": "Procedimentos de contabilidade de custos em soluções tecnológicas de apoio à geração de informações de custos", "nat": "Finalístico"}, {"cod": "06660018", "ent": "Ações e programas passíveis de serem incluídos no Plano Nacional dos Direitos da Pessoa com Deficiência identificados", "mac": "Defesa dos direitos", "cat": "Gestão de políticas públicas", "serv": "", "nat": "Finalístico"}, {"cod": "09200001", "ent": "Sistema Balcão gov.br gerido", "mac": "Gestão de registros administrativos e de informações estatísticas do mercado de trabalho", "cat": "Gestão da emissão da carteira de trabalho e previdência social", "serv": "", "nat": "Finalístico"}, {"cod": "06040033", "ent": "Orientação de estágios provenientes de convênios internacionais (graduação e pós-graduação) realizada", "mac": "Educação superior", "cat": "Docência - Magistério Federal", "serv": "Compensação Previdenciária", "nat": "Finalístico"}, {"cod": "06120486", "ent": "Participação em iniciativas relacionadas à Estratégia Nacional da Infraestrutura da Qualidade (ENIQ) e nas reuniões do Comitê Técnico de Assessoramento ad hoc de Infraestrutura da Qualidade (CTIQ)", "mac": "Relações institucionais e federativas", "cat": "Articulação institucional", "serv": "", "nat": "Governança"}, {"cod": "10010001", "ent": "Implementação do programa nacional de bioinsumos acompanhadas", "mac": "Fomento ao desenvolvimento sustentável da bioeconomia", "cat": "Gestão das políticas públicas agropecuárias", "serv": "", "nat": "Finalístico"}, {"cod": "06660004", "ent": "PNS elaborado", "mac": "Defesa dos direitos", "cat": "Gestão de políticas públicas", "serv": "", "nat": "Finalístico"}, {"cod": "01500374", "ent": "Materiais patrimoniados a receber e/ou a devolver conforme autorização da Coordenadora-Geral da SE-CNDH solicitados", "mac": "Administrar bens de consumo e permanente", "cat": "Gestão Patrimonial", "serv": "Gerenciamento de bens e patrimônios", "nat": "Finalístico"}, {"cod": "02510631", "ent": "Atendimentos presenciais realizados", "mac": "Gestão dos créditos e defesa do interesse público", "cat": "Atendimento", "serv": "", "nat": "Finalístico"}, {"cod": "11170001", "ent": "Normativa para o ordenamento pesqueiro nas UC instituído", "mac": "Ações socioambientais e consolidação territorial em unidades de conservação", "cat": "Economias da Sociobiodiversidade", "serv": "Gestão da pesca artesanal", "nat": "Finalístico"}, {"cod": "06020104", "ent": "Plano de capacitação elaborado", "mac": "Gerenciar manifestações de ouvidoria", "cat": "Ouvidoria", "serv": "", "nat": "Finalístico"}, {"cod": "01230178", "ent": "Planejamento Contábil Realizado", "mac": "Administrar recursos financeiros", "cat": "Contabilidade", "serv": "", "nat": "Finalístico"}, {"cod": "01011207", "ent": "Relatório final de curso no Moodle emitido", "mac": "Gestão de pessoas", "cat": "Desenvolvimento e capacitação", "serv": "", "nat": "Suporte"}, {"cod": "10920029", "ent": "Estudo preliminar  de projeto relacionado ao aperfeiçoamento da Política Agrícola Nacional (PAN) realizado", "mac": "Formulação, acompanhamento e avaliação de políticas públicas", "cat": "Gestão de políticas de promoção de eficiência econômica​", "serv": "", "nat": "Finalístico"}, {"cod": "10390007", "ent": "Prestação de contas de instrumentos de parceria analisada", "mac": "Diretrizes estratégicas para comunicações", "cat": "Articulação e Parceria para Viabilização de Políticas e Comunicações", "serv": "", "nat": "Finalístico"}, {"cod": "08450084", "ent": "Pesquisas e relatórios sobre pirataria e delitos contra a Propriedade Intelectual  produzidos", "mac": "Justiça e cidadania", "cat": "Propriedade Intelectual", "serv": "", "nat": "Finalístico"}, {"cod": "02310059", "ent": "Ofícios e nota técnicas referentes à alterações na estrutura organizacional da Secretaria elaboradas", "mac": "Gestão de políticas de energia elétrica", "cat": "Normatização", "serv": "Gestão dos Contratos de Bens, Serviços e Insumos de Saúde Indígena", "nat": "Finalístico"}, {"cod": "02500008", "ent": "Programas vigentes", "mac": "Gerenciar programas e projetos", "cat": "Gestão de programas", "serv": "", "nat": "Finalístico"}, {"cod": "09250003", "ent": "Contratos do abono salarial mensalmente geridos", "mac": "Gestão de benefícios trabalhistas", "cat": "Gestão do abono salarial", "serv": "", "nat": "Finalístico"}, {"cod": "10430027", "ent": "Proposta de alteração dos Planos Básicos de Distribuição de Canais encaminhada", "mac": "Monitorar a gestão da comunicação institucional", "cat": "Gestão dos Serviços de Radiodifusão", "serv": "", "nat": "Finalístico"}, {"cod": "01140022", "ent": "Homologação de Programa de Gerenciamento de Risco-PGR  das empresas terceirizadas da instituição realizada.", "mac": "Promover saúde, segurança e qualidade de vida", "cat": "Segurança do Trabalho", "serv": "", "nat": "Finalístico"}, {"cod": "03200285", "ent": "Participação em reunião de grupo de trabalho, comitês, colegiados, sindicâncias e comissões registrada", "mac": "Monitorar a gestão do desenvolvimento organizacional e da inovação", "cat": "Grupos de Trabalho e Comitês", "serv": "Representação institucional", "nat": "Finalístico"}, {"cod": "05010080", "ent": "Estrutura de governança internacional migratória mapeada", "mac": "Gestão de políticas de justiça e cidadania", "cat": "Políticas migratórias", "serv": "INATIVAR - incompreensível", "nat": "Finalístico"}, {"cod": "08670002", "ent": "Fluxo de operações especiais realizado", "mac": "Enfrentamento e prevenção à corrupção", "cat": "Ações investigativas", "serv": "", "nat": "Finalístico"}, {"cod": "06750002", "ent": "Conteúdo para redes sociais sobre o curso da vida divulgado", "mac": "Comunicação institucional", "cat": "Publicidade e mídias institucionais", "serv": "", "nat": "Governança"}, {"cod": "05010055", "ent": "Parecer de processo de naturalização analisado", "mac": "Gestão de políticas de justiça e cidadania", "cat": "Políticas migratórias", "serv": "INATIVAR - incompreensível", "nat": "Finalístico"}, {"cod": "08450083", "ent": "Reunião do Conselho Nacional de Combate à Pirataria secretariada", "mac": "Justiça e cidadania", "cat": "Propriedade Intelectual", "serv": "", "nat": "Finalístico"}, {"cod": "02060019", "ent": "Manual de fiscalização revisado", "mac": "Gestão de políticas de transporte aquaviário", "cat": "Transporte Aquaviário – Fiscalização", "serv": "Monitoramento do mapeamento e gerência dos processos", "nat": "Finalístico"}, {"cod": "13140003", "ent": "Vítimas atendidas", "mac": "Cultura", "cat": "Combate à Tortura", "serv": "Acompanhamento de atendimento a vítimas de tortura", "nat": "Finalístico"}, {"cod": "13120001", "ent": "Recepção e execução dos procedimentos padronizados de atendimento realizados", "mac": "Prestação de assistência jurídica integral e gratuita, individual e coletiva", "cat": "Atendimento inicial ao público para prestação de assistência jurídica", "serv": "Gestão do atendimento inicial ao público", "nat": "Finalístico"}, {"cod": "13590002", "ent": "Site do PROMAQ atualizado", "mac": "Apoio à produção agropecuária", "cat": "Implementação de políticas de apoio à produção e comercialização", "serv": "", "nat": "Finalístico"}, {"cod": "10410007", "ent": "Indicadores de políticas públicas monitorados", "mac": "Administrar suporte técnico", "cat": "Inclusão Digital e Acesso à Informação", "serv": "Expansão e Massificação das Comunicações", "nat": "Finalístico"}, {"cod": "12240195", "ent": "Novo painel preparado", "mac": "Transparência das contas públicas", "cat": "Informações e relatórios contábeis, orçamentários e fiscais da União", "serv": "Divulgação de informações contábeis e fiscais da Federação", "nat": "Finalístico"}, {"cod": "01400803", "ent": "Pagamentos dos contratos vigentes efetivados", "mac": "Gerenciar contratações", "cat": "Gestão de contratos", "serv": "Gestão e Fiscalização de Contratos", "nat": "Finalístico"}, {"cod": "13060002", "ent": "Sistema de ciência documental desenvolvido, testado, homologado e implantado", "mac": "Gestão do processo administrativo digital", "cat": "GERIR ATIVIDADES DE PROTOCOLO, TRIAGEM E PREPARO DE PROCESSO ADMINISTRATIVO DIGITAL", "serv": "CONTROLAR A CONFIGURAÇÃO E PADRONIZAÇÃO DOS SISTEMAS DE CONTROLE PROCESSUAL", "nat": "Finalístico"}, {"cod": "10930005", "ent": "Monitoramento de política econômica realizado", "mac": "Gestão integrada das políticas econômicas", "cat": "Análise de políticas econômicas", "serv": "", "nat": "Finalístico"}, {"cod": "11630001", "ent": "Estudos e pesquisas técnicas sobre tecnologias e seus impactos na proteção de dados e privacidade elaboradas", "mac": "Gestão da política de proteção de dados pessoais", "cat": "Tecnologia e Pesquisa em Proteção de Dados Pessoais", "serv": "", "nat": "Finalístico"}, {"cod": "11790019", "ent": "Relatórios físico financeiros organizados", "mac": "Planejamento e gestão de recursos externos", "cat": "Gestão de Compensação Ambiental", "serv": "Gestão Operacional Termos de Compromissos com Empreendedores", "nat": "Finalístico"}, {"cod": "09770023", "ent": "Relatório técnico com análise comparativa dos indicadores frente a metas, padrões de referência e evolução histórica detalhado", "mac": "Gestão de política de exploração e produção de petróleo e gás natural", "cat": "Promoção da sustentabilidade", "serv": "", "nat": "Finalístico"}, {"cod": "02500048", "ent": "Representação da IFES junto à Secretaria de Educação Superior do MEC; Acompanhamento e execução do fluxo administrativo do PET na IFES, realizados", "mac": "Gerenciar programas e projetos", "cat": "Gestão de programas", "serv": "", "nat": "Finalístico"}, {"cod": "11760015", "ent": "Apoio aos estados da região nordeste do Brasil na realização do serviço de acompanhamento familiar para inclusão produtiva rural (SAFISP) por meio de TED", "mac": "Gestão de políticas públicas de segurança alimentar e nutricional", "cat": "Promoção de ações de inclusão produtiva rural", "serv": "Formalização de TED para apoiar os estados da região nordeste do Brasil na realização do serviço de acompanhamento familiar para inclusão produtiva rural", "nat": "Finalístico"}, {"cod": "02300078", "ent": "Realização de Pesquisa", "mac": "Consultoria e assessoramento jurídico", "cat": "Consultoria jurídica institucional", "serv": "", "nat": "Finalístico"}, {"cod": "12480012", "ent": "Conformidade de Registro de Gestão realizada", "mac": "Financiamento público e operações de fomento e subvenções", "cat": "Provimento de subvenções a operações de crédito", "serv": "Acompanhar e controlar os programas de financiamento e subvenção", "nat": "Finalístico"}, {"cod": "14010001", "ent": "Analise de indicadores realizada", "mac": "Gestão pública orientada por evidências", "cat": "Gestão de indicadores", "serv": "Levantamento de indicadores", "nat": "Finalístico"}, {"cod": "11270014", "ent": "Procuradorias-Regionais, quanto a ferramentas de transação, atendidas", "mac": "Gestão integrada da representação e defesa da fazenda nacional", "cat": "Atuação na mediação e transação tributária para redução de litígios", "serv": "", "nat": "Finalístico"}, {"cod": "10390001", "ent": "Processo de planos de aplicação de recursos analisado", "mac": "Diretrizes estratégicas para comunicações", "cat": "Articulação e Parceria para Viabilização de Políticas e Comunicações", "serv": "Diretrizes Estratégicas para Comunicações", "nat": "Finalístico"}, {"cod": "06880255", "ent": "Demandas dos GGEs/DDRs captadas, analisadas e registradas", "mac": "Gestão da política de sustentabilidade, inovação e qualidade da produção agropecuária", "cat": "Gestão de parcerias institucionais", "serv": "", "nat": "Finalístico"}, {"cod": "07930021", "ent": "Informações territoriais das terras indígenas jurisdicionadas aglutinadas", "mac": "Proteção de terras indígenas e de povos isolados e de recente contato", "cat": "Gerir informação territorial", "serv": "Compensação Previdenciária", "nat": "Finalístico"}, {"cod": "09450012", "ent": "Termos de Execução Descentralizada (TED) gerenciados e fiscalizados", "mac": "Gestão da política da aviação civil", "cat": "Planejamento, monitoramento e avaliação do Setor de Aviação Civil", "serv": "", "nat": "Finalístico"}, {"cod": "09590004", "ent": "Complementações na legislação de navegação interior proposta", "mac": "Gestão da política de navegação marítima", "cat": "Gestão da Política de Navegação interior", "serv": "", "nat": "Finalístico"}, {"cod": "02310234", "ent": "Notas Técnicas de projetos de Lei elaboradas", "mac": "Gestão de políticas públicas", "cat": "Normatização", "serv": "Assessoramento Jurídico em projetos de leis e normativos sobre Atenção à Saúde Integral", "nat": "Finalístico"}, {"cod": "10200025", "ent": "Apresentação técnica na Reunião Climática da Região Nordeste realizada", "mac": "Gestão de informações meteorológicas nacionais", "cat": "Meteorologia Geral e Aplicada", "serv": "", "nat": "Finalístico"}, {"cod": "01030036", "ent": "Instrução de processo à CGPLAN para verificação realizada", "mac": "Gerenciar informações cadastrais de pessoal", "cat": "RH - Monitoramento e Controle", "serv": "", "nat": "Finalístico"}, {"cod": "12600004", "ent": "Estudo de impacto sobre proposta de acordo e convênios elaborado", "mac": "Gestão da política de desenvolvimento de atividades espaciais", "cat": "Fomento e comercialização no setor espacial", "serv": "Apoio à celebração de acordos de cessão de uso de infraestrutura espacial pública", "nat": "Finalístico"}, {"cod": "02310133", "ent": "Nota técnica elaborada", "mac": "Gestão de trabalho em saúde", "cat": "Normatização", "serv": "Gestão dos Contratos de Bens, Serviços e Insumos de Saúde Indígena", "nat": "Finalístico"}], "semServico": [{"cod": "03200342", "ent": "Outras deliberações do Conselho Diretor geridas", "mac": "Monitorar a gestão do desenvolvimento organizacional e da inovação", "cat": "Grupos de Trabalho e Comitês", "serv": "", "nat": "Finalístico"}, {"cod": "01150323", "ent": "Certificação de disponibilidade financeira solicitada", "mac": "Gestão de transferências e repasses de recursos da União", "cat": "Gestão das transferências voluntárias", "serv": "", "nat": "Finalístico"}, {"cod": "06630107", "ent": "Participação nas reuniões ou nos projetos propostos da Comissão Nacional do Risco de Fauna (CNRF), vinculada ao Comitê Nacional de Prevenção de Acidentes Aeronáuticos (CNPAA) realizadas.", "mac": "Relações institucionais e federativas", "cat": "Representação institucional", "serv": "", "nat": "Governança"}, {"cod": "01210259", "ent": "Propostas Fundo a Fundo de   recursos de Emenda Parlamentar/Programa pré-classificadas", "mac": "Planejamento e orçamento", "cat": "Orçamento", "serv": "", "nat": "Governança"}, {"cod": "11120014", "ent": "Plano de mídia aprovado e veiculação autorizada", "mac": "Desenvolvimento estratégico do setor turístico", "cat": "Promoção, Estruturação e Qualificação de Produtos, Serviços e Destinos Turísticos", "serv": "", "nat": "Finalístico"}, {"cod": "06120509", "ent": "Despachos diários com a direção do departamento realizados", "mac": "Relações institucionais e federativas", "cat": "Articulação institucional", "serv": "", "nat": "Governança"}, {"cod": "01080064", "ent": "Contratação de estagiários na modalidade não obrigatório realizada", "mac": "Gestão de pessoas", "cat": "Recrutamento e seleção de pessoal", "serv": "", "nat": "Suporte"}, {"cod": "08400033", "ent": "Reunião realizada com Euroclima e GIZ para discutir escopo da cooperação realizada", "mac": "Economia verde, descarbonização e bioindústria", "cat": "Descarbonização", "serv": "", "nat": "Finalístico"}, {"cod": "06010224", "ent": "Conteúdo da página revisado periodicamente", "mac": "Comunicação institucional", "cat": "Comunicação institucional", "serv": "", "nat": "Governança"}, {"cod": "06120108", "ent": "Articulação intersetorial com as demais secretarias do Ministério da Saúde sobre o Pacto Nacional da Transmissão Vertical realizada", "mac": "Relações institucionais e federativas", "cat": "Articulação institucional", "serv": "", "nat": "Governança"}, {"cod": "01010047", "ent": "Informação sobre atributos de cargos, carreiras e planos Especiais  fornecida a área demandante", "mac": "Gestão de pessoas", "cat": "Desenvolvimento e capacitação", "serv": "", "nat": "Suporte"}, {"cod": "05050067", "ent": "Estudos relacionados a OSINT realizados", "mac": "Segurança pública", "cat": "Combate à corrupção e à lavagem de dinheiro", "serv": "", "nat": "Finalístico"}, {"cod": "02590421", "ent": "Responder as dúvidas que surgem dos executores do programa de emissão de certificado de origem", "mac": "Gestão de pessoas", "cat": "Gestão de sistemas informatizados", "serv": "", "nat": "Suporte"}, {"cod": "11600045", "ent": "Termo aditivo ao Contrato  de prestação de serviços entre o Ministério do Desenvolvimento Social, Família e Combate à Fome e a CAIXA para ações de transferência direta de renda assinado", "mac": "Gestão da transferência condicionada e direta de renda", "cat": "Gestão de contratos de prestação de serviços das ações de transferência direta de renda", "serv": "", "nat": "Finalístico"}, {"cod": "01700291", "ent": "Apresentação padrão dos resultados dos indicadores calculados para o monitoramento dos exames de biologia molecular para clamídia e gnococo atualizada", "mac": "Gestão do conhecimento e da informação corporativa", "cat": "Gestão documental e arquivística", "serv": "", "nat": "Governança"}, {"cod": "01070346", "ent": "Registro nos sistemas governamentais e gerenciais das movimentações, cessões, requisições, reconduções; remoções; redistribuições; licenças ou afastamentos concedidos aos servidores realizado", "mac": "Gestão de pessoas", "cat": "Cadastro e registro de pessoal", "serv": "", "nat": "Suporte"}, {"cod": "01220114", "ent": "Valores das Decisões Judiciais homologados", "mac": "Gestão financeira, contábil e de custos", "cat": "Gestão de Pagamentos", "serv": "", "nat": "Suporte"}, {"cod": "08720020", "ent": "Rondas e inspeções de segurança em áreas e instalações consideradas sensíveis realizada", "mac": "Gestão da logística pública", "cat": "Atividades de inteligência e segurança institucional", "serv": "", "nat": "Suporte"}, {"cod": "09140046", "ent": "Demandas de saneamento processual enviadas pelo SEMUR no sistema eletrônico de processos (CPMR) cumpridas", "mac": "Inspeção do trabalho", "cat": "Fiscalização do trabalho", "serv": "", "nat": "Finalístico"}, {"cod": "01230027", "ent": "Seguro garantia e caução registrada/baixada", "mac": "Gestão financeira, contábil e de custos", "cat": "Contabilidade", "serv": "", "nat": "Suporte"}, {"cod": "01210681", "ent": "Solução de Melhorias e Correções para o SIOP homologada", "mac": "Planejamento e orçamento", "cat": "Orçamento", "serv": "", "nat": "Governança"}, {"cod": "03200366", "ent": "Documentos e subsídios para o Clube de Paris preparados.", "mac": "Relações institucionais e federativas", "cat": "Grupos de Trabalho e Comitês", "serv": "", "nat": "Governança"}, {"cod": "10140006", "ent": "Informações sobre o setor de gás natural no mercado internacional solicitadas", "mac": "Gestão da política de petróleo, gás natural e biocombustíveis", "cat": "Gestão da política nacional de gás natural", "serv": "", "nat": "Finalístico"}, {"cod": "11480015", "ent": "Relatório da Avaliação da Cultura de Riscos e Controles concluído", "mac": "Gerenciar riscos corporativos", "cat": "Gestão de Riscos", "serv": "", "nat": "Finalístico"}, {"cod": "01400615", "ent": "Processo de solicitação de renovação contratual enviado", "mac": "Gestão da logística pública", "cat": "Gestão de contratos", "serv": "", "nat": "Suporte"}, {"cod": "01910047", "ent": "Atividades de capacitação em Desenho Industrial aplicadas", "mac": "Atividade econômica", "cat": "Propriedade Industrial - Desenhos Industriais", "serv": "", "nat": "Finalístico"}, {"cod": "01010157", "ent": "Planejamento de Plano de Trabalho em curso", "mac": "Gestão de pessoas", "cat": "Desenvolvimento e capacitação", "serv": "", "nat": "Suporte"}, {"cod": "01010225", "ent": "Ações do seminário de integração ofertadas", "mac": "Gestão de pessoas", "cat": "Desenvolvimento e capacitação", "serv": "", "nat": "Suporte"}, {"cod": "10180008", "ent": "Recuperação da vegetação nativa e recomposição florestal apoiados", "mac": "Gestão da política de fomento à produção agropecuária", "cat": "Fomento à produção agropecuária", "serv": "", "nat": "Finalístico"}, {"cod": "09050001", "ent": "Propostas de alteração em resoluções de órgãos colegiados elaboradas", "mac": "Promoção à inserção e permanência no trabalho", "cat": "Gestão das ações e serviços sine", "serv": "", "nat": "Finalístico"}, {"cod": "01010465", "ent": "Processo de Reconhecimento de Saberes e Competências analisado", "mac": "Gestão de pessoas", "cat": "Desenvolvimento e capacitação", "serv": "", "nat": "Suporte"}, {"cod": "02510319", "ent": "Atendimento presencial realizado", "mac": "Relações institucionais e federativas", "cat": "Atendimento", "serv": "", "nat": "Governança"}, {"cod": "08400007", "ent": "Gerir o Hub de Descarbonizaçao da Indústria Brasileira (HDIB) - Cooperaçao internacional com governos britãnico e alemão (entre outros)  para assistência internacional no apoio à concretização das ambições de descarbonização industrial e de industrialização verde do Brasil.", "mac": "Economia verde, descarbonização e bioindústria", "cat": "Descarbonização", "serv": "", "nat": "Finalístico"}, {"cod": "01490030", "ent": "Gêneros alimentícios armazenados e distribuídos", "mac": "Gestão da logística pública", "cat": "Planejamento logístico", "serv": "", "nat": "Suporte"}, {"cod": "02380007", "ent": "Demandas de assessoramento jurídico das Procuradorias-Gerais Adjuntas coordenadas", "mac": "Gestão jurídica institucional", "cat": "Consultoria e assessoramento jurídico", "serv": "", "nat": "Governança"}, {"cod": "09800005", "ent": "Atividades da COP relacionadas ao setor mineral coordenadas", "mac": "Gestão do planejamento setorial de mineração", "cat": "Desenvolvimento socioambiental do setor mineral", "serv": "", "nat": "Finalístico"}, {"cod": "01400535", "ent": "Desenvolvimento Técnico na Área de Contratações de Serviços, Bens e Insumos de Saúde Indígena realizados", "mac": "Gestão da logística pública", "cat": "Gestão de contratos", "serv": "", "nat": "Suporte"}, {"cod": "02460005", "ent": "Componente de Risco e Gestão do Risco Atuarial do Manual da Supervisão Permanente analisado", "mac": "Avaliação e controle", "cat": "Fiscalização/Supervisão", "serv": "", "nat": "Governança"}, {"cod": "01030017", "ent": "Informação sobre recolhimento de PSS encaminhado à CGOF", "mac": "Avaliação e controle", "cat": "RH - Monitoramento e Controle", "serv": "", "nat": "Governança"}, {"cod": "01070366", "ent": "Relatório das demandas específicas de relatórios de Cargos em Comissão e das Funções de Confiança, Funções Comissionadas Técnicas e Gratificações Temporárias elaborado", "mac": "Gestão de pessoas", "cat": "Cadastro e registro de pessoal", "serv": "", "nat": "Suporte"}, {"cod": "01150166", "ent": "Autorização para celebração do instrumento solicitada", "mac": "Gestão de transferências e repasses de recursos da União", "cat": "Gestão das transferências voluntárias", "serv": "", "nat": "Finalístico"}, {"cod": "06880120", "ent": "Comitê de Impacto do Estado do Rio Grande do Norte Instituído", "mac": "Relações institucionais e federativas", "cat": "Gestão de parcerias institucionais", "serv": "", "nat": "Governança"}, {"cod": "06570008", "ent": "Alinhamento para a gestão de ações prioritárias realizado", "mac": "Participação e controle social", "cat": "Controle social na saúde indígena", "serv": "", "nat": "Governança"}, {"cod": "11250038", "ent": "Averbação pré-executória realizada", "mac": "Gestão da dívida ativa", "cat": "Gestão dos Créditos e Defesa do Interesse Público", "serv": "", "nat": "Finalístico"}, {"cod": "06010225", "ent": "Fornecer subsídios técnicos para comunicação e divulgação das ações de enfrentamento das zoonoses, acidentes causados por animais peçonhentos e doenças de transmissão hídrica e alimentar", "mac": "Comunicação institucional", "cat": "Comunicação institucional", "serv": "", "nat": "Governança"}, {"cod": "01080138", "ent": "Profissionais contratados que compõem o quadro de profissionais do DGCI no MS", "mac": "Gestão de pessoas", "cat": "Recrutamento e seleção de pessoal", "serv": "", "nat": "Suporte"}, {"cod": "10160005", "ent": "Viabilidade de interligações internacionais definidas", "mac": "Gestão da política de transição energética e planejamento", "cat": "Outorgas transmissão e distribuição e Planejamento da Transmissão", "serv": "", "nat": "Finalístico"}, {"cod": "09960003", "ent": "Relatórios relacionados ao setor energético para envio a órgãos internacionais produzidos", "mac": "Gestão da política de transição energética e planejamento", "cat": "Acesso à informação", "serv": "", "nat": "Finalístico"}, {"cod": "09410085", "ent": "Participação ativa nas Oficinas de Estudos Temáticos da Estratégia Brasil 2050 (EBR2050), realizadas pelo Ministério da Gestão e da Inovação em Serviços Públicos (MGI)", "mac": "Gestão da política de transição energética e planejamento", "cat": "Planejamento energético", "serv": "", "nat": "Finalístico"}, {"cod": "09680145", "ent": "Materiais técnicos apresentados às instâncias decisórias, com atas, notas de reunião e recomendações de replanejamento ou reforço de ações estratégicas.", "mac": "Gestão estratégica e governança", "cat": "Planejamento, monitoramento e avaliação", "serv": "", "nat": "Governança"}, {"cod": "01050291", "ent": "Cadastro das informações pessoais e funcionais dos servidores ativos nos sistemas estruturantes, efetuado.", "mac": "Gestão de pessoas", "cat": "Pagamento de Pessoal", "serv": "", "nat": "Suporte"}, {"cod": "03500208", "ent": "Subsídios técnicos para definição da pauta de insumos utilizados no diagnóstico laboratorial da tuberculose, micoses endêmicas e micobactérias não tuberculosas elaborados", "mac": "Gestão do conhecimento e da informação corporativa", "cat": "Pesquisa, estudos e estatísticas", "serv": "", "nat": "Governança"}, {"cod": "02500003", "ent": "Relatório com análise do indicador de desempenho elaborado", "mac": "Gestão do conhecimento e da informação corporativa", "cat": "Gestão de programas", "serv": "", "nat": "Governança"}, {"cod": "11120034", "ent": "Ações do Projeto Brasil, Turismo Responsável Realizadas (Desenvolvimento do turismo responsável em comunidades indígenas e evento bianual Encontro de Turismo Responsável)", "mac": "Desenvolvimento estratégico do setor turístico", "cat": "Promoção, Estruturação e Qualificação de Produtos, Serviços e Destinos Turísticos", "serv": "", "nat": "Finalístico"}, {"cod": "11250006", "ent": "Trilhas de detecção de não conformidade desenvolvidas", "mac": "Gestão da dívida ativa", "cat": "Gestão dos Créditos e Defesa do Interesse Público", "serv": "", "nat": "Finalístico"}, {"cod": "02570083", "ent": "Dados sobre as despesas previstas e comprometidas das unidades do Ministério solicitados", "mac": "Gestão do conhecimento e da informação corporativa", "cat": "Gestão da informação corporativa", "serv": "", "nat": "Governança"}, {"cod": "09350021", "ent": "Exposições dos artisticas  residentes, nos espaços  da instituição, realizadas.", "mac": "Preservação e difusão do patrimônio histórico, artístico e cultural", "cat": "Preservação dos acervos históricos, culturais e artísticos", "serv": "", "nat": "Finalístico"}, {"cod": "01050100", "ent": "Informações prestadas", "mac": "Gestão de pessoas", "cat": "Pagamento de Pessoal", "serv": "", "nat": "Suporte"}, {"cod": "06120538", "ent": "Desenvolvimento da cadeia de valor de minerais estratégicos no Brasil articulados", "mac": "Gestão da politica de geologia, mineração e transformação mineral", "cat": "Articulação institucional", "serv": "", "nat": "Finalístico"}, {"cod": "01010311", "ent": "Bolsistas de apoio técnico acompanhados", "mac": "Gestão de pessoas", "cat": "Desenvolvimento e capacitação", "serv": "", "nat": "Suporte"}]};
const PAIRS_SEED = [{"id1": 526, "cod1": 8620034, "st1": "Ativa", "e1": "Denúncia de denúncia de indícios de irregularidades nas importações para fins de licenciamento não automático analisada", "id2": 563, "cod2": 8620042, "st2": "Ativa", "e2": "Validação da análise de denúncia de indícios de irregularidades nas importações para fins de licenciamento não automático realizada", "sim": 0.9709, "macro": "Comércio exterior", "cat": "Análise de dados operacionais e padronização de procedimentos", "serv": "", "sin": "analisado: analisada <> analise", "sug": "Validação da análise de denúncia de indícios de irregularidades nas importações para fins de licenciamento não automático realizada"}, {"id1": 578, "cod1": 10930002, "st1": "Ativa", "e1": "Pleitos de política econômica analisados", "id2": 579, "cod2": 10930001, "st2": "Ativa", "e2": "Política econômica avaliada", "sim": 0.979, "macro": "Gestão integrada das políticas econômicas", "cat": "Análise de políticas econômicas", "serv": "", "sin": "analisado: analisados <> avaliada", "sug": "Pleitos de política econômica analisados"}, {"id1": 595, "cod1": 1020094, "st1": "Ativa", "e1": "Minuta de alteração de normativos elaborada", "id2": 675, "cod2": 1020019, "st2": "Ativa", "e2": "Atos normativos elaborados", "sim": 0.9749, "macro": "Gestão de pessoas", "cat": "Análise e aplicação da legislação de pessoal", "serv": "Produção de conteúdo sobre provimento de pessoal na APF", "sin": "elaborado: elaborada <> elaborados", "sug": ""}, {"id1": 795, "cod1": 8560005, "st1": "Ativa", "e1": "Atas de reuniões do Grupo de Trabalho de Revisão da Tarifa Externa Comum (GT TEC) elaboradas", "id2": 811, "cod2": 8560004, "st2": "Ativa", "e2": "Reunião do Grupo de Trabalho de Revisão da Tarifa Externa Comum (GT TEC) convocada", "sim": 0.9708, "macro": "Comércio exterior", "cat": "Análise política comercial", "serv": "Análise Política Comercial", "sin": "reuniao: reunioes <> reuniao", "sug": ""}, {"id1": 835, "cod1": 13180001, "st1": "Ativa", "e1": "Encaminhamento à rede socioassistencial, de saúde e órgãos locais de proteção realizado.", "id2": 843, "cod2": 13180011, "st2": "Ativa", "e2": "Encaminhamentos e solicitações à rede socioassistencial realizados", "sim": 0.9736, "macro": "Atuar na defesa e promoção dos direitos humanos", "cat": "Aperfeiçoar o diálogo externo", "serv": "Articulação e encaminhamento à rede socioassistencial, de saúde e de proteção", "sin": "realizado: realizado <> realizados", "sug": ""}, {"id1": 888, "cod1": 11800017, "st1": "Ativa", "e1": "Oficinas e demais encontros presenciais com gestores, servidores e colaboradores dos estados e municípios da gestão do Programa Bolsa Família realizada", "id2": 889, "cod2": 11800018, "st2": "Ativa", "e2": "Lives públicas com gestores, servidores e colaboradores dos estados e municípios da gestão do Programa Bolsa Família realizadas", "sim": 0.9799, "macro": "Gestão da transferência condicionada e direta de renda", "cat": "Apoio à gestão descentralizada do Programa Bolsa Família", "serv": "Realização de eventos/encontros sobre programa Bolsa Família", "sin": "realizado: realizada <> realizadas", "sug": ""}, {"id1": 1350, "cod1": 6120124, "st1": "Ativa", "e1": "Gestão da execução de Cooperação Técnica Nacional e Internacional  no âmbito da vigilância da covid-19, influenza e outros vírus respiratórios realizada", "id2": 1442, "cod2": 6120123, "st2": "Ativa", "e2": "Planejamento de Cooperação Técnica Nacional e Internacional no âmbito da vigilânciada covid-19, influenza e outros vírus respiratórios realizado", "sim": 0.9799, "macro": "Relações institucionais e federativas", "cat": "Articulação institucional", "serv": "", "sin": "realizado: realizada <> realizado", "sug": ""}, {"id1": 1601, "cod1": 6120646, "st1": "Ativa", "e1": "Fórum de Gestores e Gestoras de Políticas para Pessoas com Deficiência realizado", "id2": 1609, "cod2": 6120327, "st2": "Ativa", "e2": "Gestão do Fórum de Gestores e Gestoras de Políticas para Pessoas com Deficiência realizada", "sim": 0.971, "macro": "Relações institucionais e federativas", "cat": "Articulação institucional", "serv": "Articulação institucional", "sin": "realizado: realizado <> realizada", "sug": "Gestão do Fórum de Gestores e Gestoras de Políticas para Pessoas com Deficiência realizada"}, {"id1": 1601, "cod1": 6120646, "st1": "Ativa", "e1": "Fórum de Gestores e Gestoras de Políticas para Pessoas com Deficiência realizado", "id2": 1685, "cod2": 6120647, "st2": "Ativa", "e2": "Gestão do Fórum de Gestores e Gestoras de Políticas para Pessoas com Deficiência realizada ", "sim": 0.9863, "macro": "Relações institucionais e federativas", "cat": "Articulação institucional", "serv": "Articulação institucional", "sin": "realizado: realizado <> realizada", "sug": "Gestão do Fórum de Gestores e Gestoras de Políticas para Pessoas com Deficiência realizada "}, {"id1": 2284, "cod1": 3760031, "st1": "Ativa", "e1": "Monitoramento do acesso aos medicamentos e insumos estratégicos no âmbito do Componente Especializado da Assistência Farmacêutica (CEAF) realizado", "id2": 2311, "cod2": 3760016, "st2": "Ativa", "e2": "Análise técnica dos medicamentos incorporados no âmbito do Componente Especializado da Assistência Farmacêutica (CEAF) realizada", "sim": 0.9725, "macro": "Gestão de insumos estratégicos aplicados à saúde", "cat": "Assistência farmacêutica e insumos estratégicos para o SUS", "serv": "", "sin": "realizado: realizado <> realizada", "sug": ""}, {"id1": 2401, "cod1": 2510247, "st1": "Ativa", "e1": "Atendimento telefônico realizado", "id2": 2757, "cod2": 2510791, "st2": "Ativa", "e2": "Realização de  atendimento telefônico", "sim": 0.9799, "macro": "Relações institucionais e federativas", "cat": "Atendimento", "serv": "", "sin": "realizado: realizado <> realizacao", "sug": "Realização de  atendimento telefônico"}, {"id1": 2401, "cod1": 2510247, "st1": "Ativa", "e1": "Atendimento telefônico realizado", "id2": 2838, "cod2": 2510842, "st2": "Ativa", "e2": "Realização de  atendimento telefônico realizada", "sim": 0.9804, "macro": "Relações institucionais e federativas", "cat": "Atendimento", "serv": "", "sin": "realizado: realizado <> realizada, realizacao", "sug": "Realização de  atendimento telefônico realizada"}, {"id1": 43, "cod1": 11240049, "st1": "Ativa", "e1": "Ações de acompanhamento e fiscalização do Programa de Aquisição de Alimentos (PAA) preparadas", "id2": 88, "cod2": 11240058, "st2": "Ativa", "e2": "Ações de acompanhamento e fiscalização do Programa de Aquisição de Alimentos (PAA) monitoradas", "sim": 0.9753, "macro": "Gestão de políticas públicas de segurança alimentar e nutricional", "cat": "Acesso à alimentação adequada e saudável", "serv": "", "sin": "", "sug": ""}, {"id1": 45, "cod1": 11240050, "st1": "Ativa", "e1": "Relatório final do Plano Anual de Acompanhamento do  Programa de Aquisição de Alimentos (PAA) elaborado", "id2": 47, "cod2": 11240048, "st2": "Ativa", "e2": "Plano Anual de Acompanhamento do Programa de Aquisição de Alimentos (PAA) elaborado", "sim": 0.9749, "macro": "Gestão de políticas públicas de segurança alimentar e nutricional", "cat": "Acesso à alimentação adequada e saudável", "serv": "", "sin": "", "sug": ""}, {"id1": 47, "cod1": 11240048, "st1": "Ativa", "e1": "Plano Anual de Acompanhamento do Programa de Aquisição de Alimentos (PAA) elaborado", "id2": 87, "cod2": 11240056, "st2": "Ativa", "e2": "Estudos avaliativos do Programa de Aquisição de Alimentos (PAA) elaborado", "sim": 0.9703, "macro": "Gestão de políticas públicas de segurança alimentar e nutricional", "cat": "Acesso à alimentação adequada e saudável", "serv": "", "sin": "", "sug": ""}, {"id1": 62, "cod1": 11240110, "st1": "Ativa", "e1": "Estudo de pagamentos aos beneficiários fornecedores à comissão de fiscalização do contrato com o Banco do Brasil nº 73/2023 elaborado", "id2": 76, "cod2": 11240109, "st2": "Ativa", "e2": "Estudo de cadastros e cartões emitidos à comissão de fiscalização do contrato com o Banco do Brasil nº 73/2023 elaborado", "sim": 0.9772, "macro": "Gestão de políticas públicas de segurança alimentar e nutricional", "cat": "Acesso à alimentação adequada e saudável", "serv": "Gestão do Sistema de Informação do Programa de Aquisição de Alimentos (SISPAA)", "sin": "", "sug": ""}, {"id1": 65, "cod1": 11240004, "st1": "Ativa", "e1": " Termo de Execução Descentralizada (TED) para mapeamento e qualificação dos Equipamentos de Segurança Alimentar e Nutricional (EqSAN) formalizado ", "id2": 93, "cod2": 11240006, "st2": "Ativa", "e2": "Termo de Execução Descentralizada  para mapeamento e qualificação dos Equipamentos de Segurança Alimentar e Nutricional (EqSAN) finalizado ", "sim": 0.9722, "macro": "Promoção do acesso e do consumo de alimentação adequada e saudável", "cat": "Acesso à alimentação adequada e saudável", "serv": "", "sin": "", "sug": ""}, {"id1": 77, "cod1": 11240107, "st1": "Ativa", "e1": "Gestores em dúvidas sobre pagamentos aos beneficiários fornecedores do Programa de Aquisição de Alimentos atendidos", "id2": 91, "cod2": 11240106, "st2": "Ativa", "e2": "Gestores em dúvidas sobre emissão de cartões  do Programa de Aquisição de Alimentos atendidos", "sim": 0.9751, "macro": "Gestão de políticas públicas de segurança alimentar e nutricional", "cat": "Acesso à alimentação adequada e saudável", "serv": "Gestão do Sistema de Informação do Programa de Aquisição de Alimentos (SISPAA)", "sin": "", "sug": ""}, {"id1": 85, "cod1": 11240114, "st1": "Ativa", "e1": "Estudos sobre a execução do Programa Fomento Rural preparados", "id2": 99, "cod2": 11240115, "st2": "Ativa", "e2": "Estudos e avaliações do Programa Fomento rural acompanhados", "sim": 0.9889, "macro": "Gestão de políticas públicas de segurança alimentar e nutricional", "cat": "Acesso à alimentação adequada e saudável", "serv": "Plano Anual de Acompanhamento e Fiscalização do Programa de Aquisição de Alimentos (PAA)", "sin": "", "sug": ""}, {"id1": 94, "cod1": 11240099, "st1": "Ativa", "e1": " Atividades para capacitação, orientação e avaliação com entes federativos - unidades executoras do PAA, remotas e presenciais promovidas", "id2": 97, "cod2": 11240100, "st2": "Ativa", "e2": "Orientação a parceiros entes federativos - unidades executoras do PAA, remotas e presenciais realizadas", "sim": 0.9812, "macro": "Gestão de políticas públicas de segurança alimentar e nutricional", "cat": "Acesso à alimentação adequada e saudável", "serv": "Atendimento ao público referente ao Programa de Aquisição de Alimentos (PAA)", "sin": "", "sug": ""}, {"id1": 125, "cod1": 9960078, "st1": "Ativa", "e1": "Subsídio para atendimento a pedido de acesso à informação solicitado", "id2": 132, "cod2": 9960071, "st2": "Ativa", "e2": "Pedidos de acesso à informação respondidos", "sim": 0.9772, "macro": "Participação e controle social", "cat": "Acesso à informação", "serv": "Gerenciar demandas do serviço de informação ao cidadão", "sin": "", "sug": ""}, {"id1": 126, "cod1": 9960076, "st1": "Ativa", "e1": "Elaboração e implementação de fluxos de trabalho Serviço de Acesso à Informação realizados", "id2": 133, "cod2": 9960077, "st2": "Ativa", "e2": "Divulgação e orientação sobre fluxos de trabalho Serviço de Acesso à Informação realizados", "sim": 0.9838, "macro": "Participação e controle social", "cat": "Acesso à informação", "serv": "Gerenciar demandas do serviço de informação ao cidadão", "sin": "", "sug": ""}, {"id1": 127, "cod1": 9960075, "st1": "Ativa", "e1": "Dados gerenciais do serviço de acesso à informação extraídos", "id2": 132, "cod2": 9960071, "st2": "Ativa", "e2": "Pedidos de acesso à informação respondidos", "sim": 0.9711, "macro": "Participação e controle social", "cat": "Acesso à informação", "serv": "Gerenciar demandas do serviço de informação ao cidadão", "sin": "", "sug": ""}, {"id1": 129, "cod1": 9960073, "st1": "Ativa", "e1": "Qualidade da resposta a pedido de acesso à informação monitorada", "id2": 131, "cod2": 9960072, "st2": "Ativa", "e2": "Prazos de respostas a pedido de acesso à informação monitorado", "sim": 0.9827, "macro": "Participação e controle social", "cat": "Acesso à informação", "serv": "Gerenciar demandas do serviço de informação ao cidadão", "sin": "", "sug": ""}, {"id1": 129, "cod1": 9960073, "st1": "Ativa", "e1": "Qualidade da resposta a pedido de acesso à informação monitorada", "id2": 132, "cod2": 9960071, "st2": "Ativa", "e2": "Pedidos de acesso à informação respondidos", "sim": 0.9734, "macro": "Participação e controle social", "cat": "Acesso à informação", "serv": "Gerenciar demandas do serviço de informação ao cidadão", "sin": "", "sug": ""}, {"id1": 132, "cod1": 9960071, "st1": "Ativa", "e1": "Pedidos de acesso à informação respondidos", "id2": 140, "cod2": 9960069, "st2": "Ativa", "e2": "Subsídio para atendimento a pedido de acesso à informação prestado", "sim": 0.9795, "macro": "Participação e controle social", "cat": "Acesso à informação", "serv": "Gerenciar demandas do serviço de informação ao cidadão", "sin": "", "sug": ""}, {"id1": 298, "cod1": 13810018, "st1": "Ativa", "e1": "Informações da DPF para o BGU disponibilizadas", "id2": 308, "cod2": 13810019, "st2": "Ativa", "e2": "Informações da DPFe disponibilizadas", "sim": 0.9779, "macro": "Gestão de passivos", "cat": "Acompanhamento da evolução da dívida pública", "serv": "Elaboração e disponibilização de informações e estatísticas da DPF", "sin": "", "sug": ""}, {"id1": 298, "cod1": 13810018, "st1": "Ativa", "e1": "Informações da DPF para o BGU disponibilizadas", "id2": 311, "cod2": 13810020, "st2": "Ativa", "e2": "Informações da Dívida Garantida disponibilizadas", "sim": 0.9706, "macro": "Gestão de passivos", "cat": "Acompanhamento da evolução da dívida pública", "serv": "Elaboração e disponibilização de informações e estatísticas da DPF", "sin": "", "sug": ""}, {"id1": 308, "cod1": 13810019, "st1": "Ativa", "e1": "Informações da DPFe disponibilizadas", "id2": 311, "cod2": 13810020, "st2": "Ativa", "e2": "Informações da Dívida Garantida disponibilizadas", "sim": 0.9852, "macro": "Gestão de passivos", "cat": "Acompanhamento da evolução da dívida pública", "serv": "Elaboração e disponibilização de informações e estatísticas da DPF", "sin": "", "sug": ""}, {"id1": 308, "cod1": 13810019, "st1": "Ativa", "e1": "Informações da DPFe disponibilizadas", "id2": 324, "cod2": 13810024, "st2": "Ativa", "e2": "Informações para o COPOM disponibilizadas", "sim": 0.9715, "macro": "Gestão de passivos", "cat": "Acompanhamento da evolução da dívida pública", "serv": "Elaboração e disponibilização de informações e estatísticas da DPF", "sin": "", "sug": ""}, {"id1": 311, "cod1": 13810020, "st1": "Ativa", "e1": "Informações da Dívida Garantida disponibilizadas", "id2": 324, "cod2": 13810024, "st2": "Ativa", "e2": "Informações para o COPOM disponibilizadas", "sim": 0.9709, "macro": "Gestão de passivos", "cat": "Acompanhamento da evolução da dívida pública", "serv": "Elaboração e disponibilização de informações e estatísticas da DPF", "sin": "", "sug": ""}, {"id1": 339, "cod1": 13810030, "st1": "Ativa", "e1": "Metodologia de Cálculo dos Indicadores da Dívida Pública atualizada", "id2": 340, "cod2": 13810033, "st2": "Ativa", "e2": "Metodologia de Cálculo dos Títulos da Dívida Interna atualizada", "sim": 0.9763, "macro": "Gestão de passivos", "cat": "Acompanhamento da evolução da dívida pública", "serv": "Elaboração e disponibilização de informações e estatísticas da DPF", "sin": "", "sug": ""}, {"id1": 363, "cod1": 10440006, "st1": "Ativa", "e1": "Acompanhamento da implantação da Rede Privativa da Administração Pública Federal realizado", "id2": 365, "cod2": 10440005, "st2": "Ativa", "e2": "Planejamento da implantação da Rede Privativa da Administração Pública Federal realizado", "sim": 0.984, "macro": "Normatização e efetivação das políticas de comunicações", "cat": "Acompanhamento e Avaliação de Programas e Ações de Comunicações", "serv": "", "sin": "", "sug": ""}, {"id1": 393, "cod1": 12830003, "st1": "Ativa", "e1": "DGT - Demonstrativo de Gastos Tributários - PLOA publicado", "id2": 395, "cod2": 12830002, "st2": "Ativa", "e2": "DGT - Demonstrativo de Gastos Tributários - Bases Efetivas publicado", "sim": 0.9899, "macro": "Gestão das políticas tributária e aduaneira", "cat": "Acompanhar e subsidiar a avaliação das políticas públicas implementadas com benefício fiscal", "serv": "Mensurar gastos tributários ", "sin": "", "sug": ""}, {"id1": 451, "cod1": 12680002, "st1": "Ativa", "e1": "Previsão de dividendos para o PLOA atualizado", "id2": 455, "cod2": 12680001, "st2": "Ativa", "e2": "Previsão de dividendos para o PLDO atualizado", "sim": 0.9718, "macro": "Gestão de ativos", "cat": "Administração de haveres mobiliários", "serv": "Controle Estratégico dos Haveres Mobiliários", "sin": "", "sug": ""}, {"id1": 453, "cod1": 12680006, "st1": "Ativa", "e1": "Proventos a receber conciliados e atualizados", "id2": 459, "cod2": 12680009, "st2": "Ativa", "e2": "Demais Contas Conciliadas e atualizadas", "sim": 0.9783, "macro": "Gestão de ativos", "cat": "Administração de haveres mobiliários", "serv": "Controle Estratégico dos Haveres Mobiliários", "sin": "", "sug": ""}, {"id1": 488, "cod1": 11920002, "st1": "Ativa", "e1": "Lançamento do Laboratório de Inovação em Políticas Públicas para Sistemas Alimentares Saudáveis e Sustentáveis (AlimentaLAB) realizado", "id2": 489, "cod2": 11920003, "st2": "Ativa", "e2": "Laboratório de Inovação em Políticas Públicas para Sistemas Alimentares Saudáveis e Sustentáveis (AlimentaLAB) implementado", "sim": 0.9742, "macro": "Promoção do acesso e do consumo de alimentação adequada e saudável", "cat": "Alimentação Saudável", "serv": "", "sin": "", "sug": ""}, {"id1": 499, "cod1": 10370021, "st1": "Ativa", "e1": "Ato de Concentração (AC) ordinário concluído", "id2": 505, "cod2": 10370007, "st2": "Ativa", "e2": "Ato de Concentração (AC) sumário concluído", "sim": 0.9718, "macro": "Competitividade e política regulatória", "cat": "Análise de condutas e atos de concentração", "serv": "", "sin": "", "sug": ""}, {"id1": 523, "cod1": 8620014, "st1": "Ativa", "e1": "Ambientes operacionais dos bancos de dados de comércio exterior implementados", "id2": 544, "cod2": 8620013, "st2": "Ativa", "e2": "Suporte em ambientes operacionais dos bancos de dados de comércio exterior realizado", "sim": 0.9772, "macro": "Comércio exterior", "cat": "Análise de dados operacionais e padronização de procedimentos", "serv": "", "sin": "", "sug": ""}];

/* ---------- similaridade (client-side, sem embeddings) ---------- */
const norm = s => (s||"").toString().toLowerCase().normalize("NFD")
  .replace(/[̀-ͯ]/g,"").replace(/[^a-z0-9\s]/g," ").replace(/\s+/g," ").trim();
const toks = s => new Set(norm(s).split(" ").filter(Boolean));
function jaccard(a,b){ const A=toks(a),B=toks(b); if(!A.size||!B.size) return 0;
  let inter=0; A.forEach(t=>{ if(B.has(t)) inter++; }); return inter/(A.size+B.size-inter); }
function trigrams(s){ const n=norm(s).replace(/\s/g,""); const g=new Set();
  for(let i=0;i<n.length-2;i++) g.add(n.slice(i,i+3)); return g; }
function dice(a,b){ const A=trigrams(a),B=trigrams(b); if(!A.size||!B.size) return 0;
  let inter=0; A.forEach(t=>{ if(B.has(t)) inter++; }); return (2*inter)/(A.size+B.size); }
function simEntrega(a,b){ return 0.5*dice(a,b)+0.5*jaccard(a,b); }
function eq(a,b){ return norm(a)===norm(b); }

/* diff destaque (palavras da a que não estão na b) */
function diffHi(a,b){ const setB=toks(b);
  return (a||"").toString().split(/(\s+)/).map((t,i)=>{ const n=norm(t);
    return (n&&!setB.has(n))
      ? <mark key={i} className="cr-df">{t}</mark> : <span key={i}>{t}</span>; }); }

/* ---------- CSV ---------- */
function parseCSV(text){
  text=text.replace(/^﻿/,"");
  const lines=text.split(/\r?\n/).filter(l=>l.length);
  if(!lines.length) return [];
  const delim = (lines[0].split(";").length > lines[0].split(",").length) ? ";" : ",";
  const split=(line)=>{ const out=[]; let cur="",q=false;
    for(let i=0;i<line.length;i++){ const c=line[i];
      if(c==='"'){ if(q&&line[i+1]==='"'){cur+='"';i++;} else q=!q; }
      else if(c===delim&&!q){ out.push(cur); cur=""; }
      else cur+=c; }
    out.push(cur); return out.map(s=>s.trim()); };
  const hdr=split(lines[0]);
  return lines.slice(1).map(l=>{ const c=split(l); const o={}; hdr.forEach((h,i)=>o[h]=c[i]||""); return o; });
}
function toCSV(rows, cols){
  const esc=v=>{ v=(v==null?"":String(v)); return /[";\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v; };
  const head=cols.join(";");
  const body=rows.map(r=>cols.map(c=>esc(r[c])).join(";")).join("\n");
  return "﻿"+head+"\n"+body;
}
function download(name, text){
  const blob=new Blob([text],{type:"text/csv;charset=utf-8"});
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob);
  a.download=name; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1500);
}

const COLS_EXPORT = ["Código","Entrega","MacroProcesso","Categoria de Serviço","Serviço",
  "Origem do Serviço","Tratamento","Decisão Revisor","Código Mantido","Código Desativado",
  "Texto Mesclado","Serviço Sugerido IA","Justificativa IA","Revisor","Aba","Registrado em"];

/* ============================================================ */
function CentralRevisao({ onClose, embutido, flags }){
  const [aba,setAba]=useState("similares");
  const [revisor,setRevisor]=useState("");
  const [banco,setBanco]=useState(SAMPLE.corpus);           // corpus p/ similaridade
  const [semServico,setSemServico]=useState(SAMPLE.semServico);
  const [pares,setPares]=useState(PAIRS_SEED);
  const [decisoes,setDecisoes]=useState([]);                // linhas p/ exportar
  const [toast,setToast]=useState(null);
  const fileRef=useRef(null);

  function flash(t){ setToast(t); setTimeout(()=>setToast(null),1900); }
  function registrar(linhas){ setDecisoes(d=>[...d, ...linhas]); }

  /* importar CSV: detecta se é planilha de pares (id1/id2/sim) ou banco de entregas */
  function importCSV(ev){
    const f=ev.target.files[0]; if(!f) return;
    const rd=new FileReader();
    rd.onload=()=>{
      const rows=parseCSV(rd.result);
      if(!rows.length){ flash("CSV vazio."); return; }
      const k=Object.keys(rows[0]).map(x=>x.toLowerCase());
      if(k.includes("id1")&&k.includes("id2")){
        const p=rows.map(r=>({
          id1:r.id1,cod1:r.cod1||r.codigo1||"",st1:r.st1||"Ativa",e1:r.e1||r.entrega1||"",
          id2:r.id2,cod2:r.cod2||r.codigo2||"",st2:r.st2||"Ativa",e2:r.e2||r.entrega2||"",
          sim:parseFloat(r.sim||r.similaridade||"0"),macro:r.macro||"",cat:r.cat||"",
          serv:r.serv||r.servico||"",sin:r.sin||"",sug:r.sug||"",
        })).filter(x=>x.e1&&x.e2);
        setPares(p); flash(`${p.length} pares carregados para revisão.`);
      } else {
        const norm2=r=>({
          cod:r["Código"]||r.codigo||r.cod||"",
          ent:r["Entrega"]||r.entrega||r.ent||"",
          mac:r["MacroProcesso"]||r.macroprocesso||r.mac||"",
          cat:r["Categoria de Serviço"]||r.categoria||r.cat||"",
          serv:r["Serviço"]||r.servico||r.serv||"",
          nat:r["Natureza"]||r.natureza||r.nat||"",
        });
        const all=rows.map(norm2).filter(x=>x.ent);
        setBanco(all);
        setSemServico(all.filter(x=>!x.serv));
        flash(`${all.length} entregas carregadas (${all.filter(x=>!x.serv).length} sem serviço).`);
      }
    };
    rd.readAsText(f); ev.target.value="";
  }

  function exportar(){
    if(!decisoes.length){ flash("Nenhuma decisão registrada ainda."); return; }
    const stamp=new Date().toISOString().slice(0,19).replace("T"," ");
    const rows=decisoes.map(d=>({...d, "Revisor":d["Revisor"]||revisor||"—", "Registrado em":d["Registrado em"]||stamp}));
    download(`decisoes_central_revisao_${(revisor||"revisor").replace(/\s+/g,"_")}.csv`, toCSV(rows,COLS_EXPORT));
    flash(`${rows.length} decisões exportadas — mescle na planilha pelo Código.`);
  }

  const [filtro,setFiltro]=useState({mac:null,cat:null});
  const paresF=useMemo(()=>pares.filter(p=>(!filtro.mac||p.macro===filtro.mac)&&(!filtro.cat||p.cat===filtro.cat)),[pares,filtro]);
  const semServicoF=useMemo(()=>semServico.filter(x=>(!filtro.mac||x.mac===filtro.mac)&&(!filtro.cat||x.cat===filtro.cat)),[semServico,filtro]);
  const sinalizacoes=flags||[];
  const abas=[
    ["similares","Similares",GitMerge, paresF.length],
    ["servicos","Serviços",Sparkles, semServicoF.length],
    ["sinalizacoes","Sinalizações",Flag, sinalizacoes.length],
  ];

  return (
    <div className={embutido?"cr-root cr-embutido":"cr-root"}>
      <style>{CSS}</style>
      {!embutido && <div className="cr-stripe"><span style={{background:C.green}}/><span style={{background:C.yellow}}/><span style={{background:C.primary}}/></div>}

      <header className="cr-head">
        <div className="cr-title">
          <span className="cr-badge"><ClipboardCheck size={18}/></span>
          <div>
            <div className="cr-h1">Central de Revisão de Entregas</div>
            <div className="cr-h2">Curadoria do banco · nenhuma alteração é automática, tudo aguarda sua confirmação</div>
          </div>
        </div>
        <div className="cr-headr">
          <div className="cr-rev">
            <label>Revisor</label>
            <input value={revisor} onChange={e=>setRevisor(e.target.value)} placeholder="Seu nome"/>
          </div>
          <button className="cr-back" onClick={onClose} title="Voltar à página principal"><ChevronLeft size={15}/> Voltar ao catálogo</button>
        </div>
      </header>

      <div className="cr-databar">
        <span className="cr-dstat"><Layers size={13}/> {banco.length} entregas no banco · {pares.length} pares · {semServico.length} sem serviço</span>
        <div className="cr-dactions">
          <input ref={fileRef} type="file" accept=".csv" style={{display:"none"}} onChange={importCSV}/>
        </div>
      </div>

      <nav className="cr-tabs">
        {abas.map(([id,lbl,Ic,n])=>(
          <button key={id} className={`cr-tab ${aba===id?"on":""}`} onClick={()=>setAba(id)}>
            <Ic size={15}/> {lbl}{n!=null && <span className="cr-tabn">{n}</span>}
          </button>
        ))}
      </nav>

      <div className="cr-body">
        <ArvoreBanco banco={banco} pares={pares} semServico={semServico} filtro={filtro} onFiltrar={setFiltro}/>
        <main className="cr-main">
          {aba==="similares" && <AbaSimilares pares={paresF} revisor={revisor} onDecidir={registrar} flash={flash}/>}
          {aba==="servicos"  && <AbaServicos itens={semServicoF} revisor={revisor} onAprovar={registrar} flash={flash}/>}
          {aba==="sinalizacoes" && <QualidadePanel embutido flags={sinalizacoes} onClose={onClose}/>}
        </main>
      </div>

      {toast && <div className="cr-toast">{toast}</div>}
    </div>
  );
}

/* ============================================================
   PAINEL LATERAL — árvore do banco (natureza › macro › categoria)
   ============================================================ */
function ArvoreBanco({ banco, pares, semServico, filtro, onFiltrar }){
  const [aberto,setAberto]=useState(true);
  const [exp,setExp]=useState({});
  const [q,setQ]=useState("");
  const tree=useMemo(()=>{
    // "a revisar" = sem serviço OU envolvida em algum par de similaridade
    const revCods=new Set();
    (pares||[]).forEach(p=>{ if(p.cod1) revCods.add(String(p.cod1)); if(p.cod2) revCods.add(String(p.cod2)); });
    (semServico||[]).forEach(s=>{ if(s.cod) revCods.add(String(s.cod)); });
    const fonte=(banco||[]).filter(e=> !e.serv || revCods.has(String(e.cod)));
    const root=new Map();
    fonte.forEach(e=>{
      const nat=e.nat||"—", mac=e.mac||"—", cat=e.cat||"—";
      if(!root.has(nat)) root.set(nat,new Map());
      const macs=root.get(nat);
      if(!macs.has(mac)) macs.set(mac,new Map());
      const cats=macs.get(mac);
      cats.set(cat,(cats.get(cat)||0)+1);
    });
    return root;
  },[banco,pares,semServico]);
  const nq=norm(q);
  const toggle=k=>setExp(o=>({...o,[k]:!o[k]}));

  if(!aberto) return (
    <div className="cr-tree collapsed">
      <button className="cr-tree-open" onClick={()=>setAberto(true)} title="Mostrar organização do banco">
        <Layers size={15}/><span>Organização do banco</span>
      </button>
    </div>
  );
  return (
    <aside className="cr-tree">
      <div className="cr-tree-h">
        <div className="cr-tree-t"><Layers size={14}/> Organização do banco</div>
        <button className="cr-tree-x" onClick={()=>setAberto(false)} title="Recolher painel"><ChevronRight size={16}/></button>
      </div>
      <div className="cr-tree-search"><Search size={13} color={C.faint}/>
        <input value={q} onChange={e=>setQ(e.target.value)} placeholder="Filtrar macro/categoria…"/></div>
      {filtro&&(filtro.mac||filtro.cat) && <div className="cr-tfilt-active"><span title={filtro.cat||filtro.mac}>Filtrando: {filtro.cat||filtro.mac}</span><button onClick={()=>onFiltrar&&onFiltrar({mac:null,cat:null})}><X size={12}/></button></div>}
      <div className="cr-tree-body">
        {[...tree.keys()].sort().map(nat=>{
          const macs=tree.get(nat); const nk="n:"+nat;
          const natTot=[...macs.values()].reduce((s,cats)=>s+[...cats.values()].reduce((a,b)=>a+b,0),0);
          const style=NAT[nat]||{cor:C.faint,soft:C.bg};
          return (
            <div key={nat} className="cr-tnat">
              <button className="cr-trow lvl0" onClick={()=>toggle(nk)}>
                {exp[nk]?<ChevronDown size={13}/>:<ChevronRight size={13}/>}
                <span className="cr-tdot" style={{background:style.cor}}/>
                <span className="cr-tlabel"><b>{nat}</b></span><span className="cr-tcount">{natTot}</span>
              </button>
              {exp[nk] && [...macs.keys()].sort().map(mac=>{
                const cats=macs.get(mac); const mk=nk+"|m:"+mac;
                const macTot=[...cats.values()].reduce((a,b)=>a+b,0);
                if(nq && !norm(mac).includes(nq) && ![...cats.keys()].some(c=>norm(c).includes(nq))) return null;
                return (
                  <div key={mac}>
                    <div className={`cr-trow lvl1 ${filtro&&filtro.mac===mac&&!filtro.cat?"filt":""}`}>
                      <button className="cr-texp" onClick={()=>toggle(mk)} title={exp[mk]?"recolher":"expandir"}>{exp[mk]?<ChevronDown size={12}/>:<ChevronRight size={12}/>}</button>
                      <button className="cr-tfilt" onClick={()=>onFiltrar&&onFiltrar({mac,cat:null})} title="Filtrar similares e serviços por este macroprocesso"><span className="cr-tlabel" title={mac}>{mac}</span><span className="cr-tcount">{macTot}</span></button>
                    </div>
                    {exp[mk] && [...cats.keys()].sort()
                      .filter(c=>!nq||norm(c).includes(nq)||norm(mac).includes(nq))
                      .map(cat=>(
                        <button key={cat} className={`cr-trow lvl2 cr-tfilt ${filtro&&filtro.cat===cat?"filt":""}`} onClick={()=>onFiltrar&&onFiltrar({mac,cat})}>
                          <span className="cr-tlabel" title={cat}>{cat}</span><span className="cr-tcount">{cats.get(cat)}</span>
                        </button>
                      ))}
                  </div>
                );
              })}
            </div>
          );
        })}
        {!tree.size && <div className="cr-tree-empty">Nada a revisar no momento — sem entregas pendentes de serviço nem pares similares. Importe o banco/pares (CSV) para carregar.</div>}
      </div>
    </aside>
  );
}

/* ============================================================
   ABA 1 — SIMILARES (faixa 0,97–0,99)
   ============================================================ */
function AbaSimilares({ pares, revisor, onDecidir, flash }){
  const [i,setI]=useState(0);
  const [feitas,setFeitas]=useState({});
  const [merge,setMerge]=useState("");
  const [mergeOn,setMergeOn]=useState(false);
  const key=d=>d.id1+"_"+d.id2;

  useEffect(()=>{ setMergeOn(false); setMerge(pares[i]?.sug||""); },[i,pares]);

  const total=pares.length;
  const nFeitas=Object.keys(feitas).length;
  const pct=total?Math.round(100*nFeitas/total):0;

  if(!total) return <Vazio icon={GitMerge} titulo="Sem pares para revisar"
    texto="Importe a planilha de pares (colunas id1/id2/sim) pelo botão “Importar CSV” no topo. Ela sai do pipeline de detecção de quase-duplicatas (faixa 0,97–0,99)."/>;

  if(nFeitas>=total) return <div className="cr-done">
    <div className="cr-done-ic"><Check size={40}/></div>
    <div className="cr-done-t">Todos os {total} pares revisados</div>
    <div className="cr-done-s">Exporte as decisões no topo para mesclar na planilha.</div>
    <button className="cr-ghost" onClick={()=>{setFeitas({});setI(0);}}><RotateCcw size={13}/> Revisar de novo</button>
  </div>;

  const d=pares[i];
  const dec=feitas[key(d)];

  function avancar(){ let j=i; for(let c=0;c<total;c++){ j=(j+1)%total; if(!feitas[key(pares[j])]){ setI(j); return; } } }
  function decidir(acao, extra={}){
    const stamp=new Date().toISOString().slice(0,19).replace("T"," ");
    const linhas=[];
    if(acao==="manter_1"||acao==="manter_2"){
      const mant = acao==="manter_1"?d.cod1:d.cod2;
      const desat= acao==="manter_1"?d.cod2:d.cod1;
      const eMant= acao==="manter_1"?d.e1:d.e2;
      const eDes = acao==="manter_1"?d.e2:d.e1;
      linhas.push(row(mant,eMant,d,{ "Tratamento":"Manter","Decisão Revisor":"Manter (duplicata resolvida)","Código Mantido":mant,"Código Desativado":desat }));
      linhas.push(row(desat,eDes,d,{ "Tratamento":"Retirar","Decisão Revisor":"Desativar (duplicata)","Código Mantido":mant,"Código Desativado":desat }));
    } else if(acao==="nao_dup"){
      linhas.push(row(d.cod1,d.e1,d,{ "Decisão Revisor":"Não são duplicatas" }));
      linhas.push(row(d.cod2,d.e2,d,{ "Decisão Revisor":"Não são duplicatas" }));
    } else if(acao==="mesclar"){
      linhas.push(row(d.cod1,d.e1,d,{ "Tratamento":"Retirar","Decisão Revisor":"Mesclada","Código Desativado":d.cod1,"Texto Mesclado":extra.texto }));
      linhas.push(row(d.cod2,d.e2,d,{ "Tratamento":"Retirar","Decisão Revisor":"Mesclada","Código Desativado":d.cod2,"Texto Mesclado":extra.texto }));
      linhas.push(row("(nova)",extra.texto,d,{ "Tratamento":"Em análise","Decisão Revisor":"Nova entrega (mescla)","Texto Mesclado":extra.texto }));
    }
    linhas.forEach(l=>{ l["Revisor"]=revisor||"—"; l["Aba"]="Similares"; l["Registrado em"]=stamp; });
    onDecidir(linhas);
    setFeitas(f=>({...f,[key(d)]:acao}));
    flash(rotulo(acao));
    avancar();
  }

  return (
    <div className="cr-sim">
      <div className="cr-simhead">
        <div className="cr-bar"><i style={{width:pct+"%"}}/></div>
        <div className="cr-barrow"><span><b>{nFeitas}</b> de {total} revisados</span><span>{pct}%</span></div>
      </div>

      <div className="cr-card">
        <div className="cr-meta">
          {d.macro && <span className="cr-chain-i"><em>Macroprocesso</em>{d.macro}</span>}
          {d.cat && <><ChevronRight size={12} className="cr-chain-sep"/><span className="cr-chain-i"><em>Categoria de serviço</em>{d.cat}</span></>}
          <ChevronRight size={12} className="cr-chain-sep"/><span className="cr-chain-i"><em>Serviço</em>{d.serv||"—"}</span>
          <span className="cr-sim-pct">similaridade {(Number(d.sim)*100).toFixed(1)}%</span>
        </div>

        <div className="cr-cols-lbl"><span className="cr-tagent">Entrega</span> compare os dois textos e escolha qual manter</div>
        <div className="cr-cols">
          <button className={`cr-opt ${dec==="manter_1"?"sel":""}`} onClick={()=>decidir("manter_1")}>
            <div className="cr-opt-top"><span className="cr-key">1</span><span className="cr-cod">cód. {d.cod1||"—"}</span><Badge st={d.st1}/></div>
            <div className="cr-txt">{diffHi(d.e1,d.e2)}</div>
          </button>
          <button className={`cr-opt ${dec==="manter_2"?"sel":""}`} onClick={()=>decidir("manter_2")}>
            <div className="cr-opt-top"><span className="cr-key">2</span><span className="cr-cod">cód. {d.cod2||"—"}</span><Badge st={d.st2}/></div>
            <div className="cr-txt">{diffHi(d.e2,d.e1)}</div>
          </button>
        </div>

        {d.sin && <div className="cr-hint">Diferença é só grafia/plural: <b>{d.sin}</b></div>}

        <div className="cr-acts">
          <button className="cr-act" onClick={()=>decidir("manter_1")}><span className="cr-k">1</span>Manter a 1 · retirar a 2</button>
          <button className="cr-act" onClick={()=>decidir("manter_2")}><span className="cr-k">2</span>Manter a 2 · retirar a 1</button>
          <button className="cr-act dup" onClick={()=>decidir("nao_dup")}><span className="cr-k">3</span>Não são duplicatas</button>
          <button className="cr-act merge" onClick={()=>{setMergeOn(true);}}><span className="cr-k">4</span>Mesclar num texto novo</button>
        </div>

        {mergeOn && <div className="cr-mergebox">
          <div className="cr-mergel">Texto da nova entrega {d.sug?"— sugestão do pipeline, ajuste se quiser":""}</div>
          <textarea value={merge} onChange={e=>setMerge(e.target.value)}/>
          <button className="cr-primary" disabled={!merge.trim()} onClick={()=>decidir("mesclar",{texto:merge.trim()})}>Confirmar mescla <ArrowRight size={14}/></button>
        </div>}

        <div className="cr-row2">
          <button onClick={avancar}>Pular este par</button>
        </div>
      </div>
    </div>
  );
}
function rotulo(a){ return a==="manter_1"?"Mantida a entrega 1.":a==="manter_2"?"Mantida a entrega 2.":a==="nao_dup"?"Marcadas como distintas.":"Mescla registrada."; }
function row(cod,ent,d,extra){ return {
  "Código":cod,"Entrega":ent,"MacroProcesso":d.macro||"","Categoria de Serviço":d.cat||"",
  "Serviço":d.serv||"","Origem do Serviço":"","Tratamento":"","Decisão Revisor":"",
  "Código Mantido":"","Código Desativado":"","Texto Mesclado":"","Serviço Sugerido IA":"","Justificativa IA":"",
  ...extra }; }
function Badge({st}){ const c=st==="Ativa"?"ativa":st==="Inativa"?"inativa":"nao";
  return <span className={`cr-b cr-b-${c}`}>{st||"—"}</span>; }

/* ============================================================
   ABA 2 — SERVIÇOS (sugestão IA + aprovação humana)
   ============================================================ */
function AbaServicos({ itens, revisor, onAprovar, flash }){
  const [estado,setEstado]=useState({});   // cod -> {sug,just,val,load,err,aprovado}
  const [q,setQ]=useState("");
  const lista = useMemo(()=>{ const n=norm(q);
    return itens.filter(it=> !n || norm(it.ent).includes(n) || norm(it.mac).includes(n) || norm(it.cat).includes(n)); },[itens,q]);

  async function sugerir(it){
    const cod=it.cod;
    setEstado(s=>({...s,[cod]:{...(s[cod]||{}),load:true,err:null}}));
    try{
      const r=await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST", headers:{"Content-Type":"application/json"},
        body:JSON.stringify({ model:"claude-sonnet-4-6", max_tokens:1000,
          messages:[{role:"user",content:promptServico(it)}] })
      });
      const data=await r.json();
      const txt=(data.content||[]).map(b=>b.type==="text"?b.text:"").join("").replace(/```json|```/g,"").trim();
      const obj=JSON.parse(txt);
      setEstado(s=>({...s,[cod]:{ sug:obj.servico_sugerido||"", just:obj.justificativa||"",
        val:(s[cod]&&s[cod].val)||obj.servico_sugerido||"", load:false }}));
    }catch(e){
      setEstado(s=>({...s,[cod]:{...(s[cod]||{}),load:false,
        err:"Sem IA aqui — preencha o serviço à mão, ou rode o script offline sugerir_servico_comIA.py."}}));
    }
  }
  function setVal(cod,v){ setEstado(s=>({...s,[cod]:{...(s[cod]||{}),val:v}})); }
  function aprovar(it){
    const st=estado[it.cod]||{}; const val=(st.val||"").trim();
    if(!val){ flash("Escreva ou aceite um serviço antes de aprovar."); return; }
    const stamp=new Date().toISOString().slice(0,19).replace("T"," ");
    onAprovar([{
      "Código":it.cod,"Entrega":it.ent,"MacroProcesso":it.mac,"Categoria de Serviço":it.cat,
      "Serviço":val,"Origem do Serviço": st.sug&&val===st.sug ? "ia":"manual","Tratamento":"",
      "Decisão Revisor":"Serviço aprovado","Código Mantido":"","Código Desativado":"","Texto Mesclado":"",
      "Serviço Sugerido IA":st.sug||"","Justificativa IA":st.just||"","Revisor":revisor||"—","Aba":"Serviços","Registrado em":stamp,
    }]);
    setEstado(s=>({...s,[it.cod]:{...st,aprovado:true}}));
    flash(`Serviço aprovado para ${it.cod}.`);
  }

  if(!itens.length) return <Vazio icon={Sparkles} titulo="Nenhuma entrega sem serviço"
    texto="Importe o banco (CSV) para listar as entregas que ainda não têm Serviço preenchido. Hoje ~55% do banco está nessa situação."/>;

  return (
    <div className="cr-serv-wrap">
      <div className="cr-serv-intro">
        <ShieldCheck size={15}/> A IA <b>sugere</b> um serviço padronizado a partir do macroprocesso, da categoria e da entrega. Nada entra no banco sem sua aprovação.
      </div>
      <div className="cr-serv-search">
        <Search size={16} color={C.faint}/>
        <input value={q} onChange={e=>setQ(e.target.value)} placeholder="Filtrar por entrega, macroprocesso ou categoria…"/>
        <span className="cr-serv-count">{lista.length} de {itens.length}</span>
      </div>
      <div className="cr-serv-list">
        {lista.map(it=>{ const st=estado[it.cod]||{}; return (
          <div className={`cr-serv-row ${st.aprovado?"ok":""}`} key={it.cod}>
            <div className="cr-serv-ctx">
              <div className="cr-chain">
                {it.mac && <span className="cr-chain-i"><em>Macroprocesso</em>{it.mac}</span>}
                {it.cat && <><ChevronRight size={12} className="cr-chain-sep"/><span className="cr-chain-i"><em>Categoria de serviço</em>{it.cat}</span></>}
              </div>
              <div className="cr-serv-entrow"><span className="cr-tagent">Entrega</span><span className="cr-serv-ent">{it.ent}</span></div>
              <div className="cr-serv-cod">cód. {it.cod}</div>
            </div>
            <div className="cr-serv-act">
              {st.aprovado
                ? <div className="cr-serv-done"><Check size={15}/> {st.val} <span>aprovado</span></div>
                : <>
                  <div className="cr-serv-inrow">
                    <input className="cr-serv-in" value={st.val||""} placeholder="Serviço sugerido / editável"
                      onChange={e=>setVal(it.cod,e.target.value)}/>
                    <button className="cr-serv-ia" onClick={()=>sugerir(it)} disabled={st.load}>
                      <Wand2 size={13}/> {st.load?"Sugerindo…":"Sugerir com IA"}</button>
                    <button className="cr-serv-ok" onClick={()=>aprovar(it)}><Check size={14}/> Aprovar</button>
                  </div>
                  {st.just && <div className="cr-serv-just"><b>Justificativa:</b> {st.just}</div>}
                  {st.err && <div className="cr-serv-err">{st.err}</div>}
                </>}
            </div>
          </div>
        ); })}
      </div>
    </div>
  );
}
function promptServico(it){
  return `Você é especialista em gestão pública e arquitetura de serviços no setor público brasileiro.
Sugira UM nome de serviço padronizado e reutilizável para a entrega abaixo. O serviço deve representar a natureza do trabalho que gera a entrega — nem tão específico quanto uma entrega, nem genérico demais. Comece por substantivo de ação/área (ex.: "Elaboração de relatórios de acompanhamento"). 3 a 8 palavras. Não invente siglas, sistemas ou normas fora do contexto.

Macroprocesso: ${it.mac}
Categoria de serviço: ${it.cat}
Entrega: ${it.ent}

Responda APENAS com JSON válido, sem texto antes ou depois:
{"servico_sugerido":"...", "justificativa":"uma frase curta"}`;
}

/* ============================================================
   ABA 3 — NOVA ENTREGA (checagem de similaridade nos 4 níveis)
   ============================================================ */
function AbaNova({ banco, vocab, revisor, onInserir, flash }){
  const [nat,setNat]=useState("Finalístico");
  const [mac,setMac]=useState("");
  const [cat,setCat]=useState("");
  const [serv,setServ]=useState("");
  const [ent,setEnt]=useState("");
  const [inserida,setInserida]=useState(false);

  const candidatos=useMemo(()=>{
    if(norm(ent).length<6) return [];
    const scored=banco.map(r=>{
      const s=simEntrega(ent,r.ent);
      const nMac=eq(mac,r.mac), nCat=eq(cat,r.cat), nServ=serv&&r.serv?eq(serv,r.serv):false;
      // score composto: texto domina, mesmo macro/cat/serviço puxa pra cima
      const score=s + (nMac?0.05:0) + (nCat?0.05:0) + (nServ?0.03:0);
      return {...r, s, score, nMac, nCat, nServ};
    }).filter(x=>x.s>0.35).sort((a,b)=>b.score-a.score).slice(0,6);
    return scored;
  },[ent,mac,cat,serv,banco]);

  const topo=candidatos[0];
  // limiares na escala desta métrica léxica (quase-duplicata real ~0,6-0,9; idêntica ~0,9+)
  const veredito = !topo ? "ok"
    : (topo.s>=0.90 && topo.nCat) ? "bloqueio"
    : (topo.s>=0.72) ? "atencao"
    : (topo.s>=0.55) ? "revisar" : "ok";

  function inserir(){
    if(veredito==="bloqueio"){ flash("Praticamente idêntica a uma entrega existente — resolva antes de inserir."); return; }
    if(!ent.trim()||!mac||!cat){ flash("Preencha ao menos macroprocesso, categoria e o texto da entrega."); return; }
    const stamp=new Date().toISOString().slice(0,19).replace("T"," ");
    onInserir([{
      "Código":"(nova)","Entrega":ent.trim(),"MacroProcesso":mac,"Categoria de Serviço":cat,
      "Serviço":serv,"Origem do Serviço":serv?"manual":"","Tratamento":"Em análise",
      "Decisão Revisor": topo? `Nova (mais próxima ${(topo.s*100).toFixed(0)}%: ${topo.cod})` : "Nova entrega",
      "Código Mantido":"","Código Desativado":"","Texto Mesclado":"","Serviço Sugerido IA":"","Justificativa IA":"",
      "Revisor":revisor||"—","Aba":"Nova entrega","Registrado em":stamp,
    }]);
    setInserida(true);
    flash("Nova entrega registrada — aguarda curadoria da Frente 1.");
  }

  return (
    <div className="cr-nova">
      <div className="cr-nova-form">
        <div className="cr-nova-t"><FilePlus2 size={16}/> Propor nova entrega ao banco</div>
        <p className="cr-nova-sub">Antes de inserir, o sistema compara com o que já existe nos 4 níveis: macroprocesso → categoria (processo) → serviço → entrega. Assim evitamos duplicar o que já está no banco.</p>

        <label className="cr-l">Natureza</label>
        <div className="cr-natrow">
          {Object.keys(NAT).map(n=>(
            <button key={n} className={`cr-natchip ${nat===n?"on":""}`}
              style={nat===n?{background:NAT[n].cor,borderColor:NAT[n].cor,color:"#fff"}:{}}
              onClick={()=>setNat(n)}>{n}</button>
          ))}
        </div>

        <label className="cr-l">Macroprocesso <span className="cr-lvl">nível 1</span></label>
        <input list="cr-macros" className="cr-in" value={mac} onChange={e=>setMac(e.target.value)} placeholder="Comece a digitar…"/>
        <datalist id="cr-macros">{vocab.macros.map(m=><option key={m} value={m}/>)}</datalist>

        <label className="cr-l">Categoria de serviço (processo) <span className="cr-lvl">nível 2</span></label>
        <input list="cr-cats" className="cr-in" value={cat} onChange={e=>setCat(e.target.value)} placeholder="Comece a digitar…"/>
        <datalist id="cr-cats">{vocab.cats.map(c=><option key={c} value={c}/>)}</datalist>

        <label className="cr-l">Serviço</label>
        <input list="cr-servs" className="cr-in" value={serv} onChange={e=>setServ(e.target.value)} placeholder="Deixe vazio se ainda não há serviço"/>
        <datalist id="cr-servs">{vocab.servicos.slice(0,1200).map(s=><option key={s} value={s}/>)}</datalist>

        <label className="cr-l">Texto da entrega <span className="cr-lvl">nível 4</span></label>
        <textarea className="cr-ta" value={ent} onChange={e=>{setEnt(e.target.value);setInserida(false);}} placeholder="Ex.: Relatório de gestão consolidado elaborado"/>

        <button className={`cr-primary lg ${veredito==="bloqueio"?"blocked":""}`} onClick={inserir} disabled={inserida}>
          {inserida ? <><Check size={15}/> Registrada — aguarda curadoria</> : <><Plus size={15}/> Inserir no banco (aguarda curadoria)</>}
        </button>
      </div>

      <div className="cr-nova-check">
        <div className="cr-check-h"><Layers size={15}/> Já existe algo parecido?</div>
        {norm(ent).length<6
          ? <div className="cr-check-empty">Digite o texto da entrega para ver as mais próximas no banco.</div>
          : <>
            <div className={`cr-verd v-${veredito}`}>
              {veredito==="bloqueio" && <><AlertTriangle size={15}/> Praticamente idêntica a uma entrega existente na mesma categoria. Resolva antes de inserir (talvez seja caso de mesclar).</>}
              {veredito==="atencao" && <><AlertTriangle size={15}/> Muito parecida com uma existente. Confirme que é realmente distinta.</>}
              {veredito==="revisar" && <><Flag size={15}/> Há entregas próximas — vale conferir a lista antes de inserir.</>}
              {veredito==="ok" && <><Check size={15}/> Nada muito próximo. Parece uma entrega nova.</>}
            </div>
            <div className="cr-cands">
              {candidatos.map((c,i)=>(
                <div className="cr-cand" key={c.cod+i}>
                  <div className="cr-cand-top">
                    <span className="cr-cand-sim" style={simColor(c.s)}>{(c.s*100).toFixed(0)}%</span>
                    <span className="cr-cand-cod">{c.cod}</span>
                    <div className="cr-cand-levels">
                      <span className={c.nMac?"hit":""} title="Macroprocesso">M</span>
                      <span className={c.nCat?"hit":""} title="Categoria">C</span>
                      <span className={c.nServ?"hit":""} title="Serviço">S</span>
                    </div>
                  </div>
                  <div className="cr-cand-ent">{diffHi(c.ent,ent)}</div>
                  <div className="cr-cand-path">{c.mac} <ChevronRight size={10}/> {c.cat}{c.serv?<> <ChevronRight size={10}/> {c.serv}</>:null}</div>
                </div>
              ))}
              {!candidatos.length && <div className="cr-check-empty">Nenhuma entrega parecida encontrada no banco carregado.</div>}
            </div>
          </>}
      </div>
    </div>
  );
}
function simColor(s){ return s>=0.90?{background:C.coralSoft,color:C.coral}
  : s>=0.72?{background:C.amberSoft,color:C.amber}:{background:C.primarySoft,color:C.primary}; }

/* ---------- util ui ---------- */
function Vazio({icon:Ic,titulo,texto}){
  return <div className="cr-vazio"><span className="cr-vazio-ic"><Ic size={28}/></span>
    <div className="cr-vazio-t">{titulo}</div><div className="cr-vazio-s">{texto}</div></div>;
}

/* ============================================================ CSS */
const CSS = `
.cr-root{position:fixed;inset:60px 0 0 0;z-index:60;background:${C.bg};display:flex;flex-direction:column;
  font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:${C.ink};overflow:hidden;}
.cr-embutido{position:static;inset:auto;z-index:auto;height:auto;min-height:640px;border:1px solid ${C.line};border-radius:14px;overflow:hidden;}
.cr-stripe{display:flex;height:4px;flex-shrink:0;} .cr-stripe span{flex:1;}
.cr-head{background:#fff;border-bottom:1px solid ${C.line};padding:12px 22px;display:flex;justify-content:space-between;align-items:center;gap:16px;flex-wrap:wrap;flex-shrink:0;}
.cr-title{display:flex;align-items:center;gap:12px;}
.cr-badge{width:38px;height:38px;border-radius:10px;background:${C.primarySoft};color:${C.primary};display:flex;align-items:center;justify-content:center;flex-shrink:0;}
.cr-h1{font-size:17px;font-weight:800;color:${C.navy};letter-spacing:-.01em;}
.cr-h2{font-size:11.5px;color:${C.sub};margin-top:2px;max-width:640px;}
.cr-headr{display:flex;align-items:flex-end;gap:12px;}
.cr-rev{display:flex;flex-direction:column;gap:2px;}
.cr-rev label{font-size:9.5px;text-transform:uppercase;letter-spacing:.06em;color:${C.faint};font-weight:700;}
.cr-rev input{height:34px;border:1px solid ${C.line};border-radius:8px;padding:0 11px;font-size:13px;font-family:inherit;width:170px;background:${C.surface};color:${C.ink};}
.cr-x{width:34px;height:34px;border:1px solid ${C.line};background:#fff;border-radius:8px;cursor:pointer;display:flex;align-items:center;justify-content:center;color:${C.sub};}
.cr-x:hover{background:${C.bg};}
.cr-databar{background:#fff;border-bottom:1px solid ${C.line};padding:8px 22px;display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;flex-shrink:0;}
.cr-dstat{font-size:12px;color:${C.sub};display:flex;align-items:center;gap:6px;}
.cr-dactions{display:flex;gap:8px;}
.cr-ghost{display:inline-flex;align-items:center;gap:6px;height:32px;border:1px solid ${C.line};background:#fff;border-radius:8px;padding:0 12px;font-size:12.5px;font-family:inherit;color:${C.ink};cursor:pointer;}
.cr-ghost:hover{border-color:${C.primary};color:${C.primary};}
.cr-cnt{background:${C.primary};color:#fff;border-radius:20px;padding:1px 7px;font-size:11px;margin-left:2px;}
.cr-tabs{background:#fff;border-bottom:1px solid ${C.line};padding:0 22px;display:flex;gap:4px;flex-shrink:0;}
.cr-tab{display:inline-flex;align-items:center;gap:7px;background:none;border:0;border-bottom:2px solid transparent;padding:11px 14px;font-size:13.5px;font-family:inherit;color:${C.sub};cursor:pointer;font-weight:600;}
.cr-tab:hover{color:${C.ink};} .cr-tab.on{color:${C.primary};border-bottom-color:${C.primary};}
.cr-tabn{background:${C.bg};border:1px solid ${C.line};border-radius:20px;padding:0 7px;font-size:11px;color:${C.sub};}
.cr-tab.on .cr-tabn{background:${C.primarySoft};border-color:transparent;color:${C.primary};}
.cr-body{flex:1;display:flex;min-height:0;overflow:hidden;}
.cr-main{flex:1;overflow:auto;padding:20px 22px 40px;min-width:0;}

/* painel lateral recolhível — árvore do banco */
.cr-tree{width:288px;flex-shrink:0;background:#fff;border-right:1px solid ${C.line};display:flex;flex-direction:column;min-height:0;}
.cr-tree.collapsed{width:44px;align-items:center;padding-top:14px;}
.cr-tree-open{writing-mode:vertical-rl;transform:rotate(180deg);display:flex;align-items:center;gap:8px;background:none;border:0;cursor:pointer;color:${C.sub};font-size:12px;font-weight:700;font-family:inherit;padding:8px 4px;}
.cr-tree-open:hover{color:${C.primary};}
.cr-tree-h{display:flex;align-items:center;justify-content:space-between;padding:12px 12px 8px;border-bottom:1px solid ${C.line};}
.cr-tree-t{display:flex;align-items:center;gap:7px;font-size:12px;font-weight:800;color:${C.navy};text-transform:uppercase;letter-spacing:.03em;}
.cr-tree-x{width:26px;height:26px;border:1px solid ${C.line};background:#fff;border-radius:7px;cursor:pointer;display:flex;align-items:center;justify-content:center;color:${C.sub};}
.cr-tree-x:hover{border-color:${C.primary};color:${C.primary};}
.cr-tree-search{display:flex;align-items:center;gap:7px;margin:10px 12px;background:${C.bg};border:1px solid ${C.line};border-radius:8px;padding:0 10px;height:32px;}
.cr-tree-search input{flex:1;border:0;outline:0;background:none;font-size:12.5px;font-family:inherit;color:${C.ink};}
.cr-tree-body{flex:1;overflow:auto;padding:0 8px 16px;}
.cr-trow{width:100%;display:flex;align-items:center;gap:6px;background:none;border:0;cursor:pointer;font-family:inherit;text-align:left;border-radius:7px;padding:6px 7px;color:${C.ink};font-size:12.5px;}
.cr-trow:hover{background:${C.bg};}
.cr-trow.lvl0{font-weight:700;margin-top:4px;}
.cr-trow.lvl1{padding-left:20px;color:${C.sub};}
.cr-trow.lvl2{padding-left:40px;color:${C.faint};cursor:default;font-size:12px;}
.cr-trow.lvl2:hover{background:none;}
.cr-tdot{width:9px;height:9px;border-radius:50%;flex-shrink:0;}
.cr-tlabel{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.cr-tcount{flex-shrink:0;background:${C.bg};color:${C.faint};border-radius:20px;padding:1px 8px;font-size:10.5px;font-weight:700;}
.cr-trow.lvl0 .cr-tcount{background:${C.primarySoft};color:${C.primaryDark};}
.cr-tree-empty{padding:20px 12px;font-size:12px;color:${C.faint};line-height:1.5;}

/* pills destacadas (macro / categoria / serviço) */
.cr-pill-mac{border-left:3px solid ${C.primary};} .cr-pill-mac em{color:${C.primary}!important;}
.cr-pill-cat{border-left:3px solid ${C.amber};} .cr-pill-cat em{color:${C.amber}!important;}
.cr-pill-serv{border-left:3px solid ${C.green};} .cr-pill-serv em{color:${C.green}!important;}
.cr-cols-lbl{display:flex;align-items:center;gap:9px;font-size:12px;color:${C.faint};margin:4px 0 9px;}
.cr-tagent{background:${C.navy};color:#fff;font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.06em;padding:3px 9px;border-radius:20px;}

/* botão voltar ao catálogo (rotulado) */
.cr-x{width:auto!important;gap:6px;padding:0 12px 0 9px!important;font-size:12.5px;font-weight:700;font-family:inherit;color:${C.sub};}
.cr-x:hover{border-color:${C.primary};color:${C.primary};}

/* aba Serviços — cartões no padrão dos Similares */
.cr-serv-row{border-radius:12px!important;box-shadow:0 1px 2px rgba(16,40,80,.04);}
.cr-serv-ctx{display:flex;flex-direction:column;gap:8px;}
.cr-serv-pills{display:flex;gap:8px;flex-wrap:wrap;align-items:center;}
.cr-serv-entrow{display:flex;align-items:flex-start;gap:9px;}
.cr-serv-entrow .cr-serv-ent{margin:0;font-weight:600;color:${C.ink};}

/* similares */
.cr-sim{max-width:860px;margin:0 auto;}
.cr-simhead{margin-bottom:14px;}
.cr-bar{height:7px;background:${C.line};border-radius:20px;overflow:hidden;} .cr-bar i{display:block;height:100%;background:${C.primary};transition:width .3s;}
.cr-barrow{display:flex;justify-content:space-between;font-size:12px;color:${C.sub};margin-top:5px;} .cr-barrow b{color:${C.primaryDark};}
.cr-card{background:#fff;border:1px solid ${C.line};border-radius:14px;padding:18px 20px 16px;box-shadow:0 1px 2px rgba(0,0,0,.03);}
.cr-meta{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:8px;}
.cr-pill{background:${C.bg};border-radius:20px;padding:3px 11px;font-size:12px;color:${C.sub};}
.cr-pill em{font-style:normal;font-size:9px;text-transform:uppercase;letter-spacing:.05em;color:${C.faint};font-weight:700;margin-right:6px;}
.cr-sim-pct{margin-left:auto;font-weight:700;color:${C.primaryDark};background:${C.primarySoft};padding:3px 11px;border-radius:20px;font-size:12px;}
.cr-serv{font-size:12px;color:${C.faint};margin-bottom:12px;padding-bottom:12px;border-bottom:1px solid ${C.line};}
.cr-cols{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
@media(max-width:640px){.cr-cols{grid-template-columns:1fr;}}
.cr-opt{text-align:left;border:1.5px solid ${C.line};border-radius:10px;padding:13px;cursor:pointer;background:#fff;transition:.12s;font-family:inherit;}
.cr-opt:hover{border-color:${C.primary};} .cr-opt.sel{border-color:${C.primary};background:${C.primarySoft};}
.cr-opt-top{display:flex;align-items:center;gap:8px;margin-bottom:8px;}
.cr-key{width:22px;height:22px;border-radius:6px;background:${C.ink};color:#fff;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;}
.cr-opt.sel .cr-key{background:${C.primary};}
.cr-cod{font-size:12px;color:${C.faint};font-family:'JetBrains Mono',monospace;}
.cr-b{font-size:10.5px;font-weight:700;padding:2px 8px;border-radius:6px;margin-left:auto;}
.cr-b-ativa{background:${C.greenSoft};color:${C.green};} .cr-b-inativa{background:#F1EFE8;color:#5F5E5A;} .cr-b-nao{background:${C.amberSoft};color:${C.amber};}
.cr-txt{font-size:14.5px;line-height:1.5;}
.cr-df{background:${C.amberSoft};color:${C.amber};border-radius:3px;padding:0 2px;font-weight:600;}
.cr-hint{margin-top:12px;font-size:12.5px;color:${C.sub};background:${C.bg};border-radius:8px;padding:8px 11px;} .cr-hint b{color:${C.amber};}
.cr-acts{display:grid;grid-template-columns:1fr 1fr;gap:9px;margin-top:14px;}
@media(max-width:640px){.cr-acts{grid-template-columns:1fr;}}
.cr-act{display:flex;align-items:center;gap:9px;border:1.5px solid ${C.line};background:#fff;border-radius:9px;padding:11px 12px;font-size:13.5px;font-family:inherit;color:${C.ink};cursor:pointer;text-align:left;transition:.12s;}
.cr-act:hover{border-color:${C.sub};background:#FAFBFC;}
.cr-act .cr-k{width:20px;height:20px;border-radius:5px;background:${C.bg};color:${C.sub};display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;flex-shrink:0;}
.cr-act.dup:hover{border-color:${C.coral};} .cr-act.merge:hover{border-color:${C.amber};}
.cr-mergebox{margin-top:12px;background:${C.amberSoft};border:1px solid #EBD3A8;border-radius:10px;padding:12px;}
.cr-mergel{font-size:12.5px;color:${C.amber};font-weight:700;margin-bottom:7px;}
.cr-mergebox textarea{width:100%;min-height:66px;border:1.5px solid #E7C77A;border-radius:8px;padding:10px;font-family:inherit;font-size:14px;resize:vertical;color:${C.ink};}
.cr-row2{margin-top:12px;} .cr-row2 button{border:1px solid ${C.line};background:#fff;border-radius:8px;height:36px;padding:0 16px;font-size:13px;font-family:inherit;color:${C.sub};cursor:pointer;} .cr-row2 button:hover{background:#FAFBFC;}
.cr-primary{display:inline-flex;align-items:center;gap:7px;margin-top:9px;height:38px;border:0;background:${C.primary};color:#fff;border-radius:8px;padding:0 16px;font-size:13.5px;font-family:inherit;font-weight:600;cursor:pointer;}
.cr-primary:hover{background:${C.primaryDark};} .cr-primary:disabled{opacity:.5;cursor:default;}
.cr-primary.lg{height:44px;font-size:14.5px;width:100%;justify-content:center;margin-top:16px;}
.cr-primary.blocked{background:${C.coral};}

.cr-done{text-align:center;padding:60px 20px;max-width:520px;margin:0 auto;}
.cr-done-ic{width:74px;height:74px;border-radius:50%;background:${C.greenSoft};color:${C.green};display:flex;align-items:center;justify-content:center;margin:0 auto 16px;}
.cr-done-t{font-size:19px;font-weight:700;} .cr-done-s{font-size:13.5px;color:${C.sub};margin:6px 0 18px;}

/* serviços */
.cr-serv-wrap{max-width:900px;margin:0 auto;}
.cr-serv-intro{display:flex;align-items:center;gap:8px;background:${C.greenSoft};border:1px solid #BFE2C5;border-radius:10px;padding:10px 13px;font-size:12.5px;color:#1B5E27;margin-bottom:14px;}
.cr-serv-intro b{color:${C.green};}
.cr-serv-search{display:flex;align-items:center;gap:9px;background:#fff;border:1px solid ${C.line};border-radius:10px;padding:0 13px;height:42px;margin-bottom:12px;}
.cr-serv-search input{flex:1;border:0;outline:0;font-size:14px;font-family:inherit;background:none;color:${C.ink};}
.cr-serv-count{font-size:12px;color:${C.faint};}
.cr-serv-list{display:flex;flex-direction:column;gap:10px;}
.cr-serv-row{background:#fff;border:1px solid ${C.line};border-radius:12px;padding:14px 16px;display:grid;grid-template-columns:1fr;gap:12px;}
.cr-serv-row.ok{border-color:#BFE2C5;background:${C.greenSoft};}
.cr-serv-path{font-size:11px;color:${C.faint};display:flex;align-items:center;gap:3px;flex-wrap:wrap;}
.cr-serv-ent{font-size:14.5px;line-height:1.45;margin:3px 0;}
.cr-serv-cod{font-size:11px;color:${C.faint};font-family:'JetBrains Mono',monospace;}
.cr-serv-inrow{display:flex;gap:8px;flex-wrap:wrap;align-items:center;}
.cr-serv-in{flex:1;min-width:220px;height:38px;border:1.5px solid ${C.line};border-radius:8px;padding:0 11px;font-size:13.5px;font-family:inherit;color:${C.ink};}
.cr-serv-in:focus{border-color:${C.primary};outline:0;}
.cr-serv-ia{display:inline-flex;align-items:center;gap:6px;height:38px;border:1px solid ${C.primary};background:${C.primarySoft};color:${C.primaryDark};border-radius:8px;padding:0 13px;font-size:12.5px;font-family:inherit;font-weight:600;cursor:pointer;}
.cr-serv-ia:disabled{opacity:.6;cursor:default;}
.cr-serv-ok{display:inline-flex;align-items:center;gap:6px;height:38px;border:0;background:${C.green};color:#fff;border-radius:8px;padding:0 14px;font-size:12.5px;font-family:inherit;font-weight:600;cursor:pointer;}
.cr-serv-just{font-size:12px;color:${C.sub};margin-top:8px;background:${C.bg};border-radius:7px;padding:7px 10px;} .cr-serv-just b{color:${C.ink};}
.cr-serv-err{font-size:12px;color:${C.coral};margin-top:8px;background:${C.coralSoft};border-radius:7px;padding:7px 10px;}
.cr-serv-done{display:flex;align-items:center;gap:8px;color:${C.green};font-weight:600;font-size:14px;} .cr-serv-done span{font-size:11px;background:#fff;border:1px solid #BFE2C5;border-radius:20px;padding:1px 8px;font-weight:700;}

/* nova */
.cr-nova{max-width:1040px;margin:0 auto;display:grid;grid-template-columns:1fr 1fr;gap:18px;align-items:start;}
@media(max-width:820px){.cr-nova{grid-template-columns:1fr;}}
.cr-nova-form,.cr-nova-check{background:#fff;border:1px solid ${C.line};border-radius:14px;padding:18px 20px;}
.cr-nova-t{display:flex;align-items:center;gap:8px;font-size:15px;font-weight:700;color:${C.navy};}
.cr-nova-sub{font-size:12.5px;color:${C.sub};line-height:1.5;margin:8px 0 16px;}
.cr-l{display:block;font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:${C.faint};font-weight:700;margin:14px 0 6px;}
.cr-lvl{text-transform:none;letter-spacing:0;color:${C.primary};background:${C.primarySoft};border-radius:20px;padding:1px 8px;font-size:10px;margin-left:6px;}
.cr-in,.cr-ta{width:100%;border:1.5px solid ${C.line};border-radius:8px;padding:10px 11px;font-size:14px;font-family:inherit;color:${C.ink};}
.cr-in:focus,.cr-ta:focus{border-color:${C.primary};outline:0;}
.cr-ta{min-height:70px;resize:vertical;}
.cr-natrow{display:flex;gap:7px;flex-wrap:wrap;}
.cr-natchip{border:1.5px solid ${C.line};background:#fff;border-radius:20px;padding:6px 14px;font-size:12.5px;font-family:inherit;cursor:pointer;color:${C.sub};font-weight:600;}
.cr-check-h{display:flex;align-items:center;gap:8px;font-size:14px;font-weight:700;color:${C.navy};margin-bottom:12px;}
.cr-check-empty{font-size:13px;color:${C.faint};padding:24px 0;text-align:center;}
.cr-verd{display:flex;align-items:flex-start;gap:8px;border-radius:10px;padding:11px 13px;font-size:13px;line-height:1.45;margin-bottom:14px;font-weight:500;}
.v-ok{background:${C.greenSoft};color:#1B5E27;} .v-revisar{background:${C.primarySoft};color:${C.primaryDark};}
.v-atencao{background:${C.amberSoft};color:${C.amber};} .v-bloqueio{background:${C.coralSoft};color:${C.coral};}
.cr-cands{display:flex;flex-direction:column;gap:9px;}
.cr-cand{border:1px solid ${C.line};border-radius:10px;padding:11px 12px;}
.cr-cand-top{display:flex;align-items:center;gap:8px;margin-bottom:6px;}
.cr-cand-sim{font-size:12px;font-weight:800;border-radius:6px;padding:2px 8px;}
.cr-cand-cod{font-size:11px;color:${C.faint};font-family:'JetBrains Mono',monospace;}
.cr-cand-levels{margin-left:auto;display:flex;gap:4px;}
.cr-cand-levels span{width:20px;height:20px;border-radius:5px;background:${C.bg};color:${C.faint};display:flex;align-items:center;justify-content:center;font-size:10.5px;font-weight:800;border:1px solid ${C.line};}
.cr-cand-levels span.hit{background:${C.greenSoft};color:${C.green};border-color:#BFE2C5;}
.cr-cand-ent{font-size:13.5px;line-height:1.45;}
.cr-cand-path{font-size:11px;color:${C.faint};margin-top:5px;display:flex;align-items:center;gap:3px;flex-wrap:wrap;}

.cr-vazio{max-width:460px;margin:60px auto;text-align:center;}
.cr-vazio-ic{width:64px;height:64px;border-radius:16px;background:${C.primarySoft};color:${C.primary};display:flex;align-items:center;justify-content:center;margin:0 auto 14px;}
.cr-vazio-t{font-size:17px;font-weight:700;} .cr-vazio-s{font-size:13px;color:${C.sub};line-height:1.55;margin-top:6px;}

.cr-toast{position:fixed;bottom:22px;left:50%;transform:translateX(-50%);background:${C.ink};color:#fff;padding:11px 18px;border-radius:9px;font-size:13px;z-index:70;box-shadow:0 6px 24px rgba(0,0,0,.22);}

/* botões voltar (Central): X ícone + botão destacado */
.cr-x{width:34px!important;height:34px;padding:0!important;gap:0;color:${C.sub};}
.cr-back{display:inline-flex;align-items:center;gap:6px;height:34px;border:1px solid ${C.primary};background:${C.primary};color:#fff;border-radius:8px;padding:0 13px;font-size:12.5px;font-weight:700;font-family:inherit;cursor:pointer;}
.cr-back:hover{background:${C.primaryDark};border-color:${C.primaryDark};}


/* destaques leves (macro/categoria/serviço) */
.cr-meta-i{font-size:11.5px;color:${C.sub};display:inline-flex;align-items:baseline;gap:5px;}
.cr-meta-i em{font-style:normal;font-size:9px;text-transform:uppercase;letter-spacing:.04em;color:${C.faint};font-weight:700;}
.cr-serv-meta{display:flex;gap:16px;flex-wrap:wrap;}
.cr-serv-meta span{font-size:11.5px;color:${C.sub};display:inline-flex;gap:5px;align-items:baseline;}
.cr-serv-meta em{font-style:normal;font-size:9px;text-transform:uppercase;letter-spacing:.04em;color:${C.faint};font-weight:700;}


/* encadeamento macro › categoria › serviço */
.cr-chain{display:inline-flex;align-items:center;gap:7px;flex-wrap:wrap;}
.cr-chain-i{display:inline-flex;flex-direction:column;gap:1px;background:${C.bg};border:1px solid ${C.line};border-radius:8px;padding:4px 10px;font-size:12.5px;color:${C.ink};font-weight:600;line-height:1.15;}
.cr-chain-i em{font-style:normal;font-size:8.5px;text-transform:uppercase;letter-spacing:.05em;color:${C.faint};font-weight:700;}
.cr-chain-sep{color:${C.faint};flex-shrink:0;}
/* filtro na árvore */
.cr-trow.lvl1{display:flex;align-items:center;gap:0;padding:0!important;}
.cr-texp{border:0;background:none;cursor:pointer;color:${C.sub};display:flex;padding:6px 4px 6px 20px;}
.cr-tfilt{flex:1;display:flex;align-items:center;gap:6px;border:0;background:none;cursor:pointer;font-family:inherit;text-align:left;color:inherit;border-radius:7px;padding:6px 7px;}
.cr-tfilt:hover{background:${C.primarySoft};}
.cr-trow.lvl2.cr-tfilt{width:100%;padding-left:40px;color:${C.faint};}
.cr-trow.filt>.cr-tfilt,.cr-trow.lvl2.filt{background:${C.primarySoft};color:${C.primaryDark};font-weight:700;}
.cr-tfilt-active{display:flex;align-items:center;gap:8px;margin:0 12px 8px;background:${C.primarySoft};border:1px solid ${C.primary};border-radius:8px;padding:6px 10px;font-size:11.5px;color:${C.primaryDark};}
.cr-tfilt-active span{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.cr-tfilt-active button{border:0;background:none;cursor:pointer;color:${C.primaryDark};display:flex;}

`;

return CentralRevisao;
})();



/* ===== módulo embutido: OrganogramaDimensionamento (tela cheia) =====
   Sincronizado com organograma-completo.jsx (merge de 14/07/2026).
   Wiring: aberto por setOrganogramaOpen(true) — fechado por onClose. ===== */
/* ============================================================
   TOKENS DE DESIGN (mesmos usados no resto do portal)
============================================================ */
const T = {
  bg: "#F4F6F9", surface: "#FFFFFF", ink: "#1B1B1B", sub: "#55606E", faint: "#8C97A6",
  line: "#E4E7EB", conn: "#CBD3DD", primary: "#1351B4", primaryDark: "#0C326F",
  primarySoft: "#E8EEF9", green: "#168821", greenSoft: "#E3F2E5", navy: "#13315C",
  yellow: "#FFCD07", amber: "#B86E00", amberSoft: "#FBEEDB", coral: "#9E3B1F",
  coralSoft: "#FAECE7", blue: "#155BCB", blueSoft: "#E6EFFB", grey: "#9AA6B2",
  greySoft: "#EDF0F3",
};

const NATUREZA_COLORS = {
  "Finalístico": { cor: "#168821", soft: "#E3F2E5" },
  "Finalística": { cor: "#168821", soft: "#E3F2E5" },
  "Governança":  { cor: "#B86E00", soft: "#FBEEDB" },
  "Suporte":     { cor: "#5F5E5A", soft: "#F1EFE8" },
  "Gestão":      { cor: "#155BCB", soft: "#E6EFFB" },
};
const natureza = (nat) => NATUREZA_COLORS[nat] || { cor: "#8C97A6", soft: "#F1F3F5" };

const STATUS = {
  feito:      { cor: T.green,  soft: T.greenSoft,  txt: "DFT concluído" },
  andamento:  { cor: T.amber,  soft: T.amberSoft,  txt: "DFT em andamento" },
  sem:        { cor: T.grey,   soft: T.greySoft,   txt: "sem DFT" },
  englobado:  { cor: T.green,  soft: T.greenSoft,  txt: "No DFT da coordenação-geral" },
  estr:       { cor: T.sub,    soft: "#fff",       txt: "" },
};

const API_BASE = "https://portal-backend-bftz.onrender.com";

/* ============================================================
   NORMALIZAÇÃO DE TEXTO (para busca/filtro sem acento)
   -------------------------------------------------------------
   norm() já foi declarado no topo do arquivo (linha 227) — não
   redeclarar aqui, senão esbuild falha com "already declared".
============================================================ */

/* ============================================================
   CONSTRUÇÃO DA ÁRVORE — recebe o JSON cru do backend e monta
   os nós com: tipo (orgao/depto/unidade), status do DFT, e a
   HERANÇA de DFT abaixo de uma coordenação-geral (nível 4).
============================================================ */
function buildNode(raw, parentPath, depth, herdaDe) {
  const uid = parentPath ? `${parentPath}/${raw.sigla}` : raw.sigla;
  const temFilhos = Array.isArray(raw.filhos) && raw.filhos.length > 0;

  // se este nó é uma CG (nível 4) com DFT próprio, ele passa a "doar"
  // esse DFT para todos os descendentes — a menos que já venha herdado de mais acima
  const heranca = herdaDe || (depth === 4 && raw.temDFT
    ? { uid, sig: raw.sigla, nome: raw.nome }
    : null);

  const filhos = temFilhos
    ? raw.filhos.map((f) => buildNode(f, uid, depth + 1, heranca))
    : [];

  // colegiados (comitês/comissões/conselhos) não entram na conta de completude
  const totalFilhos = temFilhos
    ? filhos.reduce((s, f) => s + (isColegiado(f) ? 0 : (f.cobTot || (f.tipo === "unidade" ? 1 : 0))), 0)
    : 1;
  const temCobertura = !!raw.temDFT || !!herdaDe;
  const feitoFilhos = temFilhos
    ? filhos.reduce((s, f) => s + (isColegiado(f) ? 0 : (f.cobFeito || 0)), 0)
    : (temCobertura ? 1 : 0);

  const node = {
    uid,
    sig: raw.sigla || raw.codigo,
    nome: raw.nome,
    tipo: temFilhos ? "depto" : "unidade",
    status: raw.temDFT ? "feito" : (herdaDe ? "englobado" : (temFilhos ? "estr" : "sem")),
    depth,
    cobFeito: feitoFilhos,
    cobTot: totalFilhos,
    dimzvl: !temFilhos,
    sel: true,
    filhos,
    cobConcl: feitoFilhos,
    cobAnd: 0,
    englobado: !!herdaDe,
    cov: herdaDe || null,
    // cod: código do organograma (cruzado com DIM_UORG) — usado pelo relatório
    // de unidade para carregar dados reais via /api/organograma-agregado.
    // idU: id_unidade do SISDIP — usado para buscar entregas via
    // /api/entregas-unidade/:idU. Ver GUIA-integracao-dados-sisdip.md.
    cod: raw.codigo,
    idU: raw.idUnidadeSisdip ?? null,
  };

  // dados de dimensionamento (efetivo/estimado/necessidade/observação)
  if (raw.temDFT && raw.indicadores) {
    const ind = raw.indicadores;
    node.dash = {
      nec: (ind.qt_media_pessoas_atual ?? 0) - (ind.qt_media_pessoas_estimada ?? 0),
      efet: ind.qt_media_pessoas_efetiva ?? 0,
      estim: ind.qt_media_pessoas_estimada ?? 0,
      abono: 0,
      apoio: ind.qt_pessoas_apoio ?? 0,
      dft: ind.ds_dimensionamento || `${raw.sigla}. Ref: ${ind.dt_fim || ""}`,
      periodo: ind.dt_fim || "",
      obs: ind.observacao || ind.ds_observacao || null,
    };
  }
  return node;
}

/* Carrega o organograma + entregas do backend */
async function carregarOrganograma() {
  const resp = await fetch(`${API_BASE}/api/organograma`);
  if (!resp.ok) throw new Error(`Falha ao carregar organograma: HTTP ${resp.status}`);
  const raw = await resp.json();
  const entregasMap = {};

  // coleta entregas embutidas em qualquer nível da árvore recebida
  function coletaEntregas(n) {
    if (!n || typeof n !== "object") return;
    const sig = n.sigla || n.sig;
    if (sig && Array.isArray(n.entregas) && n.entregas.length) {
      entregasMap[sig] = n.entregas.map((x) => ({
        cd: x.cd || x.cod || x.codigo || "",
        nm: x.nm || x.ent || x.entrega || x.nome || "",
        serv: x.serv || x.servico || "",
        mac: x.mac || x.macro || x.macroprocesso || "",
        cat: x.cat || x.categoria || "",
        nat: x.nat || x.natureza || "",
        noBanco: x.noBanco !== false,
      }));
    }
    (n.filhos || n.children || []).forEach(coletaEntregas);
  }
  const orgaosRaw = Array.isArray(raw) ? raw : raw.orgaos || [];
  orgaosRaw.forEach(coletaEntregas);
  if (raw && raw.entregas && typeof raw.entregas === "object" && !Array.isArray(raw.entregas)) {
    Object.assign(entregasMap, raw.entregas);
  }
  // reforço: tenta um endpoint dedicado de entregas, se existir
  try {
    const r2 = await fetch(`${API_BASE}/api/entregas`);
    if (r2.ok) {
      const e2 = await r2.json();
      if (e2 && typeof e2 === "object" && !Array.isArray(e2)) {
        for (const k in e2) if (!entregasMap[k] && Array.isArray(e2[k])) entregasMap[k] = e2[k];
      }
    }
  } catch (_e) { /* endpoint opcional */ }

  return {
    orgaos: orgaosRaw.map((m) => ({ ...buildNode(m, "", 1, null), tipo: "orgao" })),
    entregas: entregasMap,
  };
}

/* Necessidade de pessoal → rótulo/cor */
function leituraNecessidade(nec) {
  if (nec == null) return { tipo: "na", cor: T.faint, soft: T.bg, Icon: AlertCircle };
  if (nec > 0) return { tipo: "deficit", curto: `+${nec}`, cor: T.coral, soft: T.coralSoft, Icon: TrendingUp };
  if (nec < 0) return { tipo: "superavit", curto: `${nec}`, cor: T.blue, soft: T.blueSoft, Icon: CheckCircle2 };
  return { tipo: "equilibrio", curto: "0", cor: T.green, soft: T.greenSoft, Icon: Sparkles };
}

/* Assessorias/gabinetes ficam como "staff" ao lado do nó pai, não como galho */
function isStaff(n) {
  const sig = n.sig || "", nome = n.nome || "";
  if (["ASCOM", "AECI", "ASPAR", "ASPAD", "CONJUR", "OUV", "CGGM", "GATEL", "AECON", "AECOR"].includes(sig)) return true;
  return /^\s*(assessoria|gabinete|ouvidoria|consultoria|corregedoria)\b/i.test(nome)
    || /assessor|consultoria jur|conjur|ouvidor|corregedor/i.test(nome);
}
const podeClicar = (n) => n.tipo === "unidade" || n.tipo === "ramo-sel";

/* Comitês/comissões/conselhos → agrupam num cartão "Colegiados" */
function isColegiado(n) {
  if (n.tipo === "orgao") return false;
  const nome = n.nome || "";
  const texto = nome + " " + (n.sig || "");
  return /\b(comit|comiss|comis|conselho|c[aâ]mara|junta|colegiad)/i.test(texto)
    || /^\s*com\b/i.test(nome); // abreviação "COM ..." usada no SIAPE
}

/* ============================================================
   HELPERS DE VISÃO PERSONALIZADA (modo edição: mover/ocultar)
============================================================ */
const clone = (o) => JSON.parse(JSON.stringify(o));
function acharNo(raiz, uid) {
  let achado = null, pai = null;
  (function andar(n, p) {
    if (achado) return;
    if (n.uid === uid) { achado = n; pai = p; return; }
    (n.filhos || []).forEach((c) => andar(c, n));
  })(raiz, null);
  return { node: achado, parent: pai };
}
function ehDescendente(a, b) {
  let achou = false;
  (function andar(n) { if (achou) return; if (n === b) { achou = true; return; } (n.filhos || []).forEach(andar); })(a);
  return achou;
}
function cloneRaso(n) {
  const c = {};
  for (const k in n) if (k !== "filhos") c[k] = n[k];
  c.filhos = [];
  return c;
}
/** aplica remoções e movimentações de uma visão personalizada sobre a árvore oficial */
function aplicarPersonalizacao(raiz, ops) {
  if (!ops || (!(ops.removed?.length) && !(ops.moved && Object.keys(ops.moved).length))) return raiz;
  const r = clone(raiz);
  const removidos = new Set(ops.removed || []);
  (function podar(n) {
    if (!n.filhos) return;
    n.filhos = n.filhos.filter((c) => !removidos.has(c.uid));
    n.filhos.forEach(podar);
  })(r);
  const movidos = ops.moved || {};
  for (const id in movidos) {
    const origem = acharNo(r, id), destino = acharNo(r, movidos[id]);
    if (!origem.node || !destino.node || !origem.parent) continue;
    if (ehDescendente(origem.node, destino.node) || origem.node === destino.node) continue; // evita ciclo
    origem.parent.filhos = origem.parent.filhos.filter((c) => c !== origem.node);
    destino.node.filhos = destino.node.filhos || [];
    destino.node.filhos.push(origem.node);
  }
  return r;
}
/** filtra a árvore para mostrar só quem tem DFT + o caminho até lá */
function filtrarSoComDft(n) {
  const filhos = (n.filhos || []).map(filtrarSoComDft).filter(Boolean);
  if (n.status === "feito" || n.tipo === "orgao" || filhos.length > 0) {
    const c = cloneRaso(n);
    c.filhos = filhos;
    return c;
  }
  return null;
}

/* ============================================================
   FORMATAÇÃO DE PERÍODO ("mm/aaaa" ou "d–d/aaaa" → por extenso)
============================================================ */
const MESES = ["", "janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
  "agosto", "setembro", "outubro", "novembro", "dezembro"];
function formatarPeriodo(v) {
  if (!v) return v;
  const s = String(v);
  let m = s.match(/(\d{1,2})\s*[–-]\s*(\d{1,2})\s*\/\s*(\d{4})/);
  if (m) {
    const a = +m[1], b = +m[2], ano = m[3];
    if (MESES[a] && MESES[b]) return a === b ? `${MESES[a]} de ${ano}` : `${MESES[a]} a ${MESES[b]} de ${ano}`;
  }
  m = s.match(/(\d{1,2})\s*\/\s*(\d{4})/);
  return m && MESES[+m[1]] ? `${MESES[+m[1]]} de ${m[2]}` : v;
}

/* ============================================================
   NÓ DA ÁRVORE (recursivo) — órgão / depto / ramo / colegiado / unidade
============================================================ */
function NoOrganograma({
  n, raiz, parentUid, colapsados, toggle, sel, setSel, match, qOn,
  lente, metricas, onTipEnter, onTipLeave, edit,
}) {
  const todosFilhos = n.filhos || [];
  const staff = todosFilhos.filter(isStaff);
  const normais0 = todosFilhos.filter((c) => !staff.includes(c));
  const isCluster = n.tipo === "cluster";
  const colegiados = isCluster ? [] : normais0.filter(isColegiado);

  // agrupa colegiados num único cartão "Colegiados" quando há 2+
  const normais = colegiados.length >= 2
    ? [
        ...normais0.filter((c) => !colegiados.includes(c)),
        {
          uid: "clu:" + n.uid, tipo: "cluster", sig: "Colegiados",
          nome: `${colegiados.length} unidades colegiadas`, filhos: colegiados,
          cobTot: 0, cobConcl: 0, cobAnd: 0,
        },
      ]
    : normais0;

  const temFilhosVisiveis = normais.length > 0;
  const colapsado = isCluster ? !colapsados.has(n.uid) : colapsados.has(n.uid);
  const selecionado = sel && sel.uid === n.uid;
  const esmaecido = qOn && n.tipo === "unidade" && !match(n);
  const cobertos = (n.cobConcl || 0) + (n.cobAnd || 0);

  let classe = "oc-node", conteudo = null, clicavel = false;

  if (n.tipo === "orgao") {
    classe += " oc-orgao";
    conteudo = (
      <Fragment>
        <div className="oc-sig">{n.sig}</div>
        <div className="oc-nome">{n.nome}</div>
      </Fragment>
    );
  } else if (n.tipo === "depto") {
    classe += " oc-depto" + (n.dash ? " oc-depto-dft" : "");
    clicavel = !!n.dash; // coordenações-gerais com DFT próprio abrem o painel
    conteudo = (
      <Fragment>
        <span className="oc-depto-sig">{n.sig}</span>
        {n.dash && <CheckCircle2 size={11} color={T.green} />}
        <span className="oc-depto-cob">{cobertos}/{n.cobTot}</span>
      </Fragment>
    );
  } else if (n.tipo === "cluster") {
    classe += " oc-cluster";
    conteudo = (
      <Fragment>
        <div className="oc-cluster-h"><Layers size={13} /><span className="oc-cluster-t">Colegiados</span></div>
        <div className="oc-cluster-sub">{n.nome}</div>
      </Fragment>
    );
  } else {
    const st = STATUS[n.status] || STATUS.sem;
    const nec = leituraNecessidade(n.dash ? n.dash.nec : null);
    classe += " oc-unidade oc-" + n.status;
    clicavel = true;
    conteudo = (
      <Fragment>
        <div className="oc-u-head">
          <span className="oc-sig">{n.sig}</span>
          {n.status === "andamento" ? (
            <span className="oc-nec" style={{ background: T.amberSoft, color: T.amber }}><AlertCircle size={10} /></span>
          ) : n.status === "englobado" ? (
            <span className="oc-nec oc-nec-eng" title="Englobada no DFT da coordenação-geral">
              <Layers size={9} /> CG
            </span>
          ) : n.dash ? (
            <span className="oc-nec" style={{ background: nec.soft, color: nec.cor }}>{nec.curto}</span>
          ) : (
            <span className="oc-nec oc-nec-sem">—</span>
          )}
        </div>
        <div className="oc-nome">{n.nome}</div>
      </Fragment>
    );
  }

  const lensStyle = isCluster ? null : estiloLente(n, lente, metricas ? metricas[n.uid] : null);
  const emEdicao = edit && edit.on;

  return (
    <li className={raiz ? "oc-raiz" : ""}>
      <div className="oc-row">
        <div
          className={`${classe}${selecionado ? " on" : ""}${esmaecido ? " dim" : ""}${emEdicao && !isCluster && !raiz ? " oc-editable" : ""}`}
          draggable={!!(emEdicao && !isCluster && !raiz)}
          onDragStart={emEdicao && !isCluster && !raiz ? (e) => { e.stopPropagation(); e.dataTransfer.setData("text/plain", n.uid); } : undefined}
          onDragOver={emEdicao && !isCluster ? (e) => { e.preventDefault(); e.currentTarget.classList.add("oc-dragover"); } : undefined}
          onDragLeave={emEdicao && !isCluster ? (e) => e.currentTarget.classList.remove("oc-dragover") : undefined}
          onDrop={emEdicao && !isCluster ? (e) => {
            e.preventDefault(); e.stopPropagation(); e.currentTarget.classList.remove("oc-dragover");
            const origemUid = e.dataTransfer.getData("text/plain");
            if (origemUid && origemUid !== n.uid) edit.move(origemUid, n.uid);
          } : undefined}
          onClick={emEdicao ? undefined : (isCluster ? () => toggle(n.uid) : (clicavel ? () => setSel(n) : undefined))}
          data-ocuid={n.uid}
          data-ocparent={parentUid || undefined}
          onMouseEnter={(!isCluster && !emEdicao && onTipEnter) ? (e) => onTipEnter(n, e) : undefined}
          onMouseLeave={onTipLeave}
          style={{ ...(clicavel || isCluster || emEdicao ? { cursor: emEdicao && !raiz ? "grab" : "pointer" } : {}), ...(lensStyle || {}) }}
        >
          {emEdicao && !isCluster && !raiz && (
            <button className="oc-del" title="Ocultar desta visão"
              onClick={(e) => { e.stopPropagation(); edit.remove(n.uid); }}>✕</button>
          )}
          {conteudo}
          {temFilhosVisiveis && (
            <button className="oc-toggle" title={colapsado ? "expandir" : "recolher"}
              onClick={(e) => { e.stopPropagation(); toggle(n.uid); }}>
              {colapsado ? <span className="oc-plus">+{normais.length}</span> : <ChevronDown size={13} />}
            </button>
          )}
        </div>

        {staff.length > 0 && !colapsado && (
          <div className="oc-staff">
            {staff.map((s) => {
              const clicavelStaff = podeClicar(s);
              const sSel = sel && sel.uid === s.uid;
              const sLens = estiloLente(s, lente, metricas ? metricas[s.uid] : null);
              return (
                <button key={s.uid} className={`oc-staffchip${sSel ? " on" : ""}`}
                  onMouseEnter={onTipEnter ? (e) => onTipEnter(s, e) : undefined}
                  onMouseLeave={onTipLeave}
                  onClick={clicavelStaff ? () => setSel(s) : undefined}
                  style={{ ...(clicavelStaff ? { cursor: "pointer" } : { cursor: "default" }), ...(sLens || {}) }}>
                  <span className="oc-staff-sig">{s.sig}</span>
                  <span className="oc-staff-nome">{s.nome}</span>
                </button>
              );
            })}
          </div>
        )}
      </div>

      {temFilhosVisiveis && (
        <ColapsavelSuave open={!colapsado}>
          <ul>
            {normais.map((c) => (
              <NoOrganograma key={c.uid} n={c} parentUid={n.uid}
                colapsados={colapsados} toggle={toggle} sel={sel} setSel={setSel}
                match={match} qOn={qOn} lente={lente} metricas={metricas}
                onTipEnter={onTipEnter} onTipLeave={onTipLeave} edit={edit} />
            ))}
          </ul>
        </ColapsavelSuave>
      )}
    </li>
  );
}

/** Grid animado (0fr → 1fr) para abrir/fechar galhos com transição suave */
function ColapsavelSuave({ open, children }) {
  const [montado, setMontado] = useState(open);
  const [aberto, setAberto] = useState(open);
  useEffect(() => {
    if (open) {
      setMontado(true);
      const raf = requestAnimationFrame(() => requestAnimationFrame(() => setAberto(true)));
      return () => cancelAnimationFrame(raf);
    }
    setAberto(false);
    const t = setTimeout(() => setMontado(false), 320);
    return () => clearTimeout(t);
  }, [open]);
  if (!montado) return null;
  return <div className={`oc-kids${aberto ? " open" : ""}`}>{children}</div>;
}

/* ============================================================
   LENTES DE COLORIZAÇÃO (déficit / cobertura / catálogo)
============================================================ */
const LENTES = [
  { id: "status", label: "Status do DFT" },
  { id: "deficit", label: "Necessidade de pessoal" },
  { id: "cobertura", label: "Cobertura do DFT" },
  { id: "catalogo", label: "Qualidade do catálogo" },
];
const LENTE_DESC = {
  status: "cor = situação do dimensionamento",
  deficit: "cor = déficit (coral) ou superávit (azul) de pessoal; intensidade = tamanho relativo",
  cobertura: "cor = proporção de unidades com DFT",
  catalogo: "cor = proporção de entregas sem serviço atribuído",
};
function estiloLente(n, lenteId, metrica) {
  if (!lenteId || lenteId === "status" || n.tipo === "orgao") return null;
  const base = { background: "#fff", borderColor: T.line, borderStyle: "dashed" };

  if (lenteId === "deficit") {
    let nec = null, estim = 0;
    if (n.tipo === "unidade") { if (!n.dash) return base; nec = n.dash.nec; estim = n.dash.estim || 0; }
    else { if (!metrica || !metrica.dashN) return base; nec = metrica.necSum; estim = metrica.estimSum; }
    if (nec === 0) return { background: T.greenSoft, borderColor: T.green + "66", borderStyle: "solid" };
    const razao = estim > 0 ? Math.abs(nec) / estim : 1;
    const alpha = razao < 0.1 ? "1F" : razao < 0.25 ? "3D" : "5C";
    const cor = nec > 0 ? T.coral : T.blue;
    return { background: cor + alpha, borderColor: cor + "88", borderStyle: "solid" };
  }
  if (lenteId === "cobertura") {
    const total = n.cobTot || 0, feitos = (n.cobConcl || 0) + (n.cobAnd || 0);
    if (!total || feitos === 0) return base;
    const p = feitos / total;
    const alpha = p >= 0.85 ? "5C" : p >= 0.5 ? "38" : "1F";
    return { background: T.green + alpha, borderColor: T.green + "88", borderStyle: "solid" };
  }
  if (lenteId === "catalogo") {
    if (!metrica || !metrica.entTot) return base;
    const p = metrica.semServ / metrica.entTot;
    if (p === 0) return { background: T.greenSoft, borderColor: T.green + "66", borderStyle: "solid" };
    if (p <= 0.25) return { background: T.amber + "26", borderColor: T.amber + "77", borderStyle: "solid" };
    if (p <= 0.5) return { background: T.amber + "47", borderColor: T.amber + "99", borderStyle: "solid" };
    return { background: T.coral + "42", borderColor: T.coral + "88", borderStyle: "solid" };
  }
  return null;
}
function legendaLente(lenteId) {
  if (lenteId === "deficit") return [
    [T.coral + "5C", "déficit alto"], [T.coral + "1F", "déficit leve"], [T.greenSoft, "equilíbrio"],
    [T.blue + "3D", "superávit"], [null, "sem dado"],
  ];
  if (lenteId === "cobertura") return [
    [T.green + "5C", "≥ 85%"], [T.green + "38", "50–84%"], [T.green + "1F", "1–49%"], [null, "0%"],
  ];
  if (lenteId === "catalogo") return [
    [T.greenSoft, "100% com serviço"], [T.amber + "26", "até 25% sem"], [T.amber + "47", "até 50% sem"],
    [T.coral + "42", "> 50% sem"], [null, "sem entregas"],
  ];
  return [[T.greenSoft, "DFT concluído"], [T.amberSoft, "em andamento"], [null, "sem DFT"]];
}

/* soma recursiva de métricas (entregas/déficit) por nó, para as lentes e o tooltip */
function agregarMetricas(n, acc) {
  const entregasNo = (n.tipo === "unidade" || n.tipo === "ramo-sel") ? (window.__entregasPorSigla?.[n.sig] || []) : [];
  const m = {
    entTot: entregasNo.length,
    semServ: entregasNo.filter((x) => !x.serv).length,
    foraCat: entregasNo.filter((x) => !x.noBanco).length,
    necSum: 0, estimSum: 0, efetSum: 0, dashN: 0,
  };
  if (n.dash) { m.necSum += n.dash.nec || 0; m.estimSum += n.dash.estim || 0; m.efetSum += n.dash.efet || 0; m.dashN++; }
  (n.filhos || []).forEach((c) => {
    const cm = agregarMetricas(c, acc);
    m.entTot += cm.entTot; m.semServ += cm.semServ; m.foraCat += cm.foraCat;
    m.necSum += cm.necSum; m.estimSum += cm.estimSum; m.efetSum += cm.efetSum; m.dashN += cm.dashN;
  });
  acc[n.uid] = m;
  return m;
}

/* marca uids "abertos por padrão": todo depth>=2 começa expandido */
function abrirPorPadrao(n, set) {
  if (n.filhos && n.filhos.length) {
    if (n.depth >= 2) set.add(n.uid);
    n.filhos.forEach((c) => abrirPorPadrao(c, set));
  }
  return set;
}

/* ============================================================
   CONECTORES ORTOGONAIS (SVG desenhado via medição real do DOM)
============================================================ */
function Conectores({ chartRef, zoom, depA, depB }) {
  const [estado, setEstado] = useState({ paths: [], w: 0, h: 0 });
  const zoomRef = useRef(zoom);
  zoomRef.current = zoom;

  useEffect(() => {
    const box = chartRef.current;
    if (!box) return;
    let raf = 0, prazo = 0, ativo = true;

    function medir() {
      const escala = zoomRef.current || 1;
      const boxRect = box.getBoundingClientRect();
      const nos = box.querySelectorAll("[data-ocuid]");
      const pos = {};
      nos.forEach((el) => {
        if (el.closest(".oc-kids:not(.open)")) return; // ignora galhos recolhidos
        const r = el.getBoundingClientRect();
        pos[el.getAttribute("data-ocuid")] = {
          x: (r.left - boxRect.left + r.width / 2) / escala,
          t: (r.top - boxRect.top) / escala,
          b: (r.top - boxRect.top + r.height) / escala,
        };
      });
      const caminhos = [];
      nos.forEach((el) => {
        const uid = el.getAttribute("data-ocuid"), pUid = el.getAttribute("data-ocparent");
        if (!pUid || !pos[uid] || !pos[pUid]) return;
        const pai = pos[pUid], filho = pos[uid];
        const meio = (pai.b + 6 + filho.t) / 2;
        const ax = pai.x, bx = filho.x, dx = bx - ax;
        const raio = Math.max(0, Math.min(9, Math.abs(dx) / 2, (filho.t - pai.b - 6) / 2));
        let d;
        if (Math.abs(dx) < 1.5) {
          d = `M ${ax.toFixed(1)} ${(pai.b + 6).toFixed(1)} L ${ax.toFixed(1)} ${filho.t.toFixed(1)}`;
        } else {
          const sinal = dx > 0 ? 1 : -1;
          d = `M ${ax.toFixed(1)} ${(pai.b + 6).toFixed(1)} `
            + `L ${ax.toFixed(1)} ${(meio - raio).toFixed(1)} `
            + `Q ${ax.toFixed(1)} ${meio.toFixed(1)} ${(ax + sinal * raio).toFixed(1)} ${meio.toFixed(1)} `
            + `L ${(bx - sinal * raio).toFixed(1)} ${meio.toFixed(1)} `
            + `Q ${bx.toFixed(1)} ${meio.toFixed(1)} ${bx.toFixed(1)} ${(meio + raio).toFixed(1)} `
            + `L ${bx.toFixed(1)} ${filho.t.toFixed(1)}`;
        }
        caminhos.push(d);
      });
      setEstado({ paths: caminhos, w: box.scrollWidth, h: box.scrollHeight });
    }
    function loop() {
      if (!ativo) { raf = 0; return; }
      medir();
      if (performance.now() < prazo) raf = requestAnimationFrame(loop); else raf = 0;
    }
    function agendar() { prazo = performance.now() + 480; if (!raf) loop(); }
    agendar();
    const ro = typeof ResizeObserver !== "undefined" ? new ResizeObserver(agendar) : null;
    ro?.observe(box);
    window.addEventListener("resize", agendar);
    return () => { ativo = false; raf && cancelAnimationFrame(raf); ro?.disconnect(); window.removeEventListener("resize", agendar); };
  }, [depA, depB]);

  if (!estado.w || !estado.h) return null;
  return (
    <svg className="oc-svg" width={estado.w} height={estado.h} viewBox={`0 0 ${estado.w} ${estado.h}`} aria-hidden="true">
      {estado.paths.map((d, i) => (
        <path key={i} d={d} fill="none" stroke={T.conn} strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
      ))}
    </svg>
  );
}

/* ============================================================
   TOOLTIP flutuante ao passar o mouse sobre um nó
============================================================ */
function Tooltip({ tip }) {
  if (!tip) return null;
  const { n, m, x, y } = tip;
  const st = STATUS[n.status] || STATUS.estr;
  const nec = n.dash ? leituraNecessidade(n.dash.nec) : null;
  const cobertos = (n.cobConcl || 0) + (n.cobAnd || 0);
  const left = Math.min(Math.max(x, 155), (typeof window !== "undefined" ? window.innerWidth : 1200) - 155);

  return (
    <div className="oc-tip" style={{ left, top: y - 10 }}>
      <div className="oc-tip-h">
        <span className="oc-tip-sig">{n.sig}</span>
        {st.txt && <span className="oc-tip-st" style={{ background: st.soft, color: st.cor }}>{st.txt}</span>}
      </div>
      <div className="oc-tip-nome">{n.nome}</div>
      {n.dash && (
        <div className="oc-tip-grid">
          <div><b>{n.dash.efet}</b><span>efetivo</span></div>
          <div><b>{n.dash.estim}</b><span>estimado</span></div>
          <div>
            <b style={{ color: nec.tipo === "deficit" ? "#F5A08B" : nec.tipo === "superavit" ? "#8FB6F2" : "#8FD79A" }}>
              {nec.tipo === "deficit" ? `faltam ${n.dash.nec}` : nec.tipo === "superavit" ? `${-n.dash.nec} a mais` : "equilíbrio"}
            </b>
            <span>necessidade</span>
          </div>
          <div><b>{formatarPeriodo(n.dash.periodo)}</b><span>período do DFT</span></div>
        </div>
      )}
      {n.cobTot > 0 && (
        <div className="oc-tip-row"><i style={{ background: T.green }} />{cobertos}/{n.cobTot} unidades com DFT</div>
      )}
      {m && m.entTot > 0 && (
        <div className="oc-tip-row">
          <i style={{ background: m.semServ ? T.yellow : T.green }} />
          {m.semServ} de {m.entTot} entregas sem serviço{m.foraCat ? ` · ${m.foraCat} fora do catálogo` : ""}
        </div>
      )}
      {podeClicar(n) && <div className="oc-tip-hint">clique para abrir o painel da unidade</div>}
    </div>
  );
}

/* ============================================================
   DONUT de completude do DFT (cabeçalho do órgão)
============================================================ */
function DonutCompletude({ concl, and: andamento, tot }) {
  const R = 26, C = 2 * Math.PI * R;
  const pConcl = tot ? concl / tot : 0, pAnd = tot ? andamento / tot : 0;
  const pct = Math.round((pConcl + pAnd) * 100);
  return (
    <div className="od-donut">
      <svg width="66" height="66" viewBox="0 0 66 66">
        <circle cx="33" cy="33" r={R} fill="none" stroke={T.greySoft} strokeWidth="9" />
        <circle cx="33" cy="33" r={R} fill="none" stroke={T.amber} strokeWidth="9"
          strokeDasharray={`${pAnd * C} ${C}`} strokeDashoffset={-pConcl * C}
          transform="rotate(-90 33 33)" strokeLinecap="butt" />
        <circle cx="33" cy="33" r={R} fill="none" stroke={T.green} strokeWidth="9"
          strokeDasharray={`${pConcl * C} ${C}`} transform="rotate(-90 33 33)" strokeLinecap="butt" />
        <text x="33" y="34" textAnchor="middle" fontSize="15" fontWeight="800" fill={T.navy}>{pct}%</text>
        <text x="33" y="45" textAnchor="middle" fontSize="7.5" fill={T.faint}>com DFT</text>
      </svg>
    </div>
  );
}

/* ============================================================
   PAINEL DA UNIDADE — abas "Entregas" e "Relatório completo"
============================================================ */
/* ============================================================
   PAINEL DA UNIDADE — abas "Entregas" e "Relatório completo"
   (versão sincronizada com o bundle.js de produção em 14/07/2026 —
   ver GUIA-frontend-portal.md para o histórico desta correção)
============================================================ */
function PainelUnidade({
  org, unidade, entregas, entregasCarregando, overrides, onInserir, flash, onClose,
  panelH, setPanelH, maxPanel, setMaxPanel, onMontarNoCatalogo,
  origem, onListaPropria,
}) {
  function iniciarArraste(e) {
    e.preventDefault();
    const y0 = e.clientY, altura0 = panelH || 360;
    function mover(ev) { const delta = y0 - ev.clientY; setPanelH(Math.max(160, Math.min((window.innerHeight || 900) - 72, altura0 + delta))); }
    function soltar() { window.removeEventListener("pointermove", mover); window.removeEventListener("pointerup", soltar); }
    window.addEventListener("pointermove", mover); window.addEventListener("pointerup", soltar);
  }

  const dash = unidade.dash;
  const nec = leituraNecessidade(dash ? dash.nec : null);
  const semDftProprio = unidade.status === "sem" || unidade.status === "englobado";

  // O relatório é o conteúdo principal do painel; a lista de entregas que o
  // sustenta fica recolhida abaixo dele, para quem quiser conferir item a item.
  const [entregasAbertas, setEntregasAbertas] = useState(false);
  const [obsAberta, setObsAberta] = useState(false);
  const observacao = dash && dash.obs ? dash.obs : null;

  const [filtro, setFiltro] = useState("");
  const [rascunhos, setRascunhos] = useState({});
  const listaFiltrada = useMemo(() => {
    const q = norm(filtro);
    return entregas
      .map((e) => ({ ...e, serv: overrides[e.cd] || e.serv }))
      .filter((e) => !q || norm(e.nm).includes(q) || norm(e.mac).includes(q) || norm(e.cat).includes(q) || norm(e.serv).includes(q));
  }, [entregas, filtro, overrides]);
  const semServico = listaFiltrada.filter((e) => !e.serv).length;
  const foraCatalogo = listaFiltrada.filter((e) => !e.noBanco).length;
  const status = STATUS[unidade.status] || STATUS.sem;

  // srcDoc do relatório: SEMPRE modo real, usando o código de organograma da
  // própria unidade. Não há mais fallback para unidades-amostra fictícias
  // (cgfor/cgped/cgace/ditec) — se a unidade não tiver `cod`, o relatório
  // mostra "sem dados" em vez de simular uma unidade que não é a selecionada.
  // Ver GUIA-frontend-portal.md, seção "Correção de dados ilustrativos".
  const relatorioSrcDoc = (() => {
    const html = window.__RELATORIO_HTML__ || "";
    const par = (k) => (k == null ? "" : encodeURIComponent(String(k)));
    const hashReal =
      "unit=real" +
      "&codigo=" + par(unidade.cod) +
      "&idU=" + par(unidade.idU) +
      "&sig=" + par(unidade.sig) +
      "&nome=" + par(unidade.nome) +
      "&crumb=" + par(unidade.uid.split("/").join(" › ")) +
      "&temDFT=" + (unidade.status === "feito" ? "1" : "0");
    return html.replace("</head>", `<script>location.hash='${hashReal}';<\/script></head>`);
  })();

  return (
    <div className="od-painel">
      <div className="od-p-grip" onPointerDown={iniciarArraste} title="Arraste para redimensionar o painel">
        <span className="od-p-grip-bar" />
      </div>

      {/* faixa de contexto quando a unidade selecionada está englobada numa CG */}
      {origem && (
        <div className="od-p-engbanner">
          <Layers size={13} />
          <span><b>{origem.sig}</b> está englobada no DFT de <b>{unidade.sig}</b>. Esta lista refere-se ao dimensionamento da coordenação-geral.</span>
          {onListaPropria && <button className="od-p-engbtn" onClick={onListaPropria}>criar lista própria de {origem.sig}</button>}
        </div>
      )}
      {!origem && unidade.status === "englobado" && unidade.cov && (
        <div className="od-p-engbanner">
          <Layers size={13} />
          <span>Lista própria de <b>{unidade.sig}</b>. A unidade permanece englobada no DFT de <b>{unidade.cov.sig}</b>.</span>
        </div>
      )}

      {/* cabeçalho compacto de uma linha */}
      <div className="od-p-head2">
        <div className="od-p-idline">
          <span className="od-p-crumb2">{unidade.uid.split("/").join("  ›  ")}</span>
          <span className="od-p-title2">{unidade.nome}</span>
        </div>
        <div className="od-p-status" style={{ background: status.soft, color: status.cor }}>
          <CheckCircle2 size={12} /> {status.txt}
        </div>
        <button className="od-p-max" onClick={() => setMaxPanel((v) => !v)} title={maxPanel ? "Restaurar tamanho" : "Sobrepor todo o organograma"}>
          {maxPanel ? <Minimize2 size={14} /> : <Maximize2 size={13} />}
        </button>
        <button className="od-p-close" onClick={onClose}><X size={15} /></button>
      </div>

      {/* unidade sem DFT próprio: o caminho de ação vem antes de tudo */}
      {semDftProprio && (
        <div className="od-p-sem">
          <div className="od-p-sem-tx">
            <AlertCircle size={16} /> Esta unidade ainda <b>não realizou DFT próprio</b>. Você pode montar a descrição de área desta unidade no catálogo.
          </div>
          <button className="od-p-montar" onClick={() => onMontarNoCatalogo && onMontarNoCatalogo(org.nome, unidade.nome)}>
            <PlusCircle size={14} /> Montar descrição no catálogo <ArrowLeft size={13} style={{ transform: "rotate(180deg)" }} />
          </button>
        </div>
      )}

      {/* ===== o relatório é o conteúdo principal ===== */}
      <div className="od-p-reportwrap">
        <iframe title="Relatório da unidade" className="od-p-reportframe" srcDoc={relatorioSrcDoc} />
      </div>

      {/* ===== as entregas que sustentam o relatório, recolhidas ===== */}
      {!semDftProprio && (
        <Fragment>
          <button className={"od-p-expand" + (entregasAbertas ? " on" : "")} onClick={() => setEntregasAbertas((v) => !v)}>
            {entregasAbertas ? <ChevronDown size={15} /> : <ChevronRight size={15} />}
            <FileText size={14} />
            <span>Ver as {listaFiltrada.length} entregas desta unidade</span>
            {semServico > 0 && <span className="od-p-expand-chip warn">{semServico} sem serviço</span>}
            {foraCatalogo > 0 && <span className="od-p-expand-chip mut">{foraCatalogo} fora do catálogo</span>}
          </button>

          {entregasAbertas && (
            <Fragment>
              {dash && (
                <div className="od-p-inds">
                  <span><TrendingUp size={11} /> último: <b>{formatarPeriodo(dash.periodo)}</b></span>
                  <span>
                    <nec.Icon size={11} /> necessidade:{" "}
                    <b style={{ color: nec.cor }}>
                      {nec.tipo === "deficit" ? `faltam ${dash.nec}` : nec.tipo === "superavit" ? `${-dash.nec} a mais` : "equilíbrio"}
                    </b>
                  </span>
                  <span><Users size={11} /> <b>{dash.efet}</b> efetivas · est. <b>{dash.estim}</b></span>
                  <button className="od-p-bilink"
                    onClick={() => flash("Integração Power BI — próxima fase. Abrirá o painel do DFT filtrado por esta unidade.")}>
                    <BarChart2 size={11} /> Power BI <ExternalLink size={11} />
                  </button>
                  {observacao && (
                    <button className={"od-p-obsbtn" + (obsAberta ? " on" : "")} onClick={() => setObsAberta((v) => !v)}>
                      ⓘ observação
                    </button>
                  )}
                </div>
              )}

              {obsAberta && observacao && (
                <div className="od-p-obsbox">
                  <span>📝</span>
                  <div><b>Observação do dimensionamento — </b>{observacao}</div>
                </div>
              )}

              {!dash && (
                <div className="od-p-andamento">
                  <TrendingUp size={15} /> Dimensionamento em andamento — sem resultados consolidados. As entregas já levantadas aparecem abaixo.
                </div>
              )}

              <div className="od-tbl-head">
                <div className="od-tbl-t">
                  <FileText size={15} /> Entregas do dimensionamento
                  <span className="od-tbl-chips">
                    <span>{listaFiltrada.length} entregas</span>
                    {semServico > 0 && <span className="warn">{semServico} sem serviço</span>}
                    {foraCatalogo > 0 && <span className="mut">{foraCatalogo} fora do catálogo</span>}
                  </span>
                </div>
                <div className="od-tbl-search">
                  <Search size={14} color={T.faint} />
                  <input value={filtro} onChange={(e) => setFiltro(e.target.value)} placeholder="Filtrar entregas…" />
                </div>
              </div>

              <div className="od-tbl-wrap">
                <table className="od-tbl">
                  <thead>
                    <tr>
                      <th className="col-nat" title="Natureza">Nat.</th>
                      <th className="col-mac">Macroprocesso</th>
                      <th className="col-cat">Categoria de serviço</th>
                      <th className="col-serv">Serviço</th>
                      <th className="col-ent od-th-ent">Entrega</th>
                    </tr>
                  </thead>
                  <tbody>
                    {listaFiltrada.map((e, i) => {
                      const inseridoAgora = !!overrides[e.cd];
                      return (
                        <tr key={e.cd + i} className={e.noBanco ? "" : "fora"}>
                          <td className="col-nat">
                            <span className="od-natdot" title={e.nat || "—"} style={{ background: e.nat ? natureza(e.nat).cor : "#CBD3DD" }} />
                          </td>
                          <td className="col-mac" title={e.mac}>{e.mac || <span className="od-falta">—</span>}</td>
                          <td className="col-cat" title={e.cat}>{e.cat || <span className="od-falta">—</span>}</td>
                          <td className="col-serv">
                            {e.serv ? (
                              <span className={inseridoAgora ? "od-serv-new" : ""}>
                                {e.serv}
                                {inseridoAgora && <span className="od-serv-tag"><CheckCircle2 size={10} /> inserido</span>}
                              </span>
                            ) : (
                              <div className="od-add">
                                <input value={rascunhos[e.cd] || ""} placeholder="Preencher serviço…"
                                  onChange={(ev) => setRascunhos((r) => ({ ...r, [e.cd]: ev.target.value }))}
                                  onKeyDown={(ev) => { if (ev.key === "Enter") { onInserir(unidade, e, rascunhos[e.cd]); setRascunhos((r) => ({ ...r, [e.cd]: "" })); } }} />
                                <button onClick={() => { onInserir(unidade, e, rascunhos[e.cd]); setRascunhos((r) => ({ ...r, [e.cd]: "" })); }}>
                                  <PlusCircle size={13} />
                                </button>
                              </div>
                            )}
                          </td>
                          <td className="col-ent od-ent-cell">
                            <div className="od-ent-nm">{e.nm}</div>
                            <div className="od-ent-cd">{e.cd}{!e.noBanco && <span className="od-fora">fora do catálogo</span>}</div>
                          </td>
                        </tr>
                      );
                    })}
                    {!listaFiltrada.length && (
                      <tr>
                        <td colSpan={5} className="od-tbl-empty">
                          {entregasCarregando ? (
                            <Fragment><span className="od-tbl-loading-spin" />&nbsp;Carregando entregas…</Fragment>
                          ) : (
                            "Nenhuma entrega para o filtro."
                          )}
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </Fragment>
          )}
        </Fragment>
      )}
    </div>
  );
}
/* ============================================================
   APP RAIZ — organograma completo (header, busca, árvore, painel)
============================================================ */
/* ============================================================
   CSS do módulo — extraído do bundle.js de produção. Interpola
   os tokens de T definidos no topo deste arquivo.
============================================================ */
const OD_CSS = `
.od-root{position:fixed;inset:60px 0 0 0;z-index:60;background:${T.bg};display:flex;flex-direction:column;
  font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:${T.ink};overflow:hidden;}
.od-stripe{display:flex;height:4px;flex-shrink:0;} .od-stripe span{flex:1;}
.od-head{background:#fff;border-bottom:1px solid ${T.line};padding:12px 22px;display:flex;justify-content:space-between;align-items:center;gap:16px;flex-wrap:wrap;flex-shrink:0;}
.od-title{display:flex;align-items:center;gap:12px;}
.od-badge{width:38px;height:38px;border-radius:10px;background:${T.primarySoft};color:${T.primary};display:flex;align-items:center;justify-content:center;}
.od-h1{font-size:17px;font-weight:800;color:${T.navy};letter-spacing:-.01em;}
.od-h2{font-size:11.5px;color:${T.sub};margin-top:2px;}

/* seletor de organiza\xE7\xE3o */
.od-orgpick{display:flex;align-items:center;gap:8px;margin-left:18px;padding-left:18px;border-left:1px solid ${T.line};}
.od-orgpick select{height:34px;border:1px solid ${T.line};border-radius:8px;padding:0 10px;font-size:12.5px;font-family:inherit;color:${T.ink};background:#fff;max-width:340px;cursor:pointer;}
.od-orgpick select:hover{border-color:${T.primary};}
.od-orgimport{display:inline-flex;align-items:center;gap:6px;height:34px;border:1px dashed ${T.line};background:${T.bg};border-radius:8px;padding:0 11px;font-size:12px;font-family:inherit;color:${T.sub};cursor:pointer;}
.od-orgimport:hover{border-color:${T.primary};color:${T.primary};}

/* descri\xE7\xE3o de \xE1rea da unidade */
.od-area{margin-top:18px;border-top:1px solid ${T.line};padding-top:16px;}
.od-area-h{display:flex;align-items:center;gap:7px;font-size:13px;font-weight:800;color:${T.navy};margin-bottom:10px;}
.od-area-ta{width:100%;min-height:72px;border:1px solid ${T.line};border-radius:10px;padding:10px 12px;font-size:13px;font-family:inherit;color:${T.ink};resize:vertical;box-sizing:border-box;}
.od-area-ta:focus{outline:none;border-color:${T.primary};}
.od-area-tools{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:10px;}
.od-area-search{display:flex;align-items:center;gap:8px;flex:1;min-width:220px;background:${T.bg};border:1px solid ${T.line};border-radius:9px;padding:0 12px;height:36px;}
.od-area-search input{flex:1;border:0;outline:0;background:none;font-size:13px;font-family:inherit;color:${T.ink};}
.od-area-clear{border:0;background:none;cursor:pointer;color:${T.faint};display:flex;}
.od-area-pull{display:inline-flex;align-items:center;gap:6px;height:36px;border:1px solid ${T.primary};background:${T.primarySoft};color:${T.primaryDark};border-radius:9px;padding:0 13px;font-size:12.5px;font-weight:700;font-family:inherit;cursor:pointer;}
.od-area-pull:hover{background:${T.primary};color:#fff;}
.od-area-results{margin-top:8px;border:1px solid ${T.line};border-radius:10px;overflow:hidden;background:#fff;}
.od-area-res{width:100%;display:flex;flex-direction:column;gap:2px;align-items:flex-start;text-align:left;background:#fff;border:0;border-bottom:1px solid ${T.line};padding:8px 12px;cursor:pointer;font-family:inherit;}
.od-area-res:last-child{border-bottom:0;} .od-area-res:hover{background:${T.bg};}
.od-area-res:disabled{opacity:.5;cursor:default;}
.od-area-res-nm{font-size:12.5px;color:${T.ink};}
.od-area-res-cd{font-size:11px;color:${T.faint};}
.od-area-chips{display:flex;flex-wrap:wrap;gap:8px;margin-top:12px;}
.od-area-empty{font-size:12.5px;color:${T.faint};}
.od-area-chip{display:inline-flex;align-items:center;gap:7px;background:${T.greenSoft};color:${T.navy};border:1px solid ${T.green}44;border-radius:20px;padding:5px 6px 5px 12px;font-size:12px;}
.od-area-chip button{border:0;background:none;cursor:pointer;color:${T.sub};display:flex;}
.od-area-chip button:hover{color:${T.coral};}
.od-headr{display:flex;align-items:flex-end;gap:12px;}
.od-rev{display:flex;flex-direction:column;gap:2px;}
.od-rev label{font-size:9.5px;text-transform:uppercase;letter-spacing:.06em;color:${T.faint};font-weight:700;}
.od-rev input{height:34px;border:1px solid ${T.line};border-radius:8px;padding:0 11px;font-size:13px;font-family:inherit;width:150px;color:${T.ink};}
.od-ghost{display:inline-flex;align-items:center;gap:6px;height:34px;border:1px solid ${T.line};background:#fff;border-radius:8px;padding:0 12px;font-size:12.5px;font-family:inherit;color:${T.ink};cursor:pointer;}
.od-ghost:hover{border-color:${T.primary};color:${T.primary};}
.od-cnt{background:${T.primary};color:#fff;border-radius:20px;padding:1px 7px;font-size:11px;}
.od-x{width:34px;height:34px;border:1px solid ${T.line};background:#fff;border-radius:8px;cursor:pointer;display:flex;align-items:center;justify-content:center;color:${T.sub};}
.od-x:hover{background:${T.bg};}

.od-completude{background:#fff;border-bottom:1px solid ${T.line};padding:10px 22px;display:flex;align-items:center;gap:16px;flex-wrap:wrap;flex-shrink:0;}
.od-donut{flex-shrink:0;}
.od-completude-txt{min-width:200px;}
.od-completude-t{font-size:13px;font-weight:700;color:${T.navy};}
.od-completude-legend{display:flex;gap:14px;flex-wrap:wrap;font-size:12px;color:${T.sub};margin-top:4px;align-items:center;}
.od-completude-legend span{display:inline-flex;align-items:center;gap:5px;} .od-completude-legend i{width:10px;height:10px;border-radius:3px;}
.od-completude-legend b{color:${T.ink};font-weight:800;}
.od-completude-den{color:${T.faint};font-size:11px;}
.od-controls-r{display:flex;align-items:center;gap:8px;margin-left:auto;flex-wrap:wrap;}
.od-search{display:flex;align-items:center;gap:8px;background:${T.bg};border:1px solid ${T.line};border-radius:9px;padding:0 12px;height:34px;min-width:180px;}
.od-search input{flex:1;border:0;outline:0;background:none;font-size:13px;font-family:inherit;color:${T.ink};}
.od-mini{height:34px;border:1px solid ${T.line};background:#fff;border-radius:8px;padding:0 11px;font-size:12px;font-family:inherit;color:${T.sub};cursor:pointer;}
.od-mini:hover{border-color:${T.primary};color:${T.primary};}
.od-zoom{display:flex;gap:4px;}
.od-zoom button{width:32px;height:32px;border:1px solid ${T.line};background:#fff;border-radius:7px;cursor:pointer;display:flex;align-items:center;justify-content:center;color:${T.sub};}
.od-zoom button:hover{border-color:${T.primary};color:${T.primary};}

.od-chartwrap{flex:1;overflow:auto;padding:30px 22px;min-height:0;}
.od-chart{display:inline-block;transform-origin:top center;transition:transform .15s;}
.oc-tree{padding:0;margin:0;} .oc-tree ul{display:flex;justify-content:center;padding-top:24px;margin:0;position:relative;list-style:none;}
.oc-tree li{list-style:none;position:relative;padding:24px 8px 0;text-align:center;}
.oc-tree li::before,.oc-tree li::after{content:'';position:absolute;top:0;right:50%;border-top:2px solid ${T.conn};width:50%;height:24px;}
.oc-tree li::after{right:auto;left:50%;border-left:2px solid ${T.conn};}
.oc-tree li:only-child::before,.oc-tree li:only-child::after{display:none;}
.oc-tree li:only-child{padding-top:24px;}
.oc-tree li:first-child::before,.oc-tree li:last-child::after{border:0 none;}
.oc-tree li:last-child::before{border-right:2px solid ${T.conn};border-radius:0 6px 0 0;}
.oc-tree li:first-child::after{border-radius:6px 0 0 0;}
.oc-tree ul ul::before{content:'';position:absolute;top:0;left:50%;border-left:2px solid ${T.conn};width:0;height:24px;}
.oc-tree>li.oc-raiz{padding-top:0;} .oc-tree>li.oc-raiz::before,.oc-tree>li.oc-raiz::after{display:none;}

.oc-node{position:relative;display:inline-block;vertical-align:top;border-radius:11px;padding:10px 15px;min-width:118px;max-width:184px;
  border:1.5px solid ${T.line};background:#fff;box-shadow:0 1px 3px rgba(13,49,111,.06);text-align:center;transition:.12s;}
.oc-node.dim{opacity:.3;} .oc-node.on{box-shadow:0 0 0 3px ${T.primary}33,0 3px 10px rgba(13,49,111,.18);}
.oc-sig{font-size:15px;font-weight:800;color:${T.navy};letter-spacing:-.01em;}
.oc-nome{font-size:10px;color:${T.sub};line-height:1.28;margin-top:3px;}

.oc-orgao{background:${T.navy};border-color:${T.navy};padding:13px 22px;min-width:140px;}
.oc-orgao .oc-sig{color:#fff;font-size:18px;} .oc-orgao .oc-nome{color:#C7D2E4;}

.oc-ramo{background:${T.primarySoft};border-color:#C6D6F0;}
.oc-ramo .oc-sig{color:${T.primaryDark};}
.oc-ramo-top{display:flex;align-items:center;justify-content:center;gap:6px;}
.oc-selflag{width:20px;height:20px;border-radius:6px;background:#fff;border:1px solid #C6D6F0;display:inline-flex;align-items:center;justify-content:center;cursor:pointer;}
.oc-selflag:hover{background:${T.greenSoft};}
.oc-cob{margin-top:6px;} .oc-cob i{display:block;height:4px;border-radius:20px;background:${T.green};margin-bottom:3px;}
.oc-cob{position:relative;height:auto;} .oc-cob::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;border-radius:20px;background:${T.line};z-index:0;}
.oc-cob i{position:relative;z-index:1;} .oc-cob span{font-size:9.5px;color:${T.primary};font-weight:700;}

.oc-depto{background:${T.greySoft};border-color:${T.line};padding:5px 11px;min-width:0;border-radius:20px;display:inline-flex;align-items:center;gap:7px;}
.oc-depto-sig{font-size:11.5px;font-weight:800;color:${T.sub};} .oc-depto-cob{font-size:10px;color:${T.faint};font-weight:700;}

.oc-unidade{border-width:1.5px;}
.oc-unidade.oc-feito{background:${T.greenSoft};border-color:#BEE0C4;}
.oc-unidade.oc-andamento{background:${T.amberSoft};border-color:#EBD3A8;}
.oc-unidade.oc-sem{background:#fff;border-color:${T.line};border-style:dashed;}
.oc-unidade.oc-sem .oc-sig{color:${T.grey};}
.oc-u-head{display:flex;align-items:center;justify-content:center;gap:6px;}
.oc-nec{font-size:11px;font-weight:800;border-radius:6px;padding:1px 6px;font-family:'JetBrains Mono',monospace;display:inline-flex;align-items:center;gap:2px;}
.oc-nec-sem{background:${T.greySoft};color:${T.grey};}

.oc-toggle{position:absolute;bottom:-11px;left:50%;transform:translateX(-50%);height:22px;min-width:22px;padding:0 6px;border-radius:20px;
  border:1.5px solid ${T.line};background:#fff;color:${T.sub};cursor:pointer;display:flex;align-items:center;justify-content:center;box-shadow:0 1px 3px rgba(0,0,0,.1);z-index:2;}
.oc-toggle:hover{border-color:${T.primary};color:${T.primary};} .oc-plus{font-size:11px;font-weight:800;}

.od-panelwrap{max-height:56%;overflow:auto;background:#fff;border-top:2px solid ${T.line};box-shadow:0 -4px 16px rgba(13,49,111,.06);flex-shrink:0;}
.od-painel{max-width:1080px;margin:0 auto;padding:20px 22px 30px;}
.od-p-head{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;margin-bottom:16px;}
.od-p-crumb{font-size:11px;color:${T.faint};font-weight:600;}
.od-p-title{font-size:20px;font-weight:800;color:${T.navy};letter-spacing:-.01em;margin:4px 0 0;}
.od-p-headr{display:flex;align-items:center;gap:10px;}
.od-p-status{font-size:12px;font-weight:700;border-radius:20px;padding:5px 12px;display:inline-flex;align-items:center;gap:6px;white-space:nowrap;}
.od-p-close{width:32px;height:32px;border:1px solid ${T.line};background:#fff;border-radius:8px;cursor:pointer;display:flex;align-items:center;justify-content:center;color:${T.sub};}
.od-p-close:hover{background:${T.bg};}
.od-p-sem{display:flex;align-items:center;gap:9px;background:${T.greySoft};border:1px solid ${T.line};border-radius:11px;padding:14px 16px;font-size:13.5px;color:${T.sub};}
.od-p-cards{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:22px;}
@media(max-width:820px){.od-p-cards{grid-template-columns:repeat(2,1fr);}}
.od-mc{background:#fff;border:1px solid ${T.line};border-radius:12px;padding:13px 14px;text-align:left;}
.od-mc-l{font-size:10.5px;text-transform:uppercase;letter-spacing:.04em;color:${T.faint};font-weight:700;display:flex;align-items:center;gap:5px;}
.od-mc-v{font-size:20px;font-weight:800;color:${T.navy};margin:6px 0 2px;}
.od-mc-arrow{font-size:12px;font-weight:600;color:${T.sub};} .od-mc-s{font-size:11px;color:${T.sub};line-height:1.35;}
.od-bi{cursor:pointer;font-family:inherit;background:${T.primarySoft};border-color:transparent;}
.od-bi:hover{background:${T.primary};} .od-bi:hover *{color:#fff!important;}
.od-bi-cta{font-size:14px;font-weight:700;color:${T.primaryDark};margin:8px 0 3px;display:flex;align-items:center;gap:6px;}
.od-p-andamento{display:flex;align-items:center;gap:9px;background:${T.amberSoft};border:1px solid #EBD3A8;border-radius:11px;padding:12px 14px;font-size:13px;color:${T.amber};margin-bottom:20px;}
.od-tbl-head{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:10px;}
.od-tbl-t{font-size:14px;font-weight:700;color:${T.navy};display:flex;align-items:center;gap:8px;}
.od-tbl-chips{display:inline-flex;gap:6px;margin-left:4px;}
.od-tbl-chips span{font-size:11px;font-weight:600;background:${T.bg};color:${T.sub};border-radius:20px;padding:2px 9px;}
.od-tbl-chips .warn{background:${T.amberSoft};color:${T.amber};} .od-tbl-chips .mut{background:${T.coralSoft};color:${T.coral};}
.od-tbl-search{display:flex;align-items:center;gap:7px;background:#fff;border:1px solid ${T.line};border-radius:8px;padding:0 11px;height:36px;}
.od-tbl-search input{border:0;outline:0;font-size:13px;font-family:inherit;background:none;width:180px;color:${T.ink};}
.od-tbl-wrap{border:1px solid ${T.line};border-radius:12px;overflow:hidden;background:#fff;}
.od-tbl{width:100%;border-collapse:collapse;font-size:13px;}
.od-tbl thead th{position:sticky;top:0;background:${T.bg};text-align:left;font-size:10.5px;text-transform:uppercase;letter-spacing:.04em;color:${T.faint};font-weight:700;padding:9px 12px;border-bottom:1px solid ${T.line};}
.od-tbl td{padding:10px 12px;border-bottom:1px solid ${T.line};vertical-align:top;}
.od-tbl tr:last-child td{border-bottom:0;} .od-tbl tbody tr:hover{background:#FAFBFC;} .od-tbl tr.fora{background:#FDFBF9;}
.col-ent{width:34%;} .col-nat{width:11%;} .col-mac{width:18%;} .col-cat{width:17%;} .col-serv{width:20%;}
.od-ent-nm{font-weight:600;line-height:1.4;color:${T.ink};}
.od-ent-cd{font-size:10.5px;color:${T.faint};font-family:'JetBrains Mono',monospace;margin-top:2px;display:flex;align-items:center;gap:6px;}
.od-fora{background:${T.coralSoft};color:${T.coral};border-radius:5px;padding:0 6px;font-family:inherit;font-weight:700;}
.od-natpill{font-size:11px;font-weight:700;border-radius:6px;padding:2px 8px;white-space:nowrap;}
.col-mac,.col-cat{font-size:12px;color:${T.sub};line-height:1.35;}
.od-falta{color:${T.faint};display:inline-flex;align-items:center;gap:4px;font-size:11.5px;}
.od-serv-new{color:${T.green};font-weight:600;}
.od-serv-tag{font-size:10px;background:${T.greenSoft};color:${T.green};border-radius:20px;padding:1px 7px;margin-left:6px;font-weight:700;display:inline-flex;align-items:center;gap:3px;}
.od-add{display:flex;gap:5px;}
.od-add input{flex:1;min-width:90px;height:32px;border:1.5px solid ${T.amberSoft};background:#FFFDF8;border-radius:7px;padding:0 9px;font-size:12.5px;font-family:inherit;color:${T.ink};}
.od-add input:focus{border-color:${T.primary};outline:0;background:#fff;}
.od-add button{width:32px;height:32px;border:0;background:${T.primary};color:#fff;border-radius:7px;cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0;}
.od-add button:hover{background:${T.primaryDark};}
.od-tbl-empty{text-align:center;color:${T.faint};padding:26px;font-size:13px;}
.od-toast{position:fixed;bottom:22px;left:50%;transform:translateX(-50%);background:${T.ink};color:#fff;padding:11px 18px;border-radius:9px;font-size:13px;z-index:70;box-shadow:0 6px 24px rgba(0,0,0,.22);max-width:90vw;text-align:center;}


/* ===== ajustes v3 ===== */
/* destaque da entrega — sem divis\xF3ria, tom suave */
.od-th-ent{color:${T.primaryDark}!important;}
.od-ent-cell{border-left:0!important;background:${T.primarySoft}55!important;}
.od-ent-cell .od-ent-nm{font-weight:700;color:${T.navy};}

/* bot\xF5es voltar (Organograma): X \xEDcone + bot\xE3o destacado */
.od-x{width:34px!important;height:34px;padding:0!important;gap:0;color:${T.sub};display:inline-flex;align-items:center;justify-content:center;}
.od-back{display:inline-flex;align-items:center;gap:6px;height:34px;border:1px solid ${T.primary};background:${T.primary};color:#fff;border-radius:8px;padding:0 13px;font-size:12.5px;font-weight:700;font-family:inherit;cursor:pointer;}
.od-back:hover{background:${T.primaryDark};border-color:${T.primaryDark};}

/* estado: nenhuma organiza\xE7\xE3o escolhida */
.od-choose{flex:1;display:flex;align-items:center;justify-content:center;padding:40px;}
.od-choose-card{max-width:430px;text-align:center;background:#fff;border:1px solid ${T.line};border-radius:16px;padding:36px 30px;box-shadow:0 10px 30px rgba(13,49,111,.05);}
.od-choose-ic{width:64px;height:64px;border-radius:16px;background:${T.primarySoft};color:${T.primary};display:inline-flex;align-items:center;justify-content:center;margin-bottom:14px;}
.od-choose-card h3{margin:0 0 8px;font-size:18px;color:${T.navy};}
.od-choose-card p{margin:0;font-size:13.5px;color:${T.sub};line-height:1.55;}

/* organograma horizontal (com pan) + staff lateral */
.od-chartwrap{cursor:grab;}
.od-chartwrap.grabbing{cursor:grabbing;}
.oc-row{display:flex;align-items:center;gap:8px;justify-content:center;flex-wrap:nowrap;}
.oc-staff{display:flex;flex-direction:column;gap:4px;align-items:flex-start;position:relative;padding-left:8px;border-left:1px dashed ${T.blue};}
.oc-staff::before{content:'staff';position:absolute;left:6px;top:-11px;font-size:8px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:${T.faint};}
.oc-staffchip{display:flex;flex-direction:column;align-items:flex-start;gap:0;background:#fff;border:1px dashed ${T.blue};border-radius:8px;padding:3px 8px;font-family:inherit;max-width:160px;cursor:pointer;}
.oc-staffchip:hover{background:${T.blueSoft};}
.oc-staffchip.on{background:${T.blueSoft};border-style:solid;}
.oc-staff-sig{font-size:10.5px;font-weight:800;color:${T.blue};}
.oc-staff-nome{display:none;}

/* painel redimension\xE1vel / sobrepon\xEDvel */
.od-panelwrap{max-height:none!important;overflow:auto;background:#fff;border-top:2px solid ${T.line};box-shadow:0 -6px 20px rgba(13,49,111,.08);flex-shrink:0;position:relative;}
.od-panelwrap.max{position:absolute;left:0;right:0;top:58px;bottom:0;height:auto!important;z-index:8;}
.od-p-grip{display:flex;align-items:center;justify-content:center;height:16px;cursor:ns-resize;background:${T.bg};border-bottom:1px solid ${T.line};position:sticky;top:0;z-index:2;}
.od-p-grip-bar{width:46px;height:4px;border-radius:3px;background:${T.grey};}
.od-p-grip:hover .od-p-grip-bar{background:${T.primary};}
.od-p-max{width:30px;height:30px;border:1px solid ${T.line};background:#fff;border-radius:8px;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;color:${T.sub};}
.od-p-max:hover{border-color:${T.primary};color:${T.primary};}

/* n\xF3s menores + s\xF3 sigla + \xE1rvore menos alta */
.oc-node{padding:7px 10px!important;min-width:auto!important;width:auto!important;gap:2px;}
.oc-orgao,.oc-ramo,.oc-unidade,.oc-depto{min-width:auto!important;width:auto!important;max-width:230px;}
.oc-sig{font-size:12px;}
.oc-row{gap:9px;}
/* Natureza discreta (ponto) + planilha leve */
.col-nat{width:34px!important;text-align:center;}
.od-natdot{display:inline-block;width:11px;height:11px;border-radius:50%;vertical-align:middle;}
.od-tbl td,.od-tbl th{border-right:1px solid ${T.line};}
.od-tbl td:last-child,.od-tbl th:last-child{border-right:0;}
.od-tbl th{background:${T.bg};}


/* montar no cat\xE1logo (unidade sem DFT) */
.od-p-sem{flex-direction:column;align-items:flex-start!important;gap:12px!important;}
.od-p-sem-tx{display:flex;align-items:flex-start;gap:9px;}
.od-p-montar{display:inline-flex;align-items:center;gap:8px;border:0;background:${T.primary};color:#fff;border-radius:9px;padding:9px 15px;font-size:13px;font-weight:700;font-family:inherit;cursor:pointer;}
.od-p-montar:hover{background:${T.primaryDark};}

/* ===== ajustes v4: lentes \xB7 tooltip \xB7 conectores SVG \xB7 transi\xE7\xF5es ===== */

/* barra de lentes */
.od-lens{background:#fff;border-bottom:1px solid ${T.line};padding:7px 22px;display:flex;align-items:center;gap:12px;flex-wrap:wrap;flex-shrink:0;}
.od-lens-lb{display:inline-flex;align-items:center;gap:6px;font-size:10.5px;font-weight:800;text-transform:uppercase;letter-spacing:.06em;color:${T.faint};}
.od-lens-seg{display:inline-flex;background:${T.bg};border:1px solid ${T.line};border-radius:9px;padding:3px;gap:2px;flex-wrap:wrap;}
.od-lens-seg button{border:0;background:none;border-radius:7px;padding:5px 11px;font-size:12px;font-weight:600;font-family:inherit;color:${T.sub};cursor:pointer;transition:color .15s,background .15s,box-shadow .15s;}
.od-lens-seg button:hover{color:${T.primary};}
.od-lens-seg button.on{background:#fff;color:${T.primaryDark};font-weight:700;box-shadow:0 1px 3px rgba(13,49,111,.14);}
.od-lens-leg{display:flex;align-items:center;gap:11px;flex-wrap:wrap;margin-left:auto;font-size:11px;color:${T.sub};}
.od-lens-leg span{display:inline-flex;align-items:center;gap:5px;white-space:nowrap;}
.od-lens-leg i{width:12px;height:12px;border-radius:4px;border:1.5px solid ${T.conn};display:inline-block;flex-shrink:0;}
.od-lens-leg i.dsh{border-style:dashed;background:#fff;}
/* a lente troca cores por estilo inline; suaviza a transi\xE7\xE3o */
.oc-node,.oc-staffchip{transition:background .3s ease,border-color .3s ease,box-shadow .16s ease,transform .16s ease,opacity .2s ease;}

/* tooltip rico */
.oc-tip{position:fixed;transform:translate(-50%,-100%);z-index:80;background:${T.ink};color:#fff;border-radius:10px;padding:10px 13px;min-width:215px;max-width:300px;box-shadow:0 10px 30px rgba(0,0,0,.28);pointer-events:none;text-align:left;animation:ocTipIn .14s ease;}
@keyframes ocTipIn{from{opacity:0;transform:translate(-50%,calc(-100% + 5px));}to{opacity:1;transform:translate(-50%,-100%);}}
.oc-tip::after{content:'';position:absolute;left:50%;bottom:-4px;transform:translateX(-50%) rotate(45deg);width:9px;height:9px;background:${T.ink};border-radius:2px;}
.oc-tip-h{display:flex;align-items:center;gap:8px;flex-wrap:wrap;}
.oc-tip-sig{font-size:14px;font-weight:800;letter-spacing:-.01em;}
.oc-tip-st{font-size:10px;font-weight:700;border-radius:20px;padding:2px 8px;white-space:nowrap;}
.oc-tip-nome{font-size:11.5px;color:#C9D2DE;margin-top:2px;line-height:1.35;}
.oc-tip-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px 12px;margin-top:9px;padding-top:9px;border-top:1px solid rgba(255,255,255,.14);}
.oc-tip-grid div{display:flex;flex-direction:column;min-width:0;}
.oc-tip-grid b{font-size:13px;font-weight:800;line-height:1.25;}
.oc-tip-grid span{font-size:9px;color:#9AA6B5;text-transform:uppercase;letter-spacing:.06em;margin-top:1px;}
.oc-tip-row{display:flex;align-items:center;gap:6px;font-size:11.5px;color:#DDE3EA;margin-top:8px;line-height:1.3;}
.oc-tip-row i{width:8px;height:8px;border-radius:2px;flex-shrink:0;}
.oc-tip-hint{margin-top:8px;font-size:10px;color:#8D99A8;font-style:italic;}

/* conectores SVG substituem os pseudo-elementos CSS */
.oc-tree li::before,.oc-tree li::after,.oc-tree ul ul::before{display:none!important;}
.od-chart{position:relative;}
.oc-tree{position:relative;z-index:1;}
.oc-svg{position:absolute;top:0;left:0;z-index:0;pointer-events:none;}
.oc-tree li,.oc-tree li:only-child{padding-top:30px;}
.oc-tree ul{padding-top:30px;}
.oc-tree>li.oc-raiz{padding-top:0;}
/* respiro inferior p/ o bal\xE3ozinho +N n\xE3o ser cortado quando o n\xF3 est\xE1 recolhido */
.oc-row{padding-bottom:13px;}
.oc-svg{overflow:visible;}

/* expand/collapse animado */
.oc-kids{display:grid;grid-template-rows:0fr;opacity:0;transition:grid-template-rows .3s cubic-bezier(.25,.7,.3,1),opacity .22s ease;}
.oc-kids.open{grid-template-rows:1fr;opacity:1;}
.oc-kids>ul{overflow:hidden;min-height:0;}

/* microintera\xE7\xF5es nos n\xF3s */
.oc-node:hover{transform:translateY(-1px);box-shadow:0 4px 12px rgba(13,49,111,.15);}
.oc-node.dim:hover{transform:none;box-shadow:0 1px 3px rgba(13,49,111,.06);}
.oc-toggle{transition:border-color .15s,color .15s,transform .15s;}
.oc-toggle:hover{transform:translateX(-50%) scale(1.12);}

/* ===== v5: estetica do organograma — cartoes com sigla+nome, canvas pontilhado, conectores ortogonais ===== */
.od-chartwrap{background-image:radial-gradient(${T.conn}66 1.1px,transparent 1.1px);background-size:22px 22px;background-position:-4px -4px;}
.oc-node{padding:9px 13px 8px!important;min-width:112px!important;max-width:170px;border-radius:12px;
  box-shadow:0 1px 2px rgba(13,49,111,.05),0 5px 16px -8px rgba(13,49,111,.12);}
.oc-sig{font-size:12.5px;font-weight:800;letter-spacing:.01em;}
.oc-nome{display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;
  font-size:9px;line-height:1.32;margin:3px auto 0;color:${T.faint};font-weight:600;max-width:150px;}
.oc-orgao{padding:12px 24px 11px!important;min-width:180px!important;max-width:250px;border-radius:14px;
  background:linear-gradient(160deg,${T.navy} 0%,#144A9C 100%);border-color:${T.navy};
  box-shadow:0 8px 22px -8px rgba(12,50,111,.5);}
.oc-orgao .oc-sig{font-size:16px;}
.oc-orgao .oc-nome{color:#B9C7DE;font-size:9.5px;max-width:210px;}
.oc-ramo .oc-nome{color:#5E77A3;}
.oc-unidade.oc-sem .oc-nome{color:${T.grey};opacity:.85;}
.oc-cob{margin-top:7px;}
.oc-svg path{stroke:${T.conn};stroke-width:1.4px;}
.oc-toggle{border-color:${T.conn};font-size:10px;box-shadow:0 1px 4px rgba(13,49,111,.14);}
.oc-toggle:hover{border-color:${T.primary};background:${T.primarySoft};}
.oc-staffchip{border-radius:9px;padding:4px 9px;}
.oc-staff-nome{display:-webkit-box;-webkit-line-clamp:1;-webkit-box-orient:vertical;overflow:hidden;
  font-size:8.5px;color:${T.faint};font-weight:600;max-width:150px;text-align:left;}
.oc-tree li,.oc-tree li:only-child{padding-top:34px;}
.oc-tree ul{padding-top:34px;}
.oc-row{gap:11px;}
.oc-node.on{box-shadow:0 0 0 3px ${T.primary}2E,0 6px 18px -6px rgba(13,49,111,.3);}

/* ===== v6: anti-sobreposicao + busca global ===== */
.od-chart{width:max-content;}
.oc-tree ul{min-width:max-content;}
.oc-tree li{flex-shrink:0;}
.oc-row{flex-shrink:0;}
.oc-kids{grid-template-columns:minmax(max-content,1fr);}
.oc-kids>ul{min-width:max-content;}

.od-search{position:relative;}
.od-gdrop{position:absolute;top:calc(100% + 6px);left:0;min-width:320px;max-width:440px;max-height:350px;overflow:auto;
  background:#fff;border:1px solid ${T.line};border-radius:12px;box-shadow:0 14px 36px rgba(13,49,111,.18);z-index:45;padding:5px;text-align:left;}
.od-gdrop-h{font-size:9.5px;font-weight:800;text-transform:uppercase;letter-spacing:.07em;color:${T.faint};padding:7px 9px 4px;}
.od-gdrop-it{display:flex;align-items:baseline;gap:8px;width:100%;text-align:left;border:0;background:none;border-radius:8px;padding:7px 9px;cursor:pointer;font-family:inherit;}
.od-gdrop-it:hover{background:${T.primarySoft};}
.od-gdrop-sig{font-size:12px;font-weight:800;color:${T.navy};white-space:nowrap;}
.od-gdrop-nm{font-size:11px;color:${T.sub};overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;min-width:0;}
.od-gdrop-org{font-size:10px;font-weight:700;color:${T.primaryDark};background:${T.primarySoft};border-radius:20px;padding:1px 7px;white-space:nowrap;}
.od-gdrop-vazio{padding:10px;font-size:12px;color:${T.faint};}
.od-gdrop-mais{padding:7px 9px;font-size:10.5px;color:${T.faint};border-top:1px solid ${T.line};margin-top:3px;}
.od-gsearch-hero{margin-top:18px;height:42px;background:#fff;border-color:${T.conn};}
.od-gsearch-hero input{font-size:13.5px;}
.od-gsearch-hero .od-gdrop{left:0;right:0;min-width:0;max-width:none;}
.od-choose-card{overflow:visible;position:relative;}

/* ===== v7: aglomerado de colegiados (comites/comissoes/conselhos) ===== */
.oc-cluster{background:repeating-linear-gradient(135deg,${T.bg},${T.bg} 7px,#EEF1F6 7px,#EEF1F6 14px)!important;
  border:1.5px dashed ${T.conn}!important;border-radius:12px;padding:8px 14px 9px!important;min-width:120px!important;max-width:180px;
  box-shadow:none!important;cursor:pointer;}
.oc-cluster:hover{border-color:${T.primary}!important;background:${T.primarySoft}!important;transform:translateY(-1px);}
.oc-cluster-h{display:flex;align-items:center;justify-content:center;gap:6px;color:${T.sub};}
.oc-cluster-t{font-size:11.5px;font-weight:800;letter-spacing:.02em;color:${T.sub};}
.oc-cluster-sub{font-size:9px;color:${T.faint};font-weight:600;margin-top:2px;}
.oc-cluster .oc-toggle{border-style:dashed;}

/* ===== v8: unidade englobada no DFT da coordenacao-geral ===== */
.oc-unidade.oc-englobado{background:${T.greenSoft};border-color:#BEE0C4;border-style:dashed;}
.oc-unidade.oc-englobado .oc-sig{color:${T.navy};}
.oc-nec-eng{background:${T.greenSoft}!important;color:${T.green}!important;display:inline-flex;align-items:center;gap:2px;}

/* ===== v9: CG clicavel + lista englobada no painel ===== */
.oc-depto-dft{border-color:#BEE0C4;background:linear-gradient(180deg,#fff,#F4FBF5);cursor:pointer;}
.oc-depto-dft:hover{transform:translateY(-2px);box-shadow:0 4px 8px rgba(13,49,111,.08),0 10px 24px -8px rgba(22,136,33,.25);}
.od-p-engbanner{display:flex;align-items:center;gap:9px;background:${T.greenSoft};border-bottom:1px solid #D3EBD7;
  padding:8px 18px;font-size:11.5px;color:#0F5C1A;}
.od-p-engbanner b{font-weight:800;}
.od-p-engbanner span{flex:1;}
.od-p-engbtn{border:1px dashed #9CCBA4;background:none;border-radius:8px;padding:4px 10px;font-size:10.5px;font-weight:700;
  color:${T.green};cursor:pointer;font-family:inherit;white-space:nowrap;}
.od-p-engbtn:hover{background:#fff;}
.od-p-relatorio{display:inline-flex;align-items:center;gap:6px;border:1px solid ${T.primary};background:${T.primary};color:#fff;
  border-radius:9px;padding:6px 12px;font-size:11.5px;font-weight:700;cursor:pointer;font-family:inherit;margin-right:8px;}
.od-p-relatorio:hover{background:${T.primaryDark};border-color:${T.primaryDark};}
.od-tbl-loading-spin{display:inline-block;width:11px;height:11px;border-radius:50%;
  border:2px solid ${T.line};border-top-color:${T.primary};animation:od-spin .7s linear infinite;vertical-align:-1.5px;}
@keyframes od-spin{to{transform:rotate(360deg);}}

/* ===== v12: painel com abas Entregas / Relatorio completo ===== */
.od-p-head2{display:flex;align-items:center;gap:12px;margin-bottom:2px;}
.od-p-idline{flex:1;min-width:0;display:flex;align-items:baseline;gap:10px;}
.od-p-crumb2{font-size:10px;letter-spacing:.04em;text-transform:uppercase;color:${T.faint};font-weight:700;white-space:nowrap;}
.od-p-title2{font-size:16px;font-weight:800;color:${T.navy};overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.od-p-tabbar{display:flex;align-items:center;gap:14px;padding-bottom:10px;margin-bottom:12px;border-bottom:1px solid ${T.line};flex-wrap:wrap;}
.od-p-tabs{display:inline-flex;gap:3px;background:${T.bg};border-radius:10px;padding:3px;flex-shrink:0;}
.od-p-tab{border:0;background:transparent;border-radius:8px;padding:7px 14px;font-size:12px;font-weight:700;cursor:pointer;
  font-family:inherit;color:${T.sub};display:flex;align-items:center;gap:6px;white-space:nowrap;}
.od-p-tab.on{background:#fff;color:${T.primaryDark};box-shadow:0 1px 3px rgba(13,49,111,.12);}
.od-p-tabn{font-family:monospace;background:${T.bg};border-radius:6px;padding:1px 6px;font-size:10.5px;}
.od-p-tab.on .od-p-tabn{background:${T.bg};}
/* bloco recolhido das entregas, abaixo do relatório */
.od-p-expand{display:flex;align-items:center;gap:8px;width:100%;margin-top:14px;background:#fff;
  border:1px solid ${T.line};border-radius:10px;padding:11px 13px;font-family:inherit;font-size:12.5px;
  font-weight:700;color:${T.navy};cursor:pointer;transition:all .14s;text-align:left;}
.od-p-expand:hover{border-color:${T.primary};color:${T.primaryDark};}
.od-p-expand.on{border-color:${T.primary};background:${T.primarySoft};border-bottom-left-radius:0;border-bottom-right-radius:0;}
.od-p-expand-chip{font-size:10px;font-weight:700;border-radius:20px;padding:2px 8px;}
.od-p-expand-chip.warn{background:#FBF3DC;color:#9A6A00;}
.od-p-expand-chip.mut{background:${T.bg};color:${T.sub};}
.od-p-expand-chip:first-of-type{margin-left:auto;}
.od-p-inds{display:flex;align-items:center;gap:14px;font-size:11px;color:${T.sub};font-weight:600;overflow:hidden;flex-wrap:wrap;min-width:0;padding:12px 2px 4px;}
.od-p-inds span{white-space:nowrap;}
.od-p-inds b{color:${T.navy};}
.od-p-bilink{border:0;background:none;color:${T.primary};font-weight:700;font-size:11px;cursor:pointer;font-family:inherit;
  white-space:nowrap;display:inline-flex;align-items:center;gap:4px;padding:0;}
.od-p-obsbtn{flex-shrink:0;border:0;background:none;color:${T.faint};cursor:pointer;font-size:11px;font-weight:700;padding:0;margin-left:auto;}
.od-p-obsbtn.on{color:${T.primary};}
.od-p-protonote{font-size:11px;color:${T.faint};font-weight:600;}
.od-p-obsbox{display:flex;gap:8px;align-items:flex-start;background:#FBF9F1;border:1px solid #EDE3C4;border-radius:9px;
  padding:9px 12px;margin-bottom:12px;font-size:11px;color:#7A6A1F;font-weight:600;line-height:1.5;}
.od-p-obsbox b{color:#5C4E12;}
.od-p-reportwrap{flex:1;min-height:0;display:flex;border:1px solid ${T.line};border-radius:12px;overflow:hidden;background:#FCFCFD;}
.od-p-reportframe{width:100%;height:100%;border:0;display:block;min-height:420px;}
.od-painel{display:flex;flex-direction:column;height:100%;}
.od-p-cards,.od-p-sem,.od-p-andamento{flex-shrink:0;}
.od-tbl-wrap{flex:1;min-height:0;overflow:auto;}

/* ===== v10: painel mais leve + botao focar ===== */
.od-p-head{padding-top:10px!important;padding-bottom:10px!important;}
.od-p-crumb{font-size:10px!important;opacity:.75;}
.od-p-title{font-size:15px!important;line-height:1.25!important;margin-top:1px!important;}
.od-p-status{font-size:10px!important;padding:2px 9px!important;}
.od-p-engbanner{padding:6px 18px!important;font-size:11px!important;gap:7px!important;}
.od-mini-on{background:${T.primary}!important;border-color:${T.primary}!important;color:#fff!important;}

/* ===== v11: modo edicao + so com DFT ===== */
.od-mini-edit{background:${T.amber}!important;border-color:${T.amber}!important;color:#fff!important;}
.oc-editable{outline:1.5px dashed transparent;transition:outline-color .12s;}
.oc-editable:hover{outline-color:${T.amber};}
.oc-node.oc-dragover{outline:2px solid ${T.primary}!important;outline-offset:2px;background:${T.primarySoft}!important;}
.oc-del{position:absolute;top:-9px;right:-9px;width:19px;height:19px;border-radius:50%;border:1px solid #E3B9AC;
  background:#fff;color:#9E3B1F;font-size:11px;font-weight:800;line-height:16px;cursor:pointer;padding:0;z-index:6;
  box-shadow:0 1px 4px rgba(13,49,111,.18);}
.oc-del:hover{background:#9E3B1F;color:#fff;}
.od-editbanner{display:flex;align-items:center;gap:9px;background:${T.amberSoft};border-bottom:1px solid #EAD9BD;
  color:${T.amber};font-size:11.5px;font-weight:600;padding:6px 16px;}

`;

function OrganogramaDimensionamento({ selPrincipal, onClose, onMontarNoCatalogo }) {
  const [carregando, setCarregando] = useState(true);
  const [erro, setErro] = useState(null);
  const [dados, setDados] = useState({ orgaos: [], entregas: {} });

  useEffect(() => {
    let cancelado = false;
    setCarregando(true); setErro(null);
    carregarOrganograma()
      .then((d) => { if (!cancelado) { setDados(d); window.__entregasPorSigla = d.entregas; setCarregando(false); } })
      .catch((e) => { if (!cancelado) { console.error("Erro ao carregar organograma:", e); setErro(e.message || "Erro desconhecido"); setCarregando(false); } });
    return () => { cancelado = true; };
  }, []);

  const [orgIdx, setOrgIdx] = useState(null);
  const orgAtual = orgIdx != null ? dados.orgaos[orgIdx] : null;

  const [busca, setBusca] = useState("");
  const [selecao, setSelecao] = useState(null);
  const [colapsados, setColapsados] = useState(new Set());
  const [zoom, setZoom] = useState(1);
  const [foco, setFoco] = useState(true);                 // "Focar" — liga/desliga o auto-centralizar
  const [soDft, setSoDft] = useState(false);
  const [editando, setEditando] = useState(false);
  const [ops, setOps] = useState({ removed: [], moved: {} }); // visão personalizada (por navegador)

  const [revisor, setRevisor] = useState("");
  const [inseridos, setInseridos] = useState([]);
  const [overridesPorUnidade, setOverridesPorUnidade] = useState({});
  const [servicosPorUnidade, setServicosPorUnidade] = useState({});
  const [anexoAberto, setAnexoAberto] = useState(false);
  const [maxPanel, setMaxPanel] = useState(false);
  const [panelH, setPanelH] = useState(360);
  const [toast, setToast] = useState(null);
  const [lente, setLente] = useState("status");
  const [tip, setTip] = useState(null);

  const chartRef = useRef(null);
  const metricas = useMemo(() => { const acc = {}; orgAtual && agregarMetricas(orgAtual, acc); return acc; }, [orgIdx]);

  function mostrarToast(msg) { setToast(msg); setTimeout(() => setToast(null), 2200); }
  const entregasSelecao = selecao && servicosPorUnidade[selecao.uid] ? servicosPorUnidade[selecao.uid].entregas : [];

  /* ao trocar de órgão: reabre tudo, fecha painel, carrega visão personalizada salva */
  useEffect(() => {
    if (orgAtual) setColapsados(abrirPorPadrao(orgAtual, new Set()));
    setSelecao(null); setMaxPanel(false); setTip(null); setEditando(false); setSoDft(false);
    let salvo = { removed: [], moved: {} };
    try { if (orgAtual) { const s = window.localStorage.getItem("odops:" + orgAtual.uid); if (s) salvo = JSON.parse(s); } } catch (_e) {}
    setOps(salvo);
  }, [orgIdx]);

  /* persiste a visão personalizada por órgão */
  useEffect(() => {
    try { orgAtual && window.localStorage.setItem("odops:" + orgAtual.uid, JSON.stringify(ops)); } catch (_e) {}
  }, [ops, orgIdx]);

  /* ============================================================
     "FOCAR" — centraliza a unidade clicada e ajusta o zoom
     -------------------------------------------------------------
     CORREÇÕES DESTA SESSÃO:
     1) a busca do elemento no DOM não usa mais um seletor CSS
        montado por concatenação de texto (`querySelector('[data-ocuid="'+uid+'"]')`),
        que lançava exceção silenciosa para uids com aspas ou
        certos caracteres especiais — agora comparamos o atributo
        diretamente, o que nunca falha por escaping.
     2) se o nó ainda não estiver no DOM no instante do clique
        (ex.: layout ainda assentando), tenta de novo por até
        6 quadros de animação em vez de desistir na primeira vez.
     3) o alvo de rolagem é sempre limitado (clamp) aos limites
        reais de scrollWidth/scrollHeight do contêiner, então
        nunca tenta rolar para um ponto inexistente.
     4) loga no console (`[organograma] Focar: ...`) o que foi
        calculado — antes a falha era 100% muda.
  ============================================================ */
  useEffect(() => {
    if (!foco || !selecao) return;
    const wrap = chartRef.current?.parentElement; // .od-chartwrap (contêiner com scroll)
    if (!wrap) return;

    function encontrarNo() {
      const nos = wrap.querySelectorAll("[data-ocuid]");
      for (let i = 0; i < nos.length; i++) if (nos[i].getAttribute("data-ocuid") === selecao.uid) return nos[i];
      return null;
    }

    let tentativas = 0;
    function centralizar() {
      const el = encontrarNo();
      if (!el) {
        if (tentativas < 6) { tentativas++; requestAnimationFrame(centralizar); return; }
        console.warn("[organograma] Focar: unidade não encontrada no DOM após várias tentativas — uid:", selecao.uid);
        return;
      }
      const er = el.getBoundingClientRect(), wr = wrap.getBoundingClientRect();
      let alvoX = wrap.scrollLeft + (er.left - wr.left) - (wr.width / 2 - er.width / 2);
      let alvoY = wrap.scrollTop + (er.top - wr.top) - (wr.height * 0.34 - er.height / 2);
      const maxX = Math.max(0, wrap.scrollWidth - wrap.clientWidth);
      const maxY = Math.max(0, wrap.scrollHeight - wrap.clientHeight);
      alvoX = Math.max(0, Math.min(alvoX, maxX));
      alvoY = Math.max(0, Math.min(alvoY, maxY));
      try { wrap.scrollTo({ left: alvoX, top: alvoY, behavior: "smooth" }); }
      catch (_e) { wrap.scrollLeft = alvoX; wrap.scrollTop = alvoY; }
      console.info("[organograma] Focar: uid=", selecao.uid,
        "alvo(x,y)=", Math.round(alvoX), Math.round(alvoY),
        "limites(maxX,maxY)=", Math.round(maxX), Math.round(maxY),
        "scrollWidth/Height=", wrap.scrollWidth, wrap.scrollHeight);
    }

    if (zoom > 1) { setZoom(0.85); requestAnimationFrame(() => requestAnimationFrame(centralizar)); }
    else requestAnimationFrame(centralizar);
  }, [selecao, foco]);

  function onTipEnter(n, e) {
    const r = e.currentTarget.getBoundingClientRect();
    setTip({ n, m: metricas[n.uid], x: r.left + r.width / 2, y: r.top });
  }
  const onTipLeave = () => setTip(null);

  /* pan por arrastar (clique e arraste no fundo do canvas) */
  const arrasteRef = useRef({ on: false, x: 0, y: 0, sl: 0, st: 0 });
  function iniciarPan(e) {
    if (e.target.closest?.(".oc-node,.oc-staffchip,.oc-toggle,button")) return;
    const wrap = chartRef.current?.parentElement;
    if (!wrap) return;
    setTip(null);
    arrasteRef.current = { on: true, x: e.clientX, y: e.clientY, sl: wrap.scrollLeft, st: wrap.scrollTop };
    wrap.classList.add("grabbing");
  }
  function moverPan(e) {
    const a = arrasteRef.current;
    if (!a.on) return;
    const wrap = chartRef.current.parentElement;
    wrap.scrollLeft = a.sl - (e.clientX - a.x);
    wrap.scrollTop = a.st - (e.clientY - a.y);
  }
  function soltarPan() {
    arrasteRef.current.on = false;
    chartRef.current?.parentElement?.classList.remove("grabbing");
  }

  function toggleColapso(uid) {
    setColapsados((s) => { const n = new Set(s); n.has(uid) ? n.delete(uid) : n.add(uid); return n; });
  }
  function recolherTudo() { setColapsados(new Set()); }
  function expandirTudo() { orgAtual && setColapsados(abrirPorPadrao(orgAtual, new Set())); }

  /* modo edição: mover / ocultar / restaurar visão personalizada */
  function moverUnidade(id, destino) { if (id === destino) return; setOps((p) => ({ removed: p.removed || [], moved: { ...(p.moved || {}), [id]: destino } })); }
  function ocultarUnidade(uid) { setOps((p) => ({ removed: [...new Set([...(p.removed || []), uid])], moved: p.moved || {} })); }
  function restaurarVisao() { setOps({ removed: [], moved: {} }); }

  function registrarServico(unidade, entrega, valorDigitado) {
    const valor = (valorDigitado || "").trim();
    if (!valor) return;
    setOverridesPorUnidade((prev) => ({ ...prev, [entrega.cd]: valor }));
    const agora = new Date().toISOString().slice(0, 19).replace("T", " ");
    setInseridos((lista) => [...lista, {
      Código: entrega.cd, Entrega: entrega.nm, Serviço: valor, "Origem do Serviço": "manual",
      Natureza: entrega.nat, MacroProcesso: entrega.mac, "Categoria de Serviço": entrega.cat,
      Órgão: orgAtual.sig, Unidade: unidade.sig, DFT: unidade.dash ? unidade.dash.dft : "(em andamento)",
      Revisor: revisor || "—", "Registrado em": agora,
    }]);
    mostrarToast(`Serviço registrado para ${entrega.cd} — vai para a planilha compartilhada.`);
  }

  const textoBusca = norm(busca);
  const bate = (n) => !textoBusca || norm(n.sig).includes(textoBusca) || norm(n.nome).includes(textoBusca);

  /* visão renderizada = árvore oficial + personalizações + filtro "Só com DFT" */
  const raizVisivel = useMemo(() => {
    if (!orgAtual) return orgAtual;
    let t = aplicarPersonalizacao(orgAtual, ops);
    if (soDft) t = filtrarSoComDft(t) || { ...cloneRaso(t), filhos: [] };
    return t;
  }, [orgAtual, ops, soDft]);

  /* busca global (órgãos + unidades) com dropdown de resultados */
  const resultadosBusca = useMemo(() => {
    const q = norm(busca || "");
    if (!q || q.length < 2) return null;
    const orgs = [], unidades = [];
    (dados.orgaos || []).forEach((org, i) => {
      if (norm(org.sig).includes(q) || norm(org.nome).includes(q)) orgs.push({ o: org, i });
      const pilha = (org.filhos || []).map((c) => [c, org]);
      for (let g = 0; pilha.length && g < 3e4; g++) {
        const [n, pai] = pilha.pop();
        (n.filhos || []).forEach((c) => pilha.push([c, n]));
        if (norm(n.sig).includes(q) || norm(n.nome).includes(q)) unidades.push({ u: n, o: org, i, p: pai });
      }
    });
    return { os: orgs.slice(0, 8), us: unidades.slice(0, 12), nu: unidades.length };
  }, [busca, dados]);

  const [buscaAberta, setBuscaAberta] = useState(false);
  const [listaPropriaDe, setListaPropriaDe] = useState(new Set());

  function selecionarResultado(idx, textoBuscaUnidade, uidPaiCluster) {
    setOrgIdx(idx); setBusca(textoBuscaUnidade || ""); setBuscaAberta(false);
    if (textoBuscaUnidade) setTimeout(() => setColapsados(uidPaiCluster ? new Set(["clu:" + uidPaiCluster]) : new Set()), 60);
  }

  function DropdownBusca() {
    return (
      <div className="od-gdrop">
        {resultadosBusca.os.length > 0 && <div className="od-gdrop-h">Órgãos</div>}
        {resultadosBusca.os.map((r) => (
          <button key={"o" + r.o.uid} className="od-gdrop-it" onMouseDown={(e) => { e.preventDefault(); selecionarResultado(r.i, ""); }}>
            <span className="od-gdrop-sig">{r.o.sig}</span><span className="od-gdrop-nm">{r.o.nome}</span>
          </button>
        ))}
        {resultadosBusca.us.length > 0 && <div className="od-gdrop-h">Unidades</div>}
        {resultadosBusca.us.map((r, k) => (
          <button key={"u" + k + r.u.uid} className="od-gdrop-it"
            onMouseDown={(e) => { e.preventDefault(); selecionarResultado(r.i, r.u.sig, r.p ? r.p.uid : null); }}>
            <span className="od-gdrop-sig">{r.u.sig}</span><span className="od-gdrop-nm">{r.u.nome}</span>
            <span className="od-gdrop-org">{r.o.sig}</span>
          </button>
        ))}
        {!resultadosBusca.os.length && !resultadosBusca.us.length && <div className="od-gdrop-vazio">Nada encontrado.</div>}
        {resultadosBusca.nu > 12 && <div className="od-gdrop-mais">+{resultadosBusca.nu - 12} unidades correspondem — refine a busca</div>}
      </div>
    );
  }

  /* unidade "efetiva" a exibir no painel: se a selecionada está englobada e o
     usuário não pediu lista própria, mostra o painel da coordenação-geral */
  const coberturaDaSelecao = selecao && selecao.status === "englobado" && !listaPropriaDe.has(selecao.uid) && selecao.cov
    ? (function acharCov(raiz, uid) {
        if (!raiz) return null;
        if (raiz.uid === uid) return raiz;
        let r = null;
        (raiz.filhos || []).some((c) => (r = acharCov(c, uid), !!r));
        return r;
      })(orgAtual, selecao.cov.uid)
    : null;
  const unidadeEfetiva = coberturaDaSelecao || selecao;

  // Busca as entregas da unidade selecionada sob demanda, via API real
  // (substituiu o antigo dados.entregas[sig] estático). Cache por idU em
  // `entregasPorUnidade` evita rebuscar ao reabrir a mesma unidade na
  // sessão. Ver GUIA-integracao-dados-sisdip.md, rotas-entregas-unidade.js.
  const [entregasPorUnidade, setEntregasPorUnidade] = useState({});
  const [idUCarregando, setIdUCarregando] = useState(null);
  useEffect(() => {
    const idU = unidadeEfetiva && unidadeEfetiva.idU;
    if (idU == null) return;
    if (entregasPorUnidade[idU] !== undefined) return;
    let cancelado = false;
    setIdUCarregando(idU);
    fetch(`${API_BASE}/api/entregas-unidade/${idU}`)
      .then((resp) => { if (!resp.ok) throw new Error("HTTP " + resp.status); return resp.json(); })
      .then((dadosResp) => {
        if (cancelado) return;
        const lista = (dadosResp.entregas || []).map((x) => ({
          cd: x.codigo_entrega || String(x.id_entrega || ""),
          nm: x.nome_entrega || "",
          serv: x.servico || "",
          mac: x.macroprocesso || "",
          cat: x.categoria_servico || "",
          nat: "",
          noBanco: true,
          horasMes: x.horas_mes,
          qtDemandaReprimida: x.qt_demanda_reprimida,
          qtResultado: x.qt_resultado,
          qtMeta: x.qt_meta,
        }));
        setEntregasPorUnidade((p) => ({ ...p, [idU]: lista }));
      })
      .catch((err) => {
        if (!cancelado) {
          console.error("[entregas-unidade] falha ao buscar:", err);
          setEntregasPorUnidade((p) => ({ ...p, [idU]: [] }));
        }
      })
      .finally(() => { if (!cancelado) setIdUCarregando((x) => (x === idU ? null : x)); });
    return () => { cancelado = true; };
  }, [unidadeEfetiva && unidadeEfetiva.idU]);

  if (carregando) {
    return (
      <div className="od-root" style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100vh", flexDirection: "column", gap: 16 }}>
        <style>{OD_CSS}</style>
        <div>Carregando organograma...</div>
        <button onClick={onClose} style={{ padding: "8px 16px", borderRadius: 8, border: `1px solid ${T.line}`, background: "#fff", cursor: "pointer" }}>Fechar</button>
      </div>
    );
  }
  if (erro) {
    return (
      <div className="od-root" style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100vh", flexDirection: "column", gap: 16 }}>
        <style>{OD_CSS}</style>
        <div style={{ color: T.coral }}>Não foi possível carregar o organograma: {erro}</div>
        <div style={{ color: T.sub, fontSize: 13 }}>Verifique a conexão com o backend ({API_BASE}).</div>
        <button onClick={onClose} style={{ padding: "8px 16px", borderRadius: 8, border: `1px solid ${T.line}`, background: "#fff", cursor: "pointer" }}>Fechar</button>
      </div>
    );
  }

  return (
    <div className="od-root">
      <style>{OD_CSS}</style>
      <div className="od-stripe"><span style={{ background: T.green }} /><span style={{ background: T.yellow }} /><span style={{ background: T.primary }} /></div>

      <header className="od-head">
        <div className="od-title">
          <span className="od-badge"><MapPin size={18} /></span>
          <div>
            <div className="od-h1">Organizações</div>
            <div className="od-h2">Cobertura do DFT por unidade · dados do SISDIP</div>
          </div>
          <div className="od-orgpick">
            <select value={orgIdx ?? ""} onChange={(e) => setOrgIdx(e.target.value === "" ? null : Number(e.target.value))}>
              <option value="">Selecione uma organização…</option>
              {dados.orgaos
                .map((o, i) => ({ o, i }))
                .sort((a, b) => String(a.o.sig || a.o.nome || "").localeCompare(String(b.o.sig || b.o.nome || ""), "pt"))
                .map(({ o, i }) => <option key={o.uid} value={i}>{o.sig} — {o.nome}</option>)}
            </select>
          </div>
        </div>
        <div className="od-headr">
          <div className="od-rev"><label>Revisor</label><input value={revisor} onChange={(e) => setRevisor(e.target.value)} placeholder="Seu nome" /></div>
          <button className="od-back" onClick={onClose} title="Voltar à página principal"><ArrowLeft size={15} /> Voltar ao catálogo</button>
        </div>
      </header>

      {orgAtual ? (
        <Fragment>
          <div className="od-completude">
            <DonutCompletude concl={orgAtual.cobConcl} and={orgAtual.cobAnd} tot={orgAtual.cobTot} />
            <div className="od-completude-txt">
              <div className="od-completude-t">Completude do DFT — {orgAtual.nome}</div>
              <div className="od-completude-legend">
                <span><i style={{ background: T.green }} /> <b>{orgAtual.cobConcl}</b> concluídas</span>
                <span><i style={{ background: T.amber }} /> <b>{orgAtual.cobAnd}</b> em andamento</span>
                <span><i style={{ background: T.grey }} /> <b>{orgAtual.cobTot - orgAtual.cobConcl - orgAtual.cobAnd}</b> sem DFT</span>
                <span className="od-completude-den">de {orgAtual.cobTot} unidades · colegiados fora do cálculo</span>
              </div>
            </div>

            <div className="od-controls-r">
              <div className="od-search">
                <Search size={15} color={T.faint} />
                <input value={busca} onChange={(e) => { setBusca(e.target.value); setBuscaAberta(true); }}
                  onFocus={() => setBuscaAberta(true)} onBlur={() => setTimeout(() => setBuscaAberta(false), 160)}
                  placeholder="Buscar unidade ou órgão…" />
                {buscaAberta && resultadosBusca && <DropdownBusca />}
              </div>
              <button className="od-mini" onClick={recolherTudo}>Recolher</button>
              <button className="od-mini" onClick={expandirTudo}>Expandir</button>
              <button className={foco ? "od-mini od-mini-on" : "od-mini"} onClick={() => setFoco((v) => !v)}
                title="Ao clicar numa unidade, centraliza nela e ajusta o zoom para mostrar a vizinhança">
                <MapPin size={12} /> Focar
              </button>
              <button className={soDft ? "od-mini od-mini-on" : "od-mini"} onClick={() => setSoDft((v) => !v)}
                title="Mostrar apenas as unidades com DFT e o caminho até elas">◉ Só com DFT</button>
              <button className={editando ? "od-mini od-mini-edit" : "od-mini"} onClick={() => setEditando((v) => !v)}
                title="Arraste cartões para reorganizar e use ✕ para ocultar unidades desta visão">✎ Editar</button>
              {editando && <button className="od-mini" onClick={restaurarVisao} title="Desfazer todas as personalizações desta visão">↺ Restaurar</button>}
              <div className="od-zoom">
                <button onClick={() => setZoom((z) => Math.max(0.6, +(z - 0.1).toFixed(2)))}><ZoomOut size={14} /></button>
                <button onClick={() => setZoom(1)}><Maximize2 size={13} /></button>
                <button onClick={() => setZoom((z) => Math.min(1.4, +(z + 0.1).toFixed(2)))}><ZoomIn size={14} /></button>
              </div>
            </div>
          </div>

            <div className="od-lens">
              <span className="od-lens-lb"><Layers size={13} /> Lente</span>
              <div className="od-lens-seg">
                {LENTES.map((l) => (
                  <button key={l.id} className={lente === l.id ? "on" : ""} onClick={() => setLente(l.id)} title={LENTE_DESC[l.id]}>{l.label}</button>
                ))}
              </div>
              <div className="od-lens-leg">
                {legendaLente(lente).map(([cor, texto], i) => (
                  <span key={i}><i className={cor ? "" : "dsh"} style={cor ? { background: cor } : undefined} />{texto}</span>
                ))}
              </div>
            </div>

            {editando && (
              <div className="od-editbanner">
                <MapPin size={13} />
                <span>Modo edição: <b>arraste</b> um cartão sobre a nova unidade superior · <b>✕</b> oculta · vale só para esta visão, no seu navegador.</span>
                {soDft && <span style={{ marginLeft: "auto", opacity: 0.8 }}>(filtro "Só com DFT" ativo)</span>}
              </div>
            )}

            <div className="od-chartwrap" ref={(el) => { /* wrapRef indireto: chartRef aponta pro filho, ver useEffect acima */ }}
              onScroll={onTipLeave} onPointerDown={iniciarPan} onPointerMove={moverPan} onPointerUp={soltarPan} onPointerLeave={soltarPan}>
              <div className="od-chart" ref={chartRef} style={{ transform: `scale(${zoom})` }}>
                <Conectores chartRef={chartRef} zoom={zoom} depA={orgIdx} depB={colapsados} />
                <ul className="oc-tree">
                  <NoOrganograma n={raizVisivel} raiz colapsados={colapsados} toggle={toggleColapso}
                    sel={selecao} setSel={setSelecao} match={bate} qOn={!!textoBusca} lente={lente}
                    metricas={metricas} onTipEnter={onTipEnter} onTipLeave={onTipLeave}
                    edit={{ on: editando, move: moverUnidade, remove: ocultarUnidade }} />
                </ul>
              </div>
            </div>

          {selecao && (
            <div className={`od-panelwrap${maxPanel ? " max" : ""}`} style={maxPanel ? undefined : { height: panelH }}>
              <PainelUnidade
                key={unidadeEfetiva.uid}
                org={orgAtual}
                unidade={unidadeEfetiva}
                entregas={entregasPorUnidade[unidadeEfetiva.idU] || []}
                entregasCarregando={idUCarregando === unidadeEfetiva.idU}
                overrides={overridesPorUnidade}
                onInserir={registrarServico}
                flash={mostrarToast}
                origem={coberturaDaSelecao ? selecao : null}
                onListaPropria={coberturaDaSelecao ? () => setListaPropriaDe((s) => new Set(s).add(selecao.uid)) : null}
                onClose={() => setSelecao(null)}
                panelH={panelH} setPanelH={setPanelH}
                maxPanel={maxPanel} setMaxPanel={setMaxPanel}
                onMontarNoCatalogo={onMontarNoCatalogo}
              />
            </div>
          )}
        </Fragment>
      ) : (
        <div className="od-choose">
          <div className="od-choose-card">
            <span className="od-choose-ic"><MapPin size={30} /></span>
            <h3>Escolha uma organização</h3>
            <p>Selecione um órgão no seletor acima — ou busque abaixo por órgão ou unidade.</p>
            <div className="od-search od-gsearch-hero">
              <Search size={15} color={T.faint} />
              <input value={busca} onChange={(e) => { setBusca(e.target.value); setBuscaAberta(true); }}
                onFocus={() => setBuscaAberta(true)} onBlur={() => setTimeout(() => setBuscaAberta(false), 160)}
                placeholder="Buscar órgão ou unidade…" />
              {buscaAberta && resultadosBusca && <DropdownBusca />}
            </div>
          </div>
        </div>
      )}

      <Tooltip tip={tip} />
      {toast && <div className="od-toast">{toast}</div>}
    </div>
  );
}

/* ============================================================================
   APP SEPARADO — Organizações & DFT (acesso próprio: organizacoes.html)
   Reaproveita o OrganogramaDimensionamento por inteiro, com um header fino
   próprio (60px, mesma altura do offset do .od-root) e link de volta ao
   Catálogo de Serviços.
============================================================================ */
export function OrganizacoesApp(){
  return (
    <div className="pxo-app">
      <style>{`
        .pxo-app{font-family:'Raleway',-apple-system,'Segoe UI',sans-serif;}
        .pxo-head{position:fixed;top:0;left:0;right:0;height:57px;box-sizing:border-box;background:#fff;border-bottom:1px solid ${C.line};display:flex;align-items:center;justify-content:space-between;padding:0 24px;z-index:95;}
        .pxo-stripe{position:fixed;top:57px;left:0;right:0;height:3px;background:linear-gradient(90deg,#168821 0 33%,#FFCD07 33% 66%,#1351B4 66% 100%);z-index:95;}
        .pxo-brand{display:flex;align-items:center;gap:11px;}
        .pxo-brand-w{font-size:16px;font-weight:800;color:${C.navy};letter-spacing:-.2px;}
        .pxo-brand-s{font-size:10px;color:${C.faint};font-weight:600;letter-spacing:.04em;}
        .pxo-back{display:flex;align-items:center;gap:7px;border:1px solid ${C.line};background:#fff;color:${C.primary};font-family:inherit;font-size:12px;font-weight:700;padding:8px 14px;border-radius:9px;cursor:pointer;text-decoration:none;}
        .pxo-back:hover{border-color:${C.primary};}
      `}</style>
      <header className="pxo-head">
        <div className="pxo-brand">
          <svg viewBox="0 0 32 24" width="30" height="23"><circle cx="7" cy="7" r="4" fill={C.yellow}/><circle cx="16" cy="6" r="4.4" fill={C.green}/><circle cx="25" cy="7" r="4" fill={C.primary}/><path d="M2 22c1.5-5 9-5 10.5 0z" fill={C.yellow}/><path d="M10.5 22c1.5-6 9.5-6 11 0z" fill={C.green}/><path d="M20 22c1.5-5 9-5 10.5 0z" fill={C.primary}/></svg>
          <div><div className="pxo-brand-w">Organizações &amp; DFT</div><div className="pxo-brand-s">Organograma · dimensionamento · relatório por unidade</div></div>
        </div>
        <a className="pxo-back" href="./index.html">Catálogo de Serviços →</a>
      </header>
      <div className="pxo-stripe"/>
      <OrganogramaDimensionamento selPrincipal={[]}
        onClose={()=>{ window.location.href="./index.html"; }}
        onMontarNoCatalogo={()=>{ window.location.href="./index.html"; }}/>
    </div>
  );
}
