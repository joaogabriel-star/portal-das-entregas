const fs=require('fs');
const src=fs.readFileSync('portal-frontend.jsx','utf8');

// extrai as funcoes reais do portal, sem reimplementar
function bloco(nome){
  const i=src.indexOf('function '+nome+'(');
  if(i<0) throw new Error('nao achei '+nome);
  let j=src.indexOf('{',i), n=0;
  for(let k=j;k<src.length;k++){ if(src[k]==='{')n++; else if(src[k]==='}'){n--; if(!n) return src.slice(i,k+1);} }
}
const STOP=src.match(/const STOPWORDS=new Set\(\[[^\]]*\]\);/)[0];

const banco=JSON.parse(fs.readFileSync('banco.json','utf8'));
const ENTREGAS=banco.map(x=>({codigo:x.c,entrega:x.e,atividade:x.a||"",servico:(x.s&&x.s!==x.k)?x.s:"",categoria:x.k,macro:x.m,natureza:x.n}));

const ctx=new Function('ENTREGAS',`
  const norm=s=>(s||"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"");
  ${STOP}
  ${bloco('stem')}
  ${bloco('tokenize')}
  let _IDF=null;
  ${bloco('construirIDF')}
  ${bloco('buscarCandidatas')}
  return {buscarCandidatas,tokenize};
`)(ENTREGAS);

const buscar=ctx.buscarCandidatas;

// ---- caso A: atribuicao derivada mecanicamente da propria entrega ----
// "Portaria emitida" -> "emitir portaria"  (aproximacao grosseira mas honesta)
function paraInfinitivo(ent){
  const p=ent.trim().replace(/\.$/,'').split(/\s+/);
  const ult=p[p.length-1].toLowerCase();
  const map={ada:'ar',ado:'ar',adas:'ar',ados:'ar',ida:'ir',ido:'ir',idas:'ir',idos:'ir'};
  let verbo=null;
  for(const suf of Object.keys(map).sort((a,b)=>b.length-a.length)){
    if(ult.endsWith(suf)){ verbo=ult.slice(0,-suf.length)+map[suf]; break; }
  }
  if(!verbo) return null;
  return verbo+' '+p.slice(0,-1).join(' ').toLowerCase();
}

const amostra=[];
for(let i=0;i<ENTREGAS.length && amostra.length<300;i+=137){
  const e=ENTREGAS[i]; const q=paraInfinitivo(e.entrega);
  if(q && q.split(/\s+/).length>=3) amostra.push({e,q});
}

function recall(k){
  let ok=0;
  amostra.forEach(({e,q})=>{ const r=buscar(q,k); if(r.some(x=>x.codigo===e.codigo)) ok++; });
  return (ok/amostra.length*100).toFixed(1);
}
console.log('=== A) recall com atribuicao derivada da entrega ('+amostra.length+' casos) ===');
[20,50,80,150].forEach(k=>console.log('  recall@'+String(k).padEnd(4)+' = '+recall(k)+'%'));

// ---- caso B: atribuicoes REAIS da Enap, com gabarito conferido a mao ----
const reais=[
  {t:'planejar, coordenar e supervisionar a execução logística e operacional dos concursos públicos sob responsabilidade da Enap',
   esperado:['Concursos públicos formalizados']},
  {t:'coordenar a elaboração de normas, editais, conteúdos programáticos e materiais instrucionais de concursos públicos',
   esperado:['Análise de solicitações de concursos públicos realizada','Matriz com pontuação da solicitação de concurso elaborada']},
  {t:'prestar consultoria técnica e apoio institucional a órgãos para implantação de processos seletivos e concursos públicos',
   esperado:['Consultoria técnica e orientação sobre aspectos formais da solicitação de concurso realizada']},
  {t:'planejar, acompanhar e fiscalizar contratos, convênios e parcerias firmados com outras instituições',
   esperado:['Contratos firmados ']},
];
console.log('\n=== B) atribuicoes reais da Enap (gabarito manual) ===');
reais.forEach(({t,esperado})=>{
  const r80=buscar(t,80).map(x=>x.entrega);
  const achou=esperado.filter(x=>r80.some(y=>y.trim()===x.trim()));
  console.log('  '+(achou.length===esperado.length?'OK   ':'FALHA')+' ['+achou.length+'/'+esperado.length+'] '+t.slice(0,58)+'…');
  esperado.filter(x=>!r80.some(y=>y.trim()===x.trim())).forEach(x=>console.log('        nao veio em 80: "'+x.trim().slice(0,52)+'"'));
});
