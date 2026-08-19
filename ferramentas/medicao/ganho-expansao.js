const fs=require('fs');
const src=fs.readFileSync('portal-frontend.jsx','utf8');
function bloco(n){ const i=src.indexOf('function '+n+'('); let j=src.indexOf('{',i),c=0;
  for(let k=j;k<src.length;k++){ if(src[k]==='{')c++; else if(src[k]==='}'){c--; if(!c) return src.slice(i,k+1);} } }
const STOP=src.match(/const STOPWORDS=new Set\(\[[^\]]*\]\);/)[0];
const banco=JSON.parse(fs.readFileSync('banco.json','utf8'));
const ENTREGAS=banco.map(x=>({codigo:x.c,entrega:x.e,atividade:x.a||"",servico:(x.s&&x.s!==x.k)?x.s:"",categoria:x.k,macro:x.m,natureza:x.n}));
const {buscarCandidatas:buscar}=new Function('ENTREGAS',`
  const norm=s=>(s||"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"");
  ${STOP} ${bloco('stem')} ${bloco('tokenize')} let _IDF=null; ${bloco('construirIDF')} ${bloco('buscarCandidatas')}
  return {buscarCandidatas};`)(ENTREGAS);

/* Simula a etapa de expansão: o que um modelo produziria ao reescrever a frase
   do gestor no vocabulário do catálogo (resultado no particípio, termos do
   domínio). Escrito à mão aqui só para VALIDAR a arquitetura antes de construir. */
const casos=[
 {alvo:'Concursos públicos formalizados',
  cru:'abertura de certames para ingresso de novos servidores no quadro',
  exp:'concurso público formalizado provimento de cargos ingresso de pessoal autorização de concurso'},
 {alvo:'Materiais didáticos e instrucionais desenvolvidos',
  cru:'produção de apostilas e conteúdo para as capacitações da escola',
  exp:'material didático instrucional desenvolvido conteúdo de curso capacitação'},
 {alvo:'Contratos firmados',
  cru:'celebração de ajustes e parcerias com outras instituições',
  exp:'contrato firmado convênio celebrado parceria institucional acordo de cooperação'},
 {alvo:'Ações de mitigação de riscos implementadas',
  cru:'adoção de medidas para reduzir ameaças ao andamento dos trabalhos',
  exp:'ação de mitigação de riscos implementada gestão de riscos plano de contingência'},
 {alvo:'Consultoria técnica e orientação sobre aspectos formais da solicitação de concurso realizada',
  cru:'orientar os órgãos sobre a forma correta de pedir autorização de certame',
  exp:'consultoria técnica orientação sobre aspectos formais da solicitação de concurso realizada'},
 {alvo:'Ranqueamento e definição de prioridades para autorizações de concursos realizados',
  cru:'organizar a fila de pedidos de concurso por ordem de prioridade',
  exp:'ranqueamento e definição de prioridades para autorização de concurso'},
];
const pos=(q,alvo,k)=>{ const r=buscar(q,k).map(x=>x.entrega.trim()); return r.findIndex(y=>y===alvo.trim()); };
let a=0,b=0;
console.log('=== sem expansão  ->  com expansão (posição dentro de 80) ===');
casos.forEach(({cru,exp,alvo})=>{
  const p1=pos(cru,alvo,80), p2=pos(exp,alvo,80);
  if(p1>=0)a++; if(p2>=0)b++;
  console.log('  '+(p1<0?'AUSENTE':'pos '+(p1+1)).padEnd(9)+' -> '+(p2<0?'AUSENTE':'pos '+(p2+1)).padEnd(9)+' | '+alvo.slice(0,44)+'…');
});
console.log('\n  recuperados dentro de 80: '+a+'/6  ->  '+b+'/6');
