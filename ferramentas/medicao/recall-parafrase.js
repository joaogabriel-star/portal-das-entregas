const fs=require('fs');
const src=fs.readFileSync('portal-frontend.jsx','utf8');
function bloco(nome){ const i=src.indexOf('function '+nome+'('); let j=src.indexOf('{',i),n=0;
  for(let k=j;k<src.length;k++){ if(src[k]==='{')n++; else if(src[k]==='}'){n--; if(!n) return src.slice(i,k+1);} } }
const STOP=src.match(/const STOPWORDS=new Set\(\[[^\]]*\]\);/)[0];
const banco=JSON.parse(fs.readFileSync('banco.json','utf8'));
const ENTREGAS=banco.map(x=>({codigo:x.c,entrega:x.e,atividade:x.a||"",servico:(x.s&&x.s!==x.k)?x.s:"",categoria:x.k,macro:x.m,natureza:x.n}));
const {buscarCandidatas:buscar}=new Function('ENTREGAS',`
  const norm=s=>(s||"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"");
  ${STOP} ${bloco('stem')} ${bloco('tokenize')} let _IDF=null; ${bloco('construirIDF')} ${bloco('buscarCandidatas')}
  return {buscarCandidatas};`)(ENTREGAS);

// paráfrases com vocabulário DIFERENTE da entrega — o caso que realmente testa
const casos=[
  {q:'abertura de certames para ingresso de novos servidores no quadro',           alvo:'Concursos públicos formalizados'},
  {q:'produção de apostilas e conteúdo para as capacitações da escola',            alvo:'Materiais didáticos e instrucionais desenvolvidos'},
  {q:'celebração de ajustes e parcerias com outras instituições',                  alvo:'Contratos firmados'},
  {q:'adoção de medidas para reduzir ameaças ao andamento dos trabalhos',          alvo:'Ações de mitigação de riscos implementadas'},
  {q:'orientar os órgãos sobre a forma correta de pedir autorização de certame',   alvo:'Consultoria técnica e orientação sobre aspectos formais da solicitação de concurso realizada'},
  {q:'organizar a fila de pedidos de concurso por ordem de prioridade',            alvo:'Ranqueamento e definição de prioridades para autorizações de concursos realizados'},
];
console.log('=== paráfrases com vocabulário diferente ===');
let achou20=0,achou80=0,achou200=0;
casos.forEach(({q,alvo})=>{
  const pos=k=>{ const r=buscar(q,k).map(x=>x.entrega.trim()); const i=r.findIndex(y=>y===alvo.trim()); return i; };
  const p200=pos(200);
  const marca=p200<0?'AUSENTE':'pos '+(p200+1);
  if(p200>=0&&p200<20)achou20++; if(p200>=0&&p200<80)achou80++; if(p200>=0)achou200++;
  console.log('  '+String(marca).padEnd(10)+' <- "'+q.slice(0,52)+'…"');
  if(p200<0) console.log('             alvo nunca recuperado: "'+alvo.slice(0,58)+'"');
});
const n=casos.length;
console.log('\n  dentro de 20 : '+achou20+'/'+n);
console.log('  dentro de 80 : '+achou80+'/'+n+'   <- e o que vai ao modelo hoje');
console.log('  dentro de 200: '+achou200+'/'+n);
