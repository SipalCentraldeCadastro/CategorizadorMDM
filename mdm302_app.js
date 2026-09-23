/* Categorizador MDM 302 - v7.8 | arquivo irmao de index.html | sem CDN, sem fetch, sem API, sem credencial */
/* motor de categorizacao, similaridade, fluxos 301/302/303/304, painel admin e semente de regra */

/* ================= CATALOGOS E REGRAS EMBARCADAS ================= */
let TIPOS = [
 ["ERSA","Peca de Reposicao","Componentes para manutencao e reparo de maquinas e equipamentos (rolamentos, motores, correias, sensores)"],
 ["HIBE","Material Auxiliar ou de Consumo","Uso no dia a dia da operacao e manutencao (lubrificantes, produtos de limpeza, EPIs, material de escritorio)"],
 ["NLAG","Material Nao Estocavel","Itens nao estocaveis, nao palpaveis ou sucata"]
];
/* ================= CATALOGOS OFICIAIS DA BASE DE CONHECIMENTO (atualizados) =================
   Unica fonte de codigo valido para sugerir, gravar e exportar:
   650 TIPOS (3) | 663 Grupo de Mercadoria - 302 (20) | 627 Grupo de Compradores (18).
   Codigo que existe no historico da base 302 do site e NAO consta nestes catalogos e
   DESCONTINUADO: nunca sugerido, nunca gravado, nunca exportado, nem por edicao manual. */
/* ================= CATALOGOS OFICIAIS - PLANILHAS ANEXADAS (FONTE PRIMARIA, v7) =================
   Grupo_de_Mercadoria BASE.xlsx (20 codigos) e Grupo_de_Compradores BASE.xlsx (18 codigos) lidos
   linha a linha e embarcados abaixo. Sao a fonte primaria de codigo valido: nada fora destas listas
   (ou de uma planilha recarregada pelo administrador) pode ser sugerido, escolhido ou exportado.
   Codigo do historico da base 302 que nao consta aqui e tratado como fora de catalogo e o motor
   procura o mais proximo disponivel por familia, substantivo ou categoria. */
/* v7.8: catalogos 650/663/627 movidos para mdm302_catalogos.js (arquivo irmao) */
let GM_OFF = {}, GC_OFF = {};
const GM_FAM = {}, GM_TXT = {}, GC_TXT = {};
SHEET_GM.forEach(function(r){ GM_OFF[r[0]]=r[1]; GM_FAM[r[0]]=r[2]; GM_TXT[r[0]]=r[3] });
SHEET_GC.forEach(function(r){ GC_OFF[r[0]]=r[1]; GC_TXT[r[0]]=r[2] });
const GM_LEG = {"90702":"UTENSILIO/COPA-COZINHA","90703":"ALIMENTACAO","90704":"MOVEIS/UTENSILIOS","90706":"LABORATORIO","90710":"UNIFORMES","90005":"IMOBILIZADO ANTIGO","90713":"grupo antigo","10001":"grupo antigo","10071":"grupo antigo","10072":"grupo antigo","80001":"IMOB ELETRONICO","80002":"IMOB IMOVEIS","80004":"IMOB SOFTWARE","80005":"IMOB VEICULOS","80006":"IMOB MOBILIARIO"};
const GC_LEG = {"MR1":"MANUTENCAO E REPOSICAO","AL1":"ALIMENTACAO","AU1":"UTENSILIOS/COPA","IA1":"ITENS ADMINISTRATIVOS","MQ1":"MAQUINAS/PECAS","EP1":"EPI","ME2":"MEDICAMENTOS","IS1":"INFORMATICA/SISTEMAS","LB1":"LUBRIFICANTES","OP1":"OBRAS E PROJETOS","SH1":"SUPRIMENTOS HARDWARE","IN1":"INSUMOS (REVENDA)","I02":"MERCADORIAS - IMPORTACAO","IM1":"IMPORTACAO GRAOS","SE3":"SEGURO PATRIMONIAL","SL1":"SERVICOS LOGISTICOS","SL3":"SERVICOS LOCACAO","SV4":"SERVICOS DE VIGILANCIA","SG5":"grupo antigo","F01":"grupo antigo","IF1":"grupo antigo","II2":"grupo antigo","TI1":"grupo antigo","FM1":"grupo antigo","MC4":"grupo antigo"};
/* ===== v7.1: a planilha oficial manda. Codigo que consta nela nao pode ficar marcado como
   legado nem como descontinuado, e nao pode ter "sucessor": ele proprio e o codigo valido. ===== */
(function(){
 for(const k in GM_OFF){ if(GM_LEG[k]) delete GM_LEG[k] }
 for(const k in GC_OFF){ if(GC_LEG[k]) delete GC_LEG[k] }
})();
/* view apenas de leitura (oficial + descontinuado). A VALIDACAO NUNCA usa esta view. */
let GM = Object.assign({}, GM_LEG, GM_OFF);
let GC = Object.assign({}, GC_LEG, GC_OFF);

/* ===== REGISTRO DE CODIGOS DESCONTINUADOS: trava absoluta de sugestao e exportacao =====
   vale para os catalogos 650 / 663 / 627 E para os Grupos de Compradores. Um codigo aqui
   nunca e sugerido, nunca entra em lista de escolha e nunca e exportado. */
const DISC = {tm:{}, gm:{}, gc:{}};
const DISC_SHEET = {tm:{}, gm:{}, gc:{}};
const DISC_RX = /(DESCONTINU|INATIV|OBSOLET|NAO USAR|NAO UTILIZAR|BLOQUEAD|CANCELAD|ENCERRAD|DESATIV|EXPIRAD|SEM USO|DEPRECAD|SUBSTITUIDO POR|NAO CADASTRAR)/;
let DPEND = [];
(function(){ for(const k in GM_LEG) DISC.gm[k]=GM_LEG[k]||1; for(const k in GC_LEG) DISC.gc[k]=GC_LEG[k]||1; })();
function discMark(f,c,why){ if(!f||!c||!DISC[f]) return; DISC[f][c]=why||1;
 if(f==="gm"){ delete GM_OFF[c] } else if(f==="gc"){ delete GC_OFF[c] }
 else if(f==="tm"){ TIPOS=TIPOS.filter(function(t){ return t[0]!==c }) } }
function discFlush(f){ DPEND.forEach(function(p){ DISC_SHEET[f][p[0]]=p[1]; discMark(f,p[0],p[1]) }); DPEND=[]; }
function discCount(){ return Object.keys(DISC.gc).length+Object.keys(DISC.gm).length }
/* ====== CATALOGOS CARREGAVEIS: a lista oficial passa a vir das planilhas da base de conhecimento ======
   O usuario carrega 650 / 663 / 627 na area administrativa e a trava de codigo passa a valer para a
   lista nova, sem depender de nova versao do arquivo. Fallback: catalogos embutidos acima. */
const CAT_EMB = {tm:TIPOS.slice(), gm:Object.assign({},GM_OFF), gc:Object.assign({},GC_OFF)};
let CATSRC = {tm:SHEET_SRC.tm, gm:SHEET_SRC.gm, gc:SHEET_SRC.gc};
const CAT_LS = "mdm302_catalogos_v7_1";
function rebuildViews(){
 /* codigo marcado como descontinuado na planilha oficial nunca volta para o catalogo ativo */
 for(const k in DISC_SHEET.gm) delete GM_OFF[k];
 for(const k in DISC_SHEET.gc) delete GC_OFF[k];
 GM = Object.assign({}, GM_LEG, GM_OFF);
 GC = Object.assign({}, GC_LEG, GC_OFF);
 for(const k in GM_OFF){ if(GM_LEG[k]) delete GM_LEG[k]; if(DISC.gm[k] && !DISC_SHEET.gm[k]) delete DISC.gm[k] }
 for(const k in GC_OFF){ if(GC_LEG[k]) delete GC_LEG[k]; if(DISC.gc[k] && !DISC_SHEET.gc[k]) delete DISC.gc[k] }
 /* descricoes dos descontinuados seguem visiveis para leitura, mas fora de qualquer sugestao */
 for(const k in DISC.gm){ if(!GM[k]) GM[k]=(typeof DISC.gm[k]==="string"?DISC.gm[k]:"grupo descontinuado") }
 for(const k in DISC.gc){ if(!GC[k]) GC[k]=(typeof DISC.gc[k]==="string"?DISC.gc[k]:"grupo descontinuado") }
}
/* le pares codigo/descricao de qualquer planilha, aceitando "HIBE - MAT.AUX." ou colunas separadas */
function catPairs(rows){
 const out=[], seen={};
 DPEND=[];
 (rows||[]).forEach(function(r){
  if(!r) return;
  const cells=(Array.isArray(r)?r:[r]).map(function(c){return String(c==null?"":c).replace(/\u00a0/g," ").trim()}).filter(function(c){return c!==""});
  if(!cells.length) return;
  let code="", desc="";
  for(let i=0;i<cells.length && !code;i++){
   const U=cells[i].toUpperCase();
   const m=U.match(/^([A-Z0-9]{2,6})(?:\s*[-\u2013\u2014:|\/]\s*(.*))?$/);
   if(m && /^(\d{5,6}|[A-Z]{1,3}[0-9]{1,2}|[A-Z]{3,5})$/.test(m[1])){
    code=m[1]; desc=(m[2]||"").trim() || cells.slice(i+1).join(" - ");
   }
  }
  if(!code) return;
  if(/^(CODIGO|COD|CODE|TIPO|GRUPO)$/.test(code)) return;
  /* linha marcada como descontinuada na planilha oficial: registra e NUNCA entra no catalogo */
  if(DISC_RX.test(norm(cells.join(" ")))){ seen[code]=1; DPEND.push([code,(desc||"").trim()||1]); return; }
  desc=desc.replace(new RegExp("^"+code+"\\s*[-:|]\\s*"),"").trim();
  if(seen[code]) return; seen[code]=1;
  out.push([code, desc||code]);
 });
 return out;
}
function detectCat(pairs){
 let d=0,l=0,g=0;
 pairs.forEach(function(p){ const c=p[0];
  if(/^\d{5,6}$/.test(c)) d++; else if(/^[A-Z]{1,3}[0-9]{1,2}$/.test(c)) g++; else if(/^[A-Z]{3,5}$/.test(c)) l++; });
 if(d>=g && d>=l && d) return "gm";
 if(g>=l && g) return "gc";
 return l? "tm" : "";
}
function applyCat(f, pairs, src){
 if(!pairs || !pairs.length) return 0;
 if(f==="tm"){
  const hint={}; CAT_EMB.tm.forEach(function(t){hint[t[0]]=[t[1],t[2]]});
  TIPOS = pairs.map(function(p){ const h=hint[p[0]]; return [p[0], h?h[0]:(p[1]||p[0]), h?h[1]:(p[1]||"")] });
 } else if(f==="gm"){ const o={}; pairs.forEach(function(p){o[p[0]]=p[1]}); GM_OFF=o; }
 else if(f==="gc"){ const o={}; pairs.forEach(function(p){o[p[0]]=p[1]}); GC_OFF=o; }
 else return 0;
 CATSRC[f]=src||"planilha carregada";
 discFlush(f);
 rebuildViews(); saveCat();
 try{ sanitizePriors() }catch(e){}
 return pairs.length;
}
function saveCat(){
 try{ localStorage.setItem(CAT_LS, JSON.stringify({tm:TIPOS, gm:GM_OFF, gc:GC_OFF, src:CATSRC, dt:new Date().toISOString()})) }catch(e){}
}
function loadCatSaved(){
 try{
  const s=localStorage.getItem(CAT_LS); if(!s) return false;
  const o=JSON.parse(s);
  if(o.tm && o.tm.length) TIPOS=o.tm;
  if(o.gm && Object.keys(o.gm).length) GM_OFF=o.gm;
  if(o.gc && Object.keys(o.gc).length) GC_OFF=o.gc;
  if(o.src) CATSRC=o.src;
  rebuildViews(); return true;
 }catch(e){ return false }
}
function resetCat(){
 TIPOS=CAT_EMB.tm.slice(); GM_OFF=Object.assign({},CAT_EMB.gm); GC_OFF=Object.assign({},CAT_EMB.gc);
 CATSRC={tm:SHEET_SRC.tm,gm:SHEET_SRC.gm,gc:SHEET_SRC.gc};
 rebuildViews(); try{localStorage.removeItem(CAT_LS)}catch(e){}
}
/* tipo padrao valido: usado quando a evidencia nao aponta um tipo do catalogo */
function tmDefault(){
 if(TIPOS.some(function(t){return t[0]==="HIBE"})) return "HIBE";
 const cnt={}; (S.base||[]).forEach(function(r){ if(r.tm && codeOk("tm",r.tm)) cnt[r.tm]=(cnt[r.tm]||0)+1 });
 const ks=Object.keys(cnt).sort(function(a,b){return cnt[b]-cnt[a]});
 return ks[0] || (TIPOS[0]? TIPOS[0][0] : "");
}
const GM_IMOB = ["80003","80001","80002","80004","80005","80006"];
let SUC_GM = {"90702":"90701","90704":"90701"};
let SUC_GC = {"MR1":"EQ1","MQ1":"EQ1","AU1":"MU1","IA1":"MU1","ME2":"VM1","LB1":"PQ1","OP1":"MC1"};
/* v7.1: sucessor so vale para codigo que realmente saiu da planilha oficial */
(function(){ for(const k in SUC_GM){ if(GM_OFF[k]) delete SUC_GM[k] } for(const k in SUC_GC){ if(GC_OFF[k]) delete SUC_GC[k] } })();
/* ===== APROXIMACAO PARA CODIGO FORA DAS PLANILHAS (v7) =====
   Codigo do historico que nao consta nas planilhas anexadas nao e descartado em silencio:
   o motor procura o mais proximo disponivel nas listas validas, nesta ordem
   1 sucessor oficial declarado  2 mesma familia da planilha  3 semelhanca de descricao/substantivo
   4 para Grupo de Compradores, o predominante do Grupo de Mercadoria escolhido.
   O caminho usado fica registrado em APX e aparece no painel de transparencia. */
const APX = {tm:{}, gm:{}, gc:{}};
function famGm(c){
 const s=String(c||"");
 if(GM_FAM[s]) return GM_FAM[s];
 if(/^907/.test(s)) return "DEMAIS MATERIAIS";
 if(/^9010/.test(s)) return "MATERIAL INSTALACAO";
 if(/^9015/.test(s)) return "MATERIAL CONSTRUCAO";
 if(/^900/.test(s)) return "MATERIAL MANUTENCAO";
 if(/^80/.test(s)) return "IMOBILIZADO";
 if(/^24/.test(s)) return "EMBALAGENS";
 if(/^29/.test(s)) return "COMBUSTIVEIS";
 if(/^14/.test(s)) return "PREPARADOS P/ANIMAIS";
 return "";
}
function apxDesc(f,c){
 if(f==="gm") return GM_LEG[c] || (GM[c]||"") ;
 if(f==="gc") return GC_LEG[c] || (GC[c]||"") ;
 return "";
}
function apxSim(a,b){
 a=norm(a||""); b=norm(b||"");
 if(!a||!b) return 0;
 if(a===b) return 1;
 let s=0;
 try{ s=diceSet(tri(a),tri(b)) }catch(e){ s=0 }
 const ta=a.split(/\s+/), tb=b.split(/\s+/);
 let hit=0; ta.forEach(function(t){ if(t.length>=4 && tb.some(function(u){return u.indexOf(t)===0||t.indexOf(u)===0})) hit++ });
 return Math.max(s, ta.length? hit/ta.length*0.9 : 0);
}
function apxGm(c){
 c=String(c||"");
 if(!c) return {k:"",why:""};
 if(codeOk("gm",c)) return {k:c, why:""};
 if(APX.gm[c]) return APX.gm[c];
 const fam=famGm(c), d=apxDesc("gm",c);
 const imobLeg=/^80/.test(c) || /IMOBILIZ/.test(norm(d));
 let bk="", bs=0, bw="";
 Object.keys(GM_OFF).forEach(function(k){
  const kimob=GM_IMOB.indexOf(k)>=0;
  if(kimob!==imobLeg) return;                     /* imobilizado so aproxima imobilizado */
  let s=0, w=[];
  if(fam && famGm(k)===fam){ s+=0.46; w.push("mesma familia "+fam) }
  const ds=apxSim(d, GM_OFF[k]+" "+(GM_FAM[k]||""));
  if(ds>0){ s+=0.44*ds; w.push("semelhanca de descricao "+Math.round(ds*100)+"%") }
  if(String(k).substr(0,3)===c.substr(0,3)){ s+=0.18; w.push("mesma faixa de codigo") }
  const use=(GM2GC[k]||[]).reduce(function(t,x){return t+x[1]},0);
  s+=Math.min(0.10, use/1200);
  if(s>bs){ bs=s; bk=k; bw=w.join(" + ") }
 });
 if(bk && bs>=0.30){
  const r={k:bk, why:"aproximacao: "+c+" nao consta na planilha e o mais proximo disponivel e "+bk+" "+GM_OFF[bk]+" ("+(bw||"familia/categoria")+")"};
  APX.gm[c]=r; return r;
 }
 const r0={k:"", why:"codigo "+c+" nao consta na planilha e nao ha equivalente proximo - campo segue para REVISAR"};
 APX.gm[c]=r0; return r0;
}
function apxGc(c, gm){
 c=String(c||"");
 if(!c) return {k:"",why:""};
 if(codeOk("gc",c)) return {k:c, why:""};
 const key=c+"|"+(gm||"");
 if(APX.gc[key]) return APX.gc[key];
 const d=apxDesc("gc",c);
 let bk="", bs=0, bw="";
 Object.keys(GC_OFF).forEach(function(k){
  let s=0, w=[];
  const ds=apxSim(d, GC_OFF[k]);
  if(ds>0){ s+=0.72*ds; w.push("semelhanca de descricao "+Math.round(ds*100)+"%") }
  if(k.substr(0,2)===c.substr(0,2)){ s+=0.16; w.push("mesma familia de sigla") }
  if(gm && (GM2GC[gm]||[]).some(function(x){return x[0]===k})){ s+=0.24; w.push("usado no Grupo de Mercadoria "+gm) }
  if(s>bs){ bs=s; bk=k; bw=w.join(" + ") }
 });
 if((!bk || bs<0.34) && gm){
  const lst=(GM2GC[gm]||[]).filter(function(x){return codeOk("gc",x[0])});
  if(lst.length){ bk=lst[0][0]; bs=0.4; bw="predominante do Grupo de Mercadoria "+gm }
 }
 if(bk && bs>=0.45){
  const r={k:bk, why:"aproximacao: "+c+" nao consta na planilha e o mais proximo disponivel e "+bk+" "+GC_OFF[bk]+" ("+(bw||"familia/categoria")+")"};
  APX.gc[key]=r; return r;
 }
 const r0={k:"", why:"codigo "+c+" nao consta na planilha e nao ha equivalente proximo - campo segue para REVISAR"};
 APX.gc[key]=r0; return r0;
}
function apxWhy(f,c,gm){
 if(f==="gm"){ const r=APX.gm[String(c)]; return r? r.why : "" }
 if(f==="gc"){ const r=APX.gc[String(c)+"|"+(gm||"")]||APX.gc[String(c)+"|"]; return r? r.why : "" }
 return "";
}
const GM2GC = {"14001":[["MC1",5],["IP2",2],["MR1",1],["AL1",1]],"24001":[["MC1",8],["ME2",6],["MR1",2]],"24002":[["MR1",1],["MC1",1]],"29001":[["MR1",1],["CO1",1]],"29002":[["CO1",3],["CO2",2]],"90001":[["MR1",7584],["MC1",79],["ME1",55],["MQ1",39],["ME2",26],["EQ1",7]],"90002":[["MR1",2883],["MQ1",120],["MC1",18],["AU1",5],["EQ1",2],["MU1",1]],"90003":[["MR1",898],["AU1",335],["MQ1",52],["MC1",34],["IA1",8],["ME2",4]],"90004":[["MC1",1630],["MR1",1057],["ME1",25],["OP1",22],["MU1",14],["ME2",8]],"90005":[["MU1",10],["MR1",10],["MC1",8],["ME1",3],["ME2",1],["OP1",1]],"90101":[["MC1",100],["MR1",79],["ME1",33],["MU1",6],["ME2",5],["IS1",1]],"90151":[["MR1",129],["MC1",118],["OP1",48],["ME1",12],["ME2",5],["MU1",3]],"90701":[["MC1",436],["MR1",227],["ME2",70],["MU1",13],["EQ1",11],["ME1",9]],"90703":[["AL1",108],["MC1",5],["IA1",1],["ME2",1],["MR1",1]],"90705":[["MR1",462],["ML1",161],["MC1",82],["IA1",12],["ME2",11],["AU1",9]],"90706":[["LB2",183],["MC1",23],["PQ1",7],["MU1",1],["ME2",1],["CO1",1]],"90707":[["MR1",1125],["MC1",451],["ME2",377],["IA1",181],["MU1",11],["IS1",6]],"90708":[["MR1",311],["LB1",141],["MC1",26],["IA1",2],["MQ1",1],["CO1",1]],"90709":[["IS1",51],["MC1",21],["SH1",21],["MR1",15],["ME2",4],["ME1",2]],"90711":[["MR1",362],["EP1",200],["MC1",52],["AU1",13],["UF1",1]],"90712":[["ME2",4],["MC1",2],["AL1",2]]};
const HEAD_GCV_RAW = "ABRACADEIRA:MR1:165:169;ACABAMENTO:MR1:6:8;ACENDEDOR:MC1:1:2;ACESSORIOS:MR1:9:15;ACIONADOR:MR1:2:6;ACOPLADOR:MR1:3:3;ACOPLAMENTO:MR1:63:90;ACUMULADOR:MR1:8:8;ADAM4PDM:MR1:2:4;ADAPTADOR:MR1:24:52;ADESIVO:MR1:57:83;ADITIVO:MR1:17:33;AGENDA:ME2:2:3;AGUA:MC1:2:5;AGULHA:MR1:79:80;AGULHAO:MR1:3:3;ALARGADOR:MR1:2:2;ALAVANCA:MR1:11:13;ALCA:MR1:6:6;ALCOOL:MR1:32:39;ALGODAO:MR1:26:27;ALICATE:MR1:21:34;ALIMENTOS:AL1:91:98;ALONGADOR:MR1:2:2;ALTERNADOR:MR1:8:10;AMACIANTE:ML1:2:3;AMORTECEDOR:MR1:15:20;AMOSTRA:LB2:2:2;AMPERIMETRO:MR1:2:2;ANALISADOR:MR1:3:3;ANEL:MR1:362:378;ANILHA:MC1:37:42;ANTENA:MR1:4:4;ANTI:MR1:3:3;ANTIFERRUGEM:MR1:6:6;APARELHO:MC1:3:4;APLICADOR:MR1:2:3;APOIO:ME2:4:9;APONTADOR:ME2:2:2;ARAME:MR1:19:31;ARCO:MR1:1:2;AREIA:MC1:2:2;ARGAMASSA:MR1:7:14;ARO:MR1:6:6;AROMA:AL1:2:2;AROMATIZADOR/DIFUSOR:ML1:1:2;ARRUELA:MR1:264:271;ARTICULACAO:MR1:8:8;ASA:MR1:2:2;ASPERSOR:MR1:3:3;ASSENTO:MC1:4:5;ASSOALHO:MR1:8:8;ATADURA:MR1:38:39;ATUADOR:MR1:9:10;AUTOMATICO:MR1:7:8;AUTOTRAFO:MR1:1:2;AVENTAL:EP1:2:4;BACIA:MC1:4:8;BALANCA:MR1:3:4;BALANCIM:MR1:2:2;BALAO:LB2:10:13;BALDE:MR1:41:55;BANCO:MR1:2:3;BANDA:MR1:3:3;BANDEIRA:ME2:3:3;BANDEJA:MR1:3:5;BANNER:ME2:2:3;BARBANTE:ME2:2:2;BARRA:MC1:81:155;BARRAMENTO:MR1:16:18;BARREIRA:MR1:4:4;BASE:MR1:18:32;BASTAO/VARA:MC1:2:2;BATENTE:MR1:4:6;BATERIA:MR1:38:61;BEBIDA:AL1:10:11;BECKER:LB2:4:4;BICO:MR1:167:172;BIELA:MR1:3:3;BIELETA:MR1:4:5;BITS:MC1:2:4;BLOCO:MR1:58:67;BLOQUEIO:MR1:4:4;BOBINA:MR1:20:28;BOCAL:MR1:9:9;BOIA:MR1:5:7;BOLSA:MC1:2:3;BOMBA:MR1:57:68;BOMBONA:ME2:3:3;BORNE:MR1:38:42;BORRACHA:MR1:19:25;BOTAO:MR1:50:58;BOTOEIRA:MR1:4:5;BRACADEIRA:MR1:6:6;BRACO:MR1:13:14;BRIDA:MR1:2:2;BRINDES:MC1:1:2;BRITA:MC1:6:14;BROCA:MR1:113:117;BRONZINA:MR1:2:3;BROXA:MR1:5:8;BUCHA:MR1:403:422;BUJAO:MR1:35:35;BURRINHO:MR1:2:2;BUZINA:MR1:6:6;CABECOTE:MR1:15:17;CABO:MR1:211:278;CABRA:MR1:2:2;CACAMBA:MR1:4:6;CADARCO:MR1:3:3;CADEADO:MC1:10:16;CADERNO:ME2:7:7;CAIBRO:MC1:5:8;CAIXA:MC1:134:234;CAL:MC1:3:5;CALCA:EP1:2:3;CALCADO:EP1:73:74;CALCO:MR1:33:38;CALCULADORA:ME2:2:3;CALHA:MR1:2:2;CAMARA:MR1:20:21;CAMERA:MR1:6:7;CAMISA:MR1:6:8;CAMISETA:EP1:8:9;CANALETA:MC1:7:11;CANALIZACAO:MR1:3:4;CANECA:MR1:19:27;CANETA:MR1:171:248;CANTONEIRA:MR1:3:5;CAPA:MR1:46:72;CAPACETE:MR1:152:167;CAPACITOR:MR1:43:43;CAPCTETE:MR1:2:2;CAPO:MR1:2:2;CAPOTA:MR1:3:6;CARBURADOR:MQ1:2:3;CARCACA:MR1:7:9;CARNEIRA:EP1:4:4;CARREGADOR:MR1:4:12;CARRETEL:MR1:2:2;CARRINHO:MC1:4:5;CARTAO:ME2:2:5;CARTOLA:MR1:2:2;CARTUCHO:MR1:245:253;CARV:MR1:1:2;CASQUILHO:MR1:3:3;CASTANHA:MR1:3:3;CATALISADOR:MR1:14:21;CATRACA:MR1:4:4;CAVADEIRA:MC1:5:6;CAVALETE:ME2:2:4;CELULA:MR1:7:9;CESTO:MR1:2:4;CHAPA:MC1:101:187;CHAPEU:ME2:2:4;CHAVE:MC1:127:266;CHAVEIRO:MC1:2:2;CHAVETA:MC1:30:33;CHICOTE:MR1:19:23;CHIP:MC1:5:5;CHUMBADOR:MC1:24:30;CHUVEIRO:MR1:4:5;CILINDRO:MR1:49:55;CIMENTO:MC1:2:3;CINTA:MR1:23:31;CINTO:EP1:4:4;CINTURAO:EP1:1:4;CLAVICULARIO/ARMARIO:ME2:1:2;CLIP:ME2:1:3;CLIPE:MC1:15:23;CLIPES:MC1:32:61;CNV:MC1:8:10;COBERTURA:MR1:2:3;COIFA:AU1:2:2;COLA:ME2:4:4;COLAR:MR1:38:40;COLETE:MC1:3:6;COLHER:MR1:6:8;COLUNA:MR1:7:14;COMBUSTIVEL:CO2:2:2;COMPRESSA:MC1:1:3;COMPRESSOR:MR1:16:17;COMUTADOR:MR1:2:2;CONCRETO:MC1:2:2;CONDENSADOR:MR1:3:4;CONDUTOR:MR1:4:4;CONE:MR1:6:10;CONECTOR:MR1:189:198;CONECTOR/PLUG:MR1:7:12;CONECTOR/PLUGUE:MR1:39:39;CONEXAO:MR1:43:67;CONTADOR:MR1:4:4;CONTATO:MR1:7:7;CONTATOR:MR1:95:98;CONTRA:MR1:13:14;CONTROLADOR:MR1:20:23;CONTROLE:IS1:2:6;CONVERSOR:MR1:7:13;COPO:MR1:2:3;COPO/TACA/XICARA:MU1:1:3;CORANTE:ME2:1:3;CORDA:MR1:8:13;CORDAO:ME2:3:6;CORDOALHA:MR1:5:5;COROA:MR1:3:5;CORPO:MR1:4:4;CORREIA:MR1:114:132;CORRENTE:MR1:37:40;CORTINA/PERSIANA:MC1:1:2;COSSINETE:MR1:6:6;COTOVELO:MR1:32:32;COXIM:MR1:13:15;CREMALHEIRA:MR1:5:5;CREME:ME2:2:3;CRUZETA:MR1:2:3;CUBA:MR1:4:4;CUBETA:LB2:2:2;CUBO:MR1:19:22;CUICA:AU1:1:2;CURATIVO:MC1:2:3;CURVA:MR1:132:133;DAMPER:MR1:2:2;DECALQUE:MR1:5:5;DEFLETOR:MR1:6:6;DENSIMENTRO:LB2:2:2;DENSIMETRO:LB2:2:2;DESANDADOR:MC1:2:4;DESCARBONIZANTE:AU1:1:2;DESCARGA:MR1:2:2;DESEMPENADEIRA:MR1:6:9;DESENGRAXANTE:MR1:6:11;DESINCRUSTANTE:ML1:2:4;DESINFETANTE:MR1:60:69;DESODORANTE:ML1:5:8;DESODORIZADOR:ML1:10:16;DETECTOR:MR1:6:8;DETERGENTE:MR1:160:163;DIAFRAGMA:MR1:1:2;DILUENTE:MC1:10:12;DIODO:MR1:4:4;DISCO:MR1:126:137;DISJUNTOR:MR1:180:185;DISPENSER:MC1:1:3;DISPLAY:SH1:1:2;DISTRIBUIDOR:MR1:3:3;DOBRADICA:MC1:13:28;DOSADOR:MR1:2:3;DUTO:MR1:3:5;EIXO:MR1:62:70;ELASTICO:MR1:4:5;ELETROCALHA:MR1:6:10;ELETRODO:MR1:74:85;ELETRODUTO:MR1:34:43;ELO:MR1:19:19;EMBALAGEM:MR1:2:4;EMBORRACHAMENTO:MR1:2:2;EMBREAGEM:MR1:5:5;EMENDA:MR1:32:34;EMISSOR:MR1:4:4;EMPUNHADURA:MC1:2:2;ENGATE:MR1:16:18;ENGRENAGEM:MR1:46:49;ENVELOPE:ME2:3:5;ENXADA:MR1:5:5;EPOXI:MC1:4:6;ESCADA:MR1:4:9;ESCAPADOR:MR1:2:2;ESCORA:MR1:2:3;ESCOVA:MC1:26:46;ESFIGMOMANOMETRO:ME2:1:2;ESGUICHO:MR1:4:6;ESMERILHADEIRA:MR1:5:5;ESPACADOR:MR1:11:13;ESPAGUETE:MC1:6:9;ESPARADRAPO:MR1:14:15;ESPATULA:MR1:6:7;ESPELHO:AU1:2:5;ESPONJA:ML1:3:4;ESPUMA:MC1:4:9;ESQUADRO:MR1:3:6;ESTATOR:MR1:8:9;ESTEIRA:MR1:4:4;ESTICADOR:MC1:13:21;ESTILETE:ME2:3:5;ESTOPA:MC1:4:4;ETIQUETA:MR1:9:14;EVAPORADOR:MR1:2:2;EXAUSTOR:MR1:7:7;EXPOSITOR/DISPLAY:MC1:2:2;EXTENSAO:MR1:2:3;EXTENSOR:MR1:4:8;EXTINTOR:MC1:10:13;EXTRATOR:MR1:2:3;FACA:MR1:6:7;FACAO:MR1:1:3;FAROL:MR1:19:23;FECHADURA:MR1:14:20;FECHO:MR1:3:4;FERRAMENTAS:MC1:5:5;FERRO:MR1:1:2;FERTILIZANTE:MC1:2:2;FILT:MR1:8:10;FILTRANTE:MR1:38:38;FILTRO:MR1:568:626;FIO:MR1:3:7;FITA:MR1:27:69;FIXACAO:MR1:4:4;FIXADOR:MR1:6:8;FLANELA:ML1:4:4;FLANGE:MR1:69:71;FLEXIVEL:AU1:2:2;FLORES:MC1:3:4;FLUIDO:MR1:7:9;FLUXO:MR1:2:3;FONE:IS1:5:7;FONTE:MR1:24:28;FORMULARIO:ME2:5:8;FORQUILHA:MR1:3:3;FORRO:MC1:5:8;FOTOCONDUTOR:IS1:1:3;FRASCO:LB2:5:7;FREIO:MR1:3:3;FRUTA:AL1:2:3;FUNDO:MR1:1:2;FUNIL:MR1:3:3;FURADEIRA:MR1:4:4;FUSIVEL:MC1:58:116;FUSO:MR1:2:2;GANCHO:MR1:4:6;GARFO:MR1:4:7;GARRA:MR1:4:5;GAS:MR1:6:10;GAVETA:MR1:2:2;GAXETA:MR1:25:25;GAZE:MR1:10:10;GEL:ML1:2:2;GERADOR:MR1:2:2;GESSO:MR1:3:4;GRADE:MR1:6:11;GRAFITE:IA1:2:4;GRAMPEADOR:ME2:4:5;GRAMPO:MR1:46:68;GRAXA:LB1:26:35;GRAXEIRA:MR1:7:7;GRELHA:MR1:7:8;GUARDA:MR1:3:4;GUARNICAO:MR1:8:10;GUIA:MR1:16:19;HASTE:MR1:15:19;HELICE:MR1:3:4;HERBICIDA:IF1:3:6;HIDROMETRO:MR1:2:2;HORIMETRO:MR1:1:2;HUB:IS1:2:2;IDENTIFICADOR:MR1:3:3;IMPERMEABILIZANTE:MC1:12:16;IMPULSOR:MR1:8:8;IND:MR1:3:3;INDICADOR:MR1:4:4;INDUZIDO:MR1:3:3;INSERTO:MR1:10:10;INSETICIDA:MC1:5:12;INTERFACE:MR1:4:4;INTERRUPTOR:MR1:39:64;INTERRUPTOR/TOMADA:MR1:5:8;INVERSOR:MR1:9:12;ISOLADOR:MR1:22:24;ISOLAMENTO:MR1:4:5;ISOLANTE:MR1:4:5;JALECO:EP1:18:20;JANELA:MC1:3:7;JOGO:MR1:2:6;JOGO/CONJUNTO:MR1:105:143;JOGO/CONJUNTO/KIT:MC1:34:60;JUGULAR:EP1:3:4;JUMPER:MR1:2:2;JUNCAO:MR1:2:4;JUNTA:MR1:89:110;KIT:MR1:92:102;LACRE:MC1:5:8;LAMINA:MR1:17:21;LAMPADA:MR1:298:375;LANTERNA:MR1:5:11;LAPIS:MC1:59:65;LAVATORIO:MC1:2:2;LEITO:MR1:5:11;LEITOR:EQ1:1:2;LENCOL:MC1:4:10;LENCOL/COBERTOR/MANTA:MU1:2:3;LENTE:MR1:10:18;LETREIRO:MC1:4:4;LIGACAO:MR1:2:2;LIMA:MR1:7:20;LIMITADOR:MR1:4:4;LIMPA:MC1:2:7;LIMPADOR:ML1:40:67;LINHA:MR1:4:8;LIQUIDO:MR1:2:2;LIXA:MC1:21:32;LIXEIRA:ML1:6:6;LONA:MC1:8:19;LONGARINA:MR1:6:8;LUMINARIA:MR1:23:38;LUVA:MR1:99:232;LUZ:MR1:1:2;MACACAO:EP1:6:6;MACACO:MC1:4:5;MACANETA:MR1:9:12;MACARICO:MC1:3:6;MACHO:MR1:39:39;MANCAL:MR1:108:113;MANCHAO:MR1:8:8;MANDRIL:MR1:13:14;MANETE:MR1:3:3;MANGA:MR1:6:9;MANGOTE:MR1:6:8;MANGUEIRA:MR1:78:89;MANICOTO:MR1:2:2;MANILHA:MR1:5:11;MANIPULO:MR1:4:4;MANOMETRO:AU1:1:2;MANOMETRO/VACUOMETRO:MR1:4:5;MANOPLA:MR1:3:4;MANTA:MC1:6:8;MANUAL:MR1:2:3;MAO:MR1:5:10;MAQUINA:MR1:2:2;MARCADOR:ME2:4:4;MARRETA:ME2:1:3;MARTELETE:MC1:5:7;MARTELO:MC1:10:11;MASCARA:MR1:67:69;MASSA:MC1:47:53;MEDICAMENTO:MR1:172:191;MEDIDOR:MR1:6:12;MEGOMETRO:MR1:2:3;MEIA:MR1:3:7;MEIO:MR1:2:3;MESA:ME2:1:3;METALICO/N:MR1:32:32;METALICO/NAO:MC1:56:58;MICRO:MR1:6:6;MICROMETRO:MR1:5:5;MICRORRUPTOR:MR1:1:2;MISTURA:AL1:1:2;MISTURADOR:EQ1:1:2;MOCHILA:ME2:5:6;MOD:MR1:2:3;MODULO:MR1:24:30;MODULO/CARTAO/PLACA:MR1:27:36;MOLA:MR1:50:52;MOLDE:MR1:7:8;MOLDURA:MR1:3:7;MONITOR:MQ1:2:3;MONTAGE:MR1:2:2;MOTOBOMBA:MR1:2:2;MOTOR:MR1:35:42;MOTORREDUTOR:MR1:6:6;MOTRIZ:MR1:2:2;MOUSE:ME2:1:4;MUFLA:MR1:2:3;MUNHAO:MR1:3:3;NIPLE:MC1:19:23;NIVEL:MC1:6:10;NOME:MR1:27:50;OCULOS:MR1:89:97;OLEO:MR1:318:451;OLHAL:MR1:5:6;PAINEL:MR1:10:23;PAINEL/QUADRO:MC1:5:8;PALETE/PALLET/ESTRADO:MR1:2:3;PALHA:ML1:1:2;PALHETA:MR1:9:15;PANO:MC1:9:19;PAPEL:MC1:20:47;PAPELAO:ME2:4:6;PAQUIMETRO:MC1:2:3;PARA:MR1:21:24;PARABRISA:MR1:8:9;PARAFUSADEIRA:MC1:1:2;PARAFUSO:MR1:1356:1385;PARALAMA:MR1:1:2;PAREDE:MR1:2:2;PARTE:MR1:76:93;PASSA:MR1:3:3;PASTA:MR1:7:12;PASTILHA:MR1:21:25;PATCH:MR1:2:2;PATIM:MR1:5:6;PECAS:MR1:2:2;PEDAL:MR1:6:6;PEDESTAL:MR1:2:2;PEDRA:MC1:3:3;PELICULA:MR1:3:5;PEN:IS1:3:3;PENEIRA:ME2:7:14;PENTE:ME2:2:8;PERFIL:MC1:80:139;PERFILADO:MR1:6:12;PERFURADOR:ME2:1:2;PERNEIRA:EP1:1:3;PERNO:MR1:2:2;PESTANA:MR1:3:5;PIA:ME2:1:3;PICARETA:MC1:2:3;PINCA:MR1:4:5;PINCEL:MC1:5:8;PINCEL/TRINCHA:MC1:14:21;PINHAO:MR1:10:10;PINO:MR1:93:122;PIRULITO:MR1:4:4;PISCA:AU1:2:3;PISO:MC1:18:32;PISTAO:MR1:6:6;PISTOLA:MR1:6:7;PIVO:MR1:5:10;PLACA:MR1:60:106;PLANTA:MC1:4:4;PLAQUETA:MR1:2:2;PLASTICO:MC1:2:2;PLATO:MR1:6:7;PLUGUE:MR1:3:3;PNEU:MR1:31:55;POLIA:MR1:20:22;PONTA:MR1:2:2;PONTEIRA:MR1:6:6;PORCA:MR1:339:352;PORTA:MR1:35:62;POSICIONADOR:MR1:6:6;POSTE:MR1:8:11;POTE:ML1:1:2;POTENCIOMETRO:MR1:3:3;PRANCHA/PRANCHAO:OP1:1:2;PRANCHETA:ME2:1:2;PRATO:MR1:1:3;PREGO:MR1:16:29;PRESILHA:MR1:5:7;PRESSOSTATO:MR1:5:5;PRIMER:MR1:10:11;PRISIONEIRO:MR1:4:4;PROMOCIONAL:ME2:2:2;PROTECAO:MR1:22:29;PROTETOR:MR1:20:31;PROVETA:LB2:4:4;PRUMO:MR1:5:6;PULVERIZADOR:MR1:3:4;PURGADOR:MR1:15:15;PUXADOR:MR1:2:3;QUADRO:MC1:9:24;QUIMICO:LB2:127:166;QUINTA:MQ1:1:2;RACK:MR1:1:2;RADIADOR:MR1:3:4;RALO:MC1:2:2;RAMAL:MR1:2:2;RASPADOR:MR1:8:8;RATICIDA:ME2:2:4;REATOR:MC1:7:8;REBITADOR:MC1:1:2;REBITE:MR1:9:9;REBOLO:MR1:6:6;RECEPTOR:SH1:1:2;REDUCAO:MR1:31:32;REDUTOR:MR1:6:7;REFIL:MR1:3:8;REFLETOR:MC1:15:22;REFORCO:MR1:4:4;REGISTRO:MR1:15:22;REGUA:MR1:6:10;REGULADOR:MR1:17:19;REJUNTE:MR1:11:13;RELE:MR1:48:54;RELOGIO:AU1:3:4;REMENDO:MR1:3:3;REMOVEDOR:MR1:2:2;REPARO:MR1:39:46;RESERVATORIO:MR1:4:5;RESFRIADOR:AU1:1:2;RESIDUO:MC1:3:5;RESINA:MR1:11:14;RESISTENCIA:MR1:3:3;RESISTOR:MR1:3:3;RESPIRADOR:MC1:2:4;RESPIRO:MR1:5:5;RETENTOR:MR1:270:280;RETIFICA:MC1:2:2;RETIFICADOR:MR1:2:3;RETROVISOR:MR1:7:7;REVESTIMENTO:MR1:4:7;RIPA:MC1:4:7;RODA:MR1:18:22;RODAPE:MR1:2:3;RODO:ML1:4:6;ROLAMENTO:MR1:606:632;ROLAMENTOS:MR1:2:2;ROLDANA:MR1:4:6;ROLETE:MR1:33:33;ROLO:MC1:26:32;ROTEADOR:MC1:3:6;ROTOR:MR1:12:12;ROTULA:MR1:4:4;ROTULADOR/ETIQUETADOR:ME2:1:2;ROUPA:MC1:5:7;SABAO:MR1:40:49;SABONETE:IA1:108:167;SACA:MR1:3:3;SACO:ML1:13:21;SACOLA:MC1:5:6;SAIDA:MR1:11:19;SAPATA:MR1:12:14;SAPATILHA:MR1:4:5;SEDE:MR1:6:6;SELADOR:MR1:3:4;SELO:MR1:18:19;SEMI:MR1:2:2;SENSOR:MR1:125:139;SERINGA:MR1:81:81;SERRA:MR1:12:22;SIFAO:MC1:2:2;SILENCIOSO:MR1:3:3;SILICONE:MC1:5:6;SINALEIRO:MR1:3:3;SINALIZADOR:MR1:5:11;SINCRONIZADO:MR1:3:3;SIRENE:MR1:5:5;SOLDA:MR1:3:5;SOLDAVEL:MR1:2:2;SOLENOIDE:MR1:6:7;SOLUCAO:MR1:1:2;SOLVENTE:MC1:6:13;SONDA:AU1:3:3;SOPRADOR:MR1:4:6;SOQUETE:MC1:68:89;SOQUETE/RECEPTACULO:MC1:7:14;SOQUETEIRA:MC1:3:5;SUPORTE:MR1:68:121;SWITCH:MC1:4:5;TABUA:MR1:9:18;TALABARTE:MR1:4:7;TALHADEIRA/PONTEIRO:MR1:10:10;TALISCA:MR1:2:2;TAMBOR:MR1:2:3;TAMPA:MR1:99:134;TAMPAO:MR1:19:39;TANQUE:MR1:8:9;TAPETE/CAPACHO:MC1:8:9;TECIDO:ME2:9:10;TECLADO:MR1:5:5;TEE:MR1:22:30;TELA:MC1:22:37;TELHA:MC1:14:30;TENSIONADOR:MR1:2:3;TENSOR:MR1:5:5;TERMINAL:MR1:353:373;TERMOMETRO:MC1:78:157;TERMOSTATO:MR1:8:8;TESOURA:MC1:148:183;TIJOLO:MR1:10:17;TINTA:MR1:103:176;TIRA:MR1:2:2;TIRANTE:MR1:2:3;TOALHA:ME2:1:3;TOCHA:MR1:2:3;TOMADA:MC1:48:80;TONER:MR1:94:106;TORNEIRA:MR1:26:34;TORQUES:MR1:2:4;TOUCA:EP1:2:4;TRAFO:MC1:6:7;TRANSCEIVER:MR1:3:3;TRANSDUTOR:MR1:4:5;TRANSFORMADOR:MR1:2:2;TRANSMISSOR:MR1:13:15;TRANSMISSOR/MEDIDOR:MR1:2:2;TRAVA:MR1:11:13;TRELICA:MR1:4:5;TRENA:MR1:16:24;TRILHO:MR1:8:8;TUBO:MR1:93:150;TUBO/MANGUEIRA/ESPIRAL/JARDIM:MR1:2:3;TURBOCOMPRESSOR:AU1:1:2;UNIAO:MR1:56:57;UNIDADE:MR1:7:8;UTENSILIO:MU1:2:7;VALVULA:MR1:202:214;VARETA:MR1:6:6;VASELINA:MC1:2:3;VASO:ME2:2:3;VASSOURA:MR1:150:158;VASSOURAO:MR1:2:2;VEDACAO:MR1:56:64;VEDADOR:MR1:4:4;VEDANTE:MR1:1:2;VELA:MQ1:2:5;VENTILADOR:MR1:36:49;VERGALHAO:MR1:8:20;VERNIZ:MR1:4:10;VIDRO:MR1:12:13;VIGA:MC1:2:5;VIRABREQUIM:AU1:2:2;VISOR:MR1:4:4";
const HEAD_GCV = (function(){const o={};HEAD_GCV_RAW.split(";").forEach(function(s){if(!s)return;const p=s.split(":");if(GC_OFF[p[1]])o[p[0]]={k:p[1],n:+p[2],tot:+p[3]}});return o})();
const RULE_GC = [
 /* v7.1: familias amarradas aos 30 codigos do 627 oficial */
 [/(LUVA|CAPACETE|OCULOS DE PROTECAO|MASCARA|RESPIRADOR|PROTETOR AURICULAR|ABAFADOR|BOTINA|BOTA DE SEGURANCA|CINTO DE SEGURANCA|TALABARTE|PERNEIRA|AVENTAL|PROTETOR FACIAL|PROTETOR SOLAR|EPI)/,"EP1","equipamento de protecao individual"],
 [/(PAPEL A4|PAPEL SULFITE|CANETA|LAPIS|CLIPS|GRAMPEADOR|GRAMPO |ENVELOPE|PASTA AZ|PASTA SUSPENSA|MARCA TEXTO|CALCULADORA|PRANCHETA|PERFURADOR|BLOCO DE ANOTACAO|POST IT|APONTADOR|CORRETIVO|EXPEDIENTE)/,"ME2","material de expediente"],
 [/(OLEO LUBRIFICANTE|OLEO HIDRAULICO|OLEO DE MOTOR|OLEO 15W40|OLEO 20W50|GRAXA|LUBRIFICANTE|FLUIDO HIDRAULICO|FLUIDO DE FREIO)/,"LB1","lubrificante"],
 [/(NOTEBOOK|COMPUTADOR|DESKTOP|MONITOR|TECLADO|MOUSE|IMPRESSORA|NOBREAK|SSD|HEADSET|WEBCAM|ESTABILIZADOR)/,"SH1","informatica - hardware"],
 [/(TONER|CARTUCHO|PENDRIVE|CD |DVD |CABO USB|SUPRIMENTO DE INFORMATICA)/,"IS1","informatica - suprimentos"],
 [/(SWITCH |ROTEADOR|ACCESS POINT|CABO DE REDE|PATCH PANEL|PATCH CORD|FIBRA OPTICA|CONECTOR RJ45)/,"TI1","internet / rede"],
 [/(PNEU|CAMARA DE AR|PASTILHA DE FREIO|LONA DE FREIO|AMORTECEDOR|EMBREAGEM|VELA DE IGNICAO|BATERIA AUTOMOTIVA|PARABRISA|RETROVISOR|FILTRO DE COMBUSTIVEL|FILTRO DE AR DO MOTOR)/,"AU1","peca automotiva"],
 [/(ROLAMENTO|RETENTOR|MANCAL|POLIA|CORREIA|CORRENTE TRANSPORTADORA|ENGRENAGEM|REDUTOR|MANGUEIRA HIDRAULICA|VALVULA|ATUADOR|ROLETE)/,"MR1","manutencao e reposicao (M.R.O.)"],
 [/(REFEICAO|LANCHE|MARMITA|CAFE |ACUCAR|BISCOITO|LEITE |AGUA MINERAL|SUCO |ALIMENTO)/,"AL1","produto alimenticio"],
 [/(INSETICIDA|FUNGICIDA|HERBICIDA|FORMICIDA|RATICIDA|DEFENSIVO)/,"IF1","inseticida / fungicida"],
 [/(FERTILIZANTE|ADUBO|SEMENTE|CALCARIO AGRICOLA|UREIA|INSUMO AGRICOLA)/,"IA1","insumo agricola"],
 [/(OBRA PRONTA|SERVICO DE OBRA|EMPREITADA)/,"OP1","obra pronta"],
 [/(MAQUINA |TRATOR|COLHEITADEIRA|EMPILHADEIRA|CAMINHAO|VEICULO|CARRETA)/,"MQ1","maquinas e veiculos"],
 [/(BIG BAG|SACARIA|SACO DE RAFIA|BOBINA PLASTICA|FILME STRETCH|CAIXA DE PAPELAO|EMBALAGEM INDUSTRIAL)/,"II2","insumo de industria / embalagem"],
 [/(DETERGENTE|DESINFETANTE|SABAO|VASSOURA|RODO |ESPONJA|LUSTRA|AMACIANTE|AGUA SANITARIA|LIMPADOR|ODORIZADOR|DESODORIZADOR|SACO DE LIXO|PANO DE CHAO|FLANELA|CERA LIQUIDA)/,"ML1","material de limpeza"],
 [/(PAPEL HIGIENICO|SABONETE|ALCOOL GEL|TOALHA DE PAPEL|HIGIENIZADOR|ABSORVENTE|FRALDA|COPO DESCARTAVEL)/,"MG3","material de higiene"],
 [/(UNIFORME|CAMISA |CAMISETA|CALCA |JAQUETA|JALECO|BLUSA|MACACAO|AGASALHO)/,"UF1","vestimenta / uniforme"],
 [/(CABO ELETRICO|DISJUNTOR|CONTATOR|FUSIVEL|LAMPADA|REATOR|TOMADA|INTERRUPTOR|TRANSFORMADOR|ELETRODUTO|CONDULETE|LUMINARIA|REFLETOR|SOQUETE|ELETROCALHA|INVERSOR DE FREQUENCIA|CHAVE SECCIONADORA|FIO ELETRICO)/,"ME1","material eletrico"],
 [/(REAGENTE|PROVETA|BECKER|BURETA|PIPETA|ERLENMEYER|FENOLFTALEINA|SULFATO|HIDROXIDO|CLORETO|ACIDO |PLACA DE PETRI|MEIO DE CULTURA|BALAO VOLUMETRICO|DESSECADOR)/,"LB2","material de laboratorio"],
 [/(SODA CAUSTICA|POLIMERO|FLOCULANTE|ANTIESPUMANTE|BIOCIDA|COAGULANTE|PERMANGANATO|PEROXIDO|LUBRIFICANTE|GRAXA|SOLVENTE|AGUARRAS|THINNER)/,"PQ1","produto quimico"],
 [/(MEDICAMENTO|VACINA|SERINGA|AGULHA DESCARTAVEL|ANTIBIOTICO|ANALGESICO|POMADA|CURATIVO|VERMIFUGO)/,"VM1","vacina / medicamento"],
 [/(OLEO DIESEL|DIESEL|GASOLINA|ETANOL COMBUSTIVEL|ARLA)/,"CO2","combustivel"],
 [/(LENHA|CAVACO|BIOMASSA)/,"LC1","lenha / cavaco"],
 [/(IMPLEMENTO AGRICOLA|ARADO|GRADE AGRICOLA|PLANTADEIRA|SEMEADEIRA)/,"IP1","implemento agricola"],
 [/(RACAO|SAL MINERAL|CARRAPATICIDA|BRINCO BOVINO)/,"IP2","insumo pecuario"],
 [/(BRINDE|CONFRATERNIZACAO|COFFEE BREAK|TROFEU)/,"EV1","evento / confraternizacao"],
 [/(PASSAGEM AEREA|HOSPEDAGEM|DIARIA DE HOTEL)/,"PH1","passagem / hospedagem"],
 [/(CADEIRA|ARMARIO|ESTANTE|GAVETEIRO|SOFA|COLCHAO|TRAVESSEIRO|LENCOL|CANECA|GARRAFA TERMICA|BANDEJA|TALHER)/,"MU1","movel / utensilio"]
];
function ruleGc(txt){ for(let i=0;i<RULE_GC.length;i++){ if(RULE_GC[i][0].test(txt)) return {k:RULE_GC[i][1], why:RULE_GC[i][2]} } return null }
const RULE_GM = [
 /* v7: familias especificas amarradas aos 20 codigos da planilha Grupo_de_Mercadoria BASE.xlsx */
 [/(PAPEL A4|PAPEL SULFITE|PAPEL SUL|CANETA|LAPIS|BORRACHA ESCOLAR|CLIPS|CLIPE |GRAMPEADOR|GRAMPO |ENVELOPE|PASTA AZ|PASTA SUSPENSA|TONER|CARTUCHO|MARCA TEXTO|CALCULADORA|PRANCHETA|PERFURADOR|BLOCO DE ANOTACAO|POST IT|APONTADOR|CORRETIVO)/,"90707","material de expediente"],
 [/(LUVA|CAPACETE|OCULOS DE PROTECAO|OCULOS|MASCARA|RESPIRADOR|PROTETOR AURICULAR|ABAFADOR|BOTINA|BOTA DE SEGURANCA|CINTO DE SEGURANCA|TALABARTE|PERNEIRA|AVENTAL|PROTETOR FACIAL|PROTETOR SOLAR)/,"90711","equipamento de protecao individual"],
 [/(DETERGENTE|DESINFETANTE|VASSOURA|RODO |SABAO|SAPONACEO|AGUA SANITARIA|ESPONJA|PANO DE LIMPEZA|PAPEL TOALHA|PAPEL HIGIENICO|SACO DE LIXO|LUSTRA MOVEIS|LIMPADOR|MULTIUSO)/,"90705","material de limpeza"],
 [/(CABO FLEXIVEL|CABO ELETRICO|CABO PP|FIO ELETRICO|DISJUNTOR|LAMPADA|TOMADA|INTERRUPTOR|CONTATOR|REATOR|ELETRODUTO|CONDULETE|LUMINARIA|REFLETOR|CANALETA|SOQUETE|ELETROCALHA|INVERSOR DE FREQUENCIA|CHAVE SECCIONADORA)/,"90101","material de instalacao eletrica"],
 [/(CIMENTO|AREIA |BRITA|TIJOLO|BLOCO DE CONCRETO|ARGAMASSA|CONCRETO|TELHA|VERGALHAO|COMPENSADO|GESSO|MASSA CORRIDA|IMPERMEABILIZANTE|CAL HIDRATADA|MADEIRA SERRADA)/,"90151","material de construcao"],
 [/(OLEO LUBRIFICANTE|OLEO HIDRAULICO|OLEO DE MOTOR|OLEO 15W40|OLEO 20W50|GRAXA|LUBRIFICANTE|FLUIDO DE FREIO|FLUIDO HIDRAULICO)/,"90708","oleo / lubrificante"],
 [/(NOTEBOOK|COMPUTADOR|DESKTOP|MONITOR|TECLADO|MOUSE|IMPRESSORA|NOBREAK|SWITCH |ROTEADOR|SSD|PENDRIVE|CABO DE REDE|HEADSET|WEBCAM)/,"90709","material de informatica"],
 [/(VACINA|MEDICAMENTO|SERINGA|AGULHA DESCARTAVEL|VERMIFUGO|CARRAPATICIDA|BRINCO BOVINO)/,"14001","insumo pecuario / veterinario"],
 [/(BRINDE|CONFRATERNIZACAO|COFFEE BREAK|TROFEU|KIT EVENTO)/,"90712","evento / confraternizacao"],
 [/(RACAO|SAL MINERAL|SUPLEMENTO MINERAL)/,"14001","preparado para animais"],
 [/(EMBALAGEM|SACARIA|CAIXA DE PAPELAO|FILME STRETCH)/,"24001","embalagem"],
 [/(OLEO DIESEL|DIESEL S10|GASOLINA|ETANOL COMBUSTIVEL)/,"29002","combustivel fossil"],
 [/(LENHA|CAVACO|BIOMASSA)/,"29003","combustivel lenhoso"],
 [/(CADEIRA|MESA DE ESCRITORIO|ARMARIO|ESTANTE|GAVETEIRO|SOFA|COLCHAO|TRAVESSEIRO|LENCOL|CANECA|GARRAFA TERMICA|BANDEJA|TALHER|COPA|COZINHA)/,"90701","utensilio - 90702 e 90704 descontinuados"],
 [/(REAGENTE|PROVETA|BECKER|BURETA|PIPETA|ERLENMEYER|VIDRARIA)/,"90701","utensilio de laboratorio - 90706 descontinuado"],
 [/(UNIFORME|CAMISA |CAMISETA|CALCA |JALECO|JAQUETA)/,"90711","vestimenta - 90710 descontinuado"]
];
/* ===== v7: FAMILIA OFICIAL GRUPO DE MERCADORIA -> GRUPO DE COMPRADORES =====
   usada como ultimo apoio antes de sugerir REVISAR, no lugar da predominancia geral do 302:
   mantem o Grupo de Compradores coerente com a familia do Grupo de Mercadoria escolhido.
   Sugestao de apoio a decisao - a escolha final e sempre do usuario. */
const GM_GC_FAM = {"14001":["MC1","IP2","MR1"],"24001":["MC1","ME2","MR1"],"24002":["MR1","MC1"],"29001":["MR1","CO1"],"29002":["CO1","CO2"],"90001":["MR1","MC1","ME1"],"90002":["MR1","MQ1","MC1"],"90003":["MR1","AU1","MQ1"],"90004":["MC1","MR1","ME1"],"90005":["MU1","MR1","MC1"],"90101":["MC1","MR1","ME1"],"90151":["MR1","MC1","OP1"],"90701":["MC1","MR1","ME2"],"90703":["AL1","MC1","IA1"],"90705":["MR1","ML1","MC1"],"90706":["LB2","MC1","PQ1"],"90707":["MR1","MC1","ME2"],"90708":["MR1","LB1","MC1"],"90709":["IS1","MC1","SH1"],"90711":["MR1","EP1","MC1"],"90712":["ME2","MC1","AL1"]};
function famGc(gm, txt){
 if(!codeOk("gm",gm)) return {k:"", c:0.25, via:""};
 const rg=(typeof ruleGc==="function")? ruleGc(txt||"") : null;
 const fam=(GM_GC_FAM[gm]||[]).filter(function(c){ return codeOk("gc",c) });
 if(rg && codeOk("gc",rg.k) && (!fam.length || fam.indexOf(rg.k)>=0))
  return {k:rg.k, c:0.80, via:"regra de familia ("+rg.why+") coerente com o Grupo de Mercadoria "+gm};
 if(fam.length)
  return {k:fam[0], c:0.66, via:"familia do Grupo de Mercadoria "+gm+" "+(GM_OFF[gm]||"")+" na planilha oficial sugere "+fam[0]+" "+(GC_OFF[fam[0]]||"")+" - sugerimos confirmar"};
 const g=gcFromGm(gm);
 if(g) return {k:g, c:0.55, via:"predominancia do Grupo de Mercadoria "+gm+" - sugerimos confirmar"};
 return {k:"", c:0.25, via:"nenhum codigo do 627 sustentado pelo Grupo de Mercadoria "+gm+" - sugerimos REVISAR"};
}
function ruleGm(txt){ for(let i=0;i<RULE_GM.length;i++){ if(RULE_GM[i][0].test(txt)) return {k:RULE_GM[i][1], why:RULE_GM[i][2]} } return null }
const FIXOS = {gci:"NORM", gciNlag:"NLAG", setor:"80", aprovado:"APROVADO"};
const COLS4MDG = ["Descrição do material","Tipo de material","Grupo de Mercadorias","Grupo de compradores","Grupo categoria Item","Setor de atividade","Cadastro aprovado?","Classe de avaliação"];
/* codigos presentes no historico da base 302 que nao constam nos catalogos oficiais 663/627 */
const FORA663 = Object.keys(GM_LEG);
const FORA627 = Object.keys(GC_LEG);
const GCI_OPC = ["NORM","NLAG"];
/* ================= CODIGO PURO (regra central: campo classificado guarda SEMPRE o codigo) =================
   qualquer valor que chegue como "90701 - FERRAMENTA/UTENSILIO" (planilha do 4MDG, base carregada,
   colagem manual) e reduzido ao codigo antes de entrar no motor, na tela e na exportacao */
function codeOf(f,v){
 if(v===undefined||v===null) return "";
 let s=String(v).replace(/\u00a0/g," ").trim().replace(/^["']+|["']+$/g,"");
 if(!s) return "";
 const U=s.toUpperCase();
 if(f==="setor"){ const m=U.match(/\d{1,3}/); return m?m[0]:"" }
 if(f==="gm"){ const m=U.match(/\b\d{5,6}\b/); return m?m[0]:U.split(/[\s\-\u2013\u2014:|;,\/]+/)[0] }
 if(f==="tm"){ for(let i=0;i<TIPOS.length;i++){ if(new RegExp("(^|[^A-Z])"+TIPOS[i][0]+"([^A-Z]|$)").test(U)) return TIPOS[i][0] } return U.split(/[\s\-\u2013\u2014:|;,\/]+/)[0] }
 if(f==="gci"){ if(/NLAG/.test(U)) return "NLAG"; if(/NORM/.test(U)) return "NORM"; return U.split(/[\s\-\u2013\u2014:|;,\/]+/)[0] }
 if(f==="gc"){ const m=U.match(/\b[A-Z]{1,3}[0-9]{1,2}\b/); return m?m[0]:U.split(/[\s\-\u2013\u2014:|;,\/]+/)[0] }
 if(f==="aprovado"){ return U.split(/[\s\-\u2013\u2014:|;,\/]+/)[0] }
 if(f==="ca"){ const m=U.match(/\b[0-9]{3,5}\b/); return m?m[0]:U.split(/[\s\-\u2013\u2014:|;,\/]+/)[0].slice(0,8) }
 return s;
}
function codeSafe(f,v){ const c=codeOf(f,v); return codeOk(f,c)?c:"" }
/* forma de codigo valida mas ausente do catalogo oficial = descontinuado (FM1, MC4, MR1, 90710...) */
function descont(f,c){
 if(!c) return false;
 if(f==="gm") return !GM_OFF[c] && /^\d{5,6}$/.test(c);
 if(f==="gc") return !GC_OFF[c] && /^[A-Z]{1,3}[0-9]{1,2}$/.test(c);
 if(f==="tm") return !TIPOS.some(function(t){return t[0]===c});
 return false;
}
function sucOf(f,c){
 c=String(c||""); if(!c) return "";
 const map=(f==="gm"?SUC_GM:(f==="gc"?SUC_GC:{}));
 const s=map[c];
 if(s && codeOk(f,s)){
  if(APX[f]) APX[f][f==="gc"? c+"|" : c]={k:s, why:"sucessor oficial de "+c+" na lista valida: "+s+" "+(f==="gm"?(GM_OFF[s]||""):(GC_OFF[s]||""))};
  return s;
 }
 if(f==="gm"){ const a=apxGm(c); return a.k||"" }
 if(f==="gc"){ const a=apxGc(c,""); return a.k||"" }
 return "";
}
function codeOk(f,c){
 if(!c) return false;
 if(DISC[f] && DISC[f][c]) return false;   /* descontinuado: nunca sugerido, nunca exportado */
 if(f==="tm") return TIPOS.some(function(t){return t[0]===c});
 if(f==="gm") return !!GM_OFF[c];    /* SO codigo do catalogo 663 */
 if(f==="gc") return !!GC_OFF[c];    /* SO codigo do catalogo 627 */
 if(f==="gci") return GCI_OPC.indexOf(c)>=0;
 if(f==="ca") return !!CA_OFF[c];   /* SO classe presente na base/catalogo carregado */
 if(f==="setor") return /^\d{1,3}$/.test(c);
 if(f==="aprovado") return /^[A-Z]+$/.test(c);
 return true;
}
/* fonte do codigo, so para o painel de transparencia (nunca exportado) */
const CAT_SRC = { tm:"catalogo 650 - Materiais Indiretos", gm:"catalogo 663 - Grupo de Mercadoria - 302",
                  gc:"catalogo 627 - Grupo de Compradores", gci:"padrao fixo do fluxo 302",
                  setor:"padrao fixo do fluxo 302", aprovado:"padrao fixo do fluxo 302",
                  ca:"coluna CLASSE DE AVALIACAO da base de materiais" };
const SETOR_OPC = ["80","70","65","60","50","40","10"];

/* ================================================================
   v7.4 - CLASSE DE AVALIACAO (campo novo do motor)
   O catalogo nasce da coluna CLASSE DE AVALIACAO da base de materiais
   (Indiretos ou Imobilizados) e pode ser editado pelo administrador.
   Nenhum codigo e inventado: sem a coluna na base e sem cadastro do
   administrador, o campo sai como REVISAR - a decisao e do usuario.
   ================================================================ */
const CA_KEY="mdm302.ca.v74";
let CA_OFF={};            /* codigo -> descricao */
let CA_SRC="";            /* de onde veio o catalogo */
function caCatLoad(){
 try{ const r=JSON.parse(localStorage.getItem(CA_KEY)||"null");
  if(r&&r.cat){ CA_OFF=r.cat; CA_SRC=r.src||"catalogo salvo no navegador" } }catch(e){}
}
function caCatSave(){ try{ localStorage.setItem(CA_KEY,JSON.stringify({cat:CA_OFF,src:CA_SRC})) }catch(e){} }
function caDesc(c){ return CA_OFF[c]||"" }
function caCount(){ return Object.keys(CA_OFF).length }
function caLearn(v){
 const c=codeOf("ca",v); if(!c) return "";
 if(!CA_OFF[c]){ CA_OFF[c]="classe lida da base de materiais"; if(!CA_SRC) CA_SRC="coluna CLASSE DE AVALIACAO da base de materiais"; caCatSave() }
 return c;
}
/* indice de aprendizado da classe de avaliacao a partir da base */
let CA_IX={sub:{},gm:{},tmgm:{},exato:{},n:0};
function caTop(o){ let k="",n=0,t=0; for(const x in o){ t+=o[x]; if(o[x]>n){n=o[x];k=x} } return k?{k:k,n:n,p:n/t,t:t}:null }
/* v7.8: versao 7.3 preservada e renomeada - era sobreposta pela declaracao mais nova abaixo */
function caBuild_legado_v73(){
 CA_IX={sub:{},gm:{},tmgm:{},exato:{},n:0};
 if(!S.base||!S.base.length) return;
 for(let i=0;i<S.base.length;i++){
  const r=S.base[i]; const c=caLearn(r.ca||""); if(!c) continue; r.ca=c;
  CA_IX.n++;
  const sub=(r._t&&r._t[0])? r._t[0] : (toks((r.d||"")+" "+(r.c||""))[0]||"");
  const gm=r.gm||"", tm=r.tm||"";
  const push=function(o,k){ if(!k) return; if(!o[k]) o[k]={}; o[k][c]=(o[k][c]||0)+1 };
  push(CA_IX.sub, sub+"|"+gm); push(CA_IX.gm, gm); push(CA_IX.tmgm, tm+"|"+gm);
  const ex=norm(r.d||"").replace(/\s+/g," ").trim(); if(ex) CA_IX.exato[ex]=c;
 }
}
/* decide a classe de avaliacao: gemeo exato > substantivo+GM > tipo+GM > GM > REVISAR */
/* v7.8: versao 7.3 preservada e renomeada - era sobreposta pela declaracao mais nova abaixo */
function caPick_legado_v73(desc,sub,gm,tm,twinRec){
 if(twinRec && twinRec.ca && codeOk("ca",twinRec.ca))
  return {c:twinRec.ca,conf:0.95,via:"cadastro existente na base com a mesma descricao: classe "+twinRec.ca+" "+caDesc(twinRec.ca),lista:"coluna CLASSE DE AVALIACAO da base carregada",kind:"exata"};
 if(!CA_IX.n) return {c:"",conf:0,via:"a base carregada nao traz a coluna CLASSE DE AVALIACAO - campo marcado como REVISAR, sem chute",lista:caCount()? "catalogo de classes cadastrado pelo administrador" : "nenhuma lista de classe de avaliacao carregada",kind:"revisar"};
 const ex=norm(desc||"").replace(/\s+/g," ").trim();
 if(ex && CA_IX.exato[ex] && codeOk("ca",CA_IX.exato[ex]))
  return {c:CA_IX.exato[ex],conf:0.94,via:"descricao identica na base: classe "+CA_IX.exato[ex],lista:"coluna CLASSE DE AVALIACAO da base carregada",kind:"exata"};
 const t1=caTop(CA_IX.sub[(sub||"")+"|"+(gm||"")]||{});
 if(t1 && codeOk("ca",t1.k) && t1.t>=3)
  return {c:t1.k,conf:Math.min(0.90,0.55+0.35*t1.p),via:"substantivo \""+sub+"\" com Grupo de Mercadoria "+gm+": "+t1.n+" de "+t1.t+" materiais da base com a classe "+t1.k+" ("+Math.round(t1.p*100)+"%)",lista:"coluna CLASSE DE AVALIACAO da base carregada",kind:t1.p>=0.8?"exata":"aproximada"};
 const t2=caTop(CA_IX.tmgm[(tm||"")+"|"+(gm||"")]||{});
 if(t2 && codeOk("ca",t2.k) && t2.t>=5)
  return {c:t2.k,conf:Math.min(0.80,0.50+0.30*t2.p),via:"tipo de material "+tm+" com Grupo de Mercadoria "+gm+": "+t2.n+" de "+t2.t+" materiais da base com a classe "+t2.k+" ("+Math.round(t2.p*100)+"%)",lista:"coluna CLASSE DE AVALIACAO da base carregada",kind:"aproximada"};
 const t3=caTop(CA_IX.gm[gm||""]||{});
 if(t3 && codeOk("ca",t3.k) && t3.t>=5)
  return {c:t3.k,conf:Math.min(0.72,0.45+0.27*t3.p),via:"Grupo de Mercadoria "+gm+": "+t3.n+" de "+t3.t+" materiais da base com a classe "+t3.k+" ("+Math.round(t3.p*100)+"%) - evidencia mais ampla, sugerimos confirmar",lista:"coluna CLASSE DE AVALIACAO da base carregada",kind:"aproximada"};
 return {c:"",conf:0,via:"a base tem a coluna CLASSE DE AVALIACAO, mas nenhuma evidencia suficiente para este material - campo marcado como REVISAR",lista:"coluna CLASSE DE AVALIACAO da base carregada",kind:"revisar"};
}

/* ================================================================
   v7.4 - BASES EMBUTIDAS NO PROPRIO ARQUIVO
   Slots preenchidos com o conteudo das planilhas oficiais: quando ha
   linhas embutidas, o sistema abre pronto para categorizar, sem upload.
   O upload do administrador continua sobrescrevendo em memoria/sessao.
   ================================================================ */
const BASE_EMB = {
 ind:{ nome:"Base de Materiais Indiretos", rows:[], meta:"BASEINDIRETO.xlsx - 25.009 materiais - 9 campos, com Classe de Avaliacao - 10/09/2026" },
 imo:{ nome:"Base de Materiais Imobilizados", rows:[], meta:"BASEIMOBILIZADO.xlsx - 1.390 materiais - 8 campos, sem Classe de Avaliacao - 10/09/2026" }
};
let BASE_ATIVA="ind";
function baseEmbHas(k){ return !!(BASE_EMB[k] && BASE_EMB[k].rows && BASE_EMB[k].rows.length) }
function baseEmbTotal(){ return (BASE_EMB.ind.rows.length||0)+(BASE_EMB.imo.rows.length||0) }
function baseEmbRecs(k){
 return (BASE_EMB[k].rows||[]).map(function(o){
  return {d:o.d||"", c:o.c||o.d||"", l:o.l||"", tm:codeSafe("tm",o.tm), gm:codeSafe("gm",o.gm),
          gc:codeSafe("gc",o.gc), gci:codeSafe("gci",o.gci), sa:codeSafe("setor",o.sa),
          cod:String(o.cod||""), ncm:(typeof ncmClean==="function"? ncmClean(o.ncm||"") : (o.ncm||"")),
          ca:caLearn(o.ca||"")};
 });
}
function baseEmbApply(k,quiet){
 if(!baseEmbHas(k)) return false;
 BASE_ATIVA=k; S.base=baseEmbRecs(k); BASE_ORIGEM=BASE_EMB[k].nome+" embutida no arquivo"; buildIndex();
 try{ setBaseState() }catch(e){}
 try{ if(!quiet) toast(BASE_EMB[k].nome+" embutida no arquivo: "+S.base.length.toLocaleString("pt-BR")+" materiais prontos, sem upload") }catch(e){}
 return true;
}
function baseAtivaNome(){ return (BASE_EMB[BASE_ATIVA]||{}).nome || "Base de Materiais Indiretos" }


/* Regra de Imobilizado - IN-CPR-01.001 / CPC 27 */
const IMOB = {
 criterios:["Essencia fisica: bem tangivel mantido para uso na producao, servicos, aluguel ou fins administrativos",
            "Vida util esperada superior a 1 ano (mais de um periodo contabil)",
            "Beneficios economicos futuros provaveis e custo mensuravel com confiabilidade"],
 excecoes:["NOTEBOOK","LAPTOP","CELULAR","SMARTPHONE","TABLET","COLETOR","EQUIPAMENTO LABORATORIAL"],
 excTok:["NOTEBOOK","LAPTOP","CELULAR","SMARTPHONE","TABLET","COLETOR","IPHONE","IPAD","MACBOOK"],
 bens:["EMPILHADEIRA","TRATOR","CAMINHAO","CAMINHONETE","AUTOMOVEL","VEICULO","ONIBUS","CARRETA","REBOQUE","COLHEDORA","RETROESCAVADEIRA","ESCAVADEIRA","MOTONIVELADORA","PA CARREGADEIRA","GERADOR","COMPRESSOR","TORNO","PRENSA","BALANCA","EMPACOTADORA","ESTEIRA","TRANSFORMADOR","SUBESTACAO","ELEVADOR","PONTE ROLANTE","COMPUTADOR","DESKTOP","SERVIDOR","MONITOR","IMPRESSORA","NOBREAK","SWITCH","ROTEADOR","PROJETOR","MESA","CADEIRA","ARMARIO","ESTANTE","GAVETEIRO","SOFA","GELADEIRA","FREEZER","MICROONDAS","BEBEDOURO","AR CONDICIONADO","SPLIT","EMPILHADOR","MOTOR ELETRICO COMPLETO","MAQUINA","EQUIPAMENTO","VEICULAR"],
 peca:["PARAFUSO","PORCA","ARRUELA","ROLAMENTO","RETENTOR","CORREIA","FILTRO","MANGUEIRA","BUCHA","ANEL","EIXO","BICO","VALVULA","SENSOR","CABO","MANCAL","JUNTA","KIT","REPARO","PECA","COMPONENTE","ELEMENTO","LAMPADA","FUSIVEL","DISJUNTOR","CONTATOR","TERMINAL","CONECTOR","ABRACADEIRA","PINO","MOLA","TAMPA","SUPORTE","ENGRENAGEM","PISTAO","CAMISA","VEDACAO","GAXETA","ESCOVA","PASTILHA","LONA","DISCO","PNEU","CAMARA","BOMBA","GRAXA","OLEO","TINTA","PARA","REPOSICAO","CARTUCHO","TONER","BATERIA","CARREGADOR","FONTE","TECLADO","MOUSE","CARCACA","MODULO","PLACA","BARRA","CHAPA","TUBO","PERFIL","CONEXAO","LUVA","COTOVELO","CURVA","REDUCAO","FLANGE","NIPLE","UNIAO"],
 baixoValor:["VENTILADOR","CAFETEIRA","FERRAMENTA MANUAL","CADEIRA","PERIFERICO"]
};

/* dicionario de abreviacoes e sinonimos do vocabulario interno */
const ABREV = {
 "ROLAM":"ROLAMENTO","ROL":"ROLAMENTO","RET":"RETENTOR","PARAF":"PARAFUSO","PRF":"PARAFUSO","PFS":"PARAFUSO","SXT":"SEXTAVADO","SEXT":"SEXTAVADO","SXI":"SEXTAVADO INTERNO",
 "AZ":"AZUL","VD":"VERDE","VM":"VERMELHO","BR":"BRANCO","PT":"PRETO","AM":"AMARELO","CZ":"CINZA","LAR":"LARANJA",
 "LUB":"LUBRIFICANTE","LUBR":"LUBRIFICANTE","LUBRIF":"LUBRIFICANTE","GRX":"GRAXA","OL":"OLEO",
 "MANG":"MANGUEIRA","VALV":"VALVULA","CIL":"CILINDRO","HID":"HIDRAULICO","HIDR":"HIDRAULICO","PNEUM":"PNEUMATICO",
 "ELET":"ELETRICO","ELETR":"ELETRICO","ELTD":"ELETRODUTO","ELETROC":"ELETROCALHA","MEC":"MECANICO","AUT":"AUTOMATICO","AUTOM":"AUTOMATICO",
 "CX":"CAIXA","CAB":"CABO","CONEX":"CONEXAO","ABRAC":"ABRACADEIRA","ARR":"ARRUELA","PRC":"PORCA","ANL":"ANEL",
 "EQPTO":"EQUIPAMENTO","EQUIP":"EQUIPAMENTO","MAQ":"MAQUINA","MOT":"MOTOR","BOMB":"BOMBA","COMPR":"COMPRESSOR",
 "SEG":"SEGURANCA","PROT":"PROTETOR","AURIC":"AURICULAR","OCUL":"OCULOS","MASC":"MASCARA","BOT":"BOTINA","UNIF":"UNIFORME",
 "LIMP":"LIMPEZA","DET":"DETERGENTE","DESINF":"DESINFETANTE","SAB":"SABONETE","PAP":"PAPEL","EXPED":"EXPEDIENTE",
 "IMP":"IMPRESSAO","INFO":"INFORMATICA","NTB":"NOTOBOOK","NOTEB":"NOTEBOOK","CEL":"CELULAR","COMP":"COMPUTADOR",
 "INOX":"INOXIDAVEL","GALV":"GALVANIZADO","PLAST":"PLASTICO","BORR":"BORRACHA","MAD":"MADEIRA","ACR":"ACRILICO","ACRIL":"ACRILICO",
 "TRAT":"TRATOR","CAM":"CAMINHAO","VEIC":"VEICULO","EMPILH":"EMPILHADEIRA","COLH":"COLHEDORA",
 "PC":"PECA","UN":"UNIDADE","PCT":"PACOTE","CJ":"CONJUNTO","CONJ":"CONJUNTO","REP":"REPARO","MONT":"MONTAGEM",
 "P/":"PARA","C/":"COM","S/":"SEM"
};
const HEAD_PART = ["PARAFUSO","PORCA","ARRUELA","ROLAMENTO","RETENTOR","CORREIA","FILTRO","MANGUEIRA","BUCHA","ANEL","EIXO","BICO","VALVULA","SENSOR","MANCAL","JUNTA","ENGRENAGEM","PISTAO","VEDACAO","GAXETA","MOLA","PINO","TERMINAL","CONECTOR","ABRACADEIRA","DISJUNTOR","CONTATOR","RELE","FUSIVEL","BOMBA","CILINDRO","ACOPLAMENTO","POLIA","CORRENTE","ELEMENTO","FLANGE","COTOVELO","CURVA","REDUCAO","NIPLE","TRAVA","GRAXEIRO","CAMISA","VOLANTE","TURBINA","INJETOR","ALTERNADOR","ARRANQUE","EMBREAGEM","AMORTECEDOR","PASTILHA","LONA","ROTOR","ESTATOR","SAPATA"];
const HEAD_NLAG = ["SUCATA","SERVICO","LOCACAO","FRETE","LICENCA","SOFTWARE","CURSO","TREINAMENTO","ASSINATURA","MANUTENCAO","CONSERTO","REFORMA","INSTALACAO","CALIBRACAO","ADESIVAGEM","IMPRESSAO"];
const GM_MANUT = ["90001","90002","90003","90004","90101","90151"];
const STOP = new Set(["DE","DA","DO","DOS","DAS","E","OU","A","O","AS","OS","EM","NO","NA","COM","SEM","PARA","POR","AO","UM","UMA","NCM","COD","REF","TIPO","MATERIAL","MOD","UNIDADE","PECA","PC"]);

/* lexico aprendido da base 302 (24.344 registros) - semente embarcada */
const LEXRAW = "PARAFUSO:1394:E:67:90001:86:MR1:98;2POL:906:H:57:90001:50:MR1:70;SXT:749:E:80:90001:85:MR1:97;FILTRO:720:E:76:90002:43:MR1:89;4POL:718:H:64:90001:44:MR1:70;ROLAMENTO:678:E:94:90002:80:MR1:95;PVC:669:H:76:90004:44:MR1:70;BSP:590:H:51:90001:56:MR1:76;OLEO:552:H:81:90708:70:MR1:72;UNC:540:E:85:90001:86:MR1:94;BUCHA:522:E:82:90001:74:MR1:97;NCM:477:E:81:90001:58:MR1:79;CABO:473:E:58:90001:49:MR1:70;LIS:413:E:50:90002:54:MR1:77;PARA:412:E:55:90001:41:MR1:58;ANEL:400:E:92:90001:63:MR1:96;COLHEDORA:388:E:100:90002:99:MR1:100;NAT:384:H:88:90004:52:MR1:60;PORCA:380:E:62:90001:81:MR1:97;LAMPADA:378:H:68:90001:61:MR1:80;KIT:374:E:78:90001:56:MR1:85;ELET:373:E:54:90001:81:MR1:94;CHAVE:364:H:73:90701:70:MC1:57;NBR:354:E:84:90001:67:MR1:84;TERMINAL:354:E:67:90001:91:MR1:94;90GR:353:H:69:90001:57:MR1:81;CORREIA:349:E:93:90002:75:MR1:95;8POL:348:H:56:90001:45:MR1:69;EXT:332:E:64:90001:65:MR1:87;LUVA:331:H:92:90004:46:MR1:51;RED:319:H:61:90001:32:MR1:77;CAIXA:309:H:83:90004:60:MC1:50;RETENTOR:298:E:98:90001:86:MR1:96;CIL:298:E:80:90001:66:MR1:92;5MM:293:E:55:90001:61:MR1:76;ESF:279:E:92:90002:83:MR1:97;COM:279:E:51:90001:42:MR1:56;ARRUELA:275:E:82:90002:63:MR1:97;SEG:262:H:95:90711:86:MR1:78;VALVULA:256:E:90:90001:70:MR1:94;CANETA:251:H:100:90707:98:MR1:68;CARTUCHO:249:H:96:90707:95:MR1:98;MOTOR:247:E:81:90001:53:MR1:84;ACO:238:E:50:90001:47:MR1:76;TONER:234:H:99:90707:98:MR1:89;TAMPA:234:E:59:90001:37:MR1:75;TUBO:232:H:63:90001:41:MR1:74;BICO:232:E:84:90001:71:MR1:89;1POL:231:H:61:90001:48:MR1:68;LIQ:228:H:98:90705:43:MR1:71;FOFO:227:E:62:90001:69:MR1:87;CHAPA:217:H:55:90004:57:MC1:51;ABRACADEIRA:215:E:69:90001:87:MR1:100;20L:204:H:100:90708:62:MR1:73;TRAT:200:E:80:90002:77:MR1:100;COMPR:198:H:62:90001:78:MR1:95;TINTA:197:H:93:90004:68:MR1:65;BALDE:197:H:99:90705:89:MR1:82;BARRA:194:H:70:90004:34:MC1:44;IND:193:H:68:90001:38:MR1:65;EIXO:193:E:91:90001:68:MR1:93;COD:192:E:82:90001:63:MR1:86;BORR:192:E:74:90002:69:MR1:84;RIG:189:E:87:90002:85:MR1:98;16POL:188:H:55:90001:47:MR1:69;DISJUNTOR:187:E:64:90001:94:MR1:98;TESOURA:184:H:100:90707:83:MC1:79;PLACA:183:H:53:90004:41:MR1:53;UNF:181:E:100:90001:99:MR1:100;CAPACETE:176:H:99:90711:99:MR1:87;NPT:176:H:60:90001:48:MR1:79;PINO:176:E:65:90001:39:MR1:79;MANCAL:176:E:93:90001:89:MR1:97;CONECTOR:176:E:66:90001:73:MR1:94;ROSC:174:H:91:90004:68:MC1:52;AUT:170:E:66:90002:52:MR1:93;SENSOR:169:E:90:90001:54:MR1:91;8MM:169:E:51:90001:62:MR1:80;INT:169:H:65:90004:41:MR1:77;ROSCA:169:E:81:90001:78:MR1:84;BRG:166:E:100:90001:100:MR1:100;CAM:165:H:78:90003:88:MR1:98;VASSOURA:163:H:100:90705:98:MR1:94;MET:161:H:81:90004:55:MR1:72;CURVA:160:H:62:90001:86:MR1:98;25MM:159:H:55:90001:54:MR1:81;10MM:157:E:66:90001:64:MR1:82;SUPORTE:157:E:52:90001:30:MR1:57;BOMBA:157:E:72:90001:38:MR1:80;SABONETE:156:H:100:90707:97:IA1:69;DISCO:156:H:51:90001:76:MR1:94;IMP:155:H:94:90707:74:MR1:90;220V:155:H:58:90001:46:MR1:68;FITA:155:H:94:90001:35:MR1:56;INC:154:H:97:90004:34:MR1:62;2MM:151:H:73:90001:46:MR1:70;PERFIL:148:H:81:90004:84:MC1:53;COMB:148:H:57:90002:33:MR1:67;LUBRIF:147:H:99:90708:98:MR1:98;100:146:H:62:90001:36:MR1:60;LIG:145:H:98:90004:92:MC1:74;SACO:140:H:100:90702:77:AU1:62;INOX:138:E:70:90001:62:MR1:78;PLAST:135:H:93:90004:34:MR1:47;LED:133:H:80:90004:66:MR1:50;MANGUEIRA:130:E:74:90001:47:MR1:85;FUSIVEL:129:H:70:90004:50:MR1:52;CONT:128:E:59:90001:55:MR1:88;6MM:128:H:55:90001:69:MR1:79;PORTA:127:H:53:90004:34:MR1:48;JUNTA:126:E:71:90001:40:MR1:86;DETERGENTE:124:H:100:90705:95:MR1:94;COTOVELO:124:E:56:90001:86:MR1:100;10A:121:H:65:90004:42:MR1:62;BROCA:119:E:57:90001:92:MR1:97;50MM:119:H:61:90001:49:MR1:76;PROTETOR:119:H:89:90711:78:MR1:87;ELEMENTO:119:E:75:90001:56:MR1:95;FLEX:119:H:61:90004:39:MR1:61;SINT:118:H:97:90004:52:MR1:66;ALG:117:H:100:90710:66:UF1:65;AUX:116:E:58:90001:76:MR1:97;LIXO:115:H:100:90702:80:AU1:68;5MM2:114:H:56:90001:58:MR1:64;PRETO:114:H:77:90707:35:MR1:53;TERMOMETRO:113:H:100:90707:73:MR1:58;CONIC:113:E:92:90002:71:MR1:100;PNEU:112:H:53:90003:66:MR1:67;PARAF:111:H:60:90001:53:MR1:70;CAMISA:111:H:86:90710:78:UF1:76;MOD:111:E:63:90001:47:MR1:68;MAD:111:H:96:90004:38:MR1:58;MEDICAM:108:H:100:90707:100:MR1:100;TRATOR:108:E:89:90002:81:MR1:100;12MM:108:H:51:90001:56:MR1:70;EMB:108:H:100:90004:49:MR1:60;UNX:107:H:100:90711:55:EP1:52;CAPA:107:H:67:90001:19:MR1:50;WEG:107:E:88:90001:62:MR1:76;BRANCO:107:H:76:90004:21:MR1:36;HID:106:E:75:90002:42:MR1:89;CORTE:106:E:63:90001:74:MR1:82;ACL:106:E:97:90001:98:MR1:98;500ML:106:H:100:90705:42:MR1:49;UNIAO:105:H:52:90001:61:MR1:81;EST:104:H:90:90701:46:MC1:55;PROT:104:H:97:90711:77:MR1:74;MANG:104:E:80:90001:81:MR1:86;FLANGE:102:E:66:90001:83:MR1:95;JOHN:102:E:61:90001:57:MR1:85;150:102:E:67:90003:45:MR1:80;SXI:102:E:100:90001:100:MR1:100;CAB:101:H:76:90001:51:MR1:86;OCULOS:101:H:100:90711:97:MR1:88;MEDICAMENTO:101:H:100:90707:93:MR1:65;ELAST:100:E:63:90001:65:MR1:80;1KG:100:H:96:90004:29:MR1:60;DEERE:100:E:60:90001:59:MR1:86;4MM:99:H:54:90001:56:MR1:77;690V:99:E:64:90001:91:MR1:99;CONTATOR:98:E:72:90001:93:MR1:97;TIPO:98:E:65:90001:48:MR1:65;VEDA:97:E:81:90001:70:MR1:82;AGULHA:97:H:97:90707:90:MR1:95;MAN:95:E:51:90001:66:MR1:76;PERF:94:H:84:90004:68:MR1:66;ELETRODO:93:E:86:90001:91:MR1:88;SEGURANCA:93:H:92:90711:89:MR1:97;MIN:93:H:89:90708:88:MR1:84;PAGINAS:92:H:100:90707:100:MR1:100;LUB:92:H:79:90708:68:MR1:93;PLUG:91:H:54:90001:52:MR1:93;24V:91:E:55:90001:48:MR1:64;RTG:91:H:81:90004:63:MC1:56;MOL:91:E:77:90001:73:MR1:90;ISOL:90:H:68:90001:67:MR1:67;TOMADA:90:H:92:90004:91:MC1:52;OLH:90:H:64:90001:97:MR1:99;SEXT:89:H:74:90001:79:MR1:90;HSS:89:H:57:90001:78:MR1:84;3POL:88:H:65:90001:47:MR1:76;MOLA:88:E:80:90001:43:MR1:86;BLOCO:87:E:52:90001:55:MR1:71;PRATO:87:H:95:90702:87:MR1:77;4X0:87:H:100:90707:94:MR1:100;FIX:87:E:60:90001:76:MR1:95;VEL:87:E:67:90001:56:MR1:69;CIRC:87:E:63:90001:68:MR1:89;FIO:86:E:60:90001:53:MR1:67;CL8:86:E:100:90001:100:MR1:100;16MM:86:E:60:90001:65:MR1:78;MANUAL:86:E:65:90001:60:MR1:87;MEC:86:H:53:90001:44:MR1:67;PES:86:H:74:90001:26:MR1:48;SOLDA:86:H:66:90711:40:MR1:77;MASCARA:86:H:98:90711:98:MR1:97;PAPEL:85:H:96:90707:41:MC1:52;45GR:85:E:52:90001:85:MR1:98;ANILHA:85:H:91:90004:91:MR1:55;100UN:85:H:99:90702:65:AU1:47;SERINGA:84:H:100:90707:98:MR1:99;LUBR:83:H:99:90708:96:LB1:67;FREIO:83:E:71:90003:64:MR1:73;W22:83:E:93:90001:67:MR1:88;18L:83:H:99:90004:66:MR1:73;160:83:E:83:90001:48:MR1:83;MACHO:82:E:91:90001:83:MR1:88;A105:82:H:57:90001:45:MR1:99;750V:82:E:63:90001:84:MR1:82;CONEXAO:82:E:77:90001:61:MR1:71;250V:81:H:79:90004:53:MR1:54;HELIC:81:E:51:90001:95:MR1:99;1NA:81:E:64:90001:68:MR1:94;CILINDRO:80:E:79:90001:51:MR1:85;TAM:80:H:99:90711:78:MR1:38;ELETROD:80:H:86:90004:46:MR1:65;ACRIL:80:H:98:90004:89:MR1:54;12V:80:H:61:90003:30:MR1:60;AGUA:79:H:58:90001:23:MR1:56;22MM:79:H:56:90001:49:MR1:81;ADAPTADOR:79:H:73:90004:51:MR1:49;TRAVA:79:E:87:90001:51:MR1:90;ELTD:78:H:77:90001:54:MR1:79;ACOPLAMENTO:78:E:53:90004:37:MR1:64;VEDACAO:78:E:95:90002:60:MR1:92;1KV:77:E:58:90001:82:MR1:84;3MM:76:H:71:90001:38:MR1:62;SOQ:76:H:97:90701:87:MC1:92;AZUL:76:H:83:90707:25:MR1:42;AURIC:75:H:100:90711:100:MR1:100;INTERRUPTOR:75:H:53:90004:36:MR1:65;GALV:75:H:57:90004:43:MR1:53;MANGA:74:H:84:90710:68:UF1:66;COND:73:H:60:90001:34:MR1:82;BASE:72:H:57:90001:33:MR1:63;GRAMPO:72:H:69:90004:57:MR1:72;BORRACHA:72:E:78:90001:46:MR1:79;NXR:71:E:79:90003:75:MR1:80;FLG:71:E:73:90001:75:MR1:86;VED:71:E:90:90001:52:MR1:94;LAPIS:70:H:99:90707:97:MC1:90;1MM:70:H:60:90001:53:MR1:83;CLIPES:70:H:99:90707:96:IA1:51;BATERIA:70:E:61:90001:40:MR1:67;TUB:69:E:57:90001:80:MR1:88;220:69:E:52:90001:32:MR1:77;ELETRODUTO:68:E:60:90001:79:MR1:90;20MM:68:H:65:90001:54:MR1:82;CONC:68:H:94:90711:35:MR1:87;ROLO:68:H:75:90004:41:MC1:47;ESCOVA:68:H:74:90701:35:MR1:47;100MM:68:H:72:90001:34:MR1:76;BOTAO:68:H:51:90001:43:MR1:79;ADESIVO:68:H:79:90001:28:MR1:68;VALV:67:E:81:90001:64:MR1:96;MASSA:67:H:97:90004:82:MC1:78;ENC:66:E:61:90001:61:MR1:98;CARBONO:66:E:76:90001:61:MR1:68;COMBUSTIVEL:66:E:79:90002:41:MR1:79;DIR:66:H:53:90001:83:MR1:94;DESINFETANTE:66:H:100:90705:100:MR1:91;ESP:66:H:59:90001:35:MR1:70;CORRENTE:66:E:88:90001:61:MR1:88;SHELL:66:H:100:90708:98:MR1:97;MINI:65:E:54:90001:60:MR1:88;20A:64:H:70:90004:42:MR1:73;PISO:63:H:87:90004:75:MR1:52;GUARDANAPO:63:H:100:90702:100:MR1:89;TELA:63:H:68:90004:49:MC1:43;CARREG:62:E:60:90002:52:MR1:100;FERRO:62:E:53:90001:34:MR1:60;5KG:62:E:50:90001:52:MR1:79;32MM:62:H:71:90001:37:MR1:52;AUTO:61:H:70:90003:36:MR1:84;40MM:61:H:61:90001:51:MR1:85;RELE:61:E:67:90001:51:MR1:80;DESC:61:H:90:90707:34:MR1:79;EMEN:61:H:61:90004:61:MC1:52;84879000:61:E:100:90001:100:MR1:100;500:61:H:84:90004:62:MR1:79;RODA:61:E:75:90001:34:MR1:82;PRESSAO:60:E:92:90001:57:MR1:88;ABS:60:H:90:90711:37:MR1:72;GRAXEIRO:60:E:100:90001:100:MR1:100;4X1:60:H:50:90001:55:MR1:75;150LBS:60:H:57:90001:82:MR1:92;AUTOM:59:H:83:90003:56:MR1:97;PRESS:59:E:76:90001:69:MR1:85;CONJUNTO:59:E:66:90001:41:MR1:56;7MM:59:H:63:90001:44:MR1:69;200:59:E:68:90001:47:MR1:71;PARTIDA:59:E:88:90003:42:MR1:88;FENDA:59:H:73:90701:61:MR1:63;CAMINHAO:58:H:71:90003:86:MR1:98;TOM:58:H:91:90004:86:MC1:55;PEAD:58:H:71:90711:40:MR1:69;EMENDA:58:H:50:90001:59:MR1:90;32A:58:E:55:90001:69:MR1:81;BOTINA:58:H:100:90711:100:EP1:98;150MM:58:H:81:90004:26:MR1:64;COR:57:H:56:90001:33:MR1:47;CONCHA:57:H:100:90711:93:MR1:95;SELO:57:E:74:90001:60:MR1:75;ELETROC:56:H:93:90004:73:MR1:88;MULTIUSO:56:H:100:90705:59:MR1:50;MULT:56:H:91:90004:41:MR1:82;6POL:56:H:71:90001:43:MR1:68;HIDR:56:H:57:90707:36:MR1:68;POL:56:H:54:90001:39:MR1:80;DIG:55:H:84:90707:58:MR1:73;500G:55:H:93:90706:24:MR1:40;CONCR:55:H:71:90001:73:MR1:85;CALCA:55:H:98:90710:91:UF1:82;COPO:54:H:89:90701:30:MC1:48;5A2:54:H:87:90004:69:MC1:69;16A:54:H:57:90001:52:MR1:72;200MM:54:H:85:90702:22:MR1:83;MODULO:53:H:64:90001:34:MR1:79;COBRE:53:E:85:90001:58:MR1:64;L70F:53:E:55:90002:51:MR1:98;COLA:53:H:85:90001:55:MR1:55;ASTM:53:E:85:90001:66:MR1:77;300MM:53:H:89:90705:43:MR1:70;GAS:52:H:60:90001:37:MR1:65;9MM:52:H:69:90001:42:MR1:67;400MM:52:H:81:90702:38:MR1:85;ACV:52:H:94:90701:88:MC1:81;METAL:52:E:52:90001:75:MR1:87;INSULOK:52:E:52:90001:83:MR1:100;VIDRO:52:H:65:90004:19:MR1:50;MATERIAL:52:E:73:90001:60:MR1:62;LIMPADOR:52:H:65:90705:56:ML1:33;6000MM:52:H:71:90004:62:MR1:56;PRETA:52:H:92:90707:56:MR1:38;IMPRESSORA:51:H:100:90707:98:MR1:71;ROL:51:E:51:90001:33:MR1:86;18MM:51:H:71:90001:76:MR1:71;TRAV:51:E:78:90001:82:MR1:92;CPVC:51:H:53:90001:69:MR1:90;RSF:51:E:84:90001:98:MR1:100";
const LEX = {};
LEXRAW.split(";").forEach(function(r){
 const p = r.split(":"); if(p.length<8) return;
 const tm = p[2]==="E"?"ERSA":(p[2]==="H"?"HIBE":"NLAG");
 LEX[p[0]] = {n:+p[1], tm:tm, ptm:+p[3]/100, gm:p[4], pgm:+p[5]/100, gc:p[6], pgc:+p[7]/100};
});

/* ================= UTIL ================= */
const $ = function(s){return document.querySelector(s)};
const $$ = function(s){return Array.prototype.slice.call(document.querySelectorAll(s))};
function toast(m){const t=$("#toast");t.textContent=m;t.classList.add("on");setTimeout(function(){t.classList.remove("on")},2200)}
function esc(s){return String(s==null?"":s).replace(/[&<>"]/g,function(c){return {"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c]})}
function deacc(s){return String(s||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"")}
function norm(s){
 let t = deacc(String(s||"")).toUpperCase();
 t = t.replace(/([0-9])[,.]([0-9])/g,"$1.$2");
 t = t.replace(/["']/g," POL ").replace(/POLEGADAS?/g," POL ");
 t = t.replace(/[^A-Z0-9./]+/g," ").replace(/\s+/g," ").trim();
 return t;
}
function expand(t){ return ABREV[t]||t; }
function toks(s){
 const out=[]; const base=norm(s).split(" ");
 for(let i=0;i<base.length;i++){
  let w=base[i]; if(!w) continue;
  w.split(" ").forEach(function(x){
   if(!x||STOP.has(x)) return;
   if(x.length<2) return;
   if(x.length>4 && /S$/.test(x) && !/SS$/.test(x)) x=x.slice(0,-1); /* singular */
   out.push(x);
  });
 }
 return out;
}
function tri(s){const t=" "+norm(s).replace(/\s+/g," ")+" ";const S=new Set();for(let i=0;i<t.length-2;i++)S.add(t.substr(i,3));return S}
function diceSet(A,B){if(!A.size||!B.size)return 0;let c=0;A.forEach(function(x){if(B.has(x))c++});return 2*c/(A.size+B.size)}
function lev(a,b){
 a=norm(a);b=norm(b);
 if(a===b) return 1;
 const m=a.length,n=b.length; if(!m||!n) return 0;
 if(Math.abs(m-n)>Math.max(m,n)*0.6) return 0;
 let prev=new Array(n+1),cur=new Array(n+1);
 for(let j=0;j<=n;j++)prev[j]=j;
 for(let i=1;i<=m;i++){cur[0]=i;
  for(let j=1;j<=n;j++){cur[j]=Math.min(prev[j]+1,cur[j-1]+1,prev[j-1]+(a[i-1]===b[j-1]?0:1))}
  const t=prev;prev=cur;cur=t;
 }
 return 1-prev[n]/Math.max(m,n);
}

/* ================= INDICE E APRENDIZADO ================= */
const S = { ex:[], ncmIx:null, ncmIxN:-1, bm:null, bmN:-1, base:[], idx:new Map(), df:new Map(), pref:new Map(), lexBase:new Map(), headBase:new Map(), fb:[], log:[], results:[], pwd:"cadastro302" };

function buildIndex(){
 S.idx=new Map(); S.df=new Map(); S.lexBase=new Map(); S.pref=new Map();
 S.headBase=new Map();
 for(let i=0;i<S.base.length;i++){
  const r=S.base[i];
  r._t = toks((r.d||"")+" "+(r.c||""));
  const uniq=new Set(r._t);
  const hd0=r._t[0]||"";
  if(hd0 && (r.tm||r.gm||r.gc)){
   let he=S.headBase.get(hd0); if(!he){he={n:0,tm:{},gm:{},gc:{}};S.headBase.set(hd0,he)}
   he.n++;
   if(r.tm) he.tm[r.tm]=(he.tm[r.tm]||0)+1;
   if(r.gm) he.gm[r.gm]=(he.gm[r.gm]||0)+1;
   if(r.gc) he.gc[r.gc]=(he.gc[r.gc]||0)+1;
  }
  uniq.forEach(function(t){
   if(!S.idx.has(t)) S.idx.set(t,[]);
   S.idx.get(t).push(i);
   S.df.set(t,(S.df.get(t)||0)+1);
   const p4=t.substr(0,4);
   if(!S.pref.has(p4)) S.pref.set(p4,[]);
   if(S.pref.get(p4).indexOf(t)<0) S.pref.get(p4).push(t);
   if(r.tm||r.gm||r.gc){
    let e=S.lexBase.get(t); if(!e){e={n:0,tm:{},gm:{},gc:{}};S.lexBase.set(t,e)}
    e.n++;
    if(r.tm) e.tm[r.tm]=(e.tm[r.tm]||0)+1;
    if(r.gm) e.gm[r.gm]=(e.gm[r.gm]||0)+1;
    if(r.gc) e.gc[r.gc]=(e.gc[r.gc]||0)+1;
   }
  });
 }
 try{ caBuild() }catch(e){}
}
function idf(t){
 const N = S.base.length || 24344;
 const d = S.df.get(t) || (LEX[t]? LEX[t].n : 1);
 return Math.log(1 + N/(1+d));
}
function lexLook(tk){
 const lb=S.lexBase.get(tk); if(lb&&lb.n>=3) return {src:"base",e:lb};
 if(LEX[tk]) return {src:"seed",e:LEX[tk]};
 if(tk.length>=3){
  const pf=S.pref.get(tk.substr(0,Math.min(4,tk.length)));
  if(pf){ let bt=null,bn=0;
   pf.forEach(function(t2){ const x=S.lexBase.get(t2); if(x&&x.n>bn&&(t2.indexOf(tk)===0||tk.indexOf(t2)===0)){bn=x.n;bt=x} });
   if(bt&&bn>=3) return {src:"base",e:bt,soft:true};
  }
  let sk=null,sn=0;
  for(const k in LEX){ if((k.indexOf(tk)===0||tk.indexOf(k)===0)&&Math.min(k.length,tk.length)>=3&&LEX[k].n>sn){sn=LEX[k].n;sk=k} }
  if(sk) return {src:"seed",e:LEX[sk],soft:true};
 }
 return null;
}
function best(o){let k=null,v=-1,tot=0;for(const x in o){tot+=o[x];if(o[x]>v){v=o[x];k=x}}return {k:k,p:tot?v/tot:0}}

/* candidatos por indice invertido */
function candidates(qt){
 const score=new Map();
 const uq=Array.from(new Set(qt));
 function add(list,w){ if(!list) return; if(list.length > Math.max(600, S.base.length*0.30)) return;
  for(let i=0;i<list.length;i++) score.set(list[i],(score.get(list[i])||0)+w) }
 uq.forEach(function(t){
  const w=idf(t);
  add(S.idx.get(t), w);
  const pf=S.pref.get(t.substr(0,4));
  if(pf) pf.forEach(function(t2){ if(t2!==t) add(S.idx.get(t2), w*0.55) });
 });
 const arr=Array.from(score.entries()).sort(function(a,b){return b[1]-a[1]}).slice(0,260);
 return arr.map(function(x){return x[0]});
}
/* casamento tolerante a abreviacao, prefixo e erro de digitacao */
function tokMatch(a,b){
 if(a===b) return 1;
 const m=Math.min(a.length,b.length);
 if(m>=2 && (a.substr(0,m)===b.substr(0,m))) return m>=4?0.9:0.84;
 if(m>=4){ const l=lev(a,b); if(l>=0.8) return 0.75; if(l>=0.68) return 0.55 }
 return 0;
}
function cov(A,B){
 let num=0,den=0;
 for(let i=0;i<A.length;i++){
  const w=idf(A[i]); den+=w; let bst=0;
  for(let j=0;j<B.length;j++){ const s=tokMatch(A[i],B[j]); if(s>bst){bst=s; if(bst===1)break} }
  num+=w*bst;
 }
 return den? num/den : 0;
}
function simScore(qt,qs,r){
 const A=Array.from(new Set(qt)), B=Array.from(new Set(r._t||[]));
 if(!A.length||!B.length) return 0;
 const cq=cov(A,B), cr=cov(B,A);
 const tr=diceSet(qs, r._tri||(r._tri=tri((r.d||"")+" "+(r.c||""))));
 const mx=Math.max(cq,cr), mn=Math.min(cq,cr);
 let sc = 0.52*mx + 0.24*mn + 0.24*tr;
 /* contencao: descricao abreviada contida na existente (CANETA AZ vs CANETA AZUL) */
 if(cq>=0.86 || cr>=0.86) sc = Math.max(sc, 0.76 + 0.16*tr);
 if(A.length<=3 && B.length<=3) sc = Math.max(sc, 0.6*sc + 0.4*lev(A.join(" "),B.join(" ")));
 return Math.max(0,Math.min(1,sc));
}
function neighbors(desc, extra, k){
 if(!S.base.length) return [];
 const qt = toks(desc+" "+(extra||""));
 const qs = tri(desc+" "+(extra||""));
 const cand = candidates(qt);
 const out=[];
 for(let i=0;i<cand.length;i++){ const r=S.base[cand[i]]; out.push({r:r, s:simScore(qt,qs,r)}) }
 out.sort(function(a,b){return b.s-a.s});
 return out.slice(0,k||18);
}

/* ================= PRIOR DO SUBSTANTIVO PRINCIPAL (real, base 302 - 24.344 registros) ================= */
const HEADRAW = "ABRACADEIRA:169:E:78:90001:85:MR1:98;ACABAMENTO:8:H:100:90004:88:MR1:75;ACENDEDOR:2:H:100:90701:50:MC1:50;ACESSORIOS:15:H:93:90004:73:MR1:60;ACIONADOR:6:H:50:90004:33:MR1:33;ACOPLADOR:3:E:100:90001:100:MR1:100;ACOPLAMENTO:90:E:62:90001:49:MR1:70;ACUMULADOR:8:E:100:90002:62:MR1:100;ADAM4PDM:4:H:100:90003:50:MR1:50;ADAPTADOR:52:H:69:90004:48:MR1:46;ADESIVO:83:H:86:90001:48:MR1:69;ADITIVO:33:H:100:90151:33:MR1:52;AGENDA:3:H:100:90707:100:ME2:67;AGUA:5:H:80:90705:60:MC1:40;AGULHA:80:H:99:90707:96:MR1:99;AGULHAO:3:H:100:90707:100:MR1:100;ALARGADOR:2:H:100:90003:100:MR1:100;ALAVANCA:13:E:92:90001:46:MR1:85;ALCA:6:E:67:90001:100:MR1:100;ALCOOL:39:H:97:90707:79:MR1:82;ALGODAO:27:H:100:90707:100:MR1:96;ALICATE:34:H:85:90701:85:MR1:62;ALIMENTOS:98:H:77:90703:94:AL1:93;ALONGADOR:2:H:100:90004:100:MR1:100;ALTERNADOR:10:E:100:90002:90:MR1:80;AMACIANTE:3:H:100:90705:100:ML1:67;AMORTECEDOR:20:E:75:90003:45:MR1:75;AMOSTRA:2:H:100:90706:100:LB2:100;AMPERIMETRO:2:E:100:90001:50:MR1:100;ANALISADOR:3:H:67:90003:33:MR1:100;ANEL:378:E:96:90001:58:MR1:96;ANILHA:42:H:95:90004:95:MC1:88;ANTENA:4:E:50:90001:50:MR1:100;ANTI:3:H:100:90002:67:MR1:100;ANTIFERRUGEM:6:H:100:90001:83:MR1:100;APARELHO:4:H:100:14001:100:MC1:75;APLICADOR:3:H:100:90001:67:MR1:67;APOIO:9:E:56:90002:56:ME2:44;APONTADOR:2:H:100:90707:100:ME2:100;ARAME:31:H:94:90004:90:MR1:61;ARCO:2:E:50:90003:50:MR1:50;AREIA:2:H:100:90151:100:MC1:100;ARGAMASSA:14:H:100:90151:93:MR1:50;ARO:6:E:100:90001:33:MR1:100;AROMA:2:H:100:90703:100:AL1:100;AROMATIZADOR/DIFUSOR:2:H:100:90705:50:ML1:50;ARRUELA:271:E:82:90002:65:MR1:97;ARTICULACAO:8:E:100:90001:75:MR1:100;ASA:2:E:100:90002:100:MR1:100;ASPERSOR:3:E:100:90001:100:MR1:100;ASSENTO:5:H:80:90004:80:MC1:80;ASSOALHO:8:E:100:90002:88:MR1:100;ATADURA:39:H:100:90707:97:MR1:97;ATUADOR:10:E:60:90001:50:MR1:90;AUTOMATICO:8:H:88:90003:100:MR1:88;AUTOTRAFO:2:E:100:90001:50:MR1:50;AVENTAL:4:H:100:90711:75:EP1:50;BACIA:8:H:100:90004:62:MC1:50;BALANCA:4:H:50:90701:25:MR1:75;BALANCIM:2:E:100:90001:100:MR1:100;BALAO:13:H:100:90706:77:LB2:77;BALDE:55:H:100:90705:95:MR1:75;BANCO:3:E:100:90001:67:MR1:67;BANDA:3:E:100:90002:67:MR1:100;BANDEIRA:3:H:100:90707:100:ME2:100;BANDEJA:5:H:60:90004:60:MR1:60;BANNER:3:H:100:90707:100:ME2:67;BARBANTE:2:H:100:90707:100:ME2:100;BARRA:155:H:76:90004:44:MC1:52;BARRAMENTO:18:H:50:90001:44:MR1:89;BARREIRA:4:E:100:90002:100:MR1:100;BASE:32:H:78:90004:41:MR1:56;BASTAO/VARA:2:H:100:90701:100:MC1:100;BATENTE:6:E:83:90003:50:MR1:67;BATERIA:61:H:51:90707:38:MR1:62;BEBIDA:11:H:82:90703:100:AL1:91;BECKER:4:H:100:90706:100:LB2:100;BICO:172:E:96:90001:82:MR1:97;BIELA:3:E:100:90001:67:MR1:100;BIELETA:5:E:80:90003:40:MR1:80;BITS:4:H:100:90701:50:MC1:50;BLOCO:67:E:61:90001:69:MR1:87;BLOQUEIO:4:E:75:90001:75:MR1:100;BOBINA:28:E:64:90001:50:MR1:71;BOCAL:9:E:100:90001:89:MR1:100;BOIA:7:H:71:90001:43:MR1:71;BOLSA:3:H:100:90701:100:MC1:67;BOMBA:68:E:90:90001:44:MR1:84;BOMBONA:3:H:100:24001:100:ME2:100;BORNE:42:E:64:90001:64:MR1:90;BORRACHA:25:E:80:90001:36:MR1:76;BOTAO:58:E:52:90004:48:MR1:86;BOTOEIRA:5:E:80:90001:80:MR1:80;BRACADEIRA:6:E:100:90002:100:MR1:100;BRACO:14:E:100:90002:64:MR1:93;BRIDA:2:E:100:90001:100:MR1:100;BRINDES:2:H:100:90707:100:MC1:50;BRITA:14:H:93:90151:64:MC1:43;BROCA:117:E:57:90001:93:MR1:97;BRONZINA:3:E:100:90001:67:MR1:67;BROXA:8:H:100:90004:62:MR1:62;BUCHA:422:E:86:90001:74:MR1:95;BUJAO:35:E:74:90001:71:MR1:100;BURRINHO:2:H:100:90003:100:MR1:100;BUZINA:6:H:50:90003:50:MR1:100;CABECOTE:17:E:71:90002:53:MR1:88;CABO:278:E:74:90001:64:MR1:76;CABRA:2:H:100:90701:100:MR1:100;CACAMBA:6:E:67:90001:33:MR1:67;CADARCO:3:E:100:90001:100:MR1:100;CADEADO:16:H:100:90701:75:MC1:62;CADERNO:7:H:100:90707:100:ME2:100;CAIBRO:8:H:100:90004:88:MC1:62;CAIXA:234:H:83:90004:72:MC1:57;CAL:5:H:80:90004:60:MC1:60;CALCA:3:H:100:90711:67:EP1:67;CALCADO:74:H:100:90711:99:EP1:99;CALCO:38:E:97:90001:42:MR1:87;CALCULADORA:3:H:67:90707:67:ME2:67;CALHA:2:E:100:90001:100:MR1:100;CAMARA:21:E:52:90003:81:MR1:95;CAMERA:7:H:71:90709:71:MR1:86;CAMISA:8:E:88:90001:62:MR1:75;CAMISETA:9:H:100:90711:89:EP1:89;CANALETA:11:H:73:90101:45:MC1:64;CANALIZACAO:4:H:50:90004:50:MR1:75;CANECA:27:E:74:90002:67:MR1:70;CANETA:248:H:100:90707:99:MR1:69;CANTONEIRA:5:H:80:90004:100:MR1:60;CAPA:72:H:54:90002:26:MR1:64;CAPACETE:167:H:100:90711:100:MR1:91;CAPACITOR:43:E:100:90101:91:MR1:100;CAPCTETE:2:H:100:90711:100:MR1:100;CAPO:2:E:100:90002:100:MR1:100;CAPOTA:6:E:83:90002:50:MR1:50;CARBURADOR:3:E:100:90002:67:MQ1:67;CARCACA:9:E:100:90002:56:MR1:78;CARNEIRA:4:H:100:90711:100:EP1:100;CARREGADOR:12:E:50:90001:50:MR1:33;CARRETEL:2:H:50:90001:100:MR1:100;CARRINHO:5:H:80:90701:40:MC1:80;CARTAO:5:H:80:90707:80:ME2:40;CARTOLA:2:E:100:90001:100:MR1:100;CARTUCHO:253:H:97:90707:96:MR1:97;CARV:2:E:50:90001:50:MR1:50;CASQUILHO:3:E:100:90002:33:MR1:100;CASTANHA:3:E:100:90001:67:MR1:100;CATALISADOR:21:H:90:90001:67:MR1:67;CATRACA:4:E:100:90001:75:MR1:100;CAVADEIRA:6:H:100:90701:100:MC1:83;CAVALETE:4:H:50:90001:25:ME2:50;CELULA:9:E:100:90001:78:MR1:78;CESTO:4:E:50:90707:50:MR1:50;CHAPA:187:H:62:90004:65:MC1:54;CHAPEU:4:H:100:90711:50:ME2:50;CHAVE:266:H:67:90701:65:MC1:48;CHAVEIRO:2:H:100:90707:100:MC1:100;CHAVETA:33:H:91:90004:82:MC1:91;CHICOTE:23:E:74:90002:57:MR1:83;CHIP:5:H:100:90709:80:MC1:100;CHUMBADOR:30:H:90:90151:57:MC1:80;CHUVEIRO:5:H:100:90004:80:MR1:80;CILINDRO:55:E:76:90001:53:MR1:89;CIMENTO:3:H:100:90004:67:MC1:67;CINTA:31:E:61:90001:61:MR1:74;CINTO:4:H:100:90711:100:EP1:100;CINTURAO:4:H:100:90711:75:EP1:25;CLAVICULARIO/ARMARIO:2:H:100:90707:50:ME2:50;CLIP:3:H:100:90707:33:ME2:33;CLIPE:23:H:100:90707:83:MC1:65;CLIPES:61:H:100:90707:98:MC1:52;CNV:10:H:90:90004:80:MC1:80;COBERTURA:3:E:100:90002:100:MR1:67;COIFA:2:E:100:90003:100:AU1:100;COLA:4:H:100:90707:100:ME2:100;COLAR:40:E:60:90151:58:MR1:95;COLETE:6:H:100:90711:100:MC1:50;COLHER:8:H:100:90004:75:MR1:75;COLUNA:14:H:71:90151:36:MR1:50;COMBUSTIVEL:2:H:100:29002:100:CO2:100;COMPRESSA:3:H:100:90707:100:MC1:33;COMPRESSOR:17:E:59:90003:47:MR1:94;COMUTADOR:2:E:100:90002:100:MR1:100;CONCRETO:2:H:100:90004:50:MC1:100;CONDENSADOR:4:E:100:90002:75:MR1:75;CONDUTOR:4:E:50:90003:50:MR1:100;CONE:10:E:80:90002:60:MR1:60;CONECTOR:198:E:68:90001:80:MR1:95;CONECTOR/PLUG:12:H:50:90001:42:MR1:58;CONECTOR/PLUGUE:39:E:79:90001:90:MR1:100;CONEXAO:67:E:76:90001:57:MR1:64;CONTADOR:4:E:50:90001:50:MR1:100;CONTATO:7:E:100:90002:100:MR1:100;CONTATOR:98:E:72:90001:93:MR1:97;CONTRA:14:E:100:90001:64:MR1:93;CONTROLADOR:23:E:78:90001:78:MR1:87;CONTROLE:6:H:83:90707:33:IS1:33;CONVERSOR:13:E:54:90001:46:MR1:54;COPO:3:E:67:90002:33:MR1:67;COPO/TACA/XICARA:3:H:100:90707:100:MU1:33;CORANTE:3:H:67:90004:67:ME2:33;CORDA:13:H:62:90701:54:MR1:62;CORDAO:6:H:83:90001:33:ME2:50;CORDOALHA:5:E:100:90001:100:MR1:100;COROA:5:E:60:90002:60:MR1:60;CORPO:4:E:100:90001:100:MR1:100;CORREIA:132:E:87:90002:52:MR1:86;CORRENTE:40:E:85:90001:65:MR1:92;CORTINA/PERSIANA:2:H:100:90004:50:MC1:50;COSSINETE:6:E:100:90001:100:MR1:100;COTOVELO:32:H:78:90001:91:MR1:100;COXIM:15:E:73:90002:40:MR1:87;CREMALHEIRA:5:E:60:90001:40:MR1:100;CREME:3:H:100:90707:100:ME2:67;CRUZETA:3:E:67:90003:67:MR1:67;CUBA:4:H:100:90004:100:MR1:100;CUBETA:2:H:100:90706:100:LB2:100;CUBO:22:E:82:90001:36:MR1:86;CUICA:2:E:100:90003:50:AU1:50;CURATIVO:3:H:100:90707:100:MC1:67;CURVA:133:H:63:90001:88:MR1:99;DAMPER:2:E:100:90001:100:MR1:100;DECALQUE:5:E:100:90002:100:MR1:100;DEFLETOR:6:E:83:90002:67:MR1:100;DENSIMENTRO:2:H:100:90706:100:LB2:100;DENSIMETRO:2:H:100:90706:100:LB2:100;DESANDADOR:4:H:75:90701:75:MC1:50;DESCARBONIZANTE:2:H:50:90003:100:AU1:50;DESCARGA:2:E:100:90001:100:MR1:100;DESEMPENADEIRA:9:H:89:90001:67:MR1:67;DESENGRAXANTE:11:H:100:90705:55:MR1:55;DESINCRUSTANTE:4:H:100:90705:100:ML1:50;DESINFETANTE:69:H:100:90705:99:MR1:87;DESODORANTE:8:H:100:90705:88:ML1:62;DESODORIZADOR:16:H:94:90705:94:ML1:62;DETECTOR:8:H:50:90001:62:MR1:75;DETERGENTE:163:H:100:90705:100:MR1:98;DIAFRAGMA:2:E:100:90001:50:MR1:50;DILUENTE:12:H:100:90001:50:MC1:83;DIODO:4:E:100:90001:100:MR1:100;DISCO:137:H:55:90001:74:MR1:92;DISJUNTOR:185:E:64:90001:94:MR1:97;DISPENSER:3:H:100:90004:33:MC1:33;DISPLAY:2:H:50:90709:50:SH1:50;DISTRIBUIDOR:3:E:100:90001:100:MR1:100;DOBRADICA:28:H:68:90004:68:MC1:46;DOSADOR:3:H:100:90001:67:MR1:67;DUTO:5:E:60:90001:40:MR1:60;EIXO:70:E:81:90001:44:MR1:89;ELASTICO:5:E:60:90002:20:MR1:80;ELETROCALHA:10:H:100:90004:80:MR1:60;ELETRODO:85:E:85:90001:91:MR1:87;ELETRODUTO:43:E:60:90001:72:MR1:79;ELO:19:E:100:90001:63:MR1:100;EMBALAGEM:4:H:75:90001:50:MR1:50;EMBORRACHAMENTO:2:E:100:90001:100:MR1:100;EMBREAGEM:5:E:80:90002:40:MR1:100;EMENDA:34:E:74:90001:94:MR1:94;EMISSOR:4:E:100:90002:100:MR1:100;EMPUNHADURA:2:H:100:90701:50:MC1:100;ENGATE:18:E:89:90001:67:MR1:89;ENGRENAGEM:49:E:90:90002:45:MR1:94;ENVELOPE:5:H:100:90707:80:ME2:60;ENXADA:5:H:100:90701:100:MR1:100;EPOXI:6:H:100:90002:33:MC1:67;ESCADA:9:H:78:90701:78:MR1:44;ESCAPADOR:2:E:100:90002:100:MR1:100;ESCORA:3:E:67:90001:67:MR1:67;ESCOVA:46:H:91:90701:43:MC1:57;ESFIGMOMANOMETRO:2:H:100:90707:100:ME2:50;ESGUICHO:6:H:83:90004:67:MR1:67;ESMERILHADEIRA:5:H:100:90701:100:MR1:100;ESPACADOR:13:E:85:90002:62:MR1:85;ESPAGUETE:9:H:89:90004:89:MC1:67;ESPARADRAPO:15:H:100:90707:100:MR1:93;ESPATULA:7:H:100:90004:43:MR1:86;ESPELHO:5:E:60:90003:60:AU1:40;ESPONJA:4:H:100:90705:75:ML1:75;ESPUMA:9:H:89:90004:44:MC1:44;ESQUADRO:6:H:100:90701:100:MR1:50;ESTATOR:9:E:100:90001:89:MR1:89;ESTEIRA:4:E:100:90002:75:MR1:100;ESTICADOR:21:H:71:90004:62:MC1:62;ESTILETE:5:H:100:90707:100:ME2:60;ESTOPA:4:H:100:90705:100:MC1:100;ETIQUETA:14:H:64:90001:43:MR1:64;EVAPORADOR:2:E:100:90001:50:MR1:100;EXAUSTOR:7:E:100:90001:100:MR1:100;EXPOSITOR/DISPLAY:2:H:100:90707:100:MC1:100;EXTENSAO:3:H:100:90101:33:MR1:67;EXTENSOR:8:H:100:90001:25:MR1:50;EXTINTOR:13:H:100:90004:92:MC1:77;EXTRATOR:3:E:67:90002:33:MR1:67;FACA:7:E:86:90002:71:MR1:86;FACAO:3:H:100:90701:67:MR1:33;FAROL:23:E:70:90003:48:MR1:83;FECHADURA:20:H:85:90004:60:MR1:70;FECHO:4:E:75:90001:50:MR1:75;FERRAMENTAS:5:H:100:90701:100:MC1:100;FERRO:2:H:100:90002:50:MR1:50;FERTILIZANTE:2:H:100:90707:100:MC1:100;FILT:10:E:80:90001:40:MR1:80;FILTRANTE:38:E:66:90001:50:MR1:100;FILTRO:626:E:79:90002:45:MR1:91;FIO:7:H:57:90002:29:MR1:43;FITA:69:H:97:90001:38:MR1:39;FIXACAO:4:E:100:90002:75:MR1:100;FIXADOR:8:E:62:90001:38:MR1:75;FLANELA:4:H:100:90705:75:ML1:100;FLANGE:71:E:65:90001:82:MR1:97;FLEXIVEL:2:E:100:90003:100:AU1:100;FLORES:4:H:75:90004:75:MC1:75;FLUIDO:9:H:100:90001:44:MR1:78;FLUXO:3:H:100:90001:100:MR1:67;FONE:7:H:100:90709:57:IS1:71;FONTE:28:E:61:90001:86:MR1:86;FORMULARIO:8:H:100:90707:100:ME2:62;FORQUILHA:3:E:100:90001:33:MR1:100;FORRO:8:H:100:90004:88:MC1:62;FOTOCONDUTOR:3:H:67:90707:67:IS1:33;FRASCO:7:H:100:90706:71:LB2:71;FREIO:3:E:67:90001:67:MR1:100;FRUTA:3:H:67:90703:100:AL1:67;FUNDO:2:E:50:90101:50:MR1:50;FUNIL:3:H:100:90701:100:MR1:100;FURADEIRA:4:H:100:90701:100:MR1:100;FUSIVEL:116:H:74:90004:55:MC1:50;FUSO:2:E:100:90001:50:MR1:100;GANCHO:6:E:83:90001:67:MR1:67;GARFO:7:E:100:90002:71:MR1:57;GARRA:5:E:100:90001:80:MR1:80;GAS:10:H:90:90001:50:MR1:60;GAVETA:2:E:100:90001:50:MR1:100;GAXETA:25:E:84:90001:96:MR1:100;GAZE:10:H:100:90707:100:MR1:100;GEL:2:H:100:90705:100:ML1:100;GERADOR:2:E:100:90001:100:MR1:100;GESSO:4:H:100:90004:50:MR1:75;GRADE:11:E:82:90001:45:MR1:55;GRAFITE:4:H:100:90708:50:IA1:50;GRAMPEADOR:5:H:80:90707:100:ME2:80;GRAMPO:68:H:72:90004:59:MR1:68;GRAXA:35:H:97:90708:91:LB1:74;GRAXEIRA:7:E:100:90001:71:MR1:100;GRELHA:8:E:62:90001:50:MR1:88;GUARDA:4:E:75:90003:50:MR1:75;GUARNICAO:10:E:70:90001:30:MR1:80;GUIA:19:E:95:90001:58:MR1:84;HASTE:19:E:63:90001:47:MR1:79;HELICE:4:E:100:90001:75:MR1:75;HERBICIDA:6:H:100:90705:50:IF1:50;HIDROMETRO:2:E:100:90001:100:MR1:100;HORIMETRO:2:E:100:90001:50:MR1:50;HUB:2:H:100:90709:100:IS1:100;IDENTIFICADOR:3:H:100:90001:100:MR1:100;IMPERMEABILIZANTE:16:H:100:90004:56:MC1:75;IMPULSOR:8:E:75:90002:62:MR1:100;IND:3:E:67:90001:67:MR1:100;INDICADOR:4:E:75:90001:50:MR1:100;INDUZIDO:3:H:67:90003:67:MR1:100;INSERTO:10:E:90:90001:70:MR1:100;INSETICIDA:12:H:100:90705:67:MC1:42;INTERFACE:4:E:100:90001:100:MR1:100;INTERRUPTOR:64:H:55:90004:38:MR1:61;INTERRUPTOR/TOMADA:8:H:75:90004:100:MR1:62;INVERSOR:12:E:100:90001:92:MR1:75;ISOLADOR:24:E:83:90001:58:MR1:92;ISOLAMENTO:5:E:80:90002:80:MR1:80;ISOLANTE:5:E:80:90001:80:MR1:80;JALECO:20:H:95:90711:100:EP1:90;JANELA:7:H:71:90004:43:MC1:43;JOGO:6:E:50:90002:33:MR1:33;JOGO/CONJUNTO:143:E:78:90001:60:MR1:73;JOGO/CONJUNTO/KIT:60:H:68:90701:63:MC1:57;JUGULAR:4:H:75:90711:75:EP1:75;JUMPER:2:E:100:90001:100:MR1:100;JUNCAO:4:E:50:90001:50:MR1:50;JUNTA:110:E:75:90001:34:MR1:81;KIT:102:E:61:90003:41:MR1:90;LACRE:8:H:100:90707:62:MC1:62;LAMINA:21:H:67:90003:43:MR1:81;LAMPADA:375:H:69:90001:61:MR1:79;LANTERNA:11:E:64:90003:36:MR1:45;LAPIS:65:H:98:90707:97:MC1:91;LAVATORIO:2:H:100:90004:100:MC1:100;LEITO:11:H:100:90004:55:MR1:45;LEITOR:2:H:50:90701:50:EQ1:50;LENCOL:10:H:90:90707:50:MC1:40;LENCOL/COBERTOR/MANTA:3:H:100:90707:100:MU1:67;LENTE:18:H:67:90711:39:MR1:56;LETREIRO:4:H:100:90004:75:MC1:100;LIGACAO:2:E:100:90002:100:MR1:100;LIMA:20:E:65:90701:75:MR1:35;LIMITADOR:4:E:100:90001:50:MR1:100;LIMPA:7:H:86:90705:57:MC1:29;LIMPADOR:67:H:90:90705:75:ML1:60;LINHA:8:H:50:90004:38:MR1:50;LIQUIDO:2:H:100:90002:50:MR1:100;LIXA:32:H:84:90004:66:MC1:66;LIXEIRA:6:H:100:90705:100:ML1:100;LONA:19:H:68:90004:42:MC1:42;LONGARINA:8:H:50:90002:25:MR1:75;LUMINARIA:38:H:100:90004:97:MR1:61;LUVA:232:H:92:90004:57:MR1:43;LUZ:2:E:100:90002:50:MR1:50;MACACAO:6:H:100:90711:100:EP1:100;MACACO:5:H:80:90701:80:MC1:80;MACANETA:12:E:58:90003:67:MR1:75;MACARICO:6:H:67:90701:83:MC1:50;MACHO:39:E:100:90001:100:MR1:100;MANCAL:113:E:92:90001:88:MR1:96;MANCHAO:8:H:50:90003:62:MR1:100;MANDRIL:14:E:93:90001:93:MR1:93;MANETE:3:E:67:90003:67:MR1:100;MANGA:9:E:78:90001:67:MR1:67;MANGOTE:8:E:75:90002:62:MR1:75;MANGUEIRA:89:E:74:90001:38:MR1:88;MANICOTO:2:E:100:90003:100:MR1:100;MANILHA:11:H:64:90001:36:MR1:45;MANIPULO:4:E:100:90001:50:MR1:100;MANOMETRO:2:E:100:90003:50:AU1:50;MANOMETRO/VACUOMETRO:5:E:80:90001:80:MR1:80;MANOPLA:4:H:75:90001:75:MR1:75;MANTA:8:H:75:90004:75:MC1:75;MANUAL:3:H:100:90701:100:MR1:67;MAO:10:H:90:90004:80:MR1:50;MAQUINA:2:E:100:90002:50:MR1:100;MARCADOR:4:H:100:90707:100:ME2:100;MARRETA:3:H:100:90701:100:ME2:33;MARTELETE:7:H:86:90701:100:MC1:71;MARTELO:11:H:100:90701:100:MC1:91;MASCARA:69:H:100:90711:100:MR1:97;MASSA:53:H:98:90004:94:MC1:89;MEDICAMENTO:191:H:100:90707:98:MR1:90;MEDIDOR:12:E:67:90001:33:MR1:50;MEGOMETRO:3:E:67:90001:67:MR1:67;MEIA:7:H:57:90002:43:MR1:43;MEIO:3:E:100:90001:33:MR1:67;MESA:3:E:67:90707:33:ME2:33;METALICO/N:32:H:100:90004:100:MR1:100;METALICO/NAO:58:H:100:90004:100:MC1:97;MICRO:6:E:83:90001:67:MR1:100;MICROMETRO:5:E:100:90001:100:MR1:100;MICRORRUPTOR:2:H:50:90001:50:MR1:50;MISTURA:2:H:100:90703:50:AL1:50;MISTURADOR:2:E:100:90701:50:EQ1:50;MOCHILA:6:H:100:90707:83:ME2:83;MOD:3:E:67:90005:33:MR1:67;MODULO:30:E:57:90002:37:MR1:80;MODULO/CARTAO/PLACA:36:E:69:90001:53:MR1:75;MOLA:52:E:81:90001:37:MR1:96;MOLDE:8:E:100:90001:88:MR1:88;MOLDURA:7:E:71:90003:57:MR1:43;MONITOR:3:E:100:90002:67:MQ1:67;MONTAGE:2:E:100:90002:100:MR1:100;MOTOBOMBA:2:E:100:90003:50:MR1:100;MOTOR:42:E:76:90003:40:MR1:83;MOTORREDUTOR:6:E:100:90001:100:MR1:100;MOTRIZ:2:E:100:90002:50:MR1:100;MOUSE:4:H:100:90709:75:ME2:25;MUFLA:3:E:67:90001:67:MR1:67;MUNHAO:3:E:100:90001:67:MR1:100;NIPLE:23:H:83:90004:83:MC1:83;NIVEL:10:H:100:90004:70:MC1:60;NOME:50:E:58:90001:52:MR1:54;OCULOS:97:H:100:90711:99:MR1:92;OLEO:451:H:99:90708:96:MR1:71;OLHAL:6:E:100:90001:50:MR1:83;PAINEL:23:E:65:90004:35:MR1:43;PAINEL/QUADRO:8:H:75:90004:88:MC1:62;PALETE/PALLET/ESTRADO:3:H:100:24001:67:MR1:67;PALHA:2:H:100:90705:50:ML1:50;PALHETA:15:E:87:90003:60:MR1:60;PANO:19:H:100:90705:100:MC1:47;PAPEL:47:H:100:90707:51:MC1:43;PAPELAO:6:H:100:90001:67:ME2:67;PAQUIMETRO:3:H:100:90701:100:MC1:67;PARA:24:H:75:90004:58:MR1:88;PARABRISA:9:E:56:90003:67:MR1:89;PARAFUSADEIRA:2:H:100:90701:100:MC1:50;PARAFUSO:1385:E:66:90001:86:MR1:98;PARALAMA:2:E:100:90001:50:MR1:50;PAREDE:2:E:100:90002:50:MR1:100;PARTE:93:E:74:90001:71:MR1:82;PASSA:3:E:67:90002:67:MR1:100;PASTA:12:H:83:90707:58:MR1:58;PASTILHA:25:E:64:90003:68:MR1:84;PATCH:2:H:100:90004:100:MR1:100;PATIM:6:H:50:90003:100:MR1:83;PECAS:2:E:50:90002:100:MR1:100;PEDAL:6:E:67:90003:100:MR1:100;PEDESTAL:2:E:100:90001:100:MR1:100;PEDRA:3:H:100:90707:67:MC1:100;PELICULA:5:H:100:90003:40:MR1:60;PEN:3:H:100:90709:100:IS1:100;PENEIRA:14:H:71:90701:57:ME2:50;PENTE:8:H:50:90709:50:ME2:25;PERFIL:139:H:77:90004:83:MC1:58;PERFILADO:12:H:92:90004:92:MR1:50;PERFURADOR:2:H:50:90707:50:ME2:50;PERNEIRA:3:H:100:90711:33:EP1:33;PERNO:2:E:100:90001:100:MR1:100;PESTANA:5:E:60:90004:40:MR1:60;PIA:3:H:67:90004:100:ME2:33;PICARETA:3:H:67:90701:67:MC1:67;PINCA:5:E:80:90001:60:MR1:80;PINCEL:8:H:100:90707:75:MC1:62;PINCEL/TRINCHA:21:H:95:90004:67:MC1:67;PINHAO:10:E:80:90001:50:MR1:100;PINO:122:E:72:90002:39:MR1:76;PIRULITO:4:E:100:90001:50:MR1:100;PISCA:3:E:100:90003:67:AU1:67;PISO:32:H:94:90004:91:MC1:56;PISTAO:6:E:100:90002:67:MR1:100;PISTOLA:7:E:86:90001:43:MR1:86;PIVO:10:E:80:90003:80:MR1:50;PLACA:106:E:61:90002:30:MR1:57;PLANTA:4:H:100:90004:100:MC1:100;PLAQUETA:2:E:100:90002:100:MR1:100;PLASTICO:2:H:100:90707:50:MC1:100;PLATO:7:E:71:90002:86:MR1:86;PLUGUE:3:E:100:90002:67:MR1:100;PNEU:55:H:51:90003:78:MR1:56;POLIA:22:E:73:90001:77:MR1:91;PONTA:2:E:50:90001:50:MR1:100;PONTEIRA:6:E:100:90001:67:MR1:100;PORCA:352:E:62:90001:81:MR1:96;PORTA:62:H:56:90004:42:MR1:56;POSICIONADOR:6:E:100:90001:100:MR1:100;POSTE:11:H:91:90004:82:MR1:73;POTE:2:H:100:90705:50:ML1:50;POTENCIOMETRO:3:E:100:90001:67:MR1:100;PRANCHA/PRANCHAO:2:H:100:90151:100:OP1:50;PRANCHETA:2:H:100:90707:100:ME2:50;PRATO:3:E:67:90001:33:MR1:33;PREGO:29:H:76:90004:76:MR1:55;PRESILHA:7:E:86:90003:43:MR1:71;PRESSOSTATO:5:E:60:90001:60:MR1:100;PRIMER:11:H:100:90004:91:MR1:91;PRISIONEIRO:4:E:100:90001:75:MR1:100;PROMOCIONAL:2:H:100:90707:100:ME2:100;PROTECAO:29:E:66:90002:38:MR1:76;PROTETOR:31:H:68:90003:39:MR1:65;PROVETA:4:H:100:90706:100:LB2:100;PRUMO:6:H:83:90004:67:MR1:83;PULVERIZADOR:4:H:75:90001:75:MR1:75;PURGADOR:15:E:87:90001:87:MR1:100;PUXADOR:3:H:67:90004:33:MR1:67;QUADRO:24:H:75:90004:46:MC1:38;QUIMICO:166:H:99:90706:91:LB2:77;QUINTA:2:H:50:90003:100:MQ1:50;RACK:2:H:100:90707:50:MR1:50;RADIADOR:4:H:50:90003:75:MR1:75;RALO:2:H:100:90004:100:MC1:100;RAMAL:2:E:100:90002:50:MR1:100;RASPADOR:8:E:100:90001:100:MR1:100;RATICIDA:4:H:100:90705:100:ME2:50;REATOR:8:H:100:90101:100:MC1:88;REBITADOR:2:H:100:90701:100:MC1:50;REBITE:9:H:78:90004:56:MR1:100;REBOLO:6:E:83:90001:83:MR1:100;RECEPTOR:2:E:100:90709:50:SH1:50;REDUCAO:32:H:91:90004:59:MR1:97;REDUTOR:7:E:100:90001:86:MR1:86;REFIL:8:H:62:90705:38:MR1:38;REFLETOR:22:H:95:90004:91:MC1:68;REFORCO:4:E:100:90002:75:MR1:100;REGISTRO:22:H:95:90004:95:MR1:68;REGUA:10:H:100:90707:40:MR1:60;REGULADOR:19:E:100:90001:79:MR1:89;REJUNTE:13:H:100:90004:100:MR1:85;RELE:54:E:80:90001:59:MR1:89;RELOGIO:4:E:75:90003:75:AU1:75;REMENDO:3:H:67:90001:100:MR1:100;REMOVEDOR:2:E:100:90001:100:MR1:100;REPARO:46:E:89:90001:61:MR1:85;RESERVATORIO:5:E:80:90001:40:MR1:80;RESFRIADOR:2:E:50:90003:100:AU1:50;RESIDUO:5:N:60:90151:60:MC1:60;RESINA:14:H:86:90004:57:MR1:79;RESISTENCIA:3:E:100:90001:67:MR1:100;RESISTOR:3:E:100:90002:67:MR1:100;RESPIRADOR:4:H:75:90711:100:MC1:50;RESPIRO:5:E:100:90002:60:MR1:100;RETENTOR:280:E:98:90001:87:MR1:96;RETIFICA:2:H:100:90701:100:MC1:100;RETIFICADOR:3:E:100:90001:67:MR1:67;RETROVISOR:7:E:86:90002:43:MR1:100;REVESTIMENTO:7:H:71:90004:86:MR1:57;RIPA:7:H:100:90151:86:MC1:57;RODA:22:E:91:90001:45:MR1:82;RODAPE:3:H:100:90003:67:MR1:67;RODO:6:H:100:90705:83:ML1:67;ROLAMENTO:632:E:96:90002:82:MR1:96;ROLAMENTOS:2:E:100:90002:100:MR1:100;ROLDANA:6:E:83:90001:83:MR1:67;ROLETE:33:E:100:90001:64:MR1:100;ROLO:32:H:88:90004:78:MC1:81;ROTEADOR:6:H:83:90709:50:MC1:50;ROTOR:12:E:100:90001:83:MR1:100;ROTULA:4:E:100:90002:75:MR1:100;ROTULADOR/ETIQUETADOR:2:H:50:90701:50:ME2:50;ROUPA:7:H:100:90711:100:MC1:71;SABAO:49:H:100:90705:100:MR1:82;SABONETE:167:H:100:90707:99:IA1:65;SACA:3:E:100:90001:100:MR1:100;SACO:21:H:100:90705:86:ML1:62;SACOLA:6:H:100:90707:83:MC1:83;SAIDA:19:H:84:90004:89:MR1:58;SAPATA:14:E:93:90002:50:MR1:86;SAPATILHA:5:H:60:90001:80:MR1:80;SEDE:6:E:100:90001:67:MR1:100;SELADOR:4:H:50:90001:50:MR1:75;SELO:19:E:89:90001:58:MR1:95;SEMI:2:E:100:90002:50:MR1:100;SENSOR:139:E:93:90001:52:MR1:90;SERINGA:81:H:100:90707:100:MR1:100;SERRA:22:H:86:90701:45:MR1:55;SIFAO:2:H:100:90004:100:MC1:100;SILENCIOSO:3:E:100:90001:33:MR1:100;SILICONE:6:H:100:90002:50:MC1:83;SINALEIRO:3:H:67:90001:67:MR1:100;SINALIZADOR:11:H:64:90001:45:MR1:45;SINCRONIZADO:3:H:67:90002:100:MR1:100;SIRENE:5:E:60:90101:60:MR1:100;SOLDA:5:H:80:90001:60:MR1:60;SOLDAVEL:2:E:100:90151:100:MR1:100;SOLENOIDE:7:E:100:90002:57:MR1:86;SOLUCAO:2:H:100:90001:50:MR1:50;SOLVENTE:13:H:92:90004:54:MC1:46;SONDA:3:E:100:90003:100:AU1:100;SOPRADOR:6:E:50:90002:50:MR1:67;SOQUETE:89:H:83:90701:82:MC1:76;SOQUETE/RECEPTACULO:14:H:71:90004:71:MC1:50;SOQUETEIRA:5:H:80:90004:80:MC1:60;SUPORTE:121:H:51:90101:26:MR1:56;SWITCH:5:H:100:90709:80:MC1:80;TABUA:18:H:100:90004:94:MR1:50;TALABARTE:7:H:71:90711:43:MR1:57;TALHADEIRA/PONTEIRO:10:H:70:90004:50:MR1:100;TALISCA:2:E:100:90002:100:MR1:100;TAMBOR:3:E:100:90003:67:MR1:67;TAMPA:134:E:66:90001:35:MR1:74;TAMPAO:39:H:77:90004:74:MR1:49;TANQUE:9:E:67:90002:56:MR1:89;TAPETE/CAPACHO:9:H:100:90707:100:MC1:89;TECIDO:10:H:100:90707:70:ME2:90;TECLADO:5:H:60:90707:40:MR1:100;TEE:30:H:77:90004:77:MR1:73;TELA:37:H:78:90004:73:MC1:59;TELHA:30:H:90:90004:77:MC1:47;TENSIONADOR:3:E:67:90002:67:MR1:67;TENSOR:5:E:100:90002:60:MR1:100;TERMINAL:373:E:69:90001:93:MR1:95;TERMOMETRO:157:H:100:90707:76:MC1:50;TERMOSTATO:8:E:88:90001:62:MR1:100;TESOURA:183:H:100:90707:84:MC1:81;TIJOLO:17:H:100:90004:76:MR1:59;TINTA:176:H:94:90004:76:MR1:59;TIRA:2:E:100:90002:100:MR1:100;TIRANTE:3:E:100:90002:67:MR1:67;TOALHA:3:H:100:90707:67:ME2:33;TOCHA:3:E:100:90001:67:MR1:67;TOMADA:80:H:98:90004:96:MC1:60;TONER:106:H:100:90707:100:MR1:89;TORNEIRA:34:H:88:90151:74:MR1:76;TORQUES:4:H:100:90701:50:MR1:50;TOUCA:4:H:100:90711:100:EP1:50;TRAFO:7:H:86:90004:86:MC1:86;TRANSCEIVER:3:E:100:90001:67:MR1:100;TRANSDUTOR:5:H:60:90004:60:MR1:80;TRANSFORMADOR:2:E:100:90001:100:MR1:100;TRANSMISSOR:15:E:87:90001:73:MR1:87;TRANSMISSOR/MEDIDOR:2:E:100:90001:100:MR1:100;TRAVA:13:E:85:90002:54:MR1:85;TRELICA:5:H:100:90002:80:MR1:80;TRENA:24:H:100:90701:92:MR1:67;TRILHO:8:H:62:90001:75:MR1:100;TUBO:150:E:51:90004:43:MR1:62;TUBO/MANGUEIRA/ESPIRAL/JARDIM:3:E:100:90701:67:MR1:67;TURBOCOMPRESSOR:2:E:100:90003:50:AU1:50;UNIAO:57:E:81:90001:96:MR1:98;UNIDADE:8:E:88:90001:75:MR1:88;UTENSILIO:7:H:100:90707:71:MU1:29;VALVULA:214:E:93:90001:70:MR1:94;VARETA:6:E:100:90001:67:MR1:100;VASELINA:3:H:100:90707:67:MC1:67;VASO:3:H:100:90707:100:ME2:67;VASSOURA:158:H:100:90705:100:MR1:95;VASSOURAO:2:H:100:90705:100:MR1:100;VEDACAO:64:E:98:90002:72:MR1:88;VEDADOR:4:E:100:90001:50:MR1:100;VEDANTE:2:H:100:90001:50:MR1:50;VELA:5:E:100:90002:40:MQ1:40;VENTILADOR:49:E:96:90001:57:MR1:73;VERGALHAO:20:H:60:90004:65:MR1:40;VERNIZ:10:H:100:90004:50:MR1:40;VIDRO:13:H:54:90004:46:MR1:92;VIGA:5:E:60:90004:60:MC1:40;VIRABREQUIM:2:E:100:90003:50:AU1:100;VISOR:4:E:75:90001:50:MR1:100";
const HEAD = {};
HEADRAW.split(";").forEach(function(r){
 const p=r.split(":"); if(p.length<8) return;
 HEAD[p[0]]={n:+p[1], tm:(p[2]==="E"?"ERSA":(p[2]==="H"?"HIBE":"NLAG")), ptm:+p[3]/100, gm:p[4], pgm:+p[5]/100, gc:p[6], pgc:+p[7]/100};
});
/* complemento curado para substantivos tipicos com pouco historico no 302 */
const HEAD_CUR = {
 "CIMENTO":["HIBE","90151","MC1"],"AREIA":["HIBE","90151","MC1"],"BRITA":["HIBE","90151","MC1"],
 "CONCRETO":["HIBE","90151","MC1"],"MADEIRA":["HIBE","90004","MC1"],"COMPENSADO":["HIBE","90004","MC1"],
 "IMPRESSORA":["HIBE","90707","MR1"],"MONITOR":["HIBE","90709","IS1"],"MOUSE":["HIBE","90709","IS1"],
 "TECLADO":["HIBE","90709","IS1"],"CABO_REDE":["HIBE","90709","IS1"],"PENDRIVE":["HIBE","90709","IS1"],
 "ENVELOPE":["HIBE","90707","IA1"],"GRAMPEADOR":["HIBE","90707","IA1"],"PASTA":["HIBE","90707","IA1"],
 "AVENTAL":["HIBE","90711","EP1"],"PROTETOR_AURICULAR":["HIBE","90711","EP1"],"CINTO":["HIBE","90711","EP1"],
 "PERFURATRIZ":["ERSA","90002","MR1"],"MANOMETRO":["ERSA","90001","MR1"],"PRESSOSTATO":["ERSA","90001","MR1"],
 "TRANSDUTOR":["ERSA","90001","MR1"],"INJETOR":["ERSA","90002","MR1"],"TURBINA":["ERSA","90002","MR1"],
 "PISTAO":["ERSA","90002","MR1"],"BIELA":["ERSA","90002","MR1"],"VIRABREQUIM":["ERSA","90002","MR1"],
 "DESINCRUSTANTE":["HIBE","90705","ML1"],"AGUARRAS":["HIBE","90004","MC1"],"THINNER":["HIBE","90004","MC1"]
};
function headLook(hd){
 if(!hd) return null;
 const hb=S.headBase.get(hd);
 if(hb && hb.n>=2) return {src:"base", n:hb.n, tm:best(hb.tm), gm:best(hb.gm), gc:best(hb.gc)};
 if(HEAD[hd]){ const e=HEAD[hd]; return {src:"seed", n:e.n, tm:{k:e.tm,p:e.ptm}, gm:{k:e.gm,p:e.pgm}, gc:{k:e.gc,p:e.pgc}} }
 if(HEAD_CUR[hd]){ const c=HEAD_CUR[hd]; return {src:"curado", n:14, tm:{k:c[0],p:0.9}, gm:{k:c[1],p:0.9}, gc:{k:c[2],p:0.8}} }
 if(hb && hb.n===1) return {src:"base", n:1, soft:true, tm:best(hb.tm), gm:best(hb.gm), gc:best(hb.gc)};
 let bk=null,bn=0;
 S.headBase.forEach(function(v,k){
  if(v.n>bn && Math.min(k.length,hd.length)>=4 && (k.indexOf(hd)===0||hd.indexOf(k)===0)){bn=v.n;bk=k}
 });
 if(bk){ const v=S.headBase.get(bk); return {src:"base", soft:true, n:v.n, tm:best(v.tm), gm:best(v.gm), gc:best(v.gc), via:bk} }
 for(const k in HEAD){ if(Math.min(k.length,hd.length)>=4 && (k.indexOf(hd)===0||hd.indexOf(k)===0) && HEAD[k].n>bn){bn=HEAD[k].n;bk=k} }
 if(bk){ const e=HEAD[bk]; return {src:"seed", soft:true, n:e.n, tm:{k:e.tm,p:e.ptm}, gm:{k:e.gm,p:e.pgm}, gc:{k:e.gc,p:e.pgc}, via:bk} }
 return null;
}
/* confianca calibrada por campo: evidencia do historico com encolhimento bayesiano,
   sem punir o uso antes de carregar a planilha da base */
function shrink(p,n,k){ n=n||0; return (p*n + 0.34*k)/(n+k) }
function fieldConf(share, hl, f, maxSim, ev){
 const val=[], wgt=[];
 val.push(Math.max(0,Math.min(1,share||0))); wgt.push(1.0);
 if(hl && hl[f] && hl[f].k){
  const hp=shrink(hl[f].p, hl.n, 4)*(hl.soft?0.88:1);
  val.push(hp); wgt.push(Math.min(2.8, 0.9+Math.log(1+hl.n)/1.6));
 }
 if(S.base.length && maxSim>0){ val.push(Math.min(1,maxSim/0.88)); wgt.push(1.3) }
 let num=0,den=0; for(let i=0;i<val.length;i++){num+=val[i]*wgt[i];den+=wgt[i]}
 if(!den) return 0.15;
 const sup = 0.50 + 0.50*Math.max(0,Math.min(1, ev===undefined?1:ev)); /* sem evidencia real a confianca nao sobe */
 return Math.max(0.05, Math.min(0.99, (num/den)*sup));
}

/* ================= REGRA DE IMOBILIZADO ================= */
function checkImob(desc, longa){
 const t = toks(desc); const full = norm(desc+" "+(longa||""));
 const head = t.slice(0,2);
 const hasPart = t.some(function(x){return IMOB.peca.indexOf(x)>=0});
 for(let i=0;i<IMOB.excTok.length;i++){
  const e=IMOB.excTok[i];
  if(t.indexOf(e)>=0 && !hasPart) return {flag:true, lvl:"alto", why:"Excecao da norma: "+e+" e sempre imobilizado (Fluxo 304 / ZATI), independente do valor"};
 }
 for(let i=0;i<IMOB.bens.length;i++){
  const b=IMOB.bens[i];
  if(head.indexOf(b)>=0 && !hasPart) return {flag:true, lvl:"alto", why:"Bem duravel como item principal ("+b+"): vida util acima de 1 ano, avaliar Fluxo 304 / ZATI"};
  if(new RegExp("(^| )"+b+"( |$)").test(full) && !hasPart && t.length<=4) return {flag:true, lvl:"medio", why:"Descricao aponta o bem completo ("+b+") e nao um componente: validar contra a regra de imobilizado"};
 }
 if(!hasPart && /COMPLETO|NOVO|UNIDADE COMPLETA|MONTADO/.test(full) && t.length<=5)
  return {flag:true, lvl:"baixo", why:"Indicio de bem completo: confirmar se nao e imobilizado (Fluxo 304)"};
 return {flag:false, lvl:"", why:""};
}

/* ================= CLASSIFICADOR ================= */
/* ================= v10: ENTRADA UNICA PONDERADA, NCM E TRAVA DE DESCONTINUADOS ================= */
function uniDesc(list){
 const seen={}, o=[];
 (list||[]).forEach(function(t){ toks(String(t==null?"":t)).forEach(function(w){ if(!seen[w]){ seen[w]=1; o.push(w) } }) });
 return o.join(" ");
}
function pickHead(parts, tks){
 for(let i=0;i<(parts||[]).length;i++){
  const tt=toks(parts[i].t||"");
  for(let j=0;j<tt.length && j<3;j++){ const h=headLook(tt[j]); if(h && (h.n||0)>=3) return tt[j] }
 }
 for(let i=0;i<(parts||[]).length;i++){ const tt=toks(parts[i].t||""); if(tt.length) return tt[0] }
 return (tks&&tks[0])||"";
}
function ncmClean(v){ const d=String(v==null?"":v).replace(/\D/g,""); return d.length>=4? d.slice(0,8) : "" }
const NCM_MAP = {
 "271019":{gm:"90708",gc:"MC1",tm:"HIBE"},"271012":{gm:"29002",gc:"CO1",tm:"HIBE"},"271020":{gm:"29002",gc:"CO1",tm:"HIBE"},"3403":{gm:"90708",gc:"MC1",tm:"HIBE"},"27":{gm:"29002",gc:"CO1"},
 "3401":{gm:"90705",gc:"ML1",tm:"HIBE"},"3402":{gm:"90705",gc:"ML1",tm:"HIBE"},"3405":{gm:"90705",gc:"ML1"},"3808":{gm:"90705",gc:"PQ1"},"3824":{gc:"PQ1"},
 "3204":{gm:"90004",gc:"MC1"},"3208":{gm:"90004",gc:"MC1"},"3209":{gm:"90004",gc:"MC1"},"3210":{gm:"90004",gc:"MC1"},"3214":{gm:"90004",gc:"MC1"},
 "28":{gc:"PQ1"},"29":{gc:"PQ1"},"38":{gc:"PQ1"},"3002":{gc:"VM1"},"3004":{gc:"VM1"},"30":{gc:"VM1"},
 "3917":{gm:"90004",gc:"MC1"},"3923":{gm:"24001",gc:"MC1"},"3926":{gm:"90004",gc:"MC1"},
 "4009":{gm:"90001",gc:"MC1",tm:"ERSA"},"4010":{gm:"90001",gc:"EQ1",tm:"ERSA"},"4016":{gm:"90001",gc:"MC1",tm:"ERSA"},"4015":{gm:"90711",gc:"MC1",tm:"HIBE"},
 "4407":{gm:"90151",gc:"MC1"},"44":{gm:"90151",gc:"MC1"},
 "4802":{gm:"90707",gc:"MC1"},"4817":{gm:"90707",gc:"MC1"},"4820":{gm:"90707",gc:"MC1"},"4821":{gm:"90707",gc:"MC1"},"4818":{gm:"90705",gc:"ML1"},"48":{gm:"90707",gc:"MC1"},
 "6307":{gm:"90705",gc:"ML1"},"6116":{gm:"90711",gc:"MC1",tm:"HIBE"},"6403":{gm:"90711",gc:"MC1",tm:"HIBE"},"6401":{gm:"90711",gc:"MC1",tm:"HIBE"},"6506":{gm:"90711",gc:"MC1",tm:"HIBE"},
 "6804":{gm:"90701",gc:"MC1"},"6805":{gm:"90701",gc:"MC1"},"68":{gm:"90151",gc:"MC1"},"6907":{gm:"90151",gc:"MC1"},"69":{gm:"90151",gc:"MC1"},"7005":{gm:"90004",gc:"MC1"},"70":{gm:"90004",gc:"MC1"},
 "7307":{gm:"90004",gc:"MC1"},"7318":{gm:"90004",gc:"MC1"},"7306":{gm:"90151",gc:"MC1"},"72":{gm:"90151",gc:"MC1"},"73":{gm:"90004",gc:"MC1"},"76":{gm:"90004",gc:"MC1"},
 "8201":{gm:"90701",gc:"MC1"},"8203":{gm:"90701",gc:"MC1"},"8204":{gm:"90701",gc:"MC1"},"8205":{gm:"90701",gc:"MC1"},"8206":{gm:"90701",gc:"MC1"},"8207":{gm:"90701",gc:"MC1"},"82":{gm:"90701",gc:"MC1"},
 "8302":{gm:"90004",gc:"MC1"},"83":{gm:"90004",gc:"MC1"},
 "8413":{gm:"90001",gc:"EQ1",tm:"ERSA"},"8481":{gm:"90001",gc:"MC1",tm:"ERSA"},"8482":{gm:"90001",gc:"EQ1",tm:"ERSA"},"8483":{gm:"90001",gc:"EQ1",tm:"ERSA"},"8471":{gm:"90709",gc:"MC1"},"8443":{gm:"90709",gc:"MC1"},"84":{gm:"90001",gc:"EQ1",tm:"ERSA"},
 "8504":{gm:"90101",gc:"ME1"},"8536":{gm:"90101",gc:"ME1"},"8537":{gm:"90101",gc:"ME1"},"8538":{gm:"90101",gc:"ME1"},"8539":{gm:"90004",gc:"ME1"},"8544":{gm:"90101",gc:"ME1"},"8506":{gm:"90707",gc:"MC1"},"8517":{gm:"90709",gc:"MC1"},"85":{gm:"90101",gc:"ME1"},
 "8708":{gm:"90003",gc:"MC1",tm:"ERSA"},"87":{gm:"90003",gc:"MC1"},
 "9004":{gm:"90711",gc:"MC1",tm:"HIBE"},"9020":{gm:"90711",gc:"MC1",tm:"HIBE"},"9025":{gm:"90001",gc:"EQ1"},"9026":{gm:"90001",gc:"EQ1"},"9032":{gm:"90001",gc:"ME1"},"90":{gm:"90001",gc:"EQ1"},
 "9401":{gm:"90701",gc:"MU1"},"9403":{gm:"90701",gc:"MU1"},"94":{gm:"90701",gc:"MU1"},
 "9603":{gm:"90705",gc:"ML1"},"9608":{gm:"90707",gc:"MC1"},"96":{gm:"90707",gc:"MC1"},
 "23":{gm:"14001",gc:"IP2"},"3105":{gc:"IP2"},"31":{gc:"IP2"}
};
function ncmIxBuild(){
 S.ncmIx=new Map(); S.ncmIxN=(S.base||[]).length;
 (S.base||[]).forEach(function(r){
  const n=ncmClean(r.ncm); if(!n) return;
  [n.slice(0,8),n.slice(0,6),n.slice(0,4)].forEach(function(p){
   if(p.length<4) return;
   let o=S.ncmIx.get(p); if(!o){ o={tot:0,tm:{},gm:{},gc:{}}; S.ncmIx.set(p,o) }
   o.tot++;
   if(r.tm) o.tm[r.tm]=(o.tm[r.tm]||0)+1;
   if(r.gm && codeOk("gm",r.gm)) o.gm[r.gm]=(o.gm[r.gm]||0)+1;
   if(r.gc && codeOk("gc",r.gc)) o.gc[r.gc]=(o.gc[r.gc]||0)+1;
  });
 });
}
function ncmLook(n){
 if(!n) return null;
 if(!S.ncmIx || S.ncmIxN!==(S.base||[]).length) ncmIxBuild();
 const top=function(mp,tot){ let k="",v=0; for(const x in mp){ if(mp[x]>v){ v=mp[x]; k=x } } return {k:k,p:tot? v/tot:0} };
 const tr=[[n.slice(0,8),1.00],[n.slice(0,6),0.92],[n.slice(0,4),0.80]];
 for(let i=0;i<tr.length;i++){
  if(tr[i][0].length<4) continue;
  const o=S.ncmIx.get(tr[i][0]);
  if(o && o.tot>=2){
   const tm=top(o.tm,o.tot), gm=top(o.gm,o.tot), gc=top(o.gc,o.tot);
   return {lvl:tr[i][0].length, w:(2.2+Math.min(1.4,o.tot/12))*tr[i][1], src:"NCM "+tr[i][0]+" da base ("+o.tot+" registros)",
           tm:tm.p>=0.55?tm.k:"", gm:gm.p>=0.50?gm.k:"", gc:gc.p>=0.50?gc.k:""};
  }
 }
 const ks=[n.slice(0,8),n.slice(0,6),n.slice(0,4),n.slice(0,2)];
 for(let i=0;i<ks.length;i++){
  const o=NCM_MAP[ks[i]];
  if(o) return {lvl:ks[i].length, w:(ks[i].length>=6?2.4:(ks[i].length===4?2.0:1.2)), src:"NCM "+ks[i], tm:o.tm||"", gm:o.gm||"", gc:o.gc||""};
 }
 return null;
}
function gcGlobal(){
 const c={}; let tot=0;
 (S.base||[]).forEach(function(r){ if(r.gc && codeOk("gc",r.gc)){ c[r.gc]=(c[r.gc]||0)+1; tot++ } });
 if(!tot){ for(const g in GM2GC){ const li=GM2GC[g]||[]; li.forEach(function(p){ if(codeOk("gc",p[0])){ c[p[0]]=(c[p[0]]||0)+p[1]; tot+=p[1] } }) } }
 let k="",v=0; for(const x in c){ if(c[x]>v){ v=c[x]; k=x } }
 return {k:k, n:v, p:tot? v/tot:0};
}
function gcFromGm(gm){
 const li=gm?GM2GC[gm]:null; if(!li || !li.length) return "";
 let k="",v=0; li.forEach(function(p){ if(p[1]>v && codeOk("gc",p[0])){ v=p[1]; k=p[0] } });
 return k;
}
function sanitizePriors(){
 const fixGm=function(g){ if(!g) return ""; if(codeOk("gm",g)) return g; const s2=sucOf("gm",g); return (s2&&codeOk("gm",s2))?s2:"" };
 const fixGc=function(g,gm){ if(!g) return ""; if(codeOk("gc",g)) return g; const s2=sucOf("gc",g); if(s2&&codeOk("gc",s2)) return s2; return gcFromGm(gm) };
 const clean=function(o){
  if(!o||typeof o!=="object") return;
  if(Array.isArray(o)){ o[1]=fixGm(o[1]); o[2]=fixGc(o[2],o[1]); return }
  if(typeof o.gm==="string") o.gm=fixGm(o.gm);
  if(typeof o.gc==="string") o.gc=fixGc(o.gc, typeof o.gm==="string"?o.gm:"");
  ["gm","gc"].forEach(function(f){
   const v=o[f];
   if(v && typeof v==="object"){
    for(const c in v){
     if(!codeOk(f,c)){
      const sc=sucOf(f,c);
      if(sc && codeOk(f,sc)) v[sc]=(v[sc]||0)+v[c];
      delete v[c];
     }
    }
   }
  });
 };
 try{ for(const k in HEAD) clean(HEAD[k]) }catch(e){}
 try{ for(const k in HEAD_CUR) clean(HEAD_CUR[k]) }catch(e){}
 try{ if(typeof HEAD_GCV!=="undefined") for(const k in HEAD_GCV){ const o=HEAD_GCV[k]; if(!o) continue;
      if(typeof o.k==="string"){ const nk=fixGc(o.k,""); if(nk!==o.k){ o.k=nk; if(!nk) o.tot=0 } } clean(o) } }catch(e){}
 try{ for(const g in GM2GC){ const li=GM2GC[g]; if(li && li.filter) GM2GC[g]=li.filter(function(p){ return codeOk("gc",p[0]) }) } }catch(e){}
 try{ (S.ex||[]).forEach(function(e){ e.gm=fixGm(e.gm); e.gc=fixGc(e.gc,e.gm) }) }catch(e){}
}
/* ================= v11: RECUPERACAO BM25 POR SUBSTANTIVO E GC DERIVADO DO GM =================
   metodo: o substantivo principal filtra os candidatos na base 302; o ranqueamento pesa
   tokens raros (IDF/BM25) em vez de tratar todo token igual; o Grupo de Mercadoria sai da
   maioria ponderada dos vizinhos e o Grupo de Compradores e DERIVADO do GM escolhido. */
function bmBuild(){
 const B={df:{}, rows:[], post:{}, avg:0, n:0, head:new Map()};
 (S.base||[]).forEach(function(r){
  const tk=toks((r.d||"")+" "+(r.c||"")+" "+(r.l||""));
  if(!tk.length) return;
  const ix=B.rows.length;
  B.rows.push({r:r, tk:tk, len:tk.length, hd:tk[0]||""});
  const u={}; tk.forEach(function(x){ u[x]=1 });
  for(const x in u){
   B.df[x]=(B.df[x]||0)+1;
   let p=B.post[x]; if(!p){ p=B.post[x]=[] }
   if(p.length<4000) p.push(ix);
  }
  const h=tk[0]||""; if(h){ if(!B.head.has(h)) B.head.set(h,[]); B.head.get(h).push(ix) }
  B.n++;
 });
 B.avg = B.n? B.rows.reduce(function(a,x){ return a+x.len },0)/B.n : 0;
 S.bm=B; S.bmN=(S.base||[]).length;
}
function bmIdf(tk){
 const B=S.bm; if(!B||!B.n) return 1;
 const df=B.df[tk]||0;
 return Math.max(0.15, Math.log(1+(B.n-df+0.5)/(df+0.5)));
}
function bmAttr(tk){ return /\d/.test(tk) && tk.length>=2 }
function bm25Rank(parts, hd, kk){
 if(!(S.base||[]).length) return [];
 if(!S.bm || S.bmN!==(S.base||[]).length) bmBuild();
 const B=S.bm; if(!B.n) return [];
 const k1=1.5, bb=0.75;
 const qw={}, qtxt=[];
 (parts||[]).forEach(function(p){
  const tk=toks(p.t||"");
  if(tk.length) qtxt.push(p.t);
  tk.forEach(function(x){ qw[x]=Math.max(qw[x]||0, p.w) });
 });
 const qk=Object.keys(qw); if(!qk.length) return [];
 const qtri=tri(qtxt.join(" "));
 /* candidatos: substantivo principal como filtro rigido */
 let cand=[], soft=false;
 if(hd && B.head.has(hd)) cand=B.head.get(hd).slice();
 if(cand.length<3){
  /* sem massa critica no substantivo: abre para os tokens mais raros da consulta */
  soft=true;
  const raros=qk.slice().sort(function(a,b){ return bmIdf(b)-bmIdf(a) }).slice(0,6);
  const seen={};
  raros.forEach(function(x){
   const p=B.post[x]||[];
   for(let i=0;i<p.length && Object.keys(seen).length<1200;i++){ seen[p[i]]=1 }
  });
  if(hd && B.head.has(hd)) B.head.get(hd).forEach(function(i){ seen[i]=1 });
  cand=Object.keys(seen).map(Number);
 }
 const sc=[];
 for(let ci=0; ci<cand.length; ci++){
  const row=B.rows[cand[ci]]; if(!row) continue;
  const tf={}; row.tk.forEach(function(x){ tf[x]=(tf[x]||0)+1 });
  let v=0, hit=0, at=0;
  for(let j=0;j<qk.length;j++){
   const x=qk[j], f=tf[x]||0; if(!f) continue;
   v += bmIdf(x)*qw[x]*(f*(k1+1))/(f + k1*(1-bb+bb*(row.len/(B.avg||1))));
   hit++; if(bmAttr(x)) at++;
  }
  if(v<=0) continue;
  if(at) v *= (1+0.25*Math.min(3,at));      /* atributo tecnico casado (medida, norma, codigo) vale mais */
  if(row.hd===hd) v *= 1.15;
  sc.push({r:row.r, sc:v, hit:hit, soft:soft});
 }
 sc.sort(function(a,b){ return b.sc-a.sc });
 const top=sc.slice(0, kk||20);
 top.forEach(function(x){
  x.s = diceSet(qtri, tri((x.r.d||"")+" "+(x.r.c||"")+" "+(x.r.l||"")));
  x.pw = 1;
 });
 return top;
}
function deriveGc(gm, nbBm, exHit, fb, hd, txtRule, qtxt){
 if(!gm || !codeOk("gm",gm)) return {k:"", c:0.25, via:"Grupo de Mercadoria indefinido: Grupo de Compradores nao derivado - REVISAR"};
 if(exHit && exHit.s>=0.88 && exHit.e.gm===gm && codeOk("gc",exHit.e.gc))
  return {k:exHit.e.gc, c:0.93, via:"exemplo da base de conhecimento com o mesmo Grupo de Mercadoria: "+exHit.e.d};
 let fbb=null;
 (fb||[]).forEach(function(f){
  const sm=diceSet(tri(qtxt),tri(f.d));
  if(sm>=0.85 && f.gm===gm && codeOk("gc",f.gc) && (!fbb||fbb.s<sm)) fbb={s:sm,f:f};
 });
 if(fbb) return {k:fbb.f.gc, c:0.92, via:"correcao sua em material equivalente do mesmo Grupo de Mercadoria "+gm};
 const vv={}; let tot=0, n=0;
 (nbBm||[]).forEach(function(x){
  if(x.r.gm!==gm || !x.r.gc || !codeOk("gc",x.r.gc)) return;
  const w=x.sc||1; vv[x.r.gc]=(vv[x.r.gc]||0)+w; tot+=w; n++;
 });
 const b=best(vv);
 if(b.k && tot>0){
  const ag=vv[b.k]/tot;
  if(ag>=0.6) return {k:b.k, c:Math.min(0.90,0.55+0.35*ag), via:"maioria dos semelhantes dentro do Grupo de Mercadoria "+gm+": "+Math.round(ag*100)+"% de "+n+" vizinho(s) em "+b.k};
 }
 const rgc=ruleGc(txtRule);
 if(rgc && codeOk("gc",rgc.k)){
  const li=GM2GC[gm]||[];
  if(!li.length || li.some(function(p){ return p[0]===rgc.k }))
   return {k:rgc.k, c:0.82, via:"regra de familia ("+rgc.why+") compativel com o Grupo de Mercadoria "+gm};
 }
 const li=GM2GC[gm]||[];
 if(li.length){
  const tt=li.reduce(function(a,x){ return a+x[1] },0);
  const vl=li.filter(function(p){ return codeOk("gc",p[0]) });
  if(vl.length) return {k:vl[0][0], c:Math.min(0.78,0.48+0.30*(vl[0][1]/tt)), via:"predominancia do Grupo de Mercadoria "+gm+": "+vl[0][1]+" de "+tt+" registro(s) em "+vl[0][0]};
 }
 const hev=HEAD_GCV[hd];
 if(hev && hev.tot>=3 && codeOk("gc",hev.k)){
  const li2=GM2GC[gm]||[];
  if(!li2.length || li2.some(function(p){ return p[0]===hev.k }))
   return {k:hev.k, c:Math.min(0.72,0.45+0.30*(hev.n/hev.tot)), via:"historico do substantivo "+hd+" coerente com o Grupo de Mercadoria "+gm+": "+hev.n+" de "+hev.tot};
 }
 const fgc=famGc(gm, txtRule);
 if(fgc.k) return {k:fgc.k, c:fgc.c, via:fgc.via};
 const gg=gcGlobal();
 if(gg.k) return {k:gg.k, c:Math.min(0.52,0.38+0.14*gg.p), via:"sem vizinho no Grupo de Mercadoria "+gm+": predominancia geral do 302 - sugerimos confirmar"};
 return {k:"", c:0.25, via:"nenhum codigo do 627 sustentado pelo Grupo de Mercadoria "+gm+" - REVISAR"};
}

/* ================= CALIBRACAO PELA BASE 302 REAL (v7.2) =================
   Tabelas geradas da base 302 completa enviada pelo usuario (20.793 materiais com Grupo de
   Mercadoria e Grupo de Compradores validados). Para cada combinacao de substantivo + aplicacao +
   NCM + tipo/modificador guardamos o codigo majoritario, a concordancia e quantos materiais
   sustentam a evidencia. O motor vota entre os niveis, do mais especifico ao mais generico, e
   mostra na transparencia qual nivel decidiu. Nenhum codigo novo: os indices apontam somente
   para os catalogos oficiais 663 e 627 ja embarcados e validados. */
const CAL_GM=["14001","24001","24002","29001","29002","29003","90001","90002","90003","90004","90005","90101","90151","90701","90703","90705","90706","90707","90708","90709","90711","90712"];
const CAL_GC=["AL1","AU1","CO1","CO2","EP1","EQ1","EV1","IA1","IF1","II2","IP1","IP2","IS1","LB1","LB2","LC1","MC1","ME1","ME2","MG3","ML1","MQ1","MR1","MU1","OP1","PQ1","SH1","TI1","UF1","VM1"];
const CAL_TM=["ERSA","HIBE","NLAG"];
const CAL_W={A8:3.0,A6:2.4,A4:2.0,T:1.8,B6:1.6,B4:1.4,M:1.3,C:1.0,D:0.8,E:0.7,F:0.6,G:0.4};
const CAL_LVL={A8:"substantivo + aplicacao + NCM (8 digitos)",A6:"substantivo + aplicacao + NCM (6 digitos)",A4:"substantivo + aplicacao + posicao NCM",T:"substantivo + tipo + aplicacao",B6:"substantivo + NCM (6 digitos)",B4:"substantivo + posicao NCM",M:"substantivo + modificador + aplicacao",C:"substantivo + aplicacao",D:"substantivo",E:"NCM (6 digitos) + aplicacao",F:"posicao NCM + aplicacao",G:"posicao NCM"};

/* ===== descompressao sincrona embutida (gzip/deflate em javascript puro) =====
   Permite guardar os blocos grandes de texto do arquivo comprimidos, sem
   depender de rede nem de API assincrona: o conteudo volta identico ao original. */
var ZL_=[3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258];
var ZLE_=[0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0];
var ZD_=[1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577];
var ZDE_=[0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13];
var ZCO_=[16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15];
function zHuff_(lens,n){
 var i,cnt=new Int32Array(16),sym=new Int32Array(n),off=new Int32Array(16),s=0;
 for(i=0;i<n;i++) cnt[lens[i]]++;
 cnt[0]=0;
 for(i=1;i<16;i++){ off[i]=s; s+=cnt[i]; }
 for(i=0;i<n;i++) if(lens[i]) sym[off[lens[i]]++]=i;
 return {cnt:cnt,sym:sym};
}
function zInflate_(u8,p){
 var out=new Uint8Array(1<<18),olen=0,bp=p<<3,fixL=null,fixD=null;
 function grow(n){ if(olen+n<=out.length) return; var s=out.length; while(s<olen+n) s*=2; var t=new Uint8Array(s); t.set(out.subarray(0,olen)); out=t; }
 function bits(n){ var v=0,i; for(i=0;i<n;i++){ v|=((u8[bp>>3]>>(bp&7))&1)<<i; bp++; } return v; }
 function dec(h){ var code=0,first=0,index=0,len=1,cnt;
  for(;len<16;len++){ code|=bits(1); cnt=h.cnt[len];
   if(code-cnt<first) return h.sym[index+(code-first)];
   index+=cnt; first=(first+cnt)<<1; code<<=1; }
  throw new Error("codigo huffman invalido");
 }
 for(;;){
  var last=bits(1), type=bits(2), i;
  if(type===0){
   bp=(bp+7)&~7; var q=bp>>3, len=u8[q]|(u8[q+1]<<8); q+=4;
   grow(len); out.set(u8.subarray(q,q+len),olen); olen+=len; bp=(q+len)<<3;
  } else {
   var lc,dc;
   if(type===1){
    if(!fixL){ var l=new Uint8Array(288);
     for(i=0;i<144;i++)l[i]=8; for(;i<256;i++)l[i]=9; for(;i<280;i++)l[i]=7; for(;i<288;i++)l[i]=8;
     fixL=zHuff_(l,288); var d=new Uint8Array(30); for(i=0;i<30;i++)d[i]=5; fixD=zHuff_(d,30); }
    lc=fixL; dc=fixD;
   } else if(type===2){
    var hl=bits(5)+257, hd=bits(5)+1, hc=bits(4)+4, cl=new Uint8Array(19);
    for(i=0;i<hc;i++) cl[ZCO_[i]]=bits(3);
    var clh=zHuff_(cl,19), lens=new Uint8Array(hl+hd), k=0, s2, r, v;
    while(k<hl+hd){ s2=dec(clh);
     if(s2<16) lens[k++]=s2;
     else if(s2===16){ v=lens[k-1]; r=3+bits(2); while(r--) lens[k++]=v; }
     else if(s2===17){ r=3+bits(3); while(r--) lens[k++]=0; }
     else { r=11+bits(7); while(r--) lens[k++]=0; }
    }
    lc=zHuff_(lens.subarray(0,hl),hl); dc=zHuff_(lens.subarray(hl),hd);
   } else throw new Error("bloco deflate invalido");
   for(;;){
    var sy=dec(lc);
    if(sy<256){ grow(1); out[olen++]=sy; }
    else if(sy===256) break;
    else { sy-=257; var ln=ZL_[sy]+bits(ZLE_[sy]); var ds=dec(dc); var dt=ZD_[ds]+bits(ZDE_[ds]);
     grow(ln); var from=olen-dt; for(var z=0;z<ln;z++) out[olen++]=out[from+z]; }
   }
  }
  if(last) break;
 }
 return out.subarray(0,olen);
}
function zB64_(s){
 var bin=atob(String(s||"").replace(/[^A-Za-z0-9+/=]/g,""));
 var u=new Uint8Array(bin.length);
 for(var i=0;i<bin.length;i++) u[i]=bin.charCodeAt(i);
 return u;
}
function zGunzip_(u8){
 if(!(u8.length>2 && u8[0]===0x1f && u8[1]===0x8b)) return u8;
 var flg=u8[3], p=10;
 if(flg&4){ p+=2+(u8[p]|(u8[p+1]<<8)); }
 if(flg&8){ while(u8[p]!==0) p++; p++; }
 if(flg&16){ while(u8[p]!==0) p++; p++; }
 if(flg&2){ p+=2; }
 return zInflate_(u8,p);
}
function GZS(b64){ return new TextDecoder("utf-8").decode(zGunzip_(zB64_(b64))); }

/* payload colunar: V1 (colunas fixas da base de indiretos) e V2 (colunas no proprio payload) */
function embDecodePay(txt){
 var nl=txt.indexOf("\n"), tag=txt.slice(0,nl), rest=txt.slice(nl+1), keys, i;
 if(tag==="V2"){ var nl2=rest.indexOf("\n"); keys=rest.slice(0,nl2).split(","); rest=rest.slice(nl2+1); }
 else keys=["d","c","l","tm","gm","gc","ncm","cod","ca"];
 var cut=rest.indexOf("\x1d"), dl=rest.slice(0,cut).split("\n"), body=rest.slice(cut+1), map={};
 for(i=0;i<dl.length;i++) if(dl[i]) map[dl[i].charAt(0)]=dl[i].slice(1);
 body=body.replace(/[\u0100-\u2FFF]/g,function(ch){ var v=map[ch]; return v===undefined?ch:v });
 var cols=body.split("\x1e"), parts=[];
 for(i=0;i<cols.length;i++) parts.push(cols[i].split("\n"));
 var n=parts[0].length, out=new Array(n), r, k, o, v2;
 for(r=0;r<n;r++){ o={};
  for(k=0;k<keys.length;k++){ v2=parts[k]?parts[k][r]:""; if(v2) o[keys[k]] = v2.indexOf("\x1f")<0 ? v2 : v2.replace(/\x1f/g,"\n"); }
  out[r]=o;
 }
 return out;
}

function embDecodeV1(txt){ return embDecodePay(txt) }

const CALRAW = GZS(CAT_RAW_GZ);
const CAL = {};
(function(){
 CALRAW.split("\n").forEach(function(bl){
  const p=bl.indexOf("="); if(p<0) return;
  const tag=bl.slice(0,p), T={};
  bl.slice(p+1).split(";").forEach(function(e){
   if(!e) return;
   const a=e.split(":"); const n=a.length; if(n<7) return;
   T[a.slice(0,n-6).join(":")]=[+a[n-6],+a[n-5],+a[n-4],+a[n-3],+a[n-2],+a[n-1]];
  });
  CAL[tag]=T;
 });
})();
const CAL_GEN={COMPONENTE:1,PRODUTO:1,PECA:1,CONJUNTO:1,MATERIAL:1,ITEM:1,ELEMENTO:1,ACESSORIO:1,DISPOSITIVO:1,ESPECIFICO:1};
function calNorm(s){ return String(s==null?"":s).normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/\s+/g," ").trim().toUpperCase() }
function calAttr(t,name){ const m=t.match(new RegExp(name+"[^:;]*:\\s*([^;]+)")); return m? m[1].trim() : "" }
const CAL_VEI=/(CAMINHAO|ONIBUS|VEICULO|CAMIONETE|PICKUP|CARRETA|AUTOMOVEL|CARRO|VAN|MOTOCICLETA|MOTO |REBOQUE|CAVALO MECANICO|HONDA|SCANIA|VOLVO|MERCEDES|IVECO|FORD|VOLKSWAGEN|TOYOTA|HILUX|S10)/;
const CAL_MAQ=/(TRATOR|COLHEDORA|COLHEITADEIRA|COLHED|PLANTADEIRA|PULVERIZADOR|MAQUINA|ESTEIRA|ESCAVADEIRA|CARREGADEIRA|MOTONIVELADORA|GRADE|SUBSOLADOR|SEMEADORA|TRANSBORDO|CARREGADORA|MOENDA|ROCADEIRA|AGRICOLA|IMPLEMENTO|GUINCHO|RETROESCAVADEIRA|TRATORES)/;
function calApc(a){ if(!a) return ""; try{ const _e=eqapCls(a); if(_e) return _e }catch(e){} if(CAL_VEI.test(a)) return "VEI"; if(CAL_MAQ.test(a)) return "MAQ"; return "EQP" }
/* ================================================================
   v7.3 - EQUIVALENCIA DE APLICACAO (APOIO, NAO REGRA)
   Tabela de apoio que liga o termo de aplicacao ou o modelo de
   maquina/veiculo/equipamento citado na descricao longa a uma
   familia (MAQ/VEI/EQP) e, opcionalmente, a um Grupo de Mercadoria
   e um Grupo de Compradores de apoio. Vem preenchida com o que a
   base 302 real mostra e e editavel pelo administrador na aba
   Regras e Catalogos. Onde a base 302 tem evidencia, a base decide
   e o apoio fica apenas registrado no painel de transparencia.
   ================================================================ */
const EQAP_KEY="mdm302.eqap.v73";
const EQAP_CLS={MAQ:"maquina ou implemento",VEI:"veiculo",EQP:"equipamento ou instalacao"};
const EQAP_EMB=[{"t":"TRATOR","cls":"MAQ","gm":"90002","gc":"MR1","n":610,"p":84,"pg":65,"pc":84,"o":"base"},{"t":"COLHEDORA","cls":"MAQ","gm":"90002","gc":"MR1","n":386,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"COLHEDORA CANA","cls":"MAQ","gm":"90002","gc":"MR1","n":292,"p":99,"pg":83,"pc":99,"o":"base"},{"t":"CAMINHAO","cls":"VEI","gm":"90003","gc":"MR1","n":191,"p":74,"pg":74,"pc":60,"o":"base"},{"t":"PROVA EXPLOSAO","cls":"EQP","gm":"90004","gc":"MC1","n":122,"p":99,"pg":99,"pc":80,"o":"base"},{"t":"MOTOR","cls":"","gm":"","gc":"MR1","n":106,"p":82,"pg":58,"pc":82,"o":"pendente"},{"t":"MOTOCICLETA","cls":"VEI","gm":"90003","gc":"MR1","n":72,"p":79,"pg":79,"pc":72,"o":"base"},{"t":"VEICULO","cls":"VEI","gm":"90003","gc":"","n":67,"p":79,"pg":79,"pc":46,"o":"pendente"},{"t":"COLHEITADEIRA","cls":"MAQ","gm":"90002","gc":"MR1","n":63,"p":100,"pg":71,"pc":100,"o":"base"},{"t":"PA CARREGADEIRA","cls":"MAQ","gm":"","gc":"MR1","n":59,"p":98,"pg":56,"pc":98,"o":"pendente"},{"t":"AUTOMOVEL","cls":"VEI","gm":"90003","gc":"MR1","n":56,"p":86,"pg":73,"pc":86,"o":"base"},{"t":"CARREGADEIRA","cls":"MAQ","gm":"90002","gc":"MR1","n":53,"p":92,"pg":60,"pc":92,"o":"base"},{"t":"CONTATOR","cls":"EQP","gm":"90001","gc":"MR1","n":53,"p":100,"pg":70,"pc":100,"o":"base"},{"t":"TRICANTER","cls":"EQP","gm":"90001","gc":"MR1","n":52,"p":94,"pg":81,"pc":94,"o":"base"},{"t":"CONCRETO","cls":"EQP","gm":"90001","gc":"MR1","n":51,"p":96,"pg":96,"pc":94,"o":"base"},{"t":"PULVERIZADOR","cls":"MAQ","gm":"","gc":"MR1","n":50,"p":100,"pg":52,"pc":100,"o":"pendente"},{"t":"ELETRICISTA","cls":"EQP","gm":"90711","gc":"EP1","n":47,"p":72,"pg":72,"pc":72,"o":"base"},{"t":"METAL","cls":"EQP","gm":"90001","gc":"MR1","n":44,"p":86,"pg":86,"pc":84,"o":"base"},{"t":"USO GERAL","cls":"","gm":"","gc":"","n":41,"p":39,"pg":32,"pc":39,"o":"pendente"},{"t":"ESGOTO","cls":"","gm":"","gc":"MR1","n":41,"p":73,"pg":44,"pc":73,"o":"pendente"},{"t":"INDUSTRIAL","cls":"","gm":"","gc":"","n":34,"p":56,"pg":56,"pc":50,"o":"pendente"},{"t":"ACO CARBONO","cls":"EQP","gm":"90001","gc":"MR1","n":33,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"PAREDE","cls":"EQP","gm":"90004","gc":"MC1","n":33,"p":85,"pg":85,"pc":61,"o":"base"},{"t":"BOMBA","cls":"EQP","gm":"90001","gc":"MR1","n":32,"p":97,"pg":78,"pc":97,"o":"base"},{"t":"COMPRESSOR","cls":"","gm":"","gc":"MR1","n":32,"p":72,"pg":44,"pc":72,"o":"pendente"},{"t":"ACO INOXIDAVEL","cls":"EQP","gm":"90001","gc":"MR1","n":31,"p":97,"pg":94,"pc":97,"o":"base"},{"t":"ROCADEIRA","cls":"MAQ","gm":"90001","gc":"MR1","n":31,"p":71,"pg":61,"pc":71,"o":"base"},{"t":"EIXO/EXTERNO","cls":"EQP","gm":"90001","gc":"MR1","n":31,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"VALVULA BORBOLETA","cls":"EQP","gm":"90001","gc":"MR1","n":28,"p":100,"pg":86,"pc":100,"o":"base"},{"t":"EIXO/INTERNO","cls":"EQP","gm":"90001","gc":"MR1","n":27,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"MOTOR PARTIDA","cls":"","gm":"","gc":"","n":26,"p":54,"pg":46,"pc":54,"o":"pendente"},{"t":"OPERACIONAL","cls":"","gm":"","gc":"UF1","n":26,"p":88,"pg":88,"pc":77,"o":"pendente"},{"t":"COBRIDOR CANA 2 LINHAS","cls":"MAQ","gm":"90002","gc":"MR1","n":25,"p":100,"pg":80,"pc":100,"o":"base"},{"t":"TOMBADOR","cls":"MAQ","gm":"90002","gc":"MR1","n":24,"p":100,"pg":71,"pc":100,"o":"base"},{"t":"DECANTER","cls":"EQP","gm":"90001","gc":"MR1","n":24,"p":92,"pg":88,"pc":92,"o":"base"},{"t":"JEANS","cls":"","gm":"","gc":"UF1","n":23,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"CAMINHONETE","cls":"VEI","gm":"90003","gc":"MR1","n":22,"p":86,"pg":86,"pc":77,"o":"base"},{"t":"CARRO","cls":"VEI","gm":"90003","gc":"","n":22,"p":82,"pg":82,"pc":55,"o":"pendente"},{"t":"TRATOR ESTEIRA","cls":"MAQ","gm":"","gc":"MR1","n":21,"p":100,"pg":52,"pc":100,"o":"pendente"},{"t":"ESTRUTURAS METALICAS","cls":"EQP","gm":"90004","gc":"MC1","n":21,"p":100,"pg":100,"pc":86,"o":"base"},{"t":"MOTOSSERRA","cls":"MAQ","gm":"90001","gc":"MR1","n":21,"p":95,"pg":71,"pc":95,"o":"base"},{"t":"IMPLEMENTO AGRICOLA","cls":"MAQ","gm":"90001","gc":"MR1","n":21,"p":95,"pg":76,"pc":95,"o":"base"},{"t":"IMPRESSORA","cls":"EQP","gm":"90707","gc":"MR1","n":20,"p":100,"pg":95,"pc":100,"o":"base"},{"t":"REBOQUE","cls":"VEI","gm":"90001","gc":"MR1","n":19,"p":89,"pg":79,"pc":89,"o":"base"},{"t":"JD6415","cls":"MAQ","gm":"90002","gc":"MR1","n":19,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"ACOPLAMENTO","cls":"","gm":"","gc":"MR1","n":18,"p":78,"pg":44,"pc":78,"o":"pendente"},{"t":"MADEIRA","cls":"","gm":"","gc":"MR1","n":18,"p":94,"pg":50,"pc":94,"o":"pendente"},{"t":"CENTRIFUGA ACUCAR BATELADA","cls":"EQP","gm":"90001","gc":"MR1","n":17,"p":100,"pg":94,"pc":100,"o":"base"},{"t":"AR CONDICIONADO","cls":"","gm":"","gc":"MR1","n":17,"p":88,"pg":47,"pc":88,"o":"pendente"},{"t":"PISO","cls":"EQP","gm":"90004","gc":"","n":17,"p":71,"pg":71,"pc":59,"o":"pendente"},{"t":"ALTERNADOR","cls":"","gm":"","gc":"MR1","n":17,"p":94,"pg":41,"pc":94,"o":"pendente"},{"t":"GERADOR","cls":"EQP","gm":"90001","gc":"MR1","n":17,"p":100,"pg":82,"pc":100,"o":"base"},{"t":"PORTA","cls":"","gm":"","gc":"","n":17,"p":53,"pg":53,"pc":53,"o":"pendente"},{"t":"ELEVADOR GRAO","cls":"MAQ","gm":"90002","gc":"MR1","n":17,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"EXTRUSORA","cls":"EQP","gm":"90001","gc":"MR1","n":16,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"AGUA FRIA","cls":"EQP","gm":"90004","gc":"MC1","n":16,"p":88,"pg":88,"pc":75,"o":"base"},{"t":"FAROL","cls":"","gm":"90003","gc":"","n":16,"p":63,"pg":63,"pc":50,"o":"pendente"},{"t":"MADEIRA/METAL","cls":"EQP","gm":"90004","gc":"MR1","n":15,"p":100,"pg":100,"pc":93,"o":"base"},{"t":"PICADOR","cls":"MAQ","gm":"","gc":"MR1","n":15,"p":87,"pg":40,"pc":87,"o":"pendente"},{"t":"CENTRIFUGA ACUCAR CONTINUA","cls":"EQP","gm":"90001","gc":"MR1","n":15,"p":87,"pg":87,"pc":87,"o":"base"},{"t":"ACO CARBONO/ACO INOXIDAVEL","cls":"EQP","gm":"90001","gc":"MR1","n":14,"p":100,"pg":86,"pc":100,"o":"base"},{"t":"EMPILHADEIRA","cls":"MAQ","gm":"","gc":"MR1","n":14,"p":79,"pg":43,"pc":79,"o":"pendente"},{"t":"SOPRADOR FULIGEM","cls":"EQP","gm":"90001","gc":"MR1","n":14,"p":100,"pg":86,"pc":100,"o":"base"},{"t":"INVERSOR FREQUENCIA","cls":"EQP","gm":"90001","gc":"MR1","n":14,"p":93,"pg":93,"pc":93,"o":"base"},{"t":"MAQUINA AGRICOLA","cls":"MAQ","gm":"90002","gc":"MR1","n":14,"p":100,"pg":71,"pc":100,"o":"base"},{"t":"BOMBA HIDROJATO","cls":"","gm":"","gc":"MR1","n":14,"p":79,"pg":43,"pc":79,"o":"pendente"},{"t":"TR218A","cls":"MAQ","gm":"90002","gc":"MR1","n":14,"p":100,"pg":93,"pc":100,"o":"base"},{"t":"L70F","cls":"MAQ","gm":"90002","gc":"MR1","n":14,"p":100,"pg":86,"pc":100,"o":"base"},{"t":"ELETRICO","cls":"","gm":"","gc":"","n":13,"p":54,"pg":46,"pc":54,"o":"pendente"},{"t":"COMPRESSOR AR","cls":"","gm":"90001","gc":"MR1","n":13,"p":100,"pg":62,"pc":100,"o":"base"},{"t":"ESMERILHADEIRA","cls":"","gm":"","gc":"MR1","n":13,"p":77,"pg":54,"pc":77,"o":"pendente"},{"t":"CARRETA","cls":"VEI","gm":"","gc":"MR1","n":13,"p":85,"pg":46,"pc":85,"o":"pendente"},{"t":"VALVULA","cls":"","gm":"90001","gc":"MR1","n":13,"p":92,"pg":62,"pc":92,"o":"base"},{"t":"RADIADOR","cls":"","gm":"","gc":"MR1","n":13,"p":85,"pg":46,"pc":85,"o":"pendente"},{"t":"UNIPORT","cls":"MAQ","gm":"90002","gc":"MR1","n":13,"p":92,"pg":77,"pc":92,"o":"base"},{"t":"ONIBUS","cls":"VEI","gm":"90001","gc":"MR1","n":12,"p":83,"pg":83,"pc":83,"o":"base"},{"t":"FREIO","cls":"VEI","gm":"90003","gc":"MR1","n":12,"p":92,"pg":92,"pc":92,"o":"base"},{"t":"METAIS FERROSOS/NAO FERROSOS","cls":"EQP","gm":"90001","gc":"MR1","n":12,"p":100,"pg":100,"pc":92,"o":"base"},{"t":"MECANICO","cls":"","gm":"","gc":"UF1","n":12,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"AUTOMOTIVO","cls":"","gm":"","gc":"","n":12,"p":58,"pg":50,"pc":58,"o":"pendente"},{"t":"DAGUA SUPERFICIES METALICAS","cls":"EQP","gm":"90004","gc":"MC1","n":11,"p":91,"pg":91,"pc":91,"o":"base"},{"t":"PORCELANATO","cls":"EQP","gm":"90004","gc":"MR1","n":11,"p":91,"pg":82,"pc":91,"o":"base"},{"t":"ROLAMENTO","cls":"","gm":"","gc":"MR1","n":11,"p":64,"pg":36,"pc":64,"o":"pendente"},{"t":"ALVENARIA","cls":"EQP","gm":"90004","gc":"MR1","n":11,"p":91,"pg":82,"pc":91,"o":"base"},{"t":"DISTRIBUIDOR","cls":"EQP","gm":"90001","gc":"MR1","n":11,"p":91,"pg":73,"pc":91,"o":"base"},{"t":"MULTIUSO","cls":"","gm":"","gc":"","n":11,"p":45,"pg":45,"pc":45,"o":"pendente"},{"t":"OCUPACIONAL","cls":"EQP","gm":"90711","gc":"EP1","n":10,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"ESCAVADEIRA","cls":"MAQ","gm":"","gc":"MR1","n":10,"p":100,"pg":50,"pc":100,"o":"pendente"},{"t":"RETROESCAVADEIRA","cls":"MAQ","gm":"","gc":"MR1","n":10,"p":100,"pg":40,"pc":100,"o":"pendente"},{"t":"CENTRIFUGA","cls":"","gm":"90001","gc":"MR1","n":10,"p":100,"pg":60,"pc":100,"o":"base"},{"t":"IMPACTOS/AGENTES ABRASIVOS/ESCORIANTES","cls":"EQP","gm":"90711","gc":"EP1","n":10,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"PINTURA","cls":"EQP","gm":"90004","gc":"MR1","n":10,"p":90,"pg":70,"pc":90,"o":"base"},{"t":"APARENTE","cls":"","gm":"90004","gc":"","n":10,"p":60,"pg":60,"pc":50,"o":"pendente"},{"t":"BLOCO COMANDO","cls":"MAQ","gm":"90002","gc":"MR1","n":10,"p":100,"pg":90,"pc":100,"o":"base"},{"t":"FREIO DIANTEIRO","cls":"VEI","gm":"90003","gc":"MR1","n":9,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"ALVENARIA/CONCRETO/MADEIRA","cls":"EQP","gm":"90004","gc":"MC1","n":9,"p":89,"pg":89,"pc":78,"o":"base"},{"t":"JARDIM","cls":"","gm":"","gc":"","n":9,"p":56,"pg":56,"pc":56,"o":"pendente"},{"t":"TANQUE","cls":"","gm":"","gc":"MR1","n":9,"p":89,"pg":56,"pc":89,"o":"pendente"},{"t":"BALANCA","cls":"","gm":"90001","gc":"MR1","n":9,"p":78,"pg":67,"pc":78,"o":"base"},{"t":"SECADOR","cls":"","gm":"90001","gc":"MR1","n":9,"p":100,"pg":67,"pc":100,"o":"base"},{"t":"PROVA TEMPO/GAS/VAPOR/PO","cls":"EQP","gm":"90004","gc":"MR1","n":9,"p":100,"pg":100,"pc":67,"o":"base"},{"t":"INDUSTRIAL/AUTOMOTIVO","cls":"","gm":"","gc":"MR1","n":9,"p":100,"pg":44,"pc":100,"o":"pendente"},{"t":"HONDA/NXR160 BROS ESDD","cls":"VEI","gm":"90003","gc":"MR1","n":8,"p":88,"pg":88,"pc":63,"o":"base"},{"t":"CAMINHAO/ONIBUS","cls":"VEI","gm":"90003","gc":"AU1","n":8,"p":100,"pg":100,"pc":63,"o":"base"},{"t":"BOMBA HIDRAULICA","cls":"MAQ","gm":"90002","gc":"MR1","n":8,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"CAMBIO","cls":"VEI","gm":"90003","gc":"MR1","n":8,"p":100,"pg":100,"pc":63,"o":"base"},{"t":"MOTONIVELADORA","cls":"MAQ","gm":"90002","gc":"MR1","n":8,"p":88,"pg":63,"pc":88,"o":"base"},{"t":"LIMPADOR PARABRISA","cls":"","gm":"","gc":"MR1","n":8,"p":88,"pg":50,"pc":88,"o":"pendente"},{"t":"EMBREAGEM","cls":"","gm":"","gc":"MR1","n":8,"p":75,"pg":50,"pc":75,"o":"pendente"},{"t":"RODA","cls":"","gm":"","gc":"MR1","n":8,"p":63,"pg":38,"pc":63,"o":"pendente"},{"t":"GUINDASTE","cls":"EQP","gm":"90001","gc":"MR1","n":8,"p":100,"pg":88,"pc":100,"o":"base"},{"t":"FERRO","cls":"","gm":"90004","gc":"MC1","n":8,"p":63,"pg":63,"pc":63,"o":"base"},{"t":"DETECTOR GASES","cls":"","gm":"90001","gc":"MR1","n":7,"p":86,"pg":86,"pc":86,"o":"base"},{"t":"RODOVIARIO","cls":"VEI","gm":"90003","gc":"MR1","n":7,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"HONDA/NXR150 BROS KS","cls":"VEI","gm":"90003","gc":"AU1","n":7,"p":86,"pg":86,"pc":71,"o":"base"},{"t":"MOTOR ARRANQUE","cls":"","gm":"90002","gc":"MR1","n":7,"p":100,"pg":71,"pc":100,"o":"base"},{"t":"LIXADEIRA","cls":"","gm":"","gc":"","n":7,"p":57,"pg":57,"pc":57,"o":"pendente"},{"t":"MINICARREGADEIRA","cls":"MAQ","gm":"90002","gc":"MR1","n":7,"p":86,"pg":71,"pc":86,"o":"base"},{"t":"LANTERNA","cls":"","gm":"90003","gc":"MR1","n":7,"p":100,"pg":100,"pc":86,"o":"base"},{"t":"MADEIRA/MASSA","cls":"","gm":"90004","gc":"MC1","n":7,"p":71,"pg":71,"pc":71,"o":"base"},{"t":"MINI CARREGADEIRA","cls":"MAQ","gm":"","gc":"MR1","n":7,"p":86,"pg":43,"pc":86,"o":"pendente"},{"t":"ARGAMASSA","cls":"","gm":"90151","gc":"MR1","n":7,"p":71,"pg":71,"pc":71,"o":"base"},{"t":"LIMPEZA GERAL","cls":"","gm":"90705","gc":"MR1","n":7,"p":71,"pg":71,"pc":71,"o":"base"},{"t":"BOP SONDA WEST CAPRICORN","cls":"","gm":"90001","gc":"MR1","n":7,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"MOTORREDUTOR","cls":"","gm":"90001","gc":"MR1","n":6,"p":100,"pg":67,"pc":100,"o":"base"},{"t":"CILINDRO","cls":"","gm":"90001","gc":"MR1","n":6,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"CILINDRO DIRECAO","cls":"","gm":"90002","gc":"MR1","n":6,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"HONDA VERMELHA 003508","cls":"VEI","gm":"90003","gc":"AU1","n":6,"p":67,"pg":67,"pc":67,"o":"base"},{"t":"ESCAVADEIRA HIDRAULICA","cls":"MAQ","gm":"90001","gc":"MR1","n":6,"p":100,"pg":67,"pc":100,"o":"base"},{"t":"PLANTADEIRA","cls":"MAQ","gm":"","gc":"MR1","n":6,"p":100,"pg":50,"pc":100,"o":"pendente"},{"t":"CARRINHO TRIPPER","cls":"","gm":"90001","gc":"MR1","n":6,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"PORTAO","cls":"","gm":"90004","gc":"MR1","n":6,"p":100,"pg":67,"pc":100,"o":"base"},{"t":"ATUADOR","cls":"","gm":"90001","gc":"MR1","n":6,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"CABECOTE","cls":"","gm":"90001","gc":"MR1","n":6,"p":100,"pg":83,"pc":100,"o":"base"},{"t":"RETARDANTE CHAMA","cls":"","gm":"","gc":"UF1","n":6,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"MOTOCANA","cls":"MAQ","gm":"90002","gc":"MR1","n":5,"p":100,"pg":80,"pc":100,"o":"base"},{"t":"TURBINA VAPOR SIEMENS","cls":"","gm":"90001","gc":"MR1","n":5,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"MOENDA","cls":"MAQ","gm":"90001","gc":"MR1","n":5,"p":100,"pg":80,"pc":100,"o":"base"},{"t":"KOMBI","cls":"VEI","gm":"90003","gc":"AU1","n":5,"p":80,"pg":80,"pc":80,"o":"base"},{"t":"MOTOR ELETRICO","cls":"","gm":"90001","gc":"MR1","n":5,"p":80,"pg":80,"pc":60,"o":"base"},{"t":"MOTOR DIESEL","cls":"","gm":"","gc":"MR1","n":5,"p":60,"pg":40,"pc":60,"o":"pendente"},{"t":"MOINHO","cls":"","gm":"90001","gc":"MR1","n":5,"p":80,"pg":60,"pc":80,"o":"base"},{"t":"PISTAO","cls":"","gm":"90002","gc":"MR1","n":5,"p":100,"pg":60,"pc":100,"o":"base"},{"t":"CARCACA","cls":"","gm":"90001","gc":"MR1","n":5,"p":80,"pg":80,"pc":80,"o":"base"},{"t":"CABINE","cls":"","gm":"90001","gc":"MR1","n":5,"p":80,"pg":60,"pc":80,"o":"base"},{"t":"TOYOTA HILUX","cls":"VEI","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"L70F/L90F","cls":"MAQ","gm":"","gc":"MR1","n":30,"p":97,"pg":53,"pc":97,"o":"pendente"},{"t":"AC193","cls":"MAQ","gm":"","gc":"MR1","n":35,"p":89,"pg":49,"pc":89,"o":"pendente"},{"t":"UNIDADE POTENCIA HIDRAULICA","cls":"","gm":"90001","gc":"MR1","n":5,"p":100,"pg":100,"pc":100,"o":"base"},{"t":"ENCARREGADO OPERACIONAL","cls":"","gm":"","gc":"UF1","n":5,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"TRANSMISSAO","cls":"","gm":"","gc":"","n":3,"p":100,"pg":67,"pc":100,"o":"pendente"},{"t":"MANCAL","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"ACOPLAMENTO/ENGRENAGENS","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"SISTEMA HIDRAULICO","cls":"","gm":"","gc":"","n":3,"p":67,"pg":67,"pc":67,"o":"pendente"},{"t":"EIXO PLANETARIO","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"PLATAFORMA DE CORTE","cls":"","gm":"","gc":"","n":3,"p":100,"pg":67,"pc":100,"o":"pendente"},{"t":"SKIDDER FLORESTAL","cls":"MAQ","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"FELLER FLORESTAL","cls":"MAQ","gm":"","gc":"","n":3,"p":100,"pg":67,"pc":100,"o":"pendente"},{"t":"TRATOR/PA CARREGADEIRA","cls":"MAQ","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"ESCAVADEIRA/TRATOR/CARREGADEIRA","cls":"MAQ","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"TURBO GERADOR","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"TROCADOR CALOR","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"BOMBA CENTRIFUGA","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"DESINTEGRADOR","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"REDLER","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"MOTONIVELADORA 620G","cls":"MAQ","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"BOMBA INJETORA","cls":"","gm":"","gc":"","n":3,"p":67,"pg":67,"pc":67,"o":"pendente"},{"t":"INTERRUPTOR FREIO","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"AMORTECEDOR","cls":"","gm":"","gc":"","n":3,"p":67,"pg":67,"pc":67,"o":"pendente"},{"t":"SUSPENSAO","cls":"","gm":"","gc":"","n":3,"p":67,"pg":67,"pc":67,"o":"pendente"},{"t":"FREIO TRASEIRO","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":67,"o":"pendente"},{"t":"MOLA TRASEIRA","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":67,"o":"pendente"},{"t":"CAIXA DIRECAO","cls":"","gm":"","gc":"","n":3,"p":67,"pg":67,"pc":67,"o":"pendente"},{"t":"COBRIDOR DE CANA 2 LINHAS","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"GRADE NIVELADORA AGRICOLA","cls":"MAQ","gm":"","gc":"","n":4,"p":100,"pg":50,"pc":100,"o":"pendente"},{"t":"BAZUCA AGRICOLA","cls":"MAQ","gm":"","gc":"","n":4,"p":100,"pg":75,"pc":100,"o":"pendente"},{"t":"CILINDRO HIDRAULICO","cls":"","gm":"","gc":"","n":4,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"CILINDRO ELEVACAO","cls":"","gm":"","gc":"","n":4,"p":100,"pg":75,"pc":100,"o":"pendente"},{"t":"CILINDRO INCLINACAO","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"VALVULA PNEUMATICA","cls":"","gm":"","gc":"","n":4,"p":100,"pg":50,"pc":100,"o":"pendente"},{"t":"CONSERTO PNEU","cls":"","gm":"","gc":"","n":4,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"PEDAL PARTIDA","cls":"","gm":"","gc":"","n":4,"p":100,"pg":100,"pc":75,"o":"pendente"},{"t":"TRANSPORTADOR CORREIA","cls":"","gm":"","gc":"","n":4,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"CORREIA ELEVADORA","cls":"","gm":"","gc":"","n":4,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"TURBO FILTRO","cls":"","gm":"","gc":"","n":4,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"TURBO BOMBA","cls":"","gm":"","gc":"","n":4,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"TURBINA","cls":"","gm":"","gc":"","n":4,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"IMPRESSORA LASER","cls":"","gm":"","gc":"","n":4,"p":100,"pg":100,"pc":75,"o":"pendente"},{"t":"AGROTOXICO","cls":"","gm":"","gc":"","n":4,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"MOTOR INDUSTRIAL","cls":"","gm":"","gc":"","n":4,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"CILINDRO PNEUMATICO","cls":"","gm":"","gc":"","n":4,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"VARREDEIRA","cls":"","gm":"","gc":"","n":3,"p":67,"pg":67,"pc":67,"o":"pendente"},{"t":"ESCRITORIO","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":67,"o":"pendente"},{"t":"LAPIS","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":67,"o":"pendente"},{"t":"PAPEL","cls":"","gm":"","gc":"","n":3,"p":67,"pg":67,"pc":67,"o":"pendente"},{"t":"SANITARIO","cls":"","gm":"","gc":"","n":3,"p":67,"pg":67,"pc":67,"o":"pendente"},{"t":"LIMPEZA","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":67,"o":"pendente"},{"t":"PARAFUSADEIRA","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"PROJETOR","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"SENSOR","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"EVAPORADOR","cls":"","gm":"","gc":"","n":3,"p":100,"pg":67,"pc":100,"o":"pendente"},{"t":"CONDENSADORA","cls":"","gm":"","gc":"","n":3,"p":100,"pg":67,"pc":100,"o":"pendente"},{"t":"CERAMICA ESMALTADA","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":67,"o":"pendente"},{"t":"ARMARIO","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"TORNEIRA","cls":"","gm":"","gc":"","n":3,"p":100,"pg":67,"pc":100,"o":"pendente"},{"t":"CONSTRUCAO","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":67,"o":"pendente"},{"t":"CIMENTO","cls":"","gm":"","gc":"","n":3,"p":67,"pg":67,"pc":67,"o":"pendente"},{"t":"ARGAMASSA/CONCRETO","cls":"","gm":"","gc":"","n":3,"p":100,"pg":100,"pc":100,"o":"pendente"},{"t":"LAVATORIO","cls":"","gm":"90151","gc":"MR1","n":5,"p":100,"pg":100,"pc":100,"o":"base"}];
const EQAP_MIN=5;
let EQAP=[], EQAP_IX={};
function eqapNorm(s){ return calNorm(s).replace(/[^A-Z0-9 \/\.]/g," ").replace(/\s+/g," ").trim() }
function eqapIndex(){ EQAP_IX={}; EQAP.forEach(function(e){ const k=eqapNorm(e.t); if(k) EQAP_IX[k]=e }) }
function eqapDefault(){ return EQAP_EMB.map(function(o){ return Object.assign({},o) }) }
function eqapLoad(){
 try{ const r=JSON.parse(localStorage.getItem(EQAP_KEY)||"null"); EQAP=(r&&r.length)? r : eqapDefault() }
 catch(e){ EQAP=eqapDefault() }
 eqapIndex();
}
function eqapSave(){ try{ localStorage.setItem(EQAP_KEY, JSON.stringify(EQAP)) }catch(e){} eqapIndex() }
function eqapFind(s,min){
 const a=eqapNorm(s); if(!a) return null;
 if(EQAP_IX[a]) return EQAP_IX[a];
 const m=min||4; let bk="";
 for(const k in EQAP_IX){ if(k.length<m) continue; if(a.indexOf(k)>=0 && k.length>bk.length) bk=k }
 return bk? EQAP_IX[bk] : null;
}
function eqapCls(a){ const e=eqapFind(a,4); return (e&&e.cls)? e.cls : "" }
function eqapEv(e){
 if(!e) return "";
 if(e.o==="admin") return "preenchida pelo administrador";
 if(!e.n) return "sem contagem registrada";
 return e.n+" material(is) equivalentes na base 302 (concordancia "+(e.pg||0)+"% no 663 e "+(e.pc||0)+"% no 627)";
}
function eqapPend(){ return EQAP.filter(function(e){ return !e.gm || !e.gc }) }
eqapLoad();
function calTok(s){ return calNorm(s).replace(/[^A-Z0-9 \/]/g," ").split(/\s+/).filter(Boolean) }
function calFeat(longa,curta,desc,ncm){
 const L=calNorm(longa), C=calNorm(curta), D=calNorm(desc);
 const peca=calAttr(L,"NOME DA PECA"), mod=calAttr(L,"NOME MODIFICADOR"), tip=calAttr(L,"TIPO"), apl=calAttr(L,"APLICACAO");
 let S="";
 [peca, L.split(";")[0], C, D].forEach(function(src){
  if(S||!src) return;
  const tk=calTok(src);
  for(let i=0;i<tk.length;i++){ if(tk[i].length>=3 && !CAL_GEN[tk[i]]){ S=tk[i]; break } }
 });
 const dig=String(ncm==null?"":ncm).replace(/\D/g,"");
 return {S:S, APC:calApc(apl), TIP:(calTok(tip)[0]||""), MODT:(calTok(mod)[0]||""), N4:dig.slice(0,4), N6:dig.slice(0,6), N8:dig.slice(0,8)};
}
function calKeys(f){
 return [["A8",f.S+"|"+f.APC+"|"+f.N8],["A6",f.S+"|"+f.APC+"|"+f.N6],["A4",f.S+"|"+f.APC+"|"+f.N4],
         ["T",f.S+"|"+f.TIP+"|"+f.APC],["B6",f.S+"|"+f.N6],["B4",f.S+"|"+f.N4],
         ["M",f.S+"|"+f.MODT+"|"+f.APC],["C",f.S+"|"+f.APC],["D",f.S],
         ["E",f.N6+"|"+f.APC],["F",f.N4+"|"+f.APC],["G",f.N4]];
}
function calVote(longa,curta,desc,ncm){
 const f=calFeat(longa,curta,desc,ncm), out={f:f};
 [["gm",0,1,CAL_GM],["gc",2,3,CAL_GC],["tm",4,5,CAL_TM]].forEach(function(cfg){
  const sc={}, ev={}; let tot=0;
  calKeys(f).forEach(function(kv){
   const T=CAL[kv[0]]; if(!T) return;
   const e=T[kv[1]]; if(!e) return;
   const code=cfg[3][e[cfg[1]]]; if(!code) return;
   const p=e[cfg[2]]/100, n=e[5];
   const w=CAL_W[kv[0]]*p*Math.min(1,Math.log1p(n)/Math.log(12));
   sc[code]=(sc[code]||0)+w; tot+=w;
   if(!ev[code] || ev[code].w<w) ev[code]={lvl:kv[0],key:kv[1],p:p,n:n,w:w};
  });
  let bk="", bw=0;
  for(const k in sc){ if(sc[k]>bw){ bw=sc[k]; bk=k } }
  out[cfg[0]] = bk? {k:bk,p:ev[bk].p,n:ev[bk].n,lvl:ev[bk].lvl,key:ev[bk].key,share:(tot? bw/tot:0)} : {k:"",p:0,n:0,lvl:"",key:"",share:0};
 });
 return out;
}
function calTxt(c){ if(!c||!c.k) return ""; return "calibracao pela base 302 ("+(CAL_LVL[c.lvl]||c.lvl)+": "+c.n+" material(is) equivalentes, "+Math.round(c.p*100)+"% no mesmo codigo)" }

function classify(desc, curta, longa, ncm){
 /* ===== v10: ENTRADA UNICA PONDERADA =====
    as descricoes da mesma linha descrevem o MESMO material e nunca sao tratadas como
    materiais distintos: entram juntas e cada uma vota com o peso da sua hierarquia.
    hierarquia de confianca: 1 descricao longa  2 descricao curta  3 NCM  4 descricao do material.
    digitacao manual (um campo unico) continua sendo validada isoladamente. */
 const _L=String(longa==null?"":longa).trim(), _C=String(curta==null?"":curta).trim(), _D=String(desc==null?"":desc).trim();
 const PW={longa:1.00, curta:0.82, ncm:0.66, desc:0.55};
 const dparts=[{t:_L,w:PW.longa,f:"descricao longa"},{t:_C,w:PW.curta,f:"descricao curta"},{t:_D,w:PW.desc,f:"descricao do material"}].filter(function(p){ return p.t });
 const key = desc||curta||longa||"";
 const uni = uniDesc([_L,_C,_D]);
 const t = toks(uni||key);
 const NCMV = ncmClean(ncm), NCH = NCMV? ncmLook(NCMV) : null;
 const srcHit = {base:0, head:0, ncm:0, exemplo:0, fb:0};
 /* substantivo principal define o bloco de candidatos, BM25 ranqueia dentro dele */
 const hd = pickHead(dparts, t);
 const nbBm = bm25Rank(dparts, hd, 20);
 const nb = nbBm.map(function(x){ return {r:x.r, s:x.s, eff:x.s, pw:1, sc:x.sc, soft:x.soft} });
 const nbTop = nb.slice().sort(function(a,b){ return b.s-a.s });

 const vt={},vg={},vc={}; let wsum=0;

 /* 1) vizinhos reais da base carregada */
 for(let i=0;i<nb.length;i++){
  const x=nb[i]; if(x.s<0.30) continue;
  const w=Math.pow(x.s,3)*3*(x.pw||1);
  wsum+=w;
  if(x.s>=0.55) srcHit.base=Math.max(srcHit.base,x.s);
  if(x.r.tm) vt[x.r.tm]=(vt[x.r.tm]||0)+w;
  if(x.r.gm) vg[x.r.gm]=(vg[x.r.gm]||0)+w;
  if(x.r.gc) vc[x.r.gc]=(vc[x.r.gc]||0)+w;
 }
 /* 1b) NCM da linha: peso abaixo das descricoes e acima da descricao do material */
 if(NCH){
  const w=NCH.w*PW.ncm*1.6; wsum+=w; srcHit.ncm=1;
  if(NCH.tm && codeOk("tm",NCH.tm)) vt[NCH.tm]=(vt[NCH.tm]||0)+w*0.8;
  if(NCH.gm && codeOk("gm",NCH.gm)) vg[NCH.gm]=(vg[NCH.gm]||0)+w;
  if(NCH.gc && codeOk("gc",NCH.gc)) vc[NCH.gc]=(vc[NCH.gc]||0)+w*0.95;
 }
 /* 1c) exemplos da base de conhecimento (painel de exemplos): sinal curado, peso alto */
 let exHit=null;
 (S.ex||[]).forEach(function(e){
  if(!e || !e.d) return;
  const sm=Math.max(diceSet(tri(uni||key),tri(e.d)), diceSet(tri(key),tri(e.d)));
  if(sm>0.68){
   const w=11*sm; wsum+=w;
   if(!exHit || exHit.s<sm) exHit={e:e, s:sm};
   srcHit.exemplo=Math.max(srcHit.exemplo,sm);
   if(e.tm && codeOk("tm",e.tm)) vt[e.tm]=(vt[e.tm]||0)+w;
   if(e.gm && codeOk("gm",e.gm)) vg[e.gm]=(vg[e.gm]||0)+w;
   if(e.gc && codeOk("gc",e.gc)) vc[e.gc]=(vc[e.gc]||0)+w;
  }
 });
 /* 2) reforco das correcoes aprovadas por voce (peso maior que o historico bruto) */
 for(let i=0;i<S.fb.length;i++){
  const f=S.fb[i];
  const s=Math.max(diceSet(tri(key),tri(f.d)), diceSet(tri(uni||key),tri(f.d)));
  if(s>0.72){ const w=9*s; wsum+=w; srcHit.fb=Math.max(srcHit.fb,s);
   if(f.tm) vt[f.tm]=(vt[f.tm]||0)+w;
   if(f.gm) vg[f.gm]=(vg[f.gm]||0)+w;
   if(f.gc) vc[f.gc]=(vc[f.gc]||0)+w;
  }
 }
 /* 3) substantivo principal: prior mais forte do modelo (real, tirado da base 302) */
 const hl = headLook(hd);
 if(hl && (hl.n||0)>=6) srcHit.head=Math.min(1, Math.log(1+hl.n)/4);
 if(hl){
  const w=Math.log(1+hl.n)*4.3*(hl.soft?0.6:1)*(hl.src==="curado"?1.3:1);
  if(hl.tm.k) vt[hl.tm.k]=(vt[hl.tm.k]||0)+w*hl.tm.p;
  if(hl.gm.k) vg[hl.gm.k]=(vg[hl.gm.k]||0)+w*hl.gm.p;
  if(hl.gc.k) vc[hl.gc.k]=(vc[hl.gc.k]||0)+w*hl.gc.p;
  wsum+=w;
 }
 /* 4) modificadores da descricao (lexico de tokens) */
 let lexHit=0;
 const uq=Array.from(new Set(t));
 const posW=function(tk){const p=t.indexOf(tk); return p===0?1.2:(p===1?1.5:(p===2?1.1:0.9))};
 uq.forEach(function(tk){
  const w0=idf(tk)*posW(tk);
  const L=lexLook(tk); if(!L) return;
  const soft = L.soft?0.7:1;
  lexHit += L.soft?0.6:1;
  if(L.src==="base"){
   const lb=L.e, bt2=best(lb.tm), bg2=best(lb.gm), bc2=best(lb.gc);
   const w=w0*Math.log(1+lb.n)*0.55*soft;
   if(bt2.k) vt[bt2.k]=(vt[bt2.k]||0)+w*bt2.p;
   if(bg2.k) vg[bg2.k]=(vg[bg2.k]||0)+w*bg2.p;
   if(bc2.k) vc[bc2.k]=(vc[bc2.k]||0)+w*bc2.p;
   wsum+=w;
  } else {
   const e=L.e, w=w0*Math.log(1+e.n)*0.5*soft;
   vt[e.tm]=(vt[e.tm]||0)+w*e.ptm;
   vg[e.gm]=(vg[e.gm]||0)+w*e.pgm;
   vc[e.gc]=(vc[e.gc]||0)+w*e.pgc;
   wsum+=w;
  }
 });

 /* FILTRO DE CATALOGO: evidencia com codigo fora do catalogo oficial sai do pool de sugestao */
 const vgV={}, vcV={}, vgS={}, vcS={}, discGm={}, discGc={};
 for(const k in vg){ if(codeOk("gm",k)) vgV[k]=(vgV[k]||0)+vg[k];
  else { discGm[k]=(discGm[k]||0)+vg[k]; const s=sucOf("gm",k); if(s) vgS[s]=(vgS[s]||0)+vg[k]*0.6 } }
 for(const k in vc){ if(codeOk("gc",k)) vcV[k]=(vcV[k]||0)+vc[k];
  else { discGc[k]=(discGc[k]||0)+vc[k]; const s=sucOf("gc",k); if(s) vcS[s]=(vcS[s]||0)+vc[k]*0.6 } }
 const discGmTop=best(discGm).k||"", discGcTop=best(discGc).k||"";
 const vtV={}, discTm={};
 for(const k in vt){ if(codeOk("tm",k)) vtV[k]=(vtV[k]||0)+vt[k]; else discTm[k]=(discTm[k]||0)+vt[k] }
 const bt=best(vtV), bg=best(vgV), bc=best(vcV);
 let tm=bt.k||"", gm=bg.k||"", gc=bc.k||"";
 if(GM_IMOB.indexOf(gm)>=0) gm="90001"; /* GM de imobilizado nao pertence ao 302 */

 const maxSim = nbTop.length? nbTop[0].s : 0;
 const strongHits = uq.filter(function(x){const L=lexLook(x); return L&&!L.soft&&L.e&&(L.e.n||0)>=8}).length;
 const cov = uq.length? strongHits/uq.length : 0;
 const headSup = hl? Math.min(1, Math.log(1+hl.n)/3.5)*(hl.soft?0.55:1) : 0;
 const ev = Math.min(1, 0.65*headSup + 0.35*cov + (S.base.length? 0.35*maxSim : 0));
 let cTm=fieldConf(bt.p,hl,"tm",maxSim,ev), cGm=fieldConf(bg.p,hl,"gm",maxSim,ev), cGc=fieldConf(bc.p,hl,"gc",maxSim,ev);

 /* prior dominante: substantivo com historico grande e alta pureza nao e derrubado por modificador */
 if(hl && !hl.soft && hl.n>=30){
  if(hl.tm.k && hl.tm.p>=0.75 && tm!==hl.tm.k){ tm=hl.tm.k; cTm=Math.max(cTm,fieldConf(hl.tm.p,hl,"tm",maxSim,ev)) }
  if(hl.gm.k && codeOk("gm",hl.gm.k) && hl.gm.p>=0.75 && gm!==hl.gm.k && GM_IMOB.indexOf(hl.gm.k)<0){ gm=hl.gm.k; cGm=Math.max(cGm,fieldConf(hl.gm.p,hl,"gm",maxSim,ev)) }
  if(hl.gc.k && codeOk("gc",hl.gc.k) && hl.gc.p>=0.75 && gc!==hl.gc.k){ gc=hl.gc.k; cGc=Math.max(cGc,fieldConf(hl.gc.p,hl,"gc",maxSim,ev)) }
 }

 /* regras determinísticas */
 let ruleTxt="";
 if(HEAD_NLAG.indexOf(hd)>=0 && tm!=="NLAG"){ tm="NLAG"; cTm=Math.max(cTm,0.9);
  ruleTxt="regra: "+hd+" nao e item estocavel/palpavel, Tipo ajustado para NLAG"; }
 const priorForte = !!(hl && !hl.soft && hl.n>=30 && hl.gm.p>=0.75);
 if(HEAD_PART.indexOf(hd)>=0 && !(priorForte && GM_MANUT.indexOf(hl.gm.k)<0)){
  if(GM_MANUT.indexOf(gm)<0){
   let alt=null,bv=-1;
   GM_MANUT.forEach(function(k){ if(codeOk("gm",k) && (vgV[k]||0)>bv){bv=vgV[k]||0;alt=k} });
   if(!(bv>0)) alt=null;
   if(!alt && hl && hl.gm.k && codeOk("gm",hl.gm.k) && GM_MANUT.indexOf(hl.gm.k)>=0) alt=hl.gm.k;
   if(!alt) alt="90001";
   ruleTxt="regra: item principal e componente ("+hd+"), Grupo de Mercadoria ajustado de "+gm+" para "+alt;
   gm=alt; cGm=Math.max(cGm,0.8);
  }
  if(tm==="HIBE" && hl && hl.tm.k==="ERSA" && hl.tm.p>=0.6){
   tm="ERSA"; cTm=Math.max(cTm,0.82);
   ruleTxt=(ruleTxt?ruleTxt+" | ":"")+"regra: "+hd+" e peca de reposicao no padrao da base, Tipo ajustado para ERSA";
  }
 }

 /* EPI: substantivo de protecao individual + termo de seguranca => 90711 */
 const EPI_HEAD=["LUVA","OCULOS","BOTINA","BOTA","MASCARA","PROTETOR","CAPACETE","RESPIRADOR","AVENTAL","TALABARTE","COLETE","CONCHA","JALECO","PERNEIRA","CAPUZ"];
 const EPI_SIG=["SEGURANCA","EPI","VAQUETA","NITRILICA","NITRILICO","PFF2","PFF1","AURICULAR","BALACLAVA","RASPA","ANTICORTE","SOLDADOR","CA","ABAFADOR"];
 if(EPI_HEAD.indexOf(hd)>=0 && t.some(function(x){return EPI_SIG.indexOf(x)>=0})){
  if(gm!=="90711"){ ruleTxt=(ruleTxt?ruleTxt+" | ":"")+"regra: "+hd+" com indicacao de protecao individual, Grupo de Mercadoria ajustado de "+gm+" para 90711 (EPI'S)"; }
  gm="90711"; cGm=Math.max(cGm,0.88);
  if(tm!=="HIBE"){ tm="HIBE"; cTm=Math.max(cTm,0.85) }
 }

 /* gemeo praticamente identico na base: adota a classificacao ja cadastrada */
 let twin=null;
 if(maxSim>=0.95 && nb[0].r.tm){
  twin=nb[0].r;
  const ttm=codeSafe("tm",twin.tm), tgm=codeSafe("gm",twin.gm), tgc=codeSafe("gc",twin.gc);
  if(ttm) tm=ttm; if(tgm) gm=tgm; if(tgc) gc=tgc;
  cTm=ttm?0.97:cTm; cGm=tgm?0.97:cGm; cGc=tgc?0.97:cGc;
  ruleTxt=(ruleTxt?ruleTxt+" | ":"")+"material praticamente identico ja cadastrado ("+(twin.cod?twin.cod+" ":"")+twin.d+"): classificacao copiada do cadastro existente";
 }
 if(!wsum){ cTm=cGm=cGc=0.15 }

 /* ===== v11: GRUPO DE MERCADORIA POR MAIORIA PONDERADA DOS VIZINHOS =====
    candidatos filtrados pelo substantivo, ranqueados por BM25 (tokens raros pesam mais);
    a decisao vem da concordancia entre os k vizinhos e abstem quando ela e baixa */
 let agGm=0, apoio="";
 if(!twin && nbBm.length){
  const vv={}; let tot=0, n=0;
  nbBm.slice(0,12).forEach(function(x){
   if(!x.r.gm || !codeOk("gm",x.r.gm) || GM_IMOB.indexOf(x.r.gm)>=0) return;
   const w=x.sc||1; vv[x.r.gm]=(vv[x.r.gm]||0)+w; tot+=w; n++;
  });
  const bgv=best(vv);
  if(bgv.k && tot>0){
   agGm=vv[bgv.k]/tot;
   const cc=Math.min(0.95, 0.45+0.50*agGm)*(nbBm[0].soft?0.85:1);
   if(agGm>=0.5 || !codeOk("gm",gm)){
    if(gm!==bgv.k) ruleTxt=(ruleTxt?ruleTxt+" | ":"")+"maioria dos "+n+" materiais semelhantes do substantivo "+(hd||"-")+": "+Math.round(agGm*100)+"% do peso em "+bgv.k;
    gm=bgv.k; cGm=Math.max(cGm*0.6, cc);
   } else if(gm===bgv.k){ cGm=Math.max(cGm,cc) }
  }
  if(agGm && agGm<0.45 && !priorForte && !(exHit && exHit.s>=0.88)){
   gm=""; cGm=0.25;
   ruleTxt=(ruleTxt?ruleTxt+" | ":"")+"concordancia entre semelhantes abaixo do minimo ("+Math.round(agGm*100)+"%): Grupo de Mercadoria fica para REVISAR";
  }
  apoio = nbBm.slice(0,3).map(function(x){ return (x.r.cod?x.r.cod+" ":"")+String(x.r.d||"").slice(0,34)+" ["+(x.r.gm||"-")+"/"+(x.r.gc||"-")+"] "+Math.round(x.s*100)+"%" }).join(" ; ");
 }
 /* base carregada, nenhum semelhante e nenhuma regra: melhor abster do que arriscar codigo */
 if(!nbBm.length && (S.base||[]).length && !twin && !priorForte && !ruleTxt && !NCH && srcHit.head<0.2 && !(exHit && exHit.s>=0.80)){
  gm=""; cGm=0.22;
  ruleTxt="nenhum material semelhante na base 302 e nenhuma regra aplicavel: Grupo de Mercadoria e Grupo de Compradores ficam para REVISAR";
 }
 /* NCM entra como desempate e veto, nunca como voto isolado */
 if(NCH && NCH.gm && codeOk("gm",NCH.gm)){
  if(gm && NCH.gm===gm) cGm=Math.min(0.97, cGm+0.06*(1-cGm));
  else if(gm && agGm && agGm<0.70) cGm=Math.max(0.22, cGm*0.85);
  else if(!gm){ gm=NCH.gm; cGm=Math.min(0.55, 0.35+0.20*(NCH.lvl>=6?1:0.6));
   ruleTxt=(ruleTxt?ruleTxt+" | ":"")+"sem semelhante na base: "+NCH.src+" define o Grupo de Mercadoria - confirmar" }
 }

 /* ===== TRAVA DE CATALOGO - GRUPO DE COMPRADORES SO PODE SAIR DO 627 =====
    ordem: gemeo cadastrado / regra de familia / historico valido do substantivo /
    evidencia valida / predominancia valida do grupo de mercadoria / sucessor / em branco (REVISAR) */
 const txtRule=" "+norm(desc+" "+(curta||"")+" "+(longa||""))+" ";
 let gcVia = (twin && codeOk("gc",gc)) ? "copiado do cadastro existente, codigo valido do 627" : "";
 if(!(twin && codeOk("gc",gc))){
  /* ===== v11: GRUPO DE COMPRADORES DERIVADO DO GRUPO DE MERCADORIA =====
     o GC nunca mais e estimado direto do texto em paralelo ao GM - era a origem das
     divergencias. Ordem: exemplo curado / correcao sua / maioria dos semelhantes dentro
     do GM / regra de familia compativel / predominancia do GM / historico do substantivo. */
  const dg = deriveGc(gm, nbBm, exHit, S.fb, hd, txtRule, uni||key);
  gc = dg.k; cGc = dg.k? Math.min(dg.c, Math.max(0.28, cGm)) : 0.25;
  gcVia = dg.via;
 }
 if(gc && !codeOk("gc",gc)){
  /* trava absoluta: codigo descontinuado nunca e sugerido nem exportado - tenta sucessor valido,
     depois predominancia do grupo de mercadoria, depois predominancia geral */
  const _s=sucOf("gc",gc), _fg=gcFromGm(gm), _gg=gcGlobal();
  if(_s && codeOk("gc",_s)){ gcVia="sucessor valido de "+gc+" no catalogo 627 - confirmar"; gc=_s; cGc=Math.min(Math.max(cGc,0.50),0.62) }
  else if(_fg){ gcVia="predominancia valida no Grupo de Mercadoria "+gm+" (codigo anterior descontinuado) - confirmar"; gc=_fg; cGc=Math.min(Math.max(cGc,0.45),0.60) }
  else if(_gg.k){ gcVia="predominancia geral do 302 restrita ao 627 (codigo anterior descontinuado) - confirmar"; gc=_gg.k; cGc=Math.min(0.55,0.38+0.17*_gg.p) }
  else { gc=""; cGc=Math.min(cGc,0.30); gcVia="nenhum codigo do catalogo 627 com confianca suficiente - REVISAR" }
 }

 /* ===== TRAVA DE CATALOGO - GRUPO DE MERCADORIA SO PODE SAIR DO 663 ===== */
 let gmVia = codeOk("gm",gm) ? (twin?"copiado do cadastro existente, codigo valido do 663":"evidencia do historico com codigo valido do 663") : "";
 if(!codeOk("gm",gm)){
  const rgm=ruleGm(txtRule), s=sucOf("gm",discGmTop)||best(vgS).k;
  if(rgm){ gm=rgm.k; cGm=Math.max(cGm,0.70); gmVia="regra de familia ("+rgm.why+") -> "+rgm.k+" "+GM_OFF[rgm.k]+", codigo do catalogo 663" }
  else if(s){ gm=s; cGm=Math.min(Math.max(cGm,0.5),0.62); gmVia="sucessor valido de "+(discGmTop||"codigo descontinuado")+" no catalogo 663 - confirmar" }
  else { gm=""; cGm=0.25; gmVia="nenhum codigo do catalogo 663 com confianca suficiente - REVISAR" }
 }
 /* ===== v7.2: CALIBRACAO PELA BASE 302 REAL (fonte primaria de evidencia) ===== */
 let _cal=null;
 try{ _cal=calVote(_L||"", _C||"", _D||"", ncm||"") }catch(e){ _cal=null }
 if(_cal && !twin){
  if(_cal.gm.k && codeOk("gm",_cal.gm.k) && cGm<0.90){
   if(_cal.gm.k!==gm){
    gmVia=calTxt(_cal.gm)+" -> "+_cal.gm.k+" "+(GM_OFF[_cal.gm.k]||"")+(codeOk("gm",gm)? " (no lugar de "+gm+" "+(GM_OFF[gm]||"")+", que vinha do prior geral)" : "");
    gm=_cal.gm.k;
   } else { gmVia=calTxt(_cal.gm)+" confirma "+gm+" | "+gmVia }
   cGm=Math.max(cGm, 0.55+0.35*_cal.gm.p);
  }
  if(_cal.gc.k && codeOk("gc",_cal.gc.k) && cGc<0.90){
   if(_cal.gc.k!==gc){
    gcVia=calTxt(_cal.gc)+" -> "+_cal.gc.k+" "+(GC_OFF[_cal.gc.k]||"")+(codeOk("gc",gc)? " (no lugar de "+gc+" "+(GC_OFF[gc]||"")+")" : "");
    gc=_cal.gc.k;
   } else { gcVia=calTxt(_cal.gc)+" confirma "+gc+" | "+gcVia }
   cGc=Math.max(cGc, 0.55+0.35*_cal.gc.p);
  }
 }
 /* ===== v7: REGRA DE FAMILIA/SUBSTANTIVO SOBRE CODIGO GENERICO =====
    sem base 302 carregada, o prior do historico puxa quase tudo para os grupos genericos de
    manutencao (90001/90002/90004). Quando a descricao tem familia explicita na planilha oficial
    (expediente, EPI, limpeza, instalacao eletrica, construcao, lubrificante, informatica...),
    a regra de familia prevalece sobre o generico e o Grupo de Compradores e rederivado para
    ficar coerente com o novo Grupo de Mercadoria. Sugestao de apoio - a decisao final e do usuario. */
 (function(){
  const GEN_GM = ["90001","90002","90004"];
  if(_cal && _cal.gm.k) return;  /* a base real ja decidiu */
  if(twin) return;
  const semBase = !(S.base && S.base.length);
  /* rotulos de atributo da descricao longa (EMBALAGEM:, APLICACAO:, DIMENSOES:) nao sao familia */
  const txtF = String(txtRule||"").replace(/[A-Z][A-Z\/ ]{2,24}:/g," ");
  /* peca de reposicao continua no grupo de manutencao, mesmo citando oleo, agua ou combustivel */
  const _hd = (typeof hd!=="undefined" && hd)? String(hd) : "";
  const PART_HEAD = /^(FILTRO|BOMBA|MANGUEIRA|RETENTOR|VALVULA|CILINDRO|ROLAMENTO|CORREIA|PARAFUSO|PORCA|ARRUELA|MOTOR|SENSOR|ENGRENAGEM|JUNTA|VEDACAO|PISTAO|MANCAL|POLIA|CORRENTE|BUCHA|EIXO|KIT|COMPRESSOR|REDUTOR|ROLETE|ELEMENTO|MANIFOLD|TUBULACAO|ROTOR|ESTATOR)/;
  if(PART_HEAD.test(_hd)) return;
  /* a familia e lida do substantivo e das primeiras palavras da descricao, nao dos atributos do final */
  const toks = txtF.split(/\s+/).filter(Boolean).slice(0,5).join(" ");
  const alvo = (_hd+" "+toks).trim();
  const rgm2 = ruleGm(alvo);
  if(!rgm2 || !codeOk("gm",rgm2.k) || rgm2.k===gm) return;
  if(GEN_GM.indexOf(gm)<0 && codeOk("gm",gm)) return;
  if(!semBase && cGm>=0.80) return;
  const antigo = gm;
  gm = rgm2.k;
  cGm = Math.max(cGm, semBase? 0.76 : 0.72);
  gmVia = "regra de familia ("+rgm2.why+") -> "+rgm2.k+" "+(GM_OFF[rgm2.k]||"")+", codigo da planilha oficial"+
          (antigo? " (no lugar do generico "+antigo+" "+(GM_OFF[antigo]||"")+" vindo do historico)" : "");
  const ng = famGc(gm, alvo);
  if(ng.k && ng.k!==gc){ gc = ng.k; cGc = Math.min(Math.max(cGc,0.55), ng.c); gcVia = ng.via }
 })();
 /* ===== v7.1: AJUSTE DE ESPECIFICIDADE DO GRUPO DE COMPRADORES =====
    o historico do 302 usa MC1 como grupo amplo e ele acabava vencendo por volume. Havendo um
    codigo especifico do 627 coerente com a familia do material e com o Grupo de Mercadoria
    escolhido (EPI -> EP1, expediente -> ME2, lubrificante -> LB1, informatica -> SH1/IS1/TI1,
    peca de reposicao -> MR1 ...), ele prevalece. Continua sendo sugestao: a decisao e do usuario. */
 (function(){
  const AMPLO = ["MC1"];
  if(_cal && _cal.gc.k) return;  /* a base real ja decidiu */
  if(!gc || AMPLO.indexOf(gc)<0 || !codeOk("gm",gm)) return;
  const _hd2=(typeof hd!=="undefined" && hd)? String(hd) : "";
  const t2=String(txtRule||"").replace(/[A-Z][A-Z\/ ]{2,24}:/g," ");
  const alvo2=(_hd2+" "+t2.split(/\s+/).filter(Boolean).slice(0,6).join(" ")).trim();
  const sp=famGc(gm, alvo2);
  if(sp.k && AMPLO.indexOf(sp.k)<0 && sp.c>=0.66 && codeOk("gc",sp.k)){
   const ant=gc;
   gc=sp.k; cGc=Math.min(Math.max(cGc,0.55), sp.c);
   gcVia="ajuste de especificidade: "+sp.via+" - o historico do 302 traz "+ant+" "+(GC_OFF[ant]||"")+" como grupo amplo para esse substantivo; sugerimos confirmar qual dos dois vale na sua governanca";
  }
 })();
 /* ===== TRAVA + FALLBACK DO TIPO DE MATERIAL: so codigo do 650, e nunca em branco quando o catalogo existe =====
    ordem: gemeo / evidencia valida / peca de reposicao / nao estocavel / tipo padrao valido do catalogo */
 if(apoio) gmVia = (gmVia? gmVia+" | " : "")+"apoio: "+apoio;
 let tmVia = codeOk("tm",tm) ? (twin?"copiado do cadastro existente, codigo valido do 650":"evidencia do historico com codigo valido do 650") : "";
 if(!codeOk("tm",tm)){
  const has=function(c){return TIPOS.some(function(t){return t[0]===c})};
  const pecaSig = (hd && HEAD_PART.indexOf(hd)>=0) || /(ROLAMENTO|RETENTOR|CORREIA|ENGRENAGEM|MANCAL|BUCHA|PINHAO|VALVULA|BOMBA|MOTOR |SENSOR|CONTATOR|PALHETA|CILINDRO|PISTAO|ROLETE|CHAVETA|ACOPLAMENTO|POLIA|EIXO|CAMISA DE|KIT REPARO|VEDACAO)/.test(txtRule);
  const naoEstSig = /(SERVICO|MAO DE OBRA|LOCACAO|ALUGUEL|DIARIA|SUCATA|FRETE|MANUTENCAO PREDIAL|TREINAMENTO|SOFTWARE|LICENCA)/.test(txtRule);
  if(pecaSig && has("ERSA")){ tm="ERSA"; cTm=Math.max(cTm,0.72); tmVia="regra: componente de maquina/equipamento -> ERSA do catalogo 650" }
  else if(naoEstSig && has("NLAG")){ tm="NLAG"; cTm=Math.max(cTm,0.70); tmVia="regra: item nao estocavel/servico -> NLAG do catalogo 650" }
  else {
   const dflt=tmDefault();
   if(dflt){ tm=dflt; cTm=Math.max(Math.min(cTm,0.62),0.5); tmVia="tipo padrao valido do catalogo 650 ("+dflt+") - evidencia insuficiente para outro tipo" }
   else { tm=""; cTm=0.25; tmVia="catalogo 650 vazio - carregue a planilha de tipos na area administrativa" }
  }
  if(Object.keys(discTm).length) tmVia+=" | codigo de tipo fora do catalogo ignorado: "+Object.keys(discTm).join(", ");
 }
 const discAll = {gm:Object.keys(discGm), gc:Object.keys(discGc)};

 const imob = checkImob(key, longa);
 /* ===== v10: a confianca reflete a COMBINACAO das fontes, nao um campo isolado =====
    fontes independentes: base 302, substantivo do modelo, NCM, exemplo da base de
    conhecimento e correcao aprovada. Cada fonte concordante reforca a decisao. */
 const _agree=(srcHit.base>=0.55?1:0)+(srcHit.head>=0.2?1:0)+(srcHit.ncm?1:0)+(srcHit.exemplo>=0.68?1:0)+(srcHit.fb>=0.72?1:0);
 const _bst=_agree>=2? Math.min(0.18, 0.06*(_agree-1)) : 0;
 const _up=function(v){ return (v>=0.30 && _bst>0)? Math.min(0.97, v + _bst*(1-v)) : v };
 if(exHit && exHit.s>=0.90){
  /* exemplo curado praticamente identico decide a categoria */
  if(exHit.e.tm && codeOk("tm",exHit.e.tm)){ if(exHit.e.tm!==tm){ tm=exHit.e.tm; tmVia="exemplo da base de conhecimento: "+exHit.e.d } cTm=Math.max(cTm,0.93) }
  if(exHit.e.gm && codeOk("gm",exHit.e.gm)){ if(exHit.e.gm!==gm){ gm=exHit.e.gm; gmVia="exemplo da base de conhecimento: "+exHit.e.d } cGm=Math.max(cGm,0.93) }
  if(exHit.e.gc && codeOk("gc",exHit.e.gc)){ if(exHit.e.gc!==gc){ gc=exHit.e.gc; gcVia="exemplo da base de conhecimento: "+exHit.e.d } cGc=Math.max(cGc,0.93) }
 }
 else if(exHit && exHit.s>=0.80){
  if(exHit.e.tm===tm) cTm=Math.max(cTm,0.85);
  if(exHit.e.gm===gm) cGm=Math.max(cGm,0.85);
  if(exHit.e.gc===gc) cGc=Math.max(cGc,0.85);
 }
 cTm=_up(cTm); cGm=_up(cGm); cGc=_up(cGc);
 let conf = 0.30*cTm + 0.40*cGm + 0.30*cGc;
 if(imob.flag && imob.lvl==="alto") conf*=0.9;
 conf = Math.max(0.05, Math.min(0.99, conf));

 /* divergencia entre as tres descricoes */
 let div="";
 if(desc && curta){ const s=diceSet(tri(desc),tri(curta)); if(s<0.45) div="Descricao Curta divergente da Descricao do Material"; }
 if(desc && longa){ const s2=diceSet(tri(desc),tri(longa)); if(s2<0.20 && norm(longa).length>12) div=(div?div+" e ":"")+"Descricao Longa pouco aderente"; }

 /* duplicidade */
 const dups = nb.filter(function(x){return x.s>=0.60}).slice(0,5).map(function(x){
  const p=Math.round(x.s*100);
  return {pct:p, d:x.r.d, cod:x.r.cod||"", tm:x.r.tm, gm:x.r.gm, gc:x.r.gc,
          risk: p>=90?"provavel":(p>=75?"possivel":"improvavel")};
 });

 /* por que o modelo decidiu assim */
 const why=[];
 if(hl) why.push("substantivo principal "+(hl.via?hd+" tratado como "+hl.via:hd)+": "+hl.n+(hl.n>1?" materiais":" material")+" no historico da base 302, "+
   Math.round(hl.tm.p*100)+"% "+hl.tm.k+" / "+Math.round(hl.gm.p*100)+"% "+hl.gm.k+" / "+Math.round(hl.gc.p*100)+"% "+hl.gc.k+
   (hl.src==="base"?" (base carregada)":" (modelo embarcado)"));
 else why.push("substantivo principal \""+hd+"\" nao existe no historico: classificacao por semelhanca, revisar antes de integrar");
 if(nb.length && nb[0].s>=0.55) why.push("material mais parecido na base: "+nb[0].r.d+" ("+Math.round(nb[0].s*100)+"%)");
 const strong = uq.filter(function(x){const L=lexLook(x); return L&&!L.soft&&L.e&&L.e.n>=8}).slice(0,4);
 if(strong.length) why.push("termos que pesaram: "+strong.join(", "));
 if(ruleTxt) why.push(ruleTxt);
 if(discAll.gc.length) why.push("codigo(s) de Grupo de Compradores descontinuado(s) ignorado(s): "+discAll.gc.map(function(k){return k+(GC_LEG[k]?" ("+GC_LEG[k]+")":"")}).join(", ")+" - nao consta(m) no catalogo 627");
 if(discAll.gm.length) why.push("codigo(s) de Grupo de Mercadoria descontinuado(s) ignorado(s): "+discAll.gm.map(function(k){return k+(GM_LEG[k]?" ("+GM_LEG[k]+")":"")}).join(", ")+" - nao consta(m) no catalogo 663");
 /* ===== v7.3: APOIO POR EQUIVALENCIA DE APLICACAO (apoio, nunca regra) =====
    a tabela de equivalencia so entra depois de tudo: onde a base 302 tem evidencia, a base
    decide e o apoio fica apenas registrado no painel de transparencia. Sem evidencia da base
    e com confianca baixa, ela entra como desempate - sugestao de apoio, a decisao e do usuario. */
 let _eqLog=null;
 (function(){
  try{
   const _tt=calNorm((_L||"")+" ; "+(_C||"")+" ; "+(_D||""));
   const _ap=calAttr(_tt,"APLICACAO")||calAttr(_tt,"APLICACOES")||"";
   const _mo=calAttr(_tt,"MODELO")||"";
   const _e=(_ap? eqapFind(_ap,4):null) || (_mo? eqapFind(_mo,4):null) || eqapFind(_tt,6);
   if(!_e) return;
   const _via="equivalencia de aplicacao \"" + _e.t + "\"" + (_e.cls? " ("+(EQAP_CLS[_e.cls]||_e.cls)+")" : "") + " - " + eqapEv(_e);
   _eqLog={t:_e.t,cls:_e.cls||"",gm:_e.gm||"",gc:_e.gc||"",o:_e.o||"base",n:_e.n||0,pg:_e.pg||0,pc:_e.pc||0,txt:_via,apGm:"",apGc:"",campo:(_ap?"APLICACAO":(_mo?"MODELO":"descricao"))};
   /* a base so "decide" quando a evidencia dela e firme; evidencia fraca abre espaco para o apoio */
   const _fGm=!!(_cal&&_cal.gm&&_cal.gm.k), _fGc=!!(_cal&&_cal.gc&&_cal.gc.k);
   const _bGm=_fGm && (((_cal.gm.p||0)>=0.55 && (_cal.gm.n||0)>=5) || (_cal.gm.share||0)>=0.60);
   const _bGc=_fGc && (((_cal.gc.p||0)>=0.55 && (_cal.gc.n||0)>=5) || (_cal.gc.share||0)>=0.60);
   const _fr=(_e.o!=="admin" && (_e.n||0)<EQAP_MIN);
   if(!_e.gm) _eqLog.apGm="sem Grupo de Mercadoria de apoio preenchido para este termo - a linha esta disponivel para preencher no painel administrativo";
   else if(!codeOk("gm",_e.gm)) _eqLog.apGm="apoio ignorado: "+_e.gm+" nao consta no catalogo 663 em uso";
   else if(_fr || twin || _bGm || cGm>=0.60) _eqLog.apGm="apoio nao aplicado ("+(_fr?"a equivalencia tem menos de 5 materiais de evidencia na base - preencha a linha no painel para ela virar apoio":(twin?"cadastro existente":(_bGm?"a base 302 tem evidencia firme para este material":"confianca ja suficiente")))+"); registrado apenas aqui";
   else { const _a=gm; gm=_e.gm; cGm=Math.min(Math.max(cGm,0.52),0.62);
    gmVia="apoio por "+_via+" -> "+_e.gm+" "+(GM_OFF[_e.gm]||"")+(_a?" (no lugar de "+_a+" "+(GM_OFF[_a]||"")+")":"")+" - sugestao de apoio, sugerimos confirmar";
    _eqLog.apGm="aplicado como apoio: a base 302 nao tinha evidencia firme para este material"+(_fGm?" (evidencia fraca: "+Math.round((_cal.gm.p||0)*100)+"% em "+(_cal.gm.n||0)+" material(is))":""); }
   if(!_e.gc) _eqLog.apGc="sem Grupo de Compradores de apoio preenchido para este termo - a linha esta disponivel para preencher no painel administrativo";
   else if(!codeOk("gc",_e.gc)) _eqLog.apGc="apoio ignorado: "+_e.gc+" nao consta no catalogo 627 em uso";
   else if(_fr || twin || _bGc || cGc>=0.60) _eqLog.apGc="apoio nao aplicado ("+(_fr?"a equivalencia tem menos de 5 materiais de evidencia na base - preencha a linha no painel para ela virar apoio":(twin?"cadastro existente":(_bGc?"a base 302 tem evidencia firme para este material":"confianca ja suficiente")))+"); registrado apenas aqui";
   else { const _b=gc; gc=_e.gc; cGc=Math.min(Math.max(cGc,0.52),0.62);
    gcVia="apoio por "+_via+" -> "+_e.gc+" "+(GC_OFF[_e.gc]||"")+(_b?" (no lugar de "+_b+" "+(GC_OFF[_b]||"")+")":"")+" - sugestao de apoio, sugerimos confirmar";
    _eqLog.apGc="aplicado como apoio: a base 302 nao tinha evidencia firme para este material"+(_fGc?" (evidencia fraca: "+Math.round((_cal.gc.p||0)*100)+"% em "+(_cal.gc.n||0)+" material(is))":""); }
  }catch(e){ _eqLog=null }
 })();
 /* ===== v7.4: CLASSE DE AVALIACAO (campo novo, mesma logica de evidencia dos demais) ===== */
 let imoc=null;
 try{ imoc=imoConf({desc:key, curta:_C, longa:_L, ncm:NCMV, tm:tm, gm:gm, twin:!!twin}) }catch(e){ imoc=null }
 let ca="", cCa=0, caVia="", caInfo=null;
 try{
  const _sub=(hd&&hd.tk)? hd.tk : "";
  caInfo=caPick(_D||_C||"", _sub, gm, tm, twin||null, gc, (function(){ try{ var _tk=toks(uni||key); return _tk[1]||"" }catch(e){ return "" } })());
  ca=codeOk("ca",caInfo.c)? caInfo.c : "";
  cCa=ca? caInfo.conf : 0;
  caVia=caInfo.via + (ca? " - sugestao, sugerimos confirmar" : (CA_COLSTATE===false? " - a planilha da base carregada nao traz a coluna CLASSE DE AVALIACAO: sem essa coluna nao ha evidencia para derivar a classe, por isso o campo sai como REVISAR" : (Object.keys(CA_OFF||{}).length? " - nenhum material parecido na base tem classe preenchida, por isso o campo sai como REVISAR, sem chute" : " - catalogo de Classe de Avaliacao vazio: o campo sai como REVISAR ate a coluna vir na base ou as classes serem cadastradas na area do administrador")));
 }catch(e){ ca=""; cCa=0; caVia="nao foi possivel avaliar a classe de avaliacao para este material - campo marcado como REVISAR"; caInfo={lista:"-",kind:"revisar",via:caVia} }
 /* ===== v7: regra GM (+GC) -> Classe de Avaliacao, do painel do administrador ===== */
 let _carInfo=null, dupCurta=null;
 try{
  if(CAR.rows.length){
   const _cr=carPick(gm,gc);
   if(_cr && _cr.ca){ ca=_cr.ca; cCa=Math.max(cCa||0,_cr.conf); caVia=_cr.via+" | "+caVia; caInfo=caInfo||{}; caInfo.kind="regra"; caInfo.lista=_cr.lista; _carInfo={kind:"regra",ca:_cr.ca}; }
   else if(_cr && _cr.cands && _cr.cands.length){ ca=""; cCa=0; caVia="regra da planilha: o Grupo de Mercadoria "+gm+" admite mais de uma Classe de Avaliacao ("+_cr.cands.join(", ")+"): campo em aberto para escolha, sem chute | "+caVia; caInfo=caInfo||{}; caInfo.kind="revisar"; _carInfo={kind:"ambigua",cands:_cr.cands}; }
   else { caVia="a planilha de regra nao cobre o Grupo de Mercadoria "+(gm||"nao definido")+": a classe abaixo e apenas sugestao pela evidencia da base, nao regra | "+caVia; _carInfo={kind:"sem-regra"}; }
  }
 }catch(e){ _carInfo=null }
 try{ dupCurta=curDupFind((typeof _C!=="undefined"&&_C)?_C:(curta||"")) }catch(e){ dupCurta=null }
 if(caVia) why.push("Classe de Avaliacao: "+caVia);
 if(gcVia) why.push("Grupo de Compradores: "+gcVia);
 if(gmVia) why.push("Grupo de Mercadoria: "+gmVia);
 if(tmVia) why.push("Tipo de Material: "+tmVia);
 if(!S.base.length) why.push("planilha da base ainda nao carregada: duplicidade limitada ao modelo embarcado");
 const just = why.join(" | ");

 /* regra central: o campo classificado guarda SEMPRE o codigo puro */
 tm=codeOf("tm",tm); gm=codeOf("gm",gm); gc=codeOf("gc",gc);
 if(!codeOk("tm",tm)) tm="";
 if(!codeOk("gm",gm)) gm="";
 if(!codeOk("gc",gc)) gc="";
 const rev=[]; if(!tm) rev.push("Tipo de Material"); if(!gm) rev.push("Grupo de Mercadoria"); if(!gc) rev.push("Grupo de Compradores");

 return {
  desc:key, curta:curta||"", longa:longa||"",
  tm:tm, gm:gm, gc:gc, gci: codeOf("gci", tm==="NLAG"?FIXOS.gciNlag:FIXOS.gci), setor: codeOf("setor",FIXOS.setor), aprovado: codeOf("aprovado",FIXOS.aprovado),
  ref: nbTop.length? {cod:(nbTop[0].r.cod||""), d:nbTop[0].r.d, tm:codeOf("tm",nbTop[0].r.tm), gm:codeOf("gm",nbTop[0].r.gm), gc:codeOf("gc",nbTop[0].r.gc), sim:Math.round(nbTop[0].s*100)} : null,
  ncm: NCMV||"", ncmVia: NCH? NCH.src : "", agGm: Math.round((agGm||0)*100),
  top3: nbBm.slice(0,3).map(function(x){ return {cod:(x.r.cod||""), d:x.r.d||"", c:x.r.c||"", ca:x.r.ca||"", gm:x.r.gm||"", gc:x.r.gc||"", sim:Math.round(x.s*100)} }), fontes: srcHit, exemplo: exHit? {d:exHit.e.d, sim:Math.round(exHit.s*100)} : null,
  head: hl? {tk:hd, n:hl.n, src:hl.src, soft:!!hl.soft, via:hl.via||"", tm:hl.tm.k, ptm:hl.tm.p, gm:hl.gm.k, pgm:hl.gm.p, gc:hl.gc.k, pgc:hl.gc.p} : {tk:hd, n:0, src:"nenhum"},
  regra: ruleTxt||"",
  conf:conf, cTm:cTm, cGm:cGm, cGc:cGc, hd:hd, hn:(hl?hl.n:0),
  imob:imob, imoc:imoc, dups:dups, div:div, just:just, why:why,
  via:{tm:tmVia, gm:gmVia, gc:gcVia, ca:caVia}, disc:discAll, rev:rev, eqap:_eqLog, ca:ca, cCa:cCa, caInfo:caInfo, dupCurta:dupCurta, carInfo:_carInfo,
  flag: ((imoc && (imoc.estado==="imob" || imoc.conflito)) || imob.flag || conf<0.55 || (dups[0]&&dups[0].pct>=75) || !!div || rev.length>0)
 };
}

/* ================= PERSISTENCIA (IndexedDB) ================= */
const DB={n:"mdm302",v:1,h:null};
function db(){return new Promise(function(res,rej){
 if(DB.h) return res(DB.h);
 const rq=indexedDB.open(DB.n,DB.v);
 rq.onupgradeneeded=function(e){const d=e.target.result;
  if(!d.objectStoreNames.contains("kv")) d.createObjectStore("kv");
 };
 rq.onsuccess=function(e){DB.h=e.target.result;res(DB.h)};
 rq.onerror=function(){rej(rq.error)};
})}
function kvSet(k,v){return db().then(function(d){return new Promise(function(res,rej){
 const t=d.transaction("kv","readwrite");t.objectStore("kv").put(v,k);t.oncomplete=function(){res(true)};t.onerror=function(){rej(t.error)};
})})}
function kvGet(k){return db().then(function(d){return new Promise(function(res,rej){
 const t=d.transaction("kv","readonly");const r=t.objectStore("kv").get(k);
 r.onsuccess=function(){res(r.result)};r.onerror=function(){rej(r.error)};
})})}

function saveBase(){try{kvSet("baseStamp",EMB_STAMP)}catch(e){}return kvSet("base",S.base.map(function(r){return {d:r.d,c:r.c,l:r.l,tm:r.tm,gm:r.gm,gc:r.gc,gci:r.gci,sa:r.sa,cod:r.cod,ca:r.ca||"",ncm:r.ncm||""}}))}
function saveFb(){return kvSet("fb",S.fb)}
function saveLog(){return kvSet("log",S.log.slice(-4000))}

/* ================= LEITURA DE ARQUIVO (sem dependencia externa) ================= */
function csvParse(txt){
 const rows=[];let row=[],cell="",q=false;
 txt=txt.replace(/^\uFEFF/,"");
 const sep=(txt.split("\n")[0].split(";").length>txt.split("\n")[0].split(",").length)?";":",";
 for(let i=0;i<txt.length;i++){
  const c=txt[i];
  if(q){ if(c==="\""){ if(txt[i+1]==="\""){cell+="\"";i++} else q=false } else cell+=c }
  else if(c==="\""){q=true}
  else if(c===sep){row.push(cell);cell=""}
  else if(c==="\n"){row.push(cell);cell="";if(row.some(function(x){return String(x).trim()!==""}))rows.push(row);row=[]}
  else if(c!=="\r"){cell+=c}
 }
 row.push(cell); if(row.some(function(x){return String(x).trim()!==""}))rows.push(row);
 return rows;
}
function unxml(s){
 return String(s==null?"":s)
  .replace(/&#x([0-9a-fA-F]+);/g,function(_,h){return String.fromCodePoint(parseInt(h,16))})
  .replace(/&#(\d+);/g,function(_,d){return String.fromCodePoint(+d)})
  .replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/&quot;/g,"\"").replace(/&apos;/g,"'").replace(/&#39;/g,"'").replace(/&amp;/g,"&");
}
async function inflateRaw(buf){
 if(typeof DecompressionStream==="undefined") throw new Error("nav");
 const ds=new DecompressionStream("deflate-raw");
 const st=new Blob([buf]).stream().pipeThrough(ds);
 return new Uint8Array(await new Response(st).arrayBuffer());
}
async function xlsxRead(ab){
 const u=new Uint8Array(ab), dv=new DataView(ab);
 let eo=-1;
 for(let i=u.length-22;i>=Math.max(0,u.length-66000);i--){ if(dv.getUint32(i,true)===0x06054b50){eo=i;break} }
 if(eo<0) throw new Error("zip");
 const cnt=dv.getUint16(eo+10,true), cdo=dv.getUint32(eo+16,true);
 let p=cdo; const files={};
 const dec=new TextDecoder("utf-8");
 for(let i=0;i<cnt;i++){
  if(dv.getUint32(p,true)!==0x02014b50) break;
  const method=dv.getUint16(p+10,true), csz=dv.getUint32(p+20,true);
  const nl=dv.getUint16(p+28,true), el=dv.getUint16(p+30,true), cl=dv.getUint16(p+32,true);
  const lho=dv.getUint32(p+42,true);
  const name=dec.decode(u.subarray(p+46,p+46+nl));
  const lnl=dv.getUint16(lho+26,true), lel=dv.getUint16(lho+28,true);
  const start=lho+30+lnl+lel;
  const raw=u.subarray(start,start+csz);
  files[name]={method:method,raw:raw};
  p+=46+nl+el+cl;
 }
 async function txt(nm){ const f=files[nm]; if(!f) return ""; const b=(f.method===0)?f.raw:await inflateRaw(f.raw); return dec.decode(b) }
 const ss=[]; const sst=await txt("xl/sharedStrings.xml");
 if(sst){ const m=sst.match(/<si[\s\S]*?<\/si>/g)||[]; m.forEach(function(si){
   const ts=si.match(/<t[^>]*>([\s\S]*?)<\/t>/g)||[];
   ss.push(unxml(ts.map(function(t){return t.replace(/<[^>]+>/g,"")}).join("")));
 })}
 let sheetName=null;
 for(const k in files){ if(/^xl\/worksheets\/sheet\d+\.xml$/.test(k)){ if(!sheetName||k<sheetName) sheetName=k } }
 const sx=await txt(sheetName||"xl/worksheets/sheet1.xml");
 const rows=[];
 const rw=sx.match(/<row[\s\S]*?<\/row>|<row[^>]*\/>/g)||[];
 rw.forEach(function(r){
  const cells=r.match(/<c[\s\S]*?<\/c>|<c[^>]*\/>/g)||[];
  const out=[];
  cells.forEach(function(c){
   const ref=(c.match(/r="([A-Z]+)\d+"/)||[])[1]||"";
   let ci=0; for(let i=0;i<ref.length;i++) ci=ci*26+(ref.charCodeAt(i)-64);
   ci=ci?ci-1:out.length;
   const ty=(c.match(/t="([^"]+)"/)||[])[1]||"n";
   let v="";
   if(ty==="inlineStr"){ const t=c.match(/<t[^>]*>([\s\S]*?)<\/t>/); v=t?t[1]:"" }
   else { const m=c.match(/<v>([\s\S]*?)<\/v>/); v=m?m[1]:""; if(ty==="s") v=ss[+v]||"" }
   v=unxml(v);
   while(out.length<ci) out.push("");
   out[ci]=v;
  });
  if(out.some(function(x){return String(x).trim()!==""})) rows.push(out);
 });
 return rows;
}
async function readAny(f){
 const nm=f.name.toLowerCase();
 if(nm.endsWith(".xlsx")) return await xlsxRead(await f.arrayBuffer());
 return csvParse(await f.text());
}

/* ================= ESCRITA XLSX (nativa) ================= */
const CRCT=(function(){const t=[];for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=(c&1)?(0xEDB88320^(c>>>1)):(c>>>1);t[n]=c>>>0}return t})();
function crc32(u){let c=0^-1;for(let i=0;i<u.length;i++)c=(c>>>8)^CRCT[(c^u[i])&0xFF];return (c^-1)>>>0}
function xlsxWrite(sheetRows, sheetTitle){
 const enc=new TextEncoder();
 function esx(s){return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}
 let sd="";
 sheetRows.forEach(function(r,ri){
  sd+="<row r=\""+(ri+1)+"\">";
  r.forEach(function(v,ci){
   let ref=""; let n=ci+1; while(n>0){const m=(n-1)%26;ref=String.fromCharCode(65+m)+ref;n=Math.floor((n-1)/26)}
   sd+="<c r=\""+ref+(ri+1)+"\" t=\"inlineStr\"><is><t xml:space=\"preserve\">"+esx(v)+"</t></is></c>";
  });
  sd+="</row>";
 });
 const files=[
  ["[Content_Types].xml","<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>"],
  ["_rels/.rels","<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>"],
  ["xl/workbook.xml","<?xml version=\"1.0\" encoding=\"UTF-8\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\""+esx(sheetTitle||"Resultado").substr(0,31)+"\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>"],
  ["xl/_rels/workbook.xml.rels","<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>"],
  ["xl/worksheets/sheet1.xml","<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"+sd+"</sheetData></worksheet>"]
 ];
 const chunks=[],cd=[]; let off=0;
 files.forEach(function(f){
  const nb=enc.encode(f[0]), db2=enc.encode(f[1]);
  const cr=crc32(db2);
  const lh=new Uint8Array(30+nb.length), lv=new DataView(lh.buffer);
  lv.setUint32(0,0x04034b50,true); lv.setUint16(4,20,true); lv.setUint16(6,0,true); lv.setUint16(8,0,true);
  lv.setUint16(10,0,true); lv.setUint16(12,0,true); lv.setUint32(14,cr,true);
  lv.setUint32(18,db2.length,true); lv.setUint32(22,db2.length,true);
  lv.setUint16(26,nb.length,true); lv.setUint16(28,0,true);
  lh.set(nb,30);
  chunks.push(lh,db2);
  const ch=new Uint8Array(46+nb.length), cv=new DataView(ch.buffer);
  cv.setUint32(0,0x02014b50,true); cv.setUint16(4,20,true); cv.setUint16(6,20,true);
  cv.setUint16(8,0,true); cv.setUint16(10,0,true); cv.setUint16(12,0,true); cv.setUint16(14,0,true);
  cv.setUint32(16,cr,true); cv.setUint32(20,db2.length,true); cv.setUint32(24,db2.length,true);
  cv.setUint16(28,nb.length,true); cv.setUint16(30,0,true); cv.setUint16(32,0,true);
  cv.setUint16(34,0,true); cv.setUint16(36,0,true); cv.setUint32(38,0,true); cv.setUint32(42,off,true);
  ch.set(nb,46); cd.push(ch);
  off += lh.length + db2.length;
 });
 let cdl=0; cd.forEach(function(c){cdl+=c.length});
 const eo=new Uint8Array(22), ev=new DataView(eo.buffer);
 ev.setUint32(0,0x06054b50,true); ev.setUint16(8,cd.length,true); ev.setUint16(10,cd.length,true);
 ev.setUint32(12,cdl,true); ev.setUint32(16,off,true);
 return new Blob(chunks.concat(cd,[eo]),{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
}
function download(blob,name){
 const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=name;
 document.body.appendChild(a); a.click(); setTimeout(function(){URL.revokeObjectURL(a.href);a.remove()},1500);
}
function dlCsv(rows,name){
 const t=rows.map(function(r){return r.map(function(c){const s=String(c==null?"":c);return /[";\n]/.test(s)?"\""+s.replace(/"/g,"\"\"")+"\"":s}).join(";")}).join("\r\n");
 download(new Blob(["\uFEFF"+t],{type:"text/csv;charset=utf-8"}),name);
}

/* ================= MAPEAMENTO DE COLUNAS ================= */
function mapCols(head){
 const h=head.map(function(x){return norm(x)});
 const find=function(fn){for(let i=0;i<h.length;i++) if(fn(h[i])) return i; return -1};
 const iCurta=find(function(x){return /CURTA|SANEAD/.test(x)});
 const iLonga=find(function(x){return /LONGA|PADRONIZAD/.test(x)});
 let iDesc=find(function(x){return /DESCR/.test(x) && !/CURTA|LONGA|SANEAD|PADRONIZAD/.test(x)});
 if(iDesc<0) iDesc=find(function(x){return /MATERIAL|ITEM|SOLICITAD|TEXTO/.test(x) && !/TIPO|CODIGO/.test(x)});
 if(iDesc<0) iDesc = iCurta>=0?iCurta:0;
 return {
  d:iDesc, c:iCurta, l:iLonga,
  tm:find(function(x){return /(TIPO|^MTART$|^TM$)/.test(x) && !/AVALIA/.test(x)}),
  gm:find(function(x){return /(MERCADORIA|^MATKL$|^GRP MERC|^GRUPO MERC|^GM$|^G.M.$)/.test(x)}),
  gc:find(function(x){return /(COMPRADOR|^EKGRP$|^GRP COMPR|^GRUPO COMPR|^GC$|^G.C.$)/.test(x)}),
  gci:find(function(x){return /CATEGORIA/.test(x) && !/AVALIA/.test(x)}),
  sa:find(function(x){return /(SETOR|^SPART$|^SA$)/.test(x)}),
  ncm:find(function(x){return /NCM|NBM|CLASSIFICACAO FISCAL|CLASS FISCAL|COD FISCAL/.test(x)}),
  ca:find(function(x){return /(CLASSE.*AVALIA|AVALIA.*CLASSE|^BKLAS|^CL\.?\s*AVAL|^CTG\.?\s*AVAL|^CLASSE\s*AVAL|CLASSE.*CONTABIL|VALUATION.*CLASS|^C\.?\s*AVALIA|CATEGORIA.*AVALIA)/.test(x) && !/GRUPO|MERCADORIA|COMPRADOR/.test(x)}),
  cod:find(function(x){return /^(CODIGO|COD|MATERIAL|SAP|NUMERO)/.test(x) && !/DESCR/.test(x)})
 };
}
let CA_COLSTATE=null;
function looksHeader(r){ return r.some(function(c){return /DESCR|TIPO|MERCADORIA|COMPRADOR|MATERIAL|CLASSE|AVALIA|^CODIGO|GRUPO|SETOR|NCM|BKLAS/.test(norm(c))}) }

/* ================= RENDER ================= */
function cCls(v){ return v>=0.80?"ok":(v>=0.60?"wa":"ri") }
function cHex(v){ return v>=0.80?"#1a7f4b":(v>=0.60?"#9a6b00":"#b3261e") }
function pctTag(p){ return "<span class=\"tag "+cCls(p)+"\">"+Math.round(p*100)+"%</span>" }
function tmDesc(c){ const t=TIPOS.filter(function(x){return x[0]===c})[0]; return t?t[1]:"" }
function gmDesc(c){ return GM[c]||"" }
function gcDesc(c){ return GC[c]||"" }
function catNote(f,c){
 if(f==="gm"){ if(!c) return "nenhum codigo do catalogo 663 com confianca suficiente - revisar manualmente";
  if(descont(f,c)) return "codigo descontinuado, fora do catalogo 663 - nao pode ser usado"; if(!GM_OFF[c]) return "codigo fora do catalogo 663"; return "" }
 if(f==="gc"){ if(!c) return "nenhum codigo do catalogo 627 com confianca suficiente - revisar manualmente";
  if(descont(f,c)) return "codigo descontinuado, fora do catalogo 627 - nao pode ser usado"; if(!GC_OFF[c]) return "codigo fora do catalogo 627"; return "" }
 if(f==="tm" && !c) return "nenhum codigo do catalogo 650 com confianca suficiente - revisar manualmente";
 if(f==="ca"){ if(!caCount()) return "nenhuma lista de classe de avaliacao carregada ainda - traga a coluna CLASSE DE AVALIACAO na base ou cadastre as classes no painel do administrador";
  if(!c) return "nenhuma classe de avaliacao sustentada pela base para este material - revisar manualmente";
  if(!CA_OFF[c]) return "classe fora da lista carregada - bloqueada para sugestao e exportacao"; return "" }
 return "";
}
/* ===== TRANSPARENCIA v7: qual lista foi consultada e se houve aproximacao ===== */
function listaDe(f){
 const nome=(typeof CAT_SRC!=="undefined" && CAT_SRC[f])? CAT_SRC[f] : "";
 const src=(typeof CATSRC!=="undefined" && CATSRC[f])? CATSRC[f] : ((typeof SHEET_SRC!=="undefined" && SHEET_SRC[f])||"embutido");
 return nome+" - lista consultada: "+src;
}
function matchKind(r,f){
 if(f!=="tm" && f!=="gm" && f!=="gc") return {k:"fixo", txt:"valor fixo do fluxo 302"};
 const c=fldVal(r,f);
 if(!c) return {k:"revisar", txt:"nenhum codigo da lista sustentado pela evidencia - sugerimos REVISAR"};
 if(!codeOk(f,c)) return {k:"revisar", txt:"codigo fora da lista valida - bloqueado para sugestao e exportacao"};
 const v=((r&&r.via)?r.via[f]:"")||"";
 const apw=(f==="gm"||f==="gc")? apxWhy(f, (r&&r.disc&&r.disc[f]&&r.disc[f][0])||"", r?r.gm:"") : "";
 if(/aproxima|sucessor|familia|predominancia|predominante|derivado|coerente com|geral do 302/i.test(v) || apw)
  return {k:"aprox", txt:"aproximacao (mais proximo por familia/categoria): "+(v||apw)};
 return {k:"exato", txt:"codigo exato encontrado na lista consultada"};
}
function fldVal(r,f){ return f==="tm"?r.tm:(f==="gm"?r.gm:(f==="gc"?r.gc:(f==="ca"?(r.ca||""):(f==="gci"?r.gci:(f==="setor"?r.setor:r.aprovado))))) }
function fldDesc(f,c){ return f==="tm"?tmDesc(c):(f==="gm"?gmDesc(c):(f==="gc"?gcDesc(c):(f==="ca"?caDesc(c):"padrao fixo do fluxo 302"))) }
function optList(f,cur){
 let ks=[];
 if(f==="tm") ks=TIPOS.map(function(t){return [t[0],t[1]]});
 if(f==="gm") ks=Object.keys(GM_OFF).filter(function(k){return GM_IMOB.indexOf(k)<0}).map(function(k){return [k,GM_OFF[k]]});
 if(f==="gc") ks=Object.keys(GC_OFF).map(function(k){return [k,GC_OFF[k]]});
 if(f==="ca") ks=Object.keys(CA_OFF).sort().map(function(k){return [k,CA_OFF[k]]});
 if(f==="gci") ks=GCI_OPC.map(function(k){return [k,k==="NORM"?"padrao do fluxo 302":"item nao estocavel"]});
 if(f==="setor") ks=SETOR_OPC.map(function(k){return [k,k==="80"?"padrao do fluxo 302":"outro setor visto no historico"]});
 if(f==="aprovado") ks=[[FIXOS.aprovado,"padrao do fluxo 302"],["PENDENTE","aguardando aprovacao"]];
 /* nenhum valor fora do catalogo oficial entra na lista, nem como "valor atual" */
 if(cur && codeOk(f,cur) && !ks.some(function(x){return x[0]===cur})) ks.unshift([cur,"valor atual"]);
 if((f==="gm"||f==="gc"||f==="tm"||f==="ca") && !cur) ks.unshift(["","em branco - revisar"]);
 return ks.map(function(x){return "<option value=\""+esc(x[0])+"\""+(x[0]===cur?" selected":"")+">"+esc(x[0]+(x[1]?" - "+x[1]:""))+"</option>"}).join("");
}
function fldHtml(i,f,label,code,desc,conf,fixed){
 const note=catNote(f,code);
 let h="<div class=\"fld"+(fixed?" fix":"")+"\" data-i=\""+i+"\" data-f=\""+f+"\">";
 h+="<span class=\"fl\">"+esc(label)+"</span>";
 h+= code ? "<span class=\"fc\">"+esc(code)+"</span>" : "<span class=\"fc rev\">REVISAR</span>";
 h+="<span class=\"fd\">"+esc(desc||note||"sem descricao no catalogo")+"</span>";
 if(note && desc) h+="<span class=\"fd\" style=\"color:#9a6b00\">"+esc(note)+"</span>";
 h+="<span class=\"fx\">";
 h+= (conf===null||conf===undefined) ? "<span class=\"cf ok\">padrao pre-preenchido</span>"
                                    : "<span class=\"cf "+cCls(conf)+"\">confianca "+Math.round(conf*100)+"%</span>";
 h+="<button class=\"btn gho fedit\">editar</button></span>";
 h+="</div>";
 return h;
}
function dupsHtml(r){
 if(!r.dups.length) return "";
 let h="<div style=\"margin-top:12px\"><b style=\"font-size:13px\">Materiais parecidos ja cadastrados</b><div class=\"tw\"><table><thead><tr><th>%</th><th>Risco</th><th>Codigo SAP</th><th>Material</th><th>Classificacao atual</th></tr></thead><tbody>";
 r.dups.forEach(function(d){
  h+="<tr><td class=\"mono\"><b>"+d.pct+"%</b></td><td><span class=\"tag "+(d.pct>=90?"ri":(d.pct>=75?"wa":""))+"\">"+d.risk+"</span></td><td class=\"mono\">"+esc(d.cod||"-")+"</td><td>"+esc(d.d)+"</td><td class=\"mono\">"+esc([d.tm,d.gm,d.gc].filter(Boolean).join(" / "))+"</td></tr>";
 });
 return h+"</tbody></table></div></div>";
}
function fldLabel(f){ return {tm:"Tipo de Material",gm:"Grupo de Mercadoria",gc:"Grupo de Compradores",ca:"Classe de Avaliacao",gci:"Grupo Categoria Item",setor:"Setor de Atividade",aprovado:"Cadastro aprovado?"}[f] }
function origem(r,f){
 if(f==="gci"||f==="setor"||f==="aprovado") return "valor fixo do fluxo 302, pre-preenchido sem sugestao";
 const hd=r.head||{tk:"",n:0}, parts=[];
 if(hd.n){
  const p = f==="tm"?hd.ptm:(f==="gm"?hd.pgm:hd.pgc);
  const k = f==="tm"?hd.tm:(f==="gm"?hd.gm:hd.gc);
  parts.push("substantivo \""+(hd.via||hd.tk)+"\": "+hd.n+" registro(s) do historico 302, "+Math.round((p||0)*100)+"% com "+k+
   (hd.src==="base"?" (planilha carregada)":(hd.src==="curado"?" (lista curada)":" (modelo embarcado da base 302)")));
 } else parts.push("sem historico para o substantivo \""+(hd.tk||"?")+"\"");
 if(r.ref) parts.push("registro de referencia: "+(r.ref.cod?r.ref.cod+" - ":"")+r.ref.d+" ("+r.ref.sim+"% de semelhanca, cadastrado como "+[r.ref.tm,r.ref.gm,r.ref.gc].filter(Boolean).join(" / ")+")");
 if(r.regra) parts.push(r.regra);
 if(r.via && r.via[f]) parts.push("caminho ate o codigo valido: "+r.via[f]);
 if(r.disc && r.disc[f] && r.disc[f].length) parts.push("descartado por estar fora do catalogo oficial: "+r.disc[f].map(function(k){
   const d=(f==="gm"?GM_LEG:GC_LEG)[k]; return k+(d?" ("+d+")":"")+" - descontinuado" }).join(", "));
 if((f==="gm"||f==="gc") && !(f==="gm"?r.gm:r.gc)) parts.push("nenhum codigo do catalogo "+(f==="gm"?"663":"627")+" atingiu confianca minima: campo em branco para revisao manual");
 return parts.join("  |  ");
}
function caTranspRow(r){
 const ci=r.caInfo||{lista:"-",kind:"revisar",via:"-"};
 const c=r.ca||"", cf=(r.cCa||0);
 const kind = ci.kind==="exata"? "correspondencia exata na base" : (ci.kind==="aproximada"? "aproximacao por evidencia da base" : "sem evidencia - REVISAR");
 return "<tr><td>"+esc(fldLabel("ca"))+"</td><td><span class=\"cd\">"+esc(c||"REVISAR")+"</span></td><td>"+esc(caDesc(c)||catNote("ca",c)||"-")+"</td>"+
  "<td><span class=\"cx\">"+esc((CAT_SRC.ca||"classe de avaliacao")+" - lista consultada: "+(ci.lista||"-"))+"</span></td>"+
  "<td><span class=\"cx\">"+esc(kind)+"</span></td>"+
  "<td>"+(c? "<b style=\"color:"+cHex(cf)+"\">"+Math.round(cf*100)+"%</b>" : "<span class=\"tag wa\">revisar</span>")+"</td>"+
  "<td><span class=\"cx\">"+esc(ci.via||"-")+"</span></td></tr>";
}
function transpHtml(r){
 const fs=[["tm",r.tm,r.cTm],["gm",r.gm,r.cGm],["gc",r.gc,r.cGc],["gci",r.gci,null],["setor",r.setor,null],["aprovado",r.aprovado,null]];
 let h="<details style=\"margin-top:12px\"><summary>painel de transparencia - de onde veio cada codigo (nao vai para a exportacao)</summary><div>";
 h+="<div class=\"tw\"><table><thead><tr><th>Campo</th><th>Codigo exportado</th><th>Descricao (so leitura)</th><th>Lista consultada</th><th>Correspondencia</th><th>Confianca</th><th>Como o motor chegou nesse codigo</th></tr></thead><tbody>";
 fs.forEach(function(x){
  h+="<tr><td>"+esc(fldLabel(x[0]))+"</td><td><span class=\"cd\">"+esc(String(x[1]))+"</span></td><td>"+esc(fldDesc(x[0],x[1])||catNote(x[0],x[1])||"-")+"</td>"+
     "<td><span class=\"cx\">"+esc(listaDe(x[0]))+"</span></td>"+
     "<td><span class=\"cx\">"+esc(matchKind(r,x[0]).txt)+"</span></td>"+
     "<td>"+(x[2]===null?"<span class=\"tag ok\">fixo</span>":"<b style=\"color:"+cHex(x[2])+"\">"+Math.round(x[2]*100)+"%</b>")+"</td>"+
     "<td><span class=\"cx\">"+esc(origem(r,x[0]))+"</span></td></tr>";
 });
 h+=caTranspRow(r);
 h+=imoTranspRow(r);
 h+="</tbody></table></div>"+eqapTranspHtml(r)+"<div class=\"hint\" style=\"margin-top:8px\">A planilha exportada leva apenas a coluna Codigo exportado. Descricao, catalogo, confianca e origem ficam somente nesta tela.</div></div></details>";
 return h;
}
function transpAllHtml(){
 const lim=300, rs=S.results.slice(0,lim);
 let h="<details style=\"margin-top:14px\"><summary>painel de transparencia do lote - origem do codigo por material (nao vai para a exportacao)</summary><div>";
 h+="<div class=\"tw\"><table><thead><tr><th>#</th><th>Material</th><th>Codigos exportados</th><th>Descricoes (so leitura)</th><th>Confianca por campo</th><th>Lista e correspondencia</th><th>Origem da decisao</th></tr></thead><tbody>";
 rs.forEach(function(r,i){
  h+="<tr><td class=\"mono\">"+(i+1)+"</td><td>"+esc(r.desc)+"</td>"+
     "<td><span class=\"cd\">"+esc(r.tm+" / "+r.gm+" / "+r.gc+" / "+r.gci+" / "+r.setor)+"</span></td>"+
     "<td><span class=\"cx\">"+esc([tmDesc(r.tm),gmDesc(r.gm)||catNote("gm",r.gm),gcDesc(r.gc)||catNote("gc",r.gc)].filter(Boolean).join(" | "))+"</span></td>"+
     "<td><span class=\"cx\">tipo "+Math.round(r.cTm*100)+"% | mercadoria "+Math.round(r.cGm*100)+"% | compradores "+Math.round(r.cGc*100)+"%</span></td>"+
     "<td><span class=\"cx\">"+esc(listaDe("gm")+" | "+matchKind(r,"gm").txt+"  ||  "+listaDe("gc")+" | "+matchKind(r,"gc").txt)+"</span></td>"+
     "<td><span class=\"cx\">"+esc(origem(r,"gm"))+"</span></td></tr>";
 });
 h+="</tbody></table></div>"+eqapLoteHtml();
 if(S.results.length>lim) h+="<div class=\"hint\" style=\"margin-top:8px\">mostrando os "+lim+" primeiros de "+S.results.length+" materiais; a analise completa em XLSX traz todos.</div>";
 h+="</div></details>";
 return h;
}
function cardHtml(r,i){
 const pc=Math.round(r.conf*100);
 let h="<div class=\"card\" style=\"margin-bottom:14px\">";
 h+="<div style=\"margin-bottom:12px\"><span class=\"fl\" style=\"font-size:10.5px;color:#5b6675;text-transform:uppercase;letter-spacing:.6px;font-weight:700\">Material analisado</span><div style=\"font-size:16px;font-weight:700\">"+esc(r.desc)+"</div></div>";
 h+="<div class=\"gconf\"><span class=\"gv "+cCls(r.conf)+"\">"+pc+"%</span><div class=\"gb\"><b>confianca global da classificacao</b>"+
    "<div class=\"bar\" style=\"margin:6px 0\"><i style=\"width:"+pc+"%;background:"+cHex(r.conf)+"\"></i></div>"+
    "<span class=\"hint\">"+(r.conf>=0.80?"pode seguir para o cadastro":(r.conf>=0.60?"confira os campos em amarelo antes de integrar":"revisao obrigatoria antes de integrar"))+"</span></div></div>";
 h+="<div class=\"fldgrid\" style=\"margin-top:14px\">";
 h+=fldHtml(i,"tm","Tipo de Material",r.tm,tmDesc(r.tm),r.cTm,false);
 h+=fldHtml(i,"gm","Grupo de Mercadoria",r.gm,gmDesc(r.gm),r.cGm,false);
 h+=fldHtml(i,"gc","Grupo de Compradores",r.gc,gcDesc(r.gc),r.cGc,false);
 h+=fldHtml(i,"ca","Classe de Avaliacao",r.ca||"",caDesc(r.ca||""),(r.cCa||0),false);
 h+=fldHtml(i,"gci","Grupo Categoria Item",r.gci,"padrao fixo do fluxo 302",null,true);
 h+=fldHtml(i,"setor","Setor de Atividade",r.setor,"padrao fixo do fluxo 302",null,true);
 h+=fldHtml(i,"aprovado","Cadastro aprovado?",r.aprovado,"padrao fixo do fluxo 302",null,true);
 h+="</div>";
 h+=dupCurtaBanner(r);
 h+=imoBanner(r);
 if(r.imob.flag && !r.imoc) h+="<div class=\"note "+(r.imob.lvl==="alto"?"ri":"wa")+"\" style=\"margin-top:12px\"><b>Possivel Imobilizado - avaliar Fluxo 304/ZATI.</b> "+esc(r.imob.why)+"</div>";
 if(r.div) h+="<div class=\"note wa\" style=\"margin-top:10px\">"+esc(r.div)+"</div>";
 h+=simHtml(r);
 h+=dupsHtml(r);
 h+=transpHtml(r);
 h+="<details style=\"margin-top:12px\"><summary>por que essa classificacao e essa confianca</summary><div><ul style=\"margin:0 0 0 18px;padding:0;font-size:12.5px;color:#46525f\">"+
    (r.why||[r.just]).map(function(x){return "<li style=\"margin-bottom:4px\">"+esc(x)+"</li>"}).join("")+"</ul></div></details>";
 h+="</div>";
 return h;
}
function kpiHtml(rows){
 const n=rows.length, ex=rows.filter(function(r){return r.flag}).length;
 const im=rows.filter(function(r){return r.imob.flag}).length;
 const dp=rows.filter(function(r){return r.dups[0]&&r.dups[0].pct>=75}).length;
 const lc=rows.filter(function(r){return r.conf<0.60}).length;
 const avg=n?Math.round(rows.reduce(function(a,r){return a+r.conf},0)/n*100):0;
 return "<div class=\"grid g3\" style=\"margin-bottom:14px\">"+
  "<div class=\"kpi\"><b>"+n+"</b><span>materiais analisados</span></div>"+
  "<div class=\"kpi\"><b style=\"color:"+cHex(avg/100)+"\">"+avg+"%</b><span>confianca media</span></div>"+
  "<div class=\"kpi\"><b style=\"color:#9a6b00\">"+im+"</b><span>alerta de imobilizado</span></div>"+
  "<div class=\"kpi\"><b style=\"color:#b3261e\">"+dp+"</b><span>duplicidade provavel/possivel</span></div>"+
  "<div class=\"kpi\"><b>"+lc+"</b><span>confianca abaixo de 60%</span></div>"+
  "<div class=\"kpi\"><b>"+(n-ex)+"</b><span>prontos sem revisao</span></div></div>";
}
function cdCell(f,code,conf){
 const d=fldDesc(f,code), note=catNote(f,code);
 let h= code ? "<span class=\"cd\">"+esc(code)+"</span>" : "<span class=\"cd rev\">REVISAR</span>";
 h+="<span class=\"cx\">"+esc(d||note||"sem descricao no catalogo")+"</span>";
 if(conf!==null&&conf!==undefined) h+="<span class=\"cx\" style=\"color:"+cHex(conf)+";font-weight:700\">"+Math.round(conf*100)+"%</span>";
 return h;
}
function tableHtml(rows,onlyFlag){
 let h="<div class=\"tw\"><table><thead><tr><th>#</th><th>Material</th><th>Tipo de Material</th><th>Grupo de Mercadoria</th><th>Grupo de Compradores</th><th>Cat. Item</th><th>Setor</th><th>Confianca</th><th>Imobilizado</th><th>Duplicidade</th><th></th></tr></thead><tbody>";
 let shown=0;
 rows.forEach(function(r,i){
  if(onlyFlag && !r.flag) return;
  shown++;
  const d=r.dups[0];
  h+="<tr>"+
   "<td class=\"mono\">"+(i+1)+"</td>"+
   "<td>"+esc(r.desc)+(r.div?"<span class=\"cx\" style=\"color:#9a6b00\">"+esc(r.div)+"</span>":"")+"</td>"+
   "<td>"+cdCell("tm",r.tm,r.cTm)+"</td>"+
   "<td>"+cdCell("gm",r.gm,r.cGm)+"</td>"+
   "<td>"+cdCell("gc",r.gc,r.cGc)+"</td>"+
   "<td><span class=\"cd\">"+esc(r.gci)+"</span></td>"+
   "<td><span class=\"cd\">"+esc(r.setor)+"</span></td>"+
   "<td><b style=\"color:"+cHex(r.conf)+"\">"+Math.round(r.conf*100)+"%</b></td>"+
   "<td>"+(r.imoc? imoTag(r.imoc) : (r.imob.flag?"<span class=\"tag "+(r.imob.lvl==="alto"?"ri":"wa")+"\">"+r.imob.lvl+"</span>":"<span class=\"tag ok\">nao</span>"))+"</td>"+
   "<td>"+(d?"<span class=\"tag "+(d.pct>=90?"ri":(d.pct>=75?"wa":""))+"\">"+d.pct+"% "+d.risk+"</span><span class=\"cx\">"+esc((d.cod?d.cod+" - ":"")+d.d)+"</span>":"<span class=\"tag ok\">nenhuma</span>")+"</td>"+
   "<td><button class=\"btn gho rowopen\" data-i=\""+i+"\">abrir</button></td>"+
   "</tr>";
 });
 if(!shown) h+="<tr><td colspan=\"11\" class=\"hint\">nenhum item nesse filtro</td></tr>";
 return h+"</tbody></table></div>";
}
function renderOut(){
 const box=$("#out");
 if(!S.results.length){ box.innerHTML="<div class=\"empty\">Nenhum material analisado ainda. Comece pelo passo 1.</div>"; return }
 const oe=$("#onlyEx").checked;
 if(S.results.length===1) box.innerHTML=cardHtml(S.results[0],0);
 else box.innerHTML=kpiHtml(S.results)+"<div id=\"detail\"></div>"+tableHtml(S.results,oe)+transpAllHtml();
 bindOut(); buildTsv();
}
function bindOut(){
 $$("#out .fedit").forEach(function(b){ b.onclick=function(){ openEdit(b.closest(".fld")) } });
 $$("#out .rowopen").forEach(function(b){ b.onclick=function(){
  const i=+b.dataset.i, dv=$("#detail");
  dv.innerHTML=cardHtml(S.results[i],i); bindOut();
  dv.scrollIntoView({behavior:"smooth",block:"nearest"});
 }});
}
function openEdit(el){
 const i=+el.dataset.i, f=el.dataset.f, r=S.results[i], cur=fldVal(r,f);
 el.innerHTML="<span class=\"fl\">"+esc(el.querySelector(".fl")?el.querySelector(".fl").textContent:"")+"</span>";
 const lbl={tm:"Tipo de Material",gm:"Grupo de Mercadoria",gc:"Grupo de Compradores",ca:"Classe de Avaliacao",gci:"Grupo Categoria Item",setor:"Setor de Atividade",aprovado:"Cadastro aprovado?"}[f];
 el.innerHTML="<span class=\"fl\">"+esc(lbl)+"</span><select class=\"esel\">"+optList(f,cur)+"</select>"+
  "<span class=\"fx\"><button class=\"btn esave\">salvar</button><button class=\"btn gho ecanc\">cancelar</button></span>";
 el.querySelector(".esave").onclick=function(){ applyEdit(i,f,el.querySelector(".esel").value) };
 el.querySelector(".ecanc").onclick=function(){ renderOut() };
}
function applyEdit(i,f,v){
 const r=S.results[i];
 v=codeOf(f,v);
 if(f==="tm"){ r.tm=v; r.cTm=0.99; r.gci = codeOf("gci", v==="NLAG"?FIXOS.gciNlag:FIXOS.gci) }
 else if(f==="gm"){ r.gm=v; r.cGm=0.99 }
 else if(f==="gc"){ r.gc=v; r.cGc=0.99 }
 else if(f==="gci") r.gci=v;
 else if(f==="setor") r.setor=v;
 else if(f==="aprovado") r.aprovado=v;
 r.conf = 0.30*r.cTm + 0.40*r.cGm + 0.30*r.cGc;
 r.flag = r.imob.flag || r.conf<0.60 || (r.dups[0]&&r.dups[0].pct>=75) || !!r.div;
 if(f==="tm"||f==="gm"||f==="gc"){
  S.fb.push({d:r.desc,tm:r.tm,gm:r.gm,gc:r.gc,t:Date.now()}); saveFb();
  S.log.push({t:Date.now(),d:r.desc,tm:r.tm,gm:r.gm,gc:r.gc,conf:Math.round(r.conf*100),src:"usuario",ev:"correcao aprendida ("+f+" = "+v+")"});
  toast("correcao registrada e aprendida");
 } else {
  S.log.push({t:Date.now(),d:r.desc,tm:r.tm,gm:r.gm,gc:r.gc,conf:Math.round(r.conf*100),src:"usuario",ev:"ajuste manual ("+f+" = "+v+")"});
  toast("campo ajustado");
 }
 saveLog(); renderLog(); renderOut();
}

/* ================= FLUXO DE ANALISE ================= */
/* celula exportada: SO codigo do catalogo oficial. Codigo descontinuado/invalido sai em branco. */
function expCell(f,v){ const c=codeSafe(f,v); if(f==="ca") return c? c : "REVISAR"; return (f==="setor"||f==="gm")? (/^\d+$/.test(c)? Number(c) : c) : c }
function rows4mdg(skipEx){
 const out=[COLS4MDG.slice()];
 S.results.forEach(function(r){
  if(skipEx && r.flag) return;
  /* somente codigo em cada coluna de classificacao - nenhuma descricao, nenhuma concatenacao */
  out.push([ String(r.desc||"").trim(),
             expCell("tm",r.tm),
             expCell("gm",r.gm),
             expCell("gc",r.gc),
             expCell("gci",r.gci),
             expCell("setor",r.setor),
             expCell("aprovado",r.aprovado),
             expCell("ca",r.ca||"") ]);
 });
 return out;
}
/* auditoria da exportacao: garante que nenhuma celula de classificacao levou descricao */
function auditExport(skipEx){
 const rs=rows4mdg(skipEx), fs=["tm","gm","gc","gci","setor","aprovado","ca"], bad=[], rev=[];
 for(let i=1;i<rs.length;i++){
  for(let k=0;k<fs.length;k++){
   const raw=rs[i][k+1], v=String(raw===undefined||raw===null?"":raw);
   if(v===""||v==="REVISAR"){ rev.push({linha:i, campo:COLS4MDG[k+1]}); continue }
   const c=codeOf(fs[k],v);
   if(!codeOk(fs[k],c) || /[a-z]/.test(v) || v.indexOf(" ")>=0 || v.indexOf("-")>=0)
    bad.push({linha:i, campo:COLS4MDG[k+1], valor:v, motivo: descont(fs[k],c)?"codigo descontinuado, fora do catalogo oficial":"valor nao e codigo do catalogo"});
  }
 }
 return {total:rs.length-1, bad:bad, rev:rev};
}
function rowsFull(){
 const out=[["Descricao analisada","Descricao curta","Tipo de material","Descricao do tipo","Grupo de Mercadorias","Descricao do grupo de mercadorias","Grupo de compradores","Descricao do grupo de compradores","Grupo categoria Item","Setor de atividade","Cadastro aprovado?","Classe de Avaliacao","Descricao da classe de avaliacao","Confianca classe de avaliacao %","Confianca tipo %","Confianca grupo mercadoria %","Confianca grupo compradores %","Confianca global %","Alerta imobilizado","Motivo imobilizado","Duplicidade %","Risco","Material candidato","Divergencia descricoes","Justificativa","Conferencia de imobilizado","Nivel de evidencia imobilizado","Regra acionada (IN-CPR-01.001 / RES Risco 194)","Conflito nas duas bases"]];
 S.results.forEach(function(r){
  const d=r.dups[0];
  out.push([r.desc,r.curta,codeOf("tm",r.tm),tmDesc(r.tm),codeOf("gm",r.gm),gmDesc(r.gm),codeOf("gc",r.gc),gcDesc(r.gc),codeOf("gci",r.gci),codeOf("setor",r.setor),codeOf("aprovado",r.aprovado),
            codeSafe("ca",r.ca||""),caDesc(r.ca||""),Math.round((r.cCa||0)*100),
            Math.round(r.cTm*100),Math.round(r.cGm*100),Math.round(r.cGc*100),Math.round(r.conf*100),
            r.imob.flag?"SIM":"NAO", r.imob.why, d?d.pct:"", d?d.risk:"", d?((d.cod?d.cod+" - ":"")+d.d):"", r.div, r.just,
            (r.imoc? imoVerdTxt(r.imoc)+" ("+Math.round(r.imoc.conf*100)+"%)" : ""),
            (r.imoc? (r.imoc.nivel||"nenhum nivel de evidencia casou") : ""),
            (r.imoc? r.imoc.regra : ""),
            (r.imoc && r.imoc.conflito? "SIM":"NAO")]);
 });
 return out;
}
function buildTsv(){
 const sk=$("#skipEx")?$("#skipEx").checked:false;
 $("#tsv").value = rows4mdg(sk).map(function(r){return r.join("\t")}).join("\n");
}
function needResults(){ if(!S.results.length){toast("analise algum material primeiro");return false} return true }
function copyTsv(withHead){
 if(!needResults()) return;
 const rs=rows4mdg($("#skipEx").checked); const data=withHead?rs:rs.slice(1);
 const txt=data.map(function(r){return r.join("\t")}).join("\n");
 (navigator.clipboard?navigator.clipboard.writeText(txt):Promise.reject()).then(function(){toast((data.length-(withHead?1:0))+" linhas copiadas no layout do cadastro em massa")},
  function(){const ta=$("#tsv");ta.value=txt;ta.select();document.execCommand("copy");toast("copiado")});
}
function logResults(ev){
 S.results.forEach(function(r){ S.log.push({t:Date.now(),d:r.desc,tm:r.tm,gm:r.gm,gc:r.gc,conf:Math.round(r.conf*100),src:"IA",ev:ev}) });
 saveLog(); renderLog();
}
function analyzeList(items){
 const box=$("#out");
 box.innerHTML="<div class=\"sk\" style=\"width:40%\"></div><div class=\"sk\" style=\"width:100%;margin-top:10px\"></div><div class=\"sk\" style=\"width:92%;margin-top:8px\"></div>";
 const out=[]; let i=0;
 return new Promise(function(res){
  function step(){
   const t0=performance.now();
   while(i<items.length && performance.now()-t0<80){
    const it=items[i++]; if(!it.d && !it.c) continue;
    out.push(classify(it.d,it.c,it.l,it.n||it.ncm||""));
   }
   if(i<items.length){
    box.innerHTML="<div class=\"row\"><b>"+i+" de "+items.length+"</b> analisados</div><div class=\"bar\"><i style=\"width:"+Math.round(i/items.length*100)+"%\"></i></div>";
    setTimeout(step,0);
   } else {
    S.results=out; renderOut(); logResults("analise em lote"); res(out);
   }
  }
  setTimeout(step,10);
 });
}
/* ================= BASE ================= */
function setBaseState(){
 const n=S.base.length;
 $("#baseDot").className = "dot"+(n?"":" off");
 $("#baseTxt").textContent = n? (n.toLocaleString("pt-BR")+" materiais na base") : "base nao carregada";
 const tks=S.lexBase.size;
 $("#baseKpis").innerHTML =
  "<div class=\"kpi\"><b>"+n.toLocaleString("pt-BR")+"</b><span>materiais na base</span></div>"+
  "<div class=\"kpi\"><b>"+tks.toLocaleString("pt-BR")+"</b><span>termos aprendidos</span></div>"+
  "<div class=\"kpi\"><b>"+S.fb.length+"</b><span>correcoes suas</span></div>"+
  "<div class=\"kpi\"><b>"+S.log.length+"</b><span>registros no log</span></div>";
}
async function loadBaseFile(f, append){
 toast("lendo arquivo...");
 const rows=await readAny(f);
 if(!rows.length) return toast("arquivo vazio");
 const head=looksHeader(rows[0])?rows[0]:null;
 const m=mapCols(head||rows[0].map(function(_,i){return "c"+i}));
 const data=head?rows.slice(1):rows;
 const recs=[];
 data.forEach(function(r){
  const d=(r[m.d]||"").toString().trim(); if(!d) return;
  recs.push({d:d, c:(m.c>=0?r[m.c]:"")||"", l:(m.l>=0?r[m.l]:"")||"",
   tm:codeSafe("tm", m.tm>=0?r[m.tm]:""),
   gm:codeSafe("gm", m.gm>=0?r[m.gm]:""),
   gc:codeSafe("gc", m.gc>=0?r[m.gc]:""),
   gci:codeSafe("gci", m.gci>=0?r[m.gci]:""),
   sa:codeSafe("setor", m.sa>=0?r[m.sa]:""),
   cod:((m.cod>=0?r[m.cod]:"")||"").toString().trim(),
   ca:caLearn(m.ca>=0?r[m.ca]:""),
   ncm:ncmClean(m.ncm>=0?r[m.ncm]:"")});
 });
 if(!recs.length) return toast("nenhuma descricao reconhecida no arquivo");
 const sujas = data.filter(function(r){
  const g=(m.gm>=0?String(r[m.gm]||""):""), t=(m.tm>=0?String(r[m.tm]||""):"");
  return (g && g.trim()!==codeOf("gm",g)) || (t && t.trim().toUpperCase()!==codeOf("tm",t));
 }).length;
 S.base = append? S.base.concat(recs) : recs;
 CA_COLSTATE = (m.ca>=0); BASE_ORIGEM="planilha carregada pelo administrador";
 buildIndex(); await saveBase(); setBaseState();
 if(!CA_COLSTATE) setTimeout(function(){ toast("planilha sem coluna CLASSE DE AVALIACAO: as classes saem como REVISAR ate a coluna vir na base ou os codigos serem cadastrados no catalogo do admin"); },2600);
 S.log.push({t:Date.now(),d:f.name,tm:"",gm:"",gc:"",conf:"",src:"Augusto",ev:(append?"base acrescentada: ":"base substituida: ")+recs.length+" registros"});
 saveLog(); renderLog();
 toast(recs.length.toLocaleString("pt-BR")+" materiais carregados e persistidos"+(sujas?" | "+sujas+" linha(s) vinham com codigo+descricao e foram reduzidas ao codigo":""));
}

/* ================= LOG / REGRAS ================= */
function renderLog(){
 const tb=$("#logTb tbody"); const l=S.log.slice(-400).reverse();
 tb.innerHTML=l.map(function(x){
  return "<tr><td class=\"mono\">"+new Date(x.t).toLocaleString("pt-BR")+"</td><td>"+esc(x.d)+"</td><td class=\"code\">"+esc(x.tm||"")+"</td><td class=\"code\">"+esc(x.gm||"")+"</td><td class=\"code\">"+esc(x.gc||"")+"</td><td>"+esc(x.conf===""?"":x.conf+"%")+"</td><td>"+esc(x.src)+"</td><td>"+esc(x.ev)+"</td></tr>";
 }).join("");
}
function renderRegras(){
 $("#tbTipos tbody").innerHTML=TIPOS.map(function(t){return "<tr><td class=\"code\">"+t[0]+"</td><td>"+t[1]+"</td><td>"+t[2]+"</td></tr>"}).join("");
  $("#tbGm tbody").innerHTML=Object.keys(GM_OFF).map(function(k){return "<tr><td class=\"code\">"+k+"</td><td>"+esc(GM_OFF[k])+(GM_IMOB.indexOf(k)>=0?" <span class=\"tag wa\">imobilizado / 304</span>":"")+"</td></tr>"}).join("")+Object.keys(GM_LEG).map(function(k){return "<tr class=\"leg\"><td class=\"code\">"+k+"</td><td>"+esc(GM_LEG[k])+" <span class=\"tag ri\">descontinuado - nunca sugerido</span></td></tr>"}).join("");
  $("#tbGc tbody").innerHTML=Object.keys(GC_OFF).map(function(k){return "<tr><td class=\"code\">"+k+"</td><td>"+esc(GC_OFF[k])+"</td></tr>"}).join("")+Object.keys(GC_LEG).map(function(k){return "<tr class=\"leg\"><td class=\"code\">"+k+"</td><td>"+esc(GC_LEG[k])+" <span class=\"tag ri\">descontinuado"+(SUC_GC[k]?", sucessor valido "+SUC_GC[k]:"")+"</span></td></tr>"}).join("");
 $("#regraImob").innerHTML =
  "<label>Tres requisitos cumulativos</label><ul>"+IMOB.criterios.map(function(c){return "<li>"+esc(c)+"</li>"}).join("")+"</ul>"+
  "<label>Excecoes sempre imobilizado, independente do valor</label><div>"+IMOB.excecoes.map(function(e){return "<span class=\"tag wa\" style=\"margin:2px\">"+esc(e)+"</span>"}).join("")+"</div>"+
  "<label>Substantivos de bem duravel monitorados</label><div class=\"mono\" style=\"color:#68748a\">"+esc(IMOB.bens.join(", "))+"</div>"+
  "<div class=\"note\" style=\"margin-top:12px\">Quando a descricao traz um termo de peca ou componente (parafuso, filtro, kit, reparo...), o alerta nao dispara: trata-se de material de reposicao do Fluxo 302.</div>";
}

/* ================= EVENTOS ================= */
$$("nav button").forEach(function(b){
 b.onclick=function(){
  $$("nav button").forEach(function(x){x.classList.remove("on")}); b.classList.add("on");
  $$("main > section").forEach(function(s){s.classList.add("hidden")});
  $("#t-"+b.dataset.t).classList.remove("hidden");
 };
});
function runUnico(){
 const d=$("#u1").value.trim(); if(!d) return toast("digite o nome do material");
 S.results=[classify(d,$("#u2").value.trim(),$("#u3").value.trim())];
 renderOut(); logResults("analise unitaria");
 $("#stepRes").scrollIntoView({behavior:"smooth",block:"start"});
}
$("#bUnico").onclick=runUnico;
$("#u1").onkeydown=function(e){ if(e.key==="Enter") runUnico() };
$("#bExemplo").onclick=function(){
 $("#u1").value="ROLAMENTO AUT RL 22210E";
 $("#u2").value="ROLAMENTO AUT RL 22210E";
 $("#u3").value="ROLAMENTO; TIPO: AUTOCOMPENSADOR ROLO; DIMENSOES: 50X90X23MM; REFERENCIA: 22210E";
 runUnico();
};
$("#bLote").onclick=function(){
 const lines=$("#loteIn").value.split("\n").map(function(x){return x.trim()}).filter(Boolean);
 if(!lines.length) return toast("cole ao menos um material");
 const items=lines.map(function(l){const p=l.split(";");return {d:(p[0]||"").trim(),c:(p[1]||"").trim(),l:(p[2]||"").trim(),n:ncmClean(p[3]||"")}});
 analyzeList(items).then(function(){ $("#stepRes").scrollIntoView({behavior:"smooth",block:"start"}) });
};
$("#onlyEx").onchange=renderOut;
let upFile=null;
$("#drop").onclick=function(){$("#file").click()};
$("#drop").ondragover=function(e){e.preventDefault();this.classList.add("hot")};
$("#drop").ondragleave=function(){this.classList.remove("hot")};
$("#drop").ondrop=function(e){e.preventDefault();this.classList.remove("hot");upFile=e.dataTransfer.files[0];if(upFile){$("#dropInfo").textContent=upFile.name;$("#bUp").disabled=false}};
$("#file").onchange=function(){upFile=this.files[0];if(upFile){$("#dropInfo").textContent=upFile.name;$("#bUp").disabled=false}};
$("#bUp").onclick=async function(){
 if(!upFile) return;
 try{
  const rows=await readAny(upFile);
  const head=looksHeader(rows[0])?rows[0]:null;
  const m=mapCols(head||rows[0].map(function(_,i){return "c"+i}));
  const data=head?rows.slice(1):rows;
  const items=data.map(function(r){return {d:((r[m.d]||"")+"").trim(), c:m.c>=0?((r[m.c]||"")+"").trim():"", l:m.l>=0?((r[m.l]||"")+"").trim():"", n:m.ncm>=0?ncmClean(r[m.ncm]):""}}).filter(function(x){return x.d||x.c});
  if(!items.length) return toast("nenhuma coluna de descricao reconhecida na planilha");
  toast(items.length+" materiais lidos"+(head?" (colunas reconhecidas)":""));
  await analyzeList(items);
  $("#stepRes").scrollIntoView({behavior:"smooth",block:"start"});
 }catch(e){ toast("nao foi possivel ler o arquivo; salve como XLSX ou CSV e tente novamente") }
};
$("#bXlsModelo").onclick=function(){ if(!needResults())return; download(xlsxWrite(rows4mdg($("#skipEx").checked),"Etapa 5 - Central de Cadastro"),"cadastro-massa-4mdg-"+Date.now()+".xlsx") };
$("#bCsvModelo").onclick=function(){ if(!needResults())return; dlCsv(rows4mdg($("#skipEx").checked),"cadastro-massa-4mdg-"+Date.now()+".csv") };
$("#bXlsFull").onclick=function(){ if(!needResults())return; download(xlsxWrite(rowsFull(),"Analise 302"),"analise-302-"+Date.now()+".xlsx") };
$("#bCopy").onclick=function(){ copyTsv(true) };
$("#skipEx").onchange=buildTsv;
$("#bAudit").onclick=function(){
 if(!needResults()) return;
 const a=auditExport($("#skipEx").checked);
 const revTxt = a.rev.length ? " <b>"+a.rev.length+" celula(s) em branco marcada(s) para revisao</b> (nenhum codigo do catalogo atingiu confianca minima): "+a.rev.slice(0,12).map(function(b){return "linha "+b.linha+" - "+b.campo}).join(" ; ") : "";
 if(!a.bad.length) $("#auditOut").innerHTML="<div class=\"note "+(a.rev.length?"wa":"ok")+"\" style=\"margin-top:10px\"><b>Exportacao conferida:</b> "+a.total+" linha(s). Toda celula preenchida contem codigo puro e presente nos catalogos oficiais (650 Tipo, 663 Grupo de Mercadoria, 627 Grupo de Compradores, NORM e 80 fixos). Nenhuma descricao e nenhum codigo descontinuado."+revTxt+"</div>";
 else $("#auditOut").innerHTML="<div class=\"note ri\" style=\"margin-top:10px\"><b>"+a.bad.length+" celula(s) fora do catalogo oficial:</b> "+a.bad.slice(0,12).map(function(b){return "linha "+b.linha+" - "+b.campo+" = "+esc(b.valor)+" ("+b.motivo+")"}).join(" ; ")+revTxt+"</div>";
};
$("#bBaseFile").onclick=function(){$("#baseFile").click()};
$("#baseFile").onchange=async function(){ const f=this.files[0]; if(!f) return; $("#baseFileInfo").textContent=f.name;
 try{ await loadBaseFile(f,$("#baseAppend").checked) }catch(e){ toast("falha ao ler a base; verifique se e XLSX ou CSV") } };
$("#bAdm").onclick=function(){
 if($("#admPwd").value!==S.pwd) return toast("senha incorreta");
 $("#admLock").classList.add("hidden"); $("#admArea").classList.remove("hidden"); try{ catUnlock() }catch(e){}
 $("#admInfo").textContent="Base: "+S.base.length+" registros | termos aprendidos: "+S.lexBase.size+" | substantivos aprendidos: "+S.headBase.size+" | correcoes: "+S.fb.length+" | log: "+S.log.length+" eventos.";
};
$("#bRelearn").onclick=function(){ buildIndex(); setBaseState(); toast("aprendizado reconstruido a partir da base") };
$("#bExpBase").onclick=function(){ dlCsv([["descricao_material","descricao_curta","descricao_longa","tipo_material","grupo_mercadorias","grupo_compradores","grupo_categoria_item","setor_atividade","codigo"]].concat(S.base.map(function(r){return [r.d,r.c,r.l,r.tm,r.gm,r.gc,r.gci,r.sa,r.cod]})),"base-302.csv") };
$("#bExpFb").onclick=function(){ dlCsv([["descricao","tipo","grupo_mercadorias","grupo_compradores","quando"]].concat(S.fb.map(function(f){return [f.d,f.tm,f.gm,f.gc,new Date(f.t).toLocaleString("pt-BR")]})),"correcoes.csv") };
$("#bClrBase").onclick=async function(){ if(!confirm("Limpar a base carregada? Suas correcoes e o log serao mantidos.")) return; S.base=[]; buildIndex(); await saveBase(); setBaseState(); toast("base limpa") };
$("#bChgPwd").onclick=async function(){ const p=prompt("Nova senha do administrador:"); if(!p) return; S.pwd=p; await kvSet("pwd",p); toast("senha atualizada") };
$("#bExpLog").onclick=function(){ dlCsv([["quando","descricao","tipo","gm","gc","confianca","origem","evento"]].concat(S.log.map(function(x){return [new Date(x.t).toLocaleString("pt-BR"),x.d,x.tm,x.gm,x.gc,x.conf,x.src,x.ev]})),"log-decisoes.csv") };
$("#bClrLog").onclick=async function(){ if(!confirm("Limpar o log de decisoes?")) return; S.log=[]; await saveLog(); renderLog(); toast("log limpo") };

/* ================= CATALOGOS OFICIAIS: CARGA E PAINEL ================= */
function renderCat(){
 const nomes={tm:"Tipo de Material (650)", gm:"Grupo de Mercadoria (663)", gc:"Grupo de Compradores (627)"};
 const qtd={tm:TIPOS.length, gm:Object.keys(GM_OFF).length, gc:Object.keys(GC_OFF).length};
 const kp=$("#catKpis"); if(kp) kp.innerHTML=["tm","gm","gc"].map(function(f){
  return "<div class=\"kpi\"><b>"+qtd[f]+"</b><span>"+esc(nomes[f])+"</span><span class=\"mono\" style=\"text-transform:none\">"+esc(CATSRC[f])+"</span></div>" }).join("");
 const inf=$("#catInfo");
 if(inf) inf.innerHTML="<b>Codigos aceitos hoje &mdash; Tipo:</b> "+TIPOS.map(function(t){return t[0]}).join(", ")
  +" &nbsp;<b>Grupo de Mercadoria:</b> "+Object.keys(GM_OFF).join(", ")
  +" &nbsp;<b>Grupo de Compradores:</b> "+Object.keys(GC_OFF).join(", ")
  +"<br>Qualquer outro codigo que aparecer no historico da base e tratado como descontinuado: nao e sugerido, nao entra na lista de edicao e sai em branco na exportacao, marcado como REVISAR.";
 if(typeof renderRegras==="function") renderRegras();
}
async function loadCatFiles(files){
 let msg=[];
 for(let i=0;i<files.length;i++){
  const f=files[i];
  try{
   const rows=await readAny(f);
   const pairs=catPairs(rows);
   const kind=detectCat(pairs);
   if(!kind || !pairs.length){ msg.push(f.name+": nenhum codigo reconhecido"); continue }
   const n=applyCat(kind, pairs, f.name);
   msg.push(f.name+" -> "+({tm:"Tipo de Material (650)",gm:"Grupo de Mercadoria (663)",gc:"Grupo de Compradores (627)"}[kind])+": "+n+" codigos");
  }catch(e){ msg.push(f.name+": falha na leitura") }
 }
 buildIndex(); renderCat(); try{ renderRegras() }catch(e){} try{ renderCatAdm() }catch(e){} renderOut();
 toast(msg.join(" | "));
}
$("#bCatFile").onclick=function(){ $("#catFile").click() };
$("#catFile").onchange=async function(){ const fs=this.files; if(!fs||!fs.length) return; await loadCatFiles(fs); this.value="" };

/* ================= MANUTENCAO DOS CATALOGOS PELO ADMINISTRADOR (v7) ================= */
const CATF_NOME = {tm:"Tipo de Material (650)", gm:"Grupo de Mercadoria (663)", gc:"Grupo de Compradores (627)"};
function catAdmSel(){ const s=$("#catAdmF"); return s? s.value : "gm" }
function catShape(f,c){
 if(f==="gm") return /^\d{5,6}$/.test(c);
 if(f==="gc") return /^[A-Z]{1,3}[0-9]{1,2}$/.test(c);
 return /^[A-Z]{3,5}$/.test(c);
}
function catSet(f,code,desc){
 code=String(code||"").toUpperCase().trim(); desc=String(desc||"").trim();
 if(!code){ toast("informe o codigo"); return false }
 if(!catShape(f,code)){ toast("formato de codigo invalido para "+CATF_NOME[f]); return false }
 if(f==="tm"){
  const i=TIPOS.findIndex(function(t){return t[0]===code});
  if(i>=0) TIPOS[i]=[code, desc||TIPOS[i][1], TIPOS[i][2]];
  else TIPOS.push([code, desc||code, "incluido pelo administrador"]);
 } else if(f==="gm"){
  GM_OFF[code]=desc||GM_OFF[code]||code;
  if(DISC.gm[code]) delete DISC.gm[code];
  if(DISC_SHEET.gm[code]) delete DISC_SHEET.gm[code];
  if(GM_LEG[code]) delete GM_LEG[code];
  delete APX.gm[code];
 } else {
  GC_OFF[code]=desc||GC_OFF[code]||code;
  if(DISC.gc[code]) delete DISC.gc[code];
  if(DISC_SHEET.gc[code]) delete DISC_SHEET.gc[code];
  if(GC_LEG[code]) delete GC_LEG[code];
  Object.keys(APX.gc).forEach(function(k){ if(k.indexOf(code+"|")===0) delete APX.gc[k] });
 }
 CATSRC[f]="ajustado pelo administrador sobre "+(CATSRC[f]||SHEET_SRC[f]);
 catRefresh("codigo "+code+" gravado em "+CATF_NOME[f]);
 return true;
}
function catDel(f,code){
 code=String(code||"").toUpperCase().trim();
 if(!code){ toast("informe o codigo"); return false }
 if(f==="tm") TIPOS=TIPOS.filter(function(t){return t[0]!==code});
 else if(f==="gm"){ if(!GM_OFF[code]){ toast("codigo nao esta na lista ativa"); return false } discMark("gm",code,"removido pelo administrador") }
 else { if(!GC_OFF[code]){ toast("codigo nao esta na lista ativa"); return false } discMark("gc",code,"removido pelo administrador") }
 if(f==="gm") APX.gm={}; if(f==="gc") APX.gc={};
 CATSRC[f]="ajustado pelo administrador sobre "+(CATSRC[f]||SHEET_SRC[f]);
 catRefresh("codigo "+code+" removido de "+CATF_NOME[f]+" - deixa de ser sugerido e de ser exportado");
 return true;
}
function catRefresh(msg){
 rebuildViews(); saveCat();
 try{ sanitizePriors() }catch(e){}
 try{ buildIndex() }catch(e){}
 try{ renderCat() }catch(e){}
 try{ renderRegras() }catch(e){}
 try{ renderCatAdm() }catch(e){}
 try{ renderOut() }catch(e){}
 if(msg) toast(msg);
}
function renderCatAdm(){
 const tb=$("#tbCatAdm tbody"); if(!tb) return;
 const f=catAdmSel();
 let rows=[];
 if(f==="tm") rows=TIPOS.map(function(t){ return [t[0], t[1]] });
 else if(f==="gm") rows=Object.keys(GM_OFF).sort().map(function(k){ return [k, GM_OFF[k]+(GM_FAM[k]?" - familia "+GM_FAM[k]:"")] });
 else rows=Object.keys(GC_OFF).sort().map(function(k){ return [k, GC_OFF[k]] });
 tb.innerHTML = rows.map(function(r){
  return "<tr><td class=\"code\">"+esc(r[0])+"</td><td>"+esc(r[1])+"</td><td>"+
   "<button class=\"btn gho cae\" data-c=\""+esc(r[0])+"\">editar</button> "+
   "<button class=\"btn gho cax\" data-c=\""+esc(r[0])+"\">remover</button></td></tr>";
 }).join("") || "<tr><td colspan=\"3\">lista vazia - sugerimos recarregar a planilha oficial</td></tr>";
 tb.querySelectorAll("button.cae").forEach(function(b){ b.onclick=function(){
  const c=b.getAttribute("data-c");
  $("#catAdmC").value=c;
  $("#catAdmD").value=(f==="tm"? (TIPOS.find(function(t){return t[0]===c})||["",""])[1] : (f==="gm"? GM_OFF[c] : GC_OFF[c])) || "";
 }});
 tb.querySelectorAll("button.cax").forEach(function(b){ b.onclick=function(){ catDel(f, b.getAttribute("data-c")) }});
 const inf=$("#catAdmInfo");
 if(inf) inf.innerHTML="<b>"+rows.length+" codigo(s) ativos em "+esc(CATF_NOME[f])+"</b> - lista em uso: "+esc(CATSRC[f]||SHEET_SRC[f])+
  ".<br>Nenhum codigo fora desta lista pode ser sugerido, escolhido ou exportado. Codigo do historico que nao consta aqui e aproximado pelo mais proximo por familia, substantivo ou categoria, e a aproximacao fica registrada no painel de transparencia.";
}
function catUnlock(){
 const L=$("#catLock"), A=$("#catAdm");
 if(L) L.classList.add("hidden");
 if(A) A.classList.remove("hidden");
 renderCatAdm();
 try{ renderEqAp() }catch(e){}
}
/* ===== AUTOANALISE (materiais de teste das versoes anteriores) ===== */
const TESTE_V6 = [
 ["ROLAMENTO SKF 6205 2RS","ROLAMENTO; MARCA: SKF; MODELO: 6205 2RS; APLICACAO: MOTOR ELETRICO",""],
 ["PARAFUSO SEXTAVADO M12 X 50 INOX","PARAFUSO SEXTAVADO; ROSCA: M12; COMPRIMENTO: 50MM; MATERIAL: INOX",""],
 ["FILTRO DE OLEO HIDRAULICO","FILTRO; TIPO: OLEO HIDRAULICO; APLICACAO: SISTEMA HIDRAULICO",""],
 ["FILTRO DE COMBUSTIVEL MWM","FILTRO; TIPO: COMBUSTIVEL; APLICACAO: MOTOR MWM",""],
 ["OLEO LUBRIFICANTE 15W40 20L","OLEO LUBRIFICANTE; VISCOSIDADE: 15W40; EMBALAGEM: 20L",""],
 ["LUVA DE PROTECAO NITRILICA TAM 9","LUVA; MATERIAL: NITRILICA; TAMANHO: 9; USO: PROTECAO",""],
 ["PAPEL A4 75G RESMA 500 FOLHAS","PAPEL; FORMATO: A4; GRAMATURA: 75G; EMBALAGEM: RESMA",""],
 ["CIMENTO CP II 32 SACO 50KG","CIMENTO; TIPO: CP II 32; EMBALAGEM: SACO 50KG",""],
 ["TINTA ACRILICA BRANCA 18L","TINTA; TIPO: ACRILICA; COR: BRANCA; EMBALAGEM: 18L",""],
 ["CABO FLEXIVEL 2,5MM AZUL","CABO; TIPO: FLEXIVEL; SECAO: 2,5MM; COR: AZUL",""]
];
function autoAnalise(){
 const rows=[]; let fora=0, vazioSemRev=0, aprox=0, revisar=0, exato=0;
 TESTE_V6.forEach(function(t){
  let r=null;
  try{ r=classify(t[0], t[0], t[1]||"", t[2]||"") }catch(e){ r=null }
  if(!r) return;
  ["tm","gm","gc"].forEach(function(f){
   const c=fldVal(r,f), mk=matchKind(r,f);
   if(c && !codeOk(f,c)) fora++;
   if(!c){ revisar++; if(!(r.rev&&r.rev.length) && !r.flag) vazioSemRev++ }
   else if(mk.k==="aprox") aprox++; else if(mk.k==="exato") exato++;
  });
  rows.push([t[0], r.tm||"REVISAR", r.gm||"REVISAR", r.gc||"REVISAR", Math.round((r.conf||0)*100)+"%",
             matchKind(r,"gm").txt, matchKind(r,"gc").txt]);
 });
 /* varredura de todo o historico fora das planilhas: a aproximacao nunca pode devolver codigo invalido */
 let apxOk=0, apxNulo=0, apxRuim=0;
 Object.keys(GM_LEG).forEach(function(k){ const a=sucOf("gm",k); if(!a) apxNulo++; else if(codeOk("gm",a)) apxOk++; else apxRuim++ });
 Object.keys(GC_LEG).forEach(function(k){ const a=sucOf("gc",k); if(!a) apxNulo++; else if(codeOk("gc",a)) apxOk++; else apxRuim++ });
 const res={fora:fora, vazioSemRev:vazioSemRev, aprox:aprox, exato:exato, revisar:revisar, apxOk:apxOk, apxNulo:apxNulo, apxRuim:apxRuim, rows:rows};
 const box=$("#autoOut");
 if(box){
  box.innerHTML =
   "<div class=\"note "+((fora||vazioSemRev||apxRuim)?"wa":"ok")+"\"><b>Autoanalise dos "+rows.length+" materiais de teste</b><br>"+
   "codigos fora das listas: <b>"+fora+"</b> - campos vazios sem sinalizacao REVISAR: <b>"+vazioSemRev+"</b> - "+
   "correspondencia exata: <b>"+exato+"</b> - por aproximacao: <b>"+aprox+"</b> - sugeridos para REVISAR: <b>"+revisar+"</b><br>"+
   "historico fora das planilhas: "+apxOk+" com equivalente valido, "+apxNulo+" sem equivalente (segue para REVISAR), "+apxRuim+" invalidos.<br>"+
   "Listas consultadas: "+esc(listaDe("gm"))+" | "+esc(listaDe("gc"))+"</div>"+
   "<div class=\"tw\" style=\"max-height:340px;margin-top:10px\"><table><thead><tr><th>Material</th><th>Tipo</th><th>Grupo Mercadoria</th><th>Grupo Compradores</th><th>Confianca</th><th>Correspondencia GM</th><th>Correspondencia GC</th></tr></thead><tbody>"+
   rows.map(function(r){ return "<tr><td>"+esc(r[0])+"</td><td class=\"code\">"+esc(r[1])+"</td><td class=\"code\">"+esc(r[2])+"</td><td class=\"code\">"+esc(r[3])+"</td><td>"+esc(r[4])+"</td><td><span class=\"cx\">"+esc(r[5])+"</span></td><td><span class=\"cx\">"+esc(r[6])+"</span></td></tr>" }).join("")+
   "</tbody></table></div>";
 }
 return res;
}
$("#bCatReset").onclick=function(){ resetCat(); buildIndex(); renderCat(); renderOut(); catRefresh("catalogos das planilhas oficiais restaurados") };
$("#bCatAdm").onclick=function(){ if($("#catAdmPwd").value!==S.pwd) return toast("senha incorreta"); catUnlock(); toast("edicao dos catalogos liberada") };
$("#catAdmPwd").onkeydown=function(e){ if(e.key==="Enter") $("#bCatAdm").click() };
$("#catAdmF").onchange=function(){ renderCatAdm() };
$("#bCarAdd").onclick=function(){ if(carSet($("#carGm").value,$("#carGc").value,$("#carCa").value,$("#carEx").value)){ $("#carGm").value=""; $("#carGc").value=""; $("#carCa").value=""; $("#carEx").value="" } };
$("#bCarDel").onclick=function(){ carDel($("#carGm").value,$("#carGc").value) };
$("#bCarFile").onclick=function(){ $("#carFile").click() };
$("#carFile").onchange=function(){ if(this.files && this.files[0]) carImportFile(this.files[0]); this.value="" };
$("#bCarJson").onclick=function(){ try{ const b=new Blob([JSON.stringify({rows:CAR.rows,src:CAR.src,dt:new Date().toISOString()},null,1)],{type:"application/json"}); const a=document.createElement("a"); a.href=URL.createObjectURL(b); a.download="regra_classe_avaliacao.json"; a.click(); toast("regra exportada em JSON") }catch(e){ toast("nao foi possivel exportar a regra") } };
$("#bCarClear").onclick=function(){ CAR.rows=[]; CAR.src="tabela limpa pelo administrador"; carApply("tabela de regra limpa") };
$("#bCarSeed").onclick=function(){ carSeedApply(true) };
$("#bCatAdd").onclick=function(){ if(catSet(catAdmSel(), $("#catAdmC").value, $("#catAdmD").value)){ $("#catAdmC").value=""; $("#catAdmD").value="" } };
$("#bCatDel").onclick=function(){ if(catDel(catAdmSel(), $("#catAdmC").value)) $("#catAdmC").value="" };
$("#bCatAdmFile").onclick=function(){ $("#catFile").click() };
$("#bCatAdmReset").onclick=function(){ resetCat(); catRefresh("catalogos das planilhas oficiais anexadas restaurados") };
$("#bAuto").onclick=function(){ autoAnalise() };


/* ================= EQUIVALENCIA DE APLICACAO - PAINEL DO ADMINISTRADOR (v7.3) ================= */
let EQAP_ONLY_PEND=false;
function eqapOrigTxt(e){ return e.o==="admin"? "administrador" : ((!e.gm||!e.gc)? "a preencher" : "base 302") }
function renderEqAp(){
 const tb=$("#tbEqAp"); if(!tb) return;
 const body=tb.querySelector("tbody"); if(!body) return;
 const lista=(EQAP_ONLY_PEND? eqapPend() : EQAP).slice().sort(function(a,b){ return (b.n||0)-(a.n||0) });
 body.innerHTML = lista.map(function(e){
  const pend=(!e.gm||!e.gc);
  return "<tr"+(pend?" style=\"background:rgba(234,179,8,.10)\"":"")+"><td><b>"+esc(e.t)+"</b></td><td>"+esc(e.cls||"-")+"</td><td>"+esc(e.gm||"a preencher")+"</td><td>"+esc(e.gc||"a preencher")+"</td><td class=\"hint\">"+esc(eqapEv(e))+"</td><td>"+esc(eqapOrigTxt(e))+"</td>"+
   "<td><button class=\"btn gho\" data-eqed=\""+esc(e.t)+"\">editar</button> <button class=\"btn gho\" data-eqrm=\""+esc(e.t)+"\">remover</button></td></tr>";
 }).join("") || "<tr><td colspan=\"7\" class=\"hint\">nenhuma equivalencia carregada</td></tr>";
 body.querySelectorAll("[data-eqed]").forEach(function(b){ b.onclick=function(){
  const e=EQAP_IX[eqapNorm(this.getAttribute("data-eqed"))]; if(!e) return;
  if($("#eqApT")) $("#eqApT").value=e.t; if($("#eqApC")) $("#eqApC").value=e.cls||"";
  if($("#eqApGm")) $("#eqApGm").value=e.gm||""; if($("#eqApGc")) $("#eqApGc").value=e.gc||"";
 }});
 body.querySelectorAll("[data-eqrm]").forEach(function(b){ b.onclick=function(){ eqapDelT(this.getAttribute("data-eqrm")) }});
 const inf=$("#eqApInfo");
 if(inf) inf.innerHTML="<b>"+EQAP.length+"</b> equivalencias carregadas, sendo <b>"+eqapPend().length+"</b> a preencher e <b>"+EQAP.filter(function(e){return e.o==="admin"}).length+"</b> preenchidas por voce. Formato aceito na planilha: termo, familia (MAQ/VEI/EQP), GM, GC - uma linha por termo, cabecalho opcional. Sugestao: comece pelas linhas destacadas, sao as que a base 302 nao resolve sozinha.";
}
function eqapRefresh(msg){ eqapSave(); try{ renderEqAp() }catch(e){} try{ renderOut() }catch(e){} if(msg) toast(msg) }
function eqapSet(t,cls,gm,gc,orig){
 t=String(t||"").trim(); if(!t){ toast("informe o termo de aplicacao ou o modelo"); return false }
 cls=String(cls||"").toUpperCase().trim(); gm=String(gm||"").toUpperCase().trim(); gc=String(gc||"").toUpperCase().trim();
 if(cls && !EQAP_CLS[cls]){ toast("familia deve ser MAQ, VEI ou EQP"); return false }
 if(gm && !codeOk("gm",gm)){ toast(gm+" nao consta no catalogo 663 em uso - sugerimos incluir o codigo no catalogo antes"); return false }
 if(gc && !codeOk("gc",gc)){ toast(gc+" nao consta no catalogo 627 em uso - sugerimos incluir o codigo no catalogo antes"); return false }
 const k=eqapNorm(t), ex=EQAP_IX[k];
 if(ex){ ex.t=t; ex.cls=cls; ex.gm=gm; ex.gc=gc; ex.o=orig||"admin" }
 else EQAP.push({t:t,cls:cls,gm:gm,gc:gc,n:0,p:0,pg:0,pc:0,o:orig||"admin"});
 return true;
}
function eqapDelT(t){
 const k=eqapNorm(t), i=EQAP.findIndex(function(e){ return eqapNorm(e.t)===k });
 if(i<0){ toast("termo nao encontrado na tabela de apoio"); return false }
 EQAP.splice(i,1); eqapRefresh("equivalencia removida: "+t); return true;
}
function eqapFromRows(rows){
 let ok=0, ig=0;
 (rows||[]).forEach(function(r){
  if(!r) return;
  const c=(Array.isArray(r)? r : Object.keys(r).map(function(k){return r[k]})).map(function(v){ return String(v==null?"":v).trim() });
  if(!c.length || !c[0]) return;
  const h0=calNorm(c[0]);
  if(/^(TERMO|APLICACAO|MODELO)/.test(h0) && c.length>1 && /(FAMILIA|CLASSE|GRUPO|GM|GC)/.test(calNorm(c.slice(1).join(" ")))) return;
  const cls=calNorm(c[1]||"").slice(0,3), gm=(c[2]||"").toUpperCase(), gc=(c[3]||"").toUpperCase();
  if(eqapSet(c[0], EQAP_CLS[cls]?cls:"", codeOk("gm",gm)?gm:"", codeOk("gc",gc)?gc:"", "admin")) ok++; else ig++;
 });
 eqapRefresh(ok+" equivalencia(s) atualizada(s) pela planilha"+(ig?", "+ig+" linha(s) ignorada(s) por codigo fora dos catalogos":""));
}
function eqapTpl(){
 const rows=[["termo","familia (MAQ/VEI/EQP)","grupo de mercadoria (663)","grupo de compradores (627)","evidencia na base 302"]];
 eqapPend().sort(function(a,b){ return (b.n||0)-(a.n||0) }).forEach(function(e){ rows.push([e.t, e.cls||"", e.gm||"", e.gc||"", eqapEv(e)]) });
 const csv=rows.map(function(r){ return r.map(function(v){ return '"'+String(v==null?"":v).replace(/"/g,'""')+'"' }).join(";") }).join("\r\n");
 try{
  const b=new Blob(["\ufeff"+csv],{type:"text/csv;charset=utf-8"});
  const u=URL.createObjectURL(b), a=document.createElement("a");
  a.href=u; a.download="equivalencia_aplicacao_a_preencher.csv"; document.body.appendChild(a); a.click();
  setTimeout(function(){ URL.revokeObjectURL(u); a.remove() },300);
  toast("planilha das linhas a preencher gerada - devolva pelo botao Carregar equivalencias por planilha");
 }catch(e){ toast("nao foi possivel gerar o arquivo neste navegador") }
}
function eqapTranspHtml(r){
 const e=r&&r.eqap; if(!e) return "";
 const li=[];
 li.push("termo casado: <b>"+esc(e.t)+"</b>"+(e.cls?" ("+esc(EQAP_CLS[e.cls]||e.cls)+")":"")+", lido do campo "+esc(e.campo||"descricao"));
 li.push("evidencia: "+esc(eqapEv(e)));
 li.push("Grupo de Mercadoria: "+esc(e.apGm||"sem apoio"));
 li.push("Grupo de Compradores: "+esc(e.apGc||"sem apoio"));
 return "<div class=\"note\" style=\"margin-top:8px\"><b>Apoio por equivalencia de aplicacao</b> (apoio, nao regra - editavel no painel administrativo)<ul style=\"margin:6px 0 0 18px\"><li>"+li.join("</li><li>")+"</li></ul></div>";
}
function eqapLoteHtml(){
 try{
  const rs=(S.results||[]).filter(function(r){ return r&&r.eqap });
  if(!rs.length) return "";
  const ap=rs.filter(function(r){ return /aplicado como apoio/.test((r.eqap.apGm||"")+" "+(r.eqap.apGc||"")) }).length;
  const pe=rs.filter(function(r){ return /disponivel para preencher/.test((r.eqap.apGm||"")+" "+(r.eqap.apGc||"")) }).length;
  return "<div class=\"note\" style=\"margin-top:8px\">Equivalencia de aplicacao consultada em <b>"+rs.length+"</b> material(is) do lote: apoio efetivamente aplicado em <b>"+ap+"</b> (onde a base 302 nao tinha evidencia) e <b>"+pe+"</b> com termo reconhecido mas sem codigo de apoio preenchido - se quiser, complete essas linhas na aba Regras e Catalogos.</div>";
 }catch(e){ return "" }
}
if($("#bEqApAdd")) $("#bEqApAdd").onclick=function(){
 if(eqapSet($("#eqApT").value, $("#eqApC").value, $("#eqApGm").value, $("#eqApGc").value, "admin")){
  eqapRefresh("equivalencia de apoio salva - passa a valer agora, sem recarregar a pagina");
  $("#eqApT").value=""; $("#eqApGm").value=""; $("#eqApGc").value=""; $("#eqApC").value="";
 }
};
if($("#bEqApDel")) $("#bEqApDel").onclick=function(){ if(eqapDelT($("#eqApT").value)) $("#eqApT").value="" };
if($("#bEqApFile")) $("#bEqApFile").onclick=function(){ $("#eqApFile").click() };
if($("#bEqApTpl")) $("#bEqApTpl").onclick=function(){ eqapTpl() };
if($("#bEqApPend")) $("#bEqApPend").onclick=function(){ EQAP_ONLY_PEND=!EQAP_ONLY_PEND; this.textContent=EQAP_ONLY_PEND?"Ver todas as equivalencias":"Ver so as linhas a preencher"; renderEqAp() };
if($("#bEqApReset")) $("#bEqApReset").onclick=function(){ EQAP=eqapDefault(); eqapRefresh("equivalencias da base 302 restauradas") };
if($("#eqApFile")) $("#eqApFile").onchange=async function(){
 const fs=this.files; if(!fs||!fs.length) return;
 try{ for(const f of fs){ const rows=await readAny(f); eqapFromRows(rows) } }
 catch(e){ toast("nao foi possivel ler a planilha de equivalencias") }
 this.value="";
};


/* ================================================================
   REGRA DE CLASSE DE AVALIACAO (Grupo de Mercadoria + Grupo de Compradores)
   Acrescimo de regra, nunca substituicao de catalogo: os catalogos 650/663/627
   e as bases embutidas continuam exatamente como estao. A tabela comeca vazia
   e e alimentada pelo administrador (planilha XLSX/CSV ou linha por linha),
   fica no localStorage e passa a valer na hora, sem recarregar a pagina.
   ================================================================ */
var CAR_LS="mdm302_regra_ca_v1";
var CAR={rows:[],byGm:{},byPair:{},src:"nenhuma planilha de regra importada",dt:""};
/* ====== v7.8: SEMENTE DE REGRA EMBUTIDA - planilha Automatizacao.xlsx (90001/90701) ======
   Grupo de Mercadoria (+ Grupo de Compradores) -> Classe de Avaliacao, tabela inicial.
   Acrescimo, nao substituicao. Semente reconhecida por assinatura das linhas (v7.7+),
   nao pelo texto da origem - foi o que deixava tabela antiga (9001/3037) no navegador.
   Linhas da planilha NAO embutidas, aguardando validacao da Controladoria:
     fora: 90710 - Uniforme-Mat | UF1 -Uniformes -> CA 3051 - Uniforme-Mat  (GM 90710 nao existe no catalogo 663)
*/
var CAR_SEED_V="AUTOMATIZACAO-2026-09-23-R4";
var CAR_SEED_LS="mdm302_regra_ca_seed_v";
var CAR_SEED_HS="mdm302_regra_ca_seed_h";
var CAR_SEED_SRC="planilha Automatizacao.xlsx embutida neste arquivo";
var CAR_SEED=[
 {"gm": "90703", "gc": "AL1", "ca": "3118", "ex": "Alimentos (Café, açúcar, adoçante, chá, biscoitos, água mineral, arroz, feijão, óleo de cozinha, sal, leite em pó, , etc)"},
 {"gm": "90001", "gc": "MR1", "ca": "3052", "ex": "Manutenção Equipamentos (Fusíveis, disjuntores, contatores, relés, fiação elétrica, rolamentos, correias dentadas, retentores, engrenagens, filtros industriais, etc)"},
 {"gm": "90003", "gc": "MR1", "ca": "3053", "ex": "Manutenção Veiculos (Pastilhas de freio, discos de freio, velas de ignição, filtros de óleo, amortecedores, pneus, palhetas do limpador de para-brisa, lâmpadas para faróis, CARRO, MOTO OU CAMINHÃO, , etc)"},
 {"gm": "90002", "gc": "MR1", "ca": "3052", "ex": "Manutenção Maquinas (Facas de colhedora, dentes de escavadeira, esteiras, braços hidráulicos, mangueiras de alta pressão, conexões, vedações, pistões, válvulas solenoides, etc)"},
 {"gm": "90004", "gc": "MR1", "ca": "3054", "ex": "Manutenção e Reparo Predial (Tintas, rolos, pincéis, lixas, fita crepe, massa corrida, tubos de PVC, conexões, torneiras, silicone, cimento, argamassa, lâmpadas LED, tomadas, etc)"},
 {"gm": "90708", "gc": "LB1", "ca": "3040", "ex": "Óleos e Lubrificantes (Óleo de motor 5W30, óleo de motor 15W40, fluido de freio, óleo hidráulico, graxa azul, desengripante em spray, vaselina sólida, etc)"},
 {"gm": "90711", "gc": "EP1", "ca": "3035", "ex": "Epis (Botas de segurança, capacetes, óculos de proteção, luvas nitrílicas, luvas de raspa, protetor auricular, máscaras PFF2, cintos de segurança para altura, etc)"},
 {"gm": "90705", "gc": "ML1", "ca": "3038", "ex": "Limpeza (Detergente, desinfetante, água sanitária, limpa-vidros, sabão líquido, vassouras, mops, panos de microfibra, sacos de lixo, esponjas, etc)"},
 {"gm": "90707", "gc": "ME2", "ca": "3043", "ex": "Material Escritorio (Canetas, cadernos, blocos de notas, papel A4, marcadores de texto, pastas arquivo, clipes, grampeadores, fita adesiva, tesouras, etc)"},
 {"gm": "90709", "gc": "IS1", "ca": "3043", "ex": "Informatica (Mouse, teclado, fone de ouvido, webcam, cabos de rede RJ45, cabos HDMI, cabos USB, extensões elétricas, adaptadores, pen drives, etc)"},
 {"gm": "90701", "gc": "ME2", "ca": "3037", "ex": "Ferramentas (Chaves de fenda, chaves de boca, alicates, martelos, trenas, brocas para metal, brocas para concreto, discos de corte, serras copo, etc)"}
];
function carSeedRows(){ return CAR_SEED.map(function(r){ return {gm:r.gm,gc:r.gc,ca:r.ca,ex:r.ex} }) }
function carSeedHash(rows){
 var t="",h=0,i,j; rows=rows||[];
 for(i=0;i<rows.length;i++){ var r=rows[i]||{}; t+=String(r.gm||"")+"|"+String(r.gc||"")+"|"+String(r.ca||"")+";" }
 for(j=0;j<t.length;j++){ h=(h*31+t.charCodeAt(j))|0 }
 return String(h)+"."+rows.length;
}
function carGmFora(){
 var l=[],i,g;
 for(i=0;i<CAR.rows.length;i++){ g=String(CAR.rows[i].gm||"").trim(); if(g && !GM_OFF[g] && l.indexOf(g)<0) l.push(g) }
 return l;
}
function carSeedApply(av){
 CAR.rows=carSeedRows(); CAR.src=CAR_SEED_SRC;
 try{ localStorage.setItem(CAR_SEED_LS,CAR_SEED_V); localStorage.setItem(CAR_SEED_HS,carSeedHash(CAR.rows)) }catch(e){}
 carApply(av? ("regra da planilha embutida aplicada - "+CAR.rows.length+" linhas") : "");
}
function carBoot(){
 var ok=false; try{ ok=carLoad() }catch(e){ ok=false }
 if(!ok){ carSeedApply(false); return }
 var sv="",sh="";
 try{ sv=localStorage.getItem(CAR_SEED_LS)||""; sh=localStorage.getItem(CAR_SEED_HS)||"" }catch(e){}
 if(sv===CAR_SEED_V && sh===carSeedHash(CAR.rows)) return;
 var deSemente=(sh && sh===carSeedHash(CAR.rows)) || /^planilha Automatiza/i.test(String(CAR.src||""))
               || String(CAR.src||"").indexOf("embutida neste arquivo")>=0;
 if(deSemente){
  carSeedApply(false);
  setTimeout(function(){ try{ toast("regra de classe de avaliacao atualizada para a desta versao: "+CAR.rows.length+" linhas da planilha embutida") }catch(e){} },160);
  return;
 }
 setTimeout(function(){ try{
  var f=carGmFora();
  toast(f.length? ("regra salva neste navegador tem "+f.length+" grupo(s) fora do catalogo 663 ("+f.join(", ")+") - use Restaurar a regra da planilha embutida")
                : "existe regra embutida mais nova neste arquivo - use Restaurar a regra da planilha embutida");
 }catch(e){} },160);
}
function carKey(gm,gc){ return String(gm||"").trim()+"|"+String(gc||"").trim() }
function carRebuild(){
 CAR.byGm={}; CAR.byPair={};
 for(var i=0;i<CAR.rows.length;i++){
  var r=CAR.rows[i]||{}, gm=String(r.gm||"").trim(), gc=String(r.gc||"").trim(), ca=String(r.ca||"").trim();
  if(!gm) continue;
  if(!CAR.byGm[gm]) CAR.byGm[gm]={cas:{},gcs:{},ex:""};
  if(ca) CAR.byGm[gm].cas[ca]=(CAR.byGm[gm].cas[ca]||0)+1;
  if(gc) CAR.byGm[gm].gcs[gc]=ca||"";
  if(r.ex && !CAR.byGm[gm].ex) CAR.byGm[gm].ex=r.ex;
  if(gc && ca) CAR.byPair[carKey(gm,gc)]={ca:ca,ex:r.ex||""};
 }
 return CAR;
}
function carSave(){
 try{ localStorage.setItem(CAR_LS, JSON.stringify({rows:CAR.rows,src:CAR.src,dt:new Date().toISOString()})) }catch(e){}
}
function carLoad(){
 try{
  var s=localStorage.getItem(CAR_LS); if(!s) return false;
  var o=JSON.parse(s);
  if(o && o.rows && o.rows.length){ CAR.rows=o.rows; CAR.src=o.src||CAR.src; CAR.dt=o.dt||""; carRebuild(); return true }
 }catch(e){}
 return false;
}
function carApply(msg){
 carRebuild(); carSave();
 try{ renderCaRegra() }catch(e){}
 try{ if(S.results && S.results.length) renderOut() }catch(e){}
 if(msg){ try{ toast(msg) }catch(e){} }
}
function carSet(gm,gc,ca,ex){
 gm=String(gm||"").trim(); gc=String(gc||"").trim(); ca=String(ca||"").trim();
 if(!gm) { try{ toast("informe o Grupo de Mercadoria") }catch(e){} return false }
 var k=carKey(gm,gc), achou=false;
 for(var i=0;i<CAR.rows.length;i++){ if(carKey(CAR.rows[i].gm,CAR.rows[i].gc)===k){ CAR.rows[i]={gm:gm,gc:gc,ca:ca,ex:ex||CAR.rows[i].ex||""}; achou=true; break } }
 if(!achou) CAR.rows.push({gm:gm,gc:gc,ca:ca,ex:ex||""});
 CAR.src="editada no painel do administrador";
 carApply(achou?"linha de regra atualizada":"linha de regra incluida");
 return true;
}
function carDel(gm,gc){
 var k=carKey(gm,gc), n=CAR.rows.length;
 CAR.rows=CAR.rows.filter(function(r){ return carKey(r.gm,r.gc)!==k });
 if(CAR.rows.length===n){ try{ toast("linha nao encontrada na tabela de regra") }catch(e){} return false }
 carApply("linha de regra removida"); return true;
}
/* derivacao: par exato primeiro, depois o grupo de mercadoria;
   grupo sem regra NAO recebe classe automatica */
function carPick(gm,gc){
 if(!CAR.rows.length) return null;
 var g=String(gm||"").trim(); if(!g) return null;
 var byG=CAR.byGm[g]; if(!byG) return null;
 var p=CAR.byPair[carKey(g,gc)];
 if(p && p.ca) return {ca:p.ca, conf:0.95, cands:[],
   via:"regra da planilha: Grupo de Mercadoria "+g+" + Grupo de Compradores "+String(gc||"").trim()+" -> Classe de Avaliacao "+p.ca,
   lista:"tabela de regra do administrador"+(p.ex?" (exemplos orientativos: "+String(p.ex).slice(0,70)+")":"")};
 var cas=Object.keys(byG.cas);
 if(cas.length===1) return {ca:cas[0], conf:0.90, cands:[],
   via:"regra da planilha: Grupo de Mercadoria "+g+" -> Classe de Avaliacao "+cas[0]+" (unica classe prevista para este grupo)",
   lista:"tabela de regra do administrador"};
 if(cas.length>1) return {ca:"", conf:0, cands:cas, via:""};
 return null;
}
/* ---------- leitura de XLSX sem dependencia externa (ZIP + XML) ---------- */
function zipEntries(u8){
 var i, eocd=-1;
 for(i=u8.length-22;i>=0 && i>u8.length-66000;i--){ if(u8[i]===0x50&&u8[i+1]===0x4b&&u8[i+2]===0x05&&u8[i+3]===0x06){ eocd=i; break } }
 if(eocd<0) throw new Error("arquivo nao parece ser uma planilha xlsx");
 function u16(p){ return u8[p]|(u8[p+1]<<8) }
 function u32(p){ return (u8[p]|(u8[p+1]<<8)|(u8[p+2]<<16)|(u8[p+3]*16777216)) }
 var n=u16(eocd+10), off=u32(eocd+16), out={}, p=off, k;
 for(k=0;k<n;k++){
  if(!(u8[p]===0x50&&u8[p+1]===0x4b&&u8[p+2]===0x01&&u8[p+3]===0x02)) break;
  var met=u16(p+10), nl=u16(p+28), el=u16(p+30), cl=u16(p+32), lho=u32(p+42);
  var nome=new TextDecoder("utf-8").decode(u8.subarray(p+46,p+46+nl));
  var lnl=u16(lho+26), lel=u16(lho+28), ds=lho+30+lnl+lel;
  out[nome]={met:met, ds:ds};
  p=p+46+nl+el+cl;
 }
 return out;
}
function zipText(u8,ent){
 if(!ent) return "";
 if(ent.met===0) return new TextDecoder("utf-8").decode(u8.subarray(ent.ds));
 return new TextDecoder("utf-8").decode(zInflate_(u8,ent.ds));
}
function xmlUn(s){ return String(s||"").replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&#(\d+);/g,function(m,d){ return String.fromCharCode(+d) }).replace(/&amp;/g,"&") }
function xlsxMatrix(u8){
 var ents=zipEntries(u8), shared=[], i;
 var ss=zipText(u8, ents["xl/sharedStrings.xml"]);
 if(ss){ var sis=ss.match(/<si[\s>][\s\S]*?<\/si>|<si\/>/g)||[];
  for(i=0;i<sis.length;i++){ var ts=sis[i].match(/<t[^>]*>([\s\S]*?)<\/t>/g)||[]; var txt="";
   for(var j=0;j<ts.length;j++) txt+=xmlUn(ts[j].replace(/<[^>]+>/g,""));
   shared.push(txt) } }
 var nome="";
 for(var k in ents){ if(/^xl\/worksheets\/sheet\d+\.xml$/.test(k)){ if(!nome || k<nome) nome=k } }
 var sh=zipText(u8, ents[nome]); if(!sh) throw new Error("nao encontrei a primeira aba da planilha");
 var rows=sh.match(/<row[\s>][\s\S]*?<\/row>|<row[^>]*\/>/g)||[], out=[];
 for(i=0;i<rows.length;i++){
  var cs=rows[i].match(/<c[\s>][\s\S]*?<\/c>|<c[^>]*\/>/g)||[], lin=[];
  for(var c=0;c<cs.length;c++){
   var cell=cs[c], ref=(cell.match(/r="([A-Z]+)\d+"/)||[])[1]||"", t=(cell.match(/t="([^"]+)"/)||[])[1]||"";
   var col=0; for(var z=0;z<ref.length;z++) col=col*26+(ref.charCodeAt(z)-64);
   col=col>0?col-1:lin.length;
   var val="";
   if(t==="inlineStr"){ var its=cell.match(/<t[^>]*>([\s\S]*?)<\/t>/g)||[]; for(var q=0;q<its.length;q++) val+=xmlUn(its[q].replace(/<[^>]+>/g,"")) }
   else { var vm=cell.match(/<v[^>]*>([\s\S]*?)<\/v>/); if(vm){ val=xmlUn(vm[1]); if(t==="s") val=shared[+val]||"" } }
   while(lin.length<col) lin.push("");
   lin[col]=String(val).trim();
  }
  out.push(lin);
 }
 return out;
}
/* matriz (xlsx ou csv) -> linhas de regra, achando as colunas pelo cabecalho */
function carFromMatrix(mat){
 var hi=-1, cGm=-1, cGc=-1, cCa=-1, cEx=-1, i, j;
 for(i=0;i<Math.min(mat.length,15) && hi<0;i++){
  var lin=mat[i]||[];
  for(j=0;j<lin.length;j++){
   var h=String(lin[j]||"").toUpperCase();
   if(/MERCADORIA/.test(h)) cGm=j;
   else if(/COMPRADOR/.test(h)) cGc=j;
   else if(/AVALIA/.test(h)) cCa=j;
   else if(/EXEMPLO/.test(h)) cEx=j;
  }
  if(cGm>=0 && cCa>=0) hi=i; else { cGm=-1;cGc=-1;cCa=-1;cEx=-1 }
 }
 if(hi<0) throw new Error("nao encontrei as colunas de Grupo de Mercadoria e Classe de Avaliacao no cabecalho");
 function cod(s){ var m=String(s||"").match(/^\s*([A-Z0-9][A-Z0-9._-]*)/i); return m? m[1].toUpperCase() : "" }
 var rows=[], vis={};
 for(i=hi+1;i<mat.length;i++){
  var l=mat[i]||[];
  var gm=cod(l[cGm]), gc=cGc>=0? cod(l[cGc]) : "", ca=cCa>=0? cod(l[cCa]) : "", ex=cEx>=0? String(l[cEx]||"").trim() : "";
  if(!gm) continue;
  var k=carKey(gm,gc);
  if(vis[k]!==undefined){ rows[vis[k]]={gm:gm,gc:gc,ca:ca||rows[vis[k]].ca,ex:ex||rows[vis[k]].ex} }
  else { vis[k]=rows.length; rows.push({gm:gm,gc:gc,ca:ca,ex:ex}) }
 }
 if(!rows.length) throw new Error("a planilha nao trouxe nenhuma linha de regra utilizavel");
 return rows;
}
function carImportFile(file){
 if(!file) return;
 var fr=new FileReader();
 fr.onload=function(){
  try{
   var buf=fr.result, u8=new Uint8Array(buf), rows;
   if(u8.length>3 && u8[0]===0x50 && u8[1]===0x4b) rows=carFromMatrix(xlsxMatrix(u8));
   else { var txt=new TextDecoder("utf-8").decode(u8); rows=carFromMatrix(csvParse(txt)) }
   CAR.rows=rows; CAR.src=(file.name||"planilha")+" importada pelo administrador";
   carApply(rows.length+" linhas de regra importadas de "+(file.name||"planilha"));
  }catch(e){ try{ toast("nao foi possivel ler a planilha de regra: "+(e&&e.message?e.message:"formato nao reconhecido")) }catch(x){} }
 };
 fr.onerror=function(){ try{ toast("nao foi possivel ler o arquivo") }catch(e){} };
 fr.readAsArrayBuffer(file);
}
function carResumo(){
 if(!CAR.rows.length) return "nenhuma regra carregada: a Classe de Avaliacao continua saindo apenas pela evidencia da base, como sugestao. Importe a planilha de regra aqui para o motor derivar a classe pelo Grupo de Mercadoria.";
 var gms=Object.keys(CAR.byGm), comCa=0, k;
 for(k=0;k<gms.length;k++) if(Object.keys(CAR.byGm[gms[k]].cas).length) comCa++;
 var amb=0;
 for(k=0;k<gms.length;k++) if(Object.keys(CAR.byGm[gms[k]].cas).length>1) amb++;
 return CAR.rows.length+" linhas de regra | "+gms.length+" Grupos de Mercadoria cobertos | "+comCa+" com Classe de Avaliacao definida | "+amb+" com mais de uma classe possivel (campo fica em aberto) | origem: "+CAR.src+(CAR.dt?" | salva em "+CAR.dt.slice(0,16).replace("T"," "):"");
}
function renderCaRegra(){
 var tb=null; try{ tb=document.querySelector("#tbCarRegra tbody") }catch(e){}
 if(!tb) return;
 var h="", i;
 var ord=CAR.rows.slice().sort(function(a,b){ return String(a.gm).localeCompare(String(b.gm))||String(a.gc).localeCompare(String(b.gc)) });
 for(i=0;i<ord.length;i++){
  var r=ord[i];
  h+="<tr><td><span class=\"cd\">"+esc(r.gm)+"</span> <span class=\"cx\">"+esc(GM_OFF[r.gm]||"")+"</span></td>"+
     "<td>"+(r.gc?"<span class=\"cd\">"+esc(r.gc)+"</span> <span class=\"cx\">"+esc(GC_OFF[r.gc]||"")+"</span>":"<span class=\"cx\">qualquer</span>")+"</td>"+
     "<td>"+(r.ca?"<span class=\"cd\">"+esc(r.ca)+"</span>":"<span class=\"cx\">sem classe definida</span>")+"</td>"+
     "<td><span class=\"cx\">"+esc(String(r.ex||"").slice(0,90))+"</span></td>"+
     "<td><button class=\"btn gho carDel\" data-gm=\""+esc(r.gm)+"\" data-gc=\""+esc(r.gc||"")+"\">remover</button></td></tr>";
 }
 tb.innerHTML=h||"<tr><td colspan=\"5\"><span class=\"cx\">tabela de regra vazia</span></td></tr>";
 try{ var inf=document.getElementById("carInfo"); if(inf){ var _f=[]; try{ _f=carGmFora() }catch(e){} inf.innerHTML=esc(carResumo())+(_f.length? (" <b>| atencao:</b> "+_f.length+" grupo(s) fora do catalogo 663: "+esc(_f.join(", "))+" - essas linhas nao classificam nada") : "") } }catch(e){}
 try{
  var bs=document.querySelectorAll("#tbCarRegra .carDel");
  for(i=0;i<bs.length;i++) bs[i].onclick=function(){ carDel(this.getAttribute("data-gm"), this.getAttribute("data-gc")) };
 }catch(e){}
}
/* ---------- duplicidade por Descricao Curta (campo que vai para o SAP) ---------- */
var CUR_IX={map:null,src:-1};
function curNz(s){ try{ return norm(String(s||"")).replace(/[^A-Z0-9 ]/g," ").replace(/\s+/g," ").trim() }catch(e){ return String(s||"").toUpperCase().trim() } }
function curEnsure(){
 var rows=[]; try{ rows=S.base||[] }catch(e){ rows=[] }
 if(CUR_IX.map && CUR_IX.src===rows.length) return CUR_IX;
 var m={};
 for(var i=0;i<rows.length;i++){
  var r=rows[i]||{}, k=curNz(r.c||"");
  if(!k) continue;
  if(!m[k]) m[k]=[];
  if(m[k].length<4) m[k].push({cod:r.cod||"",c:r.c||"",d:r.d||"",gm:r.gm||"",gc:r.gc||"",ca:r.ca||""});
 }
 CUR_IX={map:m,src:rows.length};
 return CUR_IX;
}
function curDupFind(curta){
 var k=curNz(curta); if(!k) return null;
 var ix=curEnsure(); var l=ix.map? ix.map[k] : null;
 return (l && l.length)? l : null;
}
function dupCurtaBanner(r){
 var l=r && r.dupCurta; if(!l || !l.length) return "";
 var h="<div class=\"note ri\" style=\"margin-top:12px\"><div class=\"imoHd\"><b>Material ja existe na base</b><span class=\"tag ri\">Descricao Curta identica</span></div>";
 h+="<div class=\"cx\" style=\"margin-top:4px\">A Descricao Curta informada e a que vai para o cadastro no SAP e ja consta na base "+(l.length>1?("em "+l.length+" cadastros"):"em um cadastro")+". Confira antes de abrir um novo material - este aviso nao bloqueia o fluxo.</div>";
 for(var i=0;i<l.length;i++){
  var x=l[i];
  h+="<div class=\"cx\" style=\"margin-top:6px\"><span class=\"cd\">"+esc(x.cod||"sem codigo")+"</span> "+esc(x.c||x.d||"")+
     " &middot; Grupo de Mercadoria: "+esc(x.gm||"-")+" &middot; Grupo de Compradores: "+esc(x.gc||"-")+
     " &middot; Classe de Avaliacao: "+esc(x.ca||"nao informada")+"</div>";
 }
 h+="</div>";
 return h;
}
/* ---------- lista de materiais semelhantes (apoio a decisao de reaproveitar cadastro) ----------
   filtro de falso positivo: semelhanca superficial com categoria diferente da sugerida
   só entra na lista quando a semelhanca e alta o suficiente, e sai marcada. */
function simList(r){
 var l=(r&&r.top3)||[], out=[], i;
 for(i=0;i<l.length;i++){
  var x=l[i], s=x.sim||0;
  if(s<55) continue;
  var mesmoGm=!!(x.gm && r.gm && String(x.gm)===String(r.gm));
  var mesmaFam=!!(x.gm && r.gm && String(x.gm).slice(0,3)===String(r.gm).slice(0,3));
  if(!mesmoGm && !mesmaFam && s<75) continue;
  out.push({cod:x.cod||"",c:x.c||x.d||"",gm:x.gm||"",gc:x.gc||"",ca:x.ca||"",sim:s,mesmoGm:mesmoGm,mesmaFam:mesmaFam});
 }
 out.sort(function(a,b){ return (b.sim+(b.mesmoGm?8:0))-(a.sim+(a.mesmoGm?8:0)) });
 return out;
}
function simHtml(r){
 var l=simList(r); if(!l.length) return "";
 var h="<div class=\"note\" style=\"margin-top:12px\"><b>Materiais semelhantes na base</b>";
 h+="<div class=\"cx\" style=\"margin-top:4px\">Indice de semelhanca com o que ja esta cadastrado, para voce decidir se reaproveita um cadastro existente.</div>";
 h+="<div class=\"tw\" style=\"margin-top:8px\"><table><thead><tr><th>Semelhanca</th><th>Descricao Curta</th><th>Grupo de Mercadoria</th><th>Grupo de Compradores</th><th>Classe de Avaliacao</th></tr></thead><tbody>";
 for(var i=0;i<l.length;i++){
  var x=l[i];
  h+="<tr><td><b>"+x.sim+"%</b>"+(x.mesmoGm?" <span class=\"tag ok\">mesma categoria</span>":(x.mesmaFam?" <span class=\"tag wa\">categoria proxima</span>":" <span class=\"tag wa\">categoria diferente</span>"))+"</td>"+
     "<td>"+esc(x.c)+(x.cod?" <span class=\"cx\">("+esc(x.cod)+")</span>":"")+"</td>"+
     "<td><span class=\"cd\">"+esc(x.gm||"-")+"</span> <span class=\"cx\">"+esc(GM_OFF[x.gm]||"")+"</span></td>"+
     "<td><span class=\"cd\">"+esc(x.gc||"-")+"</span> <span class=\"cx\">"+esc(GC_OFF[x.gc]||"")+"</span></td>"+
     "<td>"+(x.ca?"<span class=\"cd\">"+esc(x.ca)+"</span>":"<span class=\"cx\">nao informada</span>")+"</td></tr>";
 }
 h+="</tbody></table></div></div>";
 return h;
}

/* ================= INIT ================= */
(async function(){
 loadCatSaved(); try{ carBoot() }catch(e){} try{ renderCaRegra() }catch(e){}
 renderRegras(); renderCat(); try{ renderCatAdm() }catch(e){} try{ renderEqAp() }catch(e){}
 try{
  const b=await kvGet("base"), st=await kvGet("baseStamp");
  if(b&&b.length&&st===EMB_STAMP){ S.base=b; BASE_ORIGEM="base salva neste navegador"; }
  else if(b&&b.length){ try{ await kvSet("base",[]) }catch(e){}
   setTimeout(function(){ try{ toast("base local de uma versao anterior foi descartada: valendo a base embutida neste arquivo, com classe de avaliacao") }catch(e){} },1400); }
  const f=await kvGet("fb"); if(f) S.fb=f;
  const l=await kvGet("log"); if(l) S.log=l;
  const p=await kvGet("pwd"); if(p) S.pwd=p;
 }catch(e){}
 buildIndex(); setBaseState(); renderLog(); renderOut();
 if(!S.base.length && !(typeof EMB_PAY!=="undefined" && EMB_PAY.ind && EMB_PAY.ind.gz)) toast("pronto para uso; carregue a base em Base e administracao para ativar a duplicidade");
})();


/* ================================================================
   V7 - CAMADA DE MOVIMENTO, FEEDBACK E CARREGAMENTO PROGRESSIVO
   Nao remove nem substitui nenhuma funcao da v6: envolve as
   existentes para acrescentar skeleton, progresso, animacao e
   feedback visual. Respeita prefers-reduced-motion.
   ================================================================ */
(function(){
 var q=function(s,r){return (r||document).querySelector(s)};
 var qa=function(s,r){return [].slice.call((r||document).querySelectorAll(s))};
 var RM=window.matchMedia?window.matchMedia("(prefers-reduced-motion: reduce)"):{matches:false};
 var raf2=function(f){ if(RM.matches) return f(); requestAnimationFrame(function(){requestAnimationFrame(f)}) };

 /* ---------- 1. BARRA DE PROGRESSO DE SCROLL / NAV FIXA / TOPO ---------- */
 var prog=q("#v7prog>i"), navEl=q("nav"), topBtn=q("#v7top");
 function onScroll(){
  var doc=document.documentElement, h=doc.scrollHeight-doc.clientHeight;
  var p=h>0?Math.min(1,doc.scrollTop/h):0;
  if(prog) prog.style.width=(p*100).toFixed(2)+"%";
  if(navEl) navEl.classList.toggle("v7stuck", doc.scrollTop>6);
  if(topBtn) topBtn.classList.toggle("on", doc.scrollTop>420);
 }
 addEventListener("scroll",onScroll,{passive:true}); onScroll();
 if(topBtn) topBtn.onclick=function(){ scrollTo({top:0,behavior:RM.matches?"auto":"smooth"}) };

 /* ---------- 2. TOASTS EM PILHA (nao bloqueiam clique, nunca cobrem o modal) ---------- */
 var origToast=window.toast;
 function tkind(s){
  var l=s.toLowerCase();
  return /falha|erro|nao reconhec|invalid|nenhum codigo|incorret|senha incorreta/.test(l) ? "ri"
   : /revisar|alerta|atencao|descontinuad|nao carregada|cole ao menos|digite o/.test(l) ? "wa"
   : /pronto|salv|limp|carregad|exportad|copiad|atualizad|restaurad|concluid|reconstru/.test(l) ? "ok" : "in";
 }
 window.toast=function(m,kind){
  var host=q("#toast");
  if(!host) return origToast&&origToast(m);
  var s=String(m==null?"":m), k=kind||tkind(s);
  var ic={ok:"\u2713",wa:"\u26A0",ri:"\u2715",in:"\u2139"}[k]||"\u2139";
  var el=document.createElement("div");
  el.className="v8toast "+k;
  el.innerHTML='<span class="v7ti"></span><span class="v7tm"></span><span class="v7tp"></span>';
  q(".v7ti",el).textContent=ic;
  q(".v7tm",el).textContent=s;
  host.appendChild(el);
  while(host.children.length>3){ host.removeChild(host.firstChild) }
  if(RM.matches) el.classList.add("on"); else requestAnimationFrame(function(){ el.classList.add("on") });
  var dur=Math.min(7000,2600+s.length*22);
  setTimeout(function(){
   el.classList.remove("on"); el.classList.add("off");
   setTimeout(function(){ if(el.parentNode) el.parentNode.removeChild(el) },RM.matches?0:260);
  },dur);
 };

 /* ---------- 3. MODAIS (foco preso, backdrop com blur, retorno de foco) ---------- */
 var ov=q("#v7ov"), lastFocus=null;
 function ovFocusables(){
  return qa("button,input,select,textarea,[href]",ov).filter(function(el){ return !el.disabled && el.offsetParent!==null });
 }
 function closeOv(){
  if(!ov) return;
  ov.classList.remove("on"); ov.setAttribute("aria-hidden","true");
  document.body.classList.remove("v8lock");
  var d=q(".v7dlg",ov);
  setTimeout(function(){ if(d&&!ov.classList.contains("on")){ d.className="v7dlg"; d.innerHTML="" } },RM.matches?0:240);
  if(lastFocus&&lastFocus.focus){ try{ lastFocus.focus() }catch(e){} }
 }
 function openOv(html,danger){
  if(!ov) return null;
  lastFocus=document.activeElement;
  var d=q(".v7dlg",ov);
  d.className="v7dlg"+(danger?" ri":""); d.innerHTML=html;
  ov.removeAttribute("aria-hidden");
  document.body.classList.add("v8lock");
  ov.classList.add("on");
  setTimeout(function(){ var f=ovFocusables()[0]; if(f) f.focus() },RM.matches?0:200);
  return d;
 }
 if(ov){
  ov.setAttribute("aria-hidden","true");
  ov.addEventListener("click",function(e){ if(e.target===ov) closeOv() });
  ov.addEventListener("keydown",function(e){
   if(e.key!=="Tab") return;
   var f=ovFocusables(); if(!f.length) return;
   var first=f[0], last=f[f.length-1];
   if(e.shiftKey && document.activeElement===first){ e.preventDefault(); last.focus() }
   else if(!e.shiftKey && document.activeElement===last){ e.preventDefault(); first.focus() }
  });
  addEventListener("keydown",function(e){ if(e.key==="Escape"&&ov.classList.contains("on")) closeOv() });
 }
 function v7ask(o){
  return new Promise(function(res){
   var d=openOv('<h3><i>'+(o.danger?"!":"?")+'</i>'+esc(o.title)+'</h3>'+
     '<div class="v7dlgb">'+(o.body||"")+'</div>'+
     '<div class="v7dlgf"><button class="btn gho" data-a="0">Cancelar</button>'+
     '<button class="btn'+(o.danger?" dgr":"")+'" data-a="1">'+esc(o.ok||"Confirmar")+'</button></div>',o.danger);
   if(!d) return res(confirm(o.title));
   qa("[data-a]",d).forEach(function(b){ b.onclick=function(){ closeOv(); res(b.dataset.a==="1") } });
  });
 }
 function v7prompt(o){
  return new Promise(function(res){
   var d=openOv('<h3><i>\u270E</i>'+esc(o.title)+'</h3>'+
     '<div class="v7dlgb"><label for="v7pv">'+esc(o.label||"")+'</label>'+
     '<input type="'+(o.type||"text")+'" id="v7pv" placeholder="'+esc(o.ph||"")+'"></div>'+
     '<div class="v7dlgf"><button class="btn gho" data-a="0">Cancelar</button>'+
     '<button class="btn" data-a="1">Salvar</button></div>');
   if(!d) return res(prompt(o.title));
   var inp=q("#v7pv",d);
   inp.onkeydown=function(e){ if(e.key==="Enter"){ var v=inp.value; closeOv(); res(v) } };
   qa("[data-a]",d).forEach(function(b){ b.onclick=function(){ var v=inp.value; closeOv(); res(b.dataset.a==="1"?v:"") } });
  });
 }
 window.v7ask=v7ask; window.v7prompt=v7prompt;

 /* ---------- 4. ESTADOS DE PROGRESSO NOS BOTOES ---------- */
 function busy(b){ if(!b) return; b._v7t0=Date.now(); b.classList.remove("v7done"); b.classList.add("v7busy"); b.setAttribute("aria-busy","true") }
 function done(b,ok){
  if(!b) return;
  var wait=Math.max(0,380-(Date.now()-(b._v7t0||0)));
  setTimeout(function(){
   b.classList.remove("v7busy"); b.removeAttribute("aria-busy");
   if(ok&&!RM.matches){ b.classList.add("v7done"); setTimeout(function(){ b.classList.remove("v7done") },760) }
  },RM.matches?0:wait);
 }

 /* ---------- 5. SKELETON SCREENS ---------- */
 function skHtml(kind,msg){
  var m=msg?'<div class="v7skmsg"><i></i>'+esc(msg)+'</div>':"";
  if(kind==="cards"){
   return '<div class="v7skwrap">'+m+
    '<div class="v7skcard"><div class="v7sk t" style="width:22%"></div><div class="v7sk h" style="width:56%"></div>'+
    '<div class="v7sk" style="width:100%"></div>'+
    '<div class="v7skfld">'+
     '<div class="v7skcard"><div class="v7sk t" style="width:45%"></div><div class="v7sk h" style="width:62%"></div><div class="v7sk t" style="width:88%"></div></div>'+
     '<div class="v7skcard"><div class="v7sk t" style="width:52%"></div><div class="v7sk h" style="width:48%"></div><div class="v7sk t" style="width:80%"></div></div>'+
     '<div class="v7skcard"><div class="v7sk t" style="width:38%"></div><div class="v7sk h" style="width:66%"></div><div class="v7sk t" style="width:74%"></div></div>'+
     '<div class="v7skcard"><div class="v7sk t" style="width:48%"></div><div class="v7sk h" style="width:42%"></div><div class="v7sk t" style="width:84%"></div></div>'+
    '</div></div></div>';
  }
  if(kind==="table"){
   var r=""; for(var i=0;i<7;i++){ r+='<div class="v7sk" style="width:'+(96-i*5)+'%;margin-top:9px"></div>' }
   return '<div class="v7skwrap">'+m+'<div class="v7skgrid">'+
    '<div class="v7skcard"><div class="v7sk h" style="width:52%"></div><div class="v7sk t" style="width:74%"></div></div>'+
    '<div class="v7skcard"><div class="v7sk h" style="width:44%"></div><div class="v7sk t" style="width:68%"></div></div>'+
    '<div class="v7skcard"><div class="v7sk h" style="width:58%"></div><div class="v7sk t" style="width:70%"></div></div>'+
    '</div><div style="margin-top:14px">'+r+'</div></div>';
  }
  if(kind==="kpi"){
   return '<div class="v7skcard"><div class="v7sk h" style="width:46%"></div><div class="v7sk t" style="width:72%"></div></div>'+
    '<div class="v7skcard"><div class="v7sk h" style="width:38%"></div><div class="v7sk t" style="width:64%"></div></div>'+
    '<div class="v7skcard"><div class="v7sk h" style="width:52%"></div><div class="v7sk t" style="width:70%"></div></div>';
  }
  return '<div class="v7skwrap">'+m+'<div class="v7sk" style="width:40%"></div><div class="v7sk" style="width:100%;margin-top:10px"></div><div class="v7sk" style="width:88%;margin-top:8px"></div></div>';
 }
 function skel(sel,kind,msg){ var b=q(sel); if(b) b.innerHTML=skHtml(kind,msg) }
 window.v7skel=skel;

 /* ---------- 6. CARREGAMENTO PROGRESSIVO DE LINHAS (lazy) ---------- */
 function lazyRows(tb,first,chunk){
  if(!tb) return;
  var body=q("tbody",tb); if(!body) return;
  var rows=qa("tr",body); if(rows.length<=first+6){ rows.slice(0,26).forEach(function(r,i){ if(RM.matches) return; r.classList.add("v7row"); r.style.animationDelay=Math.min(i,18)*26+"ms" }); return }
  var host=tb.closest(".tw")||tb;
  qa(".v7more",host.parentNode).forEach(function(x){x.remove()});
  var shown=first;
  rows.forEach(function(r,i){ if(i>=first) r.classList.add("v7hid") });
  var more=document.createElement("div"); more.className="v7more";
  more.innerHTML='<div class="v7sk" style="width:100%"></div><div class="v7sk" style="width:92%"></div>'+
   '<div class="v7cnt"></div>';
  host.parentNode.insertBefore(more,host.nextSibling);
  var cnt=q(".v7cnt",more);
  function label(){ cnt.textContent=shown.toLocaleString("pt-BR")+" de "+rows.length.toLocaleString("pt-BR")+" linhas carregadas" }
  label();
  function reveal(){
   var end=Math.min(rows.length,shown+chunk);
   for(var i=shown;i<end;i++){ var r=rows[i]; r.classList.remove("v7hid"); if(!RM.matches){ r.classList.add("v7row"); r.style.animationDelay=Math.min(i-shown,16)*22+"ms" } }
   shown=end; label();
   if(shown>=rows.length){ io.disconnect(); more.remove(); return }
   io.unobserve(more); requestAnimationFrame(function(){ try{ io.observe(more) }catch(e){} });
  }
  var io=new IntersectionObserver(function(en){ en.forEach(function(e){ if(e.isIntersecting) reveal() }) },{rootMargin:"260px"});
  io.observe(more);
 }

 /* ---------- 7. CONTAGEM ANIMADA DOS KPIs ---------- */
 function countUp(el){
  if(RM.matches||el.dataset.v7cu) return; el.dataset.v7cu="1";
  var txt=el.textContent.trim(), m=txt.match(/^(\D*?)([\d][\d.,]*)(\D*)$/);
  if(!m) return;
  var raw=m[2], hasDot=raw.indexOf(".")>=0 && raw.indexOf(",")<0;
  var num=parseFloat(raw.replace(/\./g,"").replace(",","."));
  if(!isFinite(num)||num<=0||num>9999999) return;
  var t0=performance.now(), dur=560, dec=(raw.indexOf(",")>=0)?1:0;
  function frame(t){
   var p=Math.min(1,(t-t0)/dur), e=1-Math.pow(1-p,3), v=num*e;
   el.textContent=m[1]+(dec?v.toFixed(1).replace(".",","):(hasDot?Math.round(v).toLocaleString("pt-BR"):String(Math.round(v))))+m[3];
   if(p<1) requestAnimationFrame(frame); else el.textContent=txt;
  }
  requestAnimationFrame(frame);
 }

 /* ---------- 8. POS-RENDER: entrada escalonada, lazy e contagem ---------- */
 function afterRender(name){
  if(name==="renderOut"||name==="all"){
   qa("#out .fld").forEach(function(el,i){ if(!RM.matches) el.style.animationDelay=Math.min(i,10)*45+"ms" });
   qa("#out .kpi b").forEach(countUp);
   qa("#out .bar").forEach(function(b){ b.classList.add("v7done") });
   qa("#out table").forEach(function(t){ lazyRows(t,60,60) });
  }
  if(name==="renderLog"||name==="all"){ lazyRows(q("#logTb"),40,40) }
  if(name==="setBaseState"||name==="renderCat"||name==="all"){ qa("#baseKpis .kpi b,#catKpis .kpi b").forEach(countUp) }
  if(name==="renderRegras"||name==="all"){ qa("#tbGm,#tbGc,#tbTipos").forEach(function(t){ lazyRows(t,40,40) }) }
  syncSteps();
 }
 ["renderOut","renderLog","renderCat","setBaseState","renderRegras"].forEach(function(n){
  var o=window[n]; if(typeof o!=="function") return;
  window[n]=function(){ var r=o.apply(this,arguments); try{ afterRender(n) }catch(e){} return r };
 });

 /* ---------- 9. RIPPLE E FEEDBACK DE TOQUE ---------- */
 document.addEventListener("pointerdown",function(e){
  var b=e.target.closest?e.target.closest(".btn"):null;
  if(!b||b.disabled||RM.matches) return;
  var r=b.getBoundingClientRect(), s=document.createElement("span");
  s.className="v7rip"; s.style.left=(e.clientX-r.left)+"px"; s.style.top=(e.clientY-r.top)+"px";
  b.appendChild(s); setTimeout(function(){ s.remove() },640);
 },true);

 /* ---------- 10. ENVOLVER ACOES: skeleton + progresso ---------- */
 function wrapBtn(id,opt){
  opt=opt||{}; var b=q("#"+id); if(!b) return; var orig=b.onclick; if(typeof orig!=="function") return;
  b.onclick=function(ev){
   var self=this;
   busy(b);
   if(opt.skel) skel(opt.skel[0],opt.skel[1],opt.skel[2]);
   raf2(function(){
    var r;
    try{ r=orig.call(self,ev) }catch(err){ done(b,false); throw err }
    Promise.resolve(r).then(function(){ done(b,opt.check!==false) },function(){ done(b,false) });
   });
   return false;
  };
 }
 function wrapInput(id,btnId,skelArgs){
  var el=q("#"+id); if(!el) return; var orig=el.onchange; if(typeof orig!=="function") return;
  el.onchange=function(ev){
   var self=this, b=btnId?q("#"+btnId):null;
   if(b) busy(b);
   if(skelArgs) skel(skelArgs[0],skelArgs[1],skelArgs[2]);
   var r;
   try{ r=orig.call(self,ev) }catch(err){ done(b,false); throw err }
   Promise.resolve(r).then(function(){ done(b,true) },function(){ done(b,false) });
   return r;
  };
 }
 window.v7wrapBtn=wrapBtn; window.v7wrapInput=wrapInput;
 /* analise unitaria: skeleton do cartao antes de classificar */
 var origUnico=window.runUnico;
 if(typeof origUnico==="function"){
  window.runUnico=function(){
   var d=(q("#u1")||{}).value;
   if(!d||!String(d).trim()) return origUnico.apply(this,arguments);
   var b=q("#bUnico"); busy(b);
   skel("#out","cards","analisando o material e comparando com a base...");
   var self=this, args=arguments;
   raf2(function(){
    try{ origUnico.apply(self,args) } finally { done(b,true) }
   });
  };
  var bu=q("#bUnico"); if(bu) bu.onclick=function(){ window.runUnico() };
  var u1=q("#u1"); if(u1) u1.onkeydown=function(e){ if(e.key==="Enter") window.runUnico() };
 }
 /* lote: skeleton de tabela + progresso */
 var origList=window.analyzeList;
 if(typeof origList==="function"){
  window.analyzeList=function(items){
   skel("#out","table","analisando "+(items&&items.length?items.length.toLocaleString("pt-BR"):"")+" materiais em lote...");
   return origList.apply(this,arguments);
  };
 }
 ["bLote","bUp","bExemplo"].forEach(function(id){ wrapBtn(id) });
 ["bXlsModelo","bCsvModelo","bCopy","bXlsFull","bAudit","bExpBase","bExpFb","bExpLog","bAdm","bRelearn"].forEach(function(id){ wrapBtn(id) });
 wrapBtn("bRelearn",{skel:["#baseKpis","kpi","reconstruindo o aprendizado da base..."]});
 wrapInput("baseFile","bBaseFile",["#baseKpis","kpi","lendo a planilha da base..."]);
 wrapInput("catFile","bCatFile",["#catKpis","kpi","lendo os catalogos oficiais..."]);

 /* ---------- 11. CONFIRMACOES E SENHA EM MODAL ---------- */
 function wrapConfirm(id,o){
  var b=q("#"+id); if(!b) return; var orig=b.onclick; if(typeof orig!=="function") return;
  b.onclick=function(ev){
   var self=this;
   v7ask(o).then(function(yes){
    if(!yes) return;
    var nc=window.confirm; window.confirm=function(){ return true };
    busy(b);
    var r;
    try{ r=orig.call(self,ev) } finally { setTimeout(function(){ window.confirm=nc },0) }
    Promise.resolve(r).then(function(){ done(b,true) },function(){ done(b,false) });
   });
   return false;
  };
 }
 wrapConfirm("bClrBase",{title:"Limpar a base carregada?",danger:true,ok:"Limpar base",
  body:"A base de materiais sai da memoria do navegador. <b>Suas correcoes aprendidas e o log de decisoes sao mantidos.</b>"});
 wrapConfirm("bClrLog",{title:"Limpar o log de decisoes?",danger:true,ok:"Limpar log",
  body:"O historico de analises e correcoes registradas sera apagado. A base e os catalogos permanecem."});
 wrapConfirm("bCatReset",{title:"Voltar aos catalogos embutidos?",ok:"Restaurar",
  body:"Os catalogos 650 / 663 / 627 carregados por planilha serao substituidos pelas listas embutidas no arquivo."});
 (function(){
  var b=q("#bChgPwd"); if(!b) return; var orig=b.onclick; if(typeof orig!=="function") return;
  b.onclick=function(){
   v7prompt({title:"Trocar senha do administrador",label:"Nova senha",type:"password",ph:"digite a nova senha"}).then(function(v){
    if(!v) return;
    var op=window.prompt; window.prompt=function(){ return v };
    busy(b);
    var r;
    try{ r=orig.call(b) } finally { setTimeout(function(){ window.prompt=op },0) }
    Promise.resolve(r).then(function(){ done(b,true) },function(){ done(b,false) });
   });
   return false;
  };
 })();

 /* ---------- 12. EDICAO DE CAMPO: destaque apos salvar ---------- */
 var origApply=window.applyEdit;
 if(typeof origApply==="function"){
  window.applyEdit=function(i,f,v){
   var r=origApply.apply(this,arguments);
   setTimeout(function(){
    var el=q('#out .fld[data-i="'+i+'"][data-f="'+f+'"]');
    if(el){ el.classList.add("v7flash"); el.scrollIntoView({behavior:RM.matches?"auto":"smooth",block:"nearest"});
     setTimeout(function(){ el.classList.remove("v7flash") },1200) }
   },40);
   return r;
  };
 }

 /* ---------- 13. AREA DE UPLOAD: estados de arraste ---------- */
 var dz=q("#drop");
 if(dz){
  ["dragenter","dragover"].forEach(function(ev){ dz.addEventListener(ev,function(e){ e.preventDefault(); dz.classList.add("v7over") }) });
  ["dragleave","dragend","drop"].forEach(function(ev){ dz.addEventListener(ev,function(){ dz.classList.remove("v7over") }) });
  dz.addEventListener("drop",function(){ dz.classList.add("v7ok"); setTimeout(function(){ dz.classList.remove("v7ok") },1400) });
 }
 var fi=q("#file");
 if(fi) fi.addEventListener("change",function(){ if(dz&&fi.files&&fi.files.length){ dz.classList.add("v7ok"); setTimeout(function(){ dz.classList.remove("v7ok") },1400) } });

 /* ---------- 14. ABAS: indicador deslizante e transicao entre telas ---------- */
 var ind=q(".v7navind"), steps=q("#v7steps");
 function moveInd(){
  var b=q("nav button.on"); if(!b||!ind) return;
  ind.style.width=b.offsetWidth+"px";
  ind.style.transform="translateX("+b.offsetLeft+"px)";
  ind.style.opacity="1";
 }
 function navh(){ try{ if(navEl&&document.documentElement.style.setProperty) document.documentElement.style.setProperty("--navh",navEl.offsetHeight+"px") }catch(e){} }
 addEventListener("resize",function(){ navh(); moveInd() }); navh();
 addEventListener("resize",moveInd);
 if(navEl) navEl.addEventListener("scroll",moveInd,{passive:true});
 qa("nav button").forEach(function(b){
  var orig=b.onclick;
  b.onclick=function(ev){
   var tab=b.dataset.t;
   var cur=qa("main>section").filter(function(s){ return !s.classList.contains("hidden") })[0];
   var nxt=q("#t-"+tab);
   qa("main>section").forEach(function(s){ s.classList.remove("v7in","v7out") });
   if(steps) steps.classList.toggle("v7off", tab!=="analise");
   function swap(){
    if(typeof orig==="function"){ orig.call(b,ev) }
    else{
     qa("nav button").forEach(function(x){ x.classList.remove("on") }); b.classList.add("on");
     qa("main>section").forEach(function(s){ s.classList.add("hidden") });
     if(nxt) nxt.classList.remove("hidden");
    }
    moveInd();
    if(nxt&&!RM.matches){ void nxt.offsetWidth; nxt.classList.add("v7in") }
    if(document.documentElement.scrollTop>120) scrollTo({top:0,behavior:RM.matches?"auto":"smooth"});
    afterRender("all");
   }
   if(cur&&nxt&&cur!==nxt&&!RM.matches){ cur.classList.add("v7out"); setTimeout(swap,140) } else swap();
   return false;
  };
 });
 setTimeout(moveInd,60); setTimeout(moveInd,400);

 /* ---------- 15. MINI STEPPER DO FLUXO ---------- */
 var stepMap=[["#stepEntry","1"],["#stepRes","2"],["#stepAct","3"]];
 function syncSteps(){
  if(!steps) return;
  var has = (typeof S!=="undefined" && S.results && S.results.length)?true:false;
  var c2=q('.st[data-s="2"]',steps), c3=q('.st[data-s="3"]',steps);
  if(c2) c2.classList.toggle("done",has);
  if(c3) c3.classList.toggle("done",has&&!!(q("#tsv")&&q("#tsv").value));
 }
 if(steps){
  qa(".st",steps).forEach(function(c){
   c.onclick=function(){
    var t=q(c.dataset.tg); if(!t) return;
    var y=t.getBoundingClientRect().top+scrollY-70;
    scrollTo({top:y,behavior:RM.matches?"auto":"smooth"});
   };
  });
  var targets=stepMap.map(function(p){ return q(p[0]) }).filter(Boolean);
  if(targets.length&&"IntersectionObserver" in window){
   var io2=new IntersectionObserver(function(en){
    en.forEach(function(e){
     if(!e.isIntersecting) return;
     var id="#"+e.target.id, n=(stepMap.filter(function(p){ return p[0]===id })[0]||[])[1];
     qa(".st",steps).forEach(function(c){ c.classList.toggle("on",c.dataset.s===n) });
    });
   },{rootMargin:"-45% 0px -45% 0px"});
   targets.forEach(function(t){ io2.observe(t) });
  }
 }

 /* ---------- 16. ATALHOS ---------- */
 addEventListener("keydown",function(e){
  if(e.key==="/"&&!/input|textarea|select/i.test((document.activeElement||{}).tagName||"")){
   var i=q("#u1"); if(i){ e.preventDefault(); var b=q('nav button[data-t="analise"]'); if(b&&!b.classList.contains("on")) b.onclick(); i.focus(); i.select() }
  }
 });

 /* ---------- 17. PRIMEIRA PINTURA ---------- */
 var first=q("#t-analise"); if(first&&!RM.matches) first.classList.add("v7in");
 afterRender("all");
})();


/* =====================================================================
   v7.5 - CONFERENCIA DE IMOBILIZADO POR DESCRICAO, CLASSE DE AVALIACAO
   POR DESCRICAO + GM + GC, COLUNA 8 NA EXPORTACAO E BASES EMBUTIDAS
   EM PAYLOAD COMPRIMIDO.
   Bloco 100% aditivo: nenhuma funcao, id ou fluxo das versoes anteriores
   foi removido. caBuild e caPick sao redeclarados aqui (a ultima
   declaracao vale) mantendo o mesmo contrato de entrada e de saida.
   ===================================================================== */

/* ---------- 1. bases embutidas em payload comprimido ---------- */
var EMB_STAMP = "302i-25009+imo-1390@2026-09-23-v89";
var BASE_ORIGEM = "";
var EMB_PAY = {
 ind:{ gz:EMB_IND_GZ, fmt:"V1", n:25009, nome:"Base de Materiais Indiretos - 302" },
 imo:{ gz:EMB_IMO_GZ, fmt:"V2", n:1390, nome:"Base de Materiais Imobilizados" }
};
function embB64(s){
 const bin=atob(String(s||"").replace(/[^A-Za-z0-9+/=]/g,""));
 const u=new Uint8Array(bin.length);
 for(let i=0;i<bin.length;i++) u[i]=bin.charCodeAt(i);
 return u;
}
function embTxt(b64){
 const u=embB64(b64);
 if(u.length>2 && u[0]===0x1f && u[1]===0x8b){
  if(typeof DecompressionStream!=="undefined"){
   const st=new Blob([u]).stream().pipeThrough(new DecompressionStream("gzip"));
   return new Response(st).text().catch(function(){ return GZS(b64) });
  }
  return Promise.resolve(GZS(b64));
 }
 return Promise.resolve(new TextDecoder("utf-8").decode(u));
}
function embPayLoad(k){
 const p=EMB_PAY[k];
 if(!p || !p.gz) return Promise.resolve(false);
 return embTxt(p.gz).then(function(txt){
  const arr = (p.fmt==="V1" || p.fmt==="V2" || txt.slice(0,3)==="V1\n" || txt.slice(0,3)==="V2\n") ? embDecodePay(txt) : JSON.parse(txt);
  if(Array.isArray(arr) && arr.length){ BASE_EMB[k].rows=arr; p.n=arr.length; return true }
  return false;
 }).catch(function(){ return false });
}
function embPayBoot(){
 return Promise.all([embPayLoad("ind"), embPayLoad("imo")]).then(function(r){
  if(!r[0] && !r[1]) return false;
  try{ IMO_IX.built=false; imoEnsure() }catch(e){}
  try{ if((!S.base || !S.base.length) && baseEmbHas("ind")) baseEmbApply("ind",true) }catch(e){}
  try{ if(typeof renderBasesEmb==="function") renderBasesEmb() }catch(e){}
  try{ if(typeof setBaseState==="function") setBaseState() }catch(e){}
  try{ if(typeof renderOut==="function" && S.results && S.results.length) renderOut() }catch(e){}
  return true;
 });
}
function embPayInit(){ try{ embPayBoot() }catch(e){} }
if(typeof document!=="undefined"){
 if(document.readyState==="loading") document.addEventListener("DOMContentLoaded", embPayInit);
 else setTimeout(embPayInit, 90);
}

/* ---------- 2. camada de conferencia de imobilizado ----------
   A base de imobilizados NAO sugere codigo: os GM 80001-80006 e o tipo
   ZATI nao constam nos catalogos 663/650 e por isso nunca entram como
   resultado. Ela entra apenas como cruzamento por descricao, em cinco
   niveis de evidencia, com tres estados de veredito e a regra da
   IN-CPR-01.001 / RES Risco 194 por cima. Sugestao, nunca bloqueio. */
var IMO_IX={exato:{},curta:{},submod:{},ncm:{},sub:{},nRows:0,built:false,src:-1};
/* ---------- 2b. motor de similaridade nas tres descricoes ----------
   Compara a descricao do material, a descricao curta e a descricao longa
   do item que esta sendo cadastrado contra as MESMAS tres descricoes de
   cada linha da base de imobilizados. Cada token pesa pelo inverso da sua
   frequencia na base (termo raro decide, termo generico quase nao conta),
   o score e o cosseno desses pesos e vale o melhor dos tres campos, com
   bonus quando mais de uma descricao concorda. Serve de apoio a decisao
   imobilizado x consumo: mostra os itens mais parecidos e o quanto. */
var IMO_SIM={built:false,src:-1,n:0,idf:{},inv:{},rows:[]};
function imoTkU(s){ var a=imoTk(s), o={}, r=[], i; for(i=0;i<a.length;i++){ if(!o[a[i]]){ o[a[i]]=1; r.push(a[i]) } } return r }
function imoNormV(tk,idf){ var s=0,i,w; for(i=0;i<tk.length;i++){ w=idf[tk[i]]||1.2; s+=w*w } return Math.sqrt(s)||1 }
function imoCos(inTk,inN,rowTk,idf){
 if(!rowTk.length||!inTk.length) return {s:0,mx:0};
 var set={},i,w,acc=0,mx=0;
 for(i=0;i<rowTk.length;i++) set[rowTk[i]]=1;
 for(i=0;i<inTk.length;i++){ if(set[inTk[i]]){ w=idf[inTk[i]]||1.2; acc+=w*w; if(w>mx) mx=w } }
 if(!acc) return {s:0,mx:0};
 return {s: acc/(inN*imoNormV(rowTk,idf)), mx:mx};
}
function imoSimBuild(rows){
 var i,j,t,df={},list=[];
 for(i=0;i<rows.length;i++){
  var r=rows[i]||{};
  var sd=imoTkU(r.d||r.desc||""), sc=imoTkU(r.c||""), sl=imoTkU(r.l||"");
  var all={},k;
  for(k=0;k<sd.length;k++) all[sd[k]]=1;
  for(k=0;k<sc.length;k++) all[sc[k]]=1;
  for(k=0;k<sl.length;k++) all[sl[k]]=1;
  var au=Object.keys(all);
  for(k=0;k<au.length;k++) df[au[k]]=(df[au[k]]||0)+1;
  list.push({d:r.d||r.desc||"",c:r.c||"",l:r.l||"",cod:r.cod||"",gm:r.gm||"",ncm:r.ncm||"",tk:[sd,sc,sl],all:au});
 }
 var N=Math.max(1,rows.length), idf={};
 for(t in df) idf[t]=Math.log(1+N/df[t]);
 var inv={};
 for(i=0;i<list.length;i++){ var a=list[i].all; for(j=0;j<a.length;j++){ (inv[a[j]]||(inv[a[j]]=[])).push(i) } }
 IMO_SIM={built:true,src:rows.length,n:list.length,idf:idf,inv:inv,rows:list};
 return IMO_SIM;
}
var IMO_CAMPOS=["descricao do material","descricao curta","descricao longa"];
function imoSimTop(dN,cN,lN,k){
 var IX=IMO_SIM; if(!IX.n) return [];
 var inTk=imoTkU([dN,cN,lN].filter(function(x){return !!x}).join(" "));
 if(!inTk.length) return [];
 var idf=IX.idf, inN=imoNormV(inTk,idf), i, j;
 var ord=inTk.slice().sort(function(x,y){ return (idf[y]||0)-(idf[x]||0) });
 var lim=Math.max(80, Math.floor(IX.n*0.35)), cand={}, nc=0;
 for(i=0;i<ord.length && i<16 && nc<3000;i++){
  var l=IX.inv[ord[i]]; if(!l || l.length>lim) continue;
  for(j=0;j<l.length;j++){ if(!cand[l[j]]){ cand[l[j]]=1; nc++ } }
 }
 if(!nc) return [];
 var out=[];
 for(var id in cand){
  var r=IX.rows[id], best=0, sec=0, bmx=0, campo="", sc;
  for(i=0;i<3;i++){
   sc=imoCos(inTk,inN,r.tk[i],idf);
   if(sc.s>best){ sec=best; best=sc.s; bmx=sc.mx; campo=IMO_CAMPOS[i] }
   else if(sc.s>sec) sec=sc.s;
  }
  if(best<=0) continue;
  var sco=best+(sec>=0.6*best?0.03:0);
  if(bmx<1.0) sco=Math.min(sco,0.50);
  out.push({cod:r.cod,d:r.d,c:r.c,gm:r.gm,ncm:r.ncm,campo:campo,mx:bmx,score:Math.max(0,Math.min(0.99,sco)),dupla:(sec>=0.6*best)});
 }
 out.sort(function(x,y){ return y.score-x.score });
 return out.slice(0,k||3);
}

var IMO_R={
 exc:["NOTEBOOK","LAPTOP","MACBOOK","CELULAR","SMARTPHONE","IPHONE","TABLET","IPAD","COLETOR",
      "CROMATOGRAFO","ESPECTROFOTOMETRO","AUTOCLAVE","AUDIOMETRO","MICROSCOPIO","CENTRIFUGA",
      "MUFLA","POLARIMETRO","REFRATOMETRO","CFTV","CATRACA","ALARME","EXTINTOR"],
 excTxt:"IN-CPR-01.001: excecoes que sao imobilizado independentemente do valor (notebooks, celulares, tablets, coletores de dados, equipamentos laboratoriais, de seguranca e definidos como criticos), com validacao formal da Controladoria",
 ctrl:["VENTILADOR","VENTOINHA","CAFETEIRA","LIQUIDIFICADOR","CADEIRA","BANQUETA","TECLADO","MOUSE",
       "HEADSET","HEADFONE","WEBCAM","PENDRIVE","ALICATE","MARTELO","MARRETA","TRENA","SERROTE",
       "LANTERNA","GRAMPEADOR","ESCADA","CHAVE DE FENDA","JOGO DE CHAVES","FERRAMENTA MANUAL","PEN DRIVE"],
 ctrlTxt:"IN-CPR-01.001 / RES Risco 194: bem controlado de baixo valor - vida util acima de 12 meses sem materialidade (ventiladores, cafeteiras, ferramentas manuais, cadeiras e perifericos de baixo valor): exige rastreabilidade e NAO vai para o imobilizado",
 cpc:"CPC 27 / Lei 12.973-2014 (IN-CPR-01.001): imobilizado exige os tres requisitos simultaneos - bem tangivel mantido para uso, vida util superior a um ano e beneficio economico futuro mensuravel",
 cons:"IN-CPR-01.001: material de consumo ou peca de reposicao - perde identidade fisica no uso ou tem vida util de ate dois anos: nao e imobilizado",
 sem:"CPC 27 / IN-CPR-01.001: nenhum dos tres requisitos de imobilizado foi evidenciado pela descricao",
 pec:["PARAFUSO","PORCA","ARRUELA","ROLAMENTO","RETENTOR","CORREIA","FILTRO","MANGUEIRA","JUNTA",
      "VEDACAO","BUCHA","ENGRENAGEM","PINO","MOLA","TERMINAL","ABRACADEIRA","LAMPADA","FUSIVEL",
      "CARTUCHO","ELETRODO","GRAXA","OLEO","TINTA","ANEL","KIT","REPARO","ELEMENTO"]
};
function imoNz(s){ try{ return norm(String(s||"")).replace(/\s+/g," ").trim() }catch(e){ return String(s||"").toUpperCase().trim() } }
function imoTk(s){ try{ return toks(String(s||""))||[] }catch(e){ return [] } }
function imoRowsSrc(){
 try{ if(typeof BASE_SESS!=="undefined" && BASE_SESS && BASE_SESS.imo && BASE_SESS.imo.length) return BASE_SESS.imo }catch(e){}
 try{ if(BASE_ATIVA==="imo" && S.base && S.base.length) return S.base }catch(e){}
 try{ if(BASE_EMB && BASE_EMB.imo && BASE_EMB.imo.rows && BASE_EMB.imo.rows.length) return BASE_EMB.imo.rows }catch(e){}
 return [];
}
function imoBuild(){
 const rows=imoRowsSrc();
 IMO_IX={exato:{},curta:{},submod:{},ncm:{},sub:{},nRows:rows.length,built:true,src:rows.length};
 for(let i=0;i<rows.length;i++){
  const r=rows[i]||{};
  const d=imoNz(r.d||r.desc||""), c=imoNz(r.c||""), l=imoNz(r.l||"");
  if(d) IMO_IX.exato[d]=(IMO_IX.exato[d]||0)+1;
  if(l && l!==d) IMO_IX.exato[l]=(IMO_IX.exato[l]||0)+1;
  if(c) IMO_IX.curta[c]=(IMO_IX.curta[c]||0)+1;
  const t=imoTk(d||c||l);
  if(t.length){
   IMO_IX.sub[t[0]]=(IMO_IX.sub[t[0]]||0)+1;
   if(t[1]) IMO_IX.submod[t[0]+"|"+t[1]]=(IMO_IX.submod[t[0]+"|"+t[1]]||0)+1;
  }
  const nc=String(r.ncm||"").replace(/\D/g,"");
  if(nc.length>=8) IMO_IX.ncm[nc.slice(0,8)]=(IMO_IX.ncm[nc.slice(0,8)]||0)+1;
 }
 return IMO_IX;
}
function imoEnsure(){ const rows=imoRowsSrc(), n=rows.length;
 if(!IMO_IX.built || IMO_IX.src!==n) imoBuild();
 if(!IMO_SIM.built || IMO_SIM.src!==n){ try{ imoSimBuild(rows) }catch(e){ IMO_SIM={built:true,src:n,n:0,idf:{},inv:{},rows:[]} } }
 return IMO_IX }
/* indice de descricoes da base de indiretos: usado apenas para detectar
   a mesma descricao cadastrada nas duas bases (conflito) */
var IND_EX={set:null,src:-1};
function indRowsSrc(){
 try{ if(BASE_ATIVA==="ind" && S.base && S.base.length) return S.base }catch(e){}
 try{ if(typeof BASE_SESS!=="undefined" && BASE_SESS && BASE_SESS.ind && BASE_SESS.ind.length) return BASE_SESS.ind }catch(e){}
 try{ if(BASE_EMB && BASE_EMB.ind && BASE_EMB.ind.rows && BASE_EMB.ind.rows.length) return BASE_EMB.ind.rows }catch(e){}
 return [];
}
function indEnsure(){
 const rows=indRowsSrc();
 if(IND_EX.set && IND_EX.src===rows.length) return IND_EX;
 const st={};
 for(let i=0;i<rows.length;i++){ const r=rows[i]||{}; const d=imoNz(r.d||""); if(d) st[d]=1; const c=imoNz(r.c||""); if(c) st[c]=1 }
 IND_EX={set:st, src:rows.length};
 return IND_EX;
}
/* substantivo compartilhado com o universo dos indiretos: nunca gera alerta alto */
function imoShared(tk){
 if(!tk) return false;
 try{ if(typeof LEX!=="undefined" && LEX && LEX[tk]) return true }catch(e){}
 try{ if(S.lex && S.lex[tk]) return true }catch(e){}
 try{ if(typeof CAL_SUB!=="undefined" && CAL_SUB && CAL_SUB[tk]) return true }catch(e){}
 try{ if(BASE_ATIVA==="ind" && CA_IX && CA_IX.sub){ for(const k in CA_IX.sub){ if(k.indexOf(tk+"|")===0) return true } } }catch(e){}
 return false;
}
function imoHas(list, full, tk){
 for(let i=0;i<list.length;i++){
  const w=list[i];
  if(w.indexOf(" ")>=0){ if(full.indexOf(w)>=0) return w }
  else if(tk.indexOf(w)>=0) return w;
 }
 return "";
}
function imoVerdTxt(v){
 if(!v) return "";
 if(v.estado==="imob") return "Perfil de imobilizado - confirme antes de cadastrar como indireto";
 if(v.estado==="ctrl") return "Perfil de bem controlado de baixo valor - rastreabilidade sim, imobilizado nao";
 return "Sem evidencia de imobilizado";
}
function imoConf(o){
 o=o||{};
 imoEnsure();
 const dN=imoNz(o.desc||""), cN=imoNz(o.curta||""), lN=imoNz(o.longa||"");
 const full=[dN,cN,lN].filter(function(x){return !!x}).join(" ");
 const t=imoTk(dN||cN||lN);
 const sub=t[0]||"";
 const nc=String(o.ncm||"").replace(/\D/g,"").slice(0,8);
 const temBase=IMO_IX.nRows>0;
 let nivel="", n=0, conf=0, forte=false, viaBase="";
 if(temBase){
  if(dN && IMO_IX.exato[dN]){ nivel="1 - descricao identica na base de imobilizados"; n=IMO_IX.exato[dN]; conf=0.93; forte=true }
  else if(lN && IMO_IX.exato[lN]){ nivel="1 - descricao longa identica na base de imobilizados"; n=IMO_IX.exato[lN]; conf=0.91; forte=true }
  else if(cN && (IMO_IX.curta[cN]||IMO_IX.exato[cN])){ nivel="2 - descricao curta identica na base de imobilizados"; n=(IMO_IX.curta[cN]||IMO_IX.exato[cN]); conf=0.88; forte=true }
  else if(t[0] && t[1] && IMO_IX.submod[t[0]+"|"+t[1]]){ n=IMO_IX.submod[t[0]+"|"+t[1]]; nivel="3 - substantivo + modificador iguais a um cadastro de imobilizado"; conf=Math.min(0.84,0.68+0.03*Math.min(n,5)); forte=(!imoShared(sub) && n>=2) }
  else if(nc && IMO_IX.ncm[nc]){ n=IMO_IX.ncm[nc]; nivel="4 - NCM "+nc+" presente na base de imobilizados"; conf=Math.min(0.76,0.66+0.02*Math.min(n,5)); forte=(n>=3 && !imoShared(sub)) }
  else if(sub && IMO_IX.sub[sub]){
   n=IMO_IX.sub[sub];
   const sh=imoShared(sub);
   nivel="5 - substantivo isolado "+(sh? "compartilhado com a base de indiretos (sinal fraco)" : "exclusivo da base de imobilizados");
   conf=sh? 0.35 : Math.min(0.66,0.52+0.02*Math.min(n,7));
   forte=(!sh && n>=3);
  }
  if(nivel) viaBase="cruzamento por descricao: "+nivel+" - "+n+" linha"+(n>1?"s":"")+" de "+IMO_IX.nRows+" na base de imobilizados sustentam o sinal";
 }
 const isExc=imoHas(IMO_R.exc, full, t);
 const isCtrl=imoHas(IMO_R.ctrl, full, t);
 const isPec=imoHas(IMO_R.pec, full, t);
 /* apoio por similaridade nas tres descricoes: mostra o quanto o item se
    parece com o que ja existe na base de imobilizados e com qual linha */
 var sim=[], s0=null, simTxt="";
 if(temBase){ try{ sim=imoSimTop(dN,cN,lN,3); s0=sim[0]||null }catch(e){ sim=[]; s0=null } }
 if(s0 && s0.score>=0.68){
  var pctS=Math.round(s0.score*100), fracoS=imoShared(sub);
  if(s0.score>=0.80){
   if(!nivel || conf<0.86) nivel="2 - similaridade alta nas tres descricoes ("+pctS+"% com o imobilizado "+(s0.cod||"sem codigo")+")";
   conf=Math.max(conf, Math.min(0.92, 0.60+0.35*s0.score));
   if(!fracoS && s0.mx>=1.2) forte=true;
  } else {
   if(!nivel || conf<0.70) nivel="4 - similaridade moderada nas tres descricoes ("+pctS+"%)";
   conf=Math.max(conf, Math.min(0.74, 0.45+0.35*s0.score));
  }
  simTxt="similaridade nas tres descricoes: "+pctS+"% com "+(s0.cod?("o imobilizado "+s0.cod+" - "):"")+String(s0.d||s0.c||"").slice(0,70)+" (melhor casamento pela "+s0.campo+(s0.dupla?", com mais de uma descricao concordando":"")+")";
 } else if(temBase && s0){
  simTxt="similaridade nas tres descricoes: a linha mais parecida da base de imobilizados ficou em "+Math.round(s0.score*100)+"%, abaixo do minimo para virar evidencia";
 } else if(temBase){
  simTxt="similaridade nas tres descricoes: nenhum termo em comum com a base de imobilizados";
 }
 let estado="sem", regra=IMO_R.sem, why=[];
 if(viaBase) why.push(viaBase);
 if(simTxt) why.push(simTxt);
 if(isExc){ estado="imob"; conf=Math.max(conf,0.90); regra=IMO_R.excTxt; why.push("regra da norma acionada pelo termo "+isExc+": excecao que e imobilizado independentemente do valor") }
 else if(isCtrl && !forte){ estado="ctrl"; conf=Math.max(conf,0.72); regra=IMO_R.ctrlTxt; why.push("regra da norma acionada pelo termo "+isCtrl+": bem controlado de baixo valor") }
 else if(isPec && !forte){ estado="sem"; conf=Math.min(conf||0.35,0.45); regra=IMO_R.cons; why.push("descricao de peca de reposicao ou consumivel ("+isPec+"): entra como consumo, nao como bem") }
 else if(forte){ estado="imob"; conf=Math.max(conf,0.80); regra=IMO_R.cpc }
 else if(nivel && conf>=0.55){ estado="imob"; regra=IMO_R.cpc }
 else if(nivel){ estado="sem"; conf=0.55; regra=IMO_R.sem; why.push("sinal fraco: substantivo comum as duas bases nao caracteriza imobilizado por si") }
 else { estado="sem"; conf=temBase?0.88:0.60 }
 if(estado==="sem"){
  if(!nivel) conf = isPec? 0.92 : (temBase? 0.88 : 0.62);
  else conf = 0.66;
  if(!why.length) why.push(temBase? ("nenhum dos cinco niveis de evidencia casou com as "+IMO_IX.nRows+" linhas da base de imobilizados") : "sem cruzamento de base e sem regra da norma acionada");
 }
 if(!temBase){
  conf=Math.min(conf,0.66);
  why.push("a base de imobilizados ainda nao esta carregada neste arquivo: veredito apenas pela regra da norma, sem cruzamento de linhas");
 }
 const exatoImo=!!(temBase && ((dN&&IMO_IX.exato[dN]) || (cN&&(IMO_IX.curta[cN]||IMO_IX.exato[cN])) || (lN&&IMO_IX.exato[lN])));
 let noInd=!!o.twin;
 if(exatoImo && !noInd){ try{ const ix=indEnsure().set||{}; noInd=!!(ix[dN]||ix[cN]||ix[lN]) }catch(e){} }
 const conflito=!!(exatoImo && noInd);
 if(conflito){ estado="imob"; conf=Math.max(conf,0.90); regra=IMO_R.cpc; why.push("a mesma descricao consta na base de indiretos e na base de imobilizados: o motor nao escolhe um lado") }
 return {estado:estado, conf:Math.max(0,Math.min(0.97,conf)), nivel:nivel, n:n, base:IMO_IX.nRows, sim:sim,
         regra:regra, via:why.join(" | "), conflito:conflito, txt:imoVerdTxt({estado:estado})};
}
function imoTag(v){
 if(!v) return "<span class=\"tag ok\">nao</span>";
 if(v.conflito) return "<span class=\"tag ri\">conflito</span>";
 if(v.estado==="imob") return "<span class=\"tag ri\">imobilizado</span>";
 if(v.estado==="ctrl") return "<span class=\"tag wa\">bem controlado</span>";
 return "<span class=\"tag ok\">nao</span>";
}
function imoBanner(r){
 const v=r&&r.imoc; if(!v) return "";
 const cls=v.estado==="imob"?"ri":(v.estado==="ctrl"?"wa":"ok");
 const pct=Math.round(v.conf*100);
 let h="<div class=\"note "+cls+" imoBn\" style=\"margin-top:12px\">";
 h+="<div class=\"imoHd\"><b>"+esc(imoVerdTxt(v))+"</b><span class=\"tag "+cls+"\">"+pct+"% de confianca</span></div>";
 h+="<div class=\"imoBar\"><i style=\"width:"+pct+"%\"></i></div>";
 h+="<div class=\"cx imoCx\">"+esc(v.nivel? ("nivel de evidencia: "+v.nivel) : (v.base? ("nenhum dos cinco niveis casou com as "+v.base.toLocaleString("pt-BR")+" linhas da base de imobilizados") : "base de imobilizados ainda nao carregada: veredito apenas pela regra da norma"))+"</div>";
 h+="<div class=\"cx imoCx\">regra aplicada: "+esc(v.regra)+"</div>";
 if(v.sim && v.sim.length){
  var lst=v.sim.map(function(s){ return esc((s.cod?s.cod+" - ":"")+String(s.d||s.c||"").slice(0,58)+" ("+Math.round(s.score*100)+"%)") }).join(" &middot; ");
  h+="<div class=\"cx imoCx\">imobilizados mais parecidos pelas tres descricoes: "+lst+"</div>";
 }
 if(v.conflito) h+="<div class=\"note ri\" style=\"margin-top:8px\"><b>Existe cadastro nas duas bases - decisao da Controladoria.</b> A mesma descricao consta na base de indiretos e na base de imobilizados; o motor nao escolhe um lado.</div>";
 h+="<div class=\"hint\" style=\"margin-top:6px\">Sugestao de conferencia, nunca bloqueio: divergencia de classificacao entre imobilizado e consumo e decisao da Controladoria.</div>";
 h+="</div>";
 return h;
}
function imoTranspRow(r){
 const v=r&&r.imoc; if(!v) return "";
 const cod=v.estado==="imob"?"ALERTA":(v.estado==="ctrl"?"CONTROLADO":"-");
 const lista=v.base? ("base de Materiais Imobilizados ("+v.base.toLocaleString("pt-BR")+" linhas) + IN-CPR-01.001 / RES Risco 194") : "IN-CPR-01.001 / RES Risco 194 (base de imobilizados nao carregada)";
 return "<tr><td>Conferencia de Imobilizado</td><td><span class=\"cd\">"+esc(cod)+"</span></td><td>"+esc(imoVerdTxt(v))+(v.conflito?" - existe cadastro nas duas bases":"")+"</td>"+
  "<td><span class=\"cx\">"+esc(lista)+"</span></td>"+
  "<td><span class=\"cx\">"+esc(v.nivel||"nenhum nivel de evidencia casou")+"</span></td>"+
  "<td><b style=\"color:"+cHex(v.conf)+"\">"+Math.round(v.conf*100)+"%</b></td>"+
  "<td><span class=\"cx\">"+esc(v.via||v.regra)+"</span></td></tr>";
}

/* ---------- 3. classe de avaliacao: descricao + GM + GC ---------- */
function caTop2(o){
 o=o||{};
 const ks=Object.keys(o); if(!ks.length) return null;
 let t=0; ks.forEach(function(x){ t+=o[x] });
 ks.sort(function(a,b){ return o[b]-o[a] });
 const alt=ks[1]? {k:ks[1], n:o[ks[1]], p:o[ks[1]]/t} : null;
 return {k:ks[0], n:o[ks[0]], p:o[ks[0]]/t, t:t, alt:alt};
}
function caAltTxt(x){ return (x && x.alt && x.alt.p>=0.15)? (", alternativa "+x.alt.k+" em "+Math.round(x.alt.p*100)+"%") : "" }
function caBuild(){
 CA_IX={sub:{},gm:{},tmgm:{},exato:{},sgg:{},smgg:{},n:0};
 if(!S.base||!S.base.length) return;
 for(let i=0;i<S.base.length;i++){
  const r=S.base[i]; const c=caLearn(r.ca||""); if(!c) continue; r.ca=c;
  CA_IX.n++;
  const tk=(r._t&&r._t.length)? r._t : toks((r.d||"")+" "+(r.c||""));
  const sub=tk[0]||"", mod=tk[1]||"", gm=r.gm||"", gc=r.gc||"", tm=r.tm||"";
  const push=function(o,k){ if(!k) return; if(!o[k]) o[k]={}; o[k][c]=(o[k][c]||0)+1 };
  push(CA_IX.sub, sub+"|"+gm);
  push(CA_IX.gm, gm);
  push(CA_IX.tmgm, tm+"|"+gm);
  push(CA_IX.sgg, sub+"|"+gm+"|"+gc);
  push(CA_IX.smgg, sub+"|"+mod+"|"+gm+"|"+gc);
  const ex=norm(r.d||"").replace(/\s+/g," ").trim(); if(ex) CA_IX.exato[ex]=c;
 }
}
function caPick(desc,sub,gm,tm,twinRec,gc,mod){
 const LST="coluna CLASSE DE AVALIACAO da base carregada";
 if(twinRec && twinRec.ca && codeOk("ca",twinRec.ca))
  return {c:twinRec.ca,conf:0.95,via:"nivel 1 - cadastro existente na base com a mesma descricao: classe "+twinRec.ca+" "+caDesc(twinRec.ca),lista:LST,kind:"exata",nivel:"1 - cadastro gemeo na base"};
 if(!CA_IX.n) return {c:"",conf:0,via:"a base carregada nao traz a coluna CLASSE DE AVALIACAO - campo marcado como REVISAR, sem chute",lista:caCount()? "catalogo de classes cadastrado pelo administrador" : "nenhuma lista de classe de avaliacao carregada",kind:"revisar",nivel:"sem lista de classes"};
 const ex=norm(desc||"").replace(/\s+/g," ").trim();
 if(ex && CA_IX.exato[ex] && codeOk("ca",CA_IX.exato[ex]))
  return {c:CA_IX.exato[ex],conf:0.94,via:"nivel 1 - descricao identica na base: classe "+CA_IX.exato[ex],lista:LST,kind:"exata",nivel:"1 - descricao identica na base"};
 sub=sub||""; gm=gm||""; gc=gc||""; tm=tm||""; mod=mod||"";
 let tk0="";
 try{ const _t0=toks(desc||""); tk0=_t0[0]||""; if(!mod) mod=_t0[1]||"" }catch(e){}
 const SUBS=[]; if(sub) SUBS.push(sub); if(tk0 && tk0!==sub) SUBS.push(tk0);
 if(!SUBS.length) SUBS.push("");
 const NIV=[
  {ix:CA_IX.smgg, k:function(s){ return (s&&mod&&gm&&gc)? s+"|"+mod+"|"+gm+"|"+gc : "" }, min:3, cap:0.92, a:0.60, b:0.32,
   nome:"2 - substantivo + modificador + GM + GC", txt:function(s){ return "substantivo \""+s+"\" com modificador \""+mod+"\", Grupo de Mercadoria "+gm+" e Grupo de Compradores "+gc }},
  {ix:CA_IX.sgg, k:function(s){ return (s&&gm&&gc)? s+"|"+gm+"|"+gc : "" }, min:3, cap:0.90, a:0.58, b:0.32,
   nome:"3 - substantivo + GM + GC", txt:function(s){ return "substantivo \""+s+"\" com Grupo de Mercadoria "+gm+" e Grupo de Compradores "+gc }},
  {ix:CA_IX.sub, k:function(s){ return (s||gm)? s+"|"+gm : "" }, min:3, cap:0.88, a:0.55, b:0.33,
   nome:"4 - substantivo + GM", txt:function(s){ return "substantivo \""+s+"\" com Grupo de Mercadoria "+gm }},
  {ix:CA_IX.tmgm, k:(tm||gm)? tm+"|"+gm : "", min:5, cap:0.80, a:0.50, b:0.30,
   nome:"5 - tipo de material + GM", txt:function(){ return "tipo de material "+tm+" com Grupo de Mercadoria "+gm }},
  {ix:CA_IX.gm, k:gm||"", min:5, cap:0.72, a:0.45, b:0.27,
   nome:"6 - GM isolado", txt:function(){ return "Grupo de Mercadoria "+gm }}
 ];
 for(let i=0;i<NIV.length;i++){
  const L=NIV[i];
  const KS=(typeof L.k==="function")? SUBS.map(L.k).filter(function(k){return !!k}) : (L.k? [L.k] : []);
  for(let j=0;j<KS.length;j++){
   const x=caTop2(L.ix[KS[j]]||{});
   if(!x || !codeOk("ca",x.k)) continue;
   if(x.t<L.min || x.p<0.60) continue;
   const sb=KS[j].split("|")[0];
   return {c:x.k, conf:Math.min(L.cap, L.a+L.b*x.p),
           via:"nivel "+L.nome+" - "+L.txt(sb)+": "+x.n+" de "+x.t+" materiais da base com a classe "+x.k+" ("+Math.round(x.p*100)+"%)"+caAltTxt(x)+(L.min>=5? " - evidencia mais ampla, sugerimos confirmar":""),
           lista:LST, kind:(x.p>=0.80 && i<=2)? "exata":"aproximada", nivel:L.nome,
           p:Math.round(x.p*100), t:x.t, alt:(x.alt? x.alt.k:"")};
  }
 }
 return {c:"",conf:0,via:"nenhum dos seis niveis chegou a 60% de concordancia com pelo menos 3 materiais de evidencia - campo marcado como REVISAR, sem chute",lista:LST,kind:"revisar",nivel:"sem evidencia suficiente"};
}


/* =====================================================================
   v7.6 - MONITOR DO MOTOR EM TEMPO REAL E AREA RESTRITA COM SENHA,
   SESSAO COM EXPIRACAO E BLOQUEIO POR TENTATIVAS.
   Bloco aditivo: nenhuma funcao, id ou fluxo anterior foi removido.
   ===================================================================== */

/* ---------- 1. monitor do motor (ocupa o lugar das pilulas de estado) ---------- */
var ENG={ts:0,lastN:-1,open:false,tmr:null};
var ENGC={src:-1,gmOk:0,gcOk:0,caOk:0};
function engNum(n){ try{ return Number(n||0).toLocaleString("pt-BR") }catch(e){ return String(n||0) } }
function engCov(){
 const rows=(typeof S!=="undefined" && S.base)? S.base : [];
 if(ENGC.src===rows.length) return ENGC;
 let a=0,b=0,c=0;
 for(let i=0;i<rows.length;i++){
  const r=rows[i]||{};
  if(r.gm && GM_OFF[r.gm]) a++;
  if(r.gc && GC_OFF[r.gc]) b++;
  if(r.ca && CA_OFF[r.ca]) c++;
 }
 ENGC={src:rows.length,gmOk:a,gcOk:b,caOk:c};
 return ENGC;
}
var ENGG={src:-1,list:[]};
function engCovGm(){
 const rows=(typeof S!=="undefined" && S.base)? S.base : [];
 if(ENGG.src===rows.length) return ENGG.list;
 const acc={};
 for(let i=0;i<rows.length;i++){
  const r=rows[i]||{}, g=r.gm;
  if(!g || !GM_OFF[g]) continue;
  const a=acc[g]||(acc[g]={n:0,ok:0,cls:{}});
  a.n++;
  if(r.ca && CA_OFF[r.ca]){ a.ok++; a.cls[r.ca]=(a.cls[r.ca]||0)+1 }
 }
 const list=Object.keys(acc).map(function(g){
  const a=acc[g]; let top="", tn=0;
  for(const k in a.cls){ if(a.cls[k]>tn){ tn=a.cls[k]; top=k } }
  return {gm:g, n:a.n, ok:a.ok, pct:a.n? Math.round(100*a.ok/a.n):0, top:top, topPct:a.ok? Math.round(100*tn/a.ok):0};
 }).sort(function(x,y){ return (x.pct-y.pct) || (y.n-x.n) });
 ENGG={src:rows.length, list:list};
 return list;
}
function engCovGmHtml(d){
 let L=[]; try{ L=engCovGm() }catch(e){ return "" }
 if(!d.n || !L.length) return "";
 const fracos=L.filter(function(x){ return x.pct<100 }), full=L.length-fracos.length;
 const lacuna=fracos.length>0;
 const mostra = lacuna? fracos.slice(0,6)
   : L.slice().sort(function(a,b){ return (a.topPct-b.topPct) || (b.n-a.n) }).slice(0,6);
 let h="<div class=\"engSepRow\"></div>";
 h+="<span>Base ativa</span><b>"+esc(BASE_ORIGEM||"base local")+"</b>";
 h+="<span>Grupos com classe em toda a base</span><b>"+engNum(full)+" de "+engNum(L.length)+"</b>";
 h+="<span><i>"+(lacuna? "grupos com lacuna de classe" : "grupos mais ambiguos: quanto a classe dominante cobre")+"</i></span><b>&nbsp;</b>";
 mostra.forEach(function(x){
  const nome=String(GM_OFF[x.gm]||"").split(" - ")[0].slice(0,20);
  const val=lacuna? x.pct : x.topPct;
  const dica=x.gm+" "+(GM_OFF[x.gm]||"")+" - "+engNum(x.ok)+" de "+engNum(x.n)+" materiais com classe na base"
   +(x.top? " - classe dominante "+x.top+" em "+x.topPct+"% deles" : "");
  h+="<span title=\""+esc(dica)+"\">"+esc(x.gm+" "+nome)+"</span><b>"+val+"%"+(lacuna? "" : " "+esc(x.top))+"</b>";
  h+="<div class=\"engBar\"><i style=\"width:"+val+"%\"></i></div>";
 });
 if(lacuna && fracos.length>mostra.length) h+="<span><i>e mais "+engNum(fracos.length-mostra.length)+" grupos com lacuna</i></span><b>&nbsp;</b>";
 return h;
}
function engData(){
 const d={n:0,cov:{gmOk:0,gcOk:0,caOk:0},cls:0,imo:0,gm:0,gc:0,tipos:0,termos:0,heads:0,log:0,fb:0,
          nres:0,med:0,rev:0,al:0,cf:0,caIx:0,base:"ind",baseNome:""};
 try{ d.n=(S.base||[]).length }catch(e){}
 try{ d.cov=engCov() }catch(e){}
 try{ d.cls=caCount() }catch(e){}
 try{ d.imo=imoEnsure().nRows }catch(e){}
 try{ d.gm=Object.keys(GM_OFF).length }catch(e){}
 try{ d.gc=Object.keys(GC_OFF).length }catch(e){}
 try{ d.tipos=TIPOS.length }catch(e){}
 try{ d.termos=S.lexBase.size }catch(e){}
 try{ d.heads=S.headBase.size }catch(e){}
 try{ d.log=S.log.length }catch(e){}
 try{ d.fb=S.fb.length }catch(e){}
 try{ d.caIx=CA_IX.n }catch(e){}
 try{ d.base=(typeof BASE_ATIVA!=="undefined")? BASE_ATIVA : "ind" }catch(e){}
 try{ d.baseNome=(BASE_EMB && BASE_EMB[d.base] && BASE_EMB[d.base].nome)? BASE_EMB[d.base].nome : "" }catch(e){}
 try{
  const rs=S.results||[]; d.nres=rs.length;
  let sum=0,c=0;
  for(let i=0;i<rs.length;i++){
   const r=rs[i]||{};
   if(typeof r.conf==="number"){ sum+=r.conf; c++ }
   if(!r.ca) d.rev++;
   if(r.imoc && r.imoc.estado==="imob") d.al++;
   if(r.imoc && r.imoc.conflito) d.cf++;
  }
  d.med=c? sum/c : 0;
 }catch(e){}
 return d;
}
function engHhmm(ts){
 if(!ts) return "--";
 try{ return new Date(ts).toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"}) }catch(e){ return "--" }
}
function engRender(){
 const mon=$("#engMon"); if(!mon) return;
 const d=engData();
 if(d.n!==ENG.lastN){ ENG.lastN=d.n; if(d.n) ENG.ts=Date.now() }
 const dot=$("#engDot");
 const estado = (!d.gm || !d.gc)? "ri" : (d.n? "ok" : "wa");
 if(dot) dot.className="engDot "+estado;
 const m1=$("#engM1"), m2=$("#engM2"), m3=$("#engM3");
 if(m1) m1.textContent=engNum(d.n);
 if(m2) m2.textContent=d.med? (Math.round(d.med*100)+"%") : "--";
 if(m3) m3.textContent=engNum(d.nres);
 mon.title = d.n? ("Motor pronto - "+engNum(d.n)+" linhas na base ativa") : "Base ainda nao carregada - o motor esta usando os padroes de calibracao embutidos";
 const tag=$("#engTag");
 if(tag){
  tag.className="engTag "+(estado==="ok"?"ok":(estado==="wa"?"wa":""));
  tag.textContent = estado==="ok"? "motor pronto" : (estado==="wa"? "base pendente" : "catalogo ausente");
 }
 const g=$("#engGrid");
 if(g && ENG.open){
  const pGm=d.n? d.cov.gmOk/d.n : 0, pCa=d.n? d.cov.caOk/d.n : 0;
  let h="";
  h+="<span>Base ativa</span><b>"+esc(d.base==="imo"? "Imobilizados":"Indiretos")+"</b>";
  h+="<span>Linhas carregadas</span><b>"+engNum(d.n)+"</b>";
  h+="<span>Ultima atualizacao da base</span><b>"+engHhmm(ENG.ts)+"</b>";
  h+="<div class=\"engSepRow\"></div>";
  h+="<span>Cobertura de GM valido (663)</span><b>"+(d.n? Math.round(pGm*100)+"%":"--")+"</b>";
  h+="<div class=\"engBar\"><i style=\"width:"+Math.round(pGm*100)+"%\"></i></div>";
  h+="<span>Classe de avaliacao preenchida</span><b>"+(d.n? Math.round(pCa*100)+"%":"--")+"</b>";
  h+="<div class=\"engBar\"><i style=\"width:"+Math.round(pCa*100)+"%\"></i></div>";
  h+=engCovGmHtml(d);
  h+="<span>Catalogos oficiais</span><b>"+engNum(d.gm)+" GM &middot; "+engNum(d.gc)+" GC &middot; "+engNum(d.tipos)+" tipos</b>";
  h+="<span>Classes de avaliacao no catalogo</span><b>"+engNum(d.cls)+"</b>";
  h+="<span>Linhas de imobilizado indexadas</span><b>"+engNum(d.imo)+"</b>";
  h+="<div class=\"engSepRow\"></div>";
  h+="<span>Analises nesta sessao</span><b>"+engNum(d.nres)+"</b>";
  h+="<span>Confianca media da sessao</span><b>"+(d.med? Math.round(d.med*100)+"%":"--")+"</b>";
  h+="<span>Campos em REVISAR</span><b>"+engNum(d.rev)+"</b>";
  h+="<span>Alertas de imobilizado</span><b>"+engNum(d.al)+(d.cf? " ("+engNum(d.cf)+" em conflito)":"")+"</b>";
  h+="<span>Termos aprendidos &middot; correcoes</span><b>"+engNum(d.termos)+" &middot; "+engNum(d.fb)+"</b>";
  g.innerHTML=h;
  const ft=$("#engFt");
  if(ft) ft.textContent = d.n
   ? ("Sugestao apoiada em "+engNum(d.n)+" cadastros da base ativa e "+engNum(d.log)+" eventos no log desta sessao. Codigo fora dos catalogos oficiais nunca e sugerido nem exportado.")
   : "Base ainda nao carregada neste arquivo: o motor esta trabalhando com os padroes de calibracao embutidos e com os catalogos oficiais. Codigo descontinuado nunca e sugerido nem exportado.";
 }
}
function engToggle(f){
 const pop=$("#engPop"), mon=$("#engMon"); if(!pop||!mon) return;
 ENG.open = (f===undefined)? !ENG.open : !!f;
 if(ENG.open){ pop.hidden=false; mon.classList.add("on"); mon.setAttribute("aria-expanded","true"); engRender() }
 else { pop.hidden=true; mon.classList.remove("on"); mon.setAttribute("aria-expanded","false") }
}
function engInit(){
 const mon=$("#engMon"); if(!mon) return;
 mon.addEventListener("click",function(e){ if(e.target.closest && e.target.closest("#engPop")) return; engToggle() });
 mon.addEventListener("keydown",function(e){ if(e.key==="Enter"||e.key===" "){ e.preventDefault(); engToggle() } if(e.key==="Escape") engToggle(false) });
 document.addEventListener("click",function(e){ if(ENG.open && !(e.target.closest && e.target.closest("#engMon"))) engToggle(false) });
 if(ENG.tmr) clearInterval(ENG.tmr);
 ENG.tmr=setInterval(engRender,1500);
 engRender();
}

/* ---------- 2. area restrita: senha, sessao com expiracao e bloqueio ---------- */
var ADMG={ok:false,exp:0,tries:0,lock:0,tmr:null};
function admgTtl(){ const s=$("#gateTtl"); const v=s? parseInt(s.value,10):15; return (v>0? v:15) }
function admgMs(ms){
 const t=Math.max(0,Math.ceil(ms/1000));
 return String(Math.floor(t/60)).padStart(2,"0")+":"+String(t%60).padStart(2,"0");
}
function admgMsg(txt,cls){
 const b=$("#gateMsg"); if(!b) return;
 if(!txt){ b.hidden=true; b.innerHTML=""; return }
 b.hidden=false; b.className="note gateMsg "+(cls||"");
 b.innerHTML=txt;
}
function admgApply(){
 const sec=$("#t-base"); if(!sec) return;
 const gate=$("#admGate"), sess=$("#admSess");
 if(ADMG.ok){ sec.classList.remove("admLocked"); if(gate) gate.hidden=true; if(sess) sess.hidden=false }
 else { sec.classList.add("admLocked"); if(gate) gate.hidden=false; if(sess) sess.hidden=true }
}
function admgTouch(){ if(ADMG.ok) ADMG.exp=Date.now()+admgTtl()*60000 }
function admgRelock(quiet){
 ADMG.ok=false; ADMG.exp=0;
 try{ $("#admLock").classList.remove("hidden"); $("#admArea").classList.add("hidden") }catch(e){}
 try{ if($("#catLock")){ $("#catLock").classList.remove("hidden"); $("#catAdm").classList.add("hidden") } }catch(e){}
 try{ if($("#caAdmBox")) $("#caAdmBox").classList.add("hidden") }catch(e){}
 try{ $("#gatePwd").value=""; $("#admPwd").value=""; if($("#catAdmPwd")) $("#catAdmPwd").value="" }catch(e){}
 admgApply();
 admgMsg(quiet? "" : "<b>Sessao encerrada por inatividade.</b> Informe a senha novamente para liberar o painel.","wa");
 if(!quiet){ try{ toast("sessao administrativa expirada por inatividade") }catch(e){} }
}
function admgTry(){
 const now=Date.now(), inp=$("#gatePwd"), v=inp? inp.value : "";
 if(now<ADMG.lock){
  ADMG.lock=now+300000;
  admgMsg("<b>Acesso bloqueado.</b> A tentativa reiniciou a contagem: novo desbloqueio em <b id=\"gateCd\">05:00</b>.","ri");
  if(inp) inp.value="";
  return;
 }
 if(!v){ admgMsg("Informe a senha do administrador.","wa"); return }
 if(v!==S.pwd){
  ADMG.tries++;
  if(inp){ inp.value=""; inp.focus() }
  if(ADMG.tries>=5){
   ADMG.tries=0; ADMG.lock=now+300000;
   admgMsg("<b>Cinco tentativas incorretas.</b> Acesso bloqueado por 5 minutos - desbloqueio em <b id=\"gateCd\">05:00</b>.","ri");
   try{ toast("acesso administrativo bloqueado por 5 minutos") }catch(e){}
  } else {
   admgMsg("Senha incorreta. Restam <b>"+(5-ADMG.tries)+"</b> tentativa"+((5-ADMG.tries)>1?"s":"")+" antes do bloqueio de 5 minutos.","wa");
  }
  return;
 }
 ADMG.ok=true; ADMG.tries=0; ADMG.lock=0; admgTouch();
 admgMsg("","");
 if(inp) inp.value="";
 /* libera os blocos internos reaproveitando os fluxos ja existentes */
 try{ $("#admPwd").value=S.pwd; $("#bAdm").click(); $("#admPwd").value="" }catch(e){}
 try{ if($("#catAdmPwd")){ $("#catAdmPwd").value=S.pwd; $("#bCatAdm").click(); $("#catAdmPwd").value="" } }catch(e){}
 /* garantia: se o clique nao tiver liberado os blocos internos, libera direto */
 try{ if($("#admLock")){ $("#admLock").classList.add("hidden"); $("#admArea").classList.remove("hidden") } }catch(e){}
 try{ if($("#catLock") && $("#catAdm")){ $("#catLock").classList.add("hidden"); $("#catAdm").classList.remove("hidden") } }catch(e){}
 admgApply(); engRender();
 try{ toast("painel administrativo liberado por "+admgTtl()+" minutos de inatividade") }catch(e){}
}
function admgTick(){
 const now=Date.now();
 if(ADMG.ok){
  const cd=$("#admSessCd");
  if(cd) cd.textContent=admgMs(ADMG.exp-now);
  if(now>=ADMG.exp) admgRelock(false);
  return;
 }
 if(now<ADMG.lock){
  const c=$("#gateCd");
  if(c) c.textContent=admgMs(ADMG.lock-now);
  const b=$("#bGate"); if(b) b.disabled=false;
 } else if(ADMG.lock){
  ADMG.lock=0;
  admgMsg("Bloqueio encerrado. Voce pode informar a senha novamente.","");
 }
}
function admgInit(){
 const sec=$("#t-base"); if(!sec) return;
 admgApply();
 const b=$("#bGate"); if(b) b.onclick=admgTry;
 const p=$("#gatePwd"); if(p) p.onkeydown=function(e){ if(e.key==="Enter"){ e.preventDefault(); admgTry() } };
 const t=$("#gateTtl"); if(t) t.onchange=function(){ admgTouch(); try{ toast("expiracao por inatividade: "+admgTtl()+" minutos") }catch(e){} };
 const lk=$("#bAdmLock"); if(lk) lk.onclick=function(){ admgRelock(true); try{ toast("painel administrativo bloqueado") }catch(e){} };
 sec.addEventListener("click",admgTouch,true);
 sec.addEventListener("keydown",admgTouch,true);
 sec.addEventListener("change",admgTouch,true);
 if(ADMG.tmr) clearInterval(ADMG.tmr);
 ADMG.tmr=setInterval(admgTick,1000);
}
function v76Init(){ try{ engInit() }catch(e){} try{ admgInit() }catch(e){} }
if(typeof document!=="undefined"){
 if(document.readyState==="loading") document.addEventListener("DOMContentLoaded",function(){ setTimeout(v76Init,60) });
 else setTimeout(v76Init,60);
}

