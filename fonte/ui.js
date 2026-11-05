// ===== Slide 20 (adicional): linhas consideradas no redutor =====
const LIN={
graos:{cp:['Receita Contraprestação (Transcerrados – CP Proteção, CP Fixa e Aditivo – CP Fixa)'],loc:['Receita Acessória'],no:'Receita Tarifária (Transcerrados e Aditivo); Remuneração do Ativo Financeiro e Construção (fora da base do CBS)'},
mobi:{cp:['Contraprestação Pecuniária (líquida do Desconto IFRS 09): é a base do CBS na DF em todos os anos'],loc:['Rec. Acessórias – Locação','Publicidade – Outdoor e Mercado','Naming Rights'],no:'Remuneração do Ativo Financeiro e Construção (fora da base do CBS)'},
leste:{cp:['Contraprestação OPEX','Receita Financeira IFRS 09'],loc:['Locação Espaço p/ Construção (4 linhas)','ABL Interna – Todos Terminais','Publicidade – Todos Terminais','Naming Right'],no:'Receita de Construção'},
norte:{cp:['Receita de Operação – Contábil','Receita Financeira (IFRIC 12)'],loc:[],no:'Receita Acessória (zerada) e Receita de Construção'},
sul:{cp:['Receita de Operação – Contábil','Receita Financeira (IFRIC 12)'],loc:[],no:'Receita Acessória (zerada) e Receita de Construção'},
mercosul:{cp:[],loc:[],no:'Sem redução: receita tarifária (Argentina e Brasil), serviços CUF, acessórias e recomposição'},
mcp01:{cp:['Receita Bruta (toda a receita)'],loc:[],no:'Receita de Construção'},
welfare:{cp:[],loc:[],no:'Sem redução: receita por município (saneamento)'},
lote6:{cp:[],loc:['Receita Acessória (1,5% da receita, proporção da planilha 9906)'],no:'Receita Tarifária'},
rota:{cp:[],loc:[],no:'Sem redução: Receita Tarifária (pedágio) e Receita Reequilíbrio'}};
function Pzero(){const q=Object.assign({},P);q.redutor={};q.redutorLoc={};YEARS.forEach(y=>{q.redutor[y]=0;q.redutorLoc[y]=0});return q}
function shareLoc(k,y){const Y=B.empresas[k].years[y];return(Y&&Y.comp&&Y.comp.RB&&Y.comp.RB.shareLoc)||0}
function shareOf(k,y){const Y=B.empresas[k].years[y];return(Y&&Y.comp&&Y.comp.RB&&Y.comp.RB.shareCP)||0}
function dynRedutor(){
  const cv=$('cv');const P0=Pzero();const y=2027;
  const rd=P.redutor[2027],rl=P.redutorLoc[2027],pn=v=>(v*100).toLocaleString('pt-BR',{maximumFractionDigits:2}),
   mk=(t,s,b,c,al,ex)=>`<div class="pp" style="font-size:${s}px;${b?'font-weight:700;':''}${c?'color:'+c+';':''}text-align:${al||'left'};${ex||''}">${t}</div>`;
  const hdr=['Empresa','Contraprestação (COM redutor)','Locação (COM redução)','Linhas SEM redutor','Débito 2027: receita (R$ mil) e alíquota utilizada'];
  const W=[112,215,205,263,409.8];
  const n=ORS().length,rh=n>9?30:36;
  let h=`<table id="tredut" class="tb" style="left:36.48px;top:128px;width:1204.8px"><colgroup>${W.map(w=>`<col style="width:${w}px">`).join('')}</colgroup><tr style="height:36px">${hdr.map((t,i)=>`<td style="background:#17365D;padding:2px 5px;vertical-align:middle">${mk(t,10.6,1,'#fff','left')}</td>`).join('')}</tr>`;
  let tR=0,tL=0;
  ORS().forEach((k,i)=>{const L=LIN[k],r=CY(k,y),r0=compute(B.empresas[k],k,y,P0,B.audit);
    const dR=r&&r0?r.rt.RB-r0.rt.RB:0,dL=r&&r0?r.rt.LL-r0.rt.LL:0;tR+=dR;tL+=dL;
    const bg=i%2?'#F4F7FA':'#fff',hasC=L.cp.length>0,hasL=L.loc.length>0,has=hasC||hasL;
    const bd='border-bottom:1px solid #E3E8EE;';const td=(c,bgc)=>`<td style="background:${bgc||bg};${bd}padding:2px 5px;vertical-align:middle">${c}</td>`;
    h+=`<tr style="height:${rh}px">`+td(mk(NAMES[k],10.8,1,'#17365D'))
     +td(hasC?L.cp.map(t=>mk('• '+t,10.4,0,'#293746')).join(''):mk('—',10.6,0,'#5B6573'),hasC?'#FFF4D6':null)
     +td(hasL?L.loc.map(t=>mk('• '+t,10.4,0,'#293746')).join(''):mk('—',10.6,0,'#5B6573'),hasL?'#E3F2E6':null)
     +td(mk(L.no,10,0,'#5B6573'))
     +td(debLinhas(k).ln.map(a=>mk(a[0]+': <b>'+a[1]+'</b>',10.2,0,'#293746')).join(''))+'</tr>'});
  let rbT=0;ORS().forEach(k=>{rbT+=debLinhas(k).rb});
  h+=`<tr style="height:28px"><td colspan="4" style="background:#EAF0F7;padding:2px 5px;vertical-align:middle">${mk(`Total (${n} empresas)`,10.8,1,'#17365D')}</td><td style="background:#EAF0F7;padding:2px 5px;vertical-align:middle">${mk('Receita Bruta 2027: R$ '+fmt(rbT)+' mil',10.8,1,'#17365D')}</td></tr></table>`;
  const ttl=`<div class="sh hast" style="left:36px;top:51px;width:1200px;height:33px"><div class="tx">${mk('REDUTORES | CONTRAPRESTAÇÃO '+pn(rd)+'% E LOCAÇÃO '+pn(rl)+'%',26,1,'#17365D')}</div></div>`
   +`<div class="sh hast" style="left:36px;top:87px;width:1200px;height:19px"><div class="tx">${mk('Aplicados a todos os anos (2026–2033) | CBS do cenário: '+rate(P.cbs[2027])+' | soma das receitas de cada empresa = Receita Bruta 2027',12,0,'#5B6573')}</div></div>`
   +`<div class="sh hast" style="left:36px;top:107px;width:900px;height:18px"><div class="tx">${mk('QUAIS RECEITAS RECEBEM A REDUÇÃO, EMPRESA A EMPRESA',12,1,'#17365D')}</div></div>`;
  cv.insertAdjacentHTML('beforeend',ttl+h);
  const t=$('tredut'),top=t.offsetTop+t.offsetHeight+8;
  const bl=['A CBS e o IBS da Receita Bruta caem na proporção de cada linha: alíquota efetiva = alíquota × (1 − redutor × % contraprestação − redução × % locação). Linhas reduzidas conforme a planilha Base_Redução: 44,2% nas contraprestações e receitas financeiras (e em toda a Receita Bruta da MCP01); 70% na receita acessória, locação, ABL, publicidade e naming rights.',
   'Receita tarifária, construção e as demais linhas fora da tabela não são reduzidas. Créditos, ISS e IR/CSLL não mudam (o crédito segue o % de fornecedores no Regime Normal).'];
  cv.insertAdjacentHTML('beforeend',`<div class="sh hast" style="left:36px;top:${top}px;width:1204px;height:70px"><div class="tx">${bl.map(x=>mk('• '+x,11,0,'#293746','left','margin-top:2px')).join('')}</div></div>`);
}
// ===== Aplicação =====
let CUR=0;let SL=SLIDES.slice();let BASE=null;
function fit(){const t=$('stage').getBoundingClientRect();const s=Math.min((t.width-(innerWidth<700?8:32))/1280,(t.height-(innerWidth<700?8:20))/720);$('cv').style.transform='scale('+s+')';$('box').style.width=(1280*s)+'px';$('box').style.height=(720*s)+'px'}
function mkS20(){if(SL.length>=20){BASE=SL.slice(0,20);return}const b=JSON.parse(JSON.stringify(SL[18]));b.i=20;b.el=b.el.filter(e=>e.slnum&&!e.p||e.slnum);b.el=b.el.filter(e=>e.slnum);const f=SL[18].lay.concat([]);SL.push({i:20,bg:SL[18].bg,lay:SL[18].lay,el:SL[18].el.filter(e=>e.slnum)});BASE=SL.slice()}
function mkL6(){const b=JSON.parse(JSON.stringify(BASE[12]));b.i=21;b.el.forEach(e=>{if(e.p)e.p.forEach(p=>p.runs.forEach(r=>{r.t=r.t.replace('ROTA DA INTEGRAÇÃO','LOTE 6')}))});return b}
function mkSL(){SL=BASE.slice();if(L6){const ci=SL.findIndex(s=>s.i===14);SL.splice(ci<0?SL.length:ci,0,mkL6())}}
function fxText(){const FX=[['Fonte: planilha Aliquotas_Fluxo (aba Fluxo) e DF - Mensal das empresas, 2027.','Fonte: DF - Mensal 2027 de cada empresa (planilhas de origem), recalculada com as premissas da simulação.'],['Créditos de CAPEX conforme a planilha (zerados no fluxo mensal de Grãos do Piauí, MCP01 e Rota da Integração).','Créditos de CAPEX conforme a DF - Mensal de cada empresa (planilhas de origem).'],['RT: aba Fluxo 2027,','RT: fluxo 2027 (DF - Mensal),'],['RT: planilha Fluxo 2027,','RT: fluxo 2027 (DF - Mensal),'],['Os débitos do modelo atual incluem todas as deduções da DRE.','Os débitos (Atual e RT) incluem todas as deduções da DRE, inclusive o IIBB da Mercosul.']];
 document.querySelectorAll('#cv span').forEach(s=>{let t=s.textContent,r=t;FX.forEach(([a,b])=>{if(r.includes(a))r=r.replace(a,b)});if(r!==t)s.textContent=r})}
function render(){inv();live();const S=SL[CUR];$('cv').innerHTML=slideHtml(S);$('cv').style.background=S.bg;$('box').style.background=S.bg;
 document.querySelectorAll('#cv .snv').forEach(e=>e.textContent=CUR+1);
 dynamic(S.i);
 if(L6&&S.i>=17&&S.i<=19){const bx=[...document.querySelectorAll('#cv .sh.hast')].find(e=>e.textContent.trim().startsWith('• '));if(bx)bx.style.top=(parseFloat(bx.style.top)+24)+'px'}
 if(S.i===21){document.querySelectorAll('#cv .sh.hast').forEach(e=>{if(e.textContent.trim().startsWith('Considerações'))e.style.display='none'});document.querySelectorAll('#cv table.tb').forEach(t=>{if(parseFloat(t.style.top)>570)t.style.display='none'})}
 if(L6&&S.i!==17)document.querySelectorAll('#cv span').forEach(s=>{if(s.textContent.includes('9 empresas'))s.textContent=s.textContent.replace(/\b9 empresas/g,'10 empresas')});
 if(L6)document.querySelectorAll('#cv span').forEach(s=>{if(s.textContent.includes('slides 5 a 13'))s.textContent=s.textContent.replace('slides 5 a 13','slides 5 a 14')});
 if(L6)document.querySelectorAll('#cv span').forEach(s=>{if(s.textContent.includes('o slide 16'))s.textContent=s.textContent.replace('o slide 16','o slide 17')});
 fxText();
 if(S.i===2)wireIdx();
 $('pos').textContent=(CUR+1)+' / '+SL.length;$('prev').disabled=CUR===0;$('next').disabled=CUR===SL.length-1;fit();try{history.replaceState(null,'','#'+(CUR+1))}catch(e){}}
function live(){const el=$('live');if(!el)return;try{const row=y=>{const c=CONS(y);if(!c)return '';const d=c.rt.LL-c.atual.LL,db=(c.rt.RB-c.atual.RB);return `<tr><td>${y}</td><td>${fmt(c.atual.LL)}</td><td>${fmt(c.rt.LL)}</td><td class="${cls(d)}">${fmt(d)}</td><td class="${cls(db)}">${fmt(db)}</td></tr>`};
 el.innerHTML='<table><thead><tr><th>Consolidado (R$ mil)</th><th>LL Atual</th><th>LL RT</th><th>Δ LL</th><th>Δ Receita</th></tr></thead><tbody>'+[2027,2028,2029,2031].map(row).join('')+'</tbody></table>'}catch(e){el.textContent=''}}
function go(i){CUR=Math.max(0,Math.min(SL.length-1,i));render()}
function addIdx20(){const all=[...document.querySelectorAll('#cv .sh')];const num=n=>all.find(e=>e.textContent.trim()===n&&e.classList.contains('hast'));
 const a=num('18'),b=num('19');if(!a||!b)return;const ta=parseFloat(a.style.top),tb=parseFloat(b.style.top),st=tb-ta;
 all.filter(e=>Math.abs(parseFloat(e.style.top)-tb)<1.5&&Math.abs(parseFloat(e.style.left)-parseFloat(b.style.left))<400).forEach(e=>{const c=e.cloneNode(true);c.style.top=(parseFloat(e.style.top)+st)+'px';
  c.querySelectorAll('span').forEach(sp=>{if(sp.textContent.trim()==='19')sp.textContent='20';else if(sp.textContent.trim().startsWith('Fluxo 2027 por empresa'))sp.textContent='Redutor | linhas consideradas por empresa'});e.parentNode.appendChild(c)});if(L6){const all2=[...document.querySelectorAll('#cv .sh')].filter(e=>e.classList.contains('hast'));
  const nm=n=>all2.filter(e=>e.textContent.trim()===n&&parseFloat(e.style.left)<100||false);
  const n12=all2.find(e=>e.textContent.trim()==='12'&&parseFloat(e.style.left)<100),n13=all2.find(e=>e.textContent.trim()==='13'&&parseFloat(e.style.left)<100);
  if(n12&&n13){const st3=parseFloat(n13.style.top)-parseFloat(n12.style.top),T13=parseFloat(n13.style.top);
   // renumera a coluna da direita (14..20 -> 15..21) antes de criar a nova linha
   all2.filter(e=>parseFloat(e.style.left)>600&&/^(1[4-9]|20)$/.test(e.textContent.trim())).forEach(e=>e.querySelectorAll('span').forEach(sp=>{const t=sp.textContent.trim();if(/^(1[4-9]|20)$/.test(t))sp.textContent=String(+t+1)}));
   [...document.querySelectorAll('#cv .sh')].filter(e=>Math.abs(parseFloat(e.style.top)-T13)<8&&parseFloat(e.style.left)<660&&!/^[A-E]$/.test(e.textContent.trim())).forEach(e=>{const c=e.cloneNode(true);c.style.top=(T13+st3)+'px';
    c.querySelectorAll('span').forEach(sp=>{const t=sp.textContent.trim();if(t==='13')sp.textContent='14';else if(t.startsWith('Rota'))sp.textContent='Lote 6'});e.parentNode.appendChild(c)})}}}
function wireIdx(){addIdx20();document.querySelectorAll('#cv span').forEach(sp=>{if(sp.textContent.includes('Clique no número'))sp.textContent=sp.textContent.replace('Clique no número','Clique na linha')});/* índice (slide 2): a linha inteira (número, título e caixa) é clicável */
 const all=[...document.querySelectorAll('#cv .sh')];
 all.forEach(e=>{if(!e.classList.contains('hast'))return;const m=e.textContent.trim().match(/^(\d{2})$/);if(!m||+m[1]<3||+m[1]>21)return;
  const n=+m[1],T=parseFloat(e.style.top),L=parseFloat(e.style.left);
  all.filter(q=>Math.abs(parseFloat(q.style.top)-T)<4&&parseFloat(q.style.left)>=L-1&&parseFloat(q.style.left)<L+560).forEach(q=>{q.style.cursor='pointer';q.title='Ir para o slide '+n;q.onclick=ev=>{ev.stopPropagation();go(n-1)};q.onmouseenter=()=>{q.style.filter='brightness(.93)'};q.onmouseleave=()=>{q.style.filter=''}})})}
$('prev').onclick=()=>go(CUR-1);$('next').onclick=()=>go(CUR+1);
const dock=on=>{$('drawer').classList.toggle('on',on);document.body.classList.toggle('dopen',on&&innerWidth>=1500);fit()};
$('sim').onclick=()=>dock(!$('drawer').classList.contains('on'));$('dclose').onclick=()=>dock(false);
$('rst').onclick=()=>{P=DEF();buildParams();buildEmp();buildAud();render()};
$('rst0').onclick=()=>{P=defaultParams();buildParams();buildEmp();buildAud();render()};
$('exp').onclick=()=>{const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([JSON.stringify({params:P},null,1)],{type:'application/json'}));a.download='cenario_reforma.json';a.click()};
$('imp').onclick=()=>$('file').click();
$('file').onchange=e=>{const f=e.target.files[0];if(!f)return;f.text().then(t=>{const j=JSON.parse(t);P=Object.assign(DEF(),j.params||j);P.audit=Object.assign(DEF().audit,P.audit||{});buildParams();buildEmp();buildAud();render()})};
$('l6').onchange=e=>{L6=e.target.checked;mkSL();CUR=Math.min(CUR,SL.length-1);render()};
$('pdf').onclick=()=>{const keep=CUR;let h='';for(let i=0;i<SL.length;i++){CUR=i;render();h+=`<div class="pg" style="background:${SL[i].bg}"><div class="cv2">${$('cv').innerHTML}</div></div>`}$('pr').innerHTML=h;CUR=keep;render();
 const clean=()=>{$('pr').innerHTML='';window.removeEventListener('afterprint',clean)};window.addEventListener('afterprint',clean);setTimeout(()=>window.print(),100)};
$('idx').onclick=()=>go(1);
$('full').onclick=()=>{try{document.fullscreenElement?document.exitFullscreen():document.documentElement.requestFullscreen()}catch(e){}};
window.addEventListener('resize',fit);
document.addEventListener('keydown',e=>{if(/INPUT|SELECT|TEXTAREA/.test(document.activeElement.tagName))return;if(e.key==='ArrowRight'||e.key==='PageDown')go(CUR+1);if(e.key==='ArrowLeft'||e.key==='PageUp')go(CUR-1);if(e.key==='Escape')dock(false)});
let tx=0,ty=0;document.addEventListener('touchstart',e=>{tx=e.touches[0].clientX;ty=e.touches[0].clientY},{passive:true});document.addEventListener('touchend',e=>{if($('drawer').classList.contains('on')||e.touches.length||(window.visualViewport&&visualViewport.scale>1.05))return;const dx=e.changedTouches[0].clientX-tx,dy=e.changedTouches[0].clientY-ty;if(Math.abs(dx)>70&&Math.abs(dx)>Math.abs(dy)*2)go(CUR+(dx<0?1:-1))},{passive:true});
mkS20();mkSL();buildParams();buildEmp();buildAud();
const h0=parseInt((location.hash||'').slice(1));go(h0>0?h0-1:0);
try{document.fonts.ready.then(fit)}catch(e){}
