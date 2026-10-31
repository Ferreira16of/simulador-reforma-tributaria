// ===== Slide 20 (adicional): linhas consideradas no redutor =====
const LIN={
graos:{cp:['Transcerrados – CP Proteção','Transcerrados – CP Fixa','Aditivo – CP Fixa'],no:'Receita Tarifária (Transcerrados e Aditivo), Receita Acessória, Receita de Construção, Remuneração do Ativo Financeiro'},
mobi:{cp:['Contraprestação Pecuniária'],no:'Desconto IFRS 09, Rec. Acessórias (Locação, Publicidade), Naming Rights, Remuneração do Ativo Financeiro, Receita de Construção'},
leste:{cp:['Contraprestação OPEX'],no:'Receita Financeira IFRS 09, ABL Interna, Publicidade, Locação de Espaço, Naming Right'},
norte:{cp:['Contraprestação CAPEX','Contraprestação OPEX'],no:'Receita Acessória e Receita de Construção'},
sul:{cp:['Contraprestação CAPEX','Contraprestação OPEX'],no:'Receita Acessória e Receita de Construção'},
mercosul:{cp:[],no:'Sem contraprestação: receita tarifária (Argentina e Brasil), serviços CUF, acessórias e recomposição'},
mcp01:{cp:[],no:'Sem contraprestação identificada na planilha (Receita Bruta de Serviços)'},
welfare:{cp:[],no:'Sem contraprestação: receita por município (saneamento)'},
lote6:{cp:[],no:'Sem contraprestação identificada na planilha'},
rota:{cp:[],no:'Sem contraprestação: Receita Tarifária (pedágio) e Receita Reequilíbrio'}};
function Pzero(){const q=Object.assign({},P);q.redutor={};YEARS.forEach(y=>q.redutor[y]=0);return q}
function shareOf(k,y){const Y=B.empresas[k].years[y];return(Y&&Y.comp&&Y.comp.RB&&Y.comp.RB.shareCP)||0}
function dynRedutor(){
  const cv=$('cv');const P0=Pzero();const y=2027;
  const rd=P.redutor[2027],mk=(t,s,b,c,al,ex)=>`<div class="pp" style="font-size:${s}px;${b?'font-weight:700;':''}${c?'color:'+c+';':''}text-align:${al||'left'};${ex||''}">${t}</div>`;
  const hdr=['Empresa','Linhas de receita COM redutor','Linhas SEM redutor','% da base '+2027,'% da base '+2033,'Efeito CBS+IBS '+y+' (R$ mil)','Efeito no LL '+y+' (R$ mil)'];
  const W=[130,330,360,70,70,120,124.8];
  let h=`<table class="tb" style="left:36.48px;top:144px;width:1204.8px"><colgroup>${W.map(w=>`<col style="width:${w}px">`).join('')}</colgroup><tr style="height:38px">${hdr.map((t,i)=>`<td style="background:#17365D;padding:2px 5px;vertical-align:middle">${mk(t,11.3,1,'#fff',i>2?'center':'left')}</td>`).join('')}</tr>`;
  let tR=0,tL=0;
  ORS().forEach((k,i)=>{const L=LIN[k],r=CY(k,y),r0=compute(B.empresas[k],k,y,P0,B.audit);
    const dR=r&&r0?r.rt.RB-r0.rt.RB:0,dL=r&&r0?r.rt.LL-r0.rt.LL:0;tR+=dR;tL+=dL;
    const bg=i%2?'#F4F7FA':'#fff',has=L.cp.length>0;
    const bd='border-bottom:1px solid #E3E8EE;';
    h+=`<tr style="height:46px"><td style="background:${bg};${bd}padding:2px 5px;vertical-align:middle">${mk(NAMES[k],11.3,1,'#17365D')}</td>`
     +`<td style="background:${has?'#FFF4D6':bg};${bd}padding:2px 5px;vertical-align:middle">${has?L.cp.map(t=>mk('• '+t,11.3,0,'#293746')).join(''):mk('—',11.3,0,'#5B6573')}</td>`
     +`<td style="background:${bg};${bd}padding:2px 5px;vertical-align:middle">${mk(L.no,10.7,0,'#5B6573')}</td>`
     +[shareOf(k,2027),shareOf(k,2033)].map(s=>`<td style="background:${bg};${bd}padding:2px 5px;vertical-align:middle">${mk(has?p1(s):'—',11.3,0,'#293746','center')}</td>`).join('')
     +[dR,dL].map(v=>`<td style="background:${bg};${bd}padding:2px 5px;vertical-align:middle">${mk(has?fmt(v):'—',11.3,0,has?colS(v):null,'center')}</td>`).join('')+'</tr>'});
  h+=`<tr style="height:30px"><td colspan="5" style="background:#EAF0F7;padding:2px 5px;vertical-align:middle">${mk(`Total (${ORS().length} empresas)`,11.3,1,'#17365D')}</td><td style="background:#EAF0F7;padding:2px 5px;vertical-align:middle">${mk(fmt(tR),11.3,1,colS(tR),'center')}</td><td style="background:#EAF0F7;padding:2px 5px;vertical-align:middle">${mk(fmt(tL),11.3,1,colS(tL),'center')}</td></tr></table>`;
  const mx=(a,b)=>`<div class="sh hast" style="left:${a}px;top:${b[0]}px;width:${b[1]}px;height:${b[2]}px"><div class="tx" style="padding:3.6px 7.2px">${b[3]}</div></div>`;
  const ttl=`<div class="sh hast" style="left:36px;top:51px;width:931px;height:33px"><div class="tx">${mk('REDUTOR DE '+(rd*100).toLocaleString('pt-BR',{maximumFractionDigits:2})+'% | LINHAS CONSIDERADAS POR EMPRESA',26,1,'#17365D')}</div></div>`
   +`<div class="sh hast" style="left:36px;top:87px;width:900px;height:19px"><div class="tx">${mk('Redutor aplicado à contraprestação em todos os anos (2026–2033) | CBS atual do cenário: '+rate(P.cbs[2027])+' | efeito = com redutor − sem redutor',12,0,'#5B6573')}</div></div>`
   +`<div class="sh hast" style="left:36px;top:121px;width:682px;height:18px"><div class="tx">${mk('QUAIS RECEITAS RECEBEM O REDUTOR, EMPRESA A EMPRESA',12,1,'#17365D')}</div></div>`;
  const bl=['O redutor reduz a CBS e o IBS incidentes sobre a Receita Bruta na proporção da contraprestação (coluna “% da base”): nova alíquota efetiva = alíquota × (1 − redutor × % da base).',
   'Só entram as linhas de contraprestação (pecuniária, CAPEX, OPEX, fixa e proteção). Receita tarifária, acessória, de construção e financeira não são reduzidas.',
   'Não alterado: créditos sobre custos, ISS e IR/CSLL. Mercosul, MCP01, Welfare e Rota não têm contraprestação e não sofrem o redutor.',
   'Em Norte e Sul a contraprestação (CAPEX e OPEX) é 100% da Receita Bruta tributada; por isso o redutor incide sobre toda a CBS/IBS dessas duas empresas.'];
  const bt=`<div class="sh hast" style="left:36px;top:${144+38+ORS().length*46+30+10}px;width:1204px;height:90px"><div class="tx">${bl.map(t=>mk('• '+t,12,0,'#293746','left','margin-top:2px')).join('')}</div></div>`;
  cv.insertAdjacentHTML('beforeend',ttl+h+bt);
}
// ===== Aplicação =====
let CUR=0;let SL=SLIDES.slice();let BASE=null;
function fit(){const t=$('stage').getBoundingClientRect();const s=Math.min((t.width-(innerWidth<700?8:32))/1280,(t.height-(innerWidth<700?8:20))/720);$('cv').style.transform='scale('+s+')';$('box').style.width=(1280*s)+'px';$('box').style.height=(720*s)+'px'}
function mkS20(){if(SL.length>=20){BASE=SL.slice(0,20);return}const b=JSON.parse(JSON.stringify(SL[18]));b.i=20;b.el=b.el.filter(e=>e.slnum&&!e.p||e.slnum);b.el=b.el.filter(e=>e.slnum);const f=SL[18].lay.concat([]);SL.push({i:20,bg:SL[18].bg,lay:SL[18].lay,el:SL[18].el.filter(e=>e.slnum)});BASE=SL.slice()}
function mkL6(){const b=JSON.parse(JSON.stringify(BASE[12]));b.i=21;b.el.forEach(e=>{if(e.p)e.p.forEach(p=>p.runs.forEach(r=>{r.t=r.t.replace('ROTA DA INTEGRAÇÃO','LOTE 6')}))});return b}
function mkSL(){SL=BASE.slice();if(L6)SL.push(mkL6())}
function render(){inv();const S=SL[CUR];$('cv').innerHTML=slideHtml(S);$('cv').style.background=S.bg;$('box').style.background=S.bg;
 document.querySelectorAll('#cv .snv').forEach(e=>e.textContent=S.i);
 dynamic(S.i);
 if(L6&&S.i>=17&&S.i<=19){const bx=[...document.querySelectorAll('#cv .sh.hast')].find(e=>e.textContent.trim().startsWith('• '));if(bx)bx.style.top=(parseFloat(bx.style.top)+24)+'px'}
 if(S.i===21){document.querySelectorAll('#cv .sh.hast').forEach(e=>{if(e.textContent.trim().startsWith('Considerações'))e.style.display='none'});document.querySelectorAll('#cv table.tb').forEach(t=>{if(parseFloat(t.style.top)>570)t.style.display='none'})}
 if(L6)document.querySelectorAll('#cv span').forEach(s=>{if(s.textContent.includes('9 empresas'))s.textContent=s.textContent.replace(/\b9 empresas/g,'10 empresas')});
 if(S.i===2)wireIdx();
 $('pos').textContent=(CUR+1)+' / '+SL.length;$('prev').disabled=CUR===0;$('next').disabled=CUR===SL.length-1;fit();try{history.replaceState(null,'','#'+(CUR+1))}catch(e){}}
function go(i){CUR=Math.max(0,Math.min(SL.length-1,i));render()}
function addIdx20(){const all=[...document.querySelectorAll('#cv .sh')];const num=n=>all.find(e=>e.textContent.trim()===n&&e.classList.contains('hast'));
 const a=num('18'),b=num('19');if(!a||!b)return;const ta=parseFloat(a.style.top),tb=parseFloat(b.style.top),st=tb-ta;
 all.filter(e=>Math.abs(parseFloat(e.style.top)-tb)<1.5&&Math.abs(parseFloat(e.style.left)-parseFloat(b.style.left))<400).forEach(e=>{const c=e.cloneNode(true);c.style.top=(parseFloat(e.style.top)+st)+'px';
  c.querySelectorAll('span').forEach(sp=>{if(sp.textContent.trim()==='19')sp.textContent='20';else if(sp.textContent.trim().startsWith('Fluxo 2027 por empresa'))sp.textContent='Redutor | linhas consideradas por empresa'});e.parentNode.appendChild(c)});if(L6){const all2=[...document.querySelectorAll('#cv .sh')];const n20=all2.find(e=>e.textContent.trim()==='20'&&e.classList.contains('hast'));if(n20){const st2=tb-ta;all2.filter(e=>Math.abs(parseFloat(e.style.top)-parseFloat(n20.style.top))<1.5&&Math.abs(parseFloat(e.style.left)-parseFloat(n20.style.left))<400).forEach(e=>{const c=e.cloneNode(true);c.style.top=(parseFloat(e.style.top)+st2)+'px';c.querySelectorAll('span').forEach(sp=>{if(sp.textContent.trim()==='20')sp.textContent='21';else if(sp.textContent.trim().startsWith('Redutor'))sp.textContent='Lote 6 | DRE Atual × RT (opcional)'});e.parentNode.appendChild(c)})}}}
function wireIdx(){addIdx20();/* índice (slide 2): números clicáveis */
 document.querySelectorAll('#cv .sh.hast').forEach(e=>{const m=e.textContent.trim().match(/^(\d{2})\b/);if(m&&e.textContent.length<80&&+m[1]>=3&&+m[1]<=21){e.style.cursor='pointer';e.onclick=()=>go(+m[1]-1)}})}
$('prev').onclick=()=>go(CUR-1);$('next').onclick=()=>go(CUR+1);
$('sim').onclick=()=>$('drawer').classList.toggle('on');$('dclose').onclick=()=>$('drawer').classList.remove('on');
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
document.addEventListener('keydown',e=>{if(/INPUT|SELECT|TEXTAREA/.test(document.activeElement.tagName))return;if(e.key==='ArrowRight'||e.key==='PageDown')go(CUR+1);if(e.key==='ArrowLeft'||e.key==='PageUp')go(CUR-1);if(e.key==='Escape')$('drawer').classList.remove('on')});
let tx=0,ty=0;document.addEventListener('touchstart',e=>{tx=e.touches[0].clientX;ty=e.touches[0].clientY},{passive:true});document.addEventListener('touchend',e=>{if($('drawer').classList.contains('on')||e.touches.length||(window.visualViewport&&visualViewport.scale>1.05))return;const dx=e.changedTouches[0].clientX-tx,dy=e.changedTouches[0].clientY-ty;if(Math.abs(dx)>70&&Math.abs(dx)>Math.abs(dy)*2)go(CUR+(dx<0?1:-1))},{passive:true});
mkS20();mkSL();buildParams();buildEmp();buildAud();
const h0=parseInt((location.hash||'').slice(1));go(h0>0?h0-1:0);
try{document.fonts.ready.then(fit)}catch(e){}
