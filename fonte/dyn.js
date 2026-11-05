// ===== Camada dinâmica: recalcula os números das formas/tabelas do PPT v15 =====
const OR9=['graos','mobi','leste','mercosul','norte','sul','mcp01','welfare','rota'];
const ORS=()=>L6?OR9.concat(['lote6']):OR9;
function grow(tb,base){if(!L6||!tb)return;const r=tb.rows[base];const c=r.cloneNode(true);r.parentNode.insertBefore(c,r.nextSibling)}
const RED='#C00000',GRN='#70AD47';
const rhe=x=>{const f=Math.floor(x+1e-9),dd=x-f;if(Math.abs(dd-0.5)<1e-6)return f%2===0?f:f+1;return Math.round(x)};
const p1=(v)=>{v=Math.round(v*1000)/10;if(v===0)v=0;return v.toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})+'%'};
const dec4=()=>Math.abs(P.cbs[2027]*10000-Math.round(P.cbs[2027]*10000))>1e-6||Math.abs(P.cbs[2027]*100*100-Math.round(P.cbs[2027]*100*100))>1e-6;
const rate=v=>{const f=dec4()?4:2;let x=v*100;const m=Math.pow(10,f),y=x*m,fl=Math.floor(y+1e-9),dd=y-fl;if(Math.abs(dd-0.5)<1e-7)x=((fl%2===0)?fl:fl+1)/m;return x.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:f})+'%'};
const mi=(v,d=1)=>fmt(v/1000,d);
const sgn=(v,t)=>v>1e-9?'sobe':v<-1e-9?'cai':'fica estável';
const aabs=v=>fmt(Math.abs(v));
function setP(pe,t){const sp=[...pe.querySelectorAll('span:not(.bu)')];if(sp[0]){const lit=/^•\s/.test(sp[0].textContent)&&!pe.querySelector('.bu');sp[0].textContent=(lit&&!/^•/.test(t)?'• ':'')+t;sp.slice(1).forEach(x=>x.textContent='')}}
function setTxt(el,txt,color){const sp=el.querySelector('span');if(!sp)return;sp.textContent=txt;[...el.querySelectorAll('span')].slice(1).forEach(s=>s.textContent='');if(color)sp.style.color=color}
function allEls(){return [...document.querySelectorAll('#cv .sh.hast')].map(e=>({e,x:+e.dataset.x,y:+e.dataset.y,w:+e.dataset.w,t:e.textContent.trim()}))}
// ---- slides de DRE (5 a 14)
const BL=[138.95,138.95+184.8,138.95+2*184.8,138.95+3*184.8,138.95+4*184.8,138.95+5*184.8],CW=44.36;
const FLD={'Receita Bruta':'RB','(-) Impostos':'DED','Receita Líquida':'RL','OPEX':'OPEX','Lucro Bruto':'LB','LAIR':'EBT','(-) IR':'IR','Lucro Líquido':'LL','EBITDA':'EBITDA','(+) Créd. Tributários':'CRED'};
function valDRE(r,lab){const a=r.atual,t=r.rt;if(lab==='Resultado Financeiro')return[a.FIN,t.FIN];if(lab==='EBIT')return[a.EBITX,t.EBITX];if(lab.startsWith('↳'))return[a.CT,t.CT];const f=FLD[lab];return[a[f],t[f]]}
function arrow(a,t){if(Math.abs(a)<0.5&&Math.abs(t)<0.5)return'= 0%';if(Math.abs(t-a)<1e-6)return'= 0%';if(Math.abs(a)<0.5)return t>0?'▲ novo':'▼ novo';const v=(t-a)/Math.abs(a);const s=Math.round(Math.abs(v)*100)+'%';return(v>0?'▲ ':'▼ ')+s}
const fz=(v,dd)=>Math.abs(v)<0.5*Math.pow(10,-dd)?'0':fmt(v,dd);

const RDP={graos:{f:'Receita de Pedágio',c:'Receita de Contraprestação',l:'Receita Acessória'},mobi:{f:'Receita Bruta',c:'Receita de Contraprestação',l:'Receita Acessória / Locação'},leste:{f:'Receita Bruta',c:'Receita de Contraprestação',l:'Receita Acessória / Locação'},mercosul:{f:'Receita de Pedágio e CUF',c:'Receita de Contraprestação',l:'Receita Acessória'},norte:{f:'Receita Bruta',c:'Receita de Operação e Financeira IFRIC 12',l:'Receita Acessória'},sul:{f:'Receita Bruta',c:'Receita de Operação e Financeira IFRIC 12',l:'Receita Acessória'},mcp01:{f:'Receita Bruta',c:'Receita Bruta (contraprestação)',l:'Receita Acessória'},welfare:{f:'Receita Bruta',c:'Receita de Contraprestação',l:'Receita Acessória'},rota:{f:'Receita de Pedágio',c:'Receita de Contraprestação',l:'Receita Acessória'},lote6:{f:'Receita de Pedágio',c:'Receita de Contraprestação',l:'Receita Acessória'}};
const RDZ={graos:'Remuneração do Ativo Financeiro e Construção',mobi:'Remuneração do Ativo Financeiro',mercosul:'Receitas fora da base do CBS'};

function debLinhas(k){
  const rd=(P.redutor[2027]||0),rl=(P.redutorLoc[2027]||0),r2=v=>(v*100).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})+'%',full=P.cbs[2027],mv=v=>' (R$ '+fmt(v,0)+' mil)';
  const R=B.empresas[k].years[2027].comp.RB,cp=R.shareCP||0,lo=R.shareLoc||0,den=-R.cbs/0.0925,r=CY(k,2027),rb=r?r.atual.RB:den;
  const vf=den*Math.max(0,1-cp-lo),vc=den*cp,vl=den*lo;let vz=rb-den;if(vz<0.5)vz=0;
  const t=RDP[k],ln=[],e=0.0005*rb;
  if(vf>e)ln.push([t.f+mv(vf),rate(full)]);
  if(vc>e)ln.push([t.c+mv(vc),r2(full*(1-rd))]);
  if(vl>e)ln.push([t.l+mv(vl),r2(full*(1-rl))]);
  if(vz>0)ln.push([(RDZ[k]||'Receitas fora da base do CBS')+mv(vz),r2(0)]);
  return {ln,rb};
}
function rodapeDeb(k){
  const cv=document.getElementById('cv');if(!cv||(k&&!RDP[k]))return;
  const fe=allEls().find(o=>o.e.dataset.fe||o.t.startsWith('Fonte: planilhas das empresas'));if(!fe)return;fe.e.dataset.fe=1;
  const rd=(P.redutor[2027]||0),rl=(P.redutorLoc[2027]||0),r2=v=>(v*100).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})+'%',full=P.cbs[2027],mv=v=>' (R$ '+fmt(v,0)+' mil)';
  let vf=0,vc=0,vl=0,vz=0,rbT=0;
  (k?[k]:ORS()).forEach(q=>{const R=B.empresas[q].years[2027].comp.RB,cp=R.shareCP||0,lo=R.shareLoc||0,den=-R.cbs/0.0925,r=CY(q,2027),rb=r?r.atual.RB:den;vf+=den*Math.max(0,1-cp-lo);vc+=den*cp;vl+=den*lo;vz+=rb-den;rbT+=rb});
  if(vz<0.5)vz=0;
  const t=k?RDP[k]:{f:'Receita sem redução (Pedágio, Receita Bruta)',c:'Receita de Contraprestação',l:'Receita Acessória / Locação',z:'Receitas fora da base do CBS'},ln=[];if(k)t.z=RDZ[k]||'Receitas fora da base do CBS';
  if(vf>0.0005*(vf+vc+vl)+1e-9)ln.push([t.f+mv(vf),rate(full)]);
  if(vc>0.0005*(vf+vc+vl))ln.push([t.c+mv(vc),r2(full*(1-rd))]);
  if(vl>0.0005*(vf+vc+vl))ln.push([t.l+mv(vl),r2(full*(1-rl))]);
  if(vz>0)ln.push([t.z+mv(vz),r2(0)]);
  const tx=['Débito 2027 (total = Receita Bruta de R$ '+fmt(rbT,0)+' mil):'].concat(ln.map(a=>a[0]+'. Utilizada a alíquota de '+a[1]));
  const pps=[...fe.e.querySelectorAll('.pp')];while(pps.length<tx.length){const c=pps[0].cloneNode(true);pps[pps.length-1].after(c);pps.push(c)}
  pps.forEach((pe,i)=>{if(i<tx.length){setP(pe,tx[i]);pe.style.cssText+=';display:block;text-align:left;white-space:nowrap;line-height:1.1;font-size:9px;'+(i===0?'font-weight:700;':'')}else pe.remove()});fe.e.style.height='auto';fe.e.style.display='block';fe.e.style.top='659px';
}
function dynDRE(si){
  const cons=si===14,k=cons?null:(si===21?'lote6':OR9[si-5]),d=0,sc=1;
  const els=allEls();const labs=els.filter(o=>Math.abs(o.x-34.56)<1&&o.t).sort((a,b)=>a.y-b.y);
  const eff={};let prev=null;labs.forEach(o=>{if(o.t==='% Margem')eff[Math.round(o.y)]=['M',prev];else{eff[Math.round(o.y)]=['V',o.t];prev=o.t}});
  els.forEach(o=>{const ey=eff[Math.round(o.y)];if(!ey||o.x<100)return;
    for(let bi=0;bi<6;bi++){const ci=Math.round((o.x-BL[bi])/CW);if(ci<0||ci>3||Math.abs(o.x-(BL[bi]+ci*CW))>2)continue;
      const y=2026+bi;const r=cons?CONS(y):CY(k,y);
      const[kind,lab]=ey;if(kind==='V'&&!(lab in FLD)&&lab!=='Resultado Financeiro'&&lab!=='EBIT'&&!lab.startsWith('↳'))return;
      if(!r){if(o.t!=='')setTxt(o.e,'—');return}if(kind==='V'&&!(lab in FLD)&&lab!=='Resultado Financeiro'&&lab!=='EBIT'&&!lab.startsWith('↳'))return;
      if(cons&&ci<3)o.e.querySelectorAll('span').forEach(s=>s.style.fontSize='9.3px');
      if(kind==='V'){let[a,t]=valDRE(r,lab);a/=sc;t/=sc;const dl=t-a;
        if(lab.startsWith('↳')){if(ci===0)setTxt(o.e,fz(a,d));else if(ci===1)setTxt(o.e,fz(t,d));else if(ci===2)setTxt(o.e,fz(dl,d),dl<-1e-9?RED:dl>1e-9?GRN:null);else setTxt(o.e,arrow(a,t),(t-a)<-1e-9?RED:(t-a)>1e-9?GRN:null);return}
        if(ci===0)setTxt(o.e,fz(a,d));else if(ci===1)setTxt(o.e,fz(t,d));
        else if(ci===2)setTxt(o.e,fz(dl,d),dl<-1e-9?RED:dl>1e-9?GRN:null);
        else setTxt(o.e,Math.abs(a)<1e-9?'–':p1(dl/Math.abs(a)),dl<-1e-9?RED:dl>1e-9?GRN:null);return}
      const base={'Lucro Bruto':'LB','EBITDA':'EBITDA','Lucro Líquido':'LL','EBIT':'EBITX','LAIR':'EBT'}[lab];if(!base)return;
      const ma=r.atual.RL?r.atual[base]/r.atual.RL:0,mt=r.rt.RL?r.rt[base]/r.rt.RL:0;
      if(ci===0)setTxt(o.e,p1(ma));else if(ci===1)setTxt(o.e,p1(mt));else if(ci===2)setTxt(o.e,pp(mt-ma),mt-ma>=0?GRN:RED);return}
  });
  // rodapé: Débito 2027 (alíquota por tipo de receita), substitui os quadros de alíquotas por ano
  rodapeDeb(cons?null:k);
  // quadros de alíquotas (rodapé) ou bullets de impacto
  if(!cons){document.querySelectorAll('#cv table.tb').forEach(tb=>{const bi=Math.round((parseFloat(tb.style.left)-141)/184.8);if(bi<0||bi>5)return;const y=2026+bi;
    const rows=[...tb.rows];let sum=0,anyTot=null;
    rows.forEach(tr=>{const c=tr.cells;if(c.length<3)return;const lab=c[0].textContent.trim();
      const set=v=>{const sp=c[2].querySelector('span');if(sp)sp.textContent=v};
      if(lab==='CBS'){set(rate(P.cbs[y]));sum+=P.cbs[y]}
      else if(lab==='IBS'){const v=P.ibsRef[y]*P.ibsFase[y];set(rate(v));sum+=v}
      else if(lab==='ISS'){const a=parseFloat(c[1].textContent.replace('%','').replace(',','.'))/100;const v=(isNaN(a)?0:a)*P.issFase[y];set(rate(v));sum+=v}
      else if(/^PIS\/COF/.test(lab)||/^Total/.test(lab)){if(/^Total/.test(lab))anyTot=c[2];else{const t=c[2].textContent.trim();const n=parseFloat(t.replace('%','').replace(',','.'));if(!isNaN(n)&&t!=='–')sum+=n/100}}});
    if(anyTot){const sp=anyTot.querySelector('span');if(sp)sp.textContent=rate(sum)}
  })}else{
    els.filter(o=>o.t.startsWith('• Δ Receita Bruta')).forEach(o=>{const bi=Math.round((o.x-141)/184.8);const y=2026+bi;const r=CONS(y);if(!r)return;
      const f=(lab,k2)=>{const a=r.atual[k2],t=r.rt[k2],dl=t-a;return `• Δ ${lab}: ${fz(dl,0)} mil (${p1(dl/Math.abs(a))})`};
      const ps=o.e.querySelectorAll('.pp');const arr=[f('Receita Bruta','RB'),f('EBITDA','EBITDA'),f('Lucro Líquido','LL')];ps.forEach((pe,i)=>setP(pe,arr[i].replace(/^• /,'')))})
  }
}
// ---- helpers de tabela
function tbl(n){return document.querySelector(`#cv table.tb[data-n="${n}"]`)}
function cell(tb,r,c,txt,color){const row=tb.rows[r];if(!row)return;let ci=0,td=null;for(const t of row.cells){const span=t.colSpan||1;if(c>=ci&&c<ci+span){td=t;break}ci+=span}if(!td)return;const sp=td.querySelector('span');if(!sp)return;sp.textContent=txt;if(color)sp.style.color=color}
const colS=(v)=>v<-1e-9?RED:v>1e-9?GRN:null;
function nsPct(a,t){if(Math.abs(a)<1e-9)return'–';if((a>0&&t<0)||(a<0&&t>0))return'n.s.';return p1((t-a)/Math.abs(a))}
function fitBelow(tb,noteN,srcN){const cv=document.getElementById('cv'),sc=cv.getBoundingClientRect().width/1280;const bot=(tb.getBoundingClientRect().bottom-cv.getBoundingClientRect().top)/sc;
  const nt=document.querySelector(`#cv .sh[data-n="${noteN}"]`);if(!nt)return;const h=nt.getBoundingClientRect().height/sc;const top=Math.max(parseFloat(nt.style.top),bot+6);nt.style.top=top+'px';
  const sr=document.querySelector(`#cv .sh[data-n="${srcN}"]`);if(sr){const need=top+h+2;if(parseFloat(sr.style.top)<need)sr.style.top=Math.min(need,684)+'px'}}
// ---- slide 15: fluxo mensal
function dynFlux(){const f=FLC(),r=f.rt;const tb=tbl('Table 516');if(!tb)return;grow(tbl('Table 517'),9);
  const rowsK=[['deb',2],['opex',3],['capex',4],['mov',5],['des',6],['res',7]];
  for(let m=0;m<12;m++)cell(tb,1,1+m,'0');cell(tb,1,13,'0');
  rowsK.forEach(([k,ri])=>{r[k].forEach((v,m)=>cell(tb,ri,1+m,fmt(v)));cell(tb,ri,13,fmt(sum(r[k])))});
  const t2=tbl('Table 517');
  ORS().forEach((k,i)=>{cell(t2,1+i,0,NAMES[k]);const x=FL(k).rt;['deb','opex','capex','mov','des','res'].forEach((c,j)=>cell(t2,1+i,1+j,fmt(sum(x[c]))))});
  ['deb','opex','capex','mov','des','res'].forEach((c,j)=>cell(t2,1+ORS().length,1+j,fmt(sum(r[c]))));
  if(L6)fitBelow(t2,'TextBox 515','TextBox 507');
  const ch=document.querySelector('#cv .chart[data-n="Chart 514"]');
  if(ch)ch.innerHTML=chartBars(ch,{title:'Movimentação mensal consolidada (R$ mil)',cats:MESES,series:[{name:'RT',v:r.mov,color:'#2F75B5'}],fmtv:v=>fmt(v)});
}
function chartBars(ch,o){const W=parseFloat(ch.style.width),H=parseFloat(ch.style.height);
  const pad={l:50,r:10,t:o.title?26:10,b:o.legend?44:26};const iw=W-pad.l-pad.r,ih=H-pad.t-pad.b;
  const all=o.series.flatMap(s=>s.v).concat([0]);let mn=Math.min(...all),mx=Math.max(...all);if(mx===mn)mx=mn+1;
  const nt=(()=>{const rg=mx-mn;const raw=rg/5;const mag=Math.pow(10,Math.floor(Math.log10(raw)));const n=raw/mag;return(n<1.5?1:n<3?2:n<7?5:10)*mag})();
  mn=Math.floor(mn/nt)*nt;mx=Math.ceil(mx/nt)*nt;
  const y=v=>pad.t+ih*(1-(v-mn)/(mx-mn));const ns=o.series.length,cw=iw/o.cats.length,gw=cw*0.2,bw=(cw-gw)/ns;
  let s=`<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" style="font-family:${FONT}">`;
  if(o.title)s+=`<text x="${W/2}" y="14" text-anchor="middle" font-size="11" font-weight="700" fill="#293746">${o.title}</text>`;
  for(let v=mn;v<=mx+1e-9;v+=nt){s+=`<line x1="${pad.l}" x2="${W-pad.r}" y1="${y(v)}" y2="${y(v)}" stroke="#E3E8EE"/><text x="${pad.l-6}" y="${y(v)+3}" text-anchor="end" font-size="9" fill="#5B6573">${(o.axf||(x=>fmt(x,Math.abs(nt)<1?1:0)))(v)}</text>`}
  o.cats.forEach((c,i)=>{o.series.forEach((se,j)=>{const v=se.v[i],x=pad.l+i*cw+gw/2+j*bw,y0=y(0),y1=y(v);s+=`<rect x="${x}" y="${Math.min(y0,y1)}" width="${bw}" height="${Math.max(Math.abs(y1-y0),.5)}" fill="${se.color}"/>`;
    if(o.labels)s+=`<text x="${x+bw/2}" y="${v>=0?y1-3:y1+10}" text-anchor="middle" font-size="9" fill="#293746">${o.fmtv?o.fmtv(v):v}</text>`});
    s+=`<text x="${pad.l+i*cw+cw/2}" y="${H-pad.b+(o.rot?12:13)}" text-anchor="middle" font-size="9" fill="#5B6573" ${o.rot?`transform="rotate(-35 ${pad.l+i*cw+cw/2} ${H-pad.b+12})"`:''}>${c}</text>`});
  s+=`<line x1="${pad.l}" x2="${W-pad.r}" y1="${y(0)}" y2="${y(0)}" stroke="#8A98A8"/>`;
  if(o.legend)o.series.forEach((se,j)=>{const lx=W/2-o.series.length*40+j*90;s+=`<rect x="${lx}" y="${H-12}" width="9" height="9" fill="${se.color}"/><text x="${lx+13}" y="${H-4}" font-size="9" fill="#293746">${se.name}</text>`});
  return s+'</svg>'}
// ---- slide 16: movimentação atual x RT
function dynMov(){const f=FLC(),r=f.rt,a=f.at;const tb=tbl('Table 511');if(!tb)return;
  const K=['deb','opex','capex','mov','des','res'];
  const A=K.map(k=>a[k]),T=K.map(k=>sum(r[k]));
  const d1=v=>Math.abs(v)<0.5?'–':fmt(v);
  K.forEach((k,i)=>{cell(tb,1+i,1,d1(A[i]));cell(tb,1+i,2,d1(T[i]));cell(tb,1+i,3,fmt(T[i]-A[i]),colS(T[i]-A[i]))});
  const ch=document.querySelector('#cv .chart[data-n="Chart 513"]');
  if(ch)ch.innerHTML=chartBars(ch,{title:'Atual × RT (R$ mil)',cats:['Débitos','Créditos OPEX','Créditos CAPEX','Movimentação','Desembolso','Restituição'],series:[{name:'Modelo atual',v:A.map(v=>Math.round(v)),color:'#8A98A8'},{name:'RT',v:T.map(v=>Math.round(v)),color:'#2F75B5'}],labels:1,fmtv:v=>fmt(v),legend:1,rot:1,axf:v=>fmt(v)});
  const be=allEls().find(o=>o.t.startsWith('• Modelo atual'));
  if(be){const ps=be.e.querySelectorAll('.pp');const m=(x)=>(x<-0.5?'-':'')+fmt(Math.abs(x));
   const big=T[2]>T[1]?'créditos de CAPEX':'créditos de OPEX/SG&A';
   const txt=[`Modelo atual: débitos de ${m(A[0])} mil e créditos de ${m(A[1])} mil resultam em movimentação de ${m(A[3])} mil, integralmente desembolso de caixa.`,
    `Reforma Tributária: débitos de ${m(T[0])} mil, compensados por créditos de OPEX/SG&A (${m(T[1])} mil) e de CAPEX (${m(T[2])} mil), geram movimentação líquida de ${m(T[3])} mil.`,
    `Impacto no caixa (por CNPJ): desembolso de ${m(T[4])} mil e restituição de ${m(T[5])} mil. ${T[3]>0?'A movimentação líquida positiva não elimina o desembolso, porque o crédito de uma empresa não paga o débito de outra.':'Créditos de uma empresa não pagam o débito de outra.'}`,
    `Variação da movimentação: ${m(T[3]-A[3])} mil, puxada pelos ${big}.`];
   ps.forEach((pe,i)=>{if(txt[i])setP(pe,txt[i])})}
}
// ---- slides 17-19: comparativos
function rowsCmp(y){return ORS().map(k=>({k,r:CY(k,y)}))}
function totCmp(y){return CONS(y)}
function dynCmp(si){
  const y=si===17?2026:2027;const tb=tbl('Table 511');if(!tb)return;grow(tb,10);
  const items=ORS().map(k=>({n:NAMES[k],r:CY(k,y)}));const T=CONS(y);items.push({n:'Total 9 empresas',r:T,tot:1});
  items.forEach((it,i)=>{const ri=2+i,r=it.r;cell(tb,ri,0,it.n);
    if(!r){for(let c=1;c<15;c++)cell(tb,ri,c,'—');return}
    const A=r.atual,R=r.rt;let c=1;
    const put=(txt,col)=>{cell(tb,ri,c++,txt,col)};
    const pctc=(a,t)=>{const s=nsPct(a,t);put(s,s==='n.s.'||s==='–'?null:colS(t-a))};
    const m=(a,t,dl)=>{put(fmt(a));put(fmt(t))};
    if(y===2027){m(A.RB,R.RB);pctc(A.RB,R.RB)}
    m(A.RL,R.RL);pctc(A.RL,R.RL);
    put(fmt(A.CT));put(fmt(R.CT));put(arrow(A.CT,R.CT),colS(R.CT-A.CT));
    put(fmt(A.EBITDA));put(fmt(R.EBITDA));put(fmt(R.EBITDA-A.EBITDA),colS(R.EBITDA-A.EBITDA));pctc(A.EBITDA,R.EBITDA);
    put(fmt(A.LL));put(fmt(R.LL));put(fmt(R.LL-A.LL),colS(R.LL-A.LL));pctc(A.LL,R.LL);
  });
  // bullets
  const be=allEls().find(o=>o.t.startsWith('• Em 20'));if(!be)return;
  const ps=be.e.querySelectorAll('.pp'),set=(i,t)=>{if(ps[i])setP(ps[i],t)};
  const dLL=T.rt.LL-T.atual.LL,dE=T.rt.EBITDA-T.atual.EBITDA;
  const rk=ORS().map(k=>({k,d:CY(k,y)?CY(k,y).rt.LL-CY(k,y).atual.LL:0})).sort((a,b)=>a.d-b.d);
  const nl=(y===2027&&L6)?3:2,l2=rk.slice(0,nl),s2=l2.reduce((a,o)=>a+o.d,0),cbs=rate(P.cbs[y]);const nw=nl===3?'três':'duas';
  const qtd=ORS().filter(k=>CY(k,y)).length;const x6=f=>sum(ORS().filter(k=>k!=='lote6'&&CY(k,y)).map(k=>CY(k,y).rt[f]-CY(k,y).atual[f]));
  const ct0=T.atual.CT,ct1=T.rt.CT;
  set(0,`Em ${y}, o lucro líquido somado ${y===2026?'das 9 empresas ':''}${dLL<0?'cai':'sobe'} ${aabs(dLL)} mil (${p1(dLL/Math.abs(T.atual.LL))}) no RT, e o EBITDA ${dE<0?'cai':'sobe'} ${aabs(dE)} mil (${p1(dE/Math.abs(T.atual.EBITDA))}).${y===2027&&L6?` Sem a Lote 6 as quedas eram de ${aabs(x6('LL'))} mil e ${aabs(x6('EBITDA'))} mil.`:''}${y===2027?' É o primeiro ano com a CBS plena e o PIS/COFINS extinto.':''}`);
  if(y===2026){
    set(1,`${NAMES[l2[0].k]} e ${NAMES[l2[1].k]} ${s2<0?'perdem':'variam'} ${aabs(s2)} mil${s2<dLL?', mais que a perda total':''}: ${l2.every(o=>['graos','rota'].includes(o.k))?`as duas saem do PIS/COFINS cumulativo (3,65%) para CBS de ${cbs}`:'maior pressão da carga tributária no RT'}.`);
    set(2,`O crédito tributário ${ct1>=ct0?'sobe':'cai'} de ${aabs(ct0)} mil para ${aabs(ct1)} mil (${arrow(ct0,ct1)}), mas não compensa a alíquota maior nessas duas operações.`);
    const oth=ORS().filter(k=>CY(k,y)&&!l2.some(o=>o.k===k));const so=sum(oth.map(k=>CY(k,y).rt.LL-CY(k,y).atual.LL));
    set(3,`As outras ${oth.length} empresas com operação ${so>=0?'melhoram levemente (+':'pioram ('}${aabs(so)} mil somados), pelo crédito sobre OPEX. MCP01 não opera em 2026.`);
  }else{
    const dRB=T.rt.RB-T.atual.RB,dRL=T.rt.RL-T.atual.RL;
    set(1,`A receita bruta ${dRB<0?'cai':'sobe'} ${aabs(dRB)} mil (${p1(dRB/Math.abs(T.atual.RB))}), mas a receita líquida ${dRL<0?'cai':'sobe'} só ${aabs(dRL)} mil (${p1(dRL/Math.abs(T.atual.RL))}): a maior parte da queda na receita bruta é a CBS saindo do preço ("por fora").`);
    const lst=l2.slice(0,-1).map(o=>`${NAMES[o.k]} (–${aabs(o.d)} mil)`).join(', ')+` e ${NAMES[l2[nl-1].k]} (–${aabs(l2[nl-1].d)} mil)`;
    set(2,`${lst} perdem juntas ${aabs(s2)} mil${s2<dLL?', mais que a perda total':''}: ${l2.every(o=>['graos','rota','lote6'].includes(o.k))?`as ${nw} saem do PIS/COFINS cumulativo (3,65%) para CBS de ${cbs}.`:'maior pressão da carga tributária no RT.'}`);
    set(3,`O crédito tributário ${ct1>=ct0?'sobe':'cai'} de ${aabs(ct0)} mil para ${aabs(ct1)} mil (${arrow(ct0,ct1)}), inclusive em Grãos${L6?', Rota e Lote 6':' e Rota'}, que hoje não têm crédito; ainda assim, o ganho não compensa o aumento de alíquota nessas ${nw} operações.`);
  }
}
function dynFluxEmp(){const tb=tbl('Table 511');if(!tb)return;grow(tb,10);
  const items=ORS().map(k=>({n:NAMES[k],f:FL(k)}));const fc=FLC();
  const tot={at:fc.at,rt:fc.rt};
  const rows=items.map(it=>({a:it.f.at,r:{deb:sum(it.f.rt.deb),opex:sum(it.f.rt.opex),capex:sum(it.f.rt.capex),mov:sum(it.f.rt.mov),des:sum(it.f.rt.des),res:sum(it.f.rt.res)}}));
  rows.push({a:fc.at,r:{deb:sum(fc.rt.deb),opex:sum(fc.rt.opex),capex:sum(fc.rt.capex),mov:sum(fc.rt.mov),des:sum(fc.rt.des),res:sum(fc.rt.res)}});
  rows.forEach((o,i)=>{const ri=2+i,a=o.a,r=o.r;cell(tb,ri,0,i<items.length?items[i].n:'Total '+items.length+' empresas');let c=1;const put=(t,col)=>cell(tb,ri,c++,t,col);
    put(fmt(a.deb));put(fmt(r.deb));put(fmt(r.deb-a.deb),colS(r.deb-a.deb));
    put(fmt(a.opex));put(fmt(r.opex));put(fmt(r.capex));
    put(fmt(a.mov));put(fmt(r.mov));put(fmt(r.mov-a.mov),colS(r.mov-a.mov));
    put(fmt(a.des));put(fmt(r.des));put(fmt(r.res));put(fmt(r.des-a.des),colS(r.des-a.des))});
  const be=allEls().find(o=>o.t.startsWith('• Modelo atual'));if(!be)return;
  const ps=be.e.querySelectorAll('.pp'),set=(i,t)=>{if(ps[i])setP(ps[i],t)};
  const m=x=>(x<-0.5?'-':'')+fmt(Math.abs(x)),T=rows[rows.length-1];
  const ndes=rows.slice(0,ORS().length).filter(o=>o.r.des<-0.5).length,nres=rows.slice(0,ORS().length).filter(o=>o.r.res>0.5).length;
  const both=ORS().filter((k,i)=>rows[i].r.des<-0.5&&rows[i].r.res>0.5).map(k=>NAMES[k]);set(0,`Modelo atual: as ${ORS().length} empresas desembolsam juntas ${m(T.a.des)} mil. No RT, o desembolso soma ${m(T.r.des)} mil em ${ndes} empresas e a restituição soma ${m(T.r.res)} mil em ${nres}${both.length?` (${both.join(' e ')} ${both.length>1?'têm':'tem'} os dois ao longo do ano)`:''}. A movimentação líquida de ${m(T.r.mov)} mil esconde essa assimetria entre CNPJs.`);
  const gl=rows.slice(0,ORS().length).map((o,i)=>({i,d:o.r.des-o.a.des}));const gi=gl.slice().sort((a,b)=>a.d-b.d)[0].d<-0.5?gl.slice().sort((a,b)=>a.d-b.d)[0]:gl.slice().sort((a,b)=>Math.abs(b.d)-Math.abs(a.d))[0],go_=rows[gi.i];
  set(1,`${NAMES[ORS()[gi.i]]} concentra o ${gi.d<0?'aumento':'alívio'} de saída de caixa: o desembolso vai de ${m(go_.a.des)} mil para ${m(go_.r.des)} mil, ${m(Math.abs(gi.d))} mil ${gi.d<0?'a mais':'a menos'}${go_.r.capex>0.5?`, mesmo com ${m(go_.r.capex)} mil de crédito de CAPEX no fluxo de 2027`:''}.`);
  const rs=ORS().map((k,i)=>({k,o:rows[i]})).filter(x=>x.o.r.res>0.5&&x.o.r.des>-0.5);
  const nm=rs.map(x=>NAMES[x.k]);const lst=nm.length>1?nm.slice(0,-1).join(', ')+' e '+nm[nm.length-1]:(nm[0]||'Nenhuma empresa');
  set(2,rs.length?`${lst} ${rs.length>1?'passam':'passa'} a ter restituição líquida (${m(sum(rs.map(x=>x.o.r.res)))} mil), puxada pelos créditos de CAPEX (${m(sum(rs.map(x=>x.o.r.capex)))} mil). Esse caixa depende do prazo de ressarcimento.`:'Nenhuma empresa fica com restituição líquida. O caixa de créditos depende do prazo de ressarcimento.');
}
// ---- slides 3 e 4: textos que dependem dos parâmetros
function rep(el,fn){el.querySelectorAll('span').forEach(s=>{const n=fn(s.textContent);if(n!==s.textContent)s.textContent=n})}
function dynS3(){const p2=v=>rate(v).replace('%','');
  const ib=y=>P.ibsRef[y]*P.ibsFase[y],iss=y=>0.05*P.issFase[y];
  const il=[2029,2030,2031].map(y=>rate(ib(y))).join(' / ').replace(/%/g,(m,i,s)=>m);
  const sl=[2029,2030,2031].map(y=>(iss(y)*100).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:2})+'%').join(' / ');
  allEls().forEach(o=>{
    if(o.t.includes('modelo: CBS'))rep(o.e,s=>s.replace(/modelo: CBS .*$/,`modelo: CBS ${rate(P.cbs[2027])}`+(Math.abs(P.cbs[2027]-P.cbs[2028])>1e-9?` / ${rate(P.cbs[2028])}`:'')));
    if(o.t.includes('Modelo 2029–31'))rep(o.e,s=>s.startsWith('Modelo')?`Modelo 2029–31: IBS ${il} e ISS ${sl}`:s);
    if(o.t.includes('No modelo: CBS de'))rep(o.e,s=>s.replace(/CBS de [\d,]+%/,`CBS de ${rate(P.cbs[2027])}`));
    if(o.t.includes('matriz de custeio')||o.t.includes('Crédito sobre')){const cr=ORDER.map(k=>{const E=P.empresas[k]||{};return E.credito!=null?E.credito:B.empresas[k].credito0}),lo=Math.min(...cr),hi=Math.max(...cr),g=x=>(x*100).toLocaleString('pt-BR',{maximumFractionDigits:1});
      rep(o.e,s=>s.replace(/^15% da matriz de custeio.*$/,`a parcela das compras de fornecedores do Simples Nacional não gera crédito; o % no Regime Normal vem da planilha Regime Tributário (Assertif), por empresa (de ${g(lo)}% a ${g(hi)}%, ver slide 4).`).replace(/Crédito sobre 85% dos custos e despesas elegíveis \(regime normal\)\./,`Crédito sobre ${g(lo)}% a ${g(hi)}% dos custos e despesas elegíveis, conforme o Regime Normal de cada empresa.`))}
  })}
function dynS4(){
  const FX=[['Fonte: planilha Aliquotas_Fluxo (aba Fluxo) e DF - Mensal das empresas, 2027.','Fonte: DF - Mensal 2027 de cada empresa (planilhas de origem), recalculada com as premissas da simulação.'],
   ['Créditos de CAPEX conforme a planilha (zerados no fluxo mensal de Grãos do Piauí, MCP01 e Rota da Integração).','Créditos de CAPEX conforme a DF - Mensal de cada empresa (planilhas de origem).'],
   ['RT: aba Fluxo 2027,','RT: fluxo 2027 (DF - Mensal),'],['RT: planilha Fluxo 2027,','RT: fluxo 2027 (DF - Mensal),']];
  allEls().forEach(o=>rep(o.e,t=>{let r=t;FX.forEach(([a,b])=>{if(r.includes(a))r=r.replace(a,b)});return r}));
  allEls().forEach(o=>{if(o.t.startsWith('Fonte: planilha Aliquotas_Fluxo'))rep(o.e,s=>s.replace('(aba Premissas).','(aba Premissas); % de crédito (Regime Normal) da planilha Regime Tributário (Assertif).'))});
  allEls().forEach(o=>{if(!o.t.includes('IR/CS'))return;const bx=o.e;
    // descobre a empresa pelo título da caixa acima (mesmo x, y-39)
    const ti=allEls().find(q=>Math.abs(q.x-o.x)<1&&Math.abs(q.y-(o.y-39))<3);if(!ti)return;
    const k=OR9.find(kk=>NAMES[kk]===ti.t);if(!k)return;const E=P.empresas[k]||{};const c=B.empresas[k],y0=Object.keys(c.years)[0];
    const cr=E.credito!=null?E.credito:null,ir=E.ir!=null?E.ir:null,g=x=>(x*100).toLocaleString('pt-BR',{maximumFractionDigits:2});
    bx.querySelectorAll('.pp').forEach(pe=>{const sp=pe.querySelector('span');if(!sp)return;
      if(cr!=null&&/Aplicado \d+%/.test(sp.textContent))sp.textContent=sp.textContent.replace(/Aplicado [\d,]+%/,'Aplicado '+g(cr)+'%');
      if(ir!=null&&/IR\/CS/.test(sp.textContent))sp.textContent=sp.textContent.replace(/aplicado [\d,]+%/,'aplicado '+g(ir)+'%')})})}
function dynamic(si){inv();try{
  if(si===3)dynS3();if(si===4)dynS4();
  if((si>=5&&si<=14)||si===21)dynDRE(si);
  if(si===15)dynFlux();if(si===16)dynMov();if(si===17||si===18)dynCmp(si);if(si===19)dynFluxEmp();
  if(si===20)dynRedutor();
 }catch(e){console.error('dyn',si,e)}}
