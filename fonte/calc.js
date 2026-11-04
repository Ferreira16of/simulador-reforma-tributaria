
const B=/*BASES*/;const F=/*FLOW*/;
/*MOTOR*/
const ORDER=['graos','mobi','leste','mercosul','norte','sul','mcp01','welfare','rota','lote6'];
const NAMES={};ORDER.forEach(k=>NAMES[k]=B.empresas[k].nome);
const DEF=()=>simDefaults(B);let P=DEF(),L6=false;const D=DEF();
let cache={};const inv=()=>{cache={}};
const fmt=(v,d=0)=>{if(Math.abs(v)<0.5*Math.pow(10,-d))v=0;const s=Math.abs(v).toLocaleString('pt-BR',{minimumFractionDigits:d,maximumFractionDigits:d});return v<0?'('+s+')':s};
const pc=(v,d=1)=>{v=Math.round(v*100*Math.pow(10,d))/Math.pow(10,d);if(v===0)v=0;return v.toLocaleString('pt-BR',{minimumFractionDigits:d,maximumFractionDigits:d})+'%'};
const pcs=v=>pc(v);const pp=v=>{v=Math.round(v*1000)/10;if(v===0)v=0;return v.toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})+' p.p.'};
const cls=v=>v>1e-9?'pos':v<-1e-9?'neg':'';
// ---- dados por empresa/ano
function CY(k,y){const key=k+y;if(cache[key]!==undefined)return cache[key];const r=compute(B.empresas[k],k,y,P,B.audit);
 if(r){const cr=r.rt.OPEX-r.atual.OPEX;r.atual.CT=r.atual.CRED||0;r.rt.CT=cr;const eb=k==='lote6'?'EBIT':'EBITDA';r.atual.EBITX=r.atual[eb];r.rt.EBITX=r.rt[eb];}
 return cache[key]=r;}
function keysV(){return ORDER.filter(k=>L6||k!=='lote6')}
function CONS(y){const ks=keysV();let A=null,R=null,n=0;ks.forEach(k=>{const r=CY(k,y);if(!r)return;n++;if(!A){A={};R={};for(const f in r.atual){A[f]=0;R[f]=0}}for(const f in r.atual){A[f]+=r.atual[f];R[f]+=r.rt[f]}});return A?{atual:A,rt:R,n}:null}
function FL(k){const key='f'+k;if(cache[key]!==undefined)return cache[key];return cache[key]=flow(B.empresas[k],k,P,B.audit,F[k],D)}
const sum=a=>a.reduce((x,y)=>x+y,0);
function FLC(){const ks=keysV();const o={rt:{deb:Array(12).fill(0),opex:Array(12).fill(0),capex:Array(12).fill(0),mov:Array(12).fill(0),des:Array(12).fill(0),res:Array(12).fill(0)},at:{deb:0,opex:0,capex:0,mov:0,des:0,res:0}};
 ks.forEach(k=>{const f=FL(k);if(!f)return;for(const x in o.rt)f.rt[x].forEach((v,i)=>o.rt[x][i]+=v);for(const x in o.at)o.at[x]+=f.at[x]});return o}
const tot=(f,x)=>sum(f.rt[x]);