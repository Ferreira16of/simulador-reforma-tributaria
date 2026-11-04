const CBS0=0.0925,IBS0=0.1875;
const ISS_PH0={2026:1,2027:1,2028:1,2029:.9,2030:.8,2031:.7,2032:.6,2033:0};
const IBS_PH0={2026:0,2027:0,2028:0,2029:.1,2030:.2,2031:.3,2032:.4,2033:1};
const CORE=['RB','OPEX','SGA','DA','NONOP'];
const YEARS=[2026,2027,2028,2029,2030,2031,2032,2033];
function defaultParams(){const P={cbs:{},ibsRef:{},ibsFase:{},issFase:{},redutor:{},redutorLoc:{},empresas:{},audit:{mercosulIss2026:false,mobiCredLegado:false,mobiDaCredito:false}};
 YEARS.forEach(y=>{P.cbs[y]=CBS0;P.ibsRef[y]=IBS0;P.ibsFase[y]=IBS_PH0[y];P.issFase[y]=ISS_PH0[y];P.redutor[y]=0;P.redutorLoc[y]=0;});return P;}
// Padrão da simulação: CBS 11,3468% + IBS ref. (28% − CBS), redutor 44,2% s/ contraprestação, redução 70% s/ locação, crédito = % Regime Normal por empresa
const SIM={cbs:0.113468,total:0.28,redutor:0.442,redutorLoc:0.70};
function simDefaults(B){const p=defaultParams();YEARS.forEach(y=>{p.cbs[y]=SIM.cbs;p.ibsRef[y]=SIM.total-SIM.cbs;p.redutor[y]=SIM.redutor;p.redutorLoc[y]=SIM.redutorLoc});
 if(B&&B.empresas)for(const k in B.empresas){const c=B.empresas[k].creditoRN;if(c!==undefined&&c!==null)p.empresas[k]={credito:c}}return p}
function dre(v,ir,irv){const o=Object.assign({},v);o.RL=o.RB+o.DED;o.LB=o.RL+o.OPEX+(o.CRED||0);o.EBITDA=o.LB+o.SGA;o.EBIT=o.EBITDA+o.DA;o.EBT=o.EBIT+o.FIN+o.NONOP;o.IR=(irv!==null&&irv!==undefined)?irv:-o.EBT*ir;o.LL=o.EBT+o.IR;return o;}
function compute(co,key,y,P,AUD){
 const Y=co.years[String(y)];if(!Y)return null;
 const E=(P.empresas&&P.empresas[key])||{};const c0=co.credito0;
 const c=E.credito!==undefined&&E.credito!==null?E.credito:c0;
 const ir=E.ir!==undefined&&E.ir!==null?E.ir:Y.ir_rate;const issf=E.issFator!==undefined&&E.issFator!==null?E.issFator:1;
 const cbsr=P.cbs[y]/CBS0,ibsr=(P.ibsRef[y]/IBS0)*P.ibsFase[y];const rt={};const det={cbsRB:0,ibsRB:0,credOpex:0,credOS:0,fixed:0};
 CORE.forEach(L=>{const cp=Y.comp[L];if(!cp){rt[L]=Y.rt[L];return;}
  const regC=cp.regC||0,regI=cp.regI||0;const fc=c0?1+regC*(c/c0-1):1,fi=c0?1+regI*(c/c0-1):1;
  let cbs=cp.cbs*cbsr*fc,ibs=cp.ibsUnit*ibsr*fi;
  if(L==='RB'){const rd=(P.redutor&&P.redutor[y])||0,rl=(P.redutorLoc&&P.redutorLoc[y])||0;const f=1-rd*(cp.shareCP||0)-rl*(cp.shareLoc||0);cbs*=f;ibs*=f;}rt[L]=cp.sem+cp.res+cbs+ibs;if(L==='RB'){det.cbsRB=cbs;det.ibsRB=ibs;}else if(L!=='NONOP'){det.credOpex+=cbs+ibs;if(L==='OPEX'||L==='SGA')det.credOS+=cbs+ibs;}});
 const d=Y.comp.DED;if(d)det.fixed=d.fixed||0;
 if(!d)rt.DED=Y.rt.DED;else{const ph=P.issFase[y],fx=d.fixed||0;rt.DED=fx+(ISS_PH0[y]>0?(d.def-fx)*(ph/ISS_PH0[y]):d.issFull*ph)*issf;}
 rt.CRED=0;rt.FIN=Y.rt.FIN;
 const A=P.audit||{};
 if(AUD&&A.mercosulIss2026&&key==='mercosul'&&y===2026)rt.DED+=AUD.mercosul.iss2026;
 if(AUD&&key==='mobi'){
  if(A.mobiCredLegado&&AUD.mobi.credLeg[y]!==undefined)rt.OPEX-=AUD.mobi.credLeg[y];
  if(A.mobiDaCredito&&AUD.mobi.daFix[y]){const f=AUD.mobi.daFix[y];rt.DA+=f.cbs*(P.cbs[y]/CBS0)+f.ibsUnit*((P.ibsRef[y]/IBS0)*P.ibsFase[y]);}
 }
 const a=Object.assign({},Y.atual);
 return {atual:dre(a,ir,Y.ir_atual_flat?null:a.IR),rt:dre(rt,ir,null),det};
}

// ---- Fluxo mensal 2027 (débitos/créditos por empresa). F = {deb:[12],opex:[12],capex:[12]} (valores base das planilhas)
const MESES=['Jan/27','Fev/27','Mar/27','Abr/27','Mai/27','Jun/27','Jul/27','Ago/27','Set/27','Out/27','Nov/27','Dez/27'];
function flow(co,key,P,AUD,F,Pdef){
 const r=compute(co,key,2027,P,AUD),r0=compute(co,key,2027,Pdef,null);if(!r)return null;
 const ann=x=>x.det.cbsRB+x.det.ibsRB+x.rt.DED-x.det.fixed;
 const debA=ann(r),debA0=ann(r0),opA=r.det.credOpex,opA0=r0.det.credOpex;
 const even=Array(12).fill(1/12);
 const w=(arr)=>{const t=arr?arr.reduce((a,b)=>a+b,0):0;return t?arr.map(v=>v/t):even;};
 const wd=w(F&&F.deb),wo=w(F&&F.opex);
 const cf=(P.cbs[2027]+P.ibsRef[2027]*P.ibsFase[2027])/CBS0;
 const fxA=r.det.fixed||0,wi=w(F&&F.iibb),deb=wd.map((x,i)=>x*debA+wi[i]*fxA),opex=wo.map(x=>x*opA);const Ey=(P.empresas&&P.empresas[key])||{},c0_=co.credito0,cc=(Ey.credito!==undefined&&Ey.credito!==null)?Ey.credito:c0_,y27=co.years['2027'],regO=(y27&&y27.comp&&y27.comp.OPEX&&y27.comp.OPEX.regC)||0,fcap=(c0_&&regO)?1+regO*(cc/c0_-1):1;const capex=(F&&F.capex?F.capex:Array(12).fill(0)).map(v=>v*cf*fcap);
 const mov=deb.map((v,i)=>v+opex[i]+capex[i]);
 const A=r.atual,aMov=A.DED+(A.CRED||0);
 return {rt:{deb,opex,capex,mov,des:mov.map(v=>Math.min(v,0)),res:mov.map(v=>Math.max(v,0))},
  at:{deb:A.DED,opex:A.CRED||0,capex:0,mov:aMov,des:Math.min(aMov,0),res:Math.max(aMov,0)}};
}
if(typeof module!=='undefined')module.exports={compute,defaultParams,simDefaults,SIM,YEARS,CORE,flow,MESES};
