const fs=require('fs');const B=JSON.parse(fs.readFileSync('/tmp/w/red/bases_v7.json','utf8'));const m=require('/tmp/w/v16/motor.js');
const P=m.simDefaults(B),P0=m.simDefaults(B);m.YEARS.forEach(y=>{P0.redutor[y]=0;P0.redutorLoc[y]=0});
const out={};for(const k in B.empresas){const c=B.empresas[k];const r=m.compute(c,k,2027,P,B.audit),r0=m.compute(c,k,2027,P0,B.audit);const rb=c.years['2027'].comp.RB;
 out[k]={trt:r.det.cbsRB+r.det.ibsRB,tfull:r0.det.cbsRB+r0.det.ibsRB,den:rb.cbs?-rb.cbs/0.0925:rb.sem,dedA:r.atual.DED,dedR:r.rt.DED,rbA:r.atual.RB,rbR:r.rt.RB,fixed:r.det.fixed}}
console.log(JSON.stringify(out));
