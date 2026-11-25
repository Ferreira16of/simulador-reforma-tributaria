// uso: node calc_cli.js bases.json motor.js [cenario.json|--padrao]  -> JSON {empresa:{ano:{atual,rt,det}}}
// --padrao = padrão da simulação (CBS 11,3468 / IBS ref 28-CBS / redutor 44,2 / locação 70 / crédito = Regime Normal)
const fs=require('fs');const B=JSON.parse(fs.readFileSync(process.argv[2],'utf8'));
const m=require(require('path').resolve(process.argv[3]));
let P=m.defaultParams();
if(process.argv[4]==='--padrao')P=m.simDefaults(B);
else if(process.argv[4]){const j=JSON.parse(fs.readFileSync(process.argv[4],'utf8'));const q=j.params||j;
 P=Object.assign(m.defaultParams(),q);P.audit=Object.assign(m.defaultParams().audit,q.audit||{});}
const out={};for(const k in B.empresas){out[k]={};for(const y in B.empresas[k].years){out[k][y]=m.compute(B.empresas[k],k,+y,P,B.audit);}}
console.log(JSON.stringify(out));
