// uso: node flow_cli.js bases.json motor.js flow.json [cenario.json]  -> JSON {empresa:{rt:{deb,opex,capex,mov,des,res:[12]},at:{...}}}
const fs=require('fs');const B=JSON.parse(fs.readFileSync(process.argv[2],'utf8'));
const m=require(require('path').resolve(process.argv[3]));const F=JSON.parse(fs.readFileSync(process.argv[4],'utf8'));
let P=m.simDefaults(B);const D=m.simDefaults(B);
if(process.argv[5]){const j=JSON.parse(fs.readFileSync(process.argv[5],'utf8'));const q=j.params||j;
 P=Object.assign(m.simDefaults(B),q);P.audit=Object.assign(m.defaultParams().audit,q.audit||{});}
const out={};for(const k in B.empresas){const f=m.flow(B.empresas[k],k,P,B.audit,F[k],D);if(f)out[k]=f}
console.log(JSON.stringify(out));
