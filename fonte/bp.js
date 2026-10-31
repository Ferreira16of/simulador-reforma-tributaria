// ---- parâmetros
const gv=(k,y,Q=P)=>k==='total'?Q.cbs[y]+Q.ibsRef[y]:Q[k][y];
function buildParams(){const rows=[['CBS (%)','cbs'],['IBS referência (%)','ibsRef'],['CBS + IBS referência (%)','total'],['Fase IBS (%)','ibsFase'],['Fase ISS (%)','issFase'],['Redutor s/ contraprestação (%)','redutor']];
 let h='<table><thead><tr><th></th>'+YEARS.map(y=>`<th>${y}</th>`).join('')+'</tr></thead><tbody>';
 rows.forEach(([l,k])=>{h+=`<tr><td>${l}</td>`+YEARS.map(y=>`<td><input type=number step=0.25 data-k="${k}" data-y="${y}" value="${+(gv(k,y)*100).toFixed(4)}" class="${Math.abs(gv(k,y)-gv(k,y,D))>1e-9?'chg':''}"></td>`).join('')+'</tr>'});
 $('par').innerHTML=h+'</tbody></table>';
 $('par').querySelectorAll('input').forEach(i=>i.oninput=()=>{const v=parseFloat(i.value);if(isNaN(v))return;const k=i.dataset.k,y=i.dataset.y;
  if(k==='cbs'){const tot=P.cbs[y]+P.ibsRef[y];P.cbs[y]=v/100;P.ibsRef[y]=tot-v/100}
  else if(k==='total'){P.ibsRef[y]=v/100-P.cbs[y]}
  else P[k][y]=v/100;
  ['cbs','ibsRef','total'].forEach(kk=>{const e=$('par').querySelector(`input[data-k="${kk}"][data-y="${y}"]`);if(e){if(e!==i)e.value=+(gv(kk,y)*100).toFixed(4);e.classList.toggle('chg',Math.abs(gv(kk,y)-gv(kk,y,D))>1e-9)}});
  i.classList.toggle('chg',Math.abs(gv(k,y)-gv(k,y,D))>1e-9);render()})}
function buildEmp(){let h='<table><thead><tr><th>Empresa</th><th>Crédito (%)</th><th>IR/CSLL (%)</th><th>Fator ISS</th></tr></thead><tbody>';
 ORDER.forEach(k=>{const c=B.empresas[k],E=P.empresas[k]||{},y0=Object.keys(c.years)[0];const f=(n,ph,m)=>`<td><input type=number step=any data-e="${k}" data-f="${n}" placeholder="${ph}" value="${E[n]!==undefined?+(E[n]*m).toFixed(4):''}"></td>`;
  h+=`<tr><td>${c.nome}</td>${f('credito',(c.credito0*100).toFixed(0),100)}${f('ir',(c.years[y0].ir_rate*100).toFixed(2),100)}${f('issFator','1',1)}</tr>`});
 $('emp').innerHTML=h+'</tbody></table>';
 $('emp').querySelectorAll('input').forEach(i=>i.oninput=()=>{const k=i.dataset.e,f=i.dataset.f,m=f==='issFator'?1:100;P.empresas[k]=P.empresas[k]||{};if(i.value==='')delete P.empresas[k][f];else P.empresas[k][f]=parseFloat(i.value)/m;i.classList.toggle('chg',i.value!=='');render()})}
function buildAud(){const it=[['mercosulIss2026','Mercosul 2026: manter o ISS integral no Pós-Reforma (PPT mostra 5% do ISS antes da vigência)'],['mobiCredLegado','Mobi: retirar do OPEX Pós-Reforma o crédito legado (“12. Impostos”) mantido a partir de 2027'],['mobiDaCredito','Mobi 2028+: reconhecer crédito de CBS/IBS sobre a D&A, ausente na planilha']];
 $('aud').innerHTML=it.map(([k,l])=>`<label class="t"><input type=checkbox data-a="${k}" ${P.audit[k]?'checked':''}> ${l}</label>`).join('');
 $('aud').querySelectorAll('input').forEach(i=>i.onchange=()=>{P.audit[i.dataset.a]=i.checked;render()})}
