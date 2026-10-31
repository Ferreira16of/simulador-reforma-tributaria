// ===== Renderizador: reproduz as formas do PPT v15 em HTML (canvas 1280x720) =====
const FONT='Aptos,"Segoe UI",Calibri,Arial,sans-serif';
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}
function parHtml(ps,def){
  return ps.map(p=>{
    const sz0=(p.runs[0]&&p.runs[0].sz)||p.esz||def||10;
    let st=`text-align:${p.al==='ctr'?'center':p.al==='r'?'right':p.al==='just'?'justify':'left'};font-size:${(sz0*4/3).toFixed(2)}px;line-height:${p.ls?p.ls*1.2:1.2};`;
    if(p.sb)st+=`margin-top:${p.sb.toFixed(1)}px;`;
    let pre='';
    if(p.bu){st+=`padding-left:${Math.max(p.ml,0).toFixed(1)}px;text-indent:${p.ind.toFixed(1)}px;`;pre=`<span class="bu" style="width:${Math.max(-p.ind,0)}px">${esc(p.bu)}</span>`}
    const rs=p.runs.length?p.runs.map(r=>{
      if(r.t==='\n')return '<br>';
      let s='';if(r.sz)s+=`font-size:${(r.sz*4/3).toFixed(2)}px;`;if(r.b)s+='font-weight:700;';if(r.i)s+='font-style:italic;';if(r.c)s+=`color:${r.c};`;
      return `<span style="${s}">${esc(r.t)}</span>`}).join(''):'&nbsp;';
    return `<div class="pp" style="${st}">${pre}${rs}</div>`}).join('');
}
function boxStyle(d){
  let s=`left:${d.x}px;top:${d.y}px;width:${d.w}px;height:${d.h}px;`;
  if(d.f)s+=`background:${d.f};`;
  if(d.ln){s+=`border:${Math.max(d.ln[1],.5)}px ${d.ln[2]==='solid'?'solid':'dashed'} ${d.ln[0]};box-sizing:border-box;`}
  if(d.g==='roundRect')s+=`border-radius:${(Math.min(d.w,d.h)*(d.adj||.1667)).toFixed(1)}px;`;
  if(d.g==='ellipse')s+='border-radius:50%;';
  if(d.g==='homePlate')s+=`clip-path:polygon(0 0,${Math.max(d.w-d.h/2,0)}px 0,100% 50%,${Math.max(d.w-d.h/2,0)}px 100%,0 100%);`;
  return s;
}
function elHtml(d,idx){
  if(d.chart)return `<div class="chart" data-n="${d.n}" style="left:${d.x}px;top:${d.y}px;width:${d.w}px;height:${d.h}px"></div>`;
  if(d.tbl)return tblHtml(d,idx);
  if(d.slnum&&!d.p)return `<div class="sh sn" style="left:${d.x}px;top:${d.y}px;width:${d.w}px;height:${d.h}px"><span class="snv"></span></div>`;
  if(d.g==='line'){
    const c=d.ln?d.ln[0]:'#000',t=d.ln?Math.max(d.ln[1],.75):1;
    if(d.h<=1.5||d.h<d.w/20)return `<div class="sh" style="left:${d.x}px;top:${d.y}px;width:${d.w}px;height:${t}px;background:${c}"></div>`;
    if(d.w<=1.5||d.w<d.h/20)return `<div class="sh" style="left:${d.x}px;top:${d.y}px;width:${t}px;height:${d.h}px;background:${c}"></div>`;
    return `<svg class="sh" style="left:${d.x}px;top:${d.y}px;width:${d.w}px;height:${d.h}px" viewBox="0 0 ${d.w} ${d.h}"><line x1="0" y1="${d.flipV?d.h:0}" x2="${d.w}" y2="${d.flipV?0:d.h}" stroke="${c}" stroke-width="${t}"/></svg>`;
  }
  let inner='';
  if(d.p){
    const anc=d.anc==='ctr'?'center':d.anc==='b'?'flex-end':'flex-start';
    const ins=d.ins||[7.2,3.6,7.2,3.6];
    inner=`<div class="tx${d.wrap==='none'?' nw':''}" style="justify-content:${anc};padding:${ins[1]}px ${ins[2]}px ${ins[3]}px ${ins[0]}px">${parHtml(d.p)}</div>`;
    if(d.slnum)inner=inner.replace(/>\s*[^<>]*‹nº›[^<]*</,'><span class="snv"></span><');
  }
  return `<div class="sh${d.p?' hast':''}" ${d.n?`data-n="${d.n}"`:''} data-x="${d.x}" data-y="${d.y}" data-w="${d.w}" style="${boxStyle(d)}">${inner}</div>`;
}
function tblHtml(d,idx){
  let cg=d.cols.map(w=>`<col style="width:${w}px">`).join('');
  let rows='';
  const R=d.rows;
  R.forEach((r,ri)=>{
    let tds='';
    r.c.forEach((c,ci)=>{
      if(c.hm||c.vm)return;
      let st='';if(c.f)st+=`background:${c.f};`;
      const m=c.m||[3,0,3,0];st+=`padding:${m[1]}px ${m[2]}px ${m[3]}px ${m[0]}px;vertical-align:${c.anc==='ctr'?'middle':c.anc==='b'?'bottom':'top'};`;
      [['bl','border-left'],['br','border-right'],['bt','border-top'],['bb','border-bottom']].forEach(([k,p])=>{if(c[k])st+=`${p}:${Math.max(c[k][1],.5)}px solid ${c[k][0]};`});
      const sp=(c.gs?` colspan="${c.gs}"`:'')+(c.rs?` rowspan="${c.rs}"`:'');
      tds+=`<td${sp} style="${st}">${parHtml(c.p,10)}</td>`});
    rows+=`<tr style="height:${r.h}px">${tds}</tr>`});
  return `<table class="tb" data-n="${d.n}" style="left:${d.x}px;top:${d.y}px;width:${d.w}px"><colgroup>${cg}</colgroup>${rows}</table>`;
}
function slideHtml(S){
  const own=S.el.some(e=>e.slnum);
  return (S.lay||[]).filter(e=>!(own&&e.slnum)).map(elHtml).join('')+S.el.map(elHtml).join('');
}
