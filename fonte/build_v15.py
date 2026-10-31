d='/tmp/claude-0/s/v15/'
import re
css=open(d+'css.txt').read().replace('</style>','''
#drawer{position:fixed;top:0;right:0;bottom:0;width:min(860px,96vw);background:#fff;z-index:20;transform:translateX(102%);transition:transform .25s;box-shadow:-12px 0 40px rgba(23,54,93,.2);overflow:auto;padding:24px 28px 90px}
#drawer.on{transform:none}#drawer h2{font-size:22px;font-weight:300;color:#17365D;margin:0 0 4px}#drawer h2 b{font-weight:800}
#drawer h4{margin:22px 0 8px;font-size:12px;letter-spacing:2px;text-transform:uppercase;color:#5B6573}
#drawer table{border-collapse:collapse;width:100%;font-size:13px}#drawer th,#drawer td{padding:4px 6px;border-bottom:1px solid #E3E8EE;text-align:left}
#drawer input[type=number]{width:78px;padding:4px;border:1px solid #C9D3DF;border-radius:4px;font:inherit}
#drawer input.chg{border-color:#E69138;background:#FFF4E6}.t{display:flex;gap:8px;align-items:center;margin:6px 0}
#drawer .row{display:flex;gap:8px;flex-wrap:wrap;margin-top:14px}
[data-n^="Slide Number"] .pp,.sn{text-align:right!important}[data-n^="Slide Number"]{left:1197px!important;width:45px!important}
#pr{display:none}
@page{size:1280px 720px;margin:0}
@media print{html,body{height:auto!important;background:#fff!important}.app,#drawer{display:none!important}#pr{display:block}
#pr .pg{position:relative;width:1280px;height:720px;overflow:hidden;page-break-after:always;break-after:page}#pr .pg:last-child{page-break-after:auto;break-after:auto}
#pr .pg .cv2{position:absolute;left:0;top:0;width:1280px;height:720px;font-family:Aptos,"Segoe UI",Calibri,Arial,sans-serif}
*{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
@media(max-width:700px){#stage{padding:4px}.dock{padding:6px 8px;gap:6px;flex-wrap:wrap}.dock button{padding:8px 10px}#drawer{width:100vw;padding:18px 14px 90px}}
</style>''')
sk=open(d+'skel.html').read().replace('</div>\n<script>','''</div><div id="pr"></div>
<aside id="drawer"><div style="display:flex;align-items:center;gap:12px"><h2 style="flex:1">Simular <b>cenários</b></h2><button class="ghost" id="dclose">Fechar ✕</button></div>
<p style="color:#5B6573;font-size:13px">Altere as alíquotas por ano; todos os slides recalculam. Padrão: CBS + IBS = 28,0% (CBS 11,3468% e IBS 16,6532%) e redutor 44,2%, em todos os anos. Ao mudar a CBS, o IBS ajusta para manter a soma.</p>
<h4>Empresas incluídas</h4><label class="t"><input type="checkbox" id="l6"> Incluir Lote 6 (consolidado, tabelas e slide 21)</label><h4>Parâmetros por ano</h4><div id="par"></div><h4>Por empresa (opcional)</h4><div id="emp"></div><h4>Ajustes de auditoria</h4><div id="aud"></div>
<div class="row"><button id="rst">Restaurar padrão</button><button class="ghost" id="rst0">Valores originais (CBS 9,25% / IBS 18,75% / sem redutor)</button><button class="ghost" id="exp">Exportar cenário</button><button class="ghost" id="imp">Importar</button><input type="file" id="file" accept=".json" hidden></div></aside>
<script>''',1)
sk=sk.replace('<script>\nconst SLIDES','<script>\nconst SLIDES')
# remove app_static parts: use ui.js instead; keep SLIDES const from skel
calc=open(d+'calc.js').read().replace('/*BASES*/',open('/tmp/claude-0/s/out/bases.json').read()).replace('/*FLOW*/',open('/tmp/claude-0/s/flow27_v15.json').read()).replace('/*MOTOR*/',open('/tmp/claude-0/s/motor.js').read())

h='<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Reforma Tributária</title>'+css+sk
h+='const SLIDES='+open(d+'v15.json').read()+';const $=id=>document.getElementById(id);\n'
out=h+open(d+'render.js').read()+'\n'+calc+'\n'+open(d+'bp.js').read()+'\n'+open(d+'dyn.js').read()+'\n'+open(d+'ui.js').read()+'</script></body></html>'
open(d+'v15.html','w').write(out);print(len(out))
