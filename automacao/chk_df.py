import sys,json,subprocess;sys.path.insert(0,'/tmp/claude-0/s')
from extract import extract,CFG
from lib import *
UNIT={'leste':1000.0,'mercosul':1000.0,'welfare':1000.0}
d=json.loads(subprocess.check_output(['node','/tmp/w/v16/../ppt/calc_cli.js','/tmp/w/v16/bases.json','/tmp/w/v16/motor.js']))  # defaultParams = parâmetros originais das DFs
tot=0
for k in CFG:
    rows,out=extract(k);u=UNIT.get(k,1.0)
    for y,info in sorted(out.items()):
        m=d[k].get(str(y))
        if not m: continue
        for lab,key in (('Receita Liquida','RL'),('EBITDA','EBITDA'),('EBT','EBT'),('Lucro Líquido','LL')):
            for side,mm,col in (('atual','atual',info['atual']),('rt','rt',info['rt'])):
                v=(col.get(lab) or 0)/u;mv=m[mm][key]
                if abs(v-mv)>max(1,abs(v)*0.0005):
                    tot+=1;print(k,y,side,lab,'DF',round(v,1),'motor',round(mv,1),'dif',round(mv-v,1))
print('divergências',tot)
