import openpyxl,glob,pickle,sys
f,sheetidx,out=sys.argv[1],int(sys.argv[2]),sys.argv[3]
res={}
for mode in (True,False):
    wb=openpyxl.load_workbook(f,read_only=True,data_only=mode)
    ws=wb[wb.sheetnames[sheetidx]]
    g={}
    for r,row in enumerate(ws.iter_rows(max_col=140),1):
        for c,cell in enumerate(row,1):
            if cell.value is not None: g[(r,c)]=cell.value
    res['v' if mode else 'f']=g
    res['name']=wb.sheetnames[sheetidx]
pickle.dump(res,open(out,'wb'))
print(f,res['name'],len(res['v']),len(res['f']))
