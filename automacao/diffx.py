import sys,openpyxl
a,b=sys.argv[1],sys.argv[2]
wa=openpyxl.load_workbook(a,data_only=True,read_only=True);wb=openpyxl.load_workbook(b,data_only=True,read_only=True)
print('sheets',wa.sheetnames==wb.sheetnames, set(wb.sheetnames)^set(wa.sheetnames))
tot=0
for n in wb.sheetnames:
    if n not in wa.sheetnames: continue
    A={ (r,c):v for r,row in enumerate(wa[n].iter_rows(values_only=True),1) for c,v in enumerate(row,1) if v is not None}
    B={ (r,c):v for r,row in enumerate(wb[n].iter_rows(values_only=True),1) for c,v in enumerate(row,1) if v is not None}
    d=[k for k in set(A)|set(B) if A.get(k)!=B.get(k) and not (isinstance(A.get(k),(int,float)) and isinstance(B.get(k),(int,float)) and abs(A[k]-B[k])<1e-6*max(1,abs(A[k])))]
    if d:
        d.sort();tot+=len(d);print(' ',n,len(d))
        for k in d[:int(sys.argv[3]) if len(sys.argv)>3 else 6]:print('    ',k,repr(A.get(k))[:50],'->',repr(B.get(k))[:50])
print('total',tot)
