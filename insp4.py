import openpyxl,sys,warnings
from collections import Counter
warnings.filterwarnings('ignore')
wb=openpyxl.load_workbook(sys.argv[1])
ws=wb['Planilha1']
rs=list(ws.iter_rows(min_row=2,values_only=True))
print(Counter((r[0],r[2],r[3]) for r in rs))
print('multi-dia',[r for r in rs if sum(int(x or 0) for x in r[4:11])!=1])
print('dup loja-semana',[k for k,v in Counter((r[1],r[3]) for r in rs).items() if v>1])
print('types', {type(v).__name__ for r in rs for v in r[:4]})
for r in rs[29:]: print(r)
src=openpyxl.load_workbook(sys.argv[2],data_only=True)['ROTEIRO']
rows=list(src.iter_rows(min_row=13,max_col=45,values_only=True)); h=rows[0]; i={x:k for k,x in enumerate(h) if x}
print(Counter((r[i['PROMOTOR']],r[i['COD_ROTA']],r[i['COD ROTA S3']]) for r in rows[1:]))
codes={r[1] for r in rs}
al=[r for r in rows[1:] if r[i['COD_LOJA']] in codes]
print('example stores in nov:',len(al),'of',len(codes), Counter(r[i['PROMOTOR']] for r in al))
