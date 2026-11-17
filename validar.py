import sys, openpyxl
from collections import Counter
wb = openpyxl.load_workbook(sys.argv[1])
for r in wb['RESUMO'].iter_rows(values_only=True): print(r)
ro = list(wb['ROTEIRO'].iter_rows(values_only=True))
print('sem rota:', [r[:2] for r in ro[1:] if r[12] is None])
print('freq != dias:', [(r[0], r[-2], r[-1]) for r in ro[1:] if (r[-1] or 0) != r[-2]][:20])
ag = list(wb['AGENDA POR DIA'].iter_rows(values_only=True))
print(sorted(Counter((a[0].strftime('%d/%m'), a[1]) for a in ag[1:]).items()))
