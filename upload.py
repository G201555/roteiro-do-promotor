"""Gera a planilha de upload do roteiro (layout 'Alterações consultor.xlsx') a partir do Roteiro mensal.
Uma linha por loja x semana ISO com visita; DOM..SAB = 1 nos dias marcados com X.
Uso: python upload.py <roteiro_mensal.xlsx> <saida.xlsx> [ano]"""
import sys, warnings, datetime as dt, openpyxl
from collections import OrderedDict, Counter
warnings.filterwarnings('ignore')
src, out = sys.argv[1], sys.argv[2]
ano = int(sys.argv[3]) if len(sys.argv) > 3 else 2026
ws = openpyxl.load_workbook(src, data_only=True)['ROTEIRO']
rows = list(ws.iter_rows(min_row=13, max_col=60, values_only=True))
hdr = list(rows[0]); i = {h: k for k, h in enumerate(hdr) if h}
dias = [(k, h) for k, h in enumerate(hdr) if isinstance(h, str) and '/' in h and h.split('/')[0].isdigit()]
marc = lambda v: v is not None and str(v).strip() != ''
semanas = OrderedDict(); sem_rota = []; ordem = {}
for n, r in enumerate(rows[1:]):
    if r[i['COD_LOJA']] is None: continue
    marcados = [dt.date(ano, int(h.split('/')[1]), int(h.split('/')[0])) for k, h in dias if marc(r[k])]
    if not marcados: continue
    rota = r[i['COD ROTA S3']]
    if rota is None:
        sem_rota.append((r[i['COD_LOJA']], r[i['LOJA S3']], r[i['PROMOTOR']], len(marcados))); continue
    for d in marcados:
        y, w, wd = d.isocalendar()  # wd: 1=seg ... 7=dom
        key = (int(rota), int(r[i['COD_LOJA']]), y, w)
        ordem.setdefault(key, n)
        semanas.setdefault(key, [0] * 7)[wd % 7] = 1  # índice 0=DOM, 1=SEG ... 6=SAB
wb = openpyxl.Workbook(); sh = wb.active; sh.title = 'Planilha1'
sh.append(['COD_ROTA','COD_LOJA','ANO','SEMANA_ANO','DOM','SEG','TER','QUA','QUI','SEX','SAB','DATA_INICIO','DATA_FIM'])
for key in sorted(semanas, key=lambda k: (k[0], k[2], k[3], ordem[k])):
    rota, loja, y, w = key; v = semanas[key]
    sh.append([rota, loja, str(y), str(w), str(v[0])] + v[1:6] + [str(v[6]), None, None])
    for c in sh[sh.max_row][4:11]: c.number_format = '@'
wb.save(out)
print('linhas', len(semanas), 'visitas', sum(sum(v) for v in semanas.values()))
print('por rota/semana', sorted(Counter((k[0], k[3]) for k in semanas).items()))
print('multi-dia na semana', sum(1 for v in semanas.values() if sum(v) > 1))
print('SEM COD ROTA S3 (fora do arquivo):', sem_rota)
