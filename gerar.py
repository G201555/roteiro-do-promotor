"""Gera roteiro por promotor a partir do Roteiro mensal da Penalty.
Uso: python gerar.py <roteiro_mensal.xlsx> <PROMOTOR|TODOS> <saida.xlsx>"""
import os, sys, warnings, datetime as dt, openpyxl
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Alignment
warnings.filterwarnings('ignore')
src, prom, out = sys.argv[1], sys.argv[2].strip().upper(), sys.argv[3]
todos = prom == 'TODOS'
ws = openpyxl.load_workbook(src, data_only=True)['ROTEIRO']
rows = list(ws.iter_rows(min_row=13, max_col=60, values_only=True))
hdr = list(rows[0]); i = {h: k for k, h in enumerate(hdr) if h}
dias = [h for h in hdr if isinstance(h, str) and '/' in h and h.split('/')[0].isdigit()]
dsem = {d: v for d, v in zip(hdr, list(ws.iter_rows(min_row=12, max_row=12, max_col=60, values_only=True))[0])}
mes = int(dias[0].split('/')[1]); ano = 2026
pnome = lambda r: str(r[i['PROMOTOR']]).strip().upper() if r[i['PROMOTOR']] else 'SEM PROMOTOR'
data = [r for r in rows[1:] if r[i['COD_LOJA']] is not None and (todos or pnome(r) == prom)]
assert data, 'promotor não encontrado'
data.sort(key=lambda r: pnome(r)) if todos else None  # sort estável: mantém a ordem da planilha dentro de cada promotor
marc = lambda v: v is not None and str(v).strip() != ''
cols = ['COD_LOJA','LOJA S3','BANDEIRA','SEGMENTAÇÃO','TIPO (shopping ou rua)','ENDEREÇO','BAIRRO','CIDADE','CEP','UF','CNPJFORMATADO','LÍDER','COD_ROTA','BOLA','CHUTEIRA','QTD SKUS']
pre = ['PROMOTOR'] if todos else []
outh = pre + ['COD_LOJA','LOJA','BANDEIRA','SEGMENTAÇÃO','TIPO (shopping ou rua)','ENDEREÇO','BAIRRO','CIDADE','CEP','UF','CNPJ','LÍDER','COD_ROTA','BOLA','CHUTEIRA','QTD SKUS'] + dias + ['DIAS PLANEJADOS','FREQUÊNCIA']
HF, HFont, HA = PatternFill('solid', fgColor='0F2747'), Font(bold=True, color='FFFFFF'), Alignment(horizontal='center', vertical='center', wrap_text=True)
def head(sh, h, widths):
    sh.append(h)
    for c in sh[1]: c.fill, c.font, c.alignment = HF, HFont, HA
    for k, w in enumerate(widths): sh.column_dimensions[openpyxl.utils.get_column_letter(k+1)].width = w
wb = openpyxl.Workbook(); res = wb.active; res.title = 'RESUMO'
ro = wb.create_sheet('ROTEIRO'); ag = wb.create_sheet('AGENDA POR DIA')
head(ro, outh, [34]*len(pre) + [9,52,20,13,16,40,20,16,10,5,16,18,12,7,9,9] + [6]*len(dias) + [11,11])
agenda = []; st = defaultdict(lambda: {'lojas': 0, 'com': 0, 'visitas': 0, 'rota': set(), 'lider': set()})
for r in data:
    p = pnome(r)
    vals = [r[i[c]] for c in cols]
    vals[8] = str(vals[8]).zfill(8) if vals[8] is not None else None
    vals[10] = str(vals[10]).zfill(14) if vals[10] is not None else None
    m = ['X' if marc(r[i[d]]) else None for d in dias]
    n = sum(1 for x in m if x)
    ro.append(([p] if todos else []) + vals + m + [n, r[i['FREQUÊNCIA']]])
    s = st[p]; s['lojas'] += 1; s['com'] += n > 0; s['visitas'] += n
    if r[i['COD_ROTA']]: s['rota'].add(r[i['COD_ROTA']])
    if r[i['LÍDER']]: s['lider'].add(r[i['LÍDER']])
    for d, x in zip(dias, m):
        if x: agenda.append(([p] if todos else []) + [dt.datetime(ano, int(d.split('/')[1]), int(d.split('/')[0])), dsem[d], r[i['COD_LOJA']], r[i['LOJA S3']], r[i['BANDEIRA']], r[i['BAIRRO']], r[i['CIDADE']], r[i['ENDEREÇO']]])
c0 = 17 + len(pre)
for row in ro.iter_rows(min_row=2, min_col=c0, max_col=len(outh)):
    for c in row: c.alignment = Alignment(horizontal='center')
ro.freeze_panes = 'D2' if todos else 'C2'; ro.auto_filter.ref = ro.dimensions
head(ag, pre + ['DATA','DIA','COD_LOJA','LOJA','BANDEIRA','BAIRRO','CIDADE','ENDEREÇO'], [34]*len(pre) + [11,6,9,52,20,20,16,40])
k = len(pre)
for a in sorted(agenda, key=lambda a: (a[:k], a[k], str(a[k+6]), str(a[k+5]))):
    ag.append(a); ag.cell(ag.max_row, k+1).number_format = 'DD/MM/YYYY'
ag.freeze_panes = 'A2'; ag.auto_filter.ref = ag.dimensions
nomes = {10:'outubro',11:'novembro',12:'dezembro'}
fonte = f'{os.path.basename(src)} ({os.path.dirname(os.path.abspath(src))}), versão de {dt.datetime.fromtimestamp(os.path.getmtime(src)):%d/%m/%Y}'
j = lambda s: ', '.join(sorted(s)) or None
if todos:
    head(res, ['PROMOTOR','ROTA','LÍDER','LOJAS NO ROTEIRO',f'LOJAS COM DIA MARCADO EM {nomes.get(mes, mes).upper()}','LOJAS SEM DIA MARCADO','VISITAS PLANEJADAS NO MÊS'], [38,22,34,12,16,14,14])
    for p, s in sorted(st.items()):
        res.append([p, j(s['rota']), j(s['lider']), s['lojas'], s['com'], s['lojas'] - s['com'], s['visitas']])
    tot = [sum(s[x] for s in st.values()) for x in ('lojas', 'com', 'visitas')]
    res.append(['TOTAL', None, None, tot[0], tot[1], tot[0] - tot[1], tot[2]])
    for c in res[res.max_row]: c.font = Font(bold=True)
    res.append([]); res.append(['Fonte: ' + fonte])
else:
    s = st[prom]
    head(res, ['ITEM','VALOR'], [42,80])
    for kk, v in [('Promotor', prom), ('Rota', j(s['rota'])), ('Líder', j(s['lider'])), ('Fonte', fonte),
                  ('Lojas no roteiro', s['lojas']), (f'Lojas com dia marcado em {nomes.get(mes, mes)}', s['com']),
                  ('Lojas sem dia marcado (frequência 0)', s['lojas'] - s['com']), ('Visitas planejadas no mês', s['visitas'])]:
        res.append([kk, v])
res.freeze_panes = 'A2'
wb.save(out)
print('ok', len(data), 'lojas', len(agenda), 'visitas', len(st), 'promotores')
