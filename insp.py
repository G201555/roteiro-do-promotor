import openpyxl,sys,warnings
warnings.filterwarnings('ignore')
wb=openpyxl.load_workbook(sys.argv[1])
for ws in wb.worksheets:
    print('--',ws.title,ws.max_row,ws.max_column)
    for r in ws.iter_rows(min_row=1,max_row=int(sys.argv[2]) if len(sys.argv)>2 else 30):
        print([ (c.value, c.number_format) if c.value is not None and not isinstance(c.value,str) else c.value for c in r])
    for dv in ws.data_validations.dataValidation: print('DV',dv.sqref,dv.formula1)
