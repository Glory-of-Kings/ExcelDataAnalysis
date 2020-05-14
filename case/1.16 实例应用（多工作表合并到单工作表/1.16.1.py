import openpyxl
wb=openpyxl.load_workbook('各年业绩表.xlsx')
nwb=openpyxl.Workbook()
nws=nwb.active
nws.append(['年份','月份','金额'])
for ws in wb.worksheets:
    l=[[ws.title]+[c.value for c in row] for row in ws.rows][1:-1]
    for v in l:
        nws.append(v)
nwb.save('合并1.xlsx')
