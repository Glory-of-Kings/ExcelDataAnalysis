import openpyxl
wb=openpyxl.load_workbook('test.xlsx')
ws=wb.worksheets[0]
for row in list(ws.rows)[1:]:
    l=[v.value for v in row]
    print(l[0],sum(l[1:]))