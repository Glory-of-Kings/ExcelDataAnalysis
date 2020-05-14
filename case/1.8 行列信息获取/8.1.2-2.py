import openpyxl
wb=openpyxl.load_workbook('test.xlsx')
ws=wb.worksheets[0]
for col in list(ws.columns)[1:]:
    l=[v.value for v in col]
    print(l[0],max(l[1:]))