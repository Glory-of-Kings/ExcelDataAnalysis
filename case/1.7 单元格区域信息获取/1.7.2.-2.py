import openpyxl
wb=openpyxl.load_workbook('test.xlsx')
ws=wb.active
print(['%s-%d'%(row[0],sum(row[1:])) for row in list(ws.values)[1:]])