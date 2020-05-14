import openpyxl
wb=openpyxl.load_workbook('test.xlsx')
ws=wb.worksheets[0]
for row in ws['a1:g9']:
    for c in row:
        c.value=100
wb.save('test.xlsx')