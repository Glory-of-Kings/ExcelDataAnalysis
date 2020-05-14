import openpyxl
wb=openpyxl.load_workbook('2018年.xlsx')
for sh in wb.worksheets:
    if sh.title.split('-')[0]!='北京':
        wb.remove(sh)
wb.save('北京.xlsx')