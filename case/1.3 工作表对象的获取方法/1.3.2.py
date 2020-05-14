import openpyxl
wb=openpyxl.load_workbook('各年业绩表.xlsx')
for sh in wb.worksheets:
    sh.title=sh.title+'-芝华公司'
wb.save('各年业绩表(修改后).xlsx')