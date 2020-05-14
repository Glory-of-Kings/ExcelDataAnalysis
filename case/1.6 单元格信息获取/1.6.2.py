import openpyxl
wb=openpyxl.load_workbook('各年业绩表.xlsx')
print(sum([sh['b14'].value for sh in wb.worksheets]))