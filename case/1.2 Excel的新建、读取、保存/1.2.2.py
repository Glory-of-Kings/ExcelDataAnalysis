import openpyxl
wb=openpyxl.load_workbook('我的工作簿.xlsx')
wb.save('我的工作簿-1.xlsx')
