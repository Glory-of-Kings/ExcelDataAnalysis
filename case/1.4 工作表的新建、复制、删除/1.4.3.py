import openpyxl
wb=openpyxl.load_workbook('demo3-1.xlsx')
wb.remove(wb['工资表'])
wb.save('demo3-1.xlsx')