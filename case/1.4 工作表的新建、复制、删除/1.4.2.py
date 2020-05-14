import openpyxl
wb=openpyxl.load_workbook('demo3.xlsx')
wb.copy_worksheet(wb['工资表']).title='工资表1月'
wb.save('demo3-1.xlsx')