import openpyxl
wb=openpyxl.load_workbook('test.xlsx')
ws=wb.worksheets[0]
ws.delete_cols(4,2)
ws.delete_rows(6,3)
wb.save('test.xlsx')
