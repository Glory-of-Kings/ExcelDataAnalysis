import openpyxl
wb=openpyxl.load_workbook('test.xlsx')
ws=wb.worksheets[0]
ws.append({'a':'张三','b':56,'c':'fdgsfg'})
wb.save('test.xlsx')