import openpyxl
for m in range(1,13):
    wb=openpyxl.Workbook()
    wb.save('%d月.xlsx'%m)