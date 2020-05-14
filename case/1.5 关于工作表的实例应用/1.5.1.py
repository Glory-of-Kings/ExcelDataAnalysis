import openpyxl
wb=openpyxl.Workbook()
for m in range(1,13):
    wb.create_sheet('%d月'%m)
wb.remove(wb['Sheet'])
wb.save('2019年计划表.xlsx')