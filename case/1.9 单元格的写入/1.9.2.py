import openpyxl
wb=openpyxl.Workbook()
ws=wb.active
ws.title='九九表'
for x in range(1,10):
    for y in range(1,x+1):
        ws.cell(x,y,'%d×%d=%d'%(y,x,x*y))
wb.save('九九表.xlsx')