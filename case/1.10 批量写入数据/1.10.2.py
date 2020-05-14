import openpyxl
wb=openpyxl.load_workbook('demo.xlsx')
ws=wb.active
# for r in [['%d*%d=%d'%(y,x,x*y) for y in range(1,x+1)] for x in range(1,10)]:
#     ws.append(r)
# ws.delete_rows(1)
# wb.save('demo.xlsx')
for row in ws['a1:c6']:
    for c in row:
        c.value=1
wb.save('demo.xlsx')