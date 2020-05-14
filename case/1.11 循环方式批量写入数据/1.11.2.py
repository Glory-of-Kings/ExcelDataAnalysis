import openpyxl
wb=openpyxl.load_workbook('demo.xlsx')
ws=wb.active
rngs=ws.iter_rows(min_row=2,min_col=2)
for row in rngs:
    for c in row:
        if c.value>=90:
            c.value=str(c.value)+'（优秀）'
wb.save('demo1.xlsx')