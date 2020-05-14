import openpyxl
wb=openpyxl.load_workbook('demo.xlsx')
ws=wb.active
rngs=ws.iter_rows(min_row=2,min_col=2)
for row in rngs:
        sm=sum([c.value for c in row][0:4])
        if sm>=300:
            row[-1].value='优秀'
wb.save('demo2.xlsx')