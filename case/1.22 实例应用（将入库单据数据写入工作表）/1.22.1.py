import openpyxl
wb1=openpyxl.load_workbook('入库单.xlsx',data_only=True)
wb2=openpyxl.load_workbook('数据库.xlsx')
ws1=wb1.active;ws2=wb2.active
l=list(ws1.values)
t=(l[2][1],l[2][3],l[2][5])
if not t[2] in [c.value for c in ws2['i']]:
    for row in l[5:]:
        if not None in row:
            ws2.append(row+t)
    wb2.save('数据库.xlsx')
    print('保存成功！')
else:
    print('已保存')
