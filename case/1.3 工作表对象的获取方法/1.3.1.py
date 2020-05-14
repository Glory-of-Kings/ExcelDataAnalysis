import openpyxl
wb=openpyxl.load_workbook('各年业绩表.xlsx')
ws1=wb.active#获取活动工作表
ws2=wb.worksheets[2]#以索引值方式获取工作表
ws3=wb['2012年']#以工作表名获取
# for sh in wb.worksheets:
#     print(sh)
# print(wb.sheetnames)
wb.worksheets[1].title='demo'
wb.save('各年业绩表-1.xlsx')
