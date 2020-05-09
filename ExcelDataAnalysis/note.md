## base
```python
#新建保存读取工作簿
wb = openpyxl.Workbook() #新建工作簿 W大写
wb.save('我的工作簿.xlsx') #保存工作簿
# 每次创建默认会存在一个worksheet，也可以自行创建更多的worksheet
wb.openpyxl.load_workbook('filename') #读取工作簿

#获取工作表
workbook.active #获取当前活动的工作表
workbook.worksheets[i] #以索引的方式获取工作表
workbook['sheetname'] #义工作表名获取工作表，此方法没有代码提示
workbook.worksheets #循环工作表
workbook.sheetnames #获取所有工作表名
worksheet.title #获取指定工作表名，可以返回工作表名称，也可以对其赋值来修改工作表名称，需要save才生效

#新建工作表
workbook.create_sheet() #新建工作表时会默认新建一个工作表sheet，默认工作表名为sheet1，sheet2，sheet3……
workbook.create_sheet('工作表名称',指定位置) #如果不指定位置，那么工作表默认放置在最后，指定位置从0开始

#复制工作表
wb.copy_worksheet(工作表对象) #参数必须时工作表对象而不是工作表名字
wb.copy_worksheet(wb['工资表']).title = '工资表1月' #在复制表的同时给新表重命名，对工作表进行复制，需要save

#删除工作表
wb.remove(工作表对象) #参数必须时工作表对象，参考复制工作表，同样需要save

#获取单元格
ws = wb.worksheet[0]
print(ws['A1'].value) #A1表示法
print(ws.cell(1,1).value) #坐标表示法
print([sh['b14'].value for sh in wb.worksheets]) #打印说有工作表的b14，前面加个sum可求和



```





