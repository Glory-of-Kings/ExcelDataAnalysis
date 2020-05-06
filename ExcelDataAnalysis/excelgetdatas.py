'''
author:dengwei
date:2020-05-06
descript:遍历文件夹，读取Excel文件，将信息汇总到一个Excel文件中

'''
import os
import openpyxl
from exceloperator import *
from loggerhelper import *

def get_all_excel_data(directory):
    '''
    文件夹目录
    '''

    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "total"
    rows = []
    for root, dirs, files in os.walk(directory): # 开始遍历文件
            # root 表示当前正在访问的文件夹路径
            # dirs 表示该文件夹下的子目录名list
            # files 表示该文件夹下的文件list
            # 遍历文件
            for f in files:
                src = os.path.join(root, f)
                logger.info("正在遍历文件："+src)
                if f.endswith(".xlsx"):
                    excelop = ExcelOperator(src)
                    r,c = excelop.get_row_clo_num()
                    for i in range(1,r+1):
                        row = excelop.get_row_value(i)
                        rows.append(row)
                        worksheet.append(row)
    workbook.save("total.xlsx")
    return rows