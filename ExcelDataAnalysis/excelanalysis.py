'''
author:dengwei
date:2020-05-06
descript:程序入口

'''

import sys
from excelgetdatas import *


if __name__ == "__main__":
    directory = input("请输入文件夹路径：")
    if len(directory) == 0:
        logger.info("输入的路径有误！")
        sys.exit(0)

    # 1.遍历文件夹，读取Excel文件，将信息汇总到一个Excel文件中 ------- 邓伟
    rows = get_all_excel_data(directory)

    # 2.对问题的处理人，做数据分析，输出柱形图 ------ 高国栋
