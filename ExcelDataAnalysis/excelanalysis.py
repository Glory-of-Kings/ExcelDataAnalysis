'''
author:dengwei
date:2020-05-06
descript:程序入口

'''

import sys
from excelgetdatas import *
from charttop5levelratio import *
from chartdescript import *
from chartfeedbacker import *
from chartrequestlevel import get_requestleval_piechart


if __name__ == "__main__":
    directory = input("请输入文件夹路径：")
    if len(directory) == 0:
        logger.info("输入的路径有误！")
        sys.exit(0)

    # 1.遍历文件夹，读取Excel文件，将信息汇总到一个Excel文件中 ------- 邓伟
    rows = get_all_excel_data(directory)

    # 2.对问题的处理人，做数据分析，输出柱形图 ------ 高国栋
    get_handler()

    # 3.对问题的反馈人，做数据分析，输出柱形图 ------ 邱大发
    get_feedbacker_chart()


    # 4.对问题反馈人(Top5)反馈的问题评级占比分析，输出柱形图 ------ 江坚
    dadaist = data_create(rows)
    draw_top5(dadaist)

	
	# 5.对问题描述，做数据分析，输出词云图 ------ 丁小永
    get_wordcloud()

    # 6.对问题评级，做数据分析，输出饼图 ------ 纪志昌
    get_requestleval_piechart()


    
