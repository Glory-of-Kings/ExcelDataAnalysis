python version：3.7.3 or 3.7.4
IDE:vs code
openpyxl:pip install openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple/
官方文档：https://openpyxl.readthedocs.io/en/stable/index.html

需求：对转研发分析的现场问题，做数据分析
输出：问题处理人，问题反馈人，问题描述，评级等做数据分析，输出相关图表

功能点：
1.遍历文件夹，读取Excel文件，将信息汇总到一个Excel文件中 ------- 邓伟
2.对问题的处理人，做数据分析，输出柱形图 ------ 高国栋
3.对问题的反馈人，做数据分析，输出柱形图 ------ 邱大发
4.对问题反馈人(Top5)反馈的问题评级占比分析，输出柱形图 ------ 江坚
5.对问题描述，做数据分析，输出词云图 ------ 丁小永
6.对问题评级，做数据分析，输出饼图 ------ 纪志昌

设计：
每一个功能点单独新建module[文件]
每一个图表都输出到一个单独的Sheet中
