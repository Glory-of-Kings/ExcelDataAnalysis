'''
author:dengwei
date:2020-02-29
descript:记录文本日志

'''
import os
import sys
import logging
import datetime
from pathlib import Path

os.chdir(sys.path[0])#解决相对路径的问题

dir_name = "logs"
root = os.getcwd()
dir = Path(root, dir_name)
if dir.exists() == False:
    os.chdir(root)
    os.mkdir(dir_name)


date = datetime.datetime.now()

detester = date.strftime('%Y-%m-%d')

# 获取logger对象,取名mylog
logger = logging.getLogger("mylog")
# 输出DEBUG及以上级别的信息，针对所有输出的第一层过滤
logger.setLevel(level=logging.DEBUG)

file_name = os.path.join("logs", detester+".log")

# 获取文件日志句柄并设置日志级别，第二层过滤
fh = logging.FileHandler(file_name)
fh.setLevel(logging.DEBUG)

#创建一个handler，用于将日志输出到控制台
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)

# 生成并设置文件日志格式，其中name为上面设置的mylog
formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
ch.setFormatter(formatter)

# 为logger对象添加句柄
logger.addHandler(fh)
logger.addHandler(ch)
