'''
对问题描述，做数据分析，输出词云图 ------ 丁小永

'''
import openpyxl
from wordcloud import WordCloud
import jieba
import numpy
import PIL.Image as Image
import xlsxwriter

from exceloperator import *
from loggerhelper import *

#1.将字符串切分为单个字符8q
def chinese_jieba(text):
    wordlist_jieba=jieba.cut(text)
    space_wordlist=''.join(wordlist_jieba)
    return  space_wordlist

#2.获取tital.xlsx中问题描述的值，返回str
def get_total_data():
    src = os.path.abspath('total.xlsx')
    excelop = ExcelOperator(src)
    problem_desdribe = excelop.get_col_value(7)
    return ''.join(problem_desdribe)

#3.画词云图
def print_wordcloud(text):
    # 1.将字符串切分
    text = chinese_jieba(text)
    # 2.图片遮罩层
    mask_pic = numpy.array(Image.open("pika.jpg"))
    wordcloud = WordCloud(font_path="C:/Windows/Fonts/simfang.ttf",
                          mask=mask_pic).generate(text)
    image = wordcloud.to_image()
    image.save('wordcloud.jpg');
    return image

#4.把词云图输出到excel中
def output_wordcloud():
    book = xlsxwriter.Workbook('wordcloud.xlsx')
    sheet = book.add_worksheet('wordcloud')
    sheet.insert_image('D4', 'wordcloud.jpg')
    book.close()

#if __name__ == "__main__":
def get_wordcloud():
    text = get_total_data()
    image = print_wordcloud(text)
    output_wordcloud()