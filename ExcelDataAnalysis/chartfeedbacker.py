'''
对问题的反馈人，做数据分析，输出柱形图 ------ 邱大发
'''

import openpyxl

def create_sheets(wbb):
    
    for i in range(1,10):
        wbb.create_sheet('第{}表'.format(i),i)
    wbb.remove(wbb['Sheet'])
    wbb.save('test.xlsx')

def test():
    wb = openpyxl.Workbook()
    print('test.xlsx')
    wb.save('test.xlsx')
    create_sheets(wb)
    

test()


    
        









