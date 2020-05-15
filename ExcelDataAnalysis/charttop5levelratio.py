'''
对问题反馈人(Top5)反馈的问题评级占比分析，输出柱形图 ------ 江坚
'''
import xlsxwriter


def data_create(dateResource):
    dadaist = []
    jxx_dict = {}
    jxx_dicta = {}
    jxx_dictb = {}
    jxx_dictc = {}
    jxx_dictd = {}
    for j in range(0, len(dateResource)):
        str4 = dateResource[j][3]
        str1 = str4[:str4.index('\n')]
        stra = dateResource[j][8]
        if stra == "A":
            if str1 in jxx_dicta:
                jxx_dicta[str1] = (jxx_dicta.get(str1) + 1)
            else:
                jxx_dicta.setdefault(str1, 1)
        elif stra == "B":
            if str1 in jxx_dictb:
                jxx_dictb[str1] = (jxx_dictb.get(str1) + 1)
            else:
                jxx_dictb.setdefault(str1, 1)
        elif stra == "C":
            if str1 in jxx_dictc:
                jxx_dictc[str1] = (jxx_dictc.get(str1) + 1)
            else:
                jxx_dictc.setdefault(str1, 1)
        else:
            if str1 in jxx_dictd:
                jxx_dictd[str1] = (jxx_dictd.get(str1) + 1)
            else:
                jxx_dictd.setdefault(str1, 1)
        if str1 in jxx_dict:
            jxx_dict[str1] = (jxx_dict.get(str1) + 1)
        else:
            jxx_dict.setdefault(str1, 1)
    a1 = sorted(jxx_dict.items(), key=lambda x: x[1], reverse=True)
    if len(a1) > 5:
        b1 = 5
    else:
        b1 = len(a1)
    for m in range(0, b1):
        listrange = [a1[m][0], jxx_dicta.get(a1[m][0]), jxx_dictb.get(a1[m][0]), jxx_dictc.get(a1[m][0]),
                     jxx_dictd.get(a1[m][0])]
        dadaist.append(listrange)
    return dadaist


def draw_top5(dadaist):
    # 新建一个excel文件，起名为drawChart1.xlsx
    workbook = xlsxwriter.Workbook("drawTop5.xlsx")
    # 添加一个Sheet页，不添写名字，默认为Sheet1
    worksheet = workbook.add_worksheet()
    # 准备数据
    headings = ["姓名", "A", "B", "C", "D"]
    # data = [["板面张", 78, 60], ["糖人李", 98, 89], ["炸糕徐", 88, 100]]
    data = dadaist
    # 样式
    head_style = workbook.add_format({"bold": True, "bg_color": "yellow", "align": "center", "font": 13})
    # 写数据
    worksheet.write_row("A1", headings, head_style)
    for i in range(0, len(data)):
        worksheet.write_row("A{}".format(i + 2), data[i])
    # 添加柱状图
    chart1 = workbook.add_chart({"type": "column"})
    chart1.add_series({
        "name": "=Sheet1!$B$1",  # 图例项
        "categories": "=Sheet1!$A$2:$A$6",  # X轴 Item名称
        "values": "=Sheet1!$B$2:$B$6",  # X轴Item值
        'column': {'color': 'red'},  # 图表颜色
    })
    chart1.add_series({
        "name": "=Sheet1!$C$1",
        "categories": "=Sheet1!$A$2:$A$6",
        "values": "=Sheet1!$C$2:$C$6",
        'column': {'color': 'yellow'},  # 图表颜色
    })
    chart1.add_series({
        "name": "=Sheet1!$D$1",
        "categories": "=Sheet1!$A$2:$A$6",
        "values": "=Sheet1!$D$2:$D$6",
        'column': {'color': 'blue'},  # 图表颜色
    })
    chart1.add_series({
        "name": "=Sheet1!$E$1",
        "categories": "=Sheet1!$A$2:$A$6",
        "values": "=Sheet1!$E$2:$E$6",
        'column': {'color': 'black'},  # 图表颜色
    })
    # 添加柱状图标题
    chart1.set_title({"name": "Top5反馈评级数"})
    # Y轴名称
    chart1.set_y_axis({"name": "数量"})
    # X轴名称
    chart1.set_x_axis({"name": "人名"})
    # 图表样式
    chart1.set_style(11)
    # 插入图表
    worksheet.insert_chart("B10", chart1)
    # 关闭EXCEL文件
    workbook.close()
