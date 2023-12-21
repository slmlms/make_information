import math
import os
import pathlib

import numpy as np
import xlrd

import utils.data_util as data

# 选择的模块类型，只能有openpyxl和xlwings
model_type: str = 'openpyxl'

save_path = pathlib.Path('E:\工作\盛屯二期\资料\绝缘电阻测试记录\\')
work_book_path = "E:\工作\盛屯二期\图纸\电气施工图\B04 料液池\实验报告.xlsx"

excel_template_path = "inspection_lot\实验报告\线路绝缘电阻测试记录.xlsx"


def run(rows, i):
    print("当前进程id：", os.getpid())
    # Excel模板，注意选择打开方式
    excel_template = data.switch_open_excel_template(model_type, excel_template_path)

    place = rows[0][4]

    bianhao = rows[0][6]
    # date = data.int_to_date(rows[0][5])
    # 测试地点
    excel_template.worksheets[0]["M4"].value = place
    # 编号
    excel_template.worksheets[0]["M3"].value = "15MCC-" + bianhao + "-" + "{:0>3d}".format(i)
    # excel_template.sheets[0].range("M5").value = date
    j = 1
    for row in rows:
        num = row[0]
        start = row[1]
        end = row[2]
        cable = row[3]
        # 写入模板
        excel_template.worksheets[0]["B" + str(j + 8)].value = num
        excel_template.worksheets[0]["C" + str(j + 8)].value = start
        excel_template.worksheets[0]["D" + str(j + 8)].value = end
        excel_template.worksheets[0]["E" + str(j + 8)].value = cable
        j += 1

    save_file = save_path.joinpath("15MCC-" + bianhao + "-" + "{:0>3d}".format(i) + place + '.xlsx')
    excel_template.save(save_file)
    # s.append(save_path)


if __name__ == '__main__':
    # pool = multiprocessing.Pool(processes=8)
    sheet = xlrd.open_workbook(work_book_path).sheet_by_name('实验报告')
    # s = multiprocessing.Manager().list()
    datas = []
    # 一个报告17行。向上取整
    files = math.ceil(sheet.nrows / 17)
    # 将所有的行添加到一个数组
    for i in range(1, sheet.nrows):
        datas.append(sheet.row_values(i))
    # 利用numpy对数组分割
    rows_list = np.array_split(datas, files)

    count = 1
    for rows in rows_list:
        print(len(rows))
        # run(rows, count)
        count += 1
        # pool.apply_async(run, args=(rows, i, s))
    data.close_excel()
    # pool.close()
    # pool.join()
    # data.make_pdf(set(s))
