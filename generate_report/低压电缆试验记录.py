import multiprocessing
import os
import pathlib

import docxtpl
import xlrd

import utils.data_util as data

save_path = pathlib.Path('E:\工作\庞比\报验资料\调试记录\石灰乳\低压电缆\\')
work_book_path = "E:\工作\庞比\施工图\\906刚果（金）庞比铜钴矿项目蓝图PDF版\电气仪表\电力\石灰乳及石灰石浆制备-电力\石灰乳电缆表.xls"


def run(row, i, s):
    print("当前进程id：", os.getpid())
    # 模板要放到循环里面，不然会生成同一个文件
    word_template = docxtpl.DocxTemplate(
        'E:\ideaProject\make_information\\resources\inspection_lot\实验报告\低压电缆实验记录模板.docx')
    num = row[0]
    start = row[1]
    end = row[2]
    u0 = row[4]
    cable = row[3]
    len = row[5]
    place = row[7]
    date = data.int_to_date(row[6])
    # id = sheet.cell_value(i, 8)
    context = {'id': num, 'num': num, 'start': start, 'end': end, 'u0': u0, 'cable': cable, 'len': len,
               'date': date,
               'place': place}
    print(context)
    save_file = save_path.joinpath(str(i) + '.docx')
    data.write_word_template(word_template, context, save_file)
    s.append(save_path)


if __name__ == '__main__':
    pool = multiprocessing.Pool(processes=8)
    sheet = xlrd.open_workbook(work_book_path).sheet_by_name('实验报告')
    s = multiprocessing.Manager().list()
    for i in range(1, sheet.nrows):
        row = sheet.row_values(i)
        pool.apply_async(run, args=(row, i, s))

    pool.close()
    pool.join()
    data.make_pdf(set(s))
