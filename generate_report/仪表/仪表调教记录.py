# 检验批质量验收记录
import multiprocessing as mp
import os
import pathlib
import time

import xlrd
from loguru import logger

import utils.data_util as data

excel_template_path = 'inspection_lot/仪表/仪表调试记录.xlsx'

data_path = "E:\工作\庞比\报验资料\仪表数据源.xlsx"

save_path = pathlib.Path("E:\工作\庞比\报验资料\调试记录\仪表\\")
model_type: str = 'xlwings'
sheet_name = "仪表"


def run():
    data_sheet = xlrd.open_workbook(data_path).sheet_by_name(sheet_name)
    pool = mp.Pool(processes=15)
    s = mp.Manager().list()
    for i in range(1, data_sheet.nrows):
        logger.debug("正在制作第{i}行，已完成{a}%", i=i, a=i / data_sheet.nrows * 100)
        values = data_sheet.row_values(i)
        if values[8] == 1:
            continue
        pool.apply_async(func=make_ziliao, args=(values,))
    pool.close()
    pool.join()
    s.append(save_path)
    data.make_pdf(set(s))


def make_ziliao(values: list):
    logger.debug(os.getpid())
    save_workbook = data.switch_open_excel_template(model_type, excel_template_path)
    yibiao_name = values[0]
    weihao = values[1]
    xinghao = values[2]
    zhizaochang = values[3]
    liangcheng = values[4]
    shiyong = values[5]
    zhunquedu = values[6]
    shuchu = values[7]

    save_workbook.sheets[0].range("N3").value = yibiao_name
    save_workbook.sheets[0].range("T3").value = weihao
    save_workbook.sheets[0].range("D4").value = xinghao
    save_workbook.sheets[0].range("N4").value = zhizaochang
    save_workbook.sheets[0].range("T4").value = zhunquedu
    save_workbook.sheets[0].range("D5").value = liangcheng
    save_workbook.sheets[0].range("N5").value = shiyong
    save_workbook.sheets[0].range("T5").value = shuchu
    save_name = save_path.joinpath(weihao + '.xlsx')
    save_workbook.save(save_name)
    save_workbook.close()


if __name__ == '__main__':
    start = time.time()
    data.close_excel()
    run()
    data.close_excel()
    logger.success('共耗时{t} ms', t=time.time() - start)
