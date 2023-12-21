# 检验批质量验收记录
import multiprocessing as mp
import os
import pathlib
import time

import xlrd
from loguru import logger
from openpyxl.cell import MergedCell

import utils.data_util as data
import utils.log_util as log
from resources import setting_util as setting

# 选择的模块类型，只能有openpyxl和xlwings
model_type: str = 'openpyxl'
# 施工检查记录模型
data_mapping_record = 'SGJCJL_DataMapping_JYP'
data_mapping = 'JYP_DataMapping'
cell_mapping = 'JYP_CellMapping'
# 获取配置文件
config = setting.get_config('setting.cfg')
logger.debug(config.sections())
# 记录日志
log.to_log(config, '检验批')
# 数据源路径
sheets = "电气"
workbook_path = config.get('default', 'DataSource')
sheet_names = sheets.split(',')


def run(d, titles, row, s: list):
    logger.debug(os.getpid())
    # logger.info('正在制作第{i}行，已完成{b}%', i=i, b=i / sheet.nrows * 100)
    # 子项名称
    child_name = d.get('child_name')
    # 检验批编号
    serial_number = d.get('true_number')
    # 检验批标题
    title = d.get('title')
    print(child_name, serial_number, title)
    # 检验批模板路径
    template_path = data.find_template_xlsx(config, 'JYP_template', serial_number[0:5], title)
    logger.info("检验批模板路径为{p}", p=template_path)
    # Excel模板，注意选择打开方式
    excel_template = data.switch_open_excel_template(model_type, template_path)
    # 修改excel_template sheet0单元格N5的值为“项目技术负责人”
    merged_cell = excel_template.worksheets[0]['U5'].offset(row=0, column=-1)
    if isinstance(merged_cell, MergedCell):  # 判断该单元格是否为合并单元格
        for merged_range in excel_template.worksheets[0].merged_cell_ranges:  # 循环查找该单元格所属的合并区域
            if merged_cell.coordinate in merged_range:
                # 获取合并区域左上角的单元格作为该单元格的值返回
                merged_cell = excel_template.worksheets[0].cell(row=merged_range.min_row, column=merged_range.min_col)
                merged_cell.value = "技术负责人"
    # 保存模板路径
    save_path_template = pathlib.Path(config.get('default', 'outPutDir')).joinpath(child_name).joinpath(
        data.find_save_path(config, 'JYP_outputPath', serial_number[0:5]))
    if not pathlib.Path(save_path_template).exists():
        pathlib.Path.mkdir(pathlib.Path(save_path_template), True, True)
    logger.debug(save_path_template)

    # 最终保存的文件
    save_file = save_path_template.joinpath(serial_number + child_name + title + '.xlsx')
    # 写入模板
    data.switch_write_excel_template(model_type, excel_template, cell_mapping, d, save_file)
    logger.info("保存为{f}成功！", f=save_file)
    s.append(save_path_template)

    return
    # 施工检查记录

    # data.construction_inspection_record(model_type, data_mapping_record, titles, row, serial_number,
    #                                     child_name, title, save_path_template)
    # # 工程报验表
    # engineering_inspection_form_template = setting.get_word_template('inspection_lot\报验表\报验表.docx')
    # engineering_inspection_form_number = serial_number[:6] + "C2" + serial_number[8:]
    # check_parts = child_name + d.get('acceptance_part') + d.get('sub_project_name')
    # inspection_lot = d.get('sub_project_name')
    # engineering_inspection_form_save_file = save_path_template.joinpath(
    #     serial_number + child_name + title + '报验表.docx')
    # context = {'BYBBianHao': engineering_inspection_form_number, 'YinJianBuWei': check_parts, 'JYP': inspection_lot,
    #            'BianHao': serial_number}
    # data.write_word_template(engineering_inspection_form_template, context, engineering_inspection_form_save_file)


if __name__ == '__main__':

    start = time.time()
    # 进程数量，创建进程池
    pross = 12
    pool = mp.Pool(processes=pross)
    # 全局共享变量
    s = mp.Manager().list()

    for sheet_name in sheet_names:
        # 数据源工作表
        workbook = xlrd.open_workbook(workbook_path)
        sheet = workbook.sheet_by_name(sheet_name)
        titles = data.read_titles(workbook_path, sheet_name)
        # 标题行
        logger.debug(titles)
        for i in range(1, sheet.nrows):
            row = sheet.row_values(i)
            d = data.get_object(data_mapping, titles, row)
            # 添加基本信息
            d.setdefault("general_contractor", "刚果（金）卡隆威矿业有限公司")
            d.setdefault("turnkey_project_manager", "高磊")
            d.setdefault("subcontractor", "十五冶对外工程有限公司")
            d.setdefault("subcontracting_project_manager", "冯齐名")
            logger.debug(d)

            if data.whether_to_submit(d) == True:
                continue
            # 使用异步方法并行处理，多线程尽量不要传入文件作为参数
            pool.apply_async(run, args=(d, titles, row, s))

    # 先结束进程池，后join，否则会报错
    pool.close()
    pool.join()
    data.close_excel()
    data.make_pdf(set(s))
    logger.success('共耗时{t} ms', t=time.time() - start)
