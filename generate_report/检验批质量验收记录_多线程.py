# 导入所需模块
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

# 初始化变量
model_type = 'openpyxl'
data_mapping_record = 'SGJCJL_DataMapping_JYP'
data_mapping = 'JYP_DataMapping'
cell_mapping = 'JYP_CellMapping'

# 获取配置文件和记录日志
config = setting.get_config('setting.cfg')
logger.debug(config.sections())
log.to_log(config, '检验批')

def load_data_sheet(sheet_name):
    """
    加载数据源工作表并获取标题行
    """
    workbook_path = config.get('default', 'DataSource')
    workbook = xlrd.open_workbook(workbook_path)
    sheet = workbook.sheet_by_name(sheet_name)
    titles = data.read_titles(workbook_path, sheet_name)
    return sheet, titles

def process_row(d, titles, row, shared_results):
    """
    处理单行数据，创建检验批文档及相关表格
    """
    child_name = d['child_name']
    serial_number = d['true_number']
    title = d['title']

    # 找到并读取模板
    template_path = data.find_template_xlsx(config, 'JYP_template', serial_number[0:5], title)
    excel_template = data.switch_open_excel_template(model_type, template_path)

    # 修改单元格值
    _update_merged_cell_value(excel_template.worksheets[0], 'U5', "技术负责人")

    # 保存路径处理与目录创建
    save_path_template = pathlib.Path(config.get('default', 'outPutDir')).joinpath(child_name).joinpath(
        data.find_save_path(config, 'JYP_outputPath', serial_number[0:5]))
    pathlib.Path.mkdir(save_path_template, parents=True, exist_ok=True)

    # 写入并保存模板
    save_file = save_path_template.joinpath(serial_number + child_name + title + '.xlsx')
    data.switch_write_excel_template(model_type, excel_template, cell_mapping, d, save_file)

    # 记录结果
    shared_results.append(save_path_template)

    # 创建施工检查记录和其他相关文档
    _create_construction_inspection_record(model_type, data_mapping_record, titles, row, serial_number,
                                           child_name, title, save_path_template)
    _generate_engineering_inspection_form(data_mapping, config, serial_number, child_name, title, save_path_template)


def _update_merged_cell_value(sheet, cell_address, value):
    """
    更新合并单元格的值
    """
    merged_cell = sheet[cell_address].offset(row=0, column=-1)
    if isinstance(merged_cell, MergedCell):
        for merged_range in sheet.merged_cell_ranges:
            if merged_cell.coordinate in merged_range:
                top_left_cell = sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
                top_left_cell.value = value


def _create_construction_inspection_record(model_type, data_mapping_record, titles, row, serial_number,
                                           child_name, title, save_path_template):
    """
    创建施工检查记录
    """
    data.construction_inspection_record(model_type, data_mapping_record, titles, row, serial_number,
                                        child_name, title, save_path_template)


def _generate_engineering_inspection_form(data_mapping, config, serial_number, child_name, title, save_path_template):
    """
    生成工程报验表
    """
    template_path = setting.get_word_template('inspection_lot\报验表\报验表.docx')
    form_number = serial_number[:6] + "C2" + serial_number[8:]
    check_parts = child_name + d.get('acceptance_part') + d.get('sub_project_name')
    inspection_lot = d.get('sub_project_name')
    save_file = save_path_template.joinpath(serial_number + child_name + title + '报验表.docx')

    context = {'BYBBianHao': form_number, 'YinJianBuWei': check_parts, 'JYP': inspection_lot, 'BianHao': serial_number}
    data.write_word_template(template_path, context, save_file)


if __name__ == '__main__':
    start = time.time()
    pross = 12
    pool = mp.Pool(processes=pross)
    shared_results = mp.Manager().list()

    sheet_names = config.get('default', 'sheets').split(',')

    for sheet_name in sheet_names:
        sheet, titles = load_data_sheet(sheet_name)

        for i in range(1, sheet.nrows):
            row = sheet.row_values(i)
            d = data.get_object(data_mapping, titles, row)
            d.update({
                "general_contractor": "刚果（金）卡隆威矿业有限公司",
                "turnkey_project_manager": "高磊",
                "subcontractor": "十五冶对外工程有限公司",
                "subcontracting_project_manager": "冯齐名"
            })

            if not data.whether_to_submit(d):
                pool.apply_async(process_row, args=(d, titles, row, shared_results))

    pool.close()
    pool.join()
    data.close_excel()
    data.make_pdf(set(shared_results))
    logger.success('共耗时{t} ms', t=time.time() - start)
