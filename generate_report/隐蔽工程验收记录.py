# 检验批质量验收记录
import pathlib

import xlrd
from loguru import logger

import utils.data_util as data
import utils.log_util as log
from resources import setting_util as setting

# 选择的模块类型，只能有openpyxl和xlwings
model_type: str = 'xlwings'
# 映射
cell_mapping = 'YB_CellMapping'
data_mapping = 'YB_DataMapping'
# 施工检查记录模型
data_mapping_jyp = 'SGJCJL_DataMapping_YB'
# 获取配置文件
config = setting.get_config('setting.cfg')
logger.debug(config.sections())
# 记录日志
log.to_log(config, '隐蔽')
# 数据源路径
workbook_path = config.get('default', 'DataSource')

s = set()

# 数据源工作表
name = '隐蔽工程'
sheet = xlrd.open_workbook(workbook_path).sheet_by_name(name)
# 标题行
titles = data.read_titles(workbook_path, name)
logger.debug(titles)

for i in range(1, sheet.nrows):
    logger.info('正在制作第{i}行，已完成{b}%', i=i, b=i / sheet.nrows * 100)
    row = sheet.row_values(i)
    d = data.get_object(data_mapping, titles, row)
    logger.debug(d)
    if data.whether_to_submit(d) == True:
        continue
    # 子项名称
    child_name = d.get('child_name')
    # 检验批编号
    serial_number = d.get('true_number')
    # 检验批标题
    title = '隐蔽工程验收记录'
    print(child_name, serial_number, title)
    # 检验批模板路径
    template_path = 'inspection_lot\隐蔽工程验收记录\隐蔽工程验收记录.xlsx'
    logger.info("隐蔽工程模板路径为{p}", p=template_path)
    # Excel模板，注意选择打开方式
    excel_template = data.switch_open_excel_template(model_type, template_path)
    # 保存模板路径
    save_path_template = pathlib.Path(config.get('default', 'yinBiOutPutDir')).joinpath(child_name)
    if not pathlib.Path(save_path_template).exists():
        pathlib.Path.mkdir(pathlib.Path(save_path_template), True, True)
    logger.debug(save_path_template)
    # 最终保存的文件
    save_file = save_path_template.joinpath(serial_number + child_name + title + '资料.xlsx')
    # 写入模板
    data.switch_write_excel_template(model_type, excel_template, '%s' % cell_mapping, d, save_file)
    s.add(save_path_template)

    # # 施工检查记录
    # data.construction_inspection_record(model_type, data_mapping_jyp, titles, row, serial_number, child_name, title,
    #                                     save_path_template)
    # # 工程报验表
    # engineering_inspection_form_template = setting.get_word_template('inspection_lot\报验表\报验表.docx')
    # engineering_inspection_form_number = serial_number[:6] + "C2" + serial_number[8:]
    # check_parts = d.get('acceptance_part')
    # inspection_lot = '隐蔽工程验收记录'
    # engineering_inspection_form_save_file = save_path_template.joinpath(
    #     serial_number + child_name + title + '报验表.docx')
    # context = {'BYBBianHao': engineering_inspection_form_number, 'YinJianBuWei': check_parts, 'JYP': inspection_lot,
    #            'BianHao': serial_number}
    # data.write_word_template(engineering_inspection_form_template, context, engineering_inspection_form_save_file)

data.close_excel()
data.make_pdf(s)
