import pathlib

import xlrd
from loguru import logger

import utils.data_util as data
import utils.log_util as log
from resources import setting_util as setting

# 获取配置文件
config = setting.get_config('setting.cfg')
logger.debug(config.sections())
# 记录日志
log.to_log(config, '电流互感器实验记录')
# 数据源路径
workbook_path = config.get('default', 'DataSource')
# 数据源工作表
sheet = xlrd.open_workbook(workbook_path).sheet_by_name('电流互感器实验记录')
# 标题行
titles = data.read_titles(workbook_path, '电流互感器实验记录')
# 保存路径
save_path_template = pathlib.Path(config.get('default', 'currentTransformerOutPutDir'))

mapping = 'current_transformer'

s = set()
for i in range(1, sheet.nrows):
    # 模板路径，模板要放到循环里面，不然会生成同一个文件
    word_template = setting.get_word_template('inspection_lot\实验报告\电流互感器实验报告模板.docx')
    row = sheet.row_values(i)
    d = data.get_object(mapping, titles, row)
    logger.debug(d)
    if data.whether_to_submit(d) == True:
        continue
    # 修改时间格式
    # d['date_of_manufacture'] = data.int_to_date(d.get('date_of_manufacture'))
    if not pathlib.Path(save_path_template).exists():
        pathlib.Path.mkdir(pathlib.Path(save_path_template), True, True)
    logger.debug(save_path_template)
    data.write_word_template(word_template, d, save_path_template.joinpath(str(i) + '.docx'))
    s.add(save_path_template)
data.close_excel()
data.make_pdf(s)
