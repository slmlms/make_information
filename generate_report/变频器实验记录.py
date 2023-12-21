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
log.to_log(config, '变频器试验记录')
# 数据源路径
workbook_path = config.get('default', 'DataSource')
# 数据源工作表
sheet = xlrd.open_workbook(workbook_path).sheet_by_name('变频器试验记录5wt')
# 标题行
titles = data.read_titles(workbook_path, '变频器试验记录5wt')
# 保存路径
save_path_template = pathlib.Path(config.get('default', 'inverterOutPutDir'))

mapping = 'inverter'

s = set()
for i in range(1, sheet.nrows):
    # 模板路径，模板要放到循环里面，不然会生成同一个文件
    word_template = setting.get_word_template('inspection_lot\实验报告\变频器实验报告模板.docx')
    row = sheet.row_values(i)
    d = data.get_object(mapping, titles, row)
    logger.debug(d)
    if data.whether_to_submit(d) == True:
        continue
    # 修改时间格式
    d['inverter_date_of_manufacture'] = data.int_to_date(d.get('inverter_date_of_manufacture'))
    if not pathlib.Path(save_path_template.joinpath(d["child_name"])).exists():
        pathlib.Path.mkdir(pathlib.Path(save_path_template).joinpath(d["child_name"]), True, True)
    logger.debug(save_path_template)
    data.write_word_template(word_template, d,
                             save_path_template.joinpath(d["child_name"]).joinpath(d["acceptance_part"] + '.docx'))
    s.add(save_path_template.joinpath(d["child_name"]))
data.close_excel()
data.make_pdf(s)
