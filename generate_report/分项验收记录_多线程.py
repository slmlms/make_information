# 检验批质量验收记录
import multiprocessing as mp
import os
import pathlib
import time

import xlrd
from loguru import logger

import utils.data_util as data
import utils.log_util as log
from resources import setting_util as setting

# 选择的模块类型，只能有openpyxl和xlwings
model_type: str = 'openpyxl'

data_mapping = 'FX_DataMapping'
cell_mapping = 'FX_CellMapping'
# 获取配置文件
config = setting.get_config('setting.cfg')
logger.debug(config.sections())
# 记录日志
log.to_log(config, '分项')
# 数据源路径
workbook_path = config.get('default', 'DataSource')
sheet_name = "分项工程-电气"


@logger.catch()
def run(d, s: list):
    logger.debug(os.getpid())
    # logger.info('正在制作第{i}行，已完成{b}%', i=i, b=i / sheet.nrows * 100)
    # 子项名称
    child_name = d.get('child_name')
    # 编号
    serial_number = "S2-" + "{:0>2d}".format(int(float(d.get('sub_project_code'))))
    # 标题
    title = d.get('sub_project_name')
    print(child_name, serial_number, title)
    # 将检验批区段分割为列表
    value = d.get('inspection_lot_content').split(';')
    # 模板路径
    template_path = pathlib.Path("inspection_lot/分项//" + sheet_name + ".xlsx")
    # Excel模板，注意选择打开方式
    excel_template = data.switch_open_excel_template(model_type, template_path)
    # 保存模板路径
    save_path_template = pathlib.Path(config.get('default', 'outPutDir') + sheet_name + "\\")
    if not pathlib.Path(save_path_template).exists():
        pathlib.Path.mkdir(pathlib.Path(save_path_template), True, True)
    logger.debug(save_path_template)
    # 最终保存的文件
    save_file = save_path_template.joinpath(child_name + serial_number + title + '.xlsx')
    # 写入模板
    # excel_template.sheets[0].range('B2').value = title + '分项工程质量验收记录'
    # excel_template.sheets[0].range('G6').value = "刚果(金)KALONGWE铜钴矿采冶项目" + child_name
    # excel_template.sheets[0].range('S6').value = d.get('structure_type')
    # excel_template.sheets[0].range('AA6').value = d.get('quantity_of_inspection_lot')

    excel_template.active['B2'] = title + '分项工程质量验收记录'
    excel_template.active['G6'] = "刚果(金)盛屯矿业KALONGWE5万吨/年阴极铜扩建项目\t" + child_name
    excel_template.active['S6'] = d.get('structure_type')
    excel_template.active['AA6'] = d.get('quantity_of_inspection_lot')

    num = 11
    for v in value:
        if len(value) > 13:
            continue
        # excel_template.sheets[0].range('D' + str(num)).value = v
        # excel_template.sheets[0].range('M' + str(num)).value = "合格"

        excel_template.active['D' + str(num)] = v
        excel_template.active['M' + str(num)] = "合格"
        # 先杰出合并单元格，不然后续合并的时候会报错
        excel_template.active.unmerge_cells("V" + str(num) + ":AD" + str(num))
        num = num + 1
    # 合并单元格
    # excel_template.sheets[0].range('v11:v' + str(num - 1)).api.merge()
    excel_template.active.merge_cells("V11:AD" + str(num - 1))
    data.remove_file(save_file)
    excel_template.save(save_file)
    excel_template.close()
    s.append(save_path_template)

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
    pross = 8
    pool = mp.Pool(processes=pross)
    # 全局共享变量
    s = mp.Manager().list()

    # 数据源工作表
    sheet = xlrd.open_workbook(workbook_path).sheet_by_name(sheet_name)
    titles = data.read_titles(workbook_path, sheet_name)
    # 标题行
    logger.debug(titles)
    for i in range(1, sheet.nrows):
        row = sheet.row_values(i)
        d = data.get_object(data_mapping, titles, row)
        logger.debug(d)
        if data.whether_to_submit(d) == True or d.get('quantity_of_inspection_lot') == '0' or d.get(
                'quantity_of_inspection_lot') == "":
            continue

        # 使用异步方法并行处理，多线程尽量不要传入文件作为参数
        pool.apply_async(run, args=(d, s))

    # 先结束进程池，后join，否则会报错
    pool.close()
    pool.join()
    data.close_excel()
    data.make_pdf(set(s))
    logger.success('共耗时{t} ms', t=time.time() - start)
