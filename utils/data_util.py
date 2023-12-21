
import datetime
import pathlib
import xlrd
import xlwings
from loguru import logger
import utils.office2pdf as o2p
from resources import setting_util as setting

# 将Excel的数字日期转化为标准日期
@logger.catch
def int_to_date(number: str):
    """
    将Excel的数字日期转化为标准日期

    参数:
    number (str): 需要转化的日期字符串

    返回:
    str: 转化后的标准日期字符串
    """
    dt = float(number)
    delta = datetime.timedelta(days=dt)
    today = datetime.datetime.strptime('1899-12-30', '%Y-%m-%d') + delta
    strftime = datetime.datetime.strftime(today, '%Y-%m-%d')
    logger.debug(strftime)
    return strftime


# 读取标题行:
# workbook:要读取的工作簿（数据源）
# sheetName：工作簿名称（数据源）
@logger.catch
def read_titles(workbook, sheetName):
    """
    读取标题行

    参数:
    workbook (str): 要读取的工作簿路径
    sheetName (str): 工作簿名称

    返回:
    list: 标题行数据
    """
    wb = xlrd.open_workbook(workbook)
    sheet = wb.sheet_by_name(sheetName)
    return sheet.row(0)


# 获得标题所对应的列号
# titleName:标题所在的列号
@logger.catch
def get_title_index(title_list, title_name):
    """
    获得标题所对应的列号

    参数:
    title_list (list): 标题列表
    title_name: 标题名称

    返回:
    int: 标题在列表中的列号
    """
    for i in range(title_list.__len__()):
        if title_list.__getitem__(i).value == title_name:
            return i


# 如果文件存在，删除该文件（python不支持自动覆盖，有可能是没找到）
@logger.catch
def remove_file(filePath):
    """
    如果文件存在，删除该文件

    参数:
    filePath (str): 文件路径

    返回:
    None
    """
    if pathlib.Path(filePath).exists():
        pathlib.Path(filePath).unlink()
    else:
        pass


def close_excel():
    """
    关闭Excel

    返回:
    None
    """
    # 关闭Excel
    if xlwings.apps.__len__() > 0:
        for app in xlwings.apps:
            app.kill()
            logger.debug("{id}已关闭", id=app.pid)


@logger.catch
def make_pdf(s: set):
    """
    将文件合成为PDF

    参数:
    s (set): 需要合成PDF的文件集合

    返回:
    None
    """
    logger.info('开始合成PDF……')
    for dir in s:
        logger.debug('dir:{d}', d=dir)
        o2p.run(dir)


# 查询表格模板
@logger.catch
def find_template(config, groupName, itemName, title) -> pathlib.Path:
    """
    查询表格模板

    参数:
    config (obj): 配置文件对象
    groupName (str): 组名
    itemName (str): 项名
    title (str): 标题

    返回:
    pathlib.Path: 表格模板路径
    """
    if config.has_section(groupName) & config.has_option(groupName, itemName):
        return pathlib.Path(config.get(groupName, itemName) + '/' + title + '.xls')


# 查询表格模板
@logger.catch
def find_template_xlsx(config, groupName, itemName, title) -> pathlib.Path:
    """
    查询表格模板

    参数:
    config (obj): 配置文件对象
    groupName (str): 组名
    itemName (str): 项名
    title (str): 标题

    返回:
    pathlib.Path: 表格模板路径
    """
    if config.has_section(groupName) & config.has_option(groupName, itemName):
        return pathlib.Path(config.get(groupName, itemName) + '/' + title + '.xlsx')


# 查询保存路径
@logger.catch
def find_save_path(config, groupName, itemName) -> pathlib.Path:
    """
    查询保存路径

    参数:
    config (obj): 配置文件对象
    groupName (str): 组名
    itemName (str): 项名

    返回:
    pathlib.Path: 保存路径
    """
    if config.has_section(groupName) & config.has_option(groupName, itemName):
        return pathlib.Path(config.get(groupName, itemName))


# 如果文件存在，删除该文件（python不支持自动覆盖，有可能是没找到）
@logger.catch
def remove_file(filePath):
    """
    如果文件存在，删除该文件

    参数:
    filePath (str): 文件路径

    返回:
    None
    """
    if pathlib.Path(filePath).exists():
        pathlib.Path(filePath).unlink()
    else:
        pass


# 根据配置文件获取数据库模型，并生成字典
@logger.catch
def get_object(group_name: str, titles, row_data) -> dict:
    """
    根据配置文件获取数据库模型，并生成字典

    参数:
    group_name (str): 组名
    titles (list): 标题列表
    row_data: 数据行

    返回:
    dict: 字典对象
    """
    # 获取配置文件，en_us为中英文映射，不可修改
    config = setting.get_config('setting.cfg')
    en_us = setting.get_config('en_us.cfg')
    options = config.options(group_name)
    d = dict()
    for option in options:
        # 根据配置文件获取字段名
        attr = config.get(group_name, option)
        en = en_us.get('default', option)
        index = get_title_index(titles, attr)
        if None == index: continue
        value = str(row_data[index])

        d.setdefault(en, value)
    return d


# 生成Word
@logger.catch
def write_word_template(word_template, data, save_file):
    """
    生成Word

    参数:
    word_template (obj): Word模板对象
    data (dict): 数据字典
    save_file (str): 保存文件路径

    返回:
    None
    """
    word_template.render(data)
    word_template.save(save_file)


# 使用xlwings填充Excel模板
@logger.catch
def write_excel_template_with_xlwings(excel_template, group_name, data: dict, save_file):
    """
    使用xlwings填充Excel模板

    参数:
    excel_template (obj): Excel模板对象
    group_name (str): 组名
    data (dict): 数据字典
    save_file (str): 保存文件路径

    返回:
    None
    """
    config = setting.get_config('setting.cfg')
    en_us = setting.get_config('en_us.cfg')
    options = config.options(group_name)
    for option in options:
        cell_address = config.get(group_name, option)
        en = en_us.get('default', option)
        excel_template.sheets[0].range(cell_address).value = data.get(en)
    remove_file(save_file)
    excel_template.save(save_file)
    excel_template.close()


# 使用openpyxl填充Excel模板
@logger.catch
def write_excel_template_with_openpyxl(excel_template, group_name, data: dict, save_file):
    """
    使用openpyxl填充Excel模板

    参数:
    excel_template (obj): Excel模板对象
    group_name (str): 组名
    data (dict): 数据字典
    save_file (str): 保存文件路径

    返回:
    None
    """
    config = setting.get_config('setting.cfg')
    en_us = setting.get_config('en_us.cfg')
    options = config.options(group_name)
    for option in options:
        cell_address = config.get(group_name, option)
        en = en_us.get('default', option)
        excel_template.worksheets[0][cell_address].value = data.get(en)
    remove_file(save_file)
    excel_template.save(save_file)


@logger.catch
def switch_open_excel_template(model_type, templatepath):
    """
    切换打开的Excel模板

    参数:
    model_type (str): 模型类型
    templatepath (str): 模板路径

    返回:
    obj: Excel模板对象
    """
    if model_type == 'openpyxl':
        return setting.openExcelTemplateWithOpenpyxl(templatepath)
    elif model_type == 'xlwings':
        return setting.openExcelTemplateWithXlwings(templatepath)


@logger.catch
def switch_write_excel_template(model_type, excel_template, group_name, da: dict, save_file):
    """
    切换写入的Excel模板

    参数:
    model_type (str): 模型类型
    excel_template (obj): Excel模板对象
    group_name (str): 组名
    da (dict): 数据字典
    save_file (str): 保存文件路径

    返回:
    None
    """
    if model_type == 'openpyxl':
        write_excel_template_with_openpyxl(excel_template, group_name, da, save_file)
    elif model_type == 'xlwings':
        write_excel_template_with_xlwings(excel_template, group_name, da, save_file)


# 施工检查记录
@logger.catch
def construction_inspection_record(model_type, group_name: str, titles, row, serial_number, child_name, title,
                                   save_path_template):
    """
    施工检查记录

    参数:
    model_type (str): 模型类型
    group_name (str): 组名
    titles (list): 标题列表
    row (int): 数据行数
    serial_number (str): 序列号
    child_name (str): 子名称
    title (str): 标题
    save_path_template (obj): 保存路径模板对象

    返回:
    None
    """
    construction_inspection_record_data = get_object(group_name, titles, row)
    logger.debug(construction_inspection_record_data)
    # construction_inspection_record_data['check_date'] = int_to_date(
    #     construction_inspection_record_data.get('check_date'))
    construction_inspection_record_template = switch_open_excel_template(model_type,
                                                                         'inspection_lot/施工检查记录/施工检查记录.xlsx')
    construction_inspection_record_save_file = save_path_template.joinpath(
        serial_number + child_name + title + '施工检查记录.xlsx')
    switch_write_excel_template(model_type, construction_inspection_record_template, group_name, construction_inspection_record_data,
                                construction_inspection_record_save_file)


# 是否报送
@logger.catch
def whether_to_submit(d: dict):
    """
    判断是否需要报送

    参数:
    d (dict): 数据字典

    返回:
    bool: 判断结果
    """
    if None == d: return True
    result: str = d.get('whether_to_submit')
    if result.lower() == '1': return True
    if result.lower() == '0':
        return False
    else:
        return True