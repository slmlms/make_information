import configparser
import os

import docxtpl
import openpyxl
import xlwings
from loguru import logger


@logger.catch
def get_config(config_path):
    """
    获取配置信息

    Args:
        config_path (str): 配置文件路径

    Returns:
        configparser.ConfigParser: 配置对象
    """
    current_path = os.path.dirname(__file__)
    config = configparser.ConfigParser()
    config.read(current_path + '/' + config_path, encoding='UTF-8-sig')
    return config


@logger.catch
def get_word_template(template_path):
    """
    获取Word模板

    Args:
        template_path (str): 模板文件路径

    Returns:
        docxtpl.DocxTemplate: Word模板对象
    """
    current_path = os.path.dirname(__file__)
    return docxtpl.DocxTemplate(current_path + '/' + template_path)


# 打开Excel模板
@logger.catch
def openExcelTemplateWithXlwings(templatepath):
    """
    使用xlwings打开Excel模板

    Args:
        templatepath (str): 模板文件路径

    Returns:
        xlwings.Book: 打开的Excel模板对象
    """
    current_path = os.path.dirname(__file__)
    app = xlwings.App(visible=False, add_book=False)
    app.screen_updating = False
    return app.books.open(current_path + '/' + str(templatepath))


# 打开Excel模板
@logger.catch
def openExcelTemplateWithOpenpyxl(templatepath):
    """
    使用openpyxl打开Excel模板

    Args:
        templatepath (str): 模板文件路径

    Returns:
        openpyxl.Workbook: 打开的Excel模板对象
    """
    current_path = os.path.dirname(__file__)
    return openpyxl.load_workbook(current_path + '/' + str(templatepath))


# 关闭Excel模板
@logger.catch
def closeExcelTemplate(templateWorkBook):
    """
    关闭Excel模板

    Args:
        templateWorkBook: Excel模板对象
    """
    templateWorkBook.app.quit()
    logger.info('Excel模板已关闭')