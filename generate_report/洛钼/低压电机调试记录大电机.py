import gc
import os
import pathlib
import subprocess

import openpyxl
import xlrd
from PyPDF2 import PdfFileMerger
from loguru import logger
from tqdm import tqdm

import utils.data_util as data

# 选择的模块类型，只能有openpyxl和xlwings
# model_type: str = 'openpyxl'

save_path = pathlib.Path('D:\Jobs\洛钼\调试记录\低压电机\大电机\\')
work_book_path = "D:\Jobs\洛钼\调试记录\低压电机\电动机数据源.xlsx"

excel_template_path = "D:\Jobs\洛钼\调试记录\低压电机\\3..2.2 100kW及以上低压电动机.xlsx"

# 存储生成的 Excel 文件路径的列表
generated_files = []


def toFileJoin(filePath, file):
    return os.path.join(filePath, 'pdf', file[:file.rfind('.')] + ".pdf")


@logger.catch
def excel2Pdf(filePath, excels):
    # 如果没有文件则提示后直接退出
    if len(excels) < 1:
        logger.warning("\n【无 Excel 文件】\n")
        return
    # 开始转换
    logger.info("\n【开始 Excel -> PDF 转换】")
    try:
        pdfs = []
        logger.info("打开 Excel 进程中...")
        # excel = win32com.client.Dispatch("Excel.Application")
        # excel.Visible = 0
        # excel.DisplayAlerts = False

        for i in range(len(excels)):
            logger.debug(i)
            fileName = excels[i]  # 文件名称
            fromFile = os.path.join(filePath, fileName)  # 文件地址

            logger.info("转换：" + fileName + "文件中...")
            # 某文件出错不影响其他文件打印
            try:
                cmd = ["D:\Software\Libre Offices\program\soffice.com", "--headless", "--convert-to", "pdf", fromFile,
                       "--outdir", filePath + "\\"]

                subprocess.run(cmd, encoding="utf-8")
                pdfs.append(fromFile.replace("xlsx", "pdf"))
                # wb = excel.Workbooks.Open(fromFile)
                # for j in range(1):  # 工作表数量，一个工作簿可能有多张工作表
                #     toFileName = addWorksheetsOrder(fileName)  # 生成的文件名称
                #     toFile = toFileJoin(filePath, toFileName)  # 生成的文件地址
                #
                #     ws = wb.Worksheets(j + 1)  # 若为[0]则打包后会提示越界
                #     ws.ExportAsFixedFormat(0, toFile)  # 每一张都需要打印
                #     logger.success("转换至：" + toFileName + "文件完成")
                #     pdfs.append(toFile)
            except Exception as e:
                logger.exception(e)
        # 关闭 Excel 进程
        logger.success("所有 Excel 文件已打印完毕")
        logger.success("结束 Excel 进程中...\n")
        # close_excel_by_force(excel)
        return pdfs
    except Exception as e:
        logger.exception(e)
    finally:
        gc.collect()


def addWorksheetsOrder(file):
    return file[:file.rfind('.')] + ".pdf"


def close_excel_by_force(excel):
    import win32process
    import win32api
    import win32con
    # Get the window's process id's
    hwnd = excel.Hwnd
    t, p = win32process.GetWindowThreadProcessId(hwnd)
    # Ask window nicely to close
    try:
        handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, 0, p)
        if handle:
            win32api.TerminateProcess(handle, 0)
            win32api.CloseHandle(handle)
    except:
        pass


def merge_pdfs(pdf_files, output_pdf):
    # 创建PdfFileMerger对象
    merger = PdfFileMerger()

    # 合并PDF文件
    logger.info("开始合并文件...")
    for pdf_file in pdf_files:
        merger.append(pdf_file)

    # 保存合并后的PDF文件
    merger.write(output_pdf)
    merger.close()

    # 删除转换的PDF文件
    for pdf_file in pdf_files:
        os.remove(pdf_file)
    logger.info("合并完成！")


def create_and_list_pdf_files(root_folder):
    # 遍历根目录及其子目录下的所有.docx文件，但跳过以“~$”开头的临时文件
    docx_files = [os.path.join(root, file) for root, dirs, files in os.walk(root_folder)
                  for file in files
                  if file.endswith('.xlsx') and not file.startswith('~$')]
    pdfs = excel2Pdf(root_folder, docx_files)
    if pdfs is None:
        logger.warning("【无 PDF 文件生成】")
        return []
    return pdfs


def run(sheet):
    logger.info("当前进程id：", os.getpid())
    # Excel模板，注意选择打开方式
    excel_template = openpyxl.load_workbook(excel_template_path)

    for i in range(sheet.nrows):
        if i == 0:
            continue
        row = sheet.row_values(i)
        # 生成的文件夹名称为row[0]的部分
        folder_name = ''.join(filter(lambda x: not x.isdigit(), row[0]))
        folder_path = save_path.joinpath(folder_name)
        folder_path.mkdir(parents=True, exist_ok=True)  # 创建文件夹
        if row[4] > 30:
            excel_template.worksheets[0]["C3"].value = row[1]
            excel_template.worksheets[0]["N3"].value = row[0]
            excel_template.worksheets[0]["C5"].value = row[3]
            excel_template.worksheets[0]["L5"].value = row[4]
            excel_template.worksheets[0]["C6"].value = row[5]
            excel_template.worksheets[0]["L6"].value = row[6]
            excel_template.worksheets[0]["C7"].value = row[7]
            excel_template.worksheets[0]["L7"].value = row[8]
            excel_template.worksheets[0]["C8"].value = row[10]
            excel_template.worksheets[0]["L8"].value = row[9]
            excel_template.worksheets[0]["C9"].value = row[11]

            # 生成一个唯一的文件名，以F4单元格的内容作为文件名
            file_name = f"{row[0]}{row[1]}"
            file_name = file_name.replace("/", "_")  # 替换特殊字符
            save_file = folder_path.joinpath(file_name + '.xlsx')
            generated_files.append(save_file)  # 将生成的文件路径添加到列表中
            excel_template.save(save_file)


def list_folders(directory):
    # 遍历目录下的所有文件和文件夹
    contents = os.listdir(directory)
    # 筛选出文件夹，并返回文件夹名称的集合
    return {item for item in contents if os.path.isdir(os.path.join(directory, item))}


# 使用LibreOffice的API将所有文件转换为PDF并合并
# 将文件转换为PDF
def convert_to_pdf(input_file, output_file):
    cmd = ["D:\Software\Libre Offices\program\soffice.com", "--headless", "--convert-to", "pdf", input_file, "--outdir",
           os.path.dirname(output_file)]
    subprocess.run(cmd)


if __name__ == '__main__':
    excel_instance = None  # 全局变量保存Excel实例
    wb = xlrd.open_workbook(work_book_path)
    sheet = wb.sheet_by_name("电机数据")

    run(sheet)

    data.close_excel()

    # 在循环结束后保存所有生成的 Excel 文件
    folders = list_folders(save_path)
    for folder in tqdm(folders):
        root_folder = os.path.join(save_path, folder)
        output_folder = os.path.join(save_path, folder, folder + ".pdf")
        pdfs = create_and_list_pdf_files(root_folder)
        merge_pdfs(pdfs, output_folder)
