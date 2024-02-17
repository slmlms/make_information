import concurrent.futures
import gc
import os
import threading

import pythoncom
import win32com.client
import win32com.client
from PyPDF2 import PdfFileMerger
from loguru import logger
from tqdm import tqdm


def word2Pdf(filePath, words):
    # 如果没有文件则提示后直接退出
    if len(words) < 1:
        logger.warning("\n【无 Word 文件】\n")
        return
    # 开始转换
    try:
        pdfs = []
        lock = threading.Lock()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = 0
        word.DisplayAlerts = False
        doc = None
        for i in tqdm(range(len(words))):
            with lock:
                logger.debug(i)
                fileName = words[i]  # 文件名称
                fromFile = os.path.join(filePath, fileName)  # 文件地址
                toFileName = changeSufix2Pdf(fileName)  # 生成的文件名称
                toFile = toFileJoin(filePath, toFileName)  # 生成的文件地址
                pdfs.append(toFile)
                # 某文件出错不影响其他文件打印
                try:
                    doc = word.Documents.Open(fromFile)
                    doc.SaveAs(toFile, 17)  # 生成的所有 PDF 都会在 PDF 文件夹中
                except Exception as e:
                    logger.exception(e)
                # 关闭 Word 进程
        doc.Close()
        word.Quit()
        return sort_docx_files(pdfs)
    except Exception as e:
        logger.warning(e)
    finally:
        gc.collect()


@logger.catch
def changeSufix2Pdf(file):
    return file[:file.rfind('.')] + ".pdf"


@logger.catch
def toFileJoin(filePath, file):
    return os.path.join(filePath, 'pdf', file[:file.rfind('.')] + ".pdf")


def sort_docx_files(docx_files):
    # 定义排序规则
    sort_order = {
        "分部工程质量检验评定记录-报验表": 5,
        "分部工程质量检验评定记录": 4,
        "分项质量检查验收记录-报验表": 3,
        "分项质量检查验收记录": 2,
        "检验批质量检验评定记录": 1
    }

    # 按照指定顺序排序
    return sorted(docx_files, key=lambda x: sort_order.get(os.path.basename(x), 0), reverse=True)


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


def list_folders(directory):
    # 遍历目录下的所有文件和文件夹
    contents = os.listdir(directory)
    # 筛选出文件夹，并返回文件夹名称的集合
    return {item for item in contents if os.path.isdir(os.path.join(directory, item))}


def create_and_list_pdf_files(root_folder):
    # 遍历根目录及其子目录下的所有.docx文件，但跳过以“~$”开头的临时文件
    docx_files = [os.path.join(root, file) for root, dirs, files in os.walk(root_folder)
                  for file in files
                  if file.endswith('.docx') and not file.startswith('~$')]
    # 按指定顺序排序
    sorted_docx_files = sort_docx_files(docx_files)
    pdfs = word2Pdf(root_folder, sorted_docx_files)
    if pdfs is None:
        logger.warning("【无 PDF 文件生成】")
        return []
    return pdfs


if __name__ == "__main__":
    # 指定目录路径
    directory_path = 'D:\Jobs\卡莫亚\检验批及分项\钢结构检验批生成'

    # 获取文件夹名称的集合
    folders = list_folders(directory_path)
    for folder in tqdm(folders):
        print(folder)
        # 指定目录及子目录下的所有.docx文件
        root_folder = os.path.join(directory_path, folder)
        # 指定合并后保存的文件名
        output_folder = os.path.join(directory_path, folder, folder+".pdf")
        # 执行主函数
        pdfs = create_and_list_pdf_files(root_folder)
        merge_pdfs(pdfs, output_folder)

# def process_folder(folder,directory_path):
#     root_folder = os.path.join(directory_path, folder)
#     pdfs = create_and_list_pdf_files(root_folder)
#     output_folder = os.path.join(directory_path, folder, folder + ".pdf")
#     if pdfs is not None:
#         merge_pdfs(pdfs, output_folder)
#
#     # 确保返回一个包含 PDF 列表（即使为空）和文件夹名的元组
#     return pdfs, folder
#
# def process_word2pdf_in_thread(folder, directory_path):
#     pythoncom.CoInitialize()
#
#     try:
#         root_folder = os.path.join(directory_path, folder)
#         output_folder = os.path.join(directory_path, folder, folder + ".pdf")
#         pdfs = create_and_list_pdf_files(root_folder)
#         merge_pdfs(pdfs, output_folder)
#     except Exception as e:
#         logger.warning(str(e))
#     finally:
#         pythoncom.CoUninitialize()
#
#
# if __name__ == "__main__":
#     # 指定目录路径
#     directory_path = 'D:\Jobs\卡莫亚\检验批及分项\钢结构检验批生成'
#
#     # 获取文件夹名称的集合
#     folders = list_folders(directory_path)
#
#     with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
#         futures = {executor.submit(process_word2pdf_in_thread, folder, directory_path): folder for folder in folders}
