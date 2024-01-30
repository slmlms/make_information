import gc
import os
import shutil
import tempfile
from datetime import time

import win32com
from PyPDF2 import PdfFileMerger
from loguru import logger
from win32com.client import Dispatch


def is_file_locked(file):
    # 尝试打开文件以检测其是否被锁定
    try:
        with open(file, 'r+b'):
            return False  # 如果可以成功打开，则文件未被锁定
    except IOError as e:
        if "Permission denied" in str(e):
            return True  # 如果出现权限错误，表示文件被锁定
        else:
            raise e  # 其他IO错误则直接抛出


def convert_word_to_pdf(word_path, temp_dir):
    pdf_path = os.path.join(temp_dir, os.path.basename(word_path) + ".pdf")


    word = Dispatch("Word.Application")
    doc = word.Documents.Open(word_path)

    # 确保转换后Word文档不显示
    word.Visible = False

    # 设置输出格式为PDF
    doc.SaveAs(pdf_path, FileFormat=17)  # 17代表wdFormatPDF for Word 2007+

    doc.Close()
    word.Quit()


def merge_word_to_pdf(input_paths, output_path):
    # 创建临时目录
    temp_dir = tempfile.mkdtemp()

    # 转换每个Word文档到PDF并保存到临时目录
    pdf_files = []
    for word_path in input_paths:
        pdf_path = convert_word_to_pdf(word_path, temp_dir)
        pdf_files.append(pdf_path)

    # 合并PDF文件
    merger = PdfFileMerger()
    for pdf_file in pdf_files:
        merger.append(pdf_file)

    # 保存合并后的PDF文件
    with open(output_path, 'wb') as fh:
        merger.write(fh)

    # 删除合并前的临时PDF文件和临时目录
    for pdf_file in pdf_files:
        os.remove(pdf_file)
    shutil.rmtree(temp_dir)



Jianyanpi_save_path = "D:\Jobs\卡莫亚\检验批及分项\钢结构检验批生成\\"
# 使用os.scandir()方法，它提供更好的性能和额外信息
for entry in os.scandir(Jianyanpi_save_path):
    if entry.is_dir() and not entry.is_symlink():  # 排除符号链接
        print(entry.name)  # 输出文件夹名称
        # 遍历文件夹及子文件夹下所有的docx文件，并根据路径排序
        documents_order = []
        out_path = Jianyanpi_save_path + entry.name + ".pdf"
        for root, dirs, files in os.walk(entry.path):
            # 按照文件路径排序
            files.sort(key=lambda x: x.lower())
            for file in files:
                if file.endswith(".docx"):
                    print(file)
                    documents_order.append(os.path.join(root, file))

                    print(documents_order)
            merge_word_to_pdf(documents_order, out_path)
