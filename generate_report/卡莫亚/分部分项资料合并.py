import gc
import os
import subprocess
from pathlib import Path

import fitz  # 导入PyMuPDF以处理空白页
import win32com.client
from PyPDF2 import PdfFileMerger
from loguru import logger
from tqdm import tqdm


def word2Pdf1(filePath, words):
    if len(words) < 1:
        logger.warning("\n【无 Word 文件】\n")
        return
    # 开始转换
    logger.info("\n【开始 Excel -> PDF 转换】")
    try:
        pdfs = []
        logger.info("打开 Word 进程中...")

        for i in range(len(words)):
            logger.debug(i)
            fileName = words[i]  # 文件名称
            fromFile = os.path.join(filePath, fileName)  # 文件地址
            logger.info("转换：" + fileName + "文件中...")

            # 某文件出错不影响其他文件打印
            try:
                cmd = ["D:\Software\Libre Offices\program\soffice.com", "--headless", "--convert-to", "pdf", fromFile,
                       "--outdir", Path(fileName).parent]

                proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                pdfs.append(fromFile.replace("docx", "pdf"))
                out, err = proc.communicate()

                if out is not None:
                    try:
                        out = out.decode('GBK')
                    except UnicodeDecodeError:
                        out = out.decode('utf-8')

                if err is not None:
                    try:
                        err = err.decode('utf-8')
                    except UnicodeDecodeError:
                        err = err.decode('GBK')

                print('Output:', out)
                print('Error:', err)
            except Exception as e:
                logger.exception(e)
        return sort_docx_files(pdfs)
    except Exception as e:
        logger.warning(e)
    finally:
        gc.collect()

def word2Pdf2(filePath, words):
    if len(words) < 1:
        logger.warning("\n【无 Word 文件】\n")
        return
    # 开始转换
    logger.info("\n【开始 Excel -> PDF 转换】")
    try:
        pdfs = []
        logger.info("打开 Word 进程...")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = 0
        word.DisplayAlerts = False
        doc = None
        for i in range(len(words)):
            logger.debug(i)
            fileName = words[i]  # 文件名称
            fromFile = os.path.join(filePath, fileName)  # 文件地址
            toFileName = changeSufix2Pdf(fileName)  # 生成的文件名称
            toFile = toFileJoin(filePath, toFileName)  # 生成的文件地址

            logger.info("转换：" + fileName + "文件中...")
            # 某文件出错不影响其他文件打印
            try:
                doc = word.Documents.Open(fromFile)
                doc.SaveAs(toFile, 17)  # 生成的所有 PDF 都会在 PDF 文件夹中
                pdfs.append(fromFile.replace("docx", "pdf"))
                logger.success("转换到：" + toFileName + "完成")
            except Exception as e:
                logger.exception(e)
            # 关闭 Word 进程
        logger.success("所有 Word 文件已打印完毕")
        logger.success("结束 Word 进程...\n")
        doc.Close()
        word.Quit()
        return sort_docx_files(pdfs)
    except Exception as e:
        logger.warning(e)
    finally:
        gc.collect()


# 常量定义
PDF_FORMAT = 17


def openWordApplication():
    """打开Word应用，隐藏并禁用警告"""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = 0
    word.DisplayAlerts = False
    return word


def closeWordApplication(word):
    """关闭Word应用，确保资源释放"""
    if word is not None:
        word.Quit()
        del word


def convertDocument(word, fromFile, toFile):
    """转换单个文档"""
    try:
        doc = word.Documents.Open(fromFile)
        doc.SaveAs(toFile, PDF_FORMAT)
        doc.Close()
        logger.success(f"转换到：{os.path.basename(toFile)}完成")
        return fromFile.replace("docx", "pdf")
    except Exception as e:
        logger.error(f"转换 {os.path.basename(fromFile)} 失败: {e}")
        return None


def word2Pdf(filePath, words):
    if not words:
        logger.warning("\n【无 Word 文件】\n")
        return []

    logger.info("\n【开始 Excel -> PDF 转换】")
    pdfs = []
    word = openWordApplication()
    try:
        for i, fileName in enumerate(words):
            fromFile = os.path.join(filePath, fileName)
            if not os.path.isfile(fromFile):
                logger.warning(f"{fileName} 文件不存在，跳过转换")
                continue

            toFileName = changeSufix2Pdf(fileName)
            toFile = toFileJoin(filePath, toFileName)

            logger.info(f"转换：{fileName}文件中...")
            pdfPath = convertDocument(word, fromFile, toFile)
            if pdfPath:
                pdfs.append(pdfPath)
            else:
                logger.warning(f"{fileName} 文件转换失败")

        logger.success("所有 Word 文件已转换完毕")
    except Exception as e:
        logger.error(f"转换过程中发生异常: {e}")
    finally:
        closeWordApplication(word)
        gc.collect()

    return sort_docx_files(pdfs)


@logger.catch
def changeSufix2Pdf(file):
    return file[:file.rfind('.')] + ".pdf"


@logger.catch
def toFileJoin(filePath, file):
    return os.path.join(filePath, 'pdf', file[:file.rfind('.')] + ".pdf")


def sort_docx_files(docx_files):
    # 定义排序规则
    sort_order = {
        "分部工程质量检验评定记录-报验表": 9,
        "分部工程质量检验评定记录":8,
        "分项质量检查验收记录-报验表": 7,
        "分项质量检查验收记录": 6,
        "检验批质量检验评定记录": 5
    }

    # 按照指定顺序排序
    sorted_files = sorted(docx_files, key=lambda x: sort_order.get(os.path.basename(x).split('-')[0], 0), reverse=True)


    return sorted_files



def is_blank_page(page):
    """检查页面是否为空白"""
    img = page.get_pixmap()
    if img.n >= 3:  # 颜色图像（RGB）
        return all(sum(row) / len(row) > 250 for row in zip(*[iter(img.samples)] * 3))
    else:  # 灰度图像
        return all(pixel > 250 for pixel in img.samples)


def remove_blank_pages(pdf_file, output_pdf):
    """从PDF中移除空白页"""
    doc = fitz.open(pdf_file)
    blank_pages = [i for i in range(doc.page_count) if is_blank_page(doc[i])]
    for i in reversed(blank_pages):  # 从后向前删除，避免改变页码
        doc.delete_page(i)
    doc.save(output_pdf)
    doc.close()


def merge_pdfs(pdf_files, output_pdf):
    # 创建PdfFileMerger对象
    merger = PdfFileMerger()

    # 先移除每个PDF文件中的空白页
    cleaned_pdf_files = []
    for pdf_file in pdf_files:
        cleaned_file = f"{os.path.splitext(pdf_file)[0]}_cleaned.pdf"
        remove_blank_pages(pdf_file, cleaned_file)
        cleaned_pdf_files.append(cleaned_file)

    # 合并已移除空白页的PDF文件
    logger.info("开始合并文件...")
    for cleaned_pdf_file in cleaned_pdf_files:
        merger.append(cleaned_pdf_file)

    # 保存合并后的PDF文件
    merger.write(output_pdf)
    merger.close()

    # 删除转换过程中的临时PDF文件
    for cleaned_pdf_file in cleaned_pdf_files:
        os.remove(cleaned_pdf_file)

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
    directory_path = 'D:\Jobs\卡莫亚\检验批及分项\非标设备检验批生成'
    # 获取文件夹名称的集合
    folders = list_folders(directory_path)
    for folder in tqdm(folders):
        print(folder)
        # 指定目录及子目录下的所有.docx文件
        root_folder = os.path.join(directory_path, folder)
        # 指定合并后保存的文件名
        output_folder = os.path.join(directory_path, folder, folder + ".pdf")
        # 执行主函数
        pdfs = create_and_list_pdf_files(root_folder)
        merge_pdfs(pdfs, output_folder)
