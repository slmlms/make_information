"""
【程序功能】：将【目标文件夹】内所有的 ppt、excel、word 均生成一份对应的 PDF 文件
【作者】：evgo，公众号（随风前行），Github（evgo2017）
【目标文件夹】：默认为此程序目前所在的文件夹；
                若输入路径，则为该文件夹（只转换该层，不转换子文件夹下内容）
【生成的pdf名称】：原始名称+.pdf
"""
import gc
import os
import pathlib
import re
import shutil
from datetime import datetime
import utils.data_util as data_util
import win32com.client
from PyPDF2 import PdfFileReader, PdfFileWriter
from loguru import logger


# Word


@logger.catch
def word2Pdf(filePath, words):
    # 如果没有文件则提示后直接退出
    if len(words) < 1:
        logger.warning("\n【无 Word 文件】\n")
        return
    # 开始转换
    logger.info("\n【开始 Word -> PDF 转换】")
    try:
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
                logger.success("转换到：" + toFileName + "完成")
            except Exception as e:
                logger.exception(e)
            # 关闭 Word 进程
        logger.success("所有 Word 文件已打印完毕")
        logger.success("结束 Word 进程...\n")
        doc.Close()
        word.Quit()
    except Exception as e:
        logger.warning(e)
    finally:
        gc.collect()


# Excel


@logger.catch
def excel2Pdf(filePath, excels):
    # 如果没有文件则提示后直接退出
    if len(excels) < 1:
        logger.warning("\n【无 Excel 文件】\n")
        return
    # 开始转换
    logger.info("\n【开始 Excel -> PDF 转换】")
    try:
        logger.info("打开 Excel 进程中...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = 0
        excel.DisplayAlerts = False

        for i in range(len(excels)):
            logger.debug(i)
            fileName = excels[i]  # 文件名称
            fromFile = os.path.join(filePath, fileName)  # 文件地址

            logger.info("转换：" + fileName + "文件中...")
            # 某文件出错不影响其他文件打印
            try:
                wb = excel.Workbooks.Open(fromFile)
                for j in range(1):  # 工作表数量，一个工作簿可能有多张工作表
                    toFileName = addWorksheetsOrder(fileName)  # 生成的文件名称
                    toFile = toFileJoin(filePath, toFileName)  # 生成的文件地址

                    ws = wb.Worksheets(j + 1)  # 若为[0]则打包后会提示越界
                    ws.ExportAsFixedFormat(0, toFile)  # 每一张都需要打印
                    logger.success("转换至：" + toFileName + "文件完成")
            except Exception as e:
                logger.exception(e)
        # 关闭 Excel 进程
        logger.success("所有 Excel 文件已打印完毕")
        logger.success("结束 Excel 进程中...\n")
        close_excel_by_force(excel)
    except Exception as e:
        logger.exception(e)
    finally:
        gc.collect()


# PPT


@logger.catch
def ppt2Pdf(filePath, ppts):
    # 如果没有文件则提示后直接退出
    if len(ppts) < 1:
        print("\n【无 PPT 文件】\n")
        return
    # 开始转换
    logger.info("\n【开始 PPT -> PDF 转换】")
    try:
        print("打开 PowerPoint 进程中...")
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        ppt = None
        # 某文件出错不影响其他文件打印

        for i in range(len(ppts)):
            print(i)
            fileName = ppts[i]  # 文件名称
            fromFile = os.path.join(filePath, fileName)  # 文件地址
            toFileName = changeSufix2Pdf(fileName)  # 生成的文件名称
            toFile = toFileJoin(filePath, toFileName)  # 生成的文件地址

            print("转换：" + fileName + "文件中...")
            try:
                ppt = powerpoint.Presentations.Open(fromFile, WithWindow=False)
                if ppt.Slides.Count > 0:
                    ppt.SaveAs(toFile, 32)  # 如果为空则会跳出提示框（暂时没有找到消除办法）
                    print("转换至：" + toFileName + "文件完成")
                else:
                    print("（错误，发生意外：此文件为空，跳过此文件）")
            except Exception as e:
                print(e)
        # 关闭 PPT 进程
        print("所有 PPT 文件已打印完毕")
        print("结束 PowerPoint 进程中...\n")
        ppt.Close()
        powerpoint.Quit()
    except Exception as e:
        print(e)
    finally:
        gc.collect()


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


# 修改后缀名


@logger.catch
def changeSufix2Pdf(file):
    return file[:file.rfind('.')] + ".pdf"


# 添加工作簿序号


@logger.catch
def addWorksheetsOrder(file):
    return file[:file.rfind('.')] + ".pdf"


# 转换地址

# 移动文件
@logger.catch
def mymovefile(srcfile, dstfile):
    if not os.path.isfile(srcfile):
        logger.warning("{src} not exist!", src=srcfile)
    else:
        fpath, fname = os.path.split(dstfile)  # 分离文件名和路径
        if not os.path.exists(fpath):
            os.makedirs(fpath)  # 创建路径
        if pathlib.Path(dstfile).exists():
            pathlib.Path(dstfile).unlink()
        shutil.move(srcfile, dstfile)  # 移动文件
        logger.info("move {src} -> {dst}", src=srcfile, dst=dstfile)


@logger.catch
def toFileJoin(filePath, file):
    return os.path.join(filePath, 'pdf', file[:file.rfind('.')] + ".pdf")


@logger.catch
def getFileName(filedir):
    file_list = [os.path.join(root, filespath)
                 for root, dirs, files in os.walk(filedir)
                 for filespath in files
                 if str(filespath).endswith('pdf')
                 ]
    return file_list if file_list else []


@logger.catch
def mergePDF(filePath):
    outfile = 'Merge(' + str(datetime.now().strftime('%Y-%m-%d %H-%M-%S')) + ').pdf'
    output = PdfFileWriter()
    outputPages = 0
    pdf_fileName = getFileName(filePath)
    regrx = 'Merge\(\d{4}-\d{2}-\d{2} \d{2}-\d{2}-\d{2}\).pdf'
    if pdf_fileName:
        op_ls = []
        for pdf_file in pdf_fileName:
            logger.info("路径：%s" % pdf_file)
            if not re.search(regrx, pdf_file) == None:
                continue
            # 读取源PDF文件
            op = open(pdf_file, "rb")
            input = PdfFileReader(op)

            # 获得源PDF文件中页面总数
            pageCount = input.getNumPages()
            outputPages += pageCount
            logger.info("页数：%d" % pageCount)

            # 分别将page添加到输出output中
            for iPage in range(pageCount):
                output.addPage(input.getPage(iPage))
            op_ls.append(op)
        logger.success("合并后的总页数:%d." % outputPages)
        # 写入到目标PDF文件
        outputStream = open(os.path.join(filePath, outfile), "wb")
        output.write(outputStream)
        outputStream.close()
        for op in op_ls:
            op.close()
        for pdf_tempfile in pdf_fileName:

            result = re.search(regrx, pdf_tempfile)
            if result == None:
                os.remove(pdf_tempfile)
        logger.success("PDF文件合并完成！")

    else:
        logger.warning("没有可以合并的PDF文件！")


@logger.catch
def run(filePath):
    # 开始程序

    # 目标路径，若没有输入路径则为当前路径
    if filePath == "":
        filePath = os.getcwd()

    # 将目标文件夹所有文件归类，转换时只打开一个进程
    words = []
    excels = []

    for fn in os.listdir(filePath):
        if fn.endswith(('.doc', 'docx')):
            words.append(fn)
        # if fn.endswith(('.ppt', 'pptx')):
        #     ppts.append(fn)
        if fn.endswith(('.xls', 'xlsx')):
            excels.append(fn)

    # 调用方法
    logger.info("====================开始转换====================")

    # 新建 pdf 文件夹，所有生成的 PDF 文件都放在里面
    folder = str(os.path.join(filePath, 'pdf'))
    if not os.path.exists(folder):
        os.makedirs(folder)
    folder1 = str(os.path.join(filePath, 'backup'))
    if not os.path.exists(folder1):
        os.makedirs(folder1)

    word2Pdf(filePath, words)
    excel2Pdf(filePath, excels)
    # ppt2Pdf(filePath, ppts)
    for fn in os.listdir(filePath):
        if fn.endswith(('.doc', 'docx')) | fn.endswith(('.xls', 'xlsx')):
            mymovefile(os.path.join(filePath, fn), os.path.join(folder1, fn))

    logger.info("====================转换结束====================")
    logger.info("\n====================程序结束====================")
    logger.info("\n====================开始合并PDF====================")
    mergePDF(folder)
    logger.info("\n====================PDF合并完成====================")
