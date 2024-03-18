import win32com.client as win32

def extract_pages(input_file, output_file, start_page, end_page):
    """
    提取给定范围的页面。
    参数：
    input_file: 输入的Word文件名。
    output_file: 输出的Word文件名。
    start_page: 开始页面的索引（从1开始）。
    end_page: 结束页面的索引（包含）。
    """
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(input_file)
    doc.Activate()

    # 创建一个新的文档
    new_doc = word.Documents.Add()

    # 提取页面
    for i in range(start_page, end_page + 1):
        # 选择页面
        word.Selection.GoTo(What=win32.constants.wdGoToPage, Which=win32.constants.wdGoToAbsolute, Count=i)
        word.Selection.Copy()
        # 粘贴到新文档
        new_doc.Activate()
        word.Selection.PasteAndFormat(Type=win32.constants.wdFormatOriginalFormatting)

    # 保存新文档
    new_doc.SaveAs(output_file)
    new_doc.Close()

    # 关闭原始文档
    doc.Close()


# 例子：提取第2-4页并保存
extract_pages(
    "D:\Jobs\图集规范\建筑工程资料(下载到电脑上解压)\地方版资料（陆续更新）\湖北省建筑工程施工统一用表\湖北省建筑工程施工统一用表\下册.docx",
    "D:\Jobs\图集规范\建筑工程资料(下载到电脑上解压)\地方版资料（陆续更新）\湖北省建筑工程施工统一用表\湖北省建筑工程施工统一用表\智能建筑.docx",
    152, 221)
