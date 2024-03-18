import os
from win32com import client
import tqdm
import os

# 获取所有.doc文件的路径

def get_doc_files(path):
    doc_files = []

    for root, dirs, files in os.walk(path):
        for file in files:
            if file.endswith(".doc") or file.endswith(".DOC"):
                doc_file_path = os.path.join(root, file)
                doc_files.append(doc_file_path)

    return doc_files


# 将.doc文件转换为.docx文件
def convert_doc_to_docx(doc_files):
    word = client.Dispatch("Word.Application")
    for file in tqdm.tqdm(doc_files):
        doc = word.Documents.Open(file)
        doc.SaveAs("{}x".format(file), 12)  # 12 represents the file type - .docx
        doc.Close()
    word.Quit()


# 使用函数
path = "D:\Jobs\图集规范\建筑工程资料(下载到电脑上解压)\地方版资料（陆续更新）\湖南建筑资料员表格\湖南省建筑工程全套资料表格\湖南省建筑工程全套资料表格\检验批验收表\\08(39)"  # 请替换为你的.doc文件所在的文件夹路径
doc_files = get_doc_files(path)
convert_doc_to_docx(doc_files)
