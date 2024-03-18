import os
import subprocess
import sys

import pandas as pd
from docxtpl import DocxTemplate

# 读取Excel文件
xls = pd.ExcelFile("D:\Jobs\洛钼\材料构配件设备数量清单及自检结果\仪表及消防\数据源.xlsx")

# 获取所有sheet的名称
sheets = xls.sheet_names

# 遍历每个sheet
for sheet in sheets:
    df = pd.read_excel(xls, sheet_name=sheet)

    # 按照子项名称分类
    grouped = df.groupby('子项名称')

    # 遍历每个子项名称
    for name, group in grouped:
        # 加载Word模板
        doc = DocxTemplate("D:\Jobs\洛钼\材料构配件设备数量清单及自检结果\仪表及消防\模板.docx")
        # 将数据转换为字典列表
        data = group.to_dict(orient='records')

        # 如果data的长度小于10，添加空的字典到data中，直到data的长度等于10
        while len(data) < 12:
            data.append({})
        # 将数据填充到模板中
        context = {'data': data}
        context['zi_xiang_ming_cheng'] = name
        print(context)
        doc.render(context)

        # 保存文件
        doc.save(f'D:\Jobs\洛钼\材料构配件设备数量清单及自检结果\仪表及消防\仪表部分\\{name}-{sheet}.docx')

filePath = 'D:\Jobs\洛钼\材料构配件设备数量清单及自检结果\仪表及消防\仪表部分'
encoding = sys.getdefaultencoding()
for file in os.listdir(filePath):
    if file.endswith(".docx"):
        fromFile = filePath + "\\" + file
        toFile = filePath + "\\" + file[:-5] + ".pdf"
        cmd = ["D:\Software\Libre Offices\program\soffice.com", "--headless", "--convert-to", "pdf", fromFile,
               "--outdir", filePath + "\\"]

        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        out, err = proc.communicate()

        if out is not None:
            try:
                out = out.decode('utf-8')
            except UnicodeDecodeError:
                out = out.decode('GBK')

        if err is not None:
            try:
                err = err.decode('utf-8')
            except UnicodeDecodeError:
                err = err.decode('GBK')

        print('Output:', out)
        print('Error:', err)
