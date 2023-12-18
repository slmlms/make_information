import os

import docxtpl
import openpyxl
import pandas as pd

# 指定文档和数据源路径
word_template_path = "D:\Jobs\卡莫亚\检验批及分项\模板\\5电缆敷设\电缆敷设报审、报验申请表.docx"
data_source_path = "D:\Jobs\卡莫亚\检验批及分项\模板\\5电缆敷设\电缆敷设数据源.xlsx"
excel_template = "D:\Jobs\卡莫亚\检验批及分项\模板\\5电缆敷设\分项工程-电缆敷设.xlsx"
sava_path = "D:\Jobs\卡莫亚\检验批及分项\检验批\电缆敷设\\"
# 打开数据源并读取模板数据
workbook = openpyxl.load_workbook(data_source_path)
sheet = workbook["数据源"]
df = pd.read_excel(data_source_path, sheet_name="数据源")
template_data_list = []

# 按照B列的值进行分组
groups = df.groupby('单位工程名称')

# 遍历每个分组，将H列的值加入字典中
result = {}
for name, group in groups:
    template_data = {"Danweigognchengmingcheng": str(name), "yanshoubuwei": group['验收部位'].tolist()}
    template_data_list.append(template_data)
print(template_data_list)

# 加载模板文档
template_path = os.path.abspath(word_template_path)
excel_path = os.path.abspath(excel_template)

# 填充模板文档
for template_data in template_data_list:
    for key, value in template_data.items():
        doc = docxtpl.DocxTemplate(template_path)
        xls = openpyxl.load_workbook(excel_path)
        doc.render(template_data)
        # 保存填充后的文档

        output_path = sava_path + "报审表\\" + template_data["Danweigognchengmingcheng"] + "报审表.docx"
        doc.save(output_path)

        print(f"填充后的文档已保存至 {output_path}")

        xls.worksheets[0]["S5"].value = template_data["Danweigognchengmingcheng"]
        for i in range(len(template_data["yanshoubuwei"])):
            xls.worksheets[0]["D" + str(i + 9)].value = template_data["yanshoubuwei"][i]

        xls.save(sava_path + "分项\\" + template_data["Danweigognchengmingcheng"] + "分项验收记录.xlsx")
        xls.close()
