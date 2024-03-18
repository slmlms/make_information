import openpyxl
from docxtpl import DocxTemplate

# 打开Excel表格
workbook = openpyxl.load_workbook("D:\Jobs\洛钼\调试记录\气动阀调试记录\数据源.xlsx",data_only=True)

# 加载Word模板
template = DocxTemplate("D:\Jobs\洛钼\调试记录\气动阀调试记录\气动阀调试记录.docx")

# 遍历所有Sheet
for sheet in workbook.worksheets:
    data = {}
    data['zixiangmingcheng'] = sheet.title

    i = 1
    for row in sheet.iter_rows():
        if row[0].value == "序号": continue
        value0 = str(row[0].value) if row[0].value is not None else ""
        data["xuhao" + str(i)] = value0
        value1 = str(row[1].value) if row[1].value is not None else ""
        data["famenbianhao" + str(i)] = value1
        value2 = str(row[2].value) if row[2].value is not None else ""
        data["famenmingcheng" + str(i)] = value2
        value3 = str(row[3].value) if row[3].value is not None else ""
        data["edingyali" + str(i)] = value3
        value4 = str(row[4].value) if row[4].value is not None else ""
        data['zhengdingyali' + str(i)] = value4
        value5 = row[5].value if row[5].value is not None else ""
        data['kaishijian' + str(i)] = value5
        value6 = row[6].value if row[6].value is not None else ""
        data['guanshijian' + str(i)] = value6
        value7 = str(row[7].value) if row[7].value is not None else ""
        data['qiyuanyali' + str(i)] = value7

        i += 1

    print(data)
    template.render(data)
    template.save("D:\Jobs\洛钼\调试记录\气动阀调试记录\\" + sheet.title + ".docx")

# 保存生成的新文档
# template.save("output.docx")
