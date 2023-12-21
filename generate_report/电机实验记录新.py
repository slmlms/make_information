import os
import random
import re

import openpyxl
import pandas as pd
from openpyxl.styles import Font, Alignment

workbook_path = "D:\Jobs\卡隆威\检验批数据集卡隆威.xlsx"
sheet_name = '电机实验记录'
template_path = 'D:\\Documents\PycharmProjects\\make_information\\resources\\inspection_lot\\实验报告\\3.2.3  100kW及以下低压电动机.xlsx'
save_path = 'D:\Jobs\卡隆威\验收资料\调试记录\电机试验记录\\3万吨\\'


def get_no_load_current(rated_current):
    # 使用正则表达式提取额定电流中的数字
    rated_current_digits = re.findall(r'\d+\.?\d*', rated_current)
    # 将提取到的数字转换为浮点数
    rated_current_digits = float(rated_current_digits[0])
    # 计算无负载电流
    no_load_current = round(rated_current_digits * random.uniform(0.3, 0.4), 2)
    return no_load_current


def get_load_current(rated_current):
    # 使用正则表达式提取额定电流中的数字
    rated_current_digits = re.findall(r'\d+\.?\d*', rated_current)
    # 将提取到的数字转换为浮点数
    rated_current_digits = float(rated_current_digits[0])
    # 计算负载电流
    load_current = round(rated_current_digits * random.uniform(0.65, 0.75), 2)
    load_current = round(load_current, 2)
    return load_current


def fill_template(group_name, group_data, template_path):
    # Load the template workbook
    # 加载模板工作簿
    template_wb = openpyxl.load_workbook(template_path)
    # Select the first sheet
    # 选择第一个工作表
    template_ws = template_wb.active
    max_date = pd.to_datetime(group_data['检查日期']).max().strftime('%Y年%m月%d日')
    template_ws['O1'] = max_date
    # Loop through each row in group_data
    # 循环遍历group_data中的每一行
    index = 0
    for _, row_data in group_data.iterrows():
        # Loop through each column in the row
        # 循环遍历行中的每一列
        template_ws['B' + str(index + 3)] = row_data['子项名称']
        template_ws['D' + str(index + 3)] = row_data['验收部位']
        template_ws['F' + str(index + 3)] = row_data['制造厂']
        template_ws['H' + str(index + 3)] = row_data['出厂编号']
        template_ws['J' + str(index + 3)] = row_data['出厂日期']
        template_ws['L' + str(index + 3)] = str(row_data['额定功率']) + 'kW'
        template_ws['M' + str(index + 3)] = row_data['功率因数']
        template_ws['N' + str(index + 3)] = row_data['转速']
        template_ws['O' + str(index + 3)] = row_data['额定电流']
        template_ws['P' + str(index + 3)] = '>100MΩ'
        if '液下泵' in row_data['验收部位'] or '搅拌' in row_data['验收部位']:
            template_ws['Q' + str(index + 3)] = '/'
        else:
            template_ws['Q' + str(index + 3)] = str(get_no_load_current(row_data['额定电流'])) + "A"
        template_ws['R' + str(index + 3)] = str(get_load_current(row_data['额定电流'])) + "A"
        index += 1

    # Save the filled template
    # 保存填充后的模板
    for row in template_ws.iter_rows(min_row=3, max_row=33, min_col=2, max_col=18):
        for cell in row:
            cell.font = Font(size=8)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    save_file_name = save_path + group_name
    # 判断save_file_name所在的目录是否存在，如果不存在就创建该目录
    if not os.path.exists(os.path.dirname(save_file_name)):
        os.makedirs(os.path.dirname(save_file_name))
    template_wb.save(save_file_name + ".xlsx")


# 使用pandas打开workbook_path中的sheet_name，将其转换为dataframe，变量名为data
data = pd.read_excel(workbook_path, sheet_name=sheet_name, header=0)

# 将data中“是否报送”为True和“子项名称”为NaN的行删除
data = data[(data['是否报送'] == True) | (data['子项名称'].isnull() == False)]
data['额定功率'] = data['额定功率'].str.extract('(\d+\.?\d*)', expand=False).astype(float)
data = data[data['额定功率'] < 100]
# 将data按照“子项名称”分组
grouped_data = data.groupby('子项名称')
# 100kW以下报告生成
for group_name, group_data in grouped_data:
    if len(group_data) > 30:
        sub_groups = [group_data.iloc[i:i + 30] for i in range(0, len(group_data), 30)]
        for i, sub_group in enumerate(sub_groups):
            sub_group_name = f"{group_name}_{i + 1}"
            fill_template(sub_group_name, sub_group, template_path)

    else:
        fill_template(group_name, group_data, template_path)
