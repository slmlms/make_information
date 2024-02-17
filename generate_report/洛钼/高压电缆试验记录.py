import os
import shutil
import openpyxl

# 指定目录列表
directory_list = ["D:\Jobs\洛钼\调试记录\电力电缆试验记录(10-35kV）\\10kV单芯电缆.xlsx 02月07日", "D:\Jobs\洛钼\调试记录\电力电缆试验记录(10-35kV）\\10kV多芯电缆.xlsx 02月07日"]

# 遍历每个目录
for directory in directory_list:
    # 获取目录中的所有.xlsx文件
    for filename in os.listdir(directory):
        if filename.endswith('.xlsx'):
            filepath = os.path.join(directory, filename)

            # 使用openpyxl打开.xlsx文件
            wb = openpyxl.load_workbook(filepath)
            sheet = wb.active

            # 获取F4和F7单元格的值
            f4_value = sheet['F4'].value
            f7_value = sheet['F7'].value.replace('/', '-')

            # 构建新的文件名
            new_filename = f"{f7_value}-{f4_value}.xlsx"

            # 移动文件到指定文件夹
            new_directory = os.path.join('D:\Jobs\洛钼\调试记录\电力电缆试验记录(10-35kV）\\', f7_value)
            os.makedirs(new_directory, exist_ok=True)
            new_filepath = os.path.join(new_directory, new_filename)
            shutil.move(filepath, new_filepath)
