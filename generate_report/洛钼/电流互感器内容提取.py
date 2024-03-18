import os
import re
from openpyxl import Workbook

# 创建一个 Excel 工作簿
wb = Workbook()
ws = wb.active

# 设置表头
ws.append(["产品编号", "变比"])
# 定义包含所有可能变比值的列表
allowed_ratios = [str(x) for x in range(50, 2001, 50)]

# 构造只允许提取指定数值的正则表达式
pattern_ratio = r"(?<![\d.,])(" + "|".join(allowed_ratios) + r")(?:,\s*(" + "|".join(allowed_ratios) + r"))*"

# 指定目录路径
directory = r"D:\Jobs\洛钼\设备质量证明文件 -邓玉拷\大全\东区选矿C1205702 KYN28-12 资料\零序互感器"

# 遍历指定目录下的所有文件
for file_name in os.listdir(directory):
    # 拼接文件路径并只处理文本文件
    if os.path.isfile(os.path.join(directory, file_name)) and file_name.endswith('.txt'):
        with open(os.path.join(directory, file_name), "r", encoding="utf-8") as file:
            content = file.read().replace("\n", "")

            # 提取产品编号信息
            pattern_number = r"\s*(\d{8})"
            number_match = re.search(pattern_number, content)
            number = number_match.group(1) if number_match else ""

            # 提取变比信息（只匹配指定范围内的整数序列）
            ratio_match = re.search(pattern_ratio, content)
            ratio = ratio_match.group() if ratio_match else ""

            # 将提取的信息添加到 Excel 表格中（假设提取的是以逗号分隔的数字，因此需要移除逗号以便存储）
            ratio_values = ratio.replace(",", "").split() if ratio else []
            ws.append([number] + ratio_values)



# 保存 Excel 文件
wb.save(r"D:\Jobs\洛钼\设备质量证明文件 -邓玉拷\大全\东区选矿C1205702 KYN28-12 资料\零序互感器\output.xlsx")
