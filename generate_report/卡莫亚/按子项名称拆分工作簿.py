import os
import shutil


def move_files(source_dir, target_dir):
    # 创建目标目录
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)

    # 获取所有文件
    files = os.listdir(source_dir)

    # 遍历文件
    for file in files:
        # 判断文件是否为文件夹
        if not os.path.isdir(os.path.join(source_dir, file)):
            # 获取文件名
            name = file.split(".")[0]
            # 获取文件名中的“-”之前的内容
            content = name.split("-")[0]
            # 创建目标文件夹
            target_folder = os.path.join(target_dir, content)
            if not os.path.exists(target_folder):
                os.makedirs(target_folder)
            # 移动文件到目标文件夹
            shutil.move(os.path.join(source_dir, file), target_folder)


# 设置源目录和目标目录
source_dir = "C:\\Users\shuli\Desktop\电缆敷设检验批.xlsx 12月18日"
target_dir = "D:\Jobs\卡莫亚\检验批及分项\检验批\电缆敷设\检验批"

# 调用函数
move_files(source_dir, target_dir)