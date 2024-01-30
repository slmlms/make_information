import multiprocessing
import os
import pathlib
import sys
import docxtpl
import xlrd
import utils.data_util as data

# 定义生成Word文档的函数，传入进程安全的队列来传递结果
def run(row, i, result_queue, args):
    try:
        # 获取模板路径、创建word模板对象
        template_path = pathlib.Path('E:\ideaProject\make_information\\resources\inspection_lot\实验报告\低压电缆实验记录模板.docx')
        with open(template_path, 'rb') as f:
            word_template = docxtpl.DocxTemplate(f)

        num = row[0]
        start = row[1]
        end = row[2]
        u0 = row[4]
        cable = row[3]
        len_value = row[5]
        place = row[7]
        date = data.int_to_date(row[6])

        context = {'id': num, 'num': num, 'start': start, 'end': end, 'u0': u0,
                   'cable': cable, 'len': len_value, 'date': date, 'place': place}

        save_file = save_path.joinpath(str(i) + '.docx')
        # 使用with语句确保文件正确关闭
        with open(save_file, 'wb') as output:
            doc = word_template.render(context)
            output.write(doc)

        # 将保存路径放入队列中
        result_queue.put(save_file)

    except Exception as e:
        # 在这里处理任何异常，比如记录日志或错误处理
        print(f"Error processing row {i}: {e}")

if __name__ == '__main__':
    save_path = pathlib.Path('E:\工作\庞比\报验资料\调试记录\石灰乳\低压电缆\\')
    work_book_path = "E:\工作\庞比\施工图\\906刚果（金）庞比铜钴矿项目蓝图PDF版\电气仪表\电力\石灰乳及石灰石浆制备-电力\石灰乳电缆表.xls"

    # 根据可用CPU核心数设置进程池大小
    num_processes = max(1, int(multiprocessing.cpu_count() * 2 / 3))
    pool = multiprocessing.Pool(processes=num_processes)
    sheet = xlrd.open_workbook(work_book_path).sheet_by_name('实验报告')

    # 创建一个进程安全的队列来接收每个进程的结果
    result_queue = multiprocessing.Queue()

    for i in range(1, sheet.nrows):
        row = sheet.row_values(i)
        # 将任务提交给线程池
        pool.apply_async(run, args=(row, i, result_queue))

    # 关闭线程池（不再接受新的任务）
    pool.close()

    # 等待所有任务完成，并收集结果
    processed_files = []
    while not result_queue.empty():
        processed_files.append(result_queue.get())

    # 确保所有进程都已经结束
    pool.join()

    # 处理已生成的Word文件列表（例如：将它们转换为PDF格式）
    data.make_pdf(set(processed_files), output_dir=save_path.parent)

    # ... 其他后续操作 ...
