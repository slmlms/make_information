import PyPDF2

def extract_pages(input_pdf_path, output_pdf_path, skip_pages,start_page):
    # 打开原始PDF文件（只读模式）
    with open(input_pdf_path, 'rb') as in_file:
        original_pdf = PyPDF2.PdfFileReader(in_file)

        # 创建一个新的PDF写入器，用于生成不修改原始文件的新文档
        new_pdf = PyPDF2.PdfFileWriter()

        # 从原始PDF中按指定间隔提取页面
        for i in range(start_page, original_pdf.getNumPages(), (skip_pages )):
            # 添加页面到新的PDF文档（不会影响原始PDF）
            new_pdf.addPage(original_pdf.getPage(i))

        # 将新文档保存到输出文件（过程中不会修改原始PDF）
        with open(output_pdf_path, 'wb') as out_file:
            new_pdf.write(out_file)

# 使用函数，每隔2页提取一页并保存到output.pdf，原始的input.pdf保持不变
input_pdf = "D:\Jobs\洛钼\设备质量证明文件 -邓玉拷\大全\东区选矿C1205702 KYN28-12 资料\电流互感器.pdf"
output_pdf = "D:\Jobs\洛钼\设备质量证明文件 -邓玉拷\大全\东区选矿C1205702 KYN28-12 资料\电流互感器封面.pdf"
extract_pages(input_pdf, output_pdf, 2,0)
