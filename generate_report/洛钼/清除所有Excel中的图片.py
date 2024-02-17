import fitz

def remove_editable_text(pdf_path, output_path):
    # 打开PDF文件
    with fitz.open(pdf_path) as pdf_document:
        for page_number, page in enumerate(pdf_document):
            text_blocks = page.get_text("blocks")

            for text_block in text_blocks:
                # 检查当前文本块是否是文本（非图像或其他类型）
                if isinstance(text_block[0], str):  # 或者使用更适合的条件来判断文本块是否为文本
                    flags = int(text_block[4])  # 将文本块的标志位转换为整数
                    if flags & 0x08:  # 检查是否为可编辑文本（文本块的第四位为0x08）
                        x0, y0, x1, y1 = text_block[:4]
                        page.draw_rect(fitz.Rect(x0, y0, x1, y1), fill=(1, 1, 1))

    # 保存修改后的PDF
    pdf_document.save(output_path)


# 调用函数并指定输入和输出路径
input_pdf_path = "D:\\Videos\\机电\\01-文档资料（含真题）\\09-HQ-机电-名师讲义.pdf"
output_pdf_path = "D:\\Videos\\机电\\01-文档资料（含真题）\\09-HQ-机电-名师讲义1.pdf"
remove_editable_text(input_pdf_path, output_pdf_path)
