import subprocess
import fitz  # PyMuPDF
import easyocr
import os
from tqdm import tqdm  # 导入tqdm库
import openai

# 设置你的OpenAI API密钥
openai.api_key = "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"

# 定义路径
pdf_path = "D:\Jobs\图集规范\GB 50150-2016 电气装置安装工程电气设备交接试验标准.pdf"
output_dir = 'D:/Jobs/pdf extract/outputs'
upscale_dir = 'D:/Jobs/pdf extract/upscaled_images'
result_dir = 'D:/Jobs/pdf extract/results'
result_file = os.path.join("D:/Jobs/pdf extract/results", "extracted_text.txt")

# 将PDF转换为图像
def convert_pdf_to_images(pdf_path, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    pdf_document = fitz.open(pdf_path)
    for page_num in tqdm(range(pdf_document.page_count), desc="Converting PDF to Images"):
        page = pdf_document[page_num]
        image_path = os.path.join(output_dir, f'page_{page_num + 1}.png')

        # 检查图像文件是否已经存在
        if not os.path.exists(image_path):
            image = page.get_pixmap()
            image.save(image_path)
        else:
            print(f"Image for page {page_num + 1} already exists, skipping conversion.")

    pdf_document.close()

# 执行OCR并保存结果
def perform_ocr_and_save_results(image_dir, result_dir):
    if not os.path.exists(result_dir):
        os.makedirs(result_dir)

    reader = easyocr.Reader(['ch_sim', 'en'])

    for image_file in tqdm(os.listdir(image_dir), desc="Performing OCR"):
        if image_file.endswith('.png'):
            image_path = os.path.join(image_dir, image_file)
            result_file_path = os.path.join(result_dir, f'{os.path.splitext(image_file)[0]}.txt')

            result = reader.readtext(image_path, detail=0)
            with open(result_file_path, 'w', encoding='utf-8') as result_file:
                result_file.write('\n'.join(result))

# 使用超分辨率进行图像放大
def upscale_images_with_command(input_dir, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    command = f'D:\Documents\PycharmProjects\make_information\generate_report\AI工具\models\\realesrgan-ncnn-vulkan-20220424-windows\\realesrgan-ncnn-vulkan.exe -i "{input_dir}" -o "{output_dir}"  -m D:\Documents\PycharmProjects\make_information\generate_report\AI工具\models\\realesrgan-ncnn-vulkan-20220424-windows\models\\'
    process = subprocess.Popen(['cmd', '/c', command], stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
    # 获取命令输出
    stdout, stderr = process.communicate()
    print("Standard Output:\n", stdout.decode('utf-8'))
    print("Standard Error:\n", stderr.decode('utf-8'))

# 从PDF中提取文本并保存结果
def extract_text_from_pdf(pdf_path, result_dir):
    pdf_document = fitz.open(pdf_path)

    if not os.path.exists(result_dir):
        os.makedirs(result_dir)

    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        text = page.get_text("text")  # 注意这里是get_text而不是getText

        result_file_path = os.path.join(result_dir, f'page_{page_num + 1}.txt')
        with open(result_file_path, 'w', encoding='utf-8') as result_file:
            result_file.write(text)

    pdf_document.close()

# 处理和纠正文本
def process_and_correct_text(text):
    prompt = "你是一位国家标准起草人，请修改这里面的错误并按照国家标准的格式整理文档：\n"
    input_text = prompt + text
    response = openai.Completion.create(
        model="gpt-3.5-turbo-16k",
        prompt=input_text,
        max_tokens=4096,
        stop=None
    )
    corrected_text = response.choices[0].text.strip()
    return corrected_text

# 处理文件夹中的文本文件
def process_text_files_in_folder(folder_path):
    for filename in os.listdir(folder_path):
        if filename.endswith('.txt'):
            file_path = os.path.join(folder_path, filename)
            with open(file_path, 'r', encoding='utf-8') as file:
                original_text = file.read()
            corrected_text = process_and_correct_text(original_text)
            with open(file_path, 'w', encoding='utf-8') as file:
                file.write(corrected_text)

# 主函数
def main():
    # convert_pdf_to_images(pdf_path, output_dir)
    # upscale_images_with_command(output_dir, upscale_dir)
    # perform_ocr_and_save_results(upscale_dir, result_dir)
    # extract_text_from_pdf(pdf_path,result_dir)
    process_text_files_in_folder(result_dir)

if __name__ == '__main__':
    main()
