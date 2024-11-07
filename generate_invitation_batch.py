import os
from docx import Document
from docx2pdf import convert

# 人名列表 (您可以从文件读取)
names = []
# names = ["Alice Smith", "Bob Johnson", "Charlie Brown"]  # 示例人名列表
with open("nameList.txt", 'r') as fh:
    for i in fh:
        names.append(i.strip())

# 模板Word文件路径
template_path = 'template.docx'  # 替换为您的模板文件路径
# 创建目录以保存Word和PDF文件
output_dir = "speakers_invitation"
output_dir = os.path.abspath(output_dir)
os.makedirs(output_dir, exist_ok=True)

# 处理每个人名
for name in names:
    # 读取模板
    doc = Document(template_path)

    # 遍历文档中的每一个段落和文本框，替换占位符
    for paragraph in doc.paragraphs:
        if '{{name}}' in paragraph.text:  # 假设模板中占位符为{{name}}
            paragraph.text = paragraph.text.replace('{{name}}', name)

    # 生成Word文件
    filename = f"Invitation letter to {name}.docx"
    word_file_path = os.path.join(output_dir, f'{filename}.docx')
    doc.save(word_file_path)

    # 转换为PDF文件
    #pdf_file_path = os.path.join(output_dir, f'{filename}.pdf')
    #convert(word_file_path, pdf_file_path)

    # 转换为PDF文件
    #try:
    #    convert(word_file_path, pdf_file_path)
    #    print(f"成功转换: {word_file_path} 到 {pdf_file_path}")
    #except Exception as e:
    #    print(f"转换失败: {word_file_path} 到 {pdf_file_path}，错误信息: {e}")

print("所有任务已完成。")

