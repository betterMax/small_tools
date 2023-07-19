import pdfplumber

def parse_pdf(file_path):
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                # 处理提取到的文本，可以进行进一步的处理或分析
                print(text)
    except Exception as e:
        print("解析PDF文件时出错：", str(e))

# 调用函数并传入PDF文件路径
pdf_file_path = r"D:\OneDrive\Work\理工环科\General Information\资源\水质论文\SWAT模型下“十三五”中...件对岷江流域水环境改善贡献_刘骞.pdf"
parse_pdf(pdf_file_path)
