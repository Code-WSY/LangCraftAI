import pdfplumber
import sys
import tkinter as tk
from tkinter import filedialog
def pdf_to_text(pdf_path):
    text = ''
    try:
        # 打开PDF文件
        with pdfplumber.open(pdf_path) as pdf:
            # 遍历每一页
            for page in pdf.pages:
                # 提取当前页面的文本
                page_text = page.extract_text()
                if page_text:
                    text += page_text + '\n'
    except Exception as e:
        print("处理PDF文件时出错：", e)
        sys.exit(1)
    return text

def select_pdf():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()
    return file_path
def save_text(text, pdf_path):
    save_name=pdf_path.split('/')[-1].split('.')[0]
    save_name=save_name+'.txt'
    with open(save_name, 'w', encoding='utf-8') as f:
        f.write(text)
    print('文本已保存为：', save_name)
if __name__ == '__main__':
    pdf_path = select_pdf()
    text = pdf_to_text(pdf_path)
    save_text(text, pdf_path)

