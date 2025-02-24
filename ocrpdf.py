import fitz  # PyMuPDF
from PIL import Image, ImageTk
import pytesseract
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import cv2
import numpy as np

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# 转换全角数字为半角数字的函数
def fullwidth_and_circled_to_halfwidth(text):
    # 带圈数字的映射表（①->1, ②->2,..., ⑨->9, ⑩->0）
    circled_digit_map = str.maketrans("①②③④⑤⑥⑦⑧⑨⑩", "1234567890")
    # 全角数字映射表（０１２３４５６７８９ -> 0123456789）
    fullwidth_map = str.maketrans("０１２３４５６７８９", "0123456789")
    
    # 先将带圈数字转换为半角数字
    text = text.translate(circled_digit_map)
    # 再将全角数字转换为半角数字
    text = text.translate(fullwidth_map)
    return text

# 图像预处理函数
def preprocess_image(img):
    img_cv = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2GRAY)  # 转灰度
    _, img_thresh = cv2.threshold(img_cv, 150, 255, cv2.THRESH_BINARY)  # 二值化
    return Image.fromarray(img_thresh)

# 提取 JBX 开头7位数字的函数
def extract_jbx_numbers_from_pdf(pdf_path):
    pdf_document = fitz.open(pdf_path)  # 打开PDF文件
    jbx_results = {}  # 保存每页的提取结果

    for page_number in range(len(pdf_document)):
        try:
            # 加载PDF页面并转换为高分辨率图像
            page = pdf_document.load_page(page_number)
            pix = page.get_pixmap(dpi=300)  # 提高分辨率为300 DPI
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # 图像预处理
            img = preprocess_image(img)

            # OCR 识别页面文本
            text = pytesseract.image_to_string(img, lang="jpn")
            print(f"Page {page_number + 1} OCR Output:\n{text}\n{'=' * 50}")  # 调试输出

            # 使用正则表达式匹配 JBX 开头后跟任意7个字符
            raw_matches = re.findall(r'JBX.{7}', text)
            print(f"Page {page_number + 1} Raw Matches: {raw_matches}")  # 调试输出

            # 转换结果：转换全角和带圈数字，并去除非数字字符
            converted_matches = []
            for match in raw_matches:
                jbx_part = match[3:]  # 提取 'JBX' 后的部分
                cleaned_match = fullwidth_and_circled_to_halfwidth(jbx_part)  # 转换带圈和全角数字
                cleaned_match = re.sub(r'\D', '', cleaned_match)  # 去除非数字字符

                if len(cleaned_match) == 7:  # 确保数字长度为7
                    converted_matches.append("JBX" + cleaned_match)

            print(f"Page {page_number + 1} Converted Matches: {converted_matches}")

            # 保存结果
            if converted_matches:
                jbx_results[page_number + 1] = converted_matches

        except Exception as e:
            print(f"Error on page {page_number + 1}: {e}")

    pdf_document.close()
    return jbx_results

# UI 主界面
class OCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF OCR")

        self.frame_left = tk.Frame(root, width=300, bg="lightgray")
        self.frame_left.pack(side="left", fill="both", expand=True)

        self.pdf_canvas = tk.Canvas(self.frame_left, bg="white")
        self.pdf_canvas.pack(fill="both", expand=True)

        self.frame_right = tk.Frame(root, width=400, bg="white")
        self.frame_right.pack(side="right", fill="both", expand=True)

        self.result_text = tk.Text(self.frame_right, font=("Arial", 12), wrap="word")
        self.result_text.pack(fill="both", expand=True)

        self.btn_open = ttk.Button(root, text="导入 PDF", command=self.open_pdf)
        self.btn_open.pack(side="top", pady=10)

    def open_pdf(self):
        pdf_path = filedialog.askopenfilename(title="选择PDF文件", filetypes=[("PDF 文件", "*.pdf")])

        if not pdf_path:
            messagebox.showwarning("未选择文件", "请选择一个PDF文件！")
            return

        self.display_pdf_pages(pdf_path)
        self.display_ocr_results(pdf_path)

    def display_pdf_pages(self, pdf_path):
        self.pdf_canvas.delete("all")
        pdf_document = fitz.open(pdf_path)
        y_offset = 0

        for page_number in range(min(5, len(pdf_document))):
            page = pdf_document.load_page(page_number)
            pix = page.get_pixmap(dpi=50)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img_tk = ImageTk.PhotoImage(img)

            self.pdf_canvas.create_image(10, y_offset, anchor="nw", image=img_tk)
            self.pdf_canvas.image = img_tk
            y_offset += pix.height + 10

        pdf_document.close()

    def display_ocr_results(self, pdf_path):
        self.result_text.delete(1.0, tk.END)
        jbx_results = extract_jbx_numbers_from_pdf(pdf_path)

        if not jbx_results:
            self.result_text.insert(tk.END, "未找到任何 JBX 开头的7位数字。")
            return

        for page, matches in jbx_results.items():
            self.result_text.insert(tk.END, f"第 {page} 页:\n")
            for match in matches:
                self.result_text.insert(tk.END, f"  {match}\n")
            self.result_text.insert(tk.END, "\n")

if __name__ == "__main__":
    root = tk.Tk()
    app = OCRApp(root)
    root.mainloop()
