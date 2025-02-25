import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import json

def convert_excel_to_json():
    """
    将选择的 Excel 文件转换为 JSON 文件。
    """
    # 获取 Excel 文件路径
    excel_file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not excel_file:
        messagebox.showerror("Error", "No file selected!")
        return

    # 获取 JSON 文件保存路径
    json_file = filedialog.asksaveasfilename(
        title="Save JSON File",
        defaultextension=".json",
        filetypes=[("JSON Files", "*.json")]
    )
    if not json_file:
        messagebox.showerror("Error", "No save location selected!")
        return

    try:
        # 读取 Excel 数据
        df = pd.read_excel(excel_file)

        # 转换为 JSON 格式的字典
        data = {}
        for _, row in df.iterrows():
            location = row["地点"]
            data[location] = {}
            for col in df.columns[1:]:  # 从第2列开始读取尺寸和价格
                size = str(col)
                price = row[col]
                data[location][size] = price

        # 保存为 JSON 文件
        with open(json_file, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

        messagebox.showinfo("Success", f"JSON file saved successfully:\n{json_file}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

# 创建主窗口
root = tk.Tk()
root.title("Excel to JSON Converter")

# 添加按钮
select_button = tk.Button(root, text="Select Excel and Convert", command=convert_excel_to_json, width=30, height=2)
select_button.pack(pady=20)

# 启动应用程序
root.mainloop()
