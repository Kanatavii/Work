import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import json

def parse_excel_to_json(excel_file, json_file):
    """
    将指定格式的 Excel 文件转换为 JSON 文件。
    
    :param excel_file: 输入的 Excel 文件路径
    :param json_file: 输出的 JSON 文件路径
    """
    # 读取 Excel 文件
    df = pd.read_excel(excel_file, header=None)

    # 解析分组和具体县
    regions = {}
    for col in range(1, len(df.columns)):
        group = str(df.iloc[0, col]).strip()  # 第一行分组名称
        sub_regions = str(df.iloc[1, col]).strip().split("\n")  # 第二行的具体县
        for region in sub_regions:
            regions[region] = col  # 记录县对应的列号

    # 创建 JSON 数据结构
    data = {region: {} for region in regions.keys()}

    # 从第三行开始处理价格数据
    for _, row in df.iloc[2:].iterrows():
        size = str(row[0])  # 第一列是尺寸
        weight = row[1]  # 第二列是重量范围
        for region, col in regions.items():  # 遍历每个县和对应的列号
            price = row[col]
            if pd.notna(price):  # 如果价格非空
                if size not in data[region]:
                    data[region][size] = {}
                data[region][size] = {
                    "weight": weight,
                    "price": price
                }

    # 保存到 JSON 文件
    with open(json_file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

    print(f"JSON 文件已成功保存到 {json_file}")

def convert_excel_to_json():
    """
    图形界面功能：选择 Excel 文件并保存为 JSON 文件。
    """
    # 选择 Excel 文件
    excel_file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not excel_file:
        messagebox.showerror("Error", "未选择任何 Excel 文件！")
        return

    # 选择 JSON 文件保存路径
    json_file = filedialog.asksaveasfilename(
        title="Save JSON File",
        defaultextension=".json",
        filetypes=[("JSON Files", "*.json")]
    )
    if not json_file:
        messagebox.showerror("Error", "未选择保存路径！")
        return

    try:
        parse_excel_to_json(excel_file, json_file)
        messagebox.showinfo("Success", f"JSON 文件已成功保存至：\n{json_file}")
    except Exception as e:
        messagebox.showerror("Error", f"出现错误：\n{e}")

# 创建主窗口
root = tk.Tk()
root.title("Excel 转 JSON 转换器")

# 添加按钮
select_button = tk.Button(root, text="选择 Excel 并转换", command=convert_excel_to_json, width=30, height=2)
select_button.pack(pady=20)

# 启动应用程序
root.mainloop()
