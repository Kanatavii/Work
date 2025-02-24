import win32com.client as win32
from datetime import datetime
import pandas as pd
import os
import time
import tkinter as tk
from tkinter import filedialog

# **1. 让用户选择文件**
def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="选择数据文件",
        filetypes=[("Excel 文件 (*.xlsx, *.xls)", "*.xlsx;*.xls"), ("CSV 文件 (*.csv)", "*.csv")]
    )
    if not file_path:
        print("❌ 未选择文件，程序退出。")
        exit()
    return file_path

file_path = select_file()

# **2. 生成文件名**
current_datetime = datetime.now().strftime("%Y-%m-%d-%H%M")
DOWNLOAD_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads")
excel_output_path = os.path.join(DOWNLOAD_FOLDER, f"流山入库-{current_datetime}.xlsx")
pdf_output_path = os.path.join(DOWNLOAD_FOLDER, f"流山入库-{current_datetime}.pdf")

# **3. 读取数据**
file_extension = os.path.splitext(file_path)[1].lower()
if file_extension in [".xls", ".xlsx"]:
    df = pd.read_excel(file_path, engine="openpyxl")
elif file_extension == ".csv":
    df = pd.read_csv(file_path, encoding="utf-8")
else:
    print("❌ 不支持的文件类型，程序退出。")
    exit()

# **4. 筛选数据**
filtered_df = df[df.iloc[:, 2].notna()]
final_df = filtered_df[filtered_df.iloc[:, 3].isna()]

# **5. 只保留指定的列**
columns_to_keep = ["许可时间", "回数", "送り状番号", "箱数", "转运公司", "转运备注", "现场用-函数对应", "入库时间", "取件地", "数据用"]
final_df = final_df[columns_to_keep]

# **6. 解决许可时间存储为 Excel 序列号的问题**
print("📊 许可时间列转换前：")
print(final_df["许可时间"].head())  # 检查转换前数据

# **确保许可时间列是字符串类型，转换为数值**
try:
    final_df["许可时间"] = final_df["许可时间"].astype(float)  # 转换为 float
    final_df["许可时间"] = pd.to_datetime(final_df["许可时间"], origin="1899-12-30", unit="D")
    print("✅ 许可时间转换成功！")
except Exception as e:
    print(f"⚠️ 许可时间转换失败：{e}")
    final_df["许可时间"] = pd.to_datetime(final_df["许可时间"], errors='coerce')  # 兜底转换

# **确保最终 Excel 里显示为 `YYYY-MM-DD` 格式**
final_df["许可时间"] = final_df["许可时间"].dt.strftime("%Y-%m-%d")

# **7. 保存 Excel**
final_df.to_excel(excel_output_path, index=False, engine="openpyxl")
print(f"✅ 数据已保存到 Excel：{excel_output_path}")

# **8. 使用 pywin32 处理 Excel**
print("🔄 正在格式化 Excel 并导出 PDF...")
excel = win32.DispatchEx('Excel.Application')
time.sleep(1)
wb = excel.Workbooks.Open(excel_output_path)
ws = wb.Worksheets(1)

# **9. 设置 A 列为日期格式**
ws.Columns("A").NumberFormat = "m/d/yyyy"

# **10. 设置列宽**
column_widths = {'A': 11, 'B': 4.63, 'C': 13.25, 'D': 4.63, 'E': 18, 'F': 28, 'G': 58.13, 'H': 19.88, 'I': 15, 'J': 15}
for col, width in column_widths.items():
    ws.Columns(col).ColumnWidth = width

# **11. 添加全边框**
thin_border = 2
last_row = final_df.shape[0] + 1  # +1 计算表头

for row in range(1, last_row + 1):  
    for col in range(1, 11):  
        cell = ws.Cells(row, col)
        cell.Borders.LineStyle = 1  # 强制添加边框
        cell.Borders.Weight = thin_border

# **12. 设置打印区域**
ws.PageSetup.PrintArea = f"A1:J{last_row}"
ws.PageSetup.Orientation = 2
ws.PageSetup.PaperSize = 9
ws.PageSetup.Zoom = False
ws.PageSetup.FitToPagesWide = 1
ws.PageSetup.FitToPagesTall = 1

# **13. 手动刷新 Excel**
excel.CalculateFullRebuild()

# **14. 保存并导出 PDF**
wb.Save()
wb.ExportAsFixedFormat(0, pdf_output_path)
wb.Close(False)
excel.Application.Quit()

# **15. 清理**
del excel
print(f"✅ Excel 文件已导出到: {excel_output_path}")
print(f"✅ PDF 文件已导出到: {pdf_output_path}")
