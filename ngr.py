import os
import time
import subprocess
import psutil
import requests
import pandas as pd
import win32com.client as win32
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

# **1. 获取当前用户的 Downloads 目录**
DOWNLOAD_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads")

# **2. 重要参数**
CHROME_DEBUG_PORT = 9222
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
USER_DATA_DIR = r"C:\selenium_chrome"
WEIXIN_DOC_URL = "https://doc.weixin.qq.com/sheet/e3_AOQAHAbJAFI2aphULNdRdOZragyT3"

# 获取当前时间，命名文件
current_datetime = datetime.now().strftime("%Y-%m-%d-%H%M")
excel_output_path = os.path.join(DOWNLOAD_FOLDER, f"流山入库-{current_datetime}.xlsx")
pdf_output_path = os.path.join(DOWNLOAD_FOLDER, f"流山入库-{current_datetime}.pdf")

# **3. 检查 Chrome 远程调试模式**
def is_chrome_debug_running():
    try:
        response = requests.get(f"http://127.0.0.1:{CHROME_DEBUG_PORT}/json", timeout=3)
        return response.status_code == 200
    except requests.exceptions.RequestException:
        return False

# **4. 启动 Chrome 远程调试模式**
def start_chrome():
    if is_chrome_debug_running():
        print("✅ Chrome 远程调试已在运行，直接复用会话。")
    else:
        print("🚀 Chrome 未运行，正在启动远程调试模式...")
        command = f'"{CHROME_PATH}" --remote-debugging-port={CHROME_DEBUG_PORT} --user-data-dir="{USER_DATA_DIR}"'
        subprocess.Popen(command, shell=True)
        time.sleep(5)

# **5. 运行 Chrome**
start_chrome()

# **6. 复用已登录的 Chrome**
options = webdriver.ChromeOptions()
options.debugger_address = f"127.0.0.1:{CHROME_DEBUG_PORT}"
driver = webdriver.Chrome(options=options)

print("✅ 成功连接 Chrome！")

# **7. 打开微信文档**
print("🔄 正在打开微信文档...")
driver.get(WEIXIN_DOC_URL)
time.sleep(5)

# **8. 点击菜单按钮**
try:
    print("🔄 寻找并点击菜单按钮...")
    menu_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CLASS_NAME, "wedocs-icon-tdoc-titlebar-menu"))
    )
    menu_button.click()
    print("✅ 菜单已打开")
except Exception as e:
    print("❌ 未找到菜单按钮:", e)
    driver.quit()
    exit()

# **9. 点击“导出”选项**
try:
    print("🔄 寻找并点击‘导出’选项...")
    export_option = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CLASS_NAME, "mainmenu-submenu-exportAs"))
    )
    export_option.click()
    print("✅ 选中‘导出’")
except Exception as e:
    print("❌ 未找到‘导出’选项:", e)
    driver.quit()
    exit()

# **10. 点击“本地CSV文件”**
try:
    print("🔄 寻找并点击‘本地CSV文件’...")
    csv_option = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CLASS_NAME, "mainmenu-item-export-csv"))
    )
    csv_option.click()
    print("✅ 已点击‘本地CSV文件’")
except Exception as e:
    print("❌ 未找到‘本地CSV文件’选项:", e)
    driver.quit()
    exit()

# **11. 等待文件下载**
time.sleep(10)

# **12. 获取最新下载的 CSV 文件**
csv_files = [f for f in os.listdir(DOWNLOAD_FOLDER) if f.endswith(".csv")]
if csv_files:
    csv_files.sort(key=lambda x: os.path.getctime(os.path.join(DOWNLOAD_FOLDER, x)), reverse=True)
    latest_file = os.path.join(DOWNLOAD_FOLDER, csv_files[0])
    print(f"📂 最新下载的 CSV 文件：{latest_file}")
else:
    print("⚠️ 没有找到下载的 CSV 文件，请检查下载目录！")
    driver.quit()
    exit()

# **13. 读取 CSV 并处理数据**
print("🔄 正在处理数据...")
df = pd.read_csv(latest_file)

# **14. 筛选第三列非空白的行**
filtered_df = df[df.iloc[:, 2].notna()]

# **15. 再筛选第四列为空白的行**
final_df = filtered_df[filtered_df.iloc[:, 3].isna()]

# **16. 强制确保 A 列（许可时间）为 datetime 格式**
final_df["许可时间"] = pd.to_datetime(final_df["许可时间"], errors='coerce')

# **17. 保存 Excel**
final_df.to_excel(excel_output_path, index=False)
print(f"✅ 数据已保存到 Excel：{excel_output_path}")

# **18. 使用 pywin32 处理 Excel**
print("🔄 正在格式化 Excel 并导出 PDF...")
excel = win32.DispatchEx('Excel.Application')
time.sleep(1)
wb = excel.Workbooks.Open(excel_output_path)
ws = wb.Worksheets(1)

# **19. 设置 A 列为日期格式**
ws.Columns("A").NumberFormat = "m/d/yyyy"

# **20. 设置列宽**
column_widths = {'A': 11, 'B': 4.63, 'C': 13.25, 'D': 4.63, 'E': 18, 'F': 28, 'G': 58.13, 'H': 19.88, 'I': 15, 'J': 15}
for col, width in column_widths.items():
    ws.Columns(col).ColumnWidth = width

# **21. 添加边框**
thin_border = 2
used_range = ws.UsedRange
for row in range(1, used_range.Rows.Count + 1):
    if ws.Cells(row, 1).Value is not None:
        for col in range(1, 11):
            ws.Cells(row, col).Borders.Weight = thin_border

# **22. 设置打印区域**
ws.PageSetup.PrintArea = f"A1:J{used_range.Rows.Count}"
ws.PageSetup.Orientation = 2
ws.PageSetup.PaperSize = 9
ws.PageSetup.Zoom = False
ws.PageSetup.FitToPagesWide = 1
ws.PageSetup.FitToPagesTall = 1

# **23. 保存并导出 PDF**
wb.Save()
wb.ExportAsFixedFormat(0, pdf_output_path)
wb.Close(False)
excel.Application.Quit()

# **24. 清理**
del excel
print(f"✅ Excel 文件已导出到: {excel_output_path}")
print(f"✅ PDF 文件已导出到: {pdf_output_path}")

# **25. 关闭 Selenium**
driver.quit()
print("🚀 脚本执行完毕！")
