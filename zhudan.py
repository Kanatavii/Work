import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import subprocess

# 定义主窗口
def select_file():
    # 打开文件选择对话框，选择导入文件
    import_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    
    if import_file_path:
        # 显示选择的导入文件路径
        file_label.config(text=f"导入文件: {import_file_path}")
        process_file(import_file_path)
    else:
        messagebox.showerror("错误", "未选择文件")

# 处理文件函数
def process_file(import_file_path):
    try:
        # 判断文件类型并选择合适的引擎
        if import_file_path.endswith('.xls'):
            import_data = pd.read_excel(import_file_path, engine='xlrd')  # 使用xlrd读取.xls文件
        else:
            import_data = pd.read_excel(import_file_path, engine='openpyxl')  # 使用openpyxl读取.xlsx文件

        # 创建一个固定结构的DataFrame用于导出数据，并保证所有列都存在
        export_data = pd.DataFrame({
            '到港时间': import_data['预计到港日期'] if '预计到港日期' in import_data else [''] * len(import_data),
            '搬入时间': [''] * len(import_data),
            '许可时间函数对应': [''] * len(import_data),
            '入库时间': [''] * len(import_data),
            '出库指示时间/入金确认时间': [''] * len(import_data),
            '转运日期': [''] * len(import_data),
            '转运单号': [''] * len(import_data),
            '回数': [''] * len(import_data),
            'FBA': import_data['FBA进仓编号'].str[:12] if 'FBA进仓编号' in import_data else [''] * len(import_data),
            '空运/海运': [''] * len(import_data),
            '取件地': [''] * len(import_data),
            '主单号': import_data['MAWB番号'] if 'MAWB番号' in import_data else [''] * len(import_data),
            '送り状番号': import_data['送り状番号'] if '送り状番号' in import_data else [''] * len(import_data),
            '箱数': import_data['PKG'] if 'PKG' in import_data else [''] * len(import_data),
            '重量(KG)': import_data['WEIGHT(KG)'] if 'WEIGHT(KG)' in import_data else [''] * len(import_data),
            '立方数': import_data['收货立方'] if '收货立方' in import_data else [''] * len(import_data),
            '尺寸': [''] * len(import_data),
            '转运公司': [''] * len(import_data),
            '转运备注': [''] * len(import_data),
            '现场用-函数对应': [''] * len(import_data),
            '乐天匹配用辅助函数': [''] * len(import_data),
            '郵便番号': import_data['荷受人郵便番号'] if '荷受人郵便番号' in import_data else [''] * len(import_data),
            '会社名/个人': import_data['荷受人漢字名'] if '荷受人漢字名' in import_data else [''] * len(import_data),
            '荷受人住所': import_data['荷受人住所'] if '荷受人住所' in import_data else [''] * len(import_data),
            '荷受人电话': import_data['荷受人電話番号'] if '荷受人電話番号' in import_data else [''] * len(import_data),
            '乐天5位数担当者': import_data['荷受人担当者'] if '荷受人担当者' in import_data else [''] * len(import_data)
        })

        # 处理‘额外服务’列，转移到‘转运备注’列，并处理‘合并’情况
        if '额外服务' in import_data:
            export_data['转运备注'] = import_data['额外服务']
            for index, service in import_data['额外服务'].items():
                if '合并' in str(service):
                    # 将原先合并的单元格信息分别写在分开的单元格内
                    export_data.loc[index, '转运备注'] = f"合并派送 {service}"
                    export_data.loc[index + 1, '转运备注'] = f"合并派送 {service}"

        # 生成输出文件名，添加"-数据用"后缀
        folder = os.path.dirname(import_file_path)
        filename = os.path.splitext(os.path.basename(import_file_path))[0]
        output_file_path = os.path.join(folder, f"{filename}-数据用.xlsx")
        
        # 将处理后的数据导出到新的Excel文件中，保留空列
        export_data.to_excel(output_file_path, index=False)
        
        # 显示成功消息
        messagebox.showinfo("成功", f"文件已成功保存至: {output_file_path}")
        
        # 自动打开生成的文件
        if os.name == 'nt':  # 如果是Windows系统
            os.startfile(output_file_path)
        else:  # macOS 或 Linux
            subprocess.call(['open', output_file_path])

    except Exception as e:
        messagebox.showerror("错误", f"文件处理失败: {e}")

# 创建UI界面
root = tk.Tk()
root.title("Excel 文件处理")

# 创建选择导入文件按钮
file_button = tk.Button(root, text="选择导入文件", command=select_file)
file_button.pack(pady=20)

# 显示选择的文件路径标签
file_label = tk.Label(root, text="未选择文件")
file_label.pack(pady=10)

# 运行主窗口循环
root.mainloop()
