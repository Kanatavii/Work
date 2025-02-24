import time
import openpyxl
from collections import defaultdict

# 你的 Excel 文件路径
file_path = "Z:\\UOF\\转运数据\\UOF出入库汇总表.xlsx"

# 记录计算时间的字典
column_calc_times = defaultdict(float)
formula_counts = defaultdict(int)

# 读取 Excel 文件
wb = openpyxl.load_workbook(file_path, data_only=False)

# 选择要分析的工作表
sheet_name = "货物总单"
ws = wb[sheet_name]

# 遍历所有列，分析哪一列计算最耗时
for col in ws.iter_cols():
    # 跳过完全空白的列
    if all(cell.value is None for cell in col):
        continue
    
    column_formula_count = sum(1 for cell in col if isinstance(cell.value, str) and cell.value.startswith("="))
    
    if column_formula_count == 0:
        continue
    
    start_time = time.perf_counter()
    for cell in col:
        if isinstance(cell.value, str) and cell.value.startswith("="):
            try:
                eval_formula = wb.defined_names[cell.value[1:]] if cell.value[1:] in wb.defined_names else cell.value
            except Exception:
                eval_formula = cell.value
            _ = eval_formula  # 读取公式但不计算
    end_time = time.perf_counter()
    
    column_calc_times[col[0].column_letter] = end_time - start_time
    formula_counts[col[0].column_letter] = column_formula_count

# 仅打印有公式的列的计算时间和公式数量，并按计算时间排序
sorted_columns = sorted(column_calc_times.items(), key=lambda x: x[1], reverse=True)
for col, calc_time in sorted_columns:
    print(f"Column: {col}, Calculation Time: {calc_time:.6f} seconds, Formula Count: {formula_counts[col]}")
