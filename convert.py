import pandas as pd
from tkinter import Tk, filedialog

# 定义新表格的列头
new_headers = [
    "出货仓库/Delivery Warehouse", "运输方式/TransportMode", "参考单号/Ref No",
    "目的国家/Destination Country", "收件人姓名/Recipient Name", "地址/Address",
    "电话/Phone", "SKU/SKU", "数量/Count", "收件人公司/Recipient Company",
    "州省/State Province", "城市/City", "县区/District", "备注/Remark",
    "邮政编码/Postcode", "邮箱/Email"
]

def process_excel(file_path):
    try:
        # 读取原始表格，从第二行开始读取，忽略第一行
        df = pd.read_excel(file_path, header=None, skiprows=1, dtype=str)  # 确保所有列都以字符串格式读取

        # 创建新表格
        new_df = pd.DataFrame(columns=new_headers)

        # 转换第一列内容
        def transform_transport_mode(mode):
            if mode == "佐川":
                return "YJYMT"
            elif mode == "投函":
                return "YJDM"
            else:
                return "未知"

        # 填充数据到新表格
        new_df["出货仓库/Delivery Warehouse"] = ["大阪1号仓"] * len(df)  # 固定值
        new_df["运输方式/TransportMode"] = df[0].map(transform_transport_mode)  # 第一列转换后填充到"运输方式"
        new_df["参考单号/Ref No"] = df[1]  # 第二列填充到"参考单号"
        new_df["目的国家/Destination Country"] = ["JP"] * len(df)  # 固定值
        new_df["收件人姓名/Recipient Name"] = df[6]  # 第七列填充到"收件人姓名"
        new_df["地址/Address"] = df[5]  # 第六列填充到"地址"
        new_df["电话/Phone"] = df[3]  # 第四列填充到"电话"
        new_df["邮政编码/Postcode"] = df[4]  # 第五列填充到"邮政编码"

        # 其他列可以根据需求填充默认值或保持为空

        # 保存新表格
        output_file = file_path.replace(".xlsx", "_converted.xlsx")
        new_df.to_excel(output_file, index=False)
        print(f"转换后的文件已保存为: {output_file}")

    except Exception as e:
        print(f"处理过程中出错: {e}")

# 创建文件选择界面
def select_file():
    root = Tk()
    root.withdraw()  # 隐藏主窗口

    file_path = filedialog.askopenfilename(
        title="选择一个Excel文件",
        filetypes=[("Excel 文件", "*.xlsx;*.xls")]
    )

    if file_path:
        print(f"选择的文件是: {file_path}")
        process_excel(file_path)
    else:
        print("未选择文件。")

if __name__ == "__main__":
    select_file()

