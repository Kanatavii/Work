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

        # 遍历原始数据行
        for _, row in df.iterrows():
            # 分解SKU和数量
            sku_quantity_str = row[8]  # 第9列是SKU和数量
            if pd.isna(sku_quantity_str):
                continue

            sku_quantity_pairs = sku_quantity_str.split("+")  # 以加号分隔多个SKU

            for pair in sku_quantity_pairs:
                if "*" in pair:
                    sku, quantity = pair.split("*")  # 以星号分隔SKU和数量
                else:
                    sku, quantity = pair, "1"  # 如果没有数量，默认数量为1

                # 创建新行并填充数据
                new_row = {
                    "出货仓库/Delivery Warehouse": "大阪1号仓",
                    "运输方式/TransportMode": transform_transport_mode(row[0]),  # 第一列转换后填充到"运输方式"
                    "参考单号/Ref No": row[1],  # 第二列填充到"参考单号"
                    "目的国家/Destination Country": "JP",  # 固定值
                    "收件人姓名/Recipient Name": row[6],  # 第七列填充到"收件人姓名"
                    "地址/Address": row[5],  # 第六列填充到"地址"
                    "电话/Phone": row[3],  # 第四列填充到"电话"
                    "SKU/SKU": sku.strip(),  # SKU放到第8列
                    "数量/Count": quantity.strip(),  # 数量放到第9列
                    "邮政编码/Postcode": row[4]  # 第五列填充到"邮政编码"
                }
                new_df = pd.concat([new_df, pd.DataFrame([new_row])], ignore_index=True)

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