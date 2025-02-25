import requests
import json
import tkinter as tk
from tkinter import messagebox
import threading
from concurrent.futures import ThreadPoolExecutor
import pandas as pd
import time

API_URL = "https://cargotracking.logiquest.co.jp/api/v1/house/hawb/{tracking_number}"

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36",
    "Referer": "https://cargotracking.logiquest.co.jp/",
    "Accept": "application/json, text/plain, */*",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Accept-Language": "en,zh-CN;q=0.9,zh;q=0.8,ja;q=0.7"
}

cookies = {
    "csrftoken": "84T4OdGDi5PQQI7W8awMEupK4sF7kHB4Xjk3y52meU2llqXBuhf7TC0XvSIyQElH",
    "sessionid": "7ornygl2zlyl9v03nw0v44757mxf2f8l"
}

def fetch_tracking_data(tracking_number):
    """ 直接从 XHR 请求数据，获取 h_out_datetime """
    url = API_URL.format(tracking_number=tracking_number)
    response = requests.get(url, headers=headers, cookies=cookies)
    
    print(f"\n查询单号: {tracking_number}")
    print(f"状态码: {response.status_code}")
    print(f"返回内容: {response.text[:500]}")  # 先打印前500字符，看看是什么数据

    if response.status_code == 200:
        try:
            data = response.json()  # 解析 JSON
            if "data" in data and len(data["data"]) > 0:
                return data["data"][0].get("h_out_datetime", "无结果")
            else:
                return "未找到相关数据"
        except json.JSONDecodeError:
            return "返回的不是 JSON 数据"
    else:
        return f"请求失败，状态码：{response.status_code}"

def process_number(tracking_number):
    return tracking_number.strip(), fetch_tracking_data(tracking_number.strip())

def export_to_excel(results):
    df = pd.DataFrame(results, columns=["单号", "出库时间"])
    file_path = "查询结果.xlsx"
    df.to_excel(file_path, index=False)

def search_multiple():
    input_text = text_box.get("1.0", tk.END).strip()
    if not input_text:
        messagebox.showerror("错误", "请输入至少一个单号！")
        return

    tracking_numbers = input_text.splitlines()
    start_time = time.time()

    def worker():
        results = []
        with ThreadPoolExecutor(max_workers=10) as executor:  
            futures = [executor.submit(process_number, number) for number in tracking_numbers]
            for future in futures:
                results.append(future.result())
        export_to_excel(results)
        end_time = time.time()
        elapsed_time = end_time - start_time
        num_queries = len(tracking_numbers)
        messagebox.showinfo(
            "统计结果", 
            f"总单号数量: {num_queries}\n总耗时: {elapsed_time:.2f} 秒\n平均每秒查询: {num_queries / elapsed_time:.2f} 单\n结果已导出到 查询结果.xlsx"
        )

    threading.Thread(target=worker).start()

# GUI 代码
root = tk.Tk()
root.title("LOGI批量查询")
root.geometry("600x400")

label = tk.Label(root, text="请输入单号，每行一个:")
label.pack(pady=10)

text_box = tk.Text(root, height=15, width=70)
text_box.pack(pady=10)

button = tk.Button(root, text="开始查询", command=search_multiple)
button.pack(pady=10)

root.mainloop()