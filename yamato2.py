import tkinter as tk
from tkinter import messagebox
import requests
import time
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
import json

def query_kuroneko_yamato(track_number):
    url = "https://member.kms.kuronekoyamato.co.jp/api/receive_parcel/v1/getParcelDetailInfo"
    params = {
        "X-NEKO-TRACE": "f6b6a1f3-7192-4ec8-aa67-d8066cc08845",
        "X-NEKO-DEVICE": "a8136f5d-9473-46a8-ac80-b036e7cf0cde"
    }

    headers = {
        "accept": "application/json, text/plain, */*",
        "content-type": "application/json;charset=UTF-8",
        "origin": "https://member.kms.kuronekoyamato.co.jp",
        "referer": f"https://member.kms.kuronekoyamato.co.jp/parcel/detail?pno={track_number}",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        "x-pg-id": "NRCWBRP0610",
        "x-pg_id": "NRCWBRP0610",
        "x-system-id": "NRC"
    }

    data = {
        "nmtJho": {
            "okjNo": track_number,
            "uktHasKbn": "01"
        },
        "ktuJho": {
            "iriTs": "2024-12-07T22:54:00.139+09:00",
            "tnmSbtCd": "1"
        }
    }

    response = requests.post(url, headers=headers, params=params, json=data, timeout=10)
    response.raise_for_status()

    # Print the raw response as text to see if there are any discrepancies
    print(f"Raw response text: {response.text}")

    # Attempt to load the response text as JSON
    try:
        raw_json_data = json.loads(response.text)
        print(f"Manually parsed API response: {raw_json_data}")

        # Directly access the 'result' section and print it
        result = raw_json_data["result"]
        print(f"Result content: {json.dumps(result, ensure_ascii=False)}")

        # Accessing the 'nmtJho' section explicitly to inspect its content
        nmt_jho = result.get("nmtJho", {})
        print(f"nmtJho content: {json.dumps(nmt_jho, ensure_ascii=False)}")
        print(f"nmtJho keys: {list(nmt_jho.keys())}")

        # Check for 'nmtHttJkyLst' directly
        if 'nmtHttJkyLst' in result:
            nmt_htt_jky_lst = result['nmtHttJkyLst']
            print(f"Contents of nmtHttJkyLst (Raw): {json.dumps(nmt_htt_jky_lst, ensure_ascii=False)}")

            if len(nmt_htt_jky_lst) > 0:
                print(f"First element in nmtHttJkyLst: {nmt_htt_jky_lst[0]}")
                delivery_time = nmt_htt_jky_lst[0].get("nmtStsHenMdYbtk", "未找到时间")
                print(f"Extracted Delivery Time from nmtStsHenMdYbtk: {delivery_time}")
            else:
                # If the list is empty, fallback to 'hsoJkyMdl' field
                delivery_time = "未找到时间"
                print("nmtHttJkyLst is empty or doesn't contain the time. Fallback to delivery status.")
        else:
            print("Key 'nmtHttJkyLst' does NOT exist in the response.")
            delivery_time = "未找到时间"

        # Get the delivery status from 'hsoJkyMdl'
        delivery_status = nmt_jho.get("hsoJkyMdl", "未找到配送状况")

        # Final debug prints
        print(f"Final Delivery Time: {delivery_time}")
        print(f"Delivery Status: {delivery_status}")

    except json.JSONDecodeError as e:
        print(f"Error parsing JSON: {e}")

    return {
        "单号": track_number,
        "时间": delivery_time,
        "配送状况": delivery_status
    }

def run_queries():
    track_numbers = text_input.get("1.0", tk.END).strip().splitlines()
    if not track_numbers:
        messagebox.showwarning("警告", "请输入至少一个单号")
        return

    start_time = time.time()
    results = []

    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = [executor.submit(query_kuroneko_yamato, num.strip()) for num in track_numbers if num.strip()]
        for future in futures:
            try:
                res = future.result()
                results.append(res)
            except Exception as e:
                results.append({"单号": "未知", "时间": "查询失败", "配送状况": f"错误: {e}"})

    end_time = time.time()
    elapsed = end_time - start_time

    df = pd.DataFrame(results)
    df.to_csv("results.csv", index=False, encoding="utf-8-sig")

    messagebox.showinfo("完成", f"查询完成，共处理 {len(results)} 个单号。\n耗时：{elapsed:.2f} 秒\n结果已导出到 results.csv")

root = tk.Tk()
root.title("黑猫")
root.geometry("500x400")

label = tk.Label(root, text="请输入要查询的单号（每行一个）:")
label.pack(pady=10)

text_input = tk.Text(root, width=50, height=15)
text_input.pack(pady=10)

run_button = tk.Button(root, text="运行查询", command=run_queries)
run_button.pack(pady=10)

root.mainloop()
