import os
import win32com.client
import datetime
import logging
import tkinter as tk
from tkinter import messagebox

# 配置日志记录
logging.basicConfig(filename='app.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def show_messagebox(title, message):
    """
    显示强制最前的消息框
    """
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    root.attributes("-topmost", True)  # 设置窗口为最前
    messagebox.showinfo(title, message, parent=root)
    root.destroy()  # 销毁主窗口

def check_and_create_folder(folder_path):
    """
    检查文件夹路径是否存在，如果不存在就创建它。
    """
    os.makedirs(folder_path, exist_ok=True)

try:
    # 创建 Outlook 应用程序对象
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # 获取默认账户
    default_account = namespace.Accounts.Item(1)

    # 获取默认账户的收件箱
    inbox = None
    for folder in namespace.Folders:
        if folder.Name == default_account.DisplayName:  # 获取与默认账户相关的文件夹
            inbox = folder.Folders.Item("受信トレイ")  # 获取默认账户下的收件箱（"受信トレイ"是日文版本收件箱的名称）
            break

    if inbox is None:
        raise Exception("Failed to find the Inbox folder for the default account.")

    # 获取当前时间并计算2天前的时间
    two_days_ago = datetime.datetime.now() - datetime.timedelta(days=2)

    # 筛选指定的发件人邮箱地址
    specific_emails = [
        "s.morita@plutus-thl.com",
        "y.sagan@plutus-thl.com",
        "n.teruya@plutus-thl.com",
        "h.mizukawa@plutus-thl.com"
    ]
    email_filter = " OR ".join([f'\"urn:schemas:httpmail:fromemail\" = \'{email}\'' for email in specific_emails])

    # 设置SQL过滤器，获取接收时间在过去2天内的邮件，并且来自指定发件人
    filter_str = f"@SQL=({email_filter}) AND \"urn:schemas:httpmail:datereceived\" >= '{two_days_ago.strftime('%m/%d/%Y %H:%M %p')}'"

    # 获取符合条件的邮件
    all_items = inbox.Items
    all_items.Sort("[ReceivedTime]", True)  # 按照接收时间降序排序
    filtered_items = all_items.Restrict(filter_str)

    if len(filtered_items) == 0:
        show_messagebox("没有新邮件", "未找到符合条件的邮件。")
    else:
        # 创建附件保存的文件夹
        attachment_folder = f"Z:\\UOF\\转运数据\\JHSS\\{default_account.DisplayName}"
        check_and_create_folder(attachment_folder)

        # 新许可提示框
        new_license_message = ""

        # 处理符合条件的邮件
        for item in filtered_items:
            # 获取邮件接收的时间并格式化
            received_time = item.ReceivedTime.strftime("%Y%m%d_%H%M%S")
            time_folder_path = os.path.join(attachment_folder, received_time)
            check_and_create_folder(time_folder_path)

            # 检查附件
            if item.Attachments.Count > 0:
                for attachment in item.Attachments:
                    attachment_filename = attachment.FileName
                    save_path = os.path.join(time_folder_path, attachment_filename)

                    # 保存附件
                    try:
                        attachment.SaveAsFile(save_path)
                    except Exception as e:
                        logging.error(f"Failed to save the file at {save_path}", exc_info=True)
                    else:
                        new_license_message += f"新许可: {attachment_filename}\n"

        # 显示新许可信息提示框
        if new_license_message:
            show_messagebox("有新许可", new_license_message)
        else:
            show_messagebox("没有新许可", "没有新许可.")

    # 关闭 Outlook 应用程序
    outlook.Quit()

except Exception as e:
    logging.error("An error occurred", exc_info=True)
