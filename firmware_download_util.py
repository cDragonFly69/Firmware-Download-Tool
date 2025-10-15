import os
import shutil
import subprocess
import zipfile
import tkinter as tk
from tkinter import ttk, messagebox
import sys
import pyperclip
import threading
import time
from datetime import datetime


class PrintRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, msg):
        self.text_widget.insert(tk.END, msg)
        self.text_widget.yview(tk.END)

    def flush(self):
        pass


def update_download_progress(progress_label, progress_var, error_var, total_size, copied_size):
    """更新下载进度条的函数"""
    while progress_var[0] < 100 and error_var[0] == 0:
        if total_size > 0:
            progress = min(100, (copied_size[0] / total_size) * 100)
            progress_var[0] = progress
            progress_label['text'] = f"下载进度: {int(progress_var[0])}%"
            progress_label.update_idletasks()
        time.sleep(0.1)

    if error_var[0] == 2:
        progress_label['fg'] = 'red'
        progress_label['text'] = "下载失败，请检查固件文件类型是否准确！"
    elif error_var[0] == 0:
        progress_label['fg'] = 'blue'
        progress_label['text'] = "下载完成，准备解压..."


def update_unzip_progress(progress_label, progress_var, error_var, total_size, extracted_size):
    """更新解压进度条的函数"""
    while progress_var[0] < 100 and error_var[0] == 0:
        if total_size > 0:
            progress = min(100, (extracted_size[0] / total_size) * 100)
            progress_var[0] = progress
            progress_label['text'] = f"解压进度: {int(progress_var[0])}%"
            progress_label.update_idletasks()
        time.sleep(0.1)

    if error_var[0] == 0:
        progress_label['fg'] = 'green'
        progress_label['text'] = "解压完成！请复制地址烧录固件"


def copy_with_progress(src, dst, buffer_size=64 * 1024, callback=None):
    """带进度回调的文件复制函数"""
    with open(src, 'rb') as fsrc:
        with open(dst, 'wb') as fdst:
            while True:
                buf = fsrc.read(buffer_size)
                if not buf:
                    break
                fdst.write(buf)
                if callback:
                    callback(len(buf))


def download_firmware(firmware_code, project_name, log_widget, progress_label, local_path, progress_var, error_var):
    try:
        # 尝试从"定制"路径下载
        firmware_url = os.path.join(r"\\10.11.80.6\oss 固件仓库\定制", f"{firmware_code}.zip")
        if not os.path.exists(firmware_url):
            # 如果"定制"路径不存在，尝试从"标准"路径下载
            firmware_url = os.path.join(r"\\10.11.80.6\oss 固件仓库\标准", f"{firmware_code}.zip")
            if not os.path.exists(firmware_url):
                # 如果两个路径都不存在，抛出错误
                error_var[0] = 2
                raise FileNotFoundError(f"无法找到固件文件: {firmware_code}.zip")

        # 获取文件总大小
        total_size = os.path.getsize(firmware_url)
        copied_size = [0]  # 用列表包裹以便线程间共享

        # 确保本地路径存在
        if not os.path.exists(local_path):
            os.makedirs(local_path)

        # 下载文件（带进度回调）
        destination = os.path.join(local_path, f"{firmware_code}.zip")

        def copy_callback(bytes_copied):
            copied_size[0] += bytes_copied

        # 启动下载进度更新线程
        threading.Thread(target=update_download_progress,
                         args=(progress_label, progress_var, error_var, total_size, copied_size)).start()

        # 执行带进度的文件复制
        copy_with_progress(firmware_url, destination, callback=copy_callback)

        print(f"固件已成功从 {firmware_url} 下载到 {destination}")

        # 解压并删除源文件
        unzip_and_delete(destination, local_path, log_widget, progress_var, error_var, progress_label)
    except FileNotFoundError as e:
        messagebox.showerror("错误", str(e))
    except Exception as e:
        error_var[0] = 1
        messagebox.showerror("错误", f"下载过程中发生错误: {str(e)}")


def unzip_and_delete(zip_file, local_path, log_widget, progress_var, error_var, progress_label):
    try:
        if not os.path.exists(zip_file):
            print(f"文件 {zip_file} 不存在！")
            return

        current_directory = local_path
        print(f"当前目录: {current_directory}")
        print("▓▓▓▓▓▓▓▓▓▓▓▓▓▓")
        print(f"....正在解压....请等待....")

        # 获取ZIP文件总大小
        total_size = 0
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            for file in zip_ref.infolist():
                total_size += file.file_size

        extracted_size = [0]  # 用列表包裹以便线程间共享

        # 启动解压进度更新线程
        threading.Thread(target=update_unzip_progress,
                         args=(progress_label, progress_var, error_var, total_size, extracted_size)).start()

        # 执行解压（带进度更新）
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            for file in zip_ref.infolist():
                zip_ref.extract(file, current_directory)
                extracted_size[0] += file.file_size

        os.remove(zip_file)
        print(f"已删除源文件 {zip_file}")
        print("▓▓▓▓▓▓▓▓▓▓▓▓▓▓")
        print(f"解压完成！\n请点击复制已下载固件地址烧录固件")
        progress_var[0] = 100  # 解压完成后设置进度为100%
        if progress_var[0] == 100:
            progress_label['fg'] = 'green'
            progress_label['text'] = "下载完成，请复制地址烧录！"
    except Exception as e:
        error_var[0] = 1
        messagebox.showerror("错误", str(e))


def copy_directory_path_to_clipboard(project_name_entry, folder_path_entry):
    project_name = project_name_entry.get()
    local_path = fr"D:\待测固件\{project_name}"
    folder_path_entry.delete(0, tk.END)
    folder_path_entry.insert(0, local_path)
    pyperclip.copy(local_path)
    messagebox.showinfo("成功", "目录路径已复制到剪贴板")


def delete_folder(folder_path):
    if os.path.exists(folder_path):
        shutil.rmtree(folder_path)
        messagebox.showinfo("成功", "文件夹已删除")
    else:
        messagebox.showerror("错误", "文件夹不存在")


def open_folder(folder_path):
    try:
        if os.name == 'nt':  # For Windows
            os.startfile(folder_path)
        else:  # For Unix or MacOS
            opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
            subprocess.call([opener, folder_path])
    except Exception as e:
        messagebox.showerror("错误", f"无法打开文件夹: {str(e)}")


def create_folder_from_description(description_entry, base_path):
    description = description_entry.get()
    if not description:
        messagebox.showwarning("输入错误", "请输入内容")
        return

    # 去除所有的【】符号
    description = description.replace("【", "[").replace("】", "]")

    # 解析输入内容，按 '][' 分割
    parts = description.split('][')

    # 如果分割出来的部分小于2，则表示格式错误
    if len(parts) < 2:
        messagebox.showwarning("输入错误", "输入格式不正确")
        return

    # 提取公司名和产品信息
    company_name = parts[0].strip().replace('[', '')
    product_info = parts[1].strip().replace(']', '-')

    # 处理附加信息部分（如果有的话）
    additional_info = '-'.join(parts[2:]).strip() if len(parts) > 2 else ''

    # 拼接产品信息和附加信息（如果有的话）
    if additional_info:
        product_info = f"_{product_info}_{additional_info}"

    # 当前日期
    current_date = datetime.now().strftime('%Y-%m-%d')

    # 创建文件夹名称
    folder_name = f"{current_date}_XXX_{company_name}_{product_info}"
    folder_path = os.path.join(base_path, folder_name)

    def copy_upfm(folder_path):
        if os.path.exists(folder_path):
            source_file = r"D:\任务点\UPFM模板.xlsx"
            target_file = os.path.join(folder_path, f"{folder_name}.xlsx")
            shutil.copy(source_file, target_file)
            print(f"UPFM模板已复制到 {target_file}")

        else:
            messagebox.showerror("错误", "UPFM模板文件夹不存在,请下载模板文件到指定目录：D:\\任务点\\UPFM模板.xlsx")

    # 检查文件夹是否存在，如果不存在则创建
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        # messagebox.showinfo("成功", f"文件夹已创建在：{folder_path}")
        print(f"文件夹已创建好：{folder_path} ")
        copy_upfm(folder_path)
    else:
        messagebox.showerror("错误", "文件夹已存在")


def create_ui():
    root = tk.Tk()
    root.title("固件下载工具")
    root.geometry("520x450")
    root.iconbitmap("favicon.ico")
    # root.wm_iconbitmap("dog.ico")

    left_frame = tk.Frame(root)
    left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    right_frame = tk.Frame(root)
    right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

    tk.Label(left_frame, text="OSS下载码:").pack(pady=5)
    firmware_code_entry = tk.Entry(left_frame, width=40)
    firmware_code_entry.pack(pady=5)

    def on_project_name_entry_click(event):
        """当输入框获得焦点时，如果内容是提示信息，则清除内容"""
        if project_name_entry.get() == '输入格式：公司名或其他文件夹名':
            project_name_entry.delete(0, tk.END)
            project_name_entry.config(fg='black')

    def on_project_name_entry_focusout(event):
        """当输入框失去焦点时，如果内容为空，则恢复提示信息"""
        if project_name_entry.get() == '':
            project_name_entry.insert(0, '输入格式：公司名或其他文件夹名')
            project_name_entry.config(fg='grey')

    # 创建描述输入框并设置提示信息
    tk.Label(left_frame, text="项目名称:").pack(pady=5)
    project_name_entry = tk.Entry(left_frame, width=40)

    project_name_entry.insert(0, '输入格式：公司名或其他文件夹名')
    project_name_entry.config(fg='grey')
    project_name_entry.bind('<FocusIn>', on_project_name_entry_click)
    project_name_entry.bind('<FocusOut>', on_project_name_entry_focusout)
    project_name_entry.pack(pady=5)

    log_widget = tk.Text(right_frame, height=40, width=30)
    log_widget.pack(pady=10)

    sys.stdout = PrintRedirector(log_widget)

    progress_label = tk.Label(left_frame, text="", fg="black")
    progress_label.pack(pady=5)

    def on_download_button_click():
        firmware_code = firmware_code_entry.get()
        project_name = project_name_entry.get()

        # 重置进度条和错误状态
        progress_label['text'] = ""
        progress_label['fg'] = 'black'
        log_widget.delete(1.0, tk.END)  # 清除日志信息

        if not firmware_code or not project_name:
            messagebox.showerror("错误", "请填写完整的 OSS 下载码和项目名称")
            return

        local_path = fr"D:\待测固件\{project_name}"
        progress_var = [0]  # 用列表包裹进度，以便线程间共享
        error_var = [0]  # 用列表包裹错误标志，以便线程间共享

        # 启动下载线程
        threading.Thread(target=download_firmware,
                         args=(firmware_code, project_name, log_widget, progress_label, local_path, progress_var,
                               error_var)).start()

    download_button = tk.Button(left_frame, text="一键下载", command=on_download_button_click)
    download_button.pack(pady=10)

    def on_copy_path_button_click():
        copy_directory_path_to_clipboard(project_name_entry, folder_path_entry)

    copy_path_button = tk.Button(left_frame, text="复制已下载固件地址", command=on_copy_path_button_click)
    copy_path_button.pack(pady=5)

    def on_entry_click(event):
        """当输入框获得焦点时，如果内容是提示信息，则清除内容"""
        if description_entry.get() == '输入格式：【公司名】【项目名】型号版本名':
            description_entry.delete(0, tk.END)
            description_entry.config(fg='black')

    def on_focusout(event):
        """当输入框失去焦点时，如果内容为空，则恢复提示信息"""
        if description_entry.get() == '':
            description_entry.insert(0, '输入格式：【公司名】【项目名】型号版本名')
            description_entry.config(fg='grey')

    # 创建描述输入框并设置提示信息
    description_entry = tk.Entry(left_frame, width=40)
    description_entry.pack(pady=5)
    description_entry.insert(0, '输入格式：【公司名】【项目名】型号版本名')
    description_entry.config(fg='grey')
    description_entry.bind('<FocusIn>', on_entry_click)
    description_entry.bind('<FocusOut>', on_focusout)
    description_entry.pack(pady=5)
    create_folder_button = tk.Button(left_frame, text="新建测试报告文件夹",
                                     command=lambda: create_folder_from_description(description_entry,
                                                                                    "D:\\任务点\\非统一化\\"))
    create_folder_button.pack(pady=5)

    folder_path_entry = tk.Entry(left_frame, width=40)
    folder_path_entry.pack(pady=5)

    def on_folder_path_entry_click(event):
        """当输入框获得焦点时，如果内容是提示信息，则清除内容"""
        if folder_path_entry.get() == 'D:\待测固件\你的项目名称':
            folder_path_entry.delete(0, tk.END)
            folder_path_entry.config(fg='black')

    def on_folder_path_entry_focusout(event):
        """当输入框失去焦点时，如果内容为空，则恢复提示信息"""
        if folder_path_entry.get() == '':
            folder_path_entry.insert(0, 'D:\待测固件\你的项目名称')
            folder_path_entry.config(fg='grey')

    folder_path_entry.insert(0, "D:\待测固件\你的项目名称")
    folder_path_entry.config(fg='grey')
    folder_path_entry.bind('<FocusIn>', on_folder_path_entry_click)
    folder_path_entry.bind('<FocusOut>', on_folder_path_entry_focusout)

    # 使用Frame控件将打开和删除按钮放在同一行
    folder_buttons_frame = tk.Frame(left_frame)
    folder_buttons_frame.pack(pady=5)

    def on_open_folder_button_click():
        folder_path = folder_path_entry.get()
        open_folder(folder_path)

    open_folder_button = tk.Button(folder_buttons_frame, text="打开文件夹", command=on_open_folder_button_click)
    open_folder_button.pack(side=tk.LEFT, padx=5)

    def on_delete_folder_button_click():
        folder_path = folder_path_entry.get()
        delete_folder(folder_path)

    delete_folder_button = tk.Button(folder_buttons_frame, text="删除文件夹", command=on_delete_folder_button_click)
    delete_folder_button.pack(side=tk.LEFT, padx=5)

    exit_button = tk.Button(left_frame, text="退出", command=root.quit)
    exit_button.pack(pady=5)

    root.mainloop()


create_ui()