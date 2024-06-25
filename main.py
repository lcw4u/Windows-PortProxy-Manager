import ctypes
import tkinter as tk
from tkinter import messagebox
import subprocess
import sys
import os
import functools
import re


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def run_as_admin():
    script = os.path.abspath(sys.argv[0])
    params = ' '.join([script] + sys.argv[1:])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)


def run_netsh_command(command):
    try:
        msg = subprocess.check_output(command, shell=True, stderr=subprocess.STDOUT)
        result = True
    except subprocess.CalledProcessError as e:
        msg = e.output
        result = False

    try:
        return result, msg.decode('utf-8')
    except UnicodeDecodeError:
        return result, msg.decode('gbk')


def show_port_forwarding():
    result, msg = run_netsh_command("netsh interface portproxy show all")
    text.delete(1.0, tk.END)
    text.insert(tk.END, msg)


def add_port_forwarding(listen_port_entry, forward_entry):
    try:
        listen_port = int(listen_port_entry.get())
    except ValueError:
        messagebox.showinfo("Warning", "请输入正确的端口号")
        return

    forward = forward_entry.get()
    pattern = re.compile(r'^(?P<host>[\w.-]+)[:：](?P<port>\d+)$')
    output = pattern.match(forward)
    if not output:
        messagebox.showinfo("Warning", "请输入正确的转发地址")
        return

    connect_address = output.group("host")
    connect_port = output.group("port")

    command = f"netsh interface portproxy add v4tov4 listenport={listen_port} listenaddress=0.0.0.0 connectport={connect_port} connectaddress={connect_address}"
    result, msg = run_netsh_command(command)
    if result:
        messagebox.showinfo("Result", "添加成功")
    else:
        messagebox.showinfo("Error", f"添加失败：{msg}")
    show_port_forwarding()


def delete_port_forwarding(port_entry):
    input_str = port_entry.get()
    pattern = re.compile(r'(\d+)')
    ports = pattern.findall(input_str)
    if len(ports) == 0:
        messagebox.showinfo("Warning", "请输入正确的端口号")
        return

    txt = []
    for port in ports:
        command = f"netsh interface portproxy delete v4tov4 listenport={port} listenaddress=0.0.0.0"
        result, msg = run_netsh_command(command)
        if result:
            txt.append(f"删除端口{port}成功")
        else:
            txt.append(f"删除端口{port}失败：{msg}")

    messagebox.showinfo("Result", "\n".join(txt))
    show_port_forwarding()


def center_window(window):
    window.update_idletasks()  # 更新 "requested size" 信息
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')


# Create the main window
root = tk.Tk()
root.title("Port Forwarding Manager")


def new_input(parent, row, title, comment=None):
    tk.Label(parent, text=title).grid(row=row, pady=5)
    entry = tk.Entry(parent)
    entry.grid(row=row, column=1, padx=3)
    if comment:
        tk.Label(parent, text=f"({comment})").grid(row=row, column=2)
    return entry


# 新增端口对话框
def open_add_window():
    new_window = tk.Toplevel(master=root)
    new_window.title("添加端口")
    listen_port_entry = new_input(new_window, 0, "监听端口")
    forward_entry = new_input(new_window, 1, "转发到", "格式 129.168.16.90:8080")

    submit_button = tk.Button(new_window, text="添加",
                              command=functools.partial(add_port_forwarding, listen_port_entry, forward_entry))
    submit_button.grid(row=2, column=1, pady=10)

    center_window(new_window)


# 删除端口对话框
def open_del_window():
    new_window = tk.Toplevel(master=root)
    new_window.title("删除端口")
    port_entry = new_input(new_window, 0, "端口号", "多个端口使用,隔开")

    submit_button = tk.Button(new_window, text="删除",
                              command=functools.partial(delete_port_forwarding, port_entry))
    submit_button.grid(row=1, column=1, pady=10)

    center_window(new_window)


if not is_admin():
    run_as_admin()
    sys.exit()

# Create a text widget to display the port forwardings
text = tk.Text(root, height=20, width=80)
text.pack()

# Create buttons
add_button = tk.Button(root, text="添加新的端口转发", command=open_add_window)
add_button.pack()

delete_button = tk.Button(root, text="删除端口", command=open_del_window)
delete_button.pack()

# Run the main loop
show_port_forwarding()
center_window(root)
root.mainloop()
