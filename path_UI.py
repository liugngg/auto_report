from pathlib2 import Path
import tkinter as tk
from tkinter import filedialog

path = Path.cwd()


# 采用图形界面获取文件夹路径
def get_dir():
    app = tk.Tk() # 初始化GUI程序
    app.withdraw() # 仅显示对话框，隐藏主窗口
    print("请选择文件夹：\n")
    foldPath = filedialog.askdirectory(title="请选择文件夹")
    app.destroy()
    if foldPath:
        print(f'选择的文件夹为：{foldPath}')
        return Path(foldPath)
    return None

# 采用图形界面获取文件路径
def get_file(*types, title='请选择文件：', start_dir='D:'):
    app = tk.Tk()  # 初始化GUI程序
    app.withdraw()  # 仅显示对话框，隐藏主窗口
    print(f"{title}\n")
    # filePath = filedialog.askopenfilename(title='请选择文件:', filetypes=[('TXT', '*.txt'), ('All Files', '*')], initialdir=D:\TestRecord01)
    file = filedialog.askopenfilename(title=title, filetypes=types, initialdir=start_dir)
    app.destroy()
    if file:
        print(f'选择的文件为：{file}')
        return Path(file)
    return None


if __name__ == '__main__':
    filetyps = [('TXT', '*.txt'), ('All Files', '*')]
    file = get_file(*filetyps)

    if not file:
        print("您没有选择有效的文件，已退出！")
        exit(-1)
    print(f'你选择的文件为 “{file}”。')


