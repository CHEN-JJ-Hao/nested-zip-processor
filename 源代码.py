import os
import zipfile
import threading
import shutil
import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
from tkinter import messagebox, filedialog
import subprocess
import sys
import re
import py7zr
import tarfile
from tkinter import ttk

try:
    import rarfile
    has_rar = True
except ImportError:
    has_rar = False

class Extractor:
    def __init__(self):
        self._pause = threading.Event()
        self._pause.set()
        self._stop = threading.Event()
        self.extracted_dirs = []
        self.progress_callback = None  # 用于进度显示

    def extract_archive(self, file_path, extract_to):
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        target_dir = os.path.join(extract_to, base_name)
        os.makedirs(target_dir, exist_ok=True)
        self.extracted_dirs.append(target_dir)
        self._show_progress(f"正在解压: {os.path.basename(file_path)}")
        try:
            if file_path.lower().endswith('.zip'):
                with zipfile.ZipFile(file_path, 'r') as zf:
                    for member in zf.infolist():
                        if self._stop.is_set():
                            raise Exception("用户终止了解压")
                        self._pause.wait()
                        zf.extract(member, target_dir)
            elif has_rar and file_path.lower().endswith('.rar'):
                with rarfile.RarFile(file_path, 'r') as rf:
                    for member in rf.infolist():
                        if self._stop.is_set():
                            raise Exception("用户终止了解压")
                        self._pause.wait()
                        rf.extract(member, target_dir)
            elif file_path.lower().endswith('.7z'):
                with py7zr.SevenZipFile(file_path, mode='r') as zf:
                    if self._stop.is_set():
                        raise Exception("用户终止了解压")
                    self._pause.wait()
                    zf.extractall(target_dir)
            elif file_path.lower().endswith(('.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz2')):
                with tarfile.open(file_path, 'r:*') as tf:
                    for member in tf.getmembers():
                        if self._stop.is_set():
                            raise Exception("用户终止了解压")
                        self._pause.wait()
                        tf.extract(member, target_dir)
            else:
                raise Exception(f"不支持的压缩格式: {file_path}")
        except Exception as e:
            self._show_progress("")
            raise Exception(f"{file_path} 解压失败: {e}")
        # 递归处理嵌套压缩包
        self.extract_nested_archives(target_dir)
        # 删除解压出来的压缩包（不删除原始拖拽的压缩包）
        for root, _, files in os.walk(target_dir):
            for f in files:
                if f.lower().endswith(('.zip', '.rar', '.7z', '.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz2')):
                    try:
                        os.remove(os.path.join(root, f))
                    except Exception:
                        pass
        self._show_progress("")

    def extract_nested_archives(self, folder):
        # 递归查找并解压嵌套压缩包
        while True:
            archives = []
            for dirpath, _, filenames in os.walk(folder):
                for filename in filenames:
                    if filename.lower().endswith(('.zip', '.rar', '.7z', '.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz2')):
                        archives.append(os.path.join(dirpath, filename))
            if not archives:
                break
            for archive in archives:
                if self._stop.is_set():
                    return
                self._pause.wait()
                sub_folder = os.path.splitext(archive)[0]
                os.makedirs(sub_folder, exist_ok=True)
                self._show_progress(f"正在解压: {os.path.basename(archive)}")
                try:
                    if archive.lower().endswith('.zip'):
                        with zipfile.ZipFile(archive, 'r') as zf:
                            for member in zf.infolist():
                                if self._stop.is_set():
                                    raise Exception("用户终止了解压")
                                self._pause.wait()
                                zf.extract(member, sub_folder)
                    elif has_rar and archive.lower().endswith('.rar'):
                        with rarfile.RarFile(archive, 'r') as rf:
                            for member in rf.infolist():
                                if self._stop.is_set():
                                    raise Exception("用户终止了解压")
                                self._pause.wait()
                                rf.extract(member, sub_folder)
                    elif archive.lower().endswith('.7z'):
                        with py7zr.SevenZipFile(archive, mode='r') as zf:
                            if self._stop.is_set():
                                raise Exception("用户终止了解压")
                            self._pause.wait()
                            zf.extractall(sub_folder)
                    elif archive.lower().endswith(('.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz2')):
                        with tarfile.open(archive, 'r:*') as tf:
                            for member in tf.getmembers():
                                if self._stop.is_set():
                                    raise Exception("用户终止了解压")
                                self._pause.wait()
                                tf.extract(member, sub_folder)
                    else:
                        continue
                except Exception as e:
                    self._show_progress("")
                    raise Exception(f"{archive} 解压失败: {e}")
                try:
                    os.remove(archive)
                except Exception:
                    pass
        self._show_progress("")

    def pause(self):
        self._pause.clear()

    def resume(self):
        self._pause.set()

    def stop(self):
        self._stop.set()
        self._pause.set()  # 防止线程阻塞

    def rollback(self):
        # 删除所有已解压的文件夹
        for d in reversed(self.extracted_dirs):
            if os.path.exists(d):
                shutil.rmtree(d)
        self.extracted_dirs.clear()

    def _show_progress(self, msg):
        if self.progress_callback:
            self.progress_callback(msg)

extractor = Extractor()
extract_thread = None

def start_extract(file_paths):
    global extract_thread
    extractor._stop.clear()
    extractor._pause.set()
    extractor.extracted_dirs.clear()
    def run():
        try:
            for file_path in file_paths:
                if extractor._stop.is_set():
                    return
                extract_to = os.path.dirname(file_path)
                extractor.extract_archive(file_path, extract_to)
            extractor._show_progress("解压完成，正在打开文件夹...")
            # 自动打开第一个解压目标文件夹
            open_folder(os.path.splitext(file_paths[0])[0])
            messagebox.showinfo("完成", "所有压缩包已解压完成，窗口即将关闭。")
            extractor.root.quit()
        except Exception as e:
            extractor.rollback()
            messagebox.showerror("解压错误", str(e))
            extractor._show_progress("")
    extract_thread = threading.Thread(target=run)
    extract_thread.start()

def on_drop(event):
    files = event.data
    print("拖拽内容：", files)
    file_paths = re.findall(r'\{([^}]+)\}|([^\s]+)', files)
    paths = []
    for group in file_paths:
        f = group[0] if group[0] else group[1]
        f = os.path.normpath(f)
        print("检测文件路径：", f)
        if os.path.isfile(f) and f.lower().endswith(('.zip', '.rar', '.7z', '.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz2')):
            paths.append(f)
    print("最终识别到的文件：", paths)
    if not paths:
        messagebox.showerror("错误", "请拖入一个或多个压缩包文件")
        return
    start_extract(paths, None)

def on_choose_file():
    filetypes = [
        ("压缩包", "*.zip *.rar *.7z *.tar *.tar.gz *.tgz *.tar.bz2 *.tbz2"),
        ("所有文件", "*.*")
    ]
    files = filedialog.askopenfilenames(
        title="选择一个或多个压缩包文件",
        filetypes=filetypes
    )
    paths = [f for f in files if os.path.isfile(f) and f.lower().endswith(('.zip', '.rar', '.7z', '.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz2'))]
    if not paths:
        messagebox.showerror("错误", "请选择一个或多个压缩包文件")
        return
    start_extract(paths, None)

def on_choose_folder_and_files():
    folder = filedialog.askdirectory(title="选择解压目标文件夹", initialdir=os.path.expanduser("~\\Desktop"))
    if not folder:
        return
    filetypes = [
        ("压缩包", "*.zip *.rar *.7z *.tar *.tar.gz *.tgz *.tar.bz2 *.tbz2"),
        ("所有文件", "*.*")
    ]
    files = filedialog.askopenfilenames(
        title="选择一个或多个压缩包文件",
        filetypes=filetypes
    )
    paths = [f for f in files if os.path.isfile(f) and f.lower().endswith(('.zip', '.rar', '.7z', '.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz2'))]
    if not paths:
        messagebox.showerror("错误", "请选择一个或多个压缩包文件")
        return
    start_extract(paths, folder)

def on_pause():
    extractor.pause()

def on_resume():
    extractor.resume()

def on_stop():
    extractor.stop()
    extractor.rollback()
    messagebox.showinfo("已终止", "解压已终止，已删除已解压内容。")

def open_folder(path):
    if os.path.exists(path):
        if sys.platform.startswith('win'):
            os.startfile(path)
        elif sys.platform.startswith('darwin'):
            subprocess.Popen(['open', path])
        else:
            subprocess.Popen(['xdg-open', path])

if __name__ == "__main__":
    try:
        from tkinterdnd2 import TkinterDnD
    except ImportError:
        print("请先安装 tkinterdnd2：pip install tkinterdnd2")
        exit(1)

    # 设置标准输出为UTF-8
    import sys
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    root = TkinterDnD.Tk()
    root.title("轻享 - 智能解压工具")
    root.geometry("520x390")
    root.configure(bg="#f5f5f5")

    # 顶部大标题
    title_label = tk.Label(root, text="作者：陈富豪", font=("楷体", 22, "bold"), bg="#f5f5f5", fg="#222")
    title_label.pack(pady=(18, 0))

    # 副标题
    subtitle_label = tk.Label(root, text="拖拽或选择自动解压（支持嵌套）", font=("微软雅黑", 14), bg="#f5f5f5", fg="#666")
    subtitle_label.pack(pady=(2, 8))

    # 分割线
    sep1 = tk.Frame(root, bg="#e0e0e0", height=2)
    sep1.pack(fill=tk.X, padx=30, pady=(0, 10))

    # 拖拽大区域
    drag_frame = tk.Frame(root, bg="#eaf6ff", bd=2, relief="groove", height=100, width=400)
    drag_frame.pack(pady=(10, 18))
    drag_frame.pack_propagate(False)  # 保持固定大小

    drag_label = tk.Label(
        drag_frame,
        text="拖入压缩包自动解压",
        font=("微软雅黑", 18, "bold"),
        bg="#eaf6ff",
        fg="#444"
    )
    drag_label.pack(expand=True, fill=tk.BOTH)

    # 分割线
    sep2 = tk.Frame(root, bg="#e0e0e0", height=2)
    sep2.pack(fill=tk.X, padx=30, pady=(0, 16))

    # 按钮区域（选择文件夹、暂停/继续、终止）
    btn_frame = tk.Frame(root, bg="#f5f5f5")
    btn_frame.pack(side=tk.BOTTOM, pady=18)

    # 选择文件夹按钮
    btn_choose = tk.Button(
        btn_frame, text="选择文件夹", command=on_choose_folder_and_files,
        bg="#ff6b6b", fg="white", font=("微软雅黑", 11, "bold"),
        activebackground="#ff8787", activeforeground="white",
        relief="flat", padx=8, pady=4, bd=0, width=14
    )
    btn_choose.pack(side=tk.LEFT, padx=12, ipadx=8, ipady=2)

    # 暂停/继续按钮
    is_paused = tk.BooleanVar(value=False)
    def on_pause_resume():
        if not is_paused.get():
            extractor.pause()
            btn_pause_resume.config(text="继续", bg="#f5f5f5", fg="#333", activebackground="#e0e0e0")
            is_paused.set(True)
            progress_var.set("已暂停，点击继续恢复解压")
        else:
            extractor.resume()
            btn_pause_resume.config(text="暂停", bg="#f5f5f5", fg="#333", activebackground="#e0e0e0")
            is_paused.set(False)
            progress_var.set("正在解压...")

    btn_pause_resume = tk.Button(
        btn_frame, text="暂停", command=on_pause_resume,
        bg="#f5f5f5", fg="#333", font=("微软雅黑", 11),
        activebackground="#e0e0e0", activeforeground="#333",
        relief="solid", bd=1, padx=8, pady=4, width=14
    )
    btn_pause_resume.pack(side=tk.LEFT, padx=12, ipadx=8, ipady=2)

    # 终止按钮
    btn_stop = tk.Button(
        btn_frame, text="终止", command=on_stop,
        bg="#d9d9d9", fg="#333", font=("微软雅黑", 12, "bold"),
        activebackground="#bfbfbf", activeforeground="#333",
        relief="solid", bd=1, padx=8, pady=4, width=14
    )
    btn_stop.pack(side=tk.LEFT, padx=12, ipadx=8, ipady=2)

    # 进度提示
    progress_var = tk.StringVar()
    progress_label = tk.Label(
        root,
        textvariable=progress_var,
        fg="#0984e3",
        font=("微软雅黑", 12),
        bg="#f5f5f5",
        wraplength=480,
        justify="center",
        anchor="center",
        height=2  # 让Label有两行高度
    )
    progress_label.pack(fill=tk.X, pady=16)

    def update_progress(msg):
        progress_var.set(msg)
        root.update_idletasks()
    extractor.progress_callback = update_progress

    # 拖拽支持（绑定到 drag_frame）
    drag_frame.drop_target_register(DND_FILES)
    drag_frame.dnd_bind('<<Drop>>', on_drop)
    extractor.root = root  # 绑定窗口对象

    # 修改start_extract支持自定义目标文件夹
    def start_extract(file_paths, extract_to=None):
        global extract_thread
        extractor._stop.clear()
        extractor._pause.set()
        extractor.extracted_dirs.clear()
        def run():
            try:
                for file_path in file_paths:
                    if extractor._stop.is_set():
                        return
                    # 如果指定了解压目标文件夹，则用它，否则用文件所在目录
                    target_dir = extract_to if extract_to else os.path.dirname(file_path)
                    extractor.extract_archive(file_path, target_dir)
                extractor._show_progress("解压完成，正在打开文件夹...")
                open_folder(os.path.splitext(file_paths[0])[0])
                messagebox.showinfo("完成", "所有压缩包已解压完成。")
            except Exception as e:
                extractor.rollback()
                messagebox.showerror("解压错误", str(e))
                extractor._show_progress("")
        extract_thread = threading.Thread(target=run)
        extract_thread.start()

    root.mainloop()