import os
import zipfile
import threading
import shutil
import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
from tkinter import messagebox, filedialog, ttk, Menu
import subprocess
import sys
import re
import py7zr
import tarfile
import pathlib
import io
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
        self.compression_thread = None
        self.progress_callback = None

    def _sanitize_path(self, path):
        return os.path.normpath(path)

    def extract_archive(self, file_path, extract_to):
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        safe_base_name = "".join(c for c in base_name if c.isalnum() or c in (' ', '_', '-')).rstrip()
        target_dir = os.path.join(self._sanitize_path(extract_to), safe_base_name)
        orig_target_dir = target_dir
        count = 1
        while os.path.exists(target_dir):
            target_dir = f"{orig_target_dir}_{count}"
            count += 1
        os.makedirs(target_dir, exist_ok=True)
        self.extracted_dirs.append(target_dir)
        self._show_progress(f"正在解压: {os.path.basename(file_path)}")
        try:
            if file_path.lower().endswith('.zip'):
                with zipfile.ZipFile(file_path, 'r') as zf:
                    for member in zf.infolist():
                        self._check_stop_and_pause()
                        member_filename = os.path.normpath(member.filename)
                        if os.path.isabs(member_filename) or '..' in pathlib.PurePath(member_filename).parts:
                            continue
                        try:
                            zf.extract(member, target_dir)
                        except RuntimeError as e:
                            if 'password required' in str(e).lower():
                                self._show_progress("检测到加密压缩包，暂不支持密码解压，已跳过。")
                                continue
                            else:
                                raise
            elif has_rar and file_path.lower().endswith('.rar'):
                with rarfile.RarFile(file_path, 'r') as rf:
                    for member in rf.infolist():
                        self._check_stop_and_pause()
                        member_filename = os.path.normpath(member.filename)
                        if os.path.isabs(member_filename) or '..' in pathlib.PurePath(member_filename).parts:
                            continue
                        try:
                            rf.extract(member, target_dir)
                        except rarfile.BadRarFile as e:
                            self._show_progress("RAR文件损坏，已跳过。")
                            continue
                        except rarfile.PasswordRequired:
                            self._show_progress("检测到加密RAR包，暂不支持密码解压，已跳过。")
                            continue
            elif file_path.lower().endswith('.7z'):
                with py7zr.SevenZipFile(file_path, mode='r') as zf:
                    self._check_stop_and_pause()
                    try:
                        zf.extractall(target_dir)
                    except py7zr.exceptions.PasswordRequired:
                        self._show_progress("检测到加密7z包，暂不支持密码解压，已跳过。")
            elif file_path.lower().endswith(('.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz2')):
                with tarfile.open(file_path, 'r:*') as tf:
                    for member in tf.getmembers():
                        self._check_stop_and_pause()
                        member_name = os.path.normpath(member.name)
                        if os.path.isabs(member_name) or '..' in pathlib.PurePath(member_name).parts:
                            continue
                        try:
                            tf.extract(member, target_dir)
                        except Exception as e:
                            self._show_progress(f"tar解压异常: {e}")
                            continue
            else:
                raise Exception(f"不支持的压缩格式: {file_path}")
        except Exception as e:
            self._show_progress("")
            raise Exception(f"{file_path} 解压失败: {e}")

        self.extract_nested_archives(target_dir)
        self._cleanup_extracted_archives(target_dir, file_path)

    def extract_nested_archives(self, folder):
        while True:
            archives = self._find_archives(folder)
            if not archives:
                break
            for archive in archives:
                self._check_stop_and_pause()
                sub_folder = self._get_base_folder(archive)
                orig_sub_folder = sub_folder
                count = 1
                while os.path.exists(sub_folder):
                    sub_folder = f"{orig_sub_folder}_{count}"
                    count += 1
                os.makedirs(sub_folder, exist_ok=True)
                self._show_progress(f"正在解压: {os.path.basename(archive)}")
                try:
                    self._extract_single_archive(archive, sub_folder)
                except Exception as e:
                    self._show_progress("")
                    raise Exception(f"{archive} 解压失败: {e}")
                self._safe_remove(archive)
        self._show_progress("")

    def _find_archives(self, folder):
        archives = []
        for dirpath, _, filenames in os.walk(folder):
            for filename in filenames:
                if filename.lower().endswith(('.zip', '.rar', '.7z', '.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz2')):
                    archives.append(os.path.join(dirpath, filename))
        return archives

    def _get_base_folder(self, path):
        filename = os.path.basename(path)
        for ext in ['.tar.gz', '.tgz', '.tar.bz2', '.tbz2', '.zip', '.rar', '.7z', '.tar']:
            if filename.lower().endswith(ext):
                return os.path.join(os.path.dirname(path), filename[:-len(ext)])
        return os.path.join(os.path.dirname(path), filename)

    def _extract_single_archive(self, archive, target_dir):
        if archive.lower().endswith('.zip'):
            with zipfile.ZipFile(archive, 'r') as zf:
                try:
                    zf.extractall(target_dir)
                except RuntimeError as e:
                    if 'password required' in str(e).lower():
                        self._show_progress("检测到加密压缩包，暂不支持密码解压，已跳过。")
        elif has_rar and archive.lower().endswith('.rar'):
            with rarfile.RarFile(archive, 'r') as rf:
                try:
                    rf.extractall(target_dir)
                except rarfile.PasswordRequired:
                    self._show_progress("检测到加密RAR包，暂不支持密码解压，已跳过。")
        elif archive.lower().endswith('.7z'):
            with py7zr.SevenZipFile(archive, mode='r') as zf:
                try:
                    zf.extractall(target_dir)
                except py7zr.exceptions.PasswordRequired:
                    self._show_progress("检测到加密7z包，暂不支持密码解压，已跳过。")
        elif archive.lower().endswith(('.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz2')):
            with tarfile.open(archive, 'r:*') as tf:
                try:
                    tf.extractall(target_dir)
                except Exception as e:
                    self._show_progress(f"tar解压异常: {e}")

    def _cleanup_extracted_archives(self, target_dir, original_file):
        for root, _, files in os.walk(target_dir):
            for f in files:
                file_path = os.path.join(root, f)
                if self._is_supported_archive(f) and file_path != original_file:
                    self._safe_remove(file_path)

    def _is_supported_archive(self, filename):
        return filename.lower().endswith(('.zip', '.rar', '.7z', '.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz2'))

    def _safe_remove(self, path):
        try:
            os.remove(path)
        except Exception:
            pass

    def _check_stop_and_pause(self):
        if self._stop.is_set():
            raise Exception("用户终止了操作")
        self._pause.wait()

    def pause(self):
        self._pause.clear()

    def resume(self):
        self._pause.set()

    def stop(self):
        self._stop.set()
        self._pause.set()

    def rollback(self):
        for d in reversed(self.extracted_dirs):
            if os.path.exists(d):
                shutil.rmtree(d, ignore_errors=True)
        self.extracted_dirs.clear()

    def _show_progress(self, msg):
        if self.progress_callback:
            self.progress_callback(msg)

    def compress_folder(self, folder_path, archive_path, fmt="zip"):
        self._show_progress(f"正在压缩: {os.path.basename(folder_path)}")
        try:
            if fmt == "zip":
                with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for root, dirs, files in os.walk(folder_path):
                        for file in files:
                            self._check_stop_and_pause()
                            abs_path = os.path.join(root, file)
                            rel_path = os.path.relpath(abs_path, folder_path)
                            zf.write(abs_path, rel_path)
            elif fmt == "7z":
                with py7zr.SevenZipFile(archive_path, 'w') as zf:
                    self._check_stop_and_pause()
                    zf.writeall(folder_path, arcname="")
            elif fmt == "tar":
                with tarfile.open(archive_path, "w:gz") as tf:
                    self._check_stop_and_pause()
                    tf.add(folder_path, arcname=os.path.basename(folder_path))
            elif fmt == "rar" and has_rar:
                with rarfile.RarFile(archive_path, 'w') as rf:
                    for root, dirs, files in os.walk(folder_path):
                        for file in files:
                            self._check_stop_and_pause()
                            abs_path = os.path.join(root, file)
                            rel_path = os.path.relpath(abs_path, folder_path)
                            rf.write(abs_path, rel_path)
            else:
                raise Exception("不支持的压缩格式")
            self._show_progress("压缩完成")
        except Exception as e:
            self._show_progress(f"压缩失败: {str(e)}")
            raise

    def compress_file(self, file_path, archive_path, fmt="zip"):
        self._show_progress(f"正在压缩: {os.path.basename(file_path)}")
        try:
            if fmt == "zip":
                with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                    self._check_stop_and_pause()
                    zf.write(file_path, os.path.basename(file_path))
            elif fmt == "7z":
                with py7zr.SevenZipFile(archive_path, 'w') as zf:
                    self._check_stop_and_pause()
                    zf.write(file_path, arcname=os.path.basename(file_path))
            elif fmt == "tar":
                with tarfile.open(archive_path, "w:gz") as tf:
                    self._check_stop_and_pause()
                    tf.add(file_path, arcname=os.path.basename(file_path))
            else:
                raise Exception("不支持的压缩格式")
            self._show_progress("压缩完成")
        except Exception as e:
            self._show_progress(f"压缩失败: {str(e)}")
            raise

    def extract_file(self, file_path, extract_to):
        self.extract_archive(file_path, extract_to)
    
    def extract_folder(self, folder_path, extract_to=None):
        if not extract_to:
            extract_to = folder_path
        archives = self._find_archives(folder_path)
        if not archives:
            raise Exception("所选文件夹中没有找到支持的压缩包")
        for archive in archives:
            self._check_stop_and_pause()
            sub_folder = self._get_base_folder(archive)
            orig_sub_folder = sub_folder
            count = 1
            while os.path.exists(sub_folder):
                sub_folder = f"{orig_sub_folder}_{count}"
                count += 1
            os.makedirs(sub_folder, exist_ok=True)
            self._show_progress(f"正在解压: {os.path.basename(archive)}")
            try:
                self._extract_single_archive(archive, sub_folder)
            except Exception as e:
                self._show_progress(f"解压失败: {str(e)}")
                continue
            self._safe_remove(archive)
        if self.extracted_dirs:
            open_folder(self.extracted_dirs[0])

extractor = Extractor()
extract_thread = None

def start_extract_file(file_paths, extract_to):
    global extract_thread
    extractor._stop.clear()
    extractor._pause.set()
    extractor.extracted_dirs.clear()
    extractor.compression_thread = None

    def run():
        try:
            for file_path in file_paths:
                if extractor._stop.is_set():
                    return
                extractor.extract_file(file_path, extract_to)
            extractor._show_progress("解压完成，正在打开文件夹...")
            if extractor.extracted_dirs:
                open_folder(extractor.extracted_dirs[0])
            messagebox.showinfo("提示", "所有压缩包已解压完成。")
        except Exception as e:
            extractor.rollback()
            messagebox.showerror("错误", str(e))
            extractor._show_progress("")

    extract_thread = threading.Thread(target=run, daemon=True)
    extract_thread.start()

def start_extract_folder(folder_path, extract_to=None):
    global extract_thread
    extractor._stop.clear()
    extractor._pause.set()
    extractor.extracted_dirs.clear()
    extractor.compression_thread = None

    def run():
        try:
            extractor.extract_folder(folder_path, extract_to)
            extractor._show_progress("解压完成，正在打开文件夹...")
            if extractor.extracted_dirs:
                open_folder(extractor.extracted_dirs[0])
            messagebox.showinfo("提示", "文件夹内所有压缩包已解压完成。")
        except Exception as e:
            extractor.rollback()
            messagebox.showerror("错误", str(e))
            extractor._show_progress("")

    extract_thread = threading.Thread(target=run, daemon=True)
    extract_thread.start()

def on_drop(event):
    files = event.data
    file_paths = re.findall(r'\{([^}]+)\}|([^\s]+)', files)
    paths = []
    for group in file_paths:
        f = group[0] or group[1]
        f = os.path.normpath(f)
        if os.path.isfile(f) and _is_supported_archive(f):
            paths.append(f)
        elif os.path.isdir(f):
            paths.append(f)
    if not paths:
        messagebox.showerror("错误", "请拖入压缩包文件或文件夹")
        return
    
    if all(os.path.isfile(p) for p in paths):
        start_extract_file(paths, os.path.dirname(paths[0]))
    else:
        start_extract_folder(paths[0])

def on_choose_extract_file():
    filetypes = [
        ("压缩包", "*.zip *.rar *.7z *.tar *.tar.gz *.tgz *.tar.bz2 *.tbz2"),
        ("所有文件", "*.*")
    ]
    files = filedialog.askopenfilenames(
        title="选择要解压的文件",
        filetypes=filetypes
    )
    if not files:
        return
    
    extract_to = filedialog.askdirectory(
        title="选择解压目标文件夹",
        initialdir=os.path.expanduser("~\\Desktop")
    )
    if not extract_to:
        return
    
    paths = [f for f in files if os.path.isfile(f) and _is_supported_archive(f)]
    if not paths:
        messagebox.showerror("错误", "请选择有效的压缩包文件")
        return
    
    start_extract_file(paths, extract_to)

def on_choose_extract_folder():
    folder_path = filedialog.askdirectory(
        title="选择包含压缩包的文件夹",
        initialdir=os.path.expanduser("~\\Desktop")
    )
    if not folder_path:
        return
    
    if not os.path.isdir(folder_path):
        messagebox.showerror("错误", "请选择有效的文件夹")
        return
    
    extract_to = filedialog.askdirectory(
        title="选择解压目标文件夹",
        initialdir=os.path.dirname(folder_path)
    )
    if not extract_to:
        extract_to = None
    
    start_extract_folder(folder_path, extract_to)

def on_pause_resume():
    if extractor._pause.is_set():
        extractor.pause()
        btn_pause_resume.config(text="继续")
        progress_var.set("已暂停，点击继续恢复操作")
    else:
        extractor.resume()
        btn_pause_resume.config(text="暂停")
        progress_var.set("正在处理...")

def on_stop():
    extractor.stop()
    extractor.rollback()
    messagebox.showinfo("提示", "操作已终止，已删除已解压内容。")
    progress_var.set("")

def open_folder(path):
    if os.path.exists(path):
        if sys.platform.startswith('win'):
            os.startfile(path)
        elif sys.platform.startswith('darwin'):
            subprocess.Popen(['open', path])
        else:
            subprocess.Popen(['xdg-open', path])

def _is_supported_archive(filename):
    return filename.lower().endswith(('.zip', '.rar', '.7z', '.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz2'))

def on_compress_file():
    file_path = filedialog.askopenfilename(title="选择要压缩的文件")
    if not file_path:
        return
    save_compressed_file(file_path, is_file=True)

def on_compress_folder():
    folder_path = filedialog.askdirectory(title="选择要压缩的文件夹")
    if not folder_path:
        return
    save_compressed_file(folder_path, is_file=False)

def save_compressed_file(target_path, is_file):
    filetypes = [
        ("ZIP 压缩包", "*.zip"),
        ("7Z 压缩包", "*.7z"),
        ("TAR.GZ 压缩包", "*.tar.gz"),
        ("RAR 压缩包", "*.rar") if has_rar else ("所有支持的压缩包", "*.zip *.7z *.tar.gz")
    ]
    
    default_name = os.path.basename(target_path)
    if is_file:
        default_name = os.path.splitext(default_name)[0]
    
    archive_path = filedialog.asksaveasfilename(
        title="保存压缩包",
        defaultextension=".zip",
        filetypes=filetypes,
        initialfile=default_name + ".zip"
    )
    
    if not archive_path:
        return
    
    fmt = "zip"
    if archive_path.lower().endswith(".7z"):
        fmt = "7z"
    elif archive_path.lower().endswith(".tar.gz"):
        fmt = "tar"
    elif archive_path.lower().endswith(".rar") and has_rar:
        fmt = "rar"
    
    def compress_in_thread():
        extractor.compression_thread = threading.current_thread()
        try:
            if is_file:
                extractor.compress_file(target_path, archive_path, fmt=fmt)
            else:
                # 确保folder_path存在且是一个目录
                if not os.path.exists(target_path) or not os.path.isdir(target_path):
                    raise Exception(f"无效的文件夹路径: {target_path}")
                
                # 检查压缩格式
                if fmt == "rar" and not has_rar:
                    raise Exception("不支持RAR格式，请安装rarfile库")
                
                extractor.compression_thread = threading.current_thread()
                extractor.compress_folder(target_path, archive_path, fmt=fmt)
            messagebox.showinfo("提示", f"压缩完成：{archive_path}")
        except Exception as e:
            messagebox.showerror("压缩错误", f"压缩失败: {str(e)}")
        finally:
            extractor.compression_thread = None
            progress_var.set("")
    
    threading.Thread(target=compress_in_thread, daemon=True).start()

if __name__ == "__main__":
    try:
        from tkinterdnd2 import TkinterDnD
    except ImportError:
        print("请先安装 tkinterdnd2：pip install tkinterdnd2")
        exit(1)

    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    root = TkinterDnD.Tk()
    root.title("轻享 - 智能解压工具")
    root.geometry("720x420")
    root.configure(bg="#f5f5f5")

    title_label = tk.Label(root, text="作者：陈富豪", font=("楷体", 22, "bold"), bg="#f5f5f5", fg="#222")
    title_label.pack(pady=(18, 0))

    subtitle_label = tk.Label(root, text="拖入或选择解压/缩文件/夹", font=("微软雅黑", 14), bg="#f5f5f5", fg="#666")
    subtitle_label.pack(pady=(2, 8))

    sep1 = tk.Frame(root, bg="#e0e0e0", height=2)
    sep1.pack(fill=tk.X, padx=30, pady=(0, 10))

    drag_frame = tk.Frame(root, bg="#eaf6ff", bd=2, relief="groove", height=100, width=400)
    drag_frame.pack(pady=(10, 18))
    drag_frame.pack_propagate(False)

    drag_label = tk.Label(
        drag_frame,
        text="拖入压缩包或文件夹自动处理",
        font=("微软雅黑", 18, "bold"),
        bg="#eaf6ff",
        fg="#444"
    )
    drag_label.pack(expand=True, fill=tk.BOTH)

    sep2 = tk.Frame(root, bg="#e0e0e0", height=2)
    sep2.pack(fill=tk.X, padx=30, pady=(0, 16))

    btn_frame = tk.Frame(root, bg="#f5f5f5")
    btn_frame.pack(side=tk.BOTTOM, pady=18)

    extract_menu_btn = tk.Menubutton(
        btn_frame, text="解压",
        bg="#ff6b6b", fg="white", font=("微软雅黑", 11, "bold"),
        activebackground="#ff8787", activeforeground="white",
        relief="flat", padx=8, pady=4, bd=1, width=14
    )
    extract_menu_btn.pack(side=tk.LEFT, padx=12, ipadx=8, ipady=2)
    
    extract_menu = Menu(extract_menu_btn, tearoff=0)
    extract_menu.add_command(label="解压文件", command=on_choose_extract_file)
    extract_menu.add_command(label="解压文件夹", command=on_choose_extract_folder)
    extract_menu_btn.configure(menu=extract_menu)

    compress_menu_btn = tk.Menubutton(
        btn_frame, text="压缩",
        bg="#4e9cff", fg="white", font=("微软雅黑", 11, "bold"),
        activebackground="#6eaaff", activeforeground="white",
        relief="flat", padx=8, pady=8, bd=1, width=14
    )
    compress_menu_btn.pack(side=tk.LEFT, padx=12, ipadx=8, ipady=2)
    
    compress_menu = Menu(compress_menu_btn, tearoff=0)
    compress_menu.add_command(label="压缩文件", command=on_compress_file)
    compress_menu.add_command(label="压缩文件夹", command=on_compress_folder)
    compress_menu_btn.configure(menu=compress_menu)

    btn_pause_resume = tk.Button(
        btn_frame, text="暂停", command=on_pause_resume,
        bg="#f5f5f5", fg="#333", font=("微软雅黑", 11),
        activebackground="#e0e0e0", activeforeground="#333",
        relief="solid", bd=1, padx=8, pady=4, width=14
    )
    btn_pause_resume.pack(side=tk.LEFT, padx=12, ipadx=8, ipady=2)

    btn_stop = tk.Button(
        btn_frame, text="终止", command=on_stop,
        bg="#e74c3c", fg="white", font=("微软雅黑", 12, "bold"),
        activebackground="#c0392b", activeforeground="white",
        relief="solid", bd=1, padx=8, pady=4, width=14
    )
    btn_stop.pack(side=tk.LEFT, padx=12, ipadx=8, ipady=2)

    progress_var = tk.StringVar()
    progress_label = tk.Label(
        root,
        textvariable=progress_var,
        fg="#0984e3",
        font=("微软雅黑", 12),
        bg="#f5f5f5",
        wraplength=680,
        justify="center",
        anchor="center"
    )
    progress_label.pack(fill=tk.X, pady=16)

    def update_progress(msg):
        root.after(0, progress_var.set, msg)
    extractor.progress_callback = update_progress

    drag_frame.drop_target_register(DND_FILES)
    drag_frame.dnd_bind('<<Drop>>', on_drop)
    extractor.root = root

    root.mainloop()