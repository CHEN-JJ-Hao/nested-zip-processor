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
        self.compressed_files = []
        self.compression_thread = None
        self.progress_callback = None
        self.keep_original_archives = False  # 是否保留原始压缩包
        self.flatten_single_folder = True   # 是否展平单层文件夹

    def _sanitize_path(self, path):
        return os.path.normpath(path)

    def extract_archive(self, file_path, extract_to):
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        safe_base_name = self._sanitize_filename(base_name)
        
        # 检查压缩包内结构，确定目标目录
        target_dir = self._determine_target_directory(file_path, extract_to, safe_base_name)
        
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
                    # 设置解压编码为UTF-8，防止中文乱码
                    for member in zf.infolist():
                        self._check_stop_and_pause()
                        # 修正文件名编码
                        member_filename = self._decode_filename(member.filename)
                        member_filename = os.path.normpath(member_filename)
                        if os.path.isabs(member_filename) or '..' in pathlib.PurePath(member_filename).parts:
                            continue
                        target_path = os.path.join(target_dir, member_filename)
                        if member.is_dir():
                            os.makedirs(target_path, exist_ok=True)
                        else:
                            os.makedirs(os.path.dirname(target_path), exist_ok=True)
                            with zf.open(member) as source, open(target_path, 'wb') as target:
                                shutil.copyfileobj(source, target)
            elif has_rar and file_path.lower().endswith('.rar'):
                with rarfile.RarFile(file_path, 'r') as rf:
                    for member in rf.infolist():
                        self._check_stop_and_pause()
                        member_filename = self._decode_filename(member.filename)
                        member_filename = os.path.normpath(member_filename)
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
                        member_name = self._decode_filename(member.name)
                        member_name = os.path.normpath(member_name)
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

        # 优化解压后的文件夹结构
        self.optimize_extracted_structure(target_dir)
        
        # 处理嵌套压缩包
        if not self._stop.is_set():
            self.extract_nested_archives(target_dir)
            
        # 清理原始压缩包（如果需要）
        if not self.keep_original_archives and not self._stop.is_set():
            self._cleanup_extracted_archives(target_dir, file_path)

    def _decode_filename(self, filename):
        if isinstance(filename, bytes):
            try:
                return filename.decode('utf-8')
            except UnicodeDecodeError:
                try:
                    return filename.decode('gbk')
                except UnicodeDecodeError:
                    return filename.decode('utf-8', errors='replace')
        # zipfile 3.6+ filename 已经是 str，但有些是 cp437错误解码的
        try:
            # 尝试将 cp437 错误解码的 str 转回 bytes 再用 gbk/utf-8 解
            raw = filename.encode('cp437')
            try:
                return raw.decode('utf-8')
            except UnicodeDecodeError:
                return raw.decode('gbk')
        except Exception:
            return filename

    def _determine_target_directory(self, file_path, extract_to, base_name):
        """确定解压目标目录，处理同名文件夹和松散文件情况"""
        try:
            # 不使用临时解压，而是直接检查压缩包内结构
            top_level_items = []
            is_all_files = True
            import tempfile
            temp_dir = tempfile.mkdtemp()
            
            if file_path.lower().endswith('.zip'):
                with zipfile.ZipFile(file_path, 'r') as zf:
                    for name in zf.namelist():
                        parts = name.split('/')
                        if parts and parts[0]:
                            top_level_items.append(parts[0])
            elif file_path.lower().endswith('.7z'):
                with py7zr.SevenZipFile(file_path, mode='r') as zf:
                    zf.extractall(temp_dir)
            elif file_path.lower().endswith(('.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz2')):
                with tarfile.open(file_path, 'r:*') as tf:
                    tf.extractall(temp_dir)
            else:
                raise Exception(f"不支持的压缩格式: {file_path}")
            
            # 检查临时目录中的内容
            contents = os.listdir(temp_dir)
            if len(contents) == 1:
                first_item = os.path.join(temp_dir, contents[0])
                if os.path.isdir(first_item) and os.path.basename(first_item) == base_name:
                    # 情况1: 根目录存在同名文件夹，直接使用base_name作为目标目录
                    target_dir = os.path.join(self._sanitize_path(extract_to), base_name)
                else:
                    # 情况2: 根目录是松散文件或不同名文件夹，创建base_name作为父目录
                    target_dir = os.path.join(self._sanitize_path(extract_to), base_name)
            else:
                # 多个文件或文件夹，创建base_name作为父目录
                target_dir = os.path.join(self._sanitize_path(extract_to), base_name)
        except Exception as e:
            self._show_progress(f"检查压缩包结构时出错: {e}")
            target_dir = os.path.join(self._sanitize_path(extract_to), base_name)
        finally:
            # 清理临时目录
            if 'temp_dir' in locals() and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
                
        return target_dir

    def optimize_extracted_structure(self, target_dir):
        """优化解压后的文件夹结构，减少冗余层级"""
        if not self.flatten_single_folder:
            return
            
        # 获取目标目录下的所有内容
        try:
            contents = os.listdir(target_dir)
        except Exception:
            return
            
        # 如果目录下只有一个子目录，且子目录名与目标目录名相似，则展平
        if len(contents) == 1:
            first_item = os.path.join(target_dir, contents[0])
            if os.path.isdir(first_item):
                subdir_name = os.path.basename(first_item)
                parent_name = os.path.basename(target_dir)
                
                # 检查子目录名是否与父目录名相似
                if self._are_names_similar(subdir_name, parent_name):
                    self._show_progress(f"优化文件夹结构: 展平 {subdir_name}")
                    
                    # 将子目录中的所有内容移动到父目录
                    for item in os.listdir(first_item):
                        src = os.path.join(first_item, item)
                        dst = os.path.join(target_dir, item)
                        # 如果目标已存在，则添加序号
                        if os.path.exists(dst):
                            base, ext = os.path.splitext(item)
                            count = 1
                            while os.path.exists(dst):
                                dst = os.path.join(target_dir, f"{base}_{count}{ext}")
                                count += 1
                        shutil.move(src, dst)
                    
                    # 删除空的子目录
                    os.rmdir(first_item)

    def _are_names_similar(self, name1, name2):
        """判断两个名称是否相似（忽略大小写、扩展名和常见后缀）"""
        name1 = name1.lower()
        name2 = name2.lower()
        
        # 移除常见后缀
        suffixes = ['_files', '-files', '_contents', '-contents', '_extracted', '-extracted']
        for suffix in suffixes:
            if name1.endswith(suffix):
                name1 = name1[:-len(suffix)]
            if name2.endswith(suffix):
                name2 = name2[:-len(suffix)]
                
        return name1 == name2

    def _sanitize_filename(self, filename):
        """清理文件名，移除非法字符"""
        return "".join(c for c in filename if c.isalnum() or c in (' ', '_', '-')).rstrip()

    def extract_nested_archives(self, folder):
        """处理嵌套压缩包"""
        while True:
            archives = self._find_archives(folder)
            if not archives:
                break
                
            # 按路径长度排序，优先处理内层压缩包
            archives.sort(key=lambda x: len(x.split(os.sep)))
            
            for archive in archives:
                self._check_stop_and_pause()
                
                # 获取父目录
                parent_dir = os.path.dirname(archive)
                
                # 从文件名生成子文件夹名
                sub_folder_name = self._sanitize_filename(os.path.splitext(os.path.basename(archive))[0])
                sub_folder = os.path.join(parent_dir, sub_folder_name)
                
                # 确保子文件夹名唯一
                orig_sub_folder = sub_folder
                count = 1
                while os.path.exists(sub_folder):
                    sub_folder = f"{orig_sub_folder}_{count}"
                    count += 1
                    
                os.makedirs(sub_folder, exist_ok=True)
                self._show_progress(f"正在解压嵌套文件: {os.path.basename(archive)}")
                
                try:
                    self._extract_single_archive(archive, sub_folder)
                    
                    # 优化解压后的结构
                    self.optimize_extracted_structure(sub_folder)
                    
                    # 处理子文件夹中的嵌套压缩包
                    if not self._stop.is_set():
                        self.extract_nested_archives(sub_folder)
                        
                    # 如果不需要保留原始压缩包，则删除
                    if not self.keep_original_archives and not self._stop.is_set():
                        self._safe_remove(archive)
                except Exception as e:
                    self._show_progress(f"嵌套文件解压失败: {e}")
                    continue
                    
        self._show_progress("")

    def _find_archives(self, folder):
        """查找文件夹中的所有压缩包"""
        archives = []
        for dirpath, _, filenames in os.walk(folder):
            for filename in filenames:
                if self._is_supported_archive(filename):
                    archives.append(os.path.join(dirpath, filename))
        return archives

    def _get_base_folder(self, path):
        """从压缩包路径获取基本文件夹名"""
        filename = os.path.basename(path)
        for ext in ['.tar.gz', '.tgz', '.tar.bz2', '.tbz2', '.zip', '.rar', '.7z', '.tar']:
            if filename.lower().endswith(ext):
                return os.path.join(os.path.dirname(path), filename[:-len(ext)])
        return os.path.join(os.path.dirname(path), filename)

    def _extract_single_archive(self, archive, target_dir):
        """解压单个压缩包"""
        if archive.lower().endswith('.zip'):
            with zipfile.ZipFile(archive, 'r') as zf:
                try:
                    # 修正文件名编码
                    for member in zf.infolist():
                        member.filename = self._decode_filename(member.filename)
                        if not os.path.isabs(member.filename) and '..' not in pathlib.PurePath(member.filename).parts:
                            zf.extract(member, target_dir)
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
                    # 处理文件名编码问题
                    for member in tf.getmembers():
                        member.name = self._decode_filename(member.name)
                        if not os.path.isabs(member.name) and '..' not in pathlib.PurePath(member.name).parts:
                            tf.extract(member, target_dir)
                except Exception as e:
                    self._show_progress(f"tar解压异常: {e}")

    def _cleanup_extracted_archives(self, target_dir, original_file):
        """清理解压后的压缩包文件"""
        for root, _, files in os.walk(target_dir):
            for f in files:
                file_path = os.path.join(root, f)
                if self._is_supported_archive(f) and file_path != original_file:
                    self._safe_remove(file_path)

    def _is_supported_archive(self, filename):
        """检查是否为支持的压缩格式"""
        return filename.lower().endswith(('.zip', '.rar', '.7z', '.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz2'))

    def _safe_remove(self, path):
        """安全删除文件"""
        try:
            os.remove(path)
        except Exception:
            pass

    def _check_stop_and_pause(self):
        """检查是否需要暂停或停止操作"""
        if self._stop.is_set():
            raise Exception("用户终止了操作")
        self._pause.wait()

    def pause(self):
        """暂停当前操作"""
        self._pause.clear()

    def resume(self):
        """恢复暂停的操作"""
        self._pause.set()

    def stop(self):
        """停止当前操作"""
        self._stop.set()
        self._pause.set()

    def rollback(self):
        """回滚所有操作，删除已解压的内容"""
        # 删除所有解压的目录
        for d in reversed(self.extracted_dirs):
            if os.path.exists(d):
                try:
                    shutil.rmtree(d, ignore_errors=False)
                    self._show_progress(f"已删除: {d}")
                except Exception as e:
                    self._show_progress(f"删除失败: {d}, 错误: {str(e)}")
        
        # 删除所有压缩的文件
        for f in reversed(self.compressed_files):
            if os.path.exists(f):
                try:
                    os.remove(f)
                    self._show_progress(f"已删除: {f}")
                except Exception as e:
                    self._show_progress(f"删除失败: {f}, 错误: {str(e)}")
        
        self.extracted_dirs.clear()
        self.compressed_files.clear()

    def _show_progress(self, msg):
        """显示进度信息"""
        if self.progress_callback:
            self.progress_callback(msg)

    def compress_folder(self, folder_path, archive_path, fmt="zip"):
        """压缩文件夹"""
        self.compressed_files.append(archive_path)
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
        """压缩文件"""
        self.compressed_files.append(archive_path)
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
        """解压单个文件"""
        self.extract_archive(file_path, extract_to)
    
    def extract_folder(self, folder_path, extract_to=None):
        """解压文件夹中的所有压缩包"""
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
                
                # 优化解压后的结构
                self.optimize_extracted_structure(sub_folder)
                
                # 处理嵌套压缩包
                if not self._stop.is_set():
                    self.extract_nested_archives(sub_folder)
                    
                # 如果不需要保留原始压缩包，则删除
                if not self.keep_original_archives and not self._stop.is_set():
                    self._safe_remove(archive)
            except Exception as e:
                self._show_progress(f"解压失败: {str(e)}")
                continue
                
        if self.extracted_dirs:
            open_folder(self.extracted_dirs[0])

extractor = Extractor()
extract_thread = None

def start_extract_file(file_paths, extract_to):
    global extract_thread
    extractor._stop.clear()
    extractor._pause.set()
    extractor.extracted_dirs.clear()
    extractor.compressed_files.clear()
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
    extractor.compressed_files.clear()
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
    root.configure(bg="#f0f0f0")

    # 统一字体设置
    title_font = ("微软雅黑", 22, "bold")
    subtitle_font = ("微软雅黑", 14)
    button_font = ("微软雅黑", 11)
    progress_font = ("微软雅黑", 12)

    title_label = tk.Label(root, text="作者-陈富豪", font=title_font, bg="#f0f0f0", fg="#333")
    title_label.pack(pady=(20, 5))

    subtitle_label = tk.Label(root, text="拖入或选择文件/夹进行解压/缩", font=subtitle_font, bg="#f0f0f0", fg="#666")
    subtitle_label.pack(pady=(0, 15))

    sep1 = tk.Frame(root, bg="#d0d0d0", height=1)
    sep1.pack(fill=tk.X, padx=30, pady=(0, 15))

    drag_frame = tk.Frame(root, bg="#e6f2ff", bd=2, relief="ridge", height=100, width=400)
    drag_frame.pack(pady=(0, 20))
    drag_frame.pack_propagate(False)

    drag_label = tk.Label(
        drag_frame,
        text="拖入压缩包或文件夹自动处理解压(含嵌套式解压)\n支持ZIP、RAR、7Z、TAR等格式",
        font=("微软雅黑", 12, "bold"),
        bg="#e6f2ff",
        fg="#2c5282",
        anchor="center",
        justify="center",
        wraplength=380
    )
    drag_label.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

    sep2 = tk.Frame(root, bg="#d0d0d0", height=1)
    sep2.pack(fill=tk.X, padx=30, pady=(0, 20))

    btn_frame = tk.Frame(root, bg="#f0f0f0")
    btn_frame.pack(side=tk.BOTTOM, pady=10)

    # 统一按钮样式参数
    btn_width = 14
    btn_padx = 8
    btn_pady = 8
    btn_relief = "raised"
    btn_bd = 1
    btn_active_bg = "#e0e0e0"

    # 解压菜单按钮
    extract_menu_btn = tk.Menubutton(
        btn_frame, text="解压",
        bg="#4a86e8", fg="white", font=button_font,
        activebackground=btn_active_bg, activeforeground="#333",
        relief=btn_relief, padx=btn_padx, pady=btn_pady, bd=btn_bd, width=btn_width
    )
    extract_menu_btn.pack(side=tk.LEFT, padx=5)
    extract_menu = Menu(extract_menu_btn, tearoff=0)
    extract_menu.add_command(label="解压文件", command=on_choose_extract_file)
    extract_menu.add_command(label="解压文件夹", command=on_choose_extract_folder)
    extract_menu_btn.configure(menu=extract_menu)

    # 压缩菜单按钮
    compress_menu_btn = tk.Menubutton(
        btn_frame, text="压缩",
        bg="#6aa84f", fg="white", font=button_font,
        activebackground=btn_active_bg, activeforeground="#333",
        relief=btn_relief, padx=btn_padx, pady=btn_pady, bd=btn_bd, width=btn_width
    )
    compress_menu_btn.pack(side=tk.LEFT, padx=5)
    compress_menu = Menu(compress_menu_btn, tearoff=0)
    compress_menu.add_command(label="压缩文件", command=on_compress_file)
    compress_menu.add_command(label="压缩文件夹", command=on_compress_folder)
    compress_menu_btn.configure(menu=compress_menu)

    # 暂停/继续按钮
    btn_pause_resume = tk.Button(
        btn_frame, text="暂停", command=on_pause_resume,
        bg="#f1c232", fg="#333", font=button_font,
        activebackground=btn_active_bg, activeforeground="#333",
        relief=btn_relief, padx=btn_padx, pady=btn_pady, bd=btn_bd, width=btn_width
    )
    btn_pause_resume.pack(side=tk.LEFT, padx=5)

    # 终止按钮
    btn_stop = tk.Button(
        btn_frame, text="终止", command=on_stop,
        bg="#e66465", fg="white", font=button_font,
        activebackground=btn_active_bg, activeforeground="#333",
        relief=btn_relief, padx=btn_padx, pady=btn_pady, bd=btn_bd, width=btn_width
    )
    btn_stop.pack(side=tk.LEFT, padx=5)

    # 进度显示区域
    progress_var = tk.StringVar()
    progress_label = tk.Label(
        root,
        textvariable=progress_var,
        fg="#2b6cb0",
        font=progress_font,
        bg="#f0f0f0",
        wraplength=680,
        justify="center",
        anchor="center"
    )
    progress_label.pack(fill=tk.X, pady=(10, 20))

    def update_progress(msg):
        root.after(0, progress_var.set, msg)
    extractor.progress_callback = update_progress

    # 配置拖放区域
    drag_frame.drop_target_register(DND_FILES)
    drag_frame.dnd_bind('<<Drop>>', on_drop)
    extractor.root = root

    root.mainloop()
    if extract_thread and extract_thread.is_alive():
        extractor.stop()
        extract_thread.join()

import sys