import os
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import threading
import time
from pathlib import Path
import re

class WPSCacheExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("WPS缓存提取工具")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 设置中文字体支持
        self.style = ttk.Style()
        self.style.configure("Treeview.Heading", font=("SimHei", 10, "bold"))
        self.style.configure("Treeview", font=("SimHei", 10))
        
        # 缓存路径
        self.default_cache_path = self.get_default_cache_path()
        self.target_path = tk.StringVar(value=str(Path.home() / "WPS缓存备份"))
        self.all_files = []  # 存储所有文件信息，用于搜索
        
        # 创建UI
        self.create_widgets()
        
        # 加载缓存文件列表
        self.load_cache_files()
    
    def get_default_cache_path(self):
        """获取默认的WPS缓存路径"""
        try:
            user_profile = os.environ["USERPROFILE"]
            return os.path.join(user_profile, "AppData", "Roaming", "kingsoft", "office6", "backup")
        except Exception as e:
            messagebox.showerror("错误", f"无法获取默认缓存路径: {str(e)}")
            return ""
    
    def create_widgets(self):
        """创建GUI组件"""
        # 源路径框架
        source_frame = ttk.LabelFrame(self.root, text="WPS缓存路径", padding="10")
        source_frame.pack(fill="x", padx=10, pady=5)
        
        # 盘符选择按钮
        drives_frame = ttk.Frame(source_frame)
        drives_frame.pack(side="left", padx=(0, 5))
        
        # 获取可用盘符
        drives = self.get_available_drives()
        for drive in drives[:6]:  # 最多显示6个盘符
            ttk.Button(drives_frame, text=drive, 
                      command=lambda d=drive: self.set_path_to_drive(d)).pack(side="left", padx=2)
        
        # 路径输入和按钮
        path_frame = ttk.Frame(source_frame)
        path_frame.pack(side="left", fill="x", expand=True)
        
        self.source_path_var = tk.StringVar(value=self.default_cache_path)
        ttk.Entry(path_frame, textvariable=self.source_path_var, width=70).pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        button_frame = ttk.Frame(path_frame)
        button_frame.pack(side="left")
        
        ttk.Button(button_frame, text="浏览", command=self.browse_source).pack(side="left", padx=2)
        ttk.Button(button_frame, text="搜索目录", command=self.search_cache_directory).pack(side="left", padx=2)
        
        # 文件搜索框
        search_frame = ttk.LabelFrame(self.root, text="文件搜索", padding="10")
        search_frame.pack(fill="x", padx=10, pady=5)
        
        self.search_keyword = tk.StringVar()
        # 使用占位文本的替代实现方式
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_keyword, width=50)
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.search_entry.insert(0, "输入文件名关键字搜索")
        self.search_entry.bind("<FocusIn>", self.clear_placeholder)
        self.search_entry.bind("<FocusOut>", self.restore_placeholder)
        
        ttk.Button(search_frame, text="搜索文件", command=self.search_files).pack(side="left", padx=2)
        ttk.Button(search_frame, text="显示全部", command=self.load_cache_files).pack(side="left", padx=2)
        
        # 目标路径框架
        target_frame = ttk.LabelFrame(self.root, text="目标保存路径", padding="10")
        target_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Entry(target_frame, textvariable=self.target_path, width=70).pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(target_frame, text="浏览", command=self.browse_target).pack(side="right")
        
        # 文件列表框架
        files_frame = ttk.LabelFrame(self.root, text="缓存文件列表", padding="10")
        files_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 滚动条
        scrollbar_y = ttk.Scrollbar(files_frame)
        scrollbar_y.pack(side="right", fill="y")
        
        scrollbar_x = ttk.Scrollbar(files_frame, orient="horizontal")
        scrollbar_x.pack(side="bottom", fill="x")
        
        # 列表视图
        self.file_tree = ttk.Treeview(files_frame, columns=("name", "size", "modified"), show="headings", 
                                      yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        self.file_tree.heading("name", text="文件名")
        self.file_tree.heading("size", text="大小 (KB)")
        self.file_tree.heading("modified", text="修改时间")
        
        self.file_tree.column("name", width=350)
        self.file_tree.column("size", width=100, anchor="center")
        self.file_tree.column("modified", width=200, anchor="center")
        
        self.file_tree.pack(fill="both", expand=True)
        scrollbar_y.config(command=self.file_tree.yview)
        scrollbar_x.config(command=self.file_tree.xview)
        
        # 操作按钮框架
        button_frame = ttk.Frame(self.root, padding="10")
        button_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(button_frame, text="刷新列表", command=self.load_cache_files).pack(side="left", padx=5)
        ttk.Button(button_frame, text="复制选中", command=self.copy_selected_files).pack(side="right", padx=5)
        ttk.Button(button_frame, text="全部复制", command=self.copy_all_files).pack(side="right", padx=5)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", padx=10, pady=5)
        
        # 状态标签
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(self.root, textvariable=self.status_var).pack(fill="x", padx=10, pady=5)
    
    def clear_placeholder(self, event):
        """清除占位文本"""
        if self.search_entry.get() == "输入文件名关键字搜索":
            self.search_entry.delete(0, tk.END)
    
    def restore_placeholder(self, event):
        """恢复占位文本（如果输入框为空）"""
        if not self.search_entry.get():
            self.search_entry.insert(0, "输入文件名关键字搜索")
    
    def get_available_drives(self):
        """获取系统中可用的盘符"""
        drives = []
        if os.name == 'nt':  # Windows系统
            for drive in range(65, 91):  # A-Z
                drive_letter = chr(drive) + ":"
                if os.path.exists(drive_letter):
                    drives.append(drive_letter)
        return drives
    
    def set_path_to_drive(self, drive):
        """将路径设置为指定盘符"""
        self.source_path_var.set(drive + "\\")
        self.load_cache_files()
    
    def browse_source(self):
        """浏览选择源路径"""
        path = filedialog.askdirectory(title="选择WPS缓存目录")
        if path:
            self.source_path_var.set(path)
            self.load_cache_files()
    
    def browse_target(self):
        """浏览选择目标路径"""
        path = filedialog.askdirectory(title="选择保存目录")
        if path:
            self.target_path.set(path)
    
    def search_cache_directory(self):
        """搜索WPS缓存目录"""
        # 询问用户要搜索的盘符
        drive = simpledialog.askstring("选择盘符", "请输入要搜索的盘符 (例如 C:):", initialvalue="C:")
        if not drive:
            return
        
        if len(drive) == 1:  # 如果只输入了字母，自动添加冒号
            drive += ":"
        
        if not os.path.exists(drive):
            messagebox.showerror("错误", f"盘符 {drive} 不存在")
            return
        
        self.status_var.set(f"正在 {drive} 搜索WPS缓存目录，请稍候...")
        self.root.update()
        
        # 在后台线程中执行搜索，避免UI冻结
        threading.Thread(target=self._perform_directory_search, args=(drive,), daemon=True).start()
    
    def _perform_directory_search(self, drive):
        """执行目录搜索的后台任务"""
        search_pattern = os.path.join("*", "AppData", "Roaming", "kingsoft", "office6", "backup")
        found_paths = []
        
        try:
            # 使用glob搜索可能的路径
            for path in Path(drive).glob(search_pattern):
                if path.is_dir():
                    found_paths.append(str(path))
            
            # 如果找到多个路径，让用户选择
            if len(found_paths) > 0:
                self.root.after(0, lambda: self._show_search_results(found_paths))
            else:
                self.root.after(0, lambda: messagebox.showinfo("搜索结果", "未找到WPS缓存目录"))
        
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"搜索过程中出错: {str(e)}"))
        
        self.root.after(0, lambda: self.status_var.set("就绪"))
    
    def _show_search_results(self, paths):
        """显示搜索结果供用户选择"""
        if len(paths) == 1:
            # 只有一个结果，直接使用
            self.source_path_var.set(paths[0])
            self.load_cache_files()
            messagebox.showinfo("搜索结果", f"找到WPS缓存目录:\n{paths[0]}")
        else:
            # 多个结果，让用户选择
            result_window = tk.Toplevel(self.root)
            result_window.title("选择缓存目录")
            result_window.geometry("500x300")
            result_window.transient(self.root)
            result_window.grab_set()
            
            ttk.Label(result_window, text="找到多个可能的缓存目录，请选择:").pack(padx=10, pady=10)
            
            listbox = tk.Listbox(result_window, font=("SimHei", 10), width=60, height=10)
            listbox.pack(padx=10, pady=5, fill="both", expand=True)
            
            for path in paths:
                listbox.insert("end", path)
            
            # 自动选中默认路径如果存在
            try:
                index = paths.index(self.default_cache_path)
                listbox.selection_set(index)
                listbox.see(index)
            except ValueError:
                pass
            
            def on_select():
                if listbox.curselection():
                    selected = paths[listbox.curselection()[0]]
                    self.source_path_var.set(selected)
                    self.load_cache_files()
                    result_window.destroy()
            
            ttk.Button(result_window, text="确定", command=on_select).pack(pady=10)
    
    def search_files(self):
        """搜索文件"""
        # 获取搜索关键字，排除占位文本的情况
        keyword = self.search_keyword.get().strip().lower()
        if not keyword or keyword == "输入文件名关键字搜索".lower():
            messagebox.showinfo("提示", "请输入搜索关键字")
            return
        
        # 清空现有列表
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        
        # 过滤文件
        matched_files = [f for f in self.all_files if keyword in f[0].lower()]
        
        if not matched_files:
            self.status_var.set(f"没有找到包含 '{keyword}' 的文件")
            return
        
        # 添加匹配的文件到列表
        for file in matched_files:
            self.file_tree.insert("", "end", values=file)
        
        self.status_var.set(f"找到 {len(matched_files)} 个匹配文件")
    
    def get_file_size(self, path):
        """获取文件大小（KB）"""
        try:
            size = os.path.getsize(path)
            return round(size / 1024, 2)
        except:
            return "N/A"
    
    def get_file_modified_time(self, path):
        """获取文件修改时间"""
        try:
            mtime = os.path.getmtime(path)
            return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mtime))
        except:
            return "N/A"
    
    def load_cache_files(self):
        """加载缓存文件列表"""
        # 清空现有列表
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        
        self.all_files = []  # 重置所有文件列表
        cache_path = self.source_path_var.get()
        
        if not cache_path or not os.path.exists(cache_path):
            self.status_var.set(f"错误: 缓存路径不存在 - {cache_path}")
            return
        
        try:
            # 获取所有文件
            files = [f for f in os.listdir(cache_path) if os.path.isfile(os.path.join(cache_path, f))]
            
            if not files:
                self.status_var.set("缓存目录中没有文件")
                return
            
            # 添加到列表和全局文件列表
            for file in files:
                file_path = os.path.join(cache_path, file)
                size = self.get_file_size(file_path)
                modified = self.get_file_modified_time(file_path)
                self.file_tree.insert("", "end", values=(file, size, modified))
                self.all_files.append((file, size, modified))
            
            self.status_var.set(f"已加载 {len(files)} 个文件")
        except Exception as e:
            self.status_var.set(f"加载文件时出错: {str(e)}")
            messagebox.showerror("错误", f"加载文件时出错: {str(e)}")
    
    def copy_file(self, source, target):
        """复制单个文件"""
        try:
            # 确保目标目录存在
            os.makedirs(os.path.dirname(target), exist_ok=True)
            shutil.copy2(source, target)  # 保留元数据
            return True
        except Exception as e:
            print(f"复制文件 {os.path.basename(source)} 失败: {str(e)}")
            return False
    
    def copy_all_files(self):
        """复制所有文件（在后台线程中执行）"""
        source_path = self.source_path_var.get()
        target_path = self.target_path.get()
        
        if not source_path or not os.path.exists(source_path):
            messagebox.showerror("错误", f"源路径不存在: {source_path}")
            return
        
        if not target_path:
            messagebox.showerror("错误", "请指定目标路径")
            return
        
        # 获取所有文件
        try:
            files = [f for f in os.listdir(source_path) if os.path.isfile(os.path.join(source_path, f))]
            if not files:
                messagebox.showinfo("提示", "没有文件可复制")
                return
        except Exception as e:
            messagebox.showerror("错误", f"获取文件列表失败: {str(e)}")
            return
        
        # 在后台线程中执行复制操作，避免UI冻结
        threading.Thread(target=self._copy_files_in_background, args=(source_path, target_path, files), daemon=True).start()
    
    def copy_selected_files(self):
        """复制选中的文件"""
        source_path = self.source_path_var.get()
        target_path = self.target_path.get()
        
        if not source_path or not os.path.exists(source_path):
            messagebox.showerror("错误", f"源路径不存在: {source_path}")
            return
        
        if not target_path:
            messagebox.showerror("错误", "请指定目标路径")
            return
        
        # 获取选中的文件
        selected_items = self.file_tree.selection()
        if not selected_items:
            messagebox.showinfo("提示", "请先选择要复制的文件")
            return
        
        # 提取文件名
        files = []
        for item in selected_items:
            file_name = self.file_tree.item(item, "values")[0]
            files.append(file_name)
        
        # 在后台线程中执行复制操作
        threading.Thread(target=self._copy_files_in_background, args=(source_path, target_path, files), daemon=True).start()
    
    def _copy_files_in_background(self, source_path, target_path, files):
        """后台复制文件"""
        total = len(files)
        success = 0
        failed = 0
        
        for i, file in enumerate(files):
            # 更新进度
            progress = (i + 1) / total * 100
            self.progress_var.set(progress)
            self.status_var.set(f"正在复制: {file} ({i+1}/{total})")
            
            # 复制文件
            source_file = os.path.join(source_path, file)
            target_file = os.path.join(target_path, file)
            
            if self.copy_file(source_file, target_file):
                success += 1
            else:
                failed += 1
            
            # 小延迟，让UI有机会更新
            time.sleep(0.01)
        
        # 完成后更新状态
        self.progress_var.set(100)
        self.status_var.set(f"复制完成 - 成功: {success}, 失败: {failed}")
        messagebox.showinfo("完成", f"复制完成\n成功: {success} 个文件\n失败: {failed} 个文件\n保存路径: {target_path}")

if __name__ == "__main__":
    root = tk.Tk()
    # 解决中文显示问题
    root.option_add("*Font", "SimHei 10")
    app = WPSCacheExtractor(root)
    root.mainloop()
    