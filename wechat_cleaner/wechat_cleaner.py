import os
import sys
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ctypes
import subprocess
import platform
import threading
import re

class WeChatCleaner:
    def __init__(self, root):
        self.root = root
        self.root.title("微信文件夹清理工具")
        self.root.geometry("850x700")
        self.root.resizable(False, False)
        
        # 设置窗口图标
        self.set_window_icon()
        
        # 确保中文显示正常
        self.setup_font()
        
        # 窗口居中
        self.center_window()
        
        # 微信文件夹路径
        self.default_paths = self.get_default_paths()
        self.scanned_paths = []  # 存储深度扫描到的路径
        self.all_paths = self.default_paths.copy()
        
        self.folder_states = {path: tk.BooleanVar(value=True) for path in self.all_paths}
        self.folder_sizes = {}
        self.scanning = False  # 扫描状态标记
        self.stop_scan = False  # 停止扫描标记
        self.filter_type = tk.StringVar(value="全部")  # 筛选类型
        
        # 创建UI
        self.create_widgets()
    
    def set_window_icon(self):
        """设置窗口图标"""
        try:
            # 尝试加载图标文件
            if getattr(sys, 'frozen', False):
                # 当程序被打包为EXE时
                base_dir = sys._MEIPASS
            else:
                # 开发环境
                base_dir = os.path.dirname(os.path.abspath(__file__))
            
            icon_path = os.path.join(base_dir, "wechatclean.ico")
            
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
            else:
                print(f"图标文件未找到: {icon_path}")
        except Exception as e:
            print(f"设置图标时出错: {e}")
    
    def setup_font(self):
        """设置字体以支持中文显示"""
        if platform.system() == "Windows":
            default_font = ("Microsoft YaHei", 10)
            self.root.option_add("*Font", default_font)
    
    def center_window(self):
        """使窗口在屏幕中央显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def get_default_paths(self):
        """获取默认的微信文件夹路径"""
        paths = []
        user_profile = os.environ.get('USERPROFILE', '')
        
        # 默认路径
        path1 = os.path.join(user_profile, 'xwechat_files')
        paths.append(path1)
        
        path2 = os.path.join(user_profile, 'Documents', 'WeChat Files')
        paths.append(path2)
        
        path3 = os.path.join(user_profile, 'Documents', 'xwechat_files')
        paths.append(path3)
        
        return paths
    
    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(main_frame, text="微信文件夹清理工具", font=("Microsoft YaHei", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # 筛选区域
        filter_frame = ttk.Frame(main_frame)
        filter_frame.pack(fill=tk.X, pady=(0, 10), anchor=tk.W)
        
        ttk.Label(filter_frame, text="筛选:").pack(side=tk.LEFT, padx=(0, 10))
        
        # 下拉筛选框
        filter_combobox = ttk.Combobox(
            filter_frame, 
            textvariable=self.filter_type,
            values=["全部", "默认路径", "扫描路径"],
            state="readonly",
            width=12
        )
        filter_combobox.pack(side=tk.LEFT)
        filter_combobox.bind("<<ComboboxSelected>>", lambda e: self.update_folder_list())
        
        # 文件夹列表区域
        list_frame = ttk.LabelFrame(main_frame, text="微信相关文件夹列表", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # 创建表格
        columns = ("select", "path", "size", "type")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=8)
        
        # 设置列标题和宽度
        self.tree.heading("select", text="选择删除")
        self.tree.heading("path", text="文件夹路径")
        self.tree.heading("size", text="文件夹大小")
        self.tree.heading("type", text="类型")
        
        self.tree.column("select", width=80, anchor=tk.CENTER)
        self.tree.column("path", width=550, anchor=tk.W)
        self.tree.column("size", width=100, anchor=tk.CENTER)
        self.tree.column("type", width=80, anchor=tk.CENTER)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 扫描按钮
        self.scan_btn = ttk.Button(button_frame, text="扫描文件夹大小", command=self.scan_folders)
        self.scan_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 深度扫描按钮
        self.deep_scan_btn = ttk.Button(button_frame, text="深度扫描所有驱动器", command=self.start_deep_scan)
        self.deep_scan_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 停止扫描按钮
        self.stop_btn = ttk.Button(button_frame, text="停止扫描", command=self.stop_scanning, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 选择按钮
        select_all_btn = ttk.Button(button_frame, text="全选", command=lambda: self.set_all_selection(True))
        select_all_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        deselect_all_btn = ttk.Button(button_frame, text="取消全选", command=lambda: self.set_all_selection(False))
        deselect_all_btn.pack(side=tk.LEFT)
        
        # 删除按钮
        delete_btn = ttk.Button(main_frame, text="删除选中的文件夹内容", command=self.delete_selected, 
                              style="Danger.TButton")
        delete_btn.pack(fill=tk.X, pady=(0, 20))
        
        # 风险提示区域
        risk_frame = ttk.LabelFrame(main_frame, text="⚠️ 重要提示", padding=10)
        risk_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        risk_text = """
        1. 删除操作不可逆，可能导致微信聊天记录、图片、视频等数据永久丢失！
        2. 请确保已备份重要数据，并关闭微信客户端后再进行操作。
        3. 深度扫描限制为4层文件夹，扫描过程中可点击"停止扫描"按钮终止操作。
        4. 可使用上方下拉框筛选默认路径和扫描到的路径。
        """
        risk_label = ttk.Label(risk_frame, text=risk_text, justify=tk.LEFT, wraplength=750)
        risk_label.pack(anchor=tk.W)
        
        # 联系方式区域
        contact_frame = ttk.Frame(main_frame)
        contact_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        # QQ联系方式
        qq_label = ttk.Label(contact_frame, text="联系方式：QQ 88179096")
        qq_label.pack(side=tk.LEFT, padx=(0, 20))
        
        # 网址链接
        url_label = ttk.Label(contact_frame, text="http://www.itvip.com.cn", 
                            foreground="blue", cursor="hand2")
        url_label.pack(side=tk.LEFT)
        url_label.bind("<Button-1>", self.open_url)
        
        # 初始化样式
        self.init_styles()
        
        # 初始加载文件夹列表（未扫描大小）
        self.update_folder_list()
    
    def init_styles(self):
        """初始化自定义样式"""
        style = ttk.Style()
        style.configure("Danger.TButton", foreground="red", font=("Microsoft YaHei", 10, "bold"))
        style.configure("Stop.TButton", foreground="red")
    
    def update_folder_list(self):
        """更新文件夹列表显示，支持筛选"""
        # 清空现有项
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 获取筛选类型
        filter_val = self.filter_type.get()
        
        # 添加文件夹到列表
        for path in self.all_paths:
            # 根据筛选条件决定是否显示
            if filter_val == "默认路径" and path not in self.default_paths:
                continue
            if filter_val == "扫描路径" and path in self.default_paths:
                continue
            
            exists = os.path.exists(path)
            size_str = self.folder_sizes.get(path, "未扫描") if exists else "不存在"
            selected = "✓" if self.folder_states[path].get() else ""
            
            # 确定文件夹类型
            if path in self.default_paths:
                path_type = "默认"
            else:
                path_type = "扫描"
            
            self.tree.insert("", tk.END, values=(selected, path, size_str, path_type))
        
        # 绑定双击事件用于切换选择状态
        self.tree.bind(f"<Double-1>", self.toggle_selection)
    
    def toggle_selection(self, event):
        """切换文件夹选择状态"""
        if self.scanning:  # 扫描过程中不允许更改选择
            return
            
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            item = self.tree.identify_row(event.y)
            col = int(self.tree.identify_column(event.x).replace("#", "")) - 1
            
            if col == 0:  # 只处理选择列
                path = self.tree.item(item, "values")[1]
                current_state = self.folder_states[path].get()
                self.folder_states[path].set(not current_state)
                self.update_folder_list()
    
    def set_all_selection(self, state):
        """设置所有文件夹的选择状态"""
        if self.scanning:  # 扫描过程中不允许更改选择
            return
            
        # 根据当前筛选状态选择相应的项目
        filter_val = self.filter_type.get()
        
        for path in self.all_paths:
            # 只选择当前筛选条件下可见的项目
            if (filter_val == "默认路径" and path not in self.default_paths) or \
               (filter_val == "扫描路径" and path in self.default_paths):
                continue
                
            self.folder_states[path].set(state)
        
        self.update_folder_list()
    
    def get_folder_size(self, folder):
        """计算文件夹大小，支持中途停止"""
        total_size = 0
        try:
            for dirpath, _, filenames in os.walk(folder):
                # 检查是否需要停止
                if self.stop_scan:
                    return -1
                    
                for f in filenames:
                    fp = os.path.join(dirpath, f)
                    if os.path.exists(fp):
                        total_size += os.path.getsize(fp)
                        
                # 再次检查是否需要停止
                if self.stop_scan:
                    return -1
                    
        except Exception as e:
            print(f"计算大小出错: {e}")
        return total_size
    
    def format_size(self, size_bytes):
        """格式化文件大小为易读格式"""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.2f} KB"
        elif size_bytes < 1024 * 1024 * 1024:
            return f"{size_bytes / (1024 * 1024):.2f} MB"
        else:
            return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"
    
    def scan_folders(self):
        """扫描文件夹大小，支持中途停止"""
        if self.scanning:
            messagebox.showinfo("提示", "正在进行扫描操作，请等待完成")
            return
            
        self.scanning = True
        self.stop_scan = False
        
        # 更新按钮状态
        self.scan_btn.config(state=tk.DISABLED)
        self.deep_scan_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        # 显示扫描中提示
        self.progress_window = tk.Toplevel(self.root)
        self.progress_window.title("扫描中")
        self.progress_window.geometry("300x120")
        self.progress_window.transient(self.root)
        self.progress_window.grab_set()
        
        # 设置进度窗口图标
        self.set_child_window_icon(self.progress_window)
        
        # 居中显示
        self.progress_window.update_idletasks()
        width = self.progress_window.winfo_width()
        height = self.progress_window.winfo_height()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (width // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (height // 2)
        self.progress_window.geometry(f"{width}x{height}+{x}+{y}")
        
        ttk.Label(self.progress_window, text="正在扫描文件夹大小，请稍候...", padding=10).pack()
        self.size_progress = ttk.Progressbar(self.progress_window, orient="horizontal", length=250, mode="determinate")
        self.size_progress.pack(pady=5)
        self.size_status = ttk.Label(self.progress_window, text="准备开始扫描...")
        self.size_status.pack(pady=5)
        
        # 在新线程中执行扫描，避免UI冻结
        scan_thread = threading.Thread(target=self.perform_size_scan)
        scan_thread.daemon = True
        scan_thread.start()
        
        # 检查扫描线程是否完成
        self.root.after(100, self.check_size_scan_complete)
    
    def set_child_window_icon(self, window):
        """为子窗口设置图标"""
        try:
            if getattr(sys, 'frozen', False):
                base_dir = sys._MEIPASS
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))
            
            icon_path = os.path.join(base_dir, "wechatclean.ico")
            
            if os.path.exists(icon_path):
                window.iconbitmap(icon_path)
        except:
            pass  # 子窗口图标设置失败不影响主功能
    
    def perform_size_scan(self):
        """执行大小扫描的实际操作"""
        total_paths = len(self.all_paths)
        completed = 0
        
        for path in self.all_paths:
            # 检查是否需要停止
            if self.stop_scan:
                return
                
            # 更新进度
            completed += 1
            progress = (completed / total_paths) * 100
            self.size_progress["value"] = progress
            self.size_status.config(text=f"正在扫描: {os.path.basename(path)}")
            
            if os.path.exists(path):
                size = self.get_folder_size(path)
                if size == -1:  # 被停止
                    return
                self.folder_sizes[path] = self.format_size(size)
    
    def check_size_scan_complete(self):
        """检查大小扫描是否完成"""
        if self.scanning and hasattr(self, 'progress_window') and self.progress_window.winfo_exists():
            # 继续等待
            self.root.after(100, self.check_size_scan_complete)
            return
            
        # 恢复按钮状态
        self.scanning = False
        self.scan_btn.config(state=tk.NORMAL)
        self.deep_scan_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        
        # 关闭进度窗口并更新列表
        if hasattr(self, 'progress_window') and self.progress_window.winfo_exists():
            self.progress_window.destroy()
        
        self.update_folder_list()
        
        if self.stop_scan:
            messagebox.showinfo("已停止", "文件夹大小扫描已被终止")
            self.stop_scan = False
        else:
            messagebox.showinfo("完成", "文件夹大小扫描完成")
    
    def get_all_drives(self):
        """获取系统中所有驱动器"""
        drives = []
        if platform.system() == "Windows":
            # Windows系统获取所有驱动器
            for drive in range(65, 91):  # A-Z
                drive_letter = chr(drive) + ":\\"
                if os.path.exists(drive_letter):
                    drives.append(drive_letter)
        else:
            # 其他系统默认扫描根目录
            drives.append("/")
        return drives
    
    def start_deep_scan(self):
        """开始深度扫描（在新线程中执行）"""
        if self.scanning:
            messagebox.showinfo("提示", "正在进行扫描操作，请等待完成")
            return
            
        # 确认深度扫描
        if not messagebox.askyesno("确认", "深度扫描将搜索所有驱动器中包含'wechat'字符的文件夹，\n限制扫描4层目录，可能需要几分钟时间，是否继续？"):
            return
            
        self.scanning = True
        self.stop_scan = False
        self.scanned_paths = []
        
        # 更新按钮状态
        self.scan_btn.config(state=tk.DISABLED)
        self.deep_scan_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        # 显示扫描进度窗口
        self.scan_progress_window = tk.Toplevel(self.root)
        self.scan_progress_window.title("深度扫描中")
        self.scan_progress_window.geometry("500x150")
        self.scan_progress_window.transient(self.root)
        self.scan_progress_window.grab_set()
        
        # 设置进度窗口图标
        self.set_child_window_icon(self.scan_progress_window)
        
        # 居中显示
        self.scan_progress_window.update_idletasks()
        width = self.scan_progress_window.winfo_width()
        height = self.scan_progress_window.winfo_height()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (width // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (height // 2)
        self.scan_progress_window.geometry(f"{width}x{height}+{x}+{y}")
        
        # 进度条和状态标签
        ttk.Label(self.scan_progress_window, text="正在扫描所有驱动器中的微信相关文件夹...", padding=10).pack()
        self.scan_progress = ttk.Progressbar(self.scan_progress_window, orient="horizontal", length=400, mode="indeterminate")
        self.scan_progress.pack(pady=10)
        self.scan_status = ttk.Label(self.scan_progress_window, text="准备开始扫描...")
        self.scan_status.pack(pady=5)
        
        self.scan_progress.start()
        
        # 在新线程中执行扫描，避免UI冻结
        scan_thread = threading.Thread(target=self.perform_deep_scan)
        scan_thread.daemon = True
        scan_thread.start()
        
        # 检查扫描线程是否完成
        self.root.after(100, self.check_scan_complete)
    
    def perform_deep_scan(self):
        """执行深度扫描的实际操作，限制在4层文件夹，支持停止"""
        # 正则表达式匹配包含wechat的文件夹（不区分大小写）
        pattern = re.compile(r'wechat', re.IGNORECASE)
        
        # 获取所有驱动器
        drives = self.get_all_drives()
        max_depth = 4  # 限制最大扫描深度为4层
        
        # 扫描每个驱动器
        for drive in drives:
            # 检查是否需要停止
            if self.stop_scan:
                return
                
            # 更新状态
            self.scan_status.config(text=f"正在扫描: {drive}")
            
            # 跳过默认路径所在的驱动器以提高效率
            default_drive = os.path.splitdrive(self.default_paths[0])[0] + "\\"
            if drive == default_drive:
                continue
                
            try:
                # 递归扫描驱动器，限制深度
                self.scan_directory(drive, 0, max_depth, pattern)
                
            except Exception as e:
                print(f"扫描驱动器 {drive} 时出错: {e}")
                continue
    
    def scan_directory(self, current_dir, current_depth, max_depth, pattern):
        """递归扫描目录，限制深度，支持停止"""
        # 检查是否需要停止
        if self.stop_scan:
            return
            
        if current_depth > max_depth:
            return
            
        # 检查当前目录是否包含wechat
        if pattern.search(current_dir):
            if current_dir not in self.all_paths:
                self.scanned_paths.append(current_dir)
                
        try:
            # 获取子目录
            dirs = [d for d in os.listdir(current_dir) 
                   if os.path.isdir(os.path.join(current_dir, d))]
            
            # 递归扫描子目录
            for dir_name in dirs:
                # 检查是否需要停止
                if self.stop_scan:
                    return
                    
                # 检查子目录名是否包含wechat
                if pattern.search(dir_name):
                    full_path = os.path.join(current_dir, dir_name)
                    if full_path not in self.all_paths:
                        self.scanned_paths.append(full_path)
                
                # 继续扫描下一层
                self.scan_directory(os.path.join(current_dir, dir_name), 
                                   current_depth + 1, max_depth, pattern)
                                   
        except Exception as e:
            # 忽略没有访问权限的目录
            if "访问被拒绝" not in str(e):
                print(f"扫描目录 {current_dir} 时出错: {e}")
            return
    
    def check_scan_complete(self):
        """检查深度扫描是否完成"""
        if self.scanning and hasattr(self, 'scan_progress_window') and self.scan_progress_window.winfo_exists():
            # 继续等待
            self.root.after(100, self.check_scan_complete)
            return
            
        # 恢复按钮状态
        self.scanning = False
        self.scan_btn.config(state=tk.NORMAL)
        self.deep_scan_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        
        # 扫描完成或被取消
        if hasattr(self, 'scan_progress_window') and self.scan_progress_window.winfo_exists():
            self.scan_progress.stop()
            self.scan_progress_window.destroy()
        
        if self.stop_scan:
            # 扫描被停止
            messagebox.showinfo("已停止", "深度扫描已被终止")
            self.stop_scan = False
        else:
            # 扫描正常完成
            if self.scanned_paths:
                # 去重处理
                unique_paths = list(set(self.scanned_paths))
                
                # 添加新扫描到的路径
                for path in unique_paths:
                    if path not in self.all_paths:
                        self.all_paths.append(path)
                        self.folder_states[path] = tk.BooleanVar(value=False)  # 新扫描的路径默认不选中
                
                self.update_folder_list()
                messagebox.showinfo("完成", f"深度扫描完成，共发现 {len(unique_paths)} 个相关文件夹")
            else:
                messagebox.showinfo("结果", "未发现其他包含'wechat'的文件夹")
    
    def stop_scanning(self):
        """停止当前正在进行的扫描操作"""
        if not self.scanning:
            return
            
        if messagebox.askyesno("确认停止", "确定要停止当前的扫描操作吗？"):
            self.stop_scan = True
    
    def delete_selected(self):
        """删除选中的文件夹内容"""
        if self.scanning:
            messagebox.showinfo("提示", "正在进行扫描操作，请等待完成")
            return
            
        # 获取选中的文件夹
        selected = [path for path in self.all_paths 
                   if self.folder_states[path].get() and os.path.exists(path)]
        
        if not selected:
            messagebox.showinfo("提示", "请先选择要删除的文件夹")
            return
        
        # 显示风险提示并确认
        confirm_msg = "警告：删除操作将清除以下文件夹中的所有内容，\n可能导致微信数据永久丢失！\n\n"
        confirm_msg += "\n".join(selected) + "\n\n您确定要继续吗？"
        
        if not messagebox.askyesno("确认删除", confirm_msg):
            return
        
        # 执行删除
        results = []
        for path in selected:
            try:
                # 遍历文件夹并删除内容
                for item in os.listdir(path):
                    item_path = os.path.join(path, item)
                    if os.path.isfile(item_path) or os.path.islink(item_path):
                        os.unlink(item_path)
                    elif os.path.isdir(item_path):
                        shutil.rmtree(item_path)
                results.append(f"成功清理: {path}")
            except Exception as e:
                results.append(f"清理失败 {path}: {str(e)}")
        
        # 重新扫描大小
        self.scan_folders()
        
        # 显示结果
        messagebox.showinfo("操作结果", "\n".join(results))
    
    def open_url(self, event):
        """打开网址链接"""
        url = "http://www.itvip.com.cn"
        try:
            if platform.system() == 'Windows':
                os.startfile(url)
            elif platform.system() == 'Darwin':
                subprocess.Popen(['open', url])
            else:
                subprocess.Popen(['xdg-open', url])
        except Exception as e:
            messagebox.showerror("错误", f"无法打开链接: {str(e)}")

def is_admin():
    """检查是否以管理员权限运行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """以管理员权限重新运行程序"""
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, " ".join(sys.argv), None, 1
    )

def main():
    # 检查系统
    if platform.system() != "Windows":
        messagebox.showerror("错误", "此工具仅支持Windows系统")
        return
    
    # 检查管理员权限
    if not is_admin():
        if messagebox.askyesno("权限请求", "清理微信文件夹需要管理员权限，是否以管理员身份重新运行？"):
            run_as_admin()
            return
    
    root = tk.Tk()
    app = WeChatCleaner(root)
    root.mainloop()

if __name__ == "__main__":
    main()
    