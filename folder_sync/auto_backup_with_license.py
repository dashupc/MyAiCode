import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import threading
import time
from datetime import datetime, timedelta
import sys
import hashlib
import winreg
import configparser
import pystray
from PIL import Image, ImageDraw
import random
import string

class BackupTool:
    def __init__(self, root, start_minimized=False):
        # 软件授权相关配置
        self.REGISTRATION_KEY = "backup_tool_reg_key"  # 注册表键名
        self.TRIAL_DAYS = 7  # 试用期天数
        self.is_registered = False
        self.trial_end_date = None
        
        # 主窗口设置
        self.root = root
        self.root.title("智能自动备份工具")
        self.root.geometry("700x550")
        self.root.resizable(True, True)
        
        # 窗口居中显示
        self.center_window()
        
        # 初始化日志文本组件的占位符
        self.log_text = None
        
        # 检查授权状态
        self.check_license_status()
        
        # 如果未授权且试用期已过，显示注册窗口
        if not self.is_registered and (self.trial_end_date is None or datetime.now() > self.trial_end_date):
            if not self.show_registration_window():
                sys.exit(0)  # 用户未注册且关闭了注册窗口，退出程序
        
        # 配置文件路径
        self.config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "backup_config.ini")
        if getattr(sys, 'frozen', False):
            self.config_path = os.path.join(os.path.dirname(sys.executable), "backup_config.ini")
        
        # 确定图标文件路径
        self.icon_path = self.find_icon_file()
        
        # 加载配置
        self.load_config()
        
        # 启动参数 - 是否最小化到托盘
        self.start_minimized = start_minimized
        
        # 确保中文显示正常
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("SimHei", 10))
        self.style.configure("TButton", font=("SimHei", 10))
        self.style.configure("TEntry", font=("SimHei", 10))
        self.style.configure("Footer.TLabel", font=("SimHei", 9))
        
        # 源文件夹和目标文件夹路径
        self.source_path = tk.StringVar(value=self.config.get("Paths", "source", fallback=""))
        self.dest_path = tk.StringVar(value=self.config.get("Paths", "dest", fallback=""))
        
        # 监控间隔（秒）
        self.monitor_interval = tk.IntVar(value=self.config.getint("Settings", "interval", fallback=60))
        
        # 上次监控状态（用于重启时恢复）
        self.last_monitoring_state = self.config.getboolean("Settings", "monitoring", fallback=False)
        
        # 备份状态
        self.backup_running = False
        self.monitoring = False
        self.file_hashes = {}  # 存储文件哈希值，用于检测变化
        self.tray_icon = None  # 系统托盘图标
        
        # 创建UI
        self.create_widgets()
        
        # 设置窗口图标
        self.set_window_icon()
        
        # 检查是否在启动项中并更新按钮状态
        self.update_startup_button()
        
        # 显示试用期信息（如果未注册）
        if not self.is_registered:
            days_left = (self.trial_end_date - datetime.now()).days
            if days_left >= 0:
                self.log(f"试用期剩余: {days_left + 1} 天")
            self.add_registration_menu()
        
        # 重写窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # 初始化完成后处理
        self.post_init()
    
    # 授权验证相关方法
    def check_license_status(self):
        """检查软件授权状态"""
        try:
            # 尝试从注册表读取注册信息
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\BackupTool",
                0,
                winreg.KEY_READ
            )
            
            # 检查是否已注册
            registered_value, _ = winreg.QueryValueEx(key, self.REGISTRATION_KEY)
            winreg.CloseKey(key)
            
            # 验证注册码（这里使用简单的验证，实际应用中应更复杂）
            if self.verify_registration_key(registered_value):
                self.is_registered = True
                return
        except:
            # 注册表中没有注册信息，检查试用期
            pass
        
        # 检查试用期
        self.check_trial_period()
    
    def check_trial_period(self):
        """检查试用期状态"""
        try:
            # 尝试从注册表读取试用期信息
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\BackupTool",
                0,
                winreg.KEY_READ
            )
            
            # 读取试用期结束日期
            trial_end_str, _ = winreg.QueryValueEx(key, "TrialEndDate")
            winreg.CloseKey(key)
            self.trial_end_date = datetime.strptime(trial_end_str, "%Y-%m-%d")
            
        except:
            # 首次运行，设置试用期
            self.trial_end_date = datetime.now() + timedelta(days=self.TRIAL_DAYS)
            
            # 保存试用期信息到注册表
            try:
                key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\BackupTool")
                winreg.SetValueEx(key, "TrialEndDate", 0, winreg.REG_SZ, self.trial_end_date.strftime("%Y-%m-%d"))
                winreg.CloseKey(key)
            except Exception as e:
                print(f"保存试用期信息失败: {str(e)}")
    
    def verify_registration_key(self, key):
        """验证注册码（实际应用中应使用更安全的算法）"""
        # 这里使用简单的验证逻辑，实际应用中应使用非对称加密等更安全的方式
        if len(key) != 20:
            return False
            
        # 注册码格式：XXXX-XXXX-XXXX-XXXX-XXXX
        if key[4] != '-' or key[9] != '-' or key[14] != '-' or key[19] != '-':
            return False
            
        # 简单的校验和验证
        try:
            part1 = key[:4]
            part2 = key[5:9]
            part3 = key[10:14]
            part4 = key[15:19]
            
            # 计算校验和（示例算法）
            checksum = sum(ord(c) for c in part1 + part2 + part3) % 10000
            return part4 == f"{checksum:04d}"
        except:
            return False
    
    def show_registration_window(self):
        """显示注册窗口"""
        reg_window = tk.Toplevel(self.root)
        reg_window.title("软件注册")
        reg_window.geometry("400x250")
        reg_window.resizable(False, False)
        reg_window.transient(self.root)
        reg_window.grab_set()  # 模态窗口
        self.center_window(reg_window)
        
        # 标题
        ttk.Label(reg_window, text="软件注册", font=("SimHei", 16, "bold")).pack(pady=10)
        
        # 试用期已过提示
        if self.trial_end_date is not None and datetime.now() > self.trial_end_date:
            ttk.Label(reg_window, text="试用期已结束，请输入注册码以继续使用。", 
                    font=("SimHei", 10), foreground="red").pack(pady=10)
        else:
            ttk.Label(reg_window, text="请输入注册码以解锁全部功能，或继续试用。", 
                    font=("SimHei", 10)).pack(pady=10)
        
        # 注册码输入框
        frame = ttk.Frame(reg_window)
        frame.pack(pady=10, fill=tk.X, padx=20)
        
        ttk.Label(frame, text="注册码:").pack(side=tk.LEFT)
        reg_key_var = tk.StringVar()
        ttk.Entry(frame, textvariable=reg_key_var, width=30).pack(side=tk.LEFT, padx=5)
        
        # 按钮
        btn_frame = ttk.Frame(reg_window)
        btn_frame.pack(pady=20)
        
        result = [False]  # 使用列表存储结果，以便在内部函数中修改
        
        def register():
            key = reg_key_var.get().strip()
            if self.verify_registration_key(key):
                # 保存注册信息到注册表
                try:
                    key_path = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\BackupTool")
                    winreg.SetValueEx(key_path, self.REGISTRATION_KEY, 0, winreg.REG_SZ, key)
                    winreg.CloseKey(key_path)
                    
                    self.is_registered = True
                    messagebox.showinfo("成功", "注册成功，感谢您的支持！")
                    result[0] = True
                    reg_window.destroy()
                except Exception as e:
                    messagebox.showerror("错误", f"注册失败: {str(e)}")
            else:
                messagebox.showerror("错误", "无效的注册码，请重新输入。")
        
        def continue_trial():
            if self.trial_end_date is None or datetime.now() > self.trial_end_date:
                messagebox.showerror("错误", "试用期已结束，请注册后使用。")
                return
                
            result[0] = True
            reg_window.destroy()
        
        ttk.Button(btn_frame, text="注册", command=register).pack(side=tk.LEFT, padx=10)
        
        # 只有还在试用期内才显示继续试用按钮
        if self.trial_end_date is not None and datetime.now() <= self.trial_end_date:
            ttk.Button(btn_frame, text="继续试用", command=continue_trial).pack(side=tk.LEFT, padx=10)
        
        ttk.Button(btn_frame, text="退出", command=reg_window.destroy).pack(side=tk.LEFT, padx=10)
        
        self.root.wait_window(reg_window)
        return result[0]
    
    def add_registration_menu(self):
        """添加注册菜单"""
        # 创建注册按钮
        reg_button = ttk.Button(self.root, text="注册软件", command=self.show_registration_window)
        reg_button.place(relx=0.95, rely=0.02, anchor=tk.NE)
        
        # 显示试用期剩余天数
        days_left = (self.trial_end_date - datetime.now()).days + 1
        if days_left > 0:
            trial_label = ttk.Label(
                self.root, 
                text=f"试用期剩余: {days_left} 天", 
                foreground="orange"
            )
            trial_label.place(relx=0.8, rely=0.02, anchor=tk.NE)
    
    # 以下是原有的其他方法，保持不变
    def find_icon_file(self):
        """在多个可能的位置查找backup.ico图标文件"""
        possible_paths = []
        
        try:
            # 1. 程序当前运行目录
            current_dir = os.path.dirname(os.path.abspath(__file__)) if not getattr(sys, 'frozen', False) else os.path.dirname(sys.executable)
            possible_paths.append(os.path.join(current_dir, "backup.ico"))
            
            # 2. PyInstaller临时目录
            if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
                possible_paths.append(os.path.join(sys._MEIPASS, "backup.ico"))
            
            # 3. 用户桌面
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            possible_paths.append(os.path.join(desktop, "backup.ico"))
            
            # 4. 程序数据目录
            appdata = os.getenv('APPDATA')
            if appdata:
                app_dir = os.path.join(appdata, "BackupTool")
                possible_paths.append(os.path.join(app_dir, "backup.ico"))
            
            # 检查所有可能的路径
            for path in possible_paths:
                if os.path.exists(path) and os.path.isfile(path):
                    return path
            
            # 如果都找不到，返回None
            self.log(f"在以下位置均未找到backup.ico: {', '.join(possible_paths)}")
            return None
            
        except Exception as e:
            print(f"查找图标文件时出错: {str(e)}")
            return None
    
    def set_window_icon(self):
        """设置窗口标题栏图标为backup.ico"""
        if self.icon_path and os.path.exists(self.icon_path):
            try:
                self.root.iconbitmap(self.icon_path)
                self.log(f"已加载窗口图标: {self.icon_path}")
            except Exception as e:
                self.log(f"设置窗口图标失败: {str(e)}")
        else:
            self.log(f"未找到图标文件，使用默认图标")
    
    def post_init(self):
        """初始化完成后的操作"""
        # 如果是开机启动或命令行指定最小化，则隐藏窗口到托盘
        if self.start_minimized:
            self.root.withdraw()  # 隐藏窗口
            self.create_tray_icon()  # 创建托盘图标
            
            # 自动开始监控（如果配置有效）
            self.auto_start_monitoring()
        else:
            # 正常启动，显示窗口
            self.root.deiconify()
            
            # 如果上次是监控状态，询问是否恢复
            if self.last_monitoring_state:
                if messagebox.askyesno("恢复监控", "检测到上次退出时正在监控，是否恢复监控状态？"):
                    self.toggle_monitoring()
    
    def auto_start_monitoring(self):
        """自动启动监控（用于开机启动后）"""
        source = self.source_path.get()
        dest = self.dest_path.get()
        
        # 验证路径是否有效
        if source and os.path.isdir(source) and dest and os.path.isdir(dest):
            # 延迟启动，确保系统完全就绪
            threading.Timer(5, self.start_monitoring_silently).start()
        else:
            self.log("路径配置无效，无法自动启动监控，请检查设置")
    
    def start_monitoring_silently(self):
        """无提示启动监控（用于自动启动）"""
        try:
            self.monitoring = True
            self.update_monitor_button_text()
            self.calculate_initial_hashes()
            self.log(f"自动开始监控，间隔 {self.monitor_interval.get()} 秒")
            
            # 启动监控线程
            thread = threading.Thread(target=self.monitoring_thread)
            thread.daemon = True
            thread.start()
        except Exception as e:
            self.log(f"自动启动监控失败: {str(e)}")
    
    def center_window(self, window=None):
        """使窗口在屏幕中居中显示"""
        if window is None:
            window = self.root
            
        window.update_idletasks()
        width = window.winfo_width()
        height = window.winfo_height()
        x = (window.winfo_screenwidth() // 2) - (width // 2)
        y = (window.winfo_screenheight() // 2) - (height // 2)
        window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def create_widgets(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 源文件夹选择
        ttk.Label(main_frame, text="源文件夹:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.source_path, width=50).grid(row=0, column=1, pady=5)
        ttk.Button(main_frame, text="浏览...", command=self.select_source).grid(row=0, column=2, padx=5, pady=5)
        
        # 目标文件夹选择
        ttk.Label(main_frame, text="目标文件夹:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.dest_path, width=50).grid(row=1, column=1, pady=5)
        ttk.Button(main_frame, text="浏览...", command=self.select_dest).grid(row=1, column=2, padx=5, pady=5)
        
        # 监控间隔设置
        ttk.Label(main_frame, text="监控间隔(秒):").grid(row=2, column=0, sticky=tk.W, pady=5)
        interval_frame = ttk.Frame(main_frame)
        interval_frame.grid(row=2, column=1, sticky=tk.W, pady=5)
        ttk.Entry(interval_frame, textvariable=self.monitor_interval, width=10).pack(side=tk.LEFT)
        ttk.Label(interval_frame, text="建议设置60秒以上").pack(side=tk.LEFT, padx=5)
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=10)
        
        self.backup_btn = ttk.Button(button_frame, text="立即备份", command=self.start_backup)
        self.backup_btn.pack(side=tk.LEFT, padx=5)
        
        self.monitor_btn = ttk.Button(button_frame, text="开始监控", command=self.toggle_monitoring)
        self.monitor_btn.pack(side=tk.LEFT, padx=5)
        
        # 开机启动按钮 - 将根据状态动态变化
        self.startup_btn = ttk.Button(button_frame, text="添加开机启动", command=self.toggle_startup)
        self.startup_btn.pack(side=tk.LEFT, padx=5)
        
        self.save_btn = ttk.Button(button_frame, text="保存配置", command=self.save_config)
        self.save_btn.pack(side=tk.LEFT, padx=5)
        
        # 状态标签
        ttk.Label(main_frame, text="当前状态:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(main_frame, textvariable=self.status_var).grid(row=4, column=1, sticky=tk.W, pady=5)
        
        # 开机启动状态
        ttk.Label(main_frame, text="开机启动:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.startup_status_var = tk.StringVar(value="未设置")
        ttk.Label(main_frame, textvariable=self.startup_status_var).grid(row=5, column=1, sticky=tk.W, pady=5)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=6, column=0, columnspan=3, sticky=tk.EW, pady=10)
        
        # 日志区域
        ttk.Label(main_frame, text="操作日志:").grid(row=7, column=0, sticky=tk.NW, pady=5)
        self.log_text = tk.Text(main_frame, height=12, width=70)
        self.log_text.grid(row=7, column=1, columnspan=2, pady=5, sticky=tk.NSEW)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(main_frame, command=self.log_text.yview)
        scrollbar.grid(row=7, column=3, sticky=tk.NS)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 联系方式和网址（页脚）
        footer_frame = ttk.Frame(main_frame)
        footer_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="nsew")
        
        ttk.Label(footer_frame, text="QQ: 88179096", style="Footer.TLabel").pack(side=tk.LEFT, padx=10)
        ttk.Label(footer_frame, text="网址: www.itvip.com.cn", style="Footer.TLabel").pack(side=tk.LEFT, padx=10)
        
        # 设置网格权重
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(7, weight=1)
    
    def update_startup_button(self):
        """根据当前是否在开机启动中更新按钮文本和状态显示"""
        if self.is_in_startup():
            self.startup_btn.config(text="移除开机启动")
            self.startup_status_var.set("已设置")
        else:
            self.startup_btn.config(text="添加开机启动")
            self.startup_status_var.set("未设置")
    
    def update_monitor_button_text(self):
        """更新监控按钮的文本"""
        if self.monitoring:
            self.monitor_btn.config(text="停止监控")
        else:
            self.monitor_btn.config(text="开始监控")
    
    def select_source(self):
        path = filedialog.askdirectory(title="选择源文件夹")
        if path:
            self.source_path.set(path)
            self.calculate_initial_hashes()
    
    def select_dest(self):
        path = filedialog.askdirectory(title="选择目标文件夹")
        if path:
            self.dest_path.set(path)
    
    def log(self, message):
        """添加日志信息"""
        # 确保log_text已初始化
        if self.log_text is None:
            print(f"日志组件未准备好: {message}")
            return
            
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)  # 滚动到最新内容
    
    def update_progress(self, value):
        """更新进度条"""
        self.progress_var.set(value)
        self.root.update_idletasks()
    
    def update_status(self, status):
        """更新状态文本"""
        self.status_var.set(status)
        self.root.update_idletasks()
        
        # 更新系统托盘提示
        if self.tray_icon:
            self.tray_icon.title = f"智能自动备份工具 - {status}"
    
    def get_file_hash(self, file_path):
        """计算文件的MD5哈希值，用于检测文件变化"""
        try:
            hasher = hashlib.md5()
            with open(file_path, 'rb') as f:
                while chunk := f.read(4096):
                    hasher.update(chunk)
            return hasher.hexdigest()
        except Exception as e:
            self.log(f"计算文件哈希时出错 {file_path}: {str(e)}")
            return None
    
    def calculate_initial_hashes(self):
        """计算源文件夹中所有文件的初始哈希值"""
        source = self.source_path.get()
        if not source or not os.path.isdir(source):
            return
            
        self.file_hashes = {}
        self.log("正在计算初始文件哈希值...")
        
        for root, dirs, files in os.walk(source):
            for file in files:
                file_path = os.path.join(root, file)
                rel_path = os.path.relpath(file_path, source)
                self.file_hashes[rel_path] = self.get_file_hash(file_path)
        
        self.log(f"已计算 {len(self.file_hashes)} 个文件的哈希值")
    
    def check_for_changes(self):
        """检查源文件夹是否有变化"""
        source = self.source_path.get()
        if not source or not os.path.isdir(source):
            return False, [], [], []
            
        current_hashes = {}
        new_files = []
        modified_files = []
        deleted_files = []
        
        # 检查当前文件
        for root, dirs, files in os.walk(source):
            for file in files:
                file_path = os.path.join(root, file)
                rel_path = os.path.relpath(file_path, source)
                current_hashes[rel_path] = self.get_file_hash(file_path)
                
                # 检查是否是新文件或已修改的文件
                if rel_path not in self.file_hashes:
                    new_files.append(rel_path)
                else:
                    if current_hashes[rel_path] != self.file_hashes[rel_path]:
                        modified_files.append(rel_path)
        
        # 检查已删除的文件
        for rel_path in self.file_hashes:
            if rel_path not in current_hashes:
                deleted_files.append(rel_path)
        
        # 更新哈希值字典
        self.file_hashes = current_hashes
        
        # 返回是否有变化以及变化的文件列表
        has_changes = len(new_files) > 0 or len(modified_files) > 0 or len(deleted_files) > 0
        return has_changes, new_files, modified_files, deleted_files
    
    def copy_files(self, src, dest, specific_files=None):
        """复制文件和文件夹，可指定特定文件"""
        try:
            # 获取需要处理的项目
            if specific_files is None:
                items = os.listdir(src)
                total_items = len(items)
            else:
                # 处理特定文件
                items = []
                for file in specific_files:
                    items.append(os.path.normpath(file).split(os.sep)[0])
                items = list(set(items))  # 去重
                total_items = len(items)
                
            processed_items = 0
            
            for item in items:
                if not self.backup_running:  # 检查是否需要取消备份
                    self.log("备份已取消")
                    return False
                    
                src_path = os.path.join(src, item)
                dest_path = os.path.join(dest, item)
                
                try:
                    # 如果指定了特定文件，只处理这些文件
                    process_this_item = True
                    if specific_files is not None:
                        process_this_item = any(
                            os.path.normpath(file).startswith(os.path.normpath(item)) 
                            for file in specific_files
                        )
                    
                    if process_this_item:
                        if os.path.isfile(src_path):
                            # 复制文件
                            shutil.copy2(src_path, dest_path)
                            self.log(f"已复制文件: {src_path}")
                        elif os.path.isdir(src_path):
                            # 复制文件夹
                            if not os.path.exists(dest_path):
                                os.makedirs(dest_path)
                                self.log(f"已创建文件夹: {dest_path}")
                            
                            # 如果指定了特定文件，只复制其中的特定文件
                            if specific_files is not None:
                                sub_files = [
                                    file.split(os.sep, 1)[1] 
                                    for file in specific_files 
                                    if os.path.normpath(file).startswith(os.path.normpath(item))
                                    and len(file.split(os.sep, 1)) > 1
                                ]
                                if sub_files:
                                    if not self.copy_files(src_path, dest_path, sub_files):
                                        return False
                            else:
                                # 递归复制子目录和文件
                                if not self.copy_files(src_path, dest_path):
                                    return False
                    
                    processed_items += 1
                    progress = (processed_items / total_items) * 100
                    self.update_progress(progress)
                    
                except Exception as e:
                    self.log(f"处理 {src_path} 时出错: {str(e)}")
            
            return True
            
        except Exception as e:
            self.log(f"复制过程中出错: {str(e)}")
            return False
    
    def delete_files(self, dest, files_to_delete):
        """删除目标文件夹中对应的文件"""
        for rel_path in files_to_delete:
            file_path = os.path.join(dest, rel_path)
            try:
                if os.path.isfile(file_path):
                    os.remove(file_path)
                    self.log(f"已删除文件: {file_path}")
                elif os.path.isdir(file_path):
                    # 检查目录是否为空
                    if not os.listdir(file_path):
                        os.rmdir(file_path)
                        self.log(f"已删除空目录: {file_path}")
                    else:
                        self.log(f"无法删除非空目录: {file_path}")
            except Exception as e:
                self.log(f"删除 {file_path} 时出错: {str(e)}")
    
    def backup_thread(self, specific_files=None):
        """备份线程函数"""
        source = self.source_path.get()
        dest = self.dest_path.get()
        
        # 验证路径
        if not source or not os.path.isdir(source):
            self.log("错误: 无效的源文件夹")
            self.backup_running = False
            self.backup_btn.config(text="立即备份", state=tk.NORMAL)
            self.update_status("就绪")
            return
        
        if not dest or not os.path.isdir(dest):
            self.log("错误: 无效的目标文件夹")
            self.backup_running = False
            self.backup_btn.config(text="立即备份", state=tk.NORMAL)
            self.update_status("就绪")
            return
        
        # 清空进度条
        self.update_progress(0)
        
        if specific_files is None:
            self.log("开始完整备份...")
        else:
            self.log("检测到变化，开始增量备份...")
            
        self.log(f"源文件夹: {source}")
        self.log(f"目标文件夹: {dest}")
        
        start_time = time.time()
        
        # 执行备份
        success = self.copy_files(source, dest, specific_files)
        
        # 如果有删除的文件，在目标文件夹中也删除
        if specific_files is None and hasattr(self, 'deleted_files') and self.deleted_files:
            self.delete_files(dest, self.deleted_files)
        
        end_time = time.time()
        elapsed_time = end_time - start_time
        
        if success:
            self.log(f"备份完成! 耗时: {elapsed_time:.2f} 秒")
            self.update_status("就绪")
        else:
            self.log(f"备份中断! 耗时: {elapsed_time:.2f} 秒")
            self.update_status("备份中断")
        
        # 重置备份状态
        self.backup_running = False
        self.backup_btn.config(text="立即备份", state=tk.NORMAL)
        self.update_progress(100)
    
    def start_backup(self):
        """开始备份"""
        if not self.backup_running:
            self.backup_running = True
            self.backup_btn.config(text="取消备份", state=tk.DISABLED)
            self.update_status("备份中...")
            # 在新线程中执行备份，避免UI冻结
            thread = threading.Thread(target=self.backup_thread)
            thread.daemon = True
            thread.start()
    
    def monitoring_thread(self):
        """监控线程函数，定期检查文件变化"""
        while self.monitoring:
            self.update_status(f"监控中 (间隔: {self.monitor_interval.get()}秒)")
            
            # 检查文件变化
            has_changes, new_files, modified_files, deleted_files = self.check_for_changes()
            
            if has_changes:
                self.log(f"检测到变化 - 新增: {len(new_files)}, 修改: {len(modified_files)}, 删除: {len(deleted_files)}")
                
                # 保存删除的文件列表，用于在备份时同步删除
                self.deleted_files = deleted_files
                
                # 合并需要更新的文件列表
                files_to_update = new_files + modified_files
                
                # 执行增量备份
                if not self.backup_running:
                    self.backup_running = True
                    self.backup_btn.config(text="取消备份", state=tk.DISABLED)
                    self.update_status("增量备份中...")
                    thread = threading.Thread(target=self.backup_thread, args=(files_to_update,))
                    thread.daemon = True
                    thread.start()
            
            # 等待指定的间隔时间
            for _ in range(self.monitor_interval.get()):
                if not self.monitoring:
                    break
                time.sleep(1)
    
    def toggle_monitoring(self):
        """切换监控状态"""
        if not self.monitoring:
            # 开始监控
            source = self.source_path.get()
            dest = self.dest_path.get()
            
            if not source or not os.path.isdir(source):
                messagebox.showerror("错误", "请选择有效的源文件夹")
                return
            
            if not dest or not os.path.isdir(dest):
                messagebox.showerror("错误", "请选择有效的目标文件夹")
                return
            
            interval = self.monitor_interval.get()
            if interval < 10:
                if not messagebox.askyesno("警告", "监控间隔过短可能会影响系统性能，是否继续?"):
                    return
            
            self.monitoring = True
            self.update_monitor_button_text()
            self.calculate_initial_hashes()
            self.log(f"开始监控，间隔 {interval} 秒")
            
            # 启动监控线程
            thread = threading.Thread(target=self.monitoring_thread)
            thread.daemon = True
            thread.start()
        else:
            # 停止监控
            self.monitoring = False
            self.update_monitor_button_text()
            self.log("已停止监控")
            self.update_status("就绪")
    
    def toggle_startup(self):
        """添加或移除开机启动，并更新按钮状态"""
        if self.is_in_startup():
            # 移除开机启动
            if self.remove_from_startup():
                self.log("已从开机启动中移除")
            else:
                self.log("从开机启动中移除失败")
        else:
            # 添加到开机启动
            if self.add_to_startup():
                self.log("已添加到开机启动")
            else:
                self.log("添加到开机启动失败")
        
        # 无论操作成功与否，都更新按钮状态
        self.update_startup_button()
    
    def get_executable_path(self):
        """获取当前程序的路径"""
        if getattr(sys, 'frozen', False):
            # 打包后的EXE
            return sys.executable
        else:
            # 开发环境
            return os.path.abspath(__file__)
    
    def add_to_startup(self):
        """添加到Windows开机启动，包含最小化参数"""
        try:
            # 获取程序路径和名称
            exe_path = self.get_executable_path()
            exe_name = os.path.basename(exe_path)
            
            # 添加启动参数，用于标识是开机启动
            command = f'"{exe_path}" --minimized'
            
            # 打开注册表
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_SET_VALUE
            )
            
            # 添加注册表项，包含启动参数
            winreg.SetValueEx(key, exe_name, 0, winreg.REG_SZ, command)
            winreg.CloseKey(key)
            return True
        except Exception as e:
            self.log(f"添加到开机启动失败: {str(e)}")
            return False
    
    def remove_from_startup(self):
        """从Windows开机启动中移除"""
        try:
            # 获取程序名称
            exe_path = self.get_executable_path()
            exe_name = os.path.basename(exe_path)
            
            # 打开注册表
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_SET_VALUE
            )
            
            # 删除注册表项
            winreg.DeleteValue(key, exe_name)
            winreg.CloseKey(key)
            return True
        except Exception as e:
            self.log(f"从开机启动中移除失败: {str(e)}")
            return False
    
    def is_in_startup(self):
        """检查是否已在开机启动中"""
        try:
            # 获取程序名称
            exe_path = self.get_executable_path()
            exe_name = os.path.basename(exe_path)
            
            # 打开注册表
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_READ
            )
            
            # 尝试获取值
            winreg.QueryValueEx(key, exe_name)
            winreg.CloseKey(key)
            return True
        except FileNotFoundError:
            return False
        except Exception as e:
            self.log(f"检查开机启动状态出错: {str(e)}")
            return False
    
    def load_config(self):
        """加载配置文件"""
        self.config = configparser.ConfigParser()
        if os.path.exists(self.config_path):
            self.config.read(self.config_path, encoding="utf-8")
            if "Paths" not in self.config:
                self.config.add_section("Paths")
            if "Settings" not in self.config:
                self.config.add_section("Settings")
        else:
            self.config.add_section("Paths")
            self.config.add_section("Settings")
    
    def save_config(self):
        """保存配置文件，包括当前监控状态"""
        try:
            self.config.set("Paths", "source", self.source_path.get())
            self.config.set("Paths", "dest", self.dest_path.get())
            self.config.set("Settings", "interval", str(self.monitor_interval.get()))
            self.config.set("Settings", "monitoring", str(self.monitoring))  # 保存当前监控状态
            
            with open(self.config_path, 'w', encoding="utf-8") as f:
                self.config.write(f)
            
            self.log(f"配置已保存到 {self.config_path}")
            messagebox.showinfo("成功", "配置已保存")
        except Exception as e:
            self.log(f"保存配置失败: {str(e)}")
            messagebox.showerror("错误", f"保存配置失败: {str(e)}")
    
    # 系统托盘相关功能
    def create_tray_icon(self):
        """创建系统托盘图标，使用backup.ico"""
        # 尝试加载指定的图标文件
        try:
            if self.icon_path and os.path.exists(self.icon_path):
                icon = Image.open(self.icon_path)
                self.log(f"已加载系统托盘图标: {self.icon_path}")
            else:
                # 如果找不到图标文件，创建默认图标
                self.log(f"未找到图标文件，使用默认图标")
                width = 64
                height = 64
                color1 = "blue"
                color2 = "white"
                
                icon = Image.new("RGB", (width, height), color1)
                dc = ImageDraw.Draw(icon)
                dc.rectangle(
                    (width // 4, height // 4, width * 3 // 4, height * 3 // 4),
                    fill=color2
                )
                dc.rectangle(
                    (width // 2, height // 4, width * 3 // 4, height * 3 // 4),
                    fill=color1
                )
        except Exception as e:
            # 创建默认图标作为最后的 fallback
            self.log(f"加载图标失败: {str(e)}，使用默认图标")
            width = 64
            height = 64
            color1 = "blue"
            color2 = "white"
            
            icon = Image.new("RGB", (width, height), color1)
            dc = ImageDraw.Draw(icon)
            dc.rectangle(
                (width // 4, height // 4, width * 3 // 4, height * 3 // 4),
                fill=color2
            )
            dc.rectangle(
                (width // 2, height // 4, width * 3 // 4, height * 3 // 4),
                fill=color1
            )
        
        # 创建菜单
        menu = []
        
        # 添加注册选项（如果未注册）
        if not self.is_registered:
            menu.append(pystray.MenuItem("注册软件", self.show_registration_window))
            days_left = (self.trial_end_date - datetime.now()).days + 1
            if days_left > 0:
                menu.append(pystray.MenuItem(f"试用期剩余: {days_left} 天", None, enabled=False))
            menu.append(pystray.Menu.SEPARATOR)
        
        menu.extend([
            pystray.MenuItem("显示窗口", self.show_window),
            pystray.MenuItem("立即备份", self.tray_start_backup),
            pystray.MenuItem("退出程序", self.exit_program)
        ])
        
        # 创建托盘图标
        self.tray_icon = pystray.Icon("backup_tool", icon, "智能自动备份工具", tuple(menu))
        
        # 在单独的线程中运行托盘图标
        threading.Thread(target=self.tray_icon.run, daemon=True).start()
    
    def show_window(self):
        """显示主窗口"""
        self.root.deiconify()  # 显示窗口
        self.root.lift()  # 窗口置顶
        self.center_window()  # 重新居中窗口
    
    def tray_start_backup(self):
        """从系统托盘启动备份"""
        if not self.backup_running:
            self.start_backup()
    
    def exit_program(self):
        """退出整个程序"""
        # 保存当前监控状态
        self.config.set("Settings", "monitoring", str(self.monitoring))
        with open(self.config_path, 'w', encoding="utf-8") as f:
            self.config.write(f)
        
        # 停止监控
        self.monitoring = False
        
        # 停止托盘图标
        if self.tray_icon:
            self.tray_icon.stop()
        
        # 销毁窗口
        self.root.destroy()
        
        # 退出程序
        sys.exit(0)
    
    def on_close(self):
        """窗口关闭事件处理"""
        # 保存当前监控状态
        self.config.set("Settings", "monitoring", str(self.monitoring))
        with open(self.config_path, 'w', encoding="utf-8") as f:
            self.config.write(f)
            
        if self.monitoring:
            # 如果正在监控，最小化到托盘
            self.root.withdraw()  # 隐藏窗口
            if not self.tray_icon:
                self.create_tray_icon()  # 创建托盘图标
            self.log("程序已最小化到系统托盘，右键点击图标可操作")
        else:
            # 如果没有监控，询问是否退出
            if messagebox.askyesno("确认", "确定要退出程序吗？"):
                self.exit_program()


# 生成注册码的工具函数（仅用于开发者生成注册码）
def generate_registration_key():
    """生成有效的注册码（供开发者使用）"""
    # 生成前三部分随机字符
    part1 = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
    part2 = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
    part3 = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
    
    # 计算校验和
    checksum = sum(ord(c) for c in part1 + part2 + part3) % 10000
    part4 = f"{checksum:04d}"
    
    # 组合成注册码
    return f"{part1}-{part2}-{part3}-{part4}"


def main():
    # 检查命令行参数，判断是否需要最小化启动
    start_minimized = "--minimized" in sys.argv
    
    # 确保中文显示正常
    root = tk.Tk()
    
    app = BackupTool(root, start_minimized)
    root.mainloop()


if __name__ == "__main__":
    main()
    