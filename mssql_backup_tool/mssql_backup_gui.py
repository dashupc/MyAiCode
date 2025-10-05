import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pyodbc
import os
import datetime
import threading
import logging
import platform
import time
import requests
import sys
import ctypes
import json
import schedule
from PIL import Image, ImageTk, ImageDraw, ImageFont
import pystray
from pystray import MenuItem as item
import tempfile
import atexit

class MSSQLBackupTool:
    def __init__(self, root):
        self.root = root
        self.root.title("MSSQL数据库备份工具 - V0.01")
        self.root.geometry("900x700")
        
        # 核心状态变量
        self.version = "0.01"
        self.tray_icon = None
        self.tray_thread = None
        self.tray_healthy = False
        self.tray_restarting = False
        self.default_icon = None  # 仅保留默认图标
        self.startup_complete = False
        
        # 窗口状态管理
        self.is_minimized_to_tray = False
        
        # 窗口事件绑定
        self.root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)
        self.root.bind("<Unmap>", self.on_minimize)
        self.root.bind("<Map>", self.on_restore)
        
        # 初始化变量
        self.log_text = None
        self.backup_running = False
        self.auto_backup_enabled = False
        self.selected_databases = []
        self.os_type = platform.system()
        
        # 创建默认图标（红色背景+白色DB文字）
        self._create_default_icon()
        
        # 设置字体
        self.setup_fonts()
        
        # 创建界面
        self.create_widgets()
        
        # 初始化日志
        self.setup_logging()
        
        # 窗口配置
        self.root.attributes("-topmost", True)
        self.set_window_icon()
        self.center_window()
        
        # 加载配置
        self.load_config()
        
        # 启动服务
        self.start_schedule_thread()
        self.start_monitor_thread()
        
        # 初始化系统托盘
        self.init_system_tray()
        
        # 显示引导
        self.show_startup_guide()
        
        # 注册退出清理
        atexit.register(self.cleanup_on_exit)

    def _create_default_icon(self):
        """创建默认图标（红色背景+白色DB文字，确保高可见性）"""
        try:
            icon_size = (32, 32)
            self.default_icon = Image.new('RGB', icon_size, color='red')
            draw = ImageDraw.Draw(self.default_icon)
            
            # 尝试加载中文字体，失败则用默认字体
            try:
                font = ImageFont.truetype("simsun.ttc", 16)
            except Exception as e:
                self.log(f"加载中文字体失败: {str(e)}，使用默认字体")
                font = ImageFont.load_default()
                
            # 绘制"DB"文字（数据库缩写）
            draw.text((8, 6), "DB", font=font, fill='white')
            self.log("默认图标创建成功（红色背景+白色DB文字）")
            return True
        except Exception as e:
            self.log(f"创建默认图标失败: {str(e)}")
            # 极端失败时使用纯红色图标
            self.default_icon = Image.new('RGB', (32, 32), color='red')
            return False

    def setup_fonts(self):
        """设置中文字体支持，避免乱码"""
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("SimHei", 10))
        self.style.configure("TButton", font=("SimHei", 10))
        self.style.configure("TEntry", font=("SimHei", 10))
        self.style.configure("TCombobox", font=("SimHei", 10))

    def center_window(self):
        """窗口居中显示，确保启动时在屏幕中央"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def set_window_icon(self):
        """设置窗口图标，使用默认图标"""
        try:
            # 生成临时默认图标
            temp_dir = tempfile.gettempdir()
            temp_icon = os.path.join(temp_dir, "mssql_window_icon_v001.ico")
            self.default_icon.save(temp_icon)
            self.root.iconbitmap(temp_icon)
            
            # 注册退出时删除临时文件
            atexit.register(lambda: os.remove(temp_icon) if os.path.exists(temp_icon) else None)
            self.log(f"使用默认窗口图标: {temp_icon}")
        except Exception as e:
            self.log(f"设置窗口图标失败: {str(e)}")

    def setup_logging(self):
        """初始化日志系统，记录操作和错误信息"""
        logging.basicConfig(
            filename='backup_log_v001.txt',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            encoding='utf-8'
        )
        self.log(f"程序启动 (版本: {self.version})")
        self.log(f"操作系统类型: {self.os_type}")

    def create_widgets(self):
        """创建完整界面组件，包含监控、配置、日志三个标签页"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标签页控件
        tab_control = ttk.Notebook(main_frame)
        
        self.monitor_tab = ttk.Frame(tab_control)
        self.config_tab = ttk.Frame(tab_control)
        self.log_tab = ttk.Frame(tab_control)
        
        tab_control.add(self.monitor_tab, text="监控")
        tab_control.add(self.config_tab, text="配置")
        tab_control.add(self.log_tab, text="日志")
        
        tab_control.pack(expand=1, fill="both")
        
        # 构建各标签页内容
        self.create_config_tab()
        self.create_log_tab()
        self.create_monitor_tab()

    def create_config_tab(self):
        """配置标签页：数据库连接、备份设置、系统设置（已删除图标设置）"""
        config_frame = ttk.Frame(self.config_tab, padding="10")
        config_frame.pack(fill=tk.BOTH, expand=True)
        
        # 1. 数据库连接设置
        conn_frame = ttk.LabelFrame(config_frame, text="数据库连接设置", padding="10")
        conn_frame.pack(fill=tk.X, pady=5)
        
        # 服务器地址
        ttk.Label(conn_frame, text="服务器地址:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        self.server_entry = ttk.Entry(conn_frame, width=50)
        self.server_entry.grid(row=0, column=1, sticky=tk.W+tk.E, pady=5, padx=5)
        conn_frame.columnconfigure(1, weight=1)
        ttk.Label(conn_frame, text="(格式: 192.168.1.100\\SQLEXPRESS)").grid(row=0, column=2, sticky=tk.W, pady=5, padx=5)
        
        # 用户名
        ttk.Label(conn_frame, text="用户名:").grid(row=1, column=0, sticky=tk.W, pady=5, padx=5)
        self.user_entry = ttk.Entry(conn_frame, width=50)
        self.user_entry.grid(row=1, column=1, sticky=tk.W+tk.E, pady=5, padx=5)
        
        # 密码
        ttk.Label(conn_frame, text="密码:").grid(row=2, column=0, sticky=tk.W, pady=5, padx=5)
        self.password_entry = ttk.Entry(conn_frame, width=50, show="*")
        self.password_entry.grid(row=2, column=1, sticky=tk.W+tk.E, pady=5, padx=5)
        
        # 2. 数据库选择（支持多选）
        db_frame = ttk.LabelFrame(config_frame, text="数据库选择 (可多选)", padding="10")
        db_frame.pack(fill=tk.X, pady=5)
        
        # 控制按钮
        btn_frame = ttk.Frame(db_frame)
        btn_frame.pack(side=tk.RIGHT, padx=5, pady=5)
        ttk.Button(btn_frame, text="全选", command=self.select_all_databases).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="取消全选", command=self.deselect_all_databases).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="刷新列表", command=self.refresh_databases).pack(side=tk.RIGHT, padx=5)
        
        # 数据库列表
        ttk.Label(db_frame, text="选择要备份的数据库:").pack(side=tk.LEFT, padx=5, pady=5)
        db_list_frame = ttk.Frame(db_frame)
        db_list_frame.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)
        
        self.db_listbox = tk.Listbox(db_list_frame, selectmode=tk.MULTIPLE, width=50, height=4)
        self.db_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 列表滚动条
        db_scrollbar = ttk.Scrollbar(db_list_frame, orient=tk.VERTICAL, command=self.db_listbox.yview)
        db_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.db_listbox.config(yscrollcommand=db_scrollbar.set)
        
        # 3. 备份设置
        backup_frame = ttk.LabelFrame(config_frame, text="备份设置", padding="10")
        backup_frame.pack(fill=tk.X, pady=5)
        
        # 本地保存路径
        path_frame = ttk.Frame(backup_frame)
        path_frame.grid(row=0, column=0, columnspan=3, sticky=tk.W+tk.E, pady=5, padx=5)
        
        ttk.Label(path_frame, text="本地保存路径:").pack(side=tk.LEFT, padx=5, pady=5)
        self.local_save_path_entry = ttk.Entry(path_frame, width=50)
        self.local_save_path_entry.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)
        ttk.Button(path_frame, text="浏览...", command=self.browse_local_path).pack(side=tk.LEFT, padx=5, pady=5)
        
        # 文件名前缀
        ttk.Label(backup_frame, text="备份文件名前缀:").grid(row=1, column=0, sticky=tk.W, pady=5, padx=5)
        self.filename_prefix_entry = ttk.Entry(backup_frame, width=30)
        self.filename_prefix_entry.grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)
        ttk.Label(backup_frame, text="(自动添加数据库名和时间戳)").grid(row=1, column=2, sticky=tk.W, pady=5, padx=5)
        
        # 4. 自动清理设置
        cleanup_frame = ttk.LabelFrame(config_frame, text="自动清理设置", padding="10")
        cleanup_frame.pack(fill=tk.X, pady=5)
        
        self.auto_cleanup_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            cleanup_frame, 
            text="启用自动清理过期备份", 
            variable=self.auto_cleanup_var
        ).grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        
        ttk.Label(cleanup_frame, text="保留最近:").grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)
        self.retention_days = tk.IntVar(value=30)
        ttk.Combobox(
            cleanup_frame, 
            textvariable=self.retention_days,
            values=[7, 15, 30, 60, 90, 180, 365],
            width=5
        ).grid(row=0, column=2, sticky=tk.W, pady=5, padx=5)
        ttk.Label(cleanup_frame, text="天的备份文件").grid(row=0, column=3, sticky=tk.W, pady=5, padx=5)
        ttk.Button(cleanup_frame, text="立即清理", command=self.manual_cleanup).grid(row=0, column=4, sticky=tk.W, pady=5, padx=10)
        
        # 5. 系统设置（已删除图标设置相关内容）
        system_frame = ttk.LabelFrame(config_frame, text="系统设置", padding="10")
        system_frame.pack(fill=tk.X, pady=5)
        
        # 启动和最小化设置
        self.start_visible = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            system_frame, 
            text="启动时显示主窗口（推荐）", 
            variable=self.start_visible
        ).grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        
        self.minimize_to_tray = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            system_frame, 
            text="关闭窗口时最小化到系统托盘", 
            variable=self.minimize_to_tray
        ).grid(row=1, column=0, sticky=tk.W, pady=5, padx=5)
        
        # 6. 操作按钮（已删除图标相关按钮）
        button_frame = ttk.Frame(config_frame, padding="10")
        button_frame.pack(fill=tk.X, pady=5)
        
        self.backup_button = ttk.Button(button_frame, text="开始备份", command=self.start_backup)
        self.backup_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="测试连接", command=self.test_connection).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="保存配置", command=self.save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="重启托盘", command=self.restart_tray_icon).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="托盘位置", command=self.show_tray_location).pack(side=tk.LEFT, padx=5)

    def create_log_tab(self):
        """日志标签页：显示操作日志，支持清空和导出"""
        log_frame = ttk.Frame(self.log_tab, padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # 日志控制按钮
        log_control_frame = ttk.Frame(log_frame)
        log_control_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(log_control_frame, text="清空日志", command=self.clear_log).pack(side=tk.RIGHT, padx=5)
        ttk.Button(log_control_frame, text="导出日志", command=self.export_log).pack(side=tk.RIGHT, padx=5)
        
        # 日志显示区域
        log_display_frame = ttk.LabelFrame(log_frame, text="操作日志（最近100条）", padding="10")
        log_display_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = tk.Text(log_display_frame, height=25, wrap=tk.WORD, font=("Consolas", 9))
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)
        
        # 日志滚动条
        scrollbar = ttk.Scrollbar(log_display_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 加载历史日志
        self.load_history_log()

    def create_monitor_tab(self):
        """监控标签页：显示系统状态、备份统计、最近记录"""
        monitor_frame = ttk.Frame(self.monitor_tab, padding="10")
        monitor_frame.pack(fill=tk.BOTH, expand=True)
        
        # 1. 系统状态
        status_frame = ttk.LabelFrame(monitor_frame, text="系统状态", padding="10")
        status_frame.pack(fill=tk.X, pady=5)
        
        # 程序信息
        ttk.Label(status_frame, text="程序版本:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        ttk.Label(status_frame, text=self.version).grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)
        
        ttk.Label(status_frame, text="启动时间:").grid(row=0, column=2, sticky=tk.W, pady=5, padx=20)
        self.start_time_label = ttk.Label(status_frame, text=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.start_time_label.grid(row=0, column=3, sticky=tk.W, pady=5, padx=5)
        
        # 连接状态
        ttk.Label(status_frame, text="数据库连接:").grid(row=1, column=0, sticky=tk.W, pady=5, padx=5)
        self.db_status_label = ttk.Label(status_frame, text="未检测", foreground="orange")
        self.db_status_label.grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)
        
        ttk.Label(status_frame, text="系统托盘:").grid(row=1, column=2, sticky=tk.W, pady=5, padx=20)
        self.tray_status_label = ttk.Label(status_frame, text="初始化中", foreground="orange")
        self.tray_status_label.grid(row=1, column=3, sticky=tk.W, pady=5, padx=5)
        
        # 2. 备份统计
        stats_frame = ttk.LabelFrame(monitor_frame, text="今日备份统计", padding="10")
        stats_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(stats_frame, text="成功次数:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=20)
        self.success_count_label = ttk.Label(stats_frame, text="0", foreground="green")
        self.success_count_label.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)
        
        ttk.Label(stats_frame, text="失败次数:").grid(row=0, column=2, sticky=tk.W, pady=5, padx=20)
        self.fail_count_label = ttk.Label(stats_frame, text="0", foreground="red")
        self.fail_count_label.grid(row=0, column=3, sticky=tk.W, pady=5, padx=5)
        
        # 3. 最近备份记录
        recent_frame = ttk.LabelFrame(monitor_frame, text="最近备份记录（10条）", padding="10")
        recent_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 表格列定义
        columns = ("数据库名", "备份时间", "状态", "文件大小")
        self.recent_tree = ttk.Treeview(recent_frame, columns=columns, show="headings")
        
        # 设置列标题和宽度
        for col in columns:
            self.recent_tree.heading(col, text=col)
            self.recent_tree.column(col, width=180, anchor=tk.CENTER)
        
        # 表格滚动条
        scrollbar = ttk.Scrollbar(recent_frame, orient=tk.VERTICAL, command=self.recent_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.recent_tree.configure(yscrollcommand=scrollbar.set)
        
        self.recent_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 初始化统计数据
        self.today_success_count = 0
        self.today_fail_count = 0
        self.recent_backups = []

    def log(self, message):
        """记录日志到界面和文件"""
        if self.log_text is None:
            print(f"日志: {message}")
            return
            
        # 写入界面日志
        self.log_text.config(state=tk.NORMAL)
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        # 写入文件日志
        logging.info(message)

    # ==============================
    # 数据库选择相关方法
    # ==============================
    def select_all_databases(self):
        """全选数据库列表"""
        for i in range(self.db_listbox.size()):
            self.db_listbox.selection_set(i)
        self.log("已全选所有数据库")

    def deselect_all_databases(self):
        """取消全选数据库列表"""
        self.db_listbox.selection_clear(0, tk.END)
        self.log("已取消全选所有数据库")

    def get_selected_databases(self):
        """获取选中的数据库列表"""
        selected_indices = self.db_listbox.curselection()
        databases = [self.db_listbox.get(i) for i in selected_indices]
        return databases

    # ==============================
    # 系统托盘处理（仅使用默认图标）
    # ==============================
    def get_tray_icon(self):
        """获取托盘图标（仅使用默认图标）"""
        try:
            return self.default_icon
        except Exception as e:
            self.log(f"获取托盘图标失败: {str(e)}")
            # 极端情况返回纯红色图标
            return Image.new('RGB', (32, 32), color='red')

    def init_system_tray(self):
        """初始化系统托盘，独立线程运行避免阻塞"""
        self.tray_healthy = False
        self.tray_restarting = True
        self.tray_status_label.config(text="初始化中", foreground="orange")
        
        # 停止现有托盘
        if self.tray_icon:
            try:
                self.tray_icon.stop()
                self.log("已停止现有托盘图标")
            except Exception as e:
                self.log(f"停止托盘失败: {str(e)}")
        
        # 启动新托盘线程
        self.tray_thread = threading.Thread(target=self._run_tray, daemon=False)
        self.tray_thread.start()
        
        # 等待初始化并更新状态
        threading.Thread(target=self._wait_for_tray, daemon=True).start()

    def restart_tray_icon(self):
        """重启托盘图标"""
        self.log("用户请求重启系统托盘")
        self.init_system_tray()
        self.set_window_icon()
        messagebox.showinfo("提示", "系统托盘已重启，请查看屏幕右下角通知区域")

    def _run_tray(self):
        """托盘核心运行逻辑，支持重试机制"""
        try_count = 0
        max_tries = 5  # 最多重试5次
        while try_count < max_tries and not self.tray_healthy:
            try_count += 1
            try:
                # 获取图标（仅使用默认图标）
                icon_image = self.get_tray_icon()
                
                # 创建托盘菜单
                menu = (
                    item('显示主窗口', self.show_window, default=True),
                    item('立即备份', self.start_backup_from_tray),
                    item('退出程序', self.quit_application)
                )
                
                # 创建托盘图标实例
                self.tray_icon = pystray.Icon(
                    "mssql-backup-v001",
                    icon_image,
                    f"MSSQL备份工具V{self.version}",
                    menu
                )
                
                # 标记健康状态
                self.tray_healthy = True
                self.log(f"托盘初始化成功（第{try_count}/{max_tries}次尝试）")
                
                # 运行托盘（阻塞调用，在独立线程中执行）
                self.tray_icon.run()
                
            except Exception as e:
                self.log(f"托盘运行失败（第{try_count}/{max_tries}次尝试）: {str(e)}")
                self.tray_healthy = False
                time.sleep(2)  # 等待2秒后重试
        
        # 多次重试失败处理
        if not self.tray_healthy:
            self.log(f"托盘初始化失败（已尝试{max_tries}次）")
            self.minimize_to_tray.set(False)
            self.log("已自动禁用最小化到托盘功能")

    def _wait_for_tray(self):
        """等待托盘初始化完成，超时处理"""
        timeout = 15  # 15秒超时
        interval = 0.5  # 每0.5秒检查一次
        checks = int(timeout / interval)
        
        for _ in range(checks):
            if self.tray_healthy:
                # 托盘正常初始化
                self.tray_status_label.config(text="正常（右下角）", foreground="green")
                self.tray_restarting = False
                
                # 取消窗口置顶
                self.root.after(0, lambda: self.root.attributes("-topmost", False))
                self.startup_complete = True
                
                # 启动时最小化（如果设置）
                if not self.start_visible.get():
                    self.root.after(1000, self.minimize_to_tray)
                return
            time.sleep(interval)
        
        # 超时处理
        self.tray_status_label.config(text="异常（点击重启）", foreground="red")
        self.tray_restarting = False
        self.startup_complete = True
        self.root.after(0, lambda: self.root.attributes("-topmost", False))
        
        messagebox.showwarning(
            "托盘警告",
            "系统托盘初始化超时，已禁用最小化到托盘功能。\n"
            "请点击「重启托盘」按钮修复，或保持窗口可见。"
        )

    def show_tray_notification(self, title, message):
        """显示托盘通知"""
        try:
            if self.os_type == "Windows":
                # 优先使用win10toast
                try:
                    from win10toast import ToastNotifier
                    toaster = ToastNotifier()
                    # 使用默认图标
                    toaster.show_toast(
                        title, message,
                        duration=10
                    )
                except Exception as e:
                    self.log(f"win10toast通知失败: {str(e)}")
                    # 备选：系统消息框
                    ctypes.windll.user32.MessageBoxW(0, message, title, 0x40 | 0x1)
            else:
                # 其他系统用日志替代
                self.log(f"[{title}] {message}")
        except Exception as e:
            self.log(f"显示通知失败: {str(e)}")

    def show_tray_location(self):
        """显示托盘位置说明"""
        messagebox.showinfo(
            "托盘位置指南",
            "系统托盘通常位于屏幕右下角的任务栏通知区域。\n\n"
            "1. 直接查看任务栏右侧（时间日期旁边）\n"
            "2. 若找不到，点击任务栏右侧的「^」图标（显示隐藏图标）\n"
            "3. 本程序图标为：红色背景 + 白色「DB」文字\n\n"
            "双击图标可快速显示主窗口。"
        )

    # ==============================
    # 窗口控制相关方法
    # ==============================
    def show_startup_guide(self):
        """显示启动引导提示"""
        messagebox.showinfo(
            f"MSSQL备份工具V{self.version} - 使用指南",
            "欢迎使用MSSQL数据库备份工具！\n\n"
            "📌 基础操作：\n"
            "1. 配置页面：设置数据库连接和备份路径\n"
            "2. 选择数据库：在列表中勾选需要备份的数据库\n"
            "3. 开始备份：点击「开始备份」按钮执行手动备份\n\n"
            "🖥️ 窗口控制：\n"
            "• 关闭窗口 → 最小化到系统托盘（右下角）\n"
            "• 托盘图标 → 右键可显示菜单（显示窗口/备份/退出）\n"
            "• 启动时显示窗口 → 可在系统设置中关闭\n\n"
            "⚠️ 注意：首次使用请先测试数据库连接！"
        )

    def show_window(self, icon=None, item=None):
        """从托盘显示主窗口"""
        def _show():
            self.root.deiconify()  # 恢复窗口
            self.root.lift()  # 置顶
            self.root.attributes('-topmost', True)
            self.root.attributes('-topmost', False)
            self.is_minimized_to_tray = False
            self.window_status_label.config(text="正常", foreground="green")
            self.log("从托盘恢复主窗口显示")
            
        # 确保在主线程执行UI操作
        self.root.after(0, _show)

    def start_backup_from_tray(self, icon=None, item=None):
        """从托盘菜单启动备份"""
        self.show_window(icon, item)
        # 延迟启动，确保窗口已显示
        self.root.after(500, self.start_backup)

    def hide_window(self):
        """隐藏窗口到托盘"""
        if not self.tray_healthy:
            messagebox.showwarning("警告", "系统托盘功能异常，无法最小化到托盘")
            return
            
        self.root.withdraw()  # 完全隐藏（任务栏不显示）
        self.is_minimized_to_tray = True
        self.window_status_label.config(text="已最小化到托盘", foreground="blue")
        self.show_tray_notification("程序运行中", "已最小化到系统托盘\n双击图标可恢复窗口")

    def minimize_to_tray(self):
        """窗口关闭按钮 → 最小化到托盘"""
        if not self.minimize_to_tray.get():
            # 用户禁用了最小化功能，直接退出
            self.quit_application()
            return
            
        if not self.tray_healthy and not self.tray_restarting:
            # 托盘异常时询问用户
            if messagebox.askyesno("确认", "系统托盘功能异常，是否继续最小化？\n选择「否」将退出程序。"):
                self.hide_window()
                messagebox.showinfo("提示", "程序已最小化，请从任务管理器重启（若找不到托盘）")
            else:
                self.quit_application()
        else:
            # 正常最小化
            self.hide_window()
        
        return "break"  # 阻止窗口真正关闭

    def on_minimize(self, event):
        """窗口最小化事件（点击任务栏最小化按钮）"""
        if self.root.state() == 'iconic' and self.minimize_to_tray.get():
            self.log("检测到窗口最小化事件，转移到托盘")
            self.hide_window()

    def on_restore(self, event):
        """窗口恢复事件（从任务栏或托盘恢复）"""
        if self.is_minimized_to_tray:
            self.is_minimized_to_tray = False
            self.window_status_label.config(text="正常", foreground="green")
            self.log("窗口已从最小化状态恢复")

    def quit_application(self, icon=None, item=None):
        """完全退出程序"""
        self.log("用户请求退出程序")
        
        # 确认退出
        if not messagebox.askyesno("确认退出", f"确定要退出MSSQL备份工具V{self.version}吗？"):
            return
            
        # 清理托盘
        if hasattr(self, 'tray_icon') and self.tray_icon:
            try:
                self.tray_icon.stop()
                self.log("已停止托盘图标")
            except Exception as e:
                self.log(f"停止托盘失败: {str(e)}")
                
        # 退出程序
        self.root.destroy()
        sys.exit(0)

    def cleanup_on_exit(self):
        """程序退出时清理资源"""
        self.log(f"程序正常退出（版本: {self.version}）")

    # ==============================
    # 备份功能相关方法
    # ==============================
    def browse_local_path(self):
        """选择本地备份保存路径"""
        path = filedialog.askdirectory(title="选择本地备份保存路径")
        if path:
            self.local_save_path_entry.delete(0, tk.END)
            self.local_save_path_entry.insert(0, path)
            self.log(f"已选择本地保存路径: {path}")

    def get_connection_string(self, database="master"):
        """生成数据库连接字符串"""
        server = self.server_entry.get().strip()
        user = self.user_entry.get().strip()
        password = self.password_entry.get().strip()
        
        if not server:
            messagebox.showerror("错误", "请输入数据库服务器地址")
            return None
            
        # 构建连接字符串
        if user:
            # SQL Server身份验证
            conn_str = (
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={server};"
                f"DATABASE={database};"
                f"UID={user};"
                f"PWD={password};"
                f"AutoCommit=True;"
                f"TrustServerCertificate=yes"  # 忽略证书验证（开发环境）
            )
        else:
            # Windows身份验证
            conn_str = (
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={server};"
                f"DATABASE={database};"
                f"Trusted_Connection=yes;"
                f"AutoCommit=True;"
                f"TrustServerCertificate=yes"
            )
            
        return conn_str

    def test_connection(self):
        """测试数据库连接"""
        try:
            self.log("开始测试数据库连接...")
            conn_str = self.get_connection_string()
            if not conn_str:
                return
                
            # 尝试连接
            conn = pyodbc.connect(conn_str, timeout=10)
            conn.close()
            
            self.log("数据库连接测试成功")
            messagebox.showinfo("成功", "数据库连接测试成功！")
            self.db_status_label.config(text="正常", foreground="green")
        except Exception as e:
            error_msg = f"连接失败: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)
            self.db_status_label.config(text="失败", foreground="red")

    def refresh_databases(self):
        """刷新数据库列表"""
        try:
            self.log("开始刷新数据库列表...")
            conn_str = self.get_connection_string()
            if not conn_str:
                return
                
            # 连接数据库
            conn = pyodbc.connect(conn_str, timeout=10)
            cursor = conn.cursor()
            
            # 查询非系统数据库
            cursor.execute("""
                SELECT name 
                FROM sys.databases 
                WHERE name NOT IN ('master', 'tempdb', 'model', 'msdb')
                ORDER BY name
            """)
            databases = [row[0] for row in cursor.fetchall()]
            conn.close()
            
            # 保存当前选择
            current_selection = self.get_selected_databases()
            
            # 更新列表
            self.db_listbox.delete(0, tk.END)
            for db in databases:
                self.db_listbox.insert(tk.END, db)
            
            # 恢复之前的选择
            if current_selection:
                for i, db in enumerate(databases):
                    if db in current_selection:
                        self.db_listbox.selection_set(i)
            
            self.log(f"已加载 {len(databases)} 个数据库，恢复 {len(self.db_listbox.curselection())} 个选中项")
            self.db_status_label.config(text="正常", foreground="green")
        except Exception as e:
            error_msg = f"刷新数据库失败: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)
            self.db_status_label.config(text="失败", foreground="red")

    def save_config(self):
        """保存配置到文件（不含图标设置）"""
        try:
            # 获取配置信息
            config = {
                'server': self.server_entry.get().strip(),
                'user': self.user_entry.get().strip(),
                'password': self.password_entry.get().strip(),
                'local_save_path': self.local_save_path_entry.get().strip(),
                'filename_prefix': self.filename_prefix_entry.get().strip(),
                'auto_cleanup_var': self.auto_cleanup_var.get(),
                'retention_days': self.retention_days.get(),
                'start_visible': self.start_visible.get(),
                'minimize_to_tray': self.minimize_to_tray.get(),
                'selected_databases': self.get_selected_databases(),
                'version': self.version
            }
            
            # 保存到JSON文件
            with open('backup_config_v001.json', 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
            
            self.log("配置保存成功")
            messagebox.showinfo("成功", "配置已保存到 backup_config_v001.json")
        except Exception as e:
            error_msg = f"保存配置失败: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)

    def load_config(self):
        """从文件加载配置（不含图标设置）"""
        try:
            config_path = 'backup_config_v001.json'
            if not os.path.exists(config_path):
                self.log("配置文件不存在，使用默认配置")
                return
                
            # 读取配置
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                
                # 加载基础配置
                self.server_entry.insert(0, config.get('server', ''))
                self.user_entry.insert(0, config.get('user', ''))
                self.password_entry.insert(0, config.get('password', ''))
                self.local_save_path_entry.insert(0, config.get('local_save_path', ''))
                self.filename_prefix_entry.insert(0, config.get('filename_prefix', 'backup'))
                
                # 加载清理设置
                self.auto_cleanup_var.set(config.get('auto_cleanup_var', True))
                self.retention_days.set(config.get('retention_days', 30))
                
                # 加载系统设置
                self.start_visible.set(config.get('start_visible', True))
                self.minimize_to_tray.set(config.get('minimize_to_tray', True))
                
                # 加载选中的数据库（将在刷新后恢复）
                self.selected_databases = config.get('selected_databases', [])
                self.log(f"从配置加载 {len(self.selected_databases)} 个选中数据库")
                
            self.log("配置加载成功")
            # 刷新数据库列表（恢复选中状态）
            self.refresh_databases()
        except Exception as e:
            error_msg = f"加载配置失败: {str(e)}"
            self.log(error_msg)

    def delete_old_backups(self):
        """删除过期备份文件（自动清理）"""
        try:
            if not self.auto_cleanup_var.get():
                self.log("自动清理已禁用，跳过删除过期备份")
                return 0, 0
                
            local_path = self.local_save_path_entry.get().strip()
            if not local_path or not os.path.exists(local_path):
                self.log(f"本地路径不存在: {local_path}，无法清理")
                return 0, 0
                
            # 计算过期时间
            days = self.retention_days.get()
            if days <= 0:
                self.log(f"保留天数无效: {days}，跳过清理")
                return 0, 0
                
            cutoff_time = datetime.datetime.now() - datetime.timedelta(days=days)
            self.log(f"开始自动清理：删除 {days} 天前的备份（截止 {cutoff_time.strftime('%Y-%m-%d')}）")
            
            # 文件名前缀
            prefix = self.filename_prefix_entry.get().strip() or 'backup'
            
            # 统计信息
            deleted_count = 0
            kept_count = 0
            
            # 遍历文件
            for filename in os.listdir(local_path):
                if filename.startswith(prefix) and filename.endswith('.bak'):
                    file_path = os.path.join(local_path, filename)
                    
                    # 获取文件创建时间
                    try:
                        if self.os_type == 'Windows':
                            create_time = datetime.datetime.fromtimestamp(os.path.getctime(file_path))
                        else:
                            create_time = datetime.datetime.fromtimestamp(os.path.getmtime(file_path))
                            
                        if create_time < cutoff_time:
                            # 删除过期文件
                            os.remove(file_path)
                            self.log(f"已删除过期备份: {filename}（创建时间: {create_time.strftime('%Y-%m-%d')}）")
                            deleted_count += 1
                        else:
                            kept_count += 1
                    except Exception as e:
                        self.log(f"处理文件 {filename} 失败: {str(e)}")
            
            self.log(f"清理完成：删除 {deleted_count} 个过期备份，保留 {kept_count} 个最新备份")
            return deleted_count, kept_count
        except Exception as e:
            self.log(f"自动清理失败: {str(e)}")
            return 0, 0

    def manual_cleanup(self):
        """手动触发清理操作"""
        try:
            self.log("用户手动触发清理操作...")
            deleted, kept = self.delete_old_backups()
            
            messagebox.showinfo(
                "清理完成",
                f"手动清理结果：\n"
                f"✅ 已删除过期备份：{deleted} 个\n"
                f"📦 保留最新备份：{kept} 个\n\n"
                f"清理规则：保留最近 {self.retention_days.get()} 天的备份文件"
            )
        except Exception as e:
            error_msg = f"手动清理失败: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)

    def backup_single_database(self, db_name, is_auto=False):
        """备份单个数据库"""
        try:
            local_path = self.local_save_path_entry.get().strip()
            if not local_path:
                self.log("未设置本地保存路径")
                if not is_auto:
                    messagebox.showerror("错误", "请先选择本地保存路径")
                return False
                
            # 确保路径存在
            if not os.path.exists(local_path):
                try:
                    os.makedirs(local_path)
                    self.log(f"已创建本地路径: {local_path}")
                except Exception as e:
                    error_msg = f"创建路径失败: {str(e)}"
                    self.log(error_msg)
                    if not is_auto:
                        messagebox.showerror("错误", error_msg)
                    return False
                
            # 生成备份文件名
            prefix = self.filename_prefix_entry.get().strip() or 'backup'
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f"{prefix}_{db_name}_{timestamp}.bak"
            local_full_path = os.path.join(local_path, backup_filename)
            
            # 获取连接字符串
            conn_str = self.get_connection_string()
            if not conn_str:
                return False
                
            self.log(f"开始备份数据库: {db_name}")
            
            # 连接数据库执行备份
            conn = pyodbc.connect(conn_str, timeout=30)
            conn.autocommit = True
            cursor = conn.cursor()
            
            # 构建备份SQL
            backup_sql = f"""
                BACKUP DATABASE [{db_name}] 
                TO DISK = N'{local_full_path}' 
                WITH NOFORMAT, NOINIT, 
                NAME = N'{db_name}-完整备份', 
                SKIP, NOREWIND, NOUNLOAD, STATS = 10
            """
            
            # 执行备份
            cursor.execute(backup_sql)
            
            # 等待备份完成（检查文件锁定）
            self.log(f"等待 {db_name} 备份完成...")
            if self.wait_for_backup_completion(local_full_path):
                self.log(f"{db_name} 备份完成")
            else:
                self.log(f"{db_name} 备份可能未完成（超时）")
                
            conn.close()
            
            # 验证备份文件
            if os.path.exists(local_full_path):
                file_size = os.path.getsize(local_full_path)
                size_str = f"{file_size / (1024*1024):.2f} MB"
                self.log(f"{db_name} 备份成功！文件大小: {size_str}")
                
                # 更新监控记录
                self.add_backup_record(db_name, timestamp, "成功", size_str)
                self.update_daily_stats(success=True)
                return True
            else:
                self.log(f"{db_name} 备份文件不存在: {local_full_path}")
                self.add_backup_record(db_name, timestamp, "失败", "0 MB")
                self.update_daily_stats(success=False)
                return False
                
        except Exception as e:
            error_msg = f"{db_name} 备份失败: {str(e)}"
            self.log(error_msg)
            self.add_backup_record(db_name, datetime.datetime.now().strftime("%Y%m%d_%H%M%S"), "失败", "0 MB")
            self.update_daily_stats(success=False)
            return False

    def wait_for_backup_completion(self, file_path, timeout=300):
        """等待备份完成（检查文件是否被锁定）"""
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # 尝试写入文件，能写入说明备份完成
                with open(file_path, 'a'):
                    return True
            except IOError:
                # 文件被锁定，备份中
                time.sleep(5)
            except Exception:
                # 文件未创建，等待
                time.sleep(5)
        
        self.log(f"备份超时（{timeout}秒）")
        return False

    def start_backup(self):
        """启动手动备份（独立线程）"""
        if self.backup_running:
            messagebox.showwarning("提示", "当前已有备份任务在运行")
            return
            
        self.backup_running = True
        self.backup_button.config(text="备份中...", command=None)
        
        def backup_task():
            try:
                databases = self.get_selected_databases()
                if not databases:
                    self.log("未选择任何数据库")
                    messagebox.showwarning("警告", "请先选择要备份的数据库")
                    return
                    
                total = len(databases)
                success_count = 0
                fail_count = 0
                fail_list = []
                
                self.log(f"开始批量备份：共 {total} 个数据库")
                
                # 逐个备份
                for i, db in enumerate(databases, 1):
                    self.log(f"\n===== 备份进度: {i}/{total} - {db} =====")
                    result = self.backup_single_database(db, is_auto=False)
                    if result:
                        success_count += 1
                    else:
                        fail_count += 1
                        fail_list.append(db)
                
                # 备份后清理过期文件
                deleted, kept = self.delete_old_backups()
                
                # 显示总结
                summary = f"""
备份完成！
总数据库数：{total}
成功：{success_count} 个
失败：{fail_count} 个
"""
                if fail_list:
                    summary += f"失败列表：{', '.join(fail_list)}\n"
                if self.auto_cleanup_var.get():
                    summary += f"自动清理：删除 {deleted} 个过期备份"
                
                self.log(summary)
                
                # 显示结果（仅失败时弹窗）
                if fail_count > 0:
                    messagebox.showwarning("备份完成（部分失败）", summary)
                else:
                    messagebox.showinfo("备份成功", summary)
                    self.show_tray_notification("备份成功", f"所有 {total} 个数据库备份完成！")
                
            except Exception as e:
                error_msg = f"批量备份失败: {str(e)}"
                self.log(error_msg)
                messagebox.showerror("错误", error_msg)
            finally:
                self.backup_running = False
                self.backup_button.config(text="开始备份", command=self.start_backup)
        
        # 启动备份线程
        backup_thread = threading.Thread(target=backup_task, daemon=True)
        backup_thread.start()

    def add_backup_record(self, db_name, timestamp, status, size):
        """添加备份记录到监控表格"""
        try:
            # 格式化时间
            backup_time = datetime.datetime.strptime(timestamp, "%Y%m%d_%H%M%S").strftime("%Y-%m-%d %H:%M:%S")
        except:
            backup_time = timestamp
            
        # 插入记录（保持最新10条）
        self.recent_backups.insert(0, (db_name, backup_time, status, size))
        if len(self.recent_backups) > 10:
            self.recent_backups = self.recent_backups[:10]
            
        # 更新表格
        for item in self.recent_tree.get_children():
            self.recent_tree.delete(item)
            
        # 添加新记录
        for record in self.recent_backups:
            tag = "success" if record[2] == "成功" else "fail"
            self.recent_tree.insert("", tk.END, values=record, tags=(tag,))
        
        # 设置颜色
        self.recent_tree.tag_configure("success", foreground="green")
        self.recent_tree.tag_configure("fail", foreground="red")

    def update_daily_stats(self, success=True):
        """更新今日备份统计"""
        today = datetime.date.today()
        if not hasattr(self, 'stats_date') or self.stats_date != today:
            # 新的一天，重置统计
            self.today_success_count = 0
            self.today_fail_count = 0
            self.stats_date = today
            
        if success:
            self.today_success_count += 1
        else:
            self.today_fail_count += 1
            
        # 更新显示
        self.success_count_label.config(text=str(self.today_success_count))
        self.fail_count_label.config(text=str(self.today_fail_count))

    # ==============================
    # 日志相关方法
    # ==============================
    def clear_log(self):
        """清空当前日志显示"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.log("日志已清空")

    def export_log(self):
        """导出日志到文件"""
        try:
            # 获取日志内容
            self.log_text.config(state=tk.NORMAL)
            log_content = self.log_text.get(1.0, tk.END)
            self.log_text.config(state=tk.DISABLED)
            
            # 选择保存路径
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"backup_log_export_{timestamp}.txt"
            file_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
                initialfile=default_filename
            )
            
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(log_content)
                self.log(f"日志已导出到: {file_path}")
                messagebox.showinfo("成功", f"日志已导出到:\n{file_path}")
        except Exception as e:
            error_msg = f"导出日志失败: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)

    def load_history_log(self):
        """加载历史日志（最近100条）"""
        try:
            log_path = 'backup_log_v001.txt'
            if not os.path.exists(log_path):
                self.log("历史日志文件不存在")
                return
                
            with open(log_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()[-100:]  # 只加载最后100条
                
                self.log_text.config(state=tk.NORMAL)
                for line in lines:
                    self.log_text.insert(tk.END, line)
                self.log_text.see(tk.END)
                self.log_text.config(state=tk.DISABLED)
                
            self.log(f"已加载 {len(lines)} 条历史日志")
        except Exception as e:
            self.log(f"加载历史日志失败: {str(e)}")

    # ==============================
    # 定时任务相关方法
    # ==============================
    def start_schedule_thread(self):
        """启动定时任务线程"""
        def run_schedule():
            while True:
                schedule.run_pending()
                time.sleep(1)
                
        schedule_thread = threading.Thread(target=run_schedule, daemon=True)
        schedule_thread.start()
        self.log("定时任务线程已启动")

    def start_monitor_thread(self):
        """启动监控更新线程（检查托盘状态、更新显示）"""
        def monitor():
            while True:
                # 检查托盘状态，自动修复
                if not self.tray_restarting and not self.tray_healthy and self.startup_complete:
                    self.log("检测到托盘异常，自动重启...")
                    self.init_system_tray()
                
                # 更新窗口状态显示
                if self.root.state() == 'iconic' and not self.is_minimized_to_tray:
                    self.window_status_label.config(text="最小化", foreground="orange")
                elif self.is_minimized_to_tray:
                    self.window_status_label.config(text="已最小化到托盘", foreground="blue")
                else:
                    self.window_status_label.config(text="正常", foreground="green")
                    
                time.sleep(5)  # 每5秒检查一次
                
        monitor_thread = threading.Thread(target=monitor, daemon=True)
        monitor_thread.start()
        self.log("监控线程已启动")

if __name__ == "__main__":
    # 检查必要依赖库
    required_libs = {
        'pyodbc': 'pyodbc',
        'requests': 'requests',
        'schedule': 'schedule',
        'PIL': 'Pillow',
        'pystray': 'pystray'
    }
    
    missing_libs = []
    for lib, pkg in required_libs.items():
        try:
            __import__(lib)
        except ImportError:
            missing_libs.append(pkg)
    
    if missing_libs:
        print("=" * 60)
        print("请先安装以下必要依赖库：")
        print(f"pip install {' '.join(missing_libs)}")
        print("\n如需更好的Windows通知支持，额外安装：")
        print("pip install win10toast")
        print("=" * 60)
        sys.exit(1)
        
    # 启动程序
    root = tk.Tk()
    app = MSSQLBackupTool(root)
    root.mainloop()
