import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pyodbc
import os
import datetime
import threading
import logging
import platform
import shutil
import time
import requests
import sys
import json
import signal
from urllib.parse import urljoin, urlparse
import schedule
from PIL import Image, ImageTk
import getpass
import winreg
import tempfile
import pystray
from pystray import MenuItem as item
import queue

# 确保中文显示正常
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]

class MSSQLBackupTool:
    def __init__(self, root, silent_mode=False):
        self.root = root
        self.silent_mode = silent_mode
        self.root.title("MSSQL数据库备份工具 - V0.45（滚动优化版）")
        self.root.geometry("1000x700")
        
        # 确保主线程退出时能正常关闭
        self.running = True
        
        # 窗口设置
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.bind("<Unmap>", self.on_minimize)  # 最小化事件
        
        # 核心变量初始化
        self.log_text = None
        self.backup_running = False
        self.auto_backup_enabled = False
        self.available_databases = []
        self.selected_databases = []
        self.auto_cleanup_var = tk.BooleanVar(value=True)
        self.retention_days = tk.IntVar(value=30)
        # 服务器清理相关变量
        self.server_auto_cleanup_var = tk.BooleanVar(value=True)
        self.server_retention_days = tk.IntVar(value=15)
        
        self.os_type = platform.system()
        self.db_loading = False
        self.autostart_var = tk.BooleanVar(value=self.is_autostart_enabled())
        self.minimize_to_tray = tk.BooleanVar(value=True)  # 默认最小化到托盘
        
        # 备份时间变量
        self.backup_hour = tk.StringVar(value="18")
        self.backup_minute = tk.StringVar(value="00")
        self.auto_backup_var = tk.BooleanVar(value=False)
        
        # 托盘相关变量
        self.tray_icon = None
        self.tray_initialized = False
        self.tray_queue = queue.Queue()
        self.tray_thread = None
        self.common_icon = None  # 用于统一图标的变量
        self.last_click_time = 0  # 用于检测双击的时间戳
        self.double_click_threshold = 0.3  # 双击时间阈值（秒）
        
        # 字体设置
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("SimHei", 10))
        self.style.configure("TButton", font=("SimHei", 10))
        self.style.configure("TEntry", font=("SimHei", 10))
        self.style.configure("TCombobox", font=("SimHei", 10))
        
        # 创建界面
        self.create_widgets()
        
        # 初始化日志（修复权限问题）
        self.setup_logging()
        
        # 设置窗口图标和托盘图标（使用打包的icon.ico）
        self.set_icons()
        
        # 窗口居中
        self.center_window()
        self.root.resizable(True, True)
        
        # 加载配置
        self.load_config()
        
        # 异步加载数据库列表
        self.start_async_database_load()
        
        # 启动定时任务线程
        self.start_schedule_thread()
        
        # 启动监控更新线程
        self.start_monitor_thread()
        
        # 注册信号处理函数
        self.register_signal_handlers()
        
        # 启动下次备份时间更新线程
        self.start_next_backup_update_thread()
        
        # 启动托盘消息处理线程
        self.start_tray_message_processor()
        
        # 如果是静默模式，最小化到托盘
        if self.silent_mode:
            self.log("程序以静默模式启动，最小化到系统托盘")
            self.minimize_to_system_tray()

    def register_signal_handlers(self):
        def handle_interrupt(signum, frame):
            self.log(f"收到中断信号 (信号 {signum})，正在优雅退出...")
            self.quit_application(force=True)
            
        signal.signal(signal.SIGINT, handle_interrupt)
        signal.signal(signal.SIGTERM, handle_interrupt)

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry('{}x{}+{}+{}'.format(width, height, x, y))

    def set_icons(self):
        """设置窗口图标和托盘图标，使用打包的icon.ico"""
        try:
            self.common_icon = None  # 统一的图标对象
            self.temp_icon_path = None  # 临时图标路径
            
            # 尝试从程序资源中获取图标（适用于打包后的情况）
            try:
                # 对于PyInstaller打包的程序，使用_MEIPASS获取临时资源路径
                if getattr(sys, 'frozen', False):
                    base_path = sys._MEIPASS
                else:
                    base_path = os.path.dirname(os.path.abspath(__file__))
                
                self.icon_path = os.path.join(base_path, "icon.ico")
                
                if os.path.exists(self.icon_path) and os.path.isfile(self.icon_path):
                    # 加载打包的图标作为统一图标源
                    self.common_icon = Image.open(self.icon_path)
                    self.log(f"成功加载图标: {self.icon_path}")
                else:
                    raise FileNotFoundError(f"图标文件 {self.icon_path} 不存在")
                    
            except Exception as e:
                self.log(f"加载图标失败: {str(e)}")
                # 不再生成默认图标，仅使用系统默认图标
                self.common_icon = None
                
            # 为窗口和托盘设置图标
            if self.common_icon:
                # 创建临时ICO文件用于窗口图标
                try:
                    temp_icon = tempfile.NamedTemporaryFile(suffix='.ico', delete=False)
                    self.common_icon.save(temp_icon, format='ICO')
                    temp_icon.close()
                    self.temp_icon_path = temp_icon.name
                    
                    # 设置窗口图标
                    self.root.iconbitmap(self.temp_icon_path)
                    self.window_icon = ImageTk.PhotoImage(file=self.temp_icon_path)
                    self.root.iconphoto(True, self.window_icon)
                    self.log("窗口图标设置完成")
                except Exception as e:
                    self.log(f"设置窗口图标时出错: {str(e)}")
                
                # 设置托盘图标（直接使用PIL Image对象）
                self.tray_image = self.common_icon
                self.log("托盘图标设置完成")
            
            # 初始化系统托盘
            self.init_tray_icon()
                
        except Exception as e:
            self.log(f"设置图标过程出错: {str(e)}")

    def init_tray_icon(self):
        """使用pystray初始化系统托盘，使用打包的图标"""
        if not self.tray_image and self.common_icon is None:
            self.log("没有可用图标，使用系统默认图标初始化托盘")
            # 使用系统默认图标
            self.tray_image = None
        elif not self.tray_image:
            self.tray_image = self.common_icon
            
        try:
            # 创建托盘菜单
            menu = (
                item('显示窗口', self.tray_action_show_window),
                item('立即备份', self.tray_action_start_backup),
                item('查看日志', self.tray_action_show_log),
                item('退出', self.tray_action_quit)
            )
            
            # 创建托盘图标，添加双击事件处理
            self.tray_icon = pystray.Icon(
                "mssql_backup_tool",
                self.tray_image,  # 使用打包的图标
                "MSSQL数据库备份工具",
                menu
            )
            
            # 添加双击事件处理
            self.tray_icon.on_click = self.on_tray_click
            
            # 在单独线程中运行托盘
            self.tray_thread = threading.Thread(target=self.run_tray, daemon=True)
            self.tray_thread.start()
            
            self.log("系统托盘初始化成功")
            self.tray_initialized = True
            
        except Exception as e:
            self.log(f"初始化系统托盘失败: {str(e)}")
            self.tray_initialized = False

    def on_tray_click(self, icon, item, event):
        """处理托盘图标点击事件 - 增强版双击检测"""
        current_time = time.time()
        
        # 检测双击事件（两次点击时间间隔小于阈值）
        if current_time - self.last_click_time < self.double_click_threshold:
            self.log("检测到托盘图标双击事件")
            self.tray_action_show_window()
            # 重置点击时间
            self.last_click_time = 0
        else:
            # 记录第一次点击时间
            self.last_click_time = current_time
            
            # 处理单击事件（如果点击了菜单项）
            if item:
                item()

    def run_tray(self):
        """运行托盘图标"""
        try:
            if self.tray_icon:
                self.tray_icon.run()
        except Exception as e:
            self.log(f"托盘运行出错: {str(e)}")

    def start_tray_message_processor(self):
        """启动托盘消息处理线程"""
        def process_messages():
            while self.running:
                try:
                    # 等待消息
                    action = self.tray_queue.get(timeout=1)
                    # 执行相应操作
                    if action == "show_window":
                        self.root.after(0, self.show_window)
                    elif action == "start_backup":
                        self.root.after(0, self.start_backup_from_tray)
                    elif action == "manual_cleanup":
                        self.root.after(0, self.manual_cleanup)
                    elif action == "manual_server_cleanup":
                        self.root.after(0, self.manual_server_cleanup)
                    elif action == "show_log":
                        self.root.after(0, self.show_log)
                    elif action == "quit":
                        self.root.after(0, lambda: self.quit_application(force=True))
                    self.tray_queue.task_done()
                except queue.Empty:
                    continue
                except Exception as e:
                    self.log(f"处理托盘消息出错: {str(e)}")
        
        thread = threading.Thread(target=process_messages, daemon=True)
        thread.start()

    # 托盘动作 - 使用队列将操作发送到主线程
    def tray_action_show_window(self, icon=None, item=None):
        self.tray_queue.put("show_window")

    def tray_action_start_backup(self, icon=None, item=None):
        self.tray_queue.put("start_backup")

    def tray_action_manual_cleanup(self, icon=None, item=None):
        self.tray_queue.put("manual_cleanup")
        
    def tray_action_manual_server_cleanup(self, icon=None, item=None):
        self.tray_queue.put("manual_server_cleanup")

    def tray_action_show_log(self, icon=None, item=None):
        self.tray_queue.put("show_log")

    def tray_action_quit(self, icon=None, item=None):
        self.tray_queue.put("quit")

    def show_window(self):
        """从托盘显示窗口"""
        try:
            # 确保窗口被显示
            self.root.deiconify()
            self.root.lift()
            self.root.attributes('-topmost', True)
            self.root.focus_force()
            self.root.after(1000, lambda: self.root.attributes('-topmost', False))
            self.log("从系统托盘恢复窗口显示")
        except Exception as e:
            self.log(f"显示窗口时出错: {str(e)}")
            try:
                self.root.state('normal')
                self.root.focus_force()
                self.log("使用备选方案显示窗口")
            except Exception as e2:
                self.log(f"备选方案显示窗口也失败: {str(e2)}")

    def minimize_to_system_tray(self):
        """最小化到系统托盘"""
        if self.tray_initialized and self.minimize_to_tray.get():
            self.root.withdraw()  # 隐藏窗口
            self.show_tray_notification("程序已最小化", "双击托盘图标可显示主窗口")
            self.log("程序已最小化到系统托盘")
        else:
            # 如果托盘功能不可用，正常最小化
            self.root.iconify()

    def show_tray_notification(self, title, message):
        """显示托盘通知"""
        if self.tray_initialized and self.tray_icon:
            try:
                self.tray_icon.notify(message, title)
                self.log(f"显示托盘通知: {title} - {message}")
            except Exception as e:
                self.log(f"显示托盘通知失败: {str(e)}")

    def on_minimize(self, event):
        """窗口最小化事件处理"""
        if self.root.state() == "iconic" and self.minimize_to_tray.get():
            self.minimize_to_system_tray()

    def on_close(self):
        """窗口关闭事件处理"""
        if self.minimize_to_tray.get() and self.tray_initialized:
            self.minimize_to_system_tray()
        else:
            self.quit_application()

    def setup_logging(self):
        """初始化日志系统，修复权限问题，将日志保存在用户可访问的目录"""
        try:
            # 确定合适的日志目录
            if getattr(sys, 'frozen', False):
                # 打包后的程序
                base_dir = os.path.dirname(sys.executable)
            else:
                # 开发环境
                base_dir = os.path.dirname(os.path.abspath(__file__))
            
            # 测试目录是否可写
            log_dir = base_dir
            writeable = False
            try:
                # 创建临时文件测试写入权限
                test_file = os.path.join(log_dir, "test_write_permission.tmp")
                with open(test_file, "w", encoding="utf-8") as f:
                    f.write("test")
                os.remove(test_file)
                writeable = True
            except:
                writeable = False
            
            # 如果程序目录不可写，使用用户文档目录
            if not writeable:
                log_dir = os.path.join(os.path.expanduser("~"), "Documents", "MSSQLBackupTool")
                # 确保目录存在
                os.makedirs(log_dir, exist_ok=True)
            
            # 日志文件路径
            log_file = os.path.join(log_dir, 'backup_log_v0.45.txt')
            
            # 配置日志
            logging.basicConfig(
                filename=log_file,
                level=logging.INFO,
                format='%(asctime)s - %(levelname)s - %(message)s',
                encoding='utf-8'
            )
            
            # 记录日志文件路径，方便用户查找
            self.log(f"程序启动（版本V0.45 - 滚动优化版）")
            self.log(f"日志文件保存路径: {log_file}")
            
        except Exception as e:
            # 最后的日志错误处理
            print(f"初始化日志系统失败: {str(e)}")
            # 尝试在控制台显示关键信息
            self.log_text = None  # 避免后续调用log方法出错

    def create_widgets(self):
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
        
        # 创建各标签页内容
        self.create_config_tab()
        self.create_log_tab()
        self.create_monitor_tab()

    def _on_listbox_scroll(self, event, listbox):
        """列表框滚动事件处理函数"""
        if event.delta > 0:
            # 向上滚动
            listbox.yview_scroll(-1, "units")
        else:
            # 向下滚动
            listbox.yview_scroll(1, "units")

    def create_config_tab(self):
        # 创建带滚动条的配置页面
        container = ttk.Frame(self.config_tab)
        container.pack(fill=tk.BOTH, expand=True)
        
        # 创建画布用于滚动
        canvas = tk.Canvas(container)
        
        # 垂直滚动条
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        # 当滚动框内容变化时更新滚动区域
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        
        # 在画布上创建滚动框
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 为配置页面添加鼠标滚轮事件，控制整体滚动
        def _on_canvas_scroll(event):
            if event.delta:
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            else:
                if event.num == 4:
                    canvas.yview_scroll(-1, "units")
                elif event.num == 5:
                    canvas.yview_scroll(1, "units")
        
        canvas.bind_all("<MouseWheel>", _on_canvas_scroll)
        canvas.bind_all("<Button-4>", _on_canvas_scroll)
        canvas.bind_all("<Button-5>", _on_canvas_scroll)
        
        # 布局滚动条和画布
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        # 配置内容放在滚动框内，使用紧凑布局
        config_frame = ttk.Frame(scrollable_frame, padding="10")
        config_frame.pack(fill=tk.X, expand=True)
        
        # 数据库连接设置 - 调整为用户名和密码在同一行
        conn_frame = ttk.LabelFrame(config_frame, text="数据库连接设置", padding="8")
        conn_frame.pack(fill=tk.X, pady=4)
        
        # 服务器地址行
        ttk.Label(conn_frame, text="服务器地址:").grid(row=0, column=0, sticky=tk.W, pady=3, padx=3)
        self.server_entry = ttk.Entry(conn_frame, width=40)
        self.server_entry.grid(row=0, column=1, sticky=tk.W, pady=3, padx=3, columnspan=2)
        ttk.Label(conn_frame, text="服务器地址格式示例", font=("SimHei", 8)).grid(row=0, column=3, sticky=tk.W, pady=3)
        
        # 用户名和密码在同一行
        user_pass_frame = ttk.Frame(conn_frame)
        user_pass_frame.grid(row=1, column=0, columnspan=4, sticky=tk.W, pady=3)
        
        ttk.Label(user_pass_frame, text="用户名:").pack(side=tk.LEFT, padx=(3, 3))
        self.user_entry = ttk.Entry(user_pass_frame, width=20)
        self.user_entry.pack(side=tk.LEFT, padx=(0, 15))
        
        ttk.Label(user_pass_frame, text="密码:").pack(side=tk.LEFT, padx=(3, 3))
        self.password_entry = ttk.Entry(user_pass_frame, width=20, show="*")
        self.password_entry.pack(side=tk.LEFT, padx=(0, 15))
        
        self.show_password_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            user_pass_frame, 
            text="显示密码", 
            variable=self.show_password_var,
            command=self.toggle_password_visibility
        ).pack(side=tk.LEFT)
        
        # 数据库选择区域 - 紧凑布局，刷新按钮在添加按钮上方
        db_selection_frame = ttk.LabelFrame(config_frame, text="数据库选择", padding="8")
        db_selection_frame.pack(fill=tk.X, pady=4)
        
        # 左侧：可用数据库列表
        available_frame = ttk.LabelFrame(db_selection_frame, text="可用数据库", padding="4")
        available_frame.pack(side=tk.LEFT, padx=4, fill=tk.BOTH, expand=True)
        
        self.available_listbox = tk.Listbox(available_frame, selectmode=tk.EXTENDED, width=25, height=6)
        self.available_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 为可用数据库列表添加鼠标滚轮支持
        self.available_listbox.bind(
            "<MouseWheel>", 
            lambda e: self._on_listbox_scroll(e, self.available_listbox)
        )
        self.available_listbox.bind(
            "<Button-4>", 
            lambda e: self._on_listbox_scroll(e, self.available_listbox)
        )
        self.available_listbox.bind(
            "<Button-5>", 
            lambda e: self._on_listbox_scroll(e, self.available_listbox)
        )
        
        available_scrollbar = ttk.Scrollbar(available_frame, orient=tk.VERTICAL, command=self.available_listbox.yview)
        available_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.available_listbox.config(yscrollcommand=available_scrollbar.set)
        
        # 中间：操作按钮 - 刷新按钮在最上方
        buttons_frame = ttk.Frame(db_selection_frame, padding="4")
        buttons_frame.pack(side=tk.LEFT, padx=4, fill=tk.Y)
        
        # 刷新数据库列表按钮在最上方
        ttk.Button(buttons_frame, text="刷新数据库列表", command=self.start_async_database_load).pack(fill=tk.X, pady=1)
        ttk.Button(buttons_frame, text="添加 >", command=self.add_selected_databases).pack(fill=tk.X, pady=1)
        ttk.Button(buttons_frame, text="< 移除", command=self.remove_selected_databases).pack(fill=tk.X, pady=1)
        ttk.Button(buttons_frame, text="全部添加", command=self.add_all_databases).pack(fill=tk.X, pady=1)
        ttk.Button(buttons_frame, text="全部移除", command=self.remove_all_databases).pack(fill=tk.X, pady=1)
        
        # 右侧：已选备份数据库列表
        selected_frame = ttk.LabelFrame(db_selection_frame, text="已选备份数据库", padding="4")
        selected_frame.pack(side=tk.LEFT, padx=4, fill=tk.BOTH, expand=True)
        
        self.selected_listbox = tk.Listbox(selected_frame, selectmode=tk.EXTENDED, width=25, height=6)
        self.selected_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 为已选数据库列表添加鼠标滚轮支持
        self.selected_listbox.bind(
            "<MouseWheel>", 
            lambda e: self._on_listbox_scroll(e, self.selected_listbox)
        )
        self.selected_listbox.bind(
            "<Button-4>", 
            lambda e: self._on_listbox_scroll(e, self.selected_listbox)
        )
        self.selected_listbox.bind(
            "<Button-5>", 
            lambda e: self._on_listbox_scroll(e, self.selected_listbox)
        )
        
        selected_scrollbar = ttk.Scrollbar(selected_frame, orient=tk.VERTICAL, command=self.selected_listbox.yview)
        selected_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.selected_listbox.config(yscrollcommand=selected_scrollbar.set)
        
        # 数据库加载状态指示器
        db_status_frame = ttk.Frame(db_selection_frame)
        db_status_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=4)
        
        self.db_status_indicator = ttk.Label(db_status_frame, text="准备加载数据库...", font=("SimHei", 9))
        self.db_status_indicator.pack(side=tk.LEFT, padx=3)
        
        # 备份设置 - 仅保留服务器备份模式
        backup_frame = ttk.LabelFrame(config_frame, text="备份设置（备份流程: 先备份到服务器，再下载到本地）", padding="8")
        backup_frame.pack(fill=tk.X, pady=4)
        
        # 移除备份模式选择，直接使用"先服务器再下载"模式
        
        ttk.Label(backup_frame, text="服务器IIS路径:").grid(row=1, column=0, sticky=tk.W, pady=3, padx=3)
        self.server_temp_path_entry = ttk.Entry(backup_frame, width=35)
        self.server_temp_path_entry.grid(row=1, column=1, sticky=tk.W, pady=3, padx=3)
        ttk.Label(backup_frame, text="服务器上存储备份的路径", font=("SimHei", 8)).grid(row=1, column=2, sticky=tk.W, pady=3)
        
        ttk.Label(backup_frame, text="服务器Web地址:").grid(row=2, column=0, sticky=tk.W, pady=3, padx=3)
        self.server_web_url_entry = ttk.Entry(backup_frame, width=35)
        self.server_web_url_entry.grid(row=2, column=1, sticky=tk.W, pady=3, padx=3)
        ttk.Label(backup_frame, text="用于下载备份的HTTP地址", font=("SimHei", 8)).grid(row=2, column=2, sticky=tk.W, pady=3)
        
        path_frame = ttk.Frame(backup_frame)
        path_frame.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=3, padx=3)
        
        ttk.Label(path_frame, text="本地保存路径:").pack(side=tk.LEFT, padx=3, pady=3)
        self.local_save_path_entry = ttk.Entry(path_frame, width=35)
        self.local_save_path_entry.pack(side=tk.LEFT, padx=3, pady=3, fill=tk.X, expand=True)
        ttk.Button(path_frame, text="浏览", command=self.browse_local_path).pack(side=tk.LEFT, padx=3, pady=3)
        
        ttk.Label(backup_frame, text="文件名前缀:").grid(row=4, column=0, sticky=tk.W, pady=3, padx=3)
        self.filename_prefix_entry = ttk.Entry(backup_frame, width=35)
        self.filename_prefix_entry.grid(row=4, column=1, sticky=tk.W, pady=3, padx=3)
        ttk.Label(backup_frame, text="用于区分不同备份的前缀", font=("SimHei", 8)).grid(row=4, column=2, sticky=tk.W, pady=3)
        
        # 自动清理设置 - 紧凑布局
        cleanup_frame = ttk.LabelFrame(config_frame, text="自动清理设置", padding="8")
        cleanup_frame.pack(fill=tk.X, pady=4)
        
        # 本地清理设置
        local_cleanup_frame = ttk.LabelFrame(cleanup_frame, text="本地备份清理", padding="4")
        local_cleanup_frame.pack(fill=tk.X, pady=3)
        
        local_cleanup_grid = ttk.Frame(local_cleanup_frame)
        local_cleanup_grid.pack(fill=tk.X, expand=True)
        
        ttk.Checkbutton(
            local_cleanup_grid, 
            text="启用自动删除过期本地备份", 
            variable=self.auto_cleanup_var
        ).grid(row=0, column=0, sticky=tk.W, pady=2, padx=2)
        
        ttk.Label(local_cleanup_grid, text="保留天数:").grid(row=0, column=1, sticky=tk.W, pady=2, padx=2)
        ttk.Combobox(
            local_cleanup_grid, 
            textvariable=self.retention_days,
            values=[7, 15, 30, 60, 90, 180, 365],
            width=5
        ).grid(row=0, column=2, sticky=tk.W, pady=2, padx=2)
        
        ttk.Button(local_cleanup_grid, text="立即清理", command=self.manual_cleanup).grid(
            row=0, column=3, padx=5, pady=2, sticky=tk.E)
        
        local_cleanup_grid.columnconfigure(4, weight=1)
        
        # 服务器清理设置
        server_cleanup_frame = ttk.LabelFrame(cleanup_frame, text="服务器备份清理", padding="4")
        server_cleanup_frame.pack(fill=tk.X, pady=3)
        
        server_cleanup_grid = ttk.Frame(server_cleanup_frame)
        server_cleanup_grid.pack(fill=tk.X, expand=True)
        
        ttk.Checkbutton(
            server_cleanup_grid, 
            text="启用自动删除过期服务器备份", 
            variable=self.server_auto_cleanup_var
        ).grid(row=0, column=0, sticky=tk.W, pady=2, padx=2)
        
        ttk.Label(server_cleanup_grid, text="保留天数:").grid(row=0, column=1, sticky=tk.W, pady=2, padx=2)
        ttk.Combobox(
            server_cleanup_grid, 
            textvariable=self.server_retention_days,
            values=[7, 15, 30, 60],
            width=5
        ).grid(row=0, column=2, sticky=tk.W, pady=2, padx=2)
        
        ttk.Button(server_cleanup_grid, text="立即清理", command=self.manual_server_cleanup).grid(
            row=0, column=3, padx=5, pady=2, sticky=tk.E)
        
        server_cleanup_grid.columnconfigure(4, weight=1)
        
        # 系统设置 - 调整为最小化到托盘和随Windows启动在同一行
        system_frame = ttk.LabelFrame(config_frame, text="系统设置", padding="8")
        system_frame.pack(fill=tk.X, pady=4)
        
        # 选项在同一行显示
        system_options_frame = ttk.Frame(system_frame)
        system_options_frame.pack(fill=tk.X, pady=2)
        
        # 最小化到系统托盘选项
        ttk.Checkbutton(
            system_options_frame, 
            text="最小化到系统托盘", 
            variable=self.minimize_to_tray,
            command=self.log_minimize_setting
        ).pack(side=tk.LEFT, padx=(0, 20))
        
        # 随Windows启动选项（原开机自动启动）
        ttk.Checkbutton(
            system_options_frame, 
            text="随Windows启动", 
            variable=self.autostart_var,
            command=self.toggle_autostart
        ).pack(side=tk.LEFT)
        
        # 自动备份设置 - 紧凑布局
        auto_backup_frame = ttk.LabelFrame(config_frame, text="自动备份设置", padding="8")
        auto_backup_frame.pack(fill=tk.X, pady=4)
        
        auto_backup_grid = ttk.Frame(auto_backup_frame)
        auto_backup_grid.pack(fill=tk.X, expand=True)
        
        # 自动备份开关
        auto_backup_switch = ttk.Checkbutton(
            auto_backup_grid, 
            text="启用每天自动备份", 
            variable=self.auto_backup_var,
            command=self.toggle_auto_backup
        )
        auto_backup_switch.grid(row=0, column=0, sticky=tk.W, pady=3, padx=3)
        
        # 时间选择区域
        time_frame = ttk.Frame(auto_backup_grid)
        time_frame.grid(row=0, column=1, sticky=tk.W, pady=3)
        
        ttk.Label(time_frame, text="备份时间:").pack(side=tk.LEFT, padx=2)
        
        self.hour_combobox = ttk.Combobox(
            time_frame, 
            textvariable=self.backup_hour,
            values=[f"{i:02d}" for i in range(24)],
            width=5,
            state="disabled"  # 默认禁用
        )
        self.hour_combobox.pack(side=tk.LEFT, padx=1)
        
        ttk.Label(time_frame, text=":").pack(side=tk.LEFT)
        
        self.minute_combobox = ttk.Combobox(
            time_frame, 
            textvariable=self.backup_minute,
            values=[f"{i:02d}" for i in range(0, 60, 5)],
            width=5,
            state="disabled"  # 默认禁用
        )
        self.minute_combobox.pack(side=tk.LEFT, padx=1)
        
        # 下次备份时间显示
        self.next_backup_label = ttk.Label(auto_backup_grid, text="下次备份: 未设置", font=("SimHei", 9))
        self.next_backup_label.grid(row=0, column=2, sticky=tk.W, pady=3, padx=10)
        
        # 绑定时间变化事件，实时更新下次备份时间
        self.backup_hour.trace_add("write", lambda *args: self.update_next_backup_time())
        self.backup_minute.trace_add("write", lambda *args: self.update_next_backup_time())
        
        # 按钮区域 - 紧凑排列
        button_frame = ttk.Frame(config_frame, padding="8")
        button_frame.pack(fill=tk.X, pady=4)
        
        self.backup_button = ttk.Button(button_frame, text="开始备份", command=self.start_backup)
        self.backup_button.pack(side=tk.LEFT, padx=3)
        
        # 测试按钮
        ttk.Button(button_frame, text="测试Web连接", command=self.test_web_connection).pack(side=tk.LEFT, padx=3)
        ttk.Button(button_frame, text="保存配置", command=self.save_config).pack(side=tk.LEFT, padx=3)
        ttk.Button(button_frame, text="测试数据库连接", command=self.test_connection).pack(side=tk.LEFT, padx=3)
        ttk.Button(button_frame, text="退出", command=lambda: self.quit_application(force=True)).pack(side=tk.RIGHT, padx=3)

    def create_log_tab(self):
        log_frame = ttk.Frame(self.log_tab, padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        log_control_frame = ttk.Frame(log_frame)
        log_control_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(log_control_frame, text="清空日志", command=self.clear_log).pack(side=tk.RIGHT, padx=5)
        ttk.Button(log_control_frame, text="导出日志", command=self.export_log).pack(side=tk.RIGHT, padx=5)
        
        log_display_frame = ttk.LabelFrame(log_frame, text="操作日志（版本V0.45 - 滚动优化版）", padding="10")
        log_display_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = tk.Text(log_display_frame, height=25, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)
        
        scrollbar = ttk.Scrollbar(log_display_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 为日志文本框添加鼠标滚轮支持
        self.log_text.bind("<MouseWheel>", lambda e: self.log_text.yview_scroll(int(-1*(e.delta/120)), "units"))
        self.log_text.bind("<Button-4>", lambda e: self.log_text.yview_scroll(-1, "units"))
        self.log_text.bind("<Button-5>", lambda e: self.log_text.yview_scroll(1, "units"))
        
        self.load_history_log()

    def create_monitor_tab(self):
        monitor_frame = ttk.Frame(self.monitor_tab, padding="10")
        monitor_frame.pack(fill=tk.BOTH, expand=True)
        
        # 系统状态区域 - 分为两列显示
        status_frame = ttk.LabelFrame(monitor_frame, text="系统状态（版本V0.45 - 滚动优化版）", padding="10")
        status_frame.pack(fill=tk.X, pady=5)
        
        # 创建两列布局
        status_grid = ttk.Frame(status_frame)
        status_grid.pack(fill=tk.X, expand=True)
        
        # 第一列
        col1 = ttk.Frame(status_grid)
        col1.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        
        # 第二列
        col2 = ttk.Frame(status_grid)
        col2.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        
        # 第一列内容
        ttk.Label(col1, text="程序启动时间:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        self.start_time_label = ttk.Label(col1, text=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.start_time_label.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(col1, text="数据库连接状态:").grid(row=1, column=0, sticky=tk.W, pady=5, padx=5)
        self.db_status_label = ttk.Label(col1, text="未检测", foreground="orange")
        self.db_status_label.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(col1, text="自动备份状态:").grid(row=2, column=0, sticky=tk.W, pady=5, padx=5)
        self.auto_backup_status_label = ttk.Label(col1, text="未启用", foreground="orange")
        self.auto_backup_status_label.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # 第二列内容
        ttk.Label(col2, text="下次自动备份时间:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        self.monitor_next_backup_label = ttk.Label(col2, text="未设置", foreground="orange")
        self.monitor_next_backup_label.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(col2, text="服务器清理状态:").grid(row=1, column=0, sticky=tk.W, pady=5, padx=5)
        self.server_cleanup_status_label = ttk.Label(col2, text="未启用", foreground="orange")
        self.server_cleanup_status_label.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(col2, text="已选备份数据库数:").grid(row=2, column=0, sticky=tk.W, pady=5, padx=5)
        self.selected_count_label = ttk.Label(col2, text="0")
        self.selected_count_label.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # 更新开机启动状态标签文本
        ttk.Label(col2, text="随Windows启动状态:").grid(row=3, column=0, sticky=tk.W, pady=5, padx=5)
        self.autostart_status_label = ttk.Label(col2, text="未启用", foreground="orange")
        self.autostart_status_label.grid(row=3, column=1, sticky=tk.W, pady=5)
        
        # 备份统计和清理统计放到同一行显示
        stats_container = ttk.Frame(monitor_frame)
        stats_container.pack(fill=tk.X, pady=5)
        
        # 备份统计区域（左半部分）
        stats_frame = ttk.LabelFrame(stats_container, text="备份统计", padding="10")
        stats_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        stats_grid = ttk.Frame(stats_frame)
        stats_grid.pack(fill=tk.X, expand=True)
        
        ttk.Label(stats_grid, text="今日成功备份数:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        self.success_count_label = ttk.Label(stats_grid, text="0")
        self.success_count_label.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(stats_grid, text="今日失败备份数:").grid(row=0, column=2, sticky=tk.W, pady=5, padx=5)
        self.fail_count_label = ttk.Label(stats_grid, text="0")
        self.fail_count_label.grid(row=0, column=3, sticky=tk.W, pady=5)
        
        # 清理统计区域（右半部分）
        cleanup_stats_frame = ttk.LabelFrame(stats_container, text="清理统计", padding="10")
        cleanup_stats_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        
        cleanup_grid = ttk.Frame(cleanup_stats_frame)
        cleanup_grid.pack(fill=tk.X, expand=True)
        
        ttk.Label(cleanup_grid, text="今日清理本地备份数:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        self.local_cleaned_count_label = ttk.Label(cleanup_grid, text="0")
        self.local_cleaned_count_label.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(cleanup_grid, text="今日清理服务器备份数:").grid(row=0, column=2, sticky=tk.W, pady=5, padx=5)
        self.server_cleaned_count_label = ttk.Label(cleanup_grid, text="0")
        self.server_cleaned_count_label.grid(row=0, column=3, sticky=tk.W, pady=5)
        
        # 最近备份记录区域
        recent_frame = ttk.LabelFrame(monitor_frame, text="最近备份记录", padding="10")
        recent_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        columns = ("数据库名", "备份时间", "状态", "大小")
        self.recent_tree = ttk.Treeview(recent_frame, columns=columns, show="headings")
        
        for col in columns:
            self.recent_tree.heading(col, text=col)
            self.recent_tree.column(col, width=150)
        
        self.recent_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(recent_frame, orient=tk.VERTICAL, command=self.recent_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.recent_tree.configure(yscrollcommand=scrollbar.set)
        
        # 为最近备份记录添加鼠标滚轮支持
        self.recent_tree.bind("<MouseWheel>", lambda e: self.recent_tree.yview_scroll(int(-1*(e.delta/120)), "units"))
        self.recent_tree.bind("<Button-4>", lambda e: self.recent_tree.yview_scroll(-1, "units"))
        self.recent_tree.bind("<Button-5>", lambda e: self.recent_tree.yview_scroll(1, "units"))
        
        self.today_success_count = 0
        self.today_fail_count = 0
        self.today_local_cleaned_count = 0
        self.today_server_cleaned_count = 0
        self.recent_backups = []
        
        # 启动状态更新
        self.update_status()

    def update_status(self):
        """更新状态显示"""
        # 更新服务器清理状态
        if self.server_auto_cleanup_var.get():
            self.server_cleanup_status_label.config(
                text=f"已启用（保留{self.server_retention_days.get()}天）", 
                foreground="green"
            )
        else:
            self.server_cleanup_status_label.config(text="未启用", foreground="orange")
            
        # 更新清理统计
        self.local_cleaned_count_label.config(text=str(self.today_local_cleaned_count))
        self.server_cleaned_count_label.config(text=str(self.today_server_cleaned_count))
            
        # 定期检查并更新状态
        self.root.after(5000, self.update_status)

    def start_next_backup_update_thread(self):
        """启动线程定期更新下次备份时间显示"""
        def update_next_backup_periodically():
            while self.running:
                if self.auto_backup_enabled:
                    self.update_next_backup_time()
                time.sleep(60)
                
        update_thread = threading.Thread(target=update_next_backup_periodically, daemon=True)
        update_thread.start()
        self.log("下次备份时间更新线程已启动")

    def calculate_next_backup_time(self):
        """精确计算下次备份时间"""
        if not self.auto_backup_enabled:
            return None
            
        try:
            now = datetime.datetime.now()
            hour = int(self.backup_hour.get())
            minute = int(self.backup_minute.get())
            
            # 设置下次备份时间为今天的指定时间
            next_backup = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
            
            # 如果今天的指定时间已过，则设置为明天的同一时间
            if next_backup <= now:
                next_backup += datetime.timedelta(days=1)
                
            return next_backup
        except Exception as e:
            self.log(f"计算下次备份时间出错: {str(e)}")
            return None

    def update_next_backup_time(self):
        """更新下次备份时间显示（在配置页和监控页同时显示）"""
        next_backup = self.calculate_next_backup_time()
        if next_backup:
            time_str = next_backup.strftime("%Y-%m-%d %H:%M:%S")
            # 在配置页显示
            self.next_backup_label.config(text=f"下次备份: {time_str}")
            # 在监控页显示
            self.monitor_next_backup_label.config(text=time_str, foreground="green")
            return time_str
        else:
            # 未启用自动备份时的显示
            self.next_backup_label.config(text="下次备份: 未设置")
            self.monitor_next_backup_label.config(text="未设置", foreground="orange")
        return None

    def log(self, message):
        if self.log_text is None:
            print(f"日志: {message}")
            return
            
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        logging.info(message)

    def log_minimize_setting(self):
        """记录最小化设置的变更"""
        if self.minimize_to_tray.get():
            self.log("已启用最小化到系统托盘")
        else:
            self.log("已禁用最小化到系统托盘，将正常最小化到任务栏")

    def clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.log("日志已清空")

    def export_log(self):
        try:
            # 获取实际的日志文件路径
            if hasattr(logging, 'getLogger'):
                root_logger = logging.getLogger()
                for handler in root_logger.handlers:
                    if isinstance(handler, logging.FileHandler):
                        log_file_path = handler.baseFilename
                        break
                else:
                    # 如果找不到日志文件路径，使用默认路径
                    log_file_path = os.path.join(os.path.expanduser("~"), "Documents", "MSSQLBackupTool", 'backup_log_v0.45.txt')
            else:
                log_file_path = os.path.join(os.path.expanduser("~"), "Documents", "MSSQLBackupTool", 'backup_log_v0.45.txt')
            
            # 尝试直接读取日志文件
            try:
                with open(log_file_path, 'r', encoding='utf-8') as f:
                    log_content = f.read()
            except:
                # 备选方案：从UI控件获取日志内容
                self.log_text.config(state=tk.NORMAL)
                log_content = self.log_text.get(1.0, tk.END)
                self.log_text.config(state=tk.DISABLED)
            
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"backup_log_export_{timestamp}_V0.45.txt"
            file_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
                initialfile=default_filename
            )
            
            if file_path:
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(log_content)
                self.log(f"日志已导出到: {file_path}")
                messagebox.showinfo("成功", f"日志已导出到:\n{file_path}")
                
        except Exception as e:
            error_msg = f"导出日志失败: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)

    def load_history_log(self):
        try:
            # 获取实际的日志文件路径
            if hasattr(logging, 'getLogger'):
                root_logger = logging.getLogger()
                for handler in root_logger.handlers:
                    if isinstance(handler, logging.FileHandler):
                        log_file = handler.baseFilename
                        break
                else:
                    # 如果找不到日志文件路径，使用默认路径
                    log_file = os.path.join(os.path.expanduser("~"), "Documents", "MSSQLBackupTool", 'backup_log_v0.45.txt')
            else:
                log_file = os.path.join(os.path.expanduser("~"), "Documents", "MSSQLBackupTool", 'backup_log_v0.45.txt')
            
            if os.path.exists(log_file):
                with open(log_file, 'r', encoding='utf-8') as f:
                    lines = f.readlines()[-100:]  # 只加载最后100行
                    
                    self.log_text.config(state=tk.NORMAL)
                    for line in lines:
                        self.log_text.insert(tk.END, line)
                    self.log_text.see(tk.END)
                    self.log_text.config(state=tk.DISABLED)
                self.log(f"已加载最近的日志记录（版本V0.45），日志文件路径：{log_file}")
        except Exception as e:
            self.log(f"加载历史日志失败: {str(e)}")

    def browse_local_path(self):
        path = filedialog.askdirectory(title="选择本地保存路径")
        if path:
            self.local_save_path_entry.delete(0, tk.END)
            self.local_save_path_entry.insert(0, path)

    def get_connection_string(self, database="master", timeout=10):
        server = self.server_entry.get().strip()
        user = self.user_entry.get().strip()
        password = self.password_entry.get().strip()
        
        if not server:
            messagebox.showinfo("提示", "请输入服务器地址")
            return None
            
        # 尝试多种常见的ODBC驱动名称，提高兼容性
        drivers = [
            "ODBC Driver 18 for SQL Server",
            "ODBC Driver 17 for SQL Server",
            "SQL Server Native Client 11.0",
            "SQL Server"
        ]
        
        # 先尝试用户可能已安装的驱动
        for driver in drivers:
            try:
                if user:
                    conn_str = f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};UID={user};PWD={password};AutoCommit=True;Connection Timeout={timeout}"
                else:
                    conn_str = f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};Trusted_Connection=yes;AutoCommit=True;Connection Timeout={timeout}"
                
                # 测试连接字符串格式是否有效
                pyodbc.connect(conn_str)
                return conn_str
            except:
                continue
                
        # 如果所有驱动都尝试失败，返回默认格式让用户手动修改
        if user:
            return f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={user};PWD={password};AutoCommit=True;Connection Timeout={timeout}"
        else:
            return f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes;AutoCommit=True;Connection Timeout={timeout}"

    def test_connection(self):
        def connection_task():
            try:
                self.log("正在测试数据库连接...")
                conn_str = self.get_connection_string(timeout=10)
                if not conn_str:
                    return
                    
                conn = pyodbc.connect(conn_str)
                conn.close()
                self.log("数据库连接成功!")
                self.root.after(0, lambda: messagebox.showinfo("成功", "数据库连接成功!"))
                self.root.after(0, lambda: self.db_status_label.config(text="连接正常", foreground="green"))
            except Exception as e:
                error_msg = f"数据库连接失败: {str(e)}\n请检查是否已安装ODBC驱动"
                self.log(error_msg)
                self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
                self.root.after(0, lambda: self.db_status_label.config(text="连接失败", foreground="red"))
            finally:
                self.root.after(0, lambda: self.db_status_indicator.config(text="数据库连接测试完成"))
        
        self.db_status_indicator.config(text="正在测试数据库连接...")
        threading.Thread(target=connection_task, daemon=True).start()

    def test_web_connection(self):
        try:
            self.log("正在测试Web服务器连接...")
            web_url = self.server_web_url_entry.get().strip()
            if not web_url:
                self.log("请输入服务器Web访问URL")
                return
                
            response = requests.head(web_url, timeout=10)
            if response.status_code in [200, 403]:
                self.log("Web服务器连接成功!")
                messagebox.showinfo("成功", "Web服务器连接成功!")
            else:
                self.log(f"Web服务器连接失败，状态码: {response.status_code}")
                messagebox.showerror("错误", f"Web服务器连接失败，状态码: {response.status_code}")
        except Exception as e:
            error_msg = f"Web服务器连接失败: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)

    def refresh_databases(self):
        try:
            self.log("正在刷新数据库列表...")
            conn_str = self.get_connection_string(timeout=10)
            if not conn_str:
                return
                
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            
            cursor.execute("SELECT name FROM sys.databases WHERE name NOT IN ('master', 'tempdb', 'model', 'msdb') ORDER BY name")
            self.available_databases = [row[0] for row in cursor.fetchall()]
            conn.close()
            
            self.root.after(0, self.update_available_listbox)
            
            existing_databases = []
            missing_databases = []
            for db in self.selected_databases:
                if db in self.available_databases:
                    existing_databases.append(db)
                else:
                    missing_databases.append(db)
            
            if missing_databases:
                self.selected_databases = existing_databases
                self.root.after(0, self.update_selected_listbox)
                self.log(f"已从备份列表中移除 {len(missing_databases)} 个不存在的数据库: {', '.join(missing_databases)}")
            
            status_msg = f"已加载 {len(self.available_databases)} 个可用数据库，备份列表中有 {len(self.selected_databases)} 个数据库"
            self.log(status_msg)
            self.root.after(0, lambda: self.db_status_label.config(text="连接正常", foreground="green"))
            self.root.after(0, lambda: self.db_status_indicator.config(text=status_msg))
            self.root.after(0, self.update_selected_count)
            
        except Exception as e:
            error_msg = f"刷新数据库列表失败: {str(e)}"
            self.log(error_msg)
            self.root.after(0, lambda: self.db_status_label.config(text="连接失败", foreground="red"))
            self.root.after(0, lambda: self.db_status_indicator.config(text="数据库加载失败，请检查连接"))
        finally:
            self.db_loading = False

    def start_async_database_load(self):
        if self.db_loading:
            self.log("数据库加载中，请稍候...")
            return
            
        self.db_loading = True
        self.db_status_indicator.config(text="正在加载数据库列表...（超时10秒）")
        self.log("开始异步加载数据库列表...")
        
        def async_load():
            self.refresh_databases()
            
        threading.Thread(target=async_load, daemon=True).start()

    def save_config(self):
        try:
            # 配置文件保存到日志所在目录，避免权限问题
            if hasattr(logging, 'getLogger'):
                root_logger = logging.getLogger()
                for handler in root_logger.handlers:
                    if isinstance(handler, logging.FileHandler):
                        log_dir = os.path.dirname(handler.baseFilename)
                        break
                else:
                    log_dir = os.path.join(os.path.expanduser("~"), "Documents", "MSSQLBackupTool")
            else:
                log_dir = os.path.join(os.path.expanduser("~"), "Documents", "MSSQLBackupTool")
            
            # 确保配置目录存在
            os.makedirs(log_dir, exist_ok=True)
            
            config = {
                'version': 'V0.45',
                'server': self.server_entry.get(),
                'user': self.user_entry.get(),
                'password': self.password_entry.get(),
                'server_temp_path': self.server_temp_path_entry.get(),
                'server_web_url': self.server_web_url_entry.get(),
                'local_save_path': self.local_save_path_entry.get(),
                'filename_prefix': self.filename_prefix_entry.get(),
                # 固定使用服务器模式
                'backup_mode': 'server_then_web',
                'auto_backup_enabled': self.auto_backup_var.get(),
                'backup_hour': self.backup_hour.get(),
                'backup_minute': self.backup_minute.get(),
                'auto_cleanup_var': self.auto_cleanup_var.get(),
                'retention_days': self.retention_days.get(),
                'server_auto_cleanup_var': self.server_auto_cleanup_var.get(),
                'server_retention_days': self.server_retention_days.get(),
                'selected_databases': self.selected_databases,
                'minimize_to_tray': self.minimize_to_tray.get()
            }
            
            config_file = os.path.join(log_dir, 'backup_config_v0.45.json')
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
            
            self.log(f"配置已保存到 {config_file}，备份列表中包含 {len(self.selected_databases)} 个数据库: {', '.join(self.selected_databases)}")
            messagebox.showinfo("成功", f"配置已保存，备份列表中记住了 {len(self.selected_databases)} 个数据库!")
            
            if self.auto_backup_var.get():
                self.setup_scheduled_backup()
                
            self.update_monitor_status()
            self.update_next_backup_time()
                
        except Exception as e:
            error_msg = f"保存配置失败: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)

    def load_config(self):
        try:
            # 尝试从日志目录加载配置文件
            config_file = None
            
            # 检查日志目录中的配置文件
            if hasattr(logging, 'getLogger'):
                root_logger = logging.getLogger()
                for handler in root_logger.handlers:
                    if isinstance(handler, logging.FileHandler):
                        log_dir = os.path.dirname(handler.baseFilename)
                        config_file = os.path.join(log_dir, 'backup_config_v0.45.json')
                        break
            
            # 如果日志目录中没有配置文件，检查默认位置
            if not config_file or not os.path.exists(config_file):
                # 检查旧版本配置文件
                for old_version in ['v0.44', 'v0.43', 'v0.42', 'v0.41']:
                    # 先检查日志目录
                    if 'log_dir' in locals():
                        old_config = os.path.join(log_dir, f'backup_config_{old_version}.json')
                        if os.path.exists(old_config):
                            config_file = old_config
                            self.log(f"检测到旧版本({old_version})配置文件，正在加载...")
                            break
                    
                    # 再检查程序目录
                    old_config = f'backup_config_{old_version}.json'
                    if os.path.exists(old_config):
                        config_file = old_config
                        self.log(f"检测到旧版本({old_version})配置文件，正在加载...")
                        break
                else:
                    self.log("未找到配置文件，使用默认设置")
                    return
            
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                
                self.server_entry.delete(0, tk.END)
                self.server_entry.insert(0, config.get('server', ''))
                
                self.user_entry.delete(0, tk.END)
                self.user_entry.insert(0, config.get('user', ''))
                
                self.password_entry.delete(0, tk.END)
                self.password_entry.insert(0, config.get('password', ''))
                
                self.server_temp_path_entry.delete(0, tk.END)
                self.server_temp_path_entry.insert(0, config.get('server_temp_path', ''))
                
                self.server_web_url_entry.delete(0, tk.END)
                self.server_web_url_entry.insert(0, config.get('server_web_url', ''))
                
                self.local_save_path_entry.delete(0, tk.END)
                self.local_save_path_entry.insert(0, config.get('local_save_path', ''))
                
                self.filename_prefix_entry.delete(0, tk.END)
                self.filename_prefix_entry.insert(0, config.get('filename_prefix', ''))
                
                # 忽略配置文件中的备份模式，强制使用服务器模式
                self.log("已强制设置为服务器备份模式")
                
                self.auto_backup_var.set(config.get('auto_backup_enabled', False))
                self.backup_hour.set(config.get('backup_hour', '18'))
                self.backup_minute.set(config.get('backup_minute', '00'))
                
                self.auto_cleanup_var.set(config.get('auto_cleanup_var', True))
                self.retention_days.set(config.get('retention_days', 30))
                
                # 加载服务器清理配置
                self.server_auto_cleanup_var.set(config.get('server_auto_cleanup_var', True))
                self.server_retention_days.set(config.get('server_retention_days', 15))
                
                # 加载最小化到托盘设置
                self.minimize_to_tray.set(config.get('minimize_to_tray', True))
                
                self.selected_databases = config.get('selected_databases', [])
                if self.selected_databases:
                    self.log(f"从配置文件加载了 {len(self.selected_databases)} 个已选备份数据库: {', '.join(self.selected_databases)}")
                    self.update_selected_listbox()
                else:
                    self.log("配置文件中未找到已选备份数据库记录")
            
            self.log("配置已加载（版本V0.45）")
            
            # 根据配置状态更新自动备份相关UI
            if self.auto_backup_var.get():
                self.setup_scheduled_backup()
                self.auto_backup_enabled = True
                # 启用时间选择框
                self.hour_combobox.config(state="readonly")
                self.minute_combobox.config(state="readonly")
            else:
                # 禁用时间选择框
                self.hour_combobox.config(state="disabled")
                self.minute_combobox.config(state="disabled")
                
            self.update_monitor_status()
            self.update_selected_count()
            self.update_autostart_status()
            self.update_next_backup_time()
                
        except Exception as e:
            self.log(f"加载配置失败: {str(e)}")
            self.selected_databases = []

    def toggle_password_visibility(self):
        if self.show_password_var.get():
            self.password_entry.config(show="")
        else:
            self.password_entry.config(show="*")

    def add_selected_databases(self):
        selected_indices = self.available_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("提示", "请先从左侧选择要添加的数据库")
            return
            
        added = 0
        for i in selected_indices:
            db_name = self.available_listbox.get(i)
            if db_name not in self.selected_databases:
                self.selected_databases.append(db_name)
                added += 1
        
        self.update_selected_listbox()
        self.log(f"已添加 {added} 个数据库到备份列表")
        self.update_selected_count()

    def remove_selected_databases(self):
        selected_indices = self.selected_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("提示", "请先从右侧选择要移除的数据库")
            return
            
        removed = 0
        for i in sorted(selected_indices, reverse=True):
            del self.selected_databases[i]
            removed += 1
        
        self.update_selected_listbox()
        self.log(f"已从备份列表中移除 {removed} 个数据库")
        self.update_selected_count()

    def add_all_databases(self):
        added = 0
        for db_name in self.available_databases:
            if db_name not in self.selected_databases:
                self.selected_databases.append(db_name)
                added += 1
        
        self.update_selected_listbox()
        self.log(f"已添加所有 {added} 个可用数据库到备份列表")
        self.update_selected_count()

    def remove_all_databases(self):
        count = len(self.selected_databases)
        if count == 0:
            messagebox.showinfo("提示", "备份列表已经是空的")
            return
            
        if messagebox.askyesno("确认", f"确定要从备份列表中移除所有 {count} 个数据库吗？"):
            self.selected_databases = []
            self.update_selected_listbox()
            self.log(f"已清空备份列表（共 {count} 个数据库）")
            self.update_selected_count()

    def update_available_listbox(self):
        self.available_listbox.delete(0, tk.END)
        for db in self.available_databases:
            self.available_listbox.insert(tk.END, db)

    def update_selected_listbox(self):
        self.selected_listbox.delete(0, tk.END)
        for db in self.selected_databases:
            self.selected_listbox.insert(tk.END, db)

    def update_selected_count(self):
        self.selected_count_label.config(text=str(len(self.selected_databases)))

    def is_autostart_enabled(self):
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_READ
            )
            
            value, _ = winreg.QueryValueEx(key, "MSSQLBackupTool")
            winreg.CloseKey(key)
            
            current_path = os.path.abspath(sys.argv[0])
            return value == f'"{current_path}" --silent'
        except WindowsError:
            return False

    def toggle_autostart(self):
        try:
            if self.autostart_var.get():
                current_path = os.path.abspath(sys.argv[0])
                command = f'"{current_path}" --silent'
                
                key = winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    r"Software\Microsoft\Windows\CurrentVersion\Run",
                    0,
                    winreg.KEY_SET_VALUE
                )
                
                winreg.SetValueEx(key, "MSSQLBackupTool", 0, winreg.REG_SZ, command)
                winreg.CloseKey(key)
                
                self.log("已启用随Windows启动（静默模式）")
                messagebox.showinfo("成功", "已启用随Windows启动，程序将以静默模式运行!")
            else:
                key = winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    r"Software\Microsoft\Windows\CurrentVersion\Run",
                    0,
                    winreg.KEY_SET_VALUE
                )
                
                winreg.DeleteValue(key, "MSSQLBackupTool")
                winreg.CloseKey(key)
                
                self.log("已禁用随Windows启动")
                messagebox.showinfo("成功", "已禁用随Windows启动！")
                
            self.update_autostart_status()
            
        except WindowsError as e:
            error_msg = f"修改随Windows启动设置失败: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)

    def update_autostart_status(self):
        if self.autostart_var.get():
            self.autostart_status_label.config(text="已启用（静默模式）", foreground="green")
        else:
            self.autostart_status_label.config(text="未启用", foreground="orange")

    def quit_application(self, force=False):
        # 标记程序为停止状态
        self.running = False
        
        # 保存配置
        self.save_config()
        
        # 停止托盘图标
        if self.tray_initialized and self.tray_icon:
            try:
                self.tray_icon.stop()
                self.log("已停止系统托盘")
            except Exception as e:
                self.log(f"停止托盘时出错: {str(e)}")
        
        # 清理临时图标文件
        if hasattr(self, 'temp_icon_path') and os.path.exists(self.temp_icon_path):
            try:
                os.remove(self.temp_icon_path)
                self.log("已清理临时图标文件")
            except Exception as e:
                self.log(f"清理临时图标时出错: {str(e)}")
                
        # 退出程序
        if force or messagebox.askyesno("确认退出", "确定要退出MSSQL备份工具吗？"):
            self.log("应用程序退出")
            self.root.destroy()
            sys.exit(0)

    def start_monitor_thread(self):
        def update_monitor():
            while self.running:
                self.update_monitor_status()
                time.sleep(30)
                
        monitor_thread = threading.Thread(target=update_monitor, daemon=True)
        monitor_thread.start()

    def update_monitor_status(self):
        if self.auto_backup_enabled:
            self.auto_backup_status_label.config(text="已启用", foreground="green")
        else:
            self.auto_backup_status_label.config(text="未启用", foreground="orange")
            
        self.success_count_label.config(text=str(self.today_success_count))
        self.fail_count_label.config(text=str(self.today_fail_count))
        self.selected_count_label.config(text=str(len(self.selected_databases)))
        self.update_autostart_status()

    def start_schedule_thread(self):
        def run_schedule():
            while self.running:
                schedule.run_pending()
                time.sleep(1)
                
        schedule_thread = threading.Thread(target=run_schedule, daemon=True)
        schedule_thread.start()
        self.log("定时任务线程已启动")

    def toggle_auto_backup(self):
        """切换自动备份状态"""
        if self.auto_backup_var.get():
            # 启用自动备份
            self.setup_scheduled_backup()
            self.auto_backup_enabled = True
            # 启用时间选择框
            self.hour_combobox.config(state="readonly")
            self.minute_combobox.config(state="readonly")
            self.log("已启用自动备份功能")
            # 立即计算并显示下次备份时间
            self.update_next_backup_time()
        else:
            # 禁用自动备份
            self.cancel_scheduled_backup()
            self.auto_backup_enabled = False
            # 禁用时间选择框
            self.hour_combobox.config(state="disabled")
            self.minute_combobox.config(state="disabled")
            # 更新显示
            self.next_backup_label.config(text="下次备份: 未设置")
            self.monitor_next_backup_label.config(text="未设置", foreground="orange")
            self.log("已禁用自动备份功能")
            
        self.update_monitor_status()

    def setup_scheduled_backup(self):
        self.cancel_scheduled_backup()  # 先清除已有任务
        
        # 获取用户设置的时间
        hour = int(self.backup_hour.get())
        minute = int(self.backup_minute.get())
        
        # 每天在指定时间执行备份
        schedule.every().day.at(f"{hour:02d}:{minute:02d}").do(self.run_auto_backup)
        self.log(f"已设置每天 {hour:02d}:{minute:02d} 自动备份")
        
        # 如果启用了服务器自动清理，设置定时清理任务
        if self.server_auto_cleanup_var.get():
            # 在备份完成后30分钟执行服务器清理
            cleanup_minute = (minute + 30) % 60
            cleanup_hour = hour + (1 if (minute + 30) >= 60 else 0)
            cleanup_hour %= 24
            
            schedule.every().day.at(f"{cleanup_hour:02d}:{cleanup_minute:02d}").do(self.run_server_cleanup)
            self.log(f"已设置每天 {cleanup_hour:02d}:{cleanup_minute:02d} 自动清理服务器备份")

    def cancel_scheduled_backup(self):
        schedule.clear()

    def run_auto_backup(self):
        """执行自动备份，在托盘显示通知"""
        if not self.backup_running and self.running:
            # 显示托盘通知
            self.show_tray_notification("自动备份开始", "正在执行数据库自动备份...")
            self.log("===== 自动备份任务开始 =====")
            
            auto_backup_thread = threading.Thread(target=self.perform_auto_backup_task)
            auto_backup_thread.daemon = True
            auto_backup_thread.start()
        else:
            self.log("自动备份任务取消，因为当前有备份正在运行或程序即将退出")

    def perform_auto_backup_task(self):
        """执行自动备份任务的实际逻辑"""
        try:
            self.backup_running = True
            
            # 执行备份
            result = self.perform_backup(is_auto=True)
            
            # 备份完成后显示通知
            if result:
                self.show_tray_notification("自动备份完成", "数据库自动备份已成功完成")
            else:
                self.show_tray_notification("自动备份失败", "数据库自动备份过程中出现错误")
                
        finally:
            self.backup_running = False

    def start_backup_from_tray(self):
        """从托盘菜单启动备份"""
        if not self.backup_running and self.running:
            self.show_tray_notification("手动备份开始", "正在执行数据库手动备份...")
            self.backup_running = True
            
            backup_thread = threading.Thread(target=self.perform_backup_from_tray)
            backup_thread.daemon = True
            backup_thread.start()
        else:
            self.show_tray_notification("备份正在运行", "已有备份任务正在执行中")

    def perform_backup_from_tray(self):
        """从托盘启动的备份任务"""
        try:
            self.perform_backup(is_auto=False)
        finally:
            self.backup_running = False

    def show_log(self):
        """从托盘菜单查看日志"""
        self.show_window()
        # 切换到日志标签页
        tab_control = self.root.nametowidget(self.root.winfo_children()[0]).nametowidget(self.root.winfo_children()[0].winfo_children()[0])
        tab_control.select(2)  # 日志标签页索引为2

    def delete_old_backups(self):
        """删除本地过期备份"""
        try:
            if not self.auto_cleanup_var.get():
                self.log("自动清理功能已禁用，跳过删除过期本地备份")
                return 0, 0
                
            local_save_path = self.local_save_path_entry.get().strip()
            if not local_save_path or not os.path.exists(local_save_path):
                self.log("本地保存路径不存在，无法执行自动清理")
                return 0, 0
                
            days = self.retention_days.get()
            if days <= 0:
                self.log("保留天数设置无效，跳过自动清理")
                return 0, 0
                
            cutoff_time = datetime.datetime.now() - datetime.timedelta(days=days)
            self.log(f"开始自动清理本地备份（版本V0.45）：删除 {days} 天前的备份文件")
            
            prefix = self.filename_prefix_entry.get().strip() or "backup"
            
            deleted_count = 0
            kept_count = 0
            
            for filename in os.listdir(local_save_path):
                if filename.startswith(prefix) and filename.endswith(".bak"):
                    file_path = os.path.join(local_save_path, filename)
                    
                    try:
                        if self.os_type == "Windows":
                            create_time = datetime.datetime.fromtimestamp(os.path.getctime(file_path))
                        else:
                            create_time = datetime.datetime.fromtimestamp(os.path.getmtime(file_path))
                            
                        if create_time < cutoff_time:
                            os.remove(file_path)
                            self.log(f"已删除过期本地备份: {filename}")
                            deleted_count += 1
                        else:
                            kept_count += 1
                    except Exception as e:
                        self.log(f"处理本地文件 {filename} 时出错: {str(e)}")
            
            self.log(f"本地备份清理完成：删除 {deleted_count} 个过期备份，保留 {kept_count} 个最新备份")
            
            # 更新今日清理统计
            self.today_local_cleaned_count += deleted_count
            
            return deleted_count, kept_count
            
        except Exception as e:
            self.log(f"本地备份清理过程出错: {str(e)}")
            return 0, 0

    def delete_server_backups(self, manual=False):
        """删除服务器上的过期备份"""
        try:
            if not manual and not self.server_auto_cleanup_var.get():
                self.log("服务器自动清理功能已禁用，跳过删除过期服务器备份")
                return 0, 0
                
            # 注意：此处仅保留本地日志记录
            days = self.server_retention_days.get()
            if days <= 0:
                self.log("服务器备份保留天数设置无效，跳过服务器清理")
                return 0, 0
                
            self.log(f"开始{'手动' if manual else '自动'}清理服务器备份：删除 {days} 天前的备份文件")
            self.log("服务器清理功能已调整，具体实现需根据实际环境配置")
            
            # 实际应用中需要根据您的服务器清理机制替换以下代码
            # 这里仅作为占位，返回0表示未执行实际删除操作
            deleted_count = 0
            kept_count = 0
            
            self.log(f"服务器备份清理完成：删除 {deleted_count} 个过期备份，保留 {kept_count} 个最新备份")
            
            # 更新今日清理统计
            self.today_server_cleaned_count += deleted_count
            
            return deleted_count, kept_count
                
        except Exception as e:
            self.log(f"服务器备份清理过程出错: {str(e)}")
            return 0, 0

    def manual_cleanup(self):
        """手动清理本地备份"""
        try:
            self.log("手动触发本地备份清理操作...")
            deleted, kept = self.delete_old_backups()
            
            messagebox.showinfo("清理完成", 
                              f"本地备份手动清理完成:\n"
                              f"已删除 {deleted} 个过期备份文件\n"
                              f"保留 {kept} 个最新备份文件")
                              
        except Exception as e:
            error_msg = f"本地备份手动清理失败: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)

    def manual_server_cleanup(self):
        """手动清理服务器备份"""
        try:
            self.log("手动触发服务器备份清理操作...")
            
            # 确认清理操作
            if not messagebox.askyesno("确认清理", 
                                     f"确定要清理服务器上超过 {self.server_retention_days.get()} 天的备份文件吗？\n"
                                     "此操作无法撤销，请谨慎执行！"):
                self.log("用户取消了服务器备份清理操作")
                return
                
            deleted, kept = self.delete_server_backups(manual=True)
            
            messagebox.showinfo("清理完成", 
                              f"服务器备份手动清理完成:\n"
                              f"已删除 {deleted} 个过期备份文件\n"
                              f"保留 {kept} 个最新备份文件")
                              
        except Exception as e:
            error_msg = f"服务器备份手动清理失败: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)

    def run_server_cleanup(self):
        """执行自动服务器备份清理"""
        if self.running:
            self.show_tray_notification("服务器清理开始", "正在执行服务器备份自动清理...")
            self.log("===== 服务器自动清理任务开始 =====")
            
            deleted, kept = self.delete_server_backups()
            
            if deleted > 0 or kept > 0:
                self.show_tray_notification(
                    "服务器清理完成", 
                    f"已删除 {deleted} 个过期备份，保留 {kept} 个最新备份"
                )
            self.log("===== 服务器自动清理任务结束 =====")

    def perform_backup(self, is_auto=False):
        if is_auto:
            self.log("===== 自动备份任务开始 =====")
        
        try:
            databases = self.selected_databases
            if not databases:
                self.log("备份列表中没有选择任何数据库")
                if is_auto:
                    self.log("===== 自动备份任务失败 =====")
                    self.show_tray_notification("备份失败", "备份列表中没有选择任何数据库")
                else:
                    messagebox.showwarning("警告", "请先在备份列表中添加要备份的数据库")
                return False
                
            total = len(databases)
            success_count = 0
            fail_count = 0
            fail_databases = []
            
            self.log(f"开始批量备份 {total} 个数据库...")
            
            for i, db_name in enumerate(databases, 1):
                self.log(f"\n===== 正在备份 {i}/{total}: {db_name} =====")
                result = self.backup_single_database(db_name, is_auto)
                if result:
                    success_count += 1
                else:
                    fail_count += 1
                    fail_databases.append(db_name)
            
            # 先清理本地备份
            local_deleted, local_kept = self.delete_old_backups()
            
            # 如果启用了服务器清理且是自动备份，同时清理服务器
            if is_auto and self.server_auto_cleanup_var.get():
                server_deleted, server_kept = self.delete_server_backups()
            else:
                server_deleted, server_kept = 0, 0
            
            self.log(f"\n===== 备份总结 =====")
            self.log(f"总数据库数: {total}")
            self.log(f"成功: {success_count}")
            self.log(f"失败: {fail_count}")
            if fail_databases:
                self.log(f"备份失败的数据库: {', '.join(fail_databases)}")
            if self.auto_cleanup_var.get():
                self.log(f"本地清理: 删除 {local_deleted} 个过期备份，保留 {local_kept} 个最新备份")
            if is_auto and self.server_auto_cleanup_var.get():
                self.log(f"服务器清理: 删除 {server_deleted} 个过期备份，保留 {server_kept} 个最新备份")
            
            # 备份完成后通知
            if not is_auto:
                if fail_count > 0:
                    msg = f"{success_count} 个数据库备份成功，{fail_count} 个数据库备份失败!\n"
                    if fail_databases:
                        msg += f"失败的数据库: {', '.join(fail_databases)}\n"
                    if self.auto_cleanup_var.get():
                        msg += f"本地清理: 删除 {local_deleted} 个过期备份，保留 {local_kept} 个最新备份\n"
                    msg += "请查看日志了解详细信息"
                    messagebox.showwarning("部分失败", msg)
                else:
                    msg = f"所有 {total} 个数据库备份成功!"
                    if self.auto_cleanup_var.get():
                        msg += f"\n本地清理: 删除 {local_deleted} 个过期备份，保留 {local_kept} 个最新备份"
                    messagebox.showinfo("备份成功", msg)
            else:
                # 自动备份完成后在托盘显示结果
                if fail_count > 0:
                    self.show_tray_notification(
                        "备份部分失败", 
                        f"{success_count} 个成功，{fail_count} 个失败，请查看日志"
                    )
                else:
                    self.show_tray_notification(
                        "备份成功", 
                        f"所有 {total} 个数据库备份完成"
                    )
                
            return fail_count == 0
                
        except Exception as e:
            error_msg = f"批量备份失败: {str(e)}"
            self.log(error_msg)
            
            if not is_auto:
                messagebox.showerror("错误", error_msg)
            else:
                self.show_tray_notification("备份失败", f"批量备份失败: {str(e)}")
                
            return False
        finally:
            self.backup_running = False
            self.backup_button.config(text="开始备份", command=self.start_backup)
            
            if is_auto:
                self.log("===== 自动备份任务结束 =====")
                self.update_next_backup_time()

    def backup_single_database(self, db_name, is_auto=False):
        try:
            local_save_path = self.local_save_path_entry.get().strip()
            if not local_save_path:
                self.log("请选择本地保存路径")
                return False
                
            if not os.path.exists(local_save_path):
                try:
                    os.makedirs(local_save_path)
                    self.log(f"已创建本地目录: {local_save_path}")
                except Exception as e:
                    self.log(f"创建本地目录失败: {str(e)}")
                    if not is_auto:
                        messagebox.showerror("错误", f"创建本地目录失败: {str(e)}")
                    else:
                        self.show_tray_notification("备份失败", f"创建本地目录失败: {str(e)}")
                    return False
                
            prefix = self.filename_prefix_entry.get().strip() or "backup"
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f"{prefix}_{db_name}_{timestamp}.bak"
            local_full_path = os.path.join(local_save_path, backup_filename)
            
            conn_str = self.get_connection_string(timeout=15)
            if not conn_str:
                return False
                
            # 强制使用服务器备份模式
            backup_mode = "server_then_web"
            
            self.log(f"开始备份数据库: {db_name}")
            
            # 仅保留服务器备份模式的代码
            server_temp_path = self.server_temp_path_entry.get().strip()
            server_web_url = self.server_web_url_entry.get().strip()
            
            if not server_temp_path or not server_web_url:
                self.log("服务器路径或Web URL未设置")
                return False
            
            server_temp_path = server_temp_path.replace('\\', '/')
            server_full_path = os.path.join(server_temp_path, backup_filename).replace('\\', '/')
            web_download_url = urljoin(server_web_url, backup_filename)
            
            self.log(f"服务器存储路径: {server_full_path}")
            self.log(f"Web下载地址: {web_download_url}")
            
            conn = pyodbc.connect(conn_str)
            conn.autocommit = True
            cursor = conn.cursor()
            
            backup_sql = f"BACKUP DATABASE [{db_name}] TO DISK = N'{server_full_path}' WITH NOFORMAT, NOINIT, NAME = N'{db_name}-完整 数据库 备份', SKIP, NOREWIND, NOUNLOAD, STATS = 10"
            cursor.execute(backup_sql)
            
            self.log(f"等待服务器端 {db_name} 备份完成...")
            time.sleep(30)  # 等待服务器备份完成
            conn.close()
            
            self.log(f"开始通过Web下载 {db_name} 备份文件...")
            download_success = self.download_file_from_web(web_download_url, local_full_path)
            
            if download_success:
                file_size = os.path.getsize(local_full_path)
                size_str = f"{file_size/(1024*1024):.2f} MB"
                self.log(f"数据库 {db_name} 备份并下载成功!")
                
                self.add_backup_record(db_name, timestamp, "成功", size_str)
                self.update_daily_stats(success=True)
                
                return True
            else:
                self.log(f"{db_name} 备份文件已保存到服务器，但Web下载失败")
                self.add_backup_record(db_name, timestamp, "失败", "0 MB")
                self.update_daily_stats(success=False)
                
                return False
                
        except Exception as e:
            error_msg = f"{db_name} 备份失败: {str(e)}"
            self.log(error_msg)
            self.add_backup_record(db_name, datetime.datetime.now().strftime("%Y%m%d_%H%M%S"), "失败", "0 MB")
            self.update_daily_stats(success=False)
            
            if is_auto:
                self.show_tray_notification("备份失败", error_msg)
                
            return False

    def download_file_from_web(self, web_url, local_path):
        try:
            self.log(f"尝试从Web下载文件: {web_url} 到本地: {local_path}")
            
            local_dir = os.path.dirname(local_path)
            if not os.path.exists(local_dir):
                os.makedirs(local_dir)
                
            response = requests.get(web_url, stream=True, timeout=300)
            response.raise_for_status()
            
            with open(local_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            
            if os.path.exists(local_path) and os.path.getsize(local_path) > 0:
                self.log(f"文件下载成功，大小: {os.path.getsize(local_path) / (1024*1024):.2f} MB")
                return True
            else:
                self.log("文件下载成功，但文件为空或未找到")
                return False
                
        except Exception as e:
            self.log(f"文件下载失败: {str(e)}")
            return False

    def update_daily_stats(self, success=True):
        today = datetime.date.today()
        if not hasattr(self, 'stats_date') or self.stats_date != today:
            self.today_success_count = 0
            self.today_fail_count = 0
            self.today_local_cleaned_count = 0
            self.today_server_cleaned_count = 0
            self.stats_date = today
            
        if success:
            self.today_success_count += 1
        else:
            self.today_fail_count += 1
            
        self.update_monitor_status()

    def add_backup_record(self, db_name, timestamp, status, size):
        try:
            backup_time = datetime.datetime.strptime(timestamp, "%Y%m%d_%H%M%S").strftime("%Y-%m-%d %H:%M:%S")
        except:
            backup_time = timestamp
            
        self.recent_backups.insert(0, (db_name, backup_time, status, size))
        if len(self.recent_backups) > 10:
            self.recent_backups = self.recent_backups[:10]
            
        for item in self.recent_tree.get_children():
            self.recent_tree.delete(item)
            
        for record in self.recent_backups:
            tag = "success" if record[2] == "成功" else "fail"
            self.recent_tree.insert("", tk.END, values=record, tags=(tag,))
            
        self.recent_tree.tag_configure("success", foreground="green")
        self.recent_tree.tag_configure("fail", foreground="red")

    def start_backup(self):
        if self.backup_running:
            return
            
        self.backup_running = True
        self.backup_button.config(text="备份中...", command=None)
        
        backup_thread = threading.Thread(target=self.perform_backup, args=(False,))
        backup_thread.daemon = True
        backup_thread.start()

if __name__ == "__main__":
    silent_mode = "--silent" in sys.argv
    
    required_libs = {
        'pyodbc': 'pyodbc',
        'requests': 'requests',
        'schedule': 'schedule',
        'PIL': 'Pillow',
        'matplotlib': 'matplotlib',
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
        print("=" * 60)
        sys.exit(1)
        
    # 确保在主线程中创建Tk实例
    root = tk.Tk()
    app = MSSQLBackupTool(root, silent_mode)
    root.mainloop()
