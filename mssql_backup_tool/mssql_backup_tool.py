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
from urllib.parse import urljoin
import schedule
from PIL import Image, ImageTk

class MSSQLBackupTool:
    def __init__(self, root):
        self.root = root
        self.root.title("MSSQL数据库备份工具 - V0.01")
        self.root.geometry("900x700")
        
        # 窗口设置
        self.root.protocol("WM_DELETE_WINDOW", self.quit_application)
        self.root.bind("<Unmap>", self.on_minimize)
        
        # 核心变量初始化
        self.log_text = None
        self.backup_running = False
        self.auto_backup_enabled = False
        self.selected_databases = []  # 存储多个选中的数据库名称
        self.auto_cleanup_var = tk.BooleanVar(value=True)
        self.os_type = platform.system()
        
        # 字体设置
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("SimHei", 10))
        self.style.configure("TButton", font=("SimHei", 10))
        self.style.configure("TEntry", font=("SimHei", 10))
        self.style.configure("TCombobox", font=("SimHei", 10))
        
        # 创建界面
        self.create_widgets()
        
        # 初始化日志
        self.setup_logging()
        
        # 设置窗口图标
        self.set_window_icon()
        
        # 窗口居中
        self.center_window()
        self.root.resizable(True, True)
        
        # 加载配置并刷新数据库列表（关键顺序）
        self.load_config()
        self.refresh_databases(initial_load=True)  # 初始加载时恢复多数据库选择
        
        # 启动定时任务线程
        self.start_schedule_thread()
        
        # 启动监控更新线程
        self.start_monitor_thread()

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry('{}x{}+{}+{}'.format(width, height, x, y))

    def set_window_icon(self):
        try:
            if os.path.exists("icon.ico"):
                self.root.iconbitmap("icon.ico")
            else:
                self.log("未找到icon.ico，使用默认图标")
                icon = Image.new('RGB', (32, 32), color='blue')
                icon.save("temp_icon.ico")
                self.root.iconbitmap("temp_icon.ico")
                os.remove("temp_icon.ico")
        except Exception as e:
            self.log(f"设置窗口图标失败: {str(e)}")

    def setup_logging(self):
        logging.basicConfig(
            filename='backup_log_v0.01.txt',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            encoding='utf-8'
        )
        self.log("程序启动（版本V0.01）")

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

    def create_config_tab(self):
        config_frame = ttk.Frame(self.config_tab, padding="10")
        config_frame.pack(fill=tk.BOTH, expand=True)
        
        # 数据库连接设置
        conn_frame = ttk.LabelFrame(config_frame, text="数据库连接设置", padding="10")
        conn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(conn_frame, text="服务器地址:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.server_entry = ttk.Entry(conn_frame, width=50)
        self.server_entry.grid(row=0, column=1, sticky=tk.W, pady=5)
        ttk.Label(conn_frame, text="(例如: 192.168.1.100\\SQLEXPRESS)").grid(row=0, column=2, sticky=tk.W, pady=5)
        
        ttk.Label(conn_frame, text="用户名:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.user_entry = ttk.Entry(conn_frame, width=50)
        self.user_entry.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(conn_frame, text="密码:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.password_entry = ttk.Entry(conn_frame, width=50, show="*")
        self.password_entry.grid(row=2, column=1, sticky=tk.W, pady=5)
        self.show_password_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            conn_frame, 
            text="显示密码", 
            variable=self.show_password_var,
            command=self.toggle_password_visibility
        ).grid(row=2, column=2, sticky=tk.W, pady=5)
        
        # 数据库选择（支持多选）
        db_frame = ttk.LabelFrame(config_frame, text="数据库选择 (可多选)", padding="10")
        db_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(db_frame, text="刷新数据库列表", command=self.refresh_databases).pack(side=tk.RIGHT, padx=5)
        ttk.Label(db_frame, text="选择要备份的数据库:").pack(side=tk.LEFT, padx=5, pady=5)
        
        # 创建带滚动条的数据库列表（支持显示更多数据库）
        db_list_frame = ttk.Frame(db_frame)
        db_list_frame.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)
        
        self.db_listbox = tk.Listbox(db_list_frame, selectmode=tk.MULTIPLE, width=50, height=6)  # 增加高度显示更多选项
        self.db_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        db_scrollbar = ttk.Scrollbar(db_list_frame, orient=tk.VERTICAL, command=self.db_listbox.yview)
        db_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.db_listbox.config(yscrollcommand=db_scrollbar.set)
        
        # 全选/取消全选按钮
        ttk.Button(db_frame, text="全选", command=self.select_all_databases).pack(side=tk.RIGHT, padx=5)
        ttk.Button(db_frame, text="取消全选", command=self.deselect_all_databases).pack(side=tk.RIGHT, padx=5)
        
        # 绑定选择事件，实时跟踪多数据库选择变化
        self.db_listbox.bind('<<ListboxSelect>>', self.on_database_select)
        
        # 备份设置
        backup_frame = ttk.LabelFrame(config_frame, text="备份设置", padding="10")
        backup_frame.pack(fill=tk.X, pady=5)
        
        # 备份模式选择等其他配置项保持不变...
        ttk.Label(backup_frame, text="备份模式:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        self.backup_mode = tk.StringVar(value="local")
        ttk.Radiobutton(backup_frame, text="直接本地备份", variable=self.backup_mode, value="local").grid(row=0, column=1, sticky=tk.W, pady=5)
        ttk.Radiobutton(backup_frame, text="先备份到服务器再通过Web下载", variable=self.backup_mode, value="server_then_web").grid(row=0, column=2, sticky=tk.W, pady=5)
        
        ttk.Label(backup_frame, text="服务器IIS发布路径:").grid(row=1, column=0, sticky=tk.W, pady=5, padx=5)
        self.server_temp_path_entry = ttk.Entry(backup_frame, width=40)
        self.server_temp_path_entry.grid(row=1, column=1, sticky=tk.W, pady=5)
        ttk.Label(backup_frame, text="(服务器上IIS发布的本地路径)").grid(row=1, column=2, sticky=tk.W, pady=5)
        
        ttk.Label(backup_frame, text="服务器Web访问URL:").grid(row=2, column=0, sticky=tk.W, pady=5, padx=5)
        self.server_web_url_entry = ttk.Entry(backup_frame, width=40)
        self.server_web_url_entry.grid(row=2, column=1, sticky=tk.W, pady=5)
        ttk.Label(backup_frame, text="(例如: http://192.168.1.100/backup/)").grid(row=2, column=2, sticky=tk.W, pady=5)
        
        path_frame = ttk.Frame(backup_frame)
        path_frame.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=5, padx=5)
        
        ttk.Label(path_frame, text="本地保存路径:").pack(side=tk.LEFT, padx=5, pady=5)
        self.local_save_path_entry = ttk.Entry(path_frame, width=40)
        self.local_save_path_entry.pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(path_frame, text="浏览...", command=self.browse_local_path).pack(side=tk.LEFT, padx=5, pady=5)
        
        ttk.Label(backup_frame, text="文件名前缀:").grid(row=4, column=0, sticky=tk.W, pady=5, padx=5)
        self.filename_prefix_entry = ttk.Entry(backup_frame, width=40)
        self.filename_prefix_entry.grid(row=4, column=1, sticky=tk.W, pady=5)
        ttk.Label(backup_frame, text="(将自动添加数据库名和日期时间)").grid(row=4, column=2, sticky=tk.W, pady=5)
        
        # 自动清理设置
        cleanup_frame = ttk.LabelFrame(config_frame, text="自动清理设置", padding="10")
        cleanup_frame.pack(fill=tk.X, pady=5)
        
        ttk.Checkbutton(
            cleanup_frame, 
            text="启用自动删除过期备份", 
            variable=self.auto_cleanup_var
        ).grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        
        ttk.Label(cleanup_frame, text="保留最近的天数:").grid(row=0, column=1, sticky=tk.W, pady=5)
        self.retention_days = tk.IntVar(value=30)
        ttk.Combobox(
            cleanup_frame, 
            textvariable=self.retention_days,
            values=[7, 15, 30, 60, 90, 180, 365],
            width=5
        ).grid(row=0, column=2, sticky=tk.W, pady=5, padx=2)
        
        ttk.Label(cleanup_frame, text="天的备份文件").grid(row=0, column=3, sticky=tk.W, pady=5)
        ttk.Button(cleanup_frame, text="立即清理", command=self.manual_cleanup).grid(row=0, column=4, padx=10, pady=5)
        
        # 自动备份设置
        auto_backup_frame = ttk.LabelFrame(config_frame, text="自动备份设置", padding="10")
        auto_backup_frame.pack(fill=tk.X, pady=5)
        
        self.auto_backup_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            auto_backup_frame, 
            text="启用每天自动备份", 
            variable=self.auto_backup_var,
            command=self.toggle_auto_backup
        ).grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        
        ttk.Label(auto_backup_frame, text="自动备份时间:").grid(row=0, column=1, sticky=tk.W, pady=5)
        self.backup_hour = tk.StringVar(value="18")
        self.backup_minute = tk.StringVar(value="00")
        
        ttk.Combobox(
            auto_backup_frame, 
            textvariable=self.backup_hour,
            values=[f"{i:02d}" for i in range(24)],
            width=5
        ).grid(row=0, column=2, sticky=tk.W, pady=5, padx=2)
        
        ttk.Label(auto_backup_frame, text=":").grid(row=0, column=3, sticky=tk.W, pady=5)
        
        ttk.Combobox(
            auto_backup_frame, 
            textvariable=self.backup_minute,
            values=[f"{i:02d}" for i in range(0, 60, 5)],
            width=5
        ).grid(row=0, column=4, sticky=tk.W, pady=5, padx=2)
        
        self.next_backup_label = ttk.Label(auto_backup_frame, text="下次自动备份时间: 未设置")
        self.next_backup_label.grid(row=0, column=5, sticky=tk.W, pady=5, padx=10)
        
        # 备份按钮
        button_frame = ttk.Frame(config_frame, padding="10")
        button_frame.pack(fill=tk.X, pady=5)
        
        self.backup_button = ttk.Button(button_frame, text="开始备份", command=self.start_backup)
        self.backup_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="测试Web连接", command=self.test_web_connection).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="保存配置", command=self.save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="测试数据库连接", command=self.test_connection).pack(side=tk.LEFT, padx=5)

    def create_log_tab(self):
        log_frame = ttk.Frame(self.log_tab, padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        log_control_frame = ttk.Frame(log_frame)
        log_control_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(log_control_frame, text="清空日志", command=self.clear_log).pack(side=tk.RIGHT, padx=5)
        ttk.Button(log_control_frame, text="导出日志", command=self.export_log).pack(side=tk.RIGHT, padx=5)
        
        log_display_frame = ttk.LabelFrame(log_frame, text="操作日志（版本V0.01）", padding="10")
        log_display_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = tk.Text(log_display_frame, height=25, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)
        
        scrollbar = ttk.Scrollbar(log_display_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        self.load_history_log()

    def create_monitor_tab(self):
        monitor_frame = ttk.Frame(self.monitor_tab, padding="10")
        monitor_frame.pack(fill=tk.BOTH, expand=True)
        
        status_frame = ttk.LabelFrame(monitor_frame, text="系统状态（版本V0.01）", padding="10")
        status_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(status_frame, text="程序启动时间:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        self.start_time_label = ttk.Label(status_frame, text=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.start_time_label.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(status_frame, text="数据库连接状态:").grid(row=1, column=0, sticky=tk.W, pady=5, padx=5)
        self.db_status_label = ttk.Label(status_frame, text="未检测", foreground="orange")
        self.db_status_label.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(status_frame, text="自动备份状态:").grid(row=2, column=0, sticky=tk.W, pady=5, padx=5)
        self.auto_backup_status_label = ttk.Label(status_frame, text="未启用", foreground="orange")
        self.auto_backup_status_label.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        stats_frame = ttk.LabelFrame(monitor_frame, text="备份统计", padding="10")
        stats_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(stats_frame, text="今日成功备份数:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        self.success_count_label = ttk.Label(stats_frame, text="0")
        self.success_count_label.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(stats_frame, text="今日失败备份数:").grid(row=0, column=2, sticky=tk.W, pady=5, padx=5)
        self.fail_count_label = ttk.Label(stats_frame, text="0")
        self.fail_count_label.grid(row=0, column=3, sticky=tk.W, pady=5)
        
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
        
        self.today_success_count = 0
        self.today_fail_count = 0
        self.recent_backups = []

    def toggle_password_visibility(self):
        if self.show_password_var.get():
            self.password_entry.config(show="")
        else:
            self.password_entry.config(show="*")

    def select_all_databases(self):
        """全选所有数据库"""
        for i in range(self.db_listbox.size()):
            self.db_listbox.selection_set(i)
        self.selected_databases = self.get_selected_databases()
        self.log(f"已全选 {len(self.selected_databases)} 个数据库")
            
    def deselect_all_databases(self):
        """取消全选所有数据库"""
        self.db_listbox.selection_clear(0, tk.END)
        self.selected_databases = []
        self.log("已取消全选所有数据库")

    def get_selected_databases(self):
        """获取当前选中的多个数据库列表"""
        selected_indices = self.db_listbox.curselection()
        databases = [self.db_listbox.get(i) for i in selected_indices]
        return databases

    def on_database_select(self, event):
        """实时跟踪多个数据库选择变化"""
        if not event.widget == self.db_listbox:
            return
            
        current_selection = self.get_selected_databases()
        if current_selection != self.selected_databases:
            self.selected_databases = current_selection
            # 显示选中的数据库名称，方便用户确认
            self.log(f"已选择 {len(self.selected_databases)} 个数据库: {', '.join(current_selection)}")

    def log(self, message):
        if self.log_text is None:
            print(f"日志: {message}")
            return
            
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        logging.info(message)

    def clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.log("日志已清空")

    def export_log(self):
        try:
            self.log_text.config(state=tk.NORMAL)
            log_content = self.log_text.get(1.0, tk.END)
            self.log_text.config(state=tk.DISABLED)
            
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"backup_log_export_{timestamp}_V0.01.txt"
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
            log_file = 'backup_log_v0.01.txt'
            if os.path.exists(log_file):
                with open(log_file, 'r', encoding='utf-8') as f:
                    lines = f.readlines()[-100:]
                    
                    self.log_text.config(state=tk.NORMAL)
                    for line in lines:
                        self.log_text.insert(tk.END, line)
                    self.log_text.see(tk.END)
                    self.log_text.config(state=tk.DISABLED)
                self.log("已加载最近的日志记录（版本V0.01）")
        except Exception as e:
            self.log(f"加载历史日志失败: {str(e)}")

    def browse_local_path(self):
        path = filedialog.askdirectory(title="选择本地保存路径")
        if path:
            self.local_save_path_entry.delete(0, tk.END)
            self.local_save_path_entry.insert(0, path)

    def get_connection_string(self, database="master"):
        server = self.server_entry.get().strip()
        user = self.user_entry.get().strip()
        password = self.password_entry.get().strip()
        
        if not server:
            messagebox.showerror("错误", "请输入服务器地址")
            return None
            
        if user:
            conn_str = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={user};PWD={password};AutoCommit=True"
        else:
            conn_str = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes;AutoCommit=True"
            
        return conn_str

    def test_connection(self):
        try:
            self.log("正在测试数据库连接...")
            conn_str = self.get_connection_string()
            if not conn_str:
                return
                
            conn = pyodbc.connect(conn_str)
            conn.close()
            self.log("数据库连接成功!")
            messagebox.showinfo("成功", "数据库连接成功!")
            self.db_status_label.config(text="连接正常", foreground="green")
        except Exception as e:
            error_msg = f"数据库连接失败: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)
            self.db_status_label.config(text="连接失败", foreground="red")

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

    def refresh_databases(self, initial_load=False):
        """刷新数据库列表并恢复多个选中的数据库"""
        try:
            self.log("正在刷新数据库列表...")
            conn_str = self.get_connection_string()
            if not conn_str:
                return
                
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            
            # 查询所有非系统数据库
            cursor.execute("SELECT name FROM sys.databases WHERE name NOT IN ('master', 'tempdb', 'model', 'msdb') ORDER BY name")
            databases = [row[0] for row in cursor.fetchall()]
            conn.close()
            
            # 保存当前选择（非初始加载时）
            current_selection = self.get_selected_databases()
            if current_selection and not initial_load:
                self.selected_databases = current_selection
                self.log(f"保存当前选择: {', '.join(current_selection)}")
                
            # 清空并重新填充列表
            self.db_listbox.delete(0, tk.END)
            for db in databases:
                self.db_listbox.insert(tk.END, db)
            
            # 恢复多个选中的数据库（核心修复点）
            restored_count = 0
            if self.selected_databases:
                # 遍历所有数据库，匹配并选中之前选择的多个数据库
                for i, db in enumerate(databases):
                    if db in self.selected_databases:
                        self.db_listbox.selection_set(i)
                        restored_count += 1
            
            self.log(f"共加载 {len(databases)} 个数据库，成功恢复 {restored_count} 个选中的数据库")
            self.db_status_label.config(text="连接正常", foreground="green")
            
        except Exception as e:
            error_msg = f"刷新数据库列表失败: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)
            self.db_status_label.config(text="连接失败", foreground="red")

    def save_config(self):
        """保存多个选中的数据库到配置文件"""
        try:
            # 强制更新选中的数据库列表
            self.selected_databases = self.get_selected_databases()
            
            config = {
                'version': 'V0.01',
                'server': self.server_entry.get(),
                'user': self.user_entry.get(),
                'password': self.password_entry.get(),
                'server_temp_path': self.server_temp_path_entry.get(),
                'server_web_url': self.server_web_url_entry.get(),
                'local_save_path': self.local_save_path_entry.get(),
                'filename_prefix': self.filename_prefix_entry.get(),
                'backup_mode': self.backup_mode.get(),
                'auto_backup_enabled': self.auto_backup_var.get(),
                'backup_hour': self.backup_hour.get(),
                'backup_minute': self.backup_minute.get(),
                'auto_cleanup_var': self.auto_cleanup_var.get(),
                'retention_days': self.retention_days.get(),
                'selected_databases': self.selected_databases  # 保存多个选中的数据库
            }
            
            # 写入配置文件
            config_file = 'backup_config_v0.01.json'
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
            
            self.log(f"配置已保存，包含 {len(self.selected_databases)} 个选中的数据库: {', '.join(self.selected_databases)}")
            messagebox.showinfo("成功", f"配置已保存，记住了 {len(self.selected_databases)} 个选中的数据库!")
            
            if self.auto_backup_var.get():
                self.setup_scheduled_backup()
                
            self.update_monitor_status()
                
        except Exception as e:
            error_msg = f"保存配置失败: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)

    def load_config(self):
        """从配置文件加载多个选中的数据库"""
        try:
            config_file = 'backup_config_v0.01.json'
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    
                    if config.get('version') != 'V0.01':
                        self.log("配置文件版本不匹配，将使用默认配置")
                        return
                    
                    # 加载基本配置
                    self.server_entry.delete(0, tk.END)
                    self.server_entry.insert(0, config.get('server', ''))
                    
                    self.user_entry.delete(0, tk.END)
                    self.user_entry.insert(0, config.get('user', ''))
                    
                    self.password_entry.delete(0, tk.END)
                    self.password_entry.insert(0, config.get('password', ''))
                    
                    # 其他配置项加载...
                    self.server_temp_path_entry.delete(0, tk.END)
                    self.server_temp_path_entry.insert(0, config.get('server_temp_path', ''))
                    
                    self.server_web_url_entry.delete(0, tk.END)
                    self.server_web_url_entry.insert(0, config.get('server_web_url', ''))
                    
                    self.local_save_path_entry.delete(0, tk.END)
                    self.local_save_path_entry.insert(0, config.get('local_save_path', ''))
                    
                    self.filename_prefix_entry.delete(0, tk.END)
                    self.filename_prefix_entry.insert(0, config.get('filename_prefix', ''))
                    
                    self.backup_mode.set(config.get('backup_mode', 'local'))
                    self.auto_backup_var.set(config.get('auto_backup_enabled', False))
                    self.backup_hour.set(config.get('backup_hour', '18'))
                    self.backup_minute.set(config.get('backup_minute', '00'))
                    
                    self.auto_cleanup_var.set(config.get('auto_cleanup_var', True))
                    self.retention_days.set(config.get('retention_days', 30))
                    
                    # 加载多个选中的数据库（核心修复点）
                    self.selected_databases = config.get('selected_databases', [])
                    if self.selected_databases:
                        self.log(f"从配置文件加载了 {len(self.selected_databases)} 个选中的数据库: {', '.join(self.selected_databases)}")
                    else:
                        self.log("配置文件中未找到选中的数据库记录")
                    
                self.log("配置已加载（版本V0.01）")
                
                if self.auto_backup_var.get():
                    self.setup_scheduled_backup()
                    self.auto_backup_enabled = True
                    
                self.update_monitor_status()
                    
        except Exception as e:
            self.log(f"加载配置失败: {str(e)}")
            self.selected_databases = []  # 加载失败时清空选择列表

    # 以下方法保持不变
    def normalize_path(self, path):
        path = path.replace('\\', '/')
        return path

    def delete_old_backups(self):
        try:
            if not self.auto_cleanup_var.get():
                self.log("自动清理功能已禁用，跳过删除过期备份")
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
            self.log(f"开始自动清理（版本V0.01）：删除 {days} 天前的备份文件")
            
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
                            self.log(f"已删除过期备份: {filename}")
                            deleted_count += 1
                        else:
                            kept_count += 1
                    except Exception as e:
                        self.log(f"处理文件 {filename} 时出错: {str(e)}")
            
            self.log(f"自动清理完成：删除 {deleted_count} 个过期备份，保留 {kept_count} 个最新备份")
            return deleted_count, kept_count
            
        except Exception as e:
            self.log(f"自动清理过程出错: {str(e)}")
            return 0, 0

    def manual_cleanup(self):
        try:
            self.log("手动触发清理操作...")
            deleted, kept = self.delete_old_backups()
            
            messagebox.showinfo("清理完成", 
                              f"手动清理完成:\n"
                              f"已删除 {deleted} 个过期备份文件\n"
                              f"保留 {kept} 个最新备份文件")
                              
        except Exception as e:
            error_msg = f"手动清理失败: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)

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

    def wait_for_backup_completion(self, conn, backup_file_path, timeout=300):
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                with open(backup_file_path, 'a'):
                    return True
            except IOError:
                self.log("备份进行中...")
                time.sleep(5)
            except Exception:
                time.sleep(5)
        
        self.log(f"备份超时（{timeout}秒）")
        return False

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
                    return False
                
            prefix = self.filename_prefix_entry.get().strip() or "backup"
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f"{prefix}_{db_name}_{timestamp}.bak"
            local_full_path = os.path.join(local_save_path, backup_filename)
            
            conn_str = self.get_connection_string()
            if not conn_str:
                return False
                
            backup_mode = self.backup_mode.get()
            
            self.log(f"开始备份数据库: {db_name}")
            
            if backup_mode == "local":
                full_path = self.normalize_path(local_full_path)
                self.log(f"备份文件将保存到: {full_path}")
                
                conn = pyodbc.connect(conn_str)
                conn.autocommit = True
                cursor = conn.cursor()
                
                backup_sql = f"BACKUP DATABASE [{db_name}] TO DISK = N'{full_path}' WITH NOFORMAT, NOINIT, NAME = N'{db_name}-完整 数据库 备份', SKIP, NOREWIND, NOUNLOAD, STATS = 10"
                cursor.execute(backup_sql)
                
                self.log(f"等待 {db_name} 备份完成...")
                if self.wait_for_backup_completion(conn, full_path):
                    self.log(f"{db_name} 备份完成")
                else:
                    self.log(f"{db_name} 备份可能未完成")
                    
                conn.close()
                time.sleep(3)
                
                if os.path.exists(full_path):
                    file_size = os.path.getsize(full_path)
                    size_str = f"{file_size/(1024*1024):.2f} MB"
                    self.log(f"数据库 {db_name} 备份成功!")
                    self.log(f"备份文件大小: {size_str}")
                    
                    self.add_backup_record(db_name, timestamp, "成功", size_str)
                    self.update_daily_stats(success=True)
                    return True
                else:
                    self.log(f"备份过程完成，但未找到 {db_name} 的备份文件")
                    self.add_backup_record(db_name, timestamp, "失败", "0 MB")
                    self.update_daily_stats(success=False)
                    return False
                    
            else:
                server_temp_path = self.server_temp_path_entry.get().strip()
                server_web_url = self.server_web_url_entry.get().strip()
                
                if not server_temp_path or not server_web_url:
                    self.log("服务器路径或Web URL未设置")
                    return False
                
                server_temp_path = self.normalize_path(server_temp_path)
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
                if self.server_entry.get().lower() in ['localhost', '127.0.0.1', '.']:
                    if self.wait_for_backup_completion(conn, server_full_path):
                        self.log(f"服务器端 {db_name} 备份完成")
                    else:
                        self.log(f"服务器端 {db_name} 备份可能未完成")
                else:
                    estimated_size = self.estimate_database_size(conn, db_name)
                    wait_time = max(10, min(300, estimated_size // 10))
                    self.log(f"远程服务器 {db_name} 备份中，将等待约 {wait_time} 秒...")
                    time.sleep(wait_time)
                
                conn.close()
                
                self.log(f"开始通过Web下载 {db_name} 备份文件...")
                download_success = self.download_file_from_web(web_download_url, local_full_path)
                
                if download_success:
                    self.log(f"服务器上的 {db_name} 临时备份文件将被保留")
                    
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
            return False

    def perform_backup(self, is_auto=False):
        if is_auto:
            self.log("===== 自动备份任务开始 =====")
        
        try:
            databases = self.get_selected_databases()
            if not databases:
                self.log("未选择任何要备份的数据库")
                if is_auto:
                    self.log("===== 自动备份任务失败 =====")
                else:
                    messagebox.showwarning("警告", "请先选择要备份的数据库")
                return
                
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
            
            deleted_count, kept_count = self.delete_old_backups()
            
            self.log(f"\n===== 备份总结 =====")
            self.log(f"总数据库数: {total}")
            self.log(f"成功: {success_count}")
            self.log(f"失败: {fail_count}")
            if fail_databases:
                self.log(f"备份失败的数据库: {', '.join(fail_databases)}")
            if self.auto_cleanup_var.get():
                self.log(f"自动清理: 删除 {deleted_count} 个过期备份，保留 {kept_count} 个最新备份")
            
            if not is_auto and fail_count > 0:
                msg = f"{success_count} 个数据库备份成功，{fail_count} 个数据库备份失败!\n"
                if fail_databases:
                    msg += f"失败的数据库: {', '.join(fail_databases)}\n"
                if self.auto_cleanup_var.get():
                    msg += f"自动清理: 删除 {deleted_count} 个过期备份，保留 {kept_count} 个最新备份\n"
                msg += "请查看日志了解详细信息"
                messagebox.showwarning("部分失败", msg)
            elif not is_auto:
                msg = f"所有 {total} 个数据库备份成功!"
                if self.auto_cleanup_var.get():
                    msg += f"\n自动清理: 删除 {deleted_count} 个过期备份，保留 {kept_count} 个最新备份"
                messagebox.showinfo("备份成功", msg)
                
        except Exception as e:
            error_msg = f"批量备份失败: {str(e)}"
            self.log(error_msg)
            if not is_auto:
                messagebox.showerror("错误", error_msg)
        finally:
            self.backup_running = False
            self.backup_button.config(text="开始备份", command=self.start_backup)
            
            if is_auto:
                self.log("===== 自动备份任务结束 =====")
                self.update_next_backup_time()

    def estimate_database_size(self, conn, db_name):
        try:
            cursor = conn.cursor()
            cursor.execute(f"USE {db_name}; EXEC sp_spaceused;")
            while True:
                row = cursor.fetchone()
                if row and row[0] == 'database_size':
                    size_str = row[1]
                    size_mb = float(size_str.split()[0])
                    return size_mb
                if not row:
                    break
            return 100
        except:
            return 100

    def start_backup(self):
        if self.backup_running:
            return
            
        self.backup_running = True
        self.backup_button.config(text="备份中...", command=None)
        
        backup_thread = threading.Thread(target=self.perform_backup, args=(False,))
        backup_thread.daemon = True
        backup_thread.start()

    def start_schedule_thread(self):
        def run_schedule():
            while True:
                schedule.run_pending()
                time.sleep(1)
                
        schedule_thread = threading.Thread(target=run_schedule, daemon=True)
        schedule_thread.start()
        self.log("定时任务线程已启动")

    def toggle_auto_backup(self):
        if self.auto_backup_var.get():
            self.setup_scheduled_backup()
            self.auto_backup_enabled = True
            self.log("已启用自动备份功能")
        else:
            self.cancel_scheduled_backup()
            self.auto_backup_enabled = False
            self.next_backup_label.config(text="下次自动备份时间: 未设置")
            self.log("已禁用自动备份功能")
            
        self.update_monitor_status()

    def setup_scheduled_backup(self):
        self.cancel_scheduled_backup()
        
        hour = int(self.backup_hour.get())
        minute = int(self.backup_minute.get())
        
        schedule.every().day.at(f"{hour:02d}:{minute:02d}").do(self.run_auto_backup)
        
        self.log(f"已设置每天 {hour:02d}:{minute:02d} 自动备份")
        self.update_next_backup_time()
        self.update_monitor_status()

    def cancel_scheduled_backup(self):
        schedule.clear()

    def run_auto_backup(self):
        if not self.backup_running:
            auto_backup_thread = threading.Thread(target=self.perform_backup, args=(True,))
            auto_backup_thread.daemon = True
            auto_backup_thread.start()
        else:
            self.log("自动备份任务取消，因为当前有备份正在运行")

    def update_next_backup_time(self):
        try:
            if not self.auto_backup_enabled:
                return
                
            now = datetime.datetime.now()
            hour = int(self.backup_hour.get())
            minute = int(self.backup_minute.get())
            
            next_backup = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
            if next_backup < now:
                next_backup += datetime.timedelta(days=1)
                
            next_time_str = next_backup.strftime("%Y-%m-%d %H:%M:%S")
            self.next_backup_label.config(text=f"下次自动备份时间: {next_time_str}")
            
            self.update_monitor_status()
            
        except Exception as e:
            self.log(f"更新下次备份时间失败: {str(e)}")

    def on_minimize(self, event):
        pass

    def quit_application(self):
        if messagebox.askyesno("确认退出", "确定要退出MSSQL备份工具吗？"):
            self.log("应用程序退出")
            self.root.destroy()
            sys.exit(0)

    def start_monitor_thread(self):
        def update_monitor():
            while True:
                self.update_monitor_status()
                time.sleep(30)
                
        monitor_thread = threading.Thread(target=update_monitor, daemon=True)
        monitor_thread.start()

    def update_monitor_status(self):
        if self.auto_backup_enabled:
            status_text = f"已启用，下次备份: {self.next_backup_label['text'].split(': ')[1]}"
            self.auto_backup_status_label.config(text=status_text, foreground="green")
        else:
            self.auto_backup_status_label.config(text="未启用", foreground="orange")
            
        self.success_count_label.config(text=str(self.today_success_count))
        self.fail_count_label.config(text=str(self.today_fail_count))

    def update_daily_stats(self, success=True):
        today = datetime.date.today()
        if not hasattr(self, 'stats_date') or self.stats_date != today:
            self.today_success_count = 0
            self.today_fail_count = 0
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

if __name__ == "__main__":
    required_libs = {
        'pyodbc': 'pyodbc',
        'requests': 'requests',
        'schedule': 'schedule',
        'PIL': 'Pillow'
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
        
    root = tk.Tk()
    app = MSSQLBackupTool(root)
    root.mainloop()
