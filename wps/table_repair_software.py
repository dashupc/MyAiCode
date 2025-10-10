import os
import shutil
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sys
import random
import time
from datetime import datetime
import ctypes  # 显式导入ctypes确保图标设置生效

class TableRepairSoftware:
    def __init__(self, root):
        self.root = root
        self.root.title("底层代码碎片重组")
        
        # 优先设置应用程序图标
        self.set_application_icon()
        
        self.root.geometry("400x200")  # 密码窗口大小
        self.root.resizable(False, False)  # 密码窗口不可调整大小
        
        # 窗口居中显示
        self.center_window()
        
        # 设置中文字体支持
        self.setup_fonts()
        
        # 获取当前日期作为密码（8位数字）
        self.correct_password = datetime.now().strftime("%Y%m%d")
        
        # 创建密码验证界面
        self.create_password_screen()
        
        # 主界面元素初始化
        self.main_frame = None
        
        # 搜索相关变量
        self.searching = False
        self.stop_search = False
        self.results = []  # 存储(路径, 文件名, 文件大小, 修改时间)元组
        self.selected_drives = []
        self.animation_running = False
        self.animation_end_time = 0
        self.animation_duration = 60  # 默认1分钟
        self.delay_complete = False  # 延迟是否完成
        
        # 结果显示延迟时间（秒）
        self.result_delays = {
            60: 120,    # 初级重组：2分钟
            300: 600,   # 中级重组：10分钟
            1800: 1800  # 加强重组：30分钟
        }
        
        # 动画内容列表
        self.animation_messages = [
            "正在解析磁盘分区表...",
            "扫描文件分配表...",
            "重建文件碎片索引...",
            "识别WPS缓存签名...",
            "验证文件完整性...",
            "分析数据块结构...",
            "读取扇区数据...",
            "恢复丢失的索引节点...",
            "重组分散的数据块...",
            "检查文件系统一致性...",
            "修复表格结构错误...",
            "提取有效数据片段...",
            "重构表格元数据...",
            "验证修复结果...",
            "优化数据存储结构..."
        ]
        self.animation_index = 0
        
    def set_application_icon(self):
        """设置应用程序图标，确保标题栏和任务栏显示正确"""
        try:
            # 确定图标路径（同时支持开发环境和打包后环境）
            if getattr(sys, 'frozen', False):
                # 打包后的环境（EXE运行时）
                icon_path = os.path.join(sys._MEIPASS, "xlsx.ico")
            else:
                # 开发环境（直接运行Python脚本）
                icon_path = "xlsx.ico"
            
            # 检查图标文件是否存在
            if os.path.exists(icon_path):
                # 设置Tkinter窗口图标
                self.root.iconbitmap(default=icon_path)
                
                # 对于Windows系统，使用Win32 API强制设置图标
                if sys.platform.startswith('win'):
                    myappid = 'data.repair.table.1.0'
                    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
                    
                    # 加载图标
                    icon_handle = ctypes.windll.user32.LoadImageW(
                        None, icon_path, 
                        ctypes.c_int(1),  # 图像类型为图标
                        0, 0, 
                        ctypes.c_int(0x00000010)  # 加载标志
                    )
                    
                    if icon_handle != 0:
                        # 获取窗口句柄
                        hwnd = self.root.winfo_id()
                        
                        # 设置大图标和小图标
                        ctypes.windll.user32.SendMessageW(
                            hwnd, ctypes.c_int(0x0080),  # WM_SETICON消息
                            ctypes.c_int(1),  # 大图标
                            icon_handle
                        )
                        ctypes.windll.user32.SendMessageW(
                            hwnd, ctypes.c_int(0x0080),  # WM_SETICON消息
                            ctypes.c_int(0),  # 小图标
                            icon_handle
                        )
            else:
                print(f"警告: 图标文件 {icon_path} 不存在")
                
        except Exception as e:
            print(f"图标设置错误: {str(e)}")
        
    def center_window(self):
        """使窗口在屏幕中居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        
        self.root.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
    def setup_fonts(self):
        # 确保中文正常显示
        default_font = ('SimHei', 10)
        self.root.option_add("*Font", default_font)
    
    def create_password_screen(self):
        """创建无提示的密码验证界面"""
        # 创建密码验证框架
        self.password_frame = ttk.Frame(self.root, padding="50")
        self.password_frame.pack(fill=tk.BOTH, expand=True)
        
        # 仅显示简单标题
        title_label = ttk.Label(
            self.password_frame, 
            text="请输入密码", 
            font=('SimHei', 14, 'bold'),
            foreground="#005fb8"
        )
        title_label.pack(pady=20)
        
        # 密码输入框
        input_frame = ttk.Frame(self.password_frame)
        input_frame.pack(pady=10)
        
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(
            input_frame, 
            textvariable=self.password_var, 
            show="*",  # 密码显示为*
            width=20,
            font=('SimHei', 12)
        )
        self.password_entry.pack(side=tk.LEFT, padx=10)
        self.password_entry.focus_set()  # 设置焦点
        
        # 确认按钮
        btn_frame = ttk.Frame(self.password_frame)
        btn_frame.pack(pady=10)
        
        self.submit_btn = ttk.Button(
            btn_frame, 
            text="确认", 
            command=self.verify_password,
            width=10
        )
        self.submit_btn.pack(side=tk.LEFT, padx=10)
        
        # 绑定回车键验证密码
        self.root.bind('<Return>', lambda event: self.verify_password())
    
    def verify_password(self):
        """验证密码是否正确，错误则直接退出"""
        entered_password = self.password_var.get()
        
        if entered_password == self.correct_password:
            # 密码正确，移除密码界面，显示主界面
            self.password_frame.destroy()
            self.root.geometry("900x1000")  # 恢复主窗口大小
            self.root.resizable(True, True)  # 主窗口可调整大小
            self.center_window()
            self.create_main_interface()
        else:
            # 密码错误，直接退出程序
            self.root.quit()
    
    def create_main_interface(self):
        """创建主软件界面"""
        # 创建主框架
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题部分
        title_frame = ttk.Frame(self.main_frame, padding="10")
        title_frame.pack(fill=tk.X)
        
        title_label = ttk.Label(
            title_frame, 
            text="底层代码碎片重组", 
            font=('SimHei', 16, 'bold'),
            foreground="#005fb8"
        )
        title_label.pack(pady=10)
        
        # 驱动器选择部分 - 移除了滚动条
        drive_frame = ttk.LabelFrame(self.main_frame, text="选择搜索驱动器（可多选）", padding="5")
        drive_frame.pack(fill=tk.X, pady=5)
        
        # 直接使用Frame显示驱动器，不使用滚动条（因为驱动器数量不多）
        self.drive_frame = ttk.Frame(drive_frame)
        self.drive_frame.pack(fill=tk.X, padx=5, pady=2)
        
        # 获取可用驱动器并显示为复选框
        self.update_drives()
        
        # 搜索控制按钮和动画时长选择
        control_frame = ttk.Frame(drive_frame)
        control_frame.pack(side=tk.LEFT, padx=5, pady=2)
        
        # 动画时长选择
        ttk.Label(control_frame, text="重组等级:").pack(side=tk.LEFT, padx=5)
        self.animation_duration_var = tk.IntVar(value=60)  # 默认1分钟
        
        duration_frame = ttk.Frame(control_frame)
        duration_frame.pack(side=tk.LEFT, padx=5)
        
        ttk.Radiobutton(
            duration_frame, 
            text="初级重组", 
            variable=self.animation_duration_var, 
            value=60
        ).pack(side=tk.LEFT, padx=3)
        
        ttk.Radiobutton(
            duration_frame, 
            text="中级重组", 
            variable=self.animation_duration_var, 
            value=300
        ).pack(side=tk.LEFT, padx=3)
        
        ttk.Radiobutton(
            duration_frame, 
            text="加强重组", 
            variable=self.animation_duration_var, 
            value=1800
        ).pack(side=tk.LEFT, padx=3)
        
        # 添加"停止重组"按钮
        self.stop_animation_btn = ttk.Button(
            duration_frame, 
            text="停止重组", 
            command=self.stop_animation
        )
        self.stop_animation_btn.pack(side=tk.LEFT, padx=10)
        
        # 搜索按钮
        btn_frame = ttk.Frame(control_frame)
        btn_frame.pack(side=tk.LEFT, padx=5)
        
        # 创建按钮并保存引用
        self.search_btn = ttk.Button(btn_frame, text="开始重组", command=self.start_search)
        self.search_btn.pack(side=tk.LEFT, padx=3)
        
        self.stop_btn = ttk.Button(btn_frame, text="停止重组", command=self.stop_searching)
        self.stop_btn.pack(side=tk.LEFT, padx=3)
        
        # 全选/取消全选按钮
        self.select_all_btn = ttk.Button(btn_frame, text="全选", command=self.select_all_drives)
        self.select_all_btn.pack(side=tk.LEFT, padx=3)
        
        self.deselect_all_btn = ttk.Button(btn_frame, text="取消全选", command=self.deselect_all_drives)
        self.deselect_all_btn.pack(side=tk.LEFT, padx=3)
        
        # 进度条和动画区域
        progress_frame = ttk.LabelFrame(self.main_frame, text="系统扫描状态", padding="5")
        progress_frame.pack(fill=tk.X, pady=5)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, length=100)
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)
        
        self.status_var = tk.StringVar(value="就绪")
        self.status_label = ttk.Label(progress_frame, textvariable=self.status_var)
        self.status_label.pack(anchor=tk.W, padx=5)
        
        # 代码动画显示区域 - 高度固定为260px
        self.animation_frame = ttk.LabelFrame(self.main_frame, text="底层数据解析", padding="5")
        self.animation_frame.pack(fill=tk.X, pady=5)
        self.animation_frame.configure(height=260)
        self.animation_frame.pack_propagate(False)
        
        self.animation_text = tk.Text(
            self.animation_frame, 
            bg="black", 
            fg="#00FF00", 
            font=('Consolas', 10),
            wrap=tk.WORD,
            state=tk.DISABLED
        )
        self.animation_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 结果显示部分 - 高度固定为260px
        result_frame = ttk.LabelFrame(self.main_frame, text="重组结果", padding="5")
        result_frame.pack(fill=tk.X, pady=5)
        result_frame.configure(height=260)
        result_frame.pack_propagate(False)
        
        # 结果列表上方选择计数
        selection_count_frame = ttk.Frame(result_frame)
        selection_count_frame.pack(fill=tk.X, padx=5, pady=2)
        
        self.selection_count_var = tk.StringVar(value="已选择: 0 项")
        ttk.Label(
            selection_count_frame, 
            textvariable=self.selection_count_var,
            name="selection_count"
        ).pack(side=tk.LEFT, padx=10)
        
        # 结果列表 - 显示三列
        columns = ("文件名", "文件大小", "最后修改时间")
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings")
        
        # 设置列标题和宽度
        self.result_tree.heading("文件名", text="文件名")
        self.result_tree.heading("文件大小", text="文件大小")
        self.result_tree.heading("最后修改时间", text="最后修改时间")
        
        self.result_tree.column("文件名", width=500, stretch=tk.YES)
        self.result_tree.column("文件大小", width=150, stretch=tk.YES)
        self.result_tree.column("最后修改时间", width=200, stretch=tk.YES)
        
        # 绑定选择事件
        self.result_tree.bind("<<TreeviewSelect>>", self.update_selection_count)
        
        # 添加滚动条
        scrollbar_y = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        scrollbar_x = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        self.result_tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)
        
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5,0), pady=(0,5))
        
        # 结果操作按钮区域
        result_controls = ttk.Frame(self.main_frame, padding="5")
        result_controls.pack(fill=tk.X, pady=5)
        
        # 全选和反选按钮居中显示
        center_controls = ttk.Frame(result_controls)
        center_controls.pack(expand=True)
        
        ttk.Button(
            center_controls, 
            text="全选", 
            command=self.select_all_results
        ).pack(side=tk.LEFT, padx=10)
        
        ttk.Button(
            center_controls, 
            text="反选", 
            command=self.invert_selection
        ).pack(side=tk.LEFT, padx=10)
        
        # 复制部分
        copy_frame = ttk.Frame(self.main_frame, padding="5")
        copy_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(copy_frame, text="目标目录:").pack(side=tk.LEFT, padx=5)
        
        self.target_dir_var = tk.StringVar(value="")
        self.target_dir_entry = ttk.Entry(copy_frame, textvariable=self.target_dir_var, width=70)
        self.target_dir_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        self.browse_btn = ttk.Button(copy_frame, text="浏览...", command=self.browse_target_dir)
        self.browse_btn.pack(side=tk.LEFT, padx=5)
        
        self.copy_btn = ttk.Button(copy_frame, text="复制选中项", command=self.copy_selected)
        self.copy_btn.pack(side=tk.LEFT, padx=5)
        
        # 绑定双击事件
        self.result_tree.bind("<Double-1>", lambda e: self.copy_selected())
        
        # 底部信息
        bottom_frame = ttk.Frame(self.root, padding="10")
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        ttk.Label(
            bottom_frame, 
            text="小凯数码 专业数据恢复", 
            font=('SimHei', 12, 'bold'),
            foreground="#005fb8"
        ).pack()
        
        # 初始化按钮状态
        self.set_stop_button_state(False)
        self.set_start_button_state(True)
        self.set_stop_animation_button_state(False)
        
        # 绑定窗口大小变化事件
        self.root.bind("<Configure>", self.on_window_resize)
    
    def set_start_button_state(self, state):
        """设置开始按钮状态"""
        if state:
            self.search_btn.config(state=tk.NORMAL)
        else:
            self.search_btn.config(state=tk.DISABLED)
        self.root.update_idletasks()
    
    def set_stop_button_state(self, state):
        """设置停止搜索按钮状态"""
        if state:
            self.stop_btn.config(state=tk.NORMAL)
        else:
            self.stop_btn.config(state=tk.DISABLED)
        self.root.update_idletasks()
        
    def set_stop_animation_button_state(self, state):
        """设置停止动画按钮状态"""
        if state:
            self.stop_animation_btn.config(state=tk.NORMAL)
        else:
            self.stop_animation_btn.config(state=tk.DISABLED)
        self.root.update_idletasks()
    
    def format_file_size(self, size_bytes):
        """将字节大小转换为人类可读的格式"""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.2f} KB"
        elif size_bytes < 1024 * 1024 * 1024:
            return f"{size_bytes / (1024 * 1024):.2f} MB"
        else:
            return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"
    
    def format_modification_time(self, timestamp):
        """将时间戳转换为可读的日期时间格式"""
        return datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M:%S")
    
    def on_window_resize(self, event):
        """窗口大小变化时调整元素大小"""
        if event.widget == self.root and self.main_frame:
            main_width = self.main_frame.winfo_width()
            if main_width > 0:
                self.animation_text.config(width=main_width // 7)
    
    def update_drives(self):
        """更新可用驱动器列表"""
        for widget in self.drive_frame.winfo_children():
            widget.destroy()
            
        self.drive_vars = []
        drives = []
        
        if sys.platform.startswith('win'):
            # Windows系统
            for drive in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
                drive_path = f"{drive}:/"
                if os.path.exists(drive_path):
                    drives.append(drive_path)
        else:
            # 非Windows系统
            common_roots = ['/', '/home']
            for root in common_roots:
                if os.path.exists(root):
                    drives.append(root)
        
        for drive in drives:
            var = tk.BooleanVar()
            self.drive_vars.append((var, drive))
            
            chk = ttk.Checkbutton(
                self.drive_frame, 
                text=drive, 
                variable=var,
                command=self.update_selected_drives
            )
            chk.pack(side=tk.LEFT, padx=5, pady=2)
    
    def update_selected_drives(self):
        """更新选中的驱动器列表"""
        self.selected_drives = [drive for var, drive in self.drive_vars if var.get()]
    
    def select_all_drives(self):
        """全选驱动器"""
        for var, _ in self.drive_vars:
            var.set(True)
        self.update_selected_drives()
    
    def deselect_all_drives(self):
        """取消全选驱动器"""
        for var, _ in self.drive_vars:
            var.set(False)
        self.update_selected_drives()
    
    def start_search(self):
        """开始搜索指定格式的目录"""
        # 立即更新按钮状态
        self.set_start_button_state(False)
        self.set_stop_button_state(True)
        self.set_stop_animation_button_state(True)
        
        # 检查是否选择了驱动器
        self.update_selected_drives()
        if not self.selected_drives:
            messagebox.showwarning("警告", "请至少选择一个驱动器")
            # 恢复按钮状态
            self.set_start_button_state(True)
            self.set_stop_button_state(False)
            self.set_stop_animation_button_state(False)
            return
            
        # 获取选中的动画时长和结果显示延迟
        self.animation_duration = self.animation_duration_var.get()
        self.result_delay = self.result_delays.get(self.animation_duration, 120)  # 默认2分钟
            
        # 重置状态
        self.results = []
        self.result_tree.delete(*self.result_tree.get_children())
        self.searching = True
        self.stop_search = False
        self.delay_complete = False
        self.progress_var.set(0)
        self.status_var.set(f"开始重组...将在{self.result_delay//60}分钟后显示结果")
        
        # 禁用驱动器选择相关控件
        for widget in self.drive_frame.winfo_children():
            if isinstance(widget, ttk.Checkbutton):
                widget.config(state=tk.DISABLED)
        self.select_all_btn.config(state=tk.DISABLED)
        self.deselect_all_btn.config(state=tk.DISABLED)
        
        # 启动代码动画
        self.animation_running = True
        self.animation_end_time = time.time() + self.animation_duration
        self.animation_index = 0
        self.root.after(100, self.update_animation)
        
        # 启动搜索线程
        def search_wrapper():
            # 执行搜索
            self.search()
            # 等待延迟时间
            self.wait_for_delay()
            # 显示结果并完成
            self.root.after(0, self.display_results)
            self.root.after(0, self.search_complete)
        
        search_thread = threading.Thread(target=search_wrapper)
        search_thread.daemon = True
        search_thread.start()
        
        # 启动进度更新检查
        self.root.after(100, self.check_search_progress)
    
    def wait_for_delay(self):
        """等待指定的延迟时间后再显示结果"""
        if self.stop_search:
            return
            
        # 计算剩余时间并更新状态
        start_time = time.time()
        end_time = start_time + self.result_delay
        
        while time.time() < end_time and not self.stop_search:
            remaining = int(end_time - time.time())
            minutes = remaining // 60
            seconds = remaining % 60
            
            # 定期更新状态（每10秒）
            if remaining % 10 == 0:
                self.root.after(0, lambda: self.status_var.set(
                    f"重组完成，等待结果显示...剩余{minutes}分{seconds}秒"
                ))
            
            time.sleep(1)
        
        self.delay_complete = True
    
    def display_results(self):
        """显示搜索结果"""
        if self.stop_search:
            return
            
        # 清空现有结果
        self.result_tree.delete(*self.result_tree.get_children())
        
        # 显示所有结果
        for path, file, file_size, modified_time in self.results:
            formatted_size = self.format_file_size(file_size)
            formatted_time = self.format_modification_time(modified_time)
            self.result_tree.insert("", tk.END, values=(file, formatted_size, formatted_time))
        
        count = len(self.results)
        self.status_var.set(f"重组完成，共找到 {count} 个文件")
    
    def stop_searching(self):
        """停止重组"""
        if self.searching and not self.stop_search:
            self.stop_search = True
            self.status_var.set("正在停止重组...")
            self.root.update_idletasks()
    
    def stop_animation(self):
        """停止动画"""
        if self.animation_running:
            self.animation_running = False
            self.animation_text.config(state=tk.NORMAL)
            self.animation_text.delete('1.0', tk.END)
            self.animation_text.insert(tk.END, "已停止数据重组\n")
            self.animation_text.config(state=tk.DISABLED)
            self.set_stop_animation_button_state(False)
            
            # 如果延迟未完成，直接显示已找到的结果
            if not self.delay_complete and self.results:
                self.display_results()
                
            count = len(self.result_tree.get_children())
            self.status_var.set(f"已停止，共找到 {count} 个文件")
    
    def update_animation(self):
        """更新代码动画显示"""
        if not self.animation_running or self.stop_search:
            return
            
        # 检查动画是否已达到设定时长
        if time.time() >= self.animation_end_time and not self.searching and self.delay_complete:
            self.animation_text.config(state=tk.NORMAL)
            self.animation_text.delete('1.0', tk.END)
            self.animation_text.insert(tk.END, "重组完成，数据解析结束\n")
            self.animation_text.config(state=tk.DISABLED)
            self.set_stop_animation_button_state(False)
            return
            
        # 生成随机的0和1组成的代码行
        try:
            text_width = self.animation_text.winfo_width() // 7
            if text_width < 30:
                text_width = 30
        except:
            text_width = 90
            
        animation_chars = ['0', '1', ' ', '0', '1', ' ', '0', '1', ';', ':', '#', '@', '$', '%', '&']
        line = ''.join(random.choice(animation_chars) for _ in range(text_width))
        
        # 定期插入状态信息
        if random.random() < 0.3:
            line = self.animation_messages[self.animation_index]
            self.animation_index = (self.animation_index + 1) % len(self.animation_messages)
        
        # 更新动画区域
        self.animation_text.config(state=tk.NORMAL)
        self.animation_text.insert(tk.END, line + "\n")
        
        # 保持与文本框高度匹配的行数
        try:
            text_height = self.animation_text.winfo_height() // 15
            if text_height < 5:
                text_height = 5
                
            current_lines = int(self.animation_text.index('end-1c').split('.')[0])
            if current_lines > text_height:
                self.animation_text.delete('1.0', '2.0')
        except:
            current_lines = int(self.animation_text.index('end-1c').split('.')[0])
            if current_lines > 15:
                self.animation_text.delete('1.0', '2.0')
        
        self.animation_text.see(tk.END)
        self.animation_text.config(state=tk.DISABLED)
        
        # 继续动画
        if self.animation_running and not self.stop_search:
            self.root.after(random.randint(100, 300), self.update_animation)
    
    def search(self):
        r"""搜索指定格式的WPS缓存目录"""
        try:
            target_path_suffix = os.path.join("AppData", "Roaming", "kingsoft", "office6", "backup")
            found_dirs = []
            
            # 搜索每个选中的驱动器
            for drive in self.selected_drives:
                if self.stop_search:
                    break
                    
                users_dir = os.path.join(drive, "Users")
                if not os.path.exists(users_dir) or not os.path.isdir(users_dir):
                    continue
                    
                # 遍历用户目录
                for user in os.listdir(users_dir):
                    if self.stop_search:
                        break
                        
                    user_dir = os.path.join(users_dir, user)
                    if not os.path.isdir(user_dir):
                        continue
                        
                    target_dir = os.path.join(user_dir, target_path_suffix)
                    if os.path.exists(target_dir) and os.path.isdir(target_dir):
                        found_dirs.append(target_dir)
            
            # 计算总文件数用于进度显示
            total_files = 0
            for dir_path in found_dirs:
                if self.stop_search:
                    break
                for root, _, files in os.walk(dir_path):
                    if self.stop_search:
                        break
                    total_files += len(files)
            
            # 收集找到的文件（不立即显示）
            processed_files = 0
            for dir_path in found_dirs:
                if self.stop_search:
                    break
                    
                for root, _, files in os.walk(dir_path):
                    if self.stop_search:
                        break
                        
                    for file in files:
                        if self.stop_search:
                            break
                            
                        file_path = os.path.join(root, file)
                        try:
                            file_stats = os.stat(file_path)
                            file_size = file_stats.st_size
                            modified_time = file_stats.st_mtime
                            
                            self.results.append((root, file, file_size, modified_time))
                            processed_files += 1
                            
                            # 更新进度但不显示文件
                            if total_files > 0:
                                progress = (processed_files / total_files) * 100
                                self.root.after(0, self.update_progress, progress, 
                                               f"正在搜索...已找到 {processed_files} 个文件")
                        except Exception as e:
                            print(f"获取文件信息失败: {file_path}, 错误: {e}")
        
        except Exception as e:
            self.root.after(0, self.show_error, str(e))
    
    def update_progress(self, value, status):
        """更新进度条和状态"""
        self.progress_var.set(value)
        self.status_var.set(status)
    
    def show_error(self, message):
        """显示错误信息"""
        messagebox.showerror("错误", f"搜索过程中发生错误:\n{message}")
    
    def search_complete(self):
        """搜索完成后的清理工作"""
        self.searching = False
        self.stop_search = False
        self.progress_var.set(100)
        
        # 恢复按钮状态
        self.set_start_button_state(True)
        self.set_stop_button_state(False)
        
        # 处理动画状态
        if time.time() < self.animation_end_time and self.animation_running:
            self.status_var.set(f"重组完成，共找到 {len(self.results)} 个文件")
        else:
            self.animation_running = False
            self.animation_text.config(state=tk.NORMAL)
            self.animation_text.delete('1.0', tk.END)
            self.animation_text.insert(tk.END, "重组完成，数据解析结束\n")
            self.animation_text.config(state=tk.DISABLED)
            self.set_stop_animation_button_state(False)
        
        # 启用驱动器选择相关控件
        for widget in self.drive_frame.winfo_children():
            if isinstance(widget, ttk.Checkbutton):
                widget.config(state=tk.NORMAL)
        self.select_all_btn.config(state=tk.NORMAL)
        self.deselect_all_btn.config(state=tk.NORMAL)
    
    def check_search_progress(self):
        """检查搜索进度并更新UI"""
        if self.searching:
            self.root.after(100, self.check_search_progress)
        else:
            if time.time() >= self.animation_end_time and self.animation_running:
                self.animation_running = False
                self.animation_text.config(state=tk.NORMAL)
                self.animation_text.delete('1.0', tk.END)
                self.animation_text.insert(tk.END, "扫描完成，数据解析结束\n")
                self.animation_text.config(state=tk.DISABLED)
                self.set_stop_animation_button_state(False)
            elif self.animation_running:
                self.root.after(1000, self.check_search_progress)
    
    def browse_target_dir(self):
        """浏览目标目录"""
        dir_path = filedialog.askdirectory(title="选择目标目录")
        if dir_path:
            self.target_dir_var.set(dir_path)
    
    def copy_selected(self):
        """复制选中的文件到目标位置"""
        selected_items = self.result_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要复制的文件")
            return
            
        target_dir = self.target_dir_var.get()
        if not target_dir or not os.path.exists(target_dir):
            messagebox.showwarning("警告", "请选择有效的目标目录")
            return
        
        try:
            copied_count = 0
            for item in selected_items:
                index = int(self.result_tree.index(item))
                path, filename, _, _ = self.results[index]
                source_path = os.path.join(path, filename)
                dest_path = os.path.join(target_dir, filename)
                
                if os.path.exists(dest_path):
                    if not messagebox.askyesno("确认", f"文件 '{filename}' 已存在，是否覆盖?"):
                        continue
                
                shutil.copy2(source_path, dest_path)
                copied_count += 1
            
            messagebox.showinfo("成功", f"共成功复制 {copied_count}/{len(selected_items)} 个文件到:\n{target_dir}")
            
        except Exception as e:
            messagebox.showerror("复制失败", f"复制过程中发生错误:\n{str(e)}")
    
    def select_all_results(self):
        """全选搜索结果"""
        for item in self.result_tree.selection():
            self.result_tree.selection_remove(item)
        
        for item in self.result_tree.get_children():
            self.result_tree.selection_add(item)
        
        self.update_selection_count()
    
    def invert_selection(self):
        """反选搜索结果"""
        current_selection = set(self.result_tree.selection())
        all_items = set(self.result_tree.get_children())
        
        for item in current_selection:
            self.result_tree.selection_remove(item)
        
        for item in all_items - current_selection:
            self.result_tree.selection_add(item)
        
        self.update_selection_count()
    
    def update_selection_count(self, event=None):
        """更新选择计数显示"""
        count = len(self.result_tree.selection())
        self.selection_count_var.set(f"已选择: {count} 项")

if __name__ == "__main__":
    root = tk.Tk()
    app = TableRepairSoftware(root)
    root.mainloop()
    