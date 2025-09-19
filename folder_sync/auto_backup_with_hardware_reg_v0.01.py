import os
import shutil
import webbrowser  # 新增：用于打开浏览器访问超链接
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
import wmi  # 需要安装: pip install wmi

class HardwareRegistrationSystem:
    """基于硬件信息的注册系统（v0.01版本 - 稳定版）"""
    def __init__(self):
        # 固定盐值（与注册码生成器严格同步，不可修改）
        self._fixed_salt = "BackupTool_v2.2_2024_Authorized"
        self.debug_info = []  # 保留基础调试信息存储（内部使用，不对外展示）
        
    def get_motherboard_serial(self):
        """获取主板序列号（硬件绑定核心）"""
        try:
            c = wmi.WMI()
            for board in c.Win32_BaseBoard():
                serial = board.SerialNumber.strip().upper()
                return serial if serial else self._generate_fallback_serial()
        except Exception as e:
            return self._generate_fallback_serial()
    
    def _generate_fallback_serial(self):
        """硬件信息获取失败时的备用序列号生成"""
        return 'FALLBACK-' + ''.join(random.choices(string.ascii_uppercase + string.digits, k=12))
    
    def generate_machine_code(self, motherboard_serial=None):
        """生成硬件唯一机器码（格式：XXXX-XXXX-XXXX-XXXX-XXXX）"""
        if not motherboard_serial:
            motherboard_serial = self.get_motherboard_serial()
        
        # SHA-256哈希处理确保唯一性
        hash_obj = hashlib.sha256(motherboard_serial.encode('utf-8'))
        hash_hex = hash_obj.hexdigest()
        hash_part = hash_hex[:20]  # 取前20位保证长度统一
        return '-'.join([hash_part[i:i+4] for i in range(0, 20, 4)])
    
    def generate_registration_key(self, machine_code):
        """生成与机器码绑定的注册码（v0.01稳定算法）"""
        if not machine_code or len(machine_code) != 24 or machine_code.count('-') != 4:
            raise ValueError("机器码格式错误，需为24位（含4个横线）")
        
        # 清理机器码并组合基础字符串
        cleaned_machine_code = machine_code.replace('-', '')
        base_str = cleaned_machine_code + self._fixed_salt
        
        # 哈希计算与注册码格式化
        hash_obj = hashlib.sha256(base_str.encode('utf-8'))
        hash_hex = hash_obj.hexdigest().upper()
        parts = [
            hash_hex[0:4],   # 段1：前4位
            hash_hex[8:12],  # 段2：第9-12位
            hash_hex[16:20], # 段3：第17-20位
            hash_hex[24:28], # 段4：第25-28位
            hash_hex[32:36]  # 段5：第33-36位
        ]
        return '-'.join(parts)
    
    def verify_registration_key(self, machine_code, registration_key):
        """验证注册码有效性（与生成算法严格同步）"""
        # 基础格式校验
        if not registration_key or len(registration_key) != 24 or registration_key.count('-') != 4:
            return False
        if not machine_code or len(machine_code) != 24 or machine_code.count('-') != 4:
            return False
        
        # 生成正确注册码并比对
        try:
            correct_key = self.generate_registration_key(machine_code)
            return correct_key == registration_key
        except:
            return False


class BackupTool:
    """智能自动备份工具（v0.01版本 - 稳定版）"""
    def __init__(self, root, start_minimized=False):
        # 初始化核心组件
        self.reg_system = HardwareRegistrationSystem()
        self.REG_KEY_PATH = r"Software\BackupTool"  # 注册表存储路径
        self.REGISTRATION_KEY = "RegistrationKey"  # 注册码存储键名
        self.MACHINE_CODE_KEY = "MachineCode"      # 机器码存储键名
        self.TRIAL_DAYS = 7  # 试用期天数
        self.is_registered = False
        self.trial_end_date = None
        self.machine_code = None
        
        # 主窗口基础配置
        self.root = root
        self.root.title("智能自动备份工具 v0.01")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        self.center_window()
        
        # 初始化状态与配置
        self.log_text = None
        self.check_license_status()  # 优先检查授权状态
        
        # 未注册且试用期过期时强制注册
        if not self.is_registered and (self.trial_end_date is None or datetime.now() > self.trial_end_date):
            if not self.show_registration_window():
                sys.exit(0)
        
        # 配置文件路径（兼容打包与开发环境）
        self.config_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)) if not getattr(sys, 'frozen', False) 
            else os.path.dirname(sys.executable), 
            "backup_config.ini"
        )
        self.icon_path = self.find_icon_file()
        self.load_config()
        self.start_minimized = start_minimized
        
        # 样式配置（确保中文显示正常）
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("SimHei", 10))
        self.style.configure("TButton", font=("SimHei", 10))
        self.style.configure("TEntry", font=("SimHei", 10))
        self.style.configure("Footer.TLabel", font=("SimHei", 9))
        self.style.configure("Link.TLabel", font=("SimHei", 9, "underline"), foreground="blue")  # 新增：超链接样式
        
        # 核心功能变量
        self.source_path = tk.StringVar(value=self.config.get("Paths", "source", fallback=""))
        self.dest_path = tk.StringVar(value=self.config.get("Paths", "dest", fallback=""))
        self.monitor_interval = tk.IntVar(value=self.config.getint("Settings", "interval", fallback=60))
        self.last_monitoring_state = self.config.getboolean("Settings", "monitoring", fallback=False)
        self.backup_running = False
        self.monitoring = False
        self.file_hashes = {}  # 用于检测文件变化的哈希存储
        self.tray_icon = None  # 系统托盘图标
        
        # 构建UI与初始化
        self.create_widgets()
        self.set_window_icon()
        self.update_startup_button()
        self.show_trial_info()  # 显示试用期信息
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)  # 重写关闭事件
        self.post_init()  # 初始化后处理（最小化/自动启动等）
    
    # ------------------------------
    # 新增：超链接点击事件
    # ------------------------------
    def open_url(self, url):
        """点击超链接时打开浏览器访问指定网址"""
        try:
            webbrowser.open_new(url)
            self.log(f"已打开网址: {url}")
        except Exception as e:
            self.log(f"打开网址失败: {str(e)}")
            messagebox.showerror("错误", f"无法打开网址: {str(e)}")
    
    # ------------------------------
    # 核心功能：授权与注册相关
    # ------------------------------
    def check_license_status(self):
        """检查软件授权状态（注册/试用期）"""
        # 生成当前机器码
        self.machine_code = self.reg_system.generate_machine_code()
        
        try:
            # 读取注册表中的注册信息
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.REG_KEY_PATH, 0, winreg.KEY_READ)
            registered_value, _ = winreg.QueryValueEx(key, self.REGISTRATION_KEY)
            winreg.CloseKey(key)
            
            # 验证注册码有效性
            if self.reg_system.verify_registration_key(self.machine_code, registered_value):
                self.is_registered = True
                return
        except FileNotFoundError:
            pass  # 未找到注册信息，进入试用期检查
        except Exception as e:
            self.log(f"授权检查警告: {str(e)}")
        
        # 检查试用期状态
        self.check_trial_period()
    
    def check_trial_period(self):
        """管理试用期逻辑（防篡改/机器绑定）"""
        try:
            # 读取已存储的试用期信息
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.REG_KEY_PATH, 0, winreg.KEY_READ)
            trial_end_str, _ = winreg.QueryValueEx(key, "TrialEndDate")
            stored_machine_code, _ = winreg.QueryValueEx(key, self.MACHINE_CODE_KEY)
            winreg.CloseKey(key)
            
            # 机器码不匹配时重置试用期（防复制试用期信息）
            if stored_machine_code != self.machine_code:
                raise Exception("机器码不匹配，重置试用期")
            
            self.trial_end_date = datetime.strptime(trial_end_str, "%Y-%m-%d")
        except:
            # 首次运行或异常时设置新试用期
            self.trial_end_date = datetime.now() + timedelta(days=self.TRIAL_DAYS)
            # 保存试用期信息到注册表
            try:
                key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.REG_KEY_PATH)
                winreg.SetValueEx(key, "TrialEndDate", 0, winreg.REG_SZ, self.trial_end_date.strftime("%Y-%m-%d"))
                winreg.SetValueEx(key, self.MACHINE_CODE_KEY, 0, winreg.REG_SZ, self.machine_code)
                winreg.CloseKey(key)
            except Exception as e:
                self.log(f"试用期信息保存失败: {str(e)}")
    
    def show_registration_window(self):
        """注册窗口（v0.01简洁版）"""
        reg_window = tk.Toplevel(self.root)
        reg_window.title("软件注册 - 智能自动备份工具 v0.01")
        reg_window.geometry("600x350")
        reg_window.resizable(False, False)
        reg_window.transient(self.root)
        reg_window.grab_set()  # 模态窗口（强制聚焦）
        self.center_window(reg_window)
        
        # 注册窗口内容框架
        main_frame = ttk.Frame(reg_window, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        ttk.Label(main_frame, text="软件注册", font=("SimHei", 16, "bold")).grid(row=0, column=0, columnspan=2, pady=10)
        
        # 机器码显示与复制
        ttk.Label(main_frame, text="您的机器码（用于获取注册码）:", font=("SimHei", 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
        machine_code_frame = ttk.Frame(main_frame)
        machine_code_frame.grid(row=2, column=0, columnspan=2, pady=5, sticky=tk.W)
        
        machine_code_var = tk.StringVar(value=self.machine_code)
        machine_code_entry = ttk.Entry(machine_code_frame, textvariable=machine_code_var, width=50, state="readonly")
        machine_code_entry.pack(side=tk.LEFT, padx=5)
        
        def copy_machine_code():
            reg_window.clipboard_clear()
            reg_window.clipboard_append(self.machine_code)
            messagebox.showinfo("提示", "机器码已复制到剪贴板")
        ttk.Button(machine_code_frame, text="复制", command=copy_machine_code).pack(side=tk.LEFT)
        
        # 试用期提示
        trial_frame = ttk.Frame(main_frame)
        trial_frame.grid(row=3, column=0, columnspan=2, pady=5, sticky=tk.W)
        if self.trial_end_date is not None and datetime.now() > self.trial_end_date:
            ttk.Label(trial_frame, text="试用期已结束，请输入注册码以继续使用。", font=("SimHei", 10), foreground="red").pack(anchor=tk.W)
        else:
            days_left = (self.trial_end_date - datetime.now()).days + 1
            ttk.Label(trial_frame, text=f"试用期剩余 {days_left} 天，请输入注册码以解锁全部功能。", font=("SimHei", 10)).pack(anchor=tk.W)
        
        # 注册码输入
        ttk.Label(main_frame, text="注册码:", font=("SimHei", 10)).grid(row=4, column=0, sticky=tk.W, pady=5)
        reg_key_frame = ttk.Frame(main_frame)
        reg_key_frame.grid(row=5, column=0, columnspan=2, pady=5, sticky=tk.W)
        
        reg_key_var = tk.StringVar()
        ttk.Entry(reg_key_frame, textvariable=reg_key_var, width=50).pack(side=tk.LEFT, padx=5)
        
        # 状态提示
        status_var = tk.StringVar(value="")
        ttk.Label(main_frame, textvariable=status_var, foreground="red").grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # 功能按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=7, column=0, columnspan=2, pady=10)
        result = [False]  # 用于内部函数修改结果
        
        def register():
            """执行注册逻辑"""
            key = reg_key_var.get().strip()
            status_var.set("正在验证注册码...")
            reg_window.update_idletasks()
            
            # 格式校验
            if len(key) != 24 or key.count('-') != 4:
                status_var.set(f"注册码格式错误（需24位，含4个横线）")
                return
            
            # 验证与保存
            if self.reg_system.verify_registration_key(self.machine_code, key):
                try:
                    key_path = winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.REG_KEY_PATH)
                    winreg.SetValueEx(key_path, self.REGISTRATION_KEY, 0, winreg.REG_SZ, key)
                    winreg.CloseKey(key_path)
                    
                    self.is_registered = True
                    self.license_status_var.set("已注册")
                    messagebox.showinfo("成功", "注册成功！感谢您的支持～")
                    result[0] = True
                    reg_window.destroy()
                except Exception as e:
                    status_var.set(f"注册信息保存失败: {str(e)}")
            else:
                status_var.set("注册码无效，请核对机器码后重试")
        
        def continue_trial():
            """继续试用（仅试用期内可用）"""
            if self.trial_end_date is None or datetime.now() > self.trial_end_date:
                messagebox.showerror("错误", "试用期已结束，请注册后使用")
                return
            result[0] = True
            reg_window.destroy()
        
        # 按钮布局
        ttk.Button(btn_frame, text="注册", command=register).pack(side=tk.LEFT, padx=10)
        if self.trial_end_date is not None and datetime.now() <= self.trial_end_date:
            ttk.Button(btn_frame, text="继续试用", command=continue_trial).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="退出", command=reg_window.destroy).pack(side=tk.LEFT, padx=10)
        
        self.root.wait_window(reg_window)
        return result[0]
    
    def show_trial_info(self):
        """显示试用期信息（未注册时）"""
        if not self.is_registered:
            days_left = (self.trial_end_date - datetime.now()).days + 1
            if days_left > 0:
                self.log(f"试用期剩余: {days_left} 天")
            # 添加注册按钮
            reg_button = ttk.Button(self.root, text="注册软件", command=self.show_registration_window)
            reg_button.place(relx=0.95, rely=0.02, anchor=tk.NE)
            # 试用期提示标签
            if days_left > 0:
                trial_label = ttk.Label(
                    self.root, 
                    text=f"试用期剩余: {days_left} 天", 
                    foreground="orange"
                )
                trial_label.place(relx=0.8, rely=0.02, anchor=tk.NE)
        else:
            self.log("软件已注册，感谢您的支持！")

    # ------------------------------
    # 核心功能：备份与监控相关
    # ------------------------------
    def calculate_initial_hashes(self):
        """计算源文件夹初始文件哈希（用于检测变化）"""
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
    
    def get_file_hash(self, file_path):
        """计算单个文件的MD5哈希值"""
        try:
            hasher = hashlib.md5()
            with open(file_path, 'rb') as f:
                while chunk := f.read(4096):
                    hasher.update(chunk)
            return hasher.hexdigest()
        except Exception as e:
            self.log(f"计算文件哈希失败 {file_path}: {str(e)}")
            return None
    
    def check_for_changes(self):
        """检查源文件夹是否有文件变化（新增/修改/删除）"""
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
                
                if rel_path not in self.file_hashes:
                    new_files.append(rel_path)
                elif current_hashes[rel_path] != self.file_hashes[rel_path]:
                    modified_files.append(rel_path)
        
        # 检查已删除文件
        for rel_path in self.file_hashes:
            if rel_path not in current_hashes:
                deleted_files.append(rel_path)
        
        self.file_hashes = current_hashes
        return len(new_files) + len(modified_files) + len(deleted_files) > 0, new_files, modified_files, deleted_files
    
    def copy_files(self, src, dest, specific_files=None):
        """复制文件（支持完整备份/增量备份）"""
        try:
            # 处理文件列表（去重）
            if specific_files is None:
                items = os.listdir(src)
                total_items = len(items)
            else:
                items = list(set([os.path.normpath(file).split(os.sep)[0] for file in specific_files]))
                total_items = len(items)
            
            processed_items = 0
            for item in items:
                if not self.backup_running:
                    self.log("备份已取消")
                    return False
                
                src_path = os.path.join(src, item)
                dest_path = os.path.join(dest, item)
                
                # 仅处理指定文件（增量备份时）
                process_flag = True
                if specific_files is not None:
                    process_flag = any(os.path.normpath(file).startswith(os.path.normpath(item)) for file in specific_files)
                
                if process_flag:
                    if os.path.isfile(src_path):
                        shutil.copy2(src_path, dest_path)
                        self.log(f"已复制: {src_path}")
                    elif os.path.isdir(src_path):
                        os.makedirs(dest_path, exist_ok=True)
                        self.log(f"已创建目录: {dest_path}")
                        # 增量备份时递归处理子文件
                        if specific_files is not None:
                            sub_files = [f.split(os.sep, 1)[1] for f in specific_files 
                                        if os.path.normpath(f).startswith(os.path.normpath(item)) and len(f.split(os.sep, 1)) > 1]
                            if sub_files and not self.copy_files(src_path, dest_path, sub_files):
                                return False
                        else:
                            if not self.copy_files(src_path, dest_path):
                                return False
                
                # 更新进度
                processed_items += 1
                self.update_progress((processed_items / total_items) * 100)
            return True
        except Exception as e:
            self.log(f"文件复制失败: {str(e)}")
            return False
    
    def backup_thread(self, specific_files=None):
        """备份线程（避免UI冻结）"""
        source = self.source_path.get()
        dest = self.dest_path.get()
        
        # 路径校验
        if not source or not os.path.isdir(source):
            self.log("错误: 源文件夹无效")
            self.backup_running = False
            self.backup_btn.config(text="立即备份", state=tk.NORMAL)
            self.update_status("就绪")
            return
        if not dest or not os.path.isdir(dest):
            self.log("错误: 目标文件夹无效")
            self.backup_running = False
            self.backup_btn.config(text="立即备份", state=tk.NORMAL)
            self.update_status("就绪")
            return
        
        # 执行备份
        self.update_progress(0)
        self.log("开始完整备份..." if specific_files is None else "检测到变化，开始增量备份...")
        self.log(f"源: {source} | 目标: {dest}")
        start_time = time.time()
        
        success = self.copy_files(source, dest, specific_files)
        # 同步删除目标文件夹中已删除的文件（完整备份时）
        if specific_files is None and hasattr(self, 'deleted_files') and self.deleted_files:
            self.delete_files(dest, self.deleted_files)
        
        # 备份结果处理
        elapsed = time.time() - start_time
        if success:
            self.log(f"备份完成！耗时: {elapsed:.2f} 秒")
        else:
            self.log(f"备份中断！耗时: {elapsed:.2f} 秒")
        
        # 重置状态
        self.backup_running = False
        self.backup_btn.config(text="立即备份", state=tk.NORMAL)
        self.update_progress(100)
        self.update_status("就绪")
    
    def start_backup(self):
        """启动备份（UI触发）"""
        if not self.backup_running:
            self.backup_running = True
            self.backup_btn.config(text="取消备份", state=tk.DISABLED)
            self.update_status("备份中...")
            threading.Thread(target=self.backup_thread, daemon=True).start()
    
    def monitoring_thread(self):
        """监控线程（定期检查文件变化）"""
        while self.monitoring:
            self.update_status(f"监控中（间隔: {self.monitor_interval.get()}秒）")
            # 检查文件变化
            has_changes, new_files, modified_files, deleted_files = self.check_for_changes()
            if has_changes:
                self.log(f"检测到变化 - 新增: {len(new_files)} | 修改: {len(modified_files)} | 删除: {len(deleted_files)}")
                self.deleted_files = deleted_files
                # 启动增量备份
                if not self.backup_running:
                    self.backup_running = True
                    self.backup_btn.config(text="取消备份", state=tk.DISABLED)
                    self.update_status("增量备份中...")
                    threading.Thread(target=self.backup_thread, args=((new_files + modified_files),), daemon=True).start()
            # 等待间隔（支持中途停止）
            for _ in range(self.monitor_interval.get()):
                if not self.monitoring:
                    break
                time.sleep(1)
    
    def toggle_monitoring(self):
        """切换监控状态（开始/停止）"""
        if not self.monitoring:
            # 启动监控前校验
            source = self.source_path.get()
            dest = self.dest_path.get()
            if not source or not os.path.isdir(source):
                messagebox.showerror("错误", "请选择有效的源文件夹")
                return
            if not dest or not os.path.isdir(dest):
                messagebox.showerror("错误", "请选择有效的目标文件夹")
                return
            interval = self.monitor_interval.get()
            if interval < 10 and not messagebox.askyesno("警告", "监控间隔过短可能影响性能，是否继续？"):
                return
            
            # 启动监控
            self.monitoring = True
            self.update_monitor_button_text()
            self.calculate_initial_hashes()
            self.log(f"开始监控，间隔 {interval} 秒")
            threading.Thread(target=self.monitoring_thread, daemon=True).start()
        else:
            # 停止监控
            self.monitoring = False
            self.update_monitor_button_text()
            self.log("已停止监控")
            self.update_status("就绪")

    # ------------------------------
    # 辅助功能：UI与配置相关
    # ------------------------------
    def create_widgets(self):
        """创建主窗口UI组件"""
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 1. 源文件夹选择
        ttk.Label(main_frame, text="源文件夹:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.source_path, width=50).grid(row=0, column=1, pady=5)
        ttk.Button(main_frame, text="浏览...", command=self.select_source).grid(row=0, column=2, padx=5, pady=5)
        
        # 2. 目标文件夹选择
        ttk.Label(main_frame, text="目标文件夹:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.dest_path, width=50).grid(row=1, column=1, pady=5)
        ttk.Button(main_frame, text="浏览...", command=self.select_dest).grid(row=1, column=2, padx=5, pady=5)
        
        # 3. 监控间隔设置
        ttk.Label(main_frame, text="监控间隔(秒):").grid(row=2, column=0, sticky=tk.W, pady=5)
        interval_frame = ttk.Frame(main_frame)
        interval_frame.grid(row=2, column=1, sticky=tk.W, pady=5)
        ttk.Entry(interval_frame, textvariable=self.monitor_interval, width=10).pack(side=tk.LEFT)
        ttk.Label(interval_frame, text="建议≥60秒").pack(side=tk.LEFT, padx=5)
        
        # 4. 功能按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=10)
        self.backup_btn = ttk.Button(button_frame, text="立即备份", command=self.start_backup)
        self.backup_btn.pack(side=tk.LEFT, padx=5)
        self.monitor_btn = ttk.Button(button_frame, text="开始监控", command=self.toggle_monitoring)
        self.monitor_btn.pack(side=tk.LEFT, padx=5)
        self.startup_btn = ttk.Button(button_frame, text="添加开机启动", command=self.toggle_startup)
        self.startup_btn.pack(side=tk.LEFT, padx=5)
        self.save_btn = ttk.Button(button_frame, text="保存配置", command=self.save_config)
        self.save_btn.pack(side=tk.LEFT, padx=5)
        
        # 5. 状态显示区域
        ttk.Label(main_frame, text="当前状态:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(main_frame, textvariable=self.status_var).grid(row=4, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(main_frame, text="授权状态:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.license_status_var = tk.StringVar(value="已注册" if self.is_registered else "试用中")
        ttk.Label(main_frame, textvariable=self.license_status_var).grid(row=5, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(main_frame, text="开机启动:").grid(row=6, column=0, sticky=tk.W, pady=5)
        self.startup_status_var = tk.StringVar(value="未设置")
        ttk.Label(main_frame, textvariable=self.startup_status_var).grid(row=6, column=1, sticky=tk.W, pady=5)
        
        # 6. 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=7, column=0, columnspan=3, sticky=tk.EW, pady=10)
        
        # 7. 日志区域
        ttk.Label(main_frame, text="操作日志:").grid(row=8, column=0, sticky=tk.NW, pady=5)
        self.log_text = tk.Text(main_frame, height=12, width=70)
        self.log_text.grid(row=8, column=1, columnspan=2, pady=5, sticky=tk.NSEW)
        scrollbar = ttk.Scrollbar(main_frame, command=self.log_text.yview)
        scrollbar.grid(row=8, column=3, sticky=tk.NS)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 8. 页脚信息（新增超链接）
        footer_frame = ttk.Frame(main_frame)
        footer_frame.grid(row=9, column=0, columnspan=3, pady=10, sticky="nsew")
        
        # 普通文本：版本与联系方式
        ttk.Label(footer_frame, text="智能自动备份工具 v0.01 | QQ: 88179096", style="Footer.TLabel").pack(side=tk.LEFT, padx=10)
        
        # 超链接：网址（可点击）
        url_label = ttk.Label(footer_frame, text="www.itvip.com.cn", style="Link.TLabel")
        url_label.pack(side=tk.LEFT, padx=10)
        # 绑定点击事件：点击时打开浏览器访问网址
        url_label.bind("<Button-1>", lambda e: self.open_url("http://www.itvip.com.cn"))
        # 鼠标悬浮时显示手型光标（增强交互体验）
        url_label.bind("<Enter>", lambda e: self.root.config(cursor="hand2"))
        url_label.bind("<Leave>", lambda e: self.root.config(cursor="arrow"))
        
        # 网格权重（支持窗口伸缩）
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(8, weight=1)
    
    def select_source(self):
        """选择源文件夹"""
        path = filedialog.askdirectory(title="选择需要备份的源文件夹")
        if path:
            self.source_path.set(path)
            self.calculate_initial_hashes()
    
    def select_dest(self):
        """选择目标文件夹"""
        path = filedialog.askdirectory(title="选择备份存储的目标文件夹")
        if path:
            self.dest_path.set(path)
    
    def log(self, message):
        """添加操作日志"""
        if self.log_text is None:
            print(f"日志: {message}")
            return
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
    
    def update_progress(self, value):
        """更新进度条"""
        self.progress_var.set(value)
        self.root.update_idletasks()
    
    def update_status(self, status):
        """更新状态文本"""
        self.status_var.set(status)
        self.root.update_idletasks()
        # 同步更新托盘标题
        if self.tray_icon:
            self.tray_icon.title = f"智能自动备份工具 v0.01 - {status}"
    
    def update_monitor_button_text(self):
        """更新监控按钮文本"""
        self.monitor_btn.config(text="停止监控" if self.monitoring else "开始监控")

    # ------------------------------
    # 辅助功能：系统与配置相关
    # ------------------------------
    def load_config(self):
        """加载配置文件（backup_config.ini）"""
        self.config = configparser.ConfigParser()
        if os.path.exists(self.config_path):
            self.config.read(self.config_path, encoding="utf-8")
            # 确保配置 sections 存在
            if "Paths" not in self.config:
                self.config.add_section("Paths")
            if "Settings" not in self.config:
                self.config.add_section("Settings")
        else:
            # 初始化默认配置
            self.config.add_section("Paths")
            self.config.add_section("Settings")
    
    def save_config(self):
        """保存配置文件（含监控状态）"""
        try:
            self.config.set("Paths", "source", self.source_path.get())
            self.config.set("Paths", "dest", self.dest_path.get())
            self.config.set("Settings", "interval", str(self.monitor_interval.get()))
            self.config.set("Settings", "monitoring", str(self.monitoring))
            
            with open(self.config_path, 'w', encoding="utf-8") as f:
                self.config.write(f)
            
            self.log(f"配置已保存到: {self.config_path}")
            messagebox.showinfo("成功", "配置保存完成！")
        except Exception as e:
            self.log(f"配置保存失败: {str(e)}")
            messagebox.showerror("错误", f"保存配置失败: {str(e)}")
    
    def find_icon_file(self):
        """查找程序图标（backup.ico）"""
        possible_paths = [
            os.path.join(os.path.dirname(os.path.abspath(__file__)) if not getattr(sys, 'frozen', False) else os.path.dirname(sys.executable), "backup.ico"),
            os.path.join(sys._MEIPASS, "backup.ico") if (getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')) else "",
            os.path.join(os.path.expanduser("~"), "Desktop", "backup.ico"),
            os.path.join(os.getenv('APPDATA', ""), "BackupTool", "backup.ico")
        ]
        
        for path in possible_paths:
            if path and os.path.exists(path) and os.path.isfile(path):
                return path
        
        self.log("未找到backup.ico图标文件，使用默认图标")
        return None
    
    def set_window_icon(self):
        """设置窗口图标"""
        if self.icon_path and os.path.exists(self.icon_path):
            try:
                self.root.iconbitmap(self.icon_path)
                self.log(f"已加载图标: {self.icon_path}")
            except Exception as e:
                self.log(f"图标设置失败: {str(e)}")
    
    def center_window(self, window=None):
        """窗口居中显示"""
        target_window = window or self.root
        target_window.update_idletasks()
        width = target_window.winfo_width()
        height = target_window.winfo_height()
        x = (target_window.winfo_screenwidth() // 2) - (width // 2)
        y = (target_window.winfo_screenheight() // 2) - (height // 2)
        target_window.geometry(f"{width}x{height}+{x}+{y}")
    
    def post_init(self):
        """初始化后处理（最小化/自动启动）"""
        if self.start_minimized:
            self.root.withdraw()  # 隐藏窗口到托盘
            self.create_tray_icon()
            self.auto_start_monitoring()  # 自动启动监控（需配置有效）
        else:
            self.root.deiconify()
            # 恢复上次监控状态
            if self.last_monitoring_state and messagebox.askyesno("恢复监控", "检测到上次退出时正在监控，是否恢复？"):
                self.toggle_monitoring()
    
    def auto_start_monitoring(self):
        """开机启动后自动监控（延迟5秒确保系统就绪）"""
        source = self.source_path.get()
        dest = self.dest_path.get()
        if source and os.path.isdir(source) and dest and os.path.isdir(dest):
            threading.Timer(5, self.start_monitoring_silently).start()
        else:
            self.log("路径配置无效，无法自动启动监控")
    
    def start_monitoring_silently(self):
        """无提示启动监控（用于自动启动）"""
        try:
            self.monitoring = True
            self.update_monitor_button_text()
            self.calculate_initial_hashes()
            self.log(f"自动监控启动，间隔 {self.monitor_interval.get()} 秒")
            threading.Thread(target=self.monitoring_thread, daemon=True).start()
        except Exception as e:
            self.log(f"自动监控启动失败: {str(e)}")

    # ------------------------------
    # 辅助功能：开机启动与托盘相关
    # ------------------------------
    def toggle_startup(self):
        """添加/移除开机启动"""
        if self.is_in_startup():
            if self.remove_from_startup():
                self.log("已从开机启动中移除")
            else:
                self.log("移除开机启动失败")
        else:
            if self.add_to_startup():
                self.log("已添加到开机启动")
            else:
                self.log("添加开机启动失败")
        self.update_startup_button()
    
    def add_to_startup(self):
        """添加到Windows开机启动（带最小化参数）"""
        try:
            exe_path = self.get_executable_path()
            exe_name = os.path.basename(exe_path)
            command = f'"{exe_path}" --minimized'  # 最小化启动参数
            
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(key, exe_name, 0, winreg.REG_SZ, command)
            winreg.CloseKey(key)
            return True
        except Exception as e:
            self.log(f"添加开机启动失败: {str(e)}")
            return False
    
    def remove_from_startup(self):
        """从Windows开机启动中移除"""
        try:
            exe_path = self.get_executable_path()
            exe_name = os.path.basename(exe_path)
            
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
            winreg.DeleteValue(key, exe_name)
            winreg.CloseKey(key)
            return True
        except Exception as e:
            self.log(f"移除开机启动失败: {str(e)}")
            return False
    
    def is_in_startup(self):
        """检查是否已在开机启动中"""
        try:
            exe_path = self.get_executable_path()
            exe_name = os.path.basename(exe_path)
            
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
            winreg.QueryValueEx(key, exe_name)
            winreg.CloseKey(key)
            return True
        except FileNotFoundError:
            return False
        except Exception as e:
            self.log(f"开机启动检查失败: {str(e)}")
            return False
    
    def get_executable_path(self):
        """获取程序路径（兼容打包/开发环境）"""
        return sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
    
    def create_tray_icon(self):
        """创建系统托盘图标（支持隐藏窗口操作）"""
        try:
            # 加载图标（无图标时生成默认图标）
            if self.icon_path and os.path.exists(self.icon_path):
                icon = Image.open(self.icon_path)
            else:
                # 生成默认蓝色方块图标
                icon = Image.new("RGB", (64, 64), "blue")
                dc = ImageDraw.Draw(icon)
                dc.rectangle((16, 16, 48, 48), fill="white")
                dc.rectangle((32, 16, 48, 48), fill="blue")
            
            # 托盘菜单
            menu_items = []
            # 未注册时添加注册选项
            if not self.is_registered:
                days_left = (self.trial_end_date - datetime.now()).days + 1
                menu_items.append(pystray.MenuItem("注册软件", self.show_registration_window))
                if days_left > 0:
                    menu_items.append(pystray.MenuItem(f"试用期剩余: {days_left} 天", None, enabled=False))
                menu_items.append(pystray.Menu.SEPARATOR)
            
            # 核心菜单选项
            menu_items.extend([
                pystray.MenuItem("显示窗口", self.show_window),
                pystray.MenuItem("立即备份", self.tray_start_backup),
                pystray.MenuItem("退出程序", self.exit_program)
            ])
            
            # 创建托盘图标（单独线程运行）
            self.tray_icon = pystray.Icon("backup_tool_v0.01", icon, "智能自动备份工具 v0.01", tuple(menu_items))
            threading.Thread(target=self.tray_icon.run, daemon=True).start()
        except Exception as e:
            self.log(f"托盘图标创建失败: {str(e)}")
    
    def show_window(self):
        """从托盘显示主窗口"""
        self.root.deiconify()
        self.root.lift()
        self.center_window()
    
    def tray_start_backup(self):
        """从托盘启动备份"""
        if not self.backup_running:
            self.start_backup()
    
    # ------------------------------
    # 辅助功能：文件删除与程序退出
    # ------------------------------
    def delete_files(self, dest, files_to_delete):
        """删除目标文件夹中对应已删除的文件"""
        for rel_path in files_to_delete:
            file_path = os.path.join(dest, rel_path)
            try:
                if os.path.isfile(file_path):
                    os.remove(file_path)
                    self.log(f"已删除: {file_path}")
                elif os.path.isdir(file_path) and not os.listdir(file_path):
                    os.rmdir(file_path)
                    self.log(f"已删除空目录: {file_path}")
                elif os.path.isdir(file_path):
                    self.log(f"无法删除非空目录: {file_path}")
            except Exception as e:
                self.log(f"删除文件失败 {file_path}: {str(e)}")
    
    def update_startup_button(self):
        """更新开机启动按钮文本与状态"""
        if self.is_in_startup():
            self.startup_btn.config(text="移除开机启动")
            self.startup_status_var.set("已设置")
        else:
            self.startup_btn.config(text="添加开机启动")
            self.startup_status_var.set("未设置")
    
    def on_close(self):
        """窗口关闭事件（监控中则最小化到托盘）"""
        # 保存当前监控状态
        self.config.set("Settings", "monitoring", str(self.monitoring))
        with open(self.config_path, 'w', encoding="utf-8") as f:
            self.config.write(f)
        
        if self.monitoring:
            self.root.withdraw()
            if not self.tray_icon:
                self.create_tray_icon()
            self.log("程序已最小化到系统托盘（右键可操作）")
        else:
            if messagebox.askyesno("确认", "确定要退出智能自动备份工具吗？"):
                self.exit_program()
    
    def exit_program(self):
        """彻底退出程序"""
        # 保存配置与清理
        self.config.set("Settings", "monitoring", str(self.monitoring))
        with open(self.config_path, 'w', encoding="utf-8") as f:
            self.config.write(f)
        
        # 停止监控与托盘
        self.monitoring = False
        if self.tray_icon:
            self.tray_icon.stop()
        
        # 销毁窗口并退出
        self.root.destroy()
        sys.exit(0)


# ------------------------------
# 开发者工具：注册码生成（命令行）
# ------------------------------
def generate_registration_key_for_machine(machine_code):
    """为指定机器码生成注册码（仅开发者使用）"""
    print(f"智能自动备份工具 v0.01 - 注册码生成")
    print(f"目标机器码: {machine_code}")
    try:
        reg_system = HardwareRegistrationSystem()
        reg_key = reg_system.generate_registration_key(machine_code)
        print(f"生成注册码: {reg_key}")
        return reg_key
    except Exception as e:
        print(f"生成失败: {str(e)}")
        return None


# ------------------------------
# 程序入口
# ------------------------------
def main():
    # 命令行参数处理（生成注册码/最小化启动）
    start_minimized = "--minimized" in sys.argv
    # 命令行生成注册码模式（示例：python xxx.py --generate-key ABCD-EFGH-IJKL-MNOP-QRST）
    if len(sys.argv) == 3 and sys.argv[1] == "--generate-key":
        generate_registration_key_for_machine(sys.argv[2])
        return
    
    # 正常启动GUI
    root = tk.Tk()
    app = BackupTool(root, start_minimized)
    root.mainloop()


if __name__ == "__main__":
    main()