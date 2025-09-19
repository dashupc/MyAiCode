import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import threading
import time
from datetime import datetime
import sys
import hashlib
import winreg  # 用于操作Windows注册表
import configparser  # 用于保存配置

class BackupTool:
    def __init__(self, root):
        # 设置中文字体支持
        self.root = root
        self.root.title("智能自动备份工具")
        self.root.geometry("700x500")
        self.root.resizable(True, True)
        
        # 配置文件路径
        self.config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "backup_config.ini")
        if getattr(sys, 'frozen', False):
            self.config_path = os.path.join(os.path.dirname(sys.executable), "backup_config.ini")
        
        # 加载配置
        self.load_config()
        
        # 确保中文显示正常
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("SimHei", 10))
        self.style.configure("TButton", font=("SimHei", 10))
        self.style.configure("TEntry", font=("SimHei", 10))
        
        # 源文件夹和目标文件夹路径
        self.source_path = tk.StringVar(value=self.config.get("Paths", "source", fallback=""))
        self.dest_path = tk.StringVar(value=self.config.get("Paths", "dest", fallback=""))
        
        # 监控间隔（秒）
        self.monitor_interval = tk.IntVar(value=self.config.getint("Settings", "interval", fallback=60))
        
        # 备份状态
        self.backup_running = False
        self.monitoring = False
        self.file_hashes = {}  # 存储文件哈希值，用于检测变化
        
        # 创建UI
        self.create_widgets()
        
        # 检查是否在启动项中
        self.check_startup_status()
    
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
        
        self.startup_btn = ttk.Button(button_frame, text="添加到开机启动", command=self.toggle_startup)
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
        
        # 设置网格权重，使控件可以随窗口大小调整
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(7, weight=1)
    
    def select_source(self):
        path = filedialog.askdirectory(title="选择源文件夹")
        if path:
            self.source_path.set(path)
            # 计算初始哈希值
            self.calculate_initial_hashes()
    
    def select_dest(self):
        path = filedialog.askdirectory(title="选择目标文件夹")
        if path:
            self.dest_path.set(path)
    
    def log(self, message):
        """添加日志信息"""
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
            messagebox.showerror("错误", "请选择有效的源文件夹")
            self.backup_running = False
            self.backup_btn.config(text="立即备份", state=tk.NORMAL)
            self.update_status("就绪")
            return
        
        if not dest or not os.path.isdir(dest):
            messagebox.showerror("错误", "请选择有效的目标文件夹")
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
            self.monitor_btn.config(text="停止监控")
            self.calculate_initial_hashes()
            self.log(f"开始监控，间隔 {interval} 秒")
            
            # 启动监控线程
            thread = threading.Thread(target=self.monitoring_thread)
            thread.daemon = True
            thread.start()
        else:
            # 停止监控
            self.monitoring = False
            self.monitor_btn.config(text="开始监控")
            self.log("已停止监控")
            self.update_status("就绪")
    
    def toggle_startup(self):
        """添加或移除开机启动"""
        if self.is_in_startup():
            # 移除开机启动
            if self.remove_from_startup():
                self.log("已从开机启动中移除")
                self.startup_status_var.set("未设置")
            else:
                self.log("从开机启动中移除失败")
        else:
            # 添加到开机启动
            if self.add_to_startup():
                self.log("已添加到开机启动")
                self.startup_status_var.set("已设置")
            else:
                self.log("添加到开机启动失败")
    
    def get_executable_path(self):
        """获取当前程序的路径"""
        if getattr(sys, 'frozen', False):
            # 打包后的EXE
            return sys.executable
        else:
            # 开发环境
            return os.path.abspath(__file__)
    
    def add_to_startup(self):
        """添加到Windows开机启动"""
        try:
            # 获取程序路径和名称
            exe_path = self.get_executable_path()
            exe_name = os.path.basename(exe_path)
            
            # 打开注册表
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_SET_VALUE
            )
            
            # 添加注册表项
            winreg.SetValueEx(key, exe_name, 0, winreg.REG_SZ, exe_path)
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
    
    def check_startup_status(self):
        """检查并更新开机启动状态显示"""
        if self.is_in_startup():
            self.startup_status_var.set("已设置")
        else:
            self.startup_status_var.set("未设置")
    
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
        """保存配置文件"""
        try:
            self.config.set("Paths", "source", self.source_path.get())
            self.config.set("Paths", "dest", self.dest_path.get())
            self.config.set("Settings", "interval", str(self.monitor_interval.get()))
            
            with open(self.config_path, 'w', encoding="utf-8") as f:
                self.config.write(f)
            
            self.log(f"配置已保存到 {self.config_path}")
            messagebox.showinfo("成功", "配置已保存")
        except Exception as e:
            self.log(f"保存配置失败: {str(e)}")
            messagebox.showerror("错误", f"保存配置失败: {str(e)}")

def main():
    # 确保中文显示正常
    root = tk.Tk()
    # 设置窗口图标（打包时需要确保ico文件存在）
    try:
        if getattr(sys, 'frozen', False):
            # 运行在打包后的环境中
            base_path = sys._MEIPASS
        else:
            # 运行在普通Python环境中
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        icon_path = os.path.join(base_path, "backup.ico")
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
    except:
        pass  # 图标设置失败不影响程序运行
    
    app = BackupTool(root)
    root.mainloop()

if __name__ == "__main__":
    main()
    