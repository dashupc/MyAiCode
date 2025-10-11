import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import subprocess
import threading
import time
import socket
import winreg
import os

class SMBFixer:
    def __init__(self, root):
        self.root = root
        self.root.title("SMB连接修复工具 v4.0 - 支持驱动器共享和权限设置")
        self.root.geometry("1000x800")
        self.root.resizable(True, True)
        
        # 标志位，用于控制线程
        self.is_running = False
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="SMB连接修复工具 (支持驱动器共享和权限设置)", font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, columnspan=4, pady=(0, 15))
        
        # 目标计算机IP输入
        ttk.Label(main_frame, text="目标计算机IP:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5))
        self.ip_var = tk.StringVar()
        ip_entry = ttk.Entry(main_frame, textvariable=self.ip_var, width=15)
        ip_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        ip_entry.bind('<Return>', lambda e: self.start_diagnosis())
        
        # 用户名密码（可选）
        ttk.Label(main_frame, text="用户名:").grid(row=2, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        self.username_var = tk.StringVar()
        username_entry = ttk.Entry(main_frame, textvariable=self.username_var, width=15)
        username_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(5, 0))
        
        ttk.Label(main_frame, text="密码:").grid(row=3, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        self.password_var = tk.StringVar()
        password_entry = ttk.Entry(main_frame, textvariable=self.password_var, show="*", width=15)
        password_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(5, 0))
        
        # 驱动器共享设置区域
        share_frame = ttk.LabelFrame(main_frame, text="驱动器共享设置", padding="5")
        share_frame.grid(row=1, column=2, rowspan=3, padx=(10, 0), sticky=(tk.N, tk.W, tk.E))
        share_frame.columnconfigure(1, weight=1)
        
        # 驱动器路径选择
        ttk.Label(share_frame, text="驱动器路径:").grid(row=0, column=0, sticky=tk.W)
        self.drive_path_var = tk.StringVar()
        drive_path_entry = ttk.Entry(share_frame, textvariable=self.drive_path_var, width=20)
        drive_path_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 5))
        browse_btn = ttk.Button(share_frame, text="浏览", command=self.browse_folder, width=6)
        browse_btn.grid(row=0, column=2)
        
        # 共享名称
        ttk.Label(share_frame, text="共享名称:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        self.share_name_var = tk.StringVar(value="SharedFolder")
        share_name_entry = ttk.Entry(share_frame, textvariable=self.share_name_var, width=20)
        share_name_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=(5, 0))
        
        # 共享描述
        ttk.Label(share_frame, text="描述:").grid(row=2, column=0, sticky=tk.W, pady=(5, 0))
        self.share_desc_var = tk.StringVar()
        share_desc_entry = ttk.Entry(share_frame, textvariable=self.share_desc_var, width=20)
        share_desc_entry.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=(5, 5), pady=(5, 0))
        
        # 权限设置
        perm_frame = ttk.Frame(share_frame)
        perm_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Label(perm_frame, text="权限:").grid(row=0, column=0, sticky=tk.W)
        self.perm_var = tk.StringVar(value="FULL")
        perm_combo = ttk.Combobox(perm_frame, textvariable=self.perm_var, 
                                 values=["FULL", "CHANGE", "READ"], width=10, state="readonly")
        perm_combo.grid(row=0, column=1, padx=(5, 0))
        
        # 共享按钮
        share_btn = ttk.Button(share_frame, text="创建共享", command=self.create_share)
        share_btn.grid(row=4, column=0, columnspan=3, pady=(10, 0))
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=3, rowspan=5, padx=(10, 0))
        
        # 诊断按钮
        self.diagnose_btn = ttk.Button(button_frame, text="开始诊断", command=self.start_diagnosis, width=15)
        self.diagnose_btn.pack(fill=tk.X, pady=(0, 3))
        
        # 快速修复按钮
        self.fix_btn = ttk.Button(button_frame, text="快速修复", command=self.start_fix, width=15)
        self.fix_btn.pack(fill=tk.X, pady=(0, 3))
        
        # 高级修复按钮
        self.advanced_fix_btn = ttk.Button(button_frame, text="高级修复", command=self.start_advanced_fix, width=15)
        self.advanced_fix_btn.pack(fill=tk.X, pady=(0, 3))
        
        # 打印机共享修复按钮
        self.printer_fix_btn = ttk.Button(button_frame, text="打印机共享修复", command=self.start_printer_fix, width=15)
        self.printer_fix_btn.pack(fill=tk.X, pady=(0, 3))
        
        # 测试连接按钮
        self.test_connect_btn = ttk.Button(button_frame, text="测试连接", command=self.test_connection, width=15)
        self.test_connect_btn.pack(fill=tk.X, pady=(0, 3))
        
        # 扫描网络按钮
        self.scan_btn = ttk.Button(button_frame, text="扫描网络", command=self.scan_network, width=15)
        self.scan_btn.pack(fill=tk.X, pady=(0, 3))
        
        # 管理共享按钮
        self.manage_shares_btn = ttk.Button(button_frame, text="管理共享", command=self.manage_shares, width=15)
        self.manage_shares_btn.pack(fill=tk.X, pady=(0, 3))
        
        # 停止按钮
        self.stop_btn = ttk.Button(button_frame, text="停止", command=self.stop_process, state=tk.DISABLED, width=15)
        self.stop_btn.pack(fill=tk.X, pady=(0, 3))
        
        # 清除日志按钮
        clear_btn = ttk.Button(button_frame, text="清除日志", command=self.clear_log, width=15)
        clear_btn.pack(fill=tk.X)
        
        # 进度条
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(10, 5))
        
        # 日志显示区域
        log_frame = ttk.LabelFrame(main_frame, text="诊断日志", padding="5")
        log_frame.grid(row=5, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=6, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # 绑定关闭事件
        root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def browse_folder(self):
        """浏览文件夹"""
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.drive_path_var.set(folder_path)
    
    def log_message(self, message):
        """添加日志消息"""
        if not self.is_running:
            return
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def update_status(self, message):
        """更新状态栏"""
        if not self.is_running:
            return
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def run_command(self, command, timeout=20):
        """执行命令并返回结果 - 修复编码问题"""
        if not self.is_running:
            return -1, "", "进程已停止"
        try:
            # 使用正确的编码处理Windows命令行输出
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            
            result = subprocess.run(
                command, 
                shell=True, 
                capture_output=True, 
                text=True, 
                timeout=timeout,
                encoding='utf-8',  # 明确指定编码
                errors='ignore',    # 忽略编码错误
                startupinfo=startupinfo
            )
            return result.returncode, result.stdout, result.stderr
        except subprocess.TimeoutExpired:
            return -1, "", "命令执行超时"
        except Exception as e:
            return -1, "", f"执行错误: {str(e)}"
    
    def create_share(self):
        """创建文件夹共享"""
        drive_path = self.drive_path_var.get().strip()
        share_name = self.share_name_var.get().strip()
        share_desc = self.share_desc_var.get().strip()
        permission = self.perm_var.get()
        
        if not drive_path or not share_name:
            messagebox.showerror("错误", "请填写驱动器路径和共享名称")
            return
        
        if not os.path.exists(drive_path):
            messagebox.showerror("错误", "指定的路径不存在")
            return
        
        # 确认操作
        if not messagebox.askyesno("确认", f"确定要共享路径 '{drive_path}' 为 '{share_name}' 吗？"):
            return
        
        self.log_message(f"正在创建共享: {share_name}")
        self.progress.start()
        
        thread = threading.Thread(target=self._create_share_worker, 
                                args=(drive_path, share_name, share_desc, permission))
        thread.daemon = True
        thread.start()
    
    def _create_share_worker(self, drive_path, share_name, share_desc, permission):
        """创建共享工作线程 - 修复版本"""
        try:
            self.update_status("正在创建共享...")
            
            # 确保路径使用双反斜杠
            drive_path = drive_path.replace('/', '\\')
            
            # 构建创建共享的命令
            if share_desc:
                cmd = f'net share "{share_name}"="{drive_path}" /remark:"{share_desc}"'
            else:
                cmd = f'net share "{share_name}"="{drive_path}"'
            
            self.log_message(f"执行命令: {cmd}")
            
            # 执行创建共享命令
            code, stdout, stderr = self.run_command(cmd, timeout=30)
            
            if code == 0:
                self.log_message(f"✓ 共享 '{share_name}' 创建成功")
                
                # 设置Everyone权限
                self.update_status("正在设置Everyone权限...")
                perm_mapping = {
                    "FULL": "FULL",
                    "CHANGE": "CHANGE", 
                    "READ": "READ"
                }
                actual_perm = perm_mapping.get(permission, "FULL")
                
                # 先删除可能存在的Everyone权限
                del_perm_cmd = f'net share "{share_name}" /delete:everyone'
                self.run_command(del_perm_cmd, timeout=15)
                
                # 添加Everyone权限
                perm_cmd = f'net share "{share_name}" /grant:everyone,{actual_perm}'
                self.log_message(f"执行权限命令: {perm_cmd}")
                
                perm_code, perm_stdout, perm_stderr = self.run_command(perm_cmd, timeout=20)
                
                if perm_code == 0:
                    self.log_message(f"✓ Everyone权限设置成功: {actual_perm}")
                else:
                    # 如果grant命令失败，尝试使用不同的方法
                    self.log_message(f"? 权限设置返回: {perm_stderr}")
                    self.log_message("尝试使用替代方法设置权限...")
                    
                    # 使用icacls设置文件夹权限
                    icacls_cmd = f'icacls "{drive_path}" /grant Everyone:({actual_perm})'
                    icacls_code, icacls_out, icacls_err = self.run_command(icacls_cmd, timeout=20)
                    if icacls_code == 0:
                        self.log_message(f"✓ 文件夹权限设置成功: {icacls_out}")
                    else:
                        self.log_message(f"✗ 文件夹权限设置失败: {icacls_err}")
                
                # 显示共享信息
                self.update_status("正在验证共享...")
                info_cmd = f'net share "{share_name}"'
                info_code, info_stdout, info_stderr = self.run_command(info_cmd, timeout=15)
                if info_code == 0:
                    self.log_message("共享详细信息:")
                    for line in info_stdout.split('\n'):
                        if line.strip():
                            self.log_message(f"  {line.strip()}")
                else:
                    self.log_message(f"获取共享信息失败: {info_stderr}")
                    
            else:
                self.log_message(f"✗ 共享创建失败!")
                self.log_message(f"  错误代码: {code}")
                self.log_message(f"  标准输出: {stdout}")
                self.log_message(f"  错误信息: {stderr}")
                
                # 尝试使用PowerShell方法
                self.log_message("尝试使用PowerShell创建共享...")
                ps_cmd = f'powershell "New-SmbShare -Name \'{share_name}\' -Path \'{drive_path}\' -FullAccess Everyone"'
                ps_code, ps_out, ps_err = self.run_command(ps_cmd, timeout=30)
                if ps_code == 0:
                    self.log_message("✓ PowerShell创建共享成功")
                else:
                    self.log_message(f"✗ PowerShell创建共享失败: {ps_err}")
                
        except Exception as e:
            self.log_message(f"创建共享失败: {str(e)}")
            import traceback
            self.log_message(f"详细错误: {traceback.format_exc()}")
        finally:
            self.progress.stop()
            self.update_status("共享创建完成")
    
    def manage_shares(self):
        """管理现有共享"""
        self.log_message("正在获取共享列表...")
        self.progress.start()
        self.manage_shares_btn.config(state=tk.DISABLED)
        self.is_running = True
        
        thread = threading.Thread(target=self._manage_shares_worker)
        thread.daemon = True
        thread.start()
    
    def _manage_shares_worker(self):
        """管理共享工作线程"""
        try:
            self.update_status("正在获取共享列表...")
            
            # 获取共享列表
            code, stdout, stderr = self.run_command('net share', timeout=20)
            
            if code == 0:
                self.log_message("当前共享列表:")
                lines = stdout.split('\n')
                shares = []
                for line in lines:
                    line = line.strip()
                    if line and not line.startswith('Share name') and not line.startswith('----') and line != 'Share':
                        # 提取共享名称（通常是第一列）
                        parts = line.split()
                        if parts:
                            share_name = parts[0]
                            if share_name not in ['IPC$', 'ADMIN$'] or share_name not in shares:  # 避免重复
                                shares.append(share_name)
                                self.log_message(f"  {share_name}")
                
                if shares:
                    # 询问是否删除某个共享
                    choice = messagebox.askyesno("管理共享", f"找到 {len(shares)} 个共享。是否要管理这些共享？")
                    if choice:
                        self.show_share_manager(shares)
                else:
                    self.log_message("  没有找到用户创建的共享")
            else:
                self.log_message(f"获取共享列表失败: {stderr}")
                
        except Exception as e:
            self.log_message(f"管理共享失败: {str(e)}")
        finally:
            self.progress.stop()
            self.manage_shares_btn.config(state=tk.NORMAL)
            self.is_running = False
    
    def show_share_manager(self, shares):
        """显示共享管理窗口"""
        manager_window = tk.Toplevel(self.root)
        manager_window.title("共享管理器")
        manager_window.geometry("400x300")
        manager_window.transient(self.root)
        manager_window.grab_set()
        
        # 创建列表框
        list_frame = ttk.Frame(manager_window, padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(list_frame, text="现有共享:").pack(anchor=tk.W)
        
        listbox = tk.Listbox(list_frame, height=10)
        listbox.pack(fill=tk.BOTH, expand=True, pady=(5, 10))
        
        for share in shares:
            if share not in ['IPC$', 'ADMIN$']:  # 过滤系统共享
                listbox.insert(tk.END, share)
        
        def delete_selected():
            selection = listbox.curselection()
            if selection:
                share_name = listbox.get(selection[0])
                if messagebox.askyesno("确认删除", f"确定要删除共享 '{share_name}' 吗？\n这将停止该文件夹的共享。"):
                    code, stdout, stderr = self.run_command(f'net share "{share_name}" /delete', timeout=20)
                    if code == 0:
                        messagebox.showinfo("成功", f"共享 '{share_name}' 已删除")
                        listbox.delete(selection[0])
                        self.log_message(f"✓ 共享 '{share_name}' 已删除")
                    else:
                        messagebox.showerror("错误", f"删除失败: {stderr}")
            else:
                messagebox.showwarning("警告", "请先选择要删除的共享")
        
        # 按钮框架
        button_frame = ttk.Frame(list_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        delete_btn = ttk.Button(button_frame, text="删除选中", command=delete_selected)
        delete_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        close_btn = ttk.Button(button_frame, text="关闭", command=manager_window.destroy)
        close_btn.pack(side=tk.RIGHT)
    
    def ping_test(self, ip):
        """测试网络连通性"""
        self.log_message(f"正在测试 {ip} 的网络连通性...")
        code, stdout, stderr = self.run_command(f"ping -n 1 -w 3000 {ip}")
        if code == 0:
            self.log_message("✓ 网络连通性正常")
            return True
        else:
            self.log_message("✗ 网络连接失败，请检查网络设置")
            return False
    
    def check_smb_port(self, ip):
        """检查SMB端口是否开放"""
        if not self.is_running:
            return False
        self.log_message(f"正在检查 {ip} 的SMB端口(445)...")
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(5)
            result = sock.connect_ex((ip, 445))
            sock.close()
            if result == 0:
                self.log_message("✓ SMB端口(445)已开放")
                return True
            else:
                self.log_message("✗ SMB端口(445)未开放或被防火墙阻止")
                return False
        except Exception as e:
            self.log_message(f"✗ 端口检查失败: {str(e)}")
            return False
    
    def check_smb_service(self):
        """检查本地SMB服务状态"""
        if not self.is_running:
            return False
        self.log_message("正在检查本地SMB服务状态...")
        services = [
            ("LanmanServer", "Server服务"),
            ("LanmanWorkstation", "Workstation服务"),
            ("Spooler", "Print Spooler服务"),  # 打印机服务
            ("Browser", "Computer Browser服务")
        ]
        
        all_running = True
        for service_name, display_name in services:
            if not self.is_running:
                return False
            code, stdout, stderr = self.run_command(f'sc query "{service_name}" | findstr STATE')
            if "RUNNING" in stdout.upper():
                self.log_message(f"✓ {display_name} 正在运行")
            else:
                self.log_message(f"✗ {display_name} 未运行")
                all_running = False
        
        return all_running
    
    def diagnose_connection(self):
        """诊断连接问题"""
        try:
            self.log_message("=" * 60)
            self.log_message("开始SMB连接诊断...")
            self.log_message("=" * 60)
            
            target_ip = self.ip_var.get().strip()
            if not target_ip:
                self.log_message("错误: 请先输入目标计算机IP地址")
                self.update_status("诊断完成 - 参数错误")
                self.progress.stop()
                self.diagnose_btn.config(state=tk.NORMAL)
                self.stop_btn.config(state=tk.DISABLED)
                self.is_running = False
                return
            
            # 1. 网络连通性测试
            self.update_status("正在测试网络连通性...")
            if not self.is_running:
                return
            if not self.ping_test(target_ip):
                self.update_status("诊断完成 - 网络连接失败")
                self.progress.stop()
                self.diagnose_btn.config(state=tk.NORMAL)
                self.stop_btn.config(state=tk.DISABLED)
                self.is_running = False
                return
            
            # 2. SMB端口检查
            self.update_status("正在检查SMB端口...")
            if not self.is_running:
                return
            self.check_smb_port(target_ip)
            
            # 3. 本地SMB服务检查
            self.update_status("正在检查本地SMB服务...")
            if not self.is_running:
                return
            self.check_smb_service()
            
            # 4. SMB版本兼容性检查
            self.log_message("检查SMB协议版本兼容性...")
            if not self.is_running:
                return
            self.check_smb_version()
            
            # 5. 防火墙检查
            self.log_message("检查Windows防火墙设置...")
            if not self.is_running:
                return
            self.check_firewall()
            
            # 6. 打印机共享检查
            self.log_message("检查打印机共享设置...")
            if not self.is_running:
                return
            self.check_printer_sharing()
            
            # 7. 本地共享检查
            self.log_message("检查本地共享设置...")
            if not self.is_running:
                return
            self.check_local_shares()
            
            self.log_message("\n诊断完成！根据以上信息进行相应处理。")
            self.update_status("诊断完成")
        except Exception as e:
            self.log_message(f"诊断过程中出现错误: {str(e)}")
        finally:
            self.progress.stop()
            self.diagnose_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            self.is_running = False
    
    def check_smb_version(self):
        """检查SMB版本兼容性"""
        if not self.is_running:
            return
        # 检查本地系统类型
        code, stdout, stderr = self.run_command('ver')
        if '6.1' in stdout:  # Windows 7
            self.log_message("本地系统: Windows 7 (SMB 2.0)")
        elif '10.0' in stdout:  # Windows 10/11
            self.log_message("本地系统: Windows 10/11 (SMB 3.0)")
        
        # 检查SMB配置
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                               r"SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters")
            try:
                smb2_enabled, _ = winreg.QueryValueEx(key, "SMB2")
                if smb2_enabled == 1:
                    self.log_message("✓ SMB 2.0 已启用")
                else:
                    self.log_message("✗ SMB 2.0 已禁用")
            except:
                self.log_message("✓ SMB 2.0 配置正常")
            winreg.CloseKey(key)
        except:
            self.log_message("无法读取SMB版本配置")
    
    def check_firewall(self):
        """检查防火墙设置"""
        if not self.is_running:
            return
        firewall_rules = [
            "File and Printer Sharing (SMB-In)",
            "Network Discovery (SMB-In)",
            "File and Printer Sharing (Echo Request - ICMPv4-In)"
        ]
        
        for rule in firewall_rules:
            if not self.is_running:
                return
            code, stdout, stderr = self.run_command(
                f'netsh advfirewall firewall show rule name="{rule}"'
            )
            if "enabled:yes" in stdout.lower():
                self.log_message(f"✓ 防火墙规则 '{rule}' 已启用")
            else:
                self.log_message(f"✗ 防火墙规则 '{rule}' 未启用或不存在")
    
    def check_printer_sharing(self):
        """检查打印机共享设置"""
        if not self.is_running:
            return
        self.log_message("检查本地打印机共享服务...")
        
        # 检查打印池服务
        code, stdout, stderr = self.run_command('sc query Spooler | findstr STATE')
        if "RUNNING" in stdout.upper():
            self.log_message("✓ Print Spooler服务正在运行")
        else:
            self.log_message("✗ Print Spooler服务未运行")
        
        # 检查注册表中的打印机共享设置
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                               r"SYSTEM\CurrentControlSet\Control\Print")
            share_printers, _ = winreg.QueryValueEx(key, "SharePrinters")
            winreg.CloseKey(key)
            if share_printers == 1:
                self.log_message("✓ 打印机共享已启用")
            else:
                self.log_message("✗ 打印机共享未启用")
        except:
            self.log_message("? 无法读取打印机共享配置")
    
    def check_local_shares(self):
        """检查本地共享设置"""
        if not self.is_running:
            return
        self.log_message("检查本地文件共享...")
        
        code, stdout, stderr = self.run_command('net share', timeout=20)
        if code == 0:
            lines = stdout.split('\n')
            share_count = 0
            for line in lines:
                line = line.strip()
                if line and not line.startswith('Share name') and not line.startswith('----') and line != 'Share':
                    parts = line.split()
                    if parts and parts[0] not in ['IPC$', 'ADMIN$']:  # 过滤系统共享
                        share_count += 1
                        self.log_message(f"✓ 共享: {parts[0]}")
            
            if share_count > 0:
                self.log_message(f"✓ 共找到 {share_count} 个用户共享")
            else:
                self.log_message("✗ 未找到用户创建的共享")
        else:
            self.log_message(f"✗ 获取共享列表失败: {stderr}")
    
    def fix_common_issues(self):
        """快速修复常见问题"""
        try:
            self.log_message("=" * 60)
            self.log_message("开始快速修复...")
            self.log_message("=" * 60)
            
            fixes_applied = []
            
            # 1. 启动必要的服务
            self.update_status("正在启动必要服务...")
            services_to_start = [
                ("LanmanServer", "Server服务"),
                ("LanmanWorkstation", "Workstation服务"),
                ("Browser", "Computer Browser服务"),
                ("Spooler", "Print Spooler服务")  # 添加打印机服务
            ]
            
            for service_name, display_name in services_to_start:
                if not self.is_running:
                    return
                code, stdout, stderr = self.run_command(f'net start "{service_name}"')
                if code == 0 or "already been started" in stderr.lower():
                    self.log_message(f"✓ {display_name} 启动成功")
                    fixes_applied.append(f"启动{display_name}")
                else:
                    self.log_message(f"⚠ {display_name} 启动失败: {stderr.strip()}")
            
            # 2. 设置网络发现和文件共享
            self.update_status("正在配置网络发现和文件共享...")
            commands = [
                ('netsh advfirewall firewall set rule group="network discovery" new enable=Yes', "网络发现防火墙规则"),
                ('netsh advfirewall firewall set rule group="file and printer sharing" new enable=Yes', "文件和打印机共享防火墙规则")
            ]
            
            for cmd, desc in commands:
                if not self.is_running:
                    return
                code, stdout, stderr = self.run_command(cmd)
                if code == 0:
                    self.log_message(f"✓ {desc} 配置成功")
                    fixes_applied.append(desc)
                else:
                    self.log_message(f"⚠ {desc} 配置失败")
            
            # 3. 启用SMB协议
            self.update_status("正在启用SMB协议...")
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                   r"SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters",
                                   0, winreg.KEY_SET_VALUE)
                winreg.SetValueEx(key, "SMB2", 0, winreg.REG_DWORD, 1)
                winreg.CloseKey(key)
                self.log_message("✓ SMB 2.0 协议已启用")
                fixes_applied.append("启用SMB 2.0")
            except Exception as e:
                self.log_message(f"⚠ 启用SMB 2.0 失败: {str(e)}")
            
            # 4. 刷新网络配置
            self.update_status("正在刷新网络配置...")
            refresh_commands = [
                ("nbtstat -R", "重置NetBIOS名称缓存"),
                ("ipconfig /flushdns", "刷新DNS缓存"),
                ("net use * /delete /y", "清除网络连接缓存")
            ]
            
            for cmd, desc in refresh_commands:
                if not self.is_running:
                    return
                code, stdout, stderr = self.run_command(cmd)
                if code == 0:
                    self.log_message(f"✓ {desc}")
                else:
                    self.log_message(f"⚠ {desc} 执行失败")
            
            if fixes_applied:
                self.log_message(f"\n已应用以下修复: {', '.join(fixes_applied)}")
                self.log_message("建议重启计算机以使所有更改生效")
            else:
                self.log_message("\n未发现需要修复的问题")
            
            self.update_status("快速修复完成")
        except Exception as e:
            self.log_message(f"修复过程中出现错误: {str(e)}")
        finally:
            self.progress.stop()
            self.fix_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            self.is_running = False
    
    def fix_printer_sharing(self):
        """专门修复打印机共享问题"""
        try:
            self.log_message("=" * 60)
            self.log_message("开始打印机共享修复...")
            self.log_message("=" * 60)
            
            fixes_applied = []
            
            # 1. 启动打印池服务
            self.update_status("正在启动打印池服务...")
            code, stdout, stderr = self.run_command('net start Spooler')
            if code == 0 or "already been started" in stderr.lower():
                self.log_message("✓ Print Spooler服务启动成功")
                fixes_applied.append("启动Print Spooler服务")
            else:
                self.log_message(f"✗ Print Spooler服务启动失败: {stderr}")
            
            # 2. 配置打印机共享防火墙规则
            self.update_status("正在配置打印机共享防火墙...")
            firewall_commands = [
                ('netsh advfirewall firewall set rule group="file and printer sharing" new enable=Yes', "文件和打印机共享"),
                ('netsh advfirewall firewall add rule name="Printer Sharing (LPR)" dir=in action=allow protocol=TCP localport=515', "LPR端口(515)"),
                ('netsh advfirewall firewall add rule name="Printer Sharing (LPD)" dir=in action=allow protocol=TCP localport=515 program="%SystemRoot%\\system32\\spoolsv.exe"', "LPD服务")
            ]
            
            for cmd, desc in firewall_commands:
                if not self.is_running:
                    return
                code, stdout, stderr = self.run_command(cmd)
                if code == 0:
                    self.log_message(f"✓ {desc} 防火墙规则配置成功")
                    fixes_applied.append(f"配置{desc}")
                else:
                    self.log_message(f"⚠ {desc} 防火墙规则配置失败")
            
            # 3. 启用注册表中的打印机共享
            self.update_status("正在配置注册表打印机共享设置...")
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                   r"SYSTEM\CurrentControlSet\Control\Print",
                                   0, winreg.KEY_SET_VALUE)
                winreg.SetValueEx(key, "SharePrinters", 0, winreg.REG_DWORD, 1)
                winreg.CloseKey(key)
                self.log_message("✓ 注册表打印机共享设置已启用")
                fixes_applied.append("启用注册表打印机共享")
            except Exception as e:
                self.log_message(f"⚠ 注册表打印机共享设置失败: {str(e)}")
            
            # 4. 重启打印服务
            self.update_status("正在重启打印服务...")
            restart_commands = [
                ("net stop spooler", "停止打印池服务"),
                ("net start spooler", "启动打印池服务")
            ]
            
            for cmd, desc in restart_commands:
                if not self.is_running:
                    return
                code, stdout, stderr = self.run_command(cmd, timeout=20)
                if code == 0:
                    self.log_message(f"✓ {desc}")
                else:
                    self.log_message(f"⚠ {desc} 失败")
            
            # 5. 检查SMB共享设置
            self.update_status("正在检查SMB共享设置...")
            code, stdout, stderr = self.run_command('net share')
            if code == 0:
                self.log_message("✓ SMB共享服务正常")
                if "IPC$" not in stdout:
                    # 重新创建IPC$共享
                    cmd_code, cmd_out, cmd_err = self.run_command('net share IPC$="" /grant:everyone,FULL')
                    if cmd_code == 0:
                        self.log_message("✓ IPC$共享已重新创建")
                        fixes_applied.append("重新创建IPC$共享")
                    else:
                        self.log_message(f"✗ IPC$共享创建失败: {cmd_err}")
            else:
                self.log_message(f"✗ SMB共享检查失败: {stderr}")
            
            if fixes_applied:
                self.log_message(f"\n已应用以下打印机共享修复: {', '.join(fixes_applied)}")
                self.log_message("打印机共享修复完成！")
            else:
                self.log_message("\n未发现需要修复的打印机共享问题")
            
            self.update_status("打印机共享修复完成")
        except Exception as e:
            self.log_message(f"打印机共享修复过程中出现错误: {str(e)}")
        finally:
            self.progress.stop()
            self.printer_fix_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            self.is_running = False
    
    def advanced_fix(self):
        """高级修复"""
        try:
            self.log_message("=" * 60)
            self.log_message("开始高级修复...")
            self.log_message("=" * 60)
            
            # 1. 重置TCP/IP堆栈
            self.update_status("正在重置TCP/IP堆栈...")
            self.log_message("正在重置TCP/IP堆栈...")
            if not self.is_running:
                return
            code, stdout, stderr = self.run_command("netsh int ip reset", timeout=30)
            if code == 0:
                self.log_message("✓ TCP/IP堆栈重置成功")
            else:
                self.log_message(f"✗ TCP/IP堆栈重置失败: {stderr}")
            
            # 2. 重置Winsock目录
            self.update_status("正在重置Winsock...")
            self.log_message("正在重置Winsock...")
            if not self.is_running:
                return
            code, stdout, stderr = self.run_command("netsh winsock reset", timeout=30)
            if code == 0:
                self.log_message("✓ Winsock重置成功")
            else:
                self.log_message(f"✗ Winsock重置失败: {stderr}")
            
            # 3. 配置注册表项
            self.update_status("正在配置注册表...")
            if not self.is_running:
                return
            self.configure_registry()
            
            # 4. 重新启动相关服务
            self.update_status("正在重启网络服务...")
            restart_services = [
                "LanmanServer",
                "LanmanWorkstation", 
                "Dnscache",
                "NlaSvc",
                "Spooler"
            ]
            
            for service in restart_services:
                if not self.is_running:
                    return
                self.run_command(f"net stop {service}", timeout=15)
                self.run_command(f"net start {service}", timeout=15)
            
            self.log_message("✓ 网络服务已重启")
            self.log_message("\n高级修复完成！建议立即重启计算机。")
            self.update_status("高级修复完成")
        except Exception as e:
            self.log_message(f"高级修复过程中出现错误: {str(e)}")
        finally:
            self.progress.stop()
            self.advanced_fix_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            self.is_running = False
    
    def configure_registry(self):
        """配置注册表以改善SMB和打印机连接"""
        if not self.is_running:
            return
        reg_configs = [
            # 启用不安全的来宾登录
            {
                "path": r"SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters",
                "key": "AllowInsecureGuestAuth",
                "value": 1,
                "type": winreg.REG_DWORD
            },
            # 禁用SMB签名要求（仅用于内网环境）
            {
                "path": r"SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters",
                "key": "RequireSecuritySignature",
                "value": 0,
                "type": winreg.REG_DWORD
            },
            # 启用LM兼容性级别
            {
                "path": r"SYSTEM\CurrentControlSet\Control\Lsa",
                "key": "LmCompatibilityLevel",
                "value": 1,
                "type": winreg.REG_DWORD
            },
            # 启用打印机共享
            {
                "path": r"SYSTEM\CurrentControlSet\Control\Print",
                "key": "SharePrinters",
                "value": 1,
                "type": winreg.REG_DWORD
            }
        ]
        
        success_count = 0
        for config in reg_configs:
            if not self.is_running:
                return
            try:
                key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, config["path"])
                winreg.SetValueEx(key, config["key"], 0, config["type"], config["value"])
                winreg.CloseKey(key)
                self.log_message(f"✓ 配置注册表项: {config['key']}")
                success_count += 1
            except Exception as e:
                self.log_message(f"✗ 配置注册表项失败: {config['key']} - {str(e)}")
        
        self.log_message(f"注册表配置完成 ({success_count}/{len(reg_configs)})")
    
    def scan_network(self):
        """扫描网络中的共享资源"""
        self.log_message("开始扫描网络共享资源...")
        self.progress.start()
        self.scan_btn.config(state=tk.DISABLED)
        self.is_running = True
        
        thread = threading.Thread(target=self._scan_network_worker)
        thread.daemon = True
        thread.start()
    
    def _scan_network_worker(self):
        """网络扫描工作线程"""
        try:
            self.update_status("正在扫描网络...")
            
            # 扫描网络邻居
            self.log_message("扫描网络邻居...")
            code, stdout, stderr = self.run_command('net view', timeout=30)
            if code == 0:
                self.log_message("网络邻居扫描结果:")
                for line in stdout.split('\n'):
                    if '\\\\' in line:
                        self.log_message(f"  {line.strip()}")
            else:
                self.log_message(f"网络邻居扫描失败: {stderr}")
            
            # 扫描本地共享
            self.log_message("扫描本地共享...")
            code, stdout, stderr = self.run_command('net share', timeout=20)
            if code == 0:
                self.log_message("本地共享列表:")
                for line in stdout.split('\n'):
                    if line.strip() and not line.startswith('Share name') and not line.startswith('----'):
                        self.log_message(f"  {line.strip()}")
            else:
                self.log_message(f"本地共享扫描失败: {stderr}")
            
            self.update_status("网络扫描完成")
        except Exception as e:
            self.log_message(f"网络扫描失败: {str(e)}")
        finally:
            self.progress.stop()
            self.scan_btn.config(state=tk.NORMAL)
            self.is_running = False
    
    def test_connection(self):
        """测试SMB连接"""
        target_ip = self.ip_var.get().strip()
        if not target_ip:
            messagebox.showerror("错误", "请先输入目标计算机IP地址")
            return
        
        # 尝试连接共享
        self.log_message(f"正在测试连接到 \\\\{target_ip}\\")
        self.progress.start()
        self.test_connect_btn.config(state=tk.DISABLED)
        self.is_running = True
        
        thread = threading.Thread(target=self._test_connection_worker, args=(target_ip,))
        thread.daemon = True
        thread.start()
    
    def _test_connection_worker(self, target_ip):
        """测试连接工作线程"""
        try:
            username = self.username_var.get()
            password = self.password_var.get()
            
            if username and password:
                cmd = f'net use \\\\{target_ip}\\ipc$ "{password}" /user:"{username}"'
            else:
                cmd = f'net use \\\\{target_ip}\\ipc$'
                
            code, stdout, stderr = self.run_command(cmd, timeout=20)
            if code == 0:
                self.log_message(f"✓ 成功连接到 \\\\{target_ip}\\")
                # 测试打印机连接
                self.log_message("测试打印机共享连接...")
                print_cmd = f'net use lpt2: \\\\{target_ip}\\printer_name'
                print_code, print_stdout, print_stderr = self.run_command(print_cmd, timeout=15)
                if print_code == 0:
                    self.log_message("✓ 打印机共享连接成功")
                else:
                    self.log_message("? 打印机共享连接测试: 无打印机或连接失败")
                
                # 清理连接
                self.run_command(f'net use \\\\{target_ip}\\ipc$ /delete', timeout=10)
            else:
                self.log_message(f"✗ 连接失败: {stderr}")
        except Exception as e:
            self.log_message(f"连接测试失败: {str(e)}")
        finally:
            self.progress.stop()
            self.test_connect_btn.config(state=tk.NORMAL)
            self.is_running = False
    
    def start_diagnosis(self):
        """开始诊断（在线程中运行）"""
        if self.is_running:
            messagebox.showwarning("警告", "已有进程正在运行，请先停止当前进程")
            return
        
        self.log_message("准备开始诊断...")
        self.progress.start()
        self.diagnose_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.is_running = True
        
        thread = threading.Thread(target=self.diagnose_connection)
        thread.daemon = True
        thread.start()
    
    def start_fix(self):
        """开始快速修复（在线程中运行）"""
        if self.is_running:
            messagebox.showwarning("警告", "已有进程正在运行，请先停止当前进程")
            return
            
        if messagebox.askyesno("确认", "确定要执行快速修复吗？这可能会修改系统设置。"):
            self.log_message("准备开始快速修复...")
            self.progress.start()
            self.fix_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.NORMAL)
            self.is_running = True
            
            thread = threading.Thread(target=self.fix_common_issues)
            thread.daemon = True
            thread.start()
    
    def start_advanced_fix(self):
        """开始高级修复（在线程中运行）"""
        if self.is_running:
            messagebox.showwarning("警告", "已有进程正在运行，请先停止当前进程")
            return
            
        if messagebox.askyesno("重要提示", 
                              "高级修复会重置网络配置并修改注册表，建议在管理员权限下运行。\n\n确定继续吗？"):
            self.log_message("准备开始高级修复...")
            self.progress.start()
            self.advanced_fix_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.NORMAL)
            self.is_running = True
            
            thread = threading.Thread(target=self.advanced_fix)
            thread.daemon = True
            thread.start()
    
    def start_printer_fix(self):
        """开始打印机共享修复（在线程中运行）"""
        if self.is_running:
            messagebox.showwarning("警告", "已有进程正在运行，请先停止当前进程")
            return
            
        if messagebox.askyesno("确认", "确定要执行打印机共享修复吗？这将配置打印机相关服务和防火墙规则。"):
            self.log_message("准备开始打印机共享修复...")
            self.progress.start()
            self.printer_fix_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.NORMAL)
            self.is_running = True
            
            thread = threading.Thread(target=self.fix_printer_sharing)
            thread.daemon = True
            thread.start()
    
    def stop_process(self):
        """停止当前进程"""
        self.is_running = False
        self.log_message("正在停止进程...")
        self.update_status("正在停止...")
        
        # 启用按钮
        self.diagnose_btn.config(state=tk.NORMAL)
        self.fix_btn.config(state=tk.NORMAL)
        self.advanced_fix_btn.config(state=tk.NORMAL)
        self.printer_fix_btn.config(state=tk.NORMAL)
        self.test_connect_btn.config(state=tk.NORMAL)
        self.scan_btn.config(state=tk.NORMAL)
        self.manage_shares_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        
        self.progress.stop()
        self.update_status("进程已停止")
    
    def clear_log(self):
        """清除日志"""
        self.log_text.delete(1.0, tk.END)
        self.update_status("日志已清除")
    
    def on_closing(self):
        """程序关闭时的处理"""
        if self.is_running:
            if messagebox.askyesno("确认退出", "有进程正在运行，确定要退出吗？"):
                self.is_running = False
                self.root.destroy()
        else:
            if messagebox.askokcancel("退出", "确定要退出SMB修复工具吗？"):
                self.root.destroy()

def main():
    # 检查是否以管理员权限运行
    try:
        import ctypes
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
    except:
        is_admin = False
    
    root = tk.Tk()
    app = SMBFixer(root)
    
    # 如果不是管理员权限，给出提示
    if not is_admin:
        messagebox.showwarning("权限提示", "建议以管理员权限运行此工具以获得最佳效果")
    
    root.mainloop()

if __name__ == "__main__":
    main()
