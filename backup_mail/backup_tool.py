import os
import sys
import time
import zipfile
import schedule
import threading
import configparser
import smtplib
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formatdate
from datetime import datetime
import winreg as reg
import ctypes

# 隐藏控制台窗口
def hide_console():
    if sys.platform.startswith('win32'):
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

# 配置参数
SMTP_RETRY = 3  # 重试次数
RETRY_DELAY = 10  # 重试间隔(秒)
SEND_DELAY = 15  # 分卷邮件发送间隔(秒)，避免频率限制

class BackupTool:
    def __init__(self, root):
        # 初始化主窗口
        self.root = root
        self.root.title("文件备份工具")
        self.root.geometry("650x580")
        self.root.resizable(False, False)
        
        # 设置图标
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass
        
        # 居中窗口
        self.center_window()
        
        # 配置文件路径
        self.config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.ini")
        
        # 初始化变量
        self.source_folder = StringVar()
        self.dest_email = StringVar()
        self.sender_email = StringVar()
        self.sender_password = StringVar()
        self.backup_interval = StringVar(value="1")
        self.part_size = StringVar(value="19")
        self.is_running = False
        self.backup_thread = None
        
        # 创建UI
        self.create_widgets()
        
        # 加载配置
        self.load_config()
        
        # 检查启动项
        self.check_startup_status()
        
        # 启动定时任务线程
        self.start_scheduler_thread()
    
    def center_window(self):
        """窗口居中"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def create_widgets(self):
        """创建界面组件"""
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=BOTH, expand=True)
        
        # 源文件夹选择
        ttk.Label(main_frame, text="源文件夹:").grid(row=0, column=0, sticky=W, pady=5)
        ttk.Entry(main_frame, textvariable=self.source_folder, width=55).grid(row=0, column=1, pady=5)
        ttk.Button(main_frame, text="浏览...", command=self.browse_source).grid(row=0, column=2, padx=5, pady=5)
        
        # 目标邮箱
        ttk.Label(main_frame, text="目标邮箱:").grid(row=1, column=0, sticky=W, pady=5)
        ttk.Entry(main_frame, textvariable=self.dest_email, width=55).grid(row=1, column=1, columnspan=2, sticky=W, pady=5)
        
        # 发件人邮箱
        ttk.Label(main_frame, text="发件人QQ邮箱:").grid(row=2, column=0, sticky=W, pady=5)
        ttk.Entry(main_frame, textvariable=self.sender_email, width=55).grid(row=2, column=1, columnspan=2, sticky=W, pady=5)
        
        # 邮箱授权码
        ttk.Label(main_frame, text="QQ邮箱授权码:").grid(row=3, column=0, sticky=W, pady=5)
        ttk.Entry(main_frame, textvariable=self.sender_password, show="*", width=55).grid(row=3, column=1, columnspan=2, sticky=W, pady=5)
        ttk.Label(main_frame, text="(需开启SMTP服务，在邮箱设置中获取)", foreground="gray").grid(row=4, column=1, sticky=W)
        
        # 分卷大小
        ttk.Label(main_frame, text="分卷大小(MB):").grid(row=5, column=0, sticky=W, pady=5)
        ttk.Entry(main_frame, textvariable=self.part_size, width=10).grid(row=5, column=1, sticky=W, pady=5)
        ttk.Label(main_frame, text="(建议19，不超过20)", foreground="gray").grid(row=5, column=2, sticky=W)
        
        # 备份间隔
        ttk.Label(main_frame, text="备份间隔(小时):").grid(row=6, column=0, sticky=W, pady=5)
        ttk.Entry(main_frame, textvariable=self.backup_interval, width=10).grid(row=6, column=1, sticky=W, pady=5)
        
        # 随系统启动
        self.startup_var = BooleanVar()
        ttk.Checkbutton(main_frame, text="随Windows启动", variable=self.startup_var, command=self.toggle_startup).grid(row=7, column=0, columnspan=3, sticky=W, pady=10)
        
        # 状态显示
        ttk.Label(main_frame, text="状态:").grid(row=8, column=0, sticky=W, pady=5)
        self.status_var = StringVar(value="未运行")
        ttk.Label(main_frame, textvariable=self.status_var, foreground="red").grid(row=8, column=1, sticky=W, pady=5)
        
        # 日志区域
        ttk.Label(main_frame, text="日志:").grid(row=9, column=0, sticky=NW, pady=5)
        self.log_text = Text(main_frame, width=65, height=8, state=DISABLED)
        self.log_text.grid(row=9, column=1, columnspan=2, pady=5)
        scrollbar = ttk.Scrollbar(main_frame, command=self.log_text.yview)
        scrollbar.grid(row=9, column=3, sticky=NS)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=10, column=0, columnspan=3, pady=20)
        
        self.start_button = ttk.Button(button_frame, text="开始备份服务", command=self.start_service)
        self.start_button.pack(side=LEFT, padx=10)
        
        ttk.Button(button_frame, text="立即备份", command=self.backup_now).pack(side=LEFT, padx=10)
        
        ttk.Button(button_frame, text="测试邮箱连接", command=self.test_email_connection).pack(side=LEFT, padx=10)
    
    def browse_source(self):
        """浏览选择源文件夹"""
        folder = filedialog.askdirectory()
        if folder:
            self.source_folder.set(folder)
    
    def log(self, message):
        """添加日志信息"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        
        self.log_text.config(state=NORMAL)
        self.log_text.insert(END, log_message)
        self.log_text.see(END)
        self.log_text.config(state=DISABLED)
    
    def load_config(self):
        """加载配置文件"""
        if os.path.exists(self.config_path):
            config = configparser.ConfigParser()
            config.read(self.config_path)
            
            if "Settings" in config:
                settings = config["Settings"]
                self.source_folder.set(settings.get("source_folder", ""))
                self.dest_email.set(settings.get("dest_email", ""))
                self.sender_email.set(settings.get("sender_email", ""))
                self.sender_password.set(settings.get("sender_password", ""))
                self.backup_interval.set(settings.get("backup_interval", "1"))
                self.part_size.set(settings.get("part_size", "19"))
                
                self.log("配置文件加载成功")
    
    def save_config(self):
        """保存配置文件"""
        config = configparser.ConfigParser()
        config["Settings"] = {
            "source_folder": self.source_folder.get(),
            "dest_email": self.dest_email.get(),
            "sender_email": self.sender_email.get(),
            "sender_password": self.sender_password.get(),
            "backup_interval": self.backup_interval.get(),
            "part_size": self.part_size.get()
        }
        
        with open(self.config_path, "w") as configfile:
            config.write(configfile)
        
        self.log("配置已保存")
        messagebox.showinfo("成功", "配置已保存")
    
    def check_startup_status(self):
        """检查是否随系统启动"""
        try:
            key = reg.OpenKey(reg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, reg.KEY_READ)
            value, _ = reg.QueryValueEx(key, "FileBackupTool")
            reg.CloseKey(key)
            
            if value == sys.executable:
                self.startup_var.set(True)
                return True
        except:
            pass
        
        self.startup_var.set(False)
        return False
    
    def toggle_startup(self):
        """切换是否随系统启动"""
        try:
            key = reg.OpenKey(reg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, reg.KEY_SET_VALUE)
            
            if self.startup_var.get():
                reg.SetValueEx(key, "FileBackupTool", 0, reg.REG_SZ, sys.executable)
                self.log("已设置随Windows启动")
            else:
                try:
                    reg.DeleteValue(key, "FileBackupTool")
                    self.log("已取消随Windows启动")
                except:
                    pass
            
            reg.CloseKey(key)
        except Exception as e:
            self.log(f"设置启动项失败: {str(e)}")
            messagebox.showerror("错误", f"设置启动项失败: {str(e)}")
    
    def create_split_zip(self):
        """创建分卷压缩文件"""
        if not self.source_folder.get():
            self.log("请设置源文件夹")
            return None
        
        if not os.path.exists(self.source_folder.get()):
            self.log(f"源文件夹不存在: {self.source_folder.get()}")
            return None
        
        try:
            # 获取分卷大小(MB转字节)
            try:
                part_size_mb = float(self.part_size.get())
                if part_size_mb <= 0 or part_size_mb > 20:
                    self.log("分卷大小必须大于0且不超过20MB")
                    return None
                part_size = int(part_size_mb * 1024 * 1024)
            except ValueError:
                self.log("请输入有效的分卷大小")
                return None
            
            # 创建临时压缩文件
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_zip_name = f"backup_{timestamp}"
            base_zip_path = os.path.join(os.environ["TEMP"], base_zip_name)
            
            # 创建分卷压缩
            with zipfile.ZipFile(f"{base_zip_path}.zip", 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as zipf:
                zipf.pkzip = True  # 启用分卷模式
                
                for root, dirs, files in os.walk(self.source_folder.get()):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, self.source_folder.get())
                        zipf.write(file_path, arcname)
            
            # 收集所有分卷文件
            part_files = []
            part_number = 1
            while True:
                part_ext = f".z{part_number:02d}" if part_number > 1 else ".zip"
                part_file = f"{base_zip_path}{part_ext}"
                if os.path.exists(part_file):
                    part_files.append(part_file)
                    part_number += 1
                else:
                    break
            
            self.log(f"分卷压缩完成，共 {len(part_files)} 个文件")
            return part_files
        except Exception as e:
            self.log(f"分卷压缩失败: {str(e)}")
            return None
    
    def send_single_email(self, to_email, subject, body, attachment_path=None):
        """发送单封邮件（带重试机制）"""
        for attempt in range(SMTP_RETRY):
            try:
                # 创建邮件
                msg = MIMEMultipart()
                msg['From'] = self.sender_email.get()
                msg['To'] = to_email
                msg['Subject'] = subject
                msg['Date'] = formatdate(localtime=True)
                
                # 添加正文
                msg.attach(MIMEText(body, 'plain', 'utf-8'))
                
                # 添加附件
                if attachment_path and os.path.exists(attachment_path):
                    filename = os.path.basename(attachment_path)
                    with open(attachment_path, "rb") as attachment:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(attachment.read())
                    
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f"attachment; filename*=UTF-8''{filename}",)
                    msg.attach(part)
                
                # 发送邮件（带SSL加密）
                with smtplib.SMTP_SSL("smtp.qq.com", 465, timeout=30) as server:
                    server.ehlo()  # 增强握手
                    server.login(self.sender_email.get(), self.sender_password.get())
                    server.sendmail(self.sender_email.get(), to_email, msg.as_string())
                
                return True
                
            except smtplib.SMTPAuthenticationError:
                self.log("邮箱认证失败，请检查授权码和邮箱设置")
                return False
            except smtplib.SMTPConnectError:
                self.log(f"连接SMTP服务器失败（尝试 {attempt+1}/{SMTP_RETRY}）")
                if attempt < SMTP_RETRY - 1:
                    time.sleep(RETRY_DELAY)
            except Exception as e:
                self.log(f"邮件发送失败（尝试 {attempt+1}/{SMTP_RETRY}）: {str(e)}")
                if attempt < SMTP_RETRY - 1:
                    time.sleep(RETRY_DELAY)
        
        return False
    
    def send_email_with_attachments(self, part_files):
        """发送带多个附件的邮件（分卷）"""
        if not all([self.dest_email.get(), self.sender_email.get(), self.sender_password.get()]):
            self.log("请填写完整的邮箱信息")
            return False
        
        if not part_files or len(part_files) == 0:
            self.log("没有要发送的附件")
            return False
        
        try:
            total_parts = len(part_files)
            base_subject = f"文件备份 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            
            for i, part_file in enumerate(part_files):
                self.log(f"准备发送第 {i+1}/{total_parts} 部分...")
                
                # 邮件内容
                subject = f"{base_subject} (第 {i+1}/{total_parts} 部分)"
                body = f"这是自动备份的第 {i+1}/{total_parts} 部分。\n所有部分下载后放在同一文件夹，解压第一个文件即可合并。"
                
                # 发送邮件
                if not self.send_single_email(
                    self.dest_email.get(),
                    subject,
                    body,
                    part_file
                ):
                    self.log(f"第 {i+1} 部分发送失败，终止发送流程")
                    return False
                
                self.log(f"第 {i+1}/{total_parts} 部分发送成功")
                
                # 控制发送频率，避免触发限制
                if i < total_parts - 1:
                    self.log(f"等待 {SEND_DELAY} 秒后发送下一部分...")
                    time.sleep(SEND_DELAY)
            
            return True
        except Exception as e:
            self.log(f"邮件发送过程出错: {str(e)}")
            return False
    
    def test_email_connection(self):
        """测试邮箱连接是否正常"""
        self.log("开始测试邮箱连接...")
        
        if not all([self.sender_email.get(), self.sender_password.get()]):
            messagebox.showerror("错误", "请填写发件人邮箱和授权码")
            return
        
        try:
            with smtplib.SMTP_SSL("smtp.qq.com", 465, timeout=10) as server:
                server.ehlo()
                server.login(self.sender_email.get(), self.sender_password.get())
                self.log("邮箱连接测试成功！")
                messagebox.showinfo("成功", "邮箱连接测试成功")
        except smtplib.SMTPAuthenticationError:
            self.log("邮箱认证失败，请检查授权码是否正确")
            messagebox.showerror("错误", "邮箱认证失败，请检查授权码是否正确")
        except smtplib.SMTPConnectError:
            self.log("无法连接到SMTP服务器，请检查网络或防火墙设置")
            messagebox.showerror("错误", "无法连接到SMTP服务器，请检查网络或防火墙设置")
        except Exception as e:
            self.log(f"邮箱测试失败: {str(e)}")
            messagebox.showerror("错误", f"邮箱测试失败: {str(e)}")
    
    def backup_now(self):
        """立即执行备份"""
        self.log("开始执行备份...")
        
        # 创建分卷压缩文件
        part_files = self.create_split_zip()
        if not part_files:
            return
        
        # 发送邮件
        if self.send_email_with_attachments(part_files):
            # 备份成功后删除临时文件
            for part_file in part_files:
                try:
                    os.remove(part_file)
                    self.log(f"临时文件已删除: {os.path.basename(part_file)}")
                except Exception as e:
                    self.log(f"删除临时文件失败: {str(e)}")
        
        self.log("备份操作完成")
    
    def start_service(self):
        """启动或停止备份服务"""
        if not self.is_running:
            # 检查必要配置
            if not all([self.source_folder.get(), self.dest_email.get(), 
                       self.sender_email.get(), self.sender_password.get()]):
                messagebox.showerror("错误", "请填写完整的配置信息")
                return
            
            try:
                interval = float(self.backup_interval.get())
                if interval <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("错误", "请输入有效的备份间隔时间")
                return
            
            # 启动服务
            self.is_running = True
            self.start_button.config(text="停止备份服务")
            self.status_var.set(f"运行中 (每 {interval} 小时)")
            self.status_var.set(fg="green")
            self.log(f"备份服务已启动，每 {interval} 小时执行一次")
            
            # 设置定时任务
            schedule.clear()
            schedule.every(interval).hours.do(self.backup_now)
        else:
            # 停止服务
            self.is_running = False
            self.start_button.config(text="开始备份服务")
            self.status_var.set("未运行")
            self.status_var.set(fg="red")
            self.log("备份服务已停止")
            
            # 清除定时任务
            schedule.clear()
    
    def run_scheduler(self):
        """运行定时任务调度器"""
        while True:
            if self.is_running:
                schedule.run_pending()
            time.sleep(1)
    
    def start_scheduler_thread(self):
        """启动定时任务线程"""
        self.backup_thread = threading.Thread(target=self.run_scheduler, daemon=True)
        self.backup_thread.start()

if __name__ == "__main__":
    # 隐藏控制台窗口
    hide_console()
    
    root = Tk()
    app = BackupTool(root)
    root.mainloop()
    