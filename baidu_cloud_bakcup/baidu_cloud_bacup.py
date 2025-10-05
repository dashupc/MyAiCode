import os
import sys
import time
import schedule
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import logging
from datetime import datetime
import requests
import json
import qrcode
from io import BytesIO
from PIL import Image, ImageTk

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("netdisk_backup_log.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# 百度网盘API相关配置 - 请替换为你自己的API_KEY和SECRET_KEY
API_KEY = "9Ed03IJAgtoRMwljjSVUsFOpP6t3cEBe"  # 替换为你的AppKey
SECRET_KEY = "IthBI8HZfaVkjg5KsGz5BonW5Ivfe6l9"  # 替换为你的SecretKey
REDIRECT_URI = "oob"

# 包含完整的权限范围
AUTH_URL = f"https://openapi.baidu.com/oauth/2.0/authorize?response_type=code&client_id={API_KEY}&redirect_uri={REDIRECT_URI}&scope=basic,netdisk,netdisk.read,netdisk.write,netdisk.file"
TOKEN_URL = "https://openapi.baidu.com/oauth/2.0/token"
UPLOAD_URL = "https://pan.baidu.com/rest/2.0/xpan/file?method=upload"
MKDIR_URL = "https://pan.baidu.com/rest/2.0/xpan/file?method=mkdir"
LIST_URL = "https://pan.baidu.com/rest/2.0/xpan/file?method=list"

# 百度网盘API错误码解释
ERROR_CODE_EXPLANATIONS = {
    0: "成功",
    -1: "系统错误",
    -2: "服务暂不可用",
    -3: "未知错误",
    -4: "参数错误",
    -5: "鉴权失败",
    -6: "路径不存在",
    -7: "权限不足",
    -8: "数据不存在",
    110: "文件/目录已存在",
    111: "文件过大",
    112: "路径过长",
    113: "包含非法字符",
    114: "目录不为空",
    31064: "应用未获得文件操作权限，请在开放平台配置权限并重新授权"
}

class BaiduNetdiskBackupTool:
    def __init__(self, root):
        self.root = root
        self.root.title("百度网盘定时备份工具")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        # 设置中文字体支持
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("SimHei", 10))
        self.style.configure("TButton", font=("SimHei", 10))
        self.style.configure("TCombobox", font=("SimHei", 10))
        
        # 备份参数
        self.folder_to_backup = tk.StringVar()
        self.refresh_token = tk.StringVar()
        self.access_token = ""
        self.expires_in = 0
        self.token_expire_time = 0
        self.remote_folder = tk.StringVar(value="/备份")  # 默认备份目录
        self.backup_interval = tk.StringVar(value="1")  # 默认1小时
        self.interval_unit = tk.StringVar(value="小时")
        self.backup_running = False
        
        # 创建界面
        self.create_widgets()
        
        # 加载保存的配置
        self.load_config()
        
        # 检查并刷新token
        if self.refresh_token.get():
            threading.Thread(target=self.refresh_access_token, daemon=True).start()
    
    def create_widgets(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 百度网盘授权区域
        auth_frame = ttk.LabelFrame(main_frame, text="百度网盘授权", padding="10")
        auth_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(auth_frame, text="刷新令牌(Refresh Token):").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(auth_frame, textvariable=self.refresh_token, width=50).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Button(auth_frame, text="获取授权", command=self.show_auth_qrcode).grid(row=0, column=2, padx=10, pady=5)
        ttk.Button(auth_frame, text="验证授权", command=self.verify_auth).grid(row=0, column=3, padx=10, pady=5)
        
        # 备份路径配置
        path_frame = ttk.LabelFrame(main_frame, text="备份路径", padding="10")
        path_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(path_frame, text="本地文件夹:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(path_frame, textvariable=self.folder_to_backup, width=50).grid(row=0, column=1, sticky=tk.W, pady=5)
        ttk.Button(path_frame, text="浏览...", command=self.browse_folder).grid(row=0, column=2, padx=10, pady=5)
        
        ttk.Label(path_frame, text="网盘文件夹:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(path_frame, textvariable=self.remote_folder, width=50).grid(row=1, column=1, sticky=tk.W, pady=5)
        ttk.Button(path_frame, text="创建文件夹", command=self.create_remote_folder).grid(row=1, column=2, padx=10, pady=5)
        
        # 定时设置
        timer_frame = ttk.LabelFrame(main_frame, text="定时设置", padding="10")
        timer_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(timer_frame, text="备份间隔:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(timer_frame, textvariable=self.backup_interval, width=10).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(timer_frame, text="时间单位:").grid(row=0, column=2, sticky=tk.W, pady=5, padx=5)
        interval_combobox = ttk.Combobox(
            timer_frame, 
            textvariable=self.interval_unit, 
            values=["分钟", "小时", "天"], 
            state="readonly",
            width=8
        )
        interval_combobox.grid(row=0, column=3, sticky=tk.W, pady=5)
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame, padding="10")
        button_frame.pack(fill=tk.X, pady=5)
        
        self.start_button = ttk.Button(button_frame, text="开始备份", command=self.start_backup)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="停止备份", command=self.stop_backup, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        self.backup_now_button = ttk.Button(button_frame, text="立即备份", command=self.backup_now)
        self.backup_now_button.pack(side=tk.LEFT, padx=5)
        
        self.save_config_button = ttk.Button(button_frame, text="保存配置", command=self.save_config)
        self.save_config_button.pack(side=tk.LEFT, padx=5)
        
        # 状态区域
        status_frame = ttk.LabelFrame(main_frame, text="状态日志", padding="10")
        status_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 日志文本框
        self.log_text = tk.Text(status_frame, height=10, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(status_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 底部状态
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_to_backup.set(folder)
            self.log("已选择备份文件夹: " + folder)
    
    def log(self, message):
        """在日志文本框中显示消息"""
        self.log_text.config(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)  # 滚动到最后
        self.log_text.config(state=tk.DISABLED)
        logging.info(message)
    
    def save_config(self):
        """保存配置到文件"""
        try:
            with open("netdisk_backup_config.txt", "w", encoding="utf-8") as f:
                f.write(f"REFRESH_TOKEN={self.refresh_token.get()}\n")
                f.write(f"FOLDER={self.folder_to_backup.get()}\n")
                f.write(f"REMOTE_FOLDER={self.remote_folder.get()}\n")
                f.write(f"INTERVAL={self.backup_interval.get()}\n")
                f.write(f"UNIT={self.interval_unit.get()}\n")
            self.log("配置已保存")
            messagebox.showinfo("成功", "配置已保存")
        except Exception as e:
            self.log(f"保存配置失败: {str(e)}")
            messagebox.showerror("错误", f"保存配置失败: {str(e)}")
    
    def load_config(self):
        """从文件加载配置"""
        try:
            if os.path.exists("netdisk_backup_config.txt"):
                with open("netdisk_backup_config.txt", "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if line.startswith("REFRESH_TOKEN="):
                            self.refresh_token.set(line[14:])
                        elif line.startswith("FOLDER="):
                            self.folder_to_backup.set(line[7:])
                        elif line.startswith("REMOTE_FOLDER="):
                            self.remote_folder.set(line[15:])
                        elif line.startswith("INTERVAL="):
                            self.backup_interval.set(line[9:])
                        elif line.startswith("UNIT="):
                            self.interval_unit.set(line[5:])
                self.log("配置已加载")
        except Exception as e:
            self.log(f"加载配置失败: {str(e)}")
    
    def show_auth_qrcode(self):
        """显示优化尺寸的授权二维码窗口"""
        # 调整窗口大小，确保内容完整显示
        auth_window = tk.Toplevel(self.root)
        auth_window.title("百度网盘授权 - 扫码登录")
        auth_window.geometry("600x800")  # 适中的窗口尺寸
        auth_window.resizable(False, False)
        
        # 居中显示窗口
        auth_window.update_idletasks()
        width = auth_window.winfo_width()
        height = auth_window.winfo_height()
        x = (auth_window.winfo_screenwidth() // 2) - (width // 2)
        y = (auth_window.winfo_screenheight() // 2) - (height // 2)
        auth_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
        # 显示授权说明
        ttk.Label(
            auth_window, 
            text="请使用百度网盘APP扫描下方二维码授权", 
            font=("SimHei", 14, "bold")
        ).pack(pady=15)
        
        # 生成尺寸适中的二维码（约300x300像素）
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_H,  # 高容错率
            box_size=9,  # 调整方块大小，使二维码适中
            border=4,
        )
        qr.add_data(AUTH_URL)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white")
        
        # 转换为PhotoImage
        img_byte_arr = BytesIO()
        qr_img.save(img_byte_arr, format='PNG', dpi=(300, 300))
        img_byte_arr.seek(0)
        pil_img = Image.open(img_byte_arr)
        
        # 确保二维码在窗口内完全显示
        max_size = 350  # 最大尺寸限制
        img_width, img_height = pil_img.size
        if img_width > max_size or img_height > max_size:
            ratio = max_size / max(img_width, img_height)
            new_size = (int(img_width * ratio), int(img_height * ratio))
            pil_img = pil_img.resize(new_size, Image.LANCZOS)
        
        tk_img = ImageTk.PhotoImage(pil_img)
        
        # 显示二维码
        qr_frame = ttk.Frame(auth_window, padding=15)
        qr_frame.pack(pady=10)
        qr_label = ttk.Label(
            qr_frame, 
            image=tk_img, 
            borderwidth=3,
            relief="solid",
            padding=10
        )
        qr_label.image = tk_img
        qr_label.pack()
        
        # 扫码提示
        ttk.Label(
            auth_window, 
            text="提示：请确保百度网盘APP已登录正确账号", 
            font=("SimHei", 10, "italic"),
            foreground="#666666"
        ).pack(pady=5)
        
        # 输入授权码的区域
        ttk.Label(
            auth_window, 
            text="如果无法扫描，请访问以下网址获取授权码:", 
            font=("SimHei", 11)
        ).pack(pady=10)
        
        # 创建可滚动的URL标签
        url_frame = ttk.Frame(auth_window)
        url_frame.pack(fill=tk.X, padx=30, pady=5)
        
        url_scrollbar = ttk.Scrollbar(url_frame, orient=tk.HORIZONTAL)
        url_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        url_label = ttk.Label(
            url_frame, 
            text=AUTH_URL, 
            foreground="blue", 
            cursor="hand2",
            wraplength=500
        )
        url_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        url_label.bind("<Button-1>", lambda e: self.log(f"授权网址: {AUTH_URL}"))
        
        ttk.Label(
            auth_window, 
            text="请输入授权码:", 
            font=("SimHei", 11)
        ).pack(pady=10)
        
        code_frame = ttk.Frame(auth_window, padding=5)
        code_frame.pack(fill=tk.X, padx=30)
        
        code_entry = ttk.Entry(code_frame, width=50, font=("SimHei", 14))
        code_entry.pack(fill=tk.X, expand=True)
        code_entry.focus()
        
        # 确认按钮
        def confirm_auth():
            code = code_entry.get().strip()
            if code:
                self.get_access_token(code)
                auth_window.destroy()
        
        confirm_btn = ttk.Button(
            auth_window, 
            text="确认授权", 
            command=confirm_auth,
            style="Accent.TButton"
        )
        self.style.configure("Accent.TButton", font=("SimHei", 12, "bold"), padding=10)
        confirm_btn.pack(pady=15)
    
    def get_access_token(self, code):
        """使用授权码获取访问令牌"""
        try:
            self.log("正在获取访问令牌...")
            params = {
                "grant_type": "authorization_code",
                "code": code,
                "client_id": API_KEY,
                "client_secret": SECRET_KEY,
                "redirect_uri": REDIRECT_URI
            }
            
            response = requests.get(TOKEN_URL, params=params)
            self.log(f"令牌请求响应: 状态码={response.status_code}, 内容={response.text}")
            
            result = json.loads(response.text)
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                self.refresh_token.set(result["refresh_token"])
                self.expires_in = result["expires_in"]
                self.token_expire_time = time.time() + self.expires_in - 300  # 提前5分钟刷新
                self.log("授权成功，已获取访问令牌")
                self.save_config()  # 保存refresh_token
                return True
            else:
                error_msg = result.get('error_description', '未知错误')
                self.log(f"授权失败: {error_msg}")
                return False
        except Exception as e:
            self.log(f"获取访问令牌失败: {str(e)}")
            return False
    
    def refresh_access_token(self):
        """刷新访问令牌"""
        if not self.refresh_token.get():
            self.log("没有刷新令牌，无法刷新访问令牌")
            return False
        
        try:
            self.log("正在刷新访问令牌...")
            params = {
                "grant_type": "refresh_token",
                "refresh_token": self.refresh_token.get(),
                "client_id": API_KEY,
                "client_secret": SECRET_KEY
            }
            
            response = requests.get(TOKEN_URL, params=params)
            self.log(f"刷新令牌响应: 状态码={response.status_code}, 内容={response.text}")
            
            result = json.loads(response.text)
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                # 更新refresh_token（如果有新的）
                if "refresh_token" in result:
                    self.refresh_token.set(result["refresh_token"])
                    self.save_config()
                self.expires_in = result["expires_in"]
                self.token_expire_time = time.time() + self.expires_in - 300  # 提前5分钟刷新
                self.log("访问令牌已刷新")
                return True
            else:
                error_msg = result.get('error_description', '未知错误')
                self.log(f"刷新令牌失败: {error_msg}")
                return False
        except Exception as e:
            self.log(f"刷新访问令牌失败: {str(e)}")
            return False
    
    def verify_auth(self):
        """验证授权是否有效"""
        if not self.refresh_token.get():
            self.log("请先获取授权")
            return False
            
        if time.time() > self.token_expire_time:
            return self.refresh_access_token()
        return True
    
    def create_remote_folder(self):
        """在百度网盘中创建文件夹"""
        if not self.verify_auth():
            self.log("授权无效，请重新授权")
            return False
        
        folder_path = self.remote_folder.get()
        if not folder_path:
            self.log("请输入网盘文件夹路径")
            return False
        
        # 验证路径格式
        if not folder_path.startswith('/'):
            self.log(f"路径格式错误: 路径必须以'/'开头，当前路径: {folder_path}")
            self.log("自动修正路径格式...")
            folder_path = '/' + folder_path.lstrip('/')
            self.remote_folder.set(folder_path)
            self.log(f"修正后的路径: {folder_path}")
        
        # 检查路径中是否包含非法字符
        invalid_chars = ['?', '*', ':', '"', '<', '>', '|']
        for char in invalid_chars:
            if char in folder_path:
                self.log(f"路径包含非法字符 '{char}'，请修改路径")
                return False
        
        try:
            self.log(f"正在创建网盘文件夹: {folder_path}")
            params = {
                "access_token": self.access_token,
                "path": folder_path,
                "isdir": 1
            }
            
            # 记录发送的请求参数（隐藏部分token）
            log_params = params.copy()
            if 'access_token' in log_params:
                log_params['access_token'] = log_params['access_token'][:10] + '...'
            self.log(f"发送创建文件夹请求: {log_params}")
            
            response = requests.post(MKDIR_URL, params=params)
            
            # 记录详细响应信息
            self.log(f"API响应状态码: {response.status_code}")
            self.log(f"API响应内容: {response.text}")
            
            # 处理响应
            try:
                result = json.loads(response.text)
            except json.JSONDecodeError:
                self.log("API返回的响应不是有效的JSON格式")
                return False
            
            # 解析错误码
            errno = result.get("error_code", result.get("errno", -3))
            errmsg = result.get("error_msg", result.get("errmsg", "未知错误"))
            
            error_explanation = ERROR_CODE_EXPLANATIONS.get(errno, "未知错误")
            self.log(f"创建文件夹失败: 错误码={errno}, 错误信息={errmsg}")
            self.log(f"错误解释: {error_explanation}")
            
            # 根据错误码提供解决方案
            if errno == 0:
                self.log(f"文件夹创建成功: {folder_path}")
                return True
            elif errno == 110:
                self.log(f"文件夹已存在: {folder_path}")
                return True
            elif errno == -6:
                self.log("解决方案: 请先创建上级目录，或检查路径是否正确")
            elif errno == -7 or errno == 31064:
                self.log("解决方案: 1. 登录百度网盘开放平台(https://pan.baidu.com/union)")
                self.log("           2. 进入应用管理，为你的应用添加netdisk相关权限")
                self.log("           3. 重新点击本程序的'获取授权'按钮完成授权")
            elif errno == 113:
                self.log("解决方案: 请移除路径中的特殊字符，如 ?*:\"<>| 等")
            elif errno == 112:
                self.log("解决方案: 路径过长，请缩短文件夹路径")
            
            return False
        except Exception as e:
            self.log(f"创建文件夹失败: {str(e)}")
            return False
    
    def upload_file(self, local_path, remote_path):
        """上传文件到百度网盘"""
        if not self.verify_auth():
            self.log("授权无效，请重新授权")
            return False
        
        try:
            # 先检查文件是否已存在
            remote_dir = os.path.dirname(remote_path)
            file_name = os.path.basename(remote_path)
            
            # 列出远程目录下的文件
            params = {
                "access_token": self.access_token,
                "path": remote_dir,
                "web": 1
            }
            
            response = requests.get(LIST_URL, params=params)
            self.log(f"检查文件存在性响应: 状态码={response.status_code}, 内容={response.text}")
            
            result = json.loads(response.text)
            
            # 检查文件是否已存在
            file_exists = False
            if result.get("errno") == 0 and "list" in result:
                for item in result["list"]:
                    if item.get("server_filename") == file_name and item.get("isdir") == 0:
                        # 检查文件大小是否相同
                        local_size = os.path.getsize(local_path)
                        if item.get("size") == local_size:
                            self.log(f"文件已存在且大小相同，跳过上传: {file_name}")
                            file_exists = True
                            break
            
            if file_exists:
                return True
            
            # 上传文件
            self.log(f"正在上传文件: {local_path} -> {remote_path}")
            
            params = {
                "access_token": self.access_token,
                "path": remote_path,
                "filename": os.path.basename(local_path),
                "isdir": 0
            }
            
            with open(local_path, "rb") as f:
                files = {"file": f}
                response = requests.post(UPLOAD_URL, params=params, files=files)
                self.log(f"文件上传响应: 状态码={response.status_code}, 内容={response.text}")
                
                result = json.loads(response.text)
                
                if result.get("errno") == 0:
                    self.log(f"文件上传成功: {file_name}")
                    return True
                else:
                    errno = result.get("error_code", result.get("errno", -3))
                    errmsg = result.get("error_msg", result.get("errmsg", "未知错误"))
                    error_explanation = ERROR_CODE_EXPLANATIONS.get(errno, "未知错误")
                    self.log(f"文件上传失败: 错误码={errno}, 错误信息={errmsg}, 解释={error_explanation}")
                    return False
        except Exception as e:
            self.log(f"文件上传失败: {str(e)}")
            return False
    
    def backup_files(self):
        """备份文件夹中的所有文件"""
        folder = self.folder_to_backup.get()
        if not folder or not os.path.isdir(folder):
            self.log("请选择有效的备份文件夹")
            return False
        
        # 确保远程文件夹存在
        if not self.create_remote_folder():
            return False
        
        self.log("开始执行备份...")
        self.status_var.set("正在备份...")
        
        try:
            # 获取当前时间作为备份目录的一部分
            backup_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_remote_path = f"{self.remote_folder.get()}/backup_{backup_time}"
            
            # 遍历文件夹中的所有文件
            for root_dir, dirs, files in os.walk(folder):
                # 创建对应的远程目录
                relative_dir = os.path.relpath(root_dir, folder)
                remote_dir = f"{base_remote_path}/{relative_dir}"
                
                # 处理目录分隔符
                remote_dir = remote_dir.replace(os.sep, "/")
                
                # 创建远程目录
                self.remote_folder.set(remote_dir)
                if not self.create_remote_folder():
                    self.log(f"无法创建远程目录 {remote_dir}，跳过该目录下的文件")
                    continue
                
                # 上传文件
                for file in files:
                    local_file_path = os.path.join(root_dir, file)
                    remote_file_path = f"{remote_dir}/{file}".replace(os.sep, "/")
                    
                    # 上传文件
                    if not self.upload_file(local_file_path, remote_file_path):
                        self.log(f"继续处理其他文件...")
            
            # 恢复远程文件夹设置
            original_remote_folder = self.remote_folder.get().split("/backup_")[0]
            self.remote_folder.set(original_remote_folder)
            
            self.log("备份完成")
            self.status_var.set("备份完成")
            return True
        except Exception as e:
            self.log(f"备份过程出错: {str(e)}")
            self.status_var.set("备份出错")
            return False
    
    def backup_now(self):
        """立即执行备份（在新线程中）"""
        threading.Thread(target=self.backup_files, daemon=True).start()
    
    def start_backup(self):
        """开始定时备份"""
        try:
            interval = int(self.backup_interval.get())
            if interval <= 0:
                messagebox.showerror("错误", "间隔时间必须大于0")
                return
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字作为间隔时间")
            return
        
        # 验证必要的配置
        if not self.refresh_token.get():
            messagebox.showerror("错误", "请先完成百度网盘授权")
            return
        
        if not self.folder_to_backup.get() or not os.path.isdir(self.folder_to_backup.get()):
            messagebox.showerror("错误", "请选择有效的备份文件夹")
            return
        
        # 验证授权
        if not self.verify_auth():
            messagebox.showerror("错误", "授权验证失败，请重新授权")
            return
        
        # 设置定时任务
        unit = self.interval_unit.get()
        self.log(f"开始定时备份，每{interval}{unit}执行一次")
        
        if unit == "分钟":
            schedule.every(interval).minutes.do(self.backup_files)
        elif unit == "小时":
            schedule.every(interval).hours.do(self.backup_files)
        elif unit == "天":
            schedule.every(interval).days.do(self.backup_files)
        
        self.backup_running = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.status_var.set(f"定时备份中 (每{interval}{unit})")
        
        # 启动定时任务线程
        threading.Thread(target=self.run_scheduler, daemon=True).start()
        
        # 立即执行一次备份
        self.backup_now()
    
    def stop_backup(self):
        """停止定时备份"""
        self.backup_running = False
        schedule.clear()
        self.log("已停止定时备份")
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.status_var.set("已停止")
    
    def run_scheduler(self):
        """运行定时任务调度器"""
        while self.backup_running:
            schedule.run_pending()
            time.sleep(1)

if __name__ == "__main__":
    # 检查是否安装了必要的库
    required_packages = ["requests", "qrcode", "PIL", "schedule"]
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print(f"请先安装缺少的库: pip install {' '.join(missing_packages)}")
        sys.exit(1)
    
    root = tk.Tk()
    app = BaiduNetdiskBackupTool(root)
    root.mainloop()
    