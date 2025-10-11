import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import platform
import re
import datetime
import os
import threading
import queue 
import time
import webbrowser 
import random 

# 导入底层 Windows 模块
try:
    # 尝试导入 Windows 注册表和网络接口库
    import winreg
    import netifaces 
except ImportError:
    winreg = None
    netifaces = None
    
# --- 全局常量 ---
# Windows 注册表中网络适配器类的 GUID
NIC_REG_KEY = r"SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}"
# 线程间通信队列
load_queue = queue.Queue()
# 用于缓存 GUID 到友好名称的映射
GUID_TO_NAME_MAP = {} 

# --- 核心逻辑：MAC 地址工具函数 ---

def generate_random_mac():
    """
    生成一个符合本地管理地址（Locally Administered Address, LAA）规范的随机 MAC 地址。
    """
    # 确保第二个十六进制位是 2, 6, A, 或 E，满足 LAA 规范
    random_bytes = [random.randint(0x00, 0xff) for _ in range(5)]
    
    first_byte = random.randint(0x00, 0xFF)
    laa_marker = random.choice([0x02, 0x06, 0x0A, 0x0E])
    
    # 清除第一个字节的低 4 位，然后与 LAA 标记合并
    first_byte = (first_byte & 0xF0) | laa_marker
    
    # 组合为 6 个字节的列表
    mac_bytes = [first_byte] + random_bytes
    
    # 格式化为 AA:BB:CC:DD:EE:FF 字符串
    return ":".join(f"{b:02X}" for b in mac_bytes)


# --- 核心逻辑：工具函数 (保持不变) ---

def get_nic_driver_desc(target_guid):
    """
    根据接口 GUID 快速查找其在注册表中的 DriverDesc (友好描述)。
    这是为了提供用户友好的显示。
    """
    target_guid = target_guid.strip('{}').upper()
    if target_guid in GUID_TO_NAME_MAP:
        return GUID_TO_NAME_MAP[target_guid]

    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, NIC_REG_KEY) as root_key:
            i = 0
            while True:
                try:
                    # 遍历 0000, 0001, 0002... 来找到匹配的 GUID
                    subkey_name = "%04d" % i
                    with winreg.OpenKey(root_key, subkey_name) as sub_key:
                        # NetCfgInstanceId 是 GUID
                        guid = winreg.QueryValueEx(sub_key, "NetCfgInstanceId")[0].upper().strip('{}')
                        
                        if guid == target_guid:
                            # 找到了注册表键，尝试读取 DriverDesc
                            try:
                                driver_desc = winreg.QueryValueEx(sub_key, "DriverDesc")[0]
                                GUID_TO_NAME_MAP[target_guid] = driver_desc
                                return driver_desc
                            except:
                                # 没有 DriverDesc 的网卡
                                GUID_TO_NAME_MAP[target_guid] = "未知适配器"
                                return "未知适配器"
                    i += 1
                except OSError: # 遍历到末尾
                    break
        
        # 遍历完成也没找到
        GUID_TO_NAME_MAP[target_guid] = "未识别适配器"
        return "未识别适配器"
        
    except Exception:
        # 注册表操作失败或权限问题
        return "查询失败"


def get_nic_reg_key_path(target_guid):
    """
    根据接口 GUID 查找其在注册表中的键路径 (例如 '0001')。
    这是修改 MAC 地址的必要步骤。
    """
    target_guid = target_guid.strip('{}').upper()

    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, NIC_REG_KEY) as root_key:
            i = 0
            while True:
                try:
                    subkey_name = "%04d" % i
                    with winreg.OpenKey(root_key, subkey_name) as sub_key:
                        guid = winreg.QueryValueEx(sub_key, "NetCfgInstanceId")[0].upper().strip('{}')
                        if guid == target_guid:
                            return subkey_name
                    i += 1
                except OSError:
                    break
        return None
    except Exception:
        return None

# --- 核心逻辑：获取网卡信息 (异步执行) ---

def get_network_interfaces_and_mac_threaded():
    """
    异步执行：获取网卡列表和MAC地址。
    结果通过 load_queue 返回。
    """
    interfaces = {} # { "友好名称 (ID: XXXXXXXX...)": "AA:BB:CC:DD:EE:FF", ... }
    error_msg = None
    
    if platform.system() != "Windows" or netifaces is None:
        error_msg = "错误：系统不支持或 netifaces 未正确安装。"
        load_queue.put((interfaces, error_msg))
        return

    try:
        # 遍历所有接口 ID
        for iface_id in netifaces.interfaces():
            addrs = netifaces.ifaddresses(iface_id)
            
            if netifaces.AF_LINK in addrs:
                mac_address = addrs[netifaces.AF_LINK][0].get('addr', '').upper().replace('-', ':')
                
                # 排除无效地址
                if mac_address and mac_address != '00:00:00:00:00:00' and mac_address != 'N/A':
                    
                    # 尝试同步获取友好名称 (DriverDesc)
                    driver_desc = get_nic_driver_desc(iface_id)
                    
                    # 提取 GUID 的前缀，用于显示
                    guid_prefix = iface_id.strip('{}')[:8].upper()
                    
                    # 最终显示名称格式：友好名称 (ID: XXXXXXXX...)
                    display_name = f"{driver_desc} (ID: {guid_prefix}...)"
                    
                    # 存储到字典，键是显示名称，值是 MAC 地址
                    interfaces[display_name] = mac_address

        if not interfaces:
             error_msg = "未找到任何非回环网络接口。请检查驱动和连接。"
             
    except Exception as e:
        error_msg = f"获取接口信息失败 (底层调用): {e}"
    
    load_queue.put((interfaces, error_msg))


# --- 核心逻辑：修改 MAC 地址 (保持不变) ---

def change_mac_address(interface_display_name, new_mac):
    """
    执行 MAC 地址修改操作（操作注册表并重启网卡）。
    使用 PowerShell 的 InterfaceDescription (即 DriverDesc) 进行重启，以获得最大的兼容性。
    """
    # 格式验证
    if len(new_mac) != 17 or new_mac.count(':') != 5:
        return False, "MAC地址格式不正确 (例如: 00:11:22:33:44:55)", "N/A"

    original_mac = MacChangerApp.instance.original_mac_label.cget("text")
    
    # 1. 从显示名称中提取 GUID 前缀 (ID: XXXXXXXX)
    prefix_match = re.search(r'ID: ([0-9A-F]{8})', interface_display_name, re.IGNORECASE)
    
    if not prefix_match:
        return False, "无法从接口名称中解析出 GUID。", original_mac
    
    guid_prefix = prefix_match.group(1).upper()
    
    # 2. 通过 GUID 前缀重新查找完整的 GUID 和注册表键路径，并获取 DriverDesc
    target_guid = None
    reg_key_path = None
    driver_desc_for_powershell = None 
    
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, NIC_REG_KEY) as root_key:
            i = 0
            while True:
                try:
                    subkey_name = "%04d" % i
                    with winreg.OpenKey(root_key, subkey_name) as sub_key:
                        full_guid = winreg.QueryValueEx(sub_key, "NetCfgInstanceId")[0].upper().strip('{}')
                        
                        if full_guid.startswith(guid_prefix):
                            target_guid = full_guid 
                            reg_key_path = subkey_name 
                            # 获取 DriverDesc，这是 PowerShell NetAdapter 命令最通用的定位方式
                            try:
                                driver_desc_for_powershell = winreg.QueryValueEx(sub_key, "DriverDesc")[0]
                            except:
                                driver_desc_for_powershell = interface_display_name.split(' (ID:')[0].strip() # 使用友好名称部分
                            break
                    i += 1
                except OSError:
                    break
    except Exception as e:
        return False, f"注册表查询失败: {e}", original_mac

    
    if not target_guid or not reg_key_path:
        return False, f"未在注册表中找到与前缀 {guid_prefix} 匹配的完整 GUID。", original_mac
    
    if not driver_desc_for_powershell:
         return False, "无法获取网卡的 DriverDesc (友好名称) 作为重启参数。", original_mac
    
    # 准备修改 MAC 地址的完整注册表路径
    full_key = os.path.join(NIC_REG_KEY, reg_key_path)

    # 为了避免 PowerShell 解析中的引号和特殊字符问题，对名称进行转义
    safe_driver_desc = driver_desc_for_powershell.replace('"', '`"') 
    
    try:
        # 3. 修改注册表 (MAC地址修改成功的部分，保持不变)
        mac_no_separator = new_mac.replace(':', '').replace('-', '').upper()
        
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, full_key, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, 'NetworkAddress', 0, winreg.REG_SZ, mac_no_separator)
            
        # 4. 禁用/启用网卡使修改生效 (改用 PowerShell + InterfaceDescription)
        
        # 禁用网卡
        disable_cmd = f'powershell -command "Disable-NetAdapter -InterfaceDescription \\"{safe_driver_desc}\\" -Confirm:$false"'
        subprocess.check_output(disable_cmd, shell=True, stderr=subprocess.STDOUT)
        
        time.sleep(1.5) # 增加延迟，确保禁用完成
        
        # 启用网卡
        enable_cmd = f'powershell -command "Enable-NetAdapter -InterfaceDescription \\"{safe_driver_desc}\\" -Confirm:$false"'
        subprocess.check_output(enable_cmd, shell=True, stderr=subprocess.STDOUT)
        
        log_message = f"Windows: 成功修改 {driver_desc_for_powershell} 的MAC地址并使用 PowerShell 重启网卡。"
        return True, log_message, original_mac

    except subprocess.CalledProcessError as e:
        # PowerShell 命令失败处理
        error_output = e.output.decode('gbk', errors='ignore').strip()
        
        error_msg = f"网卡重启失败 (PowerShell)。\n请手动禁用/启用网卡！\n接口描述: {driver_desc_for_powershell}\n错误: {error_output}"
        return False, error_msg, original_mac
    except PermissionError:
        error_msg = "权限不足！请以 **管理员身份** 运行本程序。"
        return False, error_msg, original_mac
    except Exception as e:
        error_msg = f"修改或注册表操作失败: {e}"
        return False, error_msg, original_mac


# --- 日志记录功能 (保持不变) ---
def save_log(log_entry):
    """将日志条目写入文件"""
    log_file = "mac_change_log.txt"
    try:
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(log_entry + "\n")
    except Exception as e:
        messagebox.showerror("日志错误", f"无法写入日志文件 '{log_file}': {e}")


# --- GUI 界面和事件处理 ---

class MacChangerApp:
    instance = None 
    
    def center_window(self, width, height):
        """计算并设置窗口在屏幕中央的位置"""
        # 获取屏幕宽度和高度
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()

        # 计算窗口的 X, Y 坐标
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)

        # 设置窗口位置
        self.master.geometry(f'{width}x{height}+{x}+{y}')

    def __init__(self, master):
        MacChangerApp.instance = self
        self.master = master
        
        # 窗口的固定大小
        WINDOW_WIDTH = 600
        WINDOW_HEIGHT = 520
        
        master.title("MAC 地址修改工具")
        master.resizable(False, False)
        
        # 在设置窗口大小和组件之前，先调用 update_idletasks 确保窗口信息就绪
        master.update_idletasks()
        self.center_window(WINDOW_WIDTH, WINDOW_HEIGHT)
        
        # --- 1. 设置图标 ---
        try:
            self.master.iconbitmap('mac.ico')
        except Exception:
            pass
        
        self.current_interfaces = {} 
        self.interface_load_thread = None

        style = ttk.Style()
        style.configure('TFrame', background='#f0f0f0')
        style.configure('TLabel', background='#f0f0f0', font=('Arial', 10))
        style.configure('TButton', font=('Arial', 10, 'bold'))

        main_frame = ttk.Frame(master, padding="10")
        main_frame.pack(fill='both', expand=True)

        # ----------------------------------------------------
        # 1. 网络接口选择
        # ----------------------------------------------------
        ttk.Label(main_frame, text="选择网络接口:").grid(row=0, column=0, sticky='w', pady=5, padx=5)
        self.interface_var = tk.StringVar()
        self.interface_combobox = ttk.Combobox(main_frame, textvariable=self.interface_var, state='readonly', width=40)
        self.interface_combobox.grid(row=0, column=1, sticky='ew', pady=5, padx=5)
        self.interface_combobox.bind('<<ComboboxSelected>>', self.on_interface_select)

        # ----------------------------------------------------
        # 2. 原MAC地址显示
        # ----------------------------------------------------
        ttk.Label(main_frame, text="原 MAC 地址:").grid(row=1, column=0, sticky='w', pady=5, padx=5)
        self.original_mac_label = ttk.Label(main_frame, text="N/A", foreground='blue', width=40, cursor="hand2") 
        self.original_mac_label.grid(row=1, column=1, sticky='w', pady=5, padx=5)
        self.original_mac_label.bind("<Double-Button-1>", self.copy_original_mac) 

        # ----------------------------------------------------
        # 3. 新 MAC 地址输入 & 随机按钮
        # ----------------------------------------------------
        ttk.Label(main_frame, text="新 MAC 地址:").grid(row=2, column=0, sticky='w', pady=5, padx=5)
        
        mac_input_frame = ttk.Frame(main_frame)
        mac_input_frame.grid(row=2, column=1, sticky='ew', pady=5, padx=5)
        mac_input_frame.columnconfigure(0, weight=1) 
        
        self.mac_entry = ttk.Entry(mac_input_frame, width=35)
        self.mac_entry.grid(row=0, column=0, sticky='ew', padx=(0, 5))
        self.mac_entry.insert(0, generate_random_mac()) 
        
        self.random_button = ttk.Button(mac_input_frame, text="随机生成", command=self.on_generate_random_mac)
        self.random_button.grid(row=0, column=1, sticky='e')
        
        # ----------------------------------------------------
        # 4. 执行按钮
        # ----------------------------------------------------
        self.change_button = ttk.Button(main_frame, text="执行修改", command=self.on_change_click, state='disabled')
        self.change_button.grid(row=3, column=0, columnspan=2, pady=15)
        
        # ----------------------------------------------------
        # 5. 状态和日志
        # ----------------------------------------------------
        self.status_label = ttk.Label(main_frame, text="程序启动，正在加载网卡信息...", foreground='black')
        self.status_label.grid(row=4, column=0, columnspan=2, pady=5)
        
        ttk.Label(main_frame, text="Windows：请务必以 **管理员身份** 运行本程序！", foreground='red', font=('Arial', 9, 'bold')).grid(row=5, column=0, columnspan=2, pady=(0, 5))
        
        ttk.Label(main_frame, text="--- 操作日志 (双击原MAC地址可复制) ---").grid(row=6, column=0, columnspan=2, pady=(10, 5))
        self.log_text = tk.Text(main_frame, height=8, width=70, state='disabled', wrap='word', font=('Courier', 9))
        self.log_text.grid(row=7, column=0, columnspan=2, sticky='nsew', padx=5)
        
        main_frame.grid_columnconfigure(1, weight=1)
        main_frame.grid_rowconfigure(7, weight=1)
        
        # ----------------------------------------------------
        # 6. 底部联系方式
        # ----------------------------------------------------
        contact_frame = ttk.Frame(main_frame, padding="5")
        contact_frame.grid(row=8, column=0, columnspan=2, sticky='ew', pady=(10, 0))
        
        ttk.Label(contact_frame, text="联系方式：QQ 88179096", font=('Arial', 9)).pack(side='left', padx=5)
        
        self.link_label = ttk.Label(contact_frame, text="www.itvip.com.cn", 
                                    foreground="blue", cursor="hand2", font=('Arial', 9, 'underline'))
        self.link_label.pack(side='left', padx=15)
        self.link_label.bind("<Button-1>", lambda e: self.open_link("http://www.itvip.com.cn"))

        contact_frame.columnconfigure(0, weight=1) 

        # 启动异步加载
        self.start_load_interfaces()
        self.master.after(100, self.check_load_interfaces)

    def open_link(self, url):
        """用默认浏览器打开指定的 URL"""
        webbrowser.open_new_tab(url)

    def copy_original_mac(self, event):
        """双击事件：复制当前显示的 MAC 地址到剪贴板"""
        mac_address = self.original_mac_label.cget("text")
        
        if mac_address != "N/A":
            self.master.clipboard_clear() 
            self.master.clipboard_append(mac_address) 
            self.log_message(f"[信息] 已将 MAC 地址 {mac_address} 复制到剪贴板。")
        else:
            self.log_message("[警告] 剪贴板复制失败：当前没有可用的 MAC 地址。")

    def on_generate_random_mac(self):
        """点击“生成随机 MAC”按钮"""
        new_mac = generate_random_mac()
        self.mac_entry.delete(0, tk.END)
        self.mac_entry.insert(0, new_mac)
        self.log_message(f"[信息] 已生成新的随机 MAC 地址：{new_mac}")

    def start_load_interfaces(self):
        """在后台线程启动网卡信息加载"""
        self.interface_load_thread = threading.Thread(target=get_network_interfaces_and_mac_threaded, daemon=True)
        self.interface_load_thread.start()
        self.log_message("[信息] 已启动后台线程加载网卡信息...")
        self.change_button.config(state='disabled')


    def check_load_interfaces(self):
        """定时检查线程是否完成加载"""
        try:
            interfaces, error_msg = load_queue.get_nowait()
            self.finalize_load_interfaces(interfaces, error_msg)
            
        except queue.Empty:
            self.master.after(100, self.check_load_interfaces)


    def finalize_load_interfaces(self, interfaces, error_msg):
        """加载完成后，更新 GUI"""
        self.current_interfaces = interfaces
        interface_names = list(interfaces.keys())
        
        if error_msg:
             self.log_message(f"[警告] {error_msg}")
        
        self.interface_combobox['values'] = interface_names
        
        if interface_names:
            self.interface_combobox.current(0) 
            self.on_interface_select(None)
            self.status_label.config(text="网卡加载成功，请选择接口并修改。", foreground='green')
            self.change_button.config(state='normal')
        else:
            self.interface_var.set("无可用网卡")
            self.original_mac_label.config(text="N/A")
            self.status_label.config(text="未找到可用网卡，请检查权限/系统环境。", foreground='red')
            self.log_message("[错误] 未找到任何可用网卡。")

    def on_interface_select(self, event):
        """选中接口时，更新原始MAC地址显示"""
        selected_interface = self.interface_var.get()
        mac = self.current_interfaces.get(selected_interface, "N/A")
        
        self.original_mac_label.config(text=mac)
        
        if mac != "N/A":
            self.log_message(f"[信息] 接口 {selected_interface} 原始MAC地址为 {mac}")
        else:
            self.log_message(f"[错误] 无法获取接口 {selected_interface} 的MAC地址。")

    def on_change_click(self):
        """执行修改操作"""
        interface = self.interface_var.get()
        new_mac = self.mac_entry.get().strip()
        original_mac = self.original_mac_label.cget("text")
        
        if not interface or interface == "无可用网卡":
            messagebox.showerror("错误", "请选择一个有效的网络接口。")
            return
        
        if original_mac == new_mac:
            messagebox.showwarning("警告", "新的MAC地址与原地址相同，无需修改。")
            return

        self.status_label.config(text="正在尝试执行修改 (需要管理员权限)...", foreground='orange')
        self.master.update()

        success, message, original_mac_logged = change_mac_address(interface, new_mac)
        
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        if success:
            log_entry = f"[{timestamp}] [成功] {interface} | 原MAC: {original_mac_logged} | 新MAC: {new_mac}"
            self.status_label.config(text="修改成功！请重新加载网卡确认。", foreground='green')
            messagebox.showinfo("成功", f"MAC地址修改成功。\n接口：{interface}\n新MAC：{new_mac}")
            
            # 成功后重新加载网卡列表，以显示新的 MAC 地址
            self.start_load_interfaces()
            self.master.after(100, self.check_load_interfaces)
            
        else:
            log_entry = f"[{timestamp}] [失败] {interface} | 原MAC: {original_mac_logged} | 新MAC: {new_mac} | 错误: {message}"
            self.status_label.config(text="修改失败，请查看错误信息。", foreground='red')
            messagebox.showerror("失败", f"MAC地址修改失败:\n{message}")

        self.log_message(log_entry)
        save_log(log_entry)


    def log_message(self, message):
        """在GUI日志区显示信息"""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')


# --- 运行主程序 ---
if __name__ == "__main__":
    if platform.system() != "Windows":
        messagebox.showerror("系统不支持", "此工具专为 Windows 设计。")
    elif netifaces is None or winreg is None:
         messagebox.showerror("环境错误", "缺少必要的 Python 库 (netifaces) 或 Windows 模块。请确保已安装 netifaces。")
    else:
        root = tk.Tk()
        root.clipboard_clear() 
        app = MacChangerApp(root)
        root.mainloop()