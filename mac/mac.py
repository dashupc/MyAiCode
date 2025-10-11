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
import sys

# --- 导入 Windows 底层模块 ---
try:
    import winreg
    import netifaces 
except ImportError:
    winreg = None
    netifaces = None
    
# --- 全局常量 ---
NIC_REG_KEY = r"SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}"
load_queue = queue.Queue()
GUID_TO_NAME_MAP = {} 

# --- 工具函数：用于处理打包后的图标路径 ---
def resource_path(relative_path):
    """获取资源的绝对路径，无论是从脚本运行还是从 PyInstaller 打包后的 EXE 运行。"""
    try:
        # PyInstaller 打包后的路径
        base_path = sys._MEIPASS
    except Exception:
        # 正常脚本运行时的路径
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- 核心逻辑：MAC/IP 地址工具函数 ---

def generate_random_mac():
    """生成一个符合本地管理地址（LAA）规范的随机 MAC 地址。"""
    random_bytes = [random.randint(0x00, 0xff) for _ in range(5)]
    first_byte = random.randint(0x00, 0xFF)
    laa_marker = random.choice([0x02, 0x06, 0x0A, 0x0E])
    first_byte = (first_byte & 0xF0) | laa_marker
    mac_bytes = [first_byte] + random_bytes
    return ":".join(f"{b:02X}" for b in mac_bytes)


def get_nic_driver_desc(target_guid):
    """根据接口 GUID 查找其在注册表中的 DriverDesc (友好描述)。"""
    target_guid = target_guid.strip('{}').upper()
    if target_guid in GUID_TO_NAME_MAP:
        return GUID_TO_NAME_MAP[target_guid]

    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, NIC_REG_KEY) as root_key:
            i = 0
            while True:
                try:
                    subkey_name = "%04d" % i
                    with winreg.OpenKey(root_key, subkey_name) as sub_key:
                        guid = winreg.QueryValueEx(sub_key, "NetCfgInstanceId")[0].upper().strip('{}')
                        
                        if guid == target_guid:
                            try:
                                driver_desc = winreg.QueryValueEx(sub_key, "DriverDesc")[0]
                                GUID_TO_NAME_MAP[target_guid] = driver_desc
                                return driver_desc
                            except:
                                GUID_TO_NAME_MAP[target_guid] = "未知适配器"
                                return "未知适配器"
                    i += 1
                except OSError: 
                    break
        
        GUID_TO_NAME_MAP[target_guid] = "未识别适配器"
        return "未识别适配器"
        
    except Exception:
        return "查询失败"


def get_network_interfaces_and_mac_threaded():
    """异步执行：获取网卡列表、MAC地址和IP地址。"""
    interfaces = {} 
    error_msg = None
    
    if platform.system() != "Windows" or netifaces is None:
        error_msg = "错误：系统不支持或 netifaces 未正确安装。"
        load_queue.put((interfaces, error_msg))
        return

    try:
        for iface_id in netifaces.interfaces():
            addrs = netifaces.ifaddresses(iface_id)
            
            mac_address = None
            ip_address = "N/A"
            
            if netifaces.AF_LINK in addrs:
                mac_address = addrs[netifaces.AF_LINK][0].get('addr', '').upper().replace('-', ':')
            
            if netifaces.AF_INET in addrs:
                ip_list = [a.get('addr') for a in addrs[netifaces.AF_INET] if a.get('addr') and a.get('addr') != '127.0.0.1']
                if ip_list:
                    ip_address = ip_list[0]
                    
            if mac_address and mac_address != '00:00:00:00:00:00' and mac_address != 'N/A':
                driver_desc = get_nic_driver_desc(iface_id)
                guid_prefix = iface_id.strip('{}')[:8].upper()
                display_name = f"{driver_desc} (ID: {guid_prefix}...)"
                
                interfaces[display_name] = {
                    'mac': mac_address,
                    'ip': ip_address
                }

        if not interfaces:
             error_msg = "未找到任何非回环网络接口。请检查驱动和连接。"
             
    except Exception as e:
        error_msg = f"获取接口信息失败 (底层调用): {e}"
    
    load_queue.put((interfaces, error_msg))


def change_mac_address(interface_display_name, new_mac):
    """执行 MAC 地址修改操作（操作注册表并重启网卡）。"""
    if len(new_mac) != 17 or new_mac.count(':') != 5:
        return False, "MAC地址格式不正确 (例如: 00:11:22:33:44:55)", "N/A"

    original_mac = MacChangerApp.instance.original_mac_label.cget("text")
    
    prefix_match = re.search(r'ID: ([0-9A-F]{8})', interface_display_name, re.IGNORECASE)
    if not prefix_match:
        return False, "无法从接口名称中解析出 GUID。", original_mac
    guid_prefix = prefix_match.group(1).upper()
    
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
                            try:
                                driver_desc_for_powershell = winreg.QueryValueEx(sub_key, "DriverDesc")[0]
                            except:
                                driver_desc_for_powershell = interface_display_name.split(' (ID:')[0].strip() 
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
    
    full_key = os.path.join(NIC_REG_KEY, reg_key_path)
    safe_driver_desc = driver_desc_for_powershell.replace('"', '`"') 
    
    try:
        # 3. 修改注册表
        mac_no_separator = new_mac.replace(':', '').replace('-', '').upper()
        
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, full_key, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, 'NetworkAddress', 0, winreg.REG_SZ, mac_no_separator)
            
        # 4. 禁用/启用网卡使修改生效 (使用 PowerShell)
        disable_cmd = f'powershell -command "Disable-NetAdapter -InterfaceDescription \\"{safe_driver_desc}\\" -Confirm:$false"'
        subprocess.check_output(disable_cmd, shell=True, stderr=subprocess.STDOUT)
        
        time.sleep(1.5) 
        
        enable_cmd = f'powershell -command "Enable-NetAdapter -InterfaceDescription \\"{safe_driver_desc}\\" -Confirm:$false"'
        subprocess.check_output(enable_cmd, shell=True, stderr=subprocess.STDOUT)
        
        log_message = f"Windows: 成功修改 {driver_desc_for_powershell} 的MAC地址并使用 PowerShell 重启网卡。"
        return True, log_message, original_mac

    except subprocess.CalledProcessError as e:
        error_output = e.output.decode('gbk', errors='ignore').strip()
        error_msg = f"网卡重启失败 (PowerShell)。\n请手动禁用/启用网卡！\n接口描述: {driver_desc_for_powershell}\n错误: {error_output}"
        return False, error_msg, original_mac
    except PermissionError:
        error_msg = "权限不足！请以 **管理员身份** 运行本程序。"
        return False, error_msg, original_mac
    except Exception as e:
        error_msg = f"修改或注册表操作失败: {e}"
        return False, error_msg, original_mac


# --- 日志记录功能 ---
def save_log(log_entry):
    """将日志条目写入文件"""
    log_file = "mac_change_log.txt"
    try:
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(log_entry + "\n")
    except Exception:
        pass


# --- GUI 界面和事件处理 ---

class MacChangerApp:
    instance = None 
    
    def center_window(self, width, height):
        """计算并设置窗口在屏幕中央的位置"""
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()

        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)

        self.master.geometry(f'{width}x{height}+{x}+{y}')

    def log_message(self, message):
        """在GUI日志区显示信息"""
        # 确保 log_text 存在 (已在 __init__ 中提前创建)
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')


    def __init__(self, master):
        MacChangerApp.instance = self
        self.master = master
        
        WINDOW_WIDTH = 650
        WINDOW_HEIGHT = 520
        
        master.title("MAC 地址修改工具")
        master.resizable(False, False)

        master.update_idletasks()
        self.center_window(WINDOW_WIDTH, WINDOW_HEIGHT)
        
        self.current_interfaces = {} 
        self.interface_load_thread = None

        style = ttk.Style()
        style.configure('TFrame', background='#f0f0f0')
        style.configure('TLabel', background='#f0f0f0', font=('Arial', 10))
        style.configure('TButton', font=('Arial', 10, 'bold'))

        main_frame = ttk.Frame(master, padding="10")
        main_frame.pack(fill='both', expand=True)

        # ----------------------------------------------------------------------
        # !!! 关键修正区域：将日志文本框和相关标签提前创建 !!!
        # 确保 log_text 在任何日志调用前存在
        # ----------------------------------------------------------------------
        # 状态和日志标签 (Row 5 - Row 8 对应的区域)
        self.status_label = ttk.Label(main_frame, text="", foreground='black') 
        self.status_label.grid(row=5, column=0, columnspan=2, pady=5)
        
        ttk.Label(main_frame, text="Windows：请务必以 **管理员身份** 运行本程序！", 
                  foreground='red', font=('Arial', 9, 'bold')).grid(row=6, column=0, columnspan=2, pady=(0, 5))
        
        ttk.Label(main_frame, text="--- 操作日志 ---").grid(row=7, column=0, columnspan=2, pady=(10, 5))
        self.log_text = tk.Text(main_frame, height=8, width=70, state='disabled', wrap='word', font=('Courier', 9))
        self.log_text.grid(row=8, column=0, columnspan=2, sticky='nsew', padx=5)
        
        self.log_message("[信息] 程序启动，正在初始化界面...")
        
        # --- 设置图标 ---
        try:
            icon_path = resource_path('mac.ico')
            if os.path.exists(icon_path):
                self.master.iconbitmap(icon_path)
                self.log_message("[信息] 窗口图标尝试加载...")
            else:
                 self.log_message(f"[警告] 未找到图标文件: {icon_path}，请确保 mac.ico 已正确打包。")
        except Exception as e:
            self.log_message(f"[警告] 图标设置失败: {e}")
            pass
        # ----------------------------------------------------------------------
        
        # 1. 网络接口选择 (Row 0)
        ttk.Label(main_frame, text="选择网络接口:").grid(row=0, column=0, sticky='w', pady=5, padx=5)
        self.interface_var = tk.StringVar()
        self.interface_combobox = ttk.Combobox(main_frame, textvariable=self.interface_var, state='readonly', width=40)
        self.interface_combobox.grid(row=0, column=1, sticky='ew', pady=5, padx=5)
        self.interface_combobox.bind('<<ComboboxSelected>>', self.on_interface_select)
        
        # 2. 原MAC地址显示 (Row 1)
        ttk.Label(main_frame, text="原 MAC 地址:").grid(row=1, column=0, sticky='w', pady=5, padx=5)
        self.original_mac_label = ttk.Label(main_frame, text="N/A", foreground='blue', width=40, cursor="hand2") 
        self.original_mac_label.grid(row=1, column=1, sticky='w', pady=5, padx=5)
        self.original_mac_label.bind("<Double-Button-1>", self.copy_original_mac) 

        # 3. IP 地址显示 (Row 2)
        ttk.Label(main_frame, text="IP 地址:").grid(row=2, column=0, sticky='w', pady=5, padx=5)
        self.ip_address_label = ttk.Label(main_frame, text="N/A", foreground='black', width=40, cursor="hand2") 
        self.ip_address_label.grid(row=2, column=1, sticky='w', pady=5, padx=5)
        self.ip_address_label.bind("<Double-Button-1>", self.copy_ip_address) 

        # 4. 新 MAC 地址输入 & 随机按钮 (Row 3)
        ttk.Label(main_frame, text="新的 MAC 地址:").grid(row=3, column=0, sticky='w', pady=5, padx=5)
        
        mac_input_frame = ttk.Frame(main_frame)
        mac_input_frame.grid(row=3, column=1, sticky='ew', pady=5, padx=5)
        mac_input_frame.columnconfigure(0, weight=1) 
        
        self.mac_entry = ttk.Entry(mac_input_frame, width=35)
        self.mac_entry.grid(row=0, column=0, sticky='ew', padx=(0, 5))
        self.mac_entry.insert(0, generate_random_mac()) 
        
        self.random_button = ttk.Button(mac_input_frame, text="随机生成", command=self.on_generate_random_mac)
        self.random_button.grid(row=0, column=1, sticky='e')
        
        # 5. 执行按钮 (Row 4)
        self.change_button = ttk.Button(main_frame, text="执行修改", command=self.on_change_click, state='disabled')
        self.change_button.grid(row=4, column=0, columnspan=2, pady=15)
        
        
        # 7. 底部联系方式 (Row 9)
        contact_frame = ttk.Frame(main_frame, padding="5")
        contact_frame.grid(row=9, column=0, columnspan=2, sticky='ew', pady=(10, 0))
        
        ttk.Label(contact_frame, text="联系方式：QQ 88179096", font=('Arial', 9)).pack(side='left', padx=5)
        
        self.link_label = ttk.Label(contact_frame, text="www.itvip.com.cn", 
                                    foreground="blue", cursor="hand2", font=('Arial', 9, 'underline'))
        self.link_label.pack(side='left', padx=15)
        self.link_label.bind("<Button-1>", lambda e: self.open_link("http://www.itvip.com.cn"))

        contact_frame.columnconfigure(0, weight=1) 

        main_frame.grid_columnconfigure(1, weight=1)
        main_frame.grid_rowconfigure(8, weight=1)

        # 启动异步加载 (状态标签已创建)
        self.start_load_interfaces()
        self.master.after(100, self.check_load_interfaces)


    def open_link(self, url):
        webbrowser.open_new_tab(url)

    def copy_original_mac(self, event):
        mac_address = self.original_mac_label.cget("text")
        
        if mac_address != "N/A":
            self.master.clipboard_clear() 
            self.master.clipboard_append(mac_address) 
            self.log_message(f"[信息] 已将 MAC 地址 {mac_address} 复制到剪贴板。")
        else:
            self.log_message("[警告] 剪贴板复制失败：当前没有可用的 MAC 地址。")

    def copy_ip_address(self, event):
        ip_address = self.ip_address_label.cget("text")
        
        if ip_address != "N/A":
            self.master.clipboard_clear() 
            self.master.clipboard_append(ip_address) 
            self.log_message(f"[信息] 已将 IP 地址 {ip_address} 复制到剪贴板。")
        else:
            self.log_message("[警告] 剪贴板复制失败：当前没有可用的 IP 地址。")

    def on_generate_random_mac(self):
        new_mac = generate_random_mac()
        self.mac_entry.delete(0, tk.END)
        self.mac_entry.insert(0, new_mac)
        self.log_message(f"[信息] 已生成新的随机 MAC 地址：{new_mac}")

    def start_load_interfaces(self):
        self.interface_load_thread = threading.Thread(target=get_network_interfaces_and_mac_threaded, daemon=True)
        self.interface_load_thread.start()
        self.log_message("[信息] 已启动后台线程加载网卡信息...")
        self.change_button.config(state='disabled')


    def check_load_interfaces(self):
        try:
            interfaces, error_msg = load_queue.get_nowait()
            self.finalize_load_interfaces(interfaces, error_msg)
            
        except queue.Empty:
            self.master.after(100, self.check_load_interfaces)


    def finalize_load_interfaces(self, interfaces, error_msg):
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
            self.ip_address_label.config(text="N/A")
            self.status_label.config(text="未找到可用网卡，请检查权限/系统环境。", foreground='red')
            self.log_message("[错误] 未找到任何可用网卡。")

    def on_interface_select(self, event):
        selected_interface = self.interface_var.get()
        data = self.current_interfaces.get(selected_interface, {'mac': 'N/A', 'ip': 'N/A'})
        
        mac = data['mac']
        ip = data['ip']
        
        self.original_mac_label.config(text=mac)
        self.ip_address_label.config(text=ip)
        
        if mac != "N/A":
            self.log_message(f"[信息] 接口 {selected_interface} 原始MAC地址为 {mac}, IP地址为 {ip}")
        else:
            self.log_message(f"[错误] 无法获取接口 {selected_interface} 的MAC地址和IP地址。")

    def on_change_click(self):
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
            
            self.start_load_interfaces()
            self.master.after(100, self.check_load_interfaces)
            
        else:
            log_entry = f"[{timestamp}] [失败] {interface} | 原MAC: {original_mac_logged} | 新MAC: {new_mac} | 错误: {message}"
            self.status_label.config(text="修改失败，请查看错误信息。", foreground='red')
            messagebox.showerror("失败", f"MAC地址修改失败:\n{message}")

        self.log_message(log_entry)
        save_log(log_entry)


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