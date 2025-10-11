import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog, ttk
import subprocess
import sys
import ctypes
import os
import threading

# --- 辅助函数：权限检查 ---

def is_admin():
    """检查程序是否以管理员权限运行。"""
    try:
        # 适用于 Windows 的权限检查
        return sys.getwindowsversion().platform == 2 and ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False # 假设非 Windows 或检查失败

# --- 后端：使用 PowerShell/CMD 执行系统操作 ---

def run_powershell_command(command, error_message="命令执行失败"):
    """
    运行一个 PowerShell 命令并捕获输出。
    """
    try:
        # 使用 -ExecutionPolicy Bypass 允许运行脚本
        # 命令之间使用分号 ; 确保顺序执行
        result = subprocess.run(
            ['powershell', '-Command', command],
            capture_output=True,
            text=True,
            check=True,
            encoding='gbk' # 尝试使用GBK/cp936编码以正确处理中文输出
        )
        return True, result.stdout
    except subprocess.CalledProcessError as e:
        # 命令执行失败（非零退出码）
        error_output = e.stderr or e.stdout or "无错误输出"
        return False, f"{error_message}。错误码: {e.returncode}\n错误输出:\n{error_output}"
    except FileNotFoundError:
        return False, "找不到 'powershell' 命令。请确保您的系统环境配置正确。"
    except Exception as e:
        return False, f"发生未知错误: {e}"

# --- 兼容性修复功能 ---

def toggle_password_protected_sharing(enable):
    """启用或禁用密码保护共享。"""
    if enable:
        # 启用密码保护共享 (0 = Use user accounts)
        command = 'Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Lsa" -Name "everyoneincludesguest" -Value 0 -Type DWord;'
        title = "启用密码保护共享"
    else:
        # 禁用密码保护共享 (1 = Guest access)
        command = 'Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Lsa" -Name "everyoneincludesguest" -Value 1 -Type DWord;'
        title = "禁用密码保护共享"
        
    success, output = run_powershell_command(command, "修改密码保护设置失败")
    messagebox.showinfo(title, "成功执行命令，可能需要重启才能完全生效。" if success else f"操作失败:\n{output}")

def enable_insecure_guest_logons():
    """针对 Windows 10/11，启用不安全的访客登录。"""
    command = 'Set-SmbClientConfiguration -EnableInsecureGuestLogons $true -Force'
    title = "启用不安全访客登录 (高兼容性/低安全性)"
    
    response = messagebox.askyesno(title, "警告：启用此设置会降低安全性，但能解决 Win10/11 无法访问无密码共享的问题。是否继续？")
    if not response:
        return
        
    success, output = run_powershell_command(command, "启用不安全访客登录失败")
    messagebox.showinfo(title, "成功启用。请注意安全风险。" if success else f"操作失败:\n{output}")

def check_network_discovery():
    """检查网络发现状态。"""
    command_check = 'Get-NetFirewallRule -DisplayGroup "网络发现" | Select-Object DisplayName, Enabled, Action | Format-List'
    
    success, output = run_powershell_command(command_check, "网络发现检查失败")
    if success:
        messagebox.showinfo("网络发现状态", f"防火墙状态（网络发现）：\n{output}\n\n请确保相关规则已启用（Enabled: True）。")
    else:
        messagebox.showerror("网络发现检查失败", output)
        
def fix_network_discovery():
    """
    一键启用网络发现的防火墙规则和相关服务，并设置网络为专用。
    """
    command_fix = (
        # 启用网络发现防火墙规则
        'Set-NetFirewallRule -DisplayGroup "网络发现" -Enabled True -ErrorAction SilentlyContinue;' 
        # 设置网络为专用
        'Set-NetConnectionProfile -InterfaceIndex (Get-NetAdapter | Select-Object -First 1).InterfaceIndex -NetworkCategory Private -ErrorAction SilentlyContinue;' 
        
        # 核心服务：确保它们在运行
        'Set-Service LanmanServer -StartupType Automatic -Status Running -ErrorAction SilentlyContinue;' # Server 服务 (核心文件共享)
        'Set-Service Browser -StartupType Automatic -Status Running -ErrorAction SilentlyContinue;' # Computer Browser 服务 (辅助网络发现)
        
        # 依赖服务：确保它们在运行
        'Set-Service FDPHost -StartupType Automatic -Status Running -ErrorAction SilentlyContinue;' 
        'Set-Service FDPublish -StartupType Automatic -Status Running -ErrorAction SilentlyContinue;' 
        'Set-Service SSDPSRV -StartupType Automatic -Status Running -ErrorAction SilentlyContinue;' 
    )
                  
    success, output = run_powershell_command(command_fix, "修复网络发现失败")
    messagebox.showinfo("修复网络发现", "网络发现、所有相关共享服务和网络配置已尝试修复。\n" if success else f"修复失败:\n{output}")

def force_enable_file_sharing_firewall():
    """
    【修复参数冲突】
    强制启用文件和打印机共享所需的防火墙规则（端口 445/139）。
    使用 Get-NetFirewallRule -DisplayGroup 然后通过 Where-Object 过滤，确保兼容性。
    """
    
    group_name = "文件和打印机共享"
    
    # 查找 DisplayGroup 包含 "文件和打印机共享" 的所有规则，
    # 并使用 Where-Object 过滤出 SMB-In 和 NB (NetBIOS) 相关的规则，然后启用。
    command = f"""
        Get-NetFirewallRule -DisplayGroup "{group_name}" | 
        Where-Object {{ $_.DisplayName -like '*SMB-In*' -or $_.DisplayName -like '*NB-*' }} | 
        Set-NetFirewallRule -Enabled True -Action Allow -Profile Private, Domain -ErrorAction SilentlyContinue;
        
        # 确保整个组被启用，以防遗漏
        Set-NetFirewallRule -DisplayGroup "{group_name}" -Enabled True -Profile Any -ErrorAction SilentlyContinue;
    """
    
    success, output = run_powershell_command(command, "强制启用文件共享防火墙失败")
    
    messagebox.showinfo("强制启用文件共享", 
                        "已成功尝试强制启用所有‘文件和打印机共享’防火墙规则。请检查您的网络配置文件是否为『专用』。" if success else f"操作失败:\n{output}")

def check_and_start_core_discovery_services():
    """
    【修复可见性】
    检查并强制启动所有核心网络发现服务，解决文件夹共享不可见问题。
    """
    
    # 这几个服务是网络发现和 Master Browser 选举的依赖核心
    services = [
        "FDPHost",        # Function Discovery Provider Host
        "FDResPub",       # Function Discovery Resource Publication
        "SSDPDSrv",       # SSDP Discovery Service (SSDP Discovery)
        "upnphost",       # UPnP Device Host
        "Browser",        # Computer Browser (Master Browser 依赖)
        "lmhosts"         # TCP/IP NetBIOS Helper (NetBIOS 辅助)
    ]
    
    # 构建命令：设置启动类型为自动并立即启动
    commands = []
    for service in services:
        commands.append(f'Set-Service -Name "{service}" -StartupType Automatic -ErrorAction SilentlyContinue;')
        commands.append(f'Start-Service -Name "{service}" -ErrorAction SilentlyContinue;')

    full_command = " ".join(commands)
    
    success, output = run_powershell_command(full_command, "强制启动核心发现服务失败")
    
    messagebox.showinfo("核心发现服务状态", 
                        "所有核心网络发现服务已强制设置为『自动』并尝试启动。请【重启电脑】后测试共享可见性。" if success else f"操作失败:\n{output}")

def force_enable_netbios():
    """
    【最终修复 NetBIOS 配置】
    强制启用所有网络适配器上的 TCP/IP 上的 NetBIOS (设置为 Default)。
    """
    
    # 查找所有活动的接口
    command_get_interfaces = 'Get-NetAdapter | Where-Object {$_.Status -eq "Up"} | Select-Object -ExpandProperty InterfaceIndex'
    
    success, interface_indices_raw = run_powershell_command(command_get_interfaces, "获取网卡索引失败")
    
    if not success:
        messagebox.showerror("NetBIOS 修复失败", interface_indices_raw)
        return

    indices = [idx.strip() for idx in interface_indices_raw.split() if idx.strip()]
    
    if not indices:
        messagebox.showinfo("NetBIOS 修复", "未找到活动的网络适配器。")
        return

    commands = []
    for index in indices:
        # 1. 确保 NetBIOS 绑定已启用 (ms_tcpip)
        commands.append(f'Set-NetAdapterBinding -InterfaceIndex {index} -ComponentID ms_tcpip -Enabled $True -ErrorAction SilentlyContinue;')
        
        # 2. 确保文件和打印机共享绑定已启用
        commands.append(f'Set-NetAdapterBinding -InterfaceIndex {index} -ComponentID ms_server -Enabled $True -ErrorAction SilentlyContinue;')

    # 3. 强制设置 NetBIOS over TCP/IP 选项为 Default (1) (注册表级别设置)
    # RegistryValue 1 对应 "Default" (遵循 DHCP/WINS 设置，如果未配置则自动使用 NetBIOS 广播)
    netbios_command = 'Get-NetAdapter | Where-Object {$_.Status -eq "Up"} | ForEach-Object { Set-NetAdapterAdvancedProperty -Name $_.Name -RegistryKeyword "DhcpNetBios" -RegistryValue 1 -ErrorAction SilentlyContinue; }'
    commands.append(netbios_command)
    
    full_command = " ".join(commands)

    success, output = run_powershell_command(full_command, "强制启用 NetBIOS 失败")
    
    if success:
        messagebox.showinfo("NetBIOS 强制启用", "已强制启用所有网卡上的 NetBIOS 辅助和文件共享绑定。\n\n【必须重启电脑】才能生效。")
    else:
        messagebox.showerror("NetBIOS 修复失败", output)

def verify_settings_status():
    """验证网络配置文件、关键服务和防火墙状态是否已正确设置。"""
    
    # 1. 检查网络配置文件 (Network Profile)
    # 使用 Get-NetConnectionProfile 检查是否为 Private 或 Domain
    command_profile = '(Get-NetConnectionProfile | Select-Object -First 1).NetworkCategory.ToString()'
    success_profile, output_profile = run_powershell_command(command_profile, "检查网络配置文件失败")
    profile_status = f"当前网络配置文件: {output_profile.strip()}" if success_profile else "网络配置文件状态: 检查失败"

    # 2. 检查核心服务状态 (FDPHost, lmhosts)
    # 检查 Function Discovery Provider Host 和 TCP/IP NetBIOS Helper
    command_services = """
    $fdp = Get-Service -Name "FDPHost" -ErrorAction SilentlyContinue;
    $lmhosts = Get-Service -Name "lmhosts" -ErrorAction SilentlyContinue;
    
    $fdpStatus = if ($fdp) { "FDPHost: " + $fdp.Status.ToString() + " (" + $fdp.StartType.ToString() + ")" } else { "FDPHost: 未找到" };
    $lmhStatus = if ($lmhosts) { "NetBIOS Helper: " + $lmhosts.Status.ToString() + " (" + $lmhosts.StartType.ToString() + ")" } else { "NetBIOS Helper: 未找到" };
    
    $fdpStatus + "`n" + $lmhStatus
    """
    success_services, output_services = run_powershell_command(command_services, "检查核心服务状态失败")
    
    service_status = "核心服务状态:\n"
    if success_services:
        service_status += output_services.strip()
    else:
        service_status += f"检查失败: {output_services}"

    # 3. 检查文件共享防火墙规则
    # 查找是否有 "文件和打印机共享" 组内启用的规则
    command_firewall = """
    $rules = Get-NetFirewallRule -DisplayGroup "文件和打印机共享" | Where-Object {$_.Enabled -eq $True -and ($_.Profile -like '*Private*' -or $_.Profile -like '*Any*')};
    if ($rules.Count -gt 0) { "文件共享防火墙规则: 找到 $($rules.Count) 条已启用规则。" } else { "文件共享防火墙规则: 未找到已启用规则。" }
    """
    success_firewall, output_firewall = run_powershell_command(command_firewall, "检查防火墙规则失败")
    firewall_status = output_firewall.strip() if success_firewall else f"防火墙状态: 检查失败 ({output_firewall})"
    
    # 组合结果并显示
    final_message = f"--- 关键设置验证结果 ---\n\n{profile_status}\n\n{service_status}\n\n{firewall_status}\n\n"
    final_message += "【FDPHost】和【NetBIOS Helper】应显示 Running (Automatic)。\n【配置文件】应为 Private 或 Domain。"
    messagebox.showinfo("系统状态验证", final_message)


# --- 共享驱动器/文件夹及设置权限 (已切换为 New-SmbShare) ---

def create_shared_folder(path, share_name, permission):
    """
    使用 New-SmbShare 创建文件夹共享并设置权限，以确保与 Windows 的最大兼容性。
    """
    
    # 1. 确定权限参数
    full_access = ''
    read_access = ''
    
    if permission == 'Full':
        # 赋予 'Everyone' 完全控制权限
        full_access = "Everyone"
        read_access = ''
        share_perm_msg = '完全控制'
    else: # Read
        # 赋予 'Everyone' 只读权限
        full_access = ''
        read_access = "Everyone"
        share_perm_msg = '只读'
        
    # 路径规范化
    normalized_path = os.path.normpath(path)
    
    # 构建 New-SmbShare 命令
    access_param = ""
    if full_access:
        access_param += f"-FullAccess \"{full_access}\""
    elif read_access:
        access_param += f"-ReadAccess \"{read_access}\""
        
    command = f"""
        # 1. 确保文件夹存在
        if (-not (Test-Path -Path '{normalized_path}')) {{ New-Item -Path '{normalized_path}' -ItemType Directory | Out-Null }};
        
        # 2. 尝试移除旧的共享（使用 Remove-SmbShare）
        Remove-SmbShare -Name "{share_name}" -Force -ErrorAction SilentlyContinue;

        # 3. 使用 New-SmbShare 创建新的共享并设置权限
        New-SmbShare -Name "{share_name}" -Path "{normalized_path}" -Description "Shared by SMB Fixer Tool" {access_param} -Force;
                     
        # 4. 强制设置 NTFS 权限：确保核心系统用户拥有访问权限
        cmd /c icacls "{normalized_path}" /inheritance:r /T 2>$null;
        cmd /c icacls "{normalized_path}" /grant SYSTEM:F /T 2>$null;
        cmd /c icacls "{normalized_path}" /grant Administrators:F /T 2>$null;
    """
    
    success, output = run_powershell_command(command, f"创建共享 {share_name} 失败")
    
    if not success:
         output += f"\n\n尝试执行的命令:\n{command}" 
         messagebox.showerror("创建驱动器/文件夹共享失败", output)
    else:
        messagebox.showinfo("创建驱动器/文件夹共享成功", 
                            f"共享 '{share_name}' 已使用 PowerShell 核心命令创建，权限为 '{share_perm_msg}'。\n\n"
                            "请重启电脑后测试访问，或确保已运行兼容性修复中的所有『修复』按钮。")


def show_create_share_dialog():
    """显示创建共享的对话框"""
    
    path = filedialog.askdirectory(title="选择或创建要共享的文件夹路径")
    if not path:
        return

    share_name = simpledialog.askstring("共享名称", "请输入共享名 (例如: MyDataShare)", parent=app)
    if not share_name or not share_name.strip():
        return
    share_name = share_name.strip()
        
    permission_choice = simpledialog.askstring("权限选择", "请输入权限级别 (Read 或 Full):", parent=app)
    if permission_choice not in ['Read', 'Full']:
        messagebox.showerror("错误", "权限级别必须是 'Read' 或 'Full'。")
        return

    # 在单独的线程中运行耗时操作，避免 GUI 冻结
    threading.Thread(target=lambda: create_shared_folder(path, share_name, permission_choice)).start()


# --- 共享打印机功能 ---

def get_printers():
    """获取本地所有打印机列表"""
    command = "Get-Printer | Select-Object Name | ForEach-Object { $_.Name }"
    success, output = run_powershell_command(command, "获取打印机列表失败")
    if success:
        return [p.strip() for p in output.split('\n') if p.strip()]
    return []

def share_printer(printer_name, share_name):
    """
    共享指定的打印机。
    """
    command_share = f"Set-Printer -Name '{printer_name}' -Shared $True -ShareName '{share_name}' -ErrorAction Stop"
    
    success, output = run_powershell_command(command_share, f"共享打印机 {printer_name} 失败")
    
    if success:
        messagebox.showinfo("共享打印机", f"打印机 '{printer_name}' 已共享为 '{share_name}'。")
    else:
        messagebox.showerror("共享打印机失败", output)

class PrinterShareDialog(tk.Toplevel):
    """共享打印机对话框"""
    def __init__(self, master):
        super().__init__(master)
        self.title("共享打印机")
        self.transient(master)
        self.grab_set() 
        
        tk.Label(self, text="选择要共享的本地打印机:").pack(pady=5)
        
        self.printer_listbox = tk.Listbox(self, selectmode=tk.SINGLE, height=5)
        self.printer_listbox.pack(padx=10, fill="x")
        
        printers = get_printers()
        if not printers:
            self.printer_listbox.insert(tk.END, "未找到本地打印机。")
        else:
            for p in printers:
                self.printer_listbox.insert(tk.END, p)
        
        tk.Label(self, text="共享名称:").pack(pady=5)
        self.share_name_entry = tk.Entry(self)
        self.share_name_entry.pack(padx=10, fill="x")
        
        tk.Button(self, text="共享打印机", command=self.do_share).pack(pady=10)
        self.protocol("WM_DELETE_WINDOW", self.master.focus_set)
        
    def do_share(self):
        selected_index = self.printer_listbox.curselection()
        share_name = self.share_name_entry.get().strip()
        
        if not selected_index or not share_name:
            messagebox.showerror("错误", "请选择打印机并输入共享名称。")
            return
            
        printer_name = self.printer_listbox.get(selected_index[0])
        
        # 避免尝试共享虚拟打印机（如 PDF），但代码中不强制限制，只处理用户选择的打印机
        threading.Thread(target=lambda: share_printer(printer_name, share_name)).start()
        self.destroy()

# --- GUI 界面 (主窗口) ---

class SMBSimpleFixer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SMB 共享兼容性与管理工具")
        self.geometry("600x750") # 调整高度以容纳新按钮

        self.check_admin_status()
        self.create_widgets()

    def check_admin_status(self):
        """检查并显示管理员权限状态"""
        if not is_admin():
            tk.Label(self, text="警告: 未以管理员身份运行!", fg="red", font=('Arial', 12, 'bold')).pack(pady=5)
            tk.Label(self, text="请右键以 '管理员身份运行'，否则无法修改系统设置。", fg="red").pack(pady=5)
        else:
             tk.Label(self, text="已检测到管理员权限，可以进行修复操作。", fg="green").pack(pady=5)

    def create_widgets(self):
        # 创建 Notebook (Tabbed Interface)
        notebook = ttk.Notebook(self)
        notebook.pack(pady=10, padx=10, expand=True, fill="both")
        
        # --- Tab 1: 兼容性修复 ---
        tab_repair = tk.Frame(notebook, padx=10, pady=10)
        notebook.add(tab_repair, text='兼容性修复 (网络/防火墙)')
        self.create_repair_tab(tab_repair)
        
        # --- Tab 2: 共享管理 ---
        tab_manage = tk.Frame(notebook, padx=10, pady=10)
        notebook.add(tab_manage, text='共享管理 (创建/权限)')
        self.create_manage_tab(tab_manage)


    def create_repair_tab(self, tab):
        
        # --- 共享权限设置区 ---
        frame_sharing = tk.LabelFrame(tab, text="共享访问和密码问题 (Win7/10/11 兼容)", padx=10, pady=10)
        frame_sharing.pack(padx=10, pady=10, fill="x")

        tk.Button(frame_sharing, text="禁用密码保护共享 (实现无密码访问)", command=lambda: threading.Thread(target=lambda: toggle_password_protected_sharing(False)).start()).pack(fill="x", pady=5)
        tk.Button(frame_sharing, text="启用密码保护共享 (恢复默认)", command=lambda: threading.Thread(target=lambda: toggle_password_protected_sharing(True)).start()).pack(fill="x", pady=5)
        
        tk.Button(frame_sharing, text="[Win10/11] 启用不安全访客登录 (解决访问被拒)", command=lambda: threading.Thread(target=enable_insecure_guest_logons).start()).pack(fill="x", pady=5)

        # --- 网络发现区 ---
        frame_discovery = tk.LabelFrame(tab, text="网络发现/连接故障", padx=10, pady=10)
        frame_discovery.pack(padx=10, pady=10, fill="x")
        
        tk.Button(frame_discovery, text="检查网络发现状态", command=lambda: threading.Thread(target=check_network_discovery).start()).pack(side=tk.LEFT, expand=True, padx=5)
        tk.Button(frame_discovery, text="一键修复网络发现", command=lambda: threading.Thread(target=fix_network_discovery).start()).pack(side=tk.RIGHT, expand=True, padx=5)
        tk.Label(frame_discovery, text="修复：防火墙、网络配置文件(专用)和所有共享服务。", fg="blue").pack(pady=5)
        
        # --- 最终强制修复区 (核心可见性) ---
        frame_force = tk.LabelFrame(tab, text="SMB 核心可见性与连接强制修复", padx=10, pady=10)
        frame_force.pack(padx=10, pady=10, fill="x")

        tk.Button(frame_force, text="强制启用文件共享防火墙规则 (解决连接问题)", command=lambda: threading.Thread(target=force_enable_file_sharing_firewall).start()).pack(fill="x", pady=5)
        
        tk.Button(frame_force, text="强制启动网络发现核心服务 (解决文件夹不可见)", 
                  command=lambda: threading.Thread(target=check_and_start_core_discovery_services).start()).pack(fill="x", pady=5)

        tk.Button(frame_force, text="强制启用网卡 NetBIOS 设置 (Master Browser 依赖)", 
                  command=lambda: threading.Thread(target=force_enable_netbios).start()).pack(fill="x", pady=5)
        
        tk.Label(frame_force, text="强烈建议：运行上述所有三个命令后，【重启电脑】。", fg="red").pack(pady=5)

        # --- 新增的验证区 ---
        frame_verify = tk.LabelFrame(tab, text="设置生效验证 (故障排除)", padx=10, pady=10)
        frame_verify.pack(padx=10, pady=10, fill="x")

        tk.Button(frame_verify, text="验证关键设置是否已生效 (点击查看状态)", 
                  command=lambda: threading.Thread(target=verify_settings_status).start()).pack(fill="x", pady=5)
        
        tk.Label(frame_verify, text="请先运行上方的所有修复按钮，然后点击此按钮查看状态。", fg="blue").pack(pady=5)


    def create_manage_tab(self, tab):
        
        # --- 驱动器/文件夹共享区 ---
        frame_drive = tk.LabelFrame(tab, text="驱动器/文件夹共享 (New-SmbShare)", padx=10, pady=10)
        frame_drive.pack(padx=10, pady=10, fill="x")

        tk.Button(frame_drive, text="创建新的共享并设置权限 (使用 New-SmbShare)", command=show_create_share_dialog).pack(fill="x", pady=5)
        tk.Label(frame_drive, text="使用 PowerShell 推荐的 New-SmbShare，最大化兼容性和可见性。").pack(pady=5)

        # --- 打印机共享区 ---
        frame_printer = tk.LabelFrame(tab, text="打印机共享 (Set-Printer)", padx=10, pady=10)
        frame_printer.pack(padx=10, pady=10, fill="x")
        
        tk.Button(frame_printer, text="共享本地打印机", command=lambda: PrinterShareDialog(self)).pack(fill="x", pady=5)
        
        # --- 提示 ---
        tk.Label(tab, text="* 访问路径示例: \\\\您的电脑名\\共享名 或 \\\\IP地址\\共享名", fg="blue").pack(pady=10)


if __name__ == "__main__":
    app = SMBSimpleFixer()
    app.mainloop()