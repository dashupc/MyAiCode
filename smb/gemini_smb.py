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
        # CMD 命令（如 net share, icacls）可以直接在 PowerShell 中运行
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
        return False, f"{error_message}。错误码: {e.returncode}\n错误输出:\n{e.stderr}"
    except FileNotFoundError:
        return False, "找不到 'powershell' 命令。请确保您的系统环境配置正确。"
    except Exception as e:
        return False, f"发生未知错误: {e}"

# --- 兼容性修复功能 ---

def toggle_password_protected_sharing(enable):
    """启用或禁用密码保护共享。"""
    if enable:
        # 启用密码保护共享 (0 = Use user accounts)
        command = f'Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Lsa" -Name "everyoneincludesguest" -Value 0 -Type DWord;'
        title = "启用密码保护共享"
    else:
        # 禁用密码保护共享 (1 = Guest access)
        command = f'Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Lsa" -Name "everyoneincludesguest" -Value 1 -Type DWord;'
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
    """一键启用网络发现的防火墙规则和相关服务。"""
    command_fix = 'Set-NetFirewallRule -DisplayGroup "网络发现" -Enabled True; ' \
                  'Set-NetConnectionProfile -InterfaceIndex (Get-NetAdapter | Select-Object -First 1).InterfaceIndex -NetworkCategory Private; ' \
                  'Set-Service FDPHost -StartupType Automatic -Status Running; ' \
                  'Set-Service FDPublish -StartupType Automatic -Status Running; '
                  
    success, output = run_powershell_command(command_fix, "修复网络发现失败")
    messagebox.showinfo("修复网络发现", "网络发现和相关服务已尝试启用和设置网络为私有。\n" if success else f"修复失败:\n{output}")

# --- 共享驱动器/文件夹及设置权限 (使用 net share + icacls 修复兼容性问题) ---

def create_shared_folder(path, share_name, permission):
    """
    创建文件夹共享，并设置Share和NTFS权限（使用兼容性更好的 net share 和 icacls）。
    """
    
    # 1. 确定共享和NTFS权限字符串
    if permission == 'Full':
        # net share: /grant:Everyone,FULL
        # icacls: /grant Everyone:F (完全控制)
        share_perm = 'FULL'
        ntfs_perm = 'F'
    else: # Read
        # net share: /grant:Everyone,READ
        # icacls: /grant Everyone:(OI)(CI)R (只读，继承容器和对象)
        share_perm = 'READ'
        ntfs_perm = '(OI)(CI)R'

    # 将 Python 路径 (/) 规范化为 Windows 路径 (\)
    normalized_path = os.path.normpath(path)
    
    # 使用 PowerShell 运行 CMD 命令
    command = f"""
        # 1. 移除旧的共享（如果存在）
        cmd /c net share {share_name} /delete /y 2>$null
        
        # 2. 创建文件夹 (如果不存在)
        if (-not (Test-Path -Path '{normalized_path}')) {{ New-Item -Path '{normalized_path}' -ItemType Directory | Out-Null }}
        
        # 3. 创建新的共享并设置共享权限 (使用 net share)
        cmd /c net share "{share_name}"="{normalized_path}" /grant:Everyone,{share_perm} /remark:"Shared by SMB Fixer"
        
        # 4. 设置 NTFS 权限 (使用 icacls - 覆盖继承并设置权限)
        # /inheritance:r: 移除继承的 ACEs
        # /T: 递归应用
        cmd /c icacls "{normalized_path}" /inheritance:r /T
        cmd /c icacls "{normalized_path}" /grant Everyone:{ntfs_perm} /T
        cmd /c icacls "{normalized_path}" /grant SYSTEM:F /T
        cmd /c icacls "{normalized_path}" /grant Administrators:F /T
    """
    
    success, output = run_powershell_command(command, f"创建共享 {share_name} 失败")
    messagebox.showinfo("创建驱动器/文件夹共享", f"共享 '{share_name}' 已创建并设置权限为 '{permission}'。" if success else f"操作失败:\n{output}")

def show_create_share_dialog():
    """显示创建共享的对话框"""
    
    path = filedialog.askdirectory(title="选择要共享的文件夹路径")
    if not path:
        return

    share_name = simpledialog.askstring("共享名称", "请输入共享名 (例如: MyDataShare)", parent=app)
    if not share_name:
        return
        
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
    # 启用打印机共享，并使用默认的共享权限
    command_share = f"Set-Printer -Name '{printer_name}' -Shared $True -ShareName '{share_name}' -ErrorAction Stop"
    
    success, output = run_powershell_command(command_share, f"共享打印机 {printer_name} 失败")
    
    if success:
        messagebox.showinfo("共享打印机", f"打印机 '{printer_name}' 已共享为 '{share_name}'。您可能仍需在打印机属性中手动处理驱动程序问题。")
    else:
        messagebox.showerror("共享打印机失败", output)

class PrinterShareDialog(tk.Toplevel):
    """共享打印机对话框"""
    def __init__(self, master):
        super().__init__(master)
        self.title("共享打印机")
        self.transient(master)
        self.grab_set() # 模态对话框
        
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
        self.protocol("WM_DELETE_WINDOW", self.master.focus_set) # 关闭时恢复主窗口焦点
        
    def do_share(self):
        selected_index = self.printer_listbox.curselection()
        share_name = self.share_name_entry.get().strip()
        
        if not selected_index or not share_name:
            messagebox.showerror("错误", "请选择打印机并输入共享名称。")
            return
            
        printer_name = self.printer_listbox.get(selected_index[0])
        
        # 在单独的线程中运行耗时操作
        threading.Thread(target=lambda: share_printer(printer_name, share_name)).start()
        self.destroy()

# --- GUI 界面 (主窗口) ---

class SMBSimpleFixer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SMB 共享兼容性与管理工具")
        self.geometry("600x500")

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
        notebook.add(tab_repair, text='兼容性修复')
        self.create_repair_tab(tab_repair)
        
        # --- Tab 2: 共享管理 ---
        tab_manage = tk.Frame(notebook, padx=10, pady=10)
        notebook.add(tab_manage, text='共享管理')
        self.create_manage_tab(tab_manage)


    def create_repair_tab(self, tab):
        
        # --- 共享权限设置区 ---
        frame_sharing = tk.LabelFrame(tab, text="共享访问和密码问题 (Win7/10/11 兼容)", padx=10, pady=10)
        frame_sharing.pack(padx=10, pady=10, fill="x")

        tk.Button(frame_sharing, text="禁用密码保护共享 (实现无密码访问)", command=lambda: threading.Thread(target=lambda: toggle_password_protected_sharing(False)).start()).pack(fill="x", pady=5)
        tk.Button(frame_sharing, text="启用密码保护共享 (恢复默认)", command=lambda: threading.Thread(target=lambda: toggle_password_protected_sharing(True)).start()).pack(fill="x", pady=5)
        
        tk.Button(frame_sharing, text="[Win10/11] 启用不安全访客登录 (解决访问被拒)", command=lambda: threading.Thread(target=enable_insecure_guest_logons).start()).pack(fill="x", pady=5)

        # --- 网络发现区 ---
        frame_discovery = tk.LabelFrame(tab, text="连接和网络发现故障", padx=10, pady=10)
        frame_discovery.pack(padx=10, pady=10, fill="x")
        
        tk.Button(frame_discovery, text="检查网络发现状态", command=lambda: threading.Thread(target=check_network_discovery).start()).pack(side=tk.LEFT, expand=True, padx=5)
        tk.Button(frame_discovery, text="一键修复网络发现", command=lambda: threading.Thread(target=fix_network_discovery).start()).pack(side=tk.RIGHT, expand=True, padx=5)


    def create_manage_tab(self, tab):
        
        # --- 驱动器/文件夹共享区 ---
        frame_drive = tk.LabelFrame(tab, text="驱动器/文件夹共享 (Net Share + ICACLS)", padx=10, pady=10)
        frame_drive.pack(padx=10, pady=10, fill="x")

        tk.Button(frame_drive, text="创建新的共享并设置权限 (推荐)", command=show_create_share_dialog).pack(fill="x", pady=5)
        tk.Label(frame_drive, text="此功能使用 CMD 工具，确保 Win7 到 Win11 的最大兼容性。").pack(pady=5)

        # --- 打印机共享区 ---
        frame_printer = tk.LabelFrame(tab, text="打印机共享 (Set-Printer)", padx=10, pady=10)
        frame_printer.pack(padx=10, pady=10, fill="x")
        
        tk.Button(frame_printer, text="共享本地打印机", command=lambda: PrinterShareDialog(self)).pack(fill="x", pady=5)
        tk.Label(frame_printer, text="* 提示: 打印机驱动程序可能仍需手动处理，尤其是在 Win11 客户端。").pack(pady=5)
        
        # --- 提示 ---
        tk.Label(tab, text="* 访问路径示例: \\\\您的电脑名\\共享名", fg="blue").pack(pady=10)


if __name__ == "__main__":
    app = SMBSimpleFixer()
    app.mainloop()