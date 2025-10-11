import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog, ttk
import subprocess
import sys
import ctypes
import os
import threading
import time

# --- 辅助函数：权限检查 ---

def is_admin():
    """检查程序是否以管理员权限运行。"""
    try:
        return sys.getwindowsversion().platform == 2 and ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False 

# --- 核心功能：创建共享 (使用临时批处理文件执行) ---

def create_shared_folder(path, share_name, permission):
    """
    将创建共享的命令写入临时 .bat 文件，然后使用 schtasks 以 SYSTEM 权限运行该批处理文件。
    这是避免引号转义错误的最佳方法。
    """
    
    # 1. 确定权限参数
    share_perm_msg = '完全控制' if permission == 'Full' else '只读'
    ntfs_perm = 'F' if permission == 'Full' else 'R' 
    share_perm = 'FULL' if permission == 'Full' else 'CHANGE' 

    normalized_path = os.path.normpath(path)
    
    # 保证文件和任务名唯一
    temp_dir = os.environ.get('TEMP') or 'C:\\Temp'
    batch_file_path = os.path.join(temp_dir, f"share_create_{os.getpid()}.bat")
    task_name = f"CreateShare_{share_name}_{os.getpid()}" 
    
    # 2. 写入批处理文件内容
    # 使用纯批处理命令，避免复杂的 CMD 逻辑转义
    batch_content = f"""
@echo off
REM 以 SYSTEM 权限执行共享创建和权限设置
REM 创建文件夹 (如果存在会跳过)
mkdir "{normalized_path}"

REM 删除旧的共享 (如果不存在会报错，但不会中断)
net share "{share_name}" /delete

REM 创建新的共享
net share "{share_name}"="{normalized_path}" /grant:Everyone,{share_perm} /remark:"Shared by Python Tool"

REM 设置 NTFS 权限
icacls "{normalized_path}" /grant Everyone:({ntfs_perm}) /T

exit /b %errorlevel%
"""
    
    # 3. 将命令写入 .bat 文件
    try:
        with open(batch_file_path, 'w', encoding='utf-8') as f:
            f.write(batch_content)
    except Exception as e:
        messagebox.showerror("文件写入错误", f"无法写入临时批处理文件: {e}")
        return

    # 4. 构建 schtasks 命令：创建任务，运行 .bat 文件
    # /tr 只需要简单引用 .bat 文件路径
    create_task_cmd = (
        f'schtasks /create /tn "{task_name}" /tr "cmd /c \\"{batch_file_path}\\"" '
        f'/sc once /st 00:00 /ru "System" /f'
    )
    
    run_task_cmd = f'schtasks /run /tn "{task_name}"'
    delete_task_cmd = f'schtasks /delete /tn "{task_name}" /f'

    output_msg = ""
    success = False
    
    try:
        # --- 任务执行阶段 ---
        
        # 1. 创建任务
        subprocess.run(create_task_cmd, check=True, shell=True, capture_output=True, encoding='gbk')
        
        # 2. 运行任务
        subprocess.run(run_task_cmd, check=True, shell=True, capture_output=True, encoding='gbk')
        time.sleep(2) # 留出时间执行批处理

        # --- 验证阶段 ---
        # 3. 验证是否成功创建
        check_cmd = f'powershell -Command "Get-SmbShare | Where-Object {{$_.Name -eq \\"{share_name}\\"}}"'
        check_result = subprocess.run(check_cmd, shell=True, capture_output=True, text=True, encoding='gbk', timeout=10)
        
        if check_result.stdout.strip():
            success = True
            output_msg = check_result.stdout
        else:
            output_msg = "SYSTEM任务执行完成，但共享未被系统服务完全识别。"
            
    except subprocess.CalledProcessError as e:
        error_output = e.stderr or e.stdout
        if isinstance(error_output, bytes):
            error_output = error_output.decode('gbk', errors='ignore')
            
        output_msg = (f"任务调度创建/运行失败。错误码: {e.returncode}\n"
                      f"错误输出:\n{error_output}")
    except Exception as e:
        output_msg = f"发生未知错误: {e}"
    finally:
        # 始终尝试删除临时文件和任务
        subprocess.run(delete_task_cmd, shell=True, capture_output=True)
        try:
            os.remove(batch_file_path)
        except:
            pass
    
    # 5. 结果反馈
    if not success:
         messagebox.showerror("创建共享失败 (最终尝试失败)", 
                              f"共享 '{share_name}' 创建失败。\n"
                              "原因：系统拒绝以任何命令行权限创建共享。\n\n"
                              f"诊断信息:\n{output_msg}")
    else:
        messagebox.showinfo("创建共享成功 (批处理文件绕过)", 
                            f"共享 '{share_name}' ({share_perm_msg}) 已使用系统最高权限创建。\n\n"
                            "请在命令提示符(CMD)中输入 'net share' **立即确认**。")


def show_create_share_dialog(app_instance):
    """显示创建共享的对话框并收集参数"""
    # 1. 选择路径
    path = filedialog.askdirectory(title="选择要共享的文件夹路径")
    if not path: return

    # 2. 获取共享名
    share_name = simpledialog.askstring("共享名称", "请输入共享名 (例如: MyDataShare)", parent=app_instance)
    if not share_name or not share_name.strip():
        messagebox.showerror("错误", "共享名不能为空。")
        return
    share_name = share_name.strip()
        
    # 3. 获取权限
    permission_choice = simpledialog.askstring("权限选择", "请输入权限级别 (Read 或 Full):", parent=app_instance)
    permission_choice = permission_choice.strip().capitalize() if permission_choice else ''
    
    if permission_choice not in ['Read', 'Full']:
        messagebox.showerror("错误", "权限级别必须是 'Read' 或 'Full'。")
        return

    # 4. 在单独的线程中运行耗时操作
    threading.Thread(target=lambda: create_shared_folder(path, share_name, permission_choice)).start()


# --- GUI 界面 (主窗口) ---

class ShareCreatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SMB 共享创建工具 (批处理终极修复版)")
        self.geometry("450x300")
        self.resizable(False, False)

        self.admin_status_label = None
        self.check_admin_status()
        self.create_widgets()

    def check_admin_status(self):
        """检查并显示管理员权限状态"""
        status_text = "警告: 未以管理员身份运行! 请右键以管理员身份运行。"
        status_color = "red"
        
        if is_admin():
            status_text = "状态: 已检测到管理员权限。"
            status_color = "green"
        
        self.admin_status_label = tk.Label(self, text=status_text, fg=status_color, font=('Arial', 12, 'bold'))
        self.admin_status_label.pack(pady=10)

    def create_widgets(self):
        
        frame = tk.Frame(self, padx=20, pady=10)
        frame.pack(padx=10, pady=10, fill="x")

        tk.Button(frame, 
                  text="[步骤 1] 设置新的驱动器/文件夹共享 (批处理 SYSTEM)", 
                  command=lambda: show_create_share_dialog(self),
                  height=2,
                  bg="#FF5722", fg="white", font=('Arial', 12, 'bold')).pack(fill="x", pady=15)

        tk.Label(frame, text="此版本通过创建临时批处理文件运行，以完美解决所有引号和权限问题。", fg="blue", wraplength=400).pack(pady=5)
        tk.Label(frame, text="如果此方法仍失败，请**彻底禁用或卸载**第三方安全软件，并检查 **gpedit.msc**。", fg="red", wraplength=400).pack(pady=5)
        tk.Label(frame, text="[步骤 2] 完成后请在 CMD 中输入 'net share' 立即确认。", fg="black").pack(pady=5)


if __name__ == "__main__":
    app = ShareCreatorApp()
    app.mainloop()