import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys

# 注册信息文件路径（需与主程序一致）
REGISTRATION_FILE = os.path.join(os.path.expanduser("~"), ".mssql_backup_tool_reg.dat")

class LicenseRemover:
    def __init__(self, root):
        self.root = root
        self.root.title("注册码删除器")
        self.root.geometry("400x200")
        self.root.resizable(False, False)
        
        # 窗口居中
        self.center_window()
        
        # 设置字体
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("SimHei", 10))
        self.style.configure("TButton", font=("SimHei", 10))
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        ttk.Label(
            main_frame, 
            text="MSSQL数据库备份工具 注册码删除器", 
            font=("SimHei", 12, "bold")
        ).pack(pady=10)
        
        # 说明文本
        ttk.Label(
            main_frame, 
            text="此工具用于删除已有的注册信息，\n使软件恢复到未注册状态（重新开始试用期）。", 
            font=("SimHei", 10),
            justify=tk.CENTER
        ).pack(pady=20)
        
        # 按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(
            btn_frame, 
            text="删除注册信息", 
            command=self.remove_license,
            style="Accent.TButton"
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            btn_frame, 
            text="检查注册状态", 
            command=self.check_license_status
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            btn_frame, 
            text="退出", 
            command=root.quit
        ).pack(side=tk.RIGHT, padx=5)
        
        # 底部信息
        ttk.Label(
            main_frame, 
            text="© 2025 版权所有", 
            font=("SimHei", 9)
        ).pack(side=tk.BOTTOM, pady=10)

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def check_license_status(self):
        """检查当前注册状态"""
        if os.path.exists(REGISTRATION_FILE):
            messagebox.showinfo("注册状态", "软件已注册，存在注册信息文件。")
        else:
            messagebox.showinfo("注册状态", "软件未注册，未找到注册信息文件。")
    
    def remove_license(self):
        """删除注册信息"""
        if not os.path.exists(REGISTRATION_FILE):
            messagebox.showinfo("提示", "未找到注册信息文件，无需删除。")
            return
            
        if messagebox.askyesno("确认删除", "确定要删除注册信息吗？\n此操作将使软件恢复到未注册状态，重新开始试用期。"):
            try:
                os.remove(REGISTRATION_FILE)
                if not os.path.exists(REGISTRATION_FILE):
                    messagebox.showinfo("成功", "注册信息已成功删除！\n请重启软件以应用更改。")
                else:
                    messagebox.showerror("失败", "删除注册信息失败，请手动删除文件：\n" + REGISTRATION_FILE)
            except Exception as e:
                messagebox.showerror("错误", f"删除注册信息失败：{str(e)}\n请手动删除文件：\n" + REGISTRATION_FILE)

if __name__ == "__main__":
    root = tk.Tk()
    # 设置按钮样式
    root.style = ttk.Style()
    root.style.configure("Accent.TButton", foreground="red")
    app = LicenseRemover(root)
    root.mainloop()
    