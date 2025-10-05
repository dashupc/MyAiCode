import tkinter as tk
from tkinter import ttk, messagebox
import hashlib
import sys

class LicenseGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("注册码生成器")
        self.root.geometry("500x300")
        self.root.resizable(False, False)
        
        # 窗口居中
        self.center_window()
        
        # 设置字体
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("SimHei", 10))
        self.style.configure("TButton", font=("SimHei", 10))
        self.style.configure("TEntry", font=("SimHei", 10))
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        ttk.Label(
            main_frame, 
            text="MSSQL数据库备份工具 注册码生成器", 
            font=("SimHei", 14, "bold")
        ).pack(pady=10)
        
        # 机器码输入
        machine_frame = ttk.Frame(main_frame)
        machine_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(machine_frame, text="机器码:").pack(side=tk.LEFT, padx=5)
        self.machine_code_entry = ttk.Entry(machine_frame, width=40)
        self.machine_code_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 生成按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=20)
        
        ttk.Button(
            btn_frame, 
            text="生成注册码", 
            command=self.generate_license
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            btn_frame, 
            text="清空", 
            command=self.clear_fields
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            btn_frame, 
            text="复制注册码", 
            command=self.copy_license
        ).pack(side=tk.RIGHT, padx=5)
        
        # 注册码显示
        license_frame = ttk.Frame(main_frame)
        license_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(license_frame, text="注册码:").pack(side=tk.LEFT, padx=5)
        self.license_entry = ttk.Entry(license_frame, width=40, state="readonly")
        self.license_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
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
    
    def generate_license(self):
        """根据机器码生成注册码"""
        machine_code = self.machine_code_entry.get().strip().replace('-', '')
        
        if not machine_code:
            messagebox.showerror("错误", "请输入机器码")
            return
            
        try:
            # 生成注册码的核心算法
            # 实际应用中可以使用更复杂的加密算法和密钥
            hash_obj = hashlib.sha256((machine_code + "LICENSE_KEY").encode())
            license_code = hash_obj.hexdigest()[:20].upper()
            
            # 添加分隔符，便于阅读
            formatted_license = '-'.join([license_code[i:i+5] for i in range(0, len(license_code), 5)])
            
            # 显示注册码
            self.license_entry.config(state="normal")
            self.license_entry.delete(0, tk.END)
            self.license_entry.insert(0, formatted_license)
            self.license_entry.config(state="readonly")
            
        except Exception as e:
            messagebox.showerror("错误", f"生成注册码失败: {str(e)}")
    
    def clear_fields(self):
        """清空输入框"""
        self.machine_code_entry.delete(0, tk.END)
        self.license_entry.config(state="normal")
        self.license_entry.delete(0, tk.END)
        self.license_entry.config(state="readonly")
    
    def copy_license(self):
        """复制注册码到剪贴板"""
        license_code = self.license_entry.get()
        if license_code:
            self.root.clipboard_clear()
            self.root.clipboard_append(license_code)
            self.root.update()
            messagebox.showinfo("成功", "注册码已复制到剪贴板")
        else:
            messagebox.showinfo("提示", "没有可复制的注册码")

if __name__ == "__main__":
    root = tk.Tk()
    app = LicenseGenerator(root)
    root.mainloop()
    