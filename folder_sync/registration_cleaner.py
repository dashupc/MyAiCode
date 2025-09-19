import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import winreg
import os
import sys
from datetime import datetime

class RegistrationCleaner:
    def __init__(self, root):
        self.root = root
        self.root.title("注册信息清除工具 - 调试专用")
        self.root.geometry("600x450")
        self.root.resizable(True, True)
        
        # 自动备份工具的注册表路径
        self.REG_KEY_PATH = r"Software\BackupTool"
        
        # 确保中文显示正常
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("SimHei", 10))
        self.style.configure("TButton", font=("SimHei", 10))
        self.style.configure("Header.TLabel", font=("SimHei", 12, "bold"))
        self.style.configure("Warning.TLabel", font=("SimHei", 10), foreground="red")
        
        # 创建UI
        self.create_widgets()
        
        # 窗口居中
        self.center_window()
        
        # 加载当前注册信息
        self.load_registration_info()
    
    def center_window(self):
        """使窗口在屏幕中居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def create_widgets(self):
        """创建GUI组件"""
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        ttk.Label(
            main_frame, 
            text="自动备份工具 - 注册信息清除器", 
            style="Header.TLabel"
        ).grid(row=0, column=0, columnspan=2, pady=10)
        
        # 警告信息
        ttk.Label(
            main_frame, 
            text="警告: 此工具用于调试，会删除所有注册信息和试用期数据！", 
            style="Warning.TLabel"
        ).grid(row=1, column=0, columnspan=2, pady=5, sticky=tk.W)
        
        # 当前注册信息区域
        ttk.Label(
            main_frame, 
            text="当前注册信息:"
        ).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=10)
        
        self.info_text = scrolledtext.ScrolledText(main_frame, height=10, width=60, wrap=tk.WORD)
        self.info_text.grid(row=3, column=0, columnspan=2, pady=5, sticky=tk.NSEW)
        self.info_text.config(state=tk.DISABLED)
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)
        
        self.delete_btn = ttk.Button(
            button_frame, 
            text="删除所有注册信息", 
            command=self.delete_registration_info,
            style="TButton"
        )
        self.delete_btn.pack(side=tk.LEFT, padx=10)
        
        self.refresh_btn = ttk.Button(
            button_frame, 
            text="刷新信息", 
            command=self.load_registration_info
        )
        self.refresh_btn.pack(side=tk.LEFT, padx=10)
        
        self.close_btn = ttk.Button(
            button_frame, 
            text="关闭", 
            command=self.root.destroy
        )
        self.close_btn.pack(side=tk.LEFT, padx=10)
        
        # 操作日志
        ttk.Label(
            main_frame, 
            text="操作日志:"
        ).grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=10)
        
        self.log_text = scrolledtext.ScrolledText(main_frame, height=5, width=60, wrap=tk.WORD)
        self.log_text.grid(row=6, column=0, columnspan=2, pady=5, sticky=tk.NSEW)
        self.log_text.config(state=tk.DISABLED)
        
        # 设置网格权重
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        main_frame.rowconfigure(6, weight=1)
    
    def log(self, message):
        """添加操作日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)  # 滚动到最新内容
        self.log_text.config(state=tk.DISABLED)
    
    def load_registration_info(self):
        """加载当前注册信息"""
        self.info_text.config(state=tk.NORMAL)
        self.info_text.delete(1.0, tk.END)
        
        try:
            # 尝试打开注册表项
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                self.REG_KEY_PATH,
                0,
                winreg.KEY_READ
            )
            
            self.info_text.insert(tk.END, f"注册表路径: HKEY_CURRENT_USER\\{self.REG_KEY_PATH}\n\n")
            
            # 枚举所有值
            value_count = winreg.QueryInfoKey(key)[1]
            if value_count == 0:
                self.info_text.insert(tk.END, "未找到注册相关值")
            else:
                for i in range(value_count):
                    value_name, value_data, value_type = winreg.EnumValue(key, i)
                    self.info_text.insert(tk.END, f"{value_name}: {value_data}\n")
            
            winreg.CloseKey(key)
            self.log("已加载当前注册信息")
            
        except FileNotFoundError:
            self.info_text.insert(tk.END, f"未找到注册表项: HKEY_CURRENT_USER\\{self.REG_KEY_PATH}")
            self.log("未找到注册信息")
        except Exception as e:
            self.info_text.insert(tk.END, f"读取注册信息时出错: {str(e)}")
            self.log(f"读取注册信息出错: {str(e)}")
        
        self.info_text.config(state=tk.DISABLED)
    
    def delete_registration_info(self):
        """删除所有注册信息"""
        # 确认对话框
        if not messagebox.askyesno(
            "确认删除", 
            "确定要删除所有注册信息和试用期数据吗？\n这将重置软件的注册状态，需要重新注册或重新计算试用期。"
        ):
            return
        
        try:
            # 尝试删除整个注册表项
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, self.REG_KEY_PATH)
            self.log(f"成功删除注册表项: HKEY_CURRENT_USER\\{self.REG_KEY_PATH}")
            messagebox.showinfo("成功", "所有注册信息已成功删除！\n请重启自动备份工具以应用更改。")
            
            # 刷新显示
            self.load_registration_info()
            
        except FileNotFoundError:
            self.log("未找到注册信息，无需删除")
            messagebox.showinfo("信息", "未找到任何注册信息")
        except Exception as e:
            self.log(f"删除注册信息时出错: {str(e)}")
            messagebox.showerror("错误", f"删除注册信息失败: {str(e)}")


def main():
    root = tk.Tk()
    app = RegistrationCleaner(root)a
    root.mainloop()


if __name__ == "__main__":
    main()
    