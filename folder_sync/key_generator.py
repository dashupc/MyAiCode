import tkinter as tk
from tkinter import ttk, messagebox
import hashlib
import string

class HardwareRegistrationSystem:
    """与自动备份工具完全一致的注册系统，确保生成的注册码有效"""
    def __init__(self):
        # 固定盐值 - 必须与自动备份工具中使用的盐值完全相同
        self._fixed_salt = "BackupTool_v2.2_2024_Authorized"
        self.debug_info = []  # 存储调试信息
        
    def get_debug_info(self):
        """获取调试信息"""
        return "\n".join(self.debug_info)
        
    def clear_debug_info(self):
        """清空调试信息"""
        self.debug_info = []
    
    def generate_registration_key(self, machine_code):
        """生成注册码 - 与自动备份工具使用完全相同的算法"""
        self.clear_debug_info()
        self.debug_info.append("开始生成注册码...")
        
        if not machine_code:
            raise ValueError("无效的机器码: 空值")
            
        # 移除机器码中的横线
        cleaned_machine_code = machine_code.replace('-', '')
        self.debug_info.append(f"机器码: {machine_code}")
        self.debug_info.append(f"移除横线后的机器码: {cleaned_machine_code}")
        
        # 验证机器码长度
        if len(cleaned_machine_code) != 20:
            raise ValueError(f"机器码格式错误，应为20位字符(去除横线后)，实际为{len(cleaned_machine_code)}位")
        
        # 生成基础字符串 - 机器码+固定盐值
        self.debug_info.append(f"使用固定盐值: {self._fixed_salt}")
        base_str = cleaned_machine_code + self._fixed_salt
        self.debug_info.append(f"组合字符串: {base_str}")
        
        # 计算哈希值
        hash_obj = hashlib.sha256(base_str.encode('utf-8'))
        hash_hex = hash_obj.hexdigest().upper()
        self.debug_info.append(f"SHA-256哈希结果: {hash_hex}")
        
        # 确保哈希值长度足够
        if len(hash_hex) < 36:
            raise ValueError(f"哈希值生成错误，长度不足: {len(hash_hex)}")
        
        # 从哈希值中提取并格式化为注册码格式
        parts = [
            hash_hex[0:4],   # 前4位
            hash_hex[8:12],  # 中间4位
            hash_hex[16:20], # 中间4位
            hash_hex[24:28], # 中间4位
            hash_hex[32:36]  # 后4位
        ]
        self.debug_info.append(f"提取的哈希部分: {parts}")
        
        registration_key = '-'.join(parts)
        self.debug_info.append(f"生成注册码: {registration_key}")
        
        # 验证注册码格式
        if len(registration_key) != 24:
            raise ValueError(f"注册码生成错误，长度应为24位，实际为{len(registration_key)}位")
            
        return registration_key


class KeyGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("自动备份工具 - 注册码生成器")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        # 创建注册系统实例
        self.reg_system = HardwareRegistrationSystem()
        
        # 确保中文显示正常
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("SimHei", 10))
        self.style.configure("TButton", font=("SimHei", 10))
        self.style.configure("TEntry", font=("SimHei", 10))
        self.style.configure("Header.TLabel", font=("SimHei", 12, "bold"))
        
        # 创建UI
        self.create_widgets()
        
        # 窗口居中
        self.center_window()
    
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
            text="自动备份工具注册码生成器", 
            style="Header.TLabel"
        ).grid(row=0, column=0, columnspan=2, pady=10)
        
        # 机器码输入
        ttk.Label(main_frame, text="请输入机器码:").grid(row=1, column=0, sticky=tk.W, pady=5)
        
        self.machine_code_var = tk.StringVar()
        # 添加示例文本，提示正确格式
        self.machine_code_entry = ttk.Entry(
            main_frame, 
            textvariable=self.machine_code_var, 
            width=50
        )
        self.machine_code_entry.grid(row=1, column=1, pady=5)
        self.machine_code_entry.insert(0, "例如: ABCD-EFGH-IJKL-MNOP-QRST")
        
        # 提示标签
        ttk.Label(
            main_frame, 
            text="机器码格式: 5组4位字符，用横线分隔 (共24个字符)",
            foreground="gray"
        ).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=2)
        
        # 生成按钮
        generate_btn = ttk.Button(
            main_frame, 
            text="生成注册码", 
            command=self.generate_key
        )
        generate_btn.grid(row=3, column=0, columnspan=2, pady=15)
        
        # 注册码结果
        ttk.Label(main_frame, text="生成的注册码:").grid(row=4, column=0, sticky=tk.W, pady=5)
        
        self.reg_key_var = tk.StringVar()
        self.reg_key_entry = ttk.Entry(
            main_frame, 
            textvariable=self.reg_key_var, 
            width=50,
            state="readonly"
        )
        self.reg_key_entry.grid(row=4, column=1, pady=5)
        
        # 复制按钮
        copy_btn = ttk.Button(
            main_frame, 
            text="复制注册码", 
            command=self.copy_to_clipboard
        )
        copy_btn.grid(row=5, column=0, columnspan=2, pady=5)
        
        # 调试信息区域
        ttk.Label(main_frame, text="生成过程信息:").grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        self.debug_text = tk.Text(main_frame, height=8, width=60, wrap=tk.WORD)
        self.debug_text.grid(row=7, column=0, columnspan=2, pady=5, sticky=tk.NSEW)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(main_frame, command=self.debug_text.yview)
        scrollbar.grid(row=7, column=2, sticky=tk.NS)
        self.debug_text.config(yscrollcommand=scrollbar.set)
        self.debug_text.config(state=tk.DISABLED)
        
        # 底部信息
        ttk.Label(
            main_frame, 
            text="注: 注册码与机器码一一对应，不同机器码需重新生成",
            foreground="gray"
        ).grid(row=8, column=0, columnspan=2, sticky=tk.W, pady=10)
        
        # 设置网格权重，使文本区域可伸缩
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(7, weight=1)
        
        # 绑定回车键生成注册码
        self.root.bind('<Return>', lambda event: self.generate_key())
    
    def generate_key(self):
        """生成注册码"""
        machine_code = self.machine_code_var.get().strip()
        
        # 简单验证机器码格式
        if not machine_code:
            messagebox.showerror("错误", "请输入机器码")
            return
            
        if len(machine_code) != 24 or machine_code.count('-') != 4:
            messagebox.showerror(
                "格式错误", 
                f"机器码格式不正确，应为24个字符(包含4个横线)，实际为{len(machine_code)}个字符"
            )
            return
        
        try:
            # 生成注册码
            reg_key = self.reg_system.generate_registration_key(machine_code)
            
            # 显示结果
            self.reg_key_var.set(reg_key)
            
            # 显示调试信息
            self.debug_text.config(state=tk.NORMAL)
            self.debug_text.delete(1.0, tk.END)
            self.debug_text.insert(tk.END, self.reg_system.get_debug_info())
            self.debug_text.config(state=tk.DISABLED)
            
        except Exception as e:
            messagebox.showerror("生成失败", f"注册码生成失败: {str(e)}")
            
            # 显示错误调试信息
            self.debug_text.config(state=tk.NORMAL)
            self.debug_text.delete(1.0, tk.END)
            self.debug_text.insert(tk.END, self.reg_system.get_debug_info())
            self.debug_text.insert(tk.END, f"\n错误: {str(e)}")
            self.debug_text.config(state=tk.DISABLED)
    
    def copy_to_clipboard(self):
        """将注册码复制到剪贴板"""
        reg_key = self.reg_key_var.get()
        if reg_key:
            self.root.clipboard_clear()
            self.root.clipboard_append(reg_key)
            messagebox.showinfo("成功", "注册码已复制到剪贴板")
        else:
            messagebox.showwarning("警告", "没有可复制的注册码，请先生成")


def main():
    root = tk.Tk()
    app = KeyGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
    