import tkinter as tk
from tkinter import ttk, messagebox
import time
import threading
import sys
import ctypes
from datetime import datetime, timedelta
# 使用pynput库的低级鼠标控制
from pynput.mouse import Controller, Button

class AutoClicker:
    def __init__(self, root):
        self.root = root
        self.root.title("定时自动点击工具")
        # 先设置尺寸但不设置位置
        self.root.geometry("450x480")
        self.root.resizable(False, False)
        
        # 确保窗口创建后立即计算并设置居中位置
        self.root.update_idletasks()  # 强制更新窗口信息
        self.center_window()  # 调用居中方法
        
        # 设置窗口置顶
        self.root.attributes("-topmost", True)
        self.always_on_top = True  # 记录置顶状态
        
        # 初始化鼠标控制器
        self.mouse = Controller()
        
        # 点击区域变量
        self.region_selected = False
        self.region = (0, 0, 0, 0)  # (x, y, width, height)
        self.click_x = 0
        self.click_y = 0
        
        # 状态变量
        self.running = False
        self.thread = None
        self.interval_sequence = [120, 20]  # 默认间隔序列：2分钟(120秒)，20秒，循环
        self.current_interval_index = 0
        self.countdown_window = None
        
        # 点击方式选择
        self.click_method = tk.IntVar(value=1)  # 1: pynput低级事件, 2: 模拟按下释放
        
        # 创建UI
        self.create_widgets()
        
        # 确保中文显示正常
        self.root.option_add("*Font", "SimHei 10")
    
    def center_window(self):
        """改进的窗口居中方法，确保准确居中"""
        # 获取屏幕尺寸
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # 获取窗口尺寸
        window_width = self.root.winfo_width()
        window_height = self.root.winfo_height()
        
        # 计算精确的居中坐标
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        
        # 应用坐标，不改变窗口尺寸
        self.root.geometry(f"+{x}+{y}")
        
        # 再次更新以确保设置生效
        self.root.update_idletasks()
    
    def setup_fonts(self):
        """设置字体以支持中文显示"""
        if sys.platform.startswith('win'):
            default_font = ('SimHei', 10)
        else:
            default_font = ('WenQuanYi Micro Hei', 10)
        return default_font
    
    def create_widgets(self):
        """创建界面组件"""
        # 窗口置顶控制
        topmost_frame = ttk.Frame(self.root)
        topmost_frame.pack(pady=5, padx=10, fill=tk.X)
        
        self.topmost_var = tk.BooleanVar(value=True)
        self.topmost_check = ttk.Checkbutton(
            topmost_frame, 
            text="窗口置顶", 
            variable=self.topmost_var,
            command=self.toggle_topmost
        )
        self.topmost_check.pack(anchor=tk.W)
        
        # 点击方式选择
        click_method_frame = ttk.LabelFrame(self.root, text="点击方式选择")
        click_method_frame.pack(pady=5, padx=10, fill=tk.X)
        
        ttk.Radiobutton(
            click_method_frame, 
            text="低级鼠标事件（推荐用于特殊应用）", 
            variable=self.click_method, 
            value=1
        ).pack(anchor=tk.W, padx=10, pady=2)
        
        ttk.Radiobutton(
            click_method_frame, 
            text="模拟鼠标按下释放（兼容性较好）", 
            variable=self.click_method, 
            value=2
        ).pack(anchor=tk.W, padx=10, pady=2)
        
        # 时间间隔设置区域
        interval_frame = ttk.LabelFrame(self.root, text="点击间隔设置(秒)")
        interval_frame.pack(pady=5, padx=10, fill=tk.X)
        
        # 第一个间隔
        ttk.Label(interval_frame, text="第一个间隔:").grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        self.interval1_var = tk.StringVar(value=str(self.interval_sequence[0]))
        self.interval1_entry = ttk.Entry(interval_frame, textvariable=self.interval1_var, width=10)
        self.interval1_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(interval_frame, text="秒").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        
        # 第二个间隔
        ttk.Label(interval_frame, text="第二个间隔:").grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
        self.interval2_var = tk.StringVar(value=str(self.interval_sequence[1]))
        self.interval2_entry = ttk.Entry(interval_frame, textvariable=self.interval2_var, width=10)
        self.interval2_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Label(interval_frame, text="秒").grid(row=1, column=2, padx=5, pady=5, sticky=tk.W)
        
        # 应用间隔按钮
        self.apply_interval_btn = ttk.Button(
            interval_frame, text="应用间隔设置", command=self.apply_interval_settings
        )
        self.apply_interval_btn.grid(row=0, column=3, rowspan=2, padx=20, pady=5)
        
        # 区域选择按钮
        self.select_region_btn = ttk.Button(
            self.root, text="选择点击区域", command=self.select_region
        )
        self.select_region_btn.pack(pady=5)
        
        # 区域信息显示
        self.region_info = ttk.Label(self.root, text="未选择区域")
        self.region_info.pack(pady=5)
        
        # 状态显示
        self.status_var = tk.StringVar(value="状态：就绪")
        self.status_label = ttk.Label(self.root, textvariable=self.status_var)
        self.status_label.pack(pady=5)
        
        # 下一次点击时间显示
        self.next_click_var = tk.StringVar(value="下次点击：--")
        self.next_click_label = ttk.Label(self.root, textvariable=self.next_click_var)
        self.next_click_label.pack(pady=5)
        
        # 当前倒计时显示
        self.countdown_var = tk.StringVar(value="倒计时：--")
        self.countdown_label = ttk.Label(self.root, textvariable=self.countdown_var, font=("SimHei", 12, "bold"))
        self.countdown_label.pack(pady=10)
        
        # 控制按钮
        self.control_frame = ttk.Frame(self.root)
        self.control_frame.pack(pady=10)
        
        self.start_btn = ttk.Button(
            self.control_frame, text="开始", command=self.start_clicking
        )
        self.start_btn.pack(side=tk.LEFT, padx=10)
        
        self.stop_btn = ttk.Button(
            self.control_frame, text="停止", command=self.stop_clicking, state=tk.DISABLED
        )
        self.stop_btn.pack(side=tk.LEFT, padx=10)
        
        # 退出按钮
        self.exit_btn = ttk.Button(self.root, text="退出", command=self.root.quit)
        self.exit_btn.pack(pady=5)
    
    def toggle_topmost(self):
        """切换窗口是否置顶"""
        self.always_on_top = self.topmost_var.get()
        self.root.attributes("-topmost", self.always_on_top)
    
    def apply_interval_settings(self):
        """应用时间间隔设置"""
        try:
            # 获取输入的间隔值
            interval1 = int(self.interval1_var.get())
            interval2 = int(self.interval2_var.get())
            
            # 验证输入是否有效
            if interval1 <= 0 or interval2 <= 0:
                messagebox.showerror("输入错误", "时间间隔必须为正数")
                return
            
            # 更新间隔序列
            self.interval_sequence = [interval1, interval2]
            messagebox.showinfo("设置成功", f"已应用新的时间间隔：{interval1}秒 和 {interval2}秒")
            
        except ValueError:
            messagebox.showerror("输入错误", "请输入有效的整数")
    
    def select_region(self):
        """选择屏幕上的点击区域"""
        # 创建一个全屏半透明窗口用于选择区域
        overlay = tk.Toplevel(self.root)
        overlay.attributes("-fullscreen", True)
        overlay.attributes("-alpha", 0.3)
        overlay.attributes("-topmost", True)  # 确保选择窗口在最上层
        overlay.configure(bg="gray")
        
        # 绘制选择框的变量
        self.start_x = None
        self.start_y = None
        self.rect = None
        
        # 创建画布用于绘制选择框
        canvas = tk.Canvas(overlay, cursor="cross")
        canvas.pack(fill=tk.BOTH, expand=True)
        
        # 鼠标按下事件
        def on_press(event):
            self.start_x = event.x_root
            self.start_y = event.y_root
            self.rect = canvas.create_rectangle(0, 0, 0, 0, outline="red", width=2)
        
        # 鼠标拖动事件
        def on_drag(event):
            current_x, current_y = event.x_root, event.y_root
            canvas.coords(self.rect, self.start_x, self.start_y, current_x, current_y)
        
        # 鼠标释放事件
        def on_release(event):
            end_x, end_y = event.x_root, event.y_root
            # 计算区域
            x1 = min(self.start_x, end_x)
            y1 = min(self.start_y, end_y)
            x2 = max(self.start_x, end_x)
            y2 = max(self.start_y, end_y)
            
            self.region = (x1, y1, x2 - x1, y2 - y1)
            # 计算区域中心作为点击点
            self.click_x = x1 + (x2 - x1) // 2
            self.click_y = y1 + (y2 - y1) // 2
            
            self.region_selected = True
            self.region_info.config(
                text=f"区域：({x1}, {y1}) 到 ({x2}, {y2})"
            )
            
            overlay.destroy()
        
        # 绑定鼠标事件
        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)
        
        # 按ESC键取消选择
        def cancel(event):
            overlay.destroy()
        
        overlay.bind("<Escape>", cancel)
    
    def show_countdown_window(self, seconds):
        """显示倒计时窗口"""
        # 如果已有倒计时窗口，先关闭
        if self.countdown_window:
            self.countdown_window.destroy()
            
        # 创建倒计时窗口
        self.countdown_window = tk.Toplevel(self.root)
        self.countdown_window.overrideredirect(True)  # 无边框窗口
        self.countdown_window.attributes("-topmost", True)  # 置顶
        self.countdown_window.attributes("-alpha", 0.8)  # 半透明
        
        # 窗口位置：点击区域上方居中
        x = self.click_x - 50
        y = self.click_y - 80
        self.countdown_window.geometry(f"100x60+{x}+{y}")
        
        # 倒计时标签
        countdown_label = tk.Label(
            self.countdown_window, 
            text=str(seconds), 
            font=("SimHei", 24, "bold"),
            fg="red",
            bg="white"
        )
        countdown_label.pack(fill=tk.BOTH, expand=True)
        
        return countdown_label
    
    def perform_click(self):
        """执行点击操作，根据选择的方式"""
        # 移动鼠标到目标位置
        self.mouse.position = (self.click_x, self.click_y)
        time.sleep(0.05)  # 短暂延迟确保鼠标移动到位
        
        click_method = self.click_method.get()
        
        if click_method == 1:
            # 方法1：使用pynput的低级鼠标事件
            self.mouse.press(Button.left)
            time.sleep(0.05)  # 模拟真实点击的按下时间
            self.mouse.release(Button.left)
        else:
            # 方法2：模拟鼠标按下和释放（更接近物理点击）
            # 这里使用ctypes调用Windows API实现更底层的控制
            def mouse_event(x, y):
                ctypes.windll.user32.SetCursorPos(x, y)
                ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)  # 鼠标按下
                time.sleep(0.05)
                ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)  # 鼠标释放
            
            mouse_event(self.click_x, self.click_y)
    
    def start_clicking(self):
        """开始自动点击"""
        # 先应用最新的间隔设置
        self.apply_interval_settings()
        
        if not self.region_selected:
            messagebox.showwarning("警告", "请先选择点击区域")
            return
        
        if self.running:
            return
        
        self.running = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.status_var.set("状态：运行中")
        
        # 启动点击线程
        self.thread = threading.Thread(target=self.click_loop, daemon=True)
        self.thread.start()
    
    def stop_clicking(self):
        """停止自动点击"""
        if not self.running:
            return
        
        self.running = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.status_var.set("状态：已停止")
        self.next_click_var.set("下次点击：--")
        self.countdown_var.set("倒计时：--")
        
        # 关闭倒计时窗口
        if self.countdown_window:
            self.countdown_window.destroy()
            self.countdown_window = None
    
    def click_loop(self):
        """点击循环逻辑"""
        while self.running:
            # 执行点击（根据选择的方式）
            self.perform_click()
            
            # 获取当前间隔并切换到下一个
            current_interval = self.interval_sequence[self.current_interval_index]
            self.current_interval_index = (self.current_interval_index + 1) % len(self.interval_sequence)
            
            # 更新下次点击时间显示
            next_click_time = (datetime.now() + timedelta(seconds=current_interval)).strftime("%H:%M:%S")
            self.next_click_var.set(f"下次点击：{next_click_time} (约{current_interval}秒后)")
            
            # 显示倒计时
            countdown_label = None
            if self.running:
                # 在主线程中创建UI元素
                self.root.after(0, lambda: self.countdown_var.set(f"倒计时：{current_interval}秒"))
                countdown_label = self.show_countdown_window(current_interval)
            
            # 倒计时循环
            remaining = current_interval
            while remaining > 0 and self.running:
                time.sleep(1)
                remaining -= 1
                
                # 更新倒计时显示
                if self.running:
                    self.root.after(0, lambda r=remaining: self.countdown_var.set(f"倒计时：{r}秒"))
                    if countdown_label:
                        self.root.after(0, lambda r=remaining: countdown_label.config(text=str(r)))
            
            # 关闭倒计时窗口
            if self.countdown_window and self.running:
                self.root.after(0, self.countdown_window.destroy)
                self.root.after(0, lambda: setattr(self, 'countdown_window', None))
        
        # 循环结束时更新状态
        self.root.after(0, lambda: self.status_var.set("状态：已停止"))
        self.root.after(0, lambda: self.next_click_var.set("下次点击：--"))
        self.root.after(0, lambda: self.countdown_var.set("倒计时：--"))

def is_admin():
    """检查是否以管理员身份运行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if __name__ == "__main__":
    # 隐藏控制台窗口（仅在Windows有效）
    if sys.platform.startswith('win'):
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    
    # 请求管理员权限运行，提高对特殊应用的控制能力
    if not is_admin():
        try:
            # 尝试以管理员身份重启程序
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, " ".join(sys.argv), None, 1
            )
            sys.exit()
        except Exception as e:
            messagebox.showwarning("权限不足", f"程序可能无法在某些应用上正常工作：{str(e)}")
    
    root = tk.Tk()
    app = AutoClicker(root)
    root.mainloop()
