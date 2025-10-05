import os
import shutil
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sys
import random

class TableRepairSoftware:
    def __init__(self, root):
        self.root = root
        self.root.title("底层碎片重组表格修复重建软件")
        self.root.geometry("900x950")  # 适当调整窗口大小
        self.root.resizable(True, True)
        
        # 窗口居中显示
        self.center_window()
        
        # 设置中文字体支持
        self.setup_fonts()
        
        # 搜索相关变量
        self.searching = False
        self.stop_search = False
        self.results = []
        self.selected_drives = []
        # 定义需要搜索的关键词
        self.keywords = ["kingsoft", "office6", "backup"]
        self.animation_running = False
        
        # 创建UI
        self.create_widgets()
        
    def center_window(self):
        """使窗口在屏幕中居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        
        self.root.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
    def setup_fonts(self):
        # 确保中文正常显示
        default_font = ('SimHei', 10)
        self.root.option_add("*Font", default_font)
        
    def create_widgets(self):
        # 标题部分
        title_frame = ttk.Frame(self.root, padding="10")
        title_frame.pack(fill=tk.X)
        
        title_label = ttk.Label(
            title_frame, 
            text="底层碎片重组表格修复重建软件", 
            font=('SimHei', 16, 'bold'),
            foreground="red"  # 标题为红色
        )
        title_label.pack(pady=10)
        
        # 关键词显示
        keywords_frame = ttk.Frame(self.root)
        keywords_frame.pack(fill=tk.X, padx=20)
        
        keywords_text = f"搜索关键词: {', '.join(self.keywords)}"
        ttk.Label(
            keywords_frame, 
            text=keywords_text, 
            font=('SimHei', 10, 'italic')
        ).pack(anchor=tk.W)
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 驱动器选择部分
        drive_frame = ttk.LabelFrame(main_frame, text="选择搜索驱动器（可多选）", padding="5")
        drive_frame.pack(fill=tk.X, pady=5)
        
        # 创建水平滚动条容器
        scroll_container = ttk.Frame(drive_frame)
        scroll_container.pack(fill=tk.X, padx=5, pady=2)  # 减少上下内边距
        
        # 创建水平滚动条
        drive_scroll = ttk.Scrollbar(scroll_container, orient=tk.HORIZONTAL)
        drive_scroll.pack(fill=tk.X, side=tk.BOTTOM)
        
        # 驱动器列表（使用Canvas实现水平滚动）
        drive_canvas = tk.Canvas(scroll_container, xscrollcommand=drive_scroll.set, height=40)  # 固定高度减少空白
        drive_canvas.pack(fill=tk.X)
        
        drive_scroll.config(command=drive_canvas.xview)
        
        self.drive_frame = ttk.Frame(drive_canvas)
        drive_canvas.create_window((0, 0), window=self.drive_frame, anchor=tk.W)
        
        # 获取可用驱动器并显示为复选框
        self.update_drives()
        
        # 更新滚动区域
        self.drive_frame.update_idletasks()
        drive_canvas.config(scrollregion=drive_canvas.bbox("all"))
        
        # 搜索控制按钮 - 移到驱动器框架下方，减少水平空白
        btn_frame = ttk.Frame(drive_frame)
        btn_frame.pack(side=tk.LEFT, padx=5, pady=2)  # 减少边距
        
        self.search_btn = ttk.Button(btn_frame, text="开始搜索", command=self.start_search)
        self.search_btn.pack(side=tk.LEFT, padx=3)  # 减少按钮间距
        
        self.stop_btn = ttk.Button(btn_frame, text="停止搜索", command=self.stop_searching, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=3)
        
        # 全选/取消全选按钮
        self.select_all_btn = ttk.Button(btn_frame, text="全选", command=self.select_all_drives)
        self.select_all_btn.pack(side=tk.LEFT, padx=3)
        
        self.deselect_all_btn = ttk.Button(btn_frame, text="取消全选", command=self.deselect_all_drives)
        self.deselect_all_btn.pack(side=tk.LEFT, padx=3)
        
        # 进度条和动画区域
        progress_frame = ttk.LabelFrame(main_frame, text="系统扫描状态", padding="5")
        progress_frame.pack(fill=tk.X, pady=5)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, length=100)
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)
        
        self.status_var = tk.StringVar(value="就绪")
        self.status_label = ttk.Label(progress_frame, textvariable=self.status_var)
        self.status_label.pack(anchor=tk.W, padx=5)
        
        # 代码动画显示区域
        self.animation_frame = ttk.LabelFrame(main_frame, text="底层数据解析", padding="5")
        self.animation_frame.pack(fill=tk.X, pady=5)
        
        # 创建黑底绿字的文本区域用于显示代码动画
        self.animation_text = tk.Text(
            self.animation_frame, 
            bg="black", 
            fg="#00FF00", 
            font=('Consolas', 10),
            height=8,
            wrap=tk.WORD,
            state=tk.DISABLED
        )
        self.animation_text.pack(fill=tk.X, padx=5, pady=5)
        
        # 结果显示部分
        result_frame = ttk.LabelFrame(main_frame, text="搜索结果 - WPS缓存目录", padding="5")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 结果列表
        columns = ("路径",)
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings")
        self.result_tree.heading("路径", text="缓存目录路径")
        self.result_tree.column("路径", width=900)
        
        # 添加滚动条
        scrollbar_y = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        scrollbar_x = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        self.result_tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)
        
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 复制部分
        copy_frame = ttk.LabelFrame(main_frame, text="复制选项", padding="5")
        copy_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(copy_frame, text="目标目录:").pack(side=tk.LEFT, padx=5)
        
        self.target_dir_var = tk.StringVar(value="")
        self.target_dir_entry = ttk.Entry(copy_frame, textvariable=self.target_dir_var, width=70)
        self.target_dir_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        self.browse_btn = ttk.Button(copy_frame, text="浏览...", command=self.browse_target_dir)
        self.browse_btn.pack(side=tk.LEFT, padx=5)
        
        self.copy_btn = ttk.Button(copy_frame, text="复制选中项", command=self.copy_selected)
        self.copy_btn.pack(side=tk.LEFT, padx=5)
        
        # 绑定双击事件
        self.result_tree.bind("<Double-1>", lambda e: self.copy_selected())
        
        # 底部信息
        bottom_frame = ttk.Frame(self.root, padding="10")
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        ttk.Label(
            bottom_frame, 
            text="小凯数码 专业数据恢复", 
            font=('SimHei', 12, 'bold'),
            foreground="red"  # 底部信息为红色
        ).pack()
        
    def update_drives(self):
        """更新可用驱动器列表，减少空白"""
        # 清除现有复选框
        for widget in self.drive_frame.winfo_children():
            widget.destroy()
            
        self.drive_vars = []
        drives = []
        
        if sys.platform.startswith('win'):
            # Windows系统
            for drive in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
                drive_path = f"{drive}:/"
                if os.path.exists(drive_path):
                    drives.append(drive_path)
        else:
            # 非Windows系统，添加常见根目录
            common_roots = ['/', '/home']
            for root in common_roots:
                if os.path.exists(root):
                    drives.append(root)
        
        # 创建复选框，减少间距
        for i, drive in enumerate(drives):
            var = tk.BooleanVar()
            self.drive_vars.append((var, drive))
            
            chk = ttk.Checkbutton(
                self.drive_frame, 
                text=drive, 
                variable=var,
                command=self.update_selected_drives
            )
            chk.pack(side=tk.LEFT, padx=5, pady=2)  # 显著减少水平和垂直间距
    
    def update_selected_drives(self):
        """更新选中的驱动器列表"""
        self.selected_drives = [drive for var, drive in self.drive_vars if var.get()]
    
    def select_all_drives(self):
        """全选驱动器"""
        for var, _ in self.drive_vars:
            var.set(True)
        self.update_selected_drives()
    
    def deselect_all_drives(self):
        """取消全选驱动器"""
        for var, _ in self.drive_vars:
            var.set(False)
        self.update_selected_drives()
    
    def start_search(self):
        """开始搜索线程"""
        if self.searching:
            return
            
        # 检查是否选择了驱动器
        self.update_selected_drives()
        if not self.selected_drives:
            messagebox.showwarning("警告", "请至少选择一个驱动器")
            return
            
        # 重置状态
        self.results = []
        self.result_tree.delete(*self.result_tree.get_children())
        self.searching = True
        self.stop_search = False
        self.progress_var.set(0)
        self.status_var.set("开始搜索...")
        
        # 启动代码动画
        self.animation_running = True
        self.root.after(100, self.update_animation)
        
        # 更新按钮状态
        self.search_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        # 禁用驱动器选择相关控件
        for widget in self.drive_frame.winfo_children():
            if isinstance(widget, ttk.Checkbutton):
                widget.config(state=tk.DISABLED)
        self.select_all_btn.config(state=tk.DISABLED)
        self.deselect_all_btn.config(state=tk.DISABLED)
        
        # 启动搜索线程
        search_thread = threading.Thread(target=self.search, args=(self.selected_drives,))
        search_thread.daemon = True
        search_thread.start()
        
        # 启动进度更新检查
        self.root.after(100, self.check_search_progress)
    
    def stop_searching(self):
        """停止搜索"""
        if self.searching:
            self.stop_search = True
            self.status_var.set("正在停止搜索...")
    
    def update_animation(self):
        """更新代码动画显示"""
        if not self.animation_running:
            return
            
        # 生成随机的0和1组成的代码行
        line_length = 90
        animation_chars = ['0', '1', ' ', '0', '1', ' ', '0', '1', ';', ':', '#', '@', '$', '%', '&']
        line = ''.join(random.choice(animation_chars) for _ in range(line_length))
        
        # 随机添加一些模拟状态的信息
        status_messages = [
            "正在解析磁盘分区表...",
            "扫描文件分配表...",
            "重建文件碎片索引...",
            "识别WPS缓存签名...",
            "验证文件完整性...",
            "分析数据块结构...",
            "读取扇区数据...",
            "恢复丢失的索引节点...",
            "重组分散的数据块..."
        ]
        
        # 每10行插入一条状态信息
        if random.random() < 0.1:  # 10%的概率插入状态信息
            line = random.choice(status_messages)
        
        # 更新动画区域
        self.animation_text.config(state=tk.NORMAL)
        self.animation_text.insert(tk.END, line + "\n")
        
        # 保持与文本框高度匹配的行数
        current_lines = int(self.animation_text.index('end-1c').split('.')[0])
        if current_lines > 8:
            self.animation_text.delete('1.0', '2.0')
        
        self.animation_text.see(tk.END)
        self.animation_text.config(state=tk.DISABLED)
        
        # 继续动画
        if self.searching:
            self.root.after(random.randint(100, 300), self.update_animation)
    
    def search(self, drives):
        """搜索同时包含所有关键词的目录"""
        try:
            # 首先获取所有文件和目录的总数，用于进度计算
            total_items = 0
            for drive in drives:
                if self.stop_search:
                    break
                for dirpath, _, filenames in os.walk(drive):
                    if self.stop_search:
                        break
                    total_items += len(filenames) + 1  # 目录本身 + 文件数
                
            processed_items = 0
            
            # 实际搜索
            for drive in drives:
                if self.stop_search:
                    break
                    
                for dirpath, _, filenames in os.walk(drive):
                    if self.stop_search:
                        break
                        
                    # 转换为小写以进行不区分大小写的匹配
                    lower_dirpath = dirpath.lower()
                    
                    # 检查当前目录是否包含所有关键词
                    contains_all = True
                    for keyword in self.keywords:
                        if keyword.lower() not in lower_dirpath:
                            contains_all = False
                            break
                            
                    if contains_all:
                        self.results.append(dirpath)
                        # 在主线程中更新UI
                        self.root.after(0, self.add_result, dirpath)
                    
                    # 更新进度
                    processed_items += len(filenames) + 1
                    if total_items > 0:
                        progress = (processed_items / total_items) * 100
                        self.root.after(0, self.update_progress, progress, 
                                       f"搜索中: {dirpath} ({processed_items}/{total_items})")
        
        except Exception as e:
            self.root.after(0, self.show_error, str(e))
        
        finally:
            # 搜索完成，更新状态
            self.root.after(0, self.search_complete)
    
    def add_result(self, path):
        """添加搜索结果到列表"""
        self.result_tree.insert("", tk.END, values=(path,))
        self.status_var.set(f"已找到 {len(self.results)} 个匹配项")
    
    def update_progress(self, value, status):
        """更新进度条和状态"""
        self.progress_var.set(value)
        self.status_var.set(status)
    
    def show_error(self, message):
        """显示错误信息"""
        messagebox.showerror("错误", f"搜索过程中发生错误:\n{message}")
    
    def search_complete(self):
        """搜索完成后的清理工作"""
        self.searching = False
        self.stop_search = False
        self.animation_running = False
        self.progress_var.set(100)
        
        # 搜索完成后在动画区域显示完成信息
        self.animation_text.config(state=tk.NORMAL)
        self.animation_text.delete('1.0', tk.END)
        self.animation_text.insert(tk.END, "扫描完成，数据解析结束\n")
        self.animation_text.config(state=tk.DISABLED)
        
        if len(self.results) == 0:
            self.status_var.set("搜索完成，未找到匹配的缓存目录")
        else:
            self.status_var.set(f"搜索完成，共找到 {len(self.results)} 个匹配的缓存目录")
        
        # 恢复按钮状态
        self.search_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        
        # 启用驱动器选择相关控件
        for widget in self.drive_frame.winfo_children():
            if isinstance(widget, ttk.Checkbutton):
                widget.config(state=tk.NORMAL)
        self.select_all_btn.config(state=tk.NORMAL)
        self.deselect_all_btn.config(state=tk.NORMAL)
    
    def check_search_progress(self):
        """检查搜索进度并更新UI"""
        if self.searching:
            self.root.after(100, self.check_search_progress)
    
    def browse_target_dir(self):
        """浏览目标目录"""
        dir_path = filedialog.askdirectory(title="选择目标目录")
        if dir_path:
            self.target_dir_var.set(dir_path)
    
    def copy_selected(self):
        """复制选中的目录到目标位置"""
        selected_items = self.result_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要复制的目录")
            return
            
        target_dir = self.target_dir_var.get()
        if not target_dir or not os.path.exists(target_dir):
            messagebox.showwarning("警告", "请选择有效的目标目录")
            return
        
        try:
            # 获取选中的路径
            selected_path = self.result_tree.item(selected_items[0])['values'][0]
            
            # 创建目标路径
            dest_path = os.path.join(target_dir, os.path.basename(selected_path))
            
            # 复制目录
            if os.path.exists(dest_path):
                if not messagebox.askyesno("确认", f"目标目录 '{dest_path}' 已存在，是否覆盖?"):
                    return
                # 先删除已存在的目录
                shutil.rmtree(dest_path)
            
            shutil.copytree(selected_path, dest_path)
            messagebox.showinfo("成功", f"目录已成功复制到:\n{dest_path}")
            
        except Exception as e:
            messagebox.showerror("复制失败", f"复制过程中发生错误:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TableRepairSoftware(root)
    root.mainloop()
    