import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import pyttsx3
import json
import threading
import time
from datetime import datetime
import os

# 音频播放库
AUDIO_AVAILABLE = False
try:
    import pygame
    pygame.mixer.init()
    AUDIO_AVAILABLE = True
except Exception as e:
    print(f"警告: pygame初始化失败 - {e}")

class TimedBroadcastApp:
    def __init__(self, root):
        self.root = root
        self.root.title("定时播音")
        self.root.geometry("1400x800")
        self.root.configure(bg='#E8F4F8')
        
        # 初始化语音引擎
        self.engine = pyttsx3.init()
        self.engine.setProperty('rate', 150)
        self.engine.setProperty('volume', 1.0)
        
        # 任务列表
        self.tasks = []
        self.running = False
        self.task_file = "broadcast_tasks.json"
        self.current_page = "定时广播"
        
        # 创建必要的文件夹结构
        self.create_folder_structure()
        
        # 创建界面
        self.create_widgets()
        
        # 加载已保存的任务
        self.load_tasks()
        
        # 启动后台检查线程
        self.start_background_thread()
    
    def create_folder_structure(self):
        """创建必要的文件夹结构"""
        folders = [
            "提示音",
            "音频文件"
        ]
        
        for folder in folders:
            if not os.path.exists(folder):
                os.makedirs(folder)
                self.log(f"已创建文件夹: {folder}") if hasattr(self, 'log_text') else None
    
    def create_widgets(self):
        # 左侧导航栏
        self.nav_frame = tk.Frame(self.root, bg='#A8D8E8', width=160)
        self.nav_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.nav_frame.pack_propagate(False)
        
        # 导航按钮
        nav_buttons = [
            ("定时广播", ""),
            ("背景音乐", ""),
            ("立即播播", ""),
            ("节假日、调休", "节假日不播或、调休"),
            ("设置", ""),
            ("语音广告 制作", "")
        ]
        
        for i, (title, subtitle) in enumerate(nav_buttons):
            btn_frame = tk.Frame(self.nav_frame, bg='#5DADE2' if i == 0 else '#A8D8E8')
            btn_frame.pack(fill=tk.X, pady=1)
            
            btn = tk.Button(btn_frame, text=title, bg='#5DADE2' if i == 0 else '#A8D8E8',
                          fg='white' if i == 0 else 'black', font=('Microsoft YaHei', 11, 'bold'),
                          bd=0, padx=10, pady=8, anchor='w',
                          command=lambda t=title: self.switch_page(t))
            btn.pack(fill=tk.X)
            
            if subtitle:
                sub_label = tk.Label(btn_frame, text=subtitle, bg='#5DADE2' if i == 0 else '#A8D8E8',
                                   fg='#FF6B35' if i == 3 else ('#555' if i == 0 else '#666'),
                                   font=('Microsoft YaHei', 8), anchor='w', padx=10)
                sub_label.pack(fill=tk.X)
        
        # 主内容区域
        self.main_frame = tk.Frame(self.root, bg='white')
        self.main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建定时广播页面
        self.create_scheduled_broadcast_page()
    
    def switch_page(self, page_name):
        """切换页面"""
        self.current_page = page_name
        # 这里可以扩展其他页面
        if page_name == "定时广播":
            self.log(f"切换到: {page_name}")
        else:
            self.log(f"功能开发中: {page_name}")
    
    def create_scheduled_broadcast_page(self):
        """创建定时广播页面"""
        # 顶部标题和控制区
        top_frame = tk.Frame(self.main_frame, bg='white')
        top_frame.pack(fill=tk.X, padx=10, pady=10)
        
        title_label = tk.Label(top_frame, text="定时广播", font=('Microsoft YaHei', 14, 'bold'),
                              bg='white', fg='#2C5F7C')
        title_label.pack(side=tk.LEFT)
        
        # 控制按钮区
        btn_frame = tk.Frame(top_frame, bg='white')
        btn_frame.pack(side=tk.RIGHT)
        
        buttons = [
            ("添加节目", self.add_task, '#5DADE2'),
            ("删除", self.delete_task, '#E74C3C'),
            ("修改", self.edit_task, '#F39C12'),
            ("复制", self.copy_task, '#9B59B6'),
            ("上移", lambda: self.move_task(-1), '#3498DB'),
            ("下移", lambda: self.move_task(1), '#3498DB'),
            ("导入", self.import_tasks, '#1ABC9C'),
            ("导出", self.export_tasks, '#1ABC9C'),
            ("启用", self.enable_task, '#27AE60'),
            ("禁用", self.disable_task, '#95A5A6')
        ]
        
        for text, cmd, color in buttons:
            btn = tk.Button(btn_frame, text=text, command=cmd, bg=color, fg='white',
                          font=('Microsoft YaHei', 9), bd=0, padx=12, pady=5, cursor='hand2')
            btn.pack(side=tk.LEFT, padx=3)
        
        # 节目单统计
        stats_frame = tk.Frame(self.main_frame, bg='#F0F8FF')
        stats_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.stats_label = tk.Label(stats_frame, text="节目单：0", font=('Microsoft YaHei', 10),
                                   bg='#F0F8FF', fg='#2C5F7C', anchor='w', padx=10)
        self.stats_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 节目列表表格
        table_frame = tk.Frame(self.main_frame, bg='white')
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 创建表格
        columns = ('节目名称', '状态', '开始时间(可多个)', '延时秒', '音频或文字', '音量', '周几/几号', '日期范围')
        self.task_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=12)
        
        # 设置列宽
        col_widths = [200, 60, 140, 70, 300, 60, 100, 120]
        for col, width in zip(columns, col_widths):
            self.task_tree.heading(col, text=col)
            self.task_tree.column(col, width=width, anchor='w' if col == '节目名称' or col == '音频或文字' else 'center')
        
        self.task_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.task_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.task_tree.configure(yscrollcommand=scrollbar.set)
        
        # 正在播放区域
        playing_frame = tk.LabelFrame(self.main_frame, text="正在播：", font=('Microsoft YaHei', 10),
                                     bg='white', fg='#2C5F7C', padx=10, pady=5)
        playing_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.playing_text = scrolledtext.ScrolledText(playing_frame, height=3, font=('Microsoft YaHei', 9),
                                                     bg='#FFFEF0', wrap=tk.WORD)
        self.playing_text.pack(fill=tk.BOTH, expand=True)
        self.playing_text.insert('1.0', "等待播放...")
        
        # 日志区域
        log_frame = tk.LabelFrame(self.main_frame, text="日志：", font=('Microsoft YaHei', 10),
                                 bg='white', fg='#2C5F7C', padx=10, pady=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, font=('Microsoft YaHei', 9),
                                                 bg='#F9F9F9', wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 底部状态栏
        status_frame = tk.Frame(self.main_frame, bg='#E8F4F8', height=30)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        status_frame.pack_propagate(False)
        
        self.status_labels = []
        status_texts = ["当前时间", "系统状态", "播放状态", "任务数量"]
        for i, text in enumerate(status_texts):
            label = tk.Label(status_frame, text=f"{text}: --", font=('Microsoft YaHei', 9),
                           bg='#5DADE2' if i % 2 == 0 else '#7EC8E3', fg='white', padx=15, pady=5)
            label.pack(side=tk.LEFT, padx=2)
            self.status_labels.append(label)
        
        # 启动状态栏更新
        self.update_status_bar()
        self.log("定时播音软件已启动")
    
    def add_task(self):
        """添加任务 - 先选择类型"""
        choice_dialog = tk.Toplevel(self.root)
        choice_dialog.title("选择节目类型")
        choice_dialog.geometry("350x200")
        choice_dialog.resizable(False, False)
        choice_dialog.transient(self.root)
        choice_dialog.grab_set()
        
        # 居中显示
        choice_dialog.update_idletasks()
        x = (choice_dialog.winfo_screenwidth() // 2) - (350 // 2)
        y = (choice_dialog.winfo_screenheight() // 2) - (200 // 2)
        choice_dialog.geometry(f"350x200+{x}+{y}")
        
        main_frame = tk.Frame(choice_dialog, padx=20, pady=20, bg='#F0F0F0')
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        title_label = tk.Label(main_frame, text="请选择节目类型", 
                              font=('Microsoft YaHei', 13, 'bold'), fg='#2C5F7C', bg='#F0F0F0')
        title_label.pack(pady=15)
        
        btn_frame = tk.Frame(main_frame, bg='#F0F0F0')
        btn_frame.pack(expand=True)
        
        # 音频节目按钮
        audio_btn = tk.Button(btn_frame, text="🎵 音频节目", 
                             command=lambda: self.open_audio_dialog(choice_dialog),
                             bg='#5DADE2', fg='white', font=('Microsoft YaHei', 11, 'bold'),
                             bd=0, padx=30, pady=12, cursor='hand2', width=15)
        audio_btn.pack(pady=8)
        
        # 语音节目按钮
        voice_btn = tk.Button(btn_frame, text="🎙️ 语音节目", 
                             command=lambda: self.open_voice_dialog(choice_dialog),
                             bg='#3498DB', fg='white', font=('Microsoft YaHei', 11, 'bold'),
                             bd=0, padx=30, pady=12, cursor='hand2', width=15)
        voice_btn.pack(pady=8)
    
    def open_audio_dialog(self, parent_dialog):
        """打开音频节目添加对话框"""
        parent_dialog.destroy()
        
        dialog = tk.Toplevel(self.root)
        dialog.title("定时广播频道 - 音频节目")
        dialog.geometry("850x620")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg='#E8E8E8')
        
        main_frame = tk.Frame(dialog, bg='#E8E8E8', padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ========== 内容区域 ==========
        content_frame = tk.LabelFrame(main_frame, text="内容", font=('Microsoft YaHei', 11, 'bold'),
                                     bg='#E8E8E8', padx=10, pady=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=5)
        
        # 节目名称
        tk.Label(content_frame, text="节目名称:", font=('Microsoft YaHei', 9), bg='#E8E8E8').grid(
            row=0, column=0, sticky='e', padx=5, pady=5)
        name_entry = tk.Entry(content_frame, font=('Microsoft YaHei', 9), width=55)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        
        # 音频文件单选
        audio_type_var = tk.StringVar(value="single")
        
        tk.Label(content_frame, text="音频文件", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=1, column=0, sticky='e', padx=5, pady=5)
        
        audio_single_frame = tk.Frame(content_frame, bg='#E8E8E8')
        audio_single_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        
        tk.Radiobutton(audio_single_frame, text="", variable=audio_type_var, value="single",
                      bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT)
        
        audio_single_entry = tk.Entry(audio_single_frame, font=('Microsoft YaHei', 9), width=35, state='readonly')
        audio_single_entry.pack(side=tk.LEFT, padx=5)
        
        time_label = tk.Label(audio_single_frame, text="00:00", font=('Microsoft YaHei', 10), bg='#E8E8E8')
        time_label.pack(side=tk.LEFT, padx=10)
        
        def select_single_audio():
            filename = filedialog.askopenfilename(
                title="选择音频文件",
                filetypes=[("音频文件", "*.mp3 *.wav *.ogg *.flac"), ("所有文件", "*.*")]
            )
            if filename:
                audio_single_entry.config(state='normal')
                audio_single_entry.delete(0, tk.END)
                audio_single_entry.insert(0, filename)
                audio_single_entry.config(state='readonly')
        
        tk.Button(audio_single_frame, text="选取...", command=select_single_audio, bg='#D0D0D0',
                 font=('Microsoft YaHei', 9), bd=1, padx=15, pady=3).pack(side=tk.LEFT, padx=5)
        
        # 音频文件夹单选
        tk.Label(content_frame, text="音频文件夹", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=2, column=0, sticky='e', padx=5, pady=5)
        
        audio_folder_frame = tk.Frame(content_frame, bg='#E8E8E8')
        audio_folder_frame.grid(row=2, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        
        tk.Radiobutton(audio_folder_frame, text="", variable=audio_type_var, value="folder",
                      bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT)
        
        audio_folder_entry = tk.Entry(audio_folder_frame, font=('Microsoft YaHei', 9), width=50, state='readonly')
        audio_folder_entry.pack(side=tk.LEFT, padx=5)
        
        def select_folder():
            foldername = filedialog.askdirectory(title="选择音频文件夹")
            if foldername:
                audio_folder_entry.config(state='normal')
                audio_folder_entry.delete(0, tk.END)
                audio_folder_entry.insert(0, foldername)
                audio_folder_entry.config(state='readonly')
        
        tk.Button(audio_folder_frame, text="选取...", command=select_folder, bg='#D0D0D0',
                 font=('Microsoft YaHei', 9), bd=1, padx=15, pady=3).pack(side=tk.LEFT, padx=5)
        
        # 播放顺序
        play_order_frame = tk.Frame(content_frame, bg='#E8E8E8')
        play_order_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=5)
        
        play_order_var = tk.StringVar(value="sequential")
        tk.Radiobutton(play_order_frame, text="顺序播试", variable=play_order_var, value="sequential",
                      bg='#E8E8E8', font=('Microsoft YaHei', 9)).pack(side=tk.LEFT, padx=10)
        tk.Label(play_order_frame, text="i", fg='blue', bg='#E8E8E8', cursor='hand2').pack(side=tk.LEFT, padx=5)
        
        tk.Radiobutton(play_order_frame, text="随机播试", variable=play_order_var, value="random",
                      bg='#E8E8E8', font=('Microsoft YaHei', 9)).pack(side=tk.LEFT, padx=10)
        tk.Label(play_order_frame, text="i", fg='blue', bg='#E8E8E8', cursor='hand2').pack(side=tk.LEFT, padx=5)
        
        # 音量设置
        volume_frame = tk.Frame(content_frame, bg='#E8E8E8')
        volume_frame.grid(row=4, column=1, columnspan=3, sticky='w', padx=5, pady=8)
        
        tk.Label(volume_frame, text="音量:", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT)
        volume_entry = tk.Entry(volume_frame, font=('Microsoft YaHei', 10), width=10)
        volume_entry.insert(0, "80")
        volume_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(volume_frame, text="25-100", font=('Microsoft YaHei', 9), bg='#E8E8E8').pack(side=tk.LEFT, padx=5)
        tk.Label(volume_frame, text="音量指南", font=('Microsoft YaHei', 9), fg='blue', 
                bg='#E8E8E8', cursor='hand2').pack(side=tk.LEFT, padx=10)
        
        # ========== 时间区域 ==========
        time_frame = tk.LabelFrame(main_frame, text="时间", font=('Microsoft YaHei', 12, 'bold'),
                                   bg='#E8E8E8', padx=15, pady=15)
        time_frame.grid(row=1, column=0, sticky='ew', pady=10)
        
        # 开始时间
        tk.Label(time_frame, text="开始时间:", font=('Microsoft YaHei', 9), bg='#E8E8E8').grid(
            row=0, column=0, sticky='e', padx=5, pady=5)
        start_time_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 9), width=50)
        start_time_entry.insert(0, "22:10:10")
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        tk.Label(time_frame, text="《可多个:如中铃》", font=('Microsoft YaHei', 8), bg='#E8E8E8').grid(
            row=0, column=2, sticky='w', padx=5)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry),
                 bg='#D0D0D0', font=('Microsoft YaHei', 8), bd=1, padx=12, pady=2).grid(row=0, column=3, padx=5)d(row=0, column=1, sticky='ew', padx=5, pady=8)
        tk.Label(time_frame, text="《可多个:如中铃》", font=('Microsoft YaHei', 9), bg='#E8E8E8').grid(
            row=0, column=2, sticky='w', padx=5)
        tk.Button(time_frame, text="设置...", bg='#D0D0D0', font=('Microsoft YaHei', 9), 
                 bd=1, padx=15, pady=3).grid(row=0, column=3, padx=5)
        
        # 间隔播报
        interval_var = tk.StringVar(value="first")
        
        interval_frame1 = tk.Frame(time_frame, bg='#E8E8E8')
        interval_frame1.grid(row=1, column=1, columnspan=2, sticky='w', padx=5, pady=5)
        
        tk.Label(time_frame, text="间隔播报:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=1, column=0, sticky='e', padx=5, pady=5)
        tk.Radiobutton(interval_frame1, text="播 n 首", variable=interval_var, value="first",
                      bg='#E8E8E8', font=('Microsoft YaHei', 9)).pack(side=tk.LEFT)
        interval_first_entry = tk.Entry(interval_frame1, font=('Microsoft YaHei', 9), width=15)
        interval_first_entry.insert(0, "1")
        interval_first_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(interval_frame1, text="(单曲时,指 n 遍)", font=('Microsoft YaHei', 9), 
                bg='#E8E8E8').pack(side=tk.LEFT, padx=5)
        
        interval_frame2 = tk.Frame(time_frame, bg='#E8E8E8')
        interval_frame2.grid(row=2, column=1, columnspan=2, sticky='w', padx=5, pady=5)
        
        tk.Radiobutton(interval_frame2, text="播 n 秒", variable=interval_var, value="seconds",
                      bg='#E8E8E8', font=('Microsoft YaHei', 9)).pack(side=tk.LEFT)
        interval_seconds_entry = tk.Entry(interval_frame2, font=('Microsoft YaHei', 9), width=15)
        interval_seconds_entry.insert(0, "600")
        interval_seconds_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(interval_frame2, text="(3600秒 = 1小时)", font=('Microsoft YaHei', 9), 
                bg='#E8E8E8').pack(side=tk.LEFT, padx=5)
        
        # 周几/几号
        tk.Label(time_frame, text="周几/几号:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=3, column=0, sticky='e', padx=5, pady=8)
        weekday_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 10), width=50)
        weekday_entry.insert(0, "每周:1234567")
        weekday_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=8)
        tk.Button(time_frame, text="选取...", bg='#D0D0D0', font=('Microsoft YaHei', 9), 
                 bd=1, padx=15, pady=3).grid(row=3, column=3, padx=5)
        
        # 日期范围
        tk.Label(time_frame, text="日期范围:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=4, column=0, sticky='e', padx=5, pady=8)
        date_range_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 10), width=50)
        date_range_entry.insert(0, "2000-01-01 00:00 ~ 2088-12-31 24:00")
        date_range_entry.grid(row=4, column=1, sticky='ew', padx=5, pady=8)
        tk.Button(time_frame, text="设置...", bg='#D0D0D0', font=('Microsoft YaHei', 9), 
                 bd=1, padx=15, pady=3).grid(row=4, column=3, padx=5)
        
        # 一次性播报
        onetime_label = tk.Label(time_frame, text="一次性播报：将节目的日期范围限定为某天",
                                font=('Microsoft YaHei', 9), bg='#E8E8E8', fg='#666')
        onetime_label.grid(row=5, column=1, columnspan=2, sticky='w', padx=5, pady=5)
        
        # ========== 其它区域 ==========
        other_frame = tk.LabelFrame(main_frame, text="其它", font=('Microsoft YaHei', 11, 'bold'),
                                    bg='#E8E8E8', padx=10, pady=10)
        other_frame.grid(row=2, column=0, sticky='ew', pady=5)
        
        # 准时/延后
        delay_var = tk.StringVar(value="ontime")
        
        tk.Label(other_frame, text="准时/延后:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=0, column=0, sticky='e', padx=5, pady=5)
        
        delay_frame = tk.Frame(other_frame, bg='#E8E8E8')
        delay_frame.grid(row=0, column=1, sticky='w', padx=5, pady=5)
        
        tk.Radiobutton(delay_frame, text="准时播 - 如果有别的节目正在播，终止他们（默认）",
                      variable=delay_var, value="ontime", bg='#E8E8E8', 
                      font=('Microsoft YaHei', 9)).pack(anchor='w')
        tk.Radiobutton(delay_frame, text="可延后 - 如果有别的节目正在播，排队等候",
                      variable=delay_var, value="delay", bg='#E8E8E8', 
                      font=('Microsoft YaHei', 9)).pack(anchor='w')
        
        # ========== 底部按钮 ==========
        button_frame = tk.Frame(main_frame, bg='#E8E8E8')
        button_frame.grid(row=3, column=0, pady=20)
        
        def save_audio_task():
            audio_type = audio_type_var.get()
            if audio_type == "single":
                audio_path = audio_single_entry.get().strip()
            else:
                audio_path = audio_folder_entry.get().strip()
            
            if not audio_path:
                messagebox.showwarning("警告", "请选择音频文件或文件夹")
                return
            
            task = {
                'name': name_entry.get().strip(),
                'time': start_time_entry.get().strip(),
                'content': audio_path,
                'type': 'audio',
                'audio_type': audio_type,
                'play_order': play_order_var.get(),
                'volume': volume_entry.get().strip() or "80",
                'interval_type': interval_var.get(),
                'interval_first': interval_first_entry.get().strip(),
                'interval_seconds': interval_seconds_entry.get().strip(),
                'weekday': weekday_entry.get().strip(),
                'date_range': date_range_entry.get().strip(),
                'delay': delay_var.get(),
                'status': '启用',
                'last_run': None
            }
            
            if not task['name'] or not task['time']:
                messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）")
                return
            
            self.tasks.append(task)
            self.update_task_list()
            self.save_tasks()
            self.log(f"已添加音频节目: {task['name']} - {task['time']}")
            dialog.destroy()
        
        tk.Button(button_frame, text="添加", command=save_audio_task, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 10, 'bold'), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 10), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        
        # 配置列权重
        content_frame.columnconfigure(1, weight=1)
        time_frame.columnconfigure(1, weight=1)
    
    def open_voice_dialog(self, parent_dialog):
        """打开语音节目添加对话框"""
        parent_dialog.destroy()
        
        dialog = tk.Toplevel(self.root)
        dialog.title("定时广播频道 - 语音节目")
        dialog.geometry("800x600")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg='#E8E8E8')
        
        main_frame = tk.Frame(dialog, bg='#E8E8E8', padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ========== 内容区域 ==========
        content_frame = tk.LabelFrame(main_frame, text="内容", font=('Microsoft YaHei', 11, 'bold'),
                                     bg='#E8E8E8', padx=10, pady=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=5)
        
        # 节目名称
        tk.Label(content_frame, text="节目名称:", font=('Microsoft YaHei', 9), bg='#E8E8E8').grid(
            row=0, column=0, sticky='e', padx=5, pady=5)
        name_entry = tk.Entry(content_frame, font=('Microsoft YaHei', 9), width=65)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        
        # 播音文字
        tk.Label(content_frame, text="播音文字:", font=('Microsoft YaHei', 9), bg='#E8E8E8').grid(
            row=1, column=0, sticky='ne', padx=5, pady=5)
        
        text_frame = tk.Frame(content_frame, bg='#E8E8E8')
        text_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        
        content_text = scrolledtext.ScrolledText(text_frame, height=5, font=('Microsoft YaHei', 9), width=65, wrap=tk.WORD)
        content_text.pack(fill=tk.BOTH, expand=True)
        
        # 提示文字
        hint_text = "请正确使用中文标点符号，而不是乱用空格、英文标点；专有名词请加引号；电话号码可用空格间隔。"
        content_text.insert('1.0', hint_text)
        content_text.config(fg='#999')
        
        def on_focus_in(event):
            if content_text.get('1.0', tk.END).strip() == hint_text:
                content_text.delete('1.0', tk.END)
                content_text.config(fg='black')
        
        def on_focus_out(event):
            if not content_text.get('1.0', tk.END).strip():
                content_text.insert('1.0', hint_text)
                content_text.config(fg='#999')
        
        content_text.bind('<FocusIn>', on_focus_in)
        content_text.bind('<FocusOut>', on_focus_out)
        
        # 播音员选择
        tk.Label(content_frame, text="播音员:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=2, column=0, sticky='e', padx=5, pady=8)
        
        voice_frame = tk.Frame(content_frame, bg='#E8E8E8')
        voice_frame.grid(row=2, column=1, columnspan=3, sticky='w', padx=5, pady=8)
        
        voice_var = tk.StringVar(value="六星-六陆女")
        voice_combo = ttk.Combobox(voice_frame, textvariable=voice_var, 
                                   values=["六星-六陆女 (系统>Win7,注网)"], 
                                   font=('Microsoft YaHei', 9), width=35, state='readonly')
        voice_combo.pack(side=tk.LEFT)
        
        # 提示音复选框和选择
        prompt_var = tk.IntVar(value=1)
        prompt_check = tk.Checkbutton(voice_frame, text="提示音", variable=prompt_var, bg='#E8E8E8',
                      font=('Microsoft YaHei', 9))
        prompt_check.pack(side=tk.LEFT, padx=20)
        
        # 音量、语速、音高设置
        settings_frame = tk.Frame(content_frame, bg='#E8E8E8')
        settings_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=5)
        
        # 音量
        tk.Label(settings_frame, text="音  量:", font=('Microsoft YaHei', 9), bg='#E8E8E8').grid(
            row=0, column=0, sticky='e', padx=5, pady=3)
        volume_entry = tk.Entry(settings_frame, font=('Microsoft YaHei', 9), width=12)
        volume_entry.insert(0, "125")
        volume_entry.grid(row=0, column=1, padx=5, pady=3)
        tk.Label(settings_frame, text="50-150,默认125,过大可能破音", font=('Microsoft YaHei', 8), 
                bg='#E8E8E8', fg='#666').grid(row=0, column=2, sticky='w', padx=5)
        
        # 提示音文件选择
        prompt_file_var = tk.StringVar(value="tone-b.mp3")
        prompt_volume_var = tk.StringVar(value="100")
        
        prompt_file_frame = tk.Frame(settings_frame, bg='#E8E8E8')
        prompt_file_frame.grid(row=0, column=3, columnspan=2, sticky='w', padx=10)
        
        prompt_file_entry = tk.Entry(prompt_file_frame, textvariable=prompt_file_var,
                                     font=('Microsoft YaHei', 8), width=15, state='readonly')
        prompt_file_entry.pack(side=tk.LEFT)
        
        tk.Label(prompt_file_frame, text=", 音量", font=('Microsoft YaHei', 8), 
                bg='#E8E8E8').pack(side=tk.LEFT, padx=2)
        
        prompt_volume_entry = tk.Entry(prompt_file_frame, textvariable=prompt_volume_var,
                                       font=('Microsoft YaHei', 8), width=5)
        prompt_volume_entry.pack(side=tk.LEFT, padx=2)
        
        def select_prompt_file():
            # 从提示音文件夹选择
            prompt_dir = "提示音"
            if not os.path.exists(prompt_dir):
                os.makedirs(prompt_dir)
            
            filename = filedialog.askopenfilename(
                title="选择提示音文件",
                initialdir=prompt_dir,
                filetypes=[("音频文件", "*.mp3 *.wav *.ogg"), ("所有文件", "*.*")]
            )
            if filename:
                # 只保存文件名，相对于提示音文件夹
                basename = os.path.basename(filename)
                prompt_file_var.set(basename)
        
        tk.Button(prompt_file_frame, text="...", command=select_prompt_file, bg='#D0D0D0', 
                 font=('Microsoft YaHei', 8), bd=1, padx=8, pady=1).pack(side=tk.LEFT, padx=2)
        
        # 语速
        tk.Label(settings_frame, text="语  速:", font=('Microsoft YaHei', 9), bg='#E8E8E8').grid(
            row=1, column=0, sticky='e', padx=5, pady=3)
        speed_entry = tk.Entry(settings_frame, font=('Microsoft YaHei', 9), width=12)
        speed_entry.insert(0, "-1")
        speed_entry.grid(row=1, column=1, padx=5, pady=3)
        tk.Label(settings_frame, text="[-8,2],默认-1,激励>0.舒缓<0", font=('Microsoft YaHei', 8), 
                bg='#E8E8E8', fg='#666').grid(row=1, column=2, sticky='w', padx=5)
        
        # 音高
        tk.Label(settings_frame, text="音  高:", font=('Microsoft YaHei', 9), bg='#E8E8E8').grid(
            row=2, column=0, sticky='e', padx=5, pady=3)
        pitch_entry = tk.Entry(settings_frame, font=('Microsoft YaHei', 9), width=12)
        pitch_entry.insert(0, "0")
        pitch_entry.grid(row=2, column=1, padx=5, pady=3)
        tk.Label(settings_frame, text="[-2,2],默认0. 变尖>0.变沉<0", font=('Microsoft YaHei', 8), 
                bg='#E8E8E8', fg='#666').grid(row=2, column=2, sticky='w', padx=5)
        
        # ========== 时间区域 ==========
        time_frame = tk.LabelFrame(main_frame, text="时间", font=('Microsoft YaHei', 11, 'bold'),
                                   bg='#E8E8E8', padx=10, pady=10)
        time_frame.grid(row=1, column=0, sticky='ew', pady=5)
        
        # 开始时间
        tk.Label(time_frame, text="开始时间:", font=('Microsoft YaHei', 9), bg='#E8E8E8').grid(
            row=0, column=0, sticky='e', padx=5, pady=5)
        start_time_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 9), width=50)
        start_time_entry.insert(0, "22:10:10")
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        tk.Label(time_frame, text="《可多个:如促销》", font=('Microsoft YaHei', 8), bg='#E8E8E8').grid(
            row=0, column=2, sticky='w', padx=5)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry),
                 bg='#D0D0D0', font=('Microsoft YaHei', 8), bd=1, padx=12, pady=2).grid(row=0, column=3, padx=5)
        
        # 播 n 遍
        tk.Label(time_frame, text="播 n 遍:", font=('Microsoft YaHei', 9), bg='#E8E8E8').grid(
            row=1, column=0, sticky='e', padx=5, pady=5)
        repeat_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 9), width=12)
        repeat_entry.insert(0, "1")
        repeat_entry.grid(row=1, column=1, sticky='w', padx=5, pady=5)
        
        # 周几/几号
        tk.Label(time_frame, text="周几/几号:", font=('Microsoft YaHei', 9), bg='#E8E8E8').grid(
            row=2, column=0, sticky='e', padx=5, pady=5)
        weekday_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 9), width=50)
        weekday_entry.insert(0, "每周:1234567")
        weekday_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=5)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_weekday_settings_dialog(weekday_entry),
                 bg='#D0D0D0', font=('Microsoft YaHei', 8), bd=1, padx=12, pady=2).grid(row=2, column=3, padx=5)
        
        # 日期范围
        tk.Label(time_frame, text="日期范围:", font=('Microsoft YaHei', 9), bg='#E8E8E8').grid(
            row=3, column=0, sticky='e', padx=5, pady=5)
        date_range_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 9), width=50)
        date_range_entry.insert(0, "2000-01-01 00:00 ~ 2088-12-31 24:00")
        date_range_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=5)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry),
                 bg='#D0D0D0', font=('Microsoft YaHei', 8), bd=1, padx=12, pady=2).grid(row=3, column=3, padx=5)
        
        # 一次性播报
        onetime_label = tk.Label(time_frame, text="一次性播报：将节目的日期范围限定为某天",
                                font=('Microsoft YaHei', 8), bg='#E8E8E8', fg='#666')
        onetime_label.grid(row=4, column=1, columnspan=2, sticky='w', padx=5, pady=3)
        
        # ========== 其它区域 ==========
        other_frame = tk.LabelFrame(main_frame, text="其它", font=('Microsoft YaHei', 12, 'bold'),
                                    bg='#E8E8E8', padx=15, pady=15)
        other_frame.grid(row=2, column=0, sticky='ew', pady=10)
        
        # 准时/延后
        delay_var = tk.StringVar(value="delay")
        
        tk.Label(other_frame, text="准时/延后:", font=('Microsoft YaHei', 9), bg='#E8E8E8').grid(
            row=0, column=0, sticky='e', padx=5, pady=3)
        
        delay_frame = tk.Frame(other_frame, bg='#E8E8E8')
        delay_frame.grid(row=0, column=1, sticky='w', padx=5, pady=3)
        
        tk.Radiobutton(delay_frame, text="准时播 - 频道内,如果有别的节目正在播，终止他们",
                      variable=delay_var, value="ontime", bg='#E8E8E8', 
                      font=('Microsoft YaHei', 8)).pack(anchor='w', pady=2)
        tk.Radiobutton(delay_frame, text="可延后 - 频道内,如果有别的节目正在播，排队等候  《促销/禁烟等》",
                      variable=delay_var, value="delay", bg='#E8E8E8', 
                      font=('Microsoft YaHei', 8)).pack(anchor='w', pady=2)
        
        # ========== 底部按钮 ==========
        button_frame = tk.Frame(main_frame, bg='#E8E8E8')
        button_frame.grid(row=3, column=0, pady=20)
        
        def save_voice_task():
            content = content_text.get('1.0', tk.END).strip()
            if content == hint_text or not content:
                messagebox.showwarning("警告", "请输入播音文字内容")
                return
            
            task = {
                'name': name_entry.get().strip(),
                'time': start_time_entry.get().strip(),
                'content': content,
                'type': 'voice',
                'voice': voice_var.get(),
                'prompt': prompt_var.get(),
                'prompt_file': prompt_file_var.get(),
                'prompt_volume': prompt_volume_var.get(),
                'volume': volume_entry.get().strip() or "125",
                'speed': speed_entry.get().strip() or "-1",
                'pitch': pitch_entry.get().strip() or "0",
                'repeat': repeat_entry.get().strip() or "1",
                'weekday': weekday_entry.get().strip(),
                'date_range': date_range_entry.get().strip(),
                'delay': delay_var.get(),
                'status': '启用',
                'last_run': None
            }
            
            if not task['name'] or not task['time']:
                messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）")
                return
            
            self.tasks.append(task)
            self.update_task_list()
            self.save_tasks()
            self.log(f"已添加语音节目: {task['name']} - {task['time']}")
            dialog.destroy()
        
        tk.Button(button_frame, text="确定", command=save_voice_task, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 10, 'bold'), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 10), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        
        # 配置列权重
        content_frame.columnconfigure(1, weight=1)
        time_frame.columnconfigure(1, weight=1)
    
    def delete_task(self):
        """删除任务"""
        selection = self.task_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要删除的节目")
            return
        
        if messagebox.askyesno("确认", "确定要删除选中的节目吗？"):
            index = self.task_tree.index(selection[0])
            task = self.tasks[index]
            self.tasks.pop(index)
            self.update_task_list()
            self.save_tasks()
            self.log(f"已删除节目: {task['name']}")
    
    def edit_task(self):
        """编辑任务"""
        selection = self.task_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要修改的节目")
            return
        
        index = self.task_tree.index(selection[0])
        task = self.tasks[index]
        
        # 根据任务类型打开对应的编辑对话框
        if task.get('type') == 'audio':
            self.edit_audio_task(index, task)
        else:
            self.edit_voice_task(index, task)
    
    def edit_audio_task(self, index, task):
        """编辑音频任务"""
        # 这里可以复用添加对话框的代码，并预填充数据
        self.log(f"编辑音频节目: {task['name']}")
        messagebox.showinfo("提示", "编辑功能开发中，请先删除后重新添加")
    
    def edit_voice_task(self, index, task):
        """编辑语音任务"""
        # 这里可以复用添加对话框的代码，并预填充数据
        self.log(f"编辑语音节目: {task['name']}")
        messagebox.showinfo("提示", "编辑功能开发中，请先删除后重新添加")
    
    def copy_task(self):
        """复制任务"""
        selection = self.task_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要复制的节目")
            return
        
        index = self.task_tree.index(selection[0])
        task = self.tasks[index].copy()
        task['name'] = task['name'] + " (副本)"
        self.tasks.append(task)
        self.update_task_list()
        self.save_tasks()
        self.log(f"已复制节目: {task['name']}")
    
    def move_task(self, direction):
        """移动任务"""
        selection = self.task_tree.selection()
        if not selection:
            return
        
        index = self.task_tree.index(selection[0])
        new_index = index + direction
        
        if 0 <= new_index < len(self.tasks):
            self.tasks[index], self.tasks[new_index] = self.tasks[new_index], self.tasks[index]
            self.update_task_list()
            self.save_tasks()
            # 重新选中
            items = self.task_tree.get_children()
            self.task_tree.selection_set(items[new_index])
    
    def import_tasks(self):
        """导入任务"""
        filename = filedialog.askopenfilename(filetypes=[("JSON文件", "*.json")])
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    imported = json.load(f)
                self.tasks.extend(imported)
                self.update_task_list()
                self.save_tasks()
                self.log(f"已导入 {len(imported)} 个节目")
            except Exception as e:
                messagebox.showerror("错误", f"导入失败: {str(e)}")
    
    def export_tasks(self):
        """导出任务"""
        filename = filedialog.asksaveasfilename(defaultextension=".json",
                                               filetypes=[("JSON文件", "*.json")])
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(self.tasks, f, ensure_ascii=False, indent=2)
                self.log(f"已导出 {len(self.tasks)} 个节目")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}")
    
    def enable_task(self):
        """启用任务"""
        selection = self.task_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要启用的节目")
            return
        
        for item in selection:
            index = self.task_tree.index(item)
            self.tasks[index]['status'] = '启用'
        
        self.update_task_list()
        self.save_tasks()
        self.log(f"已启用 {len(selection)} 个节目")
    
    def disable_task(self):
        """禁用任务"""
        selection = self.task_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要禁用的节目")
            return
        
        for item in selection:
            index = self.task_tree.index(item)
            self.tasks[index]['status'] = '禁用'
        
        self.update_task_list()
        self.save_tasks()
        self.log(f"已禁用 {len(selection)} 个节目")
    
    def show_time_settings_dialog(self, time_entry):
        """显示开始时间设置对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("开始时间")
        dialog.geometry("450x400")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg='#D7F3F5')
        
        main_frame = tk.Frame(dialog, bg='#D7F3F5', padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 24小时制
        tk.Label(main_frame, text="24小时制！", font=('Microsoft YaHei', 10, 'bold'),
                bg='#D7F3F5').pack(anchor='w', pady=5)
        
        # 单个开始时间
        single_frame = tk.Frame(main_frame, bg='#D7F3F5')
        single_frame.pack(anchor='w', pady=10)
        
        time_type_var = tk.StringVar(value="single")
        tk.Radiobutton(single_frame, text="僅1个开始时间", variable=time_type_var, value="single",
                      bg='#D7F3F5', font=('Microsoft YaHei', 9)).pack(side=tk.LEFT)
        
        single_time_entry = tk.Entry(single_frame, font=('Microsoft YaHei', 9), width=15)
        single_time_entry.insert(0, "22:10:10")
        single_time_entry.pack(side=tk.LEFT, padx=10)
        
        tk.Button(single_frame, text="确定", bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 8), bd=1, padx=10, pady=2).pack(side=tk.LEFT, padx=5)
        tk.Button(single_frame, text="取消", bg='#D0D0D0',
                 font=('Microsoft YaHei', 8), bd=1, padx=10, pady=2).pack(side=tk.LEFT)
        
        # 多个开始时间
        multi_frame = tk.Frame(main_frame, bg='#D7F3F5')
        multi_frame.pack(anchor='w', pady=10, fill=tk.BOTH, expand=True)
        
        tk.Radiobutton(multi_frame, text="多个开始时间\n(例如:促销)", variable=time_type_var, value="multi",
                      bg='#D7F3F5', font=('Microsoft YaHei', 9)).pack(anchor='w')
        
        # 时间列表
        time_list_frame = tk.Frame(multi_frame, bg='white', relief=tk.SUNKEN, bd=1)
        time_list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        time_listbox = tk.Listbox(time_list_frame, font=('Microsoft YaHei', 9), height=10)
        time_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(time_list_frame, orient=tk.VERTICAL, command=time_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        time_listbox.configure(yscrollcommand=scrollbar.set)
        
        # 右侧按钮
        btn_frame = tk.Frame(multi_frame, bg='#D7F3F5')
        btn_frame.pack(side=tk.RIGHT, padx=5)
        
        new_time_entry = tk.Entry(btn_frame, font=('Microsoft YaHei', 9), width=12)
        new_time_entry.insert(0, "13:53:25")
        new_time_entry.pack(pady=3)
        
        def add_time():
            time_val = new_time_entry.get().strip()
            if time_val:
                time_listbox.insert(tk.END, time_val)
        
        tk.Button(btn_frame, text="添加", bg='#D0D0D0', font=('Microsoft YaHei', 8),
                 bd=1, padx=15, pady=2, command=add_time).pack(pady=3)
        tk.Button(btn_frame, text="批量生成...", bg='#D0D0D0', font=('Microsoft YaHei', 8),
                 bd=1, padx=10, pady=2).pack(pady=3)
        tk.Button(btn_frame, text="删除", bg='#D0D0D0', font=('Microsoft YaHei', 8),
                 bd=1, padx=15, pady=2).pack(pady=3)
        tk.Button(btn_frame, text="清空", bg='#D0D0D0', font=('Microsoft YaHei', 8),
                 bd=1, padx=15, pady=2).pack(pady=3)
        
        # 底部按钮
        bottom_frame = tk.Frame(main_frame, bg='#D7F3F5')
        bottom_frame.pack(pady=10)
        
        def confirm_time():
            if time_type_var.get() == "single":
                time_entry.delete(0, tk.END)
                time_entry.insert(0, single_time_entry.get())
            else:
                times = [time_listbox.get(i) for i in range(time_listbox.size())]
                time_entry.delete(0, tk.END)
                time_entry.insert(0, ", ".join(times) if times else "")
            dialog.destroy()
        
        tk.Button(bottom_frame, text="确定", command=confirm_time, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 9, 'bold'), bd=1, padx=25, pady=5).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 9), bd=1, padx=25, pady=5).pack(side=tk.LEFT, padx=5)
    
    def show_weekday_settings_dialog(self, weekday_entry):
        """显示周几/几号设置对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("周几或几号")
        dialog.geometry("500x450")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg='#D7F3F5')
        
        main_frame = tk.Frame(dialog, bg='#D7F3F5', padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 周几选择
        week_frame = tk.LabelFrame(main_frame, text="周几", font=('Microsoft YaHei', 10, 'bold'),
                                  bg='#D7F3F5', padx=10, pady=10)
        week_frame.pack(fill=tk.X, pady=5)
        
        week_type_var = tk.StringVar(value="week")
        tk.Radiobutton(week_frame, text="周几", variable=week_type_var, value="week",
                      bg='#D7F3F5', font=('Microsoft YaHei', 9)).grid(row=0, column=0, sticky='w')
        
        week_vars = {}
        weekdays = [("周一", 1), ("周二", 2), ("周三", 3), ("周四", 4),
                   ("周五", 5), ("周六", 6), ("周日", 7)]
        
        for i, (day, num) in enumerate(weekdays):
            var = tk.IntVar(value=1)
            week_vars[num] = var
            row = (i // 4) + 1
            col = i % 4
            tk.Checkbutton(week_frame, text=day, variable=var, bg='#D7F3F5',
                          font=('Microsoft YaHei', 9)).grid(row=row, column=col, sticky='w', padx=10, pady=3)
        
        # 几号选择
        day_frame = tk.LabelFrame(main_frame, text="几号", font=('Microsoft YaHei', 10, 'bold'),
                                 bg='#D7F3F5', padx=10, pady=10)
        day_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        tk.Radiobutton(day_frame, text="几号", variable=week_type_var, value="day",
                      bg='#D7F3F5', font=('Microsoft YaHei', 9)).grid(row=0, column=0, sticky='w')
        
        day_vars = {}
        for i in range(1, 32):
            var = tk.IntVar(value=0)
            day_vars[i] = var
            row = (i - 1) // 6 + 1
            col = (i - 1) % 6
            tk.Checkbutton(day_frame, text=f"{i:02d}", variable=var, bg='#D7F3F5',
                          font=('Microsoft YaHei', 8)).grid(row=row, column=col, sticky='w', padx=8, pady=2)
        
        # 底部按钮
        bottom_frame = tk.Frame(main_frame, bg='#D7F3F5')
        bottom_frame.pack(pady=10)
        
        def confirm_weekday():
            if week_type_var.get() == "week":
                selected = [str(num) for num, var in week_vars.items() if var.get() == 1]
                result = "每周:" + "".join(selected) if selected else ""
            else:
                selected = [f"{num:02d}" for num, var in day_vars.items() if var.get() == 1]
                result = "每月:" + ",".join(selected) if selected else ""
            
            weekday_entry.delete(0, tk.END)
            weekday_entry.insert(0, result)
            dialog.destroy()
        
        tk.Button(bottom_frame, text="确定", command=confirm_weekday, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 9, 'bold'), bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 9), bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)
    
    def show_daterange_settings_dialog(self, date_range_entry):
        """显示日期范围设置对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("日期范围")
        dialog.geometry("450x220")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg='#D7F3F5')
        
        main_frame = tk.Frame(dialog, bg='#D7F3F5', padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 从日期
        from_frame = tk.Frame(main_frame, bg='#D7F3F5')
        from_frame.pack(pady=10)
        
        tk.Label(from_frame, text="从", font=('Microsoft YaHei', 10, 'bold'),
                bg='#D7F3F5').pack(side=tk.LEFT, padx=5)
        
        from_date_entry = tk.Entry(from_frame, font=('Microsoft YaHei', 9), width=15)
        from_date_entry.insert(0, "2000-01-01")
        from_date_entry.pack(side=tk.LEFT, padx=5)
        
        tk.Button(from_frame, text="📅", bg='#D0D0D0', font=('Microsoft YaHei', 9),
                 bd=1, padx=5, pady=2).pack(side=tk.LEFT, padx=2)
        
        from_hour_entry = tk.Entry(from_frame, font=('Microsoft YaHei', 9), width=5)
        from_hour_entry.insert(0, "0")
        from_hour_entry.pack(side=tk.LEFT, padx=2)
        
        tk.Label(from_frame, text="点", font=('Microsoft YaHei', 9),
                bg='#D7F3F5').pack(side=tk.LEFT)
        
        # 到日期
        to_frame = tk.Frame(main_frame, bg='#D7F3F5')
        to_frame.pack(pady=10)
        
        tk.Label(to_frame, text="到", font=('Microsoft YaHei', 10, 'bold'),
                bg='#D7F3F5').pack(side=tk.LEFT, padx=5)
        
        to_date_entry = tk.Entry(to_frame, font=('Microsoft YaHei', 9), width=15)
        to_date_entry.insert(0, "2088-12-31")
        to_date_entry.pack(side=tk.LEFT, padx=5)
        
        tk.Button(to_frame, text="📅", bg='#D0D0D0', font=('Microsoft YaHei', 9),
                 bd=1, padx=5, pady=2).pack(side=tk.LEFT, padx=2)
        
        to_hour_entry = tk.Entry(to_frame, font=('Microsoft YaHei', 9), width=5)
        to_hour_entry.insert(0, "24")
        to_hour_entry.pack(side=tk.LEFT, padx=2)
        
        tk.Label(to_frame, text="点", font=('Microsoft YaHei', 9),
                bg='#D7F3F5').pack(side=tk.LEFT)
        
        tk.Label(to_frame, text="0~24", font=('Microsoft YaHei', 8),
                bg='#D7F3F5', fg='#666').pack(side=tk.LEFT, padx=5)
        
        # 提示
        tk.Label(main_frame, text="比如：某节目仅元旦那天播", font=('Microsoft YaHei', 9),
                bg='#D7F3F5', fg='#666').pack(pady=10)
        
        # 底部按钮
        bottom_frame = tk.Frame(main_frame, bg='#D7F3F5')
        bottom_frame.pack(pady=10)
        
        def confirm_daterange():
            from_date = from_date_entry.get()
            from_hour = from_hour_entry.get()
            to_date = to_date_entry.get()
            to_hour = to_hour_entry.get()
            
            result = f"{from_date} {from_hour.zfill(2)}:00 ~ {to_date} {to_hour.zfill(2)}:00"
            date_range_entry.delete(0, tk.END)
            date_range_entry.insert(0, result)
            dialog.destroy()
        
        tk.Button(bottom_frame, text="确定", command=confirm_daterange, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 9, 'bold'), bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 9), bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)
    
    def update_task_list(self):
        """更新任务列表"""
        for item in self.task_tree.get_children():
            self.task_tree.delete(item)
        
        for task in self.tasks:
            task_type = task.get('type', 'voice')
            if task_type == 'audio':
                # 音频节目显示文件名
                import os
                content_preview = os.path.basename(task['content'])
            else:
                # 语音节目显示文本内容
                content_preview = task['content'][:30] + '...' if len(task['content']) > 30 else task['content']
            
            self.task_tree.insert('', tk.END, values=(
                task['name'],
                task['status'],
                task['time'],
                task['delay'],
                content_preview,
                task['volume'],
                task['weekday'],
                task['date_range']
            ))
        
        # 更新统计
        self.stats_label.config(text=f"节目单：{len(self.tasks)}")
    
    def update_status_bar(self):
        """更新状态栏"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.status_labels[0].config(text=f"当前时间: {current_time}")
        self.status_labels[1].config(text="系统状态: 运行中")
        self.status_labels[2].config(text="播放状态: 待机")
        self.status_labels[3].config(text=f"任务数量: {len(self.tasks)}")
        
        self.root.after(1000, self.update_status_bar)
    
    def start_background_thread(self):
        """启动后台线程"""
        self.running = True
        thread = threading.Thread(target=self._check_tasks, daemon=True)
        thread.start()
    
    def _check_tasks(self):
        """后台检查任务"""
        while self.running:
            now = datetime.now()
            current_time = now.strftime("%H:%M")
            current_date = now.strftime("%Y-%m-%d")
            weekday = now.isoweekday()  # 1=周一, 7=周日
            
            for task in self.tasks:
                if task['status'] != '启用':
                    continue
                
                # 检查时间匹配
                if current_time in task['time']:
                    if task['last_run'] == current_date:
                        continue
                    
                    # 检查星期
                    if str(weekday) in task['weekday'] or '1-7' in task['weekday']:
                        self.root.after(0, self._execute_broadcast, task, current_date)
            
            time.sleep(30)
    
    def _execute_broadcast(self, task, current_date):
        """执行播报"""
        task_type = task.get('type', 'voice')
        
        if task_type == 'audio':
            # 音频文件播放
            import os
            audio_path = task['content']
            
            # 如果是相对路径，转换为绝对路径
            if not os.path.isabs(audio_path):
                audio_path = os.path.abspath(audio_path)
            
            if not os.path.exists(audio_path):
                self.log(f"错误：音频文件不存在 - {audio_path}")
                return
            
            filename = os.path.basename(audio_path)
            self.playing_text.delete('1.0', tk.END)
            self.playing_text.insert('1.0', f"[{task['name']}] 正在播放音频: {filename}")
            self.log(f"播放音频: {task['name']} - {filename}")
            
            if AUDIO_AVAILABLE:
                threading.Thread(target=self._play_audio, args=(audio_path,), daemon=True).start()
            else:
                self.log("错误：pygame未安装，无法播放音频")
        else:
            # 语音播报
            self.playing_text.delete('1.0', tk.END)
            self.playing_text.insert('1.0', f"[{task['name']}] {task['content']}")
            self.log(f"开始播报: {task['name']}")
            
            # 如果启用提示音，先播放提示音
            if task.get('prompt', 0) == 1 and AUDIO_AVAILABLE:
                prompt_file = task.get('prompt_file', 'tone-b.mp3')
                prompt_path = os.path.join("提示音", prompt_file)
                if os.path.exists(prompt_path):
                    self._play_audio(prompt_path, wait=True)
            
            threading.Thread(target=self._speak, args=(task['content'],), daemon=True).start()
        
        task['last_run'] = current_date
        self.save_tasks()
    
    def _play_audio(self, audio_path, wait=False):
        """播放音频文件"""
        try:
            if not AUDIO_AVAILABLE:
                self.log("错误：pygame未安装")
                return
            
            pygame.mixer.music.load(audio_path)
            pygame.mixer.music.play()
            
            if wait:
                # 等待播放完成
                while pygame.mixer.music.get_busy():
                    time.sleep(0.1)
        except Exception as e:
            self.log(f"音频播放错误: {str(e)}")
    
    def _speak(self, text):
        """语音播报"""
        try:
            self.engine.say(text)
            self.engine.runAndWait()
        except Exception as e:
            self.log(f"播报错误: {str(e)}")
    
    def log(self, message):
        """记录日志"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"{timestamp}    {message}\n"
        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)
    
    def save_tasks(self):
        """保存任务"""
        try:
            with open(self.task_file, 'w', encoding='utf-8') as f:
                json.dump(self.tasks, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"保存失败: {str(e)}")
    
    def load_tasks(self):
        """加载任务"""
        if os.path.exists(self.task_file):
            try:
                with open(self.task_file, 'r', encoding='utf-8') as f:
                    self.tasks = json.load(f)
                self.update_task_list()
                self.log(f"已加载 {len(self.tasks)} 个节目")
            except Exception as e:
                self.log(f"加载失败: {str(e)}")

def main():
    root = tk.Tk()
    app = TimedBroadcastApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
