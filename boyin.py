import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import pyttsx3
import json
import threading
import time
from datetime import datetime
import os

# --- 全局设置 ---
# 获取脚本所在的目录，确保所有文件路径都是基于此目录
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TASK_FILE = os.path.join(SCRIPT_DIR, "broadcast_tasks.json")
PROMPT_FOLDER = os.path.join(SCRIPT_DIR, "提示音")
AUDIO_FOLDER = os.path.join(SCRIPT_DIR, "音频文件")

# 音频播放库
AUDIO_AVAILABLE = False
try:
    import pygame
    pygame.mixer.init()
    AUDIO_AVAILABLE = True
except ImportError:
    print("警告: 未安装 pygame 库 (pip install pygame)，音频播放功能将不可用。")
except Exception as e:
    print(f"警告: pygame 初始化失败 - {e}，音频播放功能将不可用。")

class TimedBroadcastApp:
    def __init__(self, root):
        self.root = root
        self.root.title("定时播音")
        self.root.geometry("1400x800")
        self.root.configure(bg='#E8F4F8')

        # 初始化语音引擎
        self.engine = None
        try:
            self.engine = pyttsx3.init()
            self.engine.setProperty('rate', 150)
            self.engine.setProperty('volume', 1.0)
        except Exception as e:
            print(f"错误：pyttsx3 语音引擎初始化失败 - {e}。语音播报功能将不可用。")

        # 任务列表
        self.tasks = []
        self.running = True
        self.task_file = TASK_FILE
        self.current_page = "定时广播"

        # 创建必要的文件夹结构
        self.create_folder_structure()

        # 创建界面
        self.create_widgets()

        # 加载已保存的任务
        self.load_tasks()

        # 启动后台检查线程
        self.start_background_thread()
        
        # 绑定关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_folder_structure(self):
        """创建必要的文件夹结构"""
        folders = [PROMPT_FOLDER, AUDIO_FOLDER]
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
            
            # --- 字体修复 1: 左侧导航栏字体调大 ---
            btn = tk.Button(btn_frame, text=title, bg='#5DADE2' if i == 0 else '#A8D8E8',
                          fg='white' if i == 0 else 'black', font=('Microsoft YaHei', 13, 'bold'),
                          bd=0, padx=10, pady=8, anchor='w',
                          command=lambda t=title: self.switch_page(t))
            btn.pack(fill=tk.X)

            if subtitle:
                # --- 字体修复 2: 副标题字体调大 ---
                sub_label = tk.Label(btn_frame, text=subtitle, bg='#5DADE2' if i == 0 else '#A8D8E8',
                                   fg='#FF6B35' if i == 3 else ('#555' if i == 0 else '#666'),
                                   font=('Microsoft YaHei', 10), anchor='w', padx=10)
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
            messagebox.showinfo("提示", f"页面 [{page_name}] 正在开发中...")
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
            self.task_tree.column(col, width=width, anchor='w' if col in ['节目名称', '音频或文字'] else 'center')

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
                                                     bg='#FFFEF0', wrap=tk.WORD, state='disabled')
        self.playing_text.pack(fill=tk.BOTH, expand=True)
        self.update_playing_text("等待播放...")

        # 日志区域
        log_frame = tk.LabelFrame(self.main_frame, text="日志：", font=('Microsoft YaHei', 10),
                                 bg='white', fg='#2C5F7C', padx=10, pady=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, font=('Microsoft YaHei', 9),
                                                 bg='#F9F9F9', wrap=tk.WORD, state='disabled')
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
        choice_dialog.geometry("350x250")
        choice_dialog.resizable(False, False)
        choice_dialog.transient(self.root)
        choice_dialog.grab_set()

        # 居中显示
        self.center_window(choice_dialog, 350, 250)
        
        main_frame = tk.Frame(choice_dialog, padx=20, pady=20, bg='#F0F0F0')
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_label = tk.Label(main_frame, text="请选择要添加的节目类型",
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
        # --- Bug修复 3: 增加窗口高度以显示按钮 ---
        dialog.geometry("850x680")
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
        tk.Label(content_frame, text="节目名称:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=0, column=0, sticky='e', padx=5, pady=5)
        name_entry = tk.Entry(content_frame, font=('Microsoft YaHei', 10), width=55)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=5)

        # 音频文件单选
        audio_type_var = tk.StringVar(value="single")

        tk.Label(content_frame, text="音频文件", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=1, column=0, sticky='e', padx=5, pady=5)

        audio_single_frame = tk.Frame(content_frame, bg='#E8E8E8')
        audio_single_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=5)

        tk.Radiobutton(audio_single_frame, text="", variable=audio_type_var, value="single",
                      bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT)

        audio_single_entry = tk.Entry(audio_single_frame, font=('Microsoft YaHei', 10), width=35, state='readonly')
        audio_single_entry.pack(side=tk.LEFT, padx=5)

        time_label = tk.Label(audio_single_frame, text="00:00", font=('Microsoft YaHei', 10), bg='#E8E8E8')
        time_label.pack(side=tk.LEFT, padx=10)

        def select_single_audio():
            filename = filedialog.askopenfilename(
                title="选择音频文件",
                initialdir=AUDIO_FOLDER,
                filetypes=[("音频文件", "*.mp3 *.wav *.ogg *.flac"), ("所有文件", "*.*")]
            )
            if filename:
                audio_single_entry.config(state='normal')
                audio_single_entry.delete(0, tk.END)
                audio_single_entry.insert(0, filename)
                audio_single_entry.config(state='readonly')

        tk.Button(audio_single_frame, text="选取...", command=select_single_audio, bg='#D0D0D0',
                 font=('Microsoft YaHei', 10), bd=1, padx=15, pady=3).pack(side=tk.LEFT, padx=5)

        # 音频文件夹单选
        tk.Label(content_frame, text="音频文件夹", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=2, column=0, sticky='e', padx=5, pady=5)

        audio_folder_frame = tk.Frame(content_frame, bg='#E8E8E8')
        audio_folder_frame.grid(row=2, column=1, columnspan=3, sticky='ew', padx=5, pady=5)

        tk.Radiobutton(audio_folder_frame, text="", variable=audio_type_var, value="folder",
                      bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT)

        audio_folder_entry = tk.Entry(audio_folder_frame, font=('Microsoft YaHei', 10), width=50, state='readonly')
        audio_folder_entry.pack(side=tk.LEFT, padx=5)

        def select_folder():
            foldername = filedialog.askdirectory(title="选择音频文件夹", initialdir=AUDIO_FOLDER)
            if foldername:
                audio_folder_entry.config(state='normal')
                audio_folder_entry.delete(0, tk.END)
                audio_folder_entry.insert(0, foldername)
                audio_folder_entry.config(state='readonly')

        tk.Button(audio_folder_frame, text="选取...", command=select_folder, bg='#D0D0D0',
                 font=('Microsoft YaHei', 10), bd=1, padx=15, pady=3).pack(side=tk.LEFT, padx=5)

        # 播放顺序
        play_order_frame = tk.Frame(content_frame, bg='#E8E8E8')
        play_order_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=5)

        play_order_var = tk.StringVar(value="sequential")
        tk.Radiobutton(play_order_frame, text="顺序播", variable=play_order_var, value="sequential",
                      bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT, padx=10)
        
        tk.Radiobutton(play_order_frame, text="随机播", variable=play_order_var, value="random",
                      bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT, padx=10)

        # 音量设置
        volume_frame = tk.Frame(content_frame, bg='#E8E8E8')
        volume_frame.grid(row=4, column=1, columnspan=3, sticky='w', padx=5, pady=8)

        tk.Label(volume_frame, text="音量:", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT)
        volume_entry = tk.Entry(volume_frame, font=('Microsoft YaHei', 10), width=10)
        volume_entry.insert(0, "80")
        volume_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(volume_frame, text="0-100", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT, padx=5)

        # ========== 时间区域 ==========
        time_frame = tk.LabelFrame(main_frame, text="时间", font=('Microsoft YaHei', 12, 'bold'),
                                   bg='#E8E8E8', padx=15, pady=15)
        time_frame.grid(row=1, column=0, sticky='ew', pady=10)

        # 开始时间
        tk.Label(time_frame, text="开始时间:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=0, column=0, sticky='e', padx=5, pady=5)
        start_time_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 10), width=50)
        start_time_entry.insert(0, "22:10:10")
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        tk.Label(time_frame, text="《可多个,用英文逗号,隔开》", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=0, column=2, sticky='w', padx=5)
        
        tk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry),
                 bg='#D0D0D0', font=('Microsoft YaHei', 10), bd=1, padx=12, pady=2).grid(row=0, column=3, padx=5)

        # 间隔播报
        interval_var = tk.StringVar(value="first")

        interval_frame1 = tk.Frame(time_frame, bg='#E8E8E8')
        interval_frame1.grid(row=1, column=1, columnspan=2, sticky='w', padx=5, pady=5)

        tk.Label(time_frame, text="间隔播报:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=1, column=0, sticky='e', padx=5, pady=5)
        tk.Radiobutton(interval_frame1, text="播 n 首", variable=interval_var, value="first",
                      bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT)
        interval_first_entry = tk.Entry(interval_frame1, font=('Microsoft YaHei', 10), width=15)
        interval_first_entry.insert(0, "1")
        interval_first_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(interval_frame1, text="(单曲时,指 n 遍)", font=('Microsoft YaHei', 10),
                bg='#E8E8E8').pack(side=tk.LEFT, padx=5)

        interval_frame2 = tk.Frame(time_frame, bg='#E8E8E8')
        interval_frame2.grid(row=2, column=1, columnspan=2, sticky='w', padx=5, pady=5)

        tk.Radiobutton(interval_frame2, text="播 n 秒", variable=interval_var, value="seconds",
                      bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT)
        interval_seconds_entry = tk.Entry(interval_frame2, font=('Microsoft YaHei', 10), width=15)
        interval_seconds_entry.insert(0, "600")
        interval_seconds_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(interval_frame2, text="(3600秒 = 1小时)", font=('Microsoft YaHei', 10),
                bg='#E8E8E8').pack(side=tk.LEFT, padx=5)

        # 周几/几号
        tk.Label(time_frame, text="周几/几号:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=3, column=0, sticky='e', padx=5, pady=8)
        weekday_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 10), width=50)
        weekday_entry.insert(0, "每周:1234567")
        weekday_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=8)
        tk.Button(time_frame, text="选取...", command=lambda: self.show_weekday_settings_dialog(weekday_entry),
                 bg='#D0D0D0', font=('Microsoft YaHei', 10),
                 bd=1, padx=15, pady=3).grid(row=3, column=3, padx=5)

        # 日期范围
        tk.Label(time_frame, text="日期范围:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=4, column=0, sticky='e', padx=5, pady=8)
        date_range_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 10), width=50)
        date_range_entry.insert(0, "2000-01-01 ~ 2099-12-31")
        date_range_entry.grid(row=4, column=1, sticky='ew', padx=5, pady=8)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry),
                 bg='#D0D0D0', font=('Microsoft YaHei', 10),
                 bd=1, padx=15, pady=3).grid(row=4, column=3, padx=5)

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
                      font=('Microsoft YaHei', 10)).pack(anchor='w')
        tk.Radiobutton(delay_frame, text="可延后 - 如果有别的节目正在播，排队等候",
                      variable=delay_var, value="delay", bg='#E8E8E8',
                      font=('Microsoft YaHei', 10)).pack(anchor='w')

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
                messagebox.showwarning("警告", "请选择音频文件或文件夹", parent=dialog)
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
                'last_run': {} # 初始化为字典
            }

            if not task['name'] or not task['time']:
                messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog)
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
        # --- Bug修复 3: 增加窗口高度以显示按钮 ---
        dialog.geometry("800x700")
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
        tk.Label(content_frame, text="节目名称:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=0, column=0, sticky='e', padx=5, pady=5)
        name_entry = tk.Entry(content_frame, font=('Microsoft YaHei', 10), width=65)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=5)

        # 播音文字
        tk.Label(content_frame, text="播音文字:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=1, column=0, sticky='ne', padx=5, pady=5)

        text_frame = tk.Frame(content_frame, bg='#E8E8E8')
        text_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=5)

        content_text = scrolledtext.ScrolledText(text_frame, height=5, font=('Microsoft YaHei', 10), width=65, wrap=tk.WORD)
        content_text.pack(fill=tk.BOTH, expand=True)

        # 提示文字
        hint_text = "请在此输入要播报的文字内容..."
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

        # 获取系统可用语音
        available_voices = []
        if self.engine:
            voices = self.engine.getProperty('voices')
            for voice in voices:
                available_voices.append(voice.name)

        voice_var = tk.StringVar(value=available_voices[0] if available_voices else "")
        voice_combo = ttk.Combobox(voice_frame, textvariable=voice_var,
                                   values=available_voices,
                                   font=('Microsoft YaHei', 10), width=35, state='readonly')
        voice_combo.pack(side=tk.LEFT)

        # 提示音复选框和选择
        prompt_var = tk.IntVar(value=1)
        prompt_check = tk.Checkbutton(voice_frame, text="提示音", variable=prompt_var, bg='#E8E8E8',
                      font=('Microsoft YaHei', 10))
        prompt_check.pack(side=tk.LEFT, padx=20)

        # 音量、语速、音高设置
        settings_frame = tk.Frame(content_frame, bg='#E8E8E8')
        settings_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=5)

        # 音量
        tk.Label(settings_frame, text="音  量:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=0, column=0, sticky='e', padx=5, pady=3)
        volume_entry = tk.Entry(settings_frame, font=('Microsoft YaHei', 10), width=12)
        volume_entry.insert(0, "1.0")
        volume_entry.grid(row=0, column=1, padx=5, pady=3)
        tk.Label(settings_frame, text="0.0-1.0, 默认1.0", font=('Microsoft YaHei', 10),
                bg='#E8E8E8', fg='#666').grid(row=0, column=2, sticky='w', padx=5)

        # 提示音文件选择
        prompt_file_var = tk.StringVar(value="tone-b.mp3")
        prompt_volume_var = tk.StringVar(value="100")

        prompt_file_frame = tk.Frame(settings_frame, bg='#E8E8E8')
        prompt_file_frame.grid(row=0, column=3, columnspan=2, sticky='w', padx=10)

        prompt_file_entry = tk.Entry(prompt_file_frame, textvariable=prompt_file_var,
                                     font=('Microsoft YaHei', 10), width=15, state='readonly')
        prompt_file_entry.pack(side=tk.LEFT)

        tk.Label(prompt_file_frame, text=", 音量", font=('Microsoft YaHei', 10),
                bg='#E8E8E8').pack(side=tk.LEFT, padx=2)

        prompt_volume_entry = tk.Entry(prompt_file_frame, textvariable=prompt_volume_var,
                                       font=('Microsoft YaHei', 10), width=5)
        prompt_volume_entry.pack(side=tk.LEFT, padx=2)

        def select_prompt_file():
            filename = filedialog.askopenfilename(
                title="选择提示音文件",
                initialdir=PROMPT_FOLDER,
                filetypes=[("音频文件", "*.mp3 *.wav *.ogg"), ("所有文件", "*.*")]
            )
            if filename:
                # 只保存文件名
                basename = os.path.basename(filename)
                prompt_file_var.set(basename)

        tk.Button(prompt_file_frame, text="...", command=select_prompt_file, bg='#D0D0D0',
                 font=('Microsoft YaHei', 10), bd=1, padx=8, pady=1).pack(side=tk.LEFT, padx=2)

        # 语速
        tk.Label(settings_frame, text="语  速:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=1, column=0, sticky='e', padx=5, pady=3)
        speed_entry = tk.Entry(settings_frame, font=('Microsoft YaHei', 10), width=12)
        speed_entry.insert(0, "150")
        speed_entry.grid(row=1, column=1, padx=5, pady=3)
        tk.Label(settings_frame, text="默认150, 数字越大越快", font=('Microsoft YaHei', 10),
                bg='#E8E8E8', fg='#666').grid(row=1, column=2, sticky='w', padx=5)

        # ========== 时间区域 ==========
        time_frame = tk.LabelFrame(main_frame, text="时间", font=('Microsoft YaHei', 11, 'bold'),
                                   bg='#E8E8E8', padx=10, pady=10)
        time_frame.grid(row=1, column=0, sticky='ew', pady=5)

        # 开始时间
        tk.Label(time_frame, text="开始时间:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=0, column=0, sticky='e', padx=5, pady=5)
        start_time_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 10), width=50)
        start_time_entry.insert(0, "22:10:10")
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        tk.Label(time_frame, text="《可多个,用英文逗号,隔开》", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=0, column=2, sticky='w', padx=5)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry),
                 bg='#D0D0D0', font=('Microsoft YaHei', 10), bd=1, padx=12, pady=2).grid(row=0, column=3, padx=5)

        # 播 n 遍
        tk.Label(time_frame, text="播 n 遍:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=1, column=0, sticky='e', padx=5, pady=5)
        repeat_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 10), width=12)
        repeat_entry.insert(0, "1")
        repeat_entry.grid(row=1, column=1, sticky='w', padx=5, pady=5)

        # 周几/几号
        tk.Label(time_frame, text="周几/几号:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=2, column=0, sticky='e', padx=5, pady=5)
        weekday_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 10), width=50)
        weekday_entry.insert(0, "每周:1234567")
        weekday_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=5)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_weekday_settings_dialog(weekday_entry),
                 bg='#D0D0D0', font=('Microsoft YaHei', 10), bd=1, padx=12, pady=2).grid(row=2, column=3, padx=5)

        # 日期范围
        tk.Label(time_frame, text="日期范围:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=3, column=0, sticky='e', padx=5, pady=5)
        date_range_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 10), width=50)
        date_range_entry.insert(0, "2000-01-01 ~ 2099-12-31")
        date_range_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=5)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry),
                 bg='#D0D0D0', font=('Microsoft YaHei', 10), bd=1, padx=12, pady=2).grid(row=3, column=3, padx=5)

        # ========== 其它区域 ==========
        other_frame = tk.LabelFrame(main_frame, text="其它", font=('Microsoft YaHei', 12, 'bold'),
                                    bg='#E8E8E8', padx=15, pady=15)
        other_frame.grid(row=2, column=0, sticky='ew', pady=10)

        # 准时/延后
        delay_var = tk.StringVar(value="delay")

        tk.Label(other_frame, text="准时/延后:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=0, column=0, sticky='e', padx=5, pady=3)

        delay_frame = tk.Frame(other_frame, bg='#E8E8E8')
        delay_frame.grid(row=0, column=1, sticky='w', padx=5, pady=3)

        tk.Radiobutton(delay_frame, text="准时播 - 频道内,如果有别的节目正在播，终止他们",
                      variable=delay_var, value="ontime", bg='#E8E8E8',
                      font=('Microsoft YaHei', 10)).pack(anchor='w', pady=2)
        tk.Radiobutton(delay_frame, text="可延后 - 频道内,如果有别的节目正在播，排队等候",
                      variable=delay_var, value="delay", bg='#E8E8E8',
                      font=('Microsoft YaHei', 10)).pack(anchor='w', pady=2)

        # ========== 底部按钮 ==========
        button_frame = tk.Frame(main_frame, bg='#E8E8E8')
        button_frame.grid(row=3, column=0, pady=20)

        def save_voice_task():
            content = content_text.get('1.0', tk.END).strip()
            if content == hint_text or not content:
                messagebox.showwarning("警告", "请输入播音文字内容", parent=dialog)
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
                'volume': volume_entry.get().strip() or "1.0",
                'speed': speed_entry.get().strip() or "150",
                'repeat': repeat_entry.get().strip() or "1",
                'weekday': weekday_entry.get().strip(),
                'date_range': date_range_entry.get().strip(),
                'delay': delay_var.get(),
                'status': '启用',
                'last_run': {} # 初始化为字典
            }

            if not task['name'] or not task['time']:
                messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog)
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
        selections = self.task_tree.selection()
        if not selections:
            messagebox.showwarning("警告", "请先选择要删除的节目")
            return

        if messagebox.askyesno("确认", f"确定要删除选中的 {len(selections)} 个节目吗？"):
            # 从后往前删除，避免索引变化问题
            indices_to_delete = sorted([self.task_tree.index(sel) for sel in selections], reverse=True)
            for index in indices_to_delete:
                task = self.tasks.pop(index)
                self.log(f"已删除节目: {task['name']}")
            self.update_task_list()
            self.save_tasks()

    def edit_task(self):
        """编辑任务"""
        selection = self.task_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要修改的节目")
            return
        if len(selection) > 1:
            messagebox.showwarning("警告", "只能同时修改一个节目")
            return
        
        messagebox.showinfo("提示", "编辑功能开发中，请先删除后重新添加")

    def copy_task(self):
        """复制任务"""
        selections = self.task_tree.selection()
        if not selections:
            messagebox.showwarning("警告", "请先选择要复制的节目")
            return

        for sel in selections:
            index = self.task_tree.index(sel)
            original_task = self.tasks[index]
            
            task_copy = json.loads(json.dumps(original_task))
            
            task_copy['name'] = task_copy['name'] + " (副本)"
            task_copy['last_run'] = {} 
            self.tasks.append(task_copy)
            self.log(f"已复制节目: {original_task['name']}")
        
        self.update_task_list()
        self.save_tasks()

    def move_task(self, direction):
        """移动任务"""
        selection = self.task_tree.selection()
        if not selection or len(selection) > 1:
            return
        
        item = selection[0]
        index = self.task_tree.index(item)
        new_index = index + direction
        
        if 0 <= new_index < len(self.tasks):
            self.tasks.insert(new_index, self.tasks.pop(index))
            self.update_task_list()
            self.save_tasks()
            
            all_items = self.task_tree.get_children()
            if all_items:
                self.task_tree.selection_set(all_items[new_index])

    def import_tasks(self):
        """导入任务"""
        filename = filedialog.askopenfilename(
            title="选择要导入的节目文件",
            filetypes=[("JSON文件", "*.json")])
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    imported = json.load(f)
                self.tasks.extend(imported)
                self.update_task_list()
                self.save_tasks()
                self.log(f"已从 {os.path.basename(filename)} 导入 {len(imported)} 个节目")
            except Exception as e:
                messagebox.showerror("错误", f"导入失败: {str(e)}")

    def export_tasks(self):
        """导出任务"""
        if not self.tasks:
            messagebox.showwarning("警告", "当前没有节目可以导出")
            return
        
        filename = filedialog.asksaveasfilename(
            title="导出节目到...",
            defaultextension=".json",
            initialfile="broadcast_backup.json",
            filetypes=[("JSON文件", "*.json")])
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(self.tasks, f, ensure_ascii=False, indent=2)
                self.log(f"已导出 {len(self.tasks)} 个节目到 {os.path.basename(filename)}")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}")

    def enable_task(self):
        """启用任务"""
        self._set_task_status('启用')

    def disable_task(self):
        """禁用任务"""
        self._set_task_status('禁用')

    def _set_task_status(self, status):
        """设置任务状态的通用函数"""
        selection = self.task_tree.selection()
        if not selection:
            messagebox.showwarning("警告", f"请先选择要{status}的节目")
            return
        
        count = 0
        for item in selection:
            index = self.task_tree.index(item)
            if self.tasks[index]['status'] != status:
                self.tasks[index]['status'] = status
                count += 1
        
        if count > 0:
            self.update_task_list()
            self.save_tasks()
            self.log(f"已{status} {count} 个节目")

    def show_time_settings_dialog(self, time_entry):
        """显示开始时间设置对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("开始时间设置")
        dialog.geometry("450x400")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg='#D7F3F5')
        self.center_window(dialog, 450, 400)

        main_frame = tk.Frame(dialog, bg='#D7F3F5', padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(main_frame, text="24小时制，格式为 HH:MM:SS", font=('Microsoft YaHei', 10, 'bold'),
                bg='#D7F3F5').pack(anchor='w', pady=5)
        
        time_list_frame = tk.LabelFrame(main_frame, text="时间列表", bg='#D7F3F5', padx=5, pady=5)
        time_list_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        listbox_frame = tk.Frame(time_list_frame)
        listbox_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        time_listbox = tk.Listbox(listbox_frame, font=('Microsoft YaHei', 10), height=10)
        time_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=time_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        time_listbox.configure(yscrollcommand=scrollbar.set)
        
        current_times = [t.strip() for t in time_entry.get().split(',') if t.strip()]
        for t in current_times:
            time_listbox.insert(tk.END, t)

        btn_frame = tk.Frame(time_list_frame, bg='#D7F3F5')
        btn_frame.pack(side=tk.RIGHT, padx=10, fill=tk.Y)

        new_time_entry = tk.Entry(btn_frame, font=('Microsoft YaHei', 10), width=12)
        new_time_entry.insert(0, datetime.now().strftime("%H:%M:%S"))
        new_time_entry.pack(pady=3)

        def add_time():
            time_val = new_time_entry.get().strip()
            try:
                time.strptime(time_val, '%H:%M:%S')
                if time_val not in time_listbox.get(0, tk.END):
                    time_listbox.insert(tk.END, time_val)
            except ValueError:
                messagebox.showerror("格式错误", "请输入有效的时间格式 HH:MM:SS", parent=dialog)
        
        def delete_time():
            selection = time_listbox.curselection()
            if selection:
                time_listbox.delete(selection[0])

        tk.Button(btn_frame, text="添加 ↑", command=add_time).pack(pady=3, fill=tk.X)
        tk.Button(btn_frame, text="删除", command=delete_time).pack(pady=3, fill=tk.X)
        tk.Button(btn_frame, text="清空", command=lambda: time_listbox.delete(0, tk.END)).pack(pady=3, fill=tk.X)

        bottom_frame = tk.Frame(main_frame, bg='#D7F3F5')
        bottom_frame.pack(pady=10)

        def confirm_time():
            times = list(time_listbox.get(0, tk.END))
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
        self.center_window(dialog, 500, 450)

        main_frame = tk.Frame(dialog, bg='#D7F3F5', padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        week_type_var = tk.StringVar(value="week")

        week_frame = tk.LabelFrame(main_frame, text="按周", font=('Microsoft YaHei', 10, 'bold'),
                                  bg='#D7F3F5', padx=10, pady=10)
        week_frame.pack(fill=tk.X, pady=5)

        tk.Radiobutton(week_frame, text="每周", variable=week_type_var, value="week",
                      bg='#D7F3F5', font=('Microsoft YaHei', 10)).grid(row=0, column=0, sticky='w')

        week_vars = {}
        weekdays = [("周一", 1), ("周二", 2), ("周三", 3), ("周四", 4),
                   ("周五", 5), ("周六", 6), ("周日", 7)]

        for i, (day, num) in enumerate(weekdays):
            var = tk.IntVar(value=1)
            week_vars[num] = var
            row = (i // 4) + 1
            col = i % 4
            tk.Checkbutton(week_frame, text=day, variable=var, bg='#D7F3F5',
                          font=('Microsoft YaHei', 10)).grid(row=row, column=col, sticky='w', padx=10, pady=3)

        day_frame = tk.LabelFrame(main_frame, text="按月", font=('Microsoft YaHei', 10, 'bold'),
                                 bg='#D7F3F5', padx=10, pady=10)
        day_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        tk.Radiobutton(day_frame, text="每月", variable=week_type_var, value="day",
                      bg='#D7F3F5', font=('Microsoft YaHei', 10)).grid(row=0, column=0, sticky='w')

        day_vars = {}
        for i in range(1, 32):
            var = tk.IntVar(value=0)
            day_vars[i] = var
            row = (i - 1) // 7 + 1
            col = (i - 1) % 7
            # --- 字体修复 2: 几号的字体调大 ---
            tk.Checkbutton(day_frame, text=f"{i:02d}", variable=var, bg='#D7F3F5',
                          font=('Microsoft YaHei', 10)).grid(row=row, column=col, sticky='w', padx=8, pady=2)

        bottom_frame = tk.Frame(main_frame, bg='#D7F3F5')
        bottom_frame.pack(pady=10)

        def confirm_weekday():
            if week_type_var.get() == "week":
                selected = [str(num) for num, var in week_vars.items() if var.get() == 1]
                result = "每周:" + "".join(sorted(selected)) if selected else ""
            else:
                selected = [f"{num:02d}" for num, var in day_vars.items() if var.get() == 1]
                result = "每月:" + ",".join(sorted(selected)) if selected else ""

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
        self.center_window(dialog, 450, 220)

        main_frame = tk.Frame(dialog, bg='#D7F3F5', padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        from_frame = tk.Frame(main_frame, bg='#D7F3F5')
        from_frame.pack(pady=10, anchor='w')

        tk.Label(from_frame, text="从", font=('Microsoft YaHei', 10, 'bold'),
                bg='#D7F3F5').pack(side=tk.LEFT, padx=5)

        from_date_entry = tk.Entry(from_frame, font=('Microsoft YaHei', 10), width=18)
        from_date_entry.pack(side=tk.LEFT, padx=5)

        to_frame = tk.Frame(main_frame, bg='#D7F3F5')
        to_frame.pack(pady=10, anchor='w')

        tk.Label(to_frame, text="到", font=('Microsoft YaHei', 10, 'bold'),
                bg='#D7F3F5').pack(side=tk.LEFT, padx=5)

        to_date_entry = tk.Entry(to_frame, font=('Microsoft YaHei', 10), width=18)
        to_date_entry.pack(side=tk.LEFT, padx=5)
        
        try:
            current_range = date_range_entry.get().split('~')
            from_date_entry.insert(0, current_range[0].strip())
            to_date_entry.insert(0, current_range[1].strip())
        except IndexError:
            from_date_entry.insert(0, "2000-01-01")
            to_date_entry.insert(0, "2099-12-31")

        tk.Label(main_frame, text="格式: YYYY-MM-DD, 例如: 2025-01-01", font=('Microsoft YaHei', 10),
                bg='#D7F3F5', fg='#666').pack(pady=10)

        bottom_frame = tk.Frame(main_frame, bg='#D7F3F5')
        bottom_frame.pack(pady=10)

        def confirm_daterange():
            from_date = from_date_entry.get().strip()
            to_date = to_date_entry.get().strip()
            
            try:
                datetime.strptime(from_date, "%Y-%m-%d")
                datetime.strptime(to_date, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("格式错误", "日期格式不正确，应为 YYYY-MM-DD", parent=dialog)
                return

            result = f"{from_date} ~ {to_date}"
            date_range_entry.delete(0, tk.END)
            date_range_entry.insert(0, result)
            dialog.destroy()

        tk.Button(bottom_frame, text="确定", command=confirm_daterange, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 9, 'bold'), bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 9), bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)

    def update_task_list(self):
        """更新任务列表"""
        selection = self.task_tree.selection()
        
        for item in self.task_tree.get_children():
            self.task_tree.delete(item)

        for task in self.tasks:
            content_preview = os.path.basename(task['content']) if task.get('type') == 'audio' else (task['content'][:30] + '...' if len(task['content']) > 30 else task['content'])
            self.task_tree.insert('', tk.END, values=(
                task.get('name', ''), task.get('status', ''), task.get('time', ''),
                task.get('delay', 'ontime'), content_preview, task.get('volume', ''),
                task.get('weekday', ''), task.get('date_range', '')
            ))
        
        if selection:
            try:
                self.task_tree.selection_set(selection)
            except tk.TclError: # item may have been deleted
                pass

        self.stats_label.config(text=f"节目单：{len(self.tasks)}")
        if hasattr(self, 'status_labels'):
            self.status_labels[3].config(text=f"任务数量: {len(self.tasks)}")

    def update_status_bar(self):
        """更新状态栏"""
        if not self.running: return
        self.status_labels[0].config(text=f"当前时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.status_labels[1].config(text="系统状态: 运行中")
        self.root.after(1000, self.update_status_bar)

    def start_background_thread(self):
        """启动后台线程"""
        thread = threading.Thread(target=self._check_tasks, daemon=True)
        thread.start()

    def _check_tasks(self):
        """后台检查任务"""
        while self.running:
            now = datetime.now()
            current_date_str = now.strftime("%Y-%m-%d")
            current_time_str = now.strftime("%H:%M:%S")
            weekday = now.isoweekday()
            day_of_month = now.day

            for task in self.tasks:
                if task.get('status') != '启用': continue

                try:
                    start_date_str, end_date_str = [d.strip() for d in task.get('date_range', '').split('~')]
                    if not (datetime.strptime(start_date_str, "%Y-%m-%d").date() <= now.date() <= datetime.strptime(end_date_str, "%Y-%m-%d").date()):
                        continue
                except (ValueError, IndexError): pass

                run_today = False
                schedule_info = task.get('weekday', '每周:1234567')
                if schedule_info.startswith("每周:") and str(weekday) in schedule_info.replace("每周:", ""):
                    run_today = True
                elif schedule_info.startswith("每月:") and f"{day_of_month:02d}" in schedule_info.replace("每月:", "").split(','):
                    run_today = True
                if not run_today: continue

                for trigger_time in [t.strip() for t in task.get('time', '').split(',')]:
                    if trigger_time == current_time_str and task.get('last_run', {}).get(trigger_time) != current_date_str:
                        self.log(f"触发任务: [{task['name']}] at {trigger_time}")
                        self.root.after(0, self._execute_broadcast, task, trigger_time)
                        break
            time.sleep(1)

    def _execute_broadcast(self, task, trigger_time):
        """执行播报"""
        self.update_playing_text(f"[{task['name']}] 正在准备播放...")
        self.status_labels[2].config(text="播放状态: 播放中")

        if task.get('type') == 'audio':
            audio_path = task['content']
            if not os.path.exists(audio_path):
                self.log(f"错误：音频文件不存在 - {audio_path}")
                self.update_playing_text("错误：音频文件不存在")
                self.status_labels[2].config(text="播放状态: 错误")
                return
            self.update_playing_text(f"[{task['name']}] 正在播放音频: {os.path.basename(audio_path)}")
            self.log(f"播放音频: {task['name']} - {os.path.basename(audio_path)}")
            if AUDIO_AVAILABLE:
                threading.Thread(target=self._play_audio, args=(audio_path, task), daemon=True).start()
        else:
            content, repeat = task['content'], int(task.get('repeat', 1))
            self.update_playing_text(f"[{task['name']}] {content}")
            self.log(f"开始播报: {task['name']} (共 {repeat} 遍)")
            threading.Thread(target=self._speak, args=(content, task), daemon=True).start()

        if not isinstance(task.get('last_run'), dict): task['last_run'] = {}
        task['last_run'][trigger_time] = datetime.now().strftime("%Y-%m-%d")
        self.save_tasks()

    def _play_audio(self, audio_path, task):
        """播放音频文件"""
        try:
            pygame.mixer.music.load(audio_path)
            pygame.mixer.music.set_volume(float(task.get('volume', 80)) / 100.0)
            pygame.mixer.music.play()
            while pygame.mixer.music.get_busy() and self.running:
                time.sleep(0.1)
        except Exception as e: self.log(f"音频播放错误: {e}")
        finally: self.root.after(0, self.on_playback_finished)

    def _speak(self, text, task):
        """语音播报"""
        if not self.engine:
            self.log("错误：语音引擎未初始化"); self.root.after(0, self.on_playback_finished); return
        try:
            if task.get('prompt', 0) == 1 and AUDIO_AVAILABLE:
                prompt_path = os.path.join(PROMPT_FOLDER, task.get('prompt_file', 'tone-b.mp3'))
                if os.path.exists(prompt_path):
                    self.log(f"播放提示音: {os.path.basename(prompt_path)}")
                    sound = pygame.mixer.Sound(prompt_path)
                    sound.play()
                    time.sleep(sound.get_length())
            
            self.engine.setProperty('volume', float(task.get('volume', 1.0)))
            self.engine.setProperty('rate', int(task.get('speed', 150)))
            if (voice_name := task.get('voice')):
                for v in self.engine.getProperty('voices'):
                    if v.name == voice_name: self.engine.setProperty('voice', v.id); break

            for i in range(int(task.get('repeat', 1))):
                if not self.running: break
                self.log(f"正在播报第 {i+1} 遍")
                self.engine.say(text)
                self.engine.runAndWait()
        except Exception as e: self.log(f"播报错误: {e}")
        finally: self.root.after(0, self.on_playback_finished)

    def on_playback_finished(self):
        """播放完成后的回调"""
        self.update_playing_text("等待下一个任务...")
        self.status_labels[2].config(text="播放状态: 待机")
        self.log("播放结束")

    def log(self, message):
        """记录日志"""
        def _log():
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} -> {message}\n")
            self.log_text.see(tk.END)
            self.log_text.config(state='disabled')
        self.root.after(0, _log)

    def update_playing_text(self, message):
        """更新正在播放区域"""
        def _update():
            self.playing_text.config(state='normal')
            self.playing_text.delete('1.0', tk.END)
            self.playing_text.insert('1.0', message)
            self.playing_text.config(state='disabled')
        self.root.after(0, _update)

    def save_tasks(self):
        """保存任务"""
        try:
            with open(self.task_file, 'w', encoding='utf-8') as f:
                json.dump(self.tasks, f, ensure_ascii=False, indent=2)
        except Exception as e: self.log(f"保存任务失败: {e}")

    def load_tasks(self):
        """加载任务"""
        if not os.path.exists(self.task_file): return
        try:
            with open(self.task_file, 'r', encoding='utf-8') as f: self.tasks = json.load(f)
            migrated = any('last_run' not in t or not isinstance(t.get('last_run'), dict) for t in self.tasks)
            if migrated:
                for task in self.tasks: task['last_run'] = {}
                self.log("检测到旧版任务数据，已自动迁移。"); self.save_tasks()
            self.update_task_list(); self.log(f"已加载 {len(self.tasks)} 个节目")
        except Exception as e: self.log(f"加载任务失败: {e}")

    def center_window(self, win, width, height):
        """窗口居中"""
        x = (win.winfo_screenwidth() // 2) - (width // 2)
        y = (win.winfo_screenheight() // 2) - (height // 2)
        win.geometry(f'{width}x{height}+{x}+{y}')

    def on_closing(self):
        """处理窗口关闭"""
        if messagebox.askokcancel("退出", "确定要退出定时播音软件吗？"):
            self.running = False
            self.save_tasks()
            if AUDIO_AVAILABLE and pygame.mixer.get_init(): pygame.mixer.quit()
            if self.engine and self.engine.isBusy(): self.engine.stop()
            self.root.destroy()

def main():
    root = tk.Tk()
    app = TimedBroadcastApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
