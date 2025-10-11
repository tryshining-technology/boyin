import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import json
import threading
import time
from datetime import datetime
import os
import random
import sys

# 尝试导入所需库
TRAY_AVAILABLE = False
try:
    from pystray import MenuItem as item, Icon
    from PIL import Image
    TRAY_AVAILABLE = True
except ImportError:
    print("警告: pystray 或 Pillow 未安装，最小化到托盘功能不可用。")

WIN32COM_AVAILABLE = False
try:
    import win32com.client
    import pythoncom
    from pywintypes import com_error
    WIN32COM_AVAILABLE = True
except ImportError:
    print("警告: pywin32 未安装，语音功能将受限。")

AUDIO_AVAILABLE = False
try:
    import pygame
    pygame.mixer.init()
    AUDIO_AVAILABLE = True
except ImportError:
    print("警告: pygame 未安装，音频播放功能将不可用。")
except Exception as e:
    print(f"警告: pygame 初始化失败 - {e}，音频播放功能将不可用。")


def resource_path(relative_path):
    """ 获取资源的绝对路径，无论是开发环境还是打包后 """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- 全局路径设置 ---
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

TASK_FILE = os.path.join(application_path, "broadcast_tasks.json")
PROMPT_FOLDER = os.path.join(application_path, "提示音")
AUDIO_FOLDER = os.path.join(application_path, "音频文件")
BGM_FOLDER = os.path.join(application_path, "文稿背景")
ICON_FILE = resource_path("icon.ico")

class TimedBroadcastApp:
    def __init__(self, root):
        self.root = root
        self.root.title("定时播音")
        self.root.geometry("1400x800")
        self.root.configure(bg='#E8F4F8')
        
        if os.path.exists(ICON_FILE):
            try:
                self.root.iconbitmap(ICON_FILE)
            except Exception as e:
                print(f"加载窗口图标失败: {e}")

        self.tasks = []
        self.running = True
        self.task_file = TASK_FILE
        self.tray_icon = None

        self.is_playing = threading.Event()
        self.playback_queue = []
        self.queue_lock = threading.Lock()

        self.create_folder_structure()
        self.create_widgets()
        self.load_tasks()
        self.start_background_thread()
        self.root.protocol("WM_DELETE_WINDOW", self.show_quit_dialog)

    def create_folder_structure(self):
        """创建所有必要的文件夹"""
        for folder in [PROMPT_FOLDER, AUDIO_FOLDER, BGM_FOLDER]:
            if not os.path.exists(folder):
                os.makedirs(folder)
                self.log(f"已创建文件夹: {folder}") if hasattr(self, 'log_text') else None

    def create_widgets(self):
        self.nav_frame = tk.Frame(self.root, bg='#A8D8E8', width=160)
        self.nav_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.nav_frame.pack_propagate(False)

        nav_buttons = [
            ("定时广播", ""), ("立即插播", ""), ("节假日", ""),
            ("语音广告 制作", ""), ("设置", "")
        ]
        for i, (title, subtitle) in enumerate(nav_buttons):
            btn_frame = tk.Frame(self.nav_frame, bg='#5DADE2' if i == 0 else '#A8D8E8')
            btn_frame.pack(fill=tk.X, pady=1)
            btn = tk.Button(btn_frame, text=title, bg='#5DADE2' if i == 0 else '#A8D8E8',
                          fg='white' if i == 0 else 'black', font=('Microsoft YaHei', 13, 'bold'),
                          bd=0, padx=10, pady=8, anchor='w', command=lambda t=title: self.switch_page(t))
            btn.pack(fill=tk.X)
            if subtitle:
                sub_label = tk.Label(btn_frame, text=subtitle, bg='#5DADE2' if i == 0 else '#A8D8E8',
                                   fg='#555' if i == 0 else '#666',
                                   font=('Microsoft YaHei', 10), anchor='w', padx=10)
                sub_label.pack(fill=tk.X)

        self.main_frame = tk.Frame(self.root, bg='white')
        self.main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.create_scheduled_broadcast_page()

    def switch_page(self, page_name):
        if page_name != "定时广播":
            messagebox.showinfo("提示", f"页面 [{page_name}] 正在开发中...")
            self.log(f"功能开发中: {page_name}")

    def create_scheduled_broadcast_page(self):
        top_frame = tk.Frame(self.main_frame, bg='white')
        top_frame.pack(fill=tk.X, padx=10, pady=10)
        title_label = tk.Label(top_frame, text="定时广播", font=('Microsoft YaHei', 14, 'bold'),
                              bg='white', fg='#2C5F7C')
        title_label.pack(side=tk.LEFT)
        btn_frame = tk.Frame(top_frame, bg='white')
        btn_frame.pack(side=tk.RIGHT)
        
        buttons = [("导入节目单", self.import_tasks, '#1ABC9C'), ("导出节目单", self.export_tasks, '#1ABC9C')]
        for text, cmd, color in buttons:
            btn = tk.Button(btn_frame, text=text, command=cmd, bg=color, fg='white',
                          font=('Microsoft YaHei', 9), bd=0, padx=12, pady=5, cursor='hand2')
            btn.pack(side=tk.LEFT, padx=3)

        stats_frame = tk.Frame(self.main_frame, bg='#F0F8FF')
        stats_frame.pack(fill=tk.X, padx=10, pady=5)
        self.stats_label = tk.Label(stats_frame, text="节目单：0", font=('Microsoft YaHei', 10),
                                   bg='#F0F8FF', fg='#2C5F7C', anchor='w', padx=10)
        self.stats_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        table_frame = tk.Frame(self.main_frame, bg='white')
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        columns = ('节目名称', '状态', '开始时间', '模式', '音频或文字', '音量', '周几/几号', '日期范围')
        self.task_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=12)
        col_widths = [200, 60, 140, 70, 300, 60, 100, 120]
        for col, width in zip(columns, col_widths):
            self.task_tree.heading(col, text=col)
            self.task_tree.column(col, width=width, anchor='w' if col in ['节目名称', '音频或文字'] else 'center')
        self.task_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.task_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.task_tree.configure(yscrollcommand=scrollbar.set)
        
        self.task_tree.bind("<Button-3>", self.show_context_menu)
        self.task_tree.bind("<Double-1>", self.on_double_click_edit)

        playing_frame = tk.LabelFrame(self.main_frame, text="正在播：", font=('Microsoft YaHei', 10),
                                     bg='white', fg='#2C5F7C', padx=10, pady=5)
        playing_frame.pack(fill=tk.X, padx=10, pady=5)
        self.playing_text = scrolledtext.ScrolledText(playing_frame, height=3, font=('Microsoft YaHei', 9),
                                                     bg='#FFFEF0', wrap=tk.WORD, state='disabled')
        self.playing_text.pack(fill=tk.BOTH, expand=True)
        self.update_playing_text("等待播放...")

        # --- 修改：日志区域布局 ---
        log_frame = tk.LabelFrame(self.main_frame, text="", font=('Microsoft YaHei', 10),
                                 bg='white', fg='#2C5F7C', padx=10, pady=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        log_header_frame = tk.Frame(log_frame, bg='white')
        log_header_frame.pack(fill=tk.X)

        log_label = tk.Label(log_header_frame, text="日志：", font=('Microsoft YaHei', 10, 'bold'),
                             bg='white', fg='#2C5F7C')
        log_label.pack(side=tk.LEFT)

        clear_log_btn = tk.Button(log_header_frame, text="清除日志", command=self.clear_log,
                                  font=('Microsoft YaHei', 8), bd=0, bg='#EAEAEA',
                                  fg='#333', cursor='hand2', padx=5, pady=0)
        clear_log_btn.pack(side=tk.LEFT, padx=10)
        # --- 修改结束 ---

        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, font=('Microsoft YaHei', 9),
                                                 bg='#F9F9F9', wrap=tk.WORD, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True)

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

        self.update_status_bar()
        self.log("定时播音软件已启动")
    
    def clear_log(self):
        """清除日志文本框中的所有内容"""
        if messagebox.askyesno("确认操作", "您确定要清空所有日志记录吗？\n此操作不可恢复。"):
            self.log_text.config(state='normal')
            self.log_text.delete('1.0', tk.END)
            self.log_text.config(state='disabled')
            self.log("日志已清空。")

    def on_double_click_edit(self, event):
        if self.task_tree.identify_row(event.y):
            self.edit_task()

    def show_context_menu(self, event):
        iid = self.task_tree.identify_row(event.y)
        
        context_menu = tk.Menu(self.root, tearoff=0, font=('Microsoft YaHei', 10))

        if iid:
            if iid not in self.task_tree.selection():
                self.task_tree.selection_set(iid)
            
            context_menu.add_command(label="▶️ 立即播放", command=self.play_now)
            context_menu.add_separator()
            context_menu.add_command(label="✏️ 修改", command=self.edit_task)
            context_menu.add_command(label="❌ 删除", command=self.delete_task)
            context_menu.add_command(label="📋 复制", command=self.copy_task)
            context_menu.add_separator()
            context_menu.add_command(label="🔼 上移", command=lambda: self.move_task(-1))
            context_menu.add_command(label="🔽 下移", command=lambda: self.move_task(1))
            context_menu.add_separator()
            context_menu.add_command(label="▶️ 启用", command=self.enable_task)
            context_menu.add_command(label="⏸️ 禁用", command=self.disable_task)

        else:
            self.task_tree.selection_set()
            context_menu.add_command(label="➕ 添加节目", command=self.add_task)
        
        context_menu.add_separator()
        context_menu.add_command(label="⏹️ 停止当前播放", command=self.stop_current_playback, state="normal")
        
        context_menu.post(event.x_root, event.y_root)
    
    # --- 新增：强制停止当前播放的方法 ---
    def _force_stop_playback(self):
        """强制停止当前所有播放活动"""
        if self.is_playing.is_set():
            self.log("接收到中断指令，正在停止当前播放...")
            if AUDIO_AVAILABLE and pygame.mixer.music.get_busy():
                pygame.mixer.music.stop()
            # 注意：强行中断正在进行的 win32com tts 较为复杂且风险高，
            # 这里主要处理最常见的长时任务——音频播放。
            # TTS通常较短，让其自然结束是更稳妥的做法。
            # 如果TTS也需要中断，需要引入更复杂的线程管理。
            self.on_playback_finished() # 重置状态
    
    # --- 修改：立即播放逻辑 ---
    def play_now(self):
        """立即中断并播放选定的任务"""
        selection = self.task_tree.selection()
        if not selection: 
            messagebox.showwarning("提示", "请先选择一个要立即播放的节目。")
            return
        
        index = self.task_tree.index(selection[0])
        task = self.tasks[index]

        self.log(f"手动触发高优先级播放: {task['name']}")
        
        # 1. 强制停止当前播放
        self._force_stop_playback()
        
        # 2. 清空队列，并将当前任务置于队首
        with self.queue_lock:
            self.playback_queue.clear()
            self.playback_queue.insert(0, (task, "manual_play"))
            self.log("播放队列已清空，新任务已置顶。")
        
        # 3. 立即处理队列
        self.root.after(0, self._process_queue)

    # --- 修改：停止当前播放的功能 ---
    def stop_current_playback(self):
        self.log("手动触发“停止当前播放”...")
        self._force_stop_playback()
        with self.queue_lock:
            if self.playback_queue:
                self.playback_queue.clear()
                self.log("等待播放的队列也已清空。")

    def add_task(self):
        choice_dialog = tk.Toplevel(self.root)
        choice_dialog.title("选择节目类型")
        choice_dialog.geometry("350x250")
        choice_dialog.resizable(False, False)
        choice_dialog.transient(self.root); choice_dialog.grab_set()
        self.center_window(choice_dialog, 350, 250)
        main_frame = tk.Frame(choice_dialog, padx=20, pady=20, bg='#F0F0F0')
        main_frame.pack(fill=tk.BOTH, expand=True)
        title_label = tk.Label(main_frame, text="请选择要添加的节目类型",
                              font=('Microsoft YaHei', 13, 'bold'), fg='#2C5F7C', bg='#F0F0F0')
        title_label.pack(pady=15)
        btn_frame = tk.Frame(main_frame, bg='#F0F0F0')
        btn_frame.pack(expand=True)
        audio_btn = tk.Button(btn_frame, text="🎵 音频节目", command=lambda: self.open_audio_dialog(choice_dialog),
                             bg='#5DADE2', fg='white', font=('Microsoft YaHei', 11, 'bold'),
                             bd=0, padx=30, pady=12, cursor='hand2', width=15)
        audio_btn.pack(pady=8)
        voice_btn = tk.Button(btn_frame, text="🎙️ 语音节目", command=lambda: self.open_voice_dialog(choice_dialog),
                             bg='#3498DB', fg='white', font=('Microsoft YaHei', 11, 'bold'),
                             bd=0, padx=30, pady=12, cursor='hand2', width=15)
        voice_btn.pack(pady=8)

    def open_audio_dialog(self, parent_dialog, task_to_edit=None, index=None):
        parent_dialog.destroy()
        is_edit_mode = task_to_edit is not None

        dialog = tk.Toplevel(self.root)
        dialog.title("修改音频节目" if is_edit_mode else "添加音频节目")
        dialog.geometry("850x750")
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set(); dialog.configure(bg='#E8E8E8')
        
        main_frame = tk.Frame(dialog, bg='#E8E8E8', padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        content_frame = tk.LabelFrame(main_frame, text="内容", font=('Microsoft YaHei', 11, 'bold'),
                                     bg='#E8E8E8', padx=10, pady=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=5)
        
        tk.Label(content_frame, text="节目名称:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=0, column=0, sticky='e', padx=5, pady=5)
        name_entry = tk.Entry(content_frame, font=('Microsoft YaHei', 10), width=55)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        
        audio_type_var = tk.StringVar(value="single")
        tk.Label(content_frame, text="音频文件", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=1, column=0, sticky='e', padx=5, pady=5)
        audio_single_frame = tk.Frame(content_frame, bg='#E8E8E8')
        audio_single_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        tk.Radiobutton(audio_single_frame, text="", variable=audio_type_var, value="single",
                      bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT)
        audio_single_entry = tk.Entry(audio_single_frame, font=('Microsoft YaHei', 10), width=35, state='readonly')
        audio_single_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(audio_single_frame, text="00:00", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT, padx=10)
        def select_single_audio():
            filename = filedialog.askopenfilename(title="选择音频文件", initialdir=AUDIO_FOLDER,
                filetypes=[("音频文件", "*.mp3 *.wav *.ogg *.flac *.m4a"), ("所有文件", "*.*")])
            if filename:
                audio_single_entry.config(state='normal'); audio_single_entry.delete(0, tk.END)
                audio_single_entry.insert(0, filename); audio_single_entry.config(state='readonly')
        tk.Button(audio_single_frame, text="选取...", command=select_single_audio, bg='#D0D0D0',
                 font=('Microsoft YaHei', 10), bd=1, padx=15, pady=3).pack(side=tk.LEFT, padx=5)
        
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
                audio_folder_entry.config(state='normal'); audio_folder_entry.delete(0, tk.END)
                audio_folder_entry.insert(0, foldername); audio_folder_entry.config(state='readonly')
        tk.Button(audio_folder_frame, text="选取...", command=select_folder, bg='#D0D0D0',
                 font=('Microsoft YaHei', 10), bd=1, padx=15, pady=3).pack(side=tk.LEFT, padx=5)
        
        play_order_frame = tk.Frame(content_frame, bg='#E8E8E8')
        play_order_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=5)
        play_order_var = tk.StringVar(value="sequential")
        tk.Radiobutton(play_order_frame, text="顺序播", variable=play_order_var, value="sequential",
                      bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(play_order_frame, text="随机播", variable=play_order_var, value="random",
                      bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT, padx=10)
        
        volume_frame = tk.Frame(content_frame, bg='#E8E8E8')
        volume_frame.grid(row=4, column=1, columnspan=3, sticky='w', padx=5, pady=8)
        tk.Label(volume_frame, text="音量:", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT)
        volume_entry = tk.Entry(volume_frame, font=('Microsoft YaHei', 10), width=10)
        volume_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(volume_frame, text="0-100", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT, padx=5)
        
        time_frame = tk.LabelFrame(main_frame, text="时间", font=('Microsoft YaHei', 12, 'bold'),
                                   bg='#E8E8E8', padx=15, pady=15)
        time_frame.grid(row=1, column=0, sticky='ew', pady=10)
        tk.Label(time_frame, text="开始时间:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=0, column=0, sticky='e', padx=5, pady=5)
        start_time_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 10), width=50)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        tk.Label(time_frame, text="《可多个,用英文逗号,隔开》", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=0, column=2, sticky='w', padx=5)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry),
                 bg='#D0D0D0', font=('Microsoft YaHei', 10), bd=1, padx=12, pady=2).grid(row=0, column=3, padx=5)
        
        interval_var = tk.StringVar(value="first")
        interval_frame1 = tk.Frame(time_frame, bg='#E8E8E8')
        interval_frame1.grid(row=1, column=1, columnspan=2, sticky='w', padx=5, pady=5)
        tk.Label(time_frame, text="间隔播报:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=1, column=0, sticky='e', padx=5, pady=5)
        tk.Radiobutton(interval_frame1, text="播 n 首", variable=interval_var, value="first",
                      bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT)
        interval_first_entry = tk.Entry(interval_frame1, font=('Microsoft YaHei', 10), width=15)
        interval_first_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(interval_frame1, text="(单曲时,指 n 遍)", font=('Microsoft YaHei', 10),
                bg='#E8E8E8').pack(side=tk.LEFT, padx=5)
        
        interval_frame2 = tk.Frame(time_frame, bg='#E8E8E8')
        interval_frame2.grid(row=2, column=1, columnspan=2, sticky='w', padx=5, pady=5)
        tk.Radiobutton(interval_frame2, text="播 n 秒", variable=interval_var, value="seconds",
                      bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT)
        interval_seconds_entry = tk.Entry(interval_frame2, font=('Microsoft YaHei', 10), width=15)
        interval_seconds_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(interval_frame2, text="(3600秒 = 1小时)", font=('Microsoft YaHei', 10),
                bg='#E8E8E8').pack(side=tk.LEFT, padx=5)
        
        tk.Label(time_frame, text="周几/几号:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=3, column=0, sticky='e', padx=5, pady=8)
        weekday_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 10), width=50)
        weekday_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=8)
        tk.Button(time_frame, text="选取...", command=lambda: self.show_weekday_settings_dialog(weekday_entry),
                 bg='#D0D0D0', font=('Microsoft YaHei', 10), bd=1, padx=15, pady=3).grid(row=3, column=3, padx=5)
        
        tk.Label(time_frame, text="日期范围:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=4, column=0, sticky='e', padx=5, pady=8)
        date_range_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 10), width=50)
        date_range_entry.grid(row=4, column=1, sticky='ew', padx=5, pady=8)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry),
                 bg='#D0D0D0', font=('Microsoft YaHei', 10), bd=1, padx=15, pady=3).grid(row=4, column=3, padx=5)
        
        other_frame = tk.LabelFrame(main_frame, text="其它", font=('Microsoft YaHei', 11, 'bold'),
                                    bg='#E8E8E8', padx=10, pady=10)
        other_frame.grid(row=2, column=0, sticky='ew', pady=5)
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

        if is_edit_mode:
            task = task_to_edit
            name_entry.insert(0, task.get('name', ''))
            start_time_entry.insert(0, task.get('time', ''))
            audio_type_var.set(task.get('audio_type', 'single'))
            if task.get('audio_type') == 'single':
                audio_single_entry.config(state='normal')
                audio_single_entry.insert(0, task.get('content', ''))
                audio_single_entry.config(state='readonly')
            else:
                audio_folder_entry.config(state='normal')
                audio_folder_entry.insert(0, task.get('content', ''))
                audio_folder_entry.config(state='readonly')
            play_order_var.set(task.get('play_order', 'sequential'))
            volume_entry.insert(0, task.get('volume', '80'))
            interval_var.set(task.get('interval_type', 'first'))
            interval_first_entry.insert(0, task.get('interval_first', '1'))
            interval_seconds_entry.insert(0, task.get('interval_seconds', '600'))
            weekday_entry.insert(0, task.get('weekday', '每周:1234567'))
            date_range_entry.insert(0, task.get('date_range', '2000-01-01 ~ 2099-12-31'))
            delay_var.set(task.get('delay', 'ontime'))
        else:
            volume_entry.insert(0, "80")
            interval_first_entry.insert(0, "1")
            interval_seconds_entry.insert(0, "600")
            weekday_entry.insert(0, "每周:1234567")
            date_range_entry.insert(0, "2000-01-01 ~ 2099-12-31")

        button_frame = tk.Frame(main_frame, bg='#E8E8E8')
        button_frame.grid(row=3, column=0, pady=20)
        
        def save_task():
            audio_path = audio_single_entry.get().strip() if audio_type_var.get() == "single" else audio_folder_entry.get().strip()
            if not audio_path: messagebox.showwarning("警告", "请选择音频文件或文件夹", parent=dialog); return
            
            new_task_data = {'name': name_entry.get().strip(), 'time': start_time_entry.get().strip(), 'content': audio_path,
                             'type': 'audio', 'audio_type': audio_type_var.get(), 'play_order': play_order_var.get(),
                             'volume': volume_entry.get().strip() or "80", 'interval_type': interval_var.get(),
                             'interval_first': interval_first_entry.get().strip(), 'interval_seconds': interval_seconds_entry.get().strip(),
                             'weekday': weekday_entry.get().strip(), 'date_range': date_range_entry.get().strip(),
                             'delay': delay_var.get(), 
                             'status': '启用' if not is_edit_mode else task_to_edit.get('status', '启用'), 
                             'last_run': {} if not is_edit_mode else task_to_edit.get('last_run', {})}
            
            if not new_task_data['name'] or not new_task_data['time']:
                messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog); return
            
            if is_edit_mode:
                self.tasks[index] = new_task_data
                self.log(f"已修改音频节目: {new_task_data['name']}")
            else:
                self.tasks.append(new_task_data)
                self.log(f"已添加音频节目: {new_task_data['name']}")
                
            self.update_task_list(); self.save_tasks(); dialog.destroy()
        
        button_text = "保存修改" if is_edit_mode else "添加"
        tk.Button(button_frame, text=button_text, command=save_task, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 10, 'bold'), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 10), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        
        content_frame.columnconfigure(1, weight=1)
        time_frame.columnconfigure(1, weight=1)

    def open_voice_dialog(self, parent_dialog, task_to_edit=None, index=None):
        parent_dialog.destroy()
        is_edit_mode = task_to_edit is not None

        dialog = tk.Toplevel(self.root)
        dialog.title("修改语音节目" if is_edit_mode else "添加语音节目")
        dialog.geometry("800x800")
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set(); dialog.configure(bg='#E8E8E8')
        
        main_frame = tk.Frame(dialog, bg='#E8E8E8', padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        content_frame = tk.LabelFrame(main_frame, text="内容", font=('Microsoft YaHei', 11, 'bold'),
                                     bg='#E8E8E8', padx=10, pady=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=5)
        
        tk.Label(content_frame, text="节目名称:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=0, column=0, sticky='w', padx=5, pady=5)
        name_entry = tk.Entry(content_frame, font=('Microsoft YaHei', 10), width=65)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        
        tk.Label(content_frame, text="播音文字:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=1, column=0, sticky='nw', padx=5, pady=5)
        text_frame = tk.Frame(content_frame, bg='#E8E8E8')
        text_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        content_text = scrolledtext.ScrolledText(text_frame, height=5, font=('Microsoft YaHei', 10), width=65, wrap=tk.WORD)
        content_text.pack(fill=tk.BOTH, expand=True)

        tk.Label(content_frame, text="播音员:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=2, column=0, sticky='w', padx=5, pady=8)
        voice_frame = tk.Frame(content_frame, bg='#E8E8E8')
        voice_frame.grid(row=2, column=1, columnspan=3, sticky='w', padx=5, pady=8)
        available_voices = self.get_available_voices()
        voice_var = tk.StringVar()
        voice_combo = ttk.Combobox(voice_frame, textvariable=voice_var, values=available_voices,
                                   font=('Microsoft YaHei', 10), width=50, state='readonly')
        voice_combo.pack(side=tk.LEFT)

        speech_params_frame = tk.Frame(content_frame, bg='#E8E8E8')
        speech_params_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=5)
        tk.Label(speech_params_frame, text="语速(-10~10):", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT, padx=(0,5))
        speed_entry = tk.Entry(speech_params_frame, font=('Microsoft YaHei', 10), width=8)
        speed_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(speech_params_frame, text="音调(-10~10):", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT, padx=(10,5))
        pitch_entry = tk.Entry(speech_params_frame, font=('Microsoft YaHei', 10), width=8)
        pitch_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(speech_params_frame, text="音量(0-100):", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT, padx=(10,5))
        volume_entry = tk.Entry(speech_params_frame, font=('Microsoft YaHei', 10), width=8)
        volume_entry.pack(side=tk.LEFT, padx=5)

        prompt_var = tk.IntVar()
        prompt_frame = tk.Frame(content_frame, bg='#E8E8E8')
        prompt_frame.grid(row=4, column=1, columnspan=3, sticky='w', padx=5, pady=5)
        tk.Checkbutton(prompt_frame, text="提示音:", variable=prompt_var, bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT)
        prompt_file_var, prompt_volume_var = tk.StringVar(), tk.StringVar()
        prompt_file_entry = tk.Entry(prompt_frame, textvariable=prompt_file_var, font=('Microsoft YaHei', 10), width=20, state='readonly')
        prompt_file_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(prompt_frame, text="...", command=lambda: self.select_file_for_entry(PROMPT_FOLDER, prompt_file_var)).pack(side=tk.LEFT)
        tk.Label(prompt_frame, text="音量(0-100):", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT, padx=(10,5))
        tk.Entry(prompt_frame, textvariable=prompt_volume_var, font=('Microsoft YaHei', 10), width=8).pack(side=tk.LEFT, padx=5)

        bgm_var = tk.IntVar()
        bgm_frame = tk.Frame(content_frame, bg='#E8E8E8')
        bgm_frame.grid(row=5, column=1, columnspan=3, sticky='w', padx=5, pady=5)
        tk.Checkbutton(bgm_frame, text="背景音乐:", variable=bgm_var, bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT)
        bgm_file_var, bgm_volume_var = tk.StringVar(), tk.StringVar()
        bgm_file_entry = tk.Entry(bgm_frame, textvariable=bgm_file_var, font=('Microsoft YaHei', 10), width=20, state='readonly')
        bgm_file_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(bgm_frame, text="...", command=lambda: self.select_file_for_entry(BGM_FOLDER, bgm_file_var)).pack(side=tk.LEFT)
        tk.Label(bgm_frame, text="音量(0-100):", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT, padx=(10,5))
        tk.Entry(bgm_frame, textvariable=bgm_volume_var, font=('Microsoft YaHei', 10), width=8).pack(side=tk.LEFT, padx=5)

        time_frame = tk.LabelFrame(main_frame, text="时间", font=('Microsoft YaHei', 11, 'bold'),
                                   bg='#E8E8E8', padx=10, pady=10)
        time_frame.grid(row=1, column=0, sticky='ew', pady=5)
        tk.Label(time_frame, text="开始时间:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=0, column=0, sticky='e', padx=5, pady=5)
        start_time_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 10), width=50)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        tk.Label(time_frame, text="《可多个,用英文逗号,隔开》", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=0, column=2, sticky='w', padx=5)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry),
                 bg='#D0D0D0', font=('Microsoft YaHei', 10), bd=1, padx=12, pady=2).grid(row=0, column=3, padx=5)
        
        tk.Label(time_frame, text="播 n 遍:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=1, column=0, sticky='e', padx=5, pady=5)
        repeat_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 10), width=12)
        repeat_entry.grid(row=1, column=1, sticky='w', padx=5, pady=5)
        
        tk.Label(time_frame, text="周几/几号:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=2, column=0, sticky='e', padx=5, pady=5)
        weekday_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 10), width=50)
        weekday_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=5)
        tk.Button(time_frame, text="选取...", command=lambda: self.show_weekday_settings_dialog(weekday_entry),
                 bg='#D0D0D0', font=('Microsoft YaHei', 10), bd=1, padx=12, pady=2).grid(row=2, column=3, padx=5)
        
        tk.Label(time_frame, text="日期范围:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=3, column=0, sticky='e', padx=5, pady=5)
        date_range_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 10), width=50)
        date_range_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=5)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry),
                 bg='#D0D0D0', font=('Microsoft YaHei', 10), bd=1, padx=12, pady=2).grid(row=3, column=3, padx=5)
        
        other_frame = tk.LabelFrame(main_frame, text="其它", font=('Microsoft YaHei', 12, 'bold'),
                                    bg='#E8E8E8', padx=15, pady=15)
        other_frame.grid(row=2, column=0, sticky='ew', pady=10)
        delay_var = tk.StringVar(value="delay")
        tk.Label(other_frame, text="准时/延后:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=0, column=0, sticky='e', padx=5, pady=3)
        delay_frame = tk.Frame(other_frame, bg='#E8E8E8')
        delay_frame.grid(row=0, column=1, sticky='w', padx=5, pady=3)
        tk.Radiobutton(delay_frame, text="准时播 - 频道内,若有别的节目正在播，终止他们",
                      variable=delay_var, value="ontime", bg='#E8E8E8',
                      font=('Microsoft YaHei', 10)).pack(anchor='w', pady=2)
        tk.Radiobutton(delay_frame, text="可延后 - 频道内,若有别的节目正在播，排队等候",
                      variable=delay_var, value="delay", bg='#E8E8E8',
                      font=('Microsoft YaHei', 10)).pack(anchor='w', pady=2)

        if is_edit_mode:
            task = task_to_edit
            name_entry.insert(0, task.get('name', ''))
            content_text.insert('1.0', task.get('content', ''))
            voice_var.set(task.get('voice', ''))
            speed_entry.insert(0, task.get('speed', '0'))
            pitch_entry.insert(0, task.get('pitch', '0'))
            volume_entry.insert(0, task.get('volume', '80'))
            prompt_var.set(task.get('prompt', 0))
            prompt_file_var.set(task.get('prompt_file', ''))
            prompt_volume_var.set(task.get('prompt_volume', '80'))
            bgm_var.set(task.get('bgm', 0))
            bgm_file_var.set(task.get('bgm_file', ''))
            bgm_volume_var.set(task.get('bgm_volume', '40'))
            start_time_entry.insert(0, task.get('time', ''))
            repeat_entry.insert(0, task.get('repeat', '1'))
            weekday_entry.insert(0, task.get('weekday', '每周:1234567'))
            date_range_entry.insert(0, task.get('date_range', '2000-01-01 ~ 2099-12-31'))
            delay_var.set(task.get('delay', 'delay'))
        else:
            speed_entry.insert(0, "0")
            pitch_entry.insert(0, "0")
            volume_entry.insert(0, "80")
            prompt_var.set(0)
            prompt_volume_var.set("80")
            bgm_var.set(0)
            bgm_volume_var.set("40")
            repeat_entry.insert(0, "1")
            weekday_entry.insert(0, "每周:1234567")
            date_range_entry.insert(0, "2000-01-01 ~ 2099-12-31")

        button_frame = tk.Frame(main_frame, bg='#E8E8E8')
        button_frame.grid(row=3, column=0, pady=20)
        
        def save_task():
            content = content_text.get('1.0', tk.END).strip()
            if not content: messagebox.showwarning("警告", "请输入播音文字内容", parent=dialog); return
            
            new_task_data = {'name': name_entry.get().strip(), 'time': start_time_entry.get().strip(), 'content': content,
                             'type': 'voice', 'voice': voice_var.get(), 
                             'speed': speed_entry.get().strip() or "0",
                             'pitch': pitch_entry.get().strip() or "0",
                             'volume': volume_entry.get().strip() or "80",
                             'prompt': prompt_var.get(), 'prompt_file': prompt_file_var.get(),
                             'prompt_volume': prompt_volume_var.get(),
                             'bgm': bgm_var.get(), 'bgm_file': bgm_file_var.get(),
                             'bgm_volume': bgm_volume_var.get(),
                             'repeat': repeat_entry.get().strip() or "1",
                             'weekday': weekday_entry.get().strip(), 'date_range': date_range_entry.get().strip(),
                             'delay': delay_var.get(), 
                             'status': '启用' if not is_edit_mode else task_to_edit.get('status', '启用'), 
                             'last_run': {} if not is_edit_mode else task_to_edit.get('last_run', {})}
            
            if not new_task_data['name'] or not new_task_data['time']:
                messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog); return
            
            if is_edit_mode:
                self.tasks[index] = new_task_data
                self.log(f"已修改语音节目: {new_task_data['name']}")
            else:
                self.tasks.append(new_task_data)
                self.log(f"已添加语音节目: {new_task_data['name']}")
                
            self.update_task_list(); self.save_tasks(); dialog.destroy()
        
        button_text = "保存修改" if is_edit_mode else "添加"
        tk.Button(button_frame, text=button_text, command=save_task, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 10, 'bold'), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 10), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        
        content_frame.columnconfigure(1, weight=1)
        time_frame.columnconfigure(1, weight=1)

    def get_available_voices(self):
        """获取系统可用语音列表的最终方法"""
        available_voices = []
        if WIN32COM_AVAILABLE:
            try:
                pythoncom.CoInitialize()
                speaker = win32com.client.Dispatch("SAPI.SpVoice")
                voices = speaker.GetVoices()
                available_voices = [v.GetDescription() for v in voices]
                # self.log("成功通过 win32com 刷新语音列表。") # 减少不必要的日志
                pythoncom.CoUninitialize()
            except Exception as e:
                self.log(f"警告: 使用 win32com 获取语音列表失败 - {e}")
                available_voices = []
        
        return available_voices
    
    def select_file_for_entry(self, initial_dir, string_var):
        """通用文件选择对话框"""
        filename = filedialog.askopenfilename(
            title="选择文件",
            initialdir=initial_dir,
            filetypes=[("音频文件", "*.mp3 *.wav *.ogg *.flac *.m4a"), ("所有文件", "*.*")]
        )
        if filename:
            string_var.set(os.path.basename(filename))

    def delete_task(self):
        selections = self.task_tree.selection()
        if not selections: messagebox.showwarning("警告", "请先选择要删除的节目"); return
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(selections)} 个节目吗？"):
            indices = sorted([self.task_tree.index(s) for s in selections], reverse=True)
            for index in indices: self.log(f"已删除节目: {self.tasks.pop(index)['name']}")
            self.update_task_list(); self.save_tasks()

    def edit_task(self):
        selection = self.task_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要修改的节目")
            return
        if len(selection) > 1:
            messagebox.showwarning("警告", "一次只能修改一个节目")
            return
        
        index = self.task_tree.index(selection[0])
        task = self.tasks[index]
        
        dummy_parent = tk.Toplevel(self.root)
        dummy_parent.withdraw()

        if task.get('type') == 'audio':
            self.open_audio_dialog(dummy_parent, task_to_edit=task, index=index)
        else:
            self.open_voice_dialog(dummy_parent, task_to_edit=task, index=index)
        
        def check_dialog_closed():
            try:
                if not dummy_parent.winfo_children():
                    dummy_parent.destroy()
                else:
                    self.root.after(100, check_dialog_closed)
            except tk.TclError:
                pass 
        self.root.after(100, check_dialog_closed)

    def copy_task(self):
        selections = self.task_tree.selection()
        if not selections: messagebox.showwarning("警告", "请先选择要复制的节目"); return
        for sel in selections:
            original = self.tasks[self.task_tree.index(sel)]
            copy = json.loads(json.dumps(original))
            copy['name'] += " (副本)"; copy['last_run'] = {}
            self.tasks.append(copy)
            self.log(f"已复制节目: {original['name']}")
        self.update_task_list(); self.save_tasks()

    def move_task(self, direction):
        sel = self.task_tree.selection()
        if not sel or len(sel) > 1: return
        index = self.task_tree.index(sel[0])
        new_index = index + direction
        if 0 <= new_index < len(self.tasks):
            self.tasks.insert(new_index, self.tasks.pop(index))
            self.update_task_list(); self.save_tasks()
            items = self.task_tree.get_children()
            if items: self.task_tree.selection_set(items[new_index])

    def import_tasks(self):
        filename = filedialog.askopenfilename(title="选择导入文件", filetypes=[("JSON文件", "*.json")])
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f: imported = json.load(f)
                self.tasks.extend(imported); self.update_task_list(); self.save_tasks()
                self.log(f"已从 {os.path.basename(filename)} 导入 {len(imported)} 个节目")
            except Exception as e: messagebox.showerror("错误", f"导入失败: {e}")

    def export_tasks(self):
        if not self.tasks: messagebox.showwarning("警告", "没有节目可以导出"); return
        filename = filedialog.asksaveasfilename(title="导出到...", defaultextension=".json",
            initialfile="broadcast_backup.json", filetypes=[("JSON文件", "*.json")])
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f: json.dump(self.tasks, f, ensure_ascii=False, indent=2)
                self.log(f"已导出 {len(self.tasks)} 个节目到 {os.path.basename(filename)}")
            except Exception as e: messagebox.showerror("错误", f"导出失败: {e}")

    def enable_task(self): self._set_task_status('启用')
    def disable_task(self): self._set_task_status('禁用')

    def _set_task_status(self, status):
        selection = self.task_tree.selection()
        if not selection: messagebox.showwarning("警告", f"请先选择要{status}的节目"); return
        count = sum(1 for i in selection if self.tasks[self.task_tree.index(i)]['status'] != status)
        for i in selection: self.tasks[self.task_tree.index(i)]['status'] = status
        if count > 0: self.update_task_list(); self.save_tasks(); self.log(f"已{status} {count} 个节目")

    def show_time_settings_dialog(self, time_entry):
        dialog = tk.Toplevel(self.root)
        dialog.title("开始时间设置"); dialog.geometry("450x400"); dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set(); dialog.configure(bg='#D7F3F5')
        self.center_window(dialog, 450, 400)
        main_frame = tk.Frame(dialog, bg='#D7F3F5', padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        tk.Label(main_frame, text="24小时制 HH:MM:SS", font=('Microsoft YaHei', 10, 'bold'),
                bg='#D7F3F5').pack(anchor='w', pady=5)
        list_frame = tk.LabelFrame(main_frame, text="时间列表", bg='#D7F3F5', padx=5, pady=5)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        box_frame = tk.Frame(list_frame)
        box_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        listbox = tk.Listbox(box_frame, font=('Microsoft YaHei', 10), height=10)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = tk.Scrollbar(box_frame, orient=tk.VERTICAL, command=listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y); listbox.configure(yscrollcommand=scrollbar.set)
        for t in [t.strip() for t in time_entry.get().split(',') if t.strip()]: listbox.insert(tk.END, t)
        btn_frame = tk.Frame(list_frame, bg='#D7F3F5')
        btn_frame.pack(side=tk.RIGHT, padx=10, fill=tk.Y)
        new_entry = tk.Entry(btn_frame, font=('Microsoft YaHei', 10), width=12)
