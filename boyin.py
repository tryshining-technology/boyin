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
        self.stop_playback_flag = threading.Event()

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

        log_frame = tk.LabelFrame(self.main_frame, text="日志：", font=('Microsoft YaHei', 10),
                                 bg='white', fg='#2C5F7C', padx=10, pady=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
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
    
    def on_double_click_edit(self, event):
        if self.task_tree.identify_row(event.y):
            self.edit_task()

    def show_context_menu(self, event):
        iid = self.task_tree.identify_row(event.y)
        is_playing = (AUDIO_AVAILABLE and pygame.mixer.music.get_busy())
        
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
        stop_state = "normal" if is_playing else "disabled"
        context_menu.add_command(label="⏹️ 停止当前播放", command=self.stop_current_playback, state=stop_state)
        
        context_menu.post(event.x_root, event.y_root)

    def play_now(self):
        selection = self.task_tree.selection()
        if not selection: return
        
        index = self.task_tree.index(selection[0])
        task = self.tasks[index]

        self.log(f"手动触发立即播放: {task['name']}")
        self._execute_broadcast(task, "manual_play")

    def stop_current_playback(self):
        self.stop_playback_flag.set()
        if AUDIO_AVAILABLE and pygame.mixer.music.get_busy():
            pygame.mixer.music.stop()
            self.log("手动停止播放。")
        self.on_playback_finished()

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
        audio_btn = tk.Button(btn_frame, text="🎵 音频节目", command=lambda: self._open_task_dialog(choice_dialog, 'audio'),
                             bg='#5DADE2', fg='white', font=('Microsoft YaHei', 11, 'bold'),
                             bd=0, padx=30, pady=12, cursor='hand2', width=15)
        audio_btn.pack(pady=8)
        voice_btn = tk.Button(btn_frame, text="🎙️ 语音节目", command=lambda: self._open_task_dialog(choice_dialog, 'voice'),
                             bg='#3498DB', fg='white', font=('Microsoft YaHei', 11, 'bold'),
                             bd=0, padx=30, pady=12, cursor='hand2', width=15)
        voice_btn.pack(pady=8)

    def _open_task_dialog(self, parent_dialog, task_type, task_to_edit=None, index=None):
        if parent_dialog:
            parent_dialog.destroy()
        is_edit_mode = task_to_edit is not None

        dialog = tk.Toplevel(self.root)
        title_prefix = "修改" if is_edit_mode else "添加"
        dialog_title = f"{title_prefix}{'音频' if task_type == 'audio' else '语音'}节目"
        dialog.title(dialog_title)
        
        dialog_geom = "850x750" if task_type == 'audio' else "800x800"
        dialog.geometry(dialog_geom)
        
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set(); dialog.configure(bg='#E8E8E8')
        
        main_frame = tk.Frame(dialog, bg='#E8E8E8', padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        if task_type == 'audio':
            dialog_vars = self._build_audio_dialog_ui(main_frame)
        else:
            dialog_vars = self._build_voice_dialog_ui(main_frame)

        if is_edit_mode:
            for key, var in dialog_vars.items():
                if key in task_to_edit:
                    var.set(task_to_edit[key])

        def save_task():
            new_task_data = {key: var.get() for key, var in dialog_vars.items()}
            
            if not new_task_data.get('name') or not new_task_data.get('time'):
                messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog); return
            
            if task_type == 'audio' and not new_task_data.get('content'):
                messagebox.showwarning("警告", "请选择音频文件或文件夹", parent=dialog); return
            
            if task_type == 'voice' and not new_task_data.get('content', '').strip():
                messagebox.showwarning("警告", "请输入播音文字", parent=dialog); return

            new_task_data['type'] = task_type
            if is_edit_mode:
                new_task_data['status'] = task_to_edit.get('status', '启用')
                new_task_data['last_run'] = task_to_edit.get('last_run', {})
            else:
                new_task_data['status'] = '启用'
                new_task_data['last_run'] = {}
            
            if is_edit_mode:
                self.tasks[index] = new_task_data
                self.log(f"已修改节目: {new_task_data['name']}")
            else:
                self.tasks.append(new_task_data)
                self.log(f"已添加节目: {new_task_data['name']}")
                
            self.update_task_list(); self.save_tasks(); dialog.destroy()

        button_frame = tk.Frame(main_frame, bg='#E8E8E8')
        button_frame.grid(row=3, column=0, pady=20)
        button_text = "保存修改" if is_edit_mode else "添加"
        tk.Button(button_frame, text=button_text, command=save_task, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 10, 'bold'), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 10), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        
    def _create_common_frames(self, parent):
        """创建通用的时间、计划、其它框架"""
        time_frame = tk.LabelFrame(parent, text="时间", font=('Microsoft YaHei', 11, 'bold'), bg='#E8E8E8', padx=10, pady=10)
        time_frame.grid(row=1, column=0, sticky='ew', pady=5)
        
        other_frame = tk.LabelFrame(parent, text="其它", font=('Microsoft YaHei', 12, 'bold'), bg='#E8E8E8', padx=15, pady=15)
        other_frame.grid(row=2, column=0, sticky='ew', pady=10)
        
        return time_frame, other_frame
        
    def _build_audio_dialog_ui(self, parent):
        content_frame = tk.LabelFrame(parent, text="内容", font=('Microsoft YaHei', 11, 'bold'),
                                     bg='#E8E8E8', padx=10, pady=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=5)
        
        vars = {
            'name': tk.StringVar(), 'content': tk.StringVar(), 'audio_type': tk.StringVar(value='single'),
            'play_order': tk.StringVar(value='sequential'), 'volume': tk.StringVar(value='80'),
            'time': tk.StringVar(), 'interval_type': tk.StringVar(value='first'),
            'interval_first': tk.StringVar(value='1'), 'interval_seconds': tk.StringVar(value='600'),
            'weekday': tk.StringVar(value='每周:1234567'), 'date_range': tk.StringVar(value='2000-01-01 ~ 2099-12-31'),
            'delay': tk.StringVar(value='ontime')
        }

        # Name
        tk.Label(content_frame, text="节目名称:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(row=0, column=0, sticky='e', padx=5, pady=5)
        tk.Entry(content_frame, textvariable=vars['name'], font=('Microsoft YaHei', 10), width=55).grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=5)

        # Audio File/Folder
        def update_content_path():
            vars['content'].set(audio_single_entry.get() if vars['audio_type'].get() == 'single' else audio_folder_entry.get())
        
        vars['audio_type'].trace_add('write', lambda *_: update_content_path())

        def create_file_selector(row, text, var_type):
            frame = tk.Frame(content_frame, bg='#E8E8E8')
            frame.grid(row=row, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
            tk.Radiobutton(frame, text="", variable=vars['audio_type'], value=var_type, bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT)
            entry = tk.Entry(frame, font=('Microsoft YaHei', 10), width=35 if var_type == 'single' else 50, state='readonly')
            entry.pack(side=tk.LEFT, padx=5)
            def selector_cmd():
                path = filedialog.askopenfilename(title=f"选择{text}", initialdir=AUDIO_FOLDER, filetypes=[("音频文件", "*.mp3 *.wav *.ogg *.flac *.m4a"), ("所有文件", "*.*")]) if var_type == 'single' else filedialog.askdirectory(title=f"选择{text}", initialdir=AUDIO_FOLDER)
                if path:
                    entry.config(state='normal'); entry.delete(0, tk.END); entry.insert(0, path); entry.config(state='readonly'); update_content_path()
            tk.Button(frame, text="选取...", command=selector_cmd, bg='#D0D0D0', font=('Microsoft YaHei', 10), bd=1, padx=15, pady=3).pack(side=tk.LEFT, padx=5)
            return entry

        tk.Label(content_frame, text="音频文件", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(row=1, column=0, sticky='e', padx=5, pady=5)
        audio_single_entry = create_file_selector(1, "音频文件", "single")

        tk.Label(content_frame, text="音频文件夹", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(row=2, column=0, sticky='e', padx=5, pady=5)
        audio_folder_entry = create_file_selector(2, "音频文件夹", "folder")

        # Play order, Volume
        play_order_frame = tk.Frame(content_frame, bg='#E8E8E8')
        play_order_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=5)
        tk.Radiobutton(play_order_frame, text="顺序播", variable=vars['play_order'], value="sequential", bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(play_order_frame, text="随机播", variable=vars['play_order'], value="random", bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT, padx=10)
        
        volume_frame = tk.Frame(content_frame, bg='#E8E8E8')
        volume_frame.grid(row=4, column=1, columnspan=3, sticky='w', padx=5, pady=8)
        tk.Label(volume_frame, text="音量:", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT)
        tk.Entry(volume_frame, textvariable=vars['volume'], font=('Microsoft YaHei', 10), width=10).pack(side=tk.LEFT, padx=5)
        tk.Label(volume_frame, text="0-100", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT, padx=5)

        time_frame, other_frame = self._create_common_frames(parent)
        self._build_time_and_schedule_ui(time_frame, vars)

        # Interval specific to audio
        interval_frame1 = tk.Frame(time_frame, bg='#E8E8E8')
        interval_frame1.grid(row=1, column=1, columnspan=2, sticky='w', padx=5, pady=5)
        tk.Label(time_frame, text="间隔播报:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(row=1, column=0, sticky='e', padx=5, pady=5)
        tk.Radiobutton(interval_frame1, text="播 n 首", variable=vars['interval_type'], value="first", bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT)
        tk.Entry(interval_frame1, textvariable=vars['interval_first'], font=('Microsoft YaHei', 10), width=15).pack(side=tk.LEFT, padx=5)
        tk.Label(interval_frame1, text="(单曲时,指 n 遍)", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT, padx=5)
        interval_frame2 = tk.Frame(time_frame, bg='#E8E8E8')
        interval_frame2.grid(row=2, column=1, columnspan=2, sticky='w', padx=5, pady=5)
        tk.Radiobutton(interval_frame2, text="播 n 秒", variable=vars['interval_type'], value="seconds", bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT)
        tk.Entry(interval_frame2, textvariable=vars['interval_seconds'], font=('Microsoft YaHei', 10), width=15).pack(side=tk.LEFT, padx=5)
        tk.Label(interval_frame2, text="(3600秒 = 1小时)", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT, padx=5)
        
        self._build_other_settings_ui(other_frame, vars)
        return vars
        
    def _build_voice_dialog_ui(self, parent):
        content_frame = tk.LabelFrame(parent, text="内容", font=('Microsoft YaHei', 11, 'bold'), bg='#E8E8E8', padx=10, pady=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=5)

        vars = {
            'name': tk.StringVar(), 'content': tk.StringVar(), 'voice': tk.StringVar(),
            'speed': tk.StringVar(value='0'), 'pitch': tk.StringVar(value='0'), 'volume': tk.StringVar(value='80'),
            'prompt': tk.IntVar(value=0), 'prompt_file': tk.StringVar(), 'prompt_volume': tk.StringVar(value='80'),
            'bgm': tk.IntVar(value=0), 'bgm_file': tk.StringVar(), 'bgm_volume': tk.StringVar(value='40'),
            'time': tk.StringVar(), 'repeat': tk.StringVar(value='1'),
            'weekday': tk.StringVar(value='每周:1234567'), 'date_range': tk.StringVar(value='2000-01-01 ~ 2099-12-31'),
            'delay': tk.StringVar(value='delay')
        }

        # Row 0, 1: Name, Content Text
        tk.Label(content_frame, text="节目名称:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(row=0, column=0, sticky='w', padx=5, pady=5)
        tk.Entry(content_frame, textvariable=vars['name'], font=('Microsoft YaHei', 10)).grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        tk.Label(content_frame, text="播音文字:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(row=1, column=0, sticky='nw', padx=5, pady=5)
        text_frame = tk.Frame(content_frame, bg='#E8E8E8')
        text_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        content_widget = scrolledtext.ScrolledText(text_frame, height=5, font=('Microsoft YaHei', 10), width=65, wrap=tk.WORD)
        content_widget.pack(fill=tk.BOTH, expand=True)
        # Link content_widget to the 'content' StringVar manually
        vars['content'].get = lambda: content_widget.get('1.0', 'end-1c')
        vars['content'].set = lambda val: content_widget.delete('1.0', tk.END) or content_widget.insert('1.0', val)

        # Row 2: Voice
        tk.Label(content_frame, text="播音员:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(row=2, column=0, sticky='w', padx=5, pady=8)
        ttk.Combobox(content_frame, textvariable=vars['voice'], values=self.get_available_voices(), font=('Microsoft YaHei', 10), width=50, state='readonly').grid(row=2, column=1, columnspan=3, sticky='w', padx=5, pady=8)
        
        # Row 3: Speech params
        speech_params_frame = tk.Frame(content_frame, bg='#E8E8E8')
        speech_params_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=5)
        tk.Label(speech_params_frame, text="语速(-10~10):", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT, padx=(0,5))
        tk.Entry(speech_params_frame, textvariable=vars['speed'], font=('Microsoft YaHei', 10), width=8).pack(side=tk.LEFT, padx=5)
        tk.Label(speech_params_frame, text="音调(-10~10):", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT, padx=(10,5))
        tk.Entry(speech_params_frame, textvariable=vars['pitch'], font=('Microsoft YaHei', 10), width=8).pack(side=tk.LEFT, padx=5)
        tk.Label(speech_params_frame, text="音量(0-100):", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT, padx=(10,5))
        tk.Entry(speech_params_frame, textvariable=vars['volume'], font=('Microsoft YaHei', 10), width=8).pack(side=tk.LEFT, padx=5)

        # Row 4, 5: Prompt, BGM
        self._create_file_picker_row(content_frame, 4, "提示音:", vars['prompt'], vars['prompt_file'], vars['prompt_volume'], PROMPT_FOLDER)
        self._create_file_picker_row(content_frame, 5, "背景音乐:", vars['bgm'], vars['bgm_file'], vars['bgm_volume'], BGM_FOLDER)

        time_frame, other_frame = self._create_common_frames(parent)
        self._build_time_and_schedule_ui(time_frame, vars, is_voice=True)
        self._build_other_settings_ui(other_frame, vars)
        return vars
        
    def _create_file_picker_row(self, parent, row, label_text, check_var, file_var, vol_var, initial_dir):
        """Helper to create a file picker row for voice dialog"""
        frame = tk.Frame(parent, bg='#E8E8E8')
        frame.grid(row=row, column=1, columnspan=3, sticky='w', padx=5, pady=5)
        tk.Checkbutton(frame, text=label_text, variable=check_var, bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT)
        tk.Entry(frame, textvariable=file_var, font=('Microsoft YaHei', 10), width=20, state='readonly').pack(side=tk.LEFT, padx=5)
        tk.Button(frame, text="...", command=lambda: self.select_file_for_entry(initial_dir, file_var, [("音频文件", "*.mp3 *.wav *.ogg")])).pack(side=tk.LEFT)
        tk.Label(frame, text="音量(0-100):", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT, padx=(10,5))
        tk.Entry(frame, textvariable=vol_var, font=('Microsoft YaHei', 10), width=8).pack(side=tk.LEFT, padx=5)

    def _build_time_and_schedule_ui(self, parent, vars, is_voice=False):
        """Builds the common time and schedule UI parts for dialogs"""
        tk.Label(parent, text="开始时间:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(row=0, column=0, sticky='e', padx=5, pady=5)
        tk.Entry(parent, textvariable=vars['time'], font=('Microsoft YaHei', 10)).grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        tk.Label(parent, text="《可多个,用英文逗号,隔开》", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(row=0, column=2, sticky='w', padx=5)
        tk.Button(parent, text="设置...", command=lambda: self.show_time_settings_dialog(vars['time']), bg='#D0D0D0', font=('Microsoft YaHei', 10), bd=1, padx=12, pady=2).grid(row=0, column=3, padx=5)
        
        if is_voice:
            tk.Label(parent, text="播 n 遍:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(row=1, column=0, sticky='e', padx=5, pady=5)
            tk.Entry(parent, textvariable=vars['repeat'], font=('Microsoft YaHei', 10), width=12).grid(row=1, column=1, sticky='w', padx=5, pady=5)
            start_row = 2
        else:
            start_row = 3 # Audio dialog has interval rows at 1 and 2

        tk.Label(parent, text="周几/几号:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(row=start_row, column=0, sticky='e', padx=5, pady=8)
        tk.Entry(parent, textvariable=vars['weekday'], font=('Microsoft YaHei', 10)).grid(row=start_row, column=1, sticky='ew', padx=5, pady=8)
        tk.Button(parent, text="选取...", command=lambda: self.show_weekday_settings_dialog(vars['weekday']), bg='#D0D0D0', font=('Microsoft YaHei', 10), bd=1, padx=15, pady=3).grid(row=start_row, column=3, padx=5)
        
        tk.Label(parent, text="日期范围:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(row=start_row + 1, column=0, sticky='e', padx=5, pady=8)
        tk.Entry(parent, textvariable=vars['date_range'], font=('Microsoft YaHei', 10)).grid(row=start_row + 1, column=1, sticky='ew', padx=5, pady=8)
        tk.Button(parent, text="设置...", command=lambda: self.show_daterange_settings_dialog(vars['date_range']), bg='#D0D0D0', font=('Microsoft YaHei', 10), bd=1, padx=15, pady=3).grid(row=start_row + 1, column=3, padx=5)

    def _build_other_settings_ui(self, parent, vars):
        """Builds the common 'Other' settings UI for dialogs"""
        tk.Label(parent, text="准时/延后:", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(row=0, column=0, sticky='e', padx=5, pady=3)
        delay_frame = tk.Frame(parent, bg='#E8E8E8')
        delay_frame.grid(row=0, column=1, sticky='w', padx=5, pady=3)
        tk.Radiobutton(delay_frame, text="准时播 - 频道内,若有别的节目正在播，终止他们", variable=vars['delay'], value="ontime", bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(anchor='w', pady=2)
        tk.Radiobutton(delay_frame, text="可延后 - 频道内,若有别的节目正在播，排队等候", variable=vars['delay'], value="delay", bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(anchor='w', pady=2)
        
    def get_available_voices(self):
        available_voices = []
        if WIN32COM_AVAILABLE:
            try:
                pythoncom.CoInitialize()
                speaker = win32com.client.Dispatch("SAPI.SpVoice")
                voices = speaker.GetVoices()
                available_voices = [v.GetDescription() for v in voices]
                self.log("成功通过 win32com 刷新语音列表。")
                pythoncom.CoUninitialize()
            except Exception as e:
                self.log(f"警告: 使用 win32com 获取语音列表失败 - {e}")
                available_voices = []
        return available_voices
    
    def select_file_for_entry(self, initial_dir, string_var, file_types):
        filename = filedialog.askopenfilename(title="选择文件", initialdir=initial_dir, filetypes=file_types)
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
        if not selection or len(selection) > 1: return
        
        index = self.task_tree.index(selection[0])
        task = self.tasks[index]
        
        dummy_parent = tk.Toplevel(self.root)
        dummy_parent.withdraw()
        self._open_task_dialog(dummy_parent, task.get('type'), task_to_edit=task, index=index)
        
        def check_dialog_closed():
            try:
                if not dummy_parent.winfo_children(): dummy_parent.destroy()
                else: self.root.after(100, check_dialog_closed)
            except tk.TclError: pass 
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

    def show_time_settings_dialog(self, time_entry_var):
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
        for t in [t.strip() for t in time_entry_var.get().split(',') if t.strip()]: listbox.insert(tk.END, t)
        btn_frame = tk.Frame(list_frame, bg='#D7F3F5')
        btn_frame.pack(side=tk.RIGHT, padx=10, fill=tk.Y)
        new_entry = tk.Entry(btn_frame, font=('Microsoft YaHei', 10), width=12)
        new_entry.insert(0, datetime.now().strftime("%H:%M:%S")); new_entry.pack(pady=3)
        def add_time():
            try:
                val = new_entry.get().strip()
                time.strptime(val, '%H:%M:%S')
                if val not in listbox.get(0, tk.END): listbox.insert(tk.END, val)
            except ValueError: messagebox.showerror("格式错误", "请输入有效的时间格式 HH:MM:SS", parent=dialog)
        def del_time():
            if listbox.curselection(): listbox.delete(listbox.curselection()[0])
        tk.Button(btn_frame, text="添加 ↑", command=add_time).pack(pady=3, fill=tk.X)
        tk.Button(btn_frame, text="删除", command=del_time).pack(pady=3, fill=tk.X)
        tk.Button(btn_frame, text="清空", command=lambda: listbox.delete(0, tk.END)).pack(pady=3, fill=tk.X)
        bottom_frame = tk.Frame(main_frame, bg='#D7F3F5')
        bottom_frame.pack(pady=10)
        def confirm():
            time_entry_var.set(", ".join(list(listbox.get(0, tk.END)))); dialog.destroy()
        tk.Button(bottom_frame, text="确定", command=confirm, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 9, 'bold'), bd=1, padx=25, pady=5).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 9), bd=1, padx=25, pady=5).pack(side=tk.LEFT, padx=5)

    def show_weekday_settings_dialog(self, weekday_var):
        dialog = tk.Toplevel(self.root); dialog.title("周几或几号")
        dialog.geometry("500x520"); dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set(); dialog.configure(bg='#D7F3F5')
        self.center_window(dialog, 500, 520)
        main_frame = tk.Frame(dialog, bg='#D7F3F5', padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        week_type_var = tk.StringVar(value="week")
        week_frame = tk.LabelFrame(main_frame, text="按周", font=('Microsoft YaHei', 10, 'bold'),
                                  bg='#D7F3F5', padx=10, pady=10)
        week_frame.pack(fill=tk.X, pady=5)
        tk.Radiobutton(week_frame, text="每周", variable=week_type_var, value="week",
                      bg='#D7F3F5', font=('Microsoft YaHei', 10)).grid(row=0, column=0, sticky='w')
        weekdays = [("周一", 1), ("周二", 2), ("周三", 3), ("周四", 4), ("周五", 5), ("周六", 6), ("周日", 7)]
        week_vars = {num: tk.IntVar(value=1) for day, num in weekdays}
        for i, (day, num) in enumerate(weekdays):
            tk.Checkbutton(week_frame, text=day, variable=week_vars[num], bg='#D7F3F5',
                          font=('Microsoft YaHei', 10)).grid(row=(i // 4) + 1, column=i % 4, sticky='w', padx=10, pady=3)
        day_frame = tk.LabelFrame(main_frame, text="按月", font=('Microsoft YaHei', 10, 'bold'),
                                 bg='#D7F3F5', padx=10, pady=10)
        day_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        tk.Radiobutton(day_frame, text="每月", variable=week_type_var, value="day",
                      bg='#D7F3F5', font=('Microsoft YaHei', 10)).grid(row=0, column=0, sticky='w')
        day_vars = {i: tk.IntVar(value=0) for i in range(1, 32)}
        for i in range(1, 32):
            tk.Checkbutton(day_frame, text=f"{i:02d}", variable=day_vars[i], bg='#D7F3F5',
                          font=('Microsoft YaHei', 10)).grid(row=((i - 1) // 7) + 1, column=(i - 1) % 7, sticky='w', padx=8, pady=2)
        bottom_frame = tk.Frame(main_frame, bg='#D7F3F5')
        bottom_frame.pack(pady=10)
        def confirm():
            if week_type_var.get() == "week":
                selected = sorted([str(n) for n, v in week_vars.items() if v.get()])
                result = "每周:" + "".join(selected)
            else:
                selected = sorted([f"{n:02d}" for n, v in day_vars.items() if v.get()])
                result = "每月:" + ",".join(selected)
            weekday_var.set(result if selected else ""); dialog.destroy()
        tk.Button(bottom_frame, text="确定", command=confirm, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 9, 'bold'), bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 9), bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)

    def show_daterange_settings_dialog(self, date_range_var):
        dialog = tk.Toplevel(self.root)
        dialog.title("日期范围"); dialog.geometry("450x220"); dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set(); dialog.configure(bg='#D7F3F5')
        self.center_window(dialog, 450, 220)
        main_frame = tk.Frame(dialog, bg='#D7F3F5', padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        from_frame = tk.Frame(main_frame, bg='#D7F3F5')
        from_frame.pack(pady=10, anchor='w')
        tk.Label(from_frame, text="从", font=('Microsoft YaHei', 10, 'bold'), bg='#D7F3F5').pack(side=tk.LEFT, padx=5)
        from_date_entry = tk.Entry(from_frame, font=('Microsoft YaHei', 10), width=18)
        from_date_entry.pack(side=tk.LEFT, padx=5)
        to_frame = tk.Frame(main_frame, bg='#D7F3F5')
        to_frame.pack(pady=10, anchor='w')
        tk.Label(to_frame, text="到", font=('Microsoft YaHei', 10, 'bold'), bg='#D7F3F5').pack(side=tk.LEFT, padx=5)
        to_date_entry = tk.Entry(to_frame, font=('Microsoft YaHei', 10), width=18)
        to_date_entry.pack(side=tk.LEFT, padx=5)
        try:
            start, end = date_range_var.get().split('~')
            from_date_entry.insert(0, start.strip()); to_date_entry.insert(0, end.strip())
        except (ValueError, IndexError):
            from_date_entry.insert(0, "2000-01-01"); to_date_entry.insert(0, "2099-12-31")
        tk.Label(main_frame, text="格式: YYYY-MM-DD", font=('Microsoft YaHei', 10),
                bg='#D7F3F5', fg='#666').pack(pady=10)
        bottom_frame = tk.Frame(main_frame, bg='#D7F3F5')
        bottom_frame.pack(pady=10)
        def confirm():
            try:
                start, end = from_date_entry.get().strip(), to_date_entry.get().strip()
                datetime.strptime(start, "%Y-%m-%d"); datetime.strptime(end, "%Y-%m-%d")
                date_range_var.set(f"{start} ~ {end}"); dialog.destroy()
            except ValueError: messagebox.showerror("格式错误", "日期格式不正确, 应为 YYYY-MM-DD", parent=dialog)
        tk.Button(bottom_frame, text="确定", command=confirm, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 9, 'bold'), bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 9), bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)
        
    # ... (其他辅助函数和主逻辑函数，请从下面完整代码中获取)
    

if __name__ == "__main__":
    root = tk.Tk()
    app = TimedBroadcastApp(root)
    root.mainloop()
