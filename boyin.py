import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog, simpledialog
import json
import threading
import time
from datetime import datetime
import os
import random
import sys
import getpass
import base64

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
    print("警告: pywin32 未安装，语音和开机启动功能将受限。")

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
SETTINGS_FILE = os.path.join(application_path, "settings.json")
HOLIDAY_FILE = os.path.join(application_path, "holidays.json")
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
        self.holidays = []
        self.settings = {}
        self.running = True
        self.tray_icon = None
        self.is_locked = False
        
        self.drag_start_item = None

        self.is_playing = threading.Event()
        self.playback_queue = []
        self.queue_lock = threading.Lock()
        
        self.playback_lock = threading.Lock()
        self.current_playback_token = None
        
        self.pages = {}
        self.nav_buttons = {}
        self.current_page = None

        self.create_folder_structure()
        self.load_settings()
        self.create_widgets()
        self.load_tasks()
        self.load_holidays()
        
        self.start_background_thread()
        self.root.protocol("WM_DELETE_WINDOW", self.show_quit_dialog)
        self.start_tray_icon_thread()
        
        if self.settings.get("lock_on_start", False) and self.settings.get("lock_password_b64", ""):
            self.root.after(100, self.perform_initial_lock)

        if self.settings.get("start_minimized", False):
            self.root.after(100, self.hide_to_tray)

    def create_folder_structure(self):
        for folder in [PROMPT_FOLDER, AUDIO_FOLDER, BGM_FOLDER]:
            if not os.path.exists(folder):
                os.makedirs(folder)

    def create_widgets(self):
        self.nav_frame = tk.Frame(self.root, bg='#A8D8E8', width=160)
        self.nav_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.nav_frame.pack_propagate(False)

        nav_button_titles = ["定时广播", "节假日", "设置"]
        
        for i, title in enumerate(nav_button_titles):
            btn_frame = tk.Frame(self.nav_frame, bg='#A8D8E8')
            btn_frame.pack(fill=tk.X, pady=1)
            btn = tk.Button(btn_frame, text=title, bg='#A8D8E8',
                          fg='black', font=('Microsoft YaHei', 22, 'bold'),
                          bd=0, padx=10, pady=8, anchor='w', command=lambda t=title: self.switch_page(t))
            btn.pack(fill=tk.X)
            self.nav_buttons[title] = btn
        
        self.main_frame = tk.Frame(self.root, bg='white')
        self.pages["定时广播"] = self.main_frame
        self.create_scheduled_broadcast_page()

        self.current_page = self.main_frame
        self.switch_page("定时广播")

    def switch_page(self, page_name):
        if self.is_locked:
            self.log("界面已锁定，请先解锁。")
            return
            
        if self.current_page:
            self.current_page.pack_forget()

        for title, btn in self.nav_buttons.items():
            btn.config(bg='#A8D8E8', fg='black')
            btn.master.config(bg='#A8D8E8')

        target_frame = None
        if page_name == "定时广播":
            target_frame = self.pages["定时广播"]
        elif page_name == "节假日":
            if page_name not in self.pages:
                self.pages[page_name] = self.create_holiday_page()
            target_frame = self.pages[page_name]
        elif page_name == "设置":
            if page_name not in self.pages:
                self.pages[page_name] = self.create_settings_page()
            target_frame = self.pages[page_name]
        else:
            self.log(f"功能开发中: {page_name}")
            target_frame = self.pages["定时广播"]
            page_name = "定时广播"

        target_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.current_page = target_frame
        
        selected_btn = self.nav_buttons[page_name]
        selected_btn.config(bg='#5DADE2', fg='white')
        selected_btn.master.config(bg='#5DADE2')

    def create_scheduled_broadcast_page(self):
        page_frame = self.pages["定时广播"]
        font_11 = ('Microsoft YaHei', 11)

        top_frame = tk.Frame(page_frame, bg='white')
        top_frame.pack(fill=tk.X, padx=10, pady=10)
        title_label = tk.Label(top_frame, text="定时广播", font=('Microsoft YaHei', 14, 'bold'),
                              bg='white', fg='#2C5F7C')
        title_label.pack(side=tk.LEFT)
        
        add_btn = tk.Button(top_frame, text="添加节目", command=self.add_task, bg='#3498DB', fg='white',
                              font=font_11, bd=0, padx=12, pady=5, cursor='hand2')
        add_btn.pack(side=tk.LEFT, padx=10)

        self.top_right_btn_frame = tk.Frame(top_frame, bg='white')
        self.top_right_btn_frame.pack(side=tk.RIGHT)

        batch_buttons = [
            ("全部启用", self.enable_all_tasks, '#27AE60'),
            ("全部禁用", self.disable_all_tasks, '#F39C12'),
            ("统一音量", self.set_uniform_volume, '#8E44AD'),
            ("清空节目", self.clear_all_tasks, '#C0392B')
        ]
        for text, cmd, color in batch_buttons:
            btn = tk.Button(self.top_right_btn_frame, text=text, command=cmd, bg=color, fg='white',
                          font=font_11, bd=0, padx=12, pady=5, cursor='hand2')
            btn.pack(side=tk.LEFT, padx=3)

        self.lock_button = tk.Button(self.top_right_btn_frame, text="锁定", command=self.toggle_lock_state, bg='#E74C3C', fg='white',
                                     font=font_11, bd=0, padx=12, pady=5, cursor='hand2')
        self.lock_button.pack(side=tk.LEFT, padx=3)

        io_buttons = [("导入节目单", self.import_tasks, '#1ABC9C'), ("导出节目单", self.export_tasks, '#1ABC9C')]
        for text, cmd, color in io_buttons:
            btn = tk.Button(self.top_right_btn_frame, text=text, command=cmd, bg=color, fg='white',
                          font=font_11, bd=0, padx=12, pady=5, cursor='hand2')
            btn.pack(side=tk.LEFT, padx=3)

        stats_frame = tk.Frame(page_frame, bg='#F0F8FF')
        stats_frame.pack(fill=tk.X, padx=10, pady=5)
        self.stats_label = tk.Label(stats_frame, text="节目单：0", font=('Microsoft YaHei', 11),
                                   bg='#F0F8FF', fg='#2C5F7C', anchor='w', padx=10)
        self.stats_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        table_frame = tk.Frame(page_frame, bg='white')
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        columns = ('节目名称', '状态', '开始时间', '模式', '音频或文字', '音量', '周几/几号', '日期范围')
        self.task_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=12, selectmode='extended')
        
        style = ttk.Style()
        style.configure("Treeview.Heading", font=('Microsoft YaHei', 11, 'bold'))
        style.configure("Treeview", font=('Microsoft YaHei', 11), rowheight=28)

        self.task_tree.heading('节目名称', text='节目名称')
        self.task_tree.column('节目名称', width=200, anchor='w')
        self.task_tree.heading('状态', text='状态')
        self.task_tree.column('状态', width=70, anchor='center', stretch=tk.NO)
        self.task_tree.heading('开始时间', text='开始时间')
        self.task_tree.column('开始时间', width=100, anchor='center', stretch=tk.NO)
        self.task_tree.heading('模式', text='模式')
        self.task_tree.column('模式', width=70, anchor='center', stretch=tk.NO)
        self.task_tree.heading('音频或文字', text='音频或文字')
        self.task_tree.column('音频或文字', width=300, anchor='w')
        self.task_tree.heading('音量', text='音量')
        self.task_tree.column('音量', width=70, anchor='center', stretch=tk.NO)
        self.task_tree.heading('周几/几号', text='周几/几号')
        self.task_tree.column('周几/几号', width=100, anchor='center')
        self.task_tree.heading('日期范围', text='日期范围')
        self.task_tree.column('日期范围', width=120, anchor='center')

        self.task_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.task_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.task_tree.configure(yscrollcommand=scrollbar.set)
        
        self.task_tree.bind("<Button-3>", self.show_context_menu)
        self.task_tree.bind("<Double-1>", self.on_double_click_edit)
        self._enable_drag_selection(self.task_tree)

        playing_frame = tk.LabelFrame(page_frame, text="正在播：", font=('Microsoft YaHei', 11),
                                     bg='white', fg='#2C5F7C', padx=10, pady=5)
        playing_frame.pack(fill=tk.X, padx=10, pady=5)
        self.playing_text = scrolledtext.ScrolledText(playing_frame, height=3, font=('Microsoft YaHei', 11),
                                                     bg='#FFFEF0', wrap=tk.WORD, state='disabled')
        self.playing_text.pack(fill=tk.BOTH, expand=True)
        self.update_playing_text("等待播放...")

        log_frame = tk.LabelFrame(page_frame, text="", font=('Microsoft YaHei', 11),
                                 bg='white', fg='#2C5F7C', padx=10, pady=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        log_header_frame = tk.Frame(log_frame, bg='white')
        log_header_frame.pack(fill=tk.X)
        log_label = tk.Label(log_header_frame, text="日志：", font=('Microsoft YaHei', 11, 'bold'),
                             bg='white', fg='#2C5F7C')
        log_label.pack(side=tk.LEFT)
        self.clear_log_btn = tk.Button(log_header_frame, text="清除日志", command=self.clear_log,
                                       font=('Microsoft YaHei', 8), bd=0, bg='#EAEAEA',
                                       fg='#333', cursor='hand2', padx=5, pady=0)
        self.clear_log_btn.pack(side=tk.LEFT, padx=10)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, font=('Microsoft YaHei', 11),
                                                 bg='#F9F9F9', wrap=tk.WORD, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True)

        status_frame = tk.Frame(page_frame, bg='#E8F4F8', height=30)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        status_frame.pack_propagate(False)
        self.status_labels = []
        status_texts = ["当前时间", "系统状态", "播放状态", "任务数量"]
        for i, text in enumerate(status_texts):
            label = tk.Label(status_frame, text=f"{text}: --", font=font_11,
                           bg='#5DADE2' if i % 2 == 0 else '#7EC8E3', fg='white', padx=15, pady=5)
            label.pack(side=tk.LEFT, padx=2)
            self.status_labels.append(label)

        self.update_status_bar()
        self.log("定时播音软件已启动")
    
    def create_holiday_page(self):
        page_frame = tk.Frame(self.root, bg='white')

        top_frame = tk.Frame(page_frame, bg='white')
        top_frame.pack(fill=tk.X, padx=10, pady=10)
        title_label = tk.Label(top_frame, text="节假日", font=('Microsoft YaHei', 14, 'bold'),
                              bg='white', fg='#2C5F7C')
        title_label.pack(side=tk.LEFT)
        
        desc_label = tk.Label(page_frame, text="节假日不播放 (手动和立即播任务除外)", font=('Microsoft YaHei', 11),
                              bg='white', fg='#555')
        desc_label.pack(anchor='w', padx=10, pady=(0, 10))

        content_frame = tk.Frame(page_frame, bg='white')
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        table_frame = tk.Frame(content_frame, bg='white')
        table_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        columns = ('节假日名称', '状态', '开始日期时间', '结束日期时间')
        self.holiday_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15, selectmode='extended')
        
        self.holiday_tree.heading('节假日名称', text='节假日名称')
        self.holiday_tree.column('节假日名称', width=250, anchor='w')
        self.holiday_tree.heading('状态', text='状态')
        self.holiday_tree.column('状态', width=100, anchor='center')
        self.holiday_tree.heading('开始日期时间', text='开始日期时间')
        self.holiday_tree.column('开始日期时间', width=200, anchor='center')
        self.holiday_tree.heading('结束日期时间', text='结束日期时间')
        self.holiday_tree.column('结束日期时间', width=200, anchor='center')

        self.holiday_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.holiday_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.holiday_tree.configure(yscrollcommand=scrollbar.set)
        
        self.holiday_tree.bind("<Double-1>", lambda e: self.edit_holiday())
        self.holiday_tree.bind("<Button-3>", self.show_holiday_context_menu)
        self._enable_drag_selection(self.holiday_tree)

        action_frame = tk.Frame(content_frame, bg='white', padx=10)
        action_frame.pack(side=tk.RIGHT, fill=tk.Y)

        btn_font = ('Microsoft YaHei', 11)
        btn_width = 10 
        
        buttons_config = [
            ("添加", self.add_holiday),
            ("修改", self.edit_holiday),
            ("删除", self.delete_holiday),
            (None, None),
            ("全部启用", self.enable_all_holidays),
            ("全部禁用", self.disable_all_holidays),
            (None, None),
            ("导入节日", self.import_holidays),
            ("导出节日", self.export_holidays),
            ("清空节日", self.clear_all_holidays),
        ]

        for text, cmd in buttons_config:
            if text is None:
                tk.Frame(action_frame, height=20, bg='white').pack()
                continue
            
            tk.Button(action_frame, text=text, command=cmd, font=btn_font, width=btn_width, pady=5).pack(pady=5)

        self.update_holiday_list()
        return page_frame

    def create_settings_page(self):
        settings_frame = tk.Frame(self.root, bg='white')

        title_label = tk.Label(settings_frame, text="系统设置", font=('Microsoft YaHei', 14, 'bold'),
                               bg='white', fg='#2C5F7C')
        title_label.pack(anchor='w', padx=20, pady=20)

        general_frame = tk.LabelFrame(settings_frame, text="通用设置", font=('Microsoft YaHei', 12, 'bold'),
                                      bg='white', padx=15, pady=10)
        general_frame.pack(fill=tk.X, padx=20, pady=10)
        
        self.autostart_var = tk.BooleanVar(value=self.settings.get("autostart", False))
        self.start_minimized_var = tk.BooleanVar(value=self.settings.get("start_minimized", False))
        self.lock_on_start_var = tk.BooleanVar(value=self.settings.get("lock_on_start", False))

        tk.Checkbutton(general_frame, text="登录windows后自动启动", variable=self.autostart_var, 
                       font=('Microsoft YaHei', 11), bg='white', anchor='w', 
                       command=self._handle_autostart_setting).pack(fill=tk.X, pady=5)
        tk.Checkbutton(general_frame, text="启动后最小化到系统托盘", variable=self.start_minimized_var,
                       font=('Microsoft YaHei', 11), bg='white', anchor='w',
                       command=self.save_settings).pack(fill=tk.X, pady=5)
        
        lock_frame = tk.Frame(general_frame, bg='white')
        lock_frame.pack(fill=tk.X, pady=5)
        
        self.lock_on_start_cb = tk.Checkbutton(lock_frame, text="启动软件后立即锁定", variable=self.lock_on_start_var,
                       font=('Microsoft YaHei', 11), bg='white', anchor='w',
                       command=self._handle_lock_on_start_toggle)
        self.lock_on_start_cb.pack(side=tk.LEFT)
        
        tk.Label(lock_frame, text="(请先在主界面设置锁定密码)", font=('Microsoft YaHei', 9), bg='white', fg='grey').pack(side=tk.LEFT, padx=5)

        self.clear_password_btn = tk.Button(general_frame, text="清除锁定密码", font=('Microsoft YaHei', 11),
                                              command=self.handle_settings_clear_password)
        self.clear_password_btn.pack(pady=10)
        if not self.settings.get("lock_password_b64"):
            self.clear_password_btn.config(state=tk.DISABLED)

        power_frame = tk.LabelFrame(settings_frame, text="电源管理", font=('Microsoft YaHei', 12, 'bold'),
                                    bg='white', padx=15, pady=10)
        power_frame.pack(fill=tk.X, padx=20, pady=10)

        self.daily_shutdown_enabled_var = tk.BooleanVar(value=self.settings.get("daily_shutdown_enabled", False))
        self.daily_shutdown_time_var = tk.StringVar(value=self.settings.get("daily_shutdown_time", "23:00:00"))
        self.weekly_shutdown_enabled_var = tk.BooleanVar(value=self.settings.get("weekly_shutdown_enabled", False))
        self.weekly_shutdown_time_var = tk.StringVar(value=self.settings.get("weekly_shutdown_time", "23:30:00"))
        self.weekly_shutdown_days_var = tk.StringVar(value=self.settings.get("weekly_shutdown_days", "每周:12345"))
        self.weekly_reboot_enabled_var = tk.BooleanVar(value=self.settings.get("weekly_reboot_enabled", False))
        self.weekly_reboot_time_var = tk.StringVar(value=self.settings.get("weekly_reboot_time", "22:00:00"))
        self.weekly_reboot_days_var = tk.StringVar(value=self.settings.get("weekly_reboot_days", "每周:67"))

        daily_frame = tk.Frame(power_frame, bg='white')
        daily_frame.pack(fill=tk.X, pady=4)
        tk.Checkbutton(daily_frame, text="每天关机", variable=self.daily_shutdown_enabled_var, 
                       font=('Microsoft YaHei', 11), bg='white', command=self.save_settings).pack(side=tk.LEFT)
        time_entry_daily = tk.Entry(daily_frame, textvariable=self.daily_shutdown_time_var, 
                                    font=('Microsoft YaHei', 11), width=15)
        time_entry_daily.pack(side=tk.LEFT, padx=10)
        tk.Button(daily_frame, text="设置", font=('Microsoft YaHei', 11), command=lambda: self.show_single_time_dialog(self.daily_shutdown_time_var)
                  ).pack(side=tk.LEFT)

        weekly_frame = tk.Frame(power_frame, bg='white')
        weekly_frame.pack(fill=tk.X, pady=4)
        tk.Checkbutton(weekly_frame, text="每周关机", variable=self.weekly_shutdown_enabled_var, 
                       font=('Microsoft YaHei', 11), bg='white', command=self.save_settings).pack(side=tk.LEFT)
        days_entry_weekly = tk.Entry(weekly_frame, textvariable=self.weekly_shutdown_days_var,
                                     font=('Microsoft YaHei', 11), width=20)
        days_entry_weekly.pack(side=tk.LEFT, padx=(10,5))
        time_entry_weekly = tk.Entry(weekly_frame, textvariable=self.weekly_shutdown_time_var,
                                     font=('Microsoft YaHei', 11), width=15)
        time_entry_weekly.pack(side=tk.LEFT, padx=5)
        tk.Button(weekly_frame, text="设置", font=('Microsoft YaHei', 11), command=lambda: self.show_power_week_time_dialog(
            "设置每周关机", self.weekly_shutdown_days_var, self.weekly_shutdown_time_var)).pack(side=tk.LEFT)

        reboot_frame = tk.Frame(power_frame, bg='white')
        reboot_frame.pack(fill=tk.X, pady=4)
        tk.Checkbutton(reboot_frame, text="每周重启", variable=self.weekly_reboot_enabled_var,
                       font=('Microsoft YaHei', 11), bg='white', command=self.save_settings).pack(side=tk.LEFT)
        days_entry_reboot = tk.Entry(reboot_frame, textvariable=self.weekly_reboot_days_var,
                                     font=('Microsoft YaHei', 11), width=20)
        days_entry_reboot.pack(side=tk.LEFT, padx=(10,5))
        time_entry_reboot = tk.Entry(reboot_frame, textvariable=self.weekly_reboot_time_var,
                                     font=('Microsoft YaHei', 11), width=15)
        time_entry_reboot.pack(side=tk.LEFT, padx=5)
        tk.Button(reboot_frame, text="设置", font=('Microsoft YaHei', 11), command=lambda: self.show_power_week_time_dialog(
            "设置每周重启", self.weekly_reboot_days_var, self.weekly_reboot_time_var)).pack(side=tk.LEFT)

        return settings_frame

    def toggle_lock_state(self):
        if self.is_locked:
            self._prompt_for_password_unlock()
        else:
            password_b64 = self.settings.get("lock_password_b64", "")
            if not password_b64:
                self._prompt_for_password_set()
            else:
                self._apply_lock()

    def _apply_lock(self):
        self.is_locked = True
        self.lock_button.config(text="解锁", bg='#2ECC71')
        self._set_ui_lock_state(tk.DISABLED)
        self.log("界面已锁定。")

    def _apply_unlock(self):
        self.is_locked = False
        self.lock_button.config(text="锁定", bg='#E74C3C')
        self._set_ui_lock_state(tk.NORMAL)
        self.log("界面已解锁。")

    def perform_initial_lock(self):
        self.log("根据设置，软件启动时自动锁定。")
        self._apply_lock()

    def _prompt_for_password_set(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("首次锁定，请设置密码")
        dialog.geometry("350x250"); dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()
        self.center_window(dialog, 350, 250)
        
        tk.Label(dialog, text="请设置一个锁定密码 (最多6位)", font=('Microsoft YaHei', 11)).pack(pady=10)
        
        tk.Label(dialog, text="输入密码:", font=('Microsoft YaHei', 11)).pack(pady=(5,0))
        pass_entry1 = tk.Entry(dialog, show='*', width=25, font=('Microsoft YaHei', 11))
        pass_entry1.pack()

        tk.Label(dialog, text="确认密码:", font=('Microsoft YaHei', 11)).pack(pady=(10,0))
        pass_entry2 = tk.Entry(dialog, show='*', width=25, font=('Microsoft YaHei', 11))
        pass_entry2.pack()

        def confirm():
            p1 = pass_entry1.get()
            p2 = pass_entry2.get()
            if not p1: messagebox.showerror("错误", "密码不能为空。", parent=dialog); return
            if len(p1) > 6: messagebox.showerror("错误", "密码不能超过6位。", parent=dialog); return
            if p1 != p2: messagebox.showerror("错误", "两次输入的密码不一致。", parent=dialog); return
            
            encoded_pass = base64.b64encode(p1.encode('utf-8')).decode('utf-8')
            self.settings["lock_password_b64"] = encoded_pass
            self.save_settings()
            
            if "设置" in self.pages and hasattr(self, 'clear_password_btn'):
                self.clear_password_btn.config(state=tk.NORMAL)

            messagebox.showinfo("成功", "密码设置成功，界面即将锁定。", parent=dialog)
            dialog.destroy()
            self._apply_lock()

        btn_frame = tk.Frame(dialog); btn_frame.pack(pady=20)
        tk.Button(btn_frame, text="确定", command=confirm, font=('Microsoft YaHei', 11)).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="取消", command=dialog.destroy, font=('Microsoft YaHei', 11)).pack(side=tk.LEFT, padx=10)

    def _prompt_for_password_unlock(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("解锁界面")
        dialog.geometry("400x180"); dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()
        self.center_window(dialog, 400, 180)
        
        tk.Label(dialog, text="请输入密码以解锁", font=('Microsoft YaHei', 11)).pack(pady=10)
        
        pass_entry = tk.Entry(dialog, show='*', width=25, font=('Microsoft YaHei', 11))
        pass_entry.pack(pady=5)
        pass_entry.focus_set()

        def is_password_correct():
            entered_pass = pass_entry.get()
            encoded_entered_pass = base64.b64encode(entered_pass.encode('utf-8')).decode('utf-8')
            return encoded_entered_pass == self.settings.get("lock_password_b64")

        def confirm():
            if is_password_correct():
                dialog.destroy()
                self._apply_unlock()
            else:
                messagebox.showerror("错误", "密码不正确！", parent=dialog)
        
        def clear_password_action():
            if not is_password_correct():
                messagebox.showerror("错误", "密码不正确！无法清除。", parent=dialog)
                return
            
            if messagebox.askyesno("确认操作", "您确定要清除锁定密码吗？\n此操作不可恢复。", parent=dialog):
                self._clear_password_logic()
                messagebox.showinfo("成功", "锁定密码已成功清除。", parent=dialog)
                dialog.destroy()
                self._apply_unlock()

        btn_frame = tk.Frame(dialog); btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="确定", command=confirm, font=('Microsoft YaHei', 11)).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="清除密码", command=clear_password_action, font=('Microsoft YaHei', 11)).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="取消", command=dialog.destroy, font=('Microsoft YaHei', 11)).pack(side=tk.LEFT, padx=5)
        dialog.bind('<Return>', lambda event: confirm())

    def _clear_password_logic(self):
        """Core logic for clearing password, without UI interaction."""
        self.settings["lock_password_b64"] = ""
        self.lock_on_start_var.set(False)
        self.save_settings()
        if hasattr(self, 'clear_password_btn'):
            self.clear_password_btn.config(state=tk.DISABLED)
        self.log("锁定密码已清除。")

    def handle_settings_clear_password(self):
        """Handler for the 'Clear Password' button on the settings page."""
        if messagebox.askyesno("确认操作", "您确定要清除锁定密码吗？\n此操作不可恢复。"):
            self._clear_password_logic()
            messagebox.showinfo("成功", "锁定密码已成功清除。")

    def _handle_lock_on_start_toggle(self):
        if not self.settings.get("lock_password_b64"):
            if self.lock_on_start_var.get():
                messagebox.showwarning("无法启用", "您还未设置锁定密码。\n\n请返回“定时广播”页面，点击“锁定”按钮来首次设置密码。")
                self.root.after(50, lambda: self.lock_on_start_var.set(False))
        else:
            self.save_settings()

    def _set_ui_lock_state(self, state):
        self._set_widget_state_recursively(self.nav_frame, state)
        for page in self.pages.values():
             if page and page.winfo_exists():
                self._set_widget_state_recursively(page, state)
    
    def _set_widget_state_recursively(self, parent_widget, state):
        for child in parent_widget.winfo_children():
            if child == self.lock_button:
                continue
            
            try:
                if isinstance(child, (ttk.Widget, ttk.Treeview)):
                    child.state(['disabled'] if state == tk.DISABLED else ['!disabled'])
                else:
                    child.config(state=state)
            except tk.TclError:
                pass
            
            if child.winfo_children():
                self._set_widget_state_recursively(child, state)
    
    def clear_log(self):
        if messagebox.askyesno("确认操作", "您确定要清空所有日志记录吗？\n此操作不可恢复。"):
            self.log_text.config(state='normal')
            self.log_text.delete('1.0', tk.END)
            self.log_text.config(state='disabled')
            self.log("日志已清空。")

    def on_double_click_edit(self, event):
        if self.is_locked: return
        if self.task_tree.identify_row(event.y):
            self.edit_task()

    def show_context_menu(self, event):
        if self.is_locked: return
        iid = self.task_tree.identify_row(event.y)
        context_menu = tk.Menu(self.root, tearoff=0, font=('Microsoft YaHei', 11))

        if iid:
            if iid not in self.task_tree.selection():
                self.task_tree.selection_set(iid)
            
            context_menu.add_command(label="立即播放", command=self.play_now)
            context_menu.add_separator()
            context_menu.add_command(label="修改", command=self.edit_task)
            context_menu.add_command(label="删除", command=self.delete_task)
            context_menu.add_command(label="复制", command=self.copy_task)
            context_menu.add_separator()
            context_menu.add_command(label="置顶", command=self.move_task_to_top)
            context_menu.add_command(label="上移", command=lambda: self.move_task(-1))
            context_menu.add_command(label="下移", command=lambda: self.move_task(1))
            context_menu.add_command(label="置末", command=self.move_task_to_bottom)
            context_menu.add_separator()
            context_menu.add_command(label="启用", command=self.enable_task)
            context_menu.add_command(label="禁用", command=self.disable_task)
        else:
            self.task_tree.selection_set()
            context_menu.add_command(label="添加节目", command=self.add_task)
        
        context_menu.add_separator()
        context_menu.add_command(label="停止当前播放", command=self.stop_current_playback)
        context_menu.post(event.x_root, event.y_root)

    def _stop_and_cleanup_playback(self):
        with self.playback_lock:
            self.current_playback_token = None
            if not self.is_playing.is_set():
                return
            if AUDIO_AVAILABLE:
                pygame.mixer.music.stop()
                pygame.mixer.stop()
            self.is_playing.clear()
            self.root.after(0, lambda: self.status_labels[2].config(text="播放状态: 待机"))
            self.root.after(0, lambda: self.update_playing_text("等待播放..."))
            self.log("播放已停止。")

    def play_now(self):
        selection = self.task_tree.selection()
        if not selection: messagebox.showwarning("提示", "请先选择一个要立即播放的节目。"); return
        index = self.task_tree.index(selection[0])
        task = self.tasks[index]
        self.play_task_immediately(task)
        
    def play_task_immediately(self, task):
        self._stop_and_cleanup_playback()
        with self.queue_lock:
            self.playback_queue.clear()
            self.playback_queue.insert(0, (task, "manual_play"))
            self.log(f"手动触发高优先级播放: {task['name']}")
        self.root.after(10, self._process_queue)

    def stop_current_playback(self):
        self.log("手动触发“停止当前播放”...")
        with self.queue_lock:
            self.playback_queue.clear()
            self.log("等待播放的队列也已清空。")
        self._stop_and_cleanup_playback()

    def add_task(self):
        choice_dialog = tk.Toplevel(self.root)
        choice_dialog.title("选择节目类型")
        choice_dialog.geometry("350x280")
        choice_dialog.resizable(False, False)
        choice_dialog.transient(self.root); choice_dialog.grab_set()
        self.center_window(choice_dialog, 350, 280)
        main_frame = tk.Frame(choice_dialog, padx=20, pady=20, bg='#F0F0F0')
        main_frame.pack(fill=tk.BOTH, expand=True)
        title_label = tk.Label(main_frame, text="请选择要添加的节目类型",
                              font=('Microsoft YaHei', 13, 'bold'), fg='#2C5F7C', bg='#F0F0F0')
        title_label.pack(pady=15)
        btn_frame = tk.Frame(main_frame, bg='#F0F0F0')
        btn_frame.pack(expand=True)
        audio_btn = tk.Button(btn_frame, text="🎵 音频节目",
                             bg='#5DADE2', fg='white', font=('Microsoft YaHei', 12, 'bold'),
                             bd=0, padx=30, pady=12, cursor='hand2', width=15, command=lambda: self.open_audio_dialog(choice_dialog))
        audio_btn.pack(pady=8)
        voice_btn = tk.Button(btn_frame, text="🎙️ 语音节目",
                             bg='#3498DB', fg='white', font=('Microsoft YaHei', 12, 'bold'),
                             bd=0, padx=30, pady=12, cursor='hand2', width=15, command=lambda: self.open_voice_dialog(choice_dialog))
        voice_btn.pack(pady=8)

    def open_audio_dialog(self, parent_dialog, task_to_edit=None, index=None):
        parent_dialog.destroy()
        is_edit_mode = task_to_edit is not None
        dialog = tk.Toplevel(self.root)
        dialog.title("修改音频节目" if is_edit_mode else "添加音频节目")
        dialog.geometry("950x850")
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set(); dialog.configure(bg='#E8E8E8')
        main_frame = tk.Frame(dialog, bg='#E8E8E8', padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        content_frame = tk.LabelFrame(main_frame, text="内容", font=('Microsoft YaHei', 12, 'bold'),
                                     bg='#E8E8E8', padx=10, pady=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=5)
        font_spec = ('Microsoft YaHei', 11)
        
        tk.Label(content_frame, text="节目名称:", font=font_spec, bg='#E8E8E8').grid(row=0, column=0, sticky='e', padx=5, pady=5)
        name_entry = tk.Entry(content_frame, font=font_spec, width=55)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        audio_type_var = tk.StringVar(value="single")
        tk.Label(content_frame, text="音频文件", font=font_spec, bg='#E8E8E8').grid(row=1, column=0, sticky='e', padx=5, pady=5)
        audio_single_frame = tk.Frame(content_frame, bg='#E8E8E8')
        audio_single_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        tk.Radiobutton(audio_single_frame, text="", variable=audio_type_var, value="single", bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT)
        audio_single_entry = tk.Entry(audio_single_frame, font=font_spec, width=35)
        audio_single_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(audio_single_frame, text="00:00", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=10)
        def select_single_audio():
            filename = filedialog.askopenfilename(title="选择音频文件", initialdir=AUDIO_FOLDER, filetypes=[("音频文件", "*.mp3 *.wav *.ogg *.flac *.m4a"), ("所有文件", "*.*")])
            if filename: audio_single_entry.delete(0, tk.END); audio_single_entry.insert(0, filename)
        tk.Button(audio_single_frame, text="选取...", command=select_single_audio, bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=3).pack(side=tk.LEFT, padx=5)
        tk.Label(content_frame, text="音频文件夹", font=font_spec, bg='#E8E8E8').grid(row=2, column=0, sticky='e', padx=5, pady=5)
        audio_folder_frame = tk.Frame(content_frame, bg='#E8E8E8')
        audio_folder_frame.grid(row=2, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        tk.Radiobutton(audio_folder_frame, text="", variable=audio_type_var, value="folder", bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT)
        audio_folder_entry = tk.Entry(audio_folder_frame, font=font_spec, width=50)
        audio_folder_entry.pack(side=tk.LEFT, padx=5)
        def select_folder():
            foldername = filedialog.askdirectory(title="选择音频文件夹", initialdir=AUDIO_FOLDER)
            if foldername: audio_folder_entry.delete(0, tk.END); audio_folder_entry.insert(0, foldername)
        tk.Button(audio_folder_frame, text="选取...", command=select_folder, bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=3).pack(side=tk.LEFT, padx=5)
        play_order_frame = tk.Frame(content_frame, bg='#E8E8E8')
        play_order_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=5)
        play_order_var = tk.StringVar(value="sequential")
        tk.Radiobutton(play_order_frame, text="顺序播", variable=play_order_var, value="sequential", bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(play_order_frame, text="随机播", variable=play_order_var, value="random", bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT, padx=10)
        volume_frame = tk.Frame(content_frame, bg='#E8E8E8')
        volume_frame.grid(row=4, column=1, columnspan=3, sticky='w', padx=5, pady=8)
        tk.Label(volume_frame, text="音量:", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT)
        volume_entry = tk.Entry(volume_frame, font=font_spec, width=10)
        volume_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(volume_frame, text="0-100", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=5)
        time_frame = tk.LabelFrame(main_frame, text="时间", font=('Microsoft YaHei', 12, 'bold'), bg='#E8E8E8', padx=15, pady=15)
        time_frame.grid(row=1, column=0, sticky='ew', pady=10)
        tk.Label(time_frame, text="开始时间:", font=font_spec, bg='#E8E8E8').grid(row=0, column=0, sticky='e', padx=5, pady=5)
        start_time_entry = tk.Entry(time_frame, font=font_spec, width=50)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        tk.Label(time_frame, text="《可多个,用英文逗号,隔开》", font=font_spec, bg='#E8E8E8').grid(row=0, column=2, sticky='w', padx=5)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=2).grid(row=0, column=3, padx=5)
        interval_var = tk.StringVar(value="first")
        interval_frame1 = tk.Frame(time_frame, bg='#E8E8E8')
        interval_frame1.grid(row=1, column=1, columnspan=2, sticky='w', padx=5, pady=5)
        tk.Label(time_frame, text="间隔播报:", font=font_spec, bg='#E8E8E8').grid(row=1, column=0, sticky='e', padx=5, pady=5)
        tk.Radiobutton(interval_frame1, text="播 n 首", variable=interval_var, value="first", bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT)
        interval_first_entry = tk.Entry(interval_frame1, font=font_spec, width=15)
        interval_first_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(interval_frame1, text="(单曲时,指 n 遍)", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=5)
        interval_frame2 = tk.Frame(time_frame, bg='#E8E8E8')
        interval_frame2.grid(row=2, column=1, columnspan=2, sticky='w', padx=5, pady=5)
        tk.Radiobutton(interval_frame2, text="播 n 秒", variable=interval_var, value="seconds", bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT)
        interval_seconds_entry = tk.Entry(interval_frame2, font=font_spec, width=15)
        interval_seconds_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(interval_frame2, text="(3600秒 = 1小时)", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=5)
        tk.Label(time_frame, text="周几/几号:", font=font_spec, bg='#E8E8E8').grid(row=3, column=0, sticky='e', padx=5, pady=8)
        weekday_entry = tk.Entry(time_frame, font=font_spec, width=50)
        weekday_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=8)
        tk.Button(time_frame, text="选取...", command=lambda: self.show_weekday_settings_dialog(weekday_entry), bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=3).grid(row=3, column=3, padx=5)
        tk.Label(time_frame, text="日期范围:", font=font_spec, bg='#E8E8E8').grid(row=4, column=0, sticky='e', padx=5, pady=8)
        date_range_entry = tk.Entry(time_frame, font=font_spec, width=50)
        date_range_entry.grid(row=4, column=1, sticky='ew', padx=5, pady=8)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry), bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=3).grid(row=4, column=3, padx=5)
        
        other_frame = tk.LabelFrame(main_frame, text="其它", font=('Microsoft YaHei', 12, 'bold'), bg='#E8E8E8', padx=10, pady=10)
        other_frame.grid(row=2, column=0, sticky='ew', pady=5)
        delay_var = tk.StringVar(value="ontime")
        tk.Label(other_frame, text="模式:", font=font_spec, bg='#E8E8E8').grid(row=0, column=0, sticky='ne', padx=5, pady=5)
        delay_frame = tk.Frame(other_frame, bg='#E8E8E8')
        delay_frame.grid(row=0, column=1, sticky='w', padx=5, pady=5)
        tk.Radiobutton(delay_frame, text="准时播 - 如果有别的节目正在播，终止他们（默认）", variable=delay_var, value="ontime", bg='#E8E8E8', font=font_spec).pack(anchor='w')
        tk.Radiobutton(delay_frame, text="可延后 - 如果有别的节目正在播，排队等候", variable=delay_var, value="delay", bg='#E8E8E8', font=font_spec).pack(anchor='w')
        tk.Radiobutton(delay_frame, text="立即播 - 添加后停止其他节目,立即播放此节目", variable=delay_var, value="immediate", bg='#E8E8E8', font=font_spec).pack(anchor='w')

        if is_edit_mode:
            task = task_to_edit
            name_entry.insert(0, task.get('name', ''))
            start_time_entry.insert(0, task.get('time', ''))
            audio_type_var.set(task.get('audio_type', 'single'))
            if task.get('audio_type') == 'single': audio_single_entry.insert(0, task.get('content', ''))
            else: audio_folder_entry.insert(0, task.get('content', ''))
            play_order_var.set(task.get('play_order', 'sequential'))
            volume_entry.insert(0, task.get('volume', '80'))
            interval_var.set(task.get('interval_type', 'first'))
            interval_first_entry.insert(0, task.get('interval_first', '1'))
            interval_seconds_entry.insert(0, task.get('interval_seconds', '600'))
            weekday_entry.insert(0, task.get('weekday', '每周:1234567'))
            date_range_entry.insert(0, task.get('date_range', '2000-01-01 ~ 2099-12-31'))
            delay_var.set(task.get('delay', 'ontime'))
        else:
            volume_entry.insert(0, "80"); interval_first_entry.insert(0, "1"); interval_seconds_entry.insert(0, "600")
            weekday_entry.insert(0, "每周:1234567"); date_range_entry.insert(0, "2000-01-01 ~ 2099-12-31")

        button_frame = tk.Frame(main_frame, bg='#E8E8E8')
        button_frame.grid(row=3, column=0, pady=20)
        def save_task():
            audio_path = audio_single_entry.get().strip() if audio_type_var.get() == "single" else audio_folder_entry.get().strip()
            if not audio_path: messagebox.showwarning("警告", "请选择音频文件或文件夹", parent=dialog); return
            is_valid_time, time_msg = self._normalize_multiple_times_string(start_time_entry.get().strip())
            if not is_valid_time: messagebox.showwarning("格式错误", time_msg, parent=dialog); return
            is_valid_date, date_msg = self._normalize_date_range_string(date_range_entry.get().strip())
            if not is_valid_date: messagebox.showwarning("格式错误", date_msg, parent=dialog); return
            
            play_mode = delay_var.get()
            play_this_task_now = (play_mode == 'immediate')
            saved_delay_type = 'ontime' if play_mode == 'immediate' else play_mode

            new_task_data = {'name': name_entry.get().strip(), 'time': time_msg, 'content': audio_path, 'type': 'audio', 'audio_type': audio_type_var.get(), 'play_order': play_order_var.get(), 'volume': volume_entry.get().strip() or "80", 'interval_type': interval_var.get(), 'interval_first': interval_first_entry.get().strip(), 'interval_seconds': interval_seconds_entry.get().strip(), 'weekday': weekday_entry.get().strip(), 'date_range': date_msg, 'delay': saved_delay_type, 'status': '启用' if not is_edit_mode else task_to_edit.get('status', '启用'), 'last_run': {} if not is_edit_mode else task_to_edit.get('last_run', {})}
            if not new_task_data['name'] or not new_task_data['time']: messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog); return
            
            if is_edit_mode: self.tasks[index] = new_task_data; self.log(f"已修改音频节目: {new_task_data['name']}")
            else: self.tasks.append(new_task_data); self.log(f"已添加音频节目: {new_task_data['name']}")
            
            self.update_task_list(); self.save_tasks(); dialog.destroy()

            if play_this_task_now:
                self.play_task_immediately(new_task_data)

        button_text = "保存修改" if is_edit_mode else "添加"
        tk.Button(button_frame, text=button_text, command=save_task, bg='#5DADE2', fg='white', font=('Microsoft YaHei', 11, 'bold'), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="取消", command=dialog.destroy, bg='#D0D0D0', font=('Microsoft YaHei', 11), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        content_frame.columnconfigure(1, weight=1); time_frame.columnconfigure(1, weight=1)

    def open_voice_dialog(self, parent_dialog, task_to_edit=None, index=None):
        parent_dialog.destroy()
        is_edit_mode = task_to_edit is not None
        dialog = tk.Toplevel(self.root)
        dialog.title("修改语音节目" if is_edit_mode else "添加语音节目")
        dialog.geometry("950x900")
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set(); dialog.configure(bg='#E8E8E8')
        main_frame = tk.Frame(dialog, bg='#E8E8E8', padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        content_frame = tk.LabelFrame(main_frame, text="内容", font=('Microsoft YaHei', 12, 'bold'), bg='#E8E8E8', padx=10, pady=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=5)
        font_spec = ('Microsoft YaHei', 11)
        
        tk.Label(content_frame, text="节目名称:", font=font_spec, bg='#E8E8E8').grid(row=0, column=0, sticky='w', padx=5, pady=5)
        name_entry = tk.Entry(content_frame, font=font_spec, width=65)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        tk.Label(content_frame, text="播音文字:", font=font_spec, bg='#E8E8E8').grid(row=1, column=0, sticky='nw', padx=5, pady=5)
        text_frame = tk.Frame(content_frame, bg='#E8E8E8')
        text_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        content_text = scrolledtext.ScrolledText(text_frame, height=5, font=font_spec, width=65, wrap=tk.WORD)
        content_text.pack(fill=tk.BOTH, expand=True)
        tk.Label(content_frame, text="播音员:", font=font_spec, bg='#E8E8E8').grid(row=2, column=0, sticky='w', padx=5, pady=8)
        voice_frame = tk.Frame(content_frame, bg='#E8E8E8')
        voice_frame.grid(row=2, column=1, columnspan=3, sticky='w', padx=5, pady=8)
        available_voices = self.get_available_voices()
        voice_var = tk.StringVar()
        voice_combo = ttk.Combobox(voice_frame, textvariable=voice_var, values=available_voices, font=font_spec, width=50, state='readonly')
        voice_combo.pack(side=tk.LEFT)
        speech_params_frame = tk.Frame(content_frame, bg='#E8E8E8')
        speech_params_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=5)
        tk.Label(speech_params_frame, text="语速(-10~10):", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=(0,5))
        speed_entry = tk.Entry(speech_params_frame, font=font_spec, width=8)
        speed_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(speech_params_frame, text="音调(-10~10):", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=(10,5))
        pitch_entry = tk.Entry(speech_params_frame, font=font_spec, width=8)
        pitch_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(speech_params_frame, text="音量(0-100):", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=(10,5))
        volume_entry = tk.Entry(speech_params_frame, font=font_spec, width=8)
        volume_entry.pack(side=tk.LEFT, padx=5)
        prompt_var = tk.IntVar()
        prompt_frame = tk.Frame(content_frame, bg='#E8E8E8')
        prompt_frame.grid(row=4, column=1, columnspan=3, sticky='w', padx=5, pady=5)
        tk.Checkbutton(prompt_frame, text="提示音:", variable=prompt_var, bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT)
        prompt_file_var, prompt_volume_var = tk.StringVar(), tk.StringVar()
        prompt_file_entry = tk.Entry(prompt_frame, textvariable=prompt_file_var, font=font_spec, width=20)
        prompt_file_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(prompt_frame, text="...", command=lambda: self.select_file_for_entry(PROMPT_FOLDER, prompt_file_var)).pack(side=tk.LEFT)
        tk.Label(prompt_frame, text="音量(0-100):", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=(10,5))
        tk.Entry(prompt_frame, textvariable=prompt_volume_var, font=font_spec, width=8).pack(side=tk.LEFT, padx=5)
        bgm_var = tk.IntVar()
        bgm_frame = tk.Frame(content_frame, bg='#E8E8E8')
        bgm_frame.grid(row=5, column=1, columnspan=3, sticky='w', padx=5, pady=5)
        tk.Checkbutton(bgm_frame, text="背景音乐:", variable=bgm_var, bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT)
        bgm_file_var, bgm_volume_var = tk.StringVar(), tk.StringVar()
        bgm_file_entry = tk.Entry(bgm_frame, textvariable=bgm_file_var, font=font_spec, width=20)
        bgm_file_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(bgm_frame, text="...", command=lambda: self.select_file_for_entry(BGM_FOLDER, bgm_file_var)).pack(side=tk.LEFT)
        tk.Label(bgm_frame, text="音量(0-100):", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=(10,5))
        tk.Entry(bgm_frame, textvariable=bgm_volume_var, font=font_spec, width=8).pack(side=tk.LEFT, padx=5)
        time_frame = tk.LabelFrame(main_frame, text="时间", font=('Microsoft YaHei', 12, 'bold'), bg='#E8E8E8', padx=10, pady=10)
        time_frame.grid(row=1, column=0, sticky='ew', pady=5)
        tk.Label(time_frame, text="开始时间:", font=font_spec, bg='#E8E8E8').grid(row=0, column=0, sticky='e', padx=5, pady=5)
        start_time_entry = tk.Entry(time_frame, font=font_spec, width=50)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        tk.Label(time_frame, text="《可多个,用英文逗号,隔开》", font=font_spec, bg='#E8E8E8').grid(row=0, column=2, sticky='w', padx=5)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=2).grid(row=0, column=3, padx=5)
        tk.Label(time_frame, text="播 n 遍:", font=font_spec, bg='#E8E8E8').grid(row=1, column=0, sticky='e', padx=5, pady=5)
        repeat_entry = tk.Entry(time_frame, font=font_spec, width=12)
        repeat_entry.grid(row=1, column=1, sticky='w', padx=5, pady=5)
        tk.Label(time_frame, text="周几/几号:", font=font_spec, bg='#E8E8E8').grid(row=2, column=0, sticky='e', padx=5, pady=5)
        weekday_entry = tk.Entry(time_frame, font=font_spec, width=50)
        weekday_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=5)
        tk.Button(time_frame, text="选取...", command=lambda: self.show_weekday_settings_dialog(weekday_entry), bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=2).grid(row=2, column=3, padx=5)
        tk.Label(time_frame, text="日期范围:", font=font_spec, bg='#E8E8E8').grid(row=3, column=0, sticky='e', padx=5, pady=5)
        date_range_entry = tk.Entry(time_frame, font=font_spec, width=50)
        date_range_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=5)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry), bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=2).grid(row=3, column=3, padx=5)
        
        other_frame = tk.LabelFrame(main_frame, text="其它", font=('Microsoft YaHei', 12, 'bold'), bg='#E8E8E8', padx=15, pady=15)
        other_frame.grid(row=2, column=0, sticky='ew', pady=10)
        delay_var = tk.StringVar(value="delay")
        tk.Label(other_frame, text="模式:", font=font_spec, bg='#E8E8E8').grid(row=0, column=0, sticky='ne', padx=5, pady=3)
        delay_frame = tk.Frame(other_frame, bg='#E8E8E8')
        delay_frame.grid(row=0, column=1, sticky='w', padx=5, pady=3)
        tk.Radiobutton(delay_frame, text="准时播 - 如果有别的节目正在播，终止他们", variable=delay_var, value="ontime", bg='#E8E8E8', font=font_spec).pack(anchor='w', pady=2)
        tk.Radiobutton(delay_frame, text="可延后 - 如果有别的节目正在播，排队等候（默认）", variable=delay_var, value="delay", bg='#E8E8E8', font=font_spec).pack(anchor='w', pady=2)
        tk.Radiobutton(delay_frame, text="立即播 - 添加后停止其他节目,立即播放此节目", variable=delay_var, value="immediate", bg='#E8E8E8', font=font_spec).pack(anchor='w', pady=2)

        if is_edit_mode:
            task = task_to_edit
            name_entry.insert(0, task.get('name', ''))
            content_text.insert('1.0', task.get('source_text', ''))
            voice_var.set(task.get('voice', ''))
            speed_entry.insert(0, task.get('speed', '0'))
            pitch_entry.insert(0, task.get('pitch', '0'))
            volume_entry.insert(0, task.get('volume', '80'))
            prompt_var.set(task.get('prompt', 0)); prompt_file_var.set(task.get('prompt_file', '')); prompt_volume_var.set(task.get('prompt_volume', '80'))
            bgm_var.set(task.get('bgm', 0)); bgm_file_var.set(task.get('bgm_file', '')); bgm_volume_var.set(task.get('bgm_volume', '40'))
            start_time_entry.insert(0, task.get('time', ''))
            repeat_entry.insert(0, task.get('repeat', '1'))
            weekday_entry.insert(0, task.get('weekday', '每周:1234567'))
            date_range_entry.insert(0, task.get('date_range', '2000-01-01 ~ 2099-12-31'))
            delay_var.set(task.get('delay', 'delay'))
        else:
            speed_entry.insert(0, "0"); pitch_entry.insert(0, "0"); volume_entry.insert(0, "80")
            prompt_var.set(0); prompt_volume_var.set("80"); bgm_var.set(0); bgm_volume_var.set("40")
            repeat_entry.insert(0, "1"); weekday_entry.insert(0, "每周:1234567"); date_range_entry.insert(0, "2000-01-01 ~ 2099-12-31")

        button_frame = tk.Frame(main_frame, bg='#E8E8E8')
        button_frame.grid(row=3, column=0, pady=20)
        
        def save_task():
            text_content = content_text.get('1.0', tk.END).strip()
            if not text_content: messagebox.showwarning("警告", "请输入播音文字内容", parent=dialog); return
            is_valid_time, time_msg = self._normalize_multiple_times_string(start_time_entry.get().strip())
            if not is_valid_time: messagebox.showwarning("格式错误", time_msg, parent=dialog); return
            is_valid_date, date_msg = self._normalize_date_range_string(date_range_entry.get().strip())
            if not is_valid_date: messagebox.showwarning("格式错误", date_msg, parent=dialog); return
            
            play_mode = delay_var.get()
            play_this_task_now = (play_mode == 'immediate')
            saved_delay_type = 'delay' if play_mode == 'immediate' else play_mode

            progress_dialog = tk.Toplevel(dialog)
            progress_dialog.title("请稍候")
            progress_dialog.geometry("300x100")
            progress_dialog.resizable(False, False)
            progress_dialog.transient(dialog); progress_dialog.grab_set()
            tk.Label(progress_dialog, text="语音文件生成中，请稍后...", font=font_spec).pack(expand=True)
            self.center_window(progress_dialog, 300, 100)
            dialog.update_idletasks()
            
            wav_filename = f"{int(time.time())}_{random.randint(1000, 9999)}.wav"
            output_path = os.path.join(AUDIO_FOLDER, wav_filename)
            voice_params = {'voice': voice_var.get(), 'speed': speed_entry.get().strip() or "0", 'pitch': pitch_entry.get().strip() or "0", 'volume': volume_entry.get().strip() or "80"}

            def _on_synthesis_complete(result):
                progress_dialog.destroy()
                if not result['success']:
                    messagebox.showerror("错误", f"无法生成语音文件: {result['error']}", parent=dialog)
                    return

                if is_edit_mode and 'wav_filename' in task_to_edit:
                    old_wav_path = os.path.join(AUDIO_FOLDER, task_to_edit['wav_filename'])
                    if os.path.exists(old_wav_path):
                        try: os.remove(old_wav_path); self.log(f"已删除旧语音文件: {task_to_edit['wav_filename']}")
                        except Exception as e: self.log(f"删除旧语音文件失败: {e}")

                new_task_data = {'name': name_entry.get().strip(), 'time': time_msg, 'type': 'voice', 'content': output_path, 'wav_filename': wav_filename, 'source_text': text_content, 'voice': voice_params['voice'], 'speed': voice_params['speed'], 'pitch': voice_params['pitch'], 'volume': voice_params['volume'], 'prompt': prompt_var.get(), 'prompt_file': prompt_file_var.get(), 'prompt_volume': prompt_volume_var.get(), 'bgm': bgm_var.get(), 'bgm_file': bgm_file_var.get(), 'bgm_volume': bgm_volume_var.get(), 'repeat': repeat_entry.get().strip() or "1", 'weekday': weekday_entry.get().strip(), 'date_range': date_msg, 'delay': saved_delay_type, 'status': '启用' if not is_edit_mode else task_to_edit.get('status', '启用'), 'last_run': {} if not is_edit_mode else task_to_edit.get('last_run', {})}
                if not new_task_data['name'] or not new_task_data['time']: messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog); return
                
                if is_edit_mode: self.tasks[index] = new_task_data; self.log(f"已修改语音节目: {new_task_data['name']}")
                else: self.tasks.append(new_task_data); self.log(f"已添加语音节目: {new_task_data['name']}")
                
                self.update_task_list(); self.save_tasks(); dialog.destroy()

                if play_this_task_now:
                    self.play_task_immediately(new_task_data)

            threading.Thread(target=self._synthesis_worker, args=(text_content, voice_params, output_path, _on_synthesis_complete), daemon=True).start()

        button_text = "保存修改" if is_edit_mode else "添加"
        tk.Button(button_frame, text=button_text, command=save_task, bg='#5DADE2', fg='white', font=('Microsoft YaHei', 11, 'bold'), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="取消", command=dialog.destroy, bg='#D0D0D0', font=('Microsoft YaHei', 11), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        content_frame.columnconfigure(1, weight=1); time_frame.columnconfigure(1, weight=1)

    def _synthesis_worker(self, text, voice_params, output_path, callback):
        try:
            success = self._synthesize_text_to_wav(text, voice_params, output_path)
            if success:
                self.root.after(0, callback, {'success': True})
            else:
                raise Exception("合成过程返回失败")
        except Exception as e:
            self.root.after(0, callback, {'success': False, 'error': str(e)})

    def _synthesize_text_to_wav(self, text, voice_params, output_path):
        if not WIN32COM_AVAILABLE:
            raise ImportError("pywin32 模块未安装，无法进行语音合成。")
        
        pythoncom.CoInitialize()
        try:
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            stream = win32com.client.Dispatch("SAPI.SpFileStream")
            stream.Open(output_path, 3, False)
            speaker.AudioOutputStream = stream
            
            all_voices = {v.GetDescription(): v for v in speaker.GetVoices()}
            if (selected_voice_desc := voice_params.get('voice')) in all_voices:
                speaker.Voice = all_voices[selected_voice_desc]

            speaker.Volume = int(voice_params.get('volume', 80))
            escaped_text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("'", "&apos;").replace('"', "&quot;")
            xml_text = f"<rate absspeed='{voice_params.get('speed', '0')}'><pitch middle='{voice_params.get('pitch', '0')}'>{escaped_text}</pitch></rate>"
            
            speaker.Speak(xml_text, 1)
            speaker.WaitUntilDone(-1)
            stream.Close()
            return True
        except Exception as e:
            self.log(f"语音合成到文件时出错: {e}")
            return False
        finally:
            pythoncom.CoUninitialize()

    def get_available_voices(self):
        if not WIN32COM_AVAILABLE: return []
        try:
            pythoncom.CoInitialize()
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            voices = [v.GetDescription() for v in speaker.GetVoices()]
            pythoncom.CoUninitialize()
            return voices
        except Exception as e:
            self.log(f"警告: 使用 win32com 获取语音列表失败 - {e}")
            return []
    
    def select_file_for_entry(self, initial_dir, string_var):
        filename = filedialog.askopenfilename(title="选择文件", initialdir=initial_dir, filetypes=[("音频文件", "*.mp3 *.wav *.ogg *.flac *.m4a"), ("所有文件", "*.*")])
        if filename: string_var.set(os.path.basename(filename))

    def delete_task(self):
        selections = self.task_tree.selection()
        if not selections: messagebox.showwarning("警告", "请先选择要删除的节目"); return
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(selections)} 个节目吗？"):
            indices = sorted([self.task_tree.index(s) for s in selections], reverse=True)
            for index in indices:
                task_to_delete = self.tasks[index]
                if task_to_delete.get('type') == 'voice' and 'wav_filename' in task_to_delete:
                    wav_path = os.path.join(AUDIO_FOLDER, task_to_delete['wav_filename'])
                    if os.path.exists(wav_path):
                        try: os.remove(wav_path); self.log(f"已删除语音文件: {task_to_delete['wav_filename']}")
                        except Exception as e: self.log(f"删除语音文件失败: {e}")
                self.log(f"已删除节目: {self.tasks.pop(index)['name']}")
            self.update_task_list(); self.save_tasks()

    def edit_task(self):
        selection = self.task_tree.selection()
        if not selection: messagebox.showwarning("警告", "请先选择要修改的节目"); return
        if len(selection) > 1: messagebox.showwarning("警告", "一次只能修改一个节目"); return
        index = self.task_tree.index(selection[0])
        task = self.tasks[index]
        dummy_parent = tk.Toplevel(self.root); dummy_parent.withdraw()
        if task.get('type') == 'audio': self.open_audio_dialog(dummy_parent, task_to_edit=task, index=index)
        else: self.open_voice_dialog(dummy_parent, task_to_edit=task, index=index)
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

            if copy.get('type') == 'voice' and 'source_text' in copy:
                wav_filename = f"{int(time.time())}_{random.randint(1000, 9999)}.wav"
                output_path = os.path.join(AUDIO_FOLDER, wav_filename)
                voice_params = {'voice': copy.get('voice'), 'speed': copy.get('speed'), 'pitch': copy.get('pitch'), 'volume': copy.get('volume')}
                try:
                    success = self._synthesize_text_to_wav(copy['source_text'], voice_params, output_path)
                    if not success: raise Exception("语音合成失败")
                    copy['content'] = output_path
                    copy['wav_filename'] = wav_filename
                    self.log(f"已为副本生成新语音文件: {wav_filename}")
                except Exception as e:
                    self.log(f"为副本生成语音文件失败: {e}")
                    continue

            self.tasks.append(copy)
            self.log(f"已复制节目: {original['name']}")
        self.update_task_list(); self.save_tasks()

    def move_task(self, direction):
        selections = self.task_tree.selection()
        if not selections or len(selections) > 1: return
        index = self.task_tree.index(selections[0])
        new_index = index + direction
        if 0 <= new_index < len(self.tasks):
            task_to_move = self.tasks.pop(index)
            self.tasks.insert(new_index, task_to_move)
            self.update_task_list(); self.save_tasks()
            items = self.task_tree.get_children()
            if items: self.task_tree.selection_set(items[new_index]); self.task_tree.focus(items[new_index])

    def move_task_to_top(self):
        selections = self.task_tree.selection()
        if not selections or len(selections) > 1: return
        index = self.task_tree.index(selections[0])
        if index > 0:
            task_to_move = self.tasks.pop(index)
            self.tasks.insert(0, task_to_move)
            self.update_task_list(); self.save_tasks()
            items = self.task_tree.get_children()
            if items: self.task_tree.selection_set(items[0]); self.task_tree.focus(items[0])

    def move_task_to_bottom(self):
        selections = self.task_tree.selection()
        if not selections or len(selections) > 1: return
        index = self.task_tree.index(selections[0])
        if index < len(self.tasks) - 1:
            task_to_move = self.tasks.pop(index)
            self.tasks.append(task_to_move)
            self.update_task_list(); self.save_tasks()
            items = self.task_tree.get_children()
            if items: self.task_tree.selection_set(items[-1]); self.task_tree.focus(items[-1])

    def import_tasks(self):
        filename = filedialog.askopenfilename(title="选择导入文件", filetypes=[("JSON文件", "*.json")], initialdir=application_path)
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f: imported = json.load(f)
                self.tasks.extend(imported); self.update_task_list(); self.save_tasks()
                self.log(f"已从 {os.path.basename(filename)} 导入 {len(imported)} 个节目")
            except Exception as e: messagebox.showerror("错误", f"导入失败: {e}")

    def export_tasks(self):
        if not self.tasks: messagebox.showwarning("警告", "没有节目可以导出"); return
        filename = filedialog.asksaveasfilename(title="导出到...", defaultextension=".json", initialfile="broadcast_backup.json", filetypes=[("JSON文件", "*.json")], initialdir=application_path)
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
        
    def enable_all_tasks(self):
        if not self.tasks: return
        for task in self.tasks: task['status'] = '启用'
        self.update_task_list(); self.save_tasks(); self.log("已启用全部节目。")
    
    def disable_all_tasks(self):
        if not self.tasks: return
        for task in self.tasks: task['status'] = '禁用'
        self.update_task_list(); self.save_tasks(); self.log("已禁用全部节目。")

    def set_uniform_volume(self):
        if not self.tasks: return
        volume = self._create_custom_input_dialog(
            title="统一音量",
            prompt="请输入统一音量值 (0-100):",
            minvalue=0,
            maxvalue=100
        )
        if volume is not None:
            for task in self.tasks: task['volume'] = str(volume)
            self.update_task_list(); self.save_tasks()
            self.log(f"已将全部节目音量统一设置为 {volume}。")
    
    def _create_custom_input_dialog(self, title, prompt, minvalue=None, maxvalue=None):
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("350x150")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        self.center_window(dialog, 350, 150)

        result = [None]
        
        tk.Label(dialog, text=prompt, font=('Microsoft YaHei', 11)).pack(pady=10)
        entry = tk.Entry(dialog, font=('Microsoft YaHei', 11), width=15, justify='center')
        entry.pack(pady=5)
        entry.focus_set()

        def on_confirm():
            try:
                value = int(entry.get())
                if (minvalue is not None and value < minvalue) or \
                   (maxvalue is not None and value > maxvalue):
                    messagebox.showerror("输入错误", f"请输入一个介于 {minvalue} 和 {maxvalue} 之间的整数。", parent=dialog)
                    return
                result[0] = value
                dialog.destroy()
            except ValueError:
                messagebox.showerror("输入错误", "请输入一个有效的整数。", parent=dialog)

        def on_cancel():
            dialog.destroy()

        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=15)
        
        tk.Button(btn_frame, text="确定", command=on_confirm, width=8).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="取消", command=on_cancel, width=8).pack(side=tk.LEFT, padx=10)
        
        dialog.bind('<Return>', lambda event: on_confirm())
        
        self.root.wait_window(dialog)
        return result[0]

    def clear_all_tasks(self):
        if not self.tasks: return
        if messagebox.askyesno("严重警告", "您确定要清空所有节目吗？\n此操作不可恢复！"):
            for task in self.tasks:
                if task.get('type') == 'voice' and 'wav_filename' in task:
                    wav_path = os.path.join(AUDIO_FOLDER, task['wav_filename'])
                    if os.path.exists(wav_path):
                        try: os.remove(wav_path)
                        except Exception as e: self.log(f"删除语音文件失败: {e}")
            
            self.tasks.clear()
            self.update_task_list(); self.save_tasks()
            self.log("已清空所有节目。")

    def show_time_settings_dialog(self, time_entry):
        dialog = tk.Toplevel(self.root)
        dialog.title("开始时间设置"); dialog.geometry("480x450"); dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set(); dialog.configure(bg='#D7F3F5')
        self.center_window(dialog, 480, 450)
        main_frame = tk.Frame(dialog, bg='#D7F3F5', padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        font_spec = ('Microsoft YaHei', 11)
        tk.Label(main_frame, text="24小时制 HH:MM:SS", font=(font_spec[0], 11, 'bold'), bg='#D7F3F5').pack(anchor='w', pady=5)
        list_frame = tk.LabelFrame(main_frame, text="时间列表", bg='#D7F3F5', padx=5, pady=5, font=font_spec)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        box_frame = tk.Frame(list_frame); box_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        listbox = tk.Listbox(box_frame, font=font_spec, height=10)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = tk.Scrollbar(box_frame, orient=tk.VERTICAL, command=listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y); listbox.configure(yscrollcommand=scrollbar.set)
        for t in [t.strip() for t in time_entry.get().split(',') if t.strip()]: listbox.insert(tk.END, t)
        btn_frame = tk.Frame(list_frame, bg='#D7F3F5')
        btn_frame.pack(side=tk.RIGHT, padx=10, fill=tk.Y)
        new_entry = tk.Entry(btn_frame, font=font_spec, width=12)
        new_entry.insert(0, datetime.now().strftime("%H:%M:%S")); new_entry.pack(pady=3)
        def add_time():
            val = new_entry.get().strip()
            normalized_time = self._normalize_time_string(val)
            if normalized_time:
                if normalized_time not in listbox.get(0, tk.END): listbox.insert(tk.END, normalized_time); new_entry.delete(0, tk.END); new_entry.insert(0, datetime.now().strftime("%H:%M:%S"))
            else: messagebox.showerror("格式错误", "请输入有效的时间格式 HH:MM:SS", parent=dialog)
        def del_time():
            if listbox.curselection(): listbox.delete(listbox.curselection()[0])
        tk.Button(btn_frame, text="添加 ↑", command=add_time, font=font_spec).pack(pady=3, fill=tk.X)
        tk.Button(btn_frame, text="删除", command=del_time, font=font_spec).pack(pady=3, fill=tk.X)
        tk.Button(btn_frame, text="清空", command=lambda: listbox.delete(0, tk.END), font=font_spec).pack(pady=3, fill=tk.X)
        bottom_frame = tk.Frame(main_frame, bg='#D7F3F5'); bottom_frame.pack(pady=10)
        def confirm():
            result = ", ".join(list(listbox.get(0, tk.END)))
            if isinstance(time_entry, tk.Entry): time_entry.delete(0, tk.END); time_entry.insert(0, result)
            self.save_settings(); dialog.destroy()
        tk.Button(bottom_frame, text="确定", command=confirm, bg='#5DADE2', fg='white', font=(font_spec[0], 11, 'bold'), bd=1, padx=25, pady=5).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0', font=font_spec, bd=1, padx=25, pady=5).pack(side=tk.LEFT, padx=5)

    def show_weekday_settings_dialog(self, weekday_var):
        dialog = tk.Toplevel(self.root); dialog.title("周几或几号")
        dialog.geometry("550x550"); dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set(); dialog.configure(bg='#D7F3F5')
        self.center_window(dialog, 550, 550)
        main_frame = tk.Frame(dialog, bg='#D7F3F5', padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        week_type_var = tk.StringVar(value="week")
        font_spec = ('Microsoft YaHei', 11)
        week_frame = tk.LabelFrame(main_frame, text="按周", font=(font_spec[0], 11, 'bold'), bg='#D7F3F5', padx=10, pady=10)
        week_frame.pack(fill=tk.X, pady=5)
        tk.Radiobutton(week_frame, text="每周", variable=week_type_var, value="week", bg='#D7F3F5', font=font_spec).grid(row=0, column=0, sticky='w')
        weekdays = [("周一", 1), ("周二", 2), ("周三", 3), ("周四", 4), ("周五", 5), ("周六", 6), ("周日", 7)]
        week_vars = {num: tk.IntVar(value=1) for day, num in weekdays}
        for i, (day, num) in enumerate(weekdays): tk.Checkbutton(week_frame, text=day, variable=week_vars[num], bg='#D7F3F5', font=font_spec).grid(row=(i // 4) + 1, column=i % 4, sticky='w', padx=10, pady=3)
        day_frame = tk.LabelFrame(main_frame, text="按月", font=(font_spec[0], 11, 'bold'), bg='#D7F3F5', padx=10, pady=10)
        day_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        tk.Radiobutton(day_frame, text="每月", variable=week_type_var, value="day", bg='#D7F3F5', font=font_spec).grid(row=0, column=0, sticky='w')
        day_vars = {i: tk.IntVar(value=0) for i in range(1, 32)}
        for i in range(1, 32): tk.Checkbutton(day_frame, text=f"{i:02d}", variable=day_vars[i], bg='#D7F3F5', font=font_spec).grid(row=((i - 1) // 7) + 1, column=(i - 1) % 7, sticky='w', padx=8, pady=2)
        bottom_frame = tk.Frame(main_frame, bg='#D7F3F5'); bottom_frame.pack(pady=10)
        current_val = weekday_var.get()
        if current_val.startswith("每周:"):
            week_type_var.set("week")
            selected_days = current_val.replace("每周:", "")
            for day_num in week_vars: week_vars[day_num].set(1 if str(day_num) in selected_days else 0)
        elif current_val.startswith("每月:"):
            week_type_var.set("day")
            selected_days = current_val.replace("每月:", "").split(',')
            for day_num in day_vars: day_vars[day_num].set(1 if f"{day_num:02d}" in selected_days else 0)
        def confirm():
            if week_type_var.get() == "week": result = "每周:" + "".join(sorted([str(n) for n, v in week_vars.items() if v.get()]))
            else: result = "每月:" + ",".join(sorted([f"{n:02d}" for n, v in day_vars.items() if v.get()]))
            if isinstance(weekday_var, tk.Entry): weekday_var.delete(0, tk.END); weekday_var.insert(0, result)
            self.save_settings(); dialog.destroy()
        tk.Button(bottom_frame, text="确定", command=confirm, bg='#5DADE2', fg='white', font=(font_spec[0], 11, 'bold'), bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0', font=font_spec, bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)

    def show_daterange_settings_dialog(self, date_range_entry):
        dialog = tk.Toplevel(self.root)
        dialog.title("日期范围"); dialog.geometry("450x250"); dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set(); dialog.configure(bg='#D7F3F5')
        self.center_window(dialog, 450, 250)
        main_frame = tk.Frame(dialog, bg='#D7F3F5', padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        font_spec = ('Microsoft YaHei', 11)
        from_frame = tk.Frame(main_frame, bg='#D7F3F5')
        from_frame.pack(pady=10, anchor='w')
        tk.Label(from_frame, text="从", font=(font_spec[0], 11, 'bold'), bg='#D7F3F5').pack(side=tk.LEFT, padx=5)
        from_date_entry = tk.Entry(from_frame, font=font_spec, width=18)
        from_date_entry.pack(side=tk.LEFT, padx=5)
        to_frame = tk.Frame(main_frame, bg='#D7F3F5')
        to_frame.pack(pady=10, anchor='w')
        tk.Label(to_frame, text="到", font=(font_spec[0], 11, 'bold'), bg='#D7F3F5').pack(side=tk.LEFT, padx=5)
        to_date_entry = tk.Entry(to_frame, font=font_spec, width=18)
        to_date_entry.pack(side=tk.LEFT, padx=5)
        try: start, end = date_range_entry.get().split('~'); from_date_entry.insert(0, start.strip()); to_date_entry.insert(0, end.strip())
        except (ValueError, IndexError): from_date_entry.insert(0, "2000-01-01"); to_date_entry.insert(0, "2099-12-31")
        tk.Label(main_frame, text="格式: YYYY-MM-DD", font=font_spec, bg='#D7F3F5', fg='#666').pack(pady=10)
        bottom_frame = tk.Frame(main_frame, bg='#D7F3F5'); bottom_frame.pack(pady=10)
        def confirm():
            start, end = from_date_entry.get().strip(), to_date_entry.get().strip()
            norm_start, norm_end = self._normalize_date_string(start), self._normalize_date_string(end)
            if norm_start and norm_end: date_range_entry.delete(0, tk.END); date_range_entry.insert(0, f"{norm_start} ~ {norm_end}"); dialog.destroy()
            else: messagebox.showerror("格式错误", "日期格式不正确, 应为 YYYY-MM-DD", parent=dialog)
        tk.Button(bottom_frame, text="确定", command=confirm, bg='#5DADE2', fg='white', font=(font_spec[0], 11, 'bold'), bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0', font=font_spec, bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)

    def show_single_time_dialog(self, time_var):
        dialog = tk.Toplevel(self.root)
        dialog.title("设置时间"); dialog.geometry("320x200"); dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set(); dialog.configure(bg='#D7F3F5')
        self.center_window(dialog, 320, 200)
        main_frame = tk.Frame(dialog, bg='#D7F3F5', padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        font_spec = ('Microsoft YaHei', 11)
        tk.Label(main_frame, text="24小时制 HH:MM:SS", font=(font_spec[0], 11, 'bold'), bg='#D7F3F5').pack(pady=5)
        time_entry = tk.Entry(main_frame, font=('Microsoft YaHei', 12), width=15, justify='center')
        time_entry.insert(0, time_var.get()); time_entry.pack(pady=10)
        def confirm():
            val = time_entry.get().strip()
            normalized_time = self._normalize_time_string(val)
            if normalized_time: time_var.set(normalized_time); self.save_settings(); dialog.destroy()
            else: messagebox.showerror("格式错误", "请输入有效的时间格式 HH:MM:SS", parent=dialog)
        bottom_frame = tk.Frame(main_frame, bg='#D7F3F5'); bottom_frame.pack(pady=10)
        tk.Button(bottom_frame, text="确定", command=confirm, bg='#5DADE2', fg='white', font=font_spec).pack(side=tk.LEFT, padx=10)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0', font=font_spec).pack(side=tk.LEFT, padx=10)

    def show_power_week_time_dialog(self, title, days_var, time_var):
        dialog = tk.Toplevel(self.root); dialog.title(title)
        dialog.geometry("580x330"); dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set(); dialog.configure(bg='#D7F3F5')
        self.center_window(dialog, 580, 330)
        font_spec = ('Microsoft YaHei', 11)
        week_frame = tk.LabelFrame(dialog, text="选择周几", font=(font_spec[0], 11, 'bold'), bg='#D7F3F5', padx=10, pady=10)
        week_frame.pack(fill=tk.X, pady=10, padx=10)
        weekdays = [("周一", 1), ("周二", 2), ("周三", 3), ("周四", 4), ("周五", 5), ("周六", 6), ("周日", 7)]
        week_vars = {num: tk.IntVar() for day, num in weekdays}
        current_days = days_var.get().replace("每周:", "")
        for day_num_str in current_days: week_vars[int(day_num_str)].set(1)
        for i, (day, num) in enumerate(weekdays): tk.Checkbutton(week_frame, text=day, variable=week_vars[num], bg='#D7F3F5', font=font_spec).grid(row=0, column=i, sticky='w', padx=10, pady=3)
        time_frame = tk.LabelFrame(dialog, text="设置时间", font=(font_spec[0], 11, 'bold'), bg='#D7F3F5', padx=10, pady=10)
        time_frame.pack(fill=tk.X, pady=10, padx=10)
        tk.Label(time_frame, text="时间 (HH:MM:SS):", font=font_spec, bg='#D7F3F5').pack(side=tk.LEFT)
        time_entry = tk.Entry(time_frame, font=font_spec, width=15)
        time_entry.insert(0, time_var.get()); time_entry.pack(side=tk.LEFT, padx=10)
        def confirm():
            selected_days = sorted([str(n) for n, v in week_vars.items() if v.get()])
            if not selected_days: messagebox.showwarning("提示", "请至少选择一天", parent=dialog); return
            normalized_time = self._normalize_time_string(time_entry.get().strip())
            if not normalized_time: messagebox.showerror("格式错误", "请输入有效的时间格式 HH:MM:SS", parent=dialog); return
            days_var.set("每周:" + "".join(selected_days)); time_var.set(normalized_time); self.save_settings(); dialog.destroy()
        bottom_frame = tk.Frame(dialog, bg='#D7F3F5'); bottom_frame.pack(pady=15)
        tk.Button(bottom_frame, text="确定", command=confirm, bg='#5DADE2', fg='white', font=font_spec).pack(side=tk.LEFT, padx=10)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0', font=font_spec).pack(side=tk.LEFT, padx=10)

def main():
    root = tk.Tk()
    app = TimedBroadcastApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
