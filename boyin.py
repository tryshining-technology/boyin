import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText, ScrolledFrame
from tkinter import messagebox, filedialog, simpledialog, font
import tkinter as tk
import subprocess
import shlex

import json
import threading
import time
from datetime import datetime, timedelta
import os
import random
import sys
import getpass
import base64
import queue
import shutil
import re
import ctypes
import hashlib
import requests
import edge_tts
import asyncio

# --- ↓↓↓ 新增代码：全局隐藏 subprocess 调用的控制台窗口 ↓↓↓ ---

# 仅在 Windows 平台上执行此操作
if sys.platform == "win32":
    # 创建一个 STARTUPINFO 结构体实例
    startupinfo = subprocess.STARTUPINFO()
    # 设置 dwFlags 来指定 wShowWindow 成员有效
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    # 设置 wShowWindow 为 SW_HIDE (0)，这将隐藏窗口
    startupinfo.wShowWindow = 0 
else:
    startupinfo = None

# 重写 subprocess.Popen 的默认行为
# 我们用一个 lambda 函数来包装原始的 Popen，并传入新的 startupinfo
_original_popen = subprocess.Popen
subprocess.Popen = lambda *args, **kwargs: _original_popen(
    *args,
    **kwargs,
    startupinfo=startupinfo
)

# --- 全局修复：启用高DPI感知 ---
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware(True)
    except Exception:
        print("警告: 无法设置DPI感知，在高分屏下布局可能出现问题。")
# --- DPI修复结束 ---

# 尝试导入所需库
TRAY_AVAILABLE = False
try:
    from pystray import MenuItem as item, Icon
    from PIL import Image, ImageTk, ImageGrab
    TRAY_AVAILABLE = True
    IMAGE_AVAILABLE = True
except ImportError:
    print("警告: pystray 或 Pillow 未安装，最小化到托盘和背景图片功能不可用。")
    TRAY_AVAILABLE = False
    IMAGE_AVAILABLE = False

WIN32_AVAILABLE = False
try:
    import win32com.client
    import pythoncom
    from pywintypes import com_error
    import winreg
    import win32gui
    import win32con
    import win32print
    import win32api
    WIN32_AVAILABLE = True
except ImportError:
    print("警告: pywin32 未安装，语音、开机启动、任务栏闪烁和密码持久化/注册功能将受限。")

AUDIO_AVAILABLE = False
try:
    import pygame
    pygame.mixer.init()
    pygame.mixer.set_num_channels(10)
    AUDIO_AVAILABLE = True
except ImportError:
    print("警告: pygame 未安装，音频播放功能将不可用。")
except Exception as e:
    print(f"警告: pygame 初始化失败 - {e}，音频播放功能将不可用。")

PSUTIL_AVAILABLE = False
try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    print("警告: psutil 未安装，无法获取机器码、强制结束进程，注册功能将受限。")

VLC_AVAILABLE = False
try:
    import vlc
    VLC_AVAILABLE = True
except (ImportError, OSError):
    print("警告: 未能在系统中找到VLC核心库。")
    print("提示: 请在电脑上安装官方VLC播放器以启用视频播放功能。")
except Exception as e:
    print(f"警告: vlc 初始化时发生未知错误 - {e}，视频播放功能不可用。")

def resource_path(relative_path):
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

WALLPAPER_CACHE_FOLDER = os.path.join(application_path, "每日壁纸")
TASK_FILE = os.path.join(application_path, "broadcast_tasks.json")
SETTINGS_FILE = os.path.join(application_path, "settings.json")
HOLIDAY_FILE = os.path.join(application_path, "holidays.json")
TODO_FILE = os.path.join(application_path, "todos.json")
SCREENSHOT_TASK_FILE = os.path.join(application_path, "screenshot_tasks.json")
EXECUTE_TASK_FILE = os.path.join(application_path, "execute_tasks.json")
PRINT_TASK_FILE = os.path.join(application_path, "print_tasks.json")
BACKUP_TASK_FILE = os.path.join(application_path, "backup_tasks.json")
DYNAMIC_VOICE_TASK_FILE = os.path.join(application_path, "dynamic_voice_tasks.json")
TIMESTAMP_FILE = os.path.join(application_path, ".timestamp.dat")

PROMPT_FOLDER = os.path.join(application_path, "提示音")
AUDIO_FOLDER = os.path.join(application_path, "音频文件")
BGM_FOLDER = os.path.join(application_path, "文稿背景")
VOICE_SCRIPT_FOLDER = os.path.join(application_path, "语音文稿")
SCREENSHOT_FOLDER = os.path.join(application_path, "截屏")
DYNAMIC_VOICE_CACHE_FOLDER = os.path.join(AUDIO_FOLDER, "动态语音缓存")

ICON_FILE = resource_path("icon.ico")
REMINDER_SOUND_FILE = os.path.join(PROMPT_FOLDER, "reminder.wav")
CHIME_FOLDER = os.path.join(AUDIO_FOLDER, "整点报时")

REGISTRY_KEY_PATH = r"Software\创翔科技\TimedBroadcastApp"
REGISTRY_PARENT_KEY_PATH = r"Software\创翔科技"
# --- ↓↓↓ 新增代码：定义一个用于签名的密钥盐 ↓↓↓ ---
# !!! 警告：请将这个字符串修改为您自己的、独一无二的复杂字符串 !!!
SECRET_SALT = "42492f00-d980-40e1-a17e-ba8094727636"
AMAP_API_KEY = "c62d9b56d92792d1d11c8544f1b547dc"
PRE_GENERATION_MINUTES = 5 # 动态语音预生成提前分钟数
SENTINEL_LOCATIONS = [
    # 文件哨兵1: 程序根目录 (相对路径)
    ('file', 'dat.sys', None, None), 
    # 文件哨兵2: 公共文档目录 (高权限，通常可写)
    ('file', os.path.join(os.environ.get("PUBLIC", r"C:\Users\Public"), 'Documents', 'dat.sys'), None, None),
    # 文件哨兵3: 系统级程序数据目录 (需要权限，失败也无妨)
    ('file', os.path.join(os.environ.get("ProgramData", r"C:\ProgramData"), 'dat.sys'), None, None),
    # 注册表哨兵: 一个不显眼的公共位置
    ('reg', r"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.log", 'Signature', None)
]
# --- ↑↑↑ 新增代码结束 ↑↑↑ ---
EDGE_TTS_VOICES = {
    # --- 中国大陆 (8个) ---
    '在线-晓晓 (女)': 'zh-CN-XiaoxiaoNeural',
    '在线-云扬 (男)': 'zh-CN-YunyangNeural',  
    '在线-晓伊 (女)': 'zh-CN-XiaoyiNeural',
    '在线-云健 (男)': 'zh-CN-YunjianNeural',
    '在线-云希 (男)': 'zh-CN-YunxiNeural',
    '在线-云夏 (男)': 'zh-CN-YunxiaNeural',
    '在线-辽宁-晓北 (女)': 'zh-CN-liaoning-XiaobeiNeural',
    '在线-陕西-晓妮 (女)': 'zh-CN-shaanxi-XiaoniNeural',
    
    # --- 中国香港 (3个) ---
    '在线-香港-曉佳 (女)': 'zh-HK-HiuGaaiNeural',
    '在线-香港-曉曼 (女)': 'zh-HK-HiuMaanNeural',
    '在线-香港-雲龍 (男)': 'zh-HK-WanLungNeural',
    
    # --- 中国台湾 (3个) ---
    '在线-台湾-曉臻 (女)': 'zh-TW-HsiaoChenNeural',
    '在线-台湾-雲哲 (男)': 'zh-TW-YunJheNeural',
    '在线-台湾-曉雨 (女)': 'zh-TW-HsiaoYuNeural',
}


class TimedBroadcastApp:
    def __init__(self, root):
        self.root = root
        self.root.title(" 创翔多功能定时播音旗舰版")
        self.root.minsize(800, 600)

        if os.path.exists(ICON_FILE):
            try:
                self.root.iconbitmap(ICON_FILE)
            except Exception as e:
                print(f"加载窗口图标失败: {e}")

        self.tasks = []
        self.holidays = []
        self.todos = []
        self.screenshot_tasks = []
        self.execute_tasks = []
        self.print_tasks = []
        self.backup_tasks = []
        self.dynamic_voice_tasks = []
        
        self.settings = {}
        self.wallpaper_enabled_var = tk.BooleanVar()
        self.wallpaper_interval_days_var = tk.StringVar(value="1")
        self.wallpaper_change_time_var = tk.StringVar(value="08:00:00")
        self.wallpaper_cache_days_var = tk.StringVar(value="7")
        self.timer_mode_var = tk.StringVar(value="countdown") # 'countdown' 或 'stopwatch'
        self.timer_duration_var = tk.StringVar(value="00:10:00")
        self.timer_infinite_var = tk.BooleanVar(value=False)
        self.timer_show_clock_var = tk.BooleanVar(value=True)
        self.timer_play_sound_var = tk.BooleanVar(value=True)
        self.timer_sound_file_var = tk.StringVar(value="")
        
        # 用于管理计时器窗口的状态
        self.timer_window = None
        self.is_fullscreen_exclusive = False
        self.timer_after_id = None

        self.running = True
        self.tray_icon = None
        self.is_locked = False
        self.is_window_pinned = False
        self.is_app_locked_down = False
        self.active_modal_dialog = None

        self.auth_info = {'status': 'Unregistered', 'message': '正在验证授权...'}
        self.machine_code = None

        self.lock_password_b64 = ""
        self.drag_start_item = None

        self.playback_command_queue = queue.Queue()
        # --- ↓↓↓ 在这里添加新代码 ↓↓↓ ---
        self.intercut_queue = queue.Queue() # 新增：专门用于插播任务的队列
        self.intercut_stop_event = threading.Event() # 新增：用于紧急停止插播的信号
        # --- ↑↑↑ 新增代码结束 ↑↑↑ ---
        self.reminder_queue = queue.Queue()
        self.is_reminder_active = False

        self.pages = {}
        self.nav_buttons = {}
        self.current_page = None
        self.current_page_name = ""
        self.main_weather_label = None # <--- 新增此行
        self.intercut_page_content = None # <--- 新增：用于存储插播页面的文字内容
        
        self.active_processes = {}

        self.last_chime_hour = -1

        self.fullscreen_window = None
        self.fullscreen_label = None
        self.image_tk_ref = None
        self.current_stop_visual_event = None

        self.video_window = None
        self.vlc_player = None
        self.vlc_list_player = None
        self.video_stop_event = None
        self.is_muted = False
        self.last_bgm_volume = 1.0

        self.create_folder_structure()
        self.load_settings()

        saved_geometry = self.settings.get("window_geometry")
        if saved_geometry:
            try:
                self.root.geometry(saved_geometry)
            except tk.TclError:
                self.root.geometry("1280x720")
        else:
            self.root.geometry("1280x720")

        self.load_lock_password()

        self._apply_global_font()
        self.check_authorization()

        self.create_widgets()
        self.load_tasks()
        self.load_holidays()
        self.load_todos()
        self.load_screenshot_tasks()
        self.load_execute_tasks()
        self.load_print_tasks()
        self.load_backup_tasks()
        self.load_dynamic_voice_tasks()

        self.start_background_threads()
        self.root.protocol("WM_DELETE_WINDOW", self.show_quit_dialog)
        self.start_tray_icon_thread()

        if self.settings.get("lock_on_start", False) and self.lock_password_b64:
            self.root.after(100, self.perform_initial_lock)
        if self.settings.get("start_minimized", False):
            self.root.after(100, self.hide_to_tray)
        if self.is_app_locked_down:
            self.root.after(100, self.perform_lockdown)
        if self.auth_info['status'] == 'Trial':
            self.root.after(500, self.show_trial_nag_screen)

    def _apply_global_font(self):
        font_name = self.settings.get("app_font", "Microsoft YaHei")
        try:
            if font_name not in font.families():
                #self.log(f"警告：字体 '{font_name}' 未在系统中找到，已回退至默认字体。")
                font_name = "Microsoft YaHei"
                self.settings["app_font"] = font_name
        except Exception:
            font_name = "Microsoft YaHei"
        self.log(f"应用全局字体: {font_name}")

        self.font_8 = (font_name, 8)
        self.font_9 = (font_name, 9)
        self.font_10 = (font_name, 10)
        self.font_11 = (font_name, 11)
        self.font_11_bold = (font_name, 11, 'bold')
        self.font_12 = (font_name, 12)
        self.font_12_bold = (font_name, 12, 'bold')
        self.font_13_bold = (font_name, 13, 'bold')
        self.font_14_bold = (font_name, 14, 'bold')
        self.font_22_bold = (font_name, 22, 'bold')

        self.root.option_add("*Font", self.font_11)
        style = ttk.Style.get_instance()
        style.configure("TButton", font=self.font_11)
        style.configure("TLabel", font=self.font_11)
        style.configure("TCheckbutton", font=self.font_11)
        style.configure("TRadiobutton", font=self.font_11)
        style.configure("TCombobox", font=self.font_11)
        style.configure("TEntry", font=self.font_11)
        font_obj = font.Font(font=self.font_11)
        row_height = font_obj.metrics("linespace") + 10
        style.configure("Treeview", font=self.font_11, rowheight=row_height)
        style.configure("Treeview.Heading", font=self.font_11_bold)
        style.configure("TLabelframe.Label", font=self.font_12_bold)

    def _save_to_registry(self, key_name, value):
        if not WIN32_AVAILABLE: return False
        try:
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, REGISTRY_KEY_PATH)
            winreg.SetValueEx(key, key_name, 0, winreg.REG_SZ, str(value))
            winreg.CloseKey(key)
            return True
        except Exception as e:
            self.log(f"错误: 无法写入注册表项 '{key_name}' - {e}")
            return False

    def _load_from_registry(self, key_name):
        if not WIN32_AVAILABLE: return None
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REGISTRY_KEY_PATH, 0, winreg.KEY_READ)
            value, _ = winreg.QueryValueEx(key, key_name)
            winreg.CloseKey(key)
            return value
        except FileNotFoundError:
            return None
        except Exception as e:
            self.log(f"错误: 无法读取注册表项 '{key_name}' - {e}")
            return None

    def load_lock_password(self):
        self.lock_password_b64 = self._load_from_registry("LockPasswordB64") or ""

    def create_folder_structure(self):
        folders_to_create = [
            PROMPT_FOLDER, AUDIO_FOLDER, BGM_FOLDER, 
            VOICE_SCRIPT_FOLDER, SCREENSHOT_FOLDER,
            WALLPAPER_CACHE_FOLDER,
            DYNAMIC_VOICE_CACHE_FOLDER
        ]
        for folder in folders_to_create:
            if not os.path.exists(folder):
                os.makedirs(folder)

    def create_widgets(self):
        self.status_frame = ttk.Frame(self.root, style='secondary.TFrame')
        self.status_frame.pack(side=BOTTOM, fill=X)
        self.create_status_bar_content()

        self.nav_frame = ttk.Frame(self.root, width=160, style='light.TFrame')
        self.nav_frame.pack(side=LEFT, fill=Y)
        self.nav_frame.pack_propagate(False)

        self.page_container = ttk.Frame(self.root)
        self.page_container.pack(side=LEFT, fill=BOTH, expand=True)

        nav_button_titles = ["定时广播", "插播语音", "节假日", "待办事项", "高级功能", "设置", "注册软件", "超级管理"]

        for i, title in enumerate(nav_button_titles):
            is_super_admin = (title == "超级管理")
            cmd = (lambda t=title: self._prompt_for_super_admin_password()) if is_super_admin else (lambda t=title: self.switch_page(t))
            
            btn = ttk.Button(self.nav_frame, text=title, bootstyle="light",
                           style='Link.TButton', command=cmd)
            btn.pack(fill=X, pady=1, ipady=8, padx=5)
            self.nav_buttons[title] = btn

# --- ↓↓↓ 从这里开始替换 ↓↓↓ ---

        # 添加一个分隔符，让底部按钮和主导航分开
        ttk.Separator(self.nav_frame, orient=HORIZONTAL).pack(side=BOTTOM, fill=X, pady=5, padx=5)

# 创建一个Frame来容纳底部的按钮
        bottom_btn_frame = ttk.Frame(self.nav_frame, style='light.TFrame')
        bottom_btn_frame.pack(side=BOTTOM, fill=X, padx=5, pady=(0, 10))

        # --- ↓↓↓ 在这里添加新按钮 ↓↓↓ ---
        # 0. 一键停止按钮
        stop_button = ttk.Button(
            bottom_btn_frame, 
            text="一键停止", 
            bootstyle="warning-outline", 
            command=self.stop_current_playback  # 直接复用已有的方法
        )
        stop_button.pack(fill=X, pady=2)
        # --- ↑↑↑ 新增代码结束 ↑↑↑ ---

        # 1. 一键静音按钮
        self.mute_button = ttk.Button(
            bottom_btn_frame, 
            text="一键静音", 
            bootstyle="info-outline", 
            command=self.toggle_mute_all
        )
        self.mute_button.pack(fill=X, pady=2)

        # 2. 最小化按钮
        minimize_button = ttk.Button(
            bottom_btn_frame, 
            text="最小化", 
            bootstyle="secondary-outline", 
            command=self.hide_to_tray
        )
        minimize_button.pack(fill=X, pady=2)
        if not TRAY_AVAILABLE:
            minimize_button.config(state=DISABLED)

        # 3. 退出按钮
        exit_button = ttk.Button(
            bottom_btn_frame, 
            text="退出", 
            bootstyle="danger-outline", 
            command=self.quit_app
        )
        exit_button.pack(fill=X, pady=2)

        # --- ↑↑↑ 替换到这里结束 ↑↑↑ ---
            
        style = ttk.Style.get_instance()
        style.configure('Link.TButton', font=self.font_13_bold, anchor='w')

        self.main_frame = ttk.Frame(self.page_container)
        self.pages["定时广播"] = self.main_frame
        self.create_scheduled_broadcast_page()
        
        advanced_page = self.create_advanced_features_page()
        self.pages["高级功能"] = advanced_page
        advanced_page.pack_forget()

        self.current_page = self.main_frame
        self.switch_page("定时广播")

        self.update_status_bar()
        self.log("创翔多功能定时播音旗舰版软件已启动")

    def create_status_bar_content(self):
        self.status_labels = []
        status_texts = ["当前时间", "系统状态", "播放状态", "任务数量", "待办事项"]

        copyright_label = ttk.Label(self.status_frame, text="© 创翔科技 ver20251116", font=self.font_11,
                                    bootstyle=(SECONDARY, INVERSE), padding=(15, 0))
        copyright_label.pack(side=RIGHT, padx=2)

        self.statusbar_unlock_button = ttk.Button(self.status_frame, text="🔓 解锁",
                                                  bootstyle="success",
                                                  command=self._prompt_for_password_unlock)

        for i, text in enumerate(status_texts):
            label = ttk.Label(self.status_frame, text=f"{text}: --", font=self.font_11,
                              bootstyle=(PRIMARY, INVERSE) if i % 2 == 0 else (SECONDARY, INVERSE),
                              padding=(15, 5))
            label.pack(side=LEFT, padx=2, fill=Y)
            self.status_labels.append(label)

    def switch_page(self, page_name):
        if self.is_app_locked_down and page_name not in ["注册软件", "超级管理"]:
            self.log("软件授权已过期，请先注册。")
            if self.current_page_name != "注册软件":
                self.root.after(10, lambda: self.switch_page("注册软件"))
            return

        if self.is_locked and page_name not in ["超级管理", "注册软件"]:
            self.log("界面已锁定，请先解锁。")
            return

        if self.current_page and self.current_page.winfo_exists():
            self.current_page.pack_forget()

        for title, btn in self.nav_buttons.items():
            btn.config(bootstyle="light")

        target_frame = None
        if page_name in self.pages and self.pages[page_name].winfo_exists():
            target_frame = self.pages[page_name]
        else:
            if page_name == "插播语音":
                target_frame = self.create_intercut_page()
            elif page_name == "节假日":
                target_frame = self.create_holiday_page()
            elif page_name == "待办事项":
                target_frame = self.create_todo_page()
            elif page_name == "设置":
                target_frame = self.create_settings_page()
            elif page_name == "注册软件":
                target_frame = self.create_registration_page()
            elif page_name == "超级管理":
                target_frame = self.create_super_admin_page()
            
            if target_frame:
                self.pages[page_name] = target_frame

        if not target_frame:
            self.log(f"错误或开发中: 无法找到页面 '{page_name}'，返回主页。")
            target_frame = self.pages["定时广播"]
            page_name = "定时广播"
        
        target_frame.pack(in_=self.page_container, fill=BOTH, expand=True)
        self.current_page = target_frame
        self.current_page_name = page_name

        if page_name == "设置":
            self._refresh_settings_ui()
        if page_name == "高级功能":
            self._refresh_wallpaper_ui()
            self._refresh_timer_ui()

        selected_btn = self.nav_buttons.get(page_name)
        if selected_btn:
            selected_btn.config(bootstyle="primary")

    def _prompt_for_super_admin_password(self):
        if self.auth_info['status'] != 'Permanent':
            messagebox.showerror("权限不足", "此功能仅对“永久授权”用户开放。\n\n请注册软件并获取永久授权后重试。", parent=self.root)
            #self.log("非永久授权用户尝试进入超级管理模块被阻止。")
            return

        dialog = ttk.Toplevel(self.root)
        dialog.title("身份验证")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        
        # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()
        # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

        result = [None]

        ttk.Label(dialog, text="请输入超级管理员密码:", font=self.font_11).pack(pady=20, padx=20)
        password_entry = ttk.Entry(dialog, show='*', font=self.font_11, width=25)
        password_entry.pack(pady=5, padx=20)
        password_entry.focus_set()

        def on_confirm():
            result[0] = password_entry.get()
            cleanup_and_destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="确定", command=on_confirm, bootstyle="primary", width=8).pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=cleanup_and_destroy, width=8).pack(side=LEFT, padx=10)
        dialog.bind('<Return>', lambda event: on_confirm())
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)

        self.center_window(dialog, parent=self.root)
        self.root.wait_window(dialog)
        
        entered_password = result[0]
        correct_password = datetime.now().strftime('%Y%m%d')

        if entered_password == correct_password:
            #self.log("超级管理员密码正确，进入管理模块。")
            self.switch_page("超级管理")
        elif entered_password is not None:
            messagebox.showerror("验证失败", "密码错误！", parent=self.root)
            #self.log("尝试进入超级管理模块失败：密码错误。")

    def create_intercut_page(self):
        page_frame = ttk.Frame(self.page_container, padding=20)
        page_frame.columnconfigure(0, weight=1)
        page_frame.rowconfigure(1, weight=1)

        title_label = ttk.Label(page_frame, text="实时插播语音", font=self.font_14_bold, bootstyle="primary")
        title_label.grid(row=0, column=0, sticky='w', pady=(0, 15))

        # --- 播音文字区域 ---
        text_lf = ttk.LabelFrame(page_frame, text="播音文字", padding=10)
        text_lf.grid(row=1, column=0, sticky='nsew')
        text_lf.columnconfigure(0, weight=1)
        text_lf.rowconfigure(0, weight=1)

        content_text = ScrolledText(text_lf, height=8, font=self.font_11, wrap=WORD)
        content_text.grid(row=0, column=0, sticky='nsew')
        # 加载上次保存在 settings 中的内容
        content_text.text.insert('1.0', self.settings.get("intercut_text", ""))
        
        # --- 文稿操作按钮 ---
        script_btn_frame = ttk.Frame(text_lf)
        script_btn_frame.grid(row=1, column=0, sticky='w', pady=(10, 0))
        
        # 为了让 simpledialog 能正确显示在最前，父窗口需要是 root
        ttk.Button(script_btn_frame, text="导入文稿", command=lambda: self._import_voice_script(content_text.text, self.root), bootstyle="outline").pack(side=LEFT)
        ttk.Button(script_btn_frame, text="导出文稿", command=lambda: self._export_voice_script(content_text.text, None, self.root), bootstyle="outline").pack(side=LEFT, padx=10)

        # --- 播音员和参数设置 ---
        params_lf = ttk.LabelFrame(page_frame, text="播音参数", padding=15)
        params_lf.grid(row=2, column=0, sticky='ew', pady=15)
        params_lf.columnconfigure(1, weight=1)

        ttk.Label(params_lf, text="播音员:").grid(row=0, column=0, sticky='w')
        available_voices = self.get_available_voices()
        voice_var = tk.StringVar()
        voice_combo = ttk.Combobox(params_lf, textvariable=voice_var, values=available_voices, font=self.font_11, state='readonly')
        voice_combo.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5)
        if available_voices:
            voice_combo.set(available_voices[0])

        ttk.Label(params_lf, text="语速 (-10~10):").grid(row=1, column=0, sticky='w', pady=5)
        speed_entry = ttk.Entry(params_lf, font=self.font_11, width=10)
        speed_entry.insert(0, "0")
        speed_entry.grid(row=1, column=1, sticky='w', padx=5, pady=5)

        ttk.Label(params_lf, text="音调 (-10~10):").grid(row=2, column=0, sticky='w', pady=5)
        pitch_entry = ttk.Entry(params_lf, font=self.font_11, width=10)
        pitch_entry.insert(0, "0")
        pitch_entry.grid(row=2, column=1, sticky='w', padx=5, pady=5)
        
        # --- 立即插播按钮 ---
        intercut_btn = ttk.Button(page_frame, text="立即插播", style="lg.success.TButton", 
                                  command=lambda: self._execute_intercut(
                                      content_text.text.get('1.0', tk.END),
                                      voice_var.get(),
                                      speed_entry.get(),
                                      pitch_entry.get()
                                  ))
        intercut_btn.grid(row=3, column=0, sticky='ew', ipady=8, pady=10)
        
        return page_frame
            
    def create_advanced_features_page(self):
        page_frame = ttk.Frame(self.page_container, padding=10)

        page_frame.rowconfigure(1, weight=1)
        page_frame.columnconfigure(0, weight=1)

        title_label = ttk.Label(page_frame, text="高级功能", font=self.font_14_bold, bootstyle="primary")
        title_label.grid(row=0, column=0, sticky='w', pady=(0, 10))

        notebook = ttk.Notebook(page_frame, bootstyle="primary")
        notebook.grid(row=1, column=0, sticky='nsew', pady=5)

        screenshot_tab = ttk.Frame(notebook, padding=10)
        execute_tab = ttk.Frame(notebook, padding=10)
        print_tab = ttk.Frame(notebook, padding=10)
        backup_tab = ttk.Frame(notebook, padding=10)
        media_tab = ttk.Frame(notebook, padding=10)
        wallpaper_tab = ttk.Frame(notebook, padding=10)
        timer_tab = ttk.Frame(notebook, padding=10)

        notebook.add(screenshot_tab, text=' 定时截屏 ')
        notebook.add(execute_tab, text=' 定时运行 ')
        notebook.add(print_tab, text=' 定时打印 ')
        notebook.add(backup_tab, text=' 定时备份 ')
        notebook.add(media_tab, text=' 媒体处理 ')
        notebook.add(wallpaper_tab, text=' 网络壁纸 ')
        notebook.add(timer_tab, text=' 计时工具 ')

        self._build_screenshot_ui(screenshot_tab)
        self._build_execute_ui(execute_tab)
        self._build_print_ui(print_tab)
        self._build_backup_ui(backup_tab)
        self._build_media_processing_ui(media_tab)
        self._build_wallpaper_ui(wallpaper_tab)
        self._build_timer_ui(timer_tab)

        return page_frame

    def enable_all_screenshot(self):
        if not self.screenshot_tasks: return
        for task in self.screenshot_tasks:
            task['status'] = '启用'
        self.update_screenshot_list()
        self.save_screenshot_tasks()
        self.log("已将 *全部* 截屏任务的状态设置为: 启用")

    def disable_all_screenshot(self):
        if not self.screenshot_tasks: return
        for task in self.screenshot_tasks:
            task['status'] = '禁用'
        self.update_screenshot_list()
        self.save_screenshot_tasks()
        self.log("已将 *全部* 截屏任务的状态设置为: 禁用")

    def enable_all_execute(self):
        if not self.execute_tasks: return
        for task in self.execute_tasks:
            task['status'] = '启用'
        self.update_execute_list()
        self.save_execute_tasks()
        self.log("已将 *全部* 运行任务的状态设置为: 启用")

    def disable_all_execute(self):
        if not self.execute_tasks: return
        for task in self.execute_tasks:
            task['status'] = '禁用'
        self.update_execute_list()
        self.save_execute_tasks()
        self.log("已将 *全部* 运行任务的状态设置为: 禁用")

    def update_screenshot_list(self):
        if not hasattr(self, 'screenshot_tree') or not self.screenshot_tree.winfo_exists(): return
        self.screenshot_tree.delete(*self.screenshot_tree.get_children())
        for task in self.screenshot_tasks:
            self.screenshot_tree.insert('', END, values=(
                task.get('name', ''),
                task.get('status', '启用'),
                task.get('time', ''),
                task.get('stop_time', ''),
                task.get('repeat_count', 1),
                task.get('interval_seconds', 0),
                task.get('weekday', ''),
                task.get('date_range', '')
            ))

    def add_screenshot_task(self):
        self.open_screenshot_dialog()

    def edit_screenshot_task(self):
        selection = self.screenshot_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要修改的截屏任务", parent=self.root)
            return
        index = self.screenshot_tree.index(selection[0])
        task_to_edit = self.screenshot_tasks[index]
        self.open_screenshot_dialog(task_to_edit=task_to_edit, index=index)

    def delete_screenshot_task(self):
        selections = self.screenshot_tree.selection()
        if not selections:
            messagebox.showwarning("提示", "请先选择要删除的截屏任务", parent=self.root)
            return
        if messagebox.askyesno("确认删除", f"确定要删除选中的 {len(selections)} 个截屏任务吗？", parent=self.root):
            indices = sorted([self.screenshot_tree.index(s) for s in selections], reverse=True)
            for index in indices:
                self.screenshot_tasks.pop(index)
            self.update_screenshot_list()
            self.save_screenshot_tasks()

    def clear_all_screenshot_tasks(self):
        if not self.screenshot_tasks: return
        if messagebox.askyesno("确认清空", "您确定要清空所有截屏任务吗？", parent=self.root):
            self.screenshot_tasks.clear()
            self.update_screenshot_list()
            self.save_screenshot_tasks()

    def _set_screenshot_status(self, status):
        selection = self.screenshot_tree.selection()
        if not selection:
            messagebox.showwarning("提示", f"请先选择要 {status} 的任务", parent=self.root)
            return
        for item_id in selection:
            index = self.screenshot_tree.index(item_id)
            self.screenshot_tasks[index]['status'] = status
        self.update_screenshot_list()
        self.save_screenshot_tasks()

    # --- ↓↓↓ 新增代码：为“定时截屏”列表添加右键菜单及相关操作函数 ↓↓↓ ---

    def show_screenshot_context_menu(self, event):
        if self.is_locked: return
        iid = self.screenshot_tree.identify_row(event.y)
        context_menu = tk.Menu(self.root, tearoff=0, font=self.font_11)

        if iid:
            if iid not in self.screenshot_tree.selection():
                self.screenshot_tree.selection_set(iid)

            context_menu.add_command(label="修改", command=self.edit_screenshot_task)
            context_menu.add_command(label="删除", command=self.delete_screenshot_task)
            context_menu.add_separator()
            context_menu.add_command(label="置顶", command=self.move_screenshot_to_top)
            context_menu.add_command(label="上移", command=lambda: self.move_screenshot_task(-1))
            context_menu.add_command(label="下移", command=lambda: self.move_screenshot_task(1))
            context_menu.add_command(label="置末", command=self.move_screenshot_to_bottom)
            context_menu.add_separator()
            context_menu.add_command(label="启用", command=lambda: self._set_screenshot_status('启用'))
            context_menu.add_command(label="禁用", command=lambda: self._set_screenshot_status('禁用'))
        else:
            self.screenshot_tree.selection_set()
            context_menu.add_command(label="添加任务", command=self.add_screenshot_task)

        context_menu.post(event.x_root, event.y_root)

    def move_screenshot_task(self, direction):
        selection = self.screenshot_tree.selection()
        if not selection or len(selection) > 1: return
        index = self.screenshot_tree.index(selection[0])
        new_index = index + direction
        if 0 <= new_index < len(self.screenshot_tasks):
            task_to_move = self.screenshot_tasks.pop(index)
            self.screenshot_tasks.insert(new_index, task_to_move)
            self.update_screenshot_list(); self.save_screenshot_tasks()
            items = self.screenshot_tree.get_children()
            if items: self.screenshot_tree.selection_set(items[new_index]); self.screenshot_tree.focus(items[new_index])

    def move_screenshot_to_top(self):
        selection = self.screenshot_tree.selection()
        if not selection or len(selection) > 1: return
        index = self.screenshot_tree.index(selection[0])
        if index > 0:
            task_to_move = self.screenshot_tasks.pop(index)
            self.screenshot_tasks.insert(0, task_to_move)
            self.update_screenshot_list(); self.save_screenshot_tasks()
            items = self.screenshot_tree.get_children()
            if items: self.screenshot_tree.selection_set(items[0]); self.screenshot_tree.focus(items[0])

    def move_screenshot_to_bottom(self):
        selection = self.screenshot_tree.selection()
        if not selection or len(selection) > 1: return
        index = self.screenshot_tree.index(selection[0])
        if index < len(self.screenshot_tasks) - 1:
            task_to_move = self.screenshot_tasks.pop(index)
            self.screenshot_tasks.append(task_to_move)
            self.update_screenshot_list(); self.save_screenshot_tasks()
            items = self.screenshot_tree.get_children()
            if items: self.screenshot_tree.selection_set(items[-1]); self.screenshot_tree.focus(items[-1])

    # --- ↑↑↑ 新增代码结束 ↑↑↑ ---

    def open_screenshot_dialog(self, task_to_edit=None, index=None):
        dialog = ttk.Toplevel(self.root)
        dialog.title("修改截屏任务" if task_to_edit else "添加截屏任务")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()
        # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=BOTH, expand=True)

        content_frame = ttk.LabelFrame(main_frame, text="内容", padding=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=2)
        content_frame.columnconfigure(1, weight=1)
        
        ttk.Label(content_frame, text="任务名称:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        name_entry = ttk.Entry(content_frame, font=self.font_11)
        name_entry.grid(row=0, column=1, columnspan=2, sticky='ew', padx=5, pady=2)
        
        ttk.Label(content_frame, text="截取张数:").grid(row=1, column=0, sticky='e', padx=5, pady=2)
        repeat_entry = ttk.Entry(content_frame, font=self.font_11)
        repeat_entry.grid(row=1, column=1, sticky='w', pady=2)
        
        ttk.Label(content_frame, text="间隔(秒):").grid(row=2, column=0, sticky='e', padx=5, pady=2)
        interval_entry = ttk.Entry(content_frame, font=self.font_11)
        interval_entry.grid(row=2, column=1, sticky='w', pady=2)

        time_frame = ttk.LabelFrame(main_frame, text="时间", padding=15)
        time_frame.grid(row=1, column=0, sticky='ew', pady=4)
        time_frame.columnconfigure(1, weight=1)
        
        ttk.Label(time_frame, text="开始时间:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        start_time_entry = ttk.Entry(time_frame, font=self.font_11)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        self._bind_mousewheel_to_entry(start_time_entry, self._handle_time_scroll)
        ttk.Label(time_frame, text="<可多个>").grid(row=0, column=2, sticky='w', padx=5)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bootstyle="outline").grid(row=0, column=3, padx=5)

        ttk.Label(time_frame, text="停止时间:").grid(row=1, column=0, sticky='e', padx=5, pady=2)
        stop_time_entry = ttk.Entry(time_frame, font=self.font_11)
        stop_time_entry.grid(row=1, column=1, sticky='w', padx=5, pady=2)
        self._bind_mousewheel_to_entry(stop_time_entry, self._handle_time_scroll)
        ttk.Label(time_frame, text="(可选)").grid(row=1, column=2, sticky='w')
        
        ttk.Label(time_frame, text="周几/几号:").grid(row=2, column=0, sticky='e', padx=5, pady=3)
        weekday_entry = ttk.Entry(time_frame, font=self.font_11)
        weekday_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=3)
        ttk.Button(time_frame, text="选取...", command=lambda: self.show_weekday_settings_dialog(weekday_entry), bootstyle="outline").grid(row=2, column=3, padx=5)
        
        ttk.Label(time_frame, text="日期范围:").grid(row=3, column=0, sticky='e', padx=5, pady=3)
        date_range_entry = ttk.Entry(time_frame, font=self.font_11)
        date_range_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=3)
        self._bind_mousewheel_to_entry(date_range_entry, self._handle_date_scroll)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry), bootstyle="outline").grid(row=3, column=3, padx=5)

        dialog_button_frame = ttk.Frame(dialog)
        dialog_button_frame.pack(pady=15)

        if task_to_edit:
            name_entry.insert(0, task_to_edit.get('name', ''))
            start_time_entry.insert(0, task_to_edit.get('time', ''))
            stop_time_entry.insert(0, task_to_edit.get('stop_time', ''))
            repeat_entry.insert(0, task_to_edit.get('repeat_count', 1))
            interval_entry.insert(0, task_to_edit.get('interval_seconds', 0))
            weekday_entry.insert(0, task_to_edit.get('weekday', '每周:1234567'))
            date_range_entry.insert(0, task_to_edit.get('date_range', '2025-01-01 ~ 2099-12-31'))
        else:
            repeat_entry.insert(0, '1')
            interval_entry.insert(0, '0')
            weekday_entry.insert(0, "每周:1234567")
            date_range_entry.insert(0, "2025-01-01 ~ 2099-12-31")

        def save_task():
            # --- ↓↓↓ 新增的输入验证模块 ↓↓↓ ---
            try:
                repeat_count = int(repeat_entry.get().strip() or 1)
                if repeat_count < 1:
                    messagebox.showerror("输入错误", "“截取张数”必须是大于或等于 1 的整数。", parent=dialog)
                    return
            except ValueError:
                messagebox.showerror("输入错误", "“截取张数”必须是一个有效的整数。", parent=dialog)
                return

            try:
                interval_seconds = int(interval_entry.get().strip() or 0)
                if interval_seconds < 0:
                    messagebox.showerror("输入错误", "“间隔(秒)”必须是大于或等于 0 的整数。", parent=dialog)
                    return
            except ValueError:
                messagebox.showerror("输入错误", "“间隔(秒)”必须是一个有效的整数。", parent=dialog)
                return

            if not weekday_entry.get().strip():
                messagebox.showerror("输入错误", "“周几/几号”规则不能为空，请点击“选取...”进行设置。", parent=dialog)
                return
            
            if not date_range_entry.get().strip():
                messagebox.showerror("输入错误", "“日期范围”不能为空，请点击“设置...”进行配置。", parent=dialog)
                return
            # --- ↑↑↑ 验证模块结束 ↑↑↑ ---
            
            is_valid_time, time_msg = self._normalize_multiple_times_string(start_time_entry.get().strip())
            if not is_valid_time: messagebox.showwarning("格式错误", time_msg, parent=dialog); return
            is_valid_date, date_msg = self._normalize_date_range_string(date_range_entry.get().strip())
            if not is_valid_date: messagebox.showwarning("格式错误", date_msg, parent=dialog); return

            new_task_data = {
                'name': name_entry.get().strip(), 'time': time_msg,
                'stop_time': self._normalize_time_string(stop_time_entry.get().strip()) or "",
                'repeat_count': int(repeat_entry.get().strip() or 1),
                'interval_seconds': int(interval_entry.get().strip() or 0),
                'weekday': weekday_entry.get().strip(), 'date_range': date_msg,
                'status': '启用' if not task_to_edit else task_to_edit.get('status', '启用'),
                'last_run': {} if not task_to_edit else task_to_edit.get('last_run', {}),
            }
            if not new_task_data['name'] or not new_task_data['time']: 
                messagebox.showwarning("警告", "请填写任务名称和开始时间", parent=dialog); return

            if task_to_edit:
                self.screenshot_tasks[index] = new_task_data
                self.log(f"已修改截屏任务: {new_task_data['name']}")
            else:
                self.screenshot_tasks.append(new_task_data)
                self.log(f"已添加截屏任务: {new_task_data['name']}")

            self.update_screenshot_list()
            self.save_screenshot_tasks()
            cleanup_and_destroy()

        button_text = "保存修改" if task_to_edit else "添加"
        ttk.Button(dialog_button_frame, text=button_text, command=save_task, bootstyle="primary").pack(side=LEFT, padx=10, ipady=5)
        ttk.Button(dialog_button_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10, ipady=5)
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)
        
        self.center_window(dialog, parent=self.root)

    def _build_screenshot_ui(self, parent_frame):
        parent_frame.columnconfigure(0, weight=1)
        parent_frame.rowconfigure(0, weight=1)

        main_content_frame = ttk.Frame(parent_frame)
        main_content_frame.grid(row=0, column=0, sticky='nsew')
        main_content_frame.columnconfigure(0, weight=1)
        main_content_frame.rowconfigure(1, weight=1)

        desc_label = ttk.Label(main_content_frame, 
                               text=f"此功能将在指定时间自动截取全屏图像，并以PNG格式保存到以下目录：\n{SCREENSHOT_FOLDER}",
                               font=self.font_10, bootstyle="secondary", wraplength=600)
        desc_label.grid(row=0, column=0, sticky='w', pady=(0, 10))

        table_frame = ttk.Frame(main_content_frame)
        table_frame.grid(row=1, column=0, sticky='nsew')
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        columns = ('任务名称', '状态', '开始时间', '停止时间', '截取张数', '间隔(秒)', '周/月规则', '日期范围')
        self.screenshot_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15, selectmode='extended', bootstyle="info")
        
        col_configs = [
            ('任务名称', 200, 'w'), ('状态', 80, 'center'), ('开始时间', 150, 'center'),
            ('停止时间', 100, 'center'), ('截取张数', 80, 'center'), ('间隔(秒)', 80, 'center'), 
            ('周/月规则', 150, 'center'), ('日期范围', 200, 'center')
        ]
        for name, width, anchor in col_configs:
            self.screenshot_tree.heading(name, text=name)
            self.screenshot_tree.column(name, width=width, anchor=anchor)

        self.screenshot_tree.grid(row=0, column=0, sticky='nsew')
        scrollbar = ttk.Scrollbar(table_frame, orient=VERTICAL, command=self.screenshot_tree.yview, bootstyle="round-info")
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.screenshot_tree.configure(yscrollcommand=scrollbar.set)

        self.screenshot_tree.bind("<Double-1>", lambda e: self.edit_screenshot_task())
        self.screenshot_tree.bind("<Button-3>", self.show_screenshot_context_menu) # <--- 添加这一行

        action_frame = ttk.Frame(parent_frame, padding=(10, 0))
        action_frame.grid(row=0, column=1, sticky='ns', padx=(10, 0))

        buttons_config = [
            ("添加任务", self.add_screenshot_task, "info"),
            ("修改任务", self.edit_screenshot_task, "success"),
            ("删除任务", self.delete_screenshot_task, "danger"),
            (None, None, None),
            ("全部启用", self.enable_all_screenshot, "outline-success"),
            ("全部禁用", self.disable_all_screenshot, "outline-warning"),
            ("清空列表", self.clear_all_screenshot_tasks, "outline-danger")
        ]
        for text, cmd, style in buttons_config:
            if text is None:
                ttk.Separator(action_frame, orient=HORIZONTAL).pack(fill=X, pady=10)
                continue
            ttk.Button(action_frame, text=text, command=cmd, bootstyle=style).pack(pady=5, fill=X)
            
        self.update_screenshot_list()
        
#第1部分
    def _build_execute_ui(self, parent_frame):
        if not PSUTIL_AVAILABLE:
            ttk.Label(parent_frame, text="错误：psutil 库未安装，无法使用此功能。", font=self.font_12_bold, bootstyle="danger").pack(pady=50)
            return

        parent_frame.columnconfigure(0, weight=1)
        parent_frame.rowconfigure(0, weight=1)

        main_content_frame = ttk.Frame(parent_frame)
        main_content_frame.grid(row=0, column=0, sticky='nsew')
        main_content_frame.columnconfigure(0, weight=1)
        main_content_frame.rowconfigure(1, weight=1)

        warning_label = ttk.Label(main_content_frame, 
                                  text="/!\\ 警告：请确保您完全信任所要运行的程序。运行未知或恶意程序可能对您的计算机安全造成严重威胁。",
                                  font=self.font_10, bootstyle="danger", wraplength=600)
        warning_label.grid(row=0, column=0, sticky='w', pady=(0, 10))

        table_frame = ttk.Frame(main_content_frame)
        table_frame.grid(row=1, column=0, sticky='nsew')
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        columns = ('任务名称', '状态', '执行时间', '停止时间', '目标程序', '参数', '周/月规则', '日期范围')
        self.execute_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15, selectmode='extended', bootstyle="danger")
        
        col_configs = [
            ('任务名称', 200, 'w'), ('状态', 80, 'center'), ('执行时间', 150, 'center'),
            ('停止时间', 100, 'center'), ('目标程序', 250, 'w'), ('参数', 150, 'w'),
            ('周/月规则', 150, 'center'), ('日期范围', 200, 'center')
        ]
        for name, width, anchor in col_configs:
            self.execute_tree.heading(name, text=name)
            self.execute_tree.column(name, width=width, anchor=anchor)

        self.execute_tree.grid(row=0, column=0, sticky='nsew')
        scrollbar = ttk.Scrollbar(table_frame, orient=VERTICAL, command=self.execute_tree.yview, bootstyle="round-danger")
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.execute_tree.configure(yscrollcommand=scrollbar.set)

        self.execute_tree.bind("<Double-1>", lambda e: self.edit_execute_task())
        self.execute_tree.bind("<Button-3>", self.show_execute_context_menu) # <--- 添加这一行

        action_frame = ttk.Frame(parent_frame, padding=(10, 0))
        action_frame.grid(row=0, column=1, sticky='ns', padx=(10, 0))

        buttons_config = [
            ("添加任务", self.add_execute_task, "info"),
            ("修改任务", self.edit_execute_task, "success"),
            ("删除任务", self.delete_execute_task, "danger"),
            (None, None, None),
            ("全部启用", self.enable_all_execute, "outline-success"),
            ("全部禁用", self.disable_all_execute, "outline-warning"),
            ("清空列表", self.clear_all_execute_tasks, "outline-danger")
        ]
        for text, cmd, style in buttons_config:
            if text is None:
                ttk.Separator(action_frame, orient=HORIZONTAL).pack(fill=X, pady=10)
                continue
            ttk.Button(action_frame, text=text, command=cmd, bootstyle=style).pack(pady=5, fill=X)
            
        self.update_execute_list()

    def _build_print_ui(self, parent_frame):
        parent_frame.columnconfigure(0, weight=1)
        parent_frame.rowconfigure(0, weight=1)

        main_content_frame = ttk.Frame(parent_frame)
        main_content_frame.grid(row=0, column=0, sticky='nsew')
        main_content_frame.columnconfigure(0, weight=1)
        main_content_frame.rowconfigure(1, weight=1)

        desc_label = ttk.Label(main_content_frame, 
                               text="此功能将在指定时间，使用指定打印机自动打印文件。",
                               font=self.font_10, bootstyle="info", wraplength=600)
        desc_label.grid(row=0, column=0, sticky='w', pady=(0, 10))

        table_frame = ttk.Frame(main_content_frame)
        table_frame.grid(row=1, column=0, sticky='nsew')
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        columns = ('任务名称', '状态', '打印时间', '打印文件', '打印机', '份数', '周/月规则', '日期范围')
        self.print_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15, selectmode='extended', bootstyle="info")
        
        col_configs = [
            ('任务名称', 200, 'w'), ('状态', 80, 'center'), ('打印时间', 150, 'center'),
            ('打印文件', 250, 'w'), ('打印机', 200, 'w'), ('份数', 60, 'center'),
            ('周/月规则', 150, 'center'), ('日期范围', 200, 'center')
        ]
        for name, width, anchor in col_configs:
            self.print_tree.heading(name, text=name)
            self.print_tree.column(name, width=width, anchor=anchor)

        self.print_tree.grid(row=0, column=0, sticky='nsew')
        scrollbar = ttk.Scrollbar(table_frame, orient=VERTICAL, command=self.print_tree.yview, bootstyle="round-info")
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.print_tree.configure(yscrollcommand=scrollbar.set)

        self.print_tree.bind("<Double-1>", lambda e: self.edit_print_task())
        self.print_tree.bind("<Button-3>", self.show_print_context_menu)

        action_frame = ttk.Frame(parent_frame, padding=(10, 0))
        action_frame.grid(row=0, column=1, sticky='ns', padx=(10, 0))

        buttons_config = [
            ("添加任务", self.add_print_task, "info"),
            ("修改任务", self.edit_print_task, "success"),
            ("删除任务", self.delete_print_task, "danger"),
            (None, None, None),
            ("全部启用", self.enable_all_print, "outline-success"),
            ("全部禁用", self.disable_all_print, "outline-warning"),
            ("清空列表", self.clear_all_print_tasks, "outline-danger")
        ]
        for text, cmd, style in buttons_config:
            if text is None:
                ttk.Separator(action_frame, orient=HORIZONTAL).pack(fill=X, pady=10)
                continue
            ttk.Button(action_frame, text=text, command=cmd, bootstyle=style).pack(pady=5, fill=X)
            
        self.update_print_list()

    def _build_backup_ui(self, parent_frame):
        parent_frame.columnconfigure(0, weight=1)
        parent_frame.rowconfigure(0, weight=1)

        main_content_frame = ttk.Frame(parent_frame)
        main_content_frame.grid(row=0, column=0, sticky='nsew')
        main_content_frame.columnconfigure(0, weight=1)
        main_content_frame.rowconfigure(1, weight=1)

        desc_label = ttk.Label(main_content_frame, 
                               text="此功能使用Windows内置的Robocopy命令，在指定时间将源文件夹备份到目标文件夹。",
                               font=self.font_10, bootstyle="info", wraplength=600)
        desc_label.grid(row=0, column=0, sticky='w', pady=(0, 10))

        table_frame = ttk.Frame(main_content_frame)
        table_frame.grid(row=1, column=0, sticky='nsew')
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        columns = ('任务名称', '状态', '备份时间', '源文件夹', '目标文件夹', '模式', '周/月规则', '日期范围')
        self.backup_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15, selectmode='extended', bootstyle="success")
        
        col_configs = [
            ('任务名称', 200, 'w'), ('状态', 80, 'center'), ('备份时间', 150, 'center'),
            ('源文件夹', 250, 'w'), ('目标文件夹', 250, 'w'), ('模式', 80, 'center'),
            ('周/月规则', 150, 'center'), ('日期范围', 200, 'center')
        ]
        for name, width, anchor in col_configs:
            self.backup_tree.heading(name, text=name)
            self.backup_tree.column(name, width=width, anchor=anchor)

        self.backup_tree.grid(row=0, column=0, sticky='nsew')
        scrollbar = ttk.Scrollbar(table_frame, orient=VERTICAL, command=self.backup_tree.yview, bootstyle="round-success")
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.backup_tree.configure(yscrollcommand=scrollbar.set)

        self.backup_tree.bind("<Double-1>", lambda e: self.edit_backup_task())
        self.backup_tree.bind("<Button-3>", self.show_backup_context_menu)

        action_frame = ttk.Frame(parent_frame, padding=(10, 0))
        action_frame.grid(row=0, column=1, sticky='ns', padx=(10, 0))

        buttons_config = [
            ("添加任务", self.add_backup_task, "info"),
            ("修改任务", self.edit_backup_task, "success"),
            ("删除任务", self.delete_backup_task, "danger"),
            (None, None, None),
            ("全部启用", self.enable_all_backup, "outline-success"),
            ("全部禁用", self.disable_all_backup, "outline-warning"),
            ("清空列表", self.clear_all_backup_tasks, "outline-danger")
        ]
        for text, cmd, style in buttons_config:
            if text is None:
                ttk.Separator(action_frame, orient=HORIZONTAL).pack(fill=X, pady=10)
                continue
            ttk.Button(action_frame, text=text, command=cmd, bootstyle=style).pack(pady=5, fill=X)
            
        self.update_backup_list()

    # --- 定时运行功能的全套方法 ---
    
    def load_execute_tasks(self):
        if not os.path.exists(EXECUTE_TASK_FILE): return
        try:
            with open(EXECUTE_TASK_FILE, 'r', encoding='utf-8') as f:
                self.execute_tasks = json.load(f)
            self.log(f"已加载 {len(self.execute_tasks)} 个运行任务")
            if hasattr(self, 'execute_tree'):
                self.update_execute_list()
        except Exception as e:
            self.log(f"加载运行任务失败: {e}")
            self.execute_tasks = []

    def save_execute_tasks(self):
        try:
            with open(EXECUTE_TASK_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.execute_tasks, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"保存运行任务失败: {e}")

    def update_execute_list(self):
        if not hasattr(self, 'execute_tree') or not self.execute_tree.winfo_exists(): return
        self.execute_tree.delete(*self.execute_tree.get_children())
        for task in self.execute_tasks:
            self.execute_tree.insert('', END, values=(
                task.get('name', ''),
                task.get('status', '启用'),
                task.get('time', ''),
                task.get('stop_time', ''),
                os.path.basename(task.get('target_path', '')),
                task.get('arguments', ''),
                task.get('weekday', ''),
                task.get('date_range', '')
            ))

    def add_execute_task(self):
        self.open_execute_dialog()

    def edit_execute_task(self):
        selection = self.execute_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要修改的运行任务", parent=self.root)
            return
        index = self.execute_tree.index(selection[0])
        task_to_edit = self.execute_tasks[index]
        self.open_execute_dialog(task_to_edit=task_to_edit, index=index)

    def delete_execute_task(self):
        selections = self.execute_tree.selection()
        if not selections:
            messagebox.showwarning("提示", "请先选择要删除的运行任务", parent=self.root)
            return
        if messagebox.askyesno("确认删除", f"确定要删除选中的 {len(selections)} 个运行任务吗？", parent=self.root):
            indices = sorted([self.execute_tree.index(s) for s in selections], reverse=True)
            for index in indices:
                self.execute_tasks.pop(index)
            self.update_execute_list()
            self.save_execute_tasks()

    def clear_all_execute_tasks(self):
        if not self.execute_tasks: return
        if messagebox.askyesno("确认清空", "您确定要清空所有运行任务吗？", parent=self.root):
            self.execute_tasks.clear()
            self.update_execute_list()
            self.save_execute_tasks()

    def _set_execute_status(self, status):
        selection = self.execute_tree.selection()
        if not selection:
            messagebox.showwarning("提示", f"请先选择要 {status} 的任务", parent=self.root)
            return
        for item_id in selection:
            index = self.execute_tree.index(item_id)
            self.execute_tasks[index]['status'] = status
        self.update_execute_list()
        self.save_execute_tasks()

    # --- ↓↓↓ 新增代码：为“定时运行”列表添加右键菜单及相关操作函数 ↓↓↓ ---

    def show_execute_context_menu(self, event):
        if self.is_locked: return
        iid = self.execute_tree.identify_row(event.y)
        context_menu = tk.Menu(self.root, tearoff=0, font=self.font_11)

        if iid:
            if iid not in self.execute_tree.selection():
                self.execute_tree.selection_set(iid)

            context_menu.add_command(label="修改", command=self.edit_execute_task)
            context_menu.add_command(label="删除", command=self.delete_execute_task)
            context_menu.add_separator()
            context_menu.add_command(label="置顶", command=self.move_execute_to_top)
            context_menu.add_command(label="上移", command=lambda: self.move_execute_task(-1))
            context_menu.add_command(label="下移", command=lambda: self.move_execute_task(1))
            context_menu.add_command(label="置末", command=lambda: self.move_execute_to_bottom)
            context_menu.add_separator()
            context_menu.add_command(label="启用", command=lambda: self._set_execute_status('启用'))
            context_menu.add_command(label="禁用", command=lambda: self._set_execute_status('禁用'))
        else:
            self.execute_tree.selection_set()
            context_menu.add_command(label="添加任务", command=self.add_execute_task)

        context_menu.post(event.x_root, event.y_root)

    def move_execute_task(self, direction):
        selection = self.execute_tree.selection()
        if not selection or len(selection) > 1: return
        index = self.execute_tree.index(selection[0])
        new_index = index + direction
        if 0 <= new_index < len(self.execute_tasks):
            task_to_move = self.execute_tasks.pop(index)
            self.execute_tasks.insert(new_index, task_to_move)
            self.update_execute_list(); self.save_execute_tasks()
            items = self.execute_tree.get_children()
            if items: self.execute_tree.selection_set(items[new_index]); self.execute_tree.focus(items[new_index])

    def move_execute_to_top(self):
        selection = self.execute_tree.selection()
        if not selection or len(selection) > 1: return
        index = self.execute_tree.index(selection[0])
        if index > 0:
            task_to_move = self.execute_tasks.pop(index)
            self.execute_tasks.insert(0, task_to_move)
            self.update_execute_list(); self.save_execute_tasks()
            items = self.execute_tree.get_children()
            if items: self.execute_tree.selection_set(items[0]); self.execute_tree.focus(items[0])

    def move_execute_to_bottom(self):
        selection = self.execute_tree.selection()
        if not selection or len(selection) > 1: return
        index = self.execute_tree.index(selection[0])
        if index < len(self.execute_tasks) - 1:
            task_to_move = self.execute_tasks.pop(index)
            self.execute_tasks.append(task_to_move)
            self.update_execute_list(); self.save_execute_tasks()
            items = self.execute_tree.get_children()
            if items: self.execute_tree.selection_set(items[-1]); self.execute_tree.focus(items[-1])

    # --- ↑↑↑ 新增代码结束 ↑↑↑ ---```

    def open_execute_dialog(self, task_to_edit=None, index=None):
        dialog = ttk.Toplevel(self.root)
        dialog.title("修改运行任务" if task_to_edit else "添加运行任务")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()
        # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=BOTH, expand=True)

        content_frame = ttk.LabelFrame(main_frame, text="内容", padding=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=2)
        content_frame.columnconfigure(1, weight=1)
        
        ttk.Label(content_frame, text="任务名称:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        name_entry = ttk.Entry(content_frame, font=self.font_11)
        name_entry.grid(row=0, column=1, columnspan=2, sticky='ew', padx=5, pady=2)

        ttk.Label(content_frame, text="目标程序:").grid(row=1, column=0, sticky='e', padx=5, pady=2)
        target_entry = ttk.Entry(content_frame, font=self.font_11)
        target_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=2)
        def select_target():
            path = filedialog.askopenfilename(title="选择可执行文件", filetypes=[("可执行文件", "*.exe *.bat *.cmd"), ("所有文件", "*.*")], parent=dialog)
            if path:
                target_entry.delete(0, END)
                target_entry.insert(0, path)
        ttk.Button(content_frame, text="浏览...", command=select_target, bootstyle="outline").grid(row=1, column=2, padx=5)

        ttk.Label(content_frame, text="命令行参数:").grid(row=2, column=0, sticky='e', padx=5, pady=2)
        args_entry = ttk.Entry(content_frame, font=self.font_11)
        args_entry.grid(row=2, column=1, columnspan=2, sticky='ew', padx=5, pady=2)
        ttk.Label(content_frame, text="(可选)", font=self.font_9, bootstyle="secondary").grid(row=3, column=1, sticky='w', padx=5)

        time_frame = ttk.LabelFrame(main_frame, text="时间", padding=15)
        time_frame.grid(row=1, column=0, sticky='ew', pady=4)
        time_frame.columnconfigure(1, weight=1)
        
        ttk.Label(time_frame, text="执行时间:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        start_time_entry = ttk.Entry(time_frame, font=self.font_11)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        self._bind_mousewheel_to_entry(start_time_entry, self._handle_time_scroll)
        ttk.Label(time_frame, text="<可多个>").grid(row=0, column=2, sticky='w', padx=5)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bootstyle="outline").grid(row=0, column=3, padx=5)

        ttk.Label(time_frame, text="停止时间:").grid(row=1, column=0, sticky='e', padx=5, pady=2)
        stop_time_entry = ttk.Entry(time_frame, font=self.font_11)
        stop_time_entry.grid(row=1, column=1, sticky='w', padx=5, pady=2)
        self._bind_mousewheel_to_entry(stop_time_entry, self._handle_time_scroll)
        ttk.Label(time_frame, text="(可选, 到达此时间将强制终止进程)").grid(row=1, column=2, columnspan=2, sticky='w')
        
        ttk.Label(time_frame, text="周几/几号:").grid(row=2, column=0, sticky='e', padx=5, pady=3)
        weekday_entry = ttk.Entry(time_frame, font=self.font_11)
        weekday_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=3)
        ttk.Button(time_frame, text="选取...", command=lambda: self.show_weekday_settings_dialog(weekday_entry), bootstyle="outline").grid(row=2, column=3, padx=5)
        
        ttk.Label(time_frame, text="日期范围:").grid(row=3, column=0, sticky='e', padx=5, pady=3)
        date_range_entry = ttk.Entry(time_frame, font=self.font_11)
        date_range_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=3)
        self._bind_mousewheel_to_entry(date_range_entry, self._handle_date_scroll)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry), bootstyle="outline").grid(row=3, column=3, padx=5)

        warning_frame = ttk.LabelFrame(main_frame, text="风险警告", padding=10, bootstyle="danger")
        warning_frame.grid(row=2, column=0, sticky='ew', pady=10)
        ttk.Label(warning_frame, text="请确保您完全信任所要运行的程序。运行未知或恶意程序可能对计算机安全造成威胁。\n设置“停止时间”将强制终止进程，可能导致数据未保存或文件损坏。", 
                  bootstyle="inverse-danger", wraplength=550, justify=LEFT).pack(fill=X)

        dialog_button_frame = ttk.Frame(dialog)
        dialog_button_frame.pack(pady=15)

        if task_to_edit:
            name_entry.insert(0, task_to_edit.get('name', ''))
            target_entry.insert(0, task_to_edit.get('target_path', ''))
            args_entry.insert(0, task_to_edit.get('arguments', ''))
            start_time_entry.insert(0, task_to_edit.get('time', ''))
            stop_time_entry.insert(0, task_to_edit.get('stop_time', ''))
            weekday_entry.insert(0, task_to_edit.get('weekday', '每周:1234567'))
            date_range_entry.insert(0, task_to_edit.get('date_range', '2025-01-01 ~ 2099-12-31'))
        else:
            weekday_entry.insert(0, "每周:1234567")
            date_range_entry.insert(0, "2025-01-01 ~ 2099-12-31")

        def save_task():
            target_path = target_entry.get().strip()
            if not target_path:
                messagebox.showerror("输入错误", "目标程序路径不能为空。", parent=dialog)
                return

            # --- ↓↓↓ 新增的输入验证模块 ↓↓↓ ---
            if not weekday_entry.get().strip():
                messagebox.showerror("输入错误", "“周几/几号”规则不能为空，请点击“选取...”进行设置。", parent=dialog)
                return
            
            if not date_range_entry.get().strip():
                messagebox.showerror("输入错误", "“日期范围”不能为空，请点击“设置...”进行配置。", parent=dialog)
                return
            # --- ↑↑↑ 验证模块结束 ↑↑↑ ---

            stop_time_str = stop_time_entry.get().strip()
            normalized_stop_time = ""
            if stop_time_str:
                normalized_stop_time = self._normalize_time_string(stop_time_str)
                if not normalized_stop_time:
                    messagebox.showerror("格式错误", "停止时间格式无效，应为 HH:MM:SS 或留空。", parent=dialog)
                    return

            is_valid_time, time_msg = self._normalize_multiple_times_string(start_time_entry.get().strip())
            if not is_valid_time: messagebox.showwarning("格式错误", time_msg, parent=dialog); return
            is_valid_date, date_msg = self._normalize_date_range_string(date_range_entry.get().strip())
            if not is_valid_date: messagebox.showwarning("格式错误", date_msg, parent=dialog); return

            new_task_data = {
                'name': name_entry.get().strip(), 'time': time_msg, 'type': 'execute',
                'stop_time': normalized_stop_time,
                'target_path': target_path, 'arguments': args_entry.get().strip(),
                'weekday': weekday_entry.get().strip(), 'date_range': date_msg,
                'status': '启用' if not task_to_edit else task_to_edit.get('status', '启用'),
                'last_run': {} if not task_to_edit else task_to_edit.get('last_run', {}),
            }
            if not new_task_data['name'] or not new_task_data['time']: 
                messagebox.showwarning("警告", "请填写任务名称和执行时间", parent=dialog); return

            if task_to_edit:
                self.execute_tasks[index] = new_task_data
                self.log(f"已修改运行任务: {new_task_data['name']}")
            else:
                self.execute_tasks.append(new_task_data)
                self.log(f"已添加运行任务: {new_task_data['name']}")

            self.update_execute_list()
            self.save_execute_tasks()
            cleanup_and_destroy()

        button_text = "保存修改" if task_to_edit else "添加"
        ttk.Button(dialog_button_frame, text=button_text, command=save_task, bootstyle="primary").pack(side=LEFT, padx=10, ipady=5)
        ttk.Button(dialog_button_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10, ipady=5)
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)
        
        self.center_window(dialog, parent=self.root)

    def load_print_tasks(self):
        if not os.path.exists(PRINT_TASK_FILE): return
        try:
            with open(PRINT_TASK_FILE, 'r', encoding='utf-8') as f:
                self.print_tasks = json.load(f)
            self.log(f"已加载 {len(self.print_tasks)} 个打印任务")
            if hasattr(self, 'print_tree'):
                self.update_print_list()
        except Exception as e:
            self.log(f"加载打印任务失败: {e}")
            self.print_tasks = []

    def save_print_tasks(self):
        try:
            with open(PRINT_TASK_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.print_tasks, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"保存打印任务失败: {e}")

    def update_print_list(self):
        if not hasattr(self, 'print_tree') or not self.print_tree.winfo_exists(): return
        self.print_tree.delete(*self.print_tree.get_children())
        for task in self.print_tasks:
            self.print_tree.insert('', END, values=(
                task.get('name', ''),
                task.get('status', '启用'),
                task.get('time', ''),
                os.path.basename(task.get('file_path', '')),
                task.get('printer_name', '默认打印机'),
                task.get('copies', 1),
                task.get('weekday', ''),
                task.get('date_range', '')
            ))

    def add_print_task(self):
        self.open_print_dialog()

    def edit_print_task(self):
        selection = self.print_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要修改的打印任务", parent=self.root)
            return
        index = self.print_tree.index(selection[0])
        task_to_edit = self.print_tasks[index]
        self.open_print_dialog(task_to_edit=task_to_edit, index=index)

    def delete_print_task(self):
        selections = self.print_tree.selection()
        if not selections:
            messagebox.showwarning("提示", "请先选择要删除的打印任务", parent=self.root)
            return
        if messagebox.askyesno("确认删除", f"确定要删除选中的 {len(selections)} 个打印任务吗？", parent=self.root):
            indices = sorted([self.print_tree.index(s) for s in selections], reverse=True)
            for index in indices:
                self.print_tasks.pop(index)
            self.update_print_list()
            self.save_print_tasks()

    def clear_all_print_tasks(self):
        if not self.print_tasks: return
        if messagebox.askyesno("确认清空", "您确定要清空所有打印任务吗？", parent=self.root):
            self.print_tasks.clear()
            self.update_print_list()
            self.save_print_tasks()

    def enable_all_print(self):
        if not self.print_tasks: return
        for task in self.print_tasks: task['status'] = '启用'
        self.update_print_list(); self.save_print_tasks()

    def disable_all_print(self):
        if not self.print_tasks: return
        for task in self.print_tasks: task['status'] = '禁用'
        self.update_print_list(); self.save_print_tasks()

    def get_printer_list(self):
        if not WIN32_AVAILABLE:
            return ["(功能受限，无法获取)"]
        try:
            printers = win32print.EnumPrinters(2)
            return [name for flags, desc, name, comment in printers]
        except Exception as e:
            self.log(f"获取打印机列表失败: {e}")
            return ["(获取失败)"]

    def open_print_dialog(self, task_to_edit=None, index=None):
        dialog = ttk.Toplevel(self.root)
        dialog.title("修改打印任务" if task_to_edit else "添加打印任务")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()

        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=BOTH, expand=True)

        content_frame = ttk.LabelFrame(main_frame, text="打印内容", padding=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=2)
        content_frame.columnconfigure(1, weight=1)
        
        ttk.Label(content_frame, text="任务名称:").grid(row=0, column=0, sticky='e', padx=5, pady=5)
        name_entry = ttk.Entry(content_frame, font=self.font_11)
        name_entry.grid(row=0, column=1, columnspan=2, sticky='ew', padx=5, pady=5)

        ttk.Label(content_frame, text="打印文件:").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        file_entry = ttk.Entry(content_frame, font=self.font_11)
        file_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=5)
        def select_file():
            path = filedialog.askopenfilename(title="选择要打印的文件", 
                                              filetypes=[("所有支持的文件", "*.pdf *.txt *.doc *.docx *.xls *.xlsx *.jpg *.png"), 
                                                         ("所有文件", "*.*")], 
                                              parent=dialog)
            if path:
                file_entry.delete(0, END)
                file_entry.insert(0, path)
        ttk.Button(content_frame, text="浏览...", command=select_file, bootstyle="outline").grid(row=1, column=2, padx=5)

        ttk.Label(content_frame, text="打印机:").grid(row=2, column=0, sticky='e', padx=5, pady=5)
        printer_var = tk.StringVar()
        printer_combo = ttk.Combobox(content_frame, textvariable=printer_var, values=self.get_printer_list(), font=self.font_11, state='readonly')
        printer_combo.grid(row=2, column=1, columnspan=2, sticky='ew', padx=5, pady=5)
        try:
            default_printer = win32print.GetDefaultPrinter()
            printer_var.set(default_printer)
        except Exception:
            if printer_combo['values']:
                printer_combo.current(0)
        
        ttk.Label(content_frame, text="打印份数:").grid(row=3, column=0, sticky='e', padx=5, pady=5)
        copies_entry = ttk.Entry(content_frame, font=self.font_11, width=10)
        copies_entry.grid(row=3, column=1, sticky='w', padx=5, pady=5)
        copies_entry.insert(0, "1")

        time_frame = ttk.LabelFrame(main_frame, text="时间规则", padding=15)
        time_frame.grid(row=1, column=0, sticky='ew', pady=4)
        time_frame.columnconfigure(1, weight=1)
        
        ttk.Label(time_frame, text="执行时间:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        start_time_entry = ttk.Entry(time_frame, font=self.font_11)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        self._bind_mousewheel_to_entry(start_time_entry, self._handle_time_scroll)
        ttk.Label(time_frame, text="<可多个>").grid(row=0, column=2, sticky='w', padx=5)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bootstyle="outline").grid(row=0, column=3, padx=5)
        
        ttk.Label(time_frame, text="周几/几号:").grid(row=1, column=0, sticky='e', padx=5, pady=3)
        weekday_entry = ttk.Entry(time_frame, font=self.font_11)
        weekday_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=3)
        ttk.Button(time_frame, text="选取...", command=lambda: self.show_weekday_settings_dialog(weekday_entry), bootstyle="outline").grid(row=1, column=3, padx=5)
        
        ttk.Label(time_frame, text="日期范围:").grid(row=2, column=0, sticky='e', padx=5, pady=3)
        date_range_entry = ttk.Entry(time_frame, font=self.font_11)
        date_range_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=3)
        self._bind_mousewheel_to_entry(date_range_entry, self._handle_date_scroll)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry), bootstyle="outline").grid(row=2, column=3, padx=5)

        dialog_button_frame = ttk.Frame(dialog)
        dialog_button_frame.pack(pady=15)

        if task_to_edit:
            name_entry.insert(0, task_to_edit.get('name', ''))
            file_entry.insert(0, task_to_edit.get('file_path', ''))
            printer_var.set(task_to_edit.get('printer_name', ''))
            copies_entry.delete(0, END)
            copies_entry.insert(0, task_to_edit.get('copies', 1))
            start_time_entry.insert(0, task_to_edit.get('time', ''))
            weekday_entry.insert(0, task_to_edit.get('weekday', '每周:1234567'))
            date_range_entry.insert(0, task_to_edit.get('date_range', '2025-01-01 ~ 2099-12-31'))
        else:
            weekday_entry.insert(0, "每周:1234567")
            date_range_entry.insert(0, "2025-01-01 ~ 2099-12-31")

        def save_task():
            file_path = file_entry.get().strip()
            if not file_path or not os.path.exists(file_path):
                messagebox.showerror("输入错误", "请选择一个有效的打印文件。", parent=dialog)
                return
            try:
                copies = int(copies_entry.get().strip())
                if copies < 1: raise ValueError
            except ValueError:
                messagebox.showerror("输入错误", "打印份数必须是大于0的整数。", parent=dialog)
                return

            is_valid_time, time_msg = self._normalize_multiple_times_string(start_time_entry.get().strip())
            if not is_valid_time: messagebox.showwarning("格式错误", time_msg, parent=dialog); return
            is_valid_date, date_msg = self._normalize_date_range_string(date_range_entry.get().strip())
            if not is_valid_date: messagebox.showwarning("格式错误", date_msg, parent=dialog); return

            new_task_data = {
                'name': name_entry.get().strip(),
                'file_path': file_path,
                'printer_name': printer_var.get(),
                'copies': copies,
                'time': time_msg,
                'weekday': weekday_entry.get().strip(),
                'date_range': date_msg,
                'status': '启用' if not task_to_edit else task_to_edit.get('status', '启用'),
                'last_run': {} if not task_to_edit else task_to_edit.get('last_run', {}),
            }
            if not new_task_data['name'] or not new_task_data['time']: 
                messagebox.showwarning("警告", "请填写任务名称和执行时间", parent=dialog); return

            if task_to_edit:
                self.print_tasks[index] = new_task_data
                self.log(f"已修改打印任务: {new_task_data['name']}")
            else:
                self.print_tasks.append(new_task_data)
                self.log(f"已添加打印任务: {new_task_data['name']}")

            self.update_print_list()
            self.save_print_tasks()
            cleanup_and_destroy()

        button_text = "保存修改" if task_to_edit else "添加"
        ttk.Button(dialog_button_frame, text=button_text, command=save_task, bootstyle="primary").pack(side=LEFT, padx=10, ipady=5)
        ttk.Button(dialog_button_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10, ipady=5)
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)
        
        self.center_window(dialog, parent=self.root)

    def show_print_context_menu(self, event):
        if self.is_locked: return
        iid = self.print_tree.identify_row(event.y)
        context_menu = tk.Menu(self.root, tearoff=0, font=self.font_11)

        if iid:
            if iid not in self.print_tree.selection():
                self.print_tree.selection_set(iid)

            context_menu.add_command(label="修改", command=self.edit_print_task)
            context_menu.add_command(label="删除", command=self.delete_print_task)
            context_menu.add_separator()
            context_menu.add_command(label="置顶", command=self.move_print_to_top)
            context_menu.add_command(label="上移", command=lambda: self.move_print_task(-1))
            context_menu.add_command(label="下移", command=lambda: self.move_print_task(1))
            context_menu.add_command(label="置末", command=self.move_print_to_bottom)
            context_menu.add_separator()
            context_menu.add_command(label="启用", command=lambda: self._set_print_status('启用'))
            context_menu.add_command(label="禁用", command=lambda: self._set_print_status('禁用'))
        else:
            self.print_tree.selection_set()
            context_menu.add_command(label="添加任务", command=self.add_print_task)

        context_menu.post(event.x_root, event.y_root)

    def _set_print_status(self, status):
        selection = self.print_tree.selection()
        if not selection:
            messagebox.showwarning("提示", f"请先选择要 {status} 的任务", parent=self.root)
            return
        for item_id in selection:
            index = self.print_tree.index(item_id)
            self.print_tasks[index]['status'] = status
        self.update_print_list()
        self.save_print_tasks()

    def move_print_task(self, direction):
        selection = self.print_tree.selection()
        if not selection or len(selection) > 1: return
        index = self.print_tree.index(selection[0])
        new_index = index + direction
        if 0 <= new_index < len(self.print_tasks):
            task_to_move = self.print_tasks.pop(index)
            self.print_tasks.insert(new_index, task_to_move)
            self.update_print_list()
            self.save_print_tasks()
            items = self.print_tree.get_children()
            if items: self.print_tree.selection_set(items[new_index]); self.print_tree.focus(items[new_index])

    def move_print_to_top(self):
        selection = self.print_tree.selection()
        if not selection or len(selection) > 1: return
        index = self.print_tree.index(selection[0])
        if index > 0:
            task_to_move = self.print_tasks.pop(index)
            self.print_tasks.insert(0, task_to_move)
            self.update_print_list()
            self.save_print_tasks()
            items = self.print_tree.get_children()
            if items: self.print_tree.selection_set(items[0]); self.print_tree.focus(items[0])

    def move_print_to_bottom(self):
        selection = self.print_tree.selection()
        if not selection or len(selection) > 1: return
        index = self.print_tree.index(selection[0])
        if index < len(self.print_tasks) - 1:
            task_to_move = self.print_tasks.pop(index)
            self.print_tasks.append(task_to_move)
            self.update_print_list()
            self.save_print_tasks()
            items = self.print_tree.get_children()
            if items: self.print_tree.selection_set(items[-1]); self.print_tree.focus(items[-1])

    def load_backup_tasks(self):
        if not os.path.exists(BACKUP_TASK_FILE): return
        try:
            with open(BACKUP_TASK_FILE, 'r', encoding='utf-8') as f:
                self.backup_tasks = json.load(f)
            self.log(f"已加载 {len(self.backup_tasks)} 个备份任务")
            if hasattr(self, 'backup_tree'):
                self.update_backup_list()
        except Exception as e:
            self.log(f"加载备份任务失败: {e}")
            self.backup_tasks = []

    def save_backup_tasks(self):
        try:
            with open(BACKUP_TASK_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.backup_tasks, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"保存备份任务失败: {e}")

    def update_backup_list(self):
        if not hasattr(self, 'backup_tree') or not self.backup_tree.winfo_exists(): return
        self.backup_tree.delete(*self.backup_tree.get_children())
        for task in self.backup_tasks:
            mode_value = task.get('backup_mode', 'mirror')
            mode_display_text = "镜像" if mode_value == 'mirror' else "增量"
            self.backup_tree.insert('', END, values=(
                task.get('name', ''),
                task.get('status', '启用'),
                task.get('time', ''),
                task.get('source_folder', ''),
                task.get('target_folder', ''),
                mode_display_text,
                task.get('weekday', ''),
                task.get('date_range', '')
            ))

    def add_backup_task(self):
        self.open_backup_dialog()

    def edit_backup_task(self):
        selection = self.backup_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要修改的备份任务", parent=self.root)
            return
        index = self.backup_tree.index(selection[0])
        task_to_edit = self.backup_tasks[index]
        self.open_backup_dialog(task_to_edit=task_to_edit, index=index)

    def delete_backup_task(self):
        selections = self.backup_tree.selection()
        if not selections:
            messagebox.showwarning("提示", "请先选择要删除的备份任务", parent=self.root)
            return
        if messagebox.askyesno("确认删除", f"确定要删除选中的 {len(selections)} 个备份任务吗？", parent=self.root):
            indices = sorted([self.backup_tree.index(s) for s in selections], reverse=True)
            for index in indices:
                self.backup_tasks.pop(index)
            self.update_backup_list()
            self.save_backup_tasks()

    def clear_all_backup_tasks(self):
        if not self.backup_tasks: return
        if messagebox.askyesno("确认清空", "您确定要清空所有备份任务吗？", parent=self.root):
            self.backup_tasks.clear()
            self.update_backup_list()
            self.save_backup_tasks()

    def enable_all_backup(self):
        if not self.backup_tasks: return
        for task in self.backup_tasks: task['status'] = '启用'
        self.update_backup_list(); self.save_backup_tasks()

    def disable_all_backup(self):
        if not self.backup_tasks: return
        for task in self.backup_tasks: task['status'] = '禁用'
        self.update_backup_list(); self.save_backup_tasks()
    
    def open_backup_dialog(self, task_to_edit=None, index=None):
        dialog = ttk.Toplevel(self.root)
        dialog.title("修改备份任务" if task_to_edit else "添加备份任务")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()

        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=BOTH, expand=True)

        content_frame = ttk.LabelFrame(main_frame, text="备份设置", padding=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=2)
        content_frame.columnconfigure(1, weight=1)
        
        ttk.Label(content_frame, text="任务名称:").grid(row=0, column=0, sticky='e', padx=5, pady=5)
        name_entry = ttk.Entry(content_frame, font=self.font_11)
        name_entry.grid(row=0, column=1, columnspan=2, sticky='ew', padx=5, pady=5)

        def select_folder_for_entry(entry_widget):
            folder = filedialog.askdirectory(title="请选择文件夹", parent=dialog)
            if folder:
                entry_widget.delete(0, END)
                entry_widget.insert(0, folder)

        ttk.Label(content_frame, text="源文件夹:").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        source_entry = ttk.Entry(content_frame, font=self.font_11)
        source_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=5)
        ttk.Button(content_frame, text="浏览...", command=lambda: select_folder_for_entry(source_entry), bootstyle="outline").grid(row=1, column=2, padx=5)

        ttk.Label(content_frame, text="目标文件夹:").grid(row=2, column=0, sticky='e', padx=5, pady=5)
        target_entry = ttk.Entry(content_frame, font=self.font_11)
        target_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=5)
        ttk.Button(content_frame, text="浏览...", command=lambda: select_folder_for_entry(target_entry), bootstyle="outline").grid(row=2, column=2, padx=5)

        ttk.Label(content_frame, text="备份模式:").grid(row=3, column=0, sticky='e', padx=5, pady=5)
        mode_var = tk.StringVar(value="mirror")
        mode_frame = ttk.Frame(content_frame)
        mode_frame.grid(row=3, column=1, columnspan=2, sticky='w', padx=5)
        ttk.Radiobutton(mode_frame, text="镜像 (完全同步，会删除多余文件)", variable=mode_var, value="mirror").pack(anchor='w')
        ttk.Radiobutton(mode_frame, text="增量 (只复制新增/修改，不删除)", variable=mode_var, value="incremental").pack(anchor='w')

        time_frame = ttk.LabelFrame(main_frame, text="时间规则", padding=15)
        time_frame.grid(row=1, column=0, sticky='ew', pady=4)
        time_frame.columnconfigure(1, weight=1)
        
        ttk.Label(time_frame, text="执行时间:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        start_time_entry = ttk.Entry(time_frame, font=self.font_11)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        self._bind_mousewheel_to_entry(start_time_entry, self._handle_time_scroll)
        ttk.Label(time_frame, text="<可多个>").grid(row=0, column=2, sticky='w', padx=5)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bootstyle="outline").grid(row=0, column=3, padx=5)
        
        ttk.Label(time_frame, text="周几/几号:").grid(row=1, column=0, sticky='e', padx=5, pady=3)
        weekday_entry = ttk.Entry(time_frame, font=self.font_11)
        weekday_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=3)
        ttk.Button(time_frame, text="选取...", command=lambda: self.show_weekday_settings_dialog(weekday_entry), bootstyle="outline").grid(row=1, column=3, padx=5)
        
        ttk.Label(time_frame, text="日期范围:").grid(row=2, column=0, sticky='e', padx=5, pady=3)
        date_range_entry = ttk.Entry(time_frame, font=self.font_11)
        date_range_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=3)
        self._bind_mousewheel_to_entry(date_range_entry, self._handle_date_scroll)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry), bootstyle="outline").grid(row=2, column=3, padx=5)

        dialog_button_frame = ttk.Frame(dialog)
        dialog_button_frame.pack(pady=15)

        if task_to_edit:
            name_entry.insert(0, task_to_edit.get('name', ''))
            source_entry.insert(0, task_to_edit.get('source_folder', ''))
            target_entry.insert(0, task_to_edit.get('target_folder', ''))
            mode_var.set(task_to_edit.get('backup_mode', 'mirror'))
            start_time_entry.insert(0, task_to_edit.get('time', ''))
            weekday_entry.insert(0, task_to_edit.get('weekday', '每周:1234567'))
            date_range_entry.insert(0, task_to_edit.get('date_range', '2025-01-01 ~ 2099-12-31'))
        else:
            weekday_entry.insert(0, "每周:1234567")
            date_range_entry.insert(0, "2025-01-01 ~ 2099-12-31")

        def save_task():
            source_folder = source_entry.get().strip()
            target_folder = target_entry.get().strip()
            if not source_folder or not os.path.isdir(source_folder):
                messagebox.showerror("输入错误", "请选择一个有效的源文件夹。", parent=dialog)
                return
            if not target_folder:
                messagebox.showerror("输入错误", "目标文件夹不能为空。", parent=dialog)
                return
            if source_folder == target_folder:
                messagebox.showerror("输入错误", "源文件夹和目标文件夹不能相同。", parent=dialog)
                return

            is_valid_time, time_msg = self._normalize_multiple_times_string(start_time_entry.get().strip())
            if not is_valid_time: messagebox.showwarning("格式错误", time_msg, parent=dialog); return
            is_valid_date, date_msg = self._normalize_date_range_string(date_range_entry.get().strip())
            if not is_valid_date: messagebox.showwarning("格式错误", date_msg, parent=dialog); return

            new_task_data = {
                'name': name_entry.get().strip(),
                'source_folder': source_folder,
                'target_folder': target_folder,
                'backup_mode': mode_var.get(),
                'time': time_msg,
                'weekday': weekday_entry.get().strip(),
                'date_range': date_msg,
                'status': '启用' if not task_to_edit else task_to_edit.get('status', '启用'),
                'last_run': {} if not task_to_edit else task_to_edit.get('last_run', {}),
            }
            if not new_task_data['name'] or not new_task_data['time']: 
                messagebox.showwarning("警告", "请填写任务名称和执行时间", parent=dialog); return

            if task_to_edit:
                self.backup_tasks[index] = new_task_data
                self.log(f"已修改备份任务: {new_task_data['name']}")
            else:
                self.backup_tasks.append(new_task_data)
                self.log(f"已添加备份任务: {new_task_data['name']}")

            self.update_backup_list()
            self.save_backup_tasks()
            cleanup_and_destroy()

        button_text = "保存修改" if task_to_edit else "添加"
        ttk.Button(dialog_button_frame, text=button_text, command=save_task, bootstyle="primary").pack(side=LEFT, padx=10, ipady=5)
        ttk.Button(dialog_button_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10, ipady=5)
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)
        
        self.center_window(dialog, parent=self.root)

    def show_backup_context_menu(self, event):
        if self.is_locked: return
        iid = self.backup_tree.identify_row(event.y)
        context_menu = tk.Menu(self.root, tearoff=0, font=self.font_11)

        if iid:
            if iid not in self.backup_tree.selection():
                self.backup_tree.selection_set(iid)

            context_menu.add_command(label="修改", command=self.edit_backup_task)
            context_menu.add_command(label="删除", command=self.delete_backup_task)
            context_menu.add_separator()
            context_menu.add_command(label="置顶", command=self.move_backup_to_top)
            context_menu.add_command(label="上移", command=lambda: self.move_backup_task(-1))
            context_menu.add_command(label="下移", command=lambda: self.move_backup_task(1))
            context_menu.add_command(label="置末", command=self.move_backup_to_bottom)
            context_menu.add_separator()
            context_menu.add_command(label="启用", command=lambda: self._set_backup_status('启用'))
            context_menu.add_command(label="禁用", command=lambda: self._set_backup_status('禁用'))
        else:
            self.backup_tree.selection_set()
            context_menu.add_command(label="添加任务", command=self.add_backup_task)

        context_menu.post(event.x_root, event.y_root)

    def _set_backup_status(self, status):
        selection = self.backup_tree.selection()
        if not selection:
            messagebox.showwarning("提示", f"请先选择要 {status} 的任务", parent=self.root)
            return
        for item_id in selection:
            index = self.backup_tree.index(item_id)
            self.backup_tasks[index]['status'] = status
        self.update_backup_list()
        self.save_backup_tasks()

    def move_backup_task(self, direction):
        selection = self.backup_tree.selection()
        if not selection or len(selection) > 1: return
        index = self.backup_tree.index(selection[0])
        new_index = index + direction
        if 0 <= new_index < len(self.backup_tasks):
            task_to_move = self.backup_tasks.pop(index)
            self.backup_tasks.insert(new_index, task_to_move)
            self.update_backup_list()
            self.save_backup_tasks()
            items = self.backup_tree.get_children()
            if items: self.backup_tree.selection_set(items[new_index]); self.backup_tree.focus(items[new_index])

    def move_backup_to_top(self):
        selection = self.backup_tree.selection()
        if not selection or len(selection) > 1: return
        index = self.backup_tree.index(selection[0])
        if index > 0:
            task_to_move = self.backup_tasks.pop(index)
            self.backup_tasks.insert(0, task_to_move)
            self.update_backup_list()
            self.save_backup_tasks()
            items = self.backup_tree.get_children()
            if items: self.backup_tree.selection_set(items[0]); self.backup_tree.focus(items[0])

    def move_backup_to_bottom(self):
        selection = self.backup_tree.selection()
        if not selection or len(selection) > 1: return
        index = self.backup_tree.index(selection[0])
        if index < len(self.backup_tasks) - 1:
            task_to_move = self.backup_tasks.pop(index)
            self.backup_tasks.append(task_to_move)
            self.update_backup_list()
            self.save_backup_tasks()
            items = self.backup_tree.get_children()
            if items: self.backup_tree.selection_set(items[-1]); self.backup_tree.focus(items[-1])

    # --- ↓↓↓ [新增] 媒体处理功能模块 (FFmpeg) - V3 最终整合版 ↓↓↓ ---

    def _build_media_processing_ui(self, parent_frame):
        # 检查ffmpeg是否存在
        ffmpeg_path = os.path.join(application_path, "ffmpeg.exe")
        if not os.path.exists(ffmpeg_path):
            warning_label = ttk.Label(parent_frame,
                                      text="错误：媒体处理功能依赖于 FFmpeg。\n\n请下载 FFmpeg，并将其中的 ffmpeg.exe 文件放置到本软件所在的文件夹内，然后重启软件。",
                                      font=self.font_12_bold, bootstyle="danger", justify="center")
            warning_label.pack(pady=50, fill=X, expand=True)
            return

        # 使用可滚动框架，防止窗口过小时内容溢出
        scrolled_frame = ScrolledFrame(parent_frame, autohide=True)
        scrolled_frame.pack(fill=BOTH, expand=True)
        container = scrolled_frame.container # 在这个 container 内部构建UI

        # 顶部说明文字
        desc_text = "此功能依赖于软件根目录下的 ffmpeg.exe，用于即时处理音视频文件。注意：同一时间只能执行一个媒体处理任务。"
        ttk.Label(container, text=desc_text, bootstyle="info").pack(fill=X, pady=(0, 15))

        # --- 功能1: 提取音频 ---
        extract_lf = ttk.LabelFrame(container, text=" 1. 从视频中提取音频 ", padding=15)
        extract_lf.pack(fill=X, pady=10)
        extract_lf.columnconfigure(1, weight=1)

        self.extract_input_var = tk.StringVar()
        self.extract_output_var = tk.StringVar()
        
        ttk.Label(extract_lf, text="源视频文件:").grid(row=0, column=0, sticky='e', padx=5, pady=5)
        ttk.Entry(extract_lf, textvariable=self.extract_input_var).grid(row=0, column=1, sticky='ew')
        ttk.Button(extract_lf, text="浏览...", bootstyle="outline", command=lambda: self._select_media_file(self.extract_input_var, "选择视频文件")).grid(row=0, column=2, padx=5)

        ttk.Label(extract_lf, text="输出音频文件:").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        ttk.Entry(extract_lf, textvariable=self.extract_output_var).grid(row=1, column=1, sticky='ew')
        ttk.Button(extract_lf, text="浏览...", bootstyle="outline", command=self._select_extract_output_file).grid(row=1, column=2, padx=5)
        
        extract_action_frame = ttk.Frame(extract_lf)
        extract_action_frame.grid(row=2, column=1, sticky='ew', pady=(10,0))
        extract_action_frame.columnconfigure(1, weight=1)
        self.extract_start_btn = ttk.Button(extract_action_frame, text="开始提取", bootstyle="success", width=12, command=self._start_extraction)
        self.extract_start_btn.grid(row=0, column=0, ipady=4)
        self.extract_progress = ttk.Progressbar(extract_action_frame, mode='determinate')
        self.extract_progress.grid(row=0, column=1, sticky='ew', padx=10)
        self.extract_status_label = ttk.Label(extract_action_frame, text="准备就绪", bootstyle="secondary")
        self.extract_status_label.grid(row=0, column=2)

        # --- 功能2: 转换视频格式 ---
        convert_lf = ttk.LabelFrame(container, text=" 2. 转换视频格式为通用MP4 ", padding=15)
        convert_lf.pack(fill=X, pady=10)
        convert_lf.columnconfigure(1, weight=1)

        self.convert_input_var = tk.StringVar()
        self.convert_output_var = tk.StringVar()

        ttk.Label(convert_lf, text="源视频文件:").grid(row=0, column=0, sticky='e', padx=5, pady=5)
        ttk.Entry(convert_lf, textvariable=self.convert_input_var).grid(row=0, column=1, sticky='ew')
        ttk.Button(convert_lf, text="浏览...", bootstyle="outline", command=lambda: self._select_media_file(self.convert_input_var, "选择视频文件")).grid(row=0, column=2, padx=5)

        ttk.Label(convert_lf, text="输出视频文件:").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        ttk.Entry(convert_lf, textvariable=self.convert_output_var).grid(row=1, column=1, sticky='ew')
        ttk.Button(convert_lf, text="浏览...", bootstyle="outline", command=self._select_convert_output_file).grid(row=1, column=2, padx=5)

        convert_action_frame = ttk.Frame(convert_lf)
        convert_action_frame.grid(row=2, column=1, sticky='ew', pady=(10,0))
        convert_action_frame.columnconfigure(1, weight=1)
        self.convert_start_btn = ttk.Button(convert_action_frame, text="开始转换", bootstyle="success", width=12, command=self._start_conversion)
        self.convert_start_btn.grid(row=0, column=0, ipady=4)
        self.convert_progress = ttk.Progressbar(convert_action_frame, mode='determinate')
        self.convert_progress.grid(row=0, column=1, sticky='ew', padx=10)
        self.convert_status_label = ttk.Label(convert_action_frame, text="准备就绪", bootstyle="secondary")
        self.convert_status_label.grid(row=0, column=2)
        
        # --- 功能3: 剪辑片段 ---
        trim_lf = ttk.LabelFrame(container, text=" 3. 剪辑音视频片段 ", padding=15)
        trim_lf.pack(fill=X, pady=10)
        trim_lf.columnconfigure(1, weight=1)
        
        self.trim_input_var = tk.StringVar()
        self.trim_output_var = tk.StringVar()
        self.trim_start_time_var = tk.StringVar(value="00:00:00")
        self.trim_end_time_var = tk.StringVar()

        ttk.Label(trim_lf, text="源文件:").grid(row=0, column=0, sticky='e', padx=5, pady=5)
        ttk.Entry(trim_lf, textvariable=self.trim_input_var).grid(row=0, column=1, sticky='ew')
        ttk.Button(trim_lf, text="浏览...", bootstyle="outline", command=lambda: self._select_media_file(self.trim_input_var, "选择音视频文件")).grid(row=0, column=2, padx=5)

        ttk.Label(trim_lf, text="输出文件:").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        ttk.Entry(trim_lf, textvariable=self.trim_output_var).grid(row=1, column=1, sticky='ew')
        ttk.Button(trim_lf, text="浏览...", bootstyle="outline", command=self._select_trim_output_file).grid(row=1, column=2, padx=5)

        time_frame = ttk.Frame(trim_lf)
        time_frame.grid(row=2, column=1, sticky='w', pady=5)
        ttk.Label(time_frame, text="开始时间:").pack(side=LEFT)
        ttk.Entry(time_frame, textvariable=self.trim_start_time_var, width=12).pack(side=LEFT, padx=5)
        ttk.Label(time_frame, text="结束时间:").pack(side=LEFT, padx=(10,0))
        ttk.Entry(time_frame, textvariable=self.trim_end_time_var, width=12).pack(side=LEFT, padx=5)
        ttk.Label(time_frame, text="(格式: HH:MM:SS 或 秒)", bootstyle="secondary").pack(side=LEFT)
        
        trim_action_frame = ttk.Frame(trim_lf)
        trim_action_frame.grid(row=3, column=1, sticky='ew', pady=(10,0))
        trim_action_frame.columnconfigure(1, weight=1)
        self.trim_start_btn = ttk.Button(trim_action_frame, text="开始剪辑", bootstyle="success", width=12, command=self._start_trimming)
        self.trim_start_btn.grid(row=0, column=0, ipady=4)
        self.trim_progress = ttk.Progressbar(trim_action_frame, mode='determinate')
        self.trim_progress.grid(row=0, column=1, sticky='ew', padx=10)
        self.trim_status_label = ttk.Label(trim_action_frame, text="准备就绪", bootstyle="secondary")
        self.trim_status_label.grid(row=0, column=2)

    def _select_media_file(self, string_var, title):
        filetypes = [("媒体文件", "*.mp4 *.mkv *.avi *.mov *.mp3 *.wav *.flac *.ts"), ("所有文件", "*.*")]
        filename = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if filename:
            string_var.set(filename)

    def _select_extract_output_file(self):
        input_file = self.extract_input_var.get()
        if not input_file:
            messagebox.showwarning("提示", "请先选择一个源视频文件。")
            return
        
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        def_name = f"{base_name}_audio.mp3"
        
        filetypes = [("MP3 Audio", "*.mp3"), ("WAV Audio", "*.wav"), ("AAC Audio", "*.aac"), ("FLAC Audio", "*.flac")]
        filename = filedialog.asksaveasfilename(title="保存提取的音频", initialfile=def_name, filetypes=filetypes, defaultextension=".mp3")
        if filename:
            self.extract_output_var.set(filename)

    def _select_convert_output_file(self):
        input_file = self.convert_input_var.get()
        if not input_file:
            messagebox.showwarning("提示", "请先选择一个源视频文件。")
            return
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        def_name = f"{base_name}_converted.mp4"
        filetypes = [("MP4 Video", "*.mp4")]
        filename = filedialog.asksaveasfilename(title="保存转换后的视频", initialfile=def_name, filetypes=filetypes, defaultextension=".mp4")
        if filename:
            self.convert_output_var.set(filename)

    def _select_trim_output_file(self):
        input_file = self.trim_input_var.get()
        if not input_file:
            messagebox.showwarning("提示", "请先选择一个源文件。")
            return
        base_name, ext = os.path.splitext(os.path.basename(input_file))
        def_name = f"{base_name}_trimmed{ext}"
        filetypes = [(f"{ext.upper()} File", f"*{ext}"), ("All Files", "*.*")]
        filename = filedialog.asksaveasfilename(title="保存剪辑后的文件", initialfile=def_name, filetypes=filetypes, defaultextension=ext)
        if filename:
            self.trim_output_var.set(filename)

    def _toggle_media_buttons(self, state):
        """统一控制所有媒体处理按钮的状态"""
        self.extract_start_btn.config(state=state)
        self.convert_start_btn.config(state=state)
        self.trim_start_btn.config(state=state)

    def _start_extraction(self):
        input_file = self.extract_input_var.get()
        output_file = self.extract_output_var.get()
        
        if not all([input_file, output_file]):
            messagebox.showerror("错误", "输入和输出文件路径都不能为空。")
            return
            
        ffmpeg_exe = os.path.join(application_path, "ffmpeg.exe")
        ext = os.path.splitext(output_file)[1].lower()
        
        command = [ffmpeg_exe, "-hide_banner", "-i", input_file, "-vn"] # -vn = No Video
        
        codec_map = {'.mp3': 'libmp3lame', '.aac': 'aac', '.wav': 'pcm_s16le', '.flac': 'flac'}
        if ext in codec_map:
            command.extend(["-c:a", codec_map[ext]])
        else:
            messagebox.showerror("错误", f"不支持的输出音频格式: {ext}")
            return
        
        command.extend(["-y", output_file])

        self._toggle_media_buttons(DISABLED)
        self.extract_progress['value'] = 0
        self.extract_status_label.config(text="正在处理...")
        
        threading.Thread(
            target=self._media_processing_worker,
            args=(command, input_file, self.extract_progress, self.extract_status_label, "提取音频"),
            daemon=True
        ).start()

    def _start_conversion(self):
        input_file = self.convert_input_var.get()
        output_file = self.convert_output_var.get()

        if not all([input_file, output_file]):
            messagebox.showerror("错误", "输入和输出文件路径都不能为空。")
            return

        ffmpeg_exe = os.path.join(application_path, "ffmpeg.exe")
        command = [
            ffmpeg_exe, "-hide_banner", "-i", input_file,
            "-c:v", "libx264",      # 使用通用性最好的 H.264 编码
            "-preset", "fast",     # 在速度和质量之间取得良好平衡
            "-pix_fmt", "yuv420p", # 确保最大的播放器兼容性
            "-c:a", "aac",          # 使用通用的 AAC 音频编码
            "-b:a", "192k",         # 合理的音频码率
            "-y", output_file
        ]
        
        self._toggle_media_buttons(DISABLED)
        self.convert_progress['value'] = 0
        self.convert_status_label.config(text="正在处理...")

        threading.Thread(
            target=self._media_processing_worker,
            args=(command, input_file, self.convert_progress, self.convert_status_label, "转换视频"),
            daemon=True
        ).start()

    def _start_trimming(self):
        input_file = self.trim_input_var.get()
        output_file = self.trim_output_var.get()
        start_time = self.trim_start_time_var.get()
        end_time = self.trim_end_time_var.get()

        if not all([input_file, output_file, start_time]):
            messagebox.showerror("错误", "输入、输出和开始时间都不能为空。")
            return
        
        ffmpeg_exe = os.path.join(application_path, "ffmpeg.exe")
        
        command = [
            ffmpeg_exe, 
            "-hide_banner", 
            "-i", input_file, 
            "-ss", start_time
        ]

        if end_time:
            command.extend(["-to", end_time])
        
        # 不再强制指定编码器，让ffmpeg自动选择，并强制重新编码以确保精度
        command.extend(["-y", output_file])

        self._toggle_media_buttons(DISABLED)
        self.trim_progress['value'] = 0
        self.trim_status_label.config(text="正在处理...")

        threading.Thread(
            target=self._media_processing_worker,
            args=(command, input_file, self.trim_progress, self.trim_status_label, "剪辑片段"),
            daemon=True
        ).start()
    
    def _parse_time_to_seconds(self, time_str):
        """将 HH:MM:SS 或纯秒数的字符串安全地转换为总秒数"""
        if not time_str:
            return None
        try:
            if ':' in time_str:
                parts = time_str.split(':')
                seconds = 0
                # 从秒开始反向计算，支持 HH:MM:SS, MM:SS, SS 等格式
                for i, part in enumerate(reversed(parts)):
                    seconds += float(part) * (60**i)
                return seconds
            else:
                return float(time_str)
        except (ValueError, TypeError):
            return None

    def _media_processing_worker(self, command, input_file, progress_widget, status_widget, operation_name):
        """通用媒体处理后台工作线程 (V3 - 修复剪辑和死锁问题)"""
        
        def update_ui(key, value):
            if key == 'progress':
                progress_widget.config(value=value)
            elif key == 'status':
                status_widget.config(text=value)

        total_duration_sec = 0.0
        try:
            # 1. 获取总时长
            ffprobe_exe = os.path.join(application_path, "ffmpeg.exe").replace("ffmpeg", "ffprobe")
            ffprobe_cmd = [ffprobe_exe, "-v", "error", "-show_entries", "format=duration", "-of", "default=noprint_wrappers=1:nokey=1", input_file]
            duration_proc = subprocess.run(ffprobe_cmd, capture_output=True, text=True)
            if duration_proc.returncode == 0 and duration_proc.stdout.strip():
                total_duration_sec = float(duration_proc.stdout.strip())
            
            # 修正剪辑操作的进度条总时长
            if operation_name == "剪辑片段":
                start_sec, end_sec = None, None
                
                try:
                    start_sec_str = command[command.index("-ss") + 1]
                    start_sec = self._parse_time_to_seconds(start_sec_str)
                except (ValueError, IndexError): pass
                
                try:
                    if "-to" in command:
                        end_sec_str = command[command.index("-to") + 1]
                        end_sec = self._parse_time_to_seconds(end_sec_str)
                except (ValueError, IndexError): pass

                if start_sec is not None and end_sec is not None and end_sec > start_sec:
                    total_duration_sec = end_sec - start_sec
                elif start_sec is not None and total_duration_sec > 0:
                    total_duration_sec = max(0, total_duration_sec - start_sec)

            # 2. 执行主命令并解析进度
            progress_command = command[:2] + ["-progress", "pipe:1"] + command[2:]
            
            process = subprocess.Popen(progress_command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding='utf-8', errors='replace')

            stderr_output = []
            def log_stderr(pipe):
                for line in iter(pipe.readline, ''):
                    stderr_output.append(line)
                pipe.close()

            stderr_thread = threading.Thread(target=log_stderr, args=(process.stderr,))
            stderr_thread.daemon = True
            stderr_thread.start()

            while True:
                line = process.stdout.readline()
                if not line and process.poll() is not None:
                    break
                if total_duration_sec > 0 and 'out_time_us' in line:
                    parts = line.strip().split('=')
                    if len(parts) == 2:
                        try:
                            current_us = int(parts[1])
                            progress = min(100, (current_us / (total_duration_sec * 1_000_000)) * 100)
                            self.root.after(0, update_ui, 'progress', progress)
                        except (ValueError, ZeroDivisionError):
                            pass
            
            process.wait()
            stderr_thread.join()
            
            if process.returncode != 0:
                full_stderr = "".join(stderr_output)
                raise Exception(f"FFmpeg 返回错误 (代码: {process.returncode})\n\n{full_stderr[-1000:]}")

            self.root.after(0, update_ui, 'progress', 100)
            self.root.after(0, update_ui, 'status', "处理成功!")
        except Exception as e:
            self.root.after(0, update_ui, 'progress', 0)
            self.root.after(0, update_ui, 'status', "失败!")
            messagebox.showerror("处理失败", f"执行“{operation_name}”操作时发生错误:\n\n{e}")
        finally:
            self.root.after(100, self._toggle_media_buttons, 'normal')

    #↑ --- [新增] 媒体处理功能模块结束 ---
    #↓以下是全套更换壁纸的功能

    def _build_wallpaper_ui(self, parent_frame):
        scrolled_frame = ScrolledFrame(parent_frame, autohide=True)
        scrolled_frame.pack(fill=BOTH, expand=True)
        container = scrolled_frame.container

        # --- 描述区 ---
        title_label = ttk.Label(container, text="网络壁纸自动更换", font=self.font_14_bold, bootstyle="primary")
        title_label.pack(anchor="w", pady=(0, 5))
        
        desc_text = "此功能会自动从网络获取高质量壁纸（必应每日壁纸），并定时为您更换桌面。\n下载的壁纸将保存在软件根目录下的“每日壁纸”文件夹内。"
        desc_label = ttk.Label(container, text=desc_text, bootstyle="secondary")
        desc_label.pack(anchor="w", pady=(0, 15), fill=X)
        
        ttk.Separator(container, orient=HORIZONTAL).pack(fill=X, pady=5)

        # --- 总开关 ---
        enable_check = ttk.Checkbutton(container, text="启用网络壁纸自动更换功能", variable=self.wallpaper_enabled_var, bootstyle="round-toggle")
        enable_check.pack(anchor="w", pady=10)

        # --- 更换规则 ---
        rule_lf = ttk.LabelFrame(container, text="更换规则", padding=15)
        rule_lf.pack(fill=X, pady=5)
        
        rule_frame = ttk.Frame(rule_lf)
        rule_frame.pack(fill=X)
        ttk.Label(rule_frame, text="每隔").pack(side=LEFT, padx=(0, 5))
        interval_entry = ttk.Entry(rule_frame, textvariable=self.wallpaper_interval_days_var, width=5, font=self.font_11)
        interval_entry.pack(side=LEFT)
        ttk.Label(rule_frame, text="天，在").pack(side=LEFT, padx=5)
        time_entry = ttk.Entry(rule_frame, textvariable=self.wallpaper_change_time_var, width=12, font=self.font_11)
        time_entry.pack(side=LEFT)
        self._bind_mousewheel_to_entry(time_entry, self._handle_time_scroll) # 复用时间滚动功能
        ttk.Label(rule_frame, text="时自动更换壁纸。").pack(side=LEFT, padx=5)

        # --- 缓存管理 ---
        cache_lf = ttk.LabelFrame(container, text="缓存管理", padding=15)
        cache_lf.pack(fill=X, pady=5)
        
        cache_frame = ttk.Frame(cache_lf)
        cache_frame.pack(fill=X)
        ttk.Label(cache_frame, text="自动清理").pack(side=LEFT, padx=(0, 5))
        cache_entry = ttk.Entry(cache_frame, textvariable=self.wallpaper_cache_days_var, width=5, font=self.font_11)
        cache_entry.pack(side=LEFT)
        ttk.Label(cache_frame, text="天前的壁纸缓存文件。").pack(side=LEFT, padx=5)

        # --- 手动操作 ---
        manual_lf = ttk.LabelFrame(container, text="手动操作", padding=15)
        manual_lf.pack(fill=X, pady=5)
        
        manual_frame = ttk.Frame(manual_lf)
        manual_frame.pack(fill=X)
        
        ttk.Button(manual_frame, text="立即获取并更换", command=self._trigger_wallpaper_change_now, bootstyle="info").pack(side=LEFT, padx=5, ipady=4)
        ttk.Button(manual_frame, text="打开壁纸文件夹", command=self._open_wallpaper_folder, bootstyle="secondary-outline").pack(side=LEFT, padx=5, ipady=4)
        ttk.Button(manual_frame, text="清理所有壁纸缓存", command=self._clear_wallpaper_cache, bootstyle="danger-outline").pack(side=LEFT, padx=5, ipady=4)

        # --- 保存按钮 ---
        save_btn = ttk.Button(container, text="保存设置", command=self._save_wallpaper_settings, bootstyle="success")
        save_btn.pack(pady=20, ipady=5)

    # --- ↓↓↓ 新增代码：所有网络壁纸功能的后台逻辑方法 ↓↓↓ ---
    
    def _save_wallpaper_settings(self):
        """保存网络壁纸页面的所有设置到 settings.json"""
        try:
            # 输入验证
            interval = int(self.wallpaper_interval_days_var.get())
            cache_days = int(self.wallpaper_cache_days_var.get())
            if interval < 1 or cache_days < 1:
                raise ValueError("天数必须大于0")
            if not self._normalize_time_string(self.wallpaper_change_time_var.get()):
                raise ValueError("时间格式不正确")
                
            self.settings['wallpaper_enabled'] = self.wallpaper_enabled_var.get()
            self.settings['wallpaper_interval_days'] = str(interval)
            self.settings['wallpaper_change_time'] = self.wallpaper_change_time_var.get()
            self.settings['wallpaper_cache_days'] = str(cache_days)
            
            self.save_settings() # 调用您已有的全局保存函数
            self.log("网络壁纸设置已保存。")
            messagebox.showinfo("成功", "网络壁纸设置已成功保存！", parent=self.root)

        except (ValueError, TypeError) as e:
            messagebox.showerror("输入错误", f"请检查输入内容是否为有效的数字和时间格式。\n\n错误: {e}", parent=self.root)

    def _trigger_wallpaper_change_now(self):
        """手动触发一次壁纸更换"""
        self.log("用户手动触发“立即更换壁纸”。")
        # 在后台线程中执行，避免UI卡顿
        threading.Thread(target=self._execute_wallpaper_task, args=(True,), daemon=True).start()

    def _open_wallpaper_folder(self):
        """打开壁纸缓存文件夹"""
        if os.path.exists(WALLPAPER_CACHE_FOLDER):
            try:
                os.startfile(WALLPAPER_CACHE_FOLDER)
            except Exception as e:
                self.log(f"打开壁纸文件夹失败: {e}")
                messagebox.showerror("错误", f"无法打开文件夹:\n{e}", parent=self.root)
        else:
            messagebox.showwarning("提示", "壁纸文件夹尚不存在，请先获取一次壁纸。", parent=self.root)

    def _clear_wallpaper_cache(self):
        """清理所有壁纸缓存"""
        if not os.path.exists(WALLPAPER_CACHE_FOLDER) or not os.listdir(WALLPAPER_CACHE_FOLDER):
            messagebox.showinfo("提示", "壁纸缓存文件夹为空，无需清理。", parent=self.root)
            return
            
        if messagebox.askyesno("确认操作", "您确定要删除“每日壁纸”文件夹下的所有图片吗？\n此操作不可恢复。", parent=self.root):
            try:
                for filename in os.listdir(WALLPAPER_CACHE_FOLDER):
                    file_path = os.path.join(WALLPAPER_CACHE_FOLDER, filename)
                    os.remove(file_path)
                self.log("已手动清理所有壁纸缓存。")
                messagebox.showinfo("成功", "所有壁纸缓存已成功清理！", parent=self.root)
            except Exception as e:
                self.log(f"清理壁纸缓存时发生错误: {e}")
                messagebox.showerror("错误", f"清理失败:\n{e}", parent=self.root)

    def _check_wallpaper_task(self, now):
        if not self.settings.get('wallpaper_enabled', False):
            return

        change_time = self.settings.get('wallpaper_change_time', '08:00:00')
        # 增加一个秒数容错，防止因微小延迟错过触发
        if now.strftime('%H:%M:%S') == change_time:
            interval_days = int(self.settings.get('wallpaper_interval_days', '1'))
            last_change_date_str = self.settings.get('wallpaper_last_change_date', '')

            should_change = False
            if not last_change_date_str:
                # 从未更换过，立即更换
                should_change = True
            else:
                try:
                    last_change_date = datetime.strptime(last_change_date_str, '%Y-%m-%d').date()
                    # 检查今天是否已经换过，避免在同一秒内重复触发
                    if now.date() > last_change_date and (now.date() - last_change_date).days >= interval_days:
                        should_change = True
                except (ValueError, TypeError):
                    # 日期格式错误，也更换一次以纠正状态
                    should_change = True

            if should_change:
                self.log("定时更换壁纸时间已到，开始执行...")
                # 在后台线程中执行，避免UI卡顿
                threading.Thread(target=self._execute_wallpaper_task, daemon=True).start()

    # --- ↓↓↓ 新增代码：核心的壁纸获取与设置函数 ↓↓↓ ---
    def _execute_wallpaper_task(self, is_manual_trigger=False):
        if not WIN32_AVAILABLE:
            self.log("错误：pywin32 库未安装，无法执行定时壁纸任务。")
            return

        try:
            # --- 1. 清理旧壁纸 ---
            # 只有在自动触发时，才检查是否需要清理
            if not is_manual_trigger:
                try:
                    cache_days = int(self.settings.get('wallpaper_cache_days', '7'))
                    # 为避免频繁清理，可以增加一个上次清理日期的判断，但目前简单实现也可以
                    for filename in os.listdir(WALLPAPER_CACHE_FOLDER):
                        file_path = os.path.join(WALLPAPER_CACHE_FOLDER, filename)
                        file_mod_time = os.path.getmtime(file_path)
                        if file_mod_time < (time.time() - cache_days * 24 * 3600):
                            os.remove(file_path)
                            self.log(f"已自动清理过期壁纸: {filename}")
                except Exception as e:
                    self.log(f"自动清理壁纸缓存时出错: {e}")

            # --- 2. 获取壁纸信息 ---
            self.log("正在从必应获取最新壁纸信息...")
            api_url = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=zh-CN"
            response = requests.get(api_url, timeout=10)
            response.raise_for_status()
            data = response.json()
            image_info = data["images"][0]
            image_url = f"https://www.bing.com{image_info['url']}"
            
            # --- 3. 下载壁纸 ---
            # 使用 URL 中的 HASH 值作为唯一文件名，避免重复
            try:
                image_hash = image_info.get('hsh', str(int(time.time())))
                image_filename = f"bing_{image_hash}.jpg"
            except:
                image_filename = f"bing_{datetime.now().strftime('%Y%m%d')}.jpg"

            image_path = os.path.join(WALLPAPER_CACHE_FOLDER, image_filename)

            if not os.path.exists(image_path):
                self.log(f"正在下载新壁纸: {image_filename} ...")
                image_response = requests.get(image_url, timeout=30, stream=True)
                image_response.raise_for_status()
                with open(image_path, 'wb') as f:
                    for chunk in image_response.iter_content(chunk_size=8192):
                        f.write(chunk)
                self.log("下载完成。")
            else:
                self.log(f"壁纸 '{image_filename}' 已存在于本地缓存。")

            # --- 4. 设置壁纸 ---
            self.log(f"正在设置桌面壁纸...")
            # 注意：路径必须是绝对路径
            abs_image_path = os.path.abspath(image_path)
            win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, abs_image_path, 1 + 2)
            self.log("桌面壁纸设置成功！")
            
            # --- 5. 更新记录 ---
            # 只有在自动触发成功后，才更新上次更换日期
            if not is_manual_trigger:
                self.settings['wallpaper_last_change_date'] = datetime.now().strftime('%Y-%m-%d')
                self.save_settings()

        except requests.exceptions.RequestException as e:
            self.log(f"获取网络壁纸失败（网络错误）: {e}")
            if is_manual_trigger:
                self.root.after(0, lambda: messagebox.showerror("错误", f"获取网络壁纸失败，请检查您的网络连接。\n\n{e}", parent=self.root))
        except Exception as e:
            self.log(f"执行壁纸任务时发生未知错误: {e}")
            if is_manual_trigger:
                self.root.after(0, lambda: messagebox.showerror("未知错误", f"执行时发生错误:\n{e}", parent=self.root))

    def _refresh_wallpaper_ui(self):
        # 刷新网络壁纸页面的UI状态，如果UI控件还未创建则直接返回
        if not hasattr(self, 'wallpaper_enabled_var'):
            return

        self.wallpaper_enabled_var.set(self.settings.get("wallpaper_enabled", False))
        self.wallpaper_interval_days_var.set(self.settings.get("wallpaper_interval_days", "1"))
        self.wallpaper_change_time_var.set(self.settings.get("wallpaper_change_time", "08:00:00"))
        self.wallpaper_cache_days_var.set(self.settings.get("wallpaper_cache_days", "7"))

    # --- ↑↑↑ 壁纸功能代码结束 ↑↑↑ ---

#↓以下是计时功能的全套代码
    def _refresh_timer_ui(self):
        if not hasattr(self, 'timer_mode_var'): return
    
        self.timer_duration_var.set(self.settings.get("timer_duration", "00:10:00"))
        self.timer_show_clock_var.set(self.settings.get("timer_show_clock", True))
        self.timer_play_sound_var.set(self.settings.get("timer_play_sound", True))
        self.timer_sound_file_var.set(self.settings.get("timer_sound_file", ""))

    def _build_timer_ui(self, parent_frame):
        # --- 1. 核心修改：改变父框架的布局为Grid，以分离滚动区和按钮区 ---
        parent_frame.rowconfigure(0, weight=1)  # 让第0行（滚动区）占据所有可用垂直空间
        parent_frame.columnconfigure(0, weight=1)

        # --- 2. 创建并放置可滚动框架 ---
        scrolled_frame = ScrolledFrame(parent_frame, autohide=True)
        scrolled_frame.grid(row=0, column=0, sticky="nsew") # 使用grid布局
        container = scrolled_frame.container

        # --- 3. 所有配置项依然放置在 container 中 (这部分代码不变) ---
        # --- 描述区 ---
        title_label = ttk.Label(container, text="全屏正/倒计时工具", font=self.font_14_bold, bootstyle="primary")
        title_label.pack(anchor="w", pady=(0, 5), padx=10)
        
        desc_text = "启动一个独立的、总在最前的全屏计时器，适用于会议、考试、活动等场景。按 ESC 键可随时退出。"
        desc_label = ttk.Label(container, text=desc_text, bootstyle="secondary")
        desc_label.pack(anchor="w", pady=(0, 15), padx=10, fill=X)
        
        ttk.Separator(container, orient=HORIZONTAL).pack(fill=X, pady=5, padx=10)

        # --- 计时设置 ---
        timer_lf = ttk.LabelFrame(container, text="计时设置", padding=15)
        timer_lf.pack(fill=X, pady=10, padx=10)

        mode_frame = ttk.Frame(timer_lf)
        mode_frame.pack(fill=X, pady=5)
        ttk.Label(mode_frame, text="模式:").pack(side=LEFT, padx=(0, 20))
        countdown_rb = ttk.Radiobutton(mode_frame, text="倒计时", variable=self.timer_mode_var, value="countdown")
        countdown_rb.pack(side=LEFT, padx=10)
        stopwatch_rb = ttk.Radiobutton(mode_frame, text="正计时", variable=self.timer_mode_var, value="stopwatch")
        stopwatch_rb.pack(side=LEFT, padx=10)

        duration_frame = ttk.Frame(timer_lf)
        duration_frame.pack(fill=X, pady=5)
        ttk.Label(duration_frame, text="目标时长:").pack(side=LEFT, padx=(0, 5))
        duration_entry = ttk.Entry(duration_frame, textvariable=self.timer_duration_var, font=self.font_11, width=12)
        duration_entry.pack(side=LEFT, padx=10)
        self._bind_mousewheel_to_entry(duration_entry, self._handle_time_scroll)
        ttk.Label(duration_frame, text="(HH:MM:SS)").pack(side=LEFT)

        infinite_check = ttk.Checkbutton(timer_lf, text="无限时长 (仅正计时可用)", variable=self.timer_infinite_var, bootstyle="round-toggle")
        infinite_check.pack(anchor="w", pady=5, padx=5)

        # --- 显示与提醒设置 ---
        options_lf = ttk.LabelFrame(container, text="附加选项", padding=15)
        options_lf.pack(fill=X, pady=10, padx=10)

        ttk.Checkbutton(options_lf, text="显示当前系统时间 (年月日星期)", variable=self.timer_show_clock_var, bootstyle="round-toggle").pack(anchor="w", pady=5)

        sound_frame = ttk.Frame(options_lf)
        sound_frame.pack(fill=X, pady=5)
        sound_check = ttk.Checkbutton(sound_frame, text="到达目标时长后播放提示音", variable=self.timer_play_sound_var, bootstyle="round-toggle")
        sound_check.pack(side=LEFT, anchor="w")
        
        sound_file_entry = ttk.Entry(sound_frame, textvariable=self.timer_sound_file_var, font=self.font_11)
        sound_file_entry.pack(side=LEFT, padx=10, expand=True, fill=X)
        
        def select_timer_sound():
            filepath = filedialog.askopenfilename(
                title="选择提示音文件",
                initialdir=PROMPT_FOLDER,
                filetypes=[("音频文件", "*.wav *.mp3"), ("所有文件", "*.*")],
                parent=self.root
            )
            if filepath:
                self.timer_sound_file_var.set(filepath)
        
        ttk.Button(sound_frame, text="选取...", command=select_timer_sound, bootstyle="outline").pack(side=LEFT)

        # --- 联动逻辑 (这部分代码不变) ---
        def update_timer_ui_states(*args):
            is_stopwatch = self.timer_mode_var.get() == 'stopwatch'
            is_infinite = self.timer_infinite_var.get()
            is_sound_enabled = self.timer_play_sound_var.get()

            infinite_check.config(state="normal" if is_stopwatch else "disabled")
            if not is_stopwatch:
                self.timer_infinite_var.set(False)
                is_infinite = False

            duration_entry.config(state="disabled" if is_stopwatch and is_infinite else "normal")
            
            can_play_sound = not (is_stopwatch and is_infinite)
            sound_check.config(state="normal" if can_play_sound else "disabled")
            
            sound_select_btn = None
            for child in sound_frame.winfo_children():
                if isinstance(child, ttk.Button):
                    sound_select_btn = child
                    break

            if can_play_sound and is_sound_enabled:
                sound_file_entry.config(state="normal")
                if sound_select_btn: sound_select_btn.config(state="normal")
            else:
                sound_file_entry.config(state="disabled")
                if sound_select_btn: sound_select_btn.config(state="disabled")
                if not can_play_sound:
                    self.timer_play_sound_var.set(False)

        self.timer_mode_var.trace_add("write", update_timer_ui_states)
        self.timer_infinite_var.trace_add("write", update_timer_ui_states)
        self.timer_play_sound_var.trace_add("write", update_timer_ui_states)
        
        self.root.after(100, update_timer_ui_states)

        # --- 4. 核心修改：创建独立的按钮框架，并放置在 parent_frame 的第1行 ---
        button_container = ttk.Frame(parent_frame, padding=(0, 10, 0, 10))
        button_container.grid(row=1, column=0, sticky="ew")
        
        # 为了让按钮居中，我们让容器的列可以扩展
        button_container.columnconfigure(0, weight=1)

        start_btn = ttk.Button(
            button_container, 
            text="启 动 全 屏 计 时 器", 
            command=self._start_timer_window, 
            bootstyle="success", 
            style="lg.TButton"
        )
        # 使用 grid 替代 pack，并且不设置 sticky，让它自然居中
        start_btn.grid(row=0, column=0, ipady=8, ipadx=50)

    def _start_timer_window(self):
        """
        读取UI设置，验证输入，并创建和启动计时器窗口。
        """
        # 0. 防止重复打开
        if self.timer_window and self.timer_window.winfo_exists():
            self.log("错误：计时器窗口已在运行中。")
            messagebox.showwarning("提示", "计时器已在运行中，请先关闭。", parent=self.root)
            return

        # 1. 读取并验证所有设置
        settings = {
            "mode": self.timer_mode_var.get(),
            "is_infinite": self.timer_infinite_var.get(),
            "show_clock": self.timer_show_clock_var.get(),
            "play_sound": self.timer_play_sound_var.get(),
            "sound_file": self.timer_sound_file_var.get().strip() or self.settings.get('reminder_sound', REMINDER_SOUND_FILE)
        }

        total_seconds = 0
        if not (settings["mode"] == "stopwatch" and settings["is_infinite"]):
            try:
                h, m, s = map(int, self.timer_duration_var.get().split(':'))
                total_seconds = h * 3600 + m * 60 + s
                if total_seconds <= 0:
                    raise ValueError("时长必须大于0")
            except Exception:
                messagebox.showerror("输入错误", "目标时长格式不正确或必须大于0秒。\n\n请使用 HH:MM:SS 格式。", parent=self.root)
                return
        
        settings["total_seconds"] = total_seconds

        # 2. 创建窗口并设置属性
        self.timer_window = ttk.Toplevel(self.root)
        self.timer_window.title("计时器")
        self.timer_window.configure(bg='black')
        self.timer_window.attributes('-fullscreen', True)
        self.timer_window.attributes('-topmost', True)

        self.root.attributes('-disabled', True) # 禁用主窗口
        self.is_fullscreen_exclusive = True # 设置“绝对霸权”标志
        self.log("全屏计时器已启动，其他全屏任务将被跳过。")

        # 绑定退出事件
        self.timer_window.bind('<Escape>', lambda e: self._close_timer_window())
        self.timer_window.protocol("WM_DELETE_WINDOW", self._close_timer_window)

        # 3. 动态计算字体大小并创建UI元素
        self.timer_window.update_idletasks()
        window_height = self.timer_window.winfo_height()
        window_width = self.timer_window.winfo_width()
        font_family = self.settings.get("app_font", "Microsoft YaHei")

        # 动态计算主计时器字体
        font_size = int(window_height * 0.5) # 从一个较大的估算值开始
        temp_font = font.Font(family=font_family, size=font_size, weight='bold')
        while temp_font.measure("00:00:00") > window_width * 0.9:
            font_size -= 5
            temp_font.config(size=font_size)
        
        main_font = (font_family, font_size, 'bold')
        small_font = (font_family, max(12, int(font_size / 8)), 'normal') # 底部时钟字体

        # 创建Label
        timer_label = ttk.Label(self.timer_window, font=main_font, foreground='white', background='black')
        timer_label.pack(expand=True)
        
        clock_label = None
        if settings['show_clock']:
            clock_label = ttk.Label(self.timer_window, font=small_font, foreground='lightgray', background='black')
            clock_label.pack(side="bottom", pady=20)
        
        # 4. 启动计时循环
        self._update_timer(timer_label, clock_label, settings)


    def _update_timer(self, timer_label, clock_label, settings):
        """
        每秒执行一次的计时器更新函数。
        """
        now = datetime.now()
        
        # 更新底部时钟（如果存在）
        if clock_label and clock_label.winfo_exists():
            week_map = {"1": "一", "2": "二", "3": "三", "4": "四", "5": "五", "6": "六", "7": "日"}
            day_of_week = week_map.get(str(now.isoweekday()), '')
            clock_str = now.strftime(f'%Y年%m月%d日  星期{day_of_week}  %H:%M:%S')
            clock_label.config(text=clock_str)

        # --- 计算主计时器时间 ---
        if 'start_time' not in settings: # 首次运行时记录起始时间
            settings['start_time'] = now

        elapsed = now - settings['start_time']
        
        time_is_up = False
        display_str = ""

        if settings['mode'] == 'countdown':
            remaining = timedelta(seconds=settings['total_seconds']) - elapsed
            if remaining.total_seconds() <= 0:
                time_is_up = True
                remaining = timedelta(seconds=0)
            
            # 格式化为 HH:MM:SS
            hours, rem = divmod(remaining.seconds, 3600)
            minutes, seconds = divmod(rem, 60)
            display_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"

        else: # stopwatch
            # 格式化为 HH:MM:SS
            hours, rem = divmod(int(elapsed.total_seconds()), 3600)
            minutes, seconds = divmod(rem, 60)
            display_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"

            if not settings['is_infinite'] and elapsed.total_seconds() >= settings['total_seconds']:
                time_is_up = True
        
        # 更新主计时器标签
        if timer_label.winfo_exists():
            timer_label.config(text=display_str)
        
        # --- 检查结束条件 ---
        if time_is_up:
            if timer_label.winfo_exists():
                timer_label.config(foreground='red') # 时间到，变红
            if settings['play_sound']:
                self._play_timer_end_sound(settings['sound_file'])
            
            # 时间到后，等待3秒自动关闭
            self.timer_after_id = self.root.after(3000, self._close_timer_window)
        else:
            # 预约下一次更新
            self.timer_after_id = self.root.after(1000, self._update_timer, timer_label, clock_label, settings)


    def _play_timer_end_sound(self, sound_file):
        """播放计时结束提示音"""
        if not AUDIO_AVAILABLE:
            self.log("警告：pygame未安装，无法播放计时结束提示音。")
            # 回退到系统蜂鸣声
            if WIN32_AVAILABLE: ctypes.windll.user32.MessageBeep(win32con.MB_OK)
            return

        try:
            if os.path.exists(sound_file):
                sound = pygame.mixer.Sound(sound_file)
                channel = pygame.mixer.find_channel(True) # 找一个空闲通道
                channel.set_volume(0.8) # 设置一个默认音量
                channel.play(sound)
                self.log(f"已播放计时结束提示音: {os.path.basename(sound_file)}")
            else:
                self.log(f"警告：计时结束提示音文件不存在: {sound_file}")
                if WIN32_AVAILABLE: ctypes.windll.user32.MessageBeep(win32con.MB_OK)
        except Exception as e:
            self.log(f"播放计时结束提示音失败: {e}")


    def _close_timer_window(self):
        """安全地关闭计时器窗口和相关资源"""
        if self.timer_after_id:
            self.root.after_cancel(self.timer_after_id)
            self.timer_after_id = None
        
        if self.timer_window and self.timer_window.winfo_exists():
            self.timer_window.destroy()
            self.timer_window = None
        
        self.root.attributes('-disabled', False) # 恢复主窗口
        self.root.focus_force()
        self.is_fullscreen_exclusive = False # 解除“绝对霸权”标志
        self.log("全屏计时器已关闭。")
#↑全套计时功能代码结束

# --- 动态语音功能的全套方法 ---

    def load_dynamic_voice_tasks(self):
        # 注意：动态语音任务是实验性功能，暂存在主任务文件里
        # 未来可以分离到 DYNAMIC_VOICE_TASK_FILE
        pass

    def save_dynamic_voice_tasks(self):
        # 数据随 self.tasks 一起保存，此函数暂时留空
        pass

    def clear_all_dynamic_voice_tasks(self):
        # 这是一个辅助函数，用于在重置软件时调用
        self.tasks = [t for t in self.tasks if t.get('type') != 'dynamic_voice']
        self.update_task_list()
        self.save_tasks()

    def open_dynamic_voice_dialog(self, parent_dialog, task_to_edit=None, index=None):
        parent_dialog.destroy()
        is_edit_mode = task_to_edit is not None
        dialog = ttk.Toplevel(self.root)
        dialog.title("修改动态语音" if is_edit_mode else "添加动态语音")
        dialog.resizable(True, True)
        dialog.minsize(800, 600)
        dialog.transient(self.root)

        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()

        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)

        content_frame = ttk.LabelFrame(main_frame, text="内容", padding=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=2)
        content_frame.columnconfigure(1, weight=1)

        ttk.Label(content_frame, text="节目名称:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        name_entry = ttk.Entry(content_frame, font=self.font_11)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        
        ttk.Label(content_frame, text="播音文稿:").grid(row=1, column=0, sticky='nw', padx=5, pady=2)
        text_frame = ttk.Frame(content_frame)
        text_frame.grid(row=1, column=1, columnspan=3, sticky='nsew', padx=5, pady=2)
        content_frame.rowconfigure(1, weight=1)
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        content_text = ScrolledText(text_frame, height=5, font=self.font_11, wrap=WORD)
        content_text.grid(row=0, column=0, sticky='nsew')
        
        script_btn_frame = ttk.Frame(content_frame)
        script_btn_frame.grid(row=2, column=1, columnspan=3, sticky='w', padx=5, pady=(5, 2))
        
        def insert_tag(tag):
            try:
                content_text.text.insert(tk.INSERT, tag)
                content_text.text.focus_set()
            except Exception as e:
                self.log(f"插入标记失败: {e}")

        tags = ["[年月日]", "[星期]", "[时间]", "[天气]", "[男]", "[女]"]
        for tag in tags:
            ttk.Button(script_btn_frame, text=tag, bootstyle="outline", command=lambda t=tag: insert_tag(t)).pack(side=LEFT, padx=2)

        params_frame = ttk.LabelFrame(main_frame, text="通用参数", padding=10)
        params_frame.grid(row=1, column=0, sticky='ew', pady=4)
        # 该Frame内部将使用pack，不再需要columnconfigure

        # --- ↓↓↓ 全新、正确的布局代码 ---
        
        # 第一行：语速、音调、音量
        speech_params_container = ttk.Frame(params_frame)
        speech_params_container.pack(fill=X, pady=3, padx=5)

        ttk.Label(speech_params_container, text="整体语速(-10~10):").pack(side=LEFT, padx=(0, 2))
        speed_entry = ttk.Entry(speech_params_container, font=self.font_11, width=5)
        speed_entry.pack(side=LEFT, padx=(0, 15))

        ttk.Label(speech_params_container, text="整体音调(-10~10):").pack(side=LEFT, padx=(0, 2))
        pitch_entry = ttk.Entry(speech_params_container, font=self.font_11, width=5)
        pitch_entry.pack(side=LEFT, padx=(0, 15))
        
        ttk.Label(speech_params_container, text="整体音量(0-100):").pack(side=LEFT, padx=(0, 2))
        volume_entry = ttk.Entry(speech_params_container, font=self.font_11, width=5)
        volume_entry.pack(side=LEFT)

        # 第二行：提示音
        prompt_container = ttk.Frame(params_frame)
        prompt_container.pack(fill=X, pady=3, padx=5)
        prompt_container.columnconfigure(1, weight=1) # 让文件路径输入框可伸缩

        prompt_var = tk.IntVar()
        ttk.Checkbutton(prompt_container, text="提示音:", variable=prompt_var, bootstyle="round-toggle").grid(row=0, column=0, sticky='w')
        
        prompt_file_var, prompt_volume_var = tk.StringVar(), tk.StringVar()
        prompt_file_entry = ttk.Entry(prompt_container, textvariable=prompt_file_var, font=self.font_11)
        prompt_file_entry.grid(row=0, column=1, sticky='ew', padx=5)
        
        ttk.Button(prompt_container, text="...", command=lambda: self.select_file_for_entry(PROMPT_FOLDER, prompt_file_var, dialog), bootstyle="outline", width=2).grid(row=0, column=2, padx=(0, 10))
        
        ttk.Label(prompt_container, text="音量:").grid(row=0, column=3, sticky='e')
        ttk.Entry(prompt_container, textvariable=prompt_volume_var, font=self.font_11, width=8).grid(row=0, column=4, sticky='w', padx=5)

        # 第三行：背景音乐
        bgm_container = ttk.Frame(params_frame)
        bgm_container.pack(fill=X, pady=3, padx=5)
        bgm_container.columnconfigure(1, weight=1) # 让文件路径输入框可伸缩

        bgm_var = tk.IntVar()
        ttk.Checkbutton(bgm_container, text="背景音乐:", variable=bgm_var, bootstyle="round-toggle").grid(row=0, column=0, sticky='w')
        
        bgm_file_var, bgm_volume_var = tk.StringVar(), tk.StringVar()
        bgm_file_entry = ttk.Entry(bgm_container, textvariable=bgm_file_var, font=self.font_11)
        bgm_file_entry.grid(row=0, column=1, sticky='ew', padx=5)
        
        ttk.Button(bgm_container, text="...", command=lambda: self.select_file_for_entry(BGM_FOLDER, bgm_file_var, dialog), bootstyle="outline", width=2).grid(row=0, column=2, padx=(0, 10))
        
        ttk.Label(bgm_container, text="音量:").grid(row=0, column=3, sticky='e')
        ttk.Entry(bgm_container, textvariable=bgm_volume_var, font=self.font_11, width=8).grid(row=0, column=4, sticky='w', padx=5)
        # --- ↑↑↑ 布局代码结束 ---

        time_frame = ttk.LabelFrame(main_frame, text="时间规则", padding=15)
        time_frame.grid(row=2, column=0, sticky='ew', pady=4)
        time_frame.columnconfigure(1, weight=1)
        
        # ... (time_frame 及其内部的代码保持不变) ...
        ttk.Label(time_frame, text="执行时间:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        start_time_entry = ttk.Entry(time_frame, font=self.font_11)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        self._bind_mousewheel_to_entry(start_time_entry, self._handle_time_scroll)
        ttk.Label(time_frame, text="<可多个>").grid(row=0, column=2, sticky='w', padx=5)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bootstyle="outline").grid(row=0, column=3, padx=5)
        
        ttk.Label(time_frame, text="周几/几号:").grid(row=1, column=0, sticky='e', padx=5, pady=3)
        weekday_entry = ttk.Entry(time_frame, font=self.font_11)
        weekday_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=3)
        ttk.Button(time_frame, text="选取...", command=lambda: self.show_weekday_settings_dialog(weekday_entry), bootstyle="outline").grid(row=1, column=3, padx=5)
        
        ttk.Label(time_frame, text="日期范围:").grid(row=2, column=0, sticky='e', padx=5, pady=3)
        date_range_entry = ttk.Entry(time_frame, font=self.font_11)
        date_range_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=3)
        self._bind_mousewheel_to_entry(date_range_entry, self._handle_date_scroll)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry), bootstyle="outline").grid(row=2, column=3, padx=5)

        dialog_button_frame = ttk.Frame(dialog)
        dialog_button_frame.pack(pady=15)

        # ... (数据加载和保存逻辑保持不变) ...
        if is_edit_mode:
            name_entry.insert(0, task_to_edit.get('name', ''))
            content_text.text.insert('1.0', task_to_edit.get('source_text', ''))
            speed_entry.insert(0, task_to_edit.get('speed', '0'))
            pitch_entry.insert(0, task_to_edit.get('pitch', '0'))
            volume_entry.insert(0, task_to_edit.get('volume', '100'))
            prompt_var.set(task_to_edit.get('prompt', 0))
            prompt_file_var.set(task_to_edit.get('prompt_file', ''))
            prompt_volume_var.set(task_to_edit.get('prompt_volume', '80'))
            bgm_var.set(task_to_edit.get('bgm', 0))
            bgm_file_var.set(task_to_edit.get('bgm_file', ''))
            bgm_volume_var.set(task_to_edit.get('bgm_volume', '20'))
            start_time_entry.insert(0, task_to_edit.get('time', ''))
            weekday_entry.insert(0, task_to_edit.get('weekday', '每周:1234567'))
            date_range_entry.insert(0, task_to_edit.get('date_range', '2025-01-01 ~ 2099-12-31'))
        else:
            speed_entry.insert(0, "0")
            pitch_entry.insert(0, "0")
            volume_entry.insert(0, "100")
            prompt_volume_var.set("80")
            bgm_volume_var.set("20")
            weekday_entry.insert(0, "每周:1234567")
            date_range_entry.insert(0, "2025-01-01 ~ 2099-12-31")

        def save_task():
            text_content = content_text.text.get('1.0', END).strip()
            if not text_content:
                messagebox.showwarning("警告", "请输入播音文稿内容", parent=dialog)
                return

            is_valid_time, time_msg = self._normalize_multiple_times_string(start_time_entry.get().strip())
            if not is_valid_time: messagebox.showwarning("格式错误", time_msg, parent=dialog); return
            is_valid_date, date_msg = self._normalize_date_range_string(date_range_entry.get().strip())
            if not is_valid_date: messagebox.showwarning("格式错误", date_msg, parent=dialog); return

            new_task_data = {
                'name': name_entry.get().strip(),
                'type': 'dynamic_voice',
                'source_text': text_content,
                'speed': speed_entry.get().strip() or "0",
                'pitch': pitch_entry.get().strip() or "0",
                'volume': volume_entry.get().strip() or "100",
                'prompt': prompt_var.get(),
                'prompt_file': prompt_file_var.get(),
                'prompt_volume': prompt_volume_var.get(),
                'bgm': bgm_var.get(),
                'bgm_file': bgm_file_var.get(),
                'bgm_volume': bgm_volume_var.get(),
                'time': time_msg,
                'weekday': weekday_entry.get().strip(),
                'date_range': date_msg,
                'delay': 'ontime', 
                'status': '启用' if not is_edit_mode else task_to_edit.get('status', '启用'),
                'last_run': {} if not is_edit_mode else task_to_edit.get('last_run', {}),
            }
            if not new_task_data['name'] or not new_task_data['time']: 
                messagebox.showwarning("警告", "请填写任务名称和执行时间", parent=dialog); return

            if is_edit_mode:
                self.tasks[index] = new_task_data
                self.log(f"已修改动态语音任务: {new_task_data['name']}")
            else:
                self.tasks.append(new_task_data)
                self.log(f"已添加动态语音任务: {new_task_data['name']}")

            self.update_task_list()
            self.save_tasks()
            cleanup_and_destroy()

        button_text = "保存修改" if is_edit_mode else "添加"
        ttk.Button(dialog_button_frame, text=button_text, command=save_task, bootstyle="primary").pack(side=LEFT, padx=10, ipady=5)
        ttk.Button(dialog_button_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10, ipady=5)
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)
        
        self.center_window(dialog, parent=self.root)
#代码结束

    def _parse_dynamic_script(self, script_text):
        segments = []
        # 使用正则表达式分割脚本，同时保留分割标记
        parts = re.split(r'(\[男\]|\[女\])', script_text)
        
        # 默认第一个片段是男声，除非脚本以[女]开头
        current_actor = '男'
        if script_text.strip().startswith('[女]'):
            current_actor = '女'

        for part in parts:
            if not part.strip():
                continue
            if part == '[男]':
                current_actor = '男'
            elif part == '[女]':
                current_actor = '女'
            else:
                segments.append({'actor': current_actor, 'text': part.strip()})
        return segments

    def _replace_dynamic_tags(self, text, trigger_time_obj):
        now = trigger_time_obj
        
        weather_info = self.main_weather_label.cget("text")
        weather_text = "未知"
        if "天气:" in weather_info and "正在" not in weather_info and "失败" not in weather_info:
            try:
                weather_text = weather_info.split(' ')[2]
            except IndexError:
                weather_text = "信息不全"

        week_map = {"1": "一", "2": "二", "3": "三", "4": "四", "5": "五", "6": "六", "7": "日"}
        day_of_week = week_map.get(str(now.isoweekday()), '')

        replacements = {
            "[年月日]": now.strftime('%Y年%m月%d日'),
            "[星期]": f"星期{day_of_week}",
            "[时间]": now.strftime('%H点%M分'),
            "[天气]": weather_text
        }
        
        for tag, value in replacements.items():
            text = text.replace(tag, value)
        return text

    def _execute_dynamic_voice_task(self, task):
        # **核心逻辑：播放前先检查缓存**
        final_audio_path = None
        
        # 检查内存中的任务对象是否有缓存路径的记录
        if 'cached_audio_path' in task and task.get('cached_audio_path'):
            # 再次确认磁盘上这个文件是否真的存在
            if os.path.exists(task['cached_audio_path']):
                final_audio_path = task['cached_audio_path']
                self.log(f"任务 '{task['name']}' 命中缓存，直接使用预生成的音频。")
            else:
                self.log(f"警告：任务 '{task['name']}' 缓存记录存在，但文件丢失，将重新生成。")
        
        # --- 如果缓存检查失败 (final_audio_path 依然是 None)，则执行旧的即时合成逻辑 ---
        if final_audio_path is None:
            self.log(f"任务 '{task['name']}' 未命中缓存，开始即时生成...")
            
            source_text = task.get('source_text', '')
            if not source_text:
                self.log(f"动态语音任务 '{task['name']}' 文稿为空，已跳过。")
                return

            segments = self._parse_dynamic_script(source_text)
            if not segments:
                self.log(f"动态语音任务 '{task['name']}' 未能解析出有效片段。")
                return

            temp_files = []
            final_audio = None
            
            try:
                from pydub import AudioSegment
                ffmpeg_path = os.path.join(application_path, "ffmpeg.exe")
                if os.path.exists(ffmpeg_path):
                    AudioSegment.converter = ffmpeg_path
            except ImportError:
                self.log("警告：pydub库未安装，无法拼接动态语音。")
                return

            try:
                for i, segment in enumerate(segments):
                    if self._is_interrupted():
                        self.log("动态语音生成被中断。")
                        return

                    self.update_playing_text(f"[{task['name']}] 正在生成第 {i+1}/{len(segments)} 段语音...")
                    
                    processed_text = self._replace_dynamic_tags(segment['text'], datetime.now())
                    
                    actor = segment['actor']
                    voice_name = '在线-云扬 (男)' if actor == '男' else '在线-晓晓 (女)'
                    voice_params = { 'voice': voice_name, 'speed': task.get('speed', '0'), 'pitch': task.get('pitch', '0') }
                    
                    temp_filename = f"temp_runtime_{int(time.time())}_{i}.mp3"
                    output_path = os.path.join(AUDIO_FOLDER, temp_filename)
                    temp_files.append(output_path)

                    synthesis_success = threading.Event()
                    error_message = ""
                    def online_callback(result):
                        nonlocal error_message
                        if not result['success']: error_message = result.get('error', '未知在线合成错误')
                        synthesis_success.set()
                    
                    s_thread = threading.Thread(target=self._synthesis_worker_edge, args=(processed_text, voice_params, output_path, online_callback))
                    s_thread.start()
                    s_thread.join()

                    if error_message: raise Exception(f"生成片段 '{processed_text}' 时失败: {error_message}")

                self.update_playing_text(f"[{task['name']}] 正在合成最终音频...")
                for file_path in temp_files:
                    segment_audio = AudioSegment.from_mp3(file_path)
                    if final_audio is None: final_audio = segment_audio
                    else: final_audio += segment_audio
                
                if final_audio is None: raise Exception("未能生成任何有效的音频片段。")

                final_audio_path_runtime = os.path.join(AUDIO_FOLDER, f"final_runtime_{int(time.time())}.wav")
                temp_files.append(final_audio_path_runtime)
                final_audio.export(final_audio_path_runtime, format="wav")
                final_audio_path = final_audio_path_runtime

            except Exception as e:
                self.log(f"!!! 即时生成动态语音任务 '{task['name']}' 失败: {e}")
                for f in temp_files:
                    if os.path.exists(f):
                        try: os.remove(f)
                        except: pass
                return
            finally:
                segment_files = [f for f in temp_files if "temp_runtime" in f]
                for f in segment_files:
                    if os.path.exists(f):
                        try: os.remove(f)
                        except: pass
        
        # --- 统一的播放逻辑 ---
        if final_audio_path and os.path.exists(final_audio_path):
            try:
                final_task = task.copy()
                final_task['content'] = final_audio_path
                final_task['repeat'] = 1
                
                self._play_voice_task_internal(final_task)
            finally:
                if "runtime" in os.path.basename(final_audio_path):
                    if os.path.exists(final_audio_path):
                        try: os.remove(final_audio_path)
                        except Exception as e_del: self.log(f"删除即时生成的最终文件失败: {e_del}")
        else:
            self.log(f"!!! 最终播放错误：找不到有效的音频文件用于任务 '{task['name']}'")

    def _pre_generate_dynamic_voice(self, task, trigger_time):
        """
        预生成动态语音任务的音频文件，并将其缓存。
        这是一个后台函数，只合成，不播放。
        """
        # 检查任务是否已经被缓存了，如果是，就直接返回，避免重复工作
        if task.get('cached_audio_path') and task.get('cached_for_time') == trigger_time:
            if os.path.exists(task.get('cached_audio_path')):
                return

        # 为这个任务和触发时间生成一个唯一的文件名
        safe_task_name = re.sub(r'[\\/*?:"<>|]', "", task['name'])
        safe_trigger_time = trigger_time.replace(":", "-")
        cache_filename = f"cache_{safe_task_name}_{safe_trigger_time}.wav"
        cache_filepath = os.path.join(DYNAMIC_VOICE_CACHE_FOLDER, cache_filename)

        self.log(f"开始为任务 '{task['name']}' ({trigger_time}) 预生成动态语音...")

        source_text = task.get('source_text', '')
        if not source_text:
            self.log(f"预生成失败：任务 '{task['name']}' 文稿为空。")
            return

        segments = self._parse_dynamic_script(source_text)
        if not segments:
            self.log(f"预生成失败：任务 '{task['name']}' 未能解析出有效片段。")
            return

        temp_segment_files = []
        final_audio = None

        try:
            from pydub import AudioSegment
            ffmpeg_path = os.path.join(application_path, "ffmpeg.exe")
            if os.path.exists(ffmpeg_path):
                AudioSegment.converter = ffmpeg_path
        except ImportError:
            self.log("预生成失败：pydub 库未安装。")
            return

        try:
            # 1. 循环合成每个片段
            for i, segment in enumerate(segments):
                # **核心修正**：使用目标触发时间来替换占位符
                target_dt_obj = datetime.now().replace(
                    hour=int(trigger_time[0:2]),
                    minute=int(trigger_time[3:5]),
                    second=int(trigger_time[6:8]),
                    microsecond=0
                )
                processed_text = self._replace_dynamic_tags(segment['text'], target_dt_obj)
                
                actor = segment['actor']
                voice_name = '在线-云扬 (男)' if actor == '男' else '在线-晓晓 (女)'
                voice_params = {
                    'voice': voice_name,
                    'speed': task.get('speed', '0'),
                    'pitch': task.get('pitch', '0')
                }

                temp_segment_filename = f"temp_pregen_{int(time.time())}_{i}.mp3"
                output_path = os.path.join(AUDIO_FOLDER, temp_segment_filename)
                temp_segment_files.append(output_path)

                synthesis_success = threading.Event()
                error_message = ""
                def online_callback(result):
                    nonlocal error_message
                    if not result['success']:
                        error_message = result.get('error', '未知在线合成错误')
                    synthesis_success.set()
                
                s_thread = threading.Thread(target=self._synthesis_worker_edge, args=(processed_text, voice_params, output_path, online_callback))
                s_thread.start()
                s_thread.join()

                if error_message:
                    raise Exception(f"生成片段时失败: {error_message}")

            # 2. 拼接所有片段
            for file_path in temp_segment_files:
                segment_audio = AudioSegment.from_mp3(file_path)
                if final_audio is None:
                    final_audio = segment_audio
                else:
                    final_audio += segment_audio
            
            if final_audio is None:
                raise Exception("未能生成任何有效的音频片段。")

            # 3. 导出到最终的缓存文件
            final_audio.export(cache_filepath, format="wav")

            # 4. 在内存中更新任务对象，记录缓存文件的路径
            task['cached_audio_path'] = cache_filepath
            task['cached_for_time'] = trigger_time 

            self.log(f"任务 '{task['name']}' 预生成成功！缓存文件: {cache_filename}")

        except Exception as e:
            self.log(f"!!! 预生成任务 '{task['name']}' 失败: {e}")
            if 'cached_audio_path' in task: del task['cached_audio_path']
            if 'cached_for_time' in task: del task['cached_for_time']
        finally:
            # 5. 清理临时的片段文件
            for f in temp_segment_files:
                if os.path.exists(f):
                    try: os.remove(f)
                    except Exception as e_del: self.log(f"删除预生成临时文件 {os.path.basename(f)} 失败: {e_del}")

#以上动态语音全套方法结束

    def create_registration_page(self):
        page_frame = ttk.Frame(self.page_container, padding=20)
        title_label = ttk.Label(page_frame, text="注册软件", font=self.font_14_bold, bootstyle="primary")
        title_label.pack(anchor=W)

        main_content_frame = ttk.Frame(page_frame)
        main_content_frame.pack(pady=10, fill=X, expand=True)

        machine_code_frame = ttk.Frame(main_content_frame)
        machine_code_frame.pack(fill=X, pady=10)
        ttk.Label(machine_code_frame, text="机器码:", font=self.font_12).pack(side=LEFT)
        machine_code_val = self.get_machine_code()
        machine_code_entry = ttk.Entry(machine_code_frame, font=self.font_12, bootstyle="danger")
        machine_code_entry.pack(side=LEFT, padx=10, fill=X, expand=True)
        machine_code_entry.insert(0, machine_code_val)
        machine_code_entry.config(state='readonly')

        reg_code_frame = ttk.Frame(main_content_frame)
        reg_code_frame.pack(fill=X, pady=10)
        ttk.Label(reg_code_frame, text="注册码:", font=self.font_12).pack(side=LEFT)
        self.reg_code_entry = ttk.Entry(reg_code_frame, font=self.font_12)
        self.reg_code_entry.pack(side=LEFT, padx=10, fill=X, expand=True)

        btn_container = ttk.Frame(main_content_frame)
        btn_container.pack(pady=20)

        register_btn = ttk.Button(btn_container, text="注 册",
                                 bootstyle="success", style='lg.TButton', command=self.attempt_registration)
        register_btn.pack(pady=5, fill=X)

        cancel_reg_btn = ttk.Button(btn_container, text="取消注册",
                                   bootstyle="danger", style='lg.TButton', command=self.cancel_registration)
        cancel_reg_btn.pack(pady=5, fill=X)
        
        style = ttk.Style.get_instance()
        style.configure('lg.TButton', font=self.font_12_bold)

        info_text = "请将您的机器码发送给软件提供商以获取注册码。\n注册码分为月度授权和永久授权两种。"
        ttk.Label(main_content_frame, text=info_text, font=self.font_10, bootstyle="secondary").pack(pady=10)

        return page_frame
        
#第2部分
    def cancel_registration(self):
        if not messagebox.askyesno("确认操作", "您确定要取消当前注册吗？\n取消后，软件将恢复到试用或过期状态。", parent=self.root):
            return

        self.log("用户请求取消注册...")
        self._save_to_registry('RegistrationStatus', '')
        self._save_to_registry('RegistrationDate', '')
        self._save_to_registry('LicenseSignature', '')
        self._save_to_registry('LastSeenSignature', '')

        self.check_authorization()

        messagebox.showinfo("操作完成", f"注册已成功取消。\n当前授权状态: {self.auth_info['message']}", parent=self.root)
        self.log(f"注册已取消。新状态: {self.auth_info['message']}")

        if self.is_app_locked_down:
            self.perform_lockdown()
        else:
            if self.current_page == self.pages.get("注册软件"):
                 self.switch_page("定时广播")

    def get_machine_code(self):
        if self.machine_code:
            return self.machine_code

        if not PSUTIL_AVAILABLE:
            messagebox.showerror("依赖缺失", "psutil 库未安装，无法获取机器码。软件将退出。", parent=self.root)
            self.root.destroy()
            sys.exit()

        try:
            mac = self._get_mac_address()
            if mac:
                substitution = str.maketrans("ABCDEF", "123456")
                numeric_mac = mac.upper().translate(substitution)
                self.machine_code = numeric_mac
                return self.machine_code
            else:
                raise Exception("未找到有效的有线或无线网络适配器。")
        except Exception as e:
            messagebox.showerror("错误", f"无法获取机器码：{e}\n软件将退出。", parent=self.root)
            self.root.destroy()
            sys.exit()

    def _get_mac_address(self):
        """
        获取一个稳定的物理MAC地址，优先有线网卡。
        这个版本不再依赖于网络连接状态('is_up')，使其更加稳定。
        """
        interfaces = psutil.net_if_addrs()
        
        mac_addresses = []
        for name, addrs in interfaces.items():
            # 过滤掉虚拟网卡和回环地址
            if 'loopback' in name.lower() or 'virtual' in name.lower() or name.startswith('vEthernet'):
                continue
            
            for addr in addrs:
                if addr.family == psutil.AF_LINK:
                    mac = addr.address.replace(':', '').replace('-', '').upper()
                    if len(mac) == 12 and mac != '000000000000':
                        is_wired = 'ethernet' in name.lower() or 'eth' in name.lower() or '本地连接' in name.lower()
                        # 赋予有线网卡更高的优先级
                        priority = 0 if is_wired else 1
                        mac_addresses.append((priority, mac, name))

        if not mac_addresses:
            return None

        # 按优先级（有线优先）、然后按名称排序，确保每次都得到相同的结果
        mac_addresses.sort()
        
        # 返回最优先的那个MAC地址
        #self.log(f"找到的最稳定MAC地址来自网卡: {mac_addresses[0][2]}")
        return mac_addresses[0][1]

    def _generate_signature(self, license_type, date_str):
        """根据机器码、授权类型、日期和密钥盐生成SHA-256签名"""
        machine_code = self.get_machine_code()
        # 将所有部分组合成一个不可变的字符串进行哈希
        data_to_hash = f"{machine_code}|{license_type}|{date_str}|{SECRET_SALT}"
        # 使用 SHA-256 算法生成十六进制格式的签名
        signature = hashlib.sha256(data_to_hash.encode('utf-8')).hexdigest()
        return signature

    def _calculate_reg_codes(self, numeric_mac_str):
        try:
            monthly_code = int(int(numeric_mac_str) * 3.14)

            reversed_mac_str = numeric_mac_str[::-1]
            permanent_val = int(reversed_mac_str) / 3.14
            permanent_code = f"{permanent_val:.2f}"

            return {'monthly': str(monthly_code), 'permanent': permanent_code}
        except (ValueError, TypeError):
            return {'monthly': None, 'permanent': None}

    # 第2部分 (替换整个函数)
    def attempt_registration(self):
        entered_code = self.reg_code_entry.get().strip()
        if not entered_code:
            messagebox.showwarning("提示", "请输入注册码。", parent=self.root)
            return

        numeric_machine_code = self.get_machine_code()
        correct_codes = self._calculate_reg_codes(numeric_machine_code)

        today_str = datetime.now().strftime('%Y-%m-%d')
        license_type = None

        if entered_code == correct_codes['monthly']:
            license_type = 'Monthly'
            messagebox.showinfo("注册成功", "恭喜您，月度授权已成功激活！", parent=self.root)
        elif entered_code == correct_codes['permanent']:
            license_type = 'Permanent'
            messagebox.showinfo("注册成功", "恭喜您，永久授权已成功激活！", parent=self.root)
        else:
            messagebox.showerror("注册失败", "您输入的注册码无效，请重新核对。", parent=self.root)
            return

        if license_type:
            # --- 核心修改：生成并保存签名 ---
            signature = self._generate_signature(license_type, today_str)
            self._save_to_registry('RegistrationStatus', license_type)
            self._save_to_registry('RegistrationDate', today_str)
            self._save_to_registry('LicenseSignature', signature) # <-- 新增：保存签名
            # --- 修改结束 ---
            self.check_authorization()

    def _create_sentinels(self):
        """
        在所有预定位置创建哨兵文件和注册表键。
        """
        #self.log("首次运行，正在创建防篡改哨兵...")
        machine_code = self.get_machine_code()
        
        for stype, path, name, _ in SENTINEL_LOCATIONS:
            try:
                if stype == 'file':
                    # 写入机器码，防止用户从别的电脑复制哨兵文件
                    with open(path, 'w') as f:
                        f.write(machine_code)
                    # 尝试将文件设置为隐藏
                    if sys.platform == "win32":
                        ctypes.windll.kernel32.SetFileAttributesW(path, 2) # 2 = FILE_ATTRIBUTE_HIDDEN
                    #self.log(f"成功创建文件哨兵: {path}")

                elif stype == 'reg':
                    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, path)
                    winreg.SetValueEx(key, name, 0, winreg.REG_SZ, machine_code)
                    winreg.CloseKey(key)
                    #self.log(f"成功创建注册表哨兵: HKEY_CURRENT_USER\\{path}\\{name}")

            except Exception as e:
                # 即使某个位置写入失败（如权限不足），也继续尝试下一个
                #self.log(f"警告: 创建哨兵失败 - {stype} at {path} - 原因: {e}")
                pass

    def _check_for_sentinels(self):
        """
        检查任何一个哨兵位置是否存在。只要找到一个，就返回True。
        """
        machine_code = self.get_machine_code()

        for stype, path, name, _ in SENTINEL_LOCATIONS:
            try:
                if stype == 'file':
                    if os.path.exists(path):
                        # 可选增强：检查文件内容是否匹配当前机器码
                        try:
                            with open(path, 'r') as f:
                                content = f.read()
                            if content == machine_code:
                                #self.log(f"检测到有效的文件哨兵: {path}")
                                return True
                        except:
                            # 文件存在但无法读取或内容不匹配，也算作一个标记
                            #self.log(f"检测到可疑的文件哨兵: {path}")
                            return True
                            
                elif stype == 'reg':
                    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_READ)
                    value, _ = winreg.QueryValueEx(key, name)
                    winreg.CloseKey(key)
                    # 可选增强：检查注册表值是否匹配
                    if value == machine_code:
                        #self.log(f"检测到有效的注册表哨兵: HKEY_CURRENT_USER\\{path}\\{name}")
                        return True

            except FileNotFoundError:
                # 这个是正常情况，意味着没找到
                continue
            except Exception as e:
                # 发生其他错误，例如权限问题，我们保守地认为哨兵可能存在
                #self.log(f"警告: 检查哨兵时发生错误 - {stype} at {path} - 原因: {e}")
                continue
        
        # 遍历完所有位置都没找到
        return False

    def check_authorization(self):
        today = datetime.now().date()
        today_str = today.strftime('%Y-%m-%d')
        
        # 1. 终极加固的时间回拨检测
        time_tampered = False
        last_seen_date_str = self._load_from_registry('LastSeenDate')
        last_seen_signature = self._load_from_registry('LastSeenSignature')

        if last_seen_date_str and last_seen_signature:
            # 如果存在上一次的记录，必须验证其签名
            expected_signature = self._generate_signature('LastSeen', last_seen_date_str)
            if last_seen_signature != expected_signature:
                # 签名不匹配，意味着 LastSeenDate 被篡改！
                #self.log("安全警告：检测到 LastSeenDate 被篡改，授权立即失效。")
                time_tampered = True
            else:
                # 签名匹配，LastSeenDate 可信，现在才进行时间比较
                try:
                    last_seen_date = datetime.strptime(last_seen_date_str, '%Y-%m-%d').date()
                    if today < last_seen_date:
                        #self.log("安全警告：检测到系统时间被回调，授权立即失效。")
                        time_tampered = True
                except (ValueError, TypeError):
                    # 日期格式错误，也视为篡改
                    time_tampered = True
        
        # 在所有检查之后，无论是否被篡改，都用今天的日期和新签名覆盖旧的记录
        new_signature_for_today = self._generate_signature('LastSeen', today_str)
        self._save_to_registry('LastSeenDate', today_str)
        self._save_to_registry('LastSeenSignature', new_signature_for_today)

        if time_tampered:
            self.auth_info = {'status': 'Expired', 'message': '授权已过期，请注册'}
            self.is_app_locked_down = True
            self.update_title_bar()
            return
            
        # 2. 读取主授权信息和签名 (这部分逻辑不变)
        status = self._load_from_registry('RegistrationStatus')
        reg_date_str = self._load_from_registry('RegistrationDate')
        stored_signature = self._load_from_registry('LicenseSignature')

        # 3. 核心验证逻辑 (这部分逻辑不变)
        if status and reg_date_str and stored_signature:
            expected_signature = self._generate_signature(status, reg_date_str)
            if stored_signature != expected_signature:
                self.log(f"安全警告：检测到无效或被篡改的授权信息 (状态: {status})。")
                self.auth_info = {'status': 'Expired', 'message': '授权信息损坏，请重新注册'}
                self.is_app_locked_down = True
            else:
                # 签名匹配，数据可信，判断有效期
                if status == 'Permanent':
                    self.auth_info = {'status': 'Permanent', 'message': '永久授权'}
                    self.is_app_locked_down = False
                elif status == 'Monthly':
                    try:
                        reg_date = datetime.strptime(reg_date_str, '%Y-%m-%d').date()
                        expiry_date = reg_date + timedelta(days=30)
                        if today > expiry_date:
                            self.auth_info = {'status': 'Expired', 'message': '授权已过期，请注册'}
                            self.is_app_locked_down = True
                        else:
                            remaining_days = (expiry_date - today).days
                            self.auth_info = {'status': 'Monthly', 'message': f'月度授权 - 剩余 {remaining_days} 天'}
                            self.is_app_locked_down = False
                    except (TypeError, ValueError):
                        self.auth_info = {'status': 'Expired', 'message': '授权信息损坏[M]'}
                        self.is_app_locked_down = True
                elif status == 'Trial':
                     try:
                        first_run_date = datetime.strptime(reg_date_str, '%Y-%m-%d').date()
                        trial_expiry_date = first_run_date + timedelta(days=3)
                        if today > trial_expiry_date:
                            self.auth_info = {'status': 'Expired', 'message': '授权已过期，请注册'}
                            self.is_app_locked_down = True
                        else:
                            remaining_days = (trial_expiry_date - today).days
                            self.auth_info = {'status': 'Trial', 'message': f'未注册 - 剩余 {remaining_days} 天'}
                            self.is_app_locked_down = False
                     except (TypeError, ValueError):
                        self.auth_info = {'status': 'Expired', 'message': '授权信息损坏[T]'}
                        self.is_app_locked_down = True
                else:
                    self.auth_info = {'status': 'Expired', 'message': '授权状态未知'}
                    self.is_app_locked_down = True
        else:
            # 首次运行判断逻辑 (这部分逻辑不变)
            if self._check_for_sentinels():
                #self.log("安全警告：检测到历史运行痕迹，试用期已结束。")
                self.auth_info = {'status': 'Expired', 'message': '授权已过期 (Tampered)'}
                self.is_app_locked_down = True
            else:
                #self.log("未找到有效授权和历史痕迹，初始化3天试用期...")
                trial_start_date = today_str
                trial_signature = self._generate_signature('Trial', trial_start_date)
                
                self._save_to_registry('RegistrationStatus', 'Trial')
                self._save_to_registry('RegistrationDate', trial_start_date)
                self._save_to_registry('LicenseSignature', trial_signature)
                
                self._create_sentinels()
                
                self.auth_info = {'status': 'Trial', 'message': '未注册 - 剩余 3 天'}
                self.is_app_locked_down = False

        self.update_title_bar()

    def perform_lockdown(self):
        messagebox.showerror("授权过期", "您的软件试用期或授权已到期，功能已受限。\n请在“注册软件”页面输入有效注册码以继续使用。", parent=self.root)
        self.log("软件因授权问题被锁定。")

        for task in self.tasks:
            task['status'] = '禁用'
        self.update_task_list()
        self.save_tasks()

        self.switch_page("注册软件")

    def show_trial_nag_screen(self):
        self.root.attributes('-disabled', True)

        dialog = ttk.Toplevel(self.root)
        dialog.title("试用版提示")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)
        
        # --- ↓↓↓ 修改 1：禁用窗口的关闭按钮(X) ↓↓↓ ---
        dialog.protocol("WM_DELETE_WINDOW", lambda: None)
        # --- ↑↑↑ 修改结束 ↑↑↑ ---

        countdown_seconds = 30

        main_frame = ttk.Frame(dialog, padding=(40, 20))
        main_frame.pack(fill=BOTH, expand=True)

        title_label = ttk.Label(main_frame, text="欢迎使用 创翔多功能定时播音", font=self.font_14_bold, bootstyle="primary")
        title_label.pack(pady=(0, 10))

        info_label = ttk.Label(main_frame, text="您当前使用的是试用版\n如果觉得本软件对您有帮助，请联系我们购买永久授权！", 
                               font=self.font_11, justify='center', anchor='center')
        info_label.pack(pady=10)
        
        contact_label = ttk.Label(main_frame, text="联系QQ: 315725445  |  微信: 18603970717", font=self.font_10)
        contact_label.pack(pady=10)

        # --- ↓↓↓ 修改 2：创建一个Label来显示倒计时，而不是按钮 ↓↓↓ ---
        countdown_label = ttk.Label(main_frame, text=f"请稍候... ({countdown_seconds})", font=self.font_12_bold, bootstyle="success")
        countdown_label.pack(pady=20, fill=X)
        # --- ↑↑↑ 修改结束 ↑↑↑ ---

        def cleanup_and_close():
            if hasattr(dialog, '_countdown_job'):
                dialog.after_cancel(dialog._countdown_job)
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()

        def update_countdown(sec_left):
            if sec_left > 0:
                # --- ↓↓↓ 修改 3：更新Label的文本 ↓↓↓ ---
                countdown_label.config(text=f"请稍候... ({sec_left})")
                # --- ↑↑↑ 修改结束 ↑↑↑ ---
                dialog._countdown_job = dialog.after(1000, lambda: update_countdown(sec_left - 1))
            else:
                cleanup_and_close()

        update_countdown(countdown_seconds)
        
        self.center_window(dialog, parent=self.root)

    def update_title_bar(self):
        self.root.title(f" 创翔多功能定时播音旗舰版 ({self.auth_info['message']})")

    def create_super_admin_page(self):
        page_frame = ttk.Frame(self.page_container, padding=20)
        title_label = ttk.Label(page_frame, text="超级管理", font=self.font_14_bold, bootstyle="danger")
        title_label.pack(anchor='w', pady=(0, 10))
        desc_label = ttk.Label(page_frame, text="警告：此处的任何操作都可能导致数据丢失或配置重置，请谨慎操作。\n(此功能仅对“永久授权”用户开放)",
                               font=self.font_11, bootstyle="danger", wraplength=700)
        desc_label.pack(anchor='w', pady=(0, 20))

        btn_frame = ttk.Frame(page_frame)
        btn_frame.pack(pady=10, fill=X)

        btn_width = 20
        btn_padding = 10

        ttk.Button(btn_frame, text="备份所有设置", command=self._backup_all_settings, bootstyle="primary", width=btn_width).pack(pady=btn_padding, fill=X, ipady=5)
        ttk.Button(btn_frame, text="还原所有设置", command=self._restore_all_settings, bootstyle="success", width=btn_width).pack(pady=btn_padding, fill=X, ipady=5)
        ttk.Button(btn_frame, text="重置软件", command=self._reset_software, bootstyle="danger", width=btn_width).pack(pady=btn_padding, fill=X, ipady=5)
        ttk.Button(btn_frame, text="卸载软件", command=self._prompt_for_uninstall, bootstyle="secondary", width=btn_width).pack(pady=btn_padding, fill=X, ipady=5)

        return page_frame

    def _prompt_for_uninstall(self):
        dialog = ttk.Toplevel(self.root)
        dialog.title("卸载软件 - 身份验证")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        
        # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()
        # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

        result = [None]

        ttk.Label(dialog, text="请输入卸载密码:", font=self.font_11).pack(pady=20, padx=20)
        password_entry = ttk.Entry(dialog, show='*', font=self.font_11, width=25)
        password_entry.pack(pady=5, padx=20)
        password_entry.focus_set()

        def on_confirm():
            result[0] = password_entry.get()
            cleanup_and_destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="确定", command=on_confirm, bootstyle="primary", width=8).pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=cleanup_and_destroy, width=8).pack(side=LEFT, padx=10)
        dialog.bind('<Return>', lambda event: on_confirm())
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)
        
        self.center_window(dialog, parent=self.root)
        self.root.wait_window(dialog)

        entered_password = result[0]
        correct_password = datetime.now().strftime('%Y%m%d')[::-1]

        if entered_password == correct_password:
            self.log("卸载密码正确，准备执行卸载操作。")
            self._perform_uninstall()
        elif entered_password is not None:
            messagebox.showerror("验证失败", "密码错误！", parent=self.root)
            self.log("尝试卸载软件失败：密码错误。")

    def _perform_uninstall(self):
        if not messagebox.askyesno(
            "！！！最终警告！！！",
            "您确定要卸载本软件吗？\n\n此操作将永久删除：\n- 所有注册表信息\n- 所有配置文件 (节目单, 设置, 节假日, 待办事项)\n- 所有数据文件夹 (音频, 提示音, 文稿等)\n\n此操作【绝对无法恢复】！\n\n点击“是”将立即开始清理。",
            icon='error',
            parent=self.root
        ):
            self.log("用户取消了卸载操作。")
            return

        self.log("开始执行卸载流程...")
        self.running = False

        if WIN32_AVAILABLE:
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, REGISTRY_KEY_PATH)
                self.log(f"成功删除注册表项: {REGISTRY_KEY_PATH}")
                try:
                    winreg.DeleteKey(winreg.HKEY_CURRENT_USER, REGISTRY_PARENT_KEY_PATH)
                    self.log(f"成功删除父级注册表项: {REGISTRY_PARENT_KEY_PATH}")
                except OSError:
                    self.log("父级注册表项非空，不作删除。")
            except FileNotFoundError:
                self.log("未找到相关注册表项，跳过删除。")
            except Exception as e:
                self.log(f"删除注册表时出错: {e}")

        folders_to_delete = [PROMPT_FOLDER, AUDIO_FOLDER, BGM_FOLDER, VOICE_SCRIPT_FOLDER, SCREENSHOT_FOLDER]
        for folder in folders_to_delete:
            if os.path.isdir(folder):
                try:
                    shutil.rmtree(folder)
                    self.log(f"成功删除文件夹: {os.path.basename(folder)}")
                except Exception as e:
                    self.log(f"删除文件夹 {os.path.basename(folder)} 时出错: {e}")

        files_to_delete = [
            TASK_FILE, SETTINGS_FILE, HOLIDAY_FILE, TODO_FILE,
            SCREENSHOT_TASK_FILE, EXECUTE_TASK_FILE
        ]
        for file in files_to_delete:
            if os.path.isfile(file):
                try:
                    os.remove(file)
                    self.log(f"成功删除文件: {os.path.basename(file)}")
                except Exception as e:
                    self.log(f"删除文件 {os.path.basename(file)} 时出错: {e}")

        self.log("软件数据清理完成。")
        messagebox.showinfo("卸载完成", "软件相关的数据和配置已全部清除。\n\n请手动删除本程序（.exe文件）以完成卸载。\n\n点击“确定”后软件将退出。", parent=self.root)

        os._exit(0)

    def _backup_all_settings(self):
        self.log("开始备份所有设置...")
        try:
            backup_data = {
                'backup_date': datetime.now().isoformat(), 
                'tasks': self.tasks, 
                'holidays': self.holidays,
                'todos': self.todos, 
                'screenshot_tasks': self.screenshot_tasks,
                'execute_tasks': self.execute_tasks,
                # --- ↓↓↓ 新增代码 ↓↓↓ ---
                'print_tasks': self.print_tasks,
                'backup_tasks': self.backup_tasks,
                # --- ↑↑↑ 新增代码结束 ↑↑↑ ---
                'settings': self.settings,
                'lock_password_b64': self._load_from_registry("LockPasswordB64")
            }
            filename = filedialog.asksaveasfilename(
                title="备份所有设置到...", defaultextension=".json",
                initialfile=f"boyin_backup_{datetime.now().strftime('%Y%m%d')}.json",
                filetypes=[("JSON Backup", "*.json")], initialdir=application_path,
                parent=self.root
            )
            if filename:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(backup_data, f, ensure_ascii=False, indent=2)
                self.log(f"所有设置已成功备份到: {os.path.basename(filename)}")
                messagebox.showinfo("备份成功", f"所有设置已成功备份到:\n{filename}", parent=self.root)
        except Exception as e:
            self.log(f"备份失败: {e}"); messagebox.showerror("备份失败", f"发生错误: {e}", parent=self.root)

    def _restore_all_settings(self):
        if not messagebox.askyesno("确认操作", "您确定要还原所有设置吗？\n当前所有配置将被立即覆盖。", parent=self.root):
            return

        self.log("开始还原所有设置...")
        filename = filedialog.askopenfilename(
            title="选择要还原的备份文件",
            filetypes=[("JSON Backup", "*.json")], initialdir=application_path,
            parent=self.root
        )
        if not filename: return

        try:
            with open(filename, 'r', encoding='utf-8') as f: backup_data = json.load(f)

            required_keys = ['tasks', 'holidays', 'settings', 'lock_password_b64']
            if not all(key in backup_data for key in required_keys):
                messagebox.showerror("还原失败", "备份文件格式无效或已损坏。", parent=self.root); return

            self.tasks = backup_data['tasks']
            self.holidays = backup_data['holidays']
            self.todos = backup_data.get('todos', [])
            self.screenshot_tasks = backup_data.get('screenshot_tasks', [])
            self.execute_tasks = backup_data.get('execute_tasks', [])
            # --- ↓↓↓ 新增代码 ↓↓↓ ---
            # 使用 .get() 来安全地加载，以兼容没有这些任务的旧备份文件
            self.print_tasks = backup_data.get('print_tasks', [])
            self.backup_tasks = backup_data.get('backup_tasks', [])
            # --- ↑↑↑ 新增代码结束 ↑↑↑ ---
            self.settings = backup_data['settings']
            self.lock_password_b64 = backup_data['lock_password_b64']

            self.save_tasks()
            self.save_holidays()
            self.save_todos()
            self.save_screenshot_tasks()
            self.save_execute_tasks()
            # --- ↓↓↓ 新增代码 ↓↓↓ ---
            self.save_print_tasks()
            self.save_backup_tasks()
            # --- ↑↑↑ 新增代码结束 ↑↑↑ ---
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)

            if self.lock_password_b64:
                self._save_to_registry("LockPasswordB64", self.lock_password_b64)
            else:
                self._save_to_registry("LockPasswordB64", "")

            self.update_task_list()
            self.update_holiday_list()
            self.update_todo_list()
            self.update_screenshot_list()
            self.update_execute_list()
            # --- ↓↓↓ 新增代码 ↓↓↓ ---
            self.update_print_list()
            self.update_backup_list()
            # --- ↑↑↑ 新增代码结束 ↑↑↑ ---
            self._refresh_settings_ui()
            
            self._apply_global_font()
            messagebox.showinfo("还原成功", "所有设置已成功还原。\n软件需要重启以应用字体更改。", parent=self.root)
            self.log("所有设置已从备份文件成功还原。")

            self.root.after(100, lambda: self.switch_page("定时广播"))

        except Exception as e:
            self.log(f"还原失败: {e}"); messagebox.showerror("还原失败", f"发生错误: {e}", parent=self.root)

    def _refresh_settings_ui(self):
        if "设置" not in self.pages or not hasattr(self, 'autostart_var'):
            return
        
        self.theme_var.set(self.settings.get("app_theme", "litera")) # <--- 新增此行
        self.font_var.set(self.settings.get("app_font", "Microsoft YaHei"))
        self.autostart_var.set(self.settings.get("autostart", False))
        self.start_minimized_var.set(self.settings.get("start_minimized", False))
        self.lock_on_start_var.set(self.settings.get("lock_on_start", False))
        self.daily_shutdown_enabled_var.set(self.settings.get("daily_shutdown_enabled", False))
        self.daily_shutdown_time_var.set(self.settings.get("daily_shutdown_time", "23:00:00"))
        self.weekly_shutdown_enabled_var.set(self.settings.get("weekly_shutdown_enabled", False))
        self.weekly_shutdown_time_var.set(self.settings.get("weekly_shutdown_time", "23:30:00"))
        self.weekly_shutdown_days_var.set(self.settings.get("weekly_shutdown_days", "每周:12345"))
        self.weekly_reboot_enabled_var.set(self.settings.get("weekly_reboot_enabled", False))
        self.weekly_reboot_time_var.set(self.settings.get("weekly_reboot_time", "22:00:00"))
        self.weekly_reboot_days_var.set(self.settings.get("weekly_reboot_days", "每周:67"))

        self.time_chime_enabled_var.set(self.settings.get("time_chime_enabled", False))
        self.time_chime_voice_var.set(self.settings.get("time_chime_voice", ""))
        self.time_chime_speed_var.set(self.settings.get("time_chime_speed", "0"))
        self.time_chime_pitch_var.set(self.settings.get("time_chime_pitch", "0"))

        self.bg_image_interval_var.set(str(self.settings.get("bg_image_interval", 6)))

        if self.lock_password_b64 and WIN32_AVAILABLE:
            self.clear_password_btn.config(state=NORMAL)
        else:
            self.clear_password_btn.config(state=DISABLED)

    def _reset_software(self):
        if not messagebox.askyesno(
            "！！！最终确认！！！",
            "您真的要重置整个软件吗？\n\n此操作将：\n- 清空所有节目单 (但保留音频文件)\n- 清空所有高级功能任务\n- 清空所有节假日和待办事项\n- 清除锁定密码\n- 重置所有系统设置 (包括字体)\n\n此操作【无法恢复】！软件将在重置后提示您重启。",
            parent=self.root
        ): return

        self.log("开始执行软件重置...")
        try:
            original_askyesno = messagebox.askyesno
            messagebox.askyesno = lambda title, message, parent: True
            self.clear_all_tasks(delete_associated_files=False)
            self.clear_all_screenshot_tasks()
            self.clear_all_execute_tasks()
            self.clear_all_holidays()
            self.clear_all_todos()
            self.clear_all_print_tasks()
            self.clear_all_backup_tasks()
            self.clear_all_dynamic_voice_tasks()
            messagebox.askyesno = original_askyesno

            self._save_to_registry("LockPasswordB64", "")

            if os.path.exists(CHIME_FOLDER):
                shutil.rmtree(CHIME_FOLDER)
                self.log("已删除整点报时缓存文件。")

            default_settings = {
                "app_font": "Microsoft YaHei",
                "autostart": False, "start_minimized": False, "lock_on_start": False,
                "daily_shutdown_enabled": False, "daily_shutdown_time": "23:00:00",
                "weekly_shutdown_enabled": False, "weekly_shutdown_days": "每周:12345", "weekly_shutdown_time": "23:30:00",
                "weekly_reboot_enabled": False, "weekly_reboot_days": "每周:67", "weekly_reboot_time": "22:00:00",
                "last_power_action_date": "",
                "time_chime_enabled": False, "time_chime_voice": "",
                "time_chime_speed": "0", "time_chime_pitch": "0",
                "bg_image_interval": 6
            }
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(default_settings, f, ensure_ascii=False, indent=2)

            self.log("软件已成功重置。软件需要重启。")
            messagebox.showinfo("重置成功", "软件已恢复到初始状态。\n\n请点击“确定”后手动关闭并重新启动软件。", parent=self.root)
        except Exception as e:
            self.log(f"重置失败: {e}"); messagebox.showerror("重置失败", f"发生错误: {e}", parent=self.root)

    def create_scheduled_broadcast_page(self):
        page_frame = self.pages["定时广播"]

        top_frame = ttk.Frame(page_frame, padding=(10, 10))
        top_frame.pack(side=TOP, fill=X)
        
        title_label = ttk.Label(top_frame, text="定时广播", font=self.font_14_bold)
        title_label.pack(side=LEFT)

        add_btn = ttk.Button(top_frame, text="添加节目", command=self.add_task, bootstyle="primary")
        add_btn.pack(side=LEFT, padx=10)

        top_right_container = ttk.Frame(top_frame)
        top_right_container.pack(side=RIGHT)

        button_row_1 = ttk.Frame(top_right_container)
        button_row_1.pack(fill=X, anchor='e')

        button_row_2 = ttk.Frame(top_right_container)
        button_row_2.pack(fill=X, anchor='e', pady=(5, 0))

        batch_buttons_row1 = [
            ("全部启用", self.enable_all_tasks, 'success'),
            ("全部禁用", self.disable_all_tasks, 'warning'),
            ("禁音频节目", lambda: self._set_tasks_status_by_type('audio', '禁用'), 'warning-outline'),
            ("禁语音节目", lambda: self._set_tasks_status_by_type('voice', '禁用'), 'warning-outline'),
            ("禁视频节目", lambda: self._set_tasks_status_by_type('video', '禁用'), 'warning-outline'),
        ]
        for text, cmd, style in batch_buttons_row1:
            btn = ttk.Button(button_row_1, text=text, command=cmd, bootstyle=style)
            btn.pack(side=LEFT, padx=3)

        batch_buttons_row2 = [
            ("统一音量", self.set_uniform_volume, 'info'),
            ("清空节目", self.clear_all_tasks, 'danger'),
            ("导入节目单", self.import_tasks, 'info-outline'),
            ("导出节目单", self.export_tasks, 'info-outline'),
        ]
        for text, cmd, style in batch_buttons_row2:
            btn = ttk.Button(button_row_2, text=text, command=cmd, bootstyle=style)
            btn.pack(side=LEFT, padx=3)
            
        self.pin_button = ttk.Button(button_row_2, text="置顶", command=self.toggle_pin_state, bootstyle="info-outline")
        self.pin_button.pack(side=LEFT, padx=3)
        
        self.lock_button = ttk.Button(button_row_2, text="锁定", command=self.toggle_lock_state, bootstyle='danger')
        self.lock_button.pack(side=LEFT, padx=3)
        if not WIN32_AVAILABLE:
            self.lock_button.config(state=DISABLED, text="锁定(Win)")

        stats_frame = ttk.Frame(page_frame, padding=(10, 5))
        stats_frame.pack(side=TOP, fill=X)
        
        # “节目单”标签，靠左显示
        self.stats_label = ttk.Label(stats_frame, text="节目单：0", font=self.font_11, bootstyle="secondary")
        self.stats_label.pack(side=LEFT)

# 新增的可点击天气标签，靠右显示
        self.main_weather_label = ttk.Label(stats_frame, text="天气: 正在获取...", font=self.font_11, bootstyle="info", cursor="hand2")
        self.main_weather_label.pack(side=RIGHT, padx=10)
        self.main_weather_label.bind("<Button-1>", self.on_weather_label_click)

        log_frame = ttk.LabelFrame(page_frame, text="", padding=(10, 5))
        log_frame.pack(side=BOTTOM, fill=X, padx=10, pady=5)

        playing_frame = ttk.LabelFrame(page_frame, text="正在播：", padding=(10, 5))
        playing_frame.pack(side=BOTTOM, fill=X, padx=10, pady=5)
        
        table_frame = ttk.Frame(page_frame, padding=(10, 5))
        table_frame.pack(side=TOP, fill=BOTH, expand=True)

        columns = ('节目名称', '状态', '开始时间', '模式', '文件或内容', '音量', '周几/几号', '日期范围')
        self.task_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=12, selectmode='extended', bootstyle="primary")

        self.task_tree.heading('节目名称', text='节目名称')
        self.task_tree.column('节目名称', width=200, anchor='w')
        self.task_tree.heading('状态', text='状态'); self.task_tree.column('状态', width=70, anchor='center', stretch=NO)
        self.task_tree.heading('开始时间', text='开始时间'); self.task_tree.column('开始时间', width=100, anchor='center', stretch=NO)
        self.task_tree.heading('模式', text='模式'); self.task_tree.column('模式', width=70, anchor='center', stretch=NO)
        self.task_tree.heading('文件或内容', text='文件或内容'); self.task_tree.column('文件或内容', width=300, anchor='w')
        self.task_tree.heading('音量', text='音量'); self.task_tree.column('音量', width=70, anchor='center', stretch=NO)
        self.task_tree.heading('周几/几号', text='周几/几号'); self.task_tree.column('周几/几号', width=100, anchor='center')
        self.task_tree.heading('日期范围', text='日期范围'); self.task_tree.column('日期范围', width=120, anchor='center')

        self.task_tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar = ttk.Scrollbar(table_frame, orient=VERTICAL, command=self.task_tree.yview, bootstyle="round")
        scrollbar.pack(side=RIGHT, fill=Y)
        self.task_tree.configure(yscrollcommand=scrollbar.set)

        self.task_tree.bind("<Button-3>", self.show_context_menu)
        self.task_tree.bind("<Double-1>", self.on_double_click_edit)
        self._enable_drag_selection(self.task_tree)

        self.playing_label = ttk.Label(playing_frame, text="等待播放...", font=self.font_11,
                                       anchor=W, justify=LEFT, padding=5, bootstyle="warning")
        self.playing_label.pack(fill=X, expand=True, ipady=4)
        self.update_playing_text("等待播放...")

        log_header_frame = ttk.Frame(log_frame)
        log_header_frame.pack(fill=X)
        log_label = ttk.Label(log_header_frame, text="日志：", font=self.font_11_bold)
        log_label.pack(side=LEFT)
        self.clear_log_btn = ttk.Button(log_header_frame, text="清除日志", command=self.clear_log,
                                        bootstyle="secondary-outline")
        self.clear_log_btn.pack(side=LEFT, padx=10)

        self.log_text = ScrolledText(log_frame, height=6, font=self.font_11,
                                                  wrap=WORD, state='disabled')
        self.log_text.pack(fill=BOTH, expand=True)

#第3部分
    def create_settings_page(self):
        settings_frame = ttk.Frame(self.page_container, padding=20)

        title_label = ttk.Label(settings_frame, text="系统设置", font=self.font_14_bold, bootstyle="primary")
        title_label.pack(anchor=W, pady=(0, 10))

        # --- 通用设置 ---
        general_frame = ttk.LabelFrame(settings_frame, text="通用设置", padding=(15, 10))
        general_frame.pack(fill=X, pady=10)

        self.autostart_var = ttk.BooleanVar()
        self.start_minimized_var = ttk.BooleanVar()
        self.lock_on_start_var = ttk.BooleanVar()
        self.bg_image_interval_var = ttk.StringVar()

        ttk.Checkbutton(general_frame, text="登录windows后自动启动", variable=self.autostart_var, bootstyle="round-toggle", command=self._handle_autostart_setting).pack(fill=X, pady=5)
        ttk.Checkbutton(general_frame, text="启动后最小化到系统托盘", variable=self.start_minimized_var, bootstyle="round-toggle", command=self.save_settings).pack(fill=X, pady=5)

        lock_on_start_frame = ttk.Frame(general_frame)
        lock_on_start_frame.pack(fill=X, pady=5)
        self.lock_on_start_cb = ttk.Checkbutton(lock_on_start_frame, text="启动软件后立即锁定", variable=self.lock_on_start_var, bootstyle="round-toggle", command=self._handle_lock_on_start_toggle)
        self.lock_on_start_cb.pack(side=LEFT)
        if not WIN32_AVAILABLE:
            self.lock_on_start_cb.config(state=DISABLED)
        ttk.Label(lock_on_start_frame, text="(请先在主界面设置锁定密码)", font=self.font_9, bootstyle="secondary").pack(side=LEFT, padx=10, anchor='w')
        self.clear_password_btn = ttk.Button(lock_on_start_frame, text="清除锁定密码", command=self.clear_lock_password, bootstyle="danger-link")
        self.clear_password_btn.pack(side=LEFT, padx=10)
        
        bg_interval_frame = ttk.Frame(general_frame)
        bg_interval_frame.pack(fill=X, pady=8)
        ttk.Label(bg_interval_frame, text="背景图片切换间隔:").pack(side=LEFT)
        interval_entry = ttk.Entry(bg_interval_frame, textvariable=self.bg_image_interval_var, font=self.font_11, width=5)
        interval_entry.pack(side=LEFT, padx=5)
        ttk.Label(bg_interval_frame, text="秒 (范围: 5-60)", font=self.font_10, bootstyle="secondary").pack(side=LEFT)
        ttk.Button(bg_interval_frame, text="确定", command=self._validate_bg_interval, bootstyle="primary-outline").pack(side=LEFT, padx=10)
        self.cancel_bg_images_btn = ttk.Button(bg_interval_frame, text="取消所有节目背景图片", command=self._cancel_all_background_images, bootstyle="info-outline")
        self.cancel_bg_images_btn.pack(side=LEFT, padx=5)
        self.restore_video_speed_btn = ttk.Button(bg_interval_frame, text="恢复所有视频节目播放速度", command=self._restore_all_video_speeds, bootstyle="info-outline")
        self.restore_video_speed_btn.pack(side=LEFT, padx=5)

        # --- 外观设置 (字体和主题) ---
        appearance_frame = ttk.Frame(general_frame)
        appearance_frame.pack(fill=X, pady=10)

        # 字体设置 (左侧)
        ttk.Label(appearance_frame, text="软件字体:").pack(side=LEFT)
        try:
            available_fonts = sorted(list(font.families()))
        except:
            available_fonts = ["Microsoft YaHei"]
        self.font_var = ttk.StringVar()
        font_combo = ttk.Combobox(appearance_frame, textvariable=self.font_var, values=available_fonts, font=self.font_10, width=20, state='readonly')
        font_combo.pack(side=LEFT, padx=(10, 5))
        font_combo.bind("<<ComboboxSelected>>", self._on_font_selected)
        restore_font_btn = ttk.Button(appearance_frame, text="恢复默认", command=self._restore_default_font, bootstyle="secondary-outline")
        restore_font_btn.pack(side=LEFT, padx=5)

        # 主题设置 (右侧)
        style = ttk.Style.get_instance()
        available_themes = sorted(style.theme_names())
        self.theme_var = ttk.StringVar()
        theme_combo = ttk.Combobox(appearance_frame, textvariable=self.theme_var, values=available_themes, font=self.font_10, width=20, state='readonly')
        theme_combo.pack(side=RIGHT, padx=10)
        theme_combo.bind("<<ComboboxSelected>>", self._on_theme_selected)
        ttk.Label(appearance_frame, text="软件主题:").pack(side=RIGHT)

        # --- 整点报时 ---
        time_chime_frame = ttk.LabelFrame(settings_frame, text="整点报时", padding=(15, 10))
        time_chime_frame.pack(fill=X, pady=10)
        self.time_chime_enabled_var = ttk.BooleanVar()
        self.time_chime_voice_var = ttk.StringVar()
        self.time_chime_speed_var = ttk.StringVar()
        self.time_chime_pitch_var = ttk.StringVar()
        chime_control_frame = ttk.Frame(time_chime_frame)
        chime_control_frame.pack(fill=X, pady=5)
        chime_control_frame.columnconfigure(1, weight=1)
        ttk.Checkbutton(chime_control_frame, text="启用整点报时功能", variable=self.time_chime_enabled_var, bootstyle="round-toggle", command=self._handle_time_chime_toggle).pack(side=LEFT)
        available_voices = self.get_available_voices()
        self.chime_voice_combo = ttk.Combobox(chime_control_frame, textvariable=self.time_chime_voice_var, values=available_voices, font=self.font_10, state='readonly')
        self.chime_voice_combo.pack(side=LEFT, padx=10, fill=X, expand=True)
        self.chime_voice_combo.bind("<<ComboboxSelected>>", lambda e: self._on_chime_params_changed(is_voice_change=True))
        params_frame = ttk.Frame(chime_control_frame)
        params_frame.pack(side=LEFT, padx=10)
        ttk.Label(params_frame, text="语速(-10~10):", font=self.font_10).pack(side=LEFT)
        speed_entry = ttk.Entry(params_frame, textvariable=self.time_chime_speed_var, font=self.font_10, width=5)
        speed_entry.pack(side=LEFT, padx=(0, 10))
        ttk.Label(params_frame, text="音调(-10~10):", font=self.font_10).pack(side=LEFT)
        pitch_entry = ttk.Entry(params_frame, textvariable=self.time_chime_pitch_var, font=self.font_10, width=5)
        pitch_entry.pack(side=LEFT)
        speed_entry.bind("<FocusOut>", self._on_chime_params_changed)
        pitch_entry.bind("<FocusOut>", self._on_chime_params_changed)

        # --- 电源管理 ---
        power_frame = ttk.LabelFrame(settings_frame, text="电源管理", padding=(15, 10))
        power_frame.pack(fill=X, pady=10)
        self.daily_shutdown_enabled_var = ttk.BooleanVar()
        self.daily_shutdown_time_var = ttk.StringVar()
        self.weekly_shutdown_enabled_var = ttk.BooleanVar()
        self.weekly_shutdown_time_var = ttk.StringVar()
        self.weekly_shutdown_days_var = ttk.StringVar()
        self.weekly_reboot_enabled_var = ttk.BooleanVar()
        self.weekly_reboot_time_var = ttk.StringVar()
        self.weekly_reboot_days_var = ttk.StringVar()

        # --- ↓↓↓ 核心修改部分：为 command 添加验证逻辑 ↓↓↓ ---
        
        def validate_and_save_settings(var_to_check=None, related_vars=None, error_msg=""):
            """通用验证和保存函数"""
            # 只有在用户尝试“启用”时才进行检查
            if var_to_check and var_to_check.get():
                for r_var in related_vars:
                    # 检查关联的输入框内容是否为空或仅包含前缀
                    val = r_var.get().strip()
                    if not val or val == "每周:":
                        messagebox.showerror("设置无效", error_msg, parent=self.root)
                        # 将开关拨回“关闭”状态
                        var_to_check.set(False) 
                        return # 终止保存
            
            # 如果验证通过或用户是“关闭”功能，则正常保存
            self.save_settings()

        daily_frame = ttk.Frame(power_frame)
        daily_frame.pack(fill=X, pady=4)
        daily_frame.columnconfigure(1, weight=1)
        # 每日关机不需要特殊验证，直接保存
        ttk.Checkbutton(daily_frame, text="每天关机    ", variable=self.daily_shutdown_enabled_var, bootstyle="round-toggle", command=self.save_settings).grid(row=0, column=0, sticky='w')
        daily_time_entry = ttk.Entry(daily_frame, textvariable=self.daily_shutdown_time_var, font=self.font_11)
        daily_time_entry.grid(row=0, column=1, sticky='we', padx=5)
        self._bind_mousewheel_to_entry(daily_time_entry, self._handle_time_scroll)
        ttk.Button(daily_frame, text="设置", bootstyle="primary-outline", command=lambda: self.show_single_time_dialog(self.daily_shutdown_time_var)).grid(row=0, column=2, sticky='e', padx=5)
        
        weekly_frame = ttk.Frame(power_frame)
        weekly_frame.pack(fill=X, pady=4)
        weekly_frame.columnconfigure(1, weight=1)
        # 每周关机：在保存前进行验证
        ttk.Checkbutton(weekly_frame, text="每周关机    ", variable=self.weekly_shutdown_enabled_var, bootstyle="round-toggle", 
                        command=lambda: validate_and_save_settings(
                            self.weekly_shutdown_enabled_var, 
                            [self.weekly_shutdown_days_var, self.weekly_shutdown_time_var],
                            "无法启用“每周关机”，因为周几或时间未设置。"
                        )).grid(row=0, column=0, sticky='w')
        weekly_days_entry = ttk.Entry(weekly_frame, textvariable=self.weekly_shutdown_days_var, font=self.font_11)
        weekly_days_entry.grid(row=0, column=1, sticky='we', padx=5)
        weekly_shutdown_time_entry = ttk.Entry(weekly_frame, textvariable=self.weekly_shutdown_time_var, font=self.font_11, width=15)
        weekly_shutdown_time_entry.grid(row=0, column=2, sticky='we', padx=5)
        self._bind_mousewheel_to_entry(weekly_shutdown_time_entry, self._handle_time_scroll)
        ttk.Button(weekly_frame, text="设置", bootstyle="primary-outline", command=lambda: self.show_power_week_time_dialog("设置每周关机", self.weekly_shutdown_days_var, self.weekly_shutdown_time_var)).grid(row=0, column=3, sticky='e', padx=5)
        
        reboot_frame = ttk.Frame(power_frame)
        reboot_frame.pack(fill=X, pady=4)
        reboot_frame.columnconfigure(1, weight=1)
        # 每周重启：在保存前进行验证
        ttk.Checkbutton(reboot_frame, text="每周重启    ", variable=self.weekly_reboot_enabled_var, bootstyle="round-toggle", 
                        command=lambda: validate_and_save_settings(
                            self.weekly_reboot_enabled_var,
                            [self.weekly_reboot_days_var, self.weekly_reboot_time_var],
                            "无法启用“每周重启”，因为周几或时间未设置。"
                        )).grid(row=0, column=0, sticky='w')
        ttk.Entry(reboot_frame, textvariable=self.weekly_reboot_days_var, font=self.font_11).grid(row=0, column=1, sticky='we', padx=5)
        weekly_reboot_time_entry = ttk.Entry(reboot_frame, textvariable=self.weekly_reboot_time_var, font=self.font_11, width=15)
        weekly_reboot_time_entry.grid(row=0, column=2, sticky='we', padx=5)
        self._bind_mousewheel_to_entry(weekly_reboot_time_entry, self._handle_time_scroll)
        ttk.Button(reboot_frame, text="设置", bootstyle="primary-outline", command=lambda: self.show_power_week_time_dialog("设置每周重启", self.weekly_reboot_days_var, self.weekly_reboot_time_var)).grid(row=0, column=3, sticky='e', padx=5)
        # --- ↑↑↑ 核心修改结束 ↑↑↑ ---

        return settings_frame

    def _restore_all_video_speeds(self):
        if not self.tasks:
            messagebox.showinfo("提示", "当前没有节目，无需操作。", parent=self.root)
            return

        count = 0
        for task in self.tasks:
            if task.get('type') == 'video':
                if task.get('playback_rate') != '1.0x (正常)':
                    task['playback_rate'] = '1.0x (正常)'
                    count += 1
        
        if count > 0:
            self.save_tasks()
            self.log(f"已成功将 {count} 个视频节目的播放速度恢复为1.0x。")
            messagebox.showinfo("操作成功", f"已成功将 {count} 个视频节目的播放速度恢复为默认值(1.0x)。", parent=self.root)
        else:
            messagebox.showinfo("提示", "所有视频节目已经是默认播放速度，无需恢复。", parent=self.root)

    def _on_font_selected(self, event):
        new_font = self.font_var.get()
        if new_font and new_font != self.settings.get("app_font", "Microsoft YaHei"):
            self.settings["app_font"] = new_font
            self.save_settings()
            self.log(f"字体已更改为 '{new_font}'。")
            self._apply_global_font()
            messagebox.showinfo("设置已保存", "字体设置已保存。\n请重启软件以使新字体完全生效。", parent=self.root)

    def _on_theme_selected(self, event=None):
        """当用户从下拉框选择一个新主题时调用"""
        new_theme = self.theme_var.get()
        if new_theme:
            try:
                style = ttk.Style.get_instance()
                style.theme_use(new_theme)
                self.log(f"软件主题已切换为: {new_theme}")
                # 保存新主题到设置
                self.settings['app_theme'] = new_theme
                self.save_settings()
            except tk.TclError:
                messagebox.showerror("错误", f"无法应用主题 '{new_theme}'。", parent=self.root)
                self.log(f"错误：切换主题 '{new_theme}' 失败。")
                # 恢复到上一个有效主题
                self.theme_var.set(style.theme_use())

    def _restore_default_font(self):
        default_font = "Microsoft YaHei"
        if self.settings.get("app_font") != default_font:
            self.settings["app_font"] = default_font
            self.save_settings()
            self.font_var.set(default_font)
            self.log("字体已恢复为默认。")
            self._apply_global_font()
            messagebox.showinfo("设置已保存", "字体已恢复为默认设置。\n请重启软件以生效。", parent=self.root)
        else:
            messagebox.showinfo("提示", "当前已是默认字体，无需恢复。", parent=self.root)

    def _validate_bg_interval(self, event=None):
        try:
            value = int(self.bg_image_interval_var.get())
            if not (5 <= value <= 60):
                raise ValueError("超出范围")
            self.settings['bg_image_interval'] = value
            self.save_settings()
            self.log(f"背景图片切换间隔已更新为 {value} 秒。")
            messagebox.showinfo("保存成功", f"背景图片切换间隔已设置为 {value} 秒。", parent=self.root)
        except (ValueError, TypeError):
            last_saved_value = str(self.settings.get("bg_image_interval", 6))
            messagebox.showerror("输入无效", "请输入一个介于 5 和 60 之间的整数。", parent=self.root)
            self.bg_image_interval_var.set(last_saved_value)

    def _cancel_all_background_images(self):
        if not self.tasks:
            messagebox.showinfo("提示", "当前没有节目，无需操作。", parent=self.root)
            return

        if messagebox.askyesno("确认操作", "您确定要取消所有节目中已设置的背景图片吗？\n此操作将取消所有任务的背景图片勾选。", parent=self.root):
            count = 0
            for task in self.tasks:
                if task.get('bg_image_enabled'):
                    task['bg_image_enabled'] = 0
                    count += 1

            if count > 0:
                self.save_tasks()
                self.log(f"已成功取消 {count} 个节目的背景图片设置。")
                messagebox.showinfo("操作成功", f"已成功取消 {count} 个节目的背景图片设置。", parent=self.root)
            else:
                messagebox.showinfo("提示", "没有节目设置了背景图片，无需操作。", parent=self.root)

    def _on_chime_params_changed(self, event=None, is_voice_change=False):
        current_voice = self.time_chime_voice_var.get()
        current_speed = self.time_chime_speed_var.get()
        current_pitch = self.time_chime_pitch_var.get()

        saved_voice = self.settings.get("time_chime_voice", "")
        saved_speed = self.settings.get("time_chime_speed", "0")
        saved_pitch = self.settings.get("time_chime_pitch", "0")

        params_changed = (current_voice != saved_voice or
                          current_speed != saved_speed or
                          current_pitch != saved_pitch)

        if self.time_chime_enabled_var.get() and params_changed:
            self.save_settings()
            if messagebox.askyesno("应用更改", "您更改了报时参数，需要重新生成全部24个报时文件。\n是否立即开始？", parent=self.root):
                self._handle_time_chime_toggle(force_regenerate=True)
            else:
                if is_voice_change: self.time_chime_voice_var.set(saved_voice)
                self.time_chime_speed_var.set(saved_speed)
                self.time_chime_pitch_var.set(saved_pitch)
        else:
            self.save_settings()

    def _handle_time_chime_toggle(self, force_regenerate=False):
        is_enabled = self.time_chime_enabled_var.get()

        if is_enabled or force_regenerate:
            selected_voice = self.time_chime_voice_var.get()
            if not selected_voice:
                messagebox.showwarning("操作失败", "请先从下拉列表中选择一个播音员。", parent=self.root)
                if not force_regenerate: self.time_chime_enabled_var.set(False)
                return

            self.save_settings()
            self.log("准备启用/更新整点报时功能，开始生成语音文件...")

            progress_dialog = ttk.Toplevel(self.root)
            progress_dialog.title("请稍候")
            progress_dialog.resizable(False, False)
            progress_dialog.transient(self.root)

            # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
            progress_dialog.attributes('-topmost', True)
            self.root.attributes('-disabled', True)
            
            def cleanup_and_destroy():
                self.root.attributes('-disabled', False)
                progress_dialog.destroy()
                self.root.focus_force()
            # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

            progress_dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)

            ttk.Label(progress_dialog, text="正在生成整点报时文件 (0/24)...", font=self.font_11).pack(pady=10, padx=20)
            progress_label = ttk.Label(progress_dialog, text="", font=self.font_10)
            progress_label.pack(pady=5, padx=20)
            
            self.center_window(progress_dialog, parent=self.root)

            threading.Thread(target=self._generate_chime_files_worker,
                             args=(selected_voice, progress_dialog, progress_label), daemon=True).start()

        elif not is_enabled and not force_regenerate:
            if messagebox.askyesno("确认操作", "您确定要禁用整点报时功能吗？\n这将删除所有已生成的报时音频文件。", parent=self.root):
                self.save_settings()
                threading.Thread(target=self._delete_chime_files_worker, daemon=True).start()
            else:
                self.time_chime_enabled_var.set(True)

    def _get_time_period_string(self, hour):
        if 0 <= hour < 6: return "凌晨"
        elif 6 <= hour < 9: return "早上"
        elif 9 <= hour < 12: return "上午"
        elif 12 <= hour < 14: return "中午"
        elif 14 <= hour < 18: return "下午"
        else: return "晚上"

    def _generate_chime_files_worker(self, voice, progress_dialog, progress_label):
        if not os.path.exists(CHIME_FOLDER):
            os.makedirs(CHIME_FOLDER)

        success = True
        try:
            for hour in range(24):
                period = self._get_time_period_string(hour)
                display_hour = hour
                if period == "下午" and hour > 12: display_hour -= 12
                elif period == "晚上" and hour > 12: display_hour -= 12

                text = f"现在时刻,北京时间{period}{display_hour}点整"
                output_path = os.path.join(CHIME_FOLDER, f"{hour:02d}.wav")

                progress_text = f"正在生成：{hour:02d}.wav ({hour + 1}/24)"
                self.root.after(0, lambda p=progress_text: progress_label.config(text=p))

                voice_params = {
                    'voice': voice,
                    'speed': self.settings.get("time_chime_speed", "0"),
                    'pitch': self.settings.get("time_chime_pitch", "0"),
                    'volume': '100'
                }
                if not self._synthesize_text_to_wav(text, voice_params, output_path):
                    raise Exception(f"生成 {hour:02d}.wav 失败")
        except Exception as e:
            success = False
            self.log(f"生成整点报时文件时出错: {e}")
            self.root.after(0, messagebox.showerror, "错误", f"生成报时文件失败：{e}", parent=self.root)
        finally:
            self.root.after(0, progress_dialog.destroy)
            self.root.after(1, lambda: self.root.attributes('-disabled', False))
            if success:
                self.log("全部整点报时文件生成完毕。")
                if self.time_chime_enabled_var.get():
                     self.root.after(0, messagebox.showinfo, "成功", "整点报时功能已启用/更新！", parent=self.root)
            else:
                self.log("整点报时功能启用失败。")
                self.settings['time_chime_enabled'] = False
                self.root.after(0, self.time_chime_enabled_var.set, False)
                self.save_settings()

    def _delete_chime_files_worker(self):
        self.log("正在禁用整点报时功能，开始删除缓存文件...")
        try:
            if os.path.exists(CHIME_FOLDER):
                shutil.rmtree(CHIME_FOLDER)
                self.log("整点报时缓存文件已成功删除。")
            else:
                self.log("未找到整点报时缓存文件夹，无需删除。")
        except Exception as e:
            self.log(f"删除整点报时文件失败: {e}")
            self.root.after(0, messagebox.showerror, "错误", f"删除报时文件失败：{e}", parent=self.root)

    def toggle_pin_state(self):
        self.is_window_pinned = not self.is_window_pinned
        
        if self.is_window_pinned:
            self.root.attributes('-topmost', True)
            self.pin_button.config(text="取消置顶", bootstyle="info")
            self.log("窗口已置顶显示。")
        else:
            self.root.attributes('-topmost', False)
            self.pin_button.config(text="置顶", bootstyle="info-outline")
            self.log("窗口已取消置顶。")

    def toggle_lock_state(self):
        if self.is_locked:
            self._prompt_for_password_unlock()
        else:
            if not self.lock_password_b64:
                self._prompt_for_password_set()
            else:
                self._apply_lock()

    def _apply_lock(self):
        self.is_locked = True
        self.lock_button.config(text="解锁", bootstyle='success')
        self._set_ui_lock_state(DISABLED)
        self.statusbar_unlock_button.pack(side=RIGHT, padx=5)
        self.log("界面已锁定。")

    def _apply_unlock(self):
        self.is_locked = False
        self.lock_button.config(text="锁定", bootstyle='danger')
        self._set_ui_lock_state(NORMAL)
        self.statusbar_unlock_button.pack_forget()
        self.log("界面已解锁。")

    def perform_initial_lock(self):
        self.log("根据设置，软件启动时自动锁定。")
        self._apply_lock()

    def _prompt_for_password_set(self):
        dialog = ttk.Toplevel(self.root)
        dialog.title("首次锁定，请设置密码")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()
        # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

        ttk.Label(dialog, text="请设置一个锁定密码 (最多6位)", font=self.font_11).pack(pady=10, padx=20)

        ttk.Label(dialog, text="输入密码:", font=self.font_11).pack(pady=(5,0))
        pass_entry1 = ttk.Entry(dialog, show='*', width=25, font=self.font_11)
        pass_entry1.pack(padx=20)

        ttk.Label(dialog, text="确认密码:", font=self.font_11).pack(pady=(10,0))
        pass_entry2 = ttk.Entry(dialog, show='*', width=25, font=self.font_11)
        pass_entry2.pack(padx=20)

        def confirm():
            p1 = pass_entry1.get()
            p2 = pass_entry2.get()
            if not p1: messagebox.showerror("错误", "密码不能为空。", parent=dialog); return
            if len(p1) > 6: messagebox.showerror("错误", "密码不能超过6位。", parent=dialog); return
            if p1 != p2: messagebox.showerror("错误", "两次输入的密码不一致。", parent=dialog); return

            encoded_pass = base64.b64encode(p1.encode('utf-8')).decode('utf-8')
            if self._save_to_registry("LockPasswordB64", encoded_pass):
                self.lock_password_b64 = encoded_pass
                if "设置" in self.pages and hasattr(self, 'clear_password_btn'):
                    self.clear_password_btn.config(state=NORMAL)
                messagebox.showinfo("成功", "密码设置成功，界面即将锁定。", parent=dialog)
                cleanup_and_destroy()
                self._apply_lock()
            else:
                messagebox.showerror("功能受限", "无法保存密码。\n此功能仅在Windows系统上支持且需要pywin32库。", parent=dialog)

        btn_frame = ttk.Frame(dialog); btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10)
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)
        
        self.center_window(dialog, parent=self.root)

    def _prompt_for_password_unlock(self):
        dialog = ttk.Toplevel(self.root)
        dialog.title("解锁界面")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()
        # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

        ttk.Label(dialog, text="请输入密码以解锁", font=self.font_11).pack(pady=10, padx=20)

        pass_entry = ttk.Entry(dialog, show='*', width=25, font=self.font_11)
        pass_entry.pack(pady=5, padx=20)
        pass_entry.focus_set()

        def is_password_correct():
            entered_pass = pass_entry.get()
            encoded_entered_pass = base64.b64encode(entered_pass.encode('utf-8')).decode('utf-8')
            return encoded_entered_pass == self.lock_password_b64

        def confirm():
            if is_password_correct():
                cleanup_and_destroy()
                self._apply_unlock()
            else:
                messagebox.showerror("错误", "密码不正确！", parent=dialog)

        def clear_password_action():
            if not is_password_correct():
                messagebox.showerror("错误", "密码不正确！无法清除。", parent=dialog)
                return

            if messagebox.askyesno("确认操作", "您确定要清除锁定密码吗？\n此操作不可恢复。", parent=dialog):
                self._perform_password_clear_logic()
                cleanup_and_destroy()
                self.root.after(50, self._apply_unlock)
                self.root.after(100, lambda: messagebox.showinfo("成功", "锁定密码已成功清除。", parent=self.root))

        btn_frame = ttk.Frame(dialog); btn_frame.pack(pady=20, padx=10, fill=X, expand=True)
        btn_frame.columnconfigure((0, 1, 2), weight=1)
        ttk.Button(btn_frame, text="确定", command=confirm, bootstyle="primary").grid(row=0, column=0, padx=5, sticky='ew')
        ttk.Button(btn_frame, text="清除密码", command=clear_password_action, bootstyle="warning").grid(row=0, column=1, padx=5, sticky='ew')
        ttk.Button(btn_frame, text="取消", command=cleanup_and_destroy).grid(row=0, column=2, padx=5, sticky='ew')
        dialog.bind('<Return>', lambda event: confirm())
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)
        
        self.center_window(dialog, parent=self.root)

    def _perform_password_clear_logic(self):
        if self._save_to_registry("LockPasswordB64", ""):
            self.lock_password_b64 = ""
            self.settings["lock_on_start"] = False

            if hasattr(self, 'lock_on_start_var'):
                self.lock_on_start_var.set(False)

            self.save_settings()

            if hasattr(self, 'clear_password_btn'):
                self.clear_password_btn.config(state=DISABLED)
            self.log("锁定密码已清除。")

    def clear_lock_password(self):
        if messagebox.askyesno("确认操作", "您确定要清除锁定密码吗？\n此操作不可恢复。", parent=self.root):
            self._perform_password_clear_logic()
            messagebox.showinfo("成功", "锁定密码已成功清除。", parent=self.root)

#第4部分
    def _handle_lock_on_start_toggle(self):
        if not self.lock_password_b64:
            if self.lock_on_start_var.get():
                messagebox.showwarning("无法启用", "您还未设置锁定密码。\n\n请返回“定时广播”页面，点击“锁定”按钮来首次设置密码。", parent=self.root)
                self.root.after(50, lambda: self.lock_on_start_var.set(False))
        else:
            self.save_settings()

    def _set_ui_lock_state(self, state):
        for title, btn in self.nav_buttons.items():
            if title in ["超级管理", "注册软件"]:
                continue
            try:
                btn.config(state=state)
            except tk.TclError:
                pass

        for page_name, page_frame in self.pages.items():
            if page_frame and page_frame.winfo_exists():
                if page_name in ["超级管理", "注册软件"]:
                    continue
                self._set_widget_state_recursively(page_frame, state)

    def _set_widget_state_recursively(self, parent_widget, state):
        special_widgets = (ttk.Scrollbar, )
        
        for child in parent_widget.winfo_children():
            if child == self.lock_button:
                continue

            if isinstance(child, special_widgets):
                continue
                
            try:
                child.config(state=state)
            except tk.TclError:
                pass

            if child.winfo_children():
                self._set_widget_state_recursively(child, state)

    def clear_log(self):
        if messagebox.askyesno("确认操作", "您确定要清空所有日志记录吗？\n此操作不可恢复。", parent=self.root):
            log_widget = self.log_text.text
            log_widget.config(state='normal')
            log_widget.delete('1.0', END)
            log_widget.config(state='disabled')
            self.log("日志已清空。")

    def on_double_click_edit(self, event):
        if self.is_locked: return
        if self.task_tree.identify_row(event.y):
            self.edit_task()

    def show_context_menu(self, event):
        if self.is_locked: return
        iid = self.task_tree.identify_row(event.y)
        context_menu = tk.Menu(self.root, tearoff=0, font=self.font_11)

        if iid:
            if iid not in self.task_tree.selection():
                self.task_tree.selection_set(iid)

            # --- ↓↓↓ 核心修改：在这里获取任务类型并决定状态 ↓↓↓ ---
            # 只有当选中单个任务时，我们才进行判断
            if len(self.task_tree.selection()) == 1:
                index = self.task_tree.index(iid)
                task = self.tasks[index]
                task_type = task.get('type')
                
                # 如果任务类型是 'bell_schedule'，则禁用“立即播放”
                play_now_state = 'disabled' if task_type == 'bell_schedule' else 'normal'
            else:
                # 如果选中了多个任务，为简单起见，也禁用“立即播放”
                play_now_state = 'disabled'
            
            context_menu.add_command(label="立即播放", command=self.play_now, state=play_now_state)
            # --- ↑↑↑ 修改结束 ↑↑↑ ---

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

    def play_now(self):
        selection = self.task_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择一个要立即播放的节目。", parent=self.root)
            return
        index = self.task_tree.index(selection[0])
        task = self.tasks[index]
        self.log(f"手动触发高优先级播放: {task['name']}")
        self.playback_command_queue.put(('PLAY_INTERRUPT', (task, "manual_play")))

    def stop_current_playback(self):
        self.log("手动触发“停止当前播放”...")
        self.playback_command_queue.put(('STOP', None))

    def add_task(self):
        if self.auth_info['status'] == 'Trial' and len(self.tasks) >= 3:
            messagebox.showerror(
                "试用版限制", 
                "试用版最多只能添加3个定时广播节目。\n\n请删除现有节目后再添加，或注册软件以解除全部限制。", 
                parent=self.root
            )
            return

        choice_dialog = ttk.Toplevel(self.root)
        choice_dialog.title("选择节目类型")
        choice_dialog.resizable(False, False)
        choice_dialog.transient(self.root)
        
        choice_dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            choice_dialog.destroy()
            self.root.focus_force()

        def open_and_cleanup(dialog_opener_func, *args):
            choice_dialog.destroy()
            self.root.attributes('-disabled', False)
            temp_parent = ttk.Toplevel(self.root)
            temp_parent.withdraw()
            dialog_opener_func(temp_parent, *args)

        main_frame = ttk.Frame(choice_dialog, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        title_label = ttk.Label(main_frame, text="请选择要添加的节目类型",
                              font=self.font_13_bold, bootstyle="primary")
        title_label.pack(pady=15)
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(expand=True, fill=X)

        audio_btn = ttk.Button(btn_frame, text="🎵→音频节目",
                             bootstyle="primary", width=20, command=lambda: open_and_cleanup(self.open_audio_dialog))
        audio_btn.pack(pady=8, ipady=8, fill=X)

        voice_btn = ttk.Button(btn_frame, text="🎤→语音节目",
                             bootstyle="info", width=20, command=lambda: open_and_cleanup(self.open_voice_dialog))
        voice_btn.pack(pady=8, ipady=8, fill=X)

        # --- ↓↓↓ 新增代码：动态语音按钮 ↓↓↓ ---
        dynamic_voice_btn = ttk.Button(btn_frame, text="💬→动态语音",
                             bootstyle="success", width=20, command=lambda: open_and_cleanup(self.open_dynamic_voice_dialog))
        dynamic_voice_btn.pack(pady=8, ipady=8, fill=X)
        # --- ↑↑↑ 新增代码结束 ↑↑↑ ---

        video_btn = ttk.Button(btn_frame, text="🎬→视频节目",
                             bootstyle="dark", width=20, command=lambda: open_and_cleanup(self.open_video_dialog))
        video_btn.pack(pady=8, ipady=8, fill=X)
        if not VLC_AVAILABLE:
            video_btn.config(state=DISABLED, text="🎬→视频节目 (VLC未安装)")

        bell_btn = ttk.Button(btn_frame, text="🔔→打铃模式",
                             bootstyle="warning", width=20, command=lambda: open_and_cleanup(self.open_bell_scheduler_dialog))
        bell_btn.pack(pady=8, ipady=8, fill=X)

        choice_dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)
        self.center_window(choice_dialog, parent=self.root)
#第5部分
    def open_bell_scheduler_dialog(self, parent_dialog, task_to_edit=None, index=None):
        if parent_dialog and parent_dialog.winfo_exists():
            parent_dialog.destroy()
        
        is_edit_mode = task_to_edit is not None
        
        dialog = ttk.Toplevel(self.root)
        dialog.title("修改打铃计划" if is_edit_mode else "校铃/厂铃时间表助手")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()

        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=BOTH, expand=True)
        main_frame.columnconfigure(0, weight=3)
        main_frame.columnconfigure(1, weight=2)

        left_frame = ttk.Frame(main_frame)
        left_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 10))

        name_lf = ttk.LabelFrame(left_frame, text="节目名称", padding=10)
        name_lf.pack(fill=X, pady=(0, 5))
        name_entry = ttk.Entry(name_lf, font=self.font_11)
        name_entry.pack(fill=X)

        bell_files_lf = ttk.LabelFrame(left_frame, text="1. 铃声文件设置", padding=10)
        bell_files_lf.pack(fill=X, pady=5)
        bell_files_lf.columnconfigure(1, weight=1)
        
        up_bell_var, down_bell_var, bell_volume_var = tk.StringVar(), tk.StringVar(), tk.StringVar(value="80")
        
        ttk.Label(bell_files_lf, text="上课/上班铃:").grid(row=0, column=0, sticky='e', padx=5)
        ttk.Entry(bell_files_lf, textvariable=up_bell_var, font=self.font_11).grid(row=0, column=1, sticky='ew')
        ttk.Button(bell_files_lf, text="选取", bootstyle="outline", width=5, command=lambda: self.select_file_for_entry(AUDIO_FOLDER, up_bell_var, dialog)).grid(row=0, column=2, padx=5)

        ttk.Label(bell_files_lf, text="下课/下班铃:").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        ttk.Entry(bell_files_lf, textvariable=down_bell_var, font=self.font_11).grid(row=1, column=1, sticky='ew')
        ttk.Button(bell_files_lf, text="选取", bootstyle="outline", width=5, command=lambda: self.select_file_for_entry(AUDIO_FOLDER, down_bell_var, dialog)).grid(row=1, column=2, padx=5)
        
        ttk.Label(bell_files_lf, text="统一音量:").grid(row=2, column=0, sticky='e', padx=5)
        ttk.Entry(bell_files_lf, textvariable=bell_volume_var, width=8, font=self.font_11).grid(row=2, column=1, sticky='w')
        ttk.Label(bell_files_lf, text="(0-100)", font=self.font_9, bootstyle="secondary").grid(row=2, column=1, sticky='w', padx=70)

        schedule_lf = ttk.LabelFrame(left_frame, text="2. 通用规则设置", padding=10)
        schedule_lf.pack(fill=X, pady=5)
        schedule_lf.columnconfigure(1, weight=1)

        weekday_var = tk.StringVar(value="每周:12345")
        daterange_var = tk.StringVar(value="2025-01-01 ~ 2099-12-31")

        ttk.Label(schedule_lf, text="周几执行:").grid(row=0, column=0, sticky='e', padx=5)
        weekday_entry_schedule = ttk.Entry(schedule_lf, textvariable=weekday_var, font=self.font_11)
        weekday_entry_schedule.grid(row=0, column=1, sticky='ew')
        ttk.Button(schedule_lf, text="选取", bootstyle="outline", width=5, command=lambda: self.show_weekday_settings_dialog(weekday_entry_schedule)).grid(row=0, column=2, padx=5)

        ttk.Label(schedule_lf, text="日期范围:").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        daterange_entry_schedule = ttk.Entry(schedule_lf, textvariable=daterange_var, font=self.font_11)
        daterange_entry_schedule.grid(row=1, column=1, sticky='ew')
        ttk.Button(schedule_lf, text="设置", bootstyle="outline", width=5, command=lambda: self.show_daterange_settings_dialog(daterange_entry_schedule)).grid(row=1, column=2, padx=5)

        class_time_lf = ttk.LabelFrame(left_frame, text="3. 时间点设置", padding=10)
        class_time_lf.pack(fill=X, pady=5)
        
        notebook = ttk.Notebook(class_time_lf)
        notebook.pack(fill=BOTH, expand=True, pady=5)

        # <--- 新增: 创建第三个Frame用于夜班 ---
        am_tab, pm_tab, night_tab = ttk.Frame(notebook, padding=10), ttk.Frame(notebook, padding=10), ttk.Frame(notebook, padding=10)
        notebook.add(am_tab, text=" 上午/白班 ")
        notebook.add(pm_tab, text=" 下午/晚班 ")
        # <--- 新增: 添加新的选项卡到Notebook ---
        notebook.add(night_tab, text=" 夜自习/夜班 ")

        def create_session_ui(parent, prefix):
            parent.columnconfigure(1, weight=1)
            
            # <--- 修改: 为夜班设置不同的默认值 ---
            default_start_time = "08:00:00"
            default_periods = "4"
            default_use_long_break = True
            if prefix == "下午":
                default_start_time = "14:00:00"
                default_periods = "3"
                default_use_long_break = False
            elif prefix == "夜间":
                default_start_time = "19:00:00"
                default_periods = "2"
                default_use_long_break = False

            vars = {
                'start_time': tk.StringVar(value=default_start_time),
                'periods': tk.StringVar(value=default_periods),
                'duration': tk.StringVar(value="45"),
                'short_break': tk.StringVar(value="10"),
                'use_long_break': tk.BooleanVar(value=default_use_long_break),
                'long_break_after': tk.StringVar(value="2"),
                'long_break_duration': tk.StringVar(value="25")
            }

            ttk.Label(parent, text=f"{prefix}第一节开始时间:").grid(row=0, column=0, sticky='e', padx=5)
            start_time_entry = ttk.Entry(parent, textvariable=vars['start_time'], width=12, font=self.font_11)
            start_time_entry.grid(row=0, column=1, columnspan=2, sticky='w')
            self._bind_mousewheel_to_entry(start_time_entry, self._handle_time_scroll)

            ttk.Label(parent, text=f"{prefix}课程/工作节数:").grid(row=1, column=0, sticky='e', padx=5, pady=2)
            ttk.Entry(parent, textvariable=vars['periods'], width=12, font=self.font_11).grid(row=1, column=1, columnspan=2, sticky='w')
            
            ttk.Label(parent, text="每节时长(分钟):").grid(row=2, column=0, sticky='e', padx=5, pady=2)
            ttk.Entry(parent, textvariable=vars['duration'], width=12, font=self.font_11).grid(row=2, column=1, columnspan=2, sticky='w')

            ttk.Label(parent, text="课间/休息时长(分钟):").grid(row=3, column=0, sticky='e', padx=5, pady=2)
            ttk.Entry(parent, textvariable=vars['short_break'], width=12, font=self.font_11).grid(row=3, column=1, columnspan=2, sticky='w')
            
            long_break_cb = ttk.Checkbutton(parent, text="启用大课间/长休息", variable=vars['use_long_break'], bootstyle="round-toggle")
            long_break_cb.grid(row=4, column=0, columnspan=3, pady=5)

            long_break_frame = ttk.Frame(parent)
            long_break_frame.grid(row=5, column=0, columnspan=3, sticky='w', padx=25)
            ttk.Label(long_break_frame, text="在第").pack(side=LEFT)
            ttk.Entry(long_break_frame, textvariable=vars['long_break_after'], width=5, font=self.font_11).pack(side=LEFT, padx=2)
            ttk.Label(long_break_frame, text="节后，休息").pack(side=LEFT)
            ttk.Entry(long_break_frame, textvariable=vars['long_break_duration'], width=5, font=self.font_11).pack(side=LEFT, padx=2)
            ttk.Label(long_break_frame, text="分钟").pack(side=LEFT)

            return vars

        am_vars = create_session_ui(am_tab, "上午")
        pm_vars = create_session_ui(pm_tab, "下午")
        # <--- 新增: 创建夜班的UI变量 ---
        night_vars = create_session_ui(night_tab, "夜间")

        right_frame = ttk.Frame(main_frame)
        right_frame.grid(row=0, column=1, sticky='nsew')
        right_frame.rowconfigure(0, weight=1)
        right_frame.columnconfigure(0, weight=1)

        preview_lf = ttk.LabelFrame(right_frame, text="4. 生成预览", padding=10)
        preview_lf.pack(fill=BOTH, expand=True)
        preview_lf.rowconfigure(0, weight=1)
        preview_lf.columnconfigure(0, weight=1)
        
        preview_text = ScrolledText(preview_lf, height=15, font=self.font_11, wrap=WORD)
        preview_text.grid(row=0, column=0, sticky='nsew')
        preview_text.text.config(state=DISABLED)
        
        bottom_frame = ttk.Frame(dialog)
        bottom_frame.pack(fill=X, padx=15, pady=(5, 10))
        
        commit_btn_text = "保存修改" if is_edit_mode else "添加至节目单"
        # <--- 修改: 将night_vars传递给_commit_bells_to_schedule ---
        commit_btn = ttk.Button(bottom_frame, text=commit_btn_text, bootstyle="success", state=DISABLED, 
                                command=lambda: self._commit_bells_to_schedule(
                                    preview_text, name_entry, up_bell_var, down_bell_var, bell_volume_var, 
                                    weekday_var, daterange_var, am_vars, pm_vars, night_vars, 
                                    dialog, cleanup_and_destroy, task_to_edit, index
                                ))

        # <--- 修改: 将night_vars传递给_generate_and_preview_bells ---
        preview_btn = ttk.Button(bottom_frame, text="生成预览", bootstyle="info", command=lambda: self._generate_and_preview_bells(
            preview_text, up_bell_var, down_bell_var, bell_volume_var, am_vars, pm_vars, night_vars, commit_btn, dialog
        ))
        preview_btn.pack(side=LEFT, padx=10, ipady=4)
        commit_btn.pack(side=LEFT, padx=10, ipady=4)

        ttk.Button(bottom_frame, text="取消", bootstyle="secondary", command=cleanup_and_destroy).pack(side=RIGHT, padx=10, ipady=4)

        if is_edit_mode:
            name_entry.insert(0, task_to_edit.get('name', ''))
            up_bell_var.set(task_to_edit.get('up_bell_file', ''))
            down_bell_var.set(task_to_edit.get('down_bell_file', ''))
            bell_volume_var.set(task_to_edit.get('volume', '80'))
            weekday_var.set(task_to_edit.get('weekday', '每周:12345'))
            daterange_var.set(task_to_edit.get('date_range', '2025-01-01 ~ 2099-12-31'))
            
            params = task_to_edit.get('schedule_params', {})
            am_params = params.get('am', {})
            pm_params = params.get('pm', {})
            # <--- 新增: 从任务数据中加载夜班设置 ---
            night_params = params.get('night', {})
            
            for key, var in am_vars.items():
                if isinstance(var, tk.BooleanVar):
                    var.set(am_params.get(key, False))
                else:
                    var.set(am_params.get(key, ''))
            
            for key, var in pm_vars.items():
                if isinstance(var, tk.BooleanVar):
                    var.set(pm_params.get(key, False))
                else:
                    var.set(pm_params.get(key, ''))
            
            # <--- 新增: 将加载的夜班设置填充到UI变量中 ---
            for key, var in night_vars.items():
                if isinstance(var, tk.BooleanVar):
                    var.set(night_params.get(key, False))
                else:
                    var.set(night_params.get(key, ''))
            
            dialog.after(100, preview_btn.invoke)
        else:
            name_entry.insert(0, "校园/工厂作息铃声")

        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)
        dialog.after(100, lambda: self.center_window(dialog, parent=self.root))

    def _generate_and_preview_bells(self, preview_text_widget, up_bell_var, down_bell_var, bell_volume_var, am_vars, pm_vars, night_vars, commit_btn, parent_dialog):
        """计算并显示铃声时间表的预览"""
        try:
            if not up_bell_var.get().strip() or not down_bell_var.get().strip():
                messagebox.showerror("输入错误", "请必须选择“上课铃声”和“下课铃声”文件。", parent=parent_dialog)
                return
            volume = int(bell_volume_var.get())
            if not (0 <= volume <= 100): raise ValueError("音量必须在 0-100 之间")

            preview_content = []
            
            def calculate_session(prefix, session_vars):
                start_time_str = self._normalize_time_string(session_vars['start_time'].get())
                if not start_time_str: raise ValueError(f"{prefix}开始时间格式错误")
                current_time = datetime.strptime(start_time_str, "%H:%M:%S")
                
                periods_str = session_vars['periods'].get().strip()
                periods = int(periods_str) if periods_str else 0
                if periods < 0: raise ValueError("节数不能为负数")
                if periods == 0: return
                
                duration_min = int(session_vars['duration'].get())
                short_break_min = int(session_vars['short_break'].get())
                use_long_break = session_vars['use_long_break'].get()
                long_break_after = int(session_vars['long_break_after'].get()) if use_long_break else -1
                long_break_duration_min = int(session_vars['long_break_duration'].get()) if use_long_break else 0

                for i in range(1, periods + 1):
                    preview_content.append(f"[{prefix}第{i}节 上课铃] {current_time.strftime('%H:%M:%S')}")
                    current_time += timedelta(minutes=duration_min)
                    preview_content.append(f"[{prefix}第{i}节 下课铃] {current_time.strftime('%H:%M:%S')}")

                    if i < periods:
                        if use_long_break and i == long_break_after:
                            current_time += timedelta(minutes=long_break_duration_min)
                        else:
                            current_time += timedelta(minutes=short_break_min)
            
            calculate_session("上午", am_vars)
            if preview_content and int(pm_vars['periods'].get().strip() or 0) > 0:
                preview_content.append("-" * 30)
            calculate_session("下午", pm_vars)
            
            # <--- 新增: 计算夜班并添加分隔符 ---
            if preview_content and int(night_vars['periods'].get().strip() or 0) > 0:
                preview_content.append("-" * 30)
            calculate_session("夜间", night_vars)

            preview_text_widget.text.config(state=NORMAL)
            preview_text_widget.text.delete('1.0', END)
            preview_text_widget.text.insert('1.0', "\n".join(preview_content))
            preview_text_widget.text.config(state=DISABLED)

            commit_btn.config(state=NORMAL if preview_content else DISABLED)

        except (ValueError, TypeError) as e:
            messagebox.showerror("输入错误", f"请检查所有时间、时长和节数是否为有效的纯数字。\n\n错误详情: {e}", parent=parent_dialog)
            commit_btn.config(state=DISABLED)
            return

    def _commit_bells_to_schedule(self, preview_text_widget, name_entry, up_bell_var, down_bell_var, bell_volume_var, weekday_var, daterange_var, am_vars, pm_vars, night_vars, parent_dialog, close_callback, task_to_edit=None, index=None):
        preview_content = preview_text_widget.text.get('1.0', END).strip()
        if not preview_content:
            messagebox.showwarning("无内容", "预览为空，无法添加。", parent=parent_dialog)
            return
            
        generated_times = []
        lines = preview_content.split('\n')
        for line in lines:
            if not line.strip() or line.startswith('-'):
                continue
            
            match = re.match(r'\[(.*?)\]\s*(\d{2}:\d{2}:\d{2})', line)
            if not match:
                continue
            
            task_name = match.group(1).strip()
            task_time = match.group(2).strip()
            bell_type = 'up' if "上课" in task_name or "上班" in task_name else 'down'
            generated_times.append({'name': task_name, 'time': task_time, 'bell_type': bell_type})
        
        if not generated_times:
            messagebox.showwarning("无内容", "未能从预览中解析出有效的时间点。", parent=parent_dialog)
            return

        new_bell_schedule_task = {
            'name': name_entry.get().strip() or "未命名打铃计划",
            'type': 'bell_schedule',
            'status': '启用' if task_to_edit is None else task_to_edit.get('status', '启用'),
            'weekday': weekday_var.get(),
            'date_range': daterange_var.get(),
            'up_bell_file': up_bell_var.get(),
            'down_bell_file': down_bell_var.get(),
            'volume': bell_volume_var.get(),
            'schedule_params': {
                'am': {k: v.get() if not isinstance(v, tk.BooleanVar) else bool(v.get()) for k, v in am_vars.items()},
                'pm': {k: v.get() if not isinstance(v, tk.BooleanVar) else bool(v.get()) for k, v in pm_vars.items()},
                # <--- 新增: 将夜班的设置也保存起来 ---
                'night': {k: v.get() if not isinstance(v, tk.BooleanVar) else bool(v.get()) for k, v in night_vars.items()}
            },
            'generated_times': generated_times,
            'last_run': {} if task_to_edit is None else task_to_edit.get('last_run', {})
        }
        
        if task_to_edit is None:
            self.tasks.append(new_bell_schedule_task)
            self.log(f"通过“打铃模式”成功添加了一个名为 '{new_bell_schedule_task['name']}' 的铃声计划。")
            messagebox.showinfo("成功", f"已成功生成并添加了一个包含 {len(generated_times)} 个时间点的铃声计划！", parent=self.root)
        else:
            self.tasks[index] = new_bell_schedule_task
            self.log(f"已成功修改打铃计划 '{new_bell_schedule_task['name']}'。")
            messagebox.showinfo("成功", "打铃计划已成功修改！", parent=self.root)
        
        self.update_task_list()
        self.save_tasks()
        close_callback()

    def open_audio_dialog(self, parent_dialog, task_to_edit=None, index=None):
        parent_dialog.destroy()
        is_edit_mode = task_to_edit is not None
        dialog = ttk.Toplevel(self.root)
        dialog.title("修改音频节目" if is_edit_mode else "添加音频节目")
        dialog.resizable(True, True)
        dialog.minsize(800, 600) #稍微增加高度以容纳新选项
        dialog.transient(self.root)

        # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()
        # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=BOTH, expand=True)

        content_frame = ttk.LabelFrame(main_frame, text="内容", padding=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=2)
        content_frame.columnconfigure(1, weight=1)

        # --- 0. 节目名称 ---
        ttk.Label(content_frame, text="节目名称:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        name_entry = ttk.Entry(content_frame, font=self.font_11)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        
        audio_type_var = tk.StringVar(value="single")
        # 用于暂存播放列表数据的变量
        self.temp_playlist_data = [] 

        # --- 1. 单文件模式 ---
        ttk.Label(content_frame, text="音频文件:").grid(row=1, column=0, sticky='e', padx=5, pady=2)
        audio_single_frame = ttk.Frame(content_frame)
        audio_single_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        audio_single_frame.columnconfigure(1, weight=1)
        
        ttk.Radiobutton(audio_single_frame, text="", variable=audio_type_var, value="single").grid(row=0, column=0, sticky='w')
        audio_single_entry = ttk.Entry(audio_single_frame, font=self.font_11)
        audio_single_entry.grid(row=0, column=1, sticky='ew', padx=5)
        
        # 根据VLC是否可用设置提示
        if VLC_AVAILABLE:
            filetypes = [("所有支持的音频", "*.mp3 *.wav *.ogg *.flac *.m4a *.wma *.ape"), ("所有文件", "*.*")]
            vlc_info_text = " (VLC支持多格式)"
        else:
            filetypes = [("支持的音频", "*.mp3 *.wav *.ogg *.flac"), ("所有文件", "*.*")]
            vlc_info_text = " (仅基础格式)"

        def select_single_audio():
            filename = filedialog.askopenfilename(
                title="选择音频文件", 
                initialdir=AUDIO_FOLDER, 
                filetypes=filetypes, 
                parent=dialog
            )
            if filename: 
                audio_single_entry.delete(0, END)
                audio_single_entry.insert(0, filename)
        
        ttk.Button(audio_single_frame, text="选取...", command=select_single_audio, bootstyle="outline").grid(row=0, column=3, padx=5)
        ttk.Label(audio_single_frame, text=vlc_info_text, font=self.font_9, bootstyle="secondary").grid(row=0, column=4, sticky='w')

        # --- 2. [新增] 自定义列表模式 ---
        # 这一行放在单文件和文件夹之间
        ttk.Label(content_frame, text="音频列表:").grid(row=2, column=0, sticky='e', padx=5, pady=2)
        playlist_frame = ttk.Frame(content_frame)
        playlist_frame.grid(row=2, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        
        ttk.Radiobutton(playlist_frame, text="", variable=audio_type_var, value="playlist").pack(side=LEFT)
        
        self.playlist_info_var = tk.StringVar(value="(包含 0 首歌曲)")
        playlist_info_label = ttk.Label(playlist_frame, textvariable=self.playlist_info_var)
        playlist_info_label.pack(side=LEFT, padx=10)
        
        def launch_editor():
            # 调用编辑器，传入当前dialog作为父窗口，以及当前的数据
            new_list = self.open_playlist_editor(dialog, self.temp_playlist_data)
            if new_list is not None: # 用户点击了确定
                self.temp_playlist_data = new_list
                self.playlist_info_var.set(f"(包含 {len(new_list)} 首歌曲)")
                # 自动选中列表模式
                audio_type_var.set("playlist")

        self.edit_playlist_btn = ttk.Button(playlist_frame, text="编辑播放列表...", command=launch_editor, bootstyle="info-outline")
        self.edit_playlist_btn.pack(side=LEFT)

        # --- 3. [修改] 文件夹模式 (包含顺序/随机播) ---
        ttk.Label(content_frame, text="音频文件夹:").grid(row=3, column=0, sticky='e', padx=5, pady=2)
        audio_folder_frame = ttk.Frame(content_frame)
        audio_folder_frame.grid(row=3, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        audio_folder_frame.columnconfigure(1, weight=1)
        
        ttk.Radiobutton(audio_folder_frame, text="", variable=audio_type_var, value="folder").grid(row=0, column=0, sticky='w')
        audio_folder_entry = ttk.Entry(audio_folder_frame, font=self.font_11)
        audio_folder_entry.grid(row=0, column=1, sticky='ew', padx=5)
        
        def select_folder(entry_widget):
            foldername = filedialog.askdirectory(title="选择文件夹", initialdir=application_path, parent=dialog)
            if foldername: entry_widget.delete(0, END); entry_widget.insert(0, foldername)
        ttk.Button(audio_folder_frame, text="选取...", command=lambda: select_folder(audio_folder_entry), bootstyle="outline").grid(row=0, column=2, padx=5)
        
        # [改动] 将播放顺序选项移动到这里
        play_order_var = tk.StringVar(value="sequential")
        self.folder_seq_rb = ttk.Radiobutton(audio_folder_frame, text="顺序播", variable=play_order_var, value="sequential")
        self.folder_seq_rb.grid(row=0, column=3, padx=(10,0))
        self.folder_rand_rb = ttk.Radiobutton(audio_folder_frame, text="随机播", variable=play_order_var, value="random")
        self.folder_rand_rb.grid(row=0, column=4, padx=5)
        
        # --- 4. 背景图片 ---
        bg_image_var = tk.IntVar(value=0)
        bg_image_path_var = tk.StringVar()
        bg_image_order_var = tk.StringVar(value="sequential")

        bg_image_frame = ttk.Frame(content_frame)
        bg_image_frame.grid(row=4, column=0, columnspan=4, sticky='w', padx=5, pady=5)
        bg_image_frame.columnconfigure(1, weight=1)
        bg_image_cb = ttk.Checkbutton(bg_image_frame, text="背景图片:", variable=bg_image_var, bootstyle="round-toggle")
        bg_image_cb.grid(row=0, column=0)
        if not IMAGE_AVAILABLE: bg_image_cb.config(state=DISABLED, text="背景图片(无库):")

        bg_image_entry = ttk.Entry(bg_image_frame, textvariable=bg_image_path_var, font=self.font_11)
        bg_image_entry.grid(row=0, column=1, sticky='ew', padx=(5,5))

        bg_image_btn_frame = ttk.Frame(bg_image_frame)
        bg_image_btn_frame.grid(row=0, column=2)
        ttk.Button(bg_image_btn_frame, text="选取...", command=lambda: select_folder(bg_image_entry), bootstyle="outline").pack(side=LEFT)
        ttk.Radiobutton(bg_image_btn_frame, text="顺序", variable=bg_image_order_var, value="sequential").pack(side=LEFT, padx=(10,0))
        ttk.Radiobutton(bg_image_btn_frame, text="随机", variable=bg_image_order_var, value="random").pack(side=LEFT)

        # --- 5. 音量 ---
        volume_frame = ttk.Frame(content_frame)
        volume_frame.grid(row=5, column=1, columnspan=3, sticky='w', padx=5, pady=3)
        ttk.Label(volume_frame, text="音量:").pack(side=LEFT)
        volume_entry = ttk.Entry(volume_frame, font=self.font_11, width=10)
        volume_entry.pack(side=LEFT, padx=5)
        ttk.Label(volume_frame, text="0-100").pack(side=LEFT, padx=5)

        # --- 6. 时间与规则 ---
        time_frame = ttk.LabelFrame(main_frame, text="时间", padding=15)
        time_frame.grid(row=1, column=0, sticky='ew', pady=4)
        time_frame.columnconfigure(1, weight=1)
        
        ttk.Label(time_frame, text="开始时间:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        start_time_entry = ttk.Entry(time_frame, font=self.font_11)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        self._bind_mousewheel_to_entry(start_time_entry, self._handle_time_scroll)
        ttk.Label(time_frame, text="<可多个>").grid(row=0, column=2, sticky='w', padx=5)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bootstyle="outline").grid(row=0, column=3, padx=5)

        # 批量添加容器
        batch_add_container = ttk.Frame(time_frame)
        batch_add_container.grid(row=0, column=4, rowspan=3, sticky='n', padx=5)
        batch_interval_frame = ttk.Frame(batch_add_container)
        batch_interval_frame.pack(pady=(0, 2))
        ttk.Label(batch_interval_frame, text="每").pack(side=LEFT)
        batch_interval_entry = ttk.Entry(batch_interval_frame, font=self.font_11, width=4)
        batch_interval_entry.pack(side=LEFT, padx=(2,2))
        ttk.Label(batch_interval_frame, text="分钟").pack(side=LEFT)
        batch_count_frame = ttk.Frame(batch_add_container)
        batch_count_frame.pack(pady=(0, 5))
        ttk.Label(batch_count_frame, text="共").pack(side=LEFT)
        batch_count_entry = ttk.Entry(batch_count_frame, font=self.font_11, width=4)
        batch_count_entry.pack(side=LEFT, padx=(2,2))
        ttk.Label(batch_count_frame, text="次   ").pack(side=LEFT)
        ttk.Button(batch_add_container, text="批量添加", 
                   command=lambda: self._apply_batch_time_addition(start_time_entry, batch_interval_entry, batch_count_entry, dialog), 
                   bootstyle="outline-info").pack(fill=X)

        # 间隔播报设置
        interval_var = tk.StringVar(value="first")
        ttk.Label(time_frame, text="间隔播报:").grid(row=1, column=0, sticky='e', padx=5, pady=2)
        
        interval_frame1 = ttk.Frame(time_frame)
        interval_frame1.grid(row=1, column=1, columnspan=2, sticky='w', padx=5, pady=2)
        # 需要保存引用以便动态修改文本
        self.lbl_interval_first = ttk.Radiobutton(interval_frame1, text="播 n 首", variable=interval_var, value="first")
        self.lbl_interval_first.pack(side=LEFT)
        interval_first_entry = ttk.Entry(interval_frame1, font=self.font_11, width=15)
        interval_first_entry.pack(side=LEFT, padx=5)
        self.lbl_interval_hint = ttk.Label(interval_frame1, text="(单曲时,指 n 遍)")
        self.lbl_interval_hint.pack(side=LEFT, padx=5)
        
        interval_frame2 = ttk.Frame(time_frame)
        interval_frame2.grid(row=2, column=1, columnspan=2, sticky='w', padx=5, pady=2)
        ttk.Radiobutton(interval_frame2, text="播 n 秒", variable=interval_var, value="seconds").pack(side=LEFT)
        interval_seconds_entry = ttk.Entry(interval_frame2, font=self.font_11, width=15)
        interval_seconds_entry.pack(side=LEFT, padx=5)
        ttk.Label(interval_frame2, text="(3600秒 = 1小时)").pack(side=LEFT, padx=5)
        
        ttk.Label(time_frame, text="周几/几号:").grid(row=3, column=0, sticky='e', padx=5, pady=3)
        weekday_entry = ttk.Entry(time_frame, font=self.font_11)
        weekday_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=3)
        ttk.Button(time_frame, text="选取...", command=lambda: self.show_weekday_settings_dialog(weekday_entry), bootstyle="outline").grid(row=3, column=3, padx=5)
        
        ttk.Label(time_frame, text="日期范围:").grid(row=4, column=0, sticky='e', padx=5, pady=3)
        date_range_entry = ttk.Entry(time_frame, font=self.font_11)
        date_range_entry.grid(row=4, column=1, sticky='ew', padx=5, pady=3)
        self._bind_mousewheel_to_entry(date_range_entry, self._handle_date_scroll)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry), bootstyle="outline").grid(row=4, column=3, padx=5)

        other_frame = ttk.LabelFrame(main_frame, text="其它", padding=10)
        other_frame.grid(row=2, column=0, sticky='ew', pady=5)
        other_frame.columnconfigure(1, weight=1)
        
        delay_var = tk.StringVar(value="ontime")
        ttk.Label(other_frame, text="模式:").grid(row=0, column=0, sticky='nw', padx=5, pady=2)
        delay_frame = ttk.Frame(other_frame)
        delay_frame.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        ttk.Radiobutton(delay_frame, text="准时播 - 如果有别的节目正在播，终止他们（默认）", variable=delay_var, value="ontime").pack(anchor='w')
        ttk.Radiobutton(delay_frame, text="可延后 - 如果有别的节目正在播，排队等候", variable=delay_var, value="delay").pack(anchor='w')
        ttk.Radiobutton(delay_frame, text="立即播 - 添加后停止其他节目,立即播放此节目", variable=delay_var, value="immediate").pack(anchor='w')
        
        dialog_button_frame = ttk.Frame(other_frame)
        dialog_button_frame.grid(row=0, column=2, sticky='se', padx=20, pady=10)

        # --- 数据加载逻辑 ---
        if is_edit_mode:
            task = task_to_edit
            name_entry.insert(0, task.get('name', ''))
            start_time_entry.insert(0, task.get('time', ''))
            audio_type_var.set(task.get('audio_type', 'single'))
            
            if task.get('audio_type') == 'single': 
                audio_single_entry.insert(0, task.get('content', ''))
            elif task.get('audio_type') == 'folder':
                audio_folder_entry.insert(0, task.get('content', ''))
            elif task.get('audio_type') == 'playlist':
                # 加载保存的播放列表
                self.temp_playlist_data = task.get('custom_playlist', [])
                self.playlist_info_var.set(f"(包含 {len(self.temp_playlist_data)} 首歌曲)")

            play_order_var.set(task.get('play_order', 'sequential'))
            volume_entry.insert(0, task.get('volume', '80'))
            interval_var.set(task.get('interval_type', 'first'))
            interval_first_entry.insert(0, task.get('interval_first', '1'))
            interval_seconds_entry.insert(0, task.get('interval_seconds', '600'))
            weekday_entry.insert(0, task.get('weekday', '每周:1234567'))
            date_range_entry.insert(0, task.get('date_range', '2025-01-01 ~ 2099-12-31'))
            delay_var.set(task.get('delay', 'ontime'))
            bg_image_var.set(task.get('bg_image_enabled', 0))
            bg_image_path_var.set(task.get('bg_image_path', ''))
            bg_image_order_var.set(task.get('bg_image_order', 'sequential'))
        else:
            volume_entry.insert(0, "80"); interval_first_entry.insert(0, "1"); interval_seconds_entry.insert(0, "600")
            weekday_entry.insert(0, "每周:1234567"); date_range_entry.insert(0, "2025-01-01 ~ 2099-12-31")

        # --- [新增] UI 联动逻辑 ---
        def toggle_controls(*args):
            atype = audio_type_var.get()
            # 1. 控制文件夹排序按钮
            folder_state = 'normal' if atype == 'folder' else 'disabled'
            self.folder_seq_rb.config(state=folder_state)
            self.folder_rand_rb.config(state=folder_state)
            
            # 2. 动态更新间隔播报的提示文本
            if atype == 'playlist':
                self.lbl_interval_first.config(text="播 n 遍")
                self.lbl_interval_hint.config(text="(指将列表完整播放n遍)")
            else:
                self.lbl_interval_first.config(text="播 n 首")
                self.lbl_interval_hint.config(text="(单曲时,指 n 遍)")

        audio_type_var.trace_add("write", toggle_controls)
        # 初始化一次状态
        self.root.after(10, toggle_controls)

        def save_task():
            # 验证音量
            try:
                volume = int(volume_entry.get().strip() or 80)
                if not (0 <= volume <= 100): raise ValueError
            except ValueError:
                messagebox.showerror("输入错误", "音量必须是 0 到 100 之间的整数。", parent=dialog); return

            # 验证间隔次数
            if interval_var.get() == 'first':
                try:
                    val = int(interval_first_entry.get().strip() or 1)
                    if val < 1: raise ValueError
                except ValueError:
                    messagebox.showerror("输入错误", "次数必须大于或等于 1。", parent=dialog); return
            else: 
                try:
                    val = int(interval_seconds_entry.get().strip() or 1)
                    if val < 1: raise ValueError
                except ValueError:
                    messagebox.showerror("输入错误", "秒数必须大于或等于 1。", parent=dialog); return

            if not weekday_entry.get().strip(): messagebox.showerror("输入错误", "周几规则不能为空", parent=dialog); return
            if not date_range_entry.get().strip(): messagebox.showerror("输入错误", "日期范围不能为空", parent=dialog); return

            # 验证音频内容
            audio_type = audio_type_var.get()
            audio_path = ""
            custom_playlist = []
            
            if audio_type == 'single':
                audio_path = audio_single_entry.get().strip()
                if not audio_path: messagebox.showwarning("警告", "请选择音频文件", parent=dialog); return
            elif audio_type == 'folder':
                audio_path = audio_folder_entry.get().strip()
                if not audio_path: messagebox.showwarning("警告", "请选择音频文件夹", parent=dialog); return
            elif audio_type == 'playlist':
                if not self.temp_playlist_data:
                    messagebox.showwarning("警告", "播放列表为空，请点击'编辑播放列表'添加歌曲。", parent=dialog); return
                custom_playlist = self.temp_playlist_data
                audio_path = None # 列表模式下，content 字段为空

            is_valid_time, time_msg = self._normalize_multiple_times_string(start_time_entry.get().strip())
            if not is_valid_time: messagebox.showwarning("格式错误", time_msg, parent=dialog); return
            is_valid_date, date_msg = self._normalize_date_range_string(date_range_entry.get().strip())
            if not is_valid_date: messagebox.showwarning("格式错误", date_msg, parent=dialog); return

            play_mode = delay_var.get()
            play_this_task_now = (play_mode == 'immediate')
            saved_delay_type = 'ontime' if play_mode == 'immediate' else play_mode

            new_task_data = {
                'name': name_entry.get().strip(), 'time': time_msg, 'type': 'audio',
                'audio_type': audio_type,
                'content': audio_path,
                'custom_playlist': custom_playlist, # 保存列表数据
                'play_order': play_order_var.get(),
                'volume': str(volume), 'interval_type': interval_var.get(),
                'interval_first': interval_first_entry.get().strip() or "1",
                'interval_seconds': interval_seconds_entry.get().strip() or "600",
                'weekday': weekday_entry.get().strip(), 'date_range': date_msg, 'delay': saved_delay_type,
                'status': '启用' if not is_edit_mode else task_to_edit.get('status', '启用'),
                'last_run': {} if not is_edit_mode else task_to_edit.get('last_run', {}),
                'bg_image_enabled': bg_image_var.get(),
                'bg_image_path': bg_image_path_var.get().strip(),
                'bg_image_order': bg_image_order_var.get()
            }
            if not new_task_data['name'] or not new_task_data['time']: messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog); return

            if is_edit_mode: self.tasks[index] = new_task_data; self.log(f"已修改音频节目: {new_task_data['name']}")
            else: self.tasks.append(new_task_data); self.log(f"已添加音频节目: {new_task_data['name']}")

            self.update_task_list(); self.save_tasks(); cleanup_and_destroy()

            if play_this_task_now:
                self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))

        button_text = "保存修改" if is_edit_mode else "添加"
        ttk.Button(dialog_button_frame, text=button_text, command=save_task, bootstyle="primary").pack(side=LEFT, padx=10, ipady=5)
        ttk.Button(dialog_button_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10, ipady=5)
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)
        self.center_window(dialog, parent=self.root)

    def open_playlist_editor(self, parent_dialog, initial_data):
        # 1. 创建编辑器窗口
        editor = ttk.Toplevel(parent_dialog)
        editor.title("播放列表编辑器")
        editor.geometry("700x500")
        editor.transient(parent_dialog) # 设置为父窗口的临时窗口
        editor.grab_set() # 独占焦点，实现模态
        editor.attributes('-topmost', True) # 保持在最前

        # 内部数据副本，避免直接修改原数据，点击保存后才生效
        current_playlist = list(initial_data)

        # --- UI 布局 ---
        main_layout = ttk.Frame(editor, padding=10)
        main_layout.pack(fill=BOTH, expand=True)

        # 左侧：列表区域
        list_frame = ttk.LabelFrame(main_layout, text=f"当前歌曲 ({len(current_playlist)} 首)", padding=5)
        list_frame.pack(side=LEFT, fill=BOTH, expand=True)
        
        listbox = tk.Listbox(list_frame, font=self.font_11, selectmode=EXTENDED, activestyle='none')
        listbox.pack(side=LEFT, fill=BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=VERTICAL, command=listbox.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        listbox.configure(yscrollcommand=scrollbar.set)

        # 右侧：按钮区域
        btn_frame = ttk.Frame(main_layout)
        btn_frame.pack(side=RIGHT, fill=Y, padx=(10, 0))

        # --- 功能函数 ---
        def refresh_list():
            listbox.delete(0, END)
            for path in current_playlist:
                # 在列表中只显示文件名，看起来更清爽
                listbox.insert(END, os.path.basename(path))
            list_frame.config(text=f"当前歌曲 ({len(current_playlist)} 首)")

        def add_files():
            # 1. 暂时取消置顶
            editor.attributes('-topmost', False)
            # 2. [关键新增] 强制刷新界面，确保系统应用了“取消置顶”的状态
            editor.update() 
            
            files = filedialog.askopenfilenames(
                title="添加音频文件",
                filetypes=[("Audio", "*.mp3 *.wav *.ogg *.flac *.m4a *.wma *.ape"), ("All", "*.*")],
                parent=editor
            )
            
            # 3. 恢复置顶
            editor.attributes('-topmost', True)
            # 4. [关键新增] 抢回输入焦点，防止焦点丢失
            editor.focus_force() 
            
            if files:
                for f in files:
                    current_playlist.append(f)
                refresh_list()
                listbox.see(END)

        def remove_selected():
            selected_indices = list(listbox.curselection())
            if not selected_indices: return
            # 从后往前删，避免索引偏移导致删错
            for i in reversed(selected_indices):
                del current_playlist[i]
            refresh_list()

        def move_up():
            selected = listbox.curselection()
            if not selected or len(selected) > 1: return # 仅支持单选移动
            idx = selected[0]
            if idx > 0:
                current_playlist[idx], current_playlist[idx-1] = current_playlist[idx-1], current_playlist[idx]
                refresh_list()
                listbox.selection_set(idx-1)
                listbox.see(idx-1)

        def move_down():
            selected = listbox.curselection()
            if not selected or len(selected) > 1: return
            idx = selected[0]
            if idx < len(current_playlist) - 1:
                current_playlist[idx], current_playlist[idx+1] = current_playlist[idx+1], current_playlist[idx]
                refresh_list()
                listbox.selection_set(idx+1)
                listbox.see(idx+1)

        def clear_all():
            if not current_playlist: return
            editor.attributes('-topmost', False)
            if messagebox.askyesno("确认", "确定要清空列表吗？", parent=editor):
                current_playlist.clear()
                refresh_list()
            editor.attributes('-topmost', True)

        # 按钮布局
        ttk.Button(btn_frame, text="添加文件", command=add_files, bootstyle="success").pack(fill=X, pady=5)
        ttk.Button(btn_frame, text="移除选中", command=remove_selected, bootstyle="warning").pack(fill=X, pady=5)
        ttk.Separator(btn_frame, orient=HORIZONTAL).pack(fill=X, pady=10)
        ttk.Button(btn_frame, text="上移 ↑", command=move_up, bootstyle="info-outline").pack(fill=X, pady=5)
        ttk.Button(btn_frame, text="下移 ↓", command=move_down, bootstyle="info-outline").pack(fill=X, pady=5)
        ttk.Separator(btn_frame, orient=HORIZONTAL).pack(fill=X, pady=10)
        ttk.Button(btn_frame, text="清空列表", command=clear_all, bootstyle="danger-outline").pack(fill=X, pady=5)

        # 底部确认区
        bottom_frame = ttk.Frame(editor, padding=(0, 10, 0, 0))
        bottom_frame.pack(side=BOTTOM, fill=X)
        
        result_container = [None] # 使用列表容器来存储返回值

        def on_confirm():
            result_container[0] = current_playlist
            editor.destroy()
        
        def on_cancel():
            editor.destroy()

        ttk.Button(bottom_frame, text="保存并返回", command=on_confirm, bootstyle="primary").pack(side=RIGHT, padx=10)
        ttk.Button(bottom_frame, text="取消", command=on_cancel, bootstyle="secondary").pack(side=RIGHT, padx=10)
        
        # 处理窗口关闭事件（点击X号等同于取消）
        editor.protocol("WM_DELETE_WINDOW", on_cancel)

        # 初始化列表显示
        refresh_list()
        
        # 窗口定位居中
        self.center_window(editor, parent=parent_dialog)
        
        # --- 核心：阻塞等待窗口关闭 ---
        self.root.wait_window(editor)
        
        # 窗口关闭后，重新确保父窗口置顶，防止层级混乱
        parent_dialog.attributes('-topmost', True)
        parent_dialog.focus_force()

        return result_container[0]

    def open_video_dialog(self, parent_dialog, task_to_edit=None, index=None):
        parent_dialog.destroy()
        is_edit_mode = task_to_edit is not None
        dialog = ttk.Toplevel(self.root)
        dialog.title("修改视频节目" if is_edit_mode else "添加视频节目")
        dialog.resizable(True, True)
        dialog.minsize(800, 580)
        dialog.transient(self.root)

        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()

        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)

        content_frame = ttk.LabelFrame(main_frame, text="内容", padding=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=2)
        content_frame.columnconfigure(1, weight=1)

        playback_frame = ttk.LabelFrame(main_frame, text="播放选项", padding=10)
        playback_frame.grid(row=1, column=0, sticky='ew', pady=4)

        time_frame = ttk.LabelFrame(main_frame, text="时间", padding=15)
        time_frame.grid(row=2, column=0, sticky='ew', pady=4)
        time_frame.columnconfigure(1, weight=1)

        other_frame = ttk.LabelFrame(main_frame, text="其它", padding=10)
        other_frame.grid(row=3, column=0, sticky='ew', pady=5)
        other_frame.columnconfigure(1, weight=1)

        ttk.Label(content_frame, text="节目名称:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        name_entry = ttk.Entry(content_frame, font=self.font_11)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=2)

        video_type_var = tk.StringVar(value="single")

        ttk.Label(content_frame, text="视频文件:").grid(row=1, column=0, sticky='e', padx=5, pady=2)
        video_single_frame = ttk.Frame(content_frame)
        video_single_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        video_single_frame.columnconfigure(1, weight=1)
        ttk.Radiobutton(video_single_frame, text="", variable=video_type_var, value="single").grid(row=0, column=0, sticky='w')
        video_single_entry = ttk.Entry(video_single_frame, font=self.font_11)
        video_single_entry.grid(row=0, column=1, sticky='ew', padx=5)

        def select_single_video():
            ftypes = [("视频文件", "*.mp4 *.mkv *.avi *.mov *.wmv *.flv"), ("所有文件", "*.*")]
            filename = filedialog.askopenfilename(title="选择视频文件", filetypes=ftypes, parent=dialog)
            if filename:
                video_single_entry.delete(0, END)
                video_single_entry.insert(0, filename)
        ttk.Button(video_single_frame, text="选取...", command=select_single_video, bootstyle="outline").grid(row=0, column=2, padx=5)

        ttk.Label(content_frame, text="(支持本地文件路径或网络URL地址)", font=self.font_9, bootstyle="info").grid(row=2, column=1, sticky='w', padx=5)

        ttk.Label(content_frame, text="视频文件夹:").grid(row=3, column=0, sticky='e', padx=5, pady=2)
        video_folder_frame = ttk.Frame(content_frame)
        video_folder_frame.grid(row=3, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        video_folder_frame.columnconfigure(1, weight=1)
        ttk.Radiobutton(video_folder_frame, text="", variable=video_type_var, value="folder").grid(row=0, column=0, sticky='w')
        video_folder_entry = ttk.Entry(video_folder_frame, font=self.font_11)
        video_folder_entry.grid(row=0, column=1, sticky='ew', padx=5)

        def select_folder(entry_widget):
            foldername = filedialog.askdirectory(title="选择文件夹", initialdir=application_path, parent=dialog)
            if foldername:
                entry_widget.delete(0, END)
                entry_widget.insert(0, foldername)
        ttk.Button(video_folder_frame, text="选取...", command=lambda: select_folder(video_folder_entry), bootstyle="outline").grid(row=0, column=2, padx=5)

        play_order_frame = ttk.Frame(content_frame)
        play_order_frame.grid(row=4, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        play_order_var = tk.StringVar(value="sequential")
        ttk.Radiobutton(play_order_frame, text="顺序播", variable=play_order_var, value="sequential").pack(side=LEFT, padx=10)
        ttk.Radiobutton(play_order_frame, text="随机播", variable=play_order_var, value="random").pack(side=LEFT, padx=10)
        
        ttk.Label(play_order_frame, text="音量:").pack(side=LEFT, padx=(20, 2))
        volume_entry = ttk.Entry(play_order_frame, font=self.font_11, width=5)
        volume_entry.pack(side=LEFT)
        ttk.Label(play_order_frame, text="(0-100)").pack(side=LEFT, padx=2)

        # --- ↓↓↓ 新增UI：自定义User-Agent输入框 ↓↓↓ ---
        custom_ua_var = tk.StringVar()
        ttk.Label(play_order_frame, text="自定义UA:").pack(side=LEFT, padx=(20, 2))
        ua_entry = ttk.Entry(play_order_frame, textvariable=custom_ua_var, font=self.font_11, width=25)
        ua_entry.pack(side=LEFT, fill=X, expand=True)
        # --- ↑↑↑ 新增结束 ↑↑↑ ---

        playback_mode_var = tk.StringVar(value="fullscreen")
        resolutions = ["640x480", "800x600", "1024x768", "1280x720", "1366x768", "1600x900", "1920x1080"]
        resolution_var = tk.StringVar(value=resolutions[2])

        playback_rates = ['0.5x', '0.75x', '1.0x (正常)', '1.25x', '1.5x', '2.0x']
        playback_rate_var = tk.StringVar(value='1.0x (正常)')

        mode_frame = ttk.Frame(playback_frame)
        mode_frame.grid(row=0, column=0, columnspan=3, sticky='w')

        resolution_combo = ttk.Combobox(mode_frame, textvariable=resolution_var, values=resolutions, font=self.font_11, width=12, state='readonly')

        def toggle_resolution_combo():
            if playback_mode_var.get() == "windowed":
                resolution_combo.config(state='readonly')
            else:
                resolution_combo.config(state='disabled')

        ttk.Radiobutton(mode_frame, text="无边框全屏", variable=playback_mode_var, value="fullscreen", command=toggle_resolution_combo).pack(side=LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="非全屏", variable=playback_mode_var, value="windowed", command=toggle_resolution_combo).pack(side=LEFT, padx=5)
        resolution_combo.pack(side=LEFT, padx=(5, 10))

        ttk.Label(mode_frame, text="倍速:").pack(side=LEFT)
        rate_combo = ttk.Combobox(mode_frame, textvariable=playback_rate_var, values=playback_rates, font=self.font_11, width=10)
        rate_combo.pack(side=LEFT, padx=2)
        ttk.Label(mode_frame, text="(0.25-4.0)", font=self.font_9, bootstyle="secondary").pack(side=LEFT, padx=2)

        toggle_resolution_combo()

        time_frame = ttk.LabelFrame(main_frame, text="时间", padding=15)
        time_frame.grid(row=2, column=0, sticky='ew', pady=4)
        time_frame.columnconfigure(1, weight=1)

        ttk.Label(time_frame, text="开始时间:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        start_time_entry = ttk.Entry(time_frame, font=self.font_11)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        self._bind_mousewheel_to_entry(start_time_entry, self._handle_time_scroll)
        ttk.Label(time_frame, text="<可多个>").grid(row=0, column=2, sticky='w', padx=5)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bootstyle="outline").grid(row=0, column=3, padx=5)
        
        batch_add_container = ttk.Frame(time_frame)
        batch_add_container.grid(row=0, column=4, rowspan=3, sticky='n', padx=5)

        batch_interval_frame = ttk.Frame(batch_add_container)
        batch_interval_frame.pack(pady=(0, 2))
        ttk.Label(batch_interval_frame, text="每").pack(side=LEFT)
        batch_interval_entry = ttk.Entry(batch_interval_frame, font=self.font_11, width=4)
        batch_interval_entry.pack(side=LEFT, padx=(2,2))
        ttk.Label(batch_interval_frame, text="分钟").pack(side=LEFT)

        batch_count_frame = ttk.Frame(batch_add_container)
        batch_count_frame.pack(pady=(0, 5))
        ttk.Label(batch_count_frame, text="共").pack(side=LEFT)
        batch_count_entry = ttk.Entry(batch_count_frame, font=self.font_11, width=4)
        batch_count_entry.pack(side=LEFT, padx=(2,2))
        ttk.Label(batch_count_frame, text="次   ").pack(side=LEFT)

        ttk.Button(batch_add_container, text="批量添加", 
                   command=lambda: self._apply_batch_time_addition(start_time_entry, batch_interval_entry, batch_count_entry, dialog), 
                   bootstyle="outline-info").pack(fill=X)

        interval_var = tk.StringVar(value="first")
        ttk.Label(time_frame, text="间隔播报:").grid(row=1, column=0, sticky='e', padx=5, pady=2)
        interval_frame1 = ttk.Frame(time_frame)
        interval_frame1.grid(row=1, column=1, columnspan=2, sticky='w', padx=5, pady=2)
        ttk.Radiobutton(interval_frame1, text="播 n 首", variable=interval_var, value="first").pack(side=LEFT)
        interval_first_entry = ttk.Entry(interval_frame1, font=self.font_11, width=15)
        interval_first_entry.pack(side=LEFT, padx=5)
        ttk.Label(interval_frame1, text="(单视频时,指 n 遍)").pack(side=LEFT, padx=5)

        interval_frame2 = ttk.Frame(time_frame)
        interval_frame2.grid(row=2, column=1, columnspan=2, sticky='w', padx=5, pady=2)
        ttk.Radiobutton(interval_frame2, text="播 n 秒", variable=interval_var, value="seconds").pack(side=LEFT)
        interval_seconds_entry = ttk.Entry(interval_frame2, font=self.font_11, width=15)
        interval_seconds_entry.pack(side=LEFT, padx=5)
        ttk.Label(interval_frame2, text="(3600秒 = 1小时)").pack(side=LEFT, padx=5)

        ttk.Label(time_frame, text="周几/几号:").grid(row=3, column=0, sticky='e', padx=5, pady=3)
        weekday_entry = ttk.Entry(time_frame, font=self.font_11)
        weekday_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=3)
        ttk.Button(time_frame, text="选取...", command=lambda: self.show_weekday_settings_dialog(weekday_entry), bootstyle="outline").grid(row=3, column=3, padx=5)

        ttk.Label(time_frame, text="日期范围:").grid(row=4, column=0, sticky='e', padx=5, pady=3)
        date_range_entry = ttk.Entry(time_frame, font=self.font_11)
        date_range_entry.grid(row=4, column=1, sticky='ew', padx=5, pady=3)
        self._bind_mousewheel_to_entry(date_range_entry, self._handle_date_scroll)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry), bootstyle="outline").grid(row=4, column=3, padx=5)

        other_frame = ttk.LabelFrame(main_frame, text="其它", padding=10)
        other_frame.grid(row=3, column=0, sticky='ew', pady=5)
        other_frame.columnconfigure(1, weight=1)

        delay_var = tk.StringVar(value="ontime")
        ttk.Label(other_frame, text="模式:").grid(row=0, column=0, sticky='nw', padx=5, pady=2)
        delay_frame = ttk.Frame(other_frame)
        delay_frame.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        ttk.Radiobutton(delay_frame, text="准时播 - 如果有别的节目正在播，终止他们（默认）", variable=delay_var, value="ontime").pack(anchor='w')
        ttk.Radiobutton(delay_frame, text="可延后 - 如果有别的节目正在播，排队等候", variable=delay_var, value="delay").pack(anchor='w')
        ttk.Radiobutton(delay_frame, text="立即播 - 添加后停止其他节目,立即播放此节目", variable=delay_var, value="immediate").pack(anchor='w')

        dialog_button_frame = ttk.Frame(other_frame)
        dialog_button_frame.grid(row=0, column=2, sticky='se', padx=20, pady=10)

        if is_edit_mode:
            task = task_to_edit
            name_entry.insert(0, task.get('name', ''))
            video_type_var.set(task.get('video_type', 'single'))
            if task.get('video_type') == 'single':
                video_single_entry.insert(0, task.get('content', ''))
            else:
                video_folder_entry.insert(0, task.get('content', ''))
            play_order_var.set(task.get('play_order', 'sequential'))
            volume_entry.insert(0, task.get('volume', '80'))
            # --- ↓↓↓ 新增：加载自定义UA ↓↓↓ ---
            custom_ua_var.set(task.get('custom_user_agent', ''))
            # --- ↑↑↑ 新增结束 ↑↑↑ ---
            playback_mode_var.set(task.get('playback_mode', 'fullscreen'))
            resolution_var.set(task.get('resolution', '1024x768'))
            playback_rate_var.set(task.get('playback_rate', '1.0x (正常)'))
            start_time_entry.insert(0, task.get('time', ''))
            interval_var.set(task.get('interval_type', 'first'))
            interval_first_entry.insert(0, task.get('interval_first', '1'))
            interval_seconds_entry.insert(0, task.get('interval_seconds', '600'))
            weekday_entry.insert(0, task.get('weekday', '每周:1234567'))
            date_range_entry.insert(0, task.get('date_range', '2025-01-01 ~ 2099-12-31'))
            delay_var.set(task.get('delay', 'ontime'))
            toggle_resolution_combo()
        else:
            volume_entry.insert(0, "80")
            interval_first_entry.insert(0, "1")
            interval_seconds_entry.insert(0, "600")
            weekday_entry.insert(0, "每周:1234567")
            date_range_entry.insert(0, "2025-01-01 ~ 2099-12-31")

        def save_task():
            try:
                volume = int(volume_entry.get().strip() or 80)
                if not (0 <= volume <= 100):
                    messagebox.showerror("输入错误", "音量必须是 0 到 100 之间的整数。", parent=dialog)
                    return
            except ValueError:
                messagebox.showerror("输入错误", "音量必须是一个有效的整数。", parent=dialog)
                return

            if interval_var.get() == 'first':
                try:
                    interval_first = int(interval_first_entry.get().strip() or 1)
                    if interval_first < 1:
                        messagebox.showerror("输入错误", "“播 n 首”的次数必须大于或等于 1。", parent=dialog)
                        return
                except ValueError:
                    messagebox.showerror("输入错误", "“播 n 首”的次数必须是一个有效的整数。", parent=dialog)
                    return
            else: 
                try:
                    interval_seconds = int(interval_seconds_entry.get().strip() or 1)
                    if interval_seconds < 1:
                        messagebox.showerror("输入错误", "“播 n 秒”的秒数必须大于或等于 1。", parent=dialog)
                        return
                except ValueError:
                    messagebox.showerror("输入错误", "“播 n 秒”的秒数必须是一个有效的整数。", parent=dialog)
                    return

            if not weekday_entry.get().strip():
                messagebox.showerror("输入错误", "“周几/几号”规则不能为空，请点击“选取...”进行设置。", parent=dialog)
                return
            
            if not date_range_entry.get().strip():
                messagebox.showerror("输入错误", "“日期范围”不能为空，请点击“设置...”进行配置。", parent=dialog)
                return
            
            video_path = video_single_entry.get().strip() if video_type_var.get() == "single" else video_folder_entry.get().strip()
            
            is_url = video_path.lower().startswith(('http://', 'https://', 'rtsp://', 'rtmp://', 'mms://'))
            
            if not video_path:
                messagebox.showwarning("警告", "请选择一个视频文件/文件夹，或输入一个网络地址", parent=dialog)
                return
            if not is_url and not os.path.exists(video_path):
                 messagebox.showwarning("警告", "本地文件或文件夹路径不存在，请重新选择。", parent=dialog)
                 return
            
            is_valid_time, time_msg = self._normalize_multiple_times_string(start_time_entry.get().strip())
            if not is_valid_time: messagebox.showwarning("格式错误", time_msg, parent=dialog); return
            is_valid_date, date_msg = self._normalize_date_range_string(date_range_entry.get().strip())
            if not is_valid_date: messagebox.showwarning("格式错误", date_msg, parent=dialog); return

            rate_input = playback_rate_var.get().strip()
            rate_match = re.match(r"(\d+(\.\d+)?)", rate_input)
            if not rate_match:
                messagebox.showwarning("输入错误", "无效的播放倍速值。", parent=dialog)
                return
            rate_str = rate_match.group(1)

            try:
                rate_val = float(rate_str)
                if not (0.25 <= rate_val <= 4.0):
                    messagebox.showwarning("输入错误", "播放倍速必须在 0.25 和 4.0 之间。", parent=dialog)
                    return
            except ValueError:
                messagebox.showwarning("输入错误", "无效的播放倍速值。", parent=dialog)
                return

            play_mode = delay_var.get()
            play_this_task_now = (play_mode == 'immediate')
            saved_delay_type = 'ontime' if play_mode == 'immediate' else play_mode

            task_name = name_entry.get().strip()
            if not task_name and not is_url:
                task_name = os.path.basename(video_path)

            new_task_data = {
                'name': task_name,
                'time': time_msg,
                'content': video_path,
                'type': 'video',
                'video_type': video_type_var.get(),
                'play_order': play_order_var.get(),
                'volume': str(volume),
                'interval_type': interval_var.get(),
                'interval_first': interval_first_entry.get().strip() or "1",
                'interval_seconds': interval_seconds_entry.get().strip() or "600",
                'playback_mode': playback_mode_var.get(),
                'resolution': resolution_var.get(),
                'playback_rate': rate_input,
                # --- ↓↓↓ 新增：保存自定义UA ↓↓↓ ---
                'custom_user_agent': custom_ua_var.get().strip(),
                # --- ↑↑↑ 新增结束 ↑↑↑ ---
                'weekday': weekday_entry.get().strip(),
                'date_range': date_msg,
                'delay': saved_delay_type,
                'status': '启用' if not is_edit_mode else task_to_edit.get('status', '启用'),
                'last_run': {} if not is_edit_mode else task_to_edit.get('last_run', {}),
            }
            if not new_task_data['name'] or not new_task_data['time']:
                messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog)
                return

            if is_edit_mode:
                self.tasks[index] = new_task_data
                self.log(f"已修改视频节目: {new_task_data['name']}")
            else:
                self.tasks.append(new_task_data)
                self.log(f"已添加视频节目: {new_task_data['name']}")

            self.update_task_list()
            self.save_tasks()
            cleanup_and_destroy()

            if play_this_task_now:
                self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))

        button_text = "保存修改" if is_edit_mode else "添加"
        ttk.Button(dialog_button_frame, text=button_text, command=save_task, bootstyle="primary").pack(side=LEFT, padx=10, ipady=5)
        ttk.Button(dialog_button_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10, ipady=5)
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)
        self.center_window(dialog, parent=self.root)

#第6部分
    def open_voice_dialog(self, parent_dialog, task_to_edit=None, index=None):
        parent_dialog.destroy()
        is_edit_mode = task_to_edit is not None
        dialog = ttk.Toplevel(self.root)
        dialog.title("修改语音节目" if is_edit_mode else "添加语音节目")
        dialog.resizable(True, True)
        dialog.minsize(800, 580)
        dialog.transient(self.root)

        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()

        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)

        content_frame = ttk.LabelFrame(main_frame, text="内容", padding=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=2)
        content_frame.columnconfigure(1, weight=1)

        ttk.Label(content_frame, text="节目名称:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        name_entry = ttk.Entry(content_frame, font=self.font_11)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        
        ttk.Label(content_frame, text="播音文字:").grid(row=1, column=0, sticky='nw', padx=5, pady=2)
        text_frame = ttk.Frame(content_frame)
        text_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        content_text = ScrolledText(text_frame, height=3, font=self.font_11, wrap=WORD)
        content_text.grid(row=0, column=0, sticky='nsew')
        
        script_btn_frame = ttk.Frame(content_frame)
        script_btn_frame.grid(row=2, column=1, columnspan=3, sticky='w', padx=5, pady=(0, 2))
        ttk.Button(script_btn_frame, text="导入文稿", command=lambda: self._import_voice_script(content_text, dialog), bootstyle="outline").pack(side=LEFT)
        ttk.Button(script_btn_frame, text="导出文稿", command=lambda: self._export_voice_script(content_text, name_entry, dialog), bootstyle="outline").pack(side=LEFT, padx=10)

        ad_btn_frame = ttk.Frame(script_btn_frame)
        ad_btn_frame.pack(side=LEFT, padx=20)

        self.ad_by_voice_btn = ttk.Button(ad_btn_frame, text="按语音长度制作广告", 
                                          command=lambda: self._create_advertisement('voice'))
        self.ad_by_voice_btn.pack(side=LEFT)

        self.ad_by_bgm_btn = ttk.Button(ad_btn_frame, text="按背景音乐长度制作广告", 
                                        command=lambda: self._create_advertisement('bgm'))
        self.ad_by_bgm_btn.pack(side=LEFT, padx=10)

        if self.auth_info['status'] != 'Permanent':
            self.ad_by_voice_btn.config(state=DISABLED)
            self.ad_by_bgm_btn.config(state=DISABLED)

        ttk.Label(content_frame, text="引擎类型:").grid(row=3, column=0, sticky='w', padx=5, pady=3)
        engine_frame = ttk.Frame(content_frame)
        engine_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5)
        
        voice_engine_var = tk.StringVar(value="local")
        
        local_rb = ttk.Radiobutton(engine_frame, text="本地语音 (SAPI)", variable=voice_engine_var, value="local")
        local_rb.pack(side=LEFT, padx=(0, 15))
        
        online_rb = ttk.Radiobutton(engine_frame, text="在线语音 (推荐)", variable=voice_engine_var, value="online")
        online_rb.pack(side=LEFT)
        
        ttk.Label(content_frame, text="播音员:").grid(row=4, column=0, sticky='w', padx=5, pady=3)
        voice_frame = ttk.Frame(content_frame)
        voice_frame.grid(row=4, column=1, columnspan=3, sticky='ew', padx=5, pady=3)
        voice_frame.columnconfigure(0, weight=1)
        
        voice_var = tk.StringVar()
        voice_combo = ttk.Combobox(voice_frame, textvariable=voice_var, font=self.font_11, state='readonly')
        voice_combo.grid(row=0, column=0, sticky='ew')
        
        def _update_voice_options(*args):
            engine = voice_engine_var.get()
            current_voice = voice_var.get()
            
            if engine == "local":
                available_voices = self.get_available_voices()
                voice_combo['values'] = available_voices
                if available_voices:
                    if current_voice in available_voices:
                        voice_var.set(current_voice)
                    else:
                        voice_var.set(available_voices[0])
                else:
                    voice_var.set("")
            else: 
                available_voices = list(EDGE_TTS_VOICES.keys())
                voice_combo['values'] = available_voices
                if available_voices:
                    if current_voice in available_voices:
                        voice_var.set(current_voice)
                    else:
                        voice_var.set(available_voices[0])
                else:
                    voice_var.set("")

        voice_engine_var.trace_add("write", _update_voice_options)
        
        speech_params_frame = ttk.Frame(voice_frame)
        speech_params_frame.grid(row=0, column=1, sticky='e', padx=(10, 0))

        ttk.Label(speech_params_frame, text="语速:").pack(side=LEFT)
        speed_entry = ttk.Entry(speech_params_frame, font=self.font_11, width=5); speed_entry.pack(side=LEFT, padx=(2, 5))
        ttk.Label(speech_params_frame, text="音调:").pack(side=LEFT)
        pitch_entry = ttk.Entry(speech_params_frame, font=self.font_11, width=5); pitch_entry.pack(side=LEFT, padx=(2, 5))
        ttk.Label(speech_params_frame, text="音量:").pack(side=LEFT)
        volume_entry = ttk.Entry(speech_params_frame, font=self.font_11, width=5); volume_entry.pack(side=LEFT, padx=(2, 0))

        prompt_var = tk.IntVar(); prompt_frame = ttk.Frame(content_frame)
        prompt_frame.grid(row=5, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        prompt_frame.columnconfigure(1, weight=1)
        ttk.Checkbutton(prompt_frame, text="提示音:", variable=prompt_var, bootstyle="round-toggle").grid(row=0, column=0, sticky='w')
        prompt_file_var, prompt_volume_var = tk.StringVar(), tk.StringVar()
        prompt_file_entry = ttk.Entry(prompt_frame, textvariable=prompt_file_var, font=self.font_11); prompt_file_entry.grid(row=0, column=1, sticky='ew', padx=5)
        ttk.Button(prompt_frame, text="...", command=lambda: self.select_file_for_entry(PROMPT_FOLDER, prompt_file_var, dialog), bootstyle="outline", width=2).grid(row=0, column=2)
        
        prompt_vol_frame = ttk.Frame(prompt_frame)
        prompt_vol_frame.grid(row=0, column=3, sticky='e')
        ttk.Label(prompt_vol_frame, text="音量(0-100):").pack(side=LEFT, padx=(10,5))
        ttk.Entry(prompt_vol_frame, textvariable=prompt_volume_var, font=self.font_11, width=8).pack(side=LEFT, padx=5)
        
        bgm_var = tk.IntVar(); bgm_frame = ttk.Frame(content_frame)
        bgm_frame.grid(row=6, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        bgm_frame.columnconfigure(1, weight=1)
        ttk.Checkbutton(bgm_frame, text="背景音乐:", variable=bgm_var, bootstyle="round-toggle").grid(row=0, column=0, sticky='w')
        bgm_file_var, bgm_volume_var = tk.StringVar(), tk.StringVar()
        bgm_file_entry = ttk.Entry(bgm_frame, textvariable=bgm_file_var, font=self.font_11); bgm_file_entry.grid(row=0, column=1, sticky='ew', padx=5)
        ttk.Button(bgm_frame, text="...", command=lambda: self.select_file_for_entry(BGM_FOLDER, bgm_file_var, dialog), bootstyle="outline", width=2).grid(row=0, column=2)
        
        bgm_vol_frame = ttk.Frame(bgm_frame)
        bgm_vol_frame.grid(row=0, column=3, sticky='e')
        ttk.Label(bgm_vol_frame, text="音量(0-100):").pack(side=LEFT, padx=(10,5))
        ttk.Entry(bgm_vol_frame, textvariable=bgm_volume_var, font=self.font_11, width=8).pack(side=LEFT, padx=5)

        bg_image_var = tk.IntVar(value=0)
        bg_image_path_var = tk.StringVar()
        bg_image_order_var = tk.StringVar(value="sequential")

        bg_image_frame = ttk.Frame(content_frame)
        bg_image_frame.grid(row=7, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        bg_image_frame.columnconfigure(1, weight=1)
        bg_image_cb = ttk.Checkbutton(bg_image_frame, text="背景图片:", variable=bg_image_var, bootstyle="round-toggle")
        bg_image_cb.grid(row=0, column=0, sticky='w')
        if not IMAGE_AVAILABLE: bg_image_cb.config(state=DISABLED, text="背景图片(Pillow未安装):")

        bg_image_entry = ttk.Entry(bg_image_frame, textvariable=bg_image_path_var, font=self.font_11)
        bg_image_entry.grid(row=0, column=1, sticky='ew', padx=5)
        
        bg_image_btn_frame = ttk.Frame(bg_image_frame)
        bg_image_btn_frame.grid(row=0, column=2, sticky='e')
        def select_folder(entry_widget):
            foldername = filedialog.askdirectory(title="选择文件夹", initialdir=application_path, parent=dialog)
            if foldername: entry_widget.delete(0, END); entry_widget.insert(0, foldername)
        ttk.Button(bg_image_btn_frame, text="选取...", command=lambda: select_folder(bg_image_entry), bootstyle="outline").pack(side=LEFT, padx=5)
        ttk.Radiobutton(bg_image_btn_frame, text="顺序", variable=bg_image_order_var, value="sequential").pack(side=LEFT, padx=(10,0))
        ttk.Radiobutton(bg_image_btn_frame, text="随机", variable=bg_image_order_var, value="random").pack(side=LEFT)

        time_frame = ttk.LabelFrame(main_frame, text="时间", padding=10)
        time_frame.grid(row=1, column=0, sticky='ew', pady=2)
        time_frame.columnconfigure(1, weight=1)
        
        ttk.Label(time_frame, text="开始时间:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        start_time_entry = ttk.Entry(time_frame, font=self.font_11)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        self._bind_mousewheel_to_entry(start_time_entry, self._handle_time_scroll)
        ttk.Label(time_frame, text="<可多个>").grid(row=0, column=2, sticky='w', padx=5)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bootstyle="outline").grid(row=0, column=3, padx=5)
        
        batch_add_container = ttk.Frame(time_frame)
        batch_add_container.grid(row=0, column=4, rowspan=3, sticky='n', padx=5)

        batch_interval_frame = ttk.Frame(batch_add_container)
        batch_interval_frame.pack(pady=(0, 2))
        ttk.Label(batch_interval_frame, text="每").pack(side=LEFT)
        batch_interval_entry = ttk.Entry(batch_interval_frame, font=self.font_11, width=4)
        batch_interval_entry.pack(side=LEFT, padx=(2,2))
        ttk.Label(batch_interval_frame, text="分钟").pack(side=LEFT)

        batch_count_frame = ttk.Frame(batch_add_container)
        batch_count_frame.pack(pady=(0, 5))
        ttk.Label(batch_count_frame, text="共").pack(side=LEFT)
        batch_count_entry = ttk.Entry(batch_count_frame, font=self.font_11, width=4)
        batch_count_entry.pack(side=LEFT, padx=(2,2))
        ttk.Label(batch_count_frame, text="次   ").pack(side=LEFT)

        ttk.Button(batch_add_container, text="批量添加", 
                   command=lambda: self._apply_batch_time_addition(start_time_entry, batch_interval_entry, batch_count_entry, dialog), 
                   bootstyle="outline-info").pack(fill=X)

        ttk.Label(time_frame, text="播 n 遍:").grid(row=1, column=0, sticky='e', padx=5, pady=2)
        repeat_entry = ttk.Entry(time_frame, font=self.font_11, width=12)
        repeat_entry.grid(row=1, column=1, sticky='w', padx=5, pady=2)
        
        ttk.Label(time_frame, text="周几/几号:").grid(row=2, column=0, sticky='e', padx=5, pady=2)
        weekday_entry = ttk.Entry(time_frame, font=self.font_11)
        weekday_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=2)
        ttk.Button(time_frame, text="选取...", command=lambda: self.show_weekday_settings_dialog(weekday_entry), bootstyle="outline").grid(row=2, column=3, padx=5)
        
        ttk.Label(time_frame, text="日期范围:").grid(row=3, column=0, sticky='e', padx=5, pady=2)
        date_range_entry = ttk.Entry(time_frame, font=self.font_11)
        date_range_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=2)
        self._bind_mousewheel_to_entry(date_range_entry, self._handle_date_scroll)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry), bootstyle="outline").grid(row=3, column=3, padx=5)

        other_frame = ttk.LabelFrame(main_frame, text="其它", padding=15)
        other_frame.grid(row=2, column=0, sticky='ew', pady=4)
        other_frame.columnconfigure(1, weight=1)
        
        delay_var = tk.StringVar(value="delay")
        ttk.Label(other_frame, text="模式:").grid(row=0, column=0, sticky='nw', padx=5, pady=2)
        delay_frame = ttk.Frame(other_frame)
        delay_frame.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        ttk.Radiobutton(delay_frame, text="准时播 - 如果有别的节目正在播，终止他们", variable=delay_var, value="ontime").pack(anchor='w', pady=1)
        ttk.Radiobutton(delay_frame, text="可延后 - 如果有别的节目正在播，排队等候（默认）", variable=delay_var, value="delay").pack(anchor='w', pady=1)
        ttk.Radiobutton(delay_frame, text="立即播 - 添加后停止其他节目,立即播放此节目", variable=delay_var, value="immediate").pack(anchor='w', pady=1)
        
        dialog_button_frame = ttk.Frame(other_frame)
        dialog_button_frame.grid(row=0, column=2, sticky='se', padx=20, pady=10)

        if is_edit_mode:
            task = task_to_edit
            name_entry.insert(0, task.get('name', ''))
            content_text.insert('1.0', task.get('source_text', ''))

            saved_voice = task.get('voice', '')
            if saved_voice in EDGE_TTS_VOICES:
                voice_engine_var.set("online")
            else:
                voice_engine_var.set("local")
            
            _update_voice_options()
            voice_var.set(saved_voice)

            speed_entry.insert(0, task.get('speed', '0'))
            pitch_entry.insert(0, task.get('pitch', '0'))
            volume_entry.insert(0, task.get('volume', '80'))
            prompt_var.set(task.get('prompt', 0)); prompt_file_var.set(task.get('prompt_file', '')); prompt_volume_var.set(task.get('prompt_volume', '80'))
            bgm_var.set(task.get('bgm', 0)); bgm_file_var.set(task.get('bgm_file', '')); bgm_volume_var.set(task.get('bgm_volume', '20'))
            start_time_entry.insert(0, task.get('time', ''))
            repeat_entry.insert(0, task.get('repeat', '1'))
            weekday_entry.insert(0, task.get('weekday', '每周:1234567'))
            date_range_entry.insert(0, task.get('date_range', '2025-01-01 ~ 2099-12-31'))
            delay_var.set(task.get('delay', 'delay'))
            bg_image_var.set(task.get('bg_image_enabled', 0))
            bg_image_path_var.set(task.get('bg_image_path', ''))
            bg_image_order_var.set(task.get('bg_image_order', 'sequential'))
        else:
            _update_voice_options()
            speed_entry.insert(0, "0"); pitch_entry.insert(0, "0"); volume_entry.insert(0, "80")
            prompt_var.set(0); prompt_volume_var.set("80"); bgm_var.set(0); bgm_volume_var.set("20")
            repeat_entry.insert(0, "1"); weekday_entry.insert(0, "每周:1234567"); date_range_entry.insert(0, "2025-01-01 ~ 2099-12-31")

        ad_params = {
            'dialog': dialog, 'name_entry': name_entry, 'content_text': content_text,
            'voice_var': voice_var, 'speed_entry': speed_entry, 'pitch_entry': pitch_entry,
            'volume_entry': volume_entry, 'prompt_var': prompt_var,
            'prompt_file_var': prompt_file_var, 'prompt_volume_var': prompt_volume_var,
            'bgm_var': bgm_var, 'bgm_file_var': bgm_file_var, 'bgm_volume_var': bgm_volume_var,
            # --- ↓↓↓ 新增：将引擎选择也传递给广告制作函数 ---
            'voice_engine_var': voice_engine_var,
        }

        self.ad_by_voice_btn.config(command=lambda: self._create_advertisement('voice', ad_params))
        self.ad_by_bgm_btn.config(command=lambda: self._create_advertisement('bgm', ad_params))

        def save_task():
            try:
                speed = int(speed_entry.get().strip() or '0')
                pitch = int(pitch_entry.get().strip() or '0')
                volume = int(volume_entry.get().strip() or '80')
                repeat = int(repeat_entry.get().strip() or '1')
                if not (-10 <= speed <= 10): messagebox.showerror("输入错误", "语速必须在 -10 到 10 之间。", parent=dialog); return
                if not (-10 <= pitch <= 10): messagebox.showerror("输入错误", "音调必须在 -10 到 10 之间。", parent=dialog); return
                if not (0 <= volume <= 100): messagebox.showerror("输入错误", "音量必须在 0 到 100 之间。", parent=dialog); return
                if repeat < 1: messagebox.showerror("输入错误", "“播 n 遍”的次数必须大于或等于 1。", parent=dialog); return
            except ValueError: messagebox.showerror("输入错误", "语速、音调、音量、播报遍数必须是有效的整数。", parent=dialog); return
            if not weekday_entry.get().strip(): messagebox.showerror("输入错误", "“周几/几号”规则不能为空...", parent=dialog); return
            if not date_range_entry.get().strip(): messagebox.showerror("输入错误", "“日期范围”不能为空...", parent=dialog); return
            
            text_content = content_text.get('1.0', END).strip()
            if not text_content: messagebox.showwarning("警告", "请输入播音文字内容", parent=dialog); return
            is_valid_time, time_msg = self._normalize_multiple_times_string(start_time_entry.get().strip())
            if not is_valid_time: messagebox.showwarning("格式错误", time_msg, parent=dialog); return
            is_valid_date, date_msg = self._normalize_date_range_string(date_range_entry.get().strip())
            if not is_valid_date: messagebox.showwarning("格式错误", date_msg, parent=dialog); return
            
            regeneration_needed = True
            selected_voice = voice_var.get()
            is_online_voice = voice_engine_var.get() == 'online'

            if is_edit_mode:
                original_task = task_to_edit
                is_original_online = original_task.get('voice', '') in EDGE_TTS_VOICES
                
                if (text_content == original_task.get('source_text') and
                    selected_voice == original_task.get('voice') and
                    speed_entry.get().strip() == original_task.get('speed', '0') and
                    pitch_entry.get().strip() == original_task.get('pitch', '0') and
                    is_online_voice == is_original_online):
                    if not is_online_voice and volume_entry.get().strip() == original_task.get('volume', '80'):
                        regeneration_needed = False
                    elif is_online_voice:
                        regeneration_needed = False
                
                if not regeneration_needed: self.log("语音内容未变更，跳过重新生成音频文件。")

            def build_task_data(audio_path, audio_filename_str):
                play_mode = delay_var.get()
                play_this_task_now = (play_mode == 'immediate')
                saved_delay_type = 'delay' if play_mode == 'immediate' else play_mode

                return {
                    'name': name_entry.get().strip(), 'time': time_msg, 'type': 'voice', 'content': audio_path,
                    'wav_filename': audio_filename_str, 'source_text': text_content, 'voice': voice_var.get(),
                    'speed': speed_entry.get().strip() or "0", 'pitch': pitch_entry.get().strip() or "0",
                    'volume': volume_entry.get().strip() or "80", 'prompt': prompt_var.get(),
                    'prompt_file': prompt_file_var.get(), 'prompt_volume': prompt_volume_var.get(),
                    'bgm': bgm_var.get(), 'bgm_file': bgm_file_var.get(), 'bgm_volume': bgm_volume_var.get(),
                    'repeat': repeat_entry.get().strip() or "1", 'weekday': weekday_entry.get().strip(),
                    'date_range': date_msg, 'delay': saved_delay_type,
                    'status': '启用' if not is_edit_mode else task_to_edit.get('status', '启用'),
                    'last_run': {} if not is_edit_mode else task_to_edit.get('last_run', {}),
                    'bg_image_enabled': bg_image_var.get(),
                    'bg_image_path': bg_image_path_var.get().strip(),
                    'bg_image_order': bg_image_order_var.get()
                }, play_this_task_now

            if not regeneration_needed:
                new_task_data, play_now_flag = build_task_data(task_to_edit.get('content'), task_to_edit.get('wav_filename'))
                if not new_task_data['name'] or not new_task_data['time']: messagebox.showwarning("警告", "请填写必要信息...", parent=dialog); return
                self.tasks[index] = new_task_data; self.log(f"已修改语音节目(未重新生成语音): {new_task_data['name']}")
                self.update_task_list(); self.save_tasks(); cleanup_and_destroy()
                if play_now_flag: self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))
                return

            progress_dialog = ttk.Toplevel(dialog)
            progress_dialog.title("请稍候")
            progress_dialog.resizable(False, False); progress_dialog.transient(dialog)
            
            progress_dialog.attributes('-topmost', True)
            dialog.attributes('-disabled', True)
            
            def cleanup_progress():
                dialog.attributes('-disabled', False)
                progress_dialog.destroy()
                dialog.focus_force()

            progress_dialog.protocol("WM_DELETE_WINDOW", cleanup_progress)
            ttk.Label(progress_dialog, text="语音文件生成中，请稍后...", font=self.font_11).pack(expand=True, padx=20, pady=20)
            self.center_window(progress_dialog, parent=dialog)
            
            if is_online_voice:
                new_audio_filename = f"{int(time.time())}_{random.randint(1000, 9999)}.mp3"
            else:
                new_audio_filename = f"{int(time.time())}_{random.randint(1000, 9999)}.wav"
            
            output_path = os.path.join(AUDIO_FOLDER, new_audio_filename)
            voice_params = {
                'voice': voice_var.get(), 'speed': speed_entry.get().strip() or "0", 
                'pitch': pitch_entry.get().strip() or "0", 'volume': volume_entry.get().strip() or "80"
            }
            
            def _on_synthesis_complete(result):
                cleanup_progress()
                if not result['success']: messagebox.showerror("错误", f"无法生成语音文件: {result['error']}", parent=dialog); return
                if is_edit_mode and 'wav_filename' in task_to_edit:
                    old_audio_path = os.path.join(AUDIO_FOLDER, task_to_edit['wav_filename'])
                    if os.path.exists(old_audio_path):
                        try: os.remove(old_audio_path); self.log(f"已删除旧语音文件: {task_to_edit['wav_filename']}")
                        except Exception as e: self.log(f"删除旧语音文件失败: {e}")
                
                new_task_data, play_now_flag = build_task_data(output_path, new_audio_filename)
                if not new_task_data['name'] or not new_task_data['time']: messagebox.showwarning("警告", "请填写必要信息...", parent=dialog); return
                if is_edit_mode: self.tasks[index] = new_task_data; self.log(f"已修改语音节目(并重新生成语音): {new_task_data['name']}")
                else: self.tasks.append(new_task_data); self.log(f"已添加语音节目: {new_task_data['name']}")
                self.update_task_list(); self.save_tasks(); cleanup_and_destroy()
                if play_now_flag: self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))
            
            # --- ↓↓↓ 核心逻辑：根据引擎选择不同的工作线程 ↓↓↓ ---
            if is_online_voice:
                s_thread = threading.Thread(target=self._synthesis_worker_edge, args=(text_content, voice_params, output_path, _on_synthesis_complete))
            else:
                s_thread = threading.Thread(target=self._synthesis_worker, args=(text_content, voice_params, output_path, _on_synthesis_complete))
            
            s_thread.daemon = True
            s_thread.start()
            # --- ↑↑↑ 核心逻辑结束 ↑↑↑ ---

        button_text = "保存修改" if is_edit_mode else "添加"
        ttk.Button(dialog_button_frame, text=button_text, command=save_task, bootstyle="primary").pack(side=LEFT, padx=10, ipady=5)
        ttk.Button(dialog_button_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10, ipady=5)
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)
        self.center_window(dialog, parent=self.root)

    def _create_advertisement(self, mode, params):
        try:
            from pydub import AudioSegment
            
            ffmpeg_path = os.path.join(application_path, "ffmpeg.exe")

            if not os.path.exists(ffmpeg_path):
                messagebox.showerror("依赖缺失", 
                                     "错误：未在软件根目录找到 ffmpeg.exe。\n\n"
                                     "请下载 FFmpeg，并将其中的 ffmpeg.exe 文件放置到本软件所在的文件夹内，然后重试。",
                                     parent=params['dialog'])
                return

            AudioSegment.converter = ffmpeg_path
        except ImportError:
            messagebox.showerror("依赖缺失", "错误: pydub 库未安装，无法使用此功能。", parent=params['dialog'])
            return
        except Exception as e:
            messagebox.showerror("初始化失败", f"加载音频处理组件时出错: {e}", parent=params['dialog'])
            return

        if not params['bgm_var'].get() or not params['bgm_file_var'].get().strip():
            messagebox.showerror("错误", "必须选择背景音乐才能制作广告。", parent=params['dialog']); return
        bgm_path = params['bgm_file_var'].get().strip()
        if not os.path.exists(bgm_path):
            messagebox.showerror("错误", f"背景音乐文件不存在：\n{bgm_path}", parent=params['dialog']); return
        text_content = params['content_text'].get('1.0', 'end').strip()
        if not text_content:
            messagebox.showerror("错误", "播音文字内容不能为空。", parent=params['dialog']); return
        try:
            voice_volume = int(params['volume_entry'].get().strip() or '80')
            bgm_volume = int(params['bgm_volume_var'].get().strip() or '20')
        except ValueError:
            messagebox.showerror("错误", "音量必须是有效的整数。", parent=params['dialog']); return

        progress_dialog = ttk.Toplevel(params['dialog'])
        progress_dialog.title("正在制作广告")
        progress_dialog.resizable(False, False)
        progress_dialog.transient(params['dialog'])
        
        progress_dialog.attributes('-topmost', True)
        params['dialog'].attributes('-disabled', True)
        
        def cleanup_progress():
            params['dialog'].attributes('-disabled', False)
            progress_dialog.destroy()
            params['dialog'].focus_force()

        progress_dialog.protocol("WM_DELETE_WINDOW", cleanup_progress)

        progress_label = ttk.Label(progress_dialog, text="正在准备...", font=self.font_11)
        progress_label.pack(pady=10, padx=20)
        progress = ttk.Progressbar(progress_dialog, length=300, mode='determinate')
        progress.pack(pady=10, padx=20)
        self.center_window(progress_dialog, parent=params['dialog'])

        # --- ↓↓↓ 核心修改区域：适配在线/离线语音生成 ↓↓↓ ---
        def worker():
            temp_audio_path = None # 可以是 .wav 或 .mp3
            try:
                self.root.after(0, lambda: progress_label.config(text="步骤1/4: 生成语音..."))
                self.root.after(0, lambda: progress.config(value=10))

                voice_params = {
                    'voice': params['voice_var'].get(),
                    'speed': params['speed_entry'].get().strip() or "0",
                    'pitch': params['pitch_entry'].get().strip() or "0",
                    'volume': '100' # 语音合成时总是用最大音量，混合时再调整
                }

                is_online_engine = params['voice_engine_var'].get() == 'online'
                
                # 根据引擎选择不同的文件名和生成函数
                if is_online_engine:
                    temp_audio_filename = f"temp_ad_{int(time.time())}.mp3"
                    temp_audio_path = os.path.join(AUDIO_FOLDER, temp_audio_filename)
                    
                    # 使用阻塞方式调用在线合成（在新线程中），等待其完成
                    synthesis_success = threading.Event()
                    error_message = ""
                    def online_callback(result):
                        nonlocal error_message
                        if result['success']:
                            synthesis_success.set()
                        else:
                            error_message = result.get('error', '未知在线合成错误')
                            synthesis_success.set()
                    
                    s_thread = threading.Thread(target=self._synthesis_worker_edge, args=(text_content, voice_params, temp_audio_path, online_callback))
                    s_thread.start()
                    s_thread.join() # 等待在线合成线程结束
                    
                    if error_message:
                        raise Exception(f"在线语音合成失败: {error_message}")

                else: # 本地 SAPI 引擎
                    temp_audio_filename = f"temp_ad_{int(time.time())}.wav"
                    temp_audio_path = os.path.join(AUDIO_FOLDER, temp_audio_filename)
                    if not self._synthesize_text_to_wav(text_content, voice_params, temp_audio_path):
                        raise Exception("本地语音合成失败！")

                self.root.after(0, lambda: progress_label.config(text="步骤2/4: 分析音频..."))
                self.root.after(0, lambda: progress.config(value=30))
                
                # pydub 可以自动处理 wav 和 mp3
                voice_audio = AudioSegment.from_file(temp_audio_path)
                bgm_audio = AudioSegment.from_file(bgm_path)
        # --- ↑↑↑ 核心修改区域结束 ↑↑↑ ---

                voice_duration_ms = len(voice_audio)
                bgm_duration_ms = len(bgm_audio)

                if voice_duration_ms == 0:
                    raise ValueError("合成的语音长度为0，无法制作广告。")

                self.root.after(0, lambda: progress_label.config(text="步骤3/4: 计算并混合音频..."))
                self.root.after(0, lambda: progress.config(value=60))

                def volume_to_db(vol_percent):
                    if vol_percent <= 0: return -120
                    return 20 * (vol_percent / 100.0) - 20

                adjusted_voice = voice_audio + volume_to_db(voice_volume)
                adjusted_bgm = bgm_audio + volume_to_db(bgm_volume)

                final_output = None

                if mode == 'voice':
                    if bgm_duration_ms < voice_duration_ms:
                        raise ValueError("背景音乐长度小于语音长度，无法制作。")
                    final_bgm_segment = adjusted_bgm[:voice_duration_ms]
                    final_output = final_bgm_segment.overlay(adjusted_voice)
                elif mode == 'bgm':
                    silence_5_sec = AudioSegment.silent(duration=5000)
                    unit_audio = adjusted_voice + silence_5_sec
                    if bgm_duration_ms < voice_duration_ms:
                         raise ValueError(f"背景音乐太短（{bgm_duration_ms/1000.0:.1f}秒），无法容纳一次完整的语音（需要 {voice_duration_ms/1000.0:.1f} 秒）。")

                    repeat_count = int(bgm_duration_ms // len(unit_audio))
                    if repeat_count == 0:
                        repeat_count = 1
                        unit_audio = adjusted_voice
                    
                    voice_canvas = AudioSegment.silent(duration=bgm_duration_ms)
                    current_pos_ms = 0
                    for i in range(repeat_count):
                        if current_pos_ms + len(unit_audio) <= bgm_duration_ms:
                            voice_canvas = voice_canvas.overlay(unit_audio, position=current_pos_ms)
                            current_pos_ms += len(unit_audio)
                        else:
                            if current_pos_ms + len(adjusted_voice) <= bgm_duration_ms:
                                voice_canvas = voice_canvas.overlay(adjusted_voice, position=current_pos_ms)
                            break
                    final_output = adjusted_bgm.overlay(voice_canvas)

                self.root.after(0, lambda: progress_label.config(text="步骤4/4: 导出MP3文件..."))
                self.root.after(0, lambda: progress.config(value=90))
                
                ad_folder = os.path.join(application_path, "导出的广告")
                if not os.path.exists(ad_folder):
                    os.makedirs(ad_folder)
                
                safe_filename = re.sub(r'[\\/*?:"<>|]', "", params['name_entry'].get().strip() or '未命名广告')
                output_filename = f"{safe_filename}_{int(time.time())}.mp3"
                output_path = os.path.join(ad_folder, output_filename)

                final_output.export(
                    output_path, format="mp3", bitrate="256k",
                    parameters=["-ar", "44100", "-id3v2_version", "3"], codec="libmp3lame"
                )

                self.root.after(0, lambda: progress.config(value=100))
                self.root.after(100, lambda: messagebox.showinfo("成功", f"广告制作成功！\n\n已保存至：\n{output_path}", parent=params['dialog']))

            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("制作失败", f"发生错误：\n{e}", parent=params['dialog']))
            
            finally:
                if temp_audio_path and os.path.exists(temp_audio_path):
                    try: os.remove(temp_audio_path)
                    except Exception as e_del: self.log(f"删除临时文件 {temp_audio_path} 失败: {e_del}")
                self.root.after(0, cleanup_progress)

        threading.Thread(target=worker, daemon=True).start()
        
#第7部分
    def _import_voice_script(self, text_widget, parent_dialog):
        filename = filedialog.askopenfilename(
            title="选择要导入的文稿",
            initialdir=VOICE_SCRIPT_FOLDER,
            filetypes=[("文本文档", "*.txt"), ("所有文件", "*.*")],
            parent=parent_dialog
        )
        if not filename:
            return

        try:
            with open(filename, 'r', encoding='utf-8') as f:
                content = f.read()
            text_widget.delete('1.0', END)
            text_widget.insert('1.0', content)
            self.log(f"已从 {os.path.basename(filename)} 成功导入文稿。")
        except Exception as e:
            messagebox.showerror("导入失败", f"无法读取文件：\n{e}", parent=parent_dialog)
            self.log(f"导入文稿失败: {e}")

    def _export_voice_script(self, text_widget, name_widget, parent_dialog):
        content = text_widget.get('1.0', END).strip()
        if not content:
            messagebox.showwarning("无法导出", "播音文字内容为空，无需导出。", parent=parent_dialog)
            return

        program_name = name_widget.get().strip()
        if program_name:
            invalid_chars = '\\/:*?"<>|'
            safe_name = "".join(c for c in program_name if c not in invalid_chars).strip()
            default_filename = f"{safe_name}.txt" if safe_name else "未命名文稿.txt"
        else:
            default_filename = "未命名文稿.txt"

        filename = filedialog.asksaveasfilename(
            title="导出文稿到...",
            initialdir=VOICE_SCRIPT_FOLDER,
            initialfile=default_filename,
            defaultextension=".txt",
            filetypes=[("文本文档", "*.txt")],
            parent=parent_dialog
        )
        if not filename:
            return

        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(content)
            self.log(f"文稿已成功导出到 {os.path.basename(filename)}。")
            messagebox.showinfo("导出成功", f"文稿已成功导出到：\n{filename}", parent=parent_dialog)
        except Exception as e:
            messagebox.showerror("导出失败", f"无法保存文件：\n{e}", parent=parent_dialog)
            self.log(f"导出文稿失败: {e}")

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
        if not WIN32_AVAILABLE:
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
        if not WIN32_AVAILABLE: return []
        try:
            pythoncom.CoInitialize()
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            voices = [v.GetDescription() for v in speaker.GetVoices()]
            pythoncom.CoUninitialize()
            return voices
        except Exception as e:
            self.log(f"警告: 使用 win32com 获取语音列表失败 - {e}")
            return []

    async def _edge_tts_async_task(self, text, voice_params, output_path):
        """
        执行 Edge TTS 异步任务的核心部分。
        """
        voice_id = EDGE_TTS_VOICES.get(voice_params.get('voice'))
        if not voice_id:
            raise ValueError(f"无效的在线语音名称: {voice_params.get('voice')}")

        # 将 -10~10 的范围映射到 Edge TTS 需要的格式
        rate_val = int(voice_params.get('speed', 0)) * 5
        pitch_val = int(voice_params.get('pitch', 0)) * 5
        
        rate_str = f"+{rate_val}%" if rate_val >= 0 else f"{rate_val}%"
        pitch_str = f"+{pitch_val}Hz" if pitch_val >= 0 else f"{pitch_val}Hz"

        communicate = edge_tts.Communicate(text, voice_id, rate=rate_str, pitch=pitch_str)
        with open(output_path, "wb") as file:
            async for chunk in communicate.stream():
                if chunk["type"] == "audio":
                    file.write(chunk["data"])

    def _synthesis_worker_edge(self, text, voice_params, output_path, callback):
        """
        在独立线程中运行 Edge TTS 异步任务的包装器。
        """
        try:
            # 创建并运行一个新的 asyncio 事件循环
            asyncio.run(self._edge_tts_async_task(text, voice_params, output_path))
            # 任务成功，通过 after 调用主线程的回调函数
            self.root.after(0, callback, {'success': True})
        except Exception as e:
            # 任务失败，记录日志并通过回调返回错误信息
            self.log(f"在线语音合成失败: {e}")
            self.root.after(0, callback, {'success': False, 'error': str(e)})

    def select_file_for_entry(self, initial_dir, string_var, parent_dialog):
        filename = filedialog.askopenfilename(title="选择文件", initialdir=initial_dir, filetypes=[("音频文件", "*.mp3 *.wav *.ogg *.flac"), ("所有文件", "*.*")], parent=parent_dialog)
        if filename: string_var.set(filename)

    def delete_task(self):
        selections = self.task_tree.selection()
        if not selections: messagebox.showwarning("警告", "请先选择要删除的节目", parent=self.root); return
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(selections)} 个节目吗？\n(关联的语音文件也将被删除)", parent=self.root):
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
        if not selection: 
            messagebox.showwarning("警告", "请先选择要修改的节目", parent=self.root)
            return
        if len(selection) > 1: 
            messagebox.showwarning("警告", "一次只能修改一个节目", parent=self.root)
            return
        
        index = self.task_tree.index(selection[0])
        task = self.tasks[index]
        
        dummy_parent = ttk.Toplevel(self.root)
        dummy_parent.withdraw()

        task_type = task.get('type')
        if task_type == 'audio':
            self.open_audio_dialog(dummy_parent, task_to_edit=task, index=index)
        elif task_type == 'voice':
            self.open_voice_dialog(dummy_parent, task_to_edit=task, index=index)
        elif task_type == 'video':
            self.open_video_dialog(dummy_parent, task_to_edit=task, index=index)
        elif task_type == 'bell_schedule':
            self.open_bell_scheduler_dialog(dummy_parent, task_to_edit=task, index=index)
        elif task_type == 'dynamic_voice':
            self.open_dynamic_voice_dialog(dummy_parent, task_to_edit=task, index=index)
        else:
            self.log(f"警告：任务 '{task.get('name')}' 类型未知，尝试使用音频编辑器打开。")
            self.open_audio_dialog(dummy_parent, task_to_edit=task, index=index)

    def copy_task(self):
        selections = self.task_tree.selection()
        if not selections: 
            messagebox.showwarning("警告", "请先选择要复制的节目", parent=self.root)
            return

        # --- ↓↓↓ 在这里添加唯一的限制逻辑 ↓↓↓ ---
        if self.auth_info['status'] == 'Trial':
            current_count = len(self.tasks)
            copy_count = len(selections)
            if current_count + copy_count > 3:
                messagebox.showerror(
                    "试用版限制", 
                    f"试用版最多只能添加3个节目。\n\n您当前已有 {current_count} 个，无法再复制 {copy_count} 个。", 
                    parent=self.root
                )
                return # 终止复制
        # --- ↑↑↑ 限制逻辑结束 ↑↑↑ ---

        for sel in selections:
            original = self.tasks[self.task_tree.index(sel)]
            copy = json.loads(json.dumps(original))
            copy['name'] += " (副本)"
            copy['last_run'] = {}

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
        self.update_task_list()
        self.save_tasks()

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
        filename = filedialog.askopenfilename(title="选择导入文件", filetypes=[("JSON文件", "*.json")], initialdir=application_path, parent=self.root)
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f: imported = json.load(f)

                if not isinstance(imported, list) or \
                   (imported and (not isinstance(imported[0], dict) or 'time' not in imported[0] or 'type' not in imported[0])):
                    messagebox.showerror("导入失败", "文件格式不正确，看起来不是一个有效的节目单备份文件。", parent=self.root)
                    self.log(f"尝试导入格式错误的节目单文件: {os.path.basename(filename)}")
                    return

                # --- ↓↓↓ 在这里添加您的防御代码 ↓↓↓ ---
                if self.auth_info['status'] == 'Trial':
                    current_count = len(self.tasks)
                    import_count = len(imported)
                    allowed_to_add = 3 - current_count

                    if allowed_to_add <= 0:
                        messagebox.showerror("试用版限制", "试用版最多只能有3个节目，您已达到上限，无法导入。", parent=self.root)
                        return
                    
                    if import_count > allowed_to_add:
                        messagebox.showwarning(
                            "试用版限制",
                            f"试用版最多只能有3个节目。\n\n您当前已有 {current_count} 个，只能再导入 {allowed_to_add} 个。\n\n将只导入节目单中的前 {allowed_to_add} 个节目。",
                            parent=self.root
                        )
                        # 只截取允许导入的部分
                        imported = imported[:allowed_to_add]
                # --- ↑↑↑ 防御代码结束 ↑↑↑ ---

                self.tasks.extend(imported)
                self.update_task_list()
                self.save_tasks()
                self.log(f"已从 {os.path.basename(filename)} 导入 {len(imported)} 个节目")
            except Exception as e: messagebox.showerror("错误", f"导入失败: {e}", parent=self.root)

    def export_tasks(self):
        if not self.tasks: messagebox.showwarning("警告", "没有节目可以导出", parent=self.root); return
        filename = filedialog.asksaveasfilename(title="导出到...", defaultextension=".json", initialfile="broadcast_backup.json", filetypes=[("JSON文件", "*.json")], initialdir=application_path, parent=self.root)
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f: json.dump(self.tasks, f, ensure_ascii=False, indent=2)
                self.log(f"已导出 {len(self.tasks)} 个节目到 {os.path.basename(filename)}")
            except Exception as e: messagebox.showerror("错误", f"导出失败: {e}", parent=self.root)

    def enable_task(self): self._set_task_status('启用')
    def disable_task(self): self._set_task_status('禁用')

    def _set_task_status(self, status):
        selection = self.task_tree.selection()
        if not selection: messagebox.showwarning("警告", f"请先选择要{status}的节目", parent=self.root); return
        count = sum(1 for i in selection if self.tasks[self.task_tree.index(i)]['status'] != status)
        for i in selection: self.tasks[self.task_tree.index(i)]['status'] = status
        if count > 0: self.update_task_list(); self.save_tasks(); self.log(f"已{status} {count} 个节目")

#第8部分
    def _set_tasks_status_by_type(self, task_type, status):
        if not self.tasks: return

        type_name_map = {'audio': '音频', 'voice': '语音', 'video': '视频'}
        type_name = type_name_map.get(task_type, '未知')
        status_name = "启用" if status == '启用' else "禁用"

        count = 0
        for task in self.tasks:
            if task.get('type') == task_type and task.get('status') != status:
                task['status'] = status
                count += 1

        if count > 0:
            self.update_task_list()
            self.save_tasks()
            self.log(f"已将 {count} 个{type_name}节目设置为“{status_name}”状态。")
        else:
            self.log(f"没有需要状态更新的{type_name}节目。")

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
        dialog = ttk.Toplevel(self.root)
        dialog.title(title)
        dialog.resizable(False, False)
        dialog.transient(self.root)

        # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()
        # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

        result = [None]

        ttk.Label(dialog, text=prompt, font=self.font_11).pack(pady=10, padx=20)
        entry = ttk.Entry(dialog, font=self.font_11, width=15, justify='center')
        entry.pack(pady=5, padx=20)
        entry.focus_set()

        def on_confirm():
            try:
                value = int(entry.get())
                if (minvalue is not None and value < minvalue) or \
                   (maxvalue is not None and value > maxvalue):
                    messagebox.showerror("输入错误", f"请输入一个介于 {minvalue} 和 {maxvalue} 之间的整数。", parent=dialog)
                    return
                result[0] = value
                cleanup_and_destroy()
            except ValueError:
                messagebox.showerror("输入错误", "请输入一个有效的整数。", parent=dialog)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=15)

        ttk.Button(btn_frame, text="确定", command=on_confirm, bootstyle="primary", width=8).pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=cleanup_and_destroy, width=8).pack(side=LEFT, padx=10)

        dialog.bind('<Return>', lambda event: on_confirm())
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)

        self.center_window(dialog, parent=self.root)
        self.root.wait_window(dialog)
        return result[0]

    def _apply_batch_time_addition(self, start_time_entry, interval_entry, count_entry, parent_dialog):
        """处理新的批量添加时间逻辑"""
        # 1. 校验输入
        try:
            interval_min = int(interval_entry.get())
            count = int(count_entry.get())
            if interval_min < 1 or count <= 1:
                messagebox.showwarning("输入无效", "“每分钟”和“一共次数”都必须是大于1的整数。", parent=parent_dialog)
                return
        except (ValueError, TypeError):
            messagebox.showwarning("输入无效", "“每分钟”和“一共次数”必须填写大于1的整数。", parent=parent_dialog)
            return

        current_times_str = start_time_entry.get().strip()
        if not current_times_str:
            messagebox.showwarning("操作无效", "请先在“开始时间”框中至少设置一个有效的起始时间点。", parent=parent_dialog)
            return

        # 2. 获取基准时间
        first_time_str = current_times_str.split(',')[0].strip()
        base_time = self._normalize_time_string(first_time_str)
        if not base_time:
            messagebox.showerror("格式错误", f"无法识别起始时间点 '{first_time_str}'。\n请确保格式为 HH:MM:SS。", parent=parent_dialog)
            return

        # 3. 计算新的时间序列
        try:
            # 使用集合来自动处理重复的时间点
            all_times = {base_time}
            current_time_obj = datetime.strptime(base_time, "%H:%M:%S")

            for _ in range(count):
                current_time_obj += timedelta(minutes=interval_min)
                all_times.add(current_time_obj.strftime("%H:%M:%S"))

            # 4. 更新UI
            sorted_times = sorted(list(all_times))
            final_string = ", ".join(sorted_times)
            
            start_time_entry.delete(0, tk.END)
            start_time_entry.insert(0, final_string)
            self.log(f"批量生成了 {len(sorted_times)} 个时间点。")

        except Exception as e:
            messagebox.showerror("计算错误", f"生成时间序列时发生错误: {e}", parent=parent_dialog)

    def clear_all_tasks(self, delete_associated_files=True):
        if not self.tasks: return

        if delete_associated_files:
            msg = "您确定要清空所有节目吗？\n此操作将同时删除关联的语音文件，且不可恢复！"
        else:
            msg = "您确定要清空所有节目列表吗？\n（此操作不会删除音频文件）"

        if messagebox.askyesno("严重警告", msg, parent=self.root):
            files_to_delete = []
            if delete_associated_files:
                for task in self.tasks:
                    if task.get('type') == 'voice' and 'wav_filename' in task:
                        wav_filename = task.get('wav_filename')
                        if wav_filename:
                            wav_path = os.path.join(AUDIO_FOLDER, wav_filename)
                            if os.path.exists(wav_path):
                                files_to_delete.append(wav_path)

            self.tasks.clear()
            self.update_task_list()
            self.save_tasks()
            self.log("已清空所有节目列表。")

            if delete_associated_files and files_to_delete:
                for f in files_to_delete:
                    try:
                        os.remove(f)
                        self.log(f"已删除语音文件: {os.path.basename(f)}")
                    except Exception as e:
                        self.log(f"删除语音文件失败: {e}")

    def show_time_settings_dialog(self, time_entry):
        dialog = ttk.Toplevel(self.root)
        dialog.title("开始时间设置")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()
        # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=BOTH, expand=True)
        ttk.Label(main_frame, text="24小时制 HH:MM:SS", font=self.font_11_bold).pack(anchor='w', pady=5)
        list_frame = ttk.LabelFrame(main_frame, text="时间列表", padding=5)
        list_frame.pack(fill=BOTH, expand=True, pady=5)
        box_frame = ttk.Frame(list_frame); box_frame.pack(side=LEFT, fill=BOTH, expand=True)
        listbox = tk.Listbox(box_frame, font=self.font_11, height=10)
        listbox.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar = ttk.Scrollbar(box_frame, orient=VERTICAL, command=listbox.yview, bootstyle="round")
        scrollbar.pack(side=RIGHT, fill=Y); listbox.configure(yscrollcommand=scrollbar.set)

        current_times_str = ""
        if isinstance(time_entry, ttk.Entry):
            current_times_str = time_entry.get()

        for t in [t.strip() for t in current_times_str.split(',') if t.strip()]:
            listbox.insert(END, t)

        btn_frame = ttk.Frame(list_frame)
        btn_frame.pack(side=RIGHT, padx=10, fill=Y)
        new_entry = ttk.Entry(btn_frame, font=self.font_11, width=12)
        new_entry.insert(0, datetime.now().strftime("%H:%M:%S")); new_entry.pack(pady=3)
        self._bind_mousewheel_to_entry(new_entry, self._handle_time_scroll)
        def add_time():
            val = new_entry.get().strip()
            normalized_time = self._normalize_time_string(val)
            if normalized_time:
                if normalized_time not in listbox.get(0, END):
                    listbox.insert(END, normalized_time)
                    new_entry.delete(0, END)
                    new_entry.insert(0, datetime.now().strftime("%H:%M:%S"))
            else:
                messagebox.showerror("格式错误", "请输入有效的时间格式 HH:MM:SS", parent=dialog)
        def del_time():
            if listbox.curselection(): listbox.delete(listbox.curselection()[0])
        ttk.Button(btn_frame, text="添加 ↑", command=add_time).pack(pady=3, fill=X)
        ttk.Button(btn_frame, text="删除", command=del_time).pack(pady=3, fill=X)
        ttk.Button(btn_frame, text="清空", command=lambda: listbox.delete(0, END)).pack(pady=3, fill=X)
        
        bottom_frame = ttk.Frame(main_frame); bottom_frame.pack(pady=10)
        def confirm():
            result = ", ".join(list(listbox.get(0, END)))
            if isinstance(time_entry, ttk.Entry):
                time_entry.delete(0, END)
                time_entry.insert(0, result)
            cleanup_and_destroy()
        ttk.Button(bottom_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=5, ipady=5)
        ttk.Button(bottom_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=5, ipady=5)
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)
        
        self.center_window(dialog, parent=self.root)

    def show_weekday_settings_dialog(self, weekday_entry):
        dialog = ttk.Toplevel(self.root)
        dialog.title("周几或几号")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()
        # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        week_type_var = tk.StringVar(value="week")
        week_frame = ttk.LabelFrame(main_frame, text="按周", padding=10)
        week_frame.pack(fill=X, pady=5)
        ttk.Radiobutton(week_frame, text="每周", variable=week_type_var, value="week").grid(row=0, column=0, sticky='w')
        weekdays = [("周一", 1), ("周二", 2), ("周三", 3), ("周四", 4), ("周五", 5), ("周六", 6), ("周日", 7)]
        week_vars = {num: tk.IntVar(value=1) for day, num in weekdays}
        for i, (day, num) in enumerate(weekdays): ttk.Checkbutton(week_frame, text=day, variable=week_vars[num]).grid(row=(i // 4) + 1, column=i % 4, sticky='w', padx=10, pady=3)
        day_frame = ttk.LabelFrame(main_frame, text="按月", padding=10)
        day_frame.pack(fill=BOTH, expand=True, pady=5)
        ttk.Radiobutton(day_frame, text="每月", variable=week_type_var, value="day").grid(row=0, column=0, sticky='w')
        day_vars = {i: tk.IntVar(value=0) for i in range(1, 32)}
        for i in range(1, 32): ttk.Checkbutton(day_frame, text=f"{i:02d}", variable=day_vars[i]).grid(row=((i - 1) // 7) + 1, column=(i - 1) % 7, sticky='w', padx=8, pady=2)
        bottom_frame = ttk.Frame(main_frame); bottom_frame.pack(pady=10)
        current_val = weekday_entry.get()
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
            if isinstance(weekday_entry, ttk.Entry): weekday_entry.delete(0, END); weekday_entry.insert(0, result)
            cleanup_and_destroy()
        ttk.Button(bottom_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=5, ipady=5)
        ttk.Button(bottom_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=5, ipady=5)
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)

        self.center_window(dialog, parent=self.root)

#第9部分
    def show_daterange_settings_dialog(self, date_range_entry):
        dialog = ttk.Toplevel(self.root)
        dialog.title("日期范围")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()
        # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        from_frame = ttk.Frame(main_frame)
        from_frame.pack(pady=10, anchor='w')
        ttk.Label(from_frame, text="从", font=self.font_11_bold).pack(side=LEFT, padx=5)
        from_date_entry = ttk.Entry(from_frame, font=self.font_11, width=18)
        from_date_entry.pack(side=LEFT, padx=5)
        self._bind_mousewheel_to_entry(from_date_entry, self._handle_date_scroll)
        to_frame = ttk.Frame(main_frame)
        to_frame.pack(pady=10, anchor='w')
        ttk.Label(to_frame, text="到", font=self.font_11_bold).pack(side=LEFT, padx=5)
        to_date_entry = ttk.Entry(to_frame, font=self.font_11, width=18)
        to_date_entry.pack(side=LEFT, padx=5)
        self._bind_mousewheel_to_entry(to_date_entry, self._handle_date_scroll)
        try: start, end = date_range_entry.get().split('~'); from_date_entry.insert(0, start.strip()); to_date_entry.insert(0, end.strip())
        except (ValueError, IndexError): from_date_entry.insert(0, "2025-01-01"); to_date_entry.insert(0, "2099-12-31")
        ttk.Label(main_frame, text="格式: YYYY-MM-DD", font=self.font_11, bootstyle="secondary").pack(pady=10)
        bottom_frame = ttk.Frame(main_frame); bottom_frame.pack(pady=10)
        def confirm():
            start, end = from_date_entry.get().strip(), to_date_entry.get().strip()
            norm_start, norm_end = self._normalize_date_string(start), self._normalize_date_string(end)
            if norm_start and norm_end:
                date_range_entry.delete(0, END)
                date_range_entry.insert(0, f"{norm_start} ~ {norm_end}")
                cleanup_and_destroy()
            else: messagebox.showerror("格式错误", "日期格式不正确, 应为 YYYY-MM-DD", parent=dialog)
        ttk.Button(bottom_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=5, ipady=5)
        ttk.Button(bottom_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=5, ipady=5)
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)

        self.center_window(dialog, parent=self.root)

    def show_single_time_dialog(self, time_var):
        dialog = ttk.Toplevel(self.root)
        dialog.title("设置时间")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()
        # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=BOTH, expand=True)
        ttk.Label(main_frame, text="24小时制 HH:MM:SS", font=self.font_11_bold).pack(pady=5)
        time_entry = ttk.Entry(main_frame, font=self.font_12, width=15, justify='center')
        time_entry.insert(0, time_var.get()); time_entry.pack(pady=10)
        self._bind_mousewheel_to_entry(time_entry, self._handle_time_scroll)
        def confirm():
            val = time_entry.get().strip()
            normalized_time = self._normalize_time_string(val)
            if normalized_time:
                time_var.set(normalized_time)
                self.save_settings()
                cleanup_and_destroy()
            else: messagebox.showerror("格式错误", "请输入有效的时间格式 HH:MM:SS", parent=dialog)
        bottom_frame = ttk.Frame(main_frame); bottom_frame.pack(pady=10)
        ttk.Button(bottom_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=10)
        ttk.Button(bottom_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10)
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)
        
        self.center_window(dialog, parent=self.root)

    def show_power_week_time_dialog(self, title, days_var, time_var):
        dialog = ttk.Toplevel(self.root)
        dialog.title(title)
        dialog.resizable(False, False)
        dialog.transient(self.root)

        # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()
        # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

        week_frame = ttk.LabelFrame(dialog, text="选择周几", padding=10)
        week_frame.pack(fill=X, pady=10, padx=10)
        weekdays = [("周一", 1), ("周二", 2), ("周三", 3), ("周四", 4), ("周五", 5), ("周六", 6), ("周日", 7)]
        week_vars = {num: tk.IntVar() for day, num in weekdays}
        current_days = days_var.get().replace("每周:", "")
        for day_num_str in current_days: week_vars[int(day_num_str)].set(1)
        for i, (day, num) in enumerate(weekdays): ttk.Checkbutton(week_frame, text=day, variable=week_vars[num]).grid(row=0, column=i, sticky='w', padx=10, pady=3)
        
        time_frame = ttk.LabelFrame(dialog, text="设置时间", padding=10)
        time_frame.pack(fill=X, pady=10, padx=10)
        ttk.Label(time_frame, text="时间 (HH:MM:SS):").pack(side=LEFT)
        time_entry = ttk.Entry(time_frame, font=self.font_11, width=15)
        time_entry.insert(0, time_var.get()); time_entry.pack(side=LEFT, padx=10)
        self._bind_mousewheel_to_entry(time_entry, self._handle_time_scroll)
        
        def confirm():
            selected_days = sorted([str(n) for n, v in week_vars.items() if v.get()])
            if not selected_days: messagebox.showwarning("提示", "请至少选择一天", parent=dialog); return
            normalized_time = self._normalize_time_string(time_entry.get().strip())
            if not normalized_time: messagebox.showerror("格式错误", "请输入有效的时间格式 HH:MM:SS", parent=dialog); return
            days_var.set("每周:" + "".join(selected_days))
            time_var.set(normalized_time)
            self.save_settings()
            cleanup_and_destroy()

        bottom_frame = ttk.Frame(dialog); bottom_frame.pack(pady=15)
        ttk.Button(bottom_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=10)
        ttk.Button(bottom_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10)
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)

        self.center_window(dialog, parent=self.root)

    def update_task_list(self):
        if not hasattr(self, 'task_tree') or not self.task_tree.winfo_exists(): return
        selection = self.task_tree.selection()
        self.task_tree.delete(*self.task_tree.get_children())
        for task in self.tasks:
            task_type = task.get('type')

            if task_type == 'bell_schedule':
                name = "🔔 " + task.get('name', '铃声计划')
                time_count = len(task.get('generated_times', []))
                content_preview = f"包含 {time_count} 个时间点"
                self.task_tree.insert('', END, values=(
                    name,
                    task.get('status', ''),
                    "多个",
                    "准时",
                    content_preview,
                    task.get('volume', ''),
                    task.get('weekday', ''),
                    task.get('date_range', '')
                ))
            else:
                content = task.get('content', '')
                content_preview = "" 
                
                # --- ↓↓↓ 核心修改：增加对自定义列表的显示支持 ↓↓↓ ---
                if task.get('audio_type') == 'playlist':
                    count = len(task.get('custom_playlist', []))
                    content_preview = f"自定义列表 (共 {count} 首)"
                # --- ↑↑↑ 修改结束 ---
                
                elif task_type == 'voice':
                    source_text = task.get('source_text', '')
                    clean_content = source_text.replace('\n', ' ').replace('\r', '')
                    content_preview = (clean_content[:30] + '...') if len(clean_content) > 30 else clean_content
                elif content: # 对 audio (single/folder) 和 video 类型生效
                    is_url = content.lower().startswith(('http://', 'https://', 'rtsp://', 'rtmp://', 'mms://'))
                    if is_url:
                        content_preview = (content[:40] + '...') if len(content) > 40 else content
                    else:
                        content_preview = os.path.basename(content)

                display_mode = "准时" if task.get('delay') == 'ontime' else "延时"
                self.task_tree.insert('', END, values=(
                    task.get('name', ''),
                    task.get('status', ''),
                    task.get('time', ''),
                    display_mode,
                    content_preview, 
                    task.get('volume', ''),
                    task.get('weekday', ''),
                    task.get('date_range', '')
                ))

        if selection:
            try:
                valid_selection = [s for s in selection if self.task_tree.exists(s)]
                if valid_selection: self.task_tree.selection_set(valid_selection)
            except tk.TclError: pass
        self.stats_label.config(text=f"节目单：{len(self.tasks)}")
        if hasattr(self, 'status_labels'): self.status_labels[3].config(text=f"任务数量: {len(self.tasks)}")

    def update_status_bar(self):
        if not self.running: return
        now = datetime.now()
        week_map = {"1": "一", "2": "二", "3": "三", "4": "四", "5": "五", "6": "六", "7": "日"}
        day_of_week = week_map.get(str(now.isoweekday()), '')
        time_str = now.strftime(f'%Y-%m-%d 星期{day_of_week} %H:%M:%S')

        self.status_labels[0].config(text=f"当前时间: {time_str}")
        self.status_labels[1].config(text="系统状态: 运行中")
        self.root.after(1000, self.update_status_bar)

    def start_background_threads(self):
        threading.Thread(target=self._scheduler_worker, daemon=True).start()
        threading.Thread(target=self._playback_worker, daemon=True).start()
        threading.Thread(target=self._weather_worker, daemon=True).start() # <--- 新增此行
        threading.Thread(target=self._intercut_worker, daemon=True).start() # <--- 新增此行，启动插播工人线程
        self.root.after(1000, self._process_reminder_queue)

    def _check_running_processes_for_termination(self, now):
        for task_id in list(self.active_processes.keys()):
            proc_info = self.active_processes.get(task_id)
            if not proc_info: continue

            task = proc_info.get('task')
            process = proc_info.get('process')
            stop_time_str = task.get('stop_time')

            if not stop_time_str: continue

            try:
                if process.poll() is not None:
                    del self.active_processes[task_id]
                    continue
            except Exception:
                del self.active_processes[task_id]
                continue

            current_time_str = now.strftime("%H:%M:%S")
            if current_time_str >= stop_time_str:
                self.log(f"到达停止时间，正在终止任务 '{task['name']}' (PID: {process.pid})...")
                try:
                    parent = psutil.Process(process.pid)
                    for child in parent.children(recursive=True):
                        child.kill()
                    parent.kill()
                    self.log(f"任务 '{task['name']}' (PID: {process.pid}) 已被强制终止。")
                except psutil.NoSuchProcess:
                    self.log(f"尝试终止任务 '{task['name']}' 时，进程 (PID: {process.pid}) 已不存在。")
                except Exception as e:
                    self.log(f"终止任务 '{task['name']}' (PID: {process.pid}) 时发生错误: {e}")
                finally:
                    if task_id in self.active_processes:
                        del self.active_processes[task_id]

    def _scheduler_worker(self):
        while self.running:
            now = datetime.now()
            # 计算出需要预生成的时间点
            pre_generation_time = now + timedelta(minutes=PRE_GENERATION_MINUTES)

            if not self.is_app_locked_down:
                # --- 新增的预生成检查 ---
                self._check_tasks_for_pre_generation(pre_generation_time)

                # --- 原有的检查逻辑保持不变 ---
                self._check_broadcast_tasks(now)
                self._check_advanced_tasks(now)
                self._check_time_chime(now)
                self._check_todo_tasks(now)
                self._check_running_processes_for_termination(now)
                self._check_wallpaper_task(now)

            self._check_power_tasks(now)
            time.sleep(1)

    def _check_tasks_for_pre_generation(self, pre_gen_time):
        """检查是否有动态语音任务需要在指定时间点（未来）被预生成。"""
        if self._is_in_holiday(pre_gen_time):
            return

        pre_gen_time_str = pre_gen_time.strftime("%H:%M:%S")

        for task in self.tasks:
            if task.get('status') != '启用' or task.get('type') != 'dynamic_voice':
                continue

            try:
                start, end = [d.strip() for d in task.get('date_range', '').split('~')]
                if not (datetime.strptime(start, "%Y-%m-%d").date() <= pre_gen_time.date() <= datetime.strptime(end, "%Y-%m-%d").date()):
                    continue
            except (ValueError, IndexError):
                pass

            schedule = task.get('weekday', '每周:1234567')
            run_on_pre_gen_day = (schedule.startswith("每周:") and str(pre_gen_time.isoweekday()) in schedule[3:]) or \
                                 (schedule.startswith("每月:") and f"{pre_gen_time.day:02d}" in schedule[3:].split(','))
            if not run_on_pre_gen_day:
                continue

            for trigger_time in [t.strip() for t in task.get('time', '').split(',')]:
                if trigger_time == pre_gen_time_str:
                    threading.Thread(
                        target=self._pre_generate_dynamic_voice,
                        args=(task, trigger_time),
                        daemon=True
                    ).start()
                    break

    def _is_task_due(self, task, now):
        current_date_str = now.strftime("%Y-%m-%d")
        current_time_str = now.strftime("%H:%M:%S")

        if task.get('status') != '启用':
            return False, None
        
        try:
            start, end = [d.strip() for d in task.get('date_range', '').split('~')]
            if not (datetime.strptime(start, "%Y-%m-%d").date() <= now.date() <= datetime.strptime(end, "%Y-%m-%d").date()):
                return False, None
        except (ValueError, IndexError):
            pass

        schedule = task.get('weekday', '每周:1234567')
        run_today = (schedule.startswith("每周:") and str(now.isoweekday()) in schedule[3:]) or \
                    (schedule.startswith("每月:") and f"{now.day:02d}" in schedule[3:].split(','))
        if not run_today:
            return False, None

        for trigger_time in [t.strip() for t in task.get('time', '').split(',')]:
            if trigger_time == current_time_str and task.get('last_run', {}).get(trigger_time) != current_date_str:
                return True, trigger_time
        
        return False, None

    def _check_advanced_tasks(self, now):
        for task in self.screenshot_tasks:
            is_due, trigger_time = self._is_task_due(task, now)
            if is_due:
                if self._is_in_holiday(now):
                    self.log(f"跳过截屏任务 '{task['name']}'，原因：当前处于节假日期间。")
                    task.setdefault('last_run', {})[trigger_time] = now.strftime("%Y-%m-%d")
                    self.save_screenshot_tasks()
                    continue
                
                self.log(f"触发截屏任务: {task['name']}")
                threading.Thread(target=self._execute_screenshot_task, args=(task, trigger_time), daemon=True).start()
        
        for task in self.execute_tasks:
            is_due, trigger_time = self._is_task_due(task, now)
            if is_due:
                if self._is_in_holiday(now):
                    self.log(f"跳过运行任务 '{task['name']}'，原因：当前处于节假日期间。")
                    task.setdefault('last_run', {})[trigger_time] = now.strftime("%Y-%m-%d")
                    self.save_execute_tasks()
                    continue

                self.log(f"触发运行任务: {task['name']}")
                threading.Thread(target=self._execute_program_task, args=(task, trigger_time), daemon=True).start()

        for task in self.print_tasks:
            is_due, trigger_time = self._is_task_due(task, now)
            if is_due:
                if self._is_in_holiday(now):
                    self.log(f"跳过打印任务 '{task['name']}'，原因：当前处于节假日期间。")
                    task.setdefault('last_run', {})[trigger_time] = now.strftime("%Y-%m-%d")
                    self.save_print_tasks()
                    continue
                
                self.log(f"触发打印任务: {task['name']}")
                threading.Thread(target=self._execute_print_task, args=(task, trigger_time), daemon=True).start()

        for task in self.backup_tasks:
            is_due, trigger_time = self._is_task_due(task, now)
            if is_due:
                if self._is_in_holiday(now):
                    self.log(f"跳过备份任务 '{task['name']}'，原因：当前处于节假日期间。")
                    task.setdefault('last_run', {})[trigger_time] = now.strftime("%Y-%m-%d")
                    self.save_backup_tasks()
                    continue
                
                self.log(f"触发备份任务: {task['name']}")
                threading.Thread(target=self._execute_backup_task, args=(task, trigger_time), daemon=True).start()
    
    def _execute_screenshot_task(self, task, trigger_time):
        if not IMAGE_AVAILABLE:
            self.log(f"错误：Pillow库未安装，无法执行截屏任务 '{task['name']}'。")
            return
        
        try:
            repeat_count = task.get('repeat_count', 1)
            interval_seconds = task.get('interval_seconds', 0)
            stop_time_str = task.get('stop_time')

            for i in range(repeat_count):
                if stop_time_str:
                    current_time_str = datetime.now().strftime('%H:%M:%S')
                    if current_time_str >= stop_time_str:
                        self.log(f"任务 '{task['name']}' 已到达停止时间 '{stop_time_str}'，提前中止截屏。")
                        break
                
                screenshot = ImageGrab.grab()
                filename = f"Screenshot_{task['name']}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')[:-3]}.png"
                save_path = os.path.join(SCREENSHOT_FOLDER, filename)
                screenshot.save(save_path)
                self.log(f"任务 '{task['name']}' 已成功截屏 ({i+1}/{repeat_count})，保存至: {filename}")

                if i < repeat_count - 1:
                    time.sleep(interval_seconds)
            
            task.setdefault('last_run', {})[trigger_time] = datetime.now().strftime("%Y-%m-%d")
            self.save_screenshot_tasks()

        except Exception as e:
            self.log(f"执行截屏任务 '{task['name']}' 失败: {e}")

    def _execute_program_task(self, task, trigger_time):
        target_path = task.get('target_path')
        if not target_path or not os.path.exists(target_path):
            self.log(f"错误：无法执行任务 '{task['name']}'，因为目标程序路径无效或文件不存在: {target_path}")
            return
            
        try:
            import shlex
            command = [target_path]
            arguments = task.get('arguments', '')
            if arguments:
                command.extend(shlex.split(arguments))

            p = subprocess.Popen(command, cwd=os.path.dirname(target_path))
            
            task_id = f"exec_{time.time()}_{random.randint(1000,9999)}"
            self.active_processes[task_id] = {'process': p, 'task': task}
            
            self.log(f"任务 '{task['name']}' 已成功触发，进程ID: {p.pid}")
            
            task.setdefault('last_run', {})[trigger_time] = datetime.now().strftime("%Y-%m-%d")
            self.save_execute_tasks()

        except Exception as e:
            self.log(f"执行程序任务 '{task['name']}' 失败: {e}")

    def _execute_print_task(self, task, trigger_time):
        file_path = task.get('file_path')
        printer_name = task.get('printer_name')
        copies = task.get('copies', 1)

        if not file_path or not os.path.exists(file_path):
            self.log(f"错误：无法执行打印任务 '{task['name']}'，因为文件不存在: {file_path}")
            return
        
        # 确保我们有 win32print 模块可用
        if not WIN32_AVAILABLE:
            self.log(f"错误：无法执行打印任务 '{task['name']}'，因为 pywin32 模块不可用。")
            return
            
        try:
            self.log(f"准备打印 '{os.path.basename(file_path)}' {copies} 份到打印机 '{printer_name}'...")
            
            for i in range(copies):
                self.log(f"正在提交第 {i+1}/{copies} 份打印作业...")
                win32api.ShellExecute(
                    0,
                    "printto",
                    file_path,
                    f'"{printer_name}"',
                    ".",
                    0
                )
                if copies > 1:
                    time.sleep(2) 
            
            self.log(f"任务 '{task['name']}' 的所有打印作业已成功提交。")
            
            task.setdefault('last_run', {})[trigger_time] = datetime.now().strftime("%Y-%m-%d")
            self.save_print_tasks()

        except Exception as e:
            self.log(f"执行打印任务 '{task['name']}' 时发生严重错误: {e}")

    def _execute_backup_task(self, task, trigger_time):
        source = task.get('source_folder')
        target = task.get('target_folder')
        mode = task.get('backup_mode', 'mirror')

        if not source or not os.path.isdir(source):
            self.log(f"错误：无法执行备份任务 '{task['name']}'，源文件夹不存在: {source}")
            return
        if not target:
            self.log(f"错误：无法执行备份任务 '{task['name']}'，目标文件夹未指定。")
            return
        
        if not os.path.exists(target):
            try:
                os.makedirs(target)
                self.log(f"目标文件夹不存在，已自动创建: {target}")
            except Exception as e:
                self.log(f"!!! 自动创建目标文件夹失败: {e}")
                return

        try:
            self.log(f"开始执行备份任务 '{task['name']}' (模式: {mode})...")
            self.log(f"源: {source}")
            self.log(f"目标: {target}")

            command = [
                "robocopy",
                source,
                target,
                "/E",
                "/R:2",
                "/W:5",
                "/NP",
                "/TEE"
            ]

            if mode == 'mirror':
                command.append("/MIR")

            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding='utf-8', errors='replace')
            stdout, stderr = process.communicate()

            if process.returncode >= 8:
                self.log(f"!!! 备份任务 '{task['name']}' 执行失败，Robocopy 返回码: {process.returncode}")
                if stdout: self.log(f"Robocopy 输出:\n{stdout}")
                if stderr: self.log(f"Robocopy 错误:\n{stderr}")
            else:
                self.log(f"备份任务 '{task['name']}' 成功完成。")
                task.setdefault('last_run', {})[trigger_time] = datetime.now().strftime("%Y-%m-%d")
                self.save_backup_tasks()

        except Exception as e:
            self.log(f"执行备份任务 '{task['name']}' 时发生严重异常: {e}")

    def _is_in_holiday(self, check_time):
        for holiday in self.holidays:
            if holiday.get('status') != '启用':
                continue
            try:
                start_dt = datetime.strptime(holiday['start_datetime'], '%Y-%m-%d %H:%M:%S')
                end_dt = datetime.strptime(holiday['end_datetime'], '%Y-%m-%d %H:%M:%S')
                if start_dt <= check_time <= end_dt:
                    return True
            except (ValueError, KeyError):
                self.log(f"错误：节假日 '{holiday.get('name')}' 日期格式无效，已跳过。")
                continue
        return False

    def _play_chime_concurrently(self, chime_path):
        """
        以“即发即忘”的方式，在独立的音效通道上并发播放报时音，不中断主音频。
        此方法被设计为在主GUI线程中通过 after() 调用。
        """
        if not AUDIO_AVAILABLE:
            self.log("警告：Pygame未初始化，无法进行整点报时。")
            return

        try:
            self.log(f"并发播放整点报时: {os.path.basename(chime_path)}")
            # 从文件加载报时音
            chime_sound = pygame.mixer.Sound(chime_path)
            
            # 找到一个当前未被使用的音效通道
            # a reliable way to get a free channel
            channel = pygame.mixer.find_channel(True) 
            
            # 为报时音设置一个固定的、较大的音量（例如100%）
            channel.set_volume(1.0)
            
            # 在这个独立的通道上播放报时音
            channel.play(chime_sound)
            
            # 方法到此结束，不等待播放完成，主节目可以继续播放
        except Exception as e:
            self.log(f"并发播放整点报时失败: {e}")

    def _check_time_chime(self, now):
        if not self.settings.get("time_chime_enabled", False):
            return

        if now.minute == 0 and now.second == 0 and now.hour != self.last_chime_hour:
            self.last_chime_hour = now.hour

            if self._is_in_holiday(now):
                self.log("当前处于节假日，跳过整点报时。")
                return  

            chime_file = os.path.join(CHIME_FOLDER, f"{now.hour:02d}.wav")
            if os.path.exists(chime_file):
                # --- 核心修改：调用新的并发播放方法 ---
                self.root.after(0, self._play_chime_concurrently, chime_file)
            else:
                self.log(f"警告：找不到整点报时文件 {chime_file}，报时失败。")

    def _check_broadcast_tasks(self, now):
        if self._is_in_holiday(now):
            return

        tasks_to_play = []
        current_date_str = now.strftime("%Y-%m-%d")
        current_time_str = now.strftime("%H:%M:%S")

        for task in self.tasks:
            task_type = task.get('type')
            
            if task_type == 'bell_schedule':
                if task.get('status') != '启用': continue
                
                try:
                    start, end = [d.strip() for d in task.get('date_range', '').split('~')]
                    if not (datetime.strptime(start, "%Y-%m-%d").date() <= now.date() <= datetime.strptime(end, "%Y-%m-%d").date()):
                        continue
                except (ValueError, IndexError): pass
                
                schedule = task.get('weekday', '每周:1234567')
                run_today = (schedule.startswith("每周:") and str(now.isoweekday()) in schedule[3:]) or \
                            (schedule.startswith("每月:") and f"{now.day:02d}" in schedule[3:].split(','))
                if not run_today: continue

                for bell_event in task.get('generated_times', []):
                    if bell_event['time'] == current_time_str and task.get('last_run', {}).get(bell_event['time']) != current_date_str:
                        playable_task = {
                            'name': bell_event['name'],
                            'type': 'audio',
                            'audio_type': 'single',
                            'content': task['up_bell_file'] if bell_event['bell_type'] == 'up' else task['down_bell_file'],
                            'volume': task['volume'],
                            'interval_type': 'first',
                            'interval_first': '1',
                        }
                        self.playback_command_queue.put(('PLAY_INTERRUPT', (playable_task, bell_event['time'])))
                        task.setdefault('last_run', {})[bell_event['time']] = current_date_str
                        self.save_tasks()

            else:
                is_due, trigger_time = self._is_task_due(task, now)
                if is_due:
                    tasks_to_play.append((task, trigger_time))

        if not tasks_to_play:
            return

        ontime_tasks = [t for t in tasks_to_play if t[0].get('delay') == 'ontime' or t[0].get('type') == 'dynamic_voice']
        delay_tasks = [t for t in tasks_to_play if t[0].get('delay') != 'ontime' and t[0].get('type') != 'dynamic_voice']

        if ontime_tasks:
            task, trigger_time = ontime_tasks[0]
            self.log(f"准时/高优任务 '{task['name']}' 已到时间，执行高优先级中断。")
            self.playback_command_queue.put(('PLAY_INTERRUPT', (task, trigger_time)))

        for task, trigger_time in delay_tasks:
            self.log(f"延时任务 '{task['name']}' 已到时间，加入播放队列。")
            self.playback_command_queue.put(('PLAY', (task, trigger_time)))

    def _check_power_tasks(self, now):
        current_date_str = now.strftime("%Y-%m-%d")
        current_time_str = now.strftime("%H:%M:%S")
        if self.settings.get("last_power_action_date") == current_date_str: return
        action_to_take = None
        if self.settings.get("daily_shutdown_enabled") and current_time_str == self.settings.get("daily_shutdown_time"): action_to_take = ("shutdown /s /t 60", "每日定时关机")
        if not action_to_take and self.settings.get("weekly_shutdown_enabled"):
            days = self.settings.get("weekly_shutdown_days", "").replace("每周:", "")
            if str(now.isoweekday()) in days and current_time_str == self.settings.get("weekly_shutdown_time"): action_to_take = ("shutdown /s /t 60", "每周定时关机")
        if not action_to_take and self.settings.get("weekly_reboot_enabled"):
            days = self.settings.get("weekly_reboot_days", "").replace("每周:", "")
            if str(now.isoweekday()) in days and current_time_str == self.settings.get("weekly_reboot_time"): action_to_take = ("shutdown /r /t 60", "每周定时重启")
        if action_to_take:
            command, reason = action_to_take
            self.log(f"执行系统电源任务: {reason}。系统将在60秒后操作。")
            self.settings["last_power_action_date"] = current_date_str
            self.save_settings(); os.system(command)

    def _playback_worker(self):
        is_playing = False
        while self.running:
            try:
                command, data = self.playback_command_queue.get(timeout=0.1)
            except queue.Empty:
                continue

            if command == 'PLAY_INTERRUPT':
                is_playing = True
                while not self.playback_command_queue.empty():
                    try: self.playback_command_queue.get_nowait()
                    except queue.Empty: break
                self._execute_broadcast(data[0], data[1])
                is_playing = False

            elif command == 'PLAY':
                if not is_playing:
                    is_playing = True
                    self._execute_broadcast(data[0], data[1])
                    is_playing = False

            elif command == 'STOP':
                is_playing = False
                if AUDIO_AVAILABLE:
                    pygame.mixer.music.stop()
                    pygame.mixer.stop()

                if VLC_AVAILABLE and self.vlc_player:
                    self.vlc_player.stop()
                if self.video_stop_event:
                    self.video_stop_event.set()

                self.log("STOP 命令已处理，所有播放已停止。")
                self.update_playing_text("等待播放...")
                self.status_labels[2].config(text="播放状态: 待机")
                while not self.playback_command_queue.empty():
                    try: self.playback_command_queue.get_nowait()
                    except queue.Empty: break

    # 将 "原始A" 代码中的整个 _intercut_worker 函数替换为下面的版本

    def _intercut_worker(self):
        """
        专用于处理插播任务的后台线程（最终版：彻底修复死锁）。
        """
        pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
        speaker = None
        try:
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            
            while self.running:
                task_data = self.intercut_queue.get()
                
                try:
                    self.log("接收到插播任务，开始执行...")
                    was_muted = self.is_muted
                    
                    ui_elements = queue.Queue()
                    def setup_ui():
                        if not was_muted:
                            self.toggle_mute_all()
                        
                        dialog = ttk.Toplevel(self.root)
                        dialog.title("插播进行中")
                        dialog.resizable(False, False)
                        dialog.transient(self.root)
                        dialog.attributes('-topmost', True)
                        dialog.grab_set()
                        dialog.protocol("WM_DELETE_WINDOW", lambda: None)
                        
                        ttk.Label(dialog, text="正在插播中,请等待结束或紧急停止...", font=self.font_12_bold, bootstyle="info").pack(padx=40, pady=(20, 10))
                        
                        def stop_intercut_now():
                            self.log("用户请求紧急停止插播...")
                            self.intercut_stop_event.set()
                        
                        stop_btn = ttk.Button(dialog, text="紧急停止", bootstyle="danger", command=stop_intercut_now)
                        stop_btn.pack(padx=20, pady=(0, 20), fill=tk.X)
                        
                        self.center_window(dialog)
                        ui_elements.put(dialog)

                    self.root.after(0, setup_ui)
                    progress_dialog = ui_elements.get()

                    text = task_data['text']
                    params = task_data['params']
                    repeats = task_data['repeats']
                    final_text_to_speak = (text + "。 ") * repeats
                    
                    all_voices = {v.GetDescription(): v for v in speaker.GetVoices()}
                    if (voice_desc := params.get('voice')) in all_voices:
                        speaker.Voice = all_voices[voice_desc]
                    speaker.Volume = int(params.get('volume', 100))
                    escaped_text = final_text_to_speak.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                    xml_text = f"<rate absspeed='{params.get('speed', '0')}'><pitch middle='{params.get('pitch', '0')}'>{escaped_text}</pitch></rate>"
                    
                    speaker.Speak(xml_text, 1 | 2) # SVSF_ASYNC | SVSF_IS_XML

                    # --- ↓↓↓ 核心修改：用更可靠的等待机制替换旧的while循环 ↓↓↓ ---
                    
                    # 持续循环，直到语音播放完成或被手动停止
                    while True:
                        # 1. 优先检查我们的紧急停止信号
                        if self.intercut_stop_event.is_set():
                            speaker.Speak("", 3) # SVSF_PURGEBEFORESPEAK, 强制清空并停止
                            self.log("插播被用户紧急停止！")
                            break

                        # 2. 使用SAPI内置的等待方法，等待最多100毫秒
                        #    如果语音在这100毫秒内播放完了，它会返回 True
                        if speaker.WaitUntilDone(100):
                            self.log("语音引擎报告播放完成。")
                            break # 语音已正常结束，跳出循环
                        
                        # 如果100毫秒后还没结束，循环会继续，我们就可以在下一次循环开始时
                        # 再次检查紧急停止信号，这保证了高响应性。

                    # --- ↑↑↑ 核心修改结束 ↑↑↑ ---

                finally:
                    def cleanup_ui():
                        if progress_dialog and progress_dialog.winfo_exists():
                            progress_dialog.destroy()
                        if not was_muted and self.is_muted:
                             self.toggle_mute_all()
                        self.log("插播任务已完成或被中断。")
                        self.intercut_queue.task_done()
                    
                    self.root.after(0, cleanup_ui)
                    self.intercut_stop_event.clear()
        
        except Exception as e:
            self.log(f"插播工作线程初始化时发生严重错误: {e}")
        finally:
            if speaker:
                del speaker
            pythoncom.CoUninitialize()

    def _execute_intercut(self, text, voice, speed, pitch):
        text_content = text.strip()
        if not text_content:
            messagebox.showwarning("内容为空", "请输入要播报的文字内容。", parent=self.root)
            return
            
        # 保存当前文字内容到 settings 字典，以便下次加载
        self.settings["intercut_text"] = text_content
        self.save_settings() # 调用保存，写入文件

        # --- 使用自定义对话框获取次数，确保居中和模态 ---
        dialog = ttk.Toplevel(self.root)
        dialog.title("设置播放次数")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        # 使用一个队列在主线程间安全地传递结果
        result_queue = queue.Queue()

        def cleanup_and_destroy(result=None):
            result_queue.put(result)
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()

        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main_frame, text="请输入要循环播报的次数:").pack(pady=(0, 5))
        
        repeat_entry = ttk.Entry(main_frame, font=self.font_11, width=10)
        repeat_entry.pack(pady=5)
        repeat_entry.insert(0, "1")
        repeat_entry.focus_set()
        repeat_entry.selection_range(0, tk.END) # 默认选中全部文字，方便用户直接输入

        def on_confirm():
            try:
                val = int(repeat_entry.get())
                if not (1 <= val <= 100):
                    raise ValueError
                cleanup_and_destroy(val)
            except (ValueError, TypeError):
                messagebox.showerror("输入错误", "请输入一个 1 到 100 之间的整数。", parent=dialog)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="确定", command=on_confirm, bootstyle="primary").pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=lambda: cleanup_and_destroy(None)).pack(side=tk.LEFT, padx=10)

        dialog.protocol("WM_DELETE_WINDOW", lambda: cleanup_and_destroy(None))
        dialog.bind('<Return>', lambda event: on_confirm())
        
        self.center_window(dialog)
        self.root.wait_window(dialog) # 阻塞，直到用户关闭这个对话框
        
        # 从队列中获取结果
        try:
            repeat_count = result_queue.get_nowait()
        except queue.Empty:
            repeat_count = None

        # 如果用户点击了取消或关闭窗口，则中止后续操作
        if repeat_count is None:
            self.log("用户取消了插播操作。")
            return

        # --- 核心改变：将任务打包并放入队列 ---
        
        # 1. 清除任何可能残留的旧停止信号
        self.intercut_stop_event.clear()
        
        # 2. 打包任务信息
        task_data = {
            'text': text_content,
            'params': {'voice': voice, 'speed': speed, 'pitch': pitch, 'volume': '100'},
            'repeats': repeat_count
        }

        # 3. 将任务放入插播队列，后台的 _intercut_worker 线程会自动接收并处理
        self.intercut_queue.put(task_data)

    def on_weather_label_click(self, event=None):
        """处理天气标签点击事件，弹出城市输入框"""
        dialog = ttk.Toplevel(self.root)
        dialog.title("设置天气城市")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()

        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill=BOTH, expand=True)

        ttk.Label(main_frame, text="请输入城市名称 (例如: 北京, 深圳市):").pack(pady=(0, 5))
        
        city_entry = ttk.Entry(main_frame, font=self.font_11, width=30)
        city_entry.pack(pady=5)
        city_entry.insert(0, self.settings.get("weather_city", ""))
        city_entry.focus_set()

        ttk.Label(main_frame, text="留空并保存，可恢复IP自动定位。", font=self.font_9, bootstyle="secondary").pack(pady=(5, 10))

        def on_save():
            new_city = city_entry.get().strip()
            self.settings["weather_city"] = new_city
            self.save_settings()
            
            self.log(f"用户手动设置天气城市为: '{new_city}'" if new_city else "用户清空了城市设置，将恢复自动定位。")
            cleanup_and_destroy()
            
            self.main_weather_label.config(text="天气: 正在更新...")
            threading.Thread(target=self._fetch_weather_data, daemon=True).start()

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="保存", command=on_save, bootstyle="primary").pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10)

        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)
        dialog.bind('<Return>', lambda event: on_save())
        
        self.center_window(dialog)

    def _update_weather_display_threadsafe(self, text):
        """线程安全地更新界面上的天气标签"""
        if self.main_weather_label and self.main_weather_label.winfo_exists():
            self.main_weather_label.config(text=text)

    def _fetch_weather_data(self):
        """获取天气数据（智能选择城市：优先用户设置，其次IP定位）"""
        
        if not AMAP_API_KEY or AMAP_API_KEY == "此处替换为您的真实高德API Key":
            self.log("天气功能：未在代码中配置有效的API Key。")
            self.root.after(0, self._update_weather_display_threadsafe, "天气: 未配置Key (点击设置)")
            return

        city = self.settings.get("weather_city", "").strip()
        source = "用户设置"

        if not city:
            source = "IP自动定位"
            try:
                ip_url = "https://restapi.amap.com/v3/ip"
                ip_params = {"key": AMAP_API_KEY}
                response = requests.get(ip_url, params=ip_params, timeout=10)
                response.raise_for_status()
                data = response.json()
                
                if data.get("status") == "1" and isinstance(data.get("city"), str) and data.get("city"):
                    city = data["city"]
                    self.log(f"IP定位成功: {city}")
                else:
                    city = None
                    self.log(f"IP定位未能返回有效城市: {data.get('info', '未知错误')}")
            except requests.exceptions.RequestException as e:
                city = None
                self.log(f"IP定位网络请求错误: {e}")
        
        if not city:
            self.log("天气功能：无法确定城市位置。")
            self.root.after(0, self._update_weather_display_threadsafe, "天气: 定位失败 (点击设置)")
            return

        try:
            weather_url = "https://restapi.amap.com/v3/weather/weatherInfo"
            weather_params = {"key": AMAP_API_KEY, "city": city, "extensions": "base"}
            response = requests.get(weather_url, params=weather_params, timeout=10)
            response.raise_for_status()
            data = response.json()

            if data.get("status") == "1" and data.get("lives"):
                live = data["lives"][0]
                city_name, weather, temp = live.get('city'), live.get('weather'), live.get('temperature')
                wind_dir, wind_power, humidity = live.get('winddirection'), live.get('windpower'), live.get('humidity')

                display_text = f"天气: {city_name} {weather} {temp}°C {wind_dir}风 {wind_power}级 湿度:{humidity}%"
                
                self.root.after(0, self._update_weather_display_threadsafe, display_text)
                self.log(f"成功获取天气 ({source})：{display_text}")
            else:
                error_info = data.get("info", "未知天气查询错误")
                self.log(f"获取天气失败 ({source} - {city}): {error_info}")
                self.root.after(0, self._update_weather_display_threadsafe, f"天气: 查询失败 (点击修改)")
        except Exception as e:
            self.log(f"处理天气数据时出错: {e}")
            self.root.after(0, self._update_weather_display_threadsafe, "天气: 数据错误 (点击修改)")

    def _weather_worker(self):
        """后台天气更新的循环工作线程"""
        time.sleep(5)
        while self.running:
            self._fetch_weather_data()
            time.sleep(1800)

    # --- ↑↑↑ 粘贴到这里结束 ↑↑↑ ---

#第10部分
    def _execute_broadcast(self, task, trigger_time):
        # --- ↓↓↓ 新增代码：全屏冲突检查逻辑 ↓↓↓ ---
        # 1. 判断当前任务是否需要占用全屏
        task_requires_fullscreen = (
            task.get('type') == 'video' or 
            (task.get('bg_image_enabled') and task.get('bg_image_path') and os.path.isdir(task.get('bg_image_path')))
        )

        # 2. 如果计时器正在以独占模式运行，并且当前任务也需要全屏，则跳过任务
        if self.is_fullscreen_exclusive and task_requires_fullscreen:
            self.log(f"跳过任务 '{task['name']}'，因为全屏计时器正在运行中。")
            
            # 虽然跳过了，但仍然需要更新任务的“最后运行时间”，防止它在下一秒重复触发
            if trigger_time != "manual_play":
                task.setdefault('last_run', {})[trigger_time] = datetime.now().strftime("%Y-%m-%d")
                # 根据任务类型，调用对应的保存函数
                task_type_map = {
                    'audio': self.save_tasks,
                    'voice': self.save_tasks,
                    'video': self.save_tasks,
                    # 如果未来有其他全屏任务类型，在这里补充
                }
                save_function = task_type_map.get(task.get('type'))
                if save_function:
                    save_function()

            return # 直接返回，终止本次播放
        # --- ↑↑↑ 新增代码结束 ↑↑↑ ---
        self.update_playing_text(f"[{task['name']}] 正在准备播放...")
        self.status_labels[2].config(text="播放状态: 播放中")

        if trigger_time != "manual_play":
            task.setdefault('last_run', {})[trigger_time] = datetime.now().strftime("%Y-%m-%d")
            self.save_tasks()

        visual_thread = None
        stop_visual_event = None
        task_type = task.get('type')

        if task_type == 'video':
            self.video_stop_event = threading.Event()

        if task.get('bg_image_enabled') and task.get('bg_image_path') and os.path.isdir(task.get('bg_image_path')):
            if not IMAGE_AVAILABLE:
                self.log("警告：背景图片功能已启用，但 Pillow 库未安装，无法显示图片。")
            else:
                total_duration = self._get_task_total_duration(task)
                if total_duration < 10:
                    self.log(f"任务 '{task['name']}' 总时长 ({total_duration:.1f}s) 小于10秒，不加载背景图片。")
                else:
                    stop_visual_event = threading.Event()
                    self.current_stop_visual_event = stop_visual_event
                    visual_thread = threading.Thread(target=self._visual_worker, args=(task, stop_visual_event), daemon=True)
                    visual_thread.start()

        try:
            if task_type == 'audio':
                self.log(f"开始音频任务: {task['name']}")
                self._play_audio_task_internal(task)
            elif task_type == 'voice':
                self.log(f"开始语音任务: {task['name']} (共 {task.get('repeat', 1)} 遍)")
                self._play_voice_task_internal(task)
            elif task_type == 'dynamic_voice':
                self.log(f"开始动态语音任务: {task['name']}")
                self._execute_dynamic_voice_task(task)
            elif task_type == 'video':
                self.log(f"开始视频任务: {task['name']}")
                self._play_video_task_internal(task, self.video_stop_event)

        except Exception as e:
            self.log(f"播放任务 '{task['name']}' 时发生严重错误: {e}")
        finally:
            if stop_visual_event:
                stop_visual_event.set()
                self.current_stop_visual_event = None
            if visual_thread:
                visual_thread.join(timeout=1.5)

            if AUDIO_AVAILABLE:
                pygame.mixer.music.stop()
                pygame.mixer.stop()

            if VLC_AVAILABLE and self.vlc_player:
                self.vlc_player.stop()
                self.vlc_player = None
            
            if self.video_stop_event:
                self.video_stop_event = None

            self.update_playing_text("等待播放...")
            self.status_labels[2].config(text="播放状态: 待机")
            self.log(f"任务 '{task['name']}' 播放结束。")

    def _is_interrupted(self):
        try:
            command_tuple = self.playback_command_queue.get_nowait()
            command = command_tuple[0]
            if command in ['STOP', 'PLAY_INTERRUPT']:
                self.playback_command_queue.put(command_tuple)
                return True
            else:
                self.playback_command_queue.put(command_tuple)
        except queue.Empty:
            return False
        return False

    def _play_audio_task_internal(self, task):
        playlist = []
        
        # 获取基础参数
        audio_type = task.get('audio_type', 'single')
        interval_type = task.get('interval_type', 'first')
        repeat_count = int(task.get('interval_first', 1))
        duration_seconds = int(task.get('interval_seconds', 0))

        # --- 1. 根据类型构建播放列表 ---
        
        if audio_type == 'single':
            if os.path.exists(task['content']):
                if interval_type == 'first':
                    # 单文件模式：重复播放同一个文件 n 次
                    playlist = [task['content']] * repeat_count
                else: # 按秒播放
                    # 循环播放同一个文件，直到时间到（给一个足够大的列表）
                    playlist = [task['content']] * 1000 
        
        elif audio_type == 'folder':
            folder_path = task['content']
            if os.path.isdir(folder_path):
                # 支持的音频格式
                supported_extensions = ('.mp3', '.wav', '.ogg', '.flac', '.m4a', '.wma', '.ape')
                all_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.lower().endswith(supported_extensions)]
                
                if not all_files:
                    self.log(f"警告：文件夹为空或无支持的音频文件: {folder_path}")
                    return

                if task.get('play_order') == 'random':
                    random.shuffle(all_files)
                else:
                    all_files.sort() # 顺序播时按文件名排序
                
                if interval_type == 'first':
                    # 文件夹模式：播放前 n 个文件
                    playlist = all_files[:repeat_count]
                else: # 按秒播放
                    # 循环播放整个文件夹，直到时间到
                    playlist = all_files * 100 
                    
        elif audio_type == 'playlist':
            # --- [新增] 自定义列表模式逻辑 ---
            custom_list = task.get('custom_playlist', [])
            if custom_list:
                if interval_type == 'first':
                    # 自定义列表模式：将整个列表重复播放 n 遍
                    playlist = custom_list * repeat_count
                else: # 按秒播放
                    # 循环播放整个列表，直到时间到
                    playlist = custom_list * 1000

        if not playlist:
            self.log(f"错误: 播放列表为空或文件/文件夹不存在，任务 '{task['name']}' 无法播放。")
            return

        # --- 2. 开始播放 (包含VLC和Pygame两种引擎) ---
        
        if VLC_AVAILABLE:
            self.log(f"使用VLC引擎播放任务 '{task['name']}'")
            try:
                instance = vlc.Instance()
                self.vlc_player = instance.media_player_new()
                
                if self.is_muted:
                    self.vlc_player.audio_set_mute(True)
                else:
                    self.vlc_player.audio_set_mute(False)

                start_time = time.time()

                for i, audio_path in enumerate(playlist):
                    # --- [新增] 健壮性检查：文件不存在则跳过 ---
                    if not os.path.exists(audio_path):
                        self.log(f"警告：文件不存在，已跳过: {os.path.basename(audio_path)}")
                        continue
                    
                    if self._is_interrupted():
                        self.log(f"任务 '{task['name']}' 被新指令中断。")
                        break
                    
                    media = instance.media_new(audio_path)
                    self.vlc_player.set_media(media)
                    self.vlc_player.audio_set_volume(int(task.get('volume', 80)))
                    self.vlc_player.play()
                    time.sleep(0.2) # 等待VLC状态更新

                    last_text_update_time = 0
                    # 播放循环
                    while self.vlc_player.get_state() in {vlc.State.Opening, vlc.State.Playing, vlc.State.Paused}:
                        if self._is_interrupted():
                            self.vlc_player.stop()
                            break

                        now = time.time()
                        # 处理按秒播放的停止逻辑
                        if interval_type == 'seconds':
                            elapsed = now - start_time
                            if elapsed >= duration_seconds:
                                self.vlc_player.stop()
                                self.log(f"已达到 {duration_seconds} 秒播放时长限制。")
                                break
                            # 更新UI倒计时
                            if now - last_text_update_time >= 1.0:
                                remaining = int(duration_seconds - elapsed)
                                self.update_playing_text(f"[{task['name']}] {os.path.basename(audio_path)} (剩余 {remaining} 秒)")
                                last_text_update_time = now
                        else:
                            # 更新UI进度
                            if now - last_text_update_time >= 1.0:
                                self.update_playing_text(f"[{task['name']}] {os.path.basename(audio_path)} ({i+1}/{len(playlist)})")
                                last_text_update_time = now
                        
                        time.sleep(0.1)
                    
                    # 外层循环检查：如果总时间到了，跳出文件列表循环
                    if interval_type == 'seconds' and (time.time() - start_time) >= duration_seconds:
                        break
                
                self.vlc_player.stop()

            except Exception as e:
                self.log(f"使用VLC播放音频失败: {e}")
            finally:
                if self.vlc_player:
                    self.vlc_player.stop()
                    self.vlc_player = None

        else:
            # --- 回退到 Pygame 播放 ---
            if not AUDIO_AVAILABLE:
                self.log("错误: Pygame未初始化，无法播放音频。")
                return
            
            self.log(f"VLC不可用，回退到Pygame引擎播放任务 '{task['name']}'。")
            supported_pygame_formats = ('.wav', '.mp3', '.ogg')
            
            start_time = time.time()
            for i, audio_path in enumerate(playlist):
                # --- [新增] 健壮性检查：文件不存在则跳过 ---
                if not os.path.exists(audio_path):
                    self.log(f"警告：文件不存在，已跳过: {os.path.basename(audio_path)}")
                    continue

                if self._is_interrupted():
                    self.log(f"任务 '{task['name']}' 被新指令中断。")
                    return

                if not audio_path.lower().endswith(supported_pygame_formats):
                    self.log(f"警告: Pygame不支持播放 '{os.path.basename(audio_path)}'。请安装VLC播放器以支持更多格式。")
                    continue

                # UI 状态更新
                status_base = f"[{task['name']}] 正在播放: {os.path.basename(audio_path)}"
                if interval_type == 'first':
                    self.update_playing_text(f"{status_base} ({i+1}/{len(playlist)})")
                self.log(f"正在播放: {os.path.basename(audio_path)}")

                try:
                    pygame.mixer.music.load(audio_path)
                    
                    task_volume_float = float(task.get('volume', 80)) / 100.0
                    self.last_bgm_volume = task_volume_float
                    
                    if self.is_muted:
                        pygame.mixer.music.set_volume(0)
                    else:
                        pygame.mixer.music.set_volume(task_volume_float)
                    
                    pygame.mixer.music.play()

                    last_text_update_time = 0
                    while pygame.mixer.music.get_busy():
                        if self._is_interrupted():
                            pygame.mixer.music.stop()
                            return

                        # 处理按秒播放的停止逻辑
                        if interval_type == 'seconds':
                            now = time.time()
                            elapsed = now - start_time
                            if elapsed >= duration_seconds:
                                pygame.mixer.music.stop()
                                self.log(f"已达到 {duration_seconds} 秒播放时长限制。")
                                return
                            if now - last_text_update_time >= 1.0:
                                remaining_seconds = int(duration_seconds - elapsed)
                                self.update_playing_text(f"{status_base} (剩余 {remaining_seconds} 秒)")
                                last_text_update_time = now

                        time.sleep(0.1)

                    # 外层循环检查
                    if interval_type == 'seconds' and (time.time() - start_time) >= duration_seconds:
                        return
                except Exception as e:
                    self.log(f"播放音频文件 {os.path.basename(audio_path)} 失败: {e}")
                    continue

    def _play_voice_task_internal(self, task):
        if not AUDIO_AVAILABLE:
            self.log("错误: Pygame未初始化，无法播放语音。")
            return

        if task.get('prompt', 0):
            if self._is_interrupted(): return
            prompt_file_path = task.get('prompt_file', '')
            
            if os.path.isabs(prompt_file_path):
                prompt_path = prompt_file_path
            else:
                prompt_path = os.path.join(PROMPT_FOLDER, prompt_file_path)

            if os.path.exists(prompt_path):
                try:
                    self.log(f"播放提示音: {os.path.basename(prompt_path)}")
                    sound = pygame.mixer.Sound(prompt_path)
                    sound.set_volume(float(task.get('prompt_volume', 80)) / 100.0)
                    channel = pygame.mixer.find_channel(True)
                    channel.play(sound)
                    while channel and channel.get_busy():
                        if self._is_interrupted(): return
                        time.sleep(0.05)
                except Exception as e:
                    self.log(f"播放提示音失败: {e}")
            else:
                self.log(f"警告: 提示音文件不存在 - {prompt_path}")

        if task.get('bgm', 0):
            if self._is_interrupted(): return
            bgm_file_path = task.get('bgm_file', '')

            if os.path.isabs(bgm_file_path):
                bgm_path = bgm_file_path
            else:
                bgm_path = os.path.join(BGM_FOLDER, bgm_file_path)

            if os.path.exists(bgm_path):
                try:
                    self.log(f"播放背景音乐: {os.path.basename(bgm_path)}")
                    pygame.mixer.music.load(bgm_path)
                    
                    # <--- 核心修改：智能设置BGM音量并“记住”它 ---
                    bgm_volume_float = float(task.get('bgm_volume', 40)) / 100.0
                    self.last_bgm_volume = bgm_volume_float  # 记住这个BGM的正确音量

                    if self.is_muted:
                        pygame.mixer.music.set_volume(0)
                    else:
                        pygame.mixer.music.set_volume(bgm_volume_float)
                    # --- 修改结束 ---

                    pygame.mixer.music.play(-1)
                except Exception as e:
                    self.log(f"播放背景音乐失败: {e}")
            else:
                self.log(f"警告: 背景音乐文件不存在 - {bgm_path}")

        speech_path = task.get('content', '')
        if not os.path.exists(speech_path):
            self.log(f"错误: 语音文件不存在 - {speech_path}")
            return

        try:
            speech_sound = pygame.mixer.Sound(speech_path)
            speech_sound.set_volume(float(task.get('volume', 80)) / 100.0)
            repeat_count = int(task.get('repeat', 1))

            speech_channel = pygame.mixer.find_channel(True)

            for i in range(repeat_count):
                if self._is_interrupted(): return

                self.log(f"正在播报第 {i+1}/{repeat_count} 遍")
                self.update_playing_text(f"[{task['name']}] 正在播报第 {i+1}/{repeat_count} 遍...")

                speech_channel.play(speech_sound)
                while speech_channel and speech_channel.get_busy():
                    if self._is_interrupted():
                        speech_channel.stop()
                        return
                    time.sleep(0.1)

                if i < repeat_count - 1:
                    time.sleep(0.5)
        except Exception as e:
            self.log(f"播放语音内容失败: {e}")

    def _play_video_task_internal(self, task, stop_event):
        if not VLC_AVAILABLE:
            self.log("错误: python-vlc 库未安装或VLC播放器未找到，无法播放视频。")
            return

        import urllib.parse

        custom_ua = task.get('custom_user_agent', '').strip()
        # 只有当用户填写了UA时，才使用它；否则，让VLC自己决定。
        user_agent = custom_ua or "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        if custom_ua:
            self.log(f"检测到自定义User-Agent，将使用: {user_agent}")

        # 实例选项保持干净，不包含任何UA设置
        vlc_instance_options = [
            '--no-xlib', 
            '--network-caching=5000'
            '--live-caching=3000'
            '--avcodec-hw=auto'
            '--hls-segment-threads=2'
        ]
        
        content_path = task.get('content', '')
        final_content_path = content_path
        is_http_url = content_path.lower().startswith(('http://', 'https://'))
        
        if is_http_url:
            self.log("检测到HTTP/HTTPS链接，正在进行预处理以获取最终地址...")
            try:
                # 预处理时，如果用户定义了UA，就用用户的，否则用通用的
                headers = {'User-Agent': user_agent}
                response = requests.get(content_path, headers=headers, stream=True, timeout=10, allow_redirects=True)
                response.raise_for_status()
                final_content_path = response.url
                if final_content_path != content_path:
                    self.log(f"URL重定向成功！最终播放地址为: {final_content_path}")
                else:
                    self.log("URL无需重定向，使用原始地址。")
                response.close()
            except requests.exceptions.RequestException as e:
                self.log(f"!!! 预处理URL时发生网络错误: {e}")
                final_content_path = content_path
        
        main_url_part = final_content_path.split('?')[0]
        is_m3u8_playlist = main_url_part.lower().endswith(('.m3u', '.m3u8'))
        is_folder_mode = task.get('video_type') == 'folder' and os.path.isdir(content_path)
        is_playlist_mode = is_folder_mode or is_m3u8_playlist

        self.vlc_player = None
        self.vlc_list_player = None
        
        try:
            if AUDIO_AVAILABLE:
                pygame.mixer.music.stop(); pygame.mixer.stop()

            instance = vlc.Instance(vlc_instance_options)

            if is_folder_mode:
                self.log(f"检测到视频文件夹模式，正在扫描: {content_path}")
                self.vlc_list_player = instance.media_list_player_new()
                self.vlc_player = self.vlc_list_player.get_media_player()
                media_list = instance.media_list_new()
                VIDEO_EXTENSIONS = ('.mp4', '.mkv', '.avi', '.mov', '.wmv', '.flv', '.mpg', '.mpeg', '.rmvb', '.rm', '.webm', '.vob', '.ts', '.3gp')
                video_files = [os.path.join(content_path, f) for f in os.listdir(content_path) if f.lower().endswith(VIDEO_EXTENSIONS)]
                if task.get('play_order') == 'random': random.shuffle(video_files)
                else: video_files.sort()
                interval_type = task.get('interval_type', 'first')
                if interval_type == 'first':
                    repeat_count = int(task.get('interval_first', 1))
                    self.log(f"文件夹模式：应用“播 n 首”规则，将播放列表限制为前 {repeat_count} 个视频。")
                    video_files = video_files[:repeat_count]
                if not video_files: raise ValueError("视频文件夹为空或不包含支持的视频文件。")
                self.log(f"找到 {len(video_files)} 个视频文件，正在添加到播放列表...")
                for video_file in video_files:
                    media_list.add_media(instance.media_new(video_file))
                self.vlc_list_player.set_media_list(media_list)

            else: # 单个文件或网络流 (包括M3U8)
                # ---
                # --- ▼▼▼ 最终的、最精准的修正：只在需要时才添加UA ▼▼▼
                # ---
                # 1. 先创建一个不带任何选项的媒体对象
                media = instance.media_new(final_content_path)
                
                # 2. 检查用户是否在UI中填写了自定义UA
                if custom_ua:
                    # 只有当用户填写了UA时，我们才为这个媒体对象添加选项
                    self.log("检测到自定义UA，正在为媒体对象添加选项...")
                    media.add_option(f':http-user-agent={user_agent}')
                # --- ▲▲▲ 修正结束 ▲▲▲ ---
                
                if is_playlist_mode: # M3U8
                    self.log(f"检测到播放列表，启用MediaListPlayer模式。")
                    media_list = instance.media_list_new([media])
                    self.vlc_list_player = instance.media_list_player_new()
                    self.vlc_player = self.vlc_list_player.get_media_player()
                    self.vlc_list_player.set_media_list(media_list)
                else: # 普通单个文件/流
                    self.log(f"播放单个媒体文件/流: {final_content_path}")
                    self.vlc_player = instance.media_player_new()
                    self.vlc_player.set_media(media)
            
            event_manager = self.vlc_player.event_manager()
            event_manager.event_attach(vlc.EventType.MediaPlayerEncounteredError, lambda event: self.log("!!! VLC事件: 播放器遇到错误 !!!"))
            event_manager.event_attach(vlc.EventType.MediaPlayerBuffering, lambda event, new_cache: self.log(f"--- VLC事件: 正在缓冲 {new_cache:.1f}% ---"))
            event_manager.event_attach(vlc.EventType.MediaPlayerPlaying, lambda event: self.log("--- VLC事件: 状态变更为 [播放中] ---"))
            event_manager.event_attach(vlc.EventType.MediaPlayerEndReached, lambda event: self.log("--- VLC事件: 媒体播放结束 ---"))
            self.root.after(0, self._create_video_window, task, is_playlist_mode)
            time.sleep(1.0)
            if not (self.video_window and self.video_window.winfo_exists()): raise Exception("视频窗口创建失败")
            self.vlc_player.set_hwnd(self.video_window.winfo_id())
            if self.is_muted: self.vlc_player.audio_set_mute(True)
            else: self.vlc_player.audio_set_mute(False)
            self.vlc_player.audio_set_volume(int(task.get('volume', 80)))
            player_to_start = self.vlc_list_player if is_playlist_mode else self.vlc_player
            player_to_start.play()
            self.log("已发送播放指令，等待VLC引擎响应...")
            player_to_check = self.vlc_player
            start_time = time.time()
            last_text_update_time = 0
            interval_type = task.get('interval_type', 'first')
            duration_seconds = int(task.get('interval_seconds', 0))
            while player_to_check.get_state() not in {vlc.State.Ended, vlc.State.Stopped, vlc.State.Error}:
                if self._is_interrupted() or stop_event.is_set():
                    self.log("播放被手动中断。")
                    player_to_start.stop()
                    break
                now = time.time()
                if now - last_text_update_time >= 1.0:
                    current_media = player_to_check.get_media()
                    display_name = "加载中..."
                    if current_media:
                        mrl = current_media.get_mrl()
                        if mrl:
                            try:
                                decoded_mrl = urllib.parse.unquote(mrl)
                                display_name = os.path.basename(decoded_mrl)
                            except Exception: display_name = mrl
                    state = player_to_check.get_state()
                    status_text = "播放中"
                    if state == vlc.State.Buffering: status_text = "缓冲中..."
                    elif state == vlc.State.Paused: status_text = "已暂停"
                    if interval_type == 'seconds' and duration_seconds > 0:
                        elapsed = now - start_time
                        if elapsed >= duration_seconds:
                            self.log(f"已达到 {duration_seconds} 秒播放时长限制。")
                            player_to_start.stop()
                            break
                        remaining_seconds = int(duration_seconds - elapsed)
                        self.update_playing_text(f"[{task['name']}] {display_name} ({status_text} - 剩余 {remaining_seconds} 秒)")
                    else:
                        self.update_playing_text(f"[{task['name']}] {display_name} ({status_text})")
                    last_text_update_time = now
                time.sleep(0.2)
            final_state = player_to_check.get_state()
            self.log(f"播放循环结束，最终状态为: {final_state}")
        except Exception as e:
            self.log(f"播放视频任务 '{task['name']}' 时发生严重错误: {e}")
        finally:
            if self.vlc_list_player: self.vlc_list_player.stop(); self.vlc_list_player = None
            if self.vlc_player: self.vlc_player.stop(); self.vlc_player = None
            self.root.after(0, self._destroy_video_window)
            self.log(f"视频任务 '{task['name']}' 的播放逻辑清理完毕。")

    def _create_video_window(self, task, is_playlist=False):
        if self.video_window and self.video_window.winfo_exists():
            self.video_window.destroy()

        self.video_window = ttk.Toplevel(self.root)
        self.video_window.title(f"正在播放: {task['name']}")
        self.video_window.configure(bg='black')
        
        self.root.attributes('-disabled', True)
        self.video_window.attributes('-topmost', True)

        mode = task.get('playback_mode', 'fullscreen')
        if mode == 'fullscreen':
            self.video_window.attributes('-fullscreen', True)
        else:
            try:
                w, h = map(int, task.get('resolution', '1024x768').split('x'))
                x = (self.video_window.winfo_screenwidth() - w) // 2
                y = (self.video_window.winfo_screenheight() - h) // 2
                self.video_window.geometry(f'{w}x{h}+{x}+{y}')
            except Exception as e:
                self.log(f"设置视频分辨率失败: {e}, 使用默认尺寸。")
                self.video_window.geometry('1024x768')

        self.video_window.bind('<Escape>', self._handle_video_manual_stop)
        self.video_window.bind('<space>', self._handle_video_space)
        self.video_window.protocol("WM_DELETE_WINDOW", self._handle_video_manual_stop)

        # --- ↓↓↓ 核心功能：如果是播放列表，则绑定快捷键 ↓↓↓ ---
        if is_playlist:
            self.log("播放列表模式，已启用上/下一个节目快捷键 (Ctrl+Up/Down)。")
            self.video_window.bind("<Control-Up>", lambda event: self._handle_previous_track())
            self.video_window.bind("<Control-Down>", lambda event: self._handle_next_track())
        # --- ↑↑↑ 功能结束 ↑↑↑ ---

        self.video_window.focus_force()

    def _destroy_video_window(self):
        if self.video_window and self.video_window.winfo_exists():
            self.video_window.destroy()
        self.video_window = None
        # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
        self.root.attributes('-disabled', False)
        self.root.focus_force()
        # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

    def _handle_video_manual_stop(self, event=None):
        self.log("用户手动关闭视频窗口，将停止整个视频任务。")
        if self.video_stop_event:
            self.video_stop_event.set()
        if self.vlc_player:
            self.vlc_player.stop()

    def _handle_video_space(self, event=None):
        """处理空格键，切换播放/暂停"""
        if self.vlc_list_player:
            self.vlc_list_player.pause()
            self.log("快捷键触发：切换播放/暂停状态。")
        elif self.vlc_player:
            self.vlc_player.pause()
            self.log("快捷键触发：切换播放/暂停状态。")
            
    def _handle_previous_track(self, event=None):
        """处理“上一个”命令 (由 Ctrl+Up 触发)"""
        if self.vlc_list_player:
            self.vlc_list_player.previous()
            self.log("快捷键触发：切换到上一个节目。")

    def _handle_next_track(self, event=None):
        """处理“下一个”命令 (由 Ctrl+Down 触发)"""
        if self.vlc_list_player:
            self.vlc_list_player.next()
            self.log("快捷键触发：切换到下一个节目。")

    def _get_task_total_duration(self, task):
        if not AUDIO_AVAILABLE: return 0.0

        total_duration = 0.0
        try:
            if task.get('type') == 'audio':
                if task.get('interval_type') == 'seconds':
                    return float(task.get('interval_seconds', 0))

                repeat_count = int(task.get('interval_first', 1))
                if task.get('audio_type') == 'single':
                    if os.path.exists(task['content']):
                        sound = pygame.mixer.Sound(task['content'])
                        total_duration = sound.get_length() * repeat_count
                else:
                    folder_path = task['content']
                    if os.path.isdir(folder_path):
                        all_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.lower().endswith(('.mp3', '.wav', '.ogg', '.flac', '.m4a'))]
                        playlist = all_files[:repeat_count]
                        for audio_path in playlist:
                            if os.path.exists(audio_path):
                                sound = pygame.mixer.Sound(audio_path)
                                total_duration += sound.get_length()

            elif task.get('type') == 'voice':
                speech_path = task.get('content', '')
                if os.path.exists(speech_path):
                    repeat_count = int(task.get('repeat', 1))
                    sound = pygame.mixer.Sound(speech_path)
                    total_duration = sound.get_length() * repeat_count
        except Exception as e:
            self.log(f"计算任务 '{task['name']}' 时长失败: {e}")
            return 0.0

        return total_duration

    def _visual_worker(self, task, stop_event):
        try:
            if stop_event.wait(timeout=3.0): return

            image_path = task.get('bg_image_path')
            image_order = task.get('bg_image_order', 'sequential')
            interval = float(self.settings.get("bg_image_interval", 6))

            valid_extensions = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff')
            image_files = [os.path.join(image_path, f) for f in os.listdir(image_path) if f.lower().endswith(valid_extensions)]

            if not image_files:
                self.log(f"背景图片文件夹 '{os.path.basename(image_path)}' 中没有找到有效的图片。")
                return

            if image_order == 'random':
                random.shuffle(image_files)

            self.root.after(0, self._setup_fullscreen_display)
            time.sleep(0.5)

            img_index = 0
            previous_image_path = None
            while not stop_event.is_set():
                if not self.fullscreen_window:
                    break

                current_image_path = image_files[img_index]
                self.root.after(0, self._crossfade_to_image, previous_image_path, current_image_path)

                previous_image_path = current_image_path
                img_index = (img_index + 1) % len(image_files)

                if stop_event.wait(timeout=interval):
                    break

        except Exception as e:
            self.log(f"背景图片线程出错: {e}")

        finally:
            self.root.after(0, self._destroy_fullscreen_display)
            self.log("背景图片显示已结束。")

    def _setup_fullscreen_display(self):
        if self.fullscreen_window:
            self.fullscreen_window.destroy()

        self.fullscreen_window = ttk.Toplevel(self.root)
        self.fullscreen_window.attributes('-fullscreen', True)
        self.fullscreen_window.attributes('-topmost', True)
        self.fullscreen_window.configure(bg='black', cursor='none')
        self.fullscreen_window.protocol("WM_DELETE_WINDOW", lambda: None)
        self.fullscreen_window.bind("<Escape>", self._handle_esc_press)

        # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
        self.root.attributes('-disabled', True)
        # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

        self.fullscreen_label = ttk.Label(self.fullscreen_window, background='black')
        self.fullscreen_label.pack(expand=True, fill=BOTH)

    def _handle_esc_press(self, event=None):
        self.log("用户按下ESC，手动退出背景图片显示。")
        if hasattr(self, 'current_stop_visual_event') and self.current_stop_visual_event:
            self.current_stop_visual_event.set()

    def _crossfade_to_image(self, from_path, to_path):
        if not self.fullscreen_window or not self.fullscreen_label:
            return

        TRANSITION_DURATION_MS = 800
        STEPS = 20
        DELAY_PER_STEP_MS = int(TRANSITION_DURATION_MS / STEPS)

        try:
            screen_width = self.fullscreen_window.winfo_width()
            screen_height = self.fullscreen_window.winfo_height()
            
            background = Image.new('RGBA', (screen_width, screen_height), (0, 0, 0, 255))
            
            with Image.open(to_path) as img_to_pil:
                img_to_pil.thumbnail((screen_width, screen_height), Image.Resampling.LANCZOS)
                paste_x = (screen_width - img_to_pil.width) // 2
                paste_y = (screen_height - img_to_pil.height) // 2
                
                foreground_to = background.copy()
                foreground_to.paste(img_to_pil, (paste_x, paste_y))
                img_to_rgba = foreground_to

            if from_path is None:
                self.image_tk_ref = ImageTk.PhotoImage(img_to_rgba)
                self.fullscreen_label.config(image=self.image_tk_ref)
                return

            with Image.open(from_path) as img_from_pil:
                img_from_pil.thumbnail((screen_width, screen_height), Image.Resampling.LANCZOS)
                paste_x = (screen_width - img_from_pil.width) // 2
                paste_y = (screen_height - img_from_pil.height) // 2
                
                foreground_from = background.copy()
                foreground_from.paste(img_from_pil, (paste_x, paste_y))
                img_from_rgba = foreground_from

        except Exception as e:
            self.log(f"加载过渡图片失败: {e}")
            return

        def animate_step(step):
            if not self.fullscreen_window or not hasattr(self, 'fullscreen_window') or not self.fullscreen_window.winfo_exists(): return

            alpha = step / STEPS
            blended_img = Image.blend(img_from_rgba, img_to_rgba, alpha)

            self.image_tk_ref = ImageTk.PhotoImage(blended_img)
            self.fullscreen_label.config(image=self.image_tk_ref)

            if step < STEPS:
                self.root.after(DELAY_PER_STEP_MS, animate_step, step + 1)

        animate_step(0)


    def _destroy_fullscreen_display(self):
        if self.fullscreen_window:
            self.fullscreen_window.destroy()
            self.fullscreen_window = None
            self.fullscreen_label = None
            self.image_tk_ref = None
            # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
            self.root.attributes('-disabled', False)
            self.root.focus_force()
            # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

    def log(self, message): self.root.after(0, lambda: self._log_threadsafe(message))
    
    def _log_threadsafe(self, message):
        if hasattr(self, 'log_text') and self.log_text.winfo_exists():
            log_widget = self.log_text.text
            log_widget.config(state='normal')
            log_widget.insert(END, f"{datetime.now().strftime('%H:%M:%S')} -> {message}\n")
            log_widget.see(END)
            log_widget.config(state='disabled')

    def update_playing_text(self, message): self.root.after(0, lambda: self._update_playing_text_threadsafe(message))

    def _update_playing_text_threadsafe(self, message):
        if hasattr(self, 'playing_label') and self.playing_label.winfo_exists():
            self.playing_label.config(text=message)

    def save_tasks(self):
        try:
            with open(TASK_FILE, 'w', encoding='utf-8') as f: json.dump(self.tasks, f, ensure_ascii=False, indent=2)
        except Exception as e: self.log(f"保存任务失败: {e}")

    def load_tasks(self):
        if not os.path.exists(TASK_FILE): return
        try:
            with open(TASK_FILE, 'r', encoding='utf-8') as f: self.tasks = json.load(f)
            migrated = False
            for task in self.tasks:
                if 'delay' not in task: task['delay'] = 'delay' if task.get('type') == 'voice' else 'ontime'; migrated = True
                if not isinstance(task.get('last_run'), dict): task['last_run'] = {}; migrated = True
                if task.get('type') == 'voice' and 'source_text' not in task:
                    task['source_text'] = task.get('content', '')
                    task['wav_filename'] = 'needs_regeneration'
                    migrated = True

            if migrated: self.log("旧版任务数据已迁移，部分语音节目首次播放前可能需要重新编辑保存。"); self.save_tasks()
            self.update_task_list(); self.log(f"已加载 {len(self.tasks)} 个节目")

            if self.auth_info['status'] == 'Trial' and len(self.tasks) > 3:
                messagebox.showwarning("试用版限制", "检测到节目数量超过3个限制，多余的节目将自动被移除。")
                self.tasks = self.tasks[:3] # 只保留前3个
                self.update_task_list()
                self.save_tasks() # 将截断后的列表写回文件，实现“永久”移除
                self.log("试用版限制：已将超出的节目任务移除。")

        except Exception as e: self.log(f"加载任务失败: {e}")

    def load_settings(self):
        defaults = {
            "app_font": "Microsoft YaHei",
            "app_theme": "litera", # <--- 新增此行，'litera' 是默认主题
            "autostart": False, "start_minimized": False, "lock_on_start": False,
            "daily_shutdown_enabled": False, "daily_shutdown_time": "23:00:00",
            "weekly_shutdown_enabled": False, "weekly_shutdown_days": "每周:12345", "weekly_shutdown_time": "23:30:00",
            "weekly_reboot_enabled": False, "weekly_reboot_days": "每周:67", "weekly_reboot_time": "22:00:00",
            "last_power_action_date": "",
            "time_chime_speed": "0", "time_chime_pitch": "0",
            "bg_image_interval": 6,
            "weather_city": "",
            "wallpaper_enabled": False,
            "wallpaper_interval_days": "1",
            "wallpaper_change_time": "08:00:00",
            "wallpaper_cache_days": "7",
            "wallpaper_last_change_date": "",
            "timer_duration": "00:10:00",
            "timer_show_clock": True,
            "timer_play_sound": True,
            "timer_sound_file": ""
        }
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r', encoding='utf-8') as f: self.settings = json.load(f)
                for key, value in defaults.items(): self.settings.setdefault(key, value)
            except Exception as e:
                self.log(f"加载设置失败: {e}, 将使用默认设置。")
                self.settings = defaults
        else:
            self.settings = defaults
        self.log("系统设置已加载。")

    def save_settings(self):
        if hasattr(self, 'autostart_var'):
            try:
                interval = int(self.bg_image_interval_var.get())
                if not (5 <= interval <= 60):
                    interval = self.settings.get("bg_image_interval", 6)
            except:
                interval = self.settings.get("bg_image_interval", 6)

            self.settings.update({
                "app_font": self.font_var.get(),
                "app_theme": self.theme_var.get(),
                "autostart": self.autostart_var.get(),
                "start_minimized": self.start_minimized_var.get(),
                "lock_on_start": self.lock_on_start_var.get(),
                "daily_shutdown_enabled": self.daily_shutdown_enabled_var.get(),
                "daily_shutdown_time": self.daily_shutdown_time_var.get(),
                "weekly_shutdown_enabled": self.weekly_shutdown_enabled_var.get(),
                "weekly_shutdown_days": self.weekly_shutdown_days_var.get(),
                "weekly_shutdown_time": self.weekly_shutdown_time_var.get(),
                "weekly_reboot_enabled": self.weekly_reboot_enabled_var.get(),
                "weekly_reboot_days": self.weekly_reboot_days_var.get(),
                "weekly_reboot_time": self.weekly_reboot_time_var.get(),
                "time_chime_enabled": self.time_chime_enabled_var.get(),
                "time_chime_voice": self.time_chime_voice_var.get(),
                "time_chime_speed": self.time_chime_speed_var.get(),
                "time_chime_pitch": self.time_chime_pitch_var.get(),
                "bg_image_interval": interval,
                "weather_city": self.settings.get("weather_city", ""),
                "intercut_text": self.settings.get("intercut_text", ""),
                "wallpaper_enabled": self.settings.get("wallpaper_enabled", False),
                "wallpaper_interval_days": self.settings.get("wallpaper_interval_days", "1"),
                "wallpaper_change_time": self.settings.get("wallpaper_change_time", "08:00:00"),
                "wallpaper_cache_days": self.settings.get("wallpaper_cache_days", "7"),
                "wallpaper_last_change_date": self.settings.get("wallpaper_last_change_date", ""),
                "timer_duration": self.timer_duration_var.get(),
                "timer_show_clock": self.timer_show_clock_var.get(),
                "timer_play_sound": self.timer_play_sound_var.get(),
                "timer_sound_file": self.timer_sound_file_var.get()
            })
        try:
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f: json.dump(self.settings, f, ensure_ascii=False, indent=2)
        except Exception as e: self.log(f"保存设置失败: {e}")

    def _handle_autostart_setting(self):
        self.save_settings()
        enable = self.autostart_var.get()
        if not WIN32_AVAILABLE:
            self.log("错误: 自动启动功能需要 pywin32 库。")
            if enable: self.autostart_var.set(False); self.save_settings()
            messagebox.showerror("功能受限", "未安装 pywin32 库，无法设置开机启动。", parent=self.root)
            return
        shortcut_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup', " 创翔多功能定时播音旗舰版.lnk")
        target_path = sys.executable
        try:
            if enable:
                pythoncom.CoInitialize()
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.Targetpath = target_path
                shortcut.WorkingDirectory = application_path
                shortcut.IconLocation = ICON_FILE if os.path.exists(ICON_FILE) else target_path
                shortcut.save()
                pythoncom.CoUninitialize()
                self.log("已设置开机自动启动。")
            else:
                if os.path.exists(shortcut_path): os.remove(shortcut_path); self.log("已取消开机自动启动。")
        except Exception as e:
            self.log(f"错误: 操作自动启动设置失败 - {e}")
            self.autostart_var.set(not enable); self.save_settings()
            messagebox.showerror("错误", f"操作失败: {e}", parent=self.root)

    def center_window(self, win, parent=None):
        win.update_idletasks()
        width = win.winfo_width()
        height = win.winfo_height()
        
        if parent is None:
            parent = self.root
        
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()

        x = parent_x + (parent_width // 2) - (width // 2)
        y = parent_y + (parent_height // 2) - (height // 2)

        screen_width = win.winfo_screenwidth()
        screen_height = win.winfo_screenheight()
        if x < 0: x = 0
        if y < 0: y = 0
        if x + width > screen_width: x = screen_width - width
        if y + height > screen_height: y = screen_height - height
        
        win.geometry(f'+{x}+{y}')

    def _normalize_time_string(self, time_str):
        try:
            parts = str(time_str).split(':')
            if len(parts) == 2: parts.append('00')
            if len(parts) != 3: return None
            h, m, s = int(parts[0]), int(parts[1]), int(parts[2])
            if not (0 <= h <= 23 and 0 <= m <= 59 and 0 <= s <= 59): return None
            return f"{h:02d}:{m:02d}:{s:02d}"
        except (ValueError, IndexError): return None

    def _normalize_multiple_times_string(self, times_input_str):
        if not times_input_str.strip(): return True, ""
        original_times = [t.strip() for t in times_input_str.split(',') if t.strip()]
        normalized_times, invalid_times = [], []
        for t in original_times:
            normalized = self._normalize_time_string(t)
            if normalized: normalized_times.append(normalized)
            else: invalid_times.append(t)
        if invalid_times: return False, f"以下时间格式无效: {', '.join(invalid_times)}"
        return True, ", ".join(sorted(list(set(normalized_times))))

    def _normalize_date_string(self, date_str):
        try: return datetime.strptime(date_str, "%Y-%m-%d").strftime("%Y-%m-%d")
        except ValueError: return None

    def _normalize_date_range_string(self, date_range_input_str):
        if not date_range_input_str.strip(): return True, ""
        try:
            start_str, end_str = [d.strip() for d in date_range_input_str.split('~')]
            norm_start, norm_end = self._normalize_date_string(start_str), self._normalize_date_string(end_str)
            if norm_start and norm_end: return True, f"{norm_start} ~ {norm_end}"
            invalid_parts = [p for p, n in [(start_str, norm_start), (end_str, norm_end)] if not n]
            return False, f"以下日期格式无效 (应为 YYYY-MM-DD): {', '.join(invalid_parts)}"
        except (ValueError, IndexError): return False, "日期范围格式无效，应为 'YYYY-MM-DD ~ YYYY-MM-DD'"

    def show_quit_dialog(self):
        dialog = ttk.Toplevel(self.root)
        dialog.title("确认")
        dialog.resizable(False, False); dialog.transient(self.root)

        # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()
        # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

        ttk.Label(dialog, text="您想要如何操作？", font=self.font_12).pack(pady=20, padx=40)
        btn_frame = ttk.Frame(dialog); btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="退出程序", command=lambda: [cleanup_and_destroy(), self.quit_app()], bootstyle="danger").pack(side=LEFT, padx=10)
        if TRAY_AVAILABLE: ttk.Button(btn_frame, text="最小化到托盘", command=lambda: [cleanup_and_destroy(), self.hide_to_tray()], bootstyle="primary-outline").pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10)
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)
        
        self.center_window(dialog, parent=self.root)

    def hide_to_tray(self):
        if not TRAY_AVAILABLE: messagebox.showwarning("功能不可用", "pystray 或 Pillow 库未安装，无法最小化到托盘。", parent=self.root); return
        self.root.withdraw()
        self.log("程序已最小化到系统托盘。")

    def show_from_tray(self, icon, item):
        self.root.after(0, self.root.deiconify)
        self.log("程序已从托盘恢复。")

    def quit_app(self, icon=None, item=None):
        # --- ↓↓↓ 新增/修正：在退出时写入时间戳文件 ↓↓↓ ---
        try:
            with open(TIMESTAMP_FILE, "w") as f:
                # 写入当前时间戳字符串，内容本身不重要，重要的是文件的修改时间
                f.write(str(time.time()))
        except Exception:
            # 即使写入失败，也只是降低了安全性，不应阻止程序退出。
            # 使用 pass 语句来确保 except 块在语法上是有效的。
            pass
        # --- ↑↑↑ 修正结束 ↑↑↑ ---

        if self.tray_icon: self.tray_icon.stop()
        self.running = False
        self.playback_command_queue.put(('STOP', None))

        if self.root.state() == 'normal':
            self.settings["window_geometry"] = self.root.geometry()

        self.save_tasks()
        self.save_settings()
        self.save_holidays()
        self.save_todos()
        self.save_screenshot_tasks()
        self.save_execute_tasks()
        self.save_print_tasks()
        self.save_backup_tasks()
        self.save_dynamic_voice_tasks()

        if os.path.exists(DYNAMIC_VOICE_CACHE_FOLDER):
            try:
                shutil.rmtree(DYNAMIC_VOICE_CACHE_FOLDER)
                self.log("已清空动态语音缓存。")
            except Exception as e:
                self.log(f"清空动态语音缓存失败: {e}")

        if AUDIO_AVAILABLE and pygame.mixer.get_init(): pygame.mixer.quit()

        self.root.destroy()
       
        #os._exit(0)

    def toggle_mute_all(self):
        # 1. 切换静音状态
        self.is_muted = not self.is_muted

        # 2. 更新按钮的文本和样式
        if self.is_muted:
            self.mute_button.config(text="取消静音", bootstyle="warning")
            self.log("已开启全局静音。")
        else:
            self.mute_button.config(text="一键静音", bootstyle="info-outline")
            self.log("已关闭全局静音。")

        # 3. 控制当前正在播放的 VLC 播放器
        if VLC_AVAILABLE and self.vlc_player and self.vlc_player.is_playing():
            self.vlc_player.audio_toggle_mute()

        # 4. 控制当前正在播放的 Pygame 所有音频
        if AUDIO_AVAILABLE:
            # 控制所有普通音效通道 (用于语音和提示音)
            for i in range(pygame.mixer.get_num_channels()):
                channel = pygame.mixer.Channel(i)
                if self.is_muted:
                    channel.set_volume(0)
                else:
                    channel.set_volume(1.0)
            
            # 控制专用的背景音乐通道
            if pygame.mixer.music.get_busy():
                if self.is_muted:
                    pygame.mixer.music.set_volume(0)
                else:
                    # <--- 核心修改：使用“记住”的音量来恢复 ---
                    pygame.mixer.music.set_volume(self.last_bgm_volume)
                    # --- 修改结束 ---

    def setup_tray_icon(self):
        try: image = Image.open(ICON_FILE)
        except Exception as e: image = Image.new('RGB', (64, 64), 'white'); print(f"警告: 未找到或无法加载图标文件 '{ICON_FILE}': {e}")

        menu = (
            item('显示', self.show_from_tray, default=True),
            item('退出', self.quit_app)
        )

        self.tray_icon = Icon("boyin", image, " 创翔多功能定时播音旗舰版", menu)

    def start_tray_icon_thread(self):
        if TRAY_AVAILABLE and self.tray_icon is None:
            self.setup_tray_icon()
            threading.Thread(target=self.tray_icon.run, daemon=True).start()
            self.log("系统托盘图标已启动。")

    def _enable_drag_selection(self, tree):

        def on_press(event):
            self.drag_start_item = tree.identify_row(event.y)

        def on_drag(event):
            if not self.drag_start_item:
                return

            current_item = tree.identify_row(event.y)
            if not current_item:
                return

            start_index = tree.index(self.drag_start_item)
            current_index = tree.index(current_item)

            min_idx = min(start_index, current_index)
            max_idx = max(start_index, current_index)

            all_items = tree.get_children('')
            items_to_select = all_items[min_idx : max_idx + 1]

            tree.selection_set(items_to_select)

        def on_release(event):
            self.drag_start_item = None

        tree.bind("<ButtonPress-1>", on_press, True)
        tree.bind("<B1-Motion>", on_drag, True)
        tree.bind("<ButtonRelease-1>", on_release, True)

    def create_holiday_page(self):
        page_frame = ttk.Frame(self.page_container, padding=10)
        page_frame.columnconfigure(0, weight=1)

        top_frame = ttk.Frame(page_frame)
        top_frame.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 10))
        title_label = ttk.Label(top_frame, text="节假日管理", font=self.font_14_bold, bootstyle="primary")
        title_label.pack(side=LEFT)

        desc_label = ttk.Label(page_frame, text="在节假日期间，所有“定时广播”、“整点报时”和“待办事项”都将自动暂停，节假日结束后自动恢复。",
                              font=self.font_11, bootstyle="secondary", wraplength=self.root.winfo_width() - 200)
        desc_label.grid(row=1, column=0, columnspan=2, sticky='w', pady=(0, 10))

        table_frame = ttk.Frame(page_frame)
        table_frame.grid(row=2, column=0, sticky='nsew')
        page_frame.rowconfigure(2, weight=1)

        columns = ('名称', '状态', '开始时间', '结束时间')
        self.holiday_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15, selectmode='extended', bootstyle="primary")

        self.holiday_tree.heading('名称', text='节假日名称')
        self.holiday_tree.column('名称', width=250, anchor='w')
        self.holiday_tree.heading('状态', text='状态')
        self.holiday_tree.column('状态', width=100, anchor='center')
        self.holiday_tree.heading('开始时间', text='开始时间')
        self.holiday_tree.column('开始时间', width=200, anchor='center')
        self.holiday_tree.heading('结束时间', text='结束时间')
        self.holiday_tree.column('结束时间', width=200, anchor='center')

        self.holiday_tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar = ttk.Scrollbar(table_frame, orient=VERTICAL, command=self.holiday_tree.yview, bootstyle="round")
        scrollbar.pack(side=RIGHT, fill=Y)
        self.holiday_tree.configure(yscrollcommand=scrollbar.set)

        self.holiday_tree.bind("<Double-1>", lambda e: self.edit_holiday())
        self.holiday_tree.bind("<Button-3>", self.show_holiday_context_menu)
        self._enable_drag_selection(self.holiday_tree)

        action_frame = ttk.Frame(page_frame, padding=(10, 0))
        action_frame.grid(row=2, column=1, sticky='ns')

        buttons_config = [
            ("添加", self.add_holiday, "primary"), 
            ("修改", self.edit_holiday, "info"), 
            ("删除", self.delete_holiday, "danger"),
            (None, None, None),
            ("全部启用", self.enable_all_holidays, "success-outline"), 
            ("全部禁用", self.disable_all_holidays, "warning-outline"),
            (None, None, None),
            ("导入列表", self.import_holidays, "secondary-outline"), 
            ("导出列表", self.export_holidays, "secondary-outline"), 
            ("清空列表", self.clear_all_holidays, "danger-outline")
        ]

        for text, cmd, style in buttons_config:
            if text is None:
                ttk.Separator(action_frame, orient=HORIZONTAL).pack(fill=X, pady=10)
                continue
            ttk.Button(action_frame, text=text, command=cmd, bootstyle=style).pack(pady=5, fill=X)

        self.update_holiday_list()
        return page_frame

    def save_holidays(self):
        try:
            with open(HOLIDAY_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.holidays, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"保存节假日失败: {e}")

    def load_holidays(self):
        if not os.path.exists(HOLIDAY_FILE):
            return
        try:
            with open(HOLIDAY_FILE, 'r', encoding='utf-8') as f:
                self.holidays = json.load(f)
            self.log(f"已加载 {len(self.holidays)} 个节假日设置")
            if hasattr(self, 'holiday_tree'):
                self.update_holiday_list()
        except Exception as e:
            self.log(f"加载节假日失败: {e}")
            self.holidays = []

#第11部分
    def update_holiday_list(self):
        if not hasattr(self, 'holiday_tree') or not self.holiday_tree.winfo_exists(): return
        selection = self.holiday_tree.selection()
        self.holiday_tree.delete(*self.holiday_tree.get_children())
        for holiday in self.holidays:
            self.holiday_tree.insert('', END, values=(
                holiday.get('name', ''),
                holiday.get('status', '启用'),
                holiday.get('start_datetime', ''),
                holiday.get('end_datetime', '')
            ))
        if selection:
            try:
                valid_selection = [s for s in selection if self.holiday_tree.exists(s)]
                if valid_selection: self.holiday_tree.selection_set(valid_selection)
            except tk.TclError:
                pass

    def add_holiday(self):
        self.open_holiday_dialog()

    def edit_holiday(self):
        selection = self.holiday_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要修改的节假日", parent=self.root)
            return
        index = self.holiday_tree.index(selection[0])
        holiday_to_edit = self.holidays[index]
        self.open_holiday_dialog(holiday_to_edit=holiday_to_edit, index=index)

    def delete_holiday(self):
        selections = self.holiday_tree.selection()
        if not selections:
            messagebox.showwarning("警告", "请先选择要删除的节假日", parent=self.root)
            return
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(selections)} 个节假日吗？", parent=self.root):
            indices = sorted([self.holiday_tree.index(s) for s in selections], reverse=True)
            for index in indices:
                self.holidays.pop(index)
            self.update_holiday_list()
            self.save_holidays()

    def _set_holiday_status(self, status):
        selection = self.holiday_tree.selection()
        if not selection:
            messagebox.showwarning("警告", f"请先选择要{status}的节假日", parent=self.root)
            return
        for item_id in selection:
            index = self.holiday_tree.index(item_id)
            self.holidays[index]['status'] = status
        self.update_holiday_list()
        self.save_holidays()

    def open_holiday_dialog(self, holiday_to_edit=None, index=None):
        dialog = ttk.Toplevel(self.root)
        dialog.title("修改节假日" if holiday_to_edit else "添加节假日")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        # --- ↓↓↓ 【最终BUG修复 V4】核心修改 ↓↓↓ ---
        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()
        # --- ↑↑↑ 【最终BUG修复 V4】核心修改结束 ↑↑↑ ---

        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        main_frame.columnconfigure(1, weight=1)

        ttk.Label(main_frame, text="名称:").grid(row=0, column=0, sticky='w', pady=5)
        name_entry = ttk.Entry(main_frame, font=self.font_11)
        name_entry.grid(row=0, column=1, columnspan=2, sticky='ew', pady=5)

        ttk.Label(main_frame, text="开始时间:").grid(row=1, column=0, sticky='w', pady=5)
        start_date_entry = ttk.Entry(main_frame, font=self.font_11, width=15)
        start_date_entry.grid(row=1, column=1, sticky='w', pady=5)
        self._bind_mousewheel_to_entry(start_date_entry, self._handle_date_scroll)
        start_time_entry = ttk.Entry(main_frame, font=self.font_11, width=15)
        start_time_entry.grid(row=1, column=2, sticky='w', pady=5, padx=5)
        self._bind_mousewheel_to_entry(start_time_entry, self._handle_time_scroll)

        ttk.Label(main_frame, text="结束时间:").grid(row=2, column=0, sticky='w', pady=5)
        end_date_entry = ttk.Entry(main_frame, font=self.font_11, width=15)
        end_date_entry.grid(row=2, column=1, sticky='w', pady=5)
        self._bind_mousewheel_to_entry(end_date_entry, self._handle_date_scroll)
        end_time_entry = ttk.Entry(main_frame, font=self.font_11, width=15)
        end_time_entry.grid(row=2, column=2, sticky='w', pady=5, padx=5)
        self._bind_mousewheel_to_entry(end_time_entry, self._handle_time_scroll)

        ttk.Label(main_frame, text="格式: YYYY-MM-DD", font=self.font_9, bootstyle="secondary").grid(row=3, column=1, sticky='n')
        ttk.Label(main_frame, text="格式: HH:MM:SS", font=self.font_9, bootstyle="secondary").grid(row=3, column=2, sticky='n')

        if holiday_to_edit:
            name_entry.insert(0, holiday_to_edit.get('name', ''))
            start_dt_str = holiday_to_edit.get('start_datetime', ' ')
            end_dt_str = holiday_to_edit.get('end_datetime', ' ')
            start_date, start_time = start_dt_str.split(' ') if ' ' in start_dt_str else ('', '')
            end_date, end_time = end_dt_str.split(' ') if ' ' in end_dt_str else ('', '')
            start_date_entry.insert(0, start_date)
            start_time_entry.insert(0, start_time)
            end_date_entry.insert(0, end_date)
            end_time_entry.insert(0, end_time)
        else:
            now = datetime.now()
            start_date_entry.insert(0, now.strftime('%Y-%m-%d'))
            start_time_entry.insert(0, "00:00:00")
            end_date_entry.insert(0, now.strftime('%Y-%m-%d'))
            end_time_entry.insert(0, "23:59:59")

        def save():
            name = name_entry.get().strip()
            if not name:
                messagebox.showerror("错误", "节假日名称不能为空", parent=dialog)
                return

            start_date = self._normalize_date_string(start_date_entry.get().strip())
            start_time = self._normalize_time_string(start_time_entry.get().strip())
            end_date = self._normalize_date_string(end_date_entry.get().strip())
            end_time = self._normalize_time_string(end_time_entry.get().strip())

            if not all([start_date, start_time, end_date, end_time]):
                messagebox.showerror("格式错误", "日期或时间格式不正确。\n日期: YYYY-MM-DD, 时间: HH:MM:SS", parent=dialog)
                return

            try:
                start_dt = datetime.strptime(f"{start_date} {start_time}", '%Y-%m-%d %H:%M:%S')
                end_dt = datetime.strptime(f"{end_date} {end_time}", '%Y-%m-%d %H:%M:%S')
                if start_dt >= end_dt:
                    messagebox.showerror("逻辑错误", "开始时间必须早于结束时间", parent=dialog)
                    return
            except ValueError:
                messagebox.showerror("错误", "无法解析日期时间", parent=dialog)
                return

            new_holiday_data = {
                "name": name,
                "start_datetime": start_dt.strftime('%Y-%m-%d %H:%M:%S'),
                "end_datetime": end_dt.strftime('%Y-%m-%d %H:%M:%S'),
                "status": "启用" if not holiday_to_edit else holiday_to_edit.get('status', '启用')
            }

            if holiday_to_edit:
                self.holidays[index] = new_holiday_data
            else:
                self.holidays.append(new_holiday_data)

            self.update_holiday_list()
            self.save_holidays()
            cleanup_and_destroy()

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=20)
        ttk.Button(button_frame, text="保存", command=save, bootstyle="primary", width=10).pack(side=LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=cleanup_and_destroy, width=10).pack(side=LEFT, padx=10)
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)

        self.center_window(dialog, parent=self.root)

    def show_holiday_context_menu(self, event):
        if self.is_locked: return
        iid = self.holiday_tree.identify_row(event.y)
        context_menu = tk.Menu(self.root, tearoff=0, font=self.font_11)

        if iid: # 如果点击在已有项目上
            if iid not in self.holiday_tree.selection():
                self.holiday_tree.selection_set(iid)

            context_menu.add_command(label="修改", command=self.edit_holiday)
            context_menu.add_command(label="删除", command=self.delete_holiday)
            context_menu.add_separator()
            context_menu.add_command(label="置顶", command=self.move_holiday_to_top)
            context_menu.add_command(label="上移", command=lambda: self.move_holiday(-1))
            context_menu.add_command(label="下移", command=lambda: self.move_holiday(1))
            context_menu.add_command(label="置末", command=self.move_holiday_to_bottom)
            context_menu.add_separator()
            context_menu.add_command(label="启用", command=lambda: self._set_holiday_status('启用'))
            context_menu.add_command(label="禁用", command=lambda: self._set_holiday_status('禁用'))
        else: # --- ↓↓↓ 新增的逻辑：如果点击在空白处 ↓↓↓ ---
            self.holiday_tree.selection_set() # 清空所有选择
            context_menu.add_command(label="添加节假日", command=self.add_holiday)
        # --- ↑↑↑ 新增逻辑结束 ↑↑↑ ---

        context_menu.post(event.x_root, event.y_root)

    def move_holiday(self, direction):
        selection = self.holiday_tree.selection()
        if not selection or len(selection) > 1: return
        index = self.holiday_tree.index(selection[0])
        new_index = index + direction
        if 0 <= new_index < len(self.holidays):
            item = self.holidays.pop(index)
            self.holidays.insert(new_index, item)
            self.update_holiday_list(); self.save_holidays()
            new_selection_id = self.holiday_tree.get_children()[new_index]
            self.holiday_tree.selection_set(new_selection_id)
            self.holiday_tree.focus(new_selection_id)

    def move_holiday_to_top(self):
        selection = self.holiday_tree.selection()
        if not selection or len(selection) > 1: return
        index = self.holiday_tree.index(selection[0])
        if index > 0:
            item = self.holidays.pop(index)
            self.holidays.insert(0, item)
            self.update_holiday_list(); self.save_holidays()
            new_selection_id = self.holiday_tree.get_children()[0]
            self.holiday_tree.selection_set(new_selection_id)
            self.holiday_tree.focus(new_selection_id)

    def move_holiday_to_bottom(self):
        selection = self.holiday_tree.selection()
        if not selection or len(selection) > 1: return
        index = self.holiday_tree.index(selection[0])
        if index < len(self.holidays) - 1:
            item = self.holidays.pop(index)
            self.holidays.append(item)
            self.update_holiday_list(); self.save_holidays()
            new_selection_id = self.holiday_tree.get_children()[-1]
            self.holiday_tree.selection_set(new_selection_id)
            self.holiday_tree.focus(new_selection_id)

    def enable_all_holidays(self):
        if not self.holidays: return
        for holiday in self.holidays: holiday['status'] = '启用'
        self.update_holiday_list(); self.save_holidays(); self.log("已启用全部节假日。")

    def disable_all_holidays(self):
        if not self.holidays: return
        for holiday in self.holidays: holiday['status'] = '禁用'
        self.update_holiday_list(); self.save_holidays(); self.log("已禁用全部节假日。")

    def import_holidays(self):
        filename = filedialog.askopenfilename(title="选择导入节假日文件", filetypes=[("JSON文件", "*.json")], initialdir=application_path, parent=self.root)
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f: imported = json.load(f)

                if not isinstance(imported, list) or \
                   (imported and (not isinstance(imported[0], dict) or 'start_datetime' not in imported[0] or 'end_datetime' not in imported[0])):
                    messagebox.showerror("导入失败", "文件格式不正确，看起来不是一个有效的节假日备份文件。", parent=self.root)
                    self.log(f"尝试导入格式错误的节假日文件: {os.path.basename(filename)}")
                    return

                self.holidays.extend(imported)
                self.update_holiday_list(); self.save_holidays()
                self.log(f"已从 {os.path.basename(filename)} 导入 {len(imported)} 个节假日")
            except Exception as e:
                messagebox.showerror("错误", f"导入失败: {e}", parent=self.root)

    def export_holidays(self):
        if not self.holidays:
            messagebox.showwarning("警告", "没有节假日可以导出", parent=self.root)
            return
        filename = filedialog.asksaveasfilename(title="导出节假日到...", defaultextension=".json",
                                              initialfile="holidays_backup.json", filetypes=[("JSON文件", "*.json")], initialdir=application_path, parent=self.root)
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(self.holidays, f, ensure_ascii=False, indent=2)
                self.log(f"已导出 {len(self.holidays)} 个节假日到 {os.path.basename(filename)}")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {e}", parent=self.root)

    def clear_all_holidays(self):
        if not self.holidays:
            return
        if messagebox.askyesno("严重警告", "您确定要清空所有节假日吗？\n此操作不可恢复！", parent=self.root):
            self.holidays.clear()
            self.update_holiday_list()
            self.save_holidays()
            self.log("已清空所有节假日。")

    def create_todo_page(self):
        page_frame = ttk.Frame(self.page_container, padding=10)
        page_frame.columnconfigure(0, weight=1)

        top_frame = ttk.Frame(page_frame)
        top_frame.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 10))
        title_label = ttk.Label(top_frame, text="待办事项", font=self.font_14_bold, bootstyle="primary")
        title_label.pack(side=LEFT)

        desc_label = ttk.Label(page_frame, text="到达提醒时间时会弹出窗口并播放提示音。提醒功能受节假日约束。", font=self.font_11, bootstyle="secondary")
        desc_label.grid(row=1, column=0, columnspan=2, sticky='w', pady=(0, 10))

        table_frame = ttk.Frame(page_frame)
        table_frame.grid(row=2, column=0, sticky='nsew')
        page_frame.rowconfigure(2, weight=1)

        columns = ('待办事项名称', '状态', '类型', '内容', '提醒规则')
        self.todo_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15, selectmode='extended', bootstyle="primary")

        self.todo_tree.heading('待办事项名称', text='待办事项名称')
        self.todo_tree.column('待办事项名称', width=200, anchor='w')
        self.todo_tree.heading('状态', text='状态')
        self.todo_tree.column('状态', width=80, anchor='center')
        self.todo_tree.heading('类型', text='类型')
        self.todo_tree.column('类型', width=80, anchor='center')
        self.todo_tree.heading('内容', text='内容')
        self.todo_tree.column('内容', width=300, anchor='w')
        self.todo_tree.heading('提醒规则', text='提醒规则')
        self.todo_tree.column('提醒规则', width=250, anchor='center')

        self.todo_tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar = ttk.Scrollbar(table_frame, orient=VERTICAL, command=self.todo_tree.yview, bootstyle="round")
        scrollbar.pack(side=RIGHT, fill=Y)
        self.todo_tree.configure(yscrollcommand=scrollbar.set)

        self.todo_tree.bind("<Double-1>", lambda e: self.edit_todo())
        self.todo_tree.bind("<Button-3>", self.show_todo_context_menu)
        self._enable_drag_selection(self.todo_tree)

        action_frame = ttk.Frame(page_frame, padding=(10, 0))
        action_frame.grid(row=2, column=1, sticky='ns')

        buttons_config = [
            ("添加", self.add_todo, "primary"), 
            ("修改", self.edit_todo, "info"), 
            ("删除", self.delete_todo, "danger"),
            (None, None, None),
            ("全部启用", self.enable_all_todos, "success-outline"), 
            ("全部禁用", self.disable_all_todos, "warning-outline"),
            (None, None, None),
            ("导入事项", self.import_todos, "secondary-outline"), 
            ("导出事项", self.export_todos, "secondary-outline"), 
            ("清空事项", self.clear_all_todos, "danger-outline")
        ]

        for text, cmd, style in buttons_config:
            if text is None:
                ttk.Separator(action_frame, orient=HORIZONTAL).pack(fill=X, pady=10)
                continue
            ttk.Button(action_frame, text=text, command=cmd, bootstyle=style).pack(pady=5, fill=X)

        self.update_todo_list()
        return page_frame

    def save_todos(self):
        try:
            with open(TODO_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.todos, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"保存待办事项失败: {e}")

    def load_todos(self):
        if not os.path.exists(TODO_FILE):
            return
        try:
            with open(TODO_FILE, 'r', encoding='utf-8') as f:
                self.todos = json.load(f)

            migrated = False
            for todo in self.todos:
                if 'type' not in todo:
                    todo['type'] = 'onetime'
                    migrated = True
                if todo.get('status') == '待处理':
                    todo['status'] = '启用'
                    migrated = True

            if migrated:
                self.log("检测到旧版或异常状态的待办事项数据，已自动修复。")
                self.save_todos()

            self.log(f"已加载 {len(self.todos)} 个待办事项")
            if hasattr(self, 'todo_tree'):
                self.update_todo_list()
        except Exception as e:
            self.log(f"加载待办事项失败: {e}")
            self.todos = []
#增加部分
    def load_screenshot_tasks(self):
        if not os.path.exists(SCREENSHOT_TASK_FILE): return
        try:
            with open(SCREENSHOT_TASK_FILE, 'r', encoding='utf-8') as f:
                self.screenshot_tasks = json.load(f)
            self.log(f"已加载 {len(self.screenshot_tasks)} 个截屏任务")
            if hasattr(self, 'screenshot_tree'):
                self.update_screenshot_list()
        except Exception as e:
            self.log(f"加载截屏任务失败: {e}")
            self.screenshot_tasks = []

    def save_screenshot_tasks(self):
        try:
            with open(SCREENSHOT_TASK_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.screenshot_tasks, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"保存截屏任务失败: {e}")

    def load_execute_tasks(self):
        if not os.path.exists(EXECUTE_TASK_FILE): return
        try:
            with open(EXECUTE_TASK_FILE, 'r', encoding='utf-8') as f:
                self.execute_tasks = json.load(f)
            self.log(f"已加载 {len(self.execute_tasks)} 个运行任务")
            if hasattr(self, 'execute_tree'):
                self.update_execute_list()
        except Exception as e:
            self.log(f"加载运行任务失败: {e}")
            self.execute_tasks = []

    def save_execute_tasks(self):
        try:
            with open(EXECUTE_TASK_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.execute_tasks, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"保存运行任务失败: {e}")
#增加部分结束
            
#第12部分
    def update_todo_list(self):
        if not hasattr(self, 'todo_tree') or not self.todo_tree.winfo_exists(): return
        selection = self.todo_tree.selection()
        self.todo_tree.delete(*self.todo_tree.get_children())

        active_todos_count = 0
        for todo in self.todos:
            if todo.get('status') == '启用':
                active_todos_count += 1

            content = todo.get('content', '').replace('\n', ' ').replace('\r', '')
            content_preview = (content[:30] + '...') if len(content) > 30 else content

            task_type = "一次性" if todo.get('type') == 'onetime' else "循环"

            remind_info = ""
            if task_type == '一次性':
                remind_info = todo.get('remind_datetime', '')
            else:
                times = todo.get('start_times') or "无固定时间"
                interval = todo.get('interval_minutes', 0)
                if interval > 0:
                    remind_info = f"{times} (每{interval}分钟)"
                else:
                    remind_info = times

            self.todo_tree.insert('', END, values=(
                todo.get('name', ''),
                todo.get('status', '启用'),
                task_type,
                content_preview,
                remind_info
            ))
        if selection:
            try:
                valid_selection = [s for s in selection if self.todo_tree.exists(s)]
                if valid_selection: self.todo_tree.selection_set(valid_selection)
            except tk.TclError:
                pass

        if hasattr(self, 'status_labels') and len(self.status_labels) > 4:
            self.status_labels[4].config(text=f"待办事项: {active_todos_count}")

    def add_todo(self):
        self.open_todo_dialog()

    def edit_todo(self):
        selection = self.todo_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要修改的待办事项", parent=self.root)
            return
        if len(selection) > 1:
            messagebox.showwarning("警告", "一次只能修改一个待办事项", parent=self.root)
            return
        index = self.todo_tree.index(selection[0])
        todo_to_edit = self.todos[index]
        self.open_todo_dialog(todo_to_edit=todo_to_edit, index=index)

    def delete_todo(self):
        selections = self.todo_tree.selection()
        if not selections:
            messagebox.showwarning("警告", "请先选择要删除的待办事项", parent=self.root)
            return
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(selections)} 个待办事项吗？", parent=self.root):
            indices = sorted([self.todo_tree.index(s) for s in selections], reverse=True)
            for index in indices:
                self.todos.pop(index)
            self.update_todo_list()
            self.save_todos()

    def _set_todo_status(self, status):
        selection = self.todo_tree.selection()
        if not selection:
            messagebox.showwarning("警告", f"请先选择要{status}的待办事项", parent=self.root)
            return
        for item_id in selection:
            index = self.todo_tree.index(item_id)
            self.todos[index]['status'] = status
        self.update_todo_list()
        self.save_todos()

    def open_todo_dialog(self, todo_to_edit=None, index=None):
        dialog = ttk.Toplevel(self.root)
        dialog.title("修改待办事项" if todo_to_edit else "添加待办事项")
        
        # --- ↓↓↓ 【最终BUG修复 V4.3 - 您的方案】核心修改 ↓↓↓ ---
        dialog.resizable(False, False)
        dialog.transient(self.root)

        dialog.attributes('-topmost', True)
        self.root.attributes('-disabled', True)
        
        def cleanup_and_destroy():
            self.root.attributes('-disabled', False)
            dialog.destroy()
            self.root.focus_force()
        # --- ↑↑↑ 【最终BUG修复 V4.3】核心修改结束 ↑↑↑ ---

        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        main_frame.columnconfigure(1, weight=1)

        ttk.Label(main_frame, text="名称:").grid(row=0, column=0, sticky='e', pady=5, padx=5)
        name_entry = ttk.Entry(main_frame, font=self.font_11)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', pady=5)

        ttk.Label(main_frame, text="内容:").grid(row=1, column=0, sticky='ne', pady=5, padx=5)
        content_text = ScrolledText(main_frame, height=5, font=self.font_11, wrap=WORD)
        content_text.grid(row=1, column=1, columnspan=3, sticky='ew', pady=5)

        type_var = tk.StringVar(value="onetime")
        type_frame = ttk.Frame(main_frame)
        type_frame.grid(row=2, column=1, columnspan=3, sticky='w', pady=10)

        onetime_rb = ttk.Radiobutton(type_frame, text="一次性任务", variable=type_var, value="onetime")
        onetime_rb.pack(side=LEFT, padx=10)
        recurring_rb = ttk.Radiobutton(type_frame, text="循环任务", variable=type_var, value="recurring")
        recurring_rb.pack(side=LEFT, padx=10)

        onetime_lf = ttk.LabelFrame(main_frame, text="一次性任务设置", padding=10)
        recurring_lf = ttk.LabelFrame(main_frame, text="循环任务设置", padding=10)
        recurring_lf.columnconfigure(1, weight=1)

        # --- “一次性任务”界面 ---
        ttk.Label(onetime_lf, text="执行日期:").grid(row=0, column=0, sticky='e', pady=5, padx=5)
        onetime_date_entry = ttk.Entry(onetime_lf, font=self.font_11, width=20)
        onetime_date_entry.grid(row=0, column=1, sticky='w', pady=5)
        self._bind_mousewheel_to_entry(onetime_date_entry, self._handle_date_scroll)
        
        ttk.Label(onetime_lf, text="执行时间:").grid(row=1, column=0, sticky='e', pady=5, padx=5)
        onetime_time_entry = ttk.Entry(onetime_lf, font=self.font_11, width=20)
        onetime_time_entry.grid(row=1, column=1, sticky='w', pady=5)
        self._bind_mousewheel_to_entry(onetime_time_entry, self._handle_time_scroll)

        # --- ↓↓↓ 【您的建议】为“一次性任务”界面添加占位空行，使其与“循环任务”界面等高 ↓↓↓ ---
        ttk.Label(onetime_lf, text="").grid(row=2, pady=13) # 模拟 “周几/几号” 的行高
        ttk.Label(onetime_lf, text="").grid(row=3, pady=13) # 模拟 “日期范围” 的行高
        # --- ↑↑↑ 修改结束 ↑↑↑ ---

        # --- “循环任务”界面 ---
        ttk.Label(recurring_lf, text="开始时间:").grid(row=0, column=0, sticky='e', padx=5, pady=5)
        recurring_time_entry = ttk.Entry(recurring_lf, font=self.font_11)
        recurring_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        self._bind_mousewheel_to_entry(recurring_time_entry, self._handle_time_scroll)
        ttk.Button(recurring_lf, text="设置...", command=lambda: self.show_time_settings_dialog(recurring_time_entry), bootstyle="outline").grid(row=0, column=2, padx=5)

        ttk.Label(recurring_lf, text="周几/几号:").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        recurring_weekday_entry = ttk.Entry(recurring_lf, font=self.font_11)
        recurring_weekday_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=5)
        ttk.Button(recurring_lf, text="选取...", command=lambda: self.show_weekday_settings_dialog(recurring_weekday_entry), bootstyle="outline").grid(row=1, column=2, padx=5)

        ttk.Label(recurring_lf, text="日期范围:").grid(row=2, column=0, sticky='e', padx=5, pady=5)
        recurring_daterange_entry = ttk.Entry(recurring_lf, font=self.font_11)
        recurring_daterange_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=5)
        self._bind_mousewheel_to_entry(recurring_daterange_entry, self._handle_date_scroll)
        ttk.Button(recurring_lf, text="设置...", command=lambda: self.show_daterange_settings_dialog(recurring_daterange_entry), bootstyle="outline").grid(row=2, column=2, padx=5)

        ttk.Label(recurring_lf, text="循环间隔:").grid(row=3, column=0, sticky='e', padx=5, pady=5)
        interval_frame = ttk.Frame(recurring_lf)
        interval_frame.grid(row=3, column=1, sticky='w', padx=5, pady=5)
        recurring_interval_entry = ttk.Entry(interval_frame, font=self.font_11, width=8)
        recurring_interval_entry.pack(side=LEFT)
        ttk.Label(interval_frame, text="分钟 (0表示仅在'开始时间'提醒)", font=self.font_10).pack(side=LEFT, padx=5)

        def toggle_frames(*args):
            if type_var.get() == 'onetime':
                recurring_lf.grid_forget()
                onetime_lf.grid(row=3, column=0, columnspan=4, sticky='ew', padx=5, pady=5)
            else:
                onetime_lf.grid_forget()
                recurring_lf.grid(row=3, column=0, columnspan=4, sticky='ew', padx=5, pady=5)
            
            dialog.after(1, lambda: self.center_window(dialog, parent=self.root))

        type_var.trace_add("write", toggle_frames)

        now = datetime.now()
        if todo_to_edit:
            name_entry.insert(0, todo_to_edit.get('name', ''))
            content_text.insert('1.0', todo_to_edit.get('content', ''))
            type_var.set(todo_to_edit.get('type', 'onetime'))
            dt_str = todo_to_edit.get('remind_datetime', now.strftime('%Y-%m-%d %H:%M:%S'))
            d, t = dt_str.split(' ') if ' ' in dt_str else ('', '')
            onetime_date_entry.insert(0, d)
            onetime_time_entry.insert(0, t)
            recurring_time_entry.insert(0, todo_to_edit.get('start_times', ''))
            recurring_weekday_entry.insert(0, todo_to_edit.get('weekday', '每周:1234567'))
            recurring_daterange_entry.insert(0, todo_to_edit.get('date_range', '2025-01-01 ~ 2099-12-31'))
            recurring_interval_entry.insert(0, todo_to_edit.get('interval_minutes', '0'))
        else:
            onetime_date_entry.insert(0, now.strftime('%Y-%m-%d'))
            onetime_time_entry.insert(0, (now + timedelta(minutes=5)).strftime('%H:%M:%S'))
            recurring_time_entry.insert(0, now.strftime('%H:%M:%S'))
            recurring_weekday_entry.insert(0, '每周:1234567')
            recurring_daterange_entry.insert(0, '2025-01-01 ~ 2099-12-31')
            recurring_interval_entry.insert(0, '0')

        toggle_frames()

        def save():
            name = name_entry.get().strip()
            if not name:
                messagebox.showerror("错误", "待办事项名称不能为空", parent=dialog)
                return
            new_todo_data = {
                "name": name,
                "content": content_text.get('1.0', END).strip(),
                "type": type_var.get(),
                "status": "启用" if not todo_to_edit else todo_to_edit.get('status', '启用'),
                "last_run": {} if not todo_to_edit else todo_to_edit.get('last_run', {}),
            }
            if new_todo_data['type'] == 'onetime':
                date_str = self._normalize_date_string(onetime_date_entry.get().strip())
                time_str = self._normalize_time_string(onetime_time_entry.get().strip())
                if not date_str or not time_str:
                    messagebox.showerror("格式错误", "一次性任务的日期或时间格式不正确。", parent=dialog)
                    return
                new_todo_data['remind_datetime'] = f"{date_str} {time_str}"
            else: # recurring task
                try:
                    interval = int(recurring_interval_entry.get().strip() or '0')
                    if not (0 <= interval <= 1440): raise ValueError
                except ValueError:
                    messagebox.showerror("格式错误", "循环间隔必须是 0 到 1440 之间的整数。", parent=dialog)
                    return
                
                if not recurring_weekday_entry.get().strip():
                    messagebox.showerror("输入错误", "循环任务的“周几/几号”规则不能为空。", parent=dialog)
                    return
                
                if not recurring_daterange_entry.get().strip():
                    messagebox.showerror("输入错误", "循环任务的“日期范围”不能为空。", parent=dialog)
                    return
                
                is_valid_time, time_msg = self._normalize_multiple_times_string(recurring_time_entry.get().strip())
                if not is_valid_time:
                    messagebox.showerror("格式错误", time_msg, parent=dialog); return
                is_valid_date, date_msg = self._normalize_date_range_string(recurring_daterange_entry.get().strip())
                if not is_valid_date:
                    messagebox.showerror("格式错误", date_msg, parent=dialog); return
                    
                new_todo_data['start_times'] = time_msg
                new_todo_data['weekday'] = recurring_weekday_entry.get().strip()
                new_todo_data['date_range'] = date_msg
                new_todo_data['interval_minutes'] = interval
                new_todo_data['last_interval_run'] = ""
                
            if todo_to_edit:
                self.todos[index] = new_todo_data
            else:
                self.todos.append(new_todo_data)
            self.update_todo_list()
            self.save_todos()
            cleanup_and_destroy()

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=4, pady=20)
        ttk.Button(button_frame, text="保存", command=save, bootstyle="primary", width=10).pack(side=LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=cleanup_and_destroy, width=10).pack(side=LEFT, padx=10)
        
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy)
        
        dialog.after(10, lambda: self.center_window(dialog, parent=self.root))
#第13部分
    def show_todo_context_menu(self, event):
        if self.is_locked: return
        iid = self.todo_tree.identify_row(event.y)
        context_menu = tk.Menu(self.root, tearoff=0, font=self.font_11)
        
        if iid: # 如果点击在已有项目上
            if iid not in self.todo_tree.selection():
                self.todo_tree.selection_set(iid)

            context_menu.add_command(label="修改", command=self.edit_todo)
            context_menu.add_command(label="删除", command=self.delete_todo)
            context_menu.add_separator()
            context_menu.add_command(label="置顶", command=self.move_todo_to_top)
            context_menu.add_command(label="上移", command=lambda: self.move_todo(-1))
            context_menu.add_command(label="下移", command=lambda: self.move_todo(1))
            context_menu.add_command(label="置末", command=self.move_todo_to_bottom)
            context_menu.add_separator()
            context_menu.add_command(label="启用", command=lambda: self._set_todo_status('启用'))
            context_menu.add_command(label="禁用", command=lambda: self._set_todo_status('禁用'))
        else: # --- ↓↓↓ 新增的逻辑：如果点击在空白处 ↓↓↓ ---
            self.todo_tree.selection_set() # 清空所有选择
            context_menu.add_command(label="添加待办事项", command=self.add_todo)

        context_menu.post(event.x_root, event.y_root)

    def move_todo(self, direction):
        selection = self.todo_tree.selection()
        if not selection or len(selection) > 1: return
        index = self.todo_tree.index(selection[0])
        new_index = index + direction
        if 0 <= new_index < len(self.todos):
            item = self.todos.pop(index)
            self.todos.insert(new_index, item)
            self.update_todo_list(); self.save_todos()
            new_selection_id = self.todo_tree.get_children()[new_index]
            self.todo_tree.selection_set(new_selection_id)
            self.todo_tree.focus(new_selection_id)

    def move_todo_to_top(self):
        selection = self.todo_tree.selection()
        if not selection or len(selection) > 1: return
        index = self.todo_tree.index(selection[0])
        if index > 0:
            item = self.todos.pop(index)
            self.todos.insert(0, item)
            self.update_todo_list(); self.save_todos()
            new_selection_id = self.todo_tree.get_children()[0]
            self.todo_tree.selection_set(new_selection_id)
            self.todo_tree.focus(new_selection_id)

    def move_todo_to_bottom(self):
        selection = self.todo_tree.selection()
        if not selection or len(selection) > 1: return
        index = self.todo_tree.index(selection[0])
        if index < len(self.todos) - 1:
            item = self.todos.pop(index)
            self.todos.append(item)
            self.update_todo_list(); self.save_todos()
            new_selection_id = self.todo_tree.get_children()[-1]
            self.todo_tree.selection_set(new_selection_id)
            self.todo_tree.focus(new_selection_id)

    def enable_all_todos(self):
        if not self.todos: return
        for todo in self.todos: todo['status'] = '启用'
        self.update_todo_list(); self.save_todos(); self.log("已启用全部待办事项。")

    def disable_all_todos(self):
        if not self.todos: return
        for todo in self.todos: todo['status'] = '禁用'
        self.update_todo_list(); self.save_todos(); self.log("已禁用全部待办事项。")

    def import_todos(self):
        filename = filedialog.askopenfilename(title="选择导入待办事项文件", filetypes=[("JSON文件", "*.json")], initialdir=application_path, parent=self.root)
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f: imported = json.load(f)

                if not isinstance(imported, list) or \
                   (imported and (not isinstance(imported[0], dict) or 'name' not in imported[0] or 'type' not in imported[0])):
                    messagebox.showerror("导入失败", "文件格式不正确，看起来不是一个有效的待办事项备份文件。", parent=self.root)
                    return

                self.todos.extend(imported)
                self.update_todo_list(); self.save_todos()
                self.log(f"已从 {os.path.basename(filename)} 导入 {len(imported)} 个待办事项")
            except Exception as e:
                messagebox.showerror("错误", f"导入失败: {e}", parent=self.root)

    def export_todos(self):
        if not self.todos:
            messagebox.showwarning("警告", "没有待办事项可以导出", parent=self.root)
            return
        filename = filedialog.asksaveasfilename(title="导出待办事项到...", defaultextension=".json",
                                              initialfile="todos_backup.json", filetypes=[("JSON文件", "*.json")], initialdir=application_path, parent=self.root)
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(self.todos, f, ensure_ascii=False, indent=2)
                self.log(f"已导出 {len(self.todos)} 个待办事项到 {os.path.basename(filename)}")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {e}", parent=self.root)

    def clear_all_todos(self):
        if not self.todos: return
        if messagebox.askyesno("严重警告", "您确定要清空所有待办事项吗？\n此操作不可恢复！", parent=self.root):
            self.todos.clear()
            self.update_todo_list()
            self.save_todos()
            self.log("已清空所有待办事项。")

    def _check_todo_tasks(self, now):

        now_str_dt = now.strftime('%Y-%m-%d %H:%M:%S')
        now_str_date = now.strftime('%Y-%m-%d')
        now_str_time = now.strftime('%H:%M:%S')

        for index, todo in enumerate(self.todos):
            if todo.get('status') != '启用': continue

            # --- ↓↓↓ 核心修改：将节假日检查移到每个任务的触发判断逻辑中 ↓↓↓ ---
            
            should_trigger = False
            trigger_time_for_log = "" # 用于记录是哪个时间点触发的

            if todo.get('type') == 'onetime':
                if todo.get('remind_datetime') == now_str_dt:
                    should_trigger = True
                    trigger_time_for_log = todo.get('remind_datetime')

            elif todo.get('type') == 'recurring':
                try:
                    start, end = [d.strip() for d in todo.get('date_range', '').split('~')]
                    if not (datetime.strptime(start, "%Y-%m-%d").date() <= now.date() <= datetime.strptime(end, "%Y-%m-%d").date()):
                        continue
                except (ValueError, IndexError): pass

                schedule = todo.get('weekday', '每周:1234567')
                run_today = (schedule.startswith("每周:") and str(now.isoweekday()) in schedule[3:]) or \
                            (schedule.startswith("每月:") and f"{now.day:02d}" in schedule[3:].split(','))
                if not run_today: continue

                # 检查固定时间点
                for trigger_time in [t.strip() for t in todo.get('start_times', '').split(',')]:
                    if trigger_time == now_str_time and todo.get('last_run', {}).get(trigger_time) != now_str_date:
                        should_trigger = True
                        trigger_time_for_log = trigger_time
                        todo.setdefault('last_run', {})[trigger_time] = now_str_date
                        break
                
                # 检查循环间隔
                interval = todo.get('interval_minutes', 0)
                if not should_trigger and interval > 0 and todo.get('start_times'):
                    last_run_str = todo.get('last_interval_run')
                    if last_run_str:
                        try:
                            last_run_dt = datetime.strptime(last_run_str, '%Y-%m-%d %H:%M:%S')
                            if now >= last_run_dt + timedelta(minutes=interval):
                                should_trigger = True
                                trigger_time_for_log = f"间隔循环 ({now_str_time})"
                        except ValueError: pass
            
            # --- 统一的触发/跳过逻辑 ---
            if should_trigger:
                if self._is_in_holiday(now):
                    self.log(f"跳过待办事项提醒 '{todo['name']}'，原因：当前处于节假日期间。")
                    # 对于循环任务，更新间隔计时，防止节假日后立即触发
                    if todo.get('type') == 'recurring':
                        todo['last_interval_run'] = now_str_dt
                        self.save_todos()
                    continue # 跳过此任务
                
                # 如果不是节假日，则正常触发
                self.log(f"触发待办事项提醒: {todo['name']} (规则: {trigger_time_for_log})")
                todo_with_index = todo.copy()
                todo_with_index['original_index'] = index
                self.reminder_queue.put(todo_with_index)
                
                # 更新循环任务的最后运行时间
                if todo.get('type') == 'recurring':
                    todo['last_interval_run'] = now_str_dt
                    self.save_todos()

    def _process_reminder_queue(self):
        if not self.is_reminder_active and not self.reminder_queue.empty():
            try:
                todo_task = self.reminder_queue.get_nowait()
                self.is_reminder_active = True
                self.show_todo_reminder(todo_task)
            except queue.Empty:
                pass

        self.root.after(1000, self._process_reminder_queue)

    def _play_reminder_sound(self):
        if not AUDIO_AVAILABLE:
            self.log("警告：pygame未安装，无法播放提示音。")
            return

        if os.path.exists(REMINDER_SOUND_FILE):
            try:
                sound = pygame.mixer.Sound(REMINDER_SOUND_FILE)
                channel = pygame.mixer.find_channel(True)
                channel.set_volume(0.7)
                channel.play(sound)
                self.log("已播放自定义提示音。")
                return
            except Exception as e:
                self.log(f"播放自定义提示音 {REMINDER_SOUND_FILE} 失败: {e}")

        if WIN32_AVAILABLE:
            try:
                ctypes.windll.user32.MessageBeep(win32con.MB_OK)
                self.log("已播放系统默认提示音。")
            except Exception as e:
                self.log(f"播放系统默认提示音失败: {e}")

    def show_todo_reminder(self, todo):
        self._play_reminder_sound()

        reminder_win = ttk.Toplevel(self.root)
        reminder_win.title(f"待办事项提醒 - {todo.get('name')}")
        
        reminder_win.geometry("600x480")
        reminder_win.resizable(False, False)

        self.root.attributes('-disabled', True)
        reminder_win.attributes('-topmost', True)

        reminder_win.lift()
        reminder_win.focus_force()
        reminder_win.after(1000, lambda: reminder_win.attributes('-topmost', False))

        original_index = todo.get('original_index')
        task_type = todo.get('type')

        reminder_win.columnconfigure(0, weight=1)
        reminder_win.rowconfigure(1, weight=1)

        title_label = ttk.Label(reminder_win, text=todo.get('name', '无标题'), font=self.font_14_bold, wraplength=440)
        title_label.grid(row=0, column=0, pady=(15, 10), padx=20, sticky='w')

        btn_frame = ttk.Frame(reminder_win)
        btn_frame.grid(row=2, column=0, pady=(10, 15), padx=10, sticky='ew')

        content_frame = ttk.Frame(reminder_win)
        content_frame.grid(row=1, column=0, padx=20, pady=5, sticky='nsew')
        content_frame.rowconfigure(0, weight=1)
        content_frame.columnconfigure(0, weight=1)

        content_text_widget = tk.Text(content_frame, font=self.font_11, wrap=WORD, bd=0, highlightthickness=0)
        content_text_widget.grid(row=0, column=0, sticky='nsew')
        
        scrollbar = ttk.Scrollbar(content_frame, orient=VERTICAL, command=content_text_widget.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        
        content_text_widget.config(yscrollcommand=scrollbar.set)

        content_text_widget.insert('1.0', todo.get('content', ''))
        content_text_widget.config(state='disabled')

        def close_and_release():
            self.is_reminder_active = False
            self.root.attributes('-disabled', False)
            reminder_win.destroy()
            self.root.focus_force() 

        if task_type == 'onetime':
            btn_frame.columnconfigure((0, 1, 2), weight=1)
            ttk.Button(btn_frame, text="已完成", bootstyle="success", command=lambda: handle_complete()).grid(row=0, column=0, padx=5, ipady=4, sticky='ew')
            ttk.Button(btn_frame, text="稍后提醒", bootstyle="outline-secondary", command=lambda: handle_snooze()).grid(row=0, column=1, padx=5, ipady=4, sticky='ew')
            ttk.Button(btn_frame, text="删除任务", bootstyle="danger", command=lambda: handle_delete()).grid(row=0, column=2, padx=5, ipady=4, sticky='ew')
        else:
            btn_frame.columnconfigure((0, 1), weight=1)
            ttk.Button(btn_frame, text="本次完成", bootstyle="primary", command=lambda: close_and_release()).grid(row=0, column=0, padx=5, ipady=4, sticky='ew')
            ttk.Button(btn_frame, text="删除任务", bootstyle="danger", command=lambda: handle_delete()).grid(row=0, column=1, padx=5, ipady=4, sticky='ew')
        
        def handle_complete():
            if original_index is not None and original_index < len(self.todos):
                self.todos[original_index]['status'] = '禁用'
                self.save_todos()
                self.update_todo_list()
                self.log(f"待办事项 '{todo['name']}' 已标记为完成。")
            close_and_release()

        def handle_snooze():
            minutes = simpledialog.askinteger("稍后提醒", "您想在多少分钟后再次提醒？ (1-60)", parent=reminder_win, minvalue=1, maxvalue=60, initialvalue=5)
            if minutes:
                new_remind_time = datetime.now() + timedelta(minutes=minutes)
                if original_index is not None and original_index < len(self.todos):
                    self.todos[original_index]['remind_datetime'] = new_remind_time.strftime('%Y-%m-%d %H:%M:%S')
                    self.todos[original_index]['status'] = '启用'
                    self.save_todos()
                    self.update_todo_list()
                    self.log(f"待办事项 '{todo['name']}' 已推迟 {minutes} 分钟。")
            close_and_release()

        def handle_delete():
            if messagebox.askyesno("确认删除", f"您确定要永久删除待办事项“{todo['name']}”吗？\n此操作不可恢复。", parent=reminder_win):
                if original_index is not None and original_index < len(self.todos):
                    if self.todos[original_index].get('name') == todo.get('name'):
                        self.todos.pop(original_index)
                        self.save_todos()
                        self.update_todo_list()
                        self.log(f"已删除待办事项: {todo['name']}")
                close_and_release()

        def on_closing_protocol():
            if task_type == 'onetime':
                handle_complete()
            else:
                close_and_release()

        reminder_win.protocol("WM_DELETE_WINDOW", on_closing_protocol)
        self.center_window(reminder_win, parent=self.root)

    def _bind_mousewheel_to_entry(self, entry, handler):
        entry.bind("<MouseWheel>", handler)
        entry.bind("<Button-4>", handler)
        entry.bind("<Button-5>", handler)

    def _handle_time_scroll(self, event):
        entry = event.widget
        current_val = entry.get()
        cursor_pos = entry.index(INSERT)

        try:
            dt = datetime.strptime(current_val, "%H:%M:%S")
        except ValueError:
            parts = [p.strip() for p in current_val.split(',') if p.strip()]
            if not parts: return "break"
            
            char_count = 0
            target_part_index = -1
            for i, part in enumerate(parts):
                if char_count <= cursor_pos <= char_count + len(part):
                    target_part_index = i
                    break
                char_count += len(part) + 2
            
            if target_part_index == -1: return "break"

            try:
                dt = datetime.strptime(parts[target_part_index], "%H:%M:%S")
                cursor_pos_in_part = cursor_pos - char_count
            except ValueError:
                return "break"
        else:
            cursor_pos_in_part = cursor_pos

        delta = 1 if event.num == 4 or event.delta > 0 else -1

        if 0 <= cursor_pos_in_part <= 2:
            dt += timedelta(hours=delta)
        elif 3 <= cursor_pos_in_part <= 5:
            dt += timedelta(minutes=delta)
        else:
            dt += timedelta(seconds=delta)

        new_val_part = dt.strftime("%H:%M:%S")
        
        if 'parts' in locals():
            parts[target_part_index] = new_val_part
            new_full_val = ", ".join(parts)
        else:
            new_full_val = new_val_part

        entry.delete(0, END)
        entry.insert(0, new_full_val)
        entry.icursor(cursor_pos)
        return "break"

    def _handle_date_scroll(self, event):
        entry = event.widget
        current_val = entry.get().strip()
        cursor_pos = entry.index(INSERT)

        parts = [p.strip() for p in current_val.split("~")]
        is_range_start = "~" not in current_val or cursor_pos <= len(parts[0])
        target_val = parts[0] if is_range_start else parts[1]

        try:
            dt = datetime.strptime(target_val, "%Y-%m-%d")
        except ValueError:
            return "break"

        delta = 1 if event.num == 4 or event.delta > 0 else -1

        effective_cursor_pos = cursor_pos if is_range_start else cursor_pos - (len(parts[0]) + 3)

        if 0 <= effective_cursor_pos <= 4:
            dt = dt.replace(year=dt.year + delta)
        elif 5 <= effective_cursor_pos <= 7:
            new_month = dt.month + delta
            new_year = dt.year
            if new_month > 12:
                new_month = 1; new_year += 1
            elif new_month < 1:
                new_month = 12; new_year -= 1
            
            try:
                dt = dt.replace(year=new_year, month=new_month)
            except ValueError:
                last_day_of_month = (datetime(new_year, new_month + 1, 1) - timedelta(days=1)).day
                dt = dt.replace(year=new_year, month=new_month, day=min(dt.day, last_day_of_month))
        else:
            dt += timedelta(days=delta)

        new_date_part = dt.strftime("%Y-%m-%d")
        
        if "~" in current_val:
            new_full_val = f"{new_date_part} ~ {parts[1]}" if is_range_start else f"{parts[0]} ~ {new_date_part}"
        else:
            new_full_val = new_date_part

        entry.delete(0, END)
        entry.insert(0, new_full_val)
        entry.icursor(cursor_pos)
        return "break"


def main():
    # 先加载一次设置，以获取保存的主题
    temp_settings = {}
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                temp_settings = json.load(f)
        except:
            pass # 如果加载失败，则使用默认主题

    # 使用保存的主题或默认的 'litera' 来创建窗口
    saved_theme = temp_settings.get("app_theme", "litera")
    root = ttk.Window(themename=saved_theme)
    
    app = TimedBroadcastApp(root)
    root.mainloop()

if __name__ == "__main__":
    if not WIN32_AVAILABLE:
        try:
            messagebox.showerror("核心依赖缺失", "pywin32 库未安装或损坏，软件无法运行语音、注册和锁定等核心功能，即将退出。")
        except:
            print("错误: pywin32 库未安装或损坏，无法显示图形化错误消息。")
        sys.exit(1)
    if not PSUTIL_AVAILABLE:
        try:
            messagebox.showerror("核心依赖缺失", "psutil 库未安装，软件无法获取机器码以进行授权验证，即将退出。")
        except:
            print("错误: psutil 库未安装，无法显示图形化错误消息。")
        sys.exit(1)
    main()

#第14部分
