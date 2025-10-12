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
        self.settings = {}
        self.running = True
        self.tray_icon = None
        self.is_locked = False

        self.is_playing = threading.Event()
        self.playback_queue = []
        self.queue_lock = threading.Lock()
        self.cleanup_lock = threading.Lock()
        
        self.pages = {}
        self.nav_buttons = {}
        self.current_page = None

        self.create_folder_structure()
        self.load_settings()
        self.create_widgets()
        self.load_tasks()
        
        self.start_background_thread()
        self.root.protocol("WM_DELETE_WINDOW", self.show_quit_dialog)
        self.start_tray_icon_thread()
        
        if self.settings.get("lock_on_start", False) and self.settings.get("lock_password_b64", ""):
            self.root.after(100, self.perform_initial_lock)

        if self.settings.get("start_minimized", False):
            self.root.after(100, self.hide_to_tray)

    def create_folder_structure(self):
        """创建所有必要的文件夹"""
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
        elif page_name == "设置":
            if page_name not in self.pages:
                self.pages[page_name] = self.create_settings_page()
            target_frame = self.pages[page_name]
        else:
            messagebox.showinfo("提示", f"页面 [{page_name}] 正在开发中...")
            self.log(f"功能开发中: {page_name}")
            target_frame = self.pages["定时广播"]
            page_name = "定时广播"

        target_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.current_page = target_frame
        
        selected_btn = self.nav_buttons[page_name]
        selected_btn.config(bg='#5DADE2', fg='white')
        selected_btn.master.config(bg='#5DADE2')

    def create_scheduled_broadcast_page(self):
        top_frame = tk.Frame(self.main_frame, bg='white')
        top_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 【UI调整】左侧标题
        left_frame = tk.Frame(top_frame, bg='white')
        left_frame.pack(side=tk.LEFT)
        title_label = tk.Label(left_frame, text="定时广播", font=('Microsoft YaHei', 14, 'bold'), bg='white', fg='#2C5F7C')
        title_label.pack(side=tk.LEFT, anchor='w')
        
        # 【UI调整】新增按钮
        btn_font = ('Microsoft YaHei', 11)
        add_btn = tk.Button(left_frame, text="添加节目", command=self.add_task, bg='#3498DB', fg='white', font=btn_font, bd=0, padx=12, pady=5, cursor='hand2')
        add_btn.pack(side=tk.LEFT, padx=(20, 3))
        
        enable_all_btn = tk.Button(left_frame, text="全部启用", command=self.enable_all_tasks, bg='#2ECC71', fg='white', font=btn_font, bd=0, padx=12, pady=5, cursor='hand2')
        enable_all_btn.pack(side=tk.LEFT, padx=3)

        disable_all_btn = tk.Button(left_frame, text="全部禁用", command=self.disable_all_tasks, bg='#F39C12', fg='white', font=btn_font, bd=0, padx=12, pady=5, cursor='hand2')
        disable_all_btn.pack(side=tk.LEFT, padx=3)
        
        set_volume_btn = tk.Button(left_frame, text="统一音量", command=self.set_uniform_volume, bg='#9B59B6', fg='white', font=btn_font, bd=0, padx=12, pady=5, cursor='hand2')
        set_volume_btn.pack(side=tk.LEFT, padx=3)
        
        clear_all_btn = tk.Button(left_frame, text="清空节目", command=self.clear_all_tasks, bg='#E74C3C', fg='white', font=btn_font, bd=0, padx=12, pady=5, cursor='hand2')
        clear_all_btn.pack(side=tk.LEFT, padx=3)

        # 【UI调整】右侧按钮
        self.top_right_btn_frame = tk.Frame(top_frame, bg='white')
        self.top_right_btn_frame.pack(side=tk.RIGHT)
        
        self.lock_button = tk.Button(self.top_right_btn_frame, text="锁定", command=self.toggle_lock_state, bg='#E74C3C', fg='white', font=btn_font, bd=0, padx=12, pady=5, cursor='hand2')
        self.lock_button.pack(side=tk.LEFT, padx=3)
        
        buttons = [("导入节目单", self.import_tasks, '#1ABC9C'), ("导出节目单", self.export_tasks, '#1ABC9C')]
        for text, cmd, color in buttons:
            btn = tk.Button(self.top_right_btn_frame, text=text, command=cmd, bg=color, fg='white', font=btn_font, bd=0, padx=12, pady=5, cursor='hand2')
            btn.pack(side=tk.LEFT, padx=3)

        stats_frame = tk.Frame(self.main_frame, bg='#F0F8FF')
        stats_frame.pack(fill=tk.X, padx=10, pady=5)
        self.stats_label = tk.Label(stats_frame, text="节目单：0", font=('Microsoft YaHei', 11),
                                   bg='#F0F8FF', fg='#2C5F7C', anchor='w', padx=10)
        self.stats_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        table_frame = tk.Frame(self.main_frame, bg='white')
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        columns = ('节目名称', '状态', '开始时间', '模式', '音频或文字', '音量', '周几/几号', '日期范围')
        self.task_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=12)
        
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

        playing_frame = tk.LabelFrame(self.main_frame, text="正在播：", font=('Microsoft YaHei', 11),
                                     bg='white', fg='#2C5F7C', padx=10, pady=5)
        playing_frame.pack(fill=tk.X, padx=10, pady=5)
        self.playing_text = scrolledtext.ScrolledText(playing_frame, height=3, font=('Microsoft YaHei', 11),
                                                     bg='#FFFEF0', wrap=tk.WORD, state='disabled')
        self.playing_text.pack(fill=tk.BOTH, expand=True)
        self.update_playing_text("等待播放...")

        log_frame = tk.LabelFrame(self.main_frame, text="", font=('Microsoft YaHei', 11),
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

        status_frame = tk.Frame(self.main_frame, bg='#E8F4F8', height=35) # 增加高度
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        status_frame.pack_propagate(False)
        self.status_labels = []
        status_texts = ["当前时间", "系统状态", "播放状态", "任务数量"]
        for i, text in enumerate(status_texts):
            label = tk.Label(status_frame, text=f"{text}: --", font=('Microsoft YaHei', 11), # 字体调整
                           bg='#5DADE2' if i % 2 == 0 else '#7EC8E3', fg='white', padx=15, pady=5)
            label.pack(side=tk.LEFT, padx=2)
            self.status_labels.append(label)

        self.update_status_bar()
        self.log("定时播音软件已启动")
    
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

        self.clear_password_btn = tk.Button(general_frame, text="清除锁定密码", font=('Microsoft YaHei', 11), command=self.clear_lock_password)
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
        dialog.geometry("380x180"); dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()
        self.center_window(dialog, 380, 180)
        
        tk.Label(dialog, text="请输入密码以解锁", font=('Microsoft YaHei', 11)).pack(pady=10)
        
        pass_entry = tk.Entry(dialog, show='*', width=25, font=('Microsoft YaHei', 11))
        pass_entry.pack(pady=5)
        pass_entry.focus_set()

        def clear_and_unlock():
            if check_password():
                if messagebox.askyesno("确认操作", "您确定要清除锁定密码吗？\n此操作不可恢复。", parent=dialog):
                    self.settings["lock_password_b64"] = ""
                    self.lock_on_start_var.set(False)
                    self.save_settings()
                    if hasattr(self, 'clear_password_btn'):
                        self.clear_password_btn.config(state=tk.DISABLED)
                    self.log("锁定密码已清除。")
                    messagebox.showinfo("成功", "锁定密码已成功清除。", parent=dialog)
                    dialog.destroy()
                    self._apply_unlock()

        def check_password():
            entered_pass = pass_entry.get()
            encoded_entered_pass = base64.b64encode(entered_pass.encode('utf-8')).decode('utf-8')
            if encoded_entered_pass == self.settings.get("lock_password_b64"):
                return True
            else:
                messagebox.showerror("错误", "密码不正确！", parent=dialog)
                return False

        def confirm():
            if check_password():
                dialog.destroy()
                self._apply_unlock()

        btn_frame = tk.Frame(dialog); btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="确定", command=confirm, font=('Microsoft YaHei', 11)).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="清除密码", command=clear_and_unlock, font=('Microsoft YaHei', 11)).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="取消", command=dialog.destroy, font=('Microsoft YaHei', 11)).pack(side=tk.LEFT, padx=10)
        dialog.bind('<Return>', lambda event: confirm())

    def clear_lock_password(self):
        if messagebox.askyesno("确认操作", "您确定要清除锁定密码吗？\n此操作不可恢复。"):
            self.settings["lock_password_b64"] = ""
            self.lock_on_start_var.set(False)
            self.save_settings()
            self.clear_password_btn.config(state=tk.DISABLED)
            self.log("锁定密码已清除。")
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
        self._set_widget_state_recursively(self.top_right_btn_frame, state)
        self._set_widget_state_recursively(self.main_frame, state, exclude_children_of=self.top_right_btn_frame)
        self.clear_log_btn.config(state=state)

    def _set_widget_state_recursively(self, parent_widget, state, exclude_children_of=None):
        for child in parent_widget.winfo_children():
            if child == self.lock_button or (exclude_children_of and child in exclude_children_of.winfo_children()):
                continue
            try: child.config(state=state)
            except tk.TclError: pass
            if child.winfo_children():
                self._set_widget_state_recursively(child, state, exclude_children_of=exclude_children_of)
    
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
        with self.cleanup_lock:
            if not self.is_playing.is_set(): return
            if AUDIO_AVAILABLE:
                pygame.mixer.music.stop()
                pygame.mixer.stop()
            self.is_playing.clear()
            self.root.after(0, lambda: self.status_labels[2].config(text="播放状态: 待机"))
            self.root.after(0, lambda: self.update_playing_text("等待播放..."))
            self.log("播放已停止。")

    def play_now(self, task_to_play=None):
        task = None
        if task_to_play:
            task = task_to_play
        else:
            selection = self.task_tree.selection()
            if not selection: messagebox.showwarning("提示", "请先选择一个要立即播放的节目。"); return
            index = self.task_tree.index(selection[0])
            task = self.tasks[index]

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
    
    # ... (other functions)

    def _play_audio_task(self, task):
        try:
            interval_type, duration_seconds, repeat_count = task.get('interval_type'), int(task.get('interval_seconds', 0)), int(task.get('interval_first', 1))
            playlist = []
            if task.get('audio_type') == 'single':
                if os.path.exists(task['content']): playlist = [task['content']] * repeat_count
            else:
                folder_path = task['content']
                if os.path.isdir(folder_path):
                    all_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.lower().endswith(('.mp3', '.wav', '.ogg', '.flac', '.m4a'))]
                    if task.get('play_order') == 'random': random.shuffle(all_files)
                    playlist = all_files[:repeat_count]
            if not playlist: self.log(f"错误: 音频列表为空，任务 '{task['name']}' 无法播放。"); return
            start_time = time.time()
            for audio_path in playlist:
                if not self.is_playing.is_set(): break
                self.log(f"正在播放: {os.path.basename(audio_path)}")
                self.update_playing_text(f"[{task['name']}] 正在播放: {os.path.basename(audio_path)}")
                pygame.mixer.music.load(audio_path)
                pygame.mixer.music.set_volume(float(task.get('volume', 80)) / 100.0)
                pygame.mixer.music.play()
                while pygame.mixer.music.get_busy() and self.is_playing.is_set():
                    if interval_type == 'seconds' and (time.time() - start_time) > duration_seconds: pygame.mixer.music.stop(); self.log(f"已达到 {duration_seconds} 秒播放时长限制。"); break
                    time.sleep(0.1)
                if interval_type == 'seconds' and (time.time() - start_time) > duration_seconds: break
        except Exception as e: self.log(f"音频播放错误: {e}")
        finally:
            self._stop_and_cleanup_playback()
            self.root.after(100, self._process_queue)

    def _play_voice_task(self, task):
        try:
            if not self.is_playing.is_set(): return
            if task.get('prompt', 0) and AUDIO_AVAILABLE:
                prompt_file, prompt_path = task.get('prompt_file', ''), os.path.join(PROMPT_FOLDER, task.get('prompt_file', ''))
                if os.path.exists(prompt_path):
                    if not self.is_playing.is_set(): return
                    self.log(f"播放提示音: {prompt_file}")
                    sound = pygame.mixer.Sound(prompt_path)
                    sound.set_volume(float(task.get('prompt_volume', 80)) / 100.0)
                    channel = sound.play()
                    if channel:
                        while channel.get_busy() and self.is_playing.is_set(): time.sleep(0.05)
                else: self.log(f"警告: 提示音文件不存在 - {prompt_path}")

            if not self.is_playing.is_set(): return
            
            if task.get('bgm', 0) and AUDIO_AVAILABLE:
                bgm_file, bgm_path = task.get('bgm_file', ''), os.path.join(BGM_FOLDER, task.get('bgm_file', ''))
                if os.path.exists(bgm_path):
                    self.log(f"播放背景音乐: {bgm_file}")
                    pygame.mixer.music.load(bgm_path)
                    pygame.mixer.music.set_volume(float(task.get('bgm_volume', 40)) / 100.0)
                    pygame.mixer.music.play(-1)
                else: self.log(f"警告: 背景音乐文件不存在 - {bgm_path}")

            speech_path = task.get('content', '')
            if not os.path.exists(speech_path):
                self.log(f"错误: 语音文件不存在 - {speech_path}"); return

            speech_sound = pygame.mixer.Sound(speech_path)
            speech_sound.set_volume(float(task.get('volume', 80)) / 100.0)
            repeat_count = int(task.get('repeat', 1))

            for i in range(repeat_count):
                if not self.is_playing.is_set(): break
                self.log(f"正在播报第 {i+1}/{repeat_count} 遍")
                channel = speech_sound.play()
                if channel:
                    while channel.get_busy() and self.is_playing.is_set(): time.sleep(0.1)
                if i < repeat_count - 1 and self.is_playing.is_set(): time.sleep(0.5)

        except Exception as e: self.log(f"语音任务播放错误: {e}")
        finally:
            self._stop_and_cleanup_playback()
            self.root.after(100, self._process_queue)
            
    # ... (the rest of your unchanged code)

if __name__ == "__main__":
    main()
