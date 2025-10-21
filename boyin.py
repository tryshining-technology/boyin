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

TASK_FILE = os.path.join(application_path, "broadcast_tasks.json")
SETTINGS_FILE = os.path.join(application_path, "settings.json")
HOLIDAY_FILE = os.path.join(application_path, "holidays.json")
TODO_FILE = os.path.join(application_path, "todos.json")
SCREENSHOT_TASK_FILE = os.path.join(application_path, "screenshot_tasks.json")
EXECUTE_TASK_FILE = os.path.join(application_path, "execute_tasks.json")

PROMPT_FOLDER = os.path.join(application_path, "提示音")
AUDIO_FOLDER = os.path.join(application_path, "音频文件")
BGM_FOLDER = os.path.join(application_path, "文稿背景")
VOICE_SCRIPT_FOLDER = os.path.join(application_path, "语音文稿")
SCREENSHOT_FOLDER = os.path.join(application_path, "截屏")

ICON_FILE = resource_path("icon.ico")
REMINDER_SOUND_FILE = os.path.join(PROMPT_FOLDER, "reminder.wav")
CHIME_FOLDER = os.path.join(AUDIO_FOLDER, "整点报时")

REGISTRY_KEY_PATH = r"Software\创翔科技\TimedBroadcastApp"
REGISTRY_PARENT_KEY_PATH = r"Software\创翔科技"

class TimedBroadcastApp:
    def __init__(self, root):
        self.root = root
        self.root.title(" 创翔多功能定时播音旗舰版")
        # self.root.geometry("1280x720")
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
        
        self.settings = {}
        self.running = True
        self.tray_icon = None
        self.is_locked = False
        self.is_window_pinned = False
        self.is_app_locked_down = False
        self.active_modal_dialog = None # <--- 【BUG修复】新增：追踪活动的模态对话框

        self.auth_info = {'status': 'Unregistered', 'message': '正在验证授权...'}
        self.machine_code = None

        self.lock_password_b64 = ""
        self.drag_start_item = None

        self.playback_command_queue = queue.Queue()
        self.reminder_queue = queue.Queue()
        self.is_reminder_active = False

        self.pages = {}
        self.nav_buttons = {}
        self.current_page = None
        self.current_page_name = ""
        
        self.active_processes = {}

        self.last_chime_hour = -1

        self.fullscreen_window = None
        self.fullscreen_label = None
        self.image_tk_ref = None
        self.current_stop_visual_event = None

        self.video_window = None
        self.vlc_player = None
        self.video_stop_event = None

        self.create_folder_structure()
        self.load_settings()

        # --- ↓↓↓ 新增代码：加载并应用窗口位置和大小 ↓↓↓ ---
        saved_geometry = self.settings.get("window_geometry")
        if saved_geometry:
            try:
                # 检查几何信息是否有效，防止因配置文件损坏导致启动失败
                self.root.geometry(saved_geometry)
            except tk.TclError:
                # 如果信息无效，则使用默认大小
                self.root.geometry("1280x720")
        else:
            # 首次启动或无记录时，设置一个默认大小
            self.root.geometry("1280x720")
        # --- ↑↑↑ 新增代码结束 ↑↑↑ ---

        self.load_lock_password()

        self._apply_global_font()
        self.check_authorization()

        self.create_widgets()
        self.load_tasks()
        self.load_holidays()
        self.load_todos()
        self.load_screenshot_tasks()
        self.load_execute_tasks()

        self.start_background_threads()
        self.root.protocol("WM_DELETE_WINDOW", self.show_quit_dialog)
        self.start_tray_icon_thread()

        # --- ↓↓↓ 【BUG修复】新增：启动状态监视器 ↓↓↓ ---
        self.root.after(250, self._poll_window_state)
        # --- ↑↑↑ 【BUG修复】结束 ↑↑↑ ---

        if self.settings.get("lock_on_start", False) and self.lock_password_b64:
            self.root.after(100, self.perform_initial_lock)
        if self.settings.get("start_minimized", False):
            self.root.after(100, self.hide_to_tray)
        if self.is_app_locked_down:
            self.root.after(100, self.perform_lockdown)

    # --- ↓↓↓ 【BUG修复】新增：主动轮询状态的监视器函数 ↓↓↓ ---
    def _poll_window_state(self):
        """
        通过主动轮询来监视主窗口状态，并同步模态对话框。
        这是一个比事件绑定更可靠的方法，可以绕过 grab_set() 的事件阻塞。
        """
        try:
            current_state = self.root.state()
        except tk.TclError:
            # 在程序关闭过程中，winfo_exists() 可能还返回True，但state()会失败
            return

        # 只有当状态发生变化时才执行操作
        if current_state != self._last_root_state:
            # 检查是否有活动的模态对话框
            if self.active_modal_dialog and self.active_modal_dialog.winfo_exists():
                
                # 如果主窗口被最小化了
                if current_state == 'iconic':
                    self.active_modal_dialog.withdraw()
                
                # 如果主窗口恢复正常了
                elif current_state == 'normal':
                    self.active_modal_dialog.deiconify()
            
            # 更新最后的状态记录
            self._last_root_state = current_state

        # 安排下一次检查
        if self.running:
            self.root.after(250, self._poll_window_state)
    # --- ↑↑↑ 【BUG修复】结束 ↑↑↑ ---

    def _apply_global_font(self):
        font_name = self.settings.get("app_font", "Microsoft YaHei")
        try:
            if font_name not in font.families():
                self.log(f"警告：字体 '{font_name}' 未在系统中找到，已回退至默认字体。")
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
            VOICE_SCRIPT_FOLDER, SCREENSHOT_FOLDER
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

        nav_button_titles = ["定时广播", "节假日", "待办事项", "高级功能", "设置", "注册软件", "超级管理"]

        for i, title in enumerate(nav_button_titles):
            is_super_admin = (title == "超级管理")
            cmd = (lambda t=title: self._prompt_for_super_admin_password()) if is_super_admin else (lambda t=title: self.switch_page(t))
            
            btn = ttk.Button(self.nav_frame, text=title, bootstyle="light",
                           style='Link.TButton', command=cmd)
            btn.pack(fill=X, pady=1, ipady=8, padx=5)
            self.nav_buttons[title] = btn
            
        style = ttk.Style.get_instance()
        style.configure('Link.TButton', font=self.font_13_bold, anchor='w')

        self.main_frame = ttk.Frame(self.page_container)
        self.pages["定时广播"] = self.main_frame
        self.create_scheduled_broadcast_page()
        # vvvvvv 在这里添加下面的代码 vvvvvv
        # --- 【核心修复】预创建高级功能页面 ---
        advanced_page = self.create_advanced_features_page()
        self.pages["高级功能"] = advanced_page
        # 预创建后立即隐藏它
        advanced_page.pack_forget()
        # --- 修复结束 ---
        # ^^^^^^ 在这里添加上面的代码 ^^^^^^

        self.current_page = self.main_frame
        self.switch_page("定时广播")

        self.update_status_bar()
        self.log("创翔多功能定时播音旗舰版软件已启动")

    def create_status_bar_content(self):
        self.status_labels = []
        status_texts = ["当前时间", "系统状态", "播放状态", "任务数量", "待办事项"]

        copyright_label = ttk.Label(self.status_frame, text="© 创翔科技", font=self.font_11,
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

        # 隐藏当前页面
        if self.current_page and self.current_page.winfo_exists():
            self.current_page.pack_forget()

        # 取消所有按钮的高亮
        for title, btn in self.nav_buttons.items():
            btn.config(bootstyle="light")

        # --- 【核心修复】简化页面切换逻辑 ---
        target_frame = None
        # 首先，尝试从已经创建的页面字典中获取
        if page_name in self.pages and self.pages[page_name].winfo_exists():
            target_frame = self.pages[page_name]
        else:
            # 如果字典里没有（只应该发生在节假日、设置等页面第一次被点击时），就创建它
            if page_name == "节假日":
                target_frame = self.create_holiday_page()
            elif page_name == "待办事项":
                target_frame = self.create_todo_page()
            elif page_name == "设置":
                target_frame = self.create_settings_page()
            elif page_name == "注册软件":
                target_frame = self.create_registration_page()
            elif page_name == "超级管理":
                target_frame = self.create_super_admin_page()
            
            # 创建后，存入字典以便下次使用
            if target_frame:
                self.pages[page_name] = target_frame

        # 如果最终没有找到页面，就回到默认页面
        if not target_frame:
            self.log(f"错误或开发中: 无法找到页面 '{page_name}'，返回主页。")
            target_frame = self.pages["定时广播"]
            page_name = "定时广播"
        
        # 显示目标页面
        target_frame.pack(in_=self.page_container, fill=BOTH, expand=True)
        self.current_page = target_frame
        self.current_page_name = page_name

        # 更新设置页面的UI（如果切换到设置页）
        if page_name == "设置":
            self._refresh_settings_ui()

        # 高亮当前按钮
        selected_btn = self.nav_buttons.get(page_name)
        if selected_btn:
            selected_btn.config(bootstyle="primary")

    def _prompt_for_super_admin_password(self):
        if self.auth_info['status'] != 'Permanent':
            messagebox.showerror("权限不足", "此功能仅对“永久授权”用户开放。\n\n请注册软件并获取永久授权后重试。", parent=self.root)
            self.log("非永久授权用户尝试进入超级管理模块被阻止。")
            return

        dialog = ttk.Toplevel(self.root)
        self.active_modal_dialog = dialog # <--- 【BUG修复】
        dialog.title("身份验证")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        result = [None]

        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            dialog.destroy()

        ttk.Label(dialog, text="请输入超级管理员密码:", font=self.font_11).pack(pady=20, padx=20)
        password_entry = ttk.Entry(dialog, show='*', font=self.font_11, width=25)
        password_entry.pack(pady=5, padx=20)
        password_entry.focus_set()

        def on_confirm():
            result[0] = password_entry.get()
            cleanup_and_destroy() # <--- 【BUG修复】

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="确定", command=on_confirm, bootstyle="primary", width=8).pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=cleanup_and_destroy, width=8).pack(side=LEFT, padx=10) # <--- 【BUG修复】
        dialog.bind('<Return>', lambda event: on_confirm())
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】

        self.center_window(dialog, parent=self.root)
        self.root.wait_window(dialog)
        
        entered_password = result[0]
        correct_password = datetime.now().strftime('%Y%m%d')

        if entered_password == correct_password:
            self.log("超级管理员密码正确，进入管理模块。")
            self.switch_page("超级管理")
        elif entered_password is not None:
            messagebox.showerror("验证失败", "密码错误！", parent=self.root)
            self.log("尝试进入超级管理模块失败：密码错误。")
            
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

        notebook.add(screenshot_tab, text=' 定时截屏 ')
        notebook.add(execute_tab, text=' 定时运行 ')

        self._build_screenshot_ui(screenshot_tab)
        self._build_execute_ui(execute_tab)

        return page_frame

    def enable_all_screenshot(self):
        """启用所有的定时截屏任务"""
        if not self.screenshot_tasks: return
        for task in self.screenshot_tasks:
            task['status'] = '启用'
        self.update_screenshot_list()
        self.save_screenshot_tasks()
        self.log("已将 *全部* 截屏任务的状态设置为: 启用")

    def disable_all_screenshot(self):
        """禁用所有的定时截屏任务"""
        if not self.screenshot_tasks: return
        for task in self.screenshot_tasks:
            task['status'] = '禁用'
        self.update_screenshot_list()
        self.save_screenshot_tasks()
        self.log("已将 *全部* 截屏任务的状态设置为: 禁用")

    def enable_all_execute(self):
        """启用所有的定时运行任务"""
        if not self.execute_tasks: return
        for task in self.execute_tasks:
            task['status'] = '启用'
        self.update_execute_list()
        self.save_execute_tasks()
        self.log("已将 *全部* 运行任务的状态设置为: 启用")

    def disable_all_execute(self):
        """禁用所有的定时运行任务"""
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
                task.get('stop_time', ''), # 确保这里有 stop_time
                task.get('repeat_count', 1), # 确保这里有 repeat_count
                task.get('interval_seconds', 0), # 确保这里有 interval_seconds
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

    def open_screenshot_dialog(self, task_to_edit=None, index=None):
        dialog = ttk.Toplevel(self.root)
        self.active_modal_dialog = dialog # <--- 【BUG修复】
        dialog.title("修改截屏任务" if task_to_edit else "添加截屏任务")
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()

        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            dialog.destroy()

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
        ttk.Label(time_frame, text="《可多个,用英文逗号,隔开》").grid(row=0, column=2, sticky='w', padx=5)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bootstyle="outline").grid(row=0, column=3, padx=5)

        # 增加停止时间输入框
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
            date_range_entry.insert(0, task_to_edit.get('date_range', '2000-01-01 ~ 2099-12-31'))
        else:
            repeat_entry.insert(0, '1')
            interval_entry.insert(0, '0')
            weekday_entry.insert(0, "每周:1234567")
            date_range_entry.insert(0, "2000-01-01 ~ 2099-12-31")

        def save_task():
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
            cleanup_and_destroy() # <--- 【BUG修复】

        button_text = "保存修改" if task_to_edit else "添加"
        ttk.Button(dialog_button_frame, text=button_text, command=save_task, bootstyle="primary").pack(side=LEFT, padx=10, ipady=5)
        ttk.Button(dialog_button_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10, ipady=5) # <--- 【BUG修复】
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】
        
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

    def open_execute_dialog(self, task_to_edit=None, index=None):
        dialog = ttk.Toplevel(self.root)
        self.active_modal_dialog = dialog # <--- 【BUG修复】
        dialog.title("修改运行任务" if task_to_edit else "添加运行任务")
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()

        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            dialog.destroy()

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
        ttk.Label(time_frame, text="《可多个,用英文逗号,隔开》").grid(row=0, column=2, sticky='w', padx=5)
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
                  bootstyle="inverse-danger", wraplength=450, justify=LEFT).pack(fill=X)

        dialog_button_frame = ttk.Frame(dialog)
        dialog_button_frame.pack(pady=15)

        if task_to_edit:
            name_entry.insert(0, task_to_edit.get('name', ''))
            target_entry.insert(0, task_to_edit.get('target_path', ''))
            args_entry.insert(0, task_to_edit.get('arguments', ''))
            start_time_entry.insert(0, task_to_edit.get('time', ''))
            stop_time_entry.insert(0, task_to_edit.get('stop_time', ''))
            weekday_entry.insert(0, task_to_edit.get('weekday', '每周:1234567'))
            date_range_entry.insert(0, task_to_edit.get('date_range', '2000-01-01 ~ 2099-12-31'))
        else:
            weekday_entry.insert(0, "每周:1234567")
            date_range_entry.insert(0, "2000-01-01 ~ 2099-12-31")

        def save_task():
            target_path = target_entry.get().strip()
            if not target_path:
                messagebox.showerror("输入错误", "目标程序路径不能为空。", parent=dialog)
                return

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
            cleanup_and_destroy() # <--- 【BUG修复】

        button_text = "保存修改" if task_to_edit else "添加"
        ttk.Button(dialog_button_frame, text=button_text, command=save_task, bootstyle="primary").pack(side=LEFT, padx=10, ipady=5)
        ttk.Button(dialog_button_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10, ipady=5) # <--- 【BUG修复】
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】
        
        self.center_window(dialog, parent=self.root)

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
#第2部分
    def cancel_registration(self):
        if not messagebox.askyesno("确认操作", "您确定要取消当前注册吗？\n取消后，软件将恢复到试用或过期状态。", parent=self.root):
            return

        self.log("用户请求取消注册...")
        self._save_to_registry('RegistrationStatus', '')
        self._save_to_registry('RegistrationDate', '')

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
        interfaces = psutil.net_if_addrs()
        stats = psutil.net_if_stats()

        wired_macs = []
        wireless_macs = []
        other_macs = []

        wired_keywords = ['ethernet', 'eth', '本地连接', 'local area connection']
        wireless_keywords = ['wi-fi', 'wlan', '无线网络连接', 'wireless']

        for name, addrs in interfaces.items():
            is_wired = any(keyword in name.lower() for keyword in wired_keywords)
            is_wireless = any(keyword in name.lower() for keyword in wireless_keywords)

            is_up = stats.get(name) and getattr(stats.get(name), 'isup', False)

            for addr in addrs:
                if addr.family == psutil.AF_LINK:
                    mac = addr.address.replace(':', '').replace('-', '').upper()
                    if len(mac) == 12 and mac != '000000000000':
                        mac_info = {'mac': mac, 'is_up': is_up, 'name': name}
                        if is_wired:
                            wired_macs.append(mac_info)
                        elif is_wireless:
                            wireless_macs.append(mac_info)
                        else:
                            other_macs.append(mac_info)

        wired_macs.sort(key=lambda x: x['is_up'], reverse=True)
        wireless_macs.sort(key=lambda x: x['is_up'], reverse=True)
        other_macs.sort(key=lambda x: x['is_up'], reverse=True)

        if wired_macs:
            return wired_macs[0]['mac']
        if wireless_macs:
            return wireless_macs[0]['mac']
        if other_macs:
            return other_macs[0]['mac']

        for name, addrs in interfaces.items():
            for addr in addrs:
                if addr.family == psutil.AF_LINK:
                    mac = addr.address.replace(':', '').replace('-', '').upper()
                    if len(mac) == 12 and mac != '000000000000':
                         return mac
        return None

    def _calculate_reg_codes(self, numeric_mac_str):
        try:
            monthly_code = int(int(numeric_mac_str) * 3.14)

            reversed_mac_str = numeric_mac_str[::-1]
            permanent_val = int(reversed_mac_str) / 3.14
            permanent_code = f"{permanent_val:.2f}"

            return {'monthly': str(monthly_code), 'permanent': permanent_code}
        except (ValueError, TypeError):
            return {'monthly': None, 'permanent': None}

    def attempt_registration(self):
        entered_code = self.reg_code_entry.get().strip()
        if not entered_code:
            messagebox.showwarning("提示", "请输入注册码。", parent=self.root)
            return

        numeric_machine_code = self.get_machine_code()
        correct_codes = self._calculate_reg_codes(numeric_machine_code)

        today_str = datetime.now().strftime('%Y-%m-%d')

        if entered_code == correct_codes['monthly']:
            self._save_to_registry('RegistrationStatus', 'Monthly')
            self._save_to_registry('RegistrationDate', today_str)
            messagebox.showinfo("注册成功", "恭喜您，月度授权已成功激活！", parent=self.root)
            self.check_authorization()
        elif entered_code == correct_codes['permanent']:
            self._save_to_registry('RegistrationStatus', 'Permanent')
            self._save_to_registry('RegistrationDate', today_str)
            messagebox.showinfo("注册成功", "恭喜您，永久授权已成功激活！", parent=self.root)
            self.check_authorization()
        else:
            messagebox.showerror("注册失败", "您输入的注册码无效，请重新核对。", parent=self.root)

    def check_authorization(self):
        today = datetime.now().date()
        status = self._load_from_registry('RegistrationStatus')
        reg_date_str = self._load_from_registry('RegistrationDate')

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
                self.auth_info = {'status': 'Expired', 'message': '授权信息损坏，请重新注册'}
                self.is_app_locked_down = True
        else:
            first_run_date_str = self._load_from_registry('FirstRunDate')
            if not first_run_date_str:
                self._save_to_registry('FirstRunDate', today.strftime('%Y-%m-%d'))
                self.auth_info = {'status': 'Trial', 'message': '未注册 - 剩余 3 天'}
                self.is_app_locked_down = False
            else:
                try:
                    first_run_date = datetime.strptime(first_run_date_str, '%Y-%m-%d').date()
                    trial_expiry_date = first_run_date + timedelta(days=3)
                    if today > trial_expiry_date:
                        self.auth_info = {'status': 'Expired', 'message': '授权已过期，请注册'}
                        self.is_app_locked_down = True
                    else:
                        remaining_days = (trial_expiry_date - today).days
                        self.auth_info = {'status': 'Trial', 'message': f'未注册 - 剩余 {remaining_days} 天'}
                        self.is_app_locked_down = False
                except (TypeError, ValueError):
                    self.auth_info = {'status': 'Expired', 'message': '授权信息损坏，请重新注册'}
                    self.is_app_locked_down = True

        self.update_title_bar()

    def perform_lockdown(self):
        messagebox.showerror("授权过期", "您的软件试用期或授权已到期，功能已受限。\n请在“注册软件”页面输入有效注册码以继续使用。", parent=self.root)
        self.log("软件因授权问题被锁定。")

        for task in self.tasks:
            task['status'] = '禁用'
        self.update_task_list()
        self.save_tasks()

        self.switch_page("注册软件")

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
        self.active_modal_dialog = dialog # <--- 【BUG修复】
        dialog.title("卸载软件 - 身份验证")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        result = [None]

        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            dialog.destroy()

        ttk.Label(dialog, text="请输入卸载密码:", font=self.font_11).pack(pady=20, padx=20)
        password_entry = ttk.Entry(dialog, show='*', font=self.font_11, width=25)
        password_entry.pack(pady=5, padx=20)
        password_entry.focus_set()

        def on_confirm():
            result[0] = password_entry.get()
            cleanup_and_destroy() # <--- 【BUG修复】

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="确定", command=on_confirm, bootstyle="primary", width=8).pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=cleanup_and_destroy, width=8).pack(side=LEFT, padx=10) # <--- 【BUG修复】
        dialog.bind('<Return>', lambda event: on_confirm())
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】
        
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
            self.settings = backup_data['settings']
            self.lock_password_b64 = backup_data['lock_password_b64']

            self.save_tasks()
            self.save_holidays()
            self.save_todos()
            self.save_screenshot_tasks()
            self.save_execute_tasks()
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

        # --- 顶部控件 ---
        top_frame = ttk.Frame(page_frame, padding=(10, 10))
        top_frame.pack(side=TOP, fill=X)
        
        title_label = ttk.Label(top_frame, text="定时广播", font=self.font_14_bold)
        title_label.pack(side=LEFT)

        add_btn = ttk.Button(top_frame, text="添加节目", command=self.add_task, bootstyle="primary")
        add_btn.pack(side=LEFT, padx=10)

        # --- ↓↓↓ 核心修改区域开始 ↓↓↓ ---

        # 创建一个总的按钮容器，放置在最右侧
        top_right_container = ttk.Frame(top_frame)
        top_right_container.pack(side=RIGHT)

        # 创建第一行按钮的容器
        button_row_1 = ttk.Frame(top_right_container)
        button_row_1.pack(fill=X, anchor='e')

        # 创建第二行按钮的容器
        button_row_2 = ttk.Frame(top_right_container)
        button_row_2.pack(fill=X, anchor='e', pady=(5, 0))

        # 定义第一行的按钮
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

        # 定义第二行的按钮
        batch_buttons_row2 = [
            ("统一音量", self.set_uniform_volume, 'info'),
            ("清空节目", self.clear_all_tasks, 'danger'),
            ("导入节目单", self.import_tasks, 'info-outline'),
            ("导出节目单", self.export_tasks, 'info-outline'),
        ]
        for text, cmd, style in batch_buttons_row2:
            btn = ttk.Button(button_row_2, text=text, command=cmd, bootstyle=style)
            btn.pack(side=LEFT, padx=3)
            
        # 在第二行末尾单独添加“置顶”和“锁定”按钮
        self.pin_button = ttk.Button(button_row_2, text="置顶", command=self.toggle_pin_state, bootstyle="info-outline")
        self.pin_button.pack(side=LEFT, padx=3)
        
        self.lock_button = ttk.Button(button_row_2, text="锁定", command=self.toggle_lock_state, bootstyle='danger')
        self.lock_button.pack(side=LEFT, padx=3)
        if not WIN32_AVAILABLE:
            self.lock_button.config(state=DISABLED, text="锁定(Win)")

        # --- ↑↑↑ 核心修改区域结束 ↑↑↑ ---

        stats_frame = ttk.Frame(page_frame, padding=(10, 5))
        stats_frame.pack(side=TOP, fill=X)
        self.stats_label = ttk.Label(stats_frame, text="节目单：0", font=self.font_11, bootstyle="secondary")
        self.stats_label.pack(side=LEFT, fill=X, expand=True)

        # --- 底部控件 (采用逆序 pack 技巧) ---
        log_frame = ttk.LabelFrame(page_frame, text="", padding=(10, 5))
        log_frame.pack(side=BOTTOM, fill=X, padx=10, pady=5)

        playing_frame = ttk.LabelFrame(page_frame, text="正在播：", padding=(10, 5))
        playing_frame.pack(side=BOTTOM, fill=X, padx=10, pady=5)
        
        table_frame = ttk.Frame(page_frame, padding=(10, 5))
        table_frame.pack(side=TOP, fill=BOTH, expand=True)

        # --- 填充各个区域的内容 ---
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
#第3部分
    def create_settings_page(self):
        settings_frame = ttk.Frame(self.page_container, padding=20)

        title_label = ttk.Label(settings_frame, text="系统设置", font=self.font_14_bold, bootstyle="primary")
        title_label.pack(anchor=W, pady=(0, 10))

        general_frame = ttk.LabelFrame(settings_frame, text="通用设置", padding=(15, 10))
        general_frame.pack(fill=X, pady=10)

        self.autostart_var = ttk.BooleanVar()
        self.start_minimized_var = ttk.BooleanVar()
        self.lock_on_start_var = ttk.BooleanVar()
        self.bg_image_interval_var = ttk.StringVar()

        ttk.Checkbutton(general_frame, text="登录windows后自动启动", variable=self.autostart_var, bootstyle="round-toggle", command=self._handle_autostart_setting).pack(fill=X, pady=5)
        ttk.Checkbutton(general_frame, text="启动后最小化到系统托盘", variable=self.start_minimized_var, bootstyle="round-toggle", command=self.save_settings).pack(fill=X, pady=5)

        # 使用一个新的 Frame 来容纳“启动锁定”和它的提示
        lock_on_start_frame = ttk.Frame(general_frame)
        lock_on_start_frame.pack(fill=X, pady=5)

        self.lock_on_start_cb = ttk.Checkbutton(lock_on_start_frame, text="启动软件后立即锁定", variable=self.lock_on_start_var, bootstyle="round-toggle", command=self._handle_lock_on_start_toggle)
        self.lock_on_start_cb.pack(side=LEFT)
        if not WIN32_AVAILABLE:
            self.lock_on_start_cb.config(state=DISABLED)

        # 将提示标签放在 Checkbutton 右侧
        ttk.Label(lock_on_start_frame, text="(请先在主界面设置锁定密码)", font=self.font_9, bootstyle="secondary").pack(side=LEFT, padx=10, anchor='w')

        # --- ↓↓↓ 修改部分 ↓↓↓ ---
        # 将“清除密码”按钮移动到标签后面，并设置样式
        self.clear_password_btn = ttk.Button(
            lock_on_start_frame,  # <--- 1. 父容器改为 lock_on_start_frame
            text="清除锁定密码", 
            command=self.clear_lock_password, 
            bootstyle="danger-link"  # <--- 2. 样式改为 danger-link 使文字变红
        )
        self.clear_password_btn.pack(side=LEFT, padx=10) # <--- 3. 布局改为 side=LEFT
        # --- ↑↑↑ 修改结束 ↑↑↑ ---


        # --- ↓↓↓ 核心修改区域 2 开始 ↓↓↓ ---
        
        bg_interval_frame = ttk.Frame(general_frame)
        bg_interval_frame.pack(fill=X, pady=8)
        
        ttk.Label(bg_interval_frame, text="背景图片切换间隔:").pack(side=LEFT)
        interval_entry = ttk.Entry(bg_interval_frame, textvariable=self.bg_image_interval_var, font=self.font_11, width=5)
        interval_entry.pack(side=LEFT, padx=5)
        ttk.Label(bg_interval_frame, text="秒 (范围: 5-60)", font=self.font_10, bootstyle="secondary").pack(side=LEFT)
        ttk.Button(bg_interval_frame, text="确定", command=self._validate_bg_interval, bootstyle="primary-outline").pack(side=LEFT, padx=10)

        # 将两个操作按钮放在“确定”按钮的右侧
        self.cancel_bg_images_btn = ttk.Button(bg_interval_frame, text="取消所有节目背景图片", command=self._cancel_all_background_images, bootstyle="info-outline")
        self.cancel_bg_images_btn.pack(side=LEFT, padx=5)
        
        self.restore_video_speed_btn = ttk.Button(bg_interval_frame, text="恢复所有视频节目播放速度", command=self._restore_all_video_speeds, bootstyle="info-outline")
        self.restore_video_speed_btn.pack(side=LEFT, padx=5)

        # --- ↑↑↑ 核心修改区域 2 结束 ↑↑↑ ---

        font_frame = ttk.Frame(general_frame)
        font_frame.pack(fill=X, pady=8)

        ttk.Label(font_frame, text="软件字体:").pack(side=LEFT)

        try:
            available_fonts = sorted(list(font.families()))
        except:
            available_fonts = ["Microsoft YaHei"]

        self.font_var = ttk.StringVar()

        font_combo = ttk.Combobox(font_frame, textvariable=self.font_var, values=available_fonts, font=self.font_10, width=25, state='readonly')
        font_combo.pack(side=LEFT, padx=10)
        font_combo.bind("<<ComboboxSelected>>", self._on_font_selected)

        restore_font_btn = ttk.Button(font_frame, text="恢复默认字体", command=self._restore_default_font, bootstyle="secondary-outline")
        restore_font_btn.pack(side=LEFT, padx=10)

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

        daily_frame = ttk.Frame(power_frame)
        daily_frame.pack(fill=X, pady=4)
        daily_frame.columnconfigure(1, weight=1)
        ttk.Checkbutton(daily_frame, text="每天关机    ", variable=self.daily_shutdown_enabled_var, bootstyle="round-toggle", command=self.save_settings).grid(row=0, column=0, sticky='w')
        daily_time_entry = ttk.Entry(daily_frame, textvariable=self.daily_shutdown_time_var, font=self.font_11)
        daily_time_entry.grid(row=0, column=1, sticky='we', padx=5)
        self._bind_mousewheel_to_entry(daily_time_entry, self._handle_time_scroll)
        ttk.Button(daily_frame, text="设置", bootstyle="primary-outline", command=lambda: self.show_single_time_dialog(self.daily_shutdown_time_var)).grid(row=0, column=2, sticky='e', padx=5)

        weekly_frame = ttk.Frame(power_frame)
        weekly_frame.pack(fill=X, pady=4)
        weekly_frame.columnconfigure(1, weight=1)
        ttk.Checkbutton(weekly_frame, text="每周关机    ", variable=self.weekly_shutdown_enabled_var, bootstyle="round-toggle", command=self.save_settings).grid(row=0, column=0, sticky='w')
        weekly_days_entry = ttk.Entry(weekly_frame, textvariable=self.weekly_shutdown_days_var, font=self.font_11)
        weekly_days_entry.grid(row=0, column=1, sticky='we', padx=5)
        weekly_shutdown_time_entry = ttk.Entry(weekly_frame, textvariable=self.weekly_shutdown_time_var, font=self.font_11, width=15)
        weekly_shutdown_time_entry.grid(row=0, column=2, sticky='we', padx=5)
        self._bind_mousewheel_to_entry(weekly_shutdown_time_entry, self._handle_time_scroll)
        ttk.Button(weekly_frame, text="设置", bootstyle="primary-outline", command=lambda: self.show_power_week_time_dialog("设置每周关机", self.weekly_shutdown_days_var, self.weekly_shutdown_time_var)).grid(row=0, column=3, sticky='e', padx=5)

        reboot_frame = ttk.Frame(power_frame)
        reboot_frame.pack(fill=X, pady=4)
        reboot_frame.columnconfigure(1, weight=1)
        ttk.Checkbutton(reboot_frame, text="每周重启    ", variable=self.weekly_reboot_enabled_var, bootstyle="round-toggle", command=self.save_settings).grid(row=0, column=0, sticky='w')
        ttk.Entry(reboot_frame, textvariable=self.weekly_reboot_days_var, font=self.font_11).grid(row=0, column=1, sticky='we', padx=5)
        weekly_reboot_time_entry = ttk.Entry(reboot_frame, textvariable=self.weekly_reboot_time_var, font=self.font_11, width=15)
        weekly_reboot_time_entry.grid(row=0, column=2, sticky='we', padx=5)
        self._bind_mousewheel_to_entry(weekly_reboot_time_entry, self._handle_time_scroll)
        ttk.Button(reboot_frame, text="设置", bootstyle="primary-outline", command=lambda: self.show_power_week_time_dialog("设置每周重启", self.weekly_reboot_days_var, self.weekly_reboot_time_var)).grid(row=0, column=3, sticky='e', padx=5)

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
            self.active_modal_dialog = progress_dialog # <--- 【BUG修复】
            progress_dialog.title("请稍候")
            progress_dialog.resizable(False, False)
            progress_dialog.transient(self.root); progress_dialog.grab_set()

            def cleanup_and_destroy(): # <--- 【BUG修复】
                self.active_modal_dialog = None
                progress_dialog.destroy()

            progress_dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】

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
            self.root.after(1, lambda: setattr(self, 'active_modal_dialog', None)) # <--- 【BUG修复】
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
        """切换窗口的置顶状态"""
        # 翻转当前的置顶状态
        self.is_window_pinned = not self.is_window_pinned
        
        if self.is_window_pinned:
            # 如果是True，则执行置顶操作
            self.root.attributes('-topmost', True)
            # 更新按钮的文本和样式，以便用户知道下一步是“取消置顶”
            self.pin_button.config(text="取消置顶", bootstyle="info")
            self.log("窗口已置顶显示。")
        else:
            # 如果是False，则执行取消置顶操作
            self.root.attributes('-topmost', False)
            # 恢复按钮的初始状态
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
        self.active_modal_dialog = dialog # <--- 【BUG修复】
        dialog.title("首次锁定，请设置密码")
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()

        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            dialog.destroy()

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
                cleanup_and_destroy() # <--- 【BUG修复】
                self._apply_lock()
            else:
                messagebox.showerror("功能受限", "无法保存密码。\n此功能仅在Windows系统上支持且需要pywin32库。", parent=dialog)

        btn_frame = ttk.Frame(dialog); btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10) # <--- 【BUG修复】
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】
        
        self.center_window(dialog, parent=self.root)

    def _prompt_for_password_unlock(self):
        dialog = ttk.Toplevel(self.root)
        self.active_modal_dialog = dialog # <--- 【BUG修复】
        dialog.title("解锁界面")
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()

        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            dialog.destroy()

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
                cleanup_and_destroy() # <--- 【BUG修复】
                self._apply_unlock()
            else:
                messagebox.showerror("错误", "密码不正确！", parent=dialog)

        def clear_password_action():
            if not is_password_correct():
                messagebox.showerror("错误", "密码不正确！无法清除。", parent=dialog)
                return

            if messagebox.askyesno("确认操作", "您确定要清除锁定密码吗？\n此操作不可恢复。", parent=dialog):
                self._perform_password_clear_logic()
                cleanup_and_destroy() # <--- 【BUG修复】
                self.root.after(50, self._apply_unlock)
                self.root.after(100, lambda: messagebox.showinfo("成功", "锁定密码已成功清除。", parent=self.root))

        btn_frame = ttk.Frame(dialog); btn_frame.pack(pady=20, padx=10, fill=X, expand=True)
        btn_frame.columnconfigure((0, 1, 2), weight=1)
        ttk.Button(btn_frame, text="确定", command=confirm, bootstyle="primary").grid(row=0, column=0, padx=5, sticky='ew')
        ttk.Button(btn_frame, text="清除密码", command=clear_password_action, bootstyle="warning").grid(row=0, column=1, padx=5, sticky='ew')
        ttk.Button(btn_frame, text="取消", command=cleanup_and_destroy).grid(row=0, column=2, padx=5, sticky='ew') # <--- 【BUG修复】
        dialog.bind('<Return>', lambda event: confirm())
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】
        
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
        choice_dialog = ttk.Toplevel(self.root)
        self.active_modal_dialog = choice_dialog # <--- 【BUG修复】
        choice_dialog.title("选择节目类型")
        choice_dialog.resizable(False, False)
        choice_dialog.transient(self.root); choice_dialog.grab_set()
        
        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            choice_dialog.destroy()

        def open_and_cleanup(dialog_opener_func): # <--- 【BUG修复】
            # cleanup_and_destroy() # This is now handled by the new dialogs
            dialog_opener_func(choice_dialog)

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

        video_btn = ttk.Button(btn_frame, text="🎬→视频节目",
                             bootstyle="success", width=20, command=lambda: open_and_cleanup(self.open_video_dialog))
        video_btn.pack(pady=8, ipady=8, fill=X)
        if not VLC_AVAILABLE:
            video_btn.config(state=DISABLED, text="🎬→视频节目 (VLC未安装)")

        choice_dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】
        self.center_window(choice_dialog, parent=self.root)
#第5部分
#第5部分
    def open_audio_dialog(self, parent_dialog, task_to_edit=None, index=None):
        parent_dialog.destroy()
        is_edit_mode = task_to_edit is not None
        dialog = ttk.Toplevel(self.root)
        self.active_modal_dialog = dialog # <--- 【BUG修复】
        dialog.title("修改音频节目" if is_edit_mode else "添加音频节目")
        dialog.resizable(True, True)
        dialog.minsize(800, 580)
        dialog.transient(self.root); dialog.grab_set()

        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            dialog.destroy()

        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=BOTH, expand=True)

        content_frame = ttk.LabelFrame(main_frame, text="内容", padding=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=2)
        content_frame.columnconfigure(1, weight=1)

        ttk.Label(content_frame, text="节目名称:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        name_entry = ttk.Entry(content_frame, font=self.font_11)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        audio_type_var = tk.StringVar(value="single")
        ttk.Label(content_frame, text="音频文件").grid(row=1, column=0, sticky='e', padx=5, pady=2)
        audio_single_frame = ttk.Frame(content_frame)
        audio_single_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        audio_single_frame.columnconfigure(1, weight=1)
        ttk.Radiobutton(audio_single_frame, text="", variable=audio_type_var, value="single").grid(row=0, column=0, sticky='w')
        audio_single_entry = ttk.Entry(audio_single_frame, font=self.font_11)
        audio_single_entry.grid(row=0, column=1, sticky='ew', padx=5)
        ttk.Label(audio_single_frame, text="00:00").grid(row=0, column=2, padx=10)
        def select_single_audio():
            filename = filedialog.askopenfilename(title="选择音频文件", initialdir=AUDIO_FOLDER, filetypes=[("音频文件", "*.mp3 *.wav *.ogg *.flac *.m4a"), ("所有文件", "*.*")], parent=dialog)
            if filename: audio_single_entry.delete(0, END); audio_single_entry.insert(0, filename)
        ttk.Button(audio_single_frame, text="选取...", command=select_single_audio, bootstyle="outline").grid(row=0, column=3, padx=5)
        
        ttk.Label(content_frame, text="音频文件夹").grid(row=2, column=0, sticky='e', padx=5, pady=2)
        audio_folder_frame = ttk.Frame(content_frame)
        audio_folder_frame.grid(row=2, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        audio_folder_frame.columnconfigure(1, weight=1)
        ttk.Radiobutton(audio_folder_frame, text="", variable=audio_type_var, value="folder").grid(row=0, column=0, sticky='w')
        audio_folder_entry = ttk.Entry(audio_folder_frame, font=self.font_11)
        audio_folder_entry.grid(row=0, column=1, sticky='ew', padx=5)
        def select_folder(entry_widget):
            foldername = filedialog.askdirectory(title="选择文件夹", initialdir=application_path, parent=dialog)
            if foldername: entry_widget.delete(0, END); entry_widget.insert(0, foldername)
        ttk.Button(audio_folder_frame, text="选取...", command=lambda: select_folder(audio_folder_entry), bootstyle="outline").grid(row=0, column=2, padx=5)
        
        play_order_frame = ttk.Frame(content_frame)
        play_order_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        play_order_var = tk.StringVar(value="sequential")
        ttk.Radiobutton(play_order_frame, text="顺序播", variable=play_order_var, value="sequential").pack(side=LEFT, padx=10)
        ttk.Radiobutton(play_order_frame, text="随机播", variable=play_order_var, value="random").pack(side=LEFT, padx=10)

        bg_image_var = tk.IntVar(value=0)
        bg_image_path_var = tk.StringVar()
        bg_image_order_var = tk.StringVar(value="sequential")

        bg_image_frame = ttk.Frame(content_frame)
        bg_image_frame.grid(row=4, column=0, columnspan=4, sticky='w', padx=5, pady=5)
        bg_image_frame.columnconfigure(1, weight=1)
        bg_image_cb = ttk.Checkbutton(bg_image_frame, text="背景图片:", variable=bg_image_var, bootstyle="round-toggle")
        bg_image_cb.grid(row=0, column=0)
        if not IMAGE_AVAILABLE: bg_image_cb.config(state=DISABLED, text="背景图片(Pillow未安装):")

        bg_image_entry = ttk.Entry(bg_image_frame, textvariable=bg_image_path_var, font=self.font_11)
        bg_image_entry.grid(row=0, column=1, sticky='ew', padx=(5,5))

        bg_image_btn_frame = ttk.Frame(bg_image_frame)
        bg_image_btn_frame.grid(row=0, column=2)
        ttk.Button(bg_image_btn_frame, text="选取...", command=lambda: select_folder(bg_image_entry), bootstyle="outline").pack(side=LEFT)
        ttk.Radiobutton(bg_image_btn_frame, text="顺序", variable=bg_image_order_var, value="sequential").pack(side=LEFT, padx=(10,0))
        ttk.Radiobutton(bg_image_btn_frame, text="随机", variable=bg_image_order_var, value="random").pack(side=LEFT)

        volume_frame = ttk.Frame(content_frame)
        volume_frame.grid(row=5, column=1, columnspan=3, sticky='w', padx=5, pady=3)
        ttk.Label(volume_frame, text="音量:").pack(side=LEFT)
        volume_entry = ttk.Entry(volume_frame, font=self.font_11, width=10)
        volume_entry.pack(side=LEFT, padx=5)
        ttk.Label(volume_frame, text="0-100").pack(side=LEFT, padx=5)

        time_frame = ttk.LabelFrame(main_frame, text="时间", padding=15)
        time_frame.grid(row=1, column=0, sticky='ew', pady=4)
        time_frame.columnconfigure(1, weight=1)
        
        ttk.Label(time_frame, text="开始时间:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        start_time_entry = ttk.Entry(time_frame, font=self.font_11)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        self._bind_mousewheel_to_entry(start_time_entry, self._handle_time_scroll)
        ttk.Label(time_frame, text="《可多个,用英文逗号,隔开》").grid(row=0, column=2, sticky='w', padx=5)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bootstyle="outline").grid(row=0, column=3, padx=5)
        
        interval_var = tk.StringVar(value="first")
        ttk.Label(time_frame, text="间隔播报:").grid(row=1, column=0, sticky='e', padx=5, pady=2)
        interval_frame1 = ttk.Frame(time_frame)
        interval_frame1.grid(row=1, column=1, columnspan=2, sticky='w', padx=5, pady=2)
        ttk.Radiobutton(interval_frame1, text="播 n 首", variable=interval_var, value="first").pack(side=LEFT)
        interval_first_entry = ttk.Entry(interval_frame1, font=self.font_11, width=15)
        interval_first_entry.pack(side=LEFT, padx=5)
        ttk.Label(interval_frame1, text="(单曲时,指 n 遍)").pack(side=LEFT, padx=5)
        
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
            bg_image_var.set(task.get('bg_image_enabled', 0))
            bg_image_path_var.set(task.get('bg_image_path', ''))
            bg_image_order_var.set(task.get('bg_image_order', 'sequential'))
        else:
            volume_entry.insert(0, "80"); interval_first_entry.insert(0, "1"); interval_seconds_entry.insert(0, "600")
            weekday_entry.insert(0, "每周:1234567"); date_range_entry.insert(0, "2000-01-01 ~ 2099-12-31")

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

            new_task_data = {
                'name': name_entry.get().strip(), 'time': time_msg, 'content': audio_path, 'type': 'audio',
                'audio_type': audio_type_var.get(), 'play_order': play_order_var.get(),
                'volume': volume_entry.get().strip() or "80", 'interval_type': interval_var.get(),
                'interval_first': interval_first_entry.get().strip(), 'interval_seconds': interval_seconds_entry.get().strip(),
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

            self.update_task_list(); self.save_tasks(); cleanup_and_destroy() # <--- 【BUG修复】

            if play_this_task_now:
                self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))

        button_text = "保存修改" if is_edit_mode else "添加"
        ttk.Button(dialog_button_frame, text=button_text, command=save_task, bootstyle="primary").pack(side=LEFT, padx=10, ipady=5)
        ttk.Button(dialog_button_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10, ipady=5) # <--- 【BUG修复】
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】

    def open_video_dialog(self, parent_dialog, task_to_edit=None, index=None):
        parent_dialog.destroy()
        is_edit_mode = task_to_edit is not None
        dialog = ttk.Toplevel(self.root)
        self.active_modal_dialog = dialog # <--- 【BUG修复】
        dialog.title("修改视频节目" if is_edit_mode else "添加视频节目")
        dialog.resizable(True, True)
        dialog.minsize(800, 580)
        dialog.transient(self.root)
        dialog.grab_set()

        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            dialog.destroy()

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

        # --- 填充 content_frame ---
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

        ttk.Label(content_frame, text="视频文件夹:").grid(row=2, column=0, sticky='e', padx=5, pady=2)
        video_folder_frame = ttk.Frame(content_frame)
        video_folder_frame.grid(row=2, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
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

        # --- [修改 1: 调整音量布局] ---
        play_order_frame = ttk.Frame(content_frame)
        play_order_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        play_order_var = tk.StringVar(value="sequential")
        ttk.Radiobutton(play_order_frame, text="顺序播", variable=play_order_var, value="sequential").pack(side=LEFT, padx=10)
        ttk.Radiobutton(play_order_frame, text="随机播", variable=play_order_var, value="random").pack(side=LEFT, padx=10)
        
        # 将音量控件放在同一行
        ttk.Label(play_order_frame, text="音量:").pack(side=LEFT, padx=(20, 2))
        volume_entry = ttk.Entry(play_order_frame, font=self.font_11, width=5)
        volume_entry.pack(side=LEFT)
        ttk.Label(play_order_frame, text="(0-100)").pack(side=LEFT, padx=2)

        # --- 填充 playback_frame ---
        playback_mode_var = tk.StringVar(value="fullscreen")
        resolutions = ["640x480", "800x600", "1024x768", "1280x720", "1366x768", "1600x900", "1920x1080"]
        resolution_var = tk.StringVar(value=resolutions[2])

        playback_rates = ['0.5x', '0.75x', '1.0x (正常)', '1.25x', '1.5x', '2.0x']
        playback_rate_var = tk.StringVar(value='1.0x (正常)')

        # --- [修改 2: 调整播放倍速布局] ---
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

        # 将播放倍速放在同一行
        ttk.Label(mode_frame, text="倍速:").pack(side=LEFT)
        rate_combo = ttk.Combobox(mode_frame, textvariable=playback_rate_var, values=playback_rates, font=self.font_11, width=10)
        rate_combo.pack(side=LEFT, padx=2)
        ttk.Label(mode_frame, text="(0.25-4.0)", font=self.font_9, bootstyle="secondary").pack(side=LEFT, padx=2)

        toggle_resolution_combo()

        # --- 后续布局保持不变 ---
        time_frame = ttk.LabelFrame(main_frame, text="时间", padding=15)
        time_frame.grid(row=2, column=0, sticky='ew', pady=4)
        time_frame.columnconfigure(1, weight=1)

        ttk.Label(time_frame, text="开始时间:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        start_time_entry = ttk.Entry(time_frame, font=self.font_11)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        self._bind_mousewheel_to_entry(start_time_entry, self._handle_time_scroll)
        ttk.Label(time_frame, text="《可多个,用英文逗号,隔开》").grid(row=0, column=2, sticky='w', padx=5)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bootstyle="outline").grid(row=0, column=3, padx=5)

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
            playback_mode_var.set(task.get('playback_mode', 'fullscreen'))
            resolution_var.set(task.get('resolution', '1024x768'))
            playback_rate_var.set(task.get('playback_rate', '1.0x (正常)'))
            start_time_entry.insert(0, task.get('time', ''))
            interval_var.set(task.get('interval_type', 'first'))
            interval_first_entry.insert(0, task.get('interval_first', '1'))
            interval_seconds_entry.insert(0, task.get('interval_seconds', '600'))
            weekday_entry.insert(0, task.get('weekday', '每周:1234567'))
            date_range_entry.insert(0, task.get('date_range', '2000-01-01 ~ 2099-12-31'))
            delay_var.set(task.get('delay', 'ontime'))
            toggle_resolution_combo()
        else:
            volume_entry.insert(0, "80")
            interval_first_entry.insert(0, "1")
            interval_seconds_entry.insert(0, "600")
            weekday_entry.insert(0, "每周:1234567")
            date_range_entry.insert(0, "2000-01-01 ~ 2099-12-31")

        def save_task():
            video_path = video_single_entry.get().strip() if video_type_var.get() == "single" else video_folder_entry.get().strip()
            if not video_path:
                messagebox.showwarning("警告", "请选择一个视频文件或文件夹", parent=dialog)
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

            new_task_data = {
                'name': name_entry.get().strip() or os.path.basename(video_path),
                'time': time_msg,
                'content': video_path,
                'type': 'video',
                'video_type': video_type_var.get(),
                'play_order': play_order_var.get(),
                'volume': volume_entry.get().strip() or "80",
                'interval_type': interval_var.get(),
                'interval_first': interval_first_entry.get().strip() or "1",
                'interval_seconds': interval_seconds_entry.get().strip() or "600",
                'playback_mode': playback_mode_var.get(),
                'resolution': resolution_var.get(),
                'playback_rate': rate_input,
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
            cleanup_and_destroy() # <--- 【BUG修复】

            if play_this_task_now:
                self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))

        button_text = "保存修改" if is_edit_mode else "添加"
        ttk.Button(dialog_button_frame, text=button_text, command=save_task, bootstyle="primary").pack(side=LEFT, padx=10, ipady=5)
        ttk.Button(dialog_button_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10, ipady=5) # <--- 【BUG修复】
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】

#第6部分
#第6部分
    def open_voice_dialog(self, parent_dialog, task_to_edit=None, index=None):
        parent_dialog.destroy()
        is_edit_mode = task_to_edit is not None
        dialog = ttk.Toplevel(self.root)
        self.active_modal_dialog = dialog # <--- 【BUG修复】
        dialog.title("修改语音节目" if is_edit_mode else "添加语音节目")
        dialog.resizable(True, True)
        dialog.minsize(800, 580)
        dialog.transient(self.root); dialog.grab_set()

        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            dialog.destroy()

        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)

        content_frame = ttk.LabelFrame(main_frame, text="内容", padding=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=2)
        content_frame.columnconfigure(1, weight=1)

        ttk.Label(content_frame, text="节目名称:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        name_entry = ttk.Entry(content_frame, font=self.font_11)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        
        # --- [修改 1: 减少播音文字框高度] ---
        ttk.Label(content_frame, text="播音文字:").grid(row=1, column=0, sticky='nw', padx=5, pady=2)
        text_frame = ttk.Frame(content_frame)
        text_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        content_text = ScrolledText(text_frame, height=3, font=self.font_11, wrap=WORD) # <-- 高度从5改为3
        content_text.grid(row=0, column=0, sticky='nsew')
        
        script_btn_frame = ttk.Frame(content_frame)
        script_btn_frame.grid(row=2, column=1, columnspan=3, sticky='w', padx=5, pady=(0, 2))
        ttk.Button(script_btn_frame, text="导入文稿", command=lambda: self._import_voice_script(content_text), bootstyle="outline").pack(side=LEFT)
        ttk.Button(script_btn_frame, text="导出文稿", command=lambda: self._export_voice_script(content_text, name_entry), bootstyle="outline").pack(side=LEFT, padx=10)

        # --- ↓↓↓ 新增广告制作按钮 ↓↓↓ ---
        # 创建一个专门的框架来容纳广告按钮
        ad_btn_frame = ttk.Frame(script_btn_frame)
        ad_btn_frame.pack(side=LEFT, padx=20)

        # 按语音长度制作
        self.ad_by_voice_btn = ttk.Button(ad_btn_frame, text="按语音长度制作广告", 
                                          command=lambda: self._create_advertisement('voice'))
        self.ad_by_voice_btn.pack(side=LEFT)

        # 按背景音乐长度制作
        self.ad_by_bgm_btn = ttk.Button(ad_btn_frame, text="按背景音乐长度制作广告", 
                                        command=lambda: self._create_advertisement('bgm'))
        self.ad_by_bgm_btn.pack(side=LEFT, padx=10)

        # 权限控制
        if self.auth_info['status'] != 'Permanent':
            self.ad_by_voice_btn.config(state=DISABLED)
            self.ad_by_bgm_btn.config(state=DISABLED)
            # 可以在按钮旁边加一个提示
            #ttk.Label(ad_btn_frame, text="(永久授权可用)", font=self.font_9, bootstyle="secondary").pack(side=LEFT)
        
        # --- ↑↑↑ 新增代码结束 ↑↑↑ ---        

        # --- [修改 2: 调整语速/音调/音量布局] ---
        ttk.Label(content_frame, text="播音员:").grid(row=3, column=0, sticky='w', padx=5, pady=3)
        voice_frame = ttk.Frame(content_frame)
        voice_frame.grid(row=3, column=1, columnspan=3, sticky='ew', padx=5, pady=3)
        voice_frame.columnconfigure(0, weight=1) # 让播音员列表可以伸展
        
        available_voices = self.get_available_voices()
        voice_var = tk.StringVar()
        voice_combo = ttk.Combobox(voice_frame, textvariable=voice_var, values=available_voices, font=self.font_11, state='readonly')
        voice_combo.grid(row=0, column=0, sticky='ew')
        
        # 创建一个新的框架来容纳右侧的参数输入框
        speech_params_frame = ttk.Frame(voice_frame)
        speech_params_frame.grid(row=0, column=1, sticky='e', padx=(10, 0))

        ttk.Label(speech_params_frame, text="语速:").pack(side=LEFT)
        speed_entry = ttk.Entry(speech_params_frame, font=self.font_11, width=5); speed_entry.pack(side=LEFT, padx=(2, 5))
        ttk.Label(speech_params_frame, text="音调:").pack(side=LEFT)
        pitch_entry = ttk.Entry(speech_params_frame, font=self.font_11, width=5); pitch_entry.pack(side=LEFT, padx=(2, 5))
        ttk.Label(speech_params_frame, text="音量:").pack(side=LEFT)
        volume_entry = ttk.Entry(speech_params_frame, font=self.font_11, width=5); volume_entry.pack(side=LEFT, padx=(2, 0))

        # --- 后续布局保持不变，只是行号可能需要调整 ---
        prompt_var = tk.IntVar(); prompt_frame = ttk.Frame(content_frame)
        # 原来的 speech_params_frame 在第4行，现在被合并了，所以后续控件从第5行开始
        prompt_frame.grid(row=5, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        prompt_frame.columnconfigure(1, weight=1)
        ttk.Checkbutton(prompt_frame, text="提示音:", variable=prompt_var, bootstyle="round-toggle").grid(row=0, column=0, sticky='w')
        prompt_file_var, prompt_volume_var = tk.StringVar(), tk.StringVar()
        prompt_file_entry = ttk.Entry(prompt_frame, textvariable=prompt_file_var, font=self.font_11); prompt_file_entry.grid(row=0, column=1, sticky='ew', padx=5)
        ttk.Button(prompt_frame, text="...", command=lambda: self.select_file_for_entry(PROMPT_FOLDER, prompt_file_var), bootstyle="outline", width=2).grid(row=0, column=2)
        
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
        ttk.Button(bgm_frame, text="...", command=lambda: self.select_file_for_entry(BGM_FOLDER, bgm_file_var), bootstyle="outline", width=2).grid(row=0, column=2)
        
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
        ttk.Label(time_frame, text="《可多个,用英文逗号,隔开》").grid(row=0, column=2, sticky='w', padx=5)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bootstyle="outline").grid(row=0, column=3, padx=5)
        
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
            bg_image_var.set(task.get('bg_image_enabled', 0))
            bg_image_path_var.set(task.get('bg_image_path', ''))
            bg_image_order_var.set(task.get('bg_image_order', 'sequential'))
        else:
            speed_entry.insert(0, "0"); pitch_entry.insert(0, "0"); volume_entry.insert(0, "80")
            prompt_var.set(0); prompt_volume_var.set("80"); bgm_var.set(0); bgm_volume_var.set("40")
            repeat_entry.insert(0, "1"); weekday_entry.insert(0, "每周:1234567"); date_range_entry.insert(0, "2000-01-01 ~ 2099-12-31")

        # 将所有需要用到的控件变量收集到 ad_params 字典中
        ad_params = {
            'dialog': dialog,
            'name_entry': name_entry,
            'content_text': content_text,
            'voice_var': voice_var,
            'speed_entry': speed_entry,
            'pitch_entry': pitch_entry,
            'volume_entry': volume_entry,
            'prompt_var': prompt_var,
            'prompt_file_var': prompt_file_var,
            'prompt_volume_var': prompt_volume_var,
            'bgm_var': bgm_var,
            'bgm_file_var': bgm_file_var,
            'bgm_volume_var': bgm_volume_var,
        }

        # 现在为按钮配置正确的 command，并传入 ad_params 字典
        self.ad_by_voice_btn.config(command=lambda: self._create_advertisement('voice', ad_params))
        self.ad_by_bgm_btn.config(command=lambda: self._create_advertisement('bgm', ad_params))

        def save_task():
            # ... (在所有代码的最前面)

            # --- ↓↓↓ 新增的验证逻辑 ↓↓↓ ---
            try:
                speed = int(speed_entry.get().strip() or '0')
                pitch = int(pitch_entry.get().strip() or '0')
                volume = int(volume_entry.get().strip() or '80')

                if not (-10 <= speed <= 10):
                    messagebox.showerror("输入错误", "语速必须在 -10 到 10 之间。", parent=dialog)
                    return # 中断保存
                if not (-10 <= pitch <= 10):
                    messagebox.showerror("输入错误", "音调必须在 -10 到 10 之间。", parent=dialog)
                    return # 中断保存
                if not (0 <= volume <= 100):
                    messagebox.showerror("输入错误", "音量必须在 0 到 100 之间。", parent=dialog)
                    return # 中断保存
            except ValueError:
                messagebox.showerror("输入错误", "语速、音调、音量必须是有效的整数。", parent=dialog)
                return # 中断保存
            # --- ↑↑↑ 验证逻辑结束 ↑↑↑ ---
            text_content = content_text.get('1.0', END).strip()
            if not text_content: messagebox.showwarning("警告", "请输入播音文字内容", parent=dialog); return
            is_valid_time, time_msg = self._normalize_multiple_times_string(start_time_entry.get().strip())
            if not is_valid_time: messagebox.showwarning("格式错误", time_msg, parent=dialog); return
            is_valid_date, date_msg = self._normalize_date_range_string(date_range_entry.get().strip())
            if not is_valid_date: messagebox.showwarning("格式错误", date_msg, parent=dialog); return
            regeneration_needed = True
            if is_edit_mode:
                original_task = task_to_edit
                if (text_content == original_task.get('source_text') and voice_var.get() == original_task.get('voice') and
                    speed_entry.get().strip() == original_task.get('speed', '0') and pitch_entry.get().strip() == original_task.get('pitch', '0') and
                    volume_entry.get().strip() == original_task.get('volume', '80')):
                    regeneration_needed = False; self.log("语音内容未变更，跳过重新生成WAV文件。")

            def build_task_data(wav_path, wav_filename_str):
                play_mode = delay_var.get()
                play_this_task_now = (play_mode == 'immediate')
                saved_delay_type = 'delay' if play_mode == 'immediate' else play_mode

                return {
                    'name': name_entry.get().strip(), 'time': time_msg, 'type': 'voice', 'content': wav_path,
                    'wav_filename': wav_filename_str, 'source_text': text_content, 'voice': voice_var.get(),
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
                if not new_task_data['name'] or not new_task_data['time']: messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog); return
                self.tasks[index] = new_task_data; self.log(f"已修改语音节目(未重新生成语音): {new_task_data['name']}")
                self.update_task_list(); self.save_tasks(); cleanup_and_destroy() # <--- 【BUG修复】
                if play_now_flag: self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))
                return

            progress_dialog = ttk.Toplevel(dialog)
            self.active_modal_dialog = progress_dialog # <--- 【BUG修复】
            progress_dialog.title("请稍候")
            progress_dialog.resizable(False, False); progress_dialog.transient(dialog); progress_dialog.grab_set()
            
            def cleanup_progress(): # <--- 【BUG修复】
                self.active_modal_dialog = dialog # Restore focus to the main dialog
                progress_dialog.destroy()

            progress_dialog.protocol("WM_DELETE_WINDOW", cleanup_progress) # <--- 【BUG修复】

            ttk.Label(progress_dialog, text="语音文件生成中，请稍后...", font=self.font_11).pack(expand=True, padx=20, pady=20)
            self.center_window(progress_dialog, parent=dialog)
            
            new_wav_filename = f"{int(time.time())}_{random.randint(1000, 9999)}.wav"
            output_path = os.path.join(AUDIO_FOLDER, new_wav_filename)
            voice_params = {'voice': voice_var.get(), 'speed': speed_entry.get().strip() or "0", 'pitch': pitch_entry.get().strip() or "0", 'volume': volume_entry.get().strip() or "80"}
            def _on_synthesis_complete(result):
                cleanup_progress() # <--- 【BUG修复】
                if not result['success']: messagebox.showerror("错误", f"无法生成语音文件: {result['error']}", parent=dialog); return
                if is_edit_mode and 'wav_filename' in task_to_edit:
                    old_wav_path = os.path.join(AUDIO_FOLDER, task_to_edit['wav_filename'])
                    if os.path.exists(old_wav_path):
                        try: os.remove(old_wav_path); self.log(f"已删除旧语音文件: {task_to_edit['wav_filename']}")
                        except Exception as e: self.log(f"删除旧语音文件失败: {e}")
                new_task_data, play_now_flag = build_task_data(output_path, new_wav_filename)
                if not new_task_data['name'] or not new_task_data['time']: messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog); return
                if is_edit_mode: self.tasks[index] = new_task_data; self.log(f"已修改语音节目(并重新生成语音): {new_task_data['name']}")
                else: self.tasks.append(new_task_data); self.log(f"已添加语音节目: {new_task_data['name']}")
                self.update_task_list(); self.save_tasks(); cleanup_and_destroy() # <--- 【BUG修复】
                if play_now_flag: self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))
            synthesis_thread = threading.Thread(target=self._synthesis_worker, args=(text_content, voice_params, output_path, _on_synthesis_complete))
            synthesis_thread.daemon = True; synthesis_thread.start()

        button_text = "保存修改" if is_edit_mode else "添加"
        ttk.Button(dialog_button_frame, text=button_text, command=save_task, bootstyle="primary").pack(side=LEFT, padx=10, ipady=5)
        ttk.Button(dialog_button_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10, ipady=5) # <--- 【BUG修复】
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】

    def _create_advertisement(self, mode, params):
        """
        核心广告制作函数
        mode: 'voice' 或 'bgm'
        params: 包含所有UI控件变量的字典
        """
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

        # 1. 数据验证
        if not params['bgm_var'].get() or not params['bgm_file_var'].get().strip():
            messagebox.showerror("错误", "必须选择背景音乐才能制作广告。", parent=params['dialog'])
            return

        bgm_path = params['bgm_file_var'].get().strip()
        if not os.path.exists(bgm_path):
            messagebox.showerror("错误", f"背景音乐文件不存在：\n{bgm_path}", parent=params['dialog'])
            return

        text_content = params['content_text'].get('1.0', 'end').strip()
        if not text_content:
            messagebox.showerror("错误", "播音文字内容不能为空。", parent=params['dialog'])
            return
            
        try:
            voice_volume = int(params['volume_entry'].get().strip() or '80')
            bgm_volume = int(params['bgm_volume_var'].get().strip() or '40')
        except ValueError:
            messagebox.showerror("错误", "音量必须是有效的整数。", parent=params['dialog'])
            return

        # 2. 显示进度窗口
        progress_dialog = ttk.Toplevel(params['dialog'])
        self.active_modal_dialog = progress_dialog # <--- 【BUG修复】
        progress_dialog.title("正在制作广告")
        progress_dialog.resizable(False, False)
        progress_dialog.transient(params['dialog']); progress_dialog.grab_set()
        
        def cleanup_progress(): # <--- 【BUG修复】
            self.active_modal_dialog = params['dialog']
            progress_dialog.destroy()

        progress_dialog.protocol("WM_DELETE_WINDOW", cleanup_progress) # <--- 【BUG修复】

        progress_label = ttk.Label(progress_dialog, text="正在准备...", font=self.font_11)
        progress_label.pack(pady=10, padx=20)
        progress = ttk.Progressbar(progress_dialog, length=300, mode='determinate')
        progress.pack(pady=10, padx=20)
        self.center_window(progress_dialog, parent=params['dialog'])

        # 3. 在后台线程中执行耗时操作
        def worker():
            temp_wav_path = None
            try:
                # --- 步骤 A: 生成或加载语音文件 ---
                self.root.after(0, lambda: progress_label.config(text="步骤1/4: 生成语音..."))
                self.root.after(0, lambda: progress.config(value=10))

                temp_wav_filename = f"temp_ad_{int(time.time())}.wav"
                temp_wav_path = os.path.join(AUDIO_FOLDER, temp_wav_filename)
                
                voice_params = {
                    'voice': params['voice_var'].get(),
                    'speed': params['speed_entry'].get().strip() or "0",
                    'pitch': params['pitch_entry'].get().strip() or "0",
                    'volume': '100'
                }
                
                success = self._synthesize_text_to_wav(text_content, voice_params, temp_wav_path)
                if not success:
                    raise Exception("语音合成失败！")

                # --- 步骤 B: 加载音频并获取时长 ---
                self.root.after(0, lambda: progress_label.config(text="步骤2/4: 分析音频..."))
                self.root.after(0, lambda: progress.config(value=30))

                voice_audio = AudioSegment.from_wav(temp_wav_path)
                bgm_audio = AudioSegment.from_file(bgm_path)

                voice_duration_ms = len(voice_audio)
                bgm_duration_ms = len(bgm_audio)

                if voice_duration_ms == 0:
                    raise ValueError("合成的语音长度为0，无法制作广告。")

                # --- ↓↓↓ 核心算法修改区域 (版本3) ↓↓↓ ---
                self.root.after(0, lambda: progress_label.config(text="步骤3/4: 计算并混合音频..."))
                self.root.after(0, lambda: progress.config(value=60))

                # 定义音量转换函数
                def volume_to_db(vol_percent):
                    if vol_percent <= 0: return -120
                    return 20 * (vol_percent / 100.0) - 20

                # 先调整好各自的音量
                adjusted_voice = voice_audio + volume_to_db(voice_volume)
                adjusted_bgm = bgm_audio + volume_to_db(bgm_volume)

                final_output = None

                if mode == 'voice':
                    # 按语音长度模式：只播报一次
                    if bgm_duration_ms < voice_duration_ms:
                        raise ValueError("背景音乐长度小于语音长度，无法制作。")
                    
                    # 截取背景音乐，然后叠加
                    final_bgm_segment = adjusted_bgm[:voice_duration_ms]
                    final_output = final_bgm_segment.overlay(adjusted_voice)

                elif mode == 'bgm':
                    # 按背景音乐长度模式：在BGM总时长内重复播报
                    silence_5_sec = AudioSegment.silent(duration=5000)
                    
                    # 定义一个播报单元 = 语音 + 尾部静音
                    unit_audio = adjusted_voice + silence_5_sec
                    unit_duration_ms = len(unit_audio)

                    if bgm_duration_ms < voice_duration_ms:
                         raise ValueError(f"背景音乐太短（{bgm_duration_ms/1000.0:.1f}秒），无法容纳一次完整的语音（需要 {voice_duration_ms/1000.0:.1f} 秒）。")

                    # 计算可以完整播报多少次
                    repeat_count = int(bgm_duration_ms // unit_duration_ms)
                    
                    # 如果连一次完整的“语音+静音”都放不下，就只放一次语音
                    if repeat_count == 0:
                        repeat_count = 1
                        unit_audio = adjusted_voice # 此时单元不带静音
                    
                    # 创建一个与背景音乐等长的空白“画布”
                    voice_canvas = AudioSegment.silent(duration=bgm_duration_ms)
                    
                    # 在画布上依次叠加播报单元
                    current_pos_ms = 0
                    for i in range(repeat_count):
                        # 确保下一次叠加不会超出画布范围
                        if current_pos_ms + len(unit_audio) <= bgm_duration_ms:
                            voice_canvas = voice_canvas.overlay(unit_audio, position=current_pos_ms)
                            current_pos_ms += len(unit_audio)
                        else:
                            # 如果加上静音会超出，就尝试只加语音
                            if current_pos_ms + len(adjusted_voice) <= bgm_duration_ms:
                                voice_canvas = voice_canvas.overlay(adjusted_voice, position=current_pos_ms)
                            break # 空间不足，停止添加

                    # 将填充好语音的画布与原始背景音乐混合
                    final_output = adjusted_bgm.overlay(voice_canvas)

                # --- ↑↑↑ 核心算法修改区域结束 ↑↑↑ ---

                # --- 步骤 E: 导出为MP3 ---
                self.root.after(0, lambda: progress_label.config(text="步骤4/4: 导出MP3文件..."))
                self.root.after(0, lambda: progress.config(value=90))
                
                ad_folder = os.path.join(application_path, "导出的广告")
                if not os.path.exists(ad_folder):
                    os.makedirs(ad_folder)
                
                safe_filename = re.sub(r'[\\/*?:"<>|]', "", params['name_entry'].get().strip() or '未命名广告')
                output_filename = f"{safe_filename}_{int(time.time())}.mp3"
                output_path = os.path.join(ad_folder, output_filename)

                final_output.export(
                    output_path,
                    format="mp3",
                    bitrate="256k",
                    parameters=["-ar", "44100", "-id3v2_version", "3"],
                    codec="libmp3lame"
                )

                self.root.after(0, lambda: progress.config(value=100))
                self.root.after(100, lambda: messagebox.showinfo("成功", f"广告制作成功！\n\n已保存至：\n{output_path}", parent=params['dialog']))

            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("制作失败", f"发生错误：\n{e}", parent=params['dialog']))
            
            finally:
                if temp_wav_path and os.path.exists(temp_wav_path):
                    try:
                        os.remove(temp_wav_path)
                    except Exception as e_del:
                        self.log(f"删除临时文件 {temp_wav_path} 失败: {e_del}")
                self.root.after(0, cleanup_progress) # <--- 【BUG修复】

        threading.Thread(target=worker, daemon=True).start()
        
#第7部分
#第7部分
    def _import_voice_script(self, text_widget):
        filename = filedialog.askopenfilename(
            title="选择要导入的文稿",
            initialdir=VOICE_SCRIPT_FOLDER,
            filetypes=[("文本文档", "*.txt"), ("所有文件", "*.*")],
            parent=self.root
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
            messagebox.showerror("导入失败", f"无法读取文件：\n{e}", parent=self.root)
            self.log(f"导入文稿失败: {e}")

    def _export_voice_script(self, text_widget, name_widget):
        content = text_widget.get('1.0', END).strip()
        if not content:
            messagebox.showwarning("无法导出", "播音文字内容为空，无需导出。", parent=self.root)
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
            parent=self.root
        )
        if not filename:
            return

        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(content)
            self.log(f"文稿已成功导出到 {os.path.basename(filename)}。")
            messagebox.showinfo("导出成功", f"文稿已成功导出到：\n{filename}", parent=self.root)
        except Exception as e:
            messagebox.showerror("导出失败", f"无法保存文件：\n{e}", parent=self.root)
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

    def select_file_for_entry(self, initial_dir, string_var):
        filename = filedialog.askopenfilename(title="选择文件", initialdir=initial_dir, filetypes=[("音频文件", "*.mp3 *.wav *.ogg *.flac *.m4a"), ("所有文件", "*.*")], parent=self.root)
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
        if not selection: messagebox.showwarning("警告", "请先选择要修改的节目", parent=self.root); return
        if len(selection) > 1: messagebox.showwarning("警告", "一次只能修改一个节目", parent=self.root); return
        index = self.task_tree.index(selection[0])
        task = self.tasks[index]
        dummy_parent = ttk.Toplevel(self.root)
        self.active_modal_dialog = dummy_parent # <--- 【BUG修复】
        dummy_parent.withdraw()

        task_type = task.get('type')
        if task_type == 'audio':
            self.open_audio_dialog(dummy_parent, task_to_edit=task, index=index)
        elif task_type == 'voice':
            self.open_voice_dialog(dummy_parent, task_to_edit=task, index=index)
        elif task_type == 'video':
            self.open_video_dialog(dummy_parent, task_to_edit=task, index=index)
        else:
             self.open_audio_dialog(dummy_parent, task_to_edit=task, index=index)

        def check_dialog_closed():
            try:
                if not dummy_parent.winfo_children(): 
                    self.active_modal_dialog = None # <--- 【BUG修复】
                    dummy_parent.destroy()
                else: self.root.after(100, check_dialog_closed)
            except tk.TclError: 
                self.active_modal_dialog = None # <--- 【BUG修复】
        self.root.after(100, check_dialog_closed)

    def copy_task(self):
        selections = self.task_tree.selection()
        if not selections: messagebox.showwarning("警告", "请先选择要复制的节目", parent=self.root); return
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
        filename = filedialog.askopenfilename(title="选择导入文件", filetypes=[("JSON文件", "*.json")], initialdir=application_path, parent=self.root)
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f: imported = json.load(f)

                if not isinstance(imported, list) or \
                   (imported and (not isinstance(imported[0], dict) or 'time' not in imported[0] or 'type' not in imported[0])):
                    messagebox.showerror("导入失败", "文件格式不正确，看起来不是一个有效的节目单备份文件。", parent=self.root)
                    self.log(f"尝试导入格式错误的节目单文件: {os.path.basename(filename)}")
                    return

                self.tasks.extend(imported); self.update_task_list(); self.save_tasks()
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
        self.active_modal_dialog = dialog # <--- 【BUG修复】
        dialog.title(title)
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        result = [None]

        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            dialog.destroy()

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
                cleanup_and_destroy() # <--- 【BUG修复】
            except ValueError:
                messagebox.showerror("输入错误", "请输入一个有效的整数。", parent=dialog)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=15)

        ttk.Button(btn_frame, text="确定", command=on_confirm, bootstyle="primary", width=8).pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=cleanup_and_destroy, width=8).pack(side=LEFT, padx=10) # <--- 【BUG修复】

        dialog.bind('<Return>', lambda event: on_confirm())
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】

        self.center_window(dialog, parent=self.root)
        self.root.wait_window(dialog)
        return result[0]

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
        self.active_modal_dialog = dialog # <--- 【BUG修复】
        dialog.title("开始时间设置")
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()

        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            dialog.destroy()

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
            cleanup_and_destroy() # <--- 【BUG修复】
        ttk.Button(bottom_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=5, ipady=5)
        ttk.Button(bottom_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=5, ipady=5) # <--- 【BUG修复】
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】
        
        self.center_window(dialog, parent=self.root)

    def show_weekday_settings_dialog(self, weekday_entry):
        dialog = ttk.Toplevel(self.root)
        self.active_modal_dialog = dialog # <--- 【BUG修复】
        dialog.title("周几或几号")
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()

        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            dialog.destroy()

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
            cleanup_and_destroy() # <--- 【BUG修复】
        ttk.Button(bottom_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=5, ipady=5)
        ttk.Button(bottom_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=5, ipady=5) # <--- 【BUG修复】
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】

        self.center_window(dialog, parent=self.root)

#第9部分
#第9部分
    def show_daterange_settings_dialog(self, date_range_entry):
        dialog = ttk.Toplevel(self.root)
        self.active_modal_dialog = dialog # <--- 【BUG修复】
        dialog.title("日期范围")
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()

        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            dialog.destroy()

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
        except (ValueError, IndexError): from_date_entry.insert(0, "2000-01-01"); to_date_entry.insert(0, "2099-12-31")
        ttk.Label(main_frame, text="格式: YYYY-MM-DD", font=self.font_11, bootstyle="secondary").pack(pady=10)
        bottom_frame = ttk.Frame(main_frame); bottom_frame.pack(pady=10)
        def confirm():
            start, end = from_date_entry.get().strip(), to_date_entry.get().strip()
            norm_start, norm_end = self._normalize_date_string(start), self._normalize_date_string(end)
            if norm_start and norm_end:
                date_range_entry.delete(0, END)
                date_range_entry.insert(0, f"{norm_start} ~ {norm_end}")
                cleanup_and_destroy() # <--- 【BUG修复】
            else: messagebox.showerror("格式错误", "日期格式不正确, 应为 YYYY-MM-DD", parent=dialog)
        ttk.Button(bottom_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=5, ipady=5)
        ttk.Button(bottom_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=5, ipady=5) # <--- 【BUG修复】
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】

        self.center_window(dialog, parent=self.root)

    def show_single_time_dialog(self, time_var):
        dialog = ttk.Toplevel(self.root)
        self.active_modal_dialog = dialog # <--- 【BUG修复】
        dialog.title("设置时间")
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()

        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            dialog.destroy()

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
                cleanup_and_destroy() # <--- 【BUG修复】
            else: messagebox.showerror("格式错误", "请输入有效的时间格式 HH:MM:SS", parent=dialog)
        bottom_frame = ttk.Frame(main_frame); bottom_frame.pack(pady=10)
        ttk.Button(bottom_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=10)
        ttk.Button(bottom_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10) # <--- 【BUG修复】
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】
        
        self.center_window(dialog, parent=self.root)

    def show_power_week_time_dialog(self, title, days_var, time_var):
        dialog = ttk.Toplevel(self.root)
        self.active_modal_dialog = dialog # <--- 【BUG修复】
        dialog.title(title)
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()

        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            dialog.destroy()

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
            cleanup_and_destroy() # <--- 【BUG修复】
        bottom_frame = ttk.Frame(dialog); bottom_frame.pack(pady=15)
        ttk.Button(bottom_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=10)
        ttk.Button(bottom_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10) # <--- 【BUG修复】
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】

        self.center_window(dialog, parent=self.root)

    def update_task_list(self):
        if not hasattr(self, 'task_tree') or not self.task_tree.winfo_exists(): return
        selection = self.task_tree.selection()
        self.task_tree.delete(*self.task_tree.get_children())
        for task in self.tasks:
            content = task.get('content', '')
            task_type = task.get('type')

            if task_type == 'voice':
                source_text = task.get('source_text', '')
                clean_content = source_text.replace('\n', ' ').replace('\r', '')
                content_preview = (clean_content[:30] + '...') if len(clean_content) > 30 else clean_content
            elif task_type in ['audio', 'video']:
                content_preview = os.path.basename(content)
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
        self.root.after(1000, self._process_reminder_queue)

    def _check_running_processes_for_termination(self, now):
        # 遍历活动进程字典的副本，因为我们可能会在循环中删除元素
        for task_id in list(self.active_processes.keys()):
            proc_info = self.active_processes.get(task_id)
            if not proc_info: continue

            task = proc_info.get('task')
            process = proc_info.get('process')
            stop_time_str = task.get('stop_time')

            if not stop_time_str: continue  # 如果任务没有设置停止时间，则跳过

            # 检查进程是否还在运行，如果已经自己退出了，就清理掉
            try:
                if process.poll() is not None:
                    del self.active_processes[task_id]
                    continue
            except Exception:  # 捕获所有可能的异常，例如进程不存在
                del self.active_processes[task_id]
                continue

            # 核心判断：当前时间是否到达停止时间
            current_time_str = now.strftime("%H:%M:%S")
            if current_time_str >= stop_time_str:
                self.log(f"到达停止时间，正在终止任务 '{task['name']}' (PID: {process.pid})...")
                try:
                    # 使用 psutil 强制结束进程及其所有子进程
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
                    # 无论成功与否，都从监控列表中移除
                    if task_id in self.active_processes:
                        del self.active_processes[task_id]

    def _scheduler_worker(self):
        while self.running:
            now = datetime.now()
            if not self.is_app_locked_down:
                self._check_broadcast_tasks(now)
                self._check_advanced_tasks(now)
                self._check_time_chime(now)
                self._check_todo_tasks(now)
                self._check_running_processes_for_termination(now)

            self._check_power_tasks(now)
            time.sleep(1)

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
        if self._is_in_holiday(now): return

        for task in self.screenshot_tasks:
            is_due, trigger_time = self._is_task_due(task, now)
            if is_due:
                self.log(f"触发截屏任务: {task['name']}")
                threading.Thread(target=self._execute_screenshot_task, args=(task, trigger_time), daemon=True).start()
        
        for task in self.execute_tasks:
            is_due, trigger_time = self._is_task_due(task, now)
            if is_due:
                self.log(f"触发运行任务: {task['name']}")
                threading.Thread(target=self._execute_program_task, args=(task, trigger_time), daemon=True).start()
    
    # 找到 _execute_screenshot_task 函数并替换为以下内容：
    def _execute_screenshot_task(self, task, trigger_time):
        if not IMAGE_AVAILABLE:
            self.log(f"错误：Pillow库未安装，无法执行截屏任务 '{task['name']}'。")
            return
        
        try:
            repeat_count = task.get('repeat_count', 1)
            interval_seconds = task.get('interval_seconds', 0)
            stop_time_str = task.get('stop_time') # 获取停止时间

            for i in range(repeat_count):
                # --- 【核心修复】在这里增加停止时间的判断 ---
                if stop_time_str:
                    current_time_str = datetime.now().strftime('%H:%M:%S')
                    if current_time_str >= stop_time_str:
                        self.log(f"任务 '{task['name']}' 已到达停止时间 '{stop_time_str}'，提前中止截屏。")
                        break # 退出循环
                # --- 修复结束 ---

                screenshot = ImageGrab.grab()
                filename = f"Screenshot_{task['name']}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')[:-3]}.png"
                save_path = os.path.join(SCREENSHOT_FOLDER, filename)
                screenshot.save(save_path)
                self.log(f"任务 '{task['name']}' 已成功截屏 ({i+1}/{repeat_count})，保存至: {filename}")

                if i < repeat_count - 1:
                    time.sleep(interval_seconds)
            
            task.setdefault('last_run', {})[trigger_time] = datetime.now().strftime("%Y-%m-%d")
            self.save_screenshot_tasks()

        except Exception as e: # <--- 修正：已将 except 块正确地配对到 try 之后
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
                self.log(f"触发整点报时: {now.hour:02d}点")
                self.playback_command_queue.put(('PLAY_CHIME', chime_file))
            else:
                self.log(f"警告：找不到整点报时文件 {chime_file}，报时失败。")

    def _check_broadcast_tasks(self, now):
        if self._is_in_holiday(now):
            return

        tasks_to_play = []
        for task in self.tasks:
            is_due, trigger_time = self._is_task_due(task, now)
            if is_due:
                tasks_to_play.append((task, trigger_time))

        if not tasks_to_play:
            return

        ontime_tasks = [t for t in tasks_to_play if t[0].get('delay') == 'ontime']
        delay_tasks = [t for t in tasks_to_play if t[0].get('delay') != 'ontime']

        if ontime_tasks:
            task, trigger_time = ontime_tasks[0]
            self.log(f"准时任务 '{task['name']}' 已到时间，执行高优先级中断。")
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

            elif command == 'PLAY_CHIME':
                if not AUDIO_AVAILABLE: continue
                chime_path = data
                was_playing = pygame.mixer.music.get_busy()
                if was_playing:
                    pygame.mixer.music.pause()
                    self.log("整点报时，暂停当前播放...")

                try:
                    chime_sound = pygame.mixer.Sound(chime_path)
                    chime_sound.set_volume(1.0)
                    chime_channel = pygame.mixer.find_channel(True)
                    chime_channel.play(chime_sound)
                    while chime_channel and chime_channel.get_busy():
                        time.sleep(0.1)
                except Exception as e:
                    self.log(f"播放整点报时失败: {e}")

                if was_playing:
                    pygame.mixer.music.unpause()
                    self.log("报时结束，恢复播放。")

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

#第10部分
#第10部分
    def _execute_broadcast(self, task, trigger_time):
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
        if not AUDIO_AVAILABLE:
            self.log("错误: Pygame未初始化，无法播放音频。")
            return

        interval_type = task.get('interval_type', 'first')
        duration_seconds = int(task.get('interval_seconds', 0))
        repeat_count = int(task.get('interval_first', 1))

        playlist = []
        if task.get('audio_type') == 'single':
            if os.path.exists(task['content']): playlist = [task['content']] * repeat_count
        else:
            folder_path = task['content']
            if os.path.isdir(folder_path):
                all_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.lower().endswith(('.mp3', '.wav', '.ogg', '.flac', '.m4a'))]
                if task.get('play_order') == 'random': random.shuffle(all_files)
                playlist = all_files[:repeat_count]

        if not playlist:
            self.log(f"错误: 音频列表为空，任务 '{task['name']}' 无法播放。")
            return

        start_time = time.time()
        for i, audio_path in enumerate(playlist):
            if self._is_interrupted():
                self.log(f"任务 '{task['name']}' 被新指令中断。")
                return

            if interval_type == 'first':
                status_msg = f"[{task['name']}] 正在播放: {os.path.basename(audio_path)} ({i+1}/{len(playlist)})"
                self.update_playing_text(status_msg)

            self.log(f"正在播放: {os.path.basename(audio_path)} ({i+1}/{len(playlist)})")

            try:
                pygame.mixer.music.load(audio_path)
                pygame.mixer.music.set_volume(float(task.get('volume', 80)) / 100.0)
                pygame.mixer.music.play()

                last_text_update_time = 0

                while pygame.mixer.music.get_busy():
                    if self._is_interrupted():
                        pygame.mixer.music.stop()
                        return

                    if interval_type == 'seconds':
                        now = time.time()
                        elapsed = now - start_time
                        if elapsed >= duration_seconds:
                            pygame.mixer.music.stop()
                            self.log(f"已达到 {duration_seconds} 秒播放时长限制。")
                            return
                        if now - last_text_update_time >= 1.0:
                            remaining_seconds = int(duration_seconds - elapsed)
                            status_msg = f"[{task['name']}] 正在播放: {os.path.basename(audio_path)} (剩余 {remaining_seconds} 秒)"
                            self.update_playing_text(status_msg)
                            last_text_update_time = now

                    time.sleep(0.1)

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
                    pygame.mixer.music.set_volume(float(task.get('bgm_volume', 40)) / 100.0)
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

        interval_type = task.get('interval_type', 'first')
        duration_seconds = int(task.get('interval_seconds', 0))
        repeat_count = int(task.get('interval_first', 1))

        playlist = []
        if task.get('video_type') == 'single':
            if os.path.exists(task['content']):
                playlist = [task['content']] * repeat_count
        else:
            folder_path = task['content']
            if os.path.isdir(folder_path):
                video_extensions = ('.mp4', '.mkv', '.avi', '.mov', '.wmv', '.flv')
                all_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.lower().endswith(video_extensions)]
                if task.get('play_order') == 'random':
                    random.shuffle(all_files)
                playlist = all_files[:repeat_count]

        if not playlist:
            self.log(f"错误: 视频列表为空，任务 '{task['name']}' 无法播放。")
            return

        try:
            if AUDIO_AVAILABLE:
                pygame.mixer.music.stop()
                pygame.mixer.stop()

            instance = vlc.Instance()
            self.vlc_player = instance.media_player_new()

            self.root.after(0, self._create_video_window, task)
            time.sleep(0.5)

            if not (self.video_window and self.video_window.winfo_exists()):
                self.log("错误: 视频窗口创建失败，无法播放。")
                return

            self.vlc_player.set_hwnd(self.video_window.winfo_id())

            start_time = time.time()
            for i, video_path in enumerate(playlist):
                if self._is_interrupted() or stop_event.is_set():
                    self.log(f"任务 '{task['name']}' 在播放列表循环中被中断。")
                    break

                media = instance.media_new(video_path)
                self.vlc_player.set_media(media)
                self.vlc_player.play()

                rate_input = task.get('playback_rate', '1.0').strip()
                rate_match = re.match(r"(\d+(\.\d+)?)", rate_input)
                rate_val = float(rate_match.group(1)) if rate_match else 1.0
                self.vlc_player.set_rate(rate_val)
                self.vlc_player.audio_set_volume(int(task.get('volume', 80)))
                self.log(f"设置播放速率为: {rate_val}")

                time.sleep(0.5)

                last_text_update_time = 0
                while self.vlc_player.get_state() in {vlc.State.Opening, vlc.State.Playing, vlc.State.Paused}:
                    if self._is_interrupted() or stop_event.is_set():
                        self.log(f"视频任务 '{task['name']}' 在播放期间被中断。")
                        self.vlc_player.stop()
                        break

                    now = time.time()
                    if interval_type == 'seconds':
                        elapsed = now - start_time
                        if elapsed >= duration_seconds:
                            self.log(f"已达到 {duration_seconds} 秒播放时长限制。")
                            self.vlc_player.stop()
                            break

                        if now - last_text_update_time >= 1.0:
                            remaining_seconds = int(duration_seconds - elapsed)
                            status_text = "播放中" if self.vlc_player.is_playing() else "已暂停"
                            self.update_playing_text(f"[{task['name']}] {os.path.basename(video_path)} ({status_text} - 剩余 {remaining_seconds} 秒)")
                            last_text_update_time = now
                    else:
                         if now - last_text_update_time >= 1.0:
                            status_text = "播放中" if self.vlc_player.is_playing() else "已暂停"
                            self.update_playing_text(f"[{task['name']}] {os.path.basename(video_path)} ({i+1}/{len(playlist)} - {status_text})")
                            last_text_update_time = now

                    time.sleep(0.2)

                if (interval_type == 'seconds' and (time.time() - start_time) >= duration_seconds) or stop_event.is_set():
                    break

        except Exception as e:
            self.log(f"播放视频任务 '{task['name']}' 时发生错误: {e}")
        finally:
            if self.vlc_player:
                self.vlc_player.stop()
                self.vlc_player = None

            self.root.after(0, self._destroy_video_window)
            self.log(f"视频任务 '{task['name']}' 的播放逻辑结束。")

    def _create_video_window(self, task):
        if self.video_window and self.video_window.winfo_exists():
            self.video_window.destroy()

        self.video_window = ttk.Toplevel(self.root)
        self.active_modal_dialog = self.video_window # <--- 【BUG修复】
        self.video_window.title(f"正在播放: {task['name']}")
        self.video_window.configure(bg='black')

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
        self.video_window.focus_force()

    def _destroy_video_window(self):
        if self.video_window and self.video_window.winfo_exists():
            self.video_window.destroy()
        self.video_window = None
        self.active_modal_dialog = None # <--- 【BUG修复】

    def _handle_video_manual_stop(self, event=None):
        self.log("用户手动关闭视频窗口，将停止整个视频任务。")
        if self.video_stop_event:
            self.video_stop_event.set()
        if self.vlc_player:
            self.vlc_player.stop()

    def _handle_video_space(self, event=None):
        if self.vlc_player:
            self.vlc_player.pause()
            status = "暂停" if self.vlc_player.get_state() == vlc.State.Paused else "播放"
            self.log(f"空格键按下，视频已{status}。")

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
        self.active_modal_dialog = self.fullscreen_window # <--- 【BUG修复】
        self.fullscreen_window.attributes('-fullscreen', True)
        self.fullscreen_window.attributes('-topmost', True)
        self.fullscreen_window.configure(bg='black', cursor='none')
        self.fullscreen_window.protocol("WM_DELETE_WINDOW", lambda: None)
        self.fullscreen_window.bind("<Escape>", self._handle_esc_press)

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
            self.active_modal_dialog = None # <--- 【BUG修复】

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
        except Exception as e: self.log(f"加载任务失败: {e}")

    def load_settings(self):
        defaults = {
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
                "bg_image_interval": interval
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
        
        win.geometry(f'{width}x{height}+{x}+{y}')

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
        self.active_modal_dialog = dialog # <--- 【BUG修复】
        dialog.title("确认")
        dialog.resizable(False, False); dialog.transient(self.root); dialog.grab_set()

        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            dialog.destroy()

        ttk.Label(dialog, text="您想要如何操作？", font=self.font_12).pack(pady=20, padx=40)
        btn_frame = ttk.Frame(dialog); btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="退出程序", command=lambda: [cleanup_and_destroy(), self.quit_app()], bootstyle="danger").pack(side=LEFT, padx=10)
        if TRAY_AVAILABLE: ttk.Button(btn_frame, text="最小化到托盘", command=lambda: [cleanup_and_destroy(), self.hide_to_tray()], bootstyle="primary-outline").pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=cleanup_and_destroy).pack(side=LEFT, padx=10) # <--- 【BUG修复】
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】
        
        self.center_window(dialog, parent=self.root)

    def hide_to_tray(self):
        if not TRAY_AVAILABLE: messagebox.showwarning("功能不可用", "pystray 或 Pillow 库未安装，无法最小化到托盘。", parent=self.root); return
        self.root.withdraw()
        self.log("程序已最小化到系统托盘。")

    def show_from_tray(self, icon, item):
        self.root.after(0, self.root.deiconify)
        self.log("程序已从托盘恢复。")

    def quit_app(self, icon=None, item=None):
        if self.tray_icon: self.tray_icon.stop()
        self.running = False
        self.playback_command_queue.put(('STOP', None))

        # --- ↓↓↓ 新增代码：在保存设置前，先记录当前窗口的几何信息 ↓↓↓ ---
        # 只有当窗口不是最小化状态时才保存，避免保存一个看不见的位置
        if self.root.state() == 'normal':
            self.settings["window_geometry"] = self.root.geometry()
        # --- ↑↑↑ 新增代码结束 ↑↑↑ ---

        self.save_tasks()
        self.save_settings() # 这个函数会把包含新窗口位置的整个设置字典写入文件
        self.save_holidays()
        self.save_todos()
        self.save_screenshot_tasks()
        self.save_execute_tasks()

        if AUDIO_AVAILABLE and pygame.mixer.get_init(): pygame.mixer.quit()
       
        # self.root.destroy()
        # sys.exit()
        os._exit(0)

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
            (None, None, None), # Separator
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
        self.active_modal_dialog = dialog # <--- 【BUG修复】
        dialog.title("修改节假日" if holiday_to_edit else "添加节假日")
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()

        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            dialog.destroy()

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
            cleanup_and_destroy() # <--- 【BUG修复】

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=20)
        ttk.Button(button_frame, text="保存", command=save, bootstyle="primary", width=10).pack(side=LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=cleanup_and_destroy, width=10).pack(side=LEFT, padx=10) # <--- 【BUG修复】
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】

        self.center_window(dialog, parent=self.root)

    def show_holiday_context_menu(self, event):
        if self.is_locked: return
        iid = self.holiday_tree.identify_row(event.y)
        if not iid: return

        context_menu = tk.Menu(self.root, tearoff=0, font=self.font_11)

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
        self.active_modal_dialog = dialog # <--- 【BUG修复】
        dialog.title("修改待办事项" if todo_to_edit else "添加待办事项")
        dialog.resizable(True, True)
        dialog.minsize(640, 550)
        dialog.transient(self.root)
        dialog.grab_set()

        def cleanup_and_destroy(): # <--- 【BUG修复】
            self.active_modal_dialog = None
            dialog.destroy()

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

        ttk.Label(onetime_lf, text="执行日期:").grid(row=0, column=0, sticky='e', pady=5, padx=5)
        onetime_date_entry = ttk.Entry(onetime_lf, font=self.font_11, width=20)
        onetime_date_entry.grid(row=0, column=1, sticky='w', pady=5)
        self._bind_mousewheel_to_entry(onetime_date_entry, self._handle_date_scroll)
        ttk.Label(onetime_lf, text="执行时间:").grid(row=1, column=0, sticky='e', pady=5, padx=5)
        onetime_time_entry = ttk.Entry(onetime_lf, font=self.font_11, width=20)
        onetime_time_entry.grid(row=1, column=1, sticky='w', pady=5)
        self._bind_mousewheel_to_entry(onetime_time_entry, self._handle_time_scroll)

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
            recurring_daterange_entry.insert(0, todo_to_edit.get('date_range', '2000-01-01 ~ 2099-12-31'))
            recurring_interval_entry.insert(0, todo_to_edit.get('interval_minutes', '0'))
        else:
            onetime_date_entry.insert(0, now.strftime('%Y-%m-%d'))
            onetime_time_entry.insert(0, (now + timedelta(minutes=5)).strftime('%H:%M:%S'))
            recurring_time_entry.insert(0, now.strftime('%H:%M:%S'))
            recurring_weekday_entry.insert(0, '每周:1234567')
            recurring_daterange_entry.insert(0, '2000-01-01 ~ 2099-12-31')
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
            else:
                try:
                    interval = int(recurring_interval_entry.get().strip() or '0')
                    if not (0 <= interval <= 1440): raise ValueError
                except ValueError:
                    messagebox.showerror("格式错误", "循环间隔必须是 0-1440 之间的整数。", parent=dialog)
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
            cleanup_and_destroy() # <--- 【BUG修复】

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=4, pady=20)
        ttk.Button(button_frame, text="保存", command=save, bootstyle="primary", width=10).pack(side=LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=cleanup_and_destroy, width=10).pack(side=LEFT, padx=10) # <--- 【BUG修复】
        dialog.protocol("WM_DELETE_WINDOW", cleanup_and_destroy) # <--- 【BUG修复】


#第13部分
#第13部分
    def show_todo_context_menu(self, event):
        if self.is_locked: return
        iid = self.todo_tree.identify_row(event.y)
        if not iid: return

        context_menu = tk.Menu(self.root, tearoff=0, font=self.font_11)
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
        if self._is_in_holiday(now): return

        now_str_dt = now.strftime('%Y-%m-%d %H:%M:%S')
        now_str_date = now.strftime('%Y-%m-%d')
        now_str_time = now.strftime('%H:%M:%S')

        for index, todo in enumerate(self.todos):
            if todo.get('status') != '启用': continue

            if todo.get('type') == 'onetime':
                if todo.get('remind_datetime') == now_str_dt:
                    self.log(f"触发一次性待办事项: {todo['name']}")
                    todo_with_index = todo.copy()
                    todo_with_index['original_index'] = index
                    self.reminder_queue.put(todo_with_index)

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

                triggered = False
                for trigger_time in [t.strip() for t in todo.get('start_times', '').split(',')]:
                    if trigger_time == now_str_time and todo.get('last_run', {}).get(trigger_time) != now_str_date:
                        triggered = True
                        todo.setdefault('last_run', {})[trigger_time] = now_str_date
                        break

                interval = todo.get('interval_minutes', 0)
                if not triggered and interval > 0 and todo.get('start_times'):
                    last_run_str = todo.get('last_interval_run')
                    if last_run_str:
                        try:
                            last_run_dt = datetime.strptime(last_run_str, '%Y-%m-%d %H:%M:%S')
                            if now >= last_run_dt + timedelta(minutes=interval):
                                triggered = True
                        except ValueError: pass

                if triggered:
                    self.log(f"触发循环待办事项: {todo['name']}")
                    todo_with_index = todo.copy()
                    todo_with_index['original_index'] = index
                    self.reminder_queue.put(todo_with_index)
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
        self.active_modal_dialog = reminder_win # <--- 【BUG修复】
        reminder_win.title(f"待办事项提醒 - {todo.get('name')}")
        
        # --- 核心修改：完全按照您的要求，设置一个固定的窗口尺寸 ---
        reminder_win.geometry("640x480")
        # 为了防止窗口被意外缩小，我们禁止调整大小
        reminder_win.resizable(False, False)
        # --- 修改结束 ---

        reminder_win.attributes('-topmost', True)
        reminder_win.lift()
        reminder_win.focus_force()
        reminder_win.after(1000, lambda: reminder_win.attributes('-topmost', False))

        original_index = todo.get('original_index')
        task_type = todo.get('type')

        # --- 使用我们已验证过可以稳定显示所有组件的 Grid 布局 ---
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

        # 使用原始的 Text 和 Scrollbar 组件，这是最可靠的组合
        content_text_widget = tk.Text(content_frame, font=self.font_11, wrap=WORD, bd=0, highlightthickness=0)
        content_text_widget.grid(row=0, column=0, sticky='nsew')
        
        scrollbar = ttk.Scrollbar(content_frame, orient=VERTICAL, command=content_text_widget.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        
        content_text_widget.config(yscrollcommand=scrollbar.set)

        content_text_widget.insert('1.0', todo.get('content', ''))
        content_text_widget.config(state='disabled')

        # 在按钮区内部配置按钮
        if task_type == 'onetime':
            btn_frame.columnconfigure((0, 1, 2), weight=1)
            ttk.Button(btn_frame, text="已完成", bootstyle="success", command=lambda: handle_complete()).grid(row=0, column=0, padx=5, ipady=4, sticky='ew')
            ttk.Button(btn_frame, text="稍后提醒", bootstyle="outline-secondary", command=lambda: handle_snooze()).grid(row=0, column=1, padx=5, ipady=4, sticky='ew')
            ttk.Button(btn_frame, text="删除任务", bootstyle="danger", command=lambda: handle_delete()).grid(row=0, column=2, padx=5, ipady=4, sticky='ew')
        else:
            btn_frame.columnconfigure((0, 1), weight=1)
            ttk.Button(btn_frame, text="本次完成", bootstyle="primary", command=lambda: close_and_release()).grid(row=0, column=0, padx=5, ipady=4, sticky='ew')
            ttk.Button(btn_frame, text="删除任务", bootstyle="danger", command=lambda: handle_delete()).grid(row=0, column=1, padx=5, ipady=4, sticky='ew')
        
        # --- 逻辑处理部分（没有变化） ---
        def close_and_release():
            self.is_reminder_active = False
            self.active_modal_dialog = None # <--- 【BUG修复】
            reminder_win.destroy()

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
    root = ttk.Window(themename="litera")
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
