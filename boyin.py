import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText
from tkinter import messagebox, filedialog, simpledialog, font

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

# 为DPI感知导入ctypes
try:
    import ctypes
except ImportError:
    pass


# 尝试导入所需库
TRAY_AVAILABLE = False
try:
    from pystray import MenuItem as item, Icon
    from PIL import Image, ImageTk
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
    # 为提示音和报时预留单独的通道，避免与背景音乐冲突
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
    print("警告: psutil 未安装，无法获取机器码，注册功能不可用。")

# --- 导入 VLC 库 ---
VLC_AVAILABLE = False
try:
    import vlc
    VLC_AVAILABLE = True
except ImportError:
    print("警告: python-vlc 未安装，视频播放功能不可用。")
except Exception as e:
    print(f"警告: vlc 初始化失败 - {e}，视频播放功能不可用。")


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
TODO_FILE = os.path.join(application_path, "todos.json")
PROMPT_FOLDER = os.path.join(application_path, "提示音")
AUDIO_FOLDER = os.path.join(application_path, "音频文件")
BGM_FOLDER = os.path.join(application_path, "文稿背景")
VOICE_SCRIPT_FOLDER = os.path.join(application_path, "语音文稿")
ICON_FILE = resource_path("icon.ico")

REMINDER_SOUND_FILE = os.path.join(PROMPT_FOLDER, "reminder.wav")
CHIME_FOLDER = os.path.join(AUDIO_FOLDER, "整点报时")

REGISTRY_KEY_PATH = r"Software\创翔科技\TimedBroadcastApp"
REGISTRY_PARENT_KEY_PATH = r"Software\创翔科技"


class TimedBroadcastApp:
    def __init__(self, root):
        self.root = root
        self.root.title(" 创翔多功能定时播音旗舰版")
        # 修复1: 设置初始和最小尺寸
        self.root.geometry("1024x768")
        self.root.minsize(1024, 768)

        if os.path.exists(ICON_FILE):
            try:
                self.root.iconbitmap(ICON_FILE)
            except Exception as e:
                print(f"加载窗口图标失败: {e}")

        self.tasks = []
        self.holidays = []
        self.todos = []
        self.settings = {}
        self.running = True
        self.tray_icon = None
        self.is_locked = False
        self.is_app_locked_down = False

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
        self.load_lock_password()

        self._apply_global_font()

        self.check_authorization()

        self.create_widgets()
        self.load_tasks()
        self.load_holidays()
        self.load_todos()

        self.start_background_threads()
        self.root.protocol("WM_DELETE_WINDOW", self.show_quit_dialog)
        self.start_tray_icon_thread()

        if self.settings.get("lock_on_start", False) and self.lock_password_b64:
            self.root.after(100, self.perform_initial_lock)

        if self.settings.get("start_minimized", False):
            self.root.after(100, self.hide_to_tray)

        if self.is_app_locked_down:
            self.root.after(100, self.perform_lockdown)

    def _apply_global_font(self):
        """在创建控件前，应用全局字体设置"""
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
        style.configure("Treeview", font=self.font_11, rowheight=28)
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
        for folder in [PROMPT_FOLDER, AUDIO_FOLDER, BGM_FOLDER, VOICE_SCRIPT_FOLDER]:
            if not os.path.exists(folder):
                os.makedirs(folder)

    def create_widgets(self):
        # 修复2: 移除固定高度，让其自适应内容
        self.status_frame = ttk.Frame(self.root, style='secondary.TFrame')
        self.status_frame.pack(side=BOTTOM, fill=X)
        self.create_status_bar_content()

        self.nav_frame = ttk.Frame(self.root, width=180, style='light.TFrame')
        self.nav_frame.pack(side=LEFT, fill=Y)
        self.nav_frame.pack_propagate(False)

        self.page_container = ttk.Frame(self.root)
        self.page_container.pack(side=LEFT, fill=BOTH, expand=True)

        nav_button_titles = ["定时广播", "节假日", "待办事项", "设置", "注册软件", "超级管理"]

        for i, title in enumerate(nav_button_titles):
            cmd = lambda t=title: self.switch_page(t)
            if title == "超级管理":
                cmd = self._prompt_for_super_admin_password
            
            btn = ttk.Button(self.nav_frame, text=title,
                           style='Link.TButton', command=cmd)
            btn.pack(fill=X, pady=1, ipady=8, padx=10)
            self.nav_buttons[title] = btn
            
        style = ttk.Style.get_instance()
        style.configure('Link.TButton', font=self.font_13_bold, anchor='w')

        self.main_frame = ttk.Frame(self.page_container)
        self.pages["定时广播"] = self.main_frame
        self.create_scheduled_broadcast_page()

        self.current_page = self.main_frame
        self.switch_page("定时广播")

        self.update_status_bar()
        self.log(" 创翔多功能定时播音旗舰版软件已启动")

    def create_status_bar_content(self):
        self.status_labels = []
        status_texts = ["当前时间", "系统状态", "播放状态", "任务数量", "待办事项"]

        copyright_label = ttk.Label(self.status_frame, text="© 创翔科技", font=self.font_11,
                                    bootstyle=(SECONDARY, INVERSE), padding=(15, 0))
        copyright_label.pack(side=RIGHT, padx=2)
        
        # 修复4: 在这里创建按钮，但不显示
        self.statusbar_unlock_button = ttk.Button(self.status_frame, text="🔓 解锁",
                                                  bootstyle="success",
                                                  command=self._prompt_for_password_unlock)
        
        for i, text in enumerate(status_texts):
            label = ttk.Label(self.status_frame, text=f"{text}: --", font=self.font_11,
                              bootstyle=(PRIMARY, INVERSE) if i % 2 == 0 else (SECONDARY, INVERSE),
                              padding=(15, 8)) # 增加垂直内边距
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

        if self.current_page:
            self.current_page.pack_forget()

        for title, btn in self.nav_buttons.items():
            btn.config(bootstyle="light")

        target_frame = None
        if page_name == "定时广播":
            target_frame = self.pages["定时广播"]
        elif page_name == "节假日":
            if page_name not in self.pages:
                self.pages[page_name] = self.create_holiday_page()
            target_frame = self.pages[page_name]
        elif page_name == "待办事项":
            if page_name not in self.pages:
                self.pages[page_name] = self.create_todo_page()
            target_frame = self.pages[page_name]
        elif page_name == "设置":
            if page_name not in self.pages:
                self.pages[page_name] = self.create_settings_page()
            self._refresh_settings_ui()
            target_frame = self.pages[page_name]
        elif page_name == "注册软件":
            if page_name not in self.pages:
                self.pages[page_name] = self.create_registration_page()
            target_frame = self.pages[page_name]
        elif page_name == "超级管理":
            if page_name not in self.pages:
                self.pages[page_name] = self.create_super_admin_page()
            target_frame = self.pages[page_name]
        else:
            self.log(f"功能开发中: {page_name}")
            target_frame = self.pages["定时广播"]
            page_name = "定时广播"

        target_frame.pack(in_=self.page_container, fill=BOTH, expand=True)
        self.current_page = target_frame
        self.current_page_name = page_name

        selected_btn = self.nav_buttons.get(page_name)
        if selected_btn:
            selected_btn.config(bootstyle="primary")

    def _prompt_for_super_admin_password(self):
        if self.auth_info['status'] != 'Permanent':
            messagebox.showerror("权限不足", "此功能仅对“永久授权”用户开放。\n\n请注册软件并获取永久授权后重试。")
            self.log("非永久授权用户尝试进入超级管理模块被阻止。")
            return

        dialog = ttk.Toplevel(self.root)
        dialog.title("身份验证")
        dialog.geometry("350x180")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        self.center_window(dialog, 350, 180)

        result = [None]

        ttk.Label(dialog, text="请输入超级管理员密码:", font=self.font_11).pack(pady=20)
        password_entry = ttk.Entry(dialog, show='*', font=self.font_11, width=25)
        password_entry.pack(pady=5)
        password_entry.focus_set()

        def on_confirm():
            result[0] = password_entry.get()
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="确定", command=on_confirm, bootstyle="primary", width=8).pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=on_cancel, width=8).pack(side=LEFT, padx=10)
        dialog.bind('<Return>', lambda event: on_confirm())

        self.root.wait_window(dialog)
        entered_password = result[0]

        correct_password = datetime.now().strftime('%Y%m%d')

        if entered_password == correct_password:
            self.log("超级管理员密码正确，进入管理模块。")
            self.switch_page("超级管理")
        elif entered_password is not None:
            messagebox.showerror("验证失败", "密码错误！")
            self.log("尝试进入超级管理模块失败：密码错误。")

    def create_registration_page(self):
        page_frame = ttk.Frame(self.page_container, padding=20)
        title_label = ttk.Label(page_frame, text="注册软件", font=self.font_14_bold, bootstyle="primary")
        title_label.pack(anchor=W)

        main_content_frame = ttk.Frame(page_frame)
        main_content_frame.pack(pady=10)

        machine_code_frame = ttk.Frame(main_content_frame)
        machine_code_frame.pack(fill=X, pady=10)
        ttk.Label(machine_code_frame, text="机器码:", font=self.font_12).pack(side=LEFT)
        machine_code_val = self.get_machine_code()
        machine_code_entry = ttk.Entry(machine_code_frame, font=self.font_12, width=30, bootstyle="danger")
        machine_code_entry.pack(side=LEFT, padx=10)
        machine_code_entry.insert(0, machine_code_val)
        machine_code_entry.config(state='readonly')

        reg_code_frame = ttk.Frame(main_content_frame)
        reg_code_frame.pack(fill=X, pady=10)
        ttk.Label(reg_code_frame, text="注册码:", font=self.font_12).pack(side=LEFT)
        self.reg_code_entry = ttk.Entry(reg_code_frame, font=self.font_12, width=30)
        self.reg_code_entry.pack(side=LEFT, padx=10)

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

    def cancel_registration(self):
        if not messagebox.askyesno("确认操作", "您确定要取消当前注册吗？\n取消后，软件将恢复到试用或过期状态。"):
            return

        self.log("用户请求取消注册...")
        self._save_to_registry('RegistrationStatus', '')
        self._save_to_registry('RegistrationDate', '')

        self.check_authorization()

        messagebox.showinfo("操作完成", f"注册已成功取消。\n当前授权状态: {self.auth_info['message']}")
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
            messagebox.showerror("依赖缺失", "psutil 库未安装，无法获取机器码。软件将退出。")
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
            messagebox.showerror("错误", f"无法获取机器码：{e}\n软件将退出。")
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
            messagebox.showwarning("提示", "请输入注册码。")
            return

        numeric_machine_code = self.get_machine_code()
        correct_codes = self._calculate_reg_codes(numeric_machine_code)

        today_str = datetime.now().strftime('%Y-%m-%d')

        if entered_code == correct_codes['monthly']:
            self._save_to_registry('RegistrationStatus', 'Monthly')
            self._save_to_registry('RegistrationDate', today_str)
            messagebox.showinfo("注册成功", "恭喜您，月度授权已成功激活！")
            self.check_authorization()
        elif entered_code == correct_codes['permanent']:
            self._save_to_registry('RegistrationStatus', 'Permanent')
            self._save_to_registry('RegistrationDate', today_str)
            messagebox.showinfo("注册成功", "恭喜您，永久授权已成功激活！")
            self.check_authorization()
        else:
            messagebox.showerror("注册失败", "您输入的注册码无效，请重新核对。")

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
        messagebox.showerror("授权过期", "您的软件试用期或授权已到期，功能已受限。\n请在“注册软件”页面输入有效注册码以继续使用。")
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
        dialog.title("卸载软件 - 身份验证")
        dialog.geometry("350x180")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        self.center_window(dialog, 350, 180)

        result = [None]

        ttk.Label(dialog, text="请输入卸载密码:", font=self.font_11).pack(pady=20)
        password_entry = ttk.Entry(dialog, show='*', font=self.font_11, width=25)
        password_entry.pack(pady=5)
        password_entry.focus_set()

        def on_confirm():
            result[0] = password_entry.get()
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="确定", command=on_confirm, bootstyle="primary", width=8).pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=on_cancel, width=8).pack(side=LEFT, padx=10)
        dialog.bind('<Return>', lambda event: on_confirm())

        self.root.wait_window(dialog)
        entered_password = result[0]

        correct_password = datetime.now().strftime('%Y%m%d')[::-1]

        if entered_password == correct_password:
            self.log("卸载密码正确，准备执行卸载操作。")
            self._perform_uninstall()
        elif entered_password is not None:
            messagebox.showerror("验证失败", "密码错误！", parent=self.root)
            self.log("尝试卸载软件失败：密码错误。")
#标记
    def _perform_uninstall(self):
        if not messagebox.askyesno(
            "！！！最终警告！！！",
            "您确定要卸载本软件吗？\n\n此操作将永久删除：\n- 所有注册表信息\n- 所有配置文件 (节目单, 设置, 节假日, 待办事项)\n- 所有数据文件夹 (音频, 提示音, 文稿等)\n\n此操作【绝对无法恢复】！\n\n点击“是”将立即开始清理。",
            icon='error'
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

        folders_to_delete = [PROMPT_FOLDER, AUDIO_FOLDER, BGM_FOLDER, VOICE_SCRIPT_FOLDER]
        for folder in folders_to_delete:
            if os.path.isdir(folder):
                try:
                    shutil.rmtree(folder)
                    self.log(f"成功删除文件夹: {os.path.basename(folder)}")
                except Exception as e:
                    self.log(f"删除文件夹 {os.path.basename(folder)} 时出错: {e}")

        files_to_delete = [TASK_FILE, SETTINGS_FILE, HOLIDAY_FILE, TODO_FILE]
        for file in files_to_delete:
            if os.path.isfile(file):
                try:
                    os.remove(file)
                    self.log(f"成功删除文件: {os.path.basename(file)}")
                except Exception as e:
                    self.log(f"删除文件 {os.path.basename(file)} 时出错: {e}")

        self.log("软件数据清理完成。")
        messagebox.showinfo("卸载完成", "软件相关的数据和配置已全部清除。\n\n请手动删除本程序（.exe文件）以完成卸载。\n\n点击“确定”后软件将退出。")

        os._exit(0)

    def _backup_all_settings(self):
        self.log("开始备份所有设置...")
        try:
            backup_data = {
                'backup_date': datetime.now().isoformat(), 'tasks': self.tasks, 'holidays': self.holidays,
                'todos': self.todos, 'settings': self.settings,
                'lock_password_b64': self._load_from_registry("LockPasswordB64")
            }
            filename = filedialog.asksaveasfilename(
                title="备份所有设置到...", defaultextension=".json",
                initialfile=f"boyin_backup_{datetime.now().strftime('%Y%m%d')}.json",
                filetypes=[("JSON Backup", "*.json")], initialdir=application_path
            )
            if filename:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(backup_data, f, ensure_ascii=False, indent=2)
                self.log(f"所有设置已成功备份到: {os.path.basename(filename)}")
                messagebox.showinfo("备份成功", f"所有设置已成功备份到:\n{filename}")
        except Exception as e:
            self.log(f"备份失败: {e}"); messagebox.showerror("备份失败", f"发生错误: {e}")

    def _restore_all_settings(self):
        if not messagebox.askyesno("确认操作", "您确定要还原所有设置吗？\n当前所有配置将被立即覆盖。"):
            return

        self.log("开始还原所有设置...")
        filename = filedialog.askopenfilename(
            title="选择要还原的备份文件",
            filetypes=[("JSON Backup", "*.json")], initialdir=application_path
        )
        if not filename: return

        try:
            with open(filename, 'r', encoding='utf-8') as f: backup_data = json.load(f)

            required_keys = ['tasks', 'holidays', 'settings', 'lock_password_b64']
            if not all(key in backup_data for key in required_keys):
                messagebox.showerror("还原失败", "备份文件格式无效或已损坏。"); return

            self.tasks = backup_data['tasks']
            self.holidays = backup_data['holidays']
            self.todos = backup_data.get('todos', [])
            self.settings = backup_data['settings']
            self.lock_password_b64 = backup_data['lock_password_b64']

            self.save_tasks()
            self.save_holidays()
            self.save_todos()
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)

            if self.lock_password_b64:
                self._save_to_registry("LockPasswordB64", self.lock_password_b64)
            else:
                self._save_to_registry("LockPasswordB64", "")

            self.update_task_list()
            self.update_holiday_list()
            self.update_todo_list()
            self._refresh_settings_ui()
            
            self._apply_global_font()
            messagebox.showinfo("还原成功", "所有设置已成功还原。\n软件需要重启以应用字体更改。")
            self.log("所有设置已从备份文件成功还原。")

            self.root.after(100, lambda: self.switch_page("定时广播"))

        except Exception as e:
            self.log(f"还原失败: {e}"); messagebox.showerror("还原失败", f"发生错误: {e}")

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
            "您真的要重置整个软件吗？\n\n此操作将：\n- 清空所有节目单 (但保留音频文件)\n- 清空所有节假日和待办事项\n- 清除锁定密码\n- 重置所有系统设置 (包括字体)\n\n此操作【无法恢复】！软件将在重置后提示您重启。"
        ): return

        self.log("开始执行软件重置...")
        try:
            original_askyesno = messagebox.askyesno
            messagebox.askyesno = lambda title, message: True
            self.clear_all_tasks(delete_associated_files=False)
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
            messagebox.showinfo("重置成功", "软件已恢复到初始状态。\n\n请点击“确定”后手动关闭并重新启动软件。")
        except Exception as e:
            self.log(f"重置失败: {e}"); messagebox.showerror("重置失败", f"发生错误: {e}")

    def create_scheduled_broadcast_page(self):
        page_frame = self.pages["定时广播"]

        top_frame = ttk.Frame(page_frame, padding=(10, 10))
        top_frame.pack(fill=X)
        title_label = ttk.Label(top_frame, text="定时广播", font=self.font_14_bold)
        title_label.pack(side=LEFT)

        add_btn = ttk.Button(top_frame, text="添加节目", command=self.add_task, bootstyle="primary")
        add_btn.pack(side=LEFT, padx=10)

        self.top_right_btn_frame = ttk.Frame(top_frame)
        self.top_right_btn_frame.pack(side=RIGHT)

        batch_buttons = [
            ("全部启用", self.enable_all_tasks, 'success'),
            ("全部禁用", self.disable_all_tasks, 'warning'),
            ("禁音频节目", lambda: self._set_tasks_status_by_type('audio', '禁用'), 'warning-outline'),
            ("禁语音节目", lambda: self._set_tasks_status_by_type('voice', '禁用'), 'warning-outline'),
            ("禁视频节目", lambda: self._set_tasks_status_by_type('video', '禁用'), 'warning-outline'),
            ("统一音量", self.set_uniform_volume, 'info'),
            ("清空节目", self.clear_all_tasks, 'danger')
        ]
        for text, cmd, style in batch_buttons:
            btn = ttk.Button(self.top_right_btn_frame, text=text, command=cmd, bootstyle=style)
            btn.pack(side=LEFT, padx=3)

        self.lock_button = ttk.Button(self.top_right_btn_frame, text="锁定", command=self.toggle_lock_state, bootstyle='danger')
        self.lock_button.pack(side=LEFT, padx=3)
        if not WIN32_AVAILABLE:
            self.lock_button.config(state=DISABLED, text="锁定(Win)")

        io_buttons = [("导入节目单", self.import_tasks, 'info-outline'), ("导出节目单", self.export_tasks, 'info-outline')]
        for text, cmd, style in io_buttons:
            btn = ttk.Button(self.top_right_btn_frame, text=text, command=cmd, bootstyle=style)
            btn.pack(side=LEFT, padx=3)

        stats_frame = ttk.Frame(page_frame, padding=(10, 5))
        stats_frame.pack(fill=X)
        self.stats_label = ttk.Label(stats_frame, text="节目单：0", font=self.font_11, bootstyle="secondary")
        self.stats_label.pack(side=LEFT, fill=X, expand=True)

        table_frame = ttk.Frame(page_frame, padding=(10, 5))
        table_frame.pack(fill=BOTH, expand=True)
        columns = ('节目名称', '状态', '开始时间', '模式', '文件或内容', '音量', '周几/几号', '日期范围')
        self.task_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=12, selectmode='extended', bootstyle="primary")

        self.task_tree.heading('节目名称', text='节目名称')
        self.task_tree.column('节目名称', width=200, anchor='w')
        self.task_tree.heading('状态', text='状态')
        self.task_tree.column('状态', width=70, anchor='center', stretch=NO)
        self.task_tree.heading('开始时间', text='开始时间')
        self.task_tree.column('开始时间', width=100, anchor='center', stretch=NO)
        self.task_tree.heading('模式', text='模式')
        self.task_tree.column('模式', width=70, anchor='center', stretch=NO)
        self.task_tree.heading('文件或内容', text='文件或内容')
        self.task_tree.column('文件或内容', width=300, anchor='w')
        self.task_tree.heading('音量', text='音量')
        self.task_tree.column('音量', width=70, anchor='center', stretch=NO)
        self.task_tree.heading('周几/几号', text='周几/几号')
        self.task_tree.column('周几/几号', width=100, anchor='center')
        self.task_tree.heading('日期范围', text='日期范围')
        self.task_tree.column('日期范围', width=120, anchor='center')

        self.task_tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar = ttk.Scrollbar(table_frame, orient=VERTICAL, command=self.task_tree.yview, bootstyle="round")
        scrollbar.pack(side=RIGHT, fill=Y)
        self.task_tree.configure(yscrollcommand=scrollbar.set)

        self.task_tree.bind("<Button-3>", self.show_context_menu)
        self.task_tree.bind("<Double-1>", self.on_double_click_edit)
        self._enable_drag_selection(self.task_tree)

        playing_frame = ttk.LabelFrame(page_frame, text="正在播：", padding=(10, 5))
        playing_frame.pack(fill=X, padx=10, pady=5)
        self.playing_label = ttk.Label(playing_frame, text="等待播放...", font=self.font_11,
                                       anchor=W, justify=LEFT, padding=5, bootstyle="warning")
        self.playing_label.pack(fill=X, expand=True, ipady=4)
        self.update_playing_text("等待播放...")

        log_frame = ttk.LabelFrame(page_frame, text="", padding=(10, 5))
        log_frame.pack(fill=BOTH, expand=True, padx=10, pady=5)

        log_header_frame = ttk.Frame(log_frame)
        log_header_frame.pack(fill=X)
        log_label = ttk.Label(log_header_frame, text="日志：", font=self.font_11_bold)
        log_label.pack(side=LEFT)
        self.clear_log_btn = ttk.Button(log_header_frame, text="清除日志", command=self.clear_log,
                                        bootstyle="light-outline")
        self.clear_log_btn.pack(side=LEFT, padx=10)

        self.log_text = ScrolledText(log_frame, height=6, font=self.font_11,
                                                  wrap=WORD, state='disabled', autohide=True)
        self.log_text.pack(fill=BOTH, expand=True)

    def create_settings_page(self):
        settings_frame = ttk.Frame(self.page_container, padding=20)

        title_label = ttk.Label(settings_frame, text="系统设置", font=self.font_14_bold, bootstyle="primary")
        title_label.pack(anchor=W, pady=(0, 10))

        general_frame = ttk.LabelFrame(settings_frame, text="通用设置", padding=(15, 10))
        general_frame.pack(fill=X, pady=10)

        self.autostart_var = tk.BooleanVar()
        self.start_minimized_var = tk.BooleanVar()
        self.lock_on_start_var = tk.BooleanVar()
        self.bg_image_interval_var = tk.StringVar()

        ttk.Checkbutton(general_frame, text="登录windows后自动启动", variable=self.autostart_var, bootstyle="round-toggle", command=self._handle_autostart_setting).pack(fill=X, pady=5)
        ttk.Checkbutton(general_frame, text="启动后最小化到系统托盘", variable=self.start_minimized_var, bootstyle="round-toggle", command=self.save_settings).pack(fill=X, pady=5)

        lock_and_buttons_frame = ttk.Frame(general_frame)
        lock_and_buttons_frame.pack(fill=X, pady=5)

        self.lock_on_start_cb = ttk.Checkbutton(lock_and_buttons_frame, text="启动软件后立即锁定", variable=self.lock_on_start_var, bootstyle="round-toggle", command=self._handle_lock_on_start_toggle)
        self.lock_on_start_cb.grid(row=0, column=0, sticky='w')
        if not WIN32_AVAILABLE:
            self.lock_on_start_cb.config(state=DISABLED)

        ttk.Label(lock_and_buttons_frame, text="(请先在主界面设置锁定密码)", font=self.font_9, bootstyle="secondary").grid(row=1, column=0, sticky='w', padx=20)

        self.clear_password_btn = ttk.Button(lock_and_buttons_frame, text="清除锁定密码", command=self.clear_lock_password, bootstyle="warning-outline")
        self.clear_password_btn.grid(row=0, column=1, padx=20)

        action_buttons_frame = ttk.Frame(general_frame)
        action_buttons_frame.pack(fill=X, pady=8)

        self.cancel_bg_images_btn = ttk.Button(action_buttons_frame, text="取消所有节目背景图片", command=self._cancel_all_background_images, bootstyle="info-outline")
        self.cancel_bg_images_btn.pack(side=LEFT, padx=5)
        
        self.restore_video_speed_btn = ttk.Button(action_buttons_frame, text="恢复所有视频节目播放速度", command=self._restore_all_video_speeds, bootstyle="info-outline")
        self.restore_video_speed_btn.pack(side=LEFT, padx=5)

        bg_interval_frame = ttk.Frame(general_frame)
        bg_interval_frame.pack(fill=X, pady=8)
        ttk.Label(bg_interval_frame, text="背景图片切换间隔:").pack(side=LEFT)
        interval_entry = ttk.Entry(bg_interval_frame, textvariable=self.bg_image_interval_var, font=self.font_11, width=5)
        interval_entry.pack(side=LEFT, padx=5)
        ttk.Label(bg_interval_frame, text="秒 (范围: 5-60)", font=self.font_10, bootstyle="secondary").pack(side=LEFT)
        ttk.Button(bg_interval_frame, text="确定", command=self._validate_bg_interval, bootstyle="primary-outline").pack(side=LEFT, padx=10)

        font_frame = ttk.Frame(general_frame)
        font_frame.pack(fill=X, pady=8)

        ttk.Label(font_frame, text="软件字体:").pack(side=LEFT)

        try:
            available_fonts = sorted(list(font.families()))
        except:
            available_fonts = ["Microsoft YaHei"]

        self.font_var = tk.StringVar()

        font_combo = ttk.Combobox(font_frame, textvariable=self.font_var, values=available_fonts, font=self.font_10, width=25, state='readonly')
        font_combo.pack(side=LEFT, padx=10)
        font_combo.bind("<<ComboboxSelected>>", self._on_font_selected)

        restore_font_btn = ttk.Button(font_frame, text="恢复默认字体", command=self._restore_default_font, bootstyle="secondary-outline")
        restore_font_btn.pack(side=LEFT, padx=10)

        time_chime_frame = ttk.LabelFrame(settings_frame, text="整点报时", padding=(15, 10))
        time_chime_frame.pack(fill=X, pady=10)

        self.time_chime_enabled_var = tk.BooleanVar()
        self.time_chime_voice_var = tk.StringVar()
        self.time_chime_speed_var = tk.StringVar()
        self.time_chime_pitch_var = tk.StringVar()

        chime_control_frame = ttk.Frame(time_chime_frame)
        chime_control_frame.pack(fill=X, pady=5)

        ttk.Checkbutton(chime_control_frame, text="启用整点报时功能", variable=self.time_chime_enabled_var, bootstyle="round-toggle", command=self._handle_time_chime_toggle).pack(side=LEFT)

        available_voices = self.get_available_voices()
        # 修复8: 增加 Combobox 宽度
        self.chime_voice_combo = ttk.Combobox(chime_control_frame, textvariable=self.time_chime_voice_var, values=available_voices, font=self.font_10, width=60, state='readonly')
        self.chime_voice_combo.pack(side=LEFT, padx=10)
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

        self.daily_shutdown_enabled_var = tk.BooleanVar()
        self.daily_shutdown_time_var = tk.StringVar()
        self.weekly_shutdown_enabled_var = tk.BooleanVar()
        self.weekly_shutdown_time_var = tk.StringVar()
        self.weekly_shutdown_days_var = tk.StringVar()
        self.weekly_reboot_enabled_var = tk.BooleanVar()
        self.weekly_reboot_time_var = tk.StringVar()
        self.weekly_reboot_days_var = tk.StringVar()

        daily_frame = ttk.Frame(power_frame)
        daily_frame.pack(fill=X, pady=4)
        ttk.Checkbutton(daily_frame, text="每天关机", variable=self.daily_shutdown_enabled_var, bootstyle="round-toggle", command=self.save_settings).pack(side=LEFT, padx=(0,10))
        daily_time_entry = ttk.Entry(daily_frame, textvariable=self.daily_shutdown_time_var, font=self.font_11, width=15)
        daily_time_entry.pack(side=LEFT, padx=10)
        self._bind_mousewheel_to_entry(daily_time_entry, self._handle_time_scroll)
        ttk.Button(daily_frame, text="设置", bootstyle="primary-outline", command=lambda: self.show_single_time_dialog(self.daily_shutdown_time_var)).pack(side=LEFT)

        weekly_frame = ttk.Frame(power_frame)
        weekly_frame.pack(fill=X, pady=4)
        ttk.Checkbutton(weekly_frame, text="每周关机", variable=self.weekly_shutdown_enabled_var, bootstyle="round-toggle", command=self.save_settings).pack(side=LEFT, padx=(0,10))
        ttk.Entry(weekly_frame, textvariable=self.weekly_shutdown_days_var, font=self.font_11, width=20).pack(side=LEFT, padx=(10,5))
        weekly_shutdown_time_entry = ttk.Entry(weekly_frame, textvariable=self.weekly_shutdown_time_var, font=self.font_11, width=15)
        weekly_shutdown_time_entry.pack(side=LEFT, padx=5)
        self._bind_mousewheel_to_entry(weekly_shutdown_time_entry, self._handle_time_scroll)
        ttk.Button(weekly_frame, text="设置", bootstyle="primary-outline", command=lambda: self.show_power_week_time_dialog("设置每周关机", self.weekly_shutdown_days_var, self.weekly_shutdown_time_var)).pack(side=LEFT)

        reboot_frame = ttk.Frame(power_frame)
        reboot_frame.pack(fill=X, pady=4)
        ttk.Checkbutton(reboot_frame, text="每周重启", variable=self.weekly_reboot_enabled_var, bootstyle="round-toggle", command=self.save_settings).pack(side=LEFT, padx=(0,10))
        ttk.Entry(reboot_frame, textvariable=self.weekly_reboot_days_var, font=self.font_11, width=20).pack(side=LEFT, padx=(10,5))
        weekly_reboot_time_entry = ttk.Entry(reboot_frame, textvariable=self.weekly_reboot_time_var, font=self.font_11, width=15)
        weekly_reboot_time_entry.pack(side=LEFT, padx=5)
        self._bind_mousewheel_to_entry(weekly_reboot_time_entry, self._handle_time_scroll)
        ttk.Button(reboot_frame, text="设置", bootstyle="primary-outline", command=lambda: self.show_power_week_time_dialog("设置每周重启", self.weekly_reboot_days_var, self.weekly_reboot_time_var)).pack(side=LEFT)

        return settings_frame

    def _restore_all_video_speeds(self):
        """恢复所有视频节目的播放速度为1.0x"""
        if not self.tasks:
            messagebox.showinfo("提示", "当前没有节目，无需操作。")
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
            messagebox.showinfo("操作成功", f"已成功将 {count} 个视频节目的播放速度恢复为默认值(1.0x)。")
        else:
            messagebox.showinfo("提示", "所有视频节目已经是默认播放速度，无需恢复。")

    def _on_font_selected(self, event):
        """当用户从下拉列表中选择一个新字体时调用"""
        new_font = self.font_var.get()
        if new_font and new_font != self.settings.get("app_font", "Microsoft YaHei"):
            self.settings["app_font"] = new_font
            self.save_settings()
            self.log(f"字体已更改为 '{new_font}'。")
            self._apply_global_font()
            messagebox.showinfo("设置已保存", "字体设置已保存。\n请重启软件以使新字体完全生效。")

    def _restore_default_font(self):
        """恢复默认字体"""
        default_font = "Microsoft YaHei"
        if self.settings.get("app_font") != default_font:
            self.settings["app_font"] = default_font
            self.save_settings()
            self.font_var.set(default_font) # 更新UI显示
            self.log("字体已恢复为默认。")
            self._apply_global_font()
            messagebox.showinfo("设置已保存", "字体已恢复为默认设置。\n请重启软件以生效。")
        else:
            messagebox.showinfo("提示", "当前已是默认字体，无需恢复。")

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
            messagebox.showinfo("提示", "当前没有节目，无需操作。")
            return

        if messagebox.askyesno("确认操作", "您确定要取消所有节目中已设置的背景图片吗？\n此操作将取消所有任务的背景图片勾选。"):
            count = 0
            for task in self.tasks:
                if task.get('bg_image_enabled'):
                    task['bg_image_enabled'] = 0
                    count += 1

            if count > 0:
                self.save_tasks()
                self.log(f"已成功取消 {count} 个节目的背景图片设置。")
                messagebox.showinfo("操作成功", f"已成功取消 {count} 个节目的背景图片设置。")
            else:
                messagebox.showinfo("提示", "没有节目设置了背景图片，无需操作。")

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
            if messagebox.askyesno("应用更改", "您更改了报时参数，需要重新生成全部24个报时文件。\n是否立即开始？"):
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
                messagebox.showwarning("操作失败", "请先从下拉列表中选择一个播音员。")
                if not force_regenerate: self.time_chime_enabled_var.set(False)
                return

            self.save_settings()
            self.log("准备启用/更新整点报时功能，开始生成语音文件...")

            progress_dialog = ttk.Toplevel(self.root)
            progress_dialog.title("请稍候")
            progress_dialog.geometry("350x120")
            progress_dialog.resizable(False, False)
            progress_dialog.transient(self.root); progress_dialog.grab_set()
            self.center_window(progress_dialog, 350, 120)

            ttk.Label(progress_dialog, text="正在生成整点报时文件 (0/24)...", font=self.font_11).pack(pady=10)
            progress_label = ttk.Label(progress_dialog, text="", font=self.font_10)
            progress_label.pack(pady=5)

            threading.Thread(target=self._generate_chime_files_worker,
                             args=(selected_voice, progress_dialog, progress_label), daemon=True).start()

        elif not is_enabled and not force_regenerate:
            if messagebox.askyesno("确认操作", "您确定要禁用整点报时功能吗？\n这将删除所有已生成的报时音频文件。"):
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
            self.root.after(0, messagebox.showerror, "错误", f"生成报时文件失败：{e}")
        finally:
            self.root.after(0, progress_dialog.destroy)
            if success:
                self.log("全部整点报时文件生成完毕。")
                if self.time_chime_enabled_var.get():
                     self.root.after(0, messagebox.showinfo, "成功", "整点报时功能已启用/更新！")
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
            self.root.after(0, messagebox.showerror, "错误", f"删除报时文件失败：{e}")

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
        # 修复4: 正确地显示解锁按钮
        self.statusbar_unlock_button.pack(side=RIGHT, padx=5, before=self.status_labels[0])
        self.log("界面已锁定。")

    def _apply_unlock(self):
        self.is_locked = False
        self.lock_button.config(text="锁定", bootstyle='danger')
        self._set_ui_lock_state(NORMAL)
        # 修复4: 正确地隐藏解锁按钮
        self.statusbar_unlock_button.pack_forget()
        self.log("界面已解锁。")

    def perform_initial_lock(self):
        self.log("根据设置，软件启动时自动锁定。")
        self._apply_lock()

    def _prompt_for_password_set(self):
        dialog = ttk.Toplevel(self.root)
        dialog.title("首次锁定，请设置密码")
        dialog.geometry("350x250"); dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()
        self.center_window(dialog, 350, 250)

        ttk.Label(dialog, text="请设置一个锁定密码 (最多6位)", font=self.font_11).pack(pady=10)

        ttk.Label(dialog, text="输入密码:", font=self.font_11).pack(pady=(5,0))
        pass_entry1 = ttk.Entry(dialog, show='*', width=25, font=self.font_11)
        pass_entry1.pack()

        ttk.Label(dialog, text="确认密码:", font=self.font_11).pack(pady=(10,0))
        pass_entry2 = ttk.Entry(dialog, show='*', width=25, font=self.font_11)
        pass_entry2.pack()

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
                dialog.destroy()
                self._apply_lock()
            else:
                messagebox.showerror("功能受限", "无法保存密码。\n此功能仅在Windows系统上支持且需要pywin32库。", parent=dialog)

        btn_frame = ttk.Frame(dialog); btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=LEFT, padx=10)

    def _prompt_for_password_unlock(self):
        dialog = ttk.Toplevel(self.root)
        dialog.title("解锁界面")
        # 修复3: 增加解锁对话框的高度
        dialog.geometry("400x200"); dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()
        self.center_window(dialog, 400, 200)

        ttk.Label(dialog, text="请输入密码以解锁", font=self.font_11).pack(pady=10)

        pass_entry = ttk.Entry(dialog, show='*', width=25, font=self.font_11)
        pass_entry.pack(pady=5)
        pass_entry.focus_set()

        def is_password_correct():
            entered_pass = pass_entry.get()
            encoded_entered_pass = base64.b64encode(entered_pass.encode('utf-8')).decode('utf-8')
            return encoded_entered_pass == self.lock_password_b64

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
                self._perform_password_clear_logic()
                dialog.destroy()
                self.root.after(50, self._apply_unlock)
                self.root.after(100, lambda: messagebox.showinfo("成功", "锁定密码已成功清除。", parent=self.root))

        btn_frame = ttk.Frame(dialog); btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="清除密码", command=clear_password_action).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=LEFT, padx=5)
        dialog.bind('<Return>', lambda event: confirm())
#标记
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

    def _handle_lock_on_start_toggle(self):
        if not self.lock_password_b64:
            if self.lock_on_start_var.get():
                messagebox.showwarning("无法启用", "您还未设置锁定密码。\n\n请返回“定时广播”页面，点击“锁定”按钮来首次设置密码。")
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
        for child in parent_widget.winfo_children():
            if child == self.lock_button:
                continue

            try:
                child.config(state=state)
            except tk.TclError:
                pass

            if child.winfo_children():
                self._set_widget_state_recursively(child, state)

    def clear_log(self):
        if messagebox.askyesno("确认操作", "您确定要清空所有日志记录吗？\n此操作不可恢复。"):
            self.log_text.config(state='normal')
            self.log_text.delete('1.0', END)
            self.log_text.config(state='disabled')
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
            messagebox.showwarning("提示", "请先选择一个要立即播放的节目。")
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
        choice_dialog.title("选择节目类型")
        # 修复5: 增加添加节目对话框高度
        choice_dialog.geometry("350x400")
        choice_dialog.resizable(False, False)
        choice_dialog.transient(self.root); choice_dialog.grab_set()
        self.center_window(choice_dialog, 350, 400)
        main_frame = ttk.Frame(choice_dialog, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        title_label = ttk.Label(main_frame, text="请选择要添加的节目类型",
                              font=self.font_13_bold, bootstyle="primary")
        title_label.pack(pady=15)
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(expand=True)

        audio_btn = ttk.Button(btn_frame, text="🎵 音频节目",
                             bootstyle="primary", width=15, command=lambda: self.open_audio_dialog(choice_dialog))
        audio_btn.pack(pady=8, ipady=8)

        voice_btn = ttk.Button(btn_frame, text="🎙️ 语音节目",
                             bootstyle="info", width=15, command=lambda: self.open_voice_dialog(choice_dialog))
        voice_btn.pack(pady=8, ipady=8)

        video_btn = ttk.Button(btn_frame, text="🎬 视频节目",
                             bootstyle="success", width=15, command=lambda: self.open_video_dialog(choice_dialog))
        video_btn.pack(pady=8, ipady=8)
        if not VLC_AVAILABLE:
            video_btn.config(state=DISABLED, text="🎬 视频节目 (VLC未安装)")

    def open_audio_dialog(self, parent_dialog, task_to_edit=None, index=None):
        parent_dialog.destroy()
        is_edit_mode = task_to_edit is not None
        dialog = ttk.Toplevel(self.root)
        dialog.title("修改音频节目" if is_edit_mode else "添加音频节目")
        # 修复6: 移除固定尺寸，让窗口自适应
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()

        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=BOTH, expand=True)

        content_frame = ttk.LabelFrame(main_frame, text="内容", padding=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=2)

        ttk.Label(content_frame, text="节目名称:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        name_entry = ttk.Entry(content_frame, font=self.font_11, width=55)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        audio_type_var = tk.StringVar(value="single")
        ttk.Label(content_frame, text="音频文件").grid(row=1, column=0, sticky='e', padx=5, pady=2)
        audio_single_frame = ttk.Frame(content_frame)
        audio_single_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        ttk.Radiobutton(audio_single_frame, text="", variable=audio_type_var, value="single").pack(side=LEFT)
        audio_single_entry = ttk.Entry(audio_single_frame, font=self.font_11, width=35)
        audio_single_entry.pack(side=LEFT, padx=5)
        ttk.Label(audio_single_frame, text="00:00").pack(side=LEFT, padx=10)
        def select_single_audio():
            filename = filedialog.askopenfilename(title="选择音频文件", initialdir=AUDIO_FOLDER, filetypes=[("音频文件", "*.mp3 *.wav *.ogg *.flac *.m4a"), ("所有文件", "*.*")])
            if filename: audio_single_entry.delete(0, END); audio_single_entry.insert(0, filename)
        ttk.Button(audio_single_frame, text="选取...", command=select_single_audio, bootstyle="outline").pack(side=LEFT, padx=5)
        ttk.Label(content_frame, text="音频文件夹").grid(row=2, column=0, sticky='e', padx=5, pady=2)
        audio_folder_frame = ttk.Frame(content_frame)
        audio_folder_frame.grid(row=2, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        ttk.Radiobutton(audio_folder_frame, text="", variable=audio_type_var, value="folder").pack(side=LEFT)
        audio_folder_entry = ttk.Entry(audio_folder_frame, font=self.font_11, width=50)
        audio_folder_entry.pack(side=LEFT, padx=5)
        def select_folder(entry_widget):
            foldername = filedialog.askdirectory(title="选择文件夹", initialdir=application_path)
            if foldername: entry_widget.delete(0, END); entry_widget.insert(0, foldername)
        ttk.Button(audio_folder_frame, text="选取...", command=lambda: select_folder(audio_folder_entry), bootstyle="outline").pack(side=LEFT, padx=5)
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

        bg_image_cb = ttk.Checkbutton(bg_image_frame, text="背景图片:", variable=bg_image_var, bootstyle="round-toggle")
        bg_image_cb.pack(side=LEFT)
        if not IMAGE_AVAILABLE: bg_image_cb.config(state=DISABLED, text="背景图片(Pillow未安装):")

        bg_image_entry = ttk.Entry(bg_image_frame, textvariable=bg_image_path_var, font=self.font_11, width=42)
        bg_image_entry.pack(side=LEFT, padx=(0, 5))

        ttk.Button(bg_image_frame, text="选取...", command=lambda: select_folder(bg_image_entry), bootstyle="outline").pack(side=LEFT, padx=5)

        ttk.Radiobutton(bg_image_frame, text="顺序", variable=bg_image_order_var, value="sequential").pack(side=LEFT, padx=(10,0))
        ttk.Radiobutton(bg_image_frame, text="随机", variable=bg_image_order_var, value="random").pack(side=LEFT)

        volume_frame = ttk.Frame(content_frame)
        volume_frame.grid(row=5, column=1, columnspan=3, sticky='w', padx=5, pady=3)
        ttk.Label(volume_frame, text="音量:").pack(side=LEFT)
        volume_entry = ttk.Entry(volume_frame, font=self.font_11, width=10)
        volume_entry.pack(side=LEFT, padx=5)
        ttk.Label(volume_frame, text="0-100").pack(side=LEFT, padx=5)

        time_frame = ttk.LabelFrame(main_frame, text="时间", padding=15)
        time_frame.grid(row=1, column=0, sticky='ew', pady=4)
        ttk.Label(time_frame, text="开始时间:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        start_time_entry = ttk.Entry(time_frame, font=self.font_11, width=50)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        self._bind_mousewheel_to_entry(start_time_entry, self._handle_time_scroll)
        ttk.Label(time_frame, text="《可多个,用英文逗号,隔开》").grid(row=0, column=2, sticky='w', padx=5)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bootstyle="outline").grid(row=0, column=3, padx=5)
        interval_var = tk.StringVar(value="first")
        interval_frame1 = ttk.Frame(time_frame)
        interval_frame1.grid(row=1, column=1, columnspan=2, sticky='w', padx=5, pady=2)
        ttk.Label(time_frame, text="间隔播报:").grid(row=1, column=0, sticky='e', padx=5, pady=2)
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
        weekday_entry = ttk.Entry(time_frame, font=self.font_11, width=50)
        weekday_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=3)
        ttk.Button(time_frame, text="选取...", command=lambda: self.show_weekday_settings_dialog(weekday_entry), bootstyle="outline").grid(row=3, column=3, padx=5)
        ttk.Label(time_frame, text="日期范围:").grid(row=4, column=0, sticky='e', padx=5, pady=3)
        date_range_entry = ttk.Entry(time_frame, font=self.font_11, width=50)
        date_range_entry.grid(row=4, column=1, sticky='ew', padx=5, pady=3)
        self._bind_mousewheel_to_entry(date_range_entry, self._handle_date_scroll)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry), bootstyle="outline").grid(row=4, column=3, padx=5)

        other_frame = ttk.LabelFrame(main_frame, text="其它", padding=10)
        other_frame.grid(row=2, column=0, sticky='ew', pady=5)
        delay_var = tk.StringVar(value="ontime")
        ttk.Label(other_frame, text="模式:").grid(row=0, column=0, sticky='nw', padx=5, pady=2)
        delay_frame = ttk.Frame(other_frame)
        delay_frame.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        ttk.Radiobutton(delay_frame, text="准时播 - 如果有别的节目正在播，终止他们（默认）", variable=delay_var, value="ontime").pack(anchor='w')
        ttk.Radiobutton(delay_frame, text="可延后 - 如果有别的节目正在播，排队等候", variable=delay_var, value="delay").pack(anchor='w')
        ttk.Radiobutton(delay_frame, text="立即播 - 添加后停止其他节目,立即播放此节目", variable=delay_var, value="immediate").pack(anchor='w')
        dialog_button_frame = ttk.Frame(other_frame)
        dialog_button_frame.grid(row=0, column=2, sticky='e', padx=20)
        other_frame.grid_columnconfigure(1, weight=1)

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

            self.update_task_list(); self.save_tasks(); dialog.destroy()

            if play_this_task_now:
                self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))

        button_text = "保存修改" if is_edit_mode else "添加"
        ttk.Button(dialog_button_frame, text=button_text, command=save_task, bootstyle="primary").pack(side=LEFT, padx=10, ipady=5)
        ttk.Button(dialog_button_frame, text="取消", command=dialog.destroy).pack(side=LEFT, padx=10, ipady=5)

        content_frame.columnconfigure(1, weight=1); time_frame.columnconfigure(1, weight=1)
        
        dialog.update_idletasks()
        self.center_window(dialog, dialog.winfo_width(), dialog.winfo_height())

    def open_video_dialog(self, parent_dialog, task_to_edit=None, index=None):
        parent_dialog.destroy()
        is_edit_mode = task_to_edit is not None
        dialog = ttk.Toplevel(self.root)
        dialog.title("修改视频节目" if is_edit_mode else "添加视频节目")
        # 修复6: 移除固定尺寸
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=BOTH, expand=True)

        content_frame = ttk.LabelFrame(main_frame, text="内容", padding=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=2)

        ttk.Label(content_frame, text="节目名称:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        name_entry = ttk.Entry(content_frame, font=self.font_11, width=55)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=2)

        video_type_var = tk.StringVar(value="single")

        ttk.Label(content_frame, text="视频文件:").grid(row=1, column=0, sticky='e', padx=5, pady=2)
        video_single_frame = ttk.Frame(content_frame)
        video_single_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        ttk.Radiobutton(video_single_frame, text="", variable=video_type_var, value="single").pack(side=LEFT)
        video_single_entry = ttk.Entry(video_single_frame, font=self.font_11, width=50)
        video_single_entry.pack(side=LEFT, padx=5)

        def select_single_video():
            ftypes = [("视频文件", "*.mp4 *.mkv *.avi *.mov *.wmv *.flv"), ("所有文件", "*.*")]
            filename = filedialog.askopenfilename(title="选择视频文件", filetypes=ftypes)
            if filename:
                video_single_entry.delete(0, END)
                video_single_entry.insert(0, filename)
        ttk.Button(video_single_frame, text="选取...", command=select_single_video, bootstyle="outline").pack(side=LEFT, padx=5)

        ttk.Label(content_frame, text="视频文件夹:").grid(row=2, column=0, sticky='e', padx=5, pady=2)
        video_folder_frame = ttk.Frame(content_frame)
        video_folder_frame.grid(row=2, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        ttk.Radiobutton(video_folder_frame, text="", variable=video_type_var, value="folder").pack(side=LEFT)
        video_folder_entry = ttk.Entry(video_folder_frame, font=self.font_11, width=50)
        video_folder_entry.pack(side=LEFT, padx=5)

        def select_folder(entry_widget):
            foldername = filedialog.askdirectory(title="选择文件夹", initialdir=application_path)
            if foldername:
                entry_widget.delete(0, END)
                entry_widget.insert(0, foldername)
        ttk.Button(video_folder_frame, text="选取...", command=lambda: select_folder(video_folder_entry), bootstyle="outline").pack(side=LEFT, padx=5)

        play_order_frame = ttk.Frame(content_frame)
        play_order_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        play_order_var = tk.StringVar(value="sequential")
        ttk.Radiobutton(play_order_frame, text="顺序播", variable=play_order_var, value="sequential").pack(side=LEFT, padx=10)
        ttk.Radiobutton(play_order_frame, text="随机播", variable=play_order_var, value="random").pack(side=LEFT, padx=10)

        volume_frame = ttk.Frame(content_frame)
        volume_frame.grid(row=4, column=1, columnspan=3, sticky='w', padx=5, pady=3)
        ttk.Label(volume_frame, text="音量:").pack(side=LEFT)
        volume_entry = ttk.Entry(volume_frame, font=self.font_11, width=10)
        volume_entry.pack(side=LEFT, padx=5)
        ttk.Label(volume_frame, text="0-100").pack(side=LEFT, padx=5)

        content_frame.columnconfigure(1, weight=1)

        playback_frame = ttk.LabelFrame(main_frame, text="播放选项", padding=10)
        playback_frame.grid(row=1, column=0, sticky='ew', pady=4)

        playback_mode_var = tk.StringVar(value="fullscreen")
        resolutions = ["640x480", "800x600", "1024x768", "1280x720", "1366x768", "1600x900", "1920x1080"]
        resolution_var = tk.StringVar(value=resolutions[2])

        playback_rates = ['0.5x', '0.75x', '1.0x (正常)', '1.25x', '1.5x', '2.0x']
        playback_rate_var = tk.StringVar(value='1.0x (正常)')

        mode_frame = ttk.Frame(playback_frame)
        mode_frame.grid(row=0, column=0, columnspan=3, sticky='w')

        resolution_combo = ttk.Combobox(mode_frame, textvariable=resolution_var, values=resolutions, font=self.font_11, width=15, state='readonly')

        def toggle_resolution_combo():
            if playback_mode_var.get() == "windowed":
                resolution_combo.config(state='readonly')
            else:
                resolution_combo.config(state='disabled')

        ttk.Radiobutton(mode_frame, text="无边框全屏", variable=playback_mode_var, value="fullscreen", command=toggle_resolution_combo).pack(side=LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="非全屏", variable=playback_mode_var, value="windowed", command=toggle_resolution_combo).pack(side=LEFT, padx=5)
        resolution_combo.pack(side=LEFT, padx=10)

        rate_frame = ttk.Frame(playback_frame)
        rate_frame.grid(row=1, column=0, columnspan=3, sticky='w', pady=5)
        ttk.Label(rate_frame, text="播放倍速:").pack(side=LEFT, padx=5)
        rate_combo = ttk.Combobox(rate_frame, textvariable=playback_rate_var, values=playback_rates, font=self.font_11, width=15)
        rate_combo.pack(side=LEFT)
        ttk.Label(rate_frame, text="(可手动输入0.25-4.0之间的值)", font=self.font_9, bootstyle="secondary").pack(side=LEFT, padx=5)

        toggle_resolution_combo()

        time_frame = ttk.LabelFrame(main_frame, text="时间", padding=15)
        time_frame.grid(row=2, column=0, sticky='ew', pady=4)

        ttk.Label(time_frame, text="开始时间:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        start_time_entry = ttk.Entry(time_frame, font=self.font_11, width=50)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        self._bind_mousewheel_to_entry(start_time_entry, self._handle_time_scroll)
        ttk.Label(time_frame, text="《可多个,用英文逗号,隔开》").grid(row=0, column=2, sticky='w', padx=5)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bootstyle="outline").grid(row=0, column=3, padx=5)

        interval_var = tk.StringVar(value="first")
        interval_frame1 = ttk.Frame(time_frame)
        interval_frame1.grid(row=1, column=1, columnspan=2, sticky='w', padx=5, pady=2)
        ttk.Label(time_frame, text="间隔播报:").grid(row=1, column=0, sticky='e', padx=5, pady=2)
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
        weekday_entry = ttk.Entry(time_frame, font=self.font_11, width=50)
        weekday_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=3)
        ttk.Button(time_frame, text="选取...", command=lambda: self.show_weekday_settings_dialog(weekday_entry), bootstyle="outline").grid(row=3, column=3, padx=5)

        ttk.Label(time_frame, text="日期范围:").grid(row=4, column=0, sticky='e', padx=5, pady=3)
        date_range_entry = ttk.Entry(time_frame, font=self.font_11, width=50)
        date_range_entry.grid(row=4, column=1, sticky='ew', padx=5, pady=3)
        self._bind_mousewheel_to_entry(date_range_entry, self._handle_date_scroll)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry), bootstyle="outline").grid(row=4, column=3, padx=5)

        time_frame.columnconfigure(1, weight=1)

        other_frame = ttk.LabelFrame(main_frame, text="其它", padding=10)
        other_frame.grid(row=3, column=0, sticky='ew', pady=5)

        delay_var = tk.StringVar(value="ontime")
        ttk.Label(other_frame, text="模式:").grid(row=0, column=0, sticky='nw', padx=5, pady=2)
        delay_frame = ttk.Frame(other_frame)
        delay_frame.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        ttk.Radiobutton(delay_frame, text="准时播 - 如果有别的节目正在播，终止他们（默认）", variable=delay_var, value="ontime").pack(anchor='w')
        ttk.Radiobutton(delay_frame, text="可延后 - 如果有别的节目正在播，排队等候", variable=delay_var, value="delay").pack(anchor='w')
        ttk.Radiobutton(delay_frame, text="立即播 - 添加后停止其他节目,立即播放此节目", variable=delay_var, value="immediate").pack(anchor='w')

        dialog_button_frame = ttk.Frame(other_frame)
        dialog_button_frame.grid(row=0, column=2, sticky='e', padx=20)
        other_frame.grid_columnconfigure(1, weight=1)

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
            dialog.destroy()

            if play_this_task_now:
                self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))

        button_text = "保存修改" if is_edit_mode else "添加"
        ttk.Button(dialog_button_frame, text=button_text, command=save_task, bootstyle="primary").pack(side=LEFT, padx=10, ipady=5)
        ttk.Button(dialog_button_frame, text="取消", command=dialog.destroy).pack(side=LEFT, padx=10, ipady=5)
        
        dialog.update_idletasks()
        self.center_window(dialog, dialog.winfo_width(), dialog.winfo_height())
#标记
    def open_voice_dialog(self, parent_dialog, task_to_edit=None, index=None):
        parent_dialog.destroy()
        is_edit_mode = task_to_edit is not None
        dialog = ttk.Toplevel(self.root)
        dialog.title("修改语音节目" if is_edit_mode else "添加语音节目")
        # 修复6: 移除固定尺寸
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()

        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=BOTH, expand=True)

        content_frame = ttk.LabelFrame(main_frame, text="内容", padding=10)
        content_frame.grid(row=0, column=0, sticky='ew', pady=2)

        ttk.Label(content_frame, text="节目名称:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        name_entry = ttk.Entry(content_frame, font=self.font_11, width=65)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        ttk.Label(content_frame, text="播音文字:").grid(row=1, column=0, sticky='nw', padx=5, pady=2)
        text_frame = ttk.Frame(content_frame)
        text_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        content_text = ScrolledText(text_frame, height=5, font=self.font_11, width=65, wrap=WORD, autohide=True)
        content_text.pack(fill=BOTH, expand=True)
        script_btn_frame = ttk.Frame(content_frame)
        script_btn_frame.grid(row=2, column=1, columnspan=3, sticky='w', padx=5, pady=(0, 2))
        ttk.Button(script_btn_frame, text="导入文稿", command=lambda: self._import_voice_script(content_text), bootstyle="outline").pack(side=LEFT)
        ttk.Button(script_btn_frame, text="导出文稿", command=lambda: self._export_voice_script(content_text, name_entry), bootstyle="outline").pack(side=LEFT, padx=10)
        ttk.Label(content_frame, text="播音员:").grid(row=3, column=0, sticky='w', padx=5, pady=3)
        voice_frame = ttk.Frame(content_frame)
        voice_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=3)
        available_voices = self.get_available_voices()
        voice_var = tk.StringVar()
        # 修复8: 增加 Combobox 宽度
        voice_combo = ttk.Combobox(voice_frame, textvariable=voice_var, values=available_voices, font=self.font_11, width=60, state='readonly')
        voice_combo.pack(side=LEFT)
        speech_params_frame = ttk.Frame(content_frame)
        speech_params_frame.grid(row=4, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        ttk.Label(speech_params_frame, text="语速(-10~10):").pack(side=LEFT, padx=(0,5))
        speed_entry = ttk.Entry(speech_params_frame, font=self.font_11, width=8); speed_entry.pack(side=LEFT, padx=5)
        ttk.Label(speech_params_frame, text="音调(-10~10):").pack(side=LEFT, padx=(10,5))
        pitch_entry = ttk.Entry(speech_params_frame, font=self.font_11, width=8); pitch_entry.pack(side=LEFT, padx=5)
        ttk.Label(speech_params_frame, text="音量(0-100):").pack(side=LEFT, padx=(10,5))
        volume_entry = ttk.Entry(speech_params_frame, font=self.font_11, width=8); volume_entry.pack(side=LEFT, padx=5)
        prompt_var = tk.IntVar(); prompt_frame = ttk.Frame(content_frame)
        prompt_frame.grid(row=5, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        ttk.Checkbutton(prompt_frame, text="提示音:", variable=prompt_var, bootstyle="round-toggle").pack(side=LEFT)
        prompt_file_var, prompt_volume_var = tk.StringVar(), tk.StringVar()
        prompt_file_entry = ttk.Entry(prompt_frame, textvariable=prompt_file_var, font=self.font_11, width=20); prompt_file_entry.pack(side=LEFT, padx=5)
        ttk.Button(prompt_frame, text="...", command=lambda: self.select_file_for_entry(PROMPT_FOLDER, prompt_file_var), bootstyle="outline", width=2).pack(side=LEFT)
        ttk.Label(prompt_frame, text="音量(0-100):").pack(side=LEFT, padx=(10,5))
        ttk.Entry(prompt_frame, textvariable=prompt_volume_var, font=self.font_11, width=8).pack(side=LEFT, padx=5)
        bgm_var = tk.IntVar(); bgm_frame = ttk.Frame(content_frame)
        bgm_frame.grid(row=6, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        ttk.Checkbutton(bgm_frame, text="背景音乐:", variable=bgm_var, bootstyle="round-toggle").pack(side=LEFT)
        bgm_file_var, bgm_volume_var = tk.StringVar(), tk.StringVar()
        bgm_file_entry = ttk.Entry(bgm_frame, textvariable=bgm_file_var, font=self.font_11, width=20); bgm_file_entry.pack(side=LEFT, padx=5)
        ttk.Button(bgm_frame, text="...", command=lambda: self.select_file_for_entry(BGM_FOLDER, bgm_file_var), bootstyle="outline", width=2).pack(side=LEFT)
        ttk.Label(bgm_frame, text="音量(0-100):").pack(side=LEFT, padx=(10,5))
        ttk.Entry(bgm_frame, textvariable=bgm_volume_var, font=self.font_11, width=8).pack(side=LEFT, padx=5)

        bg_image_var = tk.IntVar(value=0)
        bg_image_path_var = tk.StringVar()
        bg_image_order_var = tk.StringVar(value="sequential")

        bg_image_frame = ttk.Frame(content_frame)
        bg_image_frame.grid(row=7, column=1, columnspan=3, sticky='w', padx=5, pady=5)

        bg_image_cb = ttk.Checkbutton(bg_image_frame, text="背景图片:", variable=bg_image_var, bootstyle="round-toggle")
        bg_image_cb.pack(side=LEFT)
        if not IMAGE_AVAILABLE: bg_image_cb.config(state=DISABLED, text="背景图片(Pillow未安装):")

        bg_image_entry = ttk.Entry(bg_image_frame, textvariable=bg_image_path_var, font=self.font_11, width=32)
        bg_image_entry.pack(side=LEFT, padx=(0, 5))

        ttk.Button(bg_image_frame, text="选取...", command=lambda: select_folder(bg_image_entry), bootstyle="outline").pack(side=LEFT, padx=5)

        ttk.Radiobutton(bg_image_frame, text="顺序", variable=bg_image_order_var, value="sequential").pack(side=LEFT, padx=(10,0))
        ttk.Radiobutton(bg_image_frame, text="随机", variable=bg_image_order_var, value="random").pack(side=LEFT)

        time_frame = ttk.LabelFrame(main_frame, text="时间", padding=10)
        time_frame.grid(row=1, column=0, sticky='ew', pady=2)
        ttk.Label(time_frame, text="开始时间:").grid(row=0, column=0, sticky='e', padx=5, pady=2)
        start_time_entry = ttk.Entry(time_frame, font=self.font_11, width=50)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        self._bind_mousewheel_to_entry(start_time_entry, self._handle_time_scroll)
        ttk.Label(time_frame, text="《可多个,用英文逗号,隔开》").grid(row=0, column=2, sticky='w', padx=5)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bootstyle="outline").grid(row=0, column=3, padx=5)
        ttk.Label(time_frame, text="播 n 遍:").grid(row=1, column=0, sticky='e', padx=5, pady=2)
        repeat_entry = ttk.Entry(time_frame, font=self.font_11, width=12)
        repeat_entry.grid(row=1, column=1, sticky='w', padx=5, pady=2)
        ttk.Label(time_frame, text="周几/几号:").grid(row=2, column=0, sticky='e', padx=5, pady=2)
        weekday_entry = ttk.Entry(time_frame, font=self.font_11, width=50)
        weekday_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=2)
        ttk.Button(time_frame, text="选取...", command=lambda: self.show_weekday_settings_dialog(weekday_entry), bootstyle="outline").grid(row=2, column=3, padx=5)
        ttk.Label(time_frame, text="日期范围:").grid(row=3, column=0, sticky='e', padx=5, pady=2)
        date_range_entry = ttk.Entry(time_frame, font=self.font_11, width=50)
        date_range_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=2)
        self._bind_mousewheel_to_entry(date_range_entry, self._handle_date_scroll)
        ttk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry), bootstyle="outline").grid(row=3, column=3, padx=5)

        other_frame = ttk.LabelFrame(main_frame, text="其它", padding=15)
        other_frame.grid(row=2, column=0, sticky='ew', pady=4)
        delay_var = tk.StringVar(value="delay")
        ttk.Label(other_frame, text="模式:").grid(row=0, column=0, sticky='nw', padx=5, pady=2)
        delay_frame = ttk.Frame(other_frame)
        delay_frame.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        ttk.Radiobutton(delay_frame, text="准时播 - 如果有别的节目正在播，终止他们", variable=delay_var, value="ontime").pack(anchor='w', pady=1)
        ttk.Radiobutton(delay_frame, text="可延后 - 如果有别的节目正在播，排队等候（默认）", variable=delay_var, value="delay").pack(anchor='w', pady=1)
        ttk.Radiobutton(delay_frame, text="立即播 - 添加后停止其他节目,立即播放此节目", variable=delay_var, value="immediate").pack(anchor='w', pady=1)
        dialog_button_frame = ttk.Frame(other_frame)
        dialog_button_frame.grid(row=0, column=2, sticky='e', padx=20)
        other_frame.grid_columnconfigure(1, weight=1)

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

        def save_task():
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
                self.update_task_list(); self.save_tasks(); dialog.destroy()
                if play_now_flag: self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))
                return

            progress_dialog = ttk.Toplevel(dialog); progress_dialog.title("请稍候"); progress_dialog.geometry("300x100")
            progress_dialog.resizable(False, False); progress_dialog.transient(dialog); progress_dialog.grab_set()
            ttk.Label(progress_dialog, text="语音文件生成中，请稍后...", font=self.font_11).pack(expand=True)
            self.center_window(progress_dialog, 300, 100); dialog.update_idletasks()
            new_wav_filename = f"{int(time.time())}_{random.randint(1000, 9999)}.wav"
            output_path = os.path.join(AUDIO_FOLDER, new_wav_filename)
            voice_params = {'voice': voice_var.get(), 'speed': speed_entry.get().strip() or "0", 'pitch': pitch_entry.get().strip() or "0", 'volume': volume_entry.get().strip() or "80"}
            def _on_synthesis_complete(result):
                progress_dialog.destroy()
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
                self.update_task_list(); self.save_tasks(); dialog.destroy()
                if play_now_flag: self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))
            synthesis_thread = threading.Thread(target=self._synthesis_worker, args=(text_content, voice_params, output_path, _on_synthesis_complete))
            synthesis_thread.daemon = True; synthesis_thread.start()

        button_text = "保存修改" if is_edit_mode else "添加"
        ttk.Button(dialog_button_frame, text=button_text, command=save_task, bootstyle="primary").pack(side=LEFT, padx=10, ipady=5)
        ttk.Button(dialog_button_frame, text="取消", command=dialog.destroy).pack(side=LEFT, padx=10, ipady=5)

        content_frame.columnconfigure(1, weight=1); time_frame.columnconfigure(1, weight=1)
        
        dialog.update_idletasks()
        self.center_window(dialog, dialog.winfo_width(), dialog.winfo_height())

    def _import_voice_script(self, text_widget):
        filename = filedialog.askopenfilename(
            title="选择要导入的文稿",
            initialdir=VOICE_SCRIPT_FOLDER,
            filetypes=[("文本文档", "*.txt"), ("所有文件", "*.*")]
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
            messagebox.showerror("导入失败", f"无法读取文件：\n{e}")
            self.log(f"导入文稿失败: {e}")

    def _export_voice_script(self, text_widget, name_widget):
        content = text_widget.get('1.0', END).strip()
        if not content:
            messagebox.showwarning("无法导出", "播音文字内容为空，无需导出。")
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
            filetypes=[("文本文档", "*.txt")]
        )
        if not filename:
            return

        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(content)
            self.log(f"文稿已成功导出到 {os.path.basename(filename)}。")
            messagebox.showinfo("导出成功", f"文稿已成功导出到：\n{filename}")
        except Exception as e:
            messagebox.showerror("导出失败", f"无法保存文件：\n{e}")
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
        filename = filedialog.askopenfilename(title="选择文件", initialdir=initial_dir, filetypes=[("音频文件", "*.mp3 *.wav *.ogg *.flac *.m4a"), ("所有文件", "*.*")])
        if filename: string_var.set(os.path.basename(filename))

    def delete_task(self):
        selections = self.task_tree.selection()
        if not selections: messagebox.showwarning("警告", "请先选择要删除的节目"); return
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(selections)} 个节目吗？\n(关联的语音文件也将被删除)"):
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
        dummy_parent = ttk.Toplevel(self.root); dummy_parent.withdraw()

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

                if not isinstance(imported, list) or \
                   (imported and (not isinstance(imported[0], dict) or 'time' not in imported[0] or 'type' not in imported[0])):
                    messagebox.showerror("导入失败", "文件格式不正确，看起来不是一个有效的节目单备份文件。")
                    self.log(f"尝试导入格式错误的节目单文件: {os.path.basename(filename)}")
                    return

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
        dialog.geometry("350x150")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        self.center_window(dialog, 350, 150)

        result = [None]

        ttk.Label(dialog, text=prompt, font=self.font_11).pack(pady=10)
        entry = ttk.Entry(dialog, font=self.font_11, width=15, justify='center')
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

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=15)

        ttk.Button(btn_frame, text="确定", command=on_confirm, bootstyle="primary", width=8).pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=on_cancel, width=8).pack(side=LEFT, padx=10)

        dialog.bind('<Return>', lambda event: on_confirm())

        self.root.wait_window(dialog)
        return result[0]

    def clear_all_tasks(self, delete_associated_files=True):
        if not self.tasks: return

        if delete_associated_files:
            msg = "您确定要清空所有节目吗？\n此操作将同时删除关联的语音文件，且不可恢复！"
        else:
            msg = "您确定要清空所有节目列表吗？\n（此操作不会删除音频文件）"

        if messagebox.askyesno("严重警告", msg):
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
        # 修复7: 让窗口自适应
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()
        
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
            dialog.destroy()
        ttk.Button(bottom_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=5, ipady=5)
        ttk.Button(bottom_frame, text="取消", command=dialog.destroy).pack(side=LEFT, padx=5, ipady=5)

        dialog.update_idletasks()
        self.center_window(dialog, dialog.winfo_width(), dialog.winfo_height())


    def show_weekday_settings_dialog(self, weekday_entry):
        dialog = ttk.Toplevel(self.root); dialog.title("周几或几号")
        # 修复7: 让窗口自适应
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()
        
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
            dialog.destroy()
        ttk.Button(bottom_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=5, ipady=5)
        ttk.Button(bottom_frame, text="取消", command=dialog.destroy).pack(side=LEFT, padx=5, ipady=5)

        dialog.update_idletasks()
        self.center_window(dialog, dialog.winfo_width(), dialog.winfo_height())


    def show_daterange_settings_dialog(self, date_range_entry):
        dialog = ttk.Toplevel(self.root)
        dialog.title("日期范围")
        # 修复7: 让窗口自适应
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()
        
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
                dialog.destroy()
            else: messagebox.showerror("格式错误", "日期格式不正确, 应为 YYYY-MM-DD", parent=dialog)
        ttk.Button(bottom_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=5, ipady=5)
        ttk.Button(bottom_frame, text="取消", command=dialog.destroy).pack(side=LEFT, padx=5, ipady=5)

        dialog.update_idletasks()
        self.center_window(dialog, dialog.winfo_width(), dialog.winfo_height())


    def show_single_time_dialog(self, time_var):
        dialog = ttk.Toplevel(self.root)
        dialog.title("设置时间")
        # 修复7 & 9: 让窗口自适应
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()

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
                dialog.destroy()
            else: messagebox.showerror("格式错误", "请输入有效的时间格式 HH:MM:SS", parent=dialog)
        bottom_frame = ttk.Frame(main_frame); bottom_frame.pack(pady=10)
        ttk.Button(bottom_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=10)
        ttk.Button(bottom_frame, text="取消", command=dialog.destroy).pack(side=LEFT, padx=10)
        
        dialog.update_idletasks()
        self.center_window(dialog, dialog.winfo_width(), dialog.winfo_height())


    def show_power_week_time_dialog(self, title, days_var, time_var):
        dialog = ttk.Toplevel(self.root); dialog.title(title)
        # 修复7 & 9: 让窗口自适应
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()
        
        main_frame = ttk.Frame(dialog, padding=10)
        main_frame.pack(fill=BOTH, expand=True)

        week_frame = ttk.LabelFrame(main_frame, text="选择周几", padding=10)
        week_frame.pack(fill=X, pady=10, padx=10)
        weekdays = [("周一", 1), ("周二", 2), ("周三", 3), ("周四", 4), ("周五", 5), ("周六", 6), ("周日", 7)]
        week_vars = {num: tk.IntVar() for day, num in weekdays}
        current_days = days_var.get().replace("每周:", "")
        for day_num_str in current_days: week_vars[int(day_num_str)].set(1)
        for i, (day, num) in enumerate(weekdays): ttk.Checkbutton(week_frame, text=day, variable=week_vars[num]).grid(row=0, column=i, sticky='w', padx=10, pady=3)
        
        time_frame = ttk.LabelFrame(main_frame, text="设置时间", padding=10)
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
            dialog.destroy()
        bottom_frame = ttk.Frame(main_frame); bottom_frame.pack(pady=15)
        ttk.Button(bottom_frame, text="确定", command=confirm, bootstyle="primary").pack(side=LEFT, padx=10)
        ttk.Button(bottom_frame, text="取消", command=dialog.destroy).pack(side=LEFT, padx=10)
        
        dialog.update_idletasks()
        self.center_window(dialog, dialog.winfo_width(), dialog.winfo_height())


def main():
    # 修复：为更好的缩放效果设置DPI感知
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception as e:
        print(f"警告: 无法设置DPI感知 - {e}")

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
