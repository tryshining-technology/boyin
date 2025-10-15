import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog, simpledialog
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
# NEW: 导入 ctypes 用于调用 Windows API
import ctypes
from ctypes import wintypes

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

WIN32COM_AVAILABLE = False
try:
    import win32com.client
    import pythoncom
    from pywintypes import com_error
    import winreg
    WIN32COM_AVAILABLE = True
except ImportError:
    print("警告: pywin32 未安装，语音、开机启动和密码持久化/注册功能将受限。")

AUDIO_AVAILABLE = False
try:
    import pygame
    pygame.mixer.init()
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
# NEW: 默认提示音文件路径
NOTIFY_SOUND_FILE = os.path.join(PROMPT_FOLDER, "notify.wav")


CHIME_FOLDER = os.path.join(AUDIO_FOLDER, "整点报时")

REGISTRY_KEY_PATH = r"Software\创翔科技\TimedBroadcastApp"
REGISTRY_PARENT_KEY_PATH = r"Software\创翔科技"


class TimedBroadcastApp:
    def __init__(self, root):
        self.root = root
        self.root.title(" 创翔多功能定时播音旗舰版")
        self.root.geometry("1400x800")
        self.root.configure(bg='#E8F4F8')
        
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

        self.last_chime_hour = -1

        self.fullscreen_window = None
        self.fullscreen_label = None
        self.image_tk_ref = None 
        self.current_stop_visual_event = None

        self.notification_sound = None

        self.create_folder_structure()
        self.load_settings()
        self.load_lock_password()
        
        self.check_authorization()

        self.create_widgets()
        self.load_tasks()
        self.load_holidays()
        self.load_todos()
        
        if AUDIO_AVAILABLE and os.path.exists(NOTIFY_SOUND_FILE):
            try:
                self.notification_sound = pygame.mixer.Sound(NOTIFY_SOUND_FILE)
            except Exception as e:
                print(f"警告: 无法加载提示音文件 {NOTIFY_SOUND_FILE}: {e}")

        self.start_background_threads()
        self.root.protocol("WM_DELETE_WINDOW", self.show_quit_dialog)
        self.start_tray_icon_thread()
        
        if self.settings.get("lock_on_start", False) and self.lock_password_b64:
            self.root.after(100, self.perform_initial_lock)

        if self.settings.get("start_minimized", False):
            self.root.after(100, self.hide_to_tray)
        
        if self.is_app_locked_down:
            self.root.after(100, self.perform_lockdown)

    def _save_to_registry(self, key_name, value):
        if not WIN32COM_AVAILABLE: return False
        try:
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, REGISTRY_KEY_PATH)
            winreg.SetValueEx(key, key_name, 0, winreg.REG_SZ, str(value))
            winreg.CloseKey(key)
            return True
        except Exception as e:
            self.log(f"错误: 无法写入注册表项 '{key_name}' - {e}")
            return False

    def _load_from_registry(self, key_name):
        if not WIN32COM_AVAILABLE: return None
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
        self.status_frame = tk.Frame(self.root, bg='#E8F4F8', height=30)
        self.status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_frame.pack_propagate(False)
        self.create_status_bar_content()

        self.nav_frame = tk.Frame(self.root, bg='#A8D8E8', width=160)
        self.nav_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.nav_frame.pack_propagate(False)
        
        self.page_container = tk.Frame(self.root)
        self.page_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        nav_button_titles = ["定时广播", "节假日", "待办事项", "设置", "注册软件", "超级管理"]
        
        for i, title in enumerate(nav_button_titles):
            btn_frame = tk.Frame(self.nav_frame, bg='#A8D8E8')
            btn_frame.pack(fill=tk.X, pady=1)

            cmd = None
            if title == "超级管理":
                cmd = self._prompt_for_super_admin_password
            else:
                cmd = lambda t=title: self.switch_page(t)

            btn = tk.Button(btn_frame, text=title, bg='#A8D8E8',
                          fg='black', font=('Microsoft YaHei', 22, 'bold'),
                          bd=0, padx=10, pady=8, anchor='w', command=cmd)
            btn.pack(fill=tk.X)
            self.nav_buttons[title] = btn
        
        self.main_frame = tk.Frame(self.page_container, bg='white')
        self.pages["定时广播"] = self.main_frame
        self.create_scheduled_broadcast_page()

        self.current_page = self.main_frame
        self.switch_page("定时广播")
        
        self.update_status_bar()
        self.log(" 创翔多功能定时播音旗舰版软件已启动")

    def create_status_bar_content(self):
        self.status_labels = []
        status_texts = ["当前时间", "系统状态", "播放状态", "任务数量"]
        font_11 = ('Microsoft YaHei', 11)

        copyright_label = tk.Label(self.status_frame, text="© 创翔科技", font=font_11,
                                   bg='#5DADE2', fg='white', padx=15)
        copyright_label.pack(side=tk.RIGHT, padx=2)
        
        self.statusbar_unlock_button = tk.Button(self.status_frame, text="🔓 解锁", font=font_11,
                                                 bg='#2ECC71', fg='white', bd=0, padx=15, cursor='hand2',
                                                 command=self._prompt_for_password_unlock)

        for i, text in enumerate(status_texts):
            label = tk.Label(self.status_frame, text=f"{text}: --", font=font_11,
                           bg='#5DADE2' if i % 2 == 0 else '#7EC8E3', fg='white', padx=15, pady=5)
            label.pack(side=tk.LEFT, padx=2)
            self.status_labels.append(label)

    def switch_page(self, page_name):
        if self.is_app_locked_down and page_name not in ["注册软件", "超级管理"]:
            self.log("软件授权已过期，请先注册。")
            if self.current_page != self.pages.get("注册软件"):
                self.root.after(10, lambda: self.switch_page("注册软件"))
            return

        if self.is_locked and page_name not in ["超级管理", "注册软件"]:
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

        target_frame.pack(in_=self.page_container, fill=tk.BOTH, expand=True)
        self.current_page = target_frame
        
        selected_btn = self.nav_buttons[page_name]
        selected_btn.config(bg='#5DADE2', fg='white')
        selected_btn.master.config(bg='#5DADE2')

    def _prompt_for_super_admin_password(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("身份验证")
        dialog.geometry("350x180")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        self.center_window(dialog, 350, 180)

        result = [None] 

        tk.Label(dialog, text="请输入超级管理员密码:", font=('Microsoft YaHei', 11)).pack(pady=20)
        password_entry = tk.Entry(dialog, show='*', font=('Microsoft YaHei', 11), width=25)
        password_entry.pack(pady=5)
        password_entry.focus_set()

        def on_confirm():
            result[0] = password_entry.get()
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()

        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=20)
        tk.Button(btn_frame, text="确定", command=on_confirm, width=8).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="取消", command=on_cancel, width=8).pack(side=tk.LEFT, padx=10)
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
        page_frame = tk.Frame(self.page_container, bg='white')
        title_label = tk.Label(page_frame, text="注册软件", font=('Microsoft YaHei', 14, 'bold'), bg='white', fg='#2980B9')
        title_label.pack(anchor='w', padx=20, pady=20)
        
        main_content_frame = tk.Frame(page_frame, bg='white')
        main_content_frame.pack(padx=20, pady=10)

        font_spec = ('Microsoft YaHei', 12)
        
        machine_code_frame = tk.Frame(main_content_frame, bg='white')
        machine_code_frame.pack(fill=tk.X, pady=10)
        tk.Label(machine_code_frame, text="机器码:", font=font_spec, bg='white').pack(side=tk.LEFT)
        machine_code_val = self.get_machine_code()
        machine_code_entry = tk.Entry(machine_code_frame, font=font_spec, width=30, fg='red')
        machine_code_entry.pack(side=tk.LEFT, padx=10)
        machine_code_entry.insert(0, machine_code_val)
        machine_code_entry.config(state='readonly')

        reg_code_frame = tk.Frame(main_content_frame, bg='white')
        reg_code_frame.pack(fill=tk.X, pady=10)
        tk.Label(reg_code_frame, text="注册码:", font=font_spec, bg='white').pack(side=tk.LEFT)
        self.reg_code_entry = tk.Entry(reg_code_frame, font=font_spec, width=30)
        self.reg_code_entry.pack(side=tk.LEFT, padx=10)
        
        btn_container = tk.Frame(main_content_frame, bg='white')
        btn_container.pack(pady=20)

        register_btn = tk.Button(btn_container, text="注 册", font=('Microsoft YaHei', 12, 'bold'), 
                                 bg='#27AE60', fg='white', width=15, pady=5, command=self.attempt_registration)
        register_btn.pack(pady=5)
        
        cancel_reg_btn = tk.Button(btn_container, text="取消注册", font=('Microsoft YaHei', 12, 'bold'),
                                   bg='#E74C3C', fg='white', width=15, pady=5, command=self.cancel_registration)
        cancel_reg_btn.pack(pady=5)
        
        info_text = "请将您的机器码发送给软件提供商以获取注册码。\n注册码分为月度授权和永久授权两种。"
        tk.Label(main_content_frame, text=info_text, font=('Microsoft YaHei', 10), bg='white', fg='grey').pack(pady=10)

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
        page_frame = tk.Frame(self.page_container, bg='white')
        title_label = tk.Label(page_frame, text="超级管理", font=('Microsoft YaHei', 14, 'bold'), bg='white', fg='#C0392B')
        title_label.pack(anchor='w', padx=20, pady=20)
        desc_label = tk.Label(page_frame, text="警告：此处的任何操作都可能导致数据丢失或配置重置，请谨慎操作。",
                              font=('Microsoft YaHei', 11), bg='white', fg='red', wraplength=700)
        desc_label.pack(anchor='w', padx=20, pady=(0, 20))
        
        btn_frame = tk.Frame(page_frame, bg='white')
        btn_frame.pack(padx=20, pady=10, fill=tk.X)
        
        btn_font = ('Microsoft YaHei', 12, 'bold')
        btn_width = 20; btn_pady = 10

        tk.Button(btn_frame, text="备份所有设置", command=self._backup_all_settings,
                  font=btn_font, width=btn_width, pady=btn_pady, bg='#2980B9', fg='white').pack(pady=10)
        tk.Button(btn_frame, text="还原所有设置", command=self._restore_all_settings,
                  font=btn_font, width=btn_width, pady=btn_pady, bg='#27AE60', fg='white').pack(pady=10)
        tk.Button(btn_frame, text="重置软件", command=self._reset_software,
                  font=btn_font, width=btn_width, pady=btn_pady, bg='#E74C3C', fg='white').pack(pady=10)
        
        tk.Button(btn_frame, text="卸载软件", command=self._prompt_for_uninstall,
                  font=btn_font, width=btn_width, pady=btn_pady, bg='#34495E', fg='white').pack(pady=10)
                  
        return page_frame

    def _prompt_for_uninstall(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("卸载软件 - 身份验证")
        dialog.geometry("350x180")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        self.center_window(dialog, 350, 180)

        result = [None] 

        tk.Label(dialog, text="请输入卸载密码:", font=('Microsoft YaHei', 11)).pack(pady=20)
        password_entry = tk.Entry(dialog, show='*', font=('Microsoft YaHei', 11), width=25)
        password_entry.pack(pady=5)
        password_entry.focus_set()

        def on_confirm():
            result[0] = password_entry.get()
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()

        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=20)
        tk.Button(btn_frame, text="确定", command=on_confirm, width=8).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="取消", command=on_cancel, width=8).pack(side=tk.LEFT, padx=10)
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

        if WIN32COM_AVAILABLE:
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

            self.log("所有设置已从备份文件成功还原。")
            messagebox.showinfo("还原成功", "所有设置已成功还原并立即应用。")
            
            self.root.after(100, lambda: self.switch_page("定时广播"))

        except Exception as e:
            self.log(f"还原失败: {e}"); messagebox.showerror("还原失败", f"发生错误: {e}")
    
    def _refresh_settings_ui(self):
        if "设置" not in self.pages or not hasattr(self, 'autostart_var'):
            return

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
        
        if self.lock_password_b64 and WIN32COM_AVAILABLE:
            self.clear_password_btn.config(state=tk.NORMAL)
        else:
            self.clear_password_btn.config(state=tk.DISABLED)

    def _reset_software(self):
        if not messagebox.askyesno(
            "！！！最终确认！！！",
            "您真的要重置整个软件吗？\n\n此操作将：\n- 清空所有节目单 (但保留音频文件)\n- 清空所有节假日和待办事项\n- 清除锁定密码\n- 重置所有系统设置\n\n此操作【无法恢复】！软件将在重置后提示您重启。"
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
                "autostart": False, "start_minimized": False, "lock_on_start": False,
                "daily_shutdown_enabled": False, "daily_shutdown_time": "23:00:00",
                "weekly_shutdown_enabled": False, "weekly_shutdown_days": "每周:12345", "weekly_shutdown_time": "23:30:00",
                "weekly_reboot_enabled": False, "weekly_reboot_days": "每周:67", "weekly_reboot_time": "22:00:00",
                "last_power_action_date": "",
                "time_chime_enabled": False, "time_chime_voice": "",
                "time_chime_speed": "0", "time_chime_pitch": "0"
            }
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(default_settings, f, ensure_ascii=False, indent=2)
            
            self.log("软件已成功重置。软件需要重启。")
            messagebox.showinfo("重置成功", "软件已恢复到初始状态。\n\n请点击“确定”后手动关闭并重新启动软件。")
        except Exception as e:
            self.log(f"重置失败: {e}"); messagebox.showerror("重置失败", f"发生错误: {e}")

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
            ("禁音频节目", lambda: self._set_tasks_status_by_type('audio', '禁用'), '#E67E22'),
            ("禁语音节目", lambda: self._set_tasks_status_by_type('voice', '禁用'), '#D35400'),
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
        if not WIN32COM_AVAILABLE:
            self.lock_button.config(state=tk.DISABLED, text="锁定(Win)")

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
                                     bg='white', fg='#2C5F7C', padx=10, pady=2)
        playing_frame.pack(fill=tk.X, padx=10, pady=5)
        self.playing_label = tk.Label(playing_frame, text="等待播放...", font=('Microsoft YaHei', 11),
                                      bg='#FFFEF0', anchor='w', justify=tk.LEFT, padx=5)
        self.playing_label.pack(fill=tk.X, expand=True, ipady=4)
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

    def create_holiday_page(self):
        page_frame = tk.Frame(self.page_container, bg='white')

        top_frame = tk.Frame(page_frame, bg='white')
        top_frame.pack(fill=tk.X, padx=10, pady=10)
        title_label = tk.Label(top_frame, text="节假日", font=('Microsoft YaHei', 14, 'bold'),
                              bg='white', fg='#2C5F7C')
        title_label.pack(side=tk.LEFT)
        
        desc_label = tk.Label(page_frame, text="节假日不播放 (手动和立即播任务除外)，整点报时和待办事项也受此约束", font=('Microsoft YaHei', 11),
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
            ("添加", self.add_holiday), ("修改", self.edit_holiday), ("删除", self.delete_holiday),
            (None, None),
            ("全部启用", self.enable_all_holidays), ("全部禁用", self.disable_all_holidays),
            (None, None),
            ("导入节日", self.import_holidays), ("导出节日", self.export_holidays), ("清空节日", self.clear_all_holidays),
        ]

        for text, cmd in buttons_config:
            if text is None:
                tk.Frame(action_frame, height=20, bg='white').pack()
                continue
            
            tk.Button(action_frame, text=text, command=cmd, font=btn_font, width=btn_width, pady=5).pack(pady=5)

        self.update_holiday_list()
        return page_frame
        
    def create_settings_page(self):
        settings_frame = tk.Frame(self.page_container, bg='white')
        
        title_label = tk.Label(settings_frame, text="系统设置", font=('Microsoft YaHei', 14, 'bold'), bg='white', fg='#2C5F7C')
        title_label.pack(anchor='w', padx=20, pady=20)
        
        general_frame = tk.LabelFrame(settings_frame, text="通用设置", font=('Microsoft YaHei', 12, 'bold'), bg='white', padx=15, pady=10)
        general_frame.pack(fill=tk.X, padx=20, pady=10)
        
        self.autostart_var = tk.BooleanVar()
        self.start_minimized_var = tk.BooleanVar()
        self.lock_on_start_var = tk.BooleanVar()
        
        tk.Checkbutton(general_frame, text="登录windows后自动启动", variable=self.autostart_var, font=('Microsoft YaHei', 11), bg='white', anchor='w', command=self._handle_autostart_setting).pack(fill=tk.X, pady=5)
        tk.Checkbutton(general_frame, text="启动后最小化到系统托盘", variable=self.start_minimized_var, font=('Microsoft YaHei', 11), bg='white', anchor='w', command=self.save_settings).pack(fill=tk.X, pady=5)
        
        lock_and_buttons_frame = tk.Frame(general_frame, bg='white')
        lock_and_buttons_frame.pack(fill=tk.X, pady=5)
        
        self.lock_on_start_cb = tk.Checkbutton(lock_and_buttons_frame, text="启动软件后立即锁定", variable=self.lock_on_start_var, font=('Microsoft YaHei', 11), bg='white', anchor='w', command=self._handle_lock_on_start_toggle)
        self.lock_on_start_cb.grid(row=0, column=0, sticky='w')
        if not WIN32COM_AVAILABLE:
            self.lock_on_start_cb.config(state=tk.DISABLED)
            
        tk.Label(lock_and_buttons_frame, text="(请先在主界面设置锁定密码)", font=('Microsoft YaHei', 9), bg='white', fg='grey').grid(row=1, column=0, sticky='w', padx=20)
        
        self.clear_password_btn = tk.Button(lock_and_buttons_frame, text="清除锁定密码", font=('Microsoft YaHei', 11), command=self.clear_lock_password)
        self.clear_password_btn.grid(row=0, column=1, padx=20)
        
        self.cancel_bg_images_btn = tk.Button(lock_and_buttons_frame, text="取消所有节目背景图片", font=('Microsoft YaHei', 11), command=self._cancel_all_background_images)
        self.cancel_bg_images_btn.grid(row=0, column=2, padx=10)

        time_chime_frame = tk.LabelFrame(settings_frame, text="整点报时", font=('Microsoft YaHei', 12, 'bold'), bg='white', padx=15, pady=10)
        time_chime_frame.pack(fill=tk.X, padx=20, pady=10)
        
        self.time_chime_enabled_var = tk.BooleanVar()
        self.time_chime_voice_var = tk.StringVar()
        self.time_chime_speed_var = tk.StringVar()
        self.time_chime_pitch_var = tk.StringVar()
        
        chime_control_frame = tk.Frame(time_chime_frame, bg='white')
        chime_control_frame.pack(fill=tk.X, pady=5)

        tk.Checkbutton(chime_control_frame, text="启用整点报时功能", variable=self.time_chime_enabled_var, font=('Microsoft YaHei', 11), bg='white', anchor='w', command=self._handle_time_chime_toggle).pack(side=tk.LEFT)

        available_voices = self.get_available_voices()
        self.chime_voice_combo = ttk.Combobox(chime_control_frame, textvariable=self.time_chime_voice_var, values=available_voices, font=('Microsoft YaHei', 10), width=35, state='readonly')
        self.chime_voice_combo.pack(side=tk.LEFT, padx=10)
        self.chime_voice_combo.bind("<<ComboboxSelected>>", lambda e: self._on_chime_params_changed(is_voice_change=True))

        params_frame = tk.Frame(chime_control_frame, bg='white')
        params_frame.pack(side=tk.LEFT, padx=10)
        tk.Label(params_frame, text="语速(-10~10):", font=('Microsoft YaHei', 10), bg='white').pack(side=tk.LEFT)
        speed_entry = tk.Entry(params_frame, textvariable=self.time_chime_speed_var, font=('Microsoft YaHei', 10), width=5)
        speed_entry.pack(side=tk.LEFT, padx=(0, 10))
        tk.Label(params_frame, text="音调(-10~10):", font=('Microsoft YaHei', 10), bg='white').pack(side=tk.LEFT)
        pitch_entry = tk.Entry(params_frame, textvariable=self.time_chime_pitch_var, font=('Microsoft YaHei', 10), width=5)
        pitch_entry.pack(side=tk.LEFT)
        
        speed_entry.bind("<FocusOut>", self._on_chime_params_changed)
        pitch_entry.bind("<FocusOut>", self._on_chime_params_changed)

        power_frame = tk.LabelFrame(settings_frame, text="电源管理", font=('Microsoft YaHei', 12, 'bold'), bg='white', padx=15, pady=10)
        power_frame.pack(fill=tk.X, padx=20, pady=10)

        self.daily_shutdown_enabled_var = tk.BooleanVar()
        self.daily_shutdown_time_var = tk.StringVar()
        self.weekly_shutdown_enabled_var = tk.BooleanVar()
        self.weekly_shutdown_time_var = tk.StringVar()
        self.weekly_shutdown_days_var = tk.StringVar()
        self.weekly_reboot_enabled_var = tk.BooleanVar()
        self.weekly_reboot_time_var = tk.StringVar()
        self.weekly_reboot_days_var = tk.StringVar()

        daily_frame = tk.Frame(power_frame, bg='white')
        daily_frame.pack(fill=tk.X, pady=4)
        tk.Checkbutton(daily_frame, text="每天关机", variable=self.daily_shutdown_enabled_var, font=('Microsoft YaHei', 11), bg='white', command=self.save_settings).pack(side=tk.LEFT)
        tk.Entry(daily_frame, textvariable=self.daily_shutdown_time_var, font=('Microsoft YaHei', 11), width=15).pack(side=tk.LEFT, padx=10)
        tk.Button(daily_frame, text="设置", font=('Microsoft YaHei', 11), command=lambda: self.show_single_time_dialog(self.daily_shutdown_time_var)).pack(side=tk.LEFT)

        weekly_frame = tk.Frame(power_frame, bg='white')
        weekly_frame.pack(fill=tk.X, pady=4)
        tk.Checkbutton(weekly_frame, text="每周关机", variable=self.weekly_shutdown_enabled_var, font=('Microsoft YaHei', 11), bg='white', command=self.save_settings).pack(side=tk.LEFT)
        tk.Entry(weekly_frame, textvariable=self.weekly_shutdown_days_var, font=('Microsoft YaHei', 11), width=20).pack(side=tk.LEFT, padx=(10,5))
        tk.Entry(weekly_frame, textvariable=self.weekly_shutdown_time_var, font=('Microsoft YaHei', 11), width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(weekly_frame, text="设置", font=('Microsoft YaHei', 11), command=lambda: self.show_power_week_time_dialog("设置每周关机", self.weekly_shutdown_days_var, self.weekly_shutdown_time_var)).pack(side=tk.LEFT)

        reboot_frame = tk.Frame(power_frame, bg='white')
        reboot_frame.pack(fill=tk.X, pady=4)
        tk.Checkbutton(reboot_frame, text="每周重启", variable=self.weekly_reboot_enabled_var, font=('Microsoft YaHei', 11), bg='white', command=self.save_settings).pack(side=tk.LEFT)
        tk.Entry(reboot_frame, textvariable=self.weekly_reboot_days_var, font=('Microsoft YaHei', 11), width=20).pack(side=tk.LEFT, padx=(10,5))
        tk.Entry(reboot_frame, textvariable=self.weekly_reboot_time_var, font=('Microsoft YaHei', 11), width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(reboot_frame, text="设置", font=('Microsoft YaHei', 11), command=lambda: self.show_power_week_time_dialog("设置每周重启", self.weekly_reboot_days_var, self.weekly_reboot_time_var)).pack(side=tk.LEFT)

        return settings_frame

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
            
            progress_dialog = tk.Toplevel(self.root)
            progress_dialog.title("请稍候")
            progress_dialog.geometry("350x120")
            progress_dialog.resizable(False, False)
            progress_dialog.transient(self.root); progress_dialog.grab_set()
            self.center_window(progress_dialog, 350, 120)
            
            tk.Label(progress_dialog, text="正在生成整点报时文件 (0/24)...", font=('Microsoft YaHei', 11)).pack(pady=10)
            progress_label = tk.Label(progress_dialog, text="", font=('Microsoft YaHei', 10))
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
        self.lock_button.config(text="解锁", bg='#2ECC71')
        self._set_ui_lock_state(tk.DISABLED)
        self.statusbar_unlock_button.pack(side=tk.RIGHT, padx=5)
        self.log("界面已锁定。")

    def _apply_unlock(self):
        self.is_locked = False
        self.lock_button.config(text="锁定", bg='#E74C3C')
        self._set_ui_lock_state(tk.NORMAL)
        self.statusbar_unlock_button.pack_forget()
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
            if self._save_to_registry("LockPasswordB64", encoded_pass):
                self.lock_password_b64 = encoded_pass
                if "设置" in self.pages and hasattr(self, 'clear_password_btn'):
                    self.clear_password_btn.config(state=tk.NORMAL)
                messagebox.showinfo("成功", "密码设置成功，界面即将锁定。", parent=dialog)
                dialog.destroy()
                self._apply_lock()
            else:
                messagebox.showerror("功能受限", "无法保存密码。\n此功能仅在Windows系统上支持且需要pywin32库。", parent=dialog)

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

        btn_frame = tk.Frame(dialog); btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="确定", command=confirm, font=('Microsoft YaHei', 11)).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="清除密码", command=clear_password_action, font=('Microsoft YaHei', 11)).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="取消", command=dialog.destroy, font=('Microsoft YaHei', 11)).pack(side=tk.LEFT, padx=5)
        dialog.bind('<Return>', lambda event: confirm())

    def _perform_password_clear_logic(self):
        if self._save_to_registry("LockPasswordB64", ""):
            self.lock_password_b64 = ""
            self.settings["lock_on_start"] = False
            
            if hasattr(self, 'lock_on_start_var'):
                self.lock_on_start_var.set(False)
            
            self.save_settings()
            
            if hasattr(self, 'clear_password_btn'):
                self.clear_password_btn.config(state=tk.DISABLED)
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
        dialog.geometry("950x750")
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set(); dialog.configure(bg='#E8E8E8')
        
        main_frame = tk.Frame(dialog, bg='#E8E8E8', padx=15, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        content_frame = tk.LabelFrame(main_frame, text="内容", font=('Microsoft YaHei', 12, 'bold'),
                                     bg='#E8E8E8', padx=10, pady=5)
        content_frame.grid(row=0, column=0, sticky='ew', pady=2)
        font_spec = ('Microsoft YaHei', 11)
        
        tk.Label(content_frame, text="节目名称:", font=font_spec, bg='#E8E8E8').grid(row=0, column=0, sticky='e', padx=5, pady=2)
        name_entry = tk.Entry(content_frame, font=font_spec, width=55)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        audio_type_var = tk.StringVar(value="single")
        tk.Label(content_frame, text="音频文件", font=font_spec, bg='#E8E8E8').grid(row=1, column=0, sticky='e', padx=5, pady=2)
        audio_single_frame = tk.Frame(content_frame, bg='#E8E8E8')
        audio_single_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        tk.Radiobutton(audio_single_frame, text="", variable=audio_type_var, value="single", bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT)
        audio_single_entry = tk.Entry(audio_single_frame, font=font_spec, width=35)
        audio_single_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(audio_single_frame, text="00:00", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=10)
        def select_single_audio():
            filename = filedialog.askopenfilename(title="选择音频文件", initialdir=AUDIO_FOLDER, filetypes=[("音频文件", "*.mp3 *.wav *.ogg *.flac *.m4a"), ("所有文件", "*.*")])
            if filename: audio_single_entry.delete(0, tk.END); audio_single_entry.insert(0, filename)
        tk.Button(audio_single_frame, text="选取...", command=select_single_audio, bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=1).pack(side=tk.LEFT, padx=5)
        tk.Label(content_frame, text="音频文件夹", font=font_spec, bg='#E8E8E8').grid(row=2, column=0, sticky='e', padx=5, pady=2)
        audio_folder_frame = tk.Frame(content_frame, bg='#E8E8E8')
        audio_folder_frame.grid(row=2, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        tk.Radiobutton(audio_folder_frame, text="", variable=audio_type_var, value="folder", bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT)
        audio_folder_entry = tk.Entry(audio_folder_frame, font=font_spec, width=50)
        audio_folder_entry.pack(side=tk.LEFT, padx=5)
        def select_folder(entry_widget):
            foldername = filedialog.askdirectory(title="选择文件夹", initialdir=application_path)
            if foldername: entry_widget.delete(0, tk.END); entry_widget.insert(0, foldername)
        tk.Button(audio_folder_frame, text="选取...", command=lambda: select_folder(audio_folder_entry), bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=1).pack(side=tk.LEFT, padx=5)
        play_order_frame = tk.Frame(content_frame, bg='#E8E8E8')
        play_order_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        play_order_var = tk.StringVar(value="sequential")
        tk.Radiobutton(play_order_frame, text="顺序播", variable=play_order_var, value="sequential", bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(play_order_frame, text="随机播", variable=play_order_var, value="random", bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT, padx=10)
        
        bg_image_var = tk.IntVar(value=0)
        bg_image_path_var = tk.StringVar()
        bg_image_order_var = tk.StringVar(value="sequential")

        bg_image_frame = tk.Frame(content_frame, bg='#E8E8E8')
        bg_image_frame.grid(row=4, column=0, columnspan=4, sticky='w', padx=5, pady=5)
        
        bg_image_cb = tk.Checkbutton(bg_image_frame, text="背景图片:", variable=bg_image_var, bg='#E8E8E8', font=font_spec)
        bg_image_cb.pack(side=tk.LEFT)
        if not IMAGE_AVAILABLE: bg_image_cb.config(state=tk.DISABLED, text="背景图片(Pillow未安装):")

        bg_image_entry = tk.Entry(bg_image_frame, textvariable=bg_image_path_var, font=font_spec, width=42)
        bg_image_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        tk.Button(bg_image_frame, text="选取...", command=lambda: select_folder(bg_image_entry), bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=1).pack(side=tk.LEFT, padx=5)
        
        tk.Radiobutton(bg_image_frame, text="顺序", variable=bg_image_order_var, value="sequential", bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT, padx=(10,0))
        tk.Radiobutton(bg_image_frame, text="随机", variable=bg_image_order_var, value="random", bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT)

        volume_frame = tk.Frame(content_frame, bg='#E8E8E8')
        volume_frame.grid(row=5, column=1, columnspan=3, sticky='w', padx=5, pady=3)
        tk.Label(volume_frame, text="音量:", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT)
        volume_entry = tk.Entry(volume_frame, font=font_spec, width=10)
        volume_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(volume_frame, text="0-100", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=5)
        
        time_frame = tk.LabelFrame(main_frame, text="时间", font=('Microsoft YaHei', 12, 'bold'), bg='#E8E8E8', padx=15, pady=10)
        time_frame.grid(row=1, column=0, sticky='ew', pady=4)
        tk.Label(time_frame, text="开始时间:", font=font_spec, bg='#E8E8E8').grid(row=0, column=0, sticky='e', padx=5, pady=2)
        start_time_entry = tk.Entry(time_frame, font=font_spec, width=50)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        tk.Label(time_frame, text="《可多个,用英文逗号,隔开》", font=font_spec, bg='#E8E8E8').grid(row=0, column=2, sticky='w', padx=5)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=1).grid(row=0, column=3, padx=5)
        interval_var = tk.StringVar(value="first")
        interval_frame1 = tk.Frame(time_frame, bg='#E8E8E8')
        interval_frame1.grid(row=1, column=1, columnspan=2, sticky='w', padx=5, pady=2)
        tk.Label(time_frame, text="间隔播报:", font=font_spec, bg='#E8E8E8').grid(row=1, column=0, sticky='e', padx=5, pady=2)
        tk.Radiobutton(interval_frame1, text="播 n 首", variable=interval_var, value="first", bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT)
        interval_first_entry = tk.Entry(interval_frame1, font=font_spec, width=15)
        interval_first_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(interval_frame1, text="(单曲时,指 n 遍)", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=5)
        interval_frame2 = tk.Frame(time_frame, bg='#E8E8E8')
        interval_frame2.grid(row=2, column=1, columnspan=2, sticky='w', padx=5, pady=2)
        tk.Radiobutton(interval_frame2, text="播 n 秒", variable=interval_var, value="seconds", bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT)
        interval_seconds_entry = tk.Entry(interval_frame2, font=font_spec, width=15)
        interval_seconds_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(interval_frame2, text="(3600秒 = 1小时)", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=5)
        tk.Label(time_frame, text="周几/几号:", font=font_spec, bg='#E8E8E8').grid(row=3, column=0, sticky='e', padx=5, pady=3)
        weekday_entry = tk.Entry(time_frame, font=font_spec, width=50)
        weekday_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=3)
        tk.Button(time_frame, text="选取...", command=lambda: self.show_weekday_settings_dialog(weekday_entry), bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=1).grid(row=3, column=3, padx=5)
        tk.Label(time_frame, text="日期范围:", font=font_spec, bg='#E8E8E8').grid(row=4, column=0, sticky='e', padx=5, pady=3)
        date_range_entry = tk.Entry(time_frame, font=font_spec, width=50)
        date_range_entry.grid(row=4, column=1, sticky='ew', padx=5, pady=3)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry), bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=1).grid(row=4, column=3, padx=5)
        
        other_frame = tk.LabelFrame(main_frame, text="其它", font=('Microsoft YaHei', 12, 'bold'), bg='#E8E8E8', padx=10, pady=10)
        other_frame.grid(row=2, column=0, sticky='ew', pady=5)
        delay_var = tk.StringVar(value="ontime")
        tk.Label(other_frame, text="模式:", font=font_spec, bg='#E8E8E8').grid(row=0, column=0, sticky='nw', padx=5, pady=2)
        delay_frame = tk.Frame(other_frame, bg='#E8E8E8')
        delay_frame.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        tk.Radiobutton(delay_frame, text="准时播 - 如果有别的节目正在播，终止他们（默认）", variable=delay_var, value="ontime", bg='#E8E8E8', font=font_spec).pack(anchor='w')
        tk.Radiobutton(delay_frame, text="可延后 - 如果有别的节目正在播，排队等候", variable=delay_var, value="delay", bg='#E8E8E8', font=font_spec).pack(anchor='w')
        tk.Radiobutton(delay_frame, text="立即播 - 添加后停止其他节目,立即播放此节目", variable=delay_var, value="immediate", bg='#E8E8E8', font=font_spec).pack(anchor='w')
        dialog_button_frame = tk.Frame(other_frame, bg='#E8E8E8')
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
        tk.Button(dialog_button_frame, text=button_text, command=save_task, bg='#5DADE2', fg='white', font=('Microsoft YaHei', 11, 'bold'), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        tk.Button(dialog_button_frame, text="取消", command=dialog.destroy, bg='#D0D0D0', font=('Microsoft YaHei', 11), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        
        content_frame.columnconfigure(1, weight=1); time_frame.columnconfigure(1, weight=1)

    def open_voice_dialog(self, parent_dialog, task_to_edit=None, index=None):
        parent_dialog.destroy()
        is_edit_mode = task_to_edit is not None
        dialog = tk.Toplevel(self.root)
        dialog.title("修改语音节目" if is_edit_mode else "添加语音节目")
        dialog.geometry("950x750")
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set(); dialog.configure(bg='#E8E8E8')

        main_frame = tk.Frame(dialog, bg='#E8E8E8', padx=15, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        content_frame = tk.LabelFrame(main_frame, text="内容", font=('Microsoft YaHei', 12, 'bold'), bg='#E8E8E8', padx=10, pady=5)
        content_frame.grid(row=0, column=0, sticky='ew', pady=2)
        font_spec = ('Microsoft YaHei', 11)
        
        tk.Label(content_frame, text="节目名称:", font=font_spec, bg='#E8E8E8').grid(row=0, column=0, sticky='w', padx=5, pady=2)
        name_entry = tk.Entry(content_frame, font=font_spec, width=65)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        tk.Label(content_frame, text="播音文字:", font=font_spec, bg='#E8E8E8').grid(row=1, column=0, sticky='nw', padx=5, pady=2)
        text_frame = tk.Frame(content_frame, bg='#E8E8E8')
        text_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        content_text = scrolledtext.ScrolledText(text_frame, height=5, font=font_spec, width=65, wrap=tk.WORD)
        content_text.pack(fill=tk.BOTH, expand=True)
        script_btn_frame = tk.Frame(content_frame, bg='#E8E8E8')
        script_btn_frame.grid(row=2, column=1, columnspan=3, sticky='w', padx=5, pady=(0, 2))
        tk.Button(script_btn_frame, text="导入文稿", command=lambda: self._import_voice_script(content_text), font=('Microsoft YaHei', 10)).pack(side=tk.LEFT)
        tk.Button(script_btn_frame, text="导出文稿", command=lambda: self._export_voice_script(content_text, name_entry), font=('Microsoft YaHei', 10)).pack(side=tk.LEFT, padx=10)
        tk.Label(content_frame, text="播音员:", font=font_spec, bg='#E8E8E8').grid(row=3, column=0, sticky='w', padx=5, pady=3)
        voice_frame = tk.Frame(content_frame, bg='#E8E8E8')
        voice_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=3)
        available_voices = self.get_available_voices()
        voice_var = tk.StringVar()
        voice_combo = ttk.Combobox(voice_frame, textvariable=voice_var, values=available_voices, font=font_spec, width=50, state='readonly')
        voice_combo.pack(side=tk.LEFT)
        speech_params_frame = tk.Frame(content_frame, bg='#E8E8E8')
        speech_params_frame.grid(row=4, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        tk.Label(speech_params_frame, text="语速(-10~10):", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=(0,5))
        speed_entry = tk.Entry(speech_params_frame, font=font_spec, width=8); speed_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(speech_params_frame, text="音调(-10~10):", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=(10,5))
        pitch_entry = tk.Entry(speech_params_frame, font=font_spec, width=8); pitch_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(speech_params_frame, text="音量(0-100):", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=(10,5))
        volume_entry = tk.Entry(speech_params_frame, font=font_spec, width=8); volume_entry.pack(side=tk.LEFT, padx=5)
        prompt_var = tk.IntVar(); prompt_frame = tk.Frame(content_frame, bg='#E8E8E8')
        prompt_frame.grid(row=5, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        tk.Checkbutton(prompt_frame, text="提示音:", variable=prompt_var, bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT)
        prompt_file_var, prompt_volume_var = tk.StringVar(), tk.StringVar()
        prompt_file_entry = tk.Entry(prompt_frame, textvariable=prompt_file_var, font=font_spec, width=20); prompt_file_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(prompt_frame, text="...", command=lambda: self.select_file_for_entry(PROMPT_FOLDER, prompt_file_var)).pack(side=tk.LEFT)
        tk.Label(prompt_frame, text="音量(0-100):", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=(10,5))
        tk.Entry(prompt_frame, textvariable=prompt_volume_var, font=font_spec, width=8).pack(side=tk.LEFT, padx=5)
        bgm_var = tk.IntVar(); bgm_frame = tk.Frame(content_frame, bg='#E8E8E8')
        bgm_frame.grid(row=6, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        tk.Checkbutton(bgm_frame, text="背景音乐:", variable=bgm_var, bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT)
        bgm_file_var, bgm_volume_var = tk.StringVar(), tk.StringVar()
        bgm_file_entry = tk.Entry(bgm_frame, textvariable=bgm_file_var, font=font_spec, width=20); bgm_file_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(bgm_frame, text="...", command=lambda: self.select_file_for_entry(BGM_FOLDER, bgm_file_var)).pack(side=tk.LEFT)
        tk.Label(bgm_frame, text="音量(0-100):", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=(10,5))
        tk.Entry(bgm_frame, textvariable=bgm_volume_var, font=font_spec, width=8).pack(side=tk.LEFT, padx=5)
        
        def select_folder(entry_widget):
            foldername = filedialog.askdirectory(title="选择文件夹", initialdir=application_path)
            if foldername: entry_widget.delete(0, tk.END); entry_widget.insert(0, foldername)
            
        bg_image_var = tk.IntVar(value=0)
        bg_image_path_var = tk.StringVar()
        bg_image_order_var = tk.StringVar(value="sequential")
        
        bg_image_frame = tk.Frame(content_frame, bg='#E8E8E8')
        bg_image_frame.grid(row=7, column=1, columnspan=3, sticky='w', padx=5, pady=5)
        
        bg_image_cb = tk.Checkbutton(bg_image_frame, text="背景图片:", variable=bg_image_var, bg='#E8E8E8', font=font_spec)
        bg_image_cb.pack(side=tk.LEFT)
        if not IMAGE_AVAILABLE: bg_image_cb.config(state=tk.DISABLED, text="背景图片(Pillow未安装):")

        bg_image_entry = tk.Entry(bg_image_frame, textvariable=bg_image_path_var, font=font_spec, width=32)
        bg_image_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        tk.Button(bg_image_frame, text="选取...", command=lambda: select_folder(bg_image_entry), bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=1).pack(side=tk.LEFT, padx=5)
        
        tk.Radiobutton(bg_image_frame, text="顺序", variable=bg_image_order_var, value="sequential", bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT, padx=(10,0))
        tk.Radiobutton(bg_image_frame, text="随机", variable=bg_image_order_var, value="random", bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT)

        time_frame = tk.LabelFrame(main_frame, text="时间", font=('Microsoft YaHei', 12, 'bold'), bg='#E8E8E8', padx=10, pady=5)
        time_frame.grid(row=1, column=0, sticky='ew', pady=2)
        tk.Label(time_frame, text="开始时间:", font=font_spec, bg='#E8E8E8').grid(row=0, column=0, sticky='e', padx=5, pady=2)
        start_time_entry = tk.Entry(time_frame, font=font_spec, width=50)
        start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        tk.Label(time_frame, text="《可多个,用英文逗号,隔开》", font=font_spec, bg='#E8E8E8').grid(row=0, column=2, sticky='w', padx=5)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_time_settings_dialog(start_time_entry), bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=1).grid(row=0, column=3, padx=5)
        tk.Label(time_frame, text="播 n 遍:", font=font_spec, bg='#E8E8E8').grid(row=1, column=0, sticky='e', padx=5, pady=2)
        repeat_entry = tk.Entry(time_frame, font=font_spec, width=12)
        repeat_entry.grid(row=1, column=1, sticky='w', padx=5, pady=2)
        tk.Label(time_frame, text="周几/几号:", font=font_spec, bg='#E8E8E8').grid(row=2, column=0, sticky='e', padx=5, pady=2)
        weekday_entry = tk.Entry(time_frame, font=font_spec, width=50)
        weekday_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=2)
        tk.Button(time_frame, text="选取...", command=lambda: self.show_weekday_settings_dialog(weekday_entry), bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=1).grid(row=2, column=3, padx=5)
        tk.Label(time_frame, text="日期范围:", font=font_spec, bg='#E8E8E8').grid(row=3, column=0, sticky='e', padx=5, pady=2)
        date_range_entry = tk.Entry(time_frame, font=font_spec, width=50)
        date_range_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=2)
        tk.Button(time_frame, text="设置...", command=lambda: self.show_daterange_settings_dialog(date_range_entry), bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=1).grid(row=3, column=3, padx=5)
        
        other_frame = tk.LabelFrame(main_frame, text="其它", font=('Microsoft YaHei', 12, 'bold'), bg='#E8E8E8', padx=15, pady=10)
        other_frame.grid(row=2, column=0, sticky='ew', pady=4)
        delay_var = tk.StringVar(value="delay")
        tk.Label(other_frame, text="模式:", font=font_spec, bg='#E8E8E8').grid(row=0, column=0, sticky='nw', padx=5, pady=2)
        delay_frame = tk.Frame(other_frame, bg='#E8E8E8')
        delay_frame.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        tk.Radiobutton(delay_frame, text="准时播 - 如果有别的节目正在播，终止他们", variable=delay_var, value="ontime", bg='#E8E8E8', font=font_spec).pack(anchor='w', pady=1)
        tk.Radiobutton(delay_frame, text="可延后 - 如果有别的节目正在播，排队等候（默认）", variable=delay_var, value="delay", bg='#E8E8E8', font=font_spec).pack(anchor='w', pady=1)
        tk.Radiobutton(delay_frame, text="立即播 - 添加后停止其他节目,立即播放此节目", variable=delay_var, value="immediate", bg='#E8E8E8', font=font_spec).pack(anchor='w', pady=1)
        dialog_button_frame = tk.Frame(other_frame, bg='#E8E8E8')
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
            text_content = content_text.get('1.0', tk.END).strip()
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
                saved_delay_type = task_to_edit.get('delay', 'delay') if is_edit_mode else play_mode
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
                }

            if not regeneration_needed:
                new_task_data = build_task_data(task_to_edit.get('content'), task_to_edit.get('wav_filename'))
                if not new_task_data['name'] or not new_task_data['time']: messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog); return
                self.tasks[index] = new_task_data; self.log(f"已修改语音节目(未重新生成语音): {new_task_data['name']}")
                self.update_task_list(); self.save_tasks(); dialog.destroy()
                if delay_var.get() == 'immediate': self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))
                return
            progress_dialog = tk.Toplevel(dialog); progress_dialog.title("请稍候"); progress_dialog.geometry("300x100")
            progress_dialog.resizable(False, False); progress_dialog.transient(dialog); progress_dialog.grab_set()
            tk.Label(progress_dialog, text="语音文件生成中，请稍后...", font=font_spec).pack(expand=True)
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
                new_task_data = build_task_data(output_path, new_wav_filename)
                if not new_task_data['name'] or not new_task_data['time']: messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog); return
                if is_edit_mode: self.tasks[index] = new_task_data; self.log(f"已修改语音节目(并重新生成语音): {new_task_data['name']}")
                else: self.tasks.append(new_task_data); self.log(f"已添加语音节目: {new_task_data['name']}")
                self.update_task_list(); self.save_tasks(); dialog.destroy()
                if delay_var.get() == 'immediate': self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))
            synthesis_thread = threading.Thread(target=self._synthesis_worker, args=(text_content, voice_params, output_path, _on_synthesis_complete))
            synthesis_thread.daemon = True; synthesis_thread.start()
        
        button_text = "保存修改" if is_edit_mode else "添加"
        tk.Button(dialog_button_frame, text=button_text, command=save_task, bg='#5DADE2', fg='white', font=('Microsoft YaHei', 11, 'bold'), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        tk.Button(dialog_button_frame, text="取消", command=dialog.destroy, bg='#D0D0D0', font=('Microsoft YaHei', 11), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        
        content_frame.columnconfigure(1, weight=1); time_frame.columnconfigure(1, weight=1)

# --- 代码第一部分结束 ---
# --- NEW & REWORKED: 待办事项所有相关函数 ---

    def create_todo_page(self):
        page_frame = tk.Frame(self.page_container, bg='white')

        top_frame = tk.Frame(page_frame, bg='white')
        top_frame.pack(fill=tk.X, padx=10, pady=10)
        title_label = tk.Label(top_frame, text="待办事项", font=('Microsoft YaHei', 14, 'bold'), bg='white', fg='#2C5F7C')
        title_label.pack(side=tk.LEFT)

        desc_label = tk.Label(page_frame, text="到达提醒时间时会弹出窗口提醒，提醒功能受节假日约束。", font=('Microsoft YaHei', 11), bg='white', fg='#555')
        desc_label.pack(anchor='w', padx=10, pady=(0, 10))

        content_frame = tk.Frame(page_frame, bg='white')
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        table_frame = tk.Frame(content_frame, bg='white')
        table_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        columns = ('待办事项名称', '状态', '类型', '内容', '提醒规则')
        self.todo_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15, selectmode='extended')

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

        self.todo_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.todo_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.todo_tree.configure(yscrollcommand=scrollbar.set)

        self.todo_tree.bind("<Double-1>", lambda e: self.edit_todo())
        self.todo_tree.bind("<Button-3>", self.show_todo_context_menu)
        self._enable_drag_selection(self.todo_tree)

        action_frame = tk.Frame(content_frame, bg='white', padx=10)
        action_frame.pack(side=tk.RIGHT, fill=tk.Y)

        btn_font = ('Microsoft YaHei', 11)
        btn_width = 10
        buttons_config = [
            ("添加", self.add_todo), ("修改", self.edit_todo), ("删除", self.delete_todo),
            (None, None),
            ("全部启用", self.enable_all_todos), ("全部禁用", self.disable_all_todos),
            (None, None),
            ("导入事项", self.import_todos), ("导出事项", self.export_todos), ("清空事项", self.clear_all_todos),
        ]

        for text, cmd in buttons_config:
            if text is None:
                tk.Frame(action_frame, height=20, bg='white').pack()
                continue
            tk.Button(action_frame, text=text, command=cmd, font=btn_font, width=btn_width, pady=5).pack(pady=5)

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

    def update_todo_list(self):
        if not hasattr(self, 'todo_tree') or not self.todo_tree.winfo_exists(): return
        selection = self.todo_tree.selection()
        self.todo_tree.delete(*self.todo_tree.get_children())

        for todo in self.todos:
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
            
            self.todo_tree.insert('', tk.END, values=(
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

    def add_todo(self):
        self.open_todo_dialog()

    def edit_todo(self):
        selection = self.todo_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要修改的待办事项")
            return
        if len(selection) > 1:
            messagebox.showwarning("警告", "一次只能修改一个待办事项")
            return
        index = self.todo_tree.index(selection[0])
        todo_to_edit = self.todos[index]
        self.open_todo_dialog(todo_to_edit=todo_to_edit, index=index)

    def delete_todo(self):
        selections = self.todo_tree.selection()
        if not selections:
            messagebox.showwarning("警告", "请先选择要删除的待办事项")
            return
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(selections)} 个待办事项吗？"):
            indices = sorted([self.todo_tree.index(s) for s in selections], reverse=True)
            for index in indices:
                self.todos.pop(index)
            self.update_todo_list()
            self.save_todos()

    def _set_todo_status(self, status):
        selection = self.todo_tree.selection()
        if not selection:
            messagebox.showwarning("警告", f"请先选择要{status}的待办事项")
            return
        for item_id in selection:
            index = self.todo_tree.index(item_id)
            self.todos[index]['status'] = status
        self.update_todo_list()
        self.save_todos()

    def open_todo_dialog(self, todo_to_edit=None, index=None):
        dialog = tk.Toplevel(self.root)
        dialog.title("修改待办事项" if todo_to_edit else "添加待办事项")
        dialog.geometry("750x600")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg='#F0F8FF')
        self.center_window(dialog, 750, 600)

        font_spec = ('Microsoft YaHei', 11)
        main_frame = tk.Frame(dialog, bg='#F0F8FF', padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(main_frame, text="名称:", font=font_spec, bg='#F0F8FF').grid(row=0, column=0, sticky='e', pady=5, padx=5)
        name_entry = tk.Entry(main_frame, font=font_spec, width=60)
        name_entry.grid(row=0, column=1, columnspan=3, sticky='w', pady=5)

        tk.Label(main_frame, text="内容:", font=font_spec, bg='#F0F8FF').grid(row=1, column=0, sticky='ne', pady=5, padx=5)
        content_text = scrolledtext.ScrolledText(main_frame, height=5, font=font_spec, width=60, wrap=tk.WORD)
        content_text.grid(row=1, column=1, columnspan=3, sticky='w', pady=5)
        
        type_var = tk.StringVar(value="onetime")
        type_frame = tk.Frame(main_frame, bg='#F0F8FF')
        type_frame.grid(row=2, column=1, columnspan=3, sticky='w', pady=10)
        
        onetime_rb = tk.Radiobutton(type_frame, text="一次性任务", variable=type_var, value="onetime", bg='#F0F8FF', font=font_spec)
        onetime_rb.pack(side=tk.LEFT, padx=10)
        recurring_rb = tk.Radiobutton(type_frame, text="循环任务", variable=type_var, value="recurring", bg='#F0F8FF', font=font_spec)
        recurring_rb.pack(side=tk.LEFT, padx=10)
        
        onetime_lf = tk.LabelFrame(main_frame, text="一次性任务设置", font=font_spec, bg='#F0F8FF', padx=10, pady=10)
        recurring_lf = tk.LabelFrame(main_frame, text="循环任务设置", font=font_spec, bg='#F0F8FF', padx=10, pady=10)
        
        tk.Label(onetime_lf, text="执行日期:", font=font_spec, bg='#F0F8FF').grid(row=0, column=0, sticky='e', pady=5, padx=5)
        onetime_date_entry = tk.Entry(onetime_lf, font=font_spec, width=20)
        onetime_date_entry.grid(row=0, column=1, sticky='w', pady=5)
        tk.Label(onetime_lf, text="执行时间:", font=font_spec, bg='#F0F8FF').grid(row=1, column=0, sticky='e', pady=5, padx=5)
        onetime_time_entry = tk.Entry(onetime_lf, font=font_spec, width=20)
        onetime_time_entry.grid(row=1, column=1, sticky='w', pady=5)

        tk.Label(recurring_lf, text="开始时间:", font=font_spec, bg='#F0F8FF').grid(row=0, column=0, sticky='e', padx=5, pady=5)
        recurring_time_entry = tk.Entry(recurring_lf, font=font_spec, width=40)
        recurring_time_entry.grid(row=0, column=1, sticky='w', padx=5, pady=5)
        tk.Button(recurring_lf, text="设置...", command=lambda: self.show_time_settings_dialog(recurring_time_entry), bg='#D0D0D0', font=font_spec).grid(row=0, column=2, padx=5)
        
        tk.Label(recurring_lf, text="周几/几号:", font=font_spec, bg='#F0F8FF').grid(row=1, column=0, sticky='e', padx=5, pady=5)
        recurring_weekday_entry = tk.Entry(recurring_lf, font=font_spec, width=40)
        recurring_weekday_entry.grid(row=1, column=1, sticky='w', padx=5, pady=5)
        tk.Button(recurring_lf, text="选取...", command=lambda: self.show_weekday_settings_dialog(recurring_weekday_entry), bg='#D0D0D0', font=font_spec).grid(row=1, column=2, padx=5)

        tk.Label(recurring_lf, text="日期范围:", font=font_spec, bg='#F0F8FF').grid(row=2, column=0, sticky='e', padx=5, pady=5)
        recurring_daterange_entry = tk.Entry(recurring_lf, font=font_spec, width=40)
        recurring_daterange_entry.grid(row=2, column=1, sticky='w', padx=5, pady=5)
        tk.Button(recurring_lf, text="设置...", command=lambda: self.show_daterange_settings_dialog(recurring_daterange_entry), bg='#D0D0D0', font=font_spec).grid(row=2, column=2, padx=5)

        tk.Label(recurring_lf, text="循环间隔:", font=font_spec, bg='#F0F8FF').grid(row=3, column=0, sticky='e', padx=5, pady=5)
        interval_frame = tk.Frame(recurring_lf, bg='#F0F8FF')
        interval_frame.grid(row=3, column=1, sticky='w', padx=5, pady=5)
        recurring_interval_entry = tk.Entry(interval_frame, font=font_spec, width=8)
        recurring_interval_entry.pack(side=tk.LEFT)
        tk.Label(interval_frame, text="分钟 (0表示仅在'开始时间'提醒)", font=('Microsoft YaHei', 10), bg='#F0F8FF').pack(side=tk.LEFT, padx=5)
        
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
                "content": content_text.get('1.0', tk.END).strip(),
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
            else: # recurring
                try:
                    interval = int(recurring_interval_entry.get().strip() or '0')
                    if not (0 <= interval <= 60): raise ValueError
                except ValueError:
                    messagebox.showerror("格式错误", "循环间隔必须是 0-60 之间的整数。", parent=dialog)
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
            dialog.destroy()

        button_frame = tk.Frame(main_frame, bg='#F0F8FF')
        button_frame.grid(row=4, column=0, columnspan=4, pady=20)
        tk.Button(button_frame, text="保存", command=save, font=font_spec, width=10).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="取消", command=dialog.destroy, font=font_spec, width=10).pack(side=tk.LEFT, padx=10)
    
    def _flash_window(self):
        """调用Windows API使任务栏图标闪烁"""
        try:
            hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
            
            class FLASHWINFO(ctypes.Structure):
                _fields_ = [
                    ("cbSize", ctypes.c_uint),
                    ("hwnd", wintypes.HWND),
                    ("dwFlags", ctypes.c_ulong),
                    ("uCount", ctypes.c_uint),
                    ("dwTimeout", ctypes.c_ulong)
                ]

            info = FLASHWINFO()
            info.cbSize = ctypes.sizeof(info)
            info.hwnd = hwnd
            info.dwFlags = 2 | 12  # FLASHW_TRAY | FLASHW_TIMERNOFG
            info.uCount = 0
            info.dwTimeout = 0
            
            ctypes.windll.user32.FlashWindowEx(ctypes.byref(info))
        except Exception as e:
            print(f"警告: 无法使窗口闪烁 - {e}")

    def show_todo_reminder(self, todo):
        reminder_win = tk.Toplevel(self.root)
        reminder_win.title(f"待办事项提醒 - {todo.get('name')}")
        reminder_win.geometry("480x320")
        reminder_win.resizable(False, False)
        reminder_win.transient(self.root)
        
        self.center_window(reminder_win, 480, 320)
        reminder_win.configure(bg='#FFFFE0')

        original_index = todo.get('original_index')
        
        def on_close():
            self.is_reminder_active = False
            if original_index is not None and original_index < len(self.todos):
                if self.todos[original_index].get('status') == '待处理':
                    self.todos[original_index]['status'] = '启用'
                    self.update_todo_list()
            reminder_win.destroy()

        reminder_win.protocol("WM_DELETE_WINDOW", on_close)

        if original_index is not None and original_index < len(self.todos):
            if self.todos[original_index]['status'] != '禁用':
                self.todos[original_index]['status'] = '待处理'
                self.update_todo_list()

        title_label = tk.Label(reminder_win, text=todo.get('name', '无标题'), font=('Microsoft YaHei', 14, 'bold'), bg='#FFFFE0', wraplength=460)
        title_label.pack(pady=(15, 10))

        content_frame = tk.Frame(reminder_win, bg='white', bd=1, relief='solid')
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        content_text = scrolledtext.ScrolledText(content_frame, font=('Microsoft YaHei', 11), wrap=tk.WORD, bd=0, bg='white')
        content_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        content_text.insert('1.0', todo.get('content', ''))
        content_text.config(state='disabled')

        btn_frame = tk.Frame(reminder_win, bg='#FFFFE0')
        btn_frame.pack(pady=15)
        font_spec = ('Microsoft YaHei', 11)
        
        def _handle_complete():
            if original_index is not None and original_index < len(self.todos):
                self.todos[original_index]['status'] = '禁用'
                self.save_todos()
                self.update_todo_list()
                self.log(f"待办事项 '{todo['name']}' 已标记为完成。")
            on_close()

        def _handle_snooze():
            minutes = simpledialog.askinteger("稍后提醒", "您想在多少分钟后再次提醒？ (1-60)", parent=reminder_win, minvalue=1, maxvalue=60, initialvalue=5)
            if minutes:
                new_remind_time = datetime.now() + timedelta(minutes=minutes)
                if original_index is not None and original_index < len(self.todos):
                    self.todos[original_index]['remind_datetime'] = new_remind_time.strftime('%Y-%m-%d %H:%M:%S')
                    self.todos[original_index]['status'] = '启用' 
                    self.save_todos()
                    self.update_todo_list()
                    self.log(f"待办事项 '{todo['name']}' 已推迟 {minutes} 分钟。")
            on_close()

        def _handle_delete():
            if messagebox.askyesno("确认删除", f"您确定要永久删除待办事项“{todo['name']}”吗？\n此操作不可恢复。", parent=reminder_win):
                if original_index is not None and original_index < len(self.todos):
                    if self.todos[original_index]['name'] == todo['name']:
                        self.todos.pop(original_index)
                        self.save_todos()
                        self.update_todo_list()
                        self.log(f"已删除待办事项: {todo['name']}")
                on_close()

        task_type = todo.get('type')
        if task_type == 'onetime':
            tk.Button(btn_frame, text="已完成", font=font_spec, bg='#27AE60', fg='white', width=10, command=_handle_complete).pack(side=tk.LEFT, padx=10)
            tk.Button(btn_frame, text="稍后提醒", font=font_spec, width=10, command=_handle_snooze).pack(side=tk.LEFT, padx=10)
            tk.Button(btn_frame, text="删除任务", font=font_spec, bg='#E74C3C', fg='white', width=10, command=_handle_delete).pack(side=tk.LEFT, padx=10)
        elif task_type == 'recurring':
            tk.Button(btn_frame, text="本次完成", font=font_spec, bg='#3498DB', fg='white', width=10, command=on_close).pack(side=tk.LEFT, padx=10)
            tk.Button(btn_frame, text="删除任务", font=font_spec, bg='#E74C3C', fg='white', width=10, command=_handle_delete).pack(side=tk.LEFT, padx=10)
        else:
            tk.Button(btn_frame, text="确定", font=font_spec, bg='#3498DB', fg='white', width=10, command=on_close).pack(side=tk.LEFT, padx=10)

        if self.notification_sound:
            self.notification_sound.play()
            
        reminder_win.lift()
        reminder_win.focus_force()
        self._flash_window()
        
        reminder_win.update_idletasks()
        reminder_win.update()


def main():
    root = tk.Tk()
    app = TimedBroadcastApp(root)
    root.mainloop()

if __name__ == "__main__":
    if not WIN32COM_AVAILABLE:
        messagebox.showerror("核心依赖缺失", "pywin32 库未安装或损坏，软件无法运行注册和锁定等核心功能，即将退出。")
        sys.exit()
    if not PSUTIL_AVAILABLE:
        messagebox.showerror("核心依赖缺失", "psutil 库未安装，软件无法获取机器码以进行授权验证，即将退出。")
        sys.exit()
    main()
