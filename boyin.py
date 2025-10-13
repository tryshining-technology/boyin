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
PROMPT_FOLDER = os.path.join(application_path, "提示音")
AUDIO_FOLDER = os.path.join(application_path, "音频文件")
BGM_FOLDER = os.path.join(application_path, "文稿背景")
VOICE_SCRIPT_FOLDER = os.path.join(application_path, "语音文稿")
ICON_FILE = resource_path("icon.ico")

CHIME_FOLDER = os.path.join(AUDIO_FOLDER, "整点报时")

REGISTRY_KEY_PATH = r"Software\创翔科技\TimedBroadcastApp"
REGISTRY_PARENT_KEY_PATH = r"Software\创翔科技"


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
        self.is_app_locked_down = False
        
        self.auth_info = {'status': 'Unregistered', 'message': '正在验证授权...'}
        self.machine_code = None

        self.lock_password_b64 = ""
        
        self.drag_start_item = None
        
        self.playback_command_queue = queue.Queue()
        
        self.pages = {}
        self.nav_buttons = {}
        self.current_page = None

        self.last_chime_hour = -1

        self.create_folder_structure()
        self.load_settings()
        self.load_lock_password()
        
        self.check_authorization()

        self.create_widgets()
        self.load_tasks()
        self.load_holidays()
        
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
        # --- 整体布局调整 ---
        # 1. 状态栏在底部
        self.status_frame = tk.Frame(self.root, bg='#E8F4F8', height=30)
        self.status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_frame.pack_propagate(False)
        self.create_status_bar_content()

        # 2. 导航栏在左侧
        self.nav_frame = tk.Frame(self.root, bg='#A8D8E8', width=160)
        self.nav_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.nav_frame.pack_propagate(False)
        
        # 3. 页面容器填充剩余空间
        self.page_container = tk.Frame(self.root)
        self.page_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # --- 填充导航栏 ---
        nav_button_titles = ["定时广播", "节假日", "设置", "注册软件", "超级管理"]
        
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
        
        # --- 创建初始页面 ---
        self.main_frame = tk.Frame(self.page_container, bg='white')
        self.pages["定时广播"] = self.main_frame
        self.create_scheduled_broadcast_page()

        self.current_page = self.main_frame
        self.switch_page("定时广播")
        
        self.update_status_bar()
        self.log("定时播音软件已启动")

    def create_status_bar_content(self):
        self.status_labels = []
        status_texts = ["当前时间", "系统状态", "播放状态", "任务数量"]
        font_11 = ('Microsoft YaHei', 11)

        copyright_label = tk.Label(self.status_frame, text="© 创翔科技", font=font_11,
                                   bg='#5DADE2', fg='white', padx=15)
        copyright_label.pack(side=tk.RIGHT, padx=2)
        
        # 新增的解锁按钮，初始不显示
        self.statusbar_unlock_button = tk.Button(self.status_frame, text="🔓 解锁", font=font_11,
                                                 bg='#2ECC71', fg='white', bd=0, padx=15, cursor='hand2',
                                                 command=self._prompt_for_password_unlock)
        # self.statusbar_unlock_button will be packed/unpacked in _apply_lock/unlock

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

        # 页面现在被打包到 page_container 中
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
        # 页面父容器修改
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
        
        # 新增取消注册按钮
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

        self.check_authorization()  # 重新评估授权状态

        messagebox.showinfo("操作完成", f"注册已成功取消。\n当前授权状态: {self.auth_info['message']}")
        self.log(f"注册已取消。新状态: {self.auth_info['message']}")
        
        if self.is_app_locked_down:
            # 如果试用期已过，则执行锁定流程
            self.perform_lockdown()
        else:
            # 如果仍在试用期，确保功能可用并可以切换到主页面
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
            self.log(f"找到有线网卡 MAC: {wired_macs[0]['mac']} (来自: {wired_macs[0]['name']})")
            return wired_macs[0]['mac']
        if wireless_macs:
            self.log(f"找到无线网卡 MAC: {wireless_macs[0]['mac']} (来自: {wireless_macs[0]['name']})")
            return wireless_macs[0]['mac']
        if other_macs:
            self.log(f"找到其他类型网卡 MAC: {other_macs[0]['mac']} (来自: {other_macs[0]['name']})")
            return other_macs[0]['mac']
            
        for name, addrs in interfaces.items():
            for addr in addrs:
                if addr.family == psutil.AF_LINK:
                    mac = addr.address.replace(':', '').replace('-', '').upper()
                    if len(mac) == 12 and mac != '000000000000':
                         self.log(f"备用方案找到 MAC: {mac} (来自: {name})")
                         return mac

        self.log("错误: 未能找到任何有效的网络适配器 MAC 地址。")
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
        self.root.title(f"定时播音 ({self.auth_info['message']})")
    
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
        
        # --- 新增卸载按钮 ---
        tk.Button(btn_frame, text="卸载软件", command=self._prompt_for_uninstall,
                  font=btn_font, width=btn_width, pady=btn_pady, bg='#34495E', fg='white').pack(pady=10)
                  
        return page_frame

    def _prompt_for_uninstall(self):
        """ 弹出卸载密码输入框 """
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

        # 密码是 YYYYMMDD 的反序
        correct_password = datetime.now().strftime('%Y%m%d')[::-1]
        
        if entered_password == correct_password:
            self.log("卸载密码正确，准备执行卸载操作。")
            self._perform_uninstall()
        elif entered_password is not None:
            messagebox.showerror("验证失败", "密码错误！", parent=self.root)
            self.log("尝试卸载软件失败：密码错误。")

    def _perform_uninstall(self):
        """ 执行实际的卸载（删除文件和注册表）操作 """
        if not messagebox.askyesno(
            "！！！最终警告！！！",
            "您确定要卸载本软件吗？\n\n此操作将永久删除：\n- 所有注册表信息\n- 所有配置文件 (节目单, 设置, 节假日)\n- 所有数据文件夹 (音频, 提示音, 文稿等)\n\n此操作【绝对无法恢复】！\n\n点击“是”将立即开始清理。",
            icon='error'
        ):
            self.log("用户取消了卸载操作。")
            return

        self.log("开始执行卸载流程...")
        self.running = False # 停止后台线程

        # 1. 删除注册表
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

        # 2. 删除文件夹
        folders_to_delete = [PROMPT_FOLDER, AUDIO_FOLDER, BGM_FOLDER, VOICE_SCRIPT_FOLDER]
        for folder in folders_to_delete:
            if os.path.isdir(folder):
                try:
                    shutil.rmtree(folder)
                    self.log(f"成功删除文件夹: {os.path.basename(folder)}")
                except Exception as e:
                    self.log(f"删除文件夹 {os.path.basename(folder)} 时出错: {e}")

        # 3. 删除JSON配置文件
        files_to_delete = [TASK_FILE, SETTINGS_FILE, HOLIDAY_FILE]
        for file in files_to_delete:
            if os.path.isfile(file):
                try:
                    os.remove(file)
                    self.log(f"成功删除文件: {os.path.basename(file)}")
                except Exception as e:
                    self.log(f"删除文件 {os.path.basename(file)} 时出错: {e}")
        
        # 4. 提示用户并强制退出
        self.log("软件数据清理完成。")
        messagebox.showinfo("卸载完成", "软件相关的数据和配置已全部清除。\n\n请手动删除本程序（.exe文件）以完成卸载。\n\n点击“确定”后软件将退出。")
        
        # --- 关键修改：使用 os._exit(0) 强制退出 ---
        # 这会立即终止程序，不会执行任何后续代码或清理操作（如 quit_app 中的保存），
        # 从而防止重新生成已删除的文件。
        os._exit(0)

    def _backup_all_settings(self):
        self.log("开始备份所有设置...")
        try:
            backup_data = {
                'backup_date': datetime.now().isoformat(), 'tasks': self.tasks, 'holidays': self.holidays,
                'settings': self.settings, 'lock_password_b64': self._load_from_registry("LockPasswordB64")
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
            self.settings = backup_data['settings']
            self.lock_password_b64 = backup_data['lock_password_b64']
            
            self.save_tasks()
            self.save_holidays()
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
            
            if self.lock_password_b64:
                self._save_to_registry("LockPasswordB64", self.lock_password_b64)
            else:
                self._save_to_registry("LockPasswordB64", "")

            self.update_task_list()
            self.update_holiday_list()
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
            "您真的要重置整个软件吗？\n\n此操作将：\n- 清空所有节目单 (但保留音频文件)\n- 清空所有节假日\n- 清除锁定密码\n- 重置所有系统设置\n\n此操作【无法恢复】！软件将在重置后提示您重启。"
        ): return

        self.log("开始执行软件重置...")
        try:
            original_askyesno = messagebox.askyesno
            messagebox.askyesno = lambda title, message: True
            self.clear_all_tasks(delete_associated_files=False)
            self.clear_all_holidays()
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
        # 父容器是 self.main_frame
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

        # 状态栏相关代码已移至 create_widgets 和 create_status_bar_content
    
    def create_holiday_page(self):
        # 页面父容器修改
        page_frame = tk.Frame(self.page_container, bg='white')

        top_frame = tk.Frame(page_frame, bg='white')
        top_frame.pack(fill=tk.X, padx=10, pady=10)
        title_label = tk.Label(top_frame, text="节假日", font=('Microsoft YaHei', 14, 'bold'),
                              bg='white', fg='#2C5F7C')
        title_label.pack(side=tk.LEFT)
        
        desc_label = tk.Label(page_frame, text="节假日不播放 (手动和立即播任务除外)，整点报时也受此约束", font=('Microsoft YaHei', 11),
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
        # 页面父容器修改
        settings_frame = tk.Frame(self.page_container, bg='white')

        title_label = tk.Label(settings_frame, text="系统设置", font=('Microsoft YaHei', 14, 'bold'),
                               bg='white', fg='#2C5F7C')
        title_label.pack(anchor='w', padx=20, pady=20)

        general_frame = tk.LabelFrame(settings_frame, text="通用设置", font=('Microsoft YaHei', 12, 'bold'),
                                      bg='white', padx=15, pady=10)
        general_frame.pack(fill=tk.X, padx=20, pady=10)
        
        self.autostart_var = tk.BooleanVar()
        self.start_minimized_var = tk.BooleanVar()
        self.lock_on_start_var = tk.BooleanVar()

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
        if not WIN32COM_AVAILABLE:
            self.lock_on_start_cb.config(state=tk.DISABLED)
        
        tk.Label(lock_frame, text="(请先在主界面设置锁定密码)", font=('Microsoft YaHei', 9), bg='white', fg='grey').pack(side=tk.LEFT, padx=5)

        self.clear_password_btn = tk.Button(general_frame, text="清除锁定密码", font=('Microsoft YaHei', 11), command=self.clear_lock_password)
        self.clear_password_btn.pack(pady=10)
        
        time_chime_frame = tk.LabelFrame(settings_frame, text="整点报时", font=('Microsoft YaHei', 12, 'bold'),
                                         bg='white', padx=15, pady=10)
        time_chime_frame.pack(fill=tk.X, padx=20, pady=10)
        
        self.time_chime_enabled_var = tk.BooleanVar()
        self.time_chime_voice_var = tk.StringVar()
        self.time_chime_speed_var = tk.StringVar()
        self.time_chime_pitch_var = tk.StringVar()
        
        chime_control_frame = tk.Frame(time_chime_frame, bg='white')
        chime_control_frame.pack(fill=tk.X, pady=5)

        tk.Checkbutton(chime_control_frame, text="启用整点报时功能", variable=self.time_chime_enabled_var, 
                       font=('Microsoft YaHei', 11), bg='white', anchor='w',
                       command=self._handle_time_chime_toggle).pack(side=tk.LEFT)

        available_voices = self.get_available_voices()
        self.chime_voice_combo = ttk.Combobox(chime_control_frame, textvariable=self.time_chime_voice_var, 
                                              values=available_voices, font=('Microsoft YaHei', 10), 
                                              width=35, state='readonly')
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

        power_frame = tk.LabelFrame(settings_frame, text="电源管理", font=('Microsoft YaHei', 12, 'bold'),
                                    bg='white', padx=15, pady=10)
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
        tk.Checkbutton(daily_frame, text="每天关机", variable=self.daily_shutdown_enabled_var, 
                       font=('Microsoft YaHei', 11), bg='white', command=self.save_settings).pack(side=tk.LEFT)
        tk.Entry(daily_frame, textvariable=self.daily_shutdown_time_var, 
                 font=('Microsoft YaHei', 11), width=15).pack(side=tk.LEFT, padx=10)
        tk.Button(daily_frame, text="设置", font=('Microsoft YaHei', 11), command=lambda: self.show_single_time_dialog(self.daily_shutdown_time_var)
                  ).pack(side=tk.LEFT)

        weekly_frame = tk.Frame(power_frame, bg='white')
        weekly_frame.pack(fill=tk.X, pady=4)
        tk.Checkbutton(weekly_frame, text="每周关机", variable=self.weekly_shutdown_enabled_var, 
                       font=('Microsoft YaHei', 11), bg='white', command=self.save_settings).pack(side=tk.LEFT)
        tk.Entry(weekly_frame, textvariable=self.weekly_shutdown_days_var,
                 font=('Microsoft YaHei', 11), width=20).pack(side=tk.LEFT, padx=(10,5))
        tk.Entry(weekly_frame, textvariable=self.weekly_shutdown_time_var,
                 font=('Microsoft YaHei', 11), width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(weekly_frame, text="设置", font=('Microsoft YaHei', 11), command=lambda: self.show_power_week_time_dialog(
            "设置每周关机", self.weekly_shutdown_days_var, self.weekly_shutdown_time_var)).pack(side=tk.LEFT)

        reboot_frame = tk.Frame(power_frame, bg='white')
        reboot_frame.pack(fill=tk.X, pady=4)
        tk.Checkbutton(reboot_frame, text="每周重启", variable=self.weekly_reboot_enabled_var,
                       font=('Microsoft YaHei', 11), bg='white', command=self.save_settings).pack(side=tk.LEFT)
        tk.Entry(reboot_frame, textvariable=self.weekly_reboot_days_var,
                 font=('Microsoft YaHei', 11), width=20).pack(side=tk.LEFT, padx=(10,5))
        tk.Entry(reboot_frame, textvariable=self.weekly_reboot_time_var,
                 font=('Microsoft YaHei', 11), width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(reboot_frame, text="设置", font=('Microsoft YaHei', 11), command=lambda: self.show_power_week_time_dialog(
            "设置每周重启", self.weekly_reboot_days_var, self.weekly_reboot_time_var)).pack(side=tk.LEFT)

        return settings_frame

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
            self.save_settings() # 保存新值
            if messagebox.askyesno("应用更改", "您更改了报时参数，需要重新生成全部24个报时文件。\n是否立即开始？"):
                self._handle_time_chime_toggle(force_regenerate=True)
            else:
                # 用户取消，恢复到之前的设置
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
        # 在状态栏显示解锁按钮
        self.statusbar_unlock_button.pack(side=tk.RIGHT, padx=5)
        self.log("界面已锁定。")

    def _apply_unlock(self):
        self.is_locked = False
        self.lock_button.config(text="锁定", bg='#E74C3C')
        self._set_ui_lock_state(tk.NORMAL)
        # 从状态栏隐藏解锁按钮
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
        dialog.geometry("950x680")
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
        def select_folder():
            foldername = filedialog.askdirectory(title="选择音频文件夹", initialdir=AUDIO_FOLDER)
            if foldername: audio_folder_entry.delete(0, tk.END); audio_folder_entry.insert(0, foldername)
        tk.Button(audio_folder_frame, text="选取...", command=select_folder, bg='#D0D0D0', font=font_spec, bd=1, padx=22, pady=1).pack(side=tk.LEFT, padx=5)
        play_order_frame = tk.Frame(content_frame, bg='#E8E8E8')
        play_order_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        play_order_var = tk.StringVar(value="sequential")
        tk.Radiobutton(play_order_frame, text="顺序播", variable=play_order_var, value="sequential", bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(play_order_frame, text="随机播", variable=play_order_var, value="random", bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT, padx=10)
        volume_frame = tk.Frame(content_frame, bg='#E8E8E8')
        volume_frame.grid(row=4, column=1, columnspan=3, sticky='w', padx=5, pady=3)
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
        
        # --- MODIFICATION START: Buttons are now part of the "Other" frame ---
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
        # --- MODIFICATION END ---

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
                self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))
        
        # --- MODIFICATION START: Buttons are created in the new frame ---
        button_text = "保存修改" if is_edit_mode else "添加"
        tk.Button(dialog_button_frame, text=button_text, command=save_task, bg='#5DADE2', fg='white', font=('Microsoft YaHei', 11, 'bold'), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        tk.Button(dialog_button_frame, text="取消", command=dialog.destroy, bg='#D0D0D0', font=('Microsoft YaHei', 11), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        # --- MODIFICATION END ---
        
        content_frame.columnconfigure(1, weight=1); time_frame.columnconfigure(1, weight=1)

    def open_voice_dialog(self, parent_dialog, task_to_edit=None, index=None):
        parent_dialog.destroy()
        is_edit_mode = task_to_edit is not None
        dialog = tk.Toplevel(self.root)
        dialog.title("修改语音节目" if is_edit_mode else "添加语音节目")
        dialog.geometry("950x680")
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
        prompt_frame.grid(row=5, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        tk.Checkbutton(prompt_frame, text="提示音:", variable=prompt_var, bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT)
        prompt_file_var, prompt_volume_var = tk.StringVar(), tk.StringVar()
        prompt_file_entry = tk.Entry(prompt_frame, textvariable=prompt_file_var, font=font_spec, width=20)
        prompt_file_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(prompt_frame, text="...", command=lambda: self.select_file_for_entry(PROMPT_FOLDER, prompt_file_var)).pack(side=tk.LEFT)
        tk.Label(prompt_frame, text="音量(0-100):", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=(10,5))
        tk.Entry(prompt_frame, textvariable=prompt_volume_var, font=font_spec, width=8).pack(side=tk.LEFT, padx=5)
        
        bgm_var = tk.IntVar()
        bgm_frame = tk.Frame(content_frame, bg='#E8E8E8')
        bgm_frame.grid(row=6, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        tk.Checkbutton(bgm_frame, text="背景音乐:", variable=bgm_var, bg='#E8E8E8', font=font_spec).pack(side=tk.LEFT)
        bgm_file_var, bgm_volume_var = tk.StringVar(), tk.StringVar()
        bgm_file_entry = tk.Entry(bgm_frame, textvariable=bgm_file_var, font=font_spec, width=20)
        bgm_file_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(bgm_frame, text="...", command=lambda: self.select_file_for_entry(BGM_FOLDER, bgm_file_var)).pack(side=tk.LEFT)
        tk.Label(bgm_frame, text="音量(0-100):", font=font_spec, bg='#E8E8E8').pack(side=tk.LEFT, padx=(10,5))
        tk.Entry(bgm_frame, textvariable=bgm_volume_var, font=font_spec, width=8).pack(side=tk.LEFT, padx=5)
        
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
        
        # --- MODIFICATION START: Buttons are now part of the "Other" frame ---
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
        # --- MODIFICATION END ---

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

        def save_task():
            text_content = content_text.get('1.0', tk.END).strip()
            if not text_content:
                messagebox.showwarning("警告", "请输入播音文字内容", parent=dialog)
                return

            is_valid_time, time_msg = self._normalize_multiple_times_string(start_time_entry.get().strip())
            if not is_valid_time:
                messagebox.showwarning("格式错误", time_msg, parent=dialog)
                return
            is_valid_date, date_msg = self._normalize_date_range_string(date_range_entry.get().strip())
            if not is_valid_date:
                messagebox.showwarning("格式错误", date_msg, parent=dialog)
                return

            regeneration_needed = True
            if is_edit_mode:
                original_task = task_to_edit
                if (text_content == original_task.get('source_text') and
                    voice_var.get() == original_task.get('voice') and
                    speed_entry.get().strip() == original_task.get('speed', '0') and
                    pitch_entry.get().strip() == original_task.get('pitch', '0') and
                    volume_entry.get().strip() == original_task.get('volume', '80')):
                    regeneration_needed = False
                    self.log("语音内容未变更，跳过重新生成WAV文件。")

            def build_task_data(wav_path, wav_filename_str):
                play_mode = delay_var.get()
                saved_delay_type = task_to_edit.get('delay', 'delay') if is_edit_mode else play_mode
                
                return {
                    'name': name_entry.get().strip(), 'time': time_msg, 'type': 'voice', 
                    'content': wav_path, 'wav_filename': wav_filename_str, 
                    'source_text': text_content, 'voice': voice_var.get(), 
                    'speed': speed_entry.get().strip() or "0", 'pitch': pitch_entry.get().strip() or "0", 
                    'volume': volume_entry.get().strip() or "80", 
                    'prompt': prompt_var.get(), 'prompt_file': prompt_file_var.get(), 
                    'prompt_volume': prompt_volume_var.get(), 'bgm': bgm_var.get(), 
                    'bgm_file': bgm_file_var.get(), 'bgm_volume': bgm_volume_var.get(), 
                    'repeat': repeat_entry.get().strip() or "1", 'weekday': weekday_entry.get().strip(), 
                    'date_range': date_msg, 'delay': saved_delay_type, 
                    'status': '启用' if not is_edit_mode else task_to_edit.get('status', '启用'), 
                    'last_run': {} if not is_edit_mode else task_to_edit.get('last_run', {})
                }

            if not regeneration_needed:
                new_task_data = build_task_data(
                    task_to_edit.get('content'), task_to_edit.get('wav_filename')
                )
                if not new_task_data['name'] or not new_task_data['time']:
                    messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog); return
                
                self.tasks[index] = new_task_data
                self.log(f"已修改语音节目(未重新生成语音): {new_task_data['name']}")
                self.update_task_list(); self.save_tasks(); dialog.destroy()
                
                if delay_var.get() == 'immediate':
                     self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))
                return

            progress_dialog = tk.Toplevel(dialog)
            progress_dialog.title("请稍候")
            progress_dialog.geometry("300x100")
            progress_dialog.resizable(False, False)
            progress_dialog.transient(dialog); progress_dialog.grab_set()
            tk.Label(progress_dialog, text="语音文件生成中，请稍后...", font=font_spec).pack(expand=True)
            self.center_window(progress_dialog, 300, 100)
            dialog.update_idletasks()
            
            new_wav_filename = f"{int(time.time())}_{random.randint(1000, 9999)}.wav"
            output_path = os.path.join(AUDIO_FOLDER, new_wav_filename)
            voice_params = {
                'voice': voice_var.get(), 'speed': speed_entry.get().strip() or "0", 
                'pitch': pitch_entry.get().strip() or "0", 'volume': volume_entry.get().strip() or "80"
            }

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

                new_task_data = build_task_data(output_path, new_wav_filename)
                if not new_task_data['name'] or not new_task_data['time']:
                    messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog); return
                
                if is_edit_mode:
                    self.tasks[index] = new_task_data
                    self.log(f"已修改语音节目(并重新生成语音): {new_task_data['name']}")
                else:
                    self.tasks.append(new_task_data)
                    self.log(f"已添加语音节目: {new_task_data['name']}")
                
                self.update_task_list(); self.save_tasks(); dialog.destroy()

                if delay_var.get() == 'immediate':
                    self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))

            synthesis_thread = threading.Thread(target=self._synthesis_worker, 
                                                args=(text_content, voice_params, output_path, _on_synthesis_complete))
            synthesis_thread.daemon = True
            synthesis_thread.start()
        
        # --- MODIFICATION START: Buttons are created in the new frame ---
        button_text = "保存修改" if is_edit_mode else "添加"
        tk.Button(dialog_button_frame, text=button_text, command=save_task, bg='#5DADE2', fg='white', font=('Microsoft YaHei', 11, 'bold'), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        tk.Button(dialog_button_frame, text="取消", command=dialog.destroy, bg='#D0D0D0', font=('Microsoft YaHei', 11), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        # --- MODIFICATION END ---

        content_frame.columnconfigure(1, weight=1); time_frame.columnconfigure(1, weight=1)

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
            text_widget.delete('1.0', tk.END)
            text_widget.insert('1.0', content)
            self.log(f"已从 {os.path.basename(filename)} 成功导入文稿。")
        except Exception as e:
            messagebox.showerror("导入失败", f"无法读取文件：\n{e}")
            self.log(f"导入文稿失败: {e}")

    def _export_voice_script(self, text_widget, name_widget):
        content = text_widget.get('1.0', tk.END).strip()
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
        
        type_name = "音频" if task_type == 'audio' else "语音"
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

    def update_task_list(self):
        if not hasattr(self, 'task_tree') or not self.task_tree.winfo_exists(): return
        selection = self.task_tree.selection()
        self.task_tree.delete(*self.task_tree.get_children())
        for task in self.tasks:
            content = task.get('content', '')
            if task.get('type') == 'voice':
                source_text = task.get('source_text', '')
                clean_content = source_text.replace('\n', ' ').replace('\r', '')
                content_preview = (clean_content[:30] + '...') if len(clean_content) > 30 else clean_content
            else:
                content_preview = os.path.basename(content)
                
            display_mode = "准时" if task.get('delay') == 'ontime' else "延时"
            self.task_tree.insert('', tk.END, values=(task.get('name', ''), task.get('status', ''), task.get('time', ''), display_mode, content_preview, task.get('volume', ''), task.get('weekday', ''), task.get('date_range', '')))
        if selection:
            try: 
                valid_selection = [s for s in selection if self.task_tree.exists(s)]
                if valid_selection: self.task_tree.selection_set(valid_selection)
            except tk.TclError: pass
        self.stats_label.config(text=f"节目单：{len(self.tasks)}")
        if hasattr(self, 'status_labels'): self.status_labels[3].config(text=f"任务数量: {len(self.tasks)}")

    def update_status_bar(self):
        if not self.running: return
        self.status_labels[0].config(text=f"当前时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.status_labels[1].config(text="系统状态: 运行中")
        self.root.after(1000, self.update_status_bar)

    def start_background_threads(self):
        threading.Thread(target=self._scheduler_worker, daemon=True).start()
        threading.Thread(target=self._playback_worker, daemon=True).start()

    def _scheduler_worker(self):
        while self.running:
            now = datetime.now()
            if not self.is_app_locked_down:
                self._check_broadcast_tasks(now)
                self._check_time_chime(now)
            
            self._check_power_tasks(now)
            time.sleep(1)
    
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
        current_date_str = now.strftime("%Y-%m-%d")
        current_time_str = now.strftime("%H:%M:%S")

        is_holiday_now = self._is_in_holiday(now)
        
        tasks_to_play = []

        for task in self.tasks:
            if task.get('status') != '启用': continue
            try:
                start, end = [d.strip() for d in task.get('date_range', '').split('~')]
                if not (datetime.strptime(start, "%Y-%m-%d").date() <= now.date() <= datetime.strptime(end, "%Y-%m-%d").date()): continue
            except (ValueError, IndexError): pass
            schedule = task.get('weekday', '每周:1234567')
            run_today = (schedule.startswith("每周:") and str(now.isoweekday()) in schedule[3:]) or (schedule.startswith("每月:") and f"{now.day:02d}" in schedule[3:].split(','))
            if not run_today: continue
            
            for trigger_time in [t.strip() for t in task.get('time', '').split(',')]:
                if trigger_time == current_time_str and task.get('last_run', {}).get(trigger_time) != current_date_str:
                    
                    if is_holiday_now:
                        self.log(f"任务 '{task['name']}' 因处于节假日期间而被跳过。")
                        continue 
                    
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
                was_playing = pygame.mixer.get_busy()
                if was_playing:
                    pygame.mixer.pause()
                    self.log("整点报时，暂停当前播放...")
                
                try:
                    chime_sound = pygame.mixer.Sound(chime_path)
                    chime_sound.set_volume(1.0)
                    chime_channel = chime_sound.play()
                    while chime_channel and chime_channel.get_busy():
                        time.sleep(0.1)
                except Exception as e:
                    self.log(f"播放整点报时失败: {e}")

                if was_playing:
                    pygame.mixer.unpause()
                    self.log("报时结束，恢复播放。")

            elif command == 'STOP':
                is_playing = False
                if AUDIO_AVAILABLE:
                    pygame.mixer.music.stop()
                    pygame.mixer.stop()
                self.log("STOP 命令已处理，所有播放已停止。")
                self.update_playing_text("等待播放...")
                self.status_labels[2].config(text="播放状态: 待机")
                while not self.playback_command_queue.empty():
                    try: self.playback_command_queue.get_nowait()
                    except queue.Empty: break
    
    def _execute_broadcast(self, task, trigger_time):
        self.update_playing_text(f"[{task['name']}] 正在准备播放...")
        self.status_labels[2].config(text="播放状态: 播放中")
        
        if trigger_time != "manual_play":
            task.setdefault('last_run', {})[trigger_time] = datetime.now().strftime("%Y-%m-%d")
            self.save_tasks()
        
        try:
            if task.get('type') == 'audio':
                self.log(f"开始音频任务: {task['name']}")
                self._play_audio_task_internal(task)
            elif task.get('type') == 'voice':
                self.log(f"开始语音任务: {task['name']} (共 {task.get('repeat', 1)} 遍)")
                self._play_voice_task_internal(task)
        except Exception as e:
            self.log(f"播放任务 '{task['name']}' 时发生严重错误: {e}")
        finally:
            if AUDIO_AVAILABLE:
                pygame.mixer.music.stop()
                pygame.mixer.stop()
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
        for audio_path in playlist:
            if self._is_interrupted():
                self.log(f"任务 '{task['name']}' 被新指令中断。")
                return
            
            self.log(f"正在播放: {os.path.basename(audio_path)}")
            self.update_playing_text(f"[{task['name']}] 正在播放: {os.path.basename(audio_path)}")
            
            try:
                pygame.mixer.music.load(audio_path)
                pygame.mixer.music.set_volume(float(task.get('volume', 80)) / 100.0)
                pygame.mixer.music.play()

                while pygame.mixer.music.get_busy():
                    if self._is_interrupted():
                        pygame.mixer.music.stop()
                        return
                    if interval_type == 'seconds' and (time.time() - start_time) >= duration_seconds:
                        pygame.mixer.music.stop()
                        self.log(f"已达到 {duration_seconds} 秒播放时长限制。")
                        return
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
            prompt_file = task.get('prompt_file', '')
            prompt_path = os.path.join(PROMPT_FOLDER, prompt_file)
            if os.path.exists(prompt_path):
                try:
                    self.log(f"播放提示音: {prompt_file}")
                    sound = pygame.mixer.Sound(prompt_path)
                    sound.set_volume(float(task.get('prompt_volume', 80)) / 100.0)
                    channel = sound.play()
                    while channel and channel.get_busy():
                        if self._is_interrupted(): return
                        time.sleep(0.05)
                except Exception as e:
                    self.log(f"播放提示音失败: {e}")
            else:
                self.log(f"警告: 提示音文件不存在 - {prompt_path}")

        if task.get('bgm', 0):
            if self._is_interrupted(): return
            bgm_file = task.get('bgm_file', '')
            bgm_path = os.path.join(BGM_FOLDER, bgm_file)
            if os.path.exists(bgm_path):
                try:
                    self.log(f"播放背景音乐: {bgm_file}")
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

            for i in range(repeat_count):
                if self._is_interrupted(): return
                
                self.log(f"正在播报第 {i+1}/{repeat_count} 遍")
                self.update_playing_text(f"[{task['name']}] 正在播报第 {i+1}/{repeat_count} 遍...")
                
                channel = speech_sound.play()
                while channel and channel.get_busy():
                    if self._is_interrupted():
                        channel.stop()
                        return
                    time.sleep(0.1)
                
                if i < repeat_count - 1:
                    time.sleep(0.5)
        except Exception as e:
            self.log(f"播放语音内容失败: {e}")

    def log(self, message): self.root.after(0, lambda: self._log_threadsafe(message))
    def _log_threadsafe(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} -> {message}\n")
        self.log_text.see(tk.END); self.log_text.config(state='disabled')

    def update_playing_text(self, message): self.root.after(0, lambda: self._update_playing_text_threadsafe(message))
    def _update_playing_text_threadsafe(self, message):
        self.playing_text.config(state='normal')
        self.playing_text.delete('1.0', tk.END); self.playing_text.insert('1.0', message)
        self.playing_text.config(state='disabled')

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
            "autostart": False, "start_minimized": False, "lock_on_start": False, 
            "daily_shutdown_enabled": False, "daily_shutdown_time": "23:00:00", 
            "weekly_shutdown_enabled": False, "weekly_shutdown_days": "每周:12345", "weekly_shutdown_time": "23:30:00", 
            "weekly_reboot_enabled": False, "weekly_reboot_days": "每周:67", "weekly_reboot_time": "22:00:00", 
            "last_power_action_date": "",
            "time_chime_enabled": False, "time_chime_voice": "",
            "time_chime_speed": "0", "time_chime_pitch": "0"
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
            self.settings.update({
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
                "time_chime_pitch": self.time_chime_pitch_var.get()
            })
        try:
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f: json.dump(self.settings, f, ensure_ascii=False, indent=2)
        except Exception as e: self.log(f"保存设置失败: {e}")

    def _handle_autostart_setting(self):
        self.save_settings()
        enable = self.autostart_var.get()
        if not WIN32COM_AVAILABLE:
            self.log("错误: 自动启动功能需要 pywin32 库。")
            if enable: self.autostart_var.set(False); self.save_settings()
            messagebox.showerror("功能受限", "未安装 pywin32 库，无法设置开机启动。")
            return
        shortcut_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup', "定时播音.lnk")
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
            messagebox.showerror("错误", f"操作失败: {e}")

    def center_window(self, win, width, height):
        x = (win.winfo_screenwidth() - width) // 2
        y = (win.winfo_screenheight() - height) // 2
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
        return True, ", ".join(normalized_times)

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
        dialog = tk.Toplevel(self.root)
        dialog.title("确认")
        dialog.geometry("380x170")
        dialog.resizable(False, False); dialog.transient(self.root); dialog.grab_set()
        self.center_window(dialog, 380, 170)
        font_spec = ('Microsoft YaHei', 11)
        tk.Label(dialog, text="您想要如何操作？", font=(font_spec[0], 12), pady=20).pack()
        btn_frame = tk.Frame(dialog); btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="退出程序", command=lambda: [dialog.destroy(), self.quit_app()], font=font_spec).pack(side=tk.LEFT, padx=10)
        if TRAY_AVAILABLE: tk.Button(btn_frame, text="最小化到托盘", command=lambda: [dialog.destroy(), self.hide_to_tray()], font=font_spec).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="取消", command=dialog.destroy, font=font_spec).pack(side=tk.LEFT, padx=10)

    def hide_to_tray(self):
        if not TRAY_AVAILABLE: messagebox.showwarning("功能不可用", "pystray 或 Pillow 库未安装，无法最小化到托盘。"); return
        self.root.withdraw()
        self.log("程序已最小化到系统托盘。")

    def show_from_tray(self, icon, item):
        self.root.after(0, self.root.deiconify)
        self.log("程序已从托盘恢复。")

    def quit_app(self, icon=None, item=None):
        if self.tray_icon: self.tray_icon.stop()
        self.running = False
        self.playback_command_queue.put(('STOP', None))
        
        # 在正常退出时才保存
        self.save_tasks()
        self.save_settings()
        self.save_holidays()
        
        if AUDIO_AVAILABLE and pygame.mixer.get_init(): pygame.mixer.quit()
        self.root.destroy()
        sys.exit()

    def setup_tray_icon(self):
        try: image = Image.open(ICON_FILE)
        except Exception as e: image = Image.new('RGB', (64, 64), 'white'); print(f"警告: 未找到或无法加载图标文件 '{ICON_FILE}': {e}")
        
        # 定义统一的菜单，不再根据锁定状态切换
        menu = (
            item('显示', self.show_from_tray, default=True),
            item('退出', self.quit_app)
        )
        
        self.tray_icon = Icon("boyin", image, "定时播音", menu)

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

    def update_holiday_list(self):
        if not hasattr(self, 'holiday_tree') or not self.holiday_tree.winfo_exists(): return
        selection = self.holiday_tree.selection()
        self.holiday_tree.delete(*self.holiday_tree.get_children())
        for holiday in self.holidays:
            self.holiday_tree.insert('', tk.END, values=(
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
            messagebox.showwarning("警告", "请先选择要修改的节假日")
            return
        index = self.holiday_tree.index(selection[0])
        holiday_to_edit = self.holidays[index]
        self.open_holiday_dialog(holiday_to_edit=holiday_to_edit, index=index)

    def delete_holiday(self):
        selections = self.holiday_tree.selection()
        if not selections:
            messagebox.showwarning("警告", "请先选择要删除的节假日")
            return
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(selections)} 个节假日吗？"):
            indices = sorted([self.holiday_tree.index(s) for s in selections], reverse=True)
            for index in indices:
                self.holidays.pop(index)
            self.update_holiday_list()
            self.save_holidays()

    def _set_holiday_status(self, status):
        selection = self.holiday_tree.selection()
        if not selection:
            messagebox.showwarning("警告", f"请先选择要{status}的节假日")
            return
        for item_id in selection:
            index = self.holiday_tree.index(item_id)
            self.holidays[index]['status'] = status
        self.update_holiday_list()
        self.save_holidays()

    def open_holiday_dialog(self, holiday_to_edit=None, index=None):
        dialog = tk.Toplevel(self.root)
        dialog.title("修改节假日" if holiday_to_edit else "添加节假日")
        dialog.geometry("500x300"); dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set(); dialog.configure(bg='#F0F8FF')
        self.center_window(dialog, 500, 300)

        font_spec = ('Microsoft YaHei', 11)
        main_frame = tk.Frame(dialog, bg='#F0F8FF', padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(main_frame, text="名称:", font=font_spec, bg='#F0F8FF').grid(row=0, column=0, sticky='w', pady=5)
        name_entry = tk.Entry(main_frame, font=font_spec, width=40)
        name_entry.grid(row=0, column=1, columnspan=2, sticky='ew', pady=5)

        tk.Label(main_frame, text="开始时间:", font=font_spec, bg='#F0F8FF').grid(row=1, column=0, sticky='w', pady=5)
        start_date_entry = tk.Entry(main_frame, font=font_spec, width=15)
        start_date_entry.grid(row=1, column=1, sticky='w', pady=5)
        start_time_entry = tk.Entry(main_frame, font=font_spec, width=15)
        start_time_entry.grid(row=1, column=2, sticky='w', pady=5, padx=5)

        tk.Label(main_frame, text="结束时间:", font=font_spec, bg='#F0F8FF').grid(row=2, column=0, sticky='w', pady=5)
        end_date_entry = tk.Entry(main_frame, font=font_spec, width=15)
        end_date_entry.grid(row=2, column=1, sticky='w', pady=5)
        end_time_entry = tk.Entry(main_frame, font=font_spec, width=15)
        end_time_entry.grid(row=2, column=2, sticky='w', pady=5, padx=5)
        
        tk.Label(main_frame, text="格式: YYYY-MM-DD", font=('Microsoft YaHei', 9), bg='#F0F8FF', fg='grey').grid(row=3, column=1, sticky='n')
        tk.Label(main_frame, text="格式: HH:MM:SS", font=('Microsoft YaHei', 9), bg='#F0F8FF', fg='grey').grid(row=3, column=2, sticky='n')

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
            dialog.destroy()

        button_frame = tk.Frame(main_frame, bg='#F0F8FF')
        button_frame.grid(row=4, column=0, columnspan=3, pady=20)
        tk.Button(button_frame, text="保存", command=save, font=font_spec, width=10).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="取消", command=dialog.destroy, font=font_spec, width=10).pack(side=tk.LEFT, padx=10)

    def show_holiday_context_menu(self, event):
        if self.is_locked: return
        iid = self.holiday_tree.identify_row(event.y)
        if not iid: return

        context_menu = tk.Menu(self.root, tearoff=0, font=('Microsoft YaHei', 11))
        
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
        filename = filedialog.askopenfilename(title="选择导入节假日文件", filetypes=[("JSON文件", "*.json")], initialdir=application_path)
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f: imported = json.load(f)

                if not isinstance(imported, list) or \
                   (imported and (not isinstance(imported[0], dict) or 'start_datetime' not in imported[0] or 'end_datetime' not in imported[0])):
                    messagebox.showerror("导入失败", "文件格式不正确，看起来不是一个有效的节假日备份文件。")
                    self.log(f"尝试导入格式错误的节假日文件: {os.path.basename(filename)}")
                    return

                self.holidays.extend(imported)
                self.update_holiday_list(); self.save_holidays()
                self.log(f"已从 {os.path.basename(filename)} 导入 {len(imported)} 个节假日")
            except Exception as e:
                messagebox.showerror("错误", f"导入失败: {e}")

    def export_holidays(self):
        if not self.holidays:
            messagebox.showwarning("警告", "没有节假日可以导出")
            return
        filename = filedialog.asksaveasfilename(title="导出节假日到...", defaultextension=".json",
                                              initialfile="holidays_backup.json", filetypes=[("JSON文件", "*.json")], initialdir=application_path)
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(self.holidays, f, ensure_ascii=False, indent=2)
                self.log(f"已导出 {len(self.holidays)} 个节假日到 {os.path.basename(filename)}")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {e}")
    
    def clear_all_holidays(self):
        if not self.holidays:
            return
        if messagebox.askyesno("严重警告", "您确定要清空所有节假日吗？\n此操作不可恢复！"):
            self.holidays.clear()
            self.update_holiday_list()
            self.save_holidays()
            self.log("已清空所有节假日。")

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
