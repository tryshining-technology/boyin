import customtkinter as ctk
from tkinter import ttk, messagebox, filedialog
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
import ctypes

# --- CustomTkinter 设置 ---
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

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
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

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
        self.root.title(" 创翔多功能定时播音旗舰版")
        self.root.geometry("1400x800")
        
        try:
            dpi = ctypes.windll.user32.GetDpiForWindow(self.root.winfo_id())
            scaling_factor = dpi / 96
            self.root.tk.call('tk', 'scaling', scaling_factor)
            ctk.set_widget_scaling(scaling_factor)
            ctk.set_window_scaling(scaling_factor)
        except (AttributeError, OSError):
            pass

        if os.path.exists(ICON_FILE):
            try:
                self.root.iconbitmap(ICON_FILE)
            except Exception as e:
                print(f"加载窗口图标失败: {e}")

        self.font_nav = ctk.CTkFont(family="Microsoft YaHei", size=22, weight="bold")
        self.font_bold = ctk.CTkFont(family="Microsoft YaHei", size=12, weight="bold")
        self.font_normal = ctk.CTkFont(family="Microsoft YaHei", size=12)
        self.font_small = ctk.CTkFont(family="Microsoft YaHei", size=12)
        self.font_log = ctk.CTkFont(family="Microsoft YaHei", size=12)

        self.tasks, self.holidays, self.settings = [], [], {}
        self.running, self.is_locked, self.is_app_locked_down = True, False, False
        self.tray_icon, self.machine_code, self.drag_start_item = None, None, None
        self.lock_password_b64 = ""
        self.auth_info = {'status': 'Unregistered', 'message': '正在验证授权...'}
        self.playback_command_queue = queue.Queue()
        self.pages, self.nav_buttons = {}, {}
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
        
        if self.settings.get("lock_on_start", False) and self.lock_password_b64: self.root.after(100, self.perform_initial_lock)
        if self.settings.get("start_minimized", False): self.root.after(100, self.hide_to_tray)
        if self.is_app_locked_down: self.root.after(100, self.perform_lockdown)

    def _save_to_registry(self, key_name, value):
        if not WIN32COM_AVAILABLE: return False
        try:
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, REGISTRY_KEY_PATH)
            winreg.SetValueEx(key, key_name, 0, winreg.REG_SZ, str(value))
            winreg.CloseKey(key); return True
        except Exception as e: self.log(f"错误: 无法写入注册表项 '{key_name}' - {e}"); return False

    def _load_from_registry(self, key_name):
        if not WIN32COM_AVAILABLE: return None
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REGISTRY_KEY_PATH, 0, winreg.KEY_READ)
            value, _ = winreg.QueryValueEx(key, key_name)
            winreg.CloseKey(key); return value
        except FileNotFoundError: return None
        except Exception as e: self.log(f"错误: 无法读取注册表项 '{key_name}' - {e}"); return None

    # --- 已添加缺失的方法 ---
    def load_settings(self):
        """加载设置文件"""
        try:
            if os.path.exists(SETTINGS_FILE):
                with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                    self.settings = json.load(f)
            else:
                self.settings = {}  # 如果文件不存在，初始化为空字典
        except (json.JSONDecodeError, IOError) as e:
            # 如果文件损坏或无法读取，也初始化为空字典
            print(f"警告：无法加载设置文件 ({e})，将使用默认设置。")
            self.settings = {}
    # --- 添加结束 ---
            
    def load_lock_password(self):
        self.lock_password_b64 = self._load_from_registry("LockPasswordB64") or ""
    
    def create_folder_structure(self):
        for folder in [PROMPT_FOLDER, AUDIO_FOLDER, BGM_FOLDER, VOICE_SCRIPT_FOLDER]:
            if not os.path.exists(folder): os.makedirs(folder)

    def create_widgets(self):
        self.root.grid_rowconfigure(1, weight=1); self.root.grid_columnconfigure(1, weight=1)

        self.status_frame = ctk.CTkFrame(self.root, height=35, corner_radius=0)
        self.status_frame.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.create_status_bar_content()

        self.nav_frame = ctk.CTkFrame(self.root, width=180, corner_radius=0)
        self.nav_frame.grid(row=0, column=0, rowspan=2, sticky="nsw")
        
        self.page_container = ctk.CTkFrame(self.root, fg_color="transparent")
        self.page_container.grid(row=1, column=1, sticky="nsew", padx=10, pady=10)

        nav_button_titles = ["定时广播", "节假日", "设置", "注册软件", "超级管理"]
        self.nav_frame.grid_rowconfigure(0, minsize=20)

        for title in nav_button_titles:
            cmd = self._prompt_for_super_admin_password if title == "超级管理" else lambda t=title: self.switch_page(t)
            btn = ctk.CTkButton(self.nav_frame, text=title, font=self.font_nav, corner_radius=0, height=50, anchor='w', border_spacing=10, command=cmd)
            btn.pack(fill="x", pady=1)
            self.nav_buttons[title] = btn
        
        self.main_frame = ctk.CTkFrame(self.page_container, fg_color="transparent")
        self.pages["定时广播"] = self.main_frame
        self.create_scheduled_broadcast_page()

        self.current_page = self.main_frame
        self.switch_page("定时广播")
        
        self.update_status_bar()
        self.log("创翔多功能定时播音旗舰版软件已启动")

    def create_status_bar_content(self):
        self.status_labels, status_texts = [], ["当前时间", "系统状态", "播放状态", "任务数量"]
        ctk.CTkLabel(self.status_frame, text="© 创翔科技", font=self.font_small).pack(side="right", padx=15)
        self.statusbar_unlock_button = ctk.CTkButton(self.status_frame, text="🔓 解锁", font=self.font_small, fg_color="#2ECC71", hover_color="#27AE60", width=80, command=self._prompt_for_password_unlock)
        for text in status_texts:
            label = ctk.CTkLabel(self.status_frame, text=f"{text}: --", font=self.font_small)
            label.pack(side="left", padx=15, pady=5)
            self.status_labels.append(label)

    def switch_page(self, page_name):
        if self.is_app_locked_down and page_name not in ["注册软件", "超级管理"]:
            self.log("软件授权已过期，请先注册。")
            if self.current_page != self.pages.get("注册软件"): self.root.after(10, lambda: self.switch_page("注册软件"))
            return

        if self.is_locked and page_name not in ["超级管理", "注册软件"]: self.log("界面已锁定，请先解锁."); return
            
        if self.current_page: self.current_page.pack_forget()
        for btn in self.nav_buttons.values(): btn.configure(fg_color="transparent", text_color=("gray10", "gray90"))

        if page_name not in self.pages:
            if page_name == "节假日": self.pages[page_name] = self.create_holiday_page()
            elif page_name == "设置": self.pages[page_name] = self.create_settings_page()
            elif page_name == "注册软件": self.pages[page_name] = self.create_registration_page()
            elif page_name == "超级管理": self.pages[page_name] = self.create_super_admin_page()
        
        if page_name == "设置": self._refresh_settings_ui()
        
        target_frame = self.pages.get(page_name, self.pages["定时广播"])
        page_name = page_name if page_name in self.pages else "定时广播"

        target_frame.pack(in_=self.page_container, fill="both", expand=True)
        self.current_page = target_frame
        
        self.nav_buttons[page_name].configure(fg_color=("gray75", "gray25"), text_color=("#1A66D2", "white"))

    def _create_input_dialog(self, title, text, show_asterisk=False):
        dialog = ctk.CTkToplevel(self.root); dialog.title(title); dialog.transient(self.root); dialog.grab_set()
        result = [None]
        def on_confirm(): result[0] = entry.get(); dialog.destroy()
        def on_cancel(): dialog.destroy()
        
        main_frame = ctk.CTkFrame(dialog, fg_color="transparent"); main_frame.pack(padx=20, pady=20, expand=True, fill="both")
        ctk.CTkLabel(main_frame, text=text, font=self.font_normal).pack(pady=(0, 10))
        entry = ctk.CTkEntry(main_frame, font=self.font_normal, width=250)
        if show_asterisk: entry.configure(show="*")
        entry.pack(pady=(0, 20), ipady=5); entry.focus_set(); entry.bind("<Return>", lambda event: on_confirm())

        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent"); btn_frame.pack()
        ctk.CTkButton(btn_frame, text="确定", font=self.font_normal, width=100, command=on_confirm).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="取消", font=self.font_normal, width=100, fg_color="gray", command=on_cancel).pack(side="left", padx=10)
        
        dialog.update_idletasks()
        self.center_window(dialog)
        self.root.wait_window(dialog)
        return result[0]

    def _prompt_for_super_admin_password(self):
        entered_password = self._create_input_dialog(title="身份验证", text="请输入超级管理员密码:", show_asterisk=True)
        if entered_password == datetime.now().strftime('%Y%m%d'):
            self.log("超级管理员密码正确，进入管理模块."); self.switch_page("超级管理")
        elif entered_password is not None:
            messagebox.showerror("验证失败", "密码错误！"); self.log("尝试进入超级管理模块失败：密码错误。")

    def create_registration_page(self):
        page_frame = ctk.CTkFrame(self.page_container, fg_color="transparent")
        ctk.CTkLabel(page_frame, text="注册软件", font=self.font_bold).pack(anchor='w', padx=20, pady=20)
        
        main_content_frame = ctk.CTkFrame(page_frame); main_content_frame.pack(padx=20, pady=10)
        main_content_frame.grid_columnconfigure(0, weight=1)

        fields_frame = ctk.CTkFrame(main_content_frame, fg_color="transparent"); fields_frame.grid(row=0, column=0, pady=10, padx=20, sticky="ew")
        fields_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(fields_frame, text="机器码:", font=self.font_normal).grid(row=0, column=0, sticky='w')
        machine_code_entry = ctk.CTkEntry(fields_frame, font=self.font_normal, text_color='red'); machine_code_entry.grid(row=0, column=1, sticky="ew", padx=10)
        machine_code_entry.insert(0, self.get_machine_code()); machine_code_entry.configure(state='disabled')

        ctk.CTkLabel(fields_frame, text="注册码:", font=self.font_normal).grid(row=1, column=0, sticky='w', pady=(10,0))
        self.reg_code_entry = ctk.CTkEntry(fields_frame, font=self.font_normal); self.reg_code_entry.grid(row=1, column=1, sticky="ew", padx=10, pady=(10,0))
        
        btn_container = ctk.CTkFrame(main_content_frame, fg_color="transparent"); btn_container.grid(row=1, column=0, pady=20)
        ctk.CTkButton(btn_container, text="注 册", font=self.font_normal, fg_color="#27AE60", hover_color="#2ECC71", width=150, height=35, command=self.attempt_registration).pack(pady=5)
        ctk.CTkButton(btn_container, text="取消注册", font=self.font_normal, fg_color="#E74C3C", hover_color="#C0392B", width=150, height=35, command=self.cancel_registration).pack(pady=5)
        
        ctk.CTkLabel(main_content_frame, text="请将您的机器码发送给软件提供商以获取注册码。\n注册码分为月度授权和永久授权两种。", font=self.font_small, text_color='gray').grid(row=2, column=0, pady=10)
        return page_frame

    def cancel_registration(self):
        if not messagebox.askyesno("确认操作", "您确定要取消当前注册吗？\n取消后，软件将恢复到试用或过期状态。"): return
        self.log("用户请求取消注册..."); self._save_to_registry('RegistrationStatus', ''); self._save_to_registry('RegistrationDate', '')
        self.check_authorization()
        messagebox.showinfo("操作完成", f"注册已成功取消。\n当前授权状态: {self.auth_info['message']}")
        self.log(f"注册已取消。新状态: {self.auth_info['message']}")
        if self.is_app_locked_down: self.perform_lockdown()
        elif self.current_page == self.pages.get("注册软件"): self.switch_page("定时广播")

    def get_machine_code(self):
        if self.machine_code: return self.machine_code
        if not PSUTIL_AVAILABLE: messagebox.showerror("依赖缺失", "psutil 库未安装，无法获取机器码。软件将退出。"); self.root.destroy(); sys.exit()
        try:
            mac = self._get_mac_address()
            if not mac: raise Exception("未找到有效的网络适配器。")
            self.machine_code = mac.upper().translate(str.maketrans("ABCDEF", "123456")); return self.machine_code
        except Exception as e: messagebox.showerror("错误", f"无法获取机器码：{e}\n软件将退出。"); self.root.destroy(); sys.exit()

    def _get_mac_address(self):
        mac_list = []
        for name, addrs in psutil.net_if_addrs().items():
            for addr in addrs:
                if addr.family == psutil.AF_LINK:
                    mac = addr.address.replace(':', '').replace('-', '').upper()
                    if len(mac) == 12 and mac != '000000000000': mac_list.append(mac)
        return mac_list[0] if mac_list else None

    def _calculate_reg_codes(self, numeric_mac_str):
        try:
            monthly = str(int(int(numeric_mac_str) * 3.14))
            permanent = f"{(int(numeric_mac_str[::-1]) / 3.14):.2f}"
            return {'monthly': monthly, 'permanent': permanent}
        except (ValueError, TypeError): return {'monthly': None, 'permanent': None}

    def attempt_registration(self):
        entered_code = self.reg_code_entry.get().strip()
        if not entered_code: messagebox.showwarning("提示", "请输入注册码."); return
        correct_codes = self._calculate_reg_codes(self.get_machine_code())
        today_str = datetime.now().strftime('%Y-%m-%d')
        reg_type, msg = (None, None)
        if entered_code == correct_codes['monthly']: reg_type, msg = 'Monthly', "恭喜您，月度授权已成功激活！"
        elif entered_code == correct_codes['permanent']: reg_type, msg = 'Permanent', "恭喜您，永久授权已成功激活！"
        
        if reg_type:
            self._save_to_registry('RegistrationStatus', reg_type); self._save_to_registry('RegistrationDate', today_str)
            messagebox.showinfo("注册成功", msg); self.check_authorization()
        else: messagebox.showerror("注册失败", "您输入的注册码无效，请重新核对。")

    def check_authorization(self):
        today = datetime.now().date()
        status, reg_date_str = self._load_from_registry('RegistrationStatus'), self._load_from_registry('RegistrationDate')
        self.is_app_locked_down = True 
        if status == 'Permanent':
            self.auth_info = {'status': 'Permanent', 'message': '永久授权'}; self.is_app_locked_down = False
        elif status == 'Monthly':
            try:
                expiry_date = datetime.strptime(reg_date_str, '%Y-%m-%d').date() + timedelta(days=30)
                if today <= expiry_date:
                    remaining = (expiry_date - today).days
                    self.auth_info = {'status': 'Monthly', 'message': f'月度授权 - 剩余 {remaining} 天'}; self.is_app_locked_down = False
                else: self.auth_info = {'status': 'Expired', 'message': '授权已过期，请注册'}
            except (TypeError, ValueError): self.auth_info = {'status': 'Expired', 'message': '授权信息损坏，请重新注册'}
        else:
            first_run_str = self._load_from_registry('FirstRunDate')
            if not first_run_str:
                self._save_to_registry('FirstRunDate', today.strftime('%Y-%m-%d'))
                self.auth_info = {'status': 'Trial', 'message': '未注册 - 剩余 3 天'}; self.is_app_locked_down = False
            else:
                try:
                    trial_expiry = datetime.strptime(first_run_str, '%Y-%m-%d').date() + timedelta(days=3)
                    if today <= trial_expiry:
                        remaining = (trial_expiry - today).days
                        self.auth_info = {'status': 'Trial', 'message': f'未注册 - 剩余 {remaining} 天'}; self.is_app_locked_down = False
                    else: self.auth_info = {'status': 'Expired', 'message': '授权已过期，请注册'}
                except (TypeError, ValueError): self.auth_info = {'status': 'Expired', 'message': '授权信息损坏，请重新注册'}
        self.update_title_bar()

    def perform_lockdown(self):
        messagebox.showerror("授权过期", "您的软件试用期或授权已到期，功能已受限。\n请在“注册软件”页面输入有效注册码以继续使用。")
        self.log("软件因授权问题被锁定。")
        for task in self.tasks: task['status'] = '禁用'
        self.update_task_list(); self.save_tasks(); self.switch_page("注册软件")

    def update_title_bar(self):
        self.root.title(f" 创翔多功能定时播音旗舰版 ({self.auth_info['message']})")
    
    def create_super_admin_page(self):
        page_frame = ctk.CTkFrame(self.page_container, fg_color="transparent")
        ctk.CTkLabel(page_frame, text="超级管理", font=self.font_bold, text_color='#C0392B').pack(anchor='w', padx=20, pady=20)
        ctk.CTkLabel(page_frame, text="警告：此处的任何操作都可能导致数据丢失或配置重置，请谨慎操作。", font=self.font_normal, text_color='red', wraplength=700).pack(anchor='w', padx=20, pady=(0, 20))
        btn_frame = ctk.CTkFrame(page_frame); btn_frame.pack(padx=20, pady=10, fill="x")
        btn_conf = {'font': self.font_normal, 'width': 200, 'height': 40}
        ctk.CTkButton(btn_frame, text="备份所有设置", command=self._backup_all_settings, **btn_conf, fg_color='#2980B9', hover_color='#3498DB').pack(pady=10)
        ctk.CTkButton(btn_frame, text="还原所有设置", command=self._restore_all_settings, **btn_conf, fg_color='#27AE60', hover_color='#2ECC71').pack(pady=10)
        ctk.CTkButton(btn_frame, text="重置软件", command=self._reset_software, **btn_conf, fg_color='#E74C3C', hover_color='#C0392B').pack(pady=10)
        ctk.CTkButton(btn_frame, text="卸载软件", command=self._prompt_for_uninstall, **btn_conf, fg_color='#34495E', hover_color='#2C3E50').pack(pady=10)
        return page_frame

    def _prompt_for_uninstall(self):
        entered_password = self._create_input_dialog(title="卸载软件 - 身份验证", text="请输入卸载密码:", show_asterisk=True)
        if entered_password == datetime.now().strftime('%Y%m%d')[::-1]:
            self.log("卸载密码正确，准备执行卸载操作。"); self._perform_uninstall()
        elif entered_password is not None:
            messagebox.showerror("验证失败", "密码错误！", parent=self.root); self.log("尝试卸载软件失败：密码错误。")

    def _perform_uninstall(self):
        if not messagebox.askyesno("！！！最终警告！！！", "您确定要卸载本软件吗？\n\n此操作将永久删除：\n- 所有注册表信息\n- 所有配置文件\n- 所有数据文件夹\n\n此操作【绝对无法恢复】！", icon='error'): return
        self.log("开始执行卸载流程..."); self.running = False
        if WIN32COM_AVAILABLE:
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, REGISTRY_KEY_PATH)
                self.log(f"成功删除注册表项: {REGISTRY_KEY_PATH}")
                try: winreg.DeleteKey(winreg.HKEY_CURRENT_USER, REGISTRY_PARENT_KEY_PATH); self.log(f"成功删除父级注册表项: {REGISTRY_PARENT_KEY_PATH}")
                except OSError: self.log("父级注册表项非空或不存在，不作删除。")
            except FileNotFoundError: self.log("未找到相关注册表项，跳过删除。")
            except Exception as e: self.log(f"删除注册表时出错: {e}")
        for path in [PROMPT_FOLDER, AUDIO_FOLDER, BGM_FOLDER, VOICE_SCRIPT_FOLDER, TASK_FILE, SETTINGS_FILE, HOLIDAY_FILE]:
            try:
                if os.path.isdir(path): shutil.rmtree(path); self.log(f"成功删除文件夹: {os.path.basename(path)}")
                elif os.path.isfile(path): os.remove(path); self.log(f"成功删除文件: {os.path.basename(path)}")
            except Exception as e: self.log(f"删除路径 {os.path.basename(path)} 时出错: {e}")
        self.log("软件数据清理完成。")
        messagebox.showinfo("卸载完成", "软件相关的数据和配置已全部清除。\n\n请手动删除本程序（.exe文件）以完成卸载。\n\n点击“确定”后软件将退出。")
        os._exit(0)

    def _backup_all_settings(self):
        self.log("开始备份所有设置...")
        try:
            backup_data = {'backup_date': datetime.now().isoformat(), 'tasks': self.tasks, 'holidays': self.holidays, 'settings': self.settings, 'lock_password_b64': self._load_from_registry("LockPasswordB64")}
            filename = filedialog.asksaveasfilename(title="备份所有设置到...", defaultextension=".json", initialfile=f"boyin_backup_{datetime.now().strftime('%Y%m%d')}.json", filetypes=[("JSON Backup", "*.json")], initialdir=application_path)
            if filename:
                with open(filename, 'w', encoding='utf-8') as f: json.dump(backup_data, f, ensure_ascii=False, indent=2)
                self.log(f"所有设置已成功备份到: {os.path.basename(filename)}"); messagebox.showinfo("备份成功", f"所有设置已成功备份到:\n{filename}")
        except Exception as e: self.log(f"备份失败: {e}"); messagebox.showerror("备份失败", f"发生错误: {e}")

    def _restore_all_settings(self):
        if not messagebox.askyesno("确认操作", "您确定要还原所有设置吗？\n当前所有配置将被立即覆盖。"): return
        self.log("开始还原所有设置...")
        filename = filedialog.askopenfilename(title="选择要还原的备份文件", filetypes=[("JSON Backup", "*.json")], initialdir=application_path)
        if not filename: return
        try:
            with open(filename, 'r', encoding='utf-8') as f: backup_data = json.load(f)
            if not all(k in backup_data for k in ['tasks', 'holidays', 'settings', 'lock_password_b64']):
                messagebox.showerror("还原失败", "备份文件格式无效或已损坏."); return
            self.tasks, self.holidays, self.settings, self.lock_password_b64 = backup_data['tasks'], backup_data['holidays'], backup_data['settings'], backup_data['lock_password_b64']
            self.save_tasks(); self.save_holidays()
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f: json.dump(self.settings, f, ensure_ascii=False, indent=2)
            self._save_to_registry("LockPasswordB64", self.lock_password_b64 or "")
            self.update_task_list(); self.update_holiday_list(); self._refresh_settings_ui()
            self.log("所有设置已从备份文件成功还原。"); messagebox.showinfo("还原成功", "所有设置已成功还原并立即应用。")
            self.root.after(100, lambda: self.switch_page("定时广播"))
        except Exception as e: self.log(f"还原失败: {e}"); messagebox.showerror("还原失败", f"发生错误: {e}")
    
    def _refresh_settings_ui(self):
        if "设置" not in self.pages or not hasattr(self, 'autostart_var'): return
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
        self.clear_password_btn.configure(state="normal" if self.lock_password_b64 and WIN32COM_AVAILABLE else "disabled")

    def _reset_software(self):
        if not messagebox.askyesno("！！！最终确认！！！", "您真的要重置整个软件吗？\n\n此操作将清空所有节目单、节假日、锁定密码和系统设置，且【无法恢复】！\n软件将在重置后提示您重启。"): return
        self.log("开始执行软件重置...")
        try:
            original_askyesno = messagebox.askyesno; messagebox.askyesno = lambda title, message: True
            self.clear_all_tasks(delete_associated_files=False); self.clear_all_holidays()
            messagebox.askyesno = original_askyesno
            self._save_to_registry("LockPasswordB64", "")
            if os.path.exists(CHIME_FOLDER): shutil.rmtree(CHIME_FOLDER); self.log("已删除整点报时缓存文件。")
            default_settings = {"autostart": False, "start_minimized": False, "lock_on_start": False, "daily_shutdown_enabled": False, "daily_shutdown_time": "23:00:00", "weekly_shutdown_enabled": False, "weekly_shutdown_days": "每周:12345", "weekly_shutdown_time": "23:30:00", "weekly_reboot_enabled": False, "weekly_reboot_days": "每周:67", "weekly_reboot_time": "22:00:00", "last_power_action_date": "", "time_chime_enabled": False, "time_chime_voice": "", "time_chime_speed": "0", "time_chime_pitch": "0"}
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f: json.dump(default_settings, f, ensure_ascii=False, indent=2)
            self.log("软件已成功重置。软件需要重启。")
            messagebox.showinfo("重置成功", "软件已恢复到初始状态。\n\n请点击“确定”后手动关闭并重新启动软件。")
        except Exception as e: self.log(f"重置失败: {e}"); messagebox.showerror("重置失败", f"发生错误: {e}")

    def create_scheduled_broadcast_page(self):
        page_frame = self.pages["定时广播"]
        page_frame.grid_columnconfigure(0, weight=1); page_frame.grid_rowconfigure(2, weight=1); page_frame.grid_rowconfigure(4, weight=1)

        top_frame = ctk.CTkFrame(page_frame, fg_color="transparent"); top_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ctk.CTkLabel(top_frame, text="定时广播", font=self.font_bold).pack(side="left")
        ctk.CTkButton(top_frame, text="添加节目", font=self.font_normal, command=self.add_task).pack(side="left", padx=10)
        
        self.top_right_btn_frame = ctk.CTkFrame(top_frame, fg_color="transparent"); self.top_right_btn_frame.pack(side="right")
        batch_buttons = [("全部启用", self.enable_all_tasks, '#27AE60'), ("全部禁用", self.disable_all_tasks, '#F39C12'), ("禁音频节目", lambda: self._set_tasks_status_by_type('audio', '禁用'), '#E67E22'), ("禁语音节目", lambda: self._set_tasks_status_by_type('voice', '禁用'), '#D35400'), ("统一音量", self.set_uniform_volume, '#8E44AD'), ("清空节目", self.clear_all_tasks, '#C0392B')]
        for text, cmd, color in batch_buttons: ctk.CTkButton(self.top_right_btn_frame, text=text, command=cmd, fg_color=color, font=self.font_small, width=90).pack(side="left", padx=3)
        self.lock_button = ctk.CTkButton(self.top_right_btn_frame, text="锁定", command=self.toggle_lock_state, fg_color='#E74C3C', font=self.font_small, width=70); self.lock_button.pack(side="left", padx=3)
        if not WIN32COM_AVAILABLE: self.lock_button.configure(state="disabled", text="锁定(Win)")
        io_buttons = [("导入节目单", self.import_tasks, '#1ABC9C'), ("导出节目单", self.export_tasks, '#1ABC9C')]
        for text, cmd, color in io_buttons: ctk.CTkButton(self.top_right_btn_frame, text=text, command=cmd, fg_color=color, font=self.font_small, width=90).pack(side="left", padx=3)

        stats_frame = ctk.CTkFrame(page_frame); stats_frame.grid(row=1, column=0, sticky="ew")
        self.stats_label = ctk.CTkLabel(stats_frame, text="节目单：0", font=self.font_normal, anchor='w'); self.stats_label.pack(side="left", fill="x", expand=True, padx=10)

        table_frame = ctk.CTkFrame(page_frame); table_frame.grid(row=2, column=0, sticky="nsew", pady=5)
        table_frame.grid_columnconfigure(0, weight=1); table_frame.grid_rowconfigure(0, weight=1)
        
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview.Heading", font=self.font_bold.actual(), background="#E1E1E1", foreground="black", relief="flat")
        style.map("Treeview.Heading", background=[('active', '#C1C1C1')])
        style.configure("Treeview", font=self.font_normal.actual(), rowheight=30, background="#FFFFFF", fieldbackground="#FFFFFF", foreground="black")
        style.map('Treeview', background=[('selected', '#3470B2')], foreground=[('selected', 'white')])
        style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])

        columns = ('节目名称', '状态', '开始时间', '模式', '音频或文字', '音量', '周几/几号', '日期范围')
        self.task_tree = ttk.Treeview(table_frame, columns=columns, show='headings', selectmode='extended'); self.task_tree.grid(row=0, column=0, sticky="nsew")
        
        self.task_tree.heading('节目名称', text='节目名称'); self.task_tree.column('节目名称', width=200, anchor='w')
        self.task_tree.heading('状态', text='状态'); self.task_tree.column('状态', width=70, anchor='center', stretch=False)
        self.task_tree.heading('开始时间', text='开始时间'); self.task_tree.column('开始时间', width=100, anchor='center', stretch=False)
        self.task_tree.heading('模式', text='模式'); self.task_tree.column('模式', width=70, anchor='center', stretch=False)
        self.task_tree.heading('音频或文字', text='音频或文字'); self.task_tree.column('音频或文字', width=300, anchor='w')
        self.task_tree.heading('音量', text='音量'); self.task_tree.column('音量', width=70, anchor='center', stretch=False)
        self.task_tree.heading('周几/几号', text='周几/几号'); self.task_tree.column('周几/几号', width=100, anchor='center')
        self.task_tree.heading('日期范围', text='日期范围'); self.task_tree.column('日期范围', width=120, anchor='center')

        scrollbar = ctk.CTkScrollbar(table_frame, command=self.task_tree.yview); scrollbar.grid(row=0, column=1, sticky="ns")
        self.task_tree.configure(yscrollcommand=scrollbar.set)
        
        self.task_tree.bind("<Button-3>", self.show_context_menu); self.task_tree.bind("<Double-1>", self.on_double_click_edit)
        self._enable_drag_selection(self.task_tree)
        
        bottom_area = ctk.CTkFrame(page_frame, fg_color="transparent"); bottom_area.grid(row=3, column=0, rowspan=2, sticky="nsew")
        bottom_area.grid_columnconfigure(0, weight=1); bottom_area.grid_rowconfigure(1, weight=1)

        playing_frame = ctk.CTkFrame(bottom_area); playing_frame.grid(row=0, column=0, sticky="ew", pady=(5,5))
        ctk.CTkLabel(playing_frame, text="正在播：", font=self.font_normal).pack(side="left", padx=10)
        self.playing_label = ctk.CTkLabel(playing_frame, text="等待播放...", font=self.font_normal, anchor='w', justify="left"); self.playing_label.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.update_playing_text("等待播放...")

        log_frame = ctk.CTkFrame(bottom_area); log_frame.grid(row=1, column=0, sticky="nsew", pady=(5,0))
        log_frame.grid_columnconfigure(0, weight=1); log_frame.grid_rowconfigure(1, weight=1)
        log_header_frame = ctk.CTkFrame(log_frame, fg_color="transparent"); log_header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        ctk.CTkLabel(log_header_frame, text="日志：", font=self.font_normal).pack(side="left")
        self.clear_log_btn = ctk.CTkButton(log_header_frame, text="清除日志", command=self.clear_log, font=self.font_small, width=60, height=20); self.clear_log_btn.pack(side="left", padx=10)

        self.log_text = ctk.CTkTextbox(log_frame, font=self.font_log, wrap="word", state='disabled'); self.log_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0,10))
    
    def create_holiday_page(self):
        page_frame = ctk.CTkFrame(self.page_container, fg_color="transparent")
        page_frame.grid_columnconfigure(0, weight=1); page_frame.grid_rowconfigure(2, weight=1)

        top_frame = ctk.CTkFrame(page_frame, fg_color="transparent"); top_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 5), padx=10)
        ctk.CTkLabel(top_frame, text="节假日", font=self.font_bold).pack(side="left")
        
        ctk.CTkLabel(page_frame, text="节假日不播放 (手动和立即播任务除外)，整点报时也受此约束", font=self.font_normal, text_color="gray").grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=(0, 10))

        table_frame = ctk.CTkFrame(page_frame); table_frame.grid(row=2, column=0, sticky="nsew", padx=(10, 5))
        table_frame.grid_columnconfigure(0, weight=1); table_frame.grid_rowconfigure(0, weight=1)

        columns = ('节假日名称', '状态', '开始日期时间', '结束日期时间')
        self.holiday_tree = ttk.Treeview(table_frame, columns=columns, show='headings', selectmode='extended'); self.holiday_tree.grid(row=0, column=0, sticky="nsew")
        
        self.holiday_tree.heading('节假日名称', text='节假日名称'); self.holiday_tree.column('节假日名称', width=250, anchor='w')
        self.holiday_tree.heading('状态', text='状态'); self.holiday_tree.column('状态', width=100, anchor='center')
        self.holiday_tree.heading('开始日期时间', text='开始日期时间'); self.holiday_tree.column('开始日期时间', width=200, anchor='center')
        self.holiday_tree.heading('结束日期时间', text='结束日期时间'); self.holiday_tree.column('结束日期时间', width=200, anchor='center')

        scrollbar = ctk.CTkScrollbar(table_frame, command=self.holiday_tree.yview); scrollbar.grid(row=0, column=1, sticky="ns")
        self.holiday_tree.configure(yscrollcommand=scrollbar.set)
        
        self.holiday_tree.bind("<Double-1>", lambda e: self.edit_holiday()); self.holiday_tree.bind("<Button-3>", self.show_holiday_context_menu)
        self._enable_drag_selection(self.holiday_tree)

        action_frame = ctk.CTkFrame(page_frame); action_frame.grid(row=2, column=1, sticky="ns", padx=(5, 10))
        buttons_config = [("添加", self.add_holiday), ("修改", self.edit_holiday), ("删除", self.delete_holiday), (None, None), ("全部启用", self.enable_all_holidays), ("全部禁用", self.disable_all_holidays), (None, None), ("导入节日", self.import_holidays), ("导出节日", self.export_holidays), ("清空节日", self.clear_all_holidays)]
        for text, cmd in buttons_config:
            if text is None: ctk.CTkFrame(action_frame, height=20, fg_color="transparent").pack(); continue
            ctk.CTkButton(action_frame, text=text, command=cmd, font=self.font_normal, width=100).pack(pady=5, padx=10)

        self.update_holiday_list(); return page_frame

    def create_settings_page(self):
        settings_frame = ctk.CTkScrollableFrame(self.page_container, fg_color="transparent")
        ctk.CTkLabel(settings_frame, text="系统设置", font=self.font_bold).pack(anchor='w', padx=20, pady=20)

        general_frame = ctk.CTkFrame(settings_frame); general_frame.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(general_frame, text="通用设置", font=self.font_bold).pack(anchor="w", padx=15, pady=(10, 5))
        
        self.autostart_var, self.start_minimized_var, self.lock_on_start_var = ctk.BooleanVar(), ctk.BooleanVar(), ctk.BooleanVar()
        ctk.CTkCheckBox(general_frame, text="登录windows后自动启动", variable=self.autostart_var, font=self.font_normal, command=self._handle_autostart_setting).pack(fill="x", padx=15, pady=5)
        ctk.CTkCheckBox(general_frame, text="启动后最小化到系统托盘", variable=self.start_minimized_var, font=self.font_normal, command=self.save_settings).pack(fill="x", padx=15, pady=5)
        
        lock_frame = ctk.CTkFrame(general_frame, fg_color="transparent"); lock_frame.pack(fill="x", padx=15, pady=5)
        self.lock_on_start_cb = ctk.CTkCheckBox(lock_frame, text="启动软件后立即锁定", variable=self.lock_on_start_var, font=self.font_normal, command=self._handle_lock_on_start_toggle); self.lock_on_start_cb.pack(side="left")
        if not WIN32COM_AVAILABLE: self.lock_on_start_cb.configure(state="disabled")
        ctk.CTkLabel(lock_frame, text="(请先在主界面设置锁定密码)", font=self.font_small, text_color='gray').pack(side="left", padx=5)
        
        self.clear_password_btn = ctk.CTkButton(general_frame, text="清除锁定密码", font=self.font_normal, command=self.clear_lock_password); self.clear_password_btn.pack(pady=10)
        
        time_chime_frame = ctk.CTkFrame(settings_frame); time_chime_frame.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(time_chime_frame, text="整点报时", font=self.font_bold).pack(anchor="w", padx=15, pady=(10, 5))
        
        self.time_chime_enabled_var, self.time_chime_voice_var, self.time_chime_speed_var, self.time_chime_pitch_var = ctk.BooleanVar(), ctk.StringVar(), ctk.StringVar(), ctk.StringVar()
        chime_control_frame = ctk.CTkFrame(time_chime_frame, fg_color="transparent"); chime_control_frame.pack(fill="x", padx=15, pady=5)
        ctk.CTkCheckBox(chime_control_frame, text="启用整点报时功能", variable=self.time_chime_enabled_var, font=self.font_normal, command=self._handle_time_chime_toggle).pack(side="left")

        self.chime_voice_combo = ctk.CTkComboBox(chime_control_frame, variable=self.time_chime_voice_var, values=self.get_available_voices(), font=self.font_small, width=350, state='readonly', command=lambda e: self._on_chime_params_changed(is_voice_change=True)); self.chime_voice_combo.pack(side="left", padx=10)
        
        ctk.CTkLabel(chime_control_frame, text="语速:", font=self.font_small).pack(side="left", padx=(10,0))
        speed_entry = ctk.CTkEntry(chime_control_frame, textvariable=self.time_chime_speed_var, font=self.font_small, width=40); speed_entry.pack(side="left")
        ctk.CTkLabel(chime_control_frame, text="音调:", font=self.font_small).pack(side="left", padx=(10,0))
        pitch_entry = ctk.CTkEntry(chime_control_frame, textvariable=self.time_chime_pitch_var, font=self.font_small, width=40); pitch_entry.pack(side="left")
        
        speed_entry.bind("<FocusOut>", self._on_chime_params_changed); pitch_entry.bind("<FocusOut>", self._on_chime_params_changed)

        power_frame = ctk.CTkFrame(settings_frame); power_frame.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(power_frame, text="电源管理", font=self.font_bold).pack(anchor="w", padx=15, pady=(10, 5))
        self.daily_shutdown_enabled_var, self.daily_shutdown_time_var = ctk.BooleanVar(), ctk.StringVar()
        self.weekly_shutdown_enabled_var, self.weekly_shutdown_time_var, self.weekly_shutdown_days_var = ctk.BooleanVar(), ctk.StringVar(), ctk.StringVar()
        self.weekly_reboot_enabled_var, self.weekly_reboot_time_var, self.weekly_reboot_days_var = ctk.BooleanVar(), ctk.StringVar(), ctk.StringVar()
        
        daily_frame = ctk.CTkFrame(power_frame, fg_color="transparent"); daily_frame.pack(fill="x", pady=4, padx=15)
        ctk.CTkCheckBox(daily_frame, text="每天关机", variable=self.daily_shutdown_enabled_var, font=self.font_normal, command=self.save_settings).pack(side="left")
        ctk.CTkEntry(daily_frame, textvariable=self.daily_shutdown_time_var, font=self.font_normal, width=120).pack(side="left", padx=10)
        ctk.CTkButton(daily_frame, text="设置", font=self.font_small, command=lambda: self.show_single_time_dialog(self.daily_shutdown_time_var), width=60).pack(side="left")

        weekly_frame = ctk.CTkFrame(power_frame, fg_color="transparent"); weekly_frame.pack(fill="x", pady=4, padx=15)
        ctk.CTkCheckBox(weekly_frame, text="每周关机", variable=self.weekly_shutdown_enabled_var, font=self.font_normal, command=self.save_settings).pack(side="left")
        ctk.CTkEntry(weekly_frame, textvariable=self.weekly_shutdown_days_var, font=self.font_normal, width=150).pack(side="left", padx=(10,5))
        ctk.CTkEntry(weekly_frame, textvariable=self.weekly_shutdown_time_var, font=self.font_normal, width=120).pack(side="left", padx=5)
        ctk.CTkButton(weekly_frame, text="设置", font=self.font_small, command=lambda: self.show_power_week_time_dialog("设置每周关机", self.weekly_shutdown_days_var, self.weekly_shutdown_time_var), width=60).pack(side="left")

        reboot_frame = ctk.CTkFrame(power_frame, fg_color="transparent"); reboot_frame.pack(fill="x", pady=4, padx=15)
        ctk.CTkCheckBox(reboot_frame, text="每周重启", variable=self.weekly_reboot_enabled_var, font=self.font_normal, command=self.save_settings).pack(side="left")
        ctk.CTkEntry(reboot_frame, textvariable=self.weekly_reboot_days_var, font=self.font_normal, width=150).pack(side="left", padx=(10,5))
        ctk.CTkEntry(reboot_frame, textvariable=self.weekly_reboot_time_var, font=self.font_normal, width=120).pack(side="left", padx=5)
        ctk.CTkButton(reboot_frame, text="设置", font=self.font_small, command=lambda: self.show_power_week_time_dialog("设置每周重启", self.weekly_reboot_days_var, self.weekly_reboot_time_var), width=60).pack(side="left")
        
        return settings_frame

    def _on_chime_params_changed(self, event=None, is_voice_change=False):
        current_voice, current_speed, current_pitch = self.time_chime_voice_var.get(), self.time_chime_speed_var.get(), self.time_chime_pitch_var.get()
        saved_voice, saved_speed, saved_pitch = self.settings.get("time_chime_voice", ""), self.settings.get("time_chime_speed", "0"), self.settings.get("time_chime_pitch", "0")
        if self.time_chime_enabled_var.get() and any([current_voice != saved_voice, current_speed != saved_speed, current_pitch != saved_pitch]):
            self.save_settings()
            if messagebox.askyesno("应用更改", "您更改了报时参数，需要重新生成全部24个报时文件。\n是否立即开始？"): self._handle_time_chime_toggle(force_regenerate=True)
            else:
                if is_voice_change: self.time_chime_voice_var.set(saved_voice)
                self.time_chime_speed_var.set(saved_speed); self.time_chime_pitch_var.set(saved_pitch)
        else: self.save_settings()

    def _handle_time_chime_toggle(self, force_regenerate=False):
        if self.time_chime_enabled_var.get() or force_regenerate:
            if not self.time_chime_voice_var.get():
                messagebox.showwarning("操作失败", "请先从下拉列表中选择一个播音员。")
                if not force_regenerate: self.time_chime_enabled_var.set(False)
                return
            self.save_settings(); self.log("准备启用/更新整点报时功能...")
            progress_dialog = ctk.CTkToplevel(self.root); progress_dialog.title("请稍候"); progress_dialog.resizable(False, False); progress_dialog.transient(self.root); progress_dialog.grab_set()
            ctk.CTkLabel(progress_dialog, text="正在生成整点报时文件 (0/24)...", font=self.font_normal).pack(padx=20, pady=10)
            progress_label = ctk.CTkLabel(progress_dialog, text="", font=self.font_small); progress_label.pack(pady=5)
            progress_dialog.update_idletasks(); self.center_window(progress_dialog)
            threading.Thread(target=self._generate_chime_files_worker, args=(self.time_chime_voice_var.get(), progress_dialog, progress_label), daemon=True).start()
        elif not force_regenerate:
            if messagebox.askyesno("确认操作", "您确定要禁用整点报时功能吗？\n这将删除所有已生成的报时音频文件。"):
                self.save_settings(); threading.Thread(target=self._delete_chime_files_worker, daemon=True).start()
            else: self.time_chime_enabled_var.set(True)
    
    def _get_time_period_string(self, hour):
        if 0 <= hour < 6: return "凌晨"
        elif 6 <= hour < 9: return "早上"
        elif 9 <= hour < 12: return "上午"
        elif 12 <= hour < 14: return "中午"
        elif 14 <= hour < 18: return "下午"
        else: return "晚上"

    def _generate_chime_files_worker(self, voice, progress_dialog, progress_label):
        if not os.path.exists(CHIME_FOLDER): os.makedirs(CHIME_FOLDER)
        success = True
        try:
            for hour in range(24):
                period = self._get_time_period_string(hour)
                display_hour = hour if hour <= 12 else hour - 12
                if hour == 0: display_hour = 12
                text = f"现在时刻,北京时间{period}{display_hour}点整"; output_path = os.path.join(CHIME_FOLDER, f"{hour:02d}.wav")
                self.root.after(0, lambda p=f"正在生成：{hour:02d}.wav ({hour + 1}/24)": progress_label.configure(text=p))
                voice_params = {'voice': voice, 'speed': self.settings.get("time_chime_speed", "0"), 'pitch': self.settings.get("time_chime_pitch", "0"), 'volume': '100'}
                if not self._synthesize_text_to_wav(text, voice_params, output_path): raise Exception(f"生成 {hour:02d}.wav 失败")
        except Exception as e:
            success = False; self.log(f"生成整点报时文件时出错: {e}"); self.root.after(0, messagebox.showerror, "错误", f"生成报时文件失败：{e}")
        finally:
            self.root.after(0, progress_dialog.destroy)
            if success:
                self.log("全部整点报时文件生成完毕。")
                if self.time_chime_enabled_var.get(): self.root.after(0, messagebox.showinfo, "成功", "整点报时功能已启用/更新！")
            else:
                self.log("整点报时功能启用失败。"); self.settings['time_chime_enabled'] = False
                self.root.after(0, self.time_chime_enabled_var.set, False); self.save_settings()

    def _delete_chime_files_worker(self):
        self.log("正在禁用整点报时功能，开始删除缓存文件...")
        try:
            if os.path.exists(CHIME_FOLDER): shutil.rmtree(CHIME_FOLDER); self.log("整点报时缓存文件已成功删除。")
            else: self.log("未找到整点报时缓存文件夹，无需删除。")
        except Exception as e: self.log(f"删除整点报时文件失败: {e}"); self.root.after(0, messagebox.showerror, "错误", f"删除报时文件失败：{e}")

    def toggle_lock_state(self):
        if self.is_locked: self._prompt_for_password_unlock()
        else:
            if not self.lock_password_b64: self._prompt_for_password_set()
            else: self._apply_lock()

    def _apply_lock(self):
        self.is_locked = True; self.lock_button.configure(text="解锁", fg_color='#2ECC71')
        self._set_ui_lock_state("disabled"); self.statusbar_unlock_button.pack(side="right", padx=10)
        self.log("界面已锁定。")

    def _apply_unlock(self):
        self.is_locked = False; self.lock_button.configure(text="锁定", fg_color='#E74C3C')
        self._set_ui_lock_state("normal"); self.statusbar_unlock_button.pack_forget()
        self.log("界面已解锁。")

    def perform_initial_lock(self): self.log("根据设置，软件启动时自动锁定。"); self._apply_lock()

    def _prompt_for_password_set(self):
        dialog = ctk.CTkToplevel(self.root); dialog.title("首次锁定，请设置密码"); dialog.resizable(False, False); dialog.transient(self.root); dialog.grab_set()
        
        main_frame = ctk.CTkFrame(dialog, fg_color="transparent"); main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        ctk.CTkLabel(main_frame, text="请设置一个锁定密码 (最多6位)", font=self.font_normal).pack(pady=10)
        ctk.CTkLabel(main_frame, text="输入密码:", font=self.font_normal).pack(pady=(5,0))
        pass_entry1 = ctk.CTkEntry(main_frame, show='*', width=200, font=self.font_normal); pass_entry1.pack()
        ctk.CTkLabel(main_frame, text="确认密码:", font=self.font_normal).pack(pady=(10,0))
        pass_entry2 = ctk.CTkEntry(main_frame, show='*', width=200, font=self.font_normal); pass_entry2.pack()
        
        def confirm():
            p1, p2 = pass_entry1.get(), pass_entry2.get()
            if not p1: messagebox.showerror("错误", "密码不能为空。", parent=dialog); return
            if len(p1) > 6: messagebox.showerror("错误", "密码不能超过6位。", parent=dialog); return
            if p1 != p2: messagebox.showerror("错误", "两次输入的密码不一致。", parent=dialog); return
            encoded_pass = base64.b64encode(p1.encode('utf-8')).decode('utf-8')
            if self._save_to_registry("LockPasswordB64", encoded_pass):
                self.lock_password_b64 = encoded_pass
                if "设置" in self.pages and hasattr(self, 'clear_password_btn'): self.clear_password_btn.configure(state="normal")
                messagebox.showinfo("成功", "密码设置成功，界面即将锁定。", parent=dialog); dialog.destroy(); self._apply_lock()
            else: messagebox.showerror("功能受限", "无法保存密码。\n此功能仅在Windows系统上支持且需要pywin32库。", parent=dialog)

        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent"); btn_frame.pack(pady=20)
        ctk.CTkButton(btn_frame, text="确定", font=self.font_normal, command=confirm).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="取消", font=self.font_normal, command=dialog.destroy, fg_color="gray").pack(side="left", padx=10)
        dialog.update_idletasks(); self.center_window(dialog)

    def _prompt_for_password_unlock(self):
        dialog = ctk.CTkToplevel(self.root); dialog.title("解锁界面"); dialog.resizable(False, False); dialog.transient(self.root); dialog.grab_set()
        main_frame = ctk.CTkFrame(dialog, fg_color="transparent"); main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        ctk.CTkLabel(main_frame, text="请输入密码以解锁", font=self.font_normal).pack(pady=10)
        pass_entry = ctk.CTkEntry(main_frame, show='*', width=250, font=self.font_normal); pass_entry.pack(pady=5); pass_entry.focus_set()

        def confirm(event=None):
            if base64.b64encode(pass_entry.get().encode('utf-8')).decode('utf-8') == self.lock_password_b64: dialog.destroy(); self._apply_unlock()
            else: messagebox.showerror("错误", "密码不正确！", parent=dialog)
        def clear_password_action():
            if not base64.b64encode(pass_entry.get().encode('utf-8')).decode('utf-8') == self.lock_password_b64: messagebox.showerror("错误", "密码不正确！无法清除。", parent=dialog); return
            if messagebox.askyesno("确认操作", "您确定要清除锁定密码吗？\n此操作不可恢复。", parent=dialog):
                self._perform_password_clear_logic(); dialog.destroy(); self.root.after(50, self._apply_unlock)
                self.root.after(100, lambda: messagebox.showinfo("成功", "锁定密码已成功清除。", parent=self.root))

        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent"); btn_frame.pack(pady=20)
        ctk.CTkButton(btn_frame, text="确定", font=self.font_normal, command=confirm).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="清除密码", font=self.font_normal, command=clear_password_action).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="取消", font=self.font_normal, command=dialog.destroy, fg_color="gray").pack(side="left", padx=5)
        dialog.bind('<Return>', confirm); dialog.update_idletasks(); self.center_window(dialog)

    def _perform_password_clear_logic(self):
        if self._save_to_registry("LockPasswordB64", ""):
            self.lock_password_b64 = ""; self.settings["lock_on_start"] = False
            if hasattr(self, 'lock_on_start_var'): self.lock_on_start_var.set(False)
            self.save_settings()
            if hasattr(self, 'clear_password_btn'): self.clear_password_btn.configure(state="disabled")
            self.log("锁定密码已清除。")

    def clear_lock_password(self):
        if messagebox.askyesno("确认操作", "您确定要清除锁定密码吗？\n此操作不可恢复。", parent=self.root):
            self._perform_password_clear_logic(); messagebox.showinfo("成功", "锁定密码已成功清除。", parent=self.root)

    def _handle_lock_on_start_toggle(self):
        if not self.lock_password_b64 and self.lock_on_start_var.get():
            messagebox.showwarning("无法启用", "您还未设置锁定密码。\n\n请返回“定时广播”页面，点击“锁定”按钮来首次设置密码。")
            self.root.after(50, lambda: self.lock_on_start_var.set(False))
        else: self.save_settings()

    def _set_ui_lock_state(self, state):
        for title, btn in self.nav_buttons.items():
            if title not in ["超级管理", "注册软件"]: btn.configure(state=state)
        for page_name, page_frame in self.pages.items():
            if page_frame and page_frame.winfo_exists() and page_name not in ["超级管理", "注册软件"]:
                self._set_widget_state_recursively(page_frame, state)
    
    def _set_widget_state_recursively(self, parent_widget, state):
        for child in parent_widget.winfo_children():
            if child == self.lock_button: continue
            try:
                if isinstance(child, (ctk.CTkButton, ctk.CTkEntry, ctk.CTkCheckBox, ctk.CTkComboBox, ctk.CTkTextbox, ctk.CTkScrollbar)): child.configure(state=state)
                elif isinstance(child, ttk.Treeview): child.state(['disabled'] if state == "disabled" else ['!disabled'])
                if child.winfo_children(): self._set_widget_state_recursively(child, state)
            except Exception: pass
    
    def clear_log(self):
        if messagebox.askyesno("确认操作", "您确定要清空所有日志记录吗？\n此操作不可恢复。"):
            self.log_text.configure(state='normal'); self.log_text.delete('1.0', "end"); self.log_text.configure(state='disabled'); self.log("日志已清空。")

    def on_double_click_edit(self, event):
        if not self.is_locked and self.task_tree.identify_row(event.y): self.edit_task()

    def show_context_menu(self, event):
        if self.is_locked: return
        iid = self.task_tree.identify_row(event.y)
        from tkinter import Menu
        context_menu = Menu(self.root, tearoff=0, font=self.font_normal.actual())
        if iid:
            if iid not in self.task_tree.selection(): self.task_tree.selection_set(iid)
            context_menu.add_command(label="立即播放", command=self.play_now); context_menu.add_separator()
            context_menu.add_command(label="修改", command=self.edit_task); context_menu.add_command(label="删除", command=self.delete_task); context_menu.add_command(label="复制", command=self.copy_task); context_menu.add_separator()
            context_menu.add_command(label="置顶", command=self.move_task_to_top); context_menu.add_command(label="上移", command=lambda: self.move_task(-1)); context_menu.add_command(label="下移", command=lambda: self.move_task(1)); context_menu.add_command(label="置末", command=self.move_task_to_bottom); context_menu.add_separator()
            context_menu.add_command(label="启用", command=self.enable_task); context_menu.add_command(label="禁用", command=self.disable_task)
        else: self.task_tree.selection_set(); context_menu.add_command(label="添加节目", command=self.add_task)
        context_menu.add_separator(); context_menu.add_command(label="停止当前播放", command=self.stop_current_playback); context_menu.post(event.x_root, event.y_root)

    def play_now(self):
        selection = self.task_tree.selection()
        if not selection: messagebox.showwarning("提示", "请先选择一个要立即播放的节目."); return
        task = self.tasks[self.task_tree.index(selection[0])]
        self.log(f"手动触发高优先级播放: {task['name']}"); self.playback_command_queue.put(('PLAY_INTERRUPT', (task, "manual_play")))

    def stop_current_playback(self): self.log("手动触发“停止当前播放”..."); self.playback_command_queue.put(('STOP', None))

    def add_task(self):
        choice_dialog = ctk.CTkToplevel(self.root); choice_dialog.title("选择节目类型"); choice_dialog.resizable(False, False); choice_dialog.transient(self.root); choice_dialog.grab_set()
        main_frame = ctk.CTkFrame(choice_dialog, fg_color="transparent"); main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        ctk.CTkLabel(main_frame, text="请选择要添加的节目类型", font=self.font_bold).pack(pady=15)
        ctk.CTkButton(main_frame, text="🎵 音频节目", font=self.font_normal, height=40, command=lambda: self.open_audio_dialog(choice_dialog)).pack(pady=8, fill="x")
        ctk.CTkButton(main_frame, text="🎙️ 语音节目", font=self.font_normal, height=40, command=lambda: self.open_voice_dialog(choice_dialog)).pack(pady=8, fill="x")
        choice_dialog.update_idletasks()
        self.center_window(choice_dialog)

    def open_audio_dialog(self, parent_dialog, task_to_edit=None, index=None):
        parent_dialog.destroy(); is_edit_mode = task_to_edit is not None
        dialog = ctk.CTkToplevel(self.root); dialog.title("修改音频节目" if is_edit_mode else "添加音频节目"); dialog.resizable(False, False); dialog.transient(self.root); dialog.grab_set()
        
        main_frame = ctk.CTkFrame(dialog, fg_color="transparent"); main_frame.pack(fill="both", expand=True, padx=15, pady=10)
        content_frame = ctk.CTkFrame(main_frame); content_frame.grid(row=0, column=0, sticky='ew', pady=2)
        ctk.CTkLabel(content_frame, text="节目名称:", font=self.font_normal).grid(row=0, column=0, sticky='e', padx=5, pady=2)
        name_entry = ctk.CTkEntry(content_frame, font=self.font_normal, width=400); name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        audio_type_var = ctk.StringVar(value="single")
        ctk.CTkLabel(content_frame, text="音频文件", font=self.font_normal).grid(row=1, column=0, sticky='e', padx=5, pady=2)
        audio_single_frame = ctk.CTkFrame(content_frame, fg_color="transparent"); audio_single_frame.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        ctk.CTkRadioButton(audio_single_frame, text="", variable=audio_type_var, value="single").pack(side="left")
        audio_single_entry = ctk.CTkEntry(audio_single_frame, font=self.font_normal, width=300); audio_single_entry.pack(side="left", padx=5)
        def select_single_audio():
            filename = filedialog.askopenfilename(title="选择音频文件", initialdir=AUDIO_FOLDER, filetypes=[("音频文件", "*.mp3 *.wav *.ogg *.flac *.m4a"), ("所有文件", "*.*")])
            if filename: audio_single_entry.delete(0, "end"); audio_single_entry.insert(0, filename)
        ctk.CTkButton(audio_single_frame, text="选取...", width=80, command=select_single_audio, font=self.font_small).pack(side="left", padx=5)
        
        ctk.CTkLabel(content_frame, text="音频文件夹", font=self.font_normal).grid(row=2, column=0, sticky='e', padx=5, pady=2)
        audio_folder_frame = ctk.CTkFrame(content_frame, fg_color="transparent"); audio_folder_frame.grid(row=2, column=1, columnspan=3, sticky='ew', padx=5, pady=2)
        ctk.CTkRadioButton(audio_folder_frame, text="", variable=audio_type_var, value="folder").pack(side="left")
        audio_folder_entry = ctk.CTkEntry(audio_folder_frame, font=self.font_normal, width=300); audio_folder_entry.pack(side="left", padx=5)
        def select_folder():
            foldername = filedialog.askdirectory(title="选择音频文件夹", initialdir=AUDIO_FOLDER)
            if foldername: audio_folder_entry.delete(0, "end"); audio_folder_entry.insert(0, foldername)
        ctk.CTkButton(audio_folder_frame, text="选取...", width=80, command=select_folder, font=self.font_small).pack(side="left", padx=5)
        
        play_order_frame = ctk.CTkFrame(content_frame, fg_color="transparent"); play_order_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        play_order_var = ctk.StringVar(value="sequential")
        ctk.CTkRadioButton(play_order_frame, text="顺序播", variable=play_order_var, value="sequential", font=self.font_normal).pack(side="left", padx=10)
        ctk.CTkRadioButton(play_order_frame, text="随机播", variable=play_order_var, value="random", font=self.font_normal).pack(side="left", padx=10)
        
        volume_frame = ctk.CTkFrame(content_frame, fg_color="transparent"); volume_frame.grid(row=4, column=1, columnspan=3, sticky='w', padx=5, pady=3)
        ctk.CTkLabel(volume_frame, text="音量 (0-100):", font=self.font_normal).pack(side="left")
        volume_entry = ctk.CTkEntry(volume_frame, font=self.font_normal, width=80); volume_entry.pack(side="left", padx=5)
        
        time_frame = ctk.CTkFrame(main_frame); time_frame.grid(row=1, column=0, sticky='ew', pady=4)
        ctk.CTkLabel(time_frame, text="开始时间:", font=self.font_normal).grid(row=0, column=0, sticky='e', padx=5, pady=2)
        start_time_entry = ctk.CTkEntry(time_frame, font=self.font_normal, width=400); start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        ctk.CTkLabel(time_frame, text="多个用 , 隔开", font=self.font_small, text_color="gray").grid(row=0, column=2, sticky='w', padx=5)
        ctk.CTkButton(time_frame, text="设置...", width=80, command=lambda: self.show_time_settings_dialog(start_time_entry), font=self.font_small).grid(row=0, column=3, padx=5)
        
        interval_var = ctk.StringVar(value="first")
        interval_frame1 = ctk.CTkFrame(time_frame, fg_color="transparent"); interval_frame1.grid(row=1, column=1, columnspan=2, sticky='w', padx=5, pady=2)
        ctk.CTkLabel(time_frame, text="间隔播报:", font=self.font_normal).grid(row=1, column=0, sticky='e', padx=5, pady=2)
        ctk.CTkRadioButton(interval_frame1, text="播 n 首", variable=interval_var, value="first", font=self.font_normal).pack(side="left")
        interval_first_entry = ctk.CTkEntry(interval_frame1, font=self.font_normal, width=100); interval_first_entry.pack(side="left", padx=5)
        ctk.CTkLabel(interval_frame1, text="(单曲时,指 n 遍)", font=self.font_small, text_color="gray").pack(side="left", padx=5)
        
        interval_frame2 = ctk.CTkFrame(time_frame, fg_color="transparent"); interval_frame2.grid(row=2, column=1, columnspan=2, sticky='w', padx=5, pady=2)
        ctk.CTkRadioButton(interval_frame2, text="播 n 秒", variable=interval_var, value="seconds", font=self.font_normal).pack(side="left")
        interval_seconds_entry = ctk.CTkEntry(interval_frame2, font=self.font_normal, width=100); interval_seconds_entry.pack(side="left", padx=5)
        ctk.CTkLabel(interval_frame2, text="(3600秒 = 1小时)", font=self.font_small, text_color="gray").pack(side="left", padx=5)
        
        ctk.CTkLabel(time_frame, text="周几/几号:", font=self.font_normal).grid(row=3, column=0, sticky='e', padx=5, pady=3)
        weekday_entry = ctk.CTkEntry(time_frame, font=self.font_normal, width=400); weekday_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=3)
        ctk.CTkButton(time_frame, text="选取...", width=80, command=lambda: self.show_weekday_settings_dialog(weekday_entry), font=self.font_small).grid(row=3, column=3, padx=5)
        
        ctk.CTkLabel(time_frame, text="日期范围:", font=self.font_normal).grid(row=4, column=0, sticky='e', padx=5, pady=3)
        date_range_entry = ctk.CTkEntry(time_frame, font=self.font_normal, width=400); date_range_entry.grid(row=4, column=1, sticky='ew', padx=5, pady=3)
        ctk.CTkButton(time_frame, text="设置...", width=80, command=lambda: self.show_daterange_settings_dialog(date_range_entry), font=self.font_small).grid(row=4, column=3, padx=5)
        
        other_frame = ctk.CTkFrame(main_frame); other_frame.grid(row=2, column=0, sticky='ew', pady=5)
        delay_var = ctk.StringVar(value="ontime")
        ctk.CTkLabel(other_frame, text="模式:", font=self.font_normal).grid(row=0, column=0, sticky='nw', padx=15, pady=15)
        delay_frame = ctk.CTkFrame(other_frame, fg_color="transparent"); delay_frame.grid(row=0, column=1, sticky='w', padx=5, pady=5)
        ctk.CTkRadioButton(delay_frame, text="准时播 - 如果有别的节目正在播，终止他们（默认）", variable=delay_var, value="ontime", font=self.font_normal).pack(anchor='w')
        ctk.CTkRadioButton(delay_frame, text="可延后 - 如果有别的节目正在播，排队等候", variable=delay_var, value="delay", font=self.font_normal).pack(anchor='w')
        ctk.CTkRadioButton(delay_frame, text="立即播 - 添加后停止其他节目,立即播放此节目", variable=delay_var, value="immediate", font=self.font_normal).pack(anchor='w')
        other_frame.grid_columnconfigure(1, weight=1)

        if is_edit_mode:
            task = task_to_edit; name_entry.insert(0, task.get('name', '')); start_time_entry.insert(0, task.get('time', '')); audio_type_var.set(task.get('audio_type', 'single'))
            if task.get('audio_type') == 'single': audio_single_entry.insert(0, task.get('content', ''))
            else: audio_folder_entry.insert(0, task.get('content', ''))
            play_order_var.set(task.get('play_order', 'sequential')); volume_entry.insert(0, task.get('volume', '80')); interval_var.set(task.get('interval_type', 'first'))
            interval_first_entry.insert(0, task.get('interval_first', '1')); interval_seconds_entry.insert(0, task.get('interval_seconds', '600')); weekday_entry.insert(0, task.get('weekday', '每周:1234567')); date_range_entry.insert(0, task.get('date_range', '2000-01-01 ~ 2099-12-31')); delay_var.set(task.get('delay', 'ontime'))
        else: volume_entry.insert(0, "80"); interval_first_entry.insert(0, "1"); interval_seconds_entry.insert(0, "600"); weekday_entry.insert(0, "每周:1234567"); date_range_entry.insert(0, "2000-01-01 ~ 2099-12-31")
        
        def save_task():
            audio_path = audio_single_entry.get().strip() if audio_type_var.get() == "single" else audio_folder_entry.get().strip()
            if not audio_path: messagebox.showwarning("警告", "请选择音频文件或文件夹", parent=dialog); return
            is_valid_time, time_msg = self._normalize_multiple_times_string(start_time_entry.get().strip())
            if not is_valid_time: messagebox.showwarning("格式错误", time_msg, parent=dialog); return
            is_valid_date, date_msg = self._normalize_date_range_string(date_range_entry.get().strip())
            if not is_valid_date: messagebox.showwarning("格式错误", date_msg, parent=dialog); return
            play_mode = delay_var.get(); play_this_task_now = (play_mode == 'immediate'); saved_delay_type = 'ontime' if play_mode == 'immediate' else play_mode
            new_task_data = {'name': name_entry.get().strip(), 'time': time_msg, 'content': audio_path, 'type': 'audio', 'audio_type': audio_type_var.get(), 'play_order': play_order_var.get(), 'volume': volume_entry.get().strip() or "80", 'interval_type': interval_var.get(), 'interval_first': interval_first_entry.get().strip(), 'interval_seconds': interval_seconds_entry.get().strip(), 'weekday': weekday_entry.get().strip(), 'date_range': date_msg, 'delay': saved_delay_type, 'status': '启用' if not is_edit_mode else task_to_edit.get('status', '启用'), 'last_run': {} if not is_edit_mode else task_to_edit.get('last_run', {})}
            if not new_task_data['name'] or not new_task_data['time']: messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog); return
            if is_edit_mode: self.tasks[index] = new_task_data; self.log(f"已修改音频节目: {new_task_data['name']}")
            else: self.tasks.append(new_task_data); self.log(f"已添加音频节目: {new_task_data['name']}")
            self.update_task_list(); self.save_tasks(); dialog.destroy()
            if play_this_task_now: self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))
        
        dialog_button_frame = ctk.CTkFrame(main_frame, fg_color="transparent"); dialog_button_frame.grid(row=3, column=0, sticky='e', pady=10)
        ctk.CTkButton(dialog_button_frame, text="保存修改" if is_edit_mode else "添加", command=save_task, font=self.font_normal, height=35, width=120).pack(side="left", padx=10)
        ctk.CTkButton(dialog_button_frame, text="取消", command=dialog.destroy, font=self.font_normal, height=35, width=120, fg_color="gray").pack(side="left", padx=10)
        content_frame.columnconfigure(1, weight=1); time_frame.columnconfigure(1, weight=1)

        dialog.update_idletasks()
        self.center_window(dialog)

    def open_voice_dialog(self, parent_dialog, task_to_edit=None, index=None):
        parent_dialog.destroy(); is_edit_mode = task_to_edit is not None
        dialog = ctk.CTkToplevel(self.root); dialog.title("修改语音节目" if is_edit_mode else "添加语音节目"); dialog.resizable(False, False); dialog.transient(self.root); dialog.grab_set()
        
        main_frame = ctk.CTkFrame(dialog, fg_color="transparent"); main_frame.pack(fill="both", expand=True, padx=15, pady=10)
        main_frame.grid_rowconfigure(0, weight=1)
        
        content_frame = ctk.CTkFrame(main_frame); content_frame.grid(row=0, column=0, sticky='nsew', pady=2)
        content_frame.grid_columnconfigure(1, weight=1); content_frame.grid_rowconfigure(1, weight=1)
        
        ctk.CTkLabel(content_frame, text="节目名称:", font=self.font_normal).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        name_entry = ctk.CTkEntry(content_frame, font=self.font_normal); name_entry.grid(row=0, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        
        ctk.CTkLabel(content_frame, text="播音文字:", font=self.font_normal).grid(row=1, column=0, sticky='nw', padx=5, pady=5)
        content_text = ctk.CTkTextbox(content_frame, font=self.font_normal, wrap="word"); content_text.grid(row=1, column=1, columnspan=3, sticky='nsew', padx=5, pady=5)

        script_btn_frame = ctk.CTkFrame(content_frame, fg_color="transparent"); script_btn_frame.grid(row=2, column=1, columnspan=3, sticky='w', padx=5, pady=(0, 2))
        ctk.CTkButton(script_btn_frame, text="导入文稿", width=80, command=lambda: self._import_voice_script(content_text), font=self.font_small).pack(side="left")
        ctk.CTkButton(script_btn_frame, text="导出文稿", width=80, command=lambda: self._export_voice_script(content_text, name_entry), font=self.font_small).pack(side="left", padx=10)

        ctk.CTkLabel(content_frame, text="播音员:", font=self.font_normal).grid(row=3, column=0, sticky='w', padx=5, pady=5)
        available_voices = self.get_available_voices(); voice_var = ctk.StringVar()
        ctk.CTkComboBox(content_frame, variable=voice_var, values=available_voices, font=self.font_normal, state='readonly').grid(row=3, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        
        speech_params_frame = ctk.CTkFrame(content_frame, fg_color="transparent"); speech_params_frame.grid(row=4, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        ctk.CTkLabel(speech_params_frame, text="语速:", font=self.font_normal).pack(side="left"); speed_entry = ctk.CTkEntry(speech_params_frame, font=self.font_normal, width=60); speed_entry.pack(side="left", padx=(5,10))
        ctk.CTkLabel(speech_params_frame, text="音调:", font=self.font_normal).pack(side="left"); pitch_entry = ctk.CTkEntry(speech_params_frame, font=self.font_normal, width=60); pitch_entry.pack(side="left", padx=(5,10))
        ctk.CTkLabel(speech_params_frame, text="音量:", font=self.font_normal).pack(side="left"); volume_entry = ctk.CTkEntry(speech_params_frame, font=self.font_normal, width=60); volume_entry.pack(side="left", padx=5)
        
        prompt_var = ctk.IntVar()
        prompt_frame = ctk.CTkFrame(content_frame, fg_color="transparent"); prompt_frame.grid(row=5, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        ctk.CTkCheckBox(prompt_frame, text="提示音:", variable=prompt_var, font=self.font_normal).pack(side="left")
        prompt_file_var, prompt_volume_var = ctk.StringVar(), ctk.StringVar()
        prompt_file_entry = ctk.CTkEntry(prompt_frame, textvariable=prompt_file_var, font=self.font_normal, width=150); prompt_file_entry.pack(side="left", padx=5)
        ctk.CTkButton(prompt_frame, text="...", width=30, command=lambda: self.select_file_for_entry(PROMPT_FOLDER, prompt_file_var), font=self.font_small).pack(side="left")
        ctk.CTkLabel(prompt_frame, text="音量:", font=self.font_normal).pack(side="left", padx=(10,5)); ctk.CTkEntry(prompt_frame, textvariable=prompt_volume_var, font=self.font_normal, width=60).pack(side="left", padx=5)
        
        bgm_var = ctk.IntVar()
        bgm_frame = ctk.CTkFrame(content_frame, fg_color="transparent"); bgm_frame.grid(row=6, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        ctk.CTkCheckBox(bgm_frame, text="背景音乐:", variable=bgm_var, font=self.font_normal).pack(side="left")
        bgm_file_var, bgm_volume_var = ctk.StringVar(), ctk.StringVar()
        bgm_file_entry = ctk.CTkEntry(bgm_frame, textvariable=bgm_file_var, font=self.font_normal, width=150); bgm_file_entry.pack(side="left", padx=5)
        ctk.CTkButton(bgm_frame, text="...", width=30, command=lambda: self.select_file_for_entry(BGM_FOLDER, bgm_file_var), font=self.font_small).pack(side="left")
        ctk.CTkLabel(bgm_frame, text="音量:", font=self.font_normal).pack(side="left", padx=(10,5)); ctk.CTkEntry(bgm_frame, textvariable=bgm_volume_var, font=self.font_normal, width=60).pack(side="left", padx=5)
        
        time_frame = ctk.CTkFrame(main_frame); time_frame.grid(row=1, column=0, sticky='ew', pady=2)
        ctk.CTkLabel(time_frame, text="开始时间:", font=self.font_normal).grid(row=0, column=0, sticky='e', padx=5, pady=2)
        start_time_entry = ctk.CTkEntry(time_frame, font=self.font_normal, width=400); start_time_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        ctk.CTkLabel(time_frame, text="多个用 , 隔开", font=self.font_small, text_color="gray").grid(row=0, column=2, sticky='w', padx=5)
        ctk.CTkButton(time_frame, text="设置...", width=80, command=lambda: self.show_time_settings_dialog(start_time_entry), font=self.font_small).grid(row=0, column=3, padx=5)
        ctk.CTkLabel(time_frame, text="播 n 遍:", font=self.font_normal).grid(row=1, column=0, sticky='e', padx=5, pady=2)
        repeat_entry = ctk.CTkEntry(time_frame, font=self.font_normal, width=100); repeat_entry.grid(row=1, column=1, sticky='w', padx=5, pady=2)
        ctk.CTkLabel(time_frame, text="周几/几号:", font=self.font_normal).grid(row=2, column=0, sticky='e', padx=5, pady=2)
        weekday_entry = ctk.CTkEntry(time_frame, font=self.font_normal, width=400); weekday_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=2)
        ctk.CTkButton(time_frame, text="选取...", width=80, command=lambda: self.show_weekday_settings_dialog(weekday_entry), font=self.font_small).grid(row=2, column=3, padx=5)
        ctk.CTkLabel(time_frame, text="日期范围:", font=self.font_normal).grid(row=3, column=0, sticky='e', padx=5, pady=2)
        date_range_entry = ctk.CTkEntry(time_frame, font=self.font_normal, width=400); date_range_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=2)
        ctk.CTkButton(time_frame, text="设置...", width=80, command=lambda: self.show_daterange_settings_dialog(date_range_entry), font=self.font_small).grid(row=3, column=3, padx=5)
        
        other_frame = ctk.CTkFrame(main_frame); other_frame.grid(row=2, column=0, sticky='ew', pady=4)
        delay_var = ctk.StringVar(value="delay")
        ctk.CTkLabel(other_frame, text="模式:", font=self.font_normal).grid(row=0, column=0, sticky='nw', padx=15, pady=5)
        delay_frame = ctk.CTkFrame(other_frame, fg_color="transparent"); delay_frame.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        ctk.CTkRadioButton(delay_frame, text="准时播 - 如果有别的节目正在播，终止他们", variable=delay_var, value="ontime", font=self.font_normal).pack(anchor='w', pady=1)
        ctk.CTkRadioButton(delay_frame, text="可延后 - 如果有别的节目正在播，排队等候（默认）", variable=delay_var, value="delay", font=self.font_normal).pack(anchor='w', pady=1)
        ctk.CTkRadioButton(delay_frame, text="立即播 - 添加后停止其他节目,立即播放此节目", variable=delay_var, value="immediate", font=self.font_normal).pack(anchor='w', pady=1)
        other_frame.grid_columnconfigure(1, weight=1)

        if is_edit_mode:
            task = task_to_edit; name_entry.insert(0, task.get('name', '')); content_text.insert('1.0', task.get('source_text', '')); voice_var.set(task.get('voice', '')); speed_entry.insert(0, task.get('speed', '0')); pitch_entry.insert(0, task.get('pitch', '0')); volume_entry.insert(0, task.get('volume', '80'))
            prompt_var.set(task.get('prompt', 0)); prompt_file_var.set(task.get('prompt_file', '')); prompt_volume_var.set(task.get('prompt_volume', '80')); bgm_var.set(task.get('bgm', 0)); bgm_file_var.set(task.get('bgm_file', '')); bgm_volume_var.set(task.get('bgm_volume', '40'))
            start_time_entry.insert(0, task.get('time', '')); repeat_entry.insert(0, task.get('repeat', '1')); weekday_entry.insert(0, task.get('weekday', '每周:1234567')); date_range_entry.insert(0, task.get('date_range', '2000-01-01 ~ 2099-12-31')); delay_var.set(task.get('delay', 'delay'))
        else:
            speed_entry.insert(0, "0"); pitch_entry.insert(0, "0"); volume_entry.insert(0, "80"); prompt_var.set(0); prompt_volume_var.set("80"); bgm_var.set(0); bgm_volume_var.set("40"); repeat_entry.insert(0, "1"); weekday_entry.insert(0, "每周:1234567"); date_range_entry.insert(0, "2000-01-01 ~ 2099-12-31")

        def save_task():
            text_content = content_text.get('1.0', "end").strip()
            if not text_content: messagebox.showwarning("警告", "请输入播音文字内容", parent=dialog); return
            is_valid_time, time_msg = self._normalize_multiple_times_string(start_time_entry.get().strip())
            if not is_valid_time: messagebox.showwarning("格式错误", time_msg, parent=dialog); return
            is_valid_date, date_msg = self._normalize_date_range_string(date_range_entry.get().strip())
            if not is_valid_date: messagebox.showwarning("格式错误", date_msg, parent=dialog); return
            regeneration_needed = True
            if is_edit_mode:
                original_task = task_to_edit
                if (text_content == original_task.get('source_text') and voice_var.get() == original_task.get('voice') and speed_entry.get().strip() == original_task.get('speed', '0') and pitch_entry.get().strip() == original_task.get('pitch', '0') and volume_entry.get().strip() == original_task.get('volume', '80')):
                    regeneration_needed = False; self.log("语音内容未变更，跳过重新生成WAV文件。")
            def build_task_data(wav_path, wav_filename_str):
                return {'name': name_entry.get().strip(), 'time': time_msg, 'type': 'voice', 'content': wav_path, 'wav_filename': wav_filename_str, 'source_text': text_content, 'voice': voice_var.get(), 'speed': speed_entry.get().strip() or "0", 'pitch': pitch_entry.get().strip() or "0", 'volume': volume_entry.get().strip() or "80", 'prompt': prompt_var.get(), 'prompt_file': prompt_file_var.get(), 'prompt_volume': prompt_volume_var.get(), 'bgm': bgm_var.get(), 'bgm_file': bgm_file_var.get(), 'bgm_volume': bgm_volume_var.get(), 'repeat': repeat_entry.get().strip() or "1", 'weekday': weekday_entry.get().strip(), 'date_range': date_msg, 'delay': (task_to_edit.get('delay', 'delay') if is_edit_mode else delay_var.get()), 'status': '启用' if not is_edit_mode else task_to_edit.get('status', '启用'), 'last_run': {} if not is_edit_mode else task_to_edit.get('last_run', {})}
            if not regeneration_needed:
                new_task_data = build_task_data(task_to_edit.get('content'), task_to_edit.get('wav_filename'))
                if not new_task_data['name'] or not new_task_data['time']: messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog); return
                self.tasks[index] = new_task_data; self.log(f"已修改语音节目(未重新生成语音): {new_task_data['name']}"); self.update_task_list(); self.save_tasks(); dialog.destroy()
                if delay_var.get() == 'immediate': self.playback_command_queue.put(('PLAY_INTERRUPT', (new_task_data, "manual_play")))
                return
            progress_dialog = ctk.CTkToplevel(dialog); progress_dialog.title("请稍候"); progress_dialog.resizable(False, False); progress_dialog.transient(dialog); progress_dialog.grab_set()
            ctk.CTkLabel(progress_dialog, text="语音文件生成中，请稍后...", font=self.font_normal).pack(expand=True, padx=20, pady=20)
            progress_dialog.update_idletasks(); self.center_window(progress_dialog)
            
            new_wav_filename = f"{int(time.time())}_{random.randint(1000, 9999)}.wav"; output_path = os.path.join(AUDIO_FOLDER, new_wav_filename); voice_params = {'voice': voice_var.get(), 'speed': speed_entry.get().strip() or "0", 'pitch': pitch_entry.get().strip() or "0", 'volume': volume_entry.get().strip() or "80"}
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
            synthesis_thread = threading.Thread(target=self._synthesis_worker, args=(text_content, voice_params, output_path, _on_synthesis_complete), daemon=True); synthesis_thread.start()
        
        dialog_button_frame = ctk.CTkFrame(main_frame, fg_color="transparent"); dialog_button_frame.grid(row=3, column=0, sticky='e', pady=10)
        ctk.CTkButton(dialog_button_frame, text="保存修改" if is_edit_mode else "添加", command=save_task, font=self.font_normal, height=35, width=120).pack(side="left", padx=10)
        ctk.CTkButton(dialog_button_frame, text="取消", command=dialog.destroy, font=self.font_normal, height=35, width=120, fg_color="gray").pack(side="left", padx=10)
        time_frame.columnconfigure(1, weight=1)

        dialog.update_idletasks()
        self.center_window(dialog)

    def _import_voice_script(self, text_widget):
        filename = filedialog.askopenfilename(title="选择要导入的文稿", initialdir=VOICE_SCRIPT_FOLDER, filetypes=[("文本文档", "*.txt"), ("所有文件", "*.*")])
        if not filename: return
        try:
            with open(filename, 'r', encoding='utf-8') as f: content = f.read()
            text_widget.delete('1.0', "end"); text_widget.insert('1.0', content); self.log(f"已从 {os.path.basename(filename)} 成功导入文稿。")
        except Exception as e: messagebox.showerror("导入失败", f"无法读取文件：\n{e}"); self.log(f"导入文稿失败: {e}")

    def _export_voice_script(self, text_widget, name_widget):
        content = text_widget.get('1.0', "end").strip()
        if not content: messagebox.showwarning("无法导出", "播音文字内容为空，无需导出。"); return
        program_name = name_widget.get().strip()
        safe_name = "".join(c for c in program_name if c not in '\\/:*?"<>|').strip() if program_name else ""
        default_filename = f"{safe_name}.txt" if safe_name else "未命名文稿.txt"
        filename = filedialog.asksaveasfilename(title="导出文稿到...", initialdir=VOICE_SCRIPT_FOLDER, initialfile=default_filename, defaultextension=".txt", filetypes=[("文本文档", "*.txt")])
        if not filename: return
        try:
            with open(filename, 'w', encoding='utf-8') as f: f.write(content)
            self.log(f"文稿已成功导出到 {os.path.basename(filename)}。"); messagebox.showinfo("导出成功", f"文稿已成功导出到：\n{filename}")
        except Exception as e: messagebox.showerror("导出失败", f"无法保存文件：\n{e}"); self.log(f"导出文稿失败: {e}")

    def _synthesis_worker(self, text, voice_params, output_path, callback):
        try:
            if self._synthesize_text_to_wav(text, voice_params, output_path): self.root.after(0, callback, {'success': True})
            else: raise Exception("合成过程返回失败")
        except Exception as e: self.root.after(0, callback, {'success': False, 'error': str(e)})

    def _synthesize_text_to_wav(self, text, voice_params, output_path):
        if not WIN32COM_AVAILABLE: raise ImportError("pywin32 模块未安装，无法进行语音合成。")
        pythoncom.CoInitialize()
        try:
            speaker, stream = win32com.client.Dispatch("SAPI.SpVoice"), win32com.client.Dispatch("SAPI.SpFileStream")
            stream.Open(output_path, 3, False); speaker.AudioOutputStream = stream
            all_voices = {v.GetDescription(): v for v in speaker.GetVoices()}
            if (selected_voice_desc := voice_params.get('voice')) in all_voices: speaker.Voice = all_voices[selected_voice_desc]
            speaker.Volume = int(voice_params.get('volume', 80))
            escaped_text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("'", "&apos;").replace('"', "&quot;")
            xml_text = f"<rate absspeed='{voice_params.get('speed', '0')}'><pitch middle='{voice_params.get('pitch', '0')}'>{escaped_text}</pitch></rate>"
            speaker.Speak(xml_text, 1); speaker.WaitUntilDone(-1); stream.Close(); return True
        except Exception as e: self.log(f"语音合成到文件时出错: {e}"); return False
        finally: pythoncom.CoUninitialize()

    def get_available_voices(self):
        if not WIN32COM_AVAILABLE: return []
        try:
            pythoncom.CoInitialize(); speaker = win32com.client.Dispatch("SAPI.SpVoice")
            voices = [v.GetDescription() for v in speaker.GetVoices()]
            pythoncom.CoUninitialize(); return voices
        except Exception as e: self.log(f"警告: 使用 win32com 获取语音列表失败 - {e}"); return []
    
    def select_file_for_entry(self, initial_dir, string_var):
        filename = filedialog.askopenfilename(title="选择文件", initialdir=initial_dir, filetypes=[("音频文件", "*.mp3 *.wav *.ogg *.flac *.m4a"), ("所有文件", "*.*")])
        if filename: string_var.set(os.path.basename(filename))

    def delete_task(self):
        selections = self.task_tree.selection()
        if not selections: messagebox.showwarning("警告", "请先选择要删除的节目"); return
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(selections)} 个节目吗？\n(关联的语音文件也将被删除)"):
            indices = sorted([self.task_tree.index(s) for s in selections], reverse=True)
            for index in indices:
                task = self.tasks.pop(index)
                if task.get('type') == 'voice' and task.get('wav_filename'):
                    try:
                        wav_path = os.path.join(AUDIO_FOLDER, task['wav_filename'])
                        if os.path.exists(wav_path): os.remove(wav_path); self.log(f"已删除语音文件: {task['wav_filename']}")
                    except Exception as e: self.log(f"删除语音文件失败: {e}")
                self.log(f"已删除节目: {task['name']}")
            self.update_task_list(); self.save_tasks()

    def edit_task(self):
        selection = self.task_tree.selection()
        if not selection or len(selection) > 1: messagebox.showwarning("警告", "请选择一个节目进行修改"); return
        index = self.task_tree.index(selection[0]); task = self.tasks[index]
        dummy_parent = ctk.CTkToplevel(self.root); dummy_parent.withdraw()
        if task.get('type') == 'audio': self.open_audio_dialog(dummy_parent, task_to_edit=task, index=index)
        else: self.open_voice_dialog(dummy_parent, task_to_edit=task, index=index)
        def check_dialog(): self.root.after(100, lambda: [dummy_parent.destroy(), check_dialog()] if not dummy_parent.winfo_children() else None)
        check_dialog()

    def copy_task(self):
        selections = self.task_tree.selection()
        if not selections: messagebox.showwarning("警告", "请先选择要复制的节目"); return
        for sel in selections:
            original = self.tasks[self.task_tree.index(sel)]; copy = json.loads(json.dumps(original))
            copy['name'] += " (副本)"; copy['last_run'] = {}
            if copy.get('type') == 'voice' and 'source_text' in copy:
                wav_filename = f"{int(time.time())}_{random.randint(1000, 9999)}.wav"; output_path = os.path.join(AUDIO_FOLDER, wav_filename)
                voice_params = {k: copy.get(k) for k in ['voice', 'speed', 'pitch', 'volume']}
                try:
                    if not self._synthesize_text_to_wav(copy['source_text'], voice_params, output_path): raise Exception("语音合成失败")
                    copy['content'], copy['wav_filename'] = output_path, wav_filename; self.log(f"已为副本生成新语音文件: {wav_filename}")
                except Exception as e: self.log(f"为副本生成语音文件失败: {e}"); continue
            self.tasks.append(copy); self.log(f"已复制节目: {original['name']}")
        self.update_task_list(); self.save_tasks()

    def move_task(self, direction):
        selections = self.task_tree.selection()
        if not selections or len(selections) > 1: return
        index = self.task_tree.index(selections[0]); new_index = index + direction
        if 0 <= new_index < len(self.tasks):
            self.tasks.insert(new_index, self.tasks.pop(index)); self.update_task_list(); self.save_tasks()
            items = self.task_tree.get_children()
            if items: self.task_tree.selection_set(items[new_index]); self.task_tree.focus(items[new_index])

    def move_task_to_top(self):
        selections = self.task_tree.selection()
        if not selections or len(selections) > 1: return
        index = self.task_tree.index(selections[0])
        if index > 0:
            self.tasks.insert(0, self.tasks.pop(index)); self.update_task_list(); self.save_tasks()
            items = self.task_tree.get_children()
            if items: self.task_tree.selection_set(items[0]); self.task_tree.focus(items[0])

    def move_task_to_bottom(self):
        selections = self.task_tree.selection()
        if not selections or len(selections) > 1: return
        index = self.task_tree.index(selections[0])
        if index < len(self.tasks) - 1:
            self.tasks.append(self.tasks.pop(index)); self.update_task_list(); self.save_tasks()
            items = self.task_tree.get_children()
            if items: self.task_tree.selection_set(items[-1]); self.task_tree.focus(items[-1])

    def import_tasks(self):
        filename = filedialog.askopenfilename(title="选择导入文件", filetypes=[("JSON文件", "*.json")], initialdir=application_path)
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f: imported = json.load(f)
                if not isinstance(imported, list) or (imported and (not isinstance(imported[0], dict) or 'time' not in imported[0] or 'type' not in imported[0])):
                    messagebox.showerror("导入失败", "文件格式不正确。"); return
                self.tasks.extend(imported); self.update_task_list(); self.save_tasks(); self.log(f"已从 {os.path.basename(filename)} 导入 {len(imported)} 个节目")
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
        type_name, status_name = ("音频" if task_type == 'audio' else "语音"), ("启用" if status == '启用' else "禁用")
        count = sum(1 for task in self.tasks if task.get('type') == task_type and task.get('status') != status)
        if count > 0:
            for task in self.tasks:
                if task.get('type') == task_type: task['status'] = status
            self.update_task_list(); self.save_tasks(); self.log(f"已将 {count} 个{type_name}节目设置为“{status_name}”状态。")
        else: self.log(f"没有需要状态更新的{type_name}节目。")

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
        volume_str = self._create_input_dialog(title="统一音量", text="请输入统一音量值 (0-100):")
        if volume_str:
            try:
                volume = int(volume_str)
                if 0 <= volume <= 100:
                    for task in self.tasks: task['volume'] = str(volume)
                    self.update_task_list(); self.save_tasks(); self.log(f"已将全部节目音量统一设置为 {volume}。")
                else: messagebox.showerror("输入错误", "请输入一个介于 0 和 100 之间的整数。")
            except ValueError: messagebox.showerror("输入错误", "请输入一个有效的整数。")

    def clear_all_tasks(self, delete_associated_files=True):
        if not self.tasks: return
        msg = "您确定要清空所有节目吗？\n此操作将同时删除关联的语音文件，且不可恢复！" if delete_associated_files else "您确定要清空所有节目列表吗？\n（此操作不会删除音频文件）"
        if messagebox.askyesno("严重警告", msg):
            files_to_delete = [os.path.join(AUDIO_FOLDER, task['wav_filename']) for task in self.tasks if delete_associated_files and task.get('type') == 'voice' and task.get('wav_filename') and os.path.exists(os.path.join(AUDIO_FOLDER, task['wav_filename']))]
            self.tasks.clear(); self.update_task_list(); self.save_tasks(); self.log("已清空所有节目列表。")
            for f in files_to_delete:
                try: os.remove(f); self.log(f"已删除语音文件: {os.path.basename(f)}")
                except Exception as e: self.log(f"删除语音文件失败: {e}")

    def show_time_settings_dialog(self, time_entry_widget):
        dialog = ctk.CTkToplevel(self.root); dialog.title("开始时间设置"); dialog.resizable(False, False); dialog.transient(self.root); dialog.grab_set()
        main_frame = ctk.CTkFrame(dialog, fg_color="transparent"); main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        ctk.CTkLabel(main_frame, text="24小时制 HH:MM:SS", font=self.font_bold).pack(anchor='w', pady=5)
        
        list_frame = ctk.CTkFrame(main_frame); list_frame.pack(fill="both", expand=True, pady=5)
        list_frame.grid_columnconfigure(0, weight=1); list_frame.grid_rowconfigure(0, weight=1)

        scrollable_list = ctk.CTkScrollableFrame(list_frame, label_text="时间列表", label_font=self.font_normal); scrollable_list.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        time_labels = [ctk.CTkButton(scrollable_list, text=t, fg_color="transparent", text_color=("gray10", "gray90"), anchor="w", font=self.font_normal) for t in time_entry_widget.get().split(',') if t.strip()]
        for label in time_labels: label.pack(fill="x", pady=1)

        btn_frame = ctk.CTkFrame(list_frame, fg_color="transparent"); btn_frame.grid(row=0, column=1, padx=10, sticky="ns")
        new_entry = ctk.CTkEntry(btn_frame, font=self.font_normal, width=120); new_entry.insert(0, datetime.now().strftime("%H:%M:%S")); new_entry.pack(pady=3)
        
        def add_time():
            val = self._normalize_time_string(new_entry.get().strip())
            if val and val not in [lbl.cget("text") for lbl in time_labels]:
                label = ctk.CTkButton(scrollable_list, text=val, fg_color="transparent", text_color=("gray10", "gray90"), anchor="w", font=self.font_normal)
                label.pack(fill="x", pady=1); time_labels.append(label)
                new_entry.delete(0, "end"); new_entry.insert(0, datetime.now().strftime("%H:%M:%S"))
            elif not val: messagebox.showerror("格式错误", "请输入有效的时间格式 HH:MM:SS", parent=dialog)
        
        def clear_times():
            for widget in scrollable_list.winfo_children(): widget.destroy()
            time_labels.clear()

        ctk.CTkButton(btn_frame, text="添加 ↑", command=add_time, font=self.font_normal).pack(pady=3, fill="x")
        ctk.CTkButton(btn_frame, text="清空", command=clear_times, font=self.font_normal).pack(pady=3, fill="x")
        
        bottom_frame = ctk.CTkFrame(main_frame, fg_color="transparent"); bottom_frame.pack(pady=10)
        def confirm():
            time_entry_widget.delete(0, "end"); time_entry_widget.insert(0, ", ".join([lbl.cget("text") for lbl in time_labels]))
            self.save_settings(); dialog.destroy()
        ctk.CTkButton(bottom_frame, text="确定", command=confirm, font=self.font_bold, width=100, height=35).pack(side="left", padx=5)
        ctk.CTkButton(bottom_frame, text="取消", command=dialog.destroy, font=self.font_normal, width=100, height=35, fg_color="gray").pack(side="left", padx=5)
        
        dialog.update_idletasks()
        self.center_window(dialog)

    def show_weekday_settings_dialog(self, weekday_var_entry):
        dialog = ctk.CTkToplevel(self.root); dialog.title("周几或几号"); dialog.resizable(False, False); dialog.transient(self.root); dialog.grab_set()
        main_frame = ctk.CTkFrame(dialog, fg_color="transparent"); main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        week_type_var = ctk.StringVar(value="week")
        
        week_frame = ctk.CTkFrame(main_frame); week_frame.pack(fill="x", pady=5)
        ctk.CTkRadioButton(week_frame, text="每周", variable=week_type_var, value="week", font=self.font_normal).grid(row=0, column=0, sticky='w', padx=10, pady=10)
        weekdays = [("周一", 1), ("周二", 2), ("周三", 3), ("周四", 4), ("周五", 5), ("周六", 6), ("周日", 7)]; week_vars = {num: ctk.IntVar(value=1) for day, num in weekdays}
        for i, (day, num) in enumerate(weekdays): ctk.CTkCheckBox(week_frame, text=day, variable=week_vars[num], font=self.font_normal).grid(row=(i // 4) + 1, column=i % 4, sticky='w', padx=10, pady=3)
        
        day_frame = ctk.CTkFrame(main_frame); day_frame.pack(fill="both", expand=True, pady=5)
        ctk.CTkRadioButton(day_frame, text="每月", variable=week_type_var, value="day", font=self.font_normal).grid(row=0, column=0, sticky='w', padx=10, pady=10)
        day_vars = {i: ctk.IntVar(value=0) for i in range(1, 32)}
        for i in range(1, 32): ctk.CTkCheckBox(day_frame, text=f"{i:02d}", variable=day_vars[i], font=self.font_normal).grid(row=((i - 1) // 7) + 1, column=(i - 1) % 7, sticky='w', padx=8, pady=2)
        
        current_val = weekday_var_entry.get()
        if current_val.startswith("每周:"):
            week_type_var.set("week"); selected_days = current_val.replace("每周:", "")
            for day_num in week_vars: week_vars[day_num].set(1 if str(day_num) in selected_days else 0)
        elif current_val.startswith("每月:"):
            week_type_var.set("day"); selected_days = current_val.replace("每月:", "").split(',')
            for day_num in day_vars: day_vars[day_num].set(1 if f"{day_num:02d}" in selected_days else 0)
        
        bottom_frame = ctk.CTkFrame(main_frame, fg_color="transparent"); bottom_frame.pack(pady=10)
        def confirm():
            result = "每周:" + "".join(sorted([str(n) for n, v in week_vars.items() if v.get()])) if week_type_var.get() == "week" else "每月:" + ",".join(sorted([f"{n:02d}" for n, v in day_vars.items() if v.get()]))
            weekday_var_entry.delete(0, "end"); weekday_var_entry.insert(0, result); self.save_settings(); dialog.destroy()
        ctk.CTkButton(bottom_frame, text="确定", command=confirm, font=self.font_bold, height=35, width=120).pack(side="left", padx=5)
        ctk.CTkButton(bottom_frame, text="取消", command=dialog.destroy, font=self.font_normal, height=35, width=120, fg_color="gray").pack(side="left", padx=5)

        dialog.update_idletasks()
        self.center_window(dialog)

    def show_daterange_settings_dialog(self, date_range_entry):
        dialog = ctk.CTkToplevel(self.root); dialog.title("日期范围"); dialog.resizable(False, False); dialog.transient(self.root); dialog.grab_set()
        main_frame = ctk.CTkFrame(dialog, fg_color="transparent"); main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        from_frame = ctk.CTkFrame(main_frame, fg_color="transparent"); from_frame.pack(pady=10, anchor='w')
        ctk.CTkLabel(from_frame, text="从", font=self.font_bold).pack(side="left", padx=5)
        from_date_entry = ctk.CTkEntry(from_frame, font=self.font_normal, width=180); from_date_entry.pack(side="left", padx=5)
        
        to_frame = ctk.CTkFrame(main_frame, fg_color="transparent"); to_frame.pack(pady=10, anchor='w')
        ctk.CTkLabel(to_frame, text="到", font=self.font_bold).pack(side="left", padx=5)
        to_date_entry = ctk.CTkEntry(to_frame, font=self.font_normal, width=180); to_date_entry.pack(side="left", padx=5)
        
        try: start, end = date_range_entry.get().split('~'); from_date_entry.insert(0, start.strip()); to_date_entry.insert(0, end.strip())
        except (ValueError, IndexError): from_date_entry.insert(0, "2000-01-01"); to_date_entry.insert(0, "2099-12-31")
        ctk.CTkLabel(main_frame, text="格式: YYYY-MM-DD", font=self.font_normal, text_color='gray').pack(pady=10)
        
        bottom_frame = ctk.CTkFrame(main_frame, fg_color="transparent"); bottom_frame.pack(pady=10)
        def confirm():
            start, end = from_date_entry.get().strip(), to_date_entry.get().strip()
            norm_start, norm_end = self._normalize_date_string(start), self._normalize_date_string(end)
            if norm_start and norm_end: date_range_entry.delete(0, "end"); date_range_entry.insert(0, f"{norm_start} ~ {norm_end}"); dialog.destroy()
            else: messagebox.showerror("格式错误", "日期格式不正确, 应为 YYYY-MM-DD", parent=dialog)
        ctk.CTkButton(bottom_frame, text="确定", command=confirm, font=self.font_bold, height=35, width=120).pack(side="left", padx=5)
        ctk.CTkButton(bottom_frame, text="取消", command=dialog.destroy, font=self.font_normal, height=35, width=120, fg_color="gray").pack(side="left", padx=5)

        dialog.update_idletasks()
        self.center_window(dialog)

    def show_single_time_dialog(self, time_var):
        dialog = ctk.CTkToplevel(self.root); dialog.title("设置时间"); dialog.resizable(False, False); dialog.transient(self.root); dialog.grab_set()
        main_frame = ctk.CTkFrame(dialog, fg_color="transparent"); main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        ctk.CTkLabel(main_frame, text="24小时制 HH:MM:SS", font=self.font_bold).pack(pady=5)
        time_entry = ctk.CTkEntry(main_frame, font=self.font_normal, width=150, justify='center'); time_entry.insert(0, time_var.get()); time_entry.pack(pady=10)
        def confirm():
            normalized_time = self._normalize_time_string(time_entry.get().strip())
            if normalized_time: time_var.set(normalized_time); self.save_settings(); dialog.destroy()
            else: messagebox.showerror("格式错误", "请输入有效的时间格式 HH:MM:SS", parent=dialog)
        bottom_frame = ctk.CTkFrame(main_frame, fg_color="transparent"); bottom_frame.pack(pady=10)
        ctk.CTkButton(bottom_frame, text="确定", command=confirm, font=self.font_normal).pack(side="left", padx=10)
        ctk.CTkButton(bottom_frame, text="取消", command=dialog.destroy, font=self.font_normal, fg_color="gray").pack(side="left", padx=10)

        dialog.update_idletasks()
        self.center_window(dialog)

    def show_power_week_time_dialog(self, title, days_var, time_var):
        dialog = ctk.CTkToplevel(self.root); dialog.title(title); dialog.resizable(False, False); dialog.transient(self.root); dialog.grab_set()
        
        main_frame = ctk.CTkFrame(dialog, fg_color="transparent"); main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        week_frame = ctk.CTkFrame(main_frame,); week_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(week_frame, text="选择周几", font=self.font_bold).grid(row=0, column=0, columnspan=7, pady=5, padx=10)
        weekdays = [("周一", 1), ("周二", 2), ("周三", 3), ("周四", 4), ("周五", 5), ("周六", 6), ("周日", 7)]; week_vars = {num: ctk.IntVar() for day, num in weekdays}
        for day_num_str in days_var.get().replace("每周:", ""): week_vars[int(day_num_str)].set(1)
        for i, (day, num) in enumerate(weekdays): ctk.CTkCheckBox(week_frame, text=day, variable=week_vars[num], font=self.font_normal).grid(row=1, column=i, sticky='w', padx=10, pady=10)
        
        time_frame = ctk.CTkFrame(main_frame); time_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(time_frame, text="时间 (HH:MM:SS):", font=self.font_normal).pack(side="left", padx=10)
        time_entry = ctk.CTkEntry(time_frame, font=self.font_normal, width=150); time_entry.insert(0, time_var.get()); time_entry.pack(side="left", padx=10)
        
        def confirm():
            selected_days = sorted([str(n) for n, v in week_vars.items() if v.get()])
            if not selected_days: messagebox.showwarning("提示", "请至少选择一天", parent=dialog); return
            normalized_time = self._normalize_time_string(time_entry.get().strip())
            if not normalized_time: messagebox.showerror("格式错误", "请输入有效的时间格式 HH:MM:SS", parent=dialog); return
            days_var.set("每周:" + "".join(selected_days)); time_var.set(normalized_time); self.save_settings(); dialog.destroy()
        
        bottom_frame = ctk.CTkFrame(main_frame, fg_color="transparent"); bottom_frame.pack(pady=15)
        ctk.CTkButton(bottom_frame, text="确定", command=confirm, font=self.font_normal).pack(side="left", padx=10)
        ctk.CTkButton(bottom_frame, text="取消", command=dialog.destroy, font=self.font_normal, fg_color="gray").pack(side="left", padx=10)

        dialog.update_idletasks()
        self.center_window(dialog)

    def show_quit_dialog(self):
        dialog = ctk.CTkToplevel(self.root); dialog.title("确认"); dialog.resizable(False, False); dialog.transient(self.root); dialog.grab_set()
        
        main_frame = ctk.CTkFrame(dialog, fg_color="transparent"); main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ctk.CTkLabel(main_frame, text="您想要如何操作？", font=self.font_bold).pack(pady=20)
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent"); btn_frame.pack(pady=10)
        ctk.CTkButton(btn_frame, text="退出程序", command=lambda: [dialog.destroy(), self.quit_app()], font=self.font_normal, fg_color="#E74C3C").pack(side="left", padx=10)
        if TRAY_AVAILABLE: ctk.CTkButton(btn_frame, text="最小化到托盘", command=lambda: [dialog.destroy(), self.hide_to_tray()], font=self.font_normal).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="取消", command=dialog.destroy, font=self.font_normal, fg_color="gray").pack(side="left", padx=10)

        dialog.update_idletasks()
        self.center_window(dialog)

    def center_window(self, win):
        win.update_idletasks()
        width = win.winfo_reqwidth()
        height = win.winfo_reqheight()
        parent_x = self.root.winfo_x(); parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width(); parent_height = self.root.winfo_height()
        x = parent_x + (parent_width // 2) - (width // 2)
        y = parent_y + (parent_height // 2) - (height // 2)
        win.geometry(f'{width}x{height}+{x}+{y}')

def main():
    root = ctk.CTk()
    app = TimedBroadcastApp(root)
    root.mainloop()

if __name__ == "__main__":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError):
            print("警告: 无法设置DPI感知。")

    if not WIN32COM_AVAILABLE: messagebox.showerror("核心依赖缺失", "pywin32 库未安装或损坏，软件无法运行注册和锁定等核心功能，即将退出。"); sys.exit()
    if not PSUTIL_AVAILABLE: messagebox.showerror("核心依赖缺失", "psutil 库未安装，软件无法获取机器码以进行授权验证，即将退出。"); sys.exit()
    
    main()
