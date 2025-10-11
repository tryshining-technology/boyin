import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import json
import threading
import time
from datetime import datetime
import os
import random
import sys

# ... [此处省略所有 import 语句，它们没有变化] ...
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
        self.running = True
        self.task_file = TASK_FILE
        self.holiday_file = HOLIDAY_FILE
        self.tray_icon = None

        self.is_playing = threading.Event()
        self.playback_queue = []
        self.queue_lock = threading.Lock()
        self.holiday_pause_logged = False

        self.create_folder_structure()
        self.create_main_layout()
        self.load_tasks()
        self.load_holidays()
        self.switch_page("定时广播")
        self.start_background_thread()
        self.root.protocol("WM_DELETE_WINDOW", self.show_quit_dialog)

    def create_folder_structure(self):
        for folder in [PROMPT_FOLDER, AUDIO_FOLDER, BGM_FOLDER]:
            if not os.path.exists(folder):
                os.makedirs(folder)

    def create_main_layout(self):
        self.nav_frame = tk.Frame(self.root, bg='#A8D8E8', width=160)
        self.nav_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.nav_frame.pack_propagate(False)

        self.nav_buttons = {}
        nav_titles = ["定时广播", "节假日", "立即插播", "语音广告 制作", "设置"]
        
        for title in nav_titles:
            btn_frame = tk.Frame(self.nav_frame, bg='#A8D8E8')
            btn_frame.pack(fill=tk.X, pady=1)
            btn = tk.Button(btn_frame, text=title, bg='#A8D8E8', fg='black', 
                          font=('Microsoft YaHei', 13, 'bold'),
                          bd=0, padx=10, pady=8, anchor='w', 
                          command=lambda t=title: self.switch_page(t))
            btn.pack(fill=tk.X)
            self.nav_buttons[title] = btn_frame

        content_area = tk.Frame(self.root, bg='white')
        content_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.page_frame = tk.Frame(content_area, bg='white')
        self.page_frame.pack(fill=tk.BOTH, expand=True)

        playing_frame = tk.LabelFrame(content_area, text="正在播：", font=('Microsoft YaHei', 10),
                                     bg='white', fg='#2C5F7C', padx=10, pady=5)
        playing_frame.pack(fill=tk.X, padx=10, pady=5)
        self.playing_text = scrolledtext.ScrolledText(playing_frame, height=3, font=('Microsoft YaHei', 9),
                                                     bg='#FFFEF0', wrap=tk.WORD, state='disabled')
        self.playing_text.pack(fill=tk.BOTH, expand=True)

        log_frame = tk.Frame(content_area, bg='white', padx=10, pady=5)
        log_frame.pack(fill=tk.BOTH, expand=True)
        log_header_frame = tk.Frame(log_frame, bg='white')
        log_header_frame.pack(fill=tk.X)
        log_label = tk.Label(log_header_frame, text="日志：", font=('Microsoft YaHei', 10, 'bold'),
                             bg='white', fg='#2C5F7C')
        log_label.pack(side=tk.LEFT)
        clear_log_btn = tk.Button(log_header_frame, text="清除日志", command=self.clear_log,
                                  font=('Microsoft YaHei', 8), bd=0, bg='#EAEAEA',
                                  fg='#333', cursor='hand2', padx=5, pady=0)
        clear_log_btn.pack(side=tk.LEFT, padx=10)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, font=('Microsoft YaHei', 9),
                                                 bg='#F9F9F9', wrap=tk.WORD, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True)

        status_frame = tk.Frame(content_area, bg='#E8F4F8', height=30)
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
        self.update_playing_text("等待播放...")
        self.log("定时播音软件已启动")

    # --- 新增：核心验证和归一化函数 ---
    def _validate_and_normalize_time_list(self, time_string, parent_dialog):
        """验证并归一化一个逗号分隔的时间字符串，返回标准格式或None"""
        raw_times = [t.strip() for t in time_string.split(',') if t.strip()]
        normalized_times = []
        invalid_times = []

        for t_str in raw_times:
            parts = t_str.split(':')
            try:
                # 补全缺失的分钟和秒
                while len(parts) < 3:
                    parts.append('0')
                
                if len(parts) > 3:
                    raise ValueError("时间格式包含过多部分")

                h, m, s = [int(p) for p in parts]

                if not (0 <= h <= 23 and 0 <= m <= 59 and 0 <= s <= 59):
                    raise ValueError("时间数值超出范围")
                
                normalized_times.append(f"{h:02d}:{m:02d}:{s:02d}")
            except (ValueError, TypeError):
                invalid_times.append(t_str)

        if invalid_times:
            messagebox.showerror("时间格式错误", 
                                 f"以下时间格式无效，请修正后保存：\n\n"
                                 f"{', '.join(invalid_times)}\n\n"
                                 f"有效格式为 HH:MM:SS，例如 '09:30:00' 或 '9:30'。",
                                 parent=parent_dialog)
            return None
        
        return ", ".join(normalized_times)

    def _clear_page_frame(self):
        for widget in self.page_frame.winfo_children():
            widget.destroy()

    def switch_page(self, page_name):
        for title, frame in self.nav_buttons.items():
            is_active = (title == page_name)
            frame.config(bg='#5DADE2' if is_active else '#A8D8E8')
            frame.winfo_children()[0].config(bg='#5DADE2' if is_active else '#A8D8E8',
                                             fg='white' if is_active else 'black')
        
        self._clear_page_frame()
        
        if page_name not in ["定时广播", "节假日"]:
            messagebox.showinfo("提示", f"页面 [{page_name}] 正在开发中...")
            self.log(f"功能开发中: {page_name}")
            if self.nav_buttons["定时广播"].cget('bg') != '#5DADE2':
                 self.switch_page("定时广播")
            return

        if page_name == "定时广播":
            self.build_broadcast_view(self.page_frame)
        elif page_name == "节假日":
            self.build_holiday_view(self.page_frame)
            
    # [build_broadcast_view 和 build_holiday_view 方法保持不变]
    def build_broadcast_view(self, parent):
        top_frame = tk.Frame(parent, bg='white')
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
        stats_frame = tk.Frame(parent, bg='#F0F8FF')
        stats_frame.pack(fill=tk.X, padx=10, pady=5)
        self.stats_label = tk.Label(stats_frame, text=f"节目单：{len(self.tasks)}", font=('Microsoft YaHei', 10),
                                   bg='#F0F8FF', fg='#2C5F7C', anchor='w', padx=10)
        self.stats_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        table_frame = tk.Frame(parent, bg='white')
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
        self.update_task_list()
        self.log("切换到定时广播页面。")

    def build_holiday_view(self, parent):
        top_frame = tk.Frame(parent, bg='white')
        top_frame.pack(fill=tk.X, padx=10, pady=10)
        title_label = tk.Label(top_frame, text="节假日", font=('Microsoft YaHei', 14, 'bold'),
                              bg='white', fg='#2C5F7C')
        title_label.pack(side=tk.LEFT)
        desc_frame = tk.Frame(parent, bg='#F0F8FF')
        desc_frame.pack(fill=tk.X, padx=10, pady=5)
        desc_label = tk.Label(desc_frame, text="节假日不播放：在此处添加的时间段内，所有定时广播任务将自动暂停。", 
                              font=('Microsoft YaHei', 10), bg='#F0F8FF', fg='#2C5F7C', anchor='w', padx=10)
        desc_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        content_frame = tk.Frame(parent, bg='white')
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        btn_panel = tk.Frame(content_frame, bg='white')
        btn_panel.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        holiday_buttons = [("添加", self.add_holiday), ("修改", self.edit_holiday), ("删除", self.delete_holiday)]
        for text, cmd in holiday_buttons:
            btn = tk.Button(btn_panel, text=text, command=cmd, font=('Microsoft YaHei', 10),
                          bd=1, bg='#F0F0F0', width=8, pady=4, cursor='hand2')
            btn.pack(pady=4, fill=tk.X)
        table_frame = tk.Frame(content_frame, bg='white')
        table_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        columns = ('节假日名称', '状态', '开始日期时间', '结束日期时间')
        self.holiday_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        col_widths = [200, 80, 200, 200]
        for col, width in zip(columns, col_widths):
            self.holiday_tree.heading(col, text=col)
            self.holiday_tree.column(col, width=width, anchor='center')
        self.holiday_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.holiday_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.holiday_tree.configure(yscrollcommand=scrollbar.set)
        self.holiday_tree.bind("<Double-1>", lambda e: self.edit_holiday())
        self.update_holiday_list()
        self.log("切换到节假日页面。")

    # ...
    # [此处省略大量未修改的方法，以节约篇幅]
    # ...
    def load_holidays(self):
        if not os.path.exists(self.holiday_file): 
            self.holidays = []
            return
        try:
            with open(self.holiday_file, 'r', encoding='utf-8') as f: 
                self.holidays = json.load(f)
            self.log(f"已加载 {len(self.holidays)} 个节假日设置")
        except Exception as e: 
            self.holidays = []
            self.log(f"加载节假日文件失败: {e}")
            
    def save_holidays(self):
        try:
            with open(self.holiday_file, 'w', encoding='utf-8') as f: 
                json.dump(self.holidays, f, ensure_ascii=False, indent=2)
        except Exception as e: 
            self.log(f"保存节假日文件失败: {e}")
            
    def update_holiday_list(self):
        if not hasattr(self, 'holiday_tree'): return
        self.holiday_tree.delete(*self.holiday_tree.get_children())
        for holiday in self.holidays:
            self.holiday_tree.insert('', tk.END, values=(
                holiday.get('name', ''), 
                holiday.get('status', '启用'), 
                holiday.get('start', ''), 
                holiday.get('end', '')
            ))
            
    def add_holiday(self):
        self._open_holiday_dialog()

    def edit_holiday(self):
        if not hasattr(self, 'holiday_tree'): return
        selection = self.holiday_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择一个要修改的节假日。")
            return
        index = self.holiday_tree.index(selection[0])
        holiday_data = self.holidays[index]
        self._open_holiday_dialog(holiday_data, index)

    def delete_holiday(self):
        if not hasattr(self, 'holiday_tree'): return
        selections = self.holiday_tree.selection()
        if not selections:
            messagebox.showwarning("提示", "请先选择要删除的节假日。")
            return
        if messagebox.askyesno("确认删除", f"确定要删除选中的 {len(selections)} 个节假日吗？"):
            indices = sorted([self.holiday_tree.index(s) for s in selections], reverse=True)
            for index in indices:
                self.holidays.pop(index)
            self.save_holidays()
            self.update_holiday_list()
            self.log(f"已删除 {len(selections)} 个节假日。")

    def _open_holiday_dialog(self, holiday_data=None, index=None):
        is_edit = holiday_data is not None
        dialog = tk.Toplevel(self.root)
        dialog.title("修改节假日" if is_edit else "添加节假日")
        dialog.geometry("450x300")
        dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set()
        self.center_window(dialog, 450, 300)
        
        main_frame = tk.Frame(dialog, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(main_frame, text="节假日名称:", font=('Microsoft YaHei', 10)).grid(row=0, column=0, sticky='w', pady=5)
        name_entry = tk.Entry(main_frame, font=('Microsoft YaHei', 10), width=35)
        name_entry.grid(row=0, column=1, sticky='ew', pady=5)
        
        tk.Label(main_frame, text="开始日期时间:", font=('Microsoft YaHei', 10)).grid(row=1, column=0, sticky='w', pady=5)
        start_entry = tk.Entry(main_frame, font=('Microsoft YaHei', 10), width=35)
        start_entry.grid(row=1, column=1, sticky='ew', pady=5)

        tk.Label(main_frame, text="结束日期时间:", font=('Microsoft YaHei', 10)).grid(row=2, column=0, sticky='w', pady=5)
        end_entry = tk.Entry(main_frame, font=('Microsoft YaHei', 10), width=35)
        end_entry.grid(row=2, column=1, sticky='ew', pady=5)
        
        tk.Label(main_frame, text="格式: YYYY-MM-DD HH:MM:SS", font=('Microsoft YaHei', 9), fg='grey').grid(row=3, column=1, sticky='w')

        status_var = tk.StringVar(value="启用")
        status_check = tk.Checkbutton(main_frame, text="启用此规则", variable=status_var, onvalue="启用", offvalue="禁用",
                                      font=('Microsoft YaHei', 10))
        status_check.grid(row=4, column=1, sticky='w', pady=10)

        if is_edit:
            name_entry.insert(0, holiday_data.get('name', ''))
            start_entry.insert(0, holiday_data.get('start', ''))
            end_entry.insert(0, holiday_data.get('end', ''))
            status_var.set(holiday_data.get('status', '启用'))
        else:
            now = datetime.now()
            start_entry.insert(0, now.strftime("%Y-%m-%d 00:00:00"))
            end_entry.insert(0, now.strftime("%Y-%m-%d 23:59:59"))

        def save():
            name = name_entry.get().strip()
            start_str = start_entry.get().strip()
            end_str = end_entry.get().strip()
            status = status_var.get()
            
            if not all([name, start_str, end_str]):
                messagebox.showerror("错误", "所有字段都不能为空。", parent=dialog)
                return
            
            try:
                # --- 修改：加强验证提示 ---
                start_dt = datetime.strptime(start_str, "%Y-%m-%d %H:%M:%S")
                end_dt = datetime.strptime(end_str, "%Y-%m-%d %H:%M:%S")
                if start_dt >= end_dt:
                    messagebox.showerror("逻辑错误", "开始时间必须早于结束时间。", parent=dialog)
                    return
            except ValueError:
                messagebox.showerror("格式错误", 
                                     "日期时间格式不正确。\n\n"
                                     "请严格使用 'YYYY-MM-DD HH:MM:SS' 格式，"
                                     "例如 '2025-10-01 00:00:00'。", 
                                     parent=dialog)
                return
            
            new_data = {'name': name, 'start': start_str, 'end': end_str, 'status': status}
            
            if is_edit:
                self.holidays[index] = new_data
            else:
                self.holidays.append(new_data)
                
            self.save_holidays()
            self.update_holiday_list()
            self.log(f"已{'修改' if is_edit else '添加'}节假日: {name}")
            dialog.destroy()

        btn_frame = tk.Frame(main_frame)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=20)
        tk.Button(btn_frame, text="保存", command=save, width=10).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="取消", command=dialog.destroy, width=10).pack(side=tk.LEFT, padx=10)

    # --- 修改：任务对话框的 save_task 方法 ---
    def _save_task_handler(self, new_task_data, is_edit, index, dialog):
        # 验证和归一化时间
        normalized_time = self._validate_and_normalize_time_list(new_task_data['time'], dialog)
        if normalized_time is None: # 如果验证失败
            return
        new_task_data['time'] = normalized_time

        if not new_task_data['name']:
            messagebox.showwarning("警告", "请填写节目名称。", parent=dialog)
            return
        
        if is_edit:
            self.tasks[index] = new_task_data
            self.log(f"已修改节目: {new_task_data['name']}")
        else:
            self.tasks.append(new_task_data)
            self.log(f"已添加节目: {new_task_data['name']}")
            
        self.update_task_list()
        self.save_tasks()
        dialog.destroy()

    def open_audio_dialog(self, parent_dialog, task_to_edit=None, index=None):
        # ... [内部代码不变，只修改最后的 save_task 函数]
        # ... [此处为简洁省略，请参考之前代码]
        # ...
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
            
            # 使用新的保存处理器
            self._save_task_handler(new_task_data, is_edit_mode, index, dialog)
        # ...
        # ... [其余部分不变] ...
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
        
        button_text = "保存修改" if is_edit_mode else "添加"
        tk.Button(button_frame, text=button_text, command=save_task, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 10, 'bold'), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 10), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        
        content_frame.columnconfigure(1, weight=1)
        time_frame.columnconfigure(1, weight=1)


    def open_voice_dialog(self, parent_dialog, task_to_edit=None, index=None):
        # ... [内部代码不变，只修改最后的 save_task 函数]
        # ... [此处为简洁省略，请参考之前代码]
        # ...
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
            
            # 使用新的保存处理器
            self._save_task_handler(new_task_data, is_edit_mode, index, dialog)
        # ...
        # ... [其余部分不变] ...
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
        
        button_text = "保存修改" if is_edit_mode else "添加"
        tk.Button(button_frame, text=button_text, command=save_task, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 10, 'bold'), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 10), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        
        content_frame.columnconfigure(1, weight=1)
        time_frame.columnconfigure(1, weight=1)

    # --- 修改：时间设置对话框的 confirm 方法 ---
    def show_time_settings_dialog(self, time_entry):
        # ... [内部代码不变，只修改最后的 confirm 函数]
        # ...
        def confirm():
            time_str = ", ".join(list(listbox.get(0, tk.END)))
            # 使用验证器
            normalized_time = self._validate_and_normalize_time_list(time_str, dialog)
            if normalized_time is not None: # 验证通过
                time_entry.delete(0, tk.END)
                time_entry.insert(0, normalized_time)
                dialog.destroy()
        # ...
        # ... [其余部分不变] ...
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
        new_entry.insert(0, datetime.now().strftime("%H:%M:%S")); new_entry.pack(pady=3)
        def add_time():
            val = new_entry.get().strip()
            # 简单验证后添加
            if val:
                listbox.insert(tk.END, val)
        def del_time():
            if listbox.curselection(): listbox.delete(listbox.curselection()[0])
        tk.Button(btn_frame, text="添加 ↑", command=add_time).pack(pady=3, fill=tk.X)
        tk.Button(btn_frame, text="删除", command=del_time).pack(pady=3, fill=tk.X)
        tk.Button(btn_frame, text="清空", command=lambda: listbox.delete(0, tk.END)).pack(pady=3, fill=tk.X)
        bottom_frame = tk.Frame(main_frame, bg='#D7F3F5')
        bottom_frame.pack(pady=10)
        tk.Button(bottom_frame, text="确定", command=confirm, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 9, 'bold'), bd=1, padx=25, pady=5).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 9), bd=1, padx=25, pady=5).pack(side=tk.LEFT, padx=5)

    # ...
    # [此处省略所有未修改的 show_..._dialog, update_..., 和后台逻辑方法]
    # [请保留您自己的这些方法，它们无需修改]
    def show_weekday_settings_dialog(self, weekday_entry):
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
            weekday_entry.delete(0, tk.END); weekday_entry.insert(0, result if selected else ""); dialog.destroy()
        tk.Button(bottom_frame, text="确定", command=confirm, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 9, 'bold'), bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 9), bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)

    def show_daterange_settings_dialog(self, date_range_entry):
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
            start, end = date_range_entry.get().split('~')
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
                date_range_entry.delete(0, tk.END); date_range_entry.insert(0, f"{start} ~ {end}"); dialog.destroy()
            except ValueError: messagebox.showerror("格式错误", "日期格式不正确, 应为 YYYY-MM-DD", parent=dialog)
        tk.Button(bottom_frame, text="确定", command=confirm, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 9, 'bold'), bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 9), bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)
    
    def quit_app(self, icon=None, item=None):
        if self.tray_icon: self.tray_icon.stop()
        self.running = False
        self.save_tasks()
        self.save_holidays()
        if AUDIO_AVAILABLE and pygame.mixer.get_init(): pygame.mixer.quit()
        # 等待后台线程结束
        time.sleep(1.1)
        self.root.destroy()
        sys.exit()

    # ... [所有其他方法，包括后台逻辑 _check_tasks, _execute_broadcast, _play_audio, _speak 等都保持不变]
    def update_task_list(self):
        if not hasattr(self, 'task_tree'): return
        selection = self.task_tree.selection()
        self.task_tree.delete(*self.task_tree.get_children())
        for task in self.tasks:
            content = task.get('content', '')
            content_preview = os.path.basename(content) if task.get('type') == 'audio' else (content[:30] + '...' if len(content) > 30 else content)
            display_mode = "准时" if task.get('delay') == 'ontime' else "延时"
            self.task_tree.insert('', tk.END, values=(
                task.get('name', ''), task.get('status', ''), task.get('time', ''),
                display_mode, content_preview, task.get('volume', ''),
                task.get('weekday', ''), task.get('date_range', '')
            ))
        if selection:
            try: self.task_tree.selection_set(selection)
            except tk.TclError: pass
        if hasattr(self, 'stats_label'): self.stats_label.config(text=f"节目单：{len(self.tasks)}")
        if hasattr(self, 'status_labels') and self.status_labels: self.status_labels[3].config(text=f"任务数量: {len(self.tasks)}")

    def update_status_bar(self):
        if not self.running or not hasattr(self, 'status_labels') or not self.status_labels: return
        self.status_labels[0].config(text=f"当前时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.status_labels[1].config(text="系统状态: 运行中")
        self.root.after(1000, self.update_status_bar)

    def start_background_thread(self):
        threading.Thread(target=self._check_tasks, daemon=True).start()

    def _check_tasks(self):
        while self.running:
            if self._is_currently_holiday():
                time.sleep(1)
                continue

            now = datetime.now()
            current_date_str = now.strftime("%Y-%m-%d")
            current_time_str = now.strftime("%H:%M:%S")

            for task in self.tasks:
                if task.get('status') != '启用': continue
                try:
                    start, end = [d.strip() for d in task.get('date_range', '').split('~')]
                    if not (datetime.strptime(start, "%Y-%m-%d").date() <= now.date() <= datetime.strptime(end, "%Y-%m-%d").date()): continue
                except (ValueError, IndexError): pass
                schedule = task.get('weekday', '每周:1234567')
                run_today = (schedule.startswith("每周:") and str(now.isoweekday()) in schedule[3:]) or \
                            (schedule.startswith("每月:") and f"{now.day:02d}" in schedule[3:].split(','))
                if not run_today: continue
                for trigger_time in [t.strip() for t in task.get('time', '').split(',')]:
                    if trigger_time == current_time_str and task.get('last_run', {}).get(trigger_time) != current_date_str:
                        if task.get('delay') == 'ontime':
                            self.log(f"准时任务 '{task['name']}' 已到时间，执行高优先级中断。")
                            self.root.after(0, self._force_stop_playback)
                            with self.queue_lock:
                                self.playback_queue.clear()
                                self.playback_queue.insert(0, (task, trigger_time))
                            self.root.after(50, self._process_queue)
                        else:
                            with self.queue_lock:
                                self.playback_queue.append((task, trigger_time))
                            self.log(f"延时任务 '{task['name']}' 已到时间，加入播放队列。")
                            self.root.after(0, self._process_queue)
            time.sleep(1)

    def _process_queue(self):
        if self.is_playing.is_set(): return
        with self.queue_lock:
            if not self.playback_queue: return
            task, trigger_time = self.playback_queue.pop(0)
        self._execute_broadcast(task, trigger_time)

    def _execute_broadcast(self, task, trigger_time):
        self.is_playing.set()
        self.update_playing_text(f"[{task['name']}] 正在准备播放...")
        if hasattr(self, 'status_labels') and self.status_labels:
            self.status_labels[2].config(text="播放状态: 播放中")
        if trigger_time != "manual_play":
            if not isinstance(task.get('last_run'), dict):
                task['last_run'] = {}
            task['last_run'][trigger_time] = datetime.now().strftime("%Y-%m-%d")
            self.save_tasks()
        if task.get('type') == 'audio':
            self.log(f"开始音频任务: {task['name']}")
            threading.Thread(target=self._play_audio, args=(task,), daemon=True).start()
        else:
            self.log(f"开始语音任务: {task['name']} (共 {task.get('repeat', 1)} 遍)")
            threading.Thread(target=self._speak, args=(task.get('content', ''), task), daemon=True).start()

    def _play_audio(self, task):
        try:
            interval_type = task.get('interval_type')
            duration_seconds = int(task.get('interval_seconds', 0))
            repeat_count = int(task.get('interval_first', 1))
            playlist = []
            if task.get('audio_type') == 'single':
                if os.path.exists(task['content']):
                    playlist = [task['content']] * repeat_count
            else:
                folder_path = task['content']
                if os.path.isdir(folder_path):
                    all_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.lower().endswith(('.mp3', '.wav', '.ogg', '.flac', '.m4a'))]
                    if task.get('play_order') == 'random':
                        random.shuffle(all_files)
                    playlist = all_files[:repeat_count]

            if not playlist:
                self.log(f"错误: 音频列表为空，任务 '{task['name']}' 无法播放。"); return

            start_time = time.time()
            for audio_path in playlist:
                self.log(f"正在播放: {os.path.basename(audio_path)}")
                self.update_playing_text(f"[{task['name']}] 正在播放: {os.path.basename(audio_path)}")
                pygame.mixer.music.load(audio_path)
                pygame.mixer.music.set_volume(float(task.get('volume', 80)) / 100.0)
                pygame.mixer.music.play()
                while pygame.mixer.music.get_busy():
                    if not self.running: break
                    if interval_type == 'seconds' and (time.time() - start_time) > duration_seconds:
                        pygame.mixer.music.stop()
                        self.log(f"已达到 {duration_seconds} 秒播放时长限制。")
                        break
                    time.sleep(0.1)
                if not self.running or (interval_type == 'seconds' and (time.time() - start_time) > duration_seconds):
                    break
        except Exception as e:
            self.log(f"音频播放错误: {e}")
        finally:
            if self.running:
                self.root.after(0, self.on_playback_finished)

    def _speak(self, text, task):
        if not WIN32COM_AVAILABLE:
            self.log("错误: pywin32库不可用，无法执行语音播报。")
            if self.running: self.root.after(0, self.on_playback_finished)
            return
        pythoncom.CoInitialize()
        try:
            if task.get('bgm', 0) and AUDIO_AVAILABLE:
                bgm_file = task.get('bgm_file', '')
                bgm_path = os.path.join(BGM_FOLDER, bgm_file)
                if os.path.exists(bgm_path):
                    pygame.mixer.music.load(bgm_path)
                    pygame.mixer.music.set_volume(float(task.get('bgm_volume', 40)) / 100.0)
                    pygame.mixer.music.play(-1)
            if task.get('prompt', 0) and AUDIO_AVAILABLE:
                prompt_file = task.get('prompt_file', '')
                prompt_path = os.path.join(PROMPT_FOLDER, prompt_file)
                if os.path.exists(prompt_path):
                    sound = pygame.mixer.Sound(prompt_path)
                    sound.set_volume(float(task.get('prompt_volume', 80)) / 100.0)
                    channel = sound.play()
                    if channel:
                        while channel.get_busy() and self.running: time.sleep(0.05)
            
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            all_voices = {v.GetDescription(): v for v in speaker.GetVoices()}
            selected_voice_desc = task.get('voice')
            if selected_voice_desc in all_voices:
                speaker.Voice = all_voices[selected_voice_desc]
            speaker.Volume = int(task.get('volume', 80))
            rate = task.get('speed', '0')
            pitch = task.get('pitch', '0')
            escaped_text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("'", "&apos;").replace('"', "&quot;")
            xml_text = f"<rate absspeed='{rate}'><pitch middle='{pitch}'>{escaped_text}</pitch></rate>"
            repeat_count = int(task.get('repeat', 1))

            for i in range(repeat_count):
                if not self.running: break
                self.log(f"正在播报第 {i+1}/{repeat_count} 遍")
                speaker.Speak(xml_text, 8)
                if i < repeat_count - 1: time.sleep(0.5)

        except Exception as e:
            self.log(f"播报错误: {e}")
        finally:
            if AUDIO_AVAILABLE and pygame.mixer.music.get_busy():
                pygame.mixer.music.stop()
            pythoncom.CoUninitialize()
            if self.running:
                self.root.after(0, self.on_playback_finished)
    
    def on_playback_finished(self):
        # 确保在主线程中安全地调用
        if not self.running: return
        self.is_playing.clear()
        self.update_playing_text("等待下一个任务...")
        if hasattr(self, 'status_labels') and self.status_labels:
            self.status_labels[2].config(text="播放状态: 待机")
        self.log("播放结束")
        self.root.after(100, self._process_queue)
    
    def hide_to_tray(self):
        self.root.withdraw()
        if not self.tray_icon and TRAY_AVAILABLE:
            self.setup_tray_icon()
            threading.Thread(target=self.tray_icon.run, daemon=True).start()
            self.log("程序已最小化到系统托盘。")

    def show_from_tray(self, icon, item):
        self.tray_icon.stop()
        self.tray_icon = None
        self.root.after(0, self.root.deiconify)
        self.log("程序已从托盘恢复。")

    def setup_tray_icon(self):
        try:
            image = Image.open(ICON_FILE)
        except Exception as e:
            image = Image.new('RGB', (64, 64), 'white')
            print(f"警告: 未找到或无法加载图标文件 '{ICON_FILE}': {e}")
        menu = (item('显示', self.show_from_tray, default=True), item('退出', self.quit_app))
        self.tray_icon = Icon("boyin", image, "定时播音", menu)
        
def main():
    root = tk.Tk()
    app = TimedBroadcastApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
