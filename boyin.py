import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import pyttsx3
import json
import threading
import time
from datetime import datetime
import os

# 尝试导入 win32com，这是获取语音列表的最佳方式
try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False


# --- 全局设置 ---
# 获取脚本所在的目录，确保所有文件路径都是基于此目录
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TASK_FILE = os.path.join(SCRIPT_DIR, "broadcast_tasks.json")
PROMPT_FOLDER = os.path.join(SCRIPT_DIR, "提示音")
AUDIO_FOLDER = os.path.join(SCRIPT_DIR, "音频文件")

# 音频播放库
AUDIO_AVAILABLE = False
try:
    import pygame
    pygame.mixer.init()
    AUDIO_AVAILABLE = True
except ImportError:
    print("警告: 未安装 pygame 库 (pip install pygame)，音频播放功能将不可用。")
except Exception as e:
    print(f"警告: pygame 初始化失败 - {e}，音频播放功能将不可用。")

class TimedBroadcastApp:
    def __init__(self, root):
        self.root = root
        self.root.title("定时播音")
        self.root.geometry("1400x800")
        self.root.configure(bg='#E8F4F8')

        # 初始化语音引擎
        self.engine = None
        try:
            self.engine = pyttsx3.init(driverName='sapi5')
        except Exception as e:
            print(f"严重错误：主语音引擎 pyttsx3 初始化失败 - {e}。语音播报功能将不可用。")
            messagebox.showerror("严重错误", f"主语音引擎初始化失败: {e}\n语音播报功能将不可用。")

        # 任务列表
        self.tasks = []
        self.running = True
        self.task_file = TASK_FILE
        self.current_page = "定时广播"

        # 创建必要的文件夹结构
        self.create_folder_structure()

        # 创建界面
        self.create_widgets()

        # 加载已保存的任务
        self.load_tasks()

        # 启动后台检查线程
        self.start_background_thread()
        
        # 绑定关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_folder_structure(self):
        """创建必要的文件夹结构"""
        folders = [PROMPT_FOLDER, AUDIO_FOLDER]
        for folder in folders:
            if not os.path.exists(folder):
                os.makedirs(folder)
                self.log(f"已创建文件夹: {folder}") if hasattr(self, 'log_text') else None

    def create_widgets(self):
        # 左侧导航栏
        self.nav_frame = tk.Frame(self.root, bg='#A8D8E8', width=160)
        self.nav_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.nav_frame.pack_propagate(False)

        nav_buttons = [
            ("定时广播", ""),
            ("立即播播", ""),
            ("节假日、调休", "节假日不播或、调休"),
            ("设置", ""),
            ("语音广告 制作", "")
        ]

        for i, (title, subtitle) in enumerate(nav_buttons):
            btn_frame = tk.Frame(self.nav_frame, bg='#5DADE2' if i == 0 else '#A8D8E8')
            btn_frame.pack(fill=tk.X, pady=1)
            
            btn = tk.Button(btn_frame, text=title, bg='#5DADE2' if i == 0 else '#A8D8E8',
                          fg='white' if i == 0 else 'black', font=('Microsoft YaHei', 13, 'bold'),
                          bd=0, padx=10, pady=8, anchor='w',
                          command=lambda t=title: self.switch_page(t))
            btn.pack(fill=tk.X)

            if subtitle:
                sub_label = tk.Label(btn_frame, text=subtitle, bg='#5DADE2' if i == 0 else '#A8D8E8',
                                   fg='#FF6B35' if title == "节假日、调休" else ('#555' if i == 0 else '#666'),
                                   font=('Microsoft YaHei', 10), anchor='w', padx=10)
                sub_label.pack(fill=tk.X)

        # 主内容区域
        self.main_frame = tk.Frame(self.root, bg='white')
        self.main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 创建定时广播页面
        self.create_scheduled_broadcast_page()

    def switch_page(self, page_name):
        """切换页面"""
        self.current_page = page_name
        if page_name == "定时广播":
            self.log(f"切换到: {page_name}")
        else:
            messagebox.showinfo("提示", f"页面 [{page_name}] 正在开发中...")
            self.log(f"功能开发中: {page_name}")

    def create_scheduled_broadcast_page(self):
        """创建定时广播页面"""
        top_frame = tk.Frame(self.main_frame, bg='white')
        top_frame.pack(fill=tk.X, padx=10, pady=10)

        title_label = tk.Label(top_frame, text="定时广播", font=('Microsoft YaHei', 14, 'bold'),
                              bg='white', fg='#2C5F7C')
        title_label.pack(side=tk.LEFT)

        btn_frame = tk.Frame(top_frame, bg='white')
        btn_frame.pack(side=tk.RIGHT)

        buttons = [
            ("添加节目", self.add_task, '#5DADE2'),
            ("删除", self.delete_task, '#E74C3C'),
            ("修改", self.edit_task, '#F39C12'),
            ("复制", self.copy_task, '#9B59B6'),
            ("上移", lambda: self.move_task(-1), '#3498DB'),
            ("下移", lambda: self.move_task(1), '#3498DB'),
            ("导入", self.import_tasks, '#1ABC9C'),
            ("导出", self.export_tasks, '#1ABC9C'),
            ("启用", self.enable_task, '#27AE60'),
            ("禁用", self.disable_task, '#95A5A6')
        ]

        for text, cmd, color in buttons:
            btn = tk.Button(btn_frame, text=text, command=cmd, bg=color, fg='white',
                          font=('Microsoft YaHei', 9), bd=0, padx=12, pady=5, cursor='hand2')
            btn.pack(side=tk.LEFT, padx=3)

        stats_frame = tk.Frame(self.main_fram
