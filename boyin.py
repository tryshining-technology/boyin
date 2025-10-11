import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import json
import threading
import time
from datetime import datetime
import os
import random
import sys
import getpass

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
        
        self.pages = {}
        self.nav_buttons = {}
        self.current_page = None

        # 【新功能】创建并加载内嵌的菜单图标
        self.menu_icons = {}
        self._create_menu_icons()

        self.create_folder_structure()
        self.load_settings()
        self.create_widgets()
        self.load_tasks()
        
        self.start_background_thread()
        self.root.protocol("WM_DELETE_WINDOW", self.show_quit_dialog)
        self.start_tray_icon_thread()
        
        if self.settings.get("start_minimized", False):
            self.root.after(100, self.hide_to_tray)

    def _create_menu_icons(self):
        """创建内嵌的Base64编码图标，用于菜单的完美对齐"""
        icons_b64 = {
            "play": "R0lGODlhEAAQAPcAAAEBAQMDAxsbGzU1NUNDQ0tLS1hYWGhoaHt7e4SEhIyMjJSUlJycnKWlpaysrL29vc/Pz9/f3+fn5+/v7/f39wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAB8ALAAAAAAQABAAAAjlABEJGEiwoMGDCBMqXMiwocOHECNKnEixosWLGDNq3Mixo8ePIEOKHEmypMmTKFOqXMmypcuXMGPKnEmzps2bOHPq3Mmzp8+fQIMKHUq0qNGjSJMqXcq0qtWrWLNq3cq1q9evYMOKHUu2rNmzaNOqXcu2rdu3cOPKnUu3rt27ePPq3cu3r9+/gAMLHky4sOHDiBMrXsy4sePHkCNLnky5suXLmDNr3sy5s+fPoEOLHk26tOnTqFOrXs26tevXsGPLnk27tu3buHPr3s27t+/fwIMLH068uPHjyJMrX868ufPn0KNLn069uvXr2LNr3869u/fv4MOLH0++vPnz6NOrX8++vfv38OPLn0+/vv37+PPr38+/v///AAYo5IADhQAAOw==",
            "edit": "R0lGODlhEAAQAPcAAAEBAQMDAxsbGzU1NUNDQ0tLS1hYWGhoaHt7e4SEhIyMjJSUlJycnKWlpaysrL29vc/Pz9/f3+fn5+/v7/f39wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAB8ALAAAAAAQABAAAAjRABEJGEiwoMGDCBMqXMiwocOHECNKnEixosWLGDNq3Mixo8ePIEOKHEmypMmTKFOqXMmypcuXMGPKnEmzps2bOHPq3Mmzp8+fQIMKHUq0qNGjSJMqXcq0qtWrWLNq3cq1q9evYMOKHUu2rNmzaNOqXcu2rdu3cOPKnUu3rt27ePPq3cu3r9+/gAMLHky4sOHDiBMrXsy4sePHkCNLnky5suXLmDNr3sy5s+fPoEOLHk26tOnTqFOrXs26tevXsGPLnk27tu3buHPr3s27t+/fwIMLH068uPHjyJMrX868ufPn0KNLn069uvXr2LNr3869u/fv4MOLH0+/vPnz6NOrX8++vfv38OPLn0+/vv37+PPr38+/v///AAYo5IADhQAAOw==",
            "delete": "R0lGODlhEAAQAPcAAAEBAQMDAxsbGzU1NUNDQ0tLS1hYWGhoaHt7e4SEhIyMjJSUlJycnKWlpaysrL29vc/Pz9/f3+fn5+/v7/f39wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAB8ALAAAAAAQABAAAAjlABEJGEiwoMGDCBMqXMiwocOHECNKnEixosWLGDNq3Mixo8ePIEOKHEmypMmTKFOqXMmypcuXMGPKnEmzps2bOHPq3Mmzp8+fQIMKHUq0qNGjSJMqXcq0qtWrWLNq3cq1q9evYMOKHUu2rNmzaNOqXcu2rdu3cOPKnUu3rt27ePPq3cu3r9+/gAMLHky4sOHDiBMrXsy4sePHkCNLnky5suXLmDNr3sy5s+fPoEOLHk26tOnTqFOrXs26tevXsGPLnk27tu3buHPr3s27t+/fwIMLH068uPHjyJMrX868ufPn0KNLn069uvXr2LNr3869u/fv4MOLH0+/vPnz6NOrX8++vfv38OPLn0+/vv37+PPr38+/v///AAYo5IADhQAAOw==",
            "copy": "R0lGODlhEAAQAPcAAAEBAQMDAxsbGzU1NUNDQ0tLS1hYWGhoaHt7e4SEhIyMjJSUlJycnKWlpaysrL29vc/Pz9/f3+fn5+/v7/f39wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAB8ALAAAAAAQABAAAAjRABEJGEiwoMGDCBMqXMiwocOHECNKnEixosWLGDNq3Mixo8ePIEOKHEmypMmTKFOqXMmypcuXMGPKnEmzps2bOHPq3Mmzp8+fQIMKHUq0qNGjSJMqXcq0qtWrWLNq3cq1q9evYMOKHUu2rNmzaNOqXcu2rdu3cOPKnUu3rt27ePPq3cu3r9+/gAMLHky4sOHDiBMrXsy4sePHkCNLnky5suXLmDNr3sy5s+fPoEOLHk26tOnTqFOrXs26tevXsGPLnk27tu3buHPr3s27t+/fwIMLH068uPHjyJMrX868ufPn0KNLn069uvXr2LNr3869u/fv4MOLH0+/vPnz6NOrX8++vfv38OPLn0+/vv37+PPr38+/v///AAYo5IADhQAAOw==",
            "top": "R0lGODlhEAAQAPcAAAEBAQMDAxsbGzU1NUNDQ0tLS1hYWGhoaHt7e4SEhIyMjJSUlJycnKWlpaysrL29vc/Pz9/f3+fn5+/v7/f39wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAB8ALAAAAAAQABAAAAjRABEJGEiwoMGDCBMqXMiwocOHECNKnEixosWLGDNq3Mixo8ePIEOKHEmypMmTKFOqXMmypcuXMGPKnEmzps2bOHPq3Mmzp8+fQIMKHUq0qNGjSJMqXcq0qtWrWLNq3cq1q9evYMOKHUu2rNmzaNOqXcu2rdu3cOPKnUu3rt27ePPq3cu3r9+/gAMLHky4sOHDiBMrXsy4sePHkCNLnky5suXLmDNr3sy5s+fPoEOLHk26tOnTqFOrXs26tevXsGPLnk27tu3buHPr3s27t+/fwIMLH068uPHjyJMrX868ufPn0KNLn069uvXr2LNr3869u/fv4MOLH0+/vPnz6NOrX8++vfv38OPLn0+/vv37+PPr38+/v///AAYo5IADhQAAOw==",
            "up": "R0lGODlhEAAQAPcAAAEBAQMDAxsbGzU1NUNDQ0tLS1hYWGhoaHt7e4SEhIyMjJSUlJycnKWlpaysrL29vc/Pz9/f3+fn5+/v7/f39wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAB8ALAAAAAAQABAAAAjRABEJGEiwoMGDCBMqXMiwocOHECNKnEixosWLGDNq3Mixo8ePIEOKHEmypMmTKFOqXMmypcuXMGPKnEmzps2bOHPq3Mmzp8+fQIMKHUq0qNGjSJMqXcq0qtWrWLNq3cq1q9evYMOKHUu2rNmzaNOqXcu2rdu3cOPKnUu3rt27ePPq3cu3r9+/gAMLHky4sOHDiBMrXsy4sePHkCNLnky5suXLmDNr3sy5s+fPoEOLHk26tOnTqFOrXs26tevXsGPLnk27tu3buHPr3s27t+/fwIMLH068uPHjyJMrX868ufPn0KNLn069uvXr2LNr3869u/fv4MOLH0+/vPnz6NOrX8++vfv38OPLn0+/vv37+PPr38+/v///AAYo5IADhQAAOw==",
            "down": "R0lGODlhEAAQAPcAAAEBAQMDAxsbGzU1NUNDQ0tLS1hYWGhoaHt7e4SEhIyMjJSUlJycnKWlpaysrL29vc/Pz9/f3+fn5+/v7/f39wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAB8ALAAAAAAQABAAAAjRABEJGEiwoMGDCBMqXMiwocOHECNKnEixosWLGDNq3Mixo8ePIEOKHEmypMmTKFOqXMmypcuXMGPKnEmzps2bOHPq3Mmzp8+fQIMKHUq0qNGjSJMqXcq0qtWrWLNq3cq1q9evYMOKHUu2rNmzaNOqXcu2rdu3cOPKnUu3rt27ePPq3cu3r9+/gAMLHky4sOHDiBMrXsy4sePHkCNLnky5suXLmDNr3sy5s+fPoEOLHk26tOnTqFOrXs26tevXsGPLnk27tu3buHPr3s27t+/fwIMLH068uPHjyJMrX868ufPn0KNLn069uvXr2LNr3869u/fv4MOLH0+/vPnz6NOrX8++vfv38OPLn0+/vv37+PPr38+/v///AAYo5IADhQAAOw==",
            "bottom": "R0lGODlhEAAQAPcAAAEBAQMDAxsbGzU1NUNDQ0tLS1hYWGhoaHt7e4SEhIyMjJSUlJycnKWlpaysrL29vc/Pz9/f3+fn5+/v7/f39wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAB8ALAAAAAAQABAAAAjRABEJGEiwoMGDCBMqXMiwocOHECNKnEixosWLGDNq3Mixo8ePIEOKHEmypMmTKFOqXMmypcuXMGPKnEmzps2bOHPq3Mmzp8+fQIMKHUq0qNGjSJMqXcq0qtWrWLNq3cq1q9evYMOKHUu2rNmzaNOqXcu2rdu3cOPKnUu3rt27ePPq3cu3r9+/gAMLHky4sOHDiBMrXsy4sePHkCNLnky5suXLmDNr3sy5s+fPoEOLHk26tOnTqFOrXs26tevXsGPLnk27tu3buHPr3s27t+/fwIMLH068uPHjyJMrX868ufPn0KNLn069uvXr2LNr3869u/fv4MOLH0+/vPnz6NOrX8++vfv38OPLn0+/vv37+PPr38+/v///AAYo5IADhQAAOw==",
            "enable": "R0lGODlhEAAQAPcAAAEBAQMDAxsbGzU1NUNDQ0tLS1hYWGhoaHt7e4SEhIyMjJSUlJycnKWlpaysrL29vc/Pz9/f3+fn5+/v7/f39wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAB8ALAAAAAAQABAAAAjlABEJGEiwoMGDCBMqXMiwocOHECNKnEixosWLGDNq3Mixo8ePIEOKHEmypMmTKFOqXMmypcuXMGPKnEmzps2bOHPq3Mmzp8+fQIMKHUq0qNGjSJMqXcq0qtWrWLNq3cq1q9evYMOKHUu2rNmzaNOqXcu2rdu3cOPKnUu3rt27ePPq3cu3r9+/gAMLHky4sOHDiBMrXsy4sePHkCNLnky5suXLmDNr3sy5s+fPoEOLHk26tOnTqFOrXs26tevXsGPLnk27tu3buHPr3s27t+/fwIMLH068uPHjyJMrX868ufPn0KNLn069uvXr2LNr3869u/fv4MOLH0+/vPnz6NOrX8++vfv38OPLn0+/vv37+PPr38+/v///AAYo5IADhQAAOw==",
            "disable": "R0lGODlhEAAQAPcAAAEBAQMDAxsbGzU1NUNDQ0tLS1hYWGhoaHt7e4SEhIyMjJSUlJycnKWlpaysrL29vc/Pz9/f3+fn5+/v7/f39wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAB8ALAAAAAAQABAAAAjlABEJGEiwoMGDCBMqXMiwocOHECNKnEixosWLGDNq3Mixo8ePIEOKHEmypMmTKFOqXMmypcuXMGPKnEmzps2bOHPq3Mmzp8+fQIMKHUq0qNGjSJMqXcq0qtWrWLNq3cq1q9evYMOKHUu2rNmzaNOqXcu2rdu3cOPKnUu3rt27ePPq3cu3r9+/gAMLHky4sOHDiBMrXsy4sePHkCNLnky5suXLmDNr3sy5s+fPoEOLHk26tOnTqFOrXs26tevXsGPLnk27tu3buHPr3s27t+/fwIMLH068uPHjyJMrX868ufPn0KNLn069uvXr2LNr3869u/fv4MOLH0+/vPnz6NOrX8++vfv38OPLn0+/vv37+PPr38+/v///AAYo5IADhQAAOw==",
            "add": "R0lGODlhEAAQAPcAAAEBAQMDAxsbGzU1NUNDQ0tLS1hYWGhoaHt7e4SEhIyMjJSUlJycnKWlpaysrL29vc/Pz9/f3+fn5+/v7/f39wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAB8ALAAAAAAQABAAAAjRABEJGEiwoMGDCBMqXMiwocOHECNKnEixosWLGDNq3Mixo8ePIEOKHEmypMmTKFOqXMmypcuXMGPKnEmzps2bOHPq3Mmzp8+fQIMKHUq0qNGjSJMqXcq0qtWrWLNq3cq1q9evYMOKHUu2rNmzaNOqXcu2rdu3cOPKnUu3rt27ePPq3cu3r9+/gAMLHky4sOHDiBMrXsy4sePHkCNLnky5suXLmDNr3sy5s+fPoEOLHk26tOnTqFOrXs26tevXsGPLnk27tu3buHPr3s27t+/fwIMLH068uPHjyJMrX868ufPn0KNLn069uvXr2LNr3869u/fv4MOLH0+/vPnz6NOrX8++vfv38OPLn0+/vv37+PPr38+/v///AAYo5IADhQAAOw==",
            "stop": "R0lGODlhEAAQAPcAAAEBAQMDAxsbGzU1NUNDQ0tLS1hYWGhoaHt7e4SEhIyMjJSUlJycnKWlpaysrL29vc/Pz9/f3+fn5+/v7/f39wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAB8ALAAAAAAQABAAAAjlABEJGEiwoMGDCBMqXMiwocOHECNKnEixosWLGDNq3Mixo8ePIEOKHEmypMmTKFOqXMmypcuXMGPKnEmzps2bOHPq3Mmzp8+fQIMKHUq0qNGjSJMqXcq0qtWrWLNq3cq1q9evYMOKHUu2rNmzaNOqXcu2rdu3cOPKnUu3rt27ePPq3cu3r9+/gAMLHky4sOHDiBMrXsy4sePHkCNLnky5suXLmDNr3sy5s+fPoEOLHk26tOnTqFOrXs26tevXsGPLnk27tu3buHPr3s27t+/fwIMLH068uPHjyJMrX868ufPn0KNLn069uvXr2LNr3869u/fv4MOLH0+/vPnz6NOrX8++vfv38OPLn0+/vv37+PPr38+/v///AAYo5IADhQAAOw=="
        }
        for name, data in icons_b64.items():
            self.menu_icons[name] = tk.PhotoImage(data=data)

    def create_folder_structure(self):
        """创建所有必要的文件夹"""
        for folder in [PROMPT_FOLDER, AUDIO_FOLDER, BGM_FOLDER]:
            if not os.path.exists(folder):
                os.makedirs(folder)

    def create_widgets(self):
        self.nav_frame = tk.Frame(self.root, bg='#A8D8E8', width=160)
        self.nav_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.nav_frame.pack_propagate(False)

        # 【修改】移除 "语音广告 制作"
        nav_button_titles = ["定时广播", "节假日", "设置"]
        
        for i, title in enumerate(nav_button_titles):
            btn_frame = tk.Frame(self.nav_frame, bg='#A8D8E8')
            btn_frame.pack(fill=tk.X, pady=1)
            btn = tk.Button(btn_frame, text=title, bg='#A8D8E8',
                          fg='black', font=('Microsoft YaHei', 13, 'bold'),
                          bd=0, padx=10, pady=8, anchor='w', command=lambda t=title: self.switch_page(t))
            btn.pack(fill=tk.X)
            self.nav_buttons[title] = btn
        
        self.main_frame = tk.Frame(self.root, bg='white')
        self.pages["定时广播"] = self.main_frame
        self.create_scheduled_broadcast_page()

        self.current_page = self.main_frame
        self.switch_page("定时广播")

    def switch_page(self, page_name):
        if self.is_locked and page_name != "设置":
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
            self.lock_now_var.set(self.is_locked)
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
        title_label = tk.Label(top_frame, text="定时广播", font=('Microsoft YaHei', 14, 'bold'),
                              bg='white', fg='#2C5F7C')
        title_label.pack(side=tk.LEFT)
        
        self.top_right_btn_frame = tk.Frame(top_frame, bg='white')
        self.top_right_btn_frame.pack(side=tk.RIGHT)
        
        self.lock_button = tk.Button(self.top_right_btn_frame, text="锁定", command=self.toggle_lock_state, bg='#E74C3C', fg='white',
                                     font=('Microsoft YaHei', 9), bd=0, padx=12, pady=5, cursor='hand2')
        self.lock_button.pack(side=tk.LEFT, padx=3)

        buttons = [("导入节目单", self.import_tasks, '#1ABC9C'), ("导出节目单", self.export_tasks, '#1ABC9C')]
        for text, cmd, color in buttons:
            btn = tk.Button(self.top_right_btn_frame, text=text, command=cmd, bg=color, fg='white',
                          font=('Microsoft YaHei', 9), bd=0, padx=12, pady=5, cursor='hand2')
            btn.pack(side=tk.LEFT, padx=3)

        stats_frame = tk.Frame(self.main_frame, bg='#F0F8FF')
        stats_frame.pack(fill=tk.X, padx=10, pady=5)
        self.stats_label = tk.Label(stats_frame, text="节目单：0", font=('Microsoft YaHei', 10),
                                   bg='#F0F8FF', fg='#2C5F7C', anchor='w', padx=10)
        self.stats_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        table_frame = tk.Frame(self.main_frame, bg='white')
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        columns = ('节目名称', '状态', '开始时间', '模式', '音频或文字', '音量', '周几/几号', '日期范围')
        self.task_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=12)
        
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

        playing_frame = tk.LabelFrame(self.main_frame, text="正在播：", font=('Microsoft YaHei', 10),
                                     bg='white', fg='#2C5F7C', padx=10, pady=5)
        playing_frame.pack(fill=tk.X, padx=10, pady=5)
        self.playing_text = scrolledtext.ScrolledText(playing_frame, height=3, font=('Microsoft YaHei', 9),
                                                     bg='#FFFEF0', wrap=tk.WORD, state='disabled')
        self.playing_text.pack(fill=tk.BOTH, expand=True)
        self.update_playing_text("等待播放...")

        log_frame = tk.LabelFrame(self.main_frame, text="", font=('Microsoft YaHei', 10),
                                 bg='white', fg='#2C5F7C', padx=10, pady=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        log_header_frame = tk.Frame(log_frame, bg='white')
        log_header_frame.pack(fill=tk.X)
        log_label = tk.Label(log_header_frame, text="日志：", font=('Microsoft YaHei', 10, 'bold'),
                             bg='white', fg='#2C5F7C')
        log_label.pack(side=tk.LEFT)
        self.clear_log_btn = tk.Button(log_header_frame, text="清除日志", command=self.clear_log,
                                       font=('Microsoft YaHei', 8), bd=0, bg='#EAEAEA',
                                       fg='#333', cursor='hand2', padx=5, pady=0)
        self.clear_log_btn.pack(side=tk.LEFT, padx=10)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, font=('Microsoft YaHei', 9),
                                                 bg='#F9F9F9', wrap=tk.WORD, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True)

        status_frame = tk.Frame(self.main_frame, bg='#E8F4F8', height=30)
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
        self.log("定时播音软件已启动")
    
    def create_settings_page(self):
        settings_frame = tk.Frame(self.root, bg='white')

        title_label = tk.Label(settings_frame, text="系统设置", font=('Microsoft YaHei', 14, 'bold'),
                               bg='white', fg='#2C5F7C')
        title_label.pack(anchor='w', padx=20, pady=20)

        general_frame = tk.LabelFrame(settings_frame, text="通用设置", font=('Microsoft YaHei', 11, 'bold'),
                                      bg='white', padx=15, pady=10)
        general_frame.pack(fill=tk.X, padx=20, pady=10)
        
        self.autostart_var = tk.BooleanVar(value=self.settings.get("autostart", False))
        self.start_minimized_var = tk.BooleanVar(value=self.settings.get("start_minimized", False))
        self.lock_now_var = tk.BooleanVar(value=self.is_locked)

        tk.Checkbutton(general_frame, text="登录windows后自动启动", variable=self.autostart_var, 
                       font=('Microsoft YaHei', 10), bg='white', anchor='w', 
                       command=self._handle_autostart_setting).pack(fill=tk.X, pady=5)
        tk.Checkbutton(general_frame, text="启动后最小化到系统托盘", variable=self.start_minimized_var,
                       font=('Microsoft YaHei', 10), bg='white', anchor='w',
                       command=self.save_settings).pack(fill=tk.X, pady=5)
        tk.Checkbutton(general_frame, text="立即锁定", variable=self.lock_now_var,
                       font=('Microsoft YaHei', 10), bg='white', anchor='w',
                       command=self.toggle_lock_state).pack(fill=tk.X, pady=5)
        
        power_frame = tk.LabelFrame(settings_frame, text="电源管理", font=('Microsoft YaHei', 11, 'bold'),
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
                       font=('Microsoft YaHei', 10), bg='white', command=self.save_settings).pack(side=tk.LEFT)
        time_entry_daily = tk.Entry(daily_frame, textvariable=self.daily_shutdown_time_var, 
                                    font=('Microsoft YaHei', 10), width=15)
        time_entry_daily.pack(side=tk.LEFT, padx=10)
        tk.Button(daily_frame, text="设置", command=lambda: self.show_single_time_dialog(self.daily_shutdown_time_var)
                  ).pack(side=tk.LEFT)

        weekly_frame = tk.Frame(power_frame, bg='white')
        weekly_frame.pack(fill=tk.X, pady=4)
        tk.Checkbutton(weekly_frame, text="每周关机", variable=self.weekly_shutdown_enabled_var, 
                       font=('Microsoft YaHei', 10), bg='white', command=self.save_settings).pack(side=tk.LEFT)
        days_entry_weekly = tk.Entry(weekly_frame, textvariable=self.weekly_shutdown_days_var,
                                     font=('Microsoft YaHei', 10), width=20)
        days_entry_weekly.pack(side=tk.LEFT, padx=(10,5))
        time_entry_weekly = tk.Entry(weekly_frame, textvariable=self.weekly_shutdown_time_var,
                                     font=('Microsoft YaHei', 10), width=15)
        time_entry_weekly.pack(side=tk.LEFT, padx=5)
        tk.Button(weekly_frame, text="设置", command=lambda: self.show_power_week_time_dialog(
            "设置每周关机", self.weekly_shutdown_days_var, self.weekly_shutdown_time_var)).pack(side=tk.LEFT)

        reboot_frame = tk.Frame(power_frame, bg='white')
        reboot_frame.pack(fill=tk.X, pady=4)
        tk.Checkbutton(reboot_frame, text="每周重启", variable=self.weekly_reboot_enabled_var,
                       font=('Microsoft YaHei', 10), bg='white', command=self.save_settings).pack(side=tk.LEFT)
        days_entry_reboot = tk.Entry(reboot_frame, textvariable=self.weekly_reboot_days_var,
                                     font=('Microsoft YaHei', 10), width=20)
        days_entry_reboot.pack(side=tk.LEFT, padx=(10,5))
        time_entry_reboot = tk.Entry(reboot_frame, textvariable=self.weekly_reboot_time_var,
                                     font=('Microsoft YaHei', 10), width=15)
        time_entry_reboot.pack(side=tk.LEFT, padx=5)
        tk.Button(reboot_frame, text="设置", command=lambda: self.show_power_week_time_dialog(
            "设置每周重启", self.weekly_reboot_days_var, self.weekly_reboot_time_var)).pack(side=tk.LEFT)

        return settings_frame

    def toggle_lock_state(self):
        self.is_locked = not self.is_locked
        if self.is_locked:
            self.lock_button.config(text="解锁", bg='#2ECC71')
            self._set_ui_lock_state(tk.DISABLED)
            self.log("界面已锁定。")
        else:
            self.lock_button.config(text="锁定", bg='#E74C3C')
            self._set_ui_lock_state(tk.NORMAL)
            self.log("界面已解锁。")
        
        if "设置" in self.pages:
            self.lock_now_var.set(self.is_locked)

    def _set_ui_lock_state(self, state):
        self._set_widget_state_recursively(self.nav_frame, state)
        self._set_widget_state_recursively(self.top_right_btn_frame, state)
        self.clear_log_btn.config(state=state)

    def _set_widget_state_recursively(self, parent_widget, state):
        for child in parent_widget.winfo_children():
            if child == self.lock_button or child == self.nav_buttons.get("设置"):
                continue
            
            try:
                if child.master in [b.master for b in self.nav_buttons.values()] and child.master != self.nav_buttons.get("设置").master:
                     child.config(state=state)
                elif child.master not in [b.master for b in self.nav_buttons.values()]:
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
        context_menu = tk.Menu(self.root, tearoff=0, font=('Microsoft YaHei', 10))

        if iid:
            if iid not in self.task_tree.selection():
                self.task_tree.selection_set(iid)
            
            # 【对齐修正】: 使用统一尺寸的内嵌图标实现完美对齐
            context_menu.add_command(label="  立即播放", image=self.menu_icons['play'], compound=tk.LEFT, command=self.play_now)
            context_menu.add_separator()
            context_menu.add_command(label="  修改", image=self.menu_icons['edit'], compound=tk.LEFT, command=self.edit_task)
            context_menu.add_command(label="  删除", image=self.menu_icons['delete'], compound=tk.LEFT, command=self.delete_task)
            context_menu.add_command(label="  复制", image=self.menu_icons['copy'], compound=tk.LEFT, command=self.copy_task)
            context_menu.add_separator()
            context_menu.add_command(label="  置顶", image=self.menu_icons['top'], compound=tk.LEFT, command=self.move_task_to_top)
            context_menu.add_command(label="  上移", image=self.menu_icons['up'], compound=tk.LEFT, command=lambda: self.move_task(-1))
            context_menu.add_command(label="  下移", image=self.menu_icons['down'], compound=tk.LEFT, command=lambda: self.move_task(1))
            context_menu.add_command(label="  置末", image=self.menu_icons['bottom'], compound=tk.LEFT, command=self.move_task_to_bottom)
            context_menu.add_separator()
            context_menu.add_command(label="  启用", image=self.menu_icons['enable'], compound=tk.LEFT, command=self.enable_task)
            context_menu.add_command(label="  禁用", image=self.menu_icons['disable'], compound=tk.LEFT, command=self.disable_task)
        else:
            self.task_tree.selection_set()
            context_menu.add_command(label="  添加节目", image=self.menu_icons['add'], compound=tk.LEFT, command=self.add_task)
        
        context_menu.add_separator()
        context_menu.add_command(label="  停止当前播放", image=self.menu_icons['stop'], compound=tk.LEFT, command=self.stop_current_playback)
        
        context_menu.post(event.x_root, event.y_root)
    
    def _force_stop_playback(self):
        if self.is_playing.is_set():
            self.log("接收到中断指令，正在停止当前播放...")
            if AUDIO_AVAILABLE and pygame.mixer.music.get_busy():
                pygame.mixer.music.stop()
            self.on_playback_finished()
    
    def play_now(self):
        selection = self.task_tree.selection()
        if not selection: 
            messagebox.showwarning("提示", "请先选择一个要立即播放的节目。")
            return
        
        index = self.task_tree.index(selection[0])
        task = self.tasks[index]

        self.log(f"手动触发高优先级播放: {task['name']}")
        
        self._force_stop_playback()
        
        with self.queue_lock:
            self.playback_queue.clear()
            self.playback_queue.insert(0, (task, "manual_play"))
            self.log("播放队列已清空，新任务已置顶。")
        
        self.root.after(0, self._process_queue)

    def stop_current_playback(self):
        self.log("手动触发“停止当前播放”...")
        self._force_stop_playback()
        with self.queue_lock:
            if self.playback_queue:
                self.playback_queue.clear()
                self.log("等待播放的队列也已清空。")

    def add_task(self):
        choice_dialog = tk.Toplevel(self.root)
        choice_dialog.title("选择节目类型")
        choice_dialog.geometry("350x250")
        choice_dialog.resizable(False, False)
        choice_dialog.transient(self.root); choice_dialog.grab_set()
        self.center_window(choice_dialog, 350, 250)
        main_frame = tk.Frame(choice_dialog, padx=20, pady=20, bg='#F0F0F0')
        main_frame.pack(fill=tk.BOTH, expand=True)
        title_label = tk.Label(main_frame, text="请选择要添加的节目类型",
                              font=('Microsoft YaHei', 13, 'bold'), fg='#2C5F7C', bg='#F0F0F0')
        title_label.pack(pady=15)
        btn_frame = tk.Frame(main_frame, bg='#F0F0F0')
        btn_frame.pack(expand=True)
        audio_btn = tk.Button(btn_frame, text="🎵 音频节目", command=lambda: self.open_audio_dialog(choice_dialog),
                             bg='#5DADE2', fg='white', font=('Microsoft YaHei', 11, 'bold'),
                             bd=0, padx=30, pady=12, cursor='hand2', width=15)
        audio_btn.pack(pady=8)
        voice_btn = tk.Button(btn_frame, text="🎙️ 语音节目", command=lambda: self.open_voice_dialog(choice_dialog),
                             bg='#3498DB', fg='white', font=('Microsoft YaHei', 11, 'bold'),
                             bd=0, padx=30, pady=12, cursor='hand2', width=15)
        voice_btn.pack(pady=8)

    def open_audio_dialog(self, parent_dialog, task_to_edit=None, index=None):
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
        audio_single_entry = tk.Entry(audio_single_frame, font=('Microsoft YaHei', 10), width=35)
        audio_single_entry.pack(side=tk.LEFT, padx=5)
        tk.Label(audio_single_frame, text="00:00", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT, padx=10)
        def select_single_audio():
            filename = filedialog.askopenfilename(title="选择音频文件", initialdir=AUDIO_FOLDER,
                filetypes=[("音频文件", "*.mp3 *.wav *.ogg *.flac *.m4a"), ("所有文件", "*.*")])
            if filename:
                audio_single_entry.delete(0, tk.END)
                audio_single_entry.insert(0, filename)
        tk.Button(audio_single_frame, text="选取...", command=select_single_audio, bg='#D0D0D0',
                 font=('Microsoft YaHei', 10), bd=1, padx=15, pady=3).pack(side=tk.LEFT, padx=5)
        
        tk.Label(content_frame, text="音频文件夹", font=('Microsoft YaHei', 10), bg='#E8E8E8').grid(
            row=2, column=0, sticky='e', padx=5, pady=5)
        audio_folder_frame = tk.Frame(content_frame, bg='#E8E8E8')
        audio_folder_frame.grid(row=2, column=1, columnspan=3, sticky='ew', padx=5, pady=5)
        tk.Radiobutton(audio_folder_frame, text="", variable=audio_type_var, value="folder",
                      bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT)
        audio_folder_entry = tk.Entry(audio_folder_frame, font=('Microsoft YaHei', 10), width=50)
        audio_folder_entry.pack(side=tk.LEFT, padx=5)
        def select_folder():
            foldername = filedialog.askdirectory(title="选择音频文件夹", initialdir=AUDIO_FOLDER)
            if foldername:
                audio_folder_entry.delete(0, tk.END)
                audio_folder_entry.insert(0, foldername)
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
                audio_single_entry.insert(0, task.get('content', ''))
            else:
                audio_folder_entry.insert(0, task.get('content', ''))
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
        
        def save_task():
            audio_path = audio_single_entry.get().strip() if audio_type_var.get() == "single" else audio_folder_entry.get().strip()
            if not audio_path: messagebox.showwarning("警告", "请选择音频文件或文件夹", parent=dialog); return
            
            is_valid_time, time_msg = self._normalize_multiple_times_string(start_time_entry.get().strip())
            if not is_valid_time:
                messagebox.showwarning("格式错误", time_msg, parent=dialog)
                return
                
            is_valid_date, date_msg = self._normalize_date_range_string(date_range_entry.get().strip())
            if not is_valid_date:
                messagebox.showwarning("格式错误", date_msg, parent=dialog)
                return

            new_task_data = {'name': name_entry.get().strip(), 'time': time_msg, 'content': audio_path,
                             'type': 'audio', 'audio_type': audio_type_var.get(), 'play_order': play_order_var.get(),
                             'volume': volume_entry.get().strip() or "80", 'interval_type': interval_var.get(),
                             'interval_first': interval_first_entry.get().strip(), 'interval_seconds': interval_seconds_entry.get().strip(),
                             'weekday': weekday_entry.get().strip(), 'date_range': date_msg,
                             'delay': delay_var.get(), 
                             'status': '启用' if not is_edit_mode else task_to_edit.get('status', '启用'), 
                             'last_run': {} if not is_edit_mode else task_to_edit.get('last_run', {})}
            
            if not new_task_data['name'] or not new_task_data['time']:
                messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog); return
            
            if is_edit_mode:
                self.tasks[index] = new_task_data
                self.log(f"已修改音频节目: {new_task_data['name']}")
            else:
                self.tasks.append(new_task_data)
                self.log(f"已添加音频节目: {new_task_data['name']}")
                
            self.update_task_list(); self.save_tasks(); dialog.destroy()
        
        button_text = "保存修改" if is_edit_mode else "添加"
        tk.Button(button_frame, text=button_text, command=save_task, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 10, 'bold'), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 10), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        
        content_frame.columnconfigure(1, weight=1)
        time_frame.columnconfigure(1, weight=1)

    def open_voice_dialog(self, parent_dialog, task_to_edit=None, index=None):
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
        prompt_file_entry = tk.Entry(prompt_frame, textvariable=prompt_file_var, font=('Microsoft YaHei', 10), width=20)
        prompt_file_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(prompt_frame, text="...", command=lambda: self.select_file_for_entry(PROMPT_FOLDER, prompt_file_var)).pack(side=tk.LEFT)
        tk.Label(prompt_frame, text="音量(0-100):", font=('Microsoft YaHei', 10), bg='#E8E8E8').pack(side=tk.LEFT, padx=(10,5))
        tk.Entry(prompt_frame, textvariable=prompt_volume_var, font=('Microsoft YaHei', 10), width=8).pack(side=tk.LEFT, padx=5)

        bgm_var = tk.IntVar()
        bgm_frame = tk.Frame(content_frame, bg='#E8E8E8')
        bgm_frame.grid(row=5, column=1, columnspan=3, sticky='w', padx=5, pady=5)
        tk.Checkbutton(bgm_frame, text="背景音乐:", variable=bgm_var, bg='#E8E8E8', font=('Microsoft YaHei', 10)).pack(side=tk.LEFT)
        bgm_file_var, bgm_volume_var = tk.StringVar(), tk.StringVar()
        bgm_file_entry = tk.Entry(bgm_frame, textvariable=bgm_file_var, font=('Microsoft YaHei', 10), width=20)
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
        
        def save_task():
            content = content_text.get('1.0', tk.END).strip()
            if not content: messagebox.showwarning("警告", "请输入播音文字内容", parent=dialog); return
            
            is_valid_time, time_msg = self._normalize_multiple_times_string(start_time_entry.get().strip())
            if not is_valid_time:
                messagebox.showwarning("格式错误", time_msg, parent=dialog)
                return
                
            is_valid_date, date_msg = self._normalize_date_range_string(date_range_entry.get().strip())
            if not is_valid_date:
                messagebox.showwarning("格式错误", date_msg, parent=dialog)
                return

            new_task_data = {'name': name_entry.get().strip(), 'time': time_msg, 'content': content,
                             'type': 'voice', 'voice': voice_var.get(), 
                             'speed': speed_entry.get().strip() or "0",
                             'pitch': pitch_entry.get().strip() or "0",
                             'volume': volume_entry.get().strip() or "80",
                             'prompt': prompt_var.get(), 'prompt_file': prompt_file_var.get(),
                             'prompt_volume': prompt_volume_var.get(),
                             'bgm': bgm_var.get(), 'bgm_file': bgm_file_var.get(),
                             'bgm_volume': bgm_volume_var.get(),
                             'repeat': repeat_entry.get().strip() or "1",
                             'weekday': weekday_entry.get().strip(), 'date_range': date_msg,
                             'delay': delay_var.get(), 
                             'status': '启用' if not is_edit_mode else task_to_edit.get('status', '启用'), 
                             'last_run': {} if not is_edit_mode else task_to_edit.get('last_run', {})}
            
            if not new_task_data['name'] or not new_task_data['time']:
                messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间）", parent=dialog); return
            
            if is_edit_mode:
                self.tasks[index] = new_task_data
                self.log(f"已修改语音节目: {new_task_data['name']}")
            else:
                self.tasks.append(new_task_data)
                self.log(f"已添加语音节目: {new_task_data['name']}")
                
            self.update_task_list(); self.save_tasks(); dialog.destroy()
        
        button_text = "保存修改" if is_edit_mode else "添加"
        tk.Button(button_frame, text=button_text, command=save_task, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 10, 'bold'), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 10), bd=1, padx=40, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        
        content_frame.columnconfigure(1, weight=1)
        time_frame.columnconfigure(1, weight=1)

    def get_available_voices(self):
        available_voices = []
        if WIN32COM_AVAILABLE:
            try:
                pythoncom.CoInitialize()
                speaker = win32com.client.Dispatch("SAPI.SpVoice")
                voices = speaker.GetVoices()
                available_voices = [v.GetDescription() for v in voices]
                pythoncom.CoUninitialize()
            except Exception as e:
                self.log(f"警告: 使用 win32com 获取语音列表失败 - {e}")
                available_voices = []
        return available_voices
    
    def select_file_for_entry(self, initial_dir, string_var):
        filename = filedialog.askopenfilename(
            title="选择文件",
            initialdir=initial_dir,
            filetypes=[("音频文件", "*.mp3 *.wav *.ogg *.flac *.m4a"), ("所有文件", "*.*")]
        )
        if filename:
            string_var.set(os.path.basename(filename))

    def delete_task(self):
        selections = self.task_tree.selection()
        if not selections: messagebox.showwarning("警告", "请先选择要删除的节目"); return
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(selections)} 个节目吗？"):
            indices = sorted([self.task_tree.index(s) for s in selections], reverse=True)
            for index in indices: self.log(f"已删除节目: {self.tasks.pop(index)['name']}")
            self.update_task_list(); self.save_tasks()

    def edit_task(self):
        selection = self.task_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要修改的节目"); return
        if len(selection) > 1:
            messagebox.showwarning("警告", "一次只能修改一个节目"); return
        
        index = self.task_tree.index(selection[0])
        task = self.tasks[index]
        dummy_parent = tk.Toplevel(self.root)
        dummy_parent.withdraw()

        if task.get('type') == 'audio':
            self.open_audio_dialog(dummy_parent, task_to_edit=task, index=index)
        else:
            self.open_voice_dialog(dummy_parent, task_to_edit=task, index=index)
        
        def check_dialog_closed():
            try:
                if not dummy_parent.winfo_children():
                    dummy_parent.destroy()
                else:
                    self.root.after(100, check_dialog_closed)
            except tk.TclError:
                pass 
        self.root.after(100, check_dialog_closed)

    def copy_task(self):
        selections = self.task_tree.selection()
        if not selections: messagebox.showwarning("警告", "请先选择要复制的节目"); return
        for sel in selections:
            original = self.tasks[self.task_tree.index(sel)]
            copy = json.loads(json.dumps(original))
            copy['name'] += " (副本)"; copy['last_run'] = {}
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
            self.update_task_list()
            self.save_tasks()
            items = self.task_tree.get_children()
            if items: 
                self.task_tree.selection_set(items[new_index])
                self.task_tree.focus(items[new_index])

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
        filename = filedialog.askopenfilename(title="选择导入文件", filetypes=[("JSON文件", "*.json")])
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f: imported = json.load(f)
                self.tasks.extend(imported); self.update_task_list(); self.save_tasks()
                self.log(f"已从 {os.path.basename(filename)} 导入 {len(imported)} 个节目")
            except Exception as e: messagebox.showerror("错误", f"导入失败: {e}")

    def export_tasks(self):
        if not self.tasks: messagebox.showwarning("警告", "没有节目可以导出"); return
        filename = filedialog.asksaveasfilename(title="导出到...", defaultextension=".json",
            initialfile="broadcast_backup.json", filetypes=[("JSON文件", "*.json")])
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

    def show_time_settings_dialog(self, time_entry):
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
        
        current_times = time_entry.get().split(',')
        for t in [t.strip() for t in current_times if t.strip()]:
            listbox.insert(tk.END, t)

        btn_frame = tk.Frame(list_frame, bg='#D7F3F5')
        btn_frame.pack(side=tk.RIGHT, padx=10, fill=tk.Y)
        new_entry = tk.Entry(btn_frame, font=('Microsoft YaHei', 10), width=12)
        new_entry.insert(0, datetime.now().strftime("%H:%M:%S")); new_entry.pack(pady=3)
        
        def add_time():
            val = new_entry.get().strip()
            normalized_time = self._normalize_time_string(val)
            if normalized_time:
                if normalized_time not in listbox.get(0, tk.END):
                    listbox.insert(tk.END, normalized_time)
                    new_entry.delete(0, tk.END)
                    new_entry.insert(0, datetime.now().strftime("%H:%M:%S"))
            else:
                messagebox.showerror("格式错误", "请输入有效的时间格式 HH:MM:SS", parent=dialog)

        def del_time():
            if listbox.curselection():
                listbox.delete(listbox.curselection()[0])
        tk.Button(btn_frame, text="添加 ↑", command=add_time).pack(pady=3, fill=tk.X)
        tk.Button(btn_frame, text="删除", command=del_time).pack(pady=3, fill=tk.X)
        tk.Button(btn_frame, text="清空", command=lambda: listbox.delete(0, tk.END)).pack(pady=3, fill=tk.X)
        bottom_frame = tk.Frame(main_frame, bg='#D7F3F5')
        bottom_frame.pack(pady=10)
        def confirm():
            result = ", ".join(list(listbox.get(0, tk.END)))
            if isinstance(time_entry, tk.Entry):
                time_entry.delete(0, tk.END)
                time_entry.insert(0, result)
            self.save_settings()
            dialog.destroy()
        tk.Button(bottom_frame, text="确定", command=confirm, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 9, 'bold'), bd=1, padx=25, pady=5).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 9), bd=1, padx=25, pady=5).pack(side=tk.LEFT, padx=5)

    def show_weekday_settings_dialog(self, weekday_var):
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
        
        current_val = weekday_var.get()
        if current_val.startswith("每周:"):
            week_type_var.set("week")
            selected_days = current_val.replace("每周:", "")
            for day_num in week_vars:
                week_vars[day_num].set(1 if str(day_num) in selected_days else 0)
        elif current_val.startswith("每月:"):
            week_type_var.set("day")
            selected_days = current_val.replace("每月:", "").split(',')
            for day_num in day_vars:
                 day_vars[day_num].set(1 if f"{day_num:02d}" in selected_days else 0)

        def confirm():
            if week_type_var.get() == "week":
                selected = sorted([str(n) for n, v in week_vars.items() if v.get()])
                result = "每周:" + "".join(selected)
            else:
                selected = sorted([f"{n:02d}" for n, v in day_vars.items() if v.get()])
                result = "每月:" + ",".join(selected)
            
            if isinstance(weekday_var, tk.Entry):
                weekday_var.delete(0, tk.END)
                weekday_var.insert(0, result if selected else "")
            self.save_settings()
            dialog.destroy()
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
            start, end = from_date_entry.get().strip(), to_date_entry.get().strip()
            norm_start, norm_end = self._normalize_date_string(start), self._normalize_date_string(end)
            
            if norm_start and norm_end:
                date_range_entry.delete(0, tk.END)
                date_range_entry.insert(0, f"{norm_start} ~ {norm_end}")
                dialog.destroy()
            else:
                messagebox.showerror("格式错误", "日期格式不正确, 应为 YYYY-MM-DD", parent=dialog)

        tk.Button(bottom_frame, text="确定", command=confirm, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 9, 'bold'), bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0',
                 font=('Microsoft YaHei', 9), bd=1, padx=30, pady=6).pack(side=tk.LEFT, padx=5)

    def show_single_time_dialog(self, time_var):
        dialog = tk.Toplevel(self.root)
        dialog.title("设置时间"); dialog.geometry("300x180"); dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set(); dialog.configure(bg='#D7F3F5')
        self.center_window(dialog, 300, 180)
        main_frame = tk.Frame(dialog, bg='#D7F3F5', padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        tk.Label(main_frame, text="24小时制 HH:MM:SS", font=('Microsoft YaHei', 10, 'bold'),
                bg='#D7F3F5').pack(pady=5)
        time_entry = tk.Entry(main_frame, font=('Microsoft YaHei', 12), width=15, justify='center')
        time_entry.insert(0, time_var.get())
        time_entry.pack(pady=10)

        def confirm():
            val = time_entry.get().strip()
            normalized_time = self._normalize_time_string(val)
            if normalized_time:
                time_var.set(normalized_time)
                self.save_settings()
                dialog.destroy()
            else:
                messagebox.showerror("格式错误", "请输入有效的时间格式 HH:MM:SS", parent=dialog)

        bottom_frame = tk.Frame(main_frame, bg='#D7F3F5')
        bottom_frame.pack(pady=10)
        tk.Button(bottom_frame, text="确定", command=confirm, bg='#5DADE2', fg='white').pack(side=tk.LEFT, padx=10)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0').pack(side=tk.LEFT, padx=10)

    def show_power_week_time_dialog(self, title, days_var, time_var):
        # 【修改】增加窗口宽度
        dialog = tk.Toplevel(self.root); dialog.title(title)
        dialog.geometry("550x300"); dialog.resizable(False, False)
        dialog.transient(self.root); dialog.grab_set(); dialog.configure(bg='#D7F3F5')
        self.center_window(dialog, 550, 300)
        
        week_frame = tk.LabelFrame(dialog, text="选择周几", font=('Microsoft YaHei', 10, 'bold'),
                                  bg='#D7F3F5', padx=10, pady=10)
        week_frame.pack(fill=tk.X, pady=10, padx=10)
        
        weekdays = [("周一", 1), ("周二", 2), ("周三", 3), ("周四", 4), ("周五", 5), ("周六", 6), ("周日", 7)]
        week_vars = {num: tk.IntVar() for day, num in weekdays}
        
        current_days = days_var.get().replace("每周:", "")
        for day_num_str in current_days:
            week_vars[int(day_num_str)].set(1)

        for i, (day, num) in enumerate(weekdays):
            tk.Checkbutton(week_frame, text=day, variable=week_vars[num], bg='#D7F3F5',
                          font=('Microsoft YaHei', 10)).grid(row=0, column=i, sticky='w', padx=10, pady=3)

        time_frame = tk.LabelFrame(dialog, text="设置时间", font=('Microsoft YaHei', 10, 'bold'),
                                  bg='#D7F3F5', padx=10, pady=10)
        time_frame.pack(fill=tk.X, pady=10, padx=10)
        tk.Label(time_frame, text="时间 (HH:MM:SS):", font=('Microsoft YaHei', 10), bg='#D7F3F5').pack(side=tk.LEFT)
        time_entry = tk.Entry(time_frame, font=('Microsoft YaHei', 10), width=15)
        time_entry.insert(0, time_var.get())
        time_entry.pack(side=tk.LEFT, padx=10)

        def confirm():
            selected_days = sorted([str(n) for n, v in week_vars.items() if v.get()])
            if not selected_days:
                messagebox.showwarning("提示", "请至少选择一天", parent=dialog)
                return
            
            normalized_time = self._normalize_time_string(time_entry.get().strip())
            if not normalized_time:
                messagebox.showerror("格式错误", "请输入有效的时间格式 HH:MM:SS", parent=dialog)
                return

            days_var.set("每周:" + "".join(selected_days))
            time_var.set(normalized_time)
            self.save_settings()
            dialog.destroy()
            
        bottom_frame = tk.Frame(dialog, bg='#D7F3F5')
        bottom_frame.pack(pady=15)
        tk.Button(bottom_frame, text="确定", command=confirm, bg='#5DADE2', fg='white').pack(side=tk.LEFT, padx=10)
        tk.Button(bottom_frame, text="取消", command=dialog.destroy, bg='#D0D0D0').pack(side=tk.LEFT, padx=10)

    def update_task_list(self):
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
            try: 
                valid_selection = [s for s in selection if self.task_tree.exists(s)]
                if valid_selection:
                    self.task_tree.selection_set(valid_selection)
            except tk.TclError: pass
        self.stats_label.config(text=f"节目单：{len(self.tasks)}")
        if hasattr(self, 'status_labels'): self.status_labels[3].config(text=f"任务数量: {len(self.tasks)}")

    def update_status_bar(self):
        if not self.running: return
        self.status_labels[0].config(text=f"当前时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.status_labels[1].config(text="系统状态: 运行中")
        self.root.after(1000, self.update_status_bar)

    def start_background_thread(self):
        threading.Thread(target=self._background_worker, daemon=True).start()

    def _background_worker(self):
        while self.running:
            now = datetime.now()
            self._check_broadcast_tasks(now)
            self._check_power_tasks(now)
            time.sleep(1)

    def _check_broadcast_tasks(self, now):
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
                        self._force_stop_playback()
                        with self.queue_lock:
                            self.playback_queue.clear()
                            self.playback_queue.insert(0, (task, trigger_time))
                        self.root.after(0, self._process_queue)
                    else:
                        with self.queue_lock:
                            self.playback_queue.append((task, trigger_time))
                        self.log(f"延时任务 '{task['name']}' 已到时间，加入播放队列。")
                        self.root.after(0, self._process_queue)

    def _check_power_tasks(self, now):
        current_date_str = now.strftime("%Y-%m-%d")
        current_time_str = now.strftime("%H:%M:%S")

        if self.settings.get("last_power_action_date") == current_date_str:
            return

        action_to_take = None
        
        if self.settings.get("daily_shutdown_enabled") and current_time_str == self.settings.get("daily_shutdown_time"):
            action_to_take = ("shutdown /s /t 60", "每日定时关机")
        if not action_to_take and self.settings.get("weekly_shutdown_enabled"):
            days = self.settings.get("weekly_shutdown_days", "").replace("每周:", "")
            if str(now.isoweekday()) in days and current_time_str == self.settings.get("weekly_shutdown_time"):
                action_to_take = ("shutdown /s /t 60", "每周定时关机")
        if not action_to_take and self.settings.get("weekly_reboot_enabled"):
            days = self.settings.get("weekly_reboot_days", "").replace("每周:", "")
            if str(now.isoweekday()) in days and current_time_str == self.settings.get("weekly_reboot_time"):
                action_to_take = ("shutdown /r /t 60", "每周定时重启")

        if action_to_take:
            command, reason = action_to_take
            self.log(f"执行系统电源任务: {reason}。系统将在60秒后操作。")
            self.settings["last_power_action_date"] = current_date_str
            self.save_settings()
            os.system(command)

    def _process_queue(self):
        if self.is_playing.is_set():
            return

        with self.queue_lock:
            if not self.playback_queue:
                return
            task, trigger_time = self.playback_queue.pop(0)
        
        self._execute_broadcast(task, trigger_time)

    def _execute_broadcast(self, task, trigger_time):
        self.is_playing.set()
        self.update_playing_text(f"[{task['name']}] 正在准备播放...")
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
                if not self.is_playing.is_set(): break
                self.log(f"正在播放: {os.path.basename(audio_path)}")
                self.update_playing_text(f"[{task['name']}] 正在播放: {os.path.basename(audio_path)}")
                
                pygame.mixer.music.load(audio_path)
                pygame.mixer.music.set_volume(float(task.get('volume', 80)) / 100.0)
                pygame.mixer.music.play()

                while pygame.mixer.music.get_busy() and self.is_playing.is_set():
                    if interval_type == 'seconds' and (time.time() - start_time) > duration_seconds:
                        pygame.mixer.music.stop()
                        self.log(f"已达到 {duration_seconds} 秒播放时长限制。")
                        break
                    time.sleep(0.1)
                
                if interval_type == 'seconds' and (time.time() - start_time) > duration_seconds:
                    break
        except Exception as e:
            self.log(f"音频播放错误: {e}")
        finally:
            self.root.after(0, self.on_playback_finished)

    def _speak(self, text, task):
        if not WIN32COM_AVAILABLE:
            self.log("错误: pywin32库不可用，无法执行语音播报。")
            self.root.after(0, self.on_playback_finished)
            return
        
        pythoncom.CoInitialize()
        try:
            if not self.is_playing.is_set(): return

            if task.get('bgm', 0) and AUDIO_AVAILABLE:
                bgm_file, bgm_path = task.get('bgm_file', ''), os.path.join(BGM_FOLDER, task.get('bgm_file', ''))
                if os.path.exists(bgm_path):
                    self.log(f"播放背景音乐: {bgm_file}")
                    pygame.mixer.music.load(bgm_path)
                    pygame.mixer.music.set_volume(float(task.get('bgm_volume', 40)) / 100.0)
                    pygame.mixer.music.play(-1)
                else:
                    self.log(f"警告: 背景音乐文件不存在 - {bgm_path}")

            if task.get('prompt', 0) and AUDIO_AVAILABLE:
                prompt_file, prompt_path = task.get('prompt_file', ''), os.path.join(PROMPT_FOLDER, task.get('prompt_file', ''))
                if os.path.exists(prompt_path):
                    if not self.is_playing.is_set(): return
                    self.log(f"播放提示音: {prompt_file}")
                    sound = pygame.mixer.Sound(prompt_path)
                    sound.set_volume(float(task.get('prompt_volume', 80)) / 100.0)
                    channel = sound.play()
                    if channel:
                        while channel.get_busy() and self.is_playing.is_set():
                            time.sleep(0.05)
                else:
                    self.log(f"警告: 提示音文件不存在 - {prompt_path}")
            
            if not self.is_playing.is_set(): return
            
            try:
                speaker = win32com.client.Dispatch("SAPI.SpVoice")
            except com_error as e:
                self.log(f"严重错误: 无法初始化语音引擎! 错误: {e}"); raise

            all_voices = {v.GetDescription(): v for v in speaker.GetVoices()}
            if (selected_voice_desc := task.get('voice')) in all_voices:
                speaker.Voice = all_voices[selected_voice_desc]
            
            speaker.Volume = int(task.get('volume', 80))
            escaped_text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("'", "&apos;").replace('"', "&quot;")
            xml_text = f"<rate absspeed='{task.get('speed', '0')}'><pitch middle='{task.get('pitch', '0')}'>{escaped_text}</pitch></rate>"
            
            repeat_count = int(task.get('repeat', 1))
            self.log(f"准备播报 {repeat_count} 遍...")

            for i in range(repeat_count):
                if not self.is_playing.is_set(): break
                self.log(f"正在播报第 {i+1}/{repeat_count} 遍")
                speaker.Speak(xml_text, 8) 
                if i < repeat_count - 1 and self.is_playing.is_set():
                    time.sleep(0.5)

        except Exception as e:
            self.log(f"播报错误: {e}")
        finally:
            if AUDIO_AVAILABLE and pygame.mixer.music.get_busy():
                pygame.mixer.music.stop()
                self.log("背景音乐已停止。")
            pythoncom.CoUninitialize()
            self.root.after(0, self.on_playback_finished)

    def on_playback_finished(self):
        if self.is_playing.is_set():
            self.is_playing.clear()
            self.update_playing_text("等待下一个任务...")
            self.status_labels[2].config(text="播放状态: 待机")
            self.log("播放结束")
            self.root.after(100, self._process_queue)

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
                if 'delay' not in task:
                    task['delay'] = 'delay' if task.get('type') == 'voice' else 'ontime'
                    migrated = True
                if not isinstance(task.get('last_run'), dict):
                    task['last_run'] = {}
                    migrated = True
            if migrated:
                self.log("旧版任务数据已迁移。")
                self.save_tasks()
            self.update_task_list(); self.log(f"已加载 {len(self.tasks)} 个节目")
        except Exception as e: self.log(f"加载任务失败: {e}")

    def load_settings(self):
        defaults = {
            "autostart": False, "start_minimized": False,
            "daily_shutdown_enabled": False, "daily_shutdown_time": "23:00:00",
            "weekly_shutdown_enabled": False, "weekly_shutdown_days": "每周:12345", "weekly_shutdown_time": "23:30:00",
            "weekly_reboot_enabled": False, "weekly_reboot_days": "每周:67", "weekly_reboot_time": "22:00:00",
            "last_power_action_date": ""
        }
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                    self.settings = json.load(f)
                for key, value in defaults.items():
                    self.settings.setdefault(key, value)
            except Exception as e:
                self.log(f"加载设置失败: {e}, 将使用默认设置。")
                self.settings = defaults
        else:
            self.settings = defaults
        self.log("系统设置已加载。")

    def save_settings(self):
        if hasattr(self, 'autostart_var'):
            self.settings.update({
                "autostart": self.autostart_var.get(), "start_minimized": self.start_minimized_var.get(),
                "daily_shutdown_enabled": self.daily_shutdown_enabled_var.get(), "daily_shutdown_time": self.daily_shutdown_time_var.get(),
                "weekly_shutdown_enabled": self.weekly_shutdown_enabled_var.get(), "weekly_shutdown_days": self.weekly_shutdown_days_var.get(),
                "weekly_shutdown_time": self.weekly_shutdown_time_var.get(), "weekly_reboot_enabled": self.weekly_reboot_enabled_var.get(),
                "weekly_reboot_days": self.weekly_reboot_days_var.get(), "weekly_reboot_time": self.weekly_reboot_time_var.get()
            })
        try:
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"保存设置失败: {e}")

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
                if os.path.exists(shortcut_path):
                    os.remove(shortcut_path)
                    self.log("已取消开机自动启动。")
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
        try:
            return datetime.strptime(date_str, "%Y-%m-%d").strftime("%Y-%m-%d")
        except ValueError: return None
            
    def _normalize_date_range_string(self, date_range_input_str):
        if not date_range_input_str.strip(): return True, ""
        try:
            start_str, end_str = [d.strip() for d in date_range_input_str.split('~')]
            norm_start, norm_end = self._normalize_date_string(start_str), self._normalize_date_string(end_str)
            if norm_start and norm_end: return True, f"{norm_start} ~ {norm_end}"
            invalid_parts = [p for p, n in [(start_str, norm_start), (end_str, norm_end)] if not n]
            return False, f"以下日期格式无效 (应为 YYYY-MM-DD): {', '.join(invalid_parts)}"
        except (ValueError, IndexError):
            return False, "日期范围格式无效，应为 'YYYY-MM-DD ~ YYYY-MM-DD'"

    def show_quit_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("确认")
        dialog.geometry("350x150")
        dialog.resizable(False, False); dialog.transient(self.root); dialog.grab_set()
        self.center_window(dialog, 350, 150)
        
        tk.Label(dialog, text="您想要如何操作？", font=('Microsoft YaHei', 12), pady=20).pack()
        btn_frame = tk.Frame(dialog); btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="退出程序", command=lambda: [dialog.destroy(), self.quit_app()]).pack(side=tk.LEFT, padx=10)
        if TRAY_AVAILABLE:
            tk.Button(btn_frame, text="最小化到托盘", command=lambda: [dialog.destroy(), self.hide_to_tray()]).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=10)

    def hide_to_tray(self):
        if not TRAY_AVAILABLE:
            messagebox.showwarning("功能不可用", "pystray 或 Pillow 库未安装，无法最小化到托盘。")
            return
        self.root.withdraw()
        self.log("程序已最小化到系统托盘。")

    def show_from_tray(self, icon, item):
        self.root.after(0, self.root.deiconify)
        self.log("程序已从托盘恢复。")

    def quit_app(self, icon=None, item=None):
        if self.tray_icon: self.tray_icon.stop()
        self.running = False
        self.save_tasks()
        self.save_settings()
        if AUDIO_AVAILABLE and pygame.mixer.get_init(): pygame.mixer.quit()
        self.root.destroy()
        sys.exit()

    def setup_tray_icon(self):
        try:
            image = Image.open(ICON_FILE)
        except Exception as e:
            image = Image.new('RGB', (64, 64), 'white')
            print(f"警告: 未找到或无法加载图标文件 '{ICON_FILE}': {e}")
        
        menu = (item('显示', self.show_from_tray, default=True), item('退出', self.quit_app))
        self.tray_icon = Icon("boyin", image, "定时播音", menu)

    def start_tray_icon_thread(self):
        if TRAY_AVAILABLE and self.tray_icon is None:
            self.setup_tray_icon()
            threading.Thread(target=self.tray_icon.run, daemon=True).start()
            self.log("系统托盘图标已启动。")

def main():
    root = tk.Tk()
    app = TimedBroadcastApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
