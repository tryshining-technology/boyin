
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import pyttsx3
import json
import threading
import time
from datetime import datetime
import os

class TimedBroadcastApp:
    def __init__(self, root):
        self.root = root
        self.root.title("定时播音")
        self.root.geometry("1400x800")
        self.root.configure(bg='#E8F4F8')
        
        # 初始化语音引擎
        self.engine = pyttsx3.init()
        self.engine.setProperty('rate', 150)
        self.engine.setProperty('volume', 1.0)
        
        # 任务列表
        self.tasks = []
        self.running = False
        self.task_file = "broadcast_tasks.json"
        self.current_page = "定时广播"
        
        # 创建界面
        self.create_widgets()
        
        # 加载已保存的任务
        self.load_tasks()
        
        # 启动后台检查线程
        self.start_background_thread()
    
    def create_widgets(self):
        # 左侧导航栏
        self.nav_frame = tk.Frame(self.root, bg='#A8D8E8', width=160)
        self.nav_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.nav_frame.pack_propagate(False)
        
        # 导航按钮
        nav_buttons = [
            ("定时广播", "频道"),
            ("背景音乐", "频道"),
            ("立即播播", "频道"),
            ("节假日、调休", "节假日不播或、调休"),
            ("设置", "开关机、闭区、多路\n作息表切换、音响..."),
            ("语音广告 制作", "")
        ]
        
        for i, (title, subtitle) in enumerate(nav_buttons):
            btn_frame = tk.Frame(self.nav_frame, bg='#5DADE2' if i == 0 else '#A8D8E8')
            btn_frame.pack(fill=tk.X, pady=1)
            
            btn = tk.Button(btn_frame, text=title, bg='#5DADE2' if i == 0 else '#A8D8E8',
                          fg='white' if i == 0 else 'black', font=('Microsoft YaHei', 11, 'bold'),
                          bd=0, padx=10, pady=8, anchor='w',
                          command=lambda t=title: self.switch_page(t))
            btn.pack(fill=tk.X)
            
            if subtitle:
                sub_label = tk.Label(btn_frame, text=subtitle, bg='#5DADE2' if i == 0 else '#A8D8E8',
                                   fg='#FF6B35' if i == 3 else ('#555' if i == 0 else '#666'),
                                   font=('Microsoft YaHei', 8), anchor='w', padx=10)
                sub_label.pack(fill=tk.X)
        
        # 主内容区域
        self.main_frame = tk.Frame(self.root, bg='white')
        self.main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建定时广播页面
        self.create_scheduled_broadcast_page()
    
    def switch_page(self, page_name):
        """切换页面"""
        self.current_page = page_name
        # 这里可以扩展其他页面
        if page_name == "定时广播":
            self.log(f"切换到: {page_name}")
        else:
            self.log(f"功能开发中: {page_name}")
    
    def create_scheduled_broadcast_page(self):
        """创建定时广播页面"""
        # 顶部标题和控制区
        top_frame = tk.Frame(self.main_frame, bg='white')
        top_frame.pack(fill=tk.X, padx=10, pady=10)
        
        title_label = tk.Label(top_frame, text="定时广播", font=('Microsoft YaHei', 14, 'bold'),
                              bg='white', fg='#2C5F7C')
        title_label.pack(side=tk.LEFT)
        
        freq_label = tk.Label(top_frame, text="频道", font=('Microsoft YaHei', 10),
                            bg='white', fg='#666')
        freq_label.pack(side=tk.LEFT, padx=10)
        
        # 控制按钮区
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
            ("导出", self.export_tasks, '#1ABC9C')
        ]
        
        for text, cmd, color in buttons:
            btn = tk.Button(btn_frame, text=text, command=cmd, bg=color, fg='white',
                          font=('Microsoft YaHei', 9), bd=0, padx=12, pady=5, cursor='hand2')
            btn.pack(side=tk.LEFT, padx=3)
        
        # 节目单统计
        stats_frame = tk.Frame(self.main_frame, bg='#F0F8FF')
        stats_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.stats_label = tk.Label(stats_frame, text="节目单：0", font=('Microsoft YaHei', 10),
                                   bg='#F0F8FF', fg='#2C5F7C', anchor='w', padx=10)
        self.stats_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 节目列表表格
        table_frame = tk.Frame(self.main_frame, bg='white')
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 创建表格
        columns = ('节目名称', '状态', '开始时间(可多个)', '延时秒', '音频或文字', '音量', '周几/几号', '日期范围')
        self.task_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=12)
        
        # 设置列宽
        col_widths = [200, 60, 140, 70, 300, 60, 100, 120]
        for col, width in zip(columns, col_widths):
            self.task_tree.heading(col, text=col)
            self.task_tree.column(col, width=width, anchor='w' if col == '节目名称' or col == '音频或文字' else 'center')
        
        self.task_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.task_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.task_tree.configure(yscrollcommand=scrollbar.set)
        
        # 正在播放区域
        playing_frame = tk.LabelFrame(self.main_frame, text="正在播：", font=('Microsoft YaHei', 10),
                                     bg='white', fg='#2C5F7C', padx=10, pady=5)
        playing_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.playing_text = scrolledtext.ScrolledText(playing_frame, height=3, font=('Microsoft YaHei', 9),
                                                     bg='#FFFEF0', wrap=tk.WORD)
        self.playing_text.pack(fill=tk.BOTH, expand=True)
        self.playing_text.insert('1.0', "等待播放...")
        
        # 日志区域
        log_frame = tk.LabelFrame(self.main_frame, text="日志：", font=('Microsoft YaHei', 10),
                                 bg='white', fg='#2C5F7C', padx=10, pady=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, font=('Microsoft YaHei', 9),
                                                 bg='#F9F9F9', wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 底部状态栏
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
        
        # 启动状态栏更新
        self.update_status_bar()
        self.log("定时播音软件已启动")
    
    def add_task(self):
        """添加任务对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("添加节目")
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 节目信息输入
        fields_frame = tk.Frame(dialog, padx=20, pady=20)
        fields_frame.pack(fill=tk.BOTH, expand=True)
        
        fields = [
            ("节目名称:", "name"),
            ("开始时间:", "time"),
            ("延时秒:", "delay"),
            ("音量(0-100):", "volume"),
            ("周几/几号:", "weekday"),
            ("日期范围:", "date_range")
        ]
        
        entries = {}
        for i, (label_text, key) in enumerate(fields):
            tk.Label(fields_frame, text=label_text, font=('Microsoft YaHei', 10)).grid(row=i, column=0, sticky='w', pady=8)
            entry = tk.Entry(fields_frame, font=('Microsoft YaHei', 10), width=40)
            entry.grid(row=i, column=1, pady=8, padx=10)
            entries[key] = entry
        
        # 默认值
        entries['delay'].insert(0, "0")
        entries['volume'].insert(0, "100")
        entries['weekday'].insert(0, "1-7")
        entries['date_range'].insert(0, "全年")
        
        # 播放内容
        tk.Label(fields_frame, text="播放内容:", font=('Microsoft YaHei', 10)).grid(row=len(fields), column=0, sticky='nw', pady=8)
        content_text = scrolledtext.ScrolledText(fields_frame, height=6, font=('Microsoft YaHei', 10), width=40)
        content_text.grid(row=len(fields), column=1, pady=8, padx=10)
        
        # 状态选择
        tk.Label(fields_frame, text="状态:", font=('Microsoft YaHei', 10)).grid(row=len(fields)+1, column=0, sticky='w', pady=8)
        status_var = tk.StringVar(value="启用")
        status_combo = ttk.Combobox(fields_frame, textvariable=status_var, values=["启用", "禁用"], 
                                   font=('Microsoft YaHei', 10), width=37, state='readonly')
        status_combo.grid(row=len(fields)+1, column=1, pady=8, padx=10)
        
        def save_task():
            task = {
                'name': entries['name'].get().strip(),
                'time': entries['time'].get().strip(),
                'delay': entries['delay'].get().strip() or "0",
                'content': content_text.get('1.0', tk.END).strip(),
                'volume': entries['volume'].get().strip() or "100",
                'weekday': entries['weekday'].get().strip() or "1-7",
                'date_range': entries['date_range'].get().strip() or "全年",
                'status': status_var.get(),
                'last_run': None
            }
            
            if not task['name'] or not task['time'] or not task['content']:
                messagebox.showwarning("警告", "请填写必要信息（节目名称、开始时间、播放内容）")
                return
            
            self.tasks.append(task)
            self.update_task_list()
            self.save_tasks()
            self.log(f"已添加节目: {task['name']} - {task['time']}")
            dialog.destroy()
        
        # 按钮
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="保存", command=save_task, bg='#5DADE2', fg='white',
                 font=('Microsoft YaHei', 10), bd=0, padx=30, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="取消", command=dialog.destroy, bg='#95A5A6', fg='white',
                 font=('Microsoft YaHei', 10), bd=0, padx=30, pady=8, cursor='hand2').pack(side=tk.LEFT, padx=10)
    
    def delete_task(self):
        """删除任务"""
        selection = self.task_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要删除的节目")
            return
        
        if messagebox.askyesno("确认", "确定要删除选中的节目吗？"):
            index = self.task_tree.index(selection[0])
            task = self.tasks[index]
            self.tasks.pop(index)
            self.update_task_list()
            self.save_tasks()
            self.log(f"已删除节目: {task['name']}")
    
    def edit_task(self):
        """编辑任务"""
        self.log("编辑功能：选择节目后可进行修改")
    
    def copy_task(self):
        """复制任务"""
        selection = self.task_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要复制的节目")
            return
        
        index = self.task_tree.index(selection[0])
        task = self.tasks[index].copy()
        task['name'] = task['name'] + " (副本)"
        self.tasks.append(task)
        self.update_task_list()
        self.save_tasks()
        self.log(f"已复制节目: {task['name']}")
    
    def move_task(self, direction):
        """移动任务"""
        selection = self.task_tree.selection()
        if not selection:
            return
        
        index = self.task_tree.index(selection[0])
        new_index = index + direction
        
        if 0 <= new_index < len(self.tasks):
            self.tasks[index], self.tasks[new_index] = self.tasks[new_index], self.tasks[index]
            self.update_task_list()
            self.save_tasks()
            # 重新选中
            items = self.task_tree.get_children()
            self.task_tree.selection_set(items[new_index])
    
    def import_tasks(self):
        """导入任务"""
        filename = filedialog.askopenfilename(filetypes=[("JSON文件", "*.json")])
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    imported = json.load(f)
                self.tasks.extend(imported)
                self.update_task_list()
                self.save_tasks()
                self.log(f"已导入 {len(imported)} 个节目")
            except Exception as e:
                messagebox.showerror("错误", f"导入失败: {str(e)}")
    
    def export_tasks(self):
        """导出任务"""
        filename = filedialog.asksaveasfilename(defaultextension=".json",
                                               filetypes=[("JSON文件", "*.json")])
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(self.tasks, f, ensure_ascii=False, indent=2)
                self.log(f"已导出 {len(self.tasks)} 个节目")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}")
    
    def update_task_list(self):
        """更新任务列表"""
        for item in self.task_tree.get_children():
            self.task_tree.delete(item)
        
        for task in self.tasks:
            content_preview = task['content'][:30] + '...' if len(task['content']) > 30 else task['content']
            self.task_tree.insert('', tk.END, values=(
                task['name'],
                task['status'],
                task['time'],
                task['delay'],
                content_preview,
                task['volume'],
                task['weekday'],
                task['date_range']
            ))
        
        # 更新统计
        self.stats_label.config(text=f"节目单：{len(self.tasks)}")
    
    def update_status_bar(self):
        """更新状态栏"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.status_labels[0].config(text=f"当前时间: {current_time}")
        self.status_labels[1].config(text="系统状态: 运行中")
        self.status_labels[2].config(text="播放状态: 待机")
        self.status_labels[3].config(text=f"任务数量: {len(self.tasks)}")
        
        self.root.after(1000, self.update_status_bar)
    
    def start_background_thread(self):
        """启动后台线程"""
        self.running = True
        thread = threading.Thread(target=self._check_tasks, daemon=True)
        thread.start()
    
    def _check_tasks(self):
        """后台检查任务"""
        while self.running:
            now = datetime.now()
            current_time = now.strftime("%H:%M")
            current_date = now.strftime("%Y-%m-%d")
            weekday = now.isoweekday()  # 1=周一, 7=周日
            
            for task in self.tasks:
                if task['status'] != '启用':
                    continue
                
                # 检查时间匹配
                if current_time in task['time']:
                    if task['last_run'] == current_date:
                        continue
                    
                    # 检查星期
                    if str(weekday) in task['weekday'] or '1-7' in task['weekday']:
                        self.root.after(0, self._execute_broadcast, task, current_date)
            
            time.sleep(30)
    
    def _execute_broadcast(self, task, current_date):
        """执行播报"""
        self.playing_text.delete('1.0', tk.END)
        self.playing_text.insert('1.0', f"[{task['name']}] {task['content']}")
        self.log(f"开始播报: {task['name']}")
        
        threading.Thread(target=self._speak, args=(task['content'],), daemon=True).start()
        task['last_run'] = current_date
        self.save_tasks()
    
    def _speak(self, text):
        """语音播报"""
        try:
            self.engine.say(text)
            self.engine.runAndWait()
        except Exception as e:
            self.log(f"播报错误: {str(e)}")
    
    def log(self, message):
        """记录日志"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"{timestamp}    {message}\n"
        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)
    
    def save_tasks(self):
        """保存任务"""
        try:
            with open(self.task_file, 'w', encoding='utf-8') as f:
                json.dump(self.tasks, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"保存失败: {str(e)}")
    
    def load_tasks(self):
        """加载任务"""
        if os.path.exists(self.task_file):
            try:
                with open(self.task_file, 'r', encoding='utf-8') as f:
                    self.tasks = json.load(f)
                self.update_task_list()
                self.log(f"已加载 {len(self.tasks)} 个节目")
            except Exception as e:
                self.log(f"加载失败: {str(e)}")

def main():
    root = tk.Tk()
    app = TimedBroadcastApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
