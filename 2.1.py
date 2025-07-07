import os
import sys
import json
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import datetime
import winreg
import platform
import ctypes
from PIL import Image, ImageDraw, ImageFont
import pystray
import keyboard
import logging
from pathlib import Path

# 常量定义
APP_NAME = "懒人关机器"

# 配置文件路径
if platform.system() == "Windows":
    CONFIG_DIR = Path(os.getenv('APPDATA')) / "LazyShutdown"
else:
    CONFIG_DIR = Path.home() / ".config" / "LazyShutdown"

# 确保目录存在
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = CONFIG_DIR / "lazy_shutdown_config.json"

AUTO_START_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"

# 关机类型映射
SHUTDOWN_TYPES = {
    "关机": "shutdown /s /t 0",
    "重启": "shutdown /r /t 0",
    "注销": "shutdown /l",
    "睡眠": "rundll32.exe powrprof.dll,SetSuspendState 0,1,0",
    "休眠": "shutdown /h"
}

# 默认配置
DEFAULT_CONFIG = {
    "auto_start": False,
    "minimize_to_tray": True,
    "hide_tray_icon": False,
    "hotkey": "ctrl+alt+l",
    "schedules": [],
    "run_as_admin": True  # 添加管理员运行选项
}

# 使用ctypes定义Windows API函数
if platform.system() == "Windows":
    # 定义ShellExecute函数
    ShellExecute = ctypes.windll.shell32.ShellExecuteW
    ShellExecute.argtypes = (ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_wchar_p, 
                             ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_int)
    ShellExecute.restype = ctypes.c_void_p
    
    # 定义常量
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_SHOW = 5
    SEE_MASK_NOCLOSEPROCESS = 0x00000040

class ShutdownSchedule:
    def __init__(self, name, shutdown_type, time, days, enabled=True, one_time=False, app=None):
        self.name = name
        self.shutdown_type = shutdown_type
        self.time = time
        self.days = days
        self.enabled = enabled
        self.one_time = one_time
        self.thread = None
        self.running = False
        self.executed = False
        self.app = app  # 存储应用引用
        
    def to_dict(self):
        return {
            "name": self.name,
            "type": self.shutdown_type,
            "time": self.time,
            "days": self.days,
            "enabled": self.enabled,
            "one_time": self.one_time
        }
    
    @staticmethod
    def from_dict(data, app=None):
        return ShutdownSchedule(
            data["name"],
            data["type"],
            data["time"],
            data["days"],
            data.get("enabled", True),
            data.get("one_time", False),
            app  # 传递应用引用
        )
    
    def start(self):
        if not self.enabled or self.running:
            return
            
        self.running = True
        self.thread = threading.Thread(target=self._schedule_check, daemon=True)
        self.thread.start()
        logging.info(f"计划 '{self.name}' 已启动")
    
    def stop(self):
        if self.running:
            self.running = False
            logging.info(f"计划 '{self.name}' 已停止")
    
    def _schedule_check(self):
        while self.running:
            try:
                now = datetime.datetime.now()
                current_weekday = now.isoweekday()
                
                # 检查日期条件
                if not self.one_time and current_weekday not in self.days:
                    time.sleep(10)
                    continue
                
                # 解析计划时间
                scheduled_time = datetime.datetime.strptime(self.time, "%H:%M")
                scheduled_time = now.replace(
                    hour=scheduled_time.hour, 
                    minute=scheduled_time.minute, 
                    second=0, 
                    microsecond=0
                )
                
                # 计算时间差
                time_diff = (scheduled_time - now).total_seconds()
                
                # 处理时间差为负的情况（计划时间已过）
                if time_diff < 0:
                    if self.one_time:
                        # 一次性任务已经过期
                        self.stop()
                        if self.app and self.app.root:
                            self.app.root.after(0, self.app.remove_executed_schedule, self.name)
                        return
                    # 重复任务则计算到第二天的时间
                    time_diff += 24 * 3600
                
                # 当时间差在0-60秒之间时执行
                if 0 < time_diff <= 60:
                    logging.info(f"计划 '{self.name}' 即将执行: {self.shutdown_type} (等待 {time_diff} 秒)")
                    
                    # 精确等待到执行时间
                    start_wait = time.time()
                    while time.time() - start_wait < time_diff and self.running:
                        time.sleep(0.5)
                    
                    if self.running:
                        self.execute_shutdown()
                    
                    if self.one_time:
                        self.executed = True
                        self.stop()
                        if self.app and self.app.root:
                            self.app.root.after(0, self.app.remove_executed_schedule, self.name)
                    
                    # 执行后休眠60秒，避免重复执行
                    time.sleep(60)
                else:
                    # 时间未到，休眠10秒再次检查
                    time.sleep(10)
            except Exception as e:
                logging.error(f"计划 '{self.name}' 执行出错: {str(e)}")
                time.sleep(10)
    
    def execute_shutdown(self):
        command = SHUTDOWN_TYPES.get(self.shutdown_type, "")
        if command:
            try:
                logging.info(f"执行命令: {command}")
                
                # 使用更可靠的方式执行命令
                run_as_admin = True
                if self.app:
                    run_as_admin = self.app.config.get("run_as_admin", True)
                
                if run_as_admin:
                    # 以管理员权限执行命令
                    self.execute_as_admin(command)
                else:
                    # 普通方式执行
                    subprocess.run(command, shell=True, check=True)
                
                logging.info(f"命令执行成功: {command}")
            except Exception as e:
                error_msg = f"执行关机命令失败: {str(e)}"
                logging.error(error_msg)
    
    def execute_as_admin(self, command):
        """以管理员权限执行命令"""
        try:
            # 使用ctypes执行命令
            if platform.system() == "Windows":
                # 创建进程时请求管理员权限
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = 0
                
                # 使用CREATE_NEW_CONSOLE避免权限问题
                subprocess.run(
                    command, 
                    shell=True, 
                    startupinfo=startupinfo,
                    creationflags=subprocess.CREATE_NEW_CONSOLE
                )
            else:
                subprocess.run(command, shell=True, check=True)
        except Exception as e:
            # 如果普通方式失败，尝试使用ShellExecute提升权限
            if platform.system() == "Windows":
                try:
                    # 使用ctypes调用ShellExecute
                    ShellExecute(0, "runas", "cmd.exe", f"/c {command}", None, 0)
                except Exception as admin_e:
                    logging.error(f"使用管理员权限执行失败: {str(admin_e)}")
                    raise admin_e
            else:
                raise e

class LazyShutdownApp:
    def __init__(self, root, icon_path):
        self.root = root
        self.icon_path = icon_path  # 存储图标路径
        self.root.title(APP_NAME)        
        self.root.geometry("600x500")
        
        # 设置全局样式
        style = ttk.Style()
        style.configure(".", font=("微软雅黑", 10))
        style.configure("TButton", padding=6)
        style.configure("TCheckbutton", padding=8)
        
        self.center_window(self.root, 600, 500)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.resizable(True, True)
        self.root.bind("<Unmap>", self.on_minimize)
        
        # 加载配置
        self.config = self.load_config()
        
        # 初始化托盘图标
        self.tray_icon = None
        self.tray_running = False
        
        # 初始化日志
        self.init_logging()
        
        # 创建UI
        self.create_widgets()
        
        # 启动所有计划
        self.start_all_schedules()
        
        # 设置自动启动
        self.set_auto_start(self.config.get("auto_start", False))
        
        # 设置全局热键
        self.show_main_window_hotkey = None
        self.setup_hotkey()
        
        # 确保程序有管理员权限
        self.check_admin_privileges()
    
    def check_admin_privileges(self):
        """检查并确保程序以管理员权限运行"""
        if platform.system() == "Windows":
            try:
                # 检查当前是否以管理员权限运行
                if ctypes.windll.shell32.IsUserAnAdmin() == 0:
                    if self.config.get("run_as_admin", True):
                        logging.warning("程序未以管理员权限运行，关机操作可能失败")
                        
                        # 询问用户是否重新启动
                        if messagebox.askyesno(
                            "权限警告", 
                            "程序需要管理员权限才能可靠执行关机操作。\n是否立即以管理员权限重新启动?"
                        ):
                            self.restart_as_admin()
            except Exception as e:
                logging.error(f"检查管理员权限失败: {str(e)}")
    
    def restart_as_admin(self):
        """以管理员权限重新启动程序"""
        if platform.system() == "Windows":
            try:
                # 获取当前可执行文件路径
                exe_path = sys.executable if getattr(sys, 'frozen', False) else sys.argv[0]
                
                # 使用ctypes以管理员权限运行
                ShellExecute(0, "runas", exe_path, " ".join(sys.argv[1:]), None, 1)
                
                # 退出当前实例
                self.quit_app()
            except Exception as e:
                logging.error(f"重新启动为管理员失败: {str(e)}")
                messagebox.showerror("错误", "无法以管理员权限重新启动程序")
    
    def init_logging(self):
        """初始化日志系统"""
        log_file = CONFIG_DIR / "lazy_shutdown.log"
    
        # 确保日志目录存在
        log_file.parent.mkdir(parents=True, exist_ok=True)
    
        # 使用文件处理器并指定编码
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    
        logger = logging.getLogger()
        logger.setLevel(logging.DEBUG)
        logger.addHandler(file_handler)
    
        logging.info(f"{APP_NAME} 启动")
    
    def center_window(self, window, width, height):
        """将窗口居中显示"""
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2) - 30
        
        window.geometry(f"{width}x{height}+{x}+{y}")
    
    def setup_hotkey(self):
        """设置全局热键"""
        hotkey = self.config.get("hotkey", "ctrl+alt+l")
        
        try:
            if self.show_main_window_hotkey is not None:
                keyboard.remove_hotkey(self.show_main_window_hotkey)
        except Exception as e:
            error_msg = f"移除热键时出错: {e}"
            logging.error(error_msg)
            
        try:
            self.show_main_window_hotkey = keyboard.add_hotkey(
                hotkey, 
                self.show_main_window_from_hotkey
            )
            logging.info(f"热键 '{hotkey}' 注册成功")
        except Exception as e:
            error_msg = f"设置热键失败: {e}"
            logging.error(error_msg)
            messagebox.showerror("热键错误", f"无法设置热键 '{hotkey}'，请检查是否冲突")
    
    def show_main_window_from_hotkey(self):
        """通过热键显示主窗口"""
        if self.root.state() == 'iconic' or not self.root.winfo_viewable() or self.tray_icon:
            self.root.after(0, self.show_main_window)
    
    def create_widgets(self):
        """创建主界面UI"""
        # 主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 标题
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        
        title_label = ttk.Label(
            title_frame, 
            text="关机计划管理",
            font=("微软雅黑", 14, "bold")
        )
        title_label.pack(side=tk.LEFT)
        
        # 计划列表
        self.schedule_frame = ttk.Frame(main_frame)
        self.schedule_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建列表容器
        self.canvas = tk.Canvas(self.schedule_frame)
        self.scrollbar = ttk.Scrollbar(
            self.schedule_frame, 
            orient=tk.VERTICAL, 
            command=self.canvas.yview
        )
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 底部按钮
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 设置按钮
        settings_button = ttk.Button(
            bottom_frame,
            text="设置",
            command=self.show_settings,
            width=10
        )
        settings_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # 删除按钮
        delete_button = ttk.Button(
            bottom_frame,
            text="删除",
            command=self.show_delete_dialog,
            width=10
        )
        delete_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # 新建按钮
        new_button = ttk.Button(
            bottom_frame,
            text="新建计划",
            command=self.create_new_schedule,
            width=10
        )
        new_button.pack(side=tk.RIGHT)
        
        # 加载计划
        self.load_schedules()
    
    def load_schedules(self):
        """加载计划到界面"""
        # 清除现有计划
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        if not self.config.get("schedules", []):
            empty_label = ttk.Label(
                self.scrollable_frame,
                text="当前没有关机计划",
                font=("微软雅黑", 12),
                anchor=tk.CENTER
            )
            empty_label.pack(fill=tk.X, pady=20)
            return
        
        # 添加计划
        for idx, schedule_data in enumerate(self.config["schedules"]):
            schedule = ShutdownSchedule.from_dict(schedule_data, self)  # 传递self引用
            self.add_schedule_to_ui(schedule, idx)
    
    def add_schedule_to_ui(self, schedule, idx):
        """添加单个计划到UI"""
        frame = ttk.Frame(self.scrollable_frame, relief=tk.GROOVE, padding=10)
        frame.pack(fill=tk.X, pady=5, padx=5)
        
        # 序号
        number_label = ttk.Label(frame, text=f"{idx+1}.", width=3)
        number_label.grid(row=0, column=0, rowspan=2, padx=(0, 10))
        
        # 计划名称
        name_label = ttk.Label(frame, text=schedule.name, font=("微软雅黑", 10, "bold"))
        name_label.grid(row=0, column=1, sticky=tk.W)
        
        # 计划详情
        days_str = "每天" if len(schedule.days) == 7 else "周" + "".join(str(d) for d in schedule.days)
        details_text = f"{schedule.shutdown_type} @ {schedule.time} ({days_str})"
        
        if schedule.one_time:
            details_text += " [单次]"
            
        details_label = ttk.Label(frame, text=details_text)
        details_label.grid(row=1, column=1, sticky=tk.W)
        
        # 开关
        switch_var = tk.BooleanVar(value=schedule.enabled)
        switch = ttk.Checkbutton(
            frame, 
            text="启用" if schedule.enabled else "禁用",
            variable=switch_var,
            command=lambda s=schedule, var=switch_var: self.toggle_schedule(s, var)
        )
        switch.grid(row=0, column=2, rowspan=2, padx=10)
        
        # 右键菜单
        frame.bind("<Button-3>", lambda e, s=schedule: self.show_schedule_context_menu(e, s))
        
        # 存储引用
        schedule.ui_frame = frame
        schedule.switch_var = switch_var
    
    def toggle_schedule(self, schedule, var):
        """切换计划开关状态"""
        schedule.enabled = var.get()
        schedule.switch_var.set(schedule.enabled)
        
        for child in schedule.ui_frame.winfo_children():
            if isinstance(child, ttk.Checkbutton):
                child.config(text="启用" if schedule.enabled else "禁用")
        
        if schedule.enabled:
            schedule.start()
        else:
            schedule.stop()
        
        self.save_config()
    
    def show_schedule_context_menu(self, event, schedule):
        """显示计划右键菜单"""
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="修改计划", command=lambda: self.modify_schedule(schedule))
        menu.add_command(label="单次执行计划", command=lambda: self.one_time_execution(schedule))
        menu.add_command(label="删除计划", command=lambda: self.delete_schedule(schedule))
        menu.tk_popup(event.x_root, event.y_root)
    
    def one_time_execution(self, schedule):
        """创建单次执行计划"""
        now = datetime.datetime.now()
        next_minute = now + datetime.timedelta(minutes=1)
        time_str = next_minute.strftime("%H:%M")
        
        one_time_schedule = ShutdownSchedule(
            f"{schedule.name} (单次)",
            schedule.shutdown_type,
            time_str,
            [now.isoweekday()],
            True,
            True,
            self  # 传递应用引用
        )
        
        self.config["schedules"].append(one_time_schedule.to_dict())
        self.save_config()
        self.load_schedules()
        one_time_schedule.start()
        
        messagebox.showinfo("单次执行", f"已创建单次执行计划，将在 {time_str} 执行")
    
    def remove_executed_schedule(self, schedule_name):
        """移除已执行的一次性计划"""
        self.config["schedules"] = [s for s in self.config["schedules"] if s["name"] != schedule_name]
        self.save_config()
        self.load_schedules()
    
    def create_new_schedule(self):
        """创建新计划"""
        dialog = ScheduleDialog(self.root, "新建关机计划", self.icon_path)
        self.center_window(dialog.top, 500, 450)
        self.root.wait_window(dialog.top)
        
        if dialog.result:
            # 创建计划对象时传递应用引用
            new_schedule = ShutdownSchedule(
                dialog.result.name,
                dialog.result.shutdown_type,
                dialog.result.time,
                dialog.result.days,
                True,
                dialog.result.one_time,
                self  # 传递应用引用
            )
            
            self.config["schedules"].append(new_schedule.to_dict())
            self.save_config()
            self.load_schedules()
            new_schedule.start()
    
    def modify_schedule(self, schedule):
        """修改现有计划"""
        idx = next((i for i, s in enumerate(self.config["schedules"]) if s["name"] == schedule.name), -1)
        if idx == -1:
            return
            
        # 先停止原计划
        schedule.stop()
        
        dialog = ScheduleDialog(
            self.root, 
            "修改计划",
            self.icon_path,
            name=schedule.name,
            shutdown_type=schedule.shutdown_type,
            time=schedule.time,
            days=schedule.days
        )
        self.center_window(dialog.top, 500, 450)
        self.root.wait_window(dialog.top)
        
        if dialog.result:
            # 直接更新原计划对象而不是创建新对象
            schedule.name = dialog.result.name
            schedule.shutdown_type = dialog.result.shutdown_type
            schedule.time = dialog.result.time
            schedule.days = dialog.result.days
            schedule.one_time = dialog.result.one_time
            
            # 更新配置文件
            self.config["schedules"][idx] = schedule.to_dict()
            self.save_config()
            
            # 更新UI显示
            self.update_schedule_in_ui(schedule)
            
            if schedule.enabled:
                schedule.start()
    
    def update_schedule_in_ui(self, schedule):
        """更新单个计划在UI中的显示"""
        if not hasattr(schedule, 'ui_frame'):
            return
            
        # 清除原有内容
        for widget in schedule.ui_frame.winfo_children():
            widget.destroy()
        
        # 序号保持不变
        number_label = ttk.Label(schedule.ui_frame, text=schedule.ui_frame.winfo_children()[0].cget("text"), width=3)
        number_label.grid(row=0, column=0, rowspan=2, padx=(0, 10))
        
        # 更新计划名称
        name_label = ttk.Label(schedule.ui_frame, text=schedule.name, font=("微软雅黑", 10, "bold"))
        name_label.grid(row=0, column=1, sticky=tk.W)
        
        # 更新计划详情
        days_str = "每天" if len(schedule.days) == 7 else "周" + "".join(str(d) for d in schedule.days)
        details_text = f"{schedule.shutdown_type} @ {schedule.time} ({days_str})"
        
        if schedule.one_time:
            details_text += " [单次]"
            
        details_label = ttk.Label(schedule.ui_frame, text=details_text)
        details_label.grid(row=1, column=1, sticky=tk.W)
        
        # 更新开关
        switch_var = tk.BooleanVar(value=schedule.enabled)
        switch = ttk.Checkbutton(
            schedule.ui_frame, 
            text="启用" if schedule.enabled else "禁用",
            variable=switch_var,
            command=lambda s=schedule, var=switch_var: self.toggle_schedule(s, var)
        )
        switch.grid(row=0, column=2, rowspan=2, padx=10)
        
        # 重新绑定右键菜单
        schedule.ui_frame.bind("<Button-3>", lambda e, s=schedule: self.show_schedule_context_menu(e, s))
        
        # 更新引用
        schedule.switch_var = switch_var
    
    def delete_schedule(self, schedule):
        """删除单个计划"""
        if not messagebox.askyesno("确认删除", f"确定要删除计划 '{schedule.name}' 吗？"):
            return
        
        schedule.stop()
        self.config["schedules"] = [s for s in self.config["schedules"] if s["name"] != schedule.name]
        self.save_config()
        self.load_schedules()
    
    def show_delete_dialog(self):
        """显示批量删除对话框"""
        if not self.config.get("schedules", []):
            messagebox.showinfo("提示", "当前没有可删除的计划")
            return
            
        dialog = DeleteDialog(self.root, self.config["schedules"], self.icon_path)
        self.center_window(dialog.top, 550, 550)  # 高度增加
        self.root.wait_window(dialog.top)
        
        if dialog.selected_schedules:
            for name in dialog.selected_schedules:
                schedule = next((s for s in self.config["schedules"] if s["name"] == name), None)
                if schedule:
                    sched_obj = ShutdownSchedule.from_dict(schedule, self)
                    sched_obj.stop()
            
            self.config["schedules"] = [s for s in self.config["schedules"] if s["name"] not in dialog.selected_schedules]
            self.save_config()
            self.load_schedules()
    
    def show_settings(self):
        """显示设置窗口"""
        SettingsDialog(self.root, self.config, self, self.icon_path)
    
    def start_all_schedules(self):
        """启动所有启用的计划"""
        logging.info("启动所有计划")
        for schedule_data in self.config.get("schedules", []):
            schedule = ShutdownSchedule.from_dict(schedule_data, self)  # 传递应用引用
            if schedule.enabled:
                logging.info(f"启动计划: {schedule.name}")
                schedule.start()
    
    def stop_all_schedules(self):
        """停止所有计划"""
        logging.info("停止所有计划")
        for schedule_data in self.config.get("schedules", []):
            schedule = ShutdownSchedule.from_dict(schedule_data, self)
            schedule.stop()
    
    def set_auto_start(self, enable):
        """设置开机自启动"""
        app_name = "LazyShutdown"
        app_path = os.path.abspath(sys.argv[0])
        
        # 添加最小化参数
        if enable:
            app_path_with_args = f'"{app_path}" --minimized'
        else:
            app_path_with_args = app_path
        
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTO_START_KEY, 0, winreg.KEY_WRITE)
            
            if enable:
                winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, app_path_with_args)
                logging.info("设置开机自启动")
            else:
                try:
                    winreg.DeleteValue(key, app_name)
                    logging.info("取消开机自启动")
                except:
                    pass
            
            winreg.CloseKey(key)
            return True
        except Exception as e:
            error_msg = f"设置自启动失败: {e}"
            logging.error(error_msg)
            return False
    
    def create_tray_icon(self):
        """创建系统托盘图标"""
        if self.tray_icon:
            return
            
        # 创建托盘图标
        image = Image.new('RGBA', (64, 64), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        draw.rectangle((16, 16, 48, 48), fill="#4b6eaf")
        font = ImageFont.truetype("arial.ttf", 20) if os.name == 'nt' else ImageFont.load_default()
        draw.text((32, 32), "LS", fill="white", anchor="mm", font=font)
        
        menu = pystray.Menu(
            pystray.MenuItem("显示主界面", self.show_main_window),
            pystray.MenuItem("退出", self.quit_app)
        )
        
        self.tray_icon = pystray.Icon("lazy_shutdown", image, APP_NAME, menu)
    
    def show_main_window(self, icon=None, item=None):
        """从托盘恢复主窗口"""
        if self.tray_icon:
            try:
                self.tray_icon.stop()
            except Exception as e:
                logging.error(f"关闭托盘图标时出错: {e}")
            self.tray_icon = None
        
        self.root.deiconify()
        self.root.attributes('-topmost', True)
        self.root.after_idle(lambda: self.root.attributes('-topmost', False))
        self.root.lift()
    
    def minimize_to_tray(self):
        """最小化到托盘"""
        if not self.config.get("minimize_to_tray", True):
            return
            
        if self.config.get("hide_tray_icon", False):
            self.root.withdraw()
            logging.info("最小化到任务栏")
        else:
            self.root.withdraw()
            self.create_tray_icon()
            
            if self.tray_icon:
                self.tray_running = True
                
                # 使用单独的线程运行托盘图标
                threading.Thread(
                    target=self.run_tray_icon, 
                    daemon=True,
                    name="TrayIconThread"
                ).start()
                
                logging.info("最小化到系统托盘")
    
    def run_tray_icon(self):
        """运行托盘图标（在独立线程中）"""
        try:
            self.tray_icon.run()
        except Exception as e:
            logging.error(f"托盘图标运行出错: {str(e)}")
    
    def on_minimize(self, event):
        """处理最小化事件"""
        if event.widget == self.root and self.config.get("minimize_to_tray", True):
            self.minimize_to_tray()
    
    def quit_app(self, icon=None, item=None):
        """退出应用程序"""
        try:
            if self.tray_icon:
                self.tray_icon.stop()
        except:
            pass
            
        self.stop_all_schedules()
        logging.info("程序退出")
        self.root.destroy()
        sys.exit(0)
    
    def on_close(self):
        """处理窗口关闭事件"""
        if self.config.get("minimize_to_tray", True):
            self.minimize_to_tray()
        else:
            self.quit_app()
    
    def load_config(self):
        """加载配置文件"""
        try:
            if CONFIG_FILE.exists():
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception as e:
            error_msg = f"加载配置失败: {e}"
            logging.error(error_msg)
        return DEFAULT_CONFIG.copy()
    
    def save_config(self):
        """保存配置文件"""
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            logging.info("配置文件已保存")
            return True
        except Exception as e:
            error_msg = f"保存配置失败: {e}"
            logging.error(error_msg)
            return False

class ScheduleDialog:
    def __init__(self, parent, title, icon_path, name="", shutdown_type="关机", time="00:00", days=None):
        self.parent = parent
        self.result = None
        
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.geometry("500x450")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()
        self.top.attributes('-topmost', True)
        
        # 设置对话框图标
        if icon_path:
            try:
                self.top.iconbitmap(icon_path)
            except:
                pass
        
        # 主框架
        main_frame = ttk.Frame(self.top)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 内容框架
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        # 计划名称
        ttk.Label(content_frame, text="计划名称:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        self.name_var = tk.StringVar(value=name)
        name_entry = ttk.Entry(content_frame, textvariable=self.name_var, width=30)
        name_entry.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)
        
        # 关机类型
        ttk.Label(content_frame, text="操作类型:").grid(row=1, column=0, sticky=tk.W, pady=5, padx=5)
        self.type_var = tk.StringVar(value=shutdown_type)
        type_combo = ttk.Combobox(
            content_frame, 
            textvariable=self.type_var,
            values=list(SHUTDOWN_TYPES.keys()),
            state="readonly",
            width=27
        )
        type_combo.grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)
        
        # 执行时间
        ttk.Label(content_frame, text="执行时间:").grid(row=2, column=0, sticky=tk.W, pady=5, padx=5)
        self.time_var = tk.StringVar(value=time)
        time_entry = ttk.Entry(content_frame, textvariable=self.time_var, width=30)
        time_entry.grid(row=2, column=1, sticky=tk.W, pady=5, padx=5)
        ttk.Label(content_frame, text="格式: HH:MM (24小时制)").grid(row=3, column=1, sticky=tk.W, padx=5)
        
        # 重复日期
        ttk.Label(content_frame, text="重复日期:").grid(row=4, column=0, sticky=tk.W, pady=5, padx=5)
        
        days_frame = ttk.Frame(content_frame)
        days_frame.grid(row=4, column=1, sticky=tk.W, pady=5, padx=5)
        
        self.day_vars = []
        days_text = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
        
        for i in range(7):
            var = tk.BooleanVar(value=days and (i+1) in days if days else False)
            self.day_vars.append(var)
            row_idx = i // 4
            col_idx = i % 4
            cb = ttk.Checkbutton(days_frame, text=days_text[i], variable=var)
            cb.grid(row=row_idx, column=col_idx, padx=5, pady=2, sticky=tk.W)
        
        # 一次性任务选项
        self.one_time_var = tk.BooleanVar(value=False)
        one_time_cb = ttk.Checkbutton(
            content_frame,
            text="单次执行（执行后自动删除）",
            variable=self.one_time_var
        )
        one_time_cb.grid(row=5, column=0, columnspan=2, pady=10, sticky=tk.W, padx=5)
        
        # 按钮框架 - 固定在底部
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        
        # 确定按钮
        ok_button = ttk.Button(button_frame, text="确定", command=self.on_ok, width=10)
        ok_button.pack(side=tk.RIGHT, padx=10)
        
        # 取消按钮
        cancel_button = ttk.Button(button_frame, text="取消", command=self.top.destroy, width=10)
        cancel_button.pack(side=tk.RIGHT)
    
    def on_ok(self):
        """确定按钮处理"""
        name = self.name_var.get().strip()
        shutdown_type = self.type_var.get()
        time_str = self.time_var.get().strip()
        days = [i+1 for i, var in enumerate(self.day_vars) if var.get()]
        one_time = self.one_time_var.get()
        
        # 验证输入
        if not name:
            messagebox.showerror("错误", "请输入计划名称")
            return
            
        if not shutdown_type:
            messagebox.showerror("错误", "请选择操作类型")
            return
            
        if not time_str or not self.validate_time(time_str):
            messagebox.showerror("错误", "请输入有效的时间 (HH:MM)")
            return
            
        if not days and not one_time:
            messagebox.showerror("错误", "请至少选择一个日期或选择单次执行")
            return
            
        # 创建计划对象
        self.result = ShutdownSchedule(name, shutdown_type, time_str, days, True, one_time)
        self.top.destroy()
    
    def validate_time(self, time_str):
        """验证时间格式"""
        try:
            datetime.datetime.strptime(time_str, "%H:%M")
            return True
        except ValueError:
            return False

class DeleteDialog:
    def __init__(self, parent, schedules, icon_path):
        self.parent = parent
        self.selected_schedules = []
        
        self.top = tk.Toplevel(parent)
        self.top.title("删除计划")
        self.top.geometry("550x550")  # 高度增加
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()
        self.top.attributes('-topmost', True)
        
        # 设置对话框图标
        if icon_path:
            try:
                self.top.iconbitmap(icon_path)
            except:
                pass
        
        # 主框架
        main_frame = ttk.Frame(self.top)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 内容框架
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        # 标题
        ttk.Label(content_frame, text="选择要删除的计划:", font=("微软雅黑", 11, "bold")).pack(fill=tk.X, pady=(0, 15))
        
        # 计划列表容器
        list_container = ttk.Frame(content_frame)
        list_container.pack(fill=tk.BOTH, expand=True)
        
        # 计划列表
        self.listbox = tk.Listbox(
            list_container, 
            selectmode=tk.MULTIPLE,
            activestyle="none",
            width=50,
            height=15,
            font=("微软雅黑", 10)
        )
        scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.listbox.yview)
        self.listbox.config(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 添加计划
        for schedule in schedules:
            self.listbox.insert(tk.END, schedule["name"])
        
        # 全选按钮
        select_all_button = ttk.Button(
            content_frame,
            text="全选",
            command=self.select_all,
            width=12
        )
        select_all_button.pack(anchor=tk.W, pady=10)
        
        # 按钮框架 - 固定在底部
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(20, 10))  # 增加上边距
        
        # 确定按钮
        ok_button = ttk.Button(
            button_frame,
            text="确定",
            command=self.on_ok,
            width=15
        )
        ok_button.pack(side=tk.RIGHT, padx=10)
        
        # 取消按钮
        cancel_button = ttk.Button(
            button_frame,
            text="取消",
            command=self.top.destroy,
            width=15
        )
        cancel_button.pack(side=tk.RIGHT)
    
    def select_all(self):
        """全选计划"""
        self.listbox.selection_set(0, tk.END)
    
    def on_ok(self):
        """确定按钮处理 - 保存选择的计划"""
        selected_indices = self.listbox.curselection()
        self.selected_schedules = [self.listbox.get(i) for i in selected_indices]
        self.top.destroy()

class SettingsDialog:
    def __init__(self, parent, config, app, icon_path):
        self.parent = parent
        self.config = config
        self.app = app
        
        self.top = tk.Toplevel(parent)
        self.top.title("设置")
        self.top.geometry("500x550")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()
        self.top.attributes('-topmost', True)
        
        # 设置对话框图标
        if icon_path:
            try:
                self.top.iconbitmap(icon_path)
            except:
                pass
        
        # 主框架
        main_frame = ttk.Frame(self.top)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 内容框架
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        # 设置选项
        self.create_settings(content_frame)
        
        # 按钮框架 - 固定在底部
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        
        # 确定按钮
        ok_button = ttk.Button(button_frame, text="确定", command=self.on_ok, width=10)
        ok_button.pack(side=tk.RIGHT, padx=10)
        
        # 取消按钮
        cancel_button = ttk.Button(button_frame, text="取消", command=self.top.destroy, width=10)
        cancel_button.pack(side=tk.RIGHT)
    
    def create_settings(self, parent):
        """创建设置选项"""
        # 开机自启动
        auto_start_frame = ttk.Frame(parent)
        auto_start_frame.pack(fill=tk.X, pady=10)
        
        self.auto_start_var = tk.BooleanVar(value=self.config.get("auto_start", False))
        auto_start_cb = ttk.Checkbutton(
            auto_start_frame, 
            text="开机自动启动",
            variable=self.auto_start_var
        )
        auto_start_cb.pack(anchor=tk.W)
        
        # 最小化到托盘
        tray_frame = ttk.Frame(parent)
        tray_frame.pack(fill=tk.X, pady=10)
        
        self.tray_var = tk.BooleanVar(value=self.config.get("minimize_to_tray", True))
        tray_cb = ttk.Checkbutton(
            tray_frame, 
            text="最小化到系统托盘",
            variable=self.tray_var
        )
        tray_cb.pack(anchor=tk.W)
        
        # 隐藏托盘图标
        hide_tray_frame = ttk.Frame(parent)
        hide_tray_frame.pack(fill=tk.X, pady=10)
        
        self.hide_tray_var = tk.BooleanVar(value=self.config.get("hide_tray_icon", False))
        hide_tray_cb = ttk.Checkbutton(
            hide_tray_frame, 
            text="隐藏托盘图标（需设置热键）",
            variable=self.hide_tray_var
        )
        hide_tray_cb.pack(anchor=tk.W)
        
        # 热键设置
        hotkey_frame = ttk.Frame(parent)
        hotkey_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(hotkey_frame, text="唤醒热键:").pack(side=tk.LEFT, padx=(0, 10))
        
        self.hotkey_var = tk.StringVar(value=self.config.get("hotkey", "ctrl+alt+l"))
        hotkey_entry = ttk.Entry(hotkey_frame, textvariable=self.hotkey_var, width=20)
        hotkey_entry.pack(side=tk.LEFT)
        ttk.Label(hotkey_frame, text="(例如: ctrl+alt+l)").pack(side=tk.LEFT, padx=(5, 0))
        
        # 管理员权限
        admin_frame = ttk.Frame(parent)
        admin_frame.pack(fill=tk.X, pady=10)
        
        self.admin_var = tk.BooleanVar(value=self.config.get("run_as_admin", True))
        admin_cb = ttk.Checkbutton(
            admin_frame, 
            text="以管理员权限运行（推荐）",
            variable=self.admin_var
        )
        admin_cb.pack(anchor=tk.W)
        ttk.Label(admin_frame, text="* 需要管理员权限才能可靠执行关机操作", font=("微软雅黑", 8)).pack(anchor=tk.W, padx=25)
        
        # 关于
        about_frame = ttk.LabelFrame(parent, text="关于")
        about_frame.pack(fill=tk.X, pady=10)
        
        info_text = f"{APP_NAME} v1.7\n\n开发者: Jumao\n联系邮箱: jumaozhixing@outlook.com"
        
        info_label = ttk.Label(about_frame, text=info_text, justify=tk.LEFT, font=("微软雅黑", 9))
        info_label.pack(fill=tk.X, padx=10, pady=10)
    
    def on_ok(self):
        """保存设置并关闭窗口"""
        # 更新配置
        self.config["auto_start"] = self.auto_start_var.get()
        self.config["minimize_to_tray"] = self.tray_var.get()
        self.config["hide_tray_icon"] = self.hide_tray_var.get()
        self.config["hotkey"] = self.hotkey_var.get().strip()
        self.config["run_as_admin"] = self.admin_var.get()
        
        # 应用设置
        self.app.set_auto_start(self.config["auto_start"])
        self.app.setup_hotkey()
        self.app.save_config()
        
        # 检查是否需要管理员权限
        if self.config["run_as_admin"]:
            self.app.check_admin_privileges()
        
        self.top.destroy()

def is_admin():
    """检查是否以管理员权限运行"""
    if platform.system() == "Windows":
        try:
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except:
            return False
    return False

def main():
    # 确保单实例运行
    if platform.system() == "Windows":
        mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "LazyShutdownMutex")
        if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
            ctypes.windll.user32.MessageBoxW(0, "程序已在运行中", APP_NAME, 0)
            return
    
    # 检查启动参数
    start_minimized = "--minimized" in sys.argv
    
    # 检查是否需要以管理员权限运行
    run_as_admin = True
    try:
        # 尝试加载配置
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
                run_as_admin = config.get("run_as_admin", True)
    except:
        pass
    
    # 如果配置要求管理员权限但当前不是管理员，则请求提升权限
    if run_as_admin and not is_admin() and platform.system() == "Windows":
        try:
            # 获取当前可执行文件路径
            exe_path = sys.executable if getattr(sys, 'frozen', False) else sys.argv[0]
            
            # 使用ctypes以管理员权限运行
            ShellExecute(0, "runas", exe_path, " ".join(sys.argv[1:]), None, 1)
            sys.exit(0)
        except Exception as e:
            # 记录错误日志
            log_file = CONFIG_DIR / "lazy_shutdown.log"
            with open(log_file, "a", encoding="utf-8") as log:
                log.write(f"[{datetime.datetime.now()}] 请求管理员权限失败: {str(e)}\n")
    
    root = tk.Tk()
    
    # 立即隐藏窗口，避免闪现
    root.withdraw()
    
    try:
        if platform.system() == "Windows":
            # 优先尝试使用绝对路径
            base_path = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(base_path, "icon.ico")
            
            # 对于打包后的程序
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
                icon_path = os.path.join(base_path, "icon.ico")
                
            # 设置根窗口图标
            root.iconbitmap(icon_path)
        else:
            # Linux/macOS 可以使用其他格式
            icon_path = None
    except Exception as e:
        # 记录错误日志
        log_file = CONFIG_DIR / "lazy_shutdown.log"
        with open(log_file, "a", encoding="utf-8") as log:
            log.write(f"[{datetime.datetime.now()}] 设置图标失败: {str(e)}\n")
        icon_path = None
    
    # 存储图标路径供其他窗口使用
    root.icon_path = icon_path
    
    app = LazyShutdownApp(root, icon_path)
    
    # 如果是自启动模式，直接最小化
    if start_minimized:
        # 立即最小化到托盘，不显示窗口
        app.minimize_to_tray()
    else:
        # 显示窗口
        root.deiconify()
        
    try:
        root.mainloop()
    except KeyboardInterrupt:
        # 记录错误日志
        log_file = CONFIG_DIR / "lazy_shutdown.log"
        with open(log_file, "a", encoding="utf-8") as log:
            log.write(f"[{datetime.datetime.now()}] 程序被用户中断\n")
        app.quit_app()

if __name__ == "__main__":
    main()