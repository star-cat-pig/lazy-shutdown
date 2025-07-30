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
import logging.handlers
from pathlib import Path
import shutil
import tempfile
import webbrowser

# 常量定义
APP_NAME = "懒人关机器"
GITHUB_URL = "https://github.com/star-cat-pig/lazy-shutdown/releases/latest"

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
    "run_as_admin": True,
    "use_task_scheduler": False,
    # 守护进程配置
    "guardian_enabled": False,
    "guardian_autostart": False,
    "guardian_terminate_taskmgr": True,
    "guardian_autorestart": True,
    "guardian_hide_window": True,
    "guardian_show_window": False,
    "guardian_show_console": False  # 新增控制台显示选项
}

# Windows API函数
if platform.system() == "Windows":
    ShellExecute = ctypes.windll.shell32.ShellExecuteW
    ShellExecute.argtypes = (ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_wchar_p, 
                             ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_int)
    ShellExecute.restype = ctypes.c_void_p
    
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
        self.app = app
        
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
            app
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
                
                if not self.one_time and current_weekday not in self.days:
                    time.sleep(10)
                    continue
                
                scheduled_time = datetime.datetime.strptime(self.time, "%H:%M")
                scheduled_time = now.replace(
                    hour=scheduled_time.hour, 
                    minute=scheduled_time.minute, 
                    second=0, 
                    microsecond=0
                )
                
                time_diff = (scheduled_time - now).total_seconds()
                
                if time_diff < 0:
                    if self.one_time:
                        self.stop()
                        if self.app and self.app.root:
                            self.app.root.after(0, self.app.remove_executed_schedule, self.name)
                        return
                    time_diff += 24 * 3600
                
                if 0 < time_diff <= 60:
                    logging.info(f"计划 '{self.name}' 即将执行: {self.shutdown_type} (等待 {time_diff} 秒)")
                    
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
                    
                    time.sleep(60)
                else:
                    time.sleep(10)
            except Exception as e:
                logging.error(f"计划 '{self.name}' 执行出错: {str(e)}")
                time.sleep(10)
    
    def execute_shutdown(self):
        command = SHUTDOWN_TYPES.get(self.shutdown_type, "")
        if command:
            try:
                logging.info(f"执行命令: {command}")
                
                run_as_admin = True
                if self.app:
                    run_as_admin = self.app.config.get("run_as_admin", True)
                
                if run_as_admin:
                    self.execute_as_admin(command)
                else:
                    subprocess.run(command, shell=True, check=True)
                
                logging.info(f"命令执行成功: {command}")
            except Exception as e:
                error_msg = f"执行关机命令失败: {str(e)}"
                logging.error(error_msg)
    
    def execute_as_admin(self, command):
        try:
            if platform.system() == "Windows":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = 0
                
                subprocess.run(
                    command, 
                    shell=True, 
                    startupinfo=startupinfo,
                    creationflags=subprocess.CREATE_NEW_CONSOLE
                )
            else:
                subprocess.run(command, shell=True, check=True)
        except Exception as e:
            if platform.system() == "Windows":
                try:
                    ShellExecute(0, "runas", "cmd.exe", f"/c {command}", None, 0)
                except Exception as admin_e:
                    logging.error(f"使用管理员权限执行失败: {str(admin_e)}")
                    raise admin_e
            else:
                raise e

class LazyShutdownApp:
    def __init__(self, root, icon_path):
        self.root = root
        self.icon_path = icon_path
        self.root.title(APP_NAME)        
        self.root.geometry("600x500")
        
        style = ttk.Style()
        style.configure(".", font=("微软雅黑", 10))
        style.configure("TButton", padding=6)
        style.configure("TCheckbutton", padding=8)
        
        self.center_window(self.root, 600, 500)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.resizable(True, True)
        self.root.bind("<Unmap>", self.on_minimize)
        self.guardian_process = None
        self.guardian_monitor_running = False
        self.guardian_monitor_thread = None
        
        # 加载配置
        self.config = self.load_config()
        
        # 初始化
        self.tray_icon = None
        self.tray_running = False
        self.init_logging()
        self.create_widgets()
        self.start_all_schedules()
        self.set_auto_start(self.config.get("auto_start", False))
        self.show_main_window_hotkey = None
        self.setup_hotkey()
        self.check_admin_privileges()
        self.start_heartbeat()
        
        # 设置守护进程
        self.setup_guardian()
        self.start_guardian_monitor()

    def setup_guardian(self):
        """根据配置设置守护进程"""
        enabled = self.config.get("guardian_enabled", False)
        
        # 如果要求禁用
        if not enabled:
            self.stop_guardian_monitor()
            self.stop_guardian()
            logging.info("守护进程已禁用")
            return
        
        # 如果要求启用
        # 检查是否已经运行
        if self.guardian_process and self.guardian_process.poll() is None:
            # 如果配置发生变化，需要重启
            current_args = self.get_guardian_args()
            new_args = self.build_guardian_args()
            
            if current_args != new_args:
                logging.info("守护进程配置已变更，正在重启守护进程...")
                self.stop_guardian()
                self.start_guardian()
        else:
            # 没有运行，启动守护进程
            self.start_guardian()
        
        # 确保监控线程运行
        if not self.guardian_monitor_running:
            self.start_guardian_monitor()
    
    def get_guardian_args(self):
        """获取当前守护进程的启动参数"""
        return {
            "show_console": self.config.get("guardian_show_console", False)
        }
    
    def build_guardian_args(self):
        """构建新的守护进程启动参数"""
        return {
            "show_console": self.config.get("guardian_show_console", False)
        }
    
    def start_guardian_monitor(self):
        """启动守护进程监控线程"""
        if self.guardian_monitor_running:
            return
            
        if self.config.get("guardian_enabled", False):
            self.guardian_monitor_running = True
            self.guardian_monitor_thread = threading.Thread(
                target=self.monitor_guardian,
                daemon=True
            )
            self.guardian_monitor_thread.start()
            logging.info("启动守护进程监控线程")
    
    def stop_guardian_monitor(self):
        """停止守护进程监控"""
        if self.guardian_monitor_running:
            self.guardian_monitor_running = False
            logging.info("停止守护进程监控")
    
    def monitor_guardian(self):
        """监控守护进程状态"""
        while self.guardian_monitor_running:
            time.sleep(10)  # 每10秒检查一次
            
            # 检查配置是否仍然启用守护进程
            if not self.config.get("guardian_enabled", False):
                self.stop_guardian_monitor()
                break
                
            # 检查守护进程是否在运行
            if self.guardian_process is None or self.guardian_process.poll() is not None:
                logging.warning("守护进程未运行，尝试重新启动")
                self.start_guardian()
    
    def start_guardian(self):
        """启动守护进程"""
        # 如果已经在运行，先停止
        if self.guardian_process and self.guardian_process.poll() is None:
            self.stop_guardian()
        
        try:
            if getattr(sys, 'frozen', False):
                base_dir = os.path.dirname(sys.executable)
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))
                
            guardian_path = os.path.join(base_dir, "guardian.exe")
            
            # 构建命令行参数
            args = ["--minimized"]  # 始终传递最小化参数
            
            # 如果配置要求显示控制台，添加参数
            if self.config.get("guardian_show_console", False):
                args.append("--console")
            
            # 启动进程
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            
            self.guardian_process = subprocess.Popen(
                [guardian_path] + args,
                startupinfo=startupinfo,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            logging.info(f"守护进程已启动，参数: {args}")
        except Exception as e:
            logging.error(f"启动守护进程失败: {e}")
    
    def stop_guardian(self):
        """停止守护进程"""
        if self.guardian_process:
            try:
                self.guardian_process.terminate()
                try:
                    self.guardian_process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    self.guardian_process.kill()
                logging.info("守护进程已停止")
            except Exception as e:
                logging.error(f"停止守护进程失败: {e}")
            finally:
                self.guardian_process = None
    
    def start_heartbeat(self):
        def heartbeat():
            try:
                if self.show_main_window_hotkey is None:
                    self.setup_hotkey()
            except Exception as e:
                logging.error(f"心跳检测中热键检查失败: {e}")
            
            self.root.after(60000, heartbeat)
        
        heartbeat()
    
    def check_admin_privileges(self):
        if platform.system() == "Windows":
            try:
                if ctypes.windll.shell32.IsUserAnAdmin() == 0:
                    if self.config.get("run_as_admin", True):
                        logging.warning("程序未以管理员权限运行，关机操作可能失败")
                        
                        dialog_root = tk.Toplevel(self.root)
                        dialog_root.withdraw()
                        dialog_root.attributes('-topmost', True)
                        
                        response = messagebox.askyesno(
                            "权限警告", 
                            "程序需要管理员权限才能可靠执行关机操作。\n是否立即以管理员权限重新启动?",
                            parent=dialog_root
                        )
                        
                        dialog_root.destroy()
                        
                        if response:
                            self.restart_as_admin()
            except Exception as e:
                logging.error(f"检查管理员权限失败: {str(e)}")
    
    def restart_as_admin(self):
        if platform.system() == "Windows":
            try:
                exe_path = sys.executable if getattr(sys, 'frozen', False) else sys.argv[0]
                ShellExecute(0, "runas", exe_path, " ".join(sys.argv[1:]), None, 1)
                self.quit_app()
            except Exception as e:
                logging.error(f"重新启动为管理员失败: {str(e)}")
                messagebox.showerror("错误", "无法以管理员权限重新启动程序")
    
    def init_logging(self):
        log_file = CONFIG_DIR / "lazy_shutdown.log"
        log_file.parent.mkdir(parents=True, exist_ok=True)
    
        handler = logging.handlers.RotatingFileHandler(
            log_file, 
            maxBytes=1024*1024,
            backupCount=3,
            encoding='utf-8'
        )
        handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    
        logger = logging.getLogger()
        logger.setLevel(logging.DEBUG)
        logger.addHandler(handler)
    
        logging.info(f"{APP_NAME} 启动")
    
    def center_window(self, window, width, height):
        window.update_idletasks()
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2) - 30
        
        window.geometry(f"{width}x{height}+{x}+{y}")
    
    def setup_hotkey(self):
        hotkey = self.config.get("hotkey", "ctrl+alt+l")
        
        def hotkey_callback():
            try:
                self.show_main_window_from_hotkey()
            except Exception as e:
                logging.error(f"热键回调出错: {e}")
                self.root.after(1000, self.setup_hotkey)
        
        try:
            if self.show_main_window_hotkey is not None:
                try:
                    keyboard.remove_hotkey(self.show_main_window_hotkey)
                except:
                    pass
        
            self.show_main_window_hotkey = keyboard.add_hotkey(
                hotkey, 
                hotkey_callback,
                suppress=True
            )
            logging.info(f"热键 '{hotkey}' 注册成功")
        except Exception as e:
            logging.error(f"热键设置失败: {e}")
            self.root.after(5000, self.setup_hotkey)
    
    def show_main_window_from_hotkey(self):
        try:
            if self.root.state() == 'iconic' or not self.root.winfo_viewable() or self.tray_icon:
                self.root.after(0, self.show_main_window)
        except Exception as e:
            logging.error(f"热键唤醒失败: {e}")
            try:
                self.root.deiconify()
                self.root.lift()
            except:
                pass
    
    def create_widgets(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        
        title_label = ttk.Label(
            title_frame, 
            text="关机计划管理",
            font=("微软雅黑", 14, "bold")
        )
        title_label.pack(side=tk.LEFT)
        
        self.schedule_frame = ttk.Frame(main_frame)
        self.schedule_frame.pack(fill=tk.BOTH, expand=True)
        
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
        
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(10, 0))
        
        settings_button = ttk.Button(
            bottom_frame,
            text="设置",
            command=self.show_settings,
            width=10
        )
        settings_button.pack(side=tk.LEFT, padx=(0, 10))
        
        delete_button = ttk.Button(
            bottom_frame,
            text="删除",
            command=self.show_delete_dialog,
            width=10
        )
        delete_button.pack(side=tk.LEFT, padx=(0, 10))
        
        new_button = ttk.Button(
            bottom_frame,
            text="新建计划",
            command=self.create_new_schedule,
            width=10
        )
        new_button.pack(side=tk.RIGHT)
        
        self.load_schedules()
    
    def load_schedules(self):
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
        
        for idx, schedule_data in enumerate(self.config["schedules"]):
            schedule = ShutdownSchedule.from_dict(schedule_data, self)
            self.add_schedule_to_ui(schedule, idx)
    
    def add_schedule_to_ui(self, schedule, idx):
        frame = ttk.Frame(self.scrollable_frame, relief=tk.GROOVE, padding=10)
        frame.pack(fill=tk.X, pady=5, padx=5)
        
        number_label = ttk.Label(frame, text=f"{idx+1}.", width=3)
        number_label.grid(row=0, column=0, rowspan=2, padx=(0, 10))
        
        name_label = ttk.Label(frame, text=schedule.name, font=("微软雅黑", 10, "bold"))
        name_label.grid(row=0, column=1, sticky=tk.W)
        
        days_str = "每天" if len(schedule.days) == 7 else "周" + "".join(str(d) for d in schedule.days)
        details_text = f"{schedule.shutdown_type} @ {schedule.time} ({days_str})"
        
        if schedule.one_time:
            details_text += " [单次]"
            
        details_label = ttk.Label(frame, text=details_text)
        details_label.grid(row=1, column=1, sticky=tk.W)
        
        switch_var = tk.BooleanVar(value=schedule.enabled)
        switch = ttk.Checkbutton(
            frame, 
            text="启用" if schedule.enabled else "禁用",
            variable=switch_var,
            command=lambda s=schedule, var=switch_var: self.toggle_schedule(s, var)
        )
        switch.grid(row=0, column=2, rowspan=2, padx=10)
        
        frame.bind("<Button-3>", lambda e, s=schedule: self.show_schedule_context_menu(e, s))
        
        schedule.ui_frame = frame
        schedule.switch_var = switch_var
    
    def toggle_schedule(self, schedule, var):
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
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="修改计划", command=lambda: self.modify_schedule(schedule))
        menu.add_command(label="单次执行计划", command=lambda: self.one_time_execution(schedule))
        menu.add_command(label="删除计划", command=lambda: self.delete_schedule(schedule))
        menu.tk_popup(event.x_root, event.y_root)
    
    def one_time_execution(self, schedule):
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
            self
        )
        
        self.config["schedules"].append(one_time_schedule.to_dict())
        self.save_config()
        self.load_schedules()
        one_time_schedule.start()
        
        messagebox.showinfo("单次执行", f"已创建单次执行计划，将在 {time_str} 执行")
    
    def remove_executed_schedule(self, schedule_name):
        self.config["schedules"] = [s for s in self.config["schedules"] if s["name"] != schedule_name]
        self.save_config()
        self.load_schedules()
    
    def create_new_schedule(self):
        dialog = ScheduleDialog(self.root, "新建关机计划", self.icon_path)
        self.center_window(dialog.top, 500, 450)
        self.root.wait_window(dialog.top)
        
        if dialog.result:
            new_schedule = ShutdownSchedule(
                dialog.result.name,
                dialog.result.shutdown_type,
                dialog.result.time,
                dialog.result.days,
                True,
                dialog.result.one_time,
                self
            )
            
            self.config["schedules"].append(new_schedule.to_dict())
            self.save_config()
            self.load_schedules()
            new_schedule.start()
    
    def modify_schedule(self, schedule):
        idx = next((i for i, s in enumerate(self.config["schedules"]) if s["name"] == schedule.name), -1)
        if idx == -1:
            return
            
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
            schedule.name = dialog.result.name
            schedule.shutdown_type = dialog.result.shutdown_type
            schedule.time = dialog.result.time
            schedule.days = dialog.result.days
            schedule.one_time = dialog.result.one_time
            
            self.config["schedules"][idx] = schedule.to_dict()
            self.save_config()
            self.load_schedules()
            
            if schedule.enabled:
                schedule.start()
    
    def delete_schedule(self, schedule):
        self.root.attributes('-topmost', True)
        self.root.update()
        if not messagebox.askyesno("确认删除", f"确定要删除计划 '{schedule.name}' 吗？"):
            self.root.attributes('-topmost', False)
            return
        
        self.root.attributes('-topmost', False)
        schedule.stop()
        self.config["schedules"] = [s for s in self.config["schedules"] if s["name"] != schedule.name]
        self.save_config()
        self.load_schedules()
    
    def show_delete_dialog(self):
        if not self.config.get("schedules", []):
            self.root.attributes('-topmost', True)
            self.root.update()
            messagebox.showinfo("提示", "当前没有可删除的计划")
            self.root.attributes('-topmost', False)
            return
            
        dialog = DeleteDialog(self.root, self.config["schedules"], self.icon_path)
        self.center_window(dialog.top, 550, 550)
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
        SettingsDialog(self.root, self.config, self, self.icon_path)
    
    def start_all_schedules(self):
        logging.info("启动所有计划")
        for schedule_data in self.config.get("schedules", []):
            schedule = ShutdownSchedule.from_dict(schedule_data, self)
            if schedule.enabled:
                logging.info(f"启动计划: {schedule.name}")
                schedule.start()
    
    def stop_all_schedules(self):
        logging.info("停止所有计划")
        for schedule_data in self.config.get("schedules", []):
            schedule = ShutdownSchedule.from_dict(schedule_data, self)
            schedule.stop()
    
    def set_auto_start(self, enable):
        if platform.system() != "Windows":
            return False
            
        app_name = "LazyShutdown"
        app_path = os.path.abspath(sys.argv[0])
        
        if enable and self.config.get("use_task_scheduler", False):
            return self.set_task_scheduler(enable)
        
        return self.set_registry_auto_start(enable)
    
    def set_task_scheduler(self, enable):
        app_name = "LazyShutdown"
        app_path = os.path.abspath(sys.argv[0])
        
        try:
            if enable:
                xml_template = f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>{APP_NAME}</Description>
  </RegistrationInfo>
  <Principals>
    <Principal id="Author">
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
    </LogonTrigger>
  </Triggers>
  <Actions Context="Author">
    <Exec>
      <Command>"{app_path}"</Command>
      <Arguments>--minimized</Arguments>
    </Exec>
  </Actions>
</Task>
"""
                xml_path = os.path.join(tempfile.gettempdir(), 'lazy_shutdown_task.xml')
                with open(xml_path, 'w', encoding='utf-16') as f:
                    f.write(xml_template)
                
                result = subprocess.run(
                    ['schtasks', '/create', '/tn', app_name, '/xml', xml_path, '/f'],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                
                if result.returncode == 0:
                    logging.info("已创建任务计划实现开机自启动")
                    return True
                else:
                    logging.error(f"任务计划创建失败: {result.stderr.decode('gbk')}")
                    return False
            else:
                result = subprocess.run(
                    ['schtasks', '/delete', '/tn', app_name, '/f'],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                
                if result.returncode == 0:
                    logging.info("已删除开机自启动任务")
                    return True
                else:
                    logging.error(f"任务计划删除失败: {result.stderr.decode('gbk')}")
                    return False
        except Exception as e:
            logging.error(f"任务计划操作失败: {e}")
            return False
    
    def set_registry_auto_start(self, enable):
        app_name = "LazyShutdown"
        app_path = os.path.abspath(sys.argv[0])
        
        if enable:
            app_path_with_args = f'"{app_path}" --minimized'
        else:
            app_path_with_args = app_path
        
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTO_START_KEY, 0, winreg.KEY_WRITE)
            
            if enable:
                winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, app_path_with_args)
                logging.info("设置开机自启动（注册表方式）")
            else:
                try:
                    winreg.DeleteValue(key, app_name)
                    logging.info("取消开机自启动（注册表方式）")
                except:
                    pass
            
            winreg.CloseKey(key)
            return True
        except Exception as e:
            error_msg = f"设置自启动失败: {e}"
            logging.error(error_msg)
            return False
    
    def create_tray_icon(self):
        if self.tray_icon:
            return
            
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
                
                threading.Thread(
                    target=self.run_tray_icon, 
                    daemon=True,
                    name="TrayIconThread"
                ).start()
                
                logging.info("最小化到系统托盘")
    
    def run_tray_icon(self):
        try:
            self.tray_icon.run()
        except Exception as e:
            logging.error(f"托盘图标运行出错: {str(e)}")
    
    def on_minimize(self, event):
        if event.widget == self.root and self.config.get("minimize_to_tray", True):
            self.minimize_to_tray()
    
    def quit_app(self, icon=None, item=None):
        try:
            if self.tray_icon:
                self.tray_icon.stop()
        except:
            pass
            
        self.stop_guardian_monitor()
        self.stop_guardian()
        self.stop_all_schedules()
        logging.info("程序退出")
        self.root.destroy()
        sys.exit(0)
    
    def on_close(self):
        if self.config.get("minimize_to_tray", True):
            self.minimize_to_tray()
        else:
            self.quit_app()
    
    def load_config(self):
        try:
            if CONFIG_FILE.exists():
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception as e:
            error_msg = f"加载配置失败: {e}"
            logging.error(error_msg)
        return DEFAULT_CONFIG.copy()
    
    def save_config(self):
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
        
        if icon_path:
            try:
                self.top.iconbitmap(icon_path)
            except:
                pass
        
        main_frame = ttk.Frame(self.top)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        ttk.Label(content_frame, text="计划名称:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        self.name_var = tk.StringVar(value=name)
        name_entry = ttk.Entry(content_frame, textvariable=self.name_var, width=30)
        name_entry.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)
        
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
        
        ttk.Label(content_frame, text="执行时间:").grid(row=2, column=0, sticky=tk.W, pady=5, padx=5)
        self.time_var = tk.StringVar(value=time)
        time_entry = ttk.Entry(content_frame, textvariable=self.time_var, width=30)
        time_entry.grid(row=2, column=1, sticky=tk.W, pady=5, padx=5)
        ttk.Label(content_frame, text="格式: HH:MM (24小时制)").grid(row=3, column=1, sticky=tk.W, padx=5)
        
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
        
        self.one_time_var = tk.BooleanVar(value=False)
        one_time_cb = ttk.Checkbutton(
            content_frame,
            text="单次执行（执行后自动删除）",
            variable=self.one_time_var
        )
        one_time_cb.grid(row=5, column=0, columnspan=2, pady=10, sticky=tk.W, padx=5)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        
        ok_button = ttk.Button(button_frame, text="确定", command=self.on_ok, width=10)
        ok_button.pack(side=tk.RIGHT, padx=10)
        
        cancel_button = ttk.Button(button_frame, text="取消", command=self.top.destroy, width=10)
        cancel_button.pack(side=tk.RIGHT)
    
    def on_ok(self):
        name = self.name_var.get().strip()
        shutdown_type = self.type_var.get()
        time_str = self.time_var.get().strip()
        days = [i+1 for i, var in enumerate(self.day_vars) if var.get()]
        one_time = self.one_time_var.get()
        
        if not name:
            self.top.attributes('-topmost', True)
            self.top.update()
            messagebox.showerror("错误", "请输入计划名称", parent=self.top)
            self.top.attributes('-topmost', False)
            return
            
        if not shutdown_type:
            self.top.attributes('-topmost', True)
            self.top.update()
            messagebox.showerror("错误", "请选择操作类型", parent=self.top)
            self.top.attributes('-topmost', False)
            return
            
        if not time_str or not self.validate_time(time_str):
            self.top.attributes('-topmost', True)
            self.top.update()
            messagebox.showerror("错误", "请输入有效的时间 (HH:MM)", parent=self.top)
            self.top.attributes('-topmost', False)
            return
            
        if not days and not one_time:
            self.top.attributes('-topmost', True)
            self.top.update()
            messagebox.showerror("错误", "请至少选择一个日期或选择单次执行", parent=self.top)
            self.top.attributes('-topmost', False)
            return
            
        self.result = ShutdownSchedule(name, shutdown_type, time_str, days, True, one_time)
        self.top.destroy()
    
    def validate_time(self, time_str):
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
        self.top.geometry("550x550")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()
        self.top.attributes('-topmost', True)
        
        if icon_path:
            try:
                self.top.iconbitmap(icon_path)
            except:
                pass
        
        main_frame = ttk.Frame(self.top)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        ttk.Label(content_frame, text="选择要删除的计划:", font=("微软雅黑", 11, "bold")).pack(fill=tk.X, pady=(0, 15))
        
        list_container = ttk.Frame(content_frame)
        list_container.pack(fill=tk.BOTH, expand=True)
        
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
        
        for schedule in schedules:
            self.listbox.insert(tk.END, schedule["name"])
        
        select_all_button = ttk.Button(
            content_frame,
            text="全选",
            command=self.select_all,
            width=12
        )
        select_all_button.pack(anchor=tk.W, pady=10)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(20, 10))
        
        ok_button = ttk.Button(
            button_frame,
            text="确定",
            command=self.on_ok,
            width=15
        )
        ok_button.pack(side=tk.RIGHT, padx=10)
        
        cancel_button = ttk.Button(
            button_frame,
            text="取消",
            command=self.top.destroy,
            width=15
        )
        cancel_button.pack(side=tk.RIGHT)
    
    def select_all(self):
        self.listbox.selection_set(0, tk.END)
    
    def on_ok(self):
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
        
        self.center_window(self.top, 500, 550)
        
        if icon_path:
            try:
                self.top.iconbitmap(icon_path)
            except:
                pass
        
        # 主容器
        main_container = ttk.Frame(self.top)
        main_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 滚动区域容器
        scroll_container = ttk.Frame(main_container)
        scroll_container.pack(fill=tk.BOTH, expand=True)
        
        # 创建画布和滚动条
        self.canvas = tk.Canvas(scroll_container)
        scrollbar = ttk.Scrollbar(scroll_container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
)
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        # 设置快速滚动
        self.canvas.configure(yscrollincrement=30)
        
        # 绑定鼠标滚轮事件
        self.scrollable_frame.bind("<Enter>", self._bound_to_mousewheel)
        self.scrollable_frame.bind("<Leave>", self._unbound_to_mousewheel)
        
        # 布局
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 创建标签页
        notebook = ttk.Notebook(self.scrollable_frame)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 常规标签页
        general_frame = ttk.Frame(notebook)
        notebook.add(general_frame, text="常规")
        self.create_general_settings(general_frame)
        
        # 安全标签页
        security_frame = ttk.Frame(notebook)
        notebook.add(security_frame, text="安全")
        self.create_security_settings(security_frame)
        
        # 关于标签页
        about_frame = ttk.Frame(notebook)
        notebook.add(about_frame, text="关于")
        self.create_about_settings(about_frame)
        
        # 按钮框架 - 固定在底部
        button_frame = ttk.Frame(main_container)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 0))
        
        # 确定按钮
        ok_button = ttk.Button(button_frame, text="确定", command=self.on_ok, width=10)
        ok_button.pack(side=tk.RIGHT, padx=10)
        
        # 取消按钮
        cancel_button = ttk.Button(button_frame, text="取消", command=self.top.destroy, width=10)
        cancel_button.pack(side=tk.RIGHT)
    
    def _bound_to_mousewheel(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
    
    def _unbound_to_mousewheel(self, event):
        self.canvas.unbind_all("<MouseWheel>")
    
    def _on_mousewheel(self, event):
        # 快速滚动（每次滚动3倍距离）
        self.canvas.yview_scroll(int(-3*(event.delta/120)), "units")
        
    def center_window(self, window, width, height):
        window.update_idletasks()
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2) - 30
        
        window.geometry(f"{width}x{height}+{x}+{y}")
    
    def create_general_settings(self, parent):
        auto_start_frame = ttk.LabelFrame(parent, text="开机自启动")
        auto_start_frame.pack(fill=tk.X, pady=10, padx=5)
        
        self.startup_var = tk.StringVar()
        
        registry_rb = ttk.Radiobutton(
            auto_start_frame, 
            text="注册表自启动",
            variable=self.startup_var,
            value="registry"
        )
        registry_rb.pack(anchor=tk.W, padx=10, pady=5)
        
        task_rb = ttk.Radiobutton(
            auto_start_frame, 
            text="任务计划自启动（避免UAC弹窗）",
            variable=self.startup_var,
            value="task"
        )
        task_rb.pack(anchor=tk.W, padx=10, pady=5)
        
        none_rb = ttk.Radiobutton(
            auto_start_frame, 
            text="无自启动",
            variable=self.startup_var,
            value="none"
        )
        none_rb.pack(anchor=tk.W, padx=10, pady=5)
        
        if self.config.get("use_task_scheduler", False):
            self.startup_var.set("task")
        elif self.config.get("auto_start", False):
            self.startup_var.set("registry")
        else:
            self.startup_var.set("none")
        
        task_note = ttk.Label(
            auto_start_frame, 
            text="* 使用任务计划启动可避免UAC弹窗，但需要管理员权限设置\n"
                 "* 选择新方式时，旧的自启动方式会自动清理",
            font=("微软雅黑", 8),
            justify=tk.LEFT
        )
        task_note.pack(anchor=tk.W, padx=20, pady=(0, 5))
        
        tray_frame = ttk.LabelFrame(parent, text="界面设置")
        tray_frame.pack(fill=tk.X, pady=10, padx=5)
        
        self.tray_var = tk.BooleanVar(value=self.config.get("minimize_to_tray", True))
        tray_cb = ttk.Checkbutton(
            tray_frame, 
            text="最小化到系统托盘",
            variable=self.tray_var
        )
        tray_cb.pack(anchor=tk.W, padx=10, pady=5)
        
        self.hide_tray_var = tk.BooleanVar(value=self.config.get("hide_tray_icon", False))
        hide_tray_cb = ttk.Checkbutton(
            tray_frame, 
            text="隐藏托盘图标（需设置热键）",
            variable=self.hide_tray_var
        )
        hide_tray_cb.pack(anchor=tk.W, padx=10, pady=5)
        
        hotkey_frame = ttk.Frame(tray_frame)
        hotkey_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(hotkey_frame, text="唤醒热键:").pack(side=tk.LEFT, padx=(0, 10))
        
        self.hotkey_var = tk.StringVar(value=self.config.get("hotkey", "ctrl+alt+l"))
        hotkey_entry = ttk.Entry(hotkey_frame, textvariable=self.hotkey_var, width=20)
        hotkey_entry.pack(side=tk.LEFT)
        ttk.Label(hotkey_frame, text="(例如: ctrl+alt+l)").pack(side=tk.LEFT, padx=(5, 0))
        
        perm_frame = ttk.LabelFrame(parent, text="权限设置")
        perm_frame.pack(fill=tk.X, pady=10, padx=5)
        
        self.admin_var = tk.BooleanVar(value=self.config.get("run_as_admin", True))
        admin_cb = ttk.Checkbutton(
            perm_frame, 
            text="以管理员权限运行（推荐）",
            variable=self.admin_var
        )
        admin_cb.pack(anchor=tk.W, padx=10, pady=5)
        
        uac_note = ttk.Label(
            perm_frame, 
            text="* UAC（用户账户控制）是Windows的安全功能，会弹窗询问权限\n"
                 "* 使用任务计划启动可以避免UAC弹窗，但需要管理员权限设置",
            font=("微软雅黑", 8),
            justify=tk.LEFT
        )
        uac_note.pack(anchor=tk.W, padx=20, pady=(0, 5))
    
    def create_security_settings(self, parent):
        security_frame = ttk.LabelFrame(parent, text="安全设置")
        security_frame.pack(fill=tk.BOTH, expand=True, pady=10, padx=5)
        
        self.guardian_enabled_var = tk.BooleanVar(
            value=self.config.get("guardian_enabled", False)
        )
        guardian_enabled_cb = ttk.Checkbutton(
            security_frame, 
            text="启用守护进程",
            variable=self.guardian_enabled_var,
            command=self.toggle_guardian_settings
        )
        guardian_enabled_cb.pack(anchor=tk.W, padx=10, pady=5)
        
        self.guardian_settings_frame = ttk.Frame(security_frame)
        self.guardian_settings_frame.pack(fill=tk.X, padx=20, pady=(0, 10))
        
        self.guardian_autostart_var = tk.BooleanVar(
            value=self.config.get("guardian_autostart", False)
        )
        autostart_cb = ttk.Checkbutton(
            self.guardian_settings_frame, 
            text="开机自启动守护进程",
            variable=self.guardian_autostart_var
        )
        autostart_cb.pack(anchor=tk.W, pady=2)
        
        self.guardian_autorestart_var = tk.BooleanVar(
            value=self.config.get("guardian_autorestart", True)
        )
        autorestart_cb = ttk.Checkbutton(
            self.guardian_settings_frame, 
            text="自动重启主程序",
            variable=self.guardian_autorestart_var
        )
        autorestart_cb.pack(anchor=tk.W, pady=2)
        
        self.guardian_terminate_taskmgr_var = tk.BooleanVar(
            value=self.config.get("guardian_terminate_taskmgr", True)
        )
        terminate_cb = ttk.Checkbutton(
            self.guardian_settings_frame, 
            text="终止任务管理器进程",
            variable=self.guardian_terminate_taskmgr_var
        )
        terminate_cb.pack(anchor=tk.W, pady=2)
        
        self.guardian_hide_window_var = tk.BooleanVar(
            value=self.config.get("guardian_hide_window", True)
        )
        hide_window_cb = ttk.Checkbutton(
            self.guardian_settings_frame, 
            text="隐藏守护进程窗口",
            variable=self.guardian_hide_window_var
        )
        hide_window_cb.pack(anchor=tk.W, pady=2)
        
        # 新增：控制台显示选项
        self.guardian_show_console_var = tk.BooleanVar(
            value=self.config.get("guardian_show_console", False)
        )
        show_console_cb = ttk.Checkbutton(
            self.guardian_settings_frame, 
            text="显示守护进程控制台（类似CMD）",
            variable=self.guardian_show_console_var
        )
        show_console_cb.pack(anchor=tk.W, pady=2)
        
        self.toggle_guardian_settings()
    
    def create_about_settings(self, parent):
        about_frame = ttk.LabelFrame(parent, text="关于")
        about_frame.pack(fill=tk.BOTH, expand=True, pady=10, padx=5)
        
        info_lines = [
            f"{APP_NAME} v1.0",
            "",
            "开发者: Jumao",
            "联系邮箱: jumaozhixing@outlook.com",
            "",
            "GitHub项目地址:"
        ]
        
        for line in info_lines:
            ttk.Label(about_frame, text=line, justify=tk.LEFT, font=("微软雅黑", 9)).pack(anchor=tk.W, padx=10, pady=2)
        
        github_frame = ttk.Frame(about_frame)
        github_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        github_label = ttk.Label(
            github_frame, 
            text=GITHUB_URL, 
            foreground="blue", 
            cursor="hand2",
            font=("微软雅黑", 9, "underline")
        )
        github_label.pack(side=tk.LEFT)
        
        copy_button = ttk.Button(
            github_frame,
            text="复制",
            width=5,
            command=lambda: self.copy_to_clipboard(GITHUB_URL)
        )
        copy_button.pack(side=tk.RIGHT, padx=(10, 0))
        
        github_label.bind("<Button-1>", lambda e: webbrowser.open(GITHUB_URL))
    
    def toggle_guardian_settings(self):
        state = "normal" if self.guardian_enabled_var.get() else "disabled"
        for child in self.guardian_settings_frame.winfo_children():
            if isinstance(child, ttk.Checkbutton):
                child.state(["!disabled" if state == "normal" else "disabled"])

    def copy_to_clipboard(self, text):
        self.top.clipboard_clear()
        self.top.clipboard_append(text)
        messagebox.showinfo("已复制", "链接已复制到剪贴板")
    
    def on_ok(self):
        current_option = self.startup_var.get()
        
        if current_option != "registry":
            self.app.set_registry_auto_start(False)
        
        if current_option != "task":
            self.app.set_task_scheduler(False)
        
        if current_option == "registry":
            self.app.set_registry_auto_start(True)
        elif current_option == "task":
            self.app.set_task_scheduler(True)
        
        self.config["auto_start"] = (current_option == "registry")
        self.config["use_task_scheduler"] = (current_option == "task")
        self.config["minimize_to_tray"] = self.tray_var.get()
        self.config["hide_tray_icon"] = self.hide_tray_var.get()
        self.config["hotkey"] = self.hotkey_var.get().strip()
        self.config["run_as_admin"] = self.admin_var.get()
        
        # 更新守护进程配置
        self.config["guardian_enabled"] = self.guardian_enabled_var.get()
        self.config["guardian_autostart"] = self.guardian_autostart_var.get()
        self.config["guardian_autorestart"] = self.guardian_autorestart_var.get()
        self.config["guardian_terminate_taskmgr"] = self.guardian_terminate_taskmgr_var.get()
        self.config["guardian_hide_window"] = self.guardian_hide_window_var.get()
        self.config["guardian_show_console"] = self.guardian_show_console_var.get()  # 新增控制台显示选项
        
        self.app.setup_hotkey()
        self.app.setup_guardian()
        self.app.save_config()
        
        if self.config["run_as_admin"]:
            self.app.check_admin_privileges()
        
        self.top.destroy()

def is_admin():
    if platform.system() == "Windows":
        try:
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except:
            return False
    return False

def main():
    if platform.system() == "Windows":
        mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "LazyShutdownMutex")
        if ctypes.windll.kernel32.GetLastError() == 183:
            ctypes.windll.user32.MessageBoxW(0, "程序已在运行中", APP_NAME, 0)
            return
    
    start_minimized = "--minimized" in sys.argv
    
    run_as_admin = True
    try:
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
                run_as_admin = config.get("run_as_admin", True)
    except:
        pass
    
    if run_as_admin and not is_admin() and platform.system() == "Windows":
        try:
            exe_path = sys.executable if getattr(sys, 'frozen', False) else sys.argv[0]
            ShellExecute(0, "runas", exe_path, " ".join(sys.argv[1:]), None, 1)
            sys.exit(0)
        except Exception as e:
            log_file = CONFIG_DIR / "lazy_shutdown.log"
            with open(log_file, "a", encoding="utf-8") as log:
                log.write(f"[{datetime.datetime.now()}] 请求管理员权限失败: {str(e)}\n")
    
    root = tk.Tk()
    root.withdraw()
    
    try:
        if platform.system() == "Windows":
            base_path = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(base_path, "icon.ico")
            
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
                icon_path = os.path.join(base_path, "icon.ico")
                
            root.iconbitmap(icon_path)
        else:
            icon_path = None
    except Exception as e:
        log_file = CONFIG_DIR / "lazy_shutdown.log"
        with open(log_file, "a", encoding="utf-8") as log:
            log.write(f"[{datetime.datetime.now()}] 设置图标失败: {str(e)}\n")
        icon_path = None
    
    root.icon_path = icon_path
    
    app = LazyShutdownApp(root, icon_path)
    
    if start_minimized:
        app.minimize_to_tray()
    else:
        root.deiconify()
        
    try:
        root.mainloop()
    except KeyboardInterrupt:
        log_file = CONFIG_DIR / "lazy_shutdown.log"
        with open(log_file, "a", encoding="utf-8") as log:
            log.write(f"[{datetime.datetime.now()}] 程序被用户中断\n")
        app.quit_app()

if __name__ == "__main__":
    main()