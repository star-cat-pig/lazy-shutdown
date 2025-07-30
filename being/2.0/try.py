import os, sys, time, psutil, subprocess, ctypes, json, logging, argparse
from datetime import datetime
import tkinter as tk
from threading import Thread

APP_NAME = "懒人关机器"
MAIN_EXE = "懒人关机器.exe"
CHECK_INTERVAL = 5
TASK_MANAGERS = ["taskmgr.exe", "processhacker.exe", "procexp.exe", "procexp64.exe"]

def is_guardian_running():
    """检查是否已有守护进程实例在运行（排除自身）"""
    current_pid = os.getpid()
    current_exe = sys.executable.lower()
    
    for proc in psutil.process_iter(['pid', 'name', 'exe', 'cmdline']):
        try:
            # 跳过自身进程
            if proc.pid == current_pid:
                continue
                
            # 获取进程的可执行文件路径
            proc_exe = (proc.info.get('exe') or '').lower()
            
            # 如果进程的可执行文件路径与当前进程相同，跳过（可能是同一个程序的不同实例）
            if proc_exe == current_exe:
                continue
                
            # 检查进程名和命令行
            name = (proc.info.get('name') or '').lower()
            cmdline = proc.info.get('cmdline') or []
            
            # 判断是否是守护进程
            is_guardian = False
            if "guardian" in name or "懒人关机器_守护进程" in name:
                is_guardian = True
            elif any("guardian.py" in arg for arg in cmdline) or any("guardian.exe" in arg for arg in cmdline):
                is_guardian = True
            
            if is_guardian:
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue
            
    return False

# Windows API
if sys.platform == "win32":
    SW_HIDE, SW_SHOW = 0, 5
    ShellExecuteW = ctypes.windll.shell32.ShellExecuteW
    ShellExecuteW.argtypes = (ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_wchar_p,
                              ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_int)
    
    # 控制台相关API
    GetConsoleWindow = ctypes.windll.kernel32.GetConsoleWindow
    AllocConsole = ctypes.windll.kernel32.AllocConsole
    FreeConsole = ctypes.windll.kernel32.FreeConsole
    SetConsoleTitleW = ctypes.windll.kernel32.SetConsoleTitleW
    GetStdHandle = ctypes.windll.kernel32.GetStdHandle
    SetStdHandle = ctypes.windll.kernel32.SetStdHandle
    STD_OUTPUT_HANDLE = -11
    STD_ERROR_HANDLE = -12

def load_config():
    cfg = {
        "terminate_taskmgr": True, 
        "autorestart": True, 
        "hide_window": True, 
        "show_window": False,
        "show_console": False  # 控制台显示选项
    }
    path = os.path.join(os.getenv('APPDATA'), "LazyShutdown", "lazy_shutdown_config.json")
    try:
        with open(path, encoding='utf-8') as f:
            j = json.load(f)
            cfg.update({
                "terminate_taskmgr": j.get("guardian_terminate_taskmgr", True),
                "autorestart":      j.get("guardian_autorestart", True),
                "hide_window":      j.get("guardian_hide_window", True),
                "show_window": j.get("guardian_show_window", False),
                "show_console": j.get("guardian_show_console", False)  # 读取控制台配置
            })
    except:
        pass
    return cfg

def setup_logging(hide_window, show_window, show_console):
    log_dir = os.path.join(os.getenv('APPDATA'), "LazyShutdown", "Logs")
    os.makedirs(log_dir, exist_ok=True)
    fh = logging.FileHandler(os.path.join(log_dir, "Guardian.log"), encoding='utf-8')
    fh.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'))
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.addHandler(fh)
    
    # 如果配置要求显示窗口或控制台，添加控制台日志处理器
    if show_window or show_console:
        ch = logging.StreamHandler()
        ch.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'))
        root.addHandler(ch)

class GuardianWindow:
    def __init__(self, title):
        self.root = tk.Tk()
        self.root.title(title)
        self.root.geometry("800x400")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # 创建文本区域和滚动条
        frame = tk.Frame(self.root)
        frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.text_area = tk.Text(frame, wrap=tk.WORD, yscrollcommand=scrollbar.set)
        self.text_area.pack(fill=tk.BOTH, expand=True)
        self.text_area.config(state=tk.DISABLED, font=("Consolas", 10))
        
        scrollbar.config(command=self.text_area.yview)
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("守护进程正在运行...")
        status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 自定义日志处理器
        self.log_handler = CustomLogHandler(self)
        self.log_handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'))
        logging.getLogger().addHandler(self.log_handler)
        
        # 设置窗口图标（如果可用）
        try:
            if hasattr(sys, '_MEIPASS'):
                icon_path = os.path.join(sys._MEIPASS, "icon.ico")
                self.root.iconbitmap(icon_path)
        except:
            pass

    def on_close(self):
        # 最小化而不是关闭
        self.root.iconify()

    def add_log(self, record):
        self.text_area.config(state=tk.NORMAL)
        
        # 根据日志级别设置颜色
        if record.levelno >= logging.ERROR:
            tag = "error"
            color = "red"
        elif record.levelno >= logging.WARNING:
            tag = "warning"
            color = "orange"
        else:
            tag = "info"
            color = "black"
        
        # 添加带颜色的文本
        self.text_area.insert(tk.END, self.log_handler.format(record) + "\n", tag)
        self.text_area.see(tk.END)
        self.text_area.config(state=tk.DISABLED)
        
        # 配置标签颜色
        self.text_area.tag_config(tag, foreground=color)

    def update_status(self, message):
        self.status_var.set(message)

    def run(self):
        self.root.mainloop()

class CustomLogHandler(logging.Handler):
    def __init__(self, window):
        super().__init__()
        self.window = window
        self.setLevel(logging.INFO)
        self.formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
    
    def emit(self, record):
        try:
            self.window.add_log(record)
        except Exception:
            self.handleError(record)
    
    def format(self, record):
        return self.formatter.format(record)

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    if not is_admin() and sys.platform == "win32":
        exe = sys.executable if getattr(sys, 'frozen', False) else sys.argv[0]
        args = " ".join(arg for arg in sys.argv[1:] if arg != "--minimized")
        ShellExecuteW(None, "runas", exe, args, None, SW_SHOW)
        sys.exit(0)

def find_main():
    for p in psutil.process_iter(['pid','name','exe']):
        name = p.info['name'] or ''
        exe  = os.path.basename(p.info.get('exe') or '')
        if name.lower() == MAIN_EXE.lower() or exe.lower() == MAIN_EXE.lower():
            return p
    return None

def start_main():
    base = getattr(sys, 'frozen', False) and os.path.dirname(sys.executable) or os.path.dirname(__file__)
    path = os.path.join(base, MAIN_EXE)
    if not os.path.exists(path):
        logging.error("找不到主程序: %s", path)
        return None
    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = SW_HIDE
    return subprocess.Popen([path, "--minimized"], startupinfo=si, creationflags=subprocess.CREATE_NO_WINDOW)

def kill_taskmgr():
    for p in psutil.process_iter(['name']):
        nm = (p.info['name'] or '').lower()
        if any(tm == nm for tm in TASK_MANAGERS):
            try:
                p.kill()
                logging.info("终止任务管理器: %s", nm)
            except Exception as e:
                logging.warning("无法终止任务管理器 %s: %s", nm, str(e))

def create_console():
    """创建控制台窗口并重定向输出"""
    if sys.platform != "win32":
        return
    
    # 检查是否已有控制台
    if GetConsoleWindow() == 0:
        # 没有控制台，创建一个新的
        AllocConsole()
    
    # 设置控制台标题
    SetConsoleTitleW(f"{APP_NAME} - 守护进程控制台")
    
    # 重定向标准输出
    sys.stdout = open('CONOUT$', 'w', encoding='utf-8')
    sys.stderr = open('CONOUT$', 'w', encoding='utf-8')
    
    # 打印控制台标题
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"[{now}] {APP_NAME} 守护进程控制台已启动")
    print("=" * 70)
    print(f"守护进程状态监控 | 日志级别: INFO")
    print(f"按 Ctrl+C 安全退出守护进程")
    print("=" * 70)

def close_console():
    """关闭控制台窗口"""
    if sys.platform == "win32" and GetConsoleWindow() != 0:
        # 恢复标准输出
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__
        
        # 关闭控制台
        FreeConsole()

def main():
    # 检查是否已有实例运行（排除自身）
    if is_guardian_running():
        logging.info("检测到已有守护进程实例运行，本实例将退出")
        # 静默退出，不显示任何提示
        sys.exit(0)
    
    # 加载配置
    cfg = load_config()
    setup_logging(cfg["hide_window"], cfg["show_window"], cfg["show_console"])
    
    # 解析命令行参数
    parser = argparse.ArgumentParser(description=f"{APP_NAME} 守护进程")
    parser.add_argument("--minimized", action="store_true", help="以最小化模式启动")
    parser.add_argument("--console", action="store_true", help="强制显示控制台窗口")
    parser.add_argument("--debug", action="store_true", help="启用调试日志")
    args = parser.parse_args()
    
    # 如果指定了--console参数，覆盖配置
    if args.console:
        cfg["show_console"] = True
    
    # 如果配置要求显示控制台，创建控制台
    if cfg["show_console"]:
        create_console()
    
    # 提升管理员权限
    run_as_admin()
    
    # 设置日志级别
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
        logging.debug("调试模式已启用")
    
    logging.info("守护进程启动, 配置: %s", cfg)
    
    # 如果配置要求显示窗口，创建并运行窗口
    window = None
    if cfg["show_window"]:
        window = GuardianWindow(f"{APP_NAME} - 守护进程")
        # 在新线程中运行窗口
        window_thread = Thread(target=window.run, daemon=True)
        window_thread.start()
        logging.info("守护进程窗口已启动")
    
    last_config_check = time.time()
    
    try:
        while True:
            try:
                # 每分钟检查一次配置是否更新
                if time.time() - last_config_check > 60:
                    new_cfg = load_config()
                    if new_cfg != cfg:
                        logging.info("检测到配置更新，重新加载配置")
                        cfg = new_cfg
                        setup_logging(cfg["hide_window"], cfg["show_window"], cfg["show_console"])
                    last_config_check = time.time()
                
                # 检查主程序是否运行
                main_process = find_main()
                if not main_process and cfg["autorestart"]:
                    logging.warning("主程序未运行，正在启动...")
                    start_main()
                elif main_process:
                    logging.debug("主程序运行中: PID=%d", main_process.pid)

                # 检查并终止任务管理器
                if cfg["terminate_taskmgr"]:
                    kill_taskmgr()
                
                # 等待下一次检查
                time.sleep(CHECK_INTERVAL)
            except Exception as e:
                logging.exception("主循环错误: %s, 10秒后重试", str(e))
                time.sleep(10)
    except KeyboardInterrupt:
        logging.info("用户中断守护进程")
        print("\n[安全退出] 正在关闭守护进程...")
    except Exception as e:
        logging.exception("严重错误导致守护进程退出: %s", str(e))
        print(f"\n[错误] 严重错误: {str(e)}")
    finally:
        # 清理资源
        close_console()
        logging.info("守护进程已退出")
        print("守护进程已安全退出")

if __name__ == "__main__":
    main()