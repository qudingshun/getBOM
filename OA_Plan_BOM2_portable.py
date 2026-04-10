
# 在文件顶部添加这行导入
try:
    from webdriver_manager.chrome import ChromeDriverManager
    WEBDRIVER_MANAGER_AVAILABLE = True
except ImportError:
    WEBDRIVER_MANAGER_AVAILABLE = False

import subprocess  # 添加这个导入用于清理进程

def kill_chrome_processes():
    """清理残留的Chrome和ChromeDriver进程"""
    try:
        subprocess.call(['taskkill', '/F', '/IM', 'chrome.exe', '/T'],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, shell=True)
        subprocess.call(['taskkill', '/F', '/IM', 'chromedriver.exe', '/T'],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, shell=True)
        pause("after_kill_processes")
        print("已清理残留进程")
    except:
        pass
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver as ChromeWebDriver  # noqa: F401
import pandas as pd
import time
import os
import glob
import json
import getpass
import base64
import hashlib
from io import StringIO
import sys
import urllib3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# 禁用SSL警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

try:
    import pyautogui

    PYAUTOGUI_AVAILABLE = True
except ImportError:
    PYAUTOGUI_AVAILABLE = False

try:
    import pyperclip

    PYPERCLIP_AVAILABLE = True
except ImportError:
    PYPERCLIP_AVAILABLE = False


# ========== 一处修改的统一配置区 ==========
RUN_CONFIG = {
    # OA查询配置
    "oa_year": "2026",
    "oa_months": ["05", "04"],  # 支持多个，格式两位字符串：01~12
    "oa_factory_code": "M03",   # OA页面工厂编码，如 M03

    # 帆软BOM查询配置
    "fanruan_factory_code": "664",  # 帆软参数P_ORGID对应编码
    "material_excel_path": "",      # 可选：本地物料编码Excel，设置后跳过OA

    # Cookie文件（可选）：放在脚本同目录
    # 文件内容格式：浏览器导出的cookie JSON数组
    "oa_cookie_file": "oa_cookies.json",
    "fanruan_cookie_file": "fanruan_cookies.json",
}

RUNTIME_CONFIG = {}
RUNTIME_CONFIG_FILE = "runtime_config.json"
YEAR_OPTIONS = [str(y) for y in range(2024, 2100)]
MONTH_OPTIONS = [f"{m:02d}" for m in range(1, 13)]
OA_FACTORY_OPTIONS = [f"M0{i}" for i in range(2, 9)]  # M02~M08
FANRUAN_FACTORY_OPTIONS = [str(i) for i in range(662, 670)]  # 662~669

# 各业务节拍（秒）
SLEEP_DEFAULTS = {
    "after_kill_processes": 1.0,
    "after_cookie_apply": 0.5,
    "after_login_submit": 0.5,
    "after_oa_enter_plan_page": 0.5,
    "after_oa_set_factory": 0.3,
    "after_oa_set_year": 0.2,
    "after_oa_set_month": 0.2,
    "after_oa_submit_query": 0.3,
    "after_open_fanruan_page": 0.3,
    "before_check_fanruan_login": 0.3,
    "after_fanruan_cookie_apply": 0.3,
    "after_fanruan_login": 0.5,
    "per_material_loop_start": 0.3,
    "after_set_fanruan_factory": 0.3,
    "after_fill_material_code": 0.2,
    "after_click_search": 0.3,
    "wait_download_poll": 0.3,
}
SLEEP_CONFIG = dict(SLEEP_DEFAULTS)
SLEEP_OPTIONS = [f"{x/10:.1f}" for x in range(1, 21)]  # 0.1~2.0


def pause(name):
    """按命名节拍暂停。"""
    time.sleep(float(SLEEP_CONFIG.get(name, 0.3)))


def _get_local_key():
    """基于本机信息生成本地转译密钥。"""
    seed = f"{os.getenv('USERNAME', '')}|{os.getenv('COMPUTERNAME', '')}|BOM_CFG_V1"
    return hashlib.sha256(seed.encode("utf-8")).digest()


def encode_local_password(plain_text):
    """将密码转译后再存储（非明文，非强加密）。"""
    if not plain_text:
        return ""
    key = _get_local_key()
    raw = plain_text.encode("utf-8")
    xored = bytes([b ^ key[i % len(key)] for i, b in enumerate(raw)])
    return "enc1:" + base64.urlsafe_b64encode(xored).decode("ascii")


def decode_local_password(encoded_text):
    """读取本地转译密码并还原。兼容历史明文。"""
    if not encoded_text:
        return ""
    if not str(encoded_text).startswith("enc1:"):
        return str(encoded_text)
    try:
        payload = str(encoded_text)[5:]
        key = _get_local_key()
        raw = base64.urlsafe_b64decode(payload.encode("ascii"))
        plain = bytes([b ^ key[i % len(key)] for i, b in enumerate(raw)])
        return plain.decode("utf-8")
    except Exception:
        return ""


def get_runtime_config_path():
    return os.path.join(get_project_dir(), RUNTIME_CONFIG_FILE)


def load_persisted_runtime_config():
    path = get_runtime_config_path()
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return {}
        # 自动还原转译密码；兼容历史明文
        if "oa_password" in data:
            data["oa_password"] = decode_local_password(data.get("oa_password", ""))
        if "fanruan_password" in data:
            data["fanruan_password"] = decode_local_password(data.get("fanruan_password", ""))
        return data
    except Exception:
        return {}


def save_persisted_runtime_config(cfg):
    path = get_runtime_config_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        print(f"已保存运行配置: {path}")
    except Exception as e:
        print(f"保存运行配置失败: {e}")


def get_credentials(prefix):
    """从环境变量读取账号密码，避免明文写入代码。"""
    # 启动弹窗输入优先
    if prefix == "OA":
        username = RUNTIME_CONFIG.get("oa_username", "").strip()
        password = RUNTIME_CONFIG.get("oa_password", "")
        if username and password:
            return username, password
    if prefix == "FANRUAN":
        username = RUNTIME_CONFIG.get("fanruan_username", "").strip()
        password = RUNTIME_CONFIG.get("fanruan_password", "")
        if username and password:
            return username, password

    username = os.getenv(f"{prefix}_USERNAME", "")
    password = os.getenv(f"{prefix}_PASSWORD", "")
    return username, password


def prompt_manual_credentials(system_name):
    """Cookie失败后手工输入账号密码。"""
    print(f"\n{system_name} 需要登录，请手工输入账号密码：")
    username = input(f"{system_name} 用户名: ").strip()
    password = getpass.getpass(f"{system_name} 密码(输入不回显): ").strip()
    return username, password


def save_cookies_to_file(driver, cookie_file):
    """将当前会话Cookie保存到JSON文件。"""
    try:
        cookies = driver.get_cookies()
        with open(cookie_file, "w", encoding="utf-8") as f:
            json.dump(cookies, f, ensure_ascii=False, indent=2)
        print(f"已更新Cookie文件: {cookie_file}（{len(cookies)}条）")
    except Exception as e:
        print(f"保存Cookie失败: {e}")


def show_runtime_config_dialog():
    """运行前弹窗配置：基础配置+节拍配置（分页）。"""
    result = {}
    persisted = load_persisted_runtime_config()
    root = tk.Tk()
    root.title("BOM导出参数设置")
    root.geometry("640x780")
    root.resizable(False, False)

    outer = ttk.Frame(root, padding=10)
    outer.pack(fill="both", expand=True)

    notebook = ttk.Notebook(outer)
    notebook.pack(fill="both", expand=True)

    base_tab = ttk.Frame(notebook, padding=10)
    pace_tab = ttk.Frame(notebook, padding=10)
    notebook.add(base_tab, text="基础配置")
    notebook.add(pace_tab, text="节拍配置")

    ttk.Label(base_tab, text="OA 用户名").grid(row=0, column=0, sticky="w")
    oa_user = ttk.Entry(base_tab, width=40)
    oa_user.grid(row=0, column=1, pady=4, sticky="w")
    oa_user.insert(0, persisted.get("oa_username", ""))

    ttk.Label(base_tab, text="OA 密码").grid(row=1, column=0, sticky="w")
    oa_pwd = ttk.Entry(base_tab, width=40, show="*")
    oa_pwd.grid(row=1, column=1, pady=4, sticky="w")
    oa_pwd.insert(0, persisted.get("oa_password", ""))

    ttk.Label(base_tab, text="帆软 用户名").grid(row=2, column=0, sticky="w")
    fr_user = ttk.Entry(base_tab, width=40)
    fr_user.grid(row=2, column=1, pady=4, sticky="w")
    fr_user.insert(0, persisted.get("fanruan_username", ""))

    ttk.Label(base_tab, text="帆软 密码").grid(row=3, column=0, sticky="w")
    fr_pwd = ttk.Entry(base_tab, width=40, show="*")
    fr_pwd.grid(row=3, column=1, pady=4, sticky="w")
    fr_pwd.insert(0, persisted.get("fanruan_password", ""))

    ttk.Label(base_tab, text="年份").grid(row=4, column=0, sticky="w")
    year_var = tk.StringVar(value=str(persisted.get("oa_year", RUN_CONFIG["oa_year"])))
    year_combo = ttk.Combobox(base_tab, width=38, textvariable=year_var, values=YEAR_OPTIONS, state="readonly")
    year_combo.grid(row=4, column=1, pady=4, sticky="w")
    if year_var.get() in YEAR_OPTIONS:
        year_combo.current(YEAR_OPTIONS.index(year_var.get()))
    else:
        year_combo.current(0)

    ttk.Label(base_tab, text="月份（单格输入，逗号分隔，如 03,04）").grid(row=5, column=0, sticky="w")
    months_var = tk.StringVar(value=",".join(persisted.get("oa_months", RUN_CONFIG["oa_months"])))
    months_entry = ttk.Entry(base_tab, width=40, textvariable=months_var)
    months_entry.grid(row=5, column=1, pady=4, sticky="w")

    ttk.Label(base_tab, text="OA工厂编码").grid(row=6, column=0, sticky="w")
    oa_factory_var = tk.StringVar(value=persisted.get("oa_factory_code", RUN_CONFIG["oa_factory_code"]))
    oa_factory_combo = ttk.Combobox(
        base_tab, width=38, textvariable=oa_factory_var, values=OA_FACTORY_OPTIONS, state="readonly"
    )
    oa_factory_combo.grid(row=6, column=1, pady=4, sticky="w")
    if oa_factory_var.get() in OA_FACTORY_OPTIONS:
        oa_factory_combo.current(OA_FACTORY_OPTIONS.index(oa_factory_var.get()))
    else:
        oa_factory_combo.current(0)

    ttk.Label(base_tab, text="帆软工厂编码").grid(row=7, column=0, sticky="w")
    fr_factory_var = tk.StringVar(value=persisted.get("fanruan_factory_code", RUN_CONFIG["fanruan_factory_code"]))
    fr_factory_combo = ttk.Combobox(
        base_tab, width=38, textvariable=fr_factory_var, values=FANRUAN_FACTORY_OPTIONS, state="readonly"
    )
    fr_factory_combo.grid(row=7, column=1, pady=4, sticky="w")
    if fr_factory_var.get() in FANRUAN_FACTORY_OPTIONS:
        fr_factory_combo.current(FANRUAN_FACTORY_OPTIONS.index(fr_factory_var.get()))
    else:
        fr_factory_combo.current(0)

    ttk.Label(base_tab, text="本地物料Excel（可选，设置后跳过OA）").grid(row=8, column=0, sticky="w")
    # 本地Excel路径为本次临时选择，不从历史配置回填
    material_excel_var = tk.StringVar(value="")
    material_excel_entry = ttk.Entry(base_tab, width=34, textvariable=material_excel_var)
    material_excel_entry.grid(row=8, column=1, pady=4, sticky="w")

    def pick_material_excel():
        file_path = filedialog.askopenfilename(
            title="选择物料编码Excel",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")],
        )
        if file_path:
            material_excel_var.set(file_path)

    ttk.Button(base_tab, text="浏览", command=pick_material_excel).grid(row=8, column=2, padx=6, sticky="w")

    remember_username_var = tk.BooleanVar(value=bool(persisted.get("remember_username", True)))
    remember_password_var = tk.BooleanVar(value=bool(persisted.get("remember_password", False)))
    wait_before_exit_var = tk.BooleanVar(value=bool(persisted.get("wait_before_exit", False)))
    ttk.Checkbutton(base_tab, text="保存用户名", variable=remember_username_var).grid(
        row=9, column=0, columnspan=2, sticky="w", pady=(6, 2)
    )
    ttk.Checkbutton(base_tab, text="记住密码（会转译保存在本地配置）", variable=remember_password_var).grid(
        row=10, column=0, columnspan=2, sticky="w", pady=(2, 2)
    )
    ttk.Checkbutton(base_tab, text="结束后等待按回车再关闭", variable=wait_before_exit_var).grid(
        row=11, column=0, columnspan=2, sticky="w", pady=(2, 6)
    )
    ttk.Label(base_tab, text="提示：账号密码可留空，仅使用Cookie。").grid(
        row=12, column=0, columnspan=2, sticky="w", pady=(10, 4)
    )

    ttk.Label(pace_tab, text="节拍配置（各个 sleep，单位秒）").grid(
        row=0, column=0, columnspan=2, sticky="w", pady=(0, 6)
    )
    sleep_vars = {}
    sleep_labels = {
        "after_kill_processes": "清理进程后等待",
        "after_cookie_apply": "Cookie应用后等待",
        "after_login_submit": "提交登录后等待",
        "after_oa_enter_plan_page": "进入OA计划页后等待",
        "after_oa_set_factory": "OA切换工厂后等待",
        "after_oa_set_year": "OA设置年份后等待",
        "after_oa_set_month": "OA设置月份后等待",
        "after_oa_submit_query": "OA提交查询后等待",
        "after_open_fanruan_page": "打开帆软页后等待",
        "before_check_fanruan_login": "检查帆软登录前等待",
        "after_fanruan_cookie_apply": "帆软Cookie应用后等待",
        "after_fanruan_login": "帆软登录完成后等待",
        "per_material_loop_start": "每个料号循环起始等待",
        "after_set_fanruan_factory": "帆软切换工厂后等待",
        "after_fill_material_code": "填料号后等待",
        "after_click_search": "点击查询后等待",
        "wait_download_poll": "下载轮询间隔",
    }
    row_idx = 1
    for key, title in sleep_labels.items():
        ttk.Label(pace_tab, text=title).grid(row=row_idx, column=0, sticky="w")
        default_v = persisted.get("sleep_config", {}).get(key, SLEEP_DEFAULTS[key])
        v = tk.StringVar(value=str(default_v))
        cb = ttk.Combobox(pace_tab, width=38, textvariable=v, values=SLEEP_OPTIONS)
        cb.grid(row=row_idx, column=1, pady=2, sticky="w")
        sleep_vars[key] = v
        row_idx += 1

    def on_ok():
        months = [m.strip().zfill(2) for m in months_var.get().split(",") if m.strip()]
        months = [m for m in months if m in MONTH_OPTIONS]
        months = list(dict.fromkeys(months))
        if not months:
            messagebox.showerror("错误", "月份不能为空")
            return
        try:
            sleep_config = {k: float(v.get().strip()) for k, v in sleep_vars.items()}
        except ValueError:
            messagebox.showerror("错误", "节拍参数必须是数字")
            return
        if any(sec <= 0 for sec in sleep_config.values()):
            messagebox.showerror("错误", "节拍参数必须大于0")
            return
        result.update({
            "oa_username": oa_user.get().strip(),
            "oa_password": oa_pwd.get().strip(),
            "fanruan_username": fr_user.get().strip(),
            "fanruan_password": fr_pwd.get().strip(),
            "oa_year": year_var.get().strip(),
            "oa_months": months,
            "oa_factory_code": oa_factory_var.get().strip(),
            "fanruan_factory_code": fr_factory_var.get().strip(),
            "material_excel_path": material_excel_var.get().strip(),
            "sleep_config": sleep_config,
            "remember_username": remember_username_var.get(),
            "remember_password": remember_password_var.get(),
            "wait_before_exit": wait_before_exit_var.get(),
            "confirmed": True
        })
        root.destroy()

    def on_cancel():
        result["confirmed"] = False
        root.destroy()

    btn_frame = ttk.Frame(outer)
    btn_frame.pack(fill="x", pady=(8, 0))
    ttk.Button(btn_frame, text="开始运行", command=on_ok).pack(side="left", padx=8)
    ttk.Button(btn_frame, text="取消", command=on_cancel).pack(side="left", padx=8)

    root.protocol("WM_DELETE_WINDOW", on_cancel)
    root.mainloop()
    return result


def apply_cookies_from_file(driver, base_url, cookie_file):
    """读取cookie文件并写入浏览器。"""
    if not os.path.exists(cookie_file):
        return False

    try:
        with open(cookie_file, "r", encoding="utf-8") as f:
            cookies = json.load(f)

        if not isinstance(cookies, list):
            print(f"Cookie文件格式错误（应为数组）: {cookie_file}")
            return False

        driver.get(base_url)
        added = 0
        for cookie in cookies:
            if not isinstance(cookie, dict):
                continue
            if "name" not in cookie or "value" not in cookie:
                continue
            clean_cookie = {
                k: v for k, v in cookie.items()
                if k in {"name", "value", "domain", "path", "expiry", "httpOnly", "secure", "sameSite"}
            }
            try:
                driver.add_cookie(clean_cookie)
                added += 1
            except Exception:
                # 域不匹配、过期等cookie直接跳过
                continue

        if added == 0:
            print(f"未成功写入任何Cookie: {cookie_file}")
            return False

        driver.get(base_url)
        pause("after_cookie_apply")
        print(f"已加载Cookie: {cookie_file}（{added}条）")
        return True
    except Exception as e:
        print(f"加载Cookie失败: {e}")
        return False


def resolve_chrome_binary():
    """自动定位 Chrome 可执行文件，避免写死用户目录。"""
    env_path = os.getenv("CHROME_BINARY")
    if env_path and os.path.exists(env_path):
        return env_path

    common_candidates = [
        os.path.join(os.environ.get("PROGRAMFILES", ""), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
    ]

    for candidate in common_candidates:
        if candidate and os.path.exists(candidate):
            return candidate

    return None


def resolve_portable_chrome_binary(project_dir):
    """优先定位内置便携版 Chrome（EXE/脚本均可用）。"""
    candidates = []

    if project_dir:
        candidates.extend([
            os.path.join(project_dir, "portable_chrome", "chrome.exe"),
            os.path.join(project_dir, "chrome-win64", "chrome.exe"),
            os.path.join(project_dir, "chrome-win", "chrome.exe"),
            os.path.join(project_dir, "Chrome-bin", "chrome.exe"),
        ])

    # EXE目录（打包后）
    exe_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else ""
    if exe_dir:
        candidates.extend([
            os.path.join(exe_dir, "portable_chrome", "chrome.exe"),
            os.path.join(exe_dir, "chrome-win64", "chrome.exe"),
            os.path.join(exe_dir, "chrome-win", "chrome.exe"),
            os.path.join(exe_dir, "Chrome-bin", "chrome.exe"),
        ])

    # PyInstaller onefile 临时解压目录
    meipass_dir = getattr(sys, "_MEIPASS", "")
    if meipass_dir:
        candidates.extend([
            os.path.join(meipass_dir, "portable_chrome", "chrome.exe"),
            os.path.join(meipass_dir, "chrome-win64", "chrome.exe"),
            os.path.join(meipass_dir, "chrome-win", "chrome.exe"),
            os.path.join(meipass_dir, "Chrome-bin", "chrome.exe"),
        ])

    # 当前工作目录兜底
    candidates.extend([
        os.path.join(os.getcwd(), "portable_chrome", "chrome.exe"),
        os.path.join(os.getcwd(), "chrome-win64", "chrome.exe"),
        os.path.join(os.getcwd(), "chrome-win", "chrome.exe"),
        os.path.join(os.getcwd(), "Chrome-bin", "chrome.exe"),
    ])

    seen = set()
    for path in candidates:
        norm = os.path.normcase(os.path.abspath(path))
        if norm in seen:
            continue
        seen.add(norm)
        if os.path.exists(path):
            return path

    return None


def resolve_chromedriver_path(project_dir):
    """优先使用环境变量，其次在常见目录查找 chromedriver。"""
    env_path = os.getenv("CHROMEDRIVER_PATH")
    if env_path and os.path.exists(env_path):
        return env_path

    candidates = []
    if project_dir:
        candidates.append(os.path.join(project_dir, "chromedriver.exe"))

    # EXE目录（打包后）
    exe_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else ""
    if exe_dir:
        candidates.append(os.path.join(exe_dir, "chromedriver.exe"))

    # PyInstaller onefile临时解压目录
    meipass_dir = getattr(sys, "_MEIPASS", "")
    if meipass_dir:
        candidates.append(os.path.join(meipass_dir, "chromedriver.exe"))

    # 当前工作目录
    candidates.append(os.path.join(os.getcwd(), "chromedriver.exe"))

    # 去重后按顺序尝试
    seen = set()
    for path in candidates:
        norm = os.path.normcase(os.path.abspath(path))
        if norm in seen:
            continue
        seen.add(norm)
        if os.path.exists(path):
            return path

    return None


def get_project_dir():
    """获取Windows项目文件夹路径（脚本所在目录）"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def create_driver(download_dir):
    """
    创建Chrome浏览器实例
    """
    # 清理残留进程
    kill_chrome_processes()

    chrome_options = Options()
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--disable-software-rasterizer')
    chrome_options.add_argument('--disable-background-networking')

    project_dir = get_project_dir()

    # 1) 优先定位内置便携版 Chrome；2) 回退系统 Chrome
    chrome_binary = resolve_portable_chrome_binary(project_dir) or resolve_chrome_binary()
    if chrome_binary:
        chrome_options.binary_location = chrome_binary
        print(f"使用Chrome路径: {chrome_options.binary_location}")
    else:
        raise Exception(
            "找不到可用 Chrome 浏览器。\n"
            "请执行以下任意一种方式：\n"
            "1) 在程序目录放置便携版浏览器（portable_chrome\\chrome.exe）\n"
            "2) 安装系统 Chrome 浏览器\n"
            "3) 设置环境变量 CHROME_BINARY 指向 chrome.exe"
        )

    # 使用独立用户数据目录
    user_data_dir = os.path.join(download_dir, "chrome_profile")
    if not os.path.exists(user_data_dir):
        os.makedirs(user_data_dir)
    chrome_options.add_argument(f'--user-data-dir={user_data_dir}')

    # 启用下载弹窗配置
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "profile.default_content_settings.popups": 0,
        "safebrowsing.enabled": True,
        "safebrowsing.disable_download_protection": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    # 优先使用 webdriver-manager
    if WEBDRIVER_MANAGER_AVAILABLE:
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            from selenium.webdriver.chrome.service import Service as ChromeService

            print("正在使用 webdriver-manager 自动下载/管理 ChromeDriver...")

            # 使用项目目录 cache，避免权限/中文路径问题
            cache_dir = os.path.join(project_dir, ".webdriver_cache")
            if not os.path.exists(cache_dir):
                os.makedirs(cache_dir)

            driver_path = ChromeDriverManager(cache_dir=cache_dir).install()
            print(f"ChromeDriver 下载路径: {driver_path}")

            service = ChromeService(executable_path=driver_path)
            driver = webdriver.Chrome(service=service, options=chrome_options)
            print("ChromeDriver 启动成功（自动管理）")
            return driver
        except Exception as e:
            print(f"webdriver-manager 失败: {e}")

    # 回退到本地 ChromeDriver（环境变量或项目目录）
    project_dir = get_project_dir()
    chromedriver_path = resolve_chromedriver_path(project_dir)

    if not chromedriver_path:
        raise Exception(
            "找不到 ChromeDriver。请执行以下任意一种方式：\n"
            "1) 安装 webdriver-manager（pip install webdriver-manager）\n"
            "2) 将 chromedriver.exe 放到项目根目录\n"
            "3) 设置环境变量 CHROMEDRIVER_PATH 指向 chromedriver.exe"
        )

    print(f"使用本地 ChromeDriver: {chromedriver_path}")
    service = Service(executable_path=chromedriver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    print("ChromeDriver 启动成功")
    return driver








def scrape_production_plan():
    """
    从惠而浦OA系统抓取生产计划数据
    选择波轮工厂，获取3月和4月数据
    """

    project_dir = get_project_dir()
    print(f"项目文件夹: {project_dir}")

    download_dir = os.path.join(project_dir, "downloads")
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)
        print(f"创建下载文件夹: {download_dir}")

    # 创建浏览器实例
    driver = create_driver(download_dir)
    driver.implicitly_wait(4)

    try:
        # ========== 可选模式：本地Excel直跑帆软，跳过OA ==========
        local_material_excel = str(RUN_CONFIG.get("material_excel_path", "")).strip()
        if local_material_excel:
            print("\n" + "=" * 10)
            print("已启用本地Excel直跑模式：跳过OA")
            print("=" * 10)
            print(f"本地物料Excel: {local_material_excel}")

            if not os.path.exists(local_material_excel):
                print("本地物料Excel路径不存在，自动回退到OA流程。")
            else:
                material_codes = get_material_codes_from_simple_excel(local_material_excel)
                if material_codes:
                    print(f"从本地Excel读取到 {len(material_codes)} 个料号")
                    print(f"前10个料号: {material_codes[:10]}")

                    query_fanruan_report(
                        driver,
                        material_codes,
                        download_dir,
                        RUN_CONFIG["fanruan_factory_code"],
                        project_dir
                    )
                    return True
                else:
                    print("本地Excel未读取到有效料号，自动回退到OA流程。")

        # ========== 第一部分：OA系统抓取数据 ==========
        print("\n" + "=" * 10)
        print("第一部分：OA系统抓取生产计划")
        print("=" * 10)

        def needs_oa_login():
            """检测 OA 是否需要登录。"""
            current_url = driver.current_url.lower()
            username_candidates = driver.find_elements(
                By.XPATH,
                "//input[@id='loginid' or @name='loginid' or contains(@placeholder,'账号') or contains(@placeholder,'用户名')]",
            )
            password_candidates = driver.find_elements(
                By.XPATH,
                "//input[@id='userpassword' or @name='userpassword' or @type='password']",
            )

            has_login_form = len(username_candidates) > 0 and len(password_candidates) > 0
            on_login_like_page = ("login" in current_url) and ("planview.jsp" not in current_url)
            return has_login_form or on_login_like_page

        oa_root_url = "https://oa.whirlpool-china.com/"
        oa_plan_url = "https://oa.whirlpool-china.com/sanyorpt/product/planview.jsp"
        oa_username, oa_password = get_credentials("OA")
        oa_cookie_path = os.path.join(project_dir, RUN_CONFIG["oa_cookie_file"])

        print("正在访问OA系统...")
        driver.get(oa_root_url)
        wait = WebDriverWait(driver, 5)

        # 优先尝试Cookie登录
        if needs_oa_login():
            print("检测到未登录，尝试Cookie登录...")
            apply_cookies_from_file(driver, oa_root_url, oa_cookie_path)

        if needs_oa_login():
            print("Cookie未生效，回退账号密码登录...")
            if not oa_username or not oa_password:
                oa_username, oa_password = prompt_manual_credentials("OA")
                if not oa_username or not oa_password:
                    raise Exception("OA未提供可用账号密码，无法继续。")
            try:
                username_input = wait.until(EC.presence_of_element_located((By.ID, "loginid")))
                password_input = driver.find_element(By.ID, "userpassword")
                login_btn = driver.find_element(By.ID, "login")
            except Exception:
                username_input = wait.until(EC.presence_of_element_located((By.NAME, "loginid")))
                password_input = driver.find_element(By.NAME, "userpassword")
                login_btn = driver.find_element(By.XPATH, "//input[@type='submit' or @value='登录']")

            username_input.clear()
            username_input.send_keys(oa_username)
            password_input.clear()
            password_input.send_keys(oa_password)
            login_btn.click()
            pause("after_login_submit")

            if needs_oa_login():
                raise Exception("OA登录后仍停留在登录页，请检查账号密码或网络/权限。")
            print("OA登录成功")
            save_cookies_to_file(driver, oa_cookie_path)
        else:
            print("检测到已登录OA，跳过登录步骤")

        print("正在访问生产计划页面...")
        driver.get(oa_plan_url)
        pause("after_oa_enter_plan_page")

        print("使用JavaScript选择OA工厂...")
        wait.until(EC.presence_of_element_located((By.NAME, "orgid")))

        driver.execute_script("""
            var select = document.getElementsByName('orgid')[0];
            select.value = arguments[0];
            var event = document.createEvent('HTMLEvents');
            event.initEvent('change', true, false);
            select.dispatchEvent(event);
        """, RUN_CONFIG["oa_factory_code"])
        print(f"已选择OA工厂({RUN_CONFIG['oa_factory_code']})")
        pause("after_oa_set_factory")

        # 按配置循环抓取多个月份
        table_by_month = {}
        for month in RUN_CONFIG["oa_months"]:
            month = str(month).zfill(2)
            print(f"设置月份: {RUN_CONFIG['oa_year']}-{month}")
            driver.execute_script("""
                var yearSelect = document.getElementsByName('styear')[0];
                yearSelect.value = arguments[0];
                var yEvent = document.createEvent('HTMLEvents');
                yEvent.initEvent('change', true, false);
                yearSelect.dispatchEvent(yEvent);
            """, str(RUN_CONFIG["oa_year"]))
            pause("after_oa_set_year")

            driver.execute_script("""
                var monSelect = document.getElementsByName('stmon')[0];
                monSelect.value = arguments[0];
                var mEvent = document.createEvent('HTMLEvents');
                mEvent.initEvent('change', true, false);
                monSelect.dispatchEvent(mEvent);
            """, month)
            pause("after_oa_set_month")

            print(f"查询 {RUN_CONFIG['oa_year']}-{month} 数据...")
            try:
                confirm_btn = driver.find_element(By.NAME, "submitbtn")
                confirm_btn.click()
            except Exception:
                driver.execute_script("document.forms[0].submit();")
            pause("after_oa_submit_query")

            print(f"提取 {RUN_CONFIG['oa_year']}-{month} 数据...")
            table_by_month[month] = extract_table_data(driver)

        print("保存数据...")
        valid_month_data = {
            m: df for m, df in table_by_month.items()
            if df is not None and len(df) > 0
        }
        for m in RUN_CONFIG["oa_months"]:
            m2 = str(m).zfill(2)
            if m2 in valid_month_data:
                print(f"{m2}月数据: {len(valid_month_data[m2])} 行")
            else:
                print(f"{m2}月无数据")

        excel_path = None
        if valid_month_data:
            excel_path = save_to_excel(valid_month_data, project_dir)
        else:
            print("未获取到数据")
            return False

        # ========== 第二部分：帆软报表系统循环查询 ==========
        if excel_path and os.path.exists(excel_path):
            print("\n" + "=" * 10)
            print("第二部分：帆软报表系统料号查询")
            print("=" * 10)

            material_codes = get_material_codes_from_excel(excel_path)

            if material_codes:
                print(f"从Excel中提取到 {len(material_codes)} 个料号")
                print(f"前10个料号: {material_codes[:10]}")

                query_fanruan_report(
                    driver,
                    material_codes,
                    download_dir,
                    RUN_CONFIG["fanruan_factory_code"],
                    project_dir
                )
            else:
                print("未从Excel中提取到料号，跳过帆软报表查询")

        return True

    except Exception as e:
        print(f"发生错误: {e}")
        import traceback
        traceback.print_exc()

        screenshot_path = os.path.join(project_dir, "error_screenshot.png")
        driver.save_screenshot(screenshot_path)
        print(f"已保存错误截图: {screenshot_path}")
        raise

    finally:
        if RUNTIME_CONFIG.get("wait_before_exit", False):
            input("\n按Enter键关闭浏览器...")
        driver.quit()


def extract_table_data(driver):
    """提取页面表格数据"""
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "maintable"))
        )

        table = driver.find_element(By.ID, "maintable")
        table_html = table.get_attribute('outerHTML')
        html_io = StringIO(table_html)
        df = pd.read_html(html_io)[0]
        df = df.dropna(how='all')

        print(f"提取到 {len(df)} 行数据")
        print("前5行预览:")
        print(df.head())

        return df

    except Exception as e:
        print(f"提取表格数据失败: {e}")
        import traceback
        traceback.print_exc()
        return None


def save_to_excel(month_data, project_dir):
    """保存数据到项目文件夹下的计划.xlsx文件（支持多个月份）"""
    file_path = os.path.join(project_dir, "计划.xlsx")

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        for month in sorted(month_data.keys()):
            df = month_data[month]
            sheet_name = f"{month}月"
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"{month}月数据已写入 '{sheet_name}' Sheet ({len(df)} 行)")

    print(f"数据已保存至: {file_path}")

    if os.path.exists(file_path):
        file_size = os.path.getsize(file_path)
        print(f"文件大小: {file_size} 字节")
        return file_path
    else:
        print("警告: 文件保存可能失败")
        return None


def get_material_codes_from_simple_excel(file_path):
    """从本地Excel读取物料编码（第一列，第二行开始）。"""
    try:
        # 按用户约定读取：第一列，从第二行开始
        df = pd.read_excel(file_path, header=None)
        if df.empty:
            print(f"本地Excel为空: {file_path}")
            return []
        if len(df.columns) < 1:
            print(f"本地Excel没有任何列: {file_path}")
            return []

        # 第二行开始 => 下标1开始
        codes = df.iloc[1:, 0].dropna().astype(str).str.strip()
        invalid_values = ['nan', 'None', '', 'NaN', 'null', 'NULL', '-']
        codes = codes[~codes.isin(invalid_values)]
        codes = sorted(list(set(codes.tolist())))
        return codes
    except Exception as e:
        print(f"读取本地物料Excel失败: {e}")
        return []


def get_material_codes_from_excel(file_path):
    """从Excel文件中提取所有料号（物料编码）"""
    material_codes = []

    try:
        xls = pd.ExcelFile(file_path)

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            possible_names = ['1']

            target_col = None
            for col in df.columns:
                col_str = str(col).strip()
                if col_str in possible_names:
                    target_col = col
                    break
                for name in possible_names:
                    if name.lower() in col_str.lower() or col_str.lower() in name.lower():
                        target_col = col
                        break
                if target_col:
                    break

            if target_col is None:
                print(f"在 {sheet_name} 中未找到物料编码列，列名: {list(df.columns)}")
                continue

            print(f"在 {sheet_name} 中找到物料编码列: {target_col}")

            codes = df[target_col].dropna().astype(str).str.strip()
            invalid_values = ['nan', 'None', '', 'NaN', 'null', 'NULL', '-']
            codes = codes[~codes.isin(invalid_values)]
            codes = codes.unique().tolist()

            print(f"从 {sheet_name} 提取到 {len(codes)} 个料号")
            material_codes.extend(codes)

        material_codes = sorted(list(set(material_codes)))

        return material_codes

    except Exception as e:
        print(f"读取Excel文件失败: {e}")
        import traceback
        traceback.print_exc()
        return []


def query_fanruan_report(driver, material_codes, download_dir, fanruan_factory_code, project_dir):
    """访问帆软报表系统，循环查询料号并导出Excel"""
    wait = WebDriverWait(driver, 4)
    fanruan_username, fanruan_password = get_credentials("FANRUAN")
    fanruan_cookie_path = os.path.join(project_dir, RUN_CONFIG["fanruan_cookie_file"])

    report_url = "https://rpt.whirlpool-china.com:7027/decision/v10/entry/access/old-platform-reportlet-entry-366?width=1690&height=524"
    print(f"\n正在访问帆软报表系统...")
    driver.get(report_url)
    pause("after_open_fanruan_page")

    # 检查是否需要登录
    try:
        pause("before_check_fanruan_login")
        login_elements = driver.find_elements(By.XPATH,
                                              "//input[@type='text' or @type='password' or contains(@placeholder, '用户名') or contains(@placeholder, '账号')]")
        current_url = driver.current_url
        print(f"当前URL: {current_url}")

        if 'login' in current_url.lower() or len(login_elements) > 0:
            print("检测到登录页面，先尝试Cookie登录...")
            apply_cookies_from_file(driver, report_url, fanruan_cookie_path)
            pause("after_fanruan_cookie_apply")
            current_url = driver.current_url
            login_elements = driver.find_elements(By.XPATH,
                                                  "//input[@type='text' or @type='password' or contains(@placeholder, '用户名') or contains(@placeholder, '账号')]")

            if 'login' in current_url.lower() or len(login_elements) > 0:
                print("Cookie未生效，回退账号密码登录...")
                if not fanruan_username or not fanruan_password:
                    fanruan_username, fanruan_password = prompt_manual_credentials("帆软")
                    if not fanruan_username or not fanruan_password:
                        raise Exception("帆软未提供可用账号密码，无法继续。")
                fanruan_login(driver, fanruan_username, fanruan_password)
                save_cookies_to_file(driver, fanruan_cookie_path)
        else:
            print("已登录或无需登录")

    except Exception as e:
        print(f"检查登录状态时出错: {e}")
        try:
            if fanruan_username and fanruan_password:
                fanruan_login(driver, fanruan_username, fanruan_password)
                save_cookies_to_file(driver, fanruan_cookie_path)
        except:
            pass

    # 等待页面完全加载
    pause("after_fanruan_login")

    # ========== 修改：检查当前工厂，如果不是664则选择 ==========
    factory_checked = False  # 标记是否已检查/设置工厂

    success_count = 0
    fail_count = 0

    for index, material_code in enumerate(material_codes, 1):
        print(f"\n[{index}/{len(material_codes)}] 正在查询料号: {material_code}")

        try:
            pause("per_material_loop_start")

            # 1. 检查并设置工厂（只在第一次执行）
            if not factory_checked:
                try:
                    # 获取当前工厂值，如果不是目标工厂则设置
                    result = driver.execute_script("""
                        var widget = _g().parameterEl.getWidgetByName('P_ORGID');

                        if (widget) {
                            var currentVal = widget.getValue();

                            if (currentVal !== arguments[0]) {
                                widget.setValue(arguments[0]);
                                widget.fireEvent("afteredit", arguments[0]);
                                return '工厂从 ' + currentVal + ' 切换到 ' + arguments[0];
                            } else {
                                return '工厂已经是 ' + arguments[0] + '，无需切换';
                            }
                        } else {
                            return '未找到工厂选择控件 P_ORGID';
                        }
                    """, fanruan_factory_code)
                    print(f"  ✓ {result}")

                    pause("after_set_fanruan_factory")
                    factory_checked = True

                except Exception as e:
                    print(f"  ! 设置工厂时出错: {e}")
                    factory_checked = True  # 即使出错也标记为已检查，避免重复尝试

            # 2. 输入料号
            try:
                item_input = wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '[widgetname="P_ITEMID"] input'))
                )
                item_input.clear()
                item_input.send_keys(material_code)
                print(f"  ✓ 已输入料号: {material_code}")
            except Exception as e:
                driver.execute_script(f"""
                    var widget = _g().parameterEl.getWidgetByName('P_ITEMID');
                    if(widget){{
                        widget.setValue('{material_code}');
                        widget.fireEvent("afteredit", "{material_code}");
                    }}
                """)
                print(f"  ✓ 已输入料号(JS): {material_code}")

            pause("after_fill_material_code")

            # 3. 点击查询按钮
            try:
                search_btn = wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, '[widgetname="SEARCH"]'))
                )
                search_btn.click()
                print("  ✓ 已点击查询按钮")
            except Exception as e:
                driver.execute_script("""
                    var widget = _g().parameterEl.getWidgetByName('SEARCH');
                    if(widget) {
                        widget.fireEvent('click');
                    } else {
                        _g().parameterCommit();
                    }
                """)
                print("  ✓ 已点击查询按钮(JS)")

            pause("after_click_search")

            # 4. 导出Excel
            export_success = export_excel_with_menu(driver, material_code, download_dir)

            if export_success:
                success_count += 1
            else:
                fail_count += 1

        except Exception as e:
            print(f"  ✗ 查询料号 {material_code} 时出错: {e}")
            fail_count += 1
            try:
                project_dir = get_project_dir()
                screenshot_path = os.path.join(project_dir, f"error_{material_code}.png")
                driver.save_screenshot(screenshot_path)
            except:
                pass
            continue

    print(f"\n{'=' * 50}")
    print(f"帆软报表查询完成:")
    print(f"  成功: {success_count} 个")
    print(f"  失败: {fail_count} 个")
    print(f"  文件保存位置: {download_dir}")
    print(f"{'=' * 50}")
def fanruan_login(driver, username, password):
    """帆软报表系统登录"""
    print(f"  正在登录帆软报表系统...")

    wait = WebDriverWait(driver, 3)

    try:
        username_input = None

        try:
            username_input = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//input[@placeholder='用户名' or @placeholder='账号' or @placeholder='用户名/账号']")
            ))
        except:
            pass

        if not username_input:
            try:
                inputs = driver.find_elements(By.XPATH, "//input[@type='text']")
                if inputs:
                    username_input = inputs[0]
            except:
                pass

        if not username_input:
            try:
                username_input = driver.find_element(By.NAME, "username")
            except:
                pass

        if not username_input:
            raise Exception("无法找到用户名输入框")

        password_input = None

        try:
            password_input = driver.find_element(By.XPATH, "//input[@type='password']")
        except:
            pass

        if not password_input:
            try:
                password_input = driver.find_element(By.NAME, "password")
            except:
                pass

        if not password_input:
            raise Exception("无法找到密码输入框")

        login_btn = None

        try:
            login_btn = driver.find_element(By.XPATH,
                                            "//button[contains(text(), '登录') or contains(text(), 'Login') or contains(text(), '提交')]")
        except:
            pass

        if not login_btn:
            try:
                login_btn = driver.find_element(By.XPATH, "//input[@type='submit']")
            except:
                pass

        if not login_btn:
            try:
                login_btn = driver.find_element(By.CSS_SELECTOR,
                                                ".login-btn, .submit-btn, .btn-login, button[type='button']")
            except:
                pass

        username_input.clear()
        username_input.send_keys(username)
        print(f"  已输入用户名: {username}")

        password_input.clear()
        password_input.send_keys(password)
        print(f"  已输入密码")

        if login_btn:
            login_btn.click()
            print(f"  已点击登录按钮")
        else:
            password_input.send_keys(u'\ue007')
            print(f"  已按回车键登录")

        pause("after_login_submit")
        print(f"  登录完成")

    except Exception as e:
        print(f"  登录失败: {e}")
        try:
            driver.save_screenshot("fanruan_login_error.png")
            print(f"  已保存登录错误截图: fanruan_login_error.png")
        except:
            pass
        raise


def export_excel_with_menu(driver, material_code, download_dir):
    """
    使用 FineReport JS 接口导出，不依赖屏幕坐标
    """
    try:
        print("  开始执行导出操作...")
        safe_material_code = "".join([c for c in material_code if c.isalnum() or c in ('-', '_')]).rstrip() or "unknown"

        existing_files = set(glob.glob(os.path.join(download_dir, "*")))
        print("  触发导出：Excel -> 原样导出")

        exported = driver.execute_script("""
            try {
                if (typeof _g === "function") {
                    _g().exportReportToExcel('simple');
                    return true;
                }
            } catch (e) {}
            try {
                if (typeof _g === "function") {
                    _g().exportReportToExcel('simple_isExcel2003');
                    return true;
                }
            } catch (e) {}
            return false;
        """)

        if not exported:
            raise Exception("未能调用 _g().exportReportToExcel('simple')，请确认页面已加载完成。")

        saved_file = wait_for_download_and_rename(download_dir, existing_files, safe_material_code)
        print(f"  ✓ 文件保存完成: {os.path.basename(saved_file)}")
        return True

    except Exception as e:
        print(f"  ✗ 导出Excel失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def wait_for_download_and_rename(download_dir, previous_files, target_name, timeout=60):
    """
    等待下载完成并按料号重命名文件。
    """
    end_time = time.time() + timeout
    latest_file = None

    while time.time() < end_time:
        current_files = set(glob.glob(os.path.join(download_dir, "*")))
        new_files = list(current_files - previous_files)

        # Chrome 临时下载文件
        temp_files = [f for f in new_files if f.lower().endswith(".crdownload")]
        if temp_files:
            pause("wait_download_poll")
            continue

        # 下载完成后常见扩展
        done_files = [
            f for f in new_files
            if os.path.isfile(f) and f.lower().endswith((".xls", ".xlsx", ".csv"))
        ]
        if done_files:
            latest_file = max(done_files, key=os.path.getmtime)
            break

        pause("wait_download_poll")

    if not latest_file:
        raise Exception("下载超时，未检测到导出文件。")

    _, ext = os.path.splitext(latest_file)
    target_path = os.path.join(download_dir, f"{target_name}{ext.lower()}")
    if os.path.exists(target_path):
        os.remove(target_path)
    os.replace(latest_file, target_path)
    return target_path


def check_dependencies():
    """检查必要的依赖是否已安装"""
    required_modules = [
        "selenium",
        "pandas",
        "openpyxl",
        "webdriver_manager",
    ]
    missing = []

    for module_name in required_modules:
        try:
            __import__(module_name)
        except ImportError:
            missing.append(module_name)

    if missing:
        print("缺少依赖，脚本无法继续运行。")
        print("缺失模块: " + ", ".join(missing))
        print("请在 PyCharm 终端执行：")
        print(
            "pip install selenium pandas openpyxl webdriver-manager"
        )
        return False

    return True


if __name__ == "__main__":
    #print("=" * 10)
    #print("惠而浦生产计划数据抓取 + 帆软报表查询工具")
    #print("=" * 10)
    #print(f"运行环境: Windows")
    #print(f"屏幕分辨率: 2560x1600")
    #print(f"抓取月份: 3月、4月")
    #print(f"下载弹窗: 已启用（可修改文件名）")
    #print()

    if not check_dependencies():
        exit(1)

    # 启动前参数弹窗（为后续EXE使用做准备）
    try:
        ui_cfg = show_runtime_config_dialog()
        if not ui_cfg.get("confirmed"):
            print("用户取消运行。")
            exit(0)
        RUNTIME_CONFIG.update(ui_cfg)
        RUN_CONFIG["oa_year"] = ui_cfg["oa_year"]
        RUN_CONFIG["oa_months"] = ui_cfg["oa_months"]
        RUN_CONFIG["oa_factory_code"] = ui_cfg["oa_factory_code"]
        RUN_CONFIG["fanruan_factory_code"] = ui_cfg["fanruan_factory_code"]
        RUN_CONFIG["material_excel_path"] = ui_cfg.get("material_excel_path", "").strip()
        SLEEP_CONFIG.update(SLEEP_DEFAULTS)
        SLEEP_CONFIG.update(ui_cfg.get("sleep_config", {}))

        # 持久化上次输入（可选择是否保存密码）
        to_save = {
            "oa_year": ui_cfg.get("oa_year", RUN_CONFIG["oa_year"]),
            "oa_months": ui_cfg.get("oa_months", RUN_CONFIG["oa_months"]),
            "oa_factory_code": ui_cfg.get("oa_factory_code", RUN_CONFIG["oa_factory_code"]),
            "fanruan_factory_code": ui_cfg.get("fanruan_factory_code", RUN_CONFIG["fanruan_factory_code"]),
            "sleep_config": dict(SLEEP_CONFIG),
            "remember_username": ui_cfg.get("remember_username", True),
            "remember_password": ui_cfg.get("remember_password", False),
            "wait_before_exit": ui_cfg.get("wait_before_exit", False),
        }
        if ui_cfg.get("remember_username", True):
            to_save["oa_username"] = ui_cfg.get("oa_username", "")
            to_save["fanruan_username"] = ui_cfg.get("fanruan_username", "")
        else:
            to_save["oa_username"] = ""
            to_save["fanruan_username"] = ""
        if ui_cfg.get("remember_password", False):
            to_save["oa_password"] = encode_local_password(ui_cfg.get("oa_password", ""))
            to_save["fanruan_password"] = encode_local_password(ui_cfg.get("fanruan_password", ""))
        else:
            to_save["oa_password"] = ""
            to_save["fanruan_password"] = ""
        save_persisted_runtime_config(to_save)
    except Exception as e:
        print(f"弹窗配置失败，继续使用默认配置: {e}")

    success = scrape_production_plan()

    if success:
        print("\n全部任务执行成功！")
    else:
        print("\n任务执行失败，请检查错误信息")