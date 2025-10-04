# spotlight.py — Spotlight clone for Windows with Everything integration
# Dependencies: pip install pywin32 pillow keyboard rapidfuzz pyperclip
# Additional: Download Everything from https://www.voidtools.com/downloads/
# Enable Everything's HTTP server: Tools > Options > HTTP Server (check "Enable HTTP Server")

import os
import subprocess
import pathlib
import threading
import tkinter as tk
from tkinter import font, Canvas
from tkinter import simpledialog, Toplevel, messagebox
import webbrowser
import re
import ctypes
from ctypes import wintypes
import sys
import time
import fnmatch
import json
import sqlite3
import urllib.request
import urllib.parse
import queue
import concurrent.futures

# Optional Pillow
try:
    from PIL import Image, ImageTk, ImageDraw
except Exception:
    Image = ImageTk = ImageDraw = None

# Optional pywin32
HAS_PYWIN32 = False
try:
    import win32com.client
    import win32ui
    import win32gui
    import win32con
    import win32api
    HAS_PYWIN32 = True
except Exception:
    win32com = win32ui = win32gui = win32con = win32api = None

# Optional rapidfuzz
try:
    from rapidfuzz import process, fuzz
except Exception:
    process = None
    fuzz = None

# Optional keyboard and pyperclip
try:
    import keyboard
except Exception:
    keyboard = None
try:
    import pyperclip
except Exception:
    pyperclip = None

COLORS = {
    "win_bg": "#070707",
    "entry_bg": "#0d0d0d",
    "entry_border": "#151515",
    "entry_fg": "#bfffbf",
    "placeholder": "#444444",
    "result_bg": "#0c0c0c",
    "result_alt": "#0f0f0f",
    "selected_bg": "#062b00",
    "selected_border": "#00ff7f",
    "selected_fg": "#c8ffd0",
    "subtle_text": "#7f7f7f"
}

SYSTEM_APPS = [
    # Basic Windows Apps
    {"name": "Notepad", "exe": "notepad.exe"},
    {"name": "Calculator", "exe": "calc.exe"},
    {"name": "Paint", "exe": "mspaint.exe"},
    {"name": "Command Prompt", "exe": "cmd.exe"},
    {"name": "Windows PowerShell", "exe": "powershell.exe"},
    {"name": "Task Manager", "exe": "taskmgr.exe"},
    {"name": "Control Panel", "exe": "control.exe"},
    {"name": "Registry Editor", "exe": "regedit.exe"},
    {"name": "System Information", "exe": "msinfo32.exe"},
    {"name": "Device Manager", "exe": "devmgmt.msc"},
    {"name": "Disk Management", "exe": "diskmgmt.msc"},
    {"name": "Services", "exe": "services.msc"},
    {"name": "Event Viewer", "exe": "eventvwr.msc"},
    {"name": "Windows Explorer", "exe": "explorer.exe"},
    {"name": "Run Dialog", "exe": "rundll32.exe shell32.dll,#61"},
    {"name": "System Configuration", "exe": "msconfig.exe"},
    {"name": "Resource Monitor", "exe": "resmon.exe"},
    {"name": "Performance Monitor", "exe": "perfmon.exe"},
    {"name": "Disk Cleanup", "exe": "cleanmgr.exe"},
    {"name": "Character Map", "exe": "charmap.exe"},
    {"name": "Magnifier", "exe": "magnify.exe"},
    {"name": "On-Screen Keyboard", "exe": "osk.exe"},
    {"name": "Snipping Tool", "exe": "snippingtool.exe"},
    {"name": "Windows Security", "exe": "windowsdefender://"},
    {"name": "Settings", "exe": "ms-settings:"},
    {"name": "Microsoft Store", "exe": "ms-windows-store:"},
    {"name": "Windows Terminal", "exe": "wt.exe"},
    {"name": "PowerShell 7", "exe": "pwsh.exe"},
    {"name": "Windows Subsystem for Linux", "exe": "wsl.exe"},
    {"name": "Remote Desktop Connection", "exe": "mstsc.exe"},
    {"name": "Windows Media Player", "exe": "wmplayer.exe"},
    {"name": "Sound Recorder", "exe": "soundrecorder.exe"},
    {"name": "Steps Recorder", "exe": "psr.exe"},
    {"name": "System File Checker", "exe": "sfc.exe"},
    {"name": "Memory Diagnostic", "exe": "mdsched.exe"},
    {"name": "Windows Update", "exe": "ms-settings:windowsupdate"},
    {"name": "Network Connections", "exe": "ncpa.cpl"},
    {"name": "Programs and Features", "exe": "appwiz.cpl"},
    {"name": "User Accounts", "exe": "netplwiz.exe"},
    {"name": "Group Policy Editor", "exe": "gpedit.msc"},
    {"name": "Local Security Policy", "exe": "secpol.msc"},
    {"name": "Computer Management", "exe": "compmgmt.msc"},
    {"name": "Disk Defragmenter", "exe": "dfrgui.exe"},
    {"name": "Windows Firewall", "exe": "wf.msc"},
    {"name": "Certificate Manager", "exe": "certmgr.msc"},
    {"name": "Local Users and Groups", "exe": "lusrmgr.msc"},
    {"name": "Task Scheduler", "exe": "taskschd.msc"},
    {"name": "Component Services", "exe": "dcomcnfg.exe"},
    {"name": "System Properties", "exe": "sysdm.cpl"},
    {"name": "Display Settings", "exe": "desk.cpl"},
    {"name": "Mouse Properties", "exe": "main.cpl"},
    {"name": "Keyboard Properties", "exe": "main.cpl @1"},
    {"name": "Sound Properties", "exe": "mmsys.cpl"},
    {"name": "Power Options", "exe": "powercfg.cpl"},
    {"name": "Date and Time", "exe": "timedate.cpl"},
    {"name": "Region Settings", "exe": "intl.cpl"},
    {"name": "Fonts", "exe": "fonts"},
    {"name": "Printers", "exe": "printers"},
    {"name": "Network", "exe": "network"},
    {"name": "Recycle Bin", "exe": "shell:RecycleBinFolder"},
    {"name": "Desktop", "exe": "shell:Desktop"},
    {"name": "Documents", "exe": "shell:Personal"},
    {"name": "Downloads", "exe": "shell:Downloads"},
    {"name": "Pictures", "exe": "shell:MyPictures"},
    {"name": "Music", "exe": "shell:MyMusic"},
    {"name": "Videos", "exe": "shell:MyVideo"},
    {"name": "Startup Folder", "exe": "shell:Startup"},
    {"name": "Recent Items", "exe": "shell:Recent"},
    {"name": "Quick Access", "exe": "shell:Quick access"},
    {"name": "This PC", "exe": "shell:MyComputerFolder"},
    {"name": "Network Places", "exe": "shell:NetworkPlacesFolder"},
    {"name": "Administrative Tools", "exe": "shell:Administrative Tools"},
    {"name": "System32", "exe": "shell:System"},
    {"name": "Program Files", "exe": "shell:ProgramFiles"},
    {"name": "Windows Folder", "exe": "shell:Windows"},
    {"name": "Temp Folder", "exe": "shell:Cache"},
    {"name": "AppData", "exe": "shell:AppData"},
    {"name": "Start Menu", "exe": "shell:Start Menu"},
    {"name": "Common Start Menu", "exe": "shell:Common Start Menu"},
    {"name": "Send To", "exe": "shell:SendTo"},
    {"name": "Templates", "exe": "shell:Templates"},
    {"name": "Cookies", "exe": "shell:Cookies"},
    {"name": "History", "exe": "shell:History"},
    {"name": "Favorites", "exe": "shell:Favorites"},
    {"name": "Profile", "exe": "shell:Profile"},
    {"name": "Public", "exe": "shell:Public"},
    {"name": "Common Documents", "exe": "shell:Common Documents"},
    {"name": "Common Pictures", "exe": "shell:CommonPictures"},
    {"name": "Common Music", "exe": "shell:CommonMusic"},
    {"name": "Common Video", "exe": "shell:CommonVideo"},
    {"name": "Common Desktop", "exe": "shell:Common Desktop"},
    {"name": "Fonts Folder", "exe": "shell:Fonts"},
    {"name": "System Folder", "exe": "shell:System"},
    {"name": "Windows Logs", "exe": "shell:Windows\\Logs"},
    {"name": "Drivers", "exe": "shell:System32\\drivers"},
    {"name": "Hosts File", "exe": "shell:System32\\drivers\\etc"},
]

# Globals
apps = []
icon_cache = {}
search_window = None
entry = None
canvas = None
placeholder_label = None
result_widgets = []
search_results = []
selected_index = -1
_search_after_id = None
_origin_x = None
_origin_y = None
backspace_empty_count = 0
_hotkey_registered = False
_default_icon = None
everything_available = False
DB_PATH = os.path.join(os.path.dirname(__file__), "file_index.db")
PREFS_PATH = os.path.join(os.path.dirname(__file__), "spotlight_prefs.json")
INDEX_PATHS = [
    os.path.expandvars(r"%APPDATA%\Microsoft\Windows\Start Menu\Programs"),
    r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
]
EVERYTHING_CLI_PATH = None
EVERYTHING_HTTP_PORT = None
HOTKEY = "win+space"
MAX_RESULTS = 8
SEARCH_DEBOUNCE_MS = 120
WINDOW_WIDTH = 820
ENTRY_HEIGHT = 72
RESULT_ITEM_HEIGHT = 56
ICON_SIZE = 32

# Detect installed browsers once so chooser can show options quickly
BROWSER_CANDIDATES = []
import shutil

def detect_browsers():
    """Populate BROWSER_CANDIDATES with available browsers (label, exe_or_None).
    exe_or_None == None means use webbrowser.open (Default browser).
    """
    mapping = {
        "Brave": [r"C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe",
                  r"C:\\Program Files (x86)\\BraveSoftware\\Brave-Browser\\Application\\brave.exe",
                  "brave.exe"],
        "Edge": [r"C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe",
                 r"C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
                 "msedge.exe"],
        "Chrome": [r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
                   r"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
                   "chrome.exe"],
        "Firefox": [r"C:\\Program Files\\Mozilla Firefox\\firefox.exe",
                    r"C:\\Program Files (x86)\\Mozilla Firefox\\firefox.exe",
                    "firefox.exe"]
    }

    results = []
    # Default browser always first
    results.append(("Default browser", None))
    for label, paths in mapping.items():
        found = None
        for p in paths:
            if p.endswith('.exe'):
                if os.path.exists(p):
                    found = p
                    break
            else:
                # try PATH lookup
                which = shutil.which(p)
                if which:
                    found = which
                    break
        if found:
            results.append((label, found))

    # Always include system PDF app fallback
    results.append(("System PDF app", "__START__"))
    BROWSER_CANDIDATES[:] = results

# Run detection immediately (cheap filesystem checks)
try:
    detect_browsers()
except Exception:
    BROWSER_CANDIDATES = [("Default browser", None), ("System PDF app", "__START__")]

def search_everything_cli(query, file_type="*", max_results=MAX_RESULTS):
    """Search using Everything's CLI (es.exe) if available"""
    if not everything_available or not EVERYTHING_CLI_PATH or not os.path.exists(EVERYTHING_CLI_PATH):
        return []
    try:
        if file_type == "folder":
            search_query = f"folder: {query}"
        else:
            search_query = query

        cmd = [EVERYTHING_CLI_PATH, "-n", str(max_results), search_query]
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace',
            timeout=2
        )
        paths = [line.strip() for line in result.stdout.strip().splitlines() if line.strip()]
        return paths[:max_results]
    except subprocess.TimeoutExpired:
        print("[everything] search timeout")
        return []
    except Exception as e:
        print(f"[everything] search error: {e}")
        return []

def search_everything_http(query, file_type="*", max_results=MAX_RESULTS):
    """Search using Everything's HTTP server (alternative method)"""
    # If port isn't configured, skip HTTP search
    if EVERYTHING_HTTP_PORT is None:
        return []
    try:
        # Build query parameters
        if file_type == "folder":
            search_query = f"folder: {query}"
        else:
            search_query = query

        params = urllib.parse.urlencode({
            'search': search_query,
            'count': max_results,
            'json': 1
        })

        url = f"http://localhost:{EVERYTHING_HTTP_PORT}/?{params}"

        with urllib.request.urlopen(url, timeout=2) as response:
            data = json.loads(response.read().decode('utf-8'))
            results = data.get('results', [])
            # Combine path and name for full path
            return [os.path.join(r['path'], r['name']) for r in results if 'path' in r and 'name' in r]
    except Exception as e:
        print(f"[everything] HTTP search failed: {e}")
        return []

# ========== WINDOWS SEARCH INDEX INTEGRATION (CUSTOM BUILT-IN) ==========

def search_windows_index(query, file_type="*", max_results=MAX_RESULTS):
    """Search using Windows built-in Search Index via pywin32 (fast, no external tools)"""
    if not HAS_PYWIN32:
        return []
    
    try:
        connection = win32com.client.Dispatch("ADODB.Connection")
        recordset = win32com.client.Dispatch("ADODB.Recordset")
        connection.Open("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';")
        
        where = "WHERE "
        if file_type == "folder":
            where += f"System.ItemName LIKE '%{query}%' AND System.Kind = 'folder'"
        else:
            where += f"System.FileName LIKE '%{query}%'"
        
        sql = f"SELECT TOP {max_results} System.ItemPathDisplay FROM SYSTEMINDEX {where}"
        
        recordset.Open(sql, connection, 0, 1)  # adOpenForwardOnly, adLockReadOnly
        
        paths = []
        while not recordset.EOF:
            paths.append(recordset.Fields.Item("System.ItemPathDisplay").Value)
            recordset.MoveNext()
        
        recordset.Close()
        connection.Close()
        return paths
    
    except Exception as e:
        print(f"[windows_index] search error: {e}")
        return []

def threaded_everything_search(query, file_type="file", callback=None):
    """Search Everything in a background thread"""
    def _search():
        # Try CLI first (faster and more reliable)
        results = search_everything_cli(query, file_type, MAX_RESULTS)
        
        # Fallback to HTTP if CLI fails and port is set
        if not results and EVERYTHING_HTTP_PORT is not None:
            results = search_everything_http(query, file_type, MAX_RESULTS)
        
        # Fallback to Windows Index if available
        if not results:
            results = search_windows_index(query, file_type, MAX_RESULTS)
        
        # Final fallback to native Python search
        if not results:
            results = native_file_search(query, file_type, MAX_RESULTS)
        
        if callback:
            search_window.after(0, lambda: callback(results))
    
    threading.Thread(target=_search, daemon=True).start()

def native_file_search(query, file_type="*", max_results=MAX_RESULTS):
    """Fallback native Python file search (slower but always works)"""
    if not query.strip():
        return []
    
    patterns = {
        "folder": None,
        "file": "*"
    }
    pattern = patterns.get(file_type, "*")
    
    # Limit search to common user paths to improve speed (change to get_all_drives() for full but slower)
    search_paths = get_search_paths()
    results = []
    
    for path in search_paths:
        if not os.path.exists(path):
            continue
        
        try:
            for root, dirs, files in os.walk(path):
                try:
                    if file_type == "folder":
                        for d in dirs:
                            if query.lower() in d.lower():
                                results.append(os.path.join(root, d))
                    else:
                        for f in files:
                            if pattern and not fnmatch.fnmatch(f.lower(), pattern.lower()):
                                continue
                            if query.lower() in f.lower():
                                results.append(os.path.join(root, f))
                    
                    if len(results) >= max_results:
                        return results[:max_results]
                except Exception:
                    continue
        except Exception:
            continue
    
    return results[:max_results]

# ========== LOCAL SQLITE DB SEARCH (OPTIONAL) ==========

def db_search(query, mode="file", limit=MAX_RESULTS):
    try:
        if not os.path.exists(DB_PATH):
            return []
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        q = f"%{query}%"
        if mode == "folder":
            cur.execute("SELECT path FROM files WHERE is_directory=1 AND name LIKE ? LIMIT ?", (q, limit))
        else:
            cur.execute("SELECT path FROM files WHERE is_directory=0 AND name LIKE ? LIMIT ?", (q, limit))
        rows = cur.fetchall()
        return [r[0] for r in rows]
    except Exception:
        return []
    finally:
        try:
            conn.close()
        except Exception:
            pass

def threaded_db_search(query, mode, callback):
    def _run():
        paths = db_search(query, mode, MAX_RESULTS)
        search_window.after(0, lambda: callback(paths))
    threading.Thread(target=_run, daemon=True).start()

# ========== HELPER FUNCTIONS ==========

def get_all_drives():
    """Get all available drives on the system"""
    if HAS_PYWIN32:
        try:
            drives = win32api.GetLogicalDriveStrings()
            drive_list = [d.strip() for d in drives.split('\000') if d.strip()]
            return drive_list
        except Exception:
            pass
    
    drive_list = []
    for letter in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
        drive = f"{letter}:\\"
        if os.path.exists(drive):
            drive_list.append(drive)
    return drive_list

def get_search_paths():
    """Get limited search paths for native search to improve speed"""
    user_home = os.path.expanduser("~")
    paths = [
        user_home,
        os.path.join(user_home, "Desktop"),
        os.path.join(user_home, "Documents"),
        os.path.join(user_home, "Downloads"),
        os.path.join(user_home, "Pictures"),
        os.path.join(user_home, "Music"),
        os.path.join(user_home, "Videos"),
        r"C:\Program Files",
        r"C:\Program Files (x86)",
        r"C:\Windows",
        # Add more common paths if needed
    ]
    return list(set(p for p in paths if os.path.exists(p)))

def enable_blur(hwnd):
    class ACCENTPOLICY(ctypes.Structure):
        _fields_ = [("AccentState", ctypes.c_int),
                    ("AccentFlags", ctypes.c_int),
                    ("GradientColor", ctypes.c_int),
                    ("AnimationId", ctypes.c_int)]
    class WINCOMPATTRDATA(ctypes.Structure):
        _fields_ = [("Attribute", ctypes.c_int),
                    ("Data", ctypes.POINTER(ACCENTPOLICY)),
                    ("SizeOfData", ctypes.c_size_t)]
    try:
        accent = ACCENTPOLICY()
        accent.AccentState = 3
        accent.AccentFlags = 2
        accent.GradientColor = 0x01000000
        data = WINCOMPATTRDATA()
        data.Attribute = 19
        data.Data = ctypes.pointer(accent)
        data.SizeOfData = ctypes.sizeof(accent)
        SetWindowCompositionAttribute = ctypes.windll.user32.SetWindowCompositionAttribute
        SetWindowCompositionAttribute.argtypes = [wintypes.HWND, ctypes.POINTER(WINCOMPATTRDATA)]
        SetWindowCompositionAttribute(hwnd, ctypes.byref(data))
    except Exception:
        pass

def round_corners(hwnd):
    try:
        DWM_WINDOW_CORNER_PREFERENCE = 33
        DWMWCP_ROUND = 2
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, DWM_WINDOW_CORNER_PREFERENCE,
            ctypes.byref(ctypes.c_int(DWMWCP_ROUND)),
            ctypes.sizeof(ctypes.c_int(DWMWCP_ROUND))
        )
    except Exception:
        pass

def create_default_icon():
    """Create a fallback icon if extraction fails."""
    global _default_icon
    if _default_icon is not None:
        return _default_icon
    try:
        if Image is None or ImageTk is None or ImageDraw is None:
            _default_icon = None
            return None
        img = Image.new("RGBA", (ICON_SIZE, ICON_SIZE), (60, 60, 60, 255))
        draw = ImageDraw.Draw(img)
        draw.text((ICON_SIZE//3, ICON_SIZE//4), "A", fill=(200, 200, 200, 255), font=None)
        _default_icon = ImageTk.PhotoImage(img)
        return _default_icon
    except:
        return None

def extract_icon(path, size=ICON_SIZE):
    if not os.path.exists(path) or not HAS_PYWIN32:
        return create_default_icon()
    if path in icon_cache:
        return icon_cache[path]

    try:
        hicon = None
        for i in range(5):
            large, small = win32gui.ExtractIconEx(path, i)
            if small:
                hicon = small[0]
                if large:
                    win32gui.DestroyIcon(large[0])
                break
            elif large:
                hicon = large[0]
                break

        if not hicon:
            icon_cache[path] = create_default_icon()
            return icon_cache[path]

        hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
        hbmp = win32ui.CreateBitmap()
        hbmp.CreateCompatibleBitmap(hdc, size, size)
        hdc_mem = hdc.CreateCompatibleDC()
        hdc_mem.SelectObject(hbmp)
        win32gui.DrawIconEx(hdc_mem.GetHandleOutput(), 0, 0, hicon, size, size, 0, 0, win32con.DI_NORMAL)

        bmpinfo = hbmp.GetInfo()
        bmpstr = hbmp.GetBitmapBits(True)
        if Image is None:
            win32gui.DestroyIcon(hicon)
            icon_cache[path] = None
            return None
        img = Image.frombuffer("RGB", (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
                               bmpstr, "raw", "BGRX", 0, 1)
        img = img.resize((size, size), Image.LANCZOS)

        win32gui.DestroyIcon(hicon)

        tk_img = ImageTk.PhotoImage(img)
        icon_cache[path] = tk_img
        return tk_img
    except Exception as e:
        print(f"[icon] failed for {path}: {e}")
        fallback = create_default_icon()
        icon_cache[path] = fallback
        return fallback

def index_apps(extract_icons=False):
    global apps
    apps = []
    if HAS_PYWIN32:
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
        except Exception:
            shell = None

    for base in INDEX_PATHS:
        if os.path.exists(base):
            for root, _, files in os.walk(base):
                for f in files:
                    if f.lower().endswith(".lnk"):
                        full_path = os.path.join(root, f)
                        target = None
                        if HAS_PYWIN32 and 'shell' in locals() and shell is not None:
                            try:
                                shortcut = shell.CreateShortcut(full_path)
                                target = shortcut.Targetpath
                            except Exception:
                                target = None
                        icon_img = None
                        if extract_icons and target:
                            icon_img = extract_icon(target) or extract_icon(full_path)
                        apps.append({
                            "name": os.path.splitext(f)[0],
                            "path": full_path,
                            "target": target,
                            "type": "app",
                            "icon": icon_img
                        })

    system32 = os.path.expandvars(r"%windir%\system32")
    for app in SYSTEM_APPS:
        exe_path = os.path.join(system32, app["exe"])
        if os.path.exists(exe_path):
            icon_img = extract_icon(exe_path) if extract_icons and HAS_PYWIN32 else None
            apps.append({
                "name": app["name"],
                "path": exe_path,
                "target": exe_path,
                "type": "system",
                "icon": icon_img
            })

    print(f"[spotlight] indexed {len(apps)} apps")

def preload_icons_background():
    """Preload all icons in background after Tk is ready."""
    time.sleep(0.5)
    print("[spotlight] preloading icons in background...")
    for app in apps:
        if app["icon"] is None:
            app["icon"] = extract_icon(app["path"]) or extract_icon(app.get("target", ""))
    print("[spotlight] icon preloading complete")

def is_url(text: str) -> bool:
    return bool(re.match(r"^(https?://)?([\w\-]+\.)+[\w\-]+", text.strip()))

def open_url(url: str):
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    webbrowser.open(url)

def calculate(expr: str):
    expr = expr.strip()
    if re.fullmatch(r"[0-9+\-*/().\s]+", expr):
        try:
            return str(eval(expr, {"__builtins__": {}}, {}))
        except Exception:
            return None
    return None

def safe_copy(text: str):
    try:
        pyperclip.copy(text)
        return True
    except Exception:
        return False

def compute_center_position(w: int, h: int):
    sw = search_window.winfo_screenwidth()
    sh = search_window.winfo_screenheight()
    return (sw - w) // 2, max(0, (sh - h) // 3)

def set_origin_for_entry():
    global _origin_x, _origin_y
    _origin_x, _origin_y = compute_center_position(WINDOW_WIDTH, ENTRY_HEIGHT)

def create_file_result(full_path, file_type):
    """Create a result dictionary for a file/folder"""
    name = os.path.basename(full_path)
    is_dir = os.path.isdir(full_path)
    
    if is_dir:
        icon_char = "📁"
        result_type = "folder"
    else:
        icon_char = "📄"
        result_type = "file"
    
    def _open_with_brave(url):
        candidates = [
            r"C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe",
            r"C:\\Program Files (x86)\\BraveSoftware\\Brave-Browser\\Application\\brave.exe",
            "brave.exe"
        ]
        for exe in candidates:
            try:
                subprocess.Popen([exe, url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                return True
            except Exception:
                continue
        return False

    def _open_with_edge(url):
        for exe in [
            r"C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe",
            r"C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
            "msedge.exe"
        ]:
            try:
                subprocess.Popen([exe, url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                return True
            except Exception:
                continue
        return False

    def _open_with_chrome(url):
        for exe in [
            r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
            r"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
            "chrome.exe"
        ]:
            try:
                subprocess.Popen([exe, url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                return True
            except Exception:
                continue
        return False

    def _open_with_firefox(url):
        for exe in [
            r"C:\\Program Files\\Mozilla Firefox\\firefox.exe",
            r"C:\\Program Files (x86)\\Mozilla Firefox\\firefox.exe",
            "firefox.exe"
        ]:
            try:
                subprocess.Popen([exe, url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                return True
            except Exception:
                continue
        return False

    def load_prefs():
        try:
            if os.path.exists(PREFS_PATH):
                with open(PREFS_PATH, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass
        return {}

    def save_prefs(prefs):
        try:
            with open(PREFS_PATH, "w", encoding="utf-8") as f:
                json.dump(prefs, f)
        except Exception:
            pass

    def show_pdf_chooser(p):
        url = pathlib.Path(p).resolve().as_uri()
        chooser = Toplevel(search_window)
        chooser.title("Open PDF with…")
        chooser.configure(bg=COLORS["win_bg"])
        chooser.attributes("-topmost", True)
        chooser.resizable(False, False)
        chooser.geometry("420x260+{}+{}".format(_origin_x + 160, _origin_y + 100))
        chooser.grab_set()
        chooser.focus_force()
        title = tk.Label(chooser, text=os.path.basename(p), bg=COLORS["win_bg"], fg=COLORS["entry_fg"], font=("Segoe UI", 11))
        title.pack(pady=(12, 8))

        # Use pre-detected browsers to render chooser instantly
        options = list(BROWSER_CANDIDATES)

        # Listbox is keyboard-focusable and supports arrow navigation
        list_frame = tk.Frame(chooser, bg=COLORS["win_bg"])
        list_frame.pack(fill="both", expand=True, padx=12)

        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        lb = tk.Listbox(list_frame, selectmode="single", activestyle="none",
                        bg=COLORS["entry_bg"], fg=COLORS["entry_fg"], highlightthickness=0,
                        bd=0, exportselection=False, yscrollcommand=scrollbar.set, font=("Segoe UI", 10))
        scrollbar.config(command=lb.yview)
        scrollbar.pack(side="right", fill="y")
        lb.pack(fill="both", expand=True)

        for label, exe in options:
            lb.insert(tk.END, label)

        # Select first item and focus listbox so arrow keys work immediately
        if lb.size() > 0:
            lb.selection_set(0)
            lb.activate(0)

        # Ensure listbox gets focus slightly after the chooser is shown
        def _focus_listbox():
            try:
                lb.focus_set()
                # make sure selection is visible
                if lb.curselection():
                    lb.see(lb.curselection()[0])
                else:
                    lb.see(0)
            except Exception:
                pass
        chooser.after(30, _focus_listbox)

        remember_var = tk.IntVar(value=0)
        remember = tk.Checkbutton(chooser, text="Always use this option for PDFs",
                                  variable=remember_var,
                                  bg=COLORS["win_bg"], fg=COLORS["entry_fg"], selectcolor=COLORS["entry_bg"], anchor="w")
        remember.pack(pady=(6, 4), padx=12, anchor="w")

        def do_open(always=False):
            sel = lb.curselection()
            idx = int(sel[0]) if sel else 0
            label, exe = options[idx]
            succeeded = False
            try:
                # If None -> default browser via webbrowser
                if exe is None:
                    succeeded = webbrowser.open(pathlib.Path(p).resolve().as_uri(), new=0, autoraise=True)
                    print(f"[open] attempted default browser for {p}: {succeeded}")
                elif exe == "__START__":
                    try:
                        os.startfile(p)
                        succeeded = True
                        print(f"[open] used os.startfile for {p}")
                    except Exception as e:
                        print(f"[open] os.startfile failed: {e}")
                else:
                    # Try launching exe with file path (some browsers prefer a path over file:// URI)
                    try:
                        subprocess.Popen([exe, str(pathlib.Path(p).resolve())], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                        succeeded = True
                        print(f"[open] launched {exe} {p}")
                    except Exception as e:
                        print(f"[open] failed to Popen {exe} with path, trying URI: {e}")
                        try:
                            subprocess.Popen([exe, pathlib.Path(p).resolve().as_uri()], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                            succeeded = True
                            print(f"[open] launched {exe} with URI {p}")
                        except Exception as e2:
                            print(f"[open] failed to Popen {exe} with URI: {e2}")
            except Exception as e:
                print(f"[open] unexpected error opening {p} with {label}: {e}")

            if not succeeded:
                # Try os.startfile as a final fallback
                try:
                    os.startfile(p)
                    succeeded = True
                    print(f"[open] fallback os.startfile succeeded for {p}")
                except Exception as e:
                    print(f"[open] fallback os.startfile failed: {e}")

            if remember_var.get() == 1 or always:
                prefs = load_prefs()
                prefs["pdf_default"] = options[idx][0]
                save_prefs(prefs)

            chooser.destroy()

        # Buttons
        btn_frame = tk.Frame(chooser, bg=COLORS["win_bg"])
        btn_frame.pack(fill="x", pady=(0,10), padx=12)
        once_btn = tk.Button(btn_frame, text="Open Once", command=lambda: do_open(always=False), bg=COLORS["entry_bg"], fg=COLORS["entry_fg"], relief="flat")
        always_btn = tk.Button(btn_frame, text="Always & Open", command=lambda: do_open(always=True), bg=COLORS["selected_bg"], fg=COLORS["selected_fg"], relief="flat")
        once_btn.pack(side="right", padx=(6,0))
        always_btn.pack(side="right")

        # Keyboard bindings: Enter/double-click to open, Escape to cancel
        def _on_enter(ev=None):
            do_open(always=False)

        def _on_double(ev=None):
            do_open(always=False)

        def _on_escape(ev=None):
            chooser.destroy()

        def _move_down(ev=None):
            try:
                cur = lb.curselection()
                idx = int(cur[0]) if cur else -1
                if idx < lb.size() - 1:
                    lb.selection_clear(0, tk.END)
                    lb.selection_set(idx + 1)
                    lb.activate(idx + 1)
                    lb.see(idx + 1)
                return "break"
            except Exception:
                return None

        def _move_up(ev=None):
            try:
                cur = lb.curselection()
                idx = int(cur[0]) if cur else 0
                if idx > 0:
                    lb.selection_clear(0, tk.END)
                    lb.selection_set(idx - 1)
                    lb.activate(idx - 1)
                    lb.see(idx - 1)
                return "break"
            except Exception:
                return None

        lb.bind('<Double-Button-1>', _on_double)
        lb.bind('<Return>', _on_enter)
        lb.bind('<Escape>', _on_escape)
        lb.bind('<Down>', _move_down)
        lb.bind('<Up>', _move_up)

        # Ensure the chooser is visible and gets focus immediately
        chooser.lift()
        chooser.focus_force()

    def get_browser_exe_by_label(lbl):
        for label, exe in BROWSER_CANDIDATES:
            if label == lbl:
                return exe
        return None

    def try_open_with_label(label, exe, filepath, url):
        """Try to open filepath/url using exe (None means default browser), return True on success"""
        try:
            if exe is None:
                # default browser
                ok = webbrowser.open(url, new=0, autoraise=True)
                print(f"[open] attempted default browser for {filepath}: {ok}")
                return bool(ok)
            if exe == "__START__":
                os.startfile(filepath)
                print(f"[open] used os.startfile for {filepath}")
                return True
            # explicit exe path
            try:
                subprocess.Popen([exe, url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                print(f"[open] launched {exe} {url}")
                return True
            except Exception as e:
                print(f"[open] failed to Popen {exe}: {e}")
                return False
        except Exception as e:
            print(f"[open] unexpected error opening {filepath} with {label}: {e}")
            return False

    def open_action():
        if not os.path.exists(full_path):
            return
        try:
            if is_dir:
                os.startfile(full_path)
            else:
                # Open Explorer with the file selected
                subprocess.run(["explorer", "/select,", full_path], check=False)
        except Exception as e:
            print(f"[spotlight] Error opening {full_path}: {e}")
            try:
                os.startfile(full_path)
            except:
                pass
    
    return {
        "name": name,
        "path": full_path,
        "type": result_type,
        "icon": icon_char,
        "action": open_action
    }

# ========== UI CREATION ==========

def create_search_window():
    global search_window, entry, canvas, placeholder_label, _default_icon
    search_window = tk.Tk()
    search_window.withdraw()
    search_window.overrideredirect(True)
    search_window.attributes("-topmost", True)
    search_window.configure(bg=COLORS["win_bg"])
    
    hwnd = ctypes.windll.user32.GetParent(search_window.winfo_id())
    style = ctypes.windll.user32.GetWindowLongW(hwnd, -20)
    style |= 0x00000080
    style &= ~0x00040000
    ctypes.windll.user32.SetWindowLongW(hwnd, -20, style)

    search_window.update_idletasks()
    set_origin_for_entry()

    try:
        from PIL import ImageDraw
        img = Image.new("RGBA", (ICON_SIZE, ICON_SIZE), (60, 60, 60))
        draw = ImageDraw.Draw(img)
        draw.text((ICON_SIZE//3, ICON_SIZE//4), "A", fill=(200, 200, 200))
        _default_icon = ImageTk.PhotoImage(img)
    except:
        _default_icon = None

    try:
        hwnd = search_window.winfo_id()
        enable_blur(hwnd)
        round_corners(hwnd)
    except Exception:
        pass

    x, y = _origin_x, _origin_y
    search_window.geometry(f"{WINDOW_WIDTH}x{ENTRY_HEIGHT}+{x}+{y}")

    canvas = Canvas(search_window, bg=COLORS["win_bg"], highlightthickness=0, bd=0)
    canvas.pack(fill="both", expand=True)

    canvas.create_text(48, ENTRY_HEIGHT // 2, text="🔍", font=("Segoe UI", 16), fill=COLORS["subtle_text"])

    frame = tk.Frame(search_window, bg=COLORS["entry_bg"])
    frame.place(x=90, y=18, width=WINDOW_WIDTH - 200, height=ENTRY_HEIGHT - 36)

    entry_font = font.Font(family="Consolas", size=16)
    entry = tk.Entry(frame, font=entry_font, bg=COLORS["entry_bg"], fg=COLORS["entry_fg"],
                     insertbackground=COLORS["selected_border"], relief="flat", bd=0,
                     selectbackground=COLORS["selected_bg"], selectforeground=COLORS["selected_fg"])
    entry.pack(fill="both", expand=True)

    search_hint = "apps, calc, url"
    placeholder_label = tk.Label(frame, text=search_hint,
                                 font=("Consolas", 13), bg=COLORS["entry_bg"],
                                 fg=COLORS["placeholder"], anchor="w")
    placeholder_label.place(x=4, y=6)

    setup_bindings()

def clear_results():
    global result_widgets
    for w in result_widgets:
        w["frame"].destroy()
    result_widgets = []

def show_results(results):
    global result_widgets, search_results, selected_index
    clear_results()
    search_results = results
    selected_index = 0 if results else -1

    if not results:
        search_window.geometry(f"{WINDOW_WIDTH}x{ENTRY_HEIGHT}+{_origin_x}+{_origin_y}")
        return

    for r in results:
        if r.get("icon") is None:
            r["icon"] = extract_icon(r.get("path", "")) or _default_icon

    new_h = ENTRY_HEIGHT + len(results) * RESULT_ITEM_HEIGHT + 20
    sw = search_window.winfo_screenheight()
    origin_y = min(_origin_y, sw - new_h - 20)
    search_window.geometry(f"{WINDOW_WIDTH}x{new_h}+{_origin_x}+{origin_y}")

    top_y = ENTRY_HEIGHT + 8
    for i, r in enumerate(results):
        bg = COLORS["result_bg"] if (i % 2 == 0) else COLORS["result_alt"]
        frame = tk.Frame(search_window, bg=bg, height=RESULT_ITEM_HEIGHT, bd=0)
        frame.place(x=16, y=top_y + i * RESULT_ITEM_HEIGHT, width=WINDOW_WIDTH - 32, height=RESULT_ITEM_HEIGHT)

        icon_to_show = r.get("icon") or _default_icon
        if isinstance(icon_to_show, ImageTk.PhotoImage):
            icon_label = tk.Label(frame, image=icon_to_show, bg=bg)
            icon_label.image = icon_to_show
        else:
            icon_text = icon_to_show if isinstance(icon_to_show, str) else "📋"
            icon_label = tk.Label(frame, text=icon_text, font=("Segoe UI Symbol", 16), bg=bg, fg=COLORS["entry_fg"])
        icon_label.place(x=12, y=12, width=36)

        title_label = tk.Label(frame, text=r["name"], font=("Segoe UI", 12), bg=bg, fg=COLORS["entry_fg"], anchor="w")
        title_label.place(x=56, y=8, width=WINDOW_WIDTH - 220)

        subtype = tk.Label(frame, text=r.get("type", ""), font=("Segoe UI", 9), bg=bg, fg=COLORS["subtle_text"], anchor="w")
        subtype.place(x=56, y=30, width=WINDOW_WIDTH - 220)

        def _on_click(ev, idx=i):
            select_result(idx)
            launch_selected()
        for w in (frame, icon_label, title_label, subtype):
            w.bind("<Button-1>", _on_click)

        result_widgets.append({"frame": frame, "icon": icon_label, "title": title_label, "sub": subtype})
    select_result(selected_index)

def select_result(idx):
    global selected_index
    if not result_widgets: return
    idx = max(0, min(idx, len(result_widgets) - 1))
    selected_index = idx
    for i, widgets in enumerate(result_widgets):
        f, icon, title, sub = widgets["frame"], widgets["icon"], widgets["title"], widgets["sub"]
        if i == idx:
            f.configure(bg=COLORS["selected_bg"])
            title.configure(bg=COLORS["selected_bg"], fg=COLORS["selected_fg"])
            sub.configure(bg=COLORS["selected_bg"], fg=COLORS["selected_border"])
        else:
            bg = COLORS["result_bg"] if (i % 2 == 0) else COLORS["result_alt"]
            f.configure(bg=bg)
            title.configure(bg=bg, fg=COLORS["entry_fg"])
            sub.configure(bg=bg, fg=COLORS["subtle_text"])

# ========== SEARCH LOGIC ==========

def perform_search():
    q = entry.get().strip()
    if q.strip() == ":resetpdf":
        try:
            if os.path.exists(PREFS_PATH):
                prefs = {}
                with open(PREFS_PATH, "w", encoding="utf-8") as f:
                    json.dump(prefs, f)
            show_results([{ "name": "PDF preference cleared.", "type": "info", "icon": "✔", "action": lambda: None }])
        except Exception:
            show_results([{ "name": "Could not clear preference.", "type": "error", "icon": "⚠️", "action": lambda: None }])
        return
    if q: placeholder_label.place_forget()
    else:
        placeholder_label.place(x=4, y=6)
        show_results([]); return

    results = []
    calc = calculate(q)
    if calc:
        results.append({"name": f"{q} → {calc}", "type": "calc", "icon": "🧮", "action": lambda v=calc: safe_copy(v)})
    if is_url(q):
        results.append({"name": q, "type": "url", "icon": "🌐", "action": lambda u=q: open_url(u)})
    
    # Regular app search
    if process and fuzz and apps:
        names = [a["name"] for a in apps]
        matches = process.extract(q, names, scorer=fuzz.WRatio, limit=MAX_RESULTS)
        for m, score, _ in matches:
            if score >= 40:
                a = next((x for x in apps if x["name"] == m), None)
                if a:
                    action = (lambda p=a["path"]: os.startfile(p)) if a["type"] != "system" else (lambda p=a["path"]: win32api.ShellExecute(0, "open", p, None, None, 1))
                    results.append({"name": a["name"], "type": a["type"], "icon": a["icon"], "action": action})
    else:
        for a in apps:
            if q.lower() in a["name"].lower():
                action = (lambda p=a["path"]: os.startfile(p)) if a["type"] != "system" else (lambda p=a["path"]: win32api.ShellExecute(0, "open", p, None, None, 1))
                results.append({"name": a["name"], "type": a["type"], "icon": a["icon"], "action": action})
    
    if not results:
        results.append({"name": f"Search web for '{q}'", "type": "web", "icon": "🔎", "action": lambda q=q: open_url("https://www.google.com/search?q=" + q)})
    show_results(results[:MAX_RESULTS])

def _debounced_search(event):
    global _search_after_id
    if _search_after_id: search_window.after_cancel(_search_after_id)
    if event.keysym not in ("Up", "Down", "Return", "Escape"):
        _search_after_id = search_window.after(SEARCH_DEBOUNCE_MS, perform_search)

# ========== NAVIGATION ==========

def on_key_nav(event):
    global selected_index
    if event.keysym == "Down":
        if search_results: select_result(min(selected_index + 1, len(search_results) - 1))
        return "break"
    elif event.keysym == "Up":
        if search_results: select_result(max(selected_index - 1, 0))
        return "break"
    elif event.keysym == "Return": launch_selected(); return "break"
    elif event.keysym == "Escape": hide(); return "break"

def launch_selected():
    if 0 <= selected_index < len(search_results):
        a = search_results[selected_index]
        try:
            # Call action first to avoid making the chooser wait while hiding/destroying TK
            if callable(a.get("action")):
                a["action"]()
        except Exception as e:
            print("[spotlight] launch error:", e)
        finally:
            # hide the launcher afterwards
            hide()

# ========== SHOW / HIDE ==========

def show():
    if not search_window: create_search_window()
    set_origin_for_entry()
    search_window.geometry(f"{WINDOW_WIDTH}x{ENTRY_HEIGHT}+{_origin_x}+{_origin_y}")
    entry.delete(0, tk.END)
    perform_search()
    search_window.deiconify()
    search_window.lift()
    hwnd = search_window.winfo_id()
    ctypes.windll.user32.SetForegroundWindow(hwnd)
    entry.focus_set()
    try:
        for a in (0.0, 0.25, 0.5, 0.8, 0.95):
            search_window.attributes("-alpha", a); search_window.update()
    except Exception: pass

def hide():
    if not search_window: return
    try:
        for a in (0.9, 0.6, 0.3, 0.0):
            search_window.attributes("-alpha", a); search_window.update()
    except Exception: pass
    search_window.withdraw()

# ========== BINDINGS ==========

def setup_bindings():
    global backspace_empty_count
    entry.bind("<KeyRelease>", _debounced_search)
    entry.bind("<KeyPress-Up>", on_key_nav)
    entry.bind("<KeyPress-Down>", on_key_nav)
    entry.bind("<KeyPress-Return>", on_key_nav)
    entry.bind("<KeyPress-Escape>", on_key_nav)

    def _backspace_close(ev):
        global backspace_empty_count
        if entry.get().strip() == "":
            backspace_empty_count += 1
            if backspace_empty_count >= 2:
                backspace_empty_count = 0
                hide()
                return "break"
        else:
            backspace_empty_count = 0

    entry.bind("<KeyRelease-BackSpace>", _backspace_close)

    def on_focus_out(ev):
        search_window.after(80, lambda: hide() if search_window.focus_get() is None else None)
    search_window.bind("<FocusOut>", on_focus_out)

# ========== HOTKEY ==========

def register_hotkey():
    global _hotkey_registered
    try:
        keyboard.unhook_all_hotkeys()
    except:
        pass
    try:
        keyboard.add_hotkey(HOTKEY, lambda: _tk_call(show))
        _hotkey_registered = True
        print(f"[spotlight] hotkey {HOTKEY} registered")
    except Exception as e:
        print("[spotlight] hotkey registration failed:", e)
        _hotkey_registered = False

def _tk_call(fn): 
    if search_window: 
        search_window.after(0, fn)

def hotkey_thread():
    global _hotkey_registered
    while True:
        if not _hotkey_registered:
            register_hotkey()
        time.sleep(5)

# ========== MAIN ==========


import pystray
from PIL import Image, ImageDraw

def create_tray_icon():
    def on_show():
        _tk_call(show)

    def on_restart_hotkey():
        register_hotkey()

    def on_exit():
        print("[spotlight] exiting...")
        os._exit(0)

    def create_image():
        image = Image.new('RGB', (64, 64), (0, 0, 0))
        d = ImageDraw.Draw(image)
        d.text((16, 16), '🔍', fill=(255, 255, 255))
        return image

    icon = pystray.Icon("spotlight")
    icon.icon = create_image()
    icon.title = "Spotlight"
    icon.menu = pystray.Menu(
        pystray.MenuItem("Show Spotlight", on_show),
        pystray.MenuItem("Restart Hotkey", on_restart_hotkey),
        pystray.MenuItem("Exit", on_exit)
    )
    icon.run()


def main():
    print("[spotlight] Using local DB and native file scan")
    print("[spotlight] Basic app search, calculator, and URL support")
    
    # Index apps WITHOUT icons first (safe before Tk)
    index_apps(extract_icons=False)
    
    # Create UI (now Tk exists)
    create_search_window()
    hide()
    
    # Now index WITH icons
    index_apps(extract_icons=True)
    
    # Preload remaining icons in background (non-blocking)
    threading.Thread(target=preload_icons_background, daemon=True).start()
    
    # Start hotkey watcher
    threading.Thread(target=hotkey_thread, daemon=True).start()
    threading.Thread(target=create_tray_icon, daemon=True).start()
    
    print(f"[spotlight] ready — press {HOTKEY} to open")
    search_window.mainloop()

if __name__ == "__main__":
    main()
