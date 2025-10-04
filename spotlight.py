# spotlight.py — Spotlight clone for Windows (FULLY FIXED)
# Dependencies: pip install pywin32 pillow keyboard rapidfuzz pyperclip
# spotlight.py — Spotlight clone for Windows
# Dependencies (optional): pywin32 pillow keyboard rapidfuzz pyperclip

import os
import threading
import tkinter as tk
from tkinter import font, Canvas
import webbrowser
import re
import ctypes
from ctypes import wintypes
import sys
import time

# Optional Pillow
try:
    from PIL import Image, ImageTk, ImageDraw
except Exception:
    Image = ImageTk = ImageDraw = None

# Optional pywin32 (Windows-only features)
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
INDEX_PATHS = [
    os.path.expandvars(r"%APPDATA%\Microsoft\Windows\Start Menu\Programs"),
    r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
]
HOTKEY = "ctrl+space"
MAX_RESULTS = 8
SEARCH_DEBOUNCE_MS = 120
WINDOW_WIDTH = 820
ENTRY_HEIGHT = 72
RESULT_ITEM_HEIGHT = 56
ICON_SIZE = 32

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

# ---------- SYSTEM APPS ----------
SYSTEM_APPS = [
    {"name": "Notepad", "exe": "notepad.exe"},
    {"name": "Calculator", "exe": "calc.exe"},
    {"name": "Paint", "exe": "mspaint.exe"},
    {"name": "Command Prompt", "exe": "cmd.exe"},
    {"name": "Windows PowerShell", "exe": "powershell.exe"},
    {"name": "Task Manager", "exe": "taskmgr.exe"},
    {"name": "Control Panel", "exe": "control.exe"},
    {"name": "Registry Editor", "exe": "regedit.exe"},
    {"name": "System Information", "exe": "msinfo32.exe"},
]

# ---------- GLOBALS ----------
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
_default_icon = None  # Will be set after Tk init

# ---------- WINDOWS BLUR & ROUNDED ----------
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

# ---------- ICON EXTRACTION ----------
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
        # Fallback to text if PIL.ImageDraw fails
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

# ---------- INDEX REAL APPS ----------
def index_apps(extract_icons=False):
    global apps
    apps = []
    # If pywin32 is available, resolve .lnk targets; otherwise add .lnk entries without resolving
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

    # Add system apps
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

# ---------- BACKGROUND ICON LOADER ----------
def preload_icons_background():
    """Preload all icons in background after Tk is ready."""
    time.sleep(0.5)  # Ensure UI is stable
    print("[spotlight] preloading icons in background...")
    for app in apps:
        if app["icon"] is None:
            app["icon"] = extract_icon(app["path"]) or extract_icon(app.get("target", ""))
    print("[spotlight] icon preloading complete")

# ---------- HELPERS ----------
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

# ---------- UI CREATION ----------
def create_search_window():
    global search_window, entry, canvas, placeholder_label, _default_icon
    search_window = tk.Tk()
    search_window.withdraw()
    search_window.overrideredirect(True)
    search_window.attributes("-topmost", True)
    search_window.configure(bg=COLORS["win_bg"])
    
    # Make it a tool window to avoid taskbar
    hwnd = ctypes.windll.user32.GetParent(search_window.winfo_id())
    style = ctypes.windll.user32.GetWindowLongW(hwnd, -20)  # GWL_EXSTYLE
    style |= 0x00000080  # WS_EX_TOOLWINDOW
    style &= ~0x00040000  # Remove WS_EX_APPWINDOW
    ctypes.windll.user32.SetWindowLongW(hwnd, -20, style)

    search_window.update_idletasks()
    set_origin_for_entry()

    # Initialize default icon now that Tk exists
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

    placeholder_label = tk.Label(frame, text="Type to search (apps, calc, url)…",
                                 font=("Consolas", 13), bg=COLORS["entry_bg"],
                                 fg=COLORS["placeholder"], anchor="w")
    placeholder_label.place(x=4, y=6)

    setup_bindings()

# ---------- RESULTS ----------
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

    # Ensure icons are loaded (fallback if missing)
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
            icon_label = tk.Label(frame, text="", font=("Segoe UI Symbol", 16), bg=bg, fg=COLORS["entry_fg"])
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

# ---------- SEARCH ----------
def perform_search():
    q = entry.get().strip()
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

# ---------- NAVIGATION ----------
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
        hide()
        try:
            if callable(a.get("action")): a["action"]()
        except Exception as e:
            print("[spotlight] launch error:", e)

# ---------- SHOW / HIDE ----------
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

# ---------- BINDINGS ----------
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

# ---------- HOTKEY ----------
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

# ---------- MAIN ----------
def main():
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
    
    print(f"[spotlight] ready — press {HOTKEY} to open")
    search_window.mainloop()

if __name__ == "__main__":
    main()