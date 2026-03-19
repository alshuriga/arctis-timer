"""
Arctis 7 Auto-Off Timer
Monitors Windows audio and controls the SteelSeries Arctis 7 inactivity timer
via USB HID. Runs as a system tray application.
"""

import ctypes
import ctypes.wintypes as wintypes
import os
import sys
import threading
import time
import json
import tkinter as tk
from tkinter import ttk
import winreg

import hid
import pystray
from PIL import Image, ImageDraw
from pycaw.pycaw import AudioUtilities, IAudioMeterInformation
from win11toast import toast

# --- Win32 Registry for Auto-Start ---
REG_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
REG_NAME = "ArctisTimer"

def get_executable_path():
    """Returns the full path to the current executable or script."""
    if getattr(sys, 'frozen', False):
        return sys.executable
    return os.path.abspath(sys.argv[0])

def is_autostart_enabled() -> bool:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH, 0, winreg.KEY_READ) as key:
            value, _ = winreg.QueryValueEx(key, REG_NAME)
            # Compare current path with stored path (allow for small quotes/path diffs)
            curr = get_executable_path().lower().strip('"')
            stored = value.lower().strip('"')
            return curr in stored
    except FileNotFoundError:
        return False
    except Exception:
        return False

def set_autostart(enabled: bool):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH, 0, winreg.KEY_SET_VALUE) as key:
            if enabled:
                path = f'"{get_executable_path()}"'
                winreg.SetValueEx(key, REG_NAME, 0, winreg.REG_SZ, path)
                print(f"[REG] Auto-start enabled: {path}")
            else:
                try:
                    winreg.DeleteValue(key, REG_NAME)
                    print("[REG] Auto-start disabled")
                except FileNotFoundError:
                    pass
    except Exception as e:
        print(f"[REG] Error managing auto-start: {e}")

import pystray
from PIL import Image, ImageDraw
from pycaw.pycaw import AudioUtilities, IAudioMeterInformation
from win11toast import toast

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

APP_NAME = "Arctis 7 Timer"
APP_VERSION = "1.0.0"

ARCTIS_VID = 0x1038
ARCTIS_PIDS = {
    0x1260: "Arctis 7",
    0x12AD: "Arctis 7 (2019)",
    0x1252: "Arctis Pro (2019)",
    0x1280: "Arctis Pro GameDAC",
}

# HID control interface on the Arctis 7 USB receiver
TARGET_USAGE_PAGE = 0xFF43
TARGET_USAGE      = 0x0202
HID_REPORT_ID     = 0x06
HID_REPORT_SIZE   = 31    # bytes, including report ID
CMD_SET_TIMER     = 0x51

# Win32 flags for CreateFile
GENERIC_READ         = 0x80000000
GENERIC_WRITE        = 0x40000000
FILE_SHARE_READ      = 0x00000001
FILE_SHARE_WRITE     = 0x00000002
OPEN_EXISTING        = 3
INVALID_HANDLE_VALUE = wintypes.HANDLE(-1).value

_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
_user32   = ctypes.WinDLL("user32",   use_last_error=True)

class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.UINT),
        ("dwTime", wintypes.DWORD),
    ]

def get_idle_time() -> float:
    """Return seconds since last mouse/keyboard activity."""
    lii = LASTINPUTINFO()
    lii.cbSize = ctypes.sizeof(LASTINPUTINFO)
    if _user32.GetLastInputInfo(ctypes.byref(lii)):
        # dwTime is in ms since system start. GetTickCount is also ms since system start.
        millis = _kernel32.GetTickCount() - lii.dwTime
        return millis / 1000.0
    return 0.0

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

SETTINGS_DIR  = os.path.join(os.environ.get("APPDATA", "."), "ArctisTimer")
SETTINGS_FILE = os.path.join(SETTINGS_DIR, "settings.json")

DEFAULTS = {
    "inactive_timer_minutes": 1,    # 1 = minimum, 0 = never (controls how long before headset auto-offs)
    "silence_duration_seconds": 30, # seconds of silence before activating the timer
    "active_duration_seconds": 2,   # seconds of audio before deactivating the timer
    "silence_threshold": 0.001,     # peak audio level considered silent
    "notifications_enabled": True,  # global toggle
    "afk_timeout_minutes": 10,      # If idle this long, force inactivity timer
    "detection_mode": "Both",       # "Both" | "Silence only" | "Idle only"
    "notification_mode": "Both",    # "Both" | "Silence only" | "Idle only"
}

DETECTION_MODES = ["Both", "Silence only", "Idle only"]


def load_settings() -> dict:
    try:
        with open(SETTINGS_FILE, "r") as f:
            data = json.load(f)
        # Fill in any missing keys with defaults
        return {**DEFAULTS, **data}
    except (FileNotFoundError, json.JSONDecodeError):
        return dict(DEFAULTS)


def save_settings(settings: dict):
    os.makedirs(SETTINGS_DIR, exist_ok=True)
    with open(SETTINGS_FILE, "w") as f:
        json.dump(settings, f, indent=2)


# ---------------------------------------------------------------------------
# HID / Arctis controller
# ---------------------------------------------------------------------------

class ArctisController:
    def __init__(self):
        self._path: bytes | None = None
        self._lock = threading.Lock()
        self._find_device()

    def _find_device(self):
        for pid in ARCTIS_PIDS:
            for dev in hid.enumerate(ARCTIS_VID, pid):
                if dev["usage_page"] == TARGET_USAGE_PAGE and dev["usage"] == TARGET_USAGE:
                    self._path = dev["path"]
                    print(f"[HID] Found {ARCTIS_PIDS[pid]} on {self._path}")
                    return
        print("[HID] Device not found — is the USB dongle plugged in?")

    def _open_handle(self):
        if self._path is None:
            return None
        path_str = self._path.decode("utf-8") if isinstance(self._path, bytes) else self._path
        h = _kernel32.CreateFileW(
            path_str,
            GENERIC_READ | GENERIC_WRITE,
            FILE_SHARE_READ | FILE_SHARE_WRITE,
            None, OPEN_EXISTING, 0, None,
        )
        return None if h == INVALID_HANDLE_VALUE else h

    def set_inactivity_timer(self, minutes: int) -> bool:
        """Send the inactivity-timer HID command. minutes=0 means Never."""
        with self._lock:
            if self._path is None:
                self._find_device()
                if self._path is None:
                    return False

            report = bytearray(HID_REPORT_SIZE)
            report[0] = HID_REPORT_ID
            report[1] = CMD_SET_TIMER
            report[2] = minutes

            h = self._open_handle()
            if h is None:
                print(f"[HID] CreateFile failed: {ctypes.get_last_error()}")
                self._path = None   # re-scan next time
                return False
            try:
                buf = (ctypes.c_byte * len(report))(*report)
                bw  = wintypes.DWORD(0)
                ok  = _kernel32.WriteFile(h, buf, len(report), ctypes.byref(bw), None)
                if not ok:
                    print(f"[HID] WriteFile error {ctypes.get_last_error()}")
                    return False
                print(f"[HID] Wrote {bw.value} bytes → timer={minutes}min")
                return bw.value > 0
            finally:
                _kernel32.CloseHandle(h)


# ---------------------------------------------------------------------------
# Audio monitor
# ---------------------------------------------------------------------------

class AudioMonitor(threading.Thread):
    def __init__(self, arctis: ArctisController, settings: dict):
        super().__init__(daemon=True)
        self.arctis   = arctis
        self.settings = settings          # live dict — changes apply on next poll
        self._stop_event = threading.Event()
        self._state  = "UNKNOWN"          # "ACTIVE" | "INACTIVE"
        self._silence_start: float | None = None
        self._active_start:  float | None = None

    def stop(self):
        self._stop_event.set()

    @staticmethod
    def _peek() -> float:
        """Return the highest audio peak across all active sessions."""
        try:
            sessions = AudioUtilities.GetAllSessions()
            peak = 0.0
            for session in sessions:
                if session.Process:
                    meter = session._ctl.QueryInterface(IAudioMeterInformation)
                    peak = max(peak, meter.GetPeakValue())
            return peak
        except Exception:
            return 0.0

    def run(self):
        # On startup, immediately mark device as Active/Stay On
        print("[Audio] Initializing... Sending STAY-ACTIVE signal")
        self.arctis.set_inactivity_timer(0)
        self._state = "ACTIVE"
        
        while not self._stop_event.is_set():
            s = self.settings
            threshold       = s["silence_threshold"]
            silence_needed  = s["silence_duration_seconds"]
            active_needed   = s["active_duration_seconds"]
            inactive_min    = s["inactive_timer_minutes"]
            poll            = 2.0

            playing = self._peek() > threshold
            idle    = get_idle_time()
            
            mode = s.get("detection_mode", "Both")
            rule_silence = mode in ["Both", "Silence only"]
            rule_idle    = mode in ["Both", "Idle only"]
            
            # AFK is only relevant if idle rule is on
            afk = rule_idle and (idle >= (s.get("afk_timeout_minutes", 10) * 60))
            # Sound is only relevant if silence rule is on
            sound_playing = rule_silence and playing
            
            now = time.time()

            # We stay ACTIVE if sound is playing (and we aren't AFK)
            # OR if we are only in Idle mode and not yet AFK
            if (sound_playing and not afk) or (not rule_silence and not afk):
                self._silence_start = None
                if self._active_start is None:
                    self._active_start = now
                if self._state != "ACTIVE" and (now - self._active_start) >= active_needed:
                    ts = time.strftime("%H:%M:%S")
                    print(f"[{ts}] Active state detected → timer OFF")
                    if self.arctis.set_inactivity_timer(0):
                        self._state = "ACTIVE"
                        self._notify("🎧 Active", "Headset will stay on.", "Both")
            else:
                self._active_start = None
                if self._silence_start is None:
                    self._silence_start = now
                elapsed = now - self._silence_start
                
                # Determine trigger reason
                reason = "Silence"
                if afk:
                    reason = "Idle"
                if playing and afk:
                    reason = "Idle (playing)"

                if self._state != "INACTIVE" and elapsed >= silence_needed:
                    ts = time.strftime("%H:%M:%S")
                    print(f"[{ts}] {reason} trigger ({elapsed:.0f}s) → timer={inactive_min}min")
                    if self.arctis.set_inactivity_timer(inactive_min):
                        self._state = "INACTIVE"
                        self._notify(f"🔇 {reason} detected",
                                     f"Auto-off in {inactive_min} min.", reason)

            self._stop_event.wait(poll)

    def _notify(self, title: str, message: str, trigger_type: str):
        """
        trigger_type should be "Silence", "Idle", or "Both"/"Active"
        """
        if not self.settings.get("notifications_enabled", True):
            return
            
        n_mode = self.settings.get("notification_mode", "Both")
        if n_mode == "Silence only" and "Idle" in trigger_type:
            return
        if n_mode == "Idle only" and "Silence" in trigger_type:
            return
            
        try:
            toast(APP_NAME, f"{title} — {message}",
                  audio={"src": "ms-winsoundevent:Notification.Default", "silent": "true"})
        except Exception:
            pass


# ---------------------------------------------------------------------------
# ToolTips
# ---------------------------------------------------------------------------

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.id = None
        self.x = self.y = 0
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(1000, self.showtip)

    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)

    def showtip(self, event=None):
        if self.tip_window or not self.text:
            return
        
        # Calculate cursor position for more accurate placement
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        
        self.tip_window = tw = tk.Toplevel(self.widget)
        # Hide the system window decorations
        tw.wm_overrideredirect(True)
        # Ensure it stays on top of the topmost parent
        tw.wm_attributes("-topmost", True)
        # Position the window
        tw.wm_geometry("+%d+%d" % (x, y))
        
        label = tk.Label(tw, text=self.text, justify='left',
                         background="#ffffff", foreground="#1c1c1e", 
                         highlightthickness=1, highlightbackground="#d1d1d6",
                         relief='flat', border=0, padx=8, pady=5, 
                         font=("Segoe UI Variable Text", 9))
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tip_window
        self.tip_window = None
        if tw:
            tw.destroy()


# ---------------------------------------------------------------------------
# Settings window (Tkinter)
# ---------------------------------------------------------------------------

# --- UI Theme (macOS Light) ---
ACCENT   = "#007AFF"   # macOS blue
BG_DARK  = "#F2F2F7"   # system light gray
BG_CARD  = "#FFFFFF"   # white cards
FG_TEXT  = "#1C1C1E"   # label dark
FG_DIM   = "#8E8E93"   # system gray
BORDER   = "#D1D1D6"   # outer border grey
# ---------------------------------------------------------------------------

def open_settings_window(settings: dict, on_save):
    """Open the settings dialog with a modern, stylish look."""
    win = tk.Toplevel()
    win.title(f"{APP_NAME} — Settings")
    win.configure(bg=BG_DARK)
    win.resizable(False, True) # Allow vertical resize in case of layout overflow
    win.attributes("-topmost", True)

    # Apply style to ttk components
    style = ttk.Style()
    style.theme_use('clam')
    style.configure("TSpinbox", fieldbackground=BG_CARD, background=BG_CARD, 
                    foreground=FG_TEXT, bordercolor=BORDER, arrowcolor=ACCENT)
    
    # Modern Combobox style
    style.configure("TCombobox", fieldbackground=BG_CARD, background=BG_CARD, 
                    foreground=FG_TEXT, bordercolor=BORDER, arrowcolor=ACCENT)
    win.option_add("*TCombobox*Listbox.background", BG_CARD)
    win.option_add("*TCombobox*Listbox.foreground", FG_TEXT)
    win.option_add("*TCombobox*Listbox.selectBackground", ACCENT)
    
    # Modern Checkbutton style
    style.configure("TCheckbutton", background=BG_CARD, foreground=FG_TEXT, font=("Segoe UI Variable Text", 9))
    style.map("TCheckbutton", background=[('active', BG_CARD)], foreground=[('active', ACCENT)])

    def heading(parent, text):
        f = tk.Frame(parent, bg=BG_DARK)
        f.pack(fill="x", pady=(15, 5))
        tk.Label(f, text=text, bg=BG_DARK, fg=ACCENT,
                 font=("Segoe UI Variable Display", 10, "bold")).pack(side="left")

    def row(parent, lbl_text, var, from_, to_, unit="", help_text=""):
        frame = tk.Frame(parent, bg=BG_CARD, pady=10, padx=12, 
                         highlightthickness=1, highlightbackground=BORDER)
        frame.pack(fill="x", pady=2)
        
        lbl = tk.Label(frame, text=lbl_text, bg=BG_CARD, fg=FG_TEXT,
                 font=("Segoe UI Variable Text", 9), width=28, anchor="w")
        lbl.pack(side="left")
        
        sb = ttk.Spinbox(frame, from_=from_, to=to_, textvariable=var,
                         width=6, font=("Segoe UI Variable Text", 9))
        sb.pack(side="left", padx=(0, 8))
        
        if unit:
            u_lbl = tk.Label(frame, text=unit, bg=BG_CARD, fg=FG_DIM,
                     font=("Segoe UI Variable Text", 9))
            u_lbl.pack(side="left")
            
        if help_text:
            ToolTip(frame, help_text)

    def row_combo(parent, lbl_text, var, options, help_text=""):
        frame = tk.Frame(parent, bg=BG_CARD, pady=10, padx=12, 
                         highlightthickness=1, highlightbackground=BORDER)
        frame.pack(fill="x", pady=2)
        lbl = tk.Label(frame, text=lbl_text, bg=BG_CARD, fg=FG_TEXT,
                 font=("Segoe UI Variable Text", 9), width=28, anchor="w")
        lbl.pack(side="left")
        cb = ttk.Combobox(frame, textvariable=var, values=options, state="readonly", width=12)
        cb.pack(side="left", padx=(0, 8))
        if help_text:
            ToolTip(frame, help_text)
            ToolTip(lbl, help_text)
            ToolTip(cb, help_text)

    def row_meter(parent, lbl_text, threshold_var):
        frame = tk.Frame(parent, bg=BG_CARD, pady=10, padx=12, 
                         highlightthickness=1, highlightbackground=BORDER)
        frame.pack(fill="x", pady=2)
        tk.Label(frame, text=lbl_text, bg=BG_CARD, fg=FG_TEXT,
                 font=("Segoe UI Variable Text", 9), width=28, anchor="w").pack(side="left")
        
        c = tk.Canvas(frame, width=150, height=20, bg="#E5E5EA", highlightthickness=0)
        c.pack(side="left", padx=5)
        
        # Internal meter objects
        bar = c.create_rectangle(0, 0, 0, 20, fill=ACCENT, outline="")
        mark = c.create_line(0, 0, 0, 20, fill="#FF3B30", width=2)
        
        def update_meter():
            try:
                # Get current peak from AudioMonitor
                peak = AudioMonitor._peek()
                thresh = threshold_var.get()
                
                # Scale for visualization (linear for simplicity, maybe log better but let's see)
                # Max peak usually 1.0 but small values matter here
                # We'll zoom into the 0.0 to 0.1 range roughly
                w = 150
                peak_w = min(w, peak * 500) # zoom in
                thresh_x = min(w, thresh * 500)
                
                c.coords(bar, 0, 0, peak_w, 20)
                c.coords(mark, thresh_x, 0, thresh_x, 20)
                
                if win.winfo_exists():
                    win.after(100, update_meter)
            except Exception:
                pass
                
        update_meter()
        ToolTip(frame, "Visual peak meter. The red line is your current silence threshold.")

    # Variables
    v_inactive   = tk.IntVar(value=settings["inactive_timer_minutes"])
    v_silence_s  = tk.IntVar(value=settings["silence_duration_seconds"])
    v_active_s   = tk.IntVar(value=settings["active_duration_seconds"])
    v_threshold  = tk.DoubleVar(value=settings["silence_threshold"])
    v_afk        = tk.IntVar(value=settings.get("afk_timeout_minutes", 10))
    v_det_mode   = tk.StringVar(value=settings.get("detection_mode", "Both"))
    v_notif_mode = tk.StringVar(value=settings.get("notification_mode", "Both"))
    v_notifs     = tk.BooleanVar(value=settings.get("notifications_enabled", True))
    v_autostart  = tk.BooleanVar(value=is_autostart_enabled())

    pad = tk.Frame(win, bg=BG_DARK, padx=25, pady=20)
    pad.pack(fill="both", expand=True)

    header_lbl = tk.Label(pad, text="App Preferences", bg=BG_DARK, fg=FG_TEXT,
                          font=("Segoe UI Variable Display", 14, "bold"))
    header_lbl.pack(anchor="w", pady=(0, 10))

    heading(pad, "AUTO-OFF LOGIC")
    row_combo(pad, "Detection mode", v_det_mode, DETECTION_MODES,
              "Choose whether the app monitors silence, idle time, or both.")
    row(pad, "Inactivity timer (when silent)", v_inactive, 1, 90, "min",
        "Sets how many minutes after silence until the headset turns itself off.")
    row(pad, "AFK / Idle timeout", v_afk, 1, 1440, "min",
        "Sets how many minutes of no activity before headset turns off.")

    heading(pad, "TIMING & DETECTION")
    row_meter(pad, "Live peak meter (red=limit)", v_threshold)
    row(pad, "Silence detection duration",     v_silence_s, 5, 3600, "sec",
        "How long audio must stay quiet before the auto-off timer starts.")
    row(pad, "Audio detection duration",       v_active_s,  1, 60,   "sec",
        "How long audio must play before the headset cancels any pending auto-off timer.")
    row(pad, "Silence threshold (peak)",       v_threshold, 0.0001, 0.1, "",
        "The sensitivity for audio detection (lower = more sensitive).")

    heading(pad, "INTERFACE")
    row_combo(pad, "Notification mode", v_notif_mode, DETECTION_MODES,
              "Choose which events should trigger a desktop notification.")

    notif_card = tk.Frame(pad, bg=BG_CARD, pady=10, padx=12, 
                          highlightthickness=1, highlightbackground=BORDER)
    notif_card.pack(fill="x", pady=2)
    
    cb_notif = ttk.Checkbutton(
        notif_card,
        text="Enable desktop notifications (Global)",
        variable=v_notifs,
        style="TCheckbutton"
    )
    cb_notif.pack(side="left")
    ToolTip(notif_card, "Globally turn Windows notifications on or off.")

    startup_card = tk.Frame(pad, bg=BG_CARD, pady=10, padx=12, 
                            highlightthickness=1, highlightbackground=BORDER)
    startup_card.pack(fill="x", pady=2)
    
    cb_start = ttk.Checkbutton(
        startup_card,
        text="Run automatically at Windows login",
        variable=v_autostart,
        style="TCheckbutton"
    )
    cb_start.pack(side="left")
    ToolTip(startup_card, "Enable or disable automatic start when you log into Windows.")

    def on_save_click():
        settings["inactive_timer_minutes"]    = v_inactive.get()
        settings["silence_duration_seconds"]  = v_silence_s.get()
        settings["active_duration_seconds"]   = v_active_s.get()
        settings["silence_threshold"]         = round(v_threshold.get(), 6)
        settings["afk_timeout_minutes"]       = v_afk.get()
        settings["detection_mode"]            = v_det_mode.get()
        settings["notification_mode"]         = v_notif_mode.get()
        settings["notifications_enabled"]     = bool(v_notifs.get())
        save_settings(settings)
        set_autostart(v_autostart.get())
        if on_save:
            on_save()
        win.destroy()

    def on_reset_click():
        v_inactive.set(DEFAULTS["inactive_timer_minutes"])
        v_silence_s.set(DEFAULTS["silence_duration_seconds"])
        v_active_s.set(DEFAULTS["active_duration_seconds"])
        v_threshold.set(DEFAULTS["silence_threshold"])
        v_afk.set(DEFAULTS["afk_timeout_minutes"])
        v_det_mode.set(DEFAULTS["detection_mode"])
        v_notif_mode.set(DEFAULTS["notification_mode"])
        v_notifs.set(DEFAULTS["notifications_enabled"])

    btn_frame = tk.Frame(win, bg=BG_DARK, pady=20, padx=25)
    btn_frame.pack(fill="x")

    save_btn = tk.Button(
        btn_frame, text="Apply Changes", command=on_save_click,
        bg=ACCENT, fg="#FFFFFF", font=("Segoe UI Variable Text", 9, "bold"),
        relief="flat", padx=20, pady=8, cursor="hand2",
        activebackground="#0056b3", activeforeground="#FFFFFF",
    )
    save_btn.pack(side="right")
    
    cancel_btn = tk.Button(
        btn_frame, text="Cancel", command=win.destroy,
        bg=BG_DARK, fg=FG_DIM, font=("Segoe UI Variable Text", 9),
        relief="flat", padx=15, pady=8, cursor="hand2",
        activebackground=BG_CARD, activeforeground=FG_TEXT
    )
    cancel_btn.pack(side="right", padx=(0, 10))

    reset_btn = tk.Button(
        btn_frame, text="Reset Defaults", command=on_reset_click,
        bg=BG_DARK, fg="#FF3B30", font=("Segoe UI Variable Text", 9),
        relief="flat", padx=10, pady=8, cursor="hand2",
        activebackground=BG_CARD, activeforeground="#D70015"
    )
    reset_btn.pack(side="left")

    # Center on screen
    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"+{x}+{y}")
    win.grab_set()

# ---------------------------------------------------------------------------
# Tray icon
# ---------------------------------------------------------------------------

def _make_icon_image() -> Image.Image:
    size  = 64
    img   = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw  = ImageDraw.Draw(img)
    green = (16, 185, 129, 255)
    white = (255, 255, 255, 255)

    # Headband arc
    draw.arc([10, 6, 54, 44], start=200, end=340, fill=green, width=6)
    # Left stem
    draw.line([10, 32, 10, 46], fill=green, width=5)
    # Right stem
    draw.line([54, 32, 54, 46], fill=green, width=5)
    # Left ear cup
    draw.ellipse([4, 42, 22, 58], fill=green)
    # Right ear cup
    draw.ellipse([42, 42, 60, 58], fill=green)

    return img


class TrayApp:
    def __init__(self):
        self.settings = load_settings()
        self.arctis   = ArctisController()
        self.monitor  = AudioMonitor(self.arctis, self.settings)

        # Hidden tkinter root for running the settings window on the main thread
        self._tk_root = tk.Tk()
        self._tk_root.withdraw()
        self._tk_root.overrideredirect(True)

        self._icon = pystray.Icon(
            name   = APP_NAME,
            icon   = _make_icon_image(),
            title  = APP_NAME,
            menu   = pystray.Menu(
                pystray.MenuItem("Settings", self._open_settings),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Exit", self._exit),
            ),
        )

    def _open_settings(self, icon=None, item=None):
        """Schedule the settings window on the tkinter main thread."""
        self._tk_root.after(0, lambda: open_settings_window(
            self.settings,
            on_save=lambda: print("[Settings] Saved:", self.settings),
        ))

    def _exit(self, icon=None, item=None):
        self.monitor.stop()
        self._icon.stop()
        self._tk_root.after(0, self._tk_root.destroy)

    def run(self):
        self.monitor.start()
        # Run pystray in a daemon thread; tkinter mainloop stays on main thread
        tray_thread = threading.Thread(target=self._icon.run, daemon=True)
        tray_thread.start()
        # Tkinter mainloop (needed to pump settings window events)
        try:
            self._tk_root.mainloop()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    app = TrayApp()
    app.run()


if __name__ == "__main__":
    main()
