"""
Arctis 7 Auto-Off Timer
Monitors Windows audio and controls the SteelSeries Arctis 7 inactivity timer
via USB HID. Runs as a system tray application.
"""

import ctypes
import ctypes.wintypes as wintypes
import json
import os
import threading
import time
import tkinter as tk
from tkinter import ttk

import hid
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
    "notifications_enabled": True,  # show Windows toast notifications on state change
}


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
        while not self._stop_event.is_set():
            s = self.settings
            threshold       = s["silence_threshold"]
            silence_needed  = s["silence_duration_seconds"]
            active_needed   = s["active_duration_seconds"]
            inactive_min    = s["inactive_timer_minutes"]
            poll            = 2.0

            playing = self._peek() > threshold
            now     = time.time()

            if playing:
                self._silence_start = None
                if self._active_start is None:
                    self._active_start = now
                if self._state != "ACTIVE" and (now - self._active_start) >= active_needed:
                    ts = time.strftime("%H:%M:%S")
                    print(f"[{ts}] Audio detected → timer OFF (never)")
                    if self.arctis.set_inactivity_timer(0):
                        self._state = "ACTIVE"
                        self._notify("🎧 Audio detected", "Headset will stay on.")
            else:
                self._active_start = None
                if self._silence_start is None:
                    self._silence_start = now
                elapsed = now - self._silence_start
                if self._state != "INACTIVE" and elapsed >= silence_needed:
                    ts = time.strftime("%H:%M:%S")
                    print(f"[{ts}] Silence {elapsed:.0f}s → timer={inactive_min}min")
                    if self.arctis.set_inactivity_timer(inactive_min):
                        self._state = "INACTIVE"
                        self._notify("🔇 Silence detected",
                                     f"Auto-off in {inactive_min} min.")

            self._stop_event.wait(poll)

    def _notify(self, title: str, message: str):
        if not self.settings.get("notifications_enabled", True):
            return
        try:
            toast(APP_NAME, f"{title} — {message}",
                  audio={"src": "ms-winsoundevent:Notification.Default", "silent": "true"})
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Settings window (Tkinter)
# ---------------------------------------------------------------------------

ACCENT   = "#10b981"   # emerald green
BG_DARK  = "#1e1e2e"
BG_CARD  = "#2a2a3d"
FG_TEXT  = "#e2e8f0"
FG_DIM   = "#94a3b8"


def open_settings_window(settings: dict, on_save):
    """Open the settings dialog on the calling thread (must be main thread)."""
    win = tk.Toplevel()
    win.title(f"{APP_NAME} — Settings")
    win.configure(bg=BG_DARK)
    win.resizable(False, False)
    win.attributes("-topmost", True)

    # Icon
    try:
        win.iconbitmap(default="")
    except Exception:
        pass

    def label(parent, text, **kw):
        return tk.Label(parent, text=text, bg=BG_DARK, fg=FG_DIM,
                        font=("Segoe UI", 9), **kw)

    def heading(parent, text):
        tk.Label(parent, text=text, bg=BG_DARK, fg=FG_TEXT,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(12, 2))

    def row(parent, lbl_text, var, from_, to_, unit=""):
        frame = tk.Frame(parent, bg=BG_CARD, pady=6, padx=10)
        frame.pack(fill="x", pady=3)
        tk.Label(frame, text=lbl_text, bg=BG_CARD, fg=FG_TEXT,
                 font=("Segoe UI", 9), width=30, anchor="w").pack(side="left")
        sb = ttk.Spinbox(frame, from_=from_, to=to_, textvariable=var,
                         width=6, font=("Segoe UI", 9))
        sb.pack(side="left", padx=(0, 6))
        if unit:
            tk.Label(frame, text=unit, bg=BG_CARD, fg=FG_DIM,
                     font=("Segoe UI", 9)).pack(side="left")

    # Variables
    v_inactive   = tk.IntVar(value=settings["inactive_timer_minutes"])
    v_silence_s  = tk.IntVar(value=settings["silence_duration_seconds"])
    v_active_s   = tk.IntVar(value=settings["active_duration_seconds"])
    v_threshold  = tk.DoubleVar(value=settings["silence_threshold"])
    v_notifs     = tk.BooleanVar(value=settings.get("notifications_enabled", True))

    pad = tk.Frame(win, bg=BG_DARK, padx=20, pady=16)
    pad.pack(fill="both", expand=True)

    heading(pad, "Auto-Off")
    row(pad, "Inactive timer (after silence)", v_inactive, 1, 90, "min")

    heading(pad, "Silence Detection")
    row(pad, "Silence detect duration",        v_silence_s, 5, 3600, "sec")
    row(pad, "Audio duration to cancel timer", v_active_s,  1, 60,   "sec")
    row(pad, "Silence threshold (peak level)", v_threshold, 0.0001, 0.1, "")

    heading(pad, "Notifications")
    notif_frame = tk.Frame(pad, bg=BG_CARD, pady=6, padx=10)
    notif_frame.pack(fill="x", pady=3)
    ttk.Checkbutton(
        notif_frame,
        text="Show toast notifications on state change",
        variable=v_notifs,
    ).pack(anchor="w")

    def on_save_click():
        settings["inactive_timer_minutes"]    = v_inactive.get()
        settings["silence_duration_seconds"]  = v_silence_s.get()
        settings["active_duration_seconds"]   = v_active_s.get()
        settings["silence_threshold"]         = round(v_threshold.get(), 6)
        settings["notifications_enabled"]     = bool(v_notifs.get())
        save_settings(settings)
        if on_save:
            on_save()
        win.destroy()

    btn_frame = tk.Frame(win, bg=BG_DARK)
    btn_frame.pack(fill="x", padx=20, pady=(0, 16))

    save_btn = tk.Button(
        btn_frame, text="Save", command=on_save_click,
        bg=ACCENT, fg="white", font=("Segoe UI", 9, "bold"),
        relief="flat", padx=16, pady=6, cursor="hand2",
        activebackground="#059669", activeforeground="white",
    )
    save_btn.pack(side="right")
    tk.Button(
        btn_frame, text="Cancel", command=win.destroy,
        bg=BG_CARD, fg=FG_DIM, font=("Segoe UI", 9),
        relief="flat", padx=16, pady=6, cursor="hand2",
    ).pack(side="right", padx=(0, 8))

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
