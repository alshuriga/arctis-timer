# Arctis 7 Timer

A lightweight Windows system tray utility that automatically manages the **SteelSeries Arctis 7** inactivity timer based on real-time audio activity — replacing the unreliable auto-off feature in SteelSeries GG.

## How It Works

| Audio state | Action |
|---|---|
| Audio playing for ≥ 2 seconds | Sets headset timer to **Never** (stays on) |
| Silence for ≥ 30 seconds | Sets headset timer to **1 minute** (headset turns off after ~1 min) |

Both values are configurable via the tray Settings window.

## Requirements

- Windows 10/11
- SteelSeries Arctis 7 (original `0x1260` or 2019 edition `0x12AD`) — USB dongle plugged in
- Python 3.10+ (for development only; end-users run the EXE)

## Installation (from EXE)

1. Download `ArctisTimer.exe` from [Releases](../../releases).
2. Run it — it will appear in your system tray.
3. To start it automatically with Windows, place a shortcut to `ArctisTimer.exe` in:
   ```
   %APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
   ```

## Usage

- **Right-click tray icon → Settings** — configure timers and silence threshold.
- **Right-click tray icon → Exit** — quit the app.
- Settings are saved to `%APPDATA%\ArctisTimer\settings.json`.

### Settings

| Setting | Default | Description |
|---|---|---|
| Inactive timer | `1` min | Minutes before headset auto-offs when silence is detected (1–90) |
| Silence duration | `30` sec | How long audio must be silent before activating the timer |
| Audio duration | `2` sec | How long audio must play before cancelling the timer |
| Silence threshold | `0.001` | Peak audio level below which audio is considered silent |

## Building from Source

### 1. Clone the repo

```bash
git clone https://github.com/yourname/arctis-timer.git
cd arctis-timer
```

### 2. Set up Python environment

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Test (dev run)

```bash
python arctis_auto_timer.py
```

### 4. Build EXE

```bash
pyinstaller ArctisTimer.spec
```

The output will be in `dist\ArctisTimer.exe` — a **single self-contained executable**, no Python install required.

> **Note:** PyInstaller bundles all dependencies. The first build may take a minute.
> If Windows Defender flags the EXE, add an exclusion for the `dist\` folder.

## Supported Devices

| Device | PID |
|---|---|
| Arctis 7 (original) | `0x1260` |
| Arctis 7 (2019) | `0x12AD` |
| Arctis Pro (2019) | `0x1252` |
| Arctis Pro GameDAC | `0x1280` |

## Technical Notes

- Uses **Windows HID via `CreateFile` + `WriteFile`** (not hidapi's `write()`, which fails on this device class under Windows).
- Targets HID interface 5, usage page `0xFF43`, usage `0x0202` — the SteelSeries vendor-specific control collection.
- Command format: `[0x06, 0x51, minutes, 0x00 * 28]` (31 bytes total, report ID `0x06`).
- Audio monitoring uses **pycaw** (Windows Core Audio API).

## License

MIT
