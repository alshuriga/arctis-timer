# Arctis 7 Timer

A lightweight Windows system tray utility that automatically manages the **SteelSeries Arctis 7** inactivity timer based on real-time audio activity and user presence — replacing the unreliable auto-off feature in SteelSeries GG.

## Key Features

- **Smart Audio Monitoring**: Detects real-time audio peaks to keep the headset active.
- **AFK / Idle Detection**: Automatically triggers the auto-off timer if you are away (no keyboard/mouse input), even if music is playing.
- **Customizable Modes**: Choose to monitor only silence, only idle time, or both.
- **Live Peak Meter**: Calibrate your silence threshold visually with a real-time audio meter in the settings.
- **Premium Light UI**: A clean, modern white theme with intuitive controls and help tooltips.
- **Direct HID Control**: Bypasses the SteelSeries software to send hardware commands directly to the USB receiver.

## How It Works

| State Trigger | Action |
|---|---|
| Active audio / User activity | Sets headset timer to **Never** (stays on) |
| Silence / Idle detected (per settings) | Sets headset timer to **X minutes** (triggers auto-off) |

## Usage

- **Right-click tray icon → Settings**: Configure all parameters.
- **Live Peak Meter**: The red line shows your current threshold. Adjust it until background noise stays to the left of the line.
- **Notifications**: Enable popups to know exactly when the headset is scheduled to turn off.

### Configuration Parameters

| Parameter | Default | Description |
|---|---|---|
| Detection Mode | `Both` | Choose to monitor silence, idle time, or both. |
| AFK Timeout | `10` min | Minutes of no keyboard/mouse input before triggering auto-off. |
| Inactive Timer | `1` min | Hardware timer set on the headset when triggered (1–90). |
| Silence Duration | `30` sec | Delay before silence triggers the inactive timer. |
| Silence Threshold | `0.001` | Sensitivity for audio detection (tuned via Peak Meter). |

## Installation (from EXE)

1. Download `ArctisTimer.exe` from the latest [Release](https://github.com/alshuriga/arctis-timer/releases).
2. Run it — it will appear in your system tray.
3. **Auto-start**: To start with Windows, place a shortcut to `ArctisTimer.exe` in:
   `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup`

## Building from Source

### 1. Clone the repo

```bash
git clone https://github.com/alshuriga/arctis-timer.git
cd arctis-timer
```

### 2. Set up environment

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Build EXE

```bash
pyinstaller ArctisTimer.spec
```
The output will be in `dist\ArctisTimer.exe`.

## Technical Notes

- Uses **Windows HID via `CreateFile` + `WriteFile`** on the SteelSeries vendor-specific control collection (`0xFF43` / `0x0202`).
- Command ID `0x06` (Report ID), Sub-command `0x51` (Set Timer).
- Audio monitoring uses **pycaw** (Core Audio API).
- Idle detection uses **Win32 User32 `GetLastInputInfo`**.

## Supported Devices

- Arctis 7 (original / 2019)
- Arctis Pro (2019 / GameDAC)
- Most 0x1038 SteelSeries wireless receivers using usage page 0xFF43.

## License

MIT
