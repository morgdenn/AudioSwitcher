# AudioSwitcher

A lightweight Windows system tray utility for quickly switching between audio output devices.

## Features

- **Left-click** the tray icon to cycle through available audio output devices
- **Right-click** for a context menu to select a specific device, refresh the device list, or exit
- Toast notifications confirm each device switch
- Tray icon tooltip shows the currently active device
- Checkmark in the menu indicates the active device
- No admin rights required

## Requirements

- Windows 10/11
- Python 3.10+ (for running from source or building)

## Usage

### Standalone executable

Download or build `AudioSwitcher.exe` and run it. The app appears in the system tray.

**To run on startup**, create a shortcut to `AudioSwitcher.exe` and place it in:

```
%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
```

### Run from source

```bash
pip install -r requirements.txt
python audio_switcher.py
```

## Building

Run the build script from the project root:

```bat
build.bat
```

This installs all dependencies and produces `dist\AudioSwitcher.exe` — a self-contained ~15 MB executable with no installer needed.

## Dependencies

| Package | Purpose |
|---|---|
| `comtypes` | Direct access to Windows Core Audio COM interfaces |
| `pystray` | System tray icon and menu |
| `Pillow` | Tray icon image generation |
| `pyinstaller` | Builds the standalone executable |

## How it works

AudioSwitcher uses the Windows Core Audio API via COM (`IMMDeviceEnumerator`, `IPolicyConfig`) to enumerate active output devices and set the default playback device across all audio roles (console, multimedia, and communications) simultaneously.
