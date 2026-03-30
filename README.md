# iPhone Photo Backup

Transfers photos from an iPhone to Windows over USB with resume support. Handles freezes and disconnections by timing out per-file and retrying — safe to interrupt and re-run at any time.

## Features

- Resume interrupted transfers (already-copied files are skipped)
- Automatic retry (up to 3×) for files that fail or time out
- Disk space check before starting
- Handles iPhone disconnection mid-transfer gracefully
- Colorized console output with progress bar and ETA
- Detailed transfer log saved alongside your photos

## Requirements

- Windows 10/11
- Python 3.9+
- iPhone connected via USB with "Trust This Computer" accepted

## Installation

```bat
pip install -r requirements.txt
```

Dependencies: `pywin32`, `colorama`

## Usage

### Option A — Double-click (simplest)

Run `run.bat`. Photos are saved to `%USERPROFILE%\Pictures\iPhone Photos`.

### Option B — Command line

```bat
# Default destination (Pictures\iPhone Photos)
python iphone_downloader.py

# Custom destination
python iphone_downloader.py D:\MyPhotos

# Clear saved state and start fresh
python iphone_downloader.py --reset

# Custom destination + reset
python iphone_downloader.py D:\MyPhotos --reset
```

### Option C — Standalone .exe (no Python needed)

Build a self-contained executable:

```bat
build.bat
```

Output: `dist\iPhone Backup.exe` — copy to any Windows machine and run directly.

## How it works

1. Connects to the iPhone via Windows Shell MTP (no iTunes required)
2. Enumerates all photo albums and files
3. Copies each file through a temp folder, polling until the file size stabilises before moving it to its final location
4. Saves progress to `transfer_state.json` so interrupted runs resume where they left off

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| "iPhone not found" | Unlock the phone, check USB, tap **Trust** on the iPhone prompt |
| Photos missing after scan | Unplug, relock/unlock, reconnect and retry |
| File keeps failing | Run again — it retries up to 3× per file |
| Want to re-copy everything | Run with `--reset` to clear saved state |

Transfer logs are written to the destination folder as `transfer_YYYYMMDD_HHMMSS.log`.
