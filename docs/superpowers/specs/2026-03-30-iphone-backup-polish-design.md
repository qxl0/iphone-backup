# iPhone Backup — Polished Console UI Design

**Date:** 2026-03-30
**Goal:** Make `iphone_downloader.py` usable by non-technical Windows users with no setup required.

---

## Overview

Rewrite the console output of `iphone_downloader.py` to be clear, friendly, and visually polished. Package it as a standalone `.exe` so users double-click and it just works — no Python, no terminal knowledge required.

---

## File Structure

```
iphone_downloader.py    ← modified: polished UI, folder prompt, elapsed time
build.bat               ← new: one-click PyInstaller packager
requirements.txt        ← new: colorama, pywin32
```

`copy_folder.py` and `diagnose.py` remain as developer tools and are not packaged.

---

## Console UI

Uses `colorama` for cross-Windows color support. All output goes through a thin UI layer that formats messages consistently.

### States

**1. Launch — folder prompt**
```
╔══════════════════════════════╗
║   📱  iPhone Photo Backup    ║
╚══════════════════════════════╝

Save photos to:
  C:\Users\<name>\Pictures\iPhone Photos

Press Enter to use this folder,
or type a new path and press Enter:
> _
```

**2. Connecting**
```
⟳ Looking for your iPhone...
```

**3. Copying (progress bar redraws in place using `\r` — no scrolling)**
```
✓ iPhone connected — 18,332 photos found

Progress:
  [████████████████░░░░░░]  72%
  13,199 / 18,332 copied  •  ~6 min left

  ✓ IMG_4667.HEIC
  ✓ IMG_4668.MOV
  → IMG_4669.HEIC ...
```

**4. Done**
```
✓ Backup complete!
  18,332 photos copied
  3 already up to date (skipped)
  Completed in 1 hour 42 minutes

Saved to: C:\Users\<name>\Pictures\iPhone Photos

Press any key to close...
```

### Colors
- Cyan: header banner
- Green: success messages (`✓`)
- Yellow: prompts and in-progress indicators
- Red: errors (`✗`)
- Dark grey: completed file lines (dimmed to reduce noise)
- White: current file being copied

---

## Folder Selection

On every launch, show the default destination and let the user accept or override:
- Press **Enter** → use default (`Pictures\iPhone Photos`)
- Type a path + Enter → use that path
- The chosen path is used for that session only (not persisted)

---

## Error Handling

All errors shown in plain English. No stack traces visible to users (logged to file only).

| Scenario | User message | Action available |
|---|---|---|
| iPhone not found | "Make sure it's plugged in, unlocked, and tap Trust on your iPhone screen" | Press R to retry |
| iPhone disconnected mid-transfer | "iPhone disconnected. Your progress has been saved." | Press R to retry |
| Not enough disk space | Shows available vs estimated needed. "Free up space and run again — already-copied photos will be skipped." | Any key to quit |
| No new photos | "No new photos found — you're all backed up!" | Any key to quit |
| Destination folder not writable | "Can't write to [folder]. Try choosing a different folder." | Any key to quit |

Resume support is preserved: re-running always skips files already successfully transferred.

---

## Elapsed Time

Timer starts when copying begins (after iPhone is found). On completion:
```
Completed in 1 hour 42 minutes
```
Format: hours and minutes if ≥ 1 hour, otherwise just minutes.

---

## Packaging

### build.bat
```bat
pyinstaller --onefile --console --name "iPhone Backup" iphone_downloader.py
```
Output: `dist\iPhone Backup.exe` — single file, ~15 MB, no Python installation needed.

### requirements.txt
```
pywin32
colorama
pyinstaller
```

---

## Out of Scope

- Windows toast notifications (not needed)
- Persistent settings file (folder choice is per-session)
- GUI window (console is sufficient)
- Automatic launch on iPhone connection
