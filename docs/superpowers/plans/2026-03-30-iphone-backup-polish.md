# iPhone Backup Polish Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Polish `iphone_downloader.py` into a standalone double-click `.exe` that non-technical Windows users can run without any setup.

**Architecture:** Modify `iphone_downloader.py` in-place — add a colorama UI layer, replace console logging with colored print statements, add a folder prompt and elapsed time. PyInstaller packages everything into a single `.exe`.

**Tech Stack:** Python 3, colorama, pywin32, pyinstaller, msvcrt (stdlib), pytest

---

## File Map

| File | Action | Responsibility |
|---|---|---|
| `iphone_downloader.py` | Modify | All logic + UI — add colorama helpers, folder prompt, progress bar, plain English errors |
| `tests/test_ui.py` | Create | Unit tests for pure UI helper functions |
| `requirements.txt` | Create | Dev + build dependencies |
| `build.bat` | Create | One-click PyInstaller packager |

---

## Task 1: requirements.txt + colorama bootstrap

**Files:**
- Create: `requirements.txt`
- Modify: `iphone_downloader.py` (imports section, top of file)

- [ ] **Step 1: Create requirements.txt**

```
pywin32
colorama
pyinstaller
pytest
```

- [ ] **Step 2: Install dependencies**

```bash
pip install colorama pytest
```

Expected: installs without errors.

- [ ] **Step 3: Add colorama imports and init to iphone_downloader.py**

At the top of `iphone_downloader.py`, after the existing imports, add:

```python
import ctypes
import msvcrt
import colorama
from colorama import Fore, Style

# Enable UTF-8 output and ANSI colors in Windows console
try:
    ctypes.windll.kernel32.SetConsoleOutputCP(65001)
    ctypes.windll.kernel32.SetConsoleCP(65001)
except Exception:
    pass
colorama.init()

# ── Color shortcuts ────────────────────────────────────────────────────────────
CYAN   = Fore.CYAN
GREEN  = Fore.GREEN
YELLOW = Fore.YELLOW
RED    = Fore.RED
DIM    = Style.DIM
BRIGHT = Style.BRIGHT
RESET  = Style.RESET_ALL
```

- [ ] **Step 4: Verify import works**

```bash
python -c "import iphone_downloader; print('OK')"
```

Expected: `OK`

- [ ] **Step 5: Commit**

```bash
git add requirements.txt iphone_downloader.py
git commit -m "chore: add colorama dependency and console UTF-8 bootstrap"
```

---

## Task 2: Pure UI helper functions + tests

These are pure functions with no side effects — easy to test.

**Files:**
- Create: `tests/__init__.py` (empty)
- Create: `tests/test_ui.py`
- Modify: `iphone_downloader.py` (add three functions after color constants)

- [ ] **Step 1: Create tests directory**

```bash
mkdir tests
touch tests/__init__.py
```

- [ ] **Step 2: Write failing tests**

Create `tests/test_ui.py`:

```python
import sys
sys.path.insert(0, ".")
from iphone_downloader import format_elapsed, format_bar, format_eta


class TestFormatElapsed:
    def test_less_than_a_minute(self):
        assert format_elapsed(45) == "less than a minute"

    def test_one_minute(self):
        assert format_elapsed(60) == "1 minute"

    def test_plural_minutes(self):
        assert format_elapsed(180) == "3 minutes"

    def test_one_hour_zero_minutes(self):
        assert format_elapsed(3600) == "1 hour 0 minutes"

    def test_one_hour_one_minute(self):
        assert format_elapsed(3660) == "1 hour 1 minute"

    def test_hours_and_minutes(self):
        assert format_elapsed(6120) == "1 hour 42 minutes"

    def test_two_hours(self):
        assert format_elapsed(7200) == "2 hours 0 minutes"


class TestFormatBar:
    def test_empty(self):
        assert format_bar(0, 100, width=10) == "[░░░░░░░░░░]"

    def test_full(self):
        assert format_bar(100, 100, width=10) == "[██████████]"

    def test_half(self):
        assert format_bar(50, 100, width=10) == "[█████░░░░░]"

    def test_zero_total(self):
        assert format_bar(0, 0, width=4) == "[░░░░]"


class TestFormatEta:
    def test_no_progress_yet(self):
        result = format_eta(0, 100, elapsed=10.0)
        assert result == "estimating..."

    def test_zero_elapsed(self):
        result = format_eta(10, 100, elapsed=0.0)
        assert result == "estimating..."

    def test_less_than_one_minute(self):
        # 50 done in 10s → rate 5/s → 50 remaining → 10s left
        result = format_eta(50, 100, elapsed=10.0)
        assert result == "< 1 min left"

    def test_minutes(self):
        # 10 done in 60s → rate 1/6 per s → 540 remaining → 540s = 9 min
        result = format_eta(10, 100, elapsed=60.0)
        assert result == "~9 min left"

    def test_hours(self):
        # 10 done in 600s → rate 1/60 per s → 90 remaining → 5400s = 90 min = 1h 30m
        result = format_eta(10, 100, elapsed=600.0)
        assert result == "~1h 30m left"
```

- [ ] **Step 3: Run tests — verify they fail**

```bash
pytest tests/test_ui.py -v
```

Expected: `ImportError` or `AttributeError` — functions don't exist yet.

- [ ] **Step 4: Add the three functions to iphone_downloader.py**

Add after the color constants block:

```python
# ── UI helper functions ────────────────────────────────────────────────────────

def format_elapsed(seconds: int) -> str:
    """Return human-readable elapsed time. E.g. '1 hour 42 minutes'."""
    if seconds < 60:
        return "less than a minute"
    total_minutes = seconds // 60
    hours = total_minutes // 60
    minutes = total_minutes % 60
    if hours > 0:
        h = f"{hours} hour{'s' if hours != 1 else ''}"
        m = f"{minutes} minute{'s' if minutes != 1 else ''}"
        return f"{h} {m}"
    return f"{minutes} minute{'s' if minutes != 1 else ''}"


def format_bar(done: int, total: int, width: int = 22) -> str:
    """Return a text progress bar. E.g. '[████░░░░░░]'."""
    if total == 0:
        return "[" + "░" * width + "]"
    filled = int(width * done / total)
    return "[" + "█" * filled + "░" * (width - filled) + "]"


def format_eta(done: int, total: int, elapsed: float) -> str:
    """Return estimated time remaining. E.g. '~6 min left'."""
    if done == 0 or elapsed == 0:
        return "estimating..."
    rate = done / elapsed  # files per second
    remaining_seconds = (total - done) / rate
    if remaining_seconds < 60:
        return "< 1 min left"
    remaining_minutes = int(remaining_seconds // 60)
    if remaining_minutes >= 60:
        h = remaining_minutes // 60
        m = remaining_minutes % 60
        return f"~{h}h {m}m left"
    return f"~{remaining_minutes} min left"
```

- [ ] **Step 5: Run tests — verify they pass**

```bash
pytest tests/test_ui.py -v
```

Expected: all 15 tests PASS.

- [ ] **Step 6: Commit**

```bash
git add tests/ iphone_downloader.py
git commit -m "feat: add format_elapsed, format_bar, format_eta helpers with tests"
```

---

## Task 3: Banner, folder prompt, press_any_key

**Files:**
- Modify: `iphone_downloader.py` (add three functions after UI helpers)

- [ ] **Step 1: Add print_banner, prompt_dest, press_any_key to iphone_downloader.py**

Add after the `format_eta` function:

```python
def print_banner() -> None:
    """Print the app header banner."""
    print(f"\n{CYAN}{BRIGHT}╔══════════════════════════════╗{RESET}")
    print(f"{CYAN}{BRIGHT}║   \U0001f4f1  iPhone Photo Backup    \u2551{RESET}")
    print(f"{CYAN}{BRIGHT}╚══════════════════════════════╝{RESET}\n")


def prompt_dest() -> Path:
    """Show default destination, let user override. Returns chosen Path."""
    default = Path.home() / "Pictures" / "iPhone Photos"
    print(f"Save photos to:")
    print(f"  {YELLOW}{default}{RESET}\n")
    print(f"Press {GREEN}Enter{RESET} to use this folder,")
    print(f"or type a new path and press Enter:")
    print(f"> ", end="", flush=True)
    user_input = input().strip()
    if user_input:
        chosen = Path(user_input)
    else:
        chosen = default
    print()
    return chosen


def press_any_key(msg: str = "Press any key to close...") -> None:
    """Wait for a single keypress."""
    print(f"\n{DIM}{msg}{RESET}", flush=True)
    msvcrt.getwch()
```

- [ ] **Step 2: Verify no syntax errors**

```bash
python -c "import iphone_downloader; print('OK')"
```

Expected: `OK`

- [ ] **Step 3: Commit**

```bash
git add iphone_downloader.py
git commit -m "feat: add print_banner, prompt_dest, press_any_key UI functions"
```

---

## Task 4: Silence the console StreamHandler

Currently `setup_logging` prints INFO messages to console via a StreamHandler. We're replacing console output with colorama — keep the StreamHandler for WARNING+ only (so genuine errors still surface).

**Files:**
- Modify: `iphone_downloader.py` (`setup_logging` function)

- [ ] **Step 1: Update setup_logging in iphone_downloader.py**

Replace the existing `setup_logging` function with:

```python
def setup_logging(dest_dir: Path) -> logging.Logger:
    dest_dir.mkdir(parents=True, exist_ok=True)
    log_file = dest_dir / f"transfer_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

    logger = logging.getLogger("iphone_downloader")
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")

    # File handler — captures everything (DEBUG+) for troubleshooting
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    # Console handler — WARNING+ only; UI uses colorama print statements
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.WARNING)
    ch.setFormatter(fmt)
    ch.stream.reconfigure(encoding="utf-8", errors="replace")
    logger.addHandler(ch)

    return logger
```

- [ ] **Step 2: Verify no syntax errors**

```bash
python -c "import iphone_downloader; print('OK')"
```

Expected: `OK`

- [ ] **Step 3: Commit**

```bash
git add iphone_downloader.py
git commit -m "refactor: suppress INFO console logging — UI uses colorama prints"
```

---

## Task 5: Rewrite transfer_photos with progress bar and error handling

This is the main UI rewrite. Pre-enumerate all files, then copy with colorama progress display, disk space check, and iPhone disconnection detection.

**Files:**
- Modify: `iphone_downloader.py` (`transfer_photos` function — full replacement)

- [ ] **Step 1: Replace transfer_photos in iphone_downloader.py**

Replace the entire `transfer_photos` function with:

```python
def transfer_photos(dest_dir: Path, log, reset: bool = False) -> None:
    state = TransferState(dest_dir / "transfer_state.json")
    if reset:
        state.reset()
        log.info("State reset — starting fresh.")

    pythoncom.CoInitialize()
    try:
        shell = _make_shell()

        # ── Find iPhone (with retry) ───────────────────────────────────────────
        print(f"{YELLOW}⟳ Looking for your iPhone...{RESET}", flush=True)
        while True:
            iphone = find_iphone(shell, log)
            if iphone:
                break
            print(f"\r{RED}✗ iPhone not found.{RESET}                    ")
            print(f"\nMake sure your iPhone is:")
            print(f"  \u2022 Plugged in via USB")
            print(f"  \u2022 Turned on and unlocked")
            print(f"  \u2022 Showing \"Trust This Computer?\" \u2014 tap {GREEN}Trust{RESET}\n")
            print(f"Press {GREEN}R{RESET} to retry, or any other key to quit...")
            key = msvcrt.getwch()
            if key.lower() != "r":
                return
            print(f"\n{YELLOW}⟳ Looking for your iPhone...{RESET}", flush=True)
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            pythoncom.CoInitialize()
            shell = _make_shell()

        print(f"\r{GREEN}✓ iPhone found!{RESET}                         ")

        # ── Find photo root ────────────────────────────────────────────────────
        print(f"{YELLOW}⟳ Scanning photos...{RESET}", flush=True)
        dcim = find_photos_root(shell, iphone, log)
        if dcim is None:
            print(f"\r{RED}✗ Could not find photos on your iPhone.{RESET}")
            print(f"\nTry unplugging, relocking, unlocking, and reconnecting.")
            press_any_key()
            return

        # ── Pre-enumerate all files ────────────────────────────────────────────
        subfolders = _folder_items(shell, dcim)
        all_items: list[tuple[str, Path, object, str]] = []
        for subfolder_item in subfolders:
            sub_name = subfolder_item.Name
            dest_sub = dest_dir / sub_name
            files = _folder_items(shell, subfolder_item)
            if not files:
                files = _refetch_album(sub_name, log)
            for file_item in files:
                all_items.append((sub_name, dest_sub, file_item, file_item.Name))

        total = len(all_items)
        print(f"\r{GREEN}✓ Found {total:,} photos.{RESET}                ")

        if total == 0:
            print(f"\n{GREEN}✓ No new photos found \u2014 you\u2019re all backed up!{RESET}")
            press_any_key()
            return

        # ── Disk space check ───────────────────────────────────────────────────
        try:
            free = shutil.disk_usage(dest_dir).free
            already_done = sum(
                1 for sub, _, _, fname in all_items
                if state.is_done(f"{sub}/{fname}")
            )
            remaining_count = total - already_done
            estimated_needed = remaining_count * 3_000_000  # ~3 MB avg
            if estimated_needed > free:
                print(f"\n{RED}✗ Not enough disk space.{RESET}")
                print(f"  Available : {free / 1e9:.1f} GB")
                print(f"  Estimated : {estimated_needed / 1e9:.1f} GB needed")
                print(f"\nFree up space and run again \u2014 already-copied photos will be skipped.")
                press_any_key()
                return
        except Exception:
            pass  # best-effort check

        # ── Copy loop ──────────────────────────────────────────────────────────
        print()
        copied = skipped = failed = 0
        start_time = time.time()
        last_sub: str | None = None

        for sub_name, dest_sub, file_item, filename in all_items:
            dest_sub.mkdir(exist_ok=True)
            key       = f"{sub_name}/{filename}"
            dest_path = dest_sub / filename
            done_so_far = copied + skipped + failed

            # Album header when subfolder changes
            if sub_name != last_sub:
                if last_sub is not None:
                    print()
                print(f"  {DIM}[{sub_name}]{RESET}")
                last_sub = sub_name

            # Skip already-done files
            if state.is_done(key):
                skipped += 1
                continue
            if dest_path.exists() and dest_path.stat().st_size > 0:
                state.mark_done(key, dest_path.stat().st_size)
                skipped += 1
                continue

            retries = state.get_retries(key)
            if retries >= MAX_RETRIES:
                failed += 1
                log.warning(f"Max retries reached: {key}")
                continue

            # Print in-progress line (no newline — overwritten on completion)
            elapsed = time.time() - start_time
            bar = format_bar(done_so_far, total)
            eta = format_eta(done_so_far, total, elapsed)
            pct = int(done_so_far / total * 100) if total else 0
            progress = f"{DIM}{bar} {pct}%  {eta}{RESET}"
            print(
                f"\r  {YELLOW}\u2192{RESET} {filename}  {progress}    ",
                end="", flush=True
            )

            ok = copy_with_timeout(
                shell, file_item, dest_sub, filename, TIMEOUT_PER_FILE, log
            )

            if ok:
                size = dest_path.stat().st_size if dest_path.exists() else 0
                state.mark_done(key, size)
                copied += 1
                # Overwrite in-progress line with dimmed done line
                print(f"\r  {DIM}{GREEN}\u2713{RESET}{DIM} {filename}{RESET}" + " " * 40)
                log.debug(f"\u2713 {key} ({size:,} bytes)")
            else:
                # Check if iPhone disconnected
                if find_iphone(shell, log) is None:
                    print(f"\r\n{YELLOW}iPhone disconnected. Your progress has been saved.{RESET}")
                    print(f"\nPlug it back in and run the backup again to continue.")
                    press_any_key()
                    return
                state.increment_retries(key)
                failed += 1
                print(f"\r  {RED}\u2717 {filename}{RESET}" + " " * 40)
                log.warning(f"\u2717 {key} \u2014 retry {retries + 1}/{MAX_RETRIES}")

        # ── Summary ────────────────────────────────────────────────────────────
        elapsed_total = int(time.time() - start_time)
        print()

        if copied == 0 and skipped == total:
            print(f"\n{GREEN}\u2713 No new photos found \u2014 you\u2019re all backed up!{RESET}")
        else:
            print(f"\n{GREEN}{BRIGHT}\u2713 Backup complete!{RESET}")
            if copied:
                print(f"  {copied:,} photos copied")
            if skipped:
                print(f"  {skipped:,} already up to date (skipped)")
            if failed:
                print(f"  {RED}{failed:,} failed{RESET} \u2014 run again to retry")
            print(f"  Completed in {format_elapsed(elapsed_total)}")

        print(f"\n{DIM}Saved to: {dest_dir}{RESET}")
        if failed:
            log.warning(f"{failed} file(s) failed after {MAX_RETRIES} retries.")

    except OSError as exc:
        msg = str(exc).lower()
        if "access" in msg or "denied" in msg or "permission" in msg:
            print(f"\n{RED}\u2717 Can't write to {dest_dir}.{RESET}")
            print(f"  Try choosing a different folder.")
            log.error(f"Permission denied: {exc}")
        else:
            print(f"\n{RED}\u2717 Unexpected error. Details saved to the log file.{RESET}")
            log.exception("Unexpected OSError")
        press_any_key()
    finally:
        pythoncom.CoUninitialize()
```

- [ ] **Step 2: Verify no syntax errors**

```bash
python -c "import iphone_downloader; print('OK')"
```

Expected: `OK`

- [ ] **Step 3: Run existing tests still pass**

```bash
pytest tests/test_ui.py -v
```

Expected: all 15 tests PASS.

- [ ] **Step 4: Commit**

```bash
git add iphone_downloader.py
git commit -m "feat: rewrite transfer_photos with colorama progress bar and plain-English error handling"
```

---

## Task 6: Rewrite main()

**Files:**
- Modify: `iphone_downloader.py` (`main` function — full replacement)

- [ ] **Step 1: Replace main() in iphone_downloader.py**

Replace the entire `main` function with:

```python
def main() -> None:
    args  = sys.argv[1:]
    reset = "--reset" in args
    args  = [a for a in args if a != "--reset"]

    print_banner()

    # Destination: use CLI arg if provided, otherwise interactive prompt
    if args:
        dest_dir = Path(args[0])
        print(f"Destination: {YELLOW}{dest_dir}{RESET}\n")
    else:
        dest_dir = prompt_dest()

    log = setup_logging(dest_dir)
    log.debug(f"Destination: {dest_dir}  reset={reset}")

    try:
        transfer_photos(dest_dir, log, reset=reset)
    except KeyboardInterrupt:
        print(f"\n\n{YELLOW}Stopped. Progress saved \u2014 run again to resume.{RESET}")
        press_any_key()
    except Exception:
        log.exception("Unexpected error")
        print(f"\n{RED}\u2717 Something went wrong. Details saved to the log file in:{RESET}")
        print(f"  {dest_dir}")
        press_any_key()
```

- [ ] **Step 2: Verify no syntax errors**

```bash
python -c "import iphone_downloader; print('OK')"
```

Expected: `OK`

- [ ] **Step 3: Run all tests**

```bash
pytest tests/ -v
```

Expected: all 15 tests PASS.

- [ ] **Step 4: Smoke test (iPhone not needed — just check startup)**

```bash
python iphone_downloader.py --help 2>&1 || python iphone_downloader.py
```

You should see the banner and folder prompt. Press Ctrl+C to exit — it should print "Stopped. Progress saved." cleanly.

- [ ] **Step 5: Commit**

```bash
git add iphone_downloader.py
git commit -m "feat: rewrite main() with banner, folder prompt, and clean interrupt handling"
```

---

## Task 7: build.bat

**Files:**
- Create: `build.bat`

- [ ] **Step 1: Create build.bat**

```bat
@echo off
chcp 65001 > nul
echo Building iPhone Backup.exe...
echo.

pip install pyinstaller colorama pywin32 --quiet

pyinstaller ^
  --onefile ^
  --console ^
  --name "iPhone Backup" ^
  --hidden-import win32com.client ^
  --hidden-import win32com.shell ^
  --hidden-import pythoncom ^
  iphone_downloader.py

echo.
if exist "dist\iPhone Backup.exe" (
    echo Done! Your file is at:
    echo   dist\iPhone Backup.exe
) else (
    echo Build failed. Check the output above for errors.
)
echo.
pause
```

- [ ] **Step 2: Run the build**

```bash
cmd /c build.bat
```

Expected: `dist\iPhone Backup.exe` is created (~15–25 MB).

- [ ] **Step 3: Smoke test the .exe**

Double-click `dist\iPhone Backup.exe`. You should see the banner and folder prompt. Press any key or Ctrl+C to quit.

- [ ] **Step 4: Commit**

```bash
git add build.bat
git commit -m "feat: add PyInstaller build.bat for standalone exe packaging"
```

---

## Self-Review

**Spec coverage check:**

| Spec requirement | Covered by |
|---|---|
| Double-clickable .exe | Task 7 (build.bat → dist/iPhone Backup.exe) |
| No Python setup needed | Task 7 (--onefile PyInstaller) |
| Banner | Task 3 (print_banner) |
| Folder prompt with default | Task 3 (prompt_dest), Task 6 (main) |
| Colorama colors | Task 1 (bootstrap), Task 5 (transfer_photos) |
| Progress bar redraws in place | Task 5 (\r overwrite) |
| Dimmed completed files | Task 5 (DIM color) |
| iPhone not found + retry | Task 5 (retry loop with R key) |
| iPhone disconnected mid-transfer | Task 5 (find_iphone check after failure) |
| Disk space check | Task 5 (shutil.disk_usage) |
| No new photos message | Task 5 (total == 0 path) |
| Dest folder not writable | Task 5 (OSError handler) |
| Elapsed time on completion | Task 5 (format_elapsed) |
| "Press any key to close" | Task 3 (press_any_key), Task 5 (called at end) |
| Log file for debugging | Task 4 (FileHandler kept at DEBUG) |
| Resume support preserved | Task 5 (TransferState unchanged) |
| requirements.txt | Task 1 |

All spec requirements covered. ✓
