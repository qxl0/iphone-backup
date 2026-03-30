#!/usr/bin/env python3
"""
iPhone Photo Downloader
-----------------------
Automatically transfers photos from iPhone to Windows with resume support.
Handles freezes and disconnections by timing out per-file and retrying.

Requirements:
    pip install pywin32

Usage:
    python iphone_downloader.py                  # saves to Pictures/iPhone Photos
    python iphone_downloader.py D:/MyPhotos      # saves to a custom folder
    python iphone_downloader.py --reset          # clear saved state and start fresh
"""

import os
import sys
import json
import time
import logging
import threading
import shutil
import pythoncom
import win32com.client
from pathlib import Path
from datetime import datetime
from typing import Optional

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
colorama.init(autoreset=True)

# ── Color shortcuts ────────────────────────────────────────────────────────────
CYAN   = Fore.CYAN
GREEN  = Fore.GREEN
YELLOW = Fore.YELLOW
RED    = Fore.RED
DIM    = Style.DIM
BRIGHT = Style.BRIGHT
RESET  = Style.RESET_ALL

# ── UI helper functions ────────────────────────────────────────────────────────────

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
    user_input = input("> ").strip()
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

# ── Configuration ──────────────────────────────────────────────────────────────
TIMEOUT_PER_FILE  = 120   # seconds to wait for a single file before giving up
MAX_RETRIES       = 3     # how many times to retry a failed file
POLL_INTERVAL     = 1.0   # seconds between file-size checks
STABLE_CHECKS     = 3     # consecutive equal-size readings = file is done
# ───────────────────────────────────────────────────────────────────────────────


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


# ── State tracking (resume support) ───────────────────────────────────────────

class TransferState:
    def __init__(self, state_file: Path):
        self.path = state_file
        self.data = self._load()

    def _load(self):
        if self.path.exists():
            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {"done": {}, "failed": {}}

    def save(self):
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, indent=2)

    def is_done(self, key: str) -> bool:
        return key in self.data["done"]

    def mark_done(self, key: str, size: int):
        self.data["done"][key] = {"size": size, "time": datetime.now().isoformat()}
        self.data["failed"].pop(key, None)
        self.save()

    def get_retries(self, key: str) -> int:
        return self.data["failed"].get(key, {}).get("retries", 0)

    def increment_retries(self, key: str):
        entry = self.data["failed"].setdefault(key, {"retries": 0})
        entry["retries"] += 1
        self.save()

    def reset(self):
        self.data = {"done": {}, "failed": {}}
        self.save()


# ── Shell / MTP helpers ────────────────────────────────────────────────────────

def _make_shell():
    """Create a Shell.Application COM object (must be called after CoInitialize)."""
    return win32com.client.Dispatch("Shell.Application")


def _folder_items(shell, item):
    """Return a list of child items inside a Shell folder item."""
    try:
        folder = shell.NameSpace(item.Path)
        if folder is not None:
            return list(folder.Items())
    except Exception:
        pass
    try:
        folder = item.GetFolder
        if folder is not None:
            return list(folder.Items())
    except Exception:
        pass
    return []


def _refetch_album(album_name: str, log):
    """Re-navigate from scratch to a specific album, returning its file items.
    Used as a fallback when the cached shell returns 0 items for an album."""
    RETRIES = 3
    for attempt in range(1, RETRIES + 1):
        try:
            pythoncom.CoInitialize()
            shell = win32com.client.Dispatch("Shell.Application")
            computer = shell.NameSpace(17)
            for dev in computer.Items():
                n = dev.Name.lower()
                if "iphone" not in n and "apple" not in n:
                    continue
                for child in _folder_items(shell, dev):
                    if "internal" in child.Name.lower() or "storage" in child.Name.lower():
                        storage = child
                        break
                else:
                    children = _folder_items(shell, dev)
                    storage = children[0] if children else None
                if storage is None:
                    continue
                for album in _folder_items(shell, storage):
                    if album.Name == album_name:
                        items = _folder_items(shell, album)
                        if items:
                            log.info(f"  [retry {attempt}] found {len(items)} file(s) via fresh shell")
                            return items
            time.sleep(2)
        except Exception as exc:
            log.debug(f"  _refetch_album attempt {attempt} error: {exc}")
            time.sleep(2)
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
    return []


def find_iphone(shell, log):
    """Find the iPhone device under 'This PC' (CSIDL_DRIVES = 17)."""
    computer = shell.NameSpace(17)
    if computer is None:
        return None
    for item in computer.Items():
        name = item.Name.lower()
        if "iphone" in name or ("apple" in name and "icloud" not in name):
            log.info(f"Found device: {item.Name}")
            return item
    return None


def find_photos_root(shell, iphone_item, log):
    """Navigate to the photo root folder on the iPhone.

    Handles two layouts:
      Layout A: iPhone → Internal Storage → DCIM → 100APPLE …
      Layout B: iPhone → Internal Storage → 201707__ … (no DCIM level)
    Returns the folder whose direct children are the per-album subfolders.
    """
    root_items = _folder_items(shell, iphone_item)

    storage_item = None
    for item in root_items:
        n = item.Name.lower()
        if "internal" in n or "storage" in n:
            storage_item = item
            break
        if item.Name.upper() == "DCIM":
            log.info(f"Found DCIM at root: {item.Path}")
            return item

    if storage_item is None and root_items:
        storage_item = root_items[0]  # take first child as fallback

    if storage_item is None:
        return None

    storage_children = _folder_items(shell, storage_item)

    # Layout A: DCIM subfolder exists
    for item in storage_children:
        if item.Name.upper() == "DCIM":
            log.info(f"Found DCIM under storage: {item.Name}")
            return item

    # Layout B: no DCIM — album folders are directly inside Internal Storage
    # Confirm by checking whether any child looks like a date/album folder
    if storage_children:
        log.info(
            f"No DCIM folder found; using Internal Storage directly "
            f"({len(storage_children)} album folder(s) visible)"
        )
        return storage_item

    return None


# ── File copy with timeout ─────────────────────────────────────────────────────

def copy_with_timeout(shell, src_item, dest_dir: Path, filename: str,
                      timeout: int, log) -> bool:
    """
    Copy src_item into dest_dir using Shell CopyHere.
    Polls until the file stabilises, then moves it to its final location.
    Returns True on success.
    """
    tmp_dir  = dest_dir / "_tmp_transfer"
    tmp_dir.mkdir(exist_ok=True)
    tmp_file  = tmp_dir / filename
    final_file = dest_dir / filename

    # Remove stale temp file from a previous interrupted attempt
    if tmp_file.exists():
        try:
            tmp_file.unlink()
        except Exception:
            pass

    result = {"error": None}

    def do_copy():
        try:
            pythoncom.CoInitialize()
            dest_ns = shell.NameSpace(str(tmp_dir))
            if dest_ns is None:
                result["error"] = "Cannot get destination namespace"
                return
            # Flags: 4 = no UI dialog, 16 = no overwrite confirmation
            dest_ns.CopyHere(src_item, 4 | 16)
        except Exception as exc:
            result["error"] = str(exc)
        finally:
            pythoncom.CoUninitialize()

    t = threading.Thread(target=do_copy, daemon=True)
    t.start()

    deadline     = time.time() + timeout
    last_size    = -1
    stable_count = 0

    while time.time() < deadline:
        time.sleep(POLL_INTERVAL)

        if result["error"]:
            log.debug(f"    copy error: {result['error']}")
            break

        if tmp_file.exists():
            try:
                size = tmp_file.stat().st_size
                if size > 0 and size == last_size:
                    stable_count += 1
                    if stable_count >= STABLE_CHECKS:
                        shutil.move(str(tmp_file), str(final_file))
                        return True
                else:
                    stable_count = 0
                    last_size = size
            except Exception:
                pass
    else:
        log.debug(f"    timeout waiting for {filename}")

    # Cleanup failed temp file
    try:
        if tmp_file.exists():
            tmp_file.unlink()
    except Exception:
        pass
    return False


# ── Main transfer logic ────────────────────────────────────────────────────────

def transfer_photos(dest_dir: Path, log, reset: bool = False) -> None:
    dest_dir.mkdir(parents=True, exist_ok=True)
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
        copied = skipped = failed = perm_failed = 0
        start_time = time.time()
        last_sub = None  # type: Optional[str]

        for sub_name, dest_sub, file_item, filename in all_items:
            dest_sub.mkdir(exist_ok=True)
            key       = f"{sub_name}/{filename}"
            dest_path = dest_sub / filename
            done_so_far = copied + skipped + failed + perm_failed

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
                perm_failed += 1
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

        if copied == 0 and failed == 0 and skipped + perm_failed == total:
            print(f"\n{GREEN}\u2713 No new photos found \u2014 you\u2019re all backed up!{RESET}")
        else:
            print(f"\n{GREEN}{BRIGHT}\u2713 Backup complete!{RESET}")
            if copied:
                print(f"  {copied:,} photos copied")
            if skipped:
                print(f"  {skipped:,} already up to date (skipped)")
            if failed:
                print(f"  {RED}{failed:,} failed{RESET} \u2014 run again to retry")
            if perm_failed:
                print(f"  {RED}{perm_failed:,} could not be copied{RESET} (tried {MAX_RETRIES}x, giving up)")
            print(f"  Completed in {format_elapsed(elapsed_total)}")

        print(f"\n{DIM}Saved to: {dest_dir}{RESET}")
        if failed or perm_failed:
            log.warning(f"{failed} file(s) failed this run; {perm_failed} permanently skipped after {MAX_RETRIES} retries.")

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


# ── Entry point ────────────────────────────────────────────────────────────────

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


if __name__ == "__main__":
    main()
