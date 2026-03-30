"""
copy_folder.py — copy a specific iPhone album folder to Windows.
Usage: python copy_folder.py 202602__
"""
import sys
import time
import shutil
import threading
import pythoncom
import win32com.client
from pathlib import Path

TIMEOUT_PER_FILE = 120
POLL_INTERVAL    = 1.0
STABLE_CHECKS    = 3
DEST_ROOT        = Path.home() / "Pictures" / "iPhone Photos"


def make_shell():
    return win32com.client.Dispatch("Shell.Application")


def folder_items(shell, item):
    try:
        f = shell.NameSpace(item.Path)
        if f: return list(f.Items())
    except Exception:
        pass
    try:
        f = item.GetFolder
        if f: return list(f.Items())
    except Exception:
        pass
    return []


def find_album(shell, album_name):
    computer = shell.NameSpace(17)
    for dev in computer.Items():
        if "iphone" not in dev.Name.lower() and "apple" not in dev.Name.lower():
            continue
        # find Internal Storage
        for child in folder_items(shell, dev):
            if "internal" in child.Name.lower() or "storage" in child.Name.lower():
                storage = child
                break
        else:
            storage = folder_items(shell, dev)[0] if folder_items(shell, dev) else None
        if storage is None:
            continue
        for album in folder_items(shell, storage):
            if album.Name == album_name:
                return album
    return None


def copy_file(shell, src_item, dest_dir: Path, filename: str) -> bool:
    tmp_dir  = dest_dir / "_tmp_transfer"
    tmp_dir.mkdir(exist_ok=True)
    tmp_file   = tmp_dir / filename
    final_file = dest_dir / filename

    if tmp_file.exists():
        try: tmp_file.unlink()
        except Exception: pass

    result = {"error": None}

    def do_copy():
        try:
            pythoncom.CoInitialize()
            ns = shell.NameSpace(str(tmp_dir))
            if ns is None:
                result["error"] = "no namespace"
                return
            ns.CopyHere(src_item, 4 | 16)
        except Exception as e:
            result["error"] = str(e)
        finally:
            pythoncom.CoUninitialize()

    t = threading.Thread(target=do_copy, daemon=True)
    t.start()

    deadline   = time.time() + TIMEOUT_PER_FILE
    last_size  = -1
    stable     = 0

    while time.time() < deadline:
        time.sleep(POLL_INTERVAL)
        if result["error"]:
            break
        if tmp_file.exists():
            try:
                size = tmp_file.stat().st_size
                if size > 0 and size == last_size:
                    stable += 1
                    if stable >= STABLE_CHECKS:
                        shutil.move(str(tmp_file), str(final_file))
                        return True
                else:
                    stable = 0
                    last_size = size
            except Exception:
                pass

    try:
        if tmp_file.exists(): tmp_file.unlink()
    except Exception:
        pass
    return False


def main():
    album_name = sys.argv[1] if len(sys.argv) > 1 else "202602__"
    dest_dir   = DEST_ROOT / album_name
    dest_dir.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()
    try:
        shell = make_shell()
        print(f"Looking for album: {album_name} ...")
        album = find_album(shell, album_name)
        if album is None:
            print("Album not found. Make sure iPhone is connected and unlocked.")
            return

        files = folder_items(shell, album)
        print(f"Found {len(files)} file(s) in {album_name}")

        copied = skipped = failed = 0
        for f in files:
            fname = f.Name
            dest  = dest_dir / fname
            if dest.exists() and dest.stat().st_size > 0:
                print(f"  skip  {fname}")
                skipped += 1
                continue
            print(f"  -> {fname}", end="", flush=True)
            ok = copy_file(shell, f, dest_dir, fname)
            if ok:
                size = dest.stat().st_size if dest.exists() else 0
                print(f"  OK  ({size:,} bytes)")
                copied += 1
            else:
                print(f"  FAILED")
                failed += 1

        print(f"\nDone. Copied={copied}  Skipped={skipped}  Failed={failed}")
        print(f"Saved to: {dest_dir}")
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
