"""
diagnose.py — shows exactly what Windows sees on your connected iPhone.
Run this with the iPhone connected and unlocked, then share the output.
"""
import pythoncom
import win32com.client

def dump_items(shell, item, indent=0, max_depth=4):
    if indent > max_depth:
        return
    prefix = "  " * indent
    try:
        folder = shell.NameSpace(item.Path)
        if folder is None:
            try:
                folder = item.GetFolder
            except Exception:
                pass
        if folder is None:
            print(f"{prefix}  [cannot open]")
            return
        children = list(folder.Items())
        for child in children:
            try:
                size = child.Size
            except Exception:
                size = ""
            print(f"{prefix}  [{child.Name}]  path={child.Path}  size={size}")
            if indent < max_depth - 1:
                dump_items(shell, child, indent + 1, max_depth)
    except Exception as e:
        print(f"{prefix}  [error: {e}]")

def main():
    pythoncom.CoInitialize()
    shell = win32com.client.Dispatch("Shell.Application")
    computer = shell.NameSpace(17)  # This PC

    print("=== Scanning 'This PC' for Apple/iPhone devices ===\n")
    found = False
    for item in computer.Items():
        name = item.Name.lower()
        if "iphone" in name or "apple" in name or "ipad" in name:
            found = True
            print(f"DEVICE: {item.Name}")
            print(f"  Path : {item.Path}")
            dump_items(shell, item, indent=0, max_depth=4)
            print()

    if not found:
        print("No iPhone/Apple device found under 'This PC'.")
        print("\nAll devices found:")
        for item in computer.Items():
            print(f"  {item.Name}  —  {item.Path}")

    pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
