import os
import shutil
import sys
import threading
import ctypes
from pathlib import Path
from PIL import Image
import pystray

# Constants
WM_SYSCOMMAND    = 0x0112
SC_MONITORPOWER  = 0xF170
HWND_BROADCAST   = 0xFFFF
APP_NAME = 'MonitorOff'
ICON_PATH = 'icon.ico'

def turn_off_monitor():
    ctypes.windll.user32.PostMessageW(HWND_BROADCAST, WM_SYSCOMMAND, SC_MONITORPOWER, 2)

def add_to_startup():
    """Creates a shortcut in the Startup folder."""
    startup_dir = Path(os.environ["APPDATA"]) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
    script_path = Path(sys.argv[0]).resolve()

    if script_path.suffix.lower() == '.exe':
        target = script_path
    else:
        # If running .py, point to python executable + script
        target = Path(sys.executable)
    
    shortcut = startup_dir / f"{APP_NAME}.url"
    with open(shortcut, 'w') as f:
        f.write(f"[InternetShortcut]\n")
        f.write(f"URL=file:///{script_path.as_posix()}\n")
        f.write(f"IconFile={ICON_PATH}\n")
        f.write("IconIndex=0\n")

def on_menu_click(icon, item):
    text = item.text
    if text == 'Turn Off':
        threading.Thread(target=turn_off_monitor, daemon=True).start()
    elif text == 'Add to Startup':
        add_to_startup()
    elif text == 'Exit':
        icon.stop()
        sys.exit()

def main():
    tray_icon = Image.open(ICON_PATH)

    icon = pystray.Icon(
        name='monitor_off',
        icon=tray_icon,
        title=APP_NAME,
        menu=pystray.Menu(
            pystray.MenuItem('Turn Off', on_menu_click),
            pystray.MenuItem('Add to Startup', on_menu_click),
            pystray.MenuItem('Exit', on_menu_click)
        )
    )
    icon.run()

if __name__ == '__main__':
    main()