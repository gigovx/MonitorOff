import os
import sys
import threading
import ctypes
from pathlib import Path
from PIL import Image
import pystray

# Win32 message constants
WM_SYSCOMMAND   = 0x0112
SC_MONITORPOWER = 0xF170
HWND_BROADCAST  = 0xFFFF

APP_NAME      = 'MonitorOff'
ICON_FILENAME = 'icon.ico'  # this is bundled via PyInstaller

def turn_off_monitor():
    ctypes.windll.user32.PostMessageW(
        HWND_BROADCAST,
        WM_SYSCOMMAND,
        SC_MONITORPOWER,
        2
    )

def add_to_startup():
    """
    Drops a .url shortcut pointing at this exe into the user's Startup folder.
    """
    startup_dir = Path(os.environ["APPDATA"]) / \
                  "Microsoft" / "Windows" / "Start Menu" / \
                  "Programs" / "Startup"
    exe_path = Path(sys.executable if getattr(sys, 'frozen', False)
                    else sys.argv[0]).resolve()

    shortcut = startup_dir / f"{APP_NAME}.url"
    with open(shortcut, 'w') as f:
        f.write("[InternetShortcut]\n")
        f.write(f"URL=file:///{exe_path.as_posix()}\n")
        f.write(f"IconFile={exe_path.as_posix()}\n")
        f.write("IconIndex=0\n")

def on_menu_click(icon, item):
    if item.text == 'Turn Off':
        threading.Thread(target=turn_off_monitor, daemon=True).start()
    elif item.text == 'Add to Startup':
        add_to_startup()
    elif item.text == 'Exit':
        icon.stop()
        sys.exit()

def load_icon():
    """
    Locate the bundled icon:
      - when frozen, files are unpacked to sys._MEIPASS
      - otherwise look next to the script
    """
    if getattr(sys, 'frozen', False):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).parent
    return Image.open(base / ICON_FILENAME)

def main():
    tray_icon = load_icon()
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