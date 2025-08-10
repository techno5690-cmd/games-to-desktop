Python 3.10.6 (tags/v3.10.6:9c7b4bd, Aug  1 2022, 21:53:49) [MSC v.1932 64 bit (AMD64)] on win32
Type "help", "copyright", "credits" or "license()" for more information.
import os
import re
import time
import winshell
from win32com.client import Dispatch
from tkinter import Tk, messagebox

# ==== CONFIG ====
PIRATED_DIR = r"D:\pirated games"
STEAM_DIR = r"D:\SteamLibrary\steamapps\common"
DESKTOP = os.path.join(os.environ["USERPROFILE"], "Desktop")
LOG_FILE = os.path.join(DESKTOP, "shortcut_log.txt")

# Words to remove from names
CLEAN_PATTERNS = [
    r"steam\.rip", r"\.com", r"fitgirl", r"installer", r"setup", r"install",
    r"\(.*?\)", r"\[.*?\]", r"v\d+(\.\d+)*"
]

SKIP_EXE = [
    "UnityCrashHandler", "CrashSender", "vcredist", "dxsetup", "setup", "uninstall", "helper"
]


# ==== FUNCTIONS ====
def clean_name(name: str) -> str:
    """Cleans folder or shortcut names for comparison."""
    name = name.lower()
    name = name.replace("-", " ").replace("_", " ")
    for pattern in CLEAN_PATTERNS:
        name = re.sub(pattern, "", name, flags=re.IGNORECASE)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def get_existing_desktop_games():
    """Reads all existing .lnk files on desktop and returns a set of cleaned names."""
    existing = set()
    for file in os.listdir(DESKTOP):
        if file.lower().endswith(".lnk"):
            existing.add(clean_name(os.path.splitext(file)[0]))
    return existing


def find_main_exe(folder):
    """Finds the largest non-skipped exe file in folder and subfolders."""
    largest_file = None
    largest_size = 0
    for root, _, files in os.walk(folder):
        for f in files:
            if f.lower().endswith(".exe") and not any(skip.lower() in f.lower() for skip in SKIP_EXE):
                size = os.path.getsize(os.path.join(root, f))
                if size > largest_size:
                    largest_file = os.path.join(root, f)
                    largest_size = size
    return largest_file


def create_shortcut(name, target):
    """Creates a shortcut on the desktop."""
    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(os.path.join(DESKTOP, f"{name}.lnk"))
    shortcut.TargetPath = target
    shortcut.WorkingDirectory = os.path.dirname(target)
    shortcut.Save()


def log_action(action, game):
    """Logs action to file."""
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"{action}: {game}\n")


def popup(msg):
    """Shows a temporary popup message."""
    root = Tk()
    root.withdraw()
    messagebox.showinfo("Game Shortcut Creator", msg)
    root.destroy()


# ==== MAIN ====
if __name__ == "__main__":
    # Popup countdown
    root = Tk()
    root.withdraw()
    for i in range(3, 0, -1):
        messagebox.showinfo("Shortcut Creator", f"Starting in {i} seconds...")
        time.sleep(1)
    root.destroy()

    existing_games = get_existing_desktop_games()

    all_game_dirs = []
    for base_dir in [PIRATED_DIR, STEAM_DIR]:
        if os.path.exists(base_dir):
            all_game_dirs.extend(
                [os.path.join(base_dir, d) for d in os.listdir(base_dir)
                 if os.path.isdir(os.path.join(base_dir, d))]
            )

    for folder in all_game_dirs:
        game_name = os.path.basename(folder)
        cleaned_game_name = clean_name(game_name)

        if cleaned_game_name in existing_games:
            msg = f"Skipping: {game_name} (Already exists)"
            print(msg)
            popup(msg)
            log_action("SKIPPED", game_name)
            continue

        main_exe = find_main_exe(folder)
        if main_exe:
            create_shortcut(game_name, main_exe)
            msg = f"Added shortcut: {game_name}"
            print(msg)
            popup(msg)
            existing_games.add(cleaned_game_name)
            log_action("ADDED", game_name)
        else:
            msg = f"No executable found for: {game_name}"
            print(msg)
            popup(msg)
            log_action("NO EXE", game_name)

    popup("All done! See shortcut_log.txt for details.")
    print("Done. See shortcut_log.txt for details.")
