import os
import re
import sys
import time
import platform
import subprocess
from typing import Optional

# Optional Windows-specific COM import
try:
    from win32com.client import Dispatch  # type: ignore
    HAS_WIN32COM = True
except Exception:
    Dispatch = None  # type: ignore
    HAS_WIN32COM = False

# Optional Tkinter popups
try:
    from tkinter import Tk, messagebox  # type: ignore
    HAS_TK = True
except Exception:
    Tk = None  # type: ignore
    messagebox = None  # type: ignore
    HAS_TK = False


# ==== CONFIG ====
IS_WINDOWS = platform.system() == "Windows"
HOME_DIR = os.path.expanduser("~")

# Allow overriding via environment; provide platform-appropriate defaults
DEFAULT_PIRATED_DIR = r"D:\pirated games" if IS_WINDOWS else os.path.join(HOME_DIR, "Games")
DEFAULT_STEAM_DIR = (
    r"D:\SteamLibrary\steamapps\common"
    if IS_WINDOWS
    else os.path.join(HOME_DIR, ".local", "share", "Steam", "steamapps", "common")
)

PIRATED_DIR = os.environ.get("PIRATED_DIR", DEFAULT_PIRATED_DIR)
STEAM_DIR = os.environ.get("STEAM_DIR", DEFAULT_STEAM_DIR)


def get_desktop_path() -> str:
    if IS_WINDOWS:
        # Try winshell if available; otherwise fallback to USERPROFILE/Desktop
        desktop = os.path.join(os.environ.get("USERPROFILE", HOME_DIR), "Desktop")
        return desktop
    # Non-Windows: try xdg-user-dir DESKTOP
    try:
        completed = subprocess.run(
            ["xdg-user-dir", "DESKTOP"],
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            check=False,
            text=True,
        )
        candidate = completed.stdout.strip()
        if candidate and os.path.isabs(candidate):
            return candidate
    except Exception:
        pass
    # Fallback
    return os.path.join(HOME_DIR, "Desktop")


DESKTOP = get_desktop_path()
LOG_FILE = os.path.join(DESKTOP if os.path.isdir(DESKTOP) else HOME_DIR, "shortcut_log.txt")

# Words to remove from names
CLEAN_PATTERNS = [
    r"steam\.rip",
    r"\.com",
    r"fitgirl",
    r"installer",
    r"setup",
    r"install",
    r"\\(.*?\\)",
    r"\\[.*?\\]",
    r"v\\d+(\\.\\d+)*",
]

SKIP_EXE = [
    "UnityCrashHandler",
    "CrashSender",
    "vcredist",
    "dxsetup",
    "setup",
    "uninstall",
    "helper",
    "steamerrorreporter",
]


# ==== UTILITIES ====

def clean_name(name: str) -> str:
    name = name.lower()
    name = name.replace("-", " ").replace("_", " ")
    for pattern in CLEAN_PATTERNS:
        name = re.sub(pattern, "", name, flags=re.IGNORECASE)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def get_existing_desktop_games() -> set[str]:
    """Reads existing shortcuts on desktop (.lnk on Windows, .desktop elsewhere)."""
    existing: set[str] = set()
    if not os.path.isdir(DESKTOP):
        return existing
    for file_name in os.listdir(DESKTOP):
        lower = file_name.lower()
        if lower.endswith(".lnk") or lower.endswith(".desktop"):
            existing.add(clean_name(os.path.splitext(file_name)[0]))
    return existing


def is_executable_file(path: str) -> bool:
    if not os.path.isfile(path):
        return False
    if IS_WINDOWS:
        return path.lower().endswith(".exe")
    return os.access(path, os.X_OK)


def should_skip_exe(file_name: str) -> bool:
    lower = file_name.lower()
    return any(skip.lower() in lower for skip in SKIP_EXE)


def find_main_executable(folder: str) -> Optional[str]:
    """Find the largest plausible executable in folder and descendants."""
    largest_file: Optional[str] = None
    largest_size = 0
    for root, _, files in os.walk(folder):
        for file_name in files:
            if not is_executable_file(os.path.join(root, file_name)):
                continue
            if should_skip_exe(file_name):
                continue
            full_path = os.path.join(root, file_name)
            try:
                size = os.path.getsize(full_path)
            except OSError:
                continue
            if size > largest_size:
                largest_file = full_path
                largest_size = size
    return largest_file


def create_shortcut(name: str, target: str) -> None:
    """Create a desktop shortcut: .lnk on Windows, .desktop otherwise."""
    if not os.path.isdir(DESKTOP):
        os.makedirs(DESKTOP, exist_ok=True)
    if IS_WINDOWS:
        if not HAS_WIN32COM or Dispatch is None:
            raise RuntimeError(
                "Windows shortcut creation requires pywin32 (win32com). Please install it."
            )
        shell = Dispatch("WScript.Shell")
        shortcut_path = os.path.join(DESKTOP, f"{name}.lnk")
        shortcut = shell.CreateShortcut(shortcut_path)
        shortcut.TargetPath = target
        shortcut.WorkingDirectory = os.path.dirname(target)
        shortcut.Save()
        return
    # Non-Windows: create .desktop launcher
    shortcut_path = os.path.join(DESKTOP, f"{name}.desktop")
    working_dir = os.path.dirname(target)
    contents = (
        "[Desktop Entry]\n"
        "Type=Application\n"
        f"Name={name}\n"
        f"Exec=\"{target}\"\n"
        f"Path={working_dir}\n"
        "Terminal=false\n"
    )
    with open(shortcut_path, "w", encoding="utf-8") as f:
        f.write(contents)
    # Make it executable
    try:
        os.chmod(shortcut_path, 0o755)
    except Exception:
        pass


def log_action(action: str, game: str) -> None:
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{action}: {game}\n")
    except Exception:
        # As a last resort, print if logging fails
        print(f"{action}: {game}")


def popup(message: str) -> None:
    """Shows a popup if Tk is available and a display is likely; otherwise prints."""
    if HAS_TK and Tk is not None and messagebox is not None:
        try:
            root = Tk()
            root.withdraw()
            messagebox.showinfo("Game Shortcut Creator", message)
            root.destroy()
            return
        except Exception:
            # Fallback to print
            pass
    print(message)


# ==== MAIN ====
if __name__ == "__main__":
    # Optional countdown
    for i in range(3, 0, -1):
        popup(f"Starting in {i} seconds...")
        time.sleep(1)

    existing_games = get_existing_desktop_games()

    all_game_dirs: list[str] = []
    for base_dir in [PIRATED_DIR, STEAM_DIR]:
        if not base_dir:
            continue
        if os.path.exists(base_dir):
            try:
                for entry in os.listdir(base_dir):
                    full = os.path.join(base_dir, entry)
                    if os.path.isdir(full):
                        all_game_dirs.append(full)
            except Exception:
                pass

    for folder in all_game_dirs:
        game_name = os.path.basename(folder)
        cleaned_game_name = clean_name(game_name)

        if cleaned_game_name in existing_games:
            msg = f"Skipping: {game_name} (Already exists)"
            print(msg)
            popup(msg)
            log_action("SKIPPED", game_name)
            continue

        main_exec = find_main_executable(folder)
        if main_exec:
            try:
                create_shortcut(game_name, main_exec)
                msg = f"Added shortcut: {game_name}"
                print(msg)
                popup(msg)
                existing_games.add(cleaned_game_name)
                log_action("ADDED", game_name)
            except Exception as exc:
                msg = f"Failed to create shortcut for {game_name}: {exc}"
                print(msg)
                popup(msg)
                log_action("FAILED", game_name)
        else:
            msg = f"No executable found for: {game_name}"
            print(msg)
            popup(msg)
            log_action("NO EXE", game_name)

    popup("All done! See shortcut_log.txt for details.")
    print("Done. See shortcut_log.txt for details.")
