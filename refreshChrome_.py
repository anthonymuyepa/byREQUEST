import os
import time
import socket
import subprocess
import sys
import ctypes
from datetime import datetime


# ====== CONFIG ======
TARGET_URL = "https://www.doifcuhb.org/connect/web/guest/home?p_p_id=58&p_p_lifecycle=0&p_p_state=maximized&p_p_mode=view&saveLastPath=0&_58_struts_action=%2Flogin%2Flogin#/groups?dynamicGroupsView=ACTIVE&manageGroupsView=STATIC&staticGroupsView=GROUPS"

LOG_PATH = r"C:\Users\hillary\Documents\byREQUEST\EventLog.log"

chrome_binary_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

chrome_user_data_dir = r"C:\ChromeDebugProfile"
profile_directory = "Default"


# ====== LOGGING ======
def log_event(category, message):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{ts}\t{category}\t{message}"

    try:
        os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception as e:
        print(f"[warn] could not write log: {e}")

    print(line)


# ====== WINDOWS HELPERS ======
user32 = ctypes.windll.user32

VK_RETURN = 0x0D
VK_TAB = 0x09
KEYEVENTF_KEYUP = 0x0002
SW_RESTORE = 9


def kill_all_chrome_windows():
    """
    Kill all Chrome processes but keep saved profile data.
    """
    if sys.platform == "win32":
        subprocess.run(
            ["taskkill", "/F", "/IM", "chrome.exe"],
            capture_output=True,
            text=True
        )
        log_event("11", "(Chrome) Killed existing Chrome processes")
        time.sleep(3)
    else:
        log_event("11", "(Chrome) Kill Chrome skipped - Windows only")


def launch_chrome():
    """
    Launch Chrome directly to the DOI login page.
    """
    os.makedirs(chrome_user_data_dir, exist_ok=True)

    cmd = [
        chrome_binary_path,
        f"--user-data-dir={chrome_user_data_dir}",
        f"--profile-directory={profile_directory}",
        "--no-first-run",
        "--disable-extensions",
        "--disable-session-crashed-bubble",
        "--hide-crash-restore-bubble",
        TARGET_URL
    ]

    log_event("11", f"(Chrome) Launching Chrome to login page")
    log_event("11", f"(Chrome) Using profile: {chrome_user_data_dir}")

    subprocess.Popen(
        cmd,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


def bring_chrome_to_front():
    """
    Find a Chrome window and bring it to the front.
    """

    found_hwnd = ctypes.c_void_p(0)

    EnumWindowsProc = ctypes.WINFUNCTYPE(
        ctypes.c_bool,
        ctypes.c_void_p,
        ctypes.c_void_p
    )

    def callback(hwnd, lparam):
        if not user32.IsWindowVisible(hwnd):
            return True

        length = user32.GetWindowTextLengthW(hwnd)
        if length == 0:
            return True

        buff = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buff, length + 1)

        title = buff.value.lower()

        if "chrome" in title or "sign in" in title or "welcome" in title or "doifcuhb" in title:
            found_hwnd.value = hwnd
            return False

        return True

    user32.EnumWindows(EnumWindowsProc(callback), 0)

    if found_hwnd.value:
        log_event("11", "(Chrome) Bringing Chrome window to front")

        user32.ShowWindow(found_hwnd.value, SW_RESTORE)
        time.sleep(0.5)

        try:
            user32.SetForegroundWindow(found_hwnd.value)
        except Exception:
            pass

        time.sleep(1)
        return True

    log_event("5", "(Chrome) Could not find Chrome window")
    return False


def press_key(vk_code):
    user32.keybd_event(vk_code, 0, 0, 0)
    time.sleep(0.1)
    user32.keybd_event(vk_code, 0, KEYEVENTF_KEYUP, 0)


def press_enter():
    log_event("11", "(Login) Pressing Enter to submit login form")
    press_key(VK_RETURN)


def press_tab_then_enter():
    """
    Backup option:
    If Enter alone does not submit, Tab may move focus to the Sign In button,
    then Enter clicks it.
    """
    log_event("11", "(Login) Pressing Tab then Enter")
    press_key(VK_TAB)
    time.sleep(0.3)
    press_key(VK_RETURN)


# ====== MAIN ======
def main():
    try:
        log_event("11", "(Main) Starting Chrome login helper")
        log_event("11", "(Main) No Selenium / no ChromeDriver attach will be used")

        kill_all_chrome_windows()
        launch_chrome()

        log_event("11", "(Main) Waiting for login page to load and saved password to fill...")
        time.sleep(10)

        bring_chrome_to_front()

        # First attempt: Enter usually submits the login form if password is filled
        press_enter()

        log_event("11", "(Main) Waiting after Enter...")
        time.sleep(8)

        # Optional backup: if Enter did not work, try Tab + Enter
        # You can comment this out if Enter alone works.
        press_tab_then_enter()

        log_event("11", "(Main) Login click/submit attempt completed")
        log_event("11", "(Main) Chrome remains open")

        # Keep script alive so ByRequest can see status
        while True:
            time.sleep(30)

    except KeyboardInterrupt:
        log_event("11", "(Main) Script stopped by user")

    except Exception as e:
        log_event("5", f"(Main) Error: {e}")

        import traceback
        traceback.print_exc()

    finally:
        log_event("11", "(Main) Done")


if __name__ == "__main__":
    main()