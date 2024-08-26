

"""
Spotify Ad Muter

Author: Alofte
Contact: https://www.linkedin.com/in/alofte-py-090680304/
GitHub: https://github.com/Alofte/spotify-ad-muter
License: GNU GENERAL PUBLIC LICENSE

Description:
This script is designed to automatically mute advertisements(or whatever you want) on Spotify by detecting ad titles
and controlling the audio session.

Features:
- Automatically detects and mutes Spotify advertisements.
- Hotkeys to manage the mute list:
  - Ctrl + Shift + F1: Add current song to the mute list.
  - Ctrl + Shift + F3: Remove the last added song from the mute list.
  - Ctrl + Shift + F4: Reset the mute list.
"""


import win32gui
import win32process
import psutil
import time
import keyboard
import pycaw.pycaw as pycaw
from threading import Lock, Thread
import json
import os
import sys
import winreg as reg
import logging

# Determine the path for the mute list file
if getattr(sys, 'frozen', False):
    # The application is frozen (running as an .exe)
    application_path = os.path.dirname(sys.executable)
else:
    # Running as a script
    application_path = os.path.dirname(os.path.abspath(__file__))

MUTE_LIST_FILE = os.path.join(application_path, "mute_list.json")
mute_list_lock = Lock()

# Setting up logging
logging.basicConfig(filename=os.path.join(application_path, 'spotify_ad_muter.log'),
                    level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')


def add_to_startup(file_path=""):
    if file_path == "":
        file_path = os.path.abspath(sys.argv[0])
    key = reg.HKEY_CURRENT_USER
    key_value = r"Software\Microsoft\Windows\CurrentVersion\Run"
    open = reg.OpenKey(key, key_value, 0, reg.KEY_ALL_ACCESS)
    reg.SetValueEx(open, "Spotify Ad Muter", 0, reg.REG_SZ, file_path)
    reg.CloseKey(open)


def load_mute_list():
    if not os.path.exists(MUTE_LIST_FILE):
        with open(MUTE_LIST_FILE, "w") as f:
            json.dump([], f)  # Initialize with an empty list
    with open(MUTE_LIST_FILE, "r") as f:
        return json.load(f)


def save_mute_list(mute_list):
    with open(MUTE_LIST_FILE, "w") as f:
        json.dump(mute_list, f)


mute_list = load_mute_list()


def get_spotify_window():
    def callback(hwnd, hwnds):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            try:
                process = psutil.Process(pid)
                if process.name().lower() == "spotify.exe":
                    title = win32gui.GetWindowText(hwnd)
                    if title and title != "Spotify":
                        hwnds.append(hwnd)
            except psutil.NoSuchProcess:
                pass
        return True

    hwnds = []
    win32gui.EnumWindows(callback, hwnds)
    return hwnds[0] if hwnds else None


def mute_spotify():
    sessions = pycaw.AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session.SimpleAudioVolume
        if session.Process and session.Process.name() == "Spotify.exe":
            volume.SetMute(1, None)
            logging.info("Spotify muted")


def unmute_spotify():
    sessions = pycaw.AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session.SimpleAudioVolume
        if session.Process and session.Process.name() == "Spotify.exe":
            volume.SetMute(0, None)
            logging.info("Spotify unmuted")


def add_to_mute_list():
    logging.debug("Hotkey F1 pressed")
    hwnd = get_spotify_window()
    if hwnd:
        title = win32gui.GetWindowText(hwnd)
        with mute_list_lock:
            if title and title not in mute_list:
                mute_list.append(title)
                save_mute_list(mute_list)
                logging.info(f"Added '{title}' to mute list")


def remove_last_from_mute_list():
    logging.debug("Hotkey F3 pressed")
    with mute_list_lock:
        if mute_list:
            removed_item = mute_list.pop()
            save_mute_list(mute_list)
            logging.info(f"Removed '{removed_item}' from mute list")
        else:
            logging.info("Mute list is already empty")


def reset_mute_list():
    global mute_list
    logging.debug("Hotkey F4 pressed")
    with mute_list_lock:
        mute_list = []
        save_mute_list(mute_list)
        logging.info("Mute list reset")


def hotkey_listener():
    keyboard.add_hotkey('ctrl+shift+f1', add_to_mute_list)
    keyboard.add_hotkey('ctrl+shift+f3', remove_last_from_mute_list)
    keyboard.add_hotkey('ctrl+shift+f4', reset_mute_list)
    keyboard.wait()  # Keep the thread alive to listen for hotkeys


def main():
    # Start the hotkey listener in a separate thread
    hotkey_thread = Thread(target=hotkey_listener, daemon=True)
    hotkey_thread.start()

    # Set the script's priority to high
    p = psutil.Process(os.getpid())
    p.nice(psutil.HIGH_PRIORITY_CLASS)

    last_title = None
    while True:
        hwnd = get_spotify_window()
        if hwnd:
            title = win32gui.GetWindowText(hwnd)
            with mute_list_lock:
                if title in mute_list or title.lower() == "advertisement":
                    if title != last_title:
                        mute_spotify()
                        last_title = title
                else:
                    if last_title is not None:
                        unmute_spotify()
                        last_title = None  # Reset last_title to ensure unmute works properly
        time.sleep(1)


if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        # The application is frozen (running as an .exe)
        application_path = os.path.dirname(sys.executable)
    else:
        # Running as a script
        application_path = os.path.dirname(os.path.abspath(__file__))

    # Add to startup
    add_to_startup(os.path.join(application_path, "SpotifyAdMuter.exe"))

    main()
