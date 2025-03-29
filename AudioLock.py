import tkinter as tk
import threading
import time
import subprocess
import os
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
from pynput import keyboard
import ctypes
import sys

def run_as_admin():
    """Restart the script with elevated privileges"""
    try:
        if ctypes.windll.shell32.IsUserAnAdmin():
            print("DEBUG: Script is running as administrator.")
            return  # Already running as admin

        # Relaunch the script with admin rights
        params = " ".join([f'"{arg}"' for arg in sys.argv])
        response = ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)

        if response <= 32:
            print("Failed to elevate to admin. Exiting...")
            sys.exit(1)

        sys.exit(0)
    except Exception as e:
        print(f"Error while trying to elevate privileges: {e}")
        sys.exit(1)

# Ensure the script is running as administrator
run_as_admin()

# Get audio device and volume control
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = interface.QueryInterface(IAudioEndpointVolume)

# Function to set volume level
def set_volume(value):
    volume_level = int(value) / 100
    volume.SetMasterVolumeLevelScalar(volume_level, None)

# Function to monitor and lock volume
def lock_volume():
    global lock_enabled
    while lock_enabled:
        current_level = volume.GetMasterVolumeLevelScalar()
        target_level = int(volume_slider.get()) / 100
        if round(current_level, 2) != round(target_level, 2):
            volume.SetMasterVolumeLevelScalar(target_level, None)
        time.sleep(0.5)  # Adjusting frequency of checks

# Volume Lock Toggle
def toggle_lock():
    global lock_enabled, lock_thread
    if lock_enabled:
        lock_enabled = False
        lock_button.config(text="Lock Volume")
    else:
        lock_enabled = True
        lock_thread = threading.Thread(target=lock_volume, daemon=True)
        lock_thread.start()
        lock_button.config(text="Unlock Volume")

# Function to intercept volume key presses
def on_press(key):
    if not lock_enabled:  # Allow volume adjustment if lock is disabled
        if key == keyboard.Key.media_volume_up:
            volume_slider.set(volume_slider.get() + 2)
            set_volume(volume_slider.get())
            return False  # Prevent default system handling
        elif key == keyboard.Key.media_volume_down:
            volume_slider.set(volume_slider.get() - 2)
            set_volume(volume_slider.get())
            return False

# Keyboard listener (runs in a thread)
def listen_for_keys():
    with keyboard.Listener(on_press=on_press) as listener:
        listener.join()

# Functions to disable/enable HID service using PowerShell
def disable_hid_service():
    print("Disabling HID Service via PowerShell...")
    try:
        result = subprocess.run(
            ["powershell", "-Command", "Set-Service -Name 'HidServ' -StartupType Disabled; Stop-Service -Name 'HidServ'"],
            check=True,
            shell=True,
            capture_output=True,
            text=True
        )
        print("HID Service disabled successfully.")
        print(f"PowerShell Output: {result.stdout}")
    except subprocess.CalledProcessError as e:
        print(f"Failed to disable HID Service. Error: {e}")
        print(f"PowerShell Error Output: {e.stderr}")

def enable_hid_service():
    print("Enabling HID Service via PowerShell...")
    try:
        result = subprocess.run(
            ["powershell", "-Command", "Set-Service -Name 'HidServ' -StartupType Automatic; Start-Service -Name 'HidServ'"],
            check=True,
            shell=True,
            capture_output=True,
            text=True
        )
        print("HID Service enabled successfully.")
        print(f"PowerShell Output: {result.stdout}")
    except subprocess.CalledProcessError as e:
        print(f"Failed to enable HID Service. Error: {e}")
        print(f"PowerShell Error Output: {e.stderr}")

# Function to toggle HID service based on checkbox
def toggle_hid_service():
    print(f"DEBUG: toggle_hid_service() called. Checkbox state: {hid_var.get()}")
    if hid_var.get():
        print("DEBUG: Disabling HID service...")
        disable_hid_service()  # Disable HID service
    else:
        print("DEBUG: Enabling HID service...")
        enable_hid_service()  # Enable HID service

# GUI Setup
root = tk.Tk()
root.title("AudioLock")
root.geometry("300x300")  

tk.Label(root, text="Set Volume Level (%)", font=("System", 12)).pack(pady=10)

# Volume Slider
volume_slider = tk.Scale(
    root, 
    from_=0, 
    to=100, 
    orient="horizontal", 
    command=set_volume, 
    length=150,
    sliderlength=10 
)
volume_slider.set(int(volume.GetMasterVolumeLevelScalar() * 100))
volume_slider.pack(pady=10)

# Lock/Unlock Button
lock_enabled = False
lock_button = tk.Button(root, text="Lock Volume", command=toggle_lock, font=("System", 10))
lock_button.pack(pady=10)

# HID Service Toggle Checkbox
hid_var = tk.BooleanVar()
hid_var.trace_add("write", lambda *args: toggle_hid_service())  # Call toggle_hid_service() on change
hid_checkbox = tk.Checkbutton(root, text="Disable HID Service", variable=hid_var)
hid_checkbox.pack(pady=10)

# Start keyboard listener in background
keyboard_thread = threading.Thread(target=listen_for_keys, daemon=True)
keyboard_thread.start()

root.mainloop()
