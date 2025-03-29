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
from PIL import Image, ImageTk  

def run_as_admin():
    """Restart the script with elevated privileges"""
    try:
        if ctypes.windll.shell32.IsUserAnAdmin():
            print("DEBUG: Script is running as administrator.")
            return 

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
root.geometry("320x250")

# Set the background color of the root window
root.configure(bg="#07003a")

# Get the path to the .ico file (works for both .py and .exe)
base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
icon_path = os.path.join(base_path, "icon.ico")

# Load the .ico file
icon_image = Image.open(icon_path)
icon_photo = ImageTk.PhotoImage(icon_image)

# Set the icon for the application window
root.iconphoto(True, icon_photo)

# Volume Slider Label
volume_label = tk.Label(root, text="Set Volume Level: 50%", font=("System", 12), bg="#07003a", fg="white")
volume_label.pack(pady=5)

# Volume Slider
def update_volume_label(value):
    volume_label.config(text=f"Set Volume Level: {value}%", bg="#07003a", fg="white")
    set_volume(value)

volume_slider = tk.Scale(
    root, 
    from_=0, 
    to=100, 
    orient="horizontal", 
    command=update_volume_label,  # Update label dynamically
    length=150,
    sliderlength=10,
    bg="#07003a",  # Background color of the slider
    fg="white",    # Text color of the slider
    highlightbackground="#07003a",  # Border color of the slider
    troughcolor="#333366"  # Color of the slider trough
)
volume_slider.set(int(volume.GetMasterVolumeLevelScalar() * 100))
volume_slider.pack(pady=10)

# Lock/Unlock Button
lock_enabled = False
lock_button = tk.Button(root, text="Lock Volume", command=toggle_lock, font=("System", 10), bg="#07003a", fg="white")
lock_button.pack(pady=10)

# HID Service Toggle Checkbox
hid_var = tk.BooleanVar()
hid_var.trace_add("write", lambda *args: toggle_hid_service())  # Call toggle_hid_service() on change
hid_checkbox = tk.Checkbutton(root, text="Disable HID Service", variable=hid_var, bg="#07003a", fg="white", selectcolor="#333366")
hid_checkbox.pack(pady=5)

# Add a frame to group the "Developed by:" and "@sprata" labels
developer_frame = tk.Frame(root, bg="#07003a")
developer_frame.pack(side="bottom", anchor="se", padx=2, pady=2)

# Add "Developed by:" in grey with a smaller font size
developer_label = tk.Label(developer_frame, text="Developed by:", font=("Arial", 8), fg="grey", bg="#07003a")
developer_label.pack(side="left")

# Add "@sprata" in the default color with a slightly larger font size
name_label = tk.Label(developer_frame, text="@sprata", font=("System", 10), fg="white", bg="#07003a")
name_label.pack(side="left")

# Start keyboard listener in background
keyboard_thread = threading.Thread(target=listen_for_keys, daemon=True)
keyboard_thread.start()

root.mainloop()
