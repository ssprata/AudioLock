import tkinter as tk
import threading
import time
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
from pynput import keyboard

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
    if key == keyboard.Key.media_volume_up:
        volume_slider.set(volume_slider.get() + 2)  # Adjust value
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

# GUI Setup
root = tk.Tk()
root.title("Volume Lock")
root.geometry("300x150")

tk.Label(root, text="Set Volume Level (%)").pack()

# Volume Slider
volume_slider = tk.Scale(root, from_=0, to=100, orient="horizontal", command=set_volume)
volume_slider.set(int(volume.GetMasterVolumeLevelScalar() * 100))
volume_slider.pack()

# Lock/Unlock Button
lock_enabled = False
lock_button = tk.Button(root, text="Lock Volume", command=toggle_lock)
lock_button.pack()

# Start keyboard listener in background
keyboard_thread = threading.Thread(target=listen_for_keys, daemon=True)
keyboard_thread.start()

root.mainloop()
