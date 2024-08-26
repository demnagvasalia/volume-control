import os
import sys
import shutil
import keyboard
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL, CoInitialize
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from win32com.client import Dispatch

# Function to get the volume controller
def get_volume_controller():
    CoInitialize()  # Initialize COM library
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    return volume

# Function to play/pause media
def play_pause():
    keyboard.send('play/pause media')

# Function to stop media
def stop_media():
    keyboard.send('stop media')

# Function to skip to the next song
def next_song():
    keyboard.send('media next track')

# Function to go to the previous song
def previous_song():
    keyboard.send('media previous track')

# Function to increase volume
def volume_up():
    volume = get_volume_controller()
    current_volume = volume.GetMasterVolumeLevelScalar()
    volume.SetMasterVolumeLevelScalar(min(1.0, current_volume + 0.05), None)  # Increase by 5%

# Function to decrease volume
def volume_down():
    volume = get_volume_controller()
    current_volume = volume.GetMasterVolumeLevelScalar()
    volume.SetMasterVolumeLevelScalar(max(0.0, current_volume - 0.05), None)  # Decrease by 5%

# Function to create a shortcut in the Startup folder
def add_to_startup():
    startup_path = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
    script_path = os.path.realpath(sys.argv[0])
    shortcut_path = os.path.join(startup_path, 'AudioControl.lnk')
    
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.TargetPath = script_path
    shortcut.WorkingDirectory = os.path.dirname(script_path)
    shortcut.IconLocation = script_path
    shortcut.save()

# Bind hotkeys
keyboard.add_hotkey('ctrl+space', play_pause)
keyboard.add_hotkey('ctrl+shift+space', stop_media)
keyboard.add_hotkey('ctrl+up', volume_up)
keyboard.add_hotkey('ctrl+down', volume_down)
keyboard.add_hotkey('ctrl+right', next_song)
keyboard.add_hotkey('ctrl+left', previous_song)

# Add the script to startup if it's not already there
add_to_startup()

print("Running... Press control + escape + enter to exit.")
keyboard.wait('ctrl+esc+enter')  # Keeps the script running
