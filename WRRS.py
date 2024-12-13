import ctypes
from ctypes import wintypes

import win32api, win32con

import pystray
from pystray import MenuItem as item
from PIL import Image

import sys
import os

import winreg

from utils import is_dark_theme, key_exists, create_registry_key



# MARK: Constants
version = "0.2.0"
REGISTRY_PATH = r"Software\WRRS\Settings"



# MARK: write_excluded_rates_to_registry()
def write_excluded_rates_to_registry(excluded_rates):
    try:
        excluded_rates_str = ",".join(map(str, excluded_rates))

        if not key_exists(REGISTRY_PATH):
            create_registry_key(REGISTRY_PATH)

        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REGISTRY_PATH, 0, winreg.KEY_WRITE)
        winreg.SetValueEx(key, "ExcludedHzRates", 0, winreg.REG_SZ, excluded_rates_str)
        winreg.CloseKey(key)
    except Exception as e:
        print(f"Error writing to registry: {e}")



# MARK: read_excluded_rates_from_registry()
def read_excluded_rates_from_registry():
    try:
        if not key_exists(REGISTRY_PATH):
            create_registry_key(REGISTRY_PATH)

        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REGISTRY_PATH, 0, winreg.KEY_READ)
        
        try:
            value, _ = winreg.QueryValueEx(key, "ExcludedHzRates")
            excluded_rates = list(map(int, value.split(",")))
        except FileNotFoundError:
            excluded_rates = []

        winreg.CloseKey(key)

        # print(f"read_excluded_rates_from_registry(): {excluded_rates}")

        return excluded_rates
    
    except Exception as e:
        print(f"Error reading from registry: {e}")
        return []



# MARK: get_available_refresh_rates()
def get_available_refresh_rates(device):
    refresh_rates = set()
    i = 0
    while True:
        try:
            devmode = win32api.EnumDisplaySettings(device, i)
            refresh_rates.add(devmode.DisplayFrequency)
            i += 1
        except Exception:
            break
    return sorted(refresh_rates)



# MARK: get_monitors_info()
def get_monitors_info():
    monitors = []

    # Callback function for EnumDisplayMonitors
    def monitor_enum_proc(hMonitor, hdcMonitor, lprcMonitor, dwData):
        monitor_info = win32api.GetMonitorInfo(hMonitor)
        device = monitor_info.get('Device', None)

        if device:
            devmode = win32api.EnumDisplaySettings(device, win32con.ENUM_CURRENT_SETTINGS)
            available_refresh_rates = get_available_refresh_rates(device)
            monitors.append({
                "Device": device,
                "Resolution": f"{devmode.PelsWidth}x{devmode.PelsHeight}",
                "RefreshRate": devmode.DisplayFrequency,
                "AvailableRefreshRates": available_refresh_rates
            })
        return True

    # Define the callback type
    MonitorEnumProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HMONITOR, wintypes.HDC, ctypes.POINTER(wintypes.RECT), wintypes.LPARAM)

    # Load the function from user32.dll
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    enum_display_monitors = user32.EnumDisplayMonitors
    enum_display_monitors.argtypes = [wintypes.HDC, ctypes.POINTER(wintypes.RECT), MonitorEnumProc, wintypes.LPARAM]
    enum_display_monitors.restype = wintypes.BOOL

    # Call EnumDisplayMonitors
    enum_display_monitors(None, None, MonitorEnumProc(monitor_enum_proc), 0)

    return monitors



# MARK: change_refresh_rate()
def change_refresh_rate(device, refresh_rate):
    devmode = win32api.EnumDisplaySettings(device, win32con.ENUM_CURRENT_SETTINGS)
    devmode.DisplayFrequency = refresh_rate
    result = win32api.ChangeDisplaySettingsEx(device, devmode)

    if result == win32con.DISP_CHANGE_SUCCESSFUL:
        print(f"Successfully changed the refresh rate of {device} to {refresh_rate} Hz.")

        # Refresh the icon menu
        icon.menu = pystray.Menu(*create_menu(get_monitors_info()))
    else:
        print(f"Failed to change the refresh rate of {device}.")



# MARK: refresh_monitors()
def refresh_monitors():
    icon.menu = pystray.Menu(*create_menu(get_monitors_info()))



def toggle_excluded_rate_ext(rate):
        
        excluded_rates = read_excluded_rates_from_registry()

        if rate in excluded_rates:
            excluded_rates.remove(rate)
        else:
            excluded_rates.append(rate)

        write_excluded_rates_to_registry(excluded_rates)

        # refresh_monitors()  # Refresh monitors to update tray menu
        # Refresh the icon menu
        icon.menu = pystray.Menu(*create_menu(get_monitors_info()))



# MARK: create_menu()
def create_menu(monitors_info):
    
    def change_rate_action(device, rate):
        return lambda _: change_refresh_rate(device, rate)

    def refresh_action():
        return lambda _: refresh_monitors()
    
    def toggle_excluded_rate(rate):
        return lambda _: toggle_excluded_rate_ext(rate)
    

    excluded_rates = read_excluded_rates_from_registry()
    # excluded_rates.sort()  # Sort the rates in ascending order

    monitor_menu = []

    for monitor in monitors_info:
        # Add monitor name
        monitor_name = monitor['Device'].replace("\\\\.\\", "")  # Видаляємо \\.\ з початку рядка

        monitor_menu.append(pystray.MenuItem(
            monitor_name,
            None,
            enabled=False
        ))
        for rate in monitor['AvailableRefreshRates']:
            if rate not in excluded_rates:
                # Add checkmark or radio button for current refresh rate
                is_current_rate = (rate == monitor['RefreshRate'])
                monitor_menu.append(pystray.MenuItem(
                    f"{rate} Hz",
                    change_rate_action(monitor['Device'], rate),
                    checked=lambda item, is_current_rate=is_current_rate: is_current_rate
                ))
        # Add a separator after each monitor's refresh rates
        monitor_menu.append(pystray.Menu.SEPARATOR)

    # Add refresh option
    monitor_menu.append(pystray.MenuItem(
        "Refresh",
        refresh_action()
    ))

    all_rates = set()
    for monitor in monitors_info:
        all_rates.update(monitor['AvailableRefreshRates'])
    all_rates = sorted(all_rates)

    # Add a dropdown menu for all available refresh rates
    all_rates_menu = []

    for rate in all_rates:
        all_rates_menu.append(pystray.MenuItem(
            f"{rate} Hz",
            checked=lambda item, rate=rate: rate not in excluded_rates,
            action=toggle_excluded_rate(rate)
        ))

    monitor_menu.append(pystray.MenuItem(
        "Exclude",
        pystray.Menu(*all_rates_menu)
    ))

    # monitor_menu.append(pystray.Menu.SEPARATOR)

    # Add exit option
    monitor_menu.append(pystray.MenuItem(
        "Exit",
        lambda _: icon.stop()
    ))

    return monitor_menu



# MARK: Main
if __name__ == "__main__":

    # MARK: Load icon
    if getattr(sys, 'frozen', False):
        # Якщо програма запущена як EXE, шлях до іконки відносно до виконуваного файлу
        if is_dark_theme():
            icon_path = os.path.join(sys._MEIPASS, 'icon_light.png')
        else:
            icon_path = os.path.join(sys._MEIPASS, 'icon_dark.png')
    else: 
        # Якщо програма запущена з Python, використовуємо поточну директорію
        if is_dark_theme():
            icon_path = 'icons/icon_light.png' 
        else:
            icon_path = 'icons/icon_dark.png'

    icon_image = Image.open(icon_path)


    # Create system tray icon
    icon = pystray.Icon(name="Windows Refresh Rate Switcher", 
                        icon=icon_image, 
                        title=f"Refresh Rate Switcher v{version}")

    icon.menu = pystray.Menu(*create_menu(get_monitors_info()))

    icon.run()





