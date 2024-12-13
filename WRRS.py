import ctypes
from ctypes import wintypes

import win32api
import win32con
import win32gui

import threading

import pystray
from pystray import MenuItem as item
from PIL import Image

import sys
import os


# import customtkinter as ctk


# MARK: Constants
version = "0.1.4"
# excluded_rates = {23, 24, 50, 56, 59, 67, 70, 71, 119}
excluded_rates = {23, 56, 59, 67, 70, 71, 119}


# Пошук шляху до іконки, якщо програма працює як .exe
if getattr(sys, 'frozen', False):
    # Якщо програма запущена як EXE, шлях до іконки відносно до виконуваного файлу
    icon_path = os.path.join(sys._MEIPASS, 'icon_light.png')
else:
    # Якщо програма запущена з Python, використовуємо поточну директорію
    icon_path = 'icons/icon_light.png'



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
    global monitors_info
    monitors_info = get_monitors_info()
    icon.menu = pystray.Menu(*create_menu(monitors_info))



# MARK: show_monitors_info_message_box()
def show_info_message_box(monitors_info):
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    MessageBoxW = user32.MessageBoxW
    MessageBoxW.argtypes = [wintypes.HWND, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.UINT]


    info_text = ""

    for index, monitor in enumerate(monitors_info):
        info_text += f"Monitor {index + 1}:\n"
        # info_text += f"  Device: {monitor['Device']}\n"
        info_text += f"  Resolution: {monitor['Resolution']}\n"
        info_text += f"  Refresh Rate: {monitor['RefreshRate']} Hz\n"
        info_text += f"  Available Refresh Rates: {', '.join(map(str, monitor['AvailableRefreshRates']))} Hz\n"
        info_text += "\n"

    info_text += f"Excluded Refresh Rates: {', '.join(map(str, sorted(excluded_rates)))} Hz"

    # MessageBoxW(None, info_text, "Monitor Information", 0)
    # Run the message box in a separate thread
    threading.Thread(target=MessageBoxW, args=(None, info_text, f"Windows Refresh Rate Switcher v{version}", 0)).start()



# MARK: create_menu()
def create_menu(monitors_info):
    def change_rate_action(device, rate):
        return lambda _: change_refresh_rate(device, rate)

    def show_info_action():
        return lambda _: show_info_message_box(monitors_info)

    def refresh_action():
        return lambda _: refresh_monitors()


    monitor_menu = []

    # monitor_menu.append(pystray.MenuItem(
    #     f"WRRS v{version}",
    #     None,
    #     enabled=False
    # ))
    # monitor_menu.append(pystray.Menu.SEPARATOR)

    for monitor in monitors_info:

        monitor_name = monitor['Device'].replace("\\\\.\\", "")  # Видаляємо \\.\ з початку рядка
        # Add monitor name
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

    # Add show info option
    monitor_menu.append(pystray.MenuItem(
        "Info",
        show_info_action()
    ))

    monitor_menu.append(pystray.Menu.SEPARATOR)

    # Add exit option
    monitor_menu.append(pystray.MenuItem(
        "Quit",
        lambda _: icon.stop()
    ))

    return monitor_menu





# MARK: Main
if __name__ == "__main__":

    # Load icon image
    icon_image = Image.open(icon_path)

    monitors_info = get_monitors_info()
    # Create system tray icon
    icon = pystray.Icon(name="Windows Refresh Rate Switcher", 
                        icon=icon_image, 
                        title=f"Refresh Rate Switcher v{version}")

    icon.menu = pystray.Menu(*create_menu(monitors_info))

    icon.run()






