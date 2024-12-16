import ctypes
from ctypes import wintypes

import win32api, win32con

import pystray
from pystray import MenuItem as item
from PIL import Image

import sys
import os

import winreg

from reg_utils import is_dark_theme, key_exists, create_reg_key
from toast import show_notification
import config

import threading

import time
import keyboard  # Add this import for keyboard hotkey support

import screen_brightness_control as sbc

import json




# MARK: write_excluded_rates_to_registry()
def write_excluded_rates_to_registry(excluded_rates):
    try:
        excluded_rates_str = ",".join(map(str, excluded_rates))

        if not key_exists(config.REGISTRY_PATH):
            create_reg_key(config.REGISTRY_PATH)

        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, config.REGISTRY_PATH, 0, winreg.KEY_WRITE)
        winreg.SetValueEx(key, "ExcludedHzRates", 0, winreg.REG_SZ, excluded_rates_str)
        winreg.CloseKey(key)
    except Exception as e:
        print(f"Error writing to registry: {e}")

# MARK: read_excluded_rates_from_registry()
def read_excluded_rates_from_registry():
    try:
        if not key_exists(config.REGISTRY_PATH):
            create_reg_key(config.REGISTRY_PATH)

        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, config.REGISTRY_PATH, 0, winreg.KEY_READ)
        
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

    sbc_info = sbc.list_monitors_info()
    for index, monitor in enumerate(monitors):
        monitor["name"] = sbc_info[index]["name"]
        monitor["model"] = sbc_info[index]["model"]
        monitor["serial"] = sbc_info[index]["serial"]
        monitor["manufacturer"] = sbc_info[index]["manufacturer"]
        monitor["manufacturer_id"] = sbc_info[index]["manufacturer_id"]

        if monitor["manufacturer"] is None:
            monitor["display_name"] = f"DISPLAY{index + 1}"
        else:
            monitor["display_name"] = f"{monitor["manufacturer"]} ({index + 1})"

    # for monitor in monitors:
    #     print(monitor)
    print("get_monitors_info()")

    return monitors



# MARK: change_refresh_rate()
def change_refresh_rate(monitor, refresh_rate):

    if refresh_rate == monitor["RefreshRate"]:
        print(f"Monitor {monitor['Device']} is already set to {refresh_rate} Hz.")
        return

    device = monitor["Device"]

    devmode = win32api.EnumDisplaySettings(device, win32con.ENUM_CURRENT_SETTINGS)
    devmode.DisplayFrequency = refresh_rate
    result = win32api.ChangeDisplaySettingsEx(device, devmode)

    if result == win32con.DISP_CHANGE_SUCCESSFUL:
        print(f"Successfully changed the refresh rate of {device} to {refresh_rate} Hz.")
    else:
        print(f"Failed to change the refresh rate of {device}.")


# MARK: change_refresh_rate_with_brightness_restore()
def change_refresh_rate_with_brightness_restore(monitor, refresh_rate, refresh=False):
    
    brightness_before = sbc.get_brightness(display=monitor["name"])
    # print(f"Current brightness of {monitor['name']}: {brightness_before}")

    change_refresh_rate(monitor, refresh_rate)

    # Refresh the icon menu
    if refresh:
        icon.menu = pystray.Menu(*create_menu(get_monitors_info()))

    # Restore brightness in a separate thread
    def restore_brightness():
        time.sleep(5)
        brightness_after = sbc.get_brightness(display=monitor["name"])
        if brightness_after != brightness_before:
            sbc.set_brightness(*brightness_before, display=monitor["name"])
            print(f"Restored brightness of {monitor['name']} from {brightness_after} to {brightness_before}")

    threading.Thread(target=restore_brightness).start()



# MARK: refresh_tray()
def refresh_tray():
    icon.menu = pystray.Menu(*create_menu(get_monitors_info()))


# MARK: toggle_excluded_rate_ext()
def toggle_excluded_rate_ext(rate):
        excluded_rates = read_excluded_rates_from_registry()

        if rate in excluded_rates:
            excluded_rates.remove(rate)
        else:
            excluded_rates.append(rate)

        write_excluded_rates_to_registry(excluded_rates)

        # Refresh tray menu
        icon.menu = pystray.Menu(*create_menu(get_monitors_info()))
        # refresh_tray()



# MARK: write_preset_to_registry()
def write_preset_to_registry(preset_name, serial, refresh_rate):
    try:
        if not key_exists(config.REGISTRY_PATH):
            create_reg_key(config.REGISTRY_PATH)

        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, config.REGISTRY_PATH, 0, winreg.KEY_WRITE)
        winreg.SetValueEx(key, f"Preset_{preset_name}", 0, winreg.REG_SZ, f"{serial},{refresh_rate}")
        winreg.CloseKey(key)
    except Exception as e:
        print(f"Error writing preset to registry: {e}")

# MARK: read_presets_from_registry()
def read_presets_from_registry():
    presets = {}
    try:
        if not key_exists(config.REGISTRY_PATH):
            create_reg_key(config.REGISTRY_PATH)

        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, config.REGISTRY_PATH, 0, winreg.KEY_READ)
        i = 0
        while True:
            try:
                value_name, value_data, _ = winreg.EnumValue(key, i)
                if value_name.startswith("Preset_"):
                    preset_name = value_name.replace("Preset_", "")
                    serial, refresh_rate = value_data.split(",")
                    presets[preset_name] = {"serial": serial, "refresh_rate": int(refresh_rate)}
                i += 1
            except OSError:
                break
        winreg.CloseKey(key)
    except Exception as e:
        print(f"Error reading presets from registry: {e}")

    print(f"read_presets_from_registry(): {presets}")
    return presets



# MARK: read_profiles_from_reg()
def read_profiles_from_reg():
        registry_path = r"Software\WRRS\Settings"

        def get_profile_value(profile_name):
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, registry_path) as key:
                    json_data = winreg.QueryValueEx(key, profile_name)[0]
                    return json.loads(json_data)
            except (FileNotFoundError, OSError, json.JSONDecodeError):
                return {}

        p1 = get_profile_value("Profile1")
        p2 = get_profile_value("Profile2")
        p3 = get_profile_value("Profile3")

        return p1, p2, p3



# MARK: set_profile()
def set_profile(preset):

    monitors_info = get_monitors_info()

    for monitor_p in preset:
        for monitor_i in monitors_info:
            if monitor_i["serial"] == monitor_p["serial"]:
                change_refresh_rate_with_brightness_restore(monitor_i, monitor_p["RefreshRate"])
                break
    icon.menu = pystray.Menu(*create_menu(get_monitors_info()))



# MARK: create_menu()
def create_menu(monitors_info):
    
    def change_rate_action(monitor, rate, refresh=True):
        return lambda _: change_refresh_rate_with_brightness_restore(monitor, rate, refresh)


    def refresh_action():
        return lambda _: refresh_tray()
    
    def toggle_excluded_rate(rate):
        return lambda _: toggle_excluded_rate_ext(rate)
    

    excluded_rates = read_excluded_rates_from_registry()
    # excluded_rates.sort()  # Sort the rates in ascending order

    monitor_menu = []

    for monitor in monitors_info:

        # Add monitor name
        # monitor_name = monitor['Device'].replace("\\\\.\\", "")  # Видаляємо \\.\ з початку рядка
        monitor_name = monitor["display_name"]

        monitor_menu.append(pystray.MenuItem(
            text = monitor_name,
            action = None,
            enabled=False
        ))

        for rate in monitor['AvailableRefreshRates']:
            if rate not in excluded_rates:
                # Add checkmark or radio button for current refresh rate
                is_current_rate = (rate == monitor['RefreshRate'])
                monitor_menu.append(pystray.MenuItem(
                    text = f"{rate} Hz",
                    action = change_rate_action(monitor, rate),
                    checked=lambda item, is_current_rate=is_current_rate: is_current_rate
                ))
        # Add a separator after each monitor's refresh rates
        monitor_menu.append(pystray.Menu.SEPARATOR)


    monitor_menu.append(pystray.Menu.SEPARATOR)

    profile_1, profile_2, profile_3 = read_profiles_from_reg()
    
    if profile_1:
        monitor_menu.append(pystray.MenuItem(
            text = f"Profile1 (Ctrl+Alt+1)",
            action = lambda _: set_profile(profile_1),
        ))

    if profile_2:
        monitor_menu.append(pystray.MenuItem(
            text = f"Profile2 (Ctrl+Alt+2)",
            action = lambda _: set_profile(profile_2),
        ))

    if profile_3:
        monitor_menu.append(pystray.MenuItem(
            text = f"Profile3 (Ctrl+Alt+3)",
            action = lambda _: set_profile(profile_3),
        ))


    def save_profile(preset_number):
        presets = []
        for monitor in monitors_info:
            presets.append({
                "serial": monitor["serial"],
                "RefreshRate": monitor["RefreshRate"]
            })

        json_data = json.dumps(presets)

        # Шлях до реєстру
        registry_path = r"Software\WRRS\Settings"
        key_name = f"Profile{preset_number}"

        # Запис у реєстр
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, registry_path) as key:
            winreg.SetValueEx(key, key_name, 0, winreg.REG_SZ, json_data)
        
        set_hotkeys()
        icon.menu = pystray.Menu(*create_menu(get_monitors_info()))


    def clear_all_profiles():
        registry_path = r"Software\WRRS\Settings"
        profile_keys = "Profile1", "Profile2", "Profile3"

        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, registry_path) as key:
            for profile_key in profile_keys:
                try:
                    winreg.DeleteValue(key, profile_key)
                except FileNotFoundError:
                    print(f"{profile_key} does not exist in the registry.")
        
        set_hotkeys()
        icon.menu = pystray.Menu(*create_menu(get_monitors_info()))


    monitor_menu.append(pystray.Menu.SEPARATOR)

    # Add refresh option
    monitor_menu.append(pystray.MenuItem(
        text = "Refresh",
        # refresh_action(),
        action = lambda _: refresh_tray(),
        default=True
    ))

    all_rates = set()
    for monitor in monitors_info:
        all_rates.update(monitor['AvailableRefreshRates'])
    all_rates = sorted(all_rates)

    # Add a dropdown menu for all available refresh rates
    all_rates_menu = []

    all_rates_menu.append(pystray.MenuItem(text = "Exclude", action = None, enabled=False))

    for rate in all_rates:
        all_rates_menu.append(pystray.MenuItem(
            f"{rate} Hz",
            checked=lambda item, rate=rate: rate not in excluded_rates,
            action=toggle_excluded_rate(rate)
        ))
    
    all_rates_menu.append(pystray.Menu.SEPARATOR)
    all_rates_menu.append(pystray.MenuItem(text = "Save to Profile1", action = lambda _: save_profile(1)))
    all_rates_menu.append(pystray.MenuItem(text = "Save to Profile2", action = lambda _: save_profile(2)))
    all_rates_menu.append(pystray.MenuItem(text = "Save to Profile3", action = lambda _: save_profile(3)))
    all_rates_menu.append(pystray.Menu.SEPARATOR)
    all_rates_menu.append(pystray.MenuItem("Clear all profiles", clear_all_profiles))


    monitor_menu.append(pystray.MenuItem(
        "Options",
        pystray.Menu(*all_rates_menu)
    ))


    # Add exit option
    monitor_menu.append(pystray.MenuItem(
        "Quit",
        lambda _: icon.stop()
    ))

    return monitor_menu


def set_hotkeys():
    profile_1, profile_2, profile_3 = read_profiles_from_reg()
    print("set_hotkeys()")

    if profile_1: 
        if 'ctrl+alt+1' not in keyboard._hotkeys:
            keyboard.add_hotkey('ctrl+alt+1', lambda: set_profile(profile_1))
            # keyboard.add_hotkey('ctrl+alt+1', lambda: print(1))
            print("Hotkey 1 added")
    elif 'ctrl+alt+1' in keyboard._hotkeys:
        keyboard.remove_hotkey('ctrl+alt+1')
        print("Hotkey 1 removed")


    if profile_2: 
        if 'ctrl+alt+2' not in keyboard._hotkeys:
            keyboard.add_hotkey('ctrl+alt+2', lambda: set_profile(profile_2))
            # keyboard.add_hotkey('ctrl+alt+2', lambda: print(2))
            print("Hotkey 2 added")
    elif 'ctrl+alt+2' in keyboard._hotkeys:
        keyboard.remove_hotkey('ctrl+alt+2')
        print("Hotkey 2 removed")


    if profile_3: 
        if 'ctrl+alt+3' not in keyboard._hotkeys:
            keyboard.add_hotkey('ctrl+alt+3', lambda: set_profile(profile_3))
            # keyboard.add_hotkey('ctrl+alt+3', lambda: print(3))
            print("Hotkey 3 added")
    elif 'ctrl+alt+3' in keyboard._hotkeys:
        keyboard.remove_hotkey('ctrl+alt+3')
        print("Hotkey 3 removed")



# MARK: Main
if __name__ == "__main__":

    if getattr(sys, 'frozen', False):
        # Якщо програма запущена як EXE, шлях до іконки відносно до виконуваного файлу
        icon_path = os.path.join(sys._MEIPASS, 'icon_color.ico')
        # if is_dark_theme():
        #     icon_path = os.path.join(sys._MEIPASS, 'icon_light.ico')
        # else:
        #     icon_path = os.path.join(sys._MEIPASS, 'icon_dark.ico')
    else:
        # Якщо програма запущена з Python, використовуємо поточну директорію
        icon_path = 'icons/icon_color_dev.ico'
        # if is_dark_theme():
        #     icon_path = 'icons/icon_light.ico' 
        # else:
        #     icon_path = 'icons/icon_dark.ico'

    icon_image = Image.open(icon_path)

    monitors_info = get_monitors_info()
    # Create system tray icon
    icon = pystray.Icon(name="WRRS",
                        title=f"Refresh Rate Switcher v{config.version}",
                        icon=icon_image,
                        menu=pystray.Menu(*create_menu(monitors_info))
                        )

    set_hotkeys()
 
    icon.run()






