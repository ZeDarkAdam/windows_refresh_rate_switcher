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
import config

from winotify import Notification, audio

import threading

import time
import keyboard  # Add this import for keyboard hotkey support




import screen_brightness_control as sbc
















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


    for monitor in monitors:
        print(monitor)


    return monitors








# MARK: change_refresh_rate()
def change_refresh_rate(monitor, refresh_rate):

    device = monitor["Device"]
    # print(device, refresh_rate)

    devmode = win32api.EnumDisplaySettings(device, win32con.ENUM_CURRENT_SETTINGS)
    devmode.DisplayFrequency = refresh_rate
    result = win32api.ChangeDisplaySettingsEx(device, devmode)

    if result == win32con.DISP_CHANGE_SUCCESSFUL:
        print(f"Successfully changed the refresh rate of {device} to {refresh_rate} Hz.")

        # Refresh the icon menu
        icon.menu = pystray.Menu(*create_menu(get_monitors_info()))
    else:
        print(f"Failed to change the refresh rate of {device}.")


# MARK: change_refresh_rate_with_brightness_restore()
def change_refresh_rate_with_brightness_restore(monitor, refresh_rate):
    
    brightness_before = sbc.get_brightness(display=monitor["name"])
    print(f"Current brightness of {monitor['name']}: {brightness_before}")


    change_refresh_rate(monitor, refresh_rate)

    # Restore brightness
    # time.sleep(5)
    # sbc.set_brightness(*brightness, display=monitor["name"])

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

def print_test():
    print("Test")














# MARK: show_message()
def show_message(title, message):

    # icon_path = "C:\Projects\GitHub\windows_refresh_rate_switcher\icons\icon_color.ico"

    toast = Notification(
        app_id = 'Refresh Rate Switcher',
        title = title,
        msg = message,
        duration = "short",
        # icon = icon_path,
    )
    toast.set_audio(audio.Default, loop=False)
    toast.show()












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











# MARK: create_menu()
def create_menu(monitors_info):
    
    def change_rate_action(monitor, rate):
        return lambda _: change_refresh_rate_with_brightness_restore(monitor, rate)

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
















    def toggle_second_monitor_rate():
        second_monitor = monitors_info[1] if len(monitors_info) > 1 else None
        if second_monitor:
            current_rate = second_monitor['RefreshRate']
            new_rate = 72 if current_rate == 60 else 60
            change_refresh_rate(second_monitor['Device'], new_rate)
            show_message("Refresh Rate Switcher", f"Changed 2nd monitor refresh rate to {new_rate} Hz.")
            # threading.Thread(target=show_message, args=("Refresh Rate Switcher", f"Changed 2nd monitor refresh rate to {new_rate} Hz.")).start()

    # Add toggle option for second monitor's refresh rate
    monitor_menu.append(pystray.MenuItem(
        "Toggle 2nd Monitor 60/72 Hz",
        lambda _: toggle_second_monitor_rate(),
        # default=True
    ))





















    def save_preset_action(monitor):
        return lambda _: save_preset(monitor)

    def load_preset_action(preset_name):
        return lambda _: load_preset(preset_name)

    def save_preset(monitor):
        preset_name = f"Preset_{monitor['display_name']}"
        write_preset_to_registry(preset_name, monitor["serial"], monitor["RefreshRate"])
        show_message("Refresh Rate Switcher", f"Preset '{preset_name}' saved.")

    def load_preset(preset_name):
        presets = read_presets_from_registry()
        if preset_name in presets:
            preset = presets[preset_name]
            for monitor in monitors_info:
                if monitor["serial"] == preset["serial"]:
                    change_refresh_rate_with_brightness_restore(monitor, preset["refresh_rate"])
                    show_message("Refresh Rate Switcher", f"Preset '{preset_name}' loaded.")
                    break

    # Add save preset option for each monitor
    for monitor in monitors_info:
        monitor_menu.append(pystray.MenuItem(
            f"Save Preset for {monitor['display_name']}",
            save_preset_action(monitor)
        ))

    # Add load preset options
    presets = read_presets_from_registry()
    for preset_name in presets.keys():
        monitor_menu.append(pystray.MenuItem(
            f"Load {preset_name}",
            load_preset_action(preset_name)
        ))























    # monitor_menu.append(pystray.Menu.SEPARATOR)

    # Add exit option
    monitor_menu.append(pystray.MenuItem(
        "Quit",
        lambda _: icon.stop()
    ))

    return monitor_menu






# MARK: set_all_monitors_to_60hz()
def set_all_monitors_to_60hz():
    monitors_info = get_monitors_info()
    for monitor in monitors_info:
        change_refresh_rate_with_brightness_restore(monitor, 60)
    show_message("Refresh Rate Switcher", "All monitors set to 60 Hz.")










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



    # Register global hotkey (Ctrl+Alt+S) to set all monitors to 60 Hz
    keyboard.add_hotkey('ctrl+alt+1', set_all_monitors_to_60hz)




    icon.run()





    

    

    # m_info = get_monitors_info()

    # for monitor in m_info:
    #     print(monitor)






