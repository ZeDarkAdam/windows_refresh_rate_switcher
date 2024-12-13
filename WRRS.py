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


import customtkinter as ctk




















import winreg

# Шлях до реєстру
REGISTRY_PATH = r"Software\WRRS\Settings"

# Перевірка, чи існує ключ реєстру
def key_exists():
    try:
        winreg.OpenKey(winreg.HKEY_CURRENT_USER, REGISTRY_PATH, 0, winreg.KEY_READ)
        return True
    except FileNotFoundError:
        return False

# Створення ключа реєстру
def create_registry_key():
    try:
        winreg.CreateKey(winreg.HKEY_CURRENT_USER, REGISTRY_PATH)
    except Exception as e:
        print(f"Error creating registry key: {e}")




# Збереження масиву виключених герцовок
def write_excluded_rates_to_registry(excluded_rates):
    try:
        excluded_rates_str = ",".join(map(str, excluded_rates))
        if not key_exists():
            create_registry_key()
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REGISTRY_PATH, 0, winreg.KEY_WRITE)
        winreg.SetValueEx(key, "ExcludedHzRates", 0, winreg.REG_SZ, excluded_rates_str)
        winreg.CloseKey(key)
    except Exception as e:
        print(f"Error writing to registry: {e}")

# Читання масиву виключених герцовок з реєстру
def read_excluded_rates_from_registry():
    try:
        if not key_exists():
            create_registry_key()
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REGISTRY_PATH, 0, winreg.KEY_READ)
        value, _ = winreg.QueryValueEx(key, "ExcludedHzRates")
        winreg.CloseKey(key)
        excluded_rates = list(map(int, value.split(",")))
        return excluded_rates
    
    except Exception as e:
        print(f"Error reading from registry: {e}")
        return []



# Додавання нового елемента в інтерфейс
def add_item_to_frame(item, frame, label_list, button_list):
    frame.grid_columnconfigure(1, weight=1)  # Ensure the button is aligned to the right
    label = ctk.CTkLabel(frame, text=f"{item} Hz", padx=5, anchor="w")
    button = ctk.CTkButton(frame, text="Remove", width=80, height=24, command=lambda: remove_item_from_frame(item, frame, label, button, label_list, button_list))

    label.grid(row=len(label_list), column=0, pady=(0, 10), sticky="w")
    button.grid(row=len(button_list), column=1, pady=(0, 10), padx=5, sticky="e")

    label_list.append(label)
    button_list.append(button)

# Оновлення реєстру після видалення елемента
def update_registry(label_list):
    global excluded_rates
    excluded_rates = [int(label.cget("text").replace(" Hz", "")) for label in label_list]
    write_excluded_rates_to_registry(excluded_rates)

# Видалення елемента з інтерфейсу
def remove_item_from_frame(item, frame, label, button, label_list, button_list):
    label.destroy()
    button.destroy()
    label_list.remove(label)
    button_list.remove(button)
    update_registry(label_list)

    refresh_monitors()  # Refresh monitors to update tray menu






























# MARK: Constants
version = "0.1.3"
# excluded_rates = {23, 24, 50, 56, 59, 67, 70, 71, 119}
# excluded_rates = {23, 56, 59, 67, 70, 71, 119}
excluded_rates = read_excluded_rates_from_registry()



















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
    info_text += f"WRRS v{version}"
    info_text += "\n\n"

    

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
    threading.Thread(target=MessageBoxW, args=(None, info_text, "Info", 0)).start()














# MARK: show_info_ctk()
# def show_info_ctk(monitors_info):
#     # create a new window
#     root = ctk.CTk()
#     root.title(f"Windows Refresh Rate Switcher v{version}")
    
#     # set the window size
#     window_width = 500
#     window_height = 400

#     # Get the screen width and height
#     screen_width = root.winfo_screenwidth()
#     screen_height = root.winfo_screenheight()

#     # Calculate the x and y coordinates for the Tk root window
#     x_coordinate = (screen_width - window_width) // 2
#     y_coordinate = (screen_height - window_height) // 2

#     # Set the window size and position
#     root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

#     # Create a text box to display the information
#     info_text = ctk.CTkTextbox(root, width=550, height=350, wrap="word")
#     info_text.pack(pady=10, padx=10)

#     # Form the text information about the monitors
#     display_text = ""
#     for index, monitor in enumerate(monitors_info):
#         display_text += f"Monitor {index + 1}:\n"
#         display_text += f"  Resolution: {monitor['Resolution']}\n"
#         display_text += f"  Refresh Rate: {monitor['RefreshRate']} Hz\n"
#         display_text += f"  Available Refresh Rates: {', '.join(map(str, monitor['AvailableRefreshRates']))} Hz\n"
#         display_text += "\n"

#     # display_text += f"Excluded Refresh Rates: {', '.join(map(str, sorted(excluded_rates)))} Hz"

#     # Add the text to the text box
#     info_text.insert("1.0", display_text)
#     info_text.configure(state="disabled")  # Make the text box read-only




#     # -----------------------------------------------------------------------------------------------


#     # Список елементів
#     label_list = []
#     button_list = []


#     # Функція для додавання нової герцовки
#     def add_rate():
#         try:
#             rate = int(entry.get())
#             if rate not in [int(label.cget("text").replace(" Hz", "")) for label in label_list]:
#                 add_item_to_frame(rate, scroll_frame, label_list, button_list)
#                 entry.delete(0, ctk.END)
#                 update_registry(label_list)
#                 refresh_list()  # Refresh the list to maintain sorted order
#                 refresh_monitors()  # Refresh monitors to update tray menu
#             else:
#                 print("Rate already exists.")

#         except ValueError:
#             print("Invalid input. Please enter a valid number.")

#     # Функція для оновлення списку
#     def refresh_list():
#         for label in label_list:
#             label.destroy()
#         for button in button_list:
#             button.destroy()
#         label_list.clear()
#         button_list.clear()
#         excluded_rates = read_excluded_rates_from_registry()
#         excluded_rates.sort()  # Sort the rates in ascending order
#         for rate in excluded_rates:
#             add_item_to_frame(rate, scroll_frame, label_list, button_list)

#     # Frame для entry і add_button
#     input_frame = ctk.CTkFrame(root)
#     input_frame.pack(pady=10)

#     # Поле введення для нової герцовки
#     entry = ctk.CTkEntry(input_frame, placeholder_text="Enter Hz rate")
#     entry.grid(row=0, column=0, padx=(0, 10))

#     # Кнопка для додавання нової герцовки
#     add_button = ctk.CTkButton(input_frame, text="Add", command=add_rate, width=10)
#     add_button.grid(row=0, column=1)

#     # Створення скроллінгового фрейму
#     scroll_frame = ctk.CTkScrollableFrame(root)
#     scroll_frame.pack(padx=20, pady=20, fill="both", expand=True)

#     # Завантаження існуючих герцовок з реєстру
#     excluded_rates = read_excluded_rates_from_registry()
#     excluded_rates.sort()  # Sort the rates in ascending order
#     for rate in excluded_rates:
#         add_item_to_frame(rate, scroll_frame, label_list, button_list)


#     # -----------------------------------------------------------------------------------------------

#     # Start the event loop
#     root.mainloop()

















def show_info_ctk(monitors_info):
    # Create a new window
    root = ctk.CTk()
    root.title(f"Windows Refresh Rate Switcher v{version}")
    
    # Set the window size
    window_width = 700
    window_height = 450

    # Get the screen width and height
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Calculate the x and y coordinates for the Tk root window
    x_coordinate = (screen_width - window_width) // 2
    y_coordinate = (screen_height - window_height) // 2

    # Set the window size and position
    root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

    # Configure the root window to use grid layout
    root.grid_columnconfigure(0, weight=1)  # Allow the left frame to expand
    root.grid_columnconfigure(1, weight=1)  # Allow the right frame to expand
    root.grid_rowconfigure(0, weight=1)     # Allow the row to expand vertically

    # Create a frame for the left section (textbox)
    left_frame = ctk.CTkFrame(root)
    left_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

    # Create a text box to display the information
    info_text = ctk.CTkTextbox(left_frame, wrap="word")
    info_text.pack(pady=10, padx=10, fill="both", expand=True)

    # Form the text information about the monitors
    display_text = ""
    for index, monitor in enumerate(monitors_info):
        display_text += f"Monitor {index + 1}:\n"
        display_text += f"  Resolution: {monitor['Resolution']}\n"
        display_text += f"  Refresh Rate: {monitor['RefreshRate']} Hz\n"
        display_text += f"  Available Refresh Rates: {', '.join(map(str, monitor['AvailableRefreshRates']))} Hz\n"
        display_text += "\n"

    # Add the text to the text box
    info_text.insert("1.0", display_text)
    info_text.configure(state="disabled")  # Make the text box read-only

    # Create a frame for the right section (controls and buttons)
    right_frame = ctk.CTkFrame(root)
    right_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)

    # List for labels and buttons
    label_list = []
    button_list = []

    # Function to add a new refresh rate
    def add_rate():
        try:
            rate = int(entry.get())
            if rate not in [int(label.cget("text").replace(" Hz", "")) for label in label_list]:
                add_item_to_frame(rate, scroll_frame, label_list, button_list)
                entry.delete(0, ctk.END)
                update_registry(label_list)
                refresh_list()  # Refresh the list to maintain sorted order
                refresh_monitors()  # Refresh monitors to update tray menu
            else:
                print("Rate already exists.")
        except ValueError:
            print("Invalid input. Please enter a valid number.")

    # Function to refresh the list of excluded rates
    def refresh_list():
        for label in label_list:
            label.destroy()
        for button in button_list:
            button.destroy()
        label_list.clear()
        button_list.clear()
        excluded_rates = read_excluded_rates_from_registry()
        excluded_rates.sort()  # Sort the rates in ascending order
        for rate in excluded_rates:
            add_item_to_frame(rate, scroll_frame, label_list, button_list)

    # Create an input frame for the entry and add button
    input_frame = ctk.CTkFrame(right_frame)
    input_frame.pack(pady=10)

    # Entry for new refresh rate
    entry = ctk.CTkEntry(input_frame, placeholder_text="Enter Hz rate")
    entry.grid(row=0, column=0, padx=(0, 10))

    # Button to add a new refresh rate
    add_button = ctk.CTkButton(input_frame, text="Add to exclude list", command=add_rate, width=10)
    add_button.grid(row=0, column=1)

    # Create a scrollable frame for the excluded rates
    scroll_frame = ctk.CTkScrollableFrame(right_frame)
    scroll_frame.pack(padx=20, pady=20, fill="both", expand=True)

    # Load existing refresh rates from registry
    excluded_rates = read_excluded_rates_from_registry()
    excluded_rates.sort()  # Sort the rates in ascending order
    for rate in excluded_rates:
        add_item_to_frame(rate, scroll_frame, label_list, button_list)

    # Start the event loop
    root.mainloop()

























# MARK: create_menu()
def create_menu(monitors_info):
    def change_rate_action(device, rate):
        return lambda _: change_refresh_rate(device, rate)

    def show_info_action():
        # return lambda _: show_info_message_box(monitors_info)
        return lambda _: threading.Thread(target=show_info_ctk, args=(monitors_info,)).start()

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
        # Add monitor name
        monitor_menu.append(pystray.MenuItem(
            f"{monitor['Device']}",
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
        "Exit",
        lambda _: icon.stop()
    ))

    return monitor_menu





# MARK: Main
if __name__ == "__main__":

    # Load icon image
    icon_image = Image.open(icon_path)

    monitors_info = get_monitors_info()
    # Create system tray icon
    icon = pystray.Icon(name="Monitor Refresh Rate Switcher", 
                        icon=icon_image, 
                        title="Refresh Rate Switcher")

    icon.menu = pystray.Menu(*create_menu(monitors_info))

    icon.run()






