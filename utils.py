import win32api
import win32con

def is_dark_theme():
    try:
        key = win32api.RegOpenKeyEx(win32con.HKEY_CURRENT_USER, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize', 0, win32con.KEY_READ)
        value, _ = win32api.RegQueryValueEx(key, 'AppsUseLightTheme')
        win32api.RegCloseKey(key)
        return value == 0
    except Exception:
        return False