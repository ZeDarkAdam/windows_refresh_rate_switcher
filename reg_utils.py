import win32api, win32con

import winreg

def is_dark_theme():
    try:
        key = win32api.RegOpenKeyEx(win32con.HKEY_CURRENT_USER, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize', 0, win32con.KEY_READ)
        value, _ = win32api.RegQueryValueEx(key, 'AppsUseLightTheme')
        win32api.RegCloseKey(key)
        return value == 0
    except Exception:
        return False
    

    
# Перевірка, чи існує ключ реєстру
def key_exists(path):
    try:
        winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_READ)
        return True
    except FileNotFoundError:
        return False

# Створення ключа реєстру
def create_reg_key(path):
    try:
        winreg.CreateKey(winreg.HKEY_CURRENT_USER, path)
    except Exception as e:
        print(f"Error creating registry key: {e}")



if __name__ == '__main__':
    print(is_dark_theme())