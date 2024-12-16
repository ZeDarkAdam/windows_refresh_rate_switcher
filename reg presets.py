import json
import winreg

# Дані моніторів
monitors = [
    {"serial": "Dell U2419H", "RefreshRate": 60},
    {"serial": "ASUS VG249Q", "RefreshRate": 144},
]

# Конвертація в JSON
json_data = json.dumps(monitors)

# Шлях до реєстру
registry_path = r"Software\WRRS\Settings"
key_name = "MonitorPresets"

# Запис у реєстр
with winreg.CreateKey(winreg.HKEY_CURRENT_USER, registry_path) as key:
    winreg.SetValueEx(key, key_name, 0, winreg.REG_SZ, json_data)









# Читання значення
with winreg.OpenKey(winreg.HKEY_CURRENT_USER, registry_path) as key:
    json_data = winreg.QueryValueEx(key, key_name)[0]

# Конвертація з JSON
monitors = json.loads(json_data)

# Виведення даних
for monitor in monitors:
    print(f"Name: {monitor['serial']}, Refresh Rate: {monitor['RefreshRate']} Hz")
