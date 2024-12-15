from winotify import Notification, audio


icon_path = "C:\Projects\GitHub\windows_refresh_rate_switcher\icons\icon_color.ico"

toast = Notification(
    app_id = 'Winotify',
    title = 'Hello, World!',
    msg = 'This is a notification from Winotify!',
    duration = "short",
    icon = icon_path,
)

toast.set_audio(audio.Default, loop=False)

toast.show()