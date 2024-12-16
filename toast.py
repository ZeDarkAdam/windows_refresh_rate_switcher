from winotify import Notification, audio


def show_notification(title, message):

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



if __name__ == '__main__':
    show_notification("Test", "Test message")

