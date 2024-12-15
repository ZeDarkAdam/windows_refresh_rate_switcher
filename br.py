import screen_brightness_control as sbc


# get the brightness
brightness = sbc.get_brightness()
print(brightness)

# get the brightness for the primary monitor
primary = sbc.get_brightness(display=0)
print(primary)

# set the brightness to 100%
# sbc.set_brightness(100)

# set the brightness to 100% for the primary monitor
# sbc.get_brightness(100, display=0)


# show the current brightness for each detected monitor
# for monitor in sbc.list_monitors():
#     print(monitor, ':', sbc.get_brightness(display=monitor), '%')




monitors_info = sbc.list_monitors_info()
print(monitors_info)
for index, monitor in enumerate(monitors_info):

    if monitor["manufacturer"] is None:
        print(f"DISPLAY{index + 1}")
    else:
        print(f"{monitor["manufacturer"]} ({index + 1})")

    print(sbc.get_brightness(display=monitor["name"]))



# test2 = sbc.list_monitors_info()
# print(test2[0])

