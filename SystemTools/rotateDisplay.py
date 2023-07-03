import win32api
import win32con


def rotate_display(display_num, rotate_choice):
    display_num = display_num.split(":")[0]
    display_num = int(display_num)
    display_num = display_num -1
    rotation_val=""

    if (rotate_choice != None):
        if (rotate_choice == "180"):
            rotation_val=win32con.DMDO_180
        elif(rotate_choice == "90"):
            rotation_val=win32con.DMDO_270
        elif (rotate_choice == "270"):   
            rotation_val=win32con.DMDO_90
        elif (rotate_choice == "0"):
            rotation_val=win32con.DMDO_DEFAULT

    device = win32api.EnumDisplayDevices(None,display_num)
    dm = win32api.EnumDisplaySettings(device.DeviceName,win32con.ENUM_CURRENT_SETTINGS)
    if((dm.DisplayOrientation + rotation_val)%2==1):
        dm.PelsWidth, dm.PelsHeight = dm.PelsHeight, dm.PelsWidth   
    dm.DisplayOrientation = rotation_val

    win32api.ChangeDisplaySettingsEx(device.DeviceName,dm)