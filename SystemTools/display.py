import win32api
import win32con
import pywintypes
from util import PLATFORM_SYSTEM

"""      DISPLAY CHANGE STUFF       """

#DISPLAY_DEVICE.StateFlags
DISPLAY_DEVICE_ACTIVE = 0x1
DISPLAY_DEVICE_MULTI_DRIVER = 0x2
DISPLAY_DEVICE_PRIMARY_DEVICE = 0x4
DISPLAY_DEVICE_MIRRORING_DRIVER = 0x8
DISPLAY_DEVICE_VGA_COMPATIBLE = 0x10
DISPLAY_DEVICE_REMOVABLE = 0x20
DISPLAY_DEVICE_DISCONNECT = 0x2000000
DISPLAY_DEVICE_REMOTE = 0x4000000
DISPLAY_DEVICE_MODESPRUNED = 0x8000000

#EnumDisplaySettingsEx.iModeNum
ENUM_CURRENT_SETTINGS = -1
ENUM_REGISTRY_SETTINGS = -2

#ChangeDisplaySettingsEx.dwflags
CDS_NONE                 = 0x00000000
CDS_UPDATEREGISTRY       = 0x00000001
CDS_TEST                 = 0x00000002
CDS_FULLSCREEN           = 0x00000004
CDS_GLOBAL               = 0x00000008
CDS_SET_PRIMARY          = 0x00000010
CDS_VIDEOPARAMETERS      = 0x00000020
CDS_ENABLE_UNSAFE_MODES  = 0x00000100
CDS_DISABLE_UNSAFE_MODES = 0x00000200
CDS_RESET                = 0x40000000
CDS_RESET_EX             = 0x20000000
CDS_NORESET              = 0x10000000
# primary_device = None
# primary_settings = None

class Display:
  #  primary_device = None
   # primary_settings = None
    #def __init__(self, display_num, display_name, display_orientation):
    #    self.display_num = display_num
    #    self.display_name = display_name
    #    self.display_orientation = display_orientation
    #   #primary_device = None
    #   #primary_settings = None

    def rotate_display(display_num, rotate_choice):
        """ Rotates the display. """
        if PLATFORM_SYSTEM == "Windows":
            display_num = int(display_num.split(":")[0]) - 1
            rotation_values = {
                "180": win32con.DMDO_180,
                "90": win32con.DMDO_270,
                "270": win32con.DMDO_90,
                "0": win32con.DMDO_DEFAULT
            }
            rotation_val = rotation_values.get(rotate_choice)
    
            device = win32api.EnumDisplayDevices(None, display_num)
            dm = win32api.EnumDisplaySettings(device.DeviceName, win32con.ENUM_CURRENT_SETTINGS)
    
            if rotation_val is not None:
                if (dm.DisplayOrientation + rotation_val) % 2 == 1:
                    dm.PelsWidth, dm.PelsHeight = dm.PelsHeight, dm.PelsWidth
                dm.DisplayOrientation = rotation_val
    
            win32api.ChangeDisplaySettingsEx(device.DeviceName, dm)





    def change_primary(monitornum):
        monitornum = monitornum.split(":")[0]
        primary = rf"\\.\DISPLAY{monitornum}"
        primary_settings = None
     #   print("After splitting", monitornum)

        # Find the settings of the new primary display
        i = 0
        while True:
            try:
                device = win32api.EnumDisplayDevices(None, i)
            except pywintypes.error:
                break
            
            if device.DeviceName == primary:
              #  primary_device = device
                primary_settings = win32api.EnumDisplaySettingsEx(device.DeviceName, ENUM_CURRENT_SETTINGS, 0)
                break

            i += 1

        # Update all the positions of the displays relative to the new primary display
        i = 0
        while True:
            try:
                device = win32api.EnumDisplayDevices(None, i)
            except pywintypes.error:
                break
            
            if device.StateFlags & DISPLAY_DEVICE_ACTIVE != 0:
               # print (device.DeviceName)
                settings = win32api.EnumDisplaySettingsEx(device.DeviceName, ENUM_CURRENT_SETTINGS, 0)
               # print (settings.PelsWidth, "x", settings.PelsHeight, "@", settings.Position_x, ",", settings.Position_y)

                settings.Position_x -= primary_settings.Position_x
                settings.Position_y -= primary_settings.Position_y

                # CDS_UPDATEREGISTRY | CDS_NORESET = Don't make the changes yet until we've updated all the displays
                if device.DeviceName == primary:
                    win32api.ChangeDisplaySettingsEx(device.DeviceName, settings, CDS_SET_PRIMARY | CDS_UPDATEREGISTRY | CDS_NORESET)
                else:
                    win32api.ChangeDisplaySettingsEx(device.DeviceName, settings, CDS_UPDATEREGISTRY | CDS_NORESET)

            i += 1

        ## Update the displays with the registry settings
        win32api.ChangeDisplaySettingsEx(None, None)