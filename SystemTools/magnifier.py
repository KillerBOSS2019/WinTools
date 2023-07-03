import re
import winreg
from util import runWindowsCMD as runCommandLine


class WindowsRegistry:
    """Class WindowsRegistry is using for easy manipulating Windows registry.
    Methods
    -------
    query_value(full_path : str)
        Check value for existing.
    get_value(full_path : str)
        Get value's data.
    set_value(full_path : str, value : str, value_type='REG_SZ' : str)
        Create a new value with data or set data to an existing value.
    delete_value(full_path : str)
        Delete an existing value.
    query_key(full_path : str)
        Check key for existing.
    delete_key(full_path : str)
        Delete an existing key(only without subkeys).
    Examples:
        WindowsRegistry.set_value('HKCU/Software/Microsoft/Windows/CurrentVersion/Run', 'Program', r'"c:\Dir1\program.exe"')
        WindowsRegistry.delete_value('HKEY_CURRENT_USER/Software/Microsoft/Windows/CurrentVersion/Run/Program')
    """
    @staticmethod
    def __parse_data(full_path):
        full_path = re.sub(r'/', r'\\', full_path)
        hive = re.sub(r'\\.*$', '', full_path)
        if not hive:
            raise ValueError('Invalid \'full_path\' param.')
        if len(hive) <= 4:
            if hive == 'HKLM':
                hive = 'HKEY_LOCAL_MACHINE'
            elif hive == 'HKCU':
                hive = 'HKEY_CURRENT_USER'
            elif hive == 'HKCR':
                hive = 'HKEY_CLASSES_ROOT'
            elif hive == 'HKU':
                hive = 'HKEY_USERS'
        reg_key = re.sub(r'^[A-Z_]*\\', '', full_path)
        reg_key = re.sub(r'\\[^\\]+$', '', reg_key)
        reg_value = re.sub(r'^.*\\', '', full_path)

        return hive, reg_key, reg_value

    @staticmethod
    def query_value(full_path):
        value_list = WindowsRegistry.__parse_data(full_path)
        try:
            opened_key = winreg.OpenKey(getattr(winreg, value_list[0]), value_list[1], 0, winreg.KEY_READ)
            winreg.QueryValueEx(opened_key, value_list[2])
            winreg.CloseKey(opened_key)
            return True
        except WindowsError:
            return False

    @staticmethod
    def get_value(full_path):
        value_list = WindowsRegistry.__parse_data(full_path)
        try:
            opened_key = winreg.OpenKey(getattr(winreg, value_list[0]), value_list[1], 0, winreg.KEY_READ)
            value_of_value, value_type = winreg.QueryValueEx(opened_key, value_list[2])
            winreg.CloseKey(opened_key)
            return value_of_value
        except WindowsError:
            return None

    @staticmethod
    def set_value(full_path, value, value_type='REG_SZ'):
        value_list = WindowsRegistry.__parse_data(full_path)
        try:
            winreg.CreateKey(getattr(winreg, value_list[0]), value_list[1])
            opened_key = winreg.OpenKey(getattr(winreg, value_list[0]), value_list[1], 0, winreg.KEY_WRITE)
            winreg.SetValueEx(opened_key, value_list[2], 0, getattr(winreg, value_type), value)
            winreg.CloseKey(opened_key)
            return True
        except WindowsError:
            return False
        
    @staticmethod
    def query_key(full_path):
        value_list = WindowsRegistry.__parse_data(full_path)
        try:
            opened_key = winreg.OpenKey(getattr(winreg, value_list[0]), value_list[1] + r'\\' + value_list[2], 0, winreg.KEY_READ)
            winreg.CloseKey(opened_key)
            return True
        except WindowsError:
            return False

magpath = r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\\"
is_on=WindowsRegistry.get_value(magpath+"\RunningState")







def magnifier_control(action, amount=None):
    match action:
        case "Zoom In":
            mag_level(amount)
        case "Zoom Out":
            mag_level(amount)
        case "Dock":
            mag_mode(1)
        case "Full Screen":
            mag_mode(2)
        case "Lens":
            mag_mode(3)
        case "Invert Colors":
            mag_invert()
            # pyautogui.hotkey('ctrl', 'alt', 'i')
        case "Start":
            mag_on()
        case "Exit":
            try:
              runCommandLine("""wmic process where "name='magnify.exe'" delete""")
            except:
                pass


def mag_on():
    """Check if Magnify is on
    - If magnify is off, this turns it on. 
       Always returns True
    """
    is_on=WindowsRegistry.get_value(magpath+"\RunningState")
    print("magifer is on ? ", is_on )
    if is_on:
        return True
    if not is_on:
        runCommandLine('magnify')
        return True
    
    
def mag_mode(mode=3):
    """ 
    -   1 = Docked
    -   2 = Full Screen
    -   3 = Lens
    """

    exists = WindowsRegistry.query_value(magpath + "MagnificationMode")
    if exists:
        WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\MagnificationMode", mode, value_type='REG_DWORD')



def mag_invert():
    """ Magnifier Invert
    - Toggles On / Off
    """
    exists = WindowsRegistry.query_value(magpath + "Invert")
    if exists:
        current = WindowsRegistry.get_value(magpath + "Invert")
        if current == 1:
            WindowsRegistry.set_value(r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\Invert", 0, value_type='REG_DWORD')
        if current ==0:
            WindowsRegistry.set_value(r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\Invert", 1, value_type='REG_DWORD')

            
def mag_increments(amount):
    """Setting Zoom Increments - MAX 400%, MIN 5%"""
    exists = WindowsRegistry.query_value(magpath + "ZoomIncrement")
    if exists:
        if amount:
            WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\ZoomIncrement", amount, value_type='REG_DWORD')
    
def mag_level(amount, onhold=False):
    """Set Zoom Levels 
    - 1600 is MAX (overriden to 1000)
    - Amount = Zoom Set
    - On Hold Amount = Increment Zoom
    """
    themin=20
    themax=1000
    current = WindowsRegistry.get_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\Magnification")
    exists = WindowsRegistry.query_value(magpath + "Magnification")
    if not onhold:
         if exists:
             if amount:
                    if amount >themax:
                        mag_on()
                        WindowsRegistry.set_value(r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\Magnification", themax, value_type='REG_DWORD')
                    else:
                         mag_on()
                         WindowsRegistry.set_value(r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\Magnification", amount, value_type='REG_DWORD')
    if onhold:
        if exists:
            if current < themin:
                 print("less than 20")
                 WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\Magnification", 20, value_type='REG_DWORD')
            elif current > themax:
                try:
                    WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\Magnification", 1600, value_type='REG_DWORD')
                except:
                    pass
            else:
                try:
                    WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\Magnification", current + amount, value_type='REG_DWORD')
                except:
                    pass
            
def text_smoothing(switch):
    """Text Smoothing
    - Switch = On / Off
    """
    exists = WindowsRegistry.query_value(magpath + "UseBitmapSmoothing")
    if exists:
        if switch =="On":
            WindowsRegistry.set_value(r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\UseBitmapSmoothing", 1, value_type='REG_DWORD')
        elif switch == "Off":
            WindowsRegistry.set_value(r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\UseBitmapSmoothing", 0, value_type='REG_DWORD')


def magnifer_dimensions(x=None, y=None, onhold=False):
    """Can set a MIN/MAX Value to avoid this
    -  onhold = int value increment
    """
    exists = WindowsRegistry.query_value(magpath + "LensWidth")
    themax=95
    themin=20
    if not onhold:
        if x:
            if exists:
                WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensWidth", x, value_type='REG_DWORD')
        if y:
            if exists:
                WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensHeight", y, value_type='REG_DWORD')
                
    if onhold:
        if exists:
            if x:
                 current = WindowsRegistry.get_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensWidth")
                 if current < themin:
                     print("less than 20")
                     WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensWidth", 20, value_type='REG_DWORD')
                 elif current > themax:
                     WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensWidth", 95, value_type='REG_DWORD')
                 else:
                     WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensWidth", current + onhold, value_type='REG_DWORD')
            if y:
                 current = WindowsRegistry.get_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensHeight")
                 if current < themin:
                     print("less than 20")
                     WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensHeight", 20, value_type='REG_DWORD')
                 elif current > themax:
                     WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensHeight", 95, value_type='REG_DWORD')
                 else:
                     WindowsRegistry.set_value("HKEY_CURRENT_USER\SOFTWARE\Microsoft\ScreenMagnifier\LensHeight", current + onhold, value_type='REG_DWORD')