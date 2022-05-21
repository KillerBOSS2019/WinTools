from constants import TPClient
from pyvda import AppView, VirtualDesktop, get_virtual_desktops


def current_vd():
    current_d = VirtualDesktop.current()
    TPClient.stateUpdate(f"KillerBOSS.TP.Plugins.virtualdesktop.current_vd", f"[{current_d.number}] {current_d.name}")
    return current_d

def vd_pinn_app(pinned=True):
    if pinned:
        AppView.current()
        AppView.current().pin()
    if not pinned:
        current_window = AppView.current()
        target_desktop = VirtualDesktop(current_vd().number)
        current_window.move(target_desktop)
        
    

def virtual_desktop(target_desktop=None, move=False, pinned=False):
    number_of_active_desktops = len(get_virtual_desktops())
    """Retrieving VD Number from name"""
    target_desktop = int(target_desktop[target_desktop.find("[") + 1: target_desktop.find("]")])

    if target_desktop == "Next":
        if move:
            current_window = AppView.current()
            target_desktop = VirtualDesktop(target_desktop)
            current_window.move(target_desktop) 
            if pinned:
                AppView.current().pin()
        elif not move:
            if target_desktop <= number_of_active_desktops:
                VirtualDesktop(target_desktop).go()
            else:
                print("too many")
        
    elif target_desktop == "Previous":
        if move:
            AppView.current().move(VirtualDesktop(target_desktop)) 
            if pinned:
                AppView.current().pin()
                
        elif not move:
            if target_desktop > 0:
                if pinned:
                    AppView.current().pin()
                VirtualDesktop(target_desktop).go()
            else:
                pass
            
    elif target_desktop != "Previous" or "Next":  
        if move:
            if pinned:
                AppView.current().move(VirtualDesktop(target_desktop)) 
                vd_pinn_app()
                VirtualDesktop(target_desktop).go()
            else:
                AppView.current().move(VirtualDesktop(target_desktop)) 
        elif not move:
            if target_desktop > 0:
                VirtualDesktop(target_desktop).go()
                if pinned:
                    vd_pinn_app()
    current_vd()
    
def rename_vd(name, number=None):
    number_of_active_desktops = len(get_virtual_desktops())
    new_vd = VirtualDesktop(number_of_active_desktops)
    VirtualDesktop.rename(new_vd, name)
    

def create_vd(name=None):
    if name:
        VirtualDesktop.create()
        rename_vd(name)
    else:
        VirtualDesktop.create()
        


def remove_vd(remove, fallbacknum=None):
    if fallbacknum:
        try:
            fall_back =VirtualDesktop(fallbacknum)
            remove_vd = VirtualDesktop(remove) 
            VirtualDesktop.remove(remove_vd, fall_back)
        except ValueError as err:
            return err
    else:
        try:
            remove_vd=VirtualDesktop(remove)
            VirtualDesktop.remove(remove_vd)
        except ValueError as err:
            return err
    current_vd()



def get_vd_name(number):
    name = VirtualDesktop(number).name
    if not name:
        name = f"Desktop {number}"
    return name

old_vd = []
def vd_check():
    global old_vd
    vdlist=[]
    vdlist.append("Next")
    vdlist.append("Previous")
    virtual_desks = get_virtual_desktops()
    for item in virtual_desks:
        if not item.name:
            vdlist.append(f"[{item.number}] Desktop {item.number}")
        else:
            vdlist.append(f"[{item.number}] {item.name}")
    if vdlist != old_vd:
        TPClient.choiceUpdate("KillerBOSS.TP.Plugins.virtualdesktop.actionchoice", vdlist)
        old_vd = vdlist