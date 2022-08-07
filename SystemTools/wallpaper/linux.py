from gi.repository import Gio

SCHEMA = 'org.gnome.desktop.background'
KEY = 'picture-uri'

def change_background(filename):
    gsettings = Gio.Settings.new(SCHEMA)
    print(gsettings.get_string(KEY))
    print(gsettings.set_string(KEY, "file://" + filename))
    gsettings.apply()
    print(gsettings.get_string(KEY))