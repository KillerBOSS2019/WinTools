import os
from appscript import app, mactypes
from subprocess import call


image_alpha = True


def set_wallpaper(image_file_with_path):

    filepath = os.path.abspath(image_file_with_path)

    try:
        app('Finder').desktop_picture.set(mactypes.File(filepath))
        return True

    except Exception:
        return False