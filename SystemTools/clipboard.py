## Clipboard
from io import BytesIO
from util import PLATFORM_SYSTEM
import subprocess

match PLATFORM_SYSTEM:
    case "Windows":
        import win32clipboard
    case "Linux":
        pass
    case "Darwin":
        pass
    


class ClipBoard:
    def copy_image_to_clipboard(image):
        if PLATFORM_SYSTEM == "Windows":
            bio = BytesIO()
            image.save(bio, 'BMP')
            data = bio.getvalue()[14:]  # removing some headers
            bio.close()
            ClipBoard.send_to_clipboard(win32clipboard.CF_DIB, data)

        # This currently works with Fedora 36 and saves image to clipboard
        if PLATFORM_SYSTEM == "Linux":
            memory = BytesIO()
            image.save(memory, format="png")
            output = subprocess.Popen(("xclip", "-selection", "clipboard", "-t", "image/png", "-i"),
                                      stdin=subprocess.PIPE)
            # write image to stdin
            output.stdin.write(memory.getvalue())
            output.stdin.close()

            # BACKUP OPTIONS
            # Option # 1
            # os.system(f"xclip -selection clipboard -t image/png -i {path + '/image.png'}")
            # os.system("xclip -selection clipboard -t image/png -i temp_file.png")

            # Option #2  - https://stackoverflow.com/questions/56618983/how-do-i-copy-a-pil-picture-to-clipboard
            # might be able to use module called klemboard ??

        # This needs tested/worked on...
        if PLATFORM_SYSTEM == "Darwin":
            # Option #1
            # os.system(f"pbcopy < {path + '/image.png'}")

            # Option #2

            subprocess.run(
                ["osascript", "-e", 'set the clipboard to (read (POSIX file "image.jpg") as JPEG picture)'])


    def send_to_clipboard(clip_type, data):
        match PLATFORM_SYSTEM:
            case "Windows":
                if clip_type == "text":
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardText(data)
                    win32clipboard.CloseClipboard()
                else:
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardData(clip_type, data)
                    win32clipboard.CloseClipboard()

            case "Linux":
                try:
                    subprocess.run(['xclip', '-selection', 'clipboard'], input=data.encode(), check=True)
                except subprocess.CalledProcessError as err:
                    print("Error:", err)
                    return err
                
            case "Darwin":
                try:
                    subprocess.run(['pbcopy'], input=data.encode(), check=True)
                except subprocess.CalledProcessError as err:
                    print("Error:", err)
                    return err
                
            case _:
                print("Unsupported operating system.")
                return None
            
            
    def get_clipboard_data():
        match PLATFORM_SYSTEM:
            case "Windows":
                try:
                    win32clipboard.OpenClipboard()
                    data = win32clipboard.GetClipboardData()
                    win32clipboard.CloseClipboard()
                    return data
                except TypeError as err:
                    return "Invalid Clipboard Data"

            case "Linux":
                command = ["xclip", "-selection", "clipboard", "-o"]

            case "Darwin":
                command = ["pbpaste"]
            case _:
                print("Unsupported operating system.")
                return None
            
        try:
            process = subprocess.Popen(command, stdout=subprocess.PIPE)
            data, _ = process.communicate()
            return data.decode().strip()
        
        except subprocess.CalledProcessError as err:
            print("Error:", err)
            return "Invalid Clipboard Data"

