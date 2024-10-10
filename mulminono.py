import tkinter as tk
from ctypes import windll, Structure, byref
from ctypes.wintypes import RECT, HWND, DWORD
import webbrowser
import pystray
from  PIL import Image
import threading
import sys
import os

# User Window Data
class AppBarData(Structure):
    _fields_ = [("cbSize", DWORD),
                ("hWnd", HWND),
                ("uCallbackMessage", DWORD),
                ("uEdge", DWORD),
                ("rc", RECT),
                ("lParam", DWORD)]

def get_taskbar_height():
    try:
        appbar_data = AppBarData()
        appbar_data.cbSize = DWORD(36)
        if windll.shell32.SHAppBarMessage(5, byref(appbar_data)) == 0:
            raise WindowsError("Failed fetching Taskbar Data")
        taskbar_height = appbar_data.rc.bottom - appbar_data.rc.top
        if appbar_data.uEdge == 3 or appbar_data.uEdge == 1:
            return taskbar_height
    except Exception as e:
        print(f"Failed fetching Taskbar Height: {e}")
    return 0

# Mulminono
class OverlayApp:
    def __init__(self, root):
        # Variables
        self.screen_width = root.winfo_screenwidth()
        self.screen_height = root.winfo_screenheight()
        self.taskbar_height = get_taskbar_height()
        self.overlays = []
        self.color = "black"
        self.mode = "corners"
        self.size = 75
        self.isOverlayed = True
        self.isExiting = False

        # Window
        ## Title Bar
        self.root = root
        self.root.title("MulmiNoNo")
        self.root.iconbitmap(self.resource_path("mulminono.ico"))

        ## Menu Bar
        self.menu_bar = tk.Menu(root)
        self.root.config(menu=self.menu_bar)

        ### Overlay Menu
        overlay_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Overlay", menu=overlay_menu)
        overlay_menu.add_command(label="Show Overlay (Space)", command=self.show_overlays)
        overlay_menu.add_command(label="Hide Overlay (Space)", command=self.hide_overlays)

        ### Color Menu
        color_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Color", menu=color_menu)
        color_menu.add_command(label="Set Color to Black (B)", command=lambda: self.set_color("black"))
        color_menu.add_command(label="Set Color to White (W)", command=lambda: self.set_color(f'#EEEEEE'))

        ### Draw Mode Menu
        mode_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Draw Mode", menu=mode_menu)
        mode_menu.add_command(label="Draw at Corners (1)", command=lambda: self.set_draw_mode("corners"))
        mode_menu.add_command(label="Draw on Sides (2)", command=lambda: self.set_draw_mode("sides"))
        mode_menu.add_command(label="Draw All (3)", command=lambda: self.set_draw_mode("eight"))

        ### Size Menu
        size_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Size", menu=size_menu)
        size_menu.add_command(label="25% (-)", command=lambda: self.set_size(25))
        size_menu.add_command(label="50% (-)", command=lambda: self.set_size(50))
        size_menu.add_command(label="75% (+)", command=lambda: self.set_size(75))
        size_menu.add_command(label="100% (+)", command=lambda: self.set_size(100))

        ### About Menu
        about_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="About", menu=about_menu)
        about_menu.add_command(label="About", command=lambda: self.open_github())

        # Shortcuts
        #### Space : Overlay
        self.root.bind('<space>', lambda e: self.space_pressed())
        
        #### B/b : Black , W/w : White
        self.root.bind('<b>', lambda e: self.set_color("black"))
        self.root.bind('<B>', lambda e: self.set_color("black"))
        self.root.bind('<w>', lambda e: self.set_color(f'#EEEEEE'))
        self.root.bind('<W>', lambda e: self.set_color(f'#EEEEEE'))
        
        #### 1/2/3 : Draw Mode
        self.root.bind('<Key-1>', lambda e: self.set_draw_mode("corners"))
        self.root.bind('<Key-2>', lambda e: self.set_draw_mode("sides"))
        self.root.bind('<Key-3>', lambda e: self.set_draw_mode("eight"))

        #### +/-/= : Size
        self.root.bind('<plus>', lambda e: self.set_size(self.size + 25))
        self.root.bind('<equal>', lambda e: self.set_size(self.size + 25))
        self.root.bind('<minus>', lambda e: self.set_size(self.size - 25))

        #### Q/q : Exit
        self.root.bind('<q>', lambda e: self.exit_program())
        self.root.bind('<Q>', lambda e: self.exit_program())

        # Tray Mode
        self.create_tray_icon()
        self.root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)

        # Show Overlays
        self.show_overlays()


    # Overlay
    ### Set Overlays to Clickthroughable
    def set_window_clickthrough(self, hwnd):
        try:
            WS_EX_LAYERED = 0x00080000
            WS_EX_TRANSPARENT = 0x00000020
            GWL_EXSTYLE = -20
            style = windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            if windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style | WS_EX_LAYERED | WS_EX_TRANSPARENT) == 0:
                raise WindowsError("Failed Setting Window Style")
        except Exception as e:
            print(f"Erorr on Overlay Window Setting: {e}")

    ### Create Overlays
    def create_overlay(self, x, y):
        width = self.size
        height = self.size
        overlay = tk.Toplevel(self.root)
        overlay.overrideredirect(True)
        overlay.attributes('-topmost', True)
        overlay.attributes('-alpha', 0.9)

        overlay.wm_attributes("-transparentcolor", "white")
        overlay.update_idletasks()
        hwnd = windll.user32.GetParent(overlay.winfo_id())
        self.set_window_clickthrough(hwnd)

        canvas = tk.Canvas(overlay, width=self.size, height=self.size, highlightthickness=0, bg="white")
        canvas.pack()

        canvas.create_rectangle(0, 0, width, height, outline="black", fill=self.color, width=0)

        self.overlays.append(overlay)
        overlay.geometry(f"{width}x{height}+{x}+{y}")

    ### Show Overlays
    def show_overlays(self):
        self.isOverlayed = True
        if self.overlays:
            return

        width, height = self.size, self.size
        margin = 15
        positions = []

        if self.mode == "corners":
            positions = [
                (margin, margin),  # Left Top
                (self.screen_width - width - margin, margin),  # Right Top
                (margin, self.screen_height - height - self.taskbar_height - margin),  # Left Bottom
                (self.screen_width - width - margin, self.screen_height - height - self.taskbar_height- margin),  # Right Bottom
            ]

        elif self.mode == "sides":
            positions = [
                (self.screen_width // 2 - width // 2 - margin, margin),  #23 Center Top
                (self.screen_width // 2 - width // 2 - margin, self.screen_height - height - self.taskbar_height - margin),  # Center Bottom
                (margin, self.screen_height // 2 - height // 2 - self.taskbar_height // 2 - margin // 2),  # Left Center
                (self.screen_width - width - margin, self.screen_height // 2 - height // 2 - self.taskbar_height // 2 - margin // 2)  # Right Center
            ]

        elif self.mode == "eight":
            positions = [
                (self.screen_width // 2 - width // 2 - margin, margin),  # Top Center
                (self.screen_width // 2 - width // 2 - margin, self.screen_height - height - self.taskbar_height - margin),  # Bottom Center
                (margin, self.screen_height // 2 - height // 2 - self.taskbar_height // 2 - margin // 2),  # Left Center
                (self.screen_width - width - margin, self.screen_height // 2 - height // 2 - self.taskbar_height // 2 - margin // 2),  # Right Center
                (margin, margin),  # Left Top
                (self.screen_width - width - margin, margin),  # Right Top
                (margin, self.screen_height - height - self.taskbar_height - margin),  # Left Bottom
                (self.screen_width - width - margin, self.screen_height - height - self.taskbar_height- margin),  # Right Bottom
            ]            

        for x, y in positions:
            self.create_overlay(x, y)

    ### Hide Overlays
    def hide_overlays(self):
        self.isOverlayed = False
        for overlay in self.overlays:
            overlay.destroy()
        self.overlays = []

    ## Overlay Options
    ### Color
    def set_color(self, color):
        self.color = color
        self.outline = color
        self.hide_overlays()
        self.show_overlays()

    ### Draw Mode
    def set_draw_mode(self, mode):
        self.mode = mode
        self.hide_overlays()
        self.show_overlays()

    ### Size
    def set_size(self, size):
        if (size > 100):
            size = 100
        elif size < 25:
            size = 25
        self.size = size
        self.hide_overlays()
        self.show_overlays()

    ### Show / Hide
    def space_pressed(self):
        if (self.isOverlayed):
            self.hide_overlays()
        else:
            self.show_overlays()

    # About
    ### Open Github README
    def open_github(self):
        try:
            webbrowser.open("https://github.com/TaeAhnK/MulmiNoNo")
        except Exception as e:
            print(f"Failed Opening README Page: {e}")

    # Program
    ### Exit Program
    def exit_program(self):
        if self.icon:
            self.icon.stop()
        self.cleanup()

    ### Cleanup and Quit Program
    def cleanup(self):
        if not self.isExiting:
            self.isExiting = True
            self.hide_overlays()
            self.root.quit()
            self.root.destroy()

    ### Get Icon Path
    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)
    
    # Tray
    ### Create Tray Icon and Window
    def create_tray_icon(self):
        self.icon = pystray.Icon("overlay_app")
        self.icon.icon = Image.open(self.resource_path("mulminono.ico"))
        self.icon.menu = pystray.Menu(
            pystray.MenuItem("Open Window", self.restore_window),
            pystray.MenuItem("Quit", self.exit_program)
        )

    ### Minimize
    def minimize_to_tray(self):
        self.root.withdraw()
        if self.icon is None:
            self.create_tray_icon()
        threading.Thread(target=self.icon.run, daemon=True).start()

    ### Restore back to Window Mode
    def restore_window(self, icon, item):
        self.root.after(0, self.root.deiconify)
        if self.icon:
            self.icon.stop()
            self.icon = None
    

if __name__ == "__main__":
    root = tk.Tk()
    app = OverlayApp(root)
    root.geometry("300x0")
    try:
        root.mainloop()
    finally:
        app.cleanup()
