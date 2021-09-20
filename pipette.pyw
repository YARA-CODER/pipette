import colorsys
import ctypes
import os
import tkinter as tk
import tkinter.ttk as ttk
from collections import namedtuple
from tkinter import colorchooser

import win32con
import win32gui
from PIL import Image, ImageDraw, ImageTk
from pynput.mouse import Button, Listener
from pywintypes import error as pywintypes_error

# TODO List:
# --------------------------------------------------------------------------------
# Change: tk.Entry -> tk.Spinbox (x3) for RGB, CMYK etc.
# --------------------------------------------------------------------------------


class App(tk.Tk):
    __PROGRAM_PATH = os.getcwd()

    def __init__(self):
        super().__init__()
        self.iconbitmap(default=os.path.join('img', 'pipette_icon.ico'))
        self.cursor = os.path.join('img', 'cursor.cur')
        self.frmUpper = tk.LabelFrame(self, borderwidth=0)
        self.frmColor = tk.LabelFrame(self.frmUpper, borderwidth=2, height=170, text='Color and Pipette')
        self.frmButtons = tk.Frame(self.frmColor)
        self.frmList = tk.LabelFrame(self.frmUpper, borderwidth=2, width=200, text='Color List')
        self.frmOutput = tk.LabelFrame(self, borderwidth=2, height=200, width=340, text='Representation in color modes')
        self.frmStatusbar = tk.Frame(self)

        self.color_sample = tk.Frame(self.frmColor, width=160, height=120, bg="white", borderwidth=2, relief=tk.RIDGE)

        self.pipette_icon = tk.PhotoImage(file=os.path.join(self.__PROGRAM_PATH, 'img', 'pipette_icon_24px.png'))
        self.btn_pipette = tk.Button(self.frmButtons, text="      Pipette", image=self.pipette_icon, compound=tk.LEFT, anchor=tk.W, command=self.startCursorTracking)

        self.picker_icon = tk.PhotoImage(file=os.path.join(self.__PROGRAM_PATH, 'img', 'rgb_wheel_24px.png'))
        self.btn_picker = tk.Button(self.frmButtons, text="CC", image=self.picker_icon, command=self.chooseColor)

        self.frmUpper.pack(fill=tk.X, side=tk.TOP)
        self.frmColor.pack(fill=tk.X, side=tk.LEFT, expand=True, padx=5)
        self.frmList.pack(fill=tk.Y, side=tk.RIGHT, padx=(0, 5))
        self.frmOutput.pack(fill=tk.BOTH, side=tk.TOP, expand=False, padx=5, pady=5)
        self.frmStatusbar.pack(fill=tk.X, side=tk.BOTTOM, expand=False)
        self.color_sample.pack(fill=tk.BOTH, side=tk.TOP, expand=True, padx=5, pady=5)
        self.frmButtons.pack(fill=tk.X, side=tk.TOP, expand=True, padx=0, pady=0)
        self.btn_pipette.pack(fill=tk.X, side=tk.LEFT, expand=True, padx=5, pady=5)
        self.btn_picker.pack(side=tk.RIGHT, ipadx=2, ipady=2, padx=(0, 5))

        self.ColorList = ttk.Treeview(self.frmList, columns=('RGB_CODE'), show='tree', height=8)
        self.ColorList.pack(fill=tk.BOTH, side=tk.TOP, expand=True, padx=5, pady=5)
        self.ColorList.column('#0', width=48, anchor=tk.W)
        self.ColorList.column('RGB_CODE', width=100, anchor=tk.W)

        self.frmR1 = tk.Frame(self.frmOutput)
        self.frmR2 = tk.Frame(self.frmOutput)
        self.frmR3 = tk.Frame(self.frmOutput)
        self.frmR4 = tk.Frame(self.frmOutput)
        self.frmR5 = tk.Frame(self.frmOutput)

        rows = [self.frmR1, self.frmR2, self.frmR3, self.frmR4, self.frmR5]
        [r.pack(fill=tk.X, side=tk.TOP, padx=20, pady=5) for r in rows]

        self.val_hex = tk.StringVar()
        self.label1 = tk.Label(self.frmR1, text='HEX   ').pack(side=tk.LEFT)
        self.tbx_val_hex = tk.Entry(self.frmR1, textvariable=self.val_hex).pack(fill=tk.X, side=tk.LEFT, expand=True)

        self.val_rgb = tk.StringVar()
        self.label2 = tk.Label(self.frmR2, text='RGB   ').pack(side=tk.LEFT)
        self.tbx_val_rgb = tk.Entry(self.frmR2, textvariable=self.val_rgb).pack(fill=tk.X, side=tk.LEFT, expand=True)

        self.val_cmyk = tk.StringVar()
        self.label3 = tk.Label(self.frmR3, text='CMYK').pack(side=tk.LEFT)
        self.tbx_val_cmyk = tk.Entry(self.frmR3, textvariable=self.val_cmyk).pack(fill=tk.X, side=tk.LEFT, expand=True)

        self.val_hsv = tk.StringVar()
        self.label4 = tk.Label(self.frmR4, text='HSV   ').pack(side=tk.LEFT)
        self.tbx_val_hsv = tk.Entry(self.frmR4, textvariable=self.val_hsv).pack(fill=tk.X, side=tk.LEFT, expand=True)

        self.val_hsl = tk.StringVar()
        self.label5 = tk.Label(self.frmR5, text='HSL   ').pack(side=tk.LEFT)
        self.tbx_val_hsl = tk.Entry(self.frmR5, textvariable=self.val_hsl).pack(fill=tk.X, side=tk.LEFT, expand=True)

        self.debug1_text = tk.StringVar()
        self.debug1 = tk.Label(self.frmStatusbar, bd=1, relief=tk.SUNKEN, width=15, textvariable=self.debug1_text, anchor=tk.W)
        self.debug1.pack(fill=tk.X, side=tk.LEFT)

        self.debug2_text = tk.StringVar()
        self.debug2 = tk.Label(self.frmStatusbar, bd=1, relief=tk.SUNKEN, width=15, textvariable=self.debug2_text)
        self.debug2.pack(fill=tk.X, side=tk.LEFT)

        self.debug3_text = tk.StringVar()
        self.debug3 = tk.Label(self.frmStatusbar, bd=1, relief=tk.SUNKEN, width=15, textvariable=self.debug3_text)
        self.debug3.pack(fill=tk.X, side=tk.LEFT)

        self.geometry('360x435')
        self.resizable(False, False)
        # set window to always on top
        self.attributes('-topmost', True)
        self.update_idletasks()

        self.center()
        self.debug1_text.set('Ready')

        self.ColorList.bind('<<TreeviewSelect>>', self._treeSelected)

        self.thumbnails = []

    def addToList(self, rgb):
        thumbnail = Image.new('RGB', (16, 16))
        ImageDraw.Draw(thumbnail).rectangle([(0, 0), (15, 15)], fill=rgb, outline=(0, 0, 0), width=1)
        self.thumbnails.append(ImageTk.PhotoImage(thumbnail))
        self.ColorList.insert('', 'end', image=self.thumbnails[-1], value=(f'({rgb[0]:03d},{rgb[1]:03d},{rgb[2]:03d})'))

    def center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry('{}x{}+{}+{}'.format(w, h, x, y))

    def rgb2hex(self, rgb):
        return '#%02x%02x%02x' % rgb

    def rgb2cmyk(self, rgb):
        rgb_scale = 255
        cmyk_scale = 100
        r, g, b = rgb

        if (r == 0) and (g == 0) and (b == 0):
            # black
            return 0, 0, 0, cmyk_scale

        # rgb [0,255] -> cmy [0,1]
        c = 1 - r / float(rgb_scale)
        m = 1 - g / float(rgb_scale)
        y = 1 - b / float(rgb_scale)

        # extract out k [0,1]
        min_cmy = min(c, m, y)
        c = (c - min_cmy)
        y = (y - min_cmy)
        m = (m - min_cmy)
        k = min_cmy

        CMYK = c * cmyk_scale, m * cmyk_scale, y * cmyk_scale, k * cmyk_scale
        CMYK = tuple([round(v, 2) for v in CMYK])
        # rescale to the range [0,cmyk_scale]
        return CMYK

    def rgb2hsv(self, rgb, target='hsv'):
        rgb = [c / 255 for c in rgb]
        r, g, b = rgb
        if target == 'hsv':
            hsv = colorsys.rgb_to_hsv(r, g, b)
        elif target == 'hls':
            hsv = colorsys.rgb_to_hls(r, g, b)
        hsv = [round(c, 2) for c in hsv]
        return hsv

    def setCursor(self):
        cursors = [32651, 32648, 32646, 32516, 32514, 32512, 32513, 32649, 32642, 32643, 32645, 32644, 32515]

        self.hcursors_saved = []
        for cur in cursors:

            new_cursor = win32gui.LoadImage(0, self.cursor, win32con.IMAGE_CURSOR, 0, 0, win32con.LR_LOADFROMFILE)
            hcursor = win32gui.LoadImage(0, cur, win32con.IMAGE_CURSOR, 0, 0, win32con.LR_SHARED)
            hcursor_image = ctypes.windll.user32.CopyImage(hcursor, win32con.IMAGE_CURSOR, 0, 0, win32con.LR_COPYFROMRESOURCE)

            self.hcursors_saved.append((hcursor_image, cur))

            ctypes.windll.user32.SetSystemCursor(new_cursor, cur)

    def restoreCursors(self):
        if self.hcursors_saved:
            for cur in self.hcursors_saved:
                ctypes.windll.user32.SetSystemCursor(cur[0], cur[1])

    def pixelColor(self, i_x, i_y):
        i_desktop_window_id = win32gui.GetDesktopWindow()
        i_desktop_window_dc = win32gui.GetWindowDC(i_desktop_window_id)
        long_colour = win32gui.GetPixel(i_desktop_window_dc, i_x, i_y)
        i_colour = int(long_colour)
        win32gui.ReleaseDC(i_desktop_window_id, i_desktop_window_dc)

        return (i_colour & 0xff), ((i_colour >> 8) & 0xff), ((i_colour >> 16) & 0xff)

    def on_move(self, x, y):
        try:
            self.color_sample.configure(bg=self.rgb2hex(self.pixelColor(x, y)))
        except pywintypes_error:
            pass
        self.debug2_text.set((x, y))

    def on_click(self, x, y, button, pressed):
        if button == Button.left:
            self.restoreCursors()
            self.updateForm(self.pixelColor(x, y))
            self.addToList(self.pixelColor(x, y))
            self.debug1_text.set('Ready')
            self.debug2_text.set('')
            self.listener.stop()

    def updateForm(self, rgb):
        self.color_sample.configure(bg=self.rgb2hex(rgb))
        self.val_hex.set(self.rgb2hex(rgb))
        self.val_rgb.set('rgb' + repr(rgb))
        self.val_cmyk.set('cmyk' + repr(self.rgb2cmyk(rgb)))

        self.val_hsv.set(self.rgb2hsv(rgb))
        self.val_hsl.set(self.rgb2hsv(rgb, target='hls'))

    def win32_event_filter(self, msg, data):
        # msg: 513,514-LMB, 516,517-RMB, 519-MiddleMB
        self.listener._suppress = True if msg in (513, 514, 516, 517) else False
        return True

    def startCursorTracking(self):
        self.setCursor()
        self.debug1_text.set('Picking color...')
        self.listener = Listener(on_move=self.on_move, on_click=self.on_click, win32_event_filter=self.win32_event_filter)
        self.listener.start()

    def chooseColor(self):
        self.debug1_text.set('Choose color')
        init_color = self.color_sample.config()['background'][4]
        color_pick = colorchooser.askcolor(initialcolor=init_color)
        try:
            color_pick = tuple([int(c) for c in color_pick[0]])
            self.debug2_text.set(color_pick)
            self.updateForm(color_pick)
            self.addToList(color_pick)
        except TypeError:
            pass
        finally:
            self.debug1_text.set('Ready')

    def _treeSelected(self, event):
        # try using namedtuple here ;) or at AppClass level
        rgb = event.widget.item(event.widget.focus())['values'][0]
        rgb = (int(rgb[1:4]), int(rgb[5:8]), int(rgb[9:12]))
        self.updateForm(rgb)


if __name__ == '__main__':
    app = App()
    app.title("Filip's Pipette")
    app.mainloop()
