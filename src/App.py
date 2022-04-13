import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk
from frames import SwitchHandler as topframe
from frames import Sizegriper as btmframe
from Util import Util
import ctypes
import os
import sys

class DataAggregationTool(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('データ集計ツール')
        self.define_style()
        self.resizable(1, 1)
        
    def setup(self):
        self.iconbitmap(default=Util.get_modifiedpath(os.path.join('resources', 'icon.ico')))
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(True)
        except:
            pass


    def define_style(self):
        self.style = ttk.Style(self)
        #self.style.theme_use('vista')
        self.font = ('Meiryo UI', 14)
        self.style.configure('TLabel', font=self.font)
        self.style.configure('TButton', font=self.font)
        self.style.configure('Exec.TButton', font=('Meiryo UI', 16, 'bold'), foreground='#ff6347')
        self.option_add("*TCombobox*Listbox.Font", self.font)
        #self.style.configure('TCombobox.Font', font=self.font)
        #self.style.configure('TEntry.Text', font=self.font)
        #self.style.configure('TFrame', background='green')
        self.style.configure('TLabelframe.Label', font=self.font)

        
if __name__ == '__main__':
    app = DataAggregationTool()
    topframe.SwitchHandler(app)
    btmframe.Sizegriper(app)
    app.setup()
    app.mainloop()
    