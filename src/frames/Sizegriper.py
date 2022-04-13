import tkinter as tk
from tkinter import ttk
from Util import Util

class Sizegriper(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)

        self.grid(row=2, sticky=tk.SE)
        self.create_widgets()
        Util.configure_grid(self.master, self, list(range(0,1)))
        
    def create_widgets(self):       
        sizegrip = ttk.Sizegrip(self)
        sizegrip.grid(row=0, sticky=tk.SE)
       
