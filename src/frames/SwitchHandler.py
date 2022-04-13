import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from frames import TemplateExecutor as mainframe
from Util import Util

class SwitchHandler(ttk.LabelFrame):
    def __init__(self, master):
        super().__init__(master)
        self['text'] = '実行処理'

        self.grid(row=0, pady=(30, 20), padx=100, sticky=tk.NSEW)
        self.create_widgets()
        Util.configure_grid(self.master, self, list(range(0,1)))
        
        
    def create_widgets(self):
        exec_comb = ttk.Combobox(self, values=['交通費集計', '勤怠集計'], font=self.master.font, state='readonly')
        exec_comb.current(0)
        exec_comb.pack(ipady=5, ipadx=10, pady=(10, 20), padx=10)
        
        # initialize frames
        self.frames = {}
        self.frames[0] = mainframe.TemplateExecutor(self.master, Util.total_trans_expences)
        self.frames[1] = mainframe.TemplateExecutor(self.master, Util.total_attend_records)
        # Default
        self.change_frame(0)
        
        exec_comb.bind("<<ComboboxSelected>>", self.on_select)  


    def on_select(self, event):      
        match event.widget.get():
            case '交通費集計':
                self.change_frame(0)
            case '勤怠集計':
                self.change_frame(1)
            case _:
                # ここには到達しないような想定
                messagebox.showerror('エラー', '想定外の値です。')
                self.change_frame(0)
                
    
    def change_frame(self, idx):
        frame = self.frames[idx]
        frame.tkraise()
        
        
