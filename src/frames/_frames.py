import os
import tkinter as tk
from tkinter import (
    ttk,
    messagebox,
    filedialog,
)

import util
import core
from config import (
    Constants as const,
    ConfigLoader,
    ConfigWriter,
)

class Sizegriper(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)

        self.grid(row=2, sticky=tk.SE)
        self.__create_widgets()
        util.configure_grid(self.master, self, list(range(0,1)))
        
    def __create_widgets(self):       
        sizegrip = ttk.Sizegrip(self)
        sizegrip.grid(row=0, sticky=tk.SE)
       

class SwitchHandler(ttk.LabelFrame):
    def __init__(self, master):
        super().__init__(master)
        self.config = ConfigLoader()
        self['text'] = self.config.get_handler_heading()
        self.combo_values = self.config.get_handler_combo_values()
        self.grid(row=0, pady=(30, 20), padx=100, sticky=tk.NSEW)
        self.__create_widgets()
        util.configure_grid(self.master, self, list(range(0,1)))
               
    def __create_widgets(self):
        exec_comb = ttk.Combobox(self, values=self.combo_values,
                                font=self.master.font, state='readonly')
        exec_comb.current(0)
        exec_comb.pack(ipady=5, ipadx=10, pady=(10, 20), padx=10)
        
        # initialize frames
        self.frames = {}
        self.frames[0] = TemplateMainFrame(self.master, core.write_trans_expenses)
        self.frames[1] = TemplateMainFrame(self.master, core.write_business_expenses)
        # Default
        self.__change_frame(0)
        
        exec_comb.bind("<<ComboboxSelected>>", self.__on_select)  

    def __on_select(self, event):      
        match event.widget.get():
            case self.combo_values(0):
                self.__change_frame(0)
            case self.combo_values(1):
                self.__change_frame(1)
            case _:
                # never reach here
                messagebox.showerror(const.ERROR, const.ERR001)
                self.__change_frame(0)
                   
    def __change_frame(self, idx):
        frame = self.frames[idx]
        frame.tkraise()


class TemplateMainFrame(ttk.Frame):
    def __init__(self, master, exec_method):
        super().__init__(master)
        self.config = ConfigLoader()
        self.exec_method = exec_method

        self.grid(row=1, padx=10, sticky=tk.NSEW)
        self.__create_widgets()
        util.configure_grid(self.master, self, list(range(0,3)))
        
    def __create_widgets(self):       
        self.in_dir_path_var = tk.StringVar()
        self.in_dir_path_var.set(self.config.get_in_dir())
        in_dir_label = ttk.Label(self, text=self.config.get_main_heading1(), anchor='e', width=14)
        in_dir_box = ttk.Entry(self, textvariable=self.in_dir_path_var, width=30, font=self.master.font)
        in_dir_btn = ttk.Button(self, text=const.REFER, command=lambda: self.__select_folder(self.in_dir_path_var))
        in_dir_label.grid(row=0, column=0, pady=10)
        in_dir_box.grid(row=0, column=1, sticky=tk.EW, padx=5)
        in_dir_btn.grid(row=0, column=2, padx=5)
        
        self.out_dir_path_var = tk.StringVar()
        self.out_dir_path_var.set(self.config.get_out_dir())
        out_dir_label = ttk.Label(self, text=self.config.get_main_heading2(), anchor='e', width=14)
        out_dir_box = ttk.Entry(self, textvariable=self.out_dir_path_var, width=30, font=self.master.font)
        out_dir_btn = ttk.Button(self, text=const.REFER, command=lambda: self.__select_folder(self.out_dir_path_var))
        out_dir_label.grid(row=1, column=0, pady=10)
        out_dir_box.grid(row=1, column=1, sticky=tk.EW, padx=5)
        out_dir_btn.grid(row=1, column=2, padx=5)
        
        self.out_file_name_var = tk.StringVar()
        self.out_file_name_var.set(self.config.get_out_file())
        out_file_label = ttk.Label(self, text=self.config.get_main_heading3(), anchor='e', width=14)
        out_file_box = ttk.Entry(self, textvariable=self.out_file_name_var, font=self.master.font)
        self.out_file_label_suffix = ttk.Label(self, text='.xlsx', anchor='w')
        out_file_label.grid(row=2, column=0, pady=10)
        out_file_box.grid(row=2, column=1, sticky=tk.EW, padx=5)
        self.out_file_label_suffix.grid(row=2, column=2, sticky=tk.W)
        
        self.app_btn = ttk.Button(self, width=15, text=const.EXECUTE, style='Exec.TButton',
            command=lambda: self.__execute(self.in_dir_path_var.get(),
                                         self.out_dir_path_var.get(),
                                         self.out_file_name_var.get(),
                                         self.out_file_label_suffix.cget('text'))
            ).grid(row=3, column=1, pady=10, ipady=4)

    def __select_folder(self, folder_path_var):
        path = filedialog.askdirectory()
        folder_path_var.set(os.path.abspath(path))
    
    def __execute(self, in_dir_path, out_dir_path, fname, out_file_suffix):
        if not in_dir_path or not fname:
            messagebox.showwarning(const.WARNING, const.WARN001)
            return
        
        self.__overwrite_config()

        fname = fname + out_file_suffix
        ret = self.exec_method(in_dir_path, out_dir_path, fname)
        if ret == const.PRE_CHECK_ERROR or ret == const.EXCEPTION_OCCURED:
            return
        
        if ret == 0:
            messagebox.showinfo(const.DONE, const.INFO001)
        else:
            messagebox.showwarning(const.DONE, f"{const.WARN003}\nエラーファイル数=[{ret}]")

    def __overwrite_config(self):
        ans = messagebox.askyesno('確認', '今回の設定を残しますか？')
        if ans == True:
            writer = ConfigWriter()
            writer.save(self.in_dir_path_var.get(), self.out_dir_path_var.get(), self.out_file_name_var.get())

        