import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from Util import Util

class TemplateExecutor(ttk.Frame):
    def __init__(self, master, exec_method):
        super().__init__(master)
        self.exec_method = exec_method

        self.grid(row=1, padx=10, sticky=tk.NSEW)
        self.create_widgets()
        Util.configure_grid(self.master, self, list(range(0,3)))
        
    def create_widgets(self):       
        in_folder_path_var = tk.StringVar()
        in_folder_label = ttk.Label(self, text='入力フォルダ指定:', anchor='e', width=14)
        in_folder_box = ttk.Entry(self, textvariable=in_folder_path_var, width=30, font=self.master.font)
        in_folder_btn = ttk.Button(self, text='参照', command=lambda: self.select_folder(in_folder_path_var))
        in_folder_label.grid(row=0, column=0, pady=10)
        in_folder_box.grid(row=0, column=1, sticky=tk.EW, padx=5)
        in_folder_btn.grid(row=0, column=2, padx=5)
        
        out_folder_path_var = tk.StringVar()
        out_folder_label = ttk.Label(self, text='出力フォルダ指定:', anchor='e', width=14)
        out_folder_box = ttk.Entry(self, textvariable=out_folder_path_var, width=30, font=self.master.font)
        out_folder_btn = ttk.Button(self, text='参照', command=lambda: self.select_folder(out_folder_path_var))
        out_folder_label.grid(row=1, column=0, pady=10)
        out_folder_box.grid(row=1, column=1, sticky=tk.EW, padx=5)
        out_folder_btn.grid(row=1, column=2, padx=5)
        
        out_file_name_var = tk.StringVar()
        out_file_label = ttk.Label(self, text='出力ファイル名:', anchor='e', width=14)
        out_file_box = ttk.Entry(self, textvariable=out_file_name_var, font=self.master.font)
        out_file_label_suffix = ttk.Label(self, text='.xlsx', anchor='w')
        out_file_label.grid(row=2, column=0, pady=10)
        out_file_box.grid(row=2, column=1, sticky=tk.EW, padx=5)
        out_file_label_suffix.grid(row=2, column=2, sticky=tk.W)
        
        app_btn = ttk.Button(self, width=15, text='実行', style='Exec.TButton',
            command=lambda: self.execute(in_folder_path_var.get(),
                                         out_folder_path_var.get(),
                                         out_file_name_var.get(),
                                         out_file_label_suffix.cget('text'))
            ).grid(row=3, column=1, pady=10, ipady=4)

    def select_folder(self, folder_path_var):
        path = filedialog.askdirectory()
        folder_path_var.set(os.path.abspath(path))
    
    # ボタン押下時のアクション
    def execute(self, in_folder_path, out_folder_path, fname, out_file_suffix):
        if not in_folder_path or not fname:
            messagebox.showwarning('注意', '入力フォルダと出力ファイル名の両方を入力してください！')
            return
        
        fname = fname + out_file_suffix
        result_file = self.exec_method(in_folder_path, out_folder_path, fname)
        if result_file is None:
            return
        else:
            messagebox.showinfo('完了', "作成が完了しました！\n出力ファイル=" + result_file)
