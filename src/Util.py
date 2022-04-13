import pandas as pd
import os
import sys
import glob
from tkinter import messagebox

class Util:
    @staticmethod
    def total_trans_expences(in_folder_path, out_folder_path, out_file_name):
        target_files = glob.glob(os.path.join(in_folder_path, "*交通費*.xlsx"))
        if len(target_files) == 0:
            messagebox.showwarning('注意', '処理対象のファイルが見つかりませんでした！')
            return None

        info_summary = []
        for f in target_files:
            try:
                df_raw = pd.read_excel(f, sheet_name='交通費', header=None, skiprows=6, usecols='E:M')
                empid = df_raw.iat[0,0]
                empname = df_raw.iat[0,6]
                totalexpence = df_raw.iat[-1,7]
                info = (int(empid), empname.replace('　', ' '), totalexpence)
                info_summary.append(info)
            except FileNotFoundError as e:
                messagebox.showerror('エラー', f"File does not exist! File={f},Error={e}")
            except Exception as e:
                messagebox.showerror('エラー', f"Unexpected Error occured! Error={e}")

        sorted_info_summary = sorted(info_summary, key=lambda x:x[0])
        write_df = pd.DataFrame(sorted_info_summary, columns=['社員ID', '社員名', '交通費合計'])
        output_file = os.path.join(out_folder_path, out_file_name)
        write_df.to_excel(output_file, index=False)
        return output_file
    
    @staticmethod
    def total_attend_records(in_folder_path, out_folder_path, out_file_name):
        messagebox.showinfo('未実装', 'TBA')
        return None

    @staticmethod
    def configure_grid(window, frame, row_idxs, col_idxs=[]):
        for i in row_idxs:
            frame.rowconfigure(i, weight=1)
        
        col_idxs = row_idxs if len(col_idxs) == 0 else col_idxs
        for j in col_idxs:
            frame.columnconfigure(j, weight=1)
        
        window.rowconfigure(0, weight=1)
        window.columnconfigure(0, weight=1)

    @staticmethod
    def get_modifiedpath(relative_path):
        try:
            # only valid for executing .exe (able to get temporary folder path)
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(os.path.dirname(sys.argv[0]))
        
        return os.path.join(base_path, relative_path)
