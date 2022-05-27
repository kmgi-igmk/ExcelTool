import pandas as pd
import os
import sys
import glob
import traceback
from tkinter import messagebox
from datetime import datetime

import util

def write_trans_expenses(in_folder_path, out_folder_path, out_file_name):
    target_files = glob.glob(os.path.join(in_folder_path, "*交通費*.xlsx"))
    if len(target_files) == 0:
        messagebox.showwarning('注意', '処理対象のファイルが見つかりませんでした!')
        return None
    try:
        emp_df = pd.read_excel(util.get_modifiedpath(os.path.join('resources', 'employees.xlsx'), False), sheet_name='employees')
        emps = {x[0]:tuple(x) for x in emp_df.values}
        columns = ['社員ID', '社員名']
        months = [f"{i}月" for i in range(1,13)]
        columns += months
        output_file = os.path.join(out_folder_path, out_file_name)
        if os.path.isfile(output_file):
            return __update_summaryfile(target_files, emps, columns, output_file)
        else:
            return __create_summaryfile(target_files, emps, columns, output_file)

    except PermissionError as e:
        messagebox.showerror('エラー', f"処理対象のファイルが開かれている可能性があります!\nファイルを閉じてから再度実行してください\nエラー内容=[{e}]")
        return None
    except Exception as e:
        messagebox.showerror('エラー', f"エラーが発生しました!\nエラー内容=[{e}]\n詳細=[{traceback.format_exc()}]")
        return None

def write_attend_records(in_folder_path, out_folder_path, out_file_name):
    messagebox.showinfo('未実装', 'TBA')
    return None

def __create_summaryfile(files, emp_map, header, ret_filepath):
    info_summary = []
    file_months = []
    for f in files:
        filename = os.path.basename(f)
        file_months.append(datetime.strptime(filename[0:6],'%Y%m').month)
        if len(list(set(file_months))) > 1:
            raise ValueError('処理月が異なるファイルが含まれています!')

        empid, totalexpense = __get_writteninfo(f)
        emp_info = emp_map.get(int(empid))
        if emp_info is None: raise ValueError(f"[{empid}]番は存在しません・・")

        month = list(set(file_months))[0]
        expenses = [totalexpense if x == month else 0 for x in range(1,13)]
        info_summary.append(emp_info + tuple(expenses))
    
    enrolled_ids = [tup[0] for tup in info_summary]
    info_summary += __createzerolist(emp_map, enrolled_ids)
    sorted_info_summary = sorted(info_summary, key=lambda x:x[0])
    write_df = pd.DataFrame(sorted_info_summary, columns=header)
    write_df.to_excel(ret_filepath, index=False)
    return f"作成ファイル=[{ret_filepath}]"

def __update_summaryfile(files, emp_map, header, ret_filepath):
    original_df = pd.read_excel(ret_filepath, sheet_name='Sheet1')
    originals = {x[0]:list(x) for x in original_df.values}
    
    info_summary = []
    file_months = []
    for f in files:
        filename = os.path.basename(f)
        file_months.append(datetime.strptime(filename[0:6],'%Y%m').month)
        if len(list(set(file_months))) > 1:
            raise ValueError('処理月が異なるファイルが含まれています!')
        
        empid, totalexpense = __get_writteninfo(f)
        emp_info = emp_map.get(int(empid))
        if emp_info is None: raise ValueError(f"[{empid}]番は存在しません・・")

        expenses = originals.get(int(empid))
        month = list(set(file_months))[0]
        expenses[month + 1] = totalexpense
        info_summary.append(tuple(expenses))

    enrolled_ids = list(originals.keys())
    info_summary += __createzerolist(emp_map, enrolled_ids)
    exclude_ids = [tup[0] for tup in info_summary]
    info_summary += __filterlist(originals, exclude_ids)
    sorted_info_summary = sorted(info_summary, key=lambda x:x[0])
    write_df = pd.DataFrame(sorted_info_summary, columns=header)
    write_df.to_excel(ret_filepath, index=False)
    return f"更新ファイル=[{ret_filepath}]"

def __get_writteninfo(filepath):
    df_raw = pd.read_excel(filepath, sheet_name='交通費', header=None, skiprows=6, usecols='E:M')
    eid = df_raw.iat[0,0]
    expense = df_raw.iat[-1,7]
    return eid, expense

def __createzerolist(emp_map, excludes):
    zerolist = [0] * 12
    return [emp + tuple(zerolist) for emp in emp_map.values() if emp[0] not in excludes]

def __filterlist(orig_map, excludes):
    return [tuple(orig) for orig in orig_map.values() if orig[0] not in excludes]
