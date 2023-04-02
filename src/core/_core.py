import pandas as pd
import os
import glob
import traceback
import re
from tkinter import messagebox
from datetime import datetime

from config import (
    Constants as const,
    ConfigLoader,
)
from dto import (
    ExpenseSummary,
)
import util

def write_trans_expenses(in_folder_path, out_folder_path, out_file_name)  -> int:
    target_files = glob.glob(os.path.join(in_folder_path, "*交通費*.xlsx"))
    if len(target_files) == 0:
        messagebox.showwarning(const.WARNING, const.WARN002)
        return const.PRE_CHECK_ERROR

    yyyymms = [os.path.basename(f)[0:6] for f in target_files if re.match('[0-9]{6}', os.path.basename(f)[0:6])]
    if not util.is_list_element_unique(yyyymms):
        messagebox.showerror(const.ERROR, const.ERR002)
        return const.PRE_CHECK_ERROR
    
    target_month = datetime.strptime(yyyymms[0],'%Y%m').month

    try:
        emp_df = pd.read_excel(util.get_modifiedpath(os.path.join('resources', 'employees.xlsx'), False), sheet_name='employees', usecols='A:B')
        emps = {x[0]:tuple(x) for x in emp_df.values}
        columns = ConfigLoader().get_columns()
        months = [f"{i}月" for i in range(1,13)]
        columns += months
        output_file = os.path.join(out_folder_path, out_file_name)
        report_file_no_extension = os.path.splitext(os.path.basename(out_file_name))[0]
        report_file = os.path.join(out_folder_path, f"{target_month}_{report_file_no_extension}.txt")
        if os.path.isfile(output_file):
            return __update_summaryfile(target_files, emps, columns, output_file, target_month, report_file)
        else:
            return __create_summaryfile(target_files, emps, columns, output_file, target_month, report_file)

    except PermissionError as e:
        messagebox.showerror(const.ERROR, f"{const.ERR003}\nエラー内容=[{e}]")
        return const.EXCEPTION_OCCURED
    except Exception as e:
        messagebox.showerror(const.ERROR, f"{const.ERR004}\nエラー内容=[{e}]\nエラー詳細=[{traceback.format_exc()}]")
        return const.EXCEPTION_OCCURED

def write_business_expenses(in_folder_path, out_folder_path, out_file_name) -> int:
    messagebox.showinfo(const.NA, const.INFO999)
    return const.PRE_CHECK_ERROR

#TODO 登録と更新でもう少し共通化したい
def __create_summaryfile(files, emp_map, header, ret_filepath, target_month, report_file) -> int:
    tmp_summarys = {}
    for f in files:
        empid, expense, filename = __fetch_file_info(f)
        emp_info = emp_map.get(empid)
        if emp_info is None:
            tmp_summarys[empid] = ExpenseSummary(empid, emp_info, expense, target_month, filename, err_msg='マスタファイルに存在しません。')
            continue
        
        if empid in tmp_summarys.keys():
            old = tmp_summarys.get(empid)
            key = __calc_reverse_seq(empid, tmp_summarys.keys())
            tmp_summarys[key] = ExpenseSummary(empid, old.get_emp_info(), old.get_expense_in_file(), target_month, old.get_filename(), err_msg='社員番号が重複しています。')
            # overwrite existing element. old -> new
            tmp_summarys[empid] = ExpenseSummary(empid, emp_info, expense, target_month, filename, err_msg='社員番号が重複しています。')
            continue

        tmp_summarys[empid] = ExpenseSummary(empid, emp_info, expense, target_month, filename)

    __export_report_file(tmp_summarys, report_file)

    count_include_err = len(tmp_summarys)
    ok_summarys = [tmp_summ.get_expense_summary() for tmp_summ in tmp_summarys.values() if tmp_summ.get_err_msg() is None]
    count_exclude_err = len(ok_summarys)

    empids = [summ_tupl[0] for summ_tupl in ok_summarys]
    ok_summarys += __create_zerolist(emp_map, set(empids))

    sorted_summarys = sorted(ok_summarys, key=lambda x:x[0])
    __write_summarys_to_excel(sorted_summarys, header, ret_filepath)

    return count_include_err - count_exclude_err

def __update_summaryfile(files, emp_map, header, ret_filepath, target_month, report_file) -> int:
    original_df = pd.read_excel(ret_filepath, sheet_name='Sheet1')
    originals = {x[0]:list(x) for x in original_df.values}
    
    tmp_summarys = {}
    for f in files:
        empid, expense, filename = __fetch_file_info(f)
        expenses = []
        emp_info = emp_map.get(empid)
        expenses_with_info = originals.get(empid)
        if expenses_with_info is not None:
            # remove first 2 elems which is empid and empname
            expenses = expenses_with_info[2:]
            # index starts with 0, range is 0 ~ 11, i.e. [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
            expenses[target_month - 1] = expense
        
        if emp_info is None:
            tmp_summarys[empid] = ExpenseSummary(empid, emp_info, expense, target_month, filename, expenses=expenses, err_msg='マスタファイルに存在しません。')
            continue
        
        if empid in tmp_summarys.keys():
            old = tmp_summarys.get(empid)
            key = __calc_reverse_seq(empid, tmp_summarys.keys())
            tmp_summarys[key] = ExpenseSummary(empid, old.get_emp_info(), old.get_expense_in_file(), target_month, old.get_filename(), expenses=old.get_expenses(), err_msg='社員番号が重複しています。')
            # overwrite existing element. old -> new
            tmp_summarys[empid] = ExpenseSummary(empid, emp_info, expense, target_month, filename, expenses=expenses, err_msg='社員番号が重複しています。')
            continue

        tmp_summarys[empid] = ExpenseSummary(empid, emp_info, expense, target_month, filename, expenses=expenses)

    __export_report_file(tmp_summarys, report_file)

    count_include_err = len(tmp_summarys)
    ok_summarys = [tmp_summ.get_expense_summary() for tmp_summ in tmp_summarys.values() if tmp_summ.get_err_msg() is None]
    count_exclude_err = len(ok_summarys)
    
    empids = [summ_tupl[0] for summ_tupl in ok_summarys]
    orig_empids = list(originals.keys())
    inherit_empids = [id for id in orig_empids if id not in empids]

    empids += orig_empids
    ok_summarys += __create_zerolist(emp_map, set(empids))
    
    inherit_summarys = [originals.get(id) for id in inherit_empids]
    ok_summarys += inherit_summarys

    sorted_summarys = sorted(ok_summarys, key=lambda x:x[0])
    __write_summarys_to_excel(sorted_summarys, header, ret_filepath)

    return count_include_err - count_exclude_err

def __fetch_file_info(filepath):
    df_raw = pd.read_excel(filepath, sheet_name='交通費', header=None, skiprows=6, usecols='E:M')
    eid = df_raw.iat[0,0]
    expense = df_raw.iat[-1,7]
    return int(eid), expense, os.path.basename(filepath)

def __create_zerolist(emp_map:dict, excludes:set) -> list:
    zerolist = [0] * 12
    return [emp + tuple(zerolist) for emp in emp_map.values() if emp[0] not in excludes]

def __calc_reverse_seq(empid, empids) -> int:
    minseq = min(empids)
    return -empid - 1 if minseq == -empid else -empid

def __export_report_file(tmp_summarys:dict, filepath:str):
    reports = [tmp_summ.get_report() for tmp_summ in tmp_summarys.values()]
    sorted_reports = sorted(reports)
    util.export_report(sorted_reports, filepath)

def __write_summarys_to_excel(contents:list, header:list, filepath:str):
    write_df = pd.DataFrame(contents, columns=header)
    write_df.to_excel(filepath, index=False)
    