import os
import sys

def configure_grid(window, frame, row_idxs, col_idxs=[]):
    for i in row_idxs:
        frame.rowconfigure(i, weight=1)
    
    col_idxs = row_idxs if len(col_idxs) == 0 else col_idxs
    for j in col_idxs:
        frame.columnconfigure(j, weight=1)
    
    window.rowconfigure(0, weight=1)
    window.columnconfigure(0, weight=1)


def get_modifiedpath(relative_path, isPackedFile=True):
    try:
        # only valid for executing .exe (able to get temporary folder path)
        base_path = sys._MEIPASS if isPackedFile else os.path.abspath(os.path.dirname(sys.argv[0]))
    except Exception:
        base_path = os.path.abspath(os.path.dirname(sys.argv[0]))
    
    return os.path.join(base_path, relative_path)

def is_list_element_unique(list):
    return True if len(set(list)) == 1 else False

def get_duplicate_vals(list):
    counted = set()
    return [x for x in list if x in counted or (counted.add(x) or False)]

def export_report(msgs, filename):
    reports = map(lambda x: x + '\n', msgs)
    filepath = get_modifiedpath(filename, False)
    with open(filepath, 'w', encoding='utf-8') as f:
        f.writelines(reports)

