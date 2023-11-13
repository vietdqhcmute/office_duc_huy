from libs.Functions import *
import numpy as np
import pandas as pd
from openpyxl.utils import column_index_from_string
import openpyxl, docx, shutil
from os.path import join

def get_workbook_path():
    wb_path = input("Hãy nhập đường dẫn file excel: ")
    return format_path(wb_path)

def get_sheet_name(wb_path):
    print("File excel gồm các sheet: ")
    dict_sheet = {}
    count = 1
    for sheet in openpyxl.load_workbook(wb_path).sheetnames:
        print(str(count) + ".", sheet)
        dict_sheet[count] = sheet
        count +=1
    option = int(input("Hãy chọn sheet sử dụng: "))
    return dict_sheet[option]

def get_columns_to_format():
    cot_so = (input("Hãy nhập các cột bạn muốn định dạng số cách nhau bằng dấu \",\": "))
    cot_so_idx = []
    if len(cot_so)>0:
        if " " in cot_so:
            cot_so = cot_so.replace(" ","")
        cot_so = cot_so.split(",")
        for char in cot_so:
            cot_so_idx.append(column_index_from_string(char))
    return cot_so_idx

def Mail_Merged():
    wb_path = get_workbook_path()
    ws_name = get_sheet_name(wb_path)
    cot_so_idx = get_columns_to_format()
    # ... rest of the function ...
