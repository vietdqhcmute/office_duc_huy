from libs.Functions import *
from libs.doc_function import *
from libs.excel_function import *
import os, openpyxl, shutil, docx
from os.path import join
from docx.shared import Pt, Inches
from libs.Mail_Merged_v2 import Mail_Merged
import pandas as pd
import math
import pyautogui as pag
import re
import xlwings as xlw
import tkinter as tk
from tkinter import ttk

wb = xlw.books.active
ws = wb.sheets["Danh sách lo dat"]
for cell in ws[3:708,16]:
    if cell.value != None:
        lst_gcn = cell.value.split("\n")
        lst_ngay = []
        for value in lst_gcn:
            if "ngày" in value:
                while "ngày" in value:
                    value = value[value.find("ngày")+4:]
                lst_ngay.append(value)
                
        print(lst_ngay)