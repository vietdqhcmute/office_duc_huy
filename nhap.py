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
from PIL import Image, ImageDraw, ImageFont

folder_hinh = input("Hãy nhập folder chứa hình: ")
folder_hinh = format_path(folder_hinh)

excel_path = r'c:\Users\Huy\OneDrive\RaSoatDoDac\Khu500\20230928_TONG_380HA 500HA LA_(Sửa danh mục hồ sơ kèm theo).xlsx'
ws = openpyxl.load_workbook(excel_path)['Tổng hợp']

folder_xuat = input("Hãy nhập folder chứa kết quả: ")
folder_xuat = format_path(folder_xuat)

width_ratio = 0.8
font_family = 'arial.ttf'

def get_fontsize(img, text, width_ratio):
    fontsize = 1
    font = ImageFont.truetype("arial.ttf", fontsize)
    while True: 
        if font.getlength(text) < width_ratio*img.size[0]:
    # iterate until the text size is just larger than the criteria
            fontsize += 1
            font = ImageFont.truetype("arial.ttf", fontsize, encoding= "utf-8")
        else:
            break
    return font


dict_text = {}
for rid in range(3, 552):
    mbv = lcell(ws, rid, "C")
    ten = lcell(ws, rid, "D")
    if mbv != "None":
        dict_text[mbv] = "MBV_{}\n{}".format(mbv,ten)
for filename in os.listdir(folder_hinh):
    ma = filename.replace(".png","")
    img = Image.open(os.path.join(folder_hinh, filename))
    draw = ImageDraw.Draw(img)
    font = get_fontsize(img, dict_text[ma], width_ratio)
    draw.multiline_textbbox((2,0), text = dict_text[ma], font = font)    
    img.save(os.path.join(folder_xuat, filename))
