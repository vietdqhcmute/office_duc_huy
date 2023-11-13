import os
import openpyxl
import pandas as pd
import pyautogui as pag
import re
from PIL import Image, ImageFont
from docx.shared import Pt, Inches
from libs.Functions import format_path
from libs.doc_function import *
from libs.excel_function import *
from libs.Mail_Merged_v2 import Mail_Merged

def get_fontsize(img, text, width_ratio):
    fontsize = 1
    font = ImageFont.truetype("arial.ttf", fontsize)
    while True:
        # rest of the function code here

def main():
    folder_hinh = input("Hãy nhập folder chứa hình: ")
    folder_hinh = format_path(folder_hinh)

    excel_path = r'c:\Users\Huy\OneDrive\RaSoatDoDac\Khu500\20230928_TONG_380HA 500HA LA_(Sửa danh mục hồ sơ kèm theo).xlsx'
    ws = openpyxl.load_workbook(excel_path)['Tổng hợp']

    folder_xuat = input("Hãy nhập folder chứa kết quả: ")
    folder_xuat = format_path(folder_xuat)

    width_ratio = 0.8
    font_family = 'arial.ttf'

if __name__ == "__main__":
    main()
