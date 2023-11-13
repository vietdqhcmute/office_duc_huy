#Hàm liên quan đến sheet
def isheet(excel_path, sheet_index):
    import os, openpyxl
    wb = openpyxl.load_workbook(excel_path)
    ws = wb[wb.sheetnames[sheet_index]]
    return ws
def lsheet(excel_path, sheet_name):
    import os, openpyxl
    wb = openpyxl.load_workbook(excel_path)
    ws = wb[sheet_name]
    return ws
#Hàm đọc cell
def lcell (ws,row_index, column_letter):
    value = str(ws[str(column_letter) + str(row_index)].value)
    return value
def icell (ws,row_index, column_index):
    value = str(ws.cell(row = row_index, column = column_index).value)
    return value
def cell_coordinate(ws, row_index, column_index):
    coordinate = ws.cell(row = row_index, column = column_index).coordinate
    return coordinate
#Hàm ghi lên cell
def w_icell(ws, row_index, column_index, content):
    ws.cell(row = row_index, column = column_index).value = content
def w_lcell(ws, row_index, column_letter, content):
    ws[str(column_letter) + str(row_index)] = content