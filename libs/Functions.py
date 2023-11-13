#Các hàm định dạng

    #kiểm tra có phải là số hay không
def check_so(num): 
    num = num.replace(".","")
    return num.isnumeric()

    #Định dạng lại số theo kiểu Việt Nam
def num_format(value): 
    if "." in value:
        num = float(value)
        num = str(f"{num:,.1f}")
        _lst_char = list(num)
        for i in range(0, len(_lst_char)):
            if _lst_char[i] == ",":     
                _lst_char[i] = "."
            elif _lst_char[i] == ".":
                _lst_char[i] = ","
        num = "".join(_lst_char)
        return num
    else:
        num = int(value)
        num = str(f"{num:,}")
        _lst_char = list(num)
        for i in range(0, len(_lst_char)):
            if _lst_char[i] == ",":     
                _lst_char[i] = "."
            elif _lst_char[i] == ".":
                _lst_char[i] = ","
        num = "".join(_lst_char)
        return num
    
    #Hàm chỉnh lại chuỗi path khi thả file vào chương trình
def format_path(_path):
    if "& " in _path or "'" in _path or '"' in _path:
        _path = _path.replace("'","")
        _path = _path.replace('"',"")
        _path = _path.replace('& ',"")
    return _path
    
    #Hàm chuyển số thành chữ VNĐ
def VND(so_tien):
    if so_tien == 0: return "Không đồng"
    elif so_tien < 0:
        neg = True
        so_tien = abs(so_tien)
    else: neg = False
    x1 = ["không","một","hai","ba","bốn","năm","sáu","bảy","tám","chín"]
    x2 = ["","nghìn","triệu","tỉ","nghìn tỉ","triệu tỉ","tỉ tỉ"]
    x3 = str(so_tien)
    x4 = len(x3)
    x5 = [0]*x4
    so_tien = ""
    for i in range(x4):
        so = int(x3[x4-i-1])
        if so==0 and i%3==0 and x5[i]==0:
            for j in range(i+1,x4):
                so1 = int(x3[x4-j-1])
                if so1!=0: break
            if (j-i)//3>0:
                for k in range(i,((j-1)//3)*3+i):x5[k] = 1
    for i in range(x4):
        so = int(x3[x4-i-1])
        if x5[i]==1: continue
        if i%3==0 and i>0: so_tien = x2[i//3]+" "+ so_tien
        if i%3==1: so_tien = "mươi " + so_tien
        if i%3==2: so_tien = "trăm " + so_tien
        so_tien = x1[so] + " " + so_tien
    so_tien = so_tien.replace("không mươi","linh").replace("linh không","").replace("mươi không","mươi")
    so_tien = so_tien.replace("một mươi","mười").replace("mươi năm","mười lăm").replace("mươi một","mười một")
    so_tien = so_tien.replace("mười năm","mười lăm").strip().split(", ")
    so_tien[-1] = so_tien[-1].replace("linh","lẻ")
    so_tien = ", ".join(so_tien)
    if neg: return "Âm " + so_tien + " đồng"
    else: return so_tien.capitalize() + " đồng"

#Hàm xử lý file
    #Hàm đổi tên
def rename(folder_path, oldname, newname):
    import os
    from os.path import join
    os.rename(join(folder_path, oldname), join(folder_path, newname))

    #Hàm copy
def copy(src_path, des_path):
    import shutil
    shutil.copy2(src_path, des_path)