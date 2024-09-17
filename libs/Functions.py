def clean_path(_path):
    if "& " in _path or "'" in _path or '"' in _path:
        _path = _path.replace("'","")
        _path = _path.replace('"',"")
        _path = _path.replace('& ',"")
    return _path

def get_vnd_word(so_tien):
    x1 = ["không","một","hai","ba","bốn","năm","sáu","bảy","tám","chín"]
    x2 = ["","nghìn","triệu","tỉ","nghìn tỉ","triệu tỉ","tỉ tỉ"]
    return x1, x2

def VND(so_tien):
    if so_tien == 0: return "Không đồng"
    elif so_tien < 0:
        neg = True
        so_tien = abs(so_tien)
    else: neg = False
    x1, x2 = get_vnd_word(so_tien)
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
