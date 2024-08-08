import os

import re
import tkinter as tk
import xlrd3 
import xml.etree.ElementTree
import json

from pathlib import *
from decimal import *
from tkinter.filedialog import *
from tkinter.messagebox import *


FLAG_CLIENT = 'clinet'
FLAG_SERVERS ='servers'
MIN_ROW = 5
SHEET_FLAGS_COUNT = 4

def convert_to_float(value):
    if isinstance(value, Decimal):
        return float(value)
    return value

def open_excel(file_path):
    try:
        return xlrd3.open_workbook(file_path)
    except Exception as e:
        print(f"open {file_path.name}错误信息:{e}")
        return None

def writeHeadToFile(fd, cells_flag, cell):
    headList = []
    for index, item in enumerate(cell):
        if cells_flag[index]:
            # print(f'{index}\t{item}')
            headList.append(item)
    txt_line = '\t'.join(str(item) for item in headList)
    # print(f'txt_line:{txt_line}')
    fd.write(txt_line+'\n')

def writeBodyToFile(fd, cells_flag, cells_type, cell):
    bodyList = []
    for index, item in enumerate(cell):
        if cells_flag[index]:
            if 'uint8' == cells_type[index] or 'uint16' == cells_type[index] or 'uint32' == cells_type[index] :
                bodyList.append(int(item))
            elif 'float' == cells_type[index] :
                bodyList.append(Decimal(float(item)).quantize(Decimal('.01'), rounding=ROUND_DOWN))
            elif 'string' == cells_type[index] :
                bodyList.append(item)
            elif  'array' == cells_type[index] :
                res_list = item.split(r'|')
                for it in res_list :
                    pattern = re.compile(r'\d+;\d+')
                    m = pattern.match(it)
                    if not m: 
                        print(f"\t\t{cells_type[index]}数据格式错误")
                        exit(1)
                bodyList.append(item)
            else :
                print(f'unknow type :{cells_type[index]}')
                exit(1)
            # print(f'index:{index}\ttype:{cells_type[index]}\t{item}') 
    txt_line = '\t'.join(str(item) for item in bodyList)
    # print(f'txt_line:{txt_line}')
    fd.write(txt_line+'\n')

def toTxt(output_dir, file_name, obj_name, flag, tables):
    lenght = len(tables)
    if lenght <= MIN_ROW:
        print(f'{file_name} row count = {lenght} ')
        return False
    
    # showinfo(f'output_dir:{output_dir}')
    out_file_path = PurePath.joinpath(Path(output_dir),flag,file_name+'.json')
    # showinfo(f'out_file_path:{out_file_path}')
    try:
        file = open(out_file_path, "w+", encoding='utf-8')
    except Exception as e:
        print(f"\t\tcreate {out_file_path}失败,错误信息:{e}")
        return False
    
    file.truncate()

    cells_name = tables[0]
    cells_flag = []
    if FLAG_CLIENT == flag:
        cells_flag = tables[3]
    elif FLAG_SERVERS == flag :
        cells_flag = tables[4]
    
    cells_type = tables[2]
    
    # writeHeadToFile(file, cells_flag, tables[0])
    # writeHeadToFile(file, cells_flag, tables[2])
    datas = [] 
    for row in range(MIN_ROW,lenght): 
        # writeBodyToFile(file, cells_flag, cells_type, tables[row])
        item_data = {}
        for index, item in enumerate(tables[row]):
            print(f'name:{cells_name[index]}\ttype:{cells_type[index]}\tvalue:{item}')
            if cells_flag[index]:
                if 'uint8' == cells_type[index] or 'uint16' == cells_type[index] or 'uint32' == cells_type[index] :
                    # bodyList.append(int(item))
                    item_data[cells_name[index]] = int(item)
                elif 'float' == cells_type[index] :
                    # bodyList.append(Decimal(float(item)).quantize(Decimal('.01'), rounding=ROUND_DOWN))
                    fvalue = Decimal(float(item)).quantize(Decimal('.01'), rounding=ROUND_DOWN)
                    item_data[cells_name[index]] = convert_to_float(fvalue)
                elif 'string' == cells_type[index] :
                    # bodyList.append(item)
                    print(item)
                    item_data[cells_name[index]] = item
                elif 'array_coord' == cells_type[index] :
                    # bodyList.append(item)
                    if not item :
                        item_data[cells_name[index]] = []
                    else :
                        res_list = item.split(r'|')
                        coord_list = []
                        for it in res_list :
                            pattern = re.compile(r'[+-]?\d+;[+-]?\d+')
                            m = pattern.match(it)
                            if m:
                                coord = m.group().split(';')
                                tmp = {}
                                tmp['x'] = int(coord[0])
                                tmp['y'] = int(coord[1])
                                coord_list.append(tmp)
                            else :
                                print(f"\t\t{cells_type[index]}数据格式错误")
                                exit(1)
                        # bodyList.append(item)
                        print(coord_list)
                        item_data[cells_name[index]] = coord_list
                elif 'array_item' == cells_type[index] :
                    # bodyList.append(item)
                    if not item :
                        item_data[cells_name[index]] = []
                    else :
                        res_list = item.split(r'|')
                        item_list = []
                        for it in res_list :
                            pattern = re.compile(r'[+]?\d+;[+]?\d+;[+]?\d+')
                            m = pattern.match(it)
                            if m:
                                coord = m.group().split(';')
                                tmp = {}
                                tmp['gtype'] = int(coord[0])
                                tmp['gid'] = int(coord[1])
                                tmp['count'] = int(coord[2])
                                item_list.append(tmp)
                            else :
                                print(f"\t\t{cells_type[index]}数据格式错误")
                                exit(1)
                        # bodyList.append(item)
                        print(item_list)
                        item_data[cells_name[index]] = item_list
                elif  'array' == cells_type[index] :
                    if not item :
                        item_data[cells_name[index]] = ""
                    else :
                        res_list = item.split(r'|')
                        for it in res_list :
                            pattern = re.compile(r'[+-]?\d+;[+-]?\d+')
                            m = pattern.match(it)
                            if not m: 
                                print(f"\t\t{cells_type[index]}数据格式错误")
                                exit(1)
                            print(m.group())
                            print(type(m.group()))
                        # bodyList.append(item)
                        print(item)
                        item_data[cells_name[index]] = item
                else :
                    print(f'unknow type :{cells_type[index]}')
                    exit(1)
            # print(f'index:{index}\ttype:{cells_type[index]}\t{item}') 
        datas.append(item_data)
    if FLAG_CLIENT == flag:
        outdata = {}
        outdata[obj_name]= datas
        json_string = json.dumps(outdata, ensure_ascii=False)
        file.write(json_string)
    elif FLAG_SERVERS == flag :
        json_string = json.dumps(datas, ensure_ascii=False)
        file.write(json_string)

    file.close()
    return True  

def excel_table_by_index(file_path, out_file_path):
    data = open_excel(file_path)
    if data is None:
        print(f"打开{file_path.name}失败")
        return False

    sheet_names = data.sheet_names()
    # showinfo(f"sheets{sheet_names} len:{len(sheet_names)}")
    if len(sheet_names) <= 0:
        print(f"\t\t{file_path.name}sheet is null")
        return False
    
    for sheet_name in sheet_names :
        sheet_flags = sheet_name.split('_')
        if len(sheet_flags) != SHEET_FLAGS_COUNT :
            print(f"\t\tsheet:{sheet_name} sheet_flags invalid")
            continue
        if sheet_flags[0] != '1' :
            continue

        table = data.sheet_by_name(sheet_name)
        row_data_list = []
        for row in range(table.nrows):
            columns = table.row_values(row) 
            row_data_list.append(columns)
        
        bRet = toTxt(out_file_path, sheet_flags[1],sheet_flags[2], FLAG_CLIENT, row_data_list)
        if False == bRet :
            print(f'\t\t{FLAG_CLIENT} {sheet_name} {sheet_flags[1]} toTxt failed')
            return False
        print(f"\t\tsheet:{sheet_name} {FLAG_CLIENT}  toTxt success")

        bRet = toTxt(out_file_path, sheet_flags[1], sheet_flags[2], FLAG_SERVERS, row_data_list)
        if False == bRet :
            print(f'\t\t{sheet_name} {FLAG_SERVERS} {sheet_flags[1]} toTxt failed')
            return False
        print(f"\t\tsheet:{sheet_name} {FLAG_SERVERS} toTxt success")
    return True

def get_files(directory_path, out_file_path):
    cur_dir = Path(directory_path)
    # for root, dirs, files in os.walk(directory_path):
    files = [file for file in cur_dir.glob("**/*")]
    for file in files:
        if file.is_file():
            # showinfo(f'file:{file.name}')
            file_name = file.stem
            if file_name.startswith('.') or  file_name.startswith('~'):
                continue 
            file_suffix = file.suffix
            if '.xlsx' != file_suffix.lower():
                # print(f"{file} suffix is not .xlsx")
                continue
            # file_path = PurePath.joinpath(root, file)
            print(f'++++++++>>>>>>start {file}')
            if excel_table_by_index(file, out_file_path) :
                print(f'++++++++>>>>>>end {file} success')
            else : 
                print(f'++++++++>>>>>>end {file} failed')
                return False
    return True
# def main():
#     print("=====main")
#     client_dir = os.path.join(DIRECTORY_OUTPUT_PATH, FLAG_CLIENT) 
#     if False == os.path.exists(client_dir):
#         os.mkdir(client_dir)
   
#     servers_dir = os.path.join(DIRECTORY_OUTPUT_PATH, FLAG_SERVERS)
#     if False == os.path.exists(servers_dir):
#         os.mkdir(servers_dir)
#     get_files(DIRECTORY_SRC_PATH)

def startTask():
    directory_src_path = src_dir.get()
    directory_output_path = output_dir.get()
    print(directory_src_path)
    print(directory_output_path)
    
    client_dir = os.path.join(directory_output_path, FLAG_CLIENT) 
    if not Path(client_dir).exists():
        Path.mkdir(Path(client_dir)) 
    
    servers_dir = os.path.join(directory_output_path, FLAG_SERVERS)
    if not Path(servers_dir).exists():
        Path.mkdir(Path(servers_dir))
    
    if get_files(directory_src_path, directory_output_path):
        print(showinfo("information","导出完成！！！"))
    else :
        print(showinfo("information","导出失败！！！"))

def openSrcDir():
    fileDir = askdirectory()  # 选择目录，返回目录名
    if fileDir.strip() != '':
        src_dir.set(fileDir)  # 设置变量src_dir的值
    else:
        print("do not choose Dir")
    
    # cur_dir = Path(src_dir.get())
    # files = [file for file in cur_dir.glob("**/*")]
    # for file in files:
    #     if file.is_file():
    #         print(file)
    #         print(file.name)  # 返回文件名+文件后缀
    #         print(file.stem)  # 返回文件名
    #         print(file.suffix)  # 返回文件后缀

def openOutputDir():
    fileDir = askdirectory()  # 选择目录，返回目录名
    if fileDir.strip() != '':
        output_dir.set(fileDir)  # 设置变量output_dir的值
    else:
        print("do not choose Dir")

if __name__ == "__main__":
    # main()

    root = tk.Tk()
    root.title("excel2txt")
    #  root.geometry('800x600')
    root.resizable(False, False)
    src_dir = tk.StringVar()
    output_dir = tk.StringVar()

    # 设置源目录
    tk.Label(root, text='Excel目录').grid(row=1, column=0, padx=5, pady=5) # 创建label 提示这是选择目录
    tk.Entry(root, textvariable=src_dir).grid(row=1, column=1, padx=5, pady=5) # 创建Entry，显示选择的目录
    tk.Button(root, text='打开目录', command=openSrcDir).grid(row=1, column=2, padx=5, pady=5)

    # 选择目标目录
    tk.Label(root, text='导出目录').grid(row=2, column=0, padx=5, pady=5) # 创建label 提示这是选择目录
    tk.Entry(root, textvariable=output_dir).grid(row=2, column=1, padx=5, pady=5) # 创建Entry，显示选择的目录
    tk.Button(root, text='打开目录', command=openOutputDir).grid(row=2, column=2, padx=5, pady=5)
    
    tk.Button(root, text='导出', command=startTask).grid(row=3, column=2, padx=5, pady=5)
    
    root.mainloop()

