"""
 多个excel，多个sheet数据合并
 生成的文件名为excel-merge...xlsx
"""
import os
import time
import openpyxl
import xlrd

# 写入的excel是有已经存在表头
has_title_row = False


# 获取当前时间
def get_current_time():
    return time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime())


# 获取excel文件列表
def get_file_path(path, ignore_file):
    # 文件列表
    file_list = []
    # 获取该目录下所有的文件名称和目录名称
    dir_or_files = os.listdir(path)
    for dir_file in dir_or_files:
        # 获取目录或者文件的路径
        dir_file_path = os.path.join(path, dir_file)
        # 该路径是文件并且文件类型为xlsx
        if not os.path.isdir(dir_file_path) and (
                dir_file_path.endswith('.xlsx') or dir_file_path.endswith('.xls')) and ignore_file not in dir_file_path:
            file_list.append(dir_file_path)
    return file_list


# 获取xlsx的列表数据
def get_xlsx_data(file_path):
    new_rows = []
    # 载入xlsx文件
    wb = openpyxl.load_workbook(file_path)
    global has_title_row
    # 循环每个sheet
    for ws_i, ws in enumerate(wb):
        # 循环每一行
        for row_i, row in enumerate(ws.rows):
            # 判断表头是否已存在
            if row_i == 0:
                if has_title_row:
                    continue
                else:
                    has_title_row = True
            write_row = []
            # 循环每一行的所有列
            for col in row:
                # 获取的数据放入write_row列表
                write_row.append(col.value)
            new_rows.append(write_row)
    return new_rows


# 获取xls的列表数据
def get_xls_data(file_path):
    new_rows = []
    # 打开文件，获取excel文件的workbook（工作簿）对象
    workbook = xlrd.open_workbook(file_path)  # 文件路径
    # 获取所有sheet的名字
    sheet_names = workbook.sheet_names()
    global has_title_row
    for sheet_index, sheet_name in enumerate(sheet_names):
        # 通过sheet名获得sheet对象
        worksheet = workbook.sheet_by_name(sheet_name)
        nrows = worksheet.nrows  # 获取该表总行数
        ncols = worksheet.ncols  # 获取该表总列数
        for row_index in range(nrows):
            # 判断表头是否已存在
            if row_index == 0:
                if has_title_row:
                    continue
                else:
                    has_title_row = True
            write_row = []
            for col_index in range(ncols):
                cell = worksheet.cell_value(row_index, col_index)
                write_row.append(cell)
            new_rows.append(write_row)
    return new_rows
