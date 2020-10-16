"""
 多个excel，多个sheet数据合并
 生成的文件名为excel-merge...xlsx
"""
import openpyxl
from openpyxl import Workbook
import os
import time


# 获取当前时间
def get_current_time():
    return time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime())


# 获取当前路径下的所有xlsx文件
def get_file_path():
    # 路径为当前代码路径
    root_path = '.'
    # 文件列表
    file_list = []
    # 获取该目录下所有的文件名称和目录名称
    dir_or_files = os.listdir(root_path)
    for dir_file in dir_or_files:
        # 获取目录或者文件的路径
        dir_file_path = os.path.join(root_path, dir_file)
        # 该路径是文件并且文件类型为xlsx
        if not os.path.isdir(dir_file_path) and dir_file_path.endswith('.xlsx') and not dir_file_path.startswith(
                './excel-merge'):
            file_list.append(dir_file_path)
    return file_list


# 实例化
wb_write = Workbook()

# 激活 worksheet
wbs = wb_write.active
for i, file_name in enumerate(get_file_path()):
    # 载入xlsx文件
    wb = openpyxl.load_workbook(file_name)
    for ws_i, ws in enumerate(wb):
        # 循环每一行
        for row_i, row in enumerate(ws.rows):
            # 第一个文件之后的文件过滤表头
            if i != 0 and ws_i == 0 and row_i == 0:
                continue
            if i == 0 and ws_i != 0 and row_i == 0:
                continue
            write_row = []
            # 循环每一行的所有列
            for col in row:
                # 获取的数据放入write_row列表
                write_row.append(col.value)
            # 写入激活的 worksheet
            wbs.append(write_row)

# 保存文件
wb_write.save('excel-merge' + get_current_time() + '.xlsx')
