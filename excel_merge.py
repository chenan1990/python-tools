"""
 多个excel，多个sheet数据合并
 生成的文件名为excel-merge...xlsx
"""
import openpyxl
import xlrd
from openpyxl import Workbook
from utils.common_functions import get_file_path, get_current_time


# 实例化
wb_write = Workbook()

# 激活 worksheet
wbs = wb_write.active
files = get_file_path('./merge', 'excel-merge')
for i, file_name in enumerate(files):
    # 载入xlsx文件
    wb = openpyxl.load_workbook(file_name)
    # 循环每个sheet
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
wb_write.save('./merge/excel-merge' + get_current_time() + '.xlsx')
