"""
    pip install openpyxl
    pip install xlrd
    pip install tqdm
    多个excel，多个sheet数据合并
    生成的文件名为excel-merge...xlsx
"""
import openpyxl
import xlrd
from tqdm import tqdm
from openpyxl import Workbook
from utils.common_functions import get_file_path, get_current_time

# 实例化
wb_write = Workbook()
# 激活 worksheet
wbs = wb_write.active
# 写入的excel是有已经存在表头
has_title_row = False


def write_merge_excel():
    files = get_file_path('./merge', 'excel-merge')
    pbar = tqdm(files)
    for file_index, file_path in enumerate(pbar):
        if file_path.endswith('xls'):
            _get_xls_data(file_path)
        else:
            _get_xlsx_data(file_path)
    # 保存文件
    wb_write.save('./merge/excel-merge' + get_current_time() + '.xlsx')


def _get_xlsx_data(file_path):
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
            # 写入激活的 worksheet
            wbs.append(write_row)


def _get_xls_data(file_path):
    # 打开文件，获取excel文件的workbook（工作簿）对象
    workbook = xlrd.open_workbook(file_path)  # 文件路径
    # 获取所有sheet的名字
    sheet_names = workbook.sheet_names()
    global has_title_row
    for sheet_index, sheet_name in enumerate(sheet_names):
        # 通过sheet名获得sheet对象
        worksheet = workbook.sheet_by_name(sheet_name)
        nrows = worksheet.nrows  # 获取该表总行数
        for row_index in range(nrows):
            # 判断表头是否已存在
            if row_index == 0:
                if has_title_row:
                    continue
                else:
                    has_title_row = True
            # 写入激活的 worksheet
            wbs.append(worksheet.row_values(row_index))


# 执行
write_merge_excel()
