"""
    pip install openpyxl
    pip install xlrd
    pip install tqdm
    多个excel，多个sheet数据合并
    生成的文件名为excel-merge...xlsx
"""

from tqdm import tqdm
from openpyxl import Workbook
from utils.common_functions import get_file_path, get_current_time, get_xls_data, get_xlsx_data

# 实例化
wb_write = Workbook()
# 激活 worksheet
wbs = wb_write.active


def write_merge_excel():
    files = get_file_path('./merge', 'excel-merge')
    pbar = tqdm(files)
    for file_index, file_path in enumerate(pbar):
        if file_path.endswith('xls'):
            new_rows = get_xls_data(file_path)
        else:
            new_rows = get_xlsx_data(file_path)
        for row in new_rows:
            wbs.append(row)
    # 保存文件
    wb_write.save('./merge/excel-merge' + get_current_time() + '.xlsx')


# 执行
write_merge_excel()
