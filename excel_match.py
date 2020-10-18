import openpyxl
import re
from tqdm import tqdm
from openpyxl import Workbook
from utils.common_functions import get_file_path, get_current_time, get_xls_data, get_xlsx_data

# 实例化
wb_write = Workbook()
# 激活 worksheet
wbs = wb_write.active


def write_merge_excel():
    files = get_file_path('./match', 'excel-match')
    files.sort()
    pbar = tqdm(files)
    data_list = []
    for file_index, file_path in enumerate(pbar):
        pattern_result = re.search(r'/match/(.*)\.xls', file_path)
        colums = pattern_result.group(1).split('-')
        # 匹配的列
        macth_col_index = colums[1]
        # 保留的列
        retain_colums = [int(col_num_str) if col_num_str != '*' else col_num_str for col_num_str in
                         colums[2].split(',')]

        if file_path.endswith('xls'):
            new_rows = get_xls_data(file_path, False)
        else:
            new_rows = get_xlsx_data(file_path, False)
        data_map = {}
        for row_values in new_rows:
            new_values = [row_value for row_index, row_value in enumerate(row_values) if
                          row_index + 1 in retain_colums or '*' in retain_colums]
            data_map[row_values[int(macth_col_index) - 1]] = new_values
        data_list.append(data_map)

    for match_value, row_list in data_list[0].items():
        new_row = []
        new_row += row_list
        for i in range(1, len(data_list)):
            match_row = ['' for i_col in list(data_list[i].values())[0]]
            if match_value in data_list[i].keys():
                match_row = data_list[i][match_value]
            new_row += match_row
        wbs.append(new_row)
        # 保存文件
    wb_write.save('./match/excel-match' + get_current_time() + '.xlsx')


write_merge_excel()
