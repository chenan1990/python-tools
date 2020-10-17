import openpyxl
from openpyxl import Workbook
from utils.common_functions import get_file_path, get_current_time


def format_excel_data(args):
    data_list = []
    # 实例化
    wb_write = Workbook()

    # 激活 worksheet
    wbs = wb_write.active
    for i, file_name in enumerate(get_file_path('./match', 'excel-match')):
        # 载入xlsx文件
        wb = openpyxl.load_workbook(file_name)
        # 循环每个sheet
        for ws_i, ws in enumerate(wb):
            data_map = {}
            for row in ws.rows:
                row_values = [col.value for col in row]
                data_map[row_values[args[i] - 1]] = row_values
            data_list.append(data_map)

    return data_list


def excel_match(*args):
    data_list = format_excel_data(args)
    print(data_list)

excel_match(1, 1)


# 保存文件
# wb_write.save('excel-merge' + get_current_time() + '.xlsx')
