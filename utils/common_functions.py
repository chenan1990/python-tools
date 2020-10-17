"""
 多个excel，多个sheet数据合并
 生成的文件名为excel-merge...xlsx
"""
import os
import time


# 获取当前时间
def get_current_time():
    return time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime())



def get_file_path(path, ignore_file):
    # 文件列表
    file_list = []
    # 获取该目录下所有的文件名称和目录名称
    dir_or_files = os.listdir(path)
    for dir_file in dir_or_files:
        # 获取目录或者文件的路径
        dir_file_path = os.path.join(path, dir_file)
        # 该路径是文件并且文件类型为xlsx
        if not os.path.isdir(dir_file_path) and dir_file_path.endswith('.xlsx') and ignore_file not in dir_file_path:
            file_list.append(dir_file_path)
    return file_list
