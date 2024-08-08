import os
import shutil
from openpyxl import Workbook
import re
import glob
def check_dir(dir, exp_basename='异常文件汇总'):
    dir = os.path.join(exp_basename, dir)
    if not os.path.exists(dir):
        os.makedirs(dir)
        return dir
    shutil.rmtree(dir)
    os.makedirs(dir)
    return dir
def check_file(file, sheet_name="放线数据汇总"):
    file = os.path.join(os.getcwd(), file)
    if os.path.exists(file):
        os.remove(file)
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    wb.save(file)
def find_match_files_recursion(parent_dir, re_pattern):
    file_set = set()
    file_list = []
    for root, _, files in os.walk(parent_dir):
        for file in files:
            if re.match(re_pattern, file) is not None and file not in file_set:
                file_set.add(file)
                file_list.append(os.path.join(root, file))
    return file_list
def find_match_txt_recursion(parent_dir, re_pattern):
    for root, _, files in os.walk(parent_dir):
        for file in files:
            if re.match(re_pattern, file) is not None:
                return os.path.join(root, file)
    return None
def ignore_hidden_files(src, names):
    """忽略隐藏文件和目录"""
    ignored_names = []
    for name in names:
        if name.startswith('~'):
            ignored_names.append(name)
    return set(ignored_names)

if __name__ == '__main__':
    path = r'F:\专题库\原数据\放线'
    project_name_pattern = r'\d{4}[放F]\d{2}[A-Z]{1}\d{3}.*\.(?i:txt)$'