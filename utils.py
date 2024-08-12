import shutil
from openpyxl import Workbook
import re
from pathlib import Path
def check_dir(dir : Path, exp_basename=Path('异常文件汇总')) -> Path:
    dir = exp_basename.joinpath(dir)
    if not dir.exists():
        dir.mkdir(parents=True)
        return dir
    shutil.rmtree(dir)
    dir.mkdir(parents=True)
    return dir
def check_file(file, sheet_name="放线数据汇总"):
    if file.exists():
        file.unlink()  
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    wb.save(file)
def find_match_files_recursion(parent_dir : Path, re_pattern : str, suffix="*.doc*"):
    file_set = set()
    file_list = []
    files = parent_dir.rglob(suffix)
    for file in files:
        # print(file)
        if re.match(re_pattern, file.name) is not None and file not in file_set:
            file_set.add(file)
            file_list.append(file.resolve())
    return file_list
def find_match_txt_recursion(parent_dir : Path, re_pattern) -> Path:
    files = parent_dir.rglob("*.txt")
    for file in files:
            if re.match(re_pattern, file.name) is not None:
                return file.resolve()
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