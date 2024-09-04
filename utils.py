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
        print(file)
        if re.match(re_pattern, file.name) is not None and file.name not in file_set:
            file_set.add(file.name)
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
def validate_building_high(construct_project, buildings_high) -> str:
    if buildings_high is None:
        return None
    # 匹配数字
    number_pattern = re.compile(r"\d+\.\d*(?!#)")
    # 类似于A-1#的编码
    normal_pattern = re.compile(r"[A-Z][-]?[A-Z]?\d*")
    ## 类似于1#的编码,要注意排除A-1#的情况
    number_sign = re.compile(r"(?<![A-Z])\d+#")
    building_list = buildings_high.split()
    if len(building_list) == 1:
        match_result = number_pattern.search(string=building_list[0])
        if match_result is not None and match_result != '' :
            return match_result.group()
        else :
            return None
    for building_high in building_list:
        if construct_project in building_high:
            if number_pattern.search(building_high) is not None:
                return number_pattern.search(building_high).group()
            else :
                return None
        elif normal_pattern.search(construct_project) is not None:
            self_code = normal_pattern.search(construct_project).group()
            if self_code in building_high:
                if number_pattern.search(building_high) is not None:
                    return number_pattern.search(building_high).group()
                else :
                    return None
        elif number_sign.search(construct_project) is not None:
            self_code = number_sign.search(construct_project).group()
            if self_code in building_high:
                if number_pattern.search(building_high) is not None:
                    return number_pattern.search(building_high).group()
                else :
                    return None
    return None
if __name__ == '__main__':
    path = Path(r"F:\专题库\原数据\放线txt补充")
    dest = Path(r"F:\专题库\原数据\放线补充")
    if not dest.exists():
        dest.mkdir(parents=True, exist_ok=True)
    files = path.glob("*.txt")
    for file in files:
        print(file.name.split(".")[0])
        dest.joinpath(file.name.split(".")[0]).mkdir(parents=True, exist_ok=True)
        shutil.copy(file, dest.joinpath(file.name.split(".")[0]).joinpath(file.name))
