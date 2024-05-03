import os
import re

def delete_cfg_files(folder_path):
    """删除文件夹中所有的.cfg文件。"""
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.cfg'):
            os.remove(os.path.join(folder_path, filename))
            print(f"Deleted '{filename}'")

def chinese_to_arabic(chinese_num):
    """将中文数字转换为阿拉伯数字。"""
    chinese_num_map = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '七': 7, '八': 8, '九': 9, '十': 10,
                       '百': 100, '千': 1000, '万': 10000, '亿': 100000000}
    result = 0
    temp_num = 0
    for char in chinese_num:
        if char in chinese_num_map:
            temp_num += chinese_num_map[char]
        else:
            result += temp_num * chinese_num_map[char]
            temp_num = 0
    result += temp_num
    return result

def extract_and_rename_files(folder_path):
    """遍历文件夹并重命名其中的文件，新文件名由提取的数字组成（不包括6和六），若原名包含'答案'则在新名后加'a'。"""
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        base, extension = os.path.splitext(filename)
        # 检测文件名中是否包含'答案'
        has_answer = '答案' in base
        # 使用正则表达式匹配阿拉伯数字和中文数字，排除6和六
        numbers = re.findall(r'[0-57-9]|[一二三四五七八九十百千万亿]', base)
        if not numbers and not has_answer:
            continue  # 如果没有数字且不包含'答案'则跳过
        # 将中文数字转换为阿拉伯数字
        numbers = [str(chinese_to_arabic(num) if num.isdigit() == False else int(num)) for num in numbers]
        # 拼接新的文件名
        new_filename = ''.join(numbers) + ('a' if has_answer else '') + extension
        new_file_path = os.path.join(folder_path, new_filename)
        # 重命名文件
        os.rename(file_path, new_file_path)
        print(f"Renamed '{filename}' to '{new_filename}'")

# 使用示例
folder_path = 'source_folder_path'  # 替换为源文件夹路径
delete_cfg_files(folder_path)  # 首先删除所有的.cfg文件
extract_and_rename_files(folder_path)  # 然后进行文件重命名
