import re
from collections import Counter
import pandas as pd
import os

def count_words_in_file(file_path):
    """读取TXT文件并统计每个英文单词的出现频率，排除只有一个字母的单词。"""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read().lower()
            words = re.findall(r'\b[a-z]{2,}\b', content)
            word_count = Counter(words)
        return word_count
    except FileNotFoundError:
        print(f"文件 {file_path} 未找到。")
        return None
    except Exception as e:
        print(f"处理文件时发生错误：{e}")
        return None

def save_word_counts_to_excel(word_counts, output_file):
    """将单词计数结果保存到Excel文件中。"""
    df = pd.DataFrame(word_counts.items(), columns=['Word', 'Count'])
    df.sort_values(by='Count', ascending=False, inplace=True)
    df.to_excel(output_file, index=False)
    print(f"数据已保存到 {output_file}")

def process_folder(folder_path, output_folder):
    """处理指定文件夹内的所有TXT文件，每个文件生成一个Excel表格，并生成一个总的Excel表格。"""
    total_word_counts = Counter()
    try:
        for filename in os.listdir(folder_path):
            if filename.lower().endswith('.txt'):
                file_path = os.path.join(folder_path, filename)
                word_counts = count_words_in_file(file_path)
                if word_counts:
                    total_word_counts.update(word_counts)
                    output_excel_file = os.path.join(output_folder, f"{filename[:-4]}_word_counts.xlsx")
                    save_word_counts_to_excel(word_counts, output_excel_file)
    except FileNotFoundError:
        print(f"文件夹 {folder_path} 未找到。")
    except Exception as e:
        print(f"处理文件夹时发生错误：{e}")
    # 最后，保存总的Excel表格
    if total_word_counts:
        total_output_file = os.path.join(output_folder, 'total_word_counts.xlsx')
        save_word_counts_to_excel(total_word_counts, total_output_file)



# 使用示例
input_folder = 'txt_path'  # 指定TXT文件的文件夹路径
output_excel_folder = 'excel_path'  # 指定Excel输出的文件夹路径
process_folder(input_folder, output_excel_folder)
