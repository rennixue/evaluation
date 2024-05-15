import pandas as pd
import numpy as np
import requests
from PyPDF2 import PdfReader
from docx import Document
from pptx import Presentation
import re
import os

    
def contains_number(s):
    if re.search(r'\d', s):
        return True
    return False

def find_split_code(chars, mixed_string):
    # 使用正则表达式匹配数字和非数字字符：当数字和字母都出现时才判断是TRUE
    digits = re.findall(r'\d+', mixed_string)
    letters = re.findall(r'\D+', mixed_string)
    letters_lower = [x.lower() for x in letters]
    letters_upper = [x.upper() for x in letters]
    res = digits+letters_lower+letters_upper
    res = [x.strip() for x in res]
    ans = 0 # 记录找到的字符个数
    for item in res:
        for char in chars:
            if item in char:
                ans += 1
                break
    return ans==len(res)

def sliding_window(chars, window_size, course_code):
    # 滑动窗口寻找
    ans = False 
    # print(course_code,chars)
    if course_code in chars:
        return True
    if course_code.replace(" ", "") in chars:
        return True
    for i in range(len(chars) - window_size + 1):
        ans = find_split_code(chars[i:i + window_size], course_code)
        if ans:
            break
    return ans

def check_nan(course_code, summary):
    if summary:
        chars = summary.split()
        return sliding_window(chars, 10, course_code)
    else:
        return None

def search_file_by_prefix(folder_path, file_prefix):
    # 列出文件夹中的所有文件
    for filename in os.listdir(folder_path):
        # 检查文件名是否以给定的前缀开头
        if filename.startswith(file_prefix):
            # 如果是，打印文件名和内容
            # print(f"找到文件: {filename}")
            # 读取文件内容
            with open(os.path.join(folder_path, filename), 'r', encoding='utf-8') as file:
                content = file.read()
                # print(content)
            return content
    # 如果文件没有找到，打印一条消息
    print(f"没有找到以{file_prefix}开头的文件。")
    # temp.append(file_prefix)
    return None

##############################################################################################
# 存放txt的文件夹路径
folder_path = './text'

# 读取文件
df = pd.read_excel('PUB_course_info_by_summary_and_agg_method_15may_1410.xlsx', sheet_name='output')
df_real = pd.read_excel('PUB_course_info_by_summary_and_agg_method_15may_1410.xlsx', sheet_name='real')
df_raw = pd.read_excel('PUB_course_info_by_summary_and_agg_method_15may_1410.xlsx', sheet_name='raw')

# 筛选相似度在0.9以下，且人工不为0的数据
select_df1 = df[(df['similarity_course_code']<0.9) & (df['course_code_org']!=0) & (pd.isna(df['course_code_org'])==False)]
# 根据order_id筛选原始数据，查找人工code是否出现在txt里面，打标
select_file_df = df_raw[df_raw['order_id'].isin(select_df1['order_id'])]
df_raw['course_code_valid'] = np.nan
for index, row in select_file_df.iterrows():
    # print(index)
    order_id = row['order_id']
    file_id = row['file_id']
    course_code = select_df1.loc[select_df1['order_id']==order_id, 'course_code_org'].iloc[0]
    content = search_file_by_prefix(folder_path, str(order_id)+'-'+str(file_id))
    df_raw.loc[index, 'course_code_valid'] = check_nan(course_code, content)
    
# 汇总raw的打标到real里打标
df_raw['file_id_split'] = df_raw.apply(lambda row: str(row['file_id']).split('_')[0], axis=1)
select_file_df_real = df_real[df_real['order_id'].isin(select_df1['order_id'])]
df_real['course_code_valid'] = np.nan
for index, row in select_file_df_real.iterrows():
    file_id = row['file_id']
    file_df = df_raw[df_raw['file_id']==file_id]
    if file_df['course_code_valid'].sum() > 0:
        df_real.loc[index, 'course_code_valid'] = True
    else:
        df_real.loc[index, 'course_code_valid'] = False
        
# 汇总real的打标到output里打标
df['org_code_valid'] = np.nan
for index, row in select_df1.iterrows():
    order_id = row['order_id']
    file_df = select_file_df_real[select_file_df_real['order_id']==order_id]
    if file_df['course_code_valid'].sum() > 0:
        df.loc[index, 'org_code_valid'] = True
    else:
        df.loc[index, 'org_code_valid'] = False
        
# 保存文件
writer = pd.ExcelWriter('result.xlsx', engine='xlsxwriter')
df_raw.to_excel(writer, sheet_name='raw', index=False, na_rep='<NA>')
df_real.to_excel(writer, sheet_name='real', index=False, na_rep='<NA>')
df.to_excel(writer, sheet_name='output', index=False, na_rep='<NA>')
writer.save()