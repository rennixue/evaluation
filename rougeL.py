import lawrouge
import pandas as pd

def rouge_similarity(str1, str2):
    '''
    用rougeL计算两个字符串的相似度
    '''
    if pd.isna(str1) or pd.isna(str2):
        return 0
    str1 = str(str1).lower().replace(" ", "")
    str2 = str(str2).lower().replace(" ", "")
    rouge = lawrouge.Rouge()
    scores = rouge.get_scores([str1], [str2], avg=2)
    return scores['f']

df = pd.read_excel('simp_result_6may_1106.xlsx', sheet_name='Sheet1')
df['similarity'] = df.apply(lambda row: rouge_similarity(row['course_code'], row['couse_code_alg']), axis=1)
df.to_excel('simp_result_similarity.xlsx', index=False)