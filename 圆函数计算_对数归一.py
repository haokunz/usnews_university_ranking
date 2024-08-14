import numpy as np
import pandas as pd

# 读取Excel文件中的"合并1"工作表
file_path = "//大学排名/翻译件/ranking_Translation.xlsx"
df = pd.read_excel(file_path, sheet_name='合并1')

# 只取US News排名前104的大学
df_top106 = df[df['USNEWS排名'] <= 106]

# 填充录取率中的NA值
df_top106['录取率'] = df_top106['录取率'].fillna(df_top106['录取率'].mean())
df_top106['国际生录取率'] = df_top106['国际生录取率'].fillna(df_top106['国际生录取率'].mean())

# 录取率转换为表示选择难度的指标
df_top106['Transformed_Acceptance_Rate'] = 1 - df_top106['录取率']

# 确保所有排名列都是数值类型
df_top106['USNEWS排名'] = pd.to_numeric(df_top106['USNEWS排名'], errors='coerce')
df_top106['2024_QS_World_University_Rankings'] = pd.to_numeric(df_top106['2024_QS_World_University_Rankings'], errors='coerce')
df_top106['TIMES排名'] = pd.to_numeric(df_top106['TIMES排名'], errors='coerce')
df_top106['软科排名'] = pd.to_numeric(df_top106['软科排名'], errors='coerce')

# 处理数据类型转换后的NA值
df_top106['USNEWS排名'] = df_top106['USNEWS排名'].fillna(df_top106['USNEWS排名'].mean())
df_top106['2024_QS_World_University_Rankings'] = df_top106['2024_QS_World_University_Rankings'].fillna(df_top106['2024_QS_World_University_Rankings'].mean())
df_top106['TIMES排名'] = df_top106['TIMES排名'].fillna(df_top106['TIMES排名'].mean())
df_top106['软科排名'] = df_top106['软科排名'].fillna(df_top106['软科排名'].mean())

# 生成随机的录取人数，录取率越低的学校，录取人数越少
np.random.seed(42)  # 固定随机种子以便于复现
df_top106['录取人数'] = (1 - df_top106['录取率']) * np.random.randint(1, 3001, size=len(df_top106))

# 反向归一化录取人数，使得录取人数最多的归一化后最小
df_top106['Normalized_Enrollment'] = 1 - (df_top106['录取人数'] - df_top106['录取人数'].min()) / (df_top106['录取人数'].max() - df_top106['录取人数'].min())

# 对数归一化函数
def log_normalize(series):
    return np.log(series + 1)

# 对数归一化所有排名和转换后的录取率
df_top106['USNews_Rank_LogNorm'] = log_normalize(df_top106['USNEWS排名'])
df_top106['QS_Rank_LogNorm'] = log_normalize(df_top106['2024_QS_World_University_Rankings'])
df_top106['TIMES_Rank_LogNorm'] = log_normalize(df_top106['TIMES排名'])
df_top106['RUANKE_Rank_LogNorm'] = log_normalize(df_top106['软科排名'])
df_top106['Acceptance_Rate_LogNorm'] = log_normalize(df_top106['Transformed_Acceptance_Rate'])
df_top106['Enrollment_LogNorm'] = df_top106['Normalized_Enrollment']

# 设定权重：4个排名各12.5%，录取率25%，录取人数25%
weights = np.array([0.125, 0.125, 0.125, 0.125, 0.25, 0.25])

# 应用圆函数计算得分
def calculate_weighted_score_log(row):
    n = 6  # 指标的数量
    theta = np.linspace(0, 2 * np.pi, n, endpoint=False)
    r = np.array([
        row['USNews_Rank_LogNorm'],
        row['QS_Rank_LogNorm'],
        row['TIMES_Rank_LogNorm'],
        row['RUANKE_Rank_LogNorm'],
        row['Acceptance_Rate_LogNorm'],
        row['Enrollment_LogNorm']
    ])
    x = r * np.cos(theta)
    y = r * np.sin(theta)
    distances = np.sqrt(x**2 + y**2)
    weighted_distances = distances * weights
    score = 1 / weighted_distances.sum()  # 分数越大越好
    return score

# 计算每个大学的得分
df_top106['Score'] = df_top106.apply(calculate_weighted_score_log, axis=1)

# 计算赋分
max_score = df_top106['Score'].max()
df_top106['赋分'] = df_top106['Score'] * (10000 / max_score)

# 按得分排序（分数越高排名越高）
final_ranking = df_top106.sort_values(by='Score', ascending=False).reset_index(drop=True)

# 保存最终结果到Excel
output_file_path = "//大学排名计算/final_ranking_log_normalized.xlsx"
final_ranking.to_excel(output_file_path, index=False)

# 显示前几所大学的排名和得分
final_ranking_preview = final_ranking[['学校名称', 'Score', '赋分', '录取人数', 'Normalized_Enrollment']].head(10)
final_ranking_preview
