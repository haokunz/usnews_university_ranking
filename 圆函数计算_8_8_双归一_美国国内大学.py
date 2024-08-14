import numpy as np
import pandas as pd

# 读取Excel文件中的工作表
file_path = "///大学排名计算/USNEWS_US_UNI__2023_fixed_log_number_normalized.xlsx"
df = pd.read_excel(file_path)

# 填充缺失值
df['录取率'] = df['录取率'].fillna(df['录取率'].mean())
df['offer_重复次数'] = df['offer_重复次数'].fillna(df['offer_重复次数'].mean())

# 录取率转换为选择难度的指标（反向处理）
df['Transformed_Acceptance_Rate'] = 1 - df['录取率']

# 确保所有排名列都是数值类型
df['排名'] = pd.to_numeric(df['排名'], errors='coerce')

# 处理数据类型转换后的NA值
df['排名'] = df['排名'].fillna(df['排名'].mean())

# 计算31-50位National University的offer重复次数均值
mean_31_50_national = df[(df['排名'] >= 31) & (df['排名'] <= 50) & (df['学校类型'] == 'national-universities')]['offer_重复次数'].mean()

# 替代相应排名的学校offer重复次数，并取整
df['Adjusted_录取人数_31_50'] = df.apply(
    lambda row: int(mean_31_50_national) if row['学校类型'] == 'national-universities' and 31 <= row['排名'] <= 50 and row['offer_重复次数'] < 100 else int(row['offer_重复次数']),
    axis=1
)

# 将德克萨斯大学奥斯丁分校的录取人数调整为录取人数_均值_31_50的一半
df.loc[df['学校名称'] == '德克萨斯大学奥斯汀分校', 'Adjusted_录取人数_31_50'] = int(mean_31_50_national / 2)

# 增加列：原始数据
df['Original_录取人数'] = df['offer_重复次数']

# 调整后的录取人数列
df['录取人数_均值_31_50'] = df['Adjusted_录取人数_31_50']

# 归一化函数（线性归一并反向处理）
def linear_normalize(series, reverse=False):
    normalized = (series - series.min()) / (series.max() - series.min())
    if reverse:
        normalized = 1 - normalized
    return normalized

# 归一化函数（对数归一并反向处理）
def log_normalize(series, reverse=False):
    normalized = np.log1p(series - series.min()) / np.log1p(series.max() - series.min())
    if reverse:
        normalized = 1 - normalized
    return normalized

# 设定权重组合
weights = np.array([0.5, 0.25, 0.25])

# 定义计算赋分的函数
def calculate_scores(df, score_col_name, rank_norm_func, offer_norm_func, acceptance_norm_func, offer_col_name, liberal_arts_coefficient):
    # 对排名进行归一化
    df['Rank_Norm'] = rank_norm_func(df['排名'], reverse=True)
    # 对 offer 重复次数进行归一化
    df['Offer_Norm'] = offer_norm_func(df[offer_col_name], reverse=True)
    # 对录取率进行归一化
    df['Acceptance_Rate_Norm'] = acceptance_norm_func(df['Transformed_Acceptance_Rate'], reverse=False)
    
    # 校正 offer 重复次数
    df['Adjusted_Offer_Repeats'] = df[offer_col_name] * (df['Rank_Norm'])**2 * (df['Acceptance_Rate_Norm'])**2

    # 对校正后的 offer 重复次数进行归一化
    df['Adjusted_Offer_Norm'] = offer_norm_func(df['Adjusted_Offer_Repeats'], reverse=False)

    df[score_col_name + '_Score'] = df.apply(lambda row: np.dot(
        np.array([row['Rank_Norm'], row['Adjusted_Offer_Norm'], row['Acceptance_Rate_Norm']]), weights), axis=1)

    df.loc[df['学校类型'] == 'national-liberal-arts-colleges', score_col_name + '_Score'] *= liberal_arts_coefficient

    # 获取普林斯顿大学、加州大学戴维斯分校和罗格斯大学新布朗斯维克(主校区)的得分
    princeton_score = df.loc[df['学校名称'] == '普林斯顿大学', score_col_name + '_Score'].values[0]
    davis_score = df.loc[df['学校名称'] == '加州大学戴维斯分校', score_col_name + '_Score'].values[0]
    rutgers_score = df.loc[df['学校名称'] == '罗格斯大学新布朗斯维克(主校区)', score_col_name + '_Score'].values[0]

    # 确保普林斯顿大学得分最大，加州大学戴维斯分校得分最小，罗格斯大学得分为100
    max_score = princeton_score
    min_score = rutgers_score
    
    # 重新定义赋分
    df[score_col_name + '_赋分'] = 100 + ((df[score_col_name + '_Score'] - min_score) / (max_score - min_score)) * (10000 - 100)
    
    # 确保赋分的最低分和最高分
    df[score_col_name + '_赋分'] = df[score_col_name + '_赋分'].clip(lower=100, upper=10000)
    
    df[score_col_name + '_赋分'] = df[score_col_name + '_赋分'].round().astype(int)
    
    return df

# 计算原始录取人数情况下的赋分
df_original_linear_log_06 = calculate_scores(df.copy(), 'Original_Linear_Log_06', linear_normalize, log_normalize, log_normalize, 'offer_重复次数', 0.6)
df_original_log_log_06 = calculate_scores(df.copy(), 'Original_Log_Log_06', log_normalize, log_normalize, log_normalize, 'offer_重复次数', 0.6)
df_original_linear_log_065 = calculate_scores(df.copy(), 'Original_Linear_Log_065', linear_normalize, log_normalize, log_normalize, 'offer_重复次数', 0.65)
df_original_log_log_065 = calculate_scores(df.copy(), 'Original_Log_Log_065', log_normalize, log_normalize, log_normalize, 'offer_重复次数', 0.65)

# 计算均值_31_50情况下的赋分
df_mean_31_50_linear_log_06 = calculate_scores(df.copy(), '均值_31_50_Linear_Log_06', linear_normalize, log_normalize, log_normalize, '录取人数_均值_31_50', 0.6)
df_mean_31_50_log_log_06 = calculate_scores(df.copy(), '均值_31_50_Log_Log_06', log_normalize, log_normalize, log_normalize, '录取人数_均值_31_50', 0.6)
df_mean_31_50_linear_log_065 = calculate_scores(df.copy(), '均值_31_50_Linear_Log_065', linear_normalize, log_normalize, log_normalize, '录取人数_均值_31_50', 0.65)
df_mean_31_50_log_log_065 = calculate_scores(df.copy(), '均值_31_50_Log_Log_065', log_normalize, log_normalize, log_normalize, '录取人数_均值_31_50', 0.65)

output_file_path = "///final_ranking_with_two_normalization_3.xlsx"
with pd.ExcelWriter(output_file_path) as writer:
    df_original_linear_log_06.to_excel(writer, sheet_name='Original_Linear_Log_06', index=False)
    df_original_log_log_06.to_excel(writer, sheet_name='Original_Log_Log_06', index=False)
    df_original_linear_log_065.to_excel(writer, sheet_name='Original_Linear_Log_065', index=False)
    df_original_log_log_065.to_excel(writer, sheet_name='Original_Log_Log_065', index=False)
    df_mean_31_50_linear_log_06.to_excel(writer, sheet_name='均值_31_50_Linear_Log_06', index=False)
    df_mean_31_50_log_log_06.to_excel(writer, sheet_name='均值_31_50_Log_log_06', index=False)
    df_mean_31_50_linear_log_065.to_excel(writer, sheet_name='均值_31_50_Linear_log_065', index=False)
    df_mean_31_50_log_log_065.to_excel(writer, sheet_name='均值_31_50_Log_log_065', index=False)
