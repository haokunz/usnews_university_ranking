import numpy as np
import pandas as pd

# 读取Excel文件中的"合并2"工作表
file_path = "//大学排名计算/USNEWS_US_UNI__2023_fixed_log_number_normalized.xlsx"
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

# 计算30-53位National University的offer重复次数均值
mean_30_53_national = df[(df['排名'] >= 30) & (df['排名'] <= 53) & (df['学校类型'] == 'national-universities')]['offer_重复次数'].mean()

# 计算20-53位公立大学的offer重复次数均值
mean_20_53_public = df[(df['排名'] >= 20) & (df['排名'] <= 53) & (df['学校类型'] == 'national-universities')]['offer_重复次数'].mean()

# 替代相应排名的学校offer重复次数，并取整
df['Adjusted_录取人数_30_53'] = df.apply(
    lambda row: int(mean_30_53_national) if row['学校类型'] == 'national-universities' and 30 <= row['排名'] <= 53 and row['offer_重复次数'] < 100 else int(row['offer_重复次数']),
    axis=1
)
df['Adjusted_录取人数_20_53'] = df.apply(
    lambda row: int(mean_20_53_public) if row['学校类型'] == 'national-universities' and 20 <= row['排名'] <= 53 and row['offer_重复次数'] < 100 else int(row['offer_重复次数']),
    axis=1
)

# 增加列：原始数据
df['Original_录取人数'] = df['offer_重复次数']

# 调整后的录取人数列
df['录取人数_均值_30_53'] = df['Adjusted_录取人数_30_53']
df['录取人数_均值_20_53'] = df['Adjusted_录取人数_20_53']

# 归一化函数（线性归一）
def linear_normalize(series):
    return (series - series.min()) / (series.max() - series.min())

# 归一化函数（对数归一）
def log_normalize(series):
    return np.log1p(series - series.min()) / np.log1p(series.max() - series.min())

# 设定权重组合
weights = np.array([0.5, 0.25, 0.25])

# 定义计算赋分的函数
def calculate_scores(df, score_col_name, rank_norm_func, offer_norm_func, acceptance_norm_func):
    # 归一化
    df['Rank_Norm'] = rank_norm_func(df['排名'])
    df['Offer_Norm'] = offer_norm_func(df['offer_重复次数'])
    df['Acceptance_Rate_Norm'] = acceptance_norm_func(df['Transformed_Acceptance_Rate'])
    
    # 校正offer重复次数
    df['Adjusted_Offer_Repeats'] = df['offer_重复次数'] * (df['Rank_Norm'])**2 * (df['Acceptance_Rate_Norm'])**2

    # 归一化校正后的 offer 重复次数
    df['Adjusted_Offer_Norm'] = offer_norm_func(df['Adjusted_Offer_Repeats'])

    df[score_col_name + '_Score'] = df.apply(lambda row: np.dot(
        np.array([row['Rank_Norm'], row['Adjusted_Offer_Norm'], row['Acceptance_Rate_Norm']]), weights), axis=1)

    df.loc[df['学校类型'] == 'national-liberal-arts-colleges', score_col_name + '_Score'] /= 2

    max_score = df[score_col_name + '_Score'].max()
    min_score = df[score_col_name + '_Score'].min()
    df[score_col_name + '_赋分'] = 500 + ((df[score_col_name + '_Score'] - min_score) / (max_score - min_score)) * (10000 - 500)

    df[score_col_name + '_赋分'] = df[score_col_name + '_赋分'].round().astype(int)

    return df

# 计算原始录取人数情况下的赋分
df_original_linear_log = calculate_scores(df.copy(), 'Original_Linear_Log', linear_normalize, log_normalize, log_normalize)
df_original_log_log = calculate_scores(df.copy(), 'Original_Log_Log', log_normalize, log_normalize, log_normalize)

# 计算均值_30_53情况下的赋分
df_mean_30_53_linear_log = calculate_scores(df.copy(), '均值_30_53_Linear_Log', linear_normalize, log_normalize, log_normalize)
df_mean_30_53_log_log = calculate_scores(df.copy(), '均值_30_53_Log_Log', log_normalize, log_normalize, log_normalize)

# 计算均值_20_53情况下的赋分
df_mean_20_53_linear_log = calculate_scores(df.copy(), '均值_20_53_Linear_Log', linear_normalize, log_normalize, log_normalize)
df_mean_20_53_log_log = calculate_scores(df.copy(), '均值_20_53_Log_Log', log_normalize, log_normalize, log_normalize)

# 保存最终结果到Excel
with pd.ExcelWriter("//大学排名计算/final_ranking_with_multiple_sheets.xlsx") as writer:
    df_original_linear_log.to_excel(writer, sheet_name='Original_Linear_Log', index=False)
    df_original_log_log.to_excel(writer, sheet_name='Original_Log_Log', index=False)
    df_mean_30_53_linear_log.to_excel(writer, sheet_name='均值_30_53_Linear_Log', index=False)
    df_mean_30_53_log_log.to_excel(writer, sheet_name='均值_30_53_Log_Log', index=False)
    df_mean_20_53_linear_log.to_excel(writer, sheet_name='均值_20_53_Linear_Log', index=False)
    df_mean_20_53_log_log.to_excel(writer, sheet_name='均值_20_53_Log_Log', index=False)
