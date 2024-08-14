import numpy as np
import pandas as pd

# 读取Excel文件中的第一个工作表
file_path = "//学校排名数据/数据/大学排名/原件/USNEWS/College_Global_ranking_US_NEWS.xlsx"
df = pd.read_excel(file_path, sheet_name=0)

# 只取US News排名前30的大学
df_top30 = df[df['排名'] <= 30]

# 填充录取率中的NA值
df_top30['所有校区的录取率'] = df_top30['所有校区的录取率'].fillna(df_top30['所有校区的录取率'].mean())
df_top30['录取率'] = df_top30['录取率'].fillna(df_top30['录取率'].mean())

# 录取率转换为表示选择难度的指标
df_top30['Transformed_Acceptance_Rate'] = 1 - df_top30['录取率']

# 确保录取率列是数值类型
df_top30['Transformed_Acceptance_Rate'] = pd.to_numeric(df_top30['Transformed_Acceptance_Rate'], errors='coerce')

# 对数归一化录取率
def log_normalize(series):
    return np.log(series + 1)

df_top30['Acceptance_Rate_LogNorm'] = log_normalize(df_top30['Transformed_Acceptance_Rate'])

# 计算综合得分（只考虑录取率）
df_top30['Score'] = df_top30['Acceptance_Rate_LogNorm']

# 按得分排序
final_ranking_acceptance_rate_only = df_top30.sort_values(by='Score').reset_index(drop=True)

#print(final_ranking_acceptance_rate_only)

# 保存最终结果到Excel
output_file_path_acceptance_rate_only = "///学校排名数据/代码/大学排名计算/final_ranking_acceptance_rate_only.xlsx"
final_ranking_acceptance_rate_only.to_excel(output_file_path_acceptance_rate_only, index=False)

final_ranking_acceptance_rate_only[['学校名称', 'Score']]
