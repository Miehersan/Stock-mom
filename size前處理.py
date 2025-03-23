import pandas as pd
import numpy as np
import os

os.chdir('D:\文件\SCU Big data\智能交易\自用')

df = pd.read_excel('市值.xlsx', header=1,dtype={'代號': str})

# 2. 篩選代號：長度為 4 且數值介於 1111~9999 之間
df_filtered = df[df['代號'].str.match(r'^[1-9]\d{3}$')]
    # 篩選代號為 1111~9999 的資料只保留4位數且不以0開頭的代號
df_filtered = df_filtered[df_filtered.iloc[:, 74].notna()]
    # 移除第74欄 2005/01為空值的代號

# 3. 定義 LOG 運算函數，若發生錯誤，回傳該欄位的中位數
def log_transform_with_median(series):
    series = pd.to_numeric(series, errors='coerce')  # 轉換數值，非數字變為 NaN
    log_series = series.apply(lambda x: np.log10(x * 1_000_000) if pd.notnull(x) and x > 0 else np.nan)
    median_value = log_series.median()  # 計算該欄位的中位數 (排除 NaN)
    log_series.fillna(median_value, inplace=True)
    return log_series

# 4. 逐欄進行 LOG 運算（排除不需要處理的欄位）
for col in df_filtered.columns:
    if col not in ['代號', '名稱']:
        df_filtered[col] = log_transform_with_median(df_filtered[col])

# 5. 複製 "size補值" 作為 "size_rank"，並將數值轉換為排名
df_rank = df_filtered.copy()
for col in df_rank.columns:
    if col not in ['代號', '名稱']:
        df_rank[col] = df_rank[col].rank(method='min', ascending=True)  # 轉換為排名

#  將整理好的資料匯出至新的 Excel 檔案
output_path = '整理後的資料.xlsx'
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df_filtered.to_excel(writer, index=False, sheet_name='size補值')
    df_rank.to_excel(writer, index=False, sheet_name='size_rank')

print("資料已成功整理並輸出到:", output_path)
