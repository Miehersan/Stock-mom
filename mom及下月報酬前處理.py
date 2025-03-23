import pandas as pd
import numpy as np
import os

os.chdir('D:\文件\SCU Big data\智能交易\自用')

df = pd.read_excel('收盤價202201-202503.xlsx', header=1,dtype={'代號': str})

# 2. 篩選代號：長度為 4 且數值介於 1111~9999 之間
df_filtered = df[df['代號'].str.match(r'^[1-9]\d{3}$')]  # 只保留4位數且不以0開頭的代號
df_filtered = df_filtered[df_filtered.iloc[:, 74].notna()]
numeric_columns = df_filtered.columns[2:]
# 3. 定義比值 下個月月報酬 運算
def calc_ln_ratio(df_filtered):
    """
    計算 LN(後一欄 / 當前欄) 並將結果放在當前欄
    """
    df_ratio = df_filtered.copy()
    for col_index in range(2, len(df_filtered.columns) - 1):  # 從第 2 欄 (非 '代號', '名稱') 開始
        current_col = pd.to_numeric(df_filtered.iloc[:, col_index], errors='coerce')  # 當前欄
        next_col = pd.to_numeric(df_filtered.iloc[:, col_index + 1], errors='coerce')  # 後一欄
        df_ratio.iloc[:, col_index] = np.log(next_col / current_col)  # 計算 LN(後一欄 / 當前欄)
    df_ratio.iloc[:, -1] = np.nan  # 最後一欄無後續欄位，直接填 NaN
    return df_ratio

# 3. 定義計算mom
def calc_ln_12_ratio(df_filtered):
    results = []

    for index, row in df_filtered.iterrows():
        row_result = [row['代號'], row['名稱']]
    
        for i in range(11, len(numeric_columns)-1):
            current_value = row[numeric_columns[i]]
            previous_value = row[numeric_columns[i - 11]]  # 對應前一年同月的值
        
            if previous_value > 0 and current_value > 0:  # 避免分比為 0 的情況
                log_change = np.log(current_value / previous_value)
            else:
                log_change = np.nan  # 如果有問題的值，設定為 NaN
        
            row_result.append(log_change)
        results.append(row_result)
# 將結果轉成 DataFrame
    result_columns = ["代號", "名稱"] + list(numeric_columns[12:])
    result_df = pd.DataFrame(results, columns=result_columns)
    return result_df

# 4. 使用中位數補值
def fill_missing_with_median(df_filtered):
    """
    使用每一欄的中位數補 NaN
    """
    df_filled = df_filtered.copy()
    for col in df_filtered.columns:
        if col not in ['代號', '名稱']:
            median_value = df_filled[col].median(skipna=True)  # 計算中位數
            df_filled[col] = df_filled[col].fillna(median_value)  # 填補 NaN
    return df_filled

# 5. 排名處理
def rank_values(df_filtered):
    """
    對數據進行排名 (數值由小到大排序)
    """
    df_rank = df_filtered.copy()
    for col in df_filtered.columns:
        if col not in ['代號', '名稱']:
            df_rank[col] = df_filtered[col].rank(ascending=True, method='min')  # 由小到大排名
    return df_rank

# 6. 執行 LN 計算
df_ln_ratio = calc_ln_ratio(df_filtered)  # 計算 LN(後一欄 / 當前欄)
df_ln_12_ratio = calc_ln_12_ratio(df_filtered)  # 計算 LN(第N欄 / 第N-12欄)

# 7. 中位數補值
df_filled_ln_ratio = fill_missing_with_median(df_ln_ratio)
df_filled_ln_12_ratio = fill_missing_with_median(df_ln_12_ratio)

# 8. 排名處理
df_rank_ln_ratio = rank_values(df_filled_ln_ratio)
df_rank_ln_12_ratio = rank_values(df_filled_ln_12_ratio)

# 9. 匯出到 Excel
output_path = '整理後的資料.xlsx'
with pd.ExcelWriter(output_path) as writer:
    df_filled_ln_ratio.to_excel(writer, sheet_name='下個月月報酬', index=False)  # 轉出 "calc_ln_ratio"
    df_rank_ln_ratio.to_excel(writer, sheet_name='下個月月報酬_rank', index=False)  # 轉出 "calc_ln_ratio排序"
    df_filled_ln_12_ratio.to_excel(writer, sheet_name='mom補值', index=False)  # 轉出 "calc_ln_12_ratio"
    df_rank_ln_12_ratio.to_excel(writer, sheet_name='mom_rank', index=False)  # 轉出 "calc_ln_12_ratio排序"


print("資料已成功整理並輸出到:", output_path)
