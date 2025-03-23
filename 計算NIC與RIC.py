import pandas as pd
import numpy as np
import os


os.chdir('D:\文件\SCU Big data\智能交易\自用')

size_rank = pd.read_excel('size.xlsx', sheet_name='size_rank', dtype={'代號': str})
size_filled = pd.read_excel('size.xlsx', sheet_name='size補值', dtype={'代號': str})
bm_rank = pd.read_excel('bm.xlsx', sheet_name='bm_rank', dtype={'代號': str})
bm_filled = pd.read_excel('bm.xlsx', sheet_name='bm補值', dtype={'代號': str})
mom_filled = pd.read_excel('mom.xlsx', sheet_name='mom補值', dtype={'代號': str})
mom_rank = pd.read_excel('mom.xlsx', sheet_name='mom_rank', dtype={'代號': str})
next_month_rank = pd.read_excel('下個月月報酬.xlsx', sheet_name='下個月月報酬_rank', dtype={'代號': str})
next_month = pd.read_excel('下個月月報酬.xlsx', sheet_name='下個月月報酬', dtype={'代號': str})

 #  定義函數 - 計算 RIC 和 NIC
def calculate_correlation(data1, data2, start_index=2, start_index_2=None):
    """
    data1, data2: 兩個 DataFrame 進行相關係數計算
    start_index: 主要控制 data1 從第幾欄開始計算
    start_index_2: 控制 data2 從第幾欄開始計算，若為 None，則與 data1 保持一致
    """
    start_index_2 = start_index if start_index_2 is None else start_index_2
    results = {'欄位': [], '相關係數': []}
    
    for col1, col2 in zip(data1.columns[start_index:], data2.columns[start_index_2:]):
        if col1 in data2.columns:
            col1_values = pd.to_numeric(data1[col1], errors='coerce')
            col2_values = pd.to_numeric(data2[col2], errors='coerce')
            corr = col1_values.corr(col2_values)
            results['欄位'].append(col1)
            results['相關係數'].append(corr)
    
    return pd.DataFrame(results)


# 3. 計算 RIC 和 NIC
# 3.1 size.xlsx
size_ric = calculate_correlation(size_rank, next_month_rank)
size_nic = calculate_correlation(size_filled, next_month)

# 3.2 bm.xlsx
bm_ric = calculate_correlation(bm_rank, next_month_rank)
bm_nic = calculate_correlation(bm_filled, next_month)

# 3.3 mom.xlsx (資料比較少，所以起始欄位不同)
mom_ric = calculate_correlation(mom_rank, next_month_rank, start_index=2, start_index_2=14)
mom_nic = calculate_correlation(mom_filled, next_month, start_index=2, start_index_2=14)


# 4. 儲存計算結果回到 Excel 檔案
# 4.1 size.xlsx - 儲存 RIC 和 NIC
with pd.ExcelWriter('size.xlsx', engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:
    size_ric.T.to_excel(writer, index=False,header=False, sheet_name='RIC')
    size_nic.T.to_excel(writer, index=False,header=False, sheet_name='NIC')

# 4.2 bm.xlsx - 儲存 RIC 和 NIC
with pd.ExcelWriter('bm.xlsx', engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:
    bm_ric.T.to_excel(writer, index=False,header=False, sheet_name='RIC')
    bm_nic.T.to_excel(writer, index=False,header=False, sheet_name='NIC')

# 4.3 mom.xlsx - 儲存 RIC 和 NIC
with pd.ExcelWriter('mom.xlsx', engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:
    mom_ric.T.to_excel(writer, index=False,header=False, sheet_name='RIC')
    mom_nic.T.to_excel(writer, index=False,header=False, sheet_name='NIC')

print("相關係數計算完成，RIC 和 NIC 已儲存至 size.xlsx、bm.xlsx 和 mom.xlsx。")