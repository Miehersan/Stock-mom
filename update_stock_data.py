import pandas as pd
import numpy as np
import json

# 讀取 Excel 檔案
file_path = "收盤價202201-KEEP.xlsx"
df = pd.read_excel(file_path, sheet_name="Sheet")

# 轉置數據，使日期成為索引，股票代號為列
df_t = df.set_index(['代號', '名稱']).T

# 轉換索引為 datetime 格式
df_t.index = pd.to_datetime(df_t.index, format="%Y/%m")

# 計算 12 個月的對數收益率
log_return_12m = np.log(df_t / df_t.shift(12))

# 轉置回原格式
log_return_12m = log_return_12m.T.reset_index()

# 找到最新日期的數據
latest_date = log_return_12m.columns[-1]

# 篩選 12 個月對數收益率最高的前 10 檔股票
top_10_stocks = log_return_12m.nlargest(10, latest_date)[['代號', '名稱', latest_date]]

# 存成 JSON 格式（方便 GitHub Pages 使用）
data = top_10_stocks.to_dict(orient="records")

with open("stock_data.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=4)

print("✅ 最新 12 個月對數收益率已更新 stock_data.json！")
