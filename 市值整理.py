import pandas as pd
import numpy as np
import os

os.chdir('D:\文件\SCU Big data\智能交易\自用')

df = pd.read_excel('市值.xlsx', header=1,dtype={'代號': str})

df = pd.read_excel('收盤價.xlsx', header=1,dtype={'代號': str})
df_filtered = df[df['代號'].str.match(r'^[1-9]\d{3}$')]  # 只保留4位數且不以0開頭的代號
df_filtered.to_excel('篩選收盤價.xlsx')