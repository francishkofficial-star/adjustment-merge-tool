import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')

# 读取RealtimeInventory文件
inv_path = r'C:\Users\Francis LIU\Downloads\RealtimeInventory-1783414061359.xlsx'
df_inv = pd.read_excel(inv_path)

# 按商品编码分组，汇总可用库存数量
grouped = df_inv.groupby('商品编码')['可用库存数量'].sum().reset_index()

print("所有商品编码及其可用库存汇总:")
for _, row in grouped.iterrows():
    print(f"  {row['商品编码']}: {row['可用库存数量']}")
