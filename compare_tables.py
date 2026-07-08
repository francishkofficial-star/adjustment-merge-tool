import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')

df1 = pd.read_excel(r'C:\Users\Francis LIU\Downloads\调账底表_合并并添加备注_20260624.xlsx')
df2 = pd.read_excel(r'C:\Users\Francis LIU\Desktop\工作簿2.xlsx')

# 按商家ID分组求和
df1_grouped = df1.groupby('商家ID')['調賬費用'].sum().reset_index()
df2_grouped = df2.groupby('商家ID')['調賬費用'].sum().reset_index()

# 合并对比
merged = pd.merge(df1_grouped, df2_grouped, on='商家ID', how='outer', suffixes=('_表1', '_表2'))
merged = merged.fillna(0)
merged['差額'] = merged['調賬費用_表1'] - merged['調賬費用_表2']
merged['是否一致'] = merged['差額'].apply(lambda x: '一致' if abs(x) < 0.01 else '不一致')

# 统计
consistent = merged[merged['是否一致'] == '一致']
inconsistent = merged[merged['是否一致'] == '不一致']

print(f'总商家数: {len(merged)}')
print(f'金额一致: {len(consistent)}')
print(f'金额不一致: {len(inconsistent)}')
print()

if len(inconsistent) > 0:
    print('=== 不一致的商家 ===')
    print(inconsistent.to_string(index=False))
else:
    print('所有商家ID的調賬費用金额完全一致！')
