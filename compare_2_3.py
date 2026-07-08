import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')

df1 = pd.read_excel(r'C:\Users\Francis LIU\Desktop\工作簿2.xlsx')
df2 = pd.read_excel(r'C:\Users\Francis LIU\Desktop\工作簿3.xlsx')

df1_grouped = df1.groupby('商家ID')['調賬費用'].sum().reset_index()
df2_grouped = df2.groupby('商家ID')['調賬費用'].sum().reset_index()

merged = pd.merge(df1_grouped, df2_grouped, on='商家ID', how='outer', suffixes=('_表2', '_表3'))
merged = merged.fillna(0)
merged['差額'] = merged['調賬費用_表2'] - merged['調賬費用_表3']
merged['是否一致'] = merged['差額'].apply(lambda x: '一致' if abs(x) < 0.01 else '不一致')

consistent = merged[merged['是否一致'] == '一致']
inconsistent = merged[merged['是否一致'] == '不一致']

print(f'总商家数: {len(merged)}')
print(f'金额一致: {len(consistent)}')
print(f'金额不一致: {len(inconsistent)}')
print()
print(f'表2行数: {len(df1)}')
print(f'表3行数: {len(df2)}')
print()
print(f'表2商家ID唯一数: {df1["商家ID"].nunique()}')
print(f'表3商家ID唯一数: {df2["商家ID"].nunique()}')
print()

if len(inconsistent) == 0:
    print('所有商家ID的調賬費用金额完全一致！')
else:
    print('=== 不一致的商家 ===')
    print(inconsistent.to_string(index=False))
