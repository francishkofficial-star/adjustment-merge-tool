import pandas as pd
import sys

# 读取预算模板
df = pd.read_excel(r'M:\下载\Q1Q2预算总表2026.xlsx')

# 输出到文件
with open(r'C:\Users\Francis LIU\Desktop\budget_output.txt', 'w', encoding='utf-8') as f:
    f.write("Budget template first 40 rows:\n")
    f.write(df.iloc[:40, [2,3,4]].to_string())
    f.write("\n\nColumns: " + str(df.columns.tolist()) + "\n")
    
    # 输出所有行数据
    f.write("\n\nAll rows (C, D, E columns):\n")
    for idx, row in df.iterrows():
        c = row.iloc[2] if len(row) > 2 else ''
        d = row.iloc[3] if len(row) > 3 else ''
        e = row.iloc[4] if len(row) > 4 else ''
        f.write(f"Row {idx+1}: C={c}, D={d}, E={e}\n")

print("Output written to budget_output.txt")
