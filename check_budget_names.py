import pandas as pd

# 读取预算模板
df = pd.read_excel(r'M:\下载\Q1Q2预算总表2026.xlsx')

# 输出到文件
with open(r'C:\Users\Francis LIU\Desktop\budget_names.txt', 'w', encoding='utf-8') as f:
    f.write("预算模板所有行（C列-三级类目, D列-供应商）:\n")
    for idx, row in df.iterrows():
        c = row.iloc[2] if len(row) > 2 else ''
        d = row.iloc[3] if len(row) > 3 else ''
        if pd.notna(c) and str(c).strip():
            f.write(f"Row {idx+1}: C={c}, D={d}\n")

print("Output written to budget_names.txt")
