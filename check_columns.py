import pandas as pd

# 读取预算模板
df = pd.read_excel(r'M:\下载\Q1Q2预算总表2026.xlsx')

# 读取CSP文件
csp_df = pd.read_excel(r'M:\下载\20260602之前的文件\CSP商品管理模版111-20260602.xlsx')

# 输出预算模板D列
with open(r'C:\Users\Francis LIU\Desktop\budget_d_column.txt', 'w', encoding='utf-8') as f:
    f.write("预算模板D列（供应商）:\n")
    for idx, row in df.iterrows():
        d = row.iloc[3] if len(row) > 3 else ''
        if pd.notna(d) and str(d).strip() and str(d) != '供应商':
            f.write(f"Row {idx+1}: D={d}\n")

# 输出CSP文件中的新SKU
with open(r'C:\Users\Francis LIU\Desktop\csp_new_skus.txt', 'w', encoding='utf-8') as f:
    f.write("CSP文件中的所有SKU:\n")
    for idx, row in csp_df.iterrows():
        code = row.iloc[1] if len(row) > 1 else ''
        name = row.iloc[2] if len(row) > 2 else ''
        if pd.notna(code) and str(code).startswith('Keeta'):
            f.write(f"{code}: {name}\n")

print("Output files created")
