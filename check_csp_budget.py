import pandas as pd

# 读取CSP文件
csp_df = pd.read_excel(r'M:\下载\20260602之前的文件\CSP商品管理模版111-20260602.xlsx')

# 输出Q1Q2预算总表列（第11列）
with open(r'C:\Users\Francis LIU\Desktop\csp_budget_names.txt', 'w', encoding='utf-8') as f:
    f.write("CSP文件Q1Q2预算总表列名称:\n")
    for idx, row in csp_df.iterrows():
        code = row.iloc[1] if len(row) > 1 else ''
        name = row.iloc[10] if len(row) > 10 else ''  # Q1Q2预算总表列
        if pd.notna(code) and str(code).startswith('Keeta'):
            f.write(f"{code}: {name}\n")
        elif pd.notna(name) and str(name).strip():
            f.write(f"Row {idx}: {name}\n")

print("Output written to csp_budget_names.txt")
