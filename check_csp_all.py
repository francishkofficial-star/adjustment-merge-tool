import pandas as pd

# 读取CSP文件完整内容
csp_df = pd.read_excel(r'M:\下载\20260602之前的文件\CSP商品管理模版111-20260602.xlsx')

# 输出所有行
with open(r'C:\Users\Francis LIU\Desktop\csp_all_rows.txt', 'w', encoding='utf-8') as f:
    f.write("CSP文件所有行:\n")
    for idx, row in csp_df.iterrows():
        code = row.iloc[1] if len(row) > 1 else ''
        name = row.iloc[2] if len(row) > 2 else ''
        f.write(f"Row {idx}: Code={code}, Name={name}\n")

print("Output written to csp_all_rows.txt")
