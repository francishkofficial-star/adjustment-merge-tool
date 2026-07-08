import pandas as pd
import sys

# 读取CSP商品管理模版
df = pd.read_excel(r'M:\下载\20260602之前的文件\CSP商品管理模版111-20260602.xlsx')

# 输出到文件
with open(r'C:\Users\Francis LIU\Desktop\csp_output.txt', 'w', encoding='utf-8') as f:
    f.write("CSP商品管理模版前50行:\n")
    f.write(df.head(50).to_string())
    f.write("\n\n列名:\n")
    for i, col in enumerate(df.columns):
        f.write(f"{i}: {col}\n")
    
    # 尝试找到商品编码和商品名称列
    f.write("\n\n查找商品编码和商品名称列...\n")
    for col in df.columns:
        sample = str(df[col].iloc[0]) if len(df) > 0 else ''
        f.write(f"列 '{col}': 第一行值 = '{sample}'\n")

print("Output written to csp_output.txt")
