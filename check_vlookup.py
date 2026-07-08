import sys, io, glob
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import pandas as pd

sf_files = glob.glob(r'C:\Users\Francis LIU\Downloads\已*.xlsx')
df = pd.read_excel(sf_files[0], dtype=str)

# 打印有值的行看shopid在哪
print('=== 前20行 每列非空值 ===')
for idx in range(min(20, len(df))):
    print(f'--- Row {idx} ---')
    for c in range(len(df.columns)):
        val = df.iloc[idx, c]
        if pd.notna(val) and str(val).strip():
            print(f'  Col{c} [{df.columns[c]}]: [{val}]')

# 搜索包含原始SA shopid格式的值
print('\n=== 搜索含721514386的行 ===')
for c in range(len(df.columns)):
    matches = df[df.iloc[:, c].astype(str).str.contains('721514386', na=False)]
    if len(matches) > 0:
        print(f'  Col{c} [{df.columns[c]}]: {len(matches)} matches')
        for idx in matches.index[:3]:
            print(f'    Row{idx}: [{df.iloc[idx, c]}]')

print('\n=== 非空Q列值样本 ===')
non_null_q = df[df.iloc[:, 16].notna()]
print(f'  Q列非空数: {len(non_null_q)}')
if len(non_null_q) > 0:
    for v in non_null_q.iloc[:5, 16].values:
        print(f'  [{v}]')

print('\n=== 每列非空数量 ===')
for c in range(len(df.columns)):
    n = df.iloc[:, c].notna().sum()
    print(f'  Col{c} [{df.columns[c]}]: {n}')
