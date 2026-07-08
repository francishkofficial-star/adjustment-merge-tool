import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')

# 读取预算模板，看D列到底是什么（供应商列）
budget = pd.read_excel('M:\\下载\\Q1Q2预算总表2026.xlsx', header=None)
print('=== 预算模板完整结构 ===')
print('Shape:', budget.shape)

# 打印所有列的header行（行1）
print('\n=== 列标题（行1）===')
for i, v in enumerate(budget.iloc[1].tolist()):
    print(f"  列{i}({chr(65+i)}): {v}")

print('\n=== 预算模板所有行数据 ===')
for i in range(len(budget)):
    row_data = budget.iloc[i].tolist()
    print(f"行{i}: {row_data}")
