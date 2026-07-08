import pandas as pd
import os
import json

# Find xlsx file
files = [f for f in os.listdir(r'C:\Users\Francis LIU\Desktop') if f.endswith('.xlsx')]
filepath = os.path.join(r'C:\Users\Francis LIU\Desktop', files[0])

df = pd.read_excel(filepath)

# The file has: Shop ID, Amount, Currency, 备注
# We need to merge rows with same Shop ID, sum Amount, and combine 备注

# Parse 备注 to extract material*quantity
# e.g. "大膠袋*8, 紙袋*1"

def parse_bz(bz):
    items = {}
    if pd.isna(bz):
        return items
    parts = str(bz).split(',')
    for part in parts:
        part = part.strip()
        if '*' in part:
            name, qty = part.rsplit('*', 1)
            name = name.strip()
            try:
                qty = int(qty.strip())
            except:
                qty = 0
            items[name] = items.get(name, 0) + qty
    return items

merged = {}
for _, row in df.iterrows():
    sid = str(row['Shop ID']).strip()
    amt = row['Amount']
    bz = row['备注']
    if sid not in merged:
        merged[sid] = {'Amount': 0, 'Currency': row['Currency'], 'items': {}}
    merged[sid]['Amount'] += amt
    items = parse_bz(bz)
    for name, qty in items.items():
        merged[sid]['items'][name] = merged[sid]['items'].get(name, 0) + qty

# Build output rows
output = []
for sid, data in merged.items():
    # Build remark: material*qty, material*qty, ...
    remark_parts = []
    for name, qty in sorted(data['items'].items()):
        remark_parts.append(f"{name}*{qty}")
    remark = ', '.join(remark_parts)
    output.append({
        'Shop ID': sid,
        'Amount': data['Amount'],
        'Currency': data['Currency'],
        '备注': remark
    })

out_df = pd.DataFrame(output)
# Save
out_df.to_excel(r'C:\Users\Francis LIU\Desktop\Catdesk file\合并后-物料調賬底表.xlsx', index=False)
out_df.to_csv(r'C:\Users\Francis LIU\Desktop\Catdesk file\合并后-物料調賬底表.csv', index=False, encoding='utf-8-sig')

# Also save as JSON for the web page
with open(r'C:\Users\Francis LIU\Desktop\Catdesk file\merged_data.json', 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print(f"Merged {len(df)} rows into {len(out_df)} rows")
print("Saved to 合并后-物料調賬底表.xlsx and merged_data.json")
