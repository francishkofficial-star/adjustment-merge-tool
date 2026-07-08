import pandas as pd
import os

# Find xlsx file
files = [f for f in os.listdir(r'C:\Users\Francis LIU\Desktop') if f.endswith('.xlsx')]
filepath = os.path.join(r'C:\Users\Francis LIU\Desktop', files[0])

df = pd.read_excel(filepath)

# Save columns and shape to a file
with open(r'C:\Users\Francis LIU\Desktop\Catdesk file\preview.txt', 'w', encoding='utf-8') as f:
    f.write('Columns: ' + str(list(df.columns)) + '\n')
    f.write('Shape: ' + str(df.shape) + '\n\n')
    f.write(df.head(30).to_string(index=False))
    f.write('\n\n--- Full data ---\n')
    f.write(df.to_string(index=False))

# Also save as CSV
df.to_csv(r'C:\Users\Francis LIU\Desktop\Catdesk file\preview.csv', index=False, encoding='utf-8-sig')
print('Done')
