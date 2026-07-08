import os
files=[f for f in os.listdir(r'C:\Users\Francis LIU\Desktop') if '.xlsx' in f.lower()]
with open(r'C:\Users\Francis LIU\Desktop\Catdesk file\xlsx_files.txt','w',encoding='utf-8') as f:
    for fn in files:
        f.write(fn+'\n')
