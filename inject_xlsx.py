import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

html_path = r'C:\Users\Francis LIU\Desktop\Catdesk file\调账底表合并备注生成器.html'
js_path = r'C:\Users\Francis LIU\Desktop\Catdesk file\xlsx.full.min.js'

with open(html_path, 'r', encoding='utf-8') as f:
    html = f.read()

with open(js_path, 'r', encoding='utf-8') as f:
    js_content = f.read()

placeholder = '<!--XLSX_LIB_PLACEHOLDER-->'
replacement = '<script>\n' + js_content + '\n</script>'

html = html.replace(placeholder, replacement)

with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html)

print(f'Done! File size: {len(html.encode("utf-8"))} bytes')
print(f'Placeholder found: {placeholder in html}')
print(f'XLSX.read exists: {"XLSX.read" in html}')

# Verify by checking key markers
check1 = '<script>\n/*! xlsx.js' in html
check2 = 'var XLSX' in html or 'XLSX' in js_content[:200]
print(f'JS embedded at placeholder: {check1}')
print(f'XLSX variable present: {check2}')
