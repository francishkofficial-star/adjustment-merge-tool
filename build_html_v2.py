import json

with open(r'C:\Users\Francis LIU\Desktop\Catdesk file\merged_717_v2.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

html = '''<!DOCTYPE html>
<html lang="zh-HK">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>物料調賬底表 7.1-7.7 合併工具</title>
<script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: "Microsoft JhengHei", "PingFang HK", sans-serif;
  background: #f0f2f5;
  color: #333;
  padding: 20px;
}
h1 { text-align: center; margin-bottom: 8px; color: #1a1a2e; font-size: 22px; }
.subtitle { text-align: center; color: #666; margin-bottom: 20px; font-size: 14px; }
.toolbar {
  display: flex;
  justify-content: center;
  gap: 10px;
  margin-bottom: 16px;
  flex-wrap: wrap;
}
.btn {
  padding: 10px 18px;
  border: none;
  border-radius: 6px;
  font-size: 14px;
  cursor: pointer;
  transition: all 0.2s;
  font-family: inherit;
}
.btn-blue { background: #2980b9; color: #fff; }
.btn-blue:hover { background: #1a6da0; }
.btn-green { background: #27ae60; color: #fff; }
.btn-green:hover { background: #1e8449; }
.btn-purple { background: #8e44ad; color: #fff; }
.btn-purple:hover { background: #7d3c98; }
.file-input-wrapper {
  position: relative;
  overflow: hidden;
  display: inline-block;
}
.file-input-wrapper input[type=file] {
  position: absolute;
  left: -9999px;
}
.stats-bar {
  background: #fff;
  border-radius: 8px;
  padding: 12px 20px;
  margin-bottom: 16px;
  box-shadow: 0 1px 4px rgba(0,0,0,0.06);
  display: flex;
  gap: 30px;
  justify-content: center;
  flex-wrap: wrap;
  font-size: 14px;
}
.stats-bar .stat-item { text-align: center; }
.stats-bar .stat-num { font-size: 20px; font-weight: bold; color: #2980b9; }
.stats-bar .stat-label { color: #888; font-size: 12px; margin-top: 2px; }
.search-box {
  width: 100%;
  max-width: 400px;
  padding: 10px 14px;
  border: 1px solid #ddd;
  border-radius: 6px;
  font-size: 14px;
  font-family: inherit;
  margin: 0 auto 16px;
  display: block;
}
.table-wrap {
  overflow-x: auto;
  background: #fff;
  border-radius: 8px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.08);
  max-height: 70vh;
  overflow-y: auto;
}
table {
  width: 100%;
  border-collapse: collapse;
  min-width: 1600px;
}
th, td {
  padding: 8px 10px;
  border-bottom: 1px solid #e8e8e8;
  border-right: 1px solid #f0f0f0;
  font-size: 13px;
  white-space: nowrap;
}
th {
  background: #2c3e50;
  color: #fff;
  position: sticky;
  top: 0;
  z-index: 10;
  font-weight: 600;
  text-align: center;
}
td { text-align: center; }
td.shop-id { font-family: monospace; font-weight: 600; color: #2c3e50; }
td.shop-name { text-align: left; max-width: 280px; overflow: hidden; text-overflow: ellipsis; }
td.address { text-align: left; max-width: 300px; overflow: hidden; text-overflow: ellipsis; }
td.amount { text-align: right; font-weight: 600; color: #e67e22; }
td.remark { text-align: left; color: #555; white-space: normal; max-width: 250px; }
tr:hover { background: #f5f9fc; }
tr.merged-row { background: #fffde7; }
tr.merged-row:hover { background: #fff9c4; }
td.qty-zero { color: #ccc; }
td.qty-nonzero { font-weight: 600; color: #2c3e50; }
.empty-msg { text-align: center; padding: 60px; color: #999; font-size: 16px; }
.drag-area {
  border: 2px dashed #bbb;
  border-radius: 8px;
  padding: 25px;
  text-align: center;
  color: #888;
  margin-bottom: 16px;
  background: #fafbfc;
  transition: all 0.2s;
  cursor: pointer;
}
.drag-area.dragover {
  border-color: #2980b9;
  background: #eaf5fb;
}
.formula-hint {
  text-align: center;
  color: #888;
  font-size: 12px;
  margin-bottom: 12px;
}
</style>
</head>
<body>

<h1>物料調賬底表 7.1-7.7 合併工具</h1>
<p class="subtitle">自動合併相同商家ID · 金額求和 · 備註欄生成物料名稱*數量</p>

<div class="drag-area" id="dragArea">
  <p>拖拽 .xlsx 文件到此處，或點擊下方按鈕選擇文件</p>
</div>

<div class="toolbar">
  <div class="file-input-wrapper">
    <button class="btn btn-blue" onclick="document.getElementById('fileInput').click()">選擇文件</button>
    <input type="file" id="fileInput" accept=".xlsx,.xls">
  </div>
  <button class="btn btn-blue" onclick="processFile()">開始合併</button>
  <button class="btn btn-green" onclick="exportExcel()">導出 Excel</button>
  <button class="btn btn-purple" onclick="exportCSV()">導出 CSV</button>
</div>

<p class="formula-hint">調賬費用公式：大膠袋*29 + 小膠袋*22 + 單杯無紡布袋*150 + 雙杯無紡布袋*180 + 3號無紡布袋*230 + 4號無紡布袋*280 + 紙袋*200 + 紙漿杯2托*200 + 紙漿杯托-4托*200 + 飲品封口紙*10 + 雙杯裝紙袋*360 + 單杯裝紙袋*320</p>

<div class="stats-bar" id="statsBar">
  <div class="stat-item"><div class="stat-num" id="statRaw">0</div><div class="stat-label">原始行數</div></div>
  <div class="stat-item"><div class="stat-num" id="statMerged">0</div><div class="stat-label">合併後行數</div></div>
  <div class="stat-item"><div class="stat-num" id="statDup">0</div><div class="stat-label">重複ID數</div></div>
  <div class="stat-item"><div class="stat-num" id="statTotal">0</div><div class="stat-label">總金額(HKD)</div></div>
</div>

<input type="text" class="search-box" id="searchBox" placeholder="搜索商家ID / 名稱 / 備註..." oninput="filterTable()">

<div class="table-wrap" id="tableWrap">
  <div class="empty-msg">請上傳 Excel 文件後點擊「開始合併」</div>
</div>

<script>
const MATERIAL_COLS = [
  { letter: 'F', name: '大膠袋', price: 29 },
  { letter: 'G', name: '小膠袋', price: 22 },
  { letter: 'H', name: '單杯無紡布袋', price: 150 },
  { letter: 'I', name: '雙杯無紡布袋', price: 180 },
  { letter: 'J', name: '3號無紡布袋', price: 230 },
  { letter: 'K', name: '4號無紡布袋', price: 280 },
  { letter: 'L', name: '紙袋', price: 200 },
  { letter: 'M', name: '紙漿杯2托', price: 200 },
  { letter: 'N', name: '紙漿杯托-4托', price: 200 },
  { letter: 'O', name: '飲品封口紙', price: 10 },
  { letter: 'P', name: '雙杯裝紙袋', price: 360 },
  { letter: 'Q', name: '單杯裝紙袋', price: 320 },
];

let rawData = [];
let mergedData = [];
let rawRowCount = 0;
let dupCount = 0;

const PRELOADED = __DATA_PLACEHOLDER__;

function buildRemark(materials) {
  const parts = [];
  for (const col of MATERIAL_COLS) {
    const qty = materials[col.name] || 0;
    if (qty > 0) parts.push(col.name + '*' + qty);
  }
  return parts.join(', ');
}

function calcAmount(materials) {
  let total = 0;
  for (const col of MATERIAL_COLS) {
    total += (materials[col.name] || 0) * col.price;
  }
  return total;
}

function mergeData(rows) {
  const map = {};
  for (const row of rows) {
    const sid = String(row['商家ID'] || row['shop_id'] || '').trim();
    if (!sid) continue;
    if (!map[sid]) {
      map[sid] = {
        shop_id: sid,
        submit_date: row['提交日期'] || row['submit_date'] || '',
        shop_name: row['商家名稱'] || row['shop_name'] || '',
        phone: row['收貨電話'] || row['phone'] || '',
        address: row['收貨地址'] || row['address'] || '',
        materials: {}
      };
      for (const col of MATERIAL_COLS) map[sid].materials[col.name] = 0;
    }
    for (const col of MATERIAL_COLS) {
      const key = col.name;
      map[sid].materials[key] += parseInt(row[key] || 0) || 0;
    }
  }
  const result = [];
  for (const sid in map) {
    const d = map[sid];
    d.remark = buildRemark(d.materials);
    d.amount = calcAmount(d.materials);
    result.push(d);
  }
  result.sort((a, b) => String(a.shop_id).localeCompare(String(b.shop_id), undefined, {numeric:true}));
  return result;
}

function renderTable() {
  if (!mergedData.length) {
    document.getElementById('tableWrap').innerHTML = '<div class="empty-msg">無數據</div>';
    return;
  }
  let html = '<table><thead><tr>';
  html += '<th>提交日期</th><th>商家ID</th><th>商家名稱</th><th>收貨電話</th><th>收貨地址</th>';
  for (const col of MATERIAL_COLS) html += '<th>' + col.name + '</th>';
  html += '<th>調賬費用</th><th>備註</th>';
  html += '</tr></thead><tbody>';
  for (const row of mergedData) {
    const isDup = row._wasMerged;
    html += '<tr' + (isDup ? ' class="merged-row"' : '') + '>';
    html += '<td>' + esc(row.submit_date) + '</td>';
    html += '<td class="shop-id">' + esc(row.shop_id) + '</td>';
    html += '<td class="shop-name" title="' + esc(row.shop_name) + '">' + esc(row.shop_name) + '</td>';
    html += '<td>' + esc(row.phone) + '</td>';
    html += '<td class="address" title="' + esc(row.address) + '">' + esc(row.address) + '</td>';
    for (const col of MATERIAL_COLS) {
      const qty = row.materials[col.name] || 0;
      html += '<td class="' + (qty > 0 ? 'qty-nonzero' : 'qty-zero') + '">' + qty + '</td>';
    }
    html += '<td class="amount">' + row.amount + '</td>';
    html += '<td class="remark">' + esc(row.remark) + '</td>';
    html += '</tr>';
  }
  html += '</tbody></table>';
  document.getElementById('tableWrap').innerHTML = html;
}

function esc(text) {
  if (text == null) return '';
  return String(text).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function updateStats() {
  document.getElementById('statRaw').textContent = rawRowCount;
  document.getElementById('statMerged').textContent = mergedData.length;
  document.getElementById('statDup').textContent = dupCount;
  const total = mergedData.reduce((s, r) => s + r.amount, 0);
  document.getElementById('statTotal').textContent = total.toFixed(0);
}

function filterTable() {
  const q = document.getElementById('searchBox').value.toLowerCase().trim();
  const rows = document.querySelectorAll('#tableWrap tbody tr');
  rows.forEach(tr => {
    if (!q) { tr.style.display = ''; return; }
    const text = tr.textContent.toLowerCase();
    tr.style.display = text.includes(q) ? '' : 'none';
  });
}

const dragArea = document.getElementById('dragArea');
const fileInput = document.getElementById('fileInput');

dragArea.addEventListener('click', () => fileInput.click());
dragArea.addEventListener('dragover', e => { e.preventDefault(); dragArea.classList.add('dragover'); });
dragArea.addEventListener('dragleave', () => dragArea.classList.remove('dragover'));
dragArea.addEventListener('drop', e => {
  e.preventDefault();
  dragArea.classList.remove('dragover');
  if (e.dataTransfer.files.length) {
    fileInput.files = e.dataTransfer.files;
    dragArea.querySelector('p').textContent = '已選擇: ' + e.dataTransfer.files[0].name;
  }
});
fileInput.addEventListener('change', () => {
  if (fileInput.files.length) dragArea.querySelector('p').textContent = '已選擇: ' + fileInput.files[0].name;
});

function processFile() {
  if (!fileInput.files.length) { alert('請先選擇文件'); return; }
  const file = fileInput.files[0];
  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const wb = XLSX.read(data, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (json.length < 2) { alert('文件內容為空'); return; }
    const headers = json[0].map(h => String(h).trim());
    rawRowCount = json.length - 1;
    rawData = [];
    for (let i = 1; i < json.length; i++) {
      const obj = {};
      for (let j = 0; j < headers.length; j++) obj[headers[j]] = json[i][j];
      rawData.push(obj);
    }
    const idCounts = {};
    rawData.forEach(r => {
      const sid = String(r['商家ID'] || '').trim();
      if (sid) idCounts[sid] = (idCounts[sid] || 0) + 1;
    });
    dupCount = Object.values(idCounts).filter(c => c > 1).length;

    mergedData = mergeData(rawData);
    for (const r of mergedData) {
      r._wasMerged = idCounts[r.shop_id] > 1;
    }
    renderTable();
    updateStats();
  };
  reader.readAsArrayBuffer(file);
}

function exportExcel() {
  if (!mergedData.length) { alert('無數據可導出'); return; }
  const exportData = mergedData.map((r, ri) => {
    const excelRow = ri + 2; // data starts at row 2
    const obj = {
      '提交日期': r.submit_date,
      '商家ID': r.shop_id,
      '商家名稱': r.shop_name,
      '收貨電話': r.phone,
      '收貨地址': r.address,
    };
    for (const col of MATERIAL_COLS) obj[col.name] = r.materials[col.name] || 0;
    // Build formula string for amount: =F2*29+G2*22+...
    const formulaParts = [];
    for (let idx = 0; idx < MATERIAL_COLS.length; idx++) {
      const col = MATERIAL_COLS[idx];
      const excelCol = String.fromCharCode(70 + idx); // F=70, G=71, ...
      formulaParts.push(excelCol + excelRow + '*' + col.price);
    }
    obj['調賬費用'] = '=' + formulaParts.join('+');
    obj['備註'] = r.remark;
    return obj;
  });
  const ws = XLSX.utils.json_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '合併後數據');
  XLSX.writeFile(wb, '合併後-物料調賬底表7.1-7.7.xlsx');
}

function exportCSV() {
  if (!mergedData.length) { alert('無數據可導出'); return; }
  const exportData = mergedData.map(r => {
    const obj = {
      '提交日期': r.submit_date,
      '商家ID': r.shop_id,
      '商家名稱': r.shop_name,
      '收貨電話': r.phone,
      '收貨地址': r.address,
    };
    for (const col of MATERIAL_COLS) obj[col.name] = r.materials[col.name] || 0;
    obj['調賬費用'] = r.amount;
    obj['備註'] = r.remark;
    return obj;
  });
  const ws = XLSX.utils.json_to_sheet(exportData);
  const csv = XLSX.utils.sheet_to_csv(ws);
  const blob = new Blob(['\\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = '合併後-物料調賬底表7.1-7.7.csv';
  link.click();
}

(function init() {
  if (PRELOADED && PRELOADED.length) {
    mergedData = PRELOADED.map(r => ({
      shop_id: r.shop_id,
      submit_date: r.submit_date,
      shop_name: r.shop_name,
      phone: r.phone,
      address: r.address,
      materials: r.materials,
      remark: r.remark,
      amount: calcAmount(r.materials),
      _wasMerged: false
    }));
    rawRowCount = 344;
    dupCount = 28;
    renderTable();
    updateStats();
    document.getElementById('dragArea').querySelector('p').textContent = '已預載入 7.1-7.7 調賬數據（344行→316行）。也可上傳新文件重新處理。';
  }
})();
</script>

</body>
</html>'''

html = html.replace('__DATA_PLACEHOLDER__', json.dumps(data, ensure_ascii=False))

output_path = r'C:\Users\Francis LIU\Desktop\Catdesk file\物料調賬底表7.1-7.7合併工具.html'
with open(output_path, 'w', encoding='utf-8') as f:
    f.write(html)

print(f'HTML saved: {output_path}')
print(f'Size: {len(html)} bytes')
