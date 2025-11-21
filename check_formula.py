import openpyxl

# 读取公式（不是计算值）
wb = openpyxl.load_workbook('data/test1/outputs/sheetcopilot_glm-4.5-air/1_57072_output.xlsx', data_only=False)
ws = wb['Sheet2']

print('=== 检查 B 列前 10 行的公式/值 ===')
for row in range(1, 11):
    cell = ws[f'B{row}']
    print(f'{cell.coordinate}: value={cell.value}, formula={cell.value if isinstance(cell.value, str) and cell.value.startswith("=") else "No formula"}')

wb.close()
