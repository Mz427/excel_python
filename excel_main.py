from openpyxl import Workbook
from openpyxl import load_workbook
from print_date_of_week import DateOfWeek

'''
wb_week = Workbook()
ws = wb_week.active

ws['A1'] = 42
ws.append([1, 2, 3])
d = ws.cell(row=4, column=2, value=10)
d = ws.cell(row=3, column=2).value = "这是第三行, 第二列"
ws.append([1, 2, 3])
cell_range = ws['A1' : 'C2']
col_range = ws['C' : 'D']
for row in cell_range:
	for cell in row:
		print(cell.value)

for row in ws.values:
	for cell in row:
		print(cell)

for row in ws.iter_rows(min_row=1, max_col=3, max_row=2, values_only=True):
	print(row)

#tuple(ws.rows)
#tuple(ws.columns)

wb_week.save("sample.xlsx")
'''

wb_week = load_workbook("sample.xlsx")
ws_week = wb_week.active
ws_week.max_row
ws_week.max_column
l_ben_di = []
l_yi_di = []
l_tag = True

for i in range(2, ws_week.max_row + 1):
	if ws_week.cell(row = i, column = 5).value == "列名5":
		l_tag = False
		continue
	elif l_tag and ws_week.cell(row = i, column = 5).value != None:
		l_ben_di.append(ws_week.cell(row = i, column = 5).value)
	elif not l_tag:
		l_yi_di.append(ws_week.cell(row = i, column = 5).value)

print(l_ben_di)
print(l_yi_di)
