from openpyxl import Workbook
from openpyxl import load_workbook

wb_week = Workbook()
ws = wb_week.active
ws['A1'] = 42
ws.append([1, 2, 3])
wb_week.save("sample.xlsx")
