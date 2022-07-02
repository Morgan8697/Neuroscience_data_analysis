from tkinter.tix import CELL
import openpyxl
from cell import Cell
from cell import Color

path = "./data/PNN_Microglia_distance.xlsx"

wb = openpyxl.load_workbook(path)

sheet = wb.active
row_count = sheet.max_row
col_count = sheet.max_column

print(row_count)
print(col_count)

# Initializing Cells
red_cells_list = []
white_cells_list = []
for i in range(1, row_count + 1):
    curr_id = i
    curr_type = Color.WHITE
    if "PNN" in sheet.cell(row=i, column=1):
        curr_type = Color.RED
    curr_x = sheet.cell(row=i, column=2)
    curr_y = sheet.cell(row=1, column=3)
    curr_cell = Cell(curr_id,curr_x,curr_y,curr_type)
    if curr_cell.type == Color.WHITE:
        red_cells_list.append(curr_cell)
    if curr_cell
        