from math import dist
import openpyxl
from cell import Cell
from cell import Color

PATH_DATA = "Neuroscience_data_analysis\data\PNN_Microglia_distance.xlsx"
SHEET = 's512 B2 centered (1) detection '
wb = openpyxl.load_workbook(PATH_DATA)
sheet1 = wb[SHEET]
wb.active = sheet1
row_count = sheet1.max_row
col_count = sheet1.max_column

print(row_count)
print(col_count)

# Initializing Cells
red_cells_list = []
white_cells_list = []
for i in range(2, row_count + 1):
    curr_id = "Cell #" + str(i)
    curr_type = Color.WHITE
    if "PNN" in sheet1.cell(row=i, column=1).value:
        curr_type = Color.RED
    curr_x = sheet1.cell(row=i, column=2).value
    curr_y = sheet1.cell(row=i, column=3).value
    curr_cell = Cell(curr_id,curr_x,curr_y,curr_type)
    if curr_cell.type == Color.WHITE:
        red_cells_list.append(curr_cell)
    else:
        white_cells_list.append(curr_cell)

sheet2 = wb['Distances']
wb.active = sheet2

# Naming headers
col=2
for red_cell in red_cells_list:
    sheet2.cell(row=1,column=col).value=red_cell.id
    col += 1
row=2
for white_cell in white_cells_list:
    sheet2.cell(row=row,column=1).value=white_cell.id
    row += 1

# Calulating distances
for  r_index, red_cell in enumerate(red_cells_list):
    for  w_index, white_cell in enumerate(white_cells_list):
        point_red = [red_cell.position_x, red_cell.position_y]
        point_white = [white_cell.position_x, white_cell.position_y]
        distance = dist(point_red, point_white)
        sheet2.cell(row=r_index+2, column=w_index+2).value=distance
wb.save(PATH_DATA)

    