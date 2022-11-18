# This is a sample Python script.

# Press *Shift+F10* to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
# Press *Ctrl+F8* to toggle the breakpoint.
from openpyxl import Workbook, load_workbook

file = '../smell1.xlsx'
wb = load_workbook(file)
ws_nodes = wb.worksheets[1]

def get_tech(target):
    tech = ''
    for i in range(1, len(ws_nodes['B'])):
        if target == ws_nodes['B'][i].value:
            cell = 'D' + str(i)
            tech = ws_nodes[cell].value
    return tech

# Edges F column
def fill_cells(sheet_id, find_col, write_col):
    ws_edges = wb.worksheets[sheet_id]
    col_target_ids = ws_edges[find_col]
    n = len(col_target_ids)
    for i in range(1, n):
        curr_target = col_target_ids[i].value
        if curr_target is not None:
            technique = get_tech(int(curr_target))
            ws_edges[write_col + str(i+1)] = technique

def do_something_in_excel():
    fill_cells(2, 'D', 'F')
    wb.save('smell1.xlsx')

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    do_something_in_excel()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
