""" 
说明：
    运行此脚本，可以将原选课名单的学生均匀地分成[助教人数]组
    该脚本只应运行一次。在之后有新的学生选课后，应运行updategroup.py，以免已分组学生的分组发生变化。
"""

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
import os

# 找到未分组的成绩考核登记表
suffix = '成绩考核登记表.xlsx'
curr_dir_filelist = [f for f in os.listdir('.') if os.path.isfile(f)]
filename = ''
for f in curr_dir_filelist:
    if f.endswith(suffix):
        filename = f
if filename == '':
    print('表格未找到')
    exit(0)

# 打开表格，实施分组
grade_wb = load_workbook(filename)
main_ws = grade_wb.active
if len(grade_wb.worksheets):
    other_sheets = grade_wb.worksheets[1::]
    for s in other_sheets:
        grade_wb.remove(s)

NUM_GROUP = 5

def copy_sheet_header(main_ws:Worksheet, dst_ws:Worksheet):
    # determine max column
    #dst_ws.merge_cells('A1:M7')
    for r in main_ws.iter_rows(1, 7):
        for c in r:
            dst_ws.cell(c.row, c.column, c.value)
            # dst_ws.cell


sub_ws = grade_wb.create_sheet('WSC')
copy_sheet_header(main_ws, sub_ws)

grade_wb.save('test.xlsx')