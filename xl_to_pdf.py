
import os
import win32com.client
from datetime import datetime as dt
from pywintypes import com_error
import openpyxl as xl

# ! this script will close any excel window you have open.

start = dt.now()

# input path to excel files
in_path = r''  # put filepath here
out_path = os.path.join(in_path, 'pdfs')

# create directory if it doesnt exist
if not os.path.isdir(out_path):
    os.mkdir(out_path)


# all excel files in this script begin with A
names = [x[:-5]for x in os.listdir(in_path) if x.endswith('.xlsx') and x.startswith('A')]
out_list = [x[:-4] for x in os.listdir(out_path) if x.endswith('.pdf')]
conv_list = [x for x in names if x not in out_list]

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False  # prevent excel from showing the gui opening and closing

err_list = []
count = len(conv_list)
i = 0

try:
    for file in conv_list:
        open_file = os.path.join(in_path, file + '.xlsx')
        wb = excel.Workbooks.Open(open_file)
        wb_index = list(range(1, wb.WorkSheets.Count + 1))
        wb.WorkSheets(wb_index).Select()
        wb.ActiveSheet.ExportAsFixedFormat(0, out_path + f'\\{file}')
        i += 1
        wb.Close()
except com_error as e:
    err_list.append(file)
finally:
    excel.Quit()
    diff = (dt.now() - start).total_seconds()
    if count == 0:
        count = 1

    dps = diff/count
    print(f'\n{i} file(s) converted to PDF.\n\nIt took {diff} seconds.\n'
          f'{dps} seconds per sheet.')

    if len(err_list) == 1:
        print(f'There was 1 error.\n\n{err_list}')
    elif len(err_list) > 1:
        print(f'There were {len(err_list)} errors.\n\n{err_list}')
    else:
        print('There were no errors.')
