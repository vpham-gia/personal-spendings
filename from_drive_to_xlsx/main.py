from from_drive_to_xlsx.spreadsheets import GSheets, ExcelWorkbook
from tkinter import *
from tkinter.filedialog import askopenfilename

url = 'https://docs.google.com/spreadsheets/d/1rhvvDNQfRwofOBiuB3ADKV0wpqcZVJQJiuA2XRRszyo/edit#gid=1682984485'

# filename = '/Users/vinhpg/Documents/Documents-perso/finance/suivi-depenses/Dépenses 2025.xlsx'
filename = askopenfilename()
Tk().update()
ew = ExcelWorkbook(filename=filename)
latest_timestamp = ew.get_latest_timestamp()

gs = GSheets(url=url)
spendings_drive = gs.get_data_from_timestamp(timestamp=latest_timestamp)

for sh_name in spendings_drive['compte'].unique():
    try:
        print('Running script for {}'.format(sh_name))
        rows_to_append = gs.get_number_of_rows_to_append(spendings=spendings_drive, account_name=sh_name)
        ts_values, values = gs.get_timestamp_and_values(spendings=spendings_drive, account_name=sh_name)

        last_row_number = ew.get_last_row_number(ws_name=sh_name)

        ew.insert_and_copy_format(ws_name=sh_name,
                                  row_to_insert_from=last_row_number,
                                  number_rows=rows_to_append)

        ew.update_total_formula(ws_name=sh_name,
                                first_row_to_insert=last_row_number,
                                last_row_to_insert=last_row_number+rows_to_append)

        ew.copy_timestamp(ws_name=sh_name,
                          row_start=last_row_number,
                          row_end=last_row_number+rows_to_append+1,
                          list_ts_values=ts_values)

        ew.copy_spendings_values(ws_name=sh_name,
                                 row_start=last_row_number,
                                 row_end=last_row_number+rows_to_append+1,
                                 list_values=values)
    except:
        print('Error for {}'.format(sh_name))

ew.workbook.save(filename)
print('End of file {}'.format(__file__))
