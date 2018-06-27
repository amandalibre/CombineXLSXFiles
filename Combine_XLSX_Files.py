from xlrd import open_workbook
import xlsxwriter
import glob, os
import pandas as pd
from datetime import datetime
from Extra import directory, data_columns, number_format_columns

df = pd.DataFrame()

frontier_dict = []

os.chdir(directory)
for file in glob.glob("*.xlsx"):

    # open xlsx file
    book = open_workbook(file)
    sheet = book.sheet_by_index(0)

    # read header values into the list
    keys = [sheet.cell(0, col_index).value.strip() for col_index in range(sheet.ncols)]

    # read other rows and make into dict
    for row_index in range(1, sheet.nrows):
        row = {keys[col_index]: sheet.cell(row_index, col_index).value for col_index in range(sheet.ncols)}
        frontier_dict.append(row)

# Create an Excel workbook and add worksheet
workbook = xlsxwriter.Workbook(r'C:\Users\Amanda Friedman\PycharmProjects\CombineXLSXFiles\Frontier.xlsx')

# add worksheet for homepage
worksheet = workbook.add_worksheet('All Countries')
worksheet.set_column(0, 40, 18)
worksheet.set_row(0, 41)

# write header to first row
cell_format = workbook.add_format({'bold': True, 'font_color': 'black', 'font_name': 'Calibri', 'font_size': 10})
cell_format.set_pattern(1)
cell_format.set_bg_color('#BDD7EE')
cell_format.set_text_wrap()
cell_format.set_border()
cell_format.set_align('center')
cell_format.set_align('vcenter')
column_count = 0
for column_header in data_columns:
    worksheet.write(0, column_count, column_header, cell_format)
    column_count += 1

# start from second row
row = 1

# make list of key errors
dict_key_errors = []
date_format_errors = []

# add promos
for row_dict in frontier_dict:
    column_count = 0
    try:
        if str(row_dict['Date Collected']).find('-') == -1:
            if row_dict['Country'] + ' ' + row_dict['Operator'] + ': Date Format Error' not in date_format_errors:
                date_format_errors.append(row_dict['Country'] + ' ' + row_dict['Operator'] + ': Date Format Error')
    except KeyError:
        continue
    for column in range(len(data_columns)):
        cell_format = workbook.add_format(
            {'bold': False, 'font_color': 'black', 'font_name': 'Calibri', 'font_size': 10})
        cell_format.set_text_wrap()
        cell_format.set_border()
        cell_format.set_align('center')
        cell_format.set_align('vcenter')
        header_key = data_columns[column_count]
        if header_key in number_format_columns:
            cell_format.set_num_format('#,##0.00')
            try:
                worksheet.write_number(row, column_count, float(row_dict[header_key]), cell_format)
            except ValueError or KeyError:
                try:
                    worksheet.write(row, column_count, row_dict[header_key], cell_format)
                except KeyError:
                    worksheet.write(row, column_count, '', cell_format)
                    if row_dict['Country'] + ' ' + row_dict[
                        'Operator'] + ': Key Error --> ' + header_key not in dict_key_errors:
                        dict_key_errors.append(
                            row_dict['Country'] + ' ' + row_dict['Operator'] + ': Key Error --> ' + header_key)
                        print(row_dict)
        elif header_key == 'Date Collected':
            string_date = str(row_dict[header_key]).strip()
            date_time = datetime.strptime(string_date, '%d-%m-%Y')
            cell_format.set_num_format('dd/mm/yyyy')
            worksheet.write(row, column_count, date_time, cell_format)
        else:
            try:
                worksheet.write(row, column_count, row_dict[header_key], cell_format)
            except KeyError:
                worksheet.write(row, column_count, '', cell_format)
                if row_dict['Country'] + ' ' + row_dict['Operator'] + ': Key Error --> ' + header_key not in dict_key_errors:
                    dict_key_errors.append(row_dict['Country'] + ' ' + row_dict['Operator'] + ': Key Error --> ' + header_key)
                    print(row_dict)
        column_count += 1
    row += 1

for error in dict_key_errors:
    print(error)

for date_error in date_format_errors:
    print(date_error)

if not dict_key_errors and not date_format_errors:
    print('Success!')
