# Autoloader
import sys
import os
from pathlib import Path
local_path = os.path.dirname(os.path.realpath(__file__))

import xlrd
from datetime import datetime
import threading
import shutil

class DataCrossing:

    def __init__(self):
        self.organize()
        self.run()

    def run(self):
        output_path = f'{local_path}{os.sep}xlsx_files_output'
        for folder in os.listdir(output_path):
            tables_path = f'{output_path}{os.sep}{folder}'
            for table in os.listdir(tables_path):
                quarter_data = dict()
                quarter_path = f'{tables_path}{os.sep}{table}'
                for quarter in os.listdir(quarter_path):
                    self.workbook = xlrd.open_workbook(f'{quarter_path}{os.sep}{quarter}', formatting_info=True)
                    self.sheet = self.workbook.sheet_by_index(0)
                    self.rows_number = self.sheet.nrows
                    for row in range(3,self.rows_number):
                        print(row,self.get_color_of_cell(row,0))

    def get_color_of_cell(self, row, column):
        cell = self.sheet.cell(row, column)
        cif = self.sheet.cell_xf_index(row, column)
        iif = self.workbook.xf_list[cif]
        return iif.background.pattern_colour_index

    def write_header(self):
        #title =
        quarter_header_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#380015'})
        worksheet.merge_range(0,quarter_column,0,quarter_column + 5, title , quarter_header_format)
        worksheet.write(1, quarter_column, 'FILE_NAME_ROW_NUMBER')
        worksheet.write(1, quarter_column + 1, 'ORDEM_EXERC')
        worksheet.write(1, quarter_column + 2, 'DT_INI_EXERC')
        worksheet.write(1, quarter_column + 3, 'DT_FIM_EXERC')
        worksheet.write(1, quarter_column + 4, 'CD_CONTA')
        worksheet.write(1, quarter_column + 5, 'DS_CONTA')

    def organize(self):
        output_path = f'{local_path}{os.sep}xlsx_files_output'
        for item in os.listdir(output_path):
            if '.' not in item:
                continue
            refreshed_listdir = os.listdir(output_path)
            file_info = item.split('_')
            symbol = file_info[0]
            sheet_name = file_info[1]
            symbol_path = f'{output_path}{os.sep}{symbol}'
            sheet_path = f'{symbol_path}{os.sep}{sheet_name}'
            try:
                shutil.move(f'{output_path}{os.sep}{item}',f'{sheet_path}{os.sep}{item}')
                self.xlsx_to_xls(f'{sheet_path}{os.sep}{item}')
            except:
                if symbol not in refreshed_listdir:
                    os.mkdir(symbol_path)
                if sheet_name not in os.listdir(symbol_path):
                    os.mkdir(sheet_path)
                shutil.move(f'{output_path}{os.sep}{item}',f'{sheet_path}{os.sep}{item}')
                self.xlsx_to_xls(f'{sheet_path}{os.sep}{item}')

    def xlsx_to_xls(self,file_path):
        shutil.copy(file_path,f'{os.path.splitext(file_path)[0]}.xls')
        os.remove(file_path)



data = DataCrossing()