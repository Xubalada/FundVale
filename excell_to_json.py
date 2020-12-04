# Autoloader
import sys
import os
from pathlib import Path
local_path = os.path.dirname(os.path.realpath(__file__))


import xlrd
from datetime import datetime
import json

class DataExtrator:
    def __init__(self, file_name: str, sheet_number: int = 0):
        self.file_path = f"{local_path}{os.sep}xls_files_input{os.sep}{file_name}"
        self.workbook = xlrd.open_workbook(self.file_path, formatting_info=True)
        self.sheet = self.workbook.sheet_by_index(sheet_number)
        self.rows_number = self.sheet.nrows
        self.columns_number = self.sheet.ncols
        self.quarters_columns = dict()
        self.data = dict()

    def get_end_date_row(self):
        for row in range(0,self.rows_number):
            if "end date" in self.sheet.cell_value(row, 0).lower():
                return row

    def insert_quarters_in_data(self):
        date_row = self.get_end_date_row()
        for column in range(1, self.columns_number):
            if self.sheet.cell_value(date_row, column) != "":
                date = self.sheet.cell_value(date_row, column)
                datetime_value = str(datetime(*xlrd.xldate_as_tuple(date, 0)).date())
                self.data.update({datetime_value: {}})
                self.quarters_columns.update({datetime_value: column})

    def get_color_of_cell(self, row, column):
        cell = self.sheet.cell(row, column)
        cif = self.sheet.cell_xf_index(row, column)
        iif = self.workbook.xf_list[cif]
        return iif.background.pattern_colour_index

    def get_first_title_row_number(self):
        for row in range(0,self.rows_number):
            if self.get_color_of_cell(row=row,column=0) == 31:
                return row

    def insert_new_title_in_quarters(self,title):
        for quarter in self.data:
            self.data[quarter].update({title:{}})

    def create_json_structure(self):
        self.insert_quarters_in_data()
        number_of_rows = self.rows_number
        first_row = self.get_first_title_row_number()
        actual_title = self.sheet.cell_value(first_row,1)
        self.insert_new_title_in_quarters(actual_title)
        for row in range(first_row + 1 ,number_of_rows):
            if self.get_color_of_cell(row,0) == 31: #cor cinza do background
                actual_title = self.sheet.cell_value(row,1)
                self.insert_new_title_in_quarters(actual_title)
                continue
            subtitle = self.sheet.cell_value(row,0)
            for quarter in self.quarters_columns:
                column_number = self.quarters_columns[quarter]
                self.data[quarter][actual_title].update({subtitle: self.sheet.cell_value(row,column_number)})
        return self.data


if __name__ == "__main__":

    extractor = DataExtrator('PETR4.xls')
    print(extractor.create_json_structure())
# for i in range(0,200):
#     cell = sheet.cell(i,0)
#     cif = sheet.cell_xf_index(i, 0)
#     iif = workbook.xf_list[cif]
#     cbg = iif.background.pattern_colour_index
#     if cbg == 31:
#         print(i)




# # For row 0 and column 0
# for i in range(0,20):
#     if "End Date" in sheet.cell_value(i, 0):
#         date = sheet.cell_value(i,1)
#         datetime_value = datetime(*xlrd.xldate_as_tuple(date, 0))
#         print(datetime_value.date() )
#         break