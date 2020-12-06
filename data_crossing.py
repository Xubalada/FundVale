# Autoloader
import sys
import os
from pathlib import Path
local_path = os.path.dirname(os.path.realpath(__file__))

import xlrd
from datetime import datetime
import threading
import shutil
import itertools
import xlsxwriter
from itertools import cycle

class DataCrossing:

    def __init__(self):
        self.organize()
        self.run()

    def run(self):
        @staticmethod
        def get_quarter_info(quarter_path,quarter, table_data:dict):
                workbook = xlrd.open_workbook(f'{quarter_path}{os.sep}{quarter}')
                sheet = workbook.sheet_by_index(0)
                rows_number = sheet.nrows
                actual_title = ''#str(sheet.cell_value(2,0))
                actual_item = ''#str(sheet.cell_value(3,0))
                for row in range(2,rows_number):
                    if sheet.cell_value(row,0) != '':
                        if sheet.cell_value(row,8) == '*':
                            actual_title =  str(sheet.cell_value(row,0))
                            if actual_title not in table_data:
                                table_data.update({actual_title: {}})
                            continue
                        actual_item = str(sheet.cell_value(row,0))
                        if actual_item not in table_data[actual_title]:
                            table_data[actual_title].update({actual_item:[]})
                        if sheet.cell_value(row,1) != '':
                            table_data[actual_title][actual_item].append(
                                [
                                sheet.cell_value(row,2),
                                sheet.cell_value(row,5),
                                sheet.cell_value(row,6),
                                sheet.cell_value(row,8)
                                ]
                            )
                    if sheet.cell_value(row,1) != '':
                        table_data[actual_title][actual_item].append([
                            sheet.cell_value(row,2),
                            sheet.cell_value(row,5),
                            sheet.cell_value(row,6),
                            sheet.cell_value(row,8)
                        ])
        @staticmethod
        def filter_data(table_data, title, section):
            table_data[title][section]
            for values in table_data[title][section]:
                if values not in table_data[title][section]:
                    continue
                table_data[title][section] = list(filter((values).__ne__, table_data[title][section]))
                a = values[:3]
                table_data[title][section].append(a)
            for values in table_data[title][section]:
                if values not in table_data[title][section]:
                    continue
                qunatity = table_data[title][section].count(values)
                table_data[title][section] = list(filter((values).__ne__, table_data[title][section]))
                if qunatity == 1:
                    continue
                table_data[title][section].append([values,qunatity])

        output_path = f'{local_path}{os.sep}xlsx_files_output'
        #threads
        for folder in os.listdir(output_path):
            tables_path = f'{output_path}{os.sep}{folder}'
            for table in os.listdir(tables_path):
                table_data = dict()
                quarter_path = f'{tables_path}{os.sep}{table}'
                get_data_threads = list()
                print(f' ##################### Crossing Data From: {table} ##################### ')
                for quarter in os.listdir(quarter_path):
                    thread = threading.Thread(target=get_quarter_info.__func__ , args=(quarter_path,quarter,table_data))
                    thread.daemon = True
                    thread.start()
                    get_data_threads.append(thread)

                for item in get_data_threads:
                    item.join()
                filter_data_threads = list()
                for title in table_data:
                    for section in table_data[title]:
                        filter_data_threads
                        thread2 = threading.Thread(target=filter_data.__func__, args=(table_data,title,section))
                        thread2.daemon = True
                        thread2.start()
                        filter_data_threads.append(thread2)
                for item in filter_data_threads:
                    item.join()
                self.create_result_file(symbol=folder,table=table,table_data=table_data)


    def create_result_file(self,symbol,table,table_data):
        self.writer_workbook = xlsxwriter.Workbook(f'{local_path}{os.sep}xlsx_files_output{os.sep}{symbol}_{table}_results.xlsx')
        self.writer_worksheet = self.writer_workbook.add_worksheet()
        self.write_header(sheet=self.writer_worksheet,workbook=self.writer_workbook, title=f'{symbol} - {table}')
        rows_format_1 =  self.writer_workbook.add_format({'bg_color': '#EEEEEE'})
        rows_format_2 =  self.writer_workbook.add_format({'bg_color': '#DDDDDD'})
        rows_format_3 =  self.writer_workbook.add_format({'bg_color': '#CCCCCC'})
        formats = cycle([rows_format_1, rows_format_2, rows_format_3])
        title_format =  self.writer_workbook.add_format({
                        'bold': 1,
                        'border': 1,
                        'align': 'center',
                        'valign': 'vcenter',
                        'font_color': '#FFFFFF',
                        'bg_color': '#753a10'})
        row = 2
        for title in table_data:
            self.writer_worksheet.set_row(row, cell_format=title_format)
            self.writer_worksheet.write(row,0,title)
            self.writer_worksheet.write(row,5,'*')
            row +=1
            for section in table_data[title]:
                print(f'Writing:  {section}')
                rows_format = next(formats)
                self.writer_worksheet.set_row(row, cell_format=rows_format)
                self.writer_worksheet.write(row,0,section)
                if table_data[title][section] == []:
                    self.writer_worksheet.write(row,5,'-')
                    row +=1
                for item in table_data[title][section]:
                    self.writer_worksheet.write(row,1, item[0][2])
                    self.writer_worksheet.write(row,2,item[0][0])
                    self.writer_worksheet.write(row,3,item[0][1])
                    self.writer_worksheet.write(row,5,item[1])
                    row +=1
        self.writer_workbook.close()

    def write_header(self,sheet,workbook,title):
        #title =
        quarter_header_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#6e2d5f'})
        sheet.merge_range(0,1,0,5, title , quarter_header_format)
        sheet.write(1, 1, 'FILE_NAME_ROW_NUMBER')
        sheet.write(1, 2, 'ORDEM_EXERC')
        sheet.write(1, 3, 'CD_CONTA')
        sheet.write(1, 4, 'DS_CONTA')
        sheet.write(1, 5, 'OCURRENCIES')

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
            except:
                if self.symbol not in refreshed_listdir:
                    os.mkdir(symbol_path)
                if sheet_name not in os.listdir(symbol_path):
                    os.mkdir(sheet_path)
                shutil.move(f'{output_path}{os.sep}{item}',f'{sheet_path}{os.sep}{item}')





data = DataCrossing()