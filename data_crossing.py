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
        def get_quarter_info(quarter_path,quarter, table_data:dict): # each section of the quarter has the repeated values ([ord_exec,cd_conta,ds_conta,vl_conta]) removed, like duplicates decurring of the init_date difference
                workbook = xlrd.open_workbook(f'{quarter_path}{os.sep}{quarter}')
                sheet = workbook.sheet_by_index(0)
                rows_number = sheet.nrows - 1
                quarter_date = quarter.split('_')[-1].split('.')[0]
                actual_title = str(sheet.cell_value(2,0))
                actual_section = str(sheet.cell_value(3,0))
                section_temp = []
                for row in range(2,rows_number):
                    if sheet.cell_value(row,0) != '':
                        if section_temp != []:
                             table_data[actual_title][actual_section]['values'].extend(section_temp)
                             section_temp = []
                        if sheet.cell_value(row,8) == '*':
                            actual_title =  str(sheet.cell_value(row,0))
                            if actual_title not in table_data:
                                table_data.update({actual_title: {}})
                            continue
                        actual_section = str(sheet.cell_value(row,0))
                        if actual_section not in table_data[actual_title]:
                            table_data[actual_title].update({actual_section:{'values':[],'Nulls':0}})
                        if sheet.cell_value(row,1) != '':
                            values = [sheet.cell_value(row,2),sheet.cell_value(row,5),sheet.cell_value(row,6),sheet.cell_value(row,8), quarter_date]
                            if values not in section_temp:
                                section_temp.append(values)
                        else:
                            if sheet.cell_value(row,8) == '':
                                table_data[actual_title][actual_section]['Nulls'] += 1
                    if sheet.cell_value(row,1) != '':
                        values = [sheet.cell_value(row,2),sheet.cell_value(row,5),sheet.cell_value(row,6),sheet.cell_value(row,8), quarter_date]
                        if values not in section_temp:
                            section_temp.append(values)

        @staticmethod
        def filter_data(table_data, title, section):
            def none_max(dict_key1, dict_key2):
                if dict_key1 is None:
                    return dict_key2
                if dict_key2 is None:
                    return dict_key1
                return max(dict_key1, dict_key2)

            def max_dict(dict_1, dict_2):
                all_keys = dict_1.keys() | dict_2.keys()
                return  {k: none_max(dict_1.get(k), dict_2.get(k)) for k in all_keys}
            section_dict = dict() #{'[cd,ds]': {'P':{'VL': repetition,'VL2': repetition},'U':{'VL': repetition,'VL2': repetition}}, 'quarters': [dd-mm-yyyy] }
            for values_list in table_data[title][section]['values']:
                quarter = values_list[4]
                ordem_exerc = values_list[0]
                vl_conta_str = str(values_list[3])
                cd_ds_str = str(values_list[1:3])
                if cd_ds_str not in section_dict:
                    section_dict.update({cd_ds_str: {'P': {}, 'U': {}, 'total': {}, 'quarters': []}})
                if quarter not in section_dict[cd_ds_str]['quarters']:
                    section_dict[cd_ds_str]['quarters'].append(quarter)
                if 'p' in ordem_exerc.lower():
                    ordem_exerc_type = 'P'
                else:
                    ordem_exerc_type = 'U'
                if vl_conta_str not in section_dict[cd_ds_str][ordem_exerc_type]:
                    section_dict[cd_ds_str][ordem_exerc_type].update({vl_conta_str: 1})
                else:
                    section_dict[cd_ds_str][ordem_exerc_type][vl_conta_str] += 1
            for item in section_dict:
                section_dict[item]['total'].update(max_dict(section_dict[item]['P'], section_dict[item]['U']))
                if 'Investments - Long-Term' in section:
                    print('P',section_dict[item]['P'])
                    print('U',section_dict[item]['U'])
                    print('T',section_dict[item]['total'])
            table_data[title][section]['values'] = section_dict

        # @staticmethod
        # def filter_data(table_data, title, section):
        #     repeated_values = dict()
        #     for values in table_data[title][section]['values']:
        #         if values not in table_data[title][section]['values'] or len(values)<4:
        #             continue
        #         value_eikon_repetitions = table_data[title][section]['values'].count(values)
        #         table_data[title][section]['values'] = list(filter((values).__ne__, table_data[title][section]['values']))
        #         new_values = values[:3]
        #         #adiciona novamente em igual quantidade os values sem o eikon_value [orde_exec,cd_conta,ds_conta]
        #         for repetition in range(0,value_eikon_repetitions):
        #             table_data[title][section]['values'].append(new_values)
        #         repeated_values.update({str(values): value_eikon_repetitions})
        #         #print(str(values))
        #     for values in table_data[title][section]['values']:
        #         if values not in table_data[title][section]['values']:
        #             continue
        #         quantiy = table_data[title][section]['values'].count(values)
        #         table_data[title][section]['values'] = list(filter((values).__ne__, table_data[title][section]['values']))
        #         if quantiy == 1:
        #             continue
        #         sep_val_by_exec = dict() # {'[cd,ds]': 'P':{'VL': repetition,'VL2': repetition},'U':{'VL': repetition,'VL2': repetition} }
        #         repeated_values_str = ''
        #         for item in repeated_values:
        #             item_list = eval(item)
        #             query_info = values[1:3]
        #             query_info_str = str(query_info)
        #             if len(set(query_info) & set(item_list)) == 2:
        #                 if query_info_str not in exec_separator:
        #                     exec_separator.update({query_info_str: {'P':{},'U':{}}})
        #                 if 'p' in item_list[0].lower():
        #                     exec_separator[query_info_str]['P'].update({str(item_list[-1]: repeated_values[item])})
        #                 else:
        #                     exec_separator[query_info_str]['U'].update({str(item_list[-1]: repeated_values[item])})
        #         for values in sep_val_by_exec:
        #             penultimo = sep_val_by_exec[values]['P']
        #             ultimo = sep_val_by_exec[values]['U']
        #             intersection = set(penultimo) & set(ultimo)

        #               table_data[title][section]['values'].append([values,])




        output_path = f'{local_path}{os.sep}xlsx_files_output'
        #threads
        for folder in os.listdir(output_path):
            tables_path = f'{output_path}{os.sep}{folder}'
            for table in os.listdir(tables_path):
                table_data = dict()
                quarter_path = f'{tables_path}{os.sep}{table}'
                data_threads_list = list()
                print(f'Crossing Data From: {table}')
                quarters = os.listdir(quarter_path)
                for quarter in quarters:
                    thread = threading.Thread(target=get_quarter_info.__func__ , args=(quarter_path,quarter,table_data))
                    thread.daemon = True
                    thread.start()
                    data_threads_list.append(thread)
                for item in data_threads_list:
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
                number_of_quarters = len(quarters)
                self.create_result_file(symbol=folder,table=table,table_data=table_data,number_of_quarters=number_of_quarters)


    def create_result_file(self,symbol,table,table_data,number_of_quarters):
        self.writer_workbook = xlsxwriter.Workbook(f'{local_path}{os.sep}xlsx_files_output{os.sep}{symbol}{os.sep}{symbol}_{table}_results.xlsx')
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
                rows_format = next(formats)
                self.writer_worksheet.set_row(row, cell_format=rows_format)
                self.writer_worksheet.write(row,0,section)
                non_nulls_quarters_section = number_of_quarters - table_data[title][section]['Nulls']
                #print(table_data[title][section]['values']['Nulls'])
                if table_data[title][section]['values'] == {}:
                    self.writer_worksheet.write(row,3,'-')
                    self.writer_worksheet.write(row,4,non_nulls_quarters_section)
                    row +=1
                for item in table_data[title][section]['values']:
                    # if 'Investments - Long-Term' in section:
                    #     print('P',table_data[title][section]['values'][item]['P'])
                    #     print('U',table_data[title][section]['values'][item]['U'])
                    #     print('tot',table_data[title][section]['values'][item]['total'])

                    item_list = eval(item)
                    total_vl_str = str(table_data[title][section]['values'][item]['total']).replace('}','').replace('{','')
                    ultimo_str = str(table_data[title][section]['values'][item]['U']).replace('}','').replace('{','')
                    penultimo_str =  str(table_data[title][section]['values'][item]['P']).replace('}','').replace('{','')
                    quarters = table_data[title][section]['values'][item]['quarters']
                    quarters_str = str(table_data[title][section]['values'][item]['quarters']).replace(']','').replace('[','')
                    self.writer_worksheet.write(row,1, item_list[1])
                    self.writer_worksheet.write(row,2,item_list[0])
                    self.writer_worksheet.write(row,3,len(quarters))
                    self.writer_worksheet.write(row,4,non_nulls_quarters_section)
                    self.writer_worksheet.write(row,5,quarters_str)
                    self.writer_worksheet.write(row,6,total_vl_str)
                    self.writer_worksheet.write(row,7,ultimo_str)
                    self.writer_worksheet.write(row,8,penultimo_str)
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
        sheet.merge_range(0,0,0,8, title , quarter_header_format)
        sheet.write(1, 1, 'DS_CONTA')
        sheet.write(1, 2, 'CD_CONTA')
        sheet.write(1, 3, 'OCURRENCIES')
        sheet.write(1, 4, 'NON-NULLS QUARTER')
        sheet.write(1, 5, 'QUARTERS')
        sheet.write(1, 6, 'VALUES(INTERSEC)')
        sheet.write(1, 7, 'VALUES(ÚLTIMO)')
        sheet.write(1, 8, 'VALUES(PENÚLTIMO)')


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
                if symbol not in refreshed_listdir:
                    os.mkdir(symbol_path)
                if sheet_name not in os.listdir(symbol_path):
                    os.mkdir(sheet_path)
                shutil.move(f'{output_path}{os.sep}{item}',f'{sheet_path}{os.sep}{item}')





data = DataCrossing()