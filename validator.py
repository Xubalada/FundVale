# Autoloader
import sys
import os
from pathlib import Path
local_path = os.path.dirname(os.path.realpath(__file__))

from itertools import cycle
import xlsxwriter
from excell_to_json import DataExtrator
from dbquery import DBquery
import threading
import re


class Validator:

    @staticmethod
    def get_json_file(file_name,sheet_number):
        extractor = DataExtrator(file_name=file_name,sheet_number=sheet_number)
        return extractor.create_json_structure()

    @staticmethod
    def run(DEBUG:bool=False):
        DEBUG = DEBUG
        dbquery = DBquery()
        files_to_validate = os.listdir(f"{local_path}{os.sep}xls_files_input")
        for item in files_to_validate:
            symbol = os.path.splitext(item)[0]
            cd_cvm = dbquery.get_cd_cvm(symbol=symbol)
            for sheet_number in range(0,3):
                json = Validator.get_json_file(file_name=item,sheet_number=sheet_number)
                for quarter in json:
                    row = 2
                    if sheet_number == 0:
                        sheet_name = "income"
                        tables_name = ['itr_cia_aberta_dre']
                        has_dt_ini_exerc = False
                    elif sheet_number == 1:
                        sheet_name = "balance"
                        tables_name = ['itr_cia_aberta_bpa','itr_cia_aberta_bpp']
                        has_dt_ini_exerc = True
                    else:
                        sheet_name = "cache"
                        tables_name = ['itr_cia_aberta_dfc_mi']
                        has_dt_ini_exerc = False
                    if f'{symbol}_{sheet_name}_{quarter}.xlsx' in os.listdir(f'{local_path}{os.sep}xlsx_files_output'):
                        print(f'{symbol}_{sheet_name}_{quarter}.xlsx is already in xlsx_files_output, processing next item')
                        continue
                    print(f' ############################ {sheet_name.upper()} - {quarter} ############################ ')
                    workbook = xlsxwriter.Workbook(f'{local_path}{os.sep}xlsx_files_output{os.sep}{symbol}_{sheet_name}_{quarter}.xlsx')
                    worksheet = workbook.add_worksheet()
                    Validator.wirte_quarter_header(
                        workbook=workbook,
                        worksheet=worksheet,
                        quarter_column=1,
                        quarter_name=quarter
                    )
                    for title in json[quarter]:
                        #cells formats
                        title_format = workbook.add_format({
                        'bold': 1,
                        'border': 1,
                        'align': 'center',
                        'valign': 'vcenter',
                        'bg_color': '#8b387b'})
                        rows_format_1 = workbook.add_format({'bg_color': '#EEEEEE'})
                        rows_format_2 = workbook.add_format({'bg_color': '#DDDDDD'})
                        rows_format_3 = workbook.add_format({'bg_color': '#CCCCCC'})
                        formats = cycle([rows_format_1, rows_format_2, rows_format_3])
                        date_format = workbook.add_format({'num_format': 'mm/dd/yy'})
                        #code
                        worksheet.set_row(row, cell_format=title_format)
                        worksheet.write(row,0,title)
                        worksheet.write(row,8,'*')
                        row +=1
                        print(f'Group: {title}')
                        print('')
                        for section in json[quarter][title]:
                            print(f'Validating:  {section}')
                            rows_format = next(formats)
                            worksheet.set_row(row, cell_format=rows_format)
                            worksheet.write(row, 0, section)
                            if json[quarter][title][section] == '':
                                row += 1
                                continue
                            results_total = Validator.get_query_list(
                                dbquery= dbquery,
                                cd_cvm=cd_cvm,
                                quarter_date_fim_exec=quarter,
                                eikon_value= float(json[quarter][title][section]),
                                tables_name=tables_name,
                                has_dt_ini_exerc=has_dt_ini_exerc,
                                DEBUG=DEBUG
                            )
                            print(results_total)
                            for result in results_total:
                                worksheet.write(row,1, result['ref'])
                                worksheet.write(row,2, result['ordem_exerc'])
                                worksheet.write(row,3, result['dt_ini_exerc'], date_format)
                                worksheet.write(row,4, result['dt_fim_exerc'], date_format)
                                worksheet.write(row,5, result['cd_conta'])
                                worksheet.write(row,6, result['ds_conta'])
                                worksheet.write(row,7, result['vl_conta'])
                                worksheet.write(row,8, float(json[quarter][title][section]))
                                row += 1
                            if results_total == []:
                                worksheet.write(row,8, float(json[quarter][title][section]))
                                row += 1

                    workbook.close()

    def get_query_list(dbquery,cd_cvm, quarter_date_fim_exec, eikon_value, tables_name, has_dt_ini_exerc,DEBUG):
            dbquery=dbquery
            results_total = list()
            queue = list()
            result_exact = threading.Thread(
                target=dbquery.postgres_query,
                args=(cd_cvm, eikon_value, tables_name, quarter_date_fim_exec, queue, False, has_dt_ini_exerc)
            )
            result_exact.daemon = True
            result_exact.start()
            result_exact_like = []
            result_plus = []
            result_minus = []
            if DEBUG:
                if eikon_value == 0:
                    print(eikon_value)
            if eikon_value != 0:
                exact_like_value, minus_value, plus_value = Validator.get_eikon_values_to_search(eikon_value,DEBUG)
                if exact_like_value != None:
                    result_exact_like = threading.Thread(
                        target=dbquery.postgres_query,
                        args=(cd_cvm, exact_like_value, tables_name, quarter_date_fim_exec, queue, True, has_dt_ini_exerc)
                    )
                    result_exact_like.daemon= True
                    result_exact_like.start()

                result_plus = threading.Thread(
                    target=dbquery.postgres_query,
                    args=(cd_cvm, plus_value, tables_name, quarter_date_fim_exec, queue, True, has_dt_ini_exerc)
                )
                result_plus.daemon= True
                result_plus.start()

                result_minus = threading.Thread(
                    target=dbquery.postgres_query,
                    args=(cd_cvm, minus_value, tables_name, quarter_date_fim_exec, queue, True, has_dt_ini_exerc)
                )
                result_minus.daemon= True
                result_minus.start()
                if exact_like_value != None:
                    result_exact_like.join()
                result_plus.join()
                result_minus.join()
            result_exact.join()
            for result in queue:
                if result not in results_total:
                    results_total.append(result)
            #print(result_exact, result_minus, result_plus, results_total)
            return results_total

    @staticmethod
    def get_eikon_values_to_search(eikon_value,DEBUG):
        #print(eikon_value)
        #print(type(eikon_value))
        none_zero_decimal = re.sub(r'(0+)?.?[0]$','',str(eikon_value))
        if -1 < eikon_value < 1:
            decimal_potence = 10**(len(str(eikon_value).split('.')[1])+1)
            minus_value = str(((eikon_value*decimal_potence)-1)/decimal_potence)
            plus_value = str(((eikon_value*decimal_potence)+1)/decimal_potence)
            exact_like_value = None
            if DEBUG:
                print('(-1 < eikon_value < 1:) ',eikon_value, exact_like_value, minus_value, plus_value)
                print('')
            return exact_like_value, minus_value, plus_value
        elif 10 > eikon_value >= 1 or -10 < eikon_value <= -1:
            if '.' in none_zero_decimal:
                decimal_potence = 10**len(none_zero_decimal.split('.')[1])
                minus_value = str(((float(none_zero_decimal)*decimal_potence)-1)/decimal_potence)
                plus_value = str(((float(none_zero_decimal)*decimal_potence)+1)/decimal_potence)
                exact_like_value = None
                if DEBUG:
                    print('(10 > eikon_value >= 1 -> "." in none_zero_decimal:) ',eikon_value, exact_like_value, minus_value, plus_value)
                    print('')
                return exact_like_value, minus_value, plus_value
            minus_value = str(((eikon_value*10)-1)/10)
            plus_value = str(((eikon_value*10)+1)/10)
            exact_like_value = None
            if DEBUG:
                print('(10 > eikon_value >= 1:) ',eikon_value, exact_like_value, minus_value, plus_value)
                print('')
            return exact_like_value, minus_value, plus_value
        elif '.' in none_zero_decimal:
            int_part = none_zero_decimal.split('.')[0]
            exact_like_value = int_part
            minus_value = str(int(int_part)-1)
            plus_value = str(int(int_part)+1)
            if DEBUG:
                print('("." in none_zero_decimal:) ',eikon_value, exact_like_value, minus_value, plus_value)
                print('')
            return exact_like_value, minus_value, plus_value
        elif len(str(abs(int(none_zero_decimal)))) == 1:
            exact_like_value = str(int(float(none_zero_decimal))*10)
            minus_value = str((int(float(none_zero_decimal))*10)-1)
            plus_value = str((int(float(none_zero_decimal))*10)+1)
            if DEBUG:
                print('(len(str(abs(int(none_zero_decimal)))) == 1:)',eikon_value, exact_like_value, minus_value, plus_value )
                print('')
            return exact_like_value, minus_value, plus_value
        else:
            exact_like_value = none_zero_decimal
            minus_value = str(int(none_zero_decimal) - 1)
            plus_value = str(int(none_zero_decimal) + 1)
            if DEBUG:
                print('(else:) ',eikon_value, exact_like_value, minus_value, plus_value)
                print('')
            return exact_like_value, minus_value, plus_value

    @staticmethod
    def wirte_quarter_header(workbook,worksheet,quarter_column, quarter_name):
        quarter_name = quarter_name
        quarter_header_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': 'green'})
        worksheet.merge_range(0,quarter_column,0,quarter_column + 7, quarter_name , quarter_header_format)
        worksheet.write(1, quarter_column, 'ref')
        worksheet.write(1, quarter_column + 1, 'ORDEM_EXERC')
        worksheet.write(1, quarter_column + 2, 'DT_INI_EXERC')
        worksheet.write(1, quarter_column + 3, 'DT_FIM_EXERC')
        worksheet.write(1, quarter_column + 4, 'CD_CONTA')
        worksheet.write(1, quarter_column + 5, 'DS_CONTA')
        worksheet.write(1, quarter_column + 6, 'VL_CONTA')
        worksheet.write(1, quarter_column + 7, 'VL_EIKON')


Validator.run(DEBUG=True)