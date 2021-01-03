# Autoloader
import sys
import os
from pathlib import Path
local_path = os.path.dirname(os.path.realpath(__file__))

import time
import xlrd
from datetime import datetime
import threading
import itertools
import xlsxwriter
from itertools import cycle
from hashlib import sha1
import json

class FinalCrossing:


    @staticmethod
    def check_group_item(itens: list, item_list: list):
        group = itens
        group_to_return = list()
        while True:
            if len(group) > 0:
                item = group.pop(0)
            else:
                break
            group_to_return.append(item)
            item_cd = item.get('cd_conta')
            item_ds = item.get('ds_conta')
            total = len(item_list)
            count = 0
            while True:
                if count == total:
                    break
                count += 1
                if len(item_list) > 0:
                    other_itens = item_list.pop(0)
                else:
                    break
                if other_itens.get('ds_conta') == item_ds or other_itens.get('cd_conta') == item_cd:
                    group.append(other_itens)
                else:
                    item_list.append(other_itens)
        # print('################################################################################', item_list, "\n",'########' ,"\n",group_to_return, '################################################################################')
        # input()
        return group_to_return


    @staticmethod
    def is_null_group(group):
        new_group = list()
        for x in group:
            for value in x.get('values'):
                val = value.split(':')[0]
                if val in new_group:
                    continue
                new_group.append(val)
        zero = True
        for item in new_group:
            if "'0.0'" not in item:
                zero = False
        return zero


    def check_if_is_sum(item):
        quarters = [x.split('-')[1] for x in item['quarters']]
        if len(quarters) > 1:
            if quarters.count('03') == len(quarters):
                return True
        return False


    @staticmethod
    def filter_list(item_list):
        final_groups = list()
        while True:
            if len(item_list) > 0:
                item = item_list.pop(0)
            else:
                break
            if item is None:
                break
            # if 'Ajustes Acumulados de Conversão' in item['ds_conta']:
            #     if 'BMGB4' in item['empresa']:
            #         print(item)
            group = FinalCrossing.check_group_item([item], item_list)
            for a in group:
                if 'Ajustes Acumulados de Conversão' in item['ds_conta']:
                    if 'BMGB4' in item['empresa']:
                        print(group)
            if FinalCrossing.is_null_group(group):
                continue
            new_item = group.pop()
            new_item['cd_conta'] = [new_item['cd_conta']]
            new_item['ds_conta'] = [new_item['ds_conta']]
            del new_item['values']
            for group_item in group:
                intersec_quarters = list(set(group_item['quarters']) & set(new_item['quarters']))
                if len(intersec_quarters) == 0:
                    if group_item['ds_conta'] not in new_item['ds_conta']:
                        new_item['ds_conta'].append(group_item['ds_conta'])
                    if group_item['cd_conta'] not in new_item['cd_conta']:
                        new_item['cd_conta'].append(group_item['cd_conta'])
                    new_item['ocurrencies'] += len(group_item['quarters'])
                    new_item['quarters'].extend(group_item['quarters'])
            if new_item['ocurrencies']/new_item['non_nulls'] >= 0.5:
                final_groups.append(new_item)
            elif FinalCrossing.check_if_is_sum(new_item):
                final_groups.append(new_item)
        return final_groups

    @staticmethod
    def agroup_item(item):
        filtered_itens = list()
        # print(item)
        while True:
            if len(item) > 0:
                option = item.pop(0)
            else:
                break
            total = len(item)
            count = 0
            while True:
                if len(item) > 0:
                    other_option = item.pop(0)
                else:
                    filtered_itens.append(option)
                    break
                if count == total:
                    break
                count += 1
                if set(option.get('cd_conta')) == set(other_option.get('cd_conta')) and set(option.get('ds_conta')) == set(other_option.get('ds_conta')):
                    option['ocurrencies'] += other_option['ocurrencies']
                    option['non_nulls'] += other_option['non_nulls']
                    option['empresa'].extend(other_option['empresa'])
                    filtered_itens.append(option)
                else:
                    item.append(other_option)
        # print(filtered_itens)
        # input()
        return filtered_itens

    @staticmethod
    def agroup_itens_final(final_dict):
        for table in final_dict:
            for section in final_dict[table]:
                for item in final_dict[table][section]:
                    list_item = final_dict[table][section][item]
                    final_dict[table][section][item] = FinalCrossing.agroup_item(list_item)
        #             list_item = final_dict[table][section][item]
        #             try:
        #                 thread = threading.Thread(target=FinalCrossing.agroup_item, args=(list_item))
        #             except:
        #                 print(list_item)
        #             thread.daemon = True
        #             thread.start()
        #             thread_list.append(thread)
        # for threads in thread_list:
        #     threads.join()



    @staticmethod
    def table_crossing(tabela, crossed_data, tabela_path, queue):
        workbook = xlrd.open_workbook(tabela_path)
        sheet = workbook.sheet_by_index(0)
        rows_number = sheet.nrows - 1
        item_list = list()
        actual_section = None
        actual_item = None
        enterprise_name = tabela.split('_')[0]
        for row in range(2, rows_number):
            if sheet.cell_value(row, 0) != '':
                if item_list != list():
                    result = FinalCrossing.filter_list(item_list)
                    queue.append(
                        {
                            'actual_section': actual_section,
                            'result': result,
                            'actual_item': actual_item,
                            'tabela': tabela,
                        }
                    )
                if sheet.cell_value(row, 5) == '*':
                    actual_section = sheet.cell_value(row, 0)
                    continue
                actual_item = sheet.cell_value(row, 0)
            if sheet.cell_value(row, 3) == '-':
                queue.append(
                    {
                        'actual_section': actual_section,
                        'result': [],
                        'actual_item': actual_item,
                        'tabela': tabela,
                    }
                )
                continue
            item_list.append(
                {
                    'ds_conta': sheet.cell_value(row, 1),
                    'cd_conta': sheet.cell_value(row, 2),
                    'ocurrencies': int(sheet.cell_value(row, 3)),
                    'non_nulls': int(sheet.cell_value(row, 4)),
                    'quarters': sheet.cell_value(row, 5).split(','),
                    'values': sheet.cell_value(row, 6).split(','),
                    'empresa': [enterprise_name]
                }
            )


    @staticmethod
    def queue_handler(crossed_data, queue_list, stop):
        while True:
            if stop['stop']:
                break
            if queue_list != []:
                to_exec = queue_list.pop(0)
                table_type = to_exec['tabela']
                if 'cache' in table_type:
                    table_type = 'cache'
                if 'balance' in table_type:
                    table_type = 'balance'
                if 'income' in table_type:
                    table_type = 'income'
                if to_exec['actual_section'] not in crossed_data[table_type]:
                    crossed_data[table_type].update({to_exec['actual_section']: {to_exec['actual_item']: to_exec['result']}})
                elif to_exec['actual_item'] not in crossed_data[table_type][to_exec['actual_section']]:
                    crossed_data[table_type][to_exec['actual_section']].update({to_exec['actual_item']: to_exec['result']})
                else:
                    crossed_data[table_type][to_exec['actual_section']][to_exec['actual_item']].extend( to_exec['result'])

    @staticmethod
    def crosser(empresa, crossed_data, input_path, queue_list):
        empresa_path = f'{input_path}{os.sep}{empresa}'
        lista_de_tabelas = os.listdir(empresa_path)
        tabelas_para_crusar = [item for item in lista_de_tabelas if f'{empresa}_' in item]
        listas_threads_list = list()
        for tabela in tabelas_para_crusar:
            #print(tabela)
            tabela_path = f'{empresa_path}{os.sep}{tabela}'
            # FinalCrossing.table_crossing(tabela, crossed_data, tabela_path, queue_list)
            thread = threading.Thread(target=FinalCrossing.table_crossing, args=(tabela, crossed_data, tabela_path, queue_list))
            thread.daemon = True
            thread.start()
            listas_threads_list.append(thread)
        for thread in listas_threads_list:
            thread.join()


    @staticmethod
    def run(): #crossed_data = {table: {section_1: {iten_1: [], iten_1: [], iten_1: []}}}
        crossed_data = {
            'income': {},
            'cache': {},
            'balance': {}
        }
        stop = {'stop': False}
        queue_list = list()
        queue_thread = threading.Thread(target=FinalCrossing.queue_handler, args=(crossed_data, queue_list, stop))
        queue_thread.daemon = True
        queue_thread.start()
        input_path = f'{local_path}{os.sep}final_crossing_input'
        lista_de_empresas = os.listdir(input_path)
        lista_de_empresas = [item for item in lista_de_empresas if '.' not in item]
        empresa_threads_list = list()
        for empresa in lista_de_empresas:
            # FinalCrossing.crosser(empresa, crossed_data, input_path, queue_list)
            empresa_threads = threading.Thread(target=FinalCrossing.crosser, args=(empresa, crossed_data, input_path, queue_list))
            empresa_threads.daemon = True
            empresa_threads.start()
            empresa_threads_list.append(empresa_threads)
        for thread in empresa_threads_list:
            thread.join()
        stop['stop'] = True
        queue_thread.join()
        FinalCrossing.agroup_itens_final(crossed_data)
        for table in crossed_data:
            FinalCrossing.create_table(table, crossed_data[table])

    @staticmethod
    def create_table_header(workbook, worksheet, table_name):
        quarter_header_format = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#6e2d5f'
        })
        worksheet.merge_range(0,0,0,6, table_name , quarter_header_format)
        worksheet.write(1, 1, 'DS_CONTA')
        worksheet.write(1, 2, 'CD_CONTA')
        worksheet.write(1, 3, 'OCURRENCIES')
        worksheet.write(1, 4, 'NON-NULLS QUARTER')
        worksheet.write(1, 5, 'QUARTERS')
        worksheet.write(1, 6, 'ENTERPRISE')

    @staticmethod
    def create_table(table_name, table):
        print(f'{local_path}{os.sep}final_crossing_output{os.sep}{table_name}')
        writer_workbook = xlsxwriter.Workbook(f'{local_path}{os.sep}final_crossing_output{os.sep}{table_name}.xlsx')
        writer_worksheet = writer_workbook.add_worksheet()
        FinalCrossing.create_table_header(writer_workbook, writer_worksheet, table_name)
        #formating rows
        rows_format_1 =  writer_workbook.add_format({'bg_color': '#EEEEEE'})
        rows_format_2 =  writer_workbook.add_format({'bg_color': '#DDDDDD'})
        rows_format_3 =  writer_workbook.add_format({'bg_color': '#CCCCCC'})
        formats = cycle([rows_format_1, rows_format_2, rows_format_3])
        title_format =  writer_workbook.add_format({
                        'bold': 1,
                        'border': 1,
                        'align': 'center',
                        'valign': 'vcenter',
                        'font_color': '#FFFFFF',
                        'bg_color': '#753a10'})
        row = 2
        for section in table:
            writer_worksheet.set_row(row, cell_format=title_format)
            writer_worksheet.write(row, 0, section)
            writer_worksheet.write(row, 3, '*')
            row +=1
            for item in table[section]:
                rows_format = next(formats)
                writer_worksheet.set_row(row, cell_format=rows_format)
                writer_worksheet.write(row, 0, item)
                if table[section][item] == []:
                    writer_worksheet.write(row, 3, '-')
                    row += 1
                hash_list = list()
                for sub_item in table[section][item]:
                    hash = sha1(json.dumps(sub_item).encode('utf-8')).hexdigest()
                    if hash in hash_list:
                        continue
                    hash_list.append(hash)
                    writer_worksheet.write(row, 1, ', '.join(sub_item['ds_conta']))
                    writer_worksheet.write(row, 2, ', '.join(sub_item['cd_conta']))
                    writer_worksheet.write(row, 3, sub_item['ocurrencies'])
                    writer_worksheet.write(row, 4, sub_item['non_nulls'])
                    writer_worksheet.write(row, 5, ', '.join(sub_item['quarters']).replace("'",""))
                    writer_worksheet.write(row, 6, ', '.join(sub_item['empresa']))
                    row += 1
        writer_workbook.close()


f = FinalCrossing()
f.run()