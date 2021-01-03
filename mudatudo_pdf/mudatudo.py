import os
from tika import parser
import re
local_path = os.path.dirname(os.path.realpath(__file__))

for folder in os.listdir(local_path):
    if '.' in folder:
        continue
    folder_path = f'{local_path}{os.sep}{folder}'
    for pdf in os.listdir(folder_path):
        parsed_pdf = parser.from_file(f'{folder_path}{os.sep}{pdf}')
        dates = re.findall(r'rimestrais - (\d\d)/(\d\d)/(\d\d\d\d)', str(parsed_pdf))
        if dates == []:
            dates = re.findall(r'- (\d\d)/(\d\d)/(\d\d\d\d) -', str(parsed_pdf))
            try:
                print(dates[0])
            except:
                continue
        try:
            if dates[0][1] == '03':
                quarter = 'Q1'
            elif dates[0][1] == '06':
                quarter = 'Q2'
            elif dates[0][1] == '09':
                quarter = 'Q3'
            elif dates[0][1] == '12':
                quarter = 'FULL_Q4'
            file_name = f'{folder}_{dates[0][2]}_{quarter}.pdf'
            if file_name in os.listdir(folder_path):
                print(file_name)
                print(pdf)
                if file_name not in pdf:
                    print('entrou')
                    os.remove(f'{folder_path}{os.sep}{pdf}')
                continue
            os.rename(f'{folder_path}{os.sep}{pdf}',f'{folder_path}{os.sep}{file_name}')
        except:
            continue