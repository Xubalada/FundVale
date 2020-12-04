# Autoloader
import sys
import os
from pathlib import Path
path = Path(__file__).resolve()
sys.path.append(str(path.parents[0]))

from google.cloud import bigquery
from google.oauth2 import service_account

class Bigquery:
    def __init__(self):
        #self.DEBUG = DEBUG
        local_base_path = os.path.dirname(os.path.realpath(__file__))
        credentials = service_account.Credentials.from_service_account_file(
        f'{local_base_path}{os.sep}google-service-account.json')
        project_id = 'lionx-269513'
        self.client = bigquery.Client(credentials= credentials, project=project_id)

    def bg_query(
        self,
        cd_cvm:int,
        vl_conta,
        tables_name:list,
        dt_fim_exerc:str,
        with_like:bool = True,
        has_dt_ini_exerc:bool = False
    ):
        # if DEBUG:
        #     print(vl_conta)
        if has_dt_ini_exerc:
            none_column = "'None'  as "
        else:
            none_column =""
        querys = list()
        for table in tables_name:
            to_query = f"""
                SELECT FILE_NAME_ROW_NUMBER,
                ORDEM_EXERC,
                {none_column} DT_INI_EXERC,
                DT_FIM_EXERC,
                CD_CONTA ,
                DS_CONTA,
                VL_CONTA FROM cvm.{table}
                WHERE CAST(CD_CVM AS INT64) = {cd_cvm}
                AND DT_FIM_EXERC = '{dt_fim_exerc}'
                """
            if with_like:
                to_query = f'{to_query} AND CAST(VL_CONTA AS String) LIKE "{vl_conta}%"'
            else:
                to_query = f'{to_query} AND VL_CONTA = {vl_conta}'
            querys.append(to_query)
        query = ' UNION ALL '.join(querys)
        # if DEBUG:
        #     print(query)

        query_job = self.client.query(query)

        results = query_job.result() # Wait for the job to complete.
        result_dataframe = results.to_dataframe()
        results_dict = result_dataframe.to_dict(orient='records')
        return results_dict

    def get_cd_cvm(self, symbol:str):
        query_job = self.client.query(
            f"""
            SELECT cvm_code FROM b3.companies
            WHERE symbol = "{symbol}"
            LIMIT 1
            """
        )
        results = query_job.result() # Wait for the job to complete.
        result_dataframe = results.to_dataframe()
        results_dict = result_dataframe.to_dict(orient='records')
        return int(results_dict[0]['cvm_code'])

if __name__ == "__main__":
    bgquery = Bigquery()
    a = bgquery.bg_query(
        cd_cvm=9512,
        dt_fim_exerc='2020-09-30',
        vl_conta=0,
        tables_name='itr_cia_aberta_dre'
    )
    print(a)
    print(a[0]['DT_FIM_EXERC'])
    # print(bgquery.get_cd_cvm(symbol="PETR4"))