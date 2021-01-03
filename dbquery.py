# Autoloader
import sys
import os
from pathlib import Path
path = Path(__file__).resolve()
sys.path.append(str(path.parents[0]))

import queue
from postgresql import Postgres

class DBquery:

    @staticmethod
    def postgres_query(
        cd_cvm:int,
        vl_conta,
        tables_name:list,
        dt_fim_exerc:str,
        queue,
        with_like:bool = True,
        has_dt_ini_exerc:bool = False,
    ):
        postgres = Postgres()
        queue = queue
        # if DEBUG:
        #     print(vl_conta)
        if has_dt_ini_exerc:
            none_column = "'None'  as "
        else:
            none_column =""
        querys = list()
        for table in tables_name:
            to_query = f'''
                SELECT ref,
                ordem_exerc,
                {none_column} dt_ini_exerc,
                dt_fim_exerc,
                cd_conta ,
                ds_conta,
                vl_conta FROM "data".cvm_{table}
                WHERE cd_cvm = {cd_cvm}
                AND dt_fim_exerc = '{dt_fim_exerc}'
                '''
            if with_like:
                to_query = f"{to_query} AND CAST(vl_conta AS text) LIKE '{vl_conta}%'"
            else:
                to_query = f'{to_query} AND vl_conta = {vl_conta}'
            querys.append(to_query)
        query = ' UNION ALL '.join(querys)
        # if DEBUG:
        #     print(query)
        #print(query)
        results = postgres.query(sql=query,as_dict=True)
        # query_job = self.client.query(query)
        # results = query_job.result() # Wait for the job to complete.
        # result_dataframe = results.to_dataframe()
        # results_dict = result_dataframe.to_dict(orient='records')
        #print(results)
        queue.extend(results)
        postgres.close()
        return results

    @staticmethod
    def get_cd_cvm(symbol:str):
        postgres = Postgres()
        query = f'''
            SELECT cvm_code FROM "data".b3_companies
            WHERE symbol = '{symbol}'
            LIMIT 1
            '''
        result = postgres.query(sql=query,as_dict=True)
        postgres.close()
        return result[0]['cvm_code']


if __name__ == "__main__":
    dbquery = DBquery()
    b = []
    # a = dbquery.postgres_query(
    #     cd_cvm=9512,
    #     dt_fim_exerc='2020-09-30',
    #     vl_conta=70730000,
    #     tables_name=['itr_cia_aberta_dre'],
    #     queue=b
    # )
    c = dbquery.get_cd_cvm('PETR4')
    print(b)
    print(c)
   # print(a[0]['DT_FIM_EXERC'])
    # print(bgquery.get_cd_cvm(symbol="PETR4"))