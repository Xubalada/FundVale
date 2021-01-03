# Standard library imports
import os
import sys
import psycopg2
from psycopg2.extras import DictCursor


class Postgres:

    def __init__(self, dictConnection: dict = None):
        if dictConnection is None:
            self.dictConnection = {
                'POSTGRES_HOST': 'env6.lionx.ai',
                'POSTGRES_DATABASE': 'lionx',
                'POSTGRES_USER': 'lionx',
                'POSTGRES_PASSWORD': 'FAJvJ4eCVQNexTRL',
                'POSTGRES_PORT': '5432'
            }
        else :
            self.dictConnection = dictConnection
        self.connect()

    def connect(self):
        self._db = psycopg2.connect(
            host=self.dictConnection.get("POSTGRES_HOST"),
            database=self.dictConnection.get("POSTGRES_DATABASE"),
            user=self.dictConnection.get("POSTGRES_USER"),
            password=self.dictConnection.get("POSTGRES_PASSWORD"),
            port=self.dictConnection.get("POSTGRES_PORT"),
        )

    def get_value_from_row(self, row: dict, field: str):
        value = row.get(field)
        if value is not None:
            escape = ''
            if isinstance(value, str):
                value = value.replace('\\n', " ")
                value = value.replace("\\", "")
                value = value.replace("'", "")
                escape = 'E'
            return f"{escape}'{value}'"
        else:
            return 'NULL'

    def get_values_from_row(self, row: dict, fields: list):
        insert_values = []
        update_values = []
        for field in fields:
            value = self.get_value_from_row(row=row, field=field)
            insert_values.append(value)
            update_values.append(f'{field} = {value}')
        return insert_values, update_values

    def insert_update_lion(
        self,
        rows: list,
        fields: list,
        conflict_field: str,
        table: str,
        schema: str = 'public'
    ):
        transaction = ['BEGIN']
        count = 0

        def execute(transaction: list):
            transaction.append('COMMIT')
            sql = ';'.join(transaction)
            self.execute(sql=sql)
        for row in rows:
            insert_values, update_values = self.get_values_from_row(row=row, fields=fields)
            sql = f"""
                INSERT INTO \"{schema}\".{table}
                ({','.join(fields)})
                VALUES({','.join(insert_values)})
            """
            if conflict_field is not None:
                sql += f"""
                    ON CONFLICT ({conflict_field})
                    DO UPDATE SET {','.join(update_values)}
                """
            transaction.append(sql)
            count = count + 1
            if len(transaction) > 5000:
                execute(transaction)
                print(f'\r Total inserted/updated rows in table {table}: {count}', end='')
                transaction = ['BEGIN']
        if(len(transaction) > 1):
            execute(transaction)
        print(f'\r Total inserted/updated rows in table {table}: {count}\n')

    def execute(self, sql: str, attemps: int = 0):
        try:
            cur = self._db.cursor()
            cur.execute(sql)
            self._db.commit()
        except psycopg2.errors.DuplicateTable:
            print('Database already exists')
            cur.execute("ROLLBACK")
        except Exception as e:
            print(e)
            self.connect()
            if attemps < 2:
                self.execute(sql=sql, attemps=(attemps + 1))
        cur.close()
        return True

    def query(self, sql: str, as_dict: bool = False):
        rows = None
        try:
            if as_dict:
                cur = self._db.cursor(cursor_factory=DictCursor)
            else:
                cur = self._db.cursor()
            cur.execute(sql)
            rows = cur.fetchall()
        except Exception as e:
            print(e)
            return None
        to_return = []
        for row in rows:
            if as_dict:
                to_return.append(dict(row))
            else:
                to_return.append(list(row))
        return to_return

    def close(self):
        self._db.close()
