# created by ekin varli
# coding: utf-8


# modules
import sqlite3
import os
import time
from openpyxl import Workbook


class SQLite2Excel:
    def Convert():
        if 'databases' not in os.listdir('.'):
            os.mkdir('databases')

        if 'xslx' not in os.listdir('.'):
            os.mkdir('xslx')

        xslx_file = input('Enter XSLX file name (New file): ')
        database_file = input('Enter DB file name (Your .db file): ')
        table_name = input('Enter database table name: ')

        connection = sqlite3.connect(f'./databases/{database_file}')
        cursor = connection.cursor()

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = 'SQLite2Excel'

        cursor.execute(f'SELECT * FROM {table_name}')
        datas = cursor.fetchall()

        for data in datas:
            worksheet.append(data)

        workbook.save(f'./xslx/{xslx_file}')
        time.sleep(1)
        print('Compile success.')


if __name__ == '__main__':
    SQLite2Excel.Convert()
