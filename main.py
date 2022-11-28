import datetime

import pandas as pd
import openpyxl

from ContractGenerator import ContractGenerator

if __name__ == '__main__':

    #df = pd.read_csv('employees/employees_2.csv', delimiter=';', encoding='utf-8')
    df = pd.read_excel('employees/employees.xlsx', sheet_name='test-sheet', engine='openpyxl')

    print(df)
    for index, row in df.iterrows():
        contract = ContractGenerator(row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8])
        contract.generateContract()
