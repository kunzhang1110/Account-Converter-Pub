import pandas as pd
import openpyxl as pxl
from openpyxl.utils.dataframe import dataframe_to_rows

SYS_EXIT_ERROR_NUMBER = 1
CONVERTED_PATH = './Converted.xlsx'


def merge_rows_in_csv():
    TOTAL_ACCOUNT_FILE_PATH = './_data/account.csv'
    OUTPUT_FILE_PATH = './account.csv'

    input_df = pd.read_csv(TOTAL_ACCOUNT_FILE_PATH)
    input_df['Date'] = pd.to_datetime(input_df['Date'])
    output_df = input_df.copy()
    print(output_df)
    def condition_sum(col):
        if col.dtypes == 'float64':
            return sum(col.values)
        else:
            if col.name == 'Notes':
                return ' '.join(col.fillna('').values)
    output_df = output_df.groupby(['Date', 'Details', 'Currency','Category']).aggregate(
        condition_sum).reset_index().reindex(columns=input_df.columns)
    print(output_df)
    output_df.to_csv(OUTPUT_FILE_PATH,
                     index=False, encoding='utf-8-sig')


merge_rows_in_csv()
