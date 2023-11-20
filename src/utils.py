import pandas as pd
import openpyxl as pxl
from openpyxl.utils.dataframe import dataframe_to_rows

SYS_EXIT_ERROR_NUMBER = 1
CONVERTED_PATH = './Converted.xlsx'


def has_duplicates(df, subset, name=''):
    if df.duplicated(subset=subset).any():
        print(pd.concat(group for _, group in df.groupby(
            subset, dropna=False) if len(group) > 1))
        print("{} has duplicates.".format(name))
        return True
    return False


def convert_account_excel_to_csv(input_df, output_df, output_file_path):

    output_df[['Date', 'Details', 'Category', 'Notes']
              ] = input_df[['Date', 'Details', 'Category', 'Notes']]

    def get_currency_amount(row):
        s = row[['AUD', 'USD', 'CNY']][pd.notna(row[['AUD', 'USD', 'CNY']])]
        currency = s.index.values[0]
        amount = s.values[0]
        return pd.Series([currency, amount])

    output_df[['Currency', 'Amount']] = input_df.apply(
        get_currency_amount, axis=1)
    output_df.to_csv(output_file_path, index=False)


def convert_account_csv_to_excel():
    TOTAL_ACCOUNT_FILE_PATH = './_data/account.csv'
    OUTPUT_FILE_PATH = './total_acount.xlsx'
    OUTPUT_SHEET_NAME = 'Total Account'
    OUTPUT_COLUMNS = ['Date', 'Details', 'AUD', 'CNY', 'USD', 'Category',
                      'Notes']

    input_df = pd.read_csv(TOTAL_ACCOUNT_FILE_PATH)
    input_df['Date'] = pd.to_datetime(input_df['Date']).dt.date
    output_df = pd.DataFrame(columns=OUTPUT_COLUMNS)

    output_df[['Date', 'Details',
              'Category', 'Notes']] = input_df.loc[:, ['Date', 'Details',
                                                       'Category', 'Notes']]

    def get_col_amount(row):
        row_number = row.name
        currency = row['Currency']
        amount = row['Amount']
        output_df.loc[row_number, currency] = amount

    input_df.apply(get_col_amount, axis=1)

    wb = pxl.Workbook()
    ws = wb.active
    ws.title = OUTPUT_SHEET_NAME

    for r in dataframe_to_rows(output_df, index=False, header=True):
        ws.append(r)

    wb.save(filename=OUTPUT_FILE_PATH)

