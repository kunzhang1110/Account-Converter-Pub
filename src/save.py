import sys
import pandas as pd
from utils import (has_duplicates, SYS_EXIT_ERROR_NUMBER,
                   CONVERTED_PATH)
from account_rules import (VALID_CATEGORY_NAMES)

IS_DAYFIRST = True  # If use Australia date format dd/MM/yyyy
INPUT_SHEET_NAME = 'Nab Account'
INPUT_COLUMNS = ['Date', 'Details', 'AUD', 'CNY', 'USD', 'Category',
                 'Notes']
DATA_ACCOUNT_PATH = '_data/account.csv'
TOTAL_ACCOUNT_COLUMNS = ['Date', 'Details',
                         'Currency', 'Amount', 'Category', 'Notes']


def validate_input():
    any_na = converted_df[['Date', 'Details', 'Category']].isna().any(
        axis=1)  # if any of the columns contain NA
    all_na = converted_df[['AUD', 'USD', 'CNY']].isna().all(
        axis=1)  # if all of the columns contain NA
    na_rows = pd.concat([converted_df[any_na],
                        converted_df[all_na]])
    if (len(na_rows) > 0):
        print(na_rows)
        print("Above inputs have unexpected NA values.")
        return False

    if has_duplicates(converted_df, [
            'Date', 'Details', 'AUD', 'USD', 'CNY'], name='input_df'):
        return False

    if (~converted_df['Category'].isin(VALID_CATEGORY_NAMES)).any():
        print(
            converted_df[~converted_df['Category'].isin(VALID_CATEGORY_NAMES)])
        print("Above inputs have invalid Category")
        return False
    return True


def get_currency_amount(row):
    s = row[['AUD', 'USD', 'CNY']][pd.notna(row[['AUD', 'USD', 'CNY']])]
    currency = s.index.values[0]
    amount = s.values[0]
    return pd.Series([currency, amount])


converted_df = pd.read_excel(CONVERTED_PATH, sheet_name=INPUT_SHEET_NAME)[
    INPUT_COLUMNS]
converted_df = converted_df.dropna(axis='index', how='all')
converted_df['Date'] = pd.to_datetime(
    converted_df['Date'], dayfirst=IS_DAYFIRST).dt.date
if not validate_input():
    print('Program exits.')
    sys.exit(SYS_EXIT_ERROR_NUMBER)

new_account_df = pd.DataFrame(columns=TOTAL_ACCOUNT_COLUMNS)
new_account_df[['Date', 'Details', 'Category', 'Notes']
               ] = converted_df[['Date', 'Details', 'Category', 'Notes']]
new_account_df[['Currency', 'Amount']] = converted_df.apply(
    get_currency_amount, axis=1)

total_account_df = pd.read_csv(DATA_ACCOUNT_PATH)
total_account_df['Date'] = pd.to_datetime(
    total_account_df['Date'], dayfirst=IS_DAYFIRST).dt.date
total_account_df = pd.concat([total_account_df, new_account_df])
total_account_df.sort_values(by=['Date'], ascending=True, inplace=True)
total_account_df['Date'] = pd.to_datetime(
    total_account_df['Date']).dt.strftime('%d/%m/%Y')
if has_duplicates(total_account_df, [
        'Date', 'Details', 'Amount'], name='total_account_df'):
    total_account_df.drop_duplicates(inplace=True, keep='first')
    print("Duplicates removed.")

try:
    total_account_df.to_csv(DATA_ACCOUNT_PATH,
                            index=False, encoding='utf-8-sig')
except UnicodeEncodeError:
    print('Unicode Error')
    sys.exit(SYS_EXIT_ERROR_NUMBER)

print("Data saved to {}".format(DATA_ACCOUNT_PATH))
