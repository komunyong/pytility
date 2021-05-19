# https://kanoki.org/2019/02/26/compare-two-excel-files-for-difference-using-python/

""" Usage:
- Please fix all columns type in excel
- Please make sure that number of columns and rows are matched
- Please check for column index (column name) are matched
- Do compare by filename dtype_<filename>.txt that this script auto generate after first run failed
"""
import pandas as pd
import numpy as np
import functools
import collections
import warnings
import time


""" Convert column that have all 'nan' to string (NaN, NULL, False, '')
    To prevent error of pandas always treat dtypes as 'float64'
"""
def nan_to_str(df_in):
    for col in df_in.columns:
        if df_in.isnull().all()[col]:
            df_in[col] = df_in[col].astype(str)
    return df_in


start = time.process_time()

FILE_A = 'input_a.xlsx'
FILE_B = 'input_b.xlsx'

with warnings.catch_warnings():
    warnings.filterwarnings('ignore', category=FutureWarning)

    df1 = pd.read_excel(FILE_A)
    df2 = pd.read_excel(FILE_B)
    df1 = nan_to_str(df1)
    df2 = nan_to_str(df2)

    # Thrown error and print log when number of columns or rows or column types are not equal
    bool_dtypes = map(lambda t1, t2: t1 == t2, df1.dtypes, df2.dtypes)
    is_dtypes_equal = functools.reduce(lambda bt1, bt2: bt1 and bt2, bool_dtypes, True)
    if not is_dtypes_equal or df1.shape != df2.shape:
        # Debug config for pandas DataFrame
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)
        pd.set_option('display.max_colwidth', -1)

        dtype_df1_filename = 'dtypes_' + FILE_A[0:-5] + '.txt'
        dtype_df2_filename = 'dtypes_' + FILE_B[0:-5] + '.txt'
        print(f'Count of (rows, columns) for both files are NOT MATCH! [{df1.shape} and {df2.shape}]')
        print(f'Column types for both excel files are NOT EQUAL, please see log file [{dtype_df1_filename} and {dtype_df2_filename}]')
        df1.columns = [x.lower() for x in df1.columns]
        df2.columns = [x.lower() for x in df2.columns]
        df1.dtypes.to_csv(dtype_df1_filename)
        df2.dtypes.to_csv(dtype_df2_filename)
        exit()

    # Compare all cells and write to excel data sheet
    df1 = df1.replace(np.nan, '').replace('nan', '')
    df2 = df2.replace(np.nan, '').replace('nan', '')
    comparison_values = df1.values == df2.values
    rows, cols = np.where(comparison_values == False)
    df_result = df1[df1.columns].astype(str)  # Make new DataFrame to prevent error when write same cell in different type

    for item in zip(rows, cols):
        df_result.iloc[item[0], item[1]] = f'"{df1.iloc[item[0], item[1]]}" --> "{df2.iloc[item[0], item[1]]}"'
    writer = pd.ExcelWriter('reconcile_result.xlsx', engine='xlsxwriter')  # engine='openpyxl' not work with encoding

    # Write summary to another sheet
    mismatch_percent = []
    for col in comparison_values.T:
        occur_val = dict(collections.Counter(col))
        occur_val[False] = 0 if False not in occur_val else occur_val[False]
        occur_val[True] = 0 if True not in occur_val else occur_val[True]
        mismatch_percent.append((occur_val[False] / sum(occur_val.values())) * 100)

    summary = {
        'Column Headers': df_result.columns,
        'Mismatch Percentage': mismatch_percent
    }
    df_summary = pd.DataFrame(summary, columns=summary.keys())
    df_summary.to_excel(writer, sheet_name='summary', index=False, header=True)
    df_result.to_excel(writer, sheet_name='diff_data', index=False, header=True)
    writer.save()

    print('End of program!', 'time taken', time.process_time() - start, 'seconds')
