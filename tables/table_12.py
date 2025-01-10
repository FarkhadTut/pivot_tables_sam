import pandas as pd
from glob import glob
import os
from tables.path import get_all_files
from tables.utils import getIndexes
from templates.templates import get_template

FILES_STARTWITH = '12. '
FILENAME_OUT = os.path.join('out', f'{FILES_STARTWITH}.xlsx')


def concat():
    print("\nTable num:", FILES_STARTWITH)
    sheet_name = 0
    files = get_all_files(FILES_STARTWITH, sheet_name=sheet_name)
    df_total = pd.DataFrame()
    for file in files:
        df_out = crop_data(file, sheet_name)
        df_total = df_out if df_total.empty else pd.concat([df_out, df_total], axis=0)

    df_total = add_headers(df_total, FILES_STARTWITH, sheet_name)
    with pd.ExcelWriter(path=FILENAME_OUT, mode='w', engine='openpyxl') as writer:
        df_total.to_excel(writer, 
                          sheet_name=str(sheet_name),
                          index=False)
    return df_total


def crop_data(file, sheet_name):
    df = pd.read_excel(file, sheet_name)
    try:
        row_start, col_start = getIndexes(df, 'Жами')[0]
    except:
        row_start, col_start = getIndexes(df, 'ўқувчи ўрни')[0]

    mask = (df.index > row_start)
    df_out = df[mask]
    columns = df_out.columns.values.tolist()
    _, col_start = getIndexes(df, 'ўқувчи ўрни')[0]
    mask_non_empty = ~df_out[columns[columns.index(col_start)+1:]].isna().all(axis=1)
    if df_out[mask_non_empty].shape[0] != 1:
        if df_out[mask_non_empty].shape[0] == 0:
            raise Exception(f'\n\nNo data filled!\nPath:{file}\n\n')
        else:
            raise Exception(f'\n\nMultiple districts filled with data!\nPath:{file}\n\n')
    df_out = df_out[mask_non_empty]
    # df_ksz = df_ksz.tail(-2) ## minus КСЗлар  бўйича and Жами КСЗ
    df_out.dropna(how='all', inplace=True, axis=0)

    ## all tables have edited titles specific to their district which is different in template file 
    #  to be able to concatenate the headers from template making sure the first (title) columns are standardized in both
    # template and data files  
    columns = df_out.columns.values
    columns[0] = 'Column 1'
    df_out.columns = columns
    #$###
    return df_out

    
def add_headers(df_total, file_startswith, sheet_name=0):
    template_path = get_template(file_startswith)
    df_template = pd.read_excel(template_path, sheet_name)
    row, col = getIndexes(df_template, 'ўқувчи ўрни')[0]
    mask_headers = df_template.index <= row
    df_headers = df_template[mask_headers]

    ## all tables have edited titles specific to their district which is different in template file 
    #  to be able to concatenate the headers from template making sure the first (title) columns are standardized in both
    # template and data files  
    columns = df_headers.columns.values
    columns[0] = 'Column 1'
    df_headers.columns = columns
    #$###

    df_out = pd.concat([df_headers, df_total], axis=0)
    return df_out
    