import pandas as pd
from glob import glob
import os
from tables.path import get_all_files
from tables.utils import getIndexes
from templates.templates import get_template
import movecolumn as mc
from .utils import add_sum_formula

FILES_STARTWITH = '3.'
FILENAME_OUT = os.path.join('out', f'{FILES_STARTWITH}.xlsx')


def concat():
    print("\nTable num:", FILES_STARTWITH)
    sheet_name = '1'
    files = get_all_files(FILES_STARTWITH, sheet_name=sheet_name)
    df_total = pd.DataFrame()
    for file in files:
        df_out = crop_data(file, sheet_name)
        df_out = add_district_col(df_out, file)
        df_out = add_total_row(df_out)
       
        df_total = df_out if df_total.empty else pd.concat([df_out, df_total], axis=0)

 
    df_total = add_headers(df_total, FILES_STARTWITH, sheet_name)
    df_total.reset_index(inplace=True, drop=True)
    df_total = add_sum_formula(df_total, "ЖАМИ")

    with pd.ExcelWriter(path=FILENAME_OUT, mode='w', engine='openpyxl') as writer:
        df_total.to_excel(writer, 
                          sheet_name=sheet_name,
                          index=False)
    return df_total


def crop_data(file, sheet_name):
    df = pd.read_excel(file, sheet_name)
    row_start, col_start = getIndexes(df, 'бир ҳафтада 3 мартагача  узилишлар бўлади')[0]
    mask = (df.index > row_start)
    df_out = df[mask]
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
    row, col = getIndexes(df_template, 'бир ҳафтада 3 мартагача  узилишлар бўлади')[0]
    mask_headers = df_template.index <= row
    df_headers = df_template[mask_headers]

    ## adding district col name in headers 
    district_col_name = 'Туман (шаҳар) номи'
    df_headers[district_col_name] = pd.NA
    df_headers.at[row, district_col_name] = district_col_name
    df_headers = mc.MoveTo2(df_headers, district_col_name)
    #######

    ## all tables have edited titles specific to their district which is different in template file 
    #  to be able to concatenate the headers from template making sure the first (title) columns are standardized in both
    # template and data files  
    columns = df_headers.columns.values
    columns[0] = 'Column 1'
    df_headers.columns = columns
    #$###
    df_out = pd.concat([df_headers, df_total], axis=0)
  
    return df_out

def add_district_col(df, filepath):
    col_name = 'Туман (шаҳар) номи'
    district = filepath.split('\\')[7]
    df[col_name] = district 
    df = mc.MoveTo2(df, col_name)

    return df

def add_total_row(df_out):
    columns = df_out.columns.values
    df_out = pd.concat([df_out.head(1), df_out], axis=0)
    df_out.reset_index(inplace=True, drop=True)
    df_out.iloc[0] = pd.NA
    df_out.at[0, columns[1]] = "Жами"
    return df_out