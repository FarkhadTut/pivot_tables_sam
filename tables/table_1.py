import pandas as pd
from glob import glob
import os
from tables.path import get_all_files
from tables.utils import getIndexes
from templates.templates import get_template

FILES_STARTWITH = '1.'
FILENAME_OUT = os.path.join('out', f'{FILES_STARTWITH}.xlsx')

def concat():
    sheet_name = 'СЗ-1'
    print("\nTable num:", FILES_STARTWITH, f"\nSheet name: {sheet_name}")
    files = get_all_files(FILES_STARTWITH)
    df_eiz_total = pd.DataFrame()
    for i, file in enumerate(files):
        df_eiz = eiz(file)
        for col in df_eiz.columns:
            try:
                total_value = df_eiz[col].sum()
                if isinstance(total_value, int) or isinstance(total_value, float): 
                    df_eiz.at[1, col] = total_value
            except:
                pass
      
        df_eiz_total = df_eiz if df_eiz_total.empty else pd.concat([df_eiz_total, df_eiz.tail(-1)], axis=0) ## adding tail(-1) to remove subheaders from non-first dfs

    df_ksz_total = pd.DataFrame()
    for i, file in enumerate(files):
        df_ksz = ksz(file)
        for col in df_ksz.columns:
            try:
                total_value = df_ksz[col].sum()
                if isinstance(total_value, int) or isinstance(total_value, float): 
                    df_ksz.at[1, col] = total_value
            except:
                pass


        df_ksz_total = df_ksz if df_ksz_total.empty else pd.concat([df_ksz_total, df_ksz.tail(-1)], axis=0)## adding tail(-1) to remove subheaders from non-first dfs

    df_estz_total = pd.DataFrame()
    for i, file in enumerate(files):
        df_estz = estz(file)
        for col in df_estz.columns:
            try:
                total_value = df_estz[col].sum()
                if isinstance(total_value, int) or isinstance(total_value, float): 
                    df_estz.at[1, col] = total_value
            except:
                pass


        df_estz_total = df_estz if df_estz_total.empty else pd.concat([df_estz_total, df_estz.tail(-1)], axis=0)## adding tail(-1) to remove subheaders from non-first dfs
        df_estz_total.reset_index(inplace=True, drop=True)

    for col in df_eiz_total.columns:
        try:
            mask = df_eiz_total['Unnamed: 1'] == 'Жами'
            total_value = df_eiz_total[mask][col].sum()
            if isinstance(total_value, int) or isinstance(total_value, float): 
                df_eiz_total.at[0, col] = total_value
        except:
            pass
    
    for col in df_ksz_total.columns:
        try:
            mask = df_ksz_total['Unnamed: 1'] == 'Жами'
            total_value = df_ksz_total[mask][col].sum()
            if isinstance(total_value, int) or isinstance(total_value, float): 
                df_ksz_total.at[0, col] = total_value
        except:
            pass

    for col in df_estz_total.columns:
        try:
            mask = df_estz_total['Unnamed: 1'] == 'Жами'
            total_value = df_estz_total[mask][col].sum()
            if isinstance(total_value, int) or isinstance(total_value, float): 
                df_estz_total.at[0, col] = total_value
        except:
            pass

    df_total = pd.concat([df_eiz_total, df_ksz_total, df_estz_total], axis=0)



    df_total = add_headers(df_total, FILES_STARTWITH)
    df_total.columns
  
    df_total.to_excel(FILENAME_OUT, 
                      sheet_name=sheet_name,
                      index=False)
    return df_total

def eiz(file):
    df = pd.read_excel(file)
    row_start, col_start = getIndexes(df, 'ЭИЗлар бўйича')[0]
    row_end, col_end = getIndexes(df, 'КСЗлар  бўйича')[0]
    mask_eiz = (df.index >= row_start) & (df.index < row_end)
    df_eiz = df[mask_eiz]
    # df_eiz = df_eiz.tail(-2) ## minus ЭИЗ бўйича and Жами ЭИЗ
    df_eiz.dropna(how='all', inplace=True, axis=0)

    ## all tables have edited titles specific to their district which is different in template file 
    #  to be able to concatenate the headers from template making sure the first (title) columns are standardized in both
    # template and data files  
    columns = df_eiz.columns.values
    columns[0] = 'Column 1'
    df_eiz.columns = columns
    #$###
    df_eiz.reset_index(inplace=True, drop=True)
    return df_eiz

def ksz(file):
    df = pd.read_excel(file)
    row_start, col_start = getIndexes(df, 'КСЗлар  бўйича')[0]
    row_end, col_end = getIndexes(df, 'ЁСТЗ бўйича')[0]
    mask_ksz = (df.index >= row_start) & (df.index < row_end)
    df_ksz = df[mask_ksz]
    # df_ksz = df_ksz.tail(-2) ## minus КСЗлар  бўйича and Жами КСЗ
    df_ksz.dropna(how='all', inplace=True, axis=0)

    ## all tables have edited titles specific to their district which is different in template file 
    #  to be able to concatenate the headers from template making sure the first (title) columns are standardized in both
    # template and data files  
    columns = df_ksz.columns.values
    columns[0] = 'Column 1'
    df_ksz.columns = columns
    #$###
    df_ksz.reset_index(inplace=True, drop=True)
    return df_ksz

def estz(file):
    df = pd.read_excel(file)
    row_start, col_start = getIndexes(df, 'ЁСТЗ бўйича')[0]
    mask_estz = (df.index >= row_start)
    df_estz = df[mask_estz]
    # df_estz = df_estz.tail(-2) ## minus ЁСТЗлар  бўйича and Жами ЁСТЗ
    df_estz.dropna(how='all', inplace=True, axis=0)

    ## all tables have edited titles specific to their district which is different in template file 
    #  to be able to concatenate the headers from template making sure the first (title) columns are standardized in both
    # template and data files  
    columns = df_estz.columns.values
    columns[0] = 'Column 1'
    df_estz.columns = columns
    #$###
    df_estz.reset_index(inplace=True, drop=True)
    return df_estz

    
def add_headers(df_total, file_startswith):
    template_path = get_template(file_startswith)
    df_template = pd.read_excel(template_path)
    row, col = getIndexes(df_template, 'Жами')[0]
    mask_headers = df_template.index < row
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
    