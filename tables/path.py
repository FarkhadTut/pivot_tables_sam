import os 
from glob import glob
import pandas as pd
from templates.templates import get_template

DATA_FOLDER = os.path.join(os.getcwd(), 'туманлар')

def get_all_files(files_startwith, sheet_name=0):
    files = []
    districts = next(os.walk(DATA_FOLDER))[1]
    for district in districts:
        excel_files_all = os.path.join(DATA_FOLDER, district, '*.xlsx')
        for file in glob(excel_files_all):
            filename_base = os.path.basename(file)
            if filename_base.startswith(files_startwith):
                files.append(file)

    check_shapes(files, files_startwith, sheet_name=sheet_name)
    return files

def check_shapes(files, files_startwith, sheet_name=0):
    template_path = get_template(files_startwith)
    df_template = pd.read_excel(template_path, sheet_name=sheet_name)
    col_num_template = df_template.shape[1]
    for file in files:
        df = pd.read_excel(file, sheet_name=sheet_name)
        if col_num_template != df.shape[1]:
            raise Exception(f'\n\nSHAPES DO NOT MATCH!\nTable num: {files_startwith}\nPath: {file}\nShape: {df.shape}\nTemplate shape: {df_template.shape}\n\n')