import pandas as pd
import os 
from glob import glob 

def get_template(file_startswith):
    templates_path = os.path.join(os.getcwd(), 'шаблоны')
    files = glob(os.path.join(templates_path, '*.xlsx'))
    files = [f for f in files if not os.path.basename(f).startswith("~")]
    file = [f for f in files if os.path.basename(f).startswith(file_startswith)][0]
    return file
