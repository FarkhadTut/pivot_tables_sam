import pandas as pd
from glob import glob
import os
from tables.path import get_all_files

FILES_STARTWITH = '2.'


def concat():
    print("\nTable num:", FILES_STARTWITH)
    files = get_all_files(FILES_STARTWITH)