from merge import main as merge_files
from format import main as prettify_files

import warnings
warnings.filterwarnings('ignore')


def main():
    merge_files()
    prettify_files()

if __name__ == "__main__":
    main()