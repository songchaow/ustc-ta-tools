import os

OUTPUT_FILENAME = '成绩考核登记表-分组.xlsx'
INPUT_FILE_SUFFIX = '成绩考核登记表.xlsx'
NUM_GROUP = 5

def find_filename_with_suffix(s : str):
    curr_dir_filelist = [f for f in os.listdir('.') if os.path.isfile(f)]
    for f in curr_dir_filelist:
        if f.endswith(INPUT_FILE_SUFFIX):
            return f
    return ''