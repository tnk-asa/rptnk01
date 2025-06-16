import os
import re

def get_matching_files(folder_path, extensions):
    result_list = []

    # 正規表現のパターン（例：123456_AB12CD34EF56_Z）
    pattern = re.compile(r'\d{6}_[A-Za-z0-9]{16}_[A-Z].*$')

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            name, ext = os.path.splitext(file)
            if ext.lower() in [e.lower() for e in extensions] and pattern.match(name):
                result_list.append(os.path.join(root, file))

    return result_list


folder = r"C:\\Users\\P0244\Desktop\\医療機器XML"
exts = ['.txt', '.xml']
files = get_matching_files(folder, exts)

#for f in files:
#    print(f)
