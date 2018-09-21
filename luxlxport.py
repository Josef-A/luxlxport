# coding=utf-8
"""luxlxport
Luddes Exlcel Export

"""
import fnmatch
import os
import pandas as pd


# Pandas behöver xlrd för att läsa Excelfiler

def main():
    """
    Här körs själva rutinen
    """
    base_folder = u'D:\\Users\\Joffa\\PycharmProjects\\luxlxport\\data'  # Ändra till var excelfilerna ligger
    out_folder = u'D:\\Users\\Joffa\\PycharmProjects\\luxlxport\\data'  # Ändra till var du vill lägga .csv-filerna
    extension = u'.xltx'

    for root, dirnames, filenames in os.walk(base_folder):
        for filename in fnmatch.filter(filenames, u'*{}'.format(extension)):
            file_path = os.path.join(root, filename)
            df = pd.read_excel(file_path, sheet_name=u'Svar', skiprows=16)
            filename_wo_ext, ext = os.path.splitext(filename)
            csv_filename = os.path.join(out_folder, u'{}.csv'.format(filename_wo_ext))
            df.to_csv(csv_filename, encoding=u'utf8', index=False)


if __name__ == '__main__':
    main()
