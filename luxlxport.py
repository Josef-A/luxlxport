# coding=utf-8
"""luxlxport
Luddes Excel Export

"""
import fnmatch
import os
import pandas as pd


# Pandas behöver xlrd för att läsa Excelfiler

def main():
    """
    Här körs själva rutinen
    """
    # base_folder = u'D:\\Users\\Joffa\\PycharmProjects\\luxlxport\\data'  # Ändra till var excelfilerna ligger
    # out_folder = u'D:\\Users\\Joffa\\PycharmProjects\\luxlxport\\data'  # Ändra till var du vill lägga .csv-filerna
    base_folder = u'/tmp'
    out_folder = u'/tmp'
    extension = u'.xlsx'

    total = None

    for root, dirnames, filenames in os.walk(base_folder):
        for filename in fnmatch.filter(filenames, u'*{}'.format(extension)):
            file_path = os.path.join(root, filename)
            if pd.__version__ < u'0.21.0':
                df = pd.read_excel(file_path, sheetname=u'Svar', skiprows=16)
            else:
                df = pd.read_excel(file_path, sheet_name=u'Svar', skiprows=16)
            filename_wo_ext, ext = os.path.splitext(filename)
            csv_filename = os.path.join(out_folder, u'{}.csv'.format(filename_wo_ext))
            df.to_csv(csv_filename, encoding=u'utf8', index=False)

            # Lägg in GPDR-ID i en egen kolumn
            df[u'GDPR-ID'] = df[df[u'Uppgift'] == u'GDPR-ID'][u'Svar'].values[0]

            # Lägg ihop allt till en stor tabell
            if total is None:
                total = df
            else:
                total = total.append(df, ignore_index=True)

    if total is not None:
        total_csv_filename = os.path.join(out_folder, u'totala.csv')
        total.to_csv(total_csv_filename, encoding=u'utf8', index=False)

    return total


if __name__ == '__main__':
    df = main()
