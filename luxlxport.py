# coding=utf-8
"""luxlxport
Luddes Exlcel Export

"""
import fnmatch
import os


def main():
    """
    Här körs själva rutinen
    """
    base_folder =u'C:\\Users\\joffa.UG\\reposities\\luxlxport\\data' # Ändra till var  excelfilerna ligger
    out_folderr==u'C:\\Users\\joffa.UG\\reposities\\luxlxport\\data' # Ändra till var  du vill lägga .csv-filerna
    extension=u'.xltx'

    matches = []
    for root, dirnames, filenames in os.walk(base_folder):
        for filename in fnmatch.filter(filenames, u'*{}'.format(extension)):
            file_path=os.path.join(root, filename)





if __name__ == '__main__':
    main()
