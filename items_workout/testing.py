import pandas as pd
import os
import numpy as np
import scipy as scipy

from fuzzywuzzy import fuzz
from fuzzywuzzy import process


#'C:\\Users\\Anastasia Siedykh\\Desktop'
#'C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\python_projects\\grindeks_company_main\\sale_out\\exampDir\\dist'



def open_file_from_location():
    namerec = input('input filename from the desktop')
    try:
        files = os.listdir('C:\\Users\\Anastasia Siedykh\\Desktop')
        print(files)
        filestart = process.extractOne(namerec, files)
        print(filestart)
        if filestart[1] >= 80:
            os.startfile('C:\\Users\\Anastasia Siedykh\\Desktop\\' + filestart[0])

        else:
            print('Файл не найден')
    except FileNotFoundError:
        print('Файл не найден')
def get_list_from_file_in_new_row(filename):
    arr1 = []
    with open(filename, "r") as myfile:
        for line in myfile:
            arr = line.rstrip().split("\n")
            arr1.append(arr)
    return arr1

if __name__ == '__main__':
    open_file_from_location()

