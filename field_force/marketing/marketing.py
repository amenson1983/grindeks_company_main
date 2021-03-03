import sqlite3

import pandas as pd
import os
import numpy as np
import scipy as scipy

from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import requests
import pprint


# 'C:\\Users\\Anastasia Siedykh\\Desktop'
# 'C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\python_projects\\grindeks_company_main\\sale_out\\exampDir\\dist'
from pandas.tests.io.excel.test_openpyxl import openpyxl
from pandas.tests.io.excel.test_xlsxwriter import xlsxwriter


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
class gamma_classify:
    def __init__(self,date,employee,region,city, sub_region ,adress,test_status,percentage_correct,	try_,link,number):
        self.number = number
        self.sub_region = sub_region
        self.region = region
        self.link = link
        self.try_ = try_
        self.percentage_correct = percentage_correct
        self.test_status = test_status
        self.adress = adress
        self.city = city
        self.employee = employee
        self.date = date
class gamma_classify_2:
    def __init__(self, adress, number):
        self.number = number
        self.adress = adress
class gamma_classify_3:
    def __init__(self, adress):
        self.adress = adress
class gamma_workout:
    def classify_dict(self,item):
        gamma_data_classified_ = []
        adress = str(item[0])
        number = str(item[1])
        st = gamma_classify_2(adress,number)
        gamma_data_classified_.append(st)
        return gamma_data_classified_
    def classify_base_2021_from_xlxs(self, item):
        gamma_data_classified_ = []
        date = str(item[0])
        employee = str(item[1])
        region = str(item[2])
        city = str(item[3])
        sub_region = str(item[4])
        adress = str(item[5])
        test_status = str(item[6])
        percentage_correct = str(item[7])
        try_ = str(item[8])
        link = str(item[9])
        number = str(item[10])
        st = gamma_classify(date,employee,region,city, sub_region ,adress,test_status,percentage_correct,try_,link,number)
        gamma_data_classified_.append(st)
        return gamma_data_classified_
    def import_from_xlsx(self):

        path = "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\python_projects\\grindeks_company_main\\field_force\\marketing\\0.gamma_in.xlsx"
        wb_obj = openpyxl.load_workbook(path)
        sheet_obj = wb_obj.active
        rows_count = str(sheet_obj.calculate_dimension()).rsplit(':')
        rows_count = int(str(rows_count[1])[1:])
        string = []
        classified_base_gamma = []
        for row in range(1, rows_count):
            str_ = []
            for col in range(1, 3):
                cell_obj = sheet_obj.cell(row=row, column=col)
                str_.append(cell_obj.value)
            string.append(str_)
        for i in string:
            x = gamma_workout()
            string_class = x.classify_dict(i)
            classified_base_gamma.append(string_class)
        return classified_base_gamma
    def our_dict_from_xlsx(self):
        path = "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\python_projects\\grindeks_company_main\\field_force\\marketing\\0.crm_base.xlsx"
        wb_obj = openpyxl.load_workbook(path)
        sheet_obj = wb_obj.active
        rows_count = str(sheet_obj.calculate_dimension()).rsplit(':')
        rows_count = int(str(rows_count[1])[1:])
        string = []
        classified_base_gamma = []
        for row in range(1, rows_count):
            str_ = []
            for col in range(1, 2):
                cell_obj = sheet_obj.cell(row=row, column=col)
                str_.append(cell_obj.value)
            string.append(str_)
        for i in string:

            classified_base_gamma.append(i)
        return classified_base_gamma
    def get_our_data_for_gamma(self):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT distinct secondary_2020_629.organization_adress, secondary_2020_629.city_town, secondary_2020_629.ff_region from secondary_2020_629 where secondary_2020_629.month is not 'Январь' is not 'Февраль' is not 'Март' is not 'Апрель' is not 'Май' and secondary_2020_629.office_head_organization = 'Гамма-55 ПФ Аптечна мережа 9-1-1 ул. Полтавський шлях, 27, кв.2'")
            conn.commit()
            results = cursor.fetchall()
        return results


def mapping_by_adress_organization():
    z = gamma_workout()
    list_gamma = z.import_from_xlsx()
    distinct_gamma_list = []
    for i in list_gamma:
        for j in i:
            if j not in distinct_gamma_list:
                distinct_gamma_list.append([j.adress])
    x = gamma_workout()
    list_crm = x.our_dict_from_xlsx()
    mapped_list_for_append = []
    for num in range(0, len(distinct_gamma_list)):
        for entry in distinct_gamma_list:
            mapped_str = fuzz.token_sort_ratio(str(entry), str(list_crm[num]))
            if mapped_str > 70:
                mapped_list_for_append.append([entry, list_crm[num]])
    workbook = xlsxwriter.Workbook(
        'C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\python_projects\\grindeks_company_main\\field_force\\marketing\\result_map.xlsx')
    worksheet = workbook.add_worksheet('BASE')
    # Widen the first column to make the text clearer.
    # worksheet.set_column('A:A', 20)
    bold = workbook.add_format({'bold': True}, )
    worksheet.write('A1', "Gamma", bold)
    worksheet.write('B1', "Crm", bold)
    result = []
    row_index = 1
    for item in mapped_list_for_append:
        item_ = [[str(item[0]),
                  str(item[1])]]

        result.append(item_)
        worksheet.write(int(row_index), int(0), str(item[0]).replace('[', '').replace(']', '').replace("'", ""))
        worksheet.write(int(row_index), int(1), str(item[1]).replace('[', '').replace(']', '').replace("'", ""))

        row_index += 1
    workbook.close()


if __name__ == '__main__':
    #open_file_from_location()
    #mapping gamma and crm
    mapping_by_adress_organization()












