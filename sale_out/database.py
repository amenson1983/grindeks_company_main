import csv
import os
import sqlite3

import unidecode
from fuzzywuzzy import process
import pandas
import xlsxwriter
import jaydebeapi
from pandas.tests.io.excel.test_openpyxl import openpyxl
import logging
logging.basicConfig(filename='process_flow_of_job_helper.log', level=logging.INFO, format='%(asctime)s %(message)s', datefmt='%d/%m/%Y %I:%M:%S %p')
from sale_out.sqls import sql_629_from_quadra_server, sql_629_2020_from_quadra_server


def run_refresh_in_big_table_report():
    import xlwings as xw
    logging.info("Opening big tablereport - 'big_table_report_2021_new_1.xlsm' - OK")
    wb = xw.Book(
        'C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\big_table_report_ukraine\\big_table_report_2021_new_1.xlsm')
    app = wb.app
    macro_vba = app.macro("'big_table_report_2021_new_1.xlsm'!open_tabs")
    logging.info("Opening tabs in 'big_table_report_2021_new_1.xlsm' - OK")
    macro_vba()
    logging.info("Updating tabs in  big_table_report_2021_new_1.xlsm from connected transformations - OK")
    macro_vba = app.macro("'big_table_report_2021_new_1.xlsm'!refresh")
    macro_vba()
    macro_vba = app.macro("'big_table_report_2021_new_1.xlsm'!close_tabs")
    logging.info("Closing tabs in  big_table_report_2021_new_1.xlsm - OK")
    macro_vba()
    wb.save()
    wb.close()
connection_url_quadra_server = "jdbc:jtds:sqlserver://62.149.15.123:1433/medowl_grindex"
quadra_login = {
                "user": "grindex",
                "password": "xednirg",
            }
jar_path_constant = 'C:\jtds-1.3.1-dist\\jtds-1.3.1.jar'

class Kam_plan_download_structure:
    def __init__(self,item_quadra, code_sf,month_local, quarter,plan_packs, plan_euro):
        self.plan_euro = plan_euro
        self.plan_packs = plan_packs
        self.quarter = quarter
        self.month_local = month_local
        self.code_sf = code_sf
        self.item_quadra = item_quadra
class Plan_ff_main_2021_download:
    def __init__(self,month_local,item_quadra,med_rep_code,plan_packs,plan_euro):
        self.month_local = month_local
        self.item_quadra = item_quadra
        self.med_rep_code = med_rep_code
        self.plan_packs = plan_packs
        self.plan_euro = plan_euro
class Kam_plans:
    def read_kam_plan(conn):
        plan_list = []
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(f"select distinct kam_plan.item_quadra, kam_plan.code, kam_plan.month,  ymm.Квартал, kam_plan.plan_packs, kam_plan.plan_euro from kam_plan join ymm on ymm.Месяц = kam_plan.month")
            results = cursor.fetchall()
            for i in results:
                y_1 = str(i[0])
                y_2 = str(i[1])
                y_3 = str(i[2])
                y_4 = str(i[3])
                y_5 = str(i[4])
                y_6 = str(i[5]).replace(',', '.')
                z = Kam_plan_download_structure(y_1, y_2, y_3,y_4, y_5, y_6)
                plan_list.append(z)
        return plan_list
    #def fact_kams_packs(self): #TODO to get actual Kam sales packs
    #def fact_kams_euro(self): #TODO to get actual Kam sales euro
    #def plan_fact_kams_packs(self): #TODO to marry
class CEXtract_database_plan_ff: #TODO create ff_plans workout functions
    def read_plan_ff_main(conn):
        rep_plan_list = []
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            logging.info("Connecting to 'local_main_base.db' - OK")
            cursor = conn.cursor()
            cursor.execute(f"select distinct ff_plan.month_local, ff_plan.item_quadra, ff_plan.med_representative_code, ff_plan.packs_plan, ff_plan.euro_plan from ff_plan where ff_plan.year = '2021' and ff_plan.euro_plan <> 0")
            results = cursor.fetchall()
            for i in results:
                y_1 = str(i[0])
                y_2 = str(i[1])
                y_3 = str(i[2])
                y_4 = str(i[3]).replace(',', '.')
                y_5 = str(i[4]).replace(',', '.')

                z = Plan_ff_main_2021_download(y_1, y_2, y_3, y_4, y_5)
                rep_plan_list.append(z)
        return rep_plan_list
    def get_axes_for_ff_sec_plan(conn,rep_plan_list,list_months_quadra):
        base = []
        for string in rep_plan_list:
            for i in string:

                for month in list_months_quadra:
                    if i.month == month:
                        base.append([i.code_sf,i.item_quadra,i.month,i.sales_euro])
        print(base)
class CStock_quadra:
    def __init__(self,week,date,country_region,distributor,item_quadra,quantity_packs,amount_euro,num):
        self.num = num
        self.amount_euro = amount_euro
        self.quantity_packs = quantity_packs
        self.item_quadra = item_quadra
        self.distributor = distributor
        self.country_region = country_region
        self.date = date
        self.week = week


class CStock_quadra_workout:
    def classify_stock_quadra(self,i):
        stock_2021_classifyed = []


        num = i[7]
        amount_euro = i[6]
        quantity_packs = i[5]
        item_quadra = i[4]
        distributor = i[3]
        country_region = i[2]
        date = str(i[1])
        week = i[0]

        entry = CStock_quadra(week,str(date),country_region,distributor,item_quadra,quantity_packs,amount_euro,num)
        stock_2021_classifyed.append(entry)
        return stock_2021_classifyed
    def classified_stock_to_sqlite(self):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            logging.info("Connecting to 'local_main_base.db' - OK")
            cursor = conn.cursor()
            cursor.execute("DROP TABLE stock_at_distributors_wh")
            logging.info("Dropping the table 'stock_at_distributors_wh' at 'local_main_base.db' - OK")
            cursor.execute("CREATE TABLE IF NOT EXISTS stock_at_distributors_wh (week,date,country_region,distributor,item_quadra,quantity_packs,amount_euro,num);")
            logging.info("Creating the table 'stock_at_distributors_wh' at 'local_main_base.db' - OK")
            path = "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\raw_data_files\\stock_distributors_wh\\stock_at_distributors_wh.xlsx"
            logging.info("Opening the 'stock_at_distributors_wh.xlsx' - OK")
            conn.commit()
            wb_obj = openpyxl.load_workbook(path)
            sheet_obj = wb_obj.active
            rows_count = str(sheet_obj.calculate_dimension()).rsplit(':')

            rows_count = int(str(rows_count[1])[1:])
            logging.info(f"Calculating rows at 'stock_at_distributors_wh.xlsx' - number of rows - {rows_count} - OK")

            string = []
            classified_base_2021 = []
            for row in range(2, rows_count + 1):
                str_ = []
                for col in range(1, 9):
                    cell_obj = sheet_obj.cell(row=row, column=col)
                    if cell_obj is not None:
                        str_.append(cell_obj.value)

                string.append(str_)
            logging.info(
                f"Launching stock classification - OK")
            for i in string:
                x = CStock_quadra_workout()

                string_class = x.classify_stock_quadra(i)
                classified_base_2021.append(string_class)
            logging.info(
                f"Stock classified - OK")
            logging.info(
                f"Inserting stock data to base - OK")
            for string in classified_base_2021:
                for row in string:
                    strin = [row.week,row.date,row.country_region,row.distributor,row.item_quadra, row.quantity_packs,row.amount_euro,row.num]
                    cursor.execute("INSERT INTO stock_at_distributors_wh VALUES (?,?,?,?,?,?,?,?);",strin)
            conn.commit()
            print('OK, check the base')
            return classified_base_2021
    #def upload_and_classify_stock_from_quadra(self):# TODO to get how to write SQL script from VStudio
    def get_stock_for_big_table(self):
        classified_stock = self.classified_stock_to_sqlite()
        stock_list = []
        for i in classified_stock:
            for in_ in i:
                stock_list.append([in_.week,in_.date,in_.country_region,in_.distributor,in_.item_quadra,in_.quantity_packs,in_.amount_euro])
        workbook = xlsxwriter.Workbook(
            'C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\transform_files\\0.transform_stock_actual.xlsx')
        worksheet = workbook.add_worksheet()
        bold = workbook.add_format({'bold': True}, )
        worksheet.write('A1', "Неделя", bold)
        worksheet.write('B1', "Дата", bold)
        worksheet.write('C1', "Область", bold)
        worksheet.write('D1', "Дистрибьютор", bold)
        worksheet.write('E1', "Товар", bold)
        worksheet.write('F1', "Упаковки", bold)
        worksheet.write('G1', "Евро", bold)

        row_index = 1
        for item in stock_list:
            if item[0] != None:
                worksheet.write(int(row_index), int(0), str(item[0]))
                worksheet.write(int(row_index), int(1), str(item[1]))
                worksheet.write(int(row_index), int(2), str(item[2]))
                worksheet.write(int(row_index), int(3), str(item[3]))
                worksheet.write(int(row_index), int(4), str(item[4]))
                worksheet.write(int(row_index), int(5), str(item[5]))
                worksheet.write(int(row_index), int(6), str(item[6]).replace(".",","))
                print(item)
            row_index +=1
        workbook.close()


class Tertiary_download_structure:
    def __init__(self,item_kpi,year,month,brand,weight_pen,sro, pen, quantity, amount_euro, weight_sro):
        self.weight_sro = weight_sro
        self.amount_euro = amount_euro
        self.quantity = quantity
        self.pen = pen
        self.sro = sro
        self.weight_pen = weight_pen
        self.brand = brand
        self.month = month
        self.year = year
        self.item_kpi = item_kpi
class Tertiary_by_region_download_structure:
    def __init__(self,year,	period_name,region,	full_medication_name,market_org,quantity,volume,sro,weight_sro):
        self.weight_sro = weight_sro
        self.sro = sro
        self.volume = volume
        self.quantity = quantity
        self.market_org = market_org
        self.full_medication_name = full_medication_name
        self.region = region
        self.period_name = period_name
        self.year = year
class Tertiary_by_region_download_structure_for_power_bi:
    def __init__(self,item_quadra, month, region_quadra, brand, quantity,  volume, sro, weight_sro):
        self.weight_sro = weight_sro
        self.sro = sro
        self.volume = volume
        self.quantity = quantity
        self.brand = brand
        self.region_quadra = region_quadra
        self.month = month
        self.item_quadra = item_quadra
class Tertiary_by_week_download_structure_for_power_bi:
    def __init__(self,item_proxima_dirty,w01,w02,w03,w04,w05,w06,w07,w08,w09,w10,w11,w12,w13,w14,w15,w16,w17,w18):
        self.w18 = w18
        self.w17 = w17
        self.w16 = w16
        self.w15 = w15
        self.w14 = w14
        self.w13 = w13
        self.w12 = w12
        self.w11 = w11
        self.w10 = w10
        self.w09 = w09
        self.w08 = w08
        self.w07 = w07
        self.w06 = w06
        self.w05 = w05
        self.w04 = w04
        self.w03 = w03
        self.w02 = w02
        self.w01 = w01
        self.item_proxima_dirty = item_proxima_dirty

class Tertiary_by_week_download_structure_for_power_bi_workout:
    def normalize_weekly_tertiary(self):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT distinct items.item_proxima from items")
            results = cursor.fetchall()
        list_crm = []
        for i in results:
            list_crm.append(i)
        filename = "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\transform_files\\0.transform_tertiary_week_actual.xlsx"
        path = "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\raw_data_files\\tertiary_sales\\tertiary_week_actual.xlsx"
        wb_obj = openpyxl.load_workbook(path)
        sheet_obj = wb_obj.active
        rows_count = str(sheet_obj.calculate_dimension()).rsplit(':')
        print(rows_count)
        rows_count = int(str(rows_count[1])[1:])
        print(rows_count)
        logging.info(f"number of rows - {rows_count}  - OK")
        string = []
        classified_base_2021 = []
        for row in range(1, rows_count + 1):
            str_ = []
            for col in range(1, 70):
                cell_obj = sheet_obj.cell(row=row, column=col)
                str_.append(cell_obj.value)
            string.append(str_)
        for i in string[2:]:
            print(i[0])
            print(list_crm[0])
            count = 0
            item_proxima_dirty = ''
            index_ = process.extractOne(str(i[0])[7:],list_crm)
            if index_[1] > 91:
                item_proxima_dirty = str(str(i[0])[8:]).replace("('","").replace("',)","")
            else: count +=1

            print(item_proxima_dirty)

            w01 = i[2]
            w02 = i[3]
            w03 = i[4]
            w04 = i[5]
            w05 = i[6]
            w06 = i[7]
            w07 = i[8]
            w08 = i[9]
            w09 = i[10]
            w10 = i[11]
            w11 = i[12]
            w12 = i[13]
            w13 = i[14]
            w14 = i[15]
            w15 = i[16]
            w16 = i[17]
            w17 = i[18]
            w18 = i[19]
            x = Tertiary_by_week_download_structure_for_power_bi(item_proxima_dirty,w01,w02,w03,w04,w05,w06,w07,w08,w09,w10,w11,w12,w13,w14,w15,w16,w17,w18)
            string_class = x
            classified_base_2021.append(string_class)


        workbook = xlsxwriter.Workbook(filename)
        logging.info(f"Opening f'{filename}' for writing  - OK")
        worksheet = workbook.add_worksheet()

        # Widen the first column to make the text clearer.
        # worksheet.set_column('A:A', 20)
        bold = workbook.add_format({'bold': True}, )
        worksheet.write('A1', "item_proxima", bold)
        worksheet.write('B1', "01", bold)
        worksheet.write('C1', "02", bold)
        worksheet.write('D1', "03", bold)
        worksheet.write('E1', "04", bold)
        worksheet.write('F1', "05", bold)
        worksheet.write('G1', "06", bold)
        worksheet.write('H1', "07", bold)
        worksheet.write('I1', "08", bold)
        worksheet.write('J1', "09", bold)
        worksheet.write('K1', "10", bold)
        worksheet.write('L1', "11", bold)
        worksheet.write('M1', "12", bold)
        worksheet.write('N1', "13", bold)
        worksheet.write('O1', "14", bold)
        worksheet.write('P1', "15", bold)
        worksheet.write('Q1', "16", bold)
        worksheet.write('R1', "17", bold)
        worksheet.write('S1', "18", bold)
        logging.info("Adding headers - OK")

        list_base_2021 = []
        row_index = 1

        for item in classified_base_2021:
            item_ = [[str(item.item_proxima_dirty),
                      str(item.w01),
                      str(item.w02),
                      str(item.w03),
                      str(item.w04),
                      str(item.w05),
                      str(item.w06),
                      str(item.w07),
                      str(item.w08),
                      str(item.w09),
                      str(item.w10),
                      str(item.w11),
                      str(item.w12),
                      str(item.w13),
                      str(item.w14),
                      str(item.w15),
                      str(item.w16),
                      str(item.w17),
                      str(item.w18)
                      ]]

            list_base_2021.append(item_)
            worksheet.write(int(row_index), int(0), str(item.item_proxima_dirty))
            worksheet.write(int(row_index), int(1), str(item.w01))
            worksheet.write(int(row_index), int(2), str(item.w02))
            worksheet.write(int(row_index), int(3), str(item.w03))
            worksheet.write(int(row_index), int(4), str(item.w04))
            worksheet.write(int(row_index), int(5), str(item.w05))
            worksheet.write(int(row_index), int(6), str(item.w06))
            worksheet.write(int(row_index), int(7), str(item.w07))
            worksheet.write(int(row_index), int(8), str(item.w08))
            worksheet.write(int(row_index), int(9), str(item.w09))
            worksheet.write(int(row_index), int(10), str(item.w10))
            worksheet.write(int(row_index), int(11), str(item.w11))
            worksheet.write(int(row_index), int(12), str(item.w12))
            worksheet.write(int(row_index), int(13), str(item.w13))
            worksheet.write(int(row_index), int(14), str(item.w14))
            worksheet.write(int(row_index), int(15), str(item.w15))
            worksheet.write(int(row_index), int(16), str(item.w16))
            worksheet.write(int(row_index), int(17), str(item.w17))
            worksheet.write(int(row_index), int(18), str(item.w18))
            row_index += 1
        logging.info("Writing data - OK")
        workbook.close()

class Tertiary_workout:





    def classify(self,col):
        classifyed_tert = []
        year = str(col[0])
        period_name = col[1]
        region = col[2]
        full_medication_name = col[3]
        market_org = col[4]
        quantity = str(col[5])
        volume = str(col[6])
        sro = str(col[7])
        weight_sro = str(col[8])
        x = Tertiary_by_region_download_structure(str(year),	period_name,region,	full_medication_name,market_org,str(quantity),str(volume),str(sro),str(weight_sro))
        classifyed_tert.append(x)
        return classifyed_tert

    def tert_reg_to_sqlite(self):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            logging.info("Connecting to 'local_main_base.db'  - OK")
            cursor = conn.cursor()
            #cursor.execute("DROP TABLE tertiary_by_reg")
            logging.info("Dropping the table 'tertiary_by_reg' in 'local_main_base.db'  - OK")
            cursor.execute("CREATE TABLE IF NOT EXISTS tertiary_by_reg (year,	period_name,region,	full_medication_name,market_org,quantity,volume,sro,weight_sro);")
            path = "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\transform_files\\0.transform_tertiary_by_region.xlsx"
            logging.info("Creating the table 'tertiary_by_reg' in 'local_main_base.db'  - OK")
            conn.commit()
            wb_obj = openpyxl.load_workbook(path)
            sheet_obj = wb_obj.active
            rows_count = str(sheet_obj.calculate_dimension()).rsplit(':')
            print(rows_count)
            rows_count = int(str(rows_count[1])[1:])
            print(rows_count)
            logging.info(f"transform_tertiary_by_region' - number of rows - {rows_count}  - OK")
            string = []
            classified_base_2021 = []
            for row in range(1, rows_count + 1):
                str_ = []
                for col in range(1, 10):
                    cell_obj = sheet_obj.cell(row=row, column=col)
                    str_.append(cell_obj.value)
                string.append(str_)
            logging.info("Classifying received data   - OK")
            for i in string[1:]:

                x = Tertiary_workout()
                string_class = x.classify(i)
                classified_base_2021.append(string_class)
            for string in classified_base_2021:
                for row in string:
                    strin = [row.year,	row.period_name,row.region,	row.full_medication_name,row.market_org,row.quantity,row.volume,row.sro,row.weight_sro]
                    cursor.execute("INSERT INTO tertiary_by_reg VALUES (?,?,?,?,?,?,?,?,?);",strin)
            conn.commit()
            logging.info("Inserting received data to 'local_main_base.db' - OK")

    def tertiary_by_region_to_xlxs(self,list_months):
        filename = "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\transform_files\\0.transform_tertiary_by_region_converted.xlsx"
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT items.item_quadra, ymm.Месяц, UkraineMapInfo.region_quadra, items.brand, tertiary_by_reg.quantity,  tertiary_by_reg.volume, tertiary_by_reg.sro, tertiary_by_reg.weight_sro from items join tertiary_by_reg on tertiary_by_reg.full_medication_name = items.item_proxima join ymm on tertiary_by_reg.period_name = ymm.month join UkraineMapInfo on UkraineMapInfo.region_proxima = tertiary_by_reg.region where tertiary_by_reg.year = '2021' group by items.item_quadra, ymm.Месяц, UkraineMapInfo.region_quadra, items.brand")
            results = cursor.fetchall()
            tertiary_list = []
            for i in results:
                y_1 = str(i[4]).replace(".",",")
                y_2 = str(i[5]).replace(".",",")
                y_3 = str(i[6]).replace(".",",")
                y_4 = str(i[7]).replace(".",",")

                z = Tertiary_by_region_download_structure_for_power_bi(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4)
                tertiary_list.append(z)
        final_list = []
        workbook = xlsxwriter.Workbook(filename)
        logging.info("Opening '0.transform_tertiary_by_region_converted.xlsx' for writing  - OK")
        worksheet = workbook.add_worksheet()

        # Widen the first column to make the text clearer.
        # worksheet.set_column('A:A', 20)
        bold = workbook.add_format({'bold': True}, )
        worksheet.write('A1', "item_quadra", bold)
        worksheet.write('B1', "month", bold)
        worksheet.write('C1', "region", bold)
        worksheet.write('D1', "brand", bold)
        worksheet.write('E1', "quantity", bold)
        worksheet.write('F1', "volume", bold)
        worksheet.write('G1', "sro", bold)
        worksheet.write('H1', "weight_sro", bold)
        logging.info("Adding headers to '0.transform_tertiary_by_region_converted.xlsx' - OK")


        list_base_2021 = []
        row_index = 1

        for item in tertiary_list:
            item_ = [[str(item.item_quadra),
                      str(item.month),
                      str(item.region_quadra),
                      str(item.brand),
                      str(item.quantity),
                      str(item.volume),
                      str(item.sro),
                      str(item.weight_sro)]]

            list_base_2021.append(item_)
            worksheet.write(int(row_index), int(0), str(item.item_quadra))
            worksheet.write(int(row_index), int(1), str(item.month))
            worksheet.write(int(row_index), int(2), str(item.region_quadra))
            worksheet.write(int(row_index), int(3), str(item.brand))
            worksheet.write(int(row_index), int(4), str(item.quantity))
            worksheet.write(int(row_index), int(5), str(item.volume))
            worksheet.write(int(row_index), int(6), str(item.sro))
            worksheet.write(int(row_index), int(7), str(item.weight_sro))


            row_index += 1
        logging.info("Writing data to '0.transform_tertiary_by_region_converted.xlsx' - OK")
        workbook.close()

class Secondary_total_2021:
    def __init__(self,item_quadra,sales_euro):
        self.sales_euro = sales_euro
        self.item_quadra = item_quadra
class Upload_2021_base_from_quadra_for_daily_totals_distr:
    def __init__(self,year,month,week,sales_method,distributor_name, item_quadra,sale_in_quantity,sales_euro_):

        self.sales_euro_ = sales_euro_
        self.sale_in_quantity = sale_in_quantity
        self.item_quadra = item_quadra
        self.distributor_name = distributor_name
        self.sales_method = sales_method
        self.week = week
        self.month = month
        self.year = year
class CSecondary_629_quadra_classify():
    def __init__(self,year,month,ff_region,country_region,city_town,organization_name,organization_adress,sales_method,
                 product_code,item_quadra,organization_etalon_id,organization_etalon_name,distributor_etalon_name,
                 distributor_name,distributor_okpo,sales_euro_,promotion,organization_type,organization_status,
                 etalon_code_okpo,delivery_date,position_code,office_head_organization,head_office_okpo,quarter_year,half_year,
                 annual_sales_category,med_representative_name,kam_name,week,territory_name,brik_name,sale_in_quantity):
        self.position_code = position_code
        self.sale_in_quantity = sale_in_quantity
        self.brik_name = brik_name
        self.territory_name = territory_name
        self.week = week
        self.kam_name = kam_name
        self.med_representative_name = med_representative_name
        self.annual_sales_category = annual_sales_category
        self.half_year = half_year
        self.quarter_year = quarter_year
        self.head_office_okpo = head_office_okpo
        self.office_head_organization = office_head_organization
        self.delivery_date = delivery_date
        self.etalon_code_okpo = etalon_code_okpo
        self.organization_status = organization_status
        self.organization_type = organization_type
        self.promotion = promotion
        self.sales_euro_ = sales_euro_
        self.distributor_okpo = distributor_okpo
        self.distributor_name = distributor_name
        self.distributor_etalon_name = distributor_etalon_name
        self.organization_etalon_name = organization_etalon_name
        self.organization_etalon_id = organization_etalon_id
        self.item_quadra = item_quadra
        self.product_code = product_code
        self.sales_method = sales_method
        self.organization_adress = organization_adress
        self.organization_name = organization_name
        self.city_town = city_town
        self.country_region = country_region
        self.ff_region = ff_region
        self.month = month
        self.year = year
class CSecondary_629_quadra_classify_from_excel():
    def __init__(self, year, month, ff_region, country_region, city_town, organization_name,
                 organization_adress, product_group,
                 product_code, item_quadra, organization_etalon_id, organization_etalon_name,
                 distributor_etalon_name,
                 distributor_name, distributor_okpo, sales_euro, promotion, organization_type,
                 organization_status,
                 etalon_code_okpo, delivery_date,position_code, office_head_organization, head_office_okpo,
                 quarter_year, half_year,
                 annual_sales_category, med_representative_name, kam_name, week, territory_name,
                 brik_name, sales_packs):
        self.position_code = position_code
        self.product_group = product_group
        self.sales_packs = sales_packs
        self.brik_name = brik_name
        self.territory_name = territory_name
        self.week = week
        self.kam_name = kam_name
        self.med_representative_name = med_representative_name
        self.annual_sales_category = annual_sales_category
        self.half_year = half_year
        self.quarter_year = quarter_year
        self.head_office_okpo = head_office_okpo
        self.office_head_organization = office_head_organization
        self.delivery_date = delivery_date
        self.etalon_code_okpo = etalon_code_okpo
        self.organization_status = organization_status
        self.organization_type = organization_type
        self.promotion = promotion
        self.sales_euro = sales_euro
        self.distributor_okpo = distributor_okpo
        self.distributor_name = distributor_name
        self.distributor_etalon_name = distributor_etalon_name
        self.organization_etalon_name = organization_etalon_name
        self.organization_etalon_id = organization_etalon_id
        self.item_quadra = item_quadra
        self.product_code = product_code
        self.organization_adress = organization_adress
        self.organization_name = organization_name
        self.city_town = city_town
        self.country_region = country_region
        self.ff_region = ff_region
        self.month = month
        self.year = year
    def __str__(self):
        return f'{self.year}, {self.month}, {self.ff_region}, ' \
               f'{self.country_region}, {self.city_town}, {self.organization_name},' \
               f'{self.organization_adress}, {self.product_group},{self.product_code}, ' \
               f'{self.item_quadra}, {self.organization_etalon_id}, {self.organization_etalon_name},' \
               f'{self.distributor_etalon_name},{self.distributor_name},{self.distributor_okpo},' \
               f'{self.sales_euro},{self.promotion},{self.organization_type},{self.organization_status},' \
               f'{self.etalon_code_okpo}, {self.delivery_date}, {self.position_code}, {self.office_head_organization}, ' \
               f'{self.head_office_okpo},{self.quarter_year}, {self.half_year},{self.annual_sales_category},' \
               f'{self.med_representative_name},{self.kam_name},{self.week},{self.territory_name},' \
               f'{self.brik_name}, {self.sales_packs}'
    def __repr__(self):
        return f'{self.year}, {self.month}, {self.ff_region}, {self.country_region}, {self.city_town}, {self.organization_name},{self.organization_adress}, {self.product_group},{self.product_code}, {self.item_quadra}, {self.organization_etalon_id}, {self.organization_etalon_name},{self.distributor_etalon_name},{self.distributor_name},{self.distributor_okpo},{self.sales_euro},{self.promotion},{self.organization_type},{self.organization_status},{self.etalon_code_okpo}, {self.delivery_date}, {self.position_code},{self.office_head_organization}, {self.head_office_okpo},{self.quarter_year}, {self.half_year},{self.annual_sales_category},{self.med_representative_name},{self.kam_name},{self.week},{self.territory_name},{self.brik_name}, {self.sales_packs}'
class CBase_2021_quadra_workout:
    def classify_base_2021(self,base_2021):
        base_2021_classifyed = []
        for i in base_2021:
            sale_in_quantity = i[32]
            brik_name = i[31]
            territory_name = i[30]
            week = i[29]
            kam_name = i[28]
            med_representative_name = i[27]
            annual_sales_category = i[26]
            half_year = i[25]
            quarter_year = i[24]
            head_office_okpo = i[23]
            office_head_organization = i[22]
            delivery_date = i[20]
            position_code = i[21]
            etalon_code_okpo = i[19]
            organization_status = i[18]
            organization_type = i[17]
            promotion = i[16]
            sales_euro_ = i[15]
            distributor_okpo = i[14]
            distributor_name = i[13]
            distributor_etalon_name = i[12]
            organization_etalon_name = i[11]
            organization_etalon_id = i[10]
            item_quadra = i[9]
            product_code = i[8]
            sales_method = i[7]
            organization_adress = i[6]
            organization_name = i[5]
            city_town = i[4]
            country_region = i[3]
            ff_region = i[2]
            month = i[1]
            year = i[0]
            entry = CSecondary_629_quadra_classify(year, month, ff_region, country_region, city_town, organization_name, organization_adress, sales_method,
                                                   product_code, item_quadra, organization_etalon_id, organization_etalon_name, distributor_etalon_name,
                                                   distributor_name, distributor_okpo, sales_euro_, promotion, organization_type, organization_status,
                                                   etalon_code_okpo, delivery_date, position_code, office_head_organization, head_office_okpo, quarter_year, half_year,
                                                   annual_sales_category, med_representative_name, kam_name, week, territory_name, brik_name, sale_in_quantity)
            base_2021_classifyed.append(entry)

        return base_2021_classifyed
    def classify_base_2021_from_xlxs(self, item):
        base_2021_classified_ = []
        year = str(item[0])
        month = str(item[1])
        ff_region = str(item[3])
        country_region = str(item[2])
        city_town = str(item[4])
        organization_name = str(item[5])
        organization_adress = str(item[6])
        product_group = str(item[7])
        product_code = str(item[8])
        item_quadra = str(item[9])
        organization_etalon_id = str(item[10])
        organization_etalon_name = str(item[11])
        distributor_etalon_name = str(item[12])
        distributor_name = str(item[13])
        distributor_okpo = str(item[14])
        sales_euro = str(item[15]).replace(',','.')
        promotion = str(item[16])
        organization_type = str(item[17])
        organization_status = str(item[18])
        etalon_code_okpo = str(item[19])
        delivery_date = str(item[20])
        position_code = str(item[21])
        office_head_organization = str(item[22])
        head_office_okpo = str(item[23])
        quarter_year = str(item[24])
        half_year = str(item[25])
        annual_sales_category = str(item[26])
        med_representative_name = str(item[27])
        kam_name = str(item[28])
        week = str(item[29])
        territory_name = str(item[30])
        brik_name = str(item[31])
        sales_packs = str(item[32]).replace(',','.')

        st = CSecondary_629_quadra_classify_from_excel(year, month, ff_region, country_region, city_town, organization_name,
                                                       organization_adress, product_group,
                                                       product_code, item_quadra, organization_etalon_id, organization_etalon_name,
                                                       distributor_etalon_name,
                                                       distributor_name, distributor_okpo, sales_euro, promotion, organization_type,
                                                       organization_status,
                                                       etalon_code_okpo, delivery_date, position_code, office_head_organization, head_office_okpo,
                                                       quarter_year, half_year,
                                                       annual_sales_category, med_representative_name, kam_name, week, territory_name,
                                                       brik_name, sales_packs)

        base_2021_classified_.append(st)

        return base_2021_classified_
    def upload_2021_base_from_quadra(self):
        try:
            # jTDS Driver
            driver_name = "net.sourceforge.jtds.jdbc.Driver"
            connection_url = connection_url_quadra_server
            logging.info("Connecting to Quadra Server - OK")
            connection_properties = quadra_login
            logging.info("Log in - OK")
            jar_path = jar_path_constant
            connection = jaydebeapi.connect(driver_name, connection_url, connection_properties, jar_path)
            cursor = connection.cursor()
            cursor.execute(sql_629_from_quadra_server)
            logging.info("Executing SQL request - OK")
            res = cursor.fetchall()
            x = CBase_2021_quadra_workout()
            base_2021_classifyed = x.classify_base_2021(res)
            logging.info("Classifying base from Quadra - OK")
            return base_2021_classifyed
        except Exception as err:
            print(str(err))
    def upload_2021_base_from_quadra_unclass(self):
        try:
            # jTDS Driver.
            driver_name = "net.sourceforge.jtds.jdbc.Driver"
            connection_url = connection_url_quadra_server
            logging.info("Connecting to Quadra Server - OK")
            connection_properties = quadra_login
            logging.info("Log in - OK")
            jar_path = jar_path_constant
            connection = jaydebeapi.connect(driver_name, connection_url, connection_properties, jar_path)
            cursor = connection.cursor()
            cursor.execute(sql_629_from_quadra_server)
            logging.info("Executing SQL request - OK")
            res = cursor.fetchall()

            return res
        except Exception as err:
            print(str(err))
    def calculate_rows_in_2021_base_from_quadra(self):
        try:
            count = 0
            # jTDS Driver.
            driver_name = "net.sourceforge.jtds.jdbc.Driver"
            connection_url = connection_url_quadra_server
            logging.info("Connecting to Quadra Server - OK")
            connection_properties = quadra_login
            logging.info("Log in - OK")
            jar_path = jar_path_constant
            connection = jaydebeapi.connect(driver_name, connection_url, connection_properties, jar_path)
            cursor = connection.cursor()
            cursor.execute(sql_629_from_quadra_server)
            logging.info("Executing SQL request - OK")
            res = cursor.fetchall()
            status = 'Already updated database'
            status_1 = 'Need to update xlsx and database'
            for i in res:
                count += 1
            print(count)
            logging.info(f"Getting rows count number of rows {count} - OK")

            path = 'C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\0.base_update_2021_row_count.xlsx'
            logging.info(f"Writing rows count - {count} - to '0.base_update_2021_row_count.xlsx' - OK")
            wb_obj = openpyxl.load_workbook(path)
            sheet_obj = wb_obj.active
            cell_obj = sheet_obj.cell(row=1, column=1)
            if cell_obj.value == count:
                return status
            else:
                workbook = xlsxwriter.Workbook(path)
                worksheet = workbook.add_worksheet()
                # Widen the first column to make the text clearer.
                # worksheet.set_column('A:A', 20)
                bold = workbook.add_format({'bold': True}, )
                worksheet.write('A1', count, bold)
                workbook.close()
                return status_1

        except Exception as err:
            print(str(err))
    def save_base_629_2021_to_xlsx(self):
        final_list = []

        # jTDS Driver.
        driver_name = "net.sourceforge.jtds.jdbc.Driver"
        connection_url = connection_url_quadra_server
        logging.info("Connecting to Quadra server  - OK")
        connection_properties = quadra_login
        logging.info("Log in to Quadra server  - OK")
        jar_path = jar_path_constant
        connection = jaydebeapi.connect(driver_name, connection_url, connection_properties, jar_path)
        cursor = connection.cursor()
        logging.info("Running 'sql_629_from_quadra_server'-script  - OK")
        cursor.execute(sql_629_from_quadra_server)
        res = cursor.fetchall()

        x = CBase_2021_quadra_workout()
        base_2021_classifyed_ = x.classify_base_2021(res)
        sum_tot_euro = 0
        workbook = xlsxwriter.Workbook('C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\0.new_629_report_2021.xlsx')
        logging.info("Opening '0.new_629_report_2021.xlsx' for writing  - OK")
        worksheet = workbook.add_worksheet()

        # Widen the first column to make the text clearer.
        #worksheet.set_column('A:A', 20)
        bold = workbook.add_format({'bold': True},)
        worksheet.write('A1', "Год", bold)
        worksheet.write('B1', "Месяц", bold)
        worksheet.write('C1', "Область", bold)
        worksheet.write('D1', "Регион", bold)
        worksheet.write('E1', "Населенный пункт", bold)
        worksheet.write('F1', "Организация", bold)
        worksheet.write('G1', "Почтовый адрес (организация)", bold)
        worksheet.write('H1', "Группа товара", bold)
        worksheet.write('I1', "product_code", bold)
        worksheet.write('J1', 'Товар', bold)
        worksheet.write('K1', 'Организация(эталонId)', bold)
        worksheet.write('L1', 'Организация(эталон)', bold)
        worksheet.write('M1', 'Дистрибьютор (эталон)', bold)
        worksheet.write('N1', 'Дистрибьютор', bold)
        worksheet.write('O1', 'ОКПО (дистрибъютор)', bold)
        worksheet.write('P1', 'Сумма (IN)', bold)
        worksheet.write('Q1', 'Группа товара 2', bold)
        worksheet.write('R1', 'Вид организации', bold)
        worksheet.write('S1', 'Тип организации', bold)
        worksheet.write('T1', 'ОКПО', bold)
        worksheet.write('U1', 'Дата отгрузки', bold)
        worksheet.write('V1', 'Код позиции', bold)
        worksheet.write('W1', 'Гл. Офис Сети', bold)
        worksheet.write('X1', 'Гл. Офис ОКПО', bold)
        worksheet.write('Y1', 'Квартал', bold)
        worksheet.write('Z1', 'Полугодие', bold)
        worksheet.write('AA1', 'Категория товарооборота', bold)
        worksheet.write('AB1', 'Сотрудник', bold)
        worksheet.write('AC1', 'КАМ', bold)
        worksheet.write('AD1', 'Неделя', bold)
        worksheet.write('AE1', 'Территория', bold)
        worksheet.write('AF1', 'Полигон', bold)
        worksheet.write('AG1', 'Количество (IN)', bold)
        logging.info("Adding headers to '0.new_629_report_2021.xlsx' - OK")

        columns = ['year', 'month', 'ff_region', 'country_region', 'city_town', 'organization_name',
                   'organization_adress', 'sales_method',
                   'product_code', 'item_quadra', 'organization_etalon_id', 'organization_etalon_name',
                   'distributor_etalon_name',
                   'distributor_name', 'distributor_okpo', 'sales_euro_', 'promotion', 'organization_type',
                   'organization_status',
                   'etalon_code_okpo', 'delivery_date', 'position_code','office_head_organization', 'head_office_okpo',
                   'quarter_year', 'half_year',
                   'annual_sales_category', 'med_representative_name', 'kam_name', 'week', 'territory_name',
                   'brik_name', 'sale_in_quantity']

        list_base_2021 = []
        row_index = 1

        for item in base_2021_classifyed_:
            z0 = str(item.year).replace(",", "")
            y1 = str(item.sales_euro_).replace(".", ",")
            y2 = str(item.sale_in_quantity).replace(".", ",")
            item_ = [[str(z0),
                     str(item.month),
                     str(item.ff_region),
                     str(item.country_region),
                     str(item.city_town),
                     str(item.organization_name),
                     str(item.organization_adress),
                     str(item.sales_method),
                     str(item.product_code),
                     str(item.item_quadra),
                     str(item.organization_etalon_id),
                     str(item.organization_etalon_name),
                     str(item.distributor_etalon_name),
                     str(item.distributor_name),
                     str(item.distributor_okpo),
                     str(y1),
                     str(item.promotion),
                     str(item.organization_type),
                     str(item.organization_status),
                     str(item.etalon_code_okpo),
                     str(item.delivery_date),
                     str(item.position_code),
                     str(item.office_head_organization),
                     str(item.head_office_okpo),
                     str(item.quarter_year),
                     str(item.half_year),
                     str(item.annual_sales_category),
                     str(item.med_representative_name),
                     str(item.kam_name),
                     str(item.week),
                     str(item.territory_name),
                     str(item.brik_name),
                     str(y2)]]

            list_base_2021.append(item_)
            worksheet.write(int(row_index), int(0),str(z0))
            worksheet.write(int(row_index), int(1),str(item.month))
            worksheet.write(int(row_index), int(2),str(item.ff_region))
            worksheet.write(int(row_index), int(3),str(item.country_region))
            worksheet.write(int(row_index), int(4),str(item.city_town))
            worksheet.write(int(row_index), int(5),str(item.organization_name))
            worksheet.write(int(row_index), int(6),str(item.organization_adress))
            worksheet.write(int(row_index), int(7),str(item.sales_method))
            worksheet.write(int(row_index), int(8),str(item.product_code))
            worksheet.write(int(row_index), int(9),str(item.item_quadra))
            worksheet.write(int(row_index), int(10),str(item.organization_etalon_id))
            worksheet.write(int(row_index), int(11),str(item.organization_etalon_name))
            worksheet.write(int(row_index), int(12),str(item.distributor_etalon_name))
            worksheet.write(int(row_index), int(13),str(item.distributor_name))
            worksheet.write(int(row_index), int(14),str(item.distributor_okpo))
            worksheet.write(int(row_index), int(15),str(y1))
            worksheet.write(int(row_index), int(16),str(item.promotion))
            worksheet.write(int(row_index), int(17),str(item.organization_type))
            worksheet.write(int(row_index), int(18),str(item.organization_status))
            worksheet.write(int(row_index), int(19),str(item.etalon_code_okpo))
            worksheet.write(int(row_index), int(20),str(item.delivery_date))
            worksheet.write(int(row_index), int(21),str(item.position_code))
            worksheet.write(int(row_index), int(22),str(item.office_head_organization))
            worksheet.write(int(row_index), int(23),str(item.head_office_okpo))
            worksheet.write(int(row_index), int(24),str(item.quarter_year))
            worksheet.write(int(row_index), int(25),str(item.half_year))
            worksheet.write(int(row_index), int(26),str(item.annual_sales_category))
            worksheet.write(int(row_index), int(27),str(item.med_representative_name))
            worksheet.write(int(row_index), int(28),str(item.kam_name))
            worksheet.write(int(row_index), int(29),str(item.week))
            worksheet.write(int(row_index), int(30),str(item.territory_name))
            worksheet.write(int(row_index), int(31),str(item.brik_name))
            worksheet.write(int(row_index), int(32),str(y2))

            row_index +=1
        logging.info("Writing data to '0.new_629_report_2021.xlsx' - OK")
        workbook.close()
    def save_base_629_2020_to_xlsx(self):
        final_list = []

        # jTDS Driver.
        driver_name = "net.sourceforge.jtds.jdbc.Driver"
        connection_url = connection_url_quadra_server
        connection_properties = quadra_login
        jar_path = jar_path_constant
        connection = jaydebeapi.connect(driver_name, connection_url, connection_properties, jar_path)
        cursor = connection.cursor()
        cursor.execute(sql_629_2020_from_quadra_server)
        res = cursor.fetchall()
        x = CBase_2021_quadra_workout()
        base_2020_classifyed_ = x.classify_base_2021(res)
        sum_tot_euro = 0
        workbook = xlsxwriter.Workbook('C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\0.new_629_report_2020.xlsx')
        worksheet = workbook.add_worksheet()
        # Widen the first column to make the text clearer.
        #worksheet.set_column('A:A', 20)
        bold = workbook.add_format({'bold': True},)
        worksheet.write('A1', "Год", bold)
        worksheet.write('B1', "Месяц", bold)
        worksheet.write('C1', "Область", bold)
        worksheet.write('D1', "Регион", bold)
        worksheet.write('E1', "Населенный пункт", bold)
        worksheet.write('F1', "Организация", bold)
        worksheet.write('G1', "Почтовый адрес (организация)", bold)
        worksheet.write('H1', "Группа товара", bold)
        worksheet.write('I1', "product_code", bold)
        worksheet.write('J1', 'Товар', bold)
        worksheet.write('K1', 'Организация(эталонId)', bold)
        worksheet.write('L1', 'Организация(эталон)', bold)
        worksheet.write('M1', 'Дистрибьютор (эталон)', bold)
        worksheet.write('N1', 'Дистрибьютор', bold)
        worksheet.write('O1', 'ОКПО (дистрибъютор)', bold)
        worksheet.write('P1', 'Сумма (IN)', bold)
        worksheet.write('Q1', 'Группа товара 2', bold)
        worksheet.write('R1', 'Вид организации', bold)
        worksheet.write('S1', 'Тип организации', bold)
        worksheet.write('T1', 'ОКПО', bold)
        worksheet.write('U1', 'Дата отгрузки', bold)
        worksheet.write('V1', 'Код позиции', bold)
        worksheet.write('W1', 'Гл. Офис Сети', bold)
        worksheet.write('X1', 'Гл. Офис ОКПО', bold)
        worksheet.write('Y1', 'Квартал', bold)
        worksheet.write('Z1', 'Полугодие', bold)
        worksheet.write('AA1', 'Категория товарооборота', bold)
        worksheet.write('AB1', 'Сотрудник', bold)
        worksheet.write('AC1', 'КАМ', bold)
        worksheet.write('AD1', 'Неделя', bold)
        worksheet.write('AE1', 'Территория', bold)
        worksheet.write('AF1', 'Полигон', bold)
        worksheet.write('AG1', 'Количество (IN)', bold)


        columns = ['year', 'month', 'ff_region', 'country_region', 'city_town', 'organization_name',
                   'organization_adress', 'sales_method',
                   'product_code', 'item_quadra', 'organization_etalon_id', 'organization_etalon_name',
                   'distributor_etalon_name',
                   'distributor_name', 'distributor_okpo', 'sales_euro_', 'promotion', 'organization_type',
                   'organization_status',
                   'etalon_code_okpo', 'delivery_date', 'position_code','office_head_organization', 'head_office_okpo',
                   'quarter_year', 'half_year',
                   'annual_sales_category', 'med_representative_name', 'kam_name', 'week', 'territory_name',
                   'brik_name', 'sale_in_quantity']

        list_base_2020 = []
        row_index = 1

        for item in base_2020_classifyed_:
            z0 = str(item.year).replace(",", "")
            y1 = str(item.sales_euro_).replace(".", ",")
            y2 = str(item.sale_in_quantity).replace(".", ",")
            item_ = [[str(z0),
                     str(item.month),
                     str(item.ff_region),
                     str(item.country_region),
                     str(item.city_town),
                     str(item.organization_name),
                     str(item.organization_adress),
                     str(item.sales_method),
                     str(item.product_code),
                     str(item.item_quadra),
                     str(item.organization_etalon_id),
                     str(item.organization_etalon_name),
                     str(item.distributor_etalon_name),
                     str(item.distributor_name),
                     str(item.distributor_okpo),
                     str(y1),
                     str(item.promotion),
                     str(item.organization_type),
                     str(item.organization_status),
                     str(item.etalon_code_okpo),
                     str(item.delivery_date),
                     str(item.position_code),
                     str(item.office_head_organization),
                     str(item.head_office_okpo),
                     str(item.quarter_year),
                     str(item.half_year),
                     str(item.annual_sales_category),
                     str(item.med_representative_name),
                     str(item.kam_name),
                     str(item.week),
                     str(item.territory_name),
                     str(item.brik_name),
                     str(y2)]]

            list_base_2020.append(item_)
            worksheet.write(int(row_index), int(0),str(z0))
            worksheet.write(int(row_index), int(1),str(item.month))
            worksheet.write(int(row_index), int(2),str(item.ff_region))
            worksheet.write(int(row_index), int(3),str(item.country_region))
            worksheet.write(int(row_index), int(4),str(item.city_town))
            worksheet.write(int(row_index), int(5),str(item.organization_name))
            worksheet.write(int(row_index), int(6),str(item.organization_adress))
            worksheet.write(int(row_index), int(7),str(item.sales_method))
            worksheet.write(int(row_index), int(8),str(item.product_code))
            worksheet.write(int(row_index), int(9),str(item.item_quadra))
            worksheet.write(int(row_index), int(10),str(item.organization_etalon_id))
            worksheet.write(int(row_index), int(11),str(item.organization_etalon_name))
            worksheet.write(int(row_index), int(12),str(item.distributor_etalon_name))
            worksheet.write(int(row_index), int(13),str(item.distributor_name))
            worksheet.write(int(row_index), int(14),str(item.distributor_okpo))
            worksheet.write(int(row_index), int(15),str(y1))
            worksheet.write(int(row_index), int(16),str(item.promotion))
            worksheet.write(int(row_index), int(17),str(item.organization_type))
            worksheet.write(int(row_index), int(18),str(item.organization_status))
            worksheet.write(int(row_index), int(19),str(item.etalon_code_okpo))
            worksheet.write(int(row_index), int(20),str(item.delivery_date))
            worksheet.write(int(row_index), int(21),str(item.position_code))
            worksheet.write(int(row_index), int(22),str(item.office_head_organization))
            worksheet.write(int(row_index), int(23),str(item.head_office_okpo))
            worksheet.write(int(row_index), int(24),str(item.quarter_year))
            worksheet.write(int(row_index), int(25),str(item.half_year))
            worksheet.write(int(row_index), int(26),str(item.annual_sales_category))
            worksheet.write(int(row_index), int(27),str(item.med_representative_name))
            worksheet.write(int(row_index), int(28),str(item.kam_name))
            worksheet.write(int(row_index), int(29),str(item.week))
            worksheet.write(int(row_index), int(30),str(item.territory_name))
            worksheet.write(int(row_index), int(31),str(item.brik_name))
            worksheet.write(int(row_index), int(32),str(y2))

            row_index +=1

        workbook.close()
    def get_secondary_2021_by_month(self):
        path = "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\0.new_629_report_2021.xlsx"
        logging.info("Connecting to '0.new_629_report_2021.xlsx'  - OK")
        wb_obj = openpyxl.load_workbook(path)
        sheet_obj = wb_obj.active
        rows_count = str(sheet_obj.calculate_dimension()).rsplit(':')
        rows_count = int(str(rows_count[1])[2:])
        logging.info(f"Calculating rows in '0.new_629_report_2021.xlsx' - number of rows - {rows_count}  - OK")
        string = []
        classified_base_2021 = []
        for row in range(1, rows_count + 1):
            str_ = []
            for col in range(1, 34):
                cell_obj = sheet_obj.cell(row=row, column=col)
                str_.append(cell_obj.value)
            string.append(str_)
        for i in string:
            x = CBase_2021_quadra_workout()
            string_class = x.classify_base_2021_from_xlxs(i)
            classified_base_2021.append(string_class)
        return classified_base_2021
    def get_secondary_2021_by_month_from_sqlite3(self,month):
        mtd_euro = 0
        mtd_packs = 0
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            logging.info("Connecting to 'local_main_base.db'  - OK")
            cursor = conn.cursor()
            cursor.execute(f"select secondary_2021_629.sales_packs, secondary_2021_629.sales_euro from secondary_2021_629 where secondary_2021_629.month = '{month}'")

            conn.commit()
            results = cursor.fetchall()
            for i in results:
                mtd_euro +=float(i[1])
                mtd_packs += float(i[0])
            logging.info(f"Getting secondary sales MTD data: MTD packs - {mtd_packs} packs, MTD euro - {mtd_euro} euro - OK")
        return mtd_packs, mtd_euro
    def get_secondary_2020_by_month(self):
        path = "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\0.new_629_report_2020.xlsx"
        wb_obj = openpyxl.load_workbook(path)
        sheet_obj = wb_obj.active
        rows_count = str(sheet_obj.calculate_dimension()).rsplit(':')
        rows_count = int(str(rows_count[1])[2:])
        string = []
        classified_base_2020 = []
        for row in range(1, rows_count + 1):
            str_ = []
            for col in range(1, 33):
                cell_obj = sheet_obj.cell(row=row, column=col)
                str_.append(cell_obj.value)
            string.append(str_)
        for i in string:
            x = CBase_2021_quadra_workout()
            string_class = x.classify_base_2021_from_xlxs(i)
            classified_base_2020.append(string_class)
        return classified_base_2020
    def rewrite_629_2021_in_database(conn):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            logging.info("Connecting to 'local_main_base.db'  - OK")
            cursor = conn.cursor()
            cursor.execute("DROP TABLE secondary_2021_629")
            logging.info("Dropping the table 'secondary_2021_629' in 'local_main_base.db'  - OK")
            cursor.execute("CREATE TABLE IF NOT EXISTS secondary_2021_629 (year, month, ff_region, country_region, city_town, organization_name,organization_adress, product_group,product_code, item_quadra, organization_etalon_id, organization_etalon_name,distributor_etalon_name, distributor_name, distributor_okpo, sales_euro, promotion, organization_type,organization_status, etalon_code_okpo, delivery_date, position_code, office_head_organization,head_office_okpo, quarter_year, half_year, annual_sales_category,med_representative_name, kam_name, week, territory_name, brik_name, sales_packs);")
            path = "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\0.new_629_report_2021.xlsx"
            logging.info("Creating the table 'secondary_2021_629' in 'local_main_base.db'  - OK")
            logging.info("Opening '0.new_629_report_2021.xlsx'   - OK")
            conn.commit()
            wb_obj = openpyxl.load_workbook(path)
            sheet_obj = wb_obj.active
            rows_count = str(sheet_obj.calculate_dimension()).rsplit(':')
            rows_count = int(str(rows_count[1])[2:])
            logging.info(f"Calculating rows in '0.new_629_report_2021.xlsx' - number of rows - {rows_count}  - OK")
            string = []
            classified_base_2021 = []
            for row in range(1, rows_count + 1):
                str_ = []
                for col in range(1, 34):
                    cell_obj = sheet_obj.cell(row=row, column=col)
                    str_.append(cell_obj.value)
                string.append(str_)
            logging.info("Classifying received data   - OK")
            for i in string:
                x = CBase_2021_quadra_workout()
                string_class = x.classify_base_2021_from_xlxs(i)
                classified_base_2021.append(string_class)
            for string in classified_base_2021:
                for row in string:
                    strin = [row.year,row.month,row.ff_region, row.country_region, row.city_town, row.organization_name,row.organization_adress, row.product_group, row.product_code, row.item_quadra, row.organization_etalon_id, row.organization_etalon_name,row.distributor_etalon_name,row.distributor_name, row.distributor_okpo, row.sales_euro, row.promotion, row.organization_type,row.organization_status,row.etalon_code_okpo, row.delivery_date, row.position_code, row.office_head_organization, row.head_office_okpo,row.quarter_year, row.half_year,row.annual_sales_category, row.med_representative_name,row.kam_name, row.week, row.territory_name,row.brik_name, row.sales_packs]
                    cursor.execute("INSERT INTO secondary_2021_629 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);",strin)
            conn.commit()
            logging.info("Inserting received data to 'local_main_base.db' - OK")

    def rewrite_629_2020_in_database(conn):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            #cursor.execute("DROP TABLE secondary_2020_629")
            cursor.execute("CREATE TABLE IF NOT EXISTS secondary_2020_629 (year, month, ff_region, country_region, city_town, organization_name,organization_adress, product_group,product_code, item_quadra, organization_etalon_id, organization_etalon_name,distributor_etalon_name, distributor_name, distributor_okpo, sales_euro, promotion, organization_type,organization_status, etalon_code_okpo, delivery_date, position_code, office_head_organization,head_office_okpo, quarter_year, half_year, annual_sales_category,med_representative_name, kam_name, week, territory_name, brik_name, sales_packs);")
            path = "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\0.new_629_report_2020.xlsx"
            conn.commit()
            wb_obj = openpyxl.load_workbook(path)
            sheet_obj = wb_obj.active
            rows_count = str(sheet_obj.calculate_dimension()).rsplit(':')
            rows_count = int(str(rows_count[1])[2:])
            string = []
            classified_base_2021 = []
            for row in range(1, rows_count + 1):
                str_ = []
                for col in range(1, 34):
                    cell_obj = sheet_obj.cell(row=row, column=col)
                    str_.append(cell_obj.value)
                string.append(str_)
            for i in string:
                x = CBase_2021_quadra_workout()
                string_class = x.classify_base_2021_from_xlxs(i)
                classified_base_2021.append(string_class)
            for string in classified_base_2021:
                for row in string:
                    strin = [row.year,row.month,row.ff_region, row.country_region, row.city_town, row.organization_name,row.organization_adress, row.product_group, row.product_code, row.item_quadra, row.organization_etalon_id, row.organization_etalon_name,row.distributor_etalon_name,row.distributor_name, row.distributor_okpo, row.sales_euro, row.promotion, row.organization_type,row.organization_status,row.etalon_code_okpo, row.delivery_date, row.position_code, row.office_head_organization, row.head_office_okpo,row.quarter_year, row.half_year,row.annual_sales_category, row.med_representative_name,row.kam_name, row.week, row.territory_name,row.brik_name, row.sales_packs]
                    cursor.execute("INSERT INTO secondary_2020_629 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);",strin)
            conn.commit()
            print('OK')
    def get_629_2021_from_sqlite3(conn):
        total_euro = 0
        total_packs = 0
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            logging.info("Connecting to 'local_main_base.db'  - OK")
            cursor = conn.cursor()
            cursor.execute("select sum(secondary_2021_629.sales_packs),sum(secondary_2021_629.sales_euro)  from secondary_2021_629")
            logging.info("Getting data from 'secondary_2021_629' table in 'local_main_base.db'  - OK")
            conn.commit()
            results = cursor.fetchall()
            for i in results:
                total_euro +=i[1]
                total_packs += i[0]
            print(total_packs)
            print(total_euro)
            logging.info(f"Total packs {total_packs} packs, Total euro {total_euro} euro - OK")
        return total_packs, total_euro
    def annual_plans_to_sqlite3_from_xlsx(conn):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            logging.info("Connecting to 'local_main_base.db'  - OK")
            cursor = conn.cursor()
            cursor.execute("DROP TABLE big_table_plans")
            logging.info("Dropping 'big_table_plans' table in 'local_main_base.db'  - OK")
            cursor.execute("CREATE TABLE IF NOT EXISTS big_table_plans (sales_type,general_name,brand,item_kpi_report,item_quadra,plan_fact,UoM,cip_euro_hq,month,month_local,distributor,value_final);")
            path = "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\raw_data_files\\secondary_sales_plans\\annual_sales_plan_from_HQ_final.xlsx"
            logging.info("Creating 'big_table_plans' table in 'local_main_base.db'  - OK")
            conn.commit()
            logging.info("Opening 'annual_sales_plan_from_HQ_final.xlsx'  - OK")
            wb_obj = openpyxl.load_workbook(path)
            sheet_obj = wb_obj.active
            rows_count = str(sheet_obj.calculate_dimension()).rsplit(':')
            print(rows_count)
            rows_count = int(str(rows_count[1])[1:])
            logging.info(f"Calculating rows - {rows_count}  - OK")
            string = []

            for row in range(2, rows_count + 1):
                str_ = []
                for col in range(1, 34):
                    cell_obj = sheet_obj.cell(row=row, column=col)
                    str_.append(cell_obj.value)
                string.append(str_)

            for row in string:
                x = str(row[11]).replace(".",",")
                strin = [row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],x]
                cursor.execute("INSERT INTO big_table_plans VALUES (?,?,?,?,?,?,?,?,?,?,?,?);",strin)
            conn.commit()
            logging.info("Inserting data to 'big_table_plans' table in 'local_main_base.db'  - OK")

    def plans_in_packs_from_sqlite3_to_xlsx_for_big_table(conn):

        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            logging.info("Connecting to 'local_main_base.db'  - OK")
            cursor = conn.cursor()
            cursor.execute("select big_table_plans.month, big_table_plans.distributor, big_table_plans.item_quadra, sum(big_table_plans.value_final) from big_table_plans where big_table_plans.UoM = 'packs' group by big_table_plans.month, big_table_plans.distributor, big_table_plans.item_quadra")
            conn.commit()
            results_packs = cursor.fetchall()
            workbook = xlsxwriter.Workbook(
                'C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\transform_files\\0.transform_for_big_table_plans_packs.xlsx')
            worksheet = workbook.add_worksheet()
            logging.info("Opening '0.transform_for_big_table_plans_packs.xlsx' for writing - OK")
            # Widen the first column to make the text clearer.
            # worksheet.set_column('A:A', 20)
            bold = workbook.add_format({'bold': True}, )
            worksheet.write('A1', "Месяц", bold)
            worksheet.write('B1', "Дистрибьютор", bold)
            worksheet.write('C1', "Товар", bold)
            worksheet.write('D1', "Количество (IN)", bold)

            list_base_2021 = []
            row_index = 1

            for item in results_packs:
                item_ = [[
                          str(item[0]),
                          str(item[1]),
                          str(item[2]),
                          str(item[3]).replace(".",",")]]

                list_base_2021.append(item_)
                worksheet.write(int(row_index), int(0), str(item[0]))
                worksheet.write(int(row_index), int(1), str(item[1]))
                worksheet.write(int(row_index), int(2), str(item[2]))
                worksheet.write(int(row_index), int(3), str(item[3]).replace(".",","))
                row_index += 1
            logging.info("Writing data to '0.transform_for_big_table_plans_packs.xlsx' - OK")
            workbook.close()
    def plans_in_euro_from_sqlite3_to_xlsx_for_big_table(conn):

        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            logging.info("Connecting to 'local_main_base.db'  - OK")
            cursor = conn.cursor()
            cursor.execute("select big_table_plans.month, big_table_plans.distributor, big_table_plans.item_quadra, sum(big_table_plans.value_final) from big_table_plans where big_table_plans.UoM = 'euro' group by big_table_plans.month, big_table_plans.distributor, big_table_plans.item_quadra")
            conn.commit()
            results_packs = cursor.fetchall()
            workbook = xlsxwriter.Workbook(
                'C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\transform_files\\0.transform_for_big_table_plans_euro.xlsx')
            logging.info("Opening '0.transform_for_big_table_plans_euro.xlsx'  - OK")
            worksheet = workbook.add_worksheet()

            # Widen the first column to make the text clearer.
            # worksheet.set_column('A:A', 20)
            bold = workbook.add_format({'bold': True}, )
            worksheet.write('A1', "Месяц", bold)
            worksheet.write('B1', "Дистрибьютор", bold)
            worksheet.write('C1', "Товар", bold)
            worksheet.write('D1', "Сумма (IN)", bold)

            list_base_2021 = []
            row_index = 1

            for item in results_packs:
                item_ = [[
                          str(item[0]),
                          str(item[1]),
                          str(item[2]),
                          str(item[3]).replace(".",",")]]

                list_base_2021.append(item_)
                worksheet.write(int(row_index), int(0), str(item[0]))
                worksheet.write(int(row_index), int(1), str(item[1]))
                worksheet.write(int(row_index), int(2), str(item[2]))
                worksheet.write(int(row_index), int(3), str(item[3]).replace(".",","))
                row_index += 1
            logging.info("Writing data to '0.transform_for_big_table_plans_euro.xlsx'  - OK")
            workbook.close()
    def actual_sales_to_sqlite3_from_xlsx(conn):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            logging.info("Connecting to 'local_main_base.db' - OK")
            cursor = conn.cursor()
            cursor.execute("DROP TABLE big_table_actual_sales")
            logging.info("Dropping the table 'big_table_actual_sales' in database - OK")
            cursor.execute("CREATE TABLE IF NOT EXISTS big_table_actual_sales (month, item_quadra, distributor_name, country_region,sales_euro,quarter_year, half_year, week,sales_packs);")
            path = "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\0.new_629_report_2021.xlsx"
            conn.commit()
            wb_obj = openpyxl.load_workbook(path)
            logging.info("Loading data from '0.new_629_report_2021.xlsx' - OK")
            sheet_obj = wb_obj.active
            rows_count = str(sheet_obj.calculate_dimension()).rsplit(':')
            rows_count = int(str(rows_count[1])[2:])
            logging.info(f"Calculating dimensions at '0.new_629_report_2021.xlsx' - number of rows: {rows_count} - OK")
            string = []
            classified_base_2021 = []
            for row in range(2, rows_count + 1):
                str_ = []
                for col in range(1, 34):
                    cell_obj = sheet_obj.cell(row=row, column=col)
                    str_.append(cell_obj.value)
                string.append(str_)
            for i in string:
                x = CBase_2021_quadra_workout()
                string_class = x.classify_base_2021_from_xlxs(i)
                classified_base_2021.append(string_class)
            for string in classified_base_2021:
                for row in string:
                    strin = [row.month,row.item_quadra,row.distributor_name,row.ff_region,row.sales_euro,row.quarter_year, row.half_year,row.week,row.sales_packs]
                    cursor.execute("INSERT INTO big_table_actual_sales VALUES (?,?,?,?,?,?,?,?,?);",strin)
            conn.commit()
            logging.info("Inserting data to 'big_table_actual_sales' table in database - OK")
            print('OK')
    def actual_sales_from_sqlite3_to_xlsx_for_big_table(conn):

        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            logging.info("Connecting to 'local_main_base.db' - OK")
            cursor = conn.cursor()
            cursor.execute("SELECT big_table_actual_sales.month, big_table_actual_sales.week, big_table_actual_sales.item_quadra, big_table_actual_sales.distributor_name,big_table_actual_sales.country_region, sum(big_table_actual_sales.sales_packs), sum(big_table_actual_sales.sales_euro) from big_table_actual_sales group by big_table_actual_sales.month,big_table_actual_sales.week, big_table_actual_sales.item_quadra, big_table_actual_sales.distributor_name,big_table_actual_sales.country_region")
            conn.commit()
            results = cursor.fetchall()
            workbook = xlsxwriter.Workbook(
                'C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\transform_files\\0.transform_for_big_table.xlsx')
            worksheet = workbook.add_worksheet()
            logging.info("Opening '0.transform_for_big_table.xlsx' - OK")

            # Widen the first column to make the text clearer.
            # worksheet.set_column('A:A', 20)
            bold = workbook.add_format({'bold': True}, )
            worksheet.write('A1', "Месяц", bold)
            worksheet.write('B1', "Неделя", bold)
            worksheet.write('C1', "Товар", bold)
            worksheet.write('D1', "Дистрибьютор", bold)
            worksheet.write('E1', "Область", bold)
            worksheet.write('F1', "Количество (IN)", bold)
            worksheet.write('G1', "Сумма (IN)", bold)
            logging.info("0.transform_for_big_table.xlsx writing headers - OK")

            list_base_2021 = []
            row_index = 1

            for item in results:
                item_ = [[
                          str(item[0]),
                          str(item[1]),
                          str(item[2]),
                          str(item[3]),
                          str(item[4]),
                          str(item[5]),
                          str(item[6])]]

                list_base_2021.append(item_)
                worksheet.write(int(row_index), int(0), str(item[0]))
                worksheet.write(int(row_index), int(1), str(item[1]))
                worksheet.write(int(row_index), int(2), str(item[2]))
                worksheet.write(int(row_index), int(3), str(item[3]))
                worksheet.write(int(row_index), int(4), str(item[4]))
                worksheet.write(int(row_index), int(5), str(item[5]).replace(".",","))
                worksheet.write(int(row_index), int(6), str(item[6]).replace(".", ","))
                row_index += 1
            logging.info("0.transform_for_big_table.xlsx writing data - OK")
            workbook.close()
    def save_1_tramsform_for_sales_report_with_filter_to_xlsx(self):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            logging.info("Connecting to 'local_main_base.db' - OK")
            cursor = conn.cursor()
            cursor.execute("select ymm.Год,ymm.half_month, items.cip_euro, items.brand, secondary_2021_629.month, secondary_2021_629.delivery_date, items.item_kpi_report, secondary_2021_629.sales_packs, secondary_2021_629.sales_euro, secondary_2021_629.distributor_name, secondary_2021_629.distributor_etalon_name from secondary_2021_629 join ymm on ymm.Дата = secondary_2021_629.delivery_date join items on items.item_quadra = secondary_2021_629.item_quadra and secondary_2021_629.month = items.month_ru where items.year = '2021'")
            conn.commit()
            results = cursor.fetchall()
        sum_tot_euro = 0
        workbook = xlsxwriter.Workbook('C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\transform_files\\0.transform_for_1_sales_report_with_filter.xlsx')
        worksheet = workbook.add_worksheet('BASE')
        logging.info("Opening '0.transform_for_1_sales_report_with_filter.xlsx' for writing - OK")

        # Widen the first column to make the text clearer.
        #worksheet.set_column('A:A', 20)
        bold = workbook.add_format({'bold': True},)
        worksheet.write('A1', "Год", bold)
        worksheet.write('B1', "Часть месяца", bold)
        worksheet.write('C1', "Актуальная CIP", bold)
        worksheet.write('D1', "Бренд", bold)
        worksheet.write('E1', "Месяц", bold)
        worksheet.write('F1', "Дата отгрузки", bold)
        worksheet.write('G1', "Товар", bold)
        worksheet.write('H1', "Количество упаковок", bold)
        worksheet.write('I1', "Сумма евро", bold)
        worksheet.write('J1', 'Дистрибьютор', bold)
        worksheet.write('K1', 'Дистрибьютор (эталон)', bold)

        list_base_2021 = []
        row_index = 1

        for item in results:
            item_ = [[str(item[0]),
                     str(item[1]),
                     str(item[2]).replace('_',' '),
                     str(item[3]),
                     str(item[4]),
                     str(item[5]),
                     str(item[6]).replace('.',','),
                     str(item[7]).replace('.',','),
                     str(item[8]),
                     str(item[9]),
                     str(item[10])]]

            list_base_2021.append(item_)
            worksheet.write(int(row_index), int(0),str(item[0]))
            worksheet.write(int(row_index), int(1),str(item[1]).replace('_',' '))
            worksheet.write(int(row_index), int(2), str(item[2]))
            worksheet.write(int(row_index), int(3),str(item[3]))
            worksheet.write(int(row_index), int(4),str(item[4]))
            worksheet.write(int(row_index), int(5),str(item[5]).replace('.',','))
            worksheet.write(int(row_index), int(6),str(item[6]))
            worksheet.write(int(row_index), int(7),str(item[7]).replace('.',','))
            worksheet.write(int(row_index), int(8),str(item[8]).replace('.',','))
            worksheet.write(int(row_index), int(9), str(item[9]))
            worksheet.write(int(row_index), int(10), str(item[10]))
            row_index +=1
        logging.info("Writing data to '0.transform_for_1_sales_report_with_filter.xlsx' - OK")
        workbook.close()
    def get_secondary_sales_sqlite3_for_big_table(conn,selected_months,year):
        sql_ = f"select secondary_{year}_629.month,  sum(secondary_{year}_629.sales_packs), sum(secondary_{year}_629.sales_euro)from secondary_{year}_629 where secondary_{year}_629.month NOT LIKE 'Месяц'  group by secondary_{year}_629.month"
        total_euro = 0
        total_packs = 0
        list_ = []
        with sqlite3.connect(
                "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(sql_)
            conn.commit()
            results = cursor.fetchall()

            for month in selected_months:

                for i in results:

                    if month == i[0]:
                        total_euro = 0
                        total_packs = 0
                        total_euro += i[2]
                        total_packs += i[1]
                        list_.append([month, total_packs, total_euro])
            print(total_packs)
            print(total_euro)
        return list_
    def get_secondary_sales_sqlite3_for_big_table_by_item(conn,selected_months,year,item_selected):
        sql_ = f"select secondary_{year}_629.month,  sum(secondary_{year}_629.sales_packs), sum(secondary_{year}_629.sales_euro) from secondary_{year}_629 where secondary_{year}_629.month NOT LIKE 'Месяц' and secondary_{year}_629.item_quadra = '{item_selected}' group by  secondary_{year}_629.month, secondary_{year}_629.item_quadra"
        list_ = []
        with sqlite3.connect(
                "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(sql_)
            conn.commit()
            results = cursor.fetchall()

            for month in selected_months:

                for i in results:

                    if month == i[0]:
                        total_euro = 0
                        total_packs = 0

                        total_euro += i[2]
                        total_packs += i[1]
                        list_.append([month, total_packs, total_euro])
        return list_


class CTest_SAles_report_classification:
    def __init__(self,year,month,week,distr,item_quadra,sec_quantity,sec_euro):
        self.sec_euro = sec_euro
        self.sec_quantity = sec_quantity
        self.item_quadra = item_quadra
        self.distr = distr
        self.week = week
        self.month = month
        self.year = year
class CTest_SAles_report_creation:
    def test_rep(self,base_raw):
        years = []
        months = []
        items = []
        distributor_name = []
        weeks = []
        for i in base_raw:
            if i.year not in years:
                years.append(i.year)
            if i.month not in months:
                months.append(i.month)
            if i.item_quadra not in items:
                items.append(i.item_quadra)
            if i.distributor_name not in distributor_name:
                distributor_name.append(i.distributor_name)
            if i.week not in weeks:
                weeks.append(i.week)
        sales_by_product = []
        for i in base_raw:
            for year in years:
                for month in months:
                    for week in weeks:
                        for distr in distributor_name:
                            for item in items:
                                if i.year == year and i.month == month and i.week == week and i.distributor_name == distr and i.item_quadra == item:
                                    sales_by_product.append(CTest_SAles_report_classification(i.year, i.month,i.week,i.distributor_name,i.item_quadra, i.sale_in_quantity, i.sales_euro_))
        return sales_by_product



    def print_actual_MTD_sales(base_raw,year,month_local):
        report_test = CTest_SAles_report_creation()
        report = report_test.test_rep(base_raw)
        sales_euro = 0
        week = ''
        distr = ''
        item_quadra = ''
        quantity = 0
        for i in report:
            if i.year == year and i.month == month_local:
                sales_euro += i.sec_euro
                quantity += i.sec_quantity
        print(f"Год: {year}\nМесяц: {month_local}\nОбщие продажи в упаковках:", '{0:,}'.format(quantity.__round__(2)).replace(",", " "), "packs\nОбщие продажи в евро:", '{0:,}'.format(sales_euro.__round__(2)).replace(",", " "), 'euro')
class CEXtract_database_tertiary:
    def read_item(conn, year):
        tertiary_list = []
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year}")
            results = cursor.fetchall()
            for i in results:
                y_1 = str(i[4]).replace(',', '.')
                y_2 = str(i[5]).replace(',', '.')
                y_3 = str(i[6]).replace(',', '.')
                y_4 = str(i[7]).replace(',', '.')
                y_5 = str(i[8]).replace(',', '.')
                y_6 = str(i[9]).replace(',', '.')
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)
                tertiary_list.append(z)
        return tertiary_list
    def read_item_2020_w_commas(conn,year):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year}")
            results = cursor.fetchall()
            tertiary_list = []
            for i in results:
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8],i[9])
                tertiary_list.append(z)
        return tertiary_list
    def save_2020_items_to_csv(self, filename, list_months,year):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year}")
            results = cursor.fetchall()
            tertiary_list = []
            for i in results:
                y_1 = str(i[4]).replace(',', '.')
                y_2 = str(i[5]).replace(',', '.')
                y_3 = str(i[6]).replace(',', '.')
                y_4 = str(i[7]).replace(',', '.')
                y_5 = str(i[8]).replace(',', '.')
                y_6 = str(i[9]).replace(',', '.')
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)
                tertiary_list.append(z)
        final_list = []
        for month in list_months:
            for i in tertiary_list:
                if month == i.month:
                    final_list.append([{"year": str(i.year), "month": i.month, "brand": i.brand, "item": i.item_kpi,
                                        "weight_penetration": str(i.weight_pen),
                                        "sro": str(i.sro), "penetration": str(i.pen),
                                        "quantity": str(i.quantity), "amount_euro": str(i.amount_euro),"weighted_sro":str(i.weight_sro)}])
        with open(filename, "w", newline="", encoding='utf-16') as file:
            columns = ["year","month","brand","item","weight_penetration","sro","penetration","quantity","amount_euro","weighted_sro"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()
            for item in final_list:
                writer.writerows(item)
    def save_items_2020_to_csv_with_commas(self, filename, list_months,year):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year}")
            results = cursor.fetchall()
            tertiary_list = []
            for i in results:

                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8],i[9])
                tertiary_list.append(z)
        final_list = []
        for month in list_months:
            for i in tertiary_list:
                if month == i.month:
                    final_list.append([{"year": str(i.year), "month": i.month, "brand": i.brand, "item": i.item_kpi,
                                        "weight_penetration": str(i.weight_pen),
                                        "sro": str(i.sro), "penetration": str(i.pen),
                                        "quantity": str(i.quantity), "amount_euro": str(i.amount_euro),"weighted_sro":str(i.weight_sro)}])
        with open(filename, "w", newline="", encoding='utf-16') as file:
            columns = ["year","month","brand","item","weight_penetration","sro","penetration","quantity","amount_euro","weighted_sro"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()
            for item in final_list:
                writer.writerows(item)
    def read_item_2020_OTC(conn,year):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year} and items.sales_method = 'OTC'")
            results = cursor.fetchall()
            tertiary_list_otc = []
            for i in results:
                y_1 = str(i[4]).replace(',', '.')
                y_2 = str(i[5]).replace(',', '.')
                y_3 = str(i[6]).replace(',', '.')
                y_4 = str(i[7]).replace(',', '.')
                y_5 = str(i[8]).replace(',', '.')
                y_6 = str(i[9]).replace(',', '.')
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)

                tertiary_list_otc.append(z)
        return tertiary_list_otc
    def save_items_otc_to_csv_2020(self, filename,list_months,year):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year} and items.sales_method = 'OTC'")
            results = cursor.fetchall()
            tertiary_list = []
            for i in results:
                y_1 = str(i[4]).replace(',', '.')
                y_2 = str(i[5]).replace(',', '.')
                y_3 = str(i[6]).replace(',', '.')
                y_4 = str(i[7]).replace(',', '.')
                y_5 = str(i[8]).replace(',', '.')
                y_6 = str(i[9]).replace(',', '.')
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)
                tertiary_list.append(z)
        final_list = []
        for month in list_months:
            for i in tertiary_list:
                if month == i.month:
                    final_list.append([{"year": str(i.year), "month": i.month, "brand": i.brand, "item": i.item_kpi,
                                        "weight_penetration": str(i.weight_pen),
                                        "sro": str(i.sro), "penetration": str(i.pen),
                                        "quantity": str(i.quantity), "amount_euro": str(i.amount_euro),"weighted_sro":str(i.weight_sro)}])
        with open(filename, "w", newline="", encoding='utf-16') as file:
            columns = ["year","month","brand","item","weight_penetration","sro","penetration","quantity","amount_euro","weighted_sro"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()
            for item in final_list:
                writer.writerows(item)
    def read_item_2020_OTC_with_commas(conn,year):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year} and items.sales_method = 'OTC'")
            results = cursor.fetchall()
            tertiary_list_otc = []
            for i in results:
                y_1 = str(i[4])
                y_2 = str(i[5])
                y_3 = str(i[6])
                y_4 = str(i[7])
                y_5 = str(i[8])
                y_6 = str(i[9])
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)

                tertiary_list_otc.append(z)
        return tertiary_list_otc
    def save_items_otc_to_csv_2020_with_commas(self, filename,list_months,year):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year} and items.sales_method = 'OTC'")
            results = cursor.fetchall()
            tertiary_list = []
            for i in results:
                y_1 = str(i[4])
                y_2 = str(i[5])
                y_3 = str(i[6])
                y_4 = str(i[7])
                y_5 = str(i[8])
                y_6 = str(i[9])
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)
                tertiary_list.append(z)
        final_list = []
        for month in list_months:
            for i in tertiary_list:
                if month == i.month:
                    final_list.append([{"year": str(i.year), "month": i.month, "brand": i.brand, "item": i.item_kpi,
                                        "weight_penetration": str(i.weight_pen),
                                        "sro": str(i.sro), "penetration": str(i.pen),
                                        "quantity": str(i.quantity), "amount_euro": str(i.amount_euro),"weighted_sro":str(i.weight_sro)}])
        with open(filename, "w", newline="", encoding='utf-16') as file:
            columns = ["year","month","brand","item","weight_penetration","sro","penetration","quantity","amount_euro","weighted_sro"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()
            for item in final_list:
                writer.writerows(item)
    def read_item_2020_RX_with_commas(conn,year):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year} and items.sales_method = 'RX'")
            results = cursor.fetchall()
            tertiary_list_otc = []
            for i in results:
                y_1 = str(i[4])
                y_2 = str(i[5])
                y_3 = str(i[6])
                y_4 = str(i[7])
                y_5 = str(i[8])
                y_6 = str(i[9])
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)

                tertiary_list_otc.append(z)
        return tertiary_list_otc
    def save_items_RX_to_csv_2020_with_commas(self, filename,list_months,year):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year} and items.sales_method = 'RX'")
            results = cursor.fetchall()
            tertiary_list = []
            for i in results:
                y_1 = str(i[4])
                y_2 = str(i[5])
                y_3 = str(i[6])
                y_4 = str(i[7])
                y_5 = str(i[8])
                y_6 = str(i[9])
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)

                tertiary_list.append(z)
        final_list = []
        for month in list_months:
            for i in tertiary_list:
                if month == i.month:
                    final_list.append([{"year": str(i.year), "month": i.month, "brand": i.brand, "item": i.item_kpi,
                                        "weight_penetration": str(i.weight_pen),
                                        "sro": str(i.sro), "penetration": str(i.pen),
                                        "quantity": str(i.quantity), "amount_euro": str(i.amount_euro),"weighted_sro":str(i.weight_sro)}])
        with open(filename, "w", newline="", encoding='utf-16') as file:
            columns = ["year","month","brand","item","weight_penetration","sro","penetration","quantity","amount_euro","weighted_sro"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()
            for item in final_list:
                writer.writerows(item)
    def read_item_2020_RX(conn,year):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year} and items.sales_method = 'RX'")
            results = cursor.fetchall()
            tertiary_list_otc = []
            for i in results:
                y_1 = str(i[4]).replace(',', '.')
                y_2 = str(i[5]).replace(',', '.')
                y_3 = str(i[6]).replace(',', '.')
                y_4 = str(i[7]).replace(',', '.')
                y_5 = str(i[8]).replace(',', '.')
                y_6 = str(i[9]).replace(',', '.')
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)

                tertiary_list_otc.append(z)
        return tertiary_list_otc
    def save_items_RX_to_csv_2020(self, filename,list_months,year):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume, tertiary_sales.WeightSRO from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = {year} and items.sales_method = 'RX'")
            results = cursor.fetchall()
            tertiary_list = []
            for i in results:
                y_1 = str(i[4]).replace(',', '.')
                y_2 = str(i[5]).replace(',', '.')
                y_3 = str(i[6]).replace(',', '.')
                y_4 = str(i[7]).replace(',', '.')
                y_5 = str(i[8]).replace(',', '.')
                y_6 = str(i[9]).replace(',', '.')
                z = Tertiary_download_structure(i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5,y_6)
                tertiary_list.append(z)
        final_list = []
        for month in list_months:
            for i in tertiary_list:
                if month == i.month:
                    final_list.append([{"year": str(i.year), "month": i.month, "brand": i.brand, "item": i.item_kpi,
                                        "weight_penetration": str(i.weight_pen),
                                        "sro": str(i.sro), "penetration": str(i.pen),
                                        "quantity": str(i.quantity), "amount_euro": str(i.amount_euro),"weighted_sro":str(i.weight_sro)}])
        with open(filename, "w", newline="", encoding='utf-16') as file:
            columns = ["year","month","brand","item","weight_penetration","sro","penetration","quantity","amount_euro","weighted_sro"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()
            for item in final_list:
                writer.writerows(item)
    def test_secondary_2021(self,month_ru):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"select secondary_sales_2021.item_quadra, secondary_sales_2021.cip_fact from secondary_sales_2021 where secondary_sales_2021.month_ru = '{month_ru}'")
            results = cursor.fetchall()
            secondary_list = []
            for i in results:
                y_1 = str(i[0]).replace(',', '.')
                y_2 = str(i[1]).replace(',', '.')
                z = Secondary_total_2021(y_1, y_2)
                secondary_list.append(z)
            sum = 0
            for i in secondary_list:
                sum += float(i.sales_euro)
            return sum
    def test_secondary_2020(self,month_ru):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"select secondary_sales_2020.item_quadra, secondary_sales_2020.cip_fact from secondary_sales_2020 where secondary_sales_2020.month_ru = '{month_ru}'")
            results = cursor.fetchall()
            secondary_list = []
            for i in results:
                y_1 = str(i[0]).replace(',', '.')
                y_2 = str(i[1]).replace(',', '.')
                z = Secondary_total_2021(y_1, y_2)
                secondary_list.append(z)
            sum = 0
            for i in secondary_list:
                sum += float(i.sales_euro)
            return sum
