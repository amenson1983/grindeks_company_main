import logging
import sqlite3

from pandas.tests.io.excel.test_xlsxwriter import xlsxwriter

from sale_out.database import Kam_plan_download_structure, Tertiary_by_region_download_structure, Tertiary_workout

logging.basicConfig(filename='bi.log', level=logging.INFO, format='%(asctime)s %(message)s', datefmt='%d/%m/%Y %I:%M:%S %p')


class CPower_BI:
    def actual_secondary_sales_from_sqlite3(conn):
        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            logging.info("Connecting to 'local_main_base.db' - OK")
            cursor = conn.cursor()
            cursor.execute("SELECT secondary_2021_629.month, secondary_2021_629.delivery_date, secondary_2021_629.week, secondary_2021_629.item_quadra, secondary_2021_629.distributor_name, secondary_2021_629.ff_region, secondary_2021_629.office_head_organization, secondary_2021_629.position_code, secondary_2021_629.kam_name, sum(secondary_2021_629.sales_packs), sum(secondary_2021_629.sales_euro) from secondary_2021_629 group by secondary_2021_629.month, secondary_2021_629.delivery_date, secondary_2021_629.week, secondary_2021_629.item_quadra, secondary_2021_629.distributor_name,secondary_2021_629.ff_region,secondary_2021_629.office_head_organization, secondary_2021_629.position_code,secondary_2021_629.kam_name")
            conn.commit()
            results = cursor.fetchall()
            workbook = xlsxwriter.Workbook(
                'C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\power_bi_package\\0.secondary_sales_2020.xlsx')
            worksheet = workbook.add_worksheet()
            logging.info("Opening - OK")

            # Widen the first column to make the text clearer.
            # worksheet.set_column('A:A', 20)
            bold = workbook.add_format({'bold': True}, )
            worksheet.write('A1', "Месяц", bold)
            worksheet.write('B1', "Дата", bold)
            worksheet.write('C1', "Неделя", bold)
            worksheet.write('D1', "Товар", bold)
            worksheet.write('E1', "Дистрибьютор", bold)
            worksheet.write('F1', "Область", bold)
            worksheet.write('G1', "Главный Офис Сети", bold)
            worksheet.write('H1', "Код сотрудника", bold)
            worksheet.write('I1', "KAM", bold)
            worksheet.write('J1', "Упаковки (IN)", bold)
            worksheet.write('K1', "Сумма (IN)", bold)
            logging.info("Writing headers - OK")

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
                          str(item[6]),
                          str(item[7]),
                          str(item[8]),
                          str(item[9]),
                          str(item[10])]]

                list_base_2021.append(item_)
                worksheet.write(int(row_index), int(0), str(item[0]))
                worksheet.write(int(row_index), int(1), str(item[1]))
                worksheet.write(int(row_index), int(2), str(item[2]))
                worksheet.write(int(row_index), int(3), str(item[3]))
                worksheet.write(int(row_index), int(4), str(item[4]))
                worksheet.write(int(row_index), int(5), str(item[5]))
                worksheet.write(int(row_index), int(6), str(item[6]))
                worksheet.write(int(row_index), int(7), str(item[7]))
                worksheet.write(int(row_index), int(8), str(item[8]))
                worksheet.write(int(row_index), int(9), str(item[9]).replace(".",","))
                worksheet.write(int(row_index), int(10), str(item[10]).replace(".", ","))
                row_index += 1
            logging.info("Writing data - OK")
            workbook.close()
    def distinct_head_offices(conn):

        with sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            logging.info("Connecting to 'local_main_base.db' - OK")
            cursor = conn.cursor()
            cursor.execute("SELECT distinct secondary_2021_629.office_head_organization from secondary_2021_629")
            conn.commit()
            results = cursor.fetchall()
            workbook = xlsxwriter.Workbook(
                'C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\power_bi_package\\0.distinct_head_offices.xlsx')
            worksheet = workbook.add_worksheet()
            logging.info("Opening  - OK")

            # Widen the first column to make the text clearer.
            # worksheet.set_column('A:A', 20)
            bold = workbook.add_format({'bold': True}, )
            worksheet.write('A1', "Главный офис сети", bold)


            list_base_2021 = []
            row_index = 1

            for item in results:
                item_ = [[
                          str(item[0])]]

                list_base_2021.append(item_)
                worksheet.write(int(row_index), int(0), str(item[0]))

                row_index += 1
            logging.info("Writing data - OK")
            workbook.close()

    def secondary_sales_2020_from_sqlite3_to_transform_xlxs(conn):
        with sqlite3.connect(
                "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            logging.info("Connecting to 'local_main_base.db' - OK")
            cursor = conn.cursor()
            cursor.execute(
                "SELECT secondary_2020_629.month, secondary_2020_629.delivery_date, secondary_2020_629.week, secondary_2020_629.item_quadra, secondary_2020_629.distributor_name, secondary_2020_629.ff_region, secondary_2020_629.office_head_organization, secondary_2020_629.position_code, secondary_2020_629.kam_name, sum(secondary_2020_629.sales_packs), sum(secondary_2020_629.sales_euro) from secondary_2020_629 group by secondary_2020_629.month, secondary_2020_629.delivery_date, secondary_2020_629.week, secondary_2020_629.item_quadra, secondary_2020_629.distributor_name,secondary_2020_629.ff_region,secondary_2020_629.office_head_organization, secondary_2020_629.position_code,secondary_2020_629.kam_name")
            conn.commit()
            results = cursor.fetchall()
            workbook = xlsxwriter.Workbook(
                'C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\power_bi_package\\0.secondary_sales_2020.xlsx')
            worksheet = workbook.add_worksheet()
            logging.info("Opening - OK")

            # Widen the first column to make the text clearer.
            # worksheet.set_column('A:A', 20)
            bold = workbook.add_format({'bold': True}, )
            worksheet.write('A1', "Месяц", bold)
            worksheet.write('B1', "Дата", bold)
            worksheet.write('C1', "Неделя", bold)
            worksheet.write('D1', "Товар", bold)
            worksheet.write('E1', "Дистрибьютор", bold)
            worksheet.write('F1', "Область", bold)
            worksheet.write('G1', "Главный Офис Сети", bold)
            worksheet.write('H1', "Код сотрудника", bold)
            worksheet.write('I1', "KAM", bold)
            worksheet.write('J1', "Упаковки (IN)", bold)
            worksheet.write('K1', "Сумма (IN)", bold)
            logging.info("Writing headers - OK")

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
                    str(item[6]),
                    str(item[7]),
                    str(item[8]),
                    str(item[9]),
                    str(item[10])]]

                list_base_2021.append(item_)
                worksheet.write(int(row_index), int(0), str(item[0]))
                worksheet.write(int(row_index), int(1), str(item[1]))
                worksheet.write(int(row_index), int(2), str(item[2]))
                worksheet.write(int(row_index), int(3), str(item[3]))
                worksheet.write(int(row_index), int(4), str(item[4]))
                worksheet.write(int(row_index), int(5), str(item[5]))
                worksheet.write(int(row_index), int(6), str(item[6]))
                worksheet.write(int(row_index), int(7), str(item[7]))
                worksheet.write(int(row_index), int(8), str(item[8]))
                worksheet.write(int(row_index), int(9), str(item[9]).replace(".", ","))
                worksheet.write(int(row_index), int(10), str(item[10]).replace(".", ","))
                row_index += 1
            logging.info("Writing data - OK")
            workbook.close()

    def kam_plan(conn):
        plan_list = []
        with sqlite3.connect(
                "C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"select distinct kam_plan.item_quadra, kam_plan.code, kam_plan.month,  ymm.Квартал, kam_plan.plan_packs, kam_plan.plan_euro from kam_plan join ymm on ymm.Месяц = kam_plan.month")
            results = cursor.fetchall()
            for i in results:
                y_1 = str(i[0])
                y_2 = str(i[1])
                y_3 = str(i[2])
                y_4 = str(i[3])
                y_5 = str(i[4])
                y_6 = str(i[5]).replace(',', '.')
                z = Kam_plan_download_structure(y_1, y_2, y_3, y_4, y_5, y_6)
                plan_list.append(z)
        workbook = xlsxwriter.Workbook(
            'C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\power_bi_package\\0.kam_plans_2021.xlsx')
        worksheet = workbook.add_worksheet()
        logging.info("Opening - OK")

        # Widen the first column to make the text clearer.
        # worksheet.set_column('A:A', 20)
        bold = workbook.add_format({'bold': True}, )
        worksheet.write('A1', "Товар", bold)
        worksheet.write('B1', "Код", bold)
        worksheet.write('C1', "Месяц", bold)
        worksheet.write('D1', "Квартал", bold)
        worksheet.write('E1', "Упаковки", bold)
        worksheet.write('F1', "Сумма", bold)

        logging.info("Writing headers - OK")

        list_base_2021 = []
        row_index = 1

        for item in results:
            item_ = [[
                str(item[0]),
                str(item[1]),
                str(item[2]),
                str(item[3]),
                str(item[4]),
                str(item[5])]]

            list_base_2021.append(item_)
            worksheet.write(int(row_index), int(0), str(item[0]))
            worksheet.write(int(row_index), int(1), str(item[1]))
            worksheet.write(int(row_index), int(2), str(item[2]))
            worksheet.write(int(row_index), int(3), str(item[3]))
            worksheet.write(int(row_index), int(4), str(item[4]).replace(".", ","))
            worksheet.write(int(row_index), int(5), str(item[5]).replace(".", ","))
            row_index += 1
        logging.info("Writing data - OK")
        workbook.close()




if __name__ == '__main__':
    #ex = CPower_BI()
    #ex.distinct_head_offices()
    #ex.actual_secondary_sales_from_sqlite3()
    #ex.secondary_sales_2020_from_sqlite3_to_transform_xlxs()
    #ex.kam_plan()
    x = Tertiary_workout()
    x.tert_reg_to_sqlite()