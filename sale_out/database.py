import csv
import sqlite3

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

class Secondary_total_2021:
    def __init__(self,item_quadra,sales_euro):
        self.sales_euro = sales_euro
        self.item_quadra = item_quadra

class Quadra_direct_629():
    def __init__(self,year,month,ff_region,country_region,city_town,organization_name,organization_adress,sales_method,
                 product_code,item_quadra,organization_etalon_id,organization_etalon_name,distributor_etalon_name,
                 distributor_name,distributor_okpo,sales_euro_,promotion,organization_type,organization_status,
                 etalon_code_okpo,delivery_date,office_head_organization,head_office_okpo,quarter_year,half_year,
                 annual_sales_category,med_representative_name,kam_name,week,territory_name,brik_name,sale_in_quantity):
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





class CEXtract_database_tertiary:
    def read_item(conn, year):
        tertiary_list = []
        with sqlite3.connect("tertiary_sales_database.db") as conn:
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
        with sqlite3.connect("tertiary_sales_database.db") as conn:
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
        with sqlite3.connect("tertiary_sales_database.db") as conn:
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
        with sqlite3.connect("tertiary_sales_database.db") as conn:
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
        with sqlite3.connect("tertiary_sales_database.db") as conn:
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
        with sqlite3.connect("tertiary_sales_database.db") as conn:
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
        with sqlite3.connect("tertiary_sales_database.db") as conn:
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
        with sqlite3.connect("tertiary_sales_database.db") as conn:
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
        with sqlite3.connect("tertiary_sales_database.db") as conn:
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
        with sqlite3.connect("tertiary_sales_database.db") as conn:
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
        with sqlite3.connect("tertiary_sales_database.db") as conn:
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
        with sqlite3.connect("tertiary_sales_database.db") as conn:
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
        with sqlite3.connect("tertiary_sales_database.db") as conn:
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
        with sqlite3.connect("tertiary_sales_database.db") as conn:
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

