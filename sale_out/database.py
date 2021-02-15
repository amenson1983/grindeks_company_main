import csv
import sqlite3
from sale_out.tertiary_sales_class import Tertiary_sales
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
                z = Tertiary_download_structure([i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8],i[9]])
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
        with open(filename, "w", newline="", encoding='UTF') as file:
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
        with open(filename, "w", newline="", encoding='UTF') as file:
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
        with open(filename, "w", newline="", encoding='UTF') as file:
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
        with open(filename, "w", newline="", encoding='UTF') as file:
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
        with open(filename, "w", newline="", encoding='UTF') as file:
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
        with open(filename, "w", newline="", encoding='UTF') as file:
            columns = ["year","month","brand","item","weight_penetration","sro","penetration","quantity","amount_euro","weighted_sro"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()
            for item in final_list:
                writer.writerows(item)

