import csv
import sqlite3

from tertiary_sales_class import Tertiary_sales


class CEXtract_database_tertiary:
    def read_item_2020(conn):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = '2020'")
            results = cursor.fetchall()
            tertiary_list = []
            for i in results:
                y_1 = str(i[4]).replace(',', '.')
                y_2 = str(i[5]).replace(',', '.')
                y_3 = str(i[6]).replace(',', '.')
                y_4 = str(i[7]).replace(',', '.')
                y_5 = str(i[8]).replace(',', '.')
                z = ([i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5])
                tertiary_list.append(z)
        return tertiary_list
    def save_items_to_csv(self, filename, list_months):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT DISTINCT items.item_kpi_report, tertiary_sales.Year, tertiary_sales.PeriodName, items.brand, tertiary_sales.WeightPenetration,  tertiary_sales.SRO, tertiary_sales.Penetration, tertiary_sales.Quantity, tertiary_sales.Volume from items join tertiary_sales on tertiary_sales.Fullmedicationname = items.item_proxima WHERE tertiary_sales.Year = '2020'")
            results = cursor.fetchall()
            tertiary_list = []
            for i in results:
                y_1 = str(i[4]).replace(',', '.')
                y_2 = str(i[5]).replace(',', '.')
                y_3 = str(i[6]).replace(',', '.')
                y_4 = str(i[7]).replace(',', '.')
                y_5 = str(i[8]).replace(',', '.')
                z = ([i[0], i[1], i[2], i[3], y_1, y_2, y_3, y_4, y_5])
                entry = Tertiary_sales(z)
                tertiary_list.append(entry)
        final_list = []
        for month in list_months:
            for i in tertiary_list:
                if month == i.month:
                    final_list.append([{"year": str(i.year), "month": i.month, "brand": i.brand, "item": i.item,
                                        "weight_penetration": str(i.weight_penetration),
                                        "weight_sro": str(i.weight_sro), "penetration": str(i.penetration),
                                        "quantity": str(i.quantity), "amount_euro": str(i.volume_euro)}])
        with open(filename, "w", newline="", encoding='UTF') as file:
            columns = ["year","month","brand","item","weight_penetration","weight_sro","penetration","quantity","amount_euro"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()
            for item in final_list:
                writer.writerows(item)




