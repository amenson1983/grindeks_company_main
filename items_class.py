import csv
import sqlite3


class CSKU:
    def __init__(self,year,sales_method,promotion,purpose,item_proxima,item_quadra,item_sales_report,item_kpi_report,brand,month,cip_euro):
        self.year = year
        self.sales_method = sales_method
        self.promotion = promotion
        self.purpose = purpose
        self.item_proxima = item_proxima
        self.item_quadra = item_quadra
        self.item_sales_report = item_sales_report
        self.item_kpi_report = item_kpi_report
        self.brand = brand
        self.month = month
        self.cip_euro = cip_euro


    def __str__(self):
        return f"{self.year},{self.sales_method},{self.promotion},{self.purpose},{self.item_proxima},{self.item_quadra},{self.item_sales_report},{self.item_kpi_report},{self.brand},{self.month},{self.cip_euro}"
    def __repr__(self):
        return f"{self.year},{self.sales_method},{self.promotion},{self.purpose},{self.item_proxima},{self.item_quadra},{self.item_sales_report},{self.item_kpi_report},{self.brand},{self.month},{self.cip_euro}"

class SKU_WORKOUT():
    def read_item_2020(conn):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * from items where items.year = '2020'")
            results = cursor.fetchall()
            list_items_2020 = []
            for entry in results:
                year = entry[0]
                sales_method= entry[1]
                promotion= entry[2]
                purpose = entry[3]
                item_proxima = entry[4]
                item_quadra= entry[5]
                item_sales_report= entry[6]
                item_kpi_report= entry[7]
                brand= entry[8]
                month= entry[9]
                cip_euro= entry[10]
                entry_new = CSKU(year,sales_method,promotion,purpose,item_proxima,item_quadra,item_sales_report,item_kpi_report,brand,month,cip_euro)
                list_items_2020.append([{"year": entry_new.year, "sales_method":entry_new.sales_method,"promotion":entry_new.promotion,"purpose":entry_new.purpose,"item_proxima":entry_new.item_proxima,"item_quadra":entry_new.item_quadra,"item_sales_report":entry_new.item_sales_report,"item_kpi_report":entry_new.item_kpi_report,"brand":entry_new.brand,"month":entry_new.month,"cip_euro":entry_new.cip_euro}])
        return list_items_2020

    def save_items_2020_to_csv(self,list_items_2020):
        FILENAME = "items_2020.csv"
        with open(FILENAME, "w", newline="",encoding='UTF') as file:
            columns = ["year","sales_method","promotion","purpose","item_proxima","item_quadra","item_sales_report","item_kpi_report","brand","month","cip_euro"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()
            for item in list_items_2020:
            # запись нескольких строк
                writer.writerows(item)


if __name__ == '__main__':
    r = SKU_WORKOUT()
    list_2020_items = r.read_item_2020()
    r.save_items_2020_to_csv(list_2020_items)