import sqlite3

from tertiary_sales_class import Tertiary_sales


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
                entry_to_list = CSKU(year,sales_method,promotion,purpose,item_proxima,item_quadra,item_sales_report,item_kpi_report,brand,month,cip_euro)
                x = str(entry_to_list.cip_euro).replace(',','.')
                entry_new = CSKU(year,sales_method,promotion,purpose,item_proxima,item_quadra,item_sales_report,item_kpi_report,brand,month,x)
                list_items_2020.append([{"year": entry_new.year, "sales_method":entry_new.sales_method,"promotion":entry_new.promotion,"purpose":entry_new.purpose,"item_proxima":entry_new.item_proxima,"item_quadra":entry_new.item_quadra,"item_sales_report":entry_new.item_sales_report,"item_kpi_report":entry_new.item_kpi_report,"brand":entry_new.brand,"month":entry_new.month,"cip_euro":entry_new.cip_euro}])
        return list_items_2020
    def read_item_2021_local_item(conn):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * from items where items.year = '2021'")
            results = cursor.fetchall()
            list_items_2021 = []
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
                entry_to_list = CSKU(year,sales_method,promotion,purpose,item_proxima,item_quadra,item_sales_report,item_kpi_report,brand,month,cip_euro)
                x = str(entry_to_list.cip_euro).replace(',','.')
                entry_new = CSKU(year,sales_method,promotion,purpose,item_proxima,item_quadra,item_sales_report,item_kpi_report,brand,month,x)
                list_items_2021.append([{"year": entry_new.year, "sales_method":entry_new.sales_method,"promotion":entry_new.promotion,"purpose":entry_new.purpose,"item_proxima":entry_new.item_proxima,"item_quadra":entry_new.item_quadra,"item_sales_report":entry_new.item_sales_report,"item_kpi_report":entry_new.item_kpi_report,"brand":entry_new.brand,"month":entry_new.month,"cip_euro":entry_new.cip_euro}])
        return list_items_2021



class CItemsDAO:
    def insert_item(conn, sales: Tertiary_sales):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            # cursor.execute("INSERT INTO Tertiary_sales VALUES (276,'Djamala');")
            # cursor.execute(f"INSERT INTO Tertiary_sales VALUES (?,?)",(artist.item,f'{artist.brand}'))
            cursor.execute(f"INSERT INTO sales VALUES (?,?,?,?,?,?,?,?)", (f'({sales.year}', (f'{sales.month}'), (f'{sales.item}'), (f'{sales.weight_penetration}'), (f'{sales.sro}'), (f'{sales.quantity}'), (f'{sales.volume_euro}')))

            conn.commit()

    def update_item(conn, sales: Tertiary_sales):
        pass
        with mySql.connect(
                host="localhost",
                user="root",
                password="root",
                database="chinook"
        ) as conn:
            cursor = conn.cursor()
            # cursor.execute("UPDATE Tertiary_sales SET Name = 'TNMK' WHERE ArtistId = 276;")
            cursor.execute(f"UPDATE Tertiary_sales SET Name = '{items.brand}' WHERE ItemsId = {items.item};")
            conn.commit()

    def delete_item(conn):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM sales WHERE sales.item_id = '(90'")
            #cursor.execute(f"DELETE FROM Tertiary_sales WHERE ItemsId = {sales.item}")
            conn.commit()

    def read_tertiary(conn):
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT ymm.Year, ymm.Month, items.item_kpi_report, items.brand, tertiary_sales.WeightPenetration,tertiary_sales.WeightSRO , tertiary_sales.Quantity, tertiary_sales.Volume from tertiary_sales join ymm on tertiary_sales.Period = ymm.Year_monnum JOIN items on tertiary_sales.Fullmedicationname = items.item_proxima where tertiary_sales.MarketOrg = 'Grindeks  (Latvia)' and tertiary_sales.year = '2020'")
            results = cursor.fetchall()
            tertiary_list = []
            for i in results:
                y = str(i[4]).replace(',', '.')
                z = (i[0], i[1], i[2], i[3], y, i[5] , i[6], i[7])
                tertiary_list.append(z)
        return tertiary_list
