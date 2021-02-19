import sqlite3


class Plan_ff_main_2021_download:
    def __init__(self,month_local,item_quadra,med_rep_code,plan_packs,plan_euro):
        self.month_local = month_local
        self.item_quadra = item_quadra
        self.med_rep_code = med_rep_code
        self.plan_packs = plan_packs
        self.plan_euro = plan_euro








class CEXtract_database_plan:
    def read_plan_ff_main(conn):
        rep_plan_list = []
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute(f"select distinct ff_main_plan_2021.month_local, ff_main_plan_2021.item_quadra, ff_main_plan_2021.med_representative_code, ff_main_plan_2021.packs_plan, ff_main_plan_2021.euro_plan from ff_main_plan_2021 where ff_main_plan_2021.year = '2021' and ff_main_plan_2021.euro_plan <> 0")
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

#class CEXtract_database_actual:
if __name__ == '__main__':
    x = CEXtract_database_plan()
    y = x.read_plan_ff_main()
    annual_ff_plan_euro = 0
    distinct_codes = []
    for i in y:
        annual_ff_plan_euro += float(i.plan_euro)
        if i.med_rep_code not in distinct_codes:
            distinct_codes.append(i.med_rep_code)

    print(annual_ff_plan_euro, distinct_codes)