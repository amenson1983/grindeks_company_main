import csv
import tkinter
from datetime import date
from tkinter import Tk


from items_GUI import Items_GUI
from sale_out.database import CEXtract_database_tertiary, CBase_2021_quadra_workout, \
    Upload_2021_base_from_quadra_for_daily_totals_distr, CTest_SAles_report_creation
from sale_out.items_class import SKU_WORKOUT


def Main():
    root = Tk()
    ex = Items_GUI()
    root.geometry("600x550+400+50")
    main_menu = tkinter.Menu()
    file_menu = tkinter.Menu()
    save_menu = tkinter.Menu()
    tertiary_reports_2020 = tkinter.Menu()

    tertiary_reports_2020.add_command(label="Save month-sold_packs data to JSON", command=ex.save_quant_month_to_json)
    tertiary_reports_2020.add_command(label="Save month-weighted_penetration data to JSON",
                          command=ex.save_weight_pen_month_to_json)
    tertiary_reports_2020.add_command(label="Save month-weighted_penetration data to CSV",
                          command=ex.save_weight_pen_month_to_csv)
    tertiary_reports_2020.add_command(label="Save month-weighted_penetration-item data to CSV",
                          command=ex.save_weight_pen_month_item_to_csv)
    tertiary_reports_2020.add_command(label="Save relevant data PROMO to CSV",
                          command=ex.save_all_data_PROMO_items_to_csv)
    tertiary_reports_2020.add_command(label="Save relevant data PROMO to CSV with commas",
                          command=ex.save_all_data_PROMO_items_to_csv_commas)

    tertiary_reports_2020.add_command(label="Save OTC relevant data to CSV",
                          command=ex.save_otc_PROMO_items_to_csv)
    tertiary_reports_2020.add_command(label="Save OTC relevant data to CSV with commas",
                          command=ex.save_otc_PROMO_items_to_csv_with_commas)
    tertiary_reports_2020.add_command(label="Save RX relevant data to CSV with commas",
                          command=ex.save_rx_items_to_csv_with_commas)
    tertiary_reports_2020.add_command(label="Save RX relevant data to CSV",
                          command=ex.save_rx_items_to_csv)


    file_menu.add_command(label="New")
    file_menu.add_cascade(label="Tertiary reports 2021", menu=save_menu)
    file_menu.add_cascade(label="Tertiary reports 2020", menu=tertiary_reports_2020)
    file_menu.add_separator()
    file_menu.add_command(label="Exit", command=root.destroy)

    main_menu.add_cascade(label="File", menu=file_menu)
    main_menu.add_cascade(label="Edit")
    main_menu.add_cascade(label="View")

    root.config(menu=main_menu)
    root.mainloop()

def replace_commas(items_2_pos):
    items = []
    for i in items_2_pos:
        x = str(i[1]).replace(',','.')
        z = [i[0],x]
        items.append(z)
    return items
def actual_secondary_sales_2021():
    x = CEXtract_database_tertiary()
    month_list = ['Январь','Февраль','Март','Апрель', 'Май','Июнь','Июль', 'Август','Сентябрь','Октябрь', 'Ноябрь','Декабрь']
    for month in month_list:
        y = x.test_secondary_2021(month)
        print(f'Год: 2021; Месяц: {month}: \t\t\t\t', '{0:,}'.format(y.__round__(2)).replace(",", " "), 'euro')

def actual_secondary_sales_2020():
    x = CEXtract_database_tertiary()
    month_list = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь',
                  'Ноябрь', 'Декабрь']
    for month in month_list:
        y = x.test_secondary_2020(month)
        print(f'Год: 2020; Месяц: {month}: \t\t\t\t', '{0:,}'.format(y.__round__(2)).replace(",", " "), 'euro')

if __name__ == '__main__':
    #Main()
    #actual_secondary_sales_2020()
    #actual_secondary_sales_2021()
    x = CBase_2021_quadra_workout()
    #base_raw = x.upload_2021_base_from_quadra()
    #y = CTest_SAles_report_creation
    #month = date.today().month
    #if month == 2:
    #    month = 'Февраль'
    #y.print_actual_MTD_sales(base_raw, 2021, month)

    #base_2021_classifyed = x.save_base_2021_to_csv()
    x.save_base_629_2021_to_xlsx()






