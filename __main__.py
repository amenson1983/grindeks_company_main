import csv
import json
import sqlite3
import tkinter
from tkinter import Tk, BOTH, Listbox, StringVar, END, messagebox, Checkbutton,font
from tkinter.ttk import Frame, Label
import matplotlib
import matplotlib.pyplot as plt

from items_GUI import Items_GUI
from items_class import SKU_WORKOUT, CItemsDAO
from new_try.database import CEXtract_database_tertiary
from tertiary_sales_class import Tertiary_sales
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
                          command=ex.save_all_data_2020_PROMO_items_to_csv)
    tertiary_reports_2020.add_command(label="Save relevant data PROMO to CSV with commas",
                          command=ex.save_all_data_2020_PROMO_items_to_csv_commas)

    tertiary_reports_2020.add_command(label="Save OTC relevant data to CSV",
                          command=ex.save_otc_2020_PROMO_items_to_csv)
    tertiary_reports_2020.add_command(label="Save OTC relevant data to CSV with commas",
                          command=ex.save_otc_2020_PROMO_items_to_csv_with_commas)
    tertiary_reports_2020.add_command(label="Save RX relevant data to CSV with commas",
                          command=ex.save_rx_2020_items_to_csv_with_commas)
    tertiary_reports_2020.add_command(label="Save RX relevant data to CSV",
                          command=ex.save_rx_2020_items_to_csv)


    file_menu.add_command(label="New")
    file_menu.add_cascade(label="Save", menu=save_menu)
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

if __name__ == '__main__':
    Main()











