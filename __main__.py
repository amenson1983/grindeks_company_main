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


def list_work():
    root = Tk()
    ex = Items_GUI()
    root.geometry("450x350")
    root.mainloop()

def replace_commas(items_2_pos):
    items = []
    for i in items_2_pos:
        x = str(i[1]).replace(',','.')
        z = [i[0],x]
        items.append(z)
    return items

if __name__ == '__main__':
    list_work()











