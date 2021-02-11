import csv
import json
import sqlite3
import tkinter
from tkinter import Tk, BOTH, Listbox, StringVar, END, messagebox, Checkbutton,font
from tkinter.ttk import Frame, Label
import matplotlib
import matplotlib.pyplot as plt

from items_GUI import Items_GUI
from items_class import SKU_WORKOUT


def list_work():
    root = Tk()
    ex = Items_GUI()
    root.geometry("450x350")
    root.mainloop()

if __name__ == '__main__':
    r = SKU_WORKOUT()
    list_2020_items = r.read_item_2020()
    r.save_items_to_csv(list_2020_items)
    list_work()
