import csv
import json
import sqlite3
import tkinter
from tkinter import Tk, BOTH, Listbox, StringVar, END, messagebox, Checkbutton,font
from tkinter.ttk import Frame, Label
import matplotlib.pyplot as plt

from new_try.database import CEXtract_database_tertiary
from tertiary_sales_class import Tertiary_sales


class Data_GUI(Frame):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.master.title( "Tertiary sales by item")
        self.pack(fill=BOTH, expand=1)
        self.top_frame = tkinter.Frame(self.master)
        self.button_frame = tkinter.Frame(self.master)
        self.left_frame = tkinter.Frame()

        self.check_var1 = tkinter.IntVar()
        self.check_var1.set(0)
        self.check_var2 = tkinter.IntVar()
        self.check_var2.set(0)
        self.check_var3 = tkinter.IntVar()
        self.check_var3.set(0)
        self.check_var4 = tkinter.IntVar()
        self.check_var4.set(0)
        self.check_var5 = tkinter.IntVar()
        self.check_var5.set(0)
        self.check_var6 = tkinter.IntVar()
        self.check_var6.set(0)
        self.check_var7 = tkinter.IntVar()
        self.check_var7.set(0)
        self.check_var8 = tkinter.IntVar()
        self.check_var8.set(0)
        self.check_var9 = tkinter.IntVar()
        self.check_var9.set(0)
        self.check_var10 = tkinter.IntVar()
        self.check_var10.set(0)
        self.check_var11 = tkinter.IntVar()
        self.check_var11.set(0)
        self.check_var12 = tkinter.IntVar()
        self.check_var12.set(0)
        my_font = tkinter.font.Font(family='Arial', size=14, weight='bold')
        my_font1 = tkinter.font.Font(family='Arial', size=11)
        self.chb1 = tkinter.Checkbutton(self.left_frame, text='Январь', variable=self.check_var1,
                                        font=my_font1)


        self.chb2 = tkinter.Checkbutton(self.left_frame, text='Февраль', variable=self.check_var2,
                                        font=my_font1)
        self.chb3 = tkinter.Checkbutton(self.left_frame, text='Март', variable=self.check_var3,
                                        font=my_font1)
        self.chb4 = tkinter.Checkbutton(self.left_frame, text='Апрель',
                                        variable=self.check_var4, font=my_font1)
        self.chb5 = tkinter.Checkbutton(self.left_frame, text='Май', variable=self.check_var5, font=my_font1)
        self.chb6 = tkinter.Checkbutton(self.left_frame, text='Июнь', variable=self.check_var6,
                                        font=my_font1)
        self.chb7 = tkinter.Checkbutton(self.top_frame, text='Июль', variable=self.check_var7,
                                        font=my_font1)
        self.chb8 = tkinter.Checkbutton(self.top_frame, text='Август', variable=self.check_var8,
                                        font=my_font1)
        self.chb9 = tkinter.Checkbutton(self.top_frame, text='Сентябрь',
                                        variable=self.check_var9, font=my_font1)
        self.chb10 = tkinter.Checkbutton(self.top_frame, text='Октябрь', variable=self.check_var10, font=my_font1)
        self.chb11 = tkinter.Checkbutton(self.top_frame, text='Ноябрь', variable=self.check_var11,
                                        font=my_font1)

        self.chb12 = tkinter.Checkbutton(self.top_frame, text='Декабрь', variable=self.check_var12,
                                        font=my_font1,padx=15, pady=-10)
        ss = CEXtract_database_tertiary()
        pos = ss.read_item_2020()
        for i in pos:
            z = Tertiary_sales(i)
            print(z.weight_penetration)

        acts = items
        self.show_button = tkinter.Button(self.button_frame, text='Weighted penetration', command=self.show_renetration)
        self.show_button.pack(side='left')
        self.ok_button = tkinter.Button(self.button_frame, text='Sales in euro', command=self.onclick_euro)
        self.ok_button.pack(side='left')
        self.quit_button = tkinter.Button(self.button_frame, text='Quit', command=self.master.destroy)
        self.quit_button.pack(side='left')
        self.button_frame.pack()
        lb = Listbox(self,width='70',height='15')
        self.ok_button_quantity = tkinter.Button(self.button_frame, text='Sales in packs', command=self.onclick_quantity)
        self.ok_button_quantity.pack(side='left')

        i_list = []
        for i in acts:
            i_list.append(i.item)
        for item in i_list:
            lb.insert(END, item)

        lb.bind("<<ListboxSelect>>", self.onSelect)

        lb.pack(pady=15)

        self.var = StringVar()
        self.label = Label(self, text=0, textvariable=self.var)
        self.label.pack()
        self.info_var = StringVar()
        self.info_label = Label(self, text=0, textvariable=self.info_var)
        self.info_label.pack()

        self.chb1.pack(side='left')
        self.chb2.pack(side='left')
        self.chb3.pack(side='left')
        self.chb4.pack(side='left')
        self.chb5.pack(side='left')
        self.chb6.pack(side='left')
        self.chb7.pack(side='left')
        self.chb8.pack(side='left')
        self.chb9.pack(side='left')
        self.chb10.pack(side='left')
        self.chb11.pack(side='left')
        self.chb12.pack(side='left')
        self.top_frame.pack(side='bottom')
        self.left_frame.pack(side='top')



    def show_renetration(self):
        self.month = ''
        list = []
        if self.check_var1.get() == 1:
            self.month = 'Январь'
            list.append(self.month)
        if self.check_var2.get() == 1:
            self.month = 'Февраль'
            list.append(self.month)
        if self.check_var3.get() == 1:
            self.month = 'Март'
            list.append(self.month)
        if self.check_var4.get() == 1:
            self.month = 'Апрель'
            list.append(self.month)
        if self.check_var5.get() == 1:
            self.month = 'Май'
            list.append(self.month)
        if self.check_var6.get() == 1:
            self.month = 'Июнь'
            list.append(self.month)
        if self.check_var7.get() == 1:
            self.month = 'Июль'
            list.append(self.month)
        if self.check_var8.get() == 1:
            self.month = 'Август'
            list.append(self.month)
        if self.check_var9.get() == 1:
            self.month = 'Сентябрь'
            list.append(self.month)
        if self.check_var10.get() == 1:
            self.month = 'Октябрь'
            list.append(self.month)
        if self.check_var11.get() == 1:
            self.month = 'Ноябрь'
            list.append(self.month)
        if self.check_var12.get() == 1:
            self.month = 'Декабрь'
            list.append(self.month)
        x_coord = list
        list_2 = []
        for l in list:
            for i in items:
                if i.item == self.info_var.get() and i.year == 2020 and i.month == l:
                    x = str(i.weight_penetration)
                    x = x.replace(',', '.')
                    list_2.append(float(x))
        y_coord = []
        for en in list_2:
            y_coord.append(en)
        plt.title(f'Пенетрация по месяцам по выбранному SKU: \n{self.info_var.get()}')
        plt.grid(True)
        plt.plot(x_coord, y_coord,marker='s')
        plt.show()
    def onSelect(self, val):
        sender = val.widget
        idx = sender.curselection()
        value = sender.get(idx)
        x = value
        self.info_var.set(value)

        self.var.set(f'Здесь может быть еще какое-нибудь сообщение')
    def onclick_euro(self):
        self.month = ''
        self.amount_euro = 0
        list = []
        if self.check_var1.get() == 1:
            self.month = 'Январь'
            list.append(self.month)
        if self.check_var2.get() == 1:
            self.month = 'Февраль'
            list.append(self.month)
        if self.check_var3.get() == 1:
            self.month = 'Март'
            list.append(self.month)
        if self.check_var4.get() == 1:
            self.month = 'Апрель'
            list.append(self.month)
        if self.check_var5.get() == 1:
            self.month = 'Май'
            list.append(self.month)
        if self.check_var6.get() == 1:
            self.month = 'Июнь'
            list.append(self.month)
        if self.check_var7.get() == 1:
            self.month = 'Июль'
            list.append(self.month)
        if self.check_var8.get() == 1:
            self.month = 'Август'
            list.append(self.month)
        if self.check_var9.get() == 1:
            self.month = 'Сентябрь'
            list.append(self.month)
        if self.check_var10.get() == 1:
            self.month = 'Октябрь'
            list.append(self.month)
        if self.check_var11.get() == 1:
            self.month = 'Ноябрь'
            list.append(self.month)
        if self.check_var12.get() == 1:
            self.month = 'Декабрь'
            list.append(self.month)
        x_coord = list
        list_2 = []

        for l in list:
            for i in items:
                if i.item == self.info_var.get() and i.year == 2020 and i.month == l:
                    x = str(i.volume_euro)
                    x = x.replace(',','.')
                    list_2.append(float(x))
        y_coord = []
        for en in list_2:
            self.amount_euro += en
            y_coord.append(en)
        plt.title(f'Третичные продажи в евро по месяцам по SKU: \n{self.info_var.get()}')
        plt.grid(True)
        plt.plot(x_coord,y_coord,marker='s')
        plt.show()
        tkinter.messagebox.showinfo('INFO',
                                    f'Sales in euro: {self.amount_euro} euro')
    def save_quant_month_to_json(self):
        FILENAME = f"{self.info_var.get()}_month_quantity.json"
        self.month = ''
        self.quantity = 0
        list = []
        if self.check_var1.get() == 1:
            self.month = 'Январь'
            list.append(self.month)
        if self.check_var2.get() == 1:
            self.month = 'Февраль'
            list.append(self.month)
        if self.check_var3.get() == 1:
            self.month = 'Март'
            list.append(self.month)
        if self.check_var4.get() == 1:
            self.month = 'Апрель'
            list.append(self.month)
        if self.check_var5.get() == 1:
            self.month = 'Май'
            list.append(self.month)
        if self.check_var6.get() == 1:
            self.month = 'Июнь'
            list.append(self.month)
        if self.check_var7.get() == 1:
            self.month = 'Июль'
            list.append(self.month)
        if self.check_var8.get() == 1:
            self.month = 'Август'
            list.append(self.month)
        if self.check_var9.get() == 1:
            self.month = 'Сентябрь'
            list.append(self.month)
        if self.check_var10.get() == 1:
            self.month = 'Октябрь'
            list.append(self.month)
        if self.check_var11.get() == 1:
            self.month = 'Ноябрь'
            list.append(self.month)
        if self.check_var12.get() == 1:
            self.month = 'Декабрь'
            list.append(self.month)
        list_2 = []
        x_coord = list
        y_coord = []
        for l in list:
            for i in items:
                if i.item == self.info_var.get() and i.year == 2020 and i.month == l:
                    x = str(i.quantity)
                    x = x.replace(',','.')
                    list_2.append(float(x))

        for en in list_2:
            self.quantity += en
            y_coord.append(en)
        users = []
        for num in range(0,len(list)):
                users.append({"month": str(list[num]),
                    "quantity_packs": str(list_2[num])})

        strData = json.dumps(users)
        with open(FILENAME, "w", encoding='utf-8') as file:
            file.write(strData)
            tkinter.messagebox.showinfo('INFO',
                                        f'File {FILENAME} has been succesfully written!')
    def save_weight_pen_month_to_json(self):
        FILENAME = f"{self.info_var.get()}_month_weight_pen.json"
        self.month = ''
        self.quantity = 0
        list = []
        if self.check_var1.get() == 1:
            self.month = 'Январь'
            list.append(self.month)
        if self.check_var2.get() == 1:
            self.month = 'Февраль'
            list.append(self.month)
        if self.check_var3.get() == 1:
            self.month = 'Март'
            list.append(self.month)
        if self.check_var4.get() == 1:
            self.month = 'Апрель'
            list.append(self.month)
        if self.check_var5.get() == 1:
            self.month = 'Май'
            list.append(self.month)
        if self.check_var6.get() == 1:
            self.month = 'Июнь'
            list.append(self.month)
        if self.check_var7.get() == 1:
            self.month = 'Июль'
            list.append(self.month)
        if self.check_var8.get() == 1:
            self.month = 'Август'
            list.append(self.month)
        if self.check_var9.get() == 1:
            self.month = 'Сентябрь'
            list.append(self.month)
        if self.check_var10.get() == 1:
            self.month = 'Октябрь'
            list.append(self.month)
        if self.check_var11.get() == 1:
            self.month = 'Ноябрь'
            list.append(self.month)
        if self.check_var12.get() == 1:
            self.month = 'Декабрь'
            list.append(self.month)
        list_2 = []
        x_coord = list
        y_coord = []
        for l in list:
            for i in items:
                if i.item == self.info_var.get() and i.year == 2020 and i.month == l:
                    x = str(i.weight_penetration)
                    x = x.replace(',','.')
                    list_2.append(float(x))

        for en in list_2:
            self.quantity += en
            y_coord.append(en)
        users = []
        for num in range(0,len(list)):
                users.append({"month": str(list[num]),
                    "weighted_penetration": str(list_2[num])})

        strData = json.dumps(users)
        with open(FILENAME, "w") as file:
            file.write(strData)
            tkinter.messagebox.showinfo('INFO',
                                        f'File {FILENAME} has been succesfully written!')
    def get_dict_from_file_1(self):
        FILENAME = f"{self.info_var.get()}_month_quantity.json"
        with open(FILENAME, "r", encoding="UTF") as myfile:
            user_str = myfile.read()
        user_dict = json.loads(user_str)
        return user_dict
    def onclick_quantity(self):
        self.month = ''
        self.quantity = 0
        list = []
        if self.check_var1.get() == 1:
            self.month = 'Январь'
            list.append(self.month)
        if self.check_var2.get() == 1:
            self.month = 'Февраль'
            list.append(self.month)
        if self.check_var3.get() == 1:
            self.month = 'Март'
            list.append(self.month)
        if self.check_var4.get() == 1:
            self.month = 'Апрель'
            list.append(self.month)
        if self.check_var5.get() == 1:
            self.month = 'Май'
            list.append(self.month)
        if self.check_var6.get() == 1:
            self.month = 'Июнь'
            list.append(self.month)
        if self.check_var7.get() == 1:
            self.month = 'Июль'
            list.append(self.month)
        if self.check_var8.get() == 1:
            self.month = 'Август'
            list.append(self.month)
        if self.check_var9.get() == 1:
            self.month = 'Сентябрь'
            list.append(self.month)
        if self.check_var10.get() == 1:
            self.month = 'Октябрь'
            list.append(self.month)
        if self.check_var11.get() == 1:
            self.month = 'Ноябрь'
            list.append(self.month)
        if self.check_var12.get() == 1:
            self.month = 'Декабрь'
            list.append(self.month)
        list_2 = []
        x_coord = list
        y_coord = []
        for l in list:
            for i in items:
                if i.item == self.info_var.get() and i.year == 2020 and i.month == l:
                    x = str(i.quantity)
                    x = x.replace(',','.')
                    list_2.append(float(x))
        #for en in range(0, len(list_2)):
        for en in list_2:
            self.quantity += en
            y_coord.append(en)
        plt.title(f'Третичные продажи в упаковках по месяцам по SKU: \n{self.info_var.get()}')
        plt.grid(True)
        plt.bar(x_coord, y_coord, color='g')
        plt.show()
        tkinter.messagebox.showinfo('INFO',
                                    f'Sales in packs: {self.quantity} pcs')
    def save_weight_pen_month_to_csv(self):
        FILENAME = f"{self.info_var.get()}_month_weight_pen.csv"
        self.month = ''
        self.quantity = 0
        list = []
        if self.check_var1.get() == 1:
            self.month = 'Январь'
            list.append(self.month)
        if self.check_var2.get() == 1:
            self.month = 'Февраль'
            list.append(self.month)
        if self.check_var3.get() == 1:
            self.month = 'Март'
            list.append(self.month)
        if self.check_var4.get() == 1:
            self.month = 'Апрель'
            list.append(self.month)
        if self.check_var5.get() == 1:
            self.month = 'Май'
            list.append(self.month)
        if self.check_var6.get() == 1:
            self.month = 'Июнь'
            list.append(self.month)
        if self.check_var7.get() == 1:
            self.month = 'Июль'
            list.append(self.month)
        if self.check_var8.get() == 1:
            self.month = 'Август'
            list.append(self.month)
        if self.check_var9.get() == 1:
            self.month = 'Сентябрь'
            list.append(self.month)
        if self.check_var10.get() == 1:
            self.month = 'Октябрь'
            list.append(self.month)
        if self.check_var11.get() == 1:
            self.month = 'Ноябрь'
            list.append(self.month)
        if self.check_var12.get() == 1:
            self.month = 'Декабрь'
            list.append(self.month)
        list_2 = []
        x_coord = list
        y_coord = []
        for l in list:
            for i in items:
                if i.item == self.info_var.get() and i.year == 2020 and i.month == l:
                    x = str(i.weight_penetration)
                    x = x.replace(',','.')
                    list_2.append(float(x))

        for en in list_2:
            self.quantity += en
            y_coord.append(en)
        users = []
        for num in range(0,len(list)):
                users.append({"month": str(list[num]),
                    "weighted_penetration": str(list_2[num])})
        with open(FILENAME, "w", newline="") as file:
            columns = ["month", "weighted_penetration"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()

            # запись нескольких строк
            writer.writerows(users)

        print(list)
        print(list_2)
        tkinter.messagebox.showinfo('INFO',
                                    f'File {FILENAME} has been succesfully written!')
    def save_weight_pen_month_item_to_csv(self):
        FILENAME = f"{self.info_var.get()}_month_weight_pen.csv"
        self.month = ''
        self.quantity = 0
        list = []
        if self.check_var1.get() == 1:
            self.month = 'Январь'
            list.append(self.month)
        if self.check_var2.get() == 1:
            self.month = 'Февраль'
            list.append(self.month)
        if self.check_var3.get() == 1:
            self.month = 'Март'
            list.append(self.month)
        if self.check_var4.get() == 1:
            self.month = 'Апрель'
            list.append(self.month)
        if self.check_var5.get() == 1:
            self.month = 'Май'
            list.append(self.month)
        if self.check_var6.get() == 1:
            self.month = 'Июнь'
            list.append(self.month)
        if self.check_var7.get() == 1:
            self.month = 'Июль'
            list.append(self.month)
        if self.check_var8.get() == 1:
            self.month = 'Август'
            list.append(self.month)
        if self.check_var9.get() == 1:
            self.month = 'Сентябрь'
            list.append(self.month)
        if self.check_var10.get() == 1:
            self.month = 'Октябрь'
            list.append(self.month)
        if self.check_var11.get() == 1:
            self.month = 'Ноябрь'
            list.append(self.month)
        if self.check_var12.get() == 1:
            self.month = 'Декабрь'
            list.append(self.month)
        list_2 = []
        x_coord = list
        y_coord = []
        for l in list:
            for i in items:
                if i.item == self.info_var.get() and i.year == 2020 and i.month == l:
                    x = str(i.weight_penetration)
                    x = x.replace(',', '.')
                    list_2.append(float(x))

        for en in list_2:
            self.quantity += en
            y_coord.append(en)
        users = []
        for num in range(0, len(list)):
            users.append({"item": str(self.info_var.get()),
                "month": str(list[num]),
                          "weighted_penetration": str(list_2[num])})
        with open(FILENAME, "w", newline="") as file:
            columns = ["item", "month", "weighted_penetration"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()

            # запись нескольких строк
            writer.writerows(users)

        print(list)
        print(list_2)
        tkinter.messagebox.showinfo('INFO',
                                    f'File {FILENAME} has been succesfully written!')
    def save_weight_pen_month_item_RX_PROMO_to_csv(self):
        FILENAME = "Weight_pen_month_item_RX_PROMO.csv"
        self.month = ''
        self.quantity = 0
        list = []
        if self.check_var1.get() == 1:
            self.month = 'Январь'
            list.append(self.month)
        if self.check_var2.get() == 1:
            self.month = 'Февраль'
            list.append(self.month)
        if self.check_var3.get() == 1:
            self.month = 'Март'
            list.append(self.month)
        if self.check_var4.get() == 1:
            self.month = 'Апрель'
            list.append(self.month)
        if self.check_var5.get() == 1:
            self.month = 'Май'
            list.append(self.month)
        if self.check_var6.get() == 1:
            self.month = 'Июнь'
            list.append(self.month)
        if self.check_var7.get() == 1:
            self.month = 'Июль'
            list.append(self.month)
        if self.check_var8.get() == 1:
            self.month = 'Август'
            list.append(self.month)
        if self.check_var9.get() == 1:
            self.month = 'Сентябрь'
            list.append(self.month)
        if self.check_var10.get() == 1:
            self.month = 'Октябрь'
            list.append(self.month)
        if self.check_var11.get() == 1:
            self.month = 'Ноябрь'
            list.append(self.month)
        if self.check_var12.get() == 1:
            self.month = 'Декабрь'
            list.append(self.month)
        with sqlite3.connect("tertiary_sales_database.db") as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT ymm.Month, ymm.Месяц, items.item_quadra, items.brand, tertiary_sales.WeightPenetration from tertiary_sales join ymm on tertiary_sales.Period = ymm.Year JOIN items on tertiary_sales.Fullmedicationname = items.item_proxima where tertiary_sales.MarketOrg = 'Grindeks  (Latvia)' and items.PROMOTION = 'PROMO' and ymm.Year = '2020'")
            results = cursor.fetchall()
        item_list = []
        for i in results:
            x = {"month":str(i[0]),"month_cube":str(i[1]),"item_translit":str(i[2]),"item_brand":str(i[3]),"weighted_penetration":str(i[4])}
            x[4] = str(i[4]).replace(',','.')
            i = x
            item_list.append(i)

        final_list = []
        for i in item_list:
            if i.get("month_cube") in list:
                final_list.append(i)
        print(final_list)

        with open(FILENAME, "w", newline="", encoding='UTF-8') as file:
            columns = ["month","month_cube", "item_translit", "item_brand", "weighted_penetration", 4]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()

            # запись нескольких строк
            writer.writerows(final_list)
