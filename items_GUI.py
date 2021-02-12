import csv
import json
import sqlite3
import tkinter
from tkinter import font, messagebox, Checkbutton, Radiobutton, BOTH, END
import matplotlib.pyplot as plt



import csv
import sqlite3

from items_class import SKU_WORKOUT, CSKU, CItemsDAO
from new_try.database import CEXtract_database_tertiary
from tertiary_sales_class import Tertiary_sales







conn = sqlite3.connect("tertiary_sales_database.db")

items_ = CItemsDAO.read_tertiary(conn)





class Items_GUI(tkinter.Frame):

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
        self.chb1 = tkinter.Checkbutton(self.left_frame, text='Jan', variable=self.check_var1,
                                        font=my_font1)


        self.chb2 = tkinter.Checkbutton(self.left_frame, text='Feb', variable=self.check_var2,
                                        font=my_font1)
        self.chb3 = tkinter.Checkbutton(self.left_frame, text='Mar', variable=self.check_var3,
                                        font=my_font1)
        self.chb4 = tkinter.Checkbutton(self.left_frame, text='Apr',
                                        variable=self.check_var4, font=my_font1)
        self.chb5 = tkinter.Checkbutton(self.left_frame, text='May', variable=self.check_var5, font=my_font1)
        self.chb6 = tkinter.Checkbutton(self.left_frame, text='Jun', variable=self.check_var6,
                                        font=my_font1)
        self.chb7 = tkinter.Checkbutton(self.top_frame, text='Jul', variable=self.check_var7,
                                        font=my_font1)
        self.chb8 = tkinter.Checkbutton(self.top_frame, text='Aug', variable=self.check_var8,
                                        font=my_font1)
        self.chb9 = tkinter.Checkbutton(self.top_frame, text='Sep',
                                        variable=self.check_var9, font=my_font1)
        self.chb10 = tkinter.Checkbutton(self.top_frame, text='Oct', variable=self.check_var10, font=my_font1)
        self.chb11 = tkinter.Checkbutton(self.top_frame, text='Nov', variable=self.check_var11,
                                        font=my_font1)

        self.chb12 = tkinter.Checkbutton(self.top_frame, text='Dec', variable=self.check_var12,
                                        font=my_font1,padx=15, pady=-10)

        acts = SKU_WORKOUT()
        acts_ = acts.read_item_2020()
        self.show_button = tkinter.Button(self.button_frame, text='Weighted penetration', command=self.show_weighted_penetration)
        self.show_button.pack(side='left')
        self.ok_button = tkinter.Button(self.button_frame, text='Sales in euro', command=self.onclick_euro)
        self.ok_button.pack(side='left')
        self.quit_button = tkinter.Button(self.button_frame, text='Quit', command=self.master.destroy)
        self.quit_button.pack(side='left')
        self.button_frame.pack()
        lb = tkinter.Listbox(self, width='70', height='15')
        self.ok_button_quantity = tkinter.Button(self.button_frame, text='Sales in packs', command=self.onclick_quantity)
        self.ok_button_quantity.pack(side='left')

        i_list = []
        for i in range(0,len(acts_)):
            for entry in acts_[i]:
                year = entry.get('year')
                sales_method = entry.get('sales_method')
                promotion = entry.get('promotion')
                purpose = entry.get('purpose')
                item_proxima = entry.get('item_proxima')
                item_quadra = entry.get('item_quadra')
                item_sales_report = entry.get('item_sales_report')
                item_kpi_report = entry.get('item_kpi_report')
                brand = entry.get('brand')
                month = entry.get('month')
                cip_euro = entry.get('cip_euro')
                z = CSKU(year,sales_method,promotion,purpose,item_proxima,item_quadra,item_sales_report,item_kpi_report,brand,month,cip_euro)
                i_list.append(str(z.item_kpi_report))
        i_list_1 = []

        for i in i_list:
            if i not in i_list_1:
                i_list_1.append(i)
        for item in i_list_1:
            lb.insert(END, item)

        lb.bind("<<ListboxSelect>>", self.onSelect)

        lb.pack(pady=15)

        self.var = tkinter.StringVar()
        self.label = tkinter.Label(self, text=0, textvariable=self.var)
        self.label.pack()
        self.info_var = tkinter.StringVar()
        self.info_label = tkinter.Label(self, text=0, textvariable=self.info_var)
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

    def show_weighted_penetration(self):
        self.month = ''
        list = []
        if self.check_var1.get() == 1:
            self.month = 'Jan'
            list.append(self.month)
        if self.check_var2.get() == 1:
            self.month = 'Feb'
            list.append(self.month)
        if self.check_var3.get() == 1:
            self.month = 'Mar'
            list.append(self.month)
        if self.check_var4.get() == 1:
            self.month = 'Apr'
            list.append(self.month)
        if self.check_var5.get() == 1:
            self.month = 'May'
            list.append(self.month)
        if self.check_var6.get() == 1:
            self.month = 'Jun'
            list.append(self.month)
        if self.check_var7.get() == 1:
            self.month = 'Jul'
            list.append(self.month)
        if self.check_var8.get() == 1:
            self.month = 'Aug'
            list.append(self.month)
        if self.check_var9.get() == 1:
            self.month = 'Sep'
            list.append(self.month)
        if self.check_var10.get() == 1:
            self.month = 'Oct'
            list.append(self.month)
        if self.check_var11.get() == 1:
            self.month = 'Nov'
            list.append(self.month)
        if self.check_var12.get() == 1:
            self.month = 'Dec'
            list.append(self.month)
        x_coord = list
        ss = CEXtract_database_tertiary()
        pos = ss.read_item_2020()
        basic_list = []
        for i in pos:
            z = Tertiary_sales(i)
            if z.month in list and z.item == self.info_var.get():
                basic_list.append(float(z.weight_penetration))
                print(basic_list)
        y_coord = basic_list

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
            self.month = 'Jan'
            list.append(self.month)
        if self.check_var2.get() == 1:
            self.month = 'Feb'
            list.append(self.month)
        if self.check_var3.get() == 1:
            self.month = 'Mar'
            list.append(self.month)
        if self.check_var4.get() == 1:
            self.month = 'Apr'
            list.append(self.month)
        if self.check_var5.get() == 1:
            self.month = 'May'
            list.append(self.month)
        if self.check_var6.get() == 1:
            self.month = 'Jun'
            list.append(self.month)
        if self.check_var7.get() == 1:
            self.month = 'Jul'
            list.append(self.month)
        if self.check_var8.get() == 1:
            self.month = 'Aug'
            list.append(self.month)
        if self.check_var9.get() == 1:
            self.month = 'Sep'
            list.append(self.month)
        if self.check_var10.get() == 1:
            self.month = 'Oct'
            list.append(self.month)
        if self.check_var11.get() == 1:
            self.month = 'Nov'
            list.append(self.month)
        if self.check_var12.get() == 1:
            self.month = 'Dec'
            list.append(self.month)

        x_coord = list
        ss = CEXtract_database_tertiary()
        pos = ss.read_item_2020()
        basic_list = []
        for i in pos:
            z = Tertiary_sales(i)
            if z.month in list and z.item == self.info_var.get():
                basic_list.append(float(z.volume_euro))
                print(basic_list)
        for i in basic_list:
            self.amount_euro += float(i)
        y_coord = basic_list
        print(self.amount_euro)
        tkinter.messagebox.showinfo('INFO', f'Sales in euro: {self.amount_euro} euro')
        plt.title(f'Третичные продажи в евро по месяцам по SKU: \n{self.info_var.get()}')
        plt.grid(True)
        plt.plot(x_coord,y_coord,marker='s')
        plt.show()


    def save_quant_month_to_json(self):
        FILENAME = f"{self.info_var.get()}_month_quantity.json"
        self.month = ''
        self.quantity = 0
        list = []
        if self.check_var1.get() == 1:
            self.month = 'Jan'
            list.append(self.month)
        if self.check_var2.get() == 1:
            self.month = 'Feb'
            list.append(self.month)
        if self.check_var3.get() == 1:
            self.month = 'Mar'
            list.append(self.month)
        if self.check_var4.get() == 1:
            self.month = 'Apr'
            list.append(self.month)
        if self.check_var5.get() == 1:
            self.month = 'May'
            list.append(self.month)
        if self.check_var6.get() == 1:
            self.month = 'Jun'
            list.append(self.month)
        if self.check_var7.get() == 1:
            self.month = 'Jul'
            list.append(self.month)
        if self.check_var8.get() == 1:
            self.month = 'Aug'
            list.append(self.month)
        if self.check_var9.get() == 1:
            self.month = 'Sep'
            list.append(self.month)
        if self.check_var10.get() == 1:
            self.month = 'Oct'
            list.append(self.month)
        if self.check_var11.get() == 1:
            self.month = 'Nov'
            list.append(self.month)
        if self.check_var12.get() == 1:
            self.month = 'Dec'
            list.append(self.month)


        ss = CEXtract_database_tertiary()
        pos = ss.read_item_2020()
        basic_list = []
        for i in pos:
            z = Tertiary_sales(i)
            if z.month in list and z.item == self.info_var.get():
                basic_list.append(float(z.quantity))
                print(basic_list)
        for i in basic_list:
            self.quantity += float(i)

        print(self.quantity)
        users = []
        for num in range(0,len(list)):
                users.append({"month": str(list[num]),
                    "quantity_packs": str(basic_list[num])})

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
            self.month = 'Jan'
            list.append(self.month)
        if self.check_var2.get() == 1:
            self.month = 'Feb'
            list.append(self.month)
        if self.check_var3.get() == 1:
            self.month = 'Mar'
            list.append(self.month)
        if self.check_var4.get() == 1:
            self.month = 'Apr'
            list.append(self.month)
        if self.check_var5.get() == 1:
            self.month = 'May'
            list.append(self.month)
        if self.check_var6.get() == 1:
            self.month = 'Jun'
            list.append(self.month)
        if self.check_var7.get() == 1:
            self.month = 'Jul'
            list.append(self.month)
        if self.check_var8.get() == 1:
            self.month = 'Aug'
            list.append(self.month)
        if self.check_var9.get() == 1:
            self.month = 'Sep'
            list.append(self.month)
        if self.check_var10.get() == 1:
            self.month = 'Oct'
            list.append(self.month)
        if self.check_var11.get() == 1:
            self.month = 'Nov'
            list.append(self.month)
        if self.check_var12.get() == 1:
            self.month = 'Dec'
            list.append(self.month)
        list_2 = []
        ss = CEXtract_database_tertiary()
        pos = ss.read_item_2020()
        basic_list = []
        for i in pos:
            z = Tertiary_sales(i)
            if z.month in list and z.item == self.info_var.get():
                basic_list.append(float(z.weight_penetration))
                print(basic_list)
        users = []
        for num in range(0,len(list)):
                users.append({"month": str(list[num]),
                    "weighted_penetration": str(basic_list[num])})
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
            self.month = 'Jan'
            list.append(self.month)
        if self.check_var2.get() == 1:
            self.month = 'Feb'
            list.append(self.month)
        if self.check_var3.get() == 1:
            self.month = 'Mar'
            list.append(self.month)
        if self.check_var4.get() == 1:
            self.month = 'Apr'
            list.append(self.month)
        if self.check_var5.get() == 1:
            self.month = 'May'
            list.append(self.month)
        if self.check_var6.get() == 1:
            self.month = 'Jun'
            list.append(self.month)
        if self.check_var7.get() == 1:
            self.month = 'Jul'
            list.append(self.month)
        if self.check_var8.get() == 1:
            self.month = 'Aug'
            list.append(self.month)
        if self.check_var9.get() == 1:
            self.month = 'Sep'
            list.append(self.month)
        if self.check_var10.get() == 1:
            self.month = 'Oct'
            list.append(self.month)
        if self.check_var11.get() == 1:
            self.month = 'Nov'
            list.append(self.month)
        if self.check_var12.get() == 1:
            self.month = 'Dec'
            list.append(self.month)
        list_2 = []
        x_coord = list
        y_coord = []

        ss = CEXtract_database_tertiary()
        pos = ss.read_item_2020()
        basic_list = []
        for i in pos:
            z = Tertiary_sales(i)
            if z.month in list and z.item == self.info_var.get():
                basic_list.append(float(z.quantity))
                print(basic_list)
        for i in basic_list:
            self.quantity += float(i)
        y_coord = basic_list
        print(self.quantity)
        tkinter.messagebox.showinfo('INFO',
                                    f'Sales in packs: {self.quantity} pcs')
        plt.title(f'Третичные продажи в упаковках по месяцам по SKU: \n{self.info_var.get()}')
        plt.grid(True)
        plt.plot(x_coord, y_coord,  color='g')
        plt.show()

    def save_weight_pen_month_to_csv(self):
        FILENAME = f"{self.info_var.get()}_month_weight_pen.csv"
        self.month = ''
        self.quantity = 0
        list = []
        if self.check_var1.get() == 1:
            self.month = 'Jan'
            list.append(self.month)
        if self.check_var2.get() == 1:
            self.month = 'Feb'
            list.append(self.month)
        if self.check_var3.get() == 1:
            self.month = 'Mar'
            list.append(self.month)
        if self.check_var4.get() == 1:
            self.month = 'Apr'
            list.append(self.month)
        if self.check_var5.get() == 1:
            self.month = 'May'
            list.append(self.month)
        if self.check_var6.get() == 1:
            self.month = 'Jun'
            list.append(self.month)
        if self.check_var7.get() == 1:
            self.month = 'Jul'
            list.append(self.month)
        if self.check_var8.get() == 1:
            self.month = 'Aug'
            list.append(self.month)
        if self.check_var9.get() == 1:
            self.month = 'Sep'
            list.append(self.month)
        if self.check_var10.get() == 1:
            self.month = 'Oct'
            list.append(self.month)
        if self.check_var11.get() == 1:
            self.month = 'Nov'
            list.append(self.month)
        if self.check_var12.get() == 1:
            self.month = 'Dec'
            list.append(self.month)
        x_coord = list
        ss = CEXtract_database_tertiary()
        pos = ss.read_item_2020()
        basic_list = []
        for i in pos:
            z = Tertiary_sales(i)
            if z.month in list and z.item == self.info_var.get():
                basic_list.append(float(z.weight_penetration))
                print(basic_list)
        y_coord = basic_list
        users = []
        for num in range(0,len(list)):
                users.append({"month": str(list[num]),
                    "weighted_penetration": str(basic_list[num])})
        with open(FILENAME, "w", newline="") as file:
            columns = ["month", "weighted_penetration"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()

            # запись нескольких строк
            writer.writerows(users)

        tkinter.messagebox.showinfo('INFO',
                                    f'File {FILENAME} has been succesfully written!')
    def save_weight_pen_month_item_to_csv(self):
        FILENAME = f"{self.info_var.get()}_month_weight_pen.csv"
        self.month = ''
        self.quantity = 0
        list = []
        if self.check_var1.get() == 1:
            self.month = 'Jan'
            list.append(self.month)
        if self.check_var2.get() == 1:
            self.month = 'Feb'
            list.append(self.month)
        if self.check_var3.get() == 1:
            self.month = 'Mar'
            list.append(self.month)
        if self.check_var4.get() == 1:
            self.month = 'Apr'
            list.append(self.month)
        if self.check_var5.get() == 1:
            self.month = 'May'
            list.append(self.month)
        if self.check_var6.get() == 1:
            self.month = 'Jun'
            list.append(self.month)
        if self.check_var7.get() == 1:
            self.month = 'Jul'
            list.append(self.month)
        if self.check_var8.get() == 1:
            self.month = 'Aug'
            list.append(self.month)
        if self.check_var9.get() == 1:
            self.month = 'Sep'
            list.append(self.month)
        if self.check_var10.get() == 1:
            self.month = 'Oct'
            list.append(self.month)
        if self.check_var11.get() == 1:
            self.month = 'Nov'
            list.append(self.month)
        if self.check_var12.get() == 1:
            self.month = 'Dec'
            list.append(self.month)

        ss = CEXtract_database_tertiary()
        pos = ss.read_item_2020()
        basic_list = []
        for i in pos:
            z = Tertiary_sales(i)
            if z.month in list and z.item == self.info_var.get():
                basic_list.append(float(z.weight_penetration))
                print(basic_list)

        users = []
        for num in range(0, len(list)):
            users.append({"item": str(self.info_var.get()),
                "month": str(list[num]),
                          "weighted_penetration": str(basic_list[num])})
        with open(FILENAME, "w", newline="") as file:
            columns = ["item", "month", "weighted_penetration"]
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()

            # запись нескольких строк
            writer.writerows(users)
        tkinter.messagebox.showinfo('INFO',
                                    f'File {FILENAME} has been succesfully written!')
    def save_all_data_2020_PROMO_items_to_csv(self):
        FILENAME = "Weight_pen_month_item_all_PROMO_items.csv"
        self.month = ''
        self.quantity = 0
        list = []
        if self.check_var1.get() == 1:
            self.month = 'Jan'
            list.append(self.month)
        if self.check_var2.get() == 1:
            self.month = 'Feb'
            list.append(self.month)
        if self.check_var3.get() == 1:
            self.month = 'Mar'
            list.append(self.month)
        if self.check_var4.get() == 1:
            self.month = 'Apr'
            list.append(self.month)
        if self.check_var5.get() == 1:
            self.month = 'May'
            list.append(self.month)
        if self.check_var6.get() == 1:
            self.month = 'Jun'
            list.append(self.month)
        if self.check_var7.get() == 1:
            self.month = 'Jul'
            list.append(self.month)
        if self.check_var8.get() == 1:
            self.month = 'Aug'
            list.append(self.month)
        if self.check_var9.get() == 1:
            self.month = 'Sep'
            list.append(self.month)
        if self.check_var10.get() == 1:
            self.month = 'Oct'
            list.append(self.month)
        if self.check_var11.get() == 1:
            self.month = 'Nov'
            list.append(self.month)
        if self.check_var12.get() == 1:
            self.month = 'Dec'
            list.append(self.month)

        ss = CEXtract_database_tertiary()
        pos = ss.read_item_2020()

        ss.save_2020_items_to_csv(FILENAME,list)
        tkinter.messagebox.showinfo('INFO',
                                    f'File {FILENAME} has been succesfully written!')
    def save_all_data_2020_PROMO_items_to_csv_commas(self):
        FILENAME = "Weight_pen_month_item_all_PROMO_items_with_commas.csv"
        self.month = ''
        self.quantity = 0
        list = []
        if self.check_var1.get() == 1:
            self.month = 'Jan'
            list.append(self.month)
        if self.check_var2.get() == 1:
            self.month = 'Feb'
            list.append(self.month)
        if self.check_var3.get() == 1:
            self.month = 'Mar'
            list.append(self.month)
        if self.check_var4.get() == 1:
            self.month = 'Apr'
            list.append(self.month)
        if self.check_var5.get() == 1:
            self.month = 'May'
            list.append(self.month)
        if self.check_var6.get() == 1:
            self.month = 'Jun'
            list.append(self.month)
        if self.check_var7.get() == 1:
            self.month = 'Jul'
            list.append(self.month)
        if self.check_var8.get() == 1:
            self.month = 'Aug'
            list.append(self.month)
        if self.check_var9.get() == 1:
            self.month = 'Sep'
            list.append(self.month)
        if self.check_var10.get() == 1:
            self.month = 'Oct'
            list.append(self.month)
        if self.check_var11.get() == 1:
            self.month = 'Nov'
            list.append(self.month)
        if self.check_var12.get() == 1:
            self.month = 'Dec'
            list.append(self.month)

        ss = CEXtract_database_tertiary()
        pos = ss.read_item_2020_w_commas()

        ss.save_items_2020_to_csv_with_commas(FILENAME,list)
        tkinter.messagebox.showinfo('INFO',
                                    f'File {FILENAME} has been succesfully written!')
    def save_otc_2020_PROMO_items_to_csv(self):
        FILENAME = "Otc_PROMO_items_all_data2020.csv"
        self.month = ''
        self.quantity = 0
        list = []
        if self.check_var1.get() == 1:
            self.month = 'Jan'
            list.append(self.month)
        if self.check_var2.get() == 1:
            self.month = 'Feb'
            list.append(self.month)
        if self.check_var3.get() == 1:
            self.month = 'Mar'
            list.append(self.month)
        if self.check_var4.get() == 1:
            self.month = 'Apr'
            list.append(self.month)
        if self.check_var5.get() == 1:
            self.month = 'May'
            list.append(self.month)
        if self.check_var6.get() == 1:
            self.month = 'Jun'
            list.append(self.month)
        if self.check_var7.get() == 1:
            self.month = 'Jul'
            list.append(self.month)
        if self.check_var8.get() == 1:
            self.month = 'Aug'
            list.append(self.month)
        if self.check_var9.get() == 1:
            self.month = 'Sep'
            list.append(self.month)
        if self.check_var10.get() == 1:
            self.month = 'Oct'
            list.append(self.month)
        if self.check_var11.get() == 1:
            self.month = 'Nov'
            list.append(self.month)
        if self.check_var12.get() == 1:
            self.month = 'Dec'
            list.append(self.month)

        ss = CEXtract_database_tertiary()
        pos = ss.read_item_2020_OTC()

        ss.save_items_otc_to_csv_2020(FILENAME,list)
        tkinter.messagebox.showinfo('INFO',
                                    f'File {FILENAME} has been succesfully written!')