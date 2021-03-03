import json
import os
import tkinter
from datetime import date
from time import strftime
import xlwings as xw
from tkinter import font, messagebox, BOTH, END
import matplotlib.pyplot as plt
import csv
import sqlite3

from dateutil.utils import today
from pandas.tests.io.excel.test_openpyxl import openpyxl
from win32timezone import now

from items_class import SKU_WORKOUT, CSKU, CItemsDAO
from sale_out.database import CEXtract_database_tertiary, Tertiary_download_structure, CBase_2021_quadra_workout, \
    Kam_plans

conn = sqlite3.connect("C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\local_main_base.db")
items_ = CItemsDAO.read_tertiary(conn)

class Items_GUI(tkinter.Frame):

    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.master.title( "Work helper")
        self.pack(fill=BOTH)
        #self.pack(fill=BOTH, expand=1)
        self.upper_frame = tkinter.Frame(self.master)
        self.top_frame = tkinter.Frame(self.master)
        self.button_frame = tkinter.Frame(self.master)
        self.left_frame = tkinter.Frame(self.master)
        self.second_upper_frame = tkinter.Frame(self.master)
        self.radio_var = tkinter.IntVar()
        self.radio_var.set(2021)
        self.check_var1 = tkinter.IntVar()
        self.check_var1.set(0)
        self.check_var2 = tkinter.IntVar()
        self.check_var2.set(0)
        self.check_var3 = tkinter.IntVar()
        self.check_var3.set(1)
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

        self.previous_rows_count = tkinter.IntVar()
        self.previous_rows_count_label = tkinter.Label(self, text='Row count base', textvariable=self.previous_rows_count)
        self.previous_rows_count_label.pack()

        my_font = tkinter.font.Font(family='Arial', size=12, weight='bold')
        my_font1 = tkinter.font.Font(family='Arial', size=11,weight='bold')
        self.specific_data = tkinter.StringVar()
        self.specific_data_label = tkinter.Label(self, text='Specific data', textvariable=self.specific_data, font=my_font)
        self.specific_data_label.pack()


        self.chb1 = tkinter.Checkbutton(self.top_frame, text='Jan', variable=self.check_var1,
                                        font=my_font1)

        self.rb1 = tkinter.Radiobutton(self.top_frame,text='2019',variable=self.radio_var,value=2019,font=my_font1)
        self.rb2 = tkinter.Radiobutton(self.top_frame, text='2020', variable=self.radio_var, value=2020,font=my_font1)
        self.rb3 = tkinter.Radiobutton(self.top_frame, text='2021', variable=self.radio_var, value=2021,font=my_font1)

        self.chb2 = tkinter.Checkbutton(self.top_frame, text='Feb', variable=self.check_var2,
                                        font=my_font1)
        self.chb3 = tkinter.Checkbutton(self.top_frame, text='Mar', variable=self.check_var3,
                                        font=my_font1)
        self.chb4 = tkinter.Checkbutton(self.top_frame, text='Apr',
                                        variable=self.check_var4, font=my_font1)
        self.chb5 = tkinter.Checkbutton(self.top_frame, text='May', variable=self.check_var5, font=my_font1)
        self.chb6 = tkinter.Checkbutton(self.top_frame, text='Jun', variable=self.check_var6,
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
        self.show_button_pen = tkinter.Button(self.button_frame, text='Penetration', command=self.show_penetration)
        self.show_button_pen.pack(side='left')
        self.show_button_sro = tkinter.Button(self.button_frame, text='SRO', command=self.show_sro)
        self.show_button_sro.pack(side='left')
        self.show_button_weight_sro = tkinter.Button(self.button_frame, text='Weighted SRO', command=self.show_weighted_sro)
        self.show_button_weight_sro.pack(side='left')
        self.ok_button = tkinter.Button(self.button_frame, text='Sale-out in euro', command=self.onclick_euro)
        self.ok_button.pack(side='left')

        self.button_frame.pack()
        lb = tkinter.Listbox(self, width='25', height='15')
        self.ok_button_quantity = tkinter.Button(self.button_frame, text='Sale-out in packs', command=self.onclick_quantity)
        self.ok_button_quantity.pack(side='left')
        self.upper_frame.pack()
        self.secondary_euro = tkinter.Button(self.second_upper_frame, text='Update 629 xlxs', command=self.secondary_sales_euro_upload)
        self.secondary_euro.pack(side='left')
        self.update_629_base_2020_xlxs_button_sec_euro = tkinter.Button(self.second_upper_frame, text='Update 629 report xlxs 2020', command=self.Update_2020_629_xlxs)
        self.update_629_base_2020_xlxs_button_sec_euro.pack(side='left')

        self.show_button_sec_euro = tkinter.Button(self.upper_frame, text='Total secondary sales euro plan KAM', command=self.kam_plans_for_chart_from_sqlite3)
        self.show_button_sec_euro.pack(side='left')
        self.rewrite_629_base_button_sec_euro = tkinter.Button(self.second_upper_frame, text='Rewrite the base 2021', command=self.rewrite_2021_629_base)
        self.rewrite_629_base_button_sec_euro.pack(side='left')


        self.rewrite_629_base_2020_button_sec_euro = tkinter.Button(self.second_upper_frame, text='Rewrite the base 2020', command=self.rewrite_2020_629_base)
        self.rewrite_629_base_2020_button_sec_euro.pack(side='left')
        self.rewrite_629_base_button_sec_euro = tkinter.Button(self.second_upper_frame, text='Save transformations', command=self.save_transformations_for_riga_sales_report)
        self.rewrite_629_base_button_sec_euro.pack(side='left')
        self.refresh_button = tkinter.Button(self.second_upper_frame, text='Read data from Quadra server', command=self.secondary_2021_total_pack_euro,font=my_font1)
        self.refresh_button.pack(side='left')
        self.quit_button = tkinter.Button(self.second_upper_frame, text='Quit', command=self.master.destroy,font=my_font1)
        self.quit_button.pack(side='left')
        self.show_secondary_sales_list = tkinter.Button(self.upper_frame, text='Total secondary sales euro', command=self.show_secondary_sales_by_month)
        self.show_secondary_sales_list.pack(side='left')
        self.show_secondary_sales_by_item = tkinter.Button(self.upper_frame, text='Secondary sales by item euro', command=self.show_secondary_sales_by_month_by_item)
        self.show_secondary_sales_by_item.pack(side='left')



        i_list = []
        self.dict_item = {}
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
                self.dict_item.update({z.item_kpi_report:z.item_quadra})
        i_list_1 = []

        for i in i_list:
            if i not in i_list_1:
                i_list_1.append(i)
        for item in i_list_1:
            lb.insert(END, item)

        lb.bind("<<ListboxSelect>>", self.onSelect)

        lb.pack(padx=10,pady=5,fill=tkinter.BOTH,expand=False)

        self.var = tkinter.StringVar()
        self.label = tkinter.Label(self, text=0, textvariable=self.var)
        self.label.pack()
        self.tot_sec_2021_packs = tkinter.StringVar()
        self.tot_sec_2021_euro = tkinter.StringVar()
        self.mtd_sec_2021_packs = tkinter.StringVar()
        self.mtd_sec_2021_euro = tkinter.StringVar()
        self.secondary_2021_total_pack_euro() #launch secondary workout
        self.label_tot_sec_2021_packs = tkinter.Label(self, text='YTD 2021 secondary sales in packs: ', textvariable=self.tot_sec_2021_packs,font=my_font1)
        self.label_tot_sec_2021_packs.pack()
        self.label_tot_sec_2021_euro = tkinter.Label(self, text='YTD 2021 secondary sales in euro: ', textvariable=self.tot_sec_2021_euro,font=my_font1)
        self.label_tot_sec_2021_euro.pack()
        self.mtd_sec_2021_packs_label = tkinter.Label(self, text='MTD 2021 secondary sales in packs: ', textvariable=self.mtd_sec_2021_packs,font=my_font1)
        self.mtd_sec_2021_packs_label.pack()
        self.mtd_sec_2021_euro_label = tkinter.Label(self, text='MTD 2021 secondary sales in euro: ', textvariable=self.mtd_sec_2021_euro,font=my_font1)
        self.mtd_sec_2021_euro_label.pack()
        list, list_months_quadra, year = self.radiobutton_months()
        selected_period_euro = 0
        selected_period_packs = 0
        if year  == 2021:
            x = CBase_2021_quadra_workout()
            print(list_months_quadra[0])
            mtd_packs, mtd_euro = x.get_secondary_2021_by_month_from_sqlite3(str(list_months_quadra[0]))
            selected_period_euro = mtd_euro
            selected_period_packs = mtd_packs
        self.mtd_sec_2021_euro.set('MTD 2021 secondary sales in euro:   '+'{0:,}'.format(selected_period_euro.__round__(2)).replace(",", " ")+' euro')
        self.mtd_sec_2021_packs.set('MTD 2021 secondary sales in packs:   '+'{0:,}'.format(selected_period_packs.__round__(0)).replace(",", " ") + ' packs')

        self.info_var = tkinter.StringVar()
        self.info_label = tkinter.Label(self, text=0, textvariable=self.info_var,font=my_font1)
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
        self.rb1.pack(side='top')
        self.rb2.pack(side='top')
        self.rb3.pack(side='top')
        self.top_frame.pack(side='bottom')
        self.left_frame.pack(side='left')
        self.second_upper_frame.pack(side='left')
    def calculate_actual_rows_in_base_2021(self):
        x = CBase_2021_quadra_workout()
        status = x.calculate_rows_in_2021_base_from_quadra()
        st_1 = 'Already updated database'
        st_2 = 'Xlsx and database have just been updated'
        if status == st_1:
            self.previous_rows_count.set(status)
        else:
            os.startfile(
                'C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\Two In One - Wave Your Hands In The Air.mp3')
            self.secondary_sales_euro_upload()
            self.rewrite_2021_629_base()
            self.previous_rows_count.set(st_2)

            tkinter.messagebox.showinfo('INFO',
                                        f'Basic excel file for 2021 and corresponding database\nhas been successfully updated!')
    def secondary_sales_euro_upload(self):
        x = CBase_2021_quadra_workout()
        x.save_base_629_2021_to_xlsx()
        self.secondary_button_name = now()

        return self.secondary_button_name
    def secondary_2021_total_pack_euro(self):
        x = CBase_2021_quadra_workout()
        x.rewrite_629_2021_in_database()
        total_packs, total_euro = x.get_629_2021_from_sqlite3()
        self.tot_sec_2021_euro.set('YTD 2021 secondary sales in euro:   '+'{0:,}'.format(total_euro.__round__(2)).replace(",", " ")+ ' euro')
        self.tot_sec_2021_packs.set('YTD 2021 secondary sales in packs:   '+'{0:,}'.format(total_packs.__round__(0)).replace(",", " ") + ' packs')
        self.calculate_actual_rows_in_base_2021()
    def radiobutton_months(self):
        self.month = ''
        self.amount_euro = 0
        year = self.radio_var.get()
        list = []
        list_months_quadra = []
        if self.check_var1.get() == 1:
            self.month = 'Jan'
            list.append(self.month)
            list_months_quadra.append('Январь')
        if self.check_var2.get() == 1:
            self.month = 'Feb'
            list.append(self.month)
            list_months_quadra.append('Февраль')
        if self.check_var3.get() == 1:
            self.month = 'Mar'
            list.append(self.month)
            list_months_quadra.append('Март')
        if self.check_var4.get() == 1:
            self.month = 'Apr'
            list.append(self.month)
            list_months_quadra.append('Апрель')
        if self.check_var5.get() == 1:
            self.month = 'May'
            list.append(self.month)
            list_months_quadra.append('Май')
        if self.check_var6.get() == 1:
            self.month = 'Jun'
            list.append(self.month)
            list_months_quadra.append('Июнь')
        if self.check_var7.get() == 1:
            self.month = 'Jul'
            list.append(self.month)
            list_months_quadra.append('Июль')
        if self.check_var8.get() == 1:
            self.month = 'Aug'
            list.append(self.month)
            list_months_quadra.append('Август')
        if self.check_var9.get() == 1:
            self.month = 'Sep'
            list.append(self.month)
            list_months_quadra.append('Сентябрь')
        if self.check_var10.get() == 1:
            self.month = 'Oct'
            list.append(self.month)
            list_months_quadra.append('Октябрь')
        if self.check_var11.get() == 1:
            self.month = 'Nov'
            list.append(self.month)
            list_months_quadra.append('Ноябрь')
        if self.check_var12.get() == 1:
            self.month = 'Dec'
            list.append(self.month)
            list_months_quadra.append('Декабрь')
        return list, list_months_quadra, year
    def secondary_sales_euro_2021_for_chart(self):
        list, list_months_quadra, year = self.radiobutton_months()

        x_coord = list_months_quadra
        x = CBase_2021_quadra_workout()
        if year  == 2021:
            base_2021_classifyed = x.get_secondary_2021_by_month()
            basic_list = []
            for month in list_months_quadra:
                selected_period_euro = 0
                for string in base_2021_classifyed:
                    for item in string:
                        if month == item.month:
                            selected_period_euro += float(item.sales_euro)
                basic_list.append(selected_period_euro)
            y_coord = basic_list
            print(y_coord)
            self.specific_data.set(f'Specific data: {y_coord}'.replace('[', '').replace(']', ''))
            plt.title(f'Total secondary sales 2021 by month: \n{self.info_var.get()}')
            plt.grid(True)
            plt.plot(x_coord, y_coord, marker='s')
            plt.show()

        elif year  == 2020:
            base_2020_classifyed = x.get_secondary_2020_by_month()
            basic_list = []
            for month in list_months_quadra:
                selected_period_euro = 0
                for string in base_2020_classifyed:
                    for item in string:
                        if month == item.month:
                            selected_period_euro += round(float(item.sales_euro),0)
                basic_list.append(selected_period_euro)
            y_coord = basic_list
            print(y_coord)
            self.specific_data.set(f'Specific data: {y_coord}'.replace('[','').replace(']',''))
            plt.title(f'Total secondary sales 2020 by month: \n{self.info_var.get()}')
            plt.grid(True)
            plt.plot(x_coord, y_coord, marker='s')
            plt.show()
    def kam_plans_for_chart_from_sqlite3(self):
        list, list_months_quadra, year = self.radiobutton_months()

        x_coord = []
        y_coord = []
        x = Kam_plans()
        basic_list = []
        plan_list = x.read_kam_plan()


        for string in plan_list:
            if string.code_sf not in x_coord:
                x_coord.append(string.code_sf)

        for sf in x_coord:
            plan_euro_sum = 0
            for string in plan_list:
                if string.code_sf == sf:
                    for month in list_months_quadra:
                        if month == string.month_local:
                            string.plan_euro = str(string.plan_euro).replace(' ','')
                            string.plan_euro = round(float(string.plan_euro),2)
                            plan_euro_sum += string.plan_euro
            y_coord.append(plan_euro_sum)
        print(x_coord)
        print(y_coord)

        self.specific_data.set(f'Specific data: {y_coord}'.replace('[', '').replace(']', ''))
        plt.title(f'Total secondary sales plan euro for KAM 2021 by month')
        plt.grid(True)
        plt.bar(x_coord, y_coord)
        plt.show()
    def show_weighted_penetration(self):
        self.month = ''
        year = self.radio_var.get()
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
        pos = ss.read_item(year)
        basic_list = []
        for i in pos:
            z = Tertiary_download_structure(i.item_kpi,i.year,i.month,i.brand,i.weight_pen,i.sro,i.pen,i.quantity,i.amount_euro,i.weight_sro)
            if z.month in list and z.item_kpi == self.info_var.get():
                basic_list.append(float(z.weight_pen))
                print(basic_list)
        y_coord = basic_list
        self.specific_data.set(f'Specific data: {y_coord}'.replace('[','').replace(']',''))
        plt.title(f'Взвешенная пенетрация по месяцам по выбранному SKU: \n{self.info_var.get()}')
        plt.grid(True)
        plt.plot(x_coord, y_coord,marker='s')
        plt.show()
    def onSelect(self, val):
        sender = val.widget
        idx = sender.curselection()
        value = sender.get(idx)
        x = value
        self.info_var.set(value)
        self.var.set(f"Проработать словарь и выводить сюда информацию о SKU\nтакже оптимизировать обработку материнской базы, переподключив выборку\nТакже нужно создать пакет для загрузки в BI")
    def show_penetration(self):
        self.month = ''
        self.amount_euro = 0
        year = self.radio_var.get()
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
        pos = ss.read_item(year)
        basic_list = []
        for i in pos:
            z = Tertiary_download_structure(i.item_kpi, i.year, i.month, i.brand, i.weight_pen, i.sro, i.pen,
                                            i.quantity, i.amount_euro, i.weight_sro)
            if z.month in list and z.item_kpi == self.info_var.get():
                basic_list.append(float(z.pen))
                print(basic_list)

        y_coord = basic_list
        print(self.amount_euro)
        self.specific_data.set(f'Specific data: {y_coord}'.replace('[','').replace(']',''))
        plt.title(f'Пенетрация по месяцам по SKU: \n{self.info_var.get()}')
        plt.grid(True)
        plt.plot(x_coord,y_coord,marker='s')
        plt.show()
    def onclick_euro(self):
        self.month = ''
        self.amount_euro = 0
        year = self.radio_var.get()
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
        pos = ss.read_item(year)
        basic_list = []
        for i in pos:
            z = Tertiary_download_structure(i.item_kpi, i.year, i.month, i.brand, i.weight_pen, i.sro, i.pen,
                                            i.quantity, i.amount_euro, i.weight_sro)
            if z.month in list and z.item_kpi == self.info_var.get():
                basic_list.append(float(z.amount_euro))
                print(basic_list)
        for i in basic_list:
            self.amount_euro += float(i)
        y_coord = basic_list
        print(self.amount_euro)
        self.specific_data.set(f'Specific data: {y_coord}'.replace('[','').replace(']',''))
        tkinter.messagebox.showinfo('INFO', f'Sales in euro: {self.amount_euro} euro')
        plt.title(f'Третичные продажи в евро по месяцам по SKU: \n{self.info_var.get()}')
        plt.grid(True)
        plt.plot(x_coord,y_coord,marker='s')
        plt.show()
    def show_sro(self):
        self.month = ''
        self.amount_euro = 0
        list = []
        year = self.radio_var.get()
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
        pos = ss.read_item(year)
        basic_list = []
        for i in pos:
            z = Tertiary_download_structure(i.item_kpi, i.year, i.month, i.brand, i.weight_pen, i.sro, i.pen,
                                            i.quantity, i.amount_euro, i.weight_sro)
            if z.month in list and z.item_kpi == self.info_var.get():
                basic_list.append(float(z.sro))
                print(basic_list)

        y_coord = basic_list
        self.specific_data.set(f'Specific data: {y_coord}'.replace('[','').replace(']',''))

        plt.title(f'SRO: \n{self.info_var.get()}')
        plt.grid(True)
        plt.plot(x_coord,y_coord,marker='s')
        plt.show()
    def show_weighted_sro(self):
        self.month = ''
        self.amount_euro = 0
        list = []
        year = self.radio_var.get()
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
        pos = ss.read_item(year)
        basic_list = []
        for i in pos:
            z = Tertiary_download_structure(i.item_kpi, i.year, i.month, i.brand, i.weight_pen, i.sro, i.pen,
                                            i.quantity, i.amount_euro, i.weight_sro)
            if z.month in list and z.item_kpi == self.info_var.get():
                basic_list.append(float(z.weight_sro))
                print(basic_list)

        y_coord = basic_list
        self.specific_data.set(f'Specific data: {y_coord}'.replace('[','').replace(']',''))

        plt.title(f'Weighted SRO: \n{self.info_var.get()}')
        plt.grid(True)
        plt.plot(x_coord,y_coord,marker='s')
        plt.savefig(f"weighted SRO_{year}.pdf", bbox_inches='tight')
        plt.show()
    def save_quant_month_to_json(self):
        year = self.radio_var.get()
        FILENAME = f"{self.info_var.get()}_month_quantity_{year}.json"
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
        pos = ss.read_item(year)
        basic_list = []
        for i in pos:
            z = Tertiary_download_structure(i.item_kpi, i.year, i.month, i.brand, i.weight_pen, i.sro, i.pen,
                                            i.quantity, i.amount_euro, i.weight_sro)
            if z.month in list and z.item_kpi == self.info_var.get():
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
        year = self.radio_var.get()
        FILENAME = f"{self.info_var.get()}_month_weight_pen_{year}.json"
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
        pos = ss.read_item(year)
        basic_list = []
        for i in pos:
            z = Tertiary_download_structure(i.item_kpi, i.year, i.month, i.brand, i.weight_pen, i.sro, i.pen,
                                            i.quantity, i.amount_euro, i.weight_sro)
            if z.month in list and z.item_kpi == self.info_var.get():
                basic_list.append(float(z.weight_pen))
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
        year = self.radio_var.get()
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
        pos = ss.read_item(year)
        basic_list = []
        for i in pos:
            z = Tertiary_download_structure(i.item_kpi, i.year, i.month, i.brand, i.weight_pen, i.sro, i.pen,
                                            i.quantity, i.amount_euro, i.weight_sro)
            if z.month in list and z.item_kpi == self.info_var.get():
                basic_list.append(float(z.quantity))
                print(basic_list)
        for i in basic_list:
            self.quantity += float(i)
        y_coord = basic_list
        print(self.quantity)
        self.specific_data.set(f'Specific data: {y_coord}'.replace('[','').replace(']',''))
        tkinter.messagebox.showinfo('INFO',
                                    f'Sales in packs: {self.quantity} pcs')
        plt.title(f'Третичные продажи в упаковках по месяцам по SKU: \n{self.info_var.get()}')
        plt.grid(True)
        plt.plot(x_coord, y_coord,  color='g')
        plt.show()
    def save_weight_pen_month_to_csv(self):
        year = self.radio_var.get()
        FILENAME = f"{self.info_var.get()}_month_weight_pen_{year}.csv"
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
        pos = ss.read_item(year)
        basic_list = []
        for i in pos:
            z = Tertiary_download_structure(i.item_kpi, i.year, i.month, i.brand, i.weight_pen, i.sro, i.pen,
                                            i.quantity, i.amount_euro, i.weight_sro)
            if z.month in list and z.item_kpi == self.info_var.get():
                basic_list.append(float(z.weight_pen))
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
        year = self.radio_var.get()
        FILENAME = f"{self.info_var.get()}_month_weight_pen_{year}.csv"
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
        pos = ss.read_item(year)
        basic_list = []
        for i in pos:
            z = Tertiary_download_structure(i.item_kpi, i.year, i.month, i.brand, i.weight_pen, i.sro, i.pen,
                                            i.quantity, i.amount_euro, i.weight_sro)
            if z.month in list and z.item_kpi == self.info_var.get():
                basic_list.append(float(z.weight_pen))
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
    def save_all_data_PROMO_items_to_csv(self):
        year = self.radio_var.get()
        FILENAME = f"Weight_pen_month_item_all_PROMO_items_{year}.csv"
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
        pos = ss.read_item(year)
        basic_list = []

        ss.save_2020_items_to_csv(FILENAME,list,year)
        tkinter.messagebox.showinfo('INFO',
                                    f'File {FILENAME} has been succesfully written!')
    def save_all_data_PROMO_items_to_csv_commas(self):
        year = self.radio_var.get()
        FILENAME = f"Weight_pen_month_item_all_PROMO_items_with_commas_{year}.csv"
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
        pos = ss.read_item_2020_w_commas(year)

        ss.save_items_2020_to_csv_with_commas(FILENAME,list,year)
        tkinter.messagebox.showinfo('INFO',
                                    f'File {FILENAME} has been succesfully written!')
    def save_otc_PROMO_items_to_csv(self):
        year = self.radio_var.get()
        FILENAME = f"Otc_PROMO_items_all_data2020_{year}.csv"
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
        pos = ss.read_item_2020_OTC(year)

        ss.save_items_otc_to_csv_2020(FILENAME,list,year)
        tkinter.messagebox.showinfo('INFO',
                                    f'File {FILENAME} has been succesfully written!')
    def save_otc_PROMO_items_to_csv_with_commas(self):
        year = self.radio_var.get()
        FILENAME = f"Otc_PROMO_items_all_data2020_with_commas_{year}.csv"
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
        pos = ss.read_item_2020_OTC_with_commas(year)

        ss.save_items_otc_to_csv_2020_with_commas(FILENAME,list,year)
        tkinter.messagebox.showinfo('INFO',
                                    f'File {FILENAME} has been succesfully written!')
    def save_rx_items_to_csv_with_commas(self):
        year = self.radio_var.get()
        FILENAME = f"Rx_items_all_data2020_with_commas_{year}.csv"
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
        pos = ss.read_item_2020_RX_with_commas(year)

        ss.save_items_RX_to_csv_2020_with_commas(FILENAME,list,year)
        tkinter.messagebox.showinfo('INFO',
                                    f'File {FILENAME} has been succesfully written!')
    def save_rx_items_to_csv(self):
        year = self.radio_var.get()
        FILENAME = f"Rx_items_all_data_2020_{year}.csv"
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
        pos = ss.read_item_2020_RX(year)

        ss.save_items_RX_to_csv_2020(FILENAME,list,year)
        tkinter.messagebox.showinfo('INFO',
                                    f'File {FILENAME} has been succesfully written!')
    def rewrite_2021_629_base(self):
        x = CBase_2021_quadra_workout()
        x.rewrite_629_2021_in_database()
        tkinter.messagebox.showinfo('INFO',
                                    f'The base local_main_base.db has been successfully updated!')
    def Update_2020_629_xlxs(self):
        x = CBase_2021_quadra_workout()
        x.save_base_629_2020_to_xlsx()
    def save_transformations_for_riga_sales_report(self):
        x = CBase_2021_quadra_workout()
        x.save_1_tramsform_for_sales_report_with_filter_to_xlsx()
        tkinter.messagebox.showinfo('INFO', f'0.transform_for_1_sales_report_with_filter.xlsx file has been successfully updated!')
        my_xlsx_excel_file = 'C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V6\\sales_report_riga\\1.Sales report with filter_new.xlsx'
        wb = xw.Book(my_xlsx_excel_file)

    def rewrite_2020_629_base(self):
        x = CBase_2021_quadra_workout()
        x.rewrite_629_2020_in_database()
        tkinter.messagebox.showinfo('INFO',
                                    f'The base local_main_base.db has been successfully updated!')
    def show_secondary_sales_by_month(self):
        list,list_quadra,year = self.radiobutton_months()
        x = CBase_2021_quadra_workout()
        list_by_month_sales_packs_euro = x.get_secondary_sales_sqlite3_for_big_table(list_quadra,year)
        x_coord = list_quadra
        y_coord_euro = []
        y_coord_packs = []
        print(list_by_month_sales_packs_euro)
        for entry in list_by_month_sales_packs_euro:
            for month in list_quadra:
                if month == entry[0]:
                    y_coord_euro.append(entry[2])
                    y_coord_packs.append(entry[1])
        plt.title(f'Total secondary sales by month euro')
        plt.grid(True)
        plt.plot(x_coord,y_coord_euro,marker='o')
        plt.show()
    def show_secondary_sales_by_month_by_item(self):
        list,list_quadra,year = self.radiobutton_months()
        item_selected = str(self.dict_item.get(str(self.info_var.get())))
        x = CBase_2021_quadra_workout()
        list_by_month_sales_packs_euro = x.get_secondary_sales_sqlite3_for_big_table_by_item(list_quadra,year,item_selected)
        x_coord = list
        y_coord_euro = []
        y_coord_packs = []

        print(list_by_month_sales_packs_euro)
        for entry in list_by_month_sales_packs_euro:
            for month in list_quadra:
                if month == entry[0]:
                    y_coord_euro.append(entry[2])
                    y_coord_packs.append(entry[1])
        plt.title(f'Total secondary sales by month {self.info_var.get()} euro')
        plt.grid(True)
        plt.plot(x_coord,y_coord_euro,marker='o')
        plt.show()





