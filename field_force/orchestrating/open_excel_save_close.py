import xlwings as xw
my_xlsx_excel_file = 'C:\\Users\\Anastasia Siedykh\\Documents\\Backup\\KPI report\\MODULE SET V4\\TRANSFORM\\1. DISTRIBUTORS PLAN-FACT PACKS-EURO-PERCENTAGE 2021.xlsx'
wb = xw.Book(my_xlsx_excel_file)

sh = wb.sheets['2. DISTRIBUTORS PLAN-FACT']

