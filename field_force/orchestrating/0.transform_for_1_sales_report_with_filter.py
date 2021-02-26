import xlwings as xw

def main():
    wb = xw.Book.caller()
    wb.sheets[0].range('A1').value = 'Hello World!'