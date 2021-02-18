
import pyodbc

server = '62.149.15.123,1433'
database = 'medowl_grindex'
username = 'grindex'
password = 'xednirg'




cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

