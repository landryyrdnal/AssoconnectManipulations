import GetDatabase
import pandas as pd
import pyodbc

# Constantes
driver = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
path = r'DBQ=D:\Dropbox\Académie\Programmation\Assoconnect - Access\REPERTOIRE ABNT_THEILAIA - Maj 2022-05-21.mdb;'

con_string = driver+path

db = GetDatabase.get_assoconnect_data_base()


def checking_access_installation():
    msa_drivers = [x for x in pyodbc.drivers() if 'ACCESS' in x.upper()]
    if not "Microsoft Access Driver (*.mdb, *.accdb)" in msa_drivers:
        print("veuillez installer MS Access correctement sur cet appareil")
    else:
        print("MS Access correctement installé sur cet appareil")


def df_to_db_access(db:pd.DataFrame):
    conn = pyodbc.connect(driver+path)
    cursor = conn.cursor()

# checking_access_installation()
df_to_db_access(db)

