'''
Adatbázis kapcsolatokat tartalmazó fájl

Készült: 2022.08.05

Utolsó módosítás dátuma: 2022.08.05
Verzió: 1
'''

import pyodbc

# ODBC kapcsolat
def ODBC_Kapcsolat(driver, sqlserver, sqldatabase):
    '''
    ODBC Connectstring összeállítása
    Adatbázis megnyitása

    Utolsó módosítás dátuma: 2022.08.04

    '''

    kapcsolat = pyodbc.connect(
        #"DRIVER={ODBC Driver 18 for SQL Server};"
        "DRIVER={" + driver + "};"
        "SERVER=" + sqlserver + ";"
        "DATABASE=" + sqldatabase + ";"
        "Trusted_Connection=yes;"
        "Encrypt=no"
    )

    return kapcsolat

    # ***** ODBC_Kapcsolat VÉGE