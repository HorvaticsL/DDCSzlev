import pyodbc
import datumido

server = "HUNSFVAC"
database = "SAPSzallitolevelek"

#ODBC driver-t ellenőrizni kell a Windows beállításokban

cnxn = pyodbc.connect(
    #"DRIVER={ODBC Driver 18 for SQL Server};"
    "DRIVER={SQL Server Native Client 11.0};"
    "SERVER=" + server + ";"
    "DATABASE=" + database + ";"
    "Trusted_Connection=yes;"
    "Encrypt=no"
)


cursor = cnxn.cursor()
# cursor.execute("SELECT * FROM Szallitolevelek WHERE Incoterms='CPT'")

# Egy új sor hozzáadása a táblához
SAPkod = "102556"
Csomagolas = "Palettás"
TermekFajta = "CEM II/B-M(V-LL) 32,5N"
KontrollingSorrend = "14"
KontrollingAzon = "T080 II/B-M (V-L) 32,5 N"
TermekCsoport = "Cement"
Megnevezes = "CEM II/B-M(V-LL) 32,5N 25KG 1,4T PA"
Szulfat = "Kezelt"
Megjegyzes = ""
FelvezetesDatuma = datumido.mainap("kotojel")

sql = (
    "INSERT INTO SAPCikkek VALUES( '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s')"
    % (
        SAPkod,
        Csomagolas,
        TermekFajta,
        KontrollingSorrend,
        KontrollingAzon,
        TermekCsoport,
        Megnevezes,
        Szulfat,
        Megjegyzes,
        FelvezetesDatuma,
    )
)

number_of_rows = cursor.execute(sql)
cnxn.commit()


cursor.execute("SELECT * FROM SAPCikkek")

for i in cursor:
    print(i)

cnxn.close()
