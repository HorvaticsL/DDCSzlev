"""
DDCSzlev kezdőfájl
Készült: 2022.05.31

Utolsó módosítás dátuma: 2022.06.15
verzió: 01

"""
import ini_fajl as inif
import konyvtar_kezeles as dirkez
import naplozas
import excel_feldolgozasa as exelfeld
import sys
import ctypes


def foprogram():
    # Utolsó módosítás dátuma: 2022.06.15

    # Naplózás elindítása
    naplo = naplozas.naplolog()
    naplo.info("Kezdődik program")

    # INI fájl adatok beolvasása, tömbbe
    try:
        akt_konyvtar = dirkez.aktualismappa()
        ini_fajlneve = "\\DDCSzlev.ini"
        inifajl = str(akt_konyvtar) + ini_fajlneve

        # initomb változó tartalmazza az összes paraméter elemet
        initomb = inif.read_ini_file(inifajl)
    except FileNotFoundError:
        naplo.error("Az INI fájl nem található: %s", str(inifajl))
        naplo.warning("A program leállt!")
        MessageBox = ctypes.windll.user32.MessageBoxW
        MessageBox(
            None,
            "Az INI Fájl nem található!\n\nRészletek a naplófájlban!",
            "",
            0,
        )
        sys.exit(0)
    except Exception as e:
        naplo.error('Ismeretlen hiba típusa, leírás: %s: %s',
                    str(type(e)), str(e))
        naplo.warning("A program leállt!")
        MessageBox = ctypes.windll.user32.MessageBoxW
        MessageBox(
            None,
            "Ismeretlen hiba!\n\nRészletek a naplófájlban!",
            "",
            0,
        )
        sys.exit(0)

    # egy elem lekérdezése, jelen esetben a nullás
    #exefajl_neve = inif.initomb_eleme(initomb, 0)
    # print(exefajl_neve)

    # teljes lista
    # for i in range(len(initomb)):
    #    print( str(i) + " - " + str(initomb[i]))

    # **** INI fájl... blokk vége ****

    # forrás - export - EXCEL fájl átalakítása
    exelfeld.excelfajl_modositas(initomb)

    # adatbázisba mentés
    # ellenőrzés

    # program futásának vége
    naplo.info("Befejeződött program")


foprogram()
