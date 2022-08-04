"""
DDCSzlev kezdőfájl
Készült: 2022.05.31

Utolsó módosítás dátuma: 2022.06.15
verzió: 01

"""
import sys
import ctypes
import ini_fajl as inif
import konyvtar_kezeles as dirkez
import naplozas
import excel_feldolgozasa as exelfeld

def foprogram():
    # Utolsó módosítás dátuma: 2022.06.15

    # Naplózás elindítása
    logfile = naplozas.naplolog()
    logfile.info("Kezdődik program")

    # INI fájl adatok beolvasása, tömbbe
    try:
        akt_konyvtar = dirkez.aktualismappa()
        ini_fajlneve = "\\DDCSzlev.ini"
        inifajl = str(akt_konyvtar) + ini_fajlneve

        # initomb változó tartalmazza az összes paraméter elemet
        initomb = inif.read_ini_file(inifajl)
    except FileNotFoundError:
        logfile.error("Az INI fájl nem található: %s", str(inifajl))
        logfile.warning("A program leállt!")
        MessageBox = ctypes.windll.user32.MessageBoxW
        MessageBox(
            None,
            "Az INI Fájl nem található!\n\nRészletek a naplófájlban!",
            "",
            0,
        )
        sys.exit(0)
    except Exception as merror:
        logfile.error('Ismeretlen hiba típusa, leírás: %s: %s',
                    str(type(merror)), str(merror))
        logfile.warning("A program leállt!")
        MessageBox = ctypes.windll.user32.MessageBoxW
        MessageBox(
            None,
            "Ismeretlen hiba!\n\nRészletek a naplófájlban!",
            "",
            0,
        )
        sys.exit(0)

    # **** INI fájl... blokk vége ****

    # forrás - export - EXCEL fájl átalakítása
    exelfeld.excelfajl_modositas(initomb, logfile)

    # adatbázisba mentés
    # ellenőrzés

    # program futásának vége
    MessageBox = ctypes.windll.user32.MessageBoxW
    MessageBox(
        None,
        "A program befejeződött, kilépés!",
        "",
        0,
    )
    logfile.info("Befejeződött program")


foprogram()
