"""
DDCSzlev kezdőfájl
Készült: 2022.05.31

Utolsó módosítás dátuma: 2022.09.15
verzió: 02

"""
import sys
import ctypes
import generate_config as makecfg
import konyvtar_kezeles as dirkez
import naplozas
import excel_feldolgozasa as exelfeld
import db_SAPCikkek as dbsapcikk
import db_SAPSzlev as dbsapszlev


def foprogram():
    # Utolsó módosítás dátuma: 2022.06.15

    # Naplózás elindítása
    logfile = naplozas.naplolog()
    logfile.info("Kezdődik program")

    # INI fájl létrehozása
    try:
        akt_konyvtar = dirkez.aktualismappa()
        ini_fajlneve = "\\DDCSzlev_v2.ini"
        inifajl = str(akt_konyvtar) + ini_fajlneve
        # INI fájl létrehozása a futási könyvtárban
        makecfg.make_config_file(inifajl)

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
    # futás végén visszaadja az átalakított Excel fájl
    # nevét, elérési úttal együtt
    saveas_Excelfile = exelfeld.excelfajl_modositas(inifajl, logfile)

    # adatbázisba mentés - SAP cikkek
    dbsapcikk.SAPCikkek_feltoltese(inifajl, logfile)
    # adatbázisba mentés - SAP szállítólevelek
    dbsapszlev.SAPSzlev_feltoltese(saveas_Excelfile, inifajl, logfile)

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
