"""
Excel fájl tartalmának feldolgozása
A fájlban lévő adatokat úgy kell átalakítani, hogy
az adatbázisba lementhetők legyenek - helységnevek, fuvar-, útdíjak, stb.
Készült: 2022.06.13

Utolsó módosítás dátuma: 2022.06.15
verzió: 01

"""
#import openpyxl
from openpyxl import load_workbook
import ini_fajl as inif
import naplozas
import sys
import ctypes


def excelfajl_modositas(initomb):
    # Utolsó módosítás dátuma: 2022.06.15
    # Naplózás beállítása
    excelnaplo = naplozas.naplolog()
    excelnaplo.info('Exel forrásfájl feldolgozása elindult')

    # INI fájladatok beolvasása
    prgneve = inif.initomb_eleme(initomb, 0)
    exelfile = inif.initomb_eleme(initomb, 3)
    excelmappa = inif.initomb_eleme(initomb, 4)
    wbsheetneve = inif.initomb_eleme(initomb, 5)
    #ofajl=str(excelmappa) + str(exelfile)
    ofajl = str(exelfile)

    excelnaplo.info('Excelfájl: %s', str(exelfile))

    print(exelfile)
    print(excelmappa)
    print(wbsheetneve)
    print(ofajl)
    try:
        excelnaplo.info('Excel fájl megnyitása')
        wb = load_workbook(filename = ofajl)
        rangesheet = wb[wbsheetneve]
        print(rangesheet['A1'].value)

        # wb.save(ofajl)
        excelnaplo.info('Excel fájl elmentve')
    except FileNotFoundError:
        excelnaplo.error("Az XLSX fájl nem található: %s", str(ofajl))
        excelnaplo.warning("A program leállt!")
        MessageBox = ctypes.windll.user32.MessageBoxW
        MessageBox(
            None,
            "Az XLSX Fájl nem található!\n\nRészletek a naplófájlban!",
            prgneve,
            0,
        )
        sys.exit(0)
    except Exception as e:
        excelnaplo.error(
            'Ismeretlen hiba típusa, leírás: %s: %s', str(type(e)), str(e))
        excelnaplo.warning("A program leállt!")
        MessageBox = ctypes.windll.user32.MessageBoxW
        MessageBox(
            None,
            "Ismeretlen hiba!\n\nRészletek a naplófájlban!",
            prgneve,
            0,
        )
        sys.exit(0)

    # Helységnév átalakítása
    # Incoterms elágazása
    # EXW - fuvar-, útdíj és fuvar-útdíj nullázása
    # CPT - VAN útdíj, NINCS útdíj
    # VAN útdíj = képelettel kiszámolni az nettó fuvardíjat
    # NINCS útdíj - Útdíj megkeresése, képlettel kiszámítani a nettó fuvardíjat
    # Fájl mentése más néven (megmaradjon az eredetifájl), kilépés

    excelnaplo.info('Exel forrásfájl feldolgozása befejeződött')
