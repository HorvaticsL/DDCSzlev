"""
INI fájl létrehozása
Készült: 2022.09.15

Utolsó módosítás dátuma: 2022.09.15
verzió: 01

"""
import configparser


def make_config_file(inifile):

    # változó létrehozása
    cfgfile = configparser.ConfigParser(allow_no_value=True)

    # Csoport hozzáadása
    cfgfile.add_section('EXE-File')
    # Tételek a csoporthoz
    cfgfile.set('EXE-File', '; EXE fájl adatai')
    cfgfile.set('EXE-File', 'EXE-File neve', 'ExcelToSQL')
    cfgfile.set('EXE-File', 'EXE-File verzio', 'alpha_0.05')
    cfgfile.set('EXE-File', 'EXE-File datuma', '2022.09.02')

    # CSoport hozzáadása (2. verzió)
    cfgfile['SAP-Excel'] = {
        '; SAP szállítólevél adatok': '',
        'SAPExcel file neve': 'HL_2021_01(január)_original.XLSX',
        'SAPExcel mappa': 'n:/Logisztika/Programok/ExcelToSQL/Betöltőfájl/',
        'SAPExcel munkalap': 'Sheet1'
    }

    cfgfile['Helysegnev-EXCEL'] = {
        '; Helységnév adatok': '',
        'helysegnev mappa': 'n:/Logisztika/Megrendelés és fuvardíj ellenőrzések/',
        'helysegnev file neve': '!Megrendelés ellenőrzés v0.88b.xlsm',
        'helysegnev munkalap': 'Alapadatok',
        'helysegnev tartomany bal-felso': 'G4',
        'helysegnev tartomany jobb-also': 'H500'
    }

    cfgfile['Utdij-EXCEL'] = {
        '; Útddíj adatok a SAP szállítólevelekhez (TIR-UTAK, egyszeres)': '',
        'utdij mappa': 'n:/Logisztika/Adminisztracio/Cement/Archív/',
        'utdij file neve': 'Belföld-Bere,Vác-km és útdíj adatok 20210101.xlsx',
        'utdij BERE-munkalap': 'Beremend',
        'utdij VAC-munkalap': 'Vác',
        'utdij SK-munkalap': 'SK',
        'utdij tartomany bal-felso': 'A7',
        'utdij tartomany jobb-also': 'W5000',
        'utdij Kimutatasnev-munkalap': 'Kimutatásnév',
        'utdij Kimutatasnev tartomany bal-felso': 'A2',
        'utdij Kimutatasnev tartomany jobb-also': 'D500',
        'utdij SAPCikkek-munkalap': 'SAP cikkek',
        'utdij SAPCikkek tartomany bal-felso': 'A2',
        'utdij SAPCikkek tartomany jobb-also': 'I500'
    }

    cfgfile['SQL-Server'] = {
        '; SQL Server kapcsolati adatok': '',
        'ODBC Driver': 'SQL Server Native Client 11.0',
        'SQL Server name': 'HUNSFVAC',
        'SQL Database name': 'SAPSzallitolevelek',
        'SQL Szallitolevel adattábla': 'Szallitolevelek',
        'SQL SAP cikkek adattabla': 'SAPCikkek',
        'SQL TermekAthordas adattabla': 'TermekAthordas'
    }

    cfgfile['SAP-Gyarak'] = {
        '; SAP-ban lévő gyárkódok': '',
        'SAP Vác gyár': 'HU11',
        'SAP Beremend gyár': 'HU12',
        'SAP Ecser gyár': 'HU14'
    }

    cfgfile['Fuvardij-Artablak'] = {
        '; ZF49 (Projekt-árakhoz) tatozó údij adatok': '',
        'artabla mappa': 'n:/Logisztika/Cement fuvardíjak/Archív - 2021/',
        'artabla VAC-SK-OML file neve': 'SK-Ömlesztett_2021.01.01_Fuvardíj&Útdíj.xlsm',
        'artabla VAC-SK-OML munkalap': 'Részletes adatok',
        'artabla VAC-SK-OML tartomany bal-felso': 'J12',
        'artabla VAC-SK-OML tartomany jobb-also': 'CZ500',
        '; az SK-ÖML ártáblában lévő oszlopszámból ki kell vonni egyet':'',
        'artabla VAC-SK-OML tartomany utdij2x': '66',
        'artabla VAC-SK-OML tartomany utdij1x': '67',
        'artabla VAC-SK-PAL file neve': 'SK-Zsákos_Vác_2021.01.01_Fuvardíj&Útdíj.xlsm',
        'artabla VAC-SK-PAL munkalap': 'Alaptábla',
        'artabla VAC-SK-PAL tartomany bal-felso': 'J12',
        'artabla VAC-SK-PAL tartomany jobb-also': 'BZ500',
        '; az SAK-PAL ártáblában lévő oszlopszámból ki kell vonni egyet':'',
        'artabla VAC-SK-PAL tartomany utdij2x': '26',
        'artabla VAC-SK-PAL tartomany utdij12x': '25',
    }

    cfgfile['SAPkodok'] = {
        '; SAP kódok a kivételekhez': '',
        'Fuvaros-rJOHANS': '1851683',
        'Fuvaros-hJOHANS': '1832590287',
        'Fuvaros-SpeedLine': '139666',
        'Fuvaros-Nordsped': '1850020',
        'Fuvaros-Petranyi': '196864',
        'Cikk-Kemencepor': '105680'
    }

    # Config fájl mentése
    #with open(r'ExcelToSQL.ini', 'w') as configfileObj:
    with open(inifile, 'w') as configfileObj:
        cfgfile.write(configfileObj)
        configfileObj.flush()
        configfileObj.close()

    '''
    print('config fájl létrehozva')

    # Config fájl tartalmának listázása
    readfile = open('ExcelToSQL.ini', 'r')
    content = readfile.read()
    print(content)
    readfile.flush()
    readfile.close()
    '''
