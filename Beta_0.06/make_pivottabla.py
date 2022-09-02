"""
Az átalakított szállítólevelek adatok alapján
PIVOT TABLE-k létrehozása, az adatok ellenőrzésére
Készült: 2022.07.19

Utolsó módosítás dátuma: 2022.07.19
verzió: 01

"""
import pandas as pd
import sys
import ctypes


def pivottabla(prgneve, forrasfile, savefile, logfile):
    # Utolsó módosítás dátuma: 2022.07.19
    logfile.info('Pivot táblák kezdése')

    try:

        # Termékcsoport, gyár, tonna, összesítve
        logfile.info('Termékcsoport, gyár, tonna, összesítve')

        sheet_df_dictonary = pd.read_excel(
            forrasfile, engine='openpyxl', sheet_name=['Sheet1'], skiprows=0)
        Sheet1 = sheet_df_dictonary['Sheet1']

        # Pivoted into Sheet1
        tmp_df = Sheet1[['Gyar', 'Tonna', 'TermekCsoport']]
        pivot_table = tmp_df.pivot_table(
            index=['TermekCsoport', 'Gyar'],
            values=['Tonna'],
            aggfunc={'Tonna': ['sum']}
        )

        pivot_table.set_axis([' - '.join(col).rstrip('_')
                             for col in pivot_table.keys()], axis=1, inplace=True)

        Sheet1_pivot = pivot_table.reset_index()
        Sheet1_pivot.style.format('{:.3f}')

        # Renamed Sheet1_pivot to Tonna
        logfile.info('Munkalap átnevezése: Tonna')
        Tonna = Sheet1_pivot
        logfile.info('KÉSZ: Termékcsoport, gyár, tonna, összesítve')

        # **** Termékcsoport, gyár, tonna, összesítve VÉGE
        # Fuvardíjak pivot table
        Sheet2 = sheet_df_dictonary['Sheet1']

        tmp_df2 = Sheet2[['Incoterms', 'Csomagolas', 'TermekCsoport',
                          'Kimutatasnev', 'Tavolsag', 'Tonna', 'FuvarUtdijBrutto', 'Utdij', 'ATKm']]
        pivot2_table = tmp_df2.pivot_table(
            index=['Incoterms', 'Csomagolas', 'TermekCsoport', 'Kimutatasnev'],
            values=['Tavolsag', 'Tonna', 'Utdij', 'FuvarUtdijBrutto', 'ATKm'],
            aggfunc={'Tavolsag': ['sum'], 'Tonna': ['sum'], 'Utdij': [
                'sum'], 'FuvarUtdijBrutto': ['sum'], 'ATKm': ['sum']}
        )

        #pivot2_table.set_axis([flatten_column_header(col) for col in pivot_table.keys()], axis=1, inplace=True)
        #pivot2_table.set_axis([' - '.join(col).rstrip('_') for col in pivot_table.keys()], axis=1, inplace=True)
        Sheet2_pivot = pivot2_table.reset_index()

        # Changed Csomagolas to dtype str
        Sheet2_pivot['Csomagolas'] = Sheet2_pivot['Csomagolas'].astype('str')

        # Changed FuvarUtdijBrutto sum to dtype float
        Sheet2_pivot['FuvarUtdijBrutto'] = Sheet2_pivot['FuvarUtdijBrutto'].astype(
            'float')

        # Changed Utdij sum to dtype float
        Sheet2_pivot['Utdij'] = Sheet2_pivot['Utdij'].astype('float')

        # Filtered Incoterms
        Sheet2_pivot = Sheet2_pivot[Sheet2_pivot['Incoterms'].str.contains(
            'CPT', na=False)]

        # Renamed Sheet2_pivot to Fuvardij
        logfile.info('Munkalap átnevezése: Fuvardij')
        Fuvardij = Sheet2_pivot
        logfile.info('KÉSZ: Fuvardíj Pivot Talbe')

        # **** Fuvardíjak pivot table VÉGE
        # munkalapok fájlba írása, Excel fájl bezárása
        writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
        Tonna.to_excel(writer, sheet_name='Tonna')
        Fuvardij.to_excel(writer, sheet_name='Fuvardij')
        logfile.info('Tonna munkalap hozzáadása a fájlhoz')
        logfile.info('Fuvardij munkalap hozzáadása a fájlhoz')

        writer.save()
        writer.close()
        logfile.info('Fájl mentése: %s, bezárása', str(savefile))

    # **** TRY VÉGE

    except Exception as merror:
        logfile.error(
            'Ismeretlen hiba típusa, leírás: %s: %s', str(type(merror)), str(merror))
        logfile.warning("A program leállt!")
        MessageBox = ctypes.windll.user32.MessageBoxW
        MessageBox(
            None,
            "Ismeretlen hiba!\n\nRészletek a naplófájlban!",
            prgneve,
            0,
        )
        sys.exit(0)

    logfile.info('Pivot Table-k elkészültek, mentés fájlba megtörtént.')
