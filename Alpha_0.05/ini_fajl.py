"""
Paraméterben megadott INI fájl tartalmának feldologozása
Készült: 2022.05.10

Paraméterek:
fajlneve = INI fájl neve (kiterjesztéssel együtt)

Visszaadott érték (return):
tömb, amiben a paraméter adatok vannak (lista)

Utolsó módosítás dátuma: 2022.06.03

"""

def read_ini_file(fajlnev):
    """
    INI fájl összes elemének beolvasása tömbbe.
    Csak azokat az elemek kerülnek a tömbbe, amelyek '[' 
    jellel kezdődnek.

    Az INI fájl végén mindig szerepelnie kell: *** Vége ***

    Utolsó módosítás dátuma: 2022.06.03
    """

    # tömb a paraméter elemeknek
    inilist = []

    # fájl megnyitása - olvasásra, text módban
    fajlopen = open(str(fajlnev), "rt", encoding="UTF-8")

    # első sor beolvasása
    msor = fajlopen.readline()

    # addig fut, amíg a végére nem ér
    while msor != "*** Vége ***":
        # print(msor)
        if msor[0] == "[":
            msor = msor[:-1]
            iniparam = msor.split("=")

            if iniparam != "\n":
                inilist.append(iniparam)

        # következő sor beolvasása
        msor = fajlopen.readline()

    fajlopen.close()

    return inilist

# INI fájl adatait tartalmazó tömb, egy adott elemének a lekérdezése
def initomb_eleme(tombneve, elemszama):
    holvan = str(tombneve[elemszama]).find(",")
    param = str(tombneve[elemszama])[holvan+3:-2]

    return param
