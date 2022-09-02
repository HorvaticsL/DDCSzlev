from datetime import date
from datetime import datetime


def pontosido(tipus):
    """
    Pontos időt adja vissza
    Készült: 2022.05.13

    Paraméterek:
    datumora = dátum (pontokkal) és pontos idő megjelenítése, pl.: 2022.05.13 12:02:02
    datumoraegybe = dátum (pontokkal) és pontos idő megjelenítése, pl.: 20220513120202
    pontosido = digitális óra formátum megjelenítés, pl.: 12:02:22
    ora = óra
    perc = perc
    nincs paraméter ("") = szöveg formátum megjelenítés, pl.: 120222

    Visszaadott érték (return):
    string, ami tartalmazza a pontosidőt

    Utolsó módosítás dátuma: 2022.05.13

    """

    idopont = datetime.now()

    if tipus == "datumora":
        idostring = idopont.strftime("%Y.%m.%d %H:%M:%S")
    elif tipus == "datumoraegybe":
        idostring = idopont.strftime("%Y%m%d%H%M%S")
    elif tipus == "pontosido":
        idostring = idopont.strftime("%H:%M:%S")
    elif tipus == "ora":
        idostring = idopont.strftime("%H")
    elif tipus == "perc":
        idostring = idopont.strftime("%M")
    else:
        idostring = idopont.strftime("%H%M%S")

    return str(idostring)


def mainap(tipus):
    """
    Napi dátumot adja vissza
    Készült: 2022.05.13

    Paraméterek:
    ev = az évszámot adja meg
    honap = a hónap számát adja meg
    evhonap = évszám és a hónap, pl: 202205
    evhonapkotojel = évszám és a hónap, pl: 2022-05
    pont = magyar fomátum megjelenítés, pl.: 2022.05.13
    kotojel = kötőjel formátum megjelenítés, pl.: 2022-05-13
    nincs paraméter ("") = szöveg formátum megjelenítés, pl.: 20220513

    Visszaadott érték (return):
    string, ami tartalmazza a pontosidőt

    Utolsó módosítás dátuma: 2022.05.13

    """

    datum = date.today()

    if tipus == "ev":
        # YY
        datumstring = datum.strftime("%Y")
    elif tipus == "honap":
        # mm
        datumstring = datum.strftime("%m")
    elif tipus == "evhonap":
        # YYmm
        datumstring = datum.strftime("%Y%m")
    elif tipus == "evhonapkotojel":
        # YY-mm
        datumstring = datum.strftime("%Y-%m")
    elif tipus == "pont":
        # YY.mm.dd
        datumstring = datum.strftime("%Y.%m.%d")
    elif tipus == "kotojel":
        # YY-mm-dd
        datumstring = datum.strftime("%Y-%m-%d")
    else:
        # YYmmdd
        datumstring = datum.strftime("%Y%m%d")

    return str(datumstring)
