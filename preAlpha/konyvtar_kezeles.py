import os


def mappatartalma(mappa, szures):
    """
    Paraméterben megadott könyvtár tartalmának beolvasása
    Készült: 2022.04.29

    Paraméterek:
    mappa = a könyvtár teljes elérési útvonala ("\" cserélni kell "/"-re)
    szures = keresett kiterjesztés a listába (ha üres, akkor az összes fájl kell)

    Rekord adatai:
    fajlnev
    fajlmeret
    fajldatuma

    Visszaadott érték (retunn):
    tömb, amiben a rekordok vannak

    Utolsó módosítás dátuma: 2022.06.03

    """

    # rekordokat tartalmazó lista
    mrekordlist = []
    # lista ürítése
    mrekordlist.clear()

    # rekord ami tárolja a fájladatokat
    class mrekord():
        pass

    # van-e szűrési feltétel
    if len(szures) > 1:
        mvanszures = True
    else:
        mvanszures = False

    # mappa elérési útvonal végén van-e "/" jel
    mmappahossz = len(mappa)
    if mappa[mmappahossz-1] != "/":
        mappa = mappa + "/"

    # mappában lévő fájlok lekérdezése
    mfajlok = os.listdir(mappa)

    # üres rekord-lista létrehozása
    for i in range(len(mfajlok)):
        mrekordlist.append(mrekord())

    # lista első eleme a nulla
    mlistindex = 0
    for mfajl in mfajlok:
        if mvanszures:
            if mfajl.endswith(szures):
                mrekordlist[mlistindex].fajlnev = mfajl
                mrekordlist[mlistindex].fajlmeret = os.path.getsize(
                    mappa + mfajl)
                mrekordlist[mlistindex].fajldatuma = os.path.getmtime(
                    mappa + mfajl)

        if not mvanszures:
            mrekordlist[mlistindex].fajlnev = mfajl
            mrekordlist[mlistindex].fajlmeret = os.path.getsize(mappa + mfajl)
            mrekordlist[mlistindex].fajldatuma = os.path.getmtime(
                mappa + mfajl)

        mlistindex += 1

    return mrekordlist


def aktualismappa():
    """
    Aktuális mappa
    Készült: 2022.05.10

    Paraméterek:

    Visszaadott érték (return):
    mappa teljes útvonallal - string

    Utolsó módosítás dátuma: 2022.06.03

    """
    m_mappa = os.getcwd()

    return m_mappa
