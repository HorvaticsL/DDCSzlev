import logging
import datumido as di
import konyvtar_kezeles as akt_mappa


def naplolog():

    aktmappa = akt_mappa.aktualismappa()
    logfileneve = ""
    datumstring = ""
    # Dátum értéke, elválasztók nélkül
    datumstring = di.mainap("")

    logfileneve = str(aktmappa) + "/Naplo/" + "Naplo_" + datumstring + ".log"

    # print(logfileneve)

    bejegyzes = logging.getLogger("__name__")
    bejegyzes.setLevel(logging.DEBUG)

    fh = logging.FileHandler(logfileneve, encoding="UTF-8")
    fh.setLevel(logging.DEBUG)

    bejegyzesformatum = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    fh.setFormatter(bejegyzesformatum)

    bejegyzes.addHandler(fh)

    return bejegyzes
