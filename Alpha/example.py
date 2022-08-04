#*** kód részletek

# Helységnév átalakítása
# ez nem működött, de jó lesz példának

eng_helysegkeplet = "=IF(P" + sr + "='HU';IFERROR(VLOOKUP(--M" + sr + ";" + fajlhivatkozas + ";2;FALSE);O" + sr + ";)O" + sr + ")"
#hu_helysegkeplet = "=HA(P" + sr + "='HU';HAHIBA(FKERES(--M" + sr + ";" + fajlhivatkozas + ";2;HAMIS);O" + sr + ");O" + sr + ")"
hu_helysegkeplet = f'=HA(P{sr}="HU";HAHIBA(FKERES(--M{sr};{fajlhivatkozas};2;HAMIS);O{sr});O{sr})'
#=HA(P2="HU";HAHIBA(FKERES(--M2;'n:\\\\Logisztika\\\\Megrendelés és fuvardíj ellenőrzések\\\\[!Megrendelés ellenőrzés v0.88a.xlsm]Alapadatok'!$G$4:$H$500;2;HAMIS);O2);O2)
munkalap["AL" + sr].value = hu_helysegkeplet.format(munkalap["AL"+sr].row)

