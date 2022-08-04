flowchart TD
%% Itt kezdődik %%

FO([Forrás Excel<br/>fájl feldolgozása])
style FO fill:#2374f7,stroke:#000,stroke-width:2px,color:#fff

FO --- |Adatok a 2. sortól|SOR(Soronkénti<br/>feldolgozás)
SOR --- |Külső táblából|ANYAG[[Anyagkód alapján megtalálni a\n anyag megnevezést]]
SOR --- |Vevőkód alapján, külső táblából|HELY[[Helységnevek módosítása]]
SOR --- FUVD(Fuvar- útdíj és Fuvar- és útdíj számítás)
FUVD --- CPT(Incoterms = CPT)
FUVD --- EXW(Incoterms = EXW)
CPT --- KOERT[[Kondíció érték - fuvardíj - kiszámmítása]]
KOERT --- |Útdíj <> 0|FELT1(Van útdíj összeg)
KOERT --- |Útdíj = 0|FELT2(Nincs útdíj összeg)
FELT1 --- VANUT[[< Kondícióérték > =\n < Furvar + Útdíj > - < Útdíj kondícióérték >]]
FELT2 --- NOUT[[Útdíj kiszámítása]]
NOUT --- |1 - ÖML - HU - ZF49 - JOHANS - gyár|JOHANS(Útdíj keresés + beírás)
JOHANS --- JHKERT[[< Kondícióérték > =\n < Furvar + Útdíj > - < Útdíj kondícióérték >]]
NOUT --- |2 - ÖML - HU - ZF49 - 105680 - gyár|KEMENCE(Útdíj keresés + beírás)
KEMENCE --- KEMERT[[< Kondícióérték > =\n < Furvar + Útdíj > - < Útdíj kondícióérték >]]
NOUT --- |3 - ÖML - SK - ZF49 - Hosszúkód fuvaros - VÁC|HOSSZU(Útdíj keresés + beírás)
HOSSZU --- HOERT[[< Kondícióérték > =\n < Furvar + Útdíj > - < Útdíj kondícióérték >]]
NOUT --- |4 - ÖML - SK - ZF49 - Rövidkódos fuvaros - VÁC|ROVID(Útdíj keresés + beírás)
ROVID --- ROERT[[< Kondícióérték > =\n < Furvar + Útdíj > - < Útdíj kondícióérték >]]
NOUT --- |5 - PAL - SK - ZF49 - NO:SpeedLine,NortSped,Petrányi - VÁC|SKPAL(Útdíj keresés + beírás)
SKPAL --- SKPERT[[< Kondícióérték > =\n < Furvar + Útdíj > - < Útdíj kondícióérték >]]
NOUT --- |6 - PAL - SK - ZF49 - CSAK Petrányi - VÁC|SKPET(Útdíj keresés + beírás)
SKPET --- SKPETERT[[< Kondícióérték > =\n < Furvar + Útdíj > - < Útdíj kondícióérték >]]
NOUT --- |7 - PAL - AT - ZF49 - BEREMEND|ATPAL(Útdíj keresés + beírás)
ATPAL --- ATPALERT[[< Kondícióérték > =\n < Furvar + Útdíj > - < Útdíj kondícióérték >]]