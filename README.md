# CSV Bankas Izrakstu Pārveidotājs Budžeta Veidnē
Šī programa saņem CSV (Comma Seperated Values) failu, kuru var izprintēt no bankas pārskatiem un ļauj doto informāciju pārvērst pārskatāmā budžeta formātā. Programma pašlaik ir eksluzīvi pārbaudītai tikai uz Swedbank bankas izrakstiem, tāpēc var nestrādāt uz CSV iegūtiem no citiem šop pakalpojumu sniedzējiem.

## Izmantotās biliotēkas un klases
* LinkedList - Paštaisīta klase, kas funkcionē kā vienvirzienā saistītā sarakta datu struktūra ar dažām modifikācījām:
* Queue - Paštaisīta klase, kas funckionē kā rindas datu struktūra.
* xlwings - Bibliotēka, kas ļauj paplašināti apstrādāt Excel failus.
* csv - Bibliotēka, kas dod funkcionalitāti CSV failu apstrādei.
* os - Bibliotēka, kas ļauj strādāt ar Windows failiem. 

## Ievades faila prasības
* Failam nav obligāti jābūt iegūtam no Swedbankas, bet jābūt balstītam uz sekojošo CSV galveni:
  ```
  "Klienta konts","Ieraksta tips","Datums","Saņēmējs/Maksātājs","Informācija saņēmējam","Summa","Valūta","Debets/Kredīts","Arhīva kods","Maksājuma veids","Refernces numurs","Dokumenta numurs"
  ```
* "Datums" kolonai jābūt formatētam kā "DD.MM.YYYY".
* "Summa" kolonai ir jāsatur skaitliska vērtība (pat ja formatēta kā string). 
* "Debets/Kredīts" kolonai izmaksas (debit) gadījumā jāsatur simbolu "D", bet iemaksu (kredit) gadījumā simbolu "K".
* Pārējās vērtības programa pašlaik nepielieto.
