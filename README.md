# CSV Bankas Izrakstu Pārveidotājs Budžeta Veidnē
Šī programa saņem CSV (Comma Seperated Values) failu, kuru var izprintēt no bankas pārskatiem un ļauj doto informāciju pārvērst pārskatāmā budžeta formātā. Programma pašlaik ir eksluzīvi pārbaudītai tikai uz Swedbank bankas izrakstiem, tāpēc var nestrādāt uz CSV iegūtiem no citiem šop pakalpojumu sniedzējiem.

## Izmantotās biliotēkas un klases
* __LinkedList__ - <ins>Paštaisīta klase</ins>, kas funkcionē kā vienvirzienā saistītā sarakta datu struktūra ar dažām modifikācījām:
* __Queue__ - <ins>Paštaisīta klase</ins>, kas funckionē kā rindas datu struktūra.
* __xlwings__ - <ins>Bibliotēka</ins>, kas ļauj paplašināti apstrādāt Excel failus. Svarīgi, šij biblotēkai ir vajadzīgs imports (https://pypi.org/project/xlwings/0.3.2/)
* __csv__ - <ins>Bibliotēka</ins>, kas dod funkcionalitāti CSV failu apstrādei.
* __os__ - <ins>Bibliotēka</ins>, kas ļauj strādāt ar Windows failiem. 

## Ievades faila prasības
* Failam nav obligāti jābūt iegūtam no Swedbankas, bet jābūt balstītam uz sekojošo __CSV galveni__:
  ```
  "Klienta konts","Ieraksta tips","Datums","Saņēmējs/Maksātājs","Informācija saņēmējam","Summa","Valūta","Debets/Kredīts","Arhīva kods","Maksājuma veids","Refernces numurs","Dokumenta numurs"
  ```
* __"Datums"__ kolonai jābūt formatētam kā <ins>"DD.MM.YYYY"</ins>.
* __"Informācija saņēmējam"__ kolonai ir jāsatur <ins>īss paskaidrojums</ins> par darījumu.
* __"Summa"__ kolonai ir jāsatur <ins>skaitliska vērtība</ins> (pat ja tā formatēta kā string). 
* __"Debets/Kredīts"__ kolonai izmaksas (debit) gadījumā jāsatur simbolu <ins>"D"</ins>, bet iemaksu (kredit) gadījumā simbolu <ins>"K"</ins>.

Pārējās vērtības programma pašlaik nepielieto un to formāts nav svarīgs.
