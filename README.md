# CSV Bankas Izrakstu Pārveidotājs Budžeta Veidnē
Šī programa saņem CSV (Comma Seperated Values) failu, kuru var izprintēt no bankas pārskatiem un ļauj doto informāciju pārvērst pārskatāmā budžeta formātā. Programma pašlaik ir eksluzīvi pārbaudītai tikai uz Swedbank bankas izrakstiem, tāpēc var nestrādāt uz CSV iegūtiem no citiem šop pakalpojumu sniedzējiem.

## Izmantotās bibliotēkas un klases
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
* __"Debets/Kredīts"__ kolonai <ins>debetšu gadījumā jāsatur simbolu "D"</ins>, bet <ins>kredītu gadījumā (iemaksas/ieņēmumi) simbolu "K"</ins>.
* Faila otrajai rindai (rindai pec galveņu nosaukumiem) vajadzētu saturēt balanca inicializācijas vērtību. (Pēcākiem failiem arī vajadzētu saturēt šadu rindu, bet tā tiks ignorēta.)
* Pēdējās 3 rindas tiek dzēstas, jo Swedbank ģenerētais CSV fails šajās rindās satur kopējā apgrozījuma un beigu atlikuma vērtības, kas programmas ģenerētajā Excel tiek aprēķinātas patstāvīgi.
Pārējās vērtības programma pašlaik nepielieto un to formāts nav svarīgs.
