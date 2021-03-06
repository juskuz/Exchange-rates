require(gdata)        # biblioteka do wczytywania plik�w .xls
library(psych)        #z tej biblioteki korzystaj� funkcje geometric.mean i harmonic.mean
library(e1071)        #z tej biblioteki korzystaj� funkcje kurtosis i skewness
library(vioplot)      #z tej biblioteki korzysta funkcja vioplot

poprawne_daty = FALSE

print("Witaj w programie 'Kursy walut'.")
print("Mo�esz tu wy�wietli� dane statystyczne dla wybranej waluty w wybranym okresie czasu.")

while(poprawne_daty==FALSE) # dop�ki u�ytkownik nie wprowadzi poprawnych dat
{
  data1 <- as.Date(readline("Podaj dat� pocz�tkow� (RRRR-MM-DD): "))
  data2 <- as.Date(readline("Podaj dat� ko�cow� (RRRR-MM-DD): "))

  
  #data1 <- as.Date('1995-01-01') # ustawiane na sta�e
  #data2 <- as.Date('1996-06-08') # ustawiane na sta�e
  
  if (data1<as.Date('1995-01-01') || data1>as.Date('2016-06-08') || data2<as.Date('1995-01-01') || data2>as.Date('2016-06-08'))
  {
    print("B��d zakresu. Podaj poprawne dane. (od 1995-01-01 do 2016-06-08")
  }
  else #sprawd� czy data pocz�tkowa jest wcze�niej ni� ko�cowa
  {
    if (data1>data2)
    {
      print("Data ko�cowa nie mo�e by� wcze�niejsza ni� data pocz�tkowa.Wprowad� poprawne daty.")
    }
    else
    {
      poprawne_daty=T # przerwij p�tl� wczytywania
    }
  }
}

rok1 <- as.numeric(format(data1,'%Y')) # wyodr�bnij rok1 z daty1
rok2 <- as.numeric(format(data2,'%Y')) # wyodr�bnij rok2 z daty2


# do u�ywania offline
sciezka <- "/studia/rok 2/semestr 2/RPiS/Statystyka lab/poprawiona wersja programu 2" #�eby dzia�a�o bez internetu
setwd(sciezka)


# pobierz dane --> tryb online --> linijka 56 (zakomentuj/odkomentuj)
file_names=c()
years=c()
print("Trwa pobieranie danych... Prosz� czekaj.")

for (rok in rok1:rok2)
{
  dest=paste("D:/studia/rok 2/semestr 2/RPiS/Statystyka lab/poprawiona wersja programu 2/archiwum_tab_a_",rok,".xls", sep='')
  file_names <- c(file_names, dest)
  years <- c(years, rok)
  url=paste("http://www.nbp.pl/kursy/Archiwum/archiwum_tab_a_",rok,".xls", sep='')
  download.file(url, dest, "internal",quiet = FALSE, mode = "wb") #on/off downloading from internet
}

print("Pobrano pliki dla wybranego zakresu dat.")
print("Wczytuj� wszystkie dost�pne waluty w podanym zakresie dat. To mo�e zaj�� chwil�. Prosz� czekaj...")

dostepne_waluty_all <-c()
for (year in years)
{
  dest=paste("D:/studia/rok 2/semestr 2/RPiS/Statystyka lab/poprawiona wersja programu 2/archiwum_tab_a_",year,".xls", sep='')
  
  #wczytaj nag��wek (aby wyci�gna� nazwy walut)
  if(year<2000)
  {
    header <- read.xls(dest,
                       sheet=1,
                       header=FALSE,
                       perl="C:/strawberry/perl/bin/perl.exe",
                       skip=1,
                       nrow=1)
  }
  if (year>1999)
  {
    header <- read.xls(dest,
                       sheet=1,
                       header=FALSE,
                       perl="C:/strawberry/perl/bin/perl.exe",
                       nrow=1)
  }
  
  for (k in (header))
  {
    #print(toString(k))
    wal=gsub("^\\d*? ","", toString(k))
    if(wal=="10000PTE")
    {
      wal="PTE"
    }
    if (wal=="Data / date" || wal=="Data / Date" || wal== "NA" || wal =="NR TAB" || wal=="KURS ŚREDNI" 
        || wal =="Nr / No." || wal=="data" || wal =="z dnia" || wal=="Nr tabeli" || wal=="Data" 
        || wal =="Pełny numer tabeli" || wal=="nr tabeli" || wal=="pełny numer tabeli")
    {
      next
    }
    dostepne_waluty_all<-c(dostepne_waluty_all, wal)
  }
}


print("DOSTEPNE WALUTY:")
print (unique(dostepne_waluty_all))

waluta <- toupper(readline("Podaj walut�, dla kt�rej chcesz wy�wietli� dane:  ")) #konwertuj do wielkich liter
#waluta <- 'LUF'

currency_exist=is.element(waluta, dostepne_waluty_all) #sprawdz czy waluta wyst�puje w plikach

while (currency_exist==FALSE)
{
  print("Taka waluta w og�le nie istnieje w podanym zakresie dat. Mo�esz wybiera� spo�r�d nast�puj�cych: ")
  print(unique(dostepne_waluty_all))
  waluta <- toupper(readline("Podaj symbol waluty z powy�szych, aby wy�wietli� dane dla tej waluty:  "))
  currency_exist=is.element(waluta, dostepne_waluty_all) #sprawdz czy waluta wyst�puje w plikach
}

#*******************************************************************************************
# wczytaj dane z plik�w:

brak_waluty <-c()
kurs<-c()

for (year in rok1:rok2)
{
  #wczytuj kolejne pliki
  dest=paste("D:/studia/rok 2/semestr 2/RPiS/Statystyka lab/poprawiona wersja programu 2/archiwum_tab_a_",year,".xls", sep='')
  #wczytaj ca�� tabel� z pliku dla danego roku
  pomin=0
  if (year<2000)
  {
     pomin=1
  }
  
  df <- read.xls(dest,
                 sheet=1,
                 #header=TRUE,
                 perl="C:/strawberry/perl/bin/perl.exe",
                 skip=pomin,
                 blank.lines.skip=TRUE) #pomija puste linie
                 

  if(identical((grep(waluta, names(df), value=TRUE)), character(0)))  # w ktorych latach w zadanym przedziale dana waluta nie wystepuje
  {
    brak_waluty <- c(brak_waluty, year)
  }
  

  for(col in names(df)){
      if (!(identical((grep(waluta, col, value=TRUE)), character(0)))) # je�li znaleziono interesuj�c� kolumn�, to
      {
        for(val in df[[col]]){  # Iteracja przez wiersze kolumny interesujacej nas waluty
          # dosta� si� do interesuj�cej warto�ci
          if (!is.na(as.numeric(val)))
          {
            if (year!=1995)
            {
              val<-as.double(val)
            } else
            {
              if(as.double(val)>10)
              {
                val<-(as.double(val)/100) #podziel pierwsze warto�ci przez 100 aby wyr�wna� kursy
              }else
              {
                val<-as.double(val)  
              }
            }
           kurs<-c(kurs, val) #dodanie do zmiennej kurs
          }
        }
      }
  }
}

options(warn=-1) #do wy��czenia ostrze�e�

if (is.null(brak_waluty)==FALSE)
{
# informacja o braku danych
print(paste("W wybranym okresie czasu dla ", waluta, " nie istniej� dane w latach: ", sep=''))
print(brak_waluty)
} else
  {
  print(paste("W wybranym okresie czasu tzn. od ", data1, "do", data2, "istniej� dane dla waluty", waluta, sep=' '))
  }

print(paste("Wielko�� pr�by:", length(kurs)))
print(paste("Max(kurs):", max(kurs)))
print(paste("Min(kurs):", min(kurs)))
print(paste("�rednia arytmetyczna:", mean(kurs)))
print(paste("�rednia geometryczna:", geometric.mean(kurs)))
print(paste("�rednia harmoniczna:", harmonic.mean(kurs)))
print(paste("Kwantyl rz�du 1/4:", quantile(kurs, (1/4))))
print(paste("Kwantyl rz�du 3/4:", quantile(kurs, (3/4))))
print(paste("Mediana:", median(kurs)))
print("Przedzia� zmienno�ci pr�by (min i max):")
print(range(kurs))
print(paste("Wariancja:", var(kurs)))
print(paste("Odchylenie standardowe:", sd(kurs)))
#print(paste("Dominanta:", (mlv(kurs, method = "mfv"))))
print(paste("Rozst�p mi�dzykwartylowy:", IQR(kurs)))
print(paste("Kurtoza:", kurtosis(kurs)))
print(paste("Sko�no��: ", skewness(kurs)))
print("Wykresy zapisano do plik�w *.pdf.")
pdf(paste0("histogram_",rok1,"-",rok2,"_", waluta, ".pdf"), width=10, height=10)     #zapis histogramu do pliku .pdf
hist(kurs, main = paste("Histogram kurs�w ", waluta, sep=""),
     xlab = "Warto�� kursu w PLN", ylab = "Czestosc wystepowania danej warto�ci") #histogram
dev.off()  #zamkni�cie grafiki
pdf(paste0("wykres_pude�kowy_",rok1,"-",rok2,"_", waluta,".pdf"), width=10, height=10)     #zapis wykresu pude�kowego do pliku .pdf
boxplot(kurs) #wykres pude�kowy
dev.off() 
pdf(paste0("wykres_skrzypcowy_",rok1,"-",rok2,"_", waluta,".pdf"), width=10, height=10)     #zapis wykresu skrzypcowego
vioplot(kurs) #wykres skrzypcowy
dev.off()
