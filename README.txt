REJESTR INTERWENCJI — wersja startowa

Wymagania:
- Windows 10/11
- Python 3.10 lub nowszy (tylko do uruchomienia wersji źródłowej / zbudowania EXE)

URUCHOMIENIE WERSJI ŹRÓDŁOWEJ:
1. Otwórz terminal w tym katalogu.
2. Wykonaj: py -m venv .venv
3. Wykonaj: .venv\\Scripts\\activate
4. Wykonaj: pip install -r requirements.txt
5. Wykonaj: python app.py

BUDOWANIE EXE NA WINDOWS:
1. Zainstaluj Python 3.10+.
2. Uruchom build_windows.bat.
3. Gotowa aplikacja pojawi się w dist\\RejestrInterwencji\\.
4. Przenieś cały folder RejestrInterwencji na komputer docelowy.
   Nie przenoś samego EXE — razem z nim musi być folder dane, kopie_zapasowe i raporty.

BAZA:
- dane\\interwencje.db — lokalna baza SQLite.
- Nie wymaga instalowania SQL Servera ani żadnego serwera.
- Baza tworzona automatycznie przy pierwszym uruchomieniu.

ZASADY WPISÓW:
- Jeden wpis jest identyfikowany przez datę + zmianę + brygadę.
- Ponowny zapis tej samej kombinacji aktualizuje poprzednią wartość.
- Wartość 0 jest dozwolona.
- Raport pokazuje sumę interwencji na każdej zmianie (sumuje brygady), sumę dnia,
  średnią na dzień z wpisami, średnią na zmianę z wpisami i podsumowanie brygad.

UWAGA:
To jest pierwsza wersja robocza do testów funkcjonalnych. EXE należy zbudować na Windows
za pomocą dołączonego pliku BAT; środowisko przygotowujące paczkę nie kompiluje natywnie
aplikacji Windows.
