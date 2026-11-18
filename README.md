# RejestrInterwencji

Lokalna, przenośna aplikacja Windows do ewidencji interwencji.

## Pobranie gotowej aplikacji

Po każdym pushu do gałęzi `main` GitHub Actions buduje aplikację dla Windows.
Gotowy plik ZIP znajdziesz w repozytorium: **Actions → najnowszy workflow → Artifacts**.

## Uruchomienie

1. Pobierz i rozpakuj `RejestrInterwencji-Windows-Portable.zip`.
2. Uruchom `RejestrInterwencji.exe`.
3. Nie jest wymagany Python, serwer SQL ani dostęp do internetu.

Baza SQLite tworzy się automatycznie w folderze `dane` obok aplikacji.
Przenoś i archiwizuj cały folder aplikacji, nie sam plik EXE.

## Funkcje wersji startowej

- wpis: data, zmiana (1–3), brygada (I–IV), liczba interwencji;
- korekta istniejącego wpisu dla tej samej daty/zmiany/brygady;
- historia miesięczna, edycja i usuwanie;
- raport miesięczny z sumami dziennymi, zmianowymi i brygadowymi;
- eksport raportu do XLSX;
- ręczna kopia zapasowa bazy.

## Budowanie lokalne

Na Windows z Pythonem 3.10+ można uruchomić `build_windows.bat`.
Automatyczne budowanie na GitHubie nie wymaga Pythona na komputerze użytkownika.

## Motywy wyglądu

Aplikacja ma przełączane motywy jasny (beżowo-niebieski) i ciemny (grafitowo-czerwony).
Wybrany motyw jest zapisywany w pliku `konfiguracja.ini` obok EXE i przywracany przy kolejnym uruchomieniu.
