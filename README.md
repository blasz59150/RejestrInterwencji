# Rejestr Interwencji — v3.1

Portable Windows application (PySide6 + SQLite), offline.

## Co zawiera ta aktualizacja
- zakładkę „Kalendarz brygadowy” w tabelarycznym układzie miesięcznym;
- cykl 16-dniowy dla brygad I–IV;
- automatyczne przypisywanie brygady według daty i zmiany;
- liczbę zarejestrowanych interwencji w komórkach kalendarza;
- blokadę edycji historii domyślnie włączoną — zaznaczona blokada wyłącza edycję, odznaczenie ją włącza;
- kontrolę źródeł przed budowaniem EXE w GitHub Actions.

## Aktualizacja repozytorium
Wgraj/zastąp pliki z paczki w repozytorium, zachowując foldery danych użytkownika:
- `dane/` — baza SQLite;
- `kopie_zapasowe/` — kopie bazy;
- `raporty/` — wygenerowane raporty;
- `konfiguracja.ini` — ustawiony motyw, jeśli już istnieje.

Nie usuwaj tych folderów ani nie zastępuj istniejącej bazy pustą bazą z paczki.

## GitHub Actions
Po commitowaniu zmian do gałęzi `main` uruchomi się workflow:
1. **Pre-build Check** — kontrola składni Python i obecności kluczowych funkcji;
2. **Build portable EXE** — budowanie uruchomi się wyłącznie, jeśli kontrola przejdzie;
3. ZIP pojawi się w `Actions` → najnowsze uruchomienie → `Artifacts`.

Można też uruchomić workflow ręcznie przez `Actions` → workflow → `Run workflow`.

## Ważne
Kontrola Pre-build jest kontrolą statyczną źródeł. Nie zastępuje pełnego testu ręcznego aplikacji w Windows.
Po pobraniu ZIP rozpakuj go do osobnego folderu testowego. Dopiero po sprawdzeniu działania zastąp starą wersję programu. Zachowaj dotychczasowe `dane/` i `konfiguracja.ini`.
