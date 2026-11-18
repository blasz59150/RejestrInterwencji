"""Static pre-build checks for Rejestr Interwencji."""
from pathlib import Path
import ast
import sys

ROOT = Path(__file__).resolve().parents[1]
APP = ROOT / "app.py"
REQUIRED = [
    ROOT / "requirements.txt",
    ROOT / ".github" / "workflows" / "build-windows.yml",
]

def fail(message: str) -> None:
    print(f"[BŁĄD] {message}")
    raise SystemExit(1)

for path in REQUIRED:
    if not path.is_file():
        fail(f"Brakuje wymaganego pliku: {path.relative_to(ROOT)}")
if not APP.is_file():
    fail("Brakuje app.py")

source = APP.read_text(encoding="utf-8")
try:
    tree = ast.parse(source, filename=str(APP))
except SyntaxError as exc:
    fail(f"Błąd składni app.py, linia {exc.lineno}: {exc.msg}")

checks = {
    "zakładka kalendarza brygadowego": '"Kalendarz brygadowy"' in source,
    "generator widoku kalendarza": "def make_brigade_calendar_tab" in source,
    "wyliczanie zmiany w cyklu": "def brigade_shift_for_date" in source,
    "automatyczne przypisanie brygady": "def update_auto_brigade" in source,
    "zapis wpisu": "def save_entry" in source,
    "historia z checkboxem blokady": 'QCheckBox("Blokada edycji tabeli")' in source,
    "obsługa edycji historii": "def on_history_item_changed" in source,
    "poprawna logika blokady historii": (
        "QTableWidget.NoEditTriggers if locked else QTableWidget.AllEditTriggers" in source
    ),
    "zapis ustawień motywu": "def save_theme" in source and "konfiguracja.ini" in source,
}
for label, passed in checks.items():
    print(f"[{'OK' if passed else 'BŁĄD'}] {label}")
    if not passed:
        raise SystemExit(1)

# Dodatkowo sprawdź, czy w module istnieje punkt wejścia aplikacji.
if not any(isinstance(node, ast.FunctionDef) and node.name == "main" for node in tree.body):
    fail("Nie znaleziono funkcji main()")

print("\nPRE-BUILD CHECK: OK — źródła przeszły kontrolę.")
