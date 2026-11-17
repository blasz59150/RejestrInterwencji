import sys
import sqlite3
import shutil
import configparser
from pathlib import Path
from datetime import datetime, date, timedelta
from calendar import monthrange

from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QPushButton, QDateEdit, QComboBox, QSpinBox, QTabWidget, QTableWidget,
    QTableWidgetItem, QMessageBox, QGroupBox, QHeaderView, QFileDialog, QFrame,
    QLineEdit, QCheckBox
)

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    Workbook = None


APP_NAME = "Rejestr Interwencji"
BRIGADES = ["I", "II", "III", "IV"]
SHIFTS = [1, 2, 3]


def app_folder() -> Path:
    # Przy EXE dane leżą obok pliku wykonywalnego; przy uruchomieniu .py obok źródła.
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE = app_folder()
DATA_DIR = BASE / "dane"
BACKUP_DIR = BASE / "kopie_zapasowe"
REPORT_DIR = BASE / "raporty"
DB_PATH = DATA_DIR / "interwencje.db"
CONFIG_PATH = BASE / "konfiguracja.ini"


def ensure_storage():
    DATA_DIR.mkdir(exist_ok=True)
    BACKUP_DIR.mkdir(exist_ok=True)
    REPORT_DIR.mkdir(exist_ok=True)
    with sqlite3.connect(DB_PATH) as con:
        con.execute("""
            CREATE TABLE IF NOT EXISTS interventions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                work_date TEXT NOT NULL,
                shift INTEGER NOT NULL CHECK(shift BETWEEN 1 AND 3),
                brigade TEXT NOT NULL CHECK(brigade IN ('I','II','III','IV')),
                count INTEGER NOT NULL CHECK(count >= 0),
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(work_date, shift, brigade)
            )
        """)
        con.execute("CREATE INDEX IF NOT EXISTS idx_work_date ON interventions(work_date)")


def db_connect():
    return sqlite3.connect(DB_PATH)


def fmt_date(qdate):
    return qdate.toString("yyyy-MM-dd")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1120, 760)
        self.setMinimumSize(900, 620)
        self.theme = self.load_theme()
        self.schedule_start = self.load_schedule_date()
        self._loading_history = False
        self.editing_record = None
        self.build_ui()
        self.apply_theme()
        self.refresh_history()
        self.refresh_report()

    def load_theme(self):
        config = configparser.ConfigParser()
        if CONFIG_PATH.exists():
            try:
                config.read(CONFIG_PATH, encoding="utf-8")
                value = config.get("appearance", "theme", fallback="light").lower()
                return value if value in ("light", "dark") else "light"
            except (configparser.Error, OSError):
                pass
        return "light"

    def read_config(self):
        config = configparser.ConfigParser()
        if CONFIG_PATH.exists():
            try:
                config.read(CONFIG_PATH, encoding="utf-8")
            except configparser.Error:
                pass
        return config

    def save_config(self):
        config = self.read_config()
        if not config.has_section("appearance"):
            config.add_section("appearance")
        config.set("appearance", "theme", self.theme)
        if not config.has_section("schedule"):
            config.add_section("schedule")
        config.set("schedule", "brigade_i_start", self.schedule_start.toString("yyyy-MM-dd"))
        with CONFIG_PATH.open("w", encoding="utf-8") as f:
            config.write(f)

    def load_schedule_date(self):
        config = self.read_config()
        raw = config.get("schedule", "brigade_i_start", fallback="2026-11-13")
        qdate = QDate.fromString(raw, "yyyy-MM-dd")
        return qdate if qdate.isValid() else QDate(2026, 11, 13)

    def toggle_theme(self):
        self.theme = "dark" if self.theme == "light" else "light"
        self.save_config()
        self.apply_theme()

    def apply_theme(self):
        if self.theme == "light":
            c = {
                "window": "#f3efe6", "surface": "#fffdf8", "surface_alt": "#e9e4d9",
                "text": "#243247", "muted": "#64748b", "border": "#cbd3dc",
                "field": "#ffffff", "field_border": "#aebdce", "accent": "#315f8c",
                "accent_hover": "#244b73", "tab": "#e2e8ef", "table": "#fffdf8",
                "table_alt": "#f0f4f8", "header": "#d9e4ef", "metric": "#244f78",
                "secondary": "#64748b", "selection": "#c8dced"
            }
            self.theme_button.setText("🌙  Tryb ciemny")
        else:
            c = {
                "window": "#171a20", "surface": "#232833", "surface_alt": "#2b313d",
                "text": "#edf0f4", "muted": "#aab2bf", "border": "#343a46",
                "field": "#232833", "field_border": "#454d5c", "accent": "#9d252d",
                "accent_hover": "#bd303a", "tab": "#232833", "table": "#1d222b",
                "table_alt": "#242a35", "header": "#303744", "metric": "#f0c75e",
                "secondary": "#48515f", "selection": "#55343a"
            }
            self.theme_button.setText("☀  Tryb jasny")
        self.setStyleSheet(f"""
            QWidget {{ background: {c['window']}; color: {c['text']}; font-size: 14px; }}
            QMainWindow {{ background: {c['window']}; }}
            QTabWidget::pane {{ background: {c['surface']}; border: 1px solid {c['border']}; border-radius: 8px; top: -1px; }}
            QTabBar::tab {{ background: {c['tab']}; color: {c['text']}; padding: 12px 22px; margin-right: 4px; border-top-left-radius: 7px; border-top-right-radius: 7px; }}
            QTabBar::tab:selected {{ background: {c['accent']}; color: white; }}
            QGroupBox {{ background: {c['surface']}; border: 1px solid {c['border']}; border-radius: 9px; margin-top: 12px; padding: 14px; font-weight: bold; }}
            QGroupBox::title {{ subcontrol-origin: margin; left: 12px; padding: 0 6px; }}
            QLineEdit, QComboBox, QSpinBox, QDateEdit {{ background: {c['field']}; color: {c['text']}; border: 1px solid {c['field_border']}; border-radius: 6px; padding: 8px; min-height: 22px; }}
            QComboBox QAbstractItemView {{ background: {c['surface']}; color: {c['text']}; selection-background-color: {c['selection']}; }}
            QPushButton {{ background: {c['accent']}; color: white; border: 0; border-radius: 7px; padding: 10px 18px; font-weight: bold; }}
            QPushButton:hover {{ background: {c['accent_hover']}; }}
            QPushButton:disabled {{ background: {c['secondary']}; color: white; }}
            QPushButton#secondary {{ background: {c['secondary']}; color: white; }}
            QTableWidget {{ background: {c['table']}; alternate-background-color: {c['table_alt']}; color: {c['text']}; gridline-color: {c['border']}; border: 1px solid {c['border']}; border-radius: 7px; selection-background-color: {c['selection']}; selection-color: {c['text']}; }}
            QHeaderView::section {{ background: {c['header']}; color: {c['text']}; padding: 8px; border: 0; font-weight: bold; }}
            QLabel#title {{ font-size: 24px; font-weight: bold; color: {c['accent']}; }}
            QLabel#muted {{ color: {c['muted']}; }}
            QLabel#metric {{ font-size: 22px; font-weight: bold; color: {c['metric']}; }}
            QFrame#metricCard {{ background: {c['surface']}; border: 1px solid {c['border']}; border-radius: 9px; }}
        """)

    def build_ui(self):
        central = QWidget()
        outer = QVBoxLayout(central)
        header = QHBoxLayout()
        title = QLabel("REJESTR INTERWENCJI")
        title.setObjectName("title")
        subtitle = QLabel("Lokalna ewidencja • baza danych SQLite • praca offline")
        subtitle.setObjectName("muted")
        left = QVBoxLayout()
        left.addWidget(title)
        left.addWidget(subtitle)
        header.addLayout(left)
        header.addStretch()
        self.theme_button = QPushButton()
        self.theme_button.clicked.connect(self.toggle_theme)
        header.addWidget(self.theme_button)
        self.db_label = QLabel(f"Baza: {DB_PATH.name}")
        self.db_label.setObjectName("muted")
        header.addWidget(self.db_label)
        outer.addLayout(header)

        self.tabs = QTabWidget()
        self.tabs.addTab(self.make_entry_tab(), "Wprowadzanie")
        self.tabs.addTab(self.make_history_tab(), "Historia")
        self.tabs.addTab(self.make_report_tab(), "Raporty")
        outer.addWidget(self.tabs)
        self.setCentralWidget(central)

    def make_entry_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)
        box = QGroupBox("Nowy wpis / korekta wpisu")
        form = QFormLayout(box)
        self.entry_date = QDateEdit(QDate.currentDate())
        self.entry_date.setCalendarPopup(True)
        self.entry_date.setDisplayFormat("dd.MM.yyyy")
        self.entry_shift = QComboBox()
        self.entry_shift.addItems(["1", "2", "3"])
        self.entry_brigade = QComboBox()
        self.entry_brigade.addItems(BRIGADES)
        self.entry_count = QSpinBox()
        self.entry_count.setRange(0, 9999)
        self.entry_count.setValue(0)
        form.addRow("Data:", self.entry_date)
        form.addRow("Zmiana:", self.entry_shift)
        self.entry_brigade.setEnabled(False)
        self.brigade_hint = QLabel("Brygada zostanie przypisana automatycznie na podstawie grafiku.")
        self.brigade_hint.setObjectName("muted")
        form.addRow("Brygada:", self.entry_brigade)
        form.addRow("Liczba interwencji:", self.entry_count)
        layout.addWidget(box)
        self.entry_date.dateChanged.connect(self.update_brigade_for_entry)
        self.entry_shift.currentIndexChanged.connect(self.update_brigade_for_entry)

        schedule_box = QGroupBox("Ustawienia grafiku brygad")
        schedule_row = QHBoxLayout(schedule_box)
        schedule_row.addWidget(QLabel("Pierwszy dzień I zmiany brygady I:"))
        self.schedule_date_edit = QDateEdit(self.schedule_start)
        self.schedule_date_edit.setCalendarPopup(True)
        self.schedule_date_edit.setDisplayFormat("dd.MM.yyyy")
        save_schedule = QPushButton("ZAPISZ DATĘ BAZOWĄ")
        save_schedule.setObjectName("secondary")
        save_schedule.clicked.connect(self.save_schedule_settings)
        schedule_row.addWidget(self.schedule_date_edit)
        schedule_row.addWidget(save_schedule)
        schedule_row.addStretch()
        layout.addWidget(schedule_box)

        buttons = QHBoxLayout()
        save = QPushButton("ZAPISZ WPIS")
        save.clicked.connect(self.save_entry)
        clear = QPushButton("WYCZYŚĆ")
        clear.setObjectName("secondary")
        clear.clicked.connect(self.clear_entry)
        buttons.addWidget(save)
        buttons.addWidget(clear)
        buttons.addStretch()
        layout.addLayout(buttons)

        info = QLabel(
            "Wskazówka: dla tej samej daty, zmiany i brygady kolejny zapis aktualizuje istniejący wpis. "
            "Możesz więc poprawić pomyłkę bez tworzenia duplikatu."
        )
        info.setWordWrap(True)
        info.setObjectName("muted")
        layout.addWidget(info)
        layout.addWidget(self.brigade_hint)
        layout.addStretch()
        self.update_brigade_for_entry()
        return w

    def make_history_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)
        controls = QHBoxLayout()
        self.hist_month = QComboBox()
        self.hist_year = QComboBox()
        self.fill_month_year(self.hist_month, self.hist_year)
        refresh = QPushButton("ODŚWIEŻ")
        refresh.clicked.connect(self.refresh_history)
        controls.addWidget(QLabel("Miesiąc:"))
        controls.addWidget(self.hist_month)
        controls.addWidget(QLabel("Rok:"))
        controls.addWidget(self.hist_year)
        controls.addWidget(refresh)
        controls.addStretch()
        layout.addLayout(controls)
        self.history_table = QTableWidget(0, 5)
        self.history_table.setHorizontalHeaderLabels(["Data", "Dzień", "Zmiana", "Brygada", "Interwencje"])
        self.history_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.history_table.setAlternatingRowColors(True)
        self.history_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.history_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.history_table.itemChanged.connect(self.on_history_item_changed)
        layout.addWidget(self.history_table)
        row = QHBoxLayout()
        edit_btn = QPushButton("ZAŁADUJ ZAZNACZONY WPIS DO EDYCJI")
        edit_btn.clicked.connect(self.load_selected_entry)
        delete_btn = QPushButton("USUŃ WPIS")
        delete_btn.setObjectName("secondary")
        delete_btn.clicked.connect(self.delete_selected_entry)
        row.addWidget(edit_btn)
        row.addWidget(delete_btn)
        row.addStretch()
        layout.addLayout(row)
        self.history_edit_lock = QCheckBox("Blokada edycji tabeli")
        self.history_edit_lock.setChecked(True)
        self.history_edit_lock.toggled.connect(self.set_history_editable)
        layout.addWidget(self.history_edit_lock, alignment=Qt.AlignLeft)
        return w

    def make_report_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)
        controls = QHBoxLayout()
        self.rep_month = QComboBox()
        self.rep_year = QComboBox()
        self.fill_month_year(self.rep_month, self.rep_year)
        gen = QPushButton("GENERUJ RAPORT")
        gen.clicked.connect(self.refresh_report)
        export = QPushButton("EKSPORT DO EXCELA")
        export.clicked.connect(self.export_excel)
        backup = QPushButton("KOPIA ZAPASOWA")
        backup.setObjectName("secondary")
        backup.clicked.connect(self.backup_database)
        controls.addWidget(QLabel("Miesiąc:"))
        controls.addWidget(self.rep_month)
        controls.addWidget(QLabel("Rok:"))
        controls.addWidget(self.rep_year)
        controls.addWidget(gen)
        controls.addWidget(export)
        controls.addWidget(backup)
        controls.addStretch()
        layout.addLayout(controls)

        cards = QHBoxLayout()
        self.metric_total = self.metric_card("ŁĄCZNIE", "0")
        self.metric_day = self.metric_card("ŚREDNIA / DZIEŃ Z WPISAMI", "0,00")
        self.metric_shift = self.metric_card("ŚREDNIA / ZMIANĘ Z WPISAMI", "0,00")
        self.metric_days = self.metric_card("DNI Z WPISAMI", "0")
        for card in [self.metric_total, self.metric_day, self.metric_shift, self.metric_days]:
            cards.addWidget(card)
        layout.addLayout(cards)

        self.report_table = QTableWidget(0, 5)
        self.report_table.setHorizontalHeaderLabels(["Dzień", "I zmiana", "II zmiana", "III zmiana", "Suma dnia"])
        self.report_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.report_table.setAlternatingRowColors(True)
        self.report_table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.report_table, 1)
        self.brigade_summary = QLabel("Podsumowanie brygad: —")
        self.brigade_summary.setWordWrap(True)
        self.brigade_summary.setObjectName("muted")
        layout.addWidget(self.brigade_summary)
        return w

    def metric_card(self, caption, value):
        frame = QFrame()
        frame.setObjectName("metricCard")
        lay = QVBoxLayout(frame)
        cap = QLabel(caption)
        cap.setObjectName("muted")
        val = QLabel(value)
        val.setObjectName("metric")
        lay.addWidget(cap)
        lay.addWidget(val)
        frame.value_label = val
        return frame

    def fill_month_year(self, month_combo, year_combo):
        month_combo.clear()
        months = ["Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec",
                  "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień"]
        month_combo.addItems(months)
        month_combo.setCurrentIndex(date.today().month - 1)
        year_combo.clear()
        years = [str(y) for y in range(date.today().year - 5, date.today().year + 3)]
        year_combo.addItems(years)
        year_combo.setCurrentText(str(date.today().year))

    def selected_period(self, month_combo, year_combo):
        return month_combo.currentIndex() + 1, int(year_combo.currentText())

    def schedule_shift_for(self, brigade: str, work_day: date):
        offsets = {"I": 0, "II": 4, "III": 8, "IV": 12}
        base = self.schedule_start.toPython()
        cycle_day = (work_day - (base + timedelta(days=offsets[brigade]))).days % 16
        if 0 <= cycle_day <= 3:
            return 1
        if 5 <= cycle_day <= 8:
            return 3
        if 11 <= cycle_day <= 14:
            return 2
        return None

    def brigade_for(self, work_day: date, shift: int):
        matches = [b for b in BRIGADES if self.schedule_shift_for(b, work_day) == shift]
        return matches[0] if len(matches) == 1 else None

    def update_brigade_for_entry(self, *_):
        if not hasattr(self, "entry_brigade"):
            return
        work_day = self.entry_date.date().toPython()
        shift = int(self.entry_shift.currentText())
        brigade = self.brigade_for(work_day, shift)
        self.entry_brigade.clear()
        if brigade:
            self.entry_brigade.addItem(brigade)
            self.brigade_hint.setText(f"Automatycznie przypisana brygada: {brigade}")
        else:
            self.entry_brigade.addItem("Brak brygady")
            self.brigade_hint.setText("Uwaga: grafik nie wskazuje jednej brygady dla tej zmiany.")

    def save_schedule_settings(self):
        self.schedule_start = self.schedule_date_edit.date()
        self.save_config()
        self.update_brigade_for_entry()
        QMessageBox.information(self, "Grafik zapisany",
            "Zapisano datę bazową. Brygady II, III i IV są przesunięte odpowiednio o 4, 8 i 12 dni.\n"
            "Istniejące wpisy historyczne nie zostały zmienione.")

    def set_history_editable(self, unlocked):
        if unlocked:
            self.history_table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
        else:
            self.history_table.setEditTriggers(QTableWidget.NoEditTriggers)

    def on_history_item_changed(self, item):
        if self._loading_history or not self.history_edit_lock or self.history_edit_lock.isChecked():
            return
        row = item.row()
        if row < 0 or row >= len(self.history_rows):
            return
        record_id, old_date, old_shift, old_brigade, old_count = self.history_rows[row]
        try:
            if item.column() == 0:
                parsed = datetime.strptime(item.text().strip(), "%d.%m.%Y").date()
                new_date = parsed.isoformat()
            else:
                new_date = old_date
            new_shift = int(self.history_table.item(row, 2).text())
            new_brigade = self.history_table.item(row, 3).text().strip().upper()
            new_count = int(self.history_table.item(row, 4).text())
            if new_shift not in SHIFTS or new_brigade not in BRIGADES or new_count < 0:
                raise ValueError
            # Dzień tygodnia w tabeli jest wyliczany z daty i nie jest edytowalny.
            with db_connect() as con:
                con.execute("""UPDATE interventions SET work_date=?, shift=?, brigade=?, count=?, updated_at=? WHERE id=?""",
                    (new_date, new_shift, new_brigade, new_count, datetime.now().isoformat(timespec="seconds"), record_id))
        except sqlite3.IntegrityError:
            QMessageBox.warning(self, "Duplikat", "Taki wpis (data, zmiana, brygada) już istnieje. Cofam zmianę.")
            self.refresh_history()
            return
        except (ValueError, TypeError):
            QMessageBox.warning(self, "Nieprawidłowe dane",
                "Sprawdź datę (DD.MM.RRRR), zmianę (1–3), brygadę (I–IV) i nieujemną liczbę interwencji.")
            self.refresh_history()
            return
        self.refresh_history()
        self.refresh_report()

    def save_entry(self):
        d = fmt_date(self.entry_date.date())
        shift = int(self.entry_shift.currentText())
        brigade = None
        if self.editing_record and d == self.editing_record[1] and shift == self.editing_record[2]:
            # Przy korekcie samej liczby zachowujemy brygadę zapisaną historycznie,
            # nawet jeśli data bazowa grafiku została później zmieniona.
            brigade = self.editing_record[3]
        else:
            brigade = self.brigade_for(self.entry_date.date().toPython(), shift)
        if not brigade:
            QMessageBox.warning(self, "Brak brygady", "Dla wybranej daty i zmiany nie znaleziono jednej brygady w grafiku.")
            return
        count = self.entry_count.value()
        now = datetime.now().isoformat(timespec="seconds")
        with db_connect() as con:
            if self.editing_record:
                con.execute("UPDATE interventions SET work_date=?, shift=?, brigade=?, count=?, updated_at=? WHERE id=?",
                    (d, shift, brigade, count, now, self.editing_record[0]))
            else:
                con.execute("""
                    INSERT INTO interventions(work_date, shift, brigade, count, created_at, updated_at)
                    VALUES(?,?,?,?,?,?)
                    ON CONFLICT(work_date, shift, brigade)
                    DO UPDATE SET count=excluded.count, updated_at=excluded.updated_at
                """, (d, shift, brigade, count, now, now))
        self.editing_record = None
        QMessageBox.information(self, "Zapisano", "Wpis został zapisany w lokalnej bazie danych.")
        self.refresh_history()
        self.refresh_report()

    def clear_entry(self):
        self.editing_record = None
        self.entry_date.setDate(QDate.currentDate())
        self.entry_shift.setCurrentIndex(0)
        self.update_brigade_for_entry()
        self.entry_count.setValue(0)

    def refresh_history(self):
        if not hasattr(self, "history_table"):
            return
        month, year = self.selected_period(self.hist_month, self.hist_year)
        start = f"{year:04d}-{month:02d}-01"
        end = f"{year:04d}-{month:02d}-{monthrange(year, month)[1]:02d}"
        with db_connect() as con:
            rows = con.execute("""
                SELECT id, work_date, shift, brigade, count FROM interventions
                WHERE work_date BETWEEN ? AND ?
                ORDER BY work_date, shift, brigade
            """, (start, end)).fetchall()
        self._loading_history = True
        self.history_table.setRowCount(len(rows))
        for r, (record_id, d, shift, brigade, count) in enumerate(rows):
            vals = [datetime.strptime(d, "%Y-%m-%d").strftime("%d.%m.%Y"),
                    datetime.strptime(d, "%Y-%m-%d").strftime("%d"), str(shift), brigade, str(count)]
            for c, val in enumerate(vals):
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignCenter)
                if c == 1:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.history_table.setItem(r, c, item)
        self.history_rows = rows
        self._loading_history = False
        self.set_history_editable(not self.history_edit_lock.isChecked())

    def load_selected_entry(self):
        row = self.history_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Brak wyboru", "Najpierw zaznacz wpis na liście.")
            return
        record_id, d, shift, brigade, count = self.history_rows[row]
        qd = QDate.fromString(d, "yyyy-MM-dd")
        self.entry_date.setDate(qd)
        self.entry_shift.setCurrentText(str(shift))
        self.entry_brigade.setCurrentText(brigade)
        self.entry_count.setValue(count)
        self.editing_record = (record_id, d, shift, brigade, count)
        self.tabs.setCurrentIndex(0)

    def delete_selected_entry(self):
        row = self.history_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Brak wyboru", "Najpierw zaznacz wpis na liście.")
            return
        record_id, d, shift, brigade, count = self.history_rows[row]
        answer = QMessageBox.question(self, "Potwierdź usunięcie",
            f"Usunąć wpis: {d}, zmiana {shift}, brygada {brigade}, interwencje {count}?")
        if answer == QMessageBox.Yes:
            with db_connect() as con:
                con.execute("DELETE FROM interventions WHERE id=?", (record_id,))
            self.refresh_history()
            self.refresh_report()

    def report_data(self):
        month, year = self.selected_period(self.rep_month, self.rep_year)
        days = monthrange(year, month)[1]
        start = f"{year:04d}-{month:02d}-01"
        end = f"{year:04d}-{month:02d}-{days:02d}"
        with db_connect() as con:
            rows = con.execute("""
                SELECT work_date, shift, brigade, count FROM interventions
                WHERE work_date BETWEEN ? AND ?
                ORDER BY work_date, shift, brigade
            """, (start, end)).fetchall()
        by_day_shift = {(day, sh): 0 for day in range(1, days + 1) for sh in SHIFTS}
        brigade_totals = {b: 0 for b in BRIGADES}
        total = 0
        active_days = set()
        active_shifts = set()
        for d, sh, b, cnt in rows:
            day = int(d[-2:])
            by_day_shift[(day, sh)] += cnt
            brigade_totals[b] += cnt
            total += cnt
            active_days.add(day)
            active_shifts.add((day, sh))
        return month, year, days, by_day_shift, brigade_totals, total, len(active_days), len(active_shifts), rows

    def refresh_report(self):
        if not hasattr(self, "report_table"):
            return
        month, year, days, data, brigades, total, active_days, active_shifts, rows = self.report_data()
        self.report_table.setRowCount(days)
        for day in range(1, days + 1):
            vals = [str(day)] + [str(data[(day, sh)]) for sh in SHIFTS] + [
                str(sum(data[(day, sh)] for sh in SHIFTS))]
            for col, val in enumerate(vals):
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignCenter)
                self.report_table.setItem(day - 1, col, item)
        avg_day = total / active_days if active_days else 0
        avg_shift = total / active_shifts if active_shifts else 0
        self.metric_total.value_label.setText(str(total))
        self.metric_day.value_label.setText(f"{avg_day:.2f}".replace(".", ","))
        self.metric_shift.value_label.setText(f"{avg_shift:.2f}".replace(".", ","))
        self.metric_days.value_label.setText(str(active_days))
        self.brigade_summary.setText("Podsumowanie brygad: " + "   |   ".join(
            f"{b}: {brigades[b]}" for b in BRIGADES))

    def export_excel(self):
        if Workbook is None:
            QMessageBox.critical(self, "Brak biblioteki", "Brakuje openpyxl. Zainstaluj zależności z requirements.txt.")
            return
        month, year, days, data, brigades, total, active_days, active_shifts, rows = self.report_data()
        default = REPORT_DIR / f"Raport_interwencji_{year}_{month:02d}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Zapisz raport Excel", str(default), "Excel (*.xlsx)")
        if not path:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "Raport miesięczny"
        ws.append([f"RAPORT INTERWENCJI — {self.rep_month.currentText()} {year}"])
        ws.append(["Dzień", "I zmiana", "II zmiana", "III zmiana", "Suma dnia"])
        for day in range(1, days + 1):
            ws.append([day] + [data[(day, sh)] for sh in SHIFTS] +
                      [sum(data[(day, sh)] for sh in SHIFTS)])
        ws.append([])
        ws.append(["Łączna liczba interwencji", total])
        ws.append(["Średnia na dzień z wpisami", round(total / active_days, 2) if active_days else 0])
        ws.append(["Średnia na zmianę z wpisami", round(total / active_shifts, 2) if active_shifts else 0])
        ws.append(["Dni z wpisami", active_days])
        ws.append([])
        ws.append(["Brygada", "Liczba interwencji"])
        for b in BRIGADES:
            ws.append([b, brigades[b]])
        detail = wb.create_sheet("Szczegóły")
        detail.append(["Data", "Dzień", "Zmiana", "Brygada", "Interwencje"])
        for d, sh, b, cnt in rows:
            detail.append([d, int(d[-2:]), sh, b, cnt])
        for sheet in wb.worksheets:
            sheet.freeze_panes = "A3"
            sheet.auto_filter.ref = sheet.dimensions
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF", size=14)
                cell.fill = PatternFill("solid", fgColor="9D252D")
            for cell in sheet[2]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="303744")
            for col in sheet.columns:
                letter = get_column_letter(col[0].column)
                width = max(12, min(34, max(len(str(c.value or "")) for c in col) + 3))
                sheet.column_dimensions[letter].width = width
            for row_cells in sheet.iter_rows():
                for cell in row_cells:
                    cell.alignment = Alignment(vertical="center")
        wb.save(path)
        QMessageBox.information(self, "Gotowe", f"Raport zapisano:\n{path}")

    def backup_database(self):
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target = BACKUP_DIR / f"interwencje_{stamp}.db"
        shutil.copy2(DB_PATH, target)
        QMessageBox.information(self, "Kopia zapasowa", f"Utworzono kopię:\n{target}")


def main():
    ensure_storage()
    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 10))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
