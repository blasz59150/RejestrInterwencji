import sys
import sqlite3
import shutil
import configparser
from pathlib import Path
from datetime import datetime, date
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
        self.build_ui()
        self.apply_theme()
        self.refresh_history()
        self.refresh_report()
        self.refresh_brigade_calendar()

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

    def save_theme(self):
        config = configparser.ConfigParser()
        config["appearance"] = {"theme": self.theme}
        with CONFIG_PATH.open("w", encoding="utf-8") as f:
            config.write(f)

    def toggle_theme(self):
        self.theme = "dark" if self.theme == "light" else "light"
        self.save_theme()
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
        self.tabs.addTab(self.make_brigade_calendar_tab(), "Kalendarz brygadowy")
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
        self.entry_date.dateChanged.connect(self.update_auto_brigade)
        self.entry_shift.currentIndexChanged.connect(self.update_auto_brigade)
        self.entry_count = QSpinBox()
        self.entry_count.setRange(0, 9999)
        self.entry_count.setValue(0)
        form.addRow("Data:", self.entry_date)
        form.addRow("Zmiana:", self.entry_shift)
        form.addRow("Brygada:", self.entry_brigade)
        form.addRow("Liczba interwencji:", self.entry_count)
        layout.addWidget(box)

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
        layout.addStretch()
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
        self.history_lock = QCheckBox("Blokada edycji tabeli")
        self.history_lock.setChecked(True)
        self.history_lock.toggled.connect(self.set_history_editable)
        row.addWidget(self.history_lock)
        edit_btn = QPushButton("ZAŁADUJ ZAZNACZONY WPIS DO EDYCJI")
        edit_btn.clicked.connect(self.load_selected_entry)
        delete_btn = QPushButton("USUŃ WPIS")
        delete_btn.setObjectName("secondary")
        delete_btn.clicked.connect(self.delete_selected_entry)
        row.addWidget(edit_btn)
        row.addWidget(delete_btn)
        row.addStretch()
        layout.addLayout(row)
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

    def make_brigade_calendar_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)
        controls = QHBoxLayout()
        self.cal_prev = QPushButton("◀ Poprzedni")
        self.cal_prev.setObjectName("secondary")
        self.cal_next = QPushButton("Następny ▶")
        self.cal_next.setObjectName("secondary")
        self.cal_today = QPushButton("Dzisiaj")
        self.cal_month = QComboBox()
        self.cal_year = QComboBox()
        self.fill_month_year(self.cal_month, self.cal_year)
        self.cal_prev.clicked.connect(lambda: self.move_calendar_month(-1))
        self.cal_next.clicked.connect(lambda: self.move_calendar_month(1))
        self.cal_today.clicked.connect(self.calendar_today)
        self.cal_month.currentIndexChanged.connect(self.refresh_brigade_calendar)
        self.cal_year.currentIndexChanged.connect(self.refresh_brigade_calendar)
        controls.addWidget(self.cal_prev)
        controls.addWidget(self.cal_month)
        controls.addWidget(self.cal_year)
        controls.addWidget(self.cal_next)
        controls.addWidget(self.cal_today)
        controls.addStretch()
        layout.addLayout(controls)
        self.calendar_info = QLabel("Harmonogram czterech brygad • cykl 16-dniowy")
        self.calendar_info.setObjectName("muted")
        layout.addWidget(self.calendar_info)
        self.brigade_calendar_table = QTableWidget(0, 6)
        self.brigade_calendar_table.setHorizontalHeaderLabels(["Dzień", "Tydzień", "I", "II", "III", "IV"])
        self.brigade_calendar_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.brigade_calendar_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.brigade_calendar_table.setAlternatingRowColors(True)
        self.brigade_calendar_table.verticalHeader().setVisible(False)
        layout.addWidget(self.brigade_calendar_table, 1)
        self.calendar_legend = QLabel("Liczba w komórce = zarejestrowane interwencje brygady danego dnia. „–” = wolne.")
        self.calendar_legend.setWordWrap(True)
        self.calendar_legend.setObjectName("muted")
        layout.addWidget(self.calendar_legend)
        return w

    def brigade_shift_for_date(self, work_day, brigade):
        # Daty bazowe oznaczają pierwszy dzień cyklu (zmiana 1) danej brygady.
        base_dates = {"I": date(2026, 11, 13), "II": date(2026, 11, 17),
                      "III": date(2026, 11, 21), "IV": date(2026, 11, 25)}
        phase = (work_day - base_dates[brigade]).days % 16
        if 0 <= phase <= 3:
            return 1
        if 5 <= phase <= 8:
            return 3
        if 11 <= phase <= 14:
            return 2
        return None

    def update_auto_brigade(self, *_):
        if not hasattr(self, "entry_brigade"):
            return
        d = date(self.entry_date.date().year(), self.entry_date.date().month(), self.entry_date.date().day())
        shift = int(self.entry_shift.currentText())
        matches = [b for b in BRIGADES if self.brigade_shift_for_date(d, b) == shift]
        if matches:
            self.entry_brigade.setCurrentText(matches[0])

    def move_calendar_month(self, delta):
        month, year = self.selected_period(self.cal_month, self.cal_year)
        index = year * 12 + (month - 1) + delta
        new_year, new_month0 = divmod(index, 12)
        self.cal_year.setCurrentText(str(new_year))
        self.cal_month.setCurrentIndex(new_month0)
        self.refresh_brigade_calendar()

    def calendar_today(self):
        today = date.today()
        self.cal_year.setCurrentText(str(today.year))
        self.cal_month.setCurrentIndex(today.month - 1)
        self.refresh_brigade_calendar()

    def refresh_brigade_calendar(self, *_):
        if not hasattr(self, "brigade_calendar_table"):
            return
        month, year = self.selected_period(self.cal_month, self.cal_year)
        days = monthrange(year, month)[1]
        start, end = f"{year:04d}-{month:02d}-01", f"{year:04d}-{month:02d}-{days:02d}"
        with db_connect() as con:
            rows = con.execute("SELECT work_date, brigade, SUM(count) FROM interventions WHERE work_date BETWEEN ? AND ? GROUP BY work_date, brigade", (start, end)).fetchall()
        counts = {(d, b): n for d, b, n in rows}
        self.brigade_calendar_table.setRowCount(days)
        weekdays = ["Pn", "Wt", "Śr", "Cz", "Pt", "So", "N"]
        for day_num in range(1, days + 1):
            d = date(year, month, day_num)
            vals = [str(day_num), weekdays[d.weekday()]]
            for brigade in BRIGADES:
                shift = self.brigade_shift_for_date(d, brigade)
                if shift is None:
                    vals.append("–")
                else:
                    n = counts.get((d.isoformat(), brigade), 0)
                    vals.append(f"{shift}  ·  {n}" if n else str(shift))
            for col, value in enumerate(vals):
                item = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignCenter)
                if col < 2 and d.weekday() == 6:
                    item.setForeground(Qt.red)
                if d == date.today():
                    item.setBackground(Qt.green)
                self.brigade_calendar_table.setItem(day_num - 1, col, item)
        self.calendar_info.setText(f"Harmonogram brygad — {self.cal_month.currentText()} {year}. W komórkach: numer zmiany; po kropce liczba interwencji.")

    def set_history_editable(self, unlocked):
        self.history_table.setEditTriggers(QTableWidget.AllEditTriggers if unlocked else QTableWidget.NoEditTriggers)
        # Checkbox oznacza blokadę, dlatego odznaczenie odblokowuje tabelę.

    def on_history_item_changed(self, item):
        if self.history_lock.isChecked() or getattr(self, "_history_refreshing", False):
            return
        row = item.row()
        if row >= len(getattr(self, "history_rows", [])):
            return
        old = self.history_rows[row]
        try:
            d = old[0]
            shift, brigade, count = old[1], old[2], old[3]
            if item.column() == 0:
                parsed = datetime.strptime(item.text().strip(), "%d.%m.%Y").date()
                d = parsed.isoformat()
            elif item.column() == 2:
                shift = int(item.text())
                if shift not in SHIFTS: raise ValueError("Zmiana musi wynosić 1, 2 albo 3.")
            elif item.column() == 3:
                brigade = item.text().strip().upper()
                if brigade not in BRIGADES: raise ValueError("Brygada musi być I, II, III albo IV.")
            elif item.column() == 4:
                count = int(item.text())
                if count < 0: raise ValueError("Liczba interwencji nie może być ujemna.")
            elif item.column() == 1:
                raise ValueError("Dzień tygodnia jest wyliczany automatycznie z daty.")
            with db_connect() as con:
                con.execute("UPDATE interventions SET work_date=?, shift=?, brigade=?, count=?, updated_at=? WHERE work_date=? AND shift=? AND brigade=?",
                            (d, shift, brigade, count, datetime.now().isoformat(timespec="seconds"), *old[:3]))
            self.refresh_history(); self.refresh_report(); self.refresh_brigade_calendar()
        except (ValueError, sqlite3.IntegrityError) as exc:
            QMessageBox.warning(self, "Nieprawidłowa zmiana", str(exc) or "Taki wpis już istnieje.")
            self.refresh_history()

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

    def save_entry(self):
        d = fmt_date(self.entry_date.date())
        shift = int(self.entry_shift.currentText())
        work_day = date(self.entry_date.date().year(), self.entry_date.date().month(), self.entry_date.date().day())
        matches = [b for b in BRIGADES if self.brigade_shift_for_date(work_day, b) == shift]
        brigade = matches[0] if matches else self.entry_brigade.currentText()
        self.entry_brigade.setCurrentText(brigade)
        count = self.entry_count.value()
        now = datetime.now().isoformat(timespec="seconds")
        with db_connect() as con:
            con.execute("""
                INSERT INTO interventions(work_date, shift, brigade, count, created_at, updated_at)
                VALUES(?,?,?,?,?,?)
                ON CONFLICT(work_date, shift, brigade)
                DO UPDATE SET count=excluded.count, updated_at=excluded.updated_at
            """, (d, shift, brigade, count, now, now))
        QMessageBox.information(self, "Zapisano", "Wpis został zapisany w lokalnej bazie danych.")
        self.refresh_history()
        self.refresh_report()
        self.refresh_brigade_calendar()

    def clear_entry(self):
        self.entry_date.setDate(QDate.currentDate())
        self.entry_shift.setCurrentIndex(0)
        self.entry_brigade.setCurrentIndex(0)
        self.entry_count.setValue(0)

    def refresh_history(self):
        if not hasattr(self, "history_table"):
            return
        month, year = self.selected_period(self.hist_month, self.hist_year)
        start = f"{year:04d}-{month:02d}-01"
        end = f"{year:04d}-{month:02d}-{monthrange(year, month)[1]:02d}"
        with db_connect() as con:
            rows = con.execute("""
                SELECT work_date, shift, brigade, count FROM interventions
                WHERE work_date BETWEEN ? AND ?
                ORDER BY work_date, shift, brigade
            """, (start, end)).fetchall()
        self._history_refreshing = True
        self.history_table.setRowCount(len(rows))
        for r, (d, shift, brigade, count) in enumerate(rows):
            vals = [datetime.strptime(d, "%Y-%m-%d").strftime("%d.%m.%Y"),
                    str(int(d[-2:])), str(shift), brigade, str(count)]
            for c, val in enumerate(vals):
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignCenter)
                if c == 1:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.history_table.setItem(r, c, item)
        self.history_rows = rows
        self._history_refreshing = False

    def load_selected_entry(self):
        row = self.history_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Brak wyboru", "Najpierw zaznacz wpis na liście.")
            return
        d, shift, brigade, count = self.history_rows[row]
        qd = QDate.fromString(d, "yyyy-MM-dd")
        self.entry_date.setDate(qd)
        self.entry_shift.setCurrentText(str(shift))
        self.entry_brigade.setCurrentText(brigade)
        self.entry_count.setValue(count)
        self.tabs.setCurrentIndex(0)

    def delete_selected_entry(self):
        row = self.history_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Brak wyboru", "Najpierw zaznacz wpis na liście.")
            return
        d, shift, brigade, count = self.history_rows[row]
        answer = QMessageBox.question(self, "Potwierdź usunięcie",
            f"Usunąć wpis: {d}, zmiana {shift}, brygada {brigade}, interwencje {count}?")
        if answer == QMessageBox.Yes:
            with db_connect() as con:
                con.execute("DELETE FROM interventions WHERE work_date=? AND shift=? AND brigade=?",
                            (d, shift, brigade))
            self.refresh_history()
            self.refresh_report()
            self.refresh_brigade_calendar()

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
