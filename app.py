import sys
import sqlite3
import shutil
from pathlib import Path
from datetime import datetime, date
from calendar import monthrange

from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QPushButton, QDateEdit, QComboBox, QSpinBox, QTabWidget, QTableWidget,
    QTableWidgetItem, QMessageBox, QGroupBox, QHeaderView, QFileDialog, QFrame,
    QLineEdit
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
        self.setStyleSheet("""
            QWidget { background: #171a20; color: #edf0f4; font-size: 14px; }
            QMainWindow { background: #111318; }
            QTabWidget::pane { border: 1px solid #343a46; border-radius: 8px; top: -1px; }
            QTabBar::tab { background: #232833; padding: 12px 22px; margin-right: 4px;
                           border-top-left-radius: 7px; border-top-right-radius: 7px; }
            QTabBar::tab:selected { background: #9d252d; color: white; }
            QGroupBox { border: 1px solid #343a46; border-radius: 9px; margin-top: 12px;
                        padding: 14px; font-weight: bold; }
            QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 6px; }
            QLineEdit, QComboBox, QSpinBox, QDateEdit {
                background: #232833; border: 1px solid #454d5c; border-radius: 6px;
                padding: 8px; min-height: 22px;
            }
            QPushButton { background: #9d252d; border: 0; border-radius: 7px;
                          padding: 10px 18px; font-weight: bold; }
            QPushButton:hover { background: #bd303a; }
            QPushButton:disabled { background: #454a54; color: #aeb4bd; }
            QTableWidget { background: #1d222b; alternate-background-color: #242a35;
                           gridline-color: #3a414d; border: 1px solid #343a46;
                           border-radius: 7px; }
            QHeaderView::section { background: #303744; color: #f4f5f7; padding: 8px;
                                   border: 0; font-weight: bold; }
            QLabel#title { font-size: 24px; font-weight: bold; color: #f0c75e; }
            QLabel#muted { color: #aab2bf; }
            QLabel#metric { font-size: 22px; font-weight: bold; color: #f0c75e; }
            QFrame#metricCard { background: #232833; border: 1px solid #343a46;
                                border-radius: 9px; }
        """)
        self.build_ui()
        self.refresh_history()
        self.refresh_report()

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
        form.addRow("Brygada:", self.entry_brigade)
        form.addRow("Liczba interwencji:", self.entry_count)
        layout.addWidget(box)

        buttons = QHBoxLayout()
        save = QPushButton("ZAPISZ WPIS")
        save.clicked.connect(self.save_entry)
        clear = QPushButton("WYCZYŚĆ")
        clear.setStyleSheet("background:#48515f")
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
        layout.addWidget(self.history_table)
        row = QHBoxLayout()
        edit_btn = QPushButton("ZAŁADUJ ZAZNACZONY WPIS DO EDYCJI")
        edit_btn.clicked.connect(self.load_selected_entry)
        delete_btn = QPushButton("USUŃ WPIS")
        delete_btn.setStyleSheet("background:#48515f")
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
        backup.setStyleSheet("background:#48515f")
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

    def save_entry(self):
        d = fmt_date(self.entry_date.date())
        shift = int(self.entry_shift.currentText())
        brigade = self.entry_brigade.currentText()
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
        self.history_table.setRowCount(len(rows))
        for r, (d, shift, brigade, count) in enumerate(rows):
            vals = [datetime.strptime(d, "%Y-%m-%d").strftime("%d.%m.%Y"),
                    str(int(d[-2:])), str(shift), brigade, str(count)]
            for c, val in enumerate(vals):
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignCenter)
                self.history_table.setItem(r, c, item)
        self.history_rows = rows

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
