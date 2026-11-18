from pathlib import Path
p=Path('/mnt/data/ri_work/app.py')
s=p.read_text(encoding='utf-8')
s=s.replace('    QLineEdit\n)', '    QLineEdit, QCheckBox\n)')
s=s.replace('        self.refresh_history()\n        self.refresh_report()\n', '        self.refresh_history()\n        self.refresh_report()\n        self.refresh_brigade_calendar()\n',1)
s=s.replace('        self.tabs.addTab(self.make_report_tab(), "Raporty")\n', '        self.tabs.addTab(self.make_report_tab(), "Raporty")\n        self.tabs.addTab(self.make_brigade_calendar_tab(), "Kalendarz brygadowy")\n')
s=s.replace('        self.entry_brigade = QComboBox()\n        self.entry_brigade.addItems(BRIGADES)\n', '        self.entry_brigade = QComboBox()\n        self.entry_brigade.addItems(BRIGADES)\n        self.entry_date.dateChanged.connect(self.update_auto_brigade)\n        self.entry_shift.currentIndexChanged.connect(self.update_auto_brigade)\n')
s=s.replace('        self.history_table.setSelectionBehavior(QTableWidget.SelectRows)\n        layout.addWidget(self.history_table)\n', '        self.history_table.setSelectionBehavior(QTableWidget.SelectRows)\n        self.history_table.setEditTriggers(QTableWidget.NoEditTriggers)\n        self.history_table.itemChanged.connect(self.on_history_item_changed)\n        layout.addWidget(self.history_table)\n')
s=s.replace('        row = QHBoxLayout()\n        edit_btn = QPushButton("ZAŁADUJ ZAZNACZONY WPIS DO EDYCJI")\n', '        row = QHBoxLayout()\n        self.history_lock = QCheckBox("Blokada edycji tabeli")\n        self.history_lock.setChecked(True)\n        self.history_lock.toggled.connect(self.set_history_editable)\n        row.addWidget(self.history_lock)\n        edit_btn = QPushButton("ZAŁADUJ ZAZNACZONY WPIS DO EDYCJI")\n')
# add calendar tab before metric_card
needle='    def metric_card(self, caption, value):\n'
calendar='''    def make_brigade_calendar_tab(self):
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
        d = self.entry_date.date().toPython()
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

'''
s=s.replace(needle,calendar+needle)
# set refreshing flag in history refresh, populate rows with day string correct
s=s.replace('        self.history_table.setRowCount(len(rows))\n', '        self._history_refreshing = True\n        self.history_table.setRowCount(len(rows))\n')
s=s.replace('            for c, val in enumerate(vals):\n                item = QTableWidgetItem(val)\n                item.setTextAlignment(Qt.AlignCenter)\n                self.history_table.setItem(r, c, item)\n        self.history_rows = rows\n', '            for c, val in enumerate(vals):\n                item = QTableWidgetItem(val)\n                item.setTextAlignment(Qt.AlignCenter)\n                if c == 1:\n                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)\n                self.history_table.setItem(r, c, item)\n        self.history_rows = rows\n        self._history_refreshing = False\n')
# auto choose brigade before save and refresh calendar after save/delete
s=s.replace('        brigade = self.entry_brigade.currentText()\n        count = self.entry_count.value()\n', '        work_day = self.entry_date.date().toPython()\n        matches = [b for b in BRIGADES if self.brigade_shift_for_date(work_day, b) == shift]\n        brigade = matches[0] if matches else self.entry_brigade.currentText()\n        self.entry_brigade.setCurrentText(brigade)\n        count = self.entry_count.value()\n')
s=s.replace('        self.refresh_history()\n        self.refresh_report()\n\n    def clear_entry', '        self.refresh_history()\n        self.refresh_report()\n        self.refresh_brigade_calendar()\n\n    def clear_entry')
s=s.replace('            self.refresh_history()\n            self.refresh_report()\n\n    def report_data', '            self.refresh_history()\n            self.refresh_report()\n            self.refresh_brigade_calendar()\n\n    def report_data')
# report table read only
s=s.replace('        self.report_table.setAlternatingRowColors(True)\n', '        self.report_table.setAlternatingRowColors(True)\n        self.report_table.setEditTriggers(QTableWidget.NoEditTriggers)\n')
p.write_text(s,encoding='utf-8')
