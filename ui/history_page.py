import os
import subprocess
import sys

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTableWidget,
    QTableWidgetItem, QHeaderView, QMessageBox
)

from data import history_store


class HistoryPage(QWidget):
    def __init__(self):
        super().__init__()
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 24)
        root.setSpacing(16)

        header_row = QHBoxLayout()
        title = QLabel("Comparison History")
        title.setObjectName("PageTitle")
        header_row.addWidget(title)
        header_row.addStretch()
        refresh_btn = QPushButton("Refresh")
        refresh_btn.setObjectName("SecondaryButton")
        refresh_btn.clicked.connect(self.refresh)
        header_row.addWidget(refresh_btn)
        root.addLayout(header_row)

        self.table = QTableWidget()
        headers = ["Timestamp", "Meter", "Baseline", "Candidate",
                   "Total", "HIGH", "MEDIUM", "LOW", "Status", "Report"]
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.cellDoubleClicked.connect(self.open_selected_report)
        root.addWidget(self.table)

        hint = QLabel("Double-click a row to open its Excel report.")
        hint.setStyleSheet("color:#94A3B8; font-size:11px;")
        root.addWidget(hint)

        self._entries = []
        self.refresh()

    def refresh(self):
        self._entries = history_store.all_entries()
        self.table.setRowCount(len(self._entries))
        for row, e in enumerate(self._entries):
            values = [
                e["timestamp"], e["meter_number"] or "-",
                os.path.basename(e["old_file"] or ""), os.path.basename(e["new_file"] or ""),
                str(e["total_changes"]), str(e["high_changes"]),
                str(e["medium_changes"]), str(e["low_changes"]),
                e["status"], os.path.basename(e["report_path"] or "-"),
            ]
            for col, v in enumerate(values):
                self.table.setItem(row, col, QTableWidgetItem(v))

    def open_selected_report(self, row, _col):
        if row >= len(self._entries):
            return
        path = self._entries[row]["report_path"]
        if not path or not os.path.exists(path):
            QMessageBox.warning(self, "Report Not Found",
                                 "The Excel report for this entry could not be located.")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.call(["open", path])
            else:
                subprocess.call(["xdg-open", path])
        except Exception as e:
            QMessageBox.warning(self, "Could not open file", str(e))
