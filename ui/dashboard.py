from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame, QTableWidget,
    QTableWidgetItem, QHeaderView
)

from data import history_store


class Dashboard(QWidget):
    def __init__(self):
        super().__init__()
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 24)
        root.setSpacing(16)

        title = QLabel("Dashboard")
        title.setObjectName("PageTitle")
        root.addWidget(title)

        self.stats_row = QHBoxLayout()
        self.stats_row.setSpacing(16)
        root.addLayout(self.stats_row)

        recent_label = QLabel("Recent Activity")
        recent_label.setObjectName("CardTitle")
        root.addWidget(recent_label)

        self.table = QTableWidget()
        headers = ["Timestamp", "Meter", "Total Changes", "HIGH", "Status"]
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        root.addWidget(self.table)

        self.refresh()

    def _stat_box(self, value, label):
        frame = QFrame()
        frame.setObjectName("Card")
        layout = QVBoxLayout(frame)
        val = QLabel(str(value))
        val.setObjectName("StatValue")
        lbl = QLabel(label)
        lbl.setObjectName("StatLabel")
        layout.addWidget(val, alignment=Qt.AlignCenter)
        layout.addWidget(lbl, alignment=Qt.AlignCenter)
        return frame

    def refresh(self):
        while self.stats_row.count():
            item = self.stats_row.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        s = history_store.stats()
        for value, label in [(s["total"], "Total Comparisons"),
                              (s["success"], "Successful"),
                              (s["failed"], "Failed")]:
            self.stats_row.addWidget(self._stat_box(value, label))
        self.stats_row.addStretch()

        entries = history_store.all_entries()[:15]
        self.table.setRowCount(len(entries))
        for row, e in enumerate(entries):
            values = [e["timestamp"], e["meter_number"] or "-",
                      str(e["total_changes"]), str(e["high_changes"]), e["status"]]
            for col, v in enumerate(values):
                self.table.setItem(row, col, QTableWidgetItem(v))
