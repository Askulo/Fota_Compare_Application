from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout, QFrame, QLabel,
    QPushButton, QStackedWidget, QButtonGroup
)

from ui.dashboard import Dashboard
from ui.comparison_page import ComparisonPage
from ui.history_page import HistoryPage
from ui.styles import APP_STYLESHEET


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("FOTA Firmware Comparator — Smart Meter HES")
        self.resize(1180, 760)
        self.setStyleSheet(APP_STYLESHEET)

        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QHBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # ---------------- Sidebar ----------------
        sidebar = QFrame()
        sidebar.setObjectName("Sidebar")
        sidebar.setFixedWidth(210)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_layout.setSpacing(0)

        app_title = QLabel("⚡ FOTA\nCOMPARATOR")
        app_title.setObjectName("AppTitle")
        sidebar_layout.addWidget(app_title)

        self.nav_group = QButtonGroup(self)
        self.nav_group.setExclusive(True)

        self.dashboard_page = Dashboard()
        self.comparison_page = ComparisonPage()
        self.history_page = HistoryPage()

        self.stack = QStackedWidget()
        self.stack.addWidget(self.dashboard_page)   # index 0
        self.stack.addWidget(self.comparison_page)  # index 1
        self.stack.addWidget(self.history_page)     # index 2

        nav_items = [
            ("Dashboard", 0),
            ("Compare", 1),
            ("History", 2),
        ]
        for label, index in nav_items:
            btn = QPushButton(label)
            btn.setObjectName("NavButton")
            btn.setCheckable(True)
            btn.clicked.connect(lambda checked, i=index: self.switch_page(i))
            self.nav_group.addButton(btn)
            sidebar_layout.addWidget(btn)

        self.nav_group.buttons()[1].setChecked(True)  # default to Compare
        sidebar_layout.addStretch()

        version_label = QLabel("v3.0")
        version_label.setStyleSheet("color:#7FA6C9; padding:12px 18px;")
        sidebar_layout.addWidget(version_label)

        root_layout.addWidget(sidebar)
        root_layout.addWidget(self.stack, stretch=1)

        self.stack.setCurrentIndex(1)

        # keep dashboard/history fresh whenever a comparison finishes
        self.comparison_page.comparison_finished.connect(self.dashboard_page.refresh)
        self.comparison_page.comparison_finished.connect(self.history_page.refresh)

    def switch_page(self, index):
        self.stack.setCurrentIndex(index)
        if index == 0:
            self.dashboard_page.refresh()
        elif index == 2:
            self.history_page.refresh()
