APP_STYLESHEET = """
QWidget {
    background-color: #F5F7FA;
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 13px;
    color: #1F2937;
}

QFrame#Sidebar {
    background-color: #17365D;
}

QPushButton#NavButton {
    background-color: transparent;
    color: #D9E2F1;
    text-align: left;
    padding: 12px 18px;
    border: none;
    font-size: 14px;
}
QPushButton#NavButton:hover {
    background-color: #1F4E78;
}
QPushButton#NavButton:checked {
    background-color: #1F4E78;
    color: white;
    font-weight: 600;
    border-left: 4px solid #4FC3F7;
}

QLabel#AppTitle {
    color: white;
    font-size: 16px;
    font-weight: 700;
    padding: 18px;
}

QFrame#Card {
    background-color: white;
    border: 1px solid #E2E8F0;
    border-radius: 10px;
}

QLabel#CardTitle {
    font-size: 14px;
    font-weight: 700;
    color: #17365D;
}

QLabel#PageTitle {
    font-size: 20px;
    font-weight: 700;
    color: #17365D;
    padding: 4px 0 12px 0;
}

QPushButton#PrimaryButton {
    background-color: #1F4E78;
    color: white;
    border: none;
    border-radius: 6px;
    padding: 10px 22px;
    font-weight: 600;
    font-size: 13px;
}
QPushButton#PrimaryButton:hover {
    background-color: #17365D;
}
QPushButton#PrimaryButton:disabled {
    background-color: #A0AEC0;
}

QPushButton#SecondaryButton {
    background-color: white;
    color: #1F4E78;
    border: 1px solid #1F4E78;
    border-radius: 6px;
    padding: 9px 20px;
    font-weight: 600;
}
QPushButton#SecondaryButton:hover {
    background-color: #EAF2FB;
}

QProgressBar {
    border: 1px solid #E2E8F0;
    border-radius: 6px;
    background-color: #EDF2F7;
    text-align: center;
    height: 18px;
}
QProgressBar::chunk {
    background-color: #1F4E78;
    border-radius: 6px;
}

QTableWidget {
    background-color: white;
    border: 1px solid #E2E8F0;
    gridline-color: #EDF2F7;
}
QHeaderView::section {
    background-color: #17365D;
    color: white;
    padding: 6px;
    border: none;
    font-weight: 600;
}

QLabel#StatValue {
    font-size: 26px;
    font-weight: 800;
    color: #17365D;
}
QLabel#StatLabel {
    font-size: 12px;
    color: #64748B;
}

QLabel#Pill {
    border-radius: 10px;
    padding: 3px 10px;
    font-weight: 700;
    font-size: 11px;
}
"""

SEVERITY_COLORS = {
    "HIGH": "#F4CCCC",
    "MEDIUM": "#FFF2CC",
    "LOW": "#D9EAD3",
}
