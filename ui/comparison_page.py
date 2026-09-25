import os
import subprocess
import sys
from datetime import datetime

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame, QPushButton,
    QFileDialog, QProgressBar, QGridLayout, QMessageBox, QTableWidget,
    QTableWidgetItem, QDialog, QHeaderView
)

from engine import comparator_engine as engine
from data import history_store


class CompareWorker(QThread):
    finished_ok = Signal(dict)
    finished_error = Signal(str)
    progress = Signal(int, str)

    def __init__(self, old_file, new_file):
        super().__init__()
        self.old_file = old_file
        self.new_file = new_file

    def run(self):
        try:
            self.progress.emit(10, "Hashing files (SHA-256)...")
            self.progress.emit(30, "Parsing baseline export...")
            self.progress.emit(50, "Parsing candidate export...")
            self.progress.emit(70, "Comparing objects...")
            result = engine.run_comparison(self.old_file, self.new_file)
            self.progress.emit(95, "Writing Excel report...")
            self.progress.emit(100, "Done")
            self.finished_ok.emit(result)
        except Exception as e:
            self.finished_error.emit(str(e))


class FileCard(QFrame):
    def __init__(self, title):
        super().__init__()
        self.setObjectName("Card")
        self.file_path = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(8)

        self.title_label = QLabel(title)
        self.title_label.setObjectName("CardTitle")
        layout.addWidget(self.title_label)

        self.browse_btn = QPushButton("Browse XML File...")
        self.browse_btn.setObjectName("SecondaryButton")
        self.browse_btn.clicked.connect(self.browse)
        layout.addWidget(self.browse_btn)

        self.name_label = QLabel("No file selected")
        self.name_label.setWordWrap(True)
        self.name_label.setStyleSheet("color:#64748B;")
        layout.addWidget(self.name_label)

        self.meta_label = QLabel("")
        self.meta_label.setWordWrap(True)
        self.meta_label.setStyleSheet("color:#94A3B8; font-size:11px;")
        layout.addWidget(self.meta_label)

    def browse(self):
        path, _ = QFileDialog.getOpenFileName(
            self, f"Select {self.title_label.text()}", "",
            "XML Files (*.xml);;All Files (*.*)"
        )
        if path:
            self.set_file(path)

    def set_file(self, path):
        self.file_path = path
        self.name_label.setText(os.path.basename(path))
        self.name_label.setStyleSheet("color:#17365D; font-weight:600;")
        try:
            size_kb = os.path.getsize(path) / 1024
            self.meta_label.setText(f"{size_kb:.1f} KB   •   {path}")
        except OSError:
            self.meta_label.setText(path)

    def reset(self):
        self.file_path = None
        self.name_label.setText("No file selected")
        self.name_label.setStyleSheet("color:#64748B;")
        self.meta_label.setText("")


class ChangeRegisterDialog(QDialog):
    """Popup table showing every detected change (mirrors sheet 02)."""

    def __init__(self, changes, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Full Change Register")
        self.resize(1000, 550)

        layout = QVBoxLayout(self)
        table = QTableWidget()
        headers = ["Severity", "Category", "Class", "Logical Name (LN)",
                   "Field", "Baseline Value", "Candidate Value"]
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setRowCount(len(changes))
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setEditTriggers(QTableWidget.NoEditTriggers)

        for row, change in enumerate(changes):
            values = [change["Severity"].split(" - ")[0], change["Category"],
                      change["Class"], change["LN"], change["Field"],
                      change["OldValue"], change["NewValue"]]
            for col, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if col == 0:
                    from ui.styles import SEVERITY_COLORS
                    from PySide6.QtGui import QColor
                    color = SEVERITY_COLORS.get(value, "#FFFFFF")
                    item.setBackground(QColor(color))
                table.setItem(row, col, item)

        layout.addWidget(table)
        close_btn = QPushButton("Close")
        close_btn.setObjectName("SecondaryButton")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn, alignment=Qt.AlignRight)


class ComparisonPage(QWidget):
    comparison_finished = Signal()

    def __init__(self):
        super().__init__()
        self.result = None
        self.worker = None

        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 24)
        root.setSpacing(16)

        title = QLabel("FOTA Firmware Comparison")
        title.setObjectName("PageTitle")
        root.addWidget(title)

        # File cards row
        cards_row = QHBoxLayout()
        cards_row.setSpacing(16)
        self.baseline_card = FileCard("BASELINE / OLD")
        self.candidate_card = FileCard("CANDIDATE / NEW")
        cards_row.addWidget(self.baseline_card)
        cards_row.addWidget(self.candidate_card)
        root.addLayout(cards_row)

        # Compare button + progress
        action_row = QHBoxLayout()
        self.compare_btn = QPushButton("COMPARE FILES")
        self.compare_btn.setObjectName("PrimaryButton")
        self.compare_btn.clicked.connect(self.run_compare)
        action_row.addWidget(self.compare_btn)
        action_row.addStretch()
        root.addLayout(action_row)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        root.addWidget(self.progress_bar)

        self.progress_label = QLabel("")
        self.progress_label.setStyleSheet("color:#64748B;")
        root.addWidget(self.progress_label)

        # Results panel
        self.results_frame = QFrame()
        self.results_frame.setObjectName("Card")
        self.results_frame.setVisible(False)
        results_layout = QVBoxLayout(self.results_frame)
        results_layout.setContentsMargins(20, 18, 20, 18)
        results_layout.setSpacing(12)

        self.result_title = QLabel("Comparison Result")
        self.result_title.setObjectName("CardTitle")
        results_layout.addWidget(self.result_title)

        self.stats_grid = QGridLayout()
        self.stats_grid.setSpacing(18)
        results_layout.addLayout(self.stats_grid)

        self.category_label = QLabel("")
        self.category_label.setWordWrap(True)
        results_layout.addWidget(self.category_label)

        buttons_row = QHBoxLayout()
        self.view_btn = QPushButton("View Changes")
        self.view_btn.setObjectName("SecondaryButton")
        self.view_btn.clicked.connect(self.view_changes)

        self.export_btn = QPushButton("Export Excel Copy")
        self.export_btn.setObjectName("SecondaryButton")
        self.export_btn.clicked.connect(self.export_copy)

        self.open_btn = QPushButton("Open Report")
        self.open_btn.setObjectName("SecondaryButton")
        self.open_btn.clicked.connect(self.open_report)

        self.new_btn = QPushButton("New Compare")
        self.new_btn.setObjectName("PrimaryButton")
        self.new_btn.clicked.connect(self.reset)

        for b in (self.view_btn, self.export_btn, self.open_btn, self.new_btn):
            buttons_row.addWidget(b)
        buttons_row.addStretch()
        results_layout.addLayout(buttons_row)

        root.addWidget(self.results_frame)
        root.addStretch()

    # --------------------------------------------------------
    def run_compare(self):
        old_file = self.baseline_card.file_path
        new_file = self.candidate_card.file_path

        if not old_file or not new_file:
            QMessageBox.warning(self, "Missing Files",
                                 "Please select both the baseline and candidate XML files.")
            return

        self.compare_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.results_frame.setVisible(False)

        self.worker = CompareWorker(old_file, new_file)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished_ok.connect(self.on_success)
        self.worker.finished_error.connect(self.on_error)
        self.worker.start()

    def on_progress(self, value, message):
        self.progress_bar.setValue(value)
        self.progress_label.setText(message)

    def on_success(self, result):
        self.result = result
        self.compare_btn.setEnabled(True)
        history_store.add_entry(result, status="SUCCESS")
        self.render_results(result)
        self.comparison_finished.emit()

    def on_error(self, message):
        self.compare_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        history_store.add_failure(
            self.baseline_card.file_path, self.candidate_card.file_path,
            message, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        QMessageBox.critical(self, "Comparison Failed", message)
        self.comparison_finished.emit()

    # --------------------------------------------------------
    def render_results(self, result):
        self.results_frame.setVisible(True)

        # clear stats grid
        while self.stats_grid.count():
            item = self.stats_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        stats = [
            ("Objects Compared", result["comparison"]["common_objects"]),
            ("Total Changes", len(result["changes"])),
            ("HIGH", result["high"]),
            ("MEDIUM", result["medium"]),
            ("LOW", result["low"]),
        ]
        for col, (label, value) in enumerate(stats):
            box = QVBoxLayout()
            val_label = QLabel(str(value))
            val_label.setObjectName("StatValue")
            lbl_label = QLabel(label)
            lbl_label.setObjectName("StatLabel")
            box.addWidget(val_label, alignment=Qt.AlignCenter)
            box.addWidget(lbl_label, alignment=Qt.AlignCenter)
            container = QWidget()
            container.setLayout(box)
            self.stats_grid.addWidget(container, 0, col)

        cat_lines = [f"{cat}: {count}" for cat, count in
                     sorted(result["category_counts"].items(), key=lambda x: -x[1])]
        self.category_label.setText("Change Analysis  —  " + "   |   ".join(cat_lines)
                                     if cat_lines else "No differences detected.")

    # --------------------------------------------------------
    def view_changes(self):
        if not self.result:
            return
        dialog = ChangeRegisterDialog(self.result["changes"], self)
        dialog.exec()

    def export_copy(self):
        if not self.result:
            return
        dest, _ = QFileDialog.getSaveFileName(
            self, "Save Excel Report Copy As",
            os.path.basename(self.result["output_path"]), "Excel Files (*.xlsx)")
        if dest:
            import shutil
            shutil.copyfile(self.result["output_path"], dest)
            QMessageBox.information(self, "Saved", f"Report copy saved to:\n{dest}")

    def open_report(self):
        if not self.result:
            return
        path = self.result["output_path"]
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.call(["open", path])
            else:
                subprocess.call(["xdg-open", path])
        except Exception as e:
            QMessageBox.warning(self, "Could not open file", str(e))

    def reset(self):
        self.baseline_card.reset()
        self.candidate_card.reset()
        self.results_frame.setVisible(False)
        self.progress_bar.setVisible(False)
        self.progress_label.setText("")
        self.result = None
