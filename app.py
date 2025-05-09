# -*- encoding: utf-8 -*-
# @File        :   marking.py
# @Time        :   2025/05/06 01:31:18
# @Author      :   Siyou
# @Description :

import sys
import os
import zipfile
import tempfile
import shutil
import csv
import traceback
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QHBoxLayout, QPushButton, QLabel, QLineEdit,
    QTextEdit, QFileDialog, QMessageBox, QInputDialog,
    QDialog, QListWidget, QDialogButtonBox, QMenuBar, QAction
)
from PyQt5.QtGui import QIcon
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import Qt, QUrl, pyqtSlot, QPointF
from nbconvert import HTMLExporter
import nbformat
from nbformat.reader import NotJSONError
# Matplotlib for histogram
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
# ac

# Application metadata
APP_NAME = "QMPlusScoring"
APP_VERSION = "1.0.0"
APP_AUTHOR = "Siyou Li"


class HistoryDialog(QDialog):
    def __init__(self, parent, evaluations):
        super().__init__(parent)
        self.setWindowTitle("Select Previous Evaluation")
        self.setModal(True)
        layout = QVBoxLayout(self)
        self.list_widget = QListWidget()
        self.list_widget.addItems(evaluations)
        self.list_widget.itemDoubleClicked.connect(self._on_select)
        layout.addWidget(self.list_widget)
        buttons = QDialogButtonBox(QDialogButtonBox.Cancel)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _on_select(self, item):
        self.selected = item.text()
        self.accept()

class QMPlusScoring(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1800, 800)

        # History storage
        self.history = tuple()

        # Prompt startup inputs
        self._prompt_initial_inputs()
        self.students = []
        self.current_index = 0
        self.records = []

        # Build UI
        self._setup_ui()

        # Load data
        self._load_submissions(self.submissions_zip)
        self._load_reference(self.reference_nb)
        if self.students:
            self._load_current()

    def _prompt_initial_inputs(self):
        # Use non-native dialog on macOS and attach to main window
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog

        # Submissions ZIP
        submissions, _ = QFileDialog.getOpenFileName(
            self,
            "Please select the .zip file containing all student submissions",
            "",
            "Zip Files (*.zip)",
            options=options
        )
        if not submissions:
            QMessageBox.critical(self, "Error", "Submissions ZIP is required.")
            sys.exit(1)
        self.submissions_zip = submissions

        # Reference notebook
        reference, _ = QFileDialog.getOpenFileName(
            self,
            "Please select the reference .ipynb notebook for marking",
            "",
            "Notebooks (*.ipynb)",
            options=options
        )
        if not reference:
            QMessageBox.critical(self, "Error", "Reference notebook is required.")
            sys.exit(1)
        self.reference_nb = reference

        # Lab number
        lab, ok = QInputDialog.getText(
            self,
            "Experiment Number",
            "Enter the lab/experiment identifier (e.g., 'lab7' or 'lab 7'):"
        )
        if not ok or not lab.strip():
            QMessageBox.critical(self, "Error", "Lab number is required.")
            sys.exit(1)
        self.lab_key = lab.strip()

    def _setup_ui(self):
        # Menu bar with Software Details
        menubar = QMenuBar(self)
        help_menu = menubar.addMenu("Help")
        details_action = QAction("Software Details", self)
        details_action.triggered.connect(self._show_software_details)
        help_menu.addAction(details_action)
        self.setMenuBar(menubar)

        # Root container
        root = QWidget()
        root_layout = QVBoxLayout(root)

        # Top bar: show files + lab
        top_bar = QWidget()
        top_bar.setStyleSheet("QLabel { color: black; font-weight: bold; }")
        top_layout = QHBoxLayout(top_bar)
        top_layout.setContentsMargins(5, 5, 5, 5)
        top_layout.setSpacing(10)
        top_layout.addWidget(QLabel("Submissions:"))
        self.submissions_field = QLineEdit(self.submissions_zip)
        self.submissions_field.setReadOnly(True)
        top_layout.addWidget(self.submissions_field)
        top_layout.addWidget(QLabel("Reference:"))
        self.reference_field = QLineEdit(self.reference_nb)
        self.reference_field.setReadOnly(True)
        top_layout.addWidget(self.reference_field)
        top_layout.addWidget(QLabel("Lab:"))
        self.lab_field = QLineEdit(self.lab_key)
        self.lab_field.setReadOnly(True)
        top_layout.addWidget(self.lab_field)
        root_layout.addWidget(top_bar)

        # Main layout
        main_layout = QHBoxLayout()
        # Notebook viewers
        viewers = QWidget()
        view_layout = QHBoxLayout(viewers)
        self.student_view = QWebEngineView()
        self.ref_view = QWebEngineView()
        view_layout.addWidget(self.student_view)
        view_layout.addWidget(self.ref_view)
        main_layout.addWidget(viewers)

        # Side panel
        side = QWidget()
        side.setFixedWidth(300)
        side_layout = QVBoxLayout(side)
        # Student info and grading inputs
        self.info_label = QLabel("Name / ID")
        self.info_label.setStyleSheet("font-weight: bold;")
        self.score_edit = QLineEdit()
        self.score_edit.setPlaceholderText("Score")
        self.eval_edit = QTextEdit()
        self.history_btn = QPushButton("Reuse Past Eval")
        self.history_btn.clicked.connect(self._open_history)
        # Navigation and export
        nav = QHBoxLayout()
        self.prev_btn = QPushButton("Previous")
        self.next_btn = QPushButton("Next")
        nav.addWidget(self.prev_btn)
        nav.addWidget(self.next_btn)
        self.save_btn = QPushButton("Export Results")
        # Histogram canvas
        hist_label = QLabel("Score Distribution:")
        hist_label.setAlignment(Qt.AlignCenter)
        self.hist_canvas = FigureCanvas(Figure(figsize=(3,2)))
        # Assemble side panel
        side_layout.addWidget(self.info_label)
        side_layout.addWidget(self.score_edit)
        side_layout.addWidget(self.eval_edit)
        side_layout.addWidget(self.history_btn)
        side_layout.addLayout(nav)
        side_layout.addWidget(self.save_btn)
        side_layout.addWidget(hist_label)
        side_layout.addWidget(self.hist_canvas)
        side_layout.addStretch()
        main_layout.addWidget(side)
        root_layout.addLayout(main_layout)

        self.setCentralWidget(root)

        # Signal connections
        self.prev_btn.clicked.connect(self._go_previous)
        self.next_btn.clicked.connect(self._go_next)
        self.save_btn.clicked.connect(self._manual_export)
        self.student_view.page().scrollPositionChanged.connect(self._sync_scroll)

    def _show_software_details(self):
        details = (
            f"{APP_NAME} v{APP_VERSION}\n"
            f"Author: {APP_AUTHOR}\n"
            f"GitHub: https://github.com/Siyou-Li/QMplusScoring\n"
            f"Python: {sys.version.split()[0]}\n"
            f"Lab Key: {self.lab_key}\n"
            f"Students Loaded: {len(self.students)}"
        )
        QMessageBox.information(self, "Software Details", details)

    def _load_submissions(self, zip_path):
        self.tmpdir = tempfile.mkdtemp()
        with zipfile.ZipFile(zip_path, 'r') as z:
            z.extractall(self.tmpdir)
        self.students.clear()
        for d in os.scandir(self.tmpdir):
            if d.is_dir():
                name_id = d.name.split('_assignsubmission')[0]
                self.students.append((name_id, d.path))
        QMessageBox.information(self, "Loaded", f"Found {len(self.students)} students.")

    def _load_reference(self, nb_path):
        self.ref_path = nb_path

    def _extract_lab_notebook(self, student_path):
        for f in os.listdir(student_path):
            if f.endswith('.zip'):
                with zipfile.ZipFile(os.path.join(student_path, f), 'r') as z2:
                    tgt = os.path.join(student_path, 'work')
                    os.makedirs(tgt, exist_ok=True)
                    z2.extractall(tgt)
                    for root, _, files in os.walk(tgt):
                        for fn in files:
                            key = self.lab_key.lower().replace(' ', '')
                            if key in fn.lower().replace(' ', '') and fn.endswith('.ipynb'):
                                return os.path.join(root, fn)
        return None

    def _notebook_to_html(self, path):
        try:
            with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                nb = nbformat.read(f, as_version=4)
        except (NotJSONError, Exception) as err:
            QMessageBox.warning(self, "Load Error", f"Could not parse notebook: {os.path.basename(path)}")
            placeholder = '<html><body><h2>Error loading notebook</h2><p>File could not be parsed as JSON.</p></body></html>'
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.html')
            tmp.write(placeholder.encode('utf-8'))
            tmp.flush()
            return tmp.name
        exporter = HTMLExporter(template_name='classic')
        body, _ = exporter.from_notebook_node(nb)
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.html')
        tmp.write(body.encode('utf-8'))
        tmp.flush()
        return tmp.name

    def _load_current(self):
        # Save previous input
        self._save_current()
        name_id, spath = self.students[self.current_index]
        s_nb = self._extract_lab_notebook(spath)
        if not s_nb:
            QMessageBox.warning(self, "Missing", f"Lab {self.lab_key} not found for {name_id}")
            return
        s_html = self._notebook_to_html(s_nb)
        r_html = self._notebook_to_html(self.ref_path)
        self.student_view.load(QUrl.fromLocalFile(s_html))
        self.ref_view.load(QUrl.fromLocalFile(r_html))
        self.info_label.setText(name_id.replace('_', ' - '))
        # Load saved record
        recs = [r for r in self.records if r[0] == name_id and r[1] == self.current_index]
        if recs:
            last = recs[-1]
            self.score_edit.setText(last[2])
            self.eval_edit.setPlainText(last[3])
        else:
            self.score_edit.clear()
            self.eval_edit.clear()

    @pyqtSlot(QPointF)
    def _sync_scroll(self, pos):
        js = f"window.scrollTo(0, {int(pos.y())});"
        self.ref_view.page().runJavaScript(js)

    def _go_previous(self):
        if self.current_index > 0:
            self.current_index -= 1
            self._load_current()

    def _go_next(self):
        if self.current_index < len(self.students) - 1:
            self.current_index += 1
            self._load_current()

    def _save_current(self):
        name_id, idx = self.students[self.current_index]
        score = self.score_edit.text().strip()
        eval_text = self.eval_edit.toPlainText().strip()
        self.history = (*self.history, eval_text)
        if score or eval_text:
            # Remove old
            self.records = [r for r in self.records if not (r[0]==name_id and r[1]==idx)]
            self.records.append((name_id, idx, score, eval_text))

    def _open_history(self):
        #name_id, _ = self.students[self.current_index]
        evals = self.history
        if not evals:
            QMessageBox.information(self, "No History", "No previous evaluations")
            return
        dlg = HistoryDialog(self, evals)
        if dlg.exec_() == QDialog.Accepted and hasattr(dlg, 'selected'):
            self.eval_edit.setPlainText(dlg.selected)

    def _manual_export(self):
        # user-triggered export
        self._save_current()
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save results to CSV",
            "results.csv",
            "CSV Files (*.csv)"
        )
        if not path:
            return
        self._export_to_path(path)
        QMessageBox.information(self, "Exported", f"Results saved to {path}")

    def _export_to_path(self, path):
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(["Name", "ID", "Score", "Evaluation"])
            for name_id, idx, score, text in self.records:
                name, sid = name_id.split('_') if '_' in name_id else (name_id, '')
                writer.writerow([name, sid, score, text])

    def _show_software_details(self):
        details = (
            f"{APP_NAME} v{APP_VERSION}\n"
            f"Author: {APP_AUTHOR}\n"
            f"Python: {sys.version.split()[0]}\n"
            f"Lab Key: {self.lab_key}\n"
            f"Students Loaded: {len(self.students)}"
        )
        QMessageBox.information(self, "Software Details", details)
    
    def closeEvent(self, event):
        if hasattr(self, 'tmpdir'):
            shutil.rmtree(self.tmpdir, ignore_errors=True)
        super().closeEvent(event)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyleSheet(
        "QMainWindow { background: #f5f5f5; }"
        "QWidget { font-family: Arial; font-size: 14px; }"
        "QPushButton { border: none; padding: 8px; border-radius: 4px; background: #007acc; color: white; }"
        "QPushButton:hover { background: #005f99; }"
        "QLineEdit, QTextEdit { border: 1px solid #ccc; border-radius: 4px; padding: 4px; }"
    )
    app.setWindowIcon(QIcon("assets/logo.png"))
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setOrganizationName(APP_AUTHOR)
    app.setOrganizationDomain("github.com/Siyou-Li/QMplusScoring")
    app.setApplicationDisplayName(APP_NAME)
    win = QMPlusScoring()
    win.show()
    try:
        ret = app.exec_()
    except Exception:
        # Attempt to save current results on crash
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            crash_file = os.path.expanduser(f"~/homework_marker_crash_{timestamp}.csv")
            win._save_current()
            win._export_to_path(crash_file)
            QMessageBox.critical(
                win,
                "Unexpected Error",
                f"An unexpected error occurred and the session was terminated.\n"
                f"Your current grades/evaluations have been saved to:\n{crash_file}\n"
                f"Please restart the application."  )
        except Exception:
            # fallback
            traceback.print_exc()
        sys.exit(1)
    sys.exit(ret)
