import sys
import subprocess
from pathlib import Path

# ==========================
# Auto-install packages
# ==========================
def install_and_import(package, import_name=None):
    try:
        __import__(import_name if import_name else package)
    except ImportError:
        print(f"[INFO] Installing {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        __import__(import_name if import_name else package)

install_and_import("PySide6")
install_and_import("pymupdf", "fitz")
install_and_import("openpyxl")

import fitz
from openpyxl import Workbook
from PySide6.QtCore import Qt, QRect
from PySide6.QtGui import QPixmap, QImage, QPainter, QPen
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QLabel, QPushButton,
    QVBoxLayout, QWidget, QHBoxLayout, QDialog, QFormLayout,
    QLineEdit, QTextEdit, QCheckBox, QDialogButtonBox, QMessageBox,
    QListWidget, QListWidgetItem, QSplitter
)


# ==========================
# Config
# ==========================
DEFAULT_ORIENTATION = "Portrait"
CLICK_BOX_WIDTH = 20
CLICK_BOX_HEIGHT = 15


# ==========================
# Excel export
# ==========================
def export_to_excel(entries, filepath):
    wb = Workbook()

    # Annotation sheet
    ws = wb.active
    ws.title = "Annotation"
    ws.append([
        "DOMAIN", "NAME", "PAGENO", "ANNOTATION",
        "ASSIGNEDFIELD", "POSITION", "ORIENTATION"
    ])

    for row in entries:
        ws.append([
            row["DOMAIN"],
            row["NAME"],
            row["PAGENO"],
            row["ANNOTATION"],
            row["ASSIGNEDFIELD"],
            row["POSITION"],
            row["ORIENTATION"]
        ])

    # Length sheet
    ws2 = wb.create_sheet("Length")
    ws2.append(["CHAR", "LEN"])

    # Simple default widths; SAS macro can use these
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 .,()-:/_[]{}%"
    for c in chars:
        ws2.append([c, 5])

    wb.save(filepath)


# ==========================
# Entry dialog
# ==========================
class EntryDialog(QDialog):
    def __init__(self, page_no, position_text, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add Annotation Entry")

        layout = QFormLayout()

        self.domain_input = QLineEdit()
        self.name_input = QLineEdit()
        self.annotation_input = QTextEdit()
        self.assigned_chk = QCheckBox("Assigned Field")
        self.orientation_input = QLineEdit(DEFAULT_ORIENTATION)

        self.page_display = QLineEdit(str(page_no))
        self.page_display.setReadOnly(True)

        self.position_display = QLineEdit(position_text)
        self.position_display.setReadOnly(True)

        layout.addRow("Domain:", self.domain_input)
        layout.addRow("Name:", self.name_input)
        layout.addRow("Page No:", self.page_display)
        layout.addRow("Position:", self.position_display)
        layout.addRow("Orientation:", self.orientation_input)
        layout.addRow("Annotation:", self.annotation_input)
        layout.addRow(self.assigned_chk)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

    def validate_and_accept(self):
        if not self.domain_input.text().strip():
            QMessageBox.warning(self, "Missing Domain", "Please enter Domain.")
            return
        if not self.annotation_input.toPlainText().strip():
            QMessageBox.warning(self, "Missing Annotation", "Please enter Annotation.")
            return
        self.accept()

    def get_data(self):
        return {
            "DOMAIN": self.domain_input.text().strip(),
            "NAME": self.name_input.text().strip(),
            "ANNOTATION": self.annotation_input.toPlainText().strip(),
            "ASSIGNEDFIELD": "Y" if self.assigned_chk.isChecked() else "",
            "ORIENTATION": self.orientation_input.text().strip() or DEFAULT_ORIENTATION,
        }


# ==========================
# Clickable PDF label
# ==========================
class PDFLabel(QLabel):
    def __init__(self, parent_window):
        super().__init__()
        self.parent_window = parent_window
        self.setAlignment(Qt.AlignCenter)

    def mousePressEvent(self, event):
        self.parent_window.handle_pdf_click(event)


# ==========================
# Main window
# ==========================
class AnnotateCRFStudio(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("AnnotateCRF Studio")
        self.setGeometry(100, 100, 1200, 900)

        self.doc = None
        self.page_index = 0
        self.page = None
        self.current_qpixmap = None
        self.entries = []

        # Left: controls + page list
        self.open_btn = QPushButton("Open PDF")
        self.open_btn.clicked.connect(self.open_pdf)

        self.prev_btn = QPushButton("Prev")
        self.prev_btn.clicked.connect(self.prev_page)

        self.next_btn = QPushButton("Next")
        self.next_btn.clicked.connect(self.next_page)

        self.export_btn = QPushButton("Export Excel")
        self.export_btn.clicked.connect(self.export_excel)

        self.delete_btn = QPushButton("Delete Selected Entry")
        self.delete_btn.clicked.connect(self.delete_selected_entry)

        self.page_info = QLabel("Page: - / -")
        self.page_info.setAlignment(Qt.AlignCenter)

        self.entry_list = QListWidget()

        left_layout = QVBoxLayout()
        btn_row = QHBoxLayout()
        btn_row.addWidget(self.open_btn)
        btn_row.addWidget(self.prev_btn)
        btn_row.addWidget(self.next_btn)

        left_layout.addLayout(btn_row)
        left_layout.addWidget(self.page_info)
        left_layout.addWidget(self.export_btn)
        left_layout.addWidget(self.delete_btn)
        left_layout.addWidget(self.entry_list)

        left_widget = QWidget()
        left_widget.setLayout(left_layout)

        # Right: PDF view
        self.pdf_label = PDFLabel(self)
        self.pdf_label.setText("Open a PDF to start")

        right_layout = QVBoxLayout()
        right_layout.addWidget(self.pdf_label)

        right_widget = QWidget()
        right_widget.setLayout(right_layout)

        splitter = QSplitter()
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([350, 850])

        self.setCentralWidget(splitter)

    # --------------------------
    # Open/render PDF
    # --------------------------
    def open_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open PDF", "", "PDF Files (*.pdf)")
        if not file_path:
            return

        self.doc = fitz.open(file_path)
        self.page_index = 0
        self.load_page()

    def load_page(self):
        if not self.doc:
            return
        self.page = self.doc[self.page_index]
        self.render_page()
        self.page_info.setText(f"Page: {self.page_index + 1} / {len(self.doc)}")

    def render_page(self):
        if not self.page:
            return

        pix = self.page.get_pixmap()
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
        qpix = QPixmap.fromImage(img)
        self.current_qpixmap = qpix

        # draw simple markers for entries on current page
        display_pixmap = QPixmap(qpix)
        painter = QPainter(display_pixmap)
        pen = QPen(Qt.red)
        pen.setWidth(2)
        painter.setPen(pen)

        for entry in self.entries:
            if entry["PAGENO"] == str(self.page_index + 1):
                x1, y1, x2, y2 = [float(v) for v in entry["POSITION"].split(",")]
                sx, sy = self.pdf_to_screen(x1, y1)
                ex, ey = self.pdf_to_screen(x2, y2)
                painter.drawRect(QRect(int(sx), int(sy), max(1, int(ex - sx)), max(1, int(ey - sy))))

        painter.end()
        self.pdf_label.setPixmap(display_pixmap)
        self.refresh_entry_list()

    # --------------------------
    # Page navigation
    # --------------------------
    def next_page(self):
        if self.doc and self.page_index < len(self.doc) - 1:
            self.page_index += 1
            self.load_page()

    def prev_page(self):
        if self.doc and self.page_index > 0:
            self.page_index -= 1
            self.load_page()

    # --------------------------
    # Coordinate transforms
    # --------------------------
    def screen_to_pdf(self, x, y):
        if not self.page or not self.pdf_label.pixmap():
            return None, None

        rect = self.pdf_label.pixmap().rect()
        pdf_rect = self.page.rect

        scale_x = pdf_rect.width / rect.width()
        scale_y = pdf_rect.height / rect.height()

        return x * scale_x, y * scale_y

    def pdf_to_screen(self, x, y):
        if not self.page or not self.pdf_label.pixmap():
            return None, None

        rect = self.pdf_label.pixmap().rect()
        pdf_rect = self.page.rect

        scale_x = rect.width() / pdf_rect.width
        scale_y = rect.height() / pdf_rect.height

        return x * scale_x, y * scale_y

    # --------------------------
    # Click handling
    # --------------------------
    def handle_pdf_click(self, event):
        if not self.page or not self.pdf_label.pixmap():
            return

        sx = event.position().x()
        sy = event.position().y()

        pdf_x, pdf_y = self.screen_to_pdf(sx, sy)
        if pdf_x is None:
            return

        x1 = round(pdf_x, 2)
        y1 = round(pdf_y, 2)
        x2 = round(pdf_x + CLICK_BOX_WIDTH, 2)
        y2 = round(pdf_y + CLICK_BOX_HEIGHT, 2)

        position_text = f"{x1},{y1},{x2},{y2}"
        page_no = self.page_index + 1

        dialog = EntryDialog(page_no=page_no, position_text=position_text, parent=self)
        if dialog.exec() != QDialog.Accepted:
            return

        data = dialog.get_data()
        entry = {
            "DOMAIN": data["DOMAIN"],
            "NAME": data["NAME"],
            "PAGENO": str(page_no),
            "ANNOTATION": data["ANNOTATION"],
            "ASSIGNEDFIELD": data["ASSIGNEDFIELD"],
            "POSITION": position_text,
            "ORIENTATION": data["ORIENTATION"],
        }
        self.entries.append(entry)
        self.render_page()

    # --------------------------
    # Entry list
    # --------------------------
    def refresh_entry_list(self):
        self.entry_list.clear()
        for i, entry in enumerate(self.entries, start=1):
            text = (
                f"{i}. Pg {entry['PAGENO']} | {entry['DOMAIN']} | "
                f"{entry['NAME']} | {entry['POSITION']}"
            )
            item = QListWidgetItem(text)
            self.entry_list.addItem(item)

    def delete_selected_entry(self):
        row = self.entry_list.currentRow()
        if row < 0 or row >= len(self.entries):
            return
        del self.entries[row]
        self.render_page()

    # --------------------------
    # Export
    # --------------------------
    def export_excel(self):
        if not self.entries:
            QMessageBox.information(self, "No Entries", "No annotation entries to export.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Annotation Excel",
            str(Path.cwd() / "Annotation.xlsx"),
            "Excel Files (*.xlsx)"
        )
        if not file_path:
            return

        export_to_excel(self.entries, file_path)
        QMessageBox.information(self, "Success", f"Excel created:\n{file_path}")


# ==========================
# Run
# ==========================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AnnotateCRFStudio()
    window.show()
    sys.exit(app.exec())