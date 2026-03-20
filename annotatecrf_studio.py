# ====================================================================================================
# Script Name    : Annotator_Studio_DirectPDF_Bookmarks.py
#
# Description    :
#                  Desktop GUI utility for capturing annotation positions from a PDF and generating
#                  final visible comment boxes directly into a PDF without XFDF / Adobe dependency.
#
#                  Behavior:
#                    - Top mode selector: Annotation / Bookmark
#                    - Annotation mode:
#                        * Click on PDF -> enter details -> preview box appears immediately
#                        * Drag an existing preview box to move it
#                    - Bookmark mode:
#                        * Add Bookmark button opens bookmark dialog
#                        * Bookmark dialog asks only Bookmark Text, Level, Page No
#                    - Generate Final Output PDF:
#                        * annotations only, if only annotations exist
#                        * bookmarks only, if only bookmarks exist
#                        * both, if both exist
#
#                  Annotation CSV columns:
#                    DOMAIN,NAME,PAGENO,ANNOTATION,ASSIGNEDFIELD,X1,Y1,PAGEH
#
#                  Bookmark CSV columns:
#                    TITLE,LEVEL,PAGENO
# ====================================================================================================

import csv
import importlib
import os
import site
import subprocess
import sys
from collections import OrderedDict
from dataclasses import dataclass
from typing import List, Optional

# ================================
# Auto-install required packages
# ================================
def ensure_user_site():
    user_site = site.getusersitepackages()
    if user_site not in sys.path:
        sys.path.insert(0, user_site)


def install_if_missing(package_name, import_name=None):
    try:
        importlib.import_module(import_name or package_name)
    except ImportError:
        print(f"[Installing] {package_name} (user mode)...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", package_name])
        print(f"[Done] {package_name} installed.")


ensure_user_site()
install_if_missing("PyMuPDF", "fitz")
install_if_missing("PyQt5")

from PyQt5 import QtCore, QtGui, QtWidgets

QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)

import fitz

# ================================
# Settings
# ================================
DEFAULT_ZOOM = 1.30

DOMAIN_COLORS = [
    (0.75, 1.00, 1.00),   # cyan
    (0.59, 1.00, 0.59),   # green
    (1.00, 0.75, 0.61),   # peach
    (1.00, 0.80, 0.55),   # orange
    (0.86, 0.82, 1.00),   # lavender
]
DEFAULT_OTHER_FILL = (0.85, 0.85, 0.85)
NOTSUB_FILL = (1.00, 0.93, 0.55)
TEXT_COLOR = (0.0, 0.0, 0.0)
BORDER_COLOR = (0.0, 0.0, 0.0)

FONT_NORMAL = "helv"
FONT_BOLD = "hebo"

BASE_FONT_SIZE = 10
DOMAIN_FONT_SIZE = 12

TEXT_PADDING_X = 3.0
TEXT_PADDING_Y = 1.5
BOX_HEIGHT_NORMAL = 15.0
BOX_HEIGHT_DOMAIN = 18.0
MAX_WRAP_WIDTH = 240.0

BOOKMARK_PREVIEW_COLOR = QtGui.QColor("#7c3aed")
BOOKMARK_PREVIEW_TEXT_COLOR = QtGui.QColor("#5b21b6")

CHAR_WIDTHS = {
    '!': 3.43, '"': 4.82, '#': 5.65, '$': 5.65, '%': 8.99, "'": 2.47,
    '(': 3.41, ')': 3.40, '*': 4.00, '+': 5.93, ',': 2.84, '-': 3.42,
    '.': 2.88, '/': 2.91, '0': 5.65, '1': 5.66, '2': 5.64, '3': 5.64,
    '4': 5.65, '5': 5.65, '6': 5.65, '7': 5.66, '8': 5.64, '9': 5.65,
    ':': 3.42, ';': 3.41, '<': 5.95, '=': 5.95, '>': 5.91, '?': 6.20,
    '@': 9.84, 'A': 7.31, 'B': 7.31, 'C': 7.32, 'D': 7.34, 'E': 6.75,
    'F': 6.22, 'G': 7.89, 'H': 7.31, 'I': 2.89, 'J': 5.67, 'K': 7.30,
    'L': 6.21, 'M': 8.43, 'N': 7.31, 'O': 7.86, 'P': 6.78, 'Q': 7.89,
    'R': 7.29, 'S': 6.75, 'T': 6.22, 'U': 7.30, 'V': 6.76, 'W': 9.58,
    'X': 6.76, 'Y': 6.80, 'Z': 6.23, '[': 3.43, '\\': 2.87, ']': 3.41,
    '^': 5.95, '_': 5.68, '`': 3.42, 'a': 5.67, 'b': 6.21, 'c': 5.64,
    'd': 6.21, 'e': 5.64, 'f': 3.43, 'g': 6.18, 'h': 6.18, 'i': 2.86,
    'j': 2.86, 'k': 5.64, 'l': 2.87, 'm': 8.97, 'n': 6.22, 'o': 6.21,
    'p': 6.20, 'q': 6.18, 'r': 3.98, 's': 5.68, 't': 3.43, 'u': 6.21,
    'v': 5.64, 'w': 7.89, 'x': 5.65, 'y': 5.66, 'z': 5.10, '{': 4.01,
    '|': 2.87, '}': 3.98, '~': 5.94, ' ': 2.90
}

# ================================
# Helpers
# ================================
def estimate_text_width(text: str, scale: float = 1.0) -> float:
    total = 0.0
    for ch in text:
        total += CHAR_WIDTHS.get(ch, 6.0)
    return total * scale + 5 * scale


def wrap_text_by_width(text: str, max_width: float, scale: float = 1.0) -> List[str]:
    if not text:
        return [""]
    words = text.split()
    if not words:
        return [text]

    lines = []
    current = words[0]
    for word in words[1:]:
        trial = current + " " + word
        if estimate_text_width(trial, scale) <= max_width:
            current = trial
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


def rect_from_top_origin(x1: float, y1_top: float, width: float, height: float, pageh: float) -> fitz.Rect:
    return fitz.Rect(x1, y1_top, x1 + width, y1_top + height)


def get_domain_color_map(entries) -> dict:
    ordered_domains = OrderedDict()
    for e in entries:
        dom = (e.domain or "").strip().upper()
        if dom and dom != "NOTSUB" and dom not in ordered_domains:
            ordered_domains[dom] = None

    color_map = {}
    for i, dom in enumerate(ordered_domains.keys(), start=1):
        color_map[dom] = DOMAIN_COLORS[i - 1] if i <= len(DOMAIN_COLORS) else DEFAULT_OTHER_FILL
    return color_map


def compute_entry_layout(entry, color_map):
    domain_key = (entry.domain or "").strip().upper()

    fill = NOTSUB_FILL if entry.is_not_submitted else color_map.get(domain_key, DEFAULT_OTHER_FILL)
    dashed = bool(entry.is_assigned_field)
    bold = bool(entry.is_domain_annotation)
    font_size = DOMAIN_FONT_SIZE if bold else BASE_FONT_SIZE
    scale = DOMAIN_FONT_SIZE / BASE_FONT_SIZE if bold else 1.0
    base_h = BOX_HEIGHT_DOMAIN if bold else BOX_HEIGHT_NORMAL

    text = (entry.annotation or "").strip()
    raw_w = estimate_text_width(text, scale=scale)
    wrap_width = MAX_WRAP_WIDTH if raw_w > MAX_WRAP_WIDTH else raw_w
    lines = wrap_text_by_width(text, wrap_width, scale=scale)
    actual_w = max(estimate_text_width(line, scale=scale) for line in lines) if lines else raw_w
    box_w = min(MAX_WRAP_WIDTH, actual_w)
    box_h = max(base_h, len(lines) * (font_size + 1.2) + 4)

    return {
        "fill": fill,
        "dashed": dashed,
        "bold": bold,
        "font_size": font_size,
        "lines": lines,
        "box_w": box_w,
        "box_h": box_h
    }


def qcolor_from_rgb01(rgb):
    return QtGui.QColor(int(rgb[0] * 255), int(rgb[1] * 255), int(rgb[2] * 255))


def draw_box_and_text_pdf(page, rect, text_lines, fill_color, bold=False, dashed=False, font_size=10):
    page.draw_rect(rect, color=None, fill=fill_color, overlay=True)

    dashes = "[3 3] 0" if dashed else None
    page.draw_rect(rect, color=BORDER_COLOR, fill=None, width=0.8, dashes=dashes, overlay=True)

    fontname = FONT_BOLD if bold else FONT_NORMAL
    line_gap = font_size + 1.2
    text_x = rect.x0 + TEXT_PADDING_X
    text_y = rect.y0 + font_size + TEXT_PADDING_Y

    for line in text_lines:
        try:
            page.insert_text(
                fitz.Point(text_x, text_y),
                line,
                fontname=fontname,
                fontsize=font_size,
                color=TEXT_COLOR,
                overlay=True
            )
        except Exception:
            page.insert_text(
                fitz.Point(text_x, text_y),
                line,
                fontname=FONT_NORMAL,
                fontsize=font_size,
                color=TEXT_COLOR,
                overlay=True
            )
        text_y += line_gap


def sanitize_output_pdf_path(src_pdf: str):
    base, ext = os.path.splitext(src_pdf)
    return base + "_final.pdf"


# ================================
# Data Models
# ================================
@dataclass
class AnnotationEntry:
    domain: str
    name: str
    pageno: int
    annotation: str
    assignedfield: str
    x1: float
    y1: float
    pageh: float
    is_domain_annotation: bool
    is_assigned_field: bool
    is_not_submitted: bool


@dataclass
class BookmarkEntry:
    title: str
    level: int
    pageno: int


# ================================
# Dialog for annotation details
# ================================
class AnnotationDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Annotation Details")
        self.resize(700, 460)
        self.setMinimumSize(700, 460)
        self.setModal(True)

        self.setStyleSheet("""
            QDialog { background-color: #f4f8ff; border-radius: 14px; }
            QLabel { font-family: 'Times New Roman'; font-size: 12pt; color: #1c2e4a; }
            QLineEdit, QTextEdit {
                background: #ffffff; border: 1px solid #a8bfdc; border-radius: 8px;
                padding: 6px 8px; font-family: 'Times New Roman'; font-size: 12pt; color: #1a1a1a;
            }
            QLineEdit:focus, QTextEdit:focus { border: 2px solid #4d8ef7; background: #fdfefe; }
            QCheckBox { font-family: 'Times New Roman'; font-size: 12pt; color: #143b66; spacing: 8px; }
            QCheckBox::indicator { width: 18px; height: 18px; }
            QDialogButtonBox QPushButton {
                font-family: 'Times New Roman'; font-size: 12pt; font-weight: bold;
                border-radius: 10px; padding: 8px 18px; min-width: 100px;
            }
        """)

        self.domain_edit = QtWidgets.QLineEdit()
        self.name_edit = QtWidgets.QLineEdit()
        self.annotation_edit = QtWidgets.QTextEdit()
        self.annotation_edit.setMinimumHeight(140)

        self.chk_domain = QtWidgets.QCheckBox("Domain")
        self.chk_assigned = QtWidgets.QCheckBox("Assigned field")
        self.chk_notsub = QtWidgets.QCheckBox("Not Submitted")

        self.chk_domain.stateChanged.connect(self.toggle_dialog_state)
        self.chk_assigned.stateChanged.connect(self.toggle_dialog_state)
        self.chk_notsub.stateChanged.connect(self.toggle_dialog_state)

        title = QtWidgets.QLabel("Enter Annotation Metadata")
        title.setAlignment(QtCore.Qt.AlignCenter)
        title.setStyleSheet("""
            QLabel {
                background: #cfe7ff; color: #103760; font-family: 'Times New Roman';
                font-size: 15pt; font-weight: bold; padding: 12px; border-radius: 12px;
            }
        """)

        self.lbl_domain = QtWidgets.QLabel("DOMAIN")
        self.lbl_name = QtWidgets.QLabel("VARIABLE")
        self.lbl_annotation = QtWidgets.QLabel("ANNOTATION")

        form = QtWidgets.QFormLayout()
        form.setLabelAlignment(QtCore.Qt.AlignRight)
        form.setHorizontalSpacing(18)
        form.setVerticalSpacing(14)
        form.addRow(self.lbl_domain, self.domain_edit)
        form.addRow(self.lbl_name, self.name_edit)
        form.addRow(self.lbl_annotation, self.annotation_edit)

        anno_type_hdr = QtWidgets.QLabel("Annotation type")
        anno_type_hdr.setStyleSheet("""
            QLabel { font-family: 'Times New Roman'; font-size: 12pt; font-weight: bold; color: #103760; padding-top: 6px; }
        """)

        chk_row = QtWidgets.QHBoxLayout()
        chk_row.setSpacing(24)
        chk_row.addWidget(self.chk_domain)
        chk_row.addWidget(self.chk_assigned)
        chk_row.addWidget(self.chk_notsub)
        chk_row.addStretch()

        note = QtWidgets.QLabel("If none is selected, it is treated as a variable annotation for a collected field.")
        note.setWordWrap(True)
        note.setStyleSheet("QLabel { color: #556b84; font-size: 11pt; font-style: italic; }")

        self.buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        self.buttons.accepted.connect(self.validate_and_accept)
        self.buttons.rejected.connect(self.reject)

        ok_btn = self.buttons.button(QtWidgets.QDialogButtonBox.Ok)
        cancel_btn = self.buttons.button(QtWidgets.QDialogButtonBox.Cancel)
        ok_btn.setStyleSheet("QPushButton { background-color: #55c16d; color: white; } QPushButton:hover { background-color: #73d789; }")
        cancel_btn.setStyleSheet("QPushButton { background-color: #f16a6a; color: white; } QPushButton:hover { background-color: #f48f8f; }")

        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(title)
        layout.addSpacing(8)
        layout.addLayout(form)
        layout.addSpacing(8)
        layout.addWidget(anno_type_hdr)
        layout.addLayout(chk_row)
        layout.addWidget(note)
        layout.addSpacing(10)
        layout.addWidget(self.buttons)

        self.toggle_dialog_state()

    def toggle_dialog_state(self):
        is_domain = self.chk_domain.isChecked()
        is_notsub = self.chk_notsub.isChecked()

        self.lbl_name.setVisible(not is_domain)
        self.name_edit.setVisible(not is_domain)

        if is_domain:
            self.name_edit.clear()

        if is_notsub:
            self.domain_edit.setText("NOTSUB")
            self.name_edit.setText("NOTSUB")
            self.annotation_edit.setPlainText("[NOT SUBMITTED]")
            self.domain_edit.setDisabled(True)
            self.name_edit.setDisabled(True)
            self.annotation_edit.setDisabled(True)
            self.chk_domain.setChecked(False)
            self.chk_domain.setDisabled(True)
            self.chk_assigned.setChecked(False)
            self.chk_assigned.setDisabled(True)
        else:
            self.domain_edit.setDisabled(False)
            self.name_edit.setDisabled(False)
            self.annotation_edit.setDisabled(False)
            self.chk_domain.setDisabled(False)
            self.chk_assigned.setDisabled(False)

            if self.domain_edit.text().strip() == "NOTSUB":
                self.domain_edit.clear()
            if self.name_edit.text().strip() == "NOTSUB":
                self.name_edit.clear()
            if self.annotation_edit.toPlainText().strip() == "[NOT SUBMITTED]":
                self.annotation_edit.clear()

            is_domain = self.chk_domain.isChecked()
            self.lbl_name.setVisible(not is_domain)
            self.name_edit.setVisible(not is_domain)
            if is_domain:
                self.name_edit.clear()

    def validate_and_accept(self):
        if not self.domain_edit.text().strip():
            QtWidgets.QMessageBox.warning(self, "Validation", "DOMAIN is required.")
            return
        if not self.annotation_edit.toPlainText().strip():
            QtWidgets.QMessageBox.warning(self, "Validation", "ANNOTATION is required.")
            return
        if not self.chk_domain.isChecked() and not self.chk_notsub.isChecked():
            if not self.name_edit.text().strip():
                QtWidgets.QMessageBox.warning(self, "Validation", "VARIABLE is required.")
                return
        self.accept()

    def get_values(self):
        is_domain_annotation = self.chk_domain.isChecked()
        is_assigned_field = self.chk_assigned.isChecked()
        is_not_submitted = self.chk_notsub.isChecked()

        if is_not_submitted:
            return {
                "domain": "NOTSUB",
                "name": "NOTSUB",
                "annotation": "[NOT SUBMITTED]",
                "assignedfield": "",
                "is_domain_annotation": False,
                "is_assigned_field": False,
                "is_not_submitted": True
            }

        return {
            "domain": self.domain_edit.text().strip(),
            "name": "" if is_domain_annotation else self.name_edit.text().strip(),
            "annotation": self.annotation_edit.toPlainText().strip(),
            "assignedfield": "Y" if is_assigned_field else "",
            "is_domain_annotation": is_domain_annotation,
            "is_assigned_field": is_assigned_field,
            "is_not_submitted": False
        }


# ================================
# Dialog for bookmark details
# ================================
class BookmarkDialog(QtWidgets.QDialog):
    def __init__(self, page_no: int, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Bookmark Details")
        self.resize(520, 230)
        self.setMinimumSize(520, 230)
        self.setModal(True)

        self.setStyleSheet("""
            QDialog { background-color: #f4f8ff; border-radius: 14px; }
            QLabel { font-family: 'Times New Roman'; font-size: 12pt; color: #1c2e4a; }
            QLineEdit, QSpinBox {
                background: #ffffff; border: 1px solid #a8bfdc; border-radius: 8px;
                padding: 6px 8px; font-family: 'Times New Roman'; font-size: 12pt; color: #1a1a1a;
            }
            QDialogButtonBox QPushButton {
                font-family: 'Times New Roman'; font-size: 12pt; font-weight: bold;
                border-radius: 10px; padding: 8px 18px; min-width: 100px;
            }
        """)

        title_hdr = QtWidgets.QLabel("Enter Bookmark Metadata")
        title_hdr.setAlignment(QtCore.Qt.AlignCenter)
        title_hdr.setStyleSheet("""
            QLabel {
                background: #ded6ff; color: #3f2376; font-family: 'Times New Roman';
                font-size: 15pt; font-weight: bold; padding: 12px; border-radius: 12px;
            }
        """)

        self.title_edit = QtWidgets.QLineEdit()

        self.level_spin = QtWidgets.QSpinBox()
        self.level_spin.setRange(1, 9)
        self.level_spin.setValue(1)

        self.page_spin = QtWidgets.QSpinBox()
        self.page_spin.setRange(1, 999999)
        self.page_spin.setValue(page_no)

        form = QtWidgets.QFormLayout()
        form.setLabelAlignment(QtCore.Qt.AlignRight)
        form.setHorizontalSpacing(18)
        form.setVerticalSpacing(14)
        form.addRow("Bookmark Text", self.title_edit)
        form.addRow("Level", self.level_spin)
        form.addRow("Page No", self.page_spin)

        note = QtWidgets.QLabel("Bookmark will be added to the PDF outline for the selected page.")
        note.setWordWrap(True)
        note.setStyleSheet("QLabel { color: #556b84; font-size: 11pt; font-style: italic; }")

        self.buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        self.buttons.accepted.connect(self.validate_and_accept)
        self.buttons.rejected.connect(self.reject)

        ok_btn = self.buttons.button(QtWidgets.QDialogButtonBox.Ok)
        cancel_btn = self.buttons.button(QtWidgets.QDialogButtonBox.Cancel)
        ok_btn.setStyleSheet("QPushButton { background-color: #55c16d; color: white; } QPushButton:hover { background-color: #73d789; }")
        cancel_btn.setStyleSheet("QPushButton { background-color: #f16a6a; color: white; } QPushButton:hover { background-color: #f48f8f; }")

        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(title_hdr)
        layout.addSpacing(10)
        layout.addLayout(form)
        layout.addWidget(note)
        layout.addSpacing(8)
        layout.addWidget(self.buttons)

    def validate_and_accept(self):
        if not self.title_edit.text().strip():
            QtWidgets.QMessageBox.warning(self, "Validation", "Bookmark Text is required.")
            return
        self.accept()

    def get_values(self):
        return {
            "title": self.title_edit.text().strip(),
            "level": int(self.level_spin.value()),
            "pageno": int(self.page_spin.value())
        }


# ================================
# PDF display label with preview + drag move
# ================================
class PdfLabel(QtWidgets.QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = None
        self.last_click_point: Optional[QtCore.QPoint] = None
        self.drag_entry_index: Optional[int] = None
        self.drag_offset = QtCore.QPointF(0, 0)
        self.dragging = False

    def _get_entry_rect_on_screen(self, entry, color_map):
        layout = compute_entry_layout(entry, color_map)
        rect_pdf = rect_from_top_origin(entry.x1, entry.y1, layout["box_w"], layout["box_h"], entry.pageh)
        zoom = self.main_window.zoom
        return QtCore.QRectF(
            rect_pdf.x0 * zoom,
            rect_pdf.y0 * zoom,
            rect_pdf.width * zoom,
            rect_pdf.height * zoom
        )

    def mousePressEvent(self, event):
        if not self.main_window or self.main_window.page_pixmap is None:
            return

        if event.button() == QtCore.Qt.LeftButton:
            current_page = self.main_window.current_page_index + 1
            color_map = get_domain_color_map(self.main_window.entries)

            # drag existing annotation only in annotation mode
            if self.main_window.active_mode == "annotation":
                for idx in reversed(range(len(self.main_window.entries))):
                    entry = self.main_window.entries[idx]
                    if entry.pageno != current_page:
                        continue
                    rect = self._get_entry_rect_on_screen(entry, color_map)
                    if rect.contains(event.pos()):
                        self.drag_entry_index = idx
                        self.dragging = True
                        self.drag_offset = QtCore.QPointF(event.pos()) - rect.topLeft()
                        self.main_window.selected_entry_index = idx
                        self.main_window.refresh_selection_only()
                        self.update()
                        return

            self.last_click_point = event.pos()
            self.main_window.store_last_click(event.pos())
            self.update()

            if self.main_window.active_mode == "annotation":
                self.main_window.capture_annotation_point(event.pos())

    def mouseMoveEvent(self, event):
        if not self.main_window or self.drag_entry_index is None or not self.dragging:
            return

        entry = self.main_window.entries[self.drag_entry_index]
        new_top_left = QtCore.QPointF(event.pos()) - self.drag_offset
        new_x = max(0.0, new_top_left.x() / self.main_window.zoom)
        new_y = max(0.0, new_top_left.y() / self.main_window.zoom)

        entry.x1 = round(new_x, 6)
        entry.y1 = round(new_y, 6)

        self.main_window.refresh_annotation_table()
        self.main_window.select_annotation_row_silent(self.drag_entry_index)
        self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.drag_entry_index = None
            self.dragging = False

    def paintEvent(self, event):
        super().paintEvent(event)

        if not self.main_window:
            return

        painter = QtGui.QPainter(self)
        zoom = self.main_window.zoom
        current_page = self.main_window.current_page_index + 1
        color_map = get_domain_color_map(self.main_window.entries)

        # Annotation previews
        for idx, entry in enumerate(self.main_window.entries):
            if entry.pageno != current_page:
                continue

            layout = compute_entry_layout(entry, color_map)
            rect_pdf = rect_from_top_origin(entry.x1, entry.y1, layout["box_w"], layout["box_h"], entry.pageh)
            rect = QtCore.QRectF(
                rect_pdf.x0 * zoom,
                rect_pdf.y0 * zoom,
                rect_pdf.width * zoom,
                rect_pdf.height * zoom
            )

            fill_q = qcolor_from_rgb01(layout["fill"])
            fill_q.setAlpha(210)
            painter.fillRect(rect, fill_q)

            pen = QtGui.QPen(QtGui.QColor(0, 0, 0), 1)
            if layout["dashed"]:
                pen.setStyle(QtCore.Qt.DashLine)
            painter.setPen(pen)
            painter.drawRect(rect)

            font = QtGui.QFont("Arial")
            font.setPointSizeF(layout["font_size"] * zoom * 0.75)
            font.setBold(layout["bold"])
            painter.setFont(font)
            painter.setPen(QtGui.QColor(0, 0, 0))

            line_gap = (layout["font_size"] + 1.2) * zoom * 0.75
            tx = rect.left() + TEXT_PADDING_X * zoom
            ty = rect.top() + (layout["font_size"] + TEXT_PADDING_Y) * zoom * 0.75

            for line in layout["lines"]:
                painter.drawText(QtCore.QPointF(tx, ty), line)
                ty += line_gap

            mx = entry.x1 * zoom
            my = entry.y1 * zoom
            if idx == self.main_window.selected_entry_index and self.main_window.tabs.currentIndex() == 0:
                sel_pen = QtGui.QPen(QtGui.QColor("#ff00aa"), 3)
                painter.setPen(sel_pen)
                painter.drawEllipse(QtCore.QPointF(mx, my), 6, 6)
            else:
                dot_pen = QtGui.QPen(QtGui.QColor("#1d7df2"), 2)
                painter.setPen(dot_pen)
                painter.setBrush(QtGui.QBrush(QtGui.QColor("#7fc2ff")))
                painter.drawEllipse(QtCore.QPointF(mx, my), 4, 4)

        # Bookmark preview by page only
        bookmark_pen = QtGui.QPen(BOOKMARK_PREVIEW_COLOR, 2)
        bookmark_pen.setStyle(QtCore.Qt.DashLine)
        painter.setPen(bookmark_pen)

        y_base = 28
        y_gap = 18
        page_bookmarks = [b for b in self.main_window.bookmarks if b.pageno == current_page]

        for idx, bm in enumerate(page_bookmarks):
            sy = y_base + idx * y_gap
            painter.drawLine(0, sy, min(self.width(), 260), sy)

            text_pen = QtGui.QPen(BOOKMARK_PREVIEW_TEXT_COLOR, 1)
            painter.setPen(text_pen)
            f = QtGui.QFont("Arial", 9)
            is_selected = (
                self.main_window.tabs.currentIndex() == 1
                and 0 <= self.main_window.selected_bookmark_index < len(self.main_window.bookmarks)
                and self.main_window.bookmarks[self.main_window.selected_bookmark_index] == bm
            )
            f.setBold(is_selected)
            painter.setFont(f)

            offset_x = 8 + (max(1, bm.level) - 1) * 12
            painter.drawText(QtCore.QPointF(offset_x, sy - 4), f"BM L{bm.level}: {bm.title[:45]}")
            painter.setPen(bookmark_pen)

        # Last click crosshair
        if self.last_click_point:
            pen = QtGui.QPen(QtGui.QColor("#ff3b30"), 2)
            painter.setPen(pen)
            x = self.last_click_point.x()
            y = self.last_click_point.y()
            painter.drawLine(x - 7, y, x + 7, y)
            painter.drawLine(x, y - 7, x, y + 7)


# ================================
# Main Application
# ================================
class AnnotatorApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Annotator Studio - Direct PDF")
        self.resize(1280, 900)
        self.setMinimumSize(1100, 760)
        self.setStyleSheet("background-color: #f3f7fd;")

        self.doc = None
        self.open_pdf_path = ""
        self.current_page_index = 0
        self.zoom = DEFAULT_ZOOM
        self.page_pixmap = None
        self.page_rect = None

        self.entries: List[AnnotationEntry] = []
        self.bookmarks: List[BookmarkEntry] = []

        self.selected_entry_index = -1
        self.selected_bookmark_index = -1

        self.current_annotation_csv_path = ""
        self.current_bookmark_csv_path = ""

        self._suppress_annotation_selection_signal = False
        self._suppress_bookmark_selection_signal = False

        self.last_click_pdf_x = None
        self.last_click_pdf_y = None
        self.last_click_pageh = None

        self.active_mode = "annotation"

        self.build_ui()

    def build_ui(self):
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(8)

        header = QtWidgets.QLabel("Annotator Studio")
        header.setAlignment(QtCore.Qt.AlignCenter)
        header.setStyleSheet("""
            QLabel {
                color: black; background: #bfe9f7; padding: 10px 0 8px 0; border-radius: 16px;
                font-family: 'Times New Roman'; font-size: 20pt; font-weight: bold;
            }
        """)
        layout.addWidget(header)

        sub = QtWidgets.QLabel(
            "Choose mode at top. Annotation mode allows PDF click capture. Bookmark mode adds page-level bookmarks."
        )
        sub.setAlignment(QtCore.Qt.AlignCenter)
        sub.setWordWrap(True)
        sub.setStyleSheet("QLabel { color: #21466d; padding: 0 0 4px 0; font-family: 'Times New Roman'; font-size: 12pt; }")
        layout.addWidget(sub)

        # Mode row
        mode_row = QtWidgets.QHBoxLayout()
        mode_row.setSpacing(10)

        mode_lbl = QtWidgets.QLabel("Selection")
        mode_lbl.setStyleSheet("""
            QLabel {
                font-family: 'Times New Roman';
                font-size: 12pt;
                font-weight: bold;
                color: #153b68;
            }
        """)

        self.mode_combo = QtWidgets.QComboBox()
        self.mode_combo.addItems(["Annotation", "Bookmark"])
        self.mode_combo.setFixedWidth(180)
        self.mode_combo.setStyleSheet("""
            QComboBox {
                background: white;
                border: 1px solid #a8bfdc;
                border-radius: 8px;
                padding: 6px 10px;
                font-family: 'Times New Roman';
                font-size: 12pt;
            }
        """)

        mode_row.addStretch()
        mode_row.addWidget(mode_lbl)
        mode_row.addWidget(self.mode_combo)
        mode_row.addStretch()
        layout.addLayout(mode_row)

        btn_row = QtWidgets.QHBoxLayout()
        btn_row.setSpacing(8)

        btn_base = """
            QPushButton {
                font-family: 'Times New Roman'; font-weight: bold; font-size: 11pt;
                border-radius: 10px; padding: 6px 14px; min-height: 34px;
            }
        """

        def mkbtn(text, bg, hov, fg="white"):
            b = QtWidgets.QPushButton(text)
            b.setStyleSheet(btn_base + f"QPushButton {{ background-color: {bg}; color: {fg}; }} QPushButton:hover {{ background-color: {hov}; }}")
            return b

        self.btn_open = mkbtn("Open PDF", "#b3c6e7", "#d5deef", "#222")
        self.btn_prev = mkbtn("Previous", "#ad7dfc", "#c7a2fd")
        self.btn_next = mkbtn("Next", "#ad7dfc", "#c7a2fd")
        self.btn_add_bookmark = mkbtn("Add Bookmark", "#8b5cf6", "#a78bfa")
        self.btn_delete = mkbtn("Delete Selected Row", "#f16a6a", "#f48f8f")
        self.btn_save_ann_csv = mkbtn("Export Annotation CSV", "#71c95e", "#98eb84")
        self.btn_load_ann_csv = mkbtn("Load Annotation CSV", "#5bb7a6", "#7ed2c2")
        self.btn_save_bm_csv = mkbtn("Export Bookmark CSV", "#71c95e", "#98eb84")
        self.btn_load_bm_csv = mkbtn("Load Bookmark CSV", "#5bb7a6", "#7ed2c2")
        self.btn_generate_pdf = mkbtn("Generate Final Output PDF", "#ff9933", "#ffbc80", "#222")

        for b in [
            self.btn_open, self.btn_prev, self.btn_next,
            self.btn_add_bookmark, self.btn_delete,
            self.btn_save_ann_csv, self.btn_load_ann_csv,
            self.btn_save_bm_csv, self.btn_load_bm_csv,
            self.btn_generate_pdf
        ]:
            btn_row.addWidget(b)

        btn_row.addStretch()

        self.page_info = QtWidgets.QLabel("Page: -")
        self.page_info.setStyleSheet("""
            QLabel {
                background: #e6f2fb; color: #12608d; border-radius: 8px; padding: 5px 10px;
                font-family: 'Times New Roman'; font-size: 11pt; font-weight: bold;
            }
        """)
        btn_row.addWidget(self.page_info)
        layout.addLayout(btn_row)

        self.splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.setHandleWidth(10)
        self.splitter.setStretchFactor(0, 8)
        self.splitter.setStretchFactor(1, 2)

        pdf_widget = QtWidgets.QWidget()
        pdf_layout = QtWidgets.QVBoxLayout(pdf_widget)
        pdf_layout.setContentsMargins(0, 0, 0, 0)

        self.image_label = PdfLabel()
        self.image_label.main_window = self
        self.image_label.setAlignment(QtCore.Qt.AlignTop | QtCore.Qt.AlignLeft)
        self.image_label.setStyleSheet("QLabel { background: white; border: 1px solid #b8c2d0; border-radius: 12px; }")

        self.pdf_scroll = QtWidgets.QScrollArea()
        self.pdf_scroll.setWidgetResizable(False)
        self.pdf_scroll.setAlignment(QtCore.Qt.AlignCenter)
        self.pdf_scroll.setWidget(self.image_label)
        self.pdf_scroll.setStyleSheet("QScrollArea { border: 1px solid #b8c2d0; border-radius: 12px; background: #eef4fb; }")
        pdf_layout.addWidget(self.pdf_scroll)

        bottom_widget = QtWidgets.QWidget()
        bottom_layout = QtWidgets.QVBoxLayout(bottom_widget)
        bottom_layout.setContentsMargins(0, 0, 0, 0)
        bottom_layout.setSpacing(6)

        tip = QtWidgets.QLabel(
            "Blue = saved annotation points on current page | Magenta = selected annotation point | "
            "Purple dashed lines = page bookmarks preview | Red = latest click"
        )
        tip.setWordWrap(True)
        tip.setStyleSheet("QLabel { background: #fff3c9; color: #8a5b00; border-radius: 8px; padding: 5px 10px; font-family: 'Times New Roman'; font-size: 11pt; }")
        bottom_layout.addWidget(tip)

        self.tabs = QtWidgets.QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #b8c2d0; border-radius: 10px; background: white; }
            QTabBar::tab {
                font-family: 'Times New Roman'; font-size: 11pt; padding: 8px 16px;
                background: #dde8f6; border-top-left-radius: 8px; border-top-right-radius: 8px; margin-right: 4px;
            }
            QTabBar::tab:selected { background: #ffffff; font-weight: bold; }
        """)

        # Annotation tab
        ann_tab = QtWidgets.QWidget()
        ann_layout = QtWidgets.QVBoxLayout(ann_tab)
        ann_layout.setContentsMargins(6, 6, 6, 6)

        self.annotation_table = QtWidgets.QTableWidget(0, 5)
        self.annotation_table.setHorizontalHeaderLabels(["DOMAIN", "NAME", "PAGENO", "ANNOTATION", "ASSIGNED\nFIELD"])
        self.annotation_table.setSelectionBehavior(QtWidgets.QTableWidget.SelectRows)
        self.annotation_table.setSelectionMode(QtWidgets.QTableWidget.SingleSelection)
        self.annotation_table.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)
        self.annotation_table.setAlternatingRowColors(True)
        self.annotation_table.verticalHeader().setDefaultSectionSize(28)
        self.annotation_table.setWordWrap(True)
        self.annotation_table.setTextElideMode(QtCore.Qt.ElideNone)

        ann_header_view = self.annotation_table.horizontalHeader()
        ann_header_view.setDefaultAlignment(QtCore.Qt.AlignCenter)
        ann_header_view.setSectionResizeMode(0, QtWidgets.QHeaderView.Fixed)
        ann_header_view.setSectionResizeMode(1, QtWidgets.QHeaderView.Fixed)
        ann_header_view.setSectionResizeMode(2, QtWidgets.QHeaderView.Fixed)
        ann_header_view.setSectionResizeMode(3, QtWidgets.QHeaderView.Stretch)
        ann_header_view.setSectionResizeMode(4, QtWidgets.QHeaderView.Fixed)

        self.annotation_table.setColumnWidth(0, 110)
        self.annotation_table.setColumnWidth(1, 140)
        self.annotation_table.setColumnWidth(2, 80)
        self.annotation_table.setColumnWidth(4, 110)

        ann_layout.addWidget(self.annotation_table)

        # Bookmark tab
        bm_tab = QtWidgets.QWidget()
        bm_layout = QtWidgets.QVBoxLayout(bm_tab)
        bm_layout.setContentsMargins(6, 6, 6, 6)

        self.bookmark_table = QtWidgets.QTableWidget(0, 3)
        self.bookmark_table.setHorizontalHeaderLabels(["BOOKMARK TEXT", "LEVEL", "PAGE NO"])
        self.bookmark_table.setSelectionBehavior(QtWidgets.QTableWidget.SelectRows)
        self.bookmark_table.setSelectionMode(QtWidgets.QTableWidget.SingleSelection)
        self.bookmark_table.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)
        self.bookmark_table.setAlternatingRowColors(True)
        self.bookmark_table.verticalHeader().setDefaultSectionSize(28)

        bm_header_view = self.bookmark_table.horizontalHeader()
        bm_header_view.setDefaultAlignment(QtCore.Qt.AlignCenter)
        bm_header_view.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        bm_header_view.setSectionResizeMode(1, QtWidgets.QHeaderView.Fixed)
        bm_header_view.setSectionResizeMode(2, QtWidgets.QHeaderView.Fixed)

        self.bookmark_table.setColumnWidth(1, 80)
        self.bookmark_table.setColumnWidth(2, 90)

        bm_layout.addWidget(self.bookmark_table)

        self.tabs.addTab(ann_tab, "Annotations")
        self.tabs.addTab(bm_tab, "Bookmarks")

        bottom_layout.addWidget(self.tabs)

        self.splitter.addWidget(pdf_widget)
        self.splitter.addWidget(bottom_widget)
        layout.addWidget(self.splitter, 1)

        QtCore.QTimer.singleShot(0, self.init_splitter_sizes)

        # connections
        self.mode_combo.currentIndexChanged.connect(self.on_mode_changed)
        self.btn_open.clicked.connect(self.open_pdf)
        self.btn_prev.clicked.connect(self.prev_page)
        self.btn_next.clicked.connect(self.next_page)
        self.btn_add_bookmark.clicked.connect(self.add_bookmark)
        self.btn_delete.clicked.connect(self.delete_selected_row)
        self.btn_save_ann_csv.clicked.connect(self.export_annotation_csv)
        self.btn_load_ann_csv.clicked.connect(self.load_annotation_csv)
        self.btn_save_bm_csv.clicked.connect(self.export_bookmark_csv)
        self.btn_load_bm_csv.clicked.connect(self.load_bookmark_csv)
        self.btn_generate_pdf.clicked.connect(self.generate_final_output_pdf)

        self.annotation_table.itemSelectionChanged.connect(self.on_annotation_selection_changed)
        self.bookmark_table.itemSelectionChanged.connect(self.on_bookmark_selection_changed)

        self.on_mode_changed(0)

    def init_splitter_sizes(self):
        total = max(self.height() - 180, 700)
        self.splitter.setSizes([int(total * 0.78), int(total * 0.22)])

    # ------------------------------------------------------------------
    # Mode
    # ------------------------------------------------------------------
    def on_mode_changed(self, idx):
        self.active_mode = "annotation" if idx == 0 else "bookmark"

        is_annotation = self.active_mode == "annotation"

        self.btn_save_ann_csv.setVisible(is_annotation)
        self.btn_load_ann_csv.setVisible(is_annotation)

        self.btn_save_bm_csv.setVisible(not is_annotation)
        self.btn_load_bm_csv.setVisible(not is_annotation)
        self.btn_add_bookmark.setVisible(not is_annotation)

        self.tabs.setCurrentIndex(0 if is_annotation else 1)
        self.image_label.update()

    # ------------------------------------------------------------------
    # PDF open / render
    # ------------------------------------------------------------------
    def open_pdf(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Open PDF", "", "PDF Files (*.pdf)")
        if not file_path:
            return
        try:
            self.doc = fitz.open(file_path)
            self.open_pdf_path = file_path
            self.current_page_index = 0
            self.entries = []
            self.bookmarks = []
            self.selected_entry_index = -1
            self.selected_bookmark_index = -1
            self.image_label.last_click_point = None
            self.last_click_pdf_x = None
            self.last_click_pdf_y = None
            self.last_click_pageh = None
            self.refresh_annotation_table()
            self.refresh_bookmark_table()
            self.render_page()
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Unable to open PDF:\n{e}")

    def render_page(self):
        if not self.doc:
            return
        page = self.doc[self.current_page_index]
        self.page_rect = page.rect

        mat = fitz.Matrix(self.zoom, self.zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        qimg = QtGui.QImage(pix.samples, pix.width, pix.height, pix.stride, QtGui.QImage.Format_RGB888).copy()

        self.page_pixmap = QtGui.QPixmap.fromImage(qimg)
        self.image_label.setPixmap(self.page_pixmap)
        self.image_label.resize(self.page_pixmap.size())
        self.image_label.update()

        self.page_info.setText(f"Page: {self.current_page_index + 1} / {len(self.doc)}")

    def prev_page(self):
        if self.doc and self.current_page_index > 0:
            self.current_page_index -= 1
            self.image_label.last_click_point = None
            self.render_page()

    def next_page(self):
        if self.doc and self.current_page_index < len(self.doc) - 1:
            self.current_page_index += 1
            self.image_label.last_click_point = None
            self.render_page()

    # ------------------------------------------------------------------
    # Click handling
    # ------------------------------------------------------------------
    def store_last_click(self, point: QtCore.QPoint):
        if not self.doc or not self.page_pixmap or not self.page_rect:
            return

        self.last_click_pdf_x = max(0, min(point.x() / self.zoom, self.page_rect.width))
        self.last_click_pdf_y = max(0, min(point.y() / self.zoom, self.page_rect.height))
        self.last_click_pageh = float(self.page_rect.height)

    def capture_annotation_point(self, point: QtCore.QPoint):
        if not self.doc or not self.page_pixmap or not self.page_rect:
            return

        x1 = max(0, min(point.x() / self.zoom, self.page_rect.width))
        y1 = max(0, min(point.y() / self.zoom, self.page_rect.height))
        pageh = float(self.page_rect.height)

        dlg = AnnotationDialog(self)
        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return

        vals = dlg.get_values()
        entry = AnnotationEntry(
            domain=vals["domain"],
            name=vals["name"],
            pageno=self.current_page_index + 1,
            annotation=vals["annotation"],
            assignedfield=vals["assignedfield"],
            x1=round(x1, 6),
            y1=round(y1, 6),
            pageh=round(pageh, 6),
            is_domain_annotation=vals["is_domain_annotation"],
            is_assigned_field=vals["is_assigned_field"],
            is_not_submitted=vals["is_not_submitted"]
        )

        self.entries.append(entry)
        self.selected_entry_index = len(self.entries) - 1
        self.refresh_annotation_table()
        self.select_annotation_row_silent(self.selected_entry_index)
        self.image_label.update()

    # ------------------------------------------------------------------
    # Add bookmark
    # ------------------------------------------------------------------
    def add_bookmark(self):
        if not self.doc:
            QtWidgets.QMessageBox.information(self, "Info", "Open a PDF first.")
            return

        dlg = BookmarkDialog(self.current_page_index + 1, self)
        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return

        vals = dlg.get_values()
        bm = BookmarkEntry(
            title=vals["title"],
            level=vals["level"],
            pageno=vals["pageno"]
        )
        self.bookmarks.append(bm)
        self.selected_bookmark_index = len(self.bookmarks) - 1
        self.refresh_bookmark_table()
        self.select_bookmark_row_silent(self.selected_bookmark_index)
        self.image_label.update()

    # ------------------------------------------------------------------
    # Annotation table
    # ------------------------------------------------------------------
    def refresh_annotation_table(self):
        self._suppress_annotation_selection_signal = True
        self.annotation_table.setRowCount(0)
        for row_idx, e in enumerate(self.entries):
            self.annotation_table.insertRow(row_idx)
            vals = [e.domain, e.name, str(e.pageno), e.annotation, e.assignedfield]
            for col_idx, val in enumerate(vals):
                item = QtWidgets.QTableWidgetItem(val)
                if col_idx in (0, 1, 2, 4):
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                elif col_idx == 3:
                    item.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
                self.annotation_table.setItem(row_idx, col_idx, item)
        self.annotation_table.resizeRowsToContents()
        self._suppress_annotation_selection_signal = False

    def select_annotation_row_silent(self, row_idx: int):
        if row_idx < 0 or row_idx >= self.annotation_table.rowCount():
            return
        self._suppress_annotation_selection_signal = True
        self.annotation_table.selectRow(row_idx)
        self._suppress_annotation_selection_signal = False

    def refresh_selection_only(self):
        if self.selected_entry_index >= 0:
            self.select_annotation_row_silent(self.selected_entry_index)

    def on_annotation_selection_changed(self):
        if self._suppress_annotation_selection_signal:
            return

        row = self.annotation_table.currentRow()
        if 0 <= row < len(self.entries):
            self.selected_entry_index = row
            entry = self.entries[row]
            target_page_index = entry.pageno - 1
            if self.doc and target_page_index != self.current_page_index:
                self.current_page_index = target_page_index
                self.render_page()
        else:
            self.selected_entry_index = -1
        self.image_label.update()

    # ------------------------------------------------------------------
    # Bookmark table
    # ------------------------------------------------------------------
    def refresh_bookmark_table(self):
        self._suppress_bookmark_selection_signal = True
        self.bookmark_table.setRowCount(0)
        for row_idx, b in enumerate(self.bookmarks):
            self.bookmark_table.insertRow(row_idx)
            vals = [b.title, str(b.level), str(b.pageno)]
            for col_idx, val in enumerate(vals):
                item = QtWidgets.QTableWidgetItem(val)
                if col_idx in (1, 2):
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                else:
                    item.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                self.bookmark_table.setItem(row_idx, col_idx, item)
        self.bookmark_table.resizeRowsToContents()
        self._suppress_bookmark_selection_signal = False

    def select_bookmark_row_silent(self, row_idx: int):
        if row_idx < 0 or row_idx >= self.bookmark_table.rowCount():
            return
        self._suppress_bookmark_selection_signal = True
        self.bookmark_table.selectRow(row_idx)
        self._suppress_bookmark_selection_signal = False

    def on_bookmark_selection_changed(self):
        if self._suppress_bookmark_selection_signal:
            return

        row = self.bookmark_table.currentRow()
        if 0 <= row < len(self.bookmarks):
            self.selected_bookmark_index = row
            bm = self.bookmarks[row]
            target_page_index = bm.pageno - 1
            if self.doc and target_page_index != self.current_page_index:
                self.current_page_index = target_page_index
                self.render_page()
        else:
            self.selected_bookmark_index = -1
        self.image_label.update()

    # ------------------------------------------------------------------
    # Delete selected
    # ------------------------------------------------------------------
    def delete_selected_row(self):
        if self.active_mode == "annotation":
            row = self.annotation_table.currentRow()
            if row < 0:
                QtWidgets.QMessageBox.information(self, "Info", "Select an annotation row to delete.")
                return
            del self.entries[row]
            self.selected_entry_index = min(row, len(self.entries) - 1) if self.entries else -1
            self.refresh_annotation_table()
            if self.selected_entry_index >= 0:
                self.select_annotation_row_silent(self.selected_entry_index)
        else:
            row = self.bookmark_table.currentRow()
            if row < 0:
                QtWidgets.QMessageBox.information(self, "Info", "Select a bookmark row to delete.")
                return
            del self.bookmarks[row]
            self.selected_bookmark_index = min(row, len(self.bookmarks) - 1) if self.bookmarks else -1
            self.refresh_bookmark_table()
            if self.selected_bookmark_index >= 0:
                self.select_bookmark_row_silent(self.selected_bookmark_index)

        self.image_label.update()

    # ------------------------------------------------------------------
    # CSV export / load
    # ------------------------------------------------------------------
    def export_annotation_csv(self):
        if not self.entries:
            QtWidgets.QMessageBox.information(self, "No Data", "No annotations captured.")
            return

        file_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save Annotation CSV", "annotate_input.csv", "CSV Files (*.csv)"
        )
        if not file_path:
            return

        try:
            with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
                writer.writerow(["DOMAIN", "NAME", "PAGENO", "ANNOTATION", "ASSIGNEDFIELD", "X1", "Y1", "PAGEH"])
                for e in self.entries:
                    writer.writerow([
                        e.domain, e.name, e.pageno, e.annotation, e.assignedfield,
                        f"{e.x1:.6f}", f"{e.y1:.6f}", f"{e.pageh:.6f}"
                    ])
            self.current_annotation_csv_path = file_path
            QtWidgets.QMessageBox.information(self, "Success", f"Annotation CSV exported:\n{file_path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to export annotation CSV:\n{e}")

    def load_annotation_csv(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Load Annotation CSV",
            self.current_annotation_csv_path if self.current_annotation_csv_path else "",
            "CSV Files (*.csv)"
        )
        if not file_path:
            return

        try:
            loaded = []
            with open(file_path, "r", encoding="utf-8-sig", newline="") as f:
                reader = csv.DictReader(f)
                for r in reader:
                    domain = (r.get("DOMAIN") or "").strip()
                    name = (r.get("NAME") or "").strip()
                    annotation = (r.get("ANNOTATION") or "").strip()
                    assignedfield = (r.get("ASSIGNEDFIELD") or "").strip()
                    is_not_submitted = annotation.strip().upper() == "[NOT SUBMITTED]"
                    is_domain_annotation = (name == "") and not is_not_submitted
                    is_assigned_field = assignedfield.upper() == "Y"

                    loaded.append(
                        AnnotationEntry(
                            domain=domain,
                            name=name,
                            pageno=int(float(r.get("PAGENO", 0) or 0)),
                            annotation=annotation,
                            assignedfield=assignedfield,
                            x1=float(r.get("X1", 0) or 0),
                            y1=float(r.get("Y1", 0) or 0),
                            pageh=float(r.get("PAGEH", 0) or 0),
                            is_domain_annotation=is_domain_annotation,
                            is_assigned_field=is_assigned_field,
                            is_not_submitted=is_not_submitted
                        )
                    )

            self.entries = loaded
            self.current_annotation_csv_path = file_path
            self.selected_entry_index = -1
            self.refresh_annotation_table()
            self.image_label.update()
            QtWidgets.QMessageBox.information(self, "Loaded", f"Loaded annotation CSV:\n{file_path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load annotation CSV:\n{e}")

    def export_bookmark_csv(self):
        if not self.bookmarks:
            QtWidgets.QMessageBox.information(self, "No Data", "No bookmarks captured.")
            return

        file_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save Bookmark CSV", "bookmarks.csv", "CSV Files (*.csv)"
        )
        if not file_path:
            return

        try:
            with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
                writer.writerow(["TITLE", "LEVEL", "PAGENO"])
                for b in self.bookmarks:
                    writer.writerow([b.title, b.level, b.pageno])
            self.current_bookmark_csv_path = file_path
            QtWidgets.QMessageBox.information(self, "Success", f"Bookmark CSV exported:\n{file_path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to export bookmark CSV:\n{e}")

    def load_bookmark_csv(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Load Bookmark CSV",
            self.current_bookmark_csv_path if self.current_bookmark_csv_path else "",
            "CSV Files (*.csv)"
        )
        if not file_path:
            return

        try:
            loaded = []
            with open(file_path, "r", encoding="utf-8-sig", newline="") as f:
                reader = csv.DictReader(f)
                for r in reader:
                    loaded.append(
                        BookmarkEntry(
                            title=(r.get("TITLE") or "").strip(),
                            level=int(float(r.get("LEVEL", 1) or 1)),
                            pageno=int(float(r.get("PAGENO", 1) or 1))
                        )
                    )

            self.bookmarks = loaded
            self.current_bookmark_csv_path = file_path
            self.selected_bookmark_index = -1
            self.refresh_bookmark_table()
            self.image_label.update()
            QtWidgets.QMessageBox.information(self, "Loaded", f"Loaded bookmark CSV:\n{file_path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load bookmark CSV:\n{e}")

    # ------------------------------------------------------------------
    # Final PDF generation
    # ------------------------------------------------------------------
    def generate_final_output_pdf(self):
        if not self.open_pdf_path or not os.path.exists(self.open_pdf_path):
            QtWidgets.QMessageBox.warning(self, "Generate PDF", "Please open the source PDF first.")
            return

        if not self.entries and not self.bookmarks:
            QtWidgets.QMessageBox.warning(self, "Generate PDF", "No annotations or bookmarks available.")
            return

        default_out = sanitize_output_pdf_path(self.open_pdf_path)
        output_pdf, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save Final Output PDF", default_out, "PDF Files (*.pdf)"
        )
        if not output_pdf:
            return

        try:
            QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
            doc = fitz.open(self.open_pdf_path)

            # Write annotations
            if self.entries:
                color_map = get_domain_color_map(self.entries)

                for e in self.entries:
                    if e.pageno < 1 or e.pageno > len(doc):
                        continue

                    page = doc[e.pageno - 1]
                    layout = compute_entry_layout(e, color_map)
                    rect = rect_from_top_origin(e.x1, e.y1, layout["box_w"], layout["box_h"], e.pageh)

                    if rect.x1 > page.rect.width:
                        shift = rect.x1 - page.rect.width
                        rect = fitz.Rect(rect.x0 - shift, rect.y0, rect.x1 - shift, rect.y1)

                    if rect.y1 > page.rect.height:
                        shift = rect.y1 - page.rect.height
                        rect = fitz.Rect(rect.x0, rect.y0 - shift, rect.x1, rect.y1 - shift)

                    if rect.y0 < 0:
                        shift = -rect.y0
                        rect = fitz.Rect(rect.x0, rect.y0 + shift, rect.x1, rect.y1 + shift)

                    draw_box_and_text_pdf(
                        page=page,
                        rect=rect,
                        text_lines=layout["lines"],
                        fill_color=layout["fill"],
                        bold=layout["bold"],
                        dashed=layout["dashed"],
                        font_size=layout["font_size"]
                    )

            # Write bookmarks
            if self.bookmarks:
                toc = []
                for bm in self.bookmarks:
                    lvl = max(1, int(bm.level))
                    pg = max(1, min(int(bm.pageno), len(doc)))
                    title = bm.title.strip()
                    if title:
                        toc.append([lvl, title, pg])

                if toc:
                    doc.set_toc(toc)

            doc.save(output_pdf, garbage=4, deflate=True)
            doc.close()

            what_written = []
            if self.entries:
                what_written.append(f"{len(self.entries)} annotation(s)")
            if self.bookmarks:
                what_written.append(f"{len(self.bookmarks)} bookmark(s)")

            QtWidgets.QMessageBox.information(
                self,
                "Final Output PDF Generated",
                f"Final PDF created successfully:\n{output_pdf}\n\nEmbedded: {', '.join(what_written)}"
            )
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to generate final PDF:\n{e}")
        finally:
            QtWidgets.QApplication.restoreOverrideCursor()


def main():
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")
    win = AnnotatorApp()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()