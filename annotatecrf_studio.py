# ====================================================================================================
# Script Name    : annotatecrf_studio.py
#
# Description    :
#                  Desktop GUI utility for capturing annotation positions from a PDF and generating
#                  final visible comment boxes directly into a PDF without XFDF / Adobe dependency.
#
#                  Workflow:
#                    - Open PDF first
#                    - Annotation mode:
#                        * Click on PDF -> enter details -> preview box appears immediately
#                        * Drag an existing preview box to move it
#                        * Manage annotation rows + export/load CSV
#                    - Bookmark mode:
#                        * Click on PDF -> enter bookmark details for current page
#                        * Manage bookmark rows + export/load CSV
#                    - Review mode:
#                        * View annotation and bookmark tables side by side
#                        * Generate Final Output PDF
#
#                  Annotation CSV columns used internally:
#                    DOMAIN,NAME,PAGENO,ANNOTATION,ASSIGNEDFIELD,X1,Y1,PAGEH
#
#                  Bookmark CSV columns used internally:
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
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

APP_ICON_CANDIDATES = [
    os.path.join(SCRIPT_DIR, "annotator_icon.ico"),
    os.path.join(SCRIPT_DIR, "annotator_icon.png"),
]

DEFAULT_ZOOM = 1.30
DOMAIN_COLORS = [
    (0.75, 1.00, 1.00),   # 1st domain on page = cyan
    (0.59, 1.00, 0.59),   # 2nd domain on page = green
    (1.00, 0.75, 0.61),   # 3rd domain on page = peach
    (1.00, 0.80, 0.55),   # 4th domain on page = orange
    (0.86, 0.82, 1.00),   # 5th domain on page = lavender
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

MAX_WRAP_WIDTH = 520.0
RIGHT_PAGE_MARGIN = 12.0
MIN_BOX_WIDTH = 90.0
LONG_TEXT_THRESHOLD = 22
LONG_TEXT_MIN_WIDTH = 220.0

BOOKMARK_PREVIEW_COLOR = QtGui.QColor("#7c3aed")
BOOKMARK_PREVIEW_TEXT_COLOR = QtGui.QColor("#5b21b6")

ANNOTATION_CSV_COLUMNS = ["DOMAIN", "NAME", "PAGENO", "ANNOTATION", "ASSIGNEDFIELD", "X1", "Y1", "PAGEH"]
BOOKMARK_CSV_COLUMNS = ["TITLE", "LEVEL", "PAGENO"]

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


def get_page_domain_color_map(entries, pageno: int) -> dict:
    ordered_domains = OrderedDict()
    for e in entries:
        if e.pageno != pageno:
            continue
        dom = (e.domain or "").strip().upper()
        if not dom or dom == "NOTSUB":
            continue
        if dom not in ordered_domains:
            ordered_domains[dom] = None

    color_map = {}
    for i, dom in enumerate(ordered_domains.keys()):
        if i < len(DOMAIN_COLORS):
            color_map[dom] = DOMAIN_COLORS[i]
        else:
            color_map[dom] = DEFAULT_OTHER_FILL
    return color_map


def compute_entry_layout(entry, color_map, page_width=None):
    domain_key = (entry.domain or "").strip().upper()

    fill = NOTSUB_FILL if entry.is_not_submitted else color_map.get(domain_key, DEFAULT_OTHER_FILL)
    dashed = bool(entry.is_assigned_field)
    bold = bool(entry.is_domain_annotation)
    font_size = DOMAIN_FONT_SIZE if bold else BASE_FONT_SIZE
    scale = DOMAIN_FONT_SIZE / BASE_FONT_SIZE if bold else 1.0
    base_h = BOX_HEIGHT_DOMAIN if bold else BOX_HEIGHT_NORMAL

    text = (entry.annotation or "").strip()
    raw_w = estimate_text_width(text, scale=scale)

    allowed_width = MAX_WRAP_WIDTH
    if page_width is not None:
        available = max(MIN_BOX_WIDTH, page_width - entry.x1 - RIGHT_PAGE_MARGIN)
        allowed_width = min(MAX_WRAP_WIDTH, available)

    preferred_width = raw_w

    long_text = (
        len(text) >= LONG_TEXT_THRESHOLD
        or " when " in f" {text.lower()} "
        or "/" in text
    )
    if long_text:
        preferred_width = max(preferred_width, LONG_TEXT_MIN_WIDTH)

    box_w = min(allowed_width, preferred_width)
    box_w = max(MIN_BOX_WIDTH, box_w)

    lines = wrap_text_by_width(text, box_w, scale=scale)
    actual_w = max(estimate_text_width(line, scale=scale) for line in lines) if lines else raw_w
    box_w = min(allowed_width, max(box_w, actual_w))
    box_w = max(MIN_BOX_WIDTH, box_w)

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
    base, _ = os.path.splitext(src_pdf)
    return base + "_final.pdf"


def bool_from_entry(domain, name, annotation, assignedfield):
    d = (domain or "").strip().upper()
    n = (name or "").strip().upper()
    a = (annotation or "").strip().upper()
    assigned = (assignedfield or "").strip().upper()

    is_not_submitted = (d == "NOTSUB") or (n == "NOTSUB") or (a == "[NOT SUBMITTED]")
    is_domain_annotation = (not is_not_submitted) and not str(name).strip()
    is_assigned_field = assigned == "Y"
    return is_domain_annotation, is_assigned_field, is_not_submitted


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
        self.setModal(True)
        self.resize(760, 330)
        self.setMinimumSize(760, 330)
        self.setSizeGripEnabled(False)

        self.setStyleSheet("""
            QDialog { background-color: #f4f8ff; border-radius: 14px; }
            QLabel { font-family: 'Times New Roman'; font-size: 12pt; color: #1c2e4a; }
            QLineEdit, QSpinBox {
                background: #ffffff; border: 1px solid #a8bfdc; border-radius: 8px;
                padding: 6px 8px; font-family: 'Times New Roman'; font-size: 12pt; color: #1a1a1a;
                min-height: 22px;
            }
            QLineEdit:focus, QSpinBox:focus {
                border: 2px solid #8b5cf6; background: #fdfefe;
            }
            QPushButton {
                font-family: 'Times New Roman'; font-size: 12pt; font-weight: bold;
                border-radius: 10px; padding: 8px 18px; min-width: 140px; min-height: 38px;
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
        self.title_edit.setMinimumWidth(420)

        self.level_spin = QtWidgets.QSpinBox()
        self.level_spin.setRange(1, 9)
        self.level_spin.setValue(1)
        self.level_spin.setButtonSymbols(QtWidgets.QAbstractSpinBox.UpDownArrows)

        self.page_spin = QtWidgets.QSpinBox()
        self.page_spin.setRange(1, 999999)
        self.page_spin.setValue(page_no)
        self.page_spin.setButtonSymbols(QtWidgets.QAbstractSpinBox.UpDownArrows)

        form = QtWidgets.QFormLayout()
        form.setLabelAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        form.setFormAlignment(QtCore.Qt.AlignTop)
        form.setFieldGrowthPolicy(QtWidgets.QFormLayout.AllNonFixedFieldsGrow)
        form.setHorizontalSpacing(18)
        form.setVerticalSpacing(14)
        form.addRow("Bookmark Text", self.title_edit)
        form.addRow("Level", self.level_spin)
        form.addRow("Page No", self.page_spin)

        note = QtWidgets.QLabel("Bookmark will be added to the PDF outline for the selected page.")
        note.setWordWrap(True)
        note.setStyleSheet("QLabel { color: #556b84; font-size: 11pt; font-style: italic; }")

        self.ok_btn = QtWidgets.QPushButton("OK")
        self.cancel_btn = QtWidgets.QPushButton("Cancel")
        self.ok_btn.clicked.connect(self.validate_and_accept)
        self.cancel_btn.clicked.connect(self.reject)
        self.ok_btn.setStyleSheet("QPushButton { background-color: #55c16d; color: white; } QPushButton:hover { background-color: #73d789; }")
        self.cancel_btn.setStyleSheet("QPushButton { background-color: #f16a6a; color: white; } QPushButton:hover { background-color: #f48f8f; }")

        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(self.ok_btn)
        btn_row.addSpacing(12)
        btn_row.addWidget(self.cancel_btn)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)
        layout.setSizeConstraint(QtWidgets.QLayout.SetFixedSize)
        layout.addWidget(title_hdr)
        layout.addSpacing(6)
        layout.addLayout(form)
        layout.addWidget(note)
        layout.addSpacing(8)
        layout.addLayout(btn_row)

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
        page_width = self.main_window.page_rect.width if self.main_window.page_rect else None
        layout = compute_entry_layout(entry, color_map, page_width=page_width)
        rect_pdf = rect_from_top_origin(entry.x1, entry.y1, layout["box_w"], layout["box_h"], entry.pageh)
        zoom = self.main_window.zoom
        return QtCore.QRectF(
            rect_pdf.x0 * zoom,
            rect_pdf.y0 * zoom,
            rect_pdf.width * zoom,
            rect_pdf.height * zoom
        )

    def mousePressEvent(self, event):
        if not self.main_window or not self.main_window.pdf_loaded or self.main_window.page_pixmap is None:
            return

        if event.button() != QtCore.Qt.LeftButton:
            return

        current_page = self.main_window.current_page_index + 1
        color_map = get_page_domain_color_map(self.main_window.entries, current_page)

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
        elif self.main_window.active_mode == "bookmark":
            self.main_window.capture_bookmark_point()

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

        if not self.main_window or self.main_window.page_rect is None:
            return

        painter = QtGui.QPainter(self)
        zoom = self.main_window.zoom
        current_page = self.main_window.current_page_index + 1
        color_map = get_page_domain_color_map(self.main_window.entries, current_page)
        page_width = self.main_window.page_rect.width if self.main_window.page_rect else None

        # Annotation previews
        for idx, entry in enumerate(self.main_window.entries):
            if entry.pageno != current_page:
                continue

            layout = compute_entry_layout(entry, color_map, page_width=page_width)
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
            if idx == self.main_window.selected_entry_index:
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
            painter.drawLine(0, sy, min(self.width(), 320), sy)

            text_pen = QtGui.QPen(BOOKMARK_PREVIEW_TEXT_COLOR, 1)
            painter.setPen(text_pen)
            f = QtGui.QFont("Arial", 9)
            is_selected = (
                0 <= self.main_window.selected_bookmark_index < len(self.main_window.bookmarks)
                and self.main_window.bookmarks[self.main_window.selected_bookmark_index] == bm
            )
            f.setBold(is_selected)
            painter.setFont(f)

            offset_x = 8 + (max(1, bm.level) - 1) * 12
            painter.drawText(QtCore.QPointF(offset_x, sy - 4), f"BM L{bm.level}: {bm.title[:45]}")
            painter.setPen(bookmark_pen)

        # Last click crosshair
        if self.last_click_point and self.main_window.active_mode in ("annotation", "bookmark"):
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
        self.setWindowTitle("Annotated CRF Studio")
        self.resize(1280, 900)
        self.setMinimumSize(1100, 760)
        self.setStyleSheet("background-color: #f3f7fd;")

        self.apply_app_icon()

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

        self.last_click_pdf_x = None
        self.last_click_pdf_y = None
        self.last_click_pageh = None

        self.active_mode = "annotation"
        self.pdf_loaded = False
        self.has_annotation = False
        self.has_bookmark = False

        self._suppress_annotation_selection_signal = False
        self._suppress_bookmark_selection_signal = False

        self.build_ui()

    def apply_app_icon(self):
        for icon_path in APP_ICON_CANDIDATES:
            if os.path.exists(icon_path):
                icon = QtGui.QIcon(icon_path)
                self.setWindowIcon(icon)
                app = QtWidgets.QApplication.instance()
                if app is not None:
                    app.setWindowIcon(icon)
                break

    def build_ui(self):
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(8)

        header = QtWidgets.QLabel("Annotated CRF Studio")
        header.setAlignment(QtCore.Qt.AlignCenter)
        header.setStyleSheet("""
            QLabel {
                color: black; background: #bfe9f7; padding: 10px 0 8px 0; border-radius: 16px;
                font-family: 'Times New Roman'; font-size: 20pt; font-weight: bold;
            }
        """)
        layout.addWidget(header)

        subtitle = QtWidgets.QLabel(
            "One-Click Annotation & Bookmarking for MSG 2.0–Compliant aCRFs"
        )
        subtitle.setAlignment(QtCore.Qt.AlignCenter)
        subtitle.setWordWrap(True)
        subtitle.setStyleSheet("QLabel { color: #21466d; padding: 0 0 4px 0; font-family: 'Times New Roman'; font-size: 12pt; }")
        layout.addWidget(subtitle)

        contact_note = QtWidgets.QLabel(
            "For queries / suggestions / issues: Manivannan.Mathialagan@veristat.com"
        )
        contact_note.setAlignment(QtCore.Qt.AlignCenter)
        contact_note.setWordWrap(True)
        contact_note.setStyleSheet("""
            QLabel {
                background: #fff3c9;
                color: #184a78;
                border: 1px solid #e6d27d;
                border-radius: 10px;
                padding: 6px 10px;
                font-family: 'Times New Roman';
                font-size: 11pt;
                font-style: italic;
            }
        """)
        layout.addWidget(contact_note)

        btn_base = """
            QPushButton {
                font-family: 'Times New Roman'; font-weight: bold; font-size: 11pt;
                border-radius: 10px; padding: 6px 14px; min-height: 34px;
            }
            QPushButton:disabled {
                background-color: #d9d9d9;
                color: #7a7a7a;
            }
        """

        def mkbtn(text, bg, hov, fg="white"):
            b = QtWidgets.QPushButton(text)
            b.setStyleSheet(btn_base + f"""
                QPushButton {{ background-color: {bg}; color: {fg}; }}
                QPushButton:hover:!disabled {{ background-color: {hov}; }}
            """)
            return b

        # Top controls
        top_row = QtWidgets.QHBoxLayout()
        top_row.setSpacing(8)

        self.btn_open = mkbtn("Open PDF", "#3b82f6", "#5c9cff")
        self.btn_prev = mkbtn("Previous Page", "#b9b9b9", "#d0d0d0", "#333")
        self.btn_next = mkbtn("Next Page", "#ad7dfc", "#c7a2fd")

        self.btn_prev.setEnabled(False)
        self.btn_next.setEnabled(False)

        self.page_info = QtWidgets.QLabel("Page: -")
        self.page_info.setStyleSheet("""
            QLabel {
                background: #e6f2fb; color: #12608d; border-radius: 8px; padding: 5px 10px;
                font-family: 'Times New Roman'; font-size: 11pt; font-weight: bold;
            }
        """)

        top_row.addWidget(self.btn_open)
        top_row.addWidget(self.btn_prev)
        top_row.addWidget(self.btn_next)
        top_row.addStretch()
        top_row.addWidget(self.page_info)
        layout.addLayout(top_row)

        # Mode row
        mode_row = QtWidgets.QHBoxLayout()
        mode_row.setSpacing(8)

        self.btn_annotation = mkbtn("Annotation", "#2563eb", "#4a7df2")
        self.btn_bookmark = mkbtn("Bookmark", "#7c3aed", "#9b63f0")
        self.btn_review = mkbtn("Review", "#059669", "#21b58a")

        self.btn_annotation.setEnabled(False)
        self.btn_bookmark.setEnabled(False)
        self.btn_review.setEnabled(False)

        mode_row.addWidget(self.btn_annotation)
        mode_row.addWidget(self.btn_bookmark)
        mode_row.addWidget(self.btn_review)
        layout.addLayout(mode_row)

        # Main stack:
        # 0 = normal view (PDF + mode tools below)
        # 1 = review only (no PDF)
        self.main_stack = QtWidgets.QStackedWidget()
        layout.addWidget(self.main_stack, 1)

        # ------------------------------------------------------------------
        # PAGE 0: NORMAL VIEW
        # ------------------------------------------------------------------
        normal_page = QtWidgets.QWidget()
        normal_layout = QtWidgets.QVBoxLayout(normal_page)
        normal_layout.setContentsMargins(0, 0, 0, 0)
        normal_layout.setSpacing(8)

        self.splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.setHandleWidth(10)

        # PDF block
        self.pdf_widget = QtWidgets.QWidget()
        pdf_layout = QtWidgets.QVBoxLayout(self.pdf_widget)
        pdf_layout.setContentsMargins(0, 0, 0, 0)
        pdf_layout.setSpacing(4)

        self.tip = QtWidgets.QLabel(
            "Click anywhere on the PDF to add annotation or bookmark, based on the selected mode."
        )
        self.tip.setWordWrap(True)
        self.tip.setStyleSheet("QLabel { background: #eef6ff; color: #35516e; border-radius: 8px; padding: 5px 10px; font-family: 'Times New Roman'; font-size: 11pt; }")
        pdf_layout.addWidget(self.tip)

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

        self.splitter.addWidget(self.pdf_widget)

        # Bottom stack
        self.bottom_stack = QtWidgets.QStackedWidget()

        # ==========================
        # Annotation page
        # ==========================
        ann_page = QtWidgets.QWidget()
        ann_layout = QtWidgets.QVBoxLayout(ann_page)
        ann_layout.setContentsMargins(0, 0, 0, 0)
        ann_layout.setSpacing(6)

        ann_hdr = QtWidgets.QLabel("Annotations")
        ann_hdr.setAlignment(QtCore.Qt.AlignCenter)
        ann_hdr.setStyleSheet("QLabel { background: #dbeafe; color: #1e3a8a; border-radius: 8px; padding: 5px; font-weight: bold; }")
        ann_layout.addWidget(ann_hdr)

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

        ann_btn_row = QtWidgets.QHBoxLayout()
        ann_btn_row.setSpacing(8)

        self.btn_export_ann = mkbtn("Export Annotation CSV", "#2563eb", "#4a7df2")
        self.btn_load_ann = mkbtn("Load Annotation CSV", "#7c3aed", "#9b63f0")
        self.btn_delete_ann = mkbtn("Delete Selected Annotation", "#ef4444", "#f87171")

        ann_btn_row.addWidget(self.btn_export_ann)
        ann_btn_row.addWidget(self.btn_load_ann)
        ann_btn_row.addStretch()
        ann_btn_row.addWidget(self.btn_delete_ann)

        ann_layout.addLayout(ann_btn_row)
        self.bottom_stack.addWidget(ann_page)

        # ==========================
        # Bookmark page
        # ==========================
        bm_page = QtWidgets.QWidget()
        bm_layout = QtWidgets.QVBoxLayout(bm_page)
        bm_layout.setContentsMargins(0, 0, 0, 0)
        bm_layout.setSpacing(6)

        bm_hdr = QtWidgets.QLabel("Bookmarks")
        bm_hdr.setAlignment(QtCore.Qt.AlignCenter)
        bm_hdr.setStyleSheet("QLabel { background: #ede9fe; color: #5b21b6; border-radius: 8px; padding: 5px; font-weight: bold; }")
        bm_layout.addWidget(bm_hdr)

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

        bm_btn_row = QtWidgets.QHBoxLayout()
        bm_btn_row.setSpacing(8)

        self.btn_export_bm = mkbtn("Export Bookmark CSV", "#2563eb", "#4a7df2")
        self.btn_load_bm = mkbtn("Load Bookmark CSV", "#7c3aed", "#9b63f0")
        self.btn_delete_bm = mkbtn("Delete Selected Bookmark", "#ef4444", "#f87171")

        bm_btn_row.addWidget(self.btn_export_bm)
        bm_btn_row.addWidget(self.btn_load_bm)
        bm_btn_row.addStretch()
        bm_btn_row.addWidget(self.btn_delete_bm)

        bm_layout.addLayout(bm_btn_row)
        self.bottom_stack.addWidget(bm_page)

        self.splitter.addWidget(self.bottom_stack)
        normal_layout.addWidget(self.splitter)
        self.main_stack.addWidget(normal_page)

        # ------------------------------------------------------------------
        # PAGE 1: REVIEW ONLY (NO PDF)
        # ------------------------------------------------------------------
        review_page = QtWidgets.QWidget()
        review_layout = QtWidgets.QVBoxLayout(review_page)
        review_layout.setContentsMargins(0, 0, 0, 0)
        review_layout.setSpacing(6)

        review_hint = QtWidgets.QLabel("Review mode active. Check both tables below and generate the final PDF.")
        review_hint.setAlignment(QtCore.Qt.AlignCenter)
        review_hint.setStyleSheet("QLabel { color: #27496d; font-family: 'Times New Roman'; font-size: 11pt; padding: 2px 0; }")
        review_layout.addWidget(review_hint)

        review_splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        review_splitter.setChildrenCollapsible(False)
        review_splitter.setHandleWidth(8)

        ann_panel = QtWidgets.QWidget()
        ann_panel_layout = QtWidgets.QVBoxLayout(ann_panel)
        ann_panel_layout.setContentsMargins(0, 0, 0, 0)

        ann_review_hdr = QtWidgets.QLabel("Annotations")
        ann_review_hdr.setAlignment(QtCore.Qt.AlignCenter)
        ann_review_hdr.setStyleSheet("QLabel { background: #dbeafe; color: #1e3a8a; border-radius: 8px; padding: 5px; font-weight: bold; }")
        ann_panel_layout.addWidget(ann_review_hdr)

        self.review_annotation_table = QtWidgets.QTableWidget(0, 5)
        self.review_annotation_table.setHorizontalHeaderLabels(["DOMAIN", "NAME", "PAGENO", "ANNOTATION", "ASSIGNED\nFIELD"])
        self.review_annotation_table.setSelectionBehavior(QtWidgets.QTableWidget.SelectRows)
        self.review_annotation_table.setSelectionMode(QtWidgets.QTableWidget.SingleSelection)
        self.review_annotation_table.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)
        self.review_annotation_table.setAlternatingRowColors(True)
        self.review_annotation_table.verticalHeader().setDefaultSectionSize(28)
        self.review_annotation_table.setWordWrap(True)
        self.review_annotation_table.setTextElideMode(QtCore.Qt.ElideNone)

        review_ann_header = self.review_annotation_table.horizontalHeader()
        review_ann_header.setDefaultAlignment(QtCore.Qt.AlignCenter)
        review_ann_header.setSectionResizeMode(0, QtWidgets.QHeaderView.Fixed)
        review_ann_header.setSectionResizeMode(1, QtWidgets.QHeaderView.Fixed)
        review_ann_header.setSectionResizeMode(2, QtWidgets.QHeaderView.Fixed)
        review_ann_header.setSectionResizeMode(3, QtWidgets.QHeaderView.Stretch)
        review_ann_header.setSectionResizeMode(4, QtWidgets.QHeaderView.Fixed)
        self.review_annotation_table.setColumnWidth(0, 110)
        self.review_annotation_table.setColumnWidth(1, 140)
        self.review_annotation_table.setColumnWidth(2, 80)
        self.review_annotation_table.setColumnWidth(4, 110)
        ann_panel_layout.addWidget(self.review_annotation_table)

        bm_panel = QtWidgets.QWidget()
        bm_panel_layout = QtWidgets.QVBoxLayout(bm_panel)
        bm_panel_layout.setContentsMargins(0, 0, 0, 0)

        bm_review_hdr = QtWidgets.QLabel("Bookmarks")
        bm_review_hdr.setAlignment(QtCore.Qt.AlignCenter)
        bm_review_hdr.setStyleSheet("QLabel { background: #ede9fe; color: #5b21b6; border-radius: 8px; padding: 5px; font-weight: bold; }")
        bm_panel_layout.addWidget(bm_review_hdr)

        self.review_bookmark_table = QtWidgets.QTableWidget(0, 3)
        self.review_bookmark_table.setHorizontalHeaderLabels(["BOOKMARK TEXT", "LEVEL", "PAGE NO"])
        self.review_bookmark_table.setSelectionBehavior(QtWidgets.QTableWidget.SelectRows)
        self.review_bookmark_table.setSelectionMode(QtWidgets.QTableWidget.SingleSelection)
        self.review_bookmark_table.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)
        self.review_bookmark_table.setAlternatingRowColors(True)
        self.review_bookmark_table.verticalHeader().setDefaultSectionSize(28)

        review_bm_header = self.review_bookmark_table.horizontalHeader()
        review_bm_header.setDefaultAlignment(QtCore.Qt.AlignCenter)
        review_bm_header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        review_bm_header.setSectionResizeMode(1, QtWidgets.QHeaderView.Fixed)
        review_bm_header.setSectionResizeMode(2, QtWidgets.QHeaderView.Fixed)
        self.review_bookmark_table.setColumnWidth(1, 80)
        self.review_bookmark_table.setColumnWidth(2, 90)
        bm_panel_layout.addWidget(self.review_bookmark_table)

        review_splitter.addWidget(ann_panel)
        review_splitter.addWidget(bm_panel)
        review_splitter.setSizes([700, 500])

        self.btn_generate_pdf = mkbtn("Generate Final Output PDF", "#ff9933", "#ffbc80", "#222")

        review_layout.addWidget(review_splitter, 1)
        review_layout.addWidget(self.btn_generate_pdf)
        self.main_stack.addWidget(review_page)

        QtCore.QTimer.singleShot(0, self.init_splitter_sizes)

        # Connections
        self.btn_open.clicked.connect(self.open_pdf)
        self.btn_prev.clicked.connect(self.prev_page)
        self.btn_next.clicked.connect(self.next_page)

        self.btn_annotation.clicked.connect(lambda: self.switch_mode("annotation"))
        self.btn_bookmark.clicked.connect(lambda: self.switch_mode("bookmark"))
        self.btn_review.clicked.connect(self.open_review_safe)

        self.btn_generate_pdf.clicked.connect(self.generate_final_output_pdf)

        self.annotation_table.itemSelectionChanged.connect(self.on_annotation_selection_changed)
        self.bookmark_table.itemSelectionChanged.connect(self.on_bookmark_selection_changed)

        self.review_annotation_table.itemSelectionChanged.connect(self.on_review_annotation_selection_changed)
        self.review_bookmark_table.itemSelectionChanged.connect(self.on_review_bookmark_selection_changed)

        self.btn_export_ann.clicked.connect(self.export_annotations_csv)
        self.btn_load_ann.clicked.connect(self.load_annotations_csv)
        self.btn_delete_ann.clicked.connect(self.delete_selected_annotation)

        self.btn_export_bm.clicked.connect(self.export_bookmarks_csv)
        self.btn_load_bm.clicked.connect(self.load_bookmarks_csv)
        self.btn_delete_bm.clicked.connect(self.delete_selected_bookmark)

        self.switch_mode("annotation", force=True)

    def init_splitter_sizes(self):
        total = max(self.height() - 200, 700)
        if self.active_mode in ("annotation", "bookmark"):
            self.splitter.setSizes([int(total * 0.80), int(total * 0.20)])
        else:
            self.splitter.setSizes([int(total * 0.72), int(total * 0.28)])

    def set_active_mode_button_styles(self):
        base_off = 0.95

        def style_button(btn, bg, hover, active=False):
            if active:
                btn.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {bg};
                        color: white;
                        font-family: 'Times New Roman';
                        font-weight: bold;
                        font-size: 11pt;
                        border-radius: 10px;
                        padding: 6px 14px;
                        min-height: 34px;
                        border: 3px solid #111827;
                    }}
                    QPushButton:hover:!disabled {{
                        background-color: {hover};
                    }}
                    QPushButton:disabled {{
                        background-color: #d9d9d9;
                        color: #7a7a7a;
                        border: 1px solid #c8c8c8;
                    }}
                """)
            else:
                btn.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {bg};
                        color: white;
                        font-family: 'Times New Roman';
                        font-weight: bold;
                        font-size: 11pt;
                        border-radius: 10px;
                        padding: 6px 14px;
                        min-height: 34px;
                        border: 1px solid #00000022;
                    }}
                    QPushButton:hover:!disabled {{
                        background-color: {hover};
                    }}
                    QPushButton:disabled {{
                        background-color: #d9d9d9;
                        color: #7a7a7a;
                        border: 1px solid #c8c8c8;
                    }}
                """)

        style_button(self.btn_annotation, "#2563eb", "#4a7df2", self.active_mode == "annotation")
        style_button(self.btn_bookmark, "#7c3aed", "#9b63f0", self.active_mode == "bookmark")
        style_button(self.btn_review, "#059669", "#21b58a", self.active_mode == "review")

    def switch_mode(self, mode, force=False):
        if not force and mode in ("annotation", "bookmark", "review") and not self.pdf_loaded and mode != "annotation":
            return

        self.active_mode = mode
        self.set_active_mode_button_styles()

        if mode == "annotation":
            self.main_stack.setCurrentIndex(0)
            self.bottom_stack.setCurrentIndex(0)
            self.tip.setText("Click anywhere on the PDF to add annotation. Drag an existing box to reposition it.")
            QtCore.QTimer.singleShot(0, lambda: self.splitter.setSizes([int(max(self.height() - 200, 700) * 0.82), int(max(self.height() - 200, 700) * 0.18)]))
        elif mode == "bookmark":
            self.main_stack.setCurrentIndex(0)
            self.bottom_stack.setCurrentIndex(1)
            self.tip.setText("Click anywhere on the PDF to add bookmark for the current page.")
            QtCore.QTimer.singleShot(0, lambda: self.splitter.setSizes([int(max(self.height() - 200, 700) * 0.82), int(max(self.height() - 200, 700) * 0.18)]))
        else:
            self.main_stack.setCurrentIndex(1)
            self.refresh_review_tables()

        self.image_label.update()

    def open_review_safe(self):
        if not self.btn_review.isEnabled():
            return
        self.switch_mode("review")

    def update_navigation_buttons(self):
        has_pdf = self.doc is not None
        self.btn_prev.setEnabled(has_pdf and self.current_page_index > 0)
        self.btn_next.setEnabled(has_pdf and self.current_page_index < len(self.doc) - 1)

    def check_review_enable(self):
        self.btn_review.setEnabled(self.has_annotation or self.has_bookmark)

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

            self.pdf_loaded = True
            self.has_annotation = False
            self.has_bookmark = False

            self.btn_annotation.setEnabled(True)
            self.btn_bookmark.setEnabled(True)
            self.check_review_enable()

            self.refresh_annotation_table()
            self.refresh_bookmark_table()
            self.refresh_review_tables()
            self.render_page()
            self.switch_mode("annotation")
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
        self.update_navigation_buttons()

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
        self.refresh_review_tables()
        self.select_annotation_row_silent(self.selected_entry_index)
        self.image_label.update()

        self.has_annotation = True
        self.check_review_enable()

    def capture_bookmark_point(self):
        if not self.doc:
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
        self.refresh_review_tables()
        self.select_bookmark_row_silent(self.selected_bookmark_index)
        self.image_label.update()

        self.has_bookmark = True
        self.check_review_enable()

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

    def populate_review_annotation_table(self):
        self.review_annotation_table.setRowCount(0)

        for row_idx, e in enumerate(self.entries):
            self.review_annotation_table.insertRow(row_idx)
            vals = [e.domain, e.name, str(e.pageno), e.annotation, e.assignedfield]
            for col_idx, val in enumerate(vals):
                item = QtWidgets.QTableWidgetItem(val)
                if col_idx in (0, 1, 2, 4):
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                elif col_idx == 3:
                    item.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
                self.review_annotation_table.setItem(row_idx, col_idx, item)

        self.review_annotation_table.resizeRowsToContents()

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

    def on_review_annotation_selection_changed(self):
        row = self.review_annotation_table.currentRow()
        if 0 <= row < len(self.entries):
            self.selected_entry_index = row
            entry = self.entries[row]
            if self.doc:
                self.current_page_index = max(0, entry.pageno - 1)
                self.render_page()

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

    def populate_review_bookmark_table(self):
        self.review_bookmark_table.setRowCount(0)

        for row_idx, b in enumerate(self.bookmarks):
            self.review_bookmark_table.insertRow(row_idx)
            vals = [b.title, str(b.level), str(b.pageno)]
            for col_idx, val in enumerate(vals):
                item = QtWidgets.QTableWidgetItem(val)
                if col_idx in (1, 2):
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                else:
                    item.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                self.review_bookmark_table.setItem(row_idx, col_idx, item)

        self.review_bookmark_table.resizeRowsToContents()

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

    def on_review_bookmark_selection_changed(self):
        row = self.review_bookmark_table.currentRow()
        if 0 <= row < len(self.bookmarks):
            self.selected_bookmark_index = row
            bm = self.bookmarks[row]
            if self.doc:
                self.current_page_index = max(0, bm.pageno - 1)
                self.render_page()

    def refresh_review_tables(self):
        self.populate_review_annotation_table()
        self.populate_review_bookmark_table()

    # ------------------------------------------------------------------
    # Delete actions
    # ------------------------------------------------------------------
    def delete_selected_annotation(self):
        row = self.annotation_table.currentRow()
        if row < 0 or row >= len(self.entries):
            QtWidgets.QMessageBox.information(self, "Delete Annotation", "Please select an annotation row to delete.")
            return

        reply = QtWidgets.QMessageBox.question(
            self,
            "Delete Annotation",
            "Delete the selected annotation?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No
        )
        if reply != QtWidgets.QMessageBox.Yes:
            return

        del self.entries[row]
        self.selected_entry_index = -1
        self.refresh_annotation_table()
        self.refresh_review_tables()
        self.image_label.update()

        self.has_annotation = len(self.entries) > 0
        self.check_review_enable()

    def delete_selected_bookmark(self):
        row = self.bookmark_table.currentRow()
        if row < 0 or row >= len(self.bookmarks):
            QtWidgets.QMessageBox.information(self, "Delete Bookmark", "Please select a bookmark row to delete.")
            return

        reply = QtWidgets.QMessageBox.question(
            self,
            "Delete Bookmark",
            "Delete the selected bookmark?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No
        )
        if reply != QtWidgets.QMessageBox.Yes:
            return

        del self.bookmarks[row]
        self.selected_bookmark_index = -1
        self.refresh_bookmark_table()
        self.refresh_review_tables()
        self.image_label.update()

        self.has_bookmark = len(self.bookmarks) > 0
        self.check_review_enable()

    # ------------------------------------------------------------------
    # CSV export / import
    # ------------------------------------------------------------------
    def export_annotations_csv(self):
        if not self.entries:
            QtWidgets.QMessageBox.information(self, "Export Annotation CSV", "No annotations available to export.")
            return

        default_path = os.path.join(
            os.path.dirname(self.open_pdf_path) if self.open_pdf_path else SCRIPT_DIR,
            "annotation_entries.csv"
        )
        out_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Export Annotation CSV", default_path, "CSV Files (*.csv)"
        )
        if not out_path:
            return

        try:
            with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=ANNOTATION_CSV_COLUMNS)
                writer.writeheader()
                for e in self.entries:
                    writer.writerow({
                        "DOMAIN": e.domain,
                        "NAME": e.name,
                        "PAGENO": e.pageno,
                        "ANNOTATION": e.annotation,
                        "ASSIGNEDFIELD": e.assignedfield,
                        "X1": e.x1,
                        "Y1": e.y1,
                        "PAGEH": e.pageh,
                    })
            QtWidgets.QMessageBox.information(self, "Export Annotation CSV", f"Annotation CSV exported successfully:\n{out_path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Export Annotation CSV", f"Failed to export annotation CSV:\n{e}")

    def load_annotations_csv(self):
        if not self.doc:
            QtWidgets.QMessageBox.warning(self, "Load Annotation CSV", "Please open the source PDF first.")
            return

        in_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Load Annotation CSV", "", "CSV Files (*.csv)"
        )
        if not in_path:
            return

        try:
            loaded_entries = []
            with open(in_path, "r", newline="", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                expected = set(ANNOTATION_CSV_COLUMNS)
                actual = set([c.strip().upper() for c in (reader.fieldnames or [])])

                if not expected.issubset(actual):
                    raise ValueError(
                        "Annotation CSV is missing required columns.\n"
                        f"Required: {', '.join(ANNOTATION_CSV_COLUMNS)}"
                    )

                for row in reader:
                    domain = (row.get("DOMAIN") or "").strip()
                    name = (row.get("NAME") or "").strip()
                    annotation = (row.get("ANNOTATION") or "").strip()
                    assignedfield = (row.get("ASSIGNEDFIELD") or "").strip()
                    pageno = int(float(row.get("PAGENO") or 0))
                    x1 = float(row.get("X1") or 0)
                    y1 = float(row.get("Y1") or 0)
                    pageh = float(row.get("PAGEH") or (self.page_rect.height if self.page_rect else 0))

                    is_domain_annotation, is_assigned_field, is_not_submitted = bool_from_entry(
                        domain, name, annotation, assignedfield
                    )

                    loaded_entries.append(
                        AnnotationEntry(
                            domain=domain,
                            name=name,
                            pageno=pageno,
                            annotation=annotation,
                            assignedfield=assignedfield,
                            x1=x1,
                            y1=y1,
                            pageh=pageh,
                            is_domain_annotation=is_domain_annotation,
                            is_assigned_field=is_assigned_field,
                            is_not_submitted=is_not_submitted
                        )
                    )

            self.entries = loaded_entries
            self.selected_entry_index = -1
            self.refresh_annotation_table()
            self.refresh_review_tables()
            self.image_label.update()

            self.has_annotation = len(self.entries) > 0
            self.check_review_enable()

            QtWidgets.QMessageBox.information(self, "Load Annotation CSV", f"Loaded {len(self.entries)} annotation row(s).")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Load Annotation CSV", f"Failed to load annotation CSV:\n{e}")

    def export_bookmarks_csv(self):
        if not self.bookmarks:
            QtWidgets.QMessageBox.information(self, "Export Bookmark CSV", "No bookmarks available to export.")
            return

        default_path = os.path.join(
            os.path.dirname(self.open_pdf_path) if self.open_pdf_path else SCRIPT_DIR,
            "bookmark_entries.csv"
        )
        out_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Export Bookmark CSV", default_path, "CSV Files (*.csv)"
        )
        if not out_path:
            return

        try:
            with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=BOOKMARK_CSV_COLUMNS)
                writer.writeheader()
                for b in self.bookmarks:
                    writer.writerow({
                        "TITLE": b.title,
                        "LEVEL": b.level,
                        "PAGENO": b.pageno,
                    })
            QtWidgets.QMessageBox.information(self, "Export Bookmark CSV", f"Bookmark CSV exported successfully:\n{out_path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Export Bookmark CSV", f"Failed to export bookmark CSV:\n{e}")

    def load_bookmarks_csv(self):
        if not self.doc:
            QtWidgets.QMessageBox.warning(self, "Load Bookmark CSV", "Please open the source PDF first.")
            return

        in_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Load Bookmark CSV", "", "CSV Files (*.csv)"
        )
        if not in_path:
            return

        try:
            loaded_bookmarks = []
            with open(in_path, "r", newline="", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                expected = set(BOOKMARK_CSV_COLUMNS)
                actual = set([c.strip().upper() for c in (reader.fieldnames or [])])

                if not expected.issubset(actual):
                    raise ValueError(
                        "Bookmark CSV is missing required columns.\n"
                        f"Required: {', '.join(BOOKMARK_CSV_COLUMNS)}"
                    )

                for row in reader:
                    title = (row.get("TITLE") or "").strip()
                    if not title:
                        continue
                    level = int(float(row.get("LEVEL") or 1))
                    pageno = int(float(row.get("PAGENO") or 1))

                    loaded_bookmarks.append(
                        BookmarkEntry(
                            title=title,
                            level=level,
                            pageno=pageno
                        )
                    )

            self.bookmarks = loaded_bookmarks
            self.selected_bookmark_index = -1
            self.refresh_bookmark_table()
            self.refresh_review_tables()
            self.image_label.update()

            self.has_bookmark = len(self.bookmarks) > 0
            self.check_review_enable()

            QtWidgets.QMessageBox.information(self, "Load Bookmark CSV", f"Loaded {len(self.bookmarks)} bookmark row(s).")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Load Bookmark CSV", f"Failed to load bookmark CSV:\n{e}")

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

            if self.entries:
                for e in self.entries:
                    if e.pageno < 1 or e.pageno > len(doc):
                        continue

                    page = doc[e.pageno - 1]
                    color_map = get_page_domain_color_map(self.entries, e.pageno)
                    layout = compute_entry_layout(e, color_map, page_width=page.rect.width)
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