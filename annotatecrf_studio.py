# ====================================================================================================
# Script Name    : annotatecrf_studio.py
#
# Description    : Desktop GUI utility for capturing annotation positions from a PDF and generating
#                  final visible comment boxes directly into a PDF without XFDF / Adobe dependency.
#
# Version        : full updated script with annotation/bookmark management, TOC support,
#                  refer-page hyperlink correction, and CSV append/duplicate-check fixes
#
#                  Key Updates / Enhancements:
#                    [1] Added annotation copy option
#                    [2] Added annotation edit option
#                    [3] Added bookmark copy option
#                    [4] Added bookmark edit option
#                    [5] Added clickable TOC generation
#                    [6] Added refer-page annotation option
#                    [7] Added internal page hyperlinks for refer-page annotations
#                    [8] Fixed refer-page hyperlink target when TOC pages are inserted
#                    [9] Added page number adjustment when TOC is added
#                   [10] Added bookmark ordering / branch placement update support
#                   [11] Added export of three CSV outputs for review and reuse
#                   [12] Added append load support for CSVs
#                   [13] Added duplicate-check handling during CSV load
#                   [14] Allowed annotation CSV import even when position fields are blank
#                   [15] Improved annotation box sizing, wrapping, and preview behavior
#                   [16] Added connector line support and editing
#                   [17] Added page jump / go-to-page navigation
#                   [18] Added multi-selection copy support for annotations
#
#                  Workflow:
#                    - Open PDF first
#                    - Annotation mode:
#                        * Click on PDF -> enter details -> preview box appears immediately
#                        * Drag an existing preview box to move it
#                        * Copy / edit / delete annotation rows
#                        * Manage annotation rows + export/load CSV
#                    - Bookmark mode:
#                        * Click on PDF -> enter bookmark details for current page
#                        * Copy / edit / delete bookmark rows
#                        * Manage bookmark rows + export/load CSV
#                    - Review mode:
#                        * View annotation and bookmark tables side by side
#                        * Add TOC if needed
#                        * Generate Final Output PDF
#
#                  Annotation CSV columns used internally:
#                    TYPE, DOMAIN, NAME, PAGENO, ANNOTATION, ASSIGNEDFIELD,
#                    X1, Y1, PAGEH, BOX_W, BOX_H,
#                    LINE_PAGENO, LINE_X1, LINE_Y1, LINE_X2, LINE_Y2
#
#                    X1/Y1/PAGEH may be left blank during import; the tool will place the
#                    annotation at a default location and it can then be dragged manually.
#
#                  Bookmark CSV columns used internally:
#                    TITLE, LEVEL, PAGENO
#
#                  Output Files:
#                    - Final annotated PDF
#                    - Annotation CSV
#                    - Bookmark CSV
#                    - Connector line CSV / details export (if applicable)
# ====================================================================================================

import csv
import importlib
import os
import re
import subprocess
import sys
from collections import OrderedDict
from dataclasses import dataclass
from typing import List, Optional

# ================================
# Auto-install required packages
# Works on Windows and macOS
# ================================
REQUIRED_PACKAGES = [
    ("PyMuPDF", "fitz"),
    ("PyQt5", "PyQt5"),
]

def install_if_missing(package_name, import_name=None):
    module_name = import_name or package_name
    try:
        importlib.import_module(module_name)
    except ImportError:
        print(f"[Installing] {package_name} ...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
        importlib.invalidate_caches()
        importlib.import_module(module_name)
        print(f"[Done] {package_name} installed.")

for package_name, import_name in REQUIRED_PACKAGES:
    install_if_missing(package_name, import_name)

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
BOX_HEIGHT_NORMAL = 13.0
BOX_HEIGHT_DOMAIN = 16.0

MAX_WRAP_WIDTH = 520.0
RIGHT_PAGE_MARGIN = 12.0
MIN_BOX_WIDTH = 55.0
MIN_BOX_WIDTH_SINGLE = 42.0
MIN_BOX_WIDTH_DOMAIN = 70.0
MIN_BOX_WIDTH_MULTI = 90.0
LONG_TEXT_THRESHOLD = 22
LONG_TEXT_MIN_WIDTH = 220.0

BOOKMARK_PREVIEW_COLOR = QtGui.QColor("#7c3aed")
BOOKMARK_PREVIEW_TEXT_COLOR = QtGui.QColor("#5b21b6")

ANNOTATION_CSV_COLUMNS = ["TYPE", "DOMAIN", "NAME", "PAGENO", "ANNOTATION", "ASSIGNEDFIELD", "X1", "Y1", "PAGEH", "BOX_W", "BOX_H", "LINE_PAGENO", "LINE_X1", "LINE_Y1", "LINE_X2", "LINE_Y2"]
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


def clean_number(value, default=None):
    try:
        s = str(value).replace(" ", "").strip()
        if s == "":
            return default
        return float(s)
    except Exception:
        return default


def wrap_text_by_width(text: str, max_width: float, scale: float = 1.0) -> List[str]:
    if text is None:
        return [""]
    text = str(text).replace("	", "    ")
    if text == "":
        return [""]

    final_lines = []
    raw_lines = text.splitlines() or [text]

    for raw_line in raw_lines:
        if raw_line == "":
            final_lines.append("")
            continue

        indent_len = len(raw_line) - len(raw_line.lstrip(" "))
        indent = raw_line[:indent_len]
        content = raw_line[indent_len:]

        if content == "":
            final_lines.append(indent)
            continue

        words = content.split()
        current = indent + words[0]
        for word in words[1:]:
            trial = current + " " + word
            if estimate_text_width(trial, scale) <= max_width:
                current = trial
            else:
                final_lines.append(current)
                current = indent + word
        final_lines.append(current)

    return final_lines or [""]


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

    text = (entry.annotation or "").replace("	", "    ").rstrip()
    text_lines_raw = text.splitlines() or [text]
    raw_w = max((estimate_text_width(line, scale=scale) for line in text_lines_raw), default=MIN_BOX_WIDTH_MULTI)

    allowed_width = MAX_WRAP_WIDTH
    if page_width is not None:
        available = max(MIN_BOX_WIDTH_SINGLE, page_width - entry.x1 - RIGHT_PAGE_MARGIN)
        allowed_width = min(MAX_WRAP_WIDTH, available)

    single_line_min = MIN_BOX_WIDTH_DOMAIN if bold else MIN_BOX_WIDTH_SINGLE
    explicit_w = clean_number(getattr(entry, "box_w", None), None)

    if explicit_w is not None:
        box_w = max(single_line_min, min(allowed_width, explicit_w))
        lines = wrap_text_by_width(text, box_w, scale=scale)
    else:
        tight_width = raw_w + (TEXT_PADDING_X * 2.0) + 6.0
        raw_has_newline = "\n" in text
        fits_single_line = (not raw_has_newline) and (tight_width <= allowed_width)

        if fits_single_line:
            box_w = max(single_line_min, min(allowed_width, tight_width))
            lines = [text]
        else:
            preferred_width = raw_w
            long_text = (
                len(text) >= LONG_TEXT_THRESHOLD
                or " when " in f" {text.lower()} "
                or raw_has_newline
            )
            if long_text:
                preferred_width = max(preferred_width, LONG_TEXT_MIN_WIDTH)

            multi_min = MIN_BOX_WIDTH_DOMAIN if bold else MIN_BOX_WIDTH_MULTI
            box_w = max(multi_min, min(allowed_width, preferred_width))
            lines = wrap_text_by_width(text, box_w, scale=scale)
            actual_w = max((estimate_text_width(line, scale=scale) for line in lines), default=box_w)
            if len(lines) <= 1 and actual_w + (TEXT_PADDING_X * 2.0) + 6.0 <= allowed_width:
                box_w = max(single_line_min, min(allowed_width, actual_w + (TEXT_PADDING_X * 2.0) + 6.0))
                lines = [text]
            else:
                box_w = max(multi_min, min(allowed_width, max(box_w, actual_w + 2.0)))
                lines = wrap_text_by_width(text, box_w, scale=scale)

    explicit_h = clean_number(getattr(entry, "box_h", None), None)
    if explicit_h is not None:
        box_h = max(base_h, explicit_h)
    else:
        if len(lines) <= 1:
            preview_padding_top = 1.5
            preview_padding_bottom = 1.5
            box_h = max(base_h, font_size + preview_padding_top + preview_padding_bottom + 0.2)
        elif len(lines) == 2:
            preview_line_gap = font_size + 0.6
            preview_padding_top = 2.0
            preview_padding_bottom = 2.0
            box_h = max(base_h, len(lines) * preview_line_gap + preview_padding_top + preview_padding_bottom)
        else:
            preview_line_gap = font_size + 0.8
            preview_padding_top = 2.5
            preview_padding_bottom = 2.5
            box_h = max(
                base_h,
                len(lines) * preview_line_gap + preview_padding_top + preview_padding_bottom
            )

    return {
        "fill": fill,
        "dashed": dashed,
        "bold": bold,
        "font_size": font_size,
        "lines": lines,
        "box_w": box_w,
        "box_h": box_h
    }
def extract_page_reference(text: str):
    if not text:
        return None
    m = re.search(r'\b(?:for\s+annotations\s+)?(?:refer\s+to|see)?\s*page\s+(\d+)\b', text, re.IGNORECASE)
    if m:
        try:
            return int(m.group(1))
        except Exception:
            return None
    return None


def add_internal_page_link(page, rect, target_page: int):
    if not target_page or target_page < 1:
        return
    try:
        page.insert_link({
            "kind": fitz.LINK_GOTO,
            "from": rect,
            "page": target_page - 1,
            "to": fitz.Point(72, 72)
        })
    except Exception:
        pass


def get_pdf_base_output_path(pdf_path: str):
    base, _ = os.path.splitext(pdf_path)
    return base


def qcolor_from_rgb01(rgb):
    return QtGui.QColor(int(rgb[0] * 255), int(rgb[1] * 255), int(rgb[2] * 255))


def measure_textbox_height(text, width, fontname, fontsize, min_height):
    """Use PyMuPDF textbox fitting to determine a safe box height.
    Keeps single-line annotations compact while allowing long wrapped text to grow.
    """
    content = (text or "").rstrip()
    if not content:
        return float(min_height)

    inner_width = max(24.0, float(width) - (TEXT_PADDING_X * 2.0))

    # Compact path for single-line text that already fits within the box width
    if "\n" not in content and estimate_text_width(content) <= inner_width:
        return max(float(min_height), fontsize + 4.5)

    probe_doc = fitz.open()
    try:
        probe_page = probe_doc.new_page(width=max(100.0, inner_width + 20.0), height=4000.0)
        test_height = max(float(min_height), fontsize + 4.0)
        for _ in range(120):
            rect = fitz.Rect(10.0, 10.0, 10.0 + inner_width, 10.0 + test_height)
            try:
                rc = probe_page.insert_textbox(
                    rect,
                    content,
                    fontname=fontname,
                    fontsize=fontsize,
                    color=TEXT_COLOR,
                    align=fitz.TEXT_ALIGN_LEFT,
                    lineheight=1.0,
                    overlay=False
                )
            except Exception:
                rc = probe_page.insert_textbox(
                    rect,
                    content,
                    fontname=FONT_NORMAL,
                    fontsize=fontsize,
                    color=TEXT_COLOR,
                    align=fitz.TEXT_ALIGN_LEFT,
                    lineheight=1.0,
                    overlay=False
                )
            if rc >= 0:
                return max(float(min_height), test_height + 2.0)
            test_height += max(4.0, fontsize * 0.75)
        return max(float(min_height), test_height + 3.0)
    finally:
        probe_doc.close()


def compute_output_box_height(text_lines, width, bold, font_size, base_h):
    """Height for final PDF output only. Keeps preview compact while output stays safe."""
    return measure_textbox_height(
        "\n".join(text_lines or []),
        width,
        FONT_BOLD if bold else FONT_NORMAL,
        font_size,
        base_h
    )


def draw_box_and_text_pdf(page, rect, text_lines, fill_color, bold=False, dashed=False, font_size=10, link_target_page=None):
    page.draw_rect(rect, color=None, fill=fill_color, overlay=True)

    dashes = "[3 3] 0" if dashed else None
    page.draw_rect(rect, color=BORDER_COLOR, fill=None, width=0.8, dashes=dashes, overlay=True)

    fontname = FONT_BOLD if bold else FONT_NORMAL
    full_text = "\n".join(text_lines or [])
    page_ref = int(link_target_page) if link_target_page else extract_page_reference(full_text)

    line_count = max(1, len(text_lines or []))
    if line_count == 1:
        lineheight = 0.95
        visual_factor = 0.80
    elif line_count == 2:
        lineheight = 0.98
        visual_factor = 0.92
    else:
        lineheight = 1.00
        visual_factor = 0.96

    text_block_height = max(font_size + 1.0, line_count * font_size * visual_factor * lineheight)
    available_height = max(text_block_height, rect.height - 2.0)
    y_offset = max(1.0, (available_height - text_block_height) / 2.0)

    inner_rect = fitz.Rect(
        rect.x0 + TEXT_PADDING_X,
        rect.y0 + y_offset,
        rect.x1 - TEXT_PADDING_X,
        rect.y0 + y_offset + text_block_height
    )

    inserted = -1
    try:
        inserted = page.insert_textbox(
            inner_rect,
            full_text,
            fontname=fontname,
            fontsize=font_size,
            color=TEXT_COLOR,
            align=fitz.TEXT_ALIGN_LEFT,
            lineheight=lineheight,
            overlay=True
        )
    except Exception:
        try:
            inserted = page.insert_textbox(
                inner_rect,
                full_text,
                fontname=FONT_NORMAL,
                fontsize=font_size,
                color=TEXT_COLOR,
                align=fitz.TEXT_ALIGN_LEFT,
                lineheight=lineheight,
                overlay=True
            )
        except Exception:
            inserted = -1

    if inserted < 0 and text_lines:
        # Safe fallback if insert_textbox cannot fit due to font metric mismatch
        line_gap = font_size * (1.00 if line_count >= 3 else 0.98)
        text_x = rect.x0 + TEXT_PADDING_X
        text_y = rect.y0 + y_offset + font_size - 1.0
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

    if page_ref:
        # Make the whole annotation box clickable for refer-page annotations.
        add_internal_page_link(page, rect, page_ref)


def sanitize_output_pdf_path(src_pdf: str):
    base, _ = os.path.splitext(src_pdf)
    return base + "_final.pdf"

def extract_toc_entries(doc):
    toc = doc.get_toc(simple=True)
    entries = []
    for entry in toc:
        level, title, page_num = entry[:3]
        entries.append((level, title.strip(), page_num - 1))  # 0-based page numbers
    return entries

def wrap_text(text, font_size, max_width):
    words = text.split()
    lines = []
    current_line = ""
    for word in words:
        test_line = current_line + (" " if current_line else "") + word
        if fitz.get_text_length(test_line, fontsize=font_size) <= max_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            current_line = word
    if current_line:
        lines.append(current_line)
    return lines

def paginate_wrapped_entries(toc_entries, font_size, max_width, lines_per_page):
    paginated_entries = []
    current_page_entries = []
    current_line_count = 0

    for entry in toc_entries:
        level, title, target_page = entry
        indent = 20 * (level - 1)
        available_width = max_width - indent
        wrapped_lines = wrap_text(title, font_size, available_width)
        line_count = len(wrapped_lines)

        if current_line_count + line_count > lines_per_page:
            paginated_entries.append(current_page_entries)
            current_page_entries = []
            current_line_count = 0

        current_page_entries.append((entry, wrapped_lines))
        current_line_count += line_count

    if current_page_entries:
        paginated_entries.append(current_page_entries)

    return paginated_entries

def generate_toc_pages(paginated_entries, font_size, page_width, page_height):
    toc_doc = fitz.open()
    link_targets = []
    left_margin = 50
    right_margin = 60
    top_margin = 50
    y_spacing = font_size * 1.5

    toc_page_count = len(paginated_entries)

    for page_index, entries in enumerate(paginated_entries):
        page = toc_doc.new_page(width=page_width, height=page_height)
        y = top_margin
        for (level, title, target_page), wrapped_lines in entries:
            indent = 20 * (level - 1)
            x = left_margin + indent
            page_number_str = str(target_page + toc_page_count + 1)
            page_number_width = fitz.get_text_length(page_number_str, fontsize=font_size)
            max_x_for_dots = page_width - right_margin - page_number_width - 5

            first_line_y = y  # Needed for hyperlink rectangle

            for i, line in enumerate(wrapped_lines):
                line_width = fitz.get_text_length(line, fontsize=font_size)
                dots = ''
                if i == len(wrapped_lines) - 1:
                    dots_space = max_x_for_dots - (x + line_width + 10)
                    dot_count = max(0, int(dots_space / fitz.get_text_length('.', fontsize=font_size)))
                    dots = '.' * dot_count

                    # Draw line with dots and page number
                    page.insert_text((x, y), f"{line} {dots}", fontsize=font_size)
                    page.insert_text((page_width - right_margin - page_number_width, y), page_number_str, fontsize=font_size)
                else:
                    # Draw line without dots/page number
                    page.insert_text((x, y), line, fontsize=font_size)

                y += y_spacing

            rect = fitz.Rect(x, first_line_y - font_size, page_width - right_margin, y)
            link_targets.append((page_index, rect, target_page))

    return toc_doc, link_targets

def add_toc_hyperlinks(doc, link_targets, toc_page_count):
    for toc_page_index, rect, target_page in link_targets:
        doc[toc_page_index].insert_link({
            "kind": fitz.LINK_GOTO,
            "from": rect,
            "page": target_page + toc_page_count
        })

def shift_bookmark_pages(bookmarks, offset):
    shifted = []
    for bm in bookmarks:
        if len(bm) >= 3:
            level, title, page_num = bm[:3]
            shifted.append([level, title, page_num + offset])
    return shifted

def add_existing_bookmarks(doc, bookmarks, offset):
    shifted = shift_bookmark_pages(bookmarks, offset)
    doc.set_toc(shifted)


def build_toc_entries_from_bookmarks(bookmarks):
    entries = []
    for bm in bookmarks or []:
        try:
            level = max(1, int(getattr(bm, "level", 1)))
            title = str(getattr(bm, "title", "")).strip()
            target_page = max(0, int(getattr(bm, "pageno", 1)) - 1)
        except Exception:
            continue
        if title:
            entries.append((level, title, target_page))
    return entries


def compute_toc_page_count_from_bookmarks(bookmarks, font_size: int = 12, lines_per_page: int = 38):
    toc_entries = build_toc_entries_from_bookmarks(bookmarks)
    if not toc_entries:
        return 0
    width, _ = fitz.paper_size("a4")
    max_width = width - 120
    paginated_entries = paginate_wrapped_entries(toc_entries, font_size, max_width, lines_per_page)
    return len(paginated_entries)


def adjust_page_refs_in_text(text: str, page_offset: int) -> str:
    if not text or not page_offset:
        return text

    def repl(match):
        try:
            page_num = int(match.group(1))
            return f"page {page_num + page_offset}"
        except Exception:
            return match.group(0)

    return re.sub(r"\bpage\s+(\d+)\b", repl, text, flags=re.IGNORECASE)


def create_clickable_toc_pdf(input_pdf: str, output_pdf: str, font_size: int = 12, lines_per_page: int = 38):
    original = fitz.open(input_pdf)
    try:
        toc_entries = extract_toc_entries(original)
        original_bookmarks = original.get_toc()

        if not toc_entries:
            original.save(output_pdf)
            return output_pdf

        width, height = fitz.paper_size("a4")
        max_width = width - 120

        paginated_entries = paginate_wrapped_entries(toc_entries, font_size, max_width, lines_per_page)
        toc_page_count = len(paginated_entries)

        toc_pdf, link_targets = generate_toc_pages(paginated_entries, font_size, width, height)

        final = fitz.open()
        try:
            final.insert_pdf(toc_pdf)
            final.insert_pdf(original)
            add_toc_hyperlinks(final, link_targets, toc_page_count)
            add_existing_bookmarks(final, original_bookmarks, toc_page_count)
            final.save(output_pdf, garbage=4, deflate=True)
        finally:
            final.close()
            toc_pdf.close()

        return output_pdf
    finally:
        original.close()



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
    box_w: Optional[float] = None
    box_h: Optional[float] = None
    line_x1: Optional[float] = None
    line_y1: Optional[float] = None
    line_x2: Optional[float] = None
    line_y2: Optional[float] = None


@dataclass
class BookmarkEntry:
    title: str
    level: int
    pageno: int


@dataclass
class ConnectorLineEntry:
    pageno: int
    x1: float
    y1: float
    x2: float
    y2: float


def _norm_text(value):
    return str(value or "").replace("\t", "    ").strip()


def annotation_entry_key(entry):
    return (
        _norm_text(getattr(entry, "domain", "")).upper(),
        _norm_text(getattr(entry, "name", "")).upper(),
        int(clean_number(getattr(entry, "pageno", 0), 0) or 0),
        _norm_text(getattr(entry, "annotation", "")),
        _norm_text(getattr(entry, "assignedfield", "")).upper(),
        round(float(clean_number(getattr(entry, "x1", 0.0), 0.0) or 0.0), 3),
        round(float(clean_number(getattr(entry, "y1", 0.0), 0.0) or 0.0), 3),
        round(float(clean_number(getattr(entry, "pageh", 0.0), 0.0) or 0.0), 3),
        round(float(clean_number(getattr(entry, "box_w", 0.0), 0.0) or 0.0), 3),
        round(float(clean_number(getattr(entry, "box_h", 0.0), 0.0) or 0.0), 3),
        bool(getattr(entry, "is_domain_annotation", False)),
        bool(getattr(entry, "is_assigned_field", False)),
        bool(getattr(entry, "is_not_submitted", False)),
    )


def connector_line_key(line):
    return (
        int(clean_number(getattr(line, "pageno", 0), 0) or 0),
        round(float(clean_number(getattr(line, "x1", 0.0), 0.0) or 0.0), 3),
        round(float(clean_number(getattr(line, "y1", 0.0), 0.0) or 0.0), 3),
        round(float(clean_number(getattr(line, "x2", 0.0), 0.0) or 0.0), 3),
        round(float(clean_number(getattr(line, "y2", 0.0), 0.0) or 0.0), 3),
    )


def bookmark_entry_key(bookmark):
    return (
        _norm_text(getattr(bookmark, "title", "")),
        int(clean_number(getattr(bookmark, "level", 1), 1) or 1),
        int(clean_number(getattr(bookmark, "pageno", 1), 1) or 1),
    )


# ================================
# Dialog for annotation details
# ================================
class AnnotationDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, initial_data=None, dialog_title="Annotation Details"):
        super().__init__(parent)
        self.setWindowTitle(dialog_title)
        self.resize(700, 500)
        self.setMinimumSize(700, 500)
        self.setModal(True)

        self._updating_refpage_ui = False

        self.setStyleSheet("""
            QDialog { background-color: #f4f8ff; border-radius: 14px; }
            QLabel { font-family: 'Times New Roman'; font-size: 12pt; color: #1c2e4a; }
            QLineEdit, QTextEdit, QSpinBox {
                background: #ffffff; border: 1px solid #a8bfdc; border-radius: 8px;
                padding: 6px 8px; font-family: 'Times New Roman'; font-size: 12pt; color: #1a1a1a;
            }
            QLineEdit:focus, QTextEdit:focus, QSpinBox:focus { border: 2px solid #4d8ef7; background: #fdfefe; }
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
        self.chk_refpage = QtWidgets.QCheckBox("Refer Page")

        self.ref_page_label = QtWidgets.QLabel("Reference Page")
        self.ref_page_spin = QtWidgets.QSpinBox()
        self.ref_page_spin.setRange(1, 999999)
        self.ref_page_spin.setValue(1)
        self.ref_page_spin.setFixedWidth(110)

        self.chk_domain.stateChanged.connect(self.toggle_dialog_state)
        self.chk_assigned.stateChanged.connect(self.toggle_dialog_state)
        self.chk_notsub.stateChanged.connect(self.toggle_dialog_state)
        self.chk_refpage.stateChanged.connect(self.toggle_dialog_state)
        self.ref_page_spin.valueChanged.connect(self.on_ref_page_changed)

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
        chk_row.addWidget(self.chk_refpage)
        chk_row.addStretch()

        ref_row = QtWidgets.QHBoxLayout()
        ref_row.setSpacing(12)
        ref_row.addSpacing(8)
        ref_row.addWidget(self.ref_page_label)
        ref_row.addWidget(self.ref_page_spin)
        ref_row.addStretch()

        note = QtWidgets.QLabel("If none is selected, it is treated as a variable annotation for a collected field. Use Refer Page to create a standard page-reference annotation.")
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
        layout.addLayout(ref_row)
        layout.addWidget(note)
        layout.addSpacing(10)
        layout.addWidget(self.buttons)

        self.toggle_dialog_state()

        if initial_data:
            self.load_initial_data(initial_data)

    def load_initial_data(self, data):
        domain = (data.get("domain") or "").strip()
        name = (data.get("name") or "").strip()
        annotation = (data.get("annotation") or "").rstrip()
        assignedfield = (data.get("assignedfield") or "").strip().upper()
        is_domain_annotation = bool(data.get("is_domain_annotation", False))
        is_assigned_field = bool(data.get("is_assigned_field", assignedfield == "Y"))
        is_not_submitted = bool(data.get("is_not_submitted", False))
        is_refpage = domain.upper() == "REF" and name.upper() == "REF" and annotation.upper().startswith("FOR ANNOTATION REFER TO PAGE ")

        if is_refpage:
            m = re.search(r"page\s+(\d+)", annotation, re.IGNORECASE)
            if m:
                self.ref_page_spin.setValue(int(m.group(1)))
            self.chk_refpage.setChecked(True)
            self.toggle_dialog_state()
            return

        self.chk_domain.setChecked(is_domain_annotation)
        self.chk_assigned.setChecked(is_assigned_field)
        self.chk_notsub.setChecked(is_not_submitted)
        self.domain_edit.setText(domain)
        self.name_edit.setText(name)
        self.annotation_edit.setPlainText(annotation)
        self.toggle_dialog_state()

    def _set_ref_annotation_text(self):
        ref_page = int(self.ref_page_spin.value())
        self.domain_edit.setText("REF")
        self.name_edit.setText("REF")
        self.annotation_edit.setPlainText(f"For annotation details, refer to page {ref_page}.")

    def on_ref_page_changed(self):
        if self.chk_refpage.isChecked():
            self._updating_refpage_ui = True
            try:
                self._set_ref_annotation_text()
            finally:
                self._updating_refpage_ui = False

    def toggle_dialog_state(self):
        if self._updating_refpage_ui:
            return

        is_refpage = self.chk_refpage.isChecked()
        is_notsub = self.chk_notsub.isChecked()
        is_domain = self.chk_domain.isChecked()

        self.ref_page_label.setVisible(is_refpage)
        self.ref_page_spin.setVisible(is_refpage)

        if is_refpage:
            self._updating_refpage_ui = True
            try:
                self.chk_domain.setChecked(False)
                self.chk_assigned.setChecked(False)
                self.chk_notsub.setChecked(False)

                self.chk_domain.setDisabled(True)
                self.chk_assigned.setDisabled(True)
                self.chk_notsub.setDisabled(True)

                self.domain_edit.setDisabled(True)
                self.name_edit.setDisabled(True)
                self.annotation_edit.setDisabled(True)

                self.lbl_name.setVisible(True)
                self.name_edit.setVisible(True)

                self._set_ref_annotation_text()
            finally:
                self._updating_refpage_ui = False
            return

        self.chk_domain.setDisabled(False)
        self.chk_assigned.setDisabled(False)
        self.chk_notsub.setDisabled(False)

        self.domain_edit.setDisabled(False)
        self.name_edit.setDisabled(False)
        self.annotation_edit.setDisabled(False)

        if self.domain_edit.text().strip().upper() == "REF":
            self.domain_edit.clear()
        if self.name_edit.text().strip().upper() == "REF":
            self.name_edit.clear()
        if self.annotation_edit.toPlainText().replace("\t", "    ").rstrip().upper().startswith("FOR ANNOTATION DETAILS, REFER TO PAGE"):
            self.annotation_edit.clear()

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
            self.lbl_name.setVisible(True)
            self.name_edit.setVisible(True)
        else:
            if self.domain_edit.text().strip() == "NOTSUB":
                self.domain_edit.clear()
            if self.name_edit.text().strip() == "NOTSUB":
                self.name_edit.clear()
            if self.annotation_edit.toPlainText().replace("\t", "    ").rstrip() == "[NOT SUBMITTED]":
                self.annotation_edit.clear()

            self.lbl_name.setVisible(not is_domain)
            self.name_edit.setVisible(not is_domain)
            if is_domain:
                self.name_edit.clear()

    def validate_and_accept(self):
        is_refpage = self.chk_refpage.isChecked()
        if is_refpage:
            self.accept()
            return

        if not self.domain_edit.text().strip():
            QtWidgets.QMessageBox.warning(self, "Validation", "DOMAIN is required.")
            return
        if not self.annotation_edit.toPlainText().replace("\t", "    ").rstrip():
            QtWidgets.QMessageBox.warning(self, "Validation", "ANNOTATION is required.")
            return
        if not self.chk_domain.isChecked() and not self.chk_notsub.isChecked():
            if not self.name_edit.text().strip():
                QtWidgets.QMessageBox.warning(self, "Validation", "VARIABLE is required.")
                return
        self.accept()

    def get_values(self):
        is_refpage = self.chk_refpage.isChecked()
        is_domain_annotation = self.chk_domain.isChecked()
        is_assigned_field = self.chk_assigned.isChecked()
        is_not_submitted = self.chk_notsub.isChecked()

        if is_refpage:
            ref_page = int(self.ref_page_spin.value())
            return {
                "domain": "REF",
                "name": "REF",
                "annotation": f"For annotation details, refer to page {ref_page}.",
                "assignedfield": "",
                "is_domain_annotation": False,
                "is_assigned_field": False,
                "is_not_submitted": False
            }

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
            "annotation": self.annotation_edit.toPlainText().replace("\t", "    ").rstrip(),
            "assignedfield": "Y" if is_assigned_field else "",
            "is_domain_annotation": is_domain_annotation,
            "is_assigned_field": is_assigned_field,
            "is_not_submitted": False
        }


# ================================
# Dialog for bookmark details
# ================================
class BookmarkDialog(QtWidgets.QDialog):
    def __init__(self, page_no: int, parent=None, initial_data=None, dialog_title="Bookmark Details",
                 bookmark_choices=None, keep_order_default=False):
        super().__init__(parent)
        self.setWindowTitle(dialog_title)
        self.resize(860, 420)
        self.setMinimumSize(860, 420)
        self.setModal(True)

        self.bookmark_choices = list(bookmark_choices or [])
        self.keep_order_default = bool(keep_order_default)

        self.setStyleSheet("""
            QDialog { background-color: #f4f8ff; border-radius: 14px; }
            QLabel { font-family: 'Times New Roman'; font-size: 12pt; color: #1c2e4a; }
            QLineEdit, QSpinBox, QComboBox {
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

        self.place_after_combo = QtWidgets.QComboBox()
        self.place_after_combo.setMinimumWidth(470)
        if self.keep_order_default:
            self.place_after_combo.addItem("Keep current order", "__KEEP__")
        else:
            self.place_after_combo.addItem("Add at end", "__END__")
        for idx, label in self.bookmark_choices:
            self.place_after_combo.addItem(label, idx)

        form = QtWidgets.QFormLayout()
        form.setLabelAlignment(QtCore.Qt.AlignRight)
        form.setHorizontalSpacing(18)
        form.setVerticalSpacing(14)
        form.addRow("Bookmark Text", self.title_edit)
        form.addRow("Level", self.level_spin)
        form.addRow("Page No", self.page_spin)
        form.addRow("Insert After Branch", self.place_after_combo)

        note = QtWidgets.QLabel(
            "Use 'Insert After Branch' to place the bookmark under the correct bookmark tree. "
            "If you select a bookmark, the new or edited bookmark will be inserted after that bookmark "
            "and all of its child bookmarks."
        )
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

        if initial_data:
            self.title_edit.setText(initial_data.get("title", ""))
            self.level_spin.setValue(int(initial_data.get("level", 1)))
            self.page_spin.setValue(int(initial_data.get("pageno", page_no)))

            selected_after = initial_data.get("insert_after_branch")
            if selected_after == "__KEEP__":
                idx = self.place_after_combo.findData("__KEEP__")
                if idx >= 0:
                    self.place_after_combo.setCurrentIndex(idx)
            elif selected_after == "__END__":
                idx = self.place_after_combo.findData("__END__")
                if idx >= 0:
                    self.place_after_combo.setCurrentIndex(idx)
            elif selected_after is not None:
                idx = self.place_after_combo.findData(selected_after)
                if idx >= 0:
                    self.place_after_combo.setCurrentIndex(idx)

    def validate_and_accept(self):
        if not self.title_edit.text().strip():
            QtWidgets.QMessageBox.warning(self, "Validation", "Bookmark Text is required.")
            return
        self.accept()

    def get_values(self):
        return {
            "title": self.title_edit.text().strip(),
            "level": int(self.level_spin.value()),
            "pageno": int(self.page_spin.value()),
            "insert_after_branch": self.place_after_combo.currentData()
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
        self.resize_entry_index: Optional[int] = None
        self.drag_line_entry_index: Optional[int] = None
        self.drag_line_mode: Optional[str] = None
        self.drag_line_kind: Optional[str] = None
        self.drag_start = QtCore.QPointF(0, 0)

    def _get_entry_rect_on_screen(self, entry, color_map):
        page_width = self.main_window.page_rect.width if self.main_window.page_rect else None
        layout = compute_entry_layout(entry, color_map, page_width=page_width)
        rect_pdf = rect_from_top_origin(entry.x1, entry.y1, layout["box_w"], layout["box_h"], entry.pageh)
        zoom = self.main_window.zoom
        return QtCore.QRectF(rect_pdf.x0 * zoom, rect_pdf.y0 * zoom, rect_pdf.width * zoom, rect_pdf.height * zoom)

    def _line_handles(self, entry):
        z = self.main_window.zoom
        if None in (entry.line_x1, entry.line_y1, entry.line_x2, entry.line_y2):
            return None
        p1 = QtCore.QPointF(entry.line_x1 * z, entry.line_y1 * z)
        p2 = QtCore.QPointF(entry.line_x2 * z, entry.line_y2 * z)
        return p1, p2

    def _connector_line_points(self, line):
        z = self.main_window.zoom
        return QtCore.QPointF(line.x1 * z, line.y1 * z), QtCore.QPointF(line.x2 * z, line.y2 * z)

    def _point_near_line(self, pt, p1, p2, tolerance=7.0):
        line = QtCore.QLineF(p1, p2)
        if line.length() == 0:
            return QtCore.QLineF(pt, p1).length() <= tolerance
        x0, y0 = pt.x(), pt.y()
        x1, y1 = p1.x(), p1.y()
        x2, y2 = p2.x(), p2.y()
        dx, dy = x2 - x1, y2 - y1
        t = ((x0 - x1) * dx + (y0 - y1) * dy) / float(dx * dx + dy * dy)
        t = max(0.0, min(1.0, t))
        proj = QtCore.QPointF(x1 + t * dx, y1 + t * dy)
        return QtCore.QLineF(pt, proj).length() <= tolerance

    def mousePressEvent(self, event):
        if not self.main_window or not self.main_window.pdf_loaded or self.main_window.page_pixmap is None:
            return
        if event.button() != QtCore.Qt.LeftButton:
            return

        current_page = self.main_window.current_page_index + 1
        color_map = get_page_domain_color_map(self.main_window.entries, current_page)

        if self.main_window.active_mode == "annotation":
            if self.main_window.line_capture_stage is not None:
                self.last_click_point = event.pos()
                self.main_window.capture_connector_line_point(event.pos())
                self.update()
                return

            for lidx in reversed(range(len(self.main_window.lines))):
                line = self.main_window.lines[lidx]
                if line.pageno != current_page:
                    continue
                p1, p2 = self._connector_line_points(line)
                if QtCore.QLineF(event.pos(), p1).length() <= 8:
                    self.drag_line_entry_index = lidx
                    self.drag_line_kind = "separate"
                    self.drag_line_mode = "start"
                    self.main_window.selected_line_index = lidx
                    self.main_window.selected_entry_index = -1
                    self.main_window.refresh_selection_only()
                    self.update()
                    return
                if QtCore.QLineF(event.pos(), p2).length() <= 8:
                    self.drag_line_entry_index = lidx
                    self.drag_line_kind = "separate"
                    self.drag_line_mode = "end"
                    self.main_window.selected_line_index = lidx
                    self.main_window.selected_entry_index = -1
                    self.main_window.refresh_selection_only()
                    self.update()
                    return
                if self._point_near_line(QtCore.QPointF(event.pos()), p1, p2, tolerance=7.0):
                    self.drag_line_entry_index = lidx
                    self.drag_line_kind = "separate"
                    self.drag_line_mode = "whole"
                    self.drag_start = QtCore.QPointF(event.pos())
                    self.main_window.selected_line_index = lidx
                    self.main_window.selected_entry_index = -1
                    self.main_window.refresh_selection_only()
                    self.update()
                    return

            for idx in reversed(range(len(self.main_window.entries))):
                entry = self.main_window.entries[idx]
                if entry.pageno != current_page:
                    continue
                rect = self._get_entry_rect_on_screen(entry, color_map)
                resize_handle = QtCore.QRectF(rect.right() - 12, rect.bottom() - 12, 12, 12)
                line_pts = self._line_handles(entry)
                if line_pts and idx == self.main_window.selected_entry_index:
                    p1, p2 = line_pts
                    if QtCore.QLineF(event.pos(), p1).length() <= 8:
                        self.drag_line_entry_index = idx
                        self.drag_line_kind = "legacy"
                        self.drag_line_mode = "start"
                        return
                    if QtCore.QLineF(event.pos(), p2).length() <= 8:
                        self.drag_line_entry_index = idx
                        self.drag_line_kind = "legacy"
                        self.drag_line_mode = "end"
                        return
                    if self._point_near_line(QtCore.QPointF(event.pos()), p1, p2, tolerance=7.0):
                        self.drag_line_entry_index = idx
                        self.drag_line_kind = "legacy"
                        self.drag_line_mode = "whole"
                        self.drag_start = QtCore.QPointF(event.pos())
                        self.main_window.selected_entry_index = idx
                        self.main_window.refresh_selection_only()
                        self.update()
                        return
                if idx == self.main_window.selected_entry_index and resize_handle.contains(QtCore.QPointF(event.pos())):
                    self.resize_entry_index = idx
                    self.main_window.selected_entry_index = idx
                    self.update()
                    return
                if rect.contains(event.pos()):
                    self.drag_entry_index = idx
                    self.dragging = True
                    self.drag_offset = QtCore.QPointF(event.pos()) - rect.topLeft()
                    self.main_window.selected_entry_index = idx
                    self.main_window.selected_line_index = -1
                    self.main_window.refresh_selection_only()
                    self.update()
                    return

        self.last_click_point = event.pos()
        self.main_window.selected_line_index = -1
        self.main_window.store_last_click(event.pos())
        self.update()

        if self.main_window.active_mode == "annotation":
            self.main_window.capture_annotation_point(event.pos())
        elif self.main_window.active_mode == "bookmark":
            self.main_window.capture_bookmark_point()

    def mouseMoveEvent(self, event):
        if not self.main_window:
            return
        if self.drag_entry_index is not None and self.dragging:
            entry = self.main_window.entries[self.drag_entry_index]
            new_top_left = QtCore.QPointF(event.pos()) - self.drag_offset
            new_x = max(0.0, new_top_left.x() / self.main_window.zoom)
            new_y = max(0.0, new_top_left.y() / self.main_window.zoom)
            entry.x1 = round(new_x, 6)
            entry.y1 = round(new_y, 6)
            self.main_window.refresh_annotation_table()
            self.main_window.select_annotation_row_silent(self.drag_entry_index)
            self.update()
            return
        if self.resize_entry_index is not None:
            entry = self.main_window.entries[self.resize_entry_index]
            width = max(MIN_BOX_WIDTH_SINGLE, event.pos().x() / self.main_window.zoom - entry.x1)
            height = max(BOX_HEIGHT_NORMAL, event.pos().y() / self.main_window.zoom - entry.y1)
            entry.box_w = round(width, 6)
            entry.box_h = round(height, 6)
            self.main_window.refresh_annotation_table()
            self.main_window.select_annotation_row_silent(self.resize_entry_index)
            self.update()
            return
        if self.drag_line_entry_index is not None:
            x = max(0.0, event.pos().x() / self.main_window.zoom)
            y = max(0.0, event.pos().y() / self.main_window.zoom)
            if self.drag_line_kind == "separate":
                line = self.main_window.lines[self.drag_line_entry_index]
                if self.drag_line_mode == "start":
                    line.x1 = round(x, 6)
                    line.y1 = round(y, 6)
                elif self.drag_line_mode == "end":
                    line.x2 = round(x, 6)
                    line.y2 = round(y, 6)
                elif self.drag_line_mode == "whole":
                    dx = (event.pos().x() - self.drag_start.x()) / self.main_window.zoom
                    dy = (event.pos().y() - self.drag_start.y()) / self.main_window.zoom
                    line.x1 = round(line.x1 + dx, 6)
                    line.y1 = round(line.y1 + dy, 6)
                    line.x2 = round(line.x2 + dx, 6)
                    line.y2 = round(line.y2 + dy, 6)
                    self.drag_start = QtCore.QPointF(event.pos())
            else:
                entry = self.main_window.entries[self.drag_line_entry_index]
                if self.drag_line_mode == "start":
                    entry.line_x1 = round(x, 6)
                    entry.line_y1 = round(y, 6)
                elif self.drag_line_mode == "end":
                    entry.line_x2 = round(x, 6)
                    entry.line_y2 = round(y, 6)
                elif self.drag_line_mode == "whole":
                    dx = (event.pos().x() - self.drag_start.x()) / self.main_window.zoom
                    dy = (event.pos().y() - self.drag_start.y()) / self.main_window.zoom
                    entry.line_x1 = round((entry.line_x1 or 0) + dx, 6)
                    entry.line_y1 = round((entry.line_y1 or 0) + dy, 6)
                    entry.line_x2 = round((entry.line_x2 or 0) + dx, 6)
                    entry.line_y2 = round((entry.line_y2 or 0) + dy, 6)
                    self.drag_start = QtCore.QPointF(event.pos())
            self.update()
            return

    def mouseReleaseEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.drag_entry_index = None
            self.dragging = False
            self.resize_entry_index = None
            self.drag_line_entry_index = None
            self.drag_line_mode = None
            self.drag_line_kind = None

    def paintEvent(self, event):
        super().paintEvent(event)
        if not self.main_window or self.main_window.page_rect is None:
            return

        painter = QtGui.QPainter(self)
        zoom = self.main_window.zoom
        current_page = self.main_window.current_page_index + 1
        color_map = get_page_domain_color_map(self.main_window.entries, current_page)
        page_width = self.main_window.page_rect.width if self.main_window.page_rect else None

        for lidx, line in enumerate(self.main_window.lines):
            if line.pageno != current_page:
                continue
            line_pen = QtGui.QPen(QtGui.QColor("#c62828"), 2)
            painter.setPen(line_pen)
            p1, p2 = self._connector_line_points(line)
            painter.drawLine(p1, p2)
            if lidx == self.main_window.selected_line_index:
                painter.setBrush(QtGui.QBrush(QtGui.QColor("#c62828")))
                painter.setPen(QtGui.QPen(QtGui.QColor("#7f1d1d"), 1))
                painter.drawEllipse(p1, 5, 5)
                painter.drawEllipse(p2, 5, 5)

        for idx, entry in enumerate(self.main_window.entries):
            if entry.pageno != current_page:
                continue

            layout = compute_entry_layout(entry, color_map, page_width=page_width)
            rect_pdf = rect_from_top_origin(entry.x1, entry.y1, layout["box_w"], layout["box_h"], entry.pageh)
            rect = QtCore.QRectF(rect_pdf.x0 * zoom, rect_pdf.y0 * zoom, rect_pdf.width * zoom, rect_pdf.height * zoom)

            if None not in (entry.line_x1, entry.line_y1, entry.line_x2, entry.line_y2):
                line_pen = QtGui.QPen(QtGui.QColor("#c62828"), 2)
                painter.setPen(line_pen)
                p1 = QtCore.QPointF(entry.line_x1 * zoom, entry.line_y1 * zoom)
                p2 = QtCore.QPointF(entry.line_x2 * zoom, entry.line_y2 * zoom)
                painter.drawLine(p1, p2)
                if idx == self.main_window.selected_entry_index:
                    painter.setBrush(QtGui.QBrush(QtGui.QColor("#c62828")))
                    painter.setPen(QtGui.QPen(QtGui.QColor("#7f1d1d"), 1))
                    painter.drawEllipse(p1, 5, 5)
                    painter.drawEllipse(p2, 5, 5)

            fill_q = qcolor_from_rgb01(layout["fill"])
            fill_q.setAlpha(210)
            painter.fillRect(rect, fill_q)

            pen = QtGui.QPen(QtGui.QColor(60, 72, 88), 1)
            if layout["dashed"]:
                pen.setStyle(QtCore.Qt.DashLine)
            painter.setPen(pen)
            painter.drawRect(rect)

            font = QtGui.QFont("Arial")
            font.setPointSizeF(layout["font_size"] * zoom * 0.75)
            font.setBold(layout["bold"])
            painter.setFont(font)
            painter.setPen(QtGui.QColor(0, 0, 0))

            line_count = max(1, len(layout["lines"]))
            if line_count == 1:
                line_gap = (layout["font_size"] + 0.6) * zoom * 0.75
                visual_factor = 0.72
            elif line_count == 2:
                line_gap = (layout["font_size"] + 0.8) * zoom * 0.75
                visual_factor = 0.88
            else:
                line_gap = (layout["font_size"] + 1.0) * zoom * 0.75
                visual_factor = 0.94

            text_block_height = line_count * line_gap * visual_factor
            y_offset = max(0.0, (rect.height() - text_block_height) / 2.0)
            ty = rect.top() + y_offset + (layout["font_size"] * zoom * 0.75) - 1.0

            for line in layout["lines"]:
                line_indent = max(0.0, estimate_text_width(line) - estimate_text_width(line.lstrip(" "))) * zoom * 0.75
                tx = rect.left() + TEXT_PADDING_X * zoom + line_indent
                painter.drawText(QtCore.QPointF(tx, ty), line.lstrip(" "))
                ty += line_gap

            mx = entry.x1 * zoom
            my = entry.y1 * zoom
            if idx == self.main_window.selected_entry_index:
                sel_pen = QtGui.QPen(QtGui.QColor("#7c3aed"), 3)
                painter.setPen(sel_pen)
                painter.setBrush(QtCore.Qt.NoBrush)
                painter.drawRect(rect.adjusted(-2, -2, 2, 2))
                painter.setBrush(QtGui.QBrush(QtGui.QColor("#facc15")))
                painter.setPen(QtGui.QPen(QtGui.QColor("#854d0e"), 1))
                painter.drawRect(QtCore.QRectF(rect.right() - 12, rect.bottom() - 12, 12, 12))
                painter.setPen(sel_pen)
                painter.drawEllipse(QtCore.QPointF(mx, my), 6, 6)
            else:
                dot_pen = QtGui.QPen(QtGui.QColor("#1d7df2"), 2)
                painter.setPen(dot_pen)
                painter.setBrush(QtGui.QBrush(QtGui.QColor("#7fc2ff")))
                painter.drawEllipse(QtCore.QPointF(mx, my), 4, 4)

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
            is_selected = (0 <= self.main_window.selected_bookmark_index < len(self.main_window.bookmarks) and self.main_window.bookmarks[self.main_window.selected_bookmark_index] == bm)
            f.setBold(is_selected)
            painter.setFont(f)
            offset_x = 8 + (max(1, bm.level) - 1) * 12
            painter.drawText(QtCore.QPointF(offset_x, sy - 4), f"BM L{bm.level}: {bm.title[:45]}")
            painter.setPen(bookmark_pen)

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
        self.setWindowTitle("AnnotateCRF Studio")
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
        self.lines: List[ConnectorLineEntry] = []

        self.selected_entry_index = -1
        self.selected_bookmark_index = -1
        self.selected_line_index = -1

        self.last_click_pdf_x = None
        self.last_click_pdf_y = None
        self.last_click_pageh = None

        self.active_mode = "annotation"
        self.pdf_loaded = False
        self.has_annotation = False
        self.has_bookmark = False

        self._suppress_annotation_selection_signal = False
        self._suppress_bookmark_selection_signal = False
        self.line_capture_stage = None
        self.pending_line_start = None

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

        header = QtWidgets.QLabel("AnnotateCRF Studio")
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
            b.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
            b.setMinimumWidth(max(170, b.fontMetrics().horizontalAdvance(text) + 42))
            return b

        # Top controls
        top_row = QtWidgets.QHBoxLayout()
        top_row.setSpacing(8)

        self.btn_open = mkbtn("Open PDF", "#3b82f6", "#5c9cff")
        self.btn_prev = mkbtn("Previous Page", "#9ca3af", "#b6bcc7", "#1f2937")
        self.btn_next = mkbtn("Next Page", "#6b7280", "#7b8495")

        self.btn_prev.setEnabled(False)
        self.btn_next.setEnabled(False)

        self.page_jump_spin = QtWidgets.QSpinBox()
        self.page_jump_spin.setRange(1, 1)
        self.page_jump_spin.setValue(1)
        self.page_jump_spin.setEnabled(False)
        self.page_jump_spin.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
        self.page_jump_spin.setAlignment(QtCore.Qt.AlignCenter)
        self.page_jump_spin.setFixedWidth(90)
        self.page_jump_spin.setStyleSheet("""
            QSpinBox {
                background: #ffffff; color: #12608d; border: 1px solid #b8cfe4; border-radius: 8px;
                padding: 5px 8px; font-family: 'Times New Roman'; font-size: 11pt; font-weight: bold;
            }
        """)

        self.btn_go_page = mkbtn("Go", "#0284c7", "#0ea5e9")
        self.btn_go_page.setMinimumWidth(80)
        self.btn_go_page.setEnabled(False)

        self.page_info = QtWidgets.QLabel("Page: -")
        self.page_info.setStyleSheet("""
            QLabel {
                background: #e6f2fb; color: #12608d; border-radius: 8px; padding: 5px 10px;
                font-family: 'Times New Roman'; font-size: 11pt; font-weight: bold;
            }
        """)

        self.btn_terminate = mkbtn("Terminate", "#991b1b", "#b91c1c")

        top_row.addWidget(self.btn_open)
        top_row.addWidget(self.btn_prev)
        top_row.addWidget(self.btn_next)
        top_row.addWidget(self.page_jump_spin)
        top_row.addWidget(self.btn_go_page)
        top_row.addStretch()
        top_row.addWidget(self.btn_terminate)
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
        self.annotation_table.setSelectionMode(QtWidgets.QTableWidget.ExtendedSelection)
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

        ann_btn_grid = QtWidgets.QGridLayout()
        ann_btn_grid.setHorizontalSpacing(8)
        ann_btn_grid.setVerticalSpacing(8)

        self.btn_export_ann = mkbtn("Export Annotation CSV", "#2563eb", "#4a7df2")
        self.btn_load_ann = mkbtn("Load Annotation CSV", "#7c3aed", "#9b63f0")
        self.btn_edit_ann = mkbtn("Edit Selected Annotation", "#0f766e", "#159287")
        self.btn_copy_ann = mkbtn("Copy Selected Annotation", "#0369a1", "#0ea5e9")
        self.btn_copy_ann_page = mkbtn("Copy Annotation To Page", "#0f766e", "#159287")
        self.btn_line_mode = mkbtn("Draw Connector Line", "#b45309", "#c97519")
        self.btn_clear_line = mkbtn("Clear Connector Line", "#475569", "#64748b")
        self.btn_delete_ann = mkbtn("Delete Selected Annotation", "#ef4444", "#f87171")

        ann_buttons = [
            self.btn_export_ann, self.btn_load_ann, self.btn_edit_ann, self.btn_copy_ann,
            self.btn_copy_ann_page, self.btn_line_mode, self.btn_clear_line, self.btn_delete_ann
        ]
        for i, btn in enumerate(ann_buttons):
            ann_btn_grid.addWidget(btn, i // 4, i % 4)
        for col in range(4):
            ann_btn_grid.setColumnStretch(col, 1)

        ann_layout.addLayout(ann_btn_grid)
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

        bm_btn_grid = QtWidgets.QGridLayout()
        bm_btn_grid.setHorizontalSpacing(8)
        bm_btn_grid.setVerticalSpacing(8)

        self.btn_export_bm = mkbtn("Export Bookmark CSV", "#2563eb", "#4a7df2")
        self.btn_load_bm = mkbtn("Load Bookmark CSV", "#7c3aed", "#9b63f0")
        self.btn_edit_bm = mkbtn("Edit Selected Bookmark", "#0f766e", "#159287")
        self.btn_copy_bm = mkbtn("Copy Selected Bookmark", "#0369a1", "#0ea5e9")
        self.btn_delete_bm = mkbtn("Delete Selected Bookmark", "#ef4444", "#f87171")

        bm_buttons = [
            self.btn_export_bm, self.btn_load_bm, self.btn_edit_bm, self.btn_copy_bm, self.btn_delete_bm
        ]
        for i, btn in enumerate(bm_buttons):
            bm_btn_grid.addWidget(btn, i // 4, i % 4)
        for col in range(4):
            bm_btn_grid.setColumnStretch(col, 1)

        bm_layout.addLayout(bm_btn_grid)
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
        self.chk_add_toc = QtWidgets.QCheckBox("Add TOC with hyperlinks")
        self.chk_add_toc.setStyleSheet("QCheckBox { font-family: 'Times New Roman'; font-size: 11pt; color: #27496d; padding: 4px 2px; }")

        review_action_row = QtWidgets.QHBoxLayout()
        review_action_row.setSpacing(10)
        review_action_row.addWidget(self.chk_add_toc)
        review_action_row.addStretch()
        review_action_row.addWidget(self.btn_generate_pdf)

        review_layout.addWidget(review_splitter, 1)
        review_layout.addLayout(review_action_row)
        self.main_stack.addWidget(review_page)

        QtCore.QTimer.singleShot(0, self.init_splitter_sizes)

        # Connections
        self.btn_open.clicked.connect(self.open_pdf)
        self.btn_prev.clicked.connect(self.prev_page)
        self.btn_next.clicked.connect(self.next_page)
        self.btn_go_page.clicked.connect(self.go_to_page)
        self.page_jump_spin.editingFinished.connect(self.go_to_page)
        self.btn_terminate.clicked.connect(self.close)

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
        self.btn_edit_ann.clicked.connect(self.edit_selected_annotation)
        self.btn_copy_ann.clicked.connect(self.copy_selected_annotation)
        self.btn_copy_ann_page.clicked.connect(self.copy_annotation_to_page)
        self.btn_edit_bm.clicked.connect(self.edit_selected_bookmark)
        self.btn_copy_bm.clicked.connect(self.copy_selected_bookmark)
        self.btn_delete_ann.clicked.connect(self.delete_selected_annotation)
        self.btn_line_mode.clicked.connect(self.start_connector_line_mode)
        self.btn_clear_line.clicked.connect(self.clear_connector_lines)

        self.btn_export_bm.clicked.connect(self.export_bookmarks_csv)
        self.btn_load_bm.clicked.connect(self.load_bookmarks_csv)
        self.btn_delete_bm.clicked.connect(self.delete_selected_bookmark)

        self.switch_mode("annotation", force=True)
        self.update_annotation_action_buttons()

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
                        border: 3px solid #d1fae5;
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

        style_button(self.btn_annotation, "#315c8f", "#4270a7", self.active_mode == "annotation")
        style_button(self.btn_bookmark, "#6b4f8f", "#7c61a0", self.active_mode == "bookmark")
        style_button(self.btn_review, "#3f7d68", "#548f7b", self.active_mode == "review")

    def switch_mode(self, mode, force=False):
        if not force and mode in ("annotation", "bookmark", "review") and not self.pdf_loaded and mode != "annotation":
            return

        self.active_mode = mode
        if mode != "annotation":
            self.reset_line_capture_state()
        self.set_active_mode_button_styles()

        if mode == "annotation":
            self.main_stack.setCurrentIndex(0)
            self.bottom_stack.setCurrentIndex(0)
            self.tip.setText("Click anywhere on the PDF to add annotation. Drag box, resize from bottom-right handle, or add/edit a connector line for the selected annotation.")
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
        self.page_jump_spin.setEnabled(has_pdf)
        self.btn_go_page.setEnabled(has_pdf)

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
            self.lines = []
            self.selected_entry_index = -1
            self.selected_bookmark_index = -1
            self.selected_line_index = -1
            self.image_label.last_click_point = None
            self.last_click_pdf_x = None
            self.last_click_pdf_y = None
            self.last_click_pageh = None

            self.pdf_loaded = True
            self.has_annotation = False
            self.has_bookmark = False

            self.btn_annotation.setEnabled(True)
            self.btn_bookmark.setEnabled(True)
            self.btn_line_mode.setEnabled(False)
            self.btn_clear_line.setEnabled(False)
            self.page_jump_spin.setRange(1, len(self.doc))
            self.page_jump_spin.setValue(1)
            self.check_review_enable()

            self.refresh_annotation_table()
            self.refresh_bookmark_table()
            self.refresh_review_tables()
            self.render_page()
            self.switch_mode("annotation")
            self.update_annotation_action_buttons()
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
        self.page_jump_spin.blockSignals(True)
        self.page_jump_spin.setValue(self.current_page_index + 1)
        self.page_jump_spin.blockSignals(False)
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

    def go_to_page(self):
        if not self.doc:
            return
        target_page = int(self.page_jump_spin.value())
        target_page = max(1, min(len(self.doc), target_page))
        if self.current_page_index != target_page - 1:
            self.current_page_index = target_page - 1
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

        dlg = AnnotationDialog(self, dialog_title="Add Annotation")
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
            is_not_submitted=vals["is_not_submitted"],
            box_w=None,
            box_h=None,
            line_x1=None,
            line_y1=None,
            line_x2=None,
            line_y2=None
        )

        self.entries.append(entry)
        self.selected_entry_index = len(self.entries) - 1

        self.refresh_annotation_table()
        self.refresh_review_tables()
        self.select_annotation_row_silent(self.selected_entry_index)
        self.image_label.update()

        self.has_annotation = True
        self.check_review_enable()

    def get_bookmark_choice_items(self, exclude_index=None):
        items = []
        for idx, bm in enumerate(self.bookmarks):
            if exclude_index is not None and idx == exclude_index:
                continue
            indent = "    " * max(0, int(bm.level) - 1)
            label = f"L{int(bm.level)} | {indent}{bm.title} (Page {int(bm.pageno)})"
            items.append((idx, label))
        return items

    def get_branch_insert_position(self, bookmark_index):
        if bookmark_index is None:
            return len(self.bookmarks)
        if bookmark_index < 0 or bookmark_index >= len(self.bookmarks):
            return len(self.bookmarks)
        parent_level = max(1, int(self.bookmarks[bookmark_index].level))
        insert_pos = bookmark_index + 1
        while insert_pos < len(self.bookmarks) and int(self.bookmarks[insert_pos].level) > parent_level:
            insert_pos += 1
        return insert_pos

    def insert_bookmark_by_choice(self, bookmark, insert_after_branch):
        if insert_after_branch in (None, "__END__", "__KEEP__"):
            self.bookmarks.append(bookmark)
            return len(self.bookmarks) - 1
        try:
            target_index = int(insert_after_branch)
        except Exception:
            self.bookmarks.append(bookmark)
            return len(self.bookmarks) - 1
        insert_pos = self.get_branch_insert_position(target_index)
        self.bookmarks.insert(insert_pos, bookmark)
        return insert_pos

    def move_existing_bookmark_by_choice(self, current_index, insert_after_branch):
        if current_index < 0 or current_index >= len(self.bookmarks):
            return current_index
        if insert_after_branch in (None, "__KEEP__"):
            return current_index

        bookmark = self.bookmarks.pop(current_index)

        if insert_after_branch == "__END__":
            self.bookmarks.append(bookmark)
            return len(self.bookmarks) - 1

        try:
            target_index = int(insert_after_branch)
        except Exception:
            self.bookmarks.insert(current_index, bookmark)
            return current_index

        if target_index > current_index:
            target_index -= 1

        insert_pos = self.get_branch_insert_position(target_index)
        self.bookmarks.insert(insert_pos, bookmark)
        return insert_pos

    def capture_bookmark_point(self):
        if not self.doc:
            return

        dlg = BookmarkDialog(
            self.current_page_index + 1,
            self,
            bookmark_choices=self.get_bookmark_choice_items(),
            keep_order_default=False
        )
        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return

        vals = dlg.get_values()
        bm = BookmarkEntry(
            title=vals["title"],
            level=vals["level"],
            pageno=vals["pageno"]
        )
        self.selected_bookmark_index = self.insert_bookmark_by_choice(
            bm,
            vals.get("insert_after_branch")
        )

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
        self.update_annotation_action_buttons()

    def get_selected_annotation_rows(self):
        model = self.annotation_table.selectionModel()
        if model is None:
            return []
        rows = sorted({idx.row() for idx in model.selectedRows() if 0 <= idx.row() < len(self.entries)})
        return rows

    def on_annotation_selection_changed(self):
        if self._suppress_annotation_selection_signal:
            return

        selected_rows = self.get_selected_annotation_rows()
        row = selected_rows[0] if selected_rows else self.annotation_table.currentRow()
        if 0 <= row < len(self.entries):
            self.selected_entry_index = row
            self.selected_line_index = -1
            entry = self.entries[row]
            target_page_index = entry.pageno - 1
            if self.doc and target_page_index != self.current_page_index:
                self.current_page_index = target_page_index
                self.render_page()
        else:
            self.selected_entry_index = -1
        self.update_annotation_action_buttons()
        self.image_label.update()

    def on_review_annotation_selection_changed(self):
        row = self.review_annotation_table.currentRow()
        if 0 <= row < len(self.entries):
            self.selected_entry_index = row
            self.selected_line_index = -1
            entry = self.entries[row]
            if self.doc:
                self.current_page_index = max(0, entry.pageno - 1)
                self.render_page()

    def update_annotation_action_buttons(self):
        selected_rows = self.get_selected_annotation_rows()
        selected_count = len(selected_rows)
        has_single_selection = selected_count == 1 and 0 <= self.selected_entry_index < len(self.entries)
        has_any_selection = selected_count > 0 or (0 <= self.selected_entry_index < len(self.entries))
        self.btn_edit_ann.setEnabled(has_single_selection)
        self.btn_delete_ann.setEnabled(has_single_selection)
        self.btn_copy_ann.setEnabled(has_any_selection)
        self.btn_copy_ann_page.setEnabled(bool(self.pdf_loaded and has_any_selection))
        self.btn_line_mode.setEnabled(bool(self.pdf_loaded))
        self.btn_clear_line.setEnabled(bool(self.pdf_loaded and (self.lines or self.selected_line_index >= 0)))

    def reset_line_capture_state(self):
        self.line_capture_stage = None
        self.pending_line_start = None

    def start_connector_line_mode(self):
        if not self.pdf_loaded or not self.doc:
            QtWidgets.QMessageBox.information(self, "Connector Line", "Please open the source PDF first.")
            return
        self.selected_entry_index = -1
        self.selected_line_index = -1
        self.line_capture_stage = "start"
        self.pending_line_start = None
        self.tip.setText("Connector line mode: click START point and then END point on the PDF.")
        self.refresh_selection_only()
        self.image_label.update()

    def clear_connector_lines(self):
        if 0 <= self.selected_line_index < len(self.lines):
            del self.lines[self.selected_line_index]
            self.selected_line_index = -1
        elif self.lines:
            reply = QtWidgets.QMessageBox.question(
                self,
                "Clear Connector Lines",
                "No line is selected. Do you want to clear all connector lines?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                QtWidgets.QMessageBox.No,
            )
            if reply != QtWidgets.QMessageBox.Yes:
                return
            self.lines = []
        else:
            QtWidgets.QMessageBox.information(self, "Connector Line", "No connector lines available to clear.")
            return
        self.reset_line_capture_state()
        self.update_annotation_action_buttons()
        self.image_label.update()

    def capture_connector_line_point(self, point: QtCore.QPoint):
        if self.line_capture_stage is None:
            return
        x = max(0, min(point.x() / self.zoom, self.page_rect.width))
        y = max(0, min(point.y() / self.zoom, self.page_rect.height))
        if self.line_capture_stage == "start":
            self.pending_line_start = (round(x, 6), round(y, 6))
            self.line_capture_stage = "end"
            self.tip.setText("Connector line mode: click END point on the PDF.")
        else:
            sx, sy = self.pending_line_start or (round(x, 6), round(y, 6))
            self.lines.append(
                ConnectorLineEntry(
                    pageno=self.current_page_index + 1,
                    x1=sx,
                    y1=sy,
                    x2=round(x, 6),
                    y2=round(y, 6),
                )
            )
            self.selected_line_index = len(self.lines) - 1
            self.reset_line_capture_state()
            self.tip.setText("Click anywhere on the PDF to add annotation or bookmark, based on the selected mode.")
            self.update_annotation_action_buttons()
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

    def annotation_entry_to_dialog_data(self, entry):
        return {
            "domain": entry.domain,
            "name": entry.name,
            "annotation": entry.annotation,
            "assignedfield": entry.assignedfield,
            "is_domain_annotation": entry.is_domain_annotation,
            "is_assigned_field": entry.is_assigned_field,
            "is_not_submitted": entry.is_not_submitted,
        }

    def edit_selected_annotation(self):
        row = self.annotation_table.currentRow()
        if row < 0 or row >= len(self.entries):
            QtWidgets.QMessageBox.information(self, "Edit Annotation", "Please select an annotation row to edit.")
            return

        entry = self.entries[row]
        dlg = AnnotationDialog(self, initial_data=self.annotation_entry_to_dialog_data(entry), dialog_title="Edit Annotation")
        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return

        vals = dlg.get_values()
        entry.domain = vals["domain"]
        entry.name = vals["name"]
        entry.annotation = vals["annotation"]
        entry.assignedfield = vals["assignedfield"]
        entry.is_domain_annotation = vals["is_domain_annotation"]
        entry.is_assigned_field = vals["is_assigned_field"]
        entry.is_not_submitted = vals["is_not_submitted"]
        entry.box_w = None
        entry.box_h = None
        entry.line_x1 = None
        entry.line_y1 = None
        entry.line_x2 = None
        entry.line_y2 = None

        self.selected_entry_index = row
        self.refresh_annotation_table()
        self.refresh_review_tables()
        self.select_annotation_row_silent(row)
        self.image_label.update()
        self.has_annotation = len(self.entries) > 0
        self.check_review_enable()

    def copy_selected_annotation(self):
        selected_rows = self.get_selected_annotation_rows()
        if not selected_rows and 0 <= self.annotation_table.currentRow() < len(self.entries):
            selected_rows = [self.annotation_table.currentRow()]
        if not selected_rows:
            QtWidgets.QMessageBox.information(self, "Copy Annotation", "Please select one or more annotation rows to copy.")
            return

        if len(selected_rows) == 1:
            row = selected_rows[0]
            src = self.entries[row]
            dlg = AnnotationDialog(self, initial_data=self.annotation_entry_to_dialog_data(src), dialog_title="Copy Annotation")
            if dlg.exec_() != QtWidgets.QDialog.Accepted:
                return

            vals = dlg.get_values()
            x1 = round(src.x1 + 8.0, 6)
            y1 = round(src.y1 + 8.0, 6)
            if self.page_rect:
                x1 = round(min(x1, max(0.0, self.page_rect.width - MIN_BOX_WIDTH_SINGLE - RIGHT_PAGE_MARGIN)), 6)
                y1 = round(min(y1, max(0.0, self.page_rect.height - BOX_HEIGHT_NORMAL - 6.0)), 6)

            new_entry = AnnotationEntry(
                domain=vals["domain"],
                name=vals["name"],
                pageno=src.pageno,
                annotation=vals["annotation"],
                assignedfield=vals["assignedfield"],
                x1=x1,
                y1=y1,
                pageh=src.pageh,
                is_domain_annotation=vals["is_domain_annotation"],
                is_assigned_field=vals["is_assigned_field"],
                is_not_submitted=vals["is_not_submitted"],
                box_w=None,
                box_h=None,
                line_x1=None,
                line_y1=None,
                line_x2=None,
                line_y2=None
            )
            self.entries.append(new_entry)
            self.selected_entry_index = len(self.entries) - 1
            self.refresh_annotation_table()
            self.refresh_review_tables()
            self.select_annotation_row_silent(self.selected_entry_index)
            self.image_label.update()
            self.has_annotation = True
            self.check_review_enable()
            return

        added_rows = []
        per_copy_shift = 8.0
        for offset_index, row in enumerate(selected_rows, start=1):
            src = self.entries[row]
            x1 = float(src.x1) + (per_copy_shift * offset_index)
            y1 = float(src.y1) + (per_copy_shift * offset_index)
            page_w = float(self.doc[src.pageno - 1].rect.width) if self.doc else (self.page_rect.width if self.page_rect else 999999)
            page_h = float(self.doc[src.pageno - 1].rect.height) if self.doc else src.pageh
            x1 = round(min(max(0.0, x1), max(0.0, page_w - MIN_BOX_WIDTH_SINGLE - RIGHT_PAGE_MARGIN)), 6)
            y1 = round(min(max(0.0, y1), max(0.0, page_h - BOX_HEIGHT_NORMAL - 6.0)), 6)

            new_entry = AnnotationEntry(
                domain=src.domain,
                name=src.name,
                pageno=src.pageno,
                annotation=src.annotation,
                assignedfield=src.assignedfield,
                x1=x1,
                y1=y1,
                pageh=src.pageh,
                is_domain_annotation=src.is_domain_annotation,
                is_assigned_field=src.is_assigned_field,
                is_not_submitted=src.is_not_submitted,
                box_w=src.box_w,
                box_h=src.box_h,
                line_x1=None,
                line_y1=None,
                line_x2=None,
                line_y2=None
            )
            self.entries.append(new_entry)
            added_rows.append(len(self.entries) - 1)

        self.selected_entry_index = added_rows[-1] if added_rows else -1
        self.refresh_annotation_table()
        self.refresh_review_tables()
        if self.selected_entry_index >= 0:
            self.select_annotation_row_silent(self.selected_entry_index)
        self.image_label.update()
        self.has_annotation = len(self.entries) > 0
        self.check_review_enable()
        QtWidgets.QMessageBox.information(self, "Copy Annotation", f"{len(added_rows)} annotations copied on their current pages.")

    def copy_annotation_to_page(self):
        selected_rows = self.get_selected_annotation_rows()
        if not selected_rows and 0 <= self.annotation_table.currentRow() < len(self.entries):
            selected_rows = [self.annotation_table.currentRow()]
        if not selected_rows:
            QtWidgets.QMessageBox.information(self, "Copy Annotation To Page", "Please select one or more annotation rows to copy.")
            return
        if not self.doc:
            QtWidgets.QMessageBox.information(self, "Copy Annotation To Page", "Please open a PDF first.")
            return

        sources = [self.entries[row] for row in selected_rows]
        src = sources[0]
        vals = None
        if len(sources) == 1:
            dlg = AnnotationDialog(self, initial_data=self.annotation_entry_to_dialog_data(src), dialog_title="Copy Annotation To Page")
            if dlg.exec_() != QtWidgets.QDialog.Accepted:
                return
            vals = dlg.get_values()

        opt = QtWidgets.QDialog(self)
        opt.setWindowTitle("Copy Annotation To Page")
        opt.setModal(True)
        opt.resize(460, 260)
        opt.setStyleSheet("""
            QDialog { background-color: #f4f8ff; border-radius: 14px; }
            QLabel { font-family: 'Times New Roman'; font-size: 12pt; color: #1c2e4a; }
            QSpinBox {
                background: #ffffff; border: 1px solid #a8bfdc; border-radius: 8px;
                padding: 6px 8px; font-family: 'Times New Roman'; font-size: 12pt; color: #1a1a1a;
            }
            QCheckBox { font-family: 'Times New Roman'; font-size: 12pt; color: #143b66; spacing: 8px; }
            QDialogButtonBox QPushButton {
                font-family: 'Times New Roman'; font-size: 12pt; font-weight: bold;
                border-radius: 10px; padding: 8px 18px; min-width: 100px;
            }
        """)

        vbox = QtWidgets.QVBoxLayout(opt)
        title = QtWidgets.QLabel("Copy Annotation To Another Page")
        title.setAlignment(QtCore.Qt.AlignCenter)
        title.setStyleSheet("""
            QLabel {
                background: #d9f99d; color: #1f4d0f; font-family: 'Times New Roman';
                font-size: 14pt; font-weight: bold; padding: 10px; border-radius: 12px;
            }
        """)
        vbox.addWidget(title)

        form = QtWidgets.QFormLayout()
        form.setLabelAlignment(QtCore.Qt.AlignRight)
        form.setHorizontalSpacing(18)
        form.setVerticalSpacing(14)

        target_page_spin = QtWidgets.QSpinBox()
        target_page_spin.setRange(1, len(self.doc))
        target_page_spin.setValue(src.pageno)

        keep_pos_chk = QtWidgets.QCheckBox("Keep same position")
        keep_pos_chk.setChecked(True)
        offset_chk = QtWidgets.QCheckBox("Apply small offset when target page is same")
        offset_chk.setChecked(True)

        form.addRow("Target Page", target_page_spin)
        form.addRow("", keep_pos_chk)
        form.addRow("", offset_chk)
        vbox.addLayout(form)

        if len(sources) > 1:
            multi_note = QtWidgets.QLabel(f"{len(sources)} selected annotations will be copied together with their own text and metadata.")
            multi_note.setWordWrap(True)
            multi_note.setStyleSheet("QLabel { color: #22543d; font-size: 11pt; font-style: italic; }")
            vbox.addWidget(multi_note)

        note = QtWidgets.QLabel("Connector lines are not copied automatically. You can draw a new line after copying if needed.")
        note.setWordWrap(True)
        note.setStyleSheet("QLabel { color: #556b84; font-size: 11pt; font-style: italic; }")
        vbox.addWidget(note)

        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        buttons.accepted.connect(opt.accept)
        buttons.rejected.connect(opt.reject)
        ok_btn = buttons.button(QtWidgets.QDialogButtonBox.Ok)
        cancel_btn = buttons.button(QtWidgets.QDialogButtonBox.Cancel)
        ok_btn.setStyleSheet("QPushButton { background-color: #55c16d; color: white; } QPushButton:hover { background-color: #73d789; }")
        cancel_btn.setStyleSheet("QPushButton { background-color: #f16a6a; color: white; } QPushButton:hover { background-color: #f48f8f; }")
        vbox.addWidget(buttons)

        if opt.exec_() != QtWidgets.QDialog.Accepted:
            return

        target_page = int(target_page_spin.value())
        target_page_obj = self.doc[target_page - 1]
        target_page_w = float(target_page_obj.rect.width)
        target_page_h = float(target_page_obj.rect.height)

        added_rows = []
        for offset_index, src in enumerate(sources, start=1):
            if keep_pos_chk.isChecked():
                new_x1 = float(src.x1)
                new_y1 = float(src.y1)
                if target_page == src.pageno and offset_chk.isChecked():
                    new_x1 += 8.0 * offset_index
                    new_y1 += 8.0 * offset_index
            else:
                new_x1 = 40.0 + (8.0 * (offset_index - 1))
                new_y1 = 40.0 + (8.0 * (offset_index - 1))

            max_w = max(0.0, target_page_w - MIN_BOX_WIDTH_SINGLE - RIGHT_PAGE_MARGIN)
            max_h = max(0.0, target_page_h - BOX_HEIGHT_NORMAL - 6.0)
            new_x1 = round(min(max(0.0, new_x1), max_w), 6)
            new_y1 = round(min(max(0.0, new_y1), max_h), 6)

            if vals is not None:
                domain = vals["domain"]
                name = vals["name"]
                annotation = vals["annotation"]
                assignedfield = vals["assignedfield"]
                is_domain_annotation = vals["is_domain_annotation"]
                is_assigned_field = vals["is_assigned_field"]
                is_not_submitted = vals["is_not_submitted"]
            else:
                domain = src.domain
                name = src.name
                annotation = src.annotation
                assignedfield = src.assignedfield
                is_domain_annotation = src.is_domain_annotation
                is_assigned_field = src.is_assigned_field
                is_not_submitted = src.is_not_submitted

            new_entry = AnnotationEntry(
                domain=domain,
                name=name,
                pageno=target_page,
                annotation=annotation,
                assignedfield=assignedfield,
                x1=new_x1,
                y1=new_y1,
                pageh=target_page_h,
                is_domain_annotation=is_domain_annotation,
                is_assigned_field=is_assigned_field,
                is_not_submitted=is_not_submitted,
                box_w=src.box_w,
                box_h=src.box_h,
                line_x1=None,
                line_y1=None,
                line_x2=None,
                line_y2=None
            )
            self.entries.append(new_entry)
            added_rows.append(len(self.entries) - 1)

        self.selected_entry_index = added_rows[-1] if added_rows else -1
        self.refresh_annotation_table()
        self.refresh_review_tables()
        if self.selected_entry_index >= 0:
            self.select_annotation_row_silent(self.selected_entry_index)
        self.has_annotation = True
        self.check_review_enable()

        if self.current_page_index != target_page - 1:
            self.current_page_index = target_page - 1
            self.render_page()
        else:
            self.image_label.update()

        if len(added_rows) == 1:
            QtWidgets.QMessageBox.information(self, "Copy Annotation To Page", f"Annotation copied to page {target_page}.")
        else:
            QtWidgets.QMessageBox.information(self, "Copy Annotation To Page", f"{len(added_rows)} annotations copied to page {target_page}.")

    def bookmark_entry_to_dialog_data(self, entry):
        return {
            "title": entry.title,
            "level": entry.level,
            "pageno": entry.pageno,
            "insert_after_branch": "__KEEP__",
        }

    def edit_selected_bookmark(self):
        row = self.bookmark_table.currentRow()
        if row < 0 or row >= len(self.bookmarks):
            QtWidgets.QMessageBox.information(self, "Edit Bookmark", "Please select a bookmark row to edit.")
            return

        bm = self.bookmarks[row]
        dlg = BookmarkDialog(
            bm.pageno,
            self,
            initial_data=self.bookmark_entry_to_dialog_data(bm),
            dialog_title="Edit Bookmark",
            bookmark_choices=self.get_bookmark_choice_items(exclude_index=row),
            keep_order_default=True
        )
        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return

        vals = dlg.get_values()
        bm.title = vals["title"]
        bm.level = vals["level"]
        bm.pageno = vals["pageno"]

        self.selected_bookmark_index = self.move_existing_bookmark_by_choice(
            row,
            vals.get("insert_after_branch")
        )
        self.refresh_bookmark_table()
        self.refresh_review_tables()
        self.select_bookmark_row_silent(row)
        if self.doc:
            self.current_page_index = max(0, min(len(self.doc) - 1, bm.pageno - 1))
            self.render_page()
        else:
            self.image_label.update()
        self.has_bookmark = len(self.bookmarks) > 0
        self.check_review_enable()

    def copy_selected_bookmark(self):
        row = self.bookmark_table.currentRow()
        if row < 0 or row >= len(self.bookmarks):
            QtWidgets.QMessageBox.information(self, "Copy Bookmark", "Please select a bookmark row to copy.")
            return

        src = self.bookmarks[row]
        initial = self.bookmark_entry_to_dialog_data(src).copy()
        initial["pageno"] = src.pageno
        dlg = BookmarkDialog(src.pageno, self, initial_data=initial, dialog_title="Copy Bookmark")
        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return

        vals = dlg.get_values()
        new_bm = BookmarkEntry(
            title=vals["title"],
            level=vals["level"],
            pageno=vals["pageno"]
        )
        self.bookmarks.append(new_bm)
        self.selected_bookmark_index = len(self.bookmarks) - 1
        self.refresh_bookmark_table()
        self.refresh_review_tables()
        self.select_bookmark_row_silent(self.selected_bookmark_index)
        self.image_label.update()
        self.has_bookmark = True
        self.check_review_enable()

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
        self.update_annotation_action_buttons()

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
    def export_annotations_csv(self, out_path=None, show_message=True, page_offset=0, create_blank_if_none=True):
        if not self.entries and not self.lines and not create_blank_if_none:
            if show_message:
                QtWidgets.QMessageBox.information(self, "Export Annotation CSV", "No annotations available to export.")
            return False

        if out_path is None:
            default_path = os.path.join(
                os.path.dirname(self.open_pdf_path) if self.open_pdf_path else SCRIPT_DIR,
                "annotation_entries.csv"
            )
            out_path, _ = QtWidgets.QFileDialog.getSaveFileName(
                self, "Export Annotation CSV", default_path, "CSV Files (*.csv)"
            )
            if not out_path:
                return False

        try:
            with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=ANNOTATION_CSV_COLUMNS, quoting=csv.QUOTE_ALL)
                writer.writeheader()
                for e in self.entries:
                    writer.writerow({
                        "TYPE": "ANNOTATION",
                        "DOMAIN": e.domain,
                        "NAME": e.name,
                        "PAGENO": int(e.pageno) + int(page_offset),
                        "ANNOTATION": adjust_page_refs_in_text(e.annotation, page_offset),
                        "ASSIGNEDFIELD": e.assignedfield,
                        "X1": e.x1,
                        "Y1": e.y1,
                        "PAGEH": e.pageh,
                        "BOX_W": "" if e.box_w is None else e.box_w,
                        "BOX_H": "" if e.box_h is None else e.box_h,
                        "LINE_PAGENO": "",
                        "LINE_X1": "" if e.line_x1 is None else e.line_x1,
                        "LINE_Y1": "" if e.line_y1 is None else e.line_y1,
                        "LINE_X2": "" if e.line_x2 is None else e.line_x2,
                        "LINE_Y2": "" if e.line_y2 is None else e.line_y2,
                    })
                for ln in self.lines:
                    writer.writerow({
                        "TYPE": "LINE",
                        "DOMAIN": "",
                        "NAME": "",
                        "PAGENO": "",
                        "ANNOTATION": "",
                        "ASSIGNEDFIELD": "",
                        "X1": "",
                        "Y1": "",
                        "PAGEH": "",
                        "BOX_W": "",
                        "BOX_H": "",
                        "LINE_PAGENO": int(ln.pageno) + int(page_offset),
                        "LINE_X1": ln.x1,
                        "LINE_Y1": ln.y1,
                        "LINE_X2": ln.x2,
                        "LINE_Y2": ln.y2,
                    })
            if show_message:
                QtWidgets.QMessageBox.information(self, "Export Annotation CSV", f"Annotation CSV exported successfully:\n{out_path}")
            return True
        except Exception as e:
            if show_message:
                QtWidgets.QMessageBox.critical(self, "Export Annotation CSV", f"Failed to export annotation CSV:\n{e}")
            raise

    def get_pdf_page_rect(self, pageno: int):
        if not self.doc or pageno < 1 or pageno > len(self.doc):
            return None
        try:
            return self.doc[pageno - 1].rect
        except Exception:
            return None

    def get_default_annotation_position(self, pageno: int, loaded_entries=None):
        """
        Return a default (x1, y1, pageh) for imported annotations that do not
        provide coordinates. Stacks boxes vertically on the page so that they
        can be reviewed and dragged later.
        """
        rect = self.get_pdf_page_rect(pageno)
        if rect is None:
            pagew = 595.0
            pageh = 842.0
        else:
            pagew = float(rect.width)
            pageh = float(rect.height)

        existing = loaded_entries if loaded_entries is not None else self.entries
        same_page_count = sum(1 for e in existing if e.pageno == pageno)

        default_x = min(36.0, max(12.0, pagew - 140.0))
        default_y = 36.0 + (same_page_count * 26.0)

        max_y = max(12.0, pageh - 80.0)
        if default_y > max_y:
            default_y = 36.0 + ((same_page_count % 10) * 26.0)

        return round(default_x, 6), round(default_y, 6), round(pageh, 6)

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
            existing_entries = list(self.entries)
            existing_lines = list(self.lines)
            loaded_entries = list(existing_entries)
            loaded_lines = list(existing_lines)

            entry_keys = {annotation_entry_key(e) for e in loaded_entries}
            line_keys = {connector_line_key(l) for l in loaded_lines}

            added_entry_count = 0
            duplicate_entry_count = 0
            added_line_count = 0
            duplicate_line_count = 0
            missing_position_count = 0

            with open(in_path, "r", newline="", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                actual = set([c.strip().upper() for c in (reader.fieldnames or [])])

                required_min = {"DOMAIN", "NAME", "ANNOTATION"}
                missing = required_min - actual
                if missing:
                    raise ValueError(
                        "Annotation CSV is missing required columns.\n"
                        f"Minimum required: {", ".join(sorted(required_min))}\n"
                        f"Missing: {", ".join(sorted(missing))}"
                    )

                for row in reader:
                    row_type = (row.get("TYPE") or "ANNOTATION").strip().upper()

                    if row_type == "LINE":
                        line_pageno = int(clean_number(row.get("LINE_PAGENO"), 0) or 0)
                        line_x1 = clean_number(row.get("LINE_X1"), None)
                        line_y1 = clean_number(row.get("LINE_Y1"), None)
                        line_x2 = clean_number(row.get("LINE_X2"), None)
                        line_y2 = clean_number(row.get("LINE_Y2"), None)

                        if line_pageno >= 1 and None not in (line_x1, line_y1, line_x2, line_y2):
                            new_line = ConnectorLineEntry(
                                pageno=line_pageno,
                                x1=line_x1,
                                y1=line_y1,
                                x2=line_x2,
                                y2=line_y2
                            )
                            key = connector_line_key(new_line)
                            if key in line_keys:
                                duplicate_line_count += 1
                            else:
                                loaded_lines.append(new_line)
                                line_keys.add(key)
                                added_line_count += 1
                        continue

                    domain = (row.get("DOMAIN") or "").strip()
                    name = (row.get("NAME") or "").strip()
                    annotation = (row.get("ANNOTATION") or "").replace("	", "    ").rstrip()
                    assignedfield = (row.get("ASSIGNEDFIELD") or "").strip()

                    if not domain and not name and not annotation:
                        continue

                    pageno = int(clean_number(row.get("PAGENO"), 0) or 0)
                    if pageno < 1:
                        pageno = self.current_page_index + 1
                    elif self.doc and pageno > len(self.doc):
                        pageno = len(self.doc)

                    raw_x1 = clean_number(row.get("X1"), None)
                    raw_y1 = clean_number(row.get("Y1"), None)
                    raw_pageh = clean_number(row.get("PAGEH"), None)

                    box_w = clean_number(row.get("BOX_W"), None)
                    box_h = clean_number(row.get("BOX_H"), None)

                    line_x1 = clean_number(row.get("LINE_X1"), None)
                    line_y1 = clean_number(row.get("LINE_Y1"), None)
                    line_x2 = clean_number(row.get("LINE_X2"), None)
                    line_y2 = clean_number(row.get("LINE_Y2"), None)

                    if raw_x1 is None or raw_y1 is None or raw_pageh is None:
                        x1, y1, pageh = self.get_default_annotation_position(
                            pageno, loaded_entries=loaded_entries
                        )
                        missing_position_count += 1
                    else:
                        rect = self.get_pdf_page_rect(pageno)
                        if rect is None:
                            pagew = 595.0
                            pageh_actual = raw_pageh
                        else:
                            pagew = float(rect.width)
                            pageh_actual = float(rect.height)

                        x1 = max(0.0, min(float(raw_x1), max(0.0, pagew - 20.0)))
                        y1 = max(0.0, min(float(raw_y1), max(0.0, pageh_actual - 20.0)))
                        pageh = float(raw_pageh) if raw_pageh is not None else pageh_actual

                    is_domain_annotation, is_assigned_field, is_not_submitted = bool_from_entry(
                        domain, name, annotation, assignedfield
                    )

                    new_entry = AnnotationEntry(
                        domain=domain,
                        name=name,
                        pageno=pageno,
                        annotation=annotation,
                        assignedfield=assignedfield,
                        x1=round(x1, 6),
                        y1=round(y1, 6),
                        pageh=round(pageh, 6),
                        is_domain_annotation=is_domain_annotation,
                        is_assigned_field=is_assigned_field,
                        is_not_submitted=is_not_submitted,
                        box_w=box_w,
                        box_h=box_h,
                        line_x1=line_x1,
                        line_y1=line_y1,
                        line_x2=line_x2,
                        line_y2=line_y2
                    )

                    entry_key = annotation_entry_key(new_entry)
                    if entry_key in entry_keys:
                        duplicate_entry_count += 1
                        continue

                    loaded_entries.append(new_entry)
                    entry_keys.add(entry_key)
                    added_entry_count += 1

                    if None not in (line_x1, line_y1, line_x2, line_y2):
                        embedded_line = ConnectorLineEntry(
                            pageno=pageno,
                            x1=line_x1,
                            y1=line_y1,
                            x2=line_x2,
                            y2=line_y2
                        )
                        line_key = connector_line_key(embedded_line)
                        if line_key in line_keys:
                            duplicate_line_count += 1
                        else:
                            loaded_lines.append(embedded_line)
                            line_keys.add(line_key)
                            added_line_count += 1

            for e in loaded_entries:
                if None not in (e.line_x1, e.line_y1, e.line_x2, e.line_y2):
                    e.line_x1 = None
                    e.line_y1 = None
                    e.line_x2 = None
                    e.line_y2 = None

            self.entries = loaded_entries
            self.lines = loaded_lines

            self.selected_entry_index = -1
            self.selected_line_index = -1
            self.refresh_annotation_table()
            self.refresh_review_tables()
            self.image_label.update()

            self.has_annotation = len(self.entries) > 0
            self.check_review_enable()

            msg = (
                f"Annotations added: {added_entry_count}\n"
                f"Duplicate annotation rows skipped: {duplicate_entry_count}\n"
                f"Connector lines added: {added_line_count}\n"
                f"Duplicate connector lines skipped: {duplicate_line_count}\n"
                f"Total annotations: {len(self.entries)}\n"
                f"Total connector lines: {len(self.lines)}"
            )
            if missing_position_count:
                msg += f"\n\n{missing_position_count} row(s) had blank position values and were placed at default locations. You can drag them after loading."
            QtWidgets.QMessageBox.information(self, "Load Annotation CSV", msg)

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Load Annotation CSV", f"Failed to load annotation CSV:\n{e}")

    def export_bookmarks_csv(self, out_path=None, show_message=True, page_offset=0, create_blank_if_none=True):
        if not self.bookmarks and not create_blank_if_none:
            if show_message:
                QtWidgets.QMessageBox.information(self, "Export Bookmark CSV", "No bookmarks available to export.")
            return False

        if out_path is None:
            default_path = os.path.join(
                os.path.dirname(self.open_pdf_path) if self.open_pdf_path else SCRIPT_DIR,
                "bookmark_entries.csv"
            )
            out_path, _ = QtWidgets.QFileDialog.getSaveFileName(
                self, "Export Bookmark CSV", default_path, "CSV Files (*.csv)"
            )
            if not out_path:
                return False

        try:
            with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=BOOKMARK_CSV_COLUMNS)
                writer.writeheader()
                for b in self.bookmarks:
                    writer.writerow({
                        "TITLE": b.title,
                        "LEVEL": b.level,
                        "PAGENO": int(b.pageno) + int(page_offset),
                    })
            if show_message:
                QtWidgets.QMessageBox.information(self, "Export Bookmark CSV", f"Bookmark CSV exported successfully:\n{out_path}")
            return True
        except Exception as e:
            if show_message:
                QtWidgets.QMessageBox.critical(self, "Export Bookmark CSV", f"Failed to export bookmark CSV:\n{e}")
            raise

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
            existing_bookmarks = list(self.bookmarks)
            loaded_bookmarks = list(existing_bookmarks)
            bookmark_keys = {bookmark_entry_key(b) for b in loaded_bookmarks}
            added_bookmark_count = 0
            duplicate_bookmark_count = 0

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
                    level = int(clean_number(row.get("LEVEL"), 1) or 1)
                    pageno = int(clean_number(row.get("PAGENO"), 1) or 1)

                    new_bookmark = BookmarkEntry(
                        title=title,
                        level=level,
                        pageno=pageno
                    )
                    key = bookmark_entry_key(new_bookmark)
                    if key in bookmark_keys:
                        duplicate_bookmark_count += 1
                    else:
                        loaded_bookmarks.append(new_bookmark)
                        bookmark_keys.add(key)
                        added_bookmark_count += 1

            self.bookmarks = loaded_bookmarks
            self.selected_bookmark_index = -1
            self.refresh_bookmark_table()
            self.refresh_review_tables()
            self.image_label.update()

            self.has_bookmark = len(self.bookmarks) > 0
            self.check_review_enable()

            QtWidgets.QMessageBox.information(
                self,
                "Load Bookmark CSV",
                f"Bookmarks added: {added_bookmark_count}\n"
                f"Duplicate bookmark rows skipped: {duplicate_bookmark_count}\n"
                f"Total bookmarks: {len(self.bookmarks)}"
            )
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Load Bookmark CSV", f"Failed to load bookmark CSV:\n{e}")

    def export_variables_csv(self, out_path=None, show_message=True, page_offset=0, create_blank_if_none=True):
        variable_rows = OrderedDict()

        for e in self.entries:
            domain = (e.domain or "").strip()
            name = (e.name or "").strip()

            if not domain or not name:
                continue
            if domain.strip().upper() == "REF" and name.strip().upper() == "REF":
                continue
            if e.is_domain_annotation:
                continue
            if e.is_assigned_field:
                continue
            if e.is_not_submitted:
                continue

            key = (domain.upper(), name.upper())
            if key not in variable_rows:
                variable_rows[key] = {
                    "DOMAIN": domain,
                    "VARIABLE": name,
                    "pages": set()
                }
            variable_rows[key]["pages"].add(int(e.pageno) + int(page_offset))

        if not variable_rows and not create_blank_if_none:
            if show_message:
                QtWidgets.QMessageBox.information(self, "Export Variables CSV", "No variable annotations available to export.")
            return False

        if out_path is None:
            default_path = os.path.join(
                os.path.dirname(self.open_pdf_path) if self.open_pdf_path else SCRIPT_DIR,
                "variables.csv"
            )
            out_path, _ = QtWidgets.QFileDialog.getSaveFileName(
                self, "Export Variables CSV", default_path, "CSV Files (*.csv)"
            )
            if not out_path:
                return False

        try:
            with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=["DOMAIN", "VARIABLE", "PAGES"])
                writer.writeheader()
                for _, item in variable_rows.items():
                    pages = sorted(item["pages"])
                    writer.writerow({
                        "DOMAIN": item["DOMAIN"],
                        "VARIABLE": item["VARIABLE"],
                        "PAGES": ",".join(str(p) for p in pages)
                    })
            if show_message:
                QtWidgets.QMessageBox.information(self, "Export Variables CSV", f"Variables CSV exported successfully:\n{out_path}")
            return True
        except Exception as e:
            if show_message:
                QtWidgets.QMessageBox.critical(self, "Export Variables CSV", f"Failed to export variables CSV:\n{e}")
            raise

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

        temp_base_pdf = None
        try:
            QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
            doc = fitz.open(self.open_pdf_path)

            pdf_base = get_pdf_base_output_path(output_pdf)
            annotation_csv_path = pdf_base + "_annotation.csv"
            bookmark_csv_path = pdf_base + "_bookmarks.csv"
            variables_csv_path = pdf_base + "_variables.csv"

            add_toc_checked = bool(self.chk_add_toc.isChecked() and self.bookmarks)
            toc_page_offset = 0
            if add_toc_checked:
                toc_page_offset = compute_toc_page_count_from_bookmarks(self.bookmarks, font_size=12, lines_per_page=38)

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

                    annotation_text_for_output = adjust_page_refs_in_text(e.annotation, toc_page_offset)
                    layout_for_output = dict(layout)
                    layout_for_output["lines"] = wrap_text_by_width(
                        annotation_text_for_output,
                        layout["box_w"],
                        scale=(DOMAIN_FONT_SIZE / BASE_FONT_SIZE if layout["bold"] else 1.0)
                    )
                    output_base_h = BOX_HEIGHT_DOMAIN if layout["bold"] else BOX_HEIGHT_NORMAL
                    output_box_h = compute_output_box_height(
                        layout_for_output["lines"],
                        layout["box_w"],
                        layout["bold"],
                        layout["font_size"],
                        output_base_h
                    )
                    rect = rect_from_top_origin(e.x1, e.y1, layout["box_w"], output_box_h, e.pageh)

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
                        text_lines=layout_for_output["lines"],
                        fill_color=layout["fill"],
                        bold=layout["bold"],
                        dashed=layout["dashed"],
                        font_size=layout["font_size"],
                        link_target_page=extract_page_reference(e.annotation)
                    )

                    if None not in (e.line_x1, e.line_y1, e.line_x2, e.line_y2):
                        page.draw_line(
                            fitz.Point(e.line_x1, e.line_y1),
                            fitz.Point(e.line_x2, e.line_y2),
                            color=(1, 0, 0),
                            width=1.2,
                            overlay=True
                        )

            if self.lines:
                for ln in self.lines:
                    if ln.pageno < 1 or ln.pageno > len(doc):
                        continue
                    page = doc[ln.pageno - 1]
                    page.draw_line(
                        fitz.Point(ln.x1, ln.y1),
                        fitz.Point(ln.x2, ln.y2),
                        color=(1, 0, 0),
                        width=1.2,
                        overlay=True
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

            final_pdf_path = output_pdf
            toc_added = False

            if add_toc_checked:
                temp_base_pdf = pdf_base + "__base_no_toc_tmp__.pdf"
                if os.path.exists(temp_base_pdf):
                    try:
                        os.remove(temp_base_pdf)
                    except Exception:
                        pass
                doc.save(temp_base_pdf, garbage=4, deflate=True)
                doc.close()
                create_clickable_toc_pdf(temp_base_pdf, output_pdf, 12)
                final_pdf_path = output_pdf
                toc_added = toc_page_offset > 0
                try:
                    if os.path.exists(temp_base_pdf):
                        os.remove(temp_base_pdf)
                except Exception:
                    pass
                temp_base_pdf = None
            else:
                doc.save(output_pdf, garbage=4, deflate=True)
                doc.close()

            # Always create all 3 CSV files so they match the actual final PDF.
            self.export_annotations_csv(
                annotation_csv_path,
                show_message=False,
                page_offset=toc_page_offset,
                create_blank_if_none=True
            )
            self.export_bookmarks_csv(
                bookmark_csv_path,
                show_message=False,
                page_offset=toc_page_offset,
                create_blank_if_none=True
            )
            self.export_variables_csv(
                variables_csv_path,
                show_message=False,
                page_offset=toc_page_offset,
                create_blank_if_none=True
            )

            what_written = []
            if self.entries:
                what_written.append(f"{len(self.entries)} annotation(s)")
            if self.bookmarks:
                what_written.append(f"{len(self.bookmarks)} bookmark(s)")
            what_written.append("3 CSV file(s)")
            if toc_added:
                what_written.append("TOC with hyperlinks")

            QtWidgets.QMessageBox.information(
                self,
                "Final Output PDF Generated",
                f"Final PDF created successfully:\n{final_pdf_path}\n\nEmbedded/Generated: {', '.join(what_written)}"
            )
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to generate final PDF:\n{e}")
        finally:
            try:
                if temp_base_pdf and os.path.exists(temp_base_pdf):
                    os.remove(temp_base_pdf)
            except Exception:
                pass
            QtWidgets.QApplication.restoreOverrideCursor()


def main():
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")
    win = AnnotatorApp()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
