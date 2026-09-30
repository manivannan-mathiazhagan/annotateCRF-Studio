# ====================================================================================================
# Tool Name      : AnnotateCRF Studio
#
# Description    : GUI-based application for creating submission-ready annotated CRFs
#                  compliant with CDISC MSG 2.0 standards without requiring Adobe Acrobat,
#                  third-party PDF annotation software, or subscription-based tools.
#
# Purpose        : Enables Statistical Programmers and Clinical Data Standards teams to
#                  efficiently create, review, manage, and publish annotated CRFs with
#                  support for SDTM metadata integration, bookmarks, hyperlinks,
#                  annotation management, and final PDF generation.
#
# Key Features   :
#                  • CDISC Library integration for SDTM metadata retrieval
#                  • Interactive PDF annotation placement
#                  • Domain, Variable, Assigned Field, and NOTSUB annotations
#                  • Refer-page annotations with clickable hyperlinks
#                  • Annotation copy, edit, delete, and multi-select operations
#                  • Connector line creation and management
#                  • Bookmark creation, editing, copying, and hierarchy management
#                  • Automatic Table of Contents generation
#                  • Clickable PDF bookmarks and internal navigation
#                  • Single-workbook Excel export for annotations, bookmarks, connector lines, and variables
#                  • Review mode with annotation and bookmark management
#                  • Final publication-ready annotated PDF generation
#
# Input Files    :
#                  • Source CRF PDF
#                  • CDISC API Key JSON (optional for metadata integration)
#
# Output Files   :
#                  • Final Annotated PDF
#                  • Excel workbook containing Annotations, Bookmarks, Connector Lines, and Variables
#
# Standards      :
#                  • CDISC MSG 2.0
#                  • SDTMIG 3.2 / 3.3 / 3.4
#                  • CDISC Library API
#
# Developed By   : Manivannan Mathialagan
# Last Updated   : September 2026
#
# Change History  :
#                  • Sep 2026 - Added document-wide annotation font controls.
#                    Default mode uses 10 pt normal annotations and 12 pt domain annotations.
#                    Unchecking "Use default font sizes (10 / 12)" reveals independent
#                    Annotation Font and Domain Font dropdowns (8-14 pt).
#                    Custom selections are preserved when default mode is toggled back on/off.
#                    Preview, hit-testing, wrapping, box sizing, and final PDF generation all
#                    use the active font settings consistently.
#
# ====================================================================================================

import json
import os
import re
import sys
import tempfile
from collections import OrderedDict
from dataclasses import dataclass
from typing import List, Optional

# ================================
# Console-safe handling for .pyw / pythonw GUI mode
# ================================
# When this application is launched with pythonw.exe or as a .pyw file, Windows may not
# provide a valid console. Any package or print/log call that tries to write to the
# console can raise errors such as:
#   Cannot open console input buffer for writing
# These wrappers keep GUI actions such as Excel export independent of console availability.
class _NullWriter:
    def write(self, *_args, **_kwargs):
        return 0

    def flush(self):
        return None


def _ensure_console_safe_streams():
    for attr in ("stdout", "stderr"):
        stream = getattr(sys, attr, None)
        if stream is None:
            setattr(sys, attr, _NullWriter())
            continue
        try:
            stream.write("")
            stream.flush()
        except Exception:
            setattr(sys, attr, _NullWriter())


def safe_log(message: str):
    try:
        print(message)
    except Exception:
        pass


_ensure_console_safe_streams()

from PyQt5 import QtCore, QtGui, QtWidgets

QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)

import fitz
import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ================================
# Application paths and settings
# ================================
# Python source files are stored in <project>/scripts.
# Shared resources and local configuration are stored in the project root.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)

APP_ICON_CANDIDATES = [
    os.path.join(PROJECT_DIR, "annotator_icon.ico"),
    os.path.join(PROJECT_DIR, "annotator_icon.png"),
]

DEFAULT_ZOOM = 1.30
# CDISC MSG 2.0 recommended page-level colour sequence.
# RGB values:
#   Blue   = 191, 255, 255
#   Yellow = 255, 255, 150
#   Green  = 150, 255, 150
#   Orange = 255, 190, 155
#
# The sequence restarts on every CRF page and is assigned only to actual SDTM
# domains in the order in which they first occur on that page.
MSG_DOMAIN_COLORS = [
    (191 / 255.0, 255 / 255.0, 255 / 255.0),  # Blue
    (255 / 255.0, 255 / 255.0, 150 / 255.0),  # Yellow
    (150 / 255.0, 255 / 255.0, 150 / 255.0),  # Green
    (255 / 255.0, 190 / 255.0, 155 / 255.0),  # Orange
]

# Additional colours used only when more than four actual domains occur on one
# page. The first four always remain exactly in MSG 2.0 order.
MSG_EXTENDED_DOMAIN_COLORS = [
    (0.86, 0.82, 1.00),  # Lavender
    (0.80, 0.90, 1.00),  # Light blue
    (0.80, 1.00, 0.90),  # Mint
    (1.00, 0.90, 0.80),  # Light peach
]

DOMAIN_COLORS = MSG_DOMAIN_COLORS + MSG_EXTENDED_DOMAIN_COLORS
MSG_BLUE_FILL = MSG_DOMAIN_COLORS[0]
DEFAULT_OTHER_FILL = MSG_BLUE_FILL

# NOT SUBMITTED and Refer Page annotations are always Blue. They do not consume
# a position in the actual-domain colour sequence.
NOTSUB_FILL = MSG_BLUE_FILL
REFER_PAGE_FILL = MSG_BLUE_FILL
TEXT_COLOR = (0.0, 0.0, 0.0)
BORDER_COLOR = (0.0, 0.0, 0.0)

FONT_NORMAL = "helv"
FONT_BOLD = "hebo"

DEFAULT_BASE_FONT_SIZE = 10
DEFAULT_DOMAIN_FONT_SIZE = 12
FONT_SIZE_OPTIONS = list(range(8, 15))

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

ANNOTATION_EXCEL_COLUMNS = ["DOMAIN", "NAME", "PAGENO", "ANNOTATION", "ASSIGNEDFIELD", "X1", "Y1", "PAGEH", "BOX_W", "BOX_H"]
BOOKMARK_EXCEL_COLUMNS = ["TITLE", "LEVEL", "PAGENO"]
CONNECTOR_LINE_EXCEL_COLUMNS = ["PAGENO", "X1", "Y1", "X2", "Y2"]
VARIABLE_EXCEL_COLUMNS = ["DOMAIN", "VARIABLE", "PAGES"]

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
# CDISC Library SDTM metadata support
# ================================
CDISC_API_BASE = "https://library.cdisc.org/api"
CDISC_KEY_FILE_CANDIDATES = [
    os.path.join(PROJECT_DIR, "CDISC_API_KEY.json"),
    os.path.join(PROJECT_DIR, "cdisc_api_key.json"),
]
SDTMIG_STANDARD_OPTIONS = ["SDTMIG 3.4", "SDTMIG 3.3", "SDTMIG 3.2"]


def safe_text(value):
    if value is None:
        return ""
    return str(value).strip()


def load_cdisc_api_keys_for_sdtm():
    """Load primary/secondary CDISC Library keys from JSON in the project folder.

    Supported formats:
      {"cdisc_library": {"primary_key": "...", "secondary_key": "..."}}
      {"primary_key": "...", "secondary_key": "..."}
    """
    last_error = ""
    for key_file in CDISC_KEY_FILE_CANDIDATES:
        if not os.path.exists(key_file):
            continue
        try:
            with open(key_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            block = data.get("cdisc_library", data) if isinstance(data, dict) else {}
            primary = safe_text(block.get("primary_key") or block.get("primary") or block.get("api_key"))
            secondary = safe_text(block.get("secondary_key") or block.get("secondary"))
            if primary or secondary:
                return primary, secondary, key_file
            last_error = f"No primary_key/secondary_key found in {key_file}"
        except Exception as e:
            last_error = f"Unable to read {key_file}: {e}"
    raise FileNotFoundError(
        "CDISC API key file was not found or is invalid. Expected CDISC_API_KEY.json in the AnnotateCRF Studio project folder."
        + (f"\n{last_error}" if last_error else "")
    )


def cdisc_api_get(path_or_url, primary_key, secondary_key=""):
    url = safe_text(path_or_url)
    if not url.startswith("http"):
        if not url.startswith("/"):
            url = "/" + url
        url = CDISC_API_BASE + url
    errors = []
    for label, key in [("Primary", primary_key), ("Secondary", secondary_key)]:
        key = safe_text(key)
        if not key:
            continue
        try:
            r = requests.get(url, headers={"api-key": key, "Accept": "application/json"}, timeout=90)
            if r.status_code == 200:
                return r.json()
            errors.append(f"{label} {r.status_code}: {r.text[:250]}")
        except Exception as e:
            errors.append(f"{label} error: {e}")
    raise RuntimeError("Unable to access CDISC Library API. " + " | ".join(errors))


def iter_json_dicts(obj):
    if isinstance(obj, dict):
        yield obj
        for v in obj.values():
            yield from iter_json_dicts(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from iter_json_dicts(v)


def json_first_value(d, keys, default=""):
    if not isinstance(d, dict):
        return default
    lower = {str(k).lower(): k for k in d.keys()}
    for key in keys:
        lk = key.lower()
        if lk in lower:
            val = d.get(lower[lk])
            if val not in (None, "", [], {}):
                return val
    return default


def extract_api_links(data):
    links = []
    def add_link(obj, hint=""):
        if not isinstance(obj, dict):
            return
        href = obj.get("href") or obj.get("url") or obj.get("path")
        title = obj.get("title") or obj.get("name") or obj.get("label") or obj.get("submissionValue") or hint
        if href:
            links.append({"href": safe_text(href), "title": safe_text(title)})
    if isinstance(data, dict):
        for container_name in ("_links", "_embedded"):
            cont = data.get(container_name)
            if isinstance(cont, dict):
                for k, v in cont.items():
                    if isinstance(v, list):
                        for item in v:
                            add_link(item, k)
                    elif isinstance(v, dict):
                        add_link(v, k)
        # Some CDISC payloads expose plain arrays.
        for k, v in data.items():
            if isinstance(v, list) and k.lower() in {"datasets", "domains", "variables", "items"}:
                for item in v:
                    add_link(item, k)
    out, seen = [], set()
    for link in links:
        key = (link["href"], link["title"])
        if key not in seen:
            seen.add(key)
            out.append(link)
    return out


def _infer_domain_from_href(href):
    """Infer SDTM domain from CDISC Library href/source path."""
    txt = safe_text(href).upper()
    patterns = [
        r"/(?:DATASETS|DOMAINS)/([A-Z][A-Z0-9]{1,7})(?:/|$)",
        r"/(?:DATASET|DOMAIN)/([A-Z][A-Z0-9]{1,7})(?:/|$)",
        r"[?&](?:DATASET|DOMAIN)=([A-Z][A-Z0-9]{1,7})(?:&|$)",
    ]
    for pat in patterns:
        m = re.search(pat, txt)
        if m:
            return m.group(1)
    return ""


def _is_actual_sdtm_domain_code(value):
    """Keep real domain codes only; prevent variable names like AGE/ACTARM/BSTAT becoming domains."""
    dom = safe_text(value).upper()
    return bool(
        re.fullmatch(r"[A-Z][A-Z0-9]{1,3}", dom)
        or dom == "RELREC"
        or re.fullmatch(r"SUPP[A-Z0-9]{2,4}", dom)
        or re.fullmatch(r"SQ[A-Z0-9]{2,4}", dom)
    )


def _looks_like_dataset_record(d):
    if not isinstance(d, dict):
        return False
    keys = {str(k).lower() for k in d.keys()}
    dataset_keys = {
        "domain", "dataset", "datasetname", "shortname", "submissionvalue",
        "class", "structure", "purpose", "datasetlabel", "datasetclass"
    }
    variable_keys = {"variable", "variablename", "datatype", "type", "length", "core", "role", "origin"}
    return bool(keys & dataset_keys) and not bool(keys & variable_keys)


def parse_sdtm_domain_and_variable_metadata(payloads):
    """Return domain_labels and variables_by_domain from CDISC Library SDTMIG payloads.

    Important: CDISC Library variable payloads can contain NAME=AGE, NAME=ACTARM, etc.
    Those are variables, not domains. This parser uses dataset/domain payloads and href
    context first, then attaches variable records to the parent dataset/domain.
    """
    domain_labels = {}
    variables_by_domain = {}
    seen_vars = set()

    def add_domain(dom, label=""):
        dom = safe_text(dom).upper()
        label = safe_text(label)
        if not _is_actual_sdtm_domain_code(dom):
            return
        if dom not in domain_labels or (label and domain_labels.get(dom, "") in {"", dom}):
            domain_labels[dom] = label or domain_labels.get(dom, "") or dom

    def add_variable(dom, var, label="", role="", core=""):
        dom = safe_text(dom).upper()
        var = safe_text(var).upper()
        label = safe_text(label)
        if not _is_actual_sdtm_domain_code(dom):
            return
        if not var or not re.fullmatch(r"[A-Z][A-Z0-9_]{1,20}", var):
            return
        if var in {"TERM", "CODE", "CODELIST", "VALUE", "LABEL", "DOMAIN", "DATASET"}:
            return
        key = (dom, var)
        if key in seen_vars:
            return
        seen_vars.add(key)
        variables_by_domain.setdefault(dom, []).append({
            "domain": dom,
            "variable": var,
            "label": label,
            "role": safe_text(role),
            "core": safe_text(core),
        })

    # Pass 1: collect dataset/domain labels only from likely dataset records or source href.
    for payload in payloads:
        source_href = safe_text(payload.get("__source_href", "")) if isinstance(payload, dict) else ""
        source_dom = _infer_domain_from_href(source_href)
        for d in iter_json_dicts(payload):
            if not isinstance(d, dict):
                continue
            d_source = safe_text(d.get("__source_href", "")) or source_href
            href_dom = _infer_domain_from_href(d_source)
            raw_dom = json_first_value(d, ["domain", "dataset", "datasetName", "shortName", "submissionValue"], "")
            if isinstance(raw_dom, dict):
                raw_dom = json_first_value(raw_dom, ["name", "submissionValue", "label"], "")
            dom = safe_text(raw_dom).upper() or href_dom or source_dom
            label = json_first_value(d, ["label", "description", "title", "datasetLabel", "preferredTerm", "definition"], "")
            if _looks_like_dataset_record(d):
                add_domain(dom, label)
            elif href_dom and any(str(k).lower() in {"label", "description", "title", "datasetlabel"} for k in d.keys()):
                # Dataset-specific payload root may be sparse but the href tells us the domain.
                add_domain(href_dom, label)

    # Pass 2: collect variable records and attach to parent domain from explicit dataset/domain or href.
    for payload in payloads:
        source_href = safe_text(payload.get("__source_href", "")) if isinstance(payload, dict) else ""
        source_dom = _infer_domain_from_href(source_href)
        for d in iter_json_dicts(payload):
            if not isinstance(d, dict):
                continue
            keys = {str(k).lower() for k in d.keys()}
            if not (keys & {"variable", "variablename", "core", "role", "datatype", "type", "length", "origin"}):
                continue

            var = json_first_value(d, ["variable", "variableName", "name", "submissionValue"], "")
            if isinstance(var, dict):
                var = json_first_value(var, ["name", "submissionValue", "label"], "")
            var = safe_text(var).upper()

            explicit_dom = json_first_value(d, ["domain", "dataset", "datasetName", "parentDomain"], "")
            if isinstance(explicit_dom, dict):
                explicit_dom = json_first_value(explicit_dom, ["name", "submissionValue", "label"], "")
            explicit_dom = safe_text(explicit_dom).upper()
            d_href = safe_text(d.get("__source_href", "")) or source_href
            href_dom = _infer_domain_from_href(d_href)
            vdom = explicit_dom if _is_actual_sdtm_domain_code(explicit_dom) else (href_dom or source_dom)

            # Do not guess AGE -> AGE or ACTARM -> ACTARM. If no parent domain is known, skip.
            if not _is_actual_sdtm_domain_code(vdom):
                continue

            vlabel = json_first_value(d, ["label", "variableLabel", "description", "title", "definition"], "")
            role = json_first_value(d, ["role", "varRole"], "")
            core = json_first_value(d, ["core", "requirement", "mandatory"], "")
            add_domain(vdom, domain_labels.get(vdom, vdom))
            add_variable(vdom, var, vlabel, role, core)

    # Preserve the domain and variable order supplied by the CDISC Library metadata.
    # Do not alphabetically sort variables: their source order is meaningful in the UI.
    ordered_domain_labels = {}
    for dom in domain_labels:
        if dom in variables_by_domain:
            ordered_domain_labels[dom] = domain_labels.get(dom, dom)
    for dom in variables_by_domain:
        if dom not in ordered_domain_labels:
            ordered_domain_labels[dom] = domain_labels.get(dom, dom)

    return ordered_domain_labels, variables_by_domain


def _tag_payload_source(payload, href):
    """Attach source href to a payload recursively enough for parser context."""
    if isinstance(payload, dict):
        payload.setdefault("__source_href", href)
        for v in payload.values():
            if isinstance(v, dict):
                v.setdefault("__source_href", href)
            elif isinstance(v, list):
                for item in v:
                    if isinstance(item, dict):
                        item.setdefault("__source_href", href)
    return payload


def load_sdtmig_metadata_from_cdisc(standard_text):
    primary, secondary, key_file = load_cdisc_api_keys_for_sdtm()
    version_match = re.search(r"(\d+)\.(\d+)", safe_text(standard_text))
    if version_match:
        version = f"{version_match.group(1)}-{version_match.group(2)}"
    else:
        version = "3-4"

    endpoints = [
        f"/mdr/sdtmig/{version}",
        f"/mdr/sdtmig/{version}/datasets",
        f"/mdr/sdtmig/{version}/domains",
    ]
    payloads = []
    errors = []
    for endpoint in endpoints:
        try:
            payloads.append(_tag_payload_source(cdisc_api_get(endpoint, primary, secondary), endpoint))
        except Exception as e:
            errors.append(f"{endpoint}: {e}")

    # Follow dataset/domain links and then variable links.
    followed = set()
    for payload in list(payloads):
        for link in extract_api_links(payload):
            href = link.get("href", "")
            hlow = href.lower()
            if not href or href in followed:
                continue
            if any(token in hlow for token in ["/datasets/", "/domains/", "/variables", "variables"]):
                followed.add(href)
                try:
                    sub = _tag_payload_source(cdisc_api_get(href, primary, secondary), href)
                    payloads.append(sub)
                    for sub_link in extract_api_links(sub):
                        shref = sub_link.get("href", "")
                        if shref and shref not in followed and "variable" in shref.lower():
                            followed.add(shref)
                            try:
                                payloads.append(_tag_payload_source(cdisc_api_get(shref, primary, secondary), shref))
                            except Exception:
                                pass
                except Exception:
                    pass

    domain_labels, variables_by_domain = parse_sdtm_domain_and_variable_metadata(payloads)

    if not variables_by_domain:
        raise RuntimeError(
            "Unable to parse SDTM domain/variable metadata from CDISC Library.\n\n"
            "Tried:\n" + "\n".join(endpoints) +
            ("\n\nEndpoint errors:\n" + "\n".join(errors) if errors else "")
        )

    return {
        "key_file": key_file,
        "standard": standard_text,
        "domain_labels": domain_labels,
        "variables_by_domain": variables_by_domain,
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


def is_refer_page_entry(entry) -> bool:
    """Return True when an annotation is a Refer Page annotation."""
    domain = (getattr(entry, "domain", "") or "").strip().upper()
    name = (getattr(entry, "name", "") or "").strip().upper()
    annotation = (getattr(entry, "annotation", "") or "").strip()

    if domain == "REF" and name == "REF":
        return True

    return extract_page_reference(annotation) is not None and domain == "REF"


def get_page_domain_color_map(entries, pageno: int) -> dict:
    """Return the MSG 2.0 actual-domain colour map for one CRF page.

    Rules:
      • The first actual SDTM domain on the page is Blue.
      • The second is Yellow.
      • The third is Green.
      • The fourth is Orange.
      • The sequence restarts on every page.
      • Repeated annotations for the same domain retain the same page colour.
      • NOT SUBMITTED and Refer Page are always Blue and do not consume a
        position in the actual-domain sequence.
    """
    ordered_domains = []

    for entry in entries:
        if entry.pageno != pageno:
            continue

        if bool(getattr(entry, "is_not_submitted", False)):
            continue

        if is_refer_page_entry(entry):
            continue

        domain = (getattr(entry, "domain", "") or "").strip().upper()
        if not domain or domain in {"NOTSUB", "REF"}:
            continue

        if domain not in ordered_domains:
            ordered_domains.append(domain)

    color_map = {}
    for index, domain in enumerate(ordered_domains):
        color_map[domain] = DOMAIN_COLORS[index % len(DOMAIN_COLORS)]

    return color_map


def compute_entry_layout(entry, color_map, page_width=None, base_font_size=DEFAULT_BASE_FONT_SIZE, domain_font_size=DEFAULT_DOMAIN_FONT_SIZE):
    domain_key = (entry.domain or "").strip().upper()

    if entry.is_not_submitted:
        fill = NOTSUB_FILL
    elif is_refer_page_entry(entry):
        fill = REFER_PAGE_FILL
    else:
        # Actual domains follow the MSG 2.0 page-level order. A safe Blue
        # fallback is used if an imported or unusual entry has no map value.
        fill = color_map.get(domain_key, MSG_BLUE_FILL)
    dashed = bool(entry.is_assigned_field)
    bold = bool(entry.is_domain_annotation)
    font_size = domain_font_size if bold else base_font_size
    scale = domain_font_size / base_font_size if bold else 1.0
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

def build_visual_pdf_copy(input_pdf: str) -> fitz.Document:
    """Create a working PDF with every page normalized to the visible orientation.

    Some CRFs contain portrait pages with /Rotate=90 or /Rotate=270. PyMuPDF renders
    those pages visually as landscape in the GUI, but direct draw_rect/insert_textbox
    operations are applied in the underlying page coordinate system. That can make
    annotations rotate or shift in the final PDF.

    This function copies each source page into a new page using the page's visible
    rectangle. The resulting working document has rotation=0 pages whose width/height
    match what the user saw in the GUI. Final annotation coordinates then match the
    preview for both portrait and landscape pages.
    """
    src = fitz.open(input_pdf)
    out = fitz.open()

    try:
        for page_index in range(len(src)):
            src_page = src[page_index]
            visible_rect = src_page.rect  # includes page rotation as displayed
            new_page = out.new_page(
                width=float(visible_rect.width),
                height=float(visible_rect.height)
            )
            new_page.show_pdf_page(new_page.rect, src, page_index)
        return out
    except Exception:
        out.close()
        raise
    finally:
        src.close()



def place_visual_overlay_on_original_page(page, overlay_doc):
    """Overlay annotations drawn in GUI/visible coordinates onto the original page.

    The source page is not normalized, rotated, cleaned, or rebuilt, so existing PDF
    content and existing PDF annotation/comment objects are preserved.

    For pages with PDF /Rotate, PyMuPDF's page.rect is the visible GUI rectangle,
    while the underlying drawing target is the native cropbox/mediabox. The overlay
    must be placed back using the page's own rotation, not the inverse rotation.
    """
    rot = int(page.rotation or 0) % 360

    if rot == 0:
        page.show_pdf_page(page.rect, overlay_doc, 0, overlay=True)
        return

    try:
        target_rect = page.cropbox
    except Exception:
        target_rect = page.mediabox

    if target_rect.is_empty or target_rect.width <= 0 or target_rect.height <= 0:
        target_rect = page.mediabox

    # IMPORTANT:
    # Use the SAME rotation as the source page. Using inverse rotation makes the
    # overlay appear upside down/mirrored on pages like the page 18 medication log.
    page.show_pdf_page(target_rect, overlay_doc, 0, rotate=rot, overlay=True)


def visual_rect_to_native_rect(page, rect):
    """Convert a GUI/visible rectangle to the page's native rectangle.

    Used for clickable refer-page links only. Annotation drawing itself is handled
    through a visual overlay.
    """
    rot = int(page.rotation or 0) % 360
    if rot == 0:
        return rect

    pts = [
        fitz.Point(rect.x0, rect.y0),
        fitz.Point(rect.x1, rect.y0),
        fitz.Point(rect.x1, rect.y1),
        fitz.Point(rect.x0, rect.y1),
    ]
    pts = [pt * page.derotation_matrix for pt in pts]
    xs = [pt.x for pt in pts]
    ys = [pt.y for pt in pts]
    return fitz.Rect(min(xs), min(ys), max(xs), max(ys))


def clamp_rect_to_page(rect, page_width, page_height):
    """Shift annotation rect inside the visible page without resizing it."""
    if rect.x1 > page_width:
        shift = rect.x1 - page_width
        rect = fitz.Rect(rect.x0 - shift, rect.y0, rect.x1 - shift, rect.y1)
    if rect.y1 > page_height:
        shift = rect.y1 - page_height
        rect = fitz.Rect(rect.x0, rect.y0 - shift, rect.x1, rect.y1 - shift)
    if rect.x0 < 0:
        shift = -rect.x0
        rect = fitz.Rect(rect.x0 + shift, rect.y0, rect.x1 + shift, rect.y1)
    if rect.y0 < 0:
        shift = -rect.y0
        rect = fitz.Rect(rect.x0, rect.y0 + shift, rect.x1, rect.y1 + shift)
    return rect

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


def get_reference_page_size(source_doc=None, source_pdf_path: str = ""):
    """Return the first page size from the actual source PDF.

    Do not force A4. CRFs can contain mixed page orientations/sizes, and the
    generated TOC should use the same base size as the source document.
    """
    try:
        if source_doc is not None and len(source_doc) > 0:
            rect = source_doc[0].rect
            return float(rect.width), float(rect.height)
    except Exception:
        pass

    if source_pdf_path:
        try:
            tmp_doc = fitz.open(source_pdf_path)
            try:
                if len(tmp_doc) > 0:
                    rect = tmp_doc[0].rect
                    return float(rect.width), float(rect.height)
            finally:
                tmp_doc.close()
        except Exception:
            pass

    # Fallback only when no PDF page is available.
    return fitz.paper_size("a4")


def compute_toc_page_count_from_bookmarks(bookmarks, source_doc=None, source_pdf_path: str = "", font_size: int = 12, lines_per_page: int = 38):
    toc_entries = build_toc_entries_from_bookmarks(bookmarks)
    if not toc_entries:
        return 0
    width, _ = get_reference_page_size(source_doc=source_doc, source_pdf_path=source_pdf_path)
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

        width, height = get_reference_page_size(source_doc=original)
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


WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
    "CONIN$", "CONOUT$",
}


def ensure_xlsx_extension(path: str) -> str:
    path = str(path or "").strip()
    if path and not path.lower().endswith(".xlsx"):
        path += ".xlsx"
    return path


def validate_output_path(path: str):
    if not path:
        raise ValueError("Output path is blank.")

    base = os.path.basename(path).strip().rstrip(". ")
    if os.name == "nt":
        stem = os.path.splitext(base)[0].upper()
        if stem in WINDOWS_RESERVED_NAMES:
            raise ValueError(
                f"'{base}' is a reserved Windows device name. Please use a normal file name, "
                f"for example annotatecrf_workbook.xlsx."
            )

    out_dir = os.path.dirname(os.path.abspath(path)) or os.getcwd()
    if not os.path.isdir(out_dir):
        raise FileNotFoundError(f"Output folder does not exist:\n{out_dir}")

    if not os.access(out_dir, os.W_OK):
        raise PermissionError(f"No write access to output folder:\n{out_dir}")


def atomic_excel_write(out_path: str, workbook: Workbook):
    """Write the workbook through a temporary file, then replace the final output."""
    out_path = ensure_xlsx_extension(out_path)
    validate_output_path(out_path)

    out_dir = os.path.dirname(os.path.abspath(out_path)) or os.getcwd()
    fd, tmp_path = tempfile.mkstemp(prefix="._acrf_export_", suffix=".xlsx", dir=out_dir)
    os.close(fd)

    try:
        workbook.save(tmp_path)
        os.replace(tmp_path, out_path)
    except Exception:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
        raise

    return out_path


# ================================
# Dialog for annotation details
# ================================
class AnnotationDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, initial_data=None, dialog_title="Annotation Details"):
        super().__init__(parent)
        self.parent_app = parent
        self.setWindowTitle(dialog_title)
        self.resize(760, 520)
        self.setMinimumSize(760, 520)
        self.setModal(True)

        self._updating_refpage_ui = False
        self._auto_annotation = ""

        self.domain_labels = getattr(parent, "sdtm_domain_labels", {}) if parent is not None else {}
        self.variables_by_domain = getattr(parent, "sdtm_variables_by_domain", {}) if parent is not None else {}

        self.setStyleSheet("""
            QDialog { background-color: #f4f8ff; border-radius: 14px; }
            QLabel { font-family: 'Times New Roman'; font-size: 12pt; color: #1c2e4a; }
            QLineEdit, QTextEdit, QSpinBox, QComboBox {
                background: #ffffff; border: 1px solid #a8bfdc; border-radius: 8px;
                padding: 6px 8px; font-family: 'Times New Roman'; font-size: 12pt; color: #1a1a1a;
            }
            QLineEdit:focus, QTextEdit:focus, QSpinBox:focus, QComboBox:focus { border: 2px solid #4d8ef7; background: #fdfefe; }
            QCheckBox { font-family: 'Times New Roman'; font-size: 12pt; color: #143b66; spacing: 8px; }
            QCheckBox::indicator { width: 18px; height: 18px; }
            QDialogButtonBox QPushButton {
                font-family: 'Times New Roman'; font-size: 12pt; font-weight: bold;
                border-radius: 10px; padding: 8px 18px; min-width: 100px;
            }
        """)

        self.domain_combo = QtWidgets.QComboBox()
        self.domain_combo.setEditable(True)
        self.domain_combo.setInsertPolicy(QtWidgets.QComboBox.NoInsert)
        self.domain_combo.setMinimumWidth(420)

        self.name_combo = QtWidgets.QComboBox()
        self.name_combo.setEditable(True)
        self.name_combo.setInsertPolicy(QtWidgets.QComboBox.NoInsert)
        self.name_combo.setMinimumWidth(420)

        self.annotation_edit = QtWidgets.QTextEdit()
        self.annotation_edit.setMinimumHeight(140)

        self.chk_domain = QtWidgets.QCheckBox("Domain annotation")
        self.chk_assigned = QtWidgets.QCheckBox("Assigned field")
        self.chk_notsub = QtWidgets.QCheckBox("Not Submitted")
        self.chk_refpage = QtWidgets.QCheckBox("Refer Page")
        self.chk_supp = QtWidgets.QCheckBox("SUPP annotation")

        self.ref_page_label = QtWidgets.QLabel("Reference Page")
        self.ref_page_spin = QtWidgets.QSpinBox()
        self.ref_page_spin.setRange(1, 999999)
        self.ref_page_spin.setValue(1)
        self.ref_page_spin.setFixedWidth(110)

        self.populate_domain_combo()

        self.domain_combo.currentIndexChanged.connect(self.on_domain_changed)
        self.domain_combo.lineEdit().editingFinished.connect(self.on_domain_changed)
        self.name_combo.currentIndexChanged.connect(self.on_variable_changed)
        self.name_combo.lineEdit().editingFinished.connect(self.on_variable_changed)
        self.chk_domain.stateChanged.connect(self.toggle_dialog_state)
        self.chk_assigned.stateChanged.connect(self.toggle_dialog_state)
        self.chk_notsub.stateChanged.connect(self.toggle_dialog_state)
        self.chk_refpage.stateChanged.connect(self.toggle_dialog_state)
        self.chk_supp.stateChanged.connect(self.toggle_dialog_state)
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
        form.addRow(self.lbl_domain, self.domain_combo)
        form.addRow(self.lbl_name, self.name_combo)
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
        chk_row.addWidget(self.chk_supp)
        chk_row.addStretch()

        ref_row = QtWidgets.QHBoxLayout()
        ref_row.setSpacing(12)
        ref_row.addSpacing(8)
        ref_row.addWidget(self.ref_page_label)
        ref_row.addWidget(self.ref_page_spin)
        ref_row.addStretch()

        std_text = getattr(parent, "sdtm_standard_text", "SDTMIG metadata not loaded") if parent is not None else "SDTMIG metadata not loaded"
        note = QtWidgets.QLabel(
            f"SDTM metadata source: {std_text}. Select Domain and Variable from CDISC Library metadata. "
            "For domain annotation, annotation text is populated automatically as DOMAIN (Dataset Label). SUPP annotation uses text like VARIABLE in SUPPDOMAIN."
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
        else:
            self.on_domain_changed()

    def selected_domain(self):
        txt = safe_text(self.domain_combo.currentText())
        # Always trust the editable text first. If a user types a domain that is not
        # in the CDISC list, currentData() can still point to the previous selected
        # item; using currentText() prevents stale domain values.
        if txt:
            return txt.strip().upper()
        data = self.domain_combo.currentData()
        return safe_text(data).upper() if data else ""

    def selected_variable(self):
        txt = safe_text(self.name_combo.currentText())
        # Always trust the editable text first. If a user types a variable manually,
        # currentData() can still point to the previous selected item.
        if txt:
            return txt.strip().upper()
        data = self.name_combo.currentData()
        return safe_text(data).upper() if data else ""

    def domain_label(self, domain=None):
        dom = domain or self.selected_domain()
        return safe_text(self.domain_labels.get(dom, ""))

    def variable_label(self, domain=None, variable=None):
        dom = domain or self.selected_domain()
        var = variable or self.selected_variable()
        for item in self.variables_by_domain.get(dom, []):
            if safe_text(item.get("variable")).upper() == var:
                return safe_text(item.get("label"))
        return ""

    def populate_domain_combo(self):
        self.domain_combo.blockSignals(True)
        self.domain_combo.clear()
        # Preserve the domain order supplied by the loaded CDISC metadata.
        for dom in self.domain_labels.keys():
            self.domain_combo.addItem(dom, dom)
        self.domain_combo.blockSignals(False)
        self.populate_variable_combo()

    def populate_variable_combo(self):
        dom = self.selected_domain()
        self.name_combo.blockSignals(True)
        self.name_combo.clear()

        # Refresh immediately when Domain changes. Do not carry a variable from the
        # previously selected domain into the new domain. Preserve CDISC source order.
        for item in self.variables_by_domain.get(dom, []):
            var = safe_text(item.get("variable")).upper()
            if var:
                self.name_combo.addItem(var, var)

        if self.name_combo.count() > 0:
            self.name_combo.setCurrentIndex(0)
        else:
            self.name_combo.setEditText("")

        self.name_combo.blockSignals(False)

    def set_domain_value(self, domain):
        domain = safe_text(domain).upper()
        self.domain_combo.blockSignals(True)
        idx = self.domain_combo.findData(domain)
        if idx >= 0:
            self.domain_combo.setCurrentIndex(idx)
        elif domain:
            self.domain_combo.setEditText(domain)
        else:
            self.domain_combo.setEditText("")
        self.domain_combo.blockSignals(False)
        self.populate_variable_combo()

    def set_variable_value(self, variable):
        variable = safe_text(variable).upper()
        idx = self.name_combo.findData(variable)
        if idx >= 0:
            self.name_combo.setCurrentIndex(idx)
        elif variable:
            self.name_combo.setEditText(variable)

    def set_auto_annotation(self, text):
        text = safe_text(text)
        current = self.annotation_edit.toPlainText().replace("\t", "    ").rstrip()
        if not current or current == self._auto_annotation:
            self.annotation_edit.setPlainText(text)
        self._auto_annotation = text

    def force_auto_annotation(self, text):
        # Used for system-generated annotations such as SUPP and Domain annotation.
        # These should refresh immediately when Domain/Variable changes.
        text = safe_text(text)
        self.annotation_edit.setPlainText(text)
        self._auto_annotation = text

    def build_domain_annotation_text(self):
        dom = self.selected_domain()
        label = self.domain_label(dom)
        return f"{dom} ({label})" if dom and label and label != dom else dom

    def build_variable_annotation_text(self):
        # Variable annotation text should show only the SDTM variable name.
        # Domain is stored separately in the entry metadata and is used for colouring.
        var = self.selected_variable()
        dom = self.selected_domain()
        return var or dom

    def build_supp_annotation_text(self):
        # SUPP annotation display should be like: CMAESPID in SUPPCM
        # Metadata remains parent domain=CM and variable=CMAESPID.
        dom = self.selected_domain()
        var = self.selected_variable()
        if not dom and not var:
            return ""
        supp_domain = dom if dom.startswith("SUPP") else f"SUPP{dom}"
        return f"{var} in {supp_domain}" if var else f"in {supp_domain}"

    def on_domain_changed(self):
        self.populate_variable_combo()
        if self.chk_domain.isChecked():
            # Domain annotation should suggest DOMAIN (Label) only when the annotation
            # box is blank or still contains the previous auto text. Do not overwrite
            # a manually edited label such as DC (Demographics as Collected).
            self.set_auto_annotation(self.build_domain_annotation_text())
        elif not self.chk_notsub.isChecked() and not self.chk_refpage.isChecked():
            if self.chk_supp.isChecked():
                self.force_auto_annotation(self.build_supp_annotation_text())
            else:
                self.set_auto_annotation(self.build_variable_annotation_text())

    def on_variable_changed(self):
        if not self.chk_domain.isChecked() and not self.chk_notsub.isChecked() and not self.chk_refpage.isChecked():
            if self.chk_supp.isChecked():
                self.force_auto_annotation(self.build_supp_annotation_text())
            else:
                self.set_auto_annotation(self.build_variable_annotation_text())

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

        self.set_domain_value(domain)
        self.set_variable_value(name)
        self.chk_domain.setChecked(is_domain_annotation)
        self.chk_assigned.setChecked(is_assigned_field)
        self.chk_notsub.setChecked(is_not_submitted)
        self.annotation_edit.setPlainText(annotation)
        self._auto_annotation = annotation
        self.toggle_dialog_state()

    def _set_ref_annotation_text(self):
        ref_page = int(self.ref_page_spin.value())
        self.domain_combo.setEditText("REF")
        self.name_combo.setEditText("REF")
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
        is_supp = self.chk_supp.isChecked()

        self.ref_page_label.setVisible(is_refpage)
        self.ref_page_spin.setVisible(is_refpage)

        if is_refpage:
            self._updating_refpage_ui = True
            try:
                self.chk_domain.setChecked(False)
                self.chk_assigned.setChecked(False)
                self.chk_notsub.setChecked(False)
                self.chk_supp.setChecked(False)
                self.chk_domain.setDisabled(True)
                self.chk_assigned.setDisabled(True)
                self.chk_notsub.setDisabled(True)
                self.chk_supp.setDisabled(True)
                self.domain_combo.setDisabled(True)
                self.name_combo.setDisabled(True)
                self.annotation_edit.setDisabled(True)
                self.lbl_name.setVisible(True)
                self.name_combo.setVisible(True)
                self._set_ref_annotation_text()
            finally:
                self._updating_refpage_ui = False
            return

        self.chk_domain.setDisabled(False)
        self.chk_assigned.setDisabled(False)
        self.chk_notsub.setDisabled(False)
        self.chk_supp.setDisabled(False)
        self.domain_combo.setDisabled(False)
        self.name_combo.setDisabled(False)
        self.annotation_edit.setDisabled(False)

        if is_notsub:
            self.domain_combo.setEditText("NOTSUB")
            self.name_combo.setEditText("NOTSUB")
            self.annotation_edit.setPlainText("[NOT SUBMITTED]")
            self.domain_combo.setDisabled(True)
            self.name_combo.setDisabled(True)
            self.annotation_edit.setDisabled(True)
            self.chk_domain.setChecked(False)
            self.chk_domain.setDisabled(True)
            self.chk_assigned.setChecked(False)
            self.chk_assigned.setDisabled(True)
            self.chk_supp.setChecked(False)
            self.chk_supp.setDisabled(True)
            self.lbl_name.setVisible(True)
            self.name_combo.setVisible(True)
            return

        if is_domain:
            self.chk_supp.setChecked(False)
            self.chk_supp.setDisabled(True)

        self.lbl_name.setVisible(not is_domain)
        self.name_combo.setVisible(not is_domain)
        if is_domain:
            self.name_combo.setEditText("")
            # Keep domain annotation text editable so users can manually enter labels
            # for custom domains not available in CDISC metadata, e.g.
            # DC (Demographics as Collected).
            self.annotation_edit.setDisabled(False)
            self.set_auto_annotation(self.build_domain_annotation_text())
        else:
            # Standard variable/VLM annotation remains editable.
            # Only SUPP annotation changes and locks the display text to: VARIABLE in SUPPDOMAIN.
            if is_supp:
                self.annotation_edit.setDisabled(True)
                self.force_auto_annotation(self.build_supp_annotation_text())
            else:
                self.annotation_edit.setDisabled(False)
                self.set_auto_annotation(self.build_variable_annotation_text())

    def validate_and_accept(self):
        is_refpage = self.chk_refpage.isChecked()
        if is_refpage:
            self.accept()
            return

        if not self.selected_domain():
            QtWidgets.QMessageBox.warning(self, "Validation", "DOMAIN is required.")
            return
        if not self.annotation_edit.toPlainText().replace("\t", "    ").rstrip():
            QtWidgets.QMessageBox.warning(self, "Validation", "ANNOTATION is required.")
            return
        if not self.chk_domain.isChecked() and not self.chk_notsub.isChecked():
            if not self.selected_variable():
                QtWidgets.QMessageBox.warning(self, "Validation", "VARIABLE is required.")
                return
        self.accept()

    def get_values(self):
        is_refpage = self.chk_refpage.isChecked()
        is_domain_annotation = self.chk_domain.isChecked()
        is_assigned_field = self.chk_assigned.isChecked()
        is_not_submitted = self.chk_notsub.isChecked()
        is_supp = self.chk_supp.isChecked()

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

        domain = self.selected_domain()
        variable = "" if is_domain_annotation else self.selected_variable()
        if is_domain_annotation:
            # Preserve manually edited domain annotation text. Do not rebuild from
            # metadata on OK, otherwise custom labels are lost.
            annotation = self.annotation_edit.toPlainText().replace("\t", "    ").rstrip()
        elif is_supp:
            annotation = self.build_supp_annotation_text()
        else:
            annotation = self.annotation_edit.toPlainText().replace("\t", "    ").rstrip()
        return {
            "domain": domain,
            "name": variable,
            "annotation": annotation,
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
        layout = compute_entry_layout(entry, color_map, page_width=page_width, base_font_size=self.main_window.base_font_size, domain_font_size=self.main_window.domain_font_size)
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

            layout = compute_entry_layout(entry, color_map, page_width=page_width, base_font_size=self.main_window.base_font_size, domain_font_size=self.main_window.domain_font_size)
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

            # GUI preview must use exactly the same page-level MSG colour as
            # the final PDF output.
            preview_fill = layout["fill"]
            fill_q = qcolor_from_rgb01(preview_fill)
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

        self.sdtm_standard_text = "SDTMIG 3.4"
        self.sdtm_domain_labels = {}
        self.sdtm_variables_by_domain = {}
        self.sdtm_metadata_loaded = False

        # Document-wide annotation font settings. Default mode preserves the
        # original 10 pt normal / 12 pt domain annotation appearance.
        self.base_font_size = DEFAULT_BASE_FONT_SIZE
        self.domain_font_size = DEFAULT_DOMAIN_FONT_SIZE

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

        header = QtWidgets.QLabel(
            "<span style=\"font-size:17pt; font-weight:bold;\">AnnotateCRF Studio</span> "
            "<span style=\"font-size:12pt; font-weight:normal;\">"
            "(One-Click Annotation &amp; Bookmarking for MSG 2.0–Compliant aCRFs)</span>"
        )
        header.setTextFormat(QtCore.Qt.RichText)
        header.setAlignment(QtCore.Qt.AlignCenter)
        header.setWordWrap(False)
        header.setStyleSheet("""
            QLabel {
                color: black; background: #bfe9f7; padding: 10px 8px; border-radius: 16px;
                font-family: 'Times New Roman';
            }
        """)
        layout.addWidget(header)

        contact_note = QtWidgets.QLabel(
            "For queries / suggestions / issues: Manivannan.Mathi@outlook.com"
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
            b.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
            b.setMinimumHeight(46)
            b.setMinimumWidth(max(105, b.fontMetrics().horizontalAdvance(text) + 34))
            return b

        # Clean two-row toolbar: project actions on row 1, PDF navigation on row 2
        toolbar_frame = QtWidgets.QFrame()
        toolbar_frame.setObjectName("toolbarFrame")
        toolbar_frame.setStyleSheet("""
            QFrame#toolbarFrame {
                background: #f7fbff;
                border: 1px solid #d8e6f3;
                border-radius: 12px;
            }
        """)
        toolbar_layout = QtWidgets.QVBoxLayout(toolbar_frame)
        toolbar_layout.setContentsMargins(10, 8, 10, 8)
        toolbar_layout.setSpacing(8)

        top_row = QtWidgets.QHBoxLayout()
        top_row.setSpacing(8)
        nav_row = QtWidgets.QHBoxLayout()
        nav_row.setSpacing(8)

        cdisc_label = QtWidgets.QLabel("SDTM")
        cdisc_label.setStyleSheet("""
            QLabel {
                font-family: 'Times New Roman';
                font-size: 11pt;
                font-weight: bold;
                color: #1e3a8a;
                padding: 0 4px;
            }
        """)
        cdisc_label.setMinimumWidth(70)
        cdisc_label.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        cdisc_label.setAlignment(QtCore.Qt.AlignCenter)

        self.sdtmig_combo = QtWidgets.QComboBox()
        self.sdtmig_combo.addItems(SDTMIG_STANDARD_OPTIONS)
        self.sdtmig_combo.setCurrentText(self.sdtm_standard_text)
        self.sdtmig_combo.setMinimumWidth(170)
        self.sdtmig_combo.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        self.sdtmig_combo.setMinimumHeight(46)
        self.sdtmig_combo.setStyleSheet("""
            QComboBox {
                background-color: #ffffff;
                color: #12395f;
                border: 1px solid #b8cfe4;
                border-radius: 10px;
                padding: 6px 34px 6px 12px;
                font-family: 'Times New Roman';
                font-size: 11pt;
                font-weight: bold;
            }
            QComboBox:hover {
                border: 1px solid #5f9ed1;
                background-color: #f8fbff;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 28px;
                border-left: 1px solid #d4e2ef;
                border-top-right-radius: 10px;
                border-bottom-right-radius: 10px;
                background-color: #eef6ff;
            }
            QComboBox::down-arrow {
                image: none;
                width: 0px;
                height: 0px;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 7px solid #12395f;
                margin-right: 8px;
            }
            QAbstractItemView {
                background-color: #ffffff;
                color: #12395f;
                selection-background-color: #d9ecff;
                selection-color: #12395f;
                border: 1px solid #b8cfe4;
                font-family: 'Times New Roman';
                font-size: 11pt;
                padding: 4px;
            }
        """)

        self.btn_load_sdtm = mkbtn("Load Metadata", "#0f766e", "#159287")
        self.btn_load_sdtm.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)

        self.btn_open = mkbtn("Open PDF", "#3b82f6", "#5c9cff")
        self.btn_open.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)

        self.btn_import_excel = mkbtn("Import Excel", "#0f766e", "#159287")
        self.btn_import_excel.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        self.btn_import_excel.setEnabled(False)

        self.btn_export_excel = mkbtn("Export Excel", "#2563eb", "#4a7df2")
        self.btn_export_excel.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        self.btn_export_excel.setEnabled(False)

        self.btn_terminate = mkbtn("Terminate", "#991b1b", "#b91c1c")
        self.btn_terminate.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)

        # Fill the complete first toolbar row with evenly distributed controls.
        top_row.addWidget(cdisc_label, 0)
        top_row.addWidget(self.sdtmig_combo, 2)
        top_row.addWidget(self.btn_load_sdtm, 2)
        top_row.addWidget(self.btn_open, 2)
        top_row.addWidget(self.btn_import_excel, 2)
        top_row.addWidget(self.btn_export_excel, 2)

        self.btn_prev = mkbtn("Previous Page", "#9ca3af", "#b6bcc7", "#1f2937")
        self.btn_prev.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)

        self.btn_next = mkbtn("Next Page", "#6b7280", "#7b8495")
        self.btn_next.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)

        self.btn_prev.setEnabled(False)
        self.btn_next.setEnabled(False)

        self.page_jump_spin = QtWidgets.QSpinBox()
        self.page_jump_spin.setRange(1, 1)
        self.page_jump_spin.setValue(1)
        self.page_jump_spin.setEnabled(False)
        self.page_jump_spin.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
        self.page_jump_spin.setAlignment(QtCore.Qt.AlignCenter)
        self.page_jump_spin.setMinimumWidth(95)
        self.page_jump_spin.setMinimumHeight(46)
        self.page_jump_spin.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        self.page_jump_spin.setStyleSheet("""
            QSpinBox {
                background: #ffffff; color: #12608d; border: 1px solid #b8cfe4; border-radius: 8px;
                padding: 5px 8px; font-family: 'Times New Roman'; font-size: 11pt; font-weight: bold;
            }
        """)

        self.btn_go_page = mkbtn("Go", "#0284c7", "#0ea5e9")
        self.btn_go_page.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        self.btn_go_page.setEnabled(False)

        self.page_info = QtWidgets.QLabel("Page: -")
        self.page_info.setMinimumWidth(220)
        self.page_info.setMinimumHeight(46)
        self.page_info.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        self.page_info.setAlignment(QtCore.Qt.AlignCenter)
        self.page_info.setStyleSheet("""
            QLabel {
                background: #e6f2fb; color: #12608d; border-radius: 8px; padding: 5px 12px;
                font-family: 'Times New Roman'; font-size: 11pt; font-weight: bold;
            }
        """)

        # Fill the complete second toolbar row; no unused white space at either side.
        nav_row.addWidget(self.btn_prev, 2)
        nav_row.addWidget(self.btn_next, 2)
        nav_row.addWidget(self.page_jump_spin, 1)
        nav_row.addWidget(self.btn_go_page, 1)
        nav_row.addWidget(self.page_info, 3)
        nav_row.addWidget(self.btn_terminate, 2)

        toolbar_layout.addLayout(top_row)
        toolbar_layout.addLayout(nav_row)
        layout.addWidget(toolbar_frame)

        # Mode row
        mode_row = QtWidgets.QHBoxLayout()
        mode_row.setSpacing(8)

        self.btn_annotation = mkbtn("Annotation", "#2563eb", "#4a7df2")
        self.btn_bookmark = mkbtn("Bookmark", "#7c3aed", "#9b63f0")
        self.btn_review = mkbtn("Review", "#059669", "#21b58a")

        self.btn_annotation.setEnabled(False)
        self.btn_bookmark.setEnabled(False)
        self.btn_review.setEnabled(False)

        mode_row.addWidget(self.btn_annotation, 1)
        mode_row.addWidget(self.btn_bookmark, 1)
        mode_row.addWidget(self.btn_review, 1)
        layout.addLayout(mode_row)

        # Font settings: keep the standard/default view compact. Custom controls
        # appear only when the user chooses to override the default 10 / 12 sizes.
        font_settings_row = QtWidgets.QHBoxLayout()
        font_settings_row.setSpacing(10)

        self.chk_default_font_sizes = QtWidgets.QCheckBox(
            f"Use default font sizes ({DEFAULT_BASE_FONT_SIZE} / {DEFAULT_DOMAIN_FONT_SIZE})"
        )
        self.chk_default_font_sizes.setChecked(True)
        self.chk_default_font_sizes.setToolTip(
            "Checked: normal annotations use 10 pt and domain annotations use 12 pt. "
            "Uncheck to select both sizes independently."
        )
        self.chk_default_font_sizes.setStyleSheet(
            "QCheckBox { font-family: 'Times New Roman'; font-size: 11pt; "
            "color: #27496d; padding: 4px 2px; }"
        )

        self.font_controls_widget = QtWidgets.QWidget()
        custom_font_layout = QtWidgets.QHBoxLayout(self.font_controls_widget)
        custom_font_layout.setContentsMargins(0, 0, 0, 0)
        custom_font_layout.setSpacing(8)

        custom_font_layout.addWidget(QtWidgets.QLabel("Annotation Font:"))
        self.cmb_base_font = QtWidgets.QComboBox()
        self.cmb_base_font.addItems([str(v) for v in FONT_SIZE_OPTIONS])
        self.cmb_base_font.setCurrentText(str(DEFAULT_BASE_FONT_SIZE))
        self.cmb_base_font.setFixedWidth(72)
        custom_font_layout.addWidget(self.cmb_base_font)

        custom_font_layout.addWidget(QtWidgets.QLabel("Domain Font:"))
        self.cmb_domain_font = QtWidgets.QComboBox()
        self.cmb_domain_font.addItems([str(v) for v in FONT_SIZE_OPTIONS])
        self.cmb_domain_font.setCurrentText(str(DEFAULT_DOMAIN_FONT_SIZE))
        self.cmb_domain_font.setFixedWidth(72)
        custom_font_layout.addWidget(self.cmb_domain_font)

        self.font_controls_widget.setVisible(False)
        font_settings_row.addWidget(self.chk_default_font_sizes)
        font_settings_row.addWidget(self.font_controls_widget)
        font_settings_row.addStretch()
        layout.addLayout(font_settings_row)

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

        self.btn_edit_ann = mkbtn("Edit Selected Annotation", "#0f766e", "#159287")
        self.btn_copy_ann = mkbtn("Copy Selected Annotation", "#0369a1", "#0ea5e9")
        self.btn_copy_ann_page = mkbtn("Copy Annotation To Page", "#0f766e", "#159287")
        self.btn_line_mode = mkbtn("Draw Connector Line", "#b45309", "#c97519")
        self.btn_clear_line = mkbtn("Clear Connector Line", "#475569", "#64748b")
        self.btn_delete_ann = mkbtn("Delete Selected Annotation(s)", "#ef4444", "#f87171")

        ann_buttons = [
            self.btn_edit_ann, self.btn_copy_ann, self.btn_copy_ann_page,
            self.btn_line_mode, self.btn_clear_line, self.btn_delete_ann
        ]
        for i, btn in enumerate(ann_buttons):
            ann_btn_grid.addWidget(btn, i // 3, i % 3)
        for col in range(3):
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

        self.btn_edit_bm = mkbtn("Edit Selected Bookmark", "#0f766e", "#159287")
        self.btn_copy_bm = mkbtn("Copy Selected Bookmark", "#0369a1", "#0ea5e9")
        self.btn_delete_bm = mkbtn("Delete Selected Bookmark", "#ef4444", "#f87171")

        bm_buttons = [
            self.btn_edit_bm, self.btn_copy_bm, self.btn_delete_bm
        ]
        for i, btn in enumerate(bm_buttons):
            bm_btn_grid.addWidget(btn, 0, i)
        for col in range(3):
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
        self.btn_import_excel.clicked.connect(self.import_project_excel)
        self.btn_export_excel.clicked.connect(lambda: self.export_project_excel())
        self.btn_load_sdtm.clicked.connect(self.load_sdtm_metadata)
        self.btn_prev.clicked.connect(self.prev_page)
        self.btn_next.clicked.connect(self.next_page)
        self.btn_go_page.clicked.connect(self.go_to_page)
        self.page_jump_spin.editingFinished.connect(self.go_to_page)
        self.btn_terminate.clicked.connect(self.close)

        self.btn_annotation.clicked.connect(lambda: self.switch_mode("annotation"))
        self.btn_bookmark.clicked.connect(lambda: self.switch_mode("bookmark"))
        self.btn_review.clicked.connect(self.open_review_safe)

        self.btn_generate_pdf.clicked.connect(self.generate_final_output_pdf)

        self.chk_default_font_sizes.toggled.connect(self.on_default_font_sizes_toggled)
        self.cmb_base_font.currentTextChanged.connect(self.on_custom_font_size_changed)
        self.cmb_domain_font.currentTextChanged.connect(self.on_custom_font_size_changed)

        self.annotation_table.itemSelectionChanged.connect(self.on_annotation_selection_changed)
        self.bookmark_table.itemSelectionChanged.connect(self.on_bookmark_selection_changed)

        self.review_annotation_table.itemSelectionChanged.connect(self.on_review_annotation_selection_changed)
        self.review_bookmark_table.itemSelectionChanged.connect(self.on_review_bookmark_selection_changed)

        self.btn_edit_ann.clicked.connect(self.edit_selected_annotation)
        self.btn_copy_ann.clicked.connect(self.copy_selected_annotation)
        self.btn_copy_ann_page.clicked.connect(self.copy_annotation_to_page)
        self.btn_edit_bm.clicked.connect(self.edit_selected_bookmark)
        self.btn_copy_bm.clicked.connect(self.copy_selected_bookmark)
        self.btn_delete_ann.clicked.connect(self.delete_selected_annotation)
        self.btn_line_mode.clicked.connect(self.start_connector_line_mode)
        self.btn_clear_line.clicked.connect(self.clear_connector_lines)

        self.btn_delete_bm.clicked.connect(self.delete_selected_bookmark)

        self.switch_mode("annotation", force=True)
        self.update_annotation_action_buttons()

    def on_default_font_sizes_toggled(self, checked):
        """Switch between the original default font sizes and custom document-wide sizes."""
        self.font_controls_widget.setVisible(not checked)

        if checked:
            self.base_font_size = DEFAULT_BASE_FONT_SIZE
            self.domain_font_size = DEFAULT_DOMAIN_FONT_SIZE
        else:
            self.base_font_size = int(self.cmb_base_font.currentText())
            self.domain_font_size = int(self.cmb_domain_font.currentText())

        # Repaint immediately so annotation boxes, wrapping and text reflect the active sizes.
        if hasattr(self, "image_label"):
            self.image_label.update()

    def on_custom_font_size_changed(self, _value=None):
        """Apply custom font sizes only while default-font mode is disabled."""
        if self.chk_default_font_sizes.isChecked():
            return

        self.base_font_size = int(self.cmb_base_font.currentText())
        self.domain_font_size = int(self.cmb_domain_font.currentText())

        if hasattr(self, "image_label"):
            self.image_label.update()

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
    # CDISC SDTM metadata load
    # ------------------------------------------------------------------
    def load_sdtm_metadata(self):
        standard_text = self.sdtmig_combo.currentText().strip() if hasattr(self, "sdtmig_combo") else self.sdtm_standard_text
        try:
            QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
            QtWidgets.QApplication.processEvents()
            meta = load_sdtmig_metadata_from_cdisc(standard_text)
            self.sdtm_standard_text = standard_text
            self.sdtm_domain_labels = meta["domain_labels"]
            self.sdtm_variables_by_domain = meta["variables_by_domain"]
            self.sdtm_metadata_loaded = True
            dom_count = len(self.sdtm_domain_labels)
            var_count = sum(len(v) for v in self.sdtm_variables_by_domain.values())
            QtWidgets.QMessageBox.information(
                self,
                "SDTM Metadata Loaded",
                f"Loaded {dom_count} domains and {var_count} variables from {standard_text}.\n\n"
                "Annotation dialog now uses SDTM domain and variable dropdowns."
            )
        except Exception as e:
            self.sdtm_metadata_loaded = False
            QtWidgets.QMessageBox.critical(self, "SDTM Metadata Load Failed", str(e))
        finally:
            QtWidgets.QApplication.restoreOverrideCursor()

    # ------------------------------------------------------------------
    # PDF open / render
    # ------------------------------------------------------------------
    def open_pdf(self):
        # Use the same native QFileDialog behaviour as the final Save PDF dialog.
        # This avoids the slow/non-opening non-native dialog seen on some Windows/network setups.
        start_dir = getattr(self, "last_pdf_dir", "") or (os.path.dirname(self.open_pdf_path) if self.open_pdf_path else PROJECT_DIR)
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Open PDF",
            start_dir,
            "PDF Files (*.pdf)"
        )
        if not file_path:
            return
        self.last_pdf_dir = os.path.dirname(file_path)
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
            self.btn_import_excel.setEnabled(True)
            self.btn_export_excel.setEnabled(True)
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

        orientation = "Landscape" if self.page_rect.width > self.page_rect.height else "Portrait"
        self.page_info.setText(f"Page: {self.current_page_index + 1} / {len(self.doc)} ({orientation})")
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
        self.btn_delete_ann.setEnabled(has_any_selection)
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
            src_rect = self.get_pdf_page_rect(src.pageno)
            if src_rect is not None:
                x1 = round(min(x1, max(0.0, float(src_rect.width) - MIN_BOX_WIDTH_SINGLE - RIGHT_PAGE_MARGIN)), 6)
                y1 = round(min(y1, max(0.0, float(src_rect.height) - BOX_HEIGHT_NORMAL - 6.0)), 6)

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
        selected_rows = self.get_selected_annotation_rows()

        # Fallback for cases where currentRow is set but selectedRows is empty.
        if not selected_rows:
            row = self.annotation_table.currentRow()
            if 0 <= row < len(self.entries):
                selected_rows = [row]

        if not selected_rows:
            QtWidgets.QMessageBox.information(self, "Delete Annotation", "Please select one or more annotation rows to delete.")
            return

        selected_rows = sorted(set(selected_rows))
        delete_count = len(selected_rows)
        if delete_count == 1:
            confirm_text = "Delete the selected annotation?"
        else:
            confirm_text = f"Delete the selected {delete_count} annotations?"

        reply = QtWidgets.QMessageBox.question(
            self,
            "Delete Annotation",
            confirm_text,
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No
        )
        if reply != QtWidgets.QMessageBox.Yes:
            return

        # Delete from bottom to top so row indexes do not shift while deleting.
        for row in sorted(selected_rows, reverse=True):
            if 0 <= row < len(self.entries):
                del self.entries[row]

        self.selected_entry_index = -1
        self.selected_line_index = -1
        self.annotation_table.clearSelection()
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
    # Excel export
    # ------------------------------------------------------------------
    def get_pdf_page_rect(self, pageno: int):
        if not self.doc or pageno < 1 or pageno > len(self.doc):
            return None
        try:
            return self.doc[pageno - 1].rect
        except Exception:
            return None

    def get_default_annotation_position(self, pageno: int, loaded_entries=None):
        """Return a default annotation position for future workbook import support."""
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

    def _build_variable_export_rows(self, page_offset=0):
        variable_rows = OrderedDict()

        for e in self.entries:
            domain = (e.domain or "").strip()
            name = (e.name or "").strip()

            if not domain or not name:
                continue
            if domain.upper() == "REF" and name.upper() == "REF":
                continue
            if e.is_domain_annotation or e.is_assigned_field or e.is_not_submitted:
                continue

            key = (domain.upper(), name.upper())
            if key not in variable_rows:
                variable_rows[key] = {
                    "DOMAIN": domain,
                    "VARIABLE": name,
                    "pages": set(),
                }
            variable_rows[key]["pages"].add(int(e.pageno) + int(page_offset))

        rows = []
        for item in variable_rows.values():
            rows.append([
                item["DOMAIN"],
                item["VARIABLE"],
                ",".join(str(p) for p in sorted(item["pages"])),
            ])
        return rows

    @staticmethod
    def _format_excel_sheet(ws, widths=None):
        header_fill = PatternFill("solid", fgColor="D9EAF7")
        header_font = Font(bold=True)

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

        if widths:
            for idx, width in enumerate(widths, start=1):
                ws.column_dimensions[get_column_letter(idx)].width = width

    def import_project_excel(self):
        """Import a complete AnnotateCRF Excel workbook into the current PDF project."""
        if not self.doc:
            QtWidgets.QMessageBox.warning(self, "Import Excel", "Please open the source PDF before importing Excel data.")
            return

        start_dir = os.path.dirname(self.open_pdf_path) if self.open_pdf_path else PROJECT_DIR
        in_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Import AnnotateCRF Excel", start_dir, "Excel Workbook (*.xlsx)"
        )
        if not in_path:
            return

        choice = QtWidgets.QMessageBox(self)
        choice.setWindowTitle("Import Excel")
        choice.setText("How should the Excel data be imported?")
        replace_btn = choice.addButton("Load / Replace", QtWidgets.QMessageBox.AcceptRole)
        append_btn = choice.addButton("Append", QtWidgets.QMessageBox.ActionRole)
        choice.addButton(QtWidgets.QMessageBox.Cancel)
        choice.exec_()
        clicked = choice.clickedButton()
        if clicked not in (replace_btn, append_btn):
            return
        append_mode = clicked is append_btn

        try:
            wb = load_workbook(in_path, data_only=True)
            required = {"Annotations", "Bookmarks", "Connector Lines"}
            missing = required.difference(wb.sheetnames)
            if missing:
                raise ValueError("Missing required worksheet(s): " + ", ".join(sorted(missing)))

            def sheet_records(ws):
                rows = list(ws.iter_rows(values_only=True))
                if not rows:
                    return []
                headers = [str(v or "").strip().upper() for v in rows[0]]
                return [dict(zip(headers, row)) for row in rows[1:] if any(v is not None and str(v).strip() != "" for v in row)]

            imported_entries = []
            for rec in sheet_records(wb["Annotations"]):
                domain = str(rec.get("DOMAIN") or "").strip()
                name = str(rec.get("NAME") or "").strip()
                annotation = str(rec.get("ANNOTATION") or "").strip()
                assignedfield = str(rec.get("ASSIGNEDFIELD") or "").strip()
                pageno = int(clean_number(rec.get("PAGENO"), 0) or 0)
                if pageno < 1 or pageno > len(self.doc):
                    continue

                page_rect = self.get_pdf_page_rect(pageno)
                default_x, default_y, default_pageh = self.get_default_annotation_position(
                    pageno, (self.entries + imported_entries) if append_mode else imported_entries
                )
                x1 = clean_number(rec.get("X1"), default_x)
                y1 = clean_number(rec.get("Y1"), default_y)
                pageh = clean_number(rec.get("PAGEH"), default_pageh)
                box_w = clean_number(rec.get("BOX_W"), None)
                box_h = clean_number(rec.get("BOX_H"), None)
                is_domain, is_assigned, is_notsub = bool_from_entry(domain, name, annotation, assignedfield)

                imported_entries.append(AnnotationEntry(
                    domain=domain, name=name, pageno=pageno, annotation=annotation,
                    assignedfield=assignedfield, x1=float(x1), y1=float(y1), pageh=float(pageh),
                    is_domain_annotation=is_domain, is_assigned_field=is_assigned,
                    is_not_submitted=is_notsub,
                    box_w=None if box_w is None else float(box_w),
                    box_h=None if box_h is None else float(box_h),
                ))

            imported_bookmarks = []
            for rec in sheet_records(wb["Bookmarks"]):
                title = str(rec.get("TITLE") or "").strip()
                level = int(clean_number(rec.get("LEVEL"), 1) or 1)
                pageno = int(clean_number(rec.get("PAGENO"), 0) or 0)
                if title and 1 <= pageno <= len(self.doc):
                    imported_bookmarks.append(BookmarkEntry(title=title, level=max(1, level), pageno=pageno))

            imported_lines = []
            for rec in sheet_records(wb["Connector Lines"]):
                pageno = int(clean_number(rec.get("PAGENO"), 0) or 0)
                if not (1 <= pageno <= len(self.doc)):
                    continue
                coords = [clean_number(rec.get(k), None) for k in ("X1", "Y1", "X2", "Y2")]
                if any(v is None for v in coords):
                    continue
                imported_lines.append(ConnectorLineEntry(
                    pageno=pageno, x1=float(coords[0]), y1=float(coords[1]),
                    x2=float(coords[2]), y2=float(coords[3])
                ))

            if append_mode:
                ann_keys = {annotation_entry_key(e) for e in self.entries}
                bm_keys = {bookmark_entry_key(b) for b in self.bookmarks}
                line_keys = {connector_line_key(ln) for ln in self.lines}
                added_ann = [e for e in imported_entries if annotation_entry_key(e) not in ann_keys]
                added_bm = [b for b in imported_bookmarks if bookmark_entry_key(b) not in bm_keys]
                added_lines = [ln for ln in imported_lines if connector_line_key(ln) not in line_keys]
                self.entries.extend(added_ann)
                self.bookmarks.extend(added_bm)
                self.lines.extend(added_lines)
            else:
                self.entries = imported_entries
                self.bookmarks = imported_bookmarks
                self.lines = imported_lines
                added_ann, added_bm, added_lines = imported_entries, imported_bookmarks, imported_lines

            self.has_annotation = bool(self.entries)
            self.has_bookmark = bool(self.bookmarks)
            self.selected_entry_index = -1
            self.selected_bookmark_index = -1
            self.selected_line_index = -1
            self.refresh_annotation_table()
            self.refresh_bookmark_table()
            self.refresh_review_tables()
            self.check_review_enable()
            self.update_annotation_action_buttons()
            self.render_page()

            QtWidgets.QMessageBox.information(
                self, "Import Excel",
                f"Excel workbook imported successfully.\n\n"
                f"Annotations added: {len(added_ann)}\n"
                f"Bookmarks added: {len(added_bm)}\n"
                f"Connector Lines added: {len(added_lines)}"
            )
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Import Excel", f"Failed to import Excel workbook:\n{e}")

    def export_project_excel(self, out_path=None, show_message=True, page_offset=0, create_blank_if_none=True):
        """Export the complete current annotation project to one Excel workbook.

        The same export is used from both the Annotation and Bookmark pages and always
        contains the current Annotations, Bookmarks, Connector Lines, and Variables.
        """
        if isinstance(out_path, bool):
            out_path = None

        if not self.entries and not self.bookmarks and not self.lines and not create_blank_if_none:
            if show_message:
                QtWidgets.QMessageBox.information(self, "Export Excel", "No annotation data available to export.")
            return False

        if out_path is None:
            if self.open_pdf_path:
                pdf_stem = os.path.splitext(os.path.basename(self.open_pdf_path))[0]
                default_name = f"{pdf_stem}_annotatecrf.xlsx"
                default_dir = os.path.dirname(self.open_pdf_path)
            else:
                default_name = "annotatecrf_workbook.xlsx"
                default_dir = PROJECT_DIR

            default_path = os.path.join(default_dir, default_name)
            out_path, _ = QtWidgets.QFileDialog.getSaveFileName(
                self,
                "Export AnnotateCRF Excel",
                default_path,
                "Excel Workbook (*.xlsx)",
            )
            if not out_path:
                return False

        out_path = ensure_xlsx_extension(out_path)

        try:
            _ensure_console_safe_streams()

            wb = Workbook()
            ws_ann = wb.active
            ws_ann.title = "Annotations"
            ws_bm = wb.create_sheet("Bookmarks")
            ws_lines = wb.create_sheet("Connector Lines")
            ws_vars = wb.create_sheet("Variables")

            ws_ann.append(ANNOTATION_EXCEL_COLUMNS)
            for e in self.entries:
                ws_ann.append([
                    e.domain,
                    e.name,
                    int(e.pageno) + int(page_offset),
                    adjust_page_refs_in_text(e.annotation, page_offset),
                    e.assignedfield,
                    e.x1,
                    e.y1,
                    e.pageh,
                    None if e.box_w is None else e.box_w,
                    None if e.box_h is None else e.box_h,
                ])

            ws_bm.append(BOOKMARK_EXCEL_COLUMNS)
            for b in self.bookmarks:
                ws_bm.append([
                    b.title,
                    b.level,
                    int(b.pageno) + int(page_offset),
                ])

            ws_lines.append(CONNECTOR_LINE_EXCEL_COLUMNS)
            for ln in self.lines:
                ws_lines.append([
                    int(ln.pageno) + int(page_offset),
                    ln.x1,
                    ln.y1,
                    ln.x2,
                    ln.y2,
                ])

            ws_vars.append(VARIABLE_EXCEL_COLUMNS)
            for row in self._build_variable_export_rows(page_offset=page_offset):
                ws_vars.append(row)

            self._format_excel_sheet(ws_ann, [14, 18, 10, 48, 16, 12, 12, 12, 12, 12])
            self._format_excel_sheet(ws_bm, [50, 10, 10])
            self._format_excel_sheet(ws_lines, [10, 14, 14, 14, 14])
            self._format_excel_sheet(ws_vars, [14, 20, 30])

            atomic_excel_write(out_path, wb)

            if show_message:
                QtWidgets.QMessageBox.information(
                    self,
                    "Export Excel",
                    "Excel workbook exported successfully:\n"
                    f"{out_path}\n\n"
                    f"Annotations: {len(self.entries)}\n"
                    f"Bookmarks: {len(self.bookmarks)}\n"
                    f"Connector Lines: {len(self.lines)}\n"
                    f"Variables: {len(self._build_variable_export_rows(page_offset=page_offset))}",
                )
            return True

        except Exception as e:
            if show_message:
                QtWidgets.QMessageBox.critical(
                    self,
                    "Export Excel",
                    "Failed to export Excel workbook:\n"
                    f"{e}\n\n"
                    "Please check that the workbook is not already open in another application "
                    "and that the selected folder is writable.",
                )
            return False

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
        doc = None

        try:
            QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)

            # Open the original PDF as-is.
            # Do NOT normalize/rebuild/rotate pages and do NOT delete existing PDF
            # annotation/comment objects. We only add a visual overlay per page.
            doc = fitz.open(self.open_pdf_path)

            pdf_base = get_pdf_base_output_path(output_pdf)
            excel_workbook_path = pdf_base + "_annotatecrf.xlsx"

            add_toc_checked = bool(self.chk_add_toc.isChecked() and self.bookmarks)
            toc_page_offset = 0
            if add_toc_checked:
                toc_page_offset = compute_toc_page_count_from_bookmarks(
                    self.bookmarks,
                    source_doc=doc,
                    font_size=12,
                    lines_per_page=38
                )

            # Draw page-by-page. Each page gets an overlay with the same visible
            # width/height as the GUI preview. This keeps portrait and landscape
            # pages independent in the same output PDF.
            for page_index in range(len(doc)):
                page_no = page_index + 1
                page = doc[page_index]

                page_entries = [e for e in self.entries if e.pageno == page_no]
                page_lines = [ln for ln in self.lines if ln.pageno == page_no]

                if not page_entries and not page_lines:
                    continue

                visible_width = float(page.rect.width)
                visible_height = float(page.rect.height)

                overlay_doc = fitz.open()
                link_rects = []

                try:
                    overlay_page = overlay_doc.new_page(width=visible_width, height=visible_height)
                    color_map = get_page_domain_color_map(self.entries, page_no)

                    for e in page_entries:
                        layout = compute_entry_layout(e, color_map, page_width=visible_width, base_font_size=self.base_font_size, domain_font_size=self.domain_font_size)

                        annotation_text_for_output = adjust_page_refs_in_text(e.annotation, toc_page_offset)
                        layout_for_output = dict(layout)
                        layout_for_output["lines"] = wrap_text_by_width(
                            annotation_text_for_output,
                            layout["box_w"],
                            scale=(self.domain_font_size / self.base_font_size if layout["bold"] else 1.0)
                        )

                        output_base_h = BOX_HEIGHT_DOMAIN if layout["bold"] else BOX_HEIGHT_NORMAL
                        output_box_h = compute_output_box_height(
                            layout_for_output["lines"],
                            layout["box_w"],
                            layout["bold"],
                            layout["font_size"],
                            output_base_h
                        )

                        # Coordinates are the GUI/visible coordinates stored in the project data.
                        rect = fitz.Rect(
                            float(e.x1),
                            float(e.y1),
                            float(e.x1) + float(layout["box_w"]),
                            float(e.y1) + float(output_box_h)
                        )
                        rect = clamp_rect_to_page(rect, visible_width, visible_height)

                        draw_box_and_text_pdf(
                            page=overlay_page,
                            rect=rect,
                            text_lines=layout_for_output["lines"],
                            fill_color=layout["fill"],
                            bold=layout["bold"],
                            dashed=layout["dashed"],
                            font_size=layout["font_size"],
                            link_target_page=None
                        )

                        page_ref = extract_page_reference(e.annotation)
                        if page_ref:
                            link_rects.append((rect, page_ref))

                        if None not in (e.line_x1, e.line_y1, e.line_x2, e.line_y2):
                            overlay_page.draw_line(
                                fitz.Point(e.line_x1, e.line_y1),
                                fitz.Point(e.line_x2, e.line_y2),
                                color=(1, 0, 0),
                                width=1.2,
                                overlay=True
                            )

                    for ln in page_lines:
                        overlay_page.draw_line(
                            fitz.Point(ln.x1, ln.y1),
                            fitz.Point(ln.x2, ln.y2),
                            color=(1, 0, 0),
                            width=1.2,
                            overlay=True
                        )

                    # Place the visual overlay on the original page.
                    # For rotated pages, this uses the source page's same rotation.
                    place_visual_overlay_on_original_page(page, overlay_doc)

                    # Add clickable refer-page links on the original page.
                    for visual_rect, target_page in link_rects:
                        try:
                            page.insert_link({
                                "kind": fitz.LINK_GOTO,
                                "from": visual_rect_to_native_rect(page, visual_rect),
                                "page": max(0, min(int(target_page) - 1, len(doc) - 1)),
                                "to": fitz.Point(72, 72)
                            })
                        except Exception:
                            pass

                finally:
                    overlay_doc.close()

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
                doc = None

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
                doc = None

            # Create one Excel workbook matching the page numbering of the final PDF.
            self.export_project_excel(
                excel_workbook_path,
                show_message=False,
                page_offset=toc_page_offset,
                create_blank_if_none=True
            )

            what_written = []
            if self.entries:
                what_written.append(f"{len(self.entries)} annotation(s)")
            if self.bookmarks:
                what_written.append(f"{len(self.bookmarks)} bookmark(s)")
            if self.lines:
                what_written.append(f"{len(self.lines)} connector line(s)")
            what_written.append("1 Excel workbook")
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
                if doc is not None:
                    doc.close()
            except Exception:
                pass
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
