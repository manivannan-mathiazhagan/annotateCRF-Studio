# AnnotateCRF Studio

## Overview

AnnotateCRF Studio is a desktop GUI utility for creating and reviewing CDISC-compliant annotated CRFs (aCRFs) without relying on licensed PDF editing tools.

The application is built with Python, PyMuPDF, and PyQt5 and is designed to support Windows, macOS, and Linux.

AnnotateCRF Studio supports:

- Annotation box creation and positioning
- Domain-based color coding
- Bookmark creation
- Connector lines
- Excel import and export for annotations, bookmarks, and connector lines
- Append and combine annotation work from multiple users
- Multi-user annotation and review workflows
- Clickable internal page-reference links
- CDISC Library API integration
- Configurable annotation font sizes with default 10 pt / 12 pt behavior
- Final annotated PDF generation

---

## Repository Structure

```text
AnnotateCRF-STUDIO/
│
├── scripts/
│   ├── annotatecrf_studio.py
│   └── annotatecrf_studio.pyw
│
├── annotator_icon.ico
├── annotator_icon.png
│
├── CDISC_API_KEY.example.json
├── CDISC_API_KEY.json
│
├── requirements.txt
├── README.md
└── .gitignore
```

### Main Application

```text
scripts/annotatecrf_studio.py
```

This is the main application and contains the AnnotateCRF Studio functionality.

The same main Python application is used across Windows, macOS, and Linux.

### Windows Launcher

```text
scripts/annotatecrf_studio.pyw
```

This is a lightweight Windows launcher that starts the main application without displaying a console window.

The application functionality is maintained in the `.py` file so that separate versions of the application do not need to be maintained for different operating systems.

---

## Requirements

Python must be installed before running AnnotateCRF Studio.

Required Python packages are listed in:

```text
requirements.txt
```

The current core dependencies are:

```text
PyMuPDF
PyQt5
openpyxl
```

These packages provide PDF processing, the desktop GUI, and Excel workbook import/export functionality.

---

## Installation

### 1. Clone the Repository

Clone the repository using Git:

```bash
git clone <repository-url>
```

Move into the repository folder:

```bash
cd AnnotateCRF-STUDIO
```

### 2. Install Required Python Packages

Install all required packages from `requirements.txt`.

On Windows:

```bash
python -m pip install -r requirements.txt
```

On macOS or Linux:

```bash
python3 -m pip install -r requirements.txt
```

`pip` reads `requirements.txt` and installs the Python packages required by AnnotateCRF Studio.

Installation is normally required only when setting up the application for the first time or when the package requirements change.

### 3. Configure the CDISC Library API

If CDISC Library functionality is required, create the local API configuration file as described in the CDISC Library API Configuration section below.

### 4. Start AnnotateCRF Studio

Use the appropriate launch method for your operating system.

---

## How to Run

### Windows

The recommended method is to double-click:

```text
scripts/annotatecrf_studio.pyw
```

This opens the GUI without displaying a console window.

If `.pyw` files are not associated with Python:

1. Right-click `annotatecrf_studio.pyw`
2. Select **Open with**
3. Select **Choose another app**
4. Browse to `pythonw.exe`
5. Select **Always use this app to open .pyw files**

Use `pythonw.exe` rather than `python.exe` so that a console window is not displayed.

### Windows – Development or Troubleshooting

Run the main Python application directly:

```bash
python scripts/annotatecrf_studio.py
```

Running the `.py` file directly is useful for development and troubleshooting because Python errors and diagnostic messages remain visible in the terminal.

### macOS

Run:

```bash
python3 scripts/annotatecrf_studio.py
```

### Linux

Run:

```bash
python3 scripts/annotatecrf_studio.py
```

The `.pyw` launcher is intended for Windows. macOS and Linux users should run the main `.py` application.

---

## CDISC Library API Configuration

AnnotateCRF Studio supports integration with the CDISC Library API.

The repository includes an example API configuration file:

```text
CDISC_API_KEY.example.json
```

Create a copy of this file and name it:

```text
CDISC_API_KEY.json
```

Enter your CDISC Library API Primary and Secondary keys in the file.

Example:

```json
{
    "primary_key": "YOUR_PRIMARY_KEY",
    "secondary_key": "YOUR_SECONDARY_KEY"
}
```

`CDISC_API_KEY.json` contains the user's API credentials and should be kept private.

The application reads the API credentials from the local configuration file when CDISC Library functionality is used.

---

## Main Features

### Annotation Mode

Annotation Mode provides the primary aCRF annotation workflow.

Users can:

- Open a source PDF
- Click on a PDF page to add an annotation
- Enter annotation metadata through the GUI
- Preview annotation boxes immediately
- Drag annotation boxes to reposition them
- Resize selected annotation boxes
- Edit existing annotations
- Copy annotations
- Delete annotations
- Review annotation information before final generation

---

## Annotation Types

### Variable Annotation

Standard annotation for a collected field or variable.

### Domain Annotation

Used when an annotation applies to a domain rather than a specific variable.

### Assigned Field

Supports assigned-field annotations with dedicated visual styling.

### Not Submitted

Creates a standard:

```text
[NOT SUBMITTED]
```

annotation.

### Refer Page

Creates a page-reference annotation such as:

```text
For annotation refer to page X from collected field.
```

The final PDF can also contain an internal clickable link to the referenced page.

---

## Bookmark Mode

Bookmark Mode allows PDF bookmarks to be created and maintained from within the application.

Users can:

- Add bookmarks
- Specify bookmark titles
- Specify bookmark levels
- Specify target page numbers
- Edit bookmarks
- Copy bookmarks
- Delete bookmarks
- Include bookmarks in the final annotated PDF

---

## Review Mode

Review Mode provides a consolidated view of annotation and bookmark information.

It can be used to:

- Review annotations
- Review bookmarks
- Check annotation information before final generation
- Identify entries requiring adjustment
- Support collaborative review workflows

---

## Connector Lines

Connector lines can be created independently from annotation boxes.

Users can:

- Add connector lines
- Move individual line endpoints
- Move complete connector lines
- Adjust connector-line positioning
- Retain connector-line information during export and import

Connector lines are included in the final annotated PDF.

---

## Excel Import and Export

AnnotateCRF Studio supports Excel-based import and export of annotation, bookmark, and connector-line information.

Excel provides a single structured workbook for saving, reviewing, sharing, and combining annotation work across multiple users.

Excel files are read and written directly using Python through `openpyxl`.

### Excel Export

Annotation work can be exported to a single `.xlsx` workbook.

The workbook can contain separate worksheets for:

```text
Annotations
Bookmarks
Connector Lines
```

This allows the information required to recreate the annotation work to be stored in one file.

The exported workbook can be used as a backup, shared with another programmer, reviewed outside the application, used to continue annotation work on another computer, or imported back into AnnotateCRF Studio.

### Excel Import

AnnotateCRF Studio can load a previously exported Excel workbook.

During import, the application reads the relevant worksheets and restores available information such as:

- Annotations
- Annotation positions
- Bookmarks
- Connector lines
- Connector-line positions
- Page-reference information

Imported information can then be reviewed, repositioned, edited, copied, or deleted through the GUI.

### Append Excel Files

Excel workbooks from multiple users can be appended to support collaborative annotation workflows.

Example:

```text
Programmer A
     │
     └── annotation_work_A.xlsx
                  │
                  │
Programmer B      │
     │            │
     └── annotation_work_B.xlsx
                  │
                  ▼
          AnnotateCRF Studio
                  │
                  ▼
        Combined Annotation Set
```

Duplicate checking can be applied when annotation information is loaded or appended to help prevent duplicate entries.

---

## Annotation Worksheet

The `Annotations` worksheet contains annotation information such as:

```text
TYPE
DOMAIN
NAME
PAGENO
ANNOTATION
ASSIGNEDFIELD
X1
Y1
PAGEH
BOX_W
BOX_H
```

Core annotation information includes the annotation type, domain, variable or annotation name, page number, annotation text, and assigned-field indicator.

Position information includes X position, Y position, page height, annotation box width, and annotation box height.

If position information is unavailable, the application can place the annotation at a default location so that it can subsequently be positioned using the GUI.

---

## Bookmarks Worksheet

The `Bookmarks` worksheet contains:

```text
TITLE
LEVEL
PAGENO
```

where:

- `TITLE` contains the bookmark text
- `LEVEL` defines the bookmark hierarchy
- `PAGENO` defines the target PDF page

---

## Connector Lines Worksheet

The `Connector Lines` worksheet stores connector-line information such as:

```text
PAGENO
X1
Y1
X2
Y2
```

This allows connector-line positioning to be retained when annotation work is exported and subsequently imported.

---

## Cross-Platform Excel Support

AnnotateCRF Studio reads and writes `.xlsx` workbooks directly using Python and does not require Microsoft Excel to be installed.

Excel import and export functionality is designed to work on:

- Windows
- macOS
- Linux

To manually open or review a generated `.xlsx` file outside AnnotateCRF Studio, an Excel-compatible spreadsheet application is required.

Examples include:

- Microsoft Excel
- Apple Numbers
- LibreOffice Calc
- Other applications capable of reading `.xlsx` files

The available spreadsheet applications depend on the operating system and user environment.

---

## Annotation Font Size Settings

AnnotateCRF Studio provides document-wide font-size controls for annotation output.

By default, the application uses:

```text
Standard annotation font: 10 pt
Domain annotation font:   12 pt
```

The **Use default font sizes (10 / 12)** option is enabled by default. While it is enabled, the custom font-size controls remain hidden and the standard 10 pt / 12 pt settings are used.

When the default option is cleared, separate controls are displayed for:

- Standard annotation font size
- Domain annotation font size

Each font size can be selected independently from **8 pt through 14 pt**. This allows the annotation and domain font sizes to be adjusted separately when a CRF layout requires smaller or larger annotation text.

The selected font sizes are applied consistently to annotation display, text wrapping, annotation box sizing, and final PDF generation.

---

## Color and Display Behavior

AnnotateCRF Studio applies visual formatting to make annotations easier to identify and review.

Features include:

- Domain-based color coding
- Different colors for multiple domains appearing on a page
- Dedicated styling for assigned fields
- Dedicated highlighting for Not Submitted annotations
- Automatic text wrapping
- Automatic annotation box sizing
- Default annotation/domain font sizes of 10 pt / 12 pt
- Optional independent annotation and domain font-size controls from 8 pt to 14 pt
- Handling of both single-line and multi-line annotations

---

## Page-Reference Links

Page-reference annotations can point users to annotations located elsewhere in the CRF.

Example:

```text
For annotation refer to page 25 from collected field.
```

When the final PDF is generated, the page reference can contain an internal clickable link.

Selecting the link navigates directly to the referenced page within the PDF.

---

## Final PDF Output

AnnotateCRF Studio can generate a final annotated PDF containing:

- Annotation boxes
- Annotation text
- Connector lines
- PDF bookmarks
- Internal clickable page-reference links
- Searchable visible annotation content

The generated PDF can be used for aCRF review and downstream clinical programming workflows.

---

## Cross-Platform Support

AnnotateCRF Studio uses the same main Python application across supported operating systems.

### Windows

Normal GUI launch:

```text
scripts/annotatecrf_studio.pyw
```

Development or troubleshooting:

```bash
python scripts/annotatecrf_studio.py
```

### macOS

```bash
python3 scripts/annotatecrf_studio.py
```

### Linux

```bash
python3 scripts/annotatecrf_studio.py
```

Python-based PDF and Excel processing allows the core application functionality to operate without requiring licensed PDF editing software or Microsoft Excel.

---

## File and Path Handling

AnnotateCRF Studio is designed to use platform-independent file handling so that the same application can operate across Windows, macOS, and Linux.

Application resources, API configuration, PDF files, Excel workbooks, and generated outputs should be located using platform-independent Python path handling rather than operating-system-specific hard-coded paths.

---

## Troubleshooting

### Required Python Package Is Missing

Install or refresh the required packages using:

**Windows**

```bash
python -m pip install -r requirements.txt
```

**macOS / Linux**

```bash
python3 -m pip install -r requirements.txt
```

### `.pyw` Does Not Open on Windows

Confirm that Python is installed, `.pyw` is associated with `pythonw.exe`, and the packages in `requirements.txt` have been installed.

For troubleshooting, run:

```bash
python scripts/annotatecrf_studio.py
```

This allows Python errors to be viewed in the terminal.

### A Console Window Appears on Windows

The application may be running through `python.exe`.

For normal GUI use, launch:

```text
scripts/annotatecrf_studio.pyw
```

using `pythonw.exe`.

### Application Icon Does Not Appear

Confirm that the following files are available:

```text
annotator_icon.ico
annotator_icon.png
```

and that the application can locate them from the repository structure.

### Annotation Data Loads but Positions Are Different

Annotation coordinates correspond to the PDF layout from which they were created.

If position information is missing or differs from the target PDF layout, annotations may initially appear at a default or different position.

Use the GUI to move or resize the annotations as required.

### Excel Workbook Cannot Be Opened Manually

AnnotateCRF Studio itself can read and write `.xlsx` files without Microsoft Excel.

To manually open the workbook outside the application, an `.xlsx`-compatible spreadsheet application must be installed.

Examples include Microsoft Excel, Apple Numbers, and LibreOffice Calc.

### CDISC Library Functionality Does Not Work

Confirm that:

```text
CDISC_API_KEY.json
```

exists and contains valid CDISC Library API credentials.

The expected structure is:

```json
{
    "primary_key": "YOUR_PRIMARY_KEY",
    "secondary_key": "YOUR_SECONDARY_KEY"
}
```

### Application Fails on Another Machine

Confirm that Python is installed and then install the dependencies using:

```bash
python -m pip install -r requirements.txt
```

or:

```bash
python3 -m pip install -r requirements.txt
```

Also confirm that the required application files are present and that the user has permission to read input files and write to the selected output location.

---

## Change History

### September 2026

- Added configurable annotation font-size controls.
- Default font sizes remain **10 pt for standard annotations** and **12 pt for domain annotations**.
- Added a **Use default font sizes (10 / 12)** option that keeps the custom font controls hidden during normal use.
- When the default option is cleared, separate **Annotation Font** and **Domain Font** controls are displayed.
- Custom annotation and domain font sizes can be selected independently from **8 pt to 14 pt**.
- Font-size selections are applied consistently to annotation layout, wrapping, preview, and final PDF generation.

---

## Development

The main application source is:

```text
scripts/annotatecrf_studio.py
```

Use this file for development, debugging, testing, and feature changes.

The Windows launcher:

```text
scripts/annotatecrf_studio.pyw
```

should remain lightweight and only start the main application.

This avoids maintaining separate Windows and cross-platform versions of AnnotateCRF Studio.

When a new third-party Python package is introduced into the application, add it to:

```text
requirements.txt
```

so that the complete application environment can be installed using a single command.

---

## Quick Start

After cloning the repository:

### Windows

```bash
cd AnnotateCRF-STUDIO
python -m pip install -r requirements.txt
python scripts/annotatecrf_studio.py
```

For normal GUI use after setup, double-click:

```text
scripts/annotatecrf_studio.pyw
```

### macOS / Linux

```bash
cd AnnotateCRF-STUDIO
python3 -m pip install -r requirements.txt
python3 scripts/annotatecrf_studio.py
```

---

## Summary

AnnotateCRF Studio provides a Python-based desktop workflow for creating, reviewing, sharing, and generating annotated CRFs.

Key capabilities include:

- PDF annotation
- Domain and variable annotations
- Assigned-field annotations
- Not Submitted annotations
- Bookmark creation
- Connector lines
- Page-reference annotations
- Clickable internal PDF links
- Excel import and export
- Multi-user annotation workflows
- CDISC Library API integration
- Configurable annotation and domain font sizes
- Final annotated PDF generation
- Windows, macOS, and Linux support
