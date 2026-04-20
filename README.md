# AnnotateCRF Studio

## Overview
AnnotateCRF Studio is a desktop GUI utility for creating CDISC-compliant annotated CRFs (aCRFs) without relying on licensed PDF tools.

It supports:
- Annotation box creation with positioning
- Domain-based color coding
- Bookmark generation
- CSV-based import and export for multi-user workflows
- Final PDF generation with embedded annotations
- Clickable internal page-reference links
- A clean desktop GUI built with Python, PyMuPDF, and PyQt5

---

## How to Run

### Recommended
Double-click this file:

`annotatecrf_studio.pyw`

This opens the GUI directly without a black console window.

### If double-click does not work
The `.pyw` file association is not set on that machine. To fix it:

1. Right-click `annotatecrf_studio.pyw`
2. Choose **Open with**
3. Choose **Choose another app**
4. Browse to `pythonw.exe`
5. Check **Always use this app to open .pyw files**

Use `pythonw.exe`, not `python.exe`, so the app opens without a console window.

### Alternative launcher
You can also keep a BAT launcher outside the app folder.

Example structure:

```text
project_folder/
│
├── run_annotatecrf.bat
└── app/
    ├── annotatecrf_studio.pyw
    ├── annotator_icon.ico
    └── annotator_icon.png
```

Example BAT file:

```bat
@echo off
cd /d "%~dp0"
start "" pythonw "app\annotatecrf_studio.pyw"
exit
```

---

## Main Features

### Annotation Mode
- Open a PDF and click on a page to add an annotation
- Enter metadata through a dialog
- Preview annotation boxes immediately
- Drag boxes to reposition them
- Resize selected annotation boxes
- Edit, copy, and delete annotations
- Supports variable annotations, domain annotations, assigned fields, not-submitted notes, and page-reference annotations

### Bookmark Mode
- Add bookmarks directly from the PDF page
- Set bookmark title, level, and page number
- Edit, copy, and delete bookmarks
- Use bookmarks in final output PDF

### Review Mode
- Review annotation and bookmark entries side by side
- Validate entries before final generation

### Connector Lines
- Add connector lines separately
- Move line endpoints after creation
- Move whole connector lines
- Connector lines are retained in exported data and final output

### CSV Import and Export
- Export annotations to CSV
- Export bookmarks to CSV
- Load annotation CSV files
- Load bookmark CSV files
- Append multiple CSV files
- Duplicate checking is supported during load
- Missing position values can be left blank and adjusted later in the GUI

---

## Annotation Workflow

1. Open the source PDF
2. Go to **Annotation** mode
3. Click on the page where annotation is needed
4. Enter metadata in the dialog
5. Adjust box position if needed
6. Add bookmarks if needed
7. Review entries
8. Generate final output PDF

---

## Annotation Types Supported

### Variable Annotation
Standard annotation for a collected field.

### Domain Annotation
Used when the annotation applies to a domain rather than a specific variable.

### Assigned Field
Adds assigned-field behavior and visual styling.

### Not Submitted
Creates a standard `[NOT SUBMITTED]` annotation.

### Refer Page
Creates text like:

`For annotation refer to page X from collected field.`

This can also generate an internal clickable link in the final PDF.

---

## CSV Format

### Annotation CSV columns
The tool supports these internal annotation columns:

```text
TYPE,DOMAIN,NAME,PAGENO,ANNOTATION,ASSIGNEDFIELD,X1,Y1,PAGEH,BOX_W,BOX_H,LINE_PAGENO,LINE_X1,LINE_Y1,LINE_X2,LINE_Y2
```

Core fields:
- `DOMAIN`
- `NAME`
- `PAGENO`
- `ANNOTATION`
- `ASSIGNEDFIELD`

Optional fields:
- `X1`
- `Y1`
- `PAGEH`
- `BOX_W`
- `BOX_H`
- line-related columns

If `X1`, `Y1`, or `PAGEH` are blank, the tool places the annotation at a default location and you can drag it to the final position manually.

### Bookmark CSV columns

```text
TITLE,LEVEL,PAGENO
```

---

## Color and Display Behavior

- Domain-based color coding is applied by order of domain appearance on a page
- The first few domains on a page are color-coded distinctly
- Assigned fields use dashed borders
- Not Submitted annotations use dedicated highlighting
- Text wrapping and box sizing are automatically adjusted
- Single-line and multi-line annotations are handled differently for cleaner appearance

---

## Final Output

The tool can generate a final annotated PDF that includes:
- Visible annotation boxes
- Connector lines
- Bookmarks
- Internal clickable page-reference links
- Searchable flattened-style visible content for review workflows

The output file can also include exported CSV files for annotations and bookmarks, depending on workflow and usage.

---

## Dependencies

The script auto-installs required packages if missing:

- `PyMuPDF`
- `PyQt5`

Because of this, Python must be installed on the machine before running the tool.

---

## Platform Notes

### Windows
- `.pyw` is the recommended launch file
- `pythonw.exe` should be associated with `.pyw`
- A BAT launcher can be used for easier team access

### macOS
- Python setup may differ depending on system Python and Homebrew Python
- Package installation and GUI behavior may vary slightly across environments

---

## Recommended Folder Setup

```text
AnnotateCRF-Studio/
│
├── run_annotatecrf.bat
└── app/
    ├── annotatecrf_studio.pyw
    ├── annotator_icon.ico
    ├── annotator_icon.png
    └── other supporting files
```

This structure keeps the launcher separate from the application files.

---

## Troubleshooting

### The `.pyw` file does not open on double-click
The `.pyw` association is not configured. Associate it with `pythonw.exe`.

### A black console window appears
The file is being opened using `python.exe` instead of `pythonw.exe`.

### The app opens but icon does not appear
Make sure the icon file is in the same folder as the script and named as expected, such as:
- `annotator_icon.ico`
- `annotator_icon.png`

### CSV loads but entries do not appear where expected
If position columns are blank or different from the target PDF layout, the tool may place annotations at a default position. Drag them to the final location.

### App fails on another machine
Check:
- Python is installed
- Required packages can be installed
- `.pyw` is associated correctly
- file permissions allow reading and writing PDFs and CSVs

---

## Best Practice for Team Use

For personal use:
- Double-click `annotatecrf_studio.pyw`

For team or server use:
- Keep a BAT launcher
- Or package the app as an EXE for easier deployment

---

## Support

For queries, suggestions, or issues:

**Manivannan.Mathialagan@veristat.com**

---

## Summary

AnnotateCRF Studio is a desktop utility for annotation and bookmarking of PDFs for aCRF workflows.

It provides:
- PDF annotation
- Bookmark creation
- CSV import and export
- multi-user collaboration support
- final annotated PDF generation
- a GUI workflow without requiring Adobe-based editing tools

Preferred launch method:
- Double-click `annotatecrf_studio.pyw`
