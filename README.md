# Annotated CRF Studio

## Overview

**Annotated CRF Studio** is a Python-based desktop GUI application designed to enable **one-click annotation and bookmarking of CRFs** in a simple, efficient, and user-friendly workflow.

It supports **MSG 2.0–compliant annotated CRFs (aCRFs)** and streamlines annotation, review, and final PDF generation without external dependencies.

---

## Key Advantages

* No licensing required
* Enables multi-user annotation workflow using CSV files
* Eliminates file access and sharing conflicts during annotation
* Supports distributed annotation (page-wise or user-wise)
* Centralized merge using CSV import
* Lightweight, fast, and easy to deploy

---

## Key Features

### Annotation

* Click anywhere on the PDF to add annotations
* Supports:

  * Domain-level annotations
  * Variable-level annotations
  * Assigned fields (dashed border)
  * Not Submitted fields
* Drag-and-drop repositioning of annotation boxes
* Automatic text wrapping and dynamic box sizing
* Domain-based color coding

### Bookmarking

* Add bookmarks per page
* Define:

  * Bookmark text
  * Hierarchy level
  * Page number
* Visual preview in UI
* Export and import bookmark metadata

### Review Mode

* Side-by-side view of:

  * Annotation table
  * Bookmark table
* PDF viewer is hidden for better usability
* Generate final output PDF

### CSV Support (Multi-User Workflow)

This tool enables **collaborative annotation without access or sharing issues**:

* Users can:

  * Annotate assigned pages independently
  * Export their annotations as CSV
* Multiple users can work in parallel without conflicts
* Final consolidation:

  * Load all CSV files into the tool
  * Merge annotations seamlessly
* No dependency on shared file editing or locking mechanisms

---

## Installation

### Requirements

* Python 3.8 or higher
* Internet connection (first-time dependency install)

### Run

```bash
python annotatecrf_studio.py
```

Dependencies are automatically installed:

* PyMuPDF
* PyQt5

---

## Usage Workflow

### 1. Open PDF

* Click **Open PDF**
* Navigate using **Previous / Next**

---

### 2. Annotation Mode

* Select **Annotation**
* Click on PDF to place annotation
* Enter:

  * DOMAIN
  * VARIABLE (optional for domain-level)
  * ANNOTATION text

Optional selections:

* Domain
* Assigned Field
* Not Submitted

---

### 3. Bookmark Mode

* Select **Bookmark**
* Click on page
* Enter:

  * Bookmark text
  * Level
  * Page number

---

### 4. Multi-User Annotation (Recommended Workflow)

* Assign pages across team members
* Each user:

  * Annotates their assigned pages
  * Exports Annotation CSV
* Final step:

  * Load all CSV files into a single session
  * Combine annotations
  * Perform final review

---

### 5. Manage Entries

Use available buttons:

* Export CSV
* Load CSV
* Delete selected entry

---

### 6. Review

* Switch to **Review**
* Validate annotation and bookmark tables

---

### 7. Generate Output

* Click **Generate Final Output PDF**

Output file:

```
<input_file_name>_final.pdf
```

---

## Color Logic

Annotations are colored based on domain appearance order per page:

1st domain → Cyan
2nd domain → Green
3rd domain → Peach
4th domain → Orange
5th domain → Lavender


**Note:**
Preview colors shown in the tool are temporary. Final output can be aligned with domain-specific standards if required.

---

## CSV Structure

### Annotation CSV

| Column        | Description     |
| ------------- | --------------- |
| DOMAIN        | SDTM domain     |
| NAME          | Variable name   |
| PAGENO        | Page number     |
| ANNOTATION    | Annotation text |
| ASSIGNEDFIELD | Y or blank      |
| X1            | X coordinate    |
| Y1            | Y coordinate    |
| PAGEH         | Page height     |

---

### Bookmark CSV

| Column | Description     |
| ------ | --------------- |
| TITLE  | Bookmark text   |
| LEVEL  | Hierarchy level |
| PAGENO | Page number     |

---

## Notes

* Single-click annotation placement
* Review mode hides PDF for better clarity
* Drag-and-drop supported for annotation repositioning
* Designed to maximize usable PDF space
* Well-suited for distributed team workflows

---

## Limitations

* Annotation import requires X1, Y1, and PAGEH
* Domain colors are assigned per page (not globally fixed)
* Editing entries is not currently supported (delete and re-add)

---

## Future Enhancements

* Edit existing annotations and bookmarks
* Fixed domain-to-color mapping across entire document
* Template-based annotation loading
* Domain-specific formatting rules
* Batch processing support

---

## Support

For queries, suggestions, or issues:

[Manivannan.Mathialagan@veristat.com](mailto:Manivannan.Mathialagan@veristat.com)

---

## License

No external licensing required.
Designed for internal and organizational use (customize as needed).
