from __future__ import annotations

import base64
import html
import mimetypes
import re
from pathlib import Path


# =============================================================================
# AnnotateCRF Studio - HTML User Guide Builder
# =============================================================================
#
# Keep this .pyw file in your working folder together with:
#
#   build_acrf_user_guide.pyw
#   version_history.md
#   screenshot_1.png
#   screenshot_2.png
#   screenshot_3.png
#   ...
#
# Double-click this .pyw file to generate:
#
#   AnnotateCRF_Studio_User_Guide.html
#
# The output HTML is self-contained. Screenshots are embedded directly in it.
# =============================================================================


WORKING_DIR = Path(__file__).resolve().parent
VERSION_HISTORY_FILE = WORKING_DIR / "version_history.md"
OUTPUT_HTML = WORKING_DIR / "AnnotateCRF_Studio_User_Guide.html"

TOOL_ROOT = r"P:\Biostatistics\aCRF Tool"
PYTHON_APP_PATH = r"P:\Biostatistics\aCRF Tool\01_Python_Version\annotatecrf_studio.pyw"
PORTABLE_APP_PATH = r"P:\Biostatistics\aCRF Tool\02_Portable_Version\Annotate CRF Studio.exe"


PYTHON_FOLDER = "01_Python_Version"
PORTABLE_FOLDER = "02_Portable_Version"
DEMO_FOLDER = "03_Demo_Examples"
ARCHIVE_FOLDER = "04_Archives"

ARCHIVE_REMOVAL_TEXT = "The folder will be retained until 31 July 2026 and is planned for removal after that date."

SUPPORTED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp"}


# =============================================================================
# EDITABLE TOOL WORKFLOW
# =============================================================================
#
# This is the main section you can edit whenever you need to add, remove or
# rearrange workflow steps and screenshots.
#
# Each workflow item supports:
#
#   "title"       : Heading for the step
#   "description" : Main instruction text
#   "points"      : Optional bullet points
#   "screenshots" : Zero, one or multiple screenshots
#
# Each screenshot supports:
#
#   "file"    : Screenshot filename in the same folder as this .pyw
#   "heading" : Heading shown below the screenshot
#   "caption" : Explanation shown below the heading
#
# Example:
#
# {
#     "title": "Open Source CRF",
#     "description": "Select the source PDF.",
#     "points": ["Confirm the correct file.", "Review the page count."],
#     "screenshots": [
#         {
#             "file": "screenshot_2.png",
#             "heading": "Open CRF Window",
#             "caption": "Select the source CRF PDF and open it."
#         },
#         {
#             "file": "screenshot_3.png",
#             "heading": "Loaded CRF",
#             "caption": "The selected CRF is displayed in the application."
#         }
#     ]
# }
#
# You can use as many screenshots as needed under each workflow step.
# =============================================================================

WORKFLOW_STEPS = [
    {
        "title": "Main Window Before Opening a PDF",
        "description": "Launch AnnotateCRF Studio and review the main application window before loading a source CRF.",
        "points": [
            "The PDF viewer is empty until a source CRF is opened.",
            "Page count and page orientation are displayed only after a PDF is loaded.",
        ],
        "screenshots": [
            {
                "file": "screenshot_1.png",
                "heading": "Main Window",
                "caption": "AnnotateCRF Studio main window before opening a PDF.",
            }
        ],
    },
    {
        "title": "Open the Source CRF PDF",
        "description": "Open the required source CRF PDF and confirm that the page count and current page orientation are displayed.",
        "points": [
            "Confirm that the correct CRF version is selected.",
            "Review the total number of pages.",
            "Confirm whether the current page is Portrait or Landscape.",
            "Use the page number field and Go option to navigate directly to a required page.",
        ],
        "screenshots": [
            {
                "file": "screenshot_2.png",
                "heading": "Source CRF Loaded",
                "caption": "Loaded PDF showing available pages, current page number and page orientation.",
            }
        ],
    },
    {
        "title": "Load SDTM Metadata",
        "description": "Select the applicable SDTM Implementation Guide and load domain and variable metadata from the CDISC Library.",
        "points": [
            "Supported versions include SDTMIG 3.2, 3.3 and 3.4.",
            "A valid CDISC Library API key is required.",
            "Review the loaded metadata before creating annotations.",
        ],
        "screenshots": [
            {
                "file": "screenshot_3.png",
                "heading": "Load SDTM Metadata",
                "caption": "Select the SDTMIG version and load CDISC Library metadata.",
            }
        ],
    },
    {
        "title": "Add a Domain Annotation",
        "description": "Add a Domain annotation to identify the SDTM domain represented by a CRF section.",
        "points": [
            "Select the required domain.",
            "Use the Domain annotation option.",
            "Review the automatically generated domain annotation text.",
        ],
        "screenshots": [
            {
                "file": "screenshot_4.png",
                "heading": "Domain Annotation",
                "caption": "Create and place a Domain annotation.",
            }
        ],
    },
    {
        "title": "Add a Variable Annotation and Connector Line",
        "description": "Add an SDTM variable annotation and create a connector line when the annotation box must be linked to the corresponding CRF field.",
        "points": [
            "Select the applicable domain and variable.",
            "Place the variable annotation in the required location.",
            "Use the connector-line option where a direct visual link is needed.",
            "Connector-line coordinates are stored within the Annotation CSV.",
        ],
        "screenshots": [
            {
                "file": "screenshot_5.png",
                "heading": "Variable Annotation",
                "caption": "Create and place an SDTM variable annotation.",
            },
            {
                "file": "screenshot_6.png",
                "heading": "Connector Line",
                "caption": "Add a connector line between the annotation and the CRF field.",
            },
        ],
    },
    {
        "title": "Add an Assigned Field Annotation",
        "description": "Create an Assigned Field annotation for values assigned during SDTM mapping.",
        "points": [
            "Select the required domain and variable.",
            "Enable the Assigned field option.",
            "Assigned Field annotations are displayed with dashed-border formatting.",
        ],
        "screenshots": [
            {
                "file": "screenshot_7.png",
                "heading": "Assigned Field Annotation",
                "caption": "Create an Assigned Field annotation.",
            }
        ],
    },
    {
        "title": "Add a Not Submitted Annotation",
        "description": "Use the Not Submitted option for collected CRF content that will not be submitted in SDTM.",
        "points": [
            "Enable the Not Submitted option.",
            "Confirm that the annotation is displayed as [NOT SUBMITTED].",
        ],
        "screenshots": [
            {
                "file": "screenshot_8.png",
                "heading": "Not Submitted Annotation",
                "caption": "Create a Not Submitted annotation.",
            }
        ],
    },
    {
        "title": "Add a Refer Page Annotation",
        "description": "Create a Refer Page annotation when annotation details are provided on another CRF page.",
        "points": [
            "Enable the Refer Page option.",
            "Enter the required destination page number.",
            "The final PDF includes a clickable internal page link.",
        ],
        "screenshots": [
            {
                "file": "screenshot_9.png",
                "heading": "Refer Page Annotation",
                "caption": "Create a Refer Page annotation with an internal hyperlink.",
            }
        ],
    },
    {
        "title": "Add a SUPP Annotation",
        "description": "Create a Supplemental Qualifier annotation for a non-standard variable represented in a SUPP domain.",
        "points": [
            "Select the parent SDTM domain and required variable.",
            "Enable the SUPP annotation option.",
            "Review the automatically generated text, such as CMAESPID in SUPPCM.",
        ],
        "screenshots": [
            {
                "file": "screenshot_10.png",
                "heading": "SUPP Annotation",
                "caption": "Create a Supplemental Qualifier annotation.",
            }
        ],
    },
    {
        "title": "Review, Edit, Copy and Delete Annotations",
        "description": "Use the annotation review window to manage annotations after they have been created.",
        "points": [
            "Review all annotation records.",
            "Edit the selected annotation.",
            "Copy one or multiple annotations to another page.",
            "Delete one or multiple selected annotations.",
        ],
        "screenshots": [
            {
                "file": "screenshot_11.png",
                "heading": "Annotation Review Window",
                "caption": "Review, edit, copy and delete annotation records.",
            }
        ],
    },
    {
        "title": "Add and Organize Bookmarks",
        "description": "Create bookmarks and assign the required hierarchy level for PDF navigation.",
        "points": [
            "Create the Level 1 bookmark first.",
            "Add the related Level 2 bookmark below the applicable Level 1 bookmark.",
            "Confirm the bookmark title, hierarchy level and destination page number.",
            "Review the bookmark order before generating the final output.",
        ],
        "screenshots": [
            {
                "file": "screenshot_12.png",
                "heading": "Level 1 Bookmark",
                "caption": "Create the main Level 1 bookmark before adding any related lower-level bookmarks.",
            },
            {
                "file": "screenshot_13.png",
                "heading": "Level 2 Bookmark",
                "caption": "Add the related Level 2 bookmark below the applicable Level 1 bookmark.",
            },
        ],
    },
    {
        "title": "Review All Inputs and Create Final Output",
        "description": "Review annotations and bookmarks together before generating the final annotated PDF.",
        "points": [
            "Review Annotation CSV content.",
            "Review Bookmark CSV content.",
            "Confirm connector lines, page references and bookmark hierarchy.",
            "Confirm that Level 1 bookmarks appear before their related Level 2 bookmarks.",
            "Generate the final output only after completing the review.",
        ],
        "screenshots": [],
    },
    {
        "title": "Review the Generated Final Output",
        "description": "Open the generated annotated PDF and verify bookmarks, annotations, links and page orientation.",
        "points": [
            "Confirm that all annotations are visible and readable.",
            "Confirm that bookmarks and the Table of Contents are clickable.",
            "Confirm that Refer Page links navigate correctly.",
            "Confirm correct output on both Portrait and Landscape pages.",
        ],
        "screenshots": [
            {
                "file": "screenshot_14.png",
                "heading": "Generated Final Annotated PDF",
                "caption": "Final output showing Level 1 and Level 2 bookmarks, annotations and internal navigation.",
            }
        ],
    },
]


def inline_markdown(text: str) -> str:
    escaped = html.escape(text)
    escaped = re.sub(r"`([^`]+)`", r"<code>\1</code>", escaped)
    escaped = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", escaped)
    escaped = re.sub(r"\*([^*]+)\*", r"<em>\1</em>", escaped)
    return escaped


def markdown_to_html(markdown_text: str) -> str:
    """Simple built-in Markdown converter for version_history.md."""
    output: list[str] = []
    paragraph_lines: list[str] = []
    active_list: str | None = None

    def flush_paragraph() -> None:
        nonlocal paragraph_lines
        if paragraph_lines:
            combined = " ".join(x.strip() for x in paragraph_lines)
            output.append(f"<p>{inline_markdown(combined)}</p>")
            paragraph_lines = []

    def close_list() -> None:
        nonlocal active_list
        if active_list:
            output.append(f"</{active_list}>")
            active_list = None

    for raw_line in markdown_text.splitlines():
        line = raw_line.rstrip()

        if not line.strip():
            flush_paragraph()
            close_list()
            continue

        heading = re.match(r"^(#{1,6})\s+(.+)$", line)
        if heading:
            flush_paragraph()
            close_list()
            level = len(heading.group(1))
            output.append(f"<h{level}>{inline_markdown(heading.group(2))}</h{level}>")
            continue

        bullet = re.match(r"^\s*[-*]\s+(.+)$", line)
        if bullet:
            flush_paragraph()
            if active_list != "ul":
                close_list()
                output.append("<ul>")
                active_list = "ul"
            output.append(f"<li>{inline_markdown(bullet.group(1))}</li>")
            continue

        numbered = re.match(r"^\s*\d+\.\s+(.+)$", line)
        if numbered:
            flush_paragraph()
            if active_list != "ol":
                close_list()
                output.append("<ol>")
                active_list = "ol"
            output.append(f"<li>{inline_markdown(numbered.group(1))}</li>")
            continue

        close_list()
        paragraph_lines.append(line)

    flush_paragraph()
    close_list()
    return "\n".join(output)


def read_version_history() -> str:
    if not VERSION_HISTORY_FILE.exists():
        return f"""
        <div class="notice warning">
            <strong>Version history file not found.</strong>
            Add <code>{html.escape(VERSION_HISTORY_FILE.name)}</code> to the working folder
            and run this builder again.
        </div>
        """

    return markdown_to_html(VERSION_HISTORY_FILE.read_text(encoding="utf-8"))


def image_to_data_uri(path: Path) -> str:
    mime_type, _ = mimetypes.guess_type(path.name)
    mime_type = mime_type or "image/png"
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime_type};base64,{encoded}"


def build_screenshot_html(item: dict) -> str:
    filename = str(item.get("file", "")).strip()
    heading = str(item.get("heading", "")).strip() or filename
    caption = str(item.get("caption", "")).strip()

    if not filename:
        return ""

    image_path = WORKING_DIR / filename

    if not image_path.exists():
        return f"""
        <figure class="screenshot-card missing">
            <div class="image-placeholder">
                <div class="placeholder-icon">🖼️</div>
                <strong>Screenshot not found</strong>
                <span>{html.escape(filename)}</span>
            </div>
            <figcaption>
                <strong>{html.escape(heading)}</strong>
                <span>{html.escape(caption)}</span>
            </figcaption>
        </figure>
        """

    src = image_to_data_uri(image_path)

    return f"""
    <figure class="screenshot-card">
        <img src="{src}" alt="{html.escape(heading)}" class="guide-image">
        <figcaption>
            <strong>{html.escape(heading)}</strong>
            <span>{html.escape(caption)}</span>
        </figcaption>
    </figure>
    """


def build_workflow_html() -> str:
    blocks: list[str] = []

    for index, step in enumerate(WORKFLOW_STEPS, start=1):
        title = html.escape(str(step.get("title", f"Step {index}")))
        description = html.escape(str(step.get("description", "")))

        points = step.get("points", []) or []
        screenshots = step.get("screenshots", []) or []

        points_html = ""
        if points:
            points_html = "<ul>" + "".join(
                f"<li>{html.escape(str(point))}</li>" for point in points
            ) + "</ul>"

        screenshots_html = "".join(
            build_screenshot_html(screenshot)
            for screenshot in screenshots
        )

        blocks.append(
            f"""
            <article class="workflow-step">
                <div class="workflow-step-heading">
                    <span class="workflow-number">{index}</span>
                    <h3>{title}</h3>
                </div>
                <p>{description}</p>
                {points_html}
                {screenshots_html}
            </article>
            """
        )

    return "\n".join(blocks)


def build_html() -> str:
    version_history_html = read_version_history()
    workflow_html = build_workflow_html()

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<title>AnnotateCRF Studio - User Guide</title>

<style>
    :root {{
        --navy: #12385b;
        --blue: #246ca8;
        --teal: #167786;
        --green: #347c4c;
        --orange: #bb6a1b;
        --purple: #67509b;
        --background: #f4f7fa;
        --text: #1f2933;
        --muted: #607080;
        --border: #d7e1e9;
        --light-blue: #eaf4ff;
        --light-green: #eaf7ee;
        --light-orange: #fff2e3;
        --light-purple: #f2edff;
        --light-red: #fff0f0;
        --shadow: 0 10px 28px rgba(18, 56, 91, 0.11);
    }}

    * {{ box-sizing: border-box; }}

    html {{ scroll-behavior: smooth; }}

    body {{
        margin: 0;
        background: var(--background);
        color: var(--text);
        font-family: "Segoe UI", Arial, sans-serif;
        line-height: 1.65;
    }}

    .hero {{
        padding: 58px 28px;
        color: white;
        background: linear-gradient(135deg, rgba(18,56,91,.98), rgba(36,108,168,.94));
        border-bottom: 6px solid #66c4db;
    }}

    .hero-inner {{
        max-width: 1180px;
        margin: 0 auto;
    }}

    .hero-label {{
        display: inline-block;
        padding: 6px 12px;
        border: 1px solid rgba(255,255,255,.45);
        border-radius: 999px;
        background: rgba(255,255,255,.12);
        font-size: 13px;
        font-weight: 700;
        letter-spacing: .5px;
        text-transform: uppercase;
    }}

    .hero h1 {{
        margin: 16px 0 6px;
        font-size: clamp(34px, 5vw, 58px);
        line-height: 1.1;
    }}

    .hero p {{
        max-width: 850px;
        margin: 10px 0 0;
        color: #eaf6ff;
        font-size: 19px;
    }}

    .layout {{
        max-width: 1280px;
        margin: 0 auto;
        display: grid;
        grid-template-columns: 275px minmax(0, 1fr);
        gap: 26px;
        padding: 28px;
    }}

    .toc {{
        position: sticky;
        top: 18px;
        align-self: start;
        max-height: calc(100vh - 36px);
        overflow-y: auto;
        padding: 18px;
        border: 1px solid var(--border);
        border-radius: 16px;
        background: white;
        box-shadow: var(--shadow);
    }}

    .toc h2 {{
        margin: 0 0 12px;
        color: var(--navy);
        font-size: 18px;
    }}

    .toc a {{
        display: block;
        margin: 2px 0;
        padding: 8px 10px;
        color: #2b4b65;
        text-decoration: none;
        border-radius: 8px;
        font-size: 14px;
    }}

    .toc a:hover {{
        color: var(--blue);
        background: var(--light-blue);
    }}

    .toc a.toc-sub {{
        margin-left: 12px;
        padding-left: 14px;
        font-size: 13px;
        border-left: 2px solid var(--border);
    }}

    main {{ min-width: 0; }}

    .section {{
        margin-bottom: 24px;
        overflow: hidden;
        border: 1px solid var(--border);
        border-radius: 18px;
        background: white;
        box-shadow: var(--shadow);
    }}

    .section-header {{
        padding: 20px 26px;
        color: white;
        background: linear-gradient(135deg, var(--navy), var(--blue));
    }}

    .section-header.green {{
        background: linear-gradient(135deg, #235f39, var(--green));
    }}

    .section-header.orange {{
        background: linear-gradient(135deg, #8b4e13, var(--orange));
    }}

    .section-header.purple {{
        background: linear-gradient(135deg, #49366f, var(--purple));
    }}

    .section-header.teal {{
        background: linear-gradient(135deg, #145b65, var(--teal));
    }}

    .section-header h2 {{
        margin: 0;
        font-size: 25px;
    }}

    .section-body {{ padding: 26px; }}

    .path-box {{
        margin: 14px 0;
        padding: 14px 16px;
        overflow-wrap: anywhere;
        border-radius: 10px;
        background: #102f4a;
        color: #f5f9ff;
        font-family: Consolas, monospace;
        font-size: 14px;
    }}

    .folder-subsection {{
        margin: 24px 0;
        padding: 20px;
        border: 1px solid var(--border);
        border-radius: 14px;
        background: #fbfdff;
    }}

    .subsection-title {{
        display: flex;
        align-items: center;
        gap: 12px;
    }}

    .subsection-title h3 {{ margin: 0; }}

    .subsection-icon {{
        width: 42px;
        height: 42px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        border-radius: 12px;
        background: var(--light-purple);
        font-size: 22px;
        flex: 0 0 auto;
    }}

    .folder-name {{
        margin-top: 3px;
        color: var(--muted);
        font-size: 13px;
    }}

    .recommended-text {{
        color: var(--green);
        font-size: 14px;
        font-weight: 700;
    }}

    .path-label {{
        margin-top: 14px;
        margin-bottom: 6px;
        color: var(--navy);
        font-size: 13px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: .4px;
    }}

    h3 {{
        margin-top: 22px;
        margin-bottom: 10px;
        color: var(--navy);
    }}

    code {{
        padding: 2px 6px;
        border: 1px solid #d9e3eb;
        border-radius: 5px;
        background: #eef3f7;
        font-family: Consolas, monospace;
    }}

    .folder-grid,
    .feature-grid {{
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 18px;
        margin: 18px 0;
    }}

    .folder-card,
    .feature-card {{
        padding: 18px;
        border: 1px solid var(--border);
        border-radius: 14px;
        background: #fbfdff;
    }}

    .folder-card.python {{
        border-left: 6px solid var(--green);
        background: var(--light-green);
    }}

    .folder-card.portable {{
        border-left: 6px solid var(--orange);
        background: var(--light-orange);
    }}

    .folder-card.demo {{
        border-left: 6px solid var(--blue);
        background: var(--light-blue);
    }}

    .folder-card.archive {{
        border-left: 6px solid var(--purple);
        background: var(--light-purple);
    }}

    .label {{
        display: inline-block;
        margin-bottom: 8px;
        padding: 4px 9px;
        border-radius: 999px;
        background: rgba(18,56,91,.10);
        color: var(--navy);
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
    }}

    .notice {{
        margin: 18px 0;
        padding: 16px 18px;
        border-left: 6px solid;
        border-radius: 12px;
    }}

    .notice.info {{
        border-color: #3f8bc7;
        background: var(--light-blue);
    }}

    .notice.success {{
        border-color: #3c9257;
        background: var(--light-green);
    }}

    .notice.warning {{
        border-color: #d48227;
        background: var(--light-orange);
    }}

    .notice.danger {{
        border-color: #c94f4f;
        background: var(--light-red);
    }}

    .feature-group {{
        margin: 0 0 18px;
        padding: 18px 20px;
        border: 1px solid var(--border);
        border-left: 6px solid var(--blue);
        border-radius: 12px;
        background: #fbfdff;
    }}

    .feature-group:nth-child(2) {{ border-left-color: var(--teal); }}
    .feature-group:nth-child(3) {{ border-left-color: var(--orange); }}
    .feature-group:nth-child(4) {{ border-left-color: var(--purple); }}
    .feature-group:nth-child(5) {{ border-left-color: var(--green); }}
    .feature-group:nth-child(6) {{ border-left-color: var(--navy); }}

    .feature-group h3 {{ margin-top: 0; }}

    .workflow-step {{
        margin-bottom: 24px;
        padding: 20px;
        border: 1px solid var(--border);
        border-radius: 14px;
        background: #fbfdff;
    }}

    .workflow-step-heading {{
        display: flex;
        align-items: center;
        gap: 12px;
    }}

    .workflow-step-heading h3 {{
        margin: 0;
    }}

    .workflow-number {{
        min-width: 38px;
        height: 38px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        border-radius: 50%;
        background: var(--blue);
        color: white;
        font-weight: 800;
    }}

    .screenshot-card {{
        margin: 20px 0 4px;
        overflow: hidden;
        border: 1px solid var(--border);
        border-radius: 14px;
        background: white;
    }}

    .guide-image {{
        display: block;
        width: 100%;
        height: auto;
        background: #eef3f7;
    }}

    .image-placeholder {{
        min-height: 230px;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        gap: 8px;
        background: #f3f6f9;
        color: var(--muted);
    }}

    .placeholder-icon {{ font-size: 40px; }}

    figcaption {{
        display: flex;
        flex-direction: column;
        gap: 4px;
        padding: 14px 16px;
        background: #f8fbfd;
        color: var(--muted);
    }}

    figcaption strong {{ color: var(--navy); }}

    .version-history h1 {{ display: none; }}

    .footer {{
        padding: 8px 24px 38px;
        text-align: center;
        color: var(--muted);
        font-size: 13px;
    }}

    @media (max-width: 900px) {{
        .layout {{
            grid-template-columns: 1fr;
            padding: 16px;
        }}

        .toc {{
            position: relative;
            top: 0;
            max-height: none;
        }}

        .folder-grid,
        .feature-grid {{
            grid-template-columns: 1fr;
        }}
    }}

    @media print {{
        body {{ background: white; }}
        .layout {{ display: block; padding: 0; }}
        .toc {{ display: none; }}
        .section {{ box-shadow: none; break-inside: avoid; }}
    }}
</style>
</head>

<body>

<header class="hero">
    <div class="hero-inner">
        <span class="hero-label">CDISC MSG 2.0 aCRF Utility</span>
        <h1>AnnotateCRF Studio</h1>
        <p>User Guide and Version History</p>
    </div>
</header>

<div class="layout">

<nav class="toc" aria-label="Table of contents">
    <h2>Contents</h2>
    <a href="#overview">Overview</a>
    <a href="#folder-structure">Application Folder Structure</a>
    <a class="toc-sub" href="#python-version">Python Version</a>
    <a class="toc-sub" href="#portable-version">Portable Version</a>
    <a class="toc-sub" href="#demo-examples">Demo Examples</a>
    <a class="toc-sub" href="#archives">Archives</a>
    <a href="#workflow">Using AnnotateCRF Studio</a>
    <a href="#features">Application Features</a>
    <a href="#version-history">Version History</a>
</nav>

<main>

<section class="section" id="overview">
    <div class="section-header">
        <h2>Overview</h2>
    </div>

    <div class="section-body">
        <p>
            <strong>AnnotateCRF Studio</strong> is a graphical application used to create,
            review and publish submission-ready annotated CRFs. It supports CDISC MSG 2.0
            annotation conventions without requiring Adobe Acrobat or another
            subscription-based PDF annotation tool.
        </p>

        <p>
            The application supports interactive annotation placement, SDTM metadata
            integration, annotation and bookmark management, internal page references,
            mixed Portrait and Landscape pages, and final annotated PDF generation.
        </p>
    </div>
</section>

<section class="section" id="folder-structure">
    <div class="section-header purple">
        <h2>Application Folder Structure</h2>
    </div>

    <div class="section-body">
        <p>
            The shared aCRF Tool folder contains the Python application, Portable
            application, demonstration examples and temporary archive files.
        </p>

        <div class="folder-subsection" id="python-version">
            <div class="subsection-title">
                <span class="subsection-icon">🐍</span>
                <div>
                    <h3>Python Version <span class="recommended-text">(Recommended)</span></h3>
                    <div class="folder-name">Folder: <code>01_Python_Version</code></div>
                </div>
            </div>
            <p>This is the recommended version for users who can access the SAS server. Double-click the following file to open AnnotateCRF Studio:</p>
            <div class="path-label">Launch Application</div>
            <div class="path-box">{html.escape(PYTHON_APP_PATH)}</div>
            <ul>
                <li>Primary version used for routine application use.</li>
                <li>Launches the GUI without displaying a command window.</li>
                <li>Used for development, testing and validation.</li>
            </ul>
        </div>

        <div class="folder-subsection" id="portable-version">
            <div class="subsection-title">
                <span class="subsection-icon">💼</span>
                <div>
                    <h3>Portable Version</h3>
                    <div class="folder-name">Folder: <code>02_Portable_Version</code></div>
                </div>
            </div>
            <p>Use the Portable version only when accessing the tool through the SAS server is not possible. Double-click the following file:</p>
            <div class="path-label">Launch Application</div>
            <div class="path-box">{html.escape(PORTABLE_APP_PATH)}</div>
            <ul>
                <li>Standalone executable version.</li>
                <li>A separate Python installation is not required.</li>
                <li>Built from the same source code as the Python version.</li>
                <li>Routine behaviour testing is performed using the Python version.</li>
            </ul>
            <div class="notice warning">Where SAS server access is available, use <code>01_Python_Version</code>.</div>
        </div>

        <div class="folder-subsection" id="demo-examples">
            <div class="subsection-title">
                <span class="subsection-icon">📂</span>
                <div>
                    <h3>Demo Examples</h3>
                    <div class="folder-name">Folder: <code>03_Demo_Examples</code></div>
                </div>
            </div>
            <p>This folder contains similar end-to-end examples created using dummy CRFs for demonstration, training and application familiarisation.</p>
            <p>The available demo folders are:</p>
            <ul>
                <li><code>Demo_01</code></li>
                <li><code>Demo_02</code></li>
                <li><code>Demo_03</code></li>
            </ul>
            <p>Each demo folder may include:</p>
            <ul>
                <li>Source dummy CRF PDF</li>
                <li>Annotation CSV, including connector-line information</li>
                <li>Bookmark CSV</li>
                <li>Final annotated PDF</li>
            </ul>
            <div class="notice warning">Demo files contain dummy data only and must not be used for production studies.</div>
        </div>

        <div class="folder-subsection" id="archives">
            <div class="subsection-title">
                <span class="subsection-icon">🗃️</span>
                <div>
                    <h3>Archives</h3>
                    <div class="folder-name">Folder: <code>04_Archives</code></div>
                </div>
            </div>
            <p>This folder contains older application files and files previously available directly in the shared aCRF Tool folder.</p>
            <ul>
                <li>Retained temporarily to support the transition to the new structure.</li>
                <li>Provided for reference only.</li>
                <li>Routine application access should use <code>01_Python_Version</code> or <code>02_Portable_Version</code>.</li>
            </ul>
            <div class="notice danger">{html.escape(ARCHIVE_REMOVAL_TEXT)}</div>
        </div>
    </div>
</section>

<section class="section" id="workflow">
    <div class="section-header teal">
        <h2>Using AnnotateCRF Studio</h2>
    </div>

    <div class="section-body">
        <p>
            The steps below describe the standard workflow for creating and reviewing
            an annotated CRF.
        </p>

        {workflow_html}
    </div>
</section>

<section class="section" id="features">
    <div class="section-header green">
        <h2>Application Features</h2>
    </div>

    <div class="section-body">
        <div class="feature-group">
            <h3>PDF Viewing and Navigation</h3>
            <ul>
                <li>Open source CRF PDFs for annotation.</li>
                <li>Navigate directly to any page using the page number field and <strong>Go</strong> option.</li>
                <li>Display the current page number and total number of pages.</li>
                <li>Display the current page orientation as Portrait or Landscape.</li>
                <li>Support PDFs containing both Portrait and Landscape pages.</li>
            </ul>
        </div>
        <div class="feature-group">
            <h3>Annotation Creation</h3>
            <ul>
                <li>Create Domain annotations.</li>
                <li>Create SDTM Variable annotations.</li>
                <li>Create Assigned Field annotations.</li>
                <li>Create Not Submitted annotations.</li>
                <li>Create Refer Page annotations with clickable internal hyperlinks.</li>
                <li>Create SUPP annotations such as <code>CMAESPID in SUPPCM</code>.</li>
                <li>Create and manage connector lines.</li>
            </ul>
        </div>
        <div class="feature-group">
            <h3>Annotation Review and Management</h3>
            <ul>
                <li>Review all annotations in a single grid.</li>
                <li>Edit existing annotations.</li>
                <li>Copy annotations to another page.</li>
                <li>Select and copy multiple annotations together.</li>
                <li>Select and delete multiple annotations together.</li>
                <li>Resize and reposition annotations where required.</li>
            </ul>
        </div>
        <div class="feature-group">
            <h3>CDISC Library Integration</h3>
            <ul>
                <li>Load SDTM domains and variables from the CDISC Library.</li>
                <li>Support SDTMIG 3.2, SDTMIG 3.3 and SDTMIG 3.4.</li>
                <li>Reduce manual maintenance of domain and variable lists.</li>
                <li>Allow manual domain and variable entry when required.</li>
            </ul>
        </div>
        <div class="feature-group">
            <h3>Bookmarks and Table of Contents</h3>
            <ul>
                <li>Create, edit, copy and delete bookmarks.</li>
                <li>Assign and maintain bookmark hierarchy levels.</li>
                <li>Import and export bookmark CSV files.</li>
                <li>Generate a clickable Table of Contents.</li>
                <li>Create internal PDF navigation links.</li>
            </ul>
        </div>
        <div class="feature-group">
            <h3>Import, Export and Final Output</h3>
            <ul>
                <li>Import and export Annotation CSV files.</li>
                <li>Store connector-line information within the Annotation CSV.</li>
                <li>Import and export Bookmark CSV files.</li>
                <li>Review annotations and bookmarks before final generation.</li>
                <li>Generate a submission-ready annotated PDF.</li>
                <li>Preserve bookmarks, annotations, hyperlinks and mixed page orientation.</li>
            </ul>
        </div>
    </div>
</section>

<section class="section version-history" id="version-history">
    <div class="section-header">
        <h2>Version History</h2>
    </div>

    <div class="section-body">
        {version_history_html}
    </div>
</section>

</main>
</div>

<footer class="footer">
    AnnotateCRF Studio · User Guide and Version History
</footer>

</body>
</html>
"""


def show_message(title: str, message: str, error: bool = False) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        if error:
            messagebox.showerror(title, message, parent=root)
        else:
            messagebox.showinfo(title, message, parent=root)

        root.destroy()
    except Exception:
        pass


def run_builder() -> None:
    try:
        content = build_html()
        OUTPUT_HTML.write_text(content, encoding="utf-8")

        configured_images = []
        missing_images = []

        for step in WORKFLOW_STEPS:
            for screenshot in step.get("screenshots", []) or []:
                filename = str(screenshot.get("file", "")).strip()
                if not filename:
                    continue
                configured_images.append(filename)
                if not (WORKING_DIR / filename).exists():
                    missing_images.append(filename)

        message = [
            "AnnotateCRF Studio user guide created successfully.",
            "",
            "Output:",
            str(OUTPUT_HTML),
            "",
            f"Configured workflow screenshots: {len(configured_images)}",
            f"Missing screenshots: {len(missing_images)}",
        ]

        if missing_images:
            message.extend([
                "",
                "Missing screenshot placeholders were added for:",
                "\n".join(missing_images),
            ])

        if not VERSION_HISTORY_FILE.exists():
            message.extend([
                "",
                f"Warning: {VERSION_HISTORY_FILE.name} was not found.",
                "A placeholder was added to the HTML.",
            ])

        show_message(
            "AnnotateCRF Studio User Guide",
            "\n".join(message),
            error=False,
        )

    except Exception as exc:
        show_message(
            "AnnotateCRF Studio User Guide - Error",
            "The user guide could not be generated.\n\n"
            f"Error details:\n{exc}",
            error=True,
        )


if __name__ == "__main__":
    run_builder()
