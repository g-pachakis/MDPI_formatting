# MDPI Manuscript Formatter

Automatically reformats a Word manuscript (`.docx`) into the [MDPI journal template](https://www.mdpi.com/authors/layout) format — correct styles, three-line tables, numbered references, and all.

## Quick Start

**1. Install**

```bash
pip install python-docx
```

**2. Download this repo** and place `mdpi_template.docx` next to `mdpi_formatter.py`:

```
your-folder/
  mdpi_formatter.py
  mdpi_template.docx      ← get this from MDPI (included in the repo)
```

**3. Run**

```bash
python mdpi_formatter.py
```

A file picker window will open. Select your manuscript `.docx` and the formatted output will be saved in the same directory as:

```
your-manuscript_MDPI.docx
```

You can also pass the file path directly:

```bash
python mdpi_formatter.py path/to/manuscript.docx
```

## What It Does

The script reads your manuscript with `python-docx`, maps each element to the correct MDPI style, and writes the result into a copy of the official MDPI template — preserving headers, footers, line numbering, and page layout.

| Manuscript Element | MDPI Style Applied |
|---|---|
| Title | `MDPI_1.2_title` (18pt bold) |
| Authors | `MDPI_1.3_authornames` (placeholder) |
| Abstract | `MDPI_1.7_abstract` (indented, "Abstract:" bold prefix) |
| Keywords | `MDPI_1.8_keywords` (indented, "Keywords:" bold prefix) |
| Section headings (H1) | `MDPI_2.1_heading1` (12pt bold) |
| Subsection headings (H2) | `MDPI_2.2_heading2` (italic) |
| Body text (first after heading) | `MDPI_3.2_text_no_indent` |
| Body text (subsequent) | `MDPI_3.1_text` (first-line indent) |
| Table captions | `MDPI_4.1_table_caption` (9pt, "Table N." bold) |
| Table body | `MDPI_4.2_table_body` (centered, three-line borders) |
| Table footnotes | `MDPI_4.3_table_footer` (9pt) |
| Figure captions | `MDPI_5.1_figure_caption` (9pt) |
| References | `MDPI_8.1_references` (9pt, numbered list) |

All inline formatting (bold, italic, subscript, superscript) is preserved from the original.

## Requirements

- **Python** 3.8+
- **python-docx** (`pip install python-docx`)
- **tkinter** (included with Python on Windows and macOS; on Linux: `sudo apt install python3-tk`)

No other external tools needed — no pandoc, no LaTeX, no LibreOffice.

## Notes

- **Author names and affiliations** are inserted as placeholders — fill them in manually after formatting.
- **Figures** are preserved as caption-only placeholders. Insert your actual figures in the output document.
- **Tables** are rebuilt as MDPI three-line tables (thick top/bottom borders, thin header separator).
- The template file (`mdpi_template.docx`) must be in the same folder as the script. Get the latest version from your target MDPI journal's submission page.

## How It Works

1. Reads the manuscript `.docx` using `python-docx` (no pandoc needed)
2. Walks the document body in order — paragraphs and tables — classifying each element by its style and content
3. Extracts the MDPI template's `document.xml` from the `.docx` zip archive
4. Builds a new `document.xml` with the manuscript content wrapped in MDPI styles
5. Writes the result back into a copy of the template zip, preserving all headers, footers, fonts, and page settings

## License

MIT
