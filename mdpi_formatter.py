#!/usr/bin/env python3
"""
MDPI Manuscript Formatter
=========================
Reformats a manuscript .docx into MDPI journal template format.

Place mdpi_template.docx in the same folder as this script, then run:
    python mdpi_formatter.py

A file picker will open. Select your manuscript and the formatted
output will be saved next to it as <filename>_MDPI.docx.
"""

import sys
import os
import re
import zipfile
import shutil
import tempfile
import xml.etree.ElementTree as ET
from copy import deepcopy

try:
    from docx import Document as DocxDocument
    from docx.oxml.ns import qn as docx_qn
except ImportError:
    print("ERROR: python-docx is required. Install it with:")
    print("  pip install python-docx")
    sys.exit(1)


# ── XML Namespace Setup ─────────────────────────────────────────────────────
# Register ALL OOXML namespaces so ElementTree preserves them in output

NSMAP = {
    'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    'cx': 'http://schemas.microsoft.com/office/drawing/2014/chartex',
    'cx1': 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex',
    'cx2': 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex',
    'cx3': 'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex',
    'cx4': 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex',
    'cx5': 'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex',
    'cx6': 'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex',
    'cx7': 'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex',
    'cx8': 'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'aink': 'http://schemas.microsoft.com/office/drawing/2016/ink',
    'am3d': 'http://schemas.microsoft.com/office/drawing/2017/model3d',
    'o': 'urn:schemas-microsoft-com:office:office',
    'oel': 'http://schemas.microsoft.com/office/2019/extlst',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v': 'urn:schemas-microsoft-com:vml',
    'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'w16cex': 'http://schemas.microsoft.com/office/word/2018/wordml/cex',
    'w16cid': 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
    'w16': 'http://schemas.microsoft.com/office/word/2018/wordml',
    'w16du': 'http://schemas.microsoft.com/office/word/2023/wordml/word16du',
    'w16sdtdh': 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',
    'w16sdtfl': 'http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock',
    'w16se': 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
    'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'wpi': 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
}
for prefix, uri in NSMAP.items():
    ET.register_namespace(prefix, uri)

W = NSMAP['w']


def qn(tag):
    """Qualified name: qn('w:p') -> '{http://...}p'"""
    prefix, local = tag.split(':')
    return f'{{{NSMAP[prefix]}}}{local}'


# ── XML Paragraph/Run Builders ──────────────────────────────────────────────

def _make_t(parent, text):
    """Create <w:t> with space preservation."""
    t = ET.SubElement(parent, qn('w:t'))
    t.text = text or ''
    if text and (text[0] == ' ' or text[-1] == ' '):
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return t


def _add_rpr(r_el, props):
    """Add run properties to a <w:r> element. props is list of: 'b','i','sub','sup'."""
    if not props:
        return
    rpr = ET.SubElement(r_el, qn('w:rPr'))
    for p in props:
        if p in ('b', 'bi'):
            ET.SubElement(rpr, qn('w:b'))
        if p in ('i', 'bi'):
            ET.SubElement(rpr, qn('w:i'))
        if p == 'sub':
            va = ET.SubElement(rpr, qn('w:vertAlign'))
            va.set(qn('w:val'), 'subscript')
        if p == 'sup':
            va = ET.SubElement(rpr, qn('w:vertAlign'))
            va.set(qn('w:val'), 'superscript')


def make_para(style_id, text='', bold_prefix=None):
    """Create a <w:p> with MDPI style and plain text content."""
    p = ET.Element(qn('w:p'))
    ppr = ET.SubElement(p, qn('w:pPr'))
    ET.SubElement(ppr, qn('w:pStyle')).set(qn('w:val'), style_id)
    if bold_prefix:
        r = ET.SubElement(p, qn('w:r'))
        _add_rpr(r, ['b'])
        _make_t(r, bold_prefix)
    if text:
        r = ET.SubElement(p, qn('w:r'))
        _make_t(r, text)
    return p


def make_styled_para(style_id, bold=False, italic=False, text=''):
    """Create a paragraph with a single styled run."""
    p = ET.Element(qn('w:p'))
    ppr = ET.SubElement(p, qn('w:pPr'))
    ET.SubElement(ppr, qn('w:pStyle')).set(qn('w:val'), style_id)
    r = ET.SubElement(p, qn('w:r'))
    props = []
    if bold:
        props.append('b')
    if italic:
        props.append('i')
    if props:
        _add_rpr(r, props)
    _make_t(r, text)
    return p


def copy_runs_to_para(style_id, source_para_element, bold_prefix=None):
    """Create an MDPI-styled paragraph, copying all runs from a source <w:p> element.
    Handles both ElementTree and lxml (python-docx) source elements."""
    p = ET.Element(qn('w:p'))
    ppr = ET.SubElement(p, qn('w:pPr'))
    ET.SubElement(ppr, qn('w:pStyle')).set(qn('w:val'), style_id)

    if bold_prefix:
        r = ET.SubElement(p, qn('w:r'))
        _add_rpr(r, ['b'])
        _make_t(r, bold_prefix)

    # Source may be lxml (python-docx) or ElementTree — serialize to bridge
    try:
        from lxml import etree as lxml_etree
        is_lxml = isinstance(source_para_element, lxml_etree._Element)
    except ImportError:
        is_lxml = False

    if is_lxml:
        for src_run in source_para_element.findall(docx_qn('w:r')):
            run_xml = lxml_etree.tostring(src_run, encoding='unicode')
            new_r = ET.fromstring(run_xml)
            p.append(new_r)
    else:
        for src_run in source_para_element.findall(qn('w:r')):
            new_r = deepcopy(src_run)
            p.append(new_r)

    return p


# ── MDPI Table Builder ──────────────────────────────────────────────────────

def build_mdpi_table(src_table_element, table_width=7857, indent=2608):
    """Convert a source <w:tbl> element into an MDPI three-line table.
    Reads cell text from source, builds new table with MDPI styling."""
    ns = {'w': W}
    src_rows = src_table_element.findall(docx_qn('w:tr'))
    if not src_rows:
        return None

    # Extract cell texts
    all_rows = []
    for tr in src_rows:
        cells = []
        for tc in tr.findall(docx_qn('w:tc')):
            text_parts = []
            for t_el in tc.iter(docx_qn('w:t')):
                text_parts.append(t_el.text or '')
            cells.append(''.join(text_parts).strip())
        all_rows.append(cells)

    if not all_rows:
        return None

    headers = all_rows[0]
    data_rows = all_rows[1:]
    ncols = len(headers)
    if ncols == 0:
        return None

    col_w = table_width // ncols
    remainder = table_width - (col_w * ncols)

    tbl = ET.Element(qn('w:tbl'))
    tpr = ET.SubElement(tbl, qn('w:tblPr'))
    tw = ET.SubElement(tpr, qn('w:tblW'))
    tw.set(qn('w:w'), str(table_width))
    tw.set(qn('w:type'), 'dxa')
    ti = ET.SubElement(tpr, qn('w:tblInd'))
    ti.set(qn('w:w'), str(indent))
    ti.set(qn('w:type'), 'dxa')

    # Three-line borders: thick top, thick bottom
    brd = ET.SubElement(tpr, qn('w:tblBorders'))
    for side in ['top', 'bottom']:
        b = ET.SubElement(brd, qn(f'w:{side}'))
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '8')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), 'auto')

    tl = ET.SubElement(tpr, qn('w:tblLayout'))
    tl.set(qn('w:type'), 'fixed')
    tcm = ET.SubElement(tpr, qn('w:tblCellMar'))
    for side in ['left', 'right']:
        s = ET.SubElement(tcm, qn(f'w:{side}'))
        s.set(qn('w:w'), '0')
        s.set(qn('w:type'), 'dxa')

    grid = ET.SubElement(tbl, qn('w:tblGrid'))
    for i in range(ncols):
        gc = ET.SubElement(grid, qn('w:gridCol'))
        gc.set(qn('w:w'), str(col_w + (1 if i < remainder else 0)))

    def add_row(cells, is_header=False):
        cells = list(cells) + [''] * max(0, ncols - len(cells))
        cells = cells[:ncols]
        tr = ET.SubElement(tbl, qn('w:tr'))
        for j, cell_text in enumerate(cells):
            tc = ET.SubElement(tr, qn('w:tc'))
            tcp = ET.SubElement(tc, qn('w:tcPr'))
            tcw = ET.SubElement(tcp, qn('w:tcW'))
            tcw.set(qn('w:w'), str(col_w + (1 if j < remainder else 0)))
            tcw.set(qn('w:type'), 'dxa')
            if is_header:
                tcb = ET.SubElement(tcp, qn('w:tcBorders'))
                bb = ET.SubElement(tcb, qn('w:bottom'))
                bb.set(qn('w:val'), 'single')
                bb.set(qn('w:sz'), '4')
                bb.set(qn('w:space'), '0')
                bb.set(qn('w:color'), 'auto')
            va = ET.SubElement(tcp, qn('w:vAlign'))
            va.set(qn('w:val'), 'center')
            cp = ET.SubElement(tc, qn('w:p'))
            cppr = ET.SubElement(cp, qn('w:pPr'))
            ET.SubElement(cppr, qn('w:pStyle')).set(qn('w:val'), 'MDPI42tablebody')
            sp = ET.SubElement(cppr, qn('w:spacing'))
            sp.set(qn('w:line'), '240')
            sp.set(qn('w:lineRule'), 'auto')
            cr = ET.SubElement(cp, qn('w:r'))
            if is_header:
                _add_rpr(cr, ['b'])
            _make_t(cr, cell_text)

    add_row(headers, is_header=True)
    for row in data_rows:
        add_row(row)
    return tbl


# ── Manuscript Reader ───────────────────────────────────────────────────────

def get_para_text(p_element):
    """Get full text from a <w:p> element (works with both lxml and ElementTree)."""
    texts = []
    # Use docx_qn for lxml source elements
    try:
        for t in p_element.iter(docx_qn('w:t')):
            texts.append(t.text or '')
    except Exception:
        for t in p_element.iter(qn('w:t')):
            texts.append(t.text or '')
    return ''.join(texts)


def get_para_style(p_element):
    """Get the style ID from a <w:p> element (works with both lxml and ElementTree)."""
    try:
        ppr = p_element.find(docx_qn('w:pPr'))
        if ppr is not None:
            ps = ppr.find(docx_qn('w:pStyle'))
            if ps is not None:
                return ps.get(docx_qn('w:val'), '')
    except Exception:
        ppr = p_element.find(qn('w:pPr'))
        if ppr is not None:
            ps = ppr.find(qn('w:pStyle'))
            if ps is not None:
                return ps.get(qn('w:val'), '')
    return ''


def classify_paragraph(p_element, is_after_refs_heading):
    """Classify a manuscript paragraph into an MDPI content type.
    Returns (type_str, extra_info_dict)."""
    text = get_para_text(p_element).strip()
    style = get_para_style(p_element)

    if not text:
        return 'blank', {}

    # Headings
    if style in ('Heading1', 'Heading 1'):
        if text.lower() == 'abstract':
            return 'abstract_heading', {}
        elif 'references' in text.lower():
            return 'references_heading', {'text': text}
        else:
            return 'heading1', {'text': text}

    if style in ('Heading2', 'Heading 2'):
        return 'heading2', {'text': text}

    if style in ('Heading3', 'Heading 3'):
        return 'heading3', {'text': text}

    # Bibliography / reference entries
    if style in ('Bibliography',):
        return 'reference', {}

    # Keywords
    if text.startswith('Keywords:') or text.startswith('Keywords :'):
        return 'keywords', {'text': re.sub(r'^Keywords\s*:\s*', '', text)}

    # Table caption: starts with "Table N." (with period)
    if re.match(r'^Table\s+\d+\.', text):
        return 'table_caption', {'text': text}

    # Table footer: starts with * (footnote marker)
    if text.startswith('*') and len(text) < 300:
        return 'table_footer', {'text': text}

    # Figure placeholder: contains [Figure N]
    if re.match(r'^\[?Figure\s+\d+', text):
        return 'figure_placeholder', {}

    # References section: lines starting with [N] after the References heading
    if is_after_refs_heading and re.match(r'^\[\d+\]', text):
        return 'reference', {}

    # Default: body paragraph
    return 'paragraph', {}


def read_manuscript(manuscript_path):
    """Read the manuscript and return a list of content items in document order.
    Each item is a dict with 'type' and 'element' (the source XML element)."""
    doc = DocxDocument(manuscript_path)
    body = doc.element.body
    items = []
    is_after_refs = False
    is_after_heading = False  # Track abstract specially
    in_abstract = False

    for child in body:
        tag = child.tag.split('}')[-1]

        if tag == 'p':
            ctype, info = classify_paragraph(child, is_after_refs)

            if ctype == 'abstract_heading':
                in_abstract = True
                continue  # Skip the heading itself; we'll prefix "Abstract:"

            if ctype == 'references_heading':
                is_after_refs = True
                items.append(
                    {'type': 'heading1', 'element': child, 'text': info['text']})
                continue

            if in_abstract and ctype == 'paragraph':
                items.append({'type': 'abstract', 'element': child})
                in_abstract = False
                continue

            if ctype == 'keywords':
                in_abstract = False

            if ctype == 'blank':
                continue

            items.append({'type': ctype, 'element': child, **info})

        elif tag == 'tbl':
            items.append({'type': 'table', 'element': child})

        elif tag == 'sectPr':
            items.append({'type': 'sectPr', 'element': child})

    return items


# ── Document Builder ────────────────────────────────────────────────────────

def build_document(template_path, items, output_path):
    """Build the output MDPI document by:
    1. Extracting the template's document.xml
    2. Building new body content with MDPI styles
    3. Replacing document.xml in the template zip
    """

    # Parse template's document.xml to get root element with all namespace declarations
    with zipfile.ZipFile(template_path, 'r') as z:
        doc_xml = z.read('word/document.xml')

    tree = ET.ElementTree(ET.fromstring(doc_xml))
    root = tree.getroot()
    body = root.find(qn('w:body'))

    # Save sectPr from template
    sect_pr = body.find(qn('w:sectPr'))
    if sect_pr is not None:
        sect_pr = deepcopy(sect_pr)
        # Remove bidi (right-to-left) — manuscript is English
        bidi = sect_pr.find(qn('w:bidi'))
        if bidi is not None:
            sect_pr.remove(bidi)

    # Clear template body
    for child in list(body):
        body.remove(child)

    # ── Build new body ──
    # Article type
    body.append(make_styled_para('MDPI11articletype', italic=True, text='Article'))

    after_heading = False
    title_added = False

    for item in items:
        itype = item['type']
        el = item.get('element')

        if itype == 'sectPr':
            continue  # We add our own at the end

        # Title: first paragraph (before any heading)
        if not title_added and itype == 'paragraph':
            # Probably the title if we haven't seen a heading yet
            text = get_para_text(el).strip()
            if text:
                body.append(make_styled_para('MDPI12title', bold=True, text=text))
                # Author placeholder
                body.append(_build_author_placeholder())
                title_added = True
                after_heading = False
                continue

        if not title_added and itype in ('abstract', 'keywords', 'heading1'):
            # If title wasn't a separate paragraph, add placeholder
            body.append(make_styled_para('MDPI12title', bold=True, text='[Title]'))
            body.append(_build_author_placeholder())
            title_added = True

        if itype == 'abstract':
            body.append(copy_runs_to_para(
                'MDPI17abstract', el, bold_prefix='Abstract: '))
            after_heading = False

        elif itype == 'keywords':
            kw_text = item.get('text', get_para_text(el))
            body.append(make_para('MDPI18keywords', kw_text, bold_prefix='Keywords: '))
            after_heading = False

        elif itype == 'heading1':
            text = item.get('text', get_para_text(el))
            body.append(make_styled_para('MDPI21heading1', bold=True, text=text))
            after_heading = True

        elif itype == 'heading2':
            text = item.get('text', get_para_text(el))
            body.append(make_styled_para('MDPI22heading2', italic=True, text=text))
            after_heading = True

        elif itype == 'heading3':
            text = item.get('text', get_para_text(el))
            body.append(make_styled_para('MDPI23heading3', text=text))
            after_heading = True

        elif itype == 'paragraph':
            style = 'MDPI32textnoindent' if after_heading else 'MDPI31text'
            after_heading = False
            body.append(copy_runs_to_para(style, el))

        elif itype == 'table_caption':
            text = item.get('text', get_para_text(el))
            m = re.match(r'^(Table\s+\d+\.?\s*)', text)
            if m:
                rest = text[len(m.group(1)):]
                body.append(make_para('MDPI41tablecaption',
                            rest, bold_prefix=m.group(1)))
            else:
                body.append(make_para('MDPI41tablecaption', text))
            after_heading = False

        elif itype == 'table':
            tbl = build_mdpi_table(el)
            if tbl is not None:
                body.append(tbl)
            after_heading = False

        elif itype == 'table_footer':
            text = item.get('text', get_para_text(el))
            body.append(make_para('MDPI43tablefooter', text))
            after_heading = False

        elif itype == 'figure_placeholder':
            body.append(copy_runs_to_para('MDPI51figurecaption', el))
            after_heading = False

        elif itype == 'reference':
            body.append(copy_runs_to_para('MDPI32textnoindent', el))
            after_heading = False

    # Restore sectPr
    if sect_pr is not None:
        body.append(sect_pr)

    # ── Write output ──
    # Serialize the new document.xml
    new_doc_xml = ET.tostring(root, encoding='unicode', xml_declaration=False)
    new_doc_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + new_doc_xml

    # Copy the template zip, replacing document.xml
    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
        tmp_path = tmp.name

    with zipfile.ZipFile(template_path, 'r') as zin:
        with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    zout.writestr(item, new_doc_xml.encode('utf-8'))
                elif item.filename == 'word/_rels/settings.xml.rels':
                    # Remove broken local file references from template
                    text = data.decode('utf-8')
                    text = re.sub(
                        r'<Relationship[^>]*Target="file:///[^"]*"[^/]*/>', '', text)
                    zout.writestr(item, text.encode('utf-8'))
                else:
                    zout.writestr(item, data)

    # Move to final output (copy + delete for cross-device compatibility)
    shutil.copy2(tmp_path, output_path)
    os.remove(tmp_path)


def _build_author_placeholder():
    """Build author names paragraph with superscript affiliations."""
    p = ET.Element(qn('w:p'))
    ppr = ET.SubElement(p, qn('w:pPr'))
    ET.SubElement(ppr, qn('w:pStyle')).set(qn('w:val'), 'MDPI13authornames')
    parts = [
        ('Firstname Lastname ', None),
        ('1', ['sup']),
        (', Firstname Lastname ', None),
        ('2', ['sup']),
        (' and Firstname Lastname ', None),
        ('2,', ['sup']),
        ('*', None),
    ]
    for text, props in parts:
        r = ET.SubElement(p, qn('w:r'))
        if props:
            _add_rpr(r, props)
        _make_t(r, text)
    return p


# ── File Selection ──────────────────────────────────────────────────────────

def pick_manuscript():
    """Open a file picker dialog. Returns path or None."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        path = filedialog.askopenfilename(
            title="Select your manuscript (.docx)",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
            initialdir=os.path.dirname(os.path.abspath(__file__))
        )
        root.destroy()
        return path if path else None
    except Exception:
        # Fallback for headless environments
        path = input("Enter manuscript .docx path: ").strip().strip('"').strip("'")
        return path if path else None


def find_template():
    """Find mdpi_template.docx next to the script."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for name in ['mdpi_template.docx', 'MDPI_template.docx']:
        candidate = os.path.join(script_dir, name)
        if os.path.isfile(candidate):
            return candidate
    return None


# ── Main ────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  MDPI Manuscript Formatter")
    print("=" * 60)

    # ── Manuscript ──
    if len(sys.argv) >= 2 and os.path.isfile(sys.argv[1]):
        manuscript_path = os.path.abspath(sys.argv[1])
    else:
        manuscript_path = pick_manuscript()

    if not manuscript_path or not os.path.isfile(manuscript_path):
        print("\nERROR: No manuscript selected.")
        input("Press Enter to exit...")
        sys.exit(1)

    manuscript_path = os.path.abspath(manuscript_path)
    manuscript_dir = os.path.dirname(manuscript_path)
    manuscript_name = os.path.splitext(os.path.basename(manuscript_path))[0]

    # ── Template ──
    template_path = find_template()
    if template_path is None:
        # Also check next to the manuscript
        for name in ['mdpi_template.docx', 'MDPI_template.docx']:
            candidate = os.path.join(manuscript_dir, name)
            if os.path.isfile(candidate):
                template_path = candidate
                break

    if template_path is None:
        print("\nERROR: mdpi_template.docx not found.")
        print("Place it in the same folder as this script.")
        input("Press Enter to exit...")
        sys.exit(1)

    # ── Output ──
    output_path = os.path.join(manuscript_dir, f"{manuscript_name}_MDPI.docx")

    print(f"\n  Manuscript : {manuscript_path}")
    print(f"  Template   : {template_path}")
    print(f"  Output     : {output_path}")
    print()

    # ── Process ──
    print("[1/3] Reading manuscript...")
    items = read_manuscript(manuscript_path)

    types = {}
    for it in items:
        types[it['type']] = types.get(it['type'], 0) + 1
    for t, c in sorted(types.items()):
        print(f"       {t}: {c}")

    print("\n[2/3] Building MDPI-formatted document...")
    build_document(template_path, items, output_path)

    print("[3/3] Done!")
    print(f"\n  >>> Saved to: {output_path}")
    print()
    input("Press Enter to exit...")


if __name__ == '__main__':
    main()
