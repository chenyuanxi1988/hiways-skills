#!/usr/bin/env python
"""
海华永泰律师事务所 - 法律文书模板生成器
Generate Hiways Law Firm formatted .docx from Markdown or structured content.

Usage:
  python generate.py <input.md> [output.docx]
  python generate.py --json <config.json> [output.docx]
"""

import sys
import os
import re
import json
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from lxml import etree
import markdown
from bs4 import BeautifulSoup
from docx.enum.section import WD_ORIENT

# ── Constants ──
ACCENT_COLOR = RGBColor(0x10, 0x6F, 0x86)
BLACK = RGBColor(0x00, 0x00, 0x00)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT_LATIN = 'FangSong'   # 仿宋（与参考文件一致）
FONT_CN = 'FangSong'      # 仿宋

# Page: A4, margins from reference template
PAGE_WIDTH = Cm(21.0)
PAGE_HEIGHT = Cm(29.7)
MARGIN_TOP = Cm(2.54)
MARGIN_BOTTOM = Cm(2.54)
MARGIN_LEFT = Cm(1.91)
MARGIN_RIGHT = Cm(1.91)

# Font sizes in Pt
SIZE_BODY = Pt(12)       # 24 half-points (152400 EMU)
SIZE_HEADING = Pt(14)    # 28 half-points (177800 EMU)
SIZE_SMALL = Pt(9)
SIZE_COVER_INFO = Pt(12)  # 信息区 12pt（与参考文件一致）

# Skill assets directory
SKILL_DIR = Path(__file__).resolve().parent.parent
LOGO_PATH = SKILL_DIR / 'assets' / 'hiways-logo.jpeg'

# ── Document Setup ──
def create_document():
    """Create a blank document with template styles configured."""
    doc = Document()

    # Page setup
    section = doc.sections[0]
    section.page_width = PAGE_WIDTH
    section.page_height = PAGE_HEIGHT
    section.top_margin = MARGIN_TOP
    section.bottom_margin = MARGIN_BOTTOM
    section.left_margin = MARGIN_LEFT
    section.right_margin = MARGIN_RIGHT

    # Configure Normal style
    style = doc.styles['Normal']
    style.font.name = FONT_LATIN
    style.font.size = SIZE_BODY
    style.font.color.rgb = BLACK
    style.element.rPr.rFonts.set(qn('w:eastAsia'), FONT_CN)
    # Line spacing ~1.15
    pf = style.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.15

    # Set east Asian font on all styles
    for s in doc.styles:
        if s.type is not None:
            try:
                s.font.name = FONT_LATIN
                s.element.rPr.rFonts.set(qn('w:eastAsia'), FONT_CN)
            except Exception:
                pass

    return doc


def set_run_font(run, size=SIZE_BODY, bold=False, color=BLACK, font_latin=FONT_LATIN, font_cn=FONT_CN):
    """Configure a run with template-compliant font settings."""
    run.font.name = font_latin
    run.font.size = size
    run.bold = bold
    run.font.color.rgb = color
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} />')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_cn)
    rFonts.set(qn('w:ascii'), font_latin)
    rFonts.set(qn('w:hAnsi'), font_latin)


def add_paragraph(doc, text, size=SIZE_BODY, bold=False, color=BLACK, alignment=None, space_before=0, space_after=0, first_line_indent=None):
    """Add a formatted paragraph."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    if alignment is not None:
        p.alignment = alignment
    pf = p.paragraph_format
    if space_before:
        pf.space_before = Pt(space_before)
    if space_after:
        pf.space_after = Pt(space_after)
    pf.line_spacing = 1.15
    if first_line_indent:
        pf.first_line_indent = first_line_indent
    return p, run


def add_cover_page(doc, title_lines, metadata):
    """Add Hiways cover page matching the reference document layout.

    Reference layout (from 临港项目法律服务方案 v2_20260319.docx):
    Page 1 (Cover):
      - 4 empty paragraphs
      - LOGO: ~6.47cm × ~3.25cm, right-shifted ~5.4cm
      - 4 empty paragraphs
      - Title lines: 14pt bold 仿宋, centered, 1.5x line spacing
      - 3 empty paragraphs
      - "上海市海华永泰律师事务所": 14pt bold 仿宋, centered, 1.5x
      - 7 empty paragraphs
      - "专业赢得尊重，携手成就理想": 12pt bold #106F86, centered
      - 4 empty paragraphs
      - City list: 12pt 仿宋, spread across page
      - Page break

    Page 2 (Info page):
      - 起草人/电话/邮箱/接收人/日期: LEFT-aligned, 12pt 仿宋
      - 主题: LEFT-aligned, 12pt bold
      - Body content follows

    title_lines: list of str (e.g., [' 工程项目', '法律服务方案'])
    metadata: dict with keys: drafter, phone, email, recipient, date, subject
    """
    # ── Page 1: Cover ──

    # 5 empty paragraphs at top (LOGO floats, need to push text down)
    for _ in range(5):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15

    # LOGO: floating/anchored positioning so it won't be clipped by line height
    # Reference uses inline with exact line height, but anchor is more robust
    if LOGO_PATH.exists():
        # Add a minimal paragraph as anchor host
        p = doc.add_paragraph()
        # Set paragraph to small line height so it doesn't take much space
        pPr = p._element.get_or_add_pPr()
        pPr.append(parse_xml(f'<w:spacing {nsdecls("w")} w:line="240" w:lineRule="exact"/>'))

        # Build the anchor drawing XML
        # Position: right-shifted ~5.35cm from left margin, top aligned with this paragraph
        run = p.add_run()
        run.add_picture(str(LOGO_PATH), width=Cm(6.47))
        set_run_font(run, size=SIZE_BODY)
        rPr = run._element.get_or_add_rPr()
        rPr.append(parse_xml(f'<w:noProof {nsdecls("w")}/>'))

        # Convert the inline image to anchor (floating)
        # Find the drawing element we just created
        drawing = run._element.find('.//' + qn('w:drawing'))
        if drawing is not None:
            inline_elem = drawing.find(qn('wp:inline'))
            if inline_elem is not None:
                # Get image dimensions from inline
                extent = inline_elem.find(qn('wp:extent'))
                cx = extent.get('cx') if extent is not None else '2329200'
                cy = extent.get('cy') if extent is not None else '1178860'
                docPr = inline_elem.find(qn('wp:docPr'))
                doc_id = docPr.get('id', '1') if docPr is not None else '1'
                doc_name = docPr.get('name', 'Picture 1') if docPr is not None else 'Picture 1'
                cNvGraphicFramePr = inline_elem.find(qn('wp:cNvGraphicFramePr'))
                graphic = inline_elem.find(qn('a:graphic'))

                # Build anchor element
                anchor_xml = (
                    f'<wp:anchor {nsdecls("wp")} {nsdecls("a")} {nsdecls("pic")} {nsdecls("r")}'
                    f' distT="0" distB="0" distL="0" distR="0"'
                    f' simplePos="0" relativeHeight="251658240"'
                    f' behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
                    f'<wp:simplePos x="0" y="0"/>'
                    f'<wp:positionH relativeFrom="margin">'
                    f'<wp:posOffset>1926000</wp:posOffset>'  # ~5.35cm from left margin
                    f'</wp:positionH>'
                    f'<wp:positionV relativeFrom="paragraph">'
                    f'<wp:posOffset>0</wp:posOffset>'
                    f'</wp:positionV>'
                    f'<wp:extent cx="{cx}" cy="{cy}"/>'
                    f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
                    f'<wp:wrapNone/>'
                    f'<wp:docPr id="{doc_id}" name="{doc_name}"/>'
                    f'<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
                )
                # Copy the graphic element
                if graphic is not None:
                    graphic_xml = etree.tostring(graphic, encoding='unicode')
                    anchor_xml += graphic_xml
                anchor_xml += '</wp:anchor>'

                anchor_elem = parse_xml(anchor_xml)

                # Replace inline with anchor in the drawing element
                drawing.remove(inline_elem)
                drawing.append(anchor_elem)

    # 7 empty paragraphs after logo (since LOGO floats, need extra space to push text down)
    for _ in range(7):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15

    # Filter out title lines that duplicate the firm name (template adds it separately)
    firm_keywords = ('海华永泰', 'HIWAYS', 'Hiways')
    filtered_titles = [t for t in title_lines if not any(kw in t for kw in firm_keywords)]

    # If no title lines remain after filtering, use subject and split smartly
    if not filtered_titles:
        subject = metadata.get('subject', '')
        if subject:
            # Smart split: break at "——", "——", or around 15-20 chars
            if '——' in subject:
                parts = subject.split('——', 1)
                filtered_titles = [parts[0].strip(), parts[1].strip()]
            elif '—' in subject:
                parts = subject.split('—', 1)
                filtered_titles = [parts[0].strip(), parts[1].strip()]
            elif len(subject) > 18:
                # Try to break at a natural point
                mid = len(subject) // 2
                # Find nearest punctuation
                for offset in range(min(10, mid)):
                    for pos in [mid + offset, mid - offset]:
                        if 0 < pos < len(subject) and subject[pos] in '，、；：':
                            filtered_titles = [subject[:pos], subject[pos:]]
                            break
                    if filtered_titles:
                        break
                if not filtered_titles:
                    filtered_titles = [subject]
            else:
                filtered_titles = [subject]
        else:
            filtered_titles = ['法律服务方案']

    # Title lines: 14pt bold 仿宋, centered, 1.5x line spacing
    for line in filtered_titles:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.5
        run = p.add_run(line)
        set_run_font(run, size=SIZE_HEADING, bold=True)

    # 2 empty paragraphs before firm name (reduced from 3)
    for _ in range(2):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.15

    # Firm name: 14pt bold 仿宋, centered, 1.5x
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run('上海市海华永泰律师事务所')
    set_run_font(run, size=SIZE_HEADING, bold=True)

    # 5 empty paragraphs before slogan (reduced from 7)
    for _ in range(5):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15

    # Slogan: 12pt bold #106F86, centered
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6.7)
    p.paragraph_format.line_spacing = Pt(16.2)
    run = p.add_run('专业赢得尊重，携手成就理想')
    set_run_font(run, size=SIZE_BODY, bold=True, color=ACCENT_COLOR)
    # Set letter spacing (w:spacing val=6 twips ≈ reference's spacing)
    rPr = run._element.get_or_add_rPr()
    rPr.append(parse_xml(f'<w:spacing {nsdecls("w")} w:val="6"/>'))

    # 2 empty paragraphs before city list (reduced from 4)
    for _ in range(2):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15

    # City list: 12pt 仿宋, single paragraph with natural wrapping
    cities = '上海｜北京｜长沙｜郑州｜昆明｜成都｜济南｜沈阳｜哈尔滨｜南京｜武汉｜深圳｜银川｜西安｜重庆｜常州｜乌鲁木齐｜宁波｜泰州｜青岛｜贵阳｜杭州｜温州｜苏州｜合肥｜珠海｜南昌｜无锡｜天津｜大连｜洛杉矶｜西雅图｜伦敦｜巴黎｜悉尼'
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = Pt(18.9)
    run = p.add_run(cities)
    set_run_font(run, size=SIZE_BODY)
    # Slightly compress letter spacing for denser layout
    rPr = run._element.get_or_add_rPr()
    rPr.append(parse_xml(f'<w:spacing {nsdecls("w")} w:val="-5"/>'))

    # ── Page 2: Info page ──
    # Page break
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_break(WD_BREAK.PAGE)

    # Metadata fields: LEFT-aligned, 12pt 仿宋 (matching reference)
    meta_fields = [
        ('起草人：', metadata.get('drafter', '陈元熹律师 上海市海华永泰律师事务所 合伙人')),
        ('电话：', metadata.get('phone', '18616552585')),
        ('邮箱：', metadata.get('email', 'alex@hiwayslaw.com')),
        ('接收人：', metadata.get('recipient', '')),
        ('日期：', metadata.get('date', '')),
    ]

    for label, value in meta_fields:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6.3)
        run = p.add_run(label + value)
        set_run_font(run, size=SIZE_BODY)

    # Subject line: LEFT-aligned, 12pt bold
    if metadata.get('subject'):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6.3)
        run = p.add_run('主题：' + metadata['subject'])
        set_run_font(run, size=SIZE_BODY, bold=True)

    # Empty paragraph before body
    doc.add_paragraph()

    return doc


def add_page_header(doc):
    """Add '保密版本' header to all sections."""
    for section in doc.sections:
        header = section.header
        header.is_linked_to_previous = False
        p = header.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run('保密版本')
        set_run_font(run, size=Pt(8), color=RGBColor(0x99, 0x99, 0x99))


def parse_markdown_content(md_text):
    """Parse markdown content and return structured elements.

    Returns a list of elements, each with type and data:
    - {'type': 'h1', 'text': '...'}
    - {'type': 'h2', 'text': '...'}
    - {'type': 'h3', 'text': '...'}
    - {'type': 'para', 'text': '...', 'bold_ranges': [(start, end), ...]}
    - {'type': 'horizontal_rule'}
    - {'type': 'table', 'headers': [...], 'rows': [[...], ...]}
    - {'type': 'list', 'items': [{'text': '...', 'bold_ranges': [...]}, ...], 'ordered': bool}
    - {'type': 'block_meta', 'lines': [('label', 'value'), ...]}
    """
    # Convert markdown to HTML, then parse
    html = markdown.markdown(md_text, extensions=['tables', 'fenced_code'])

    # Fix: markdown puts code blocks in <pre><code>, ignore those for now
    # Use BeautifulSoup to parse
    soup = BeautifulSoup(html, 'html.parser')
    elements = []

    for child in soup.children:
        if not child.name:
            continue

        if child.name in ('h1', 'h2', 'h3'):
            elements.append({
                'type': child.name,
                'text': child.get_text(strip=True)
            })

        elif child.name == 'hr':
            elements.append({'type': 'horizontal_rule'})

        elif child.name == 'p':
            text = child.get_text()
            # Parse bold ranges
            bold_ranges = []
            pos = 0
            for content in child.contents:
                if hasattr(content, 'name') and content.name in ('strong', 'b'):
                    start = pos
                    t = content.get_text()
                    end = start + len(t)
                    bold_ranges.append((start, end))
                    pos = end
                else:
                    pos += len(content.string) if content.string else 0
            elements.append({
                'type': 'para',
                'text': text,
                'bold_ranges': bold_ranges
            })

        elif child.name == 'table':
            headers = []
            rows = []
            thead = child.find('thead')
            tbody = child.find('tbody')
            if thead:
                for th in thead.find_all('th'):
                    headers.append(th.get_text(strip=True))
            else:
                # Use first row as header
                first_row = child.find('tr')
                if first_row:
                    for th in first_row.find_all('th'):
                        headers.append(th.get_text(strip=True))
                    if not headers:
                        for td in first_row.find_all('td'):
                            headers.append(td.get_text(strip=True))
            if tbody:
                for tr in tbody.find_all('tr'):
                    row = [td.get_text(strip=True) for td in tr.find_all('td')]
                    if row:
                        rows.append(row)
            else:
                all_rows = child.find_all('tr')
                start_idx = 1 if headers else 0
                for tr in all_rows[start_idx:]:
                    row = [td.get_text(strip=True) for td in tr.find_all('td')]
                    if row:
                        rows.append(row)
            if headers or rows:
                elements.append({
                    'type': 'table',
                    'headers': headers,
                    'rows': rows
                })

        elif child.name in ('ol', 'ul'):
            items = []
            for li in child.find_all('li'):
                text = li.get_text()
                bold_ranges = []
                pos = 0
                for content in li.contents:
                    if hasattr(content, 'name') and content.name in ('strong', 'b'):
                        start = pos
                        t = content.get_text()
                        end = start + len(t)
                        bold_ranges.append((start, end))
                        pos = end
                    else:
                        pos += len(content.string) if content.string else 0
                items.append({'text': text, 'bold_ranges': bold_ranges})
            elements.append({
                'type': 'list',
                'items': items,
                'ordered': child.name == 'ol'
            })

    return elements


def is_heading_like(text):
    """Check if text looks like a Chinese legal heading (一、/ （一）/ 1.)"""
    return bool(re.match(r'^[一二三四五六七八九十]、|^（[一二三四五六七八九十]）|^\d+[\.\、]', text))


def is_form_label(text):
    """Check if text looks like a form label / metadata field that should be LEFT-aligned.

    Matches patterns like:
      - **交接日期：** 2026年__月__日
      - **移交人声明：**
      - **移交人签名：** ____  **签署时间：** ____
      - 交接编号：TF-INV-001-REC-2026____

    Heuristics: starts with bold label + colon, or contains fill-in blanks (____).
    """
    # Has fill-in blanks (4+ consecutive underscores)
    if re.search(r'_{4,}', text):
        return True
    # Starts with a bold label ending in colon (Chinese or English)
    stripped = text.strip()
    if re.match(r'^\*{0,2}[^*]+\*{0,2}[：:]', stripped):
        return True
    return False


def add_content_elements(doc, elements):
    """Add parsed markdown elements to the document in template format."""
    for elem in elements:
        t = elem['type']

        if t == 'h1':
            # Main title - 14pt bold centered
            add_paragraph(doc, elem['text'], size=SIZE_HEADING, bold=True,
                         alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=12, space_after=6)

        elif t == 'h2':
            # Section heading (一、二、三...) - 14pt bold
            add_paragraph(doc, elem['text'], size=SIZE_HEADING, bold=True,
                         space_before=18, space_after=6)

        elif t == 'h3':
            # Sub-section heading (（一）（二）...) - 12pt bold
            add_paragraph(doc, elem['text'], size=SIZE_BODY, bold=True,
                         space_before=12, space_after=4)

        elif t == 'horizontal_rule':
            # Skip: 不在章节间插入分隔横线
            pass

        elif t == 'para':
            # Check if this paragraph looks like a heading in plain text
            text = elem['text']
            if is_heading_like(text):
                add_paragraph(doc, text, size=SIZE_BODY, bold=True,
                             space_before=12, space_after=4)
            elif is_form_label(text):
                # Form labels / metadata fields: LEFT-aligned, no first-line indent
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                pf = p.paragraph_format
                pf.line_spacing = 1.15

                bold_ranges = elem.get('bold_ranges', [])
                if bold_ranges:
                    pos = 0
                    for start, end in bold_ranges:
                        if start > pos:
                            run = p.add_run(text[pos:start])
                            set_run_font(run, size=SIZE_BODY)
                        run = p.add_run(text[start:end])
                        set_run_font(run, size=SIZE_BODY, bold=True)
                        pos = end
                    if pos < len(text):
                        run = p.add_run(text[pos:])
                        set_run_font(run, size=SIZE_BODY)
                else:
                    # Check for ** markers in raw text
                    clean_text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
                    if clean_text != text:
                        parts = re.split(r'(\*\*.+?\*\*)', text)
                        for part in parts:
                            if part.startswith('**') and part.endswith('**'):
                                run = p.add_run(part[2:-2])
                                set_run_font(run, size=SIZE_BODY, bold=True)
                            else:
                                run = p.add_run(part)
                                set_run_font(run, size=SIZE_BODY)
                    else:
                        run = p.add_run(text)
                        set_run_font(run, size=SIZE_BODY)
            else:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                pf = p.paragraph_format
                pf.line_spacing = 1.15
                pf.first_line_indent = Cm(0.74)  # ~2 chars

                bold_ranges = elem.get('bold_ranges', [])
                if bold_ranges:
                    pos = 0
                    for start, end in bold_ranges:
                        if start > pos:
                            run = p.add_run(text[pos:start])
                            set_run_font(run, size=SIZE_BODY)
                        run = p.add_run(text[start:end])
                        set_run_font(run, size=SIZE_BODY, bold=True)
                        pos = end
                    if pos < len(text):
                        run = p.add_run(text[pos:])
                        set_run_font(run, size=SIZE_BODY)
                else:
                    # Check for ** markers in raw text
                    clean_text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
                    if clean_text != text:
                        # Has bold markers, split accordingly
                        parts = re.split(r'(\*\*.+?\*\*)', text)
                        for part in parts:
                            if part.startswith('**') and part.endswith('**'):
                                run = p.add_run(part[2:-2])
                                set_run_font(run, size=SIZE_BODY, bold=True)
                            else:
                                run = p.add_run(part)
                                set_run_font(run, size=SIZE_BODY)
                    else:
                        run = p.add_run(text)
                        set_run_font(run, size=SIZE_BODY)

        elif t == 'table':
            headers = elem.get('headers', [])
            rows = elem.get('rows', [])
            if not headers and not rows:
                continue

            # Add a small space before table
            add_paragraph(doc, '', space_before=6, space_after=2)

            num_cols = len(headers) if headers else (len(rows[0]) if rows else 1)
            table = doc.add_table(rows=1 + len(rows), cols=num_cols)
            table.style = 'Table Grid'
            table.alignment = WD_TABLE_ALIGNMENT.CENTER

            # ── Fixed layout: prevent Word from auto-expanding columns beyond page ──
            tbl = table._tbl
            tblPr = tbl.tblPr
            if tblPr is None:
                tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}/>')
                tbl.insert(0, tblPr)

            # Remove existing tblLayout / tblW if any
            for existing in tblPr.findall(qn('w:tblLayout')):
                tblPr.remove(existing)
            for existing in tblPr.findall(qn('w:tblW')):
                tblPr.remove(existing)

            # Set fixed layout
            tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))

            # Column width profiles in TWIPS (1 cm ≈ 567 twips / dxa)
            # Available width: 21.0 - 1.91 - 1.91 = 17.18cm ≈ 9743 twips
            COL_WIDTHS_TWIPS = {
                6: [510, 907, 3289, 1021, 1588, 2313],     # total = 9628
                7: [510, 1134, 1985, 1021, 1021, 2155, 1917], # total = 9743
            }
            width_twips = COL_WIDTHS_TWIPS.get(num_cols,
                        [int(9743 / num_cols)] * num_cols)
            total_twips = sum(width_twips)

            # Set table total width to fixed
            tblPr.append(parse_xml(
                f'<w:tblW {nsdecls("w")} w:type="dxa" w:w="{total_twips}"/>'
            ))

            # Set gridCol widths (in twips / dxa)
            tblGrid = tbl.find(qn('w:tblGrid'))
            if tblGrid is not None:
                gridCols = tblGrid.findall(qn('w:gridCol'))
                for i, gc in enumerate(gridCols):
                    if i < len(width_twips):
                        gc.set(qn('w:w'), str(width_twips[i]))

            # Set cell widths (cell.width uses EMU internally) & compact cell margins
            width_emu = [tw * 635 for tw in width_twips]  # 1 twip = 635 EMU
            for ri, row in enumerate(table.rows):
                for ci, cell in enumerate(row.cells):
                    if ci < len(width_emu):
                        cell.width = width_emu[ci]
                    # Compact cell margins: top/bottom 1pt, left/right 2pt
                    tc = cell._tc
                    tcPr = tc.get_or_add_tcPr()
                    for existing in tcPr.findall(qn('w:tcMar')):
                        tcPr.remove(existing)
                    tcPr.append(parse_xml(
                        f'<w:tcMar {nsdecls("w")}>'
                        f'<w:top w:w="20" w:type="dxa"/>'
                        f'<w:left w:w="40" w:type="dxa"/>'
                        f'<w:bottom w:w="20" w:type="dxa"/>'
                        f'<w:right w:w="40" w:type="dxa"/>'
                        f'</w:tcMar>'
                    ))

            # Header row
            if headers:
                hdr_cells = table.rows[0].cells
                for i, h in enumerate(headers):
                    if i < num_cols:
                        p = hdr_cells[i].paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = p.add_run(h)
                        set_run_font(run, size=Pt(9), bold=True)
                        # Light gray background
                        shading = parse_xml(
                            f'<w:shd {nsdecls("w")} w:fill="E8E8E8" w:val="clear"/>'
                        )
                        hdr_cells[i]._element.get_or_add_tcPr().append(shading)

            # Data rows
            for ri, row_data in enumerate(rows):
                row_cells = table.rows[ri + 1].cells
                for ci, cell_text in enumerate(row_data):
                    if ci < num_cols:
                        p = row_cells[ci].paragraphs[0]
                        # Center-align short columns (序号, 提供形态, 备注)
                        if ci in (0, 3) or (num_cols == 7 and ci == 6):
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = p.add_run(str(cell_text))
                        set_run_font(run, size=Pt(8))

            # Space after table
            add_paragraph(doc, '', space_before=6)

        elif t == 'list':
            is_ordered = elem.get('ordered', False)
            for idx, item in enumerate(elem.get('items', [])):
                prefix = f'{idx + 1}. ' if is_ordered else '  '
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                pf = p.paragraph_format
                pf.line_spacing = 1.15
                pf.left_indent = Cm(0.74)

                text = item['text']
                bold_ranges = item.get('bold_ranges', [])
                if bold_ranges:
                    run = p.add_run(prefix)
                    set_run_font(run, size=SIZE_BODY)
                    pos = 0
                    for start, end in bold_ranges:
                        if start > pos:
                            run = p.add_run(text[pos:start])
                            set_run_font(run, size=SIZE_BODY)
                        run = p.add_run(text[start:end])
                        set_run_font(run, size=SIZE_BODY, bold=True)
                        pos = end
                    if pos < len(text):
                        run = p.add_run(text[pos:])
                        set_run_font(run, size=SIZE_BODY)
                else:
                    run = p.add_run(prefix + text)
                    set_run_font(run, size=SIZE_BODY)

    return doc


def add_footer_info(doc, metadata):
    """Add closing information at the end of the document."""
    # Blank space
    add_paragraph(doc, '', space_before=24)

    # Separator
    add_paragraph(doc, '─' * 40, size=Pt(8), color=RGBColor(0xCC, 0xCC, 0xCC),
                 alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=12, space_after=12)

    if metadata.get('closing_name'):
        add_paragraph(doc, metadata['closing_name'], size=SIZE_BODY,
                     alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_after=4)

    add_paragraph(doc, '上海市海华永泰律师事务所', size=SIZE_BODY,
                 alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_after=4)

    if metadata.get('date'):
        add_paragraph(doc, metadata['date'], size=SIZE_BODY,
                     alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_after=12)

    # Disclaimer
    if metadata.get('disclaimer'):
        p = doc.add_paragraph()
        run = p.add_run(metadata['disclaimer'])
        set_run_font(run, size=Pt(9), color=RGBColor(0x99, 0x99, 0x99))


def extract_markdown_header(md_text):
    """Extract cover metadata from markdown header section (before first ## heading).

    Returns (metadata_dict, body_text).
    metadata_dict contains: title_lines, doc_id, classification, doc_date,
                            recipient, drafter, subject, date
    """
    metadata = {}
    body_text = md_text

    # Split at first ## heading - everything before is header
    body_match = re.search(r'\n##\s+', md_text)
    if body_match:
        header_text = md_text[:body_match.start()].strip()
        body_text = md_text[body_match.start():].strip()
    else:
        header_text = md_text
        body_text = ''

    lines = header_text.split('\n')

    # Extract title lines (# headings at top)
    title_lines = []
    i = 0
    while i < len(lines) and lines[i].strip().startswith('#'):
        title_lines.append(lines[i].lstrip('#').strip())
        i += 1

    if title_lines:
        metadata['title_lines'] = title_lines

    # Parse metadata fields after the HR separator
    current_key = None
    for line in lines:
        line = line.strip()
        if not line or line.startswith('#'):
            continue

        # Skip horizontal rules
        if set(line) <= {'─', '─', '－', '-', '—'} or all(c in '─＝－—' for c in line if c.strip()):
            continue

        # Match "标签：值" or "**标签：** 值" format
        m = re.match(r'^(?:\*\*)?(.+?)(?:\*\*)?[：:]\s*(.*)', line)
        if m:
            key = m.group(1).strip().strip('*').strip()
            val = re.sub(r'^\*+\s*', '', m.group(2).strip()).strip()
            key_lower = key.lower()

            if '文件编号' in key or 'doc' in key_lower:
                metadata['doc_id'] = val
            elif '密级' in key or 'classification' in key_lower:
                metadata['classification'] = val
            elif '日期' in key or 'date' in key_lower:
                metadata['date'] = val
            elif '收文' in key or 'recipient' in key_lower or '接收' in key:
                metadata['recipient'] = val
            elif '发文' in key or 'drafter' in key_lower or '起草' in key:
                metadata['drafter'] = val
            elif '主题' in key or 'subject' in key_lower:
                metadata['subject'] = val

    # Use subject as title if no explicit title lines found
    if not metadata.get('title_lines') and metadata.get('subject'):
        metadata['title_lines'] = [metadata['subject']]

    return metadata, body_text


def generate_from_markdown(md_path, output_path, metadata=None):
    """Generate a formatted .docx from a markdown file.

    Automatically extracts cover metadata from the markdown header section
    (everything before the first ## heading). Alternatively, metadata can be
    provided via frontmatter (--- delimited YAML block) or as a dict argument.

    Metadata keys:
      - title_lines: list of str (cover title lines)
      - drafter, phone, email, recipient, date, subject
      - doc_id, classification
      - closing_name, disclaimer
    """
    if metadata is None:
        metadata = {}

    with open(md_path, 'r', encoding='utf-8') as f:
        md_text = f.read()

    # Step 1: Try YAML frontmatter
    fm = {}
    if md_text.startswith('---'):
        parts = md_text.split('---', 2)
        if len(parts) >= 3:
            for line in parts[1].strip().split('\n'):
                if ':' in line:
                    key, val = line.split(':', 1)
                    fm[key.strip()] = val.strip()
            md_text = parts[2]

    # Step 2: Extract header metadata from markdown body
    header_meta, body_text = extract_markdown_header(md_text)

    # Merge: CLI args > frontmatter > header extraction > defaults
    metadata = {**header_meta, **fm, **metadata}

    # Parse markdown body only
    elements = parse_markdown_content(body_text)

    # Create document
    doc = create_document()

    # Build title lines
    title = metadata.get('title_lines', None)
    if not title:
        title = ['法律服务方案']

    # Cover page (now includes info page with subject line)
    add_cover_page(doc, title, metadata)

    # Body content
    add_content_elements(doc, elements)

    # Footer
    add_footer_info(doc, metadata)

    # Page header
    add_page_header(doc)

    # Save
    doc.save(output_path)
    print(f'Generated: {output_path}')
    return output_path


# ── CLI ──
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else input_path.rsplit('.', 1)[0] + '.docx'

    generate_from_markdown(input_path, output_path)
