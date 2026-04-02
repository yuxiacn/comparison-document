#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Document Comparison Tool
Supports formats: PDF, DOCX, PPTX, TXT
Output: Comparison_File1_VS_File2.docx (landscape, diff-only view, word-level highlighting)
"""

import sys
import os
import re
import difflib
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Version number
# Format: V{major}.{minor} Build{YYYYMMDD}.{revision}
# Update Rules:
#   - Major updates: increment major version (e.g., V2.0 → V3.0)
#   - Feature changes: increment minor version (e.g., V2.0 → V2.1)
#   - Bug fixes: increment revision (e.g., V2.0 Build20260403.1 → V2.0 Build20260403.2)
#   - Update date and revision with each modification
VERSION = "V2.0 Build20260403.3"

# Reader registry
READERS = {}


def register_reader(ext):
    def decorator(func):
        READERS[ext.lower()] = func
        return func
    return decorator


def estimate_visual_lines(text, chars_per_line=80):
    """
    Estimate the visual line count of text (simulating Word auto-wrap)
    Consider mixed content, wide characters take ~2x narrow character width
    """
    if not text.strip():
        return 1  # Empty line counts as one line
    
    # Calculate effective character count (wide characters count as 2 width)
    effective_chars = 0
    for char in text:
        if ord(char) > 127:  # Non-ASCII characters (wide)
            effective_chars += 2
        else:
            effective_chars += 1
    
    # Calculate required lines
    lines_needed = max(1, (effective_chars + chars_per_line - 1) // chars_per_line)
    return lines_needed


@register_reader('.txt')
def read_txt(path, merge_lines=False):
    """Read TXT file, return content list and location info (paragraph number, page number=1, None)"""
    with open(path, 'r', encoding='utf-8') as f:
        text = f.read()
    
    if merge_lines:
        # Merge consecutive non-empty lines (handle auto-wrap cases)
        lines = text.splitlines()
        paragraphs = []
        current_para = []
        
        for line in lines:
            stripped = line.rstrip()
            if not stripped:
                if current_para:
                    paragraphs.append(' '.join(current_para))
                    current_para = []
            else:
                current_para.append(stripped)
        
        if current_para:
            paragraphs.append(' '.join(current_para))
        
        lines = paragraphs
    else:
        lines = [ln.rstrip() for ln in text.splitlines() if ln.strip()]
    
    # TXT has no real page number and line number
    location_info = [(i + 1, 1, None) for i in range(len(lines))]
    return lines, location_info


def estimate_paragraph_pages(doc):
    """
    Estimate page number for each paragraph
    Returns: [page_number, ...] corresponding to paragraphs
    """
    # Get page settings
    section = doc.sections[0] if doc.sections else None
    
    if section:
        # Page height (inches)
        page_height = section.page_height.inches if section.page_height else 11
        # Top and bottom margins
        top_margin = section.top_margin.inches if section.top_margin else 1
        bottom_margin = section.bottom_margin.inches if section.bottom_margin else 1
        # Available height
        available_height = page_height - top_margin - bottom_margin
    else:
        available_height = 9  # Default available height
    
    page_numbers = []
    current_page = 1
    current_page_used_height = 0
    
    # Estimate lines per inch for 12pt text (about 5-6 lines)
    lines_per_inch = 5.5
    
    for para in doc.paragraphs:
        text = para.text.rstrip()
        
        # Get font size
        font_size = 12
        try:
            if para.runs and para.runs[0].font.size:
                font_size = para.runs[0].font.size.pt
        except:
            pass
        
        # Calculate paragraph height (inches)
        # Larger font means larger line height
        line_height = (font_size / 12) * (1 / lines_per_inch)
        
        # Estimate paragraph lines (simple estimation)
        effective_chars = sum(2 if ord(c) > 127 else 1 for c in text)
        chars_per_line = 80  # Simplified estimation
        lines_needed = max(1, (effective_chars + chars_per_line - 1) // chars_per_line)
        
        para_height = lines_needed * line_height
        
        # Paragraph spacing before and after
        space_before = 0
        space_after = 0
        try:
            if para.paragraph_format.space_before:
                space_before = para.paragraph_format.space_before.pt / 72  # Convert to inches
            if para.paragraph_format.space_after:
                space_after = para.paragraph_format.space_after.pt / 72
        except:
            pass
        
        total_height = space_before + para_height + space_after
        
        # Check if page break needed
        if current_page_used_height + total_height > available_height:
            current_page += 1
            current_page_used_height = total_height
        else:
            current_page_used_height += total_height
        
        page_numbers.append(current_page)
    
    return page_numbers


@register_reader('.docx')
def read_docx(path, use_precise=True):
    """
    Read docx file, return content list and location info
    Location format: [(paragraph_number, page_number, line_number), ...]
    DOCX has no visual line numbers, so line_number is None
    """
    doc = Document(path)
    lines = [para.text.rstrip() for para in doc.paragraphs]
    
    # Estimate page number
    page_numbers = estimate_paragraph_pages(doc)
    
    # Build location info (paragraph number, page number, None)
    location_info = []
    for i, page_num in enumerate(page_numbers):
        location_info.append((i + 1, page_num, None))
    
    return lines, location_info


@register_reader('.pdf')
def read_pdf(path, merge_lines=True, merge_across_pages=True):
    """
    Read PDF file, return content list and location info (paragraph number, page number, line number)
    
    Args:
        merge_lines: whether to merge consecutive non-empty lines into a paragraph
        merge_across_pages: whether to merge paragraphs across pages
    
    Returns:
        paragraphs: list of text content
        location_info: [(paragraph_num, page_num, line_num), ...]
    """
    try:
        import pdfplumber
    except ImportError:
        raise ImportError("Please install pdfplumber for PDF support: pip install pdfplumber")
    
    paragraphs = []
    location_info = []
    paragraph_counter = 0
    
    # Pattern to identify page numbers: standalone numeric lines (1-4 digits)
    import re
    page_number_pattern = re.compile(r'^\s*\d{1,4}\s*$')
    # Pattern to identify visual line numbers at line start: (spaces+digits+space or dot)
    line_number_pattern = re.compile(r'^(\s*\d+)[\.\s]\s*')
    
    # Collect raw line info from all pages first
    all_pages_lines = []  # [(page_num, lines), ...]
    
    with pdfplumber.open(path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            text = page.extract_text()
            if not text:
                all_pages_lines.append((page_num, []))
                continue
            
            page_lines = text.splitlines()
            processed_lines = []
            
            for line in page_lines:
                line = line.rstrip()
                if not line:
                    continue
                # Filter page number lines
                if page_number_pattern.match(line):
                    continue
                
                # Extract visual line number at line start
                visual_line_num = None
                match = line_number_pattern.match(line)
                if match:
                    try:
                        visual_line_num = int(match.group(1).strip())
                        line = line[match.end():].lstrip()
                    except ValueError:
                        pass
                
                if line:
                    processed_lines.append((line, visual_line_num))
            
            all_pages_lines.append((page_num, processed_lines))
    
    # Cross-page paragraph merging
    if merge_across_pages and merge_lines:
        # Merge all page lines, then process paragraphs uniformly
        all_lines = []
        for page_num, lines in all_pages_lines:
            for line, line_num in lines:
                all_lines.append((line, line_num, page_num))
        
        # Uniform paragraph merging
        current_paragraph = []
        current_line_num = None
        current_page = None
        
        i = 0
        while i < len(all_lines):
            line, visual_line_num, page_num = all_lines[i]
            
            # Check if new paragraph starts
            is_new_para = False
            
            if not current_paragraph:
                # First paragraph starts
                is_new_para = True
            else:
                # Check if current line belongs to new paragraph
                stripped = line.lstrip()
                indent = len(line) - len(stripped)
                
                # Condition 1: obvious indent (4+ spaces)
                if indent >= 4:
                    is_new_para = True
                # Condition 2: previous line ends with period etc, current starts with uppercase or digit
                else:
                    prev_line = current_paragraph[-1]
                    if prev_line and prev_line[-1] in '.!?':
                        if stripped and (stripped[0].isupper() or stripped[0].isdigit()):
                            is_new_para = True
                    # Condition 3: line number reset (obviously smaller)
                    if visual_line_num and current_line_num:
                        if visual_line_num < current_line_num and visual_line_num < 10:
                            is_new_para = True
            
            if is_new_para and current_paragraph:
                # Save current paragraph
                paragraph_counter += 1
                para_text = ' '.join(current_paragraph)
                paragraphs.append(para_text)
                location_info.append((paragraph_counter, current_page or page_num, current_line_num))
                
                # Start new paragraph
                current_paragraph = [line]
                current_line_num = visual_line_num
                current_page = page_num
            else:
                # Continue current paragraph
                if not current_paragraph:
                    current_line_num = visual_line_num
                    current_page = page_num
                current_paragraph.append(line)
            
            i += 1
        
        # Save last paragraph
        if current_paragraph:
            paragraph_counter += 1
            para_text = ' '.join(current_paragraph)
            paragraphs.append(para_text)
            location_info.append((paragraph_counter, current_page, current_line_num))
    
    else:
        # Process page by page, no cross-page merging
        for page_num, lines in all_pages_lines:
            if merge_lines:
                current_paragraph = []
                current_line_num = None
                
                for line, visual_line_num in lines:
                    if not line:
                        if current_paragraph:
                            paragraph_counter += 1
                            para_text = ' '.join(current_paragraph)
                            paragraphs.append(para_text)
                            location_info.append((paragraph_counter, page_num, current_line_num))
                            current_paragraph = []
                            current_line_num = None
                        continue
                    
                    # Check new paragraph
                    is_new_para = False
                    if current_paragraph:
                        stripped = line.lstrip()
                        indent = len(line) - len(stripped)
                        if indent >= 4:
                            is_new_para = True
                        prev_line = current_paragraph[-1]
                        if prev_line and prev_line[-1] in '.!?':
                            if stripped and not stripped[0].islower():
                                is_new_para = True
                    
                    if is_new_para and current_paragraph:
                        paragraph_counter += 1
                        para_text = ' '.join(current_paragraph)
                        paragraphs.append(para_text)
                        location_info.append((paragraph_counter, page_num, current_line_num))
                        current_paragraph = [line]
                        current_line_num = visual_line_num
                    else:
                        if not current_paragraph:
                            current_line_num = visual_line_num
                        current_paragraph.append(line)
                
                # Save last paragraph
                if current_paragraph:
                    paragraph_counter += 1
                    para_text = ' '.join(current_paragraph)
                    paragraphs.append(para_text)
                    location_info.append((paragraph_counter, page_num, current_line_num))
            else:
                # No merging, each line independent
                for line, visual_line_num in lines:
                    if line:
                        paragraph_counter += 1
                        paragraphs.append(line)
                        location_info.append((paragraph_counter, page_num, visual_line_num))
    
    return paragraphs, location_info


@register_reader('.pptx')
def read_pptx(path):
    """Read PPTX file, return content list and location info (paragraph number, slide number, None)"""
    from pptx import Presentation
    prs = Presentation(path)
    lines = []
    location_info = []
    paragraph_counter = 0
    
    for slide_num, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                for ln in shape.text.splitlines():
                    line = ln.rstrip()
                    if line:  # Only process non-empty lines
                        paragraph_counter += 1
                        lines.append(line)
                        location_info.append((paragraph_counter, slide_num, None))
    
    return lines, location_info


def read_document(path, use_precise=True, merge_lines=True, merge_across_pages=True):
    """
    Read document content
    use_precise: for DOCX, whether to use precise visual line number calculation
    merge_lines: for PDF/TXT, whether to merge consecutive lines
    merge_across_pages: for PDF, whether to merge paragraphs across pages
    """
    ext = Path(path).suffix.lower()
    if ext not in READERS:
        raise ValueError(f"Unsupported file format: {ext}. Currently supported: {', '.join(READERS.keys())}")
    
    reader = READERS[ext]
    # Pass appropriate parameters for different formats
    if ext == '.docx':
        return reader(path, use_precise=use_precise)
    elif ext == '.pdf':
        return reader(path, merge_lines=merge_lines, merge_across_pages=merge_across_pages)
    elif ext == '.txt':
        return reader(path, merge_lines=merge_lines)
    else:
        return reader(path)


def set_run_font(run, font_family='"Times New Roman", "Noto Serif", Georgia, serif', size=Pt(10), bold=False):
    """Set font family for the run"""
    run.font.name = font_family
    run.font.size = size
    if bold:
        run.font.bold = True


def add_colored_run(paragraph, text, rgb, bold=False):
    run = paragraph.add_run(text)
    set_run_font(run, bold=bold)
    run.font.color.rgb = RGBColor(*rgb)
    return run


def add_page_number(paragraph):
    """Add page number field to paragraph"""
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')

    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')

    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    return run


def tokenize_text(text):
    """
    Split text into words and separators (spaces/punctuation etc).
    Preserve all content for later reassembly.
    """
    return re.findall(r'\S+|\s+', text)


def split_sentences(text):
    """
    Split text by sentences, keeping delimiters.
    Sentence delimiters: period, question mark, exclamation mark
    """
    if not text.strip():
        return []
    # Split by sentence delimiters, keeping delimiters
    parts = re.split(r'([.!?])', text)
    sentences = []
    current = ''
    for part in parts:
        current += part
        if re.match(r'[.!?]$', part):
            # Encountered sentence end, complete a sentence
            sentences.append(current)
            current = ''
    if current.strip():
        # Remaining content (possibly incomplete sentence) also as a sentence
        sentences.append(current)
    return sentences


def word_diff_runs(text1, text2):
    """
    Word-level fine-grained difference analysis, supporting sentence-level missing detection.
    Returns left_runs and right_runs, each is a list of (text, is_diff, is_placeholder).
    is_placeholder=True indicates a missing content placeholder [Missing Sentence]
    
    Logic:
    1. First split text into sentences
    2. Compare at sentence level
    3. When one side has a complete sentence (with delimiter) and the other does not, show placeholder
    """
    # First try sentence-level comparison
    sents1 = split_sentences(text1)
    sents2 = split_sentences(text2)
    
    # If both can be split into sentences, do sentence-level comparison
    if len(sents1) > 0 and len(sents2) > 0:
        sm_sents = difflib.SequenceMatcher(None, sents1, sents2, autojunk=False)
        opcodes = sm_sents.get_opcodes()
        
        # Check for sentence-level delete/insert
        has_sentence_level_diff = any(
            tag in ('delete', 'insert') for tag, _, _, _, _ in opcodes
        )
        
        if has_sentence_level_diff:
            left_runs = []
            right_runs = []
            
            for tag, i1, i2, j1, j2 in opcodes:
                if tag == 'equal':
                    for k in range(i1, min(i2, len(sents1))):
                        left_runs.append((sents1[k], False, False))
                    for k in range(j1, min(j2, len(sents2))):
                        right_runs.append((sents2[k], False, False))
                elif tag == 'replace':
                    # Replace: different content on both sides
                    max_len = max(i2 - i1, j2 - j1)
                    for k in range(max_len):
                        if i1 + k < i2 and i1 + k < len(sents1):
                            left_runs.append((sents1[i1 + k], True, False))
                        if j1 + k < j2 and j1 + k < len(sents2):
                            right_runs.append((sents2[j1 + k], True, False))
                elif tag == 'delete':
                    # Left has sentence, right missing entire sentence
                    for k in range(i1, min(i2, len(sents1))):
                        left_runs.append((sents1[k], True, False))
                        right_runs.append(('[Missing Sentence]', True, True))
                elif tag == 'insert':
                    # Right has sentence, left missing entire sentence
                    for k in range(j1, min(j2, len(sents2))):
                        left_runs.append(('[Missing Sentence]', True, True))
                        right_runs.append((sents2[k], True, False))
            
            # Merge consecutive [Missing Sentence] placeholders
            def merge_consecutive_placeholders(runs):
                if not runs:
                    return runs
                merged = [runs[0]]
                for i in range(1, len(runs)):
                    prev_text, prev_diff, prev_placeholder = merged[-1]
                    curr_text, curr_diff, curr_placeholder = runs[i]
                    # If previous is placeholder and current is also placeholder, skip current
                    if prev_placeholder and curr_placeholder:
                        continue
                    merged.append((curr_text, curr_diff, curr_placeholder))
                return merged

            left_runs = merge_consecutive_placeholders(left_runs)
            right_runs = merge_consecutive_placeholders(right_runs)
            
            return left_runs, right_runs
    
    # Fallback to word-level comparison (no placeholders)
    tokens1 = tokenize_text(text1)
    tokens2 = tokenize_text(text2)
    sm = difflib.SequenceMatcher(None, tokens1, tokens2, autojunk=False)

    left_runs = []
    right_runs = []

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            seg1 = ''.join(tokens1[i1:i2])
            left_runs.append((seg1, False, False))
            seg2 = ''.join(tokens2[j1:j2])
            right_runs.append((seg2, False, False))
        elif tag == 'replace':
            seg1 = ''.join(tokens1[i1:i2])
            left_runs.append((seg1, True, False))
            seg2 = ''.join(tokens2[j1:j2])
            right_runs.append((seg2, True, False))
        elif tag == 'delete':
            seg1 = ''.join(tokens1[i1:i2])
            left_runs.append((seg1, True, False))
        elif tag == 'insert':
            seg2 = ''.join(tokens2[j1:j2])
            right_runs.append((seg2, True, False))

    if not left_runs:
        left_runs.append(('', False, False))
    if not right_runs:
        right_runs.append(('', False, False))

    return left_runs, right_runs


def get_word_line_number_offset(doc_path):
    """
    Try to read Word document line number settings, return line number start offset.
    If document has line numbers enabled, return line number start value; otherwise return None.
    """
    try:
        doc = Document(doc_path)
        for section in doc.sections:
            # Try to get line number settings
            sectPr = section._sectPr
            if hasattr(sectPr, 'lnNumType') and sectPr.lnNumType is not None:
                # Document has line numbers enabled
                lnNumType = sectPr.lnNumType
                start = getattr(lnNumType, 'start', 1)
                return int(start) if start else 1
    except Exception:
        pass
    return None


def build_diff_report(lines1, lines2, location_info1=None, location_info2=None):
    """
    Use difflib.SequenceMatcher to analyze differences, return only diff rows.
    Each element is (tag, left_loc, left_text, right_loc, right_text)
    tag values: replace, delete, insert
    location format: (paragraph_number, page_number) or None
    """
    sm = difflib.SequenceMatcher(None, lines1, lines2, autojunk=False)
    opcodes = sm.get_opcodes()
    
    # If no location info provided, use default paragraph index+1
    if location_info1 is None:
        location_info1 = [(i + 1, 1) for i in range(len(lines1))]
    if location_info2 is None:
        location_info2 = [(i + 1, 1) for i in range(len(lines2))]

    rows = []
    for tag, i1, i2, j1, j2 in opcodes:
        if tag == 'equal':
            continue
        elif tag == 'replace':
            max_len = max(i2 - i1, j2 - j1)
            for k in range(max_len):
                left_exists = i1 + k < i2
                right_exists = j1 + k < j2
                left_loc = location_info1[i1 + k] if left_exists else None
                right_loc = location_info2[j1 + k] if right_exists else None
                ltext = lines1[i1 + k] if left_exists else ""
                rtext = lines2[j1 + k] if right_exists else ""
                if left_exists and right_exists:
                    rows.append(('replace', left_loc, ltext, right_loc, rtext))
                elif left_exists and not right_exists:
                    rows.append(('delete', left_loc, ltext, None, ""))
                elif not left_exists and right_exists:
                    rows.append(('insert', None, "", right_loc, rtext))
        elif tag == 'delete':
            for k in range(i2 - i1):
                left_loc = location_info1[i1 + k]
                ltext = lines1[i1 + k]
                rows.append(('delete', left_loc, ltext, None, ""))
        elif tag == 'insert':
            for k in range(j2 - j1):
                right_loc = location_info2[j1 + k]
                rtext = lines2[j1 + k]
                rows.append(('insert', None, "", right_loc, rtext))

    return rows


def set_cell_width(cell, width_inches):
    """Force set cell width via underlying XML"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(int(width_inches * 1440)))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)


def set_table_column_widths(table, widths_inches):
    """
    Precisely set table column widths and total width via underlying XML.
    widths_inches: list of width values in inches for each column
    """
    tbl = table._tbl
    tblGrid = tbl.tblGrid
    gridCols = tblGrid.gridCol_lst
    for idx, w in enumerate(widths_inches):
        if idx < len(gridCols):
            gridCols[idx].set(qn('w:w'), str(int(w * 1440)))
            gridCols[idx].set(qn('w:type'), 'dxa')
        if idx < len(table.columns):
            table.columns[idx].width = Inches(w)

    tblPr = tbl.tblPr
    tblW_list = tblPr.xpath('./w:tblW')
    if tblW_list:
        tblW = tblW_list[0]
    else:
        tblW = OxmlElement('w:tblW')
        tblPr.append(tblW)
    total_width = sum(widths_inches)
    tblW.set(qn('w:w'), str(int(total_width * 1440)))
    tblW.set(qn('w:type'), 'dxa')


def generate_docx(rows, name1, name2, output_path):
    doc = Document()

    # Set to landscape orientation with standard 1 inch margins
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)

    # Add footer page number
    footer = section.footer
    footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = add_page_number(footer_para)
    set_run_font(run, size=Pt(9))

    # Color definitions
    color_left = (0, 112, 192)    # Blue
    color_right = (255, 0, 0)     # Red
    color_mark_replace = (255, 140, 0)  # Orange
    color_gray = (100, 100, 100)
    color_placeholder = (0, 176, 80)  # Green, for missing sentences

    # First line: Document Comparison Report
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_title.add_run('Document Comparison Report')
    set_run_font(run, size=Pt(18), bold=True)

    # Second line: two document names
    p_names = doc.add_paragraph()
    p_names.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_names.add_run(name1)
    set_run_font(run, size=Pt(14), bold=True)
    run.font.color.rgb = RGBColor(*color_left)
    run = p_names.add_run(' VS ')
    set_run_font(run, size=Pt(14), bold=True)
    run = p_names.add_run(name2)
    set_run_font(run, size=Pt(14), bold=True)
    run.font.color.rgb = RGBColor(*color_right)

    # Third line: generation date and time
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M')
    p_time = doc.add_paragraph()
    p_time.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_time.add_run(now_str)
    set_run_font(run, size=Pt(12), bold=True)
    run.font.color.rgb = RGBColor(*color_gray)

    # Empty line between time and legend
    doc.add_paragraph()

    # Legend description
    p = doc.add_paragraph()
    run = p.add_run('Legend: ')
    set_run_font(run, bold=False)
    run = p.add_run(' = : identical content (hidden)  ')
    set_run_font(run, bold=False)
    run = p.add_run('≠ ')
    set_run_font(run, bold=False)
    run.font.color.rgb = RGBColor(*color_mark_replace)
    run = p.add_run(': modified content  ')
    set_run_font(run, bold=False)
    run = p.add_run('- ')
    set_run_font(run, bold=False)
    run.font.color.rgb = RGBColor(*color_left)
    run = p.add_run(': deleted content  ')
    set_run_font(run, bold=False)
    run = p.add_run('+ ')
    set_run_font(run, bold=False)
    run.font.color.rgb = RGBColor(*color_right)
    run = p.add_run(': added content  ')
    set_run_font(run, bold=False)
    run = p.add_run('P: page number  L: line number')
    set_run_font(run, bold=False)

    if not rows:
        p = doc.add_paragraph()
        run = p.add_run('Both documents are identical, no differences found.')
        set_run_font(run, size=Pt(12))
        run.font.color.rgb = RGBColor(128, 128, 128)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.save(output_path)
        return

    # Table: 5 columns (total width 8.8 inches, fits 1-inch margin landscape page)
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    table.autofit = False
    table.allow_autofit = False

    col_widths = [0.75, 3.50, 0.60, 0.75, 3.50]
    set_table_column_widths(table, col_widths)

    # Set headers - display page number-paragraph number
    hdr_cells = table.rows[0].cells
    headers = ['Location', name1, 'Mark', 'Location', name2]
    header_colors = [None, color_left, None, None, color_right]
    for idx, text in enumerate(headers):
        cell = hdr_cells[idx]
        set_cell_width(cell, col_widths[idx])
        p = cell.paragraphs[0]
        p.clear()
        run = p.add_run(text)
        set_run_font(run, bold=True)
        if header_colors[idx]:
            run.font.color.rgb = RGBColor(*header_colors[idx])
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for tag, left_loc, ltext, right_loc, rtext in rows:
        row_cells = table.add_row().cells

        # Location 1 - Format: Ppage number-Lline number (or Ppage number-paragraph number)
        cell = row_cells[0]
        set_cell_width(cell, col_widths[0])
        p = cell.paragraphs[0]
        p.clear()
        if left_loc is not None:
            # Handle tuple (para_num, page_num, line_num) or (para_num, page_num)
            if len(left_loc) >= 3:
                para_num, page_num, line_num = left_loc[0], left_loc[1], left_loc[2]
                if line_num:
                    loc_text = f"P{page_num}-L{line_num}"
                else:
                    loc_text = f"P{page_num}-{para_num}"
            else:
                para_num, page_num = left_loc[0], left_loc[1]
                loc_text = f"P{page_num}-{para_num}"
        else:
            loc_text = ""
        run = p.add_run(loc_text)
        set_run_font(run, size=Pt(8))
        run.font.color.rgb = RGBColor(*color_gray)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Location 2 - Format: Ppage number-Lline number (or Ppage number-paragraph number)
        cell = row_cells[3]
        set_cell_width(cell, col_widths[3])
        p = cell.paragraphs[0]
        p.clear()
        if right_loc is not None:
            # Handle tuple or triplet
            if len(right_loc) >= 3:
                para_num, page_num, line_num = right_loc[0], right_loc[1], right_loc[2]
                if line_num:
                    loc_text = f"P{page_num}-L{line_num}"
                else:
                    loc_text = f"P{page_num}-{para_num}"
            else:
                para_num, page_num = right_loc[0], right_loc[1]
                loc_text = f"P{page_num}-{para_num}"
        else:
            loc_text = ""
        run = p.add_run(loc_text)
        set_run_font(run, size=Pt(8))
        run.font.color.rgb = RGBColor(*color_gray)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Content 1
        cell = row_cells[1]
        set_cell_width(cell, col_widths[1])
        p1 = cell.paragraphs[0]
        p1.clear()
        if tag == 'replace':
            runs1, _ = word_diff_runs(ltext, rtext)
            for text_seg, is_diff, is_placeholder in runs1:
                if is_placeholder:
                    # Missing placeholder: green bold
                    add_colored_run(p1, text_seg, color_placeholder, bold=True)
                elif is_diff:
                    add_colored_run(p1, text_seg, color_left, bold=True)
                else:
                    run = p1.add_run(text_seg)
                    set_run_font(run)
        elif tag == 'delete':
            add_colored_run(p1, ltext, color_left, bold=True)
            # Right side empty for whole row, no marker added
        elif tag == 'insert':
            # Left side empty for whole row, no marker added
            pass
        else:
            run = p1.add_run(ltext)
            set_run_font(run)

        # Marker
        cell = row_cells[2]
        set_cell_width(cell, col_widths[2])
        p_mark = cell.paragraphs[0]
        p_mark.clear()
        p_mark.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if tag == 'replace':
            add_colored_run(p_mark, '≠', color_mark_replace, bold=True)
        elif tag == 'delete':
            add_colored_run(p_mark, '-', color_left, bold=True)
        elif tag == 'insert':
            add_colored_run(p_mark, '+', color_right, bold=True)

        # Content 2
        cell = row_cells[4]
        set_cell_width(cell, col_widths[4])
        p2 = cell.paragraphs[0]
        p2.clear()
        if tag == 'replace':
            _, runs2 = word_diff_runs(ltext, rtext)
            for text_seg, is_diff, is_placeholder in runs2:
                if is_placeholder:
                    # Missing placeholder: green bold
                    add_colored_run(p2, text_seg, color_placeholder, bold=True)
                elif is_diff:
                    add_colored_run(p2, text_seg, color_right, bold=True)
                else:
                    run = p2.add_run(text_seg)
                    set_run_font(run)
        elif tag == 'insert':
            add_colored_run(p2, rtext, color_right, bold=True)
        elif tag == 'delete':
            # Delete row: right side empty, no marker added
            pass
        else:
            run = p2.add_run(rtext)
            set_run_font(run)

    doc.save(output_path)


def main():
    import argparse
    
    # Program header
    print("=" * 70)
    print(f"compare_docs.py - Document Difference Comparison Tool {VERSION}")
    print("Supports PDF/DOCX/PPTX/TXT comparison, generates Word diff report")
    print("Author: Yu Xia  E-mail: yuxiacn@qq.com")
    print("=" * 70)
    print()
    
    parser = argparse.ArgumentParser(
        description='Document Comparison Tool - Compare two documents and generate Word diff report',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Supported formats:
  PDF   - Supports cross-page paragraph merging, page number filtering, visual line number extraction
  DOCX  - Supports paragraph-level comparison, estimated page number
  PPTX  - Slide text comparison
  TXT   - Plain text comparison

Examples:
  python compare_docs.py paper1.pdf paper2.pdf
  python compare_docs.py report1.docx report2.docx --calibrate
  python compare_docs.py doc1.txt doc2.txt --no-merge
        '''
    )
    parser.add_argument('file1', help='First document path')
    parser.add_argument('file2', help='Second document path')
    parser.add_argument('--calibrate', action='store_true', help='Calibration mode: output paragraph location info for debugging')
    parser.add_argument('--no-merge', action='store_true', help='PDF/TXT files: do not merge consecutive lines (compare by original lines)')
    parser.add_argument('--no-page-merge', action='store_true', help='PDF: do not merge paragraphs across pages (process each page independently)')
    
    args = parser.parse_args()
    
    file1 = args.file1
    file2 = args.file2
    use_precise = True
    merge_lines = not args.no_merge  # Default merge, disable with --no-merge
    merge_across_pages = not args.no_page_merge  # Default cross-page merge, disable with --no-page-merge
    
    if not os.path.exists(file1):
        print(f"Error: File does not exist: {file1}")
        sys.exit(1)
    if not os.path.exists(file2):
        print(f"Error: File does not exist: {file2}")
        sys.exit(1)

    name1 = Path(file1).stem
    name2 = Path(file2).stem
    output_name = f"Comparison_{name1}_VS_{name2}.docx"
    output_path = os.path.join(os.getcwd(), output_name)

    print(f"Reading {file1} ...")
    result1 = read_document(file1, use_precise=use_precise, merge_lines=merge_lines, merge_across_pages=merge_across_pages)
    if isinstance(result1, tuple):
        lines1, location_info1 = result1
    else:
        lines1 = result1
        location_info1 = None
    print(f"  Total {len(lines1)} paragraphs")
    
    # Calibration mode: display location info for first few paragraphs
    if args.calibrate and location_info1:
        print("\nCalibration info (first 5 paragraphs):")
        for i in range(min(5, len(lines1))):
            loc = location_info1[i]
            para_num, page_num, line_num = loc[0], loc[1], loc[2] if len(loc) > 2 else None
            if line_num:
                print(f"  paragraph {i+1}: Page{page_num}-L{line_num}")
            else:
                print(f"  paragraph {i+1}: Page{page_num}-{para_num}")
        print()

    print(f"Reading {file2} ...")
    result2 = read_document(file2, use_precise=use_precise, merge_lines=merge_lines, merge_across_pages=merge_across_pages)
    if isinstance(result2, tuple):
        lines2, location_info2 = result2
    else:
        lines2 = result2
        location_info2 = None
    print(f"  Total {len(lines2)} paragraphs")

    print("Analyzing differences ...")
    rows = build_diff_report(lines1, lines2, location_info1, location_info2)

    print("Generating report ...")
    generate_docx(rows, name1, name2, output_path)

    diff_count = len(rows)
    print(f"Diff row count: {diff_count}")
    print(f"Comparison report saved to: {output_path}")


if __name__ == '__main__':
    main()
