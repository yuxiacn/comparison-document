# Document Comparison Tool

A Python tool for comparing documents in PDF, DOCX, PPTX, and TXT formats, generating Word format difference reports.

**Current Version: V2.0 Build20260403.1**

## Features

- **Multi-format support**: PDF, DOCX, PPTX, TXT
- **PDF intelligent processing**: Cross-page paragraph merging, page number filtering, visual line number extraction
- **Landscape layout**: Display only difference rows
- **Word-level highlighting**: Blue for deletions, Red for additions
- **Sentence-level missing detection**: Show `[Missing Sentence]` (green)
- **Auto-merge consecutive placeholders**

## Version Format

Format: `V{major}.{minor} Build{YYYYMMDD}.{revision}`

- **Major**: Increment on major updates (e.g., V2.0 → V3.0)
- **Minor**: Increment on feature changes (e.g., V2.0 → V2.1)
- **Revision**: Increment on bug fixes (e.g., V2.0 Build20260403.1 → V2.0 Build20260403.2)

### Auto Version Update

```bash
# Update date and revision only
python bump_version.py

# Update minor version (e.g., V2.0 → V2.1)
python bump_version.py minor
```

## Installation

### Required Libraries

```bash
pip install python-docx pdfplumber python-pptx
```

**Required Python Libraries:**
- `python-docx` - For reading and writing Word documents
- `pdfplumber` - For extracting text from PDF files
- `python-pptx` - For reading PowerPoint presentations
- `difflib` - Built-in library for sequence comparison
- `re` - Built-in regular expression library
- `datetime` - Built-in date/time library
- `pathlib` - Built-in path manipulation library
- `argparse` - Built-in command-line argument parsing
- `sys`, `os` - Built-in system libraries

## Usage

```bash
python comparedocen.py <file1> <file2>
```

Output: `Comparison_File1_VS_File2.docx`

### Examples

```bash
# Compare PDF files
python comparedocen.py paper1.pdf paper2.pdf

# Compare Word documents
python comparedocen.py report1.docx report2.docx

# Calibration mode (debug paragraph locations)
python comparedocen.py doc1.pdf doc2.pdf --calibrate

# Do not merge consecutive lines (PDF/TXT)
python comparedocen.py doc1.txt doc2.txt --no-merge

# Do not merge paragraphs across pages (PDF)
python comparedocen.py doc1.pdf doc2.pdf --no-page-merge
```

## Comparison Report Format

The generated Word report includes:
- **Document Comparison Report** (title)
- **Comparison files**: File1 VS File2
- **Generation time**
- **Legend**: Explanation of symbols
- **Difference table**:
  - Location (P{page}-L{line} or P{page}-{paragraph})
  - Content from File 1
  - Mark (=, ≠, -, +)
  - Location
  - Content from File 2

## Contact

**For Chinese version, please contact the author via email:**

**Author**: Yu Xia  
**Email**: yuxiacn@qq.com

## License

This tool is provided as-is for document comparison purposes.
