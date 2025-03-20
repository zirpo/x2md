# x2md - Universal File to Markdown Converter

Convert various file formats to Markdown easily.

## Supported Formats

- **TXT files** - Plain text files converted with heading detection
- **CSV files** - Tabular data converted to Markdown tables
- **XLSX/XLS files** - Excel spreadsheets with multi-sheet support
- **DOCX files** - Microsoft Word documents with formatting preservation
- **PDF files** - PDF documents with table extraction support

## Future Planned Support

- MSG (Outlook) files

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/x2md.git
   cd x2md
   ```

2. Create and activate a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Unified Converter

The main `x2md.py` script detects file type and converts accordingly:

```bash
# Basic usage - single file
python x2md.py test/input/test_document.txt

# Specify output file
python x2md.py test/input/test_data.csv -o test/output/result.md

# Excel specific: convert a particular sheet
python x2md.py test/input/test_data.xlsx -s "Sales" -o test/output/result_excel.md

# Process all files in a directory (with automatic organization)
python x2md.py test/input
# This will create test/input/md_results for output and move originals to test/input/processed

# Process all files in a directory with a custom output location
python x2md.py test/input -d test/output

# Process all files in a directory and its subdirectories
python x2md.py test/input -r
```

### Format-Specific Converters

Each format has its own dedicated converter that can be used directly:

```bash
# TXT to Markdown
python txt2md.py input.txt -o output.md

# CSV to Markdown
python csv2md.py input.csv -o output.md

# XLSX to Markdown
python xlsx2md.py input.xlsx -o output.md
python xlsx2md.py input.xlsx -s "Sheet1" -o output.md  # Specific sheet

# DOCX to Markdown
python docx2md.py input.docx -o output.md

# PDF to Markdown
python pdf2md.py input.pdf -o output.md
```

## Features

### TXT Converter
- Intelligently detects and formats headings
- Preserves paragraph structure
- Recognizes subsections and lists

### CSV Converter
- Converts CSV data to properly formatted Markdown tables
- Preserves column alignment (left for text, right for numbers)

### XLSX Converter
- Supports multi-sheet Excel files
- Option to convert specific sheets
- Preserves table formatting

### DOCX Converter
- Preserves document structure including headings
- Handles text formatting (bold, italic)
- Converts tables to Markdown format
- Processes hyperlinks

### PDF Converter
- Extracts text content by page
- Detects and converts tables
- Intelligent heading detection
- Handles multi-page documents

## Dependencies

- pandas, tabulate: For table handling (CSV and Excel)
- openpyxl: For Excel file processing
- python-docx: For Word document processing
- PyPDF2, pdfplumber: For PDF processing

## License

MIT License
