# x2md - User Guide

## What is x2md?

x2md is a simple tool that converts various file formats into Markdown - a lightweight markup language that's easy to read, write, and share. With x2md, you can transform your documents, spreadsheets, and PDFs into clean, portable Markdown files.

## Why Use Markdown?

- **Simplicity**: Easy to read and write
- **Portability**: Works across platforms and devices
- **Compatibility**: Supported by many websites and tools
- **Version control friendly**: Easy to track changes
- **Focus on content**: Less distraction from formatting

## Quick Start

### Installation

1. Make sure you have Python 3.8 or newer installed
2. If you've downloaded the source code (which is currently the only way to install):
   ```
   # Navigate to the x2md directory
   cd path/to/x2md
   
   # Install in development mode
   python -m pip install -e .
   ```

Note: The package is not yet available on PyPI, so you cannot install it with `pip install x2md`.

### Basic Usage

Convert any supported file with a single command:

```
x2md your_file.extension
```

The result will be displayed on screen. To save to a file, add `-o output.md`:

```
x2md your_file.extension -o result.md
```

## Supported File Types

### Text Files (.txt)
Plain text files are converted with intelligent heading detection:
```
x2md document.txt -o document.md
```

### CSV Files (.csv)
Spreadsheet data is converted to well-formatted Markdown tables:
```
x2md data.csv -o data.md
```

### Excel Files (.xlsx, .xls)
Excel spreadsheets are converted with all sheets included:
```
x2md spreadsheet.xlsx -o spreadsheet.md
```

To convert only a specific sheet:
```
x2md spreadsheet.xlsx -s "Sheet1" -o spreadsheet.md
```

### Word Documents (.docx)
Word documents are converted with formatting preserved:
```
x2md document.docx -o document.md
```

### PDF Files (.pdf)
PDF files are converted with text and tables extracted:
```
x2md document.pdf -o document.md
```

## Examples

### Converting a Text File
Starting with this text file:
```
Title

This is a paragraph with some content.

Another paragraph with more information.
```

Will produce this Markdown:
```markdown
## Title

This is a paragraph with some content.

Another paragraph with more information.
```

### Converting a Spreadsheet
A CSV or Excel file containing:
```
Name,Age,City
John,34,New York
Maria,28,Chicago
```

Will produce this Markdown table:
```markdown
| Name  | Age | City      |
|-------|-----|-----------|
| John  | 34  | New York  |
| Maria | 28  | Chicago   |
```

## Tips for Best Results

1. **Text files**: Well-structured text with clear paragraph breaks works best
2. **Spreadsheets**: Clean headers and consistent data types improve table formatting
3. **Word documents**: Simple formatting and standard styles convert most accurately
4. **PDFs**: Text-based PDFs (not scanned documents) work best

## Troubleshooting

- If conversion fails, check that your file isn't password-protected
- Make sure you're using a supported file format
- For Excel files, try specifying a sheet name with `-s "Sheet1"`
- For large files, allow more time for processing

## Getting Help

Run with `--help` to see all available options:
```
x2md --help
```

For specific formats, try:
```
txt2md --help
csv2md --help
xlsx2md --help
docx2md --help
pdf2md --help
```

## What's Next?

After converting to Markdown, you can:
- Include the content in websites or documentation
- Share the file with others (it's plain text!)
- Edit in any text editor or Markdown-specific editor
- Convert to other formats using tools like Pandoc
