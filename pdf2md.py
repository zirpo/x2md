#!/usr/bin/env python3
"""
PDF to Markdown converter

This script converts PDF files to Markdown format.
It handles both text content and tables from PDF files.
"""

import argparse
import sys
from pathlib import Path
import re
import PyPDF2
import pdfplumber


class PDF2Markdown:
    """Convert PDF files to Markdown format."""

    def __init__(self, input_path, output_path=None):
        """
        Initialize the converter.

        Args:
            input_path (str): Path to the input PDF file
            output_path (str, optional): Path to the output Markdown file
        """
        self.input_path = Path(input_path)
        self.output_path = Path(output_path) if output_path else None
        
        # Validate input file exists and is a PDF file
        if not self.input_path.exists():
            raise FileNotFoundError(f"Input file not found: {self.input_path}")
        
        if self.input_path.suffix.lower() != '.pdf':
            raise ValueError(f"Input file must be a PDF file: {self.input_path}")

    def _detect_headings(self, text):
        """
        Detect headings in the PDF text.
        
        Args:
            text (str): Text to analyze
            
        Returns:
            str: Text with markdown heading markers
        """
        lines = text.split('\n')
        processed_lines = []
        
        for i, line in enumerate(lines):
            stripped = line.strip()
            
            # Skip empty lines
            if not stripped:
                processed_lines.append(line)
                continue
                
            # Simple heading detection heuristic
            is_isolated = (i == 0 or not lines[i-1].strip()) and \
                          (i == len(lines)-1 or not lines[i+1].strip())
                          
            if (is_isolated and len(stripped) < 50 and 
                (not stripped[-1] in '.!?,;' or stripped[-1] == ':') and 
                stripped[0].isupper()):
                
                # Determine heading level based on length
                if len(stripped) < 20:
                    processed_lines.append(f"## {stripped}")
                else:
                    processed_lines.append(f"### {stripped}")
            else:
                processed_lines.append(line)
                
        return '\n'.join(processed_lines)

    def _extract_text_with_pypdf2(self):
        """
        Extract text from PDF using PyPDF2.
        
        Returns:
            str: Extracted text content
        """
        text_parts = []
        
        with open(self.input_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                text = page.extract_text()
                
                if text:
                    # Add page number as heading
                    text_parts.append(f"## Page {page_num + 1}\n\n{text}")
                else:
                    text_parts.append(f"## Page {page_num + 1}\n\n*No extractable text content*")
        
        return "\n\n".join(text_parts)

    def _extract_tables_with_pdfplumber(self):
        """
        Extract tables from PDF using pdfplumber.
        
        Returns:
            list: List of extracted markdown tables with page info
        """
        table_parts = []
        
        with pdfplumber.open(self.input_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                
                for table_num, table in enumerate(tables):
                    if table:
                        # Convert table to markdown
                        md_rows = []
                        
                        # Header row
                        header = table[0]
                        md_rows.append("| " + " | ".join([str(cell or "") for cell in header]) + " |")
                        
                        # Separator row
                        md_rows.append("| " + " | ".join(["---" for _ in header]) + " |")
                        
                        # Data rows
                        for row in table[1:]:
                            md_rows.append("| " + " | ".join([str(cell or "") for cell in row]) + " |")
                        
                        # Add table to parts with heading
                        table_parts.append(f"### Table {table_num + 1} (Page {page_num + 1})\n\n" + 
                                          "\n".join(md_rows))
        
        return table_parts

    def convert(self):
        """
        Convert the PDF file to Markdown format.
        
        Returns:
            str: Markdown representation of the PDF
        """
        try:
            # Extract text from PDF
            text_content = self._extract_text_with_pypdf2()
            
            # Apply heading detection
            text_content = self._detect_headings(text_content)
            
            # Extract tables from PDF
            table_content = self._extract_tables_with_pdfplumber()
            
            # Combine text and tables
            if table_content:
                markdown = text_content + "\n\n## Tables\n\n" + "\n\n".join(table_content)
            else:
                markdown = text_content
            
            # Write to output file if specified
            if self.output_path:
                with open(self.output_path, 'w', encoding='utf-8') as f:
                    f.write(markdown)
                print(f"Successfully converted {self.input_path} to {self.output_path}")
            
            return markdown
            
        except Exception as e:
            print(f"Error converting PDF to Markdown: {e}", file=sys.stderr)
            raise


def main():
    """Parse command line arguments and convert PDF to Markdown."""
    parser = argparse.ArgumentParser(description='Convert PDF files to Markdown format')
    parser.add_argument('input_file', help='Input PDF file')
    parser.add_argument('-o', '--output', help='Output Markdown file')
    args = parser.parse_args()
    
    try:
        converter = PDF2Markdown(args.input_file, args.output)
        result = converter.convert()
        
        # If no output file specified, print to stdout
        if not args.output:
            print(result)
            
        return 0
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
