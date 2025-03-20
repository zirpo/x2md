#!/usr/bin/env python3
"""
X to Markdown converter

This script converts various file formats to Markdown.
Currently supports:
- CSV files
- TXT files
- XLSX/XLS files
- DOCX files
- PDF files

Future support planned for:
- MSG files
"""

import argparse
import os
import sys
from pathlib import Path
import mimetypes
import importlib

# Import existing converters
# We use try/except to handle potential import errors gracefully
try:
    from csv2md import CSV2Markdown
except ImportError:
    CSV2Markdown = None

try:
    from txt2md import TXT2Markdown
except ImportError:
    TXT2Markdown = None

try:
    from xlsx2md import XLSX2Markdown
except ImportError:
    XLSX2Markdown = None

try:
    from docx2md import DOCX2Markdown
except ImportError:
    DOCX2Markdown = None

try:
    from pdf2md import PDF2Markdown
except ImportError:
    PDF2Markdown = None


class FormatDetector:
    """Detect file format and return appropriate converter class."""

    @staticmethod
    def detect_format(file_path):
        """
        Detect the format of a file based on extension and/or content.
        
        Args:
            file_path (str): Path to the input file
            
        Returns:
            str: Detected format (csv, txt, xlsx, docx, pdf, msg)
        """
        file_path = Path(file_path)
        extension = file_path.suffix.lower()
        
        # Simple extension-based detection
        format_map = {
            '.csv': 'csv',
            '.txt': 'txt',
            '.xlsx': 'xlsx',
            '.xls': 'xlsx',
            '.xlsm': 'xlsx',
            '.docx': 'docx',
            '.pdf': 'pdf',
            '.msg': 'msg'
        }
        
        detected_format = format_map.get(extension)
        
        if not detected_format:
            # Fallback to mime type detection
            mime_type, _ = mimetypes.guess_type(file_path)
            
            if mime_type:
                if mime_type.startswith('text/'):
                    detected_format = 'txt'
                elif mime_type == 'application/vnd.ms-excel' or mime_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
                    detected_format = 'xlsx'
                elif mime_type == 'application/msword' or mime_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
                    detected_format = 'docx'
                elif mime_type == 'application/pdf':
                    detected_format = 'pdf'
        
        return detected_format


class X2Markdown:
    """Convert various file formats to Markdown."""

    def __init__(self, input_path, output_path=None, **kwargs):
        """
        Initialize the converter.

        Args:
            input_path (str): Path to the input file
            output_path (str, optional): Path to the output Markdown file
            **kwargs: Additional format-specific options
        """
        self.input_path = Path(input_path)
        self.output_path = Path(output_path) if output_path else None
        self.options = kwargs
        self.format = FormatDetector.detect_format(input_path)
        
        # Validate input file exists
        if not self.input_path.exists():
            raise FileNotFoundError(f"Input file not found: {self.input_path}")
        
        # Validate format detection
        if not self.format:
            raise ValueError(f"Could not detect format for file: {self.input_path}")

    def convert(self):
        """
        Convert the file to Markdown format.
        
        Returns:
            str: Markdown representation of the file
        """
        try:
            # Select appropriate converter based on format
            if self.format == 'csv' and CSV2Markdown:
                converter = CSV2Markdown(self.input_path, self.output_path)
            elif self.format == 'txt' and TXT2Markdown:
                converter = TXT2Markdown(self.input_path, self.output_path)
            elif self.format == 'xlsx' and XLSX2Markdown:
                sheet_name = self.options.get('sheet')
                converter = XLSX2Markdown(self.input_path, self.output_path, sheet_name)
            elif self.format == 'docx' and DOCX2Markdown:
                converter = DOCX2Markdown(self.input_path, self.output_path)
            elif self.format == 'pdf' and PDF2Markdown:
                converter = PDF2Markdown(self.input_path, self.output_path)
            else:
                raise NotImplementedError(f"Conversion for {self.format} files is not yet implemented")
            
            # Call the converter's convert method
            return converter.convert()
            
        except Exception as e:
            print(f"Error converting {self.input_path} to Markdown: {e}", file=sys.stderr)
            raise


def main():
    """Parse command line arguments and convert file to Markdown."""
    parser = argparse.ArgumentParser(description='Convert various file formats to Markdown')
    parser.add_argument('input_file', help='Input file')
    parser.add_argument('-o', '--output', help='Output Markdown file')
    parser.add_argument('-s', '--sheet', help='Specific sheet to convert (for Excel files)')
    args = parser.parse_args()
    
    try:
        # Automatically determine output file name if not specified
        if not args.output:
            input_path = Path(args.input_file)
            output_path = input_path.with_suffix('.md')
        else:
            output_path = args.output
            
        # Create converter and convert file
        converter = X2Markdown(args.input_file, output_path, sheet=args.sheet)
        result = converter.convert()
        
        if not args.output:
            print(result)
            
        return 0
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
