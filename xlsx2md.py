#!/usr/bin/env python3
"""
XLSX to Markdown converter

This script converts Excel files to Markdown table format.
"""

import argparse
import sys
from pathlib import Path
import pandas as pd


class XLSX2Markdown:
    """Convert Excel files to Markdown tables."""

    def __init__(self, input_path, output_path=None, sheet_name=None):
        """
        Initialize the converter.

        Args:
            input_path (str): Path to the input Excel file
            output_path (str, optional): Path to the output Markdown file
            sheet_name (str, optional): Specific sheet to convert (converts all sheets if not specified)
        """
        self.input_path = Path(input_path)
        self.output_path = Path(output_path) if output_path else None
        self.sheet_name = sheet_name
        
        # Validate input file exists and is an Excel file
        if not self.input_path.exists():
            raise FileNotFoundError(f"Input file not found: {self.input_path}")
        
        valid_extensions = ['.xlsx', '.xls', '.xlsm']
        if self.input_path.suffix.lower() not in valid_extensions:
            raise ValueError(f"Input file must be an Excel file: {self.input_path}")

    def convert(self):
        """
        Convert the Excel file to Markdown format.
        
        Returns:
            str: Markdown table representation of the Excel data
        """
        try:
            # Read the Excel file using pandas
            if self.sheet_name:
                # Read specific sheet
                df_dict = {self.sheet_name: pd.read_excel(self.input_path, sheet_name=self.sheet_name)}
            else:
                # Read all sheets
                df_dict = pd.read_excel(self.input_path, sheet_name=None)
            
            # Convert each sheet to markdown
            markdown_parts = []
            for sheet_name, df in df_dict.items():
                # Add sheet name as header
                markdown_parts.append(f"## Sheet: {sheet_name}\n")
                
                # Convert DataFrame to markdown table
                if not df.empty:
                    table_md = df.to_markdown(index=False, tablefmt="pipe")
                    markdown_parts.append(table_md)
                else:
                    markdown_parts.append("*Empty sheet*")
                
                # Add separator between sheets
                markdown_parts.append("\n")
            
            # Combine all parts
            markdown = "\n".join(markdown_parts)
            
            # Write to output file if specified
            if self.output_path:
                with open(self.output_path, 'w', encoding='utf-8') as f:
                    f.write(markdown)
            
            return markdown
            
        except Exception as e:
            print(f"Error converting Excel to Markdown: {e}", file=sys.stderr)
            raise


def main():
    """Parse command line arguments and convert Excel to Markdown."""
    parser = argparse.ArgumentParser(description='Convert Excel files to Markdown tables')
    parser.add_argument('input_file', help='Input Excel file')
    parser.add_argument('-o', '--output', help='Output Markdown file')
    parser.add_argument('-s', '--sheet', help='Specific sheet to convert')
    args = parser.parse_args()
    
    try:
        converter = XLSX2Markdown(args.input_file, args.output, args.sheet)
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
