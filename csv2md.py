#!/usr/bin/env python3
"""
CSV to Markdown converter

This script converts CSV files to Markdown table format.
"""

import argparse
import csv
import sys
from pathlib import Path
import pandas as pd


class CSV2Markdown:
    """Convert CSV files to Markdown tables."""

    def __init__(self, input_path, output_path=None):
        """
        Initialize the converter.

        Args:
            input_path (str): Path to the input CSV file
            output_path (str, optional): Path to the output Markdown file
        """
        self.input_path = Path(input_path)
        self.output_path = Path(output_path) if output_path else None
        
        # Validate input file exists and is a CSV
        if not self.input_path.exists():
            raise FileNotFoundError(f"Input file not found: {self.input_path}")
        
        if self.input_path.suffix.lower() != '.csv':
            raise ValueError(f"Input file must be a CSV file: {self.input_path}")

    def convert(self):
        """
        Convert the CSV file to Markdown format.
        
        Returns:
            str: Markdown table representation of the CSV data
        """
        try:
            # Read CSV file using pandas for better handling of various CSV formats
            df = pd.read_csv(self.input_path)
            
            # Convert DataFrame to markdown table with proper formatting
            markdown = df.to_markdown(index=False, tablefmt="pipe")
            
            # Write to output file if specified
            if self.output_path:
                with open(self.output_path, 'w', encoding='utf-8') as f:
                    f.write(markdown)
            
            return markdown
            
        except Exception as e:
            print(f"Error converting CSV to Markdown: {e}", file=sys.stderr)
            raise


def main():
    """Parse command line arguments and convert CSV to Markdown."""
    parser = argparse.ArgumentParser(description='Convert CSV files to Markdown tables')
    parser.add_argument('input_file', help='Input CSV file')
    parser.add_argument('-o', '--output', help='Output Markdown file')
    args = parser.parse_args()
    
    try:
        converter = CSV2Markdown(args.input_file, args.output)
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
