#!/usr/bin/env python3
"""
MSG to Markdown converter

This script converts Outlook MSG files to Markdown format.
"""

import argparse
import sys
from pathlib import Path
import re

try:
    import extract_msg
except ImportError:
    print("Error: extract_msg package is required for MSG conversion.", file=sys.stderr)
    print("Install it using: pip install extract-msg", file=sys.stderr)
    extract_msg = None


class MSG2Markdown:
    """Convert Outlook MSG files to Markdown format."""

    def __init__(self, input_path, output_path=None):
        """
        Initialize the converter.

        Args:
            input_path (str): Path to the input MSG file
            output_path (str, optional): Path to the output Markdown file
        """
        self.input_path = Path(input_path)
        self.output_path = Path(output_path) if output_path else None
        
        # Validate input file exists and is an MSG file
        if not self.input_path.exists():
            raise FileNotFoundError(f"Input file not found: {self.input_path}")
        
        if self.input_path.suffix.lower() != '.msg':
            raise ValueError(f"Input file must be an MSG file: {self.input_path}")
        
        # Check if extract_msg is available
        if extract_msg is None:
            raise ImportError("extract_msg package is required for MSG conversion.")

    def _format_body(self, body):
        """
        Format the email body as Markdown.
        
        Args:
            body (str): Email body text
            
        Returns:
            str: Formatted Markdown text
        """
        if not body:
            return ""
        
        # Split text into paragraphs (separated by empty lines)
        paragraphs = re.split(r'\n\s*\n', body)
        
        # Process each paragraph
        formatted_paragraphs = []
        for para in paragraphs:
            if para.strip():
                formatted_paragraphs.append(para.strip())
        
        # Join paragraphs with double newlines
        return '\n\n'.join(formatted_paragraphs)

    def convert(self):
        """
        Convert the MSG file to Markdown format.
        
        Returns:
            str: Markdown representation of the MSG file
        """
        try:
            # Open the MSG file
            msg = extract_msg.Message(self.input_path)
            
            # Extract basic information
            subject = msg.subject or "No Subject"
            sender = msg.sender or "Unknown Sender"
            body = msg.body or ""
            
            # Create markdown content
            markdown = f"# {subject}\n\n"
            markdown += f"From: {sender}\n\n"
            markdown += "---\n\n"
            markdown += self._format_body(body)
            
            # Close the MSG file
            msg.close()
            
            # Write to output file if specified
            if self.output_path:
                with open(self.output_path, 'w', encoding='utf-8') as f:
                    f.write(markdown)
            
            return markdown
            
        except Exception as e:
            print(f"Error converting MSG to Markdown: {e}", file=sys.stderr)
            raise


def main():
    """Parse command line arguments and convert MSG to Markdown."""
    parser = argparse.ArgumentParser(description='Convert MSG files to Markdown format')
    parser.add_argument('input_file', help='Input MSG file')
    parser.add_argument('-o', '--output', help='Output Markdown file')
    args = parser.parse_args()
    
    try:
        converter = MSG2Markdown(args.input_file, args.output)
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
