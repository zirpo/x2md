#!/usr/bin/env python3
"""
TXT to Markdown converter

This script converts plain text files to Markdown format.
"""

import argparse
import re
import sys
from pathlib import Path


class TXT2Markdown:
    """Convert plain text files to Markdown format."""

    def __init__(self, input_path, output_path=None):
        """
        Initialize the converter.

        Args:
            input_path (str): Path to the input text file
            output_path (str, optional): Path to the output Markdown file
        """
        self.input_path = Path(input_path)
        self.output_path = Path(output_path) if output_path else None
        
        # Validate input file exists and is a text file
        if not self.input_path.exists():
            raise FileNotFoundError(f"Input file not found: {self.input_path}")
        
        if self.input_path.suffix.lower() != '.txt':
            raise ValueError(f"Input file must be a text file: {self.input_path}")

    def _format_paragraphs(self, text):
        """
        Format paragraphs properly for markdown.
        
        Args:
            text (str): Input text
            
        Returns:
            str: Formatted text with markdown paragraphs
        """
        # Split text into paragraphs (separated by empty lines)
        paragraphs = re.split(r'\n\s*\n', text)
        
        # Process each paragraph
        formatted_paragraphs = []
        for para in paragraphs:
            if para.strip():
                lines = para.split('\n')
                
                # Classify lines as headings or content
                line_types = []
                for i, line in enumerate(lines):
                    stripped = line.strip()
                    if not stripped:
                        line_types.append('empty')
                        continue
                    
                    # Check if this could be a heading based on position and content
                    # First line in paragraph or preceded by empty line
                    is_isolated = (i == 0 or not lines[i-1].strip()) and \
                                 (i == len(lines)-1 or not lines[i+1].strip() or len(lines[i+1].strip()) > 50)
                    
                    # Typical heading characteristics - short, no ending punctuation (except colon)
                    if (is_isolated and len(stripped) < 50 and
                        (not stripped[-1] in '.!?,;' or stripped[-1] == ':') and
                        not re.search(r'^[0-9]+\.', stripped)):  # Not numbered list
                        line_types.append('heading')
                    # Check for subsection headings - shorter lines with no punctuation
                    elif (len(stripped) < 30 and 
                          not any(c in stripped for c in '.!?,;') and
                          stripped[0].isupper()):
                        line_types.append('subheading')
                    else:
                        line_types.append('content')
                
                # Process lines based on classification
                processed_lines = []
                for i, (line, line_type) in enumerate(zip(lines, line_types)):
                    if line_type == 'empty':
                        processed_lines.append(line)
                    elif line_type == 'heading':
                        # Main heading (level 2)
                        processed_lines.append(f"## {line.strip()}")
                    elif line_type == 'subheading':
                        # Subheading (level 3)
                        processed_lines.append(f"### {line.strip()}")
                    else:
                        # Regular content - preserve as is
                        processed_lines.append(line)
                
                formatted_para = '\n'.join(processed_lines)
                formatted_paragraphs.append(formatted_para)
        
        # Join paragraphs with double newlines
        return '\n\n'.join(formatted_paragraphs)

    def convert(self):
        """
        Convert the text file to Markdown format.
        
        Returns:
            str: Markdown representation of the text
        """
        try:
            # Read text file
            with open(self.input_path, 'r', encoding='utf-8') as f:
                text = f.read()
            
            # Format text as markdown
            markdown = self._format_paragraphs(text)
            
            # Write to output file if specified
            if self.output_path:
                with open(self.output_path, 'w', encoding='utf-8') as f:
                    f.write(markdown)
                print(f"Successfully converted {self.input_path} to {self.output_path}")
            
            return markdown
            
        except Exception as e:
            print(f"Error converting text to Markdown: {e}", file=sys.stderr)
            raise


def main():
    """Parse command line arguments and convert text to Markdown."""
    parser = argparse.ArgumentParser(description='Convert text files to Markdown format')
    parser.add_argument('input_file', help='Input text file')
    parser.add_argument('-o', '--output', help='Output Markdown file')
    args = parser.parse_args()
    
    try:
        converter = TXT2Markdown(args.input_file, args.output)
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
