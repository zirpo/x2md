#!/usr/bin/env python3
"""
EML to Markdown converter

This script converts EML email files to Markdown format.
"""

import argparse
import sys
import email
import re
from pathlib import Path
from email.header import decode_header


class EML2Markdown:
    """Convert EML email files to Markdown format."""

    def __init__(self, input_path, output_path=None):
        """
        Initialize the converter.

        Args:
            input_path (str): Path to the input EML file
            output_path (str, optional): Path to the output Markdown file
        """
        self.input_path = Path(input_path)
        self.output_path = Path(output_path) if output_path else None
        
        # Validate input file exists and is an EML file
        if not self.input_path.exists():
            raise FileNotFoundError(f"Input file not found: {self.input_path}")
        
        if self.input_path.suffix.lower() != '.eml':
            raise ValueError(f"Input file must be an EML file: {self.input_path}")

    def _decode_header_value(self, value):
        """
        Decode email header value.
        
        Args:
            value (str): Header value to decode
            
        Returns:
            str: Decoded header value
        """
        if not value:
            return ""
            
        decoded_parts = []
        for part, encoding in decode_header(value):
            if isinstance(part, bytes):
                if encoding:
                    try:
                        decoded_parts.append(part.decode(encoding))
                    except:
                        decoded_parts.append(part.decode('utf-8', errors='replace'))
                else:
                    decoded_parts.append(part.decode('utf-8', errors='replace'))
            else:
                decoded_parts.append(part)
                
        return ''.join(decoded_parts)

    def _get_email_body(self, msg):
        """
        Extract email body from email message.
        
        Args:
            msg (email.message.Message): Email message
            
        Returns:
            str: Email body text
        """
        body = ""
        
        if msg.is_multipart():
            # Find the text/plain part
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                
                # Skip attachments
                if "attachment" in content_disposition:
                    continue
                
                # Get the body text
                if content_type == "text/plain":
                    try:
                        charset = part.get_content_charset() or 'utf-8'
                        payload = part.get_payload(decode=True)
                        body = payload.decode(charset, errors='replace')
                        break
                    except Exception as e:
                        print(f"Error decoding email body: {e}", file=sys.stderr)
                        continue
        else:
            # Not multipart - get the payload directly
            try:
                charset = msg.get_content_charset() or 'utf-8'
                payload = msg.get_payload(decode=True)
                body = payload.decode(charset, errors='replace')
            except Exception as e:
                print(f"Error decoding email body: {e}", file=sys.stderr)
        
        return body

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
        Convert the EML file to Markdown format.
        
        Returns:
            str: Markdown representation of the EML file
        """
        try:
            # Read the EML file
            with open(self.input_path, 'rb') as f:
                msg = email.message_from_binary_file(f)
            
            # Extract basic information
            subject = self._decode_header_value(msg.get('Subject', 'No Subject'))
            sender = self._decode_header_value(msg.get('From', 'Unknown Sender'))
            body = self._get_email_body(msg)
            
            # Create markdown content
            markdown = f"# {subject}\n\n"
            markdown += f"From: {sender}\n\n"
            markdown += "---\n\n"
            markdown += self._format_body(body)
            
            # Write to output file if specified
            if self.output_path:
                with open(self.output_path, 'w', encoding='utf-8') as f:
                    f.write(markdown)
            
            return markdown
            
        except Exception as e:
            print(f"Error converting EML to Markdown: {e}", file=sys.stderr)
            raise


def main():
    """Parse command line arguments and convert EML to Markdown."""
    parser = argparse.ArgumentParser(description='Convert EML files to Markdown format')
    parser.add_argument('input_file', help='Input EML file')
    parser.add_argument('-o', '--output', help='Output Markdown file')
    args = parser.parse_args()
    
    try:
        converter = EML2Markdown(args.input_file, args.output)
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
