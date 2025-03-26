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
- MSG files
- EML files
"""

import argparse
import os
import sys
import shutil
from pathlib import Path
import mimetypes
import importlib
from concurrent.futures import ThreadPoolExecutor
from typing import List, Dict, Optional, Tuple

# Import existing converters
# We use try/except to handle potential import errors gracefully
try:
    # Try relative import first
    from csv2md import CSV2Markdown
except ImportError:
    try:
        # Try absolute import with script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        sys.path.insert(0, script_dir)
        from csv2md import CSV2Markdown
    except ImportError:
        CSV2Markdown = None

try:
    # Try relative import first
    from txt2md import TXT2Markdown
except ImportError:
    try:
        # Try absolute import with script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        sys.path.insert(0, script_dir)
        from txt2md import TXT2Markdown
    except ImportError:
        TXT2Markdown = None

try:
    # Try relative import first
    from xlsx2md import XLSX2Markdown
except ImportError:
    try:
        # Try absolute import with script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        sys.path.insert(0, script_dir)
        from xlsx2md import XLSX2Markdown
    except ImportError:
        XLSX2Markdown = None

try:
    # Try relative import first
    from docx2md import DOCX2Markdown
except ImportError:
    try:
        # Try absolute import with script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        sys.path.insert(0, script_dir)
        from docx2md import DOCX2Markdown
    except ImportError:
        DOCX2Markdown = None

try:
    # Try relative import first
    from pdf2md import PDF2Markdown
except ImportError:
    try:
        # Try absolute import with script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        sys.path.insert(0, script_dir)
        from pdf2md import PDF2Markdown
    except ImportError:
        PDF2Markdown = None

try:
    # Try relative import first
    from msg2md import MSG2Markdown
except ImportError:
    try:
        # Try absolute import with script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        sys.path.insert(0, script_dir)
        from msg2md import MSG2Markdown
    except ImportError:
        MSG2Markdown = None

try:
    # Try relative import first
    from eml2md import EML2Markdown
except ImportError:
    try:
        # Try absolute import with script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        sys.path.insert(0, script_dir)
        from eml2md import EML2Markdown
    except ImportError:
        EML2Markdown = None


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
            '.msg': 'msg',
            '.eml': 'eml'
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
                elif mime_type == 'message/rfc822':
                    detected_format = 'eml'
        
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
            elif self.format == 'msg' and MSG2Markdown:
                converter = MSG2Markdown(self.input_path, self.output_path)
            elif self.format == 'eml' and EML2Markdown:
                converter = EML2Markdown(self.input_path, self.output_path)
            else:
                raise NotImplementedError(f"Conversion for {self.format} files is not yet implemented")
            
            # Call the converter's convert method
            return converter.convert()
            
        except Exception as e:
            print(f"Error converting {self.input_path} to Markdown: {e}", file=sys.stderr)
            raise


def get_supported_files(directory: Path, recursive: bool = False) -> List[Path]:
    """
    Get all supported files in a directory.
    
    Args:
        directory (Path): Directory to search
        recursive (bool): Whether to search recursively
        
    Returns:
        List[Path]: List of supported files
    """
    # Define supported extensions
    supported_extensions = {
        '.csv', '.txt', '.xlsx', '.xls', '.xlsm', '.docx', '.pdf', '.msg', '.eml'
    }
    
    files = []
    
    # Use glob pattern based on recursion flag
    pattern = '**/*' if recursive else '*'
    
    # Find all files with supported extensions
    for ext in supported_extensions:
        files.extend(directory.glob(f'{pattern}{ext}'))
    
    return sorted(files)


def process_single_file(input_path: Path, output_path: Optional[Path], sheet: Optional[str]) -> int:
    """
    Process a single file.
    
    Args:
        input_path (Path): Path to input file
        output_path (Optional[Path]): Path to output file
        sheet (Optional[str]): Sheet name for Excel files
        
    Returns:
        int: 0 for success, 1 for failure
    """
    try:
        # Create converter and convert file
        converter = X2Markdown(input_path, output_path, sheet=sheet)
        result = converter.convert()
        
        if output_path is None:
            print(result)
        else:
            print(f"Successfully converted {input_path} to {output_path}")
            
        return 0
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


def process_directory(
    input_dir: Path, 
    output_dir: Optional[Path], 
    recursive: bool = False, 
    sheet: Optional[str] = None,
    max_workers: int = 4
) -> Tuple[int, int, int]:
    """
    Process all supported files in a directory.
    
    Args:
        input_dir (Path): Input directory
        output_dir (Optional[Path]): Output directory
        recursive (bool): Whether to process subdirectories
        sheet (Optional[str]): Sheet name for Excel files
        max_workers (int): Maximum number of worker threads
        
    Returns:
        Tuple[int, int, int]: (success_count, error_count, skipped_count)
    """
    # Get all supported files
    files = get_supported_files(input_dir, recursive)
    
    if not files:
        print(f"No supported files found in {input_dir}")
        return 0, 0, 0
    
    # Set up output directory structure
    if output_dir is None:
        # If no output directory specified, create md_results in the input directory
        output_dir = input_dir / "md_results"
    
    # Create output directory if it doesn't exist
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Create processed directory if it doesn't exist
    processed_dir = input_dir / "processed"
    processed_dir.mkdir(parents=True, exist_ok=True)
    
    success_count = 0
    error_count = 0
    skipped_count = 0
    
    print(f"Found {len(files)} supported files in {input_dir}")
    print(f"Output directory: {output_dir}")
    print(f"Processed files will be moved to: {processed_dir}")
    
    # Process files
    for file_path in files:
        try:
            # Skip files that are already in md_results or processed directories
            if "md_results" in str(file_path) or "processed" in str(file_path):
                print(f"Skipping file in special directory: {file_path}")
                skipped_count += 1
                continue
                
            # Determine output path
            if recursive:
                # Preserve directory structure
                rel_path = file_path.relative_to(input_dir)
                out_path = output_dir / rel_path.with_suffix('.md')
                # Create parent directories if needed
                out_path.parent.mkdir(parents=True, exist_ok=True)
                
                # Determine processed file path (preserving directory structure)
                proc_path = processed_dir / rel_path
                proc_path.parent.mkdir(parents=True, exist_ok=True)
            else:
                # Flat structure
                out_path = output_dir / file_path.with_suffix('.md').name
                proc_path = processed_dir / file_path.name
            
            # Process file
            result = process_single_file(file_path, out_path, sheet)
        
            if result == 0:
                success_count += 1
                
                # Move the original file to the processed directory
                try:
                    shutil.move(str(file_path), str(proc_path))
                    print(f"Moved {file_path} to {proc_path}")
                except Exception as move_error:
                    print(f"Warning: Could not move {file_path} to {proc_path}: {move_error}")
            else:
                error_count += 1
            
        except Exception as e:
            print(f"Error processing {file_path}: {e}", file=sys.stderr)
            error_count += 1
    
    return success_count, error_count, skipped_count


def main():
    """Parse command line arguments and convert file(s) to Markdown."""
    parser = argparse.ArgumentParser(description='Convert various file formats to Markdown')
    parser.add_argument('input_path', help='Input file or directory')
    parser.add_argument('-o', '--output', help='Output Markdown file (for single file conversion)')
    parser.add_argument('-d', '--output-dir', help='Output directory (for directory conversion)')
    parser.add_argument('-s', '--sheet', help='Specific sheet to convert (for Excel files)')
    parser.add_argument('-r', '--recursive', action='store_true', help='Process subdirectories recursively')
    args = parser.parse_args()
    
    input_path = Path(args.input_path)
    
    # Check if input path exists
    if not input_path.exists():
        print(f"Error: Input path does not exist: {input_path}", file=sys.stderr)
        return 1
    
    # If input is a directory, process all files
    if input_path.is_dir():
        output_dir = Path(args.output_dir) if args.output_dir else None
        
        # If output is specified but not a directory, show error
        if args.output and not args.output_dir:
            print("Error: When processing a directory, use -d/--output-dir instead of -o/--output", file=sys.stderr)
            return 1
            
        success, errors, skipped = process_directory(
            input_path, 
            output_dir, 
            args.recursive, 
            args.sheet
        )
        
        print(f"\nSummary: {success} files converted, {errors} errors, {skipped} skipped")
        return 0 if errors == 0 else 1
    else:
        # Process single file
        if args.output_dir:
            # If output directory is specified for a single file
            output_dir = Path(args.output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = output_dir / input_path.with_suffix('.md').name
        elif args.output:
            # If specific output file is specified
            output_path = Path(args.output)
        else:
            # Default behavior: create md_results directory in the parent directory
            parent_dir = input_path.parent
            output_dir = parent_dir / "md_results"
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = output_dir / input_path.with_suffix('.md').name
            
            # Create processed directory
            processed_dir = parent_dir / "processed"
            processed_dir.mkdir(parents=True, exist_ok=True)
        
        # Process the file
        result = process_single_file(input_path, output_path, args.sheet)
        
        # If successful and we're using the default organization, move the original file
        if result == 0 and not args.output and not args.output_dir:
            try:
                proc_path = processed_dir / input_path.name
                shutil.move(str(input_path), str(proc_path))
                print(f"Moved {input_path} to {proc_path}")
            except Exception as move_error:
                print(f"Warning: Could not move {input_path} to {proc_path}: {move_error}")
                
        return result


if __name__ == "__main__":
    sys.exit(main())
