#!/usr/bin/env python3
"""
PDF Combiner Tool

Combines all PDF files from a specified folder into a single PDF.
Usage: python tools/pdf_combiner.py <folder_name>
"""

import argparse
import sys
from pathlib import Path
from typing import List

try:
    from pypdf import PdfWriter, PdfReader
except ImportError:
    print("Error: pypdf is required. Install it with: pip install pypdf")
    sys.exit(1)


def combine_pdfs(pdf_files: List[Path], output_path: Path) -> bool:
    """
    Combine multiple PDF files into a single PDF.
    
    Args:
        pdf_files: List of paths to PDF files to combine (in order)
        output_path: Path where the combined PDF should be saved
        
    Returns:
        True if successful, False otherwise
    """
    try:
        writer = PdfWriter()
        total_pages = 0
        
        for pdf_file in pdf_files:
            if not pdf_file.exists():
                print(f"  Warning: PDF file not found: {pdf_file.name}")
                continue
            
            try:
                reader = PdfReader(str(pdf_file))
                page_count = len(reader.pages)
                
                for page in reader.pages:
                    writer.add_page(page)
                
                total_pages += page_count
                print(f"  ✓ Added {pdf_file.name} ({page_count} page(s))")
                
            except Exception as e:
                print(f"  ✗ Error reading {pdf_file.name}: {e}")
                continue
        
        # Write combined PDF
        with open(output_path, 'wb') as output_file:
            writer.write(output_file)
        
        print(f"\n  Total pages: {total_pages}")
        return True
        
    except Exception as e:
        print(f"Error combining PDFs: {e}")
        import traceback
        traceback.print_exc()
        return False


def combine_pdf_folder(folder_name: str) -> None:
    """
    Combine all PDF files from pdfs/{folder_name}/ into a single PDF.
    
    Args:
        folder_name: Name of the folder containing PDF files to combine
    """
    # Get the script directory and project root
    script_dir = Path(__file__).parent
    project_root = script_dir.parent
    pdfs_dir = project_root / "pdfs"
    input_folder = pdfs_dir / folder_name
    output_file = pdfs_dir / f"{folder_name}.pdf"
    
    # Check if input folder exists
    if not input_folder.exists():
        print(f"Error: Folder '{input_folder}' does not exist.")
        sys.exit(1)
    
    if not input_folder.is_dir():
        print(f"Error: '{input_folder}' is not a directory.")
        sys.exit(1)
    
    # Find all PDF files in the folder
    pdf_files = sorted(input_folder.glob("*.pdf"))
    
    if not pdf_files:
        print(f"Error: No PDF files found in '{input_folder}'.")
        sys.exit(1)
    
    print(f"Found {len(pdf_files)} PDF file(s) to combine:")
    for pdf_file in pdf_files:
        print(f"  - {pdf_file.name}")
    
    print(f"\n{'='*50}")
    print("Combining PDF files...")
    print(f"{'='*50}\n")
    
    if combine_pdfs(pdf_files, output_file):
        print(f"\n{'='*50}")
        print(f"✓ Success! Combined PDF saved to: {output_file}")
        print(f"{'='*50}")
    else:
        print("\nError: Failed to combine PDFs.")
        sys.exit(1)


def main():
    """Main entry point for the script."""
    parser = argparse.ArgumentParser(
        description="Combine all PDF files from a folder into a single PDF."
    )
    parser.add_argument(
        "folder_name",
        help="Name of the folder in pdfs/ containing PDF files to combine"
    )
    
    args = parser.parse_args()
    combine_pdf_folder(args.folder_name)


if __name__ == "__main__":
    main()

