#!/usr/bin/env python3
"""
PPTX to PDF Merger Tool

Converts all PPTX files from a specified folder to PDFs (preserving exact formatting)
and merges them into a single PDF file.
Usage: python tools/pptx_to_pdf_merger.py <folder_name>
"""

import argparse
import sys
import tempfile
from pathlib import Path
from typing import List

try:
    import comtypes.client
except ImportError:
    print("Error: comtypes is required. Install it with: pip install comtypes")
    sys.exit(1)

try:
    from pypdf import PdfWriter, PdfReader
except ImportError:
    print("Error: pypdf is required. Install it with: pip install pypdf")
    sys.exit(1)


def pptx_to_pdf(pptx_path: Path, pdf_path: Path) -> bool:
    """
    Convert a PPTX file to PDF using PowerPoint COM automation.
    
    Args:
        pptx_path: Path to the source PPTX file
        pdf_path: Path where the PDF should be saved
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Create PowerPoint application object
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1  # Set to 0 to run in background (may be faster)
        
        # Open the presentation
        presentation = powerpoint.Presentations.Open(str(pptx_path.absolute()))
        
        # Save as PDF
        # Format 32 = ppSaveAsPDF
        presentation.SaveAs(str(pdf_path.absolute()), 32)
        
        # Close presentation
        presentation.Close()
        
        # Quit PowerPoint
        powerpoint.Quit()
        
        return True
        
    except Exception as e:
        print(f"  Error converting {pptx_path.name} to PDF: {e}")
        try:
            powerpoint.Quit()
        except:
            pass
        return False


def merge_pdfs(pdf_files: List[Path], output_path: Path) -> bool:
    """
    Merge multiple PDF files into a single PDF.
    
    Args:
        pdf_files: List of paths to PDF files to merge (in order)
        output_path: Path where the merged PDF should be saved
        
    Returns:
        True if successful, False otherwise
    """
    try:
        writer = PdfWriter()
        
        for pdf_file in pdf_files:
            if not pdf_file.exists():
                print(f"  Warning: PDF file not found: {pdf_file.name}")
                continue
                
            reader = PdfReader(str(pdf_file))
            for page in reader.pages:
                writer.add_page(page)
        
        # Write merged PDF
        with open(output_path, 'wb') as output_file:
            writer.write(output_file)
        
        return True
        
    except Exception as e:
        print(f"Error merging PDFs: {e}")
        return False


def merge_pptx_to_pdf(folder_name: str, keep_temp: bool = False) -> None:
    """
    Convert all PPTX files from slideshows/{folder_name}/ to PDFs and merge them.
    
    Args:
        folder_name: Name of the folder containing PPTX files to merge
        keep_temp: If True, keep temporary PDF files; if False, delete them
    """
    # Get the script directory and project root
    script_dir = Path(__file__).parent
    project_root = script_dir.parent
    slideshows_dir = project_root / "slideshows"
    input_folder = slideshows_dir / folder_name
    output_file = slideshows_dir / f"{folder_name}.pdf"
    
    # Check if input folder exists
    if not input_folder.exists():
        print(f"Error: Folder '{input_folder}' does not exist.")
        sys.exit(1)
    
    if not input_folder.is_dir():
        print(f"Error: '{input_folder}' is not a directory.")
        sys.exit(1)
    
    # Find all PPTX files in the folder
    pptx_files = sorted(input_folder.glob("*.pptx"))
    
    if not pptx_files:
        print(f"Error: No PPTX files found in '{input_folder}'.")
        sys.exit(1)
    
    print(f"Found {len(pptx_files)} PPTX file(s) to convert and merge:")
    for pptx_file in pptx_files:
        print(f"  - {pptx_file.name}")
    
    # Create temporary directory for PDFs
    temp_dir = input_folder / ".temp_pdfs"
    temp_dir.mkdir(exist_ok=True)
    
    pdf_files = []
    
    # Convert each PPTX to PDF
    print(f"\n{'='*50}")
    print("Converting PPTX files to PDF...")
    print(f"{'='*50}")
    
    for i, pptx_file in enumerate(pptx_files, 1):
        print(f"\n[{i}/{len(pptx_files)}] Converting: {pptx_file.name}")
        pdf_path = temp_dir / f"{pptx_file.stem}.pdf"
        
        if pptx_to_pdf(pptx_file, pdf_path):
            if pdf_path.exists():
                pdf_files.append(pdf_path)
                print(f"  ✓ Successfully converted to PDF")
            else:
                print(f"  ✗ PDF file was not created")
        else:
            print(f"  ✗ Conversion failed")
    
    if not pdf_files:
        print("\nError: No PDF files were successfully created.")
        if not keep_temp:
            temp_dir.rmdir()
        sys.exit(1)
    
    # Merge all PDFs
    print(f"\n{'='*50}")
    print(f"Merging {len(pdf_files)} PDF file(s)...")
    print(f"{'='*50}")
    
    if merge_pdfs(pdf_files, output_file):
        print(f"\n✓ Success! Merged PDF saved to: {output_file}")
        
        # Clean up temporary PDFs if requested
        if not keep_temp:
            print("\nCleaning up temporary PDF files...")
            for pdf_file in pdf_files:
                try:
                    pdf_file.unlink()
                except Exception as e:
                    print(f"  Warning: Could not delete {pdf_file.name}: {e}")
            try:
                temp_dir.rmdir()
            except Exception as e:
                print(f"  Warning: Could not remove temp directory: {e}")
        else:
            print(f"\nTemporary PDF files kept in: {temp_dir}")
    else:
        print("\nError: Failed to merge PDFs.")
        sys.exit(1)


def main():
    """Main entry point for the script."""
    parser = argparse.ArgumentParser(
        description="Convert PPTX files to PDFs and merge them into a single PDF."
    )
    parser.add_argument(
        "folder_name",
        help="Name of the folder in slideshows/ containing PPTX files to merge"
    )
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        help="Keep temporary PDF files after merging (default: delete them)"
    )
    
    args = parser.parse_args()
    merge_pptx_to_pdf(args.folder_name, keep_temp=args.keep_temp)


if __name__ == "__main__":
    main()

