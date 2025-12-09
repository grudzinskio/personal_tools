# Personal Tools

A collection of personal utility tools and scripts for various tasks.

## Repository Structure

```
personal_tools/
├── tools/              # All tool scripts
├── slideshows/         # Folder containing subfolders with PPTX files
│   └── exam2_csc3511/  # Example: folder with PPTX files to merge
└── requirements.txt    # Python dependencies
```

## Installation

1. Clone this repository
2. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Tools

### PPTX to PDF Merger (Recommended)

Converts all PPTX files to PDFs (preserving exact formatting) and merges them into a single PDF.
This method preserves the exact appearance of slides, just like printing to PDF.

**Usage:**
```bash
python tools/pptx_to_pdf_merger.py <folder_name> [--keep-temp]
```

**Example:**
```bash
python tools/pptx_to_pdf_merger.py exam2_csc3511
```

This will:
- Find all `.pptx` files in `slideshows/exam2_csc3511/`
- Convert each PPTX to PDF using PowerPoint (preserves exact formatting)
- Merge all PDFs into a single PDF
- Save the result as `slideshows/exam2_csc3511.pdf`
- Clean up temporary PDF files (use `--keep-temp` to keep them)

**Features:**
- Preserves exact slide formatting and appearance
- Processes files in alphabetical order
- Uses PowerPoint's native PDF export (same as Print to PDF)
- Handles errors gracefully
- Provides progress feedback

**Requirements:**
- Microsoft PowerPoint must be installed (uses COM automation)
- Windows OS (COM automation is Windows-specific)

### PPTX Merger

Merges all PPTX files from a specified folder into a single presentation.

**Usage:**
```bash
python tools/pptx_merger.py <folder_name>
```

**Example:**
```bash
python tools/pptx_merger.py exam2_csc3511
```

This will:
- Find all `.pptx` files in `slideshows/exam2_csc3511/`
- Merge them into a single presentation
- Save the result as `slideshows/exam2_csc3511.pptx`

**Features:**
- Processes files in alphabetical order
- Preserves slide content (text, images, shapes)
- Handles errors gracefully
- Provides progress feedback

**Note:** This method may not preserve all formatting perfectly. For best results, use the PPTX to PDF Merger instead.

## Adding New Tools

To add a new tool:
1. Create a new Python script in the `tools/` folder
2. Add any required dependencies to `requirements.txt`
3. Update this README with usage instructions