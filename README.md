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

## Adding New Tools

To add a new tool:
1. Create a new Python script in the `tools/` folder
2. Add any required dependencies to `requirements.txt`
3. Update this README with usage instructions