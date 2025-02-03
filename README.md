# Appendix Table Generator for PowerPoint (Python-pptx)

This script extracts text and tables from PowerPoint presentations (`.pptx` files) and generates summary slides containing tables with extracted content. It processes either individual `.pptx` files or all `.pptx` files within a selected folder.

## Features
- Recursively scans slides for text and tables
- Organizes extracted content into structured tables
- Creates summary slides with extracted text
- Supports batch processing of multiple `.pptx` files

## Example Output

![Example Table Slide](https://github.com/SurgeonTalus/appendixTableGeneratorForPowerPoint-PythonPPTX/blob/main/TableExample.png)

## Requirements
- Python 3.x
- Required libraries:
  - `python-pptx`
  - `tkinter`
  - `os`

Install dependencies using:
```bash
pip install python-pptx
```

## Usage
### Running the Script
1. Run the script:
   ```bash
   python script.py
   ```
2. A file dialog will open to select a `.pptx` file or a folder containing `.pptx` files.
3. If a folder is selected, the script will ask whether to process subfolders.
4. Extracted content will be formatted into table slides and saved as a new `.pptx` file with `_Modified` appended to the original filename.

## How It Works
- **Text Extraction:**
  - Scans slides for textboxes, titles, and table contents.
  - Uses recursive traversal to handle grouped shapes.
- **Table Slide Generation:**
  - Divides extracted text into multiple slides if necessary.
  - Formats tables with appropriate headers and font sizes.

## License
This project is licensed under the MIT License.

## Author
Developed by [SurgeonTalus](https://github.com/SurgeonTalus).

