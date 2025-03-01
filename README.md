# Blue-Team-tools# Metadata Extractor

Overview

Metadata Extractor is a Python script designed to extract metadata from various file types and save the results in HTML and JSON formats.

## Supported File Types: 
- [X] Images (.jpg, .jpeg, .png, .tiff)
- [X] PDF Documents (.pdf)
- [X] Audio Files (.mp3, .flac, .wav, .m4a)
- [X] Word Documents (.docx)

## Features


- [X] -Extracts metadata from multiple file formats

- [X] Saves metadata in HTML table format for easy viewing

- [X] Optionally saves metadata in JSON format

- [X] Accepts file paths as command-line arguments or manual input

- [X] Works with drag and drop in terminal (Windows)

# Installation
## Prerequisites
Ensure you have Python 3.x installed. Then install the required dependencies:
  ```bash
  pip install exifread PyPDF2 mutagen python-docx
 ```

# Usage

Run the script using Python:
```bash
python metadataextractor.py
```
Then, enter the file path manually or drag and drop a file into the terminal and press Enter.

Alternatively, you can run the script with a file path as an argument:
```bash
python metadataextractor.py "C:\path\to\file.pdf"
```

# Example Output

## Command:
```bash
python metadataextractor.py
```

## Input:
```bash
Enter the path of the file to analyze: C:\Users\User\Documents\example.docx
```
## Output:
```bash
Metadata Extracted:
author: John Doe
created: 2025-02-26 14:15:00
last_modified_by: John Doe
modified: 2025-02-26 14:51:00
...
Metadata saved to example.docx_metadata.html
Metadata saved to example.docx_metadata.json
```

# Output Files

HTML File: example.docx_metadata.html (Automatically saved in the script directory)

JSON File: example.docx_metadata.json (Optional, prompted after extraction)

# License
This project is released under the MIT License.

# @Author
Developed by Pierluca De Santis.
For feedback or contributions, feel free to open an issue on GitHub.
