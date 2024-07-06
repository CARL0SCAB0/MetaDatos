# Extract Metadatos Script

This Python script extracts metadata from DOCX, XLSX, and PDF files located in a specified directory.

## Requirements

- Python 3.7 or higher
- Python libraries:
  - `python-docx`
  - `openpyxl`
  - `PyPDF2`

## Installation

1. Clone the repository to your local machine:

   ```sh
   git clone <repository_url>
   cd <repository_name>

    Install the required libraries using pip:

    sh

    pip install -r requirements.txt

Usage

Run the script by providing the path to the directory containing the files:


```bash
python extract_metadata.py <directory_path>
```
```bash
python extract_metadata.py /path/to/your/directory
```

The script will search for all DOCX, XLSX, and PDF files in the specified directory (and subdirectories) and display their metadata on the screen.
Example Output


```sh
Metadata for /path/to/your/directory/document.docx:
  author: Author
  title: Title
  subject: Subject
  keywords: Keywords
  last_modified_by: Last modified by
  created: Creation date
  modified: Modification date
```
