# PDF Converter for PowerPoint Presentations

This Python script automatically converts PowerPoint files (both `.ppt` and `.pptx`) to PDF format. It recursively scans a specified directory for PowerPoint files and converts each one into a PDF, then deletes the original file to conserve space.

## Features
- Converts both `.ppt` and `.pptx` files to PDF.
- Works recursively through directories.
- Automatically deletes original PowerPoint files after conversion.

## System Requirements
- Windows OS (7/8/10/11)
- Python 3.x
- Microsoft PowerPoint installed on your machine (any version from 2010 onwards)

## Installation

First, ensure that Python and pip are installed on your system. Then, install the required Python packages:

```bash
pip install pywin32
```

# Usage

To use the script, follow these steps:

1. Download or clone this repository to your local machine.
2. Open your command line interface (CLI).
3. Navigate to the directory where the script is located.
4. Run the script using Python by typing the following command in your CLI:

```bash
python pdfConverter.py
```

# Adjusting Sleep Time
The sleep_time parameter in the function call can be adjusted depending on your system's specifications. Increase the sleep_time if you experience issues with file access or COM operations, particularly on slower systems or systems with heavy I/O operations.

# Troubleshooting

- Permission Issues: Ensure you have the necessary permissions to read/write in the directory and that your PowerPoint application allows scripting.
- File Not Found: Ensure that the path entered is correct and accessible from your script's running environment.
- COM Errors: If you encounter errors related to COM objects, restart your PowerPoint application or your system to clear any lingering processes.

# Example
