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

First, ensure that Python and pip are installed on your system.
Then, Open your command line interface (CLI) and install the converter.

```bash
pip install ppt-to-pdf-converter==0.0.3
```
![스크린샷 2024-04-18 184334](https://github.com/Thongangerge/ppt-to-pdf-converter/assets/126161416/fa2c5cd4-2237-4cb5-bd3b-8c0abcf1515f)

# Usage

To use the script, follow these steps:

1. Run your python
```bash
python
```
2. import the package
```python
from jhconverter.converter import pdfConverter
```
![스크린샷 2024-04-18 184412](https://github.com/Thongangerge/ppt-to-pdf-converter/assets/126161416/6f5bfc26-fb6f-4386-afc4-cdab2e3bf589)

3. Use pdfConverter and enter ppt folder directory
```python
pdfConverter()
```
![스크린샷 2024-04-18 184436](https://github.com/Thongangerge/ppt-to-pdf-converter/assets/126161416/de8e5e5f-6aed-4b64-8b26-9f1012555702)


4. Wait until the process is finished. 
![스크린샷 2024-04-18 185728](https://github.com/Thongangerge/ppt-to-pdf-converter/assets/126161416/b5ce5513-e97b-496b-9084-75e7f9a2f101)



# Adjusting Sleep Time
The sleep_time parameter in the function call can be adjusted depending on your system's specifications. Increase the sleep_time if you experience issues with file access or COM operations, particularly on slower systems or systems with heavy I/O operations.

# Troubleshooting

- Permission Issues: Ensure you have the necessary permissions to read/write in the directory and that your PowerPoint application allows scripting.
- File Not Found: Ensure that the path entered is correct and accessible from your script's running environment.
