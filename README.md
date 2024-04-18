# PDF Converter for PowerPoint Presentations(.ppt & .pptx)

This Python script automatically converts PowerPoint files (both `.ppt` and `.pptx`) to PDF format. It recursively scans a specified directory for PowerPoint files and converts each one into a PDF, then copy and store original files into 'original' folder.
## Features
- Converts both `.ppt` and `.pptx` files to PDF.
- Copy PowerPoint files into subfolder.
- Automatically deletes original PowerPoint files.

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
![스크린샷 2024-04-18 190231](https://github.com/Thongangerge/ppt-to-pdf-converter/assets/126161416/acd3e2fa-494a-4cd5-987b-19cf43a59b91)

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

![스크린샷 2024-04-18 190104](https://github.com/Thongangerge/ppt-to-pdf-converter/assets/126161416/e6e262ab-5c7c-4ea6-b2c8-ad2b1d519cbe)

# Adjusting Sleep Time
The sleep_time parameter in the function call can be adjusted depending on your system's specifications. Increase the sleep_time if you experience issues with file access or COM operations, particularly on slower systems or systems with heavy I/O operations.
- example
```python
pdfConverter(sleep_time=5)
```
*sleep_time=10* means that converter will sleep 10 seconds every I/O operations between files. Default value is *sleep_time=2*.

# Troubleshooting
- Permission Issues: Ensure you have the necessary permissions to read/write in the directory and that your PowerPoint application allows scripting.
- File Not Found: Ensure that the path entered is correct and accessible from your script's running environment.
