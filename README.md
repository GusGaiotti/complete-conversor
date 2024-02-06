Document Conversion

## Overview
This script is a versatile tool for converting documents between various formats, including DOC, DOCX, XLS, XLSX, PDF, PPT, PPTX, RTF, and HTML. It utilizes Python libraries and Windows COM objects to facilitate these conversions.

## Installation

Ensure Python 3.x is installed, then install required packages:

```bash
pip install pandas pywin32 tabula-py pdfkit PyMuPDF comtypes pdf2docx python-pptx
```

Java Runtime Environment (JRE) is needed for `tabula`.

## Features

- Convert between DOC, DOCX, XLS, XLSX, RTF, HTML, PDF, PPT, and PPTX formats.
- Utilizes Python libraries for handling various file operations.
- Supports conversion to and from PDF for multiple file types.

## Usage

Run the script and follow the prompts to select the file type to convert, specify the file path, choose the target format, and execute the conversion.

## Note

- Requires Windows for `win32com.client` operations.
- Ensure paths and file formats are correctly specified.
- Java is required for PDF table extraction with `tabula`.

This utility simplifies document management tasks, making it easy to switch between formats for various needs.