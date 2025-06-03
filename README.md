# PPT2PDF

A simple and reliable web application to convert PowerPoint presentations (.ppt/.pptx) directly to PDF files.

## Description

This is a streamlined Flask web application that provides:
- **Multiple File Upload** - Upload single or multiple PPT/PPTX files at once
- **Drag & Drop Interface** - Simple drag & drop or click to upload files
- **Reliable PDF Conversion** - Direct PowerPoint COM automation for high-quality conversion
- **Batch Processing** - Convert multiple files in sequence with detailed progress tracking
- **Real-time Progress** - Monitor conversion progress for each file in your browser
- **Smart Downloads** - Single PDF download or ZIP file for multiple conversions
- **Robust Error Handling** - Comprehensive validation and error reporting

The application uses PowerPoint's built-in PDF export functionality through COM automation for maximum compatibility and reliability.

## Project Structure

```
PPT2PDF/
├── app.py                # Flask web application
├── simple_converter.py   # PPT to PDF converter using COM automation
├── requirements.txt      # Python dependencies
├── README.md             # Documentation
├── templates/            # HTML templates
│   ├── base.html         # Base template
│   ├── index.html        # Upload page
│   └── progress.html     # Progress page
├── static/               # Static files (CSS, JS)
├── uploads/              # Temporary upload directory
└── downloads/            # Generated PDF files
```

## Requirements

### System Requirements
- Windows operating system (tested on Windows 10)
- Microsoft PowerPoint installed

### Python Requirements
- Python 3.x
- PyWin32 library (for PowerPoint automation)
- Flask (for web application)
- Werkzeug (for file handling)

## Installation

1. Make sure you have Python 3 installed on your system. You can download it from [python.org](https://www.python.org/downloads/)

2. Install the required libraries:
   ```
   pip install -r requirements.txt
   ```

   Or install manually:
   ```
   pip install pywin32 flask werkzeug
   ```

## Quick Start (Web Application)

1. Install dependencies: `pip install -r requirements.txt`
2. Run the web app: `python app.py`
3. Open your browser to: `http://localhost:5000`
4. Upload PowerPoint files (single or multiple) and convert them to PDF!

## Usage

1. Start the web server:
   ```
   python app.py
   ```

2. Open your web browser and go to:
   ```
   http://localhost:5000
   ```

3. Use the web interface to:
   - Upload your PowerPoint files (.ppt or .pptx) - single or multiple
   - Monitor batch conversion progress in real-time
   - Download converted PDF files (single file or ZIP for multiple)
   - The web app handles file cleanup automatically

## Examples

### Single File Conversion
1. Upload `presentation.pptx` through the web interface
2. Monitor the conversion progress in real-time
3. Download `presentation.pdf` when conversion is complete

### Multiple File Conversion
1. Upload multiple files: `presentation1.pptx`, `presentation2.ppt`, `presentation3.pptx`
2. Monitor the batch conversion progress with individual file status
3. Download a ZIP file containing all converted PDFs: `converted_pdfs_[timestamp].zip`

The PDFs will contain all slides with original formatting and quality preserved.

## Features

### Reliability & Error Handling
- **COM Automation**: Uses reliable PowerPoint COM automation for direct conversion
- **Comprehensive Validation**: Validates files before conversion to prevent errors
- **Detailed Error Reporting**: Provides clear error messages for troubleshooting
- **PowerPoint Availability Check**: Verifies PowerPoint is installed and accessible at startup
- **Robust Cleanup**: Automatically cleans up temporary files even if conversion fails
- **Multiple Retry Attempts**: Tries different opening methods if initial attempt fails

### User Experience
- **Multiple File Upload**: Upload single or multiple files at once
- **Real-time Progress**: Live updates during conversion process with individual file status
- **File Size Limits**: Supports files up to 50MB each (configurable)
- **Multiple Formats**: Supports both .ppt and .pptx files
- **Drag & Drop**: Easy file upload interface for single or multiple files
- **Smart Downloads**: Single PDF download or ZIP file for batch conversions
- **Batch Processing**: Detailed progress tracking for each file in the batch

## Troubleshooting

- **Conversion fails**: Ensure that PowerPoint is properly installed and licensed
- **Web app not starting**: Make sure Flask is installed and no other application is using port 5000
- **Upload fails**: Check that the file is a valid PPT/PPTX file and under 50MB in size
- **PowerPoint errors**: Try:
  - Run the application as administrator
  - Close any running PowerPoint instances before conversion
  - Restart your computer if issues persist
- **PowerPoint window appears**: This is normal - PowerPoint may briefly appear during conversion
- **Conversion takes time**: Large presentations may take several minutes to process

## Credits

This application is based on PPT2PDF by ern (www.readern.com), simplified to provide direct PowerPoint to PDF conversion through a clean web interface.
