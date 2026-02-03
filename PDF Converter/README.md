# PDF Converter

A Python automation tool that converts Microsoft Word documents and PowerPoint presentations to PDF format using COM automation. The script scans a specified folder, converts supported files to PDF, and optionally removes the original files.

## 📋 Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [How It Works](#how-it-works)
- [Supported File Formats](#supported-file-formats)
- [Troubleshooting](#troubleshooting)
- [Code Structure](#code-structure)
- [Important Notes](#important-notes)

## ✨ Features

- **Batch Conversion**: Automatically converts multiple Word and PowerPoint files in a folder
- **Recursive Scanning**: Searches through subdirectories within the specified folder
- **Automatic Cleanup**: Removes original files after successful conversion (optional)
- **Error Handling**: Gracefully handles conversion errors without stopping the entire process
- **Format Support**: Handles both legacy (.doc, .ppt) and modern (.docx, .pptx) formats
- **Silent Operation**: Runs Word in the background (PowerPoint runs minimized)

## 🔧 Requirements

### Software Requirements

- **Windows Operating System** (required for COM automation)
- **Microsoft Word** (installed and licensed)
- **Microsoft PowerPoint** (installed and licensed)
- **Python 3.x**

### Python Dependencies

- `comtypes` - Python COM support library

## 📦 Installation

1. **Clone or download this repository**

2. **Install Python dependencies**:

   ```bash
   pip install comtypes
   ```

3. **Ensure Microsoft Office is installed**:
   - The script requires Microsoft Word and PowerPoint to be installed on your system
   - Office must be properly licensed and activated

## ⚙️ Configuration

Before running the script, you need to configure the target folder path.

Edit `PDF_Convert.py` and modify the `FOLDER` variable on line 3:

```python
FOLDER = "D:\\Downloads"  # Change this to your desired folder path
```

**Note**: Use double backslashes (`\\`) or raw strings (`r"D:\Downloads"`) for Windows paths.

## 🚀 Usage

1. **Configure the folder path** in `PDF_Convert.py` (see Configuration section)

2. **Run the script**:

   ```bash
   python PDF_Convert.py
   ```

3. **Monitor the output**: The script will print conversion status messages:
   - Success messages: `"Converted Word: <file_path>"` or `"Converted PowerPoint: <file_path>"`
   - Error messages: `"Error converting Word/PowerPoint <file_path>: <error_details>"`
   - Skipped files: `"Skipping unsupported file: <file_path>"`

## 🔍 How It Works

1. **File Discovery**: The script recursively walks through the specified folder and identifies:

   - Word documents (`.doc`, `.docx`)
   - PowerPoint presentations (`.ppt`, `.pptx`)
   - Other files (skipped with a message)

2. **Word Conversion**:

   - Creates a Word COM application instance (hidden)
   - Opens each Word document
   - Saves it as PDF using FileFormat 17 (PDF format)
   - Closes the document
   - Removes the original file

3. **PowerPoint Conversion**:

   - Creates a PowerPoint COM application instance (minimized)
   - Opens each presentation
   - Saves it as PDF using format 32 (PDF format)
   - Closes the presentation
   - Removes the original file

4. **Cleanup**: Both COM applications are properly closed after processing

## 📄 Supported File Formats

### Input Formats

- **Word Documents**: `.doc`, `.docx`
- **PowerPoint Presentations**: `.ppt`, `.pptx`

### Output Format

- **PDF**: `.pdf` (saved in the same location as the original file)

### Files That Are Skipped

- Files already in PDF format (`.pdf`)
- Any other file types not listed above

## 🐛 Troubleshooting

### Common Issues

1. **"ModuleNotFoundError: No module named 'comtypes'"**

   - **Solution**: Install comtypes using `pip install comtypes`

2. **"COM object creation failed" or "Word/PowerPoint not found"**

   - **Solution**: Ensure Microsoft Office is installed and properly configured
   - Try opening Word/PowerPoint manually to verify they work

3. **"Permission denied" errors**

   - **Solution**:
     - Ensure the files are not open in another application
     - Check that you have write permissions in the target folder
     - Run the script with appropriate administrator privileges if needed

4. **Files not being converted**

   - **Solution**:
     - Verify the folder path is correct
     - Check that files have the correct extensions (case-insensitive)
     - Ensure files are not corrupted or password-protected

5. **PowerPoint window appears during conversion**

   - **Note**: This is expected behavior. The script sets PowerPoint to minimized state, but it may briefly appear.

6. **Original files are deleted but PDFs not created**
   - **Solution**: Check for errors in the console output. The script may have encountered an error during save, but the original file was already deleted.

## 📁 Code Structure

```
PDF_Convert.py
├── FOLDER (constant) - Target folder path for scanning
├── convert_docs() - Handles Word document conversion
├── convert_ppts() - Handles PowerPoint presentation conversion
└── main() - Main execution function
    ├── File discovery (recursive walk)
    ├── Word application initialization
    ├── PowerPoint application initialization
    └── Cleanup and error handling
```

### Function Details

- **`convert_docs(word_app, doc_files)`**:

  - Takes a Word COM application object and list of Word file paths
  - Converts each file to PDF and removes the original
  - Handles errors per file without stopping the batch

- **`convert_ppts(ppt_app, ppt_files)`**:

  - Takes a PowerPoint COM application object and list of PPT file paths
  - Converts each file to PDF and removes the original
  - Handles errors per file without stopping the batch

- **`main()`**:
  - Orchestrates the entire conversion process
  - Discovers files, initializes COM applications, and ensures proper cleanup

## ⚠️ Important Notes

1. **File Deletion**: The script **permanently deletes** original files after conversion. Make sure you have backups if needed.

2. **Windows Only**: This script only works on Windows due to COM automation requirements.

3. **Office Installation Required**: Microsoft Word and PowerPoint must be installed and accessible via COM.

4. **File Format Constants**:

   - Word PDF format: `17` (wdFormatPDF)
   - PowerPoint PDF format: `32` (ppSaveAsPDF)

5. **Error Handling**: Errors are logged but don't stop the batch process. Check console output for any issues.

6. **PowerPoint Visibility**: PowerPoint runs with `Visible = True` and `WindowState = 2` (minimized), so it may briefly appear during conversion.

7. **Path Handling**: The script uses the same directory as the source file for the PDF output.

## 🔒 Safety Recommendations

- **Test on a small folder first** before processing large batches
- **Backup important files** before running the script
- **Review the folder path** carefully to avoid processing unintended directories
- **Check conversion results** before relying on automatic deletion
Feel free to submit issues, fork the repository, and create pull requests for any improvements.
