# Folder-Details

A Python utility that scans folders and subfolders to analyze document files (PDF, Word, PowerPoint) and generates a comprehensive CSV report with page/slide counts and statistical information.

## 📋 Overview

This tool recursively scans a specified parent folder and all its subfolders, counting pages in PDF files, pages in Word documents, and slides in PowerPoint presentations. It then generates a CSV report containing detailed statistics for each subfolder, including total pages/slides, file count, and median values.

## ✨ Features

- **Multi-format Support**: Analyzes PDF, Word (.doc, .docx), and PowerPoint (.ppt, .pptx) files
- **Recursive Scanning**: Automatically scans all subfolders within the specified directory
- **Statistical Analysis**: Calculates total pages/slides, file count, and median values per subfolder
- **CSV Export**: Generates a structured CSV report for easy analysis
- **Error Handling**: Gracefully handles corrupted or unreadable files with warning messages
- **Windows Integration**: Uses COM automation for accurate Word and PowerPoint page/slide counting

## 🔧 Requirements

### Python Version

- Python 3.x

### Dependencies

- `pypdf` - For reading PDF files
- `pywin32` - For Windows COM automation (Word and PowerPoint)

### System Requirements

- **Windows OS** (required for COM automation)
- Microsoft Word installed (for .doc/.docx files)
- Microsoft PowerPoint installed (for .ppt/.pptx files)

## 📦 Installation

1. Clone or download this repository:

   ```bash
   git clone <repository-url>
   cd Folder-Details
   ```

2. Install required Python packages:
   ```bash
   pip install pypdf pywin32
   ```

## 🚀 Usage

### Basic Usage

1. Open `FolderDetails.py` in a text editor
2. Modify the folder path in the `__main__` section (line 94):
   ```python
   get_folder_details_to_csv(
       r"YOUR_FOLDER_PATH_HERE"
   )
   ```
3. Run the script:
   ```bash
   python FolderDetails.py
   ```

### Example

```python
if __name__ == "__main__":
    get_folder_details_to_csv(
        r"D:\Documents\MyFolder"
    )
```

## 📊 Output

The script generates a CSV file named `Folder_Details.csv` in the same directory as the scanned folder. The CSV contains the following columns:

| Column                 | Description                                                         |
| ---------------------- | ------------------------------------------------------------------- |
| **Subfolder Name**     | Name of each subfolder within the parent directory                  |
| **Total Pages/Slides** | Sum of all pages (PDF/Word) or slides (PowerPoint) in the subfolder |
| **File Count**         | Number of document files found in the subfolder                     |
| **Median**             | Median value of pages/slides across all files in the subfolder      |

### Sample Output

```csv
Subfolder Name,Total Pages/Slides,File Count,Median
Assignment 1,45,3,15.0
Assignment 2,78,5,12.0
Project,120,8,15.0
```

## 🔍 How It Works

1. **Initialization**: Creates COM objects for Word and PowerPoint applications (hidden/background mode)
2. **Folder Scanning**: Iterates through all subfolders in the specified parent directory
3. **File Processing**: For each subfolder:
   - Recursively walks through all files
   - Identifies supported file types by extension
   - Counts pages/slides using appropriate methods:
     - **PDF**: Uses `pypdf.PdfReader` to count pages
     - **Word**: Uses COM automation to open and count pages
     - **PowerPoint**: Uses COM automation to count slides
4. **Statistics Calculation**: Computes total, file count, and median for each subfolder
5. **CSV Generation**: Writes results to `Folder_Details.csv`
6. **Cleanup**: Closes Word and PowerPoint applications

## 📁 Supported File Types

| File Type            | Extensions      | What's Counted |
| -------------------- | --------------- | -------------- |
| PDF                  | `.pdf`          | Pages          |
| Microsoft Word       | `.doc`, `.docx` | Pages          |
| Microsoft PowerPoint | `.ppt`, `.pptx` | Slides         |

## ⚠️ Important Notes

- **Windows Only**: This script requires Windows OS due to COM automation dependencies
- **Microsoft Office Required**: Word and PowerPoint must be installed for those file types to be processed
- **File Access**: The script opens files in read-only mode and does not modify them
- **Performance**: Processing large folders with many files may take some time
- **Error Handling**: Files that cannot be read will be skipped with a warning message, but processing continues
- **PowerPoint Visibility**: PowerPoint application is set to visible (`ppt_app.Visible = True`) - you may see PowerPoint windows briefly during processing

## 🛠️ Code Structure

### Functions

- `get_pdf_page_count(path)`: Counts pages in a PDF file
- `get_word_page_count(word_app, path)`: Counts pages in a Word document using COM
- `get_ppt_slide_count(ppt_app, path)`: Counts slides in a PowerPoint presentation using COM
- `get_folder_details_to_csv(parent_folder)`: Main function that orchestrates the scanning and CSV generation

## 🐛 Troubleshooting

### Common Issues

1. **"ModuleNotFoundError: No module named 'pypdf'"**

   - Solution: Install dependencies with `pip install pypdf pywin32`

2. **"ModuleNotFoundError: No module named 'win32com'"**

   - Solution: Install pywin32: `pip install pywin32`

3. **Word/PowerPoint files not being counted**

   - Ensure Microsoft Office is installed
   - Check that files are not corrupted or password-protected
   - Verify file extensions are correct (.doc, .docx, .ppt, .pptx)

4. **Permission errors**
   - Ensure you have read permissions for the target folder
   - Close any files that are open in Word/PowerPoint before running

## 📝 License

This project is provided as-is for personal and educational use.

## 👤 Author

Created for analyzing folder contents and generating detailed document statistics.

---

**Note**: Remember to update the folder path in the `__main__` section before running the script!
