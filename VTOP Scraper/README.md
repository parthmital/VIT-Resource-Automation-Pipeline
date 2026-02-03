# VTOP Scraper

A Python-based web automation tool for downloading course materials from VIT University's VTOP portal. This scraper automates the process of downloading lecture materials (ZIP files) for all faculty members associated with a course, organizing them by subject and faculty name.

## Overview

VTOP Scraper is designed to automate the tedious process of downloading course materials from VIT University's VTOP portal. Instead of manually clicking through each faculty member's page to download materials, this script:

- Automatically iterates through all faculty members for a selected course
- Downloads ZIP files containing lecture materials
- Organizes downloads into subject-specific folders
- Names files with faculty name and slot information for easy identification

## Features

- **Automated Download**: Automatically downloads materials for all faculty members in a course
- **Smart Organization**: Creates subject-specific folders (e.g., "CS2001 Data Structures")
- **Intelligent Naming**: Files are named using faculty name and slot information
- **Duplicate Handling**: Automatically handles duplicate filenames by appending counters
- **Chrome Profile Support**: Can use existing Chrome profiles for seamless authentication
- **Error Handling**: Robust error handling for timeouts and stale element references
- **Interactive Mode**: Supports processing multiple subjects in a single session

## Requirements

### Software Dependencies

- **Python 3.7+**
- **Google Chrome** browser
- **ChromeDriver** (optional - Selenium Manager can handle this automatically)

### Python Packages

- `selenium` - Web automation framework
- `pathlib` - Path manipulation (built-in)
- `glob` - File pattern matching (built-in)
- `re` - Regular expressions (built-in)
- `os` - Operating system interface (built-in)
- `time` - Time-related functions (built-in)

## Installation

1. **Clone or download this repository**

2. **Install Python dependencies**:

   ```bash
   pip install selenium
   ```

3. **Install ChromeDriver** (optional):
   - If you want to use a specific ChromeDriver version, download it from [ChromeDriver Downloads](https://chromedriver.chromium.org/downloads)
   - Ensure it matches your Chrome browser version
   - Alternatively, Selenium 4.6+ includes Selenium Manager which automatically handles ChromeDriver

## Configuration

Before running the script, you may want to configure the following constants in `VTOP_Scraper.py`:

### Download Location

```python
BASE_DOWNLOAD_ROOT = "D:\\Downloads"  # Change to your preferred download directory
```

### Chrome Configuration (Optional)

```python
CHROME_DRIVER_PATH = None  # Path to ChromeDriver executable (e.g., "C:\\chromedriver\\chromedriver.exe")
CHROME_USER_DATA_DIR = None  # Path to Chrome user data directory
CHROME_PROFILE_DIR = None  # Chrome profile directory name
```

### Timeout Settings

```python
PAGE_LOAD_TIMEOUT = 20  # Seconds to wait for page load
DOWNLOAD_TIMEOUT = 120  # Seconds to wait for file download
```

### File Extension

```python
ZIP_EXT = ".zip"  # Expected file extension for downloads
```

## Usage

1. **Run the script**:

   ```bash
   python VTOP_Scraper.py
   ```

2. **Login to VTOP**:
   - The script will open Chrome browser
   - Navigate to VTOP and log in manually
   - Open the desired subject/course page where the faculty table is visible

3. **Start scraping**:
   - Return to the terminal/command prompt
   - Press `Enter` to begin the automated download process

4. **Process multiple subjects** (optional):
   - After processing one subject, you can choose to:
     - Run again on the same subject page (type `y`)
     - Open another subject and process it (press `Enter`)
     - Quit the program (type `q`)

## How It Works

### Workflow

1. **Initialization**:
   - Sets up Chrome WebDriver with configured options
   - Configures download preferences

2. **Subject Detection**:
   - Extracts course code and title from the faculty table
   - Creates a subject-specific folder (e.g., "CS2001 Data Structures")

3. **Faculty Processing**:
   - Iterates through each faculty row in the table
   - For each faculty:
     - Clicks the "View" button to open faculty details
     - Extracts faculty name and slot information
     - Clicks "Download All Materials" button
     - Waits for ZIP file download to complete
     - Moves and renames the file to the subject folder with a descriptive name

4. **File Organization**:
   - Files are named as: `{FacultyName} {Slot}.zip`
   - Example: `Dr. John Doe A1.zip`
   - Duplicate names are handled with counters: `Dr. John Doe A1_1.zip`

### Key Functions

- `sanitize_filename()`: Cleans filenames to remove invalid characters
- `build_subject_download_dir()`: Creates subject-specific download directory
- `setup_driver()`: Configures and initializes Chrome WebDriver
- `wait_for_new_zip()`: Monitors download directory for new ZIP files
- `normalize_slot()`: Processes slot information (handles multiple slots)
- `process_all_faculties()`: Main processing loop for all faculty members

## Important Notes

1. **Authentication**: You must manually log in to VTOP before starting the scraper. The script does not handle authentication automatically.

2. **Browser State**: Make sure you're on the correct subject page (where the faculty table is visible) before pressing Enter.

3. **Download Directory**: Ensure the download directory has sufficient space and write permissions.

4. **Network Stability**: A stable internet connection is required. The script includes timeout handling, but network issues may cause downloads to fail.

5. **Chrome Updates**: If ChromeDriver is not specified, Selenium Manager will handle it automatically. However, if you're using a custom ChromeDriver path, ensure it matches your Chrome version.

6. **Rate Limiting**: Be mindful of VTOP's rate limiting. The script includes delays, but excessive use may result in temporary access restrictions.

7. **File Conflicts**: If a file with the same name already exists, the script will append a counter (e.g., `_1`, `_2`) to avoid overwriting.

## Troubleshooting

### Issue: ChromeDriver not found

**Solution**:

- Install ChromeDriver manually and set `CHROME_DRIVER_PATH`
- Or update Selenium to 4.6+ which includes Selenium Manager

### Issue: Downloads timing out

**Solution**:

- Increase `DOWNLOAD_TIMEOUT` value
- Check your internet connection
- Ensure VTOP is not experiencing high traffic

### Issue: Stale element reference errors

**Solution**: The script already handles this, but if issues persist:

- Increase wait times
- Check if VTOP's page structure has changed

### Issue: Files not being renamed/moved

**Solution**:

- Check file permissions on the download directory
- Ensure the subject folder path is valid
- Verify that ZIP files are actually being downloaded

### Issue: Wrong subject folder name

**Solution**:

- Check if the XPath selectors in `build_subject_download_dir()` match VTOP's current page structure
- VTOP may have updated their UI

### Issue: Button clicks not working

**Solution**:

- The script uses JavaScript clicks as fallback
- If issues persist, VTOP's page structure may have changed
- Check browser console for JavaScript errors
