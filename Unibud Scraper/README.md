# UniBud Scraper

This Python script, `Unibud_Scraper.py`, automates the process of scraping question papers from the UniBud website (specifically the VIT Question Bank section) and downloading them as PDF files. It uses Playwright for browser automation and handles login state management to persist sessions.

## Features

- **Automated Login**: The script guides the user through a one-time manual login process and saves the session state for future use.
- **Subject Selection**: Allows the user to specify a subject name to scrape question papers from.
- **Module-wise Downloading**: Iterates through different modules within a selected subject and downloads all available question papers as PDFs.
- **Includes Answers**: Automatically clicks the "Include Answers" option if available.
- **Pagination Handling**: Navigates through multiple pages of questions to ensure all questions are selected before generating the PDF.
- **PDF Generation and Saving**: Generates a PDF of the selected questions and saves it to a `downloads` directory with a sanitized filename.

## Requirements

- Python 3.x
- Playwright library (`playwright`)

## Setup

1.  **Install Playwright**:
    t login. This file is created after the first successful login.

- `downloads/`: A directory where the generated PDF question papers will be saved.

## Configuration

- `URL`: The target URL for the UniBud VIT Question Bank.
- `HEADLESS`: A boolean flag to control whether the browser runs in headless mode (default is `False`, meaning the browser UI will be visible). You can change this to `True` for background operation.
- `DOWNLOAD_DIR`: The directory where PDFs will be saved. Defaults to a `downloads` folder in the current working directory.

## Troubleshooting

- **TimeoutError**: If you encounter `PWTimeoutError`, it might be due to slow internet connection or changes in the UniBud website's structure. You can try increasing the `timeout_ms` values in the `safe_click` function or other `wait_for` calls, or setting `HEADLESS = False` to observe the browser's actions.
- **Login Issues**: If the saved session doesn't work, delete `unibud_state.json` and re-run the script to perform a fresh login.
- **0 question checkboxes found**: This might indicate a change in the website's HTML structure for question checkboxes. The `LOC_Q_CHECKBOXES` CSS selector might need to be updated.