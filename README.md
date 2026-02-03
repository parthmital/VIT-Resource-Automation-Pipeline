# VIT Resource Automation Pipeline

An end-to-end academic automation toolkit designed for VIT students to **collect, consolidate, analyze, and standardize course materials** across an entire semester with minimal manual effort.

This project combines four independent but tightly integrated utilities into a single, coherent workflow that replaces weeks of repetitive manual work with a few hours of automated processing.

## Problem Context

At VIT, academic resources are fragmented across multiple platforms:

- **VTOP Course Page**: Materials uploaded separately by dozens of faculty members per subject
- **UniBud**: Previous Year Questions (PYQs) distributed across modules and paginated question banks
- **Faculty Materials**: Provided in mixed formats (PDF, Word, PowerPoint), with no quality ranking

Manually:

- Downloading materials faculty-by-faculty
- Selecting questions module-by-module
- Converting files to a standard format
- Evaluating which faculty has the best notes

…takes **several weeks per semester**.

## What This Project Does

**VTOP Scholar Pipeline** automates the _entire academic resource lifecycle_:

1. **Bulk-downloads all faculty materials for a subject**
2. **Bulk-downloads all PYQs for a subject**
3. **Quantitatively analyzes faculty material depth**
4. **Standardizes everything into PDFs**
5. **Produces a clean, curated folder of the best resources**

No human-in-the-loop. No repetitive clicks. No manual sorting.

## Included Mini Projects

| Tool               | Purpose                                                            |
| ------------------ | ------------------------------------------------------------------ |
| **VTOP Scraper**   | Downloads all faculty-uploaded materials for a subject from VTOP   |
| **UniBud Scraper** | Downloads all PYQs for a subject from UniBud                       |
| **Folder Details** | Analyzes downloaded materials and generates per-faculty statistics |
| **PDF Converter**  | Converts Word and PowerPoint files to PDF                          |

Each tool is usable independently but is designed to form a **sequential pipeline**.

## 1 VTOP Scraper

### Purpose

Automates downloading **all faculty materials** for a selected subject from the **old VTOP course page**.

### Key Capabilities

- Iterates through every faculty listed for a subject
- Downloads all uploaded materials (ZIPs)
- Automatically renames files using:

```
{Faculty Name} {Slot}.zip
```

- Organizes downloads into subject-specific folders

### Important Constraints

- Works **only with the old VTOP course page**
- Requires **manual login**
- Chrome browser automation via Selenium

## 2 UniBud Scraper

### Purpose

Automates downloading **all Previous Year Questions (PYQs)** for a subject from UniBud.

### Key Capabilities

- Subject-based selection
- Iterates through all modules (typically 7 – 8)
- Selects **every question checkbox**
- Enables “Include Answers” when available
- Handles pagination
- Generates a consolidated PDF per module/subject

### Automation Details

- Uses Playwright
- Persists login state across sessions
- Requires one-time manual login

## 3 Folder Details Tool

### Purpose

Quantitatively evaluates faculty materials to identify **which faculties provide the most detailed content**.

### What It Analyzes

- PDF page counts
- Word document page counts
- PowerPoint slide counts

### Output

A CSV file containing **per-faculty statistics**:

| Subfolder (Faculty) | Total Pages/Slides | File Count | Median |
| ------------------- | ------------------ | ---------- | ------ |
| Dr. A Kumar         | 312                | 18         | 16     |
| Prof. B Sharma      | 145                | 9          | 14     |
| Dr. C Rao           | 420                | 22         | 19     |

This allows objective ranking (e.g., top 5 faculties by material depth).

### Output File

```
Folder_Details.csv
```

## 4 PDF Converter

### Purpose

Standardizes all documents into **PDF format**.

### Supported Input

- `.doc`, `.docx`
- `.ppt`, `.pptx`

### Behavior

- Converts files in-place
- Uses Microsoft Office COM automation
- Optionally deletes original files after conversion

## System Requirements

### Operating System

- **Windows only** (COM automation dependency)

### Software

- Python 3.x
- Google Chrome
- Microsoft Word
- Microsoft PowerPoint

### Python Dependencies

- `selenium`
- `playwright`
- `pypdf`
- `pywin32`
- `comtypes`

## Important Notes

- VTOP Scraper supports **only the old course page**, not the consolidated version
- Login is **manual** for both VTOP and UniBud
- COM-based tools require Microsoft Office to be installed and activated
- UniBud and VTOP automation depends on current site structure
- Original files may be deleted during PDF conversion if enabled

## Design Philosophy

- **Automation over interaction**
- **Quantitative filtering over subjective browsing**
- **Standardized outputs**
- **Zero manual repetition**

The pipeline is optimized for **scale**: more subjects do not increase effort.

## Usage Disclaimer

This project is intended for **personal academic use**.  
Users are responsible for complying with institutional policies and platform terms of service.
