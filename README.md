# SAP-Text-File-to-Excel-Importer-VBA
Excel VBA utility to automatically import multiple tab-delimited text files from a selected folder into separate worksheets with alphabetical ordering, duplicate worksheet handling, error recovery, and execution summary.

## Overview

Importing multiple text files into Excel manually can be repetitive and time-consuming.

This Excel VBA utility automates the process by importing every tab-delimited text file from a user-selected folder into separate worksheets within a single workbook.

The tool is generic, reusable, and designed to simplify bulk text file imports.

---

## Features

- Dynamic folder selection (no hardcoded paths)
- Automatic detection of all `.txt` files
- Alphabetical file import
- One worksheet per file
- Worksheet names based on file names
- Invalid worksheet name handling
- Duplicate worksheet detection (Replace / Skip / Cancel)
- Auto-adjust column widths
- Error handling with failed file reporting
- Import summary
- Execution time reporting

---

## Workflow

```text
Select Folder
      │
      ▼
Detect TXT Files
      │
      ▼
Sort Alphabetically
      │
      ▼
Import Files
      │
      ▼
Create Worksheets
      │
      ▼
Display Summary
```

---

## Demo

A short demonstration video is included in the **Demo** folder.

Workflow demonstrated:

1. Select a folder containing text files.
2. Run the VBA macro.
3. Import all files automatically.
4. Review generated worksheets.
5. Display import summary.

---

## Installation

1. Download the repository.
2. Open Microsoft Excel.
3. Press **Alt + F11**.
4. Import **ImportTextFiles.bas**.
5. Run **ImportTextFilesToSheets**.
6. Select the folder containing text files.

---

## Technologies Used

- Microsoft Excel VBA
- QueryTables
- FileDialog

---

## Repository Structure

```text
│
├── README.md
├── LICENSE
├── TextFileExtractorToExcel.bas
│
├── Demo
│   └── Many_Text_File_to_Excel_Workbook_Converter_at_once_Demo.1.mp4
│
├── Sample Input
│   ├── Opex2025.txt
│   ├── Opex2024.txt
│   └── Opex2023.txt
│   └── .
│   └── .
│   └── .... so on
│
├── Sample Output
│   └── Text_file_to_Excel_file _Importing_Template


```
## Future Improvements

- CSV file support
- Automatic delimiter detection
- Progress indicator

---

## License

This project is licensed under the MIT License.
