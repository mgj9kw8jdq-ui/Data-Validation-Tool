# Data Validation & Comparison Tool
A Python-based toolset for validating and comparing field inspection data against system records. The project contains four scripts that automate summary generation and mismatch detection for both PT and RT reports.

## Files Included
- **PT SUMMARY.py** 
  Generates a summary sheet for PT (Penetrant Testing) reports.

- **PT Comparison.py**  
  Compares PT field report summary sheet with master sheet and highlights mismatches.

- **RT SUMMARY.py**  
  Generates a summary sheet for RT (Radiographic Testing) reports.

- **RT Comparison.py**  
  Compares RT field report summary sheet with master sheet and highlights mismatches.

## Features
- Compares two Excel files row-by-row
- Highlights mismatches using color formatting
- Adds remarks for fields where values differ
- Supports large datasets (2000+ rows)
- Creates summary sheets for quick review

## Tech Used
Python, pandas, openpyxl, tqdm

## Notes
- Actual company data is **not included** due to confidentiality.
- You may add dummy Excel files with the same column structure if required.
- Column mappings and file paths can be modified inside each script.
