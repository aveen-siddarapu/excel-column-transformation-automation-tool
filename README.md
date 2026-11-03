# Excel Column Transformation & Automation Tool

## Overview
This project replicates an Excel VBA macro using Python and pandas.It dynamically rearranges and standardizes columns from a source Excel file
based on a mapping Excel file.

## Features
- Removes exact duplicate rows
- Dynamically arranges columns using the mapping file
- Handles missing columns gracefully
- Supports a configurable source file path
- Saves output in the same folder as the source file
- Automatically generates the output filename

## Project Structure
- `main.py` — Core transformation logic
- `config.py` — Configurable paths and constants
- `mapping/column_mapping.xlsx` — Column mapping definition
- `requirements.txt` — Python dependencies

## Requirements
-  Pands
-  NumPy
-  openpyxl

## How to Run
1. Update `Source_file_path` in `config.py`.
2. Ensure `mapping/column_mapping.xlsx` is present.
3. Install the dependencies listed in `requirements.txt`.
4. Run the script:
   ```bash
   python main.py
   ```

## Output
- The output file is created in the same folder as the source file.
- File name format:
  `<sourcefilename>-transformed.xlsx`
