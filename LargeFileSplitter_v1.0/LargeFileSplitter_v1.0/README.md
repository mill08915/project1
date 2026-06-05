FileSplitter - Quick Guide

What it does

FileSplitter turns a large `.csv` or `.xlsx` file into a more manageable `.xlsx` workbook.

Why use it

- Makes huge files easier to open in Excel
- Converts `.csv` to `.xlsx`
- Splits very large sheets into smaller sheets

What you get

- One output file: `<original_name>_SPLIT.xlsx`
- Multiple sheets if needed: `..._P1`, `..._P2`, etc.
- Same header row on every split sheet

How to use it

1. Put your input file in the same folder as `FileSplitter_.exe`.
2. Run the app: double-click `FileSplitter_.exe`
3. If multiple files are found, type the number for the file you want.
4. For CSV files, accept or change the detected delimiter.
5. Enter a maximum row count per sheet, or press Enter for the default `1,000,000`.
6. Wait for the output file to finish saving.

Important details

- Input types: `.csv`, `.xlsx`
- Output type: `.xlsx`
- Default max rows per sheet: `1,000,000`
- You can enter a smaller max row count at runtime
- If the delimiter is wrong, type the correct one when prompted

If you use the EXE

- The EXE runs without Python installed as it was built with PyInstaller
- Keep the EXE and input file in the same folder

Example

Input: `bigdata.csv`
Output: `bigdata_SPLIT.xlsx`

If the file is large, output sheets may be:
- `bigdata_P1`
- `bigdata_P2`
- `bigdata_P3`

Quick troubleshooting

- No file found: move your `.csv` or `.xlsx` file into the same folder as the program.
- Wrong delimiter: run it again and type the correct delimiter.
- Need smaller sheets: enter a smaller number than `1,000,000` when asked.

Run from source

Install `openpyxl` if needed:

```powershell
pip install openpyxl
```

Then run:

```powershell
python FileSplitter_.py
```
