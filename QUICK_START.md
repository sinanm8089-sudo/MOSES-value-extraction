# MOSES Value Extraction - Quick Start Guide

## System Requirements

- **Python**: 3.6 or higher
- **Operating System**: Windows, macOS, or Linux
- **Disk Space**: Minimal (< 10 MB)

## Installation (1 minute)

```bash
# Install required Python package
pip install openpyxl
```

## Usage (30 seconds)

```bash
# Basic usage
python moses_extractor.py your_moses_file.txt

# With custom output name
python moses_extractor.py your_moses_file.txt output_report.xlsx
```

## What Gets Extracted

For each damage case in your MOSES file, the software extracts:
- **Damage Case**: NONE, 1PO, 2PO, 3PO, etc.
- **GM Value**: Metacentric height in meters
- **Draft Marks**: AFT(P), AFT(S), FWD(P), FWD(S)

## Output

You'll get a professional Excel file with:
✓ All damage cases in one table
✓ Formatted headers and borders
✓ Proper number formatting
✓ Ready to use in reports

## Example Output

| Damage | GM (m) | AFT(P) (m) | AFT(S) (m) | FWD(P) (m) | FWD(S) (m) |
|--------|--------|------------|------------|------------|------------|
| NONE   | 35.86  | 2.43       | 2.43       | 1.69       | 1.69       |
| 1PO    | 35.44  | 2.43       | 2.43       | 1.70       | 1.69       |
| 2PO    | 33.12  | 2.49       | 2.26       | 1.96       | 1.73       |
| 3PO    | 31.99  | 2.59       | 2.20       | 2.08       | 1.69       |
| ...    | ...    | ...        | ...        | ...        | ...        |

## Troubleshooting

**No data extracted?**
- Verify the file is a MOSES output file
- Check that it contains "STABILITY SUMMARY" sections
- Ensure "DRAFT MARK READINGS" sections are present

**File encoding errors?**
- The script automatically tries multiple encodings (UTF-8, ASCII, Latin-1)
- If issues persist, try re-saving the file in UTF-8 format

**Need help?**
- Check the full README.md for detailed documentation
- Review the example output to verify expected format

## That's It!

Your extracted data is now ready in Excel format.
