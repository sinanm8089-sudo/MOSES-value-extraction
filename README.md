# MOSES Value Extraction Software

## Overview
This software extracts damage stability data from MOSES (Marine Operations Simulation) output files and generates professional Excel reports.

## Features
- ✅ Extracts GM (metacentric height) values for all damage cases
- ✅ Extracts draft mark readings (AFT(P), AFT(S), FWD(P), FWD(S))
- ✅ Includes intact condition (Damage = NONE)
- ✅ Processes multiple damage cases automatically
- ✅ Generates formatted Excel reports
- ✅ User-friendly web interface (React)
- ✅ Command-line interface (Python)

## Components

### 1. Python Extractor (`moses_extractor.py`)
Command-line tool for batch processing MOSES files.

**Usage:**
```bash
python moses_extractor.py <input_file> [output_file]
```

**Example:**
```bash
python moses_extractor.py out00001.txt moses_report.xlsx
```

**Requirements:**
- Python 3.6+
- openpyxl library (`pip install openpyxl`)

### 2. Web Interface (`moses_extractor.jsx`)
React-based web application with drag-and-drop file upload. **Fully client-side - no backend required!**

**Features:**
- Upload MOSES output files (.txt, .out)
- Preview extracted data in table format
- Download Excel reports (generated in browser)
- Error handling and validation
- Works completely offline once loaded

**Requirements:**
```bash
npm install xlsx lucide-react
```

## Installation

### Python Version
```bash
# Install required package
pip install openpyxl

# Run the extractor
python moses_extractor.py your_moses_file.txt
```

### Web Interface
```bash
# Install required packages
npm install xlsx lucide-react

# Import the component in your React app
# The component works entirely client-side - no backend needed!
```

The React component can be integrated into any React application. It generates Excel files directly in the browser using the SheetJS library.

## Extracted Data

The software extracts the following information for each damage case:

| Field | Description | Example |
|-------|-------------|---------|
| Damage | Damage case identifier | NONE, 1PO, 2PO, etc. |
| GM (m) | Metacentric height in meters | 35.86 |
| AFT(P) (m) | Aft port draft mark | 2.43 |
| AFT(S) (m) | Aft starboard draft mark | 2.43 |
| FWD(P) (m) | Forward port draft mark | 1.69 |
| FWD(S) (m) | Forward starboard draft mark | 1.69 |

**Note:** MEAN(P) and MEAN(S) draft marks are excluded as per requirements.

## Sample Output

### Console Output
```
Processing MOSES file: complete_moses.txt
Extracted 9 damage cases

Extracted Data:
--------------------------------------------------------------------------------
Damage: NONE       GM:  35.86 m
  Drafts: AFT(P)=2.43, AFT(S)=2.43, FWD(P)=1.69, FWD(S)=1.69
--------------------------------------------------------------------------------
Damage: 1PO        GM:  35.44 m
  Drafts: AFT(P)=2.43, AFT(S)=2.43, FWD(P)=1.70, FWD(S)=1.69
--------------------------------------------------------------------------------
...
Excel file created successfully: MOSES_Complete_Extraction.xlsx
```

### Excel Report
The Excel file includes:
- Professional formatting with headers
- Color-coded header row
- Centered, bordered cells
- Proper number formatting (2 decimal places)
- Auto-sized columns

## How It Works

### Extraction Algorithm
1. **Parse MOSES File**: Reads the entire file line by line
2. **Detect Damage Cases**: Identifies damage case headers and intact conditions
3. **Extract GM Values**: Locates stability summary sections and extracts GM
4. **Extract Draft Marks**: Finds draft mark readings for each case
5. **Associate Data**: Links GM values with corresponding draft marks
6. **Generate Report**: Creates formatted Excel file

### Pattern Recognition
The software uses regular expressions to identify:
- Damage case headers: `DAMAGE STABILITY Case-X (Compartment XXX Flooded)`
- Intact condition: `INTACT TOW CONDITION` or `Damage = NONE`
- GM values: `GM >= X.XX M XX.XX Passes`
- Draft marks: Named values like `AFT(P) X.XX`

## File Format Requirements

### Input Files
- Text-based MOSES output files (.txt, .out)
- Must contain stability summary sections
- Must contain draft mark readings sections

### Expected Sections
1. Draft mark readings table
2. Stability summary with GM values
3. Damage case identifiers (or intact condition)

## Error Handling

The software handles:
- Missing files
- Corrupted data
- Incomplete sections
- Malformed numbers
- Character encoding issues

## Customization

### Modifying Extracted Fields
To extract additional data, modify the `extract_moses_data()` function in `moses_extractor.py`.

### Changing Excel Formatting
Excel styling can be customized in the `create_excel_report()` function:
- Colors (header_fill, font colors)
- Fonts (name, size, style)
- Borders and alignment
- Number formats

### Adding New Draft Marks
To include additional draft marks, update the `draft_names` list in `create_excel_report()`.

## Troubleshooting

### No Data Extracted
- Check if file contains stability summary sections
- Verify file encoding (should be UTF-8 or ASCII)
- Ensure draft mark readings are present

### Incorrect Values
- Verify MOSES file format matches expected structure
- Check for special characters or formatting issues
- Review console output for parsing errors

### Excel File Won't Open
- Ensure openpyxl is installed correctly
- Check file permissions in output directory
- Verify sufficient disk space

## Technical Details

### Dependencies
- **Python**: 3.6 or higher
- **openpyxl**: For Excel file creation and formatting
- **re**: For regular expression pattern matching

### Performance
- Processes typical MOSES files (<5MB) in under 1 second
- Memory efficient - streams file line by line
- No external API calls or network dependencies

## License
This software is provided as-is for processing MOSES marine stability analysis data.

## Support
For issues or questions regarding the extraction software, please review:
1. This README file
2. Sample output files
3. Console error messages

## Version History
- **v1.0**: Initial release with core extraction features
- Supports multiple damage cases
- Generates formatted Excel reports
- Command-line and web interfaces

## Examples

### Example 1: Basic Extraction
```bash
python moses_extractor.py vessel_analysis.txt
```
Output: `moses_extraction.xlsx` in current directory

### Example 2: Custom Output Name
```bash
python moses_extractor.py stability_report.out vessel_stability.xlsx
```
Output: `vessel_stability.xlsx`

### Example 3: Batch Processing
```bash
for file in *.txt; do
    python moses_extractor.py "$file" "${file%.txt}_report.xlsx"
done
```

## Future Enhancements
Potential features for future versions:
- Support for additional MOSES output formats
- Graphical visualization of stability data
- Batch processing multiple files
- Export to additional formats (CSV, PDF)
- Database integration
- Automated report generation

## Contact
For technical questions about MOSES software or marine stability analysis, consult:
- MOSES User Manual
- Marine engineering references
- Stability analysis specialists
