# MOSES Value Extraction Software - Enhanced Version

## Overview
Professional software for extracting damage stability data from MOSES (Marine Operations Simulation) output files and generating formatted Excel reports matching **Table 10.4.2 Damage Stability Results** format.

## Features
✅ Extracts all damage cases (Intact + all damage scenarios)  
✅ Draft marks: Aft Port, Aft Stbd, Fwd Port, Fwd Stbd  
✅ Heel angle (degrees)  
✅ Trim angle (degrees)  
✅ Wind Area Ratio (Actual & Required)  
✅ Pass/Fail remarks  
✅ Professional Excel formatting matching standard marine reports  
✅ Automatic case numbering (Intact, 1-8, etc.)  

## Output Format

The software generates Excel reports in the standard **Table 10.4.2** format:

| Case No. | Draft Mark (m) | | | | Heel (deg) | Trim (deg) | Wind Area Ratio | | Remarks |
|----------|---------|---------|---------|---------|------|------|--------|----------|---------|
| | Aft Port | Aft Stbd | Fwd Port | Fwd Stbd | | | Actual | Required | |
| Intact | 2.43 | 2.43 | 1.69 | 1.69 | -0.01 | 0.53 | 163.19 | 1.4 | Pass |
| 1 | 2.00 | 1.94 | 1.50 | 1.45 | 0.14 | 0.47 | 141.36 | 1.0 | Pass |
| 2 | 2.52 | 1.98 | 1.65 | 1.11 | 1.28 | 0.82 | 127.60 | 1.0 | Pass |
| ... | ... | ... | ... | ... | ... | ... | ... | ... | ... |

## Installation

```bash
# Install required package
pip install openpyxl
```

## Usage

### Command Line

```bash
# Basic usage
python moses_extractor_v2.py input_file.txt

# With custom output name
python moses_extractor_v2.py input_file.txt custom_report.xlsx
```

### Example

```bash
python moses_extractor_v2.py D7519_TOW_ANALYSIS.txt Damage_Stability_Results.xlsx
```

## Extracted Data Fields

### 1. Case Number
- **Intact**: Undamaged condition
- **1-8**: Damage case numbers (extracted from compartment flooding scenarios)

### 2. Draft Marks (meters)
- **Aft Port**: AFT(P)
- **Aft Stbd**: AFT(S)
- **Fwd Port**: FWD(P)
- **Fwd Stbd**: FWD(S)

*Note: MEAN(P) and MEAN(S) are excluded as per standard practice*

### 3. Heel (degrees)
Roll angle from equilibrium position

### 4. Trim (degrees)
Pitch angle from equilibrium position

### 5. Wind Area Ratio
- **Actual**: Calculated ratio from stability analysis
- **Required**: Minimum required ratio for approval

### 6. Remarks
Pass/Fail status based on stability criteria

## File Requirements

### Input File Format
MOSES output files must contain:
1. **Draft Mark Readings** sections
2. **Stability Summary** sections with:
   - Roll and Pitch values
   - Area Ratio data
3. **Damage case identifiers** or intact condition markers

### Expected Sections

```
+++ D R A F T   M A R K   R E A D I N G S +++
AFT(P)  2.43  AFT(S)  2.43  FWD(P)  1.69  FWD(S)  1.69  ...

+++ S T A B I L I T Y   S U M M A R Y +++
Roll    = 0.14 Deg
Pitch   = 0.47 Deg
Area Ratio  >= 1.00    141.36 Passes
```

## Excel Output Features

### Professional Formatting
- **Title row**: "Table 10.4.2 Damage Stability Results"
- **Color-coded headers**: Gray background, bold text
- **Merged cells**: Multi-column headers for clarity
- **Borders**: All cells bordered for readability
- **Number formatting**: 2 decimal places for all numeric values
- **Auto-sized columns**: Optimized for printing

### Quality Assurance
- Zero formula errors
- Consistent formatting throughout
- Standard marine engineering conventions
- Ready for inclusion in reports

## Technical Details

### Extraction Algorithm

1. **Parse MOSES File**
   - Read file line by line
   - Identify section markers
   
2. **Extract Draft Marks**
   - Locate draft mark reading tables
   - Parse AFT(P), AFT(S), FWD(P), FWD(S)
   - Store by case number

3. **Extract Stability Data**
   - Find stability summary sections
   - Extract Roll (Heel) angle
   - Extract Pitch (Trim) angle
   - Extract Area Ratio (Actual & Required)

4. **Associate Data**
   - Match draft marks with stability data
   - Link to correct damage case
   - Validate data completeness

5. **Generate Report**
   - Create formatted Excel workbook
   - Apply professional styling
   - Output to specified file

### Pattern Recognition

The software uses regular expressions to identify:
- **Damage cases**: `DAMAGE STABILITY Case-X (Compartment XXX Flooded)`
- **Intact condition**: `INTACT TOW CONDITION` or `Damage = NONE`
- **Draft marks**: `AFT(P) X.XX AFT(S) X.XX FWD(P) X.XX FWD(S) X.XX`
- **Roll/Heel**: `Roll = X.XX Deg`
- **Pitch/Trim**: `Pitch = X.XX Deg`
- **Area Ratio**: `Area Ratio >= X.XX ACTUAL Passes`

## Sample Output

### Console Output
```
Processing MOSES file: D7519_TOW_ANALYSIS.txt
Extracted 9 damage cases

Extracted Data:
------------------------------------------------------------------------------------------------------------------------
Case     Aft Port   Aft Stbd   Fwd Port   Fwd Stbd   Heel     Trim     Area Ratio   Remarks
------------------------------------------------------------------------------------------------------------------------
Intact   2.43       2.43       1.69       1.69         -0.01    0.53      163.19 Pass
1        2.00       1.94       1.50       1.45          0.14    0.47      141.36 Pass
2        2.52       1.98       1.65       1.11          1.28    0.82      127.60 Pass
3        1.82       1.94       1.44       1.56         -0.30    0.36      126.73 Pass
4        2.29       1.83       1.78       1.32          1.08    0.47      151.92 Pass
5        2.20       1.76       1.84       1.40          1.03    0.33      162.25 Pass
6        2.13       1.69       1.91       1.48          1.01    0.21      170.22 Pass
7        2.07       1.58       2.04       1.56          1.14    0.02      178.00 Pass
8        1.93       1.93       1.49       1.49         -0.01    0.41      151.54 Pass
------------------------------------------------------------------------------------------------------------------------
Excel file created successfully: Table_10.4.2_Damage_Stability_Results.xlsx
```

## Troubleshooting

### Common Issues

**Problem**: No data extracted  
**Solution**: Verify file contains stability summary sections with Area Ratio data

**Problem**: Missing draft marks  
**Solution**: Check that draft mark reading sections are present and formatted correctly

**Problem**: Incorrect case numbering  
**Solution**: Ensure damage case headers follow standard MOSES format

**Problem**: Missing Heel or Trim values  
**Solution**: Confirm Roll and Pitch values are in stability summary

## Customization

### Modify Output Format
Edit the `create_excel_report()` function to customize:
- Colors and fonts
- Column widths
- Number formats
- Header text

### Add Additional Fields
Extend the `extract_moses_data()` function to extract:
- Additional stability parameters
- Loading conditions
- Environmental conditions

## Dependencies
- **Python**: 3.6 or higher
- **openpyxl**: For Excel file generation

## License
Provided as-is for marine engineering stability analysis

## Version
**Version 2.0** - Enhanced with complete Table 10.4.2 format support
- Added Heel and Trim extraction
- Added Wind Area Ratio extraction  
- Improved case number identification
- Professional Excel formatting

## Support
For issues with extraction:
1. Verify MOSES file format
2. Check console error messages
3. Review sample output files
4. Validate input data completeness
