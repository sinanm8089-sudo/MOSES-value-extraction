#!/usr/bin/env python3
"""
MOSES Value Extraction Tool
Extracts damage stability data from MOSES output files and generates Excel reports
"""

import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def extract_moses_data(file_path):
    """
    Extract damage cases, GM values, and draft marks from MOSES output file
    
    Returns:
        list: List of dictionaries containing damage, GM, and draft data
    """
    # Try multiple encodings for better compatibility
    encodings = ['utf-8', 'ascii', 'latin-1', 'cp1252']
    content = None
    
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                content = f.read()
            break
        except (UnicodeDecodeError, FileNotFoundError) as e:
            if encoding == encodings[-1]:
                raise Exception(f"Unable to read file with common encodings. Error: {e}")
            continue
    
    if content is None:
        raise Exception("Failed to read file content")
    
    results = []
    lines = content.split('\n')
    
    current_damage = None
    current_gm = None
    current_drafts = {}
    in_stability_summary = False
    in_draft_marks = False
    
    for i, line in enumerate(lines):
        line_stripped = line.strip()
        
        # Detect damage case from section headers
        if 'DAMAGE STABILITY Case-' in line_stripped:
            match = re.search(r'Case-(\d+)\s*\(Compartment\s+(\w+)\s+Flooded\)', line_stripped)
            if match:
                current_damage = match.group(2)
        elif 'INTACT TOW CONDITION' in line_stripped or 'Damage = NONE' in line_stripped:
            current_damage = 'NONE'
        
        # Detect draft mark readings section
        if '+++ D R A F T   M A R K   R E A D I N G S +++' in line_stripped:
            in_draft_marks = True
            current_drafts = {}
            continue
        
        # Parse draft marks (looking for the data line)
        if in_draft_marks and 'AFT(P)' in line_stripped and 'AFT(S)' in line_stripped:
            # Extract draft values using regex
            draft_pattern = r'(\w+\([PS]\))\s+([\d.]+)'
            matches = re.findall(draft_pattern, line_stripped)
            
            for name, value in matches:
                if name not in ['MEAN(P)', 'MEAN(S)']:
                    current_drafts[name] = float(value)
            
            in_draft_marks = False
        
        # Detect stability summary section
        if '+++ S T A B I L I T Y   S U M M A R Y +++' in line_stripped:
            in_stability_summary = True
            continue
        
        # Extract GM value from stability summary
        # Handle both formats: with and without 'Passes' text
        if in_stability_summary and 'GM' in line_stripped and '>=' in line_stripped:
            # Try pattern with 'Passes' first, then without
            match = re.search(r'GM\s+>=\s+[\d.]+\s+M\s+([\d.]+)(?:\s+Passes)?', line_stripped)
            if match:
                current_gm = float(match.group(1))
                
                # Store the complete record
                if current_damage is not None and current_drafts:
                    results.append({
                        'damage': current_damage,
                        'gm': current_gm,
                        'drafts': current_drafts.copy()
                    })
                
                in_stability_summary = False
    
    return results

def create_excel_report(data, output_path):
    """
    Create formatted Excel report from extracted MOSES data
    
    Args:
        data: List of dictionaries with damage, GM, and draft information
        output_path: Path to save the Excel file
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "MOSES Data Extraction"
    
    # Define styles
    header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
    header_alignment = Alignment(horizontal='center', vertical='center')
    
    data_font = Font(name='Arial', size=10)
    data_alignment = Alignment(horizontal='center', vertical='center')
    
    border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    
    # Title
    ws.merge_cells('A1:F1')
    ws['A1'] = 'MOSES VALUE EXTRACTION REPORT'
    ws['A1'].font = Font(name='Arial', size=14, bold=True, color='366092')
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    
    # Headers
    headers = ['Damage', 'GM (m)', 'AFT(P) (m)', 'AFT(S) (m)', 'FWD(P) (m)', 'FWD(S) (m)']
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_num)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = border
    
    # Data rows
    for row_num, record in enumerate(data, 4):
        # Damage column
        cell = ws.cell(row=row_num, column=1)
        cell.value = record['damage']
        cell.font = data_font
        cell.alignment = data_alignment
        cell.border = border
        
        # GM column
        cell = ws.cell(row=row_num, column=2)
        cell.value = record['gm']
        cell.number_format = '0.00'
        cell.font = data_font
        cell.alignment = data_alignment
        cell.border = border
        
        # Draft columns
        draft_names = ['AFT(P)', 'AFT(S)', 'FWD(P)', 'FWD(S)']
        for col_num, draft_name in enumerate(draft_names, 3):
            cell = ws.cell(row=row_num, column=col_num)
            cell.value = record['drafts'].get(draft_name, None)
            if cell.value is not None:
                cell.number_format = '0.00'
            cell.font = data_font
            cell.alignment = data_alignment
            cell.border = border
    
    # Column widths
    column_widths = [15, 12, 12, 12, 12, 12]
    for col_num, width in enumerate(column_widths, 1):
        ws.column_dimensions[get_column_letter(col_num)].width = width
    
    # Row heights
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[3].height = 20
    
    # Save workbook
    wb.save(output_path)
    print(f"Excel file created successfully: {output_path}")

def main():
    """Main execution function"""
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python moses_extractor.py <input_file> [output_file]")
        print("Example: python moses_extractor.py out00001.txt moses_output.xlsx")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else 'moses_extraction.xlsx'
    
    print(f"Processing MOSES file: {input_file}")
    
    try:
        # Extract data
        data = extract_moses_data(input_file)
    except FileNotFoundError:
        print(f"Error: File '{input_file}' not found.")
        print("Please check the file path and try again.")
        sys.exit(1)
    except Exception as e:
        print(f"Error reading file: {e}")
        sys.exit(1)
    
    if not data:
        print("\nWarning: No data was extracted from the file.")
        print("\nPossible reasons:")
        print("  - File format doesn't match expected MOSES output")
        print("  - Missing 'STABILITY SUMMARY' sections")
        print("  - Missing 'DRAFT MARK READINGS' sections")
        print("  - Incorrect damage case headers")
        print("\nPlease verify that the input file is a valid MOSES output file.")
        sys.exit(1)
    
    print(f"Extracted {len(data)} damage cases")
    
    # Display extracted data
    print("\nExtracted Data:")
    print("-" * 80)
    for record in data:
        print(f"Damage: {record['damage']:<10} GM: {record['gm']:>6.2f} m")
        aft_p = f"{record['drafts']['AFT(P)']:.2f}" if 'AFT(P)' in record['drafts'] else 'N/A'
        aft_s = f"{record['drafts']['AFT(S)']:.2f}" if 'AFT(S)' in record['drafts'] else 'N/A'
        fwd_p = f"{record['drafts']['FWD(P)']:.2f}" if 'FWD(P)' in record['drafts'] else 'N/A'
        fwd_s = f"{record['drafts']['FWD(S)']:.2f}" if 'FWD(S)' in record['drafts'] else 'N/A'
        print(f"  Drafts: AFT(P)={aft_p}, AFT(S)={aft_s}, FWD(P)={fwd_p}, FWD(S)={fwd_s}")
        print("-" * 80)
    
    # Create Excel report
    create_excel_report(data, output_file)

if __name__ == '__main__':
    main()
