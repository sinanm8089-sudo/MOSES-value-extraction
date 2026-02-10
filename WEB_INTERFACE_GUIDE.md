# MOSES Web Interface - User Guide

## Quick Start

Simply **double-click** `moses_web_interface.html` to open it in your web browser. No installation required!

## Features

✨ **Modern, Beautiful Design**
- Purple gradient background
- Smooth animations
- Responsive layout that works on any device

🎯 **Easy to Use**
- Drag and drop your MOSES files
- Or click to browse and select
- Real-time data preview
- One-click Excel download

📊 **Complete Data Extraction**
- GM (metacentric height) values
- All draft mark readings (AFT/FWD, Port/Starboard)
- Multiple damage cases
- Intact condition (NONE)

## How to Use

### Step 1: Open the Interface
Double-click `moses_web_interface.html` - it will open in your default web browser.

### Step 2: Upload Your MOSES File
You have two options:
- **Drag & Drop**: Drag your `.txt` or `.out` file onto the upload area
- **Click to Browse**: Click the upload area and select your file

### Step 3: Extract Data
Click the **"Extract Data"** button to process your file.

### Step 4: Review Results
The extracted data will appear in a formatted table showing:
- Damage case identifiers
- GM values
- Draft marks for all positions

### Step 5: Download Excel
Click **"Download Excel File"** to save the results as a formatted Excel spreadsheet.

## File Requirements

**Supported Formats:**
- `.txt` files
- `.out` files

**Required Sections in MOSES File:**
- `+++ D R A F T   M A R K   R E A D I N G S +++` sections
- `+++ S T A B I L I T Y   S U M M A R Y +++` sections
- Damage case headers OR intact condition markers

## Technical Details

**Works Offline:** Once the page loads, you can use it without internet (except for the first load which needs to download the Excel library).

**Browser Compatibility:**
- ✅ Chrome
- ✅ Firefox
- ✅ Edge
- ✅ Safari

**No Installation:** Everything runs in your browser - no Python, no npm, no servers needed!

## Troubleshooting

### "No data was extracted" Error

**Possible Causes:**
1. File format doesn't match MOSES output
2. Missing required sections (DRAFT MARKS or STABILITY SUMMARY)
3. Incorrect section headers

**Solutions:**
- Verify your file is a genuine MOSES output file
- Check that stability analysis sections are present
- Ensure damage case headers follow the standard format

### Excel Download Not Working

**Possible Cause:** First-time use requires internet to load the Excel library

**Solution:** Make sure you have internet connection when you first open the page

### Page Won't Open

**Possible Cause:** File associations

**Solution:** Right-click the HTML file → "Open with" → Choose your web browser

## Advantages Over Python Script

| Feature | Web Interface | Python Script |
|---------|--------------|---------------|
| Installation | None required | Requires Python + openpyxl |
| User Interface | Visual, drag-and-drop | Command line only |
| Data Preview | Live table view | Console text only |
| Platform | Any device with browser | Computer with Python |
| Updates | Single file to replace | Multiple files |

## Files

- **moses_web_interface.html** - Complete standalone application (all-in-one file)

That's it! Everything is contained in one HTML file for maximum simplicity.

## Support

If you encounter issues:
1. Check that your MOSES file has the required sections
2. Try opening the file in a different browser
3. Verify your MOSES file format matches the expected structure

---

**Created for:** MOSES (Marine Operations Simulation) stability analysis  
**Format:** Standalone HTML5 web application  
**Dependencies:** None (uses CDN for Excel library)
