# MAUDE Binary Columns Automator

## Overview

This automated script processes MAUDE Excel files to add binary columns for device problems, patient problems, and patient outcomes. It automatically detects unique problems and outcomes in your data and creates corresponding binary columns with proper formatting.

## Features

✅ **Automatic Detection**: Finds all unique device problems, patient problems, and patient outcomes  
✅ **Style Preservation**: Maintains all original Excel formatting, fonts, colors, and design  
✅ **Smart Processing**: Avoids creating duplicate binary columns if they already exist  
✅ **Backup Creation**: Automatically creates timestamped backups before processing  
✅ **Visual Formatting**: Applies green/red color coding for binary values (1=green, 0=red)  
✅ **Batch Processing**: Can process multiple files at once  
✅ **Error Handling**: Validates files and provides detailed feedback  

## Requirements

- Python 3.6+
- Required packages: `pandas`, `openpyxl`, `numpy`

Install dependencies:
```bash
pip install pandas openpyxl numpy
```

## Usage

### Method 1: Process files in current directory
```bash
python3 maude_binary_automator.py
```
This will automatically find and process all MAUDE Excel files in the current directory.

### Method 2: Process a specific file
```bash
python3 maude_binary_automator.py path/to/your/maude_file.xlsx
```

### Method 3: Process multiple specific files
```bash
python3 maude_binary_automator.py file1.xlsx file2.xlsx file3.xlsx
```

## What the Script Does

1. **Validates** the Excel file has the required structure (Events sheet with problem/outcome columns)
2. **Captures** all original styling and formatting
3. **Detects** unique values in:
   - Device Problem columns (Device Problem 1, Device Problem 2, etc.)
   - Patient Problem columns (Patient Problem 1, Patient Problem 2, etc.)
   - Patient Outcome columns (Patient Outcome 1, Patient Outcome 2, etc.)
4. **Creates** binary columns for each unique problem/outcome:
   - `Device_[Problem_Name]` for device problems
   - `Patient_[Problem_Name]` for patient problems  
   - `Outcome_[Outcome_Name]` for patient outcomes
5. **Applies** formatting:
   - Green cells for "1" (problem/outcome exists)
   - Red cells for "0" (problem/outcome does not exist)
   - Bold headers for binary columns
6. **Preserves** all original styling and formatting
7. **Creates** a backup file before making changes

## Output

- **Modified file**: Original file updated with binary columns
- **Backup file**: `filename_backup_YYYYMMDD_HHMMSS.xlsx`
- **Console output**: Detailed processing information and statistics

## Example Output

```
🚀 MAUDE Binary Columns Automator
========================================
📁 Found 1 file(s) to process

============================================================
Processing: MAUDEMetrics_2025-08-19_2310.xlsx
============================================================
✅ File validated successfully
📊 Loaded 1655 rows and 104 columns
🎨 Capturing original styling...
✅ Original styling captured successfully!

📋 Found unique values:
   • Device Problems: 71
   • Patient Problems: 130
   • Patient Outcomes: 6

🔧 Creating binary columns...
✅ 207 new binary columns created. Total columns now: 311
📝 Data written. Now restoring original formatting...
✅ Original formatting restored!
🎨 Applying green/red formatting to binary columns...
✅ Binary column formatting applied!
💾 Backup saved: MAUDEMetrics_2025-08-19_2310_backup_20250820_022903.xlsx

🎉 Successfully processed MAUDEMetrics_2025-08-19_2310.xlsx!
   • Original columns: 104
   • Binary columns added: 207
   • Total columns: 311
   • Total rows: 1655
```

## File Requirements

Your Excel file must have:
- An "Events" sheet
- Columns containing "Device Problem", "Patient Problem", or "Patient Outcome" in their names
- Standard MAUDE data structure

## What Needs to Be Consistent

### ✅ **REQUIRED (Must be consistent)**

#### 1. **File Format**
- **File extension**: Must be `.xlsx` (Excel format)
- **File type**: Must be a valid Excel file (not corrupted)

#### 2. **Sheet Name**
- **Sheet name**: Must be exactly `"Events"` (case-sensitive)
- **Location**: Must be the first sheet or accessible by name

#### 3. **Column Name Patterns**
The script looks for these **exact text patterns** in column names:

**Device Problems:**
- Column names must contain: `"Device Problem"`
- Examples: `"Device Problem 1"`, `"Device Problem 2"`, etc.

**Patient Problems:**
- Column names must contain: `"Patient Problem"`
- Examples: `"Patient Problem 1"`, `"Patient Problem 2"`, etc.

**Patient Outcomes:**
- Column names must contain: `"Patient Outcome"`
- Examples: `"Patient Outcome 1"`, `"Patient Outcome 2"`, etc.

### 🔄 **FLEXIBLE (Can vary)**

#### 1. **Number of Columns**
- ✅ Can have 1, 5, 10, 20, 50, or any number of problem columns
- ✅ Column numbering can be: 1, 2, 3... or 1, 2, 3, 4, 5... or any sequence

#### 2. **Data Content**
- ✅ Problem/outcome values can be any text
- ✅ Can have missing values (NaN)
- ✅ Can have special characters in problem names

#### 3. **File Structure**
- ✅ Can have additional sheets (Summary, etc.)
- ✅ Can have additional columns (Event ID, dates, etc.)
- ✅ Column order doesn't matter

#### 4. **File Naming**
- ✅ File name can be anything (script looks for "MAUDE" in name for auto-detection)
- ✅ Can be in any directory

### ❌ **What Will Cause Issues**

#### 1. **Column Name Variations**
```
❌ These WON'T work:
"Device Problems"  # Missing space
"device problem"   # Wrong case
"Device_Problem"   # Wrong separator
"DeviceProblem"    # No space
"Device Problem:"  # Extra colon
```

#### 2. **Sheet Name Issues**
```
❌ These WON'T work:
"events"           # Wrong case
"Event"            # Missing 's'
"Events Sheet"     # Extra words
```

#### 3. **File Format Issues**
```
❌ These WON'T work:
.xls               # Old Excel format
.csv               # CSV file
.txt               # Text file
```

### 🛠️ **How to Check Your File**

You can test if your file will work by running this simple check:

```python
import pandas as pd

# Check your file
file_path = "your_file.xlsx"
xl = pd.ExcelFile(file_path)
print("Sheets:", xl.sheet_names)

df = pd.read_excel(file_path, sheet_name='Events')
print("Columns with 'Device Problem':", [col for col in df.columns if 'Device Problem' in col])
print("Columns with 'Patient Problem':", [col for col in df.columns if 'Patient Problem' in col])
print("Columns with 'Patient Outcome':", [col for col in df.columns if 'Patient Outcome' in col])
```

### 📝 **Summary**

**Must be consistent:**
- ✅ `.xlsx` file format
- ✅ `"Events"` sheet name
- ✅ Column names containing `"Device Problem"`, `"Patient Problem"`, `"Patient Outcome"`

**Can vary:**
- ✅ Number of problem columns
- ✅ Problem/outcome values
- ✅ Additional data and columns
- ✅ File naming and location

The script is quite robust and handles most variations, but the column name patterns and sheet name must match exactly!

## Binary Column Naming

The script creates clean column names by:
- Removing "Binary_" prefix (if present)
- Converting spaces to underscores
- Removing special characters like commas, parentheses, and slashes
- Adding appropriate prefixes (Device_, Patient_, Outcome_)

## Safety Features

- **Automatic backups**: Creates timestamped backups before any changes
- **Duplicate detection**: Won't create duplicate binary columns
- **File validation**: Checks file structure before processing
- **Error handling**: Provides clear error messages if something goes wrong

## Troubleshooting

**"No MAUDE Excel files found"**
- Make sure your file contains "MAUDE" in the filename
- Or provide the specific file path as an argument

**"No 'Events' sheet found"**
- Ensure your Excel file has a sheet named "Events"

**"No problem/outcome columns found"**
- Verify your file has columns with "Device Problem", "Patient Problem", or "Patient Outcome" in their names

**"No new binary columns needed"**
- This is normal! It means your file already has binary columns for all problems/outcomes

## Support

The script is designed to be robust and handle various MAUDE file formats. If you encounter issues, check:
1. File structure matches MAUDE requirements
2. Python and required packages are installed
3. File is not corrupted or password-protected
