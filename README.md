# TB-GL Linker

A Python tool that automatically links Trial Balance accounts to General Ledger details by creating hyperlinks between them. The tool copies the GL data into the TB file and creates clickable references for easy navigation.

## Features

- **Intelligent File Analysis**: Automatically detects the structure of TB and GL files
- **Fuzzy Account Matching**: Uses similarity scoring to match accounts even with slight name differences
- **Flexible Structure Support**: Works with various TB and GL formats
- **One-Click Navigation**: Creates hyperlinks from TB accounts to GL detail sections
- **Self-Contained Output**: Combines both TB and GL data in a single file

## Requirements

- Python 3.6 or higher
- openpyxl library

## Installation

1. Ensure Python 3 is installed:
```bash
python3 --version
```

2. Install the required library:
```bash
pip3 install openpyxl
```

Or if you're using a virtual environment:
```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install openpyxl
```

## Usage

### Basic Usage

```bash
python3 tb_gl_linker.py TB.xlsx GL.xlsx
```

This will create a new file called `TB_linked.xlsx` with:
- All original TB data
- A new "Reference" column with hyperlinks
- GL data copied as a new sheet
- Clickable links from TB accounts to GL sections

### Advanced Usage

Specify a custom output filename:
```bash
python3 tb_gl_linker.py trial_balance.xlsx general_ledger.xlsx -o combined_report.xlsx
```

### Command Line Options

- `tb_file`: Path to the Trial Balance Excel file (required)
- `gl_file`: Path to the General Ledger Excel file (required)
- `-o, --output`: Custom output filename (optional, default: TB_file_linked.xlsx)

## How It Works

1. **File Loading**: Opens both TB and GL Excel files
2. **Structure Analysis**: 
   - Identifies TB columns (Account Code, Account Name, Debit, Credit)
   - Finds GL account sections
3. **Account Matching**: 
   - Uses fuzzy string matching to link TB accounts to GL accounts
   - Requires 80%+ similarity for a match
4. **Data Integration**:
   - Copies entire GL sheet into TB workbook
   - Adds "Reference" column with hyperlinks
5. **Output**: Saves combined workbook with working internal links

## Example Workflow

1. Prepare your files:
   - `Q1_TrialBalance.xlsx` - Your trial balance
   - `Q1_GeneralLedger.xlsx` - Your general ledger

2. Run the linker:
```bash
python3 tb_gl_linker.py Q1_TrialBalance.xlsx Q1_GeneralLedger.xlsx
```

3. Open the output file `Q1_TrialBalance_linked.xlsx`

4. Click any "View Details" link to jump to the GL account details

## Supported File Structures

### Trial Balance
The tool looks for these common column headers:
- Account/Account Code/Acct Code
- Account Name/Description
- Debit
- Credit

### General Ledger
The tool identifies account sections by:
- Account names as section headers (usually in column A)
- Transaction details following each account header
- Date/amount patterns in transaction rows

## Troubleshooting

### "Could not find header row in Trial Balance"
- Ensure your TB file has recognizable column headers
- Headers should be in the first 20 rows
- Must contain keywords like "Account", "Debit", or "Credit"

### Low matching rate
- Check for spelling differences between TB and GL account names
- The tool requires 80%+ similarity for matching
- Consider standardizing account names across files

### Module not found error
- Install openpyxl: `pip3 install openpyxl`
- Use a virtual environment if system-wide install fails

## Tips for Best Results

1. **Consistent Naming**: Use similar account names in both TB and GL
2. **Clear Headers**: Ensure column headers are clearly labeled
3. **Clean Data**: Remove extra spaces or special characters from account names
4. **File Location**: Keep TB and GL files in the same directory for easier access

## Output File Structure

The linked file contains:
- **Original TB sheet**: With new "Reference" column
- **General Ledger Detail sheet**: Complete GL data
- **Hyperlinks**: Format `#'General Ledger Detail'!A[row]`
- **Unmatched accounts**: Show "N/A" in Reference column

## License

This tool is provided as-is for accounting and bookkeeping purposes.

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Ensure your Excel files follow standard accounting formats
3. Verify Python and openpyxl are properly installed