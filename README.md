# Bank Statement Categorization Tool

Automated transaction categorization tool for bank statements. Processes PDF and CSV statements with intelligent categorization and professional Excel formatting.

## Supported Formats

- **PDF statements** - Automatically extracts and parses transaction data
- **CSV statements** - Processes exported CSV bank statements

## Features

- **Batch Processing** - Select and process multiple files at once
- **Smart Categorization** - Automatically categorizes transactions by type
- **Invoice Extraction** - Finds and extracts invoice numbers
- **Counterparty Detection** - Identifies payees and payers
- **Month-Based Coloring** - Visual grouping by month
- **Professional Excel Output** - Formatted tables with filters
- **Automatic Naming** - Output files named based on input files
- **Same-Directory Saving** - Outputs saved alongside input files

## Installation

1. Clone this repository:
```bash
git clone https://github.com/DanielZo0/Wamo-Categorisation.git
cd Wamo-Categorisation
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Quick Start (Windows)

**Double-click `Process Statements.bat`**

The tool will:
1. Open a file selection dialog
2. Let you select one or more statements (PDF or CSV)
3. Automatically process all selected files
4. Save results in the same directory as input files
5. Open file explorer to the output location when you press Enter

### Command Line (All Platforms)

```bash
python batch_statement_processor.py
```

The tool immediately opens a file browser where you can:
- Select a single file or multiple files (Ctrl/Cmd + Click)
- Mix PDF and CSV files in the same batch
- Process everything automatically

### Advanced: Direct Processing

For automation or scripting, call the processors directly:

**PDF Statements:**
```bash
python pdf_statement_processor.py statement.pdf output.xlsx
```

**CSV Statements:**
```bash
python csv_statement_processor.py statement.csv output.xlsx
```

## Output

Each statement generates an Excel file with three sheets:

### Sheet Structure

1. **SOURCE** - Raw transaction data
   - Date, Detail, Amount
   - Formatted as sortable Excel table

2. **INCOMING** - Income transactions with categorization
   - Date, Detail, Amount, Type, Invoice, Counterparty
   - Automatically categorized and analyzed

3. **OUTGOING** - Expense transactions with categorization
   - Date, Detail, Amount, Type, Invoice, Counterparty
   - Automatically categorized and analyzed

### Formatting Features

- ✓ Professional Excel tables with built-in filters
- ✓ Month-based row coloring (12 distinct colors)
- ✓ Currency formatting with thousands separators
- ✓ Date formatting (yyyy-mm-dd)
- ✓ Auto-sized columns for readability
- ✓ Clean, consistent styling

### Output Location

Files are automatically saved as `categorized_[original_name].xlsx` in the same directory as the input file.

## Transaction Categories

The script automatically recognizes and categorizes:

- Bank transfers (SCT, instant payments, internal)
- Fees and charges
- Salary and employment payments
- Loan repayments
- Tax and government payments
- Direct debits
- Insurance payments
- Retail and food purchases
- Utility payments
- And more...

## Example

### Batch Processing Multiple Files

1. Double-click `Categorize Statement.bat`
2. Select multiple statement files (hold Ctrl)
3. Watch the tool process each file:

```
============================================================
  BANK STATEMENT CATEGORIZATION TOOL
============================================================

Select one or more statement files (PDF or CSV)
Hold Ctrl/Cmd to select multiple files

3 file(s) selected
============================================================

[1/3] Processing: statement_Q3_2025.pdf
  Type: PDF
  Output: categorized_statement_Q3_2025.xlsx
------------------------------------------------------------
Processing: statement_Q3_2025.pdf
Extracting transactions from PDF...
Found 59 transactions
Categorizing transactions...
  Incoming: 10 transactions
  Outgoing: 49 transactions
Exporting to Excel...
Excel file created: categorized_statement_Q3_2025.xlsx
Complete! Output saved to: categorized_statement_Q3_2025.xlsx
  [OK] Completed successfully

[2/3] Processing: statement_export.csv
  Type: CSV
  Output: categorized_statement_export.xlsx
------------------------------------------------------------
Processing: statement_export.csv
Extracting transactions from CSV...
Found 338 transactions
Categorizing transactions...
  Incoming: 169 transactions
  Outgoing: 169 transactions
Exporting to Excel...
Excel file created: categorized_statement_export.xlsx
Complete! Output saved to: categorized_statement_export.xlsx
  [OK] Completed successfully

[3/3] Processing: july_statement.pdf
  Type: PDF
  Output: categorized_july_statement.xlsx
------------------------------------------------------------
  [OK] Completed successfully

============================================================
  PROCESSING COMPLETE
============================================================

Successful: 3
Failed: 0
Total: 3

Generated files:
  - C:\Statements\categorized_statement_Q3_2025.xlsx
  - C:\Statements\categorized_statement_export.xlsx
  - C:\Statements\categorized_july_statement.xlsx

Would you like to open the first output file? (Y/n):
```

## Requirements

- Python 3.8 or higher
- pandas
- PyPDF2
- xlsxwriter
- openpyxl

## Project Structure

```
├── batch_statement_processor.py  # Main UI - batch file selector
├── Process Statements.bat        # Windows launcher
├── common_categorization.py      # Shared categorization logic
├── pdf_statement_processor.py    # PDF statement processor
├── csv_statement_processor.py    # CSV statement processor
├── requirements.txt              # Python dependencies
└── README.md                     # This file
```

## Key Features

- **Modular Design** - Shared logic in common module
- **Batch Processing** - Process multiple files at once
- **Zero Configuration** - Just select files and go
- **Auto-Naming** - Output files named intelligently
- **Same-Directory Output** - Results saved with input files
- **Professional Formatting** - Excel tables, colors, proper types

## License

MIT License

## Author

Daniel Zo

