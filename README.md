# Wamo Bank Statement Categorization

Python script to automatically categorize transactions from Wamo bank statements. This replicates Excel-based categorization logic with automated processing of PDF statements.

## Features

- Extracts transactions from Wamo PDF bank statements
- Automatically categorizes transactions by type (transfers, fees, salaries, etc.)
- Extracts invoice numbers and counterparty information
- Splits transactions into INCOMING and OUTGOING sheets
- Color-codes transactions by month for easy visualization
- Exports to Excel format with proper formatting

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

Run the script with a Wamo PDF statement:

```bash
python wamo_categorization.py statement.pdf
```

Or specify a custom output filename:

```bash
python wamo_categorization.py statement.pdf output.xlsx
```

## Output

The script generates an Excel file with three sheets:

1. **SOURCE** - Raw transaction data extracted from the PDF
2. **INCOMING** - Positive transactions (income) with categorization
3. **OUTGOING** - Negative transactions (expenses) with categorization

Each sheet includes:
- Date
- Detail (transaction description)
- Amount
- Type (auto-categorized transaction type)
- Invoice (extracted invoice numbers)
- Counterparty (extracted payee/payer names)

Rows are color-coded by month for easy visual grouping.

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

```bash
python wamo_categorization.py statement_7068982_EUR_2025-06-01_2025-09-30.pdf output.xlsx
```

Output:
```
Processing: statement_7068982_EUR_2025-06-01_2025-09-30.pdf
Extracting transactions from PDF...
Found 59 transactions
Categorizing transactions...
  Incoming: 10 transactions
  Outgoing: 49 transactions
Exporting to Excel...
Excel file created: output.xlsx

Complete! Output saved to: output.xlsx
```

The script successfully:
- Extracted 59 transactions from a 4-page Wamo PDF statement
- Identified 10 incoming transactions (payments received, cashback)
- Identified 49 outgoing transactions (card payments, transfers, fees)
- Categorized all transactions by type
- Extracted counterparty information (merchants, recipients)
- Applied month-based color coding for easy visual analysis

## Requirements

- Python 3.8 or higher
- pandas
- PyPDF2
- xlsxwriter
- openpyxl

## License

MIT License

## Author

Daniel Zo

