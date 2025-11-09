#!/usr/bin/env python3
"""
Wamo Bank Statement Categorization Script
Replicates Excel categorization logic for Wamo bank statements
"""

import re
import pandas as pd
from datetime import datetime
from typing import List, Tuple, Optional
import PyPDF2
from pathlib import Path
import sys


# Month colors for Excel formatting
MONTH_COLORS = {
    1: "#FFCCCC", 2: "#FFE5CC", 3: "#FFFFCC",
    4: "#E5FFCC", 5: "#CCFFCC", 6: "#CCFFE5",
    7: "#CCFFFF", 8: "#CCE5FF", 9: "#CCCCFF",
    10: "#E5CCFF", 11: "#FFCCFF", 12: "#FFCCE5"
}


def parse_number(val: str) -> float:
    """
    Parse number from various formats (EU/US, with currency symbols, parentheses, etc.)
    Handles: €1,234.56, (123.45), 123-, 1.234,56
    """
    if not val or not isinstance(val, str):
        return 0.0
    
    val = str(val).strip()
    
    # Detect negative indicators
    has_parens = re.match(r'^\(.*\)$', val)
    has_trailing_minus = val.endswith('-')
    
    # Remove currency symbols and spaces
    val = re.sub(r'[\s€$£]', '', val)
    
    # Handle EU decimal format (1.234,56 or 1234,56)
    if ',' in val and '.' in val:
        # EU format: 1.234,56 -> remove dots, replace comma with dot
        val = val.replace('.', '').replace(',', '.')
    elif ',' in val and '.' not in val:
        # Could be EU decimal: 1234,56 -> replace comma with dot
        val = val.replace(',', '.')
    else:
        # US format or no decimals
        val = val.replace(',', '')
    
    # Remove parentheses and trailing minus
    val = re.sub(r'[()]', '', val).replace('-', '')
    
    try:
        num = float(val)
        return -num if (has_parens or has_trailing_minus) else num
    except ValueError:
        return 0.0


def parse_date_smart(date_str: str) -> Optional[datetime]:
    """
    Parse date from various formats
    Supports: yyyy-mm-dd, dd/mm/yyyy, dd-mm-yyyy, "30 September 2025", ISO formats
    """
    if not date_str:
        return None
    
    date_str = str(date_str).strip()
    
    # Try Wamo format: "30 September 2025"
    month_names = {
        'january': 1, 'february': 2, 'march': 3, 'april': 4,
        'may': 5, 'june': 6, 'july': 7, 'august': 8,
        'september': 9, 'october': 10, 'november': 11, 'december': 12
    }
    match = re.match(r'^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$', date_str)
    if match:
        d = int(match[1])
        month_name = match[2].lower()
        y = int(match[3])
        if month_name in month_names:
            m = month_names[month_name]
            return datetime(y, m, d)
    
    # Try ISO format yyyy-mm-dd or yyyy/mm/dd
    match = re.match(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$', date_str)
    if match:
        y, m, d = int(match[1]), int(match[2]), int(match[3])
        return datetime(y, m, d)
    
    # Try EU format dd/mm/yyyy or dd-mm-yyyy
    match = re.match(r'^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$', date_str)
    if match:
        d, m, y = int(match[1]), int(match[2]), int(match[3])
        return datetime(y, m, d)
    
    # Try pandas parsing as fallback
    try:
        return pd.to_datetime(date_str, dayfirst=True)
    except:
        return None


def get_transaction_type(detail: str) -> str:
    """
    Categorize transaction based on detail string
    """
    if not detail:
        return "other"
    
    detail_lower = detail.lower()
    
    # Wamo-specific patterns
    if re.search(r'card transaction', detail_lower):
        return "card payment"
    if re.search(r'sent money to', detail_lower):
        return "outgoing transfer"
    if re.search(r'received money from', detail_lower):
        return "incoming transfer"
    if re.search(r'wise charges', detail_lower):
        return "transfer fee"
    if re.search(r'cashback|balance_cashback', detail_lower):
        return "cashback"
    if re.search(r'card ending in', detail_lower):
        return "card transaction"
    
    # Cheques
    if re.search(r'cheque.*deposit', detail_lower):
        return "cheque deposit"
    if re.search(r'cheque.*returned', detail_lower):
        return "cheque returned fee"
    if re.search(r'cheques returned', detail_lower):
        return "cheque returned"
    if re.search(r'cheque', detail_lower):
        return "cheque payment"
    
    # Transfers
    if re.search(r'account to account', detail_lower):
        return "account transfer"
    if re.search(r'transfer between own accounts', detail_lower):
        return "internal transfer"
    if re.search(r'sct inwards', detail_lower):
        return "incoming sct transfer"
    if re.search(r'sct outwards', detail_lower):
        return "outgoing sct transfer"
    if re.search(r'instant payments inwards', detail_lower):
        return "instant payment in"
    if re.search(r'instant payment', detail_lower):
        return "instant payment"
    
    # Fees & charges
    if re.search(r'fee', detail_lower):
        return "bank fee"
    if re.search(r'charge', detail_lower):
        return "bank charge"
    if re.search(r'administration fee', detail_lower):
        return "administration fee"
    if re.search(r'standing instruction charge', detail_lower):
        return "standing instruction charge"
    if re.search(r'standing instruction', detail_lower):
        return "standing instruction"
    
    # Salaries & employment
    if re.search(r'salary', detail_lower):
        return "salary"
    if re.search(r'employment', detail_lower):
        return "employment payment"
    if re.search(r'stipendio|stipend', detail_lower):
        return "stipend/salary"
    
    # Loans & repayments
    if re.search(r'repayment.*principal', detail_lower):
        return "loan principal repayment"
    if re.search(r'repayment.*interest', detail_lower):
        return "loan interest repayment"
    if re.search(r'loan', detail_lower):
        return "loan"
    
    # Taxes & government
    if re.search(r'tax', detail_lower):
        return "tax payment"
    if re.search(r'vat', detail_lower):
        return "vat payment"
    if re.search(r'customs', detail_lower):
        return "customs payment"
    if re.search(r'government|gov', detail_lower):
        return "government payment"
    
    # ATM deposits
    if re.search(r'atm.*cash.*deposit', detail_lower):
        return "atm cash deposit"
    
    # 24x7 payments
    if re.search(r'24x7 pay', detail_lower):
        return "third party payment"
    if re.search(r'24x7 bill', detail_lower):
        return "bill payment"
    if re.search(r'24x7 mobile pay', detail_lower):
        return "mobile payment"
    
    # Direct debits
    if re.search(r'sdd outwards', detail_lower):
        return "direct debit out"
    
    # Insurance
    if re.search(r'mapfre|msv life|insurance', detail_lower):
        return "insurance payment"
    
    # Retail / food / hospitality
    if re.search(r'hotel', detail_lower):
        return "hotel payment"
    if re.search(r'catering', detail_lower):
        return "catering payment"
    if re.search(r'butcher|food|supermarket|restaurant|eat', detail_lower):
        return "food & retail"
    if re.search(r'retail', detail_lower):
        return "retail payment"
    
    # Utilities
    if re.search(r'electricity|water|gas|utility', detail_lower):
        return "utility payment"
    
    # Misc
    if re.search(r'refund', detail_lower):
        return "refund"
    if re.search(r'deposit', detail_lower):
        return "deposit"
    if re.search(r'withdrawal', detail_lower):
        return "withdrawal"
    
    return "other"


def extract_invoice(detail: str) -> str:
    """
    Extract invoice number from transaction detail
    """
    if not detail:
        return ""
    
    match = re.search(r'(invoice|inv|fatt(?:ura)?\s*nr?)\s*([0-9]+)', detail, re.IGNORECASE)
    if match:
        return f"invoice {match[2]}"
    return ""


def extract_counterparty(detail: str) -> str:
    """
    Extract counterparty name from transaction detail
    """
    if not detail:
        return ""
    
    # Wamo-specific patterns
    # "Sent money to <counterparty>"
    match = re.search(r'sent money to\s+(.+?)(?:\s+transaction:|$)', detail, re.IGNORECASE)
    if match:
        return match[1].strip()
    
    # "Received money from <counterparty>"
    match = re.search(r'received money from\s+(.+?)(?:\s+with reference|transaction:|$)', detail, re.IGNORECASE)
    if match:
        return match[1].strip()
    
    # "Card transaction of EUR issued by <merchant>"
    match = re.search(r'issued by\s+(.+?)(?:\s+card ending|transaction:|$)', detail, re.IGNORECASE)
    if match:
        return match[1].strip()
    
    # Check for tax administration reference
    if re.search(r'administratio', detail, re.IGNORECASE):
        tax_ref = re.search(r'ADMINISTRATIO\s+([0-9]+)', detail, re.IGNORECASE)
        if tax_ref:
            return tax_ref[1]
    
    # Clean common transaction prefixes/suffixes
    cleaned = detail
    patterns_to_remove = [
        r'24x7\s*pay\s*third\s*parties',
        r'24x7\s*pay',
        r'third\s*parties',
        r'payment order outwards same day',
        r'payment order outwards',
        r'account to account transfer express deposits',
        r'account to account transfer',
        r'transfer between own accounts',
        r'sct instant payments inwards',
        r'sct inwards',
        r'sct outwards',
        r'standing instruction charge',
        r'standing instruction',
        r'administration fee',
        r'unprocessed standing instruction charge',
        r'sdd outwards fee',
        r'atm cash deposit',
        r'cheque deposit.*$',
        r'cheque returned fee.*$',
        r'cheque book order fee.*$',
        r'cheque\s+\d+.*',
        r'relation:\s*[^,]+',
        r'reason:\s*[^,]+',
        r'value date\s*-\s*[0-9/]+',
        r'ref\s*:\s*[-0-9A-Za-z.]+.*$',
        r'\s+eur\s+[0-9.,]+',
    ]
    
    for pattern in patterns_to_remove:
        cleaned = re.sub(pattern, '', cleaned, flags=re.IGNORECASE)
    
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    
    # Split on common delimiters
    cleaned = re.split(r'ref\s*:|value date|relation:', cleaned, flags=re.IGNORECASE)[0].strip()
    
    # Look for company suffix
    company_match = re.search(r'\b([A-Za-z][A-Za-z &.\'-]*\s(?:ltd|limited|plc|co|company))\b', cleaned, re.IGNORECASE)
    if company_match:
        return company_match[1]
    
    # Split on EUR amount
    eur_split = re.split(r'\s+eur\s+', cleaned, flags=re.IGNORECASE)[0].strip()
    if eur_split and len(eur_split) >= 3:
        cleaned = eur_split
    
    # Look for person titles
    person_match = re.search(r'\b(Mr|Ms|Mrs|Dr)\.?\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+)?\b', cleaned)
    if person_match:
        return person_match[0]
    
    # Look for uppercase sequences
    upper_match = re.search(r'\b([A-Z][A-Z &.\'-]{2,})\b', cleaned)
    if upper_match:
        return upper_match[1]
    
    # Look for capitalized sequences
    cap_match = re.search(r'\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,4})\b', cleaned)
    if cap_match:
        return cap_match[1]
    
    # Return first 5 words
    words = cleaned.split()[:5]
    return ' '.join(words).strip()


def capitalize_first(text: str) -> str:
    """Capitalize first letter, lowercase rest"""
    if not text:
        return ""
    return text[0].upper() + text[1:].lower()


def limit_length(text: str, max_len: int = 26) -> str:
    """Limit text length"""
    if not text:
        return ""
    return text[:max_len] if len(text) > max_len else text


def extract_transactions_from_pdf(pdf_path: str) -> pd.DataFrame:
    """
    Extract transaction data from Wamo PDF statement
    Wamo format: Date | Description | Incoming | Outgoing | Balance
    """
    transactions = []
    
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            
            # Extract text from all pages
            full_text = ""
            for page in pdf_reader.pages:
                full_text += page.extract_text() + "\n"
            
            lines = full_text.split('\n')
            
            # Find transaction section (starts after "Description Incoming Outgoing Amount")
            in_transactions = False
            current_transaction = None
            
            for i, line in enumerate(lines):
                line = line.strip()
                if not line:
                    continue
                
                # Detect header row
                if re.search(r'Description\s+Incoming\s+Outgoing\s+Amount', line, re.IGNORECASE):
                    in_transactions = True
                    continue
                
                # Stop at end of statement or new section
                if in_transactions and re.search(r'(Opening Balance|Closing Balance|Total|Page \d+)', line, re.IGNORECASE):
                    in_transactions = False
                    continue
                
                if in_transactions:
                    # Wamo transaction pattern: Date description text ...amounts
                    # Date format: "30 September 2025" or "2 September 2025"
                    date_match = re.match(r'^(\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4})\s+(.+)', line)
                    
                    if date_match:
                        # Save previous transaction if exists
                        if current_transaction and current_transaction.get('description'):
                            transactions.append(current_transaction)
                        
                        # Start new transaction
                        date_str = date_match.group(1)
                        rest_of_line = date_match.group(2)
                        
                        date_obj = parse_date_smart(date_str)
                        
                        # Try to extract amounts from the line
                        # Look for numbers at the end (balance and possibly incoming/outgoing)
                        # Pattern: ...description... -123.45 1,234.56 OR ...description... 1,234.56
                        amounts = re.findall(r'[-]?[\d,]+\.\d{2}', rest_of_line)
                        
                        # The last number is always the balance
                        balance = None
                        incoming = None
                        outgoing = None
                        
                        if amounts:
                            # Last amount is balance
                            balance = parse_number(amounts[-1])
                            
                            # If there are 2 amounts, check which one is incoming/outgoing
                            if len(amounts) >= 2:
                                # Second-to-last could be incoming or outgoing
                                second_amount = parse_number(amounts[-2])
                                
                                # Wamo shows positive for incoming, with minus sign for outgoing
                                if amounts[-2].startswith('-'):
                                    outgoing = abs(second_amount)
                                else:
                                    incoming = abs(second_amount)
                        
                        # Extract description (text before amounts)
                        # Remove all amount patterns from the line
                        description = rest_of_line
                        for amt in amounts:
                            description = description.replace(amt, '')
                        description = description.strip()
                        
                        current_transaction = {
                            'Date': date_obj,
                            'description': description,
                            'incoming': incoming,
                            'outgoing': outgoing,
                            'balance': balance
                        }
                    
                    elif current_transaction:
                        # Continuation of previous transaction description
                        # Look for amounts on this line
                        amounts = re.findall(r'[-]?[\d,]+\.\d{2}', line)
                        
                        if amounts:
                            # This line might have the amounts we missed
                            balance = parse_number(amounts[-1])
                            current_transaction['balance'] = balance
                            
                            if len(amounts) >= 2:
                                second_amount = parse_number(amounts[-2])
                                if amounts[-2].startswith('-'):
                                    current_transaction['outgoing'] = abs(second_amount)
                                else:
                                    current_transaction['incoming'] = abs(second_amount)
                            
                            # Remove amounts from line and add to description
                            desc_part = line
                            for amt in amounts:
                                desc_part = desc_part.replace(amt, '')
                            desc_part = desc_part.strip()
                            if desc_part:
                                current_transaction['description'] += ' ' + desc_part
                        else:
                            # Pure description continuation
                            current_transaction['description'] += ' ' + line
            
            # Add last transaction
            if current_transaction and current_transaction.get('description'):
                transactions.append(current_transaction)
            
            # Convert to dataframe with proper Amount column
            result = []
            for trans in transactions:
                date = trans.get('Date')
                description = trans.get('description', '').strip()
                incoming = trans.get('incoming', 0) or 0
                outgoing = trans.get('outgoing', 0) or 0
                
                # Calculate signed amount (positive for incoming, negative for outgoing)
                if incoming > 0:
                    amount = incoming
                elif outgoing > 0:
                    amount = -outgoing
                else:
                    # No explicit incoming/outgoing, try to infer from description
                    amount = 0
                
                if date and description:
                    result.append({
                        'Date': date,
                        'Detail': description,
                        'Amount': amount
                    })
    
    except Exception as e:
        print(f"Error reading PDF: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()
    
    if not result:
        print("Warning: No transactions found in PDF")
        return pd.DataFrame(columns=['Date', 'Detail', 'Amount'])
    
    return pd.DataFrame(result)


def process_transactions(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Process transactions and split into incoming/outgoing
    """
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()
    
    # Split into incoming (amount >= 0) and outgoing (amount < 0)
    incoming = df[df['Amount'] >= 0].copy()
    outgoing = df[df['Amount'] < 0].copy()
    
    # Process both dataframes
    for transactions_df in [incoming, outgoing]:
        if transactions_df.empty:
            continue
        
        # Add derived columns
        transactions_df['Type'] = transactions_df['Detail'].apply(
            lambda x: limit_length(capitalize_first(get_transaction_type(str(x).lower())))
        )
        transactions_df['Invoice'] = transactions_df['Detail'].apply(
            lambda x: limit_length(capitalize_first(extract_invoice(str(x))))
        )
        transactions_df['Counterparty'] = transactions_df['Detail'].apply(
            lambda x: limit_length(capitalize_first(extract_counterparty(str(x))))
        )
        
        # Convert numeric-only counterparties to numbers
        transactions_df['Counterparty'] = transactions_df['Counterparty'].apply(
            lambda x: int(x) if str(x).isdigit() else x
        )
    
    # Sort by date
    incoming = incoming.sort_values('Date').reset_index(drop=True)
    outgoing = outgoing.sort_values('Date').reset_index(drop=True)
    
    return incoming, outgoing


def apply_month_colors(writer, sheet_name: str, df: pd.DataFrame):
    """
    Apply month-based coloring to Excel worksheet
    """
    if df.empty:
        return
    
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # Apply colors based on month
    for idx, row in df.iterrows():
        if pd.notna(row['Date']):
            month = row['Date'].month
            color = MONTH_COLORS.get(month, "#FFFFFF")
            
            # Convert hex to RGB for xlsxwriter
            hex_color = color.lstrip('#')
            
            # Create format with background color
            cell_format = workbook.add_format({'bg_color': color})
            
            # Apply to entire row (0-indexed row + 1 for header)
            excel_row = idx + 1
            for col in range(6):  # A:F columns
                worksheet.write(excel_row, col, df.iloc[idx, col], cell_format)


def export_to_excel(source_df: pd.DataFrame, incoming_df: pd.DataFrame, 
                    outgoing_df: pd.DataFrame, output_path: str):
    """
    Export dataframes to Excel with formatting
    """
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        # Write SOURCE sheet
        source_df.to_excel(writer, sheet_name='SOURCE', index=False)
        
        # Write INCOMING sheet
        if not incoming_df.empty:
            incoming_df.to_excel(writer, sheet_name='INCOMING', index=False)
            apply_month_colors(writer, 'INCOMING', incoming_df)
        else:
            pd.DataFrame(columns=['Date', 'Detail', 'Amount', 'Type', 'Invoice', 'Counterparty']).to_excel(
                writer, sheet_name='INCOMING', index=False
            )
        
        # Write OUTGOING sheet
        if not outgoing_df.empty:
            outgoing_df.to_excel(writer, sheet_name='OUTGOING', index=False)
            apply_month_colors(writer, 'OUTGOING', outgoing_df)
        else:
            pd.DataFrame(columns=['Date', 'Detail', 'Amount', 'Type', 'Invoice', 'Counterparty']).to_excel(
                writer, sheet_name='OUTGOING', index=False
            )
        
        # Auto-adjust column widths
        for sheet_name in ['SOURCE', 'INCOMING', 'OUTGOING']:
            if sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                worksheet.set_column('A:A', 12)  # Date
                worksheet.set_column('B:B', 50)  # Detail
                worksheet.set_column('C:C', 12)  # Amount
                worksheet.set_column('D:D', 26)  # Type
                worksheet.set_column('E:E', 26)  # Invoice
                worksheet.set_column('F:F', 26)  # Counterparty
    
    print(f"Excel file created: {output_path}")


def main():
    """Main execution function"""
    if len(sys.argv) < 2:
        print("Usage: python wamo_categorization.py <path_to_wamo_statement.pdf> [output.xlsx]")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "categorized_statement.xlsx"
    
    if not Path(pdf_path).exists():
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)
    
    print(f"Processing: {pdf_path}")
    
    # Extract transactions from PDF
    print("Extracting transactions from PDF...")
    df = extract_transactions_from_pdf(pdf_path)
    
    if df.empty:
        print("Error: No transactions found in PDF")
        sys.exit(1)
    
    print(f"Found {len(df)} transactions")
    
    # Process and categorize
    print("Categorizing transactions...")
    incoming_df, outgoing_df = process_transactions(df)
    
    print(f"  Incoming: {len(incoming_df)} transactions")
    print(f"  Outgoing: {len(outgoing_df)} transactions")
    
    # Export to Excel
    print("Exporting to Excel...")
    export_to_excel(df, incoming_df, outgoing_df, output_path)
    
    print(f"\nComplete! Output saved to: {output_path}")


if __name__ == "__main__":
    main()

