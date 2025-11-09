#!/usr/bin/env python3
"""
PDF Bank Statement Categorization Script
Processes PDF bank statements and categorizes transactions
"""

import re
import pandas as pd
from datetime import datetime
from typing import List, Tuple, Optional
import PyPDF2
from pathlib import Path
import sys

# Import shared categorization functions
from common_categorization import (
    MONTH_COLORS,
    parse_number,
    parse_date_smart,
    get_transaction_type,
    extract_invoice,
    extract_counterparty,
    capitalize_first,
    limit_length
)


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


def export_to_excel(source_df: pd.DataFrame, incoming_df: pd.DataFrame, 
                    outgoing_df: pd.DataFrame, output_path: str):
    """
    Export dataframes to Excel with formatting and proper tables
    """
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Define formats
        date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
        currency_format = workbook.add_format({'num_format': '#,##0.00'})
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        
        # Write SOURCE sheet
        source_df.to_excel(writer, sheet_name='SOURCE', index=False, startrow=0)
        worksheet_source = writer.sheets['SOURCE']
        
        # Format SOURCE sheet as table
        if not source_df.empty:
            last_row = len(source_df)
            worksheet_source.add_table(0, 0, last_row, 2, {
                'name': 'SOURCE_TABLE',
                'style': 'Table Style Medium 2',
                'columns': [
                    {'header': 'Date', 'format': date_format},
                    {'header': 'Detail'},
                    {'header': 'Amount', 'format': currency_format}
                ]
            })
            
            # Set column widths
            worksheet_source.set_column('A:A', 12, date_format)
            worksheet_source.set_column('B:B', 70)
            worksheet_source.set_column('C:C', 15, currency_format)
        
        # Write and format INCOMING sheet
        if not incoming_df.empty:
            incoming_df.to_excel(writer, sheet_name='INCOMING', index=False, startrow=0)
            worksheet_incoming = writer.sheets['INCOMING']
            
            last_row = len(incoming_df)
            worksheet_incoming.add_table(0, 0, last_row, 5, {
                'name': 'INCOMING_TABLE',
                'style': 'Table Style Medium 9',
                'columns': [
                    {'header': 'Date', 'format': date_format},
                    {'header': 'Detail'},
                    {'header': 'Amount', 'format': currency_format},
                    {'header': 'Type'},
                    {'header': 'Invoice'},
                    {'header': 'Counterparty'}
                ]
            })
            
            # Apply month-based row colors
            for idx, row in incoming_df.iterrows():
                if pd.notna(row['Date']):
                    month = row['Date'].month
                    color = MONTH_COLORS.get(month, "#FFFFFF")
                    date_cell_format = workbook.add_format({
                        'bg_color': color,
                        'num_format': 'yyyy-mm-dd'
                    })
                    text_cell_format = workbook.add_format({
                        'bg_color': color
                    })
                    currency_cell_format = workbook.add_format({
                        'bg_color': color,
                        'num_format': '#,##0.00'
                    })
                    
                    # Apply formatting to data rows (skip header)
                    excel_row = idx + 1
                    worksheet_incoming.write(excel_row, 0, row['Date'], date_cell_format)
                    worksheet_incoming.write(excel_row, 1, row['Detail'], text_cell_format)
                    worksheet_incoming.write(excel_row, 2, row['Amount'], currency_cell_format)
                    worksheet_incoming.write(excel_row, 3, row['Type'], text_cell_format)
                    worksheet_incoming.write(excel_row, 4, row['Invoice'], text_cell_format)
                    worksheet_incoming.write(excel_row, 5, row['Counterparty'], text_cell_format)
            
            # Set column widths
            worksheet_incoming.set_column('A:A', 12)
            worksheet_incoming.set_column('B:B', 50)
            worksheet_incoming.set_column('C:C', 15)
            worksheet_incoming.set_column('D:D', 26)
            worksheet_incoming.set_column('E:E', 26)
            worksheet_incoming.set_column('F:F', 26)
        else:
            # Create empty sheet with headers
            empty_df = pd.DataFrame(columns=['Date', 'Detail', 'Amount', 'Type', 'Invoice', 'Counterparty'])
            empty_df.to_excel(writer, sheet_name='INCOMING', index=False)
        
        # Write and format OUTGOING sheet
        if not outgoing_df.empty:
            outgoing_df.to_excel(writer, sheet_name='OUTGOING', index=False, startrow=0)
            worksheet_outgoing = writer.sheets['OUTGOING']
            
            last_row = len(outgoing_df)
            worksheet_outgoing.add_table(0, 0, last_row, 5, {
                'name': 'OUTGOING_TABLE',
                'style': 'Table Style Medium 4',
                'columns': [
                    {'header': 'Date', 'format': date_format},
                    {'header': 'Detail'},
                    {'header': 'Amount', 'format': currency_format},
                    {'header': 'Type'},
                    {'header': 'Invoice'},
                    {'header': 'Counterparty'}
                ]
            })
            
            # Apply month-based row colors
            for idx, row in outgoing_df.iterrows():
                if pd.notna(row['Date']):
                    month = row['Date'].month
                    color = MONTH_COLORS.get(month, "#FFFFFF")
                    date_cell_format = workbook.add_format({
                        'bg_color': color,
                        'num_format': 'yyyy-mm-dd'
                    })
                    text_cell_format = workbook.add_format({
                        'bg_color': color
                    })
                    currency_cell_format = workbook.add_format({
                        'bg_color': color,
                        'num_format': '#,##0.00'
                    })
                    
                    # Apply formatting to data rows (skip header)
                    excel_row = idx + 1
                    worksheet_outgoing.write(excel_row, 0, row['Date'], date_cell_format)
                    worksheet_outgoing.write(excel_row, 1, row['Detail'], text_cell_format)
                    worksheet_outgoing.write(excel_row, 2, row['Amount'], currency_cell_format)
                    worksheet_outgoing.write(excel_row, 3, row['Type'], text_cell_format)
                    worksheet_outgoing.write(excel_row, 4, row['Invoice'], text_cell_format)
                    worksheet_outgoing.write(excel_row, 5, row['Counterparty'], text_cell_format)
            
            # Set column widths
            worksheet_outgoing.set_column('A:A', 12)
            worksheet_outgoing.set_column('B:B', 50)
            worksheet_outgoing.set_column('C:C', 15)
            worksheet_outgoing.set_column('D:D', 26)
            worksheet_outgoing.set_column('E:E', 26)
            worksheet_outgoing.set_column('F:F', 26)
        else:
            # Create empty sheet with headers
            empty_df = pd.DataFrame(columns=['Date', 'Detail', 'Amount', 'Type', 'Invoice', 'Counterparty'])
            empty_df.to_excel(writer, sheet_name='OUTGOING', index=False)
    
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

