#!/usr/bin/env python3
"""
CSV Bank Statement Categorization Script
Processes CSV bank statements and categorizes transactions
"""

import re
import pandas as pd
from datetime import datetime
from typing import List, Tuple, Optional
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


def extract_transactions_from_csv(csv_path: str) -> pd.DataFrame:
    """
    Extract transaction data from BoV CSV statement
    """
    try:
        # Read the CSV file
        with open(csv_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        # Find the transaction history header
        transaction_start = -1
        for i, line in enumerate(lines):
            if 'Transaction History' in line:
                # Next line should be the column headers
                transaction_start = i + 2
                break
        
        if transaction_start == -1:
            print("Error: Could not find Transaction History header")
            return pd.DataFrame()
        
        # Read transactions from the found position
        df = pd.read_csv(csv_path, skiprows=transaction_start - 1, encoding='utf-8')
        
        # Clean column names
        df.columns = df.columns.str.strip()
        
        # Ensure we have the required columns
        if 'Date' not in df.columns or 'Detail' not in df.columns or 'Amount' not in df.columns:
            print(f"Error: Expected columns not found. Found: {df.columns.tolist()}")
            return pd.DataFrame()
        
        # Parse dates
        df['Date'] = df['Date'].apply(parse_date_smart)
        
        # Parse amounts
        df['Amount'] = df['Amount'].apply(parse_number)
        
        # Remove rows with no date
        df = df[df['Date'].notna()]
        
        # Keep only required columns
        df = df[['Date', 'Detail', 'Amount']].copy()
        
        return df
    
    except Exception as e:
        print(f"Error reading CSV: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()


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
        print("Usage: python bov_categorization.py <path_to_statement.csv> [output.xlsx]")
        sys.exit(1)
    
    csv_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "categorized_bov_statement.xlsx"
    
    if not Path(csv_path).exists():
        print(f"Error: File not found: {csv_path}")
        sys.exit(1)
    
    print(f"Processing: {csv_path}")
    
    # Extract transactions from CSV
    print("Extracting transactions from CSV...")
    df = extract_transactions_from_csv(csv_path)
    
    if df.empty:
        print("Error: No transactions found in CSV")
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

