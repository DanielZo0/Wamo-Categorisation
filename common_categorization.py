"""
Common categorization functions shared across all bank statement processors
"""

import re
from datetime import datetime
from typing import Optional


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
    Handles: €1,234.56, (123.45), 123-, 1.234,56, "1,234.56"
    """
    if not val or not isinstance(val, str):
        return 0.0
    
    val = str(val).strip().strip('"')
    
    # Detect negative indicators
    has_parens = re.match(r'^\(.*\)$', val)
    has_trailing_minus = val.endswith('-')
    is_negative = val.startswith('-')
    
    # Remove currency symbols and spaces
    val = re.sub(r'[\s€$£]', '', val)
    
    # Remove negative signs temporarily
    val = val.replace('-', '').replace('(', '').replace(')', '')
    
    # Handle EU decimal format (1.234,56 or 1234,56) vs US format (1,234.56)
    if ',' in val and '.' in val:
        # If both present, determine which is decimal separator
        comma_pos = val.rfind(',')
        dot_pos = val.rfind('.')
        if comma_pos > dot_pos:
            # EU format: 1.234,56
            val = val.replace('.', '').replace(',', '.')
        else:
            # US format: 1,234.56
            val = val.replace(',', '')
    elif ',' in val and '.' not in val:
        # Could be EU decimal or thousands separator
        # If only one comma and less than 3 digits after, it's decimal
        if val.count(',') == 1 and len(val.split(',')[1]) <= 2:
            val = val.replace(',', '.')
        else:
            val = val.replace(',', '')
    else:
        # Only dots or nothing special
        val = val.replace(',', '')
    
    try:
        num = float(val)
        return -num if (has_parens or has_trailing_minus or is_negative) else num
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
    
    # Try month name format: "30 September 2025"
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
    match = re.match(r'^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$', date_str)
    if match:
        y, m, d = int(match[1]), int(match[2]), int(match[3])
        return datetime(y, m, d)
    
    # Try EU format dd/mm/yyyy or dd-mm-yyyy
    match = re.match(r'^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$', date_str)
    if match:
        d, m, y = int(match[1]), int(match[2]), int(match[3])
        return datetime(y, m, d)
    
    # Try pandas parsing as fallback
    try:
        import pandas as pd
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
    
    # Card transactions
    if re.search(r'card transaction', detail_lower):
        return "card payment"
    if re.search(r'card ending in', detail_lower):
        return "card transaction"
    
    # Transfers
    if re.search(r'sent money to', detail_lower):
        return "outgoing transfer"
    if re.search(r'received money from', detail_lower):
        return "incoming transfer"
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
    
    # Cheques
    if re.search(r'cheque.*deposit', detail_lower):
        return "cheque deposit"
    if re.search(r'cheque.*returned', detail_lower):
        return "cheque returned fee"
    if re.search(r'cheques returned', detail_lower):
        return "cheque returned"
    if re.search(r'cheque', detail_lower):
        return "cheque payment"
    
    # Fees & charges
    if re.search(r'wise charges', detail_lower):
        return "transfer fee"
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
    if re.search(r'cashback|balance_cashback', detail_lower):
        return "cashback"
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
    
    # Pattern 1: "Sent money to <counterparty>"
    match = re.search(r'sent money to\s+(.+?)(?:\s+transaction:|$)', detail, re.IGNORECASE)
    if match:
        return match[1].strip()
    
    # Pattern 2: "Received money from <counterparty>"
    match = re.search(r'received money from\s+(.+?)(?:\s+with reference|transaction:|$)', detail, re.IGNORECASE)
    if match:
        return match[1].strip()
    
    # Pattern 3: "Card transaction of EUR issued by <merchant>"
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
    company_match = re.search(r'\b([A-Z][A-Za-z &.\'-]*\s(?:ltd|limited|plc|co|company))\b', cleaned, re.IGNORECASE)
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

