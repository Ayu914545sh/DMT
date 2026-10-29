"""
Data cleaning and validation utilities for MCA & GST ingestion.
Strict compliance with data governance, phone normalization, date parsing,
and spreadsheet formula injection prevention.
"""

import re
import math
import datetime
from decimal import Decimal, InvalidOperation
from typing import Optional, Tuple, Any, Dict, List


def clean_mobile(raw_number: Any, strict_prefix: bool = True) -> Tuple[Optional[str], str, Optional[str]]:
    """
    Robust phone number cleaning for all imported MCA and GST data.
    
    Rules:
    1. Convert value to text.
    2. Trim whitespace.
    3. Remove apostrophes, spaces, brackets, hyphens, dots, and other non-numeric characters.
    4. Remove a leading country code such as +91, 0091, or 91 (when 12 digits), or leading 0 (when 11 digits).
    5. Keep only a valid 10-digit Indian mobile number (starts with 6, 7, 8, or 9).
    6. Store invalid or ambiguous numbers as blank/null (None) and show them in the validation report.
    7. Preserve original uploaded value and store cleaned value separately.
    8. Statuses: 'valid', 'cleaned', 'missing', 'invalid', 'multiple_numbers_review'.
    
    Returns:
        (cleaned_phone_or_None, phone_status, error_reason_or_None)
    """
    if raw_number is None:
        return None, 'missing', 'Value is null or missing'

    # 1. Convert input to text before processing
    if isinstance(raw_number, float):
        if math.isnan(raw_number) or math.isinf(raw_number):
            return None, 'missing', 'Value is null or NaN'
        if raw_number.is_integer():
            raw_text = str(int(raw_number))
        else:
            raw_text = str(raw_number).strip()
    elif isinstance(raw_number, int):
        raw_text = str(raw_number)
    else:
        raw_text = str(raw_number)

    # 2. Trim leading and trailing whitespace
    trimmed = raw_text.strip()

    # Empty / blank check
    if not trimmed or trimmed.lower() in ('nan', 'none', 'null', 'na', '-', 'nil'):
        return None, 'missing', 'Value is blank or null'

    # Detect multiple phone numbers (e.g. contains separators like '/', ';', ',', 'or', 'and', '&', '|')
    multiple_delimiters = ['/', ';', ',', 'or', 'and', '&', '|']
    has_multiple = any(delim in trimmed.lower() for delim in multiple_delimiters)
    if has_multiple:
        parts = re.split(r'[/;,|]|(?:\s+(?:or|and|&)\s+)', trimmed, flags=re.IGNORECASE)
        digit_parts = [re.sub(r'\D', '', p) for p in parts if re.sub(r'\D', '', p)]
        if len(digit_parts) > 1:
            return None, 'multiple_numbers_review', f'Multiple phone numbers detected ({trimmed})'

    # Check for scientific notation (e.g. 9.87E+09)
    if 'e' in trimmed.lower():
        return None, 'invalid', f'Scientific notation cannot be verified ({trimmed})'

    # Convert Excel-style values such as '9876543210.0' into '9876543210'
    candidate = trimmed
    if re.search(r'\.0+$', candidate):
        candidate = re.sub(r'\.0+$', '', candidate)

    # Strip common text label prefixes (e.g. 'tel:', 'mob:', 'phone:', 'ph:', 'm:')
    candidate = re.sub(r'^(tel|mob|mobile|phone|ph|contact|call|m)[\s.:#-]*', '', candidate, flags=re.IGNORECASE)

    # Remove leading/trailing apostrophes, backticks, single quotes, double quotes, and hashtags
    candidate = re.sub(r"^[`'\"\s#]+", "", candidate)
    candidate = re.sub(r"[`'\"\s#]+$", "", candidate)

    was_modified = (candidate != trimmed)

    # Remove formatting characters: spaces, hyphens, brackets, dots, slashes, colons
    formatted_cleaned = re.sub(r"[\s()\-./\\[\]:,;]", "", candidate)
    if formatted_cleaned != candidate:
        was_modified = True
    candidate = formatted_cleaned

    # Handle Indian country-code prefixes with '+' or '00':
    if candidate.startswith("+91"):
        candidate = candidate[3:]
        was_modified = True
    elif candidate.startswith("0091"):
        candidate = candidate[4:]
        was_modified = True
    elif candidate.startswith("+"):
        candidate = candidate[1:]
        was_modified = True

    # Strip any remaining non-digits
    digits_only = re.sub(r'\D', '', candidate)
    if digits_only != candidate:
        was_modified = True
    candidate = digits_only

    # Handle length-based country codes and trunk prefixes:
    if len(candidate) == 13 and (candidate.startswith("910") or candidate.startswith("091")):
        candidate = candidate[3:]
        was_modified = True
    elif len(candidate) == 12 and candidate.startswith("91"):
        candidate = candidate[2:]
        was_modified = True
    elif len(candidate) == 11 and candidate.startswith("0"):
        candidate = candidate[1:]
        was_modified = True

    # Keep only a valid 10-digit Indian mobile number
    if len(candidate) != 10:
        return None, 'invalid', f'Expected exactly 10 digits, found {len(candidate)} digits ({trimmed})'

    if candidate[0] not in ('6', '7', '8', '9'):
        return None, 'invalid', f'Phone number must start with 6, 7, 8, or 9 (starts with {candidate[0]})'

    # Check for dummy/repetitive number (e.g. 9999999999, 8888888888, 7777777777)
    if re.match(r'^(\d)\1{9}$', candidate):
        return candidate, 'FLAGGED', f'Dummy or repetitive phone number ({trimmed})'

    # Determine status: 'valid' if already pristine 10 digits starting with 6-9 without modification, else 'cleaned'
    if not was_modified and trimmed == candidate:
        status = 'valid'
    else:
        status = 'cleaned'

    return candidate, status, None


def clean_phone_extended(raw_number: Any, strict_prefix: bool = True, default_ownership: str = 'unknown') -> Dict[str, Any]:
    """
    Comprehensive phone normalization adhering to Requirement 5:
    - Preserves original exact value.
    - Extracts extensions (e.g. ext 102, x45).
    - Safely splits multiple numbers without breaking formatting or country codes.
    - Normalizes Indian mobile (+91, 0091, 91, 0) and identifies landlines / international numbers.
    - Assigns ownership ('director', 'company', or 'unknown').
    - Returns rich dictionary for audit, caller list deduplication, and quality inbox.
    """
    raw_text = str(raw_number).strip() if raw_number is not None else ""
    if isinstance(raw_number, float) and (math.isnan(raw_number) or math.isinf(raw_number)):
        raw_text = ""
    elif isinstance(raw_number, float) and raw_number.is_integer():
        raw_text = str(int(raw_number))

    res = {
        'raw_value': raw_text,
        'primary_phone': None,
        'secondary_phones': [],
        'extension': "",
        'status': 'missing',
        'phone_validity': 'missing',
        'phone_type': 'missing',
        'ownership': default_ownership,
        'error_reason': None
    }

    if not raw_text or raw_text.lower() in ('nan', 'none', 'null', 'na', '-', 'nil', 'n/a'):
        res['error_reason'] = 'Value is blank or null'
        return res

    working = raw_text

    # 1. Extract extensions if present (e.g. ext 102, extn 45, x102, #45)
    ext_match = re.search(r'(?:ext(?:n|ension)?|[#x])\s*[:.]?\s*(\d{1,6})\b', working, flags=re.IGNORECASE)
    if ext_match:
        res['extension'] = ext_match.group(1)
        working = working[:ext_match.start()] + working[ext_match.end():]
        working = working.strip()

    # 2. Check for multiple numbers using delimiters
    parts = re.split(r'[/;,|]|\b(?:or|and|&)\b', working, flags=re.IGNORECASE)
    parts = [p.strip() for p in parts if p.strip()]

    valid_mobiles = []
    flagged_mobiles = []
    other_numbers = []
    reasons = []

    for part in parts:
        cleaned_part, part_status, part_reason = clean_mobile(part, strict_prefix=strict_prefix)
        if cleaned_part and part_status in ('valid', 'cleaned'):
            valid_mobiles.append((cleaned_part, part_status))
        elif cleaned_part and part_status == 'FLAGGED':
            flagged_mobiles.append((cleaned_part, part_reason))
        else:
            # Check if landline (starts with 0 or 1-5, 8-11 digits)
            digits = re.sub(r'\D', '', part)
            if digits.startswith('91') and len(digits) >= 12:
                digits = digits[2:]
            
            if (len(digits) in (10, 11) and digits.startswith('0')) or (len(digits) in (8, 9, 10) and digits[0] in ('1', '2', '3', '4', '5')):
                other_numbers.append(('landline', digits))
            elif digits:
                other_numbers.append(('invalid', digits))
            if part_reason:
                reasons.append(part_reason)

    if len(valid_mobiles) == 1 and len(parts) == 1 and not flagged_mobiles:
        num, stat = valid_mobiles[0]
        res['primary_phone'] = num
        res['status'] = stat
        res['phone_validity'] = stat
        res['phone_type'] = 'mobile'
    elif flagged_mobiles and not valid_mobiles:
        num, reason = flagged_mobiles[0]
        res['primary_phone'] = num
        res['status'] = 'FLAGGED'
        res['phone_validity'] = 'FLAGGED'
        res['phone_type'] = 'mobile'
        res['error_reason'] = reason
    elif len(valid_mobiles) >= 1:
        res['primary_phone'] = valid_mobiles[0][0]
        res['secondary_phones'] = [m[0] for m in valid_mobiles[1:]] + [num for _, num in other_numbers]
        res['status'] = 'multiple_numbers_review'
        res['phone_validity'] = 'multiple_numbers_review'
        res['phone_type'] = 'mobile'
        res['error_reason'] = f"Multiple phone numbers present ({raw_text})"
    elif other_numbers:
        first_type, first_num = other_numbers[0]
        res['primary_phone'] = first_num if first_type == 'landline' else None
        res['secondary_phones'] = [num for _, num in other_numbers[1:]]
        res['status'] = 'landline' if first_type == 'landline' else 'invalid'
        res['phone_validity'] = res['status']
        res['phone_type'] = first_type
        if first_type == 'landline':
            res['ownership'] = 'company'
        res['error_reason'] = reasons[0] if reasons else f"Invalid phone format ({raw_text})"
    else:
        res['status'] = 'invalid'
        res['phone_validity'] = 'invalid'
        res['phone_type'] = 'invalid'
        res['error_reason'] = reasons[0] if reasons else f"Invalid phone format ({raw_text})"

    return res


def clean_date(raw_date: Any, preferred_format: Optional[str] = None) -> Tuple[Optional[datetime.date], bool, Optional[str]]:
    """
    Parses date values strictly:
    - Supports datetime.date, datetime.datetime, pandas Timestamp.
    - Handles leading apostrophes e.g. "'2026-09-25".
    - Supports explicit formats (e.g. %m/%d/%Y for MCA, %Y-%m-%d for GST).
    - Returns (date_object_or_None, is_valid, error_reason_or_None).
    """
    if raw_date is None:
        return None, True, None
    
    if isinstance(raw_date, datetime.datetime):
        return raw_date.date(), True, None
    if isinstance(raw_date, datetime.date):
        return raw_date, True, None
    
    date_str = str(raw_date).strip().lstrip("`'\" ").rstrip("`'\" ")
    if not date_str or date_str.lower() in ('nan', 'none', 'null', 'na', '-'):
        return None, True, None
    
    # If preferred format supplied
    if preferred_format:
        try:
            parsed = datetime.datetime.strptime(date_str, preferred_format).date()
            return parsed, True, None
        except ValueError:
            pass
    
    # Standard format attempts
    candidate_formats = [
        "%m/%d/%Y",       # 09/24/2026 (MCA standard)
        "%Y-%m-%d",       # 2026-09-25 (GST standard)
        "%d/%m/%Y",       # 24/09/2026
        "%d-%m-%Y",       # 24-09-2026
        "%Y/%m/%d",       # 2026/09/25
        "%d-%b-%Y",       # 24-Sep-2026
        "%d-%B-%Y",       # 24-September-2026
    ]
    
    for fmt in candidate_formats:
        try:
            parsed = datetime.datetime.strptime(date_str, fmt).date()
            return parsed, True, None
        except ValueError:
            continue
    
    # Check for ISO format with time
    try:
        parsed_dt = datetime.datetime.fromisoformat(date_str.replace('Z', '+00:00'))
        return parsed_dt.date(), True, None
    except Exception:
        pass
    
    return None, False, f'INVALID_DATE_FORMAT: Unable to parse date "{date_str}"'


def clean_timestamp(raw_timestamp: Any) -> Tuple[Optional[datetime.datetime], bool, Optional[str]]:
    """
    Parses timestamp string or object into timezone-aware datetime.
    Supports ISO formats and common HTTP/RFC date strings.
    """
    if raw_timestamp is None:
        return None, True, None
    
    if isinstance(raw_timestamp, datetime.datetime):
        if datetime.timezone.utc and raw_timestamp.tzinfo is None:
            return raw_timestamp.replace(tzinfo=datetime.timezone.utc), True, None
        return raw_timestamp, True, None
    
    ts_str = str(raw_timestamp).strip().lstrip("`'\" ").rstrip("`'\" ")
    if not ts_str or ts_str.lower() in ('nan', 'none', 'null', 'na'):
        return None, True, None
    
    # ISO 8601 (MCA upload_date: 2026-09-25T03:49:37.768469+00:00)
    try:
        parsed = datetime.datetime.fromisoformat(ts_str.replace('Z', '+00:00'))
        return parsed, True, None
    except Exception:
        pass
    
    # JavaScript Date.toString() style: 'Sat Sep 26 2026 06:08:05 GMT+0200 (Central European Summer Time)'
    # Extract date and time portion up to GMT/UTC
    match = re.search(r'([A-Za-z]{3}\s+[A-Za-z]{3}\s+\d{1,2}\s+\d{4}\s+\d{2}:\d{2}:\d{2})', ts_str)
    if match:
        try:
            parsed = datetime.datetime.strptime(match.group(1), "%a %b %d %Y %H:%M:%S")
            return parsed.replace(tzinfo=datetime.timezone.utc), True, None
        except ValueError:
            pass
    
    return None, False, f'INVALID_TIMESTAMP: Unable to parse timestamp "{ts_str}"'


def clean_email(raw_email: Any) -> Tuple[Optional[str], str, Optional[str]]:
    """
    Validates and standardizes email syntax.
    Returns (cleaned_email_or_None, status, error_reason_or_None)
    where status is in ('VALID', 'FLAGGED', 'BLANK').
    """
    if raw_email is None:
        return None, 'BLANK', None
    
    val = str(raw_email).strip().lstrip("`'\" ").rstrip("`'\" ").lower()
    if not val or val in ('nan', 'none', 'null', 'na', '-', 'nil'):
        return None, 'BLANK', None
    
    email_regex = r'^[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}$'
    if re.fullmatch(email_regex, val):
        return val, 'VALID', None
    
    return None, 'FLAGGED', f'INVALID_EMAIL_SYNTAX: "{val}" is not a valid email'


def clean_capital(raw_capital: Any) -> Tuple[Optional[Decimal], bool, Optional[str]]:
    """
    Normalizes capital amount into Decimal.
    Handles currency symbols, commas, and scientific notation.
    """
    if raw_capital is None:
        return None, True, None
    
    val_str = str(raw_capital).strip().replace(',', '').replace('₹', '').replace('$', '').strip()
    if not val_str or val_str.lower() in ('nan', 'none', 'null', 'na', '-'):
        return None, True, None
    
    try:
        dec = Decimal(val_str)
        if dec < 0:
            return None, False, f'NEGATIVE_CAPITAL: Capital cannot be negative ({val_str})'
        return dec, True, None
    except InvalidOperation:
        return None, False, f'INVALID_DECIMAL: Unable to parse capital amount "{val_str}"'


def clean_pincode(raw_pincode: Any) -> Tuple[Optional[str], bool, Optional[str]]:
    """
    Validates Indian 6-digit postal code.
    """
    if raw_pincode is None:
        return None, True, None
    
    if isinstance(raw_pincode, (float, int)):
        if isinstance(raw_pincode, float) and raw_pincode.is_integer():
            pin_str = str(int(raw_pincode))
        else:
            pin_str = str(raw_pincode).split('.')[0].strip()
    else:
        pin_str = str(raw_pincode).strip().lstrip("`'\" ").rstrip("`'\" ")
    
    if not pin_str or pin_str.lower() in ('nan', 'none', 'null', 'na', '-'):
        return None, True, None
    
    # Strip any nondigits
    cleaned = re.sub(r'\D', '', pin_str)
    if len(cleaned) == 6:
        return cleaned, True, None
    
    return cleaned if cleaned else None, False, f'INVALID_PINCODE: Expected 6 digits, got "{pin_str}"'


def clean_pan(raw_pan: Any) -> Tuple[Optional[str], bool, Optional[str]]:
    """
    Validates Indian PAN: 5 letters, 4 digits, 1 letter (e.g. ASFPS3873C).
    """
    if raw_pan is None:
        return None, True, None
    
    pan_str = str(raw_pan).strip().upper().lstrip("`'\" ").rstrip("`'\" ")
    if not pan_str or pan_str.lower() in ('nan', 'none', 'null', 'na', '-'):
        return None, True, None
    
    if re.fullmatch(r'^[A-Z]{5}[0-9]{4}[A-Z]{1}$', pan_str):
        return pan_str, True, None
    
    return pan_str, False, f'INVALID_PAN_FORMAT: "{pan_str}" does not match standard 10-character PAN pattern'


def clean_gstin(raw_gst: Any) -> Tuple[Optional[str], bool, Optional[str]]:
    """
    Validates GSTIN: 15-character alphanumeric format.
    Standard format: 2 digit state code + 10 char PAN + 1 entity code + 'Z' + 1 checksum char.
    Preserves special registrations without discarding.
    """
    if raw_gst is None:
        return None, False, 'MISSING_GSTIN: GST number is required'
    
    gst_str = str(raw_gst).strip().upper().lstrip("`'\" ").rstrip("`'\" ")
    if not gst_str or gst_str.lower() in ('nan', 'none', 'null', 'na'):
        return None, False, 'MISSING_GSTIN: GST number is empty'
    
    # Standard format check
    is_standard = bool(re.fullmatch(r'^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$', gst_str))
    
    # 15 characters check
    if len(gst_str) != 15:
        return gst_str, False, f'INVALID_GSTIN_LENGTH: Expected 15 characters, got {len(gst_str)} for "{gst_str}"'
    
    return gst_str, True, None if is_standard else 'SPECIAL_GSTIN_PATTERN'


def clean_cin(raw_cin: Any) -> Tuple[Optional[str], bool, Optional[str]]:
    """
    Validates and standardizes MCA Corporate Identity Number (CIN) or LLPIN.
    Standard CIN: 21 characters e.g. U72200MH2018PTC312345
    Standard LLPIN: 7-8 characters e.g. AAA-1234 or AAB-5678
    """
    if raw_cin is None:
        return None, False, 'MISSING_CIN: Entity identifier is required'

    cin_str = str(raw_cin).strip().upper().lstrip("`'\" ").rstrip("`'\" ")
    if not cin_str or cin_str.lower() in ('nan', 'none', 'null', 'na', '-'):
        return None, False, 'MISSING_CIN: Entity identifier is empty'

    # CIN standard format: 1 char (U or L) + 5 digits + 2 state letters + 4 year digits + 3 entity code letters + 6 registration digits
    is_cin = bool(re.fullmatch(r'^[UL][0-9]{5}[A-Z]{2}[0-9]{4}[A-Z]{3}[0-9]{6}$', cin_str))
    # LLPIN format: e.g. AAA-1234, ABC-1234, or 7-8 alphanumeric
    is_llpin = bool(re.fullmatch(r'^[A-Z]{3}-?[0-9]{4}$|^[A-Z0-9]{7,8}$', cin_str))

    if is_cin or is_llpin:
        return cin_str, True, None

    # Lenient: if length is between 7 and 21 alphanumeric chars
    if len(cin_str) >= 7 and len(cin_str) <= 25:
        return cin_str, True, None

    return cin_str, False, f'INVALID_CIN_FORMAT: "{cin_str}" is not a recognized CIN or LLPIN format'


def clean_din(raw_din: Any) -> Tuple[Optional[str], bool, Optional[str]]:
    """
    Validates and standardizes Director Identification Number (DIN).
    Standard DIN: exactly 8 digits.
    """
    if raw_din is None:
        return None, True, None

    din_str = str(raw_din).strip().lstrip("`'\" ").rstrip("`'\" ")
    if not din_str or din_str.lower() in ('nan', 'none', 'null', 'na', '-'):
        return None, True, None

    # Handle float suffix .0
    if re.search(r'\.0+$', din_str):
        din_str = re.sub(r'\.0+$', '', din_str)

    cleaned = re.sub(r'\D', '', din_str)
    if len(cleaned) == 8:
        return cleaned, True, None

    if len(cleaned) > 0 and len(cleaned) <= 8:
        # Zero-pad if leading zeros were dropped by Excel
        padded = cleaned.zfill(8)
        return padded, True, None

    return cleaned if cleaned else None, False, f'INVALID_DIN: Expected 8 digits, got "{din_str}"'


def clean_name(raw_name: Any) -> Tuple[Optional[str], bool, Optional[str]]:
    """
    Cleans company, trade, legal, or director names.
    Trims whitespace, removes enclosing quotes, and normalizes internal spaces.
    Preserves legal entity suffixes and acronyms.
    """
    if raw_name is None:
        return None, True, None

    name_str = str(raw_name).strip().lstrip("`'\" ").rstrip("`'\" ")
    if not name_str or name_str.lower() in ('nan', 'none', 'null', 'na', '-'):
        return None, True, None

    # Normalize multiple whitespace characters
    cleaned = re.sub(r'\s+', ' ', name_str)
    return cleaned, True, None


def clean_address(raw_address: Any) -> Tuple[Optional[str], bool, Optional[str]]:
    """
    Cleans address strings: normalizes whitespace, removes enclosing quotes/backticks.
    """
    if raw_address is None:
        return None, True, None

    addr_str = str(raw_address).strip().lstrip("`'\" ").rstrip("`'\" ")
    if not addr_str or addr_str.lower() in ('nan', 'none', 'null', 'na', '-'):
        return None, True, None

    cleaned = re.sub(r'\s+', ' ', addr_str)
    return cleaned, True, None


def sanitize_spreadsheet_value(value: Any) -> Any:
    """
    Formula Injection Prevention for CSV and Excel exports:
    Any string starting with =, +, -, @, tab, or carriage return
    is prepended with a single quote to prevent execution in Excel/Calc.
    """
    if value is None:
        return ""
    val_str = str(value)
    if val_str and val_str[0] in ('=', '+', '-', '@', '\t', '\r'):
        return f"'{val_str}"
    return val_str

