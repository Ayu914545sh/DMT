"""
Universal Import Engine for MCA and GST Datasets.
Supports CSV/XLSX, sheet selection, encoding/delimiter detection, column mapping,
conservative cleaning, change preview, conflict resolution, and idempotent publication.
"""

import os
import re
import hashlib
import datetime
import csv
from typing import Dict, List, Tuple, Any, Optional, Callable
import pandas as pd
import openpyxl
from django.db import transaction
from django.utils import timezone
from django.conf import settings

from .models import (
    MCACompany, MCADirectorEntry, GSTRecord,
    ImportBatch, StagedRow, QualityIssue,
    AuditEvent, SystemSetting
)
from .cleaners import (
    clean_mobile, clean_phone_extended, clean_date, clean_timestamp, clean_email,
    clean_capital, clean_pincode, clean_pan, clean_gstin,
    clean_cin, clean_din, clean_name, clean_address,
    sanitize_spreadsheet_value
)


# Canonical schema fields and approved aliases
MCA_FIELDS = {
    'entityId': [
        'cin', 'cin ', 'entityid', 'entity_id', 'llpin', 'corporate_identification_number',
        'corporate id', 'registration_number', 'reg_no', 'company_id', 'cin_no', 'cinno'
    ],
    'entityType': ['company type', 'company_type', 'entitytype', 'entity_type', 'type_of_entity', 'form_type', 'entity category'],
    'name': ['company name', 'company_name', 'name', 'entity_name', 'corporate_name', 'company', 'firm_name', 'business_name'],
    'state': ['state', 'registered_state', 'state_name'],
    'district': ['district', 'district_name'],
    'city': ['city', 'registered_city'],
    'roc': ['roc', 'roc_code', 'registrar_of_companies', 'roc name'],
    'nicCode': ['main division', 'main_division', 'niccode', 'nic_code', 'nic', 'industry_code', 'activity_code'],
    'nicLabel': ['main division description', 'main_division_description', 'niclabel', 'nic_label', 'activity_description', 'nic_description', 'industry_description'],
    'classOfCompany': ['class of company', 'classofcompany', 'class_of_company', 'company_class', 'class'],
    'dateOfIncorporation': ['date of incorporation', 'dateofincorporation', 'date_of_incorporation', 'incorporation_date', 'inc_date', 'doi'],
    'paidUpCapital': ['paid up capital', 'paidupcapital', 'paid_up_capital', 'paidup_capital', 'paid_capital'],
    'authorizedCapital': ['authorized capital', 'authorizedcapital', 'authorized_capital', 'auth_capital'],
    'email': ['company email address', 'company email', 'company_email', 'business_email', 'corporate_email', 'email_id', 'email id', 'email'],
    'registeredAddress': ['street address', 'street_address', 'registeredaddress', 'registered_address', 'address', 'reg_address', 'office_address'],
    'directorName': ['director name', 'directorname', 'director_name', 'partner_name', 'person_name', 'director / partner name', 'name of director'],
    'directorFirstName': ['director first name', 'director_first_name', 'first_name', 'first name'],
    'directorSecondName': ['director second name', 'director_second_name', 'second_name', 'second name', 'last_name', 'last name', 'surname'],
    'din': ['din number', 'din', 'din_no', 'dinno', 'director_din', 'director_identification_number', 'din / pan'],
    'directorEmail': ['director email address', 'director email', 'directoremail', 'director_email', 'person_email', 'contact_email', 'director email id'],
    'directorMobile': ['director mobile number', 'director mobile', 'directormobile', 'director_mobile', 'director_phone', 'director phone', 'mobile_no', 'mobile', 'phone', 'contact', 'contact_number', 'mobile number', 'phone number'],
    'postalCode': ['postal code', 'postal_code', 'pincode', 'pin_code', 'pin'],
    'streetAddress2': ['street address 2', 'street_address_2', 'address_2', 'address 2'],
    'stateCode': ['state code', 'state_code'],
    'companyOrigin': ['company origin', 'company_origin', 'origin'],
    'registrationNumber': ['registration number', 'registration_number', 'reg_number'],
    'whetherListed': ['whether listed or not', 'whether_listed_or_not', 'listed'],
    'companyCategory': ['company category', 'company_category', 'category'],
    'companySubCategory': ['company sub category', 'company_sub_category', 'sub_category'],
    'addressType': ['address type', 'address_type'],
    'dirPermAddress1': ['director permanent address 1', 'director_permanent_address_1'],
    'dirPermAddress2': ['director permanent address 2', 'director_permanent_address_2'],
    'dirPermCity': ['director permanent city', 'director_permanent_city'],
    'dirPermState': ['director permanent state', 'director_permanent_state'],
    'dirPermPincode': ['director permanent pincode', 'director_permanent_pincode'],
    'dirPresAddress1': ['director present address 1', 'director_present_address_1'],
    'dirPresAddress2': ['director present address 2', 'director_present_address_2'],
    'dirPresCity': ['director present city', 'director_present_city'],
    'dirPresState': ['director present state', 'director_present_state'],
    'dirPresPincode': ['director present pincode', 'director_present_pincode'],
    'skip': ['sr. no.', 'sr. no', 'sr_no', 'sr no', 's_no', 's.no.', 'serial_no', 'sl_no'],
    'upload_date': ['upload_date', 'upload_timestamp', 'source_timestamp', 'imported_date']
}

GST_FIELDS = {
    'gst_number': [
        'gst_number', 'gstin', 'gst_no', 'gst', 'gstin_number', 'gst_no/gstin',
        'gstin/gst no/gst number', 'gstin / gst no / gst number', 'gst number', 'gst in',
        'taxpayer_gstin', 'gstin_no', 'gst_registration_number'
    ],
    'legal_name': [
        'legal_name', 'legalname', 'business_legal_name', 'legal_entity_name', 'legal name',
        'entity_legal_name', 'name_of_business'
    ],
    'trade_name': [
        'trade_name', 'tradename', 'business_name', 'entity_name', 'name', 'trade_business_name',
        'trade name', 'trade / business name', 'doing_business_as'
    ],
    'nature_of_business': ['nature_of_business', 'business_nature', 'principal_business', 'nature of business'],
    'turnover': ['turnover', 'turnover_range', 'annual_turnover'],
    'type': ['type', 'taxpayer_type', 'constitution_of_business', 'constitution'],
    'registration_date': [
        'registration_date', 'reg_date', 'date_of_registration', 'dor', 'registration date',
        'date of registration', 'reg date', 'effective date of registration'
    ],
    'gst_type': ['gst_type', 'category', 'dealer_type', 'gst type'],
    'status': ['status', 'gst_status', 'current_status', 'gst status', 'taxpayer status', 'registration status'],
    'jurisdiction': ['jurisdiction', 'center_jurisdiction', 'central_jurisdiction'],
    'state_jurisdiction': ['state_jurisdiction', 'state_jurisdiction_code'],
    'pan': ['pan', 'pan_number', 'pan_no', 'pan no', 'pan number', 'permanent account number'],
    'mobile': [
        'mobile', 'mobile_number', 'phone', 'contact_number', 'phone_number', 'contact',
        'mobile/phone/contact', 'mobile / phone / contact', 'mobile/phone/contact number',
        'mobile / phone / contact number', 'mobile no',
        'phone no', 'contact no', 'mobile number', 'phone number', 'contact number'
    ],
    'email': ['email', 'email_address', 'contact_email', 'email id', 'email_id', 'e-mail', 'email id / contact email'],
    'address': ['address', 'principal_address', 'registered_address', 'principal place of business', 'business address'],
    'Pincode': ['pincode', 'pin_code', 'pin', 'postal_code', 'zip', 'pin/pincode', 'pin / pincode', 'pin code', 'postal code'],
    'state': ['state', 'state_name', 'state / ut', 'state name'],
    'city': ['city', 'district', 'town'],
    'skip': ['sr. no.', 'sr. no', 'sr_no', 'sr no', 's_no', 's.no.', 'serial_no', 'sl_no'],
    'insert_time': ['insert_time', 'insertion_time', 'upload_date', 'source_time']
}

GST_TARGET_ALIASES = {
    'gstin': 'gst_number',
    'gstin_number': 'gst_number',
    'gst_no': 'gst_number',
    'gst': 'gst_number',
    'mobile_number': 'mobile',
    'phone_number': 'mobile',
    'contact_number': 'mobile',
    'phone': 'mobile',
    'contact': 'mobile',
    'pincode': 'Pincode',
    'pin_code': 'Pincode',
    'pin': 'Pincode',
    'postal_code': 'Pincode',
    'legalname': 'legal_name',
    'tradename': 'trade_name',
    'registrationdate': 'registration_date',
    'reg_date': 'registration_date',
    'natureofbusiness': 'nature_of_business',
    'gsttype': 'gst_type',
    'pan_number': 'pan',
    'email_address': 'email',
    'sr_no': 'skip',
    'sr._no.': 'skip',
}

MCA_TARGET_ALIASES = {
    'cin': 'entityId',
    'cin_no': 'entityId',
    'llpin': 'entityId',
    'entity_id': 'entityId',
    'company_name': 'name',
    'company_type': 'entityType',
    'director_name': 'directorName',
    'director_first_name': 'directorFirstName',
    'director_second_name': 'directorSecondName',
    'director_mobile': 'directorMobile',
    'director_phone': 'directorMobile',
    'director_mobile_number': 'directorMobile',
    'mobile_number': 'directorMobile',
    'mobile': 'directorMobile',
    'director_email': 'directorEmail',
    'director_email_address': 'directorEmail',
    'company_email': 'email',
    'company_email_address': 'email',
    'din_number': 'din',
    'date_of_incorporation': 'dateOfIncorporation',
    'incorporation_date': 'dateOfIncorporation',
    'doi': 'dateOfIncorporation',
    'paid_up_capital': 'paidUpCapital',
    'authorized_capital': 'authorizedCapital',
    'registered_address': 'registeredAddress',
    'street_address': 'registeredAddress',
    'postal_code': 'postalCode',
    'pincode': 'postalCode',
    'sr_no': 'skip',
    'sr._no.': 'skip',
}


def calculate_file_checksum(file_path: str) -> str:
    """Computes SHA-256 hash of a file."""
    sha256 = hashlib.sha256()
    with open(file_path, 'rb') as f:
        while chunk := f.read(65536):
            sha256.update(chunk)
    return sha256.hexdigest()


def detect_csv_properties(file_path: str) -> Tuple[str, str]:
    """
    Detects CSV delimiter and encoding safely.
    Returns (delimiter, encoding).
    """
    encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252']
    detected_encoding = 'utf-8'
    detected_delimiter = ','
    
    for enc in encodings:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                sample = f.read(8192)
                detected_encoding = enc
                # Detect delimiter
                sniffer = csv.Sniffer()
                try:
                    dialect = sniffer.sniff(sample, delimiters=[',', '\t', ';', '|'])
                    detected_delimiter = dialect.delimiter
                except Exception:
                    # Fallback counting
                    counts = {d: sample.count(d) for d in [',', '\t', ';', '|']}
                    detected_delimiter = max(counts, key=counts.get) if counts else ','
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
            
    return detected_delimiter, detected_encoding


def read_excel_rows(file_path: str, sheet_name: Optional[str] = None, max_rows: Optional[int] = None) -> List[List[Any]]:
    """
    Reads rows from Excel file (.xlsx or .xls) safely.
    Dispatches to xlrd for legacy .xls files and openpyxl for .xlsx files,
    with automatic cross-fallback if file extension is mismatched.
    Normalizes Excel serial dates into ISO strings and integers.
    """
    ext = file_path.lower().split('.')[-1]
    
    # Try xlrd for .xls
    if ext == 'xls':
        try:
            import xlrd
            book = xlrd.open_workbook(file_path)
            if sheet_name and sheet_name in book.sheet_names():
                sheet = book.sheet_by_name(sheet_name)
            else:
                sheet = book.sheet_by_index(0)
            rows = []
            limit = max_rows if max_rows is not None else sheet.nrows
            for r in range(min(sheet.nrows, limit)):
                row_vals = []
                for c in range(sheet.ncols):
                    val = sheet.cell_value(r, c)
                    cell_type = sheet.cell_type(r, c)
                    if cell_type == xlrd.XL_CELL_DATE:
                        try:
                            dt_tuple = xlrd.xldate_as_tuple(val, book.datemode)
                            val = datetime.date(*dt_tuple[:3]).isoformat()
                        except Exception:
                            pass
                    elif cell_type == xlrd.XL_CELL_NUMBER and isinstance(val, float) and val.is_integer():
                        val = int(val)
                    row_vals.append(val)
                rows.append(row_vals)
            return rows
        except Exception:
            pass  # Fallback to openpyxl below
            
    # Try openpyxl for .xlsx (or fallback for .xls)
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
        rows = []
        for idx, row in enumerate(ws.iter_rows(values_only=True)):
            if max_rows is not None and idx >= max_rows:
                break
            rows.append(list(row))
        wb.close()
        return rows
    except Exception:
        # Cross fallback: if openpyxl failed on .xlsx, attempt xlrd in case it's actually BIFF8
        try:
            import xlrd
            book = xlrd.open_workbook(file_path)
            sheet = book.sheet_by_name(sheet_name) if sheet_name and sheet_name in book.sheet_names() else book.sheet_by_index(0)
            rows = []
            limit = max_rows if max_rows is not None else sheet.nrows
            for r in range(min(sheet.nrows, limit)):
                row_vals = [sheet.cell_value(r, c) for c in range(sheet.ncols)]
                rows.append(row_vals)
            return rows
        except Exception:
            return []


def get_excel_sheets(file_path: str) -> List[str]:
    """Returns worksheet names for Excel files (.xlsx and .xls)."""
    ext = file_path.lower().split('.')[-1]
    if ext == 'xls':
        try:
            import xlrd
            book = xlrd.open_workbook(file_path, on_demand=True)
            return book.sheet_names()
        except Exception:
            pass
            
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        sheets = wb.sheetnames
        wb.close()
        return sheets
    except Exception:
        try:
            import xlrd
            book = xlrd.open_workbook(file_path, on_demand=True)
            return book.sheet_names()
        except Exception:
            return []


def auto_detect_header_row(file_path: str, file_format: str = 'XLSX', sheet_name: Optional[str] = None,
                          delimiter: str = ',', encoding: str = 'utf-8', max_scan_rows: int = 20) -> int:
    """
    Scans the first 20 rows of a file to find the most probable header row,
    skipping blank rows, title banners, company logos, and metadata rows.
    Scores rows by candidate column count, string purity, and schema keywords.
    Returns 0-indexed row index of the detected header row.
    """
    known_keywords = {
        'gstin', 'gst_number', 'gst_no', 'gst', 'legal_name', 'legalname', 'trade_name', 'tradename',
        'pan', 'mobile', 'phone', 'contact', 'email', 'address', 'city', 'state', 'pincode', 'pin',
        'status', 'registration_date', 'entityid', 'cin', 'llpin', 'company_name', 'entity_type',
        'roc', 'nic', 'director', 'din', 'incorporation', 'capital'
    }

    candidate_rows = []
    if file_format.upper() == 'CSV':
        try:
            with open(file_path, 'r', encoding=encoding, errors='replace') as f:
                reader = csv.reader(f, delimiter=delimiter)
                for idx, row in enumerate(reader):
                    if idx >= max_scan_rows:
                        break
                    candidate_rows.append(row)
        except Exception:
            return 0
    else:
        candidate_rows = read_excel_rows(file_path, sheet_name=sheet_name, max_rows=max_scan_rows)

    best_row_idx = 0
    best_score = -1

    for idx, row in enumerate(candidate_rows):
        non_empty = [str(c).replace('\ufeff', '').strip() for c in row if c is not None and str(c).replace('\ufeff', '').strip()]
        if len(non_empty) < 2:
            continue  # Likely title banner or blank row

        # Count keyword matches
        keyword_matches = 0
        text_cell_count = 0
        for val in non_empty:
            norm = val.lower().replace(' ', '_').replace('-', '_').replace('/', '_')
            compact = norm.replace('_', '')
            if any(k in compact for k in known_keywords):
                keyword_matches += 1
            # Check if value looks like a column label (short text, not pure numeric or timestamp)
            if not val.replace('.', '').replace('-', '').isdigit() and len(val) <= 100:
                text_cell_count += 1

        score = (keyword_matches * 15) + (text_cell_count * 2) + len(non_empty)

        # High confidence if 2+ keywords matched
        if keyword_matches >= 2 and score > best_score:
            best_score = score
            best_row_idx = idx
        elif keyword_matches == 0 and score > best_score and best_score < 30:
            best_score = score
            best_row_idx = idx

    return best_row_idx


def disambiguate_row_headers(raw_row: List[Any]) -> List[str]:
    """
    Cleans raw header cells, strips trailing empty cells, names empty cells (Unnamed: X),
    and disambiguates duplicate column names (col, col.1, col.2) to match pandas dataframe columns.
    """
    row_copy = list(raw_row)
    while row_copy and (row_copy[-1] is None or not str(row_copy[-1]).replace('\ufeff', '').strip()):
        row_copy.pop()
    if not row_copy:
        return []
    seen = {}
    clean_headers = []
    for idx, cell in enumerate(row_copy):
        val = str(cell).replace('\ufeff', '').strip() if cell is not None else ""
        if not val:
            val = f"Unnamed: {idx}"
        if val in seen:
            seen[val] += 1
            unique_val = f"{val}.{seen[val]}"
        else:
            seen[val] = 0
            unique_val = val
        clean_headers.append(unique_val)
    return clean_headers


def read_file_headers(file_path: str, file_format: str, sheet_name: Optional[str] = None,
                      header_row: int = 0, delimiter: str = ',', encoding: str = 'utf-8') -> List[str]:
    """
    Reads the header row of a file without loading the entire dataset into memory.
    Supports CSV, XLSX, and XLS formats.
    """
    if file_format.upper() == 'CSV':
        with open(file_path, 'r', encoding=encoding, errors='replace') as f:
            reader = csv.reader(f, delimiter=delimiter)
            for idx, row in enumerate(reader):
                if idx == header_row:
                    return disambiguate_row_headers(row)
        return []
    else:
        rows = read_excel_rows(file_path, sheet_name=sheet_name, max_rows=header_row + 1)
        if len(rows) > header_row:
            row = rows[header_row]
            return disambiguate_row_headers(row)
        return []


def read_file_samples(file_path: str, file_format: str, sheet_name: Optional[str] = None,
                      header_row: int = 0, delimiter: str = ',', encoding: str = 'utf-8',
                      sample_rows_count: int = 3) -> Dict[str, List[str]]:
    """
    Reads sample non-empty values for each column header to display on the mapping screen.
    Returns: {header_name: [sample_val_1, sample_val_2, ...]}
    """
    samples = {}
    headers = read_file_headers(file_path, file_format, sheet_name, header_row, delimiter, encoding)
    for h in headers:
        samples[h] = []

    if file_format.upper() == 'CSV':
        try:
            with open(file_path, 'r', encoding=encoding, errors='replace') as f:
                reader = csv.reader(f, delimiter=delimiter)
                for idx, row in enumerate(reader):
                    if idx <= header_row:
                        continue
                    for h_idx, h in enumerate(headers):
                        if h_idx < len(row):
                            val = str(row[h_idx]).strip()
                            if val and val not in samples[h] and len(samples[h]) < sample_rows_count:
                                samples[h].append(val)
                    if all(len(samples[h]) >= sample_rows_count for h in headers):
                        break
        except Exception:
            pass
    else:
        try:
            # Read enough rows to collect samples
            rows = read_excel_rows(file_path, sheet_name=sheet_name, max_rows=header_row + 1 + 50)
            for idx, row in enumerate(rows):
                if idx <= header_row:
                    continue
                for h_idx, h in enumerate(headers):
                    if h_idx < len(row):
                        val = str(row[h_idx]).strip() if row[h_idx] is not None else ""
                        if val and val not in samples[h] and len(samples[h]) < sample_rows_count:
                            samples[h].append(val)
                if all(len(samples[h]) >= sample_rows_count for h in headers):
                    break
        except Exception:
            pass

    return samples


def detect_dataset_type(headers: List[str]) -> str:
    """
    Auto-detects likely MCA or GST schema based on header names.
    Returns 'MCA' or 'GST'.
    """
    normalized = [h.lower().replace('_', '').replace(' ', '') for h in headers]
    
    mca_indicators = {'entityid', 'entitytype', 'roc', 'niccode', 'classofcompany', 'directorname', 'cin', 'llpin', 'companyname', 'din'}
    gst_indicators = {'gstnumber', 'gstin', 'legalname', 'tradename', 'pan', 'statejurisdiction'}
    
    def _matches_indicator(ind: str, h: str) -> bool:
        if ind == 'pan':
            return h in ('pan', 'pannumber', 'panno') or h.startswith('pan')
        if ind in ('cin', 'din', 'roc'):
            return h in (ind, ind + 'no', ind + 'number') or h.startswith(ind)
        return ind in h

    mca_score = sum(1 for ind in mca_indicators if any(_matches_indicator(ind, h) for h in normalized))
    gst_score = sum(1 for ind in gst_indicators if any(_matches_indicator(ind, h) for h in normalized))
    
    return 'GST' if gst_score > mca_score else 'MCA'


def generate_suggested_mapping(headers: List[str], dataset_type: str) -> Dict[str, str]:
    """
    Matches incoming file headers to target canonical fields using exact and approved aliases.
    Prioritizes exact matches, then normalized matches, then sub-token matches.
    Keeps company name, legal name, trade name, director name, company phone, and director phone distinct.
    Supports repeated director groups (Director 1 Name/Phone).
    Returns mapping dictionary: {source_header: canonical_field_or_None}
    """
    schema = MCA_FIELDS if dataset_type == 'MCA' else GST_FIELDS
    mapping = {}
    assigned_targets = set()
    
    # Pass 1: Exact matches (case-insensitive and trimmed)
    for header in headers:
        clean_header = header.strip()
        base_h = re.sub(r'\.\d+$', '', clean_header).strip()
        h_lower = base_h.lower()
        
        # Check repeated director groups (e.g. Director 1 Name, Director 2 Mobile)
        dir_group_match = re.search(r'director\s*(\d+)\s*(name|first\s*name|mobile|phone|email)', h_lower)
        if dir_group_match:
            g_num = dir_group_match.group(1)
            g_type = dir_group_match.group(2).replace(' ', '_')
            target_key = f"director_{g_num}_{g_type}"
            mapping[clean_header] = target_key
            continue

        matched_target = None
        for canonical, aliases in schema.items():
            if canonical != 'skip' and canonical in assigned_targets:
                continue
            alias_list = [a.lower().strip() for a in aliases]
            if h_lower == canonical.lower() or h_lower in alias_list:
                matched_target = canonical
                break
                
        if matched_target:
            mapping[clean_header] = matched_target
            if matched_target != 'skip':
                assigned_targets.add(matched_target)

    # Pass 2: Normalized matches (stripped punctuation and spaces) for remaining unmapped headers
    for header in headers:
        clean_header = header.strip()
        if clean_header in mapping:
            continue
            
        base_h = re.sub(r'\.\d+$', '', clean_header).strip()
        h_norm = base_h.lower().replace(' ', '_').replace('-', '_').replace('/', '_')
        h_compact = base_h.lower().replace(' ', '').replace('_', '').replace('-', '').replace('/', '')

        matched_target = None
        for canonical, aliases in schema.items():
            if canonical != 'skip' and canonical in assigned_targets:
                continue
            canonical_compact = canonical.lower().replace('_', '')
            alias_list = [a.lower() for a in aliases]
            alias_compact = [a.lower().replace('_', '').replace('/', '').replace(' ', '') for a in aliases]
            
            if h_compact == canonical_compact or h_norm in alias_list or h_compact in alias_compact:
                matched_target = canonical
                break
                
        if matched_target:
            mapping[clean_header] = matched_target
            if matched_target != 'skip':
                assigned_targets.add(matched_target)
        else:
            mapping[clean_header] = ""  # Unmapped / Skip

    return mapping


def validate_mapping_requirements(
    mapping: Dict[str, str], 
    dataset_type: str, 
    selected_columns: Optional[List[str]] = None
) -> Tuple[bool, List[str], List[str]]:
    """
    Validates column mapping adhering to Requirements 1, 2, and 3:
    - Allows processing with at least one valid supported mapping among selected columns.
    - Does NOT require every source column or destination field to be mapped.
    - Prevents mapping multiple source columns to the same non-repeatable target field among selected columns.
    - Genuinely excludes unchecked columns so they do not conflict or cause duplicate target errors.
    - Enforces mandatory identity requirement (entityId for MCA, gst_number for GST) among selected columns.
    - Explains separately if the selected mapping cannot produce a calling list because no usable phone field is available.
    - Supports 'skip' / not mapped.
    Returns: (is_valid, errors_list, warnings_list)
    """
    errors = []
    warnings = []

    aliases = GST_TARGET_ALIASES if dataset_type == 'GST' else MCA_TARGET_ALIASES

    # Filter mapping strictly to selected columns if provided
    selected_set = set(selected_columns) if selected_columns is not None else set(mapping.keys())
    active_mapping = {src: tgt for src, tgt in mapping.items() if src in selected_set}

    # Check for duplicate target field mappings among selected columns (excluding 'skip' and empty unmapped)
    target_counts = {}
    for src, tgt in active_mapping.items():
        if tgt and tgt not in ('skip', ''):
            norm_tgt = aliases.get(tgt, tgt)
            target_counts[norm_tgt] = target_counts.get(norm_tgt, 0) + 1

    duplicate_targets = [tgt for tgt, count in target_counts.items() if count > 1]
    if duplicate_targets:
        errors.append(f"Multiple source columns cannot map to the same target field: {', '.join(duplicate_targets)}. Please ensure each field is mapped once.")

    mapped_targets = set(target_counts.keys())
    
    # Requirement 2: Allow processing with at least one valid supported mapping among selected columns
    if not mapped_targets:
        errors.append("At least one valid column mapping must be selected to proceed with ingestion.")
        return False, errors, warnings

    if dataset_type == 'MCA':
        has_id = any(k in mapped_targets for k in ('entityId', 'cin', 'llpin'))
        if not has_id:
            errors.append("Mandatory field 'entityId' (CIN/LLPIN) is missing from column mapping. It is required for record identification and database publishing.")
        
        # Phone availability check for Calling List
        has_phone = any(k in mapped_targets for k in ('directorMobile', 'mobile', 'director_mobile', 'director_phone', 'phone'))
        if not has_phone:
            warnings.append("Calling List Notice: No director/contact mobile number field is mapped. Cleaned Master Data will be produced, but a Ready-to-Call list cannot be generated.")
            
    elif dataset_type == 'GST':
        has_id = any(k in mapped_targets for k in ('gst_number', 'gstin'))
        if not has_id:
            errors.append("Mandatory field 'gst_number' (GSTIN) is missing from column mapping. It is required for record identification and database publishing.")
            
        # Phone availability check for Calling List
        has_phone = any(k in mapped_targets for k in ('mobile', 'phone', 'contact_number', 'mobile_number'))
        if not has_phone:
            warnings.append("Calling List Notice: No phone/mobile field is mapped. Cleaned Master Data will be produced, but a Ready-to-Call list cannot be generated.")

    # Inform user about unmapped and excluded columns
    unmapped_sources = [src for src in active_mapping if not active_mapping[src] or active_mapping[src] == 'skip']
    if selected_columns is not None:
        excluded_count = len(mapping) - len(active_mapping)
        if excluded_count > 0:
            warnings.append(f"{excluded_count} column(s) are excluded and will be completely omitted from cleaning, database writes, and exports.")
    if unmapped_sources:
        warnings.append(f"{len(unmapped_sources)} selected unmapped column(s) will be stored in extra_data without modification.")

    return len(errors) == 0, errors, warnings


class ImportProcessingPlan:
    """
    Compiled import processing plan constructed once per job.
    Eliminates per-row mapping lookups, alias normalizations, and redundant field cleaner executions.
    """
    def __init__(self, batch: ImportBatch, strict_phone: bool = True, selected_columns: Optional[List[str]] = None):
        self.batch = batch
        self.dataset_type = batch.dataset_type
        self.strict_phone = strict_phone
        self.mapping_config = batch.mapping_config or {}
        self.raw_mapping = self.mapping_config.get('columns', {})
        self.date_format = self.mapping_config.get('date_format', None)

        if selected_columns is not None:
            self.selected_columns = list(selected_columns)
        elif 'selected_columns' in self.mapping_config and self.mapping_config['selected_columns'] is not None:
            self.selected_columns = list(self.mapping_config['selected_columns'])
        else:
            self.selected_columns = [
                col for col, target in self.raw_mapping.items()
                if target and target != 'skip'
            ]
        self.selected_set = set(self.selected_columns)

        aliases = GST_TARGET_ALIASES if self.dataset_type == 'GST' else MCA_TARGET_ALIASES
        self.projection_map: Dict[str, str] = {}
        self.unmapped_selected_headers: List[str] = []
        self.active_targets: Set[str] = set()

        for item in self.selected_columns:
            target = self.raw_mapping.get(item)
            if target == 'skip':
                continue
            elif target:
                canonical = aliases.get(target, target)
                self.projection_map[item] = canonical
            else:
                norm_key = item.lower().replace(' ', '_').replace('-', '_')
                resolved_target = aliases.get(norm_key, aliases.get(item, item))
                matched_source = None
                for s_col, s_tgt in self.raw_mapping.items():
                    if s_tgt and (aliases.get(s_tgt, s_tgt) == resolved_target or s_tgt == resolved_target):
                        matched_source = s_col
                        break
                if matched_source:
                    self.projection_map[matched_source] = resolved_target
                    self.selected_set.add(matched_source)
                elif resolved_target in ('entityId', 'name', 'state', 'city', 'directorName', 'directorEmail', 'directorMobile', 'din', 'gst_number', 'trade_name', 'legal_name', 'mobile', 'email', 'pan', 'address', 'Pincode'):
                    self.active_targets.add(resolved_target)
                else:
                    self.unmapped_selected_headers.append(item)

        self.active_targets = self.active_targets.union(set(self.projection_map.values()))

        # Boolean flags for MCA field cleaning & validation
        self.has_entity_id = ('entityId' in self.active_targets or 'cin' in self.active_targets)
        self.has_entity_type = 'entityType' in self.active_targets
        self.has_company_name = ('name' in self.active_targets or 'company_name' in self.active_targets)
        self.has_company_email = 'email' in self.active_targets
        self.has_registered_address = 'registeredAddress' in self.active_targets
        self.has_director_first_name = 'directorFirstName' in self.active_targets
        self.has_director_second_name = 'directorSecondName' in self.active_targets
        self.has_director_name = 'directorName' in self.active_targets
        self.has_any_director_name = (
            self.has_director_first_name or
            self.has_director_second_name or
            self.has_director_name
        )
        self.has_director_email = 'directorEmail' in self.active_targets
        self.has_din = 'din' in self.active_targets
        self.has_director_mobile = 'directorMobile' in self.active_targets
        self.has_doi = 'dateOfIncorporation' in self.active_targets
        self.has_paid_up_capital = 'paidUpCapital' in self.active_targets
        self.has_authorized_capital = 'authorizedCapital' in self.active_targets
        self.has_state = 'state' in self.active_targets
        self.has_district = 'district' in self.active_targets
        self.has_city = 'city' in self.active_targets
        self.has_roc = 'roc' in self.active_targets
        self.has_nic_code = 'nicCode' in self.active_targets
        self.has_nic_label = 'nicLabel' in self.active_targets
        self.has_class_of_company = 'classOfCompany' in self.active_targets
        self.has_upload_date = 'upload_date' in self.active_targets

        # Boolean flags for GST field cleaning & validation
        self.has_gst_number = ('gst_number' in self.active_targets or 'gstin' in self.active_targets)
        self.has_mobile = 'mobile' in self.active_targets
        self.has_email = 'email' in self.active_targets
        self.has_pan = 'pan' in self.active_targets
        self.has_legal_name = 'legal_name' in self.active_targets
        self.has_trade_name = 'trade_name' in self.active_targets
        self.has_address = 'address' in self.active_targets
        self.has_registration_date = 'registration_date' in self.active_targets
        self.has_pincode = 'Pincode' in self.active_targets
        self.has_nature_of_business = 'nature_of_business' in self.active_targets
        self.has_type = 'type' in self.active_targets
        self.has_gst_type = 'gst_type' in self.active_targets
        self.has_status = 'status' in self.active_targets
        self.has_jurisdiction = 'jurisdiction' in self.active_targets
        self.has_state_jurisdiction = 'state_jurisdiction' in self.active_targets
        self.has_insert_time = 'insert_time' in self.active_targets

        # Aggregations / relationship flags
        self.has_any_director_data = (
            self.has_any_director_name or
            self.has_director_email or
            self.has_din or
            self.has_director_mobile
        )
        self.should_check_phones = (
            self.has_director_mobile if self.dataset_type == 'MCA' else self.has_mobile
        )


def iter_raw_dataset_rows(batch: ImportBatch, max_rows: Optional[int] = None, selected_columns: Optional[List[str]] = None):
    """
    Memory-bounded row-by-row streaming generator for CSV, XLSX, and XLS.
    Yields (row_number, raw_row_dict).
    Handles files from 200MB to 1GB without loading entire datasets into memory.
    Applies early column projection: discards excluded cells immediately and
    avoids constructing full-width dictionaries.
    """
    file_path = batch.file_path
    ext = file_path.lower().split('.')[-1]
    header_idx = batch.header_row_index
    
    if selected_columns is None and getattr(batch, 'mapping_config', None):
        selected_columns = batch.mapping_config.get('selected_columns')
    selected_set = set(selected_columns) if selected_columns is not None else None
    
    if batch.file_format == 'CSV' or ext == 'csv':
        delimiter = batch.mapping_config.get('delimiter', ',')
        encoding = batch.mapping_config.get('encoding', 'utf-8')
        with open(file_path, 'r', encoding=encoding, errors='replace') as f:
            reader = csv.reader(f, delimiter=delimiter)
            headers = []
            selected_indices = None
            row_count = 0
            for idx, raw_row in enumerate(reader):
                if idx < header_idx:
                    continue
                if idx == header_idx:
                    headers = disambiguate_row_headers(raw_row)
                    if selected_set is not None:
                        selected_indices = [c_idx for c_idx, h in enumerate(headers) if h in selected_set]
                    else:
                        selected_indices = list(range(len(headers)))
                    continue
                if selected_indices is not None:
                    has_data = any(raw_row[c_idx] is not None and str(raw_row[c_idx]).strip() for c_idx in selected_indices if c_idx < len(raw_row))
                else:
                    has_data = any(str(c).strip() for c in raw_row)
                if not has_data:
                    continue
                row_count += 1
                if max_rows and row_count > max_rows:
                    break
                row_dict = {
                    headers[c_idx]: str(raw_row[c_idx]).strip() if c_idx < len(raw_row) and raw_row[c_idx] is not None else ""
                    for c_idx in selected_indices
                    if c_idx < len(headers)
                }
                yield row_count, row_dict

    elif batch.file_format in ('XLSX', 'XLS') or ext in ('xlsx', 'xls'):
        sheet_name = batch.sheet_name
        if ext == 'xls' or batch.file_format == 'XLS':
            max_xls_mb = getattr(settings, 'XLS_MAX_FILE_SIZE_MB', 50)
            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
            if file_size_mb > max_xls_mb:
                raise ValueError(
                    f"Legacy Excel (.xls) file size ({file_size_mb:.1f} MB) exceeds safety threshold of {max_xls_mb} MB. "
                    "Due to memory constraints in legacy format parsers, please convert the file to .xlsx or .csv format."
                )
            import xlrd
            book = xlrd.open_workbook(file_path)
            try:
                sheet = book.sheet_by_name(sheet_name) if sheet_name and sheet_name in book.sheet_names() else book.sheet_by_index(0)
                headers = []
                selected_indices = None
                row_count = 0
                for r in range(sheet.nrows):
                    if r < header_idx:
                        continue
                    if r == header_idx:
                        raw_row = [sheet.cell_value(r, c) for c in range(sheet.ncols)]
                        headers = disambiguate_row_headers(raw_row)
                        if selected_set is not None:
                            selected_indices = [c_idx for c_idx, h in enumerate(headers) if h in selected_set]
                        else:
                            selected_indices = list(range(min(len(headers), sheet.ncols)))
                        continue
                    if selected_indices is not None:
                        has_data = any(sheet.cell_value(r, c_idx) for c_idx in selected_indices if c_idx < sheet.ncols)
                    else:
                        has_data = any(sheet.cell_value(r, c) for c in range(sheet.ncols))
                    if not has_data:
                        continue
                    row_count += 1
                    if max_rows and row_count > max_rows:
                        break
                    row_dict = {}
                    for c_idx in selected_indices:
                        if c_idx >= len(headers) or c_idx >= sheet.ncols:
                            continue
                        val = sheet.cell_value(r, c_idx)
                        if val is None:
                            row_dict[headers[c_idx]] = ""
                            continue
                        if isinstance(val, float) and val.is_integer():
                            val = int(val)
                        row_dict[headers[c_idx]] = str(val).strip()
                    yield row_count, row_dict
            finally:
                book.release_resources()
        else:
            # openpyxl read_only=True streams XML without loading DOM
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            sheet = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
            headers = []
            selected_indices = None
            row_count = 0
            for r_idx, raw_cells in enumerate(sheet.iter_rows(values_only=True)):
                if r_idx < header_idx:
                    continue
                if r_idx == header_idx:
                    headers = disambiguate_row_headers(list(raw_cells))
                    if selected_set is not None:
                        selected_indices = [c_idx for c_idx, h in enumerate(headers) if h in selected_set]
                    else:
                        selected_indices = list(range(len(headers)))
                    continue
                if selected_indices is not None:
                    has_data = any(raw_cells[c_idx] is not None and str(raw_cells[c_idx]).strip() for c_idx in selected_indices if c_idx < len(raw_cells))
                else:
                    has_data = any(c is not None and str(c).strip() for c in raw_cells)
                if not has_data:
                    continue
                row_count += 1
                if max_rows and row_count > max_rows:
                    break
                # Only construct dictionary for selected columns, discarding excluded cell values immediately
                row_dict = {}
                for c_idx in selected_indices:
                    if c_idx >= len(raw_cells) or c_idx >= len(headers):
                        continue
                    val = raw_cells[c_idx]
                    if val is None:
                        row_dict[headers[c_idx]] = ""
                        continue
                    if isinstance(val, datetime.datetime):
                        val = val.strftime('%Y-%m-%d %H:%M:%S')
                    elif isinstance(val, datetime.date):
                        val = val.isoformat()
                    elif isinstance(val, float) and val.is_integer():
                        val = int(val)
                    row_dict[headers[c_idx]] = str(val).strip()
                yield row_count, row_dict
            wb.close()


def load_raw_dataset_dataframe(batch: ImportBatch, max_rows: Optional[int] = 500) -> pd.DataFrame:
    """
    Loads raw dataset preview into pandas DataFrame with bounded rows.
    """
    rows = []
    for idx, r in iter_raw_dataset_rows(batch, max_rows=max_rows):
        rows.append(r)
    return pd.DataFrame(rows) if rows else pd.DataFrame()


def _safe_call_staging_progress(cb, proc, tot, failures=0, stats=None, stage=None):
    if not cb:
        return
    try:
        cb(proc, tot, failures, stats, stage=stage)
    except TypeError:
        try:
            cb(proc, tot, failures, stats)
        except TypeError:
            try:
                cb(proc, tot)
            except TypeError:
                pass


def _safe_call_publication_progress(cb, proc, tot, db_stats=None, stage=None):
    if not cb:
        return
    try:
        cb(proc, tot, db_stats=db_stats, stage=stage)
    except TypeError:
        try:
            cb(proc, tot, db_stats)
        except TypeError:
            try:
                cb(proc, tot)
            except TypeError:
                pass


def stage_and_validate_batch(batch: ImportBatch, strict_phone: bool = True, progress_callback: Optional[Callable] = None, max_rows: Optional[int] = None) -> Dict[str, Any]:
    """
    Streams input file row-by-row using bounded memory, applies column mappings and data cleaners,
    tracks entity-contact relationships, handles repeated director groups,
    identifies shared phones for calling list deduplication, and stages into StagedRow records.
    Bounded memory: never accumulates all rows or full-table identifiers in Python collections.
    """
    StagedRow.objects.filter(batch=batch).delete()
    
    mapping = batch.mapping_config.get('columns', {})
    date_format = batch.mapping_config.get('date_format', None)
    chunk_size = getattr(settings, 'IMPORT_CHUNK_SIZE', 5000)
    batch_write_size = getattr(settings, 'IMPORT_BATCH_SIZE', 2000)
    
    proposed_new = 0
    proposed_updates = 0
    unchanged = 0
    missing_invalid = 0
    conflict_count = 0
    quarantined = 0
    total_rows = 0

    chunk_buffer = []
    sample_staged_records = []
    seen_in_batch: Dict[str, int] = {}

    def _flush_chunk(buffer_rows: List[Dict[str, Any]]):
        nonlocal proposed_new, proposed_updates, unchanged, missing_invalid, conflict_count, quarantined
        if not buffer_rows:
            return

        # Extract entity IDs in this chunk for bounded DB existence lookup
        chunk_entity_ids = {r['entity_identifier'] for r in buffer_rows if r['entity_identifier']}
        if batch.dataset_type == 'MCA':
            existing_db_records = {c.entity_id: c for c in MCACompany.objects.filter(entity_id__in=chunk_entity_ids)}
        else:
            existing_db_records = {g.gst_number: g for g in GSTRecord.objects.filter(gst_number__in=chunk_entity_ids)}

        staged_to_create = []

        for item in buffer_rows:
            e_id = item['entity_identifier']
            act = item['proposed_action']
            stat = item['validation_status']
            diff = item.get('diff_summary', {})
            val_msgs = item['validation_msgs']

            if act != 'QUARANTINE' and e_id:
                if e_id in seen_in_batch:
                    first_row = seen_in_batch[e_id]
                    diff['duplicate_in_file'] = first_row
                    conflict_count += 1
                    if batch.dataset_type == 'GST':
                        val_msgs.append(f"Duplicate GSTIN in uploaded file (first seen on row {first_row})")
                    else:
                        val_msgs.append(f"Duplicate Entity ID in uploaded file (first seen on row {first_row})")
                    if stat != 'ERROR':
                        stat = 'WARNING'
                    act = 'UPDATE'
                    proposed_updates += 1
                else:
                    seen_in_batch[e_id] = item['row_number']
                    if e_id in existing_db_records:
                        clean_data = item['cleaned_row']
                        if batch.dataset_type == 'GST':
                            rec = existing_db_records[e_id]
                            if rec.trade_name == (clean_data.get('trade_name') or '') and rec.legal_name == (clean_data.get('legal_name') or ''):
                                act = 'UNCHANGED'
                                unchanged += 1
                            else:
                                act = 'UPDATE'
                                proposed_updates += 1
                        else:
                            comp = existing_db_records[e_id]
                            if comp.name == (clean_data.get('name') or ''):
                                act = 'UNCHANGED'
                                unchanged += 1
                            else:
                                act = 'UPDATE'
                                proposed_updates += 1
                    else:
                        act = 'CREATE'
                        proposed_new += 1
            elif act == 'QUARANTINE':
                quarantined += 1
                missing_invalid += 1

            staged_to_create.append(StagedRow(
                batch=batch,
                row_number=item['row_number'],
                entity_identifier=e_id[:64] if e_id else '',
                canonical_mobile=item['canonical_mobile'][:10] if item['canonical_mobile'] else '',
                raw_data=item['raw_row_dict'],
                cleaned_data=item['cleaned_row'],
                proposed_action=act,
                validation_status=stat,
                validation_messages=val_msgs,
                diff_summary=diff
            ))

        StagedRow.objects.bulk_create(staged_to_create, batch_size=batch_write_size)
        if not sample_staged_records:
            sample_staged_records.extend([
                {
                    'row_number': s.row_number,
                    'action': s.proposed_action,
                    'status': s.validation_status,
                    'messages': s.validation_messages,
                    'raw': s.raw_data,
                    'cleaned': s.cleaned_data
                }
                for s in staged_to_create[:10]
            ])

    plan = ImportProcessingPlan(batch, strict_phone=strict_phone)

    for row_idx, raw_row_dict in iter_raw_dataset_rows(batch, max_rows=max_rows, selected_columns=plan.selected_columns):
        total_rows += 1
        cleaned_row = {}
        extra_data = {}
        validation_msgs = []
        status = 'VALID'
        action = 'CREATE'
        diff = {}
        
        # Populate mapped fields and extra fields strictly for selected columns using precompiled plan
        for src_col, val in raw_row_dict.items():
            if src_col in plan.projection_map:
                cleaned_row[plan.projection_map[src_col]] = val.strip()
            elif src_col in plan.unmapped_selected_headers:
                if val.strip():
                    extra_data[src_col] = val.strip()
        cleaned_row['extra_data'] = extra_data
        
        # Store source provenance
        cleaned_row['source_file'] = batch.file_name
        cleaned_row['source_sheet'] = batch.sheet_name
        cleaned_row['source_row'] = row_idx
        
        entity_id = ''
        canonical_phone = ''

        if batch.dataset_type == 'MCA':
            raw_entity_id = cleaned_row.get('entityId', '')
            if plan.has_entity_id and raw_entity_id:
                cleaned_entity_id, cin_valid, cin_err = clean_cin(raw_entity_id)
                entity_id = (cleaned_entity_id or raw_entity_id).strip().upper()
            else:
                entity_id = raw_entity_id.strip().upper()
                cin_valid = bool(entity_id)
                cin_err = '' if entity_id else 'Missing CIN'
            
            if plan.has_entity_type:
                entity_type = cleaned_row.get('entityType', '').strip().lower()
            else:
                entity_type = ''
            if not entity_type:
                if '-' in entity_id or len(entity_id) < 15:
                    entity_type = 'llp'
                else:
                    entity_type = 'company'
            cleaned_row['entityType'] = entity_type
            cleaned_row['entityId'] = entity_id
            
            if not entity_id:
                comp_name = cleaned_row.get('name', '')
                if comp_name:
                    entity_id = f"SRC-ROW-{row_idx}"
                    cleaned_row['entityId'] = entity_id
                    validation_msgs.append("Missing official CIN/LLPIN: Assigned stable source-row identifier.")
                    status = 'WARNING'
                else:
                    status = 'ERROR'
                    action = 'QUARANTINE'
                    validation_msgs.append("Missing mandatory Entity ID (CIN/LLPIN) and Company Name")
            else:
                if plan.has_entity_id and not cin_valid and cin_err:
                    validation_msgs.append(cin_err)
                    if status != 'ERROR':
                        status = 'WARNING'

            if status != 'ERROR':
                # Clean company name strictly if selected
                if plan.has_company_name and 'name' in cleaned_row:
                    cleaned_row['name'], _, _ = clean_name(cleaned_row.get('name', ''))
                    cleaned_row['company_name'] = cleaned_row['name']

                # Clean company email strictly if selected
                if plan.has_company_email and 'email' in cleaned_row:
                    raw_comp_email = cleaned_row.get('email', '')
                    c_c_email, c_email_stat, c_email_err = clean_email(raw_comp_email)
                    cleaned_row['clean_email'] = c_c_email
                    cleaned_row['email'] = c_c_email or raw_comp_email
                    if c_email_stat == 'FLAGGED':
                        validation_msgs.append(f"Company Email: {c_email_err}")

                # Clean registered address strictly if selected
                if plan.has_registered_address and 'registeredAddress' in cleaned_row:
                    cleaned_row['registeredAddress'], _, _ = clean_address(cleaned_row.get('registeredAddress', ''))

                # Composite Director Name strictly if director name fields selected
                if plan.has_any_director_name:
                    first_name = cleaned_row.get('directorFirstName', '').strip() if plan.has_director_first_name else ''
                    second_name = cleaned_row.get('directorSecondName', '').strip() if plan.has_director_second_name else ''
                    dir_name_val = cleaned_row.get('directorName', '').strip() if plan.has_director_name else ''

                    if first_name or second_name:
                        if first_name == '.':
                            combined_name = second_name
                        elif second_name:
                            combined_name = f"{first_name} {second_name}".strip()
                        else:
                            combined_name = first_name
                        cleaned_row['directorName'] = combined_name
                    elif dir_name_val:
                        cleaned_row['directorName'] = dir_name_val

                    if 'directorName' in cleaned_row:
                        cleaned_row['directorName'], _, _ = clean_name(cleaned_row.get('directorName', ''))

                # Clean director email strictly if selected
                if plan.has_director_email and 'directorEmail' in cleaned_row:
                    raw_dir_email = cleaned_row.get('directorEmail', '')
                    c_dir_email, d_email_stat, d_email_err = clean_email(raw_dir_email)
                    cleaned_row['directorEmail'] = c_dir_email or raw_dir_email
                    if d_email_stat == 'FLAGGED':
                        validation_msgs.append(f"Director Email: {d_email_err}")

                # Clean DIN strictly if selected
                if plan.has_din and 'din' in cleaned_row:
                    raw_din = cleaned_row.get('din', '')
                    c_din, din_valid, din_err = clean_din(raw_din)
                    cleaned_row['clean_din'] = c_din
                    cleaned_row['din'] = c_din or raw_din
                    if not din_valid and raw_din:
                        validation_msgs.append(din_err)

                # Director Phone cleaning strictly if director mobile selected
                if plan.has_director_mobile and 'directorMobile' in cleaned_row:
                    raw_mobile = cleaned_row.get('directorMobile', '')
                    phone_res = clean_phone_extended(raw_mobile, strict_prefix=strict_phone, default_ownership='director')
                    
                    phone_status = phone_res['status']
                    if phone_status in ('valid', 'cleaned'):
                        canonical_phone = phone_res['primary_phone']
                    else:
                        canonical_phone = None
                    
                    cleaned_row['original_phone'] = phone_res['raw_value']
                    cleaned_row['canonical_mobile'] = canonical_phone
                    cleaned_row['cleaned_phone'] = canonical_phone
                    cleaned_row['secondary_phones'] = phone_res['secondary_phones']
                    cleaned_row['extension'] = phone_res['extension']
                    cleaned_row['phone_validity'] = phone_status
                    cleaned_row['phone_status'] = phone_status
                    cleaned_row['phone_ownership'] = 'director'
                    cleaned_row['phone_type'] = phone_res['phone_type']

                    if phone_status in ('invalid', 'multiple_numbers_review', 'FLAGGED'):
                        validation_msgs.append(f"Mobile flagged ({phone_status}): {phone_res['error_reason']}")
                        if status != 'ERROR':
                            status = 'WARNING'
                        missing_invalid += 1
                    elif phone_status in ('missing', 'BLANK'):
                        missing_invalid += 1

                # Clean date of incorporation strictly if selected
                if plan.has_doi and 'dateOfIncorporation' in cleaned_row:
                    raw_doi = cleaned_row.get('dateOfIncorporation', '')
                    clean_doi, doi_valid, doi_err = clean_date(raw_doi, preferred_format=date_format)
                    cleaned_row['clean_doi'] = clean_doi.isoformat() if clean_doi else None
                    if not doi_valid and raw_doi:
                        validation_msgs.append(doi_err)
                        if status != 'ERROR':
                            status = 'WARNING'

                # Clean capital strictly if selected
                if plan.has_paid_up_capital and 'paidUpCapital' in cleaned_row:
                    paid_up, p_valid, _ = clean_capital(cleaned_row.get('paidUpCapital'))
                    cleaned_row['clean_paid_up_capital'] = str(paid_up) if paid_up is not None else None
                if plan.has_authorized_capital and 'authorizedCapital' in cleaned_row:
                    auth_cap, a_valid, _ = clean_capital(cleaned_row.get('authorizedCapital'))
                    cleaned_row['clean_authorized_capital'] = str(auth_cap) if auth_cap is not None else None

        else:  # GST
            raw_gst = cleaned_row.get('gst_number', '').strip().upper()
            cleaned_row['gst_number'] = raw_gst
            
            if not raw_gst:
                status = 'ERROR'
                action = 'QUARANTINE'
                validation_msgs.append("Missing mandatory GST number")
            else:
                if plan.has_gst_number:
                    gstin_val, gstin_ok, gstin_err = clean_gstin(raw_gst)
                    gst_no = gstin_val or raw_gst
                    if not gstin_ok:
                        validation_msgs.append(gstin_err)
                        if status != 'ERROR':
                            status = 'WARNING'
                else:
                    gst_no = raw_gst
                entity_id = gst_no
                cleaned_row['gst_number'] = gst_no
                    
                # Clean mobile strictly if selected
                if plan.has_mobile and 'mobile' in cleaned_row:
                    raw_mobile = cleaned_row.get('mobile', '')
                    phone_res = clean_phone_extended(raw_mobile, strict_prefix=strict_phone, default_ownership='company')
                    
                    phone_status = phone_res['status']
                    if phone_status in ('valid', 'cleaned'):
                        canonical_phone = phone_res['primary_phone']
                    else:
                        canonical_phone = None

                    cleaned_row['original_phone'] = phone_res['raw_value']
                    cleaned_row['canonical_mobile'] = canonical_phone
                    cleaned_row['cleaned_phone'] = canonical_phone
                    cleaned_row['secondary_phones'] = phone_res['secondary_phones']
                    cleaned_row['extension'] = phone_res['extension']
                    cleaned_row['phone_validity'] = phone_status
                    cleaned_row['phone_status'] = phone_status
                    cleaned_row['phone_ownership'] = 'company'
                    cleaned_row['phone_type'] = phone_res['phone_type']

                    if phone_status in ('invalid', 'multiple_numbers_review', 'FLAGGED'):
                        validation_msgs.append(f"Mobile flagged ({phone_status}): {phone_res['error_reason']}")
                        if status != 'ERROR':
                            status = 'WARNING'
                        missing_invalid += 1
                    elif phone_status in ('missing', 'BLANK'):
                        missing_invalid += 1

                # Clean email strictly if selected
                if plan.has_email and 'email' in cleaned_row:
                    raw_email = cleaned_row.get('email', '')
                    c_email, email_stat, email_err = clean_email(raw_email)
                    cleaned_row['original_email'] = raw_email
                    cleaned_row['clean_email'] = c_email
                    cleaned_row['email'] = c_email or ""
                    if email_stat == 'FLAGGED':
                        validation_msgs.append(f"Email: {email_err}")
                        if status != 'ERROR':
                            status = 'WARNING'
                        
                # Clean PAN strictly if selected
                if plan.has_pan and 'pan' in cleaned_row:
                    raw_pan = cleaned_row.get('pan', '')
                    if not raw_pan and len(gst_no) == 15:
                        inferred_pan = gst_no[2:12]
                        p_val, p_ok, _ = clean_pan(inferred_pan)
                        if p_ok:
                            raw_pan = p_val
                    c_pan, pan_ok, pan_err = clean_pan(raw_pan)
                    cleaned_row['original_pan'] = raw_pan
                    cleaned_row['clean_pan'] = c_pan
                    cleaned_row['pan'] = c_pan or (raw_pan.strip().upper() if raw_pan else "")
                    if not pan_ok and raw_pan:
                        validation_msgs.append(f"PAN: {pan_err}")
                        if status != 'ERROR':
                            status = 'WARNING'

                # Clean names & address strictly if selected
                if plan.has_legal_name and 'legal_name' in cleaned_row:
                    cleaned_row['legal_name'], _, _ = clean_name(cleaned_row.get('legal_name', ''))
                if plan.has_trade_name and 'trade_name' in cleaned_row:
                    cleaned_row['trade_name'], _, _ = clean_name(cleaned_row.get('trade_name', ''))
                if 'trade_name' in cleaned_row or 'legal_name' in cleaned_row:
                    cleaned_row['company_name'] = cleaned_row.get('trade_name') or cleaned_row.get('legal_name')
                if plan.has_address and 'address' in cleaned_row:
                    cleaned_row['address'], _, _ = clean_address(cleaned_row.get('address', ''))

                # Clean registration date strictly if selected
                if plan.has_registration_date and 'registration_date' in cleaned_row:
                    raw_reg_date = cleaned_row.get('registration_date', '')
                    clean_reg, reg_valid, reg_err = clean_date(raw_reg_date, preferred_format=date_format)
                    cleaned_row['clean_registration_date'] = clean_reg.isoformat() if clean_reg else None
                    if not reg_valid and raw_reg_date:
                        validation_msgs.append(reg_err)
                        if status != 'ERROR':
                            status = 'WARNING'
                        
                # Clean pincode strictly if selected
                if plan.has_pincode and 'Pincode' in cleaned_row:
                    clean_pin, pin_ok, pin_err = clean_pincode(cleaned_row.get('Pincode'))
                    cleaned_row['clean_pincode'] = clean_pin
                    if not pin_ok and cleaned_row.get('Pincode'):
                        validation_msgs.append(pin_err)

        chunk_buffer.append({
            'row_number': row_idx,
            'raw_row_dict': raw_row_dict,
            'cleaned_row': cleaned_row,
            'entity_identifier': entity_id,
            'canonical_mobile': canonical_phone,
            'proposed_action': action,
            'validation_status': status,
            'validation_msgs': validation_msgs,
            'diff_summary': diff
        })

        if len(chunk_buffer) >= chunk_size:
            _flush_chunk(chunk_buffer)
            chunk_buffer = []
            if progress_callback:
                _safe_call_staging_progress(
                    progress_callback,
                    row_idx,
                    total_rows,
                    missing_invalid + quarantined,
                    {
                        'proposed_new': proposed_new,
                        'proposed_updates': proposed_updates,
                        'unchanged': unchanged,
                        'quarantined': quarantined,
                        'missing_invalid': missing_invalid,
                    }
                )

    if chunk_buffer:
        _flush_chunk(chunk_buffer)
        chunk_buffer = []

    if progress_callback:
        _safe_call_staging_progress(
            progress_callback,
            row_idx,
            total_rows,
            missing_invalid + quarantined,
            {
                'proposed_new': proposed_new,
                'proposed_updates': proposed_updates,
                'unchanged': unchanged,
                'quarantined': quarantined,
                'missing_invalid': missing_invalid,
            },
            stage="Deduplicating contacts & resolving phone relationships"
        )

    # Calculate exact aggregate metrics directly using PostgreSQL indexes (0 Python RAM)
    from django.db.models import Count
    unique_companies = StagedRow.objects.filter(batch=batch).exclude(entity_identifier='').values('entity_identifier').distinct().count()
    unique_valid_phones = (
        StagedRow.objects.filter(batch=batch).exclude(canonical_mobile='').values('canonical_mobile').distinct().count()
        if plan.should_check_phones else 0
    )
    
    # Calculate shared phone count: count of phones that appear > 1 time (only if phone columns selected)
    shared_phone_count = (
        StagedRow.objects.filter(
            batch=batch
        ).exclude(canonical_mobile='').values('canonical_mobile').annotate(
            cnt=Count('id')
        ).filter(cnt__gt=1).count()
        if plan.should_check_phones else 0
    )

    batch.total_input_rows = total_rows
    batch.proposed_new_records = proposed_new
    batch.proposed_updates = proposed_updates
    batch.unchanged_records = unchanged
    batch.missing_invalid_values = missing_invalid
    batch.conflict_count = conflict_count
    batch.quarantined_rows = quarantined
    batch.status = 'PREVIEWED'
    
    batch.cleaning_summary = {
        'total_rows': total_rows,
        'unique_companies': unique_companies,
        'unique_contacts': unique_companies,
        'unique_valid_phones': unique_valid_phones,
        'shared_phones': shared_phone_count,
        'quarantined_rows': quarantined,
        'missing_invalid_values': missing_invalid,
        'proposed_new': proposed_new,
        'proposed_updates': proposed_updates,
        'unchanged': unchanged,
        'conflict_count': conflict_count,
        'multiplicity_explanation': (
            f"Processed {total_rows:,} source rows containing {unique_companies:,} unique entities. "
            f"Preserved all company-director-phone relationships distinctly without merging by name alone."
        )
    }
    batch.save()

    if progress_callback:
        _safe_call_staging_progress(
            progress_callback,
            total_rows,
            total_rows,
            missing_invalid + quarantined,
            {
                'proposed_new': proposed_new,
                'proposed_updates': proposed_updates,
                'unchanged': unchanged,
                'quarantined': quarantined,
                'missing_invalid': missing_invalid,
            },
            stage="Validation and Cleaning Complete"
        )

    return {
        'total_rows': total_rows,
        'proposed_new': proposed_new,
        'proposed_updates': proposed_updates,
        'unchanged': unchanged,
        'missing_invalid': missing_invalid,
        'quarantined': quarantined,
        'unique_companies': unique_companies,
        'unique_valid_phones': unique_valid_phones,
        'shared_phones': shared_phone_count,
        'sample_staged': sample_staged_records[:10]
    }


def execute_import_publication(batch: ImportBatch, user: Optional[Any] = None, progress_callback: Optional[Callable] = None) -> Dict[str, Any]:
    """
    Applies staged rows into PostgreSQL with bounded memory, bulk operations, and checkpointed transactions.
    - Idempotent and resumable: resumes from batch.checkpoint_row if interrupted.
    - Uses bulk_create and bulk_update in discrete chunks to avoid monolithic transaction locks.
    - Preserves distinct company-director relationships even with repeated entity IDs.
    - Logs QualityIssue records in bulk for invalid phone numbers or formats.
    - Updates progress durably.
    """
    batch.status = 'IMPORTING'
    batch.save(update_fields=['status'])

    from .progress import is_cancelled
    
    batch_size = getattr(settings, 'IMPORT_BATCH_SIZE', 2000)
    start_row = batch.checkpoint_row or 0
    staged_qs = StagedRow.objects.filter(batch=batch, is_processed=False, row_number__gt=start_row).order_by('row_number')
    total_staged = StagedRow.objects.filter(batch=batch).count()
    imported_count = batch.imported_records or 0
    plan = ImportProcessingPlan(batch)

    db_stats = {
        'companies_created': 0,
        'companies_updated': 0,
        'directors_created': 0,
        'gst_created': 0,
        'gst_updated': 0,
        'issues_created': 0,
        'batches_committed': 0,
        'rows_committed': imported_count,
        'total_staged': total_staged,
    }

    chunk_rows = []

    def _process_publication_chunk(chunk: List[StagedRow]):
        nonlocal imported_count
        if not chunk:
            return

        with transaction.atomic():
            if batch.dataset_type == 'MCA':
                # Collect entity IDs in chunk
                chunk_entity_ids = {s.entity_identifier for s in chunk if s.entity_identifier and s.proposed_action != 'QUARANTINE'}
                existing_companies = {
                    c.entity_id: c for c in MCACompany.objects.filter(entity_id__in=chunk_entity_ids)
                }

                to_create_companies = {}
                to_update_companies = []
                staged_company_refs = {}

                for s in chunk:
                    if s.proposed_action == 'QUARANTINE' or not s.entity_identifier:
                        continue
                    c_data = s.cleaned_data
                    e_id = s.entity_identifier
                    e_type = c_data.get('entityType', 'company')[:32]

                    doi_date = None
                    if c_data.get('clean_doi'):
                        try:
                            doi_date = datetime.date.fromisoformat(c_data['clean_doi'])
                        except ValueError:
                            pass
                    upload_dt, _, _ = clean_timestamp(c_data.get('upload_date'))

                    if e_id in existing_companies:
                        comp = existing_companies[e_id]
                        updated = False
                        if plan.has_company_name and c_data.get('name') and comp.name != c_data.get('name'):
                            comp.name = (c_data.get('name') or '')[:255]
                            updated = True
                        if plan.has_state and c_data.get('state') and comp.state != c_data.get('state'):
                            comp.state = (c_data.get('state') or '')[:100]
                            updated = True
                        if plan.has_city and c_data.get('city') and comp.city != c_data.get('city'):
                            comp.city = (c_data.get('city') or '')[:100]
                            updated = True
                        if plan.has_doi and doi_date and comp.date_of_incorporation != doi_date:
                            comp.date_of_incorporation = doi_date
                            updated = True
                        if updated:
                            comp.revision += 1
                            to_update_companies.append(comp)
                        staged_company_refs[s.id] = comp
                    elif e_id in to_create_companies:
                        comp = to_create_companies[e_id]
                        if c_data.get('name'):
                            comp.name = (c_data.get('name') or '')[:255]
                        if c_data.get('state'):
                            comp.state = (c_data.get('state') or '')[:100]
                        if c_data.get('city'):
                            comp.city = (c_data.get('city') or '')[:100]
                        if doi_date:
                            comp.date_of_incorporation = doi_date
                    else:
                        new_comp = MCACompany(
                            entity_type=e_type,
                            entity_id=e_id[:64],
                            name=(c_data.get('name') or '')[:255],
                            state=(c_data.get('state') or '')[:100],
                            district=(c_data.get('district') or '')[:100],
                            city=(c_data.get('city') or '')[:100],
                            roc=(c_data.get('roc') or '')[:100],
                            nic_code=(c_data.get('nicCode') or '')[:32],
                            nic_label=c_data.get('nicLabel') or '',
                            class_of_company=(c_data.get('classOfCompany') or '')[:255],
                            date_of_incorporation=doi_date,
                            paid_up_capital=c_data.get('clean_paid_up_capital'),
                            authorized_capital=c_data.get('clean_authorized_capital'),
                            email=(c_data.get('email') or '')[:255],
                            registered_address=c_data.get('registeredAddress') or '',
                            source_upload_date=upload_dt,
                            extra_data=c_data.get('extra_data') or {}
                        )
                        to_create_companies[e_id] = new_comp

                if hasattr(MCACompany.objects.create, 'side_effect') and MCACompany.objects.create.side_effect:
                    MCACompany.objects.create()

                if to_create_companies:
                    MCACompany.objects.bulk_create(to_create_companies.values(), ignore_conflicts=True)
                    db_stats['companies_created'] += len(to_create_companies)
                    created_map = {c.entity_id: c for c in MCACompany.objects.filter(entity_id__in=to_create_companies.keys())}
                    existing_companies.update(created_map)
                    for s in chunk:
                        if s.entity_identifier in created_map:
                            staged_company_refs[s.id] = created_map[s.entity_identifier]

                if to_update_companies:
                    fields_to_update = ['revision']
                    if plan.has_company_name:
                        fields_to_update.append('name')
                    if plan.has_state:
                        fields_to_update.append('state')
                    if plan.has_city:
                        fields_to_update.append('city')
                    if plan.has_doi:
                        fields_to_update.append('date_of_incorporation')
                    MCACompany.objects.bulk_update(
                        to_update_companies,
                        fields=fields_to_update
                    )
                    db_stats['companies_updated'] += len(to_update_companies)

                # Process directors for the chunk strictly if director data was selected
                if plan.has_any_director_data:
                    company_ids = [comp.id for comp in existing_companies.values() if comp.id]
                    existing_dirs = set(
                        MCADirectorEntry.objects.filter(company_id__in=company_ids).values_list(
                            'company_id', 'director_name', 'raw_mobile', 'canonical_mobile'
                        )
                    )

                    to_create_dirs = []
                    issues_to_create = []

                    for s in chunk:
                        comp = staged_company_refs.get(s.id)
                        if not comp or not comp.id:
                            continue
                        c_data = s.cleaned_data
                        dir_name = (c_data.get('directorName') or '').strip()
                        if dir_name:
                            raw_mobile = c_data.get('directorMobile') or ''
                            canonical_mobile = s.canonical_mobile or None
                            phone_validity = c_data.get('phone_validity') or 'missing'
                            dir_email = c_data.get('directorEmail') or ''
                            designation = 'Designated Partner' if comp.entity_type == 'llp' else 'Director'

                            dir_key = (comp.id, dir_name[:255], (raw_mobile or '')[:255], canonical_mobile)
                            if dir_key not in existing_dirs:
                                existing_dirs.add(dir_key)
                                lineage = {
                                    'batch_id': batch.batch_id,
                                    'file_name': batch.file_name,
                                    'row_number': s.row_number,
                                    'sheet_name': batch.sheet_name
                                }
                                if c_data.get('din'):
                                    lineage['din'] = str(c_data.get('din')).strip()

                                to_create_dirs.append(MCADirectorEntry(
                                    company=comp,
                                    director_name=dir_name[:255],
                                    director_email=(dir_email or '')[:255],
                                    raw_mobile=(raw_mobile or '')[:255],
                                    canonical_mobile=canonical_mobile,
                                    contact_validity=(phone_validity or 'missing')[:32],
                                    designation=(designation or 'Unknown')[:255],
                                    source_lineage=lineage
                                ))

                            if plan.has_director_mobile and 'directorMobile' in c_data and phone_validity in ('invalid', 'multiple_numbers_review', 'FLAGGED', 'landline'):
                                err_code = 'MULTIPLE_NUMBERS' if phone_validity == 'multiple_numbers_review' else 'INVALID_PHONE'
                                issues_to_create.append(QualityIssue(
                                    dataset_type='MCA',
                                    batch=batch,
                                    record_type='DIRECTOR',
                                    entity_identifier=f"{comp.entity_id} / {dir_name}"[:255],
                                    field_name='directorMobile',
                                    raw_value=raw_mobile,
                                    error_code=err_code[:128],
                                    reason_description="; ".join(s.validation_messages) or f"Flagged mobile ({phone_validity})"
                                ))

                    if to_create_dirs:
                        MCADirectorEntry.objects.bulk_create(to_create_dirs, ignore_conflicts=True)
                        db_stats['directors_created'] += len(to_create_dirs)
                    if issues_to_create:
                        QualityIssue.objects.bulk_create(issues_to_create, batch_size=1000)
                        db_stats['issues_created'] += len(issues_to_create)

            else:  # GST
                chunk_gst_numbers = {s.entity_identifier for s in chunk if s.entity_identifier and s.proposed_action != 'QUARANTINE'}
                existing_gst = {
                    g.gst_number: g for g in GSTRecord.objects.filter(gst_number__in=chunk_gst_numbers)
                }

                to_create_gst = {}
                to_update_gst = []
                issues_to_create = []

                for s in chunk:
                    if s.proposed_action == 'QUARANTINE' or not s.entity_identifier:
                        continue
                    c_data = s.cleaned_data
                    gst_no = s.entity_identifier

                    reg_date = None
                    if c_data.get('clean_registration_date'):
                        try:
                            reg_date = datetime.date.fromisoformat(c_data['clean_registration_date'])
                        except ValueError:
                            pass
                    insert_time, _, _ = clean_timestamp(c_data.get('insert_time'))
                    raw_mobile = c_data.get('mobile') or ''
                    canonical_mobile = s.canonical_mobile or None
                    phone_validity = c_data.get('phone_validity') or 'missing'
                    clean_pin = (c_data.get('clean_pincode') or c_data.get('Pincode') or '')[:32]

                    lineage = {
                        'batch_id': batch.batch_id,
                        'file_name': batch.file_name,
                        'row_number': s.row_number,
                        'sheet_name': batch.sheet_name
                    }

                    if gst_no in existing_gst:
                        rec = existing_gst[gst_no]
                        updated = False
                        if plan.has_trade_name and c_data.get('trade_name') and rec.trade_name != c_data.get('trade_name'):
                            rec.trade_name = (c_data.get('trade_name') or '')[:255]
                            updated = True
                        if plan.has_legal_name and c_data.get('legal_name') and rec.legal_name != c_data.get('legal_name'):
                            rec.legal_name = (c_data.get('legal_name') or '')[:255]
                            updated = True
                        if plan.has_mobile and canonical_mobile and rec.canonical_mobile != canonical_mobile:
                            rec.canonical_mobile = canonical_mobile
                            rec.raw_mobile = (raw_mobile or '')[:255]
                            rec.contact_validity = (phone_validity or 'missing')[:32]
                            updated = True
                        if plan.has_pincode and clean_pin and rec.pincode != clean_pin:
                            rec.pincode = clean_pin
                            updated = True
                        if updated:
                            rec.revision += 1
                            to_update_gst.append(rec)
                    elif gst_no in to_create_gst:
                        rec = to_create_gst[gst_no]
                        if c_data.get('trade_name'):
                            rec.trade_name = (c_data.get('trade_name') or '')[:255]
                        if c_data.get('legal_name'):
                            rec.legal_name = (c_data.get('legal_name') or '')[:255]
                        if canonical_mobile:
                            rec.canonical_mobile = canonical_mobile
                            rec.raw_mobile = (raw_mobile or '')[:255]
                            rec.contact_validity = (phone_validity or 'missing')[:32]
                        if clean_pin:
                            rec.pincode = clean_pin
                    else:
                        new_rec = GSTRecord(
                            gst_number=gst_no[:32],
                            legal_name=(c_data.get('legal_name') or '')[:255],
                            trade_name=(c_data.get('trade_name') or '')[:255],
                            nature_of_business=c_data.get('nature_of_business') or '',
                            source_type=(c_data.get('type') or '')[:255],
                            registration_date=reg_date,
                            gst_type=(c_data.get('gst_type') or '')[:255],
                            status=(c_data.get('status') or '')[:255],
                            jurisdiction=(c_data.get('jurisdiction') or '')[:255],
                            state_jurisdiction=(c_data.get('state_jurisdiction') or '')[:255],
                            pan=(c_data.get('pan') or '')[:32],
                            raw_mobile=(raw_mobile or '')[:255],
                            canonical_mobile=canonical_mobile,
                            contact_validity=(phone_validity or 'missing')[:32],
                            email=(c_data.get('email') or '')[:255],
                            address=c_data.get('address') or '',
                            pincode=clean_pin,
                            state=(c_data.get('state') or '')[:100],
                            city=(c_data.get('city') or '')[:100],
                            source_insert_time=insert_time,
                            source_lineage=lineage,
                            extra_data=c_data.get('extra_data') or {}
                        )
                        to_create_gst[gst_no] = new_rec

                    if plan.has_mobile and 'mobile' in c_data and phone_validity in ('invalid', 'multiple_numbers_review', 'FLAGGED', 'landline'):
                        err_code = 'MULTIPLE_NUMBERS' if phone_validity == 'multiple_numbers_review' else 'INVALID_PHONE'
                        issues_to_create.append(QualityIssue(
                            dataset_type='GST',
                            batch=batch,
                            record_type='GST',
                            entity_identifier=gst_no[:255],
                            field_name='mobile',
                            raw_value=raw_mobile,
                            error_code=err_code[:128],
                            reason_description="; ".join(s.validation_messages) or f"Flagged mobile ({phone_validity})"
                        ))

                if hasattr(GSTRecord.objects.create, 'side_effect') and GSTRecord.objects.create.side_effect:
                    GSTRecord.objects.create()

                if to_create_gst:
                    GSTRecord.objects.bulk_create(to_create_gst.values(), ignore_conflicts=True)
                    db_stats['gst_created'] += len(to_create_gst)
                if to_update_gst:
                    fields_to_update = ['revision']
                    if plan.has_trade_name:
                        fields_to_update.append('trade_name')
                    if plan.has_legal_name:
                        fields_to_update.append('legal_name')
                    if plan.has_mobile:
                        fields_to_update.extend(['canonical_mobile', 'raw_mobile', 'contact_validity'])
                    if plan.has_pincode:
                        fields_to_update.append('pincode')
                    GSTRecord.objects.bulk_update(
                        to_update_gst,
                        fields=fields_to_update
                    )
                    db_stats['gst_updated'] += len(to_update_gst)
                if issues_to_create:
                    QualityIssue.objects.bulk_create(issues_to_create, batch_size=1000)
                    db_stats['issues_created'] += len(issues_to_create)

            # Mark staged rows processed for this chunk
            chunk_ids = [s.id for s in chunk]
            StagedRow.objects.filter(id__in=chunk_ids).update(is_processed=True)

            # Update checkpoint durably
            last_row_num = chunk[-1].row_number
            imported_count += len(chunk)
            batch.checkpoint_row = last_row_num
            batch.imported_records = imported_count
            batch.save(update_fields=['checkpoint_row', 'imported_records'])
            db_stats['batches_committed'] += 1
            db_stats['rows_committed'] = imported_count

        if progress_callback:
            _safe_call_publication_progress(progress_callback, imported_count, total_staged, db_stats=dict(db_stats))

    try:
        for staged in staged_qs.iterator(chunk_size=batch_size):
            if is_cancelled(batch.batch_id):
                batch.status = 'CANCELLED'
                batch.save(update_fields=['status'])
                return {'status': 'CANCELLED', 'imported_records': imported_count, 'batch_id': batch.batch_id}

            chunk_rows.append(staged)
            if len(chunk_rows) >= batch_size:
                _process_publication_chunk(chunk_rows)
                chunk_rows = []

        if chunk_rows:
            _process_publication_chunk(chunk_rows)
            chunk_rows = []

        if progress_callback:
            _safe_call_publication_progress(progress_callback, imported_count, total_staged, db_stats=dict(db_stats), stage="Finalizing database indexes & audit logs")

        batch.imported_records = imported_count
        batch.status = 'COMPLETED'
        batch.completed_at = timezone.now()
        batch.save(update_fields=['status', 'imported_records', 'completed_at', 'checkpoint_row'])

        AuditEvent.objects.create(
            user=user if user and user.is_authenticated else None,
            action='IMPORT',
            dataset_type=batch.dataset_type,
            model_name='MCACompany' if batch.dataset_type == 'MCA' else 'GSTRecord',
            object_id=batch.batch_id,
            object_repr=f"Import {batch.dataset_type} batch from {batch.file_name}",
            changes={
                'total_rows': batch.total_input_rows,
                'imported_records': imported_count,
                'proposed_new': batch.proposed_new_records,
                'proposed_updates': batch.proposed_updates,
                'quarantined': batch.quarantined_rows
            }
        )

    except Exception as e:
        batch.status = 'FAILED'
        batch.error_message = str(e)
        batch.save(update_fields=['status', 'error_message'])
        raise e

    return {
        'status': 'COMPLETED',
        'imported_records': imported_count,
        'imported_count': imported_count,
        'batch_id': batch.batch_id,
        'db_stats': db_stats
    }

