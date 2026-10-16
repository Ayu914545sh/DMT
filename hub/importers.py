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
from typing import Dict, List, Tuple, Any, Optional
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
    clean_mobile, clean_date, clean_timestamp, clean_email,
    clean_capital, clean_pincode, clean_pan, clean_gstin,
    clean_cin, clean_din, clean_name, clean_address,
    sanitize_spreadsheet_value
)


# Canonical schema fields and approved aliases
MCA_FIELDS = {
    'entityId': ['entityid', 'entity_id', 'cin', 'llpin', 'registration_number', 'reg_no', 'company_id', 'cin_no', 'cinno', 'corporate_identification_number', 'corporate id'],
    'entityType': ['entitytype', 'entity_type', 'type_of_entity', 'form_type', 'entity category'],
    'name': ['name', 'company_name', 'entity_name', 'corporate_name', 'company', 'firm_name', 'business_name'],
    'state': ['state', 'registered_state', 'state_name'],
    'district': ['district', 'district_name'],
    'city': ['city', 'registered_city'],
    'roc': ['roc', 'roc_code', 'registrar_of_companies', 'roc name'],
    'nicCode': ['niccode', 'nic_code', 'nic', 'industry_code', 'activity_code'],
    'nicLabel': ['niclabel', 'nic_label', 'activity_description', 'nic_description', 'industry_description'],
    'classOfCompany': ['classofcompany', 'class_of_company', 'company_class', 'class'],
    'dateOfIncorporation': ['dateofincorporation', 'date_of_incorporation', 'incorporation_date', 'inc_date', 'doi', 'date of incorporation'],
    'paidUpCapital': ['paidupcapital', 'paid_up_capital', 'paidup_capital', 'paid_capital', 'paid up capital'],
    'authorizedCapital': ['authorizedcapital', 'authorized_capital', 'auth_capital', 'authorized capital'],
    'email': ['email', 'company_email', 'business_email', 'corporate_email', 'email_id', 'email id'],
    'registeredAddress': ['registeredaddress', 'registered_address', 'address', 'reg_address', 'office_address'],
    'directorName': ['directorname', 'director_name', 'partner_name', 'person_name', 'director / partner name', 'name of director'],
    'din': ['din', 'din_no', 'dinno', 'director_din', 'director_identification_number', 'din / pan'],
    'directorEmail': ['directoremail', 'director_email', 'person_email', 'contact_email', 'director email id'],
    'directorMobile': ['directormobile', 'director_mobile', 'director_phone', 'mobile_no', 'mobile', 'phone', 'contact', 'contact_number', 'mobile number', 'phone number'],
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
        'mobile/phone/contact number', 'mobile / phone / contact number', 'mobile no',
        'phone no', 'contact no', 'mobile number', 'phone number', 'contact number'
    ],
    'email': ['email', 'email_address', 'contact_email', 'email id', 'email_id', 'e-mail', 'email id / contact email'],
    'address': ['address', 'principal_address', 'registered_address', 'principal place of business', 'business address'],
    'Pincode': ['pincode', 'pin_code', 'pin', 'postal_code', 'zip', 'pin/pincode', 'pin / pincode', 'pin code', 'postal code'],
    'state': ['state', 'state_name', 'state / ut', 'state name'],
    'city': ['city', 'district', 'town'],
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
    'legalname': 'legal_name',
    'tradename': 'trade_name',
    'registrationdate': 'registration_date',
    'reg_date': 'registration_date',
    'natureofbusiness': 'nature_of_business',
    'gsttype': 'gst_type',
    'pan_number': 'pan',
    'email_address': 'email',
}

MCA_TARGET_ALIASES = {
    'cin': 'entityId',
    'llpin': 'entityId',
    'entity_id': 'entityId',
    'company_name': 'name',
    'director_name': 'directorName',
    'director_mobile': 'directorMobile',
    'director_phone': 'directorMobile',
    'mobile_number': 'directorMobile',
    'mobile': 'directorMobile',
    'director_email': 'directorEmail',
    'date_of_incorporation': 'dateOfIncorporation',
    'incorporation_date': 'dateOfIncorporation',
    'doi': 'dateOfIncorporation',
    'paid_up_capital': 'paidUpCapital',
    'authorized_capital': 'authorizedCapital',
    'registered_address': 'registeredAddress',
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
    
    mca_indicators = {'entityid', 'entitytype', 'roc', 'niccode', 'classofcompany', 'directorname'}
    gst_indicators = {'gstnumber', 'gstin', 'legalname', 'tradename', 'pan', 'statejurisdiction'}
    
    mca_score = sum(1 for ind in mca_indicators if any(ind in h for h in normalized))
    gst_score = sum(1 for ind in gst_indicators if any(ind in h for h in normalized))
    
    return 'GST' if gst_score > mca_score else 'MCA'


def generate_suggested_mapping(headers: List[str], dataset_type: str) -> Dict[str, str]:
    """
    Matches incoming file headers to target canonical fields using exact and approved aliases.
    Does not produce ambiguous mappings.
    Returns mapping dictionary: {source_header: canonical_field_or_None}
    """
    schema = MCA_FIELDS if dataset_type == 'MCA' else GST_FIELDS
    mapping = {}
    assigned_targets = set()
    
    for header in headers:
        clean_header = header.strip()
        base_h = re.sub(r'\.\d+$', '', clean_header).strip()
        h_norm = base_h.lower().replace(' ', '_').replace('-', '_').replace('/', '_')
        h_compact = base_h.lower().replace(' ', '').replace('_', '').replace('-', '').replace('/', '')
        
        # Check sub-tokens if compound header like 'Mobile/Phone/Contact'
        sub_tokens = [tok.strip().lower() for tok in base_h.replace('/', ' ').replace('-', ' ').replace('_', ' ').split() if tok.strip()]

        matched_target = None
        for canonical, aliases in schema.items():
            if canonical in assigned_targets:
                continue
            canonical_compact = canonical.lower().replace('_', '')
            alias_list = [a.lower() for a in aliases]
            alias_compact = [a.lower().replace('_', '').replace('/', '') for a in aliases]
            
            if h_compact == canonical_compact or h_norm in alias_list or h_compact in alias_compact:
                matched_target = canonical
                break
            elif any(tok in alias_list or tok in alias_compact or tok == canonical_compact for tok in sub_tokens):
                matched_target = canonical
                break
                
        if matched_target:
            mapping[clean_header] = matched_target
            assigned_targets.add(matched_target)
        else:
            mapping[clean_header] = ""  # Unmapped; stored in extra_data
            
    return mapping


def validate_mapping_requirements(mapping: Dict[str, str], dataset_type: str) -> Tuple[bool, List[str], List[str]]:
    """
    Checks if mandatory identity fields are mapped and prevents mapping multiple source
    columns to the same target field.
    MCA requires: entityId (or cin/llpin)
    GST requires: gst_number (or gstin)
    Returns: (is_valid, errors_list, warnings_list)
    """
    errors = []
    warnings = []

    aliases = GST_TARGET_ALIASES if dataset_type == 'GST' else MCA_TARGET_ALIASES

    # Check for duplicate target field mappings (after normalizing aliases)
    target_counts = {}
    for src, tgt in mapping.items():
        if tgt:
            norm_tgt = aliases.get(tgt, tgt)
            target_counts[norm_tgt] = target_counts.get(norm_tgt, 0) + 1

    duplicate_targets = [tgt for tgt, count in target_counts.items() if count > 1]
    if duplicate_targets:
        errors.append(f"Multiple source columns cannot map to the same target field: {', '.join(duplicate_targets)}. Please ensure each field is mapped once.")

    mapped_targets = set(target_counts.keys())
    
    if dataset_type == 'MCA':
        if 'entityId' not in mapped_targets and 'cin' not in mapped_targets:
            errors.append("Mandatory field 'entityId' (CIN/LLPIN) is missing from column mapping. Import cannot proceed.")
        # Warn for key unmapped optional fields
        recommended = ['name', 'directorName', 'directorMobile', 'state']
        for rec in recommended:
            if rec not in mapped_targets:
                warnings.append(f"Optional field '{rec}' is currently unmapped. Values for this field will be stored in extra_data.")
    elif dataset_type == 'GST':
        if 'gst_number' not in mapped_targets and 'gstin' not in mapped_targets:
            errors.append("Mandatory field 'gst_number' (GSTIN) is missing from column mapping. Import cannot proceed.")
        # Warn for key unmapped optional fields
        recommended = ['trade_name', 'legal_name', 'mobile', 'pan', 'state']
        for rec in recommended:
            if rec not in mapped_targets:
                warnings.append(f"Optional field '{rec}' is currently unmapped. Values for this field will be stored in extra_data.")

    # Warn for unmapped source columns
    unmapped_sources = [src for src, tgt in mapping.items() if not tgt]
    if unmapped_sources:
        warnings.append(f"{len(unmapped_sources)} column(s) left unmapped ({', '.join(unmapped_sources[:4])}{'...' if len(unmapped_sources) > 4 else ''}). Their original values are safely preserved in extra_data.")

    return len(errors) == 0, errors, warnings


def load_raw_dataset_dataframe(batch: ImportBatch, max_rows: Optional[int] = None) -> pd.DataFrame:
    """
    Loads raw dataset into pandas DataFrame with all cells as text/native types.
    Supports CSV, XLSX, and legacy XLS files safely.
    Handles blank rows and trailing empty columns without losing data.
    """
    file_path = batch.file_path
    if batch.file_format == 'CSV':
        delimiter = batch.mapping_config.get('delimiter', ',')
        encoding = batch.mapping_config.get('encoding', 'utf-8')
        df = pd.read_csv(
            file_path,
            delimiter=delimiter,
            encoding=encoding,
            header=batch.header_row_index,
            nrows=max_rows,
            dtype=str,
            keep_default_na=False
        )
    else:
        sheet = batch.sheet_name if batch.sheet_name else 0
        ext = file_path.lower().split('.')[-1]
        engine = 'xlrd' if ext == 'xls' or batch.file_format == 'XLS' else 'openpyxl'
        try:
            df = pd.read_excel(
                file_path,
                sheet_name=sheet,
                header=batch.header_row_index,
                nrows=max_rows,
                dtype=str,
                keep_default_na=False,
                engine=engine
            )
        except Exception:
            alt_engine = 'openpyxl' if engine == 'xlrd' else 'xlrd'
            df = pd.read_excel(
                file_path,
                sheet_name=sheet,
                header=batch.header_row_index,
                nrows=max_rows,
                dtype=str,
                keep_default_na=False,
                engine=alt_engine
            )
        
    # Strip column names to avoid leading/trailing whitespace mapping mismatches
    df.columns = [str(c).replace('\ufeff', '').strip() for c in df.columns]
    # Drop rows that are completely empty
    df = df.dropna(how='all')
    # Filter out empty string rows
    df = df[~df.apply(lambda row: all(str(val).strip() == '' for val in row), axis=1)]
    return df


def stage_and_validate_batch(batch: ImportBatch, strict_phone: bool = True) -> Dict[str, Any]:
    """
    Reads the input file, applies column mapping and data cleaners,
    detects changes/conflicts against current PostgreSQL state and within-file duplicates,
    and stages candidates into StagedRow records.
    Produces comprehensive preview statistics.
    """
    # Delete prior staging for this batch to ensure clean idempotent run
    StagedRow.objects.filter(batch=batch).delete()
    
    df = load_raw_dataset_dataframe(batch)
    batch.total_input_rows = len(df)
    
    mapping = batch.mapping_config.get('columns', {})
    date_format = batch.mapping_config.get('date_format', None)
    
    # Reverse mapping: canonical_field -> source_header
    target_to_source = {target: src for src, target in mapping.items() if target}
    
    proposed_new = 0
    proposed_updates = 0
    unchanged = 0
    missing_invalid = 0
    conflict_count = 0
    quarantined = 0
    
    # Pre-cache existing IDs for fast in-memory resolution
    if batch.dataset_type == 'MCA':
        existing_companies = {
            (c.entity_type, c.entity_id): c
            for c in MCACompany.objects.all()
        }
    else:
        existing_gst = {
            g.gst_number: g
            for g in GSTRecord.objects.all()
        }

    seen_gstins = {}  # Tracks {gstin: first_row_number} for in-file duplicate detection
    staged_records = []
    
    for row_idx, (_, row) in enumerate(df.iterrows(), start=1):
        raw_row_dict = {col: ("" if pd.isna(val) else str(val)) for col, val in row.items()}
        cleaned_row = {}
        extra_data = {}
        validation_msgs = []
        status = 'VALID'
        action = 'CREATE'
        diff = {}
        
        # Populate mapped fields and extra fields
        for src_col, val in raw_row_dict.items():
            canonical = mapping.get(src_col)
            if canonical:
                if batch.dataset_type == 'GST':
                    canonical = GST_TARGET_ALIASES.get(canonical, canonical)
                elif batch.dataset_type == 'MCA':
                    canonical = MCA_TARGET_ALIASES.get(canonical, canonical)
                cleaned_row[canonical] = val.strip()
            else:
                if val.strip():
                    extra_data[src_col] = val.strip()
        cleaned_row['extra_data'] = extra_data
        
        if batch.dataset_type == 'MCA':
            raw_entity_id = cleaned_row.get('entityId', '')
            entity_id, cin_valid, cin_err = clean_cin(raw_entity_id) if raw_entity_id else (None, False, 'Missing CIN')
            entity_id = (entity_id or raw_entity_id).strip().upper()
            entity_type = cleaned_row.get('entityType', '').strip().lower()
            
            # Infer entity_type if missing: CIN starts with U or L -> company; LLPIN has hyphen or 7-8 chars -> llp
            if not entity_type:
                if '-' in entity_id or len(entity_id) < 15:
                    entity_type = 'llp'
                else:
                    entity_type = 'company'
            cleaned_row['entityType'] = entity_type
            cleaned_row['entityId'] = entity_id
            
            if not entity_id:
                status = 'ERROR'
                action = 'QUARANTINE'
                validation_msgs.append("Missing mandatory Entity ID (CIN/LLPIN)")
                quarantined += 1
                missing_invalid += 1
            else:
                if not cin_valid and cin_err:
                    validation_msgs.append(cin_err)
                    if status != 'ERROR':
                        status = 'WARNING'

                # Clean company name
                cleaned_row['name'], _, _ = clean_name(cleaned_row.get('name', ''))

                # Clean company email
                raw_comp_email = cleaned_row.get('email', '')
                c_c_email, c_email_stat, c_email_err = clean_email(raw_comp_email)
                cleaned_row['clean_email'] = c_c_email
                cleaned_row['email'] = c_c_email or raw_comp_email
                if c_email_stat == 'FLAGGED':
                    validation_msgs.append(f"Company Email: {c_email_err}")

                # Clean registered address
                cleaned_row['registeredAddress'], _, _ = clean_address(cleaned_row.get('registeredAddress', ''))

                # Clean phone
                raw_mobile = cleaned_row.get('directorMobile', '')
                canonical_phone, phone_status, phone_err = clean_mobile(raw_mobile, strict_prefix=strict_phone)
                cleaned_row['original_phone'] = raw_mobile
                cleaned_row['canonical_mobile'] = canonical_phone
                cleaned_row['cleaned_phone'] = canonical_phone
                cleaned_row['phone_validity'] = phone_status
                cleaned_row['phone_status'] = phone_status
                if phone_status in ('invalid', 'multiple_numbers_review', 'FLAGGED'):
                    validation_msgs.append(f"Mobile flagged ({phone_status}): {phone_err}")
                    if status != 'ERROR':
                        status = 'WARNING'
                    missing_invalid += 1
                elif phone_status in ('missing', 'BLANK'):
                    missing_invalid += 1

                # Clean director name and email
                cleaned_row['directorName'], _, _ = clean_name(cleaned_row.get('directorName', ''))
                raw_dir_email = cleaned_row.get('directorEmail', '')
                c_dir_email, d_email_stat, d_email_err = clean_email(raw_dir_email)
                cleaned_row['directorEmail'] = c_dir_email or raw_dir_email
                if d_email_stat == 'FLAGGED':
                    validation_msgs.append(f"Director Email: {d_email_err}")

                # Clean DIN
                raw_din = cleaned_row.get('din', '')
                c_din, din_valid, din_err = clean_din(raw_din)
                cleaned_row['clean_din'] = c_din
                if not din_valid and raw_din:
                    validation_msgs.append(din_err)
                    
                # Clean date of incorporation
                raw_doi = cleaned_row.get('dateOfIncorporation', '')
                clean_doi, doi_valid, doi_err = clean_date(raw_doi, preferred_format=date_format)
                cleaned_row['clean_doi'] = clean_doi.isoformat() if clean_doi else None
                if not doi_valid:
                    validation_msgs.append(doi_err)
                    status = 'WARNING'
                    
                # Clean capital
                paid_up, p_valid, _ = clean_capital(cleaned_row.get('paidUpCapital'))
                auth_cap, a_valid, _ = clean_capital(cleaned_row.get('authorizedCapital'))
                cleaned_row['clean_paid_up_capital'] = str(paid_up) if paid_up is not None else None
                cleaned_row['clean_authorized_capital'] = str(auth_cap) if auth_cap is not None else None
                
                # Check against existing company in DB
                key = (entity_type, entity_id)
                existing = existing_companies.get(key)
                if existing:
                    # Detect if company attributes have changed
                    field_diffs = {}
                    if cleaned_row.get('name') and cleaned_row.get('name') != existing.name:
                        field_diffs['name'] = {'old': existing.name, 'new': cleaned_row.get('name')}
                    if cleaned_row.get('state') and cleaned_row.get('state') != existing.state:
                        field_diffs['state'] = {'old': existing.state, 'new': cleaned_row.get('state')}
                    if cleaned_row.get('roc') and cleaned_row.get('roc') != existing.roc:
                        field_diffs['roc'] = {'old': existing.roc, 'new': cleaned_row.get('roc')}
                    if cleaned_row.get('email') and cleaned_row.get('email') != existing.email:
                        field_diffs['email'] = {'old': existing.email, 'new': cleaned_row.get('email')}
                        
                    if field_diffs:
                        action = 'UPDATE'
                        diff['company_updates'] = field_diffs
                        proposed_updates += 1
                    else:
                        action = 'UNCHANGED'
                        unchanged += 1
                else:
                    action = 'CREATE'
                    proposed_new += 1
                    
        else:  # GST
            raw_gst = cleaned_row.get('gst_number', '').strip().upper()
            cleaned_row['gst_number'] = raw_gst
            
            if not raw_gst:
                status = 'ERROR'
                action = 'QUARANTINE'
                validation_msgs.append("Missing mandatory GST number")
                quarantined += 1
                missing_invalid += 1
            else:
                # GSTIN validation check
                gstin_val, gstin_ok, gstin_err = clean_gstin(raw_gst)
                gst_no = gstin_val or raw_gst
                cleaned_row['gst_number'] = gst_no
                if not gstin_ok:
                    validation_msgs.append(gstin_err)
                    if status != 'ERROR':
                        status = 'WARNING'
                    
                # Clean mobile using strict rules
                raw_mobile = cleaned_row.get('mobile', '')
                canonical_phone, phone_status, phone_err = clean_mobile(raw_mobile, strict_prefix=strict_phone)
                cleaned_row['original_phone'] = raw_mobile
                cleaned_row['canonical_mobile'] = canonical_phone
                cleaned_row['cleaned_phone'] = canonical_phone
                cleaned_row['phone_validity'] = phone_status
                cleaned_row['phone_status'] = phone_status
                if phone_status in ('invalid', 'multiple_numbers_review', 'FLAGGED'):
                    validation_msgs.append(f"Mobile flagged ({phone_status}): {phone_err}")
                    if status != 'ERROR':
                        status = 'WARNING'
                    missing_invalid += 1
                elif phone_status in ('missing', 'BLANK'):
                    missing_invalid += 1

                # Clean email
                raw_email = cleaned_row.get('email', '')
                c_email, email_stat, email_err = clean_email(raw_email)
                cleaned_row['original_email'] = raw_email
                cleaned_row['clean_email'] = c_email
                cleaned_row['email'] = c_email or ""
                if email_stat == 'FLAGGED':
                    validation_msgs.append(f"Email: {email_err}")
                    if status != 'ERROR':
                        status = 'WARNING'
                    
                # Clean PAN (auto-derive from GSTIN if blank)
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

                # Clean names & address
                cleaned_row['legal_name'], _, _ = clean_name(cleaned_row.get('legal_name', ''))
                cleaned_row['trade_name'], _, _ = clean_name(cleaned_row.get('trade_name', ''))
                cleaned_row['address'], _, _ = clean_address(cleaned_row.get('address', ''))

                # Clean registration date
                raw_reg_date = cleaned_row.get('registration_date', '')
                clean_reg, reg_valid, reg_err = clean_date(raw_reg_date, preferred_format=date_format)
                cleaned_row['clean_registration_date'] = clean_reg.isoformat() if clean_reg else None
                if not reg_valid:
                    validation_msgs.append(reg_err)
                    if status != 'ERROR':
                        status = 'WARNING'
                    
                # Clean pincode
                clean_pin, pin_ok, pin_err = clean_pincode(cleaned_row.get('Pincode'))
                cleaned_row['clean_pincode'] = clean_pin
                if not pin_ok and cleaned_row.get('Pincode'):
                    validation_msgs.append(pin_err)
                
                # Check for in-file duplicate GSTIN
                is_infile_dup = False
                if gst_no in seen_gstins:
                    conflict_count += 1
                    is_infile_dup = True
                    validation_msgs.append(f"Duplicate GSTIN in uploaded file (first seen on row {seen_gstins[gst_no]})")
                    diff['duplicate_in_file'] = seen_gstins[gst_no]
                    if status != 'ERROR':
                        status = 'WARNING'
                else:
                    seen_gstins[gst_no] = row_idx

                # Check existing GST record in DB
                existing = existing_gst.get(gst_no)
                if existing:
                    field_diffs = {}
                    # Safe update: blank fields do not erase populated existing values
                    if cleaned_row.get('trade_name') and cleaned_row.get('trade_name') != existing.trade_name:
                        field_diffs['trade_name'] = {'old': existing.trade_name, 'new': cleaned_row.get('trade_name')}
                    if cleaned_row.get('legal_name') and cleaned_row.get('legal_name') != existing.legal_name:
                        field_diffs['legal_name'] = {'old': existing.legal_name, 'new': cleaned_row.get('legal_name')}
                    if canonical_phone and canonical_phone != existing.canonical_mobile:
                        field_diffs['mobile'] = {'old': existing.canonical_mobile, 'new': canonical_phone}
                    if cleaned_row.get('pan') and cleaned_row.get('pan') != existing.pan:
                        field_diffs['pan'] = {'old': existing.pan, 'new': cleaned_row.get('pan')}
                    if cleaned_row.get('email') and cleaned_row.get('email') != existing.email:
                        field_diffs['email'] = {'old': existing.email, 'new': cleaned_row.get('email')}
                    if cleaned_row.get('status') and cleaned_row.get('status') != existing.status:
                        field_diffs['status'] = {'old': existing.status, 'new': cleaned_row.get('status')}
                    if clean_pin and clean_pin != existing.pincode:
                        field_diffs['pincode'] = {'old': existing.pincode, 'new': clean_pin}
                    if cleaned_row.get('address') and cleaned_row.get('address') != existing.address:
                        field_diffs['address'] = {'old': existing.address, 'new': cleaned_row.get('address')}
                        
                    if field_diffs:
                        action = 'UPDATE'
                        diff['gst_updates'] = field_diffs
                        proposed_updates += 1
                    else:
                        action = 'UNCHANGED'
                        unchanged += 1
                elif is_infile_dup:
                    action = 'UPDATE'
                    proposed_updates += 1
                else:
                    action = 'CREATE'
                    proposed_new += 1
                    
        staged_records.append(StagedRow(
            batch=batch,
            row_number=row_idx,
            raw_data=raw_row_dict,
            cleaned_data=cleaned_row,
            proposed_action=action,
            validation_status=status,
            validation_messages=validation_msgs,
            diff_summary=diff
        ))
        
    # Bulk create staged rows for high performance
    StagedRow.objects.bulk_create(staged_records, batch_size=1000)
    
    # Update batch summary
    batch.proposed_new_records = proposed_new
    batch.proposed_updates = proposed_updates
    batch.unchanged_records = unchanged
    batch.missing_invalid_values = missing_invalid
    batch.conflict_count = conflict_count
    batch.quarantined_rows = quarantined
    batch.status = 'PREVIEWED'
    batch.save()
    
    return {
        'total_rows': batch.total_input_rows,
        'proposed_new': proposed_new,
        'proposed_updates': proposed_updates,
        'unchanged': unchanged,
        'missing_invalid': missing_invalid,
        'quarantined': quarantined,
        'sample_staged': [
            {
                'row_number': s.row_number,
                'action': s.proposed_action,
                'status': s.validation_status,
                'messages': s.validation_messages,
                'raw': s.raw_data,
                'cleaned': s.cleaned_data
            }
            for s in staged_records[:10]
        ]
    }


def execute_import_publication(batch: ImportBatch, user: Optional[Any] = None) -> Dict[str, Any]:
    """
    Applies staged rows atomically into PostgreSQL.
    - Idempotent: safe against retries and duplicate runs.
    - Preserves distinct company-director relationships even with repeated entity IDs.
    - Never silently assigns GST contacts to MCA directors.
    - Logs QualityIssue records for invalid phone numbers or formats.
    - Emits AuditEvent.
    """
    batch.status = 'IMPORTING'
    batch.save()
    
    staged_rows = StagedRow.objects.filter(batch=batch).order_by('row_number')
    imported_count = 0
    issues_to_create = []
    
    try:
        with transaction.atomic():
            if batch.dataset_type == 'MCA':
                # Map to track companies created/updated in this batch
                company_map = {}
                
                for staged in staged_rows:
                    if staged.proposed_action == 'QUARANTINE':
                        continue
                        
                    cleaned = staged.cleaned_data
                    entity_type = cleaned.get('entityType', 'company')
                    entity_id = cleaned.get('entityId', '')
                    if not entity_id:
                        continue
                        
                    key = (entity_type, entity_id)
                    company = company_map.get(key)
                    
                    if not company:
                        company = MCACompany.objects.filter(entity_type=entity_type, entity_id=entity_id).first()
                        
                    # Parse dates and capital safely
                    doi_date = None
                    if cleaned.get('clean_doi'):
                        try:
                            doi_date = datetime.date.fromisoformat(cleaned['clean_doi'])
                        except ValueError:
                            pass
                            
                    paid_up = None
                    if cleaned.get('clean_paid_up_capital'):
                        try:
                            paid_up = cleaned['clean_paid_up_capital']
                        except Exception:
                            pass
                            
                    auth_cap = None
                    if cleaned.get('clean_authorized_capital'):
                        try:
                            auth_cap = cleaned['clean_authorized_capital']
                        except Exception:
                            pass
                            
                    upload_dt, _, _ = clean_timestamp(cleaned.get('upload_date'))
                    
                    if not company:
                        # Create new MCA Company
                        company = MCACompany.objects.create(
                            entity_type=entity_type,
                            entity_id=entity_id,
                            name=cleaned.get('name') or '',
                            state=cleaned.get('state') or '',
                            district=cleaned.get('district') or '',
                            city=cleaned.get('city') or '',
                            roc=cleaned.get('roc') or '',
                            nic_code=cleaned.get('nicCode') or '',
                            nic_label=cleaned.get('nicLabel') or '',
                            class_of_company=cleaned.get('classOfCompany') or '',
                            date_of_incorporation=doi_date,
                            paid_up_capital=paid_up,
                            authorized_capital=auth_cap,
                            email=cleaned.get('email') or '',
                            registered_address=cleaned.get('registeredAddress') or '',
                            source_upload_date=upload_dt,
                            extra_data=cleaned.get('extra_data') or {}
                        )
                    else:
                        # Safe update: populate missing fields, keep existing if imported is blank
                        updated = False
                        if cleaned.get('name') and company.name != cleaned.get('name'):
                            company.name = cleaned.get('name')
                            updated = True
                        if cleaned.get('state') and company.state != cleaned.get('state'):
                            company.state = cleaned.get('state')
                            updated = True
                        if cleaned.get('city') and company.city != cleaned.get('city'):
                            company.city = cleaned.get('city')
                            updated = True
                        if cleaned.get('roc') and company.roc != cleaned.get('roc'):
                            company.roc = cleaned.get('roc')
                            updated = True
                        if doi_date and company.date_of_incorporation != doi_date:
                            company.date_of_incorporation = doi_date
                            updated = True
                        if paid_up and company.paid_up_capital != paid_up:
                            company.paid_up_capital = paid_up
                            updated = True
                        if auth_cap and company.authorized_capital != auth_cap:
                            company.authorized_capital = auth_cap
                            updated = True
                        if updated:
                            company.revision += 1
                            company.save()
                            
                    company_map[key] = company
                    
                    # Attach Director / Person entry
                    dir_name = (cleaned.get('directorName') or '').strip()
                    if dir_name:
                        raw_mobile = cleaned.get('directorMobile') or ''
                        canonical_mobile = cleaned.get('canonical_mobile')
                        phone_validity = cleaned.get('phone_validity') or 'missing'
                        dir_email = cleaned.get('directorEmail') or ''
                        
                        designation = 'Designated Partner' if entity_type == 'llp' else 'Director'
                        
                        lineage = {
                            'batch_id': batch.batch_id,
                            'file_name': batch.file_name,
                            'row_number': staged.row_number,
                            'sheet_name': batch.sheet_name
                        }
                        if cleaned.get('din'):
                            lineage['din'] = str(cleaned.get('din')).strip()
                        
                        # Idempotency check: does this director entry already exist on this company?
                        existing_dir = MCADirectorEntry.objects.filter(
                            company=company,
                            director_name=dir_name,
                            canonical_mobile=canonical_mobile
                        ).first()
                        
                        if not existing_dir:
                            dir_entry = MCADirectorEntry.objects.create(
                                company=company,
                                director_name=dir_name,
                                director_email=dir_email,
                                raw_mobile=raw_mobile,
                                canonical_mobile=canonical_mobile,
                                contact_validity=phone_validity,
                                designation=designation,
                                source_lineage=lineage
                            )
                            # Create QualityIssue if invalid or multiple numbers review
                            if phone_validity in ('invalid', 'multiple_numbers_review', 'FLAGGED'):
                                err_code = 'MULTIPLE_NUMBERS' if phone_validity == 'multiple_numbers_review' else 'INVALID_PHONE'
                                issues_to_create.append(QualityIssue(
                                    dataset_type='MCA',
                                    batch=batch,
                                    record_id=dir_entry.id,
                                    record_type='DIRECTOR',
                                    entity_identifier=f"{entity_id} / {dir_name}",
                                    field_name='directorMobile',
                                    raw_value=raw_mobile,
                                    error_code=err_code,
                                    reason_description="; ".join(staged.validation_messages) or f"Flagged mobile ({phone_validity})"
                                ))
                    
                    staged.is_processed = True
                    staged.processed_record_id = str(company.id)
                    staged.save(update_fields=['is_processed', 'processed_record_id'])
                    imported_count += 1
                    
            else:  # GST
                for staged in staged_rows:
                    if staged.proposed_action == 'QUARANTINE':
                        continue
                        
                    cleaned = staged.cleaned_data
                    gst_no = (cleaned.get('gst_number') or '').strip().upper()
                    if not gst_no:
                        continue
                        
                    reg_date = None
                    if cleaned.get('clean_registration_date'):
                        try:
                            reg_date = datetime.date.fromisoformat(cleaned['clean_registration_date'])
                        except ValueError:
                            pass
                            
                    insert_time, _, _ = clean_timestamp(cleaned.get('insert_time'))
                    raw_mobile = cleaned.get('mobile') or ''
                    canonical_mobile = cleaned.get('canonical_mobile')
                    phone_validity = cleaned.get('phone_validity') or 'missing'
                    
                    lineage = {
                        'batch_id': batch.batch_id,
                        'file_name': batch.file_name,
                        'row_number': staged.row_number,
                        'sheet_name': batch.sheet_name
                    }
                    
                    gst_record = GSTRecord.objects.filter(gst_number=gst_no).first()
                    
                    if not gst_record:
                        gst_record = GSTRecord.objects.create(
                            gst_number=gst_no,
                            legal_name=cleaned.get('legal_name') or '',
                            trade_name=cleaned.get('trade_name') or '',
                            nature_of_business=cleaned.get('nature_of_business') or '',
                            source_type=cleaned.get('type') or '',
                            registration_date=reg_date,
                            gst_type=cleaned.get('gst_type') or '',
                            status=cleaned.get('status') or '',
                            jurisdiction=cleaned.get('jurisdiction') or '',
                            state_jurisdiction=cleaned.get('state_jurisdiction') or '',
                            pan=cleaned.get('pan') or '',
                            raw_mobile=raw_mobile,
                            canonical_mobile=canonical_mobile,
                            contact_validity=phone_validity,
                            email=cleaned.get('email') or '',
                            address=cleaned.get('address') or '',
                            pincode=cleaned.get('clean_pincode') or cleaned.get('Pincode') or '',
                            state=cleaned.get('state') or '',
                            city=cleaned.get('city') or '',
                            source_insert_time=insert_time,
                            source_lineage=lineage,
                            extra_data=cleaned.get('extra_data') or {}
                        )
                        if phone_validity in ('invalid', 'multiple_numbers_review', 'FLAGGED'):
                            err_code = 'MULTIPLE_NUMBERS' if phone_validity == 'multiple_numbers_review' else 'INVALID_PHONE'
                            issues_to_create.append(QualityIssue(
                                dataset_type='GST',
                                batch=batch,
                                record_id=gst_record.id,
                                record_type='GST',
                                entity_identifier=gst_no,
                                field_name='mobile',
                                raw_value=raw_mobile,
                                error_code=err_code,
                                reason_description="; ".join(staged.validation_messages) or f"Flagged mobile ({phone_validity})"
                            ))
                    else:
                        # Safe update: populate or update attributes without blanking out existing data
                        updated = False
                        if cleaned.get('trade_name') and gst_record.trade_name != cleaned.get('trade_name'):
                            gst_record.trade_name = cleaned.get('trade_name')
                            updated = True
                        if cleaned.get('legal_name') and gst_record.legal_name != cleaned.get('legal_name'):
                            gst_record.legal_name = cleaned.get('legal_name')
                            updated = True
                        if canonical_mobile and gst_record.canonical_mobile != canonical_mobile:
                            gst_record.canonical_mobile = canonical_mobile
                            gst_record.raw_mobile = raw_mobile
                            gst_record.contact_validity = phone_validity
                            updated = True
                        if cleaned.get('pan') and gst_record.pan != cleaned.get('pan'):
                            gst_record.pan = cleaned.get('pan')
                            updated = True
                        if cleaned.get('email') and gst_record.email != cleaned.get('email'):
                            gst_record.email = cleaned.get('email')
                            updated = True
                        if cleaned.get('status') and gst_record.status != cleaned.get('status'):
                            gst_record.status = cleaned.get('status')
                            updated = True
                        if cleaned.get('address') and gst_record.address != cleaned.get('address'):
                            gst_record.address = cleaned.get('address')
                            updated = True
                        if cleaned.get('city') and gst_record.city != cleaned.get('city'):
                            gst_record.city = cleaned.get('city')
                            updated = True
                        if cleaned.get('state') and gst_record.state != cleaned.get('state'):
                            gst_record.state = cleaned.get('state')
                            updated = True
                        clean_pin = cleaned.get('clean_pincode') or cleaned.get('Pincode')
                        if clean_pin and gst_record.pincode != clean_pin:
                            gst_record.pincode = clean_pin
                            updated = True
                        if reg_date and gst_record.registration_date != reg_date:
                            gst_record.registration_date = reg_date
                            updated = True
                        if cleaned.get('nature_of_business') and gst_record.nature_of_business != cleaned.get('nature_of_business'):
                            gst_record.nature_of_business = cleaned.get('nature_of_business')
                            updated = True
                        if updated:
                            gst_record.revision += 1
                            gst_record.save()
                            
                    staged.is_processed = True
                    staged.processed_record_id = str(gst_record.id)
                    staged.save(update_fields=['is_processed', 'processed_record_id'])
                    imported_count += 1
                    
            if issues_to_create:
                QualityIssue.objects.bulk_create(issues_to_create, batch_size=500)
                
            batch.imported_records = imported_count
            batch.status = 'COMPLETED'
            batch.completed_at = timezone.now()
            batch.save()
            
            # Record Audit event
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
        batch.save()
        raise e
        
    return {
        'status': 'COMPLETED',
        'imported_records': imported_count,
        'imported_count': imported_count,
        'batch_id': batch.batch_id
    }
