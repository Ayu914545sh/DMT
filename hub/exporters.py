"""
Audited and formula-injection-safe data exporter for MCA and GST datasets.
Preserves identifiers and phone numbers as text, prevents Excel macro/formula execution,
and generates both CSV and XLSX output formats.
"""

import os
import csv
from typing import List, Dict, Any, Optional
import openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from django.conf import settings
from django.utils import timezone

from openpyxl.styles import Font
from .models import MCACompany, MCADirectorEntry, GSTRecord, ExportJob, AuditEvent
from .cleaners import sanitize_spreadsheet_value


def sanitize_text(val: Any) -> str:
    """Removes invalid XML/Excel characters and applies formula protection."""
    if val is None:
        return ""
    text = str(val)
    # Remove illegal xml characters for openpyxl
    text = ILLEGAL_CHARACTERS_RE.sub('', text)
    return sanitize_spreadsheet_value(text)


def export_mca_dataset(queryset, columns: Optional[List[str]] = None, file_format: str = 'CSV', export_dir: Optional[str] = None) -> str:
    """
    Exports MCA Companies along with related director entries.
    Each company produces exactly one row in the output to match table record counts.
    """
    export_dir = export_dir or os.path.join(settings.MEDIA_ROOT, 'exports')
    os.makedirs(export_dir, exist_ok=True)
    
    timestamp_str = timezone.now().strftime('%Y%m%d_%H%M%S')
    filename = f"mca_export_{timestamp_str}.{file_format.lower()}"
    file_path = os.path.join(export_dir, filename)
    
    # Default columns if none chosen
    if not columns:
        columns = [
            'entity_id', 'entity_type', 'name', 'state', 'city', 'roc', 'date_of_incorporation',
            'paid_up_capital', 'email', 'director_name', 'canonical_mobile', 'contact_validity'
        ]
        
    rows_data = []
    # Prefetch directors for maximum efficiency
    qs = queryset.prefetch_related('directors')
    
    for company in qs:
        directors = list(company.directors.all())
        row = {}
        for col in columns:
            if col == 'director_name':
                row[col] = "; ".join(d.director_name for d in directors if d.director_name)
            elif col in ('director_email',):
                row[col] = "; ".join(d.director_email for d in directors if d.director_email)
            elif col in ('raw_mobile', 'original_phone'):
                row[col] = "; ".join(d.raw_mobile for d in directors if d.raw_mobile)
            elif col in ('canonical_mobile', 'cleaned_phone'):
                row[col] = "; ".join(d.canonical_mobile for d in directors if d.canonical_mobile)
            elif col in ('contact_validity', 'phone_status'):
                row[col] = "; ".join(d.contact_validity for d in directors if d.contact_validity)
            elif col == 'designation':
                row[col] = "; ".join(d.designation for d in directors if d.designation)
            elif col == 'din':
                row[col] = "; ".join(d.din for d in directors if d.din)
            else:
                val = getattr(company, col, "")
                row[col] = str(val) if val is not None else ""
        rows_data.append(row)
            
    if file_format.upper() == 'CSV':
        with open(file_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=columns)
            writer.writeheader()
            for r in rows_data:
                safe_r = {k: sanitize_text(v) for k, v in r.items()}
                writer.writerow(safe_r)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "MCA Export"
        ws.append(columns)
        
        # Style header row
        header_font = Font(bold=True)
        for col_idx in range(1, len(columns) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font

        # Write data rows with explicit text number formatting to preserve leading zeros & identifiers
        for r_idx, r in enumerate(rows_data, start=2):
            for c_idx, c in enumerate(columns, start=1):
                val = r.get(c, "")
                text_val = sanitize_text(val)
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.value = str(text_val) if text_val is not None else ""
                cell.number_format = '@'

        wb.save(file_path)
        wb.close()
        
    return file_path


def export_gst_dataset(queryset, columns: Optional[List[str]] = None, file_format: str = 'CSV', export_dir: Optional[str] = None) -> str:
    """
    Exports GST Registrations with formula sanitization and text formatting.
    """
    export_dir = export_dir or os.path.join(settings.MEDIA_ROOT, 'exports')
    os.makedirs(export_dir, exist_ok=True)
    
    timestamp_str = timezone.now().strftime('%Y%m%d_%H%M%S')
    filename = f"gst_export_{timestamp_str}.{file_format.lower()}"
    file_path = os.path.join(export_dir, filename)
    
    if not columns:
        columns = [
            'gst_number', 'legal_name', 'trade_name', 'status', 'source_type',
            'registration_date', 'pan', 'canonical_mobile', 'contact_validity',
            'email', 'state', 'city', 'pincode'
        ]
        
    if file_format.upper() == 'CSV':
        with open(file_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(columns)
            for rec in queryset:
                row_vals = []
                for col in columns:
                    val = getattr(rec, col, "")
                    row_vals.append(sanitize_text(val))
                writer.writerow(row_vals)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "GST Export"
        ws.append(columns)
        
        # Style header row
        header_font = Font(bold=True)
        for col_idx in range(1, len(columns) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font

        # Write data rows with explicit text number formatting to preserve leading zeros & identifiers
        for r_idx, rec in enumerate(queryset, start=2):
            for c_idx, col in enumerate(columns, start=1):
                val = getattr(rec, col, "")
                text_val = sanitize_text(val)
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.value = str(text_val) if text_val is not None else ""
                cell.number_format = '@'

        wb.save(file_path)
        wb.close()
        
    return file_path


def run_export_job(job: ExportJob) -> str:
    """
    Executes an ExportJob instance, creates the output file,
    records audit details, and updates the job record.
    """
    job.status = 'PROCESSING'
    job.save(update_fields=['status'])
    
    try:
        if job.dataset_type == 'MCA':
            qs = MCACompany.objects.filter(is_archived=False)
            # Apply filters if provided
            filters = job.filters or {}
            if 'state' in filters and filters['state']:
                qs = qs.filter(state__iexact=filters['state'])
            if 'entity_type' in filters and filters['entity_type']:
                qs = qs.filter(entity_type=filters['entity_type'])
            if 'roc' in filters and filters['roc']:
                qs = qs.filter(roc__iexact=filters['roc'])
            if 'search' in filters and filters['search']:
                s = filters['search']
                qs = qs.filter(name__icontains=s) | qs.filter(entity_id__icontains=s)
                
            path = export_mca_dataset(qs, job.columns, job.file_format)
            job.row_count = qs.count()
        else:
            qs = GSTRecord.objects.filter(is_archived=False)
            filters = job.filters or {}
            if 'state' in filters and filters['state']:
                qs = qs.filter(state__iexact=filters['state'])
            if 'status' in filters and filters['status']:
                qs = qs.filter(status__iexact=filters['status'])
            if 'search' in filters and filters['search']:
                s = filters['search']
                qs = qs.filter(legal_name__icontains=s) | qs.filter(trade_name__icontains=s) | qs.filter(gst_number__icontains=s)
                
            path = export_gst_dataset(qs, job.columns, job.file_format)
            job.row_count = qs.count()
            
        job.file_path = path
        job.status = 'COMPLETED'
        job.completed_at = timezone.now()
        job.save()
        
        # Audit log
        AuditEvent.objects.create(
            user=job.requester,
            action='EXPORT',
            dataset_type=job.dataset_type,
            model_name='MCACompany' if job.dataset_type == 'MCA' else 'GSTRecord',
            object_id=str(job.id),
            object_repr=f"Exported {job.dataset_type} ({job.row_count} records) as {job.file_format}",
            changes={'columns': job.columns, 'filters': job.filters, 'file_path': path}
        )
        
        return path
    except Exception as e:
        job.status = 'FAILED'
        job.save(update_fields=['status'])
        raise e
