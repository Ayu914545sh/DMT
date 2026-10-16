"""
Views for the Unified MCA & GST Data Management Hub.
Full-featured dashboard, multi-step import pipeline, record management,
global search, data quality inbox, exports, and audit logging.
"""

import os
import csv
import io
import json
import logging
import datetime
from django.shortcuts import render, redirect, get_object_or_404
from django.template.loader import render_to_string
from django.urls import reverse
from django.http import HttpResponse, JsonResponse, Http404
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.core.paginator import Paginator
from django.conf import settings
from django.utils import timezone
from django.db import transaction
from django.db.models import Count, Q

logger = logging.getLogger(__name__)

from .models import (
    MCACompany, MCADirectorEntry, GSTRecord,
    ImportBatch, StagedRow, QualityIssue,
    MappingTemplate, AuditEvent, ExportJob,
    BusinessLink, SystemSetting
)
from .cleaners import clean_mobile, sanitize_spreadsheet_value
from .importers import (
    calculate_file_checksum, detect_csv_properties, get_excel_sheets,
    read_file_headers, read_file_samples, detect_dataset_type,
    generate_suggested_mapping, validate_mapping_requirements,
    stage_and_validate_batch, execute_import_publication,
    auto_detect_header_row, read_excel_rows,
    MCA_FIELDS, GST_FIELDS,
    GST_TARGET_ALIASES, MCA_TARGET_ALIASES
)
from .services import perform_unified_search, resolve_quality_issue
from .forms import (
    DatasetUploadForm, MCACompanyForm, MCADirectorForm,
    GSTRecordForm, ExportRequestForm, AdminPasswordChangeForm
)
from .exporters import run_export_job, export_mca_dataset, export_gst_dataset


# ==========================================
# Authentication Views
# ==========================================

def login_view(request):
    if request.user.is_authenticated:
        return redirect('hub:dashboard')
        
    error = None
    if request.method == 'POST':
        uname = request.POST.get('username')
        pword = request.POST.get('password')
        user = authenticate(request, username=uname, password=pword)
        if user is not None:
            login(request, user)
            AuditEvent.objects.create(
                user=user,
                action='CREATE',
                dataset_type='SYSTEM',
                model_name='UserSession',
                object_id=str(user.id),
                object_repr=f"User {user.username} logged in",
                ip_address=request.META.get('REMOTE_ADDR')
            )
            return redirect(request.GET.get('next', 'hub:dashboard'))
        else:
            error = "Invalid username or password"
            
    return render(request, 'auth/login.html', {'error': error})


def logout_view(request):
    if request.user.is_authenticated:
        AuditEvent.objects.create(
            user=request.user,
            action='DELETE',
            dataset_type='SYSTEM',
            model_name='UserSession',
            object_id=str(request.user.id),
            object_repr=f"User {request.user.username} logged out",
            ip_address=request.META.get('REMOTE_ADDR')
        )
    logout(request)
    return redirect('hub:login')


# ==========================================
# Dashboard View
# ==========================================

@login_required
def dashboard_view(request):
    total_companies = MCACompany.objects.filter(entity_type='company', is_archived=False).count()
    total_llps = MCACompany.objects.filter(entity_type='llp', is_archived=False).count()
    total_directors = MCADirectorEntry.objects.count()
    total_gst = GSTRecord.objects.filter(is_archived=False).count()
    
    # Phone health statistics (supports both new lowercase and legacy uppercase statuses)
    mca_valid_phones = MCADirectorEntry.objects.filter(contact_validity__in=['valid', 'cleaned', 'VALID']).count()
    mca_flagged_phones = MCADirectorEntry.objects.filter(contact_validity__in=['invalid', 'multiple_numbers_review', 'FLAGGED']).count()
    mca_blank_phones = MCADirectorEntry.objects.filter(contact_validity__in=['missing', 'BLANK']).count()
    
    gst_valid_phones = GSTRecord.objects.filter(contact_validity__in=['valid', 'cleaned', 'VALID']).count()
    gst_flagged_phones = GSTRecord.objects.filter(contact_validity__in=['invalid', 'multiple_numbers_review', 'FLAGGED']).count()
    gst_blank_phones = GSTRecord.objects.filter(contact_validity__in=['missing', 'BLANK']).count()
    
    recent_batches = ImportBatch.objects.all().order_by('-created_at')[:5]
    open_quality_issues = QualityIssue.objects.filter(is_resolved=False).count()
    
    context = {
        'total_companies': total_companies,
        'total_llps': total_llps,
        'total_mca_entities': total_companies + total_llps,
        'total_directors': total_directors,
        'total_gst': total_gst,
        'mca_valid_phones': mca_valid_phones,
        'mca_flagged_phones': mca_flagged_phones,
        'mca_blank_phones': mca_blank_phones,
        'gst_valid_phones': gst_valid_phones,
        'gst_flagged_phones': gst_flagged_phones,
        'gst_blank_phones': gst_blank_phones,
        'recent_batches': recent_batches,
        'open_quality_issues': open_quality_issues,
    }
    return render(request, 'dashboard.html', context)


# ==========================================
# Unified Upload & Import Pipeline
# ==========================================

@login_required
def import_upload_view(request):
    """
    Step 1: Upload dataset (CSV or XLSX) with sheet/encoding options.
    """
    if request.method == 'POST':
        form = DatasetUploadForm(request.POST, request.FILES)
        if form.is_valid():
            dataset_type = form.cleaned_data['dataset_type']
            uploaded_file = form.cleaned_data['data_file']
            sheet_name = form.cleaned_data.get('sheet_name') or ''
            raw_hr = form.cleaned_data.get('header_row')
            header_row = (raw_hr - 1) if raw_hr and raw_hr >= 1 else 0
            csv_delim = form.cleaned_data.get('csv_delimiter') or ','
            csv_enc = form.cleaned_data.get('csv_encoding') or 'utf-8'

            # Empty file validation (0 bytes)
            if uploaded_file.size == 0:
                messages.error(request, "The selected file is empty (0 bytes). Please upload a valid dataset file.")
                return render(request, 'import/upload.html', {'form': form})

            # Maximum file size validation (50 MB)
            if uploaded_file.size > 50 * 1024 * 1024:
                messages.error(request, "File size exceeds 50 MB limit. Please upload a smaller file.")
                return render(request, 'import/upload.html', {'form': form})
            
            # Save uploaded file to private media directory
            upload_dir = os.path.join(settings.MEDIA_ROOT, 'imports')
            os.makedirs(upload_dir, exist_ok=True)
            
            timestamp = timezone.now().strftime('%Y%m%d_%H%M%S')
            saved_filename = f"{timestamp}_{uploaded_file.name}"
            saved_path = os.path.join(upload_dir, saved_filename)
            
            try:
                with open(saved_path, 'wb+') as destination:
                    for chunk in uploaded_file.chunks():
                        destination.write(chunk)
            except Exception as e:
                messages.error(request, f"Failed to save uploaded file: {str(e)}")
                return render(request, 'import/upload.html', {'form': form})
                    
            checksum = calculate_file_checksum(saved_path)
            ext = uploaded_file.name.split('.')[-1].lower() if '.' in uploaded_file.name else ''
            file_format = 'CSV' if ext == 'csv' else ('XLS' if ext == 'xls' else 'XLSX')
            
            # Check sheet names and corruption for Excel (.xlsx or .xls)
            if file_format in ('XLSX', 'XLS'):
                try:
                    sheets = get_excel_sheets(saved_path)
                    if not sheets:
                        raise ValueError("No worksheets found in Excel workbook.")
                    if not sheet_name:
                        sheet_name = sheets[0]
                    test_rows = read_excel_rows(saved_path, sheet_name=sheet_name, max_rows=5)
                    if not test_rows or all(not any(cell is not None and str(cell).strip() for cell in r) for r in test_rows):
                        if os.path.exists(saved_path):
                            os.remove(saved_path)
                        messages.error(request, "The uploaded Excel file contains no readable data or rows.")
                        return render(request, 'import/upload.html', {'form': form})
                except Exception as e:
                    if os.path.exists(saved_path):
                        os.remove(saved_path)
                    messages.error(request, f"Corrupted or invalid Excel file. The file could not be parsed: {str(e)}")
                    return render(request, 'import/upload.html', {'form': form})
                    
            # Auto-detect delimiter and encoding for CSV if default
            if file_format == 'CSV':
                det_delim, det_enc = detect_csv_properties(saved_path)
                csv_delim = det_delim or csv_delim
                csv_enc = det_enc or csv_enc
                try:
                    with open(saved_path, 'r', encoding=csv_enc, errors='replace') as f:
                        reader = csv.reader(f, delimiter=csv_delim)
                        test_rows = []
                        for _ in range(5):
                            try:
                                row = next(reader)
                                test_rows.append(row)
                            except StopIteration:
                                break
                    if not test_rows or all(not any(str(cell).strip() for cell in r) for r in test_rows):
                        if os.path.exists(saved_path):
                            os.remove(saved_path)
                        messages.error(request, "The uploaded CSV file contains no readable data or rows.")
                        return render(request, 'import/upload.html', {'form': form})
                except Exception as e:
                    if os.path.exists(saved_path):
                        os.remove(saved_path)
                    messages.error(request, f"Corrupted or invalid CSV file: {str(e)}")
                    return render(request, 'import/upload.html', {'form': form})

            # If user left default header row (1) or not specified, auto-detect header row
            if form.cleaned_data.get('header_row') in (1, None):
                auto_hr = auto_detect_header_row(
                    file_path=saved_path,
                    file_format=file_format,
                    sheet_name=sheet_name,
                    delimiter=csv_delim,
                    encoding=csv_enc
                )
                header_row = auto_hr
                
            headers = read_file_headers(
                file_path=saved_path,
                file_format=file_format,
                sheet_name=sheet_name,
                header_row=header_row,
                delimiter=csv_delim,
                encoding=csv_enc
            )
            if not headers:
                if os.path.exists(saved_path):
                    os.remove(saved_path)
                messages.error(request, "Could not detect headers in file. Please ensure row 1 (or the detected header row) contains column names.")
                return render(request, 'import/upload.html', {'form': form})

            # Enforce schema isolation
            detected_type = detect_dataset_type(headers)
            if dataset_type == 'GST' and detected_type == 'MCA':
                if os.path.exists(saved_path):
                    os.remove(saved_path)
                messages.error(
                    request,
                    "MCA file detected: The uploaded file contains MCA corporate columns (CIN, ROC, Director) "
                    "and cannot be imported into the GST module. Please upload a valid GST registration file or use the MCA Import module."
                )
                return render(request, 'import/upload.html', {'form': form})
            elif dataset_type == 'MCA' and detected_type == 'GST':
                if os.path.exists(saved_path):
                    os.remove(saved_path)
                messages.error(
                    request,
                    "GST file detected: The uploaded file contains GST registration columns (GSTIN, Trade Name) "
                    "and cannot be imported into the MCA module. Please upload a valid MCA file or use the GST Import module."
                )
                return render(request, 'import/upload.html', {'form': form})

            batch_id = f"BATCH-{timestamp}-{checksum[:8].upper()}"
            
            batch = ImportBatch.objects.create(
                batch_id=batch_id,
                dataset_type=dataset_type,
                file_name=uploaded_file.name,
                file_path=saved_path,
                file_checksum=checksum,
                file_format=file_format,
                sheet_name=sheet_name,
                header_row_index=header_row,
                status='UPLOADED',
                created_by=request.user,
                mapping_config={
                    'delimiter': csv_delim,
                    'encoding': csv_enc,
                    'sheet_name': sheet_name
                }
            )
            
            AuditEvent.objects.create(
                user=request.user,
                action='CREATE',
                dataset_type=dataset_type,
                model_name='ImportBatch',
                object_id=batch.batch_id,
                object_repr=f"Uploaded file {uploaded_file.name} for {dataset_type}",
                changes={'file_size': uploaded_file.size, 'checksum': checksum}
            )
            
            return redirect('hub:import_map', batch_id=batch.batch_id)
        else:
            for field, err_list in form.errors.items():
                for err in err_list:
                    messages.error(request, f"{field.replace('_', ' ').capitalize()}: {err}")
    else:
        form = DatasetUploadForm()
        
    return render(request, 'import/upload.html', {'form': form})


@login_required
def import_map_view(request, batch_id):
    """
    Step 2: Column mapping review and confirmation.
    """
    batch = get_object_or_404(ImportBatch, batch_id=batch_id)
    
    headers = read_file_headers(
        file_path=batch.file_path,
        file_format=batch.file_format,
        sheet_name=batch.sheet_name,
        header_row=batch.header_row_index,
        delimiter=batch.mapping_config.get('delimiter', ','),
        encoding=batch.mapping_config.get('encoding', 'utf-8')
    )
    
    if not headers:
        messages.error(request, "Could not detect headers in file. Please check header row selection.")
        return redirect('hub:import_upload')
        
    # Auto-detected schema suggestion
    detected_type = detect_dataset_type(headers)
    canonical_schema = MCA_FIELDS if batch.dataset_type == 'MCA' else GST_FIELDS
    templates = MappingTemplate.objects.filter(dataset_type=batch.dataset_type).order_by('name')
    available_sheets = get_excel_sheets(batch.file_path) if batch.file_format in ('XLSX', 'XLS') else []
    
    # Extract sample row values for each column
    column_samples = read_file_samples(
        file_path=batch.file_path,
        file_format=batch.file_format,
        sheet_name=batch.sheet_name,
        header_row=batch.header_row_index,
        delimiter=batch.mapping_config.get('delimiter', ','),
        encoding=batch.mapping_config.get('encoding', 'utf-8'),
        sample_rows_count=3
    )

    # Check if a template was requested
    load_template_id = request.GET.get('use_template')
    if load_template_id:
        tmpl = MappingTemplate.objects.filter(id=load_template_id, dataset_type=batch.dataset_type).first()
        if tmpl and isinstance(tmpl.mapping_config, dict):
            suggested_mapping = tmpl.mapping_config
        else:
            suggested_mapping = generate_suggested_mapping(headers, batch.dataset_type)
    elif batch.mapping_config.get('columns'):
        suggested_mapping = batch.mapping_config.get('columns')
    else:
        suggested_mapping = generate_suggested_mapping(headers, batch.dataset_type)
    
    if request.method == 'POST':
        # 1. Switch Excel worksheet if requested
        if 'switch_sheet' in request.POST and batch.status != 'COMPLETED':
            new_sheet = request.POST.get('switch_sheet', '').strip()
            if new_sheet and new_sheet in available_sheets:
                batch.sheet_name = new_sheet
                batch.header_row_index = auto_detect_header_row(
                    file_path=batch.file_path,
                    file_format=batch.file_format,
                    sheet_name=new_sheet
                )
                batch.mapping_config['sheet_name'] = new_sheet
                batch.mapping_config.pop('columns', None)
                batch.status = 'UPLOADED'
                batch.save()
                StagedRow.objects.filter(batch=batch).delete()
                messages.success(request, f"Switched to worksheet '{new_sheet}'. Column mappings and samples reloaded.")
                return redirect('hub:import_map', batch_id=batch.batch_id)

        # 2. Correct dataset type if requested
        if 'correct_dataset' in request.POST and batch.status != 'COMPLETED':
            batch.dataset_type = detected_type
            batch.mapping_config = {
                'delimiter': batch.mapping_config.get('delimiter', ','),
                'encoding': batch.mapping_config.get('encoding', 'utf-8'),
                'sheet_name': batch.mapping_config.get('sheet_name', '')
            }
            batch.status = 'UPLOADED'
            batch.save()
            StagedRow.objects.filter(batch=batch).delete()
            messages.success(request, f"Dataset target corrected to {detected_type}. Please map columns.")
            return redirect('hub:import_map', batch_id=batch.batch_id)

        # Schema isolation check: reject cross-module imports
        if batch.dataset_type == 'GST' and detected_type == 'MCA':
            messages.error(request, "MCA dataset detected: This file contains MCA company and director columns (CIN, ROC, Director) and cannot be imported into the GST module. Please upload a valid GST file or use the MCA Import module.")
            return render(request, 'import/map.html', {
                'batch': batch,
                'headers': headers,
                'suggested_mapping': suggested_mapping,
                'canonical_schema': canonical_schema,
                'detected_type': detected_type,
                'date_format': '%Y-%m-%d',
                'column_samples': column_samples,
                'templates': templates,
                'available_sheets': available_sheets,
            })
            
        if batch.dataset_type == 'MCA' and detected_type == 'GST':
            messages.error(request, "GST dataset detected: This file contains GST registration columns (GSTIN, Trade Name) and cannot be imported into the MCA module. Please upload a valid MCA file or use the GST Import module.")
            return render(request, 'import/map.html', {
                'batch': batch,
                'headers': headers,
                'suggested_mapping': suggested_mapping,
                'canonical_schema': canonical_schema,
                'detected_type': detected_type,
                'date_format': '%m/%d/%Y',
                'column_samples': column_samples,
                'templates': templates,
                'available_sheets': available_sheets,
            })

        # Collect submitted mapping with alias normalization
        submitted_mapping = {}
        aliases = GST_TARGET_ALIASES if batch.dataset_type == 'GST' else MCA_TARGET_ALIASES
        for h in headers:
            mapped_val = request.POST.get(f"map_{h}", "").strip()
            if mapped_val and mapped_val in aliases:
                mapped_val = aliases[mapped_val]
            submitted_mapping[h] = mapped_val
            
        date_format = request.POST.get('date_format', '%m/%d/%Y' if batch.dataset_type == 'MCA' else '%Y-%m-%d')
        template_name = request.POST.get('template_name', '').strip()
        
        # Validate identity requirement, duplicate targets, and collect warnings
        is_valid, errors, warnings = validate_mapping_requirements(submitted_mapping, batch.dataset_type)
        if not is_valid:
            for err in errors:
                messages.error(request, err)
            return render(request, 'import/map.html', {
                'batch': batch,
                'headers': headers,
                'suggested_mapping': submitted_mapping,
                'canonical_schema': canonical_schema,
                'detected_type': detected_type,
                'date_format': date_format,
                'column_samples': column_samples,
                'templates': templates,
                'available_sheets': available_sheets,
            })
            
        if warnings:
            for w in warnings:
                messages.warning(request, w)
            
        # Save mapping config into batch
        batch.mapping_config['columns'] = submitted_mapping
        batch.mapping_config['date_format'] = date_format
        batch.status = 'MAPPED'
        batch.save()
        
        # Save reusable template if requested
        if template_name:
            MappingTemplate.objects.update_or_create(
                name=template_name,
                dataset_type=batch.dataset_type,
                defaults={
                    'mapping_config': submitted_mapping,
                    'date_format': date_format,
                    'delimiter': batch.mapping_config.get('delimiter', ','),
                    'created_by': request.user
                }
            )
            
        # Run staging & validation
        strict_phone = SystemSetting.get_bool('STRICT_PHONE_PREFIX', default=True)
        stage_and_validate_batch(batch, strict_phone=strict_phone)
        
        return redirect('hub:import_preview', batch_id=batch.batch_id)
        
    return render(request, 'import/map.html', {
        'batch': batch,
        'headers': headers,
        'suggested_mapping': suggested_mapping,
        'canonical_schema': canonical_schema,
        'detected_type': detected_type,
        'date_format': '%m/%d/%Y' if batch.dataset_type == 'MCA' else '%Y-%m-%d',
        'column_samples': column_samples,
        'templates': templates,
        'available_sheets': available_sheets,
    })


@login_required
def import_preview_view(request, batch_id):
    """
    Step 3: Staging review, conflict inspection, and publication approval.
    """
    batch = get_object_or_404(ImportBatch, batch_id=batch_id)
    
    staged_qs = StagedRow.objects.filter(batch=batch).order_by('row_number')
    
    # Filter preview table
    status_filter = request.GET.get('status')
    if status_filter:
        staged_qs = staged_qs.filter(validation_status=status_filter)
        
    paginator = Paginator(staged_qs, 25)
    page_number = request.GET.get('page', 1)
    page_obj = paginator.get_page(page_number)
    
    context = {
        'batch': batch,
        'page_obj': page_obj,
        'status_filter': status_filter,
    }
    return render(request, 'import/preview.html', context)


@login_required
def import_reclean_view(request, batch_id):
    """
    Re-runs data cleaning and validation for a staged batch without re-uploading.
    """
    batch = get_object_or_404(ImportBatch, batch_id=batch_id)
    if batch.status == 'COMPLETED':
        messages.info(request, "This batch has already been published to the database.")
        return redirect('hub:import_batch_detail', batch_id=batch.batch_id)
    strict_phone = SystemSetting.get_bool('STRICT_PHONE_PREFIX', default=True)
    stage_and_validate_batch(batch, strict_phone=strict_phone)
    messages.success(request, "Data cleaning re-executed successfully! Preview metrics updated.")
    return redirect('hub:import_preview', batch_id=batch.batch_id)


@login_required
def import_reset_view(request, batch_id):
    """
    Resets staged records for a batch back to the column mapping stage.
    """
    batch = get_object_or_404(ImportBatch, batch_id=batch_id)
    if batch.status == 'COMPLETED':
        messages.warning(request, "Completed batches cannot be reset.")
        return redirect('hub:import_batch_detail', batch_id=batch.batch_id)
    StagedRow.objects.filter(batch=batch).delete()
    batch.status = 'UPLOADED'
    batch.proposed_new_records = 0
    batch.proposed_updates = 0
    batch.unchanged_records = 0
    batch.missing_invalid_values = 0
    batch.conflict_count = 0
    batch.quarantined_rows = 0
    batch.save()
    messages.info(request, "Batch staging reset. You can now adjust column mappings or re-run cleaning.")
    return redirect('hub:import_map', batch_id=batch.batch_id)


@login_required
def download_cleaned_file_view(request, batch_id):
    """
    Downloads a sanitized CSV of all cleaned records ready for external use.
    Does not modify the original source file.
    """
    batch = get_object_or_404(ImportBatch, batch_id=batch_id)
    staged_rows = StagedRow.objects.filter(batch=batch).order_by('row_number')
    
    response = HttpResponse(content_type='text/csv; charset=utf-8')
    filename = f"cleaned_{batch.dataset_type}_{batch.batch_id}.csv"
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    
    writer = csv.writer(response)
    if batch.dataset_type == 'GST':
        headers = [
            'Row Number', 'GSTIN', 'Legal Name', 'Trade Name', 'Cleaned Mobile',
            'Original Mobile', 'Phone Status', 'Email', 'PAN', 'Status',
            'Registration Date', 'Address', 'City', 'State', 'Pincode',
            'Validation Status', 'Validation Notes'
        ]
        writer.writerow(headers)
        for s in staged_rows:
            c = s.cleaned_data
            writer.writerow([
                s.row_number,
                sanitize_spreadsheet_value(c.get('gst_number', '')),
                sanitize_spreadsheet_value(c.get('legal_name', '')),
                sanitize_spreadsheet_value(c.get('trade_name', '')),
                sanitize_spreadsheet_value(c.get('canonical_mobile', '') or c.get('cleaned_phone', '')),
                sanitize_spreadsheet_value(c.get('original_phone', '')),
                c.get('phone_status', '') or c.get('phone_validity', ''),
                sanitize_spreadsheet_value(c.get('email', '') or c.get('clean_email', '')),
                sanitize_spreadsheet_value(c.get('pan', '') or c.get('clean_pan', '')),
                sanitize_spreadsheet_value(c.get('status', '')),
                sanitize_spreadsheet_value(c.get('clean_registration_date', '') or c.get('registration_date', '')),
                sanitize_spreadsheet_value(c.get('address', '')),
                sanitize_spreadsheet_value(c.get('city', '')),
                sanitize_spreadsheet_value(c.get('state', '')),
                sanitize_spreadsheet_value(c.get('clean_pincode', '') or c.get('Pincode', '')),
                s.validation_status,
                sanitize_spreadsheet_value("; ".join(s.validation_messages))
            ])
    else:
        headers = [
            'Row Number', 'Entity ID', 'Entity Type', 'Company Name', 'State',
            'City', 'ROC', 'Email', 'Incorporation Date', 'Paid Up Capital',
            'Authorized Capital', 'Director Name', 'Cleaned Mobile', 'Original Mobile',
            'Phone Status', 'Director Email', 'DIN', 'Validation Status', 'Validation Notes'
        ]
        writer.writerow(headers)
        for s in staged_rows:
            c = s.cleaned_data
            writer.writerow([
                s.row_number,
                sanitize_spreadsheet_value(c.get('entityId', '')),
                sanitize_spreadsheet_value(c.get('entityType', '')),
                sanitize_spreadsheet_value(c.get('name', '')),
                sanitize_spreadsheet_value(c.get('state', '')),
                sanitize_spreadsheet_value(c.get('city', '')),
                sanitize_spreadsheet_value(c.get('roc', '')),
                sanitize_spreadsheet_value(c.get('email', '') or c.get('clean_email', '')),
                sanitize_spreadsheet_value(c.get('clean_doi', '') or c.get('dateOfIncorporation', '')),
                sanitize_spreadsheet_value(c.get('clean_paid_up_capital', '') or c.get('paidUpCapital', '')),
                sanitize_spreadsheet_value(c.get('clean_authorized_capital', '') or c.get('authorizedCapital', '')),
                sanitize_spreadsheet_value(c.get('directorName', '')),
                sanitize_spreadsheet_value(c.get('canonical_mobile', '') or c.get('cleaned_phone', '')),
                sanitize_spreadsheet_value(c.get('original_phone', '')),
                c.get('phone_status', '') or c.get('phone_validity', ''),
                sanitize_spreadsheet_value(c.get('directorEmail', '')),
                sanitize_spreadsheet_value(c.get('clean_din', '') or c.get('din', '')),
                s.validation_status,
                sanitize_spreadsheet_value("; ".join(s.validation_messages))
            ])
    return response


@login_required
def import_execute_view(request, batch_id):
    """
    Step 4: Execute publication atomically.
    """
    batch = get_object_or_404(ImportBatch, batch_id=batch_id)
    if request.method == 'POST':
        try:
            execute_import_publication(batch, user=request.user)
            if batch.dataset_type == 'GST':
                messages.success(
                    request,
                    f"GST import completed successfully! {batch.imported_records} records processed: "
                    f"{batch.proposed_new_records} new imported, {batch.proposed_updates} updated, "
                    f"{batch.unchanged_records} unchanged (skipped), {batch.conflict_count} duplicate, "
                    f"{batch.quarantined_rows} rejected."
                )
                return redirect('hub:gst_list')
            else:
                messages.success(
                    request,
                    f"MCA import completed successfully! {batch.imported_records} records processed: "
                    f"{batch.proposed_new_records} new imported, {batch.proposed_updates} updated, "
                    f"{batch.unchanged_records} unchanged (skipped), {batch.conflict_count} duplicate, "
                    f"{batch.quarantined_rows} rejected."
                )
                return redirect('hub:mca_list')
        except Exception as e:
            messages.error(request, f"Import failed: {str(e)}")
            return redirect('hub:import_preview', batch_id=batch.batch_id)
            
        return redirect('hub:import_batch_detail', batch_id=batch.batch_id)
    return redirect('hub:import_preview', batch_id=batch.batch_id)


@login_required
def import_batch_detail_view(request, batch_id):
    """
    Displays completed batch results, statistics, and staged row audit.
    """
    batch = get_object_or_404(ImportBatch, batch_id=batch_id)
    staged_rows = StagedRow.objects.filter(batch=batch).order_by('row_number')
    
    action_filter = request.GET.get('action')
    if action_filter:
        staged_rows = staged_rows.filter(proposed_action=action_filter)
        
    paginator = Paginator(staged_rows, 50)
    page_obj = paginator.get_page(request.GET.get('page', 1))
    
    context = {
        'batch': batch,
        'page_obj': page_obj,
        'action_filter': action_filter,
        'quality_issues': QualityIssue.objects.filter(batch=batch)
    }
    return render(request, 'import/batch_detail.html', context)


@login_required
def import_history_view(request):
    batches = ImportBatch.objects.all().order_by('-created_at')
    dataset = request.GET.get('dataset')
    if dataset:
        batches = batches.filter(dataset_type=dataset)
        
    paginator = Paginator(batches, 15)
    page_obj = paginator.get_page(request.GET.get('page', 1))
    return render(request, 'import/history.html', {'page_obj': page_obj, 'dataset': dataset})


@login_required
def download_cleaning_report_view(request, batch_id):
    """
    Generates downloadable CSV report of every input row and its exact outcome.
    """
    batch = get_object_or_404(ImportBatch, batch_id=batch_id)
    staged_rows = StagedRow.objects.filter(batch=batch).order_by('row_number')
    
    response = HttpResponse(content_type='text/csv; charset=utf-8')
    response['Content-Disposition'] = f'attachment; filename="cleaning_report_{batch.batch_id}.csv"'
    
    writer = csv.writer(response)
    writer.writerow([
        'Row Number', 'Proposed Action', 'Validation Status', 'Validation Messages',
        'Raw Mobile', 'Cleaned Mobile', 'Phone Validity', 'Email', 'PAN', 'Entity Identifier', 'Processed Record ID', 'Unmapped Extra Data'
    ])
    
    for s in staged_rows:
        cleaned = s.cleaned_data
        entity_id = cleaned.get('entityId') or cleaned.get('gst_number') or ""
        extra_json = json.dumps(cleaned.get('extra_data', {})) if cleaned.get('extra_data') else ""
        writer.writerow([
            s.row_number,
            s.proposed_action,
            s.validation_status,
            sanitize_spreadsheet_value("; ".join(s.validation_messages)),
            sanitize_spreadsheet_value(cleaned.get('original_phone') or cleaned.get('directorMobile') or cleaned.get('mobile') or ""),
            sanitize_spreadsheet_value(cleaned.get('canonical_mobile') or cleaned.get('cleaned_phone') or ""),
            cleaned.get('phone_validity') or cleaned.get('phone_status') or "",
            sanitize_spreadsheet_value(cleaned.get('email') or cleaned.get('directorEmail') or ""),
            sanitize_spreadsheet_value(cleaned.get('pan') or ""),
            sanitize_spreadsheet_value(entity_id),
            s.processed_record_id or "",
            sanitize_spreadsheet_value(extra_json)
        ])
        
    return response


# ==========================================
# MCA Module Views
# ==========================================

def build_mca_filtered_queryset(request):
    """
    Builds a secure, parameterized, whitelisted queryset for MCA records.
    Supports individual column filters (f_*) and general search with strict AND logic.
    """
    qs = MCACompany.objects.filter(is_archived=False).prefetch_related('directors')
    col_filters = {
        'f_entity_id': '',
        'f_entity_type': '',
        'f_name': '',
        'f_state': '',
        'f_city': '',
        'f_roc': '',
        'f_email': '',
        'f_address': '',
        'f_director': '',
        'f_phone': '',
    }

    # 1. Whitelisted Column Filters
    # Entity ID (CIN / LLPIN) - treated as text preserving leading zeros
    f_entity_id = (request.GET.get('f_entity_id') or request.GET.get('f_cin') or '').strip()
    if f_entity_id:
        col_filters['f_entity_id'] = f_entity_id
        qs = qs.filter(entity_id__icontains=f_entity_id)

    # Entity Type (company / llp)
    f_entity_type = (request.GET.get('f_entity_type') or request.GET.get('f_type') or '').strip()
    if f_entity_type:
        col_filters['f_entity_type'] = f_entity_type
        qs = qs.filter(entity_type__icontains=f_entity_type)

    # Company Name (case-insensitive partial matching)
    f_name = (request.GET.get('f_name') or request.GET.get('f_company') or '').strip()
    if f_name:
        col_filters['f_name'] = f_name
        qs = qs.filter(name__icontains=f_name)

    # State
    f_state = (request.GET.get('f_state') or '').strip()
    if f_state:
        col_filters['f_state'] = f_state
        qs = qs.filter(state__icontains=f_state)

    # City
    f_city = (request.GET.get('f_city') or '').strip()
    if f_city:
        col_filters['f_city'] = f_city
        qs = qs.filter(city__icontains=f_city)

    # ROC
    f_roc = (request.GET.get('f_roc') or '').strip()
    if f_roc:
        col_filters['f_roc'] = f_roc
        qs = qs.filter(roc__icontains=f_roc)

    # Email
    f_email = (request.GET.get('f_email') or '').strip()
    if f_email:
        col_filters['f_email'] = f_email
        qs = qs.filter(email__icontains=f_email)

    # Address
    f_address = (request.GET.get('f_address') or '').strip()
    if f_address:
        col_filters['f_address'] = f_address
        qs = qs.filter(registered_address__icontains=f_address)

    # Director / Phone / Contact
    f_director = (request.GET.get('f_director') or request.GET.get('f_director_name') or '').strip()
    f_phone = (request.GET.get('f_phone') or request.GET.get('f_director_phone') or '').strip()
    if f_director:
        col_filters['f_director'] = f_director
        qs = qs.filter(
            Q(directors__director_name__icontains=f_director) |
            Q(directors__canonical_mobile__icontains=f_director) |
            Q(directors__raw_mobile__icontains=f_director) |
            Q(directors__director_email__icontains=f_director)
        ).distinct()

    if f_phone:
        col_filters['f_phone'] = f_phone
        qs = qs.filter(
            Q(directors__canonical_mobile__icontains=f_phone) |
            Q(directors__raw_mobile__icontains=f_phone)
        ).distinct()

    # 2. Backward compatibility with general search (q) and top filters
    search = request.GET.get('q', '').strip()
    if search:
        qs = qs.filter(
            Q(name__icontains=search) |
            Q(entity_id__icontains=search) |
            Q(email__icontains=search) |
            Q(city__icontains=search) |
            Q(state__icontains=search) |
            Q(roc__icontains=search) |
            Q(directors__director_name__icontains=search) |
            Q(directors__canonical_mobile__icontains=search) |
            Q(directors__raw_mobile__icontains=search)
        ).distinct()

    entity_type = request.GET.get('entity_type', '').strip()
    if entity_type and not f_entity_type:
        qs = qs.filter(entity_type__icontains=entity_type)

    state = request.GET.get('state', '').strip()
    if state and not f_state:
        qs = qs.filter(state__icontains=state)

    roc = request.GET.get('roc', '').strip()
    if roc and not f_roc:
        qs = qs.filter(roc__icontains=roc)

    phone_filter = request.GET.get('phone_status', '').strip()
    if phone_filter:
        pf_upper = phone_filter.strip().upper()
        if 'VALID' in pf_upper:
            qs = qs.filter(directors__contact_validity__in=['valid', 'cleaned', 'VALID']).distinct()
        elif 'FLAG' in pf_upper or 'INV' in pf_upper:
            qs = qs.filter(directors__contact_validity__in=['invalid', 'multiple_numbers_review', 'FLAGGED']).distinct()
        elif 'BLANK' in pf_upper or 'MISS' in pf_upper:
            qs = qs.filter(directors__contact_validity__in=['missing', 'BLANK']).distinct()
        else:
            qs = qs.filter(
                Q(directors__contact_validity__icontains=phone_filter) |
                Q(directors__canonical_mobile__icontains=phone_filter)
            ).distinct()

    # 3. Whitelisted Sorting
    sort_param = request.GET.get('sort', '').strip()
    allowed_sorts = {
        'name': 'name', '-name': '-name', 'name_asc': 'name', 'name_desc': '-name',
        'entity_id': 'entity_id', '-entity_id': '-entity_id', 'entity_id_asc': 'entity_id', 'entity_id_desc': '-entity_id',
        'entity_type': 'entity_type', '-entity_type': '-entity_type', 'entity_type_asc': 'entity_type', 'entity_type_desc': '-entity_type',
        'state': 'state', '-state': '-state', 'state_asc': 'state', 'state_desc': '-state',
        'city': 'city', '-city': '-city', 'city_asc': 'city', 'city_desc': '-city',
        'roc': 'roc', '-roc': '-roc', 'roc_asc': 'roc', 'roc_desc': '-roc',
        'date': 'date_of_incorporation', '-date': '-date_of_incorporation', 'date_asc': 'date_of_incorporation', 'date_desc': '-date_of_incorporation',
        'capital': 'paid_up_capital', '-capital': '-paid_up_capital', 'capital_asc': 'paid_up_capital', 'capital_desc': '-paid_up_capital',
        'updated_at': '-updated_at',
    }
    if sort_param in allowed_sorts:
        qs = qs.order_by(allowed_sorts[sort_param])
    else:
        qs = qs.order_by('-updated_at', 'name')

    return qs, col_filters, search, entity_type, state, roc, phone_filter, sort_param


@login_required
def mca_list_view(request):
    qs, col_filters, search, entity_type, state, roc, phone_filter, sort_param = build_mca_filtered_queryset(request)
    total_count = qs.count()

    paginator = Paginator(qs, 20)
    page_obj = paginator.get_page(request.GET.get('page', 1))

    # State and ROC lists for filter dropdowns
    states = MCACompany.objects.values_list('state', flat=True).distinct().order_by('state')[:50]
    rocs = MCACompany.objects.values_list('roc', flat=True).distinct().order_by('roc')[:50]

    context = {
        'page_obj': page_obj,
        'search': search,
        'entity_type': entity_type,
        'state': state,
        'roc': roc,
        'phone_filter': phone_filter,
        'states': [s for s in states if s],
        'rocs': [r for r in rocs if r],
        'total_count': total_count,
        'col_filters': col_filters,
        'sort': sort_param,
        'has_active_filters': bool(any(col_filters.values()) or search or entity_type or state or roc or phone_filter),
    }

    if request.headers.get('x-requested-with') == 'XMLHttpRequest' or request.GET.get('format') == 'partial':
        html = render_to_string('mca/_table_content.html', context, request=request)
        return JsonResponse({
            'success': True,
            'html': html,
            'total_count': total_count,
            'start_index': page_obj.start_index() if total_count > 0 else 0,
            'end_index': page_obj.end_index() if total_count > 0 else 0,
            'has_other_pages': page_obj.has_other_pages(),
            'page': page_obj.number,
            'num_pages': paginator.num_pages,
        })

    return render(request, 'mca/list.html', context)


@login_required
def mca_detail_view(request, company_id):
    company = get_object_or_404(
        MCACompany.objects.prefetch_related('directors', 'gst_links__gst_record'),
        id=company_id
    )
    directors = company.directors.all()
    audit_events = AuditEvent.objects.filter(
        model_name='MCACompany', object_id=str(company.id)
    )[:10]
    
    context = {
        'company': company,
        'directors': directors,
        'audit_events': audit_events,
        'business_links': company.gst_links.all()
    }
    return render(request, 'mca/detail.html', context)


@login_required
def mca_edit_view(request, company_id):
    company = get_object_or_404(MCACompany, id=company_id)
    if request.method == 'POST':
        form = MCACompanyForm(request.POST, instance=company)
        if form.is_valid():
            old_name = company.name
            form.save()
            company.revision += 1
            company.save()
            
            AuditEvent.objects.create(
                user=request.user,
                action='UPDATE',
                dataset_type='MCA',
                model_name='MCACompany',
                object_id=str(company.id),
                object_repr=f"Updated company {company.name} ({company.entity_id})",
                changes={'old_name': old_name, 'new_name': company.name}
            )
            messages.success(request, "Company details updated successfully.")
            return redirect('hub:mca_detail', company_id=company.id)
    else:
        form = MCACompanyForm(instance=company)
    return render(request, 'mca/edit.html', {'form': form, 'company': company})


@login_required
def mca_delete_view(request, company_id):
    """
    Safely delete an MCA entity and all related director records within an atomic transaction.
    Strictly protected on the backend: requires admin (is_superuser) privileges.
    Returns JSON for AJAX/fetch requests and redirects with flash messages for standard form submissions.
    """
    is_ajax = request.headers.get('x-requested-with') == 'XMLHttpRequest' or 'application/json' in request.headers.get('Accept', '')

    if not request.user.is_superuser:
        msg = "Access denied. Only administrators have permission to delete records."
        if is_ajax:
            return JsonResponse({'success': False, 'message': msg}, status=403)
        messages.error(request, msg)
        return redirect('hub:mca_list')

    if request.method != 'POST':
        msg = "Invalid request method. Delete must be performed via POST."
        if is_ajax:
            return JsonResponse({'success': False, 'message': msg}, status=405)
        messages.error(request, msg)
        return redirect('hub:mca_list')

    try:
        company = MCACompany.objects.get(id=company_id)
    except MCACompany.DoesNotExist:
        msg = "This company record does not exist or has already been deleted."
        if is_ajax:
            return JsonResponse({'success': False, 'already_deleted': True, 'message': msg}, status=404)
        messages.warning(request, msg)
        return redirect('hub:mca_list')

    company_name = company.name
    entity_id = company.entity_id

    try:
        with transaction.atomic():
            # Delete related directors first to prevent any orphaned records
            director_count, _ = MCADirectorEntry.objects.filter(company=company).delete()

            # Record audit event before deleting parent
            AuditEvent.objects.create(
                user=request.user,
                action='DELETE',
                dataset_type='MCA',
                model_name='MCACompany',
                object_id=str(company.id),
                object_repr=f"Deleted MCA Company {company_name} ({entity_id}) with {director_count} directors",
                changes={'company_name': company_name, 'entity_id': entity_id, 'directors_deleted': director_count}
            )

            # Delete the MCA company record
            company.delete()

        # Compute updated counts
        total_mca = MCACompany.objects.filter(is_archived=False).count()
        total_directors = MCADirectorEntry.objects.count()

        success_msg = f"MCA entity '{company_name}' ({entity_id}) and {director_count} related director(s) have been deleted successfully."

        if is_ajax:
            return JsonResponse({
                'success': True,
                'message': success_msg,
                'record_id': company_id,
                'stats': {
                    'total_mca': total_mca,
                    'total_directors': total_directors
                }
            })

        messages.success(request, success_msg)
        return redirect('hub:mca_list')

    except Exception as e:
        err_msg = f"Failed to delete record due to a server error: {str(e)}"
        if is_ajax:
            return JsonResponse({'success': False, 'message': err_msg}, status=500)
        messages.error(request, err_msg)
        return redirect('hub:mca_list')


@login_required
def mca_director_add_view(request, company_id):
    company = get_object_or_404(MCACompany, id=company_id)
    if request.method == 'POST':
        form = MCADirectorForm(request.POST)
        if form.is_valid():
            director = form.save(commit=False)
            director.company = company
            director.save()
            
            AuditEvent.objects.create(
                user=request.user,
                action='CREATE',
                dataset_type='MCA',
                model_name='MCADirectorEntry',
                object_id=str(director.id),
                object_repr=f"Added director {director.director_name} to {company.name}"
            )
            messages.success(request, f"Added person {director.director_name} successfully.")
            return redirect('hub:mca_detail', company_id=company.id)
    else:
        form = MCADirectorForm()
    return render(request, 'mca/director_form.html', {'form': form, 'company': company})


# ==========================================
# Dedicated MCA Navigation & Action Views
# ==========================================

@login_required
def mca_dashboard_view(request):
    """
    Dedicated MCA Module Dashboard with authoritative KPIs, director coverage,
    and phone validation metrics.
    """
    total_companies = MCACompany.objects.filter(entity_type='company', is_archived=False).count()
    total_llps = MCACompany.objects.filter(entity_type='llp', is_archived=False).count()
    total_entities = total_companies + total_llps
    total_directors = MCADirectorEntry.objects.count()

    mca_valid_phones = MCADirectorEntry.objects.filter(contact_validity__in=['valid', 'cleaned', 'VALID']).count()
    mca_flagged_phones = MCADirectorEntry.objects.filter(contact_validity__in=['invalid', 'multiple_numbers_review', 'FLAGGED']).count()
    mca_blank_phones = MCADirectorEntry.objects.filter(contact_validity__in=['missing', 'BLANK']).count()

    top_states = MCACompany.objects.filter(is_archived=False).exclude(state='').values('state').annotate(total=Count('id')).order_by('-total')[:6]
    top_rocs = MCACompany.objects.filter(is_archived=False).exclude(roc='').values('roc').annotate(total=Count('id')).order_by('-total')[:6]
    recent_companies = MCACompany.objects.filter(is_archived=False).prefetch_related('directors').order_by('-created_at')[:8]
    open_mca_issues = QualityIssue.objects.filter(dataset_type='MCA', is_resolved=False).count()

    context = {
        'mca_active_section': 'dashboard',
        'total_companies': total_companies,
        'total_llps': total_llps,
        'total_entities': total_entities,
        'total_directors': total_directors,
        'mca_valid_phones': mca_valid_phones,
        'mca_flagged_phones': mca_flagged_phones,
        'mca_blank_phones': mca_blank_phones,
        'top_states': top_states,
        'top_rocs': top_rocs,
        'recent_companies': recent_companies,
        'open_mca_issues': open_mca_issues,
    }
    return render(request, 'mca/dashboard.html', context)


@login_required
def mca_directors_view(request):
    """
    Dedicated MCA Directors / People directory with search and linked company details.
    """
    qs = MCADirectorEntry.objects.select_related('company').all()
    q = request.GET.get('q', '').strip()
    status_filter = request.GET.get('phone_status', '').strip()

    if q:
        qs = qs.filter(
            Q(director_name__icontains=q) |
            Q(director_email__icontains=q) |
            Q(canonical_mobile__icontains=q) |
            Q(raw_mobile__icontains=q) |
            Q(company__name__icontains=q) |
            Q(company__entity_id__icontains=q)
        )

    if status_filter:
        s_upper = status_filter.upper()
        if 'VALID' in s_upper:
            qs = qs.filter(contact_validity__in=['valid', 'cleaned', 'VALID'])
        elif 'FLAG' in s_upper or 'INV' in s_upper:
            qs = qs.filter(contact_validity__in=['invalid', 'multiple_numbers_review', 'FLAGGED'])
        elif 'BLANK' in s_upper or 'MISS' in s_upper:
            qs = qs.filter(contact_validity__in=['missing', 'BLANK'])
        else:
            qs = qs.filter(contact_validity__iexact=status_filter)

    paginator = Paginator(qs, 25)
    page_obj = paginator.get_page(request.GET.get('page', 1))

    context = {
        'mca_active_section': 'directors',
        'page_obj': page_obj,
        'q': q,
        'phone_status': status_filter,
        'total_count': qs.count(),
    }
    return render(request, 'mca/directors.html', context)


@login_required
def mca_upload_view(request):
    """
    Dedicated MCA-only Upload / Import view. Validates file format, restricts to MCA,
    and forwards to the column mapping step.
    """
    if request.method == 'POST':
        uploaded_file = request.FILES.get('data_file')
        sheet_name = request.POST.get('sheet_name', '').strip()
        
        if not uploaded_file:
            messages.error(request, "Please select an MCA dataset file to upload.")
            return redirect('hub:mca_upload')
            
        ext = uploaded_file.name.split('.')[-1].lower()
        if ext not in ('csv', 'xlsx', 'xls'):
            messages.error(request, "Invalid file format. Only .xlsx, .xls, and .csv files are supported for MCA data.")
            return redirect('hub:mca_upload')

        upload_dir = os.path.join(settings.MEDIA_ROOT, 'imports')
        os.makedirs(upload_dir, exist_ok=True)
        timestamp = timezone.now().strftime('%Y%m%d_%H%M%S')
        saved_filename = f"mca_{timestamp}_{uploaded_file.name}"
        saved_path = os.path.join(upload_dir, saved_filename)

        with open(saved_path, 'wb+') as destination:
            for chunk in uploaded_file.chunks():
                destination.write(chunk)

        checksum = calculate_file_checksum(saved_path)
        file_format = 'CSV' if ext == 'csv' else ('XLS' if ext == 'xls' else 'XLSX')

        if file_format in ('XLSX', 'XLS') and not sheet_name:
            sheets = get_excel_sheets(saved_path)
            if sheets:
                sheet_name = sheets[0]

        csv_delim = ','
        csv_enc = 'utf-8'
        if file_format == 'CSV':
            det_delim, det_enc = detect_csv_properties(saved_path)
            csv_delim = det_delim or ','
            csv_enc = det_enc or 'utf-8'

        auto_hr = auto_detect_header_row(
            file_path=saved_path,
            file_format=file_format,
            sheet_name=sheet_name,
            delimiter=csv_delim,
            encoding=csv_enc
        )

        batch_id = f"BATCH-MCA-{timestamp}-{checksum[:8].upper()}"
        batch = ImportBatch.objects.create(
            batch_id=batch_id,
            dataset_type='MCA',
            file_name=uploaded_file.name,
            file_path=saved_path,
            file_checksum=checksum,
            file_format=file_format,
            sheet_name=sheet_name,
            header_row_index=auto_hr,
            status='UPLOADED',
            created_by=request.user,
            mapping_config={
                'delimiter': csv_delim,
                'encoding': csv_enc,
                'sheet_name': sheet_name
            }
        )

        AuditEvent.objects.create(
            user=request.user,
            action='CREATE',
            dataset_type='MCA',
            model_name='ImportBatch',
            object_id=batch.batch_id,
            object_repr=f"Uploaded MCA file {uploaded_file.name}",
            changes={'file_size': uploaded_file.size, 'checksum': checksum}
        )

        messages.success(request, f"MCA File '{uploaded_file.name}' uploaded successfully. Please confirm column mapping.")
        return redirect('hub:import_map', batch_id=batch.batch_id)

    # Render dedicated MCA upload page
    recent_mca_batches = ImportBatch.objects.filter(dataset_type='MCA').order_by('-created_at')[:6]
    context = {
        'mca_active_section': 'upload',
        'recent_batches': recent_mca_batches,
    }
    return render(request, 'mca/upload.html', context)


@login_required
def mca_cleaning_view(request):
    """
    MCA Data Cleaning & Quality Inbox view. Displays quality issues specifically for MCA entities and directors.
    """
    issues = QualityIssue.objects.filter(dataset_type='MCA').order_by('-created_at')
    status = request.GET.get('status', 'open')
    error_code = request.GET.get('error_code', '')

    if status == 'open':
        issues = issues.filter(is_resolved=False)
    elif status == 'resolved':
        issues = issues.filter(is_resolved=True)

    if error_code:
        issues = issues.filter(error_code=error_code)

    paginator = Paginator(issues, 20)
    page_obj = paginator.get_page(request.GET.get('page', 1))

    error_codes = QualityIssue.objects.filter(dataset_type='MCA').values_list('error_code', flat=True).distinct()

    context = {
        'mca_active_section': 'cleaning',
        'page_obj': page_obj,
        'status': status,
        'error_code': error_code,
        'error_codes': [e for e in error_codes if e],
        'total_count': issues.count(),
    }
    return render(request, 'mca/cleaning.html', context)


@login_required
def mca_search_view(request):
    """
    Dedicated MCA Advanced Search working across company name, CIN, registered office,
    director name, DIN, phone number, email, and location (city/state).
    """
    q = request.GET.get('q', '').strip()
    entity_type = request.GET.get('entity_type', '').strip()
    state = request.GET.get('state', '').strip()

    qs = MCACompany.objects.filter(is_archived=False).prefetch_related('directors')

    if q:
        qs = qs.filter(
            Q(name__icontains=q) |
            Q(entity_id__icontains=q) |
            Q(registered_address__icontains=q) |
            Q(city__icontains=q) |
            Q(state__icontains=q) |
            Q(email__icontains=q) |
            Q(directors__director_name__icontains=q) |
            Q(directors__canonical_mobile__icontains=q) |
            Q(directors__raw_mobile__icontains=q) |
            Q(directors__director_email__icontains=q)
        ).distinct()

    if entity_type:
        qs = qs.filter(entity_type__iexact=entity_type)
    if state:
        qs = qs.filter(state__icontains=state)

    paginator = Paginator(qs, 20)
    page_obj = paginator.get_page(request.GET.get('page', 1))

    context = {
        'mca_active_section': 'search',
        'page_obj': page_obj,
        'q': q,
        'entity_type': entity_type,
        'state': state,
        'total_count': qs.count() if q or entity_type or state else MCACompany.objects.filter(is_archived=False).count(),
    }
    return render(request, 'mca/search.html', context)


@login_required
def mca_reports_view(request):
    """
    MCA Reports & Audited Data Export Center.
    Allows filtering by state, entity type, incorporation year, and phone validity before exporting.
    """
    if request.method == 'POST':
        entity_type = request.POST.get('entity_type', '').strip()
        state = request.POST.get('state', '').strip()
        phone_status = request.POST.get('phone_status', '').strip()
        file_format = request.POST.get('file_format', 'XLSX').upper()

        if file_format not in ('XLSX', 'CSV'):
            file_format = 'XLSX'

        qs = MCACompany.objects.filter(is_archived=False)
        if entity_type:
            qs = qs.filter(entity_type=entity_type)
        if state:
            qs = qs.filter(state__icontains=state)
        if phone_status:
            s_upper = phone_status.upper()
            if 'VALID' in s_upper:
                qs = qs.filter(directors__contact_validity__in=['valid', 'cleaned', 'VALID']).distinct()
            elif 'FLAG' in s_upper or 'INV' in s_upper:
                qs = qs.filter(directors__contact_validity__in=['invalid', 'multiple_numbers_review', 'FLAGGED']).distinct()

        job = ExportJob.objects.create(
            requester=request.user,
            dataset_type='MCA',
            file_format=file_format,
            columns=[],
            filters={'entity_type': entity_type, 'state': state, 'phone_status': phone_status},
            status='PENDING'
        )

        try:
            run_export_job(job)
            messages.success(request, f"MCA Report generated successfully ({job.row_count} records).")
            return redirect('hub:export_download', job_id=job.id)
        except Exception as e:
            messages.error(request, f"Failed to generate report: {str(e)}")
            return redirect('hub:mca_reports')

    recent_mca_exports = ExportJob.objects.filter(dataset_type='MCA', requester=request.user).order_by('-created_at')[:8]
    states = MCACompany.objects.values_list('state', flat=True).distinct().order_by('state')[:50]

    context = {
        'mca_active_section': 'reports',
        'recent_exports': recent_mca_exports,
        'states': [s for s in states if s],
    }
    return render(request, 'mca/reports.html', context)


@login_required
def mca_settings_view(request):
    """
    Dedicated MCA Module settings: Profile, password change, and configuration.
    """
    password_form = AdminPasswordChangeForm(user=request.user) if request.user.is_superuser else None
    context = {
        'mca_active_section': 'settings',
        'password_form': password_form,
    }
    return render(request, 'mca/settings.html', context)


@login_required
def mca_create_company_view(request):
    """
    Directly creates an MCA Company / LLP with validation, optional initial director, and auditing.
    """
    if request.method != 'POST':
        return redirect('hub:mca_list')

    name = request.POST.get('name', '').strip()
    entity_type = request.POST.get('entity_type', 'company').strip().lower()
    entity_id = request.POST.get('entity_id', '').strip().upper()
    email = request.POST.get('email', '').strip()
    state = request.POST.get('state', '').strip()
    city = request.POST.get('city', '').strip()
    roc = request.POST.get('roc', '').strip()
    doi_raw = request.POST.get('date_of_incorporation', '').strip()
    paid_up_raw = request.POST.get('paid_up_capital', '').strip()
    auth_cap_raw = request.POST.get('authorized_capital', '').strip()
    nic_code = request.POST.get('nic_code', '').strip()
    class_of_company = request.POST.get('class_of_company', '').strip()
    registered_address = request.POST.get('registered_address', '').strip()

    # Director details
    director_name = request.POST.get('director_name', '').strip()
    director_mobile = request.POST.get('director_mobile', '').strip()
    director_email = request.POST.get('director_email', '').strip()

    if not name or not entity_id:
        messages.error(request, "Company Name and Entity Identifier (CIN/LLPIN) are required.")
        return redirect('hub:mca_list')

    # Duplicate check
    if MCACompany.objects.filter(entity_type=entity_type, entity_id=entity_id).exists():
        messages.error(request, f"An MCA record with Identifier '{entity_id}' and type '{entity_type.upper()}' already exists.")
        return redirect('hub:mca_list')

    doi = None
    if doi_raw:
        try:
            doi = datetime.datetime.strptime(doi_raw, '%Y-%m-%d').date()
        except ValueError:
            pass

    paid_up = None
    if paid_up_raw:
        try:
            paid_up = float(paid_up_raw)
        except ValueError:
            pass

    auth_cap = None
    if auth_cap_raw:
        try:
            auth_cap = float(auth_cap_raw)
        except ValueError:
            pass

    with transaction.atomic():
        company = MCACompany.objects.create(
            name=name,
            entity_type=entity_type,
            entity_id=entity_id,
            email=email,
            state=state,
            city=city,
            roc=roc,
            date_of_incorporation=doi,
            paid_up_capital=paid_up,
            authorized_capital=auth_cap,
            nic_code=nic_code,
            class_of_company=class_of_company,
            registered_address=registered_address
        )

        if director_name:
            canon_phone, p_status, _ = clean_mobile(director_mobile, strict_prefix=True)
            MCADirectorEntry.objects.create(
                company=company,
                director_name=director_name,
                raw_mobile=director_mobile,
                canonical_mobile=canon_phone,
                contact_validity=p_status,
                director_email=director_email
            )

        AuditEvent.objects.create(
            user=request.user,
            action='CREATE',
            dataset_type='MCA',
            model_name='MCACompany',
            object_id=str(company.id),
            object_repr=f"Manually created MCA company '{name}' ({entity_id})"
        )

    messages.success(request, f"MCA entity '{name}' ({entity_id}) created successfully.")
    return redirect('hub:mca_list')


@login_required
def mca_export_json_view(request):
    """
    Exports filtered or complete MCA records as a structured, indented JSON attachment.
    """
    qs, _, _, _, _, _, _, _ = build_mca_filtered_queryset(request)

    records = []
    for comp in qs:
        comp_dict = {
            'entity_id': comp.entity_id,
            'entity_type': comp.entity_type,
            'name': comp.name,
            'state': comp.state,
            'city': comp.city,
            'roc': comp.roc,
            'nic_code': comp.nic_code,
            'class_of_company': comp.class_of_company,
            'date_of_incorporation': comp.date_of_incorporation.strftime('%Y-%m-%d') if comp.date_of_incorporation else None,
            'paid_up_capital': float(comp.paid_up_capital) if comp.paid_up_capital else None,
            'email': comp.email,
            'registered_address': comp.registered_address,
            'directors': [
                {
                    'director_name': d.director_name,
                    'director_email': d.director_email,
                    'original_phone': d.raw_mobile,
                    'cleaned_phone': d.canonical_mobile,
                    'phone_status': d.contact_validity,
                    'designation': d.designation
                }
                for d in comp.directors.all()
            ]
        }
        records.append(comp_dict)

    import json
    json_str = json.dumps(records, indent=2, ensure_ascii=False)
    timestamp = timezone.now().strftime('%Y%m%d_%H%M%S')
    response = HttpResponse(json_str, content_type='application/json')
    response['Content-Disposition'] = f'attachment; filename="mca_export_{timestamp}.json"'

    AuditEvent.objects.create(
        user=request.user,
        action='EXPORT',
        dataset_type='MCA',
        model_name='MCACompany',
        object_repr=f"Exported {len(records)} MCA records to JSON"
    )
    return response


@login_required
def mca_export_excel_view(request):
    """
    Exports filtered or all MCA records directly to Excel (.xlsx) using formula-injection protection.
    Only records matching active column filters, global search, and sorting are exported.
    """
    scope = request.GET.get('scope', 'filtered')
    if scope == 'all':
        qs = MCACompany.objects.filter(is_archived=False).prefetch_related('directors').order_by('-updated_at', 'name')
    else:
        qs, _, _, _, _, _, _, _ = build_mca_filtered_queryset(request)

    # If no records match, show "No matching records to export."
    if not qs.exists():
        if request.headers.get('x-requested-with') == 'XMLHttpRequest' or request.GET.get('format') == 'json':
            return JsonResponse({'success': False, 'message': 'No matching records to export.'}, status=400)
        messages.warning(request, "No matching records to export.")
        return redirect('hub:mca_list')

    # Generate XLSX directly
    columns = [
        'entity_id', 'entity_type', 'name', 'state', 'city', 'roc', 'date_of_incorporation',
        'paid_up_capital', 'email', 'director_name', 'canonical_mobile', 'contact_validity'
    ]
    file_path = export_mca_dataset(qs, columns=columns, file_format='XLSX')
    filename = os.path.basename(file_path)

    AuditEvent.objects.create(
        user=request.user,
        action='EXPORT',
        dataset_type='MCA',
        model_name='MCACompany',
        object_repr=f"Exported {qs.count()} MCA records to Excel ({scope})"
    )

    with open(file_path, 'rb') as f:
        response = HttpResponse(f.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response


@login_required
def mca_import_json_view(request):
    """
    Imports MCA records directly from a JSON file.
    Validates structure, cleans phone numbers, prevents duplicates, and logs audits.
    """
    if request.method != 'POST':
        return redirect('hub:mca_list')

    json_file = request.FILES.get('json_file')
    if not json_file:
        messages.error(request, "Please choose a JSON file to upload.")
        return redirect('hub:mca_list')

    try:
        import json
        content = json_file.read().decode('utf-8')
        data = json.loads(content)

        if not isinstance(data, list):
            data = [data]

        imported_count = 0
        updated_count = 0

        with transaction.atomic():
            for item in data:
                if not isinstance(item, dict):
                    continue

                entity_id = str(item.get('entity_id') or item.get('cin') or item.get('llpin') or '').strip().upper()
                name = str(item.get('name') or item.get('company_name') or '').strip()
                if not entity_id or not name:
                    continue

                entity_type = str(item.get('entity_type') or 'company').strip().lower()
                if entity_type not in ('company', 'llp'):
                    entity_type = 'company'

                company, created = MCACompany.objects.get_or_create(
                    entity_type=entity_type,
                    entity_id=entity_id,
                    defaults={
                        'name': name,
                        'state': str(item.get('state') or '').strip(),
                        'city': str(item.get('city') or '').strip(),
                        'roc': str(item.get('roc') or '').strip(),
                        'email': str(item.get('email') or '').strip(),
                        'registered_address': str(item.get('registered_address') or item.get('address') or '').strip()
                    }
                )

                if created:
                    imported_count += 1
                else:
                    updated_count += 1

                # Process directors if present
                directors = item.get('directors', [])
                if isinstance(directors, list):
                    for d in directors:
                        if isinstance(d, dict):
                            d_name = str(d.get('director_name') or d.get('name') or '').strip()
                            if not d_name:
                                continue
                            raw_phone = str(d.get('original_phone') or d.get('raw_mobile') or d.get('mobile') or '').strip()
                            canon_phone, p_status, _ = clean_mobile(raw_phone, strict_prefix=True)

                            MCADirectorEntry.objects.get_or_create(
                                company=company,
                                director_name=d_name,
                                defaults={
                                    'raw_mobile': raw_phone,
                                    'canonical_mobile': canon_phone,
                                    'contact_validity': p_status,
                                    'director_email': str(d.get('director_email') or d.get('email') or '').strip()
                                }
                            )

        AuditEvent.objects.create(
            user=request.user,
            action='IMPORT',
            dataset_type='MCA',
            model_name='MCACompany',
            object_repr=f"Imported JSON: {imported_count} new, {updated_count} existing"
        )
        messages.success(request, f"JSON Import complete: {imported_count} companies added, {updated_count} existing verified.")
    except Exception as e:
        messages.error(request, f"Failed to import JSON file: {str(e)}")

    return redirect('hub:mca_list')


@login_required
def mca_clear_all_view(request):
    """
    Clears all MCA records (Companies and linked Directors) within an atomic transaction.
    Strictly protected on the backend: requires admin (is_superuser) privileges
    and typing confirmation 'CLEAR MCA'.
    """
    if not request.user.is_superuser:
        messages.error(request, "Access denied. Only administrators can perform this action.")
        return redirect('hub:mca_list')

    if request.method != 'POST':
        return redirect('hub:mca_list')

    confirm_code = request.POST.get('confirmation_code', '').strip().upper()
    if confirm_code != 'CLEAR MCA':
        messages.error(request, "Confirmation code did not match. Clear All operation aborted.")
        return redirect('hub:mca_list')

    try:
        with transaction.atomic():
            comp_count = MCACompany.objects.count()
            dir_count = MCADirectorEntry.objects.count()

            # Delete all directors and companies
            MCADirectorEntry.objects.all().delete()
            MCACompany.objects.all().delete()

            AuditEvent.objects.create(
                user=request.user,
                action='DELETE',
                dataset_type='MCA',
                model_name='MCACompany',
                object_repr=f"Cleared all MCA records ({comp_count} companies, {dir_count} directors)",
                changes={'companies_cleared': comp_count, 'directors_cleared': dir_count}
            )

        messages.success(request, f"All MCA records cleared successfully ({comp_count} companies and {dir_count} directors removed).")
    except Exception as e:
        messages.error(request, f"Failed to clear MCA records: {str(e)}")

    return redirect('hub:mca_list')


# ==========================================
# GST Module Views
# ==========================================


@login_required
def gst_upload_view(request):
    """
    Dedicated GST-only Upload / Import view.
    Validates .xlsx, .xls, and .csv files, file size, empty/corrupted files,
    detects worksheets, auto-detects headers, rejects MCA datasets,
    creates an ImportBatch for GST, and redirects to the column mapping step.
    """
    if request.method != 'POST':
        return redirect('hub:gst_list')

    uploaded_file = request.FILES.get('data_file')
    sheet_name = request.POST.get('sheet_name', '').strip()

    if not uploaded_file:
        messages.error(request, "Please select a GST dataset file to upload (.xlsx, .xls, or .csv).")
        return redirect('hub:gst_list')

    # Empty file validation (0 bytes)
    if uploaded_file.size == 0:
        messages.error(request, "The selected file is empty (0 bytes). Please upload a valid GST file with data.")
        return redirect('hub:gst_list')

    # Maximum file size validation (50 MB)
    if uploaded_file.size > 50 * 1024 * 1024:
        messages.error(request, "File size exceeds 50 MB limit. Please upload a smaller file.")
        return redirect('hub:gst_list')

    # Extension validation
    ext = uploaded_file.name.split('.')[-1].lower() if '.' in uploaded_file.name else ''
    if ext not in ('xlsx', 'xls', 'csv'):
        messages.error(request, f"Unsupported file format (.{ext}). Only .xlsx, .xls, and .csv files are supported for GST data.")
        return redirect('hub:gst_list')

    # Save uploaded file to private media directory
    upload_dir = os.path.join(settings.MEDIA_ROOT, 'imports')
    os.makedirs(upload_dir, exist_ok=True)
    timestamp = timezone.now().strftime('%Y%m%d_%H%M%S')
    saved_filename = f"gst_{timestamp}_{uploaded_file.name}"
    saved_path = os.path.join(upload_dir, saved_filename)

    try:
        with open(saved_path, 'wb+') as destination:
            for chunk in uploaded_file.chunks():
                destination.write(chunk)
    except Exception as e:
        messages.error(request, f"Failed to save uploaded file: {str(e)}")
        return redirect('hub:gst_list')

    checksum = calculate_file_checksum(saved_path)
    file_format = 'CSV' if ext == 'csv' else ('XLS' if ext == 'xls' else 'XLSX')

    # Worksheet and parsing verification
    csv_delim = ','
    csv_enc = 'utf-8'

    if file_format in ('XLSX', 'XLS'):
        try:
            sheets = get_excel_sheets(saved_path)
            if not sheets:
                raise ValueError("No worksheets found in Excel workbook.")

            if sheet_name and sheet_name in sheets:
                active_sheet = sheet_name
            else:
                if sheet_name:
                    messages.warning(request, f"Worksheet '{sheet_name}' was not found. Defaulted to '{sheets[0]}'.")
                active_sheet = sheets[0]
                sheet_name = active_sheet

            # Test reading rows to verify the workbook is valid and not corrupted
            test_rows = read_excel_rows(saved_path, sheet_name=sheet_name, max_rows=5)
            if not test_rows or all(not any(cell is not None and str(cell).strip() for cell in r) for r in test_rows):
                if os.path.exists(saved_path):
                    os.remove(saved_path)
                messages.error(request, "The uploaded Excel file contains no readable data or rows.")
                return redirect('hub:gst_list')
        except Exception as e:
            if os.path.exists(saved_path):
                os.remove(saved_path)
            messages.error(request, f"Corrupted or invalid Excel file. The file could not be parsed: {str(e)}")
            return redirect('hub:gst_list')
    else:
        # CSV format
        det_delim, det_enc = detect_csv_properties(saved_path)
        csv_delim = det_delim or ','
        csv_enc = det_enc or 'utf-8'
        try:
            with open(saved_path, 'r', encoding=csv_enc, errors='replace') as f:
                reader = csv.reader(f, delimiter=csv_delim)
                test_rows = []
                for _ in range(5):
                    try:
                        row = next(reader)
                        test_rows.append(row)
                    except StopIteration:
                        break
            if not test_rows or all(not any(str(cell).strip() for cell in r) for r in test_rows):
                if os.path.exists(saved_path):
                    os.remove(saved_path)
                messages.error(request, "The uploaded CSV file contains no readable data or rows.")
                return redirect('hub:gst_list')
        except Exception as e:
            if os.path.exists(saved_path):
                os.remove(saved_path)
            messages.error(request, f"Corrupted or invalid CSV file: {str(e)}")
            return redirect('hub:gst_list')

    # Auto-detect header row
    auto_hr = auto_detect_header_row(
        file_path=saved_path,
        file_format=file_format,
        sheet_name=sheet_name,
        delimiter=csv_delim,
        encoding=csv_enc
    )

    # Read headers
    headers = read_file_headers(
        file_path=saved_path,
        file_format=file_format,
        sheet_name=sheet_name,
        header_row=auto_hr,
        delimiter=csv_delim,
        encoding=csv_enc
    )

    if not headers:
        if os.path.exists(saved_path):
            os.remove(saved_path)
        messages.error(request, "Could not detect headers in file. Please ensure the file contains valid column headers.")
        return redirect('hub:gst_list')

    # Detect dataset type and enforce schema isolation (Reject MCA files in GST)
    detected_type = detect_dataset_type(headers)
    if detected_type == 'MCA':
        if os.path.exists(saved_path):
            os.remove(saved_path)
        messages.error(
            request,
            "MCA file detected: The uploaded file contains MCA corporate columns (CIN, ROC, Director) "
            "and cannot be imported into the GST module. Please upload a valid GST registration file or use the MCA Import module."
        )
        return redirect('hub:gst_list')

    # Create dedicated GST ImportBatch
    batch_id = f"BATCH-GST-{timestamp}-{checksum[:8].upper()}"
    batch = ImportBatch.objects.create(
        batch_id=batch_id,
        dataset_type='GST',
        file_name=uploaded_file.name,
        file_path=saved_path,
        file_checksum=checksum,
        file_format=file_format,
        sheet_name=sheet_name,
        header_row_index=auto_hr,
        status='UPLOADED',
        created_by=request.user,
        mapping_config={
            'delimiter': csv_delim,
            'encoding': csv_enc,
            'sheet_name': sheet_name
        }
    )

    AuditEvent.objects.create(
        user=request.user,
        action='CREATE',
        dataset_type='GST',
        model_name='ImportBatch',
        object_id=batch.batch_id,
        object_repr=f"Uploaded GST file {uploaded_file.name}",
        changes={'file_size': uploaded_file.size, 'checksum': checksum}
    )

    messages.success(request, f"GST file '{uploaded_file.name}' uploaded successfully. Please confirm column mapping.")
    return redirect('hub:import_map', batch_id=batch.batch_id)


def build_gst_filtered_queryset(request):
    """
    Builds a secure, parameterized, whitelisted queryset for GST records.
    Supports individual column filters (f_*) and general search with strict AND logic.
    """
    qs = GSTRecord.objects.filter(is_archived=False)
    col_filters = {
        'f_gstin': '',
        'f_trade_name': '',
        'f_legal_name': '',
        'f_name': '',
        'f_status': '',
        'f_source_type': '',
        'f_pan': '',
        'f_phone': '',
        'f_email': '',
        'f_state': '',
        'f_city': '',
        'f_address': '',
        'f_date': '',
    }

    # 1. Whitelisted Column Filters
    # GSTIN (treated as text, preserving leading zeros)
    f_gstin = (request.GET.get('f_gstin') or request.GET.get('f_gst_number') or '').strip()
    if f_gstin:
        col_filters['f_gstin'] = f_gstin
        qs = qs.filter(gst_number__icontains=f_gstin)

    # Trade Name
    f_trade_name = (request.GET.get('f_trade_name') or '').strip()
    if f_trade_name:
        col_filters['f_trade_name'] = f_trade_name
        qs = qs.filter(trade_name__icontains=f_trade_name)

    # Legal Name
    f_legal_name = (request.GET.get('f_legal_name') or '').strip()
    if f_legal_name:
        col_filters['f_legal_name'] = f_legal_name
        qs = qs.filter(legal_name__icontains=f_legal_name)

    # Combined Name (matches trade OR legal)
    f_name = (request.GET.get('f_name') or '').strip()
    if f_name:
        col_filters['f_name'] = f_name
        qs = qs.filter(Q(trade_name__icontains=f_name) | Q(legal_name__icontains=f_name))

    # Status
    f_status = (request.GET.get('f_status') or '').strip()
    if f_status:
        col_filters['f_status'] = f_status
        qs = qs.filter(status__icontains=f_status)

    # Type (source_type)
    f_source_type = (request.GET.get('f_source_type') or request.GET.get('f_type') or '').strip()
    if f_source_type:
        col_filters['f_source_type'] = f_source_type
        qs = qs.filter(source_type__icontains=f_source_type)

    # PAN
    f_pan = (request.GET.get('f_pan') or '').strip()
    if f_pan:
        col_filters['f_pan'] = f_pan
        qs = qs.filter(pan__icontains=f_pan)

    # Phone / Mobile (treated as text, preserving leading zeros)
    f_phone = (request.GET.get('f_phone') or request.GET.get('f_mobile') or '').strip()
    if f_phone:
        col_filters['f_phone'] = f_phone
        qs = qs.filter(
            Q(canonical_mobile__icontains=f_phone) |
            Q(raw_mobile__icontains=f_phone)
        )

    # Email
    f_email = (request.GET.get('f_email') or '').strip()
    if f_email:
        col_filters['f_email'] = f_email
        qs = qs.filter(email__icontains=f_email)

    # State
    f_state = (request.GET.get('f_state') or '').strip()
    if f_state:
        col_filters['f_state'] = f_state
        qs = qs.filter(state__icontains=f_state)

    # City
    f_city = (request.GET.get('f_city') or '').strip()
    if f_city:
        col_filters['f_city'] = f_city
        qs = qs.filter(city__icontains=f_city)

    # Address
    f_address = (request.GET.get('f_address') or '').strip()
    if f_address:
        col_filters['f_address'] = f_address
        qs = qs.filter(address__icontains=f_address)

    # 2. General search (q) and top filters
    search = request.GET.get('q', '').strip()
    if search:
        qs = qs.filter(
            Q(gst_number__icontains=search) |
            Q(legal_name__icontains=search) |
            Q(trade_name__icontains=search) |
            Q(pan__icontains=search) |
            Q(canonical_mobile__icontains=search) |
            Q(raw_mobile__icontains=search) |
            Q(email__icontains=search) |
            Q(address__icontains=search) |
            Q(city__icontains=search) |
            Q(state__icontains=search) |
            Q(jurisdiction__icontains=search) |
            Q(state_jurisdiction__icontains=search) |
            Q(status__icontains=search)
        )

    status_filter = request.GET.get('status', '').strip()
    if status_filter and not f_status:
        qs = qs.filter(status__iexact=status_filter)

    state = request.GET.get('state', '').strip()
    if state and not f_state:
        qs = qs.filter(state__iexact=state)

    phone_filter = request.GET.get('phone_status', '').strip()
    if phone_filter:
        pf = phone_filter.strip().lower()
        if pf in ('valid',):
            qs = qs.filter(contact_validity__in=['valid', 'cleaned', 'VALID'])
        elif pf in ('invalid', 'flagged'):
            qs = qs.filter(contact_validity__in=['invalid', 'multiple_numbers_review', 'FLAGGED'])
        elif pf in ('missing', 'blank'):
            qs = qs.filter(contact_validity__in=['missing', 'BLANK'])
        else:
            qs = qs.filter(contact_validity__iexact=phone_filter)

    # 3. Whitelisted Sorting
    sort_param = request.GET.get('sort', '').strip()
    allowed_sorts = {
        'gstin': 'gst_number', '-gstin': '-gst_number', 'gstin_asc': 'gst_number', 'gstin_desc': '-gst_number',
        'trade_name': 'trade_name', '-trade_name': '-trade_name', 'trade_name_asc': 'trade_name', 'trade_name_desc': '-trade_name',
        'legal_name': 'legal_name', '-legal_name': '-legal_name', 'legal_name_asc': 'legal_name', 'legal_name_desc': '-legal_name',
        'status': 'status', '-status': '-status', 'status_asc': 'status', 'status_desc': '-status',
        'pan': 'pan', '-pan': '-pan', 'pan_asc': 'pan', 'pan_desc': '-pan',
        'state': 'state', '-state': '-state', 'state_asc': 'state', 'state_desc': '-state',
        'city': 'city', '-city': '-city', 'city_asc': 'city', 'city_desc': '-city',
        'date': 'registration_date', '-date': '-registration_date', 'date_asc': 'registration_date', 'date_desc': '-registration_date',
        'updated_at': '-updated_at',
    }
    if sort_param in allowed_sorts:
        qs = qs.order_by(allowed_sorts[sort_param])
    else:
        qs = qs.order_by('-updated_at', 'gst_number')

    return qs, col_filters, search, status_filter, state, phone_filter, sort_param


@login_required
def gst_list_view(request):
    qs, col_filters, search, status_filter, state, phone_filter, sort_param = build_gst_filtered_queryset(request)
    total_count = qs.count()

    paginator = Paginator(qs, 20)
    page_obj = paginator.get_page(request.GET.get('page', 1))

    states = GSTRecord.objects.values_list('state', flat=True).distinct().order_by('state')[:50]
    statuses = GSTRecord.objects.values_list('status', flat=True).distinct().order_by('status')[:20]

    total_db_count = GSTRecord.objects.count()
    context = {
        'page_obj': page_obj,
        'search': search,
        'status_filter': status_filter,
        'state': state,
        'phone_filter': phone_filter,
        'states': [s for s in states if s],
        'statuses': [st for st in statuses if st],
        'total_count': total_count,
        'total_gst_records': total_db_count,
        'has_gst_records': total_db_count > 0,
        'col_filters': col_filters,
        'sort': sort_param,
        'has_active_filters': bool(any(col_filters.values()) or search or status_filter or state or phone_filter),
    }

    if request.headers.get('x-requested-with') == 'XMLHttpRequest' or request.GET.get('format') == 'partial':
        html = render_to_string('gst/_table_content.html', context, request=request)
        return JsonResponse({
            'success': True,
            'html': html,
            'total_count': total_count,
            'start_index': page_obj.start_index() if total_count > 0 else 0,
            'end_index': page_obj.end_index() if total_count > 0 else 0,
            'has_other_pages': page_obj.has_other_pages(),
            'page': page_obj.number,
            'num_pages': paginator.num_pages,
        })

    return render(request, 'gst/list.html', context)


@login_required
def gst_detail_view(request, gst_id):
    gst = get_object_or_404(GSTRecord, id=gst_id)
    audit_events = AuditEvent.objects.filter(
        model_name='GSTRecord', object_id=str(gst.id)
    )[:10]
    total_db_count = GSTRecord.objects.count()
    return render(request, 'gst/detail.html', {
        'gst': gst,
        'audit_events': audit_events,
        'total_gst_records': total_db_count,
        'has_gst_records': total_db_count > 0,
    })


@login_required
def gst_edit_view(request, gst_id):
    gst = get_object_or_404(GSTRecord, id=gst_id)
    if request.method == 'POST':
        form = GSTRecordForm(request.POST, instance=gst)
        if form.is_valid():
            form.save()
            gst.revision += 1
            gst.save()
            
            AuditEvent.objects.create(
                user=request.user,
                action='UPDATE',
                dataset_type='GST',
                model_name='GSTRecord',
                object_id=str(gst.id),
                object_repr=f"Updated GST {gst.gst_number}",
                changes={'trade_name': gst.trade_name}
            )
            messages.success(request, "GST record updated successfully.")
            return redirect('hub:gst_detail', gst_id=gst.id)
    else:
        form = GSTRecordForm(instance=gst)
    total_db_count = GSTRecord.objects.count()
    return render(request, 'gst/edit.html', {
        'form': form,
        'gst': gst,
        'total_gst_records': total_db_count,
        'has_gst_records': total_db_count > 0,
    })


@login_required
def gst_delete_view(request, gst_id):
    """
    Safely delete a GST registration record within an atomic transaction.
    Strictly protected on the backend: requires admin (is_superuser) privileges.
    Returns JSON for AJAX/fetch requests and redirects with flash messages for standard form submissions.
    """
    is_ajax = request.headers.get('x-requested-with') == 'XMLHttpRequest' or 'application/json' in request.headers.get('Accept', '')

    if not request.user.is_superuser:
        msg = "Access denied. Only administrators have permission to delete records."
        if is_ajax:
            return JsonResponse({'success': False, 'message': msg}, status=403)
        messages.error(request, msg)
        return redirect('hub:gst_list')

    if request.method != 'POST':
        msg = "Invalid request method. Delete must be performed via POST."
        if is_ajax:
            return JsonResponse({'success': False, 'message': msg}, status=405)
        messages.error(request, msg)
        return redirect('hub:gst_list')

    try:
        gst = GSTRecord.objects.get(id=gst_id)
    except GSTRecord.DoesNotExist:
        msg = "This GST record does not exist or has already been deleted."
        if is_ajax:
            return JsonResponse({'success': False, 'already_deleted': True, 'message': msg}, status=404)
        messages.warning(request, msg)
        return redirect('hub:gst_list')

    gst_number = gst.gst_number
    trade_or_legal_name = gst.trade_name or gst.legal_name or gst_number

    try:
        with transaction.atomic():
            # Delete any attached BusinessLink entries for this GST record
            BusinessLink.objects.filter(gst_record=gst).delete()

            # Record audit event before deleting parent
            AuditEvent.objects.create(
                user=request.user,
                action='DELETE',
                dataset_type='GST',
                model_name='GSTRecord',
                object_id=str(gst.id),
                object_repr=f"Deleted GST Record {gst_number} ({trade_or_legal_name})",
                changes={'gst_number': gst_number, 'name': trade_or_legal_name}
            )

            # Delete the GST record itself
            gst.delete()

        # Compute updated counts
        total_gst = GSTRecord.objects.filter(is_archived=False).count()

        success_msg = f"GST registration '{trade_or_legal_name}' ({gst_number}) has been deleted successfully."

        if is_ajax:
            return JsonResponse({
                'success': True,
                'message': success_msg,
                'record_id': gst_id,
                'stats': {
                    'total_gst': total_gst
                }
            })

        messages.success(request, success_msg)
        return redirect('hub:gst_list')

    except Exception as e:
        err_msg = f"Failed to delete record due to a server error: {str(e)}"
        if is_ajax:
            return JsonResponse({'success': False, 'message': err_msg}, status=500)
        messages.error(request, err_msg)
        return redirect('hub:gst_list')


# ==========================================
# Dedicated GST Toolbar Action Views
# ==========================================

@login_required
def gst_create_view(request):
    """
    Directly creates a GST Registration with validation, phone cleaning, and auditing.
    """
    if request.method != 'POST':
        return redirect('hub:gst_list')

    gst_number = request.POST.get('gst_number', '').strip().upper()
    trade_name = request.POST.get('trade_name', '').strip()
    legal_name = request.POST.get('legal_name', '').strip()
    pan = request.POST.get('pan', '').strip().upper()
    mobile = request.POST.get('mobile', '').strip()
    email = request.POST.get('email', '').strip()
    status = request.POST.get('status', 'ACT').strip().upper()
    state = request.POST.get('state', '').strip()
    city = request.POST.get('city', '').strip()
    pincode = request.POST.get('pincode', '').strip()
    address = request.POST.get('address', '').strip()
    nature_of_business = request.POST.get('nature_of_business', '').strip()
    jurisdiction = request.POST.get('jurisdiction', '').strip()
    reg_date_raw = request.POST.get('registration_date', '').strip()

    if not gst_number:
        messages.error(request, "GSTIN (GST Number) is required.")
        return redirect('hub:gst_list')

    if GSTRecord.objects.filter(gst_number=gst_number).exists():
        messages.error(request, f"A GST record with GSTIN '{gst_number}' already exists.")
        return redirect('hub:gst_list')

    reg_date = None
    if reg_date_raw:
        try:
            reg_date = datetime.datetime.strptime(reg_date_raw, '%Y-%m-%d').date()
        except ValueError:
            pass

    canon_mobile, phone_status, _ = clean_mobile(mobile, strict_prefix=True)

    with transaction.atomic():
        gst = GSTRecord.objects.create(
            gst_number=gst_number,
            trade_name=trade_name,
            legal_name=legal_name,
            pan=pan,
            raw_mobile=mobile,
            canonical_mobile=canon_mobile,
            contact_validity=phone_status,
            email=email,
            status=status,
            state=state,
            city=city,
            pincode=pincode,
            address=address,
            nature_of_business=nature_of_business,
            jurisdiction=jurisdiction,
            registration_date=reg_date
        )

        AuditEvent.objects.create(
            user=request.user,
            action='CREATE',
            dataset_type='GST',
            model_name='GSTRecord',
            object_id=str(gst.id),
            object_repr=f"Manually created GST record '{gst_number}' ({trade_name or legal_name})"
        )

    messages.success(request, f"GST record '{gst_number}' created successfully.")
    return redirect('hub:gst_list')


@login_required
def gst_export_json_view(request):
    """
    Exports filtered or complete GST records as a formatted JSON attachment.
    """
    qs, _, _, _, _, _, _ = build_gst_filtered_queryset(request)

    records = []
    for g in qs:
        records.append({
            'gst_number': g.gst_number,
            'legal_name': g.legal_name,
            'trade_name': g.trade_name,
            'pan': g.pan,
            'original_phone': g.raw_mobile,
            'cleaned_phone': g.canonical_mobile,
            'phone_status': g.contact_validity,
            'email': g.email,
            'status': g.status,
            'state': g.state,
            'city': g.city,
            'pincode': g.pincode,
            'address': g.address,
            'nature_of_business': g.nature_of_business,
            'jurisdiction': g.jurisdiction,
            'registration_date': g.registration_date.strftime('%Y-%m-%d') if g.registration_date else None
        })

    import json
    json_str = json.dumps(records, indent=2, ensure_ascii=False)
    timestamp = timezone.now().strftime('%Y%m%d_%H%M%S')
    response = HttpResponse(json_str, content_type='application/json')
    response['Content-Disposition'] = f'attachment; filename="gst_export_{timestamp}.json"'

    AuditEvent.objects.create(
        user=request.user,
        action='EXPORT',
        dataset_type='GST',
        model_name='GSTRecord',
        object_repr=f"Exported {len(records)} GST records to JSON"
    )
    return response


@login_required
def gst_export_excel_view(request):
    """
    Exports filtered or all GST records directly to Excel (.xlsx) using formula-injection protection.
    Only records matching active column filters, global search, and sorting are exported.
    """
    scope = request.GET.get('scope', 'filtered')
    if scope == 'all':
        qs = GSTRecord.objects.filter(is_archived=False).order_by('-updated_at')
    else:
        qs, _, _, _, _, _, _ = build_gst_filtered_queryset(request)

    # If no records match, show "No matching records to export."
    if not qs.exists():
        if request.headers.get('x-requested-with') == 'XMLHttpRequest' or request.GET.get('format') == 'json':
            return JsonResponse({'success': False, 'message': 'No matching records to export.'}, status=400)
        messages.warning(request, "No matching records to export.")
        return redirect('hub:gst_list')

    columns = [
        'gst_number', 'legal_name', 'trade_name', 'status', 'source_type', 'pan',
        'canonical_mobile', 'raw_mobile', 'contact_validity',
        'email', 'state', 'city', 'pincode', 'registration_date'
    ]
    file_path = export_gst_dataset(qs, columns=columns, file_format='XLSX')
    filename = os.path.basename(file_path)

    AuditEvent.objects.create(
        user=request.user,
        action='EXPORT',
        dataset_type='GST',
        model_name='GSTRecord',
        object_repr=f"Exported {qs.count()} GST records to Excel ({scope})"
    )

    with open(file_path, 'rb') as f:
        response = HttpResponse(f.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response


@login_required
def gst_export_csv_view(request):
    """
    Exports filtered or all GST records directly to CSV (.csv) using formula-injection protection.
    Only records matching active column filters, global search, and sorting are exported.
    """
    scope = request.GET.get('scope', 'filtered')
    if scope == 'all':
        qs = GSTRecord.objects.filter(is_archived=False).order_by('-updated_at')
    else:
        qs, _, _, _, _, _, _ = build_gst_filtered_queryset(request)

    if not qs.exists():
        if request.headers.get('x-requested-with') == 'XMLHttpRequest' or request.GET.get('format') == 'json':
            return JsonResponse({'success': False, 'message': 'No matching records to export.'}, status=400)
        messages.warning(request, "No matching records to export.")
        return redirect('hub:gst_list')

    columns = [
        'gst_number', 'legal_name', 'trade_name', 'status', 'source_type', 'pan',
        'canonical_mobile', 'raw_mobile', 'contact_validity',
        'email', 'state', 'city', 'pincode', 'registration_date'
    ]
    file_path = export_gst_dataset(qs, columns=columns, file_format='CSV')
    filename = os.path.basename(file_path)

    AuditEvent.objects.create(
        user=request.user,
        action='EXPORT',
        dataset_type='GST',
        model_name='GSTRecord',
        object_repr=f"Exported {qs.count()} GST records to CSV ({scope})"
    )

    with open(file_path, 'rb') as f:
        response = HttpResponse(f.read(), content_type='text/csv; charset=utf-8')
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response


@login_required
def gst_import_json_view(request):
    """
    Imports GST records directly from a JSON file with validation and phone cleaning.
    """
    if request.method != 'POST':
        return redirect('hub:gst_list')

    json_file = request.FILES.get('json_file')
    if not json_file:
        messages.error(request, "Please choose a JSON file to upload.")
        return redirect('hub:gst_list')

    try:
        import json
        content = json_file.read().decode('utf-8')
        data = json.loads(content)

        if not isinstance(data, list):
            data = [data]

        imported_count = 0
        updated_count = 0

        with transaction.atomic():
            for item in data:
                if not isinstance(item, dict):
                    continue

                gst_number = str(item.get('gst_number') or item.get('gstin') or item.get('gst_no') or '').strip().upper()
                if not gst_number:
                    continue

                raw_phone = str(item.get('original_phone') or item.get('raw_mobile') or item.get('mobile') or '').strip()
                canon_phone, p_status, _ = clean_mobile(raw_phone, strict_prefix=True)

                rec, created = GSTRecord.objects.get_or_create(
                    gst_number=gst_number,
                    defaults={
                        'legal_name': str(item.get('legal_name') or '').strip(),
                        'trade_name': str(item.get('trade_name') or item.get('name') or '').strip(),
                        'pan': str(item.get('pan') or '').strip().upper(),
                        'raw_mobile': raw_phone,
                        'canonical_mobile': canon_phone,
                        'contact_validity': p_status,
                        'email': str(item.get('email') or '').strip(),
                        'status': str(item.get('status') or 'ACT').strip().upper(),
                        'state': str(item.get('state') or '').strip(),
                        'city': str(item.get('city') or '').strip(),
                        'pincode': str(item.get('pincode') or '').strip(),
                        'address': str(item.get('address') or '').strip(),
                        'nature_of_business': str(item.get('nature_of_business') or '').strip(),
                        'jurisdiction': str(item.get('jurisdiction') or '').strip(),
                    }
                )

                if created:
                    imported_count += 1
                else:
                    updated_count += 1

        AuditEvent.objects.create(
            user=request.user,
            action='IMPORT',
            dataset_type='GST',
            model_name='GSTRecord',
            object_repr=f"Imported GST JSON: {imported_count} new, {updated_count} existing"
        )
        messages.success(request, f"JSON Import complete: {imported_count} GST records added, {updated_count} existing verified.")
    except Exception as e:
        messages.error(request, f"Failed to import JSON file: {str(e)}")

    return redirect('hub:gst_list')


@login_required
def gst_clear_all_view(request):
    """
    Clears all GST records and related GST import/cleaning data within an atomic transaction.
    - Strictly protected on the backend: requires admin (is_superuser) privileges.
    - Preserves all MCA companies, MCA directors, users, audit logs, and system settings.
    - Records the action in AuditEvent with admin name, timestamp, deleted count, and module name.
    - Supports both AJAX/JSON requests and standard POST form redirects.
    - Fully atomic: rolls back if any exception occurs.
    """
    is_ajax = request.headers.get('x-requested-with') == 'XMLHttpRequest' or 'application/json' in request.headers.get('Accept', '')

    if not request.user.is_superuser:
        msg = "Access denied. Only administrators have permission to clear GST records."
        if is_ajax:
            return JsonResponse({'success': False, 'message': msg}, status=403)
        messages.error(request, msg)
        return redirect('hub:gst_list')

    if request.method != 'POST':
        msg = "Invalid request method. Clear All must be submitted via POST."
        if is_ajax:
            return JsonResponse({'success': False, 'message': msg}, status=405)
        messages.error(request, msg)
        return redirect('hub:gst_list')

    try:
        gst_count = GSTRecord.objects.count()
        if gst_count == 0:
            msg = "There are no GST records in the database to clear."
            if is_ajax:
                return JsonResponse({'success': True, 'cleared_count': 0, 'message': msg, 'redirect_url': reverse('hub:gst_list')})
            messages.info(request, msg)
            return redirect('hub:gst_list')

        with transaction.atomic():
            # 1. Clean up cross-references in BusinessLink
            BusinessLink.objects.all().delete()

            # 2. Delete GST Data Quality Issues (cleaning / validation issues)
            gst_quality_issues = QualityIssue.objects.filter(
                Q(dataset_type='GST') | Q(record_type='GST')
            )
            quality_count = gst_quality_issues.count()
            gst_quality_issues.delete()

            # 3. Delete GST Import Batches (cascades to StagedRow raw/cleaned import staging data)
            gst_batches = ImportBatch.objects.filter(dataset_type='GST')
            batches_count = gst_batches.count()
            gst_batches.delete()

            # 4. Delete all GST records
            GSTRecord.objects.all().delete()

            # 5. Record AuditEvent with admin name, timestamp, deleted count, and module name
            admin_name = request.user.get_full_name() or request.user.username
            AuditEvent.objects.create(
                user=request.user,
                action='DELETE',
                dataset_type='GST',
                model_name='GSTRecord',
                object_id='ALL',
                object_repr=f"Cleared all GST data: {gst_count} records removed",
                changes={
                    'module': 'GST',
                    'module_name': 'GST',
                    'admin_name': admin_name,
                    'admin_username': request.user.username,
                    'deleted_count': gst_count,
                    'gst_records_deleted': gst_count,
                    'quality_issues_deleted': quality_count,
                    'import_batches_deleted': batches_count,
                }
            )

        success_msg = f"Successfully cleared all {gst_count} GST record(s). All search filters, pagination, and selections have been reset."
        if is_ajax:
            return JsonResponse({
                'success': True,
                'cleared_count': gst_count,
                'message': success_msg,
                'redirect_url': reverse('hub:gst_list'),
            })

        messages.success(request, success_msg)
        return redirect('hub:gst_list')

    except Exception as e:
        logger.error(f"Failed to clear GST records: {str(e)}", exc_info=True)
        err_msg = f"Failed to clear GST records: {str(e)}"
        if is_ajax:
            return JsonResponse({'success': False, 'message': err_msg}, status=500)
        messages.error(request, err_msg)
        return redirect('hub:gst_list')


# ==========================================
# Global Search View
# ==========================================

@login_required
def global_search_view(request):
    query = request.GET.get('q', '').strip()
    dataset = request.GET.get('dataset', '')
    
    results = perform_unified_search(query, dataset_filter=dataset or None, limit=40)
    
    context = {
        'query': query,
        'dataset': dataset,
        'results': results,
    }
    return render(request, 'search/global_search.html', context)


# ==========================================
# Quality Inbox & Resolution
# ==========================================

@login_required
def quality_inbox_view(request):
    issues = QualityIssue.objects.all().order_by('-created_at')
    
    dataset = request.GET.get('dataset')
    status = request.GET.get('status')
    error_code = request.GET.get('error_code')
    
    if dataset:
        issues = issues.filter(dataset_type=dataset)
    if status == 'resolved':
        issues = issues.filter(is_resolved=True)
    elif status == 'open' or not status:
        issues = issues.filter(is_resolved=False)
        
    if error_code:
        issues = issues.filter(error_code=error_code)
        
    paginator = Paginator(issues, 25)
    page_obj = paginator.get_page(request.GET.get('page', 1))
    
    error_codes = QualityIssue.objects.values_list('error_code', flat=True).distinct()
    total_inbox_items = QualityIssue.objects.count()
    
    context = {
        'page_obj': page_obj,
        'dataset': dataset,
        'status': status or 'open',
        'error_code': error_code,
        'error_codes': error_codes,
        'total_count': issues.count(),
        'total_inbox_items': total_inbox_items,
    }
    return render(request, 'quality/inbox.html', context)


@login_required
def quality_resolve_view(request, issue_id):
    if request.method == 'POST':
        note = request.POST.get('note', '')
        corrected_phone = request.POST.get('corrected_phone', '').strip() or None
        success, err_msg = resolve_quality_issue(issue_id, request.user, resolution_note=note, corrected_phone=corrected_phone)
        if success:
            msg = f"Quality issue #{issue_id} marked as resolved."
            if corrected_phone:
                msg += " Phone number successfully corrected and validated."
            messages.success(request, msg)
        else:
            messages.error(request, f"Could not resolve issue: {err_msg or 'Validation failed'}")
    return redirect('hub:quality_inbox')


@login_required
def quality_clear_all_view(request):
    """
    Safely clears all QualityInbox items from the database.
    - Strictly protected on the backend: requires admin (is_superuser) privileges.
    - Preserves all MCA, GST, company, director, user, import history, and audit-log records.
    - Records the action in AuditEvent with admin user, timestamp, and count of removed items.
    - Supports both AJAX/JSON requests and standard POST form redirects.
    """
    is_ajax = request.headers.get('x-requested-with') == 'XMLHttpRequest' or 'application/json' in request.headers.get('Accept', '')

    if not request.user.is_superuser:
        msg = "Access denied. Only administrators have permission to clear the Quality Inbox."
        if is_ajax:
            return JsonResponse({'success': False, 'message': msg}, status=403)
        messages.error(request, msg)
        return redirect('hub:quality_inbox')

    if request.method != 'POST':
        msg = "Invalid request method. Clear All must be submitted via POST."
        if is_ajax:
            return JsonResponse({'success': False, 'message': msg}, status=405)
        messages.error(request, msg)
        return redirect('hub:quality_inbox')

    try:
        count = QualityIssue.objects.count()
        if count == 0:
            msg = "The Quality Inbox is already empty. No items were cleared."
            if is_ajax:
                return JsonResponse({'success': True, 'cleared_count': 0, 'message': msg})
            messages.info(request, msg)
            return redirect('hub:quality_inbox')

        with transaction.atomic():
            QualityIssue.objects.all().delete()

            AuditEvent.objects.create(
                user=request.user,
                action='DELETE',
                dataset_type='SYSTEM',
                model_name='QualityIssue',
                object_id='ALL',
                object_repr=f"Cleared all {count} items from Quality Inbox",
                changes={'cleared_count': count, 'scope': 'QualityIssue.all()'}
            )

        success_msg = f"Successfully cleared all {count} item(s) from the Quality Inbox."
        if is_ajax:
            return JsonResponse({'success': True, 'cleared_count': count, 'message': success_msg})

        messages.success(request, success_msg)
        return redirect('hub:quality_inbox')

    except Exception as e:
        err_msg = f"Failed to clear Quality Inbox items: {str(e)}"
        if is_ajax:
            return JsonResponse({'success': False, 'message': err_msg}, status=500)
        messages.error(request, err_msg)
        return redirect('hub:quality_inbox')


# ==========================================
# Exports Management
# ==========================================

@login_required
def export_list_view(request):
    jobs = ExportJob.objects.filter(requester=request.user).order_by('-created_at')
    form = ExportRequestForm()
    
    if request.method == 'POST':
        form = ExportRequestForm(request.POST)
        if form.is_valid():
            dataset_type = form.cleaned_data['dataset_type']
            file_format = form.cleaned_data['file_format']
            state = form.cleaned_data['state_filter']
            status = form.cleaned_data['status_filter']
            
            job = ExportJob.objects.create(
                requester=request.user,
                dataset_type=dataset_type,
                file_format=file_format,
                columns=[],  # Use defaults
                filters={'state': state, 'status': status},
                status='PENDING'
            )
            # Execute export synchronously / eager Celery
            try:
                run_export_job(job)
                messages.success(request, f"Export completed successfully ({job.row_count} records).")
            except Exception as e:
                messages.error(request, f"Export failed: {str(e)}")
                
            return redirect('hub:export_list')
            
    return render(request, 'exports/list.html', {'jobs': jobs, 'form': form})


@login_required
def export_download_view(request, job_id):
    job = get_object_or_404(ExportJob, id=job_id)
    if not job.file_path or not os.path.exists(job.file_path):
        raise Http404("Export file not found or expired")
        
    filename = os.path.basename(job.file_path)
    content_type = 'text/csv' if job.file_format == 'CSV' else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    
    with open(job.file_path, 'rb') as f:
        response = HttpResponse(f.read(), content_type=content_type)
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response


# ==========================================
# Audit History
# ==========================================

@login_required
def audit_history_view(request):
    events = AuditEvent.objects.all().order_by('-timestamp')
    action = request.GET.get('action')
    dataset = request.GET.get('dataset')
    
    if action:
        events = events.filter(action=action)
    if dataset:
        events = events.filter(dataset_type=dataset)
        
    paginator = Paginator(events, 30)
    page_obj = paginator.get_page(request.GET.get('page', 1))
    
    context = {
        'page_obj': page_obj,
        'action': action,
        'dataset': dataset,
    }
    return render(request, 'audit/list.html', context)


# ==========================================
# Administrator Settings
# ==========================================

@login_required
def settings_view(request):
    templates = MappingTemplate.objects.all()
    strict_phone_setting, _ = SystemSetting.objects.get_or_create(
        key='STRICT_PHONE_PREFIX',
        defaults={'value': 'true', 'description': 'Require Indian domestic mobile numbers to begin with 6, 7, 8, or 9'}
    )
    
    if request.method == 'POST' and request.user.is_superuser:
        strict_val = 'true' if request.POST.get('strict_phone') == 'on' else 'false'
        strict_phone_setting.value = strict_val
        strict_phone_setting.save()
        messages.success(request, "System settings updated successfully.")
        return redirect('hub:settings')
        
    password_form = AdminPasswordChangeForm(user=request.user) if request.user.is_superuser else None

    context = {
        'templates': templates,
        'strict_phone': strict_phone_setting.value.lower() in ('true', '1', 'yes'),
        'password_form': password_form,
    }
    return render(request, 'settings/settings.html', context)


@login_required
def admin_change_password_view(request):
    """
    Dedicated endpoint for Administrator Password Change.
    Strictly restricted to users with administrator privileges (is_superuser).
    Enforces current password verification, complex password rules,
    auditing, and session invalidation/re-login.
    """
    if not request.user.is_superuser:
        messages.error(request, "Access denied. Only administrators can perform this action.")
        return redirect('hub:dashboard')

    if request.method != 'POST':
        return redirect('hub:settings')

    form = AdminPasswordChangeForm(user=request.user, data=request.POST)
    if form.is_valid():
        new_password = form.cleaned_data['new_password']
        user = request.user
        user.set_password(new_password)
        user.save()

        # Audit Event for security tracking
        AuditEvent.objects.create(
            user=user,
            action='UPDATE',
            dataset_type='SYSTEM',
            model_name='AdminUser',
            object_id=str(user.id),
            object_repr=f"Password changed for admin '{user.username}'",
            changes={'action': 'password_change'},
            ip_address=request.META.get('REMOTE_ADDR')
        )

        # Invalidate existing sessions: Logging out invalidates the current session token
        logout(request)
        messages.success(request, "Password changed successfully! Existing sessions were invalidated. Please sign in with your new password.")
        return redirect('hub:login')
    else:
        # Re-render settings page with form errors
        templates = MappingTemplate.objects.all()
        strict_phone_setting, _ = SystemSetting.objects.get_or_create(
            key='STRICT_PHONE_PREFIX',
            defaults={'value': 'true', 'description': 'Require Indian domestic mobile numbers to begin with 6, 7, 8, or 9'}
        )
        context = {
            'templates': templates,
            'strict_phone': strict_phone_setting.value.lower() in ('true', '1', 'yes'),
            'password_form': form,
        }
        return render(request, 'settings/settings.html', context)
