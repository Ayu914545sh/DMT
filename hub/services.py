"""
Search, Data Quality, and Audit Services.
Multi-index and full-text search with source attribution, phone matching,
and quality resolution workflows.
"""

from typing import Dict, List, Any, Optional, Tuple
from django.db.models import Q
from .models import MCACompany, MCADirectorEntry, GSTRecord, QualityIssue, AuditEvent
from .cleaners import clean_mobile


def perform_unified_search(query: str, dataset_filter: Optional[str] = None, limit: int = 50) -> Dict[str, Any]:
    """
    Unified search across MCA Companies, Directors, and GST Registrations.
    Combines exact phone/identifier matches with text search,
    labeling results with strict dataset and lineage attribution.
    """
    clean_q = query.strip()
    if not clean_q:
        return {'mca_companies': [], 'mca_directors': [], 'gst_records': [], 'total_count': 0}
        
    canonical_phone, phone_status, _ = clean_mobile(clean_q, strict_prefix=False)
    
    results = {
        'mca_companies': [],
        'mca_directors': [],
        'gst_records': [],
        'total_count': 0
    }
    
    # 1. Search MCA Companies
    if not dataset_filter or dataset_filter == 'MCA':
        mca_comp_q = (
            Q(entity_id__iexact=clean_q) |
            Q(name__icontains=clean_q) |
            Q(email__icontains=clean_q) |
            Q(city__icontains=clean_q) |
            Q(state__icontains=clean_q)
        )
        companies = MCACompany.objects.filter(is_archived=False).filter(mca_comp_q)[:limit]
        results['mca_companies'] = list(companies)
        
        # 2. Search MCA Directors (including phone search)
        mca_dir_q = (
            Q(director_name__icontains=clean_q) |
            Q(director_email__icontains=clean_q) |
            Q(raw_mobile__icontains=clean_q)
        )
        if canonical_phone:
            mca_dir_q |= Q(canonical_mobile=canonical_phone)
            
        directors = MCADirectorEntry.objects.filter(mca_dir_q).select_related('company')[:limit]
        results['mca_directors'] = list(directors)
        
    # 3. Search GST Registrations
    if not dataset_filter or dataset_filter == 'GST':
        gst_q = (
            Q(gst_number__iexact=clean_q) |
            Q(pan__iexact=clean_q) |
            Q(trade_name__icontains=clean_q) |
            Q(legal_name__icontains=clean_q) |
            Q(email__icontains=clean_q) |
            Q(city__icontains=clean_q) |
            Q(state__icontains=clean_q) |
            Q(raw_mobile__icontains=clean_q)
        )
        if canonical_phone:
            gst_q |= Q(canonical_mobile=canonical_phone)
            
        gst_recs = GSTRecord.objects.filter(is_archived=False).filter(gst_q)[:limit]
        results['gst_records'] = list(gst_recs)
        
    results['total_count'] = len(results['mca_companies']) + len(results['mca_directors']) + len(results['gst_records'])
    return results


def resolve_quality_issue(issue_id: int, user, resolution_note: str = "", corrected_phone: Optional[str] = None) -> Tuple[bool, Optional[str]]:
    """
    Marks a QualityIssue as resolved, recording user and timestamp.
    If corrected_phone is provided, validates the corrected phone number
    strictly using clean_mobile before updating the record.
    Returns (success_boolean, error_message_or_None).
    """
    issue = QualityIssue.objects.filter(id=issue_id).first()
    if not issue:
        return False, "Issue could not be found."
        
    changes = {'note': resolution_note}

    if corrected_phone is not None and corrected_phone.strip():
        # Validate corrected number again before saving
        canon, status, err = clean_mobile(corrected_phone, strict_prefix=True)
        if status not in ('valid', 'cleaned') or not canon:
            return False, f"Invalid corrected phone number: {err or 'Must be 10 digits starting with 6-9'}"
            
        # Update target record
        if issue.record_type == 'DIRECTOR' and issue.record_id:
            dir_rec = MCADirectorEntry.objects.filter(id=issue.record_id).first()
            if dir_rec:
                changes['old_phone'] = dir_rec.raw_mobile
                changes['new_phone'] = canon
                dir_rec.raw_mobile = corrected_phone.strip()
                dir_rec.canonical_mobile = canon
                dir_rec.contact_validity = status
                dir_rec.save()
        elif issue.record_type == 'GST' and issue.record_id:
            gst_rec = GSTRecord.objects.filter(id=issue.record_id).first()
            if gst_rec:
                changes['old_phone'] = gst_rec.raw_mobile
                changes['new_phone'] = canon
                gst_rec.raw_mobile = corrected_phone.strip()
                gst_rec.canonical_mobile = canon
                gst_rec.contact_validity = status
                gst_rec.save()

    issue.is_resolved = True
    issue.resolved_by = user if user.is_authenticated else None
    issue.resolution_note = resolution_note
    from django.utils import timezone
    issue.resolved_at = timezone.now()
    issue.save()
    
    AuditEvent.objects.create(
        user=user if user.is_authenticated else None,
        action='RESOLVE_ISSUE',
        dataset_type=issue.dataset_type,
        model_name='QualityIssue',
        object_id=str(issue.id),
        object_repr=f"Resolved issue on {issue.entity_identifier} ({issue.error_code})",
        changes=changes
    )
    return True, None
