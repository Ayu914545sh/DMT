"""
Relational models for the Unified MCA & GST Data Management Webapp.
Strict schema definition with index optimization, constraints, and audit linkages.
"""

from django.db import models
from django.contrib.auth.models import User
from django.contrib.postgres.indexes import GinIndex
from django.utils import timezone
import uuid


class MCACompany(models.Model):
    """
    Authoritative record for an MCA entity (Company or LLP).
    Enforces uniqueness on normalized entity_type + entity_id.
    """
    ENTITY_TYPE_CHOICES = [
        ('company', 'Company'),
        ('llp', 'LLP'),
    ]

    id = models.BigAutoField(primary_key=True)
    entity_id = models.CharField(max_length=64, db_index=True, help_text="CIN or LLPIN (normalized uppercase)")
    entity_type = models.CharField(max_length=32, choices=ENTITY_TYPE_CHOICES, db_index=True)
    name = models.CharField(max_length=255, db_index=True)
    state = models.CharField(max_length=100, blank=True, db_index=True)
    district = models.CharField(max_length=100, blank=True)
    city = models.CharField(max_length=100, blank=True, db_index=True)
    roc = models.CharField(max_length=100, blank=True, db_index=True)
    nic_code = models.CharField(max_length=32, blank=True, db_index=True)
    nic_label = models.TextField(blank=True)
    class_of_company = models.CharField(max_length=255, blank=True, help_text="Private, Public, etc.")
    date_of_incorporation = models.DateField(null=True, blank=True, db_index=True)
    paid_up_capital = models.DecimalField(max_digits=18, decimal_places=2, null=True, blank=True)
    authorized_capital = models.DecimalField(max_digits=18, decimal_places=2, null=True, blank=True)
    email = models.CharField(max_length=255, blank=True, db_index=True)
    registered_address = models.TextField(blank=True)
    source_upload_date = models.DateTimeField(null=True, blank=True)
    
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_archived = models.BooleanField(default=False, db_index=True)
    revision = models.PositiveIntegerField(default=1)
    extra_data = models.JSONField(default=dict, blank=True, help_text="Preserved unmapped and supplementary source fields")

    class Meta:
        verbose_name = "MCA Company / LLP"
        verbose_name_plural = "MCA Companies & LLPs"
        constraints = [
            models.UniqueConstraint(
                fields=['entity_type', 'entity_id'],
                name='unique_mca_entity_type_id'
            )
        ]
        indexes = [
            models.Index(fields=['name']),
            models.Index(fields=['state', 'roc']),
            models.Index(fields=['date_of_incorporation']),
            GinIndex(fields=['name'], name='mca_comp_name_trgm_gin', opclasses=['gin_trgm_ops']),
            GinIndex(fields=['entity_id'], name='mca_comp_eid_trgm_gin', opclasses=['gin_trgm_ops']),
            GinIndex(fields=['city'], name='mca_comp_city_trgm_gin', opclasses=['gin_trgm_ops']),
            GinIndex(fields=['state'], name='mca_comp_state_trgm_gin', opclasses=['gin_trgm_ops']),
            GinIndex(fields=['roc'], name='mca_comp_roc_trgm_gin', opclasses=['gin_trgm_ops']),
            GinIndex(fields=['email'], name='mca_comp_email_trgm_gin', opclasses=['gin_trgm_ops']),
        ]
        ordering = ['-updated_at', 'name']

    def __str__(self):
        return f"{self.name} ({self.entity_id})"

    @property
    def director_count(self):
        return self.directors.count()


class MCADirectorEntry(models.Model):
    """
    Child entry linked to MCACompany.
    Preserves multiple director / person details per entity without deduplicating companies.
    """
    CONTACT_VALIDITY_CHOICES = [
        ('valid', 'Valid 10-Digit Mobile'),
        ('cleaned', 'Cleaned & Normalized'),
        ('missing', 'Missing / Blank'),
        ('invalid', 'Invalid Format / Length / Prefix'),
        ('multiple_numbers_review', 'Multiple Numbers (Manual Review)'),
        # Backward-compatible legacy aliases
        ('VALID', 'Valid 10-Digit Mobile'),
        ('FLAGGED', 'Flagged / Invalid Format'),
        ('BLANK', 'Missing / Blank'),
    ]

    id = models.BigAutoField(primary_key=True)
    company = models.ForeignKey(
        MCACompany,
        on_delete=models.CASCADE,
        related_name='directors',
        help_text="Parent MCA entity"
    )
    director_name = models.CharField(max_length=255, db_index=True)
    director_email = models.CharField(max_length=255, blank=True, db_index=True)
    raw_mobile = models.CharField(max_length=255, blank=True, help_text="Exact source phone cell value (original_phone)")
    canonical_mobile = models.CharField(
        max_length=10,
        null=True,
        blank=True,
        db_index=True,
        help_text="Exactly 10 ASCII digits or NULL (cleaned_phone)"
    )
    contact_validity = models.CharField(
        max_length=32,
        choices=CONTACT_VALIDITY_CHOICES,
        default='missing',
        db_index=True,
        help_text="phone_status: valid, cleaned, missing, invalid, multiple_numbers_review"
    )
    designation = models.CharField(
        max_length=255,
        default='Unknown',
        blank=True,
        help_text="Director, Designated Partner, or Unknown"
    )
    source_lineage = models.JSONField(
        default=dict,
        blank=True,
        help_text="Source trace: batch, file, sheet, row number, raw fields"
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    # Required field aliases: original_phone, cleaned_phone, phone_status
    @property
    def original_phone(self):
        return self.raw_mobile

    @original_phone.setter
    def original_phone(self, val):
        self.raw_mobile = str(val) if val is not None else ""

    @property
    def cleaned_phone(self):
        return self.canonical_mobile

    @cleaned_phone.setter
    def cleaned_phone(self, val):
        self.canonical_mobile = val

    @property
    def phone_status(self):
        return self.contact_validity

    @phone_status.setter
    def phone_status(self, val):
        self.contact_validity = val

    class Meta:
        verbose_name = "MCA Director / Person"
        verbose_name_plural = "MCA Directors & Persons"
        indexes = [
            models.Index(fields=['director_name']),
            models.Index(fields=['canonical_mobile']),
            models.Index(fields=['contact_validity']),
            GinIndex(fields=['director_name'], name='mca_dir_name_trgm_gin', opclasses=['gin_trgm_ops']),
            GinIndex(fields=['canonical_mobile'], name='mca_dir_phone_trgm_gin', opclasses=['gin_trgm_ops']),
            GinIndex(fields=['director_email'], name='mca_dir_email_trgm_gin', opclasses=['gin_trgm_ops']),
        ]
        ordering = ['director_name']

    def __str__(self):
        return f"{self.director_name} ({self.company.name})"


class GSTRecord(models.Model):
    """
    Authoritative record for GST registrations.
    Maintained separately from MCA data.
    """
    CONTACT_VALIDITY_CHOICES = [
        ('valid', 'Valid 10-Digit Mobile'),
        ('cleaned', 'Cleaned & Normalized'),
        ('missing', 'Missing / Blank'),
        ('invalid', 'Invalid Format / Length / Prefix'),
        ('multiple_numbers_review', 'Multiple Numbers (Manual Review)'),
        # Backward-compatible legacy aliases
        ('VALID', 'Valid 10-Digit Mobile'),
        ('FLAGGED', 'Flagged / Invalid Format'),
        ('BLANK', 'Missing / Blank'),
    ]

    id = models.BigAutoField(primary_key=True)
    gst_number = models.CharField(
        max_length=32,
        unique=True,
        db_index=True,
        help_text="Normalized uppercase GSTIN"
    )
    legal_name = models.CharField(max_length=255, blank=True, db_index=True)
    trade_name = models.CharField(max_length=255, blank=True, db_index=True)
    nature_of_business = models.TextField(blank=True)
    source_type = models.CharField(
        max_length=255,
        blank=True,
        db_index=True,
        help_text="Source 'type' column (e.g. REG)"
    )
    registration_date = models.DateField(null=True, blank=True, db_index=True)
    gst_type = models.CharField(max_length=255, blank=True, db_index=True)
    status = models.CharField(max_length=255, blank=True, db_index=True, help_text="Source status code (e.g. ACT)")
    jurisdiction = models.CharField(max_length=255, blank=True)
    state_jurisdiction = models.CharField(max_length=255, blank=True)
    pan = models.CharField(max_length=32, blank=True, db_index=True)
    raw_mobile = models.CharField(max_length=255, blank=True, help_text="Exact source phone cell value (original_phone)")
    canonical_mobile = models.CharField(
        max_length=10,
        null=True,
        blank=True,
        db_index=True,
        help_text="Exactly 10 ASCII digits or NULL (cleaned_phone)"
    )
    contact_validity = models.CharField(
        max_length=32,
        choices=CONTACT_VALIDITY_CHOICES,
        default='missing',
        db_index=True,
        help_text="phone_status: valid, cleaned, missing, invalid, multiple_numbers_review"
    )
    email = models.CharField(max_length=255, blank=True, db_index=True)
    address = models.TextField(blank=True)
    pincode = models.CharField(max_length=32, blank=True, db_index=True)
    state = models.CharField(max_length=100, blank=True, db_index=True)
    city = models.CharField(max_length=100, blank=True, db_index=True)
    source_insert_time = models.DateTimeField(null=True, blank=True)
    source_lineage = models.JSONField(default=dict, blank=True)
    
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    # Required field aliases: original_phone, cleaned_phone, phone_status
    @property
    def original_phone(self):
        return self.raw_mobile

    @original_phone.setter
    def original_phone(self, val):
        self.raw_mobile = str(val) if val is not None else ""

    @property
    def cleaned_phone(self):
        return self.canonical_mobile

    @cleaned_phone.setter
    def cleaned_phone(self, val):
        self.canonical_mobile = val

    @property
    def phone_status(self):
        return self.contact_validity

    @phone_status.setter
    def phone_status(self, val):
        self.contact_validity = val

    @property
    def gstin(self):
        return self.gst_number

    @property
    def mobile_number(self):
        return self.canonical_mobile

    is_archived = models.BooleanField(default=False, db_index=True)
    revision = models.PositiveIntegerField(default=1)
    extra_data = models.JSONField(default=dict, blank=True)

    class Meta:
        verbose_name = "GST Registration"
        verbose_name_plural = "GST Registrations"
        indexes = [
            models.Index(fields=['legal_name']),
            models.Index(fields=['trade_name']),
            models.Index(fields=['pan']),
            models.Index(fields=['canonical_mobile']),
            models.Index(fields=['status']),
            models.Index(fields=['state', 'city']),
            GinIndex(fields=['trade_name'], name='gst_trade_name_trgm_gin', opclasses=['gin_trgm_ops']),
            GinIndex(fields=['legal_name'], name='gst_legal_name_trgm_gin', opclasses=['gin_trgm_ops']),
            GinIndex(fields=['gst_number'], name='gst_number_trgm_gin', opclasses=['gin_trgm_ops']),
            GinIndex(fields=['pan'], name='gst_pan_trgm_gin', opclasses=['gin_trgm_ops']),
            GinIndex(fields=['canonical_mobile'], name='gst_phone_trgm_gin', opclasses=['gin_trgm_ops']),
            GinIndex(fields=['city'], name='gst_city_trgm_gin', opclasses=['gin_trgm_ops']),
            GinIndex(fields=['state'], name='gst_state_trgm_gin', opclasses=['gin_trgm_ops']),
            GinIndex(fields=['address'], name='gst_addr_trgm_gin', opclasses=['gin_trgm_ops']),
            GinIndex(fields=['email'], name='gst_email_trgm_gin', opclasses=['gin_trgm_ops']),
        ]
        ordering = ['-updated_at', 'gst_number']

    def __str__(self):
        disp = self.trade_name or self.legal_name or self.gst_number
        return f"{disp} ({self.gst_number})"


class ImportBatch(models.Model):
    """
    Manages an upload session, checksum verification, mapping configuration,
    and row-level progress tracking.
    """
    DATASET_CHOICES = [
        ('MCA', 'MCA (Companies & LLPs)'),
        ('GST', 'GST (Tax Registrations)'),
    ]
    STATUS_CHOICES = [
        ('UPLOADED', 'Uploaded'),
        ('MAPPED', 'Mapped'),
        ('QUEUED', 'Queued for processing'),
        ('PROCESSING', 'Processing'),
        ('STAGED', 'Staged & Validated'),
        ('PREVIEWED', 'Preview Ready'),
        ('IMPORTING', 'Importing'),
        ('COMPLETED', 'Completed'),
        ('PARTIAL', 'Completed with Issues'),
        ('FAILED', 'Failed'),
        ('CANCELLED', 'Cancelled'),
    ]

    id = models.BigAutoField(primary_key=True)
    batch_id = models.CharField(max_length=64, unique=True, db_index=True)
    dataset_type = models.CharField(max_length=16, choices=DATASET_CHOICES, db_index=True)
    file_name = models.CharField(max_length=255)
    file_path = models.CharField(max_length=512)
    file_checksum = models.CharField(max_length=64, help_text="SHA-256 hash")
    file_format = models.CharField(max_length=16, choices=[('CSV', 'CSV'), ('XLSX', 'Excel XLSX')])
    sheet_name = models.CharField(max_length=128, blank=True)
    header_row_index = models.PositiveIntegerField(default=0)
    status = models.CharField(max_length=32, choices=STATUS_CHOICES, default='UPLOADED', db_index=True)
    
    total_input_rows = models.PositiveIntegerField(default=0)
    proposed_new_records = models.PositiveIntegerField(default=0)
    proposed_updates = models.PositiveIntegerField(default=0)
    unchanged_records = models.PositiveIntegerField(default=0)
    missing_invalid_values = models.PositiveIntegerField(default=0)
    conflict_count = models.PositiveIntegerField(default=0)
    quarantined_rows = models.PositiveIntegerField(default=0)
    imported_records = models.PositiveIntegerField(default=0)
    
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    mapping_config = models.JSONField(default=dict, blank=True)
    cleaning_summary = models.JSONField(default=dict, blank=True)
    error_message = models.TextField(blank=True)
    checkpoint_row = models.PositiveIntegerField(default=0, help_text="Durable checkpoint row offset for publication crash recovery")
    retry_count = models.PositiveIntegerField(default=0)
    
    created_at = models.DateTimeField(auto_now_add=True)
    completed_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        verbose_name = "Import Batch"
        verbose_name_plural = "Import Batches"
        ordering = ['-created_at']

    def __str__(self):
        return f"{self.batch_id} [{self.dataset_type}] - {self.status}"

    @property
    def progress_percentage(self):
        if self.total_input_rows == 0:
            return 100 if self.status == 'COMPLETED' else 0
        return min(100, int((self.imported_records / self.total_input_rows) * 100))


class StagedRow(models.Model):
    """
    Preserves raw and cleaned data for each row during the import staging lifecycle.
    Never silently discards rows. Includes indexed entity and phone columns for
    fast, database-side deduplication and zero-RAM join lookups in PostgreSQL.
    """
    ACTION_CHOICES = [
        ('CREATE', 'Create New'),
        ('UPDATE', 'Update Existing'),
        ('UNCHANGED', 'Unchanged'),
        ('QUARANTINE', 'Quarantined'),
        ('ERROR', 'Error'),
    ]
    STATUS_CHOICES = [
        ('VALID', 'Valid'),
        ('WARNING', 'Warning / Flagged Phone'),
        ('ERROR', 'Error'),
        ('QUARANTINE', 'Quarantine'),
    ]

    id = models.BigAutoField(primary_key=True)
    batch = models.ForeignKey(ImportBatch, on_delete=models.CASCADE, related_name='staged_rows')
    row_number = models.PositiveIntegerField(db_index=True)
    entity_identifier = models.CharField(max_length=64, blank=True, db_index=True, help_text="Indexed entity ID (CIN, LLPIN, GSTIN)")
    canonical_mobile = models.CharField(max_length=10, blank=True, db_index=True, help_text="Indexed 10-digit phone for DB grouping")
    raw_data = models.JSONField(default=dict)
    cleaned_data = models.JSONField(default=dict)
    proposed_action = models.CharField(max_length=32, choices=ACTION_CHOICES, default='CREATE')
    validation_status = models.CharField(max_length=32, choices=STATUS_CHOICES, default='VALID')
    validation_messages = models.JSONField(default=list, blank=True)
    diff_summary = models.JSONField(default=dict, blank=True)
    is_processed = models.BooleanField(default=False, db_index=True)
    processed_record_id = models.CharField(max_length=128, blank=True)

    class Meta:
        verbose_name = "Staged Row"
        verbose_name_plural = "Staged Rows"
        ordering = ['batch', 'row_number']
        constraints = [
            models.UniqueConstraint(fields=['batch', 'row_number'], name='unique_staged_row_per_batch')
        ]
        indexes = [
            models.Index(fields=['batch', 'row_number']),
            models.Index(fields=['batch', 'is_processed']),
            models.Index(fields=['batch', 'entity_identifier']),
            models.Index(fields=['batch', 'canonical_mobile']),
            models.Index(fields=['validation_status']),
        ]

    def __str__(self):
        return f"Batch {self.batch.batch_id} Row #{self.row_number} ({self.proposed_action})"


class QualityIssue(models.Model):
    """
    Records data quality defects (e.g. invalid phone format, unparsable date)
    with reason codes for review in the Data Quality Inbox.
    """
    DATASET_CHOICES = [
        ('MCA', 'MCA'),
        ('GST', 'GST'),
    ]
    RECORD_TYPE_CHOICES = [
        ('COMPANY', 'MCA Company'),
        ('DIRECTOR', 'MCA Director / Person'),
        ('GST', 'GST Registration'),
    ]

    id = models.BigAutoField(primary_key=True)
    dataset_type = models.CharField(max_length=16, choices=DATASET_CHOICES, db_index=True)
    batch = models.ForeignKey(ImportBatch, on_delete=models.SET_NULL, null=True, blank=True, related_name='quality_issues')
    record_id = models.BigIntegerField(null=True, blank=True, help_text="ID of created record if saved")
    record_type = models.CharField(max_length=32, choices=RECORD_TYPE_CHOICES, db_index=True)
    entity_identifier = models.CharField(max_length=255, db_index=True, help_text="CIN, LLPIN, or GSTIN")
    field_name = models.CharField(max_length=128, db_index=True)
    raw_value = models.TextField(blank=True)
    error_code = models.CharField(max_length=128, db_index=True)
    reason_description = models.TextField()
    is_resolved = models.BooleanField(default=False, db_index=True)
    resolved_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='resolved_issues')
    resolution_note = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    resolved_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        verbose_name = "Quality Issue"
        verbose_name_plural = "Quality Issues"
        indexes = [
            models.Index(fields=['dataset_type', 'is_resolved']),
            models.Index(fields=['error_code']),
        ]
        ordering = ['-created_at']

    def __str__(self):
        return f"[{self.dataset_type}] {self.entity_identifier} - {self.error_code}"


class MappingTemplate(models.Model):
    """
    Reusable header mapping templates for recurring CSV/XLSX imports.
    """
    DATASET_CHOICES = [
        ('MCA', 'MCA (Companies & LLPs)'),
        ('GST', 'GST (Tax Registrations)'),
    ]

    id = models.BigAutoField(primary_key=True)
    name = models.CharField(max_length=128)
    dataset_type = models.CharField(max_length=16, choices=DATASET_CHOICES, db_index=True)
    mapping_config = models.JSONField(help_text="Dictionary of source_header -> model_field")
    date_format = models.CharField(max_length=32, blank=True, default="%m/%d/%Y")
    delimiter = models.CharField(max_length=8, blank=True, default=",")
    is_default = models.BooleanField(default=False)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Mapping Template"
        verbose_name_plural = "Mapping Templates"
        unique_together = [['name', 'dataset_type']]
        ordering = ['name']

    def __str__(self):
        return f"{self.name} ({self.dataset_type})"


class AuditEvent(models.Model):
    """
    Comprehensive audit log tracing user actions, changes, and exports.
    """
    ACTION_CHOICES = [
        ('IMPORT', 'Batch Import'),
        ('CREATE', 'Create Record'),
        ('UPDATE', 'Edit Record'),
        ('DELETE', 'Delete Record'),
        ('ARCHIVE', 'Archive Record'),
        ('RESTORE', 'Restore Record'),
        ('EXPORT', 'Data Export'),
        ('RESOLVE_ISSUE', 'Resolved Quality Issue'),
        ('MAPPING_CHANGE', 'Mapping Configuration Changed'),
    ]

    id = models.BigAutoField(primary_key=True)
    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    action = models.CharField(max_length=32, choices=ACTION_CHOICES, db_index=True)
    dataset_type = models.CharField(max_length=16, blank=True, db_index=True)
    model_name = models.CharField(max_length=64, db_index=True)
    object_id = models.CharField(max_length=128, blank=True)
    object_repr = models.CharField(max_length=255, blank=True)
    changes = models.JSONField(default=dict, blank=True)
    ip_address = models.GenericIPAddressField(null=True, blank=True)
    timestamp = models.DateTimeField(auto_now_add=True, db_index=True)

    class Meta:
        verbose_name = "Audit Event"
        verbose_name_plural = "Audit Events"
        ordering = ['-timestamp']

    def __str__(self):
        user_str = self.user.username if self.user else "System"
        return f"{self.timestamp.strftime('%Y-%m-%d %H:%M')} | {user_str} | {self.action} on {self.model_name}"


class ExportJob(models.Model):
    """
    Records and tracks authorized CSV/XLSX export requests.
    """
    STATUS_CHOICES = [
        ('PENDING', 'Pending'),
        ('PROCESSING', 'Processing'),
        ('COMPLETED', 'Completed'),
        ('FAILED', 'Failed'),
    ]

    id = models.BigAutoField(primary_key=True)
    requester = models.ForeignKey(User, on_delete=models.CASCADE)
    dataset_type = models.CharField(max_length=16, choices=[('MCA', 'MCA'), ('GST', 'GST')])
    file_format = models.CharField(max_length=8, choices=[('CSV', 'CSV'), ('XLSX', 'Excel XLSX')])
    columns = models.JSONField(default=list)
    filters = models.JSONField(default=dict, blank=True)
    row_count = models.PositiveIntegerField(default=0)
    file_path = models.CharField(max_length=512, blank=True)
    status = models.CharField(max_length=16, choices=STATUS_CHOICES, default='PENDING')
    created_at = models.DateTimeField(auto_now_add=True)
    completed_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        verbose_name = "Export Job"
        verbose_name_plural = "Export Jobs"
        ordering = ['-created_at']

    def __str__(self):
        return f"Export {self.id} [{self.dataset_type} {self.file_format}] - {self.requester.username}"


class BusinessLink(models.Model):
    """
    Reserved structure for reviewed connections between MCA Companies and GST Registrations.
    Never joins automatically; retains audit evidence and reviewer sign-off.
    """
    STATUS_CHOICES = [
        ('PROPOSED', 'Proposed Candidate'),
        ('REVIEWED', 'Reviewed'),
        ('APPROVED', 'Approved Link'),
        ('REJECTED', 'Rejected'),
    ]

    id = models.BigAutoField(primary_key=True)
    company = models.ForeignKey(MCACompany, on_delete=models.CASCADE, related_name='gst_links')
    gst_record = models.ForeignKey(GSTRecord, on_delete=models.CASCADE, related_name='mca_links')
    match_method = models.CharField(max_length=64, default='MANUAL')
    match_confidence = models.DecimalField(max_digits=5, decimal_places=2, default=1.0)
    status = models.CharField(max_length=32, choices=STATUS_CHOICES, default='PROPOSED')
    evidence = models.JSONField(default=dict, blank=True)
    reviewed_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    reviewed_at = models.DateTimeField(null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "MCA - GST Business Link"
        verbose_name_plural = "MCA - GST Business Links"
        unique_together = [['company', 'gst_record']]
        ordering = ['-created_at']

    def __str__(self):
        return f"{self.company.name} <-> {self.gst_record.gst_number} ({self.status})"


class SystemSetting(models.Model):
    """
    Dynamic system policies: strict phone validation, auto-quarantine, etc.
    """
    key = models.CharField(max_length=64, unique=True, db_index=True)
    value = models.CharField(max_length=255)
    description = models.TextField(blank=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "System Setting"
        verbose_name_plural = "System Settings"

    def __str__(self):
        return f"{self.key} = {self.value}"

    @classmethod
    def get_bool(cls, key: str, default: bool = True) -> bool:
        try:
            setting = cls.objects.filter(key=key).first()
            if setting:
                return setting.value.lower() in ('true', '1', 'yes')
        except Exception:
            pass
        return default
