"""
Django Admin configuration with company-director TabularInline,
django-import-export resources, quality issue tracking, and audit log inspection.
"""

from django.contrib import admin
from django.utils.html import format_html
from import_export import resources, fields
from import_export.admin import ImportExportModelAdmin

from .models import (
    MCACompany, MCADirectorEntry, GSTRecord,
    ImportBatch, StagedRow, QualityIssue,
    MappingTemplate, AuditEvent, ExportJob,
    BusinessLink, SystemSetting
)


# ==========================================
# django-import-export Resource definitions
# ==========================================

class MCACompanyResource(resources.ModelResource):
    class Meta:
        model = MCACompany
        fields = (
            'id', 'entity_id', 'entity_type', 'name', 'state', 'district', 'city',
            'roc', 'nic_code', 'class_of_company', 'date_of_incorporation',
            'paid_up_capital', 'authorized_capital', 'email', 'registered_address'
        )
        export_order = fields
        import_id_fields = ('entity_type', 'entity_id')


class GSTRecordResource(resources.ModelResource):
    class Meta:
        model = GSTRecord
        fields = (
            'id', 'gst_number', 'legal_name', 'trade_name', 'nature_of_business',
            'source_type', 'registration_date', 'gst_type', 'status', 'jurisdiction',
            'state_jurisdiction', 'pan', 'canonical_mobile', 'contact_validity',
            'email', 'address', 'pincode', 'state', 'city'
        )
        export_order = fields
        import_id_fields = ('gst_number',)


# ==========================================
# Inlines
# ==========================================

class MCADirectorEntryInline(admin.TabularInline):
    model = MCADirectorEntry
    extra = 0
    can_delete = True
    fields = ('director_name', 'canonical_mobile', 'raw_mobile', 'contact_validity', 'director_email', 'designation')
    readonly_fields = ('raw_mobile', 'source_lineage')
    show_change_link = True

    def get_queryset(self, request):
        return super().get_queryset(request).select_related('company')


# ==========================================
# ModelAdmins
# ==========================================

@admin.register(MCACompany)
class MCACompanyAdmin(ImportExportModelAdmin):
    resource_classes = [MCACompanyResource]
    inlines = [MCADirectorEntryInline]
    list_display = (
        'name', 'entity_id', 'entity_type_badge', 'state', 'roc',
        'date_of_incorporation', 'director_count_badge', 'is_archived', 'updated_at'
    )
    list_filter = ('entity_type', 'state', 'roc', 'class_of_company', 'is_archived')
    search_fields = ('name', 'entity_id', 'email', 'directors__director_name', 'directors__canonical_mobile')
    readonly_fields = ('created_at', 'updated_at', 'revision', 'extra_data')
    list_per_page = 25

    def get_queryset(self, request):
        return super().get_queryset(request).prefetch_related('directors')

    def entity_type_badge(self, obj):
        color = "#0d6efd" if obj.entity_type == 'company' else "#198754"
        return format_html(
            '<span style="background:{}; color:#fff; padding:3px 7px; border-radius:4px; font-weight:600; font-size:11px;">{}</span>',
            color, obj.entity_type.upper()
        )
    entity_type_badge.short_description = "Type"

    def director_count_badge(self, obj):
        count = obj.directors.count()
        return format_html(
            '<span style="background:#e9ecef; color:#212529; padding:2px 8px; border-radius:12px; font-weight:700;">{}</span>',
            count
        )
    director_count_badge.short_description = "People / Directors"


@admin.register(MCADirectorEntry)
class MCADirectorEntryAdmin(admin.ModelAdmin):
    list_display = ('director_name', 'company_link', 'canonical_mobile', 'validity_badge', 'director_email', 'designation')
    list_filter = ('contact_validity', 'designation', 'company__state')
    search_fields = ('director_name', 'canonical_mobile', 'raw_mobile', 'director_email', 'company__name', 'company__entity_id')
    readonly_fields = ('raw_mobile', 'source_lineage', 'created_at', 'updated_at')
    list_per_page = 30

    def get_queryset(self, request):
        return super().get_queryset(request).select_related('company')

    def company_link(self, obj):
        return format_html('<a href="/admin/hub/mcacompany/{}/change/">{}</a>', obj.company.id, obj.company.name)
    company_link.short_description = "Company"

    def validity_badge(self, obj):
        colors = {'VALID': '#198754', 'FLAGGED': '#dc3545', 'BLANK': '#6c757d'}
        return format_html(
            '<span style="background:{}; color:#fff; padding:2px 6px; border-radius:4px; font-size:11px;">{}</span>',
            colors.get(obj.contact_validity, '#6c757d'), obj.contact_validity
        )
    validity_badge.short_description = "Validity"


@admin.register(GSTRecord)
class GSTRecordAdmin(ImportExportModelAdmin):
    resource_classes = [GSTRecordResource]
    list_display = (
        'gst_number', 'display_name', 'status_badge', 'source_type',
        'canonical_mobile', 'validity_badge', 'state', 'city', 'registration_date'
    )
    list_filter = ('status', 'source_type', 'contact_validity', 'state')
    search_fields = ('gst_number', 'legal_name', 'trade_name', 'pan', 'canonical_mobile', 'email')
    readonly_fields = ('raw_mobile', 'source_lineage', 'created_at', 'updated_at', 'revision', 'extra_data')
    list_per_page = 25

    def display_name(self, obj):
        return obj.trade_name or obj.legal_name or "-"
    display_name.short_description = "Trade / Legal Name"

    def status_badge(self, obj):
        color = '#198754' if obj.status == 'ACT' else '#6c757d'
        return format_html(
            '<span style="background:{}; color:#fff; padding:2px 6px; border-radius:4px; font-size:11px;">{}</span>',
            color, obj.status or 'N/A'
        )
    status_badge.short_description = "Status"

    def validity_badge(self, obj):
        colors = {'VALID': '#198754', 'FLAGGED': '#dc3545', 'BLANK': '#6c757d'}
        return format_html(
            '<span style="background:{}; color:#fff; padding:2px 6px; border-radius:4px; font-size:11px;">{}</span>',
            colors.get(obj.contact_validity, '#6c757d'), obj.contact_validity
        )
    validity_badge.short_description = "Phone Validity"


@admin.register(ImportBatch)
class ImportBatchAdmin(admin.ModelAdmin):
    list_display = (
        'batch_id', 'dataset_type', 'file_name', 'file_format',
        'status_badge', 'total_input_rows', 'imported_records', 'created_at'
    )
    list_filter = ('dataset_type', 'status', 'file_format')
    search_fields = ('batch_id', 'file_name')
    readonly_fields = (
        'batch_id', 'file_checksum', 'total_input_rows', 'proposed_new_records',
        'proposed_updates', 'unchanged_records', 'missing_invalid_values',
        'quarantined_rows', 'imported_records', 'created_at', 'completed_at'
    )

    def status_badge(self, obj):
        colors = {
            'COMPLETED': '#198754', 'IMPORTING': '#0d6efd',
            'PREVIEWED': '#ffc107', 'FAILED': '#dc3545', 'UPLOADED': '#6c757d'
        }
        return format_html(
            '<span style="background:{}; color:#fff; padding:2px 7px; border-radius:4px; font-size:11px;">{}</span>',
            colors.get(obj.status, '#333'), obj.status
        )
    status_badge.short_description = "Status"


@admin.register(QualityIssue)
class QualityIssueAdmin(admin.ModelAdmin):
    list_display = (
        'entity_identifier', 'dataset_type', 'field_name', 'error_code',
        'raw_value', 'is_resolved_badge', 'created_at'
    )
    list_filter = ('dataset_type', 'is_resolved', 'error_code')
    search_fields = ('entity_identifier', 'field_name', 'raw_value', 'reason_description')
    readonly_fields = ('created_at',)

    def is_resolved_badge(self, obj):
        color = '#198754' if obj.is_resolved else '#dc3545'
        text = 'RESOLVED' if obj.is_resolved else 'OPEN'
        return format_html(
            '<span style="background:{}; color:#fff; padding:2px 6px; border-radius:4px; font-size:11px;">{}</span>',
            color, text
        )
    is_resolved_badge.short_description = "State"


@admin.register(AuditEvent)
class AuditEventAdmin(admin.ModelAdmin):
    list_display = ('timestamp', 'user', 'action', 'dataset_type', 'model_name', 'object_repr')
    list_filter = ('action', 'dataset_type', 'model_name')
    search_fields = ('object_repr', 'object_id', 'user__username')
    readonly_fields = ('timestamp', 'user', 'action', 'dataset_type', 'model_name', 'object_id', 'object_repr', 'changes', 'ip_address')

    def has_add_permission(self, request):
        return False

    def has_change_permission(self, request, obj=None):
        return False


@admin.register(MappingTemplate)
class MappingTemplateAdmin(admin.ModelAdmin):
    list_display = ('name', 'dataset_type', 'is_default', 'updated_at')
    list_filter = ('dataset_type', 'is_default')
    search_fields = ('name',)


@admin.register(ExportJob)
class ExportJobAdmin(admin.ModelAdmin):
    list_display = ('id', 'requester', 'dataset_type', 'file_format', 'row_count', 'status', 'created_at')
    list_filter = ('dataset_type', 'file_format', 'status')


@admin.register(BusinessLink)
class BusinessLinkAdmin(admin.ModelAdmin):
    list_display = ('company', 'gst_record', 'match_method', 'match_confidence', 'status', 'reviewed_by')
    list_filter = ('status', 'match_method')
    search_fields = ('company__name', 'company__entity_id', 'gst_record__gst_number')


@admin.register(SystemSetting)
class SystemSettingAdmin(admin.ModelAdmin):
    list_display = ('key', 'value', 'description', 'updated_at')
    search_fields = ('key', 'description')
