"""
Forms for dataset upload, column mapping, entity editing, and audited exports.
"""

from django import forms
from .models import (
    MCACompany, MCADirectorEntry, GSTRecord,
    ImportBatch, MappingTemplate, ExportJob
)
from .cleaners import clean_mobile


class DatasetUploadForm(forms.Form):
    DATASET_CHOICES = [
        ('MCA', 'MCA (Companies & LLPs)'),
        ('GST', 'GST (Tax Registrations)'),
    ]
    dataset_type = forms.ChoiceField(
        choices=DATASET_CHOICES,
        widget=forms.Select(attrs={'class': 'form-select', 'id': 'datasetTypeSelect'})
    )
    data_file = forms.FileField(
        widget=forms.FileInput(attrs={'class': 'form-control', 'accept': '.csv,.xlsx,.xls', 'id': 'dataFileField'}),
        help_text="Upload .csv, .xlsx, or .xls format files"
    )
    sheet_name = forms.CharField(
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'e.g. Processed Data (Leave blank for first sheet)'})
    )
    header_row = forms.IntegerField(
        required=False,
        initial=1,
        min_value=1,
        widget=forms.NumberInput(attrs={'class': 'form-control', 'placeholder': '1 for first row'})
    )
    csv_delimiter = forms.ChoiceField(
        required=False,
        choices=[(',', 'Comma ( , )'), ('\t', 'Tab ( \\t )'), (';', 'Semicolon ( ; )'), ('|', 'Pipe ( | )')],
        initial=',',
        widget=forms.Select(attrs={'class': 'form-select'})
    )
    csv_encoding = forms.ChoiceField(
        required=False,
        choices=[('utf-8', 'UTF-8'), ('utf-8-sig', 'UTF-8 with BOM'), ('latin-1', 'Latin-1 / CP1252')],
        initial='utf-8',
        widget=forms.Select(attrs={'class': 'form-select'})
    )

    def clean_data_file(self):
        file = self.cleaned_data.get('data_file')
        if file:
            ext = file.name.split('.')[-1].lower()
            if ext not in ('csv', 'xlsx', 'xls'):
                raise forms.ValidationError("Unsupported file format. Please upload a .csv, .xlsx, or .xls file.")
        return file


class MCACompanyForm(forms.ModelForm):
    class Meta:
        model = MCACompany
        fields = [
            'name', 'entity_type', 'entity_id', 'state', 'district', 'city',
            'roc', 'nic_code', 'nic_label', 'class_of_company',
            'date_of_incorporation', 'paid_up_capital', 'authorized_capital',
            'email', 'registered_address', 'is_archived'
        ]
        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'entity_type': forms.Select(attrs={'class': 'form-select'}),
            'entity_id': forms.TextInput(attrs={'class': 'form-control'}),
            'state': forms.TextInput(attrs={'class': 'form-control'}),
            'district': forms.TextInput(attrs={'class': 'form-control'}),
            'city': forms.TextInput(attrs={'class': 'form-control'}),
            'roc': forms.TextInput(attrs={'class': 'form-control'}),
            'nic_code': forms.TextInput(attrs={'class': 'form-control'}),
            'nic_label': forms.TextInput(attrs={'class': 'form-control'}),
            'class_of_company': forms.TextInput(attrs={'class': 'form-control'}),
            'date_of_incorporation': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'paid_up_capital': forms.NumberInput(attrs={'class': 'form-control'}),
            'authorized_capital': forms.NumberInput(attrs={'class': 'form-control'}),
            'email': forms.EmailInput(attrs={'class': 'form-control'}),
            'registered_address': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
            'is_archived': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
        }


class MCADirectorForm(forms.ModelForm):
    class Meta:
        model = MCADirectorEntry
        fields = ['director_name', 'raw_mobile', 'director_email', 'designation']
        widgets = {
            'director_name': forms.TextInput(attrs={'class': 'form-control'}),
            'raw_mobile': forms.TextInput(attrs={'class': 'form-control'}),
            'director_email': forms.EmailInput(attrs={'class': 'form-control'}),
            'designation': forms.TextInput(attrs={'class': 'form-control'}),
        }

    def clean(self):
        cleaned_data = super().clean()
        raw_m = cleaned_data.get('raw_mobile')
        # Validate and clean phone
        canon, status, err = clean_mobile(raw_m, strict_prefix=True)
        self.instance.raw_mobile = raw_m or ""
        self.instance.canonical_mobile = canon
        self.instance.contact_validity = status
        return cleaned_data


class GSTRecordForm(forms.ModelForm):
    class Meta:
        model = GSTRecord
        fields = [
            'gst_number', 'legal_name', 'trade_name', 'nature_of_business',
            'source_type', 'registration_date', 'gst_type', 'status',
            'jurisdiction', 'state_jurisdiction', 'pan', 'raw_mobile',
            'email', 'address', 'pincode', 'state', 'city', 'is_archived'
        ]
        widgets = {
            'gst_number': forms.TextInput(attrs={'class': 'form-control'}),
            'legal_name': forms.TextInput(attrs={'class': 'form-control'}),
            'trade_name': forms.TextInput(attrs={'class': 'form-control'}),
            'nature_of_business': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
            'source_type': forms.TextInput(attrs={'class': 'form-control'}),
            'registration_date': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'gst_type': forms.TextInput(attrs={'class': 'form-control'}),
            'status': forms.TextInput(attrs={'class': 'form-control'}),
            'jurisdiction': forms.TextInput(attrs={'class': 'form-control'}),
            'state_jurisdiction': forms.TextInput(attrs={'class': 'form-control'}),
            'pan': forms.TextInput(attrs={'class': 'form-control'}),
            'raw_mobile': forms.TextInput(attrs={'class': 'form-control'}),
            'email': forms.EmailInput(attrs={'class': 'form-control'}),
            'address': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
            'pincode': forms.TextInput(attrs={'class': 'form-control'}),
            'state': forms.TextInput(attrs={'class': 'form-control'}),
            'city': forms.TextInput(attrs={'class': 'form-control'}),
            'is_archived': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
        }

    def clean(self):
        cleaned_data = super().clean()
        raw_m = cleaned_data.get('raw_mobile')
        canon, status, err = clean_mobile(raw_m, strict_prefix=True)
        self.instance.raw_mobile = raw_m or ""
        self.instance.canonical_mobile = canon
        self.instance.contact_validity = status
        return cleaned_data


class ExportRequestForm(forms.Form):
    DATASET_CHOICES = [
        ('MCA', 'MCA Companies & Directors'),
        ('GST', 'GST Registrations'),
    ]
    FORMAT_CHOICES = [
        ('CSV', 'CSV Format'),
        ('XLSX', 'Excel (XLSX) Format'),
    ]
    dataset_type = forms.ChoiceField(choices=DATASET_CHOICES, widget=forms.Select(attrs={'class': 'form-select'}))
    file_format = forms.ChoiceField(choices=FORMAT_CHOICES, widget=forms.Select(attrs={'class': 'form-select'}))
    state_filter = forms.CharField(required=False, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Optional State filter'}))
    status_filter = forms.CharField(required=False, widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Optional Status / Type filter'}))


class AdminPasswordChangeForm(forms.Form):
    current_password = forms.CharField(
        widget=forms.PasswordInput(attrs={'class': 'form-control', 'placeholder': 'Enter current password', 'autocomplete': 'current-password'}),
        label="Current Password"
    )
    new_password = forms.CharField(
        widget=forms.PasswordInput(attrs={'class': 'form-control', 'placeholder': 'Enter new password', 'autocomplete': 'new-password'}),
        label="New Password"
    )
    confirm_new_password = forms.CharField(
        widget=forms.PasswordInput(attrs={'class': 'form-control', 'placeholder': 'Confirm new password', 'autocomplete': 'new-password'}),
        label="Confirm New Password"
    )

    def __init__(self, user, *args, **kwargs):
        self.user = user
        super().__init__(*args, **kwargs)

    def clean_current_password(self):
        current_password = self.cleaned_data.get('current_password')
        if not self.user.check_password(current_password):
            raise forms.ValidationError("The current password entered is incorrect.")
        return current_password

    def clean(self):
        cleaned_data = super().clean()
        new_password = cleaned_data.get('new_password')
        confirm_new_password = cleaned_data.get('confirm_new_password')

        if new_password and confirm_new_password:
            if new_password != confirm_new_password:
                self.add_error('confirm_new_password', "The new password and confirmation password do not match.")

            # Strong password rule validation
            # - Minimum 8 characters
            # - At least one uppercase letter
            # - At least one lowercase letter
            # - At least one number
            # - At least one special character
            errors = []
            if len(new_password) < 8:
                errors.append("at least 8 characters")
            if not any(c.isupper() for c in new_password):
                errors.append("at least one uppercase letter (A-Z)")
            if not any(c.islower() for c in new_password):
                errors.append("at least one lowercase letter (a-z)")
            if not any(c.isdigit() for c in new_password):
                errors.append("at least one number (0-9)")
            import re
            if not re.search(r'[!@#$%^&*(),.?":{}|<>_~`\-+=\[\]\\;/]', new_password):
                errors.append("at least one special character (e.g. !@#$%^&*)")

            if errors:
                self.add_error('new_password', f"Password must contain {', '.join(errors)}.")

        return cleaned_data
