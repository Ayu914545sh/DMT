from django.core.management.base import BaseCommand
from django.contrib.auth.models import User, Group
from hub.models import SystemSetting, MappingTemplate
from hub.importers import MCA_FIELDS, GST_FIELDS


class Command(BaseCommand):
    help = "Seeds initial user accounts (Administrator, Data Manager, Viewer), settings, and mapping templates."

    def handle(self, *args, **options):
        # 1. Groups & Roles
        admin_group, _ = Group.objects.get_or_create(name='Administrator')
        dm_group, _ = Group.objects.get_or_create(name='Data Manager')
        viewer_group, _ = Group.objects.get_or_create(name='Viewer')

        # 2. Administrator (admin / admin123)
        if not User.objects.filter(username='admin').exists():
            admin_user = User.objects.create_superuser('admin', 'admin@knowletive.com', 'admin123')
            admin_user.groups.add(admin_group)
            self.stdout.write(self.style.SUCCESS("Created superuser 'admin' with password 'admin123'"))
        else:
            self.stdout.write("User 'admin' already exists.")

        # 3. Data Manager (datamanager / manager123)
        if not User.objects.filter(username='datamanager').exists():
            dm_user = User.objects.create_user('datamanager', 'datamanager@knowletive.com', 'manager123', is_staff=True)
            dm_user.groups.add(dm_group)
            self.stdout.write(self.style.SUCCESS("Created staff user 'datamanager' with password 'manager123'"))

        # 4. Viewer (viewer / viewer123)
        if not User.objects.filter(username='viewer').exists():
            v_user = User.objects.create_user('viewer', 'viewer@knowletive.com', 'viewer123', is_staff=False)
            v_user.groups.add(viewer_group)
            self.stdout.write(self.style.SUCCESS("Created viewer user 'viewer' with password 'viewer123'"))

        # 5. System Settings
        SystemSetting.objects.get_or_create(
            key='STRICT_PHONE_PREFIX',
            defaults={
                'value': 'true',
                'description': 'Require Indian domestic mobile numbers to begin with 6, 7, 8, or 9'
            }
        )

        # 6. Reusable Mapping Templates
        default_mca_mapping = {field: field for field in MCA_FIELDS.keys()}
        MappingTemplate.objects.get_or_create(
            name='MCA Standard Template',
            dataset_type='MCA',
            defaults={
                'mapping_config': default_mca_mapping,
                'date_format': '%m/%d/%Y',
                'delimiter': ',',
                'is_default': True
            }
        )

        default_gst_mapping = {field: field for field in GST_FIELDS.keys()}
        MappingTemplate.objects.get_or_create(
            name='GST Standard Template',
            dataset_type='GST',
            defaults={
                'mapping_config': default_gst_mapping,
                'date_format': '%Y-%m-%d',
                'delimiter': ',',
                'is_default': True
            }
        )

        self.stdout.write(self.style.SUCCESS("Initial data and templates seeded successfully."))
