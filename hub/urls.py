from django.urls import path
from . import views

app_name = 'hub'

urlpatterns = [
    # Dashboard & Auth
    path('', views.dashboard_view, name='dashboard'),
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),

    # Universal Import Pipeline
    path('import/', views.import_upload_view, name='import_upload'),
    path('import/<str:batch_id>/map/', views.import_map_view, name='import_map'),
    path('import/<str:batch_id>/preview/', views.import_preview_view, name='import_preview'),
    path('import/<str:batch_id>/reclean/', views.import_reclean_view, name='import_reclean'),
    path('import/<str:batch_id>/reset/', views.import_reset_view, name='import_reset'),
    path('import/<str:batch_id>/execute/', views.import_execute_view, name='import_execute'),
    path('import/<str:batch_id>/detail/', views.import_batch_detail_view, name='import_batch_detail'),
    path('import/<str:batch_id>/report/', views.download_cleaning_report_view, name='download_cleaning_report'),
    path('import/<str:batch_id>/cleaned-file/', views.download_cleaned_file_view, name='download_cleaned_file'),
    path('import/history/', views.import_history_view, name='import_history'),

    # MCA Module
    path('mca/', views.mca_list_view, name='mca_list'),
    path('mca/dashboard/', views.mca_dashboard_view, name='mca_dashboard'),
    path('mca/directors/', views.mca_directors_view, name='mca_directors'),
    path('mca/upload/', views.mca_upload_view, name='mca_upload'),
    path('mca/cleaning/', views.mca_cleaning_view, name='mca_cleaning'),
    path('mca/search/', views.mca_search_view, name='mca_search'),
    path('mca/reports/', views.mca_reports_view, name='mca_reports'),
    path('mca/settings/', views.mca_settings_view, name='mca_settings'),
    path('mca/create/', views.mca_create_company_view, name='mca_create_company'),
    path('mca/export/json/', views.mca_export_json_view, name='mca_export_json'),
    path('mca/export/excel/', views.mca_export_excel_view, name='mca_export_excel'),
    path('mca/import/json/', views.mca_import_json_view, name='mca_import_json'),
    path('mca/clear-all/', views.mca_clear_all_view, name='mca_clear_all'),
    path('mca/<int:company_id>/', views.mca_detail_view, name='mca_detail'),
    path('mca/<int:company_id>/edit/', views.mca_edit_view, name='mca_edit'),
    path('mca/<int:company_id>/delete/', views.mca_delete_view, name='mca_delete'),
    path('mca/<int:company_id>/directors/add/', views.mca_director_add_view, name='mca_director_add'),

    # GST Module
    path('gst/', views.gst_list_view, name='gst_list'),
    path('gst/upload/', views.gst_upload_view, name='gst_upload'),
    path('gst/create/', views.gst_create_view, name='gst_create'),
    path('gst/export/json/', views.gst_export_json_view, name='gst_export_json'),
    path('gst/export/excel/', views.gst_export_excel_view, name='gst_export_excel'),
    path('gst/export/csv/', views.gst_export_csv_view, name='gst_export_csv'),
    path('gst/import/json/', views.gst_import_json_view, name='gst_import_json'),
    path('gst/clear-all/', views.gst_clear_all_view, name='gst_clear_all'),
    path('gst/<int:gst_id>/', views.gst_detail_view, name='gst_detail'),
    path('gst/<int:gst_id>/edit/', views.gst_edit_view, name='gst_edit'),
    path('gst/<int:gst_id>/delete/', views.gst_delete_view, name='gst_delete'),

    # Global Search
    path('search/', views.global_search_view, name='global_search'),

    # Data Quality Inbox
    path('quality/', views.quality_inbox_view, name='quality_inbox'),
    path('quality/clear-all/', views.quality_clear_all_view, name='quality_clear_all'),
    path('quality/<int:issue_id>/resolve/', views.quality_resolve_view, name='quality_resolve'),

    # Exports
    path('exports/', views.export_list_view, name='export_list'),
    path('exports/<int:job_id>/download/', views.export_download_view, name='export_download'),

    # Audit & Settings
    path('audit/', views.audit_history_view, name='audit_history'),
    path('settings/', views.settings_view, name='settings'),
    path('settings/change-password/', views.admin_change_password_view, name='admin_change_password'),
]
