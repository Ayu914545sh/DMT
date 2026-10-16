import os

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'dmt_project.settings')

try:
    from celery import Celery  # type: ignore
    app = Celery('dmt_project')
    app.config_from_object('django.conf:settings', namespace='CELERY')
    app.autodiscover_tasks()

    @app.task(bind=True, ignore_result=True)
    def debug_task(self):
        print(f'Request: {self.request!r}')
except (ImportError, ModuleNotFoundError):
    app = None
