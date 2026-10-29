import os

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'dmt_project.settings')

try:
    from celery import Celery  # type: ignore
    app = Celery('dmt_project')
    app.config_from_object('django.conf:settings', namespace='CELERY')
    app.autodiscover_tasks()

    # Explicitly import application tasks so Celery workers always register them
    try:
        import hub.tasks  # noqa: F401
    except Exception as _err:
        pass

    @app.task(bind=True, ignore_result=True)
    def debug_task(self):
        print(f'Request: {self.request!r}')
except (ImportError, ModuleNotFoundError):
    app = None
