"""
Celery asynchronous background tasks for data ingestion and exports.
Graceful execution supporting both async workers and immediate fallback.
"""

import logging
from typing import Optional, Any
from django.contrib.auth.models import User
from .models import ImportBatch, ExportJob
from .importers import execute_import_publication
from .exporters import run_export_job

logger = logging.getLogger(__name__)

# Safe import with graceful fallback if celery is missing in IDE interpreter
try:
    from celery import shared_task  # type: ignore
except (ImportError, ModuleNotFoundError):
    def shared_task(*args: Any, **kwargs: Any) -> Any:
        def decorator(func: Any) -> Any:
            def wrapper(*a: Any, **kw: Any) -> Any:
                return func(None, *a, **kw)
            setattr(wrapper, 'delay', wrapper)
            return wrapper
        if len(args) == 1 and callable(args[0]):
            return decorator(args[0])
        return decorator


@shared_task(bind=True, max_retries=3)
def process_import_batch_task(self: Any, batch_id: int, user_id: Optional[int] = None) -> Any:
    """
    Background worker task to publish a staged import batch idempotently.
    """
    try:
        batch = ImportBatch.objects.get(id=batch_id)
        user = User.objects.filter(id=user_id).first() if user_id else None
        result = execute_import_publication(batch, user=user)
        return result
    except Exception as exc:
        logger.error(f"Error publishing batch {batch_id}: {exc}", exc_info=True)
        if hasattr(self, 'retry'):
            raise self.retry(exc=exc, countdown=10)
        raise exc


@shared_task(bind=True, max_retries=2)
def process_export_job_task(self: Any, job_id: int) -> Any:
    """
    Background worker task to generate a CSV or Excel export file.
    """
    try:
        job = ExportJob.objects.get(id=job_id)
        path = run_export_job(job)
        return {'status': 'COMPLETED', 'file_path': path}
    except Exception as exc:
        logger.error(f"Error processing export job {job_id}: {exc}", exc_info=True)
        if hasattr(self, 'retry'):
            raise self.retry(exc=exc, countdown=5)
        raise exc
