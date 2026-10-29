"""
Celery asynchronous background tasks for data ingestion, staging, publication, and exports.
Graceful execution supporting both async distributed workers and synchronous fallback.
"""

import logging
from typing import Optional, Any, Dict
from django.contrib.auth.models import User
from django.conf import settings
from .models import ImportBatch, ExportJob
from .importers import stage_and_validate_batch, execute_import_publication
from .exporters import run_export_job
from .progress import update_progress, set_progress_completed, set_progress_failed, is_cancelled

logger = logging.getLogger(__name__)

# Safe import with graceful fallback if celery is missing in environment
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
def stage_and_validate_batch_task(self: Any, batch_id: str, strict_phone: bool = True) -> Dict[str, Any]:
    """
    Celery task to run memory-bounded staging and data cleaning in the background.
    """
    try:
        batch = ImportBatch.objects.filter(batch_id=batch_id).first()
        if not batch:
            set_progress_failed(batch_id, "Import batch record not found.")
            return {'status': 'FAILED', 'error': 'Batch not found'}

        if is_cancelled(batch_id):
            return {'status': 'CANCELLED'}

        import os
        if not batch.file_path or not os.path.exists(batch.file_path):
            err_msg = f"Uploaded file not accessible on worker filesystem: {batch.file_path}"
            logger.error(err_msg)
            set_progress_failed(batch_id, err_msg, step=3)
            return {'status': 'FAILED', 'error': err_msg}

        update_progress(batch_id, stage='Validating and cleaning', status='PROCESSING', force_db_save=True)

        def _progress_cb(processed: int, total: Optional[int], failures: int, interim_stats: Optional[Dict] = None, stage: Optional[str] = None):
            if is_cancelled(batch_id):
                raise RuntimeError("Import job cancelled by user.")
            st = interim_stats or {}
            update_progress(
                batch_id=batch_id,
                stage=stage or 'Validating and cleaning',
                status='PROCESSING',
                step=3,
                unit='rows',
                processed_rows=processed,
                total_rows=total,
                validation_failures=failures,
                proposed_new=st.get('proposed_new'),
                proposed_updates=st.get('proposed_updates'),
                unchanged=st.get('unchanged'),
                quarantined=st.get('quarantined'),
                missing_invalid=st.get('missing_invalid'),
            )

        summary = stage_and_validate_batch(batch, strict_phone=strict_phone, progress_callback=_progress_cb)

        if is_cancelled(batch_id):
            return {'status': 'CANCELLED'}

        set_progress_completed(batch_id, summary=summary)
        return {'status': 'COMPLETED', 'summary': summary}

    except Exception as exc:
        logger.error(f"Error staging batch {batch_id}: {exc}", exc_info=True)
        if "cancelled" in str(exc).lower():
            set_progress_failed(batch_id, "Import cancelled by user.")
            return {'status': 'CANCELLED'}
        set_progress_failed(batch_id, str(exc))
        if hasattr(self, 'retry') and not getattr(settings, 'CELERY_TASK_ALWAYS_EAGER', False):
            raise self.retry(exc=exc, countdown=10)
        raise exc


@shared_task(bind=True, max_retries=3)
def process_import_batch_task(self: Any, batch_id: Any, user_id: Optional[int] = None) -> Any:
    """
    Celery task to publish a staged import batch idempotently with checkpoints.
    """
    try:
        if isinstance(batch_id, int):
            batch = ImportBatch.objects.get(id=batch_id)
        else:
            batch = ImportBatch.objects.get(batch_id=str(batch_id))

        user = User.objects.filter(id=user_id).first() if user_id else None

        def _progress_cb(processed: int, total: int, db_stats: Optional[Dict] = None, stage: Optional[str] = None):
            update_progress(
                batch_id=batch.batch_id,
                stage=stage or 'Importing to database',
                status='IMPORTING',
                step=4,
                unit='source rows committed',
                processed_rows=processed,
                total_rows=total,
                db_stats=db_stats
            )

        result = execute_import_publication(batch, user=user, progress_callback=_progress_cb)
        set_progress_completed(batch.batch_id, summary={
            'total_rows': batch.total_input_rows,
            'proposed_new': batch.proposed_new_records,
            'proposed_updates': batch.proposed_updates,
            'unchanged': batch.unchanged_records,
            'quarantined': batch.quarantined_rows,
            'missing_invalid': batch.missing_invalid_values,
            'db_stats': result.get('db_stats') if isinstance(result, dict) else {},
        }, final_status='COMPLETED', step=4)
        return result

    except Exception as exc:
        logger.error(f"Error publishing batch {batch_id}: {exc}", exc_info=True)
        set_progress_failed(str(batch_id), str(exc))
        if hasattr(self, 'retry') and not getattr(settings, 'CELERY_TASK_ALWAYS_EAGER', False):
            raise self.retry(exc=exc, countdown=10)
        raise exc


@shared_task(bind=True, max_retries=2)
def process_export_job_task(self: Any, job_id: int) -> Any:
    """
    Celery task to generate a CSV or Excel export file with streaming memory.
    """
    try:
        job = ExportJob.objects.get(id=job_id)
        path = run_export_job(job)
        return {'status': 'COMPLETED', 'file_path': path}
    except Exception as exc:
        logger.error(f"Error processing export job {job_id}: {exc}", exc_info=True)
        if hasattr(self, 'retry') and not getattr(settings, 'CELERY_TASK_ALWAYS_EAGER', False):
            raise self.retry(exc=exc, countdown=5)
        raise exc
