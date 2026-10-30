"""
Management command to safely publish a staged import batch directly into the authoritative PostgreSQL database.
Resumable, idempotent, checkpoint-aware, memory-bounded, and includes destination database verification.

Usage:
    python manage.py publish_batch <BATCH_ID> [--batch-size 2000] [--verify] [--reset-checkpoint]
"""

import sys
import time
from django.core.management.base import BaseCommand, CommandError
from django.utils import timezone
from hub.models import ImportBatch, StagedRow, MCACompany, MCADirectorEntry, GSTRecord
from hub.importers import execute_import_publication


class Command(BaseCommand):
    help = 'Safely publish a staged import batch directly into PostgreSQL with verification.'

    def add_arguments(self, parser):
        parser.add_argument('batch_id', type=str, help='The batch_id to publish (e.g. BATCH-20261030_053330-D164EA76)')
        parser.add_argument('--batch-size', type=int, default=2000, help='Chunk size for bulk database operations (default 2000)')
        parser.add_argument('--reset-checkpoint', action='store_true', help='Reset checkpoint and re-process all staged rows from row 0')
        parser.add_argument('--verify-only', action='store_true', help='Skip publication and only verify existing destination database records')

    def handle(self, *args, **options):
        batch_id = options['batch_id']
        batch_size = options['batch_size']
        reset_checkpoint = options['reset_checkpoint']
        verify_only = options['verify_only']

        self.stdout.write(self.style.MIGRATE_HEADING(f"=== Publish Batch Runner: {batch_id} ==="))

        batch = ImportBatch.objects.filter(batch_id=batch_id).first()
        if not batch:
            raise CommandError(f"ImportBatch '{batch_id}' not found in database.")

        staged_count = StagedRow.objects.filter(batch=batch).count()
        unprocessed_count = StagedRow.objects.filter(batch=batch, is_processed=False).count()

        self.stdout.write(f"Dataset Type    : {batch.dataset_type}")
        self.stdout.write(f"Source File     : {batch.file_name}")
        self.stdout.write(f"Current Status  : {batch.status}")
        self.stdout.write(f"Staged Rows     : {staged_count:,} total ({unprocessed_count:,} unprocessed)")
        self.stdout.write(f"Checkpoint Row  : {batch.checkpoint_row or 0:,}")

        if verify_only:
            self.verify_destination(batch)
            return

        if staged_count == 0:
            raise CommandError(f"Batch {batch_id} has 0 staged rows in StagedRow table. Has Staging & Validation completed?")

        if reset_checkpoint:
            self.stdout.write(self.style.WARNING("Resetting checkpoint and staged rows to unprocessed..."))
            StagedRow.objects.filter(batch=batch).update(is_processed=False)
            batch.checkpoint_row = 0
            batch.imported_records = 0
            batch.status = 'PREVIEWED'
            batch.save(update_fields=['checkpoint_row', 'imported_records', 'status'])
            self.stdout.write(self.style.SUCCESS("Checkpoint reset complete."))

        self.stdout.write(self.style.NOTICE(f"\nStarting publication in chunks of {batch_size:,}..."))
        start_time = time.time()
        last_reported = 0

        def _progress(processed, total, db_stats=None, stage=None):
            nonlocal last_reported
            now = time.time()
            if now - last_reported >= 1.0 or processed >= total:
                last_reported = now
                pct = (processed / total * 100) if total else 0
                elapsed = now - start_time
                rate = processed / elapsed if elapsed > 0 else 0
                stats_str = ""
                if db_stats:
                    if batch.dataset_type == 'MCA':
                        stats_str = f" | Created: {db_stats.get('companies_created', 0):,} | Updated: {db_stats.get('companies_updated', 0):,} | Dirs: {db_stats.get('directors_created', 0):,}"
                    else:
                        stats_str = f" | Created: {db_stats.get('gst_created', 0):,} | Updated: {db_stats.get('gst_updated', 0):,}"
                self.stdout.write(
                    f"\r[{pct:5.1f}%] {processed:,}/{total:,} rows committed ({rate:,.0f} rows/s, elapsed: {elapsed:.1f}s){stats_str}",
                    ending=""
                )
                sys.stdout.flush()

        try:
            res = execute_import_publication(batch, progress_callback=_progress)
            self.stdout.write("\n")
            total_elapsed = time.time() - start_time
            self.stdout.write(self.style.SUCCESS(
                f"\nPublication complete in {total_elapsed:.1f}s! Total rows committed: {res.get('imported_records', 0):,}"
            ))
            if 'db_stats' in res:
                self.stdout.write(self.style.MIGRATE_LABEL("Database Writes Summary:"))
                for k, v in res['db_stats'].items():
                    self.stdout.write(f"  - {k}: {v:,}")

            # Verify destination database
            self.verify_destination(batch)

        except Exception as e:
            self.stdout.write("\n")
            self.stdout.write(self.style.ERROR(f"Error during publication: {e}"))
            raise CommandError(str(e))

    def verify_destination(self, batch):
        self.stdout.write(self.style.MIGRATE_HEADING("\n=== Destination Database Verification ==="))
        if batch.dataset_type == 'MCA':
            total_companies = MCACompany.objects.filter(is_archived=False).count()
            total_directors = MCADirectorEntry.objects.count()
            self.stdout.write(f"Total MCACompany records in DB       : {total_companies:,}")
            self.stdout.write(f"Total MCADirectorEntry records in DB : {total_directors:,}")

            # Check representative CINs from this batch
            sample_staged_cins = list(
                StagedRow.objects.filter(batch=batch).exclude(entity_identifier='').values_list('entity_identifier', flat=True)[:5]
            )
            if sample_staged_cins:
                self.stdout.write(f"\nChecking sample entity identifiers from batch: {sample_staged_cins}")
                found = MCACompany.objects.filter(entity_id__in=sample_staged_cins)
                self.stdout.write(self.style.SUCCESS(f"Found {found.count()}/{len(sample_staged_cins)} sample companies in MCACompany table:"))
                for comp in found:
                    dir_cnt = comp.directors.count()
                    self.stdout.write(f"  * CIN: {comp.entity_id} | Name: {comp.name[:35]} | State: {comp.state} | Directors: {dir_cnt}")
            else:
                self.stdout.write(self.style.WARNING("No entity identifiers found in staged rows to sample."))

        else:
            total_gst = GSTRecord.objects.count()
            self.stdout.write(f"Total GSTRecord records in DB: {total_gst:,}")
            sample_staged_gst = list(
                StagedRow.objects.filter(batch=batch).exclude(entity_identifier='').values_list('entity_identifier', flat=True)[:5]
            )
            if sample_staged_gst:
                self.stdout.write(f"\nChecking sample GST numbers from batch: {sample_staged_gst}")
                found = GSTRecord.objects.filter(gst_number__in=sample_staged_gst)
                self.stdout.write(self.style.SUCCESS(f"Found {found.count()}/{len(sample_staged_gst)} sample records in GSTRecord table:"))
                for rec in found:
                    self.stdout.write(f"  * GSTIN: {rec.gst_number} | Trade Name: {rec.trade_name[:35]} | State: {rec.state}")
