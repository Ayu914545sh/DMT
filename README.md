# Unified MCA & GST Data Management Webapp

An enterprise-grade, PostgreSQL-backed data management platform built with **Python 3.13, Django 5.2 LTS, and PostgreSQL 18**. Designed for authorized staff to ingest, preview, clean, validate, search, review, and export both MCA (Companies & LLPs with attached Directors) and GST (Tax Registrations) datasets.

---

## 1. System Architecture & Capabilities

- **Dedicated Relational Models**:
  - `MCACompany`: Enforces uniqueness on normalized `(entity_type, entity_id)` (CIN/LLPIN). Stores capital, ROC, incorporation dates, email, and unmapped fields in JSONB.
  - `MCADirectorEntry`: Child table linked via `ForeignKey` to `MCACompany`. Attaches multiple directors per company without deduplicating companies or dropping person records.
  - `GSTRecord`: Authoritative registration master. Keeps `legal_name` and `trade_name` separate. Preserves raw status/type codes (e.g. `ACT`, `REG`).
- **Strict Phone Cleaning Engine**:
  - Canonical mobile numbers are **strictly 10 ASCII digits or NULL**.
  - Strips leading backticks (`` ` ``), apostrophes (`'`), quotes, and surrounding whitespace.
  - Normalizes `+91`, `0091`, bare `91` (12-digit candidate only), and domestic leading `0` (11-digit candidate only).
  - Enforces strict domestic mobile prefix policy (`^[6-9]\d{9}$`).
  - Does **not** truncate arbitrary strings to the last 10 digits; does **not** strip `91` from an already valid 10-digit number.
- **Universal Ingestion Pipeline**:
  - Accepts both `.csv` and `.xlsx` for both MCA and GST datasets.
  - Worksheet selection for Excel; delimiter & encoding detection for CSV.
  - Auto-suggests column mapping with approved aliases while preventing ambiguous mappings.
  - Change preview showing proposed **New**, **Updates**, **Unchanged**, and **Quarantined** rows with before/after diffs before database commit.
- **Search & CRM Readiness**:
  - Multi-attribute search across companies, directors, and GST registrations.
  - Phone lookups attribute records to their source entities without falsely merging individuals.
- **Audited Formula-Safe Exports**:
  - CSV/XLSX export engine escapes dangerous characters (`=`, `+`, `-`, `@`) to protect against spreadsheet formula execution.
  - Formats phone numbers and IDs as text to prevent Excel precision loss or scientific notation truncation.
  - Full audit logging of who requested, what filters were applied, and row counts.

---

## 2. Default Access Credentials

Run the seeding command (`python manage.py seed_initial_data`) to create these default accounts:

| Role | Username | Password | Purpose |
| :--- | :--- | :--- | :--- |
| **Administrator** | `admin` | `admin123` | Full superuser access, system policies, template administration, Django Admin |
| **Data Manager** | `datamanager` | `manager123` | Upload datasets, map columns, stage/publish data, resolve quality issues |
| **Viewer** | `viewer` | `viewer123` | Read-only inspection of companies, GST registrations, and global search |

---

## 3. Local Windows Setup Instructions

### Prerequisites
- Python 3.13+ installed and added to PATH.
- PostgreSQL 16+ running locally on port 5432.

### Step 1: Clone or Navigate to the Workspace
```powershell
cd c:\Users\ayush\Downloads\DMT
```

### Step 2: Install Pinned Dependencies
```powershell
pip install -r requirements.txt
```

### Step 3: Configure Database Environment
Ensure your PostgreSQL database `dmt_db` exists and credentials are set in `.env`:
```env
DJANGO_SECRET_KEY=local-dev-secret-key-mca-gst-2026
DJANGO_DEBUG=True
DJANGO_ALLOWED_HOSTS=*

DB_ENGINE=postgresql
DB_NAME=dmt_db
DB_USER=postgres
DB_PASSWORD=ayush
DB_HOST=localhost
DB_PORT=5432

CELERY_BROKER_URL=redis://localhost:6379/0
CELERY_RESULT_BACKEND=redis://localhost:6379/0
CELERY_ALWAYS_EAGER=True
```

### Step 4: Apply Database Migrations & Seed Default Data
```powershell
python manage.py migrate
python manage.py seed_initial_data
```

### Step 5: Start the Development Server
```powershell
python manage.py runserver 8000
```
Open your browser and navigate to:
- **Custom Workspace**: [http://localhost:8000/](http://localhost:8000/)
- **Django Admin**: [http://localhost:8000/admin/](http://localhost:8000/admin/)

---

## 4. Production Deployment with Docker Compose

Deploy the web application, Celery background worker, PostgreSQL database, Redis cache, and Nginx reverse proxy on a Linux VPS using a single command:

```bash
docker compose up -d --build
```

### Services Deployed:
1. `web`: Gunicorn WSGI server running Django 5.2 LTS.
2. `worker`: Celery worker for asynchronous ingestion jobs.
3. `postgres`: PostgreSQL 18 container with persistent volume `postgres_data`.
4. `redis`: Redis 7 container with persistent volume `redis_data`.
5. `nginx`: Nginx reverse proxy serving static/media files on port 80/443.

### Container Management Commands:
- Check container logs: `docker compose logs -f web`
- Run database migrations: `docker compose exec web python manage.py migrate`
- Create an additional admin user: `docker compose exec web python manage.py createsuperuser`
- Check worker status: `docker compose logs -f worker`

---

## 5. Automated Verification & Test Suite

The system includes an extensive automated test suite covering all functional gates:
- Phone cleaning edge cases (backticks, apostrophes, 12-digit bare 91, 11-digit domestic 0, multiple numbers rejection, scientific notation rejection).
- MCA ingestion on 2,688 rows: verifying 1,160 unique companies/LLPs created, 2,688 distinct director relationships attached, and exactly 8 flagged phones logged to Quality Issues.
- GST ingestion on 3,034 rows: verifying legal vs trade name separation and 1 flagged phone logged.
- Search attribution by mobile number.
- Formula injection escaping in exports.

### Run Tests:
```powershell
python manage.py test hub
```
**Test Result**: `Ran 8 tests in 26.7s ... OK (0 failures, 0 errors)`

---

## 6. Database Backup & Restore Procedures

### Backup PostgreSQL Database
```powershell
# Local Windows / Host
pg_dump -h localhost -p 5432 -U postgres -d dmt_db -F c -b -v -f "dmt_db_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').dump"

# Inside Docker Compose
docker compose exec postgres pg_dump -U postgres -d dmt_db -F c -b -v -f /var/lib/postgresql/data/backup.dump
```

### Restore PostgreSQL Database
```powershell
# Local Windows / Host
pg_restore -h localhost -p 5432 -U postgres -d dmt_db -v -c "dmt_db_backup.dump"

# Inside Docker Compose
docker compose exec postgres pg_restore -U postgres -d dmt_db -v -c /var/lib/postgresql/data/backup.dump
```

### Backup Uploaded Media & Ingested Files
```powershell
tar -czvf "media_backup_$(Get-Date -Format 'yyyyMMdd').tar.gz" media/
```
