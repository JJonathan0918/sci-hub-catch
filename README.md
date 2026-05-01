# Sci-Hub EVA (Batch DOI Downloader)

This repo now focuses on **batch PDF downloading from DOI lists** (especially Web of Science Excel exports), using CLI scripts with retry and resume support.

> [!TIP]
> Sci-Hub availability is unstable and often protected by anti-bot systems. Success depends heavily on network path, IP quality, and valid cookies/session.

## What This Project Supports

- Batch read DOI values from one or more `.xlsx` files
- Deduplicate DOI values across files
- Download PDFs with retries and timeout controls
- Resume from interruption (`downloaded_dois.txt`)
- Track failures (`failed_dois.txt`)
- Skip existing files (`--skip-existing`)

Main scripts:

- `scripts/scihub_with_cookies.py`: single DOI workflow, cookie-based session
- `scripts/scihub_batch_excel.py`: batch workflow for Excel DOI columns

## Quick Start

### 1) Create and activate venv

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 2) Install dependencies

```powershell
python -m pip install -U pip
python -m pip install requests lxml openpyxl
```

### 3) Configure cookies

Open `scripts/scihub_with_cookies.py` and fill `COOKIES` with the values copied from browser DevTools for the target Sci-Hub host.

You can also use:

```powershell
$env:SCIHUB_COOKIES_JSON='{"name":"value","name2":"value2"}'
```

### 4) Run batch download

```powershell
.\.venv\Scripts\python.exe scripts/scihub_batch_excel.py `
  savedrecs-1-1000-search-atomic_clusters_and_photoelectron_spectroscopy.xlsx `
  machine_learning_potential_experiments_1_1000.xlsx `
  -o pdfs_batch `
  --delay 20 --resume --skip-existing
```

## Expected Excel Format

The batch script auto-detects a header cell equal to `DOI` and reads that column.

For your current files, it detects:

- header row: around row `11`
- DOI column: `Q` (17th column)

If auto-detection fails, set manually:

```powershell
--doi-column-letter Q
```

## Resume / Failure Logic

- Successful DOI (or skipped-existing DOI) is appended to `downloaded_dois.txt`
- Failed DOI is merged into `failed_dois.txt`
- `--resume` skips DOI already in `downloaded_dois.txt`
- With `--resume`, failed DOIs are also skipped by default
- Use `--retry-failed` to explicitly retry failed DOIs

## Useful Options

- `--delay 20`: sleep seconds between DOI attempts
- `--limit 50`: process only first N remaining DOIs (for testing)
- `--connect-timeout 120`: connect/TLS handshake timeout
- `--timeout 300`: response read timeout
- `--retries 4`: extra retries on transient network errors
- `--base https://sci-hub.red/`: override host

Example (more tolerant for unstable VPN):

```powershell
.\.venv\Scripts\python.exe scripts/scihub_batch_excel.py `
  machine_learning_potential_experiments_1_1000.xlsx `
  -o pdfs_batch_ml `
  --delay 20 --resume --skip-existing `
  --connect-timeout 120 --timeout 300 --retries 4
```

## Output Structure

Under your `-o` folder (for example `pdfs_batch/`):

- downloaded PDFs (`*.pdf`)
- `downloaded_dois.txt`
- `failed_dois.txt`

## Notes

- Do not commit real cookie/session values.
- Keep generated data and large PDFs out of Git (see `.gitignore`).
- If network is unstable, prefer longer timeout + retries + resume.

## GUI Notes (Optional)

The original GUI app is still present (`app.py`), but this README emphasizes the batch CLI flow.

## Building Original Desktop App

See [`building/README.md`](building/README.md)
