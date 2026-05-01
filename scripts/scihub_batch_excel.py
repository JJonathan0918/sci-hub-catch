#!/usr/bin/env python3
"""
Read DOI column from Web of Science–style .xlsx (header row contains 'DOI'),
then download PDFs using cookies/session from scripts/scihub_with_cookies.py.

Requires: pip install openpyxl

Example (use project venv so openpyxl + requests match this repo):
  .venv\\Scripts\\python scripts/scihub_batch_excel.py \\
    savedrecs-1-1000-search-atomic_clusters_and_photoelectron_spectroscopy.xlsx \\
    machine_learning_potential_experiments_1_1000.xlsx \\
    -o pdfs_batch \\
    --delay 20 --resume

Successful DOIs are appended to ``-o/downloaded_dois.txt`` (see ``--progress-file``).
Use ``--resume`` to skip those lines on the next run (interrupt-safe continuation).
With ``--resume``, DOIs listed in ``failed_dois.txt`` are also skipped by default
(``--retry-failed`` overrides). End-of-run merges new failures into ``failed_dois.txt``
without dropping older entries.

Cookies: fill COOKIES in scihub_with_cookies.py or set SCIHUB_COOKIES_JSON.
"""

from __future__ import annotations

import argparse
import importlib.util
import re
import sys
import time
from pathlib import Path

import requests

try:
    import openpyxl  # noqa: F401
except ImportError:
    print(
        "Missing dependency: openpyxl.\n"
        "  Install:  .venv\\Scripts\\python -m pip install openpyxl\n"
        "  Then run: .venv\\Scripts\\python scripts/scihub_batch_excel.py …\n"
        "(Plain `python` may point to another Python without this package.)",
        file=sys.stderr,
    )
    raise SystemExit(3) from None

_ROOT = Path(__file__).resolve().parent
_SPEC = importlib.util.spec_from_file_location(
    "scihub_with_cookies", _ROOT / "scihub_with_cookies.py")
if _SPEC is None or _SPEC.loader is None:
    raise RuntimeError("Cannot load scihub_with_cookies.py")
sch = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(sch)


def _find_wos_doi_column(ws):  # type: ignore
    """Return (header_row_1based, doi_col_1based)."""
    for ri in range(1, 501):
        row = [ws.cell(row=ri, column=c).value for c in range(1, 80)]
        for ci, val in enumerate(row, start=1):
            if val is None:
                continue
            if str(val).strip().upper() == "DOI":
                return ri, ci
    raise ValueError(
        'No header cell equals "DOI" in rows 1–500 / columns A–BZ. '
        "Use --doi-column-letter if your sheet differs."
    )


def _collect_dois(
    path: Path,
    *,
    doi_letter: str | None,
) -> list[str]:
    from openpyxl import load_workbook  # noqa: PLC0415

    wb = load_workbook(path, read_only=False, data_only=True)
    try:
        ws = wb.active
        if doi_letter:
            from openpyxl.utils.cell import column_index_from_string

            doi_col = column_index_from_string(doi_letter.strip().upper())
            header_row = 11
            for ri in range(1, 120):
                v = ws.cell(row=ri, column=doi_col).value
                if v is not None and str(v).strip().upper() == "DOI":
                    header_row = ri
                    break
        else:
            header_row, doi_col = _find_wos_doi_column(ws)

        out: list[str] = []
        seen: set[str] = set()
        r = header_row + 1
        while True:
            v = ws.cell(row=r, column=doi_col).value
            r += 1
            if v is None:
                if r > header_row + 50000:
                    break
                continue
            s = str(v).strip()
            if not s:
                continue
            if s in seen:
                continue
            seen.add(s)
            out.append(s)
    finally:
        wb.close()
    return out


def _output_name_for_doi(doi: str, pdf_url: str) -> str:
    tail = pdf_url.split("?")[0].rsplit("/", 1)[-1]
    if tail.endswith(".pdf") and tail:
        return tail
    return re.sub(r"[^\w.\-\+]+", "_", doi).strip("_")[:200] + ".pdf"


def _read_done_dois(progress_path: Path) -> set[str]:
    if not progress_path.is_file():
        return set()
    out: set[str] = set()
    for line in progress_path.read_text(encoding="utf-8").splitlines():
        t = line.strip()
        if t and not t.startswith("#"):
            out.add(t)
    return out


def _append_done_doi(progress_path: Path, doi: str) -> None:
    with progress_path.open("a", encoding="utf-8") as f:
        f.write(doi + "\n")


def _retryable_network_error(exc: BaseException) -> bool:
    if isinstance(
        exc,
        (
            requests.exceptions.Timeout,
            requests.exceptions.ConnectionError,
            requests.exceptions.ChunkedEncodingError,
            requests.exceptions.ProxyError,
        ),
    ):
        return True
    err = getattr(exc, "__cause__", None)
    if isinstance(err, TimeoutError):
        return True
    return False


def main() -> None:
    p = argparse.ArgumentParser(description="Batch Sci-Hub PDF from WoS Excel DOI column.")
    p.add_argument("xlsx_files", nargs="+", type=Path, help=".xlsx files with a DOI column")
    p.add_argument(
        "-o",
        "--out",
        type=Path,
        default=Path("pdfs_batch"),
        help="Output folder for PDF files",
    )
    p.add_argument(
        "--delay",
        type=float,
        default=4.0,
        help="Seconds to sleep after each DOI attempt (default: 4; e.g. 20 is conservative)",
    )
    p.add_argument("--base", default=None, help="Sci-Hub base URL (override script default)")
    p.add_argument(
        "--timeout",
        type=float,
        default=180.0,
        metavar="SEC",
        help="Read timeout (seconds) for each HTTP response (default: 180)",
    )
    p.add_argument(
        "--connect-timeout",
        type=float,
        default=90.0,
        metavar="SEC",
        help="Connect + TLS handshake timeout (default: 90; raise if handshake stalls)",
    )
    p.add_argument(
        "--retries",
        type=int,
        default=2,
        metavar="N",
        help="Extra attempts on timeouts/connection drops per DOI (default: 2)",
    )
    p.add_argument(
        "--doi-column-letter",
        default=None,
        metavar="COL",
        help='Force DOI column, e.g. "Q" if not auto-detected',
    )
    p.add_argument(
        "--limit",
        type=int,
        default=0,
        help="Maximum number of downloads total (0 = no limit)",
    )
    p.add_argument(
        "--skip-existing",
        action="store_true",
        help="After resolving the PDF URL, skip GET if the output file already exists",
    )
    p.add_argument(
        "--resume",
        action="store_true",
        help="Skip DOIs already recorded in the progress file (one DOI per line)",
    )
    p.add_argument(
        "--retry-failed",
        action="store_true",
        help="With --resume, still try DOIs listed in failed_dois.txt (default: skip them)",
    )
    p.add_argument(
        "--progress-file",
        default="downloaded_dois.txt",
        metavar="NAME",
        help="File under -o listing finished DOIs for --resume (default: downloaded_dois.txt)",
    )
    args = p.parse_args()

    req_timeout = (float(args.connect_timeout), float(args.timeout))

    cookies = sch._load_cookies()
    if not cookies:
        print("No cookies; fill COOKIES in scihub_with_cookies.py", file=sys.stderr)
        sys.exit(2)

    base = sch._normalize_base(args.base or sch.SCIHUB_BASE_URL)

    sess = sch.make_requests_session()

    args.out.mkdir(parents=True, exist_ok=True)
    progress_path = args.out / args.progress_file
    failed = args.out / "failed_dois.txt"
    failed_lines: list[str] = []
    done = 0

    # Merge DOIs from all spreadsheets (preserve order; skip duplicates across files).
    merged: list[str] = []
    seen_union: set[str] = set()
    for xlsx in args.xlsx_files:
        if not xlsx.is_file():
            print(f"Skip (not found): {xlsx}", file=sys.stderr)
            continue
        try:
            part = _collect_dois(xlsx, doi_letter=args.doi_column_letter)
        except Exception as ex:
            print(f"{xlsx}: {ex}", file=sys.stderr)
            continue
        print(f"== {xlsx.name}: {len(part)} DOIs in sheet")
        for d in part:
            if d not in seen_union:
                seen_union.add(d)
                merged.append(d)
    print(f"Total distinct DOIs: {len(merged)}")

    if args.resume:
        already = _read_done_dois(progress_path)
        before = len(merged)
        merged = [d for d in merged if d not in already]
        print(
            f"--resume: skipped {before - len(merged)} DOIs already in {progress_path.name}; "
            f"{len(merged)} remaining"
        )
        if not args.retry_failed and failed.is_file():
            failed_prev = _read_done_dois(failed)
            if failed_prev:
                before2 = len(merged)
                merged = [d for d in merged if d not in failed_prev]
                print(
                    f"--resume: skipped {before2 - len(merged)} DOIs listed in {failed.name} "
                    f"(use --retry-failed to try them again); {len(merged)} remaining"
                )

    stopped_at_limit = False
    for doi in merged:
        if args.limit and done >= args.limit:
            failed_lines.append(f"# stopped at limit {args.limit}")
            stopped_at_limit = True
            break

        for attempt in range(args.retries + 1):
            try:
                pdf_url = sch.sci_hub_post_and_resolve_pdf_url(
                    doi,
                    sess,
                    cookies,
                    base,
                    req_timeout,
                )
                name = _output_name_for_doi(doi, pdf_url)
                out_file = args.out / name

                if args.skip_existing and out_file.exists():
                    print(f"Skip existing file: {out_file.name}")
                    _append_done_doi(progress_path, doi)
                    done += 1
                    break

                pdf_bytes = sch.sci_hub_get_pdf_bytes(
                    sess, cookies, pdf_url, req_timeout
                )
                out_file.write_bytes(pdf_bytes)
                _append_done_doi(progress_path, doi)
                print(f"OK {done + 1}: {doi} → {out_file.name}")
                done += 1
                break
            except RuntimeError as e:
                print(f"FAIL: {doi}: {e}", file=sys.stderr)
                failed_lines.append(doi)
                break
            except requests.exceptions.RequestException as e:
                if attempt < args.retries and _retryable_network_error(e):
                    wait = 5.0 * (2**attempt)
                    print(
                        f"Transient error for {doi!r} (attempt {attempt + 1}/{args.retries}): {e!s}; "
                        f"retry in {wait:.0f}s",
                        file=sys.stderr,
                    )
                    time.sleep(wait)
                    continue
                print(f"FAIL: {doi}: {e}", file=sys.stderr)
                failed_lines.append(doi)
                break

        time.sleep(args.delay)

    new_failed_dois = [x for x in failed_lines if x and not x.startswith("#")]
    comment_lines = [x for x in failed_lines if x.startswith("#")]
    prev_failed = _read_done_dois(failed) if failed.is_file() else set()
    merged_failed = sorted(prev_failed | set(new_failed_dois))
    failed.write_text(
        "\n".join(merged_failed + comment_lines) + ("\n" if merged_failed or comment_lines else ""),
        encoding="utf-8",
    )
    if stopped_at_limit:
        print(f"Stopped at --limit {args.limit}")
    print(f"Finished. Saved under {args.out.resolve()}")
    print(
        f"This run new failures: {len(new_failed_dois)}; "
        f"total unique in {failed.name}: {len(merged_failed)}"
    )


if __name__ == "__main__":
    main()
