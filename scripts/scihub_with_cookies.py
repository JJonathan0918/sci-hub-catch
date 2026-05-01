#!/usr/bin/env python3
"""
After passing the site's bot check in a browser, copy cookies from
DevTools → Application → Cookies → (this host). Prefer copying **every**
cookie row for that host (session + anti-bot/CDN/etc.); do not rely on a
single cookie named ``session`` only.

Paste name→value pairs below.

Usage (from repo root):
  .venv\\Scripts\\python scripts/scihub_with_cookies.py --doi 10.1016/j.surfrep.2007.07.001
  .venv\\Scripts\\python scripts/scihub_with_cookies.py --doi 10.1016/... --save .

Do not commit real cookies; keep this file local or use env SCIHUB_COOKIES_JSON.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests
import urllib3
from lxml.etree import HTML

# ---------------------------------------------------------------------------
# 1) Fill cookies here: for the Sci-Hub host you use, copy **all** cookies from
#    DevTools (not only one "session" cookie). Or set env:
#    SCIHUB_COOKIES_JSON='{"name":"value",...}'
# ---------------------------------------------------------------------------
COOKIES: dict[str, str] = {
    "__ddg1_": "7lybxCtZo56g5Zrx2Jex",
    "__ddg10_": "1777434275",
    "__ddg8_": "0UatHRVPeLyKSuGq",
    "__ddg9_": "43.245.223.80",
    "PHPSESSID": "122a6180fd9e00f3e02fbed397c71488",
    "refresh": "1777434274.0741",
    "session": "c1485df392ebdf8b31ed64397d82e935",
}

# Base Sci-Hub URL (no trailing path). Must match the site you used in browser.
SCIHUB_BASE_URL = "https://sci-hub.red/"

# Match SciHubEVA default PDF location XPaths.
PDF_XPATHS = [
    '//div[@id="article"]/embed[1]',
    '//div[@id="article"]/iframe[1]',
    '//embed[@id="pdf"]',
]

# Optional: set to your browser UA if the site is picky (else requests default).
USER_AGENT: str | None = None

REQUESTS_VERIFY_SSL = False

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def make_requests_session() -> requests.Session:
    """
    Build a Session for Sci-Hub requests.

    By default ``trust_env=False`` so ``HTTP_PROXY`` / ``HTTPS_PROXY`` from the
    shell are ignored (common cause of ProxyError to 127.0.0.1 when no proxy
    is actually listening). Set env ``SCIHUB_TRUST_ENV_PROXY=1`` to use system
    proxy variables again.
    """
    sess = requests.Session()
    sess.verify = REQUESTS_VERIFY_SSL
    trust = os.environ.get("SCIHUB_TRUST_ENV_PROXY", "").strip().lower()
    sess.trust_env = trust in ("1", "true", "yes", "on")
    if USER_AGENT:
        sess.headers["User-Agent"] = USER_AGENT
    return sess


def _resolve_to_absolute(href: str, response_url: str) -> str:
    return urljoin(response_url, href.strip())


def _load_cookies() -> dict[str, str]:
    raw = os.environ.get("SCIHUB_COOKIES_JSON", "").strip()
    if raw:
        data = json.loads(raw)
        if not isinstance(data, dict):
            raise SystemExit("SCIHUB_COOKIES_JSON must be a JSON object of strings.")
        return {str(k): str(v) for k, v in data.items()}
    return dict(COOKIES)


def _normalize_base(url: str) -> str:
    u = url.strip()
    if not u.endswith("/"):
        u += "/"
    return u


def _extract_pdf_url(html_bytes: bytes, response_url: str) -> str | None:
    tree = HTML(html_bytes)
    if tree is None:
        return None

    # 1) Legacy: embed / iframe under #article.
    for xp in PDF_XPATHS:
        found = tree.xpath(xp)
        if found:
            src = found[0].attrib.get("src")
            if src:
                return _resolve_to_absolute(src, response_url)

    # 2) Newer pages: <meta name="citation_pdf_url" content="/path/file.pdf">.
    for m in tree.xpath('//meta[@content]'):
        name = (m.get('name') or '').lower()
        if name != 'citation_pdf_url':
            continue
        raw = (m.get('content') or '').strip()
        if raw:
            return _resolve_to_absolute(raw, response_url)

    return None


def _guess_filename(query: str, pdf_url: str) -> str:
    tail = urlparse(pdf_url).path.rsplit("/", 1)[-1]
    safe = tail if tail.endswith(".pdf") else None
    if safe:
        return safe
    q = re.sub(r"[^\w.\-\+]+", "_", query)[:120]
    return f"{q}.pdf"


def sci_hub_post_and_resolve_pdf_url(
    doi: str,
    sess: requests.Session,
    cookies: dict[str, str],
    base_normalized: str,
    timeout: float | tuple[float, float],
    *,
    dump_html_path: str | None = None,
) -> str:
    """
    POST Sci-Hub once and parse the PDF URL from HTML (no GET yet).
    Raises RuntimeError on failure.
    """
    print(f"POST {base_normalized} request={doi!r} …")
    r = sess.post(
        base_normalized,
        data={"request": doi},
        cookies=cookies,
        timeout=timeout,
    )
    print(f"HTTP {r.status_code}, body length {len(r.content)}")
    if dump_html_path:
        Path(dump_html_path).write_bytes(r.content)
        print(f"Wrote HTML: {dump_html_path}")

    if r.status_code != 200:
        raise RuntimeError(f"POST failed: HTTP {r.status_code}")

    lower = r.text.lower()
    if "robot" in lower and "human" in lower:
        print(
            "Warning: HTML still looks like a challenge page.",
            file=sys.stderr,
        )

    pdf_url = _extract_pdf_url(r.content, r.url)
    if not pdf_url:
        snippet = r.text[:1200].replace("\n", " ")
        print(f"HTML preview: {snippet!r}", file=sys.stderr)
        raise RuntimeError(
            "No PDF URL in HTML (embed/iframe or citation_pdf_url meta)."
        )
    return pdf_url


def sci_hub_get_pdf_bytes(
    sess: requests.Session,
    cookies: dict[str, str],
    pdf_url: str,
    timeout: float | tuple[float, float],
) -> bytes:
    pr = sess.get(pdf_url, cookies=cookies, timeout=timeout)
    if pr.status_code != 200:
        raise RuntimeError(f"PDF GET failed: HTTP {pr.status_code}")
    ct = pr.headers.get("Content-Type", "")
    if "pdf" not in ct.lower() and not pdf_url.lower().endswith(".pdf"):
        print(f"Unexpected Content-Type: {ct!r}", file=sys.stderr)
    return pr.content


def sci_hub_download_pdf(
    doi: str,
    sess: requests.Session,
    cookies: dict[str, str],
    base_normalized: str,
    timeout: float | tuple[float, float],
    *,
    dump_html_path: str | None = None,
) -> tuple[str, bytes]:
    """
    POST Sci-Hub once, resolve PDF URL, GET body. Raises RuntimeError if the pipeline fails.
    Returns (resolved_pdf_url, pdf_bytes).
    """
    pdf_url = sci_hub_post_and_resolve_pdf_url(
        doi, sess, cookies, base_normalized, timeout, dump_html_path=dump_html_path
    )
    content = sci_hub_get_pdf_bytes(sess, cookies, pdf_url, timeout)
    return pdf_url, content


def main() -> None:
    parser = argparse.ArgumentParser(description="Sci-Hub query using browser cookies.")
    parser.add_argument(
        "--doi",
        required=True,
        help="DOI, PMID, title fragment, or full query string for ?request=…",
    )
    parser.add_argument(
        "--base",
        default=SCIHUB_BASE_URL,
        help=f"Sci-Hub base URL (default: {SCIHUB_BASE_URL!r})",
    )
    parser.add_argument(
        "--save",
        metavar="DIR",
        help="If set, download PDF into this directory.",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=120.0,
        help="HTTP timeout in seconds (default: 120)",
    )
    parser.add_argument(
        "--dump-html",
        metavar="FILE",
        help="Write the POST response HTML to this file (for debugging XPath).",
    )
    args = parser.parse_args()

    cookies = _load_cookies()
    if not cookies:
        print(
            "No cookies: fill COOKIES in this file or set SCIHUB_COOKIES_JSON.",
            file=sys.stderr,
        )
        sys.exit(2)

    base = _normalize_base(args.base)
    sess = make_requests_session()

    try:
        pdf_url, content = sci_hub_download_pdf(
            args.doi,
            sess,
            cookies,
            base,
            args.timeout,
            dump_html_path=args.dump_html,
        )
    except RuntimeError as e:
        print(str(e), file=sys.stderr)
        sys.exit(1)

    print(f"PDF URL: {pdf_url}")

    if args.save:
        out_dir = Path(args.save).resolve()
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / _guess_filename(args.doi, pdf_url)
        out_path.write_bytes(content)
        print(f"Saved: {out_path}")


if __name__ == "__main__":
    main()
