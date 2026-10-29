#!/usr/bin/env python3
"""Pull the roadmap workbook from its share link and rebuild the chart.

    python3 tools/sync.py                      # uses the link below
    python3 tools/sync.py <share-or-file-url>  # or any other link

The GitHub workflow runs exactly this, so what you see here is what a failed
run would have said. Afterwards commit what changed:

    git add data gantt.html && git commit -m "Update roadmap" && git push

Nothing is written until the download has been checked: a password page, a
server error or a workbook with a renamed column leaves the committed data
untouched.
"""
import io
import os
import re
import ssl
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
import build  # noqa: E402  (shares the validation the workflow relies on)

TARGET = ROOT / "data" / "durf-roadmap.xlsx"
# The WebDAV public endpoint serves the file itself, so nothing is appended.
DEFAULT_URL = "https://surf.works.surf.nl/public.php/dav/files/WP9W6mL7QHsxRKz"

TIMEOUT = 30                     # seconds, for connecting and for each read
RETRY_WAITS = (10, 30)           # wait before the 2nd and 3rd attempt
TRANSIENT = {408, 425, 429, 500, 502, 503, 504}

# A Nextcloud/ownCloud /s/<token> link serves a viewer page; the file is one
# segment further on. Other shapes -- /public.php/dav/files/<token>, or a plain
# path to an .xlsx -- already point at the bytes, so leave them alone.
SHARE_PAGE = re.compile(r"/s/[A-Za-z0-9]+$")

# What each status most likely means for a public Nextcloud share. These are
# pointers for a person, not diagnoses: the server's own words are printed too.
HINTS = {
    401: "The share asks for a password. Remove it, or no automated sync can read the file.",
    403: "The server refuses the download. Check the share is not set to 'Hide download', "
         "and that nothing blocks the address this runs from.",
    404: "Nothing is served at this address. The share may have been deleted or its link "
         "regenerated -- copy the new WebDAV URL into the ROADMAP_SHARE_URL variable.",
    423: "The file is locked, usually because it is open for editing somewhere.",
    429: "Too many requests from this address.",
    503: "The server says it is unavailable: Nextcloud in maintenance, overloaded, or "
         "declining requests from this address. That is not about the workbook's columns. "
         "Open the URL from another network: if it downloads there, the server is refusing "
         "this address; if it fails there too, the fault is on the Nextcloud side.",
}

# Nextcloud answers 503 -- not 404 -- when a public share's link still exists but
# the file behind it does not. The giveaway is a NotFound wrapped inside the
# ServiceUnavailable. Seen for real on this project, so it gets its own advice.
SHARE_TARGET_GONE = (
    "The link still exists, but the file it pointed at does not: Nextcloud reports "
    "'ServiceUnavailable' wrapping 'NotFound'. The workbook was most likely renamed, "
    "moved, deleted or replaced (uploaded again, or 'saved as' a new file) since the "
    "link was made, which orphans it. Check that durf-roadmap.xlsx is still in the "
    "team folder, create a new public share link for it, and put its WebDAV address "
    "(.../public.php/dav/files/<token>) in the ROADMAP_SHARE_URL repository variable.")

# Response headers worth showing; anything else (cookies in particular) is not.
SHOW_HEADERS = ("server", "retry-after", "content-type", "content-length", "via",
                "www-authenticate", "location", "x-request-id", "x-powered-by")


class FetchError(Exception):
    def __init__(self, status, reason, headers=None, body=""):
        super().__init__(f"HTTP {status} {reason}" if status else reason)
        self.status = status
        self.headers = headers or {}
        self.body = body


def say(kind, message):
    """Print; on GitHub Actions also raise it as an annotation on the run."""
    if os.environ.get("GITHUB_ACTIONS"):
        print(f"::{kind}::{message.splitlines()[0]}")
    print(("ERROR: " if kind == "error" else "") + message)


def to_file_url(url):
    url = url.rstrip("/")
    return url + "/download" if SHARE_PAGE.search(url) else url


def fetch_once(url):
    request = urllib.request.Request(url, headers={"User-Agent": "durf-gantt-sync"})
    try:
        with urllib.request.urlopen(request, timeout=TIMEOUT,
                                    context=ssl.create_default_context()) as res:
            return res.read()
    except urllib.error.HTTPError as err:
        body = err.read(600).decode("utf-8", "replace")
        raise FetchError(err.code, err.reason, dict(err.headers), body) from None
    except (urllib.error.URLError, TimeoutError, ConnectionError) as err:
        reason = getattr(err, "reason", err)
        raise FetchError(0, f"could not connect ({reason})") from None


def download(url):
    """Fetch with retries on transient failures, and explain a final failure."""
    url = to_file_url(url)
    print(f"fetching {url}")
    last = None
    for attempt in range(1, len(RETRY_WAITS) + 2):
        try:
            return fetch_once(url)
        except FetchError as err:
            last = err
            print(f"  attempt {attempt}: {err}")
            gone = err.status == 503 and "notfound" in err.body.lower()
            retryable = (err.status == 0 or err.status in TRANSIENT) and not gone
            if not retryable or attempt > len(RETRY_WAITS):
                break
            wait = RETRY_WAITS[attempt - 1]
            print(f"  waiting {wait}s before trying again")
            time.sleep(wait)
    explain(last)
    raise SystemExit(1)


def explain(err):
    lines = [f"Could not download the workbook: {err}."]
    if err.status == 503 and "notfound" in err.body.lower():
        lines.append(SHARE_TARGET_GONE)
    elif err.status in HINTS:
        lines.append(HINTS[err.status])
    elif err.status == 0:
        lines.append("The connection itself failed, before the server could answer. "
                     "Check the address, and whether this network can reach the host.")
    shown = {k.lower(): v for k, v in err.headers.items() if k.lower() in SHOW_HEADERS}
    if shown:
        lines.append("Server headers: " + "; ".join(f"{k}: {v}" for k, v in shown.items()))
    snippet = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", err.body)).strip()
    if snippet:
        lines.append(f"Server said: {snippet[:300]}")
    say("error", "\n".join(lines))


def check(payload):
    """A password prompt, an expiry or a spent download limit all arrive as
    HTML with a 200, so never trust the status code on its own."""
    if payload[:4] != b"PK\x03\x04":
        head = re.sub(r"\s+", " ", payload[:200].decode("utf-8", "replace")).strip()
        say("error", "That link did not return a spreadsheet.\n"
                     f"It starts with: {head[:140]}\n"
                     "Check the share is public, unexpired, has no password and no "
                     "download limit.")
        raise SystemExit(1)

    sheets = build.read_sheets(io.BytesIO(payload))
    problems = build.validate(sheets)
    if problems:
        for problem in problems:
            say("error", problem)
        say("error", "The workbook was downloaded but not used: fix the columns above "
                     "in Nextcloud. The published chart keeps its last good version.")
        raise SystemExit(1)
    print("  sheets: " + ", ".join(k for k in sheets if k != build.META_KEY))


def main():
    url = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_URL
    payload = download(url)
    check(payload)

    before = TARGET.read_bytes() if TARGET.exists() else b""
    if payload == before:
        print("Already up to date — nothing changed.")
        return

    TARGET.write_bytes(payload)
    print(f"  wrote {TARGET.relative_to(ROOT)} ({len(payload) / 1024:.1f} kB)")
    build.main()
    print("\nDone. Commit the result:\n"
          "    git add data gantt.html && git commit -m \"Update roadmap\" && git push")


if __name__ == "__main__":
    main()
