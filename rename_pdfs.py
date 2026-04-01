"""
Academic PDF Renamer
====================
Batch-renames PDF files in a target directory using embedded metadata.
When embedded metadata is incomplete, the script searches online using
the CrossRef API — first by DOI (extracted from the PDF text), then by
a free-text search of the paper's first page.

Naming convention:  Last, First (Year) - Title.pdf
Fallbacks:
  - No year  →  Last, First - Title.pdf
  - No author →  Title.pdf
  - No title  →  file is skipped

Dependencies:
  pip install PyMuPDF requests
"""

import os
import re
import sys
import time

import fitz      # PyMuPDF — fast, reliable, handles most PDFs gracefully
import requests  # HTTP client for CrossRef API lookups


# ──────────────────────────────────────────────
#  Filename sanitization
# ──────────────────────────────────────────────

# Characters that are illegal on Windows and/or macOS file systems.
# Windows forbids:  < > : " / \ | ? *
# macOS forbids:    / and : (in Finder; : maps to / on HFS+)
ILLEGAL_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

# Maximum filename length (including .pdf extension).
# 255 is the typical FS limit; we use 250 to leave a small safety margin.
MAX_FILENAME_LENGTH = 250

# ── CrossRef API configuration ──
# The free CrossRef API is polite-use; adding a mailto gets you into the
# "polite pool" with better rate limits.  Replace with your own email.
CROSSREF_MAILTO = ""  # e.g. "you@example.com" (optional but recommended)
CROSSREF_TIMEOUT = 10  # seconds


def sanitize(text: str) -> str:
    """Remove illegal filesystem characters and collapse whitespace."""
    text = ILLEGAL_CHARS_RE.sub("", text)   # strip forbidden chars
    text = re.sub(r"\s+", " ", text)        # collapse runs of whitespace
    return text.strip()


# ──────────────────────────────────────────────
#  Date / year parsing
# ──────────────────────────────────────────────

def extract_year(raw_date: str | None) -> str | None:
    """
    Try to pull a 4-digit year from common PDF date formats.

    PDF dates may look like:
      D:20190315120000+05'30'   (official PDF spec)
      2019-03-15T12:00:00       (ISO-ish)
      March 15, 2019            (free-form)
      2019                      (just the year)

    We search for the first plausible 4-digit year (1900–2099).
    """
    if not raw_date:
        return None

    # Match the first 4-digit number that looks like a reasonable year.
    match = re.search(r"((?:19|20)\d{2})", raw_date)
    return match.group(1) if match else None


# ──────────────────────────────────────────────
#  DOI extraction from PDF text
# ──────────────────────────────────────────────

# DOI pattern: "10." followed by a registrant code (4+ digits), a slash,
# and a suffix that continues until whitespace or certain punctuation.
# We also handle the common "doi:" and "https://doi.org/" prefixes.
DOI_RE = re.compile(
    r"(?:doi\s*[:=]\s*|https?://(?:dx\.)?doi\.org/)?"  # optional prefix
    r"(10\.\d{4,}/[^\s\"<>}{)]+)",                     # the DOI itself
    re.IGNORECASE,
)


def extract_doi(pdf_path: str) -> str | None:
    """
    Scan the first 3 pages of a PDF for a DOI string.

    Returns the raw DOI (e.g. '10.1000/xyz123') or None.
    DOIs most commonly appear on the first page (header/footer area),
    but we check a few extra pages just in case.
    """
    try:
        doc = fitz.open(pdf_path)
        text = ""
        # Scan up to the first 3 pages (or fewer if the doc is shorter).
        for page_num in range(min(3, len(doc))):
            text += doc[page_num].get_text()
        doc.close()
    except Exception:
        return None

    match = DOI_RE.search(text)
    if not match:
        return None

    doi = match.group(1)

    # Clean trailing punctuation that regex may have captured.
    # DOIs shouldn't end with periods, commas, or semicolons.
    doi = doi.rstrip(".,;:")
    return doi


# ──────────────────────────────────────────────
#  CrossRef API lookups
# ──────────────────────────────────────────────

def _crossref_headers() -> dict:
    """Return headers for polite CrossRef API usage."""
    headers = {"Accept": "application/json"}
    if CROSSREF_MAILTO:
        headers["User-Agent"] = f"AcademicPDFRenamer/1.0 (mailto:{CROSSREF_MAILTO})"
    return headers


def _parse_crossref_item(item: dict) -> dict:
    """
    Extract title, author, and year from a single CrossRef 'work' item.

    CrossRef returns authors as a list of {"given": "...", "family": "..."}
    dicts.  We take the first author and build "First Last" so that
    format_author_name() can rearrange it later.
    """
    # ── Title ──
    titles = item.get("title", [])
    title = titles[0] if titles else None

    # ── Author ──
    authors = item.get("author", [])
    author = None
    if authors:
        first = authors[0]
        given = first.get("given", "")
        family = first.get("family", "")
        if given and family:
            author = f"{given} {family}"
        elif family:
            author = family

    # ── Year ──
    # CrossRef stores dates under several keys; try the most common ones.
    year = None
    for date_key in ("published-print", "published-online", "issued", "created"):
        date_parts = item.get(date_key, {}).get("date-parts", [[]])
        if date_parts and date_parts[0]:
            candidate = str(date_parts[0][0])  # first element is the year
            if re.match(r"^(19|20)\d{2}$", candidate):
                year = candidate
                break

    return {"title": title, "author": author, "year": year}


def lookup_by_doi(doi: str) -> dict | None:
    """
    Query the CrossRef API for metadata associated with a DOI.

    Returns a dict with keys 'title', 'author', 'year' — or None on failure.
    Example endpoint: https://api.crossref.org/works/10.1000/xyz123
    """
    url = f"https://api.crossref.org/works/{doi}"
    try:
        resp = requests.get(url, headers=_crossref_headers(), timeout=CROSSREF_TIMEOUT)
        if resp.status_code != 200:
            return None
        data = resp.json()
        item = data.get("message", {})
        return _parse_crossref_item(item)
    except Exception:
        return None


def search_crossref(query_text: str) -> dict | None:
    """
    Search the CrossRef API using free-form text (e.g. a paper title
    extracted from the first page).

    We take the top-1 result and hope it matches.  CrossRef's relevance
    ranking is generally excellent for title-like queries.

    Returns a dict with keys 'title', 'author', 'year' — or None.
    """
    # Truncate very long queries — CrossRef works best with concise input.
    query_text = query_text[:300]

    url = "https://api.crossref.org/works"
    params = {"query": query_text, "rows": 1}

    if CROSSREF_MAILTO:
        params["mailto"] = CROSSREF_MAILTO

    try:
        resp = requests.get(
            url, params=params, headers=_crossref_headers(), timeout=CROSSREF_TIMEOUT
        )
        if resp.status_code != 200:
            return None
        data = resp.json()
        items = data.get("message", {}).get("items", [])

        if not items:
            return None

        # CrossRef returns a relevance score.  If the score is very low,
        # the result is likely garbage — discard it.
        best = items[0]
        score = best.get("score", 0)
        if score < 1:
            return None

        return _parse_crossref_item(best)
    except Exception:
        return None


def extract_first_page_text(pdf_path: str) -> str | None:
    """
    Return the text of the first page, lightly cleaned.

    This is used as a free-text query to CrossRef when no DOI is found
    and embedded metadata is incomplete.
    """
    try:
        doc = fitz.open(pdf_path)
        if len(doc) == 0:
            doc.close()
            return None
        text = doc[0].get_text().strip()
        doc.close()
        # Collapse excessive whitespace / newlines into spaces.
        text = re.sub(r"\s+", " ", text)
        return text if text else None
    except Exception:
        return None


# ──────────────────────────────────────────────
#  Metadata extraction  (local → DOI → search)
# ──────────────────────────────────────────────

def get_metadata(pdf_path: str) -> dict:
    """
    Open a PDF and return a dict with keys 'title', 'author', 'year'.
    Values are sanitized strings or None when absent.

    Raises an exception for unreadable / encrypted files so the caller
    can log and skip them.
    """
    doc = fitz.open(pdf_path)

    # PyMuPDF exposes metadata as a plain dict with lowercase keys.
    meta = doc.metadata or {}
    doc.close()

    title = (meta.get("title") or "").strip() or None
    author = (meta.get("author") or "").strip() or None

    # Try 'creationDate' first, then 'modDate' as a fallback.
    raw_date = (meta.get("creationDate") or meta.get("modDate") or "").strip()
    year = extract_year(raw_date) if raw_date else None

    # Sanitize the text fields that will become part of the filename.
    if title:
        title = sanitize(title)
    if author:
        author = sanitize(author)

    return {"title": title, "author": author, "year": year}


def enrich_metadata(meta: dict, pdf_path: str, filename: str) -> dict:
    """
    Fill in missing metadata fields by querying the internet.

    Strategy (each step only runs if fields are still missing):
      1. Extract a DOI from the PDF text → query CrossRef by DOI.
      2. Extract text from the first page → search CrossRef by text.

    Local (embedded) metadata always takes priority — online results
    only fill in the blanks.
    """
    # Quick check: if we already have all three fields, no lookup needed.
    if meta["title"] and meta["author"] and meta["year"]:
        return meta

    online = None  # will hold the dict from CrossRef if successful

    # ── Step 1: Try DOI-based lookup ──
    doi = extract_doi(pdf_path)
    if doi:
        print(f"    ↳ Found DOI: {doi} — querying CrossRef …")
        online = lookup_by_doi(doi)
        if online:
            print(f"    ↳ CrossRef DOI lookup succeeded.")

    # ── Step 2: Fallback — free-text search ──
    if online is None or not any(online.values()):
        first_page = extract_first_page_text(pdf_path)
        if first_page:
            # Use the first ~200 chars as a query (usually contains the title).
            query = first_page[:200]
            print(f"    ↳ No DOI found — searching CrossRef by text …")
            online = search_crossref(query)
            if online:
                print(f"    ↳ CrossRef text search returned a candidate.")

    if not online:
        return meta  # nothing found online; return what we have

    # ── Merge: local metadata wins; online fills the gaps ──
    if not meta["title"] and online.get("title"):
        meta["title"] = sanitize(online["title"])
    if not meta["author"] and online.get("author"):
        meta["author"] = sanitize(online["author"])
    if not meta["year"] and online.get("year"):
        meta["year"] = online["year"]

    return meta


# ──────────────────────────────────────────────
#  Author name formatting
# ──────────────────────────────────────────────

def format_author_name(author: str) -> str:
    """
    Convert an author string to "Last, First" format.

    Handles several common patterns found in PDF metadata:
      "John Smith"            → "Smith, John"
      "John A. Smith"         → "Smith, John A."
      "Smith, John"           → "Smith, John"   (already correct)
      "Smith, John A."        → "Smith, John A." (already correct)
      "J. Smith"              → "Smith, J."
      "Smith"                 → "Smith"          (single name, kept as-is)
      "John Smith; Jane Doe"  → "Smith, John"    (use first author only)
      "John Smith and Jane"   → "Smith, John"    (use first author only)

    Only the *first* author is used when multiple are listed.
    """
    if not author:
        return author

    # ── Take only the first author if multiple are listed ──
    # Common separators: semicolons, " and ", " & ", " et al"
    first_author = re.split(r"\s*[;]\s*|\s+and\s+|\s*&\s*|\s+et\s+al", author, maxsplit=1)[0].strip()

    if not first_author:
        return author

    # ── Already in "Last, First" format ──
    if "," in first_author:
        return first_author

    # ── Single name (no spaces) — return as-is ──
    parts = first_author.split()
    if len(parts) == 1:
        return first_author

    # ── Multi-part name: treat the last token as the surname ──
    # "John A. Smith" → last="Smith", rest="John A."
    last = parts[-1]
    rest = " ".join(parts[:-1])
    return f"{last}, {rest}"


# ──────────────────────────────────────────────
#  Filename construction
# ──────────────────────────────────────────────

def build_filename(meta: dict) -> str | None:
    """
    Build the target filename (without extension) according to the rules:
      Last, First (Year) - Title
      Last, First - Title       (year missing)
      Title                     (author missing)
      None                      (title missing → skip)
    """
    title = meta["title"]
    author = meta["author"]
    year = meta["year"]

    if not title:
        return None  # signal to skip

    # Rearrange the author into "Last, First" format.
    if author:
        author = format_author_name(author)

    if author and year:
        base = f"{author} ({year}) - {title}"
    elif author:
        base = f"{author} - {title}"
    else:
        base = title

    return base


def unique_path(directory: str, base_name: str, extension: str = ".pdf") -> str:
    """
    Return a full path that does not collide with existing files.

    If 'Author (2020) - Title.pdf' already exists, try:
      Author (2020) - Title (1).pdf
      Author (2020) - Title (2).pdf
      ...
    """
    # Truncate the base name so the full filename stays within limits.
    # Account for extension length and potential suffix like " (99)".
    max_base = MAX_FILENAME_LENGTH - len(extension)
    base_name = base_name[:max_base]

    candidate = os.path.join(directory, base_name + extension)
    if not os.path.exists(candidate):
        return candidate

    # Append an incrementing counter until we find a free name.
    counter = 1
    while True:
        suffix = f" ({counter})"
        truncated = base_name[: max_base - len(suffix)]
        candidate = os.path.join(directory, truncated + suffix + extension)
        if not os.path.exists(candidate):
            return candidate
        counter += 1


# ──────────────────────────────────────────────
#  Main renaming logic
# ──────────────────────────────────────────────

def rename_pdfs(folder: str) -> None:
    """Iterate over every PDF in *folder* and rename using metadata."""

    if not os.path.isdir(folder):
        print(f"[ERROR] '{folder}' is not a valid directory.")
        sys.exit(1)

    pdf_files = [f for f in os.listdir(folder) if f.lower().endswith(".pdf")]

    if not pdf_files:
        print("[INFO] No PDF files found in the target directory.")
        return

    print(f"[INFO] Found {len(pdf_files)} PDF(s) in '{folder}'.\n")

    renamed = 0
    skipped = 0

    for filename in pdf_files:
        filepath = os.path.join(folder, filename)
        print(f"  Processing: '{filename}'")

        try:
            meta = get_metadata(filepath)
        except fitz.fitz.FileDataError:
            # Corrupted or unreadable PDF data.
            print(f"  [SKIP] '{filename}' — corrupted or unreadable PDF data.")
            skipped += 1
            continue
        except Exception as e:
            # Catch-all for password-protected files, permission errors, etc.
            print(f"  [SKIP] '{filename}' — {type(e).__name__}: {e}")
            skipped += 1
            continue

        # ── If metadata is incomplete, try online lookups ──
        if not (meta["title"] and meta["author"] and meta["year"]):
            try:
                meta = enrich_metadata(meta, filepath, filename)
            except Exception as e:
                print(f"    ↳ Online lookup failed ({type(e).__name__}: {e}) — using local metadata only.")

            # Small delay between API calls to respect CrossRef rate limits.
            time.sleep(0.5)

        base_name = build_filename(meta)

        if base_name is None:
            print(f"  [SKIP] '{filename}' — no title found in metadata or online.")
            skipped += 1
            continue

        new_path = unique_path(folder, base_name)
        new_filename = os.path.basename(new_path)

        # Don't rename if the file already has the target name.
        if new_filename == filename:
            print(f"  [OK]   '{filename}' — already correctly named.")
            continue

        try:
            os.rename(filepath, new_path)
            print(f"  [RENAMED] '{filename}'\n"
                  f"         → '{new_filename}'")
            renamed += 1
        except OSError as e:
            print(f"  [ERROR] Could not rename '{filename}' — {e}")
            skipped += 1

    # ── Summary ──
    print(f"\n{'─' * 50}")
    print(f"  Renamed : {renamed}")
    print(f"  Skipped : {skipped}")
    print(f"  Total   : {len(pdf_files)}")
    print(f"{'─' * 50}")


# ──────────────────────────────────────────────
#  Entry point
# ──────────────────────────────────────────────

if __name__ == "__main__":
    print("Academic PDF Renamer")
    print("Type 'q' or 'quit' to exit.\n")

    while True:
        try:
            target = input("Enter the path to the folder containing PDFs: ").strip()
        except (KeyboardInterrupt, EOFError):
            # Handle Ctrl+C / Ctrl+D gracefully.
            print("\nGoodbye!")
            break

        # Allow the user to quit with a keyword.
        if target.lower() in ("q", "quit", "exit"):
            print("Goodbye!")
            break

        # Allow drag-and-drop paths wrapped in quotes on Windows.
        target = target.strip('"').strip("'")

        if not target:
            print("[WARN] No path entered. Try again.\n")
            continue

        rename_pdfs(target)
        print()  # blank line before the next prompt
