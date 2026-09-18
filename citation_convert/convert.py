#!/usr/bin/env python3
"""
citation_convert v4 — Convert numbered citations to EndNote temp citations.

Uses authoritative RIS downloaded directly from CrossRef (not self-constructed
XML), eliminating hallucination risk for author names, years, and journal data.

Outputs:
  <stem>_converted.docx  — citations replaced with {Author, Year #N}
  <stem>.ris             — combined RIS file for EndNote import

doi_overrides.json (optional, place next to the docx):
  Provides DOIs for references that lack them in the reference list, and
  manual RIS entries for items without DOIs (WHO books, reports, etc.).
  See template at ~/research/tools/citation_convert/doi_overrides_template.json

Usage:
    python3 convert.py <input.docx> [--email EMAIL] [--dry-run]
"""

import argparse
import json
import re
import shutil
import sys
import time
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    import requests
except ImportError:
    sys.exit("pip install requests")


# ── 1. Extract plain text from docx ─────────────────────────────────────────

def docx_plain(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    xml = re.sub(r"</w:p>", "\n", xml)
    return re.sub(r"<[^>]+>", "", xml)


# ── 2. Extract reference list ────────────────────────────────────────────────

def _extract_doi_from_text(text):
    for pat in (r"https?://doi\.org/(\S+)", r"\bdoi:\s*(\S+)"):
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1).rstrip(".,)")
    return None


def extract_references(docx_path):
    text = docx_plain(docx_path)
    m = re.search(r"\n(?:References|REFERENCES|Bibliography)\n(.+)", text, re.DOTALL)
    if not m:
        raise ValueError("Cannot find 'References' section in document.")
    ref_text = m.group(1)

    refs = []
    # Lancet/Vancouver: [N] text
    bracket = list(re.finditer(r"\[(\d+)\]\s*(.+?)(?=\n\[|\Z)", ref_text, re.DOTALL))
    if bracket:
        for bm in bracket:
            body = " ".join(bm.group(2).split())
            refs.append({"num": int(bm.group(1)), "text": body,
                         "doi": _extract_doi_from_text(body)})
    else:
        # PLoS / APA: N. text
        for pm in re.finditer(r"(?:^|\n)(\d+)\.\s+(.+?)(?=\n\d+\.|\Z)", ref_text, re.DOTALL):
            body = " ".join(pm.group(2).split())
            refs.append({"num": int(pm.group(1)), "text": body,
                         "doi": _extract_doi_from_text(body)})

    if not refs:
        # Fallback: plain unnumbered paragraphs (Vancouver without [N] prefix).
        # Stop at figure legends / table headers / footnote blocks.
        stop_words = ("figure legends", "figures", "tables", "footnotes", "supplementary")
        for i, line in enumerate(ref_text.splitlines(), 1):
            body = line.strip()
            if not body:
                continue
            if any(body.lower().startswith(w) for w in stop_words):
                break
            body = " ".join(body.split())
            refs.append({"num": i, "text": body, "doi": _extract_doi_from_text(body)})

    if not refs:
        raise ValueError("No numbered references found. Expected '[N] ...' or 'N. ...' format.")
    return refs


# ── 3. doi_overrides.json ────────────────────────────────────────────────────

def load_overrides(docx_path):
    """
    Load <stem>.doi_overrides.json if it exists next to the docx.
    Returns dict keyed by str(ref_num).
    """
    p = Path(docx_path)
    override_path = p.parent / (p.stem + ".doi_overrides.json")
    if override_path.exists():
        with open(override_path) as f:
            data = json.load(f)
        # Remove comment key if present
        data.pop("_note", None)
        print(f"  Loaded overrides: {override_path.name} ({len(data)} entries)")
        return data
    return {}


def get_doi(ref, overrides):
    """Return DOI for a reference: text first, then overrides."""
    if ref["doi"]:
        return ref["doi"]
    entry = overrides.get(str(ref["num"]), {})
    return entry.get("doi")


def get_manual_ris(ref, overrides):
    """Return manual RIS string for a reference if provided in overrides."""
    entry = overrides.get(str(ref["num"]), {})
    return entry.get("ris")


# ── 4. Fetch RIS — CrossRef → doi.org → PubMed ──────────────────────────────

_EUTILS = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils"


def _pubmed_search(query, email):
    """Run a PubMed esearch and return the first PMID, or None."""
    try:
        r = requests.get(
            f"{_EUTILS}/esearch.fcgi",
            params={"db": "pubmed", "term": query, "retmode": "json",
                    "retmax": 1, "email": email},
            headers={"User-Agent": f"citation-convert/4.0 mailto:{email}"},
            timeout=15,
        )
        ids = r.json().get("esearchresult", {}).get("idlist", [])
        return ids[0] if ids else None
    except Exception:
        return None


def _pubmed_fetch_ris(pmid, email):
    """
    Fetch RIS from PubMed for a known PMID via the efetch API.
    PubMed returns Y1 (print date) — more accurate than CrossRef's epub PY.
    """
    try:
        r = requests.get(
            f"{_EUTILS}/efetch.fcgi",
            params={"db": "pubmed", "id": pmid, "rettype": "ris",
                    "retmode": "text", "email": email},
            headers={"User-Agent": f"citation-convert/4.0 mailto:{email}"},
            timeout=15,
        )
        if r.status_code == 200 and "TY  -" in r.text:
            return r.text.strip()
    except Exception:
        pass
    return None


def fetch_ris(doi, email):
    """
    Fetch authoritative RIS for a DOI.
    Order: CrossRef transform → doi.org → PubMed DOI search.
    PubMed uses print dates, so it also fixes epub-vs-print year issues
    for biomedical papers that CrossRef gets wrong.
    Returns RIS text string or None.
    """
    headers = {"User-Agent": f"citation-convert/4.0 mailto:{email}"}

    # 1. CrossRef transform (broadest coverage)
    try:
        r = requests.get(
            f"https://api.crossref.org/works/{doi}/transform/application/x-research-info-systems",
            headers=headers, timeout=15,
        )
        if r.status_code == 200 and "TY  -" in r.text:
            return r.text.strip()
    except Exception:
        pass

    # 2. doi.org content negotiation
    try:
        r = requests.get(
            f"https://doi.org/{doi}",
            headers={**headers, "Accept": "application/x-research-info-systems"},
            timeout=15, allow_redirects=True,
        )
        if r.status_code == 200 and "TY  -" in r.text:
            return r.text.strip()
    except Exception:
        pass

    # 3. PubMed DOI search (biomedical papers, print-year accurate)
    pmid = _pubmed_search(f"{doi}[DOI]", email)
    if pmid:
        ris = _pubmed_fetch_ris(pmid, email)
        if ris:
            print(f"    → found via PubMed PMID:{pmid}")
            return ris

    return None


def fetch_ris_no_doi(ref_text, email):
    """
    For references with no DOI: try PubMed bibliographic text search.
    Uses first-author + year as primary query, title words as fallback.
    Returns (ris_text, pmid) or (None, None).
    """
    # Query 1: first author + year (precise)
    author_m = re.search(r"^([A-Z][a-záéíóúñü\-]+)", ref_text.strip())
    year_m   = re.search(r"\b(19|20)\d{2}\b", ref_text)
    if author_m and year_m:
        query = f"{author_m.group(1)}[Author] AND {year_m.group(0)}[PDAT]"
        pmid = _pubmed_search(query, email)
        if pmid:
            ris = _pubmed_fetch_ris(pmid, email)
            if ris:
                return ris, pmid

    # Query 2: first 7 title words
    words = ref_text.split()[:7]
    query = " ".join(words) + "[Title]"
    pmid = _pubmed_search(query, email)
    if pmid:
        ris = _pubmed_fetch_ris(pmid, email)
        if ris:
            return ris, pmid

    return None, None


# ── 5. Parse RIS entry ───────────────────────────────────────────────────────

def parse_ris(ris_text):
    """
    Extract first author last name and publication year from a RIS entry.
    Returns (author_last, year_int).
    """
    author, year = "Unknown", 0

    for line in ris_text.splitlines():
        line = line.strip()
        tag = line[:2]
        val = line[6:].strip() if len(line) > 6 else ""

        if tag == "AU" and author == "Unknown" and val:
            # RIS AU format: "Last, First" or "Last, F." or "Organization Name"
            author = val.split(",")[0].strip()

        elif tag in ("PY", "DA", "Y1") and not year and val:
            # PY  - 2022  or  PY  - 2022/01/01
            m = re.search(r"\b(19|20)\d{2}\b", val)
            if m:
                year = int(m.group(0))

    return author, year


# ── 6. Auto-detect citation style ────────────────────────────────────────────

def detect_style(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    plain = re.sub(r"<[^>]+>", "", xml)
    brackets = len(re.findall(r"\[\d+\]", plain))
    sups = len(re.findall(r'vertAlign w:val="superscript"', xml))
    parens = len(re.findall(r'\(\d{1,3}(?:,\s*\d{1,3})*\)', plain))
    if parens > brackets and parens > sups:
        return "paren"
    return "bracket" if brackets >= sups else "superscript"


# ── 7. Find citation patterns ─────────────────────────────────────────────────

def _find_body_start(xml):
    for marker in (">Abstract<", ">ABSTRACT<", ">Introduction<",
                   ">INTRODUCTION<", ">Background<", ">Summary<"):
        pos = xml.find(marker)
        if pos > 0:
            para = xml.rfind("<w:p ", 0, pos)
            return para if para > 0 else pos
    return 0


def find_all_patterns(docx_path, style):
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    plain = re.sub(r"<[^>]+>", "", xml)

    patterns = set()
    if style == "bracket":
        patterns |= set(re.findall(r"\[\d+(?:,\s*\d+)+\]", plain))
        for m in re.finditer(r"\[(\d+)\]", plain):
            ctx = plain[m.end():m.end() + 3]
            if not re.match(r"\s+[A-Z]", ctx):
                patterns.add(m.group(0))
    elif style == "paren":
        body_start = _find_body_start(xml)
        ref_start  = xml.find(">References<")
        body_xml   = xml[body_start:ref_start] if ref_start > 0 else xml[body_start:]
        body_plain = re.sub(r"<[^>]+>", "", body_xml)
        # Multi-number combos are unambiguously citations
        patterns |= set(re.findall(r'\(\d{1,3}(?:,\s*\d{1,3})+\)', body_plain))
        # Single-number citations
        for m in re.finditer(r'\((\d{1,3})\)', body_plain):
            n = int(m.group(1))
            if 1 <= n <= 300:
                patterns.add(m.group(0))
    else:
        body_start = _find_body_start(xml)
        ref_start  = xml.find(">References<")
        body_xml   = xml[body_start:ref_start] if ref_start > 0 else xml[body_start:]
        for content in re.findall(
            r'<w:vertAlign w:val="superscript"[^>]*/>(?:[^<]*</w:rPr>)?<w:t[^>]*>([\d,\s]+)</w:t>',
            body_xml
        ):
            content = content.strip()
            nums = [int(n) for n in re.findall(r"\d+", content)]
            if nums and all(1 <= n <= 200 for n in nums):
                patterns.add(content)

    return patterns


# ── 8. Build citation map from RIS data ──────────────────────────────────────

def build_citation_map(refs, ris_data):
    """
    ris_data: {num: ris_text_string}
    Returns {num: {author, year, record_num, temp_cite}}
    record_num = 1..N in reference list order (fresh library → guaranteed match)
    """
    cmap = {}
    for i, ref in enumerate(sorted(refs, key=lambda r: r["num"])):
        num = ref["num"]
        rec = i + 1
        ris = ris_data.get(num, "")
        author, year = parse_ris(ris) if ris else ("Unknown", 0)
        if author == "Unknown":
            # Last fallback: first capitalised word from reference text
            m = re.search(r"[A-Z][a-záéíóúñü\-]+", ref["text"])
            author = m.group(0) if m else "Unknown"
        if not year:
            m = re.search(r"\b(19|20)\d{2}\b", ref["text"])
            year = int(m.group(0)) if m else 0

        cmap[num] = {
            "author":     author,
            "year":       year,
            "record_num": rec,
            "temp_cite":  f"{{{author}, {year} #{rec}}}",
        }
    return cmap


# ── 9. Build replacement table ───────────────────────────────────────────────

def build_replacements(patterns, cmap, style):
    repls = {}
    for pat in patterns:
        nums = [int(n) for n in re.findall(r"\d+", pat)]
        missing = [n for n in nums if n not in cmap]
        if missing:
            print(f"  WARNING: ref(s) {missing} in pattern '{pat}' not in reference list")
            continue
        if len(nums) == 1:
            repls[pat] = cmap[nums[0]]["temp_cite"]
        else:
            parts = [
                f"{cmap[n]['author']}, {cmap[n]['year']} #{cmap[n]['record_num']}"
                for n in nums
            ]
            repls[pat] = "{" + "; ".join(parts) + "}"
    return repls


# ── 10. Convert citations in docx XML ────────────────────────────────────────

PROOF = r"(?:<w:proofErr[^/]*/>\s*)*"


def _resolve_author(author_str, cmap):
    al = author_str.strip().lower()
    for num, info in cmap.items():
        if info["author"].lower() == al:
            return info
    return None


def convert_superscript_xml(xml, cmap):
    body_start = _find_body_start(xml)
    ref_start  = xml.find(">References<")
    if ref_start < 0:
        ref_start = len(xml)

    prefix = xml[:body_start]
    body   = xml[body_start:ref_start]
    tail   = xml[ref_start:]

    result = []
    pos = 0
    while pos < len(body):
        run_open = body.find("<w:r", pos)
        if run_open < 0:
            result.append(body[pos:])
            break
        result.append(body[pos:run_open])
        run_close = body.find("</w:r>", run_open)
        if run_close < 0:
            result.append(body[run_open:])
            break
        run_close += len("</w:r>")
        run_xml = body[run_open:run_close]
        pos = run_close

        if 'vertAlign w:val="superscript"' in run_xml:
            tm = re.search(r"<w:t[^>]*>([\d,\s]+)</w:t>", run_xml)
            if tm:
                content = tm.group(1).strip()
                nums  = [int(n) for n in re.findall(r"\d+", content)]
                valid = [n for n in nums if n in cmap]
                if valid:
                    if len(valid) == 1:
                        info = cmap[valid[0]]
                        cite = f"{{{info['author']}, {info['year']} #{info['record_num']}}}"
                    else:
                        parts = [
                            f"{cmap[n]['author']}, {cmap[n]['year']} #{cmap[n]['record_num']}"
                            for n in valid
                        ]
                        cite = "{" + "; ".join(parts) + "}"
                    result.append(f'<w:r><w:t xml:space="preserve"> {cite}</w:t></w:r>')
                    continue
        result.append(run_xml)

    return prefix + "".join(result) + tail


def convert_paren_xml(xml, repls, cmap):
    # Exact string replacement, longest pattern first to avoid partial matches
    for old, new in sorted(repls.items(), key=lambda x: -len(x[0])):
        xml = xml.replace(old, new)
    return xml


def convert_bracket_xml(xml, repls, cmap):
    # Pass 1: simple string replacements (longest first, bracket-safe)
    for old, new in sorted(repls.items(), key=lambda x: -len(x[0])):
        xml = xml.replace(old, new)

    # Pass 2: three-way split {  | Author | , year}
    def fix3(m):
        tag, pre, author, tail = m.group(1), m.group(2), m.group(3), m.group(4)
        info = _resolve_author(author, cmap)
        if not info:
            return m.group(0)
        cite = f"{{{author}, {info['year']} #{info['record_num']}}}"
        return f"{tag}{pre[:-1]}{cite}</w:t></w:r>"

    xml = re.sub(
        r"(<w:t[^>]*>)([^<]*\{)</w:t></w:r>" + PROOF +
        r"<w:r[^>]*><w:t[^>]*>([^<]+)</w:t></w:r>" + PROOF +
        r"<w:r[^>]*><w:t[^>]*>(, \d{4}[^<]*)\}</w:t></w:r>",
        fix3, xml, flags=re.DOTALL,
    )

    # Pass 3: four-way split { | Author | , <space> | year}
    def fix4(m):
        run_open, pre, author, year_str = m.group(1), m.group(2), m.group(3), m.group(5)
        info = _resolve_author(author, cmap)
        if not info:
            return m.group(0)
        cite = f"{{{author}, {info['year']} #{info['record_num']}}}"
        return f'{run_open}<w:t xml:space="preserve">{pre[:-1]}{cite}</w:t></w:r>'

    xml = re.sub(
        r"(<w:r[^>]*><w:t[^>]*>)([^<]*\{)</w:t></w:r>" + PROOF +
        r"<w:r[^>]*><w:t[^>]*>([^<]+)</w:t></w:r>" + PROOF +
        r"<w:r[^>]*><w:t[^>]*>(, )</w:t></w:r>" + PROOF +
        r"<w:r[^>]*><w:t[^>]*>(\d{4}[^<]*)\}</w:t></w:r>",
        fix4, xml, flags=re.DOTALL,
    )
    return xml


# ── 11. Apply to docx ────────────────────────────────────────────────────────

def apply_to_docx(src, dst, repls, cmap, style):
    import os
    tmp = str(dst) + ".tmp_convert"
    shutil.copy(src, tmp)

    with zipfile.ZipFile(tmp, "r") as zin, \
         zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                xml = data.decode("utf-8")
                if style == "superscript":
                    xml = convert_superscript_xml(xml, cmap)
                elif style == "paren":
                    xml = convert_paren_xml(xml, repls, cmap)
                else:
                    xml = convert_bracket_xml(xml, repls, cmap)
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    os.remove(tmp)

    # Validate
    with zipfile.ZipFile(dst) as z:
        raw = z.read("word/document.xml")
    try:
        ET.fromstring(raw)
    except ET.ParseError as e:
        print(f"  ERROR: XML invalid — {e}")
        return False

    plain = re.sub(r"<[^>]+>", "", raw.decode("utf-8"))
    bare  = [b for b in re.findall(r"\{[A-ZÀ-ž][^}#]{3,60}\}", plain) if "#" not in b]
    if bare:
        print(f"  WARNING: {len(bare)} citations still lack record numbers:")
        for b in set(bare):
            print(f"    {b}")
    else:
        print("  All citations have record numbers ✓")
    return True


# ── 12. Write combined RIS file ──────────────────────────────────────────────

def write_ris(ris_data, refs, ris_path):
    """Write RIS entries in reference list order (1..N)."""
    entries = []
    for ref in sorted(refs, key=lambda r: r["num"]):
        ris = ris_data.get(ref["num"])
        if ris:
            entries.append(ris.strip())
    ris_path.write_text("\n\n".join(entries) + "\n", encoding="utf-8")
    print(f"  RIS file → {ris_path}")


# ── 13. Main ─────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(
        description="Convert numbered .docx citations to EndNote temp citations (v3, RIS-based)"
    )
    ap.add_argument("docx", help="Input .docx with numbered citations")
    ap.add_argument("--output", help="Output .docx path")
    ap.add_argument("--ris",    help="Combined RIS output path")
    ap.add_argument("--email",  default="fakoredesodiq@gmail.com")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    src  = Path(args.docx).expanduser().resolve()
    stem = src.stem
    dst  = Path(args.output).resolve() if args.output else src.parent / f"{stem}_converted.docx"
    risp = Path(args.ris).resolve()    if args.ris    else src.parent / f"{stem}.ris"

    print(f"\n{'='*60}")
    print("  Citation Converter v4  (CrossRef + PubMed fallback)")
    print(f"  Input : {src}")
    print(f"{'='*60}")

    # Step 1
    print("\n[1/5] Extracting reference list...")
    refs = extract_references(src)
    print(f"  {len(refs)} references found")

    # Step 2
    print("\n[2/5] Loading DOI overrides...")
    overrides = load_overrides(src)

    # Step 3
    print("\n[3/5] Detecting citation style and patterns...")
    style    = detect_style(src)
    patterns = find_all_patterns(src, style)
    combined = [p for p in patterns if "," in p]
    print(f"  Style: {style} | {len(patterns)} patterns | {len(combined)} combined: {combined}")

    if args.dry_run:
        print("\n-- Dry run: DOI resolution preview --")
        for ref in refs:
            doi = get_doi(ref, overrides)
            manual = "MANUAL-RIS" if get_manual_ris(ref, overrides) else ""
            status = doi or manual or "MISSING"
            print(f"  [{ref['num']:2d}] {status}")
        return 0

    # Step 4: fetch RIS for every reference
    print("\n[4/5] Fetching RIS (CrossRef → doi.org → PubMed)...")
    ris_data = {}
    ok = missing = 0

    for ref in refs:
        num = ref["num"]

        # Manual RIS (WHO books, reports — from doi_overrides.json)
        manual = get_manual_ris(ref, overrides)
        if manual:
            ris_data[num] = manual
            print(f"  [{num:2d}] manual RIS")
            ok += 1
            continue

        doi = get_doi(ref, overrides)

        if doi:
            print(f"  [{num:2d}] {doi}", end="", flush=True)
            ris = fetch_ris(doi, args.email)
            if ris:
                ris_data[num] = ris
                print()
                ok += 1
            else:
                # CrossRef + doi.org + PubMed all failed — try PubMed text search
                print(" — all DOI sources failed, trying PubMed text search...")
                ris, pmid = fetch_ris_no_doi(ref["text"], args.email)
                if ris:
                    print(f"       → found via PubMed PMID:{pmid}")
                    ris_data[num] = ris
                    ok += 1
                else:
                    print(f"       → FAILED — add manual RIS to doi_overrides.json")
                    missing += 1
        else:
            # No DOI anywhere — try PubMed bibliographic search
            print(f"  [{num:2d}] no DOI — PubMed text search...", end="", flush=True)
            ris, pmid = fetch_ris_no_doi(ref["text"], args.email)
            if ris:
                print(f" found PMID:{pmid}")
                ris_data[num] = ris
                ok += 1
            else:
                print(f" not found — add to doi_overrides.json")
                missing += 1

        time.sleep(0.25)

    print(f"  {ok} fetched | {missing} still missing")
    if missing:
        print(f"  For missing refs: add DOI or manual RIS to {src.stem}.doi_overrides.json")

    # Step 5: build map, convert, write
    print("\n[5/5] Converting and writing outputs...")
    cmap  = build_citation_map(refs, ris_data)
    repls = build_replacements(patterns, cmap, style)

    # Backup first
    bak = src.parent / f"{stem}.bak.docx"
    if not bak.exists():
        shutil.copy(src, bak)
        print(f"  Backup → {bak.name}")

    write_ris(ris_data, refs, risp)
    ok = apply_to_docx(src, dst, repls, cmap, style)

    if ok:
        print(f"  Converted doc → {dst.name}")
        print(f"\n{'='*60}")
        print("  Done!\n")
        print("  Next steps:")
        print("  1. In EndNote: File → New  (fresh library for this paper)")
        print(f"  2. File → Import → File → {risp.name}")
        print("     Import Option: Reference Manager (RIS)")
        print("  3. Open the converted .docx in Word")
        print("  4. Delete the old reference list at the bottom")
        print("  5. EndNote ribbon → Update Citations and Bibliography")
        print(f"{'='*60}\n")
        return 0
    return 1


if __name__ == "__main__":
    sys.exit(main())
