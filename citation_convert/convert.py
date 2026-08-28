#!/usr/bin/env python3
"""
citation_convert — Convert numbered citations in a Word .docx to EndNote temp citations.

Generates an EndNote XML file that:
  1. Can be imported into a fresh EndNote library (File → Import → XML)
  2. Can be indexed by endnote-mcp so Claude can verify record numbers

Usage:
    python convert.py <input.docx> [options]

Options:
    --output PATH    Output .docx path (default: <stem>_converted.docx)
    --xml PATH       EndNote XML path (default: <stem>_endnote.xml)
    --email EMAIL    Email for CrossRef polite pool
    --dry-run        Show detected refs and citations without converting
"""

import argparse
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


# ── 1. Extract plain text ────────────────────────────────────────────────────

def docx_plain(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    xml = re.sub(r"</w:p>", "\n", xml)
    return re.sub(r"<[^>]+>", "", xml)


# ── 2. Extract reference list ────────────────────────────────────────────────

def _extract_doi(text):
    for pat in (r"https?://doi\.org/(\S+)", r"\bdoi:\s*(\S+)"):
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1).rstrip(".,)")
    return None


def extract_references(docx_path):
    """
    Find numbered references in the References section.
    Handles two formats automatically:
      - Lancet/Vancouver: [1] Author ... DOI
      - PLoS/APA:         1. Author ... doi:...
    Returns list of dicts: {num, text, doi}
    """
    text = docx_plain(docx_path)
    m = re.search(r"\n(?:References|REFERENCES|Bibliography)\n(.+)", text, re.DOTALL)
    if not m:
        raise ValueError("Cannot find 'References' section in document.")
    ref_text = m.group(1)

    refs = []

    # Try Lancet format: [N] text
    bracket_matches = list(re.finditer(r"\[(\d+)\]\s*(.+?)(?=\n\[|\Z)", ref_text, re.DOTALL))
    if bracket_matches:
        for bm in bracket_matches:
            body = " ".join(bm.group(2).split())
            refs.append({"num": int(bm.group(1)), "text": body, "doi": _extract_doi(body)})
    else:
        # Try PLoS/numbered format: N. text (N followed by period+space at line start)
        plos_matches = list(re.finditer(r"(?:^|\n)(\d+)\.\s+(.+?)(?=\n\d+\.|\Z)", ref_text, re.DOTALL))
        for pm in plos_matches:
            body = " ".join(pm.group(2).split())
            refs.append({"num": int(pm.group(1)), "text": body, "doi": _extract_doi(body)})

    if not refs:
        raise ValueError(
            "No numbered references found. Expected '[N] Author...' or 'N. Author...' format."
        )
    return refs


# ── 3. Auto-detect all citation patterns ────────────────────────────────────

def detect_citation_style(docx_path):
    """
    Returns 'bracket' ([N] format) or 'superscript' (PLoS/numbered superscript).
    """
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    plain = re.sub(r"<[^>]+>", "", xml)
    bracket_count = len(re.findall(r"\[\d+\]", plain))
    sup_count = len(re.findall(r'vertAlign w:val="superscript"', xml))
    return "bracket" if bracket_count >= sup_count else "superscript"


def find_all_patterns(docx_path, style="auto"):
    """
    Return every citation pattern found in the document body.
    style: 'bracket' | 'superscript' | 'auto'

    For superscript style, reads directly from superscript XML runs
    (content like "1" or "1,2" or "25,2,26") — no plain-text guessing.
    """
    if style == "auto":
        style = detect_citation_style(docx_path)

    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    plain = re.sub(r"<[^>]+>", "", xml)

    patterns = set()

    if style == "bracket":
        # Combined: [5,6] [9,10] etc.
        patterns |= set(re.findall(r"\[\d+(?:,\s*\d+)+\]", plain))
        # Single: [1] — exclude ref-list entries (followed by space + capital)
        for m in re.finditer(r"\[(\d+)\]", plain):
            ctx = plain[m.end():m.end()+3]
            if not re.match(r"\s+[A-Z]", ctx):
                patterns.add(m.group(0))

    else:  # superscript
        # Read directly from superscript XML runs — body only (skip title page)
        body_start = _find_body_start(xml)
        ref_start  = xml.find(">References<")
        body_xml   = xml[body_start:ref_start] if ref_start > 0 else xml[body_start:]

        # Each superscript run may contain "1" or "1,2" or "10,11" etc.
        sup_contents = re.findall(
            r'<w:vertAlign w:val="superscript"[^>]*/>'
            r'(?:[^<]*</w:rPr>)?'           # close rPr
            r'<w:t[^>]*>([\d,\s]+)</w:t>',  # capture the number(s)
            body_xml
        )
        for content in sup_contents:
            content = content.strip()
            nums = [int(n) for n in re.findall(r"\d+", content)]
            # Filter: must be plausible citation numbers (1..200), not years/stats
            if nums and all(1 <= n <= 200 for n in nums):
                patterns.add(content)

    return patterns, style


# ── 4. CrossRef ─────────────────────────────────────────────────────────────

def crossref_fetch(doi, email):
    url = f"https://api.crossref.org/works/{doi}"
    hdrs = {"User-Agent": f"citation-convert/2.0 mailto:{email}"}
    try:
        r = requests.get(url, headers=hdrs, timeout=15)
        if r.status_code == 200:
            return r.json()["message"]
    except Exception as e:
        print(f"    CrossRef error for {doi}: {e}")
    return None


def crossref_year(msg):
    for f in ("published", "published-print", "published-online", "issued"):
        if f in msg and "date-parts" in msg[f]:
            p = msg[f]["date-parts"]
            if p and p[0]:
                return int(p[0][0])
    return None


def crossref_first_author(msg):
    authors = msg.get("author", [])
    if not authors:
        return msg.get("institution", [{}])[0].get("name", "Unknown").split()[0]
    a = authors[0]
    return a.get("family", a.get("name", "Unknown").split()[0])


def fetch_all(refs, email):
    meta = {}
    for ref in refs:
        num, doi = ref["num"], ref["doi"]
        if not doi:
            print(f"  [{num}] No DOI — will use fallback")
            continue
        print(f"  [{num}] {doi}")
        msg = crossref_fetch(doi, email)
        if msg:
            meta[num] = {
                "msg": msg,
                "year": crossref_year(msg),
                "author": crossref_first_author(msg),
                "title": (msg.get("title", [""])[0] or ""),
                "journal": (msg.get("container-title", [""])[0] or ""),
                "doi": doi,
                "type": msg.get("type", "journal-article"),
            }
        else:
            print(f"    No CrossRef data for [{num}]")
        time.sleep(0.25)
    return meta


# ── 5. Fallback author/year from reference text ──────────────────────────────

def parse_fallback(text):
    am = re.search(r"[A-Z][a-záéíóúñü\-]+", text)
    author = am.group(0) if am else "Unknown"
    ym = re.search(r"\b(19|20)\d{2}\b", text)
    year = int(ym.group(0)) if ym else 2000
    return author, year


# ── 6. Build citation map ────────────────────────────────────────────────────

def build_map(refs, meta):
    """
    Returns {num: {author, year, record_num, temp_cite}}
    record_num = 1-indexed import order (predictable in a fresh library)
    """
    cmap = {}
    for i, ref in enumerate(sorted(refs, key=lambda r: r["num"])):
        num = ref["num"]
        rec = i + 1
        if num in meta:
            author = meta[num]["author"]
            year = meta[num]["year"] or 0
        else:
            author, year = parse_fallback(ref["text"])
        cmap[num] = {
            "author": author,
            "year": year,
            "record_num": rec,
            "temp_cite": f"{{{author}, {year} #{rec}}}",
        }
    return cmap


# ── 7. Generate EndNote XML ──────────────────────────────────────────────────

_TYPES = {
    "journal-article":    ("Journal Article", 17),
    "book":               ("Book", 6),
    "book-chapter":       ("Book Section", 5),
    "posted-content":     ("Electronic Article", 43),
    "report":             ("Report", 27),
    "proceedings-article":("Conference Paper", 47),
    "dataset":            ("Dataset", 59),
}

def _esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def generate_endnote_xml(refs, meta, cmap):
    lines = ['<?xml version="1.0" encoding="UTF-8"?>', "<xml>", "<records>"]

    for ref in sorted(refs, key=lambda r: r["num"]):
        num = ref["num"]
        rec = cmap[num]["record_num"]
        m = meta.get(num)

        if m:
            msg = m["msg"]
            tname, tnum = _TYPES.get(m["type"], ("Journal Article", 17))
            authors = msg.get("author", [])
            title   = _esc(m["title"])
            journal = _esc(m["journal"])
            year    = cmap[num]["year"]
            vol     = msg.get("volume", "")
            iss     = msg.get("issue", "")
            pages   = msg.get("page", "").replace("-", "–")
            doi     = m["doi"]
            abstract = _esc(re.sub(r"<[^>]+>", "", msg.get("abstract", ""))[:2000])
        else:
            tname, tnum = "Journal Article", 17
            authors, title, journal = [], _esc(ref["text"][:120]), ""
            year    = cmap[num]["year"]
            vol = iss = pages = doi = abstract = ""

        lines += [
            "<record>",
            f"  <rec-number>{rec}</rec-number>",
            f'  <ref-type name="{tname}">{tnum}</ref-type>',
        ]
        if authors:
            lines.append("  <contributors><authors>")
            for a in authors[:20]:
                last  = a.get("family", a.get("name", ""))
                first = a.get("given", "")
                lines.append(f"    <author>{_esc(last + (', ' + first if first else ''))}</author>")
            lines.append("  </authors></contributors>")

        lines += [
            "  <titles>",
            f"    <title>{title}</title>",
            *([ f"    <secondary-title>{journal}</secondary-title>"] if journal else []),
            "  </titles>",
        ]
        if year:
            lines.append(f"  <dates><year>{year}</year></dates>")
        if vol:
            lines.append(f"  <volume>{vol}</volume>")
        if iss:
            lines.append(f"  <number>{iss}</number>")
        if pages:
            lines.append(f"  <pages>{pages}</pages>")
        if doi:
            lines += [
                f"  <electronic-resource-num>{doi}</electronic-resource-num>",
                f"  <urls><related-urls><url>https://doi.org/{doi}</url></related-urls></urls>",
            ]
        if abstract:
            lines.append(f"  <abstract>{abstract}</abstract>")
        lines.append("</record>")

    lines += ["</records>", "</xml>"]
    return "\n".join(lines)


# ── 8. Build replacement table ───────────────────────────────────────────────

def build_replacements(patterns, cmap, style="bracket"):
    repls = {}
    for pat in patterns:
        nums = [int(n) for n in re.findall(r"\d+", pat)]
        missing = [n for n in nums if n not in cmap]
        if missing:
            print(f"  WARNING: ref(s) {missing} in pattern {pat} not in reference list")
            continue
        if len(nums) == 1:
            repls[pat] = cmap[nums[0]]["temp_cite"]
        else:
            parts = [f"{cmap[n]['author']}, {cmap[n]['year']} #{cmap[n]['record_num']}"
                     for n in nums]
            repls[pat] = "{" + "; ".join(parts) + "}"
    return repls


def _find_body_start(xml):
    """
    Return the position where actual body text begins — after the title/author
    block. Affiliation superscripts (1, 2, 3) live in the title block and must
    NOT be converted. Citations start at Abstract or Introduction.
    """
    for marker in (
        ">Abstract<", ">ABSTRACT<",
        ">Introduction<", ">INTRODUCTION<",
        ">Background<", ">BACKGROUND<",
        ">Summary<",
    ):
        pos = xml.find(marker)
        if pos > 0:
            # Step back to the start of the paragraph containing this heading
            para_start = xml.rfind("<w:p ", 0, pos)
            return para_start if para_start > 0 else pos
    return 0  # fallback: process whole document (bracket style is safe anyway)


def convert_superscript_xml(xml, cmap):
    """
    Replace superscript citation runs in body XML with plain temp-citation runs.

    Skips everything before Abstract/Introduction (affiliation numbers on the
    title page look identical to citations — must not touch them).
    Iterates run-by-run using indexOf (NOT regex DOTALL) to avoid the greedy
    match problem where .*? crosses paragraph boundaries and eats body text.
    """
    body_start = _find_body_start(xml)
    ref_start  = xml.find(">References<")
    if ref_start < 0:
        ref_start = len(xml)

    prefix = xml[:body_start]          # title page — leave completely alone
    body   = xml[body_start:ref_start] # body text — convert citations here
    tail   = xml[ref_start:]           # reference list — leave alone

    result = []
    pos = 0

    while pos < len(body):
        # Find the next opening run tag
        run_open = body.find("<w:r", pos)
        if run_open < 0:
            result.append(body[pos:])
            break

        # Append everything up to this run
        result.append(body[pos:run_open])

        # Find the matching closing tag (w:r elements are never nested)
        run_close = body.find("</w:r>", run_open)
        if run_close < 0:
            result.append(body[run_open:])
            break
        run_close += len("</w:r>")

        run_xml = body[run_open:run_close]
        pos = run_close

        # Only process runs that are superscript AND contain only digits/commas
        if ('vertAlign w:val="superscript"' in run_xml):
            tm = re.search(r"<w:t[^>]*>([\d,\s]+)</w:t>", run_xml)
            if tm:
                content = tm.group(1).strip()
                nums = [int(n) for n in re.findall(r"\d+", content)]
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


# ── 9. Convert citations in XML (handles all split patterns) ─────────────────

PROOF = r"(?:<w:proofErr[^/]*/>\s*)*"


def _resolve_author(author, cmap):
    """Find record by author last name match (case-insensitive)."""
    al = author.strip().lower()
    for num, info in cmap.items():
        if info["author"].lower() == al:
            return info
    return None


def convert_xml(xml, repls, cmap):
    # Pass 1 — simple complete string replacements (longest first)
    for old, new in sorted(repls.items(), key=lambda x: -len(x[0])):
        xml = xml.replace(old, new)

    # Pass 2 — two-way split across run boundary
    def fix2(m):
        pre, nums_str = m.group(1), m.group(2)
        nums = [int(n) for n in re.findall(r"\d+", nums_str)]
        key = "[" + ",".join(str(n) for n in nums) + "]"
        if key in repls:
            return pre + repls[key]
        return m.group(0)

    xml = re.sub(
        r"(\{[A-Za-zÀ-ž\-]+, )\[(\d+(?:,\d+)*)\]",
        fix2, xml
    )

    # Pass 3 — three-way spell-check split: {  | Author | , year}
    def fix3(m):
        tag, pre, author, tail = m.group(1), m.group(2), m.group(3), m.group(4)
        info = _resolve_author(author, cmap)
        if not info:
            return m.group(0)
        before = pre[:-1]
        cite = f"{{{author}, {info['year']} #{info['record_num']}}}"
        return f"{tag}{before}{cite}</w:t></w:r>"

    xml = re.sub(
        r"(<w:t[^>]*>)([^<]*\{)</w:t></w:r>"
        + PROOF
        + r"<w:r[^>]*><w:t[^>]*>([^<]+)</w:t></w:r>"
        + PROOF
        + r"<w:r[^>]*><w:t[^>]*>(, \d{4}[^<]*)\}</w:t></w:r>",
        fix3, xml, flags=re.DOTALL,
    )

    # Pass 4 — four-way split: { | Author | ,<space> | year}
    def fix4(m):
        run_open, pre, author, year_str = m.group(1), m.group(2), m.group(3), m.group(5)
        info = _resolve_author(author, cmap)
        if not info:
            return m.group(0)
        before = pre[:-1]
        cite = f"{{{author}, {info['year']} #{info['record_num']}}}"
        return f'{run_open}<w:t xml:space="preserve">{before}{cite}</w:t></w:r>'

    xml = re.sub(
        r"(<w:r[^>]*><w:t[^>]*>)([^<]*\{)</w:t></w:r>"
        + PROOF
        + r"<w:r[^>]*><w:t[^>]*>([^<]+)</w:t></w:r>"
        + PROOF
        + r"<w:r[^>]*><w:t[^>]*>(, )</w:t></w:r>"
        + PROOF
        + r"<w:r[^>]*><w:t[^>]*>(\d{4}[^<]*)\}</w:t></w:r>",
        fix4, xml, flags=re.DOTALL,
    )

    return xml


# ── 10. Apply to docx ────────────────────────────────────────────────────────

def apply_to_docx(src, dst, repls, cmap, style="bracket"):
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
                    # Superscript runs are replaced directly in XML structure;
                    # do NOT use repls (bare numbers would corrupt XML attributes)
                    xml = convert_superscript_xml(xml, cmap)
                else:
                    # Bracket style: safe to do string replacement ([N] is unique)
                    xml = convert_xml(xml, repls, cmap)
                data = xml.encode("utf-8")
            zout.writestr(item, data)

    os.remove(tmp)

    # Validate XML
    with zipfile.ZipFile(dst) as z:
        raw = z.read("word/document.xml")
    try:
        ET.fromstring(raw)
    except ET.ParseError as e:
        print(f"\n  ERROR: XML invalid after conversion: {e}")
        return False

    # Check for remaining bare citations
    plain = re.sub(r"<[^>]+>", "", raw.decode("utf-8"))
    bare = [b for b in re.findall(r"\{[A-ZÀ-ž][^}#]{3,60}\}", plain) if "#" not in b]
    if bare:
        print(f"\n  WARNING: {len(bare)} citations still lack record numbers:")
        for b in set(bare):
            print(f"    {b}")
    else:
        print("  All citations have record numbers ✓")

    return True


# ── 11. Main ─────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description="Convert numbered .docx citations to EndNote temp citations")
    ap.add_argument("docx", help="Input .docx with numbered citations")
    ap.add_argument("--output", help="Output .docx path")
    ap.add_argument("--xml",    help="EndNote XML output path")
    ap.add_argument("--email",  default="fakoredesodiq@gmail.com",
                    help="Email for CrossRef polite pool")
    ap.add_argument("--dry-run", action="store_true",
                    help="Show detected refs/patterns only, no conversion")
    args = ap.parse_args()

    src  = Path(args.docx).expanduser().resolve()
    stem = src.stem
    dst  = Path(args.output).resolve() if args.output else src.parent / f"{stem}_converted.docx"
    xmlp = Path(args.xml).resolve()    if args.xml    else src.parent / f"{stem}_endnote.xml"

    print(f"\n{'='*60}")
    print("  Citation Converter v2.0")
    print(f"  Input : {src}")
    print(f"{'='*60}")

    print("\n[1/5] Extracting reference list...")
    refs = extract_references(src)
    no_doi = [r["num"] for r in refs if not r["doi"]]
    print(f"  {len(refs)} references found | {len(no_doi)} without DOI: {no_doi}")

    print("\n[2/5] Detecting citation patterns...")
    patterns, style = find_all_patterns(src)
    combined = [p for p in patterns if "," in p]
    print(f"  Style detected: {style}")
    print(f"  {len(patterns)} unique patterns | {len(combined)} combined: {combined}")

    if args.dry_run:
        print("\n-- Dry run: first 5 references --")
        for r in refs[:5]:
            print(f"  [{r['num']}] DOI={r['doi']}")
            print(f"       {r['text'][:90]}...")
        return 0

    print("\n[3/5] Fetching CrossRef metadata...")
    meta = fetch_all(refs, args.email)
    print(f"  {len(meta)}/{len(refs)} fetched from CrossRef")

    print("\n[4/5] Building citation map...")
    cmap = build_map(refs, meta)
    repls = build_replacements(patterns, cmap, style)
    print(f"  {len(repls)} replacement rules built")

    print("\n[5/5] Generating outputs...")

    # EndNote XML
    xml_content = generate_endnote_xml(refs, meta, cmap)
    xmlp.write_text(xml_content, encoding="utf-8")
    print(f"  EndNote XML  → {xmlp}")

    # Converted docx
    ok = apply_to_docx(src, dst, repls, cmap, style)

    if ok:
        print(f"  Converted doc → {dst}")
        print(f"\n{'='*60}")
        print("  Done!\n")
        print("  Next steps:")
        print(f"  1. In EndNote: File → Import → File")
        print(f"     File: {xmlp}")
        print(f"     Import Option: EndNote XML")
        print(f"     Use a FRESH dedicated library for this paper")
        print(f"  2. Open {dst} in Word")
        print(f"  3. Delete the old reference list at the bottom")
        print(f"  4. EndNote ribbon → Update Citations and Bibliography")
        print(f"{'='*60}\n")
        return 0
    else:
        return 1


if __name__ == "__main__":
    sys.exit(main())
