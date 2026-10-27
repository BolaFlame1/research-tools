#!/usr/bin/env python3
"""Offline checks for the print-year logic in convert.py.

Run after touching print_year / _ref_year / parse_ris:
    python3 ~/research/tools/citation_convert/test_years.py
"""
import sys

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from convert import print_year, _ref_year, parse_ris  # noqa: E402

EPUB_AHEAD = {
    # ref 3 — Cruz-Jentoft: CrossRef PY is the online-first date, Y1 the print date
    "crossref": "TY  - JOUR\nAU  - Cruz-Jentoft, Alfonso J.\nPY  - 2018\nY1  - 2019/01/01\nVL  - 48\nER  - \n",
    "ref": "Cruz-Jentoft AJ, et al. Sarcopenia: revised European consensus on definition and "
           "diagnosis. Age Ageing. 2019;48(1):16–31.",
    "want": 2019,
    # ref 4 — Verghese
    "crossref_v": "TY  - JOUR\nAU  - Verghese, Joe\nPY  - 2012\nY1  - 2013/04/01\nVL  - 68\nER  - \n",
    "ref_v": "Verghese J, et al. Motoric cognitive risk syndrome and the risk of dementia. "
             "J Gerontol A Biol Sci Med Sci. 2013;68(4):412–418.",
    "want_v": 2013,
    # ref 12 — Krumpoch
    "crossref_k": "TY  - JOUR\nAU  - Krumpoch, Sebastian\nPY  - 2020\nY1  - 2021/01/01\nVL  - 33\nER  - \n",
    "ref_k": "Krumpoch S, et al. The effect of different test protocols and walking distances "
             "on gait speed in older persons. Aging Clin Exp Res. 2021;33(1):141–146.",
    "want_k": 2021,
}

# PY later than Y1 — the "pick the earliest year" rule gets these wrong.
# PY here is the *updated/online* year and Y1 the print year.
EPUB_LATER = [
    ("TY  - JOUR\nAU  - Witham, Miles D.\nPY  - 2020\nY1  - 2019/01/01\nVL  - 4\nER  - \n", 2019),
    ("TY  - JOUR\nAU  - X, Y.\nY1  - 2019/05/01\nPY  - 2021\nVL  - 9\nER  - \n", 2019),
]

TITLE_TRAPS = [
    # The only years in these titles are a study span and a guideline name.
    ("Gouda HN, et al. Burden of non-communicable diseases in sub-Saharan Africa, 1990–2017: "
     "results from the Global Burden of Disease Study 2017. Lancet Glob Health. 2019;7(10):e1375–e1387.", 2019),
    ("Page MJ, et al. The PRISMA 2020 statement: an updated guideline for reporting systematic "
     "reviews. BMJ. 2021;372:n71.", 2021),
    ("Stover ME, et al. Slow gait speed in older US adults. Geriatrics. 2023;8(5):95.", 2023),
]

fails = []

py = print_year(EPUB_AHEAD["crossref"])
if py != EPUB_AHEAD["want"]:
    fails.append(f"print_year(Cruz-Jentoft) = {py}, want {EPUB_AHEAD['want']}")
py_v = print_year(EPUB_AHEAD["crossref_v"])
if py_v != EPUB_AHEAD["want_v"]:
    fails.append(f"print_year(Verghese) = {py_v}, want {EPUB_AHEAD['want_v']}")
py_k = print_year(EPUB_AHEAD["crossref_k"])
if py_k != EPUB_AHEAD["want_k"]:
    fails.append(f"print_year(Krumpoch) = {py_k}, want {EPUB_AHEAD['want_k']}")

_, y = parse_ris(EPUB_AHEAD["crossref"])
if y != EPUB_AHEAD["want"]:
    fails.append(f"parse_ris(Cruz-Jentoft) year = {y}, want {EPUB_AHEAD['want']} "
                 f"— temp citation would carry the epub year")

for ref, want in TITLE_TRAPS:
    got = _ref_year(ref)
    if got != want:
        fails.append(f"_ref_year = {got}, want {want} for: {ref[:60]}")

for ris, want in EPUB_LATER:
    got = print_year(ris)
    if got != want:
        fails.append(f"print_year(PY later than Y1) = {got}, want {want}")

CASES = 8
print(f"print_year / parse_ris / _ref_year: {CASES - len(fails)}/{CASES} passed")
for f in fails:
    print("FAIL:", f)
sys.exit(1 if fails else 0)
