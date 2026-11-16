"""
Topic Volume Filler  -  Streamlit tool
--------------------------------------
Upload the master workbook + the zip of SEMrush topic CSVs, pick a country tab
(UK, Brazil, Thailand ...) and the tool fills that tab's "Topic Volume" column
using exact match first, then near/phrase match. The next column(s) show the
reference topic that was used and the match type.

* Non-English tabs: every row has an English topic AND a local-language topic.
  The tool tries both against the CSVs (file names and keywords), so it works
  whether SEMrush exported the topics in English or in the local language.
* Nothing matched? The closest keyword is suggested (orange, unverified) and
  every non-exact cell gets an Excel comment explaining what was used.
* Dog and Cat rows are never mixed (English + local-language animal words).

Run:   pip install streamlit openpyxl pandas
       streamlit run main.py
"""
import csv
import io
import re
import unicodedata
import zipfile
from collections import defaultdict
from difflib import SequenceMatcher
from functools import lru_cache

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# Constants
# ----------------------------------------------------------------------------
HEADER_SCAN_ROWS = 15                      # header row is auto-detected in the first N rows
ZW = dict.fromkeys(map(ord, "​‌‍﻿ ⁠"), " ")
SUGGEST_MIN_SCORE = 0.45                   # min similarity for a suggestion from the topic's own CSV
SUGGEST_MIN_OTHER = 0.80                   # stricter when the keyword comes from a different topic's CSV

# connector words ignored in the "loose" comparison (EN / ES / PT / DE / ID / TL ...)
STOP = set("""and the a an of for with in on to or y e o u und oder dan atau at ng sa na ang
de da do das dos del la las los el para por con com em um uma der die das fur en et di yang untuk""".split())

# animal words used to keep Dog and Cat rows apart (English + local languages)
_DOG_RAW = """dog puppy pup canine doggo doggy pooch hound k9 pupper cachorro cao cae filhote perro perrito
cachorrito hund hunde welpe anjing aso tuta cane cucciolo สุนัข หมา น้องหมา ลูกหมา"""
_CAT_RAW = """cat kitten kitty feline tabby gato gatinho gatito felino katze kater katzchen kucing pusa kuting
gatto gattino แมว ลูกแมว"""

# competitor / other pet brands to exclude from matching (only Pedigree & Whiskas allowed)
EXCLUDED_PET_BRANDS = [
    # Top multinational & competitor brands
    "royal canin", "purina one", "purina", "pro plan", "fancy feast", "friskies",
    "science plan", "science diet", "prescription diet", "hills", "hill's",
    "iams", "advance", "cesar", "sheba", "nutro", "greenies", "temptations", "dreamies",
    "eukanuba", "blue buffalo", "taste of the wild", "lily's kitchen", "lilys kitchen", "lilys",
    "james wellbeloved", "edgard & cooper", "edgard cooper", "pooch & mutt", "pooch and mutt",
    "nature's menu", "natures menu", "nature's variety", "natures variety",
    "barking heads", "meowing heads", "stella & chewy's", "stella and chewys",
    "diamond naturals", "tiki cat", "tiki dog", "ziwi peak", "ziwi", "almo nature",
    "happy dog", "happy cat", "gourmet", "felix", "bakers", "dentalife", "beneful",
    "acana", "orijen", "merrick", "applaws", "harringtons", "chappie", "winalot",
    "butchers", "butcher's", "forthglade", "scrumbles", "katkin", "untamed", "oxbow",
    "hownd", "skinners", "skinner's", "arden grange", "canagan", "symply", "crave",
    "burgess", "vitalin", "chudleys", "autarky", "supadog", "smilla", "naturo", "carnilove",
    "canidae", "weruva", "nulo", "schesir", "animonda", "bozita", "leonardo", "belcando",
    "josera", "mera", "kitekat", "perfect fit", "encore", "mon petit", "go-cat", "feringa",
    "trovet", "mjamjam", "cheshire cat", "fuzzball", "marro", "dr john", "csj", "serious good",
    "frontline", "advantix", "seresto", "bravecto", "nexgard", "drontal", "milbemax",
    # Compound phrases to avoid false positives on common English words
    "k9 dog food", "k9 natural", "k9 pet", "k9 dog", "k9 puppy",
    "wellness dog", "wellness cat", "wellness pet", "wellness core",
    "burns dog", "burns cat", "burns pet"
]
BRAND_RE = re.compile(r'\b(' + '|'.join(re.escape(b) for b in EXCLUDED_PET_BRANDS) + r')\b', re.IGNORECASE)
ALLOWED_BRANDS = {"pedigree", "whiskas", "dentastix"}


def brand_ok(kw_tokens: tuple, kw_name: str, topic_tokens: tuple, species: str) -> bool:
    """Return False if the keyword mentions a competitor pet brand (only Pedigree & Whiskas allowed)."""
    kw_str = str(kw_name or " ".join(kw_tokens)).lower()
    m = BRAND_RE.search(kw_str)
    if m:
        matched = m.group(0).lower()
        if matched not in ALLOWED_BRANDS:
            return False
    return True

COLOURS = {            # cell fill per match type
    "Exact": "C6EFCE",             # green
    "Phrase": "FFEB9C",            # amber
    "Other file": "D9D9D9",        # grey
    "Suggested": "F8CBAD",         # orange  (closest topic, unverified)
    "Not found": "FFC7CE",         # red
}
HEAD_FILL = PatternFill("solid", fgColor="D9E1F2")
RANK = {"Exact": 0, "Phrase": 1, "Exact (other file)": 2, "Phrase (other file)": 3}


# ----------------------------------------------------------------------------
# Text normalisation / matching helpers
# ----------------------------------------------------------------------------
def _stem(w: str) -> str:
    if len(w) > 4 and w.endswith("ies"):
        return w[:-3] + "y"
    if len(w) > 3 and w.endswith("s") and not w.endswith("ss"):
        return w[:-1]
    return w


def _keep(ch: str) -> bool:
    return ch.isspace() or unicodedata.category(ch)[0] in "LNM"     # letters, digits, combining marks (Thai vowels)


@lru_cache(maxsize=200000)
def toks(s) -> tuple:
    """Lower-case, accent-free (Latin only), punctuation-free, plural-insensitive token tuple."""
    s = str(s or "").split("\n")[0].translate(ZW)
    s = unicodedata.normalize("NFKD", s)
    if all(ord(c) < 0x250 or unicodedata.combining(c) for c in s):   # strip accents for Latin scripts only
        s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower().replace("&", " and ")
    s = "".join(ch if _keep(ch) else " " for ch in s)
    return tuple(_stem(w) for w in s.split())


@lru_cache(maxsize=200000)
def loose(t: tuple) -> tuple:
    """Tokens without connector words ('and', 'y', 'und', 'dan' ...)."""
    return tuple(w for w in t if w not in STOP) or t


def contains(hay: tuple, needle: tuple) -> bool:
    if not needle or not hay:
        return False
    if len(needle) == 1 and len(hay) == 1:          # unspaced scripts (Thai, CJK)
        return needle[0] in hay[0]
    n = len(needle)
    return any(hay[i:i + n] == needle for i in range(len(hay) - n + 1))


def contains_loose(hay: tuple, needle: tuple) -> bool:
    return contains(hay, needle) or contains(loose(hay), loose(needle))


DOG_WORDS = {toks(w)[0] for w in _DOG_RAW.split() if toks(w)}
CAT_WORDS = {toks(w)[0] for w in _CAT_RAW.split() if toks(w)}


def _hit(kw: tuple, words: set) -> bool:
    for t in kw:
        for w in words:
            if t == w or (len(w) >= 5 and t.startswith(w)) or (not w.isascii() and w in t):
                return True
    return False


def species_ok(kw: tuple, topic: tuple, species) -> bool:
    """False if the keyword talks about the *other* animal while the topic does not."""
    if species not in ("Dog", "Cat"):
        return True
    opp = CAT_WORDS if species == "Dog" else DOG_WORDS
    return not _hit(kw, opp) or _hit(topic, opp)


# ----------------------------------------------------------------------------
# Reading the zip / CSV files
# ----------------------------------------------------------------------------
def _fix_name(info: zipfile.ZipInfo) -> str:
    if info.flag_bits & 0x800:
        return info.filename
    try:
        return info.filename.encode("cp437").decode("utf-8")
    except Exception:
        return info.filename


def parse_csv_bytes(raw: bytes):
    """Return list of (tokens, name, volume) from a SEMrush topic CSV."""
    text = None
    for enc in ("utf-8-sig", "utf-16", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeError:
            continue
    if not text or not text.strip():
        return []
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t")
    except csv.Error:
        dialect = csv.excel
    rd = csv.DictReader(io.StringIO(text), dialect=dialect)
    if not rd.fieldnames:
        return []
    rd.fieldnames = [(h or "").strip().lower() for h in rd.fieldnames]
    name_col = next((h for h in rd.fieldnames if h in ("name", "topic", "keyword", "phrase")), rd.fieldnames[0])
    vol_col = next((h for h in rd.fieldnames if "volume" in h), None)
    if vol_col is None:
        return []
    rows = []
    for r in rd:
        nm = (r.get(name_col) or "").strip()
        try:
            v = int(float(str(r.get(vol_col, "")).replace(",", "").strip()))
        except ValueError:
            continue
        if nm:
            rows.append((toks(nm), nm, v))
    return rows


def _walk_zip(zf: zipfile.ZipFile, out: list, prefix: str = "", zip_name: str = ""):
    for info in zf.infolist():
        if info.is_dir() or "__MACOSX" in info.filename:
            continue
        name = _fix_name(info)
        low = name.lower()
        if low.endswith(".csv"):
            out.append((name.split("/")[-1], zf.read(info), zip_name))
        elif low.endswith(".zip"):                       # nested zip
            try:
                _walk_zip(zipfile.ZipFile(io.BytesIO(zf.read(info))), out, zip_name=zip_name)
            except zipfile.BadZipFile:
                pass


@st.cache_data(show_spinner=False)
def load_topic_files(blobs: tuple):
    """blobs = ((filename, bytes), ...) -> {file_stem: (filename, rows, zip_name)}"""
    found = []
    for fname, data in blobs:
        if fname.lower().endswith(".zip"):
            _walk_zip(zipfile.ZipFile(io.BytesIO(data)), found, zip_name=fname)
        elif fname.lower().endswith(".csv"):
            found.append((fname, data, fname))
    files = {}
    for fname, data, zname in found:
        stem = re.sub(r"\.csv$", "", fname, flags=re.I)
        rows = parse_csv_bytes(data)
        files[toks(stem)] = (fname, rows, zname)
    return files


# ----------------------------------------------------------------------------
# Master workbook helpers
# ----------------------------------------------------------------------------
def find_header(ws):
    """Return (header_row, {role: col}) for a country tab, or None for a global tab."""
    for r in range(1, min(ws.max_row, HEADER_SCAN_ROWS) + 1):
        heads = {c: str(ws.cell(r, c).value or "").strip() for c in range(1, ws.max_column + 1)}
        if any(h.startswith("Topic (EN)") for h in heads.values()) and \
           any(h.lower().startswith("topic volume") for h in heads.values()):
            cols = {"vol": [], "cat": None, "topic_en": None, "topic_loc": None, "prompt": None,
                    "priority_brand": None, "brand_name": None}
            for c, h in heads.items():
                hl = h.lower()
                if hl.startswith("topic volume"):
                    cols["vol"].append((c, h))
                elif hl.startswith("category"):
                    cols["cat"] = c
                elif h.startswith("Topic (EN)"):
                    cols["topic_en"] = c
                elif h.startswith("Topic (Local"):
                    cols["topic_loc"] = c
                elif hl.startswith("prompt (en)"):
                    cols["prompt"] = c
                elif hl.startswith("priority brand"):
                    cols["priority_brand"] = c
                elif hl.startswith("brand name") or hl == "brand":
                    cols["brand_name"] = c
            return r, cols
    return None


def detect_country_tabs(wb) -> dict:
    tabs = {}
    for name in wb.sheetnames:
        h = find_header(wb[name])
        if h:
            tabs[name] = h
    return tabs


# ----------------------------------------------------------------------------
# Core lookup
# ----------------------------------------------------------------------------
def _find_own(files: dict, tt: tuple):
    own = files.get(tt)
    if own is None:
        lt = loose(tt)
        own = next((fv for fk, fv in files.items() if loose(fk) == lt), None)
    if own is None:
        own = next((fv for fk, fv in files.items()
                    if contains(fk, tt) or (len(fk) > 1 and contains(tt, fk))), None)
    return own


def lookup_one(files, all_rows, exact_idx, topic_text, species, opts):
    """Exact -> Phrase (own CSV) -> Exact/Phrase (other CSVs) for ONE topic text."""
    tt = toks(topic_text)
    lt = loose(tt)
    res = {"type": "Not found", "ref": "", "volume": None, "file": "", "zip_file": "", "alts": [], "rejected": 0, "score": None}
    if not tt:
        return res

    def ok(kw, nm=""):
        if opts.get("species_guard") and not species_ok(kw, tt, species):
            res["rejected"] += 1
            return False
        if opts.get("brand_guard") and not brand_ok(kw, nm, tt, species):
            res["rejected"] += 1
            return False
        return True

    def phrase(rows):
        fwd = [(nm, v) for tk, nm, v in rows if contains_loose(tk, tt) and ok(tk, nm)]
        rev = []
        if opts.get("reverse_phrase"):
            rev = [(nm, v, len(tk)) for tk, nm, v in rows
                   if len(tk) >= 2 and contains_loose(tt, tk) and ok(tk, nm)]
        if fwd:
            fwd.sort(key=lambda x: -x[1])
            return fwd[0][0], fwd[0][1], [f"{n} ({v:,})" for n, v in fwd[1:5]]
        if rev:
            rev.sort(key=lambda x: (-x[2], -x[1]))
            return rev[0][0], rev[0][1], [f"{n} ({v:,})" for n, v, _ in rev[1:5]]
        return None

    own = _find_own(files, tt)
    if own:
        fname, rows, zname = own
        ex = [(nm, v) for tk, nm, v in rows if (tk == tt or loose(tk) == lt) and ok(tk, nm)]
        if ex:
            nm, v = max(ex, key=lambda x: x[1])
            res.update(type="Exact", ref=nm, volume=v, file=fname, zip_file=zname)
            return res
        ph = phrase(rows)
        if ph:
            res.update(type="Phrase", ref=ph[0], volume=ph[1], file=fname, zip_file=zname, alts=ph[2])
            return res

    if opts.get("other_files"):
        e = exact_idx.get(lt)
        if e and ok(toks(e[0]), e[0]):
            res.update(type="Exact (other file)", ref=e[0], volume=e[1], file=e[2], zip_file=e[3])
            return res
        cand = [(nm, v, f, zn) for tk, nm, v, f, zn in all_rows if contains_loose(tk, tt) and ok(tk, nm)]
        if cand:
            cand.sort(key=lambda x: -x[1])
            nm, v, f, zn = cand[0]
            res.update(type="Phrase (other file)", ref=nm, volume=v, file=f, zip_file=zn,
                       alts=[f"{n} ({vv:,})" for n, vv, _, _ in cand[1:5]])
            return res
    return res


def _sim(a: tuple, b: tuple) -> float:
    A, B = set(loose(a)), set(loose(b))
    jac = len(A & B) / len(A | B) if (A | B) else 0.0
    ratio = SequenceMatcher(None, " ".join(a), " ".join(b)).ratio()
    if len(a) == 1 or len(b) == 1:                    # unspaced scripts: character similarity only
        return ratio if ratio >= 0.6 else 0.0
    return 0.5 * jac + 0.5 * ratio if jac > 0 else 0.0


def suggest_closest(files, all_rows, variants, species, opts):
    """Best-effort nearest keyword when no exact/phrase match exists."""
    best = None
    for label, text in variants:
        tt = toks(text)
        if not tt:
            continue
        own = _find_own(files, tt)
        pools = [[(tk, nm, v, own[0], own[2]) for tk, nm, v in own[1]]] if own else []
        pools.append(all_rows)
        for pool in pools:
            own_pool = bool(own) and pool is pools[0]
            for tk, nm, v, f, zn in pool:
                if opts.get("species_guard") and not species_ok(tk, tt, species):
                    continue
                if opts.get("brand_guard") and not brand_ok(tk, nm, tt, species):
                    continue
                if len(tt) > 1 and len(tk) > 1 and not (set(loose(tt)) & set(loose(tk))):
                    continue                                  # cheap pre-filter before the slower ratio
                sc = _sim(tt, tk)
                thr = SUGGEST_MIN_SCORE if own_pool else SUGGEST_MIN_OTHER
                if not own_pool and (len(tt) == 1 or len(tk) == 1):
                    thr = 0.8                                  # unspaced scripts need a much closer string
                if sc >= thr and (best is None or sc > best[0]):
                    best = (sc, nm, v, f, zn, label)
            if best and own_pool:
                break                                         # a good candidate inside the topic's own CSV wins
    return best


def find(files, all_rows, exact_idx, variants, species, opts):
    """Try every topic text (EN + local) and keep the strongest match tier."""
    best, rejected = None, 0
    for label, text in variants:
        r = lookup_one(files, all_rows, exact_idx, text, species, opts)
        rejected += r["rejected"]
        r["via"] = label
        if r["type"] != "Not found" and (best is None or RANK[r["type"]] < RANK[best["type"]]):
            best = r
    if best:
        best["rejected"] = rejected
        return best
    res = {"type": "Not found", "ref": "", "volume": None, "file": "", "zip_file": "", "alts": [],
           "rejected": rejected, "score": None, "via": ""}
    if opts.get("suggest"):
        sg = suggest_closest(files, all_rows, variants, species, opts)
        if sg:
            sc, nm, v, f, zn, label = sg
            res.update(type="Suggested (closest)", ref=nm, volume=v, file=f, zip_file=zn, score=round(sc, 2), via=label)
    return res


def style_key(match_type: str) -> str:
    t = str(match_type)
    if "other file" in t:
        return "Other file"
    for k in ("Exact", "Phrase", "Suggested"):
        if t.startswith(k):
            return k
    return "Not found"


def make_comment(res, variants, n_files) -> str | None:
    topics = " / ".join(f"{lab}: '{txt}'" for lab, txt in variants)
    kind = style_key(res["type"])
    if kind == "Exact":
        return None
    if kind == "Not found":
        extra = " Brand/Animal filter rejected %d keyword(s)." % res["rejected"] if res["rejected"] else ""
        return (f"NO MATCH. Topic {topics}. Searched {n_files} CSV(s): no exact, phrase or close keyword."
                f"{extra} Left blank - please check manually.")
    head = {"Phrase": "PHRASE MATCH", "Other file": "MATCH IN ANOTHER CSV", "Suggested": "SUGGESTED (unverified)"}[kind]
    src_info = f"File: {res['file']}"
    if res.get("zip_file") and res["zip_file"] != res["file"]:
        src_info += f" (Zip: {res['zip_file']})"
    txt = f"{head}. Topic {topics}. Used keyword '{res['ref']}' = {res['volume']:,} ({src_info})."
    if kind == "Suggested":
        txt += f" No exact/phrase match; closest by similarity ({res['score']}). Please verify."
    if res["via"] and len(variants) > 1:
        txt += f" Matched using the {res['via']} topic."
    if res["alts"]:
        txt += " Other candidates: " + "; ".join(res["alts"][:3])
    return txt


# ----------------------------------------------------------------------------
# Fill the workbook
# ----------------------------------------------------------------------------
def fill_workbook(master_bytes: bytes, sheet: str, files: dict, opts: dict):
    wb = openpyxl.load_workbook(io.BytesIO(master_bytes))
    ws = wb[sheet]
    hdr_row, cols = find_header(ws)

    all_rows, exact_idx = [], {}
    for fname, rows, zname in files.values():
        for tk, nm, v in rows:
            all_rows.append((tk, nm, v, fname, zname))
            e = exact_idx.get(loose(tk))
            if e is None or v > e[1]:
                exact_idx[loose(tk)] = (nm, v, fname, zname)

    def is_generated_or_free(c):
        if c > ws.max_column:
            return True
        val = str(ws.cell(hdr_row, c).value or "").lower()
        if any(k in val for k in ("reference topic", "match type", "zip file", "source file")):
            return True
        return all(ws.cell(r, c).value in (None, "") for r in range(1, ws.max_row + 1))

    next_free = ws.max_column + 1
    targets = []        # (vol_col, primary_label, ref_col, type_col, zip_col)
    for vcol, vhead in cols["vol"]:
        primary = "Local" if ("local" in vhead.lower() and cols["topic_loc"]) else "EN"
        if is_generated_or_free(vcol + 1) and is_generated_or_free(vcol + 2) and is_generated_or_free(vcol + 3) and len(cols["vol"]) == 1:
            rcol, ycol, zcol = vcol + 1, vcol + 2, vcol + 3
        else:
            rcol, ycol, zcol = next_free, next_free + 1, next_free + 2
            next_free += 3
        suffix = f" ({primary})" if len(cols["vol"]) > 1 else ""
        for c, txt in ((rcol, f"Reference Topic (if no exact match){suffix}"),
                       (ycol, f"Match Type{suffix}"),
                       (zcol, f"Source File Name{suffix}")):
            cell = ws.cell(hdr_row, c, txt)
            cell.font = Font(bold=True)
            cell.fill = HEAD_FILL
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            ws.column_dimensions[get_column_letter(c)].width = 38 if c == rcol else (35 if c == zcol else 24)
        targets.append((vcol, primary, rcol, ycol, zcol))

    clean = lambda x: str(x).strip().split("\n")[0].replace("​", "").strip() if x else ""
    cache, log, rows_out = {}, [], []
    excluded_rows = []

    for r in range(hdr_row + 1, ws.max_row + 1):
        pb = str(ws.cell(r, cols["priority_brand"]).value or "").strip() if cols.get("priority_brand") else ""
        bn = str(ws.cell(r, cols["brand_name"]).value or "").strip() if cols.get("brand_name") else ""

        is_other_brand = False
        if opts.get("filter_other_brand_rows"):
            if pb.lower() == "others" or (bn and bn.lower() not in ("pedigree", "whiskas")):
                is_other_brand = True

        if is_other_brand:
            excluded_rows.append(r)
            for vcol, primary, rcol, ycol, zcol in targets:
                for c in (vcol, rcol, ycol, zcol):
                    ws.cell(r, c).value = None
            continue

        species = ws.cell(r, cols["cat"]).value if cols["cat"] else None
        species = str(species).strip().title() if species else None
        en = clean(ws.cell(r, cols["topic_en"]).value)
        loc = clean(ws.cell(r, cols["topic_loc"]).value) if cols["topic_loc"] else ""
        if not (en or loc):
            continue
        for vcol, primary, rcol, ycol, zcol in targets:
            variants = [("EN", en), ("Local", loc)] if primary == "EN" else [("Local", loc), ("EN", en)]
            variants = [(lab, t) for lab, t in variants if t]
            if len(variants) == 2 and toks(variants[0][1]) == toks(variants[1][1]):
                variants = variants[:1]                       # same text in both languages (UK, Philippines EN)
            key = (species, primary, tuple(t for _, t in variants))
            if key not in cache:
                res = find(files, all_rows, exact_idx, variants, species, opts)
                cache[key] = (res, variants)
                log.append({"Species": species or "", "Column": primary,
                            "Topic (EN)": en, "Topic (Local)": loc if loc != en else "",
                            "Match Type": res["type"], "Reference Topic": res["ref"], "Volume": res["volume"],
                            "Matched via": res["via"] if len(variants) > 1 else "",
                            "Similarity": res["score"],
                            "Source File Name": res["file"],
                            "Zip File": res["zip_file"],
                            "Other candidates": " | ".join(res["alts"]),
                            "Rejected (brand/animal)": res["rejected"]})
            res, variants = cache[key]
            kind = style_key(res["type"])
            vcell, rcell, ycell, zcell = ws.cell(r, vcol), ws.cell(r, rcol), ws.cell(r, ycol), ws.cell(r, zcol)
            vcell.value = res["volume"]
            vcell.number_format = "#,##0"
            if kind == "Exact":
                rcell.value = ""
                zcell.value = res["file"]
            elif kind == "Suggested":
                rcell.value = f"≈ {res['ref']} (closest, unverified)"
                zcell.value = res["file"]
            elif kind == "Not found":
                rcell.value = "No matching topic found"
                zcell.value = ""
            else:
                rcell.value = res["ref"] or "No matching topic found"
                zcell.value = res["file"]
            ycell.value = res["type"]
            for c in (vcell, rcell, ycell, zcell):
                c.fill = PatternFill("solid", fgColor=COLOURS[kind])
            note = make_comment(res, variants, len(files))
            if note:
                cm = Comment(note, "Topic Volume Tool")
                cm.width, cm.height = 380, 170
                vcell.comment = cm
            rows_out.append({"Row": r, "Species": species or "", "Topic (EN)": en, "Topic (Local)": loc,
                             "Volume": res["volume"], "Reference Topic": rcell.value,
                             "Match Type": res["type"], "Source File Name": res["file"], "Column": primary})

    if opts.get("remove_other_brand_rows") and excluded_rows:
        for r in sorted(excluded_rows, reverse=True):
            ws.delete_rows(r, 1)

    log_name = f"Match Log - {sheet}"[:31]
    if log_name in wb.sheetnames:
        del wb[log_name]
    ml = wb.create_sheet(log_name)
    log_df = pd.DataFrame(log)
    ml.append(list(log_df.columns))
    for c in ml[1]:
        c.font = Font(bold=True)
        c.fill = HEAD_FILL
    for rec in log_df.itertuples(index=False):
        ml.append([None if (isinstance(x, float) and pd.isna(x)) else x for x in rec])
    for col_idx, col_name in enumerate(log_df.columns, 1):
        letter = get_column_letter(col_idx)
        w = 32 if "zip" in col_name.lower() else (44 if "topic" in col_name.lower() or "csv" in col_name.lower() or "candidate" in col_name.lower() else 18)
        ml.column_dimensions[letter].width = w
    ml.freeze_panes = "A2"
    ml.auto_filter.ref = ml.dimensions

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), pd.DataFrame(rows_out), log_df


# ----------------------------------------------------------------------------
# Streamlit UI
# ----------------------------------------------------------------------------
def master_alignment_warnings(ws, hdr_row, cols) -> pd.DataFrame:
    """Topics that appear under both Dog and Cat in the master, or whose name mentions the other animal."""
    if not cols["cat"] or not cols["topic_en"]:
        return pd.DataFrame()
    seen = defaultdict(set)
    for r in range(hdr_row + 1, ws.max_row + 1):
        t, c = ws.cell(r, cols["topic_en"]).value, ws.cell(r, cols["cat"]).value
        if t and c:
            seen[str(t).strip().replace("​", "")].add(str(c).strip().title())
    out = []
    for t, cats in seen.items():
        words = set(toks(t))
        if len(cats) > 1:
            out.append({"Topic": t, "Issue": "Used by both Dog and Cat rows (matched separately per species)"})
        elif "Dog" in cats and words & CAT_WORDS and not words & DOG_WORDS:
            out.append({"Topic": t, "Issue": "Dog row but topic name mentions cats"})
        elif "Cat" in cats and words & DOG_WORDS and not words & CAT_WORDS:
            out.append({"Topic": t, "Issue": "Cat row but topic name mentions dogs"})
    return pd.DataFrame(out)


def main():
    st.set_page_config(page_title="Topic Volume Filler", page_icon="📊", layout="wide")
    st.title("📊 Topic Volume Filler")
    st.caption("Upload the master sheet and the SEMrush topic dump (zip of CSVs), pick a country, download the filled sheet.")

    with st.sidebar:
        st.header("Matching options")
        brand_guard = st.checkbox(
            "Exclude competitor branded topics/keywords",
            True,
            help="Rejects CSV keywords/topics mentioning other brands (Royal Canin, Purina, IAMS, Hills, Gourmet, Sheba, Cesar, etc.). Only Pedigree and Whiskas branded topics are allowed."
        )
        filter_other_brand_rows = st.checkbox(
            "Only process Pedigree & Whiskas rows (exclude 'Others' rows)",
            True,
            help="Skip rows from the master sheet where Priority Brand is 'Others' or Brand Name is not Pedigree/Whiskas (e.g. Advance, IAMS, Nutro, Greenies, Sheba, Cesar)."
        )
        remove_other_brand_rows = st.checkbox(
            "Remove excluded brand rows from filled Excel sheet",
            True,
            help="Completely deletes rows for other brands from the downloaded workbook and preview. Untick to leave them in the sheet with blank values."
        )
        species_guard = st.checkbox("Keep Dog and Cat separate", True,
                                    help="A Dog row never takes a keyword that mentions cats (and the reverse) unless the topic itself does.")
        reverse_phrase = st.checkbox("Allow reverse phrase match", True,
                                     help="Also accept a CSV keyword (2+ words) that sits inside the topic name.")
        other_files = st.checkbox("Search other CSVs if own CSV has no match", True,
                                  help="Marked grey so you can review them.")
        suggest = st.checkbox("Suggest closest topic when nothing matches", True,
                              help="Fills the nearest keyword (orange) with an Excel comment so you can verify it. Untick to leave those blank.")
        st.markdown("**Colour key**\n\n🟩 Exact  ·  🟨 Phrase  ·  ⬜ Other CSV  ·  🟧 Suggested (verify)  ·  🟥 Not found")
        st.caption("Non-English tabs: both the English and the local-language topic are tried automatically.")

    c1, c2 = st.columns(2)
    master = c1.file_uploader("1 · Master sheet (.xlsx)", type=["xlsx"])
    dumps = c2.file_uploader("2 · Topic dump (.zip of CSVs, or loose CSVs)", type=["zip", "csv"],
                             accept_multiple_files=True)
    if not master:
        st.info("Upload the master sheet to continue.")
        return

    master_bytes = master.getvalue()
    wb = openpyxl.load_workbook(io.BytesIO(master_bytes), read_only=False)
    tabs = detect_country_tabs(wb)
    if not tabs:
        st.error("No country tabs found. A country tab needs 'Topic (EN)' and 'Topic Volume ...' headers.")
        return

    sheet = st.selectbox("3 · Country / tab to fill", list(tabs), help="Only country tabs are listed; the global Dog/Cat tabs are never touched.")
    hdr_row, cols = tabs[sheet]
    ws = wb[sheet]
    n_rows = sum(1 for r in range(hdr_row + 1, ws.max_row + 1) if ws.cell(r, cols["topic_en"]).value)
    st.write(f"**{sheet}**: {n_rows} prompt rows · volume column(s): "
             + ", ".join(f"{get_column_letter(c)} ({h})" for c, h in cols["vol"]))

    warn = master_alignment_warnings(ws, hdr_row, cols)
    if len(warn):
        with st.expander(f"⚠️ Dog/Cat alignment notes in this tab ({len(warn)})"):
            st.dataframe(warn, use_container_width=True, hide_index=True)

    if not dumps:
        st.info("Upload the topic dump zip to continue.")
        return

    with st.spinner("Reading topic CSVs..."):
        files = load_topic_files(tuple((d.name, d.getvalue()) for d in dumps))
    if not files:
        st.error("No CSV files found in the upload.")
        return
    st.success(f"Loaded {len(files)} topic CSV files · {sum(len(v[1]) for v in files.values()):,} keywords")

    if st.button("▶ Fill topic volumes", type="primary"):
        opts = {
            "species_guard": species_guard,
            "brand_guard": brand_guard,
            "filter_other_brand_rows": filter_other_brand_rows,
            "remove_other_brand_rows": remove_other_brand_rows,
            "reverse_phrase": reverse_phrase,
            "other_files": other_files,
            "suggest": suggest
        }
        with st.spinner("Matching topics..."):
            st.session_state["result"] = (sheet,) + fill_workbook(master_bytes, sheet, files, opts)

    if "result" not in st.session_state:
        return
    sheet, out_bytes, preview, log = st.session_state["result"]
    kinds = log["Match Type"].map(lambda t: style_key(t))
    rows_kind = preview["Match Type"].map(lambda t: style_key(t))
    m = st.columns(6)
    m[0].metric("Unique topics", len(log))
    m[1].metric("Exact", int((kinds == "Exact").sum()))
    m[2].metric("Phrase", int((kinds == "Phrase").sum()))
    m[3].metric("Other CSV", int((kinds == "Other file").sum()))
    m[4].metric("Suggested", int((kinds == "Suggested").sum()))
    m[5].metric("Not found", int((kinds == "Not found").sum()))
    st.caption(f"Rows filled: {int((rows_kind != 'Not found').sum())} of {len(preview)} (suggested values are unverified - see cell comments)")

    st.download_button("⬇ Download filled master sheet", out_bytes,
                       file_name=f"Prompts_Master_List_{sheet.replace(' ', '_')}_FILLED.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")

    t1, t2, t3 = st.tabs(["Topic match log", "Not found", "Row preview"])
    with t1:
        st.dataframe(log, use_container_width=True, hide_index=True)
    with t2:
        nf = log[kinds == "Not found"]
        if len(nf):
            st.dataframe(nf[["Species", "Topic (EN)", "Topic (Local)"]], use_container_width=True, hide_index=True)
        else:
            st.success("Every topic matched.")
    with t3:
        st.dataframe(preview, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()