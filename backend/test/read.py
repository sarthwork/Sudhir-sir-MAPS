#!/usr/bin/env python3
"""
extract_comparables.py
======================
PDF -> table extraction -> LLM *proposes* column meanings -> Python *validates*
-> deterministic normalisation -> one Excel workbook (with audit trail).

Setup
-----
    pip install pdfplumber pandas openpyxl anthropic
    export ANTHROPIC_API_KEY=...        (Windows: setx ANTHROPIC_API_KEY ...)

Usage
-----
    # Step 1 (no LLM, no cost): check that table extraction works on your PDFs
    python extract_comparables.py --extract-only

    # Step 2: full run
    python extract_comparables.py

    Put PDFs in ./pdfs  ->  output: normalized_rental_comparables.xlsx

Workbook sheets
---------------
    Normalized : canonical dataset (+ review flag, raw values, extra fields)
    Review     : everything ambiguous / rejected / failed
    Mappings   : per column: LLM proposal, validation result, confidence
    Raw        : every extracted cell with page / table / row / column
    Tables     : every table found, with the target-detection score

Principle: the LLM is an interpreter, not an authority. It never writes data
values; it only labels columns. Every label is checked against the real values.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import os
import re
import sys
from dataclasses import dataclass, field as dc_field
from datetime import date, datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import pdfplumber

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------
PDF_DIR = Path("./")
OUT_XLSX = Path("normalized_rental_comparables.xlsx")
CACHE_FILE = Path(".llm_mapping_cache.json")
MODEL = os.environ.get("COMPARABLES_MODEL", "claude-sonnet-5-5")
PROMPT_VERSION = "v1"

AUTO_ACCEPT = 0.95          # >= : accept
REVIEW_FLOOR = 0.80         # [0.80, 0.95) : accept + review flag ; < : reject
MIN_FIT = 0.60              # min share of cells that must pass the value check
SAMPLE_N = 6                # sample values per column shown to the LLM
TARGET_MIN_SIGNALS = 3      # triage: header signal groups needed to send a table to the LLM
RENTAL_SOFT_RANGE = (1, 5000)   # outside -> flagged (not rejected)

# Canonical fields the LLM may use (anything else is rejected).
# area_basis is a separate attribute, not a field.
FIELD_DOC = {
    "floor_size": "overall floor size in sq ft, only when it has its OWN column",
    "area": "area (carpet or chargeable) used for the rent, in its OWN column",
    "floor_and_area": "ONE column holding both numbers joined by '~', e.g. '25,000 ~ 18,500' "
                      "(floor size first, area second)",
    "transaction_type": "whether the rent is Quoted / Transacted / CR (codes like Q, T, CR)",
    "rental": "monthly rent RATE per sq ft (INR/sq ft/month). NOT deposit, total rent, sale price",
    "bs_ws": "short code column such as BS/WS or F/UF (kept raw)",
    "property": "comparable property name and/or its location text (combined column -> property)",
    "location": "location text ONLY if it is a separate column from the property",
    "latitude": "latitude alone", "longitude": "longitude alone",
    "coordinates": "latitude and longitude together (Google coordinates, Lat/Long)",
    "date": "date of quote or transaction",
    "remarks": "free-text remarks / notes",
    "unknown": "anything else, or if you are not sure (serial no., deposit, ...)",
}
FIELDS = list(FIELD_DOC)
SINGLE_VALUED = set(FIELDS) - {"remarks", "unknown"}
BASES = ("Carpet", "Chargeable", "Unknown")

CANON_COLS = [
    "Source PDF", "Source Page", "Table ID", "Source Row",
    "Floor Size Sq Ft", "Area Sq Ft", "Area Basis", "Transaction Type",
    "Rental INR / Sq Ft / Month", "BS/WS", "Property", "Location",
    "Latitude", "Longitude", "Date", "Remarks",
    "Review Required", "Review Notes", "Extra Fields",
    "Floor/Area Raw", "Type Raw", "Rental Raw", "Coordinates Raw", "Date Raw",
]

# --------------------------------------------------------------------------
# VALUE PARSERS (deterministic)
# --------------------------------------------------------------------------
NUM = r"\d[\d,]*(?:\.\d+)?"
_UNITS = re.compile(r"(?i)₹|\brs\b\.?|inr|/|sq\.?\s*ft\.?|sqft|psf|per|month|mo|pm|p\.m\.|\s")


def clean(v) -> str:
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def parse_number(text) -> Optional[float]:
    t = clean(text)
    nums = re.findall(NUM, t)
    if len(nums) != 1:
        return None
    if _UNITS.sub("", re.sub(NUM, "", t, count=1)):      # leftover junk -> not a plain number
        return None
    f = float(nums[0].replace(",", "").rstrip("."))
    return int(f) if f.is_integer() else f


def split_combined(text):
    parts = re.split(r"\s*~\s*", clean(text))
    if len(parts) == 2:
        a, b = parse_number(parts[0]), parse_number(parts[1])
        if a is not None and b is not None:
            return a, b
    return None


_FULL = ["%d-%b-%Y", "%d-%b-%y", "%d %b %Y", "%d %B %Y", "%d-%B-%Y", "%d/%m/%Y", "%d/%m/%y",
         "%d.%m.%Y", "%d.%m.%y", "%d-%m-%Y", "%d-%m-%y", "%Y-%m-%d", "%b %d, %Y", "%B %d, %Y",
         "%d %b, %Y", "%d %b %y"]
_PARTIAL = ["%b-%y", "%b %Y", "%B %Y", "%b-%Y", "%B-%Y", "%b %y", "%b'%y", "%m/%Y"]


def parse_date(text):
    """-> (date|None, 'full'|'partial'|None). Day-first (Indian) convention for numeric dates."""
    t = clean(text)
    t = re.sub(r"(?i)\b(\d{1,2})(st|nd|rd|th)\b", r"\1", t)
    t = re.sub(r"(?i)\bsept\b", "Sep", t).rstrip(".")
    if not t or len(t) > 25:
        return None, None
    for fmts, kind in ((_FULL, "full"), (_PARTIAL, "partial")):
        for f in fmts:
            try:
                return datetime.strptime(t, f).date(), kind
            except ValueError:
                pass
    return None, None


def parse_coords(text):
    t = clean(text).replace("°", "")
    nums = re.findall(r"[-+]?\d+\.\d+", t)               # decimals only (integers are ambiguous)
    if len(nums) != 2:
        return None
    letters = re.sub(r"[-+\d.,\s;/:]", "", t)
    if letters and not re.fullmatch(r"[NSEWnsew]{1,2}", letters):
        return None
    lat, lon = float(nums[0]), float(nums[1])
    if re.search(r"\d\s*[sS]\b", t):
        lat = -abs(lat)
    if re.search(r"\d\s*[wW]\b", t):
        lon = -abs(lon)
    if -90 <= lat <= 90 and -180 <= lon <= 180:
        return lat, lon
    return None


def parse_ranged(text, lim):
    t = clean(text)
    if re.fullmatch(r"[-+]?\d+(\.\d+)?", t) and -lim <= float(t) <= lim:
        return float(t)
    return None


TYPE_MAP = {"q": "Quoted", "quoted": "Quoted", "t": "Transacted", "transacted": "Transacted", "cr": "CR"}


def _not_numberish(v):
    return parse_number(v) is None and parse_date(v)[0] is None and parse_coords(v) is None


CHECKERS = {
    "floor_size": lambda v: parse_number(v) is not None,
    "area": lambda v: parse_number(v) is not None,
    "rental": lambda v: parse_number(v) is not None,
    "floor_and_area": lambda v: split_combined(v) is not None,
    "date": lambda v: parse_date(v)[0] is not None,
    "latitude": lambda v: parse_ranged(v, 90) is not None,
    "longitude": lambda v: parse_ranged(v, 180) is not None,
    "coordinates": lambda v: parse_coords(v) is not None,
    "transaction_type": lambda v: bool(re.fullmatch(r"[A-Za-z][A-Za-z /.\-]{0,24}", v)) and parse_date(v)[0] is None,
    "bs_ws": lambda v: len(v) <= 15 and not re.search(r"\d", v),
    "property": _not_numberish,
    "location": _not_numberish,
}


def compute_fit(canon: str, values: list[str]) -> Optional[float]:
    vals = [v for v in values if v]
    chk = CHECKERS.get(canon)
    if not chk or not vals:
        return None
    return sum(1 for v in vals if chk(v)) / len(vals)


def header_basis(h: str) -> Optional[str]:
    """Cross-check only (not a mapping): does the header literally name one area basis?"""
    l = h.lower()
    c, ch = "carpet" in l, "chargeable" in l
    return "Carpet" if c and not ch else "Chargeable" if ch and not c else None


# --------------------------------------------------------------------------
# TABLE EXTRACTION
# --------------------------------------------------------------------------
@dataclass
class Table:
    pdf: str
    page: int
    idx: int
    strategy: str
    page_text: str
    headers: list[str]
    data: list[list[str]]
    header_rows: int

    @property
    def table_id(self):
        return f"{Path(self.pdf).stem}_p{self.page}_t{self.idx}"

    @property
    def ncols(self):
        return len(self.headers)


def looks_like_data_row(row: list[str]) -> bool:
    cells = [c for c in row if c]
    if len(cells) < 2:
        return False
    hits = 0
    for c in cells:
        if (re.search(r"\d", c) and parse_number(c) is not None) or parse_date(c)[0] \
                or parse_coords(c) or split_combined(c):
            hits += 1
    return hits >= 2


def extract_page_tables(page):
    """Strategy 1: ruled lines. Strategy 2: text alignment. (Swap in Camelot/OCR here if needed.)"""
    for name, settings in (("lines", None),
                           ("text", {"vertical_strategy": "text", "horizontal_strategy": "text"})):
        try:
            raw = page.extract_tables(settings) if settings else page.extract_tables()
        except Exception:
            raw = []
        good = [t for t in raw if t and max(len(r) for r in t) >= 4]
        if good:
            return name, good
    return None, []


def extract_pdf(path: Path):
    tables, warnings = [], []
    with pdfplumber.open(path) as pdf:
        for pno, page in enumerate(pdf.pages, 1):
            text = page.extract_text() or ""
            if len(text.strip()) < 20:
                warnings.append((pno, "NO_TEXT_LAYER: page has (almost) no text - scanned? OCR needed"))
                continue
            strategy, raw_tables = extract_page_tables(page)
            for ti, raw in enumerate(raw_tables, 1):
                n = max(len(r) for r in raw)
                rows = [[clean(c) for c in r] + [""] * (n - len(r)) for r in raw]
                rows = [r for r in rows if any(r)]
                first = next((i for i, r in enumerate(rows) if looks_like_data_row(r)), None)
                if first is None:
                    continue
                hdr = rows[max(0, first - 4):first]
                headers = [" ".join(dict.fromkeys(h[c] for h in hdr if h[c])) for c in range(n)]
                if first == 0:
                    headers = [""] * n
                tables.append(Table(path.name, pno, ti, strategy, text, headers, rows[first:], len(hdr)))
    return tables, warnings


def target_score(t: Table) -> int:
    text = " ".join(t.headers).lower()
    groups = [r"rent|lease|quot|transact", r"area|sq\.?\s*ft|floor", r"propert|locat|comparable|address",
              r"date", r"coordinat|\blat|\blong"]
    return sum(1 for g in groups if re.search(g, text))


# --------------------------------------------------------------------------
# LLM: PROPOSES a column mapping (never touches the data values)
# --------------------------------------------------------------------------
SYSTEM = """You classify the columns of a table extracted from a property rental / lease comparables report.
Use the header text AND the sample values AND neighbouring columns AND the page context.
Assign every column exactly one canonical_field from the allowed list. Never invent fields.
If you are not sure, use "unknown" with a low confidence - do not guess.
For columns about area or rent, set area_basis to Carpet, Chargeable or Unknown when the header/context says which
(otherwise null). Echo each column's header text in header_echo exactly as shown.
Field meanings:
""" + "\n".join(f"- {k}: {v}" for k, v in FIELD_DOC.items())

TOOL = {
    "name": "submit_mapping",
    "description": "Submit the canonical field for every source column.",
    "input_schema": {
        "type": "object",
        "properties": {"columns": {"type": "array", "items": {
            "type": "object",
            "properties": {
                "source_column": {"type": "integer", "description": "0-based column index as shown"},
                "header_echo": {"type": "string"},
                "canonical_field": {"type": "string", "enum": FIELDS},
                "area_basis": {"type": ["string", "null"], "enum": ["Carpet", "Chargeable", "Unknown", None]},
                "confidence": {"type": "number"},
                "reason": {"type": "string"},
            },
            "required": ["source_column", "header_echo", "canonical_field", "confidence"],
        }}},
        "required": ["columns"],
    },
}


def column_samples(t: Table):
    return [[r[c] for r in t.data if r[c]][:SAMPLE_N] for c in range(t.ncols)]


def build_payload(t: Table) -> str:
    first_hdr = next((h for h in t.headers if h), "")
    pos = t.page_text.find(first_hdr[:20]) if first_hdr else -1
    ctx = t.page_text[max(0, pos - 600):pos] if pos > 0 else t.page_text[:1500]
    lines = [f"Text on the page before the table:\n{ctx.strip()}\n", f"Table has {t.ncols} columns (0-based):"]
    for c, (h, s) in enumerate(zip(t.headers, column_samples(t))):
        lines.append(f"col {c} | header: {json.dumps(h, ensure_ascii=False)} | samples: {json.dumps(s, ensure_ascii=False)}")
    lines.append("\nFirst two full rows:")
    lines += [json.dumps(r, ensure_ascii=False) for r in t.data[:2]]
    return "\n".join(lines)


def llm_propose(client, t: Table):
    resp = client.messages.create(
        model=MODEL, max_tokens=3000, system=SYSTEM, tools=[TOOL],
        tool_choice={"type": "tool", "name": "submit_mapping"},
        messages=[{"role": "user", "content": build_payload(t)}],
    )
    for b in resp.content:
        if b.type == "tool_use":
            return b.input.get("columns", [])
    raise RuntimeError("LLM returned no tool call")


def get_mapping(client, t: Table, cache: dict, use_cache: bool):
    key = hashlib.sha256(json.dumps([PROMPT_VERSION, MODEL, t.headers, column_samples(t)],
                                    ensure_ascii=False).encode()).hexdigest()
    if use_cache and key in cache:
        return cache[key]
    props = llm_propose(client, t)
    cache[key] = props
    CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")
    return props


# --------------------------------------------------------------------------
# VALIDATION (Python decides)
# --------------------------------------------------------------------------
@dataclass
class ColMap:
    col: int
    header: str
    proposed: str = "unknown"
    canon: str = "unknown"            # final field after validation
    basis: Optional[str] = None
    llm_conf: float = 0.0
    fit: Optional[float] = None
    conf: float = 0.0
    status: str = "unmapped"          # accepted | review | rejected | unmapped
    reason: str = ""
    notes: list = dc_field(default_factory=list)

    @property
    def active(self):
        return self.status in ("accepted", "review") and self.canon != "unknown"


def _norm(s):
    return re.sub(r"[^a-z0-9]", "", s.lower())[:12]


def _demote(cm: ColMap, note: str):
    cm.notes.append(note)
    cm.canon, cm.status = "unknown", "rejected"


def validate_mapping(t: Table, proposals: list):
    flags: list[str] = []
    by_col = {}
    for p in proposals:
        try:
            c = int(p.get("source_column"))
        except (TypeError, ValueError):
            flags.append("BAD_COLUMN_INDEX")
            continue
        if 0 <= c < t.ncols and c not in by_col:
            by_col[c] = p
        else:
            flags.append(f"COLUMN_INDEX_OUT_OF_RANGE_OR_DUPLICATE({c})")

    cols = list(zip(*t.data)) if t.data else []
    cmaps = []
    for c, h in enumerate(t.headers):
        cm = ColMap(col=c, header=h)
        p = by_col.get(c)
        values = [r[c] for r in t.data]
        if p is None:
            cm.notes.append("LLM_OMITTED_COLUMN")
            cmaps.append(cm)
            continue
        cm.proposed = str(p.get("canonical_field", "")).strip()
        cm.reason = str(p.get("reason", ""))[:300]
        try:
            cm.llm_conf = min(1.0, max(0.0, float(p.get("confidence", 0))))
        except (TypeError, ValueError):
            cm.llm_conf = 0.0
        b = p.get("area_basis")
        cm.basis = b if b in BASES else None

        if cm.proposed not in FIELDS:
            _demote(cm, f"INVALID_FIELD({cm.proposed})")
            cmaps.append(cm)
            continue
        if cm.proposed == "unknown":
            cmaps.append(cm)
            continue
        if h and _norm(str(p.get("header_echo", ""))) != _norm(h):
            _demote(cm, "HEADER_ECHO_MISMATCH (possible column-index error)")
            cmaps.append(cm)
            continue

        cm.canon = cm.proposed
        cm.fit = compute_fit(cm.canon, values)
        if cm.fit is not None and cm.fit < MIN_FIT:
            _demote(cm, f"VALUE_MISMATCH (only {cm.fit:.0%} of cells fit '{cm.canon}')")
            cmaps.append(cm)
            continue
        cm.conf = cm.llm_conf if cm.fit is None else 0.6 * cm.llm_conf + 0.4 * cm.fit
        if cm.conf < REVIEW_FLOOR:
            _demote(cm, f"LOW_CONFIDENCE({cm.conf:.2f})")
        elif cm.conf < AUTO_ACCEPT:
            cm.status = "review"
        else:
            cm.status = "accepted"

        # area-basis cross-check against the header's literal wording
        if cm.canon in ("area", "floor_and_area", "rental"):
            hb = header_basis(h)
            if hb and cm.basis in ("Carpet", "Chargeable") and hb != cm.basis:
                cm.basis = "Unknown"
                cm.notes.append("BASIS_CONFLICT (LLM vs header text) -> Unknown")
                flags.append(f"BASIS_CONFLICT col {c}")
            elif hb and cm.basis in (None, "Unknown"):
                cm.basis = hb
                cm.notes.append("basis taken from header wording")
        cmaps.append(cm)

    # duplicate single-valued fields: keep the best, demote the rest, flag the table
    for f in SINGLE_VALUED:
        same = sorted((cm for cm in cmaps if cm.active and cm.canon == f), key=lambda x: -x.conf)
        for loser in same[1:]:
            _demote(loser, f"DUPLICATE_MAPPING({f})")
            flags.append(f"DUPLICATE_MAPPING({f})")
    # conflicting representations of the same thing
    if any(cm.active and cm.canon == "floor_and_area" for cm in cmaps):
        for cm in cmaps:
            if cm.active and cm.canon in ("floor_size", "area"):
                _demote(cm, "CONFLICTS_WITH floor_and_area")
                flags.append("FLOOR_AREA_CONFLICT")
    if any(cm.active and cm.canon == "coordinates" for cm in cmaps):
        for cm in cmaps:
            if cm.active and cm.canon in ("latitude", "longitude"):
                _demote(cm, "CONFLICTS_WITH coordinates")
                flags.append("COORD_CONFLICT")

    active = {cm.canon for cm in cmaps if cm.active}
    is_target = "rental" in active and bool(active & {"area", "floor_and_area", "date"})

    bases = {cm.basis for cm in cmaps if cm.active and cm.canon in ("area", "floor_and_area", "rental")
             and cm.basis in ("Carpet", "Chargeable")}
    if len(bases) > 1:
        flags.append("AREA_BASIS_MISMATCH between area and rental columns")
    for cm in cmaps:
        if cm.status == "review":
            flags.append(f"MAPPING_REVIEW({cm.canon}:{cm.conf:.2f})")
    return cmaps, sorted(set(flags)), is_target


def revalidate_continuation(cmaps: list[ColMap], t: Table) -> bool:
    for cm in cmaps:
        if not cm.active:
            continue
        cm.fit = compute_fit(cm.canon, [r[cm.col] for r in t.data])
        if cm.fit is not None and cm.fit < MIN_FIT:
            if cm.canon in ("rental", "date", "area", "floor_and_area"):
                return False
            _demote(cm, "VALUE_MISMATCH on continuation page")
    return True


# --------------------------------------------------------------------------
# NORMALISATION (deterministic)
# --------------------------------------------------------------------------
def normalize_table(t: Table, cmaps: list[ColMap], table_notes: list[str], review: list):
    area_basis = next((cm.basis for cm in cmaps if cm.active and cm.canon in ("area", "floor_and_area")
                       and cm.basis in ("Carpet", "Chargeable")), "Unknown")
    out = []
    for r_i, row in enumerate(t.data, 1):
        rec = {k: None for k in CANON_COLS}
        rec.update({"Source PDF": t.pdf, "Source Page": t.page, "Table ID": t.table_id,
                    "Source Row": r_i, "Area Basis": area_basis})
        notes, extras, remarks = list(table_notes), {}, []

        def fail(what, raw):
            notes.append(f"{what}_UNPARSED")
            review.append(dict(Severity="warning", **loc(t), Issue=f"{what}_UNPARSED",
                               Detail=f"row {r_i}: raw value {raw!r}"))

        for cm in cmaps:
            v = row[cm.col] if cm.col < len(row) else ""
            if not v:
                continue
            f = cm.canon if cm.active else "unknown"
            if f == "floor_and_area":
                rec["Floor/Area Raw"] = v
                s = split_combined(v)
                if s:
                    rec["Floor Size Sq Ft"], rec["Area Sq Ft"] = s
                else:
                    fail("FLOOR_AREA", v)
            elif f == "floor_size":
                n = parse_number(v)
                rec["Floor Size Sq Ft"] = n
                if n is None:
                    fail("FLOOR_SIZE", v)
            elif f == "area":
                n = parse_number(v)
                rec["Area Sq Ft"] = n
                if n is None:
                    fail("AREA", v)
            elif f == "rental":
                rec["Rental Raw"] = v
                n = parse_number(v)
                rec["Rental INR / Sq Ft / Month"] = n
                if n is None:
                    fail("RENTAL", v)
                elif not (RENTAL_SOFT_RANGE[0] <= n <= RENTAL_SOFT_RANGE[1]):
                    notes.append("RENTAL_OUT_OF_EXPECTED_RANGE")
            elif f == "transaction_type":
                rec["Type Raw"] = v
                m = TYPE_MAP.get(v.lower().strip(". "))
                rec["Transaction Type"] = m or v
                if not m:
                    notes.append(f"TYPE_UNRECOGNISED({v})")
            elif f == "bs_ws":
                rec["BS/WS"] = v
            elif f == "property":
                rec["Property"] = v
            elif f == "location":
                rec["Location"] = v
            elif f == "latitude":
                rec["Latitude"] = parse_ranged(v, 90)
                if rec["Latitude"] is None:
                    fail("LATITUDE", v)
                    rec["Coordinates Raw"] = v
            elif f == "longitude":
                rec["Longitude"] = parse_ranged(v, 180)
                if rec["Longitude"] is None:
                    fail("LONGITUDE", v)
                    rec["Coordinates Raw"] = (rec["Coordinates Raw"] or "") + v
            elif f == "coordinates":
                rec["Coordinates Raw"] = v
                c = parse_coords(v)
                if c:
                    rec["Latitude"], rec["Longitude"] = c
                else:
                    fail("COORDINATES", v)
            elif f == "date":
                rec["Date Raw"] = v
                d, kind = parse_date(v)
                rec["Date"] = d
                if d is None:
                    fail("DATE", v)
                elif kind == "partial":
                    notes.append("DATE_PARTIAL (day set to 1)")
            elif f == "remarks":
                remarks.append(v)
            else:  # unknown / rejected -> never forced into a canonical field
                key = cm.header or f"col_{cm.col}"
                extras[key if key not in extras else f"{key} [col {cm.col}]"] = v

        if rec["Rental INR / Sq Ft / Month"] is None and rec["Floor Size Sq Ft"] is None \
                and rec["Area Sq Ft"] is None:
            review.append(dict(Severity="info", **loc(t), Issue="ROW_SKIPPED",
                               Detail=f"row {r_i} has no rental/area/floor size: {' | '.join(c for c in row if c)[:200]}"))
            continue

        rec["Remarks"] = " | ".join(remarks) or None
        rec["Extra Fields"] = json.dumps(extras, ensure_ascii=False) if extras else None
        rec["Review Notes"] = "; ".join(dict.fromkeys(notes)) or None
        rec["Review Required"] = bool(notes)
        out.append(rec)
    return out


def loc(t: Table):
    return {"Source PDF": t.pdf, "Page": t.page, "Table ID": t.table_id}


def raw_layer(t: Table, cmaps):
    rows = []
    for r_i, row in enumerate(t.data, 1):
        for c, v in enumerate(row):
            cm = cmaps[c] if cmaps else None
            rows.append({"Source PDF": t.pdf, "Page": t.page, "Table ID": t.table_id, "Row": r_i, "Col": c,
                         "Original Header": t.headers[c], "Raw Value": v,
                         "Canonical Field": cm.canon if cm else None,
                         "Confidence": round(cm.conf, 3) if cm else None,
                         "Mapping Status": cm.status if cm else None})
    return rows


def mapping_rows(t: Table, cmaps, flags, is_target, continuation=False):
    return [{"Table ID": t.table_id, "Source PDF": t.pdf, "Page": t.page, "Col": cm.col,
             "Original Header": cm.header, "LLM Proposed": cm.proposed, "Final Field": cm.canon,
             "Area Basis": cm.basis, "LLM Conf": round(cm.llm_conf, 3),
             "Value Fit": None if cm.fit is None else round(cm.fit, 3), "Combined Conf": round(cm.conf, 3),
             "Status": cm.status, "Notes": "; ".join(cm.notes), "LLM Reason": cm.reason,
             "Continuation": continuation} for cm in cmaps]


# --------------------------------------------------------------------------
# MAIN
# --------------------------------------------------------------------------
def main():
    global PDF_DIR, OUT_XLSX, MODEL
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf-dir", default=str(PDF_DIR))
    ap.add_argument("--out", default=str(OUT_XLSX))
    ap.add_argument("--model", default=MODEL)
    ap.add_argument("--extract-only", action="store_true", help="no LLM: only dump extracted tables")
    ap.add_argument("--all-tables", action="store_true", help="send every table to the LLM (no triage)")
    ap.add_argument("--no-cache", action="store_true")
    a = ap.parse_args()
    PDF_DIR, OUT_XLSX, MODEL = Path(a.pdf_dir), Path(a.out), a.model

    pdfs = sorted(p for p in PDF_DIR.glob("*") if p.suffix.lower() == ".pdf")
    if not pdfs:
        sys.exit(f"No PDFs found in {PDF_DIR}/")

    client = None
    if not a.extract_only:
        try:
            import anthropic
            client = anthropic.Anthropic()
        except Exception as e:
            sys.exit(f"Cannot create Anthropic client ({e}). Set ANTHROPIC_API_KEY or use --extract-only.")
    cache = {}
    if CACHE_FILE.exists() and not a.no_cache:
        cache = json.loads(CACHE_FILE.read_text(encoding="utf-8"))

    normalized, review, mappings, raw, tables_sheet = [], [], [], [], []

    for pdf in pdfs:
        print(f"== {pdf.name}")
        try:
            tables, warns = extract_pdf(pdf)
        except Exception as e:
            review.append(dict(Severity="error", **{"Source PDF": pdf.name, "Page": None, "Table ID": None},
                               Issue="EXTRACTION_FAILED", Detail=str(e)))
            continue
        for pno, w in warns:
            review.append(dict(Severity="warning", **{"Source PDF": pdf.name, "Page": pno, "Table ID": None},
                               Issue=w.split(":")[0], Detail=w))

        prev = None   # state for multi-page continuation
        for t in tables:
            score = target_score(t)
            tables_sheet.append({"Table ID": t.table_id, "Source PDF": t.pdf, "Page": t.page,
                                 "Strategy": t.strategy, "Data Rows": len(t.data), "Cols": t.ncols,
                                 "Header Rows": t.header_rows, "Target Score": score,
                                 "Headers": " || ".join(t.headers),
                                 "First Row": " | ".join(t.data[0])[:200]})
            if a.extract_only:
                raw += raw_layer(t, None)
                continue

            # ---- headerless table: possible continuation of the previous page's table
            if t.header_rows == 0:
                if prev and t.page == prev["page"] + 1 and t.ncols == prev["ncols"]:
                    cmaps = copy.deepcopy(prev["cmaps"])
                    if revalidate_continuation(cmaps, t):
                        t.headers = [cm.header for cm in cmaps]
                        notes = prev["flags"] + [f"CONTINUATION of {prev['table_id']}"]
                        normalized += normalize_table(t, cmaps, notes, review)
                        raw += raw_layer(t, cmaps)
                        mappings += mapping_rows(t, cmaps, [], True, continuation=True)
                        prev.update(page=t.page, table_id=t.table_id)
                        continue
                    review.append(dict(Severity="warning", **loc(t), Issue="CONTINUATION_REJECTED",
                                       Detail="values do not fit previous page's column mapping"))
                else:
                    review.append(dict(Severity="info", **loc(t), Issue="HEADERLESS_TABLE_NOT_MATCHED",
                                       Detail=f"{t.ncols} cols, no matching previous table"))
                continue

            # ---- normal table
            if not a.all_tables and score < TARGET_MIN_SIGNALS:
                continue
            try:
                proposals = get_mapping(client, t, cache, not a.no_cache)
            except Exception as e:
                review.append(dict(Severity="error", **loc(t), Issue="LLM_FAILED", Detail=str(e)[:300]))
                raw += raw_layer(t, None)
                continue
            cmaps, flags, is_target = validate_mapping(t, proposals)
            mappings += mapping_rows(t, cmaps, flags, is_target)
            raw += raw_layer(t, cmaps)
            for cm in cmaps:
                if cm.status == "rejected":
                    review.append(dict(Severity="warning", **loc(t), Issue="COLUMN_REJECTED",
                                       Detail=f"col {cm.col} '{cm.header}' proposed '{cm.proposed}': {'; '.join(cm.notes)}"))
            for fl in flags:
                review.append(dict(Severity="warning", **loc(t), Issue="TABLE_FLAG", Detail=fl))
            if not is_target:
                review.append(dict(Severity="info", **loc(t), Issue="TABLE_NOT_IDENTIFIED_AS_TARGET",
                                   Detail="needs rental + (area|floor_and_area|date) accepted"))
                continue
            normalized += normalize_table(t, cmaps, flags, review)
            prev = {"page": t.page, "ncols": t.ncols, "cmaps": cmaps, "flags": flags, "table_id": t.table_id}

    # ---- write workbook
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xw:
        sheets = [("Normalized", pd.DataFrame(normalized, columns=CANON_COLS)),
                  ("Review", pd.DataFrame(review, columns=["Severity", "Source PDF", "Page", "Table ID", "Issue", "Detail"])),
                  ("Mappings", pd.DataFrame(mappings)),
                  ("Raw", pd.DataFrame(raw)),
                  ("Tables", pd.DataFrame(tables_sheet))]
        if a.extract_only:
            sheets = [s for s in sheets if s[0] in ("Tables", "Raw", "Review")]
        for name, df in sheets:
            df.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                width = max((len(str(c.value)) for c in list(col)[:200] if c.value is not None), default=8)
                ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 60)

    print(f"\nDone -> {OUT_XLSX}   rows: {len(normalized)}   review items: {len(review)}")


if __name__ == "__main__":
    main()