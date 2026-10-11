"""
Extract "Floor / Carpet (or Chargeable) Area / Type" tables from many PDFs
and save them into one nicely formatted Excel file.

Install:  pip install pdfplumber pandas openpyxl
Usage:    put your PDFs in a folder called "pdfs" next to this script, then
          python extract_tables.py
"""
import glob
import re

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PDF_GLOB = "./*.pdf"
OUTPUT = "extracted_tables.xlsx"

KNOWN = {
    "Source_PDF", "Page", "Sl_No", "Floor_Size", "Area", "Type", "Rental",
    "Furnishing", "Location_Remarks", "Coordinates", "Date", "Parse_Note",
}
TYPE_TOKENS = {"T", "Q", "CR", "R", "TRANSACTED", "QUOTED"}
FURN_TOKENS = {"BS", "WS", "UF", "F"}

DMS = re.compile(r"\d{1,3}°\s*\d{1,2}['’′]\s*[\d.]+[\"”″]?\s*[NSEW]?")
DEC = re.compile(r"-?\d{1,3}\.\d{5,}")
DATE_RE = re.compile(
    r"(?:CY\s*)?Q[1-4]\s*(?:CY\s*)?\d{4}|\bH[12]\s*(?:CY\s*)?\d{4}", re.I
)


# ---------------------------------------------------------------- detection
def _has_keywords(text):
    text = re.sub(r"\s+", " ", str(text or "")).lower()
    return (
        "floor" in text
        and ("carpet" in text or "chargeable" in text)
        and "type" in text
    )


def is_target_header(cells):
    """A real header row has several separate, non-empty header cells.
    (A row where the whole table is crammed into one cell is NOT a header.)"""
    nonempty = [str(c) for c in cells if c and str(c).strip()]
    return len(nonempty) >= 4 and _has_keywords(" ".join(nonempty))


def blob_text(table):
    """If the whole table was extracted as ONE giant cell, return its text."""
    if not table:
        return None
    first = table[0]
    nonempty = [str(c) for c in first if c and str(c).strip()]
    if len(nonempty) == 1 and len(nonempty[0]) > 120:
        return nonempty[0]
    return None


# first match wins, so order matters
RULES = [
    (r"sl\.?\s*no", "Sl_No"),
    (r"floor", "Floor_Size"),
    (r"rental|rent", "Rental"),
    (r"carpet|chargeable|area", "Area"),
    (r"type", "Type"),
    (r"bs\s*/\s*ws|f\s*/\s*uf", "Furnishing"),
    (r"coordinate|co-\s*ordinate|lat", "Coordinates"),
    (r"location|lease|comparable|property|remarks", "Location_Remarks"),
    (r"date", "Date"),
]


def normalize(col):
    c = re.sub(r"\s+", " ", str(col or "")).strip().lower()
    for pattern, name in RULES:
        if re.search(pattern, c):
            return name
    return c or "unknown"


def make_unique(cols):
    seen, out = {}, []
    for c in cols:
        seen[c] = seen.get(c, 0) + 1
        out.append(c if seen[c] == 1 else f"{c}_{seen[c]}")
    return out


# ------------------------------------------------------------------ cleaning
def clean_text(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    s = re.sub(r"\s+", " ", str(v)).strip()
    return None if s.lower() in ("", "nan") else s


def to_number(v):
    """'1,250' -> 1250.0. Anything that is not a plain number stays as text
    (ranges, 'NA', '85/-', ...)."""
    if v is None:
        return None
    s = str(v).replace(",", "").strip()
    return float(s) if re.fullmatch(r"\d+(\.\d+)?", s) else v


def to_date(v):
    """Real dates become Excel dates. Quarters like 'Q3 2025' / 'CY Q1 2026'
    and anything unparseable stay as the original text."""
    if v is None:
        return None
    s = str(v)
    if re.search(r"\bQ[1-4]\b|\bCY\b|\bH[12]\b", s, re.I):
        return v
    parsed = pd.to_datetime(s, dayfirst=True, errors="coerce")
    return parsed.to_pydatetime() if pd.notna(parsed) else v


# ---------------------------------------------------------------- extraction
def split_stacked(rows):
    """Some PDFs put many rows inside one cell, separated by line breaks.
    If several columns in a row agree on the number of lines, split it into
    that many rows. A single wrapped cell (e.g. a long location) is left alone."""
    out = []
    for row in rows:
        cells = [str(c).split("\n") if c else [""] for c in row]
        n = max(len(c) for c in cells)
        if n <= 1:
            out.append(row)
            continue
        counts = [len(c) for c in cells if len(c) > 1]
        if counts.count(n) >= max(2, len(counts) // 2):
            for i in range(n):
                out.append([c[i] if i < len(c) else "" for c in cells])
        else:
            out.append(row)
    return out


FLOOR = r"(?:[A-Z]\d+\s*[–-]\s*)?(?:G\s*\+\s*\d+|GF|FF|SF|LG|UG|B|\d{1,2})"
ROW_RE = re.compile(
    r"(?:\b(?P<sl>\d{1,2})\s+)?(?P<floor>" + FLOOR + r")\s+"
    r"(?P<area>\d{1,2},\d{3}|\d{2,5})\s+(?P<type>CR|T|Q|R)\s+"
    r"(?P<rental>\d{2,4}(?:\.\d+)?)\b(?P<rest>.*)"
)


def parse_blob(text):
    """Last resort for tables that came out as one giant cell: find the data
    rows by their pattern  <floor> <area> <T/Q> <rental>  in the text.
    Location text is interleaved with other columns in these blobs, so it is
    not recovered; coordinates are filled only if the count matches the rows."""
    rows = []
    for line in text.split("\n"):
        m = ROW_RE.search(line)
        if not m:
            continue
        rest = m.group("rest")
        fm = re.match(r"\s*(BS|WS|UF|F)\b", rest)
        dm = DATE_RE.search(rest)
        rows.append({
            "Sl_No": m.group("sl"),
            "Floor_Size": re.sub(r"\s+", " ", m.group("floor")).strip(),
            "Area": m.group("area"),
            "Type": m.group("type"),
            "Rental": m.group("rental"),
            "Furnishing": fm.group(1) if fm else None,
            "Date": re.sub(r"\s+", " ", dm.group(0)).strip() if dm else None,
        })
    if not rows:
        return None

    toks = [t.strip() for t in DMS.findall(text)]
    if len(toks) < 2:
        toks = DEC.findall(text)
    pairs = [f"{toks[i]}, {toks[i + 1]}" for i in range(0, len(toks) - 1, 2)]
    note = "Recovered from page text - verify against PDF (location not captured)"
    for i, r in enumerate(rows):
        r["Coordinates"] = pairs[i] if len(pairs) == len(rows) else None
        r["Parse_Note"] = note
    return pd.DataFrame(rows)


def extract_all(pdf_paths):
    import pdfplumber

    frames = []
    for path in pdf_paths:
        name = path.replace("\\", "/").split("/")[-1]
        last_header = None
        with pdfplumber.open(path) as pdf:
            for pno, page in enumerate(pdf.pages, start=1):
                tables = page.extract_tables()
                found = False

                for table in tables:
                    if not table or len(table) < 2 or blob_text(table):
                        continue
                    header, body = table[0], table[1:]

                    if is_target_header(header):
                        last_header = make_unique([normalize(h) for h in header])
                        print(f"[{name} p{pno}] header: {header}")
                    elif last_header and len(header) == len(last_header):
                        body = table  # continuation page without a header row
                    else:
                        continue

                    body = split_stacked(body)
                    df = pd.DataFrame(body, columns=last_header[: len(body[0])])
                    df["Source_PDF"] = name
                    df["Page"] = pno
                    frames.append(df)
                    found = True

                if not found:  # fall back to parsing collapsed single-cell tables
                    for table in tables:
                        blob = blob_text(table)
                        df = parse_blob(blob) if blob else None
                        if df is not None:
                            df["Source_PDF"] = name
                            df["Page"] = pno
                            frames.append(df)
                            print(f"[{name} p{pno}] recovered {len(df)} rows from collapsed table text")

    if not frames:
        raise SystemExit("No matching tables found. Check the PDFs / header rules.")
    return pd.concat(frames, ignore_index=True)


def _share(series, pred):
    return series.map(pred).mean()


def fix_columns(df):
    """Columns whose header was blank / unrecognised (merged header cells) are
    classified by what is actually in them, then merged into the right column."""
    for col in [c for c in df.columns if c not in KNOWN]:
        vals = df[col].dropna()
        if vals.empty:
            df = df.drop(columns=col)
            continue
        s = vals.astype(str)
        base = re.sub(r"_\d+$", "", col)

        if base in KNOWN:
            target = base
        elif _share(s, lambda x: bool(DMS.search(x) or DEC.search(x))) >= 0.5:
            target = "Coordinates"
        elif _share(s, lambda x: bool(DATE_RE.search(x))) >= 0.5:
            target = "Date"
        elif _share(s, lambda x: x.strip().upper() in TYPE_TOKENS) >= 0.7:
            target = "Type"
        elif _share(s, lambda x: x.strip().upper() in FURN_TOKENS) >= 0.7:
            target = "Furnishing"
        elif _share(s, lambda x: bool(re.fullmatch(r"[\d,]+(\.\d+)?", x.strip()))) >= 0.7:
            target = "Rental" if "Rental" not in df.columns else None
        else:
            target = None

        if target is None:
            continue
        if target in df.columns:
            df[target] = df[target].where(df[target].notna(), df[col])
            df = df.drop(columns=col)
        else:
            df = df.rename(columns={col: target})
    return df


def tidy(df):
    df = df.copy()
    for col in df.columns:
        if col != "Page":
            df[col] = df[col].map(clean_text)

    df = fix_columns(df)

    data_cols = [c for c in df.columns if c not in ("Source_PDF", "Page", "Parse_Note")]
    df = df.dropna(how="all", subset=data_cols)
    key = [c for c in ("Area", "Rental") if c in df.columns]
    if key:  # footnotes / stray text rows have neither area nor rental
        df = df.dropna(how="all", subset=key)
    if "Type" in df.columns:  # repeated header rows
        df = df[df["Type"].fillna("").str.lower().str.strip() != "type"]

    for col in ("Area", "Rental"):
        if col in df.columns:
            df[col] = df[col].map(to_number)
    if "Date" in df.columns:
        df["Date"] = df["Date"].map(to_date)

    preferred = [
        "Source_PDF", "Page", "Sl_No", "Floor_Size", "Area", "Type", "Rental",
        "Furnishing", "Location_Remarks", "Coordinates", "Date", "Parse_Note",
    ]
    ordered = [c for c in preferred if c in df.columns]
    ordered += [c for c in df.columns if c not in ordered]
    return df[ordered].reset_index(drop=True)


# ---------------------------------------------------------------- formatting
NICE_NAMES = {
    "Source_PDF": "Source PDF",
    "Page": "Page",
    "Sl_No": "Sl. No",
    "Floor_Size": "Floor Size",
    "Area": "Carpet / Chargeable Area (sq ft)",
    "Type": "Type (Quoted / Transacted)",
    "Rental": "Rental (INR / sq ft / month)",
    "Furnishing": "Furnishing (BS/WS or F/UF)",
    "Location_Remarks": "Location / Property / Remarks",
    "Coordinates": "Google Coordinates",
    "Date": "Date of Quote / Transaction",
    "Parse_Note": "Extraction Note",
}
WIDTHS = {
    "Source PDF": 34, "Page": 7, "Sl. No": 8, "Floor Size": 16,
    "Carpet / Chargeable Area (sq ft)": 18, "Type (Quoted / Transacted)": 16,
    "Rental (INR / sq ft / month)": 18, "Furnishing (BS/WS or F/UF)": 16,
    "Location / Property / Remarks": 45, "Google Coordinates": 28,
    "Date of Quote / Transaction": 18, "Extraction Note": 34,
}
LEFT_COLS = {"Source PDF", "Location / Property / Remarks", "Google Coordinates", "Extraction Note"}


def format_workbook(path):
    wb = load_workbook(path)
    ws = wb["Data"]

    header_fill = PatternFill("solid", start_color="1F3864")
    band_fill = PatternFill("solid", start_color="EEF2F8")
    thin = Side(style="thin", color="C8CED8")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.row_dimensions[1].height = 42
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    headers = [c.value for c in ws[1]]
    for r, row in enumerate(ws.iter_rows(min_row=2), start=2):
        for cell in row:
            name = headers[cell.column - 1]
            cell.border = border
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True,
                horizontal="left" if name in LEFT_COLS else "center",
            )
            if r % 2 == 1:
                cell.fill = band_fill
            if isinstance(cell.value, (int, float)):
                if name.startswith("Carpet"):
                    cell.number_format = "#,##0"
                elif name.startswith("Rental"):
                    cell.number_format = "#,##0.00"
            if hasattr(cell.value, "year"):
                cell.number_format = "DD-MMM-YYYY"

    for idx, name in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = WIDTHS.get(name, 18)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    s = wb["Summary"]
    for cell in s[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    s.column_dimensions["A"].width = 60
    s.column_dimensions["B"].width = 14
    for row in s.iter_rows(min_row=2):
        for cell in row:
            cell.border = border

    wb.save(path)


def main():
    pdfs = sorted(glob.glob(PDF_GLOB))
    if not pdfs:
        raise SystemExit(f"No PDFs found at {PDF_GLOB}")

    df = tidy(extract_all(pdfs))

    summary = (
        df.groupby("Source_PDF").size().reset_index(name="Rows extracted")
        .rename(columns={"Source_PDF": "PDF"})
    )
    summary.loc[len(summary)] = ["TOTAL", int(summary["Rows extracted"].sum())]

    out = df.rename(columns=NICE_NAMES)
    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        out.to_excel(writer, sheet_name="Data", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)

    format_workbook(OUTPUT)
    print(f"Saved {len(df)} rows to {OUTPUT}")


if __name__ == "__main__":
    main()