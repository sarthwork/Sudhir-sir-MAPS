import os
import re
import json
import pdfplumber
from docx import Document
from urllib.request import urlopen, Request
from urllib.parse import quote
import time

# ─── State Coordinates ────────────────────────────────────────────────────────
STATE_COORDS = {
    "Maharashtra":      {"lat": 19.7515, "lng": 75.7139},
    "Delhi":            {"lat": 28.7041, "lng": 77.1025},
    "Karnataka":        {"lat": 15.3173, "lng": 75.7139},
    "Tamil Nadu":       {"lat": 11.1271, "lng": 78.6569},
    "Telangana":        {"lat": 18.1124, "lng": 79.0193},
    "West Bengal":      {"lat": 22.9868, "lng": 87.8550},
    "Gujarat":          {"lat": 22.2587, "lng": 71.1924},
    "Rajasthan":        {"lat": 27.0238, "lng": 74.2179},
    "Uttar Pradesh":    {"lat": 26.8467, "lng": 80.9462},
    "Punjab":           {"lat": 31.1471, "lng": 75.3412},
    "Kerala":           {"lat": 10.8505, "lng": 76.2711},
    "Odisha":           {"lat": 20.9517, "lng": 85.0985},
    "Bihar":            {"lat": 25.0961, "lng": 85.3131},
    "Madhya Pradesh":   {"lat": 22.9734, "lng": 78.6569},
    "Andhra Pradesh":   {"lat": 15.9129, "lng": 79.7400},
    "Haryana":          {"lat": 29.0588, "lng": 76.0856},
    "Jharkhand":        {"lat": 23.6102, "lng": 85.2799},
    "Assam":            {"lat": 26.2006, "lng": 92.9376},
    "Chhattisgarh":     {"lat": 21.2787, "lng": 81.8661},
    "Goa":              {"lat": 15.2993, "lng": 74.1240},
    "Himachal Pradesh": {"lat": 31.1048, "lng": 77.1734},
    "Uttarakhand":      {"lat": 30.0668, "lng": 79.0193},
}

# ─── City/Locality Coordinates (for better accuracy) ───────────────────────────
LOCALITY_COORDS = {
    # Mumbai areas - key format (CityName, StateName, LocalityKeyword)
    ("Mumbai", "Maharashtra", "Churchgate"): {"lat": 18.9520, "lng": 72.8296},
    ("Mumbai", "Maharashtra", "Mulund"): {"lat": 19.1575, "lng": 72.8455},
    ("Mumbai", "Maharashtra", "Dombivli"): {"lat": 19.2183, "lng": 73.0511},
    ("Mumbai", "Maharashtra", "Wadala"): {"lat": 19.0176, "lng": 72.8479},
    ("Mumbai", "Maharashtra", "Andheri"): {"lat": 19.1136, "lng": 72.8697},
    ("Mumbai", "Maharashtra", "Bandra"): {"lat": 19.0596, "lng": 72.8295},
    ("Mumbai", "Maharashtra", "Powai"): {"lat": 19.1136, "lng": 72.9027},
    # Hyderabad areas
    ("Hyderabad", "Telangana", "Attapur"): {"lat": 17.3849, "lng": 78.6522},
    ("Hyderabad", "Telangana", "Hyderguda"): {"lat": 17.3849, "lng": 78.6522},
    ("Hyderabad", "Telangana", "Tolichowki"): {"lat": 17.4235, "lng": 78.5714},
    ("Hyderabad", "Telangana", "Shaikpet"): {"lat": 17.4235, "lng": 78.5714},
    ("Hyderabad", "Telangana", "Jubilee"): {"lat": 17.3811, "lng": 78.5079},
    # Alternative city names for Hyderabad
    ("K.V.Rangareddy", "Telangana", "Attapur"): {"lat": 17.3849, "lng": 78.6522},
    ("Rangareddy", "Telangana", "Attapur"): {"lat": 17.3849, "lng": 78.6522},
    ("Attapur", "Telangana", "Attapur"): {"lat": 17.3849, "lng": 78.6522},
}

# ─── Lightweight network helpers (pincode lookup + geocoding) ─────────────────
_PINCODE_CACHE_PATH = os.path.join(os.path.dirname(__file__), "pincode_cache.json")
_GEOCODE_CACHE_PATH = os.path.join(os.path.dirname(__file__), "geocode_cache.json")
_pincode_cache = {}
_geocode_cache = {}

def _load_json_cache(path):
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}

def _save_json_cache(path, obj):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(obj, f, indent=2, ensure_ascii=False)
    except Exception:
        pass

_pincode_cache = _load_json_cache(_PINCODE_CACHE_PATH)
_geocode_cache = _load_json_cache(_GEOCODE_CACHE_PATH)

def lookup_pincode(pincode):
    """
    Uses India Post pincode API to get State/District/PostOffice.
    Example: https://api.postalpincode.in/pincode/500048
    """
    if not pincode:
        return None
    if pincode in _pincode_cache:
        return _pincode_cache[pincode]
    try:
        url = f"https://api.postalpincode.in/pincode/{pincode}"
        req = Request(url, headers={"User-Agent": "RealEstateMap/1.0"})
        with urlopen(req, timeout=10) as r:
            data = json.loads(r.read().decode("utf-8", errors="ignore"))
        if isinstance(data, list) and data and data[0].get("Status") == "Success":
            po = (data[0].get("PostOffice") or [])
            item = po[0] if po else {}
            out = {
                "pincode": pincode,
                "state": item.get("State"),
                "district": item.get("District"),
                "division": item.get("Division"),
                "region": item.get("Region"),
                "post_office": item.get("Name"),
            }
            _pincode_cache[pincode] = out
            _save_json_cache(_PINCODE_CACHE_PATH, _pincode_cache)
            return out
    except Exception:
        return None
    return None

def geocode_address(query):
    """
    Geocode using OpenStreetMap Nominatim. Cached to avoid repeated calls.
    Returns (lat, lng) or None.
    """
    if not query:
        return None
    key = re.sub(r"\s+", " ", query.strip().lower())
    if key in _geocode_cache:
        return _geocode_cache[key]
    try:
        # Be polite to Nominatim
        time.sleep(1.0)
        url = (
            "https://nominatim.openstreetmap.org/search?"
            + f"q={quote(query)}&format=json&limit=1&addressdetails=0"
        )
        req = Request(url, headers={"User-Agent": "RealEstateMap/1.0 (tanush)"} )
        with urlopen(req, timeout=15) as r:
            data = json.loads(r.read().decode("utf-8", errors="ignore"))
        if isinstance(data, list) and data:
            lat = float(data[0]["lat"])
            lon = float(data[0]["lon"])
            out = {"lat": lat, "lng": lon}
            _geocode_cache[key] = out
            _save_json_cache(_GEOCODE_CACHE_PATH, _geocode_cache)
            return out
    except Exception:
        return None
    return None

CLIENT_KEYWORDS = [
    "kotak", "hdfc", "icici", "axis bank", "state bank", "sbi",
    "yes bank", "indusind", "pnb", "punjab national", "bank of baroda",
    "canara bank", "union bank", "idbi", "rbl bank", "federal bank",
    "shoppers stop", "lifestyle", "reliance", "tata", "infosys",
    "wipro", "accenture", "deloitte", "pwc", "kpmg",
    "raymond", "titan", "bajaj", "birla", "adani",
    "amazon", "flipkart", "myntra", "zomato", "swiggy",
    "dmart", "big bazaar", "spencer", "max fashion",
    "pantaloons", "westside", "central", "globus",
    "standard chartered", "citibank", "hsbc", "deutsche bank",
    "mahindra", "godrej", "wockhardt", "cipla", "sun pharma",
]

# ─── Text extraction ──────────────────────────────────────────────────────────
def extract_text_from_pdf(path):
    text = ""
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and cell.strip():
                                text += cell.strip() + "\n"
                # Catches text in bordered/merged boxes that tables miss
                words = page.extract_words()
                if words:
                    line = " ".join(w["text"] for w in words)
                    text += line + "\n"
    except Exception as e:
        print(f"  PDF error: {e}")
    return text

def extract_text_from_docx(path):
    try:
        doc = Document(path)
        return "\n".join([p.text for p in doc.paragraphs])
    except Exception as e:
        print(f"  DOCX error: {e}")
        return ""

def extract_text(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return extract_text_from_pdf(path)
    elif ext in [".docx", ".doc"]:
        return extract_text_from_docx(path)
    return ""

# ─── Extractors ───────────────────────────────────────────────────────────────
def extract_client_name(text):
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    # Search first 15 lines for known client keywords
    for line in lines[:15]:
        line_lower = line.lower()
        for kw in CLIENT_KEYWORDS:
            if kw in line_lower:
                # Clean up: take first 80 chars, remove extra spaces
                cleaned = re.sub(r"\s+", " ", line).strip()
                return cleaned[:80] if len(cleaned) > 80 else cleaned
    # Fallback: first non-empty non-numeric line
    for line in lines[:5]:
        if len(line) > 5 and not line[0].isdigit():
            return line[:80]
    return "Unknown Client"

def extract_address(text):
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    joined = "\n".join(lines)

    # ── Step 0: Scan ALL lines for address with floor/shop + state name ──────
    STATE_NAMES = [
        "maharashtra","delhi","karnataka","telangana","tamil nadu",
        "gujarat","rajasthan","haryana","punjab","kerala","west bengal",
        "andhra pradesh","madhya pradesh","uttar pradesh","odisha","bihar"
    ]
    for i, line in enumerate(lines):
        ll = line.lower()
        has_addr_start = re.search(
            r'\b(ground|first|second|third|\d+(?:st|nd|rd|th)?)\s*floor\b|\bshop\s*no\.?\b|\bplot\s*no\.?\b|\bunit\s*no\.?\b',
            ll
        )
        if not has_addr_start:
            continue
        # Check this line AND next line combined (address often split across 2 lines)
        next_line = lines[i + 1] if i + 1 < len(lines) else ""
        combined_line = line + " " + next_line
        has_state = any(s in combined_line.lower() for s in STATE_NAMES)
        if has_state and len(combined_line) > 20:
            cleaned = re.sub(r"\s+", " ", combined_line).strip(" .,-")
            return cleaned

    def _is_table_noise(s: str) -> bool:
        s2 = s.lower()
        # These show up in comparable/transaction tables and pollute the "address"
        noisy = [
            "prevailing", "q1", "q2", "q3", "q4", "uf", "ff", "gf",
            "transactions", "transaction", "nearest banks", "benchmark",
            "plot no", "prevailing rent", "rentals", "type (inr", "indicative",
        ]
        return any(k in s2 for k in noisy) or bool(re.search(r"\b\d{1,3}\.\d{4,}\b", s2))

    def _normalize_addr(s: str) -> str:
        s = s.replace("–", "-")
        s = re.sub(r"\s+", " ", s).strip()
        # Keep it as one line, remove stray colons
        s = s.replace(" :", ":").replace(": ", ": ")
        return s.strip(" .,-")

    # 1) Strong preference: capture the "located at <ADDRESS>" sentence/block
    # We stop at the first period after a pincode+state/city-ish tail, so we don't absorb tables.
    located_patterns = [
        r"(?:hereinafter\s*referred\s*to\s*as\s*the\s*['\"]?subject\s*property['\"]?[, ]*)?\s*(?:is\s*)?located\s*at\s*(.+)",
        r"(?:subject\s*property[^\n]{0,80}?(?:is\s*)?located\s*at)\s*(.+)",
    ]
    for pat in located_patterns:
        m = re.search(pat, joined, re.IGNORECASE)
        if not m:
            continue
        tail = m.group(1)
        # Take a limited window (next ~500 chars) to avoid swallowing whole document
        window = tail[:500]
        # If there is a period after a pincode, cut there
        cut = None
        m2 = re.search(r"(.+?\b\d{3}\s*\d{3}\b[^.]{0,80}\.)", window)
        if m2:
            cand = m2.group(1)
        else:
            # Otherwise take up to next line break-ish boundary
            cand = window.split("\n")[0]
        cand = _normalize_addr(cand)
        if cand and (extract_pincode(cand) is not None) and not _is_table_noise(cand):
            return cand

    # 2) Highest accuracy for these reports: the header box often contains the full address.
    # Sometimes it's split across 2 lines (long address line + next line with city/pincode/state).
    # Prefer header lines above "Date" and avoid narrative paragraph text.
    header_like = [
        r"^\s*((?:Ground|First|Second|Third|\d+(?:st|nd|rd|th)?)\s*[Ff]loor[^\n]{10,260}(?:\b\d{3}\s*\d{3}\b[^\n]{0,60})?(?:Telangana|Tamil\s*Nadu|Maharashtra|Karnataka|Delhi|Gujarat|West\s*Bengal|Rajasthan|Uttar\s*Pradesh|Haryana|Punjab|Kerala|Odisha|Bihar|Madhya\s*Pradesh|Andhra\s*Pradesh)\b[^\n]{0,40})\s*$",
        r"^\s*(Shop\s*No\.?[^\n]{10,260}\b\d{3}\s*\d{3}\b[^\n]{0,80})\s*$",
        r"^\s*(Plot\s*No\.?[^\n]{10,260}\b\d{3}\s*\d{3}\b[^\n]{0,80})\s*$",
        r"^\s*(Unit\s*No\.?[^\n]{10,260}\b\d{3}\s*\d{3}\b[^\n]{0,80})\s*$",
        r"^\s*(Office\s*No\.?[^\n]{10,260}\b\d{3}\s*\d{3}\b[^\n]{0,80})\s*$",
    ]
    header_candidates = []

    # Collect header lines until "Date" / narrative begins
    header_lines = []
    for line in lines[:60]:
        ll = line.lower()
        if ll.startswith("date"):
            break
        if "according to the information" in ll:
            break
        if _is_table_noise(line):
            continue
        header_lines.append(line)

    # Try single-line matches first
    for line in header_lines:
        for pat in header_like:
            m = re.match(pat, line, re.IGNORECASE)
            if m:
                cand = _normalize_addr(m.group(1))
                header_candidates.append((len(cand), cand))
                break

    # Then try 2-line combined header (fixes Tolichowki-style split)
    for i in range(len(header_lines) - 1):
        combo = _normalize_addr(f"{header_lines[i]} {header_lines[i+1]}")
        if not combo:
            continue
        if extract_pincode(combo) is None:
            continue
        if _is_table_noise(combo):
            continue
        if not re.match(r"^(Ground\s*Floor|Shop\s*No\.?|Plot\s*No\.?|Unit\s*No\.?|Office\s*No\.?)", combo, re.IGNORECASE):
            continue
        header_candidates.append((len(combo), combo))

    if header_candidates:
        header_candidates.sort(key=lambda x: x[0])
        return header_candidates[0][1]

    # 3) Next preference: a line/block that starts like a real address (Ground Floor/Shop/Unit/etc) and contains a pincode
    start_like = [
        r"(Ground\s*Floor[^\n]{10,220}\b\d{3}\s*\d{3}\b)",
        r"(Shop\s*No\.?[^\n]{10,220}\b\d{3}\s*\d{3}\b)",
        r"(Plot\s*No\.?[^\n]{10,220}\b\d{3}\s*\d{3}\b)",
        r"(Unit\s*No\.?[^\n]{10,220}\b\d{3}\s*\d{3}\b)",
        r"(Office\s*No\.?[^\n]{10,220}\b\d{3}\s*\d{3}\b)",
    ]
    for pat in start_like:
        m = re.search(pat, joined, re.IGNORECASE)
        if m:
            cand = _normalize_addr(m.group(1))
            if cand and not _is_table_noise(cand):
                return cand

    # 4) Any multi-line block that ends with a pincode, prefer longer blocks
    candidates = []
    for i in range(len(lines)):
        buf = []
        for j in range(i, min(i + 6, len(lines))):
            buf.append(lines[j])
            chunk = " ".join(buf)
            if extract_pincode(chunk):
                # de-prioritize obvious "comparable transactions" noise
                noise = any(k in chunk.lower() for k in ["transactions", "nearest banks", "hdfc", "sbi bank", "axis bank"])
                # penalize table-like chunks aggressively
                score = len(chunk) - (300 if _is_table_noise(chunk) else 0) - (80 if noise else 0)
                candidates.append((score, chunk))
                break
    if candidates:
        candidates.sort(key=lambda x: x[0], reverse=True)
        return _normalize_addr(candidates[0][1])

    # 3) Fallback patterns
    patterns = [
        r"(Ground\s*Floor[^\n]{10,180}\b\d{6}\b)",
        r"(Shop\s*No\.?\s*[\d,\s]+(?:and\s+\d+)?[,\s]+[^\n]{10,160}\b\d{6}\b)",
        r"(Plot\s*No\.?\s*[\w\-,\s]+[,\s]+[^\n]{10,160}\b\d{6}\b)",
        r"([A-Z][^\n]{15,180}\b\d{6}\b)",
    ]
    for pat in patterns:
        m = re.search(pat, joined, re.IGNORECASE)
        if m:
            addr = _normalize_addr(m.group(1).strip().replace("\n", " "))
            if addr and not _is_table_noise(addr):
                return addr
    # Last resort: any line containing "Floor" or "Shop No" with a state name
    last_resort = re.search(
        r"((?:Ground|First|Second|\d+(?:st|nd|rd|th)?)\s*[Ff]loor[^.\n]{10,200}(?:Maharashtra|Delhi|Karnataka|Telangana|Tamil\s*Nadu|Gujarat|Rajasthan|Haryana|Punjab|Kerala|West\s*Bengal)[^.\n]{0,60}\.?)",
        joined, re.IGNORECASE
    )
    if last_resort:
        return _normalize_addr(last_resort.group(1))

    shop_resort = re.search(
        r"(Shop\s*No\.?[^.\n]{10,200}(?:Maharashtra|Delhi|Karnataka|Telangana|Tamil\s*Nadu|Gujarat|Rajasthan|Haryana|Punjab|Kerala|West\s*Bengal)[^.\n]{0,60}\.?)",
        joined, re.IGNORECASE
    )
    if shop_resort:
        return _normalize_addr(shop_resort.group(1))

    return "Address not found"

def extract_rate(text):
    """
    Extract rental rate range from CWI-style PDF reports.
    Strategy:
      1. Look for bare range (e.g. '320-345' or '320 – 345') inside the
         Rental Range table cell — the cell only contains the number range,
         nothing else on that line.
      2. Look in Conclusion section for 'rental range of INR 320 – 345'.
      3. Scored fallback scanning entire doc for any INR range.
    """
    lowered = text.lower()

    # ── Priority 1: Bare number range on its own line near rental table headers ──
    # CWI PDFs extract the table cell as a line that is ONLY the range, e.g. "320-345"
    # We look for such a line within a window after "Rental Range for Subject Property"
    table_header_match = re.search(
        r"rental range for subject property[^\n]{0,80}\n(.*?\n){0,6}",
        lowered
    )
    if table_header_match:
        # Check the next 6 lines after the header for a bare range
        header_end = table_header_match.end()
        window = text[table_header_match.start(): min(len(text), table_header_match.start() + 600)]
        bare = re.search(r"^\s*(\d{2,4})\s*[-–—]\s*(\d{2,4})\s*$", window, re.MULTILINE)
        if bare:
            lo, hi = bare.group(1), bare.group(2)
            if 50 <= int(lo) <= 9999 and 50 <= int(hi) <= 9999 and int(lo) <= int(hi):
                return f"{lo} - {hi} INR/sq.ft/month"

    # ── Priority 2: "Indicative Range of Lease Rental" table row ──
    # Pattern seen in screenshots: line contains the address then next cell has range
    indicative_match = re.search(
        r"indicative\s*range\s*of\s*lease\s*rental.{0,400}?(\d{2,4})\s*[-–—]\s*(\d{2,4})",
        lowered + text,  # search lowered for keyword, grab digits from original
        re.IGNORECASE | re.DOTALL
    )
    if indicative_match:
        # re-search on original text to get actual digits (not lowered)
        m2 = re.search(
            r"indicative\s*range\s*of\s*lease\s*rental.{0,400}?(\d{2,4})\s*[-–—]\s*(\d{2,4})",
            text, re.IGNORECASE | re.DOTALL
        )
        if m2:
            lo, hi = m2.group(1), m2.group(2)
            if 50 <= int(lo) <= 9999 and 50 <= int(hi) <= 9999 and int(lo) <= int(hi):
                return f"{lo} - {hi} INR/sq.ft/month"

    # ── Priority 3: Conclusion section ──
    conclusion_match = re.search(
        r"conclusion.{0,500}?(?:rental range of|rental of|rate of)\s*(?:INR\s*)?(\d{2,4})\s*[-–—]\s*(\d{2,4})",
        text, re.IGNORECASE | re.DOTALL
    )
    if conclusion_match:
        lo, hi = conclusion_match.group(1), conclusion_match.group(2)
        if 50 <= int(lo) <= 9999 and 50 <= int(hi) <= 9999 and int(lo) <= int(hi):
            return f"{lo} - {hi} INR/sq.ft/month"

    # ── Priority 4: Scored full-doc scan (original logic, kept as fallback) ──
    patterns = [
        r"INR\s*([\d,]{2,5})\s*(?:to|[-–—])\s*([\d,]{2,5})\b",
        r"(?:rental|lease|rent)[^\n]{0,120}?([\d,]{2,5})\s*(?:to|[-–—])\s*([\d,]{2,5})\b",
        r"([\d,]{2,5})\s*[-–—]\s*([\d,]{2,5})\s*(?:INR|Rs\.?)?",
    ]
    best = None
    for pat in patterns:
        for m in re.finditer(pat, text, re.IGNORECASE):
            g = m.groups()
            if len(g) < 2 or not g[0] or not g[1]:
                continue
            lo = g[0].replace(",", "")
            hi = g[1].replace(",", "")
            try:
                lo_i, hi_i = int(lo), int(hi)
            except Exception:
                continue
            if not (50 <= lo_i <= 9999 and 50 <= hi_i <= 9999 and lo_i <= hi_i):
                continue

            start = max(0, m.start() - 250)
            end = min(len(lowered), m.end() + 250)
            ctx = lowered[start:end]

            score = 0
            if "conclusion" in ctx:           score += 10
            if "indicative" in ctx:           score += 8
            if "lease rental" in ctx:         score += 8
            if "rental range for subject" in ctx: score += 12
            if "rental range" in ctx:         score += 6
            if "subject property" in ctx:     score += 4
            if any(k in ctx for k in ("sq.ft", "sq ft", "psf", "per sq")): score += 2
            score += int((m.start() / max(1, len(text))) * 2)

            if best is None or score > best[0]:
                best = (score, lo, hi)

    if best:
        return f"{best[1]} - {best[2]} INR/sq.ft/month"
    return "Rate not found"

def extract_size(text):
    patterns = [
        r"carpet\s*area[^\d]{0,30}([\d,]+)\s*(?:sq\.?\s*ft|sqft|sft)",
        r"([\d,]+)\s*(?:sq\.?\s*ft|sqft|sft)[^\n]{0,40}(?:carpet|on\s*carpet)",
        r"area\s*of\s*([\d,]+)\s*(?:sq\.?\s*ft|sqft|sft)",
        r"(?:size|area)[^\d]{0,20}([\d,]+)\s*(?:sq\.?\s*ft|sqft|sft)",
        r"([\d,]+)\s*(?:sq\.?\s*ft|sqft|sft)",  # fallback
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = m.group(1).replace(",", "")
            if 50 <= int(val) <= 1000000:
                return val + " sq.ft"
    return "Size not found"
def extract_comparable_rentals_from_tables(filepath):
    """
    Reads comparable rentals directly from pdfplumber tables.
    Table 1 on page 2 has clean rows: [floor, size, type, rate, bs_ws, property, coords, date]
    """
    comparables = []
    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    # Check header row for comparable table signature
                    header = [str(c).lower().replace("\n", " ").strip() if c else "" for c in table[0]]
                    header_text = " ".join(header)
                    if not ("floor" in header_text and ("rental" in header_text or "basis" in header_text) and ("bs" in header_text or "ws" in header_text)):
                        continue
                    # Find column indices from header
                    col_floor    = next((i for i, h in enumerate(header) if "floor" in h), 0)
                    col_size     = next((i for i, h in enumerate(header) if "area" in h or "size" in h), 1)
                    col_type     = next((i for i, h in enumerate(header) if h.strip() == "type"), 2)
                    col_rate     = next((i for i, h in enumerate(header) if "rental" in h or "basis" in h or "inr" in h), 3)
                    col_bsws     = next((i for i, h in enumerate(header) if "bs" in h or "ws" in h), 4)
                    col_property = next((i for i, h in enumerate(header) if "comparable" in h or "property" in h or "lease" in h), 5)
                    # Date is always the LAST column in CWI tables
                    col_date     = len(header) - 1
                    # But verify — if last col has "date"/"transaction"/"quote" use it, else search
                    for ci, h in enumerate(header):
                        if "date" in h or "transaction" in h or "quote" in h:
                            col_date = ci
                            break
                    # Parse data rows
                    for row in table[1:]:
                        if not row or all(c is None or str(c).strip() == "" for c in row):
                            continue
                        def cell(idx):
                            if idx >= len(row) or row[idx] is None:
                                return ""
                            return str(row[idx]).replace("\n", " ").strip()
                        floor_val = cell(col_floor)
                        # Must be a comparable row
                        if not re.match(r"^(C\d+|GF|Ground|First|Second|Third)", floor_val, re.IGNORECASE):
                            continue
                        rate_raw  = cell(col_rate)
                        rate_m    = re.search(r"(\d{2,4})", rate_raw)
                        if not rate_m:
                            continue
                        # Clean property name — strip coordinates like 19°13'06.7"N
                        prop_raw  = cell(col_property)
                        prop_val  = re.sub(r"\d{1,2}[°\s]\d+['\s][\d.]+[\"']?\s*[NnSs].*", "", prop_raw).strip()
                        # Search EVERY cell in the row for a date pattern Q1 2026 / Q3 2024
                        date_val = "N/A"
                        for _ci in range(len(row)):
                            _raw = cell(_ci)
                            _dm = re.search(r"(?:CY\s*)?(Q[1-4]\s*\d{4})", _raw, re.IGNORECASE)
                            if _dm:
                                date_val = _dm.group(1).strip()
                                break
                        size_raw  = cell(col_size).replace(",", "")
                        size_val  = re.sub(r"[^\d]", "", size_raw)
                        comparables.append({
                            "floor":    floor_val,
                            "size":     size_val,
                            "type":     cell(col_type).upper() or "Q",
                            "rate":     rate_m.group(1),
                            "bs_ws":    cell(col_bsws).upper() or "BS",
                            "property": prop_val,
                            "date":     date_val,
                        })
    except Exception as e:
        print(f"  Comparable table extraction error: {e}")
    return comparables
def extract_comparable_rentals(text, filepath=None):
    """
    CWI PDFs extract table rows as individual lines (one cell per line).
    Strategy: find the comparable section, then collect values line by line
    and assemble rows based on known cell patterns.
    """
    # Try table extraction first (most accurate for CWI PDFs)
    if filepath and filepath.lower().endswith(".pdf"):
        result = extract_comparable_rentals_from_tables(filepath)
        if result:
            return result

    comparables = []
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    # ── Find section boundaries ──────────────────────────────────────────────
    section_start = -1
    section_end   = len(lines)

    for i, line in enumerate(lines):
        ll = line.lower()
        if section_start == -1 and re.search(r"prevailing\s*comparable\s*rental", ll):
            section_start = i
        elif section_start != -1 and re.search(
            r"(bs\s*=\s*bareshell|negotiation\s*margin|conclusion|site\s*description|subject\s*property\s*detail)",
            ll
        ):
            section_end = i
            break

    if section_start == -1:
        # Try looser header match
        for i, line in enumerate(lines):
            if re.search(r"comparable\s*rental", line, re.IGNORECASE):
                section_start = i
                break

    if section_start == -1:
        return []

    section = lines[section_start:section_end]

    # ── Cell-type classifiers ────────────────────────────────────────────────
    def is_floor(s):
        return bool(re.match(r"^(C\d+[-–]\s*\w+|Ground|First|Second|Third|\d+(?:st|nd|rd|th)?\s*[Ff]loor?)$", s, re.IGNORECASE))

    def is_size(s):
        # pure number like "4,000" or "4000" or "1,801"
        return bool(re.match(r"^[\d,]{3,7}$", s)) and 100 <= int(s.replace(",","")) <= 500000

    def is_type(s):
        return s.upper() in ("T", "Q")

    def is_rate(s):
        # 2-4 digit number in plausible rental range
        if re.match(r"^\d{2,4}$", s):
            v = int(s)
            return 50 <= v <= 9999
        return False

    def is_bsws(s):
        return s.upper() in ("BS", "WS")

    def is_date(s):
        return bool(re.match(r"^Q[1-4]\s*\d{4}$", s, re.IGNORECASE))

    def is_property_name(s):
        # At least 4 chars, starts with capital, not a number/keyword
        return (
            len(s) >= 4
            and s[0].isupper()
            and not is_floor(s)
            and not is_size(s)
            and not is_type(s)
            and not is_rate(s)
            and not is_bsws(s)
            and not is_date(s)
            and not re.match(r"^\d", s)
            and s.lower() not in ("floor", "size", "type", "date", "bs/ws", "location")
        )

    # ── State machine: walk lines and assemble rows ──────────────────────────
    # Expected order: floor → size → type → rate → bs_ws → property(1-2 lines) → date
    STATE_FLOOR  = 0
    STATE_SIZE   = 1
    STATE_TYPE   = 2
    STATE_RATE   = 3
    STATE_BSWS   = 4
    STATE_PROP   = 5
    STATE_DATE   = 6

    row   = {}
    state = STATE_FLOOR
    prop_parts = []

    def flush_row():
        if "floor" in row and "rate" in row:
            row["property"] = " ".join(prop_parts).strip() or "N/A"
            comparables.append({**row})
        row.clear()
        prop_parts.clear()

    for line in section:
        # Skip header/label lines
        if re.search(r"(floor|carpet\s*area|type|quoted|transacted|location|lease|comparable|prevailing|bs/ws|date\s*of)", line, re.IGNORECASE):
            continue

        if state == STATE_FLOOR:
            if is_floor(line):
                flush_row() if row else None
                row = {"floor": line}
                prop_parts.clear()
                state = STATE_SIZE

        elif state == STATE_SIZE:
            if is_size(line):
                row["size"] = line.replace(",", "")
                state = STATE_TYPE
            elif is_floor(line):           # missed a size — reset
                flush_row()
                row = {"floor": line}
                prop_parts.clear()

        elif state == STATE_TYPE:
            if is_type(line):
                row["type"] = line.upper()
                state = STATE_RATE
            elif is_size(line):            # duplicate size cell
                row["size"] = line.replace(",", "")

        elif state == STATE_RATE:
            if is_rate(line):
                row["rate"] = line
                state = STATE_BSWS
            elif is_floor(line):
                flush_row()
                row = {"floor": line}
                prop_parts.clear()
                state = STATE_SIZE

        elif state == STATE_BSWS:
            if is_bsws(line):
                row["bs_ws"] = line.upper()
                state = STATE_PROP
            elif is_rate(line):            # rate repeated
                row["rate"] = line

        elif state == STATE_PROP:
            if is_date(line):
                row["date"] = line
                flush_row()
                state = STATE_FLOOR
            elif is_property_name(line):
                prop_parts.append(line)
            elif is_floor(line):           # next row started — no date found
                row["date"] = "N/A"
                flush_row()
                row = {"floor": line}
                prop_parts.clear()
                state = STATE_SIZE

        elif state == STATE_DATE:
            if is_date(line):
                row["date"] = line
                flush_row()
                state = STATE_FLOOR

    # flush last row
    if row:
        row.setdefault("date", "N/A")
        flush_row()

    # ── Fallback: single-line regex if state machine found nothing ───────────
    if not comparables:
        full = "\n".join(lines)
        pattern = re.compile(
            r"(C\d+[-–]\s*\w+|Ground\s*Floor)\s+"
            r"([\d,]{3,7})\s+"
            r"(T|Q)\s+"
            r"(\d{2,4})\s+"
            r"(BS|WS)\s+"
            r"([A-Za-z][^\n]{4,50}?)\s+"
            r"(Q[1-4]\s*\d{4})",
            re.IGNORECASE
        )
        for m in pattern.finditer(full):
            comparables.append({
                "floor":    m.group(1).strip(),
                "size":     m.group(2).replace(",",""),
                "type":     m.group(3).upper(),
                "rate":     m.group(4),
                "bs_ws":    m.group(5).upper(),
                "property": m.group(6).strip(),
                "date":     m.group(7).strip(),
            })

    # Deduplicate by (floor, rate, property)
    seen, unique = set(), []
    for c in comparables:
        key = (c.get("floor",""), c.get("rate",""), c.get("property",""))
        if key not in seen:
            seen.add(key)
            unique.append(c)

    return unique
def extract_pincode(text):
    # Find 6-digit pincode (supports "500048" and "500 048")
    if not text:
        return None
    # First try spaced format
    m = re.search(r"\b(\d{3})\s*(\d{3})\b", text)
    if m:
        pin = f"{m.group(1)}{m.group(2)}"
        if pin[0] in "123456789":
            return pin
    # Then compact format
    matches = re.findall(r"\b(\d{6})\b", text)
    for pin in matches:
        if pin[0] in "123456789":
            return pin
    return None

def _clean_token(t):
    t = re.sub(r"\s+", " ", t).strip(" ,.-")
    return t

def _strip_pincode_like(s):
    # remove "500048" or "500 048" and stray short numeric tokens
    if not s:
        return s
    s = re.sub(r"\b(\d{3})\s*(\d{3})\b", "", s)
    s = re.sub(r"\b\d{6}\b", "", s)
    # remove standalone small numeric fragments like ", 320" that come from parsing noise
    s = re.sub(r"(?:^|[, ])\d{1,3}(?:$|[, ])", " ", s)
    s = re.sub(r"\s+", " ", s).strip(" ,.-")
    return s

def _strip_coordinate_noise(s):
    """
    Remove embedded coordinate-like numbers from strings, e.g.
    "Somi Reddy Nagar, 78.4297 Prevailing Hyderabad" -> "Somi Reddy Nagar, Hyderabad"
    """
    if not s:
        return s
    # remove long floats (lat/lng)
    s = re.sub(r"\b\d{1,3}\.\d{4,}\b", "", s)
    # remove common report noise tokens
    s = re.sub(r"\b(prevailing|q\d\s*\d{4})\b", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+", " ", s).strip(" ,.-")
    return s

def extract_google_coordinates(text):
    """
    Many C&W reports include a 'Google Coordinates' line.
    Returns (lat, lng) if found.
    """
    if not text:
        return None
    # Try decimal format first: 17.3939, 78.4377
    m = re.search(
        r"Google\s*Coordinates\s*[:\-]?\s*([0-9]{1,2}\.\d+)\s*[,\s]\s*([0-9]{1,3}\.\d+)",
        text, re.IGNORECASE,
    )
    if m:
        try:
            lat = float(m.group(1))
            lng = float(m.group(2))
            if -90 <= lat <= 90 and -180 <= lng <= 180:
                return (lat, lng)
        except Exception:
            pass

    # Try DMS format: 19°13'21.5"N 73°05'12.9"E
    m = re.search(
        r"Google\s*Coordinates\s*[:\-]?\s*"
        r"(\d{1,2})[°\s]\s*(\d{1,2})['\s]\s*([\d.]+)[\"']?\s*[Nn]\s*[,\s]*"
        r"(\d{1,3})[°\s]\s*(\d{1,2})['\s]\s*([\d.]+)[\"']?\s*[Ee]",
        text, re.IGNORECASE,
    )
    if m:
        try:
            lat = int(m.group(1)) + int(m.group(2))/60 + float(m.group(3))/3600
            lng = int(m.group(4)) + int(m.group(5))/60 + float(m.group(6))/3600
            if -90 <= lat <= 90 and -180 <= lng <= 180:
                return (lat, lng)
        except Exception:
            pass
    return None
    if not m:
        return None
    try:
        lat = float(m.group(1))
        lng = float(m.group(2))
        if -90 <= lat <= 90 and -180 <= lng <= 180:
            return (lat, lng)
    except Exception:
        return None
    return None

def extract_location_info(text, address):
    """
    Derive state/city/locality using:
    1. India Post pincode API (primary — gives exact state + district)
    2. Nominatim geocoding (for accurate lat/lng from full address)
    3. Google Coordinates embedded in report (most precise)
    4. Keyword hints (last resort only when pincode API fails)
    5. State center coordinates (final fallback)
    """
    combined = _clean_token((address or "") + " " + (text[:2000] or "")).lower()
    pincode = extract_pincode(address or "") or extract_pincode(text)

    state = None
    city = None
    locality = None
    lat = None
    lng = None

    # ── Step 1: Pincode API (most reliable for Indian addresses) ──────────────
    # api.postalpincode.in returns exact State and District for any 6-digit pincode
    pin_info = lookup_pincode(pincode) if pincode else None
    if pin_info:
        state = pin_info.get("state") or None
        raw_city = pin_info.get("district") or None
        # Normalize known district→city mappings
        DISTRICT_TO_CITY = {
            "K.V.Rangareddy": "Hyderabad",
            "Rangareddy": "Hyderabad",
            "K V Rangareddy": "Hyderabad",
            "Ranga Reddy": "Hyderabad",
            "Medchal Malkajgiri": "Hyderabad",
            "Hyderabad": "Hyderabad",
            "Thane": "Thane",
            "Raigad": "Navi Mumbai",
        }
        city = DISTRICT_TO_CITY.get(raw_city, raw_city)
        print(f"  📮 Pincode API [{pincode}] → state={state}, city={city}")

    # ── Step 2: Extract locality from address parts ───────────────────────────
    if address and address != "Address not found":
        parts = [_clean_token(p) for p in re.split(r",|–|-", address) if _clean_token(p)]
        parts_clean = [_strip_coordinate_noise(_strip_pincode_like(p)) for p in parts]
        parts_clean = [p for p in parts_clean if p]

        # Remove parts that match state or city (already known)
        filtered = []
        for p in parts_clean:
            pl = p.lower()
            if state and state.lower() in pl:
                continue
            if city and city.lower() in pl:
                continue
            if pl in ["india", "india."]:
                continue
            filtered.append(p)

        # Locality = last 1-2 meaningful tokens before city/state
        if filtered:
            locality = filtered[-1]
            if len(filtered) >= 2 and len(filtered[-2]) <= 30:
                locality = f"{filtered[-2]}, {filtered[-1]}"
            locality = _strip_coordinate_noise(locality)

    # ── Step 3: Keyword hints — ONLY if pincode API didn't give us city/state ──
    # These are NOT overrides. They only fill gaps.
    if not state or not city:
        LOCATION_HINTS = [
            # Mumbai localities
            ("churchgate",   "Mumbai",      "Maharashtra"),
            ("wadala",       "Mumbai",      "Maharashtra"),
            ("bandra",       "Mumbai",      "Maharashtra"),
            ("andheri",      "Mumbai",      "Maharashtra"),
            ("borivali",     "Mumbai",      "Maharashtra"),
            ("malad",        "Mumbai",      "Maharashtra"),
            ("goregaon",     "Mumbai",      "Maharashtra"),
            ("kandivali",    "Mumbai",      "Maharashtra"),
            ("dahisar",      "Mumbai",      "Maharashtra"),
            ("kurla",        "Mumbai",      "Maharashtra"),
            ("ghatkopar",    "Mumbai",      "Maharashtra"),
            ("mulund",       "Mumbai",      "Maharashtra"),
            ("vikhroli",     "Mumbai",      "Maharashtra"),
            ("powai",        "Mumbai",      "Maharashtra"),
            ("chembur",      "Mumbai",      "Maharashtra"),
            ("sion",         "Mumbai",      "Maharashtra"),
            ("dadar",        "Mumbai",      "Maharashtra"),
            ("worli",        "Mumbai",      "Maharashtra"),
            ("lower parel",  "Mumbai",      "Maharashtra"),
            ("parel",        "Mumbai",      "Maharashtra"),
            ("colaba",       "Mumbai",      "Maharashtra"),
            ("nariman point","Mumbai",      "Maharashtra"),
            # Mumbai suburbs (separate cities — NOT "Mumbai")
            ("dombivli",     "Dombivli",    "Maharashtra"),
            ("dombivali",    "Dombivli",    "Maharashtra"),
            ("kalyan",       "Kalyan",      "Maharashtra"),
            ("thane",        "Thane",       "Maharashtra"),
            ("navi mumbai",  "Navi Mumbai", "Maharashtra"),
            ("vashi",        "Navi Mumbai", "Maharashtra"),
            ("kharghar",     "Navi Mumbai", "Maharashtra"),
            ("panvel",       "Panvel",      "Maharashtra"),
            ("pune",         "Pune",        "Maharashtra"),
            ("mumbai",       "Mumbai",      "Maharashtra"),
            # Hyderabad localities
            ("tolichowki",   "Hyderabad",   "Telangana"),
            ("attapur",      "Hyderabad",   "Telangana"),
            ("gachibowli",   "Hyderabad",   "Telangana"),
            ("hitech city",  "Hyderabad",   "Telangana"),
            ("madhapur",     "Hyderabad",   "Telangana"),
            ("banjara hills","Hyderabad",   "Telangana"),
            ("jubilee hills","Hyderabad",   "Telangana"),
            ("begumpet",     "Hyderabad",   "Telangana"),
            ("secunderabad", "Hyderabad",   "Telangana"),
            ("hyderabad",    "Hyderabad",   "Telangana"),
            # Chennai localities
            ("teynampet",    "Chennai",     "Tamil Nadu"),
            ("anna salai",   "Chennai",     "Tamil Nadu"),
            ("adyar",        "Chennai",     "Tamil Nadu"),
            ("t nagar",      "Chennai",     "Tamil Nadu"),
            ("velachery",    "Chennai",     "Tamil Nadu"),
            ("chennai",      "Chennai",     "Tamil Nadu"),
            # Bengaluru localities
            ("koramangala",  "Bengaluru",   "Karnataka"),
            ("indiranagar",  "Bengaluru",   "Karnataka"),
            ("whitefield",   "Bengaluru",   "Karnataka"),
            ("hsr layout",   "Bengaluru",   "Karnataka"),
            ("electronic city","Bengaluru", "Karnataka"),
            ("bangalore",    "Bengaluru",   "Karnataka"),
            ("bengaluru",    "Bengaluru",   "Karnataka"),
            # Delhi/NCR
            ("connaught place","New Delhi",  "Delhi"),
            ("nehru place",  "New Delhi",   "Delhi"),
            ("gurgaon",      "Gurugram",    "Haryana"),
            ("gurugram",     "Gurugram",    "Haryana"),
            ("noida",        "Noida",       "Uttar Pradesh"),
            ("faridabad",    "Faridabad",   "Haryana"),
            ("delhi",        "New Delhi",   "Delhi"),
        ]
        for kw, guess_city, guess_state in LOCATION_HINTS:
            if kw in combined:
                city  = city  or guess_city   # only fills if still None
                state = state or guess_state  # only fills if still None
                print(f"  🔑 Keyword hint [{kw}] → city={city}, state={state}")
                break

    # ── Step 4: Google Coordinates from report (most precise lat/lng) ─────────
    gc = extract_google_coordinates(text)
    if gc:
        lat, lng = gc
        print(f"  📍 Google Coords from report → {lat}, {lng}")

    # ── Step 5: Nominatim geocoding for lat/lng ───────────────────────────────
    if lat is None or lng is None:
        # Build best possible query: full address first, then city+state
        if address and address != "Address not found":
            geo_query = address
        elif locality and city:
            geo_query = f"{locality}, {city}, {state}, India"
        elif city and state:
            geo_query = f"{city}, {state}, India"
        else:
            geo_query = None

        if geo_query:
            geo = geocode_address(geo_query)
            if geo:
                lat = geo["lat"]
                lng = geo["lng"]
                print(f"  🌐 Nominatim geocode → {lat}, {lng}")

    # ── Step 6: Locality/City center fallback ────────────────────────────────────
    if (lat is None or lng is None):
        # Try to match locality keywords against LOCALITY_COORDS
        matched = False
        locality_lower = (locality or "").lower()
        city_lower = (city or "").lower()
        
        for (map_city, map_state, map_locality), coords in LOCALITY_COORDS.items():
            if (city_lower == map_city.lower() and state == map_state and 
                map_locality.lower() in locality_lower):
                lat = coords["lat"]
                lng = coords["lng"]
                print(f"  🎯 Locality match [{map_locality}] → {lat}, {lng}")
                matched = True
                break
        
        # If no exact match, try city-level match
        if not matched:
            for (map_city, map_state, map_locality), coords in LOCALITY_COORDS.items():
                if (city_lower == map_city.lower() and state == map_state and 
                    map_city.lower() in locality_lower):
                    lat = coords["lat"]
                    lng = coords["lng"]
                    print(f"  🏙️  City match [{map_city}] → {lat}, {lng}")
                    matched = True
                    break
        
        # Finally state fallback
        if not matched:
            if state in STATE_COORDS:
                lat = STATE_COORDS[state]["lat"]
                lng = STATE_COORDS[state]["lng"]
                print(f"  🗺️  State center fallback → {lat}, {lng}")
            else:
                lat = 20.5937
                lng = 78.9629
                print(f"  ⚠️  India center fallback used")

    state    = state    or "Unknown"
    city     = city     or "Unknown"
    locality = locality or city

    return state, city, locality, lat, lng

# ─── Main processor ───────────────────────────────────────────────────────────
def process_file(filepath):
    print(f"  Reading: {os.path.basename(filepath)}")
    text = extract_text(filepath)

    if not text.strip():
        print(f"  WARNING: No text extracted from {filepath}")

    client = extract_client_name(text)
    address = extract_address(text)
    rate = extract_rate(text)
    size = extract_size(text)
    state, city, locality, lat, lng = extract_location_info(text, address)

    comparables = extract_comparable_rentals(text, filepath=filepath)

    result = {
        "client_name": client,
        "address": address,
        "rate": rate,
        "size": size,
        "state": state,
        "city": city,
        "locality": locality,
        "lat": lat,
        "lng": lng,
        "filename": os.path.basename(filepath),
        "is_saved": False,
        "comparables": comparables,
    }

    print(f"  → Client: {client}")
    print(f"  → Address: {address[:60]}...")
    print(f"  → Rate: {rate}")
    print(f"  → Size: {size}")
    print(f"  → Location: {locality}, {city}, {state} [{lat},{lng}]")
    return result

def process_all_files(folder="uploads"):
    results = []
    if not os.path.exists(folder):
        os.makedirs(folder)
        return results

    files = [f for f in os.listdir(folder)
             if f.lower().endswith((".pdf", ".docx", ".doc"))]

    if not files:
        print(f"  No PDF/DOCX files found in {folder}")
        return results

    for fname in files:
        fpath = os.path.join(folder, fname)
        print(f"\n📄 Processing: {fname}")
        try:
            data = process_file(fpath)
            results.append(data)
            print(f"  ✅ Done")
        except Exception as e:
            print(f"  ❌ Error: {e}")
            import traceback
            traceback.print_exc()

    print(f"\n✅ Total: {len(results)} file(s) processed")
    return results

if __name__ == "__main__":
    results = process_all_files()
    with open("extracted_data.json", "w") as f:
        json.dump(results, f, indent=2)
    print(json.dumps(results, indent=2))