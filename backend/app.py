from flask import Flask, jsonify, request, send_file , send_from_directory
from flask_cors import CORS
import json
import os
import io
from extractor import process_all_files, STATE_COORDS
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import re
import pandas as pd
import reverse_geocoder as rg
from flask import jsonify

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = "uploads"
DATA_FILE = "extracted_data.json"
SAVED_FILE = "saved_properties.json"
ADMIN_PASSWORD = "Raacushwake@10cr"

_saved_filenames = set() 

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def load_saved():
    if os.path.exists(SAVED_FILE):
        with open(SAVED_FILE) as f:
            return json.load(f)
    return []

def save_saved(data):
    with open(SAVED_FILE, "w") as f:
        json.dump(data, f, indent=2)

EXCEL_PATH = "./uploads/land_valuation_updated.xlsx"   # your file

# COLUMNS = [
#     'Sr. no.', 'Date of Valuation', 'Client Name', 'Micromarket',
#     'Google Co-ordiantes',          # note: spelling matches your header
#     'Zoning', 'Total Land Area (Acres)', 'Total Land Area (sqm)',
#     'Market Value (INRMn)', 'M.V. per acre', 'Land Zoning'
# ]

COLUMNS = ['Sr. no.', 'Date of Valuation', 'Client Name', 'Micromarket',
           'Google Co-ordiantes', 'Zoning', 'Total Land Area (Acres)',
           'Total Land Area (sqm)', 'Market Value (INR Mn)',
           'M.V. per acre', 'Land Zoning']

def _dms_to_dec(deg, minutes, seconds, hemi):
    val = float(deg) + float(minutes or 0) / 60 + float(seconds or 0) / 3600
    return -val if hemi.upper() in ("S", "W") else val

def parse_coords(val):
    """Handles decimal ('18.95, 72.82', '18.95° N, 72.82° E') and DMS
    ('19°03'34.2"N 72°50'11.6"E'). Returns (lat, lng) or (None, None)."""
    if val is None or str(val).strip() == "":
        return None, None
    s = str(val).strip()

    # DMS: degrees, minutes, optional seconds, hemisphere letter
    dms = re.findall(
        r"""(\d+(?:\.\d+)?)\s*[°º]\s*(\d+(?:\.\d+)?)\s*['′’]?\s*(?:(\d+(?:\.\d+)?)\s*["″”]?)?\s*([NSEWnsew])""",
        s)
    if len(dms) >= 2:
        a = _dms_to_dec(*dms[0])
        b = _dms_to_dec(*dms[1])
        lat, lng = (a, b) if dms[0][3].upper() in "NS" else (b, a)
    else:
        # Decimal: capture numbers with an optional hemisphere letter
        parts = re.findall(r"(-?\d+(?:\.\d+)?)\s*°?\s*([NSEWnsew])?", s)
        if len(parts) < 2:
            return None, None
        vals = []
        for num, hemi in parts[:2]:
            v = float(num)
            if hemi and hemi.upper() in ("S", "W"):
                v = -abs(v)
            vals.append((v, hemi.upper() if hemi else ""))
        (a, ha), (b, hb) = vals
        if ha in ("E", "W") or hb in ("N", "S"):
            a, b = b, a            # given as lng, lat
        lat, lng = a, b

    # Sanity check for India: lat ~6-37, lng ~68-98. Swap if reversed.
    if not (6 <= lat <= 38) and (6 <= lng <= 38):
        lat, lng = lng, lat
    return lat, lng

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

def clean(v):
    if pd.isna(v):
        return ""
    if isinstance(v, pd.Timestamp):
        return v.strftime("%d %b %Y")
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v

MICROMARKET_LOCATION = {
    # keyword (lowercase) : (city, state)
    "churchgate": ("Mumbai", "Maharashtra"),
    "mulund":     ("Mumbai", "Maharashtra"),
    "dombivli":   ("Mumbai", "Maharashtra"),
    "wadala":     ("Mumbai", "Maharashtra"),
    "andheri":    ("Mumbai", "Maharashtra"),
    "bandra":     ("Mumbai", "Maharashtra"),
    "powai":      ("Mumbai", "Maharashtra"),
    "attapur":    ("Hyderabad", "Telangana"),
    "hyderguda":  ("Hyderabad", "Telangana"),
    "tolichowki": ("Hyderabad", "Telangana"),
    "shaikpet":   ("Hyderabad", "Telangana"),
    "jubilee":    ("Hyderabad", "Telangana"),
}

def lookup_state_city(micromarket, lat, lng):
    text = str(micromarket or "").lower()

    # 1. keyword match on the Micromarket column
    for keyword, (city, state) in MICROMARKET_LOCATION.items():
        if keyword in text:
            return state, city

    # 2. optional fallback: reverse geocode from coordinates
    if lat is not None and lng is not None:
        try:
            import reverse_geocoder as rg
            res = rg.search((lat, lng), mode=1)[0]
            return res["admin1"], res["name"]
        except Exception:
            pass

    # 3. unknown
    return "", ""

def load_excel():
    df = pd.read_excel(EXCEL_PATH)
    df.columns = [str(c).strip() for c in df.columns]

    rows = []
    for _, r in df.iterrows():
        details = {col: clean(r.get(col, "")) for col in COLUMNS}
        lat, lng = parse_coords(details["Google Co-ordiantes"])
        state, city = lookup_state_city(details["Micromarket"], lat, lng)

        rows.append({
            "client_name": details["Client Name"],
            "locality": details["Micromarket"],
            "city": city,
            "state": state,
            "lat": lat,
            "lng": lng,
            # replacements for the old rate/size, if cards or filters use them
            "rate": f"{details['M.V. per acre']} INRMn/acre",
            "size": f"{details['Total Land Area (Acres)']} acres",
            "address": details["Micromarket"],
            "filename": f"{details['Sr. no.']}_{details['Client Name']}",
            "is_saved": False,
            "details": details,
            "comparables": [],
        })

    return rows

@app.route("/api/properties")
def get_properties():
    data = load_excel()
    for d in data:
        d["is_saved"] = d.get("filename") in _saved_filenames
    return jsonify(data)

@app.route("/api/state-coords")
def get_state_coords():
    return jsonify(STATE_COORDS)

@app.route("/api/verify-password", methods=["POST"])
def verify_password():
    body = request.json
    if body.get("password") == ADMIN_PASSWORD:
        return jsonify({"success": True})
    return jsonify({"success": False, "message": "Incorrect password"}), 401

@app.route("/api/save-property", methods=["POST"])
def save_property():
    prop = request.json
    filename = prop.get("filename")
    if filename:
        _saved_filenames.add(filename)
    saved = load_saved()
    existing = next((s for s in saved if s.get("filename") == filename), None)
    if existing:
        saved[saved.index(existing)] = prop
    else:
        saved.append(prop)
    save_saved(saved)
    return jsonify({"success": True})

@app.route("/api/saved-properties")
def get_saved_properties():
    return jsonify(load_saved())

@app.route("/api/refresh", methods=["POST"])
def refresh():
    if os.path.exists(DATA_FILE):
        os.remove(DATA_FILE)
    data = process_all_files(UPLOAD_FOLDER)
    with open(DATA_FILE, "w") as f:
        json.dump(data, f, indent=2)
    return jsonify({"success": True, "count": len(data)})

@app.route("/api/download-excel", methods=["GET", "POST"])
def download_excel():
    rows = load_excel()   # same source the map uses

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Properties"

    # Styles
    header_fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=12)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    normal_font = Font(name="Calibri", size=11)
    alt_fill = PatternFill(start_color="FFF2F2", end_color="FFF2F2", fill_type="solid")
    normal_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Headers (row 1)
    for col_idx, h in enumerate(COLUMNS, 1):
        c = ws.cell(row=1, column=col_idx, value=h)
        c.font, c.fill, c.alignment, c.border = header_font, header_fill, header_align, border
    ws.row_dimensions[1].height = 32

    # Data rows
    for i, prop in enumerate(rows):
        details = prop["details"]
        fill = alt_fill if i % 2 else normal_fill
        for col_idx, col in enumerate(COLUMNS, 1):
            c = ws.cell(row=i + 2, column=col_idx, value=details.get(col, ""))
            c.font, c.fill, c.alignment, c.border = normal_font, fill, cell_align, border
        ws.row_dimensions[i + 2].height = 24

    # Column widths
    widths = [8, 16, 28, 26, 24, 16, 20, 20, 20, 16, 22]
    for col_idx, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.freeze_panes = "A2"

    # Summary by State sheet
    ws2 = wb.create_sheet("Summary by State")
    for col_idx, h in enumerate(["State", "Number of Properties"], 1):
        c = ws2.cell(row=1, column=col_idx, value=h)
        c.font, c.fill, c.alignment, c.border = header_font, header_fill, header_align, border
    counts = {}
    for p in rows:
        s = p.get("state") or "Unknown"
        counts[s] = counts.get(s, 0) + 1
    for state, n in sorted(counts.items()):
        ws2.append([state, n])
    ws2.column_dimensions["A"].width = 25
    ws2.column_dimensions["B"].width = 25

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"RealEstate_Properties_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


#pdf route (depricated)

#def load_extracted():
#     if os.path.exists(DATA_FILE):
#         try:
#             with open(DATA_FILE) as f:
#                 content = f.read().strip()
#                 if not content:
#                     return []
#                 return json.loads(content)
#         except (json.JSONDecodeError, Exception):
#             return []
#     return []

# @app.route("/api/upload", methods=["POST"])
# def upload_file():
#     if "file" not in request.files:
#         return jsonify({"error": "No file"}), 400
#     f = request.files["file"]
#     if f.filename == "":
#         return jsonify({"error": "Empty filename"}), 400
#     allowed = {".pdf", ".docx", ".doc"}
#     ext = os.path.splitext(f.filename)[1].lower()
#     if ext not in allowed:
#         return jsonify({"error": "Only PDF/DOCX/DOC allowed"}), 400
#     filepath = os.path.join(UPLOAD_FOLDER, f.filename)
#     f.save(filepath)
#     from extractor import process_file
#     new_data = process_file(filepath)
#     existing = load_extracted()
#     existing = [e for e in existing if e.get("filename") != f.filename]
#     existing.append(new_data)
#     with open(DATA_FILE, "w") as fp:
#         json.dump(existing, fp, indent=2)
#     return jsonify({"success": True, "filename": f.filename})


@app.route("/")
def serve_index():
    # Serves index.html from the root folder
    return send_from_directory("./static", "index.html")

if __name__ == "__main__":
    print("🚀 Starting Real Estate Map Server...")
    print(f"📁 Upload folder: {os.path.abspath(UPLOAD_FOLDER)}")
    app.run(debug=False, host="0.0.0.0", port=5000)