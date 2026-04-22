from flask import Flask, request, jsonify, send_file
import base64, io, os, uuid, time, threading
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

app = Flask(__name__)

IMG_PX = 52
ROW_PT = 42

# Geçici dosya deposu: { token: { "path": ..., "filename": ..., "created_at": ... } }
_file_store = {}
_store_lock = threading.Lock()
TEMP_DIR = "/tmp/excel_files"
os.makedirs(TEMP_DIR, exist_ok=True)

def _cleanup_old_files(max_age_seconds=600):
    """10 dakikadan eski geçici dosyaları sil"""
    now = time.time()
    with _store_lock:
        expired = [t for t, v in _file_store.items() if now - v["created_at"] > max_age_seconds]
        for token in expired:
            try:
                os.remove(_file_store[token]["path"])
            except Exception:
                pass
            del _file_store[token]

def header_style(cell):
    cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    cell.fill      = PatternFill("solid", start_color="1F2D4E")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    s = Side(style="thin", color="FFFFFF")
    cell.border    = Border(left=s, right=s, bottom=s)

def row_style(cell, row_no):
    cell.fill      = PatternFill("solid", start_color="F5F7FA" if row_no % 2 == 0 else "FFFFFF")
    cell.font      = Font(name="Arial", size=9)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    s = Side(style="thin", color="D0D7E3")
    cell.border    = Border(left=s, right=s, top=s, bottom=s)

def resize_b64_image(b64_str, max_px=56):
    raw = base64.b64decode(b64_str)
    img = PILImage.open(io.BytesIO(raw)).convert("RGB")
    img.thumbnail((max_px, max_px), PILImage.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    buf.seek(0)
    return buf

COLUMNS = [
    ("Resim",      None,        8),
    ("Urun ID",    "prdID",    14),
    ("Kalem ID",   "itmID",    14),
    ("Etiket",     "tagPrice", 12),
    ("Satis",      "salePrice",12),
    ("Indirim %",  "discount", 10),
    ("Tarih",      "date",     14),
    ("Lokasyon",   "location", 16),
]

def build_workbook(rows):
    """Ortak workbook oluşturma fonksiyonu"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Satis Raporu"

    for col_no, (label, _, width) in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(col_no)].width = width
        cell = ws.cell(row=1, column=col_no, value=label)
        header_style(cell)
    ws.row_dimensions[1].height = 22

    for row_no, item in enumerate(rows, start=2):
        ws.row_dimensions[row_no].height = ROW_PT

        for col_no, (_, key, _) in enumerate(COLUMNS, start=1):
            if key is None:
                b64 = item.get("image_base64", "")
                if b64:
                    try:
                        buf = resize_b64_image(b64, max_px=IMG_PX)
                        xl_img = XLImage(buf)
                        xl_img.width  = IMG_PX
                        xl_img.height = IMG_PX
                        offset = pixels_to_EMU(2)
                        size   = XDRPositiveSize2D(pixels_to_EMU(IMG_PX), pixels_to_EMU(IMG_PX))
                        marker = AnchorMarker(col=0, colOff=offset, row=row_no-1, rowOff=offset)
                        xl_img.anchor = OneCellAnchor(_from=marker, ext=size)
                        ws.add_image(xl_img)
                    except Exception:
                        ws.cell(row=row_no, column=1, value="-")
                else:
                    ws.cell(row=row_no, column=1, value="-")
            else:
                cell = ws.cell(row=row_no, column=col_no, value=item.get(key, ""))
                row_style(cell, row_no)

    ws.freeze_panes = "A2"
    last_col = get_column_letter(len(COLUMNS))
    ws.auto_filter.ref = f"A1:{last_col}1"
    return wb

def parse_and_validate(request):
    body = request.get_json(force=True, silent=True)
    if not body or "data" not in body:
        return None, None
    rows = [r for r in body["data"] if any(str(v).strip() for k, v in r.items() if k != "image_base64")]
    return body, rows


# ── Mevcut endpoint (FM desktop için — base64 döndürür) ──────────────────────

@app.route("/generate-excel", methods=["POST"])
def generate_excel():
    body, rows = parse_and_validate(request)
    if body is None:
        return jsonify({"error": "JSON body eksik"}), 400
    if not rows:
        return jsonify({"error": "data listesi bos"}), 400

    try:
        wb = build_workbook(rows)
        
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        
        encoded = base64.b64encode(out.read()).decode("utf-8")
        
        return jsonify({
            "status": "ok",
            "rows": len(rows),
            "excel_b64": encoded,
            "filename": "satis_raporu.xlsx"
        })
    except Exception as e:
        return jsonify({"error": f"Excel oluşturma hatası: {str(e)}"}), 500

# ── Yeni endpoint (WebDirect için — download URL döndürür) ───────────────────






# ── Health ───────────────────────────────────────────────────────────────────
@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
