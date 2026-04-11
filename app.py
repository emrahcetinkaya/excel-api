from flask import Flask, request, jsonify
import base64, io, os
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

app = Flask(__name__)

# ── Stil yardımcıları ──────────────────────────────────────────────────────────

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

def resize_b64_image(b64_str, max_px=80):
    """Base64 resmi küçült, yeni base64 döndür."""
    raw = base64.b64decode(b64_str)
    img = PILImage.open(io.BytesIO(raw)).convert("RGB")
    img.thumbnail((max_px, max_px), PILImage.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    buf.seek(0)
    return buf

# ── Sütun tanımları ────────────────────────────────────────────────────────────
COLUMNS = [
    ("Resim",      None,           8),
    ("Ürün ID",    "prdID",        14),
    ("Kalem ID",   "itmID",        14),
    ("Etiket",     "tagPrice",     12),
    ("Satış",      "salePrice",    12),
    ("İndirim %",  "discount",     10),
    ("Tarih",      "date",         14),
    ("Lokasyon",   "location",     16),
]

# ── Ana endpoint ───────────────────────────────────────────────────────────────

@app.route("/generate-excel", methods=["POST"])
def generate_excel():
    body = request.get_json(force=True, silent=True)
    if not body or "data" not in body:
        return jsonify({"error": "JSON body eksik veya 'data' anahtarı yok"}), 400

    rows = body["data"]
    if not isinstance(rows, list) or len(rows) == 0:
        return jsonify({"error": "data listesi boş"}), 400

    wb = Workbook()
    ws = wb.active
    ws.title = "Satış Raporu"

    IMG_H = 60   # piksel yükseklik
    IMG_W = 60   # piksel genişlik

    # Başlık satırı
    for col_no, (label, _, width) in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(col_no)].width = width
        cell = ws.cell(row=1, column=col_no, value=label)
        header_style(cell)
    ws.row_dimensions[1].height = 22

    # Veri satırları
    for row_no, item in enumerate(rows, start=2):
        ws.row_dimensions[row_no].height = IMG_H * 0.75

        for col_no, (_, key, _) in enumerate(COLUMNS, start=1):
            if key is None:
                # Resim sütunu
                b64 = item.get("image_base64", "")
                if b64:
                    try:
                        buf = resize_b64_image(b64, max_px=IMG_H)
                        img = XLImage(buf)
                        img.width  = IMG_W
                        img.height = IMG_H
                        ws.add_image(img, f"A{row_no}")
                    except Exception:
                        ws.cell(row=row_no, column=1, value="—")
                else:
                    ws.cell(row=row_no, column=1, value="—")
            else:
                cell = ws.cell(row=row_no, column=col_no, value=item.get(key, ""))
                row_style(cell, row_no)

    # Başlığı dondur + filtre
    ws.freeze_panes = "A2"
    last_col = get_column_letter(len(COLUMNS))
    ws.auto_filter.ref = f"A1:{last_col}1"

    # Excel → base64
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    encoded = base64.b64encode(out.read()).decode("utf-8")

    return jsonify({
        "status":   "ok",
        "rows":     len(rows),
        "excel_b64": encoded
    })

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
