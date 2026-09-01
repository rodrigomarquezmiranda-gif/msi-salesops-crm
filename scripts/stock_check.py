"""
MSI SalesOps — Stock & Price Check + Email Notification
Corre vía GitHub Actions todos los días hábiles a las 18:30 ART.
Lee el pricelist desde Firebase, compara con el snapshot anterior,
escribe el diff a Firebase, y manda email con Excel adjunto si hay cambios.
"""
import json, re, sys, os, io, smtplib
from datetime import datetime, timezone
from urllib import request as urlrequest
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# ── Config ────────────────────────────────────────────────────────────────────
FB_BASE    = "https://msi-crm-default-rtdb.firebaseio.com"
FB_SECRET  = os.environ.get("FIREBASE_SECRET", "")
FB_PARAMS  = f"?auth={FB_SECRET}" if FB_SECRET else ""

SMTP_HOST  = "smtp.hostinger.com"
SMTP_PORT  = 465
SMTP_USER  = "sales@msicrm.com"
SMTP_PASS  = os.environ.get("SMTP_PASSWORD", "")
RECIPIENTS = [e.strip() for e in os.environ.get("RECIPIENT_EMAILS", "").split(",") if e.strip()]

# ── Firebase helpers ──────────────────────────────────────────────────────────
def fb_get(path):
    url = f"{FB_BASE}/{path}.json{FB_PARAMS}"
    try:
        with urlrequest.urlopen(url, timeout=30) as r:
            return json.loads(r.read())
    except Exception as e:
        print(f"[WARN] GET {path}: {e}")
        return None

def fb_put(path, data):
    url = f"{FB_BASE}/{path}.json{FB_PARAMS}"
    req = urlrequest.Request(url, data=json.dumps(data).encode(),
                             method="PUT", headers={"Content-Type":"application/json"})
    try:
        with urlrequest.urlopen(req, timeout=30) as r:
            print(f"[OK]   PUT {path} → {r.status}"); return True
    except Exception as e:
        print(f"[ERR]  PUT {path}: {e}"); return False

def fb_post(path, data):
    url = f"{FB_BASE}/{path}.json{FB_PARAMS}"
    req = urlrequest.Request(url, data=json.dumps(data).encode(),
                             method="POST", headers={"Content-Type":"application/json"})
    try:
        with urlrequest.urlopen(req, timeout=30) as r:
            print(f"[OK]   POST {path} → {r.status}"); return True
    except Exception as e:
        print(f"[ERR]  POST {path}: {e}"); return False

def safe_int(v):
    try: m=re.match(r'(\d+)',str(v).strip()); return int(m.group(1)) if m else 0
    except: return 0

def safe_float(v):
    try: return float(str(v).replace('$','').replace(',','').strip())
    except: return 0.0

# ── Generate Excel ────────────────────────────────────────────────────────────
def generate_excel(products_dict, changes):
    """Genera un Excel con la lista de precios y un resumen de cambios."""
    try:
        import openpyxl
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    except ImportError:
        print("[WARN] openpyxl no disponible, saltando Excel"); return None

    wb = openpyxl.Workbook()

    # ── Sheet 1: Lista de precios ──
    ws = wb.active
    ws.title = "Lista de Precios"

    NAVY   = "1F2A44"
    ORANGE = "FF6600"
    WHITE  = "FFFFFF"
    GRAY   = "F5F5F5"

    header_fill   = PatternFill("solid", fgColor=NAVY)
    orange_fill   = PatternFill("solid", fgColor=ORANGE)
    gray_fill     = PatternFill("solid", fgColor=GRAY)
    header_font   = Font(color=WHITE, bold=True, size=10)
    title_font    = Font(color=WHITE, bold=True, size=13)
    thin_border   = Border(bottom=Side(style="thin", color="CCCCCC"))

    # Title row
    ws.merge_cells("A1:F1")
    ws["A1"] = f"MSI Argentina — Lista de Precios  |  {datetime.now().strftime('%d/%m/%Y')}"
    ws["A1"].font = title_font
    ws["A1"].fill = PatternFill("solid", fgColor=ORANGE)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Headers
    headers = ["Part Number", "Nombre", "Categoría", "Precio Miami (USD)", "Stock Miami", "ETA"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    # Data
    col_widths = [18, 50, 16, 20, 14, 16]
    for i, (partNo, p) in enumerate(sorted(products_dict.items(), key=lambda x: x[1].get("category",""))):
        row = i + 3
        fill = gray_fill if i % 2 == 0 else None
        values = [partNo, p.get("name",""), p.get("category",""),
                  p.get("miamiPrice",0), p.get("miamiStock",0), p.get("miamiEta","")]
        for col, val in enumerate(values, 1):
            cell = ws.cell(row=row, column=col, value=val)
            if fill: cell.fill = fill
            cell.border = thin_border
            if col == 4 and isinstance(val, (int, float)):
                cell.number_format = '"$"#,##0.00'

    for col, width in enumerate(col_widths, 1):
        ws.column_dimensions[ws.cell(row=2, column=col).column_letter].width = width

    # ── Sheet 2: Cambios detectados ──
    if changes:
        ws2 = wb.create_sheet("Cambios")
        ws2.merge_cells("A1:E1")
        ws2["A1"] = f"Cambios detectados — {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        ws2["A1"].font = title_font
        ws2["A1"].fill = PatternFill("solid", fgColor=ORANGE)
        ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws2.row_dimensions[1].height = 24

        h2 = ["Tipo", "Part Number", "Nombre", "Valor anterior", "Valor nuevo"]
        for col, h in enumerate(h2, 1):
            cell = ws2.cell(row=2, column=col, value=h)
            cell.fill = header_fill; cell.font = header_font

        type_labels = {"price":"Precio","stock":"Stock","new":"Nuevo producto","removed":"Dado de baja"}
        for i, c in enumerate(changes):
            row = i + 3
            t = type_labels.get(c.get("type",""), c.get("type",""))
            old_val = c.get("oldPrice", c.get("oldStock", "—"))
            new_val = c.get("newPrice", c.get("newStock", "—"))
            for col, val in enumerate([t, c.get("partNo",""), c.get("name",""), old_val, new_val], 1):
                ws2.cell(row=row, column=col, value=val)

        for col, width in enumerate([14, 18, 46, 16, 16], 1):
            ws2.column_dimensions[ws2.cell(row=2, column=col).column_letter].width = width

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ── Send email ────────────────────────────────────────────────────────────────
def send_email(summary, changes, excel_bytes, date_str):
    if not RECIPIENTS:
        print("[SKIP] No recipients configured"); return
    if not SMTP_PASS:
        print("[SKIP] SMTP_PASSWORD not set"); return

    msg = MIMEMultipart()
    msg["From"]    = f"MSI Argentina <{SMTP_USER}>"
    msg["To"]      = ", ".join(RECIPIENTS)
    msg["Subject"] = f"MSI Argentina — Lista de Precios Actualizada ({date_str})"
    msg["Reply-To"] = SMTP_USER

    change_lines = ""
    if changes:
        labels = {"price":"Precio","stock":"Stock","new":"Nuevo","removed":"Baja"}
        for c in changes[:20]:
            t = labels.get(c.get("type",""), c.get("type",""))
            if c["type"] == "price":
                change_lines += f"  • [{t}] {c['partNo']} — {c['name']}: ${c.get('oldPrice')} → ${c.get('newPrice')}\n"
            elif c["type"] == "stock":
                change_lines += f"  • [{t}] {c['partNo']} — {c['name']}: {c.get('oldStock')} → {c.get('newStock')} unidades\n"
            elif c["type"] == "new":
                change_lines += f"  • [{t}] {c['partNo']} — {c['name']}: ${c.get('newPrice')}\n"
            elif c["type"] == "removed":
                change_lines += f"  • [{t}] {c['partNo']} — {c['name']}\n"
        if len(changes) > 20:
            change_lines += f"  ... y {len(changes)-20} cambios más (ver Excel adjunto)\n"

    body = f"""Estimados,

Les informamos que la lista de precios MSI Argentina ha sido actualizada.

{summary}
{change_lines}
Adjunto encontrarán la lista completa actualizada en formato Excel.

Saludos,
MSI Argentina
sales@msicrm.com
"""
    msg.attach(MIMEText(body, "plain", "utf-8"))

    if excel_bytes:
        filename = f"MSI_Lista_Precios_{date_str.replace('/','-')}.xlsx"
        part = MIMEBase("application", "octet-stream")
        part.set_payload(excel_bytes)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
        msg.attach(part)

    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(SMTP_USER, RECIPIENTS, msg.as_bytes())
        print(f"[OK]   Email enviado a {len(RECIPIENTS)} destinatario(s)")
    except Exception as e:
        print(f"[ERR]  Email failed: {e}")

# ── Main ──────────────────────────────────────────────────────────────────────
print("=== MSI Stock & Price Check ===")
now_iso  = datetime.now(timezone.utc).isoformat()
date_str = datetime.now().strftime("%d/%m/%Y")

# 1. Load pricelist
print("Loading pricelist from Firebase...")
pricelist_raw = fb_get("salesops_pricelist")
if not pricelist_raw:
    local_path = os.path.join(os.path.dirname(__file__), '..', 'pricelist_snapshot.json')
    if os.path.exists(local_path):
        with open(local_path) as f: pricelist_raw = json.load(f)
        print("Fallback: local pricelist_snapshot.json")
    else:
        print("[ERROR] No pricelist available. Exiting."); sys.exit(1)

if isinstance(pricelist_raw, str): pricelist_raw = json.loads(pricelist_raw)
raw_products = pricelist_raw.get("products", {})
print(f"Pricelist: {len(raw_products)} products")

# 2. Build new snapshot
new_snapshot = {}
for partNo, p in raw_products.items():
    new_snapshot[partNo] = {
        "name":        p.get("name",""),
        "category":    p.get("category",""),
        "miamiStock":  safe_int(p.get("miamiStock",0)),
        "miamiPrice":  safe_float(p.get("price",0)),
        "miamiEta":    p.get("miamiEta",""),
        "bondedStock": safe_int(p.get("bondedStock",0)),
    }

# 3. Load old snapshot
old_raw = fb_get("salesops_stock_snapshot")
if old_raw and isinstance(old_raw, str): old_raw = json.loads(old_raw)
old_snapshot = old_raw if isinstance(old_raw, dict) else {}
print(f"Previous snapshot: {len(old_snapshot)} products")

# 4. Compute diff
changes = []
for partNo in set(new_snapshot) | set(old_snapshot):
    new = new_snapshot.get(partNo)
    old = old_snapshot.get(partNo)
    if new and not old:
        changes.append({"type":"new","partNo":partNo,"name":new["name"],"category":new["category"],"newPrice":new["miamiPrice"],"newStock":new["miamiStock"]})
    elif old and not new:
        changes.append({"type":"removed","partNo":partNo,"name":old["name"],"category":old["category"],"oldPrice":old["miamiPrice"]})
    elif new and old:
        if abs(new["miamiPrice"] - old["miamiPrice"]) >= 0.5:
            changes.append({"type":"price","partNo":partNo,"name":new["name"],"category":new["category"],"oldPrice":old["miamiPrice"],"newPrice":new["miamiPrice"]})
        if abs(new["miamiStock"] - old["miamiStock"]) >= 1:
            changes.append({"type":"stock","partNo":partNo,"name":new["name"],"category":new["category"],"oldStock":old["miamiStock"],"newStock":new["miamiStock"]})

has_changes = len(changes) > 0
summary = f"{len(changes)} cambio(s) detectado(s)" if has_changes else "Sin cambios"
print(f"Changes: {summary}")

# 5. Build payloads
diff_payload = {
    "checkedAt": now_iso, "productCount": len(new_snapshot),
    "hasChanges": has_changes, "changeCount": len(changes),
    "changes": changes[:100], "summary": summary,
    "triggeredBy": "github-actions", "triggerMode": "auto",
}
changelog_entry = {
    "checkedAt": now_iso, "productCount": len(new_snapshot),
    "hasChanges": has_changes, "changeCount": len(changes),
    "summary": summary, "triggerMode": "auto", "triggeredBy": "github-actions",
}

# 6. Write to Firebase
print("Writing to Firebase...")
ok1 = fb_put("salesops_stock_diff",     json.dumps(diff_payload))
ok2 = fb_put("salesops_stock_snapshot", json.dumps(new_snapshot))
ok3 = fb_post("salesops_stock_changelog", changelog_entry)

# 7. Update local files
repo_root = os.path.join(os.path.dirname(__file__), '..')
with open(os.path.join(repo_root, 'stock_diff_latest.json'), 'w') as f:
    json.dump(diff_payload, f, ensure_ascii=False, indent=2)
with open(os.path.join(repo_root, 'stock_snapshot_latest.json'), 'w') as f:
    json.dump(new_snapshot, f, ensure_ascii=False, indent=2)

# 8. Send email if changes detected
if has_changes:
    print("Generating Excel and sending email...")
    excel_bytes = generate_excel(new_snapshot, changes)
    send_email(summary, changes, excel_bytes, date_str)
else:
    print("No changes — skipping email")

# Summary
print(f"\n{'='*40}")
print(f"RESULT: {summary}")
print(f"Firebase: diff={'OK' if ok1 else 'ERR'} snapshot={'OK' if ok2 else 'ERR'} changelog={'OK' if ok3 else 'ERR'}")
if not (ok1 and ok2 and ok3): sys.exit(1)
