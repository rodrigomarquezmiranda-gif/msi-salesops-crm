"""
MSI SalesOps — Stock & Price Check + Email Notification
Corre vía GitHub Actions todos los días hábiles a las 18:30 ART.
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

SMTP_HOST      = "smtp.hostinger.com"
SMTP_PORT      = 465
SMTP_USER      = "sales@msicrm.com"
SMTP_PASS      = os.environ.get("SMTP_PASSWORD", "")
FORCE_EMAIL    = os.environ.get("FORCE_EMAIL", "false").lower() == "true"
TEST_RECIPIENT = os.environ.get("TEST_RECIPIENT", "").strip()
_raw_recipients = TEST_RECIPIENT if TEST_RECIPIENT else os.environ.get("RECIPIENT_EMAILS", "")
RECIPIENTS = [e.strip() for e in _raw_recipients.split(",") if e.strip()]

# ── Firebase helpers ──────────────────────────────────────────────────────────
def fb_get(path):
    url = f"{FB_BASE}/{path}.json{FB_PARAMS}"
    try:
        with urlrequest.urlopen(url, timeout=30) as r:
            return json.loads(r.read())
    except Exception as e:
        print(f"[WARN] GET {path}: {e}"); return None

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
    try: m = re.match(r'(\d+)', str(v).strip()); return int(m.group(1)) if m else 0
    except: return 0

def safe_float(v):
    try: return float(str(v).replace('$','').replace(',','').strip())
    except: return 0.0

# ── Generate Excel (same format as CRM export) ────────────────────────────────
def generate_excel(raw_products, changes, date_str):
    try:
        import openpyxl
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    except ImportError:
        print("[WARN] openpyxl no disponible"); return None

    NAVY   = "1F2A44"
    ORANGE = "FF6600"
    WHITE  = "FFFFFF"
    LGRAY  = "F0F0F0"
    CATBG  = "2A3550"

    wb = openpyxl.Workbook()

    # ════════════════════════════════════════════════════════
    # Sheet 1 — Lista de Precios (formato idéntico al CRM)
    # ════════════════════════════════════════════════════════
    ws = wb.active
    ws.title = "Lista de Precios"

    title_font  = Font(color=WHITE, bold=True, size=13)
    note_font   = Font(color="555555", italic=True, size=9)
    hdr_fill    = PatternFill("solid", fgColor=NAVY)
    hdr_font    = Font(color=WHITE, bold=True, size=9)
    cat_fill    = PatternFill("solid", fgColor=CATBG)
    cat_font    = Font(color=ORANGE, bold=True, size=10)
    gray_fill   = PatternFill("solid", fgColor=LGRAY)
    thin        = Border(bottom=Side(style="thin", color="DDDDDD"))
    center      = Alignment(horizontal="center", vertical="center")
    NCOLS       = 13

    # Row 1 — Title
    ws.merge_cells(f"A1:M1")
    ws["A1"] = f"MSI LATAM — Lista de Precios y Stock"
    ws["A1"].font = title_font
    ws["A1"].fill = PatternFill("solid", fgColor=ORANGE)
    ws["A1"].alignment = center
    ws.row_dimensions[1].height = 26

    # Row 2 — Date
    ws.merge_cells("A2:M2")
    ws["A2"] = f"Fecha de referencia: {date_str}"
    ws["A2"].font = Font(color="333333", size=9)
    ws["A2"].alignment = Alignment(horizontal="left")

    # Rows 3-7 — Notes
    notes = [
        None,
        "Notas",
        "Miami Stock: físicamente en Miami.   Miami Bonded Stock: físicamente en Miami Bonded / FTZ.   ETA: llegada estimada.",
        "Precio: se toma el valor de la columna Miami Price; si existe un precio de Market Share menor, se usa el más bajo de los dos.",
        f"Exportado desde SalesOps CRM · {date_str}",
    ]
    for i, note in enumerate(notes, 3):
        if note:
            ws.merge_cells(f"A{i}:M{i}")
            ws[f"A{i}"] = note
            ws[f"A{i}"].font = note_font if i > 4 else Font(bold=True, size=9)

    # Row 9 — Headers
    HEADERS = ["UPC","N° de Parte","Producto","Detalle","Estado","Links",
               "Qty/Ctn","Qty/Plt","Miami Stock","Miami ETA","Miami Bonded Stock","Miami Bonded ETA","Precio"]
    for col, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=9, column=col, value=h)
        cell.fill = hdr_fill; cell.font = hdr_font; cell.alignment = center
    ws.row_dimensions[9].height = 18

    COL_W = [14, 20, 44, 36, 8, 8, 9, 9, 13, 11, 18, 16, 12]
    for col, w in enumerate(COL_W, 1):
        ws.column_dimensions[ws.cell(row=9, column=col).column_letter].width = w

    # Group products by category, preserving order
    from collections import OrderedDict
    by_cat = OrderedDict()
    for partNo, p in raw_products.items():
        cat = p.get("category") or p.get("cat") or "Other"
        by_cat.setdefault(cat, []).append((partNo, p))

    cur_row = 10
    row_idx = 0
    for cat, items in by_cat.items():
        # Category header row
        ws.merge_cells(f"A{cur_row}:M{cur_row}")
        ws[f"A{cur_row}"] = cat
        ws[f"A{cur_row}"].fill = cat_fill
        ws[f"A{cur_row}"].font = cat_font
        ws[f"A{cur_row}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[cur_row].height = 16
        cur_row += 1

        for partNo, p in items:
            fill = gray_fill if row_idx % 2 == 0 else None
            price = safe_float(p.get("price", p.get("priceR", 0)))
            vals = [
                p.get("upc",""),
                partNo,
                p.get("name",""),
                p.get("detail", p.get("marketing", p.get("description",""))),
                p.get("status", p.get("estado","")),
                "",  # Links — not available via API
                p.get("qtyCtn",""),
                p.get("qtyPlt",""),
                safe_int(p.get("miamiStock",0)),
                p.get("miamiEta",""),
                safe_int(p.get("bondedStock",0)),
                p.get("bondedEta",""),
                price,
            ]
            for col, val in enumerate(vals, 1):
                cell = ws.cell(row=cur_row, column=col, value=val)
                if fill: cell.fill = fill
                cell.border = thin
                cell.alignment = Alignment(vertical="center",
                    horizontal="center" if col in (1,6,7,8,9,10,11,12,13) else "left")
                if col == 13 and isinstance(val, (int,float)) and val:
                    cell.number_format = '"$"#,##0.00'
                if col == 9 or col == 11:
                    cell.number_format = '#,##0'
            cur_row += 1
            row_idx += 1

    # Freeze panes below header
    ws.freeze_panes = "A10"

    # ════════════════════════════════════════════════════════
    # Sheet 2 — Cambios (only if any)
    # ════════════════════════════════════════════════════════
    if changes:
        ws2 = wb.create_sheet("Cambios")

        ws2.merge_cells("A1:E1")
        ws2["A1"] = f"Cambios detectados — {date_str}"
        ws2["A1"].font = title_font
        ws2["A1"].fill = PatternFill("solid", fgColor=ORANGE)
        ws2["A1"].alignment = center
        ws2.row_dimensions[1].height = 24

        h2 = ["Tipo","N° de Parte","Producto","Valor anterior","Valor nuevo"]
        for col, h in enumerate(h2, 1):
            cell = ws2.cell(row=2, column=col, value=h)
            cell.fill = hdr_fill; cell.font = hdr_font; cell.alignment = center

        type_labels = {"price":"Precio","stock":"Stock","new":"Nuevo","removed":"Retirado"}
        TYPE_COLORS = {"new":"00875A","removed":"6B7280","price":"0057B8","stock":"444444"}
        for i, c in enumerate(changes):
            row = i + 3
            t   = type_labels.get(c.get("type",""), c.get("type",""))
            old = c.get("oldPrice", c.get("oldStock","—"))
            new = c.get("newPrice", c.get("newStock","—"))
            if isinstance(old, float): old = f"${old:,.2f}"
            if isinstance(new, float): new = f"${new:,.2f}"
            vals = [t, c.get("partNo",""), c.get("name",""), old, new]
            clr  = TYPE_COLORS.get(c.get("type",""),"333333")
            for col, val in enumerate(vals, 1):
                cell = ws2.cell(row=row, column=col, value=val)
                if i % 2 == 0: cell.fill = gray_fill
                cell.border = thin
                if col == 1: cell.font = Font(color=clr, bold=True, size=9)

        for col, w in enumerate([12,20,46,16,16],1):
            ws2.column_dimensions[ws2.cell(row=2,column=col).column_letter].width = w

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf.read()


# ── Build HTML email body ─────────────────────────────────────────────────────
def build_html_email(summary_text, changes, date_str, product_count):
    n_new      = sum(1 for c in changes if c["type"]=="new")
    n_removed  = sum(1 for c in changes if c["type"]=="removed")
    n_price_up = sum(1 for c in changes if c["type"]=="price" and c.get("newPrice",0) > c.get("oldPrice",0))
    n_price_dn = sum(1 for c in changes if c["type"]=="price" and c.get("newPrice",0) < c.get("oldPrice",0))
    n_stock    = sum(1 for c in changes if c["type"]=="stock")
    n_total    = len(changes)

    def pill(label, color, bg):
        return f'<span style="display:inline-block;background:{bg};color:{color};border-radius:4px;padding:2px 8px;font-size:11px;font-weight:700;margin-right:4px">{label}</span>'

    summary_pills = ""
    if n_price_dn: summary_pills += pill(f"▼ {n_price_dn} bajas","#fff","#166534")
    if n_price_up: summary_pills += pill(f"▲ {n_price_up} subas","#fff","#7f1d1d")
    if n_new:      summary_pills += pill(f"+ {n_new} nuevos","#fff","#1e3a5f")
    if n_removed:  summary_pills += pill(f"✕ {n_removed} retirados","#fff","#374151")
    if n_stock:    summary_pills += pill(f"~ {n_stock} stock","#fff","#44337a")

    # Changes table rows
    rows_html = ""
    TYPE_META = {
        "new":     ("#fff","#166534","NUEVO"),
        "removed": ("#fff","#4b5563","RETIRADO"),
        "price":   ("#fff","#1e3a5f","PRECIO"),
        "stock":   ("#fff","#44337a","STOCK"),
    }
    # Show price & new changes first, limit to 50
    priority = sorted(changes, key=lambda c: (0 if c["type"] in ("new","price") else 1, c.get("category","")))
    for i, c in enumerate(priority[:50]):
        meta = TYPE_META.get(c["type"], ("#fff","#555","?"))
        badge = f'<span style="background:{meta[1]};color:{meta[0]};border-radius:3px;padding:1px 6px;font-size:10px;font-weight:700">{meta[2]}</span>'
        bg = "#f9fafb" if i % 2 == 0 else "#ffffff"

        if c["type"] == "price":
            old_v = f'<span style="color:#9ca3af;text-decoration:line-through">${c.get("oldPrice",0):,.2f}</span>'
            new_v = f'<span style="color:#{"166534" if c.get("newPrice",0)<c.get("oldPrice",0) else "7f1d1d"};font-weight:700">${c.get("newPrice",0):,.2f}</span>'
        elif c["type"] == "new":
            old_v = "—"
            new_v = f'<span style="color:#166534;font-weight:700">${c.get("newPrice",0):,.2f}</span>'
        elif c["type"] == "removed":
            old_v = f'${c.get("oldPrice",0):,.2f}'
            new_v = '<span style="color:#9ca3af">—</span>'
        else:  # stock
            old_v = str(c.get("oldStock","—"))
            new_v = str(c.get("newStock","—"))

        rows_html += f"""
        <tr style="background:{bg}">
          <td style="padding:6px 10px;font-size:11px;color:#6b7280">{c.get("partNo","")}</td>
          <td style="padding:6px 10px;font-size:12px;color:#111">{c.get("name","")}</td>
          <td style="padding:6px 10px;font-size:11px;color:#6b7280;text-align:center">{c.get("category","")}</td>
          <td style="padding:6px 10px;font-size:12px;text-align:right">{old_v}</td>
          <td style="padding:6px 10px;font-size:12px;text-align:right">{new_v}</td>
          <td style="padding:6px 10px;text-align:center">{badge}</td>
        </tr>"""

    more = f'<p style="text-align:center;color:#6b7280;font-size:11px">… y {n_total-50} cambios más en el Excel adjunto</p>' if n_total > 50 else ""

    html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f3f4f6;font-family:-apple-system,Helvetica,Arial,sans-serif">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f3f4f6;padding:24px 0">
<tr><td align="center">
<table width="640" cellpadding="0" cellspacing="0" style="max-width:640px;width:100%">

  <!-- Header -->
  <tr><td style="background:#1F2A44;border-radius:10px 10px 0 0;padding:20px 28px">
    <table width="100%"><tr>
      <td><div style="color:#FF6600;font-size:18px;font-weight:700;letter-spacing:-.3px">Alerta de Cambios de Precios</div>
          <div style="color:#8b9ab5;font-size:12px;margin-top:2px">MSI LATAM · SalesOps CRM</div></td>
      <td align="right" style="color:#8b9ab5;font-size:12px;white-space:nowrap">{date_str}</td>
    </tr></table>
  </td></tr>

  <!-- Summary bar -->
  <tr><td style="background:#253047;padding:12px 28px">
    <span style="color:#e2e8f0;font-size:13px">Se detectaron <strong style="color:#fff">{n_total} cambios</strong> en la lista de precios LATAM.
    &nbsp;·&nbsp; {summary_pills}</span>
  </td></tr>

  <!-- Table -->
  <tr><td style="background:#ffffff;padding:0">
    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">
      <tr style="background:#1F2A44">
        <th style="padding:8px 10px;font-size:10px;color:#8b9ab5;font-weight:600;text-align:left">CÓDIGO</th>
        <th style="padding:8px 10px;font-size:10px;color:#8b9ab5;font-weight:600;text-align:left">PRODUCTO</th>
        <th style="padding:8px 10px;font-size:10px;color:#8b9ab5;font-weight:600;text-align:center">CAT.</th>
        <th style="padding:8px 10px;font-size:10px;color:#8b9ab5;font-weight:600;text-align:right">PRECIO ANT.</th>
        <th style="padding:8px 10px;font-size:10px;color:#8b9ab5;font-weight:600;text-align:right">PRECIO NUEVO</th>
        <th style="padding:8px 10px;font-size:10px;color:#8b9ab5;font-weight:600;text-align:center">VAR.</th>
      </tr>
      {rows_html}
    </table>
    {more}
  </td></tr>

  <!-- Attachment note -->
  <tr><td style="background:#f8fafc;border:1px solid #e2e8f0;border-top:none;padding:14px 28px">
    <table><tr>
      <td style="font-size:24px;padding-right:12px">📊</td>
      <td>
        <div style="font-weight:600;font-size:13px;color:#1e293b">MSI LATAM - Lista de Precios - {date_str.replace("/",".")}.xlsx</div>
        <div style="font-size:11px;color:#64748b">Lista completa con {product_count} productos · Hojas: Lista de Precios + Cambios</div>
      </td>
    </tr></table>
  </td></tr>

  <!-- Footer -->
  <tr><td style="background:#1F2A44;border-radius:0 0 10px 10px;padding:14px 28px;text-align:center">
    <span style="color:#8b9ab5;font-size:11px">Enviado automáticamente por MSI SalesOps CRM ·
    <a href="https://msicrm.com" style="color:#FF6600;text-decoration:none">msicrm.com</a></span>
  </td></tr>

</table>
</td></tr></table>
</body></html>"""
    return html


# ── Send email ────────────────────────────────────────────────────────────────
def send_email(summary_text, changes, excel_bytes, date_str, product_count):
    if not RECIPIENTS:
        print("[SKIP] No recipients configured"); return
    if not SMTP_PASS:
        print("[SKIP] SMTP_PASSWORD not set"); return

    msg = MIMEMultipart("alternative")
    msg["From"]         = f"MSI Argentina <{SMTP_USER}>"
    msg["To"]           = SMTP_USER
    msg["Bcc"]          = ", ".join(RECIPIENTS)
    msg["Subject"]      = f"[MSI CRM] Cambios en lista de precios LATAM — {date_str}"
    msg["Reply-To"]     = SMTP_USER
    msg["X-Mailer"]     = "MSI SalesOps CRM"
    msg["Precedence"]   = "bulk"
    msg["Auto-Submitted"] = "auto-generated"

    # Plain text fallback
    plain = f"MSI LATAM — Cambios en lista de precios\n{date_str}\n\n{summary_text}\n\nVer Excel adjunto.\n\nsales@msicrm.com"
    msg.attach(MIMEText(plain, "plain", "utf-8"))

    # HTML body
    html_body = build_html_email(summary_text, changes, date_str, product_count)
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    # Excel attachment — needs a new MIMEMultipart("mixed") wrapper
    outer = MIMEMultipart("mixed")
    outer["From"]         = msg["From"]
    outer["To"]           = msg["To"]
    outer["Bcc"]          = msg["Bcc"]
    outer["Subject"]      = msg["Subject"]
    outer["Reply-To"]     = msg["Reply-To"]
    outer["X-Mailer"]     = msg["X-Mailer"]
    outer["Precedence"]   = msg["Precedence"]
    outer["Auto-Submitted"] = msg["Auto-Submitted"]
    outer.attach(msg)

    if excel_bytes:
        filename = f"MSI LATAM - Lista de Precios - {date_str.replace('/','.') }.xlsx"
        part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(excel_bytes)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
        outer.attach(part)

    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(SMTP_USER, RECIPIENTS, outer.as_bytes())
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

# Firebase may return products as list — normalize to dict
if isinstance(raw_products, list):
    raw_products = {
        (p.get("code") or p.get("partNo") or p.get("sku") or str(i)): p
        for i, p in enumerate(raw_products) if isinstance(p, dict)
    }
print(f"Pricelist: {len(raw_products)} products")

# 2. Build snapshot (lean, for diffing)
new_snapshot = {}
for partNo, p in raw_products.items():
    new_snapshot[partNo] = {
        "name":        p.get("name",""),
        "category":    p.get("category",""),
        "miamiStock":  safe_int(p.get("miamiStock",0)),
        "miamiPrice":  safe_float(p.get("price", p.get("priceR", 0))),
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
    n = new_snapshot.get(partNo)
    o = old_snapshot.get(partNo)
    if n and not o:
        changes.append({"type":"new","partNo":partNo,"name":n["name"],"category":n["category"],"newPrice":n["miamiPrice"],"newStock":n["miamiStock"]})
    elif o and not n:
        changes.append({"type":"removed","partNo":partNo,"name":o["name"],"category":o["category"],"oldPrice":o["miamiPrice"]})
    elif n and o:
        if abs(n["miamiPrice"] - o["miamiPrice"]) >= 0.5:
            changes.append({"type":"price","partNo":partNo,"name":n["name"],"category":n["category"],"oldPrice":o["miamiPrice"],"newPrice":n["miamiPrice"]})
        if abs(n["miamiStock"] - o["miamiStock"]) >= 1:
            changes.append({"type":"stock","partNo":partNo,"name":n["name"],"category":n["category"],"oldStock":o["miamiStock"],"newStock":n["miamiStock"]})

has_changes = len(changes) > 0
summary = f"{len(changes)} cambio(s) detectado(s)" if has_changes else "Sin cambios"
print(f"Changes: {summary}")

# 5. Write to Firebase
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

print("Writing to Firebase...")
ok1 = fb_put("salesops_stock_diff",       json.dumps(diff_payload))
ok2 = fb_put("salesops_stock_snapshot",   json.dumps(new_snapshot))
ok3 = fb_post("salesops_stock_changelog", changelog_entry)

# 6. Update local snapshot files
repo_root = os.path.join(os.path.dirname(__file__), '..')
with open(os.path.join(repo_root, 'stock_diff_latest.json'), 'w') as f:
    json.dump(diff_payload, f, ensure_ascii=False, indent=2)
with open(os.path.join(repo_root, 'stock_snapshot_latest.json'), 'w') as f:
    json.dump(new_snapshot, f, ensure_ascii=False, indent=2)

# 7. Send email
if has_changes or FORCE_EMAIL:
    if FORCE_EMAIL and not has_changes:
        summary = f"Envío de prueba — {len(new_snapshot)} productos en lista"
        print("FORCE_EMAIL=true — sending regardless of changes")
    else:
        print("Generating Excel and sending email...")
    excel_bytes = generate_excel(raw_products, changes, date_str)
    send_email(summary, changes, excel_bytes, date_str, len(new_snapshot))
else:
    print("No changes — skipping email")

print(f"\n{'='*40}")
print(f"RESULT: {summary}")
print(f"Firebase: diff={'OK' if ok1 else 'ERR'} snapshot={'OK' if ok2 else 'ERR'} changelog={'OK' if ok3 else 'ERR'}")
if not (ok1 and ok2 and ok3): sys.exit(1)
