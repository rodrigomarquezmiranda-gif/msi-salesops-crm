"""
MSI SalesOps — Stock & Price Check
Corre vía GitHub Actions todos los días hábiles a las 18:30 ART.
Lee el pricelist desde Firebase, compara con el snapshot anterior,
y escribe el diff de vuelta a Firebase vía REST API.
"""
import json, re, sys, os
from datetime import datetime, timezone
from urllib import request, error

FB_BASE    = "https://msi-crm-default-rtdb.firebaseio.com"
FB_SECRET  = os.environ.get("FIREBASE_SECRET", "")   # GitHub Secret
FB_PARAMS  = f"?auth={FB_SECRET}" if FB_SECRET else ""

def fb_get(path):
    url = f"{FB_BASE}/{path}.json{FB_PARAMS}"
    try:
        with request.urlopen(url, timeout=30) as r:
            data = json.loads(r.read())
            return data
    except Exception as e:
        print(f"[WARN] GET {path} failed: {e}")
        return None

def fb_put(path, data):
    url = f"{FB_BASE}/{path}.json{FB_PARAMS}"
    payload = json.dumps(data).encode()
    req = request.Request(url, data=payload, method="PUT",
                          headers={"Content-Type": "application/json"})
    try:
        with request.urlopen(req, timeout=30) as r:
            print(f"[OK]   PUT {path} → {r.status}")
            return True
    except Exception as e:
        print(f"[ERR]  PUT {path} failed: {e}")
        return False

def fb_post(path, data):
    url = f"{FB_BASE}/{path}.json{FB_PARAMS}"
    payload = json.dumps(data).encode()
    req = request.Request(url, data=payload, method="POST",
                          headers={"Content-Type": "application/json"})
    try:
        with request.urlopen(req, timeout=30) as r:
            print(f"[OK]   POST {path} → {r.status}")
            return True
    except Exception as e:
        print(f"[ERR]  POST {path} failed: {e}")
        return False

def safe_int(v):
    try:
        m = re.match(r'(\d+)', str(v).strip())
        return int(m.group(1)) if m else 0
    except:
        return 0

def safe_float(v):
    try:
        return float(str(v).replace('$','').replace(',','').strip())
    except:
        return 0.0

# ── 1. Load pricelist from Firebase ──────────────────────────────────────────
print("=== MSI Stock Check ===")
print("Loading pricelist from Firebase...")

pricelist_raw = fb_get("salesops_pricelist")
if not pricelist_raw:
    # Fallback: read from local file in repo
    local_path = os.path.join(os.path.dirname(__file__), '..', 'pricelist_snapshot.json')
    if os.path.exists(local_path):
        with open(local_path) as f:
            pricelist_raw = json.load(f)
        print(f"Fallback: loaded pricelist from local file")
    else:
        print("[ERROR] No pricelist source available. Exiting.")
        sys.exit(1)

# Handle both string (Firebase stores as JSON string) and dict
if isinstance(pricelist_raw, str):
    pricelist_raw = json.loads(pricelist_raw)

raw_products = pricelist_raw.get("products", {})
print(f"Pricelist loaded: {len(raw_products)} products")

# ── 2. Build new snapshot ─────────────────────────────────────────────────────
new_snapshot = {}
for partNo, p in raw_products.items():
    new_snapshot[partNo] = {
        "name":        p.get("name", ""),
        "category":    p.get("category", ""),
        "miamiStock":  safe_int(p.get("miamiStock", 0)),
        "miamiPrice":  safe_float(p.get("price", 0)),
        "miamiEta":    p.get("miamiEta", ""),
        "bondedStock": safe_int(p.get("bondedStock", 0)),
    }

# ── 3. Load old snapshot from Firebase ────────────────────────────────────────
print("Loading previous snapshot from Firebase...")
old_raw = fb_get("salesops_stock_snapshot")
if old_raw and isinstance(old_raw, str):
    old_raw = json.loads(old_raw)
old_snapshot = old_raw if isinstance(old_raw, dict) else {}
print(f"Previous snapshot: {len(old_snapshot)} products")

# ── 4. Compute diff ───────────────────────────────────────────────────────────
changes = []
all_keys = set(new_snapshot) | set(old_snapshot)

for partNo in all_keys:
    new = new_snapshot.get(partNo)
    old = old_snapshot.get(partNo)
    if new and not old:
        changes.append({"type":"new","partNo":partNo,"name":new["name"],
                        "category":new["category"],"newPrice":new["miamiPrice"],
                        "newStock":new["miamiStock"]})
    elif old and not new:
        changes.append({"type":"removed","partNo":partNo,"name":old["name"],
                        "category":old["category"],"oldPrice":old["miamiPrice"]})
    elif new and old:
        price_diff = abs(new["miamiPrice"] - old["miamiPrice"])
        stock_diff = new["miamiStock"] - old["miamiStock"]
        if price_diff >= 0.5:
            changes.append({"type":"price","partNo":partNo,"name":new["name"],
                            "category":new["category"],
                            "oldPrice":old["miamiPrice"],"newPrice":new["miamiPrice"]})
        if abs(stock_diff) >= 1:
            changes.append({"type":"stock","partNo":partNo,"name":new["name"],
                            "category":new["category"],
                            "oldStock":old["miamiStock"],"newStock":new["miamiStock"]})

has_changes = len(changes) > 0
now_iso = datetime.now(timezone.utc).isoformat()
print(f"Changes found: {len(changes)}")

# Keep last changes if no new ones
last_changes = []
last_changes_at = now_iso
if old_raw and isinstance(old_raw, dict):
    pass  # old_snapshot is products, not diff payload

diff_payload = {
    "checkedAt":     now_iso,
    "productCount":  len(new_snapshot),
    "hasChanges":    has_changes,
    "changeCount":   len(changes),
    "changes":       changes[:100] if has_changes else [],
    "summary":       f"{len(changes)} cambio(s) detectado(s)" if has_changes else "Sin cambios",
    "triggeredBy":   "github-actions",
    "triggerMode":   "auto",
}

changelog_entry = {
    "checkedAt":    now_iso,
    "productCount": len(new_snapshot),
    "hasChanges":   has_changes,
    "changeCount":  len(changes),
    "summary":      diff_payload["summary"],
    "triggerMode":  "auto",
    "triggeredBy":  "github-actions",
}

# ── 5. Write to Firebase ──────────────────────────────────────────────────────
print("Writing results to Firebase...")
ok1 = fb_put("salesops_stock_diff",     json.dumps(diff_payload))
ok2 = fb_put("salesops_stock_snapshot", json.dumps(new_snapshot))
ok3 = fb_post("salesops_stock_changelog", changelog_entry)

# ── 6. Update local files for repo commit ─────────────────────────────────────
repo_root = os.path.join(os.path.dirname(__file__), '..')
with open(os.path.join(repo_root, 'stock_diff_latest.json'), 'w') as f:
    json.dump(diff_payload, f, ensure_ascii=False, indent=2)
with open(os.path.join(repo_root, 'stock_snapshot_latest.json'), 'w') as f:
    json.dump(new_snapshot, f, ensure_ascii=False, indent=2)

# Summary
print(f"\n{'='*40}")
print(f"SUMMARY: {diff_payload['summary']}")
if changes:
    for c in changes[:20]:
        if   c['type']=='price':   print(f"  [PRECIO]  {c['partNo']} {c['name']}: ${c['oldPrice']} → ${c['newPrice']}")
        elif c['type']=='stock':   print(f"  [STOCK]   {c['partNo']} {c['name']}: {c['oldStock']} → {c['newStock']}")
        elif c['type']=='new':     print(f"  [NUEVO]   {c['partNo']} {c['name']}: ${c['newPrice']}")
        elif c['type']=='removed': print(f"  [BAJA]    {c['partNo']} {c['name']}")
print(f"Firebase: diff={'OK' if ok1 else 'ERR'} snapshot={'OK' if ok2 else 'ERR'} changelog={'OK' if ok3 else 'ERR'}")

if not (ok1 and ok2 and ok3):
    sys.exit(1)
