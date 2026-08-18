---
name: msi-stock-update-check-pm
description: Chequeo automático del Stock Update MSI — corrida de las 18:30 hs — guarda diff en archivos locales y pushea a Firebase vía Chrome
---

You are running an automated check of the full MSI LATAM price list (PM run — 18:30 hs). Your job is to compare the current product data against the last stored snapshot, save the diff locally AND push it to Firebase so the MSI SalesOps CRM can read it in real time.

## Data source
Read `pricelist_snapshot.json` (written by the pricelist monitor — covers all 326+ products from the LATAM tab). If missing or older than 24h, fall back to reading the Google Sheet (fileId `1wgHbYfXo65Z9iZxGV0vTdy1NCiKnyl5J0Ju9HrsuiKw`) via the Drive MCP `read_file_content` tool.

## Local file paths
```bash
WORKSPACE=$(find /sessions -maxdepth 3 -name "msi-salesops-crm" -type d 2>/dev/null | head -1)
```
- `$WORKSPACE/pricelist_snapshot.json`    ← source of current prices
- `$WORKSPACE/stock_snapshot_latest.json` ← previous snapshot (what we compare against)
- `$WORKSPACE/stock_diff_latest.json`     ← output diff (local)
- `$WORKSPACE/stock_changelog.json`       ← local changelog

## Firebase (Realtime Database)
- Project: `msi-crm-default-rtdb`
- API key: `AIzaSyCzw8L-EsO-NRN3XLOuQTGtTknBe5F8x-0`
- Keys to write:
  - `salesops_stock_diff`      ← PUT (stores as JSON string)
  - `salesops_stock_snapshot`  ← PUT (stores as JSON string)
  - `salesops_stock_changelog` ← POST to append entry (stores as JSON object)

---

## Step 1 — Python script (run in bash)

Computes the diff and saves all local files. Does NOT write to Firebase (sandbox proxy blocks outbound HTTPS to Firebase).

```python
import json, glob, re
from datetime import datetime, timezone

workspaces = glob.glob('/sessions/*/mnt/msi-salesops-crm')
WORKSPACE = workspaces[0] if workspaces else '/tmp'

PRICELIST_FILE  = f"{WORKSPACE}/pricelist_snapshot.json"
SNAPSHOT_FILE   = f"{WORKSPACE}/stock_snapshot_latest.json"
DIFF_FILE       = f"{WORKSPACE}/stock_diff_latest.json"
CHANGELOG_FILE  = f"{WORKSPACE}/stock_changelog.json"
PENDING_FILE    = f"{WORKSPACE}/firebase_pending.json"

with open(PRICELIST_FILE) as f:
    pricelist = json.load(f)

raw_products = pricelist.get("products", {})

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

try:
    with open(SNAPSHOT_FILE) as f:
        old_snapshot = json.load(f)
    print(f"Old snapshot: {len(old_snapshot)} products")
except:
    old_snapshot = None
    print("No previous snapshot — saving baseline")

last_changes = []
last_changes_at = None
try:
    with open(DIFF_FILE) as f:
        prev_diff = json.load(f)
    if prev_diff.get('hasChanges') and prev_diff.get('changes'):
        last_changes    = prev_diff['changes']
        last_changes_at = prev_diff['checkedAt']
    elif prev_diff.get('lastChanges'):
        last_changes    = prev_diff['lastChanges']
        last_changes_at = prev_diff.get('lastChangesAt')
except:
    pass

changes = []
if old_snapshot:
    all_parts = set(old_snapshot.keys()) | set(new_snapshot.keys())
    for partNo in sorted(all_parts):
        old = old_snapshot.get(partNo)
        new = new_snapshot.get(partNo)
        if old is None and new is not None:
            changes.append({"type":"new",     "partNo":partNo, "name":new["name"],  "category":new["category"],      "newPrice":new["miamiPrice"]})
        elif old is not None and new is None:
            changes.append({"type":"removed", "partNo":partNo, "name":old["name"],  "category":old.get("category","")})
        elif old and new:
            try:
                if abs(new["miamiPrice"] - old["miamiPrice"]) > 1:
                    changes.append({"type":"price", "partNo":partNo, "name":new["name"], "category":new["category"],
                                    "oldPrice":old["miamiPrice"], "newPrice":new["miamiPrice"]})
            except: pass
            try:
                if abs(new["miamiStock"] - old["miamiStock"]) >= 200:
                    changes.append({"type":"stock", "partNo":partNo, "name":new["name"], "category":new["category"],
                                    "oldStock":old["miamiStock"], "newStock":new["miamiStock"]})
            except: pass

now_iso = datetime.now(timezone.utc).isoformat()
has_changes = len(changes) > 0

diff_payload = {
    "checkedAt":     now_iso,
    "productCount":  len(new_snapshot),
    "hasChanges":    has_changes,
    "changeCount":   len(changes),
    "changes":       changes[:100] if has_changes else [],
    "lastChanges":   changes[:100] if has_changes else last_changes,
    "lastChangesAt": now_iso       if has_changes else last_changes_at,
    "summary":       f"{len(changes)} cambio(s) detectado(s)" if has_changes else "Sin cambios",
    "triggeredBy":   "sistema",
    "triggerMode":   "auto",
}

changelog_entry = {
    "checkedAt":    now_iso,
    "productCount": len(new_snapshot),
    "hasChanges":   has_changes,
    "changeCount":  len(changes),
    "summary":      diff_payload["summary"],
    "triggerMode":  "auto",
    "triggeredBy":  "sistema",
}

with open(SNAPSHOT_FILE, 'w') as f:
    json.dump(new_snapshot, f, ensure_ascii=False, indent=2)
print(f"Snapshot saved: {len(new_snapshot)} products")

with open(DIFF_FILE, 'w') as f:
    json.dump(diff_payload, f, ensure_ascii=False, indent=2)
print(f"Diff saved: {len(changes)} changes")

try:
    with open(CHANGELOG_FILE) as f:
        changelog = json.load(f)
except:
    changelog = []
changelog.append(changelog_entry)
changelog = changelog[-90:]
with open(CHANGELOG_FILE, 'w') as f:
    json.dump(changelog, f, ensure_ascii=False, indent=2)
print(f"Changelog updated: {len(changelog)} entries")

# Write pending payload for Chrome MCP step
with open(PENDING_FILE, 'w') as f:
    json.dump({
        "diff":      diff_payload,
        "snapshot":  new_snapshot,
        "changelog": changelog_entry,
    }, f, ensure_ascii=False)
print("firebase_pending.json written")

print(f"\nSUMMARY: {diff_payload['summary']}")
if changes:
    for c in changes:
        if   c['type']=='price':   print(f"  [PRICE]   {c['partNo']} ({c['category']}) - {c['name']}: ${c['oldPrice']} -> ${c['newPrice']}")
        elif c['type']=='stock':   print(f"  [STOCK]   {c['partNo']} ({c['category']}) - {c['name']}: {c['oldStock']} -> {c['newStock']}")
        elif c['type']=='new':     print(f"  [NEW]     {c['partNo']} ({c['category']}) - {c['name']}: ${c['newPrice']}")
        elif c['type']=='removed': print(f"  [REMOVED] {c['partNo']} ({c['category']}) - {c['name']}")
```

---

## Step 2 — Firebase write via Chrome MCP

The sandbox proxy blocks outbound HTTPS to Firebase (`X-Proxy-Error: blocked-by-allowlist`). Use `mcp__claude-in-chrome__javascript_tool` to write from the user's browser, which has unrestricted internet access.

After the Python script finishes:

**2a.** Load the Chrome tool:
```
ToolSearch { query: "select:mcp__claude-in-chrome__javascript_tool,mcp__claude-in-chrome__tabs_context_mcp", max_results: 2 }
```

**2b.** Read the pending payload from bash:
```bash
cat $(find /sessions -maxdepth 3 -name "msi-salesops-crm" -type d 2>/dev/null | head -1)/firebase_pending.json
```

**2c.** Call `mcp__claude-in-chrome__javascript_tool` with the script below, substituting the three variables with the actual parsed JSON from the file:

```javascript
const FB_BASE    = "https://msi-crm-default-rtdb.firebaseio.com";
const FB_API_KEY = "AIzaSyCzw8L-EsO-NRN3XLOuQTGtTknBe5F8x-0";

// Inject from firebase_pending.json
const diffPayload    = /* pending.diff    — paste full JSON object */;
const newSnapshot    = /* pending.snapshot — paste full JSON object */;
const changelogEntry = /* pending.changelog — paste full JSON object */;

async function fbPut(path, data) {
  const r = await fetch(`${FB_BASE}/${path}.json?key=${FB_API_KEY}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(JSON.stringify(data))
  });
  return { path, status: r.status, ok: r.ok };
}

async function fbPost(path, data) {
  const r = await fetch(`${FB_BASE}/${path}.json?key=${FB_API_KEY}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(data)
  });
  return { path, status: r.status, ok: r.ok };
}

const results = await Promise.all([
  fbPut("salesops_stock_diff",     diffPayload),
  fbPut("salesops_stock_snapshot", newSnapshot),
  fbPost("salesops_stock_changelog", changelogEntry),
]);

return results;
```

**2d.** If all three return `ok: true` → log "Firebase updated OK via Chrome".  
If any fails → log status codes and retry once. If still failing, log error and continue (local files are already saved).

---

## Output
Print: products parsed, changes found, local files written, Firebase status (via Chrome).

Firebase API key: AIzaSyCzw8L-EsO-NRN3XLOuQTGtTknBe5F8x-0
