
#!/usr/bin/env python3
import csv, io, json, os, urllib.request

url = os.environ.get("SCRIPT_URL", "").strip()
token = os.environ.get("SITE_TOKEN", "").strip()
if not url:
    raise SystemExit("Missing SHEET_SCRIPT_URL secret")

sep = "&" if "?" in url else "?"
full = url + sep + "token=" + token + "&t=1"
req = urllib.request.Request(full, headers={"User-Agent": "connectly-action"})
with urllib.request.urlopen(req, timeout=90) as r:
    text = r.read().decode("utf-8", "replace")

if text.strip().startswith("{") or text.strip().startswith("<"):
    raise SystemExit("Script did not return CSV")

rows = list(csv.DictReader(io.StringIO(text)))

def pick(row, names):
    keys = list(row.keys())
    for n in names:
        for k in keys:
            if k and k.strip().lower() == n:
                v = str(row[k] or "").strip()
                if v:
                    return v
    for n in names:
        for k in keys:
            if k and n in k.lower():
                v = str(row[k] or "").strip()
                if v:
                    return v
    return ""

out = []
for row in rows:
    name = pick(row, ["name"])
    if not name:
        continue
    out.append({
        "name": name,
        "city": pick(row, ["city", "location"]),
        "niche": pick(row, ["skill", "niche", "category"]),
        "followers": pick(row, ["followers", "follower"]),
        "instagram": pick(row, ["instagram", "profile", "portfolio"]),
        "charges": pick(row, ["charges", "charge", "rate"]),
        "gender": pick(row, ["gender"]),
    })

with open("creators.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False)
print("wrote", len(out), "creators")
