"""
對每個 OID：先試 status 取 remains（可能 403），然後直接 cancel。
每筆結果都記錄，最後輸出完整 JSON manifest。
"""
import os, sys, json, time, urllib.request, urllib.parse

API_KEY = os.environ["SWNGFOG_API_KEY"]
API_URL = "https://www.swngfog.com/api/v1"

oids_str = os.environ.get("OIDS", "").strip()
if not oids_str:
    print("no OIDs"); sys.exit(0)
oids = [int(x) for x in oids_str.replace(",", " ").split() if x.strip()]
print(f"處理 {len(oids)} 個 OID（每筆 status + cancel）\n", flush=True)

results = {}
for i, oid in enumerate(oids):
    entry = {}
    # status
    try:
        r = urllib.request.Request(API_URL, data=urllib.parse.urlencode({
            "key": API_KEY, "action": "status", "order": str(oid)
        }).encode())
        v = json.loads(urllib.request.urlopen(r, timeout=15).read())
        if isinstance(v, dict):
            entry["status"] = v.get("status")
            entry["remains"] = v.get("remains")
            entry["quantity"] = v.get("quantity")
    except Exception as e:
        entry["status_err"] = str(e)[:80]
    # cancel
    try:
        r = urllib.request.Request(API_URL, data=urllib.parse.urlencode({
            "key": API_KEY, "action": "cancel", "order": str(oid)
        }).encode())
        v = json.loads(urllib.request.urlopen(r, timeout=15).read())
        entry["cancel"] = v
    except Exception as e:
        entry["cancel_err"] = str(e)[:80]
    results[oid] = entry
    if (i+1) % 100 == 0:
        print(f"  進度 {i+1}/{len(oids)}", flush=True)
    time.sleep(0.15)

print("\n=== RESULT_JSON_START ===")
print(json.dumps(results))
print("=== RESULT_JSON_END ===")
