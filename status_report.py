"""
Query swngfog status for a list of OIDs, output JSON per-OID.
Read-only, no cancel.
"""
import os, sys, json, time, urllib.request, urllib.parse

API_KEY = os.environ["SWNGFOG_API_KEY"]
API_URL = "https://www.swngfog.com/api/v1"

oids_str = os.environ.get("OIDS", "").strip()
if not oids_str:
    print("no OIDs"); sys.exit(0)
oids = [int(x) for x in oids_str.replace(",", " ").split() if x.strip()]
print(f"查 {len(oids)} 個 OID status\n")

results = {}
for i, oid in enumerate(oids):
    try:
        r = urllib.request.Request(API_URL, data=urllib.parse.urlencode({
            "key": API_KEY, "action": "status", "order": str(oid)
        }).encode())
        v = json.loads(urllib.request.urlopen(r, timeout=20).read())
        results[oid] = v
    except Exception as e:
        results[oid] = {"error": str(e)}
    if (i+1) % 50 == 0:
        print(f"  進度 {i+1}/{len(oids)}", flush=True)
    time.sleep(0.1)

print("\n=== STATUS_JSON_START ===")
print(json.dumps(results))
print("=== STATUS_JSON_END ===")
