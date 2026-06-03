"""
Cancel active swngfog orders.
讀環境變數 OIDS（逗號分隔），先查 status，只 cancel 還是 active 的。
"""
import os, sys, json, urllib.request, urllib.parse

API_KEY = os.environ["SWNGFOG_API_KEY"]
API_URL = "https://www.swngfog.com/api/v1"

oids_str = os.environ.get("OIDS", "").strip()
if not oids_str:
    print("沒有 OIDs 輸入"); sys.exit(0)
oids = [int(x) for x in oids_str.replace(",", " ").split() if x.strip()]
print(f"輸入 {len(oids)} 個 OID\n")

# 查 status（單筆查避免 batch 行為差異）
ACTIVE = {"in progress", "pending", "processing"}
active, terminated = [], []

# 改用 batch status (orders=...)
r = urllib.request.Request(API_URL, data=urllib.parse.urlencode({
    "key": API_KEY, "action": "status", "orders": ",".join(str(o) for o in oids)
}).encode())
try:
    d = json.loads(urllib.request.urlopen(r, timeout=30).read())
except Exception as e:
    print(f"status 查詢失敗: {e}"); sys.exit(1)

for k, v in d.items():
    s = (v.get("status") or "").lower() if isinstance(v, dict) else ""
    if s in ACTIVE:
        active.append((int(k), v.get("status"), v.get("remains", "?")))
    else:
        terminated.append((int(k), v.get("status", "?")))

print(f"=== 狀態統計 ===")
print(f"active : {len(active)} 筆")
print(f"終止   : {len(terminated)} 筆\n")

print(f"=== Active 訂單 ===")
for o, s, rem in active:
    print(f"  {o} | {s} | remains:{rem}")

print(f"\n=== 已終止訂單（不會 cancel）===")
for o, s in terminated[:10]:
    print(f"  {o} | {s}")
if len(terminated) > 10:
    print(f"  ... 還有 {len(terminated)-10} 筆")

if not active:
    print("\n✅ 沒有 active 訂單需要 cancel，結束")
    sys.exit(0)

# Cancel active
active_oids = [o for o, _, _ in active]
print(f"\n=== 開始 cancel {len(active_oids)} 筆 ===")
r = urllib.request.Request(API_URL, data=urllib.parse.urlencode({
    "key": API_KEY, "action": "cancel", "orders": ",".join(str(o) for o in active_oids)
}).encode())
try:
    result = json.loads(urllib.request.urlopen(r, timeout=60).read())
    print(f"\n=== Cancel 結果 ===")
    if isinstance(result, list):
        for item in result:
            print(f"  {item}")
    elif isinstance(result, dict):
        for k, v in result.items():
            print(f"  {k}: {v}")
    else:
        print(result)
except Exception as e:
    print(f"cancel 失敗: {e}")
    sys.exit(1)
