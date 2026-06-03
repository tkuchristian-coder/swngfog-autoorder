"""
Cancel active swngfog orders.
讀環境變數 OIDS（逗號分隔），先查 status，只 cancel 還是 active 的。
"""
import os, sys, json, time, urllib.request, urllib.parse

API_KEY = os.environ["SWNGFOG_API_KEY"]
API_URL = "https://www.swngfog.com/api/v1"

oids_str = os.environ.get("OIDS", "").strip()
if not oids_str:
    print("沒有 OIDs 輸入"); sys.exit(0)
oids = [int(x) for x in oids_str.replace(",", " ").split() if x.strip()]
print(f"輸入 {len(oids)} 個 OID\n")

ACTIVE = {"in progress", "pending", "processing"}
active, terminated = [], []

def query_status_single(oid):
    """單筆 status 查詢，避開 batch 403。"""
    r = urllib.request.Request(API_URL, data=urllib.parse.urlencode({
        "key": API_KEY, "action": "status", "order": str(oid)
    }).encode())
    return json.loads(urllib.request.urlopen(r, timeout=20).read())

print(f"逐筆查 status ({len(oids)} 筆)…")
for i, oid in enumerate(oids):
    try:
        v = query_status_single(oid)
        s = (v.get("status") or "").lower() if isinstance(v, dict) else ""
        if s in ACTIVE:
            active.append((oid, v.get("status"), v.get("remains", "?")))
        else:
            terminated.append((oid, v.get("status", "?")))
    except Exception as e:
        print(f"  ⚠️ {oid} status 失敗: {e}")
        terminated.append((oid, f"ERROR:{e}"))
    if (i+1) % 10 == 0:
        print(f"  進度 {i+1}/{len(oids)}")
    time.sleep(0.2)

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

# Cancel active（先試 batch，失敗則改逐筆）
active_oids = [o for o, _, _ in active]
print(f"\n=== 開始 cancel {len(active_oids)} 筆 ===")
try:
    r = urllib.request.Request(API_URL, data=urllib.parse.urlencode({
        "key": API_KEY, "action": "cancel", "orders": ",".join(str(o) for o in active_oids)
    }).encode())
    result = json.loads(urllib.request.urlopen(r, timeout=60).read())
    print(f"batch cancel 結果：")
    print(json.dumps(result, indent=2, ensure_ascii=False))
except Exception as e:
    print(f"batch cancel 失敗 ({e})，改逐筆…")
    ok = fail = 0
    for oid in active_oids:
        try:
            r = urllib.request.Request(API_URL, data=urllib.parse.urlencode({
                "key": API_KEY, "action": "cancel", "order": str(oid)
            }).encode())
            res = json.loads(urllib.request.urlopen(r, timeout=20).read())
            print(f"  {oid}: {res}")
            ok += 1
        except Exception as ee:
            print(f"  {oid}: ❌ {ee}")
            fail += 1
        time.sleep(0.2)
    print(f"\n逐筆 cancel: ✅{ok} ❌{fail}")
