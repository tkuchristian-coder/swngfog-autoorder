"""
Cancel by row / IGID / URL — 一站式 cancel 工具。

輸入 (TARGET 環境變數)：
  - 純數字 → 視為列號（多列用逗號）：例 "2902" 或 "2902,1681"
  - 其他字串 → 在 sheet 搜尋 C 欄含此字串的列：例 "DZGsIvQyBNo" 或 "b_moooooo"

流程：
  1. 找 sheet 對應列號
  2. 掃 GitHub Action log 找該列所有 fog訂單ID
  3. 逐筆查 swngfog status
  4. cancel 還是 active 的訂單
"""
import os, sys, json, time, re, io, zipfile, urllib.request, urllib.parse

API_KEY = os.environ["SWNGFOG_API_KEY"]
API_URL = "https://www.swngfog.com/api/v1"
GH_TOKEN = os.environ["GH_TOKEN"]
REPO = os.environ.get("REPO", "tkuchristian-coder/swngfog-autoorder")
TARGET = os.environ["TARGET"].strip()

# Google Sheet
import gspread
from google.oauth2.service_account import Credentials
SHEET_ID = os.environ.get("SHEET_ID", "1sbYoxrrMOPZsA2q6dPDqzU3OvoWykcsy-4ao9f9oyJk")
SHEET_TAB = os.environ.get("SHEET_TAB_NAME", "2026年3月")
creds = Credentials.from_service_account_file("credentials.json",
    scopes=["https://www.googleapis.com/auth/spreadsheets"])
ws = gspread.authorize(creds).open_by_key(SHEET_ID).worksheet(SHEET_TAB)
all_rows = ws.get_all_values()

# ── Step 1: 解析 TARGET 找列號 ──
target_rows = []
if all(p.strip().isdigit() for p in TARGET.split(",")):
    target_rows = [int(p.strip()) for p in TARGET.split(",")]
    print(f"輸入為列號: {target_rows}")
else:
    print(f"搜尋字串「{TARGET}」…")
    for idx, row in enumerate(all_rows, start=1):
        for cell in row[:9]:
            if TARGET in cell:
                target_rows.append(idx)
                print(f"  列 {idx}: {row[:9]}")
                break

if not target_rows:
    print("❌ 找不到匹配的列")
    sys.exit(1)

# ── Step 2: 掃 log 找 OIDs ──
print(f"\n=== 掃 GitHub Action log 找 OIDs（列：{target_rows}）===")
row_to_oids = {r: [] for r in target_rows}
seen = set()

def gh_get(url):
    req = urllib.request.Request(url, headers={"Authorization": f"token {GH_TOKEN}"})
    return urllib.request.urlopen(req, timeout=30).read()

# 抓最近 N 頁
all_done = False
for page in range(1, 5):
    if all_done: break
    data = json.loads(gh_get(f"https://api.github.com/repos/{REPO}/actions/runs?per_page=100&page={page}"))
    runs = data.get("workflow_runs", [])
    if not runs: break
    for r in runs:
        if r["conclusion"] != "success": continue
        if r["name"] != "swngfog Auto Order": continue  # 跳過 cancel workflow 自己
        try:
            log_bytes = gh_get(f"https://api.github.com/repos/{REPO}/actions/runs/{r['id']}/logs")
            zf = zipfile.ZipFile(io.BytesIO(log_bytes))
            for name in zf.namelist():
                if "run-orders" in name and "system" not in name:
                    content = zf.read(name).decode("utf-8", errors="ignore")
                    for row_num in target_rows:
                        marker = f"[處理] 列{row_num}"
                        idx_start = content.find(marker)
                        while idx_start >= 0:
                            idx_end = content.find("[處理] 列", idx_start + len(marker))
                            if idx_end < 0: idx_end = len(content)
                            block = content[idx_start:idx_end]
                            for m in re.finditer(r"fog訂單ID:(\d+)", block):
                                oid = int(m.group(1))
                                if oid not in seen:
                                    seen.add(oid)
                                    row_to_oids[row_num].append(oid)
                            idx_start = content.find(marker, idx_end)
                    break
        except Exception as e:
            print(f"  ⚠️ run {r['id']} log 失敗: {e}")

total_oids = sum(len(v) for v in row_to_oids.values())
print(f"\n找到 {total_oids} 個 OID:")
for row_num, oids in row_to_oids.items():
    print(f"  列 {row_num}: {len(oids)} 個")

if total_oids == 0:
    print("❌ 沒找到任何 OID")
    sys.exit(0)

all_oids = sorted(seen)

# ── Step 3: 逐筆查 status ──
print(f"\n=== 逐筆查 status ({len(all_oids)} 筆) ===")
ACTIVE = {"in progress", "pending", "processing"}
active, terminated = [], []
for i, oid in enumerate(all_oids):
    try:
        r = urllib.request.Request(API_URL, data=urllib.parse.urlencode({
            "key": API_KEY, "action": "status", "order": str(oid)
        }).encode())
        v = json.loads(urllib.request.urlopen(r, timeout=20).read())
        s = (v.get("status") or "").lower() if isinstance(v, dict) else ""
        if s in ACTIVE:
            active.append((oid, v.get("status"), v.get("remains", "?")))
        else:
            terminated.append((oid, v.get("status", "?")))
    except Exception as e:
        terminated.append((oid, f"ERROR:{e}"))
    if (i + 1) % 20 == 0:
        print(f"  進度 {i+1}/{len(all_oids)}")
    time.sleep(0.1)

print(f"\n=== Active: {len(active)} 筆 / 已終止: {len(terminated)} 筆 ===")
for o, s, rem in active[:20]:
    print(f"  active: {o} | {s} | remains:{rem}")
if len(active) > 20:
    print(f"  ...還有 {len(active)-20} 筆 active")

if not active:
    print("\n✅ 沒有 active 訂單需 cancel")
    sys.exit(0)

# ── Step 4: cancel ──
print(f"\n=== Cancel {len(active)} 筆 active ===")
active_oids = [o for o, _, _ in active]
try:
    r = urllib.request.Request(API_URL, data=urllib.parse.urlencode({
        "key": API_KEY, "action": "cancel", "orders": ",".join(str(o) for o in active_oids)
    }).encode())
    result = json.loads(urllib.request.urlopen(r, timeout=60).read())
    print("batch cancel 結果:")
    print(json.dumps(result, indent=2, ensure_ascii=False)[:3000])
except Exception as e:
    print(f"batch cancel 失敗 ({e})，逐筆…")
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
        time.sleep(0.15)
    print(f"\n✅ {ok}  ❌ {fail}")
