"""TẦNG 2 – chạy cả chương trình (có AI) trên test_data/ và chấm điểm theo từng nhãn.

Chạy:
    python tools/cham_diem.py --key KEY                    # chạy mọi câu
    python tools/cham_diem.py --key KEY --nhan ngay_chu    # chỉ câu có nhãn này (nhiều nhãn: a,b)
    python tools/cham_diem.py --key KEY --chi-cau-sai      # chỉ chạy lại câu sai lần trước
    python tools/cham_diem.py --key KEY --khong-cache      # bắt buộc gọi AI thật, không dùng câu trả lời đã lưu
    python tools/cham_diem.py --key KEY --bo-test excel    # chạy bộ thi thử trong folder excel/

Câu trả lời của AI được lưu ở test_data/output/cache_ai.json: prompt không đổi → không gọi AI lại.
"""
import argparse
import csv
import json
import os
import re
import subprocess
import sys
from collections import defaultdict

GOC = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, GOC)
from smartexcel.van_ban import bo_dau  # noqa: E402

p = argparse.ArgumentParser()
p.add_argument("--bo-test", default="test_data", help='folder bộ test (vd "excel" – bộ thi thử)')
p.add_argument("--key", default=os.environ.get("OPENROUTER_API_KEY", ""))
p.add_argument("--model", default="qwen/qwen3.7-flash")
p.add_argument("--nhan", default="")
p.add_argument("--chi-cau-sai", action="store_true")
p.add_argument("--khong-cache", action="store_true")
a = p.parse_args()

TEST = os.path.join(GOC, a.bo_test)
OUT = os.path.join(TEST, "output")
os.makedirs(OUT, exist_ok=True)
FILE_KQ_TRUOC = os.path.join(OUT, "ket_qua_lan_truoc.json")

with open(os.path.join(TEST, "queries.json"), encoding="utf-8") as f:
    tat_ca = json.load(f)
with open(os.path.join(TEST, "dap_an.json"), encoding="utf-8") as f:
    dap_an = json.load(f)
kq_truoc = {}
if os.path.exists(FILE_KQ_TRUOC):
    with open(FILE_KQ_TRUOC, encoding="utf-8") as f:
        kq_truoc = json.load(f)

# ---------- Chọn câu cần chạy ----------
chay = tat_ca
if a.nhan:
    can = set(a.nhan.split(","))
    chay = [q for q in chay if can & set(dap_an[str(q["id"])]["nhan"])]
if a.chi_cau_sai:
    chay = [q for q in chay if not kq_truoc.get(str(q["id"]), False)]
if not chay:
    print("Không có câu nào cần chạy.")
    sys.exit(0)
file_queries = os.path.join(OUT, "queries_chay.json")
with open(file_queries, "w", encoding="utf-8") as f:
    json.dump(chay, f, ensure_ascii=False)

# ---------- Chạy main.py ----------
output = os.path.join(OUT, "submission.csv")
env = dict(os.environ)
if not a.khong_cache:
    env["SMARTEXCEL_CACHE"] = os.path.join(OUT, "cache_ai.json")
lenh = [sys.executable, os.path.join(GOC, "main.py"),
        "--xlsx-dir", os.path.join(TEST, "xlsx_files"), "--queries", file_queries, "--output", output,
        "--openrouter-key", a.key, "--openrouter-model", a.model,
        "--openrouter-base-url", "https://openrouter.ai/api/v1"]
print(f"Chạy {len(chay)}/{len(tat_ca)} câu...")
ma_thoat = subprocess.run(lenh, env=env).returncode
print(f"\nmain.py thoát với mã {ma_thoat}")


# ---------- So sánh ----------
def thanh_so(s):
    s = s.strip().replace(" ", "")
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    s = s.replace(",", ".") if s.count(",") == 1 and "." not in s else s
    try:
        return float(s)
    except ValueError:
        return None


def chuan(s):
    return bo_dau(re.sub(r"\s+", " ", s)).strip(" .")


def dung_khong(ban, dung):
    x, y = thanh_so(ban), thanh_so(dung)
    if x is not None and y is not None:
        return abs(x - y) <= 1e-6 * max(1, abs(y))
    if "," in dung:   # đáp án dạng danh sách → so như tập hợp
        return {chuan(t) for t in ban.split(",")} == {chuan(t) for t in dung.split(",")}
    return chuan(ban) == chuan(dung)


with open(output, encoding="utf-8") as f:
    rows = list(csv.DictReader(f))
ids = [r["id"] for r in rows]
if sorted(ids) != sorted(str(q["id"]) for q in chay) or len(ids) != len(set(ids)):
    print("⚠️ File kết quả thiếu/trùng/thừa id!")

theo_nhan = defaultdict(lambda: [0, 0])
cau_hoi = {str(q["id"]): q["question"] for q in tat_ca}
so_dung = 0
for r in rows:
    da = dap_an[r["id"]]
    ok = dung_khong(r["answer"], da["answer"])
    so_dung += ok
    kq_truoc[r["id"]] = ok
    for n in da["nhan"]:
        theo_nhan[n][0] += ok
        theo_nhan[n][1] += 1
    print(f"{'✅' if ok else '❌'} [{r['id']}] {cau_hoi[r['id']]}\n"
          f"     bạn: {r['answer']!r}   đúng: {da['answer']!r}")

with open(FILE_KQ_TRUOC, "w", encoding="utf-8") as f:
    json.dump(kq_truoc, f, ensure_ascii=False, indent=2)

print("\nTHEO NHÃN:")
for n, (d, t) in sorted(theo_nhan.items(), key=lambda x: (x[1][0] / x[1][1], x[0])):
    print(f"  {'✅' if d == t else '❌'} {n:<22} {d}/{t}")
print(f"\nĐúng {so_dung}/{len(rows)} câu ({100 * so_dung / len(rows):.0f}%)")
tong = sum(kq_truoc.get(str(q["id"]), False) for q in tat_ca)
print(f"Tính cả các lần chạy trước: đúng {tong}/{len(tat_ca)} câu")
