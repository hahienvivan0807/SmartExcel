"""TẦNG 1 – kiểm tra Bước 1 (làm sạch), Bước 2 (tóm tắt) và phần chọn bảng của Bước 3.
Không gọi AI, miễn phí, chạy vài giây. Kết quả luôn giống nhau giữa các lần chạy.

Chạy: python tools/test_tang1.py [-v] [--bo-test excel]
      -v          in cả những kiểm tra đạt
      --bo-test   folder bộ test (mặc định test_data; folder "excel" là bộ thi thử)
"""
import argparse
import json
import os
import sys
from collections import defaultdict

import pandas as pd

GOC = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, GOC)

from smartexcel.doc_bang import doc_tat_ca  # noqa: E402
from smartexcel.hoi_dap import chon_bang  # noqa: E402
from smartexcel.tom_tat import tom_tat_tat_ca  # noqa: E402

_p = argparse.ArgumentParser()
_p.add_argument("-v", action="store_true")
_p.add_argument("--bo-test", default="test_data")
_a = _p.parse_args()
TEST = os.path.join(GOC, _a.bo_test)
CHI_TIET = _a.v


def kieu_cot(cot):
    if pd.api.types.is_bool_dtype(cot):
        return "chu"
    if pd.api.types.is_numeric_dtype(cot):
        return "so"
    if pd.api.types.is_datetime64_any_dtype(cot):
        return "ngay"
    return "chu"


def giong(thuc_te, mong_doi):
    if thuc_te is None or (not isinstance(thuc_te, str) and pd.isna(thuc_te)):
        return False
    if isinstance(mong_doi, (int, float)):
        try:
            return abs(float(thuc_te) - mong_doi) < 1e-9
        except (TypeError, ValueError):
            return False
    if isinstance(thuc_te, (pd.Timestamp,)) or hasattr(thuc_te, "date"):
        return str(pd.Timestamp(thuc_te).date()) == mong_doi
    return str(thuc_te) == mong_doi


with open(os.path.join(TEST, "tang1.json"), encoding="utf-8") as f:
    KY_VONG = json.load(f)
with open(os.path.join(TEST, "dap_an.json"), encoding="utf-8") as f:
    DAP_AN = json.load(f)
with open(os.path.join(TEST, "queries.json"), encoding="utf-8") as f:
    CAU_HOI = {str(q["id"]): q["question"] for q in json.load(f)}

cac_bang = doc_tat_ca(os.path.join(TEST, "xlsx_files"))
tom_tat_tat_ca(cac_bang)
theo_ten = {b.ten: b for b in cac_bang}

ket_qua = []                       # (đạt?, mô tả, [nhãn])
theo_nhan = defaultdict(lambda: [0, 0])


def kiem_tra(dat, mo_ta, nhan):
    ket_qua.append((dat, mo_ta))
    for n in nhan:
        theo_nhan[n][0] += dat
        theo_nhan[n][1] += 1
    if CHI_TIET or not dat:
        print(f"  {'✅' if dat else '❌'} {mo_ta}")


# ---------------- Bước 1–2: từng bảng ----------------
print("=" * 70, "\nBƯỚC 1–2: LÀM SẠCH BẢNG")
for ten, kv in KY_VONG.items():
    nhan = kv.get("nhan", [])
    print(f"[{ten}]")
    b = theo_ten.get(ten)
    kiem_tra(b is not None, f"đọc được bảng {ten}", nhan)
    if b is None:
        continue
    df = b.df
    kiem_tra(b.huong in kv["huong"], f"hướng = {b.huong} (mong đợi {'/'.join(kv['huong'])})", nhan)
    kiem_tra(len(df) == kv["so_dong"], f"số dòng = {len(df)} (mong đợi {kv['so_dong']})", nhan)
    if "cot" in kv:
        kiem_tra(list(df.columns) == kv["cot"], f"tên cột = {list(df.columns)}", nhan)
    for cot, kieu in kv.get("kieu", {}).items():
        if cot not in df.columns:
            kiem_tra(False, f'có cột "{cot}" (các cột hiện có: {list(df.columns)})', nhan)
            continue
        kiem_tra(kieu_cot(df[cot]) == kieu, f'cột "{cot}" kiểu {kieu_cot(df[cot])} (mong đợi {kieu})', nhan)
    for dong, cot, gia_tri in kv.get("o", []):
        tt = df[cot].iloc[dong] if cot in df.columns and dong < len(df) else None
        kiem_tra(giong(tt, gia_tri), f'ô [{dong}, "{cot}"] = {tt!r} (mong đợi {gia_tri!r})', nhan)
    kiem_tra(0 < len(b.tom_tat) <= 4000, f"tóm tắt dài {len(b.tom_tat)} ký tự (≤ 4000)", nhan)

# ---------------- Bước 3: chọn bảng ----------------
print("=" * 70, "\nBƯỚC 3: CHỌN BẢNG CHO TỪNG CÂU HỎI")
for qid, da in DAP_AN.items():
    chon = [b.ten for b in chon_bang(CAU_HOI[qid], cac_bang)]
    thieu = [t for t in da["bang"] if t not in chon]
    mo_ta = f"[câu {qid}] {CAU_HOI[qid]}  → chọn {chon}"
    if thieu:
        mo_ta += f"  THIẾU {thieu}"
    kiem_tra(not thieu, mo_ta, da["nhan"] + ["chon_bang"])

# ---------------- Dự phòng: chọn bảng rộng + chọn dòng liên quan ----------------
print("=" * 70, "\nDỰ PHÒNG: CHỌN BẢNG RỘNG + CHỌN DÒNG")
from smartexcel.du_phong import chon_bang_rong, chon_dong_lien_quan, KY_TU_TOI_DA_DP1  # noqa: E402
from smartexcel.hoi_dap import tim_gia_tri  # noqa: E402

for qid, da in DAP_AN.items():
    rong = chon_bang_rong(CAU_HOI[qid], cac_bang)
    ten_rong = [b.ten for b in rong]
    tong = sum(len(b.tom_tat) for b in rong)
    kiem_tra(all(t in ten_rong for t in da["bang"]) and (tong <= KY_TU_TOI_DA_DP1 or len(rong) == 1),
             f"[câu {qid}] dự phòng 1 gửi {len(rong)} bảng ({tong} ký tự), có đủ bảng cần", ["dp1_chon_bang"])

bang_thu_chi = theo_ten.get("thu_chi.xlsx::GiaoDich")
for cau, dieu_kien, nhan in [] if bang_thu_chi is None else [   # chỉ có trong test_data
    ("Danh mục Ăn uống có bao nhiêu giao dịch?", {"Danh mục": "Ăn uống"}, "dp2_1_gia_tri"),
    ("Có bao nhiêu giao dịch Chi thuộc danh mục Hóa đơn?", {"Loại": "Chi", "Danh mục": "Hóa đơn"}, "dp2_2_gia_tri"),
]:
    dong, ghi_chu = chon_dong_lien_quan(bang_thu_chi, tim_gia_tri(cau, [bang_thu_chi]))
    dat = len(dong) == 40 and all((dong[c] == v).all() for c, v in dieu_kien.items())
    kiem_tra(dat, f"dự phòng 2 lọc '{cau}' → {len(dong)} dòng, đều khớp {dieu_kien} ({ghi_chu})", [nhan])
if bang_thu_chi is not None and "van_don.xlsx::VanDon" in theo_ten:
    dong, ghi_chu = chon_dong_lien_quan(bang_thu_chi, [])
    kiem_tra(len(dong) == 40 and "dòng đầu" in ghi_chu, f"dự phòng 2 không có gì để lọc → {ghi_chu}", ["dp2_khong_loc"])
    dong, ghi_chu = chon_dong_lien_quan(theo_ten["van_don.xlsx::VanDon"], [])
    kiem_tra(len(dong) == 12 and "toàn bộ" in ghi_chu, f"dự phòng 2 bảng nhỏ → {ghi_chu}", ["dp2_bang_nho"])

# ---------------- Bước 5: chuẩn hóa đáp án ----------------
print("=" * 70, "\nBƯỚC 5: CHUẨN HÓA ĐÁP ÁN")
import numpy as np  # noqa: E402
from datetime import date, datetime  # noqa: E402
from smartexcel.chuan_hoa import chuan_hoa  # noqa: E402

CA_CHUAN_HOA = [
    # (đầu vào, mong đợi, nhãn)
    (3, "3", "so_nguyen"), (np.int64(950000), "950000", "so_nguyen"),
    (9.0, "9", "so_thap_phan"), (np.float64(30.5), "30.5", "so_thap_phan"),
    (38142.857142857145, "38142.86", "lam_tron"), (0.1 + 0.2, "0.3", "lam_tron"),
    (float("nan"), "", "rong"), (None, "", "rong"), (pd.NA, "", "rong"),
    (True, "Có", "dung_sai"), (np.bool_(False), "Không", "dung_sai"),
    (pd.Timestamp("2026-10-01"), "01/10/2026", "ngay"), (date(1968, 3, 15), "15/03/1968", "ngay"),
    (datetime(2026, 10, 1, 8, 30), "01/10/2026 08:30", "ngay"),
    (pd.Series(["Phạm Thu Hà"]), "Phạm Thu Hà", "series_1"),
    (pd.Series([9.0], index=[2]), "9", "series_1"),
    (pd.DataFrame({"a": [590]}), "590", "series_1"),
    (pd.Series(["VD0007", "VD0012"]), "VD0007, VD0012", "danh_sach"),
    (["B", "A"], "B, A", "danh_sach"), (np.array([1, 2]), "1, 2", "danh_sach"),
    (pd.Series([], dtype=float), "", "rong"),
    ("  Quần jean \n", "Quần jean", "chu"), ("Đáp án: 11.", "11", "chu"), ('"TP.HCM"', "TP.HCM", "chu"),
    ("Kết quả: Bùi Văn Tú", "Bùi Văn Tú", "chu"), ("9.0", "9", "chu_so"), ("1200000.00", "1200000", "chu_so"),
    ("8.333333", "8.33", "chu_so"), ("0312345678", "0312345678", "ma_so_0"), ("070000", "070000", "ma_so_0"),
    ("10A", "10A", "chu"), ("Đang giao...", "Đang giao...", "chu"),
    ("Đáp án: 800.000 đồng.", "800000", "chu_don_vi"), ("12,5 kg", "12.5", "chu_don_vi"),
    ("80%", "80", "chu_don_vi"), ("1,250", "1250", "chu_don_vi"), ("1.25", "1.25", "chu_so"),
    ("2015", "2015", "chu_so"), ("VD0007, VD0012", "VD0007, VD0012", "chu"),
]
for dau_vao, mong_doi, nhan in CA_CHUAN_HOA:
    try:
        tt = chuan_hoa(dau_vao)
    except Exception as e:
        tt = f"LỖI {e!r}"
    kiem_tra(tt == mong_doi, f"chuẩn hóa {type(dau_vao).__name__} {dau_vao!r:.40} → {tt!r} (mong đợi {mong_doi!r})",
             ["b5_" + nhan])

# ---------------- Tổng kết ----------------
dat = sum(k[0] for k in ket_qua)
print("=" * 70)
print("THEO NHÃN:")
for n, (d, t) in sorted(theo_nhan.items(), key=lambda x: (x[1][0] == x[1][1], x[0])):
    print(f"  {'✅' if d == t else '❌'} {n:<22} {d}/{t}")
print(f"\nTẦNG 1: đạt {dat}/{len(ket_qua)} kiểm tra")
sys.exit(0 if dat == len(ket_qua) else 1)
