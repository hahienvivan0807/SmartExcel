"""BƯỚC 2 – Tóm tắt bảng cho AI (không dùng AI, làm 1 lần / bảng).

Mỗi bảng có:
  - tom_tat: đoạn chữ ngắn mô tả bảng (tên cột chính xác, kiểu, giá trị, dòng mẫu)
  - tu_khoa: tập từ khóa (không dấu) để Bước 3 chọn bảng liên quan tới câu hỏi
"""
import csv
import io
from datetime import date, datetime

import pandas as pd

from .log import log
from .van_ban import tach_tu

LIET_KE_TOI_DA = 15      # cột chữ có ≤ 15 giá trị khác nhau → liệt kê hết
SO_VI_DU = 3             # cột chữ nhiều giá trị → đưa 3 ví dụ
DO_DAI_O = 30            # cắt chữ dài
SO_DONG_MAU = 2
COT_TOI_DA_CO_MAU = 25   # bảng rộng hơn thì bỏ dòng mẫu cho gọn
TU_KHOA_TOI_DA = 500     # cột có quá nhiều giá trị thì không lấy làm từ khóa


def tom_tat_tat_ca(cac_bang):
    tong = 0
    for bang in cac_bang:
        bang.tom_tat = tom_tat_bang(bang)
        bang.tu_khoa = tu_khoa_bang(bang)
        tong += len(bang.tom_tat)
    log(f"Đã tóm tắt {len(cac_bang)} bảng (tổng {tong} ký tự)")


# ============================================================
# Tóm tắt
# ============================================================

def tom_tat_bang(bang):
    df = bang.df
    dong = [f'Bảng "{bang.ten}" ({len(df)} dòng)']
    if bang.huong == "CHEO":
        dong[0] += " [bảng chéo]"
    if bang.tieu_de:
        # tiêu đề hay chứa thông tin quý như "(đơn vị: triệu đồng)" → cho dài hơn ô thường
        tieu_de = bang.tieu_de if len(bang.tieu_de) <= 100 else bang.tieu_de[:100] + "…"
        dong.append(f'Tiêu đề: "{tieu_de}"')
    if bang.mo_ta:
        dong.append(f"Mô tả: {bang.mo_ta}")
    dong.append("Cột:")
    for col in df.columns:
        dong.append(f'- "{col}": {mo_ta_cot(df[col])}')
    if df.shape[1] <= COT_TOI_DA_CO_MAU and len(df) > 0:
        dong.append(f"Mẫu {min(SO_DONG_MAU, len(df))} dòng (CSV):")
        dong.append(_dong_mau_csv(df))
    return "\n".join(dong)


def mo_ta_cot(cot):
    o = cot.dropna()
    trong = len(cot) - len(o)
    ghi_chu_trong = f" ({trong} ô trống)" if trong else ""
    if o.empty:
        return "trống hoàn toàn"

    # Cột số
    if pd.api.types.is_numeric_dtype(cot) and not pd.api.types.is_bool_dtype(cot):
        loai = "số nguyên" if (o % 1 == 0).all() else "số"
        if o.nunique() == 1:
            return f"{loai}, luôn là {_so(o.iloc[0])}{ghi_chu_trong}"
        return f"{loai} {_so(o.min())}–{_so(o.max())}{ghi_chu_trong}"

    # Cột ngày tháng
    if pd.api.types.is_datetime64_any_dtype(cot) or all(isinstance(x, (datetime, date)) for x in o):
        ngay = pd.to_datetime(o, errors="coerce").dropna()
        if not ngay.empty:
            return f"ngày {ngay.min().date()} → {ngay.max().date()}{ghi_chu_trong}"

    # Cột chữ (hoặc lẫn lộn)
    gia_tri = list(dict.fromkeys(_chu(x) for x in o))   # giữ thứ tự, bỏ trùng
    if len(gia_tri) == 1:
        return f'chữ, luôn là "{gia_tri[0]}"{ghi_chu_trong}'
    if len(gia_tri) == len(o) and len(gia_tri) > 5:
        # Mỗi dòng 1 giá trị riêng (mã, SĐT, tên) → liệt kê hết chỉ tốn token;
        # câu hỏi nhắc tới giá trị nào thì phần "gợi ý giá trị" ở Bước 3 đã tìm sẵn
        vi_du = ", ".join(f'"{v}"' for v in gia_tri[:SO_VI_DU])
        return f"chữ, mỗi dòng 1 giá trị riêng ({len(gia_tri)} giá trị), vd: {vi_du}{ghi_chu_trong}"
    if len(gia_tri) <= LIET_KE_TOI_DA:
        ds = ", ".join(f'"{v}"' for v in sorted(gia_tri))
        return f"chữ, chỉ gồm: {ds}{ghi_chu_trong}"
    vi_du = ", ".join(f'"{v}"' for v in gia_tri[:SO_VI_DU])
    return f"chữ, {len(gia_tri)} giá trị khác nhau, vd: {vi_du}{ghi_chu_trong}"


def _dong_mau_csv(df):
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\n")
    w.writerow(df.columns)
    for _, r in df.head(SO_DONG_MAU).iterrows():
        w.writerow([_o(x) for x in r])
    return buf.getvalue().rstrip("\n")


def _o(x):
    if x is None or (not isinstance(x, str) and pd.isna(x)):
        return ""
    if isinstance(x, (datetime, date, pd.Timestamp)):
        return str(pd.Timestamp(x).date())
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        return _so(x)
    return _chu(x)


def _so(x):
    """1200000.0 → '1200000', 8.333333 → '8.3333'"""
    x = float(x)
    if x.is_integer():
        return str(int(x))
    return f"{round(x, 4)}".rstrip("0").rstrip(".")


def _chu(x):
    s = str(x)
    return s if len(s) <= DO_DAI_O else s[:DO_DAI_O] + "…"


# ============================================================
# Từ khóa
# ============================================================

def tu_khoa_bang(bang):
    df = bang.df
    tu = (tach_tu(bang.ten) | tach_tu(" ".join(map(str, df.columns)))
          | tach_tu(bang.mo_ta) | tach_tu(bang.tieu_de))
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            continue
        uniq = df[col].dropna().unique()
        if len(uniq) <= TU_KHOA_TOI_DA:
            tu |= tach_tu(" ".join(map(str, uniq)))
    return tu
