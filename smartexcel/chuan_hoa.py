"""BƯỚC 5 – Chuẩn hóa đáp án thành 1 chuỗi ngắn gọn để ghi vào CSV.

Code AI có thể trả về đủ kiểu: số numpy, Series, DataFrame, list, ngày, chuỗi có chữ thừa…
Bước này đưa tất cả về 1 dạng thống nhất:
    9.0            → "9"
    38142.857142   → "38142.86"        (làm tròn SO_THAP_PHAN chữ số)
    Series ["A"]   → "A"               (1 phần tử → lấy phần tử đó)
    ["B", "A"]     → "B, A"            (nhiều phần tử → nối bằng ", ")
    Timestamp      → "01/10/2026"
    "Đáp án: 11."  → "11"
    "0312345678"   → "0312345678"      (mã có số 0 đầu giữ nguyên)

⚠️ Đề chưa nói rõ định dạng số thập phân / ngày / Có-Không mong muốn.
   Khi có đáp án mẫu chính thức → chỉnh các hằng số bên dưới cho khớp.
"""
import math
import re
from datetime import date, datetime
from decimal import Decimal

import numpy as np
import pandas as pd

from .doc_bang import tach_so_don_vi

SO_THAP_PHAN = 2                 # số chữ số sau dấu "." khi làm tròn
DINH_DANG_NGAY = "%d/%m/%Y"
DINH_DANG_NGAY_GIO = "%d/%m/%Y %H:%M"
DUNG_SAI = ("Có", "Không")       # True / False
NOI = ", "                       # nối nhiều kết quả

_TIEN_TO = re.compile(r"^\s*(đáp án|dap an|kết quả|ket qua|trả lời|tra loi|answer|result)\s*[:：=\-]\s*",
                      re.IGNORECASE)


def chuan_hoa(x):
    """Đáp án thô (bất kỳ kiểu gì) → chuỗi 1 dòng."""
    # --- Bảng / danh sách: 1 phần tử → lấy phần tử; nhiều → nối ---
    if isinstance(x, pd.DataFrame):
        return _noi(x.to_numpy().ravel().tolist())
    if isinstance(x, (pd.Series, pd.Index, np.ndarray)):
        return _noi(list(x.tolist() if hasattr(x, "tolist") else x))
    if isinstance(x, (list, tuple, set, frozenset)):
        return _noi(list(x))
    if isinstance(x, dict):
        return NOI.join(f"{chuan_hoa(k)}: {chuan_hoa(v)}" for k, v in x.items())

    # --- Giá trị đơn ---
    if x is None:
        return ""
    if isinstance(x, np.generic):          # np.int64, np.float64, np.bool_ → kiểu Python
        x = x.item()
    if isinstance(x, (bool, np.bool_)):
        return DUNG_SAI[0] if x else DUNG_SAI[1]
    if isinstance(x, (pd.Timestamp, datetime, date)):
        return _ngay(x)
    if isinstance(x, Decimal):
        x = float(x)
    if isinstance(x, int):
        return str(x)
    if isinstance(x, float):
        return _so(x)
    try:
        if pd.isna(x):                      # pd.NA, NaT
            return ""
    except (TypeError, ValueError):
        pass
    return _chu(str(x))


def _noi(ds):
    ds = [v for v in ds if chuan_hoa(v) != ""]
    if len(ds) == 1:
        return chuan_hoa(ds[0])
    return NOI.join(chuan_hoa(v) for v in ds)


def _so(x):
    if math.isnan(x) or math.isinf(x):
        return ""
    x = round(x, SO_THAP_PHAN)
    if x == int(x):
        return str(int(x))
    return f"{x:.{SO_THAP_PHAN}f}".rstrip("0").rstrip(".")


def _ngay(x):
    t = pd.Timestamp(x)
    if pd.isna(t):
        return ""
    if t.hour or t.minute:
        return t.strftime(DINH_DANG_NGAY_GIO)
    return t.strftime(DINH_DANG_NGAY)


def _chu(s):
    s = re.sub(r"\s+", " ", s).strip()            # gộp xuống dòng, khoảng trắng thừa
    s = _TIEN_TO.sub("", s)                        # bỏ "Đáp án:", "Answer:"…
    s = s.strip().strip("\"'“”‘’`").strip()       # bỏ ngoặc kép bao quanh
    if s.endswith(".") and not s.endswith(".."):   # bỏ dấu chấm cuối câu (giữ "...")
        s = s[:-1].rstrip()
    # "9.0" / "1200000.00" (AI tự đổi số thành chữ) → "9" / "1200000"
    if re.fullmatch(r"-?[1-9]\d*\.0+|-?0\.0+", s):
        return s.split(".")[0]
    # "38142.857142857" → làm tròn như số
    if re.fullmatch(r"-?(0|[1-9]\d*)\.\d+", s):
        return _so(float(s))
    # "800.000 đồng", "12,5 kg", "80%" (AI trả lời thẳng hay kèm đơn vị) → số trần
    # (mã có số 0 đầu như "0312345678" không bị đổi – tach_so_don_vi tự bỏ qua)
    so, _ = tach_so_don_vi(s)
    if so is not None and not isinstance(so, bool):
        return _so(float(so))
    return s
