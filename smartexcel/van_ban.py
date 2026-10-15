"""Xử lý chữ tiếng Việt: bỏ dấu, tách từ."""
import re
import unicodedata

TU_BO_QUA = {"xlsx", "xlsm", "sheet"}

# Từ chức năng trong câu hỏi – không giúp chọn bảng (đã bỏ dấu)
TU_DUNG = {"la", "bao", "nhieu", "co", "cua", "nao", "ai", "gi", "trong", "the",
           "va", "cac", "nhung", "mot", "duoc", "may", "bang", "nhat", "hay",
           "cho", "biet", "voi", "o", "tai", "ve", "khong"}


def _bo_dau_giu_hoa(s):
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D")


def bo_dau(s):
    """'Nguyễn Văn Đức' → 'nguyen van duc'"""
    return _bo_dau_giu_hoa(s).lower().strip()


def tach_tu(s):
    """Tập các từ (đã bỏ dấu) trong chuỗi. 'diem_thi.xlsx::DoanhSo' → {diem, thi, doanh, so}"""
    s = _bo_dau_giu_hoa(s)
    s = re.sub(r"([a-z])([A-Z])", r"\1 \2", s)       # DoanhSo → Doanh So
    return set(re.findall(r"[^\W_]+", s.lower())) - TU_BO_QUA
