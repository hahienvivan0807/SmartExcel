"""BƯỚC 3 – Lọc bảng, gợi ý giá trị và nhờ AI viết code pandas (làm cho mỗi câu hỏi).

    câu hỏi ──► chọn 2-3 bảng liên quan (từ khóa, không dùng AI)
            ──► tìm sẵn tên/giá trị câu hỏi nhắc tới (không dùng AI)
            ──► gửi tóm tắt + gợi ý + câu hỏi cho AI ──► nhận code
"""
import re
from collections import Counter

import pandas as pd

from .van_ban import TU_DUNG, bo_dau, tach_tu

SO_BANG_CHON = 3
TI_LE_GIU = 0.4
GOI_Y_TOI_DA = 10
GIA_TRI_TOI_DA = 5000     # cột có quá nhiều giá trị thì bỏ qua khi tìm gợi ý


# ============================================================
# 1. Chọn bảng liên quan
# ============================================================

def xep_hang_bang(cau_hoi, cac_bang):
    """Chấm điểm mỗi bảng theo số từ khóa trùng với câu hỏi, trả về [(điểm, bảng)] từ cao xuống thấp.

    Từ càng hiếm (ít bảng có) thì càng có giá trị → điểm = Σ 1 / (số bảng chứa từ).
    """
    tu_hoi = tach_tu(cau_hoi) - TU_DUNG
    so_bang_chua = Counter(t for b in cac_bang for t in b.tu_khoa)
    diem = [(sum(1 / so_bang_chua[t] for t in tu_hoi & b.tu_khoa), b) for b in cac_bang]
    return sorted(diem, key=lambda x: x[0], reverse=True)


def chon_bang(cau_hoi, cac_bang, so_bang=SO_BANG_CHON):
    """Chọn ≤ 3 bảng liên quan. Không bảng nào khớp → trả về tất cả (an toàn hơn chọn bừa)."""
    if len(cac_bang) <= so_bang:
        return list(cac_bang)
    diem = xep_hang_bang(cau_hoi, cac_bang)
    cao_nhat = diem[0][0]
    if cao_nhat == 0:
        return list(cac_bang)
    # Bỏ bảng chỉ khớp lặt vặt (vd "toán" trong "Kế toán") – giữ bảng ≥ 40% điểm bảng đầu
    return [b for d, b in diem[:so_bang] if d >= cao_nhat * TI_LE_GIU]


# ============================================================
# 2. Gợi ý giá trị
# ============================================================

def tim_gia_tri(cau_hoi, cac_bang):
    """Tìm sẵn trong bảng những giá trị chữ mà câu hỏi nhắc tới → [(tên bảng, cột, giá trị)].

    So khớp không dấu, không phân biệt hoa thường, theo nguyên từ
    (giá trị "An" không khớp với chữ "Thanh" trong câu hỏi).
    """
    hoi = " " + " ".join(re.findall(r"[^\W_]+", bo_dau(cau_hoi))) + " "
    tim_thay = []
    for b in cac_bang:
        for col in b.df.columns:
            if pd.api.types.is_numeric_dtype(b.df[col]) or pd.api.types.is_datetime64_any_dtype(b.df[col]):
                continue
            uniq = b.df[col].dropna().unique()
            if len(uniq) > GIA_TRI_TOI_DA:
                continue
            khop = []
            for v in uniq:
                v_bd = " ".join(re.findall(r"[^\W_]+", bo_dau(v)))
                if len(v_bd) >= 2 and f" {v_bd} " in hoi:
                    khop.append(str(v))
            # Giữ giá trị dài nhất: khớp cả "Ngô Minh Châu" thì bỏ "Châu"
            khop = [v for v in khop
                    if not any(v != k and bo_dau(v) in bo_dau(k) for k in khop)]
            tim_thay += [(b.ten, col, v) for v in khop]
    return tim_thay


def goi_y_gia_tri(cau_hoi, cac_bang):
    return [f'"{v}" có trong cột "{col}" của bảng "{ten}"'
            for ten, col, v in tim_gia_tri(cau_hoi, cac_bang)][:GOI_Y_TOI_DA]


# ============================================================
# 3. Nhờ AI viết code
# ============================================================

PROMPT_VIET_CODE = """Bạn là chuyên gia pandas. Viết code Python trả lời câu hỏi về bảng tính.

Quy tắc:
- Biến `dfs` là dict {tên bảng: DataFrame} đã có sẵn; `pd` là pandas. KHÔNG đọc file.
- Dùng ĐÚNG tên bảng và tên cột như trong mô tả.
- Lọc chữ theo đúng giá trị trong mô tả/gợi ý (chú ý dấu và hoa thường).
- Gán đáp án cuối cùng vào biến `answer`: 1 số hoặc 1 chuỗi ngắn.
  Nhiều kết quả → nối thành chuỗi, cách nhau bởi ", ". Không làm tròn trừ khi câu hỏi yêu cầu.
- Không print, không giải thích. Chỉ trả về code trong ```python ... ```.

Ví dụ:
Bảng "lop.xlsx::S1" (3 dòng)
Cột:
- "Tên": chữ, chỉ gồm: "An", "Bình", "Chi"
- "Điểm": số 6–9
Câu hỏi: Ai có điểm cao nhất?
```python
df = dfs["lop.xlsx::S1"]
answer = df.loc[df["Điểm"].idxmax(), "Tên"]
```"""


def tao_messages(cau_hoi, cac_bang, goi_y):
    noi_dung = "\n\n".join(b.tom_tat for b in cac_bang)
    if goi_y:
        noi_dung += "\n\nGợi ý (tìm thấy trong dữ liệu):\n" + "\n".join(f"- {g}" for g in goi_y)
    noi_dung += f"\n\nCâu hỏi: {cau_hoi}"
    return [
        {"role": "system", "content": PROMPT_VIET_CODE},
        {"role": "user", "content": noi_dung},
    ]


def lay_code(tra_loi):
    """Lấy phần code trong ```python ... ``` (AI hay viết thêm chữ bên ngoài)."""
    m = re.search(r"```(?:python|py)?\s*\n(.*?)```", tra_loi, re.S)
    return (m.group(1) if m else tra_loi).strip()


def viet_code(ai, cau_hoi, cac_bang):
    """Trả về (code, messages, bang_chon). messages dùng lại ở Bước 4 khi cần sửa lỗi."""
    bang_chon = chon_bang(cau_hoi, cac_bang)
    goi_y = goi_y_gia_tri(cau_hoi, bang_chon)
    messages = tao_messages(cau_hoi, bang_chon, goi_y)
    tra_loi = ai.hoi(messages, max_tokens=500) if ai is not None else ""
    return lay_code(tra_loi), messages, bang_chon
