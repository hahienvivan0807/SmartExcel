"""BƯỚC 4 – Chạy code AI viết; lỗi thì gửi lỗi lại cho AI sửa.

    code ──► chạy (giới hạn thời gian, chặn đọc/ghi file)
         ├─ OK, có kết quả ──► trả về answer
         └─ lỗi / kết quả rỗng ──► gửi 1 dòng lỗi + gợi ý cho AI ──► code mới (tối đa 2 lần sửa)
"""
import builtins
import math
import re
import threading
import traceback

import numpy as np
import pandas as pd

from .hoi_dap import lay_code, viet_code
from .log import log

THOI_GIAN_TOI_DA = 20     # giây cho mỗi lần chạy code
SO_LAN_SUA = 2

# Thư viện code AI được phép import
THU_VIEN_CHO_PHEP = {"pandas", "numpy", "math", "re", "datetime", "statistics",
                     "collections", "itertools", "functools", "unicodedata", "decimal"}
# Hàm có sẵn bị cấm (đọc/ghi file, chạy code lạ, thoát chương trình)
HAM_CAM = {"open", "exec", "eval", "compile", "input", "exit", "quit",
           "breakpoint", "help", "globals", "vars", "memoryview"}


class LoiChayCode(Exception):
    pass


# ============================================================
# Chạy code
# ============================================================

def _import_an_toan(ten, *args, **kwargs):
    if ten.split(".")[0] not in THU_VIEN_CHO_PHEP:
        raise ImportError(f"Không được import '{ten}'. Chỉ dùng: pandas, numpy, math, re, datetime")
    return builtins.__import__(ten, *args, **kwargs)


def _moi_truong(dfs):
    ham = {k: v for k, v in vars(builtins).items() if k not in HAM_CAM}
    ham["__import__"] = _import_an_toan
    ham["print"] = lambda *a, **k: None   # code AI có print cũng không in ra
    # Bản sao nông: pandas 3 có copy-on-write nên code AI sửa bảng cũng không hỏng bảng gốc
    ban_sao = {ten: df.copy(deep=False) for ten, df in dfs.items()}
    return {"__builtins__": ham, "dfs": ban_sao, "pd": pd, "np": np, "math": math}


def chay_code(code, dfs):
    """Chạy code, trả về giá trị biến `answer`. Lỗi → ném LoiChayCode (câu ngắn gửi được cho AI)."""
    if not code.strip():
        raise LoiChayCode("Không có code.")

    env = _moi_truong(dfs)
    ket_qua = {}

    def chay():
        try:
            exec(compile(code, "<code_ai>", "exec"), env)
            ket_qua["ok"] = True
        except BaseException:
            ket_qua["loi"] = traceback.format_exc()

    t = threading.Thread(target=chay, daemon=True)
    t.start()
    t.join(THOI_GIAN_TOI_DA)
    if t.is_alive():
        raise LoiChayCode(f"Code chạy quá {THOI_GIAN_TOI_DA} giây. Hãy viết cách đơn giản hơn, không dùng vòng lặp lớn.")
    if "loi" in ket_qua:
        raise LoiChayCode(_tom_tat_loi(ket_qua["loi"], code))
    if "answer" not in env:
        raise LoiChayCode("Code chưa gán kết quả vào biến `answer`.")

    answer = env["answer"]
    if _rong(answer):
        raise LoiChayCode(
            f"Code chạy được nhưng answer rỗng ({_mo_ta_ngan(answer)}). "
            "Có thể lọc sai giá trị (kiểm tra dấu, hoa thường, khoảng trắng) hoặc chọn sai cột.")
    return answer


def _tom_tat_loi(tb, code):
    """Chỉ lấy dòng lỗi cuối + dòng code gây lỗi – đủ để AI sửa, tốn ít token."""
    loi = tb.strip().splitlines()[-1]
    so_dong = re.findall(r'File "<code_ai>", line (\d+)', tb)
    if so_dong:
        dong_code = code.splitlines()
        n = int(so_dong[-1])
        if 0 < n <= len(dong_code):
            return f"{loi}\nTại dòng {n}: {dong_code[n - 1].strip()}"
    return loi


def _rong(x):
    if x is None:
        return True
    if isinstance(x, (pd.DataFrame, pd.Series, list, tuple, set, dict, np.ndarray)):
        return len(x) == 0
    if isinstance(x, str):
        return x.strip() == ""
    try:
        return bool(pd.isna(x))
    except (TypeError, ValueError):
        return False


def _mo_ta_ngan(x):
    if isinstance(x, (pd.DataFrame, pd.Series)):
        return f"{type(x).__name__} rỗng"
    return repr(x)[:50]


# ============================================================
# Vòng lặp: viết code → chạy → sửa
# ============================================================

def _goi_y_sua(loi, dfs, bang_dung):
    """Thêm thông tin giúp AI sửa đúng chỗ (vd sai tên cột → đưa danh sách cột thật)."""
    if "KeyError" in loi:
        cac_cot = "\n".join(f'- "{ten}": {list(dfs[ten].columns)}' for ten in bang_dung if ten in dfs)
        return f"\nTên bảng/cột đúng:\n{cac_cot}"
    return ""


def thu_va_sua(ai, qid, code, messages, dfs, ten_bang, so_lan_sua=SO_LAN_SUA, nhan=""):
    """Chạy code; lỗi thì gửi lại cho AI sửa (tối đa so_lan_sua lần).
    Trả về (answer, None) nếu thành công, (None, lỗi cuối) nếu thất bại hết."""
    loi = "Không có code."
    for lan in range(so_lan_sua + 1):
        try:
            answer = chay_code(code, dfs)
            log(f"[câu {qid}] ✅ {nhan}lần {lan + 1}: {_mo_ta_ngan(answer)}")
            return answer, None
        except LoiChayCode as e:
            loi = str(e)
            log(f"[câu {qid}] ❌ {nhan}lần {lan + 1}: {loi.splitlines()[0]}")

        if ai is None or lan == so_lan_sua:
            break
        messages = messages + [
            {"role": "assistant", "content": f"```python\n{code}\n```"},
            {"role": "user", "content": f"Code bị lỗi:\n{loi}{_goi_y_sua(loi, dfs, ten_bang)}\n"
                                        "Hãy sửa lại. Chỉ trả về code."},
        ]
        code = lay_code(ai.hoi(messages, max_tokens=500))
    return None, loi


def giai_cau_hoi(ai, qid, cau_hoi, cac_bang, dfs):
    """Cách chính: AI viết code → chạy → sửa (tối đa 3 lần).
    Trả về (answer, lỗi cuối, bảng đã chọn). answer = None nếu thất bại."""
    code, messages, bang_chon = viet_code(ai, cau_hoi, cac_bang)
    answer, loi = thu_va_sua(ai, qid, code, messages, dfs, [b.ten for b in bang_chon])
    return answer, loi, bang_chon
